// Browse - simple file explorer with basic video playback via libVLC
// Forked from MediaExplorer, but with all combine/FFmpeg edit features removed.
//
// Features:
//  - Drives view, folder view, optional recursive video search
//  - Shows ALL files (not just videos) in folder view
//  - Double-click / Enter:
//        drive/folder -> navigate
//        video file(s) -> play with libVLC (playlist)
//        non-video file -> open with default app (ShellExecuteW)
//  - Cut / Copy / Paste for files AND directories (with progress dialog & cancel)
//  - Recursive delete for directories
//  - Rename files/directories (F2, or context menu)
//  - Right-click context menu: Open, Play video, Rename, Cut, Copy, Paste, Delete,
//    Map Network Drive..., Disconnect Network Drive...
//  - Optional command-line start folder: Browse.exe C:\data\new
//
// Build: VS2022, x64, C++17
// Project .vcxproj should link to libvlc.lib/libvlccore.lib as in MediaExplorer.

#ifndef UNICODE
#  define UNICODE
#endif
#ifndef _UNICODE
#  define _UNICODE
#endif
#ifndef WIN32_LEAN_AND_MEAN
#  define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#  define NOMINMAX
#endif

#include <windows.h>
#include <windowsx.h>      // GET_X_LPARAM, GET_Y_LPARAM
#include <commctrl.h>
#include <shlwapi.h>
#include <shobjidl_core.h>
#include <propsys.h>
#include <propkey.h>
#include <wrl/client.h>
#include <winnetwk.h>
#include <shellapi.h>      // ShellExecuteW, CommandLineToArgvW

#include <string>
#include <vector>
#include <algorithm>
#include <atomic>
#include <cstdint>
#include <cwchar>
#include <climits>
#include <cstdio>
#include <cstdlib>
#include <fstream>
#include <sstream>
#include <cwctype>
#include <shlobj.h>        // SHCreateDirectoryExW
#include <cstdarg>
#include <io.h>            // _unlink

#ifndef CFSTR_PREFERREDDROPEFFECT
#define CFSTR_PREFERREDDROPEFFECT L"Preferred DropEffect"
#endif

#ifndef DROPEFFECT_COPY
#define DROPEFFECT_COPY 0x00000001
#endif
#ifndef DROPEFFECT_MOVE
#define DROPEFFECT_MOVE 0x00000002
#endif


#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "User32.lib")
#pragma comment(lib, "Gdi32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Propsys.lib")
#pragma comment(lib, "Uuid.lib")
#pragma comment(lib, "Comdlg32.lib")
#pragma comment(lib, "OleAut32.lib")
#pragma comment(lib, "Mpr.lib")        // WNetConnectionDialog / WNetDisconnectDialog
#pragma comment(lib, "Advapi32.lib")  // RegOpenKeyEx / RegQueryValueEx


#include <vlc/vlc.h>

#ifndef FIND_FIRST_EX_LARGE_FETCH
#  define FIND_FIRST_EX_LARGE_FETCH 0x00000002
#endif

using Microsoft::WRL::ComPtr;

static void LogLine(const wchar_t* fmt, ...); // forward declaration
static DWORD GetConnectedNetDriveMask();
static bool GetPersistentMappedRemotePath(wchar_t letter, std::wstring& outRemote);

// ----------------------------- Globals

static bool g_loadingFolder = false;
HINSTANCE g_hInst = NULL;
HWND g_hwndMain = NULL, g_hwndList = NULL, g_hwndVideo = NULL, g_hwndSeek = NULL;

enum class ViewKind { Drives, Folder, Search };
ViewKind g_view = ViewKind::Drives;
std::wstring g_folder;        // valid in Folder view
std::wstring g_initialPath;   // optional start folder from command line

struct Row {
    std::wstring name;     // display name (for Search, full path; for Folder, file name)
    std::wstring full;     // absolute path
    bool         isDir;
    ULONGLONG    size;
    FILETIME     modified;
    // video props
    int          vW, vH;
    ULONGLONG    vDur100ns;

    // NEW: drives-view network status
    bool         isBrokenNetDrive;
    std::wstring netRemote; // \\server\share (best-effort)

    Row() : isDir(false), size(0), vW(0), vH(0), vDur100ns(0),
        isBrokenNetDrive(false) {
        modified.dwLowDateTime = modified.dwHighDateTime = 0;
    }
};
std::vector<Row> g_rows;

// sorting
int  g_sortCol = 0;      // 0=Name,1=Type,2=Size,3=Modified,4=Resolution,5=Duration
bool g_sortAsc = true;

// VLC
libvlc_instance_t* g_vlc = NULL;
libvlc_media_player_t* g_mp = NULL;
bool                     g_inPlayback = false;
std::vector<std::wstring> g_playlist;
size_t                   g_playlistIndex = 0;
bool                     g_userDragging = false;
libvlc_time_t            g_lastLenForRange = -1;

// ----------------------------- Configuration (browse.ini)

struct AppConfig {
    std::wstring upscaleDirectory;  // unused in this fork, kept for future
    bool ffmpegAvailable = false;   // unused in this fork
    bool ffprobeAvailable = false;
    bool loggingEnabled = false;
    std::wstring loggingPath;
    std::wstring logFile;

    // NEW: default credentials for network reconnect/map (optional)
    std::wstring netUsername;
    std::wstring netPassword;
};

AppConfig g_cfg;

// fullscreen (app-managed)
bool g_fullscreen = false;
WINDOWPLACEMENT g_wpPrev;

// timers
const UINT_PTR kTimerPlaybackUI = 1;

// post-playback actions
enum class ActionType { DeleteFile, RenameFile, CopyToPath };
struct PostAction { ActionType type; std::wstring src; std::wstring param; };
std::vector<PostAction> g_post;

enum class OpResult { Success, Cancelled, Failed };

// filename clipboard for browser
enum class ClipMode { None, Copy, Move };
ClipMode g_clipMode = ClipMode::None;
std::vector<std::wstring> g_clipFiles; // absolute paths (files or dirs)


// Make a file or directory writable so DeleteFile/RemoveDirectory will work.
static void ClearReadonlyAndSystem(const std::wstring& path)
{
    DWORD attrs = GetFileAttributesW(path.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES) return;

    DWORD newAttrs = attrs & ~(FILE_ATTRIBUTE_READONLY | FILE_ATTRIBUTE_SYSTEM);
    if (newAttrs != attrs) {
        SetFileAttributesW(path.c_str(), newAttrs);
    }
}

static void SetClipboardFileDrop(const std::vector<std::wstring>& files, ClipMode mode) {
    if (files.empty()) return;

    if (!OpenClipboard(g_hwndMain)) return;
    EmptyClipboard();

    // Compute space needed: sum of (len+1) for each file plus final extra NUL
    size_t totalChars = 0;
    for (const auto& f : files) {
        totalChars += f.size() + 1; // including per-file NUL
    }
    totalChars += 1; // final double-NUL terminator

    const SIZE_T bytes = sizeof(DROPFILES) + totalChars * sizeof(wchar_t);
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (!hMem) {
        CloseClipboard();
        return;
    }

    DROPFILES* pDrop = (DROPFILES*)GlobalLock(hMem);
    if (!pDrop) {
        GlobalFree(hMem);
        CloseClipboard();
        return;
    }

    pDrop->pFiles = sizeof(DROPFILES);
    pDrop->pt.x = 0;
    pDrop->pt.y = 0;
    pDrop->fNC = FALSE;
    pDrop->fWide = TRUE;

    wchar_t* dst = (wchar_t*)((BYTE*)pDrop + sizeof(DROPFILES));
    for (const auto& f : files) {
        wcscpy_s(dst, f.size() + 1, f.c_str());
        dst += f.size() + 1;
    }
    *dst = L'\0'; // extra final NUL

    GlobalUnlock(hMem);
    SetClipboardData(CF_HDROP, hMem);

    // Set Preferred DropEffect so other apps know Copy vs Move
    UINT fmtDropEffect = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
    HGLOBAL hEffect = GlobalAlloc(GMEM_MOVEABLE, sizeof(DWORD));
    if (hEffect) {
        DWORD* pEffect = (DWORD*)GlobalLock(hEffect);
        if (pEffect) {
            *pEffect = (mode == ClipMode::Move) ? DROPEFFECT_MOVE : DROPEFFECT_COPY;
            GlobalUnlock(hEffect);
            SetClipboardData(fmtDropEffect, hEffect);
        }
        else {
            GlobalFree(hEffect);
        }
    }

    CloseClipboard();
}

static bool GetClipboardFileDrop(std::vector<std::wstring>& outFiles, ClipMode* outMode /*= nullptr*/) {
    outFiles.clear();
    if (outMode) *outMode = ClipMode::Copy;

    if (!OpenClipboard(g_hwndMain)) return false;

    HANDLE hDrop = GetClipboardData(CF_HDROP);
    if (!hDrop) {
        CloseClipboard();
        return false;
    }

    // Read Preferred DropEffect (Explorer sets MOVE for Ctrl+X, COPY for Ctrl+C)
    if (outMode) {
        UINT fmtDropEffect = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
        HANDLE hEff = GetClipboardData(fmtDropEffect);
        if (hEff) {
            DWORD* pEff = (DWORD*)GlobalLock(hEff);
            if (pEff) {
                DWORD eff = *pEff;
                GlobalUnlock(hEff);
                *outMode = (eff & DROPEFFECT_MOVE) ? ClipMode::Move : ClipMode::Copy;
            }
        }
    }

    HDROP drop = (HDROP)hDrop;
    UINT count = DragQueryFileW(drop, 0xFFFFFFFF, NULL, 0);
    for (UINT i = 0; i < count; ++i) {
        UINT len = DragQueryFileW(drop, i, NULL, 0);
        if (!len) continue;
        std::wstring path(len, L'\0');
        DragQueryFileW(drop, i, &path[0], len + 1);
        path.resize(wcslen(path.c_str()));
        if (!path.empty())
            outFiles.push_back(path);
    }

    CloseClipboard();
    return !outFiles.empty();
}

// ----------------------------- Search state (for videos only, as original)

struct SearchState {
    bool active;
    ViewKind originView;
    std::wstring originFolder;
    std::vector<std::wstring> termsLower;

    bool useExplicitScope;
    std::vector<std::wstring> explicitFolders;
    std::vector<std::wstring> explicitFiles;

    SearchState()
        : active(false),
        originView(ViewKind::Drives),
        useExplicitScope(false) {
    }
} g_search;

// ----------------------------- Async metadata fill

constexpr UINT WM_APP_META = WM_APP + 100;

struct MetaResult {
    std::wstring path;
    int          w, h;
    ULONGLONG    dur;
    uint32_t     gen;
};

std::atomic<uint32_t> g_metaGen{ 0 };
CRITICAL_SECTION      g_metaLock;
std::vector<std::wstring> g_metaTodoPaths;
HANDLE                g_metaThread = NULL;

// ----------------------------- Context-menu command IDs

enum {
    ID_CTX_OPEN = 30001,
    ID_CTX_PLAY = 30002,
    ID_CTX_RENAME = 30003,
    ID_CTX_CUT = 30004,
    ID_CTX_COPY = 30005,
    ID_CTX_PASTE = 30006,
    ID_CTX_DELETE = 30007,
    ID_CTX_MAPDRIVE = 30008,
    ID_CTX_DISCONNECT = 30009,
    // NEW
    ID_CTX_FIXDRIVE = 30010
};

// ----------------------------- Helpers

static inline bool IsDriveRoot(const std::wstring& p) {
    return p.size() == 3 &&
        ((p[0] >= L'A' && p[0] <= L'Z') || (p[0] >= L'a' && p[0] <= L'z')) &&
        p[1] == L':' && (p[2] == L'\\' || p[2] == L'/');
}

static inline std::wstring EnsureSlash(std::wstring p) {
    if (!p.empty() && p.back() != L'\\' && p.back() != L'/') p.push_back(L'\\');
    return p;
}

static void CollectSelection(std::vector<std::wstring>& outFolders,
    std::vector<std::wstring>& outFiles) {
    outFolders.clear();
    outFiles.clear();
    int idx = -1;
    while ((idx = ListView_GetNextItem(g_hwndList, idx, LVNI_SELECTED)) != -1) {
        if (idx < 0 || idx >= (int)g_rows.size()) continue;
        const Row& r = g_rows[idx];
        if (r.isDir) {
            outFolders.push_back(EnsureSlash(r.full));
        }
        else {
            outFiles.push_back(r.full);
        }
    }
}

// --- Search progress title + UI pumping

static void PumpMessagesThrottled(DWORD msInterval) {
    static DWORD s_last = 0;
    DWORD now = GetTickCount();
    if (now - s_last < msInterval) return;
    s_last = now;

    MSG msg;
    while (PeekMessageW(&msg, NULL, 0, 0, PM_REMOVE)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}

static void SetTitleSearchingFolder(const std::wstring& folder) {
    std::wstring t = L"Browse - searching ";
    t += EnsureSlash(folder);
    SetWindowTextW(g_hwndMain, t.c_str());
    PumpMessagesThrottled(50);
}

static std::wstring ParentDir(std::wstring p) {
    p = EnsureSlash(p);
    if (IsDriveRoot(p)) return L"";
    p.pop_back();
    size_t cut = p.find_last_of(L"\\/");
    if (cut == std::wstring::npos) return L"";
    return p.substr(0, cut + 1);
}

static std::wstring ToLower(const std::wstring& s) {
    std::wstring t = s;
    std::transform(t.begin(), t.end(), t.begin(), ::towlower);
    return t;
}

static std::wstring Trim(const std::wstring& s) {
    size_t start = 0, end = s.size();
    while (start < end && iswspace(s[start])) ++start;
    while (end > start && iswspace(s[end - 1])) --end;
    return s.substr(start, end - start);
}

// ----------------------------- Logging

static void InitLoggingFromConfig() {
    if (!g_cfg.loggingEnabled) return;
    if (g_cfg.loggingPath.empty()) return;

    std::wstring folder = Trim(g_cfg.loggingPath);
    if (folder.empty()) {
        g_cfg.loggingEnabled = false;
        return;
    }
    if (folder.back() != L'\\' && folder.back() != L'/')
        folder.push_back(L'\\');

    int rc = SHCreateDirectoryExW(NULL, folder.c_str(), NULL);
    if (rc != ERROR_SUCCESS && rc != ERROR_ALREADY_EXISTS && rc != ERROR_FILE_EXISTS) {
        g_cfg.loggingEnabled = false;
        return;
    }

    g_cfg.loggingPath = folder;
    g_cfg.logFile = folder + L"browse.log";
}

static void LogLine(const wchar_t* fmt, ...) {
    if (!g_cfg.loggingEnabled || g_cfg.logFile.empty()) return;

    FILE* f = _wfopen(g_cfg.logFile.c_str(), L"a, ccs=UTF-8");
    if (!f) return;

    SYSTEMTIME st; GetLocalTime(&st);
    wchar_t timeBuf[64];
    swprintf_s(timeBuf, L"%04u-%02u-%02u %02u:%02u:%02u.%03u",
        st.wYear, st.wMonth, st.wDay,
        st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);

    DWORD tid = GetCurrentThreadId();

    wchar_t msgBuf[1024];
    va_list ap;
    va_start(ap, fmt);
    _vsnwprintf_s(msgBuf, _countof(msgBuf), _TRUNCATE, fmt, ap);
    va_end(ap);

    fwprintf(f, L"%s [T%u] %s\n", timeBuf, (unsigned)tid, msgBuf);
    fclose(f);
}

static std::wstring FormatSize(ULONGLONG bytes) {
    const wchar_t* u[] = { L"B", L"KB", L"MB", L"GB", L"TB" };
    double v = (double)bytes; int i = 0;
    while (v >= 1024.0 && i < 4) { v /= 1024.0; ++i; }
    wchar_t buf[64]; swprintf_s(buf, L"%.2f %s", v, u[i]);
    return buf;
}

static std::wstring FormatFileTime(const FILETIME& ft) {
    SYSTEMTIME utc, loc;
    FileTimeToSystemTime(&ft, &utc);
    SystemTimeToTzSpecificLocalTime(NULL, &utc, &loc);
    wchar_t buf[64]; swprintf_s(buf, L"%04u-%02u-%02u %02u:%02u",
        loc.wYear, loc.wMonth, loc.wDay, loc.wHour, loc.wMinute);
    return buf;
}

static std::wstring FormatHMSms(LONGLONG ms) {
    if (ms < 0) ms = 0;
    LONGLONG s = ms / 1000, h = s / 3600, m = (s % 3600) / 60, sec = s % 60;
    wchar_t buf[64];
    if (h > 0) swprintf_s(buf, L"%lld:%02lld:%02lld", h, m, sec);
    else       swprintf_s(buf, L"%lld:%02lld", m, sec);
    return buf;
}

static std::wstring FormatDuration100ns(ULONGLONG d100) {
    return FormatHMSms((LONGLONG)(d100 / 10000ULL));
}

static std::string ToUtf8(const std::wstring& ws) {
    if (ws.empty()) return std::string();
    int n = WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), NULL, 0, NULL, NULL);
    std::string s(n, '\0');
    WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), &s[0], n, NULL, NULL);
    return s;
}

static std::wstring ExtLower(const std::wstring& p) {
    size_t dot = p.find_last_of(L'.');
    if (dot == std::wstring::npos) return L"";
    std::wstring e = p.substr(dot);
    std::transform(e.begin(), e.end(), e.begin(), ::towlower);
    return e;
}

static bool IsVideoFile(const std::wstring& path) {
    static const wchar_t* exts[] = {
        L".mp4", L".mkv", L".mov", L".avi", L".wmv", L".m4v",
        L".ts", L".m2ts", L".webm", L".flv", L".rm",
    };
    std::wstring e = ExtLower(path);
    for (size_t i = 0; i < _countof(exts); ++i)
        if (e == exts[i]) return true;
    return false;
}

// Fast cached video props
static bool GetVideoPropsFastCached(const std::wstring& path,
    int& outW, int& outH, ULONGLONG& outDur100ns) {
    outW = outH = 0; outDur100ns = 0;
    ComPtr<IShellItem2> item;
    if (FAILED(SHCreateItemFromParsingName(path.c_str(), NULL, IID_PPV_ARGS(&item)))) return false;
    ComPtr<IPropertyStore> store;
    if (FAILED(item->GetPropertyStore(GPS_FASTPROPERTIESONLY, IID_PPV_ARGS(&store)))) return false;

    PROPVARIANT v; PropVariantInit(&v);
    if (SUCCEEDED(store->GetValue(PKEY_Video_FrameWidth, &v)) && v.vt == VT_UI4) outW = (int)v.ulVal;
    PropVariantClear(&v);
    if (SUCCEEDED(store->GetValue(PKEY_Video_FrameHeight, &v)) && v.vt == VT_UI4) outH = (int)v.ulVal;
    PropVariantClear(&v);
    if (SUCCEEDED(store->GetValue(PKEY_Media_Duration, &v)) && (v.vt == VT_UI8 || v.vt == VT_UI4)) {
        outDur100ns = (v.vt == VT_UI8) ? v.uhVal.QuadPart : (ULONGLONG)v.ulVal;
    }
    PropVariantClear(&v);
    return (outW | outH | outDur100ns) != 0;
}

// Full video props
static bool GetVideoProps(const std::wstring& path,
    int& outW, int& outH, ULONGLONG& outDur100ns) {
    outW = outH = 0; outDur100ns = 0;
    ComPtr<IShellItem2> item;
    if (FAILED(SHCreateItemFromParsingName(path.c_str(), NULL, IID_PPV_ARGS(&item)))) return false;
    ComPtr<IPropertyStore> store;
    if (FAILED(item->GetPropertyStore(GPS_DEFAULT, IID_PPV_ARGS(&store)))) return false;

    PROPVARIANT v; PropVariantInit(&v);
    if (SUCCEEDED(store->GetValue(PKEY_Video_FrameWidth, &v)) && v.vt == VT_UI4) outW = (int)v.ulVal;
    PropVariantClear(&v);
    if (SUCCEEDED(store->GetValue(PKEY_Video_FrameHeight, &v)) && v.vt == VT_UI4) outH = (int)v.ulVal;
    PropVariantClear(&v);
    if (SUCCEEDED(store->GetValue(PKEY_Media_Duration, &v)) && (v.vt == VT_UI8 || v.vt == VT_UI4)) {
        outDur100ns = (v.vt == VT_UI8) ? v.uhVal.QuadPart : (ULONGLONG)v.ulVal;
    }
    PropVariantClear(&v);
    return (outW | outH | outDur100ns) != 0;
}

// Title helpers

static void SetTitlePlaying() {
    if (!g_inPlayback || g_playlist.empty()) return;

    const std::wstring& full = g_playlist[g_playlistIndex];
    const wchar_t* base = wcsrchr(full.c_str(), L'\\'); base = base ? base + 1 : full.c_str();

    libvlc_time_t cur = g_mp ? libvlc_media_player_get_time(g_mp) : 0;
    libvlc_time_t len = g_mp ? libvlc_media_player_get_length(g_mp) : 0;

    std::wstring left;
    if (g_playlist.size() <= 1) {
        left = L"(Single File) ";
    }
    else {
        wchar_t buf[64];
        swprintf_s(buf, L"(Playlist %zu of %zu) ", g_playlistIndex + 1, g_playlist.size());
        left = buf;
    }

    std::wstring t = left;
    t += base;
    t += L"  ";
    t += FormatHMSms(cur);
    t += L" / ";
    t += FormatHMSms(len);
    SetWindowTextW(g_hwndMain, t.c_str());
}

static std::wstring JoinTermsForTitle() {
    if (!g_search.active || g_search.termsLower.empty()) return L"";
    std::wstring s = L"\""; s += g_search.termsLower[0]; s += L"\"";
    for (size_t i = 1; i < g_search.termsLower.size(); ++i) {
        s += L" & \"";
        s += g_search.termsLower[i];
        s += L"\"";
    }
    return s;
}

static void SetTitleFolderOrDrives() {
    std::wstring t = L"Browse - ";
    if (g_view == ViewKind::Drives) t += L"[Drives]";
    else if (g_view == ViewKind::Folder) t += EnsureSlash(g_folder);
    else t += L"Search - " + JoinTermsForTitle();
    SetWindowTextW(g_hwndMain, t.c_str());
}

// ----------------------------- ffprobe helpers

static bool RunFfprobeCommand(const std::wstring& cmdLine,
    std::vector<std::string>& outLines) {
    outLines.clear();

    FILE* f = _wpopen(cmdLine.c_str(), L"rt");
    if (!f) return false;

    char buf[512];
    while (fgets(buf, sizeof(buf), f)) {
        size_t len = strlen(buf);
        while (len && (buf[len - 1] == '\n' || buf[len - 1] == '\r')) {
            buf[--len] = '\0';
        }
        outLines.emplace_back(buf);
    }
    int rc = _pclose(f);
    return (rc == 0);
}

static bool GetMediaInfoFromFfprobe(const std::wstring& path,
    int& outW,
    int& outH,
    std::wstring& outVideoCodec,
    std::wstring& outAudioCodec) {
    outW = outH = 0;
    outVideoCodec.clear();
    outAudioCodec.clear();
    bool gotV = false;
    bool gotA = false;

    std::wstring cmdV =
        L"ffprobe -v error "
        L"-select_streams v:0 "
        L"-show_entries stream=codec_name,width,height "
        L"-of default=noprint_wrappers=1 \"";
    cmdV += path;
    cmdV += L"\"";

    std::vector<std::string> linesV;
    if (RunFfprobeCommand(cmdV, linesV)) {
        std::string codecV;
        int wTmp = 0, hTmp = 0;

        for (const auto& line : linesV) {
            if (line.empty()) continue;
            const char* kCodec = "codec_name=";
            const size_t codecLen = sizeof("codec_name=") - 1;
            if (line.size() >= codecLen && line.compare(0, codecLen, kCodec) == 0) {
                codecV = line.substr(codecLen);
                continue;
            }
            const char* kWidth = "width=";
            const size_t widthLen = sizeof("width=") - 1;
            if (line.size() >= widthLen && line.compare(0, widthLen, kWidth) == 0) {
                wTmp = std::strtol(line.c_str() + widthLen, nullptr, 10);
                continue;
            }
            const char* kHeight = "height=";
            const size_t heightLen = sizeof("height=") - 1;
            if (line.size() >= heightLen && line.compare(0, heightLen, kHeight) == 0) {
                hTmp = std::strtol(line.c_str() + heightLen, nullptr, 10);
                continue;
            }
        }

        if (wTmp > 0 && hTmp > 0) {
            outW = wTmp;
            outH = hTmp;
        }
        if (!codecV.empty()) {
            outVideoCodec.assign(codecV.begin(), codecV.end());
        }

        gotV = (wTmp > 0 || hTmp > 0 || !codecV.empty());
    }

    std::wstring cmdA =
        L"ffprobe -v error "
        L"-select_streams a:0 "
        L"-show_entries stream=codec_name "
        L"-of default=noprint_wrappers=1 \"";
    cmdA += path;
    cmdA += L"\"";

    std::vector<std::string> linesA;
    if (RunFfprobeCommand(cmdA, linesA)) {
        std::string codecA;
        for (const auto& line : linesA) {
            if (line.empty()) continue;
            const char* kCodec = "codec_name=";
            const size_t codecLen = sizeof("codec_name=") - 1;
            if (line.size() >= codecLen && line.compare(0, codecLen, kCodec) == 0) {
                codecA = line.substr(codecLen);
                break;
            }
        }
        if (!codecA.empty()) {
            outAudioCodec.assign(codecA.begin(), codecA.end());
            gotA = true;
        }
    }

    return gotV || gotA;
}

static void ShowCurrentVideoProperties() {
    if (!g_inPlayback || g_playlist.empty()) {
        MessageBoxW(g_hwndMain, L"No video is currently playing.",
            L"Video properties", MB_OK);
        return;
    }

    const std::wstring& full = g_playlist[g_playlistIndex];

    int wShell = 0, hShell = 0;
    ULONGLONG durDummy = 0;
    GetVideoPropsFastCached(full, wShell, hShell, durDummy);

    int w = wShell;
    int h = hShell;
    std::wstring vCodec, aCodec;

    bool wasPlaying = (g_mp && libvlc_media_player_is_playing(g_mp) > 0);
    if (g_mp && wasPlaying) {
        libvlc_media_player_set_pause(g_mp, 1);
    }

    bool okFF = false;
    if (g_cfg.ffprobeAvailable) {
        LogLine(L"ffprobe: querying \"%s\"", full.c_str());
        okFF = GetMediaInfoFromFfprobe(full, w, h, vCodec, aCodec);
        LogLine(L"ffprobe result ok=%d w=%d h=%d vCodec=\"%s\" aCodec=\"%s\"",
            okFF ? 1 : 0, w, h, vCodec.c_str(), aCodec.c_str());
    }

    if (w <= 0) w = wShell;
    if (h <= 0) h = hShell;

    std::wstring msg = L"File: ";
    msg += full;
    msg += L"\n\n";

    if (w > 0 && h > 0) {
        wchar_t buf[64];
        swprintf_s(buf, L"%d x %d", w, h);
        msg += L"Resolution: ";
        msg += buf;
        msg += L"\n";
    }
    else {
        msg += L"Resolution: (unknown)\n";
    }

    msg += L"Video codec: ";
    msg += (vCodec.empty() ? L"(unknown)" : vCodec);
    msg += L"\n";

    msg += L"Audio codec: ";
    msg += (aCodec.empty() ? L"(unknown)" : aCodec);
    msg += L"\n";

    if (g_cfg.ffprobeAvailable && !okFF) {
        msg += L"\nNote: ffprobe.exe did not return information.";
    }
    else if (!g_cfg.ffprobeAvailable) {
        msg += L"\nNote: ffprobe-based details are disabled in browse.ini.";
    }

    MessageBoxW(g_hwndMain, msg.c_str(), L"Video properties", MB_OK);

    if (g_mp && wasPlaying) {
        libvlc_media_player_set_pause(g_mp, 0);
    }
}

// ----------------------------- Config from INI

static void LoadConfigFromIni() {
    wchar_t exePath[MAX_PATH] = {};
    if (!GetModuleFileNameW(NULL, exePath, MAX_PATH)) return;
    PathRemoveFileSpecW(exePath);

    std::wstring iniPath = exePath;
    iniPath += L"\\browse.ini";  // reuse same INI name

    std::wifstream in(iniPath);
    if (!in) return;

    std::wstring line;
    while (std::getline(in, line)) {
        line = Trim(line);
        if (line.empty()) continue;
        if (line[0] == L';' || line[0] == L'#') continue;
        if (line.front() == L'[' && line.back() == L']') continue;

        size_t semi = line.find(L';');
        if (semi != std::wstring::npos) {
            line = Trim(line.substr(0, semi));
            if (line.empty()) continue;
        }

        size_t eq = line.find(L'=');
        if (eq == std::wstring::npos) continue;

        std::wstring key = Trim(line.substr(0, eq));
        std::wstring val = Trim(line.substr(eq + 1));
        key = ToLower(key);

        if (key == L"upscaledirectory") {
            g_cfg.upscaleDirectory = val;
            if (!g_cfg.upscaleDirectory.empty()) {
                g_cfg.upscaleDirectory = EnsureSlash(g_cfg.upscaleDirectory);
            }
        }
        else if (key == L"ffmpegavailable") {
            std::wstring v = ToLower(val);
            g_cfg.ffmpegAvailable =
                (v == L"1" || v == L"true" || v == L"yes" || v == L"on" || v == L"y");
        }
        else if (key == L"loggingenabled") {
            std::wstring v = ToLower(val);
            g_cfg.loggingEnabled =
                (v == L"1" || v == L"true" || v == L"yes" || v == L"on" || v == L"y");
        }
        else if (key == L"loggingpath") {
            g_cfg.loggingPath = val;
        }
        else if (key == L"ffprobeavailable") {
            std::wstring v = ToLower(val);
            g_cfg.ffprobeAvailable =
                (v == L"1" || v == L"true" || v == L"yes" || v == L"on" || v == L"y");
        }
        else if (key == L"username") {
            g_cfg.netUsername = val;
        }
        else if (key == L"password") {
            g_cfg.netPassword = val;
        }

    }
    InitLoggingFromConfig();
    if (g_cfg.loggingEnabled) {
        LogLine(L"Config: upscale=\"%s\" ffmpeg=%d ffprobe=%d loggingPath=\"%s\"",
            g_cfg.upscaleDirectory.c_str(),
            g_cfg.ffmpegAvailable ? 1 : 0,
            g_cfg.ffprobeAvailable ? 1 : 0,
            g_cfg.loggingPath.c_str());
    }
}

// ----------------------------- Help (simplified)

static void ShowHelp() {
    bool wasPlaying = false;
    if (g_mp) {
        wasPlaying = (libvlc_media_player_is_playing(g_mp) > 0);
        if (wasPlaying) libvlc_media_player_set_pause(g_mp, 1);
    }

    std::wstring msg;
    msg += L"Browse - Help\n\n";

    msg += L"BROWSING\n"
        L"  Enter / Double-click : Open folder / Open file\n"
        L"                         (video files play in the built-in player)\n"
        L"  Left / Backspace     : Up one folder (from drive root -> drives)\n"
        L"  Column header click  : Sort by column (folders always first)\n\n";

    msg += L"FILES & FOLDERS\n"
        L"  F2                   : Rename selected file or folder\n"
        L"  Ctrl+A               : Select all items\n"
        L"  Ctrl+C / Ctrl+X      : Copy / Cut selected files and folders\n"
        L"  Ctrl+V               : Paste into current folder\n"
        L"  Del                  : Delete selected items (permanently)\n"
        L"  Right-click          : Context menu (Open, Play video, Rename, Cut/Copy/Paste, Delete)\n\n";

    msg += L"VIDEO PLAYBACK\n"
        L"  Enter                : Toggle fullscreen\n"
        L"  Esc                  : Exit playback\n"
        L"  Space / Tab          : Pause / Resume\n"
        L"  Left / Right         : Seek -/+10s  (Shift+Left/Right: -/+60s)\n"
        L"  Ctrl+Left / Ctrl+Right : Previous / Next in playlist\n"
        L"  Up / Down            : Volume +/-5\n"
        L"  Ctrl+P               : Show video properties\n\n";

    msg += L"WINDOW MANAGEMENT\n"
        L"  Win+D               : Show Windows desktop\n\n";

    msg += L"NETWORK DRIVES\n"
        L"  Right-click empty area in the list:\n"
        L"      Map Network Drive...\n"
        L"      Disconnect Network Drive...\n";

    MessageBoxW(g_hwndMain, msg.c_str(), L"Browse - Help", MB_OK);

    if (g_mp && wasPlaying) {
        libvlc_media_player_set_pause(g_mp, 0);
    }
}

// ----------------------------- ListView helpers

static LRESULT HandleListCustomDraw(NMLVCUSTOMDRAW* cd) {
    switch (cd->nmcd.dwDrawStage) {
    case CDDS_PREPAINT:
        return CDRF_NOTIFYITEMDRAW;

    case CDDS_ITEMPREPAINT:
        return CDRF_NOTIFYSUBITEMDRAW;

    case CDDS_ITEMPREPAINT | CDDS_SUBITEM: {
        int idx = (int)cd->nmcd.dwItemSpec;
        if (g_view == ViewKind::Drives && idx >= 0 && idx < (int)g_rows.size()) {
            if (g_rows[idx].isBrokenNetDrive) {
                cd->clrText = RGB(200, 0, 0);
            }
        }
        return CDRF_DODEFAULT;
    }
    }
    return CDRF_DODEFAULT;
}


static void LV_ResetColumns() {
    ListView_DeleteAllItems(g_hwndList);
    while (ListView_DeleteColumn(g_hwndList, 0)) {}

    LVCOLUMNW c; ZeroMemory(&c, sizeof(c));
    c.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;

    c.pszText = const_cast<wchar_t*>(L"Name");       c.cx = 740; c.iSubItem = 0; ListView_InsertColumn(g_hwndList, 0, &c);
    c.pszText = const_cast<wchar_t*>(L"Type");       c.cx = 80;  c.iSubItem = 1; ListView_InsertColumn(g_hwndList, 1, &c);
    c.pszText = const_cast<wchar_t*>(L"Size");       c.cx = 120; c.iSubItem = 2; ListView_InsertColumn(g_hwndList, 2, &c);
    c.pszText = const_cast<wchar_t*>(L"Modified");   c.cx = 240; c.iSubItem = 3; ListView_InsertColumn(g_hwndList, 3, &c);
    c.pszText = const_cast<wchar_t*>(L"Resolution"); c.cx = 140; c.iSubItem = 4; ListView_InsertColumn(g_hwndList, 4, &c);
    c.pszText = const_cast<wchar_t*>(L"Duration");   c.cx = 140; c.iSubItem = 5; ListView_InsertColumn(g_hwndList, 5, &c);
}

static void LV_Add(int rowIndex, const Row& r) {
    LVITEMW it; ZeroMemory(&it, sizeof(it));
    it.mask = LVIF_TEXT | LVIF_PARAM;
    it.iItem = rowIndex;
    it.pszText = const_cast<wchar_t*>(r.name.c_str());
    it.lParam = rowIndex;
    ListView_InsertItem(g_hwndList, &it);

    const wchar_t* typeText = L"Folder";
    std::wstring typeBuf;

    if (!r.isDir) {
//        if (IsVideoFile(r.full)) {
//            typeText = L"Video";
//        }
//        else {
            const wchar_t* ext = PathFindExtensionW(r.full.c_str());
            if (ext && *ext) {
                typeBuf = ext;
                typeText = typeBuf.c_str();
            }
            else {
                typeText = L"File";
            }
//        }
    }

    ListView_SetItemText(g_hwndList, rowIndex, 1, const_cast<wchar_t*>(typeText));

    if (!r.isDir) {
        std::wstring s = FormatSize(r.size);
        ListView_SetItemText(g_hwndList, rowIndex, 2, const_cast<wchar_t*>(s.c_str()));
    }
    if (r.modified.dwLowDateTime || r.modified.dwHighDateTime) {
        std::wstring m = FormatFileTime(r.modified);
        ListView_SetItemText(g_hwndList, rowIndex, 3, const_cast<wchar_t*>(m.c_str()));
    }
    if (!r.isDir && (r.vW > 0 || r.vH > 0)) {
        wchar_t buf[64];
        swprintf_s(buf, L"%dx%d", r.vW, r.vH);
        ListView_SetItemText(g_hwndList, rowIndex, 4, buf);
    }
    if (!r.isDir && r.vDur100ns > 0) {
        std::wstring ds = FormatDuration100ns(r.vDur100ns);
        ListView_SetItemText(g_hwndList, rowIndex, 5, const_cast<wchar_t*>(ds.c_str()));
    }
}

static void LV_Rebuild() {
    ListView_DeleteAllItems(g_hwndList);
    for (int i = 0; i < (int)g_rows.size(); ++i) LV_Add(i, g_rows[i]);
}

// ----------------------------- Sorting

static void SortRows(int col, bool asc) {
    g_sortCol = col; g_sortAsc = asc;
    std::sort(g_rows.begin(), g_rows.end(),
        [col, asc](const Row& A, const Row& B) {
            if (A.isDir != B.isDir) return A.isDir && !B.isDir; // dirs first
            switch (col) {
            case 0:
                return asc ? (_wcsicmp(A.name.c_str(), B.name.c_str()) < 0)
                    : (_wcsicmp(A.name.c_str(), B.name.c_str()) > 0);
            case 1: {
                auto typeTextForSort = [](const Row& r) -> const wchar_t* {
                    if (r.isDir) return L"Folder";

                    // Match what you display in LV_Add()
                    if (IsVideoFile(r.full)) return L"Video";

                    const wchar_t* ext = PathFindExtensionW(r.full.c_str());
                    if (ext && *ext) {
                        // Optional: skip the dot for nicer ordering (".txt" -> "txt")
                        return (ext[0] == L'.' && ext[1]) ? (ext + 1) : ext;
                    }
                    return L"File";
                    };

                const wchar_t* ta = typeTextForSort(A);
                const wchar_t* tb = typeTextForSort(B);

                int c = _wcsicmp(ta, tb);
                if (c != 0) return asc ? (c < 0) : (c > 0);

                // Tie-breaker: keep stable ordering by name (like your other columns)
                return _wcsicmp(A.name.c_str(), B.name.c_str()) < 0;
            }
            case 2:
                if (A.size != B.size) return asc ? (A.size < B.size) : (A.size > B.size);
                return _wcsicmp(A.name.c_str(), B.name.c_str()) < 0;
            case 3: {
                ULONGLONG a = ((ULONGLONG)A.modified.dwHighDateTime << 32) | A.modified.dwLowDateTime;
                ULONGLONG b = ((ULONGLONG)B.modified.dwHighDateTime << 32) | B.modified.dwLowDateTime;
                if (a != b) return asc ? (a < b) : (a > b);
                return _wcsicmp(A.name.c_str(), B.name.c_str()) < 0;
            }
            case 4: {
                ULONGLONG aa = (ULONGLONG)A.vW * (ULONGLONG)A.vH;
                ULONGLONG bb = (ULONGLONG)B.vW * (ULONGLONG)B.vH;
                if (aa != bb) return asc ? (aa < bb) : (aa > bb);
                if (A.vW != B.vW) return asc ? (A.vW < B.vW) : (A.vW > B.vW);
                return _wcsicmp(A.name.c_str(), B.name.c_str()) < 0;
            }
            case 5:
                if (A.vDur100ns != B.vDur100ns) return asc ? (A.vDur100ns < B.vDur100ns) : (A.vDur100ns > B.vDur100ns);
                return _wcsicmp(A.name.c_str(), B.name.c_str()) < 0;
            default:
                return _wcsicmp(A.name.c_str(), B.name.c_str()) < 0;
            }
        });
    LV_Rebuild();
}

// ----------------------------- Async metadata worker

static DWORD WINAPI MetaThreadProc(LPVOID) {
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    const uint32_t myGen = g_metaGen.load(std::memory_order_relaxed);

    for (;;) {
        std::wstring path;
        EnterCriticalSection(&g_metaLock);
        if (!g_metaTodoPaths.empty()) {
            path = g_metaTodoPaths.back();
            g_metaTodoPaths.pop_back();
        }
        LeaveCriticalSection(&g_metaLock);

        if (path.empty()) break;
        if (myGen != g_metaGen.load(std::memory_order_relaxed)) break;

        int w = 0, h = 0; ULONGLONG d = 0;
        GetVideoProps(path, w, h, d);
        MetaResult* r = new MetaResult{ path, w, h, d, myGen };
        PostMessageW(g_hwndMain, WM_APP_META, 0, (LPARAM)r);
    }

    CoUninitialize();
    return 0;
}

static void StartMetaWorker() {
    if (g_metaThread) { CloseHandle(g_metaThread); g_metaThread = NULL; }
    g_metaThread = CreateThread(NULL, 0, MetaThreadProc, NULL, 0, NULL);
}

static void CancelMetaWorkAndClearTodo() {
    g_metaGen.fetch_add(1, std::memory_order_relaxed);
    EnterCriticalSection(&g_metaLock);
    g_metaTodoPaths.clear();
    LeaveCriticalSection(&g_metaLock);
}

static void QueueMissingPropsAndKickWorker() {
    EnterCriticalSection(&g_metaLock);
    for (const auto& r : g_rows) {
        if (!r.isDir && r.vW == 0 && r.vH == 0 && r.vDur100ns == 0 && IsVideoFile(r.full)) {
            g_metaTodoPaths.push_back(r.full);
        }
    }
    LeaveCriticalSection(&g_metaLock);
    if (!g_metaTodoPaths.empty()) StartMetaWorker();
}

// ----------------------------- Populate views

static void ShowDrives() {
    CancelMetaWorkAndClearTodo();

    g_view = ViewKind::Drives;
    g_folder.clear();
    g_rows.clear();

    SendMessageW(g_hwndList, WM_SETREDRAW, FALSE, 0);
    LV_ResetColumns();

    DWORD connectedMask = GetConnectedNetDriveMask();

    DWORD mask = GetLogicalDrives();
    for (int i = 0; i < 26; ++i) {
        if (!(mask & (1u << i))) continue;

        wchar_t letter = (wchar_t)(L'A' + i);

        wchar_t root[4] = { letter, L':', L'\\', 0 };

        Row r;
        r.full = root;
        r.isDir = true;

        // default display
        r.name = root;

        // Detect persistent mapping and whether it's currently connected ("net use" OK)
        std::wstring remote;
        bool hasPersistent = GetPersistentMappedRemotePath(letter, remote);
        bool isConnected = ((connectedMask & (1u << i)) != 0);

        if (hasPersistent && !isConnected) {
            // This is your "mapped but not OK" drive
            r.isBrokenNetDrive = true;
            r.netRemote = remote;

            // Display as "X:" (no trailing slash) per your request
            r.name.clear();
            r.name.push_back(letter);
            r.name.push_back(L':');
        }

        g_rows.push_back(std::move(r));
    }

    SortRows(0, true);
    SendMessageW(g_hwndList, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(g_hwndList, NULL, TRUE);

    SetTitleFolderOrDrives();
}

// Folder view: shows ALL files, not only videos.
// Folder view: shows ALL files, not only videos.
static void ShowFolder(std::wstring abs) {
    CancelMetaWorkAndClearTodo();

    if (abs.size() == 2 && abs[1] == L':') abs += L'\\';
    abs = EnsureSlash(abs);
    g_view = ViewKind::Folder;
    g_folder = abs;
    g_rows.clear();

    // ------------------------------------------------------------
    // NEW: title update immediately + 1-char busy animation setup
    // Sequence: ' ' '.' 'o' 'O' ' ' ...
    // ------------------------------------------------------------
    g_loadingFolder = true;

    // (Optional but recommended) prevent user actions during load
    if (g_hwndList) EnableWindow(g_hwndList, FALSE);

    std::wstring animTitle = L"Browse - ";
    animTitle += EnsureSlash(g_folder);
    animTitle.push_back(L' '); // start on SPC (0x20)
    SetWindowTextW(g_hwndMain, animTitle.c_str());
    UpdateWindow(g_hwndMain);

    static const wchar_t kAnimFrames[4] = { L' ', L'.', L'o', L'O' };
    DWORD lastAnimTick = GetTickCount();
    int   animFrame = 1; // after ~1 second show '.' first

    SendMessageW(g_hwndList, WM_SETREDRAW, FALSE, 0);
    LV_ResetColumns();

    WIN32_FIND_DATAW fd; ZeroMemory(&fd, sizeof(fd));
    HANDLE h = FindFirstFileExW((abs + L"*").c_str(),
        FindExInfoBasic,
        &fd,
        FindExSearchNameMatch,
        NULL,
        FIND_FIRST_EX_LARGE_FETCH);

    if (h == INVALID_HANDLE_VALUE) {
        SendMessageW(g_hwndList, WM_SETREDRAW, TRUE, 0);
        InvalidateRect(g_hwndList, NULL, TRUE);

        if (g_hwndList) EnableWindow(g_hwndList, TRUE);
        g_loadingFolder = false;

        SetTitleFolderOrDrives(); // ends cleanly (no spinner char)
        return;
    }

    std::vector<Row> dirs, files;
    do {
        if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;

        Row r;
        r.name = fd.cFileName;
        r.full = abs + fd.cFileName;
        r.isDir = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        r.modified = fd.ftLastWriteTime;

        if (r.isDir) {
            dirs.push_back(r);
        }
        else {
            ULARGE_INTEGER uli;
            uli.HighPart = fd.nFileSizeHigh;
            uli.LowPart = fd.nFileSizeLow;
            r.size = uli.QuadPart;

            if (IsVideoFile(r.full)) {
                if (!GetVideoPropsFastCached(r.full, r.vW, r.vH, r.vDur100ns)) {
                    r.vW = r.vH = 0;
                    r.vDur100ns = 0;
                }
            }
            files.push_back(r);
        }

        // ------------------------------------------------------------
        // NEW: keep UI responsive + tick spinner once per second
        // ------------------------------------------------------------
        PumpMessagesThrottled(50);

        DWORD now = GetTickCount();
        if (now - lastAnimTick >= 1000) {
            animTitle.back() = kAnimFrames[animFrame & 3];
            ++animFrame;
            SetWindowTextW(g_hwndMain, animTitle.c_str());
            lastAnimTick = now;
        }

    } while (FindNextFileW(h, &fd));
    FindClose(h);

    g_rows.reserve(dirs.size() + files.size());
    g_rows.insert(g_rows.end(), dirs.begin(), dirs.end());
    g_rows.insert(g_rows.end(), files.begin(), files.end());

    SortRows(g_sortCol, g_sortAsc);

    SendMessageW(g_hwndList, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(g_hwndList, NULL, TRUE);

    QueueMissingPropsAndKickWorker();

    if (g_hwndList) EnableWindow(g_hwndList, TRUE);
    g_loadingFolder = false;

    // End cleanly (spinner removed / ends effectively on SPC)
    SetTitleFolderOrDrives();
}

// ----------------------------- Search (videos only, as original)

static bool NameContainsAllTerms(const std::wstring& full,
    const std::vector<std::wstring>& termsLower) {
    const wchar_t* base = wcsrchr(full.c_str(), L'\\'); base = base ? base + 1 : full.c_str();
    std::wstring bl = ToLower(base);
    for (size_t i = 0; i < termsLower.size(); ++i)
        if (bl.find(termsLower[i]) == std::wstring::npos) return false;
    return true;
}

static void SearchRecurseFolder(const std::wstring& folder,
    const std::vector<std::wstring>& terms,
    std::vector<Row>& out) {
    SetTitleSearchingFolder(folder);

    std::wstring pat = EnsureSlash(folder) + L"*";
    WIN32_FIND_DATAW fd; ZeroMemory(&fd, sizeof(fd));
    HANDLE h = FindFirstFileExW(pat.c_str(), FindExInfoBasic, &fd,
        FindExSearchNameMatch, NULL, FIND_FIRST_EX_LARGE_FETCH);
    if (h == INVALID_HANDLE_VALUE) return;

    do {
        if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;

        bool isDir = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        std::wstring full = EnsureSlash(folder) + fd.cFileName;

        if (isDir) {
            if (fd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) continue;
            SearchRecurseFolder(full, terms, out);
        }
        else if (IsVideoFile(full)) {
            if (NameContainsAllTerms(full, terms)) {
                Row r;
                r.name = full;
                r.full = full;
                r.isDir = false;
                r.modified = fd.ftLastWriteTime;
                ULARGE_INTEGER uli;
                uli.HighPart = fd.nFileSizeHigh;
                uli.LowPart = fd.nFileSizeLow;
                r.size = uli.QuadPart;

                GetVideoPropsFastCached(r.full, r.vW, r.vH, r.vDur100ns);
                out.push_back(r);
            }
        }
    } while (FindNextFileW(h, &fd));

    FindClose(h);
}

static void RunSearchFromOrigin(std::vector<Row>& outResults) {
    outResults.clear();

    if (g_search.useExplicitScope) {
        for (const auto& file : g_search.explicitFiles) {
            if (!IsVideoFile(file)) continue;
            if (!NameContainsAllTerms(file, g_search.termsLower)) continue;

            WIN32_FILE_ATTRIBUTE_DATA fad{};
            if (GetFileAttributesExW(file.c_str(), GetFileExInfoStandard, &fad) &&
                (fad.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {

                Row r;
                r.name = file;
                r.full = file;
                r.isDir = false;
                r.modified = fad.ftLastWriteTime;

                ULARGE_INTEGER uli{};
                uli.HighPart = fad.nFileSizeHigh;
                uli.LowPart = fad.nFileSizeLow;
                r.size = uli.QuadPart;

                GetVideoPropsFastCached(r.full, r.vW, r.vH, r.vDur100ns);
                outResults.push_back(std::move(r));
            }
        }

        for (const auto& folder : g_search.explicitFolders) {
            SetTitleSearchingFolder(folder);
            SearchRecurseFolder(folder, g_search.termsLower, outResults);
        }
        return;
    }

    if (g_search.originView == ViewKind::Drives) {
        DWORD mask = GetLogicalDrives();
        for (int i = 0; i < 26; ++i) {
            if (!(mask & (1u << i))) continue;
            wchar_t root[4] = { wchar_t(L'A' + i), L':', L'\\', 0 };
            SetTitleSearchingFolder(root);
            SearchRecurseFolder(root, g_search.termsLower, outResults);
        }
    }
    else {
        SetTitleSearchingFolder(g_search.originFolder);
        SearchRecurseFolder(g_search.originFolder, g_search.termsLower, outResults);
    }
}

static void ShowSearchResults(const std::vector<Row>& results) {
    CancelMetaWorkAndClearTodo();

    g_view = ViewKind::Search;
    g_rows = results;

    SendMessageW(g_hwndList, WM_SETREDRAW, FALSE, 0);
    LV_ResetColumns();
    SortRows(g_sortCol, g_sortAsc);
    SendMessageW(g_hwndList, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(g_hwndList, NULL, TRUE);

    std::wstring t = L"Browse - Search - " + JoinTermsForTitle();
    wchar_t buf[64]; swprintf_s(buf, L" - %zu file(s)", g_rows.size());
    t += buf;
    SetWindowTextW(g_hwndMain, t.c_str());

    QueueMissingPropsAndKickWorker();
}

static void ExitSearchToOrigin() {
    if (!g_search.active) return;
    if (g_search.originView == ViewKind::Drives) ShowDrives();
    else ShowFolder(g_search.originFolder);
    g_search = SearchState();
}

// ----------------------------- File operations

static std::wstring UniqueName(const std::wstring& folder,
    const std::wstring& base,
    const std::wstring& ext) {
    std::wstring f = EnsureSlash(folder);
    std::wstring target = f + base + ext;
    if (!PathFileExistsW(target.c_str())) return target;
    for (int i = 1; i < 10000; ++i) {
        wchar_t buf[32];
        swprintf_s(buf, L" (%d)", i);
        std::wstring t = f + base + buf + ext;
        if (!PathFileExistsW(t.c_str())) return t;
    }
    return target;
}

// Clipboard: now includes files AND directories
static void Browser_CopySelectedToClipboard(ClipMode mode) {
    g_clipFiles.clear();
    g_clipMode = ClipMode::None;
    if (g_view == ViewKind::Drives) return;

    std::vector<int> selectedIdx;
    int idx = -1;
    while ((idx = ListView_GetNextItem(g_hwndList, idx, LVNI_SELECTED)) != -1) {
        if (idx < 0 || idx >= (int)g_rows.size()) continue;
        const Row& r = g_rows[idx];
        g_clipFiles.push_back(r.full);
        selectedIdx.push_back(idx);
    }
    if (g_clipFiles.empty()) return;

    g_clipMode = mode;

    if (mode == ClipMode::Move) {
        SendMessageW(g_hwndList, WM_SETREDRAW, FALSE, 0);
        std::sort(selectedIdx.begin(), selectedIdx.end());
        for (int n = (int)selectedIdx.size() - 1; n >= 0; --n) {
            int rIdx = selectedIdx[n];
            if (rIdx >= 0 && rIdx < (int)g_rows.size()) {
                g_rows.erase(g_rows.begin() + rIdx);
                ListView_DeleteItem(g_hwndList, rIdx);
            }
        }
        SendMessageW(g_hwndList, WM_SETREDRAW, TRUE, 0);
        InvalidateRect(g_hwndList, NULL, TRUE);
    }
    // NEW: also publish to the system clipboard as CF_HDROP so that
    // Explorer / Remote Desktop etc. can see the file list.
    SetClipboardFileDrop(g_clipFiles, mode);

}

// ----------------------------- DPI helpers

typedef UINT(WINAPI* GetDpiForWindow_t)(HWND);
static int DpiScale(int px) {
    UINT dpi = 96;
    HMODULE m = GetModuleHandleW(L"user32.dll");
    if (m) {
        GetDpiForWindow_t p = (GetDpiForWindow_t)GetProcAddress(m, "GetDpiForWindow");
        if (p) dpi = p(g_hwndMain ? g_hwndMain : GetDesktopWindow());
    }
    return MulDiv(px, dpi, 96);
}

// ---------- Operation (copy/move) sub-modal window + cancellable copy support

struct OpUI {
    HWND hwnd{}, hText{}, hCancel{};
    std::atomic<bool> cancel{ false };
    BOOL* pCancelFlag{ nullptr };
} g_op;

static LRESULT CALLBACK OpProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    switch (m) {
    case WM_CREATE: {
        HFONT hf = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        RECT rc{}; GetClientRect(h, &rc);
        int margin = DpiScale(12);
        int btnW = DpiScale(100), btnH = DpiScale(28);

        g_op.hText = CreateWindowExW(
            WS_EX_TRANSPARENT, L"STATIC", L"",
            WS_CHILD | WS_VISIBLE | SS_LEFT,
            margin, margin, rc.right - 2 * margin, DpiScale(32),
            h, (HMENU)101, g_hInst, NULL);
        SendMessageW(g_op.hText, WM_SETFONT, (WPARAM)hf, TRUE);

        g_op.hCancel = CreateWindowExW(
            0, L"BUTTON", L"Cancel",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            rc.right - margin - btnW, rc.bottom - margin - btnH,
            btnW, btnH, h, (HMENU)IDCANCEL, g_hInst, NULL);
        SendMessageW(g_op.hCancel, WM_SETFONT, (WPARAM)hf, TRUE);
        return 0;
    }
    case WM_SIZE: {
        RECT rc{}; GetClientRect(h, &rc);
        int margin = DpiScale(12);
        int btnW = DpiScale(100), btnH = DpiScale(28);
        if (g_op.hText)
            MoveWindow(g_op.hText, margin, margin,
                rc.right - 2 * margin, DpiScale(32), TRUE);
        if (g_op.hCancel)
            MoveWindow(g_op.hCancel,
                rc.right - margin - btnW,
                rc.bottom - margin - btnH,
                btnW, btnH, TRUE);
        return 0;
    }
    case WM_COMMAND:
        if (LOWORD(w) == IDCANCEL) {
            if (g_op.pCancelFlag) *g_op.pCancelFlag = TRUE;
            g_op.cancel.store(true, std::memory_order_relaxed);
            DestroyWindow(h);
            return 0;
        }
        break;
    case WM_CLOSE:
        if (g_op.pCancelFlag) *g_op.pCancelFlag = TRUE;
        g_op.cancel.store(true, std::memory_order_relaxed);
        DestroyWindow(h);
        return 0;
    case WM_DESTROY:
        g_op.hwnd = NULL;
        return 0;
    }
    return DefWindowProcW(h, m, w, l);
}

static void EnsureOpWndClass() {
    static bool reg = false;
    if (!reg) {
        WNDCLASSW wc{};
        wc.lpfnWndProc = OpProc;
        wc.hInstance = g_hInst;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"OpProgressClass";
        RegisterClassW(&wc);
        reg = true;
    }
}

static HWND CreateOpWindow(const wchar_t* title) {
    EnsureOpWndClass();

    MONITORINFO mi{ sizeof(mi) };
    HMONITOR hm = MonitorFromWindow(g_hwndMain, MONITOR_DEFAULTTONEAREST);
    GetMonitorInfoW(hm, &mi);
    RECT wa = mi.rcWork;

    int W = DpiScale(560), H = DpiScale(110);
    int X = wa.left + ((wa.right - wa.left) - W) / 2;
    int Y = wa.top + ((wa.bottom - wa.top) - H) / 2;

    HWND h = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
        L"OpProgressClass", title,
        WS_POPUPWINDOW | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        X, Y, W, H, g_hwndMain, NULL, g_hInst, NULL);
    g_op.hwnd = h;
    return h;
}

static DWORD CALLBACK CopyProgressThunk(LARGE_INTEGER, LARGE_INTEGER,
    LARGE_INTEGER, LARGE_INTEGER,
    DWORD, DWORD, HANDLE, HANDLE, LPVOID lpData) {
    OpUI* op = reinterpret_cast<OpUI*>(lpData);
    PumpMessagesThrottled(10);
    if (op && op->cancel.load(std::memory_order_relaxed)) {
        return PROGRESS_CANCEL;
    }
    return PROGRESS_CONTINUE;
}

static bool SameVolume(const std::wstring& a, const std::wstring& b) {
    wchar_t va[MAX_PATH]{}, vb[MAX_PATH]{};
    if (!GetVolumePathNameW(a.c_str(), va, MAX_PATH)) return false;
    if (!GetVolumePathNameW(b.c_str(), vb, MAX_PATH)) return false;
    return _wcsicmp(va, vb) == 0;
}

// Recursive directory helpers

// Recursive delete of a directory tree.
// - Normal directories: recurse and delete contents, then remove the dir.
// - Reparse-point dirs (junctions/symlink dirs): do NOT recurse into target,
//   just remove the link itself (RemoveDirectoryW).
static bool DeleteDirectoryTree(const std::wstring& path)
{
    // Normalize: ensure trailing slash for enumeration
    std::wstring dir = EnsureSlash(path);

    // Enumerate children
    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW((dir + L"*").c_str(), &fd);
    if (h != INVALID_HANDLE_VALUE) {
        do {
            if (wcscmp(fd.cFileName, L".") == 0 ||
                wcscmp(fd.cFileName, L"..") == 0)
                continue;

            std::wstring child = dir + fd.cFileName;
            bool isDir = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            bool isReparse = (fd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;

            if (isDir) {
                if (isReparse) {
                    // Junction / symlink folder: remove the link itself, don't recurse
                    ClearReadonlyAndSystem(child);
                    if (!RemoveDirectoryW(child.c_str())) {
                        MoveFileExW(child.c_str(), NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
                    }
                }
                else {
                    // Normal directory: recurse into it
                    DeleteDirectoryTree(child);
                }
            }
            else {
                // Normal file or reparse file (symlink to file)
                ClearReadonlyAndSystem(child);
                if (!DeleteFileW(child.c_str())) {
                    MoveFileExW(child.c_str(), NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
                }
            }
        } while (FindNextFileW(h, &fd));
        FindClose(h);
    }

    // Finally delete this directory itself
    // (clear read-only/system attr first so RemoveDirectoryW succeeds)
    std::wstring dirNoSlash = path;
    // Strip trailing slash for safety (RemoveDirectory is happier with "C:\foo" than "C:\foo\")
    while (!dirNoSlash.empty() &&
        (dirNoSlash.back() == L'\\' || dirNoSlash.back() == L'/') &&
        !IsDriveRoot(dirNoSlash)) {
        dirNoSlash.pop_back();
    }

    ClearReadonlyAndSystem(dirNoSlash);

    if (!RemoveDirectoryW(dirNoSlash.c_str())) {
        MoveFileExW(dirNoSlash.c_str(), NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
        return false;
    }
    return true;
}

static bool CopyDirectoryTree(const std::wstring& srcDirIn,
    const std::wstring& dstDirIn,
    OpUI* op,
    BOOL* uiCancel) {
    std::wstring srcDir = EnsureSlash(srcDirIn);
    std::wstring dstDir = EnsureSlash(dstDirIn);

    if (!CreateDirectoryW(dstDir.c_str(), NULL)) {
        DWORD e = GetLastError();
        if (e != ERROR_ALREADY_EXISTS) {
            return false;
        }
    }

    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW((srcDir + L"*").c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) {
        return true;
    }

    do {
        if (wcscmp(fd.cFileName, L".") == 0 ||
            wcscmp(fd.cFileName, L"..") == 0)
            continue;

        if (op && op->cancel.load(std::memory_order_relaxed)) {
            FindClose(h);
            return false;
        }
        if (uiCancel && *uiCancel) {
            FindClose(h);
            return false;
        }

        std::wstring srcPath = srcDir + fd.cFileName;
        std::wstring dstPath = dstDir + fd.cFileName;

        bool isDir = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        bool isReparse = (fd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;

        if (isDir && !isReparse) {
            if (!CopyDirectoryTree(srcPath, dstPath, op, uiCancel)) {
                FindClose(h);
                return false;
            }
        }
        else {
            BOOL ok = CopyFileExW(
                srcPath.c_str(), dstPath.c_str(),
                CopyProgressThunk, op, uiCancel, 0);
            if (!ok) {
                FindClose(h);
                return false;
            }
        }
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    return true;
}

// Run the copy/move with UI and cancellation (files + directories)
// NOTE: requires: enum class OpResult { Success, Cancelled, Failed };

static OpResult RunClipboardOperationWithUI(const std::wstring& dstFolder)
{
    if (g_clipMode == ClipMode::None || g_clipFiles.empty())
        return OpResult::Failed;

    const bool   isCopy = (g_clipMode == ClipMode::Copy);
    const size_t total = g_clipFiles.size();

    bool allOk = true;
    bool cancelled = false;

    auto makeCaption = [&](size_t currentIndex) -> std::wstring {
        if (total <= 1) return std::wstring(isCopy ? L"Copying..." : L"Moving...");
        wchar_t buf[256];
        swprintf_s(buf, L"%s... %zu of %zu",
            isCopy ? L"Copying" : L"Moving",
            currentIndex, total);
        return buf;
        };

    std::wstring initialCap = makeCaption(total > 1 ? 1 : 0);
    HWND hw = CreateOpWindow(initialCap.c_str());
    if (!hw) {
        // Keep clipboard state so user can retry
        return OpResult::Failed;
    }

    HFONT hf = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
    if (g_op.hText) SendMessageW(g_op.hText, WM_SETFONT, (WPARAM)hf, TRUE);

    BOOL uiCancel = FALSE;
    g_op.pCancelFlag = &uiCancel;
    g_op.cancel.store(false, std::memory_order_relaxed);

    auto updateCaption = [&](size_t idx1based) {
        if (IsWindow(g_op.hwnd)) {
            std::wstring cap = makeCaption(idx1based);
            SetWindowTextW(hw, cap.c_str());
            UpdateWindow(hw);
        }
        };

    auto setStatusText = [&](const std::wstring& s) {
        if (IsWindow(g_op.hwnd) && g_op.hText) {
            SetWindowTextW(g_op.hText, s.c_str());
            UpdateWindow(hw);
        }
        };

    // strip trailing slashes (robust against "C:\foo\")
    auto StripTrailingSlashes = [&](std::wstring p) -> std::wstring {
        while (p.size() > 3 && (p.back() == L'\\' || p.back() == L'/')) {
            p.pop_back();
        }
        return p;
        };

    // prevent copying/moving a folder into itself / its subtree
    auto IsPrefixPathNoCase = [&](const std::wstring& parentIn, const std::wstring& childIn) -> bool {
        std::wstring parent = EnsureSlash(parentIn);
        std::wstring child = EnsureSlash(childIn);
        if (parent.size() > child.size()) return false;
        return _wcsnicmp(parent.c_str(), child.c_str(), parent.size()) == 0;
        };

    for (size_t i = 0; i < total; ++i) {
        if (g_op.cancel.load(std::memory_order_relaxed)) { cancelled = true; break; }

        std::wstring src = StripTrailingSlashes(g_clipFiles[i]);
        DWORD attrs = GetFileAttributesW(src.c_str());
        if (attrs == INVALID_FILE_ATTRIBUTES) {
            allOk = false;
            continue;
        }

        const bool isDir = (attrs & FILE_ATTRIBUTE_DIRECTORY) != 0;
        const bool isReparse = (attrs & FILE_ATTRIBUTE_REPARSE_POINT) != 0;

        // Base name (handles both '\' and '/')
        const wchar_t* base = wcsrchr(src.c_str(), L'\\');
        const wchar_t* base2 = wcsrchr(src.c_str(), L'/');
        if (!base || (base2 && base2 > base)) base = base2;
        base = base ? base + 1 : src.c_str();
        std::wstring baseName = base;

        // Build destination path WITHOUT using _wsplitpath_s for directories
        std::wstring nameBase, nameExt;
        if (isDir) {
            nameBase = baseName; // keep full folder name (including dots)
            nameExt.clear();
        }
        else {
            const wchar_t* extPtr = PathFindExtensionW(baseName.c_str());
            if (extPtr && *extPtr) {
                nameBase.assign(baseName.c_str(), extPtr - baseName.c_str());
                nameExt = extPtr; // includes '.'
            }
            else {
                nameBase = baseName;
                nameExt.clear();
            }
        }

        std::wstring dst = UniqueName(dstFolder, nameBase, nameExt);

        // Safety: don't paste a folder into itself / into its own subtree
        if (isDir) {
            if (IsPrefixPathNoCase(src, dstFolder) || IsPrefixPathNoCase(src, dst)) {
                allOk = false;
                continue;
            }
        }

        if (total > 1) updateCaption(i + 1);

        std::wstring line = (isCopy ? L"Copying " : L"Moving ");
        line += baseName;
        line += L"...";
        setStatusText(line);
        PumpMessagesThrottled(10);

        BOOL ok = FALSE;
        uiCancel = FALSE;

        if (isCopy) {
            if (isDir) {
                // If it's a reparse-point dir (junction/symlink), CopyDirectoryTree() will treat it as a file and fail.
                // We keep existing behavior: treat failure as failure.
                ok = CopyDirectoryTree(src, dst, &g_op, &uiCancel);
            }
            else {
                ok = CopyFileExW(src.c_str(), dst.c_str(),
                    CopyProgressThunk, &g_op, &uiCancel, 0);
            }
        }
        else {
            // Move
            if (SameVolume(src, dst)) {
                // Fast rename/move within same volume/share
                ok = MoveFileExW(src.c_str(), dst.c_str(),
                    MOVEFILE_REPLACE_EXISTING);
            }
            else {
                // Cross-volume: copy then delete source
                if (isDir) {
                    ok = CopyDirectoryTree(src, dst, &g_op, &uiCancel);
                    if (ok) {
                        if (!DeleteDirectoryTree(src)) {
                            // Copy succeeded but delete failed -> not a true move
                            allOk = false;
                        }
                    }
                    else {
                        // Best-effort cleanup of partial destination
                        DeleteDirectoryTree(dst);
                    }
                }
                else {
                    ok = CopyFileExW(src.c_str(), dst.c_str(),
                        CopyProgressThunk, &g_op, &uiCancel, 0);
                    if (ok) {
                        ClearReadonlyAndSystem(src);
                        if (!DeleteFileW(src.c_str())) {
                            MoveFileExW(src.c_str(), NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
                            allOk = false;
                        }
                    }
                    else {
                        DeleteFileW(dst.c_str());
                    }
                }
            }
        }

        if (!ok) {
            allOk = false;
            DWORD err = GetLastError();
            if (uiCancel || g_op.cancel.load(std::memory_order_relaxed) ||
                err == ERROR_REQUEST_ABORTED || err == ERROR_CANCELLED) {
                cancelled = true;
                break;
            }
        }

        if (g_op.cancel.load(std::memory_order_relaxed) || uiCancel) {
            cancelled = true;
            break;
        }

        setStatusText(line + L" Done");
        PumpMessagesThrottled(10);

        if (total > 1 && (i + 1) < total) updateCaption(i + 2);
    }

    // Close progress window (if user already cancelled, it may already be gone)
    if (IsWindow(g_op.hwnd)) DestroyWindow(g_op.hwnd);

    // IMPORTANT: pointer was to a local stack var
    g_op.pCancelFlag = nullptr;

    // Clear internal clipboard state (per your original design)
    g_clipFiles.clear();
    g_clipMode = ClipMode::None;

    // Refresh view if we're in a folder
    if (g_view == ViewKind::Folder) ShowFolder(g_folder);

    if (cancelled) return OpResult::Cancelled;
    return allOk ? OpResult::Success : OpResult::Failed;
}

static void Browser_PasteClipboardIntoCurrent() {
    std::wstring dstFolder;
    if (g_view == ViewKind::Folder) {
        dstFolder = g_folder;
    }
    else if (g_view == ViewKind::Search && g_search.originView == ViewKind::Folder) {
        dstFolder = g_search.originFolder;
    }
    else {
        return;
    }

    // 1) Try the real Windows clipboard (CF_HDROP) first
    std::vector<std::wstring> sysFiles;
    ClipMode sysMode = ClipMode::Copy;
    if (GetClipboardFileDrop(sysFiles, &sysMode)) {
        g_clipMode = sysMode;
        g_clipFiles = std::move(sysFiles);

        OpResult r = RunClipboardOperationWithUI(dstFolder);

        // Optional: if it was a CUT and we actually succeeded, clear clipboard
        // (prevents stale "cut" source paths being pasted again)
        if (sysMode == ClipMode::Move && r == OpResult::Success) {
            if (OpenClipboard(g_hwndMain)) {
                EmptyClipboard();
                CloseClipboard();
            }
        }
        return;
    }

    // 2) Fallback to internal clipboard (Ctrl+C/X done inside this app)
    if (g_clipMode == ClipMode::None || g_clipFiles.empty()) return;

    RunClipboardOperationWithUI(dstFolder);
}

static void Browser_DeleteSelected()
{
    if (g_view == ViewKind::Drives) return;

    if (MessageBoxW(
        g_hwndMain,
        L"Delete selected items permanently?\n"
        L"(Folders will be deleted recursively.)",
        L"Confirm Delete",
        MB_YESNO | MB_DEFBUTTON2 | MB_ICONWARNING) != IDYES)
    {
        return;
    }

    // First collect all selected paths so we don't care if the view changes
    std::vector<std::wstring> toDelete;
    int idx = -1;
    while ((idx = ListView_GetNextItem(g_hwndList, idx, LVNI_SELECTED)) != -1) {
        if (idx < 0 || idx >= (int)g_rows.size()) continue;
        toDelete.push_back(g_rows[idx].full);
    }
    if (toDelete.empty()) return;

    bool anyFailed = false;

    for (const auto& path : toDelete) {
        DWORD attrs = GetFileAttributesW(path.c_str());
        if (attrs == INVALID_FILE_ATTRIBUTES)
            continue;

        if (attrs & FILE_ATTRIBUTE_DIRECTORY) {
            if (!DeleteDirectoryTree(path)) {
                anyFailed = true;
            }
        }
        else {
            ClearReadonlyAndSystem(path);
            if (!DeleteFileW(path.c_str())) {
                MoveFileExW(path.c_str(), NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
                anyFailed = true;
            }
        }
    }

    // Refresh whichever view we're in so the UI matches disk state
    if (g_view == ViewKind::Search && g_search.active) {
        std::vector<Row> res;
        RunSearchFromOrigin(res);
        ShowSearchResults(res);
    }
    else if (g_view == ViewKind::Folder) {
        ShowFolder(g_folder);
    }
    else if (g_view == ViewKind::Drives) {
        ShowDrives();
    }

    if (anyFailed) {
        MessageBoxW(
            g_hwndMain,
            L"Some items could not be deleted (locked, in use, or permission denied).\n"
            L"They may have been queued for deletion on next reboot.",
            L"Delete",
            MB_OK | MB_ICONWARNING);
    }
}

// ----------------------------- Save-As helper (used in playback rename/copy)

static bool PromptSaveAsFrom(const std::wstring& seedPath,
    std::wstring& outPath,
    const wchar_t* titleText) {
    ComPtr<IFileSaveDialog> dlg;
    if (FAILED(CoCreateInstance(CLSID_FileSaveDialog, NULL, CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&dlg))))
        return false;

    wchar_t dir[MAX_PATH] = {};
    wcscpy_s(dir, seedPath.c_str());
    PathRemoveFileSpecW(dir);
    ComPtr<IShellItem> initFolder;
    if (SUCCEEDED(SHCreateItemFromParsingName(dir, NULL,
        IID_PPV_ARGS(&initFolder))))
        dlg->SetFolder(initFolder.Get());

    const wchar_t* base = wcsrchr(seedPath.c_str(), L'\\');
    base = base ? base + 1 : seedPath.c_str();
    dlg->SetFileName(base);

    COMDLG_FILTERSPEC spec[] = { { L"All Files", L"*.*" } };
    dlg->SetFileTypes(1, spec);
    dlg->SetTitle(titleText ? titleText : L"Save As");
    dlg->SetOptions(FOS_OVERWRITEPROMPT | FOS_FORCEFILESYSTEM);

    if (FAILED(dlg->Show(g_hwndMain))) return false;

    ComPtr<IShellItem> it;
    if (FAILED(dlg->GetResult(&it))) return false;
    PWSTR psz = NULL;
    if (FAILED(it->GetDisplayName(SIGDN_FILESYSPATH, &psz))) return false;
    outPath.assign(psz);
    CoTaskMemFree(psz);
    return true;
}

// ----------------------------- Keyword / generic input dialog

struct KwCtx {
    HWND hwnd, hEdit, hOK, hCancel;
    bool accepted;
    std::wstring text;
    std::wstring label;
    std::wstring title;
    std::wstring initial;
};
static KwCtx g_kw = { 0,0,0,0,false,L"",L"",L"",L"" };

static LRESULT CALLBACK KwEditSub(HWND h, UINT m, WPARAM w, LPARAM l,
    UINT_PTR, DWORD_PTR) {
    if (m == WM_KEYDOWN) {
        if (w == VK_RETURN) {
            PostMessageW(GetParent(h), WM_COMMAND,
                MAKELONG(IDOK, BN_CLICKED),
                (LPARAM)g_kw.hOK);
            return 0;
        }
        if (w == VK_ESCAPE) {
            PostMessageW(GetParent(h), WM_COMMAND,
                MAKELONG(IDCANCEL, BN_CLICKED),
                (LPARAM)g_kw.hCancel);
            return 0;
        }
    }
    return DefSubclassProc(h, m, w, l);
}

static void EnsureKwClassRegistered() {
    static bool reg = false;
    if (!reg) {
        WNDCLASSW wc; ZeroMemory(&wc, sizeof(wc));
        wc.lpfnWndProc = [](HWND h, UINT m, WPARAM w, LPARAM l) -> LRESULT {
            switch (m) {
            case WM_CREATE: {
                HFONT hf = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
                RECT rc; GetClientRect(h, &rc);
                int margin = DpiScale(12);
                int btnW = DpiScale(90), btnH = DpiScale(28);
                int labelH = DpiScale(20);
                int editH = DpiScale(24);

                const wchar_t* labelText =
                    g_kw.label.empty() ? L"Input:" : g_kw.label.c_str();

                HWND hLbl = CreateWindowExW(
                    0, L"STATIC", labelText,
                    WS_CHILD | WS_VISIBLE,
                    margin, margin,
                    rc.right - 2 * margin, labelH,
                    h, NULL, g_hInst, NULL);
                SendMessageW(hLbl, WM_SETFONT, (WPARAM)hf, TRUE);

                g_kw.hEdit = CreateWindowExW(
                    WS_EX_CLIENTEDGE, L"EDIT", L"",
                    WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                    margin, margin + labelH + DpiScale(6),
                    rc.right - 2 * margin - (btnW + DpiScale(10)),
                    editH,
                    h, (HMENU)201, g_hInst, NULL);
                SendMessageW(g_kw.hEdit, WM_SETFONT, (WPARAM)hf, TRUE);
                SetWindowSubclass(g_kw.hEdit, KwEditSub, 11, 0);

                if (!g_kw.initial.empty()) {
                    SetWindowTextW(g_kw.hEdit, g_kw.initial.c_str());
                    SendMessageW(g_kw.hEdit, EM_SETSEL, 0, -1);
                }

                int btnY = rc.bottom - margin - btnH;
                g_kw.hOK = CreateWindowExW(
                    0, L"BUTTON", L"OK",
                    WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                    rc.right - margin - btnW - (btnW + DpiScale(10)),
                    btnY,
                    btnW, btnH, h, (HMENU)IDOK, g_hInst, NULL);
                SendMessageW(g_kw.hOK, WM_SETFONT, (WPARAM)hf, TRUE);

                g_kw.hCancel = CreateWindowExW(
                    0, L"BUTTON", L"Cancel",
                    WS_CHILD | WS_VISIBLE,
                    rc.right - margin - btnW, btnY,
                    btnW, btnH, h, (HMENU)IDCANCEL, g_hInst, NULL);
                SendMessageW(g_kw.hCancel, WM_SETFONT, (WPARAM)hf, TRUE);

                SetFocus(g_kw.hEdit);
                return 0;
            }
            case WM_COMMAND:
                if (LOWORD(w) == IDOK) {
                    int len = GetWindowTextLengthW(g_kw.hEdit);
                    std::wstring t(len, L'\0');
                    GetWindowTextW(g_kw.hEdit, &t[0], len + 1);
                    g_kw.text = t;
                    g_kw.accepted = !g_kw.text.empty();
                    DestroyWindow(h);
                    return 0;
                }
                if (LOWORD(w) == IDCANCEL) {
                    g_kw.accepted = false;
                    DestroyWindow(h);
                    return 0;
                }
                break;
            case WM_CLOSE:
                g_kw.accepted = false;
                DestroyWindow(h);
                return 0;
            }
            return DefWindowProcW(h, m, w, l);
            };
        wc.hInstance = g_hInst;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"KwPromptClass";
        wc.style = CS_DBLCLKS;
        RegisterClassW(&wc);
        reg = true;
    }
}

static bool PromptKeyword(std::wstring& out) {
    EnsureKwClassRegistered();
    g_kw = KwCtx();
    g_kw.accepted = false;
    g_kw.label = L"Search keyword (case-insensitive):";
    g_kw.title = L"Search";
    g_kw.initial.clear();

    MONITORINFO mi; mi.cbSize = sizeof(mi);
    HMONITOR hm = MonitorFromWindow(g_hwndMain, MONITOR_DEFAULTTONEAREST);
    GetMonitorInfoW(hm, &mi);
    RECT wa = mi.rcWork;

    int W = DpiScale(600), H = DpiScale(160);
    int X = wa.left + ((wa.right - wa.left) - W) / 2;
    int Y = wa.top + ((wa.bottom - wa.top) - H) / 2;

    HWND hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
        L"KwPromptClass",
        g_kw.title.c_str(),
        WS_POPUPWINDOW | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        X, Y, W, H, g_hwndMain, NULL, g_hInst, NULL);

    SetWindowPos(hwnd, HWND_TOPMOST, X, Y, W, H, SWP_SHOWWINDOW);
    SetForegroundWindow(hwnd);

    MSG msg;
    while (IsWindow(hwnd) && GetMessageW(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    if (g_kw.accepted) {
        out = g_kw.text;
        return true;
    }
    SetForegroundWindow(g_hwndMain);
    return false;
}

static bool PromptRenameSimple(const std::wstring& currentName, std::wstring& out) {
    EnsureKwClassRegistered();
    g_kw = KwCtx();
    g_kw.accepted = false;
    g_kw.label = L"New name:";
    g_kw.title = L"Rename";
    g_kw.initial = currentName;

    MONITORINFO mi; mi.cbSize = sizeof(mi);
    HMONITOR hm = MonitorFromWindow(g_hwndMain, MONITOR_DEFAULTTONEAREST);
    GetMonitorInfoW(hm, &mi);
    RECT wa = mi.rcWork;

    int W = DpiScale(600), H = DpiScale(160);
    int X = wa.left + ((wa.right - wa.left) - W) / 2;
    int Y = wa.top + ((wa.bottom - wa.top) - H) / 2;

    HWND hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
        L"KwPromptClass",
        g_kw.title.c_str(),
        WS_POPUPWINDOW | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        X, Y, W, H, g_hwndMain, NULL, g_hInst, NULL);

    SetWindowPos(hwnd, HWND_TOPMOST, X, Y, W, H, SWP_SHOWWINDOW);
    SetForegroundWindow(hwnd);

    MSG msg;
    while (IsWindow(hwnd) && GetMessageW(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    if (g_kw.accepted) {
        out = g_kw.text;
        return true;
    }
    SetForegroundWindow(g_hwndMain);
    return false;
}

static void Browser_RenameSelected() {
    if (g_view == ViewKind::Drives) return;

    int sel = ListView_GetNextItem(g_hwndList, -1, LVNI_SELECTED);
    if (sel < 0 || sel >= (int)g_rows.size()) return;

    const Row& r = g_rows[sel];
    const wchar_t* base = wcsrchr(r.full.c_str(), L'\\');
    const wchar_t* name = base ? base + 1 : r.full.c_str();

    std::wstring newName;
    if (!PromptRenameSimple(name, newName)) return;
    newName = Trim(newName);
    if (newName.empty() || _wcsicmp(name, newName.c_str()) == 0) return;

    std::wstring newPath;
    if (base) {
        newPath.assign(r.full.c_str(), base + 1 - r.full.c_str());
        newPath += newName;
    }
    else {
        newPath = newName;
    }

    BOOL ok = MoveFileExW(
        r.full.c_str(), newPath.c_str(),
        MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING);
    if (!ok) {
        DWORD err = GetLastError();
        wchar_t buf[256];
        swprintf_s(buf, L"Rename failed (error %lu).", err);
        MessageBoxW(g_hwndMain, buf, L"Rename", MB_OK | MB_ICONERROR);
        return;
    }

    if (g_view == ViewKind::Folder) {
        ShowFolder(g_folder);
    }
    else if (g_view == ViewKind::Search && g_search.active) {
        std::vector<Row> res;
        RunSearchFromOrigin(res);
        ShowSearchResults(res);
    }
    else if (g_view == ViewKind::Drives) {
        ShowDrives();
    }
}

// ----------------------------- Playback / post actions

static void ApplyPostActionsAndRefresh() {
    for (size_t i = 0; i < g_post.size(); ++i) {
        const PostAction& a = g_post[i];
        switch (a.type) {
        case ActionType::DeleteFile: {
            BOOL ok = DeleteFileW(a.src.c_str());
            if (!ok) {
                DWORD err = GetLastError();
                MoveFileExW(a.src.c_str(), NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
                LogLine(L"PostAction DeleteFile: src=\"%s\" FAILED err=%lu (queued delete)",
                    a.src.c_str(), err);
            }
            else {
                LogLine(L"PostAction DeleteFile: src=\"%s\" OK", a.src.c_str());
            }
            break;
        }
        case ActionType::RenameFile: {
            BOOL ok = MoveFileExW(
                a.src.c_str(), a.param.c_str(),
                MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING);
            DWORD err = ok ? 0 : GetLastError();
            LogLine(L"PostAction RenameFile: src=\"%s\" dst=\"%s\" %s err=%lu",
                a.src.c_str(), a.param.c_str(), ok ? L"OK" : L"FAILED", err);
            break;
        }
        case ActionType::CopyToPath: {
            std::wstring t = L"Browse - copying file...";
            SetWindowTextW(g_hwndMain, t.c_str());
            BOOL ok = CopyFileW(a.src.c_str(), a.param.c_str(), FALSE);
            DWORD err = ok ? 0 : GetLastError();
            LogLine(L"PostAction CopyToPath: src=\"%s\" dst=\"%s\" %s err=%lu",
                a.src.c_str(), a.param.c_str(), ok ? L"OK" : L"FAILED", err);
            break;
        }
        }
    }
    g_post.clear();

    if (g_view == ViewKind::Search && g_search.active) {
        std::vector<Row> res;
        RunSearchFromOrigin(res);
        ShowSearchResults(res);
    }
    else if (g_view == ViewKind::Drives) {
        ShowDrives();
    }
    else {
        ShowFolder(g_folder);
    }
}

static void PlayIndex(size_t idx) {
    if (!g_vlc) {
        const char* args[] = { "--avcodec-hw=d3d11va", "--no-video-title-show" };
        g_vlc = libvlc_new((int)(sizeof(args) / sizeof(args[0])), args);
        g_mp = libvlc_media_player_new(g_vlc);
        libvlc_media_player_set_hwnd(g_mp, g_hwndVideo);
        libvlc_video_set_scale(g_mp, 0.f);
        libvlc_video_set_aspect_ratio(g_mp, NULL);

        libvlc_event_manager_t* em = libvlc_media_player_event_manager(g_mp);
        libvlc_event_attach(em, libvlc_MediaPlayerEndReached,
            [](const libvlc_event_t*, void*) {
                PostMessageW(g_hwndMain, WM_APP + 1, 0, 0);
            }, NULL);
    }

    g_playlistIndex = idx;
    g_lastLenForRange = -1;
    SendMessageW(g_hwndSeek, TBM_SETRANGEMAX, TRUE, 0);
    SendMessageW(g_hwndSeek, TBM_SETPOS, TRUE, 0);

    std::string u8 = ToUtf8(g_playlist[g_playlistIndex]);
    libvlc_media_t* m = libvlc_media_new_path(g_vlc, u8.c_str());
    libvlc_media_player_set_media(g_mp, m);
    libvlc_media_release(m);
    libvlc_media_player_play(g_mp);
}

static void ToggleFullscreen() {
    if (!g_inPlayback) return;
    HWND h = g_hwndMain;
    if (!g_fullscreen) {
        ZeroMemory(&g_wpPrev, sizeof(g_wpPrev));
        g_wpPrev.length = sizeof(g_wpPrev);
        GetWindowPlacement(h, &g_wpPrev);
        LONG style = GetWindowLongW(h, GWL_STYLE);
        SetWindowLongW(h, GWL_STYLE, style & ~(WS_OVERLAPPEDWINDOW));
        MONITORINFO mi; mi.cbSize = sizeof(mi);
        if (GetMonitorInfoW(MonitorFromWindow(h, MONITOR_DEFAULTTOPRIMARY), &mi)) {
            SetWindowPos(h, HWND_TOP,
                mi.rcMonitor.left, mi.rcMonitor.top,
                mi.rcMonitor.right - mi.rcMonitor.left,
                mi.rcMonitor.bottom - mi.rcMonitor.top,
                SWP_NOOWNERZORDER | SWP_FRAMECHANGED);
        }
        g_fullscreen = true;
    }
    else {
        LONG style = GetWindowLongW(h, GWL_STYLE);
        SetWindowLongW(h, GWL_STYLE, style | WS_OVERLAPPEDWINDOW);
        SetWindowPlacement(h, &g_wpPrev);
        SetWindowPos(h, NULL, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER |
            SWP_NOOWNERZORDER | SWP_FRAMECHANGED);
        g_fullscreen = false;
    }
}

static void ExitPlayback() {
    LogLine(L"ExitPlayback called: inPlayback=%d", g_inPlayback ? 1 : 0);
    if (!g_inPlayback) return;

    if (g_fullscreen) ToggleFullscreen();
    KillTimer(g_hwndMain, kTimerPlaybackUI);
    if (g_mp) libvlc_media_player_stop(g_mp);

    ShowWindow(g_hwndVideo, SW_HIDE);
    ShowWindow(g_hwndSeek, SW_HIDE);
    ShowWindow(g_hwndList, SW_SHOW);
    SetFocus(g_hwndList);
    g_inPlayback = false;

    RECT rc; GetClientRect(g_hwndMain, &rc);
    MoveWindow(g_hwndList, 0, 0, rc.right, rc.bottom, TRUE);

    ApplyPostActionsAndRefresh();
    SetTitleFolderOrDrives();
    LogLine(L"ExitPlayback finished");
}

static void NextInPlaylist() {
    if (!g_inPlayback) return;
    if (g_playlistIndex + 1 < g_playlist.size()) PlayIndex(g_playlistIndex + 1);
}

static void PrevInPlaylist() {
    if (!g_inPlayback) return;
    if (g_playlistIndex > 0) PlayIndex(g_playlistIndex - 1);
}

static void PlaySelectedVideos() {
    g_playlist.clear();
    int idx = -1;
    while ((idx = ListView_GetNextItem(g_hwndList, idx, LVNI_SELECTED)) != -1) {
        if (idx >= 0 && idx < (int)g_rows.size()) {
            const Row& it = g_rows[idx];
            if (!it.isDir && IsVideoFile(it.full)) g_playlist.push_back(it.full);
        }
    }
    if (g_playlist.empty()) return;

    g_inPlayback = true;
    ShowWindow(g_hwndList, SW_HIDE);
    ShowWindow(g_hwndSeek, SW_SHOW);
    ShowWindow(g_hwndVideo, SW_SHOW);
    SetFocus(g_hwndVideo);

    RECT rc; GetClientRect(g_hwndMain, &rc);
    const int seekH = 32;
    MoveWindow(g_hwndVideo, 0, 0, rc.right, rc.bottom - seekH, TRUE);
    MoveWindow(g_hwndSeek, 0, rc.bottom - seekH, rc.right, seekH, TRUE);

    SendMessageW(g_hwndSeek, TBM_SETRANGEMIN, TRUE, 0);
    SendMessageW(g_hwndSeek, TBM_SETRANGEMAX, TRUE, 0);
    SendMessageW(g_hwndSeek, TBM_SETPOS, TRUE, 0);

    PlayIndex(0);
    SetTimer(g_hwndMain, kTimerPlaybackUI, 200, NULL);
    SetTitlePlaying();
}

static void ActivateSelection() {
    int i = ListView_GetNextItem(g_hwndList, -1, LVNI_SELECTED);
    if (i < 0 || i >= (int)g_rows.size()) return;
    const Row& r = g_rows[i];

    // NEW: drives view broken-drive block
    if (g_view == ViewKind::Drives && r.isBrokenNetDrive) {
        MessageBeep(MB_ICONWARNING);
        // Optional: MessageBoxW(g_hwndMain, L"This mapped drive is disconnected. Right-click and choose Fix.", L"Browse", MB_OK | MB_ICONWARNING);
        return;
    }

    if (g_view == ViewKind::Drives || r.isDir) {
        if (g_view == ViewKind::Search) return;
        ShowFolder(r.full);
    }
    else {
        if (IsVideoFile(r.full)) {
            PlaySelectedVideos();
        }
        else {
            ShellExecuteW(g_hwndMain, L"open", r.full.c_str(), NULL, NULL, SW_SHOWNORMAL);
        }
    }
}

static void NavigateBack() {
    if (g_view == ViewKind::Search) { ExitSearchToOrigin(); return; }
    if (g_view == ViewKind::Drives) return;
    if (IsDriveRoot(g_folder)) { ShowDrives(); return; }
    std::wstring parent = ParentDir(g_folder);
    if (parent.empty()) ShowDrives();
    else ShowFolder(parent);
}

// ----------------------------- Playlist chooser

struct PickerCtx { HWND hwnd, hList; };
static PickerCtx g_pick = { 0,0 };

static LRESULT CALLBACK PickerProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    switch (m) {
    case WM_CREATE: {
        g_pick.hList = CreateWindowExW(
            WS_EX_CLIENTEDGE, L"LISTBOX", L"",
            WS_CHILD | WS_VISIBLE | LBS_NOTIFY | WS_VSCROLL | LBS_NOINTEGRALHEIGHT,
            0, 0, 100, 100, h, (HMENU)2001, g_hInst, NULL);
        SendMessageW(g_pick.hList, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);
        for (size_t i = 0; i < g_playlist.size(); ++i) {
            const std::wstring& p = g_playlist[i];
            const wchar_t* base = wcsrchr(p.c_str(), L'\\'); base = base ? base + 1 : p.c_str();
            SendMessageW(g_pick.hList, LB_ADDSTRING, 0, (LPARAM)base);
        }
        SendMessageW(g_pick.hList, LB_SETCURSEL, (WPARAM)g_playlistIndex, 0);
        return 0;
    }
    case WM_SIZE:
        MoveWindow(g_pick.hList, 8, 8, LOWORD(l) - 16, HIWORD(l) - 16, TRUE);
        return 0;
    case WM_COMMAND:
        if (HIWORD(w) == LBN_SELCHANGE && (HWND)l == g_pick.hList) {
            int sel = (int)SendMessageW(g_pick.hList, LB_GETCURSEL, 0, 0);
            if (sel >= 0 && sel < (int)g_playlist.size()) PlayIndex((size_t)sel);
            return 0;
        }
        if (HIWORD(w) == LBN_DBLCLK && (HWND)l == g_pick.hList) {
            DestroyWindow(h);
            return 0;
        }
        break;
    case WM_KEYDOWN:
        if (w == VK_RETURN || w == VK_ESCAPE) { DestroyWindow(h); return 0; }
        break;
    case WM_CLOSE:
        DestroyWindow(h); return 0;
    case WM_DESTROY:
        if (g_mp) libvlc_media_player_set_pause(g_mp, 0);
        return 0;
    }
    return DefWindowProcW(h, m, w, l);
}

static void ShowPlaylistChooser() {
    if (!g_inPlayback || g_playlist.empty()) return;
    if (g_mp) libvlc_media_player_set_pause(g_mp, 1);
    static bool registered = false;
    if (!registered) {
        WNDCLASSW wc; ZeroMemory(&wc, sizeof(wc));
        wc.lpfnWndProc = PickerProc;
        wc.hInstance = g_hInst;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"PlaylistPickerClass";
        RegisterClassW(&wc);
        registered = true;
    }
    RECT r; SystemParametersInfoW(SPI_GETWORKAREA, 0, &r, 0);
    int W = 520, H = 420;
    int X = r.left + ((r.right - r.left) - W) / 2;
    int Y = r.top + ((r.bottom - r.top) - H) / 2;
    HWND hwnd = CreateWindowExW(
        WS_EX_TOOLWINDOW, L"PlaylistPickerClass", L"Playlist",
        WS_POPUPWINDOW | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        X, Y, W, H, g_hwndMain, NULL, g_hInst, NULL);
    g_pick.hwnd = hwnd;

    MSG msg;
    while (IsWindow(hwnd) && GetMessageW(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}

// ----------------------------- List subclass (keyboard on list view)

static LRESULT CALLBACK ListSubclass(HWND h, UINT m, WPARAM w, LPARAM l,
    UINT_PTR, DWORD_PTR) {
    if (m == WM_GETDLGCODE) return DLGC_WANTALLKEYS;
    if (m == WM_KEYDOWN) {
        if (g_loadingFolder) return 0;

        bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;

        switch (w) {
        case VK_LEFT:
        case VK_BACK:
            NavigateBack();
            return 0;
        case VK_F1:
            ShowHelp();
            return 0;
        case VK_F2:
            Browser_RenameSelected();
            return 0;

        case 'A':
            if (ctrl) {
                for (int i = 0; i < (int)g_rows.size(); ++i) {
                    ListView_SetItemState(g_hwndList, i,
                        LVIS_SELECTED, LVIS_SELECTED);
                }
                return 0;
            }
            break;

        case 'P':
            if (ctrl) {
                PlaySelectedVideos();
                return 0;
            }
            break;

        case 'F':
            if (ctrl) {
                std::wstring kw;
                if (!PromptKeyword(kw)) return 0;
                kw = ToLower(kw);
                if (kw.empty()) return 0;

                if (g_view != ViewKind::Search) {
                    g_search.active = true;
                    g_search.originView = g_view;
                    g_search.originFolder =
                        (g_view == ViewKind::Folder ? g_folder : L"");
                    g_search.termsLower.clear();
                    g_search.termsLower.push_back(kw);

                    g_search.useExplicitScope = false;
                    g_search.explicitFolders.clear();
                    g_search.explicitFiles.clear();

                    std::vector<std::wstring> selFolders, selFiles;
                    CollectSelection(selFolders, selFiles);
                    if (!selFolders.empty() || !selFiles.empty()) {
                        g_search.useExplicitScope = true;
                        g_search.explicitFolders.swap(selFolders);
                        g_search.explicitFiles.swap(selFiles);
                    }

                    std::vector<Row> res;
                    RunSearchFromOrigin(res);
                    ShowSearchResults(res);
                }
                else {
                    g_search.termsLower.push_back(kw);
                    std::vector<Row> filtered;
                    filtered.reserve(g_rows.size());
                    for (size_t i = 0; i < g_rows.size(); ++i) {
                        if (NameContainsAllTerms(g_rows[i].full,
                            g_search.termsLower))
                            filtered.push_back(g_rows[i]);
                    }
                    ShowSearchResults(filtered);
                }
                return 0;
            }
            break;

        case 'C':
            if (ctrl) {
                Browser_CopySelectedToClipboard(ClipMode::Copy);
                return 0;
            }
            break;
        case 'X':
            if (ctrl) {
                Browser_CopySelectedToClipboard(ClipMode::Move);
                return 0;
            }
            break;
        case 'V':
            if (ctrl) {
                Browser_PasteClipboardIntoCurrent();
                return 0;
            }
            break;
        case VK_DELETE:
            Browser_DeleteSelected();
            return 0;
        }
    }
    return DefSubclassProc(h, m, w, l);
}

// ----------------------------- Video subclass (keyboard in playback)

static LRESULT CALLBACK VideoSubclass(HWND h, UINT m, WPARAM w, LPARAM l,
    UINT_PTR, DWORD_PTR) {
    if (m == WM_GETDLGCODE) return DLGC_WANTALLKEYS;
    if (m == WM_KEYDOWN && g_mp) {
        bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

        switch (w) {
        case VK_F1:     ShowHelp(); return 0;
        case VK_RETURN: ToggleFullscreen(); return 0;
        case VK_SPACE:  libvlc_media_player_set_pause(g_mp, 1); return 0;
        case VK_TAB:    libvlc_media_player_set_pause(g_mp, 0); return 0;
        case VK_ESCAPE: ExitPlayback(); return 0;

        case 'G':
            if (ctrl) { ShowPlaylistChooser(); return 0; }
            break;

        case 'P':
            if (ctrl) { ShowCurrentVideoProperties(); return 0; }
            break;

        case VK_DELETE: {
            if (!g_playlist.empty()) {
                std::wstring doomed = g_playlist[g_playlistIndex];
                std::vector<std::wstring> np;
                np.reserve(g_playlist.size());
                for (size_t i = 0; i < g_playlist.size(); ++i)
                    if (i != g_playlistIndex) np.push_back(g_playlist[i]);
                g_playlist.swap(np);
                g_post.push_back({ ActionType::DeleteFile, doomed, L"" });
                if (g_playlist.empty()) ExitPlayback();
                else if (g_playlistIndex >= g_playlist.size())
                    PlayIndex(g_playlist.size() - 1);
                else PlayIndex(g_playlistIndex);
            }
            return 0;
        }


        case 'R':
            if (ctrl && !g_playlist.empty()) {
                std::wstring cur = g_playlist[g_playlistIndex];
                std::wstring newPath;
                libvlc_media_player_set_pause(g_mp, 1);
                if (PromptSaveAsFrom(cur, newPath, L"Rename file")) {
                    if (_wcsicmp(cur.c_str(), newPath.c_str()) != 0)
                        g_post.push_back({ ActionType::RenameFile, cur, newPath });
                }
                libvlc_media_player_set_pause(g_mp, 0);
                return 0;
            }
            break;

        case 'C':
            if (ctrl && !g_playlist.empty()) {
                std::wstring cur = g_playlist[g_playlistIndex];
                std::wstring destFull;
                libvlc_media_player_set_pause(g_mp, 1);
                if (PromptSaveAsFrom(cur, destFull, L"Copy file to")) {
                    if (_wcsicmp(cur.c_str(), destFull.c_str()) != 0)
                        g_post.push_back({ ActionType::CopyToPath, cur, destFull });
                }
                libvlc_media_player_set_pause(g_mp, 0);
                return 0;
            }
            break;

        case VK_UP: {
            int v = libvlc_audio_get_volume(g_mp);
            v = (v < 0 ? 0 : v) + 5; if (v > 200) v = 200;
            libvlc_audio_set_volume(g_mp, v); return 0;
        }
        case VK_DOWN: {
            int v = libvlc_audio_get_volume(g_mp);
            v = (v < 0 ? 0 : v) - 5; if (v < 0) v = 0;
            libvlc_audio_set_volume(g_mp, v); return 0;
        }
        case VK_LEFT:
        case VK_RIGHT: {
            if (ctrl) {
                if (w == VK_RIGHT) NextInPlaylist();
                else PrevInPlaylist();
            }
            else {
                libvlc_time_t cur = libvlc_media_player_get_time(g_mp);
                libvlc_time_t len = libvlc_media_player_get_length(g_mp);
                libvlc_time_t step = shift ? 60000 : 10000;
                if (w == VK_RIGHT) cur += step;
                else cur = (cur > step ? cur - step : 0);
                if (len > 0 && cur > len) cur = len;
                libvlc_media_player_set_time(g_mp, cur);
            }
            return 0;
        }
        }
    }
    return DefSubclassProc(h, m, w, l);
}

// ----------------------------- Seek subclass

static LRESULT CALLBACK SeekSubclass(HWND h, UINT m, WPARAM w, LPARAM l,
    UINT_PTR, DWORD_PTR) {
    if (m == WM_KEYDOWN) {
        bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;

        if (w == VK_F1) { ShowHelp(); return 0; }
        if (w == VK_ESCAPE) { ExitPlayback(); return 0; }
        if (w == VK_RETURN) { ToggleFullscreen(); return 0; }
        if (w == VK_LEFT || w == VK_RIGHT || w == VK_UP || w == VK_DOWN ||
            w == VK_SPACE || w == VK_TAB || w == VK_DELETE) {
            SendMessageW(g_hwndVideo, WM_KEYDOWN, w, l);
            return 0;
        }

        if (ctrl && (w == 'R' || w == 'r')) {
            SendMessageW(g_hwndVideo, WM_KEYDOWN, 'R', 0);
            return 0;
        }
        if (ctrl && (w == 'C' || w == 'c')) {
            SendMessageW(g_hwndVideo, WM_KEYDOWN, 'C', 0);
            return 0;
        }
        if (ctrl && (w == 'G' || w == 'g')) {
            ShowPlaylistChooser();
            return 0;
        }
        if (ctrl && (w == 'P' || w == 'p')) {
            SendMessageW(g_hwndVideo, WM_KEYDOWN, 'P', 0);
            return 0;
        }
    }
    return DefSubclassProc(h, m, w, l);
}

// ----------------------------- Layout

static void OnSize(int cx, int cy) {
    if (g_inPlayback) {
        const int seekH = 32;
        if (g_hwndVideo) MoveWindow(g_hwndVideo, 0, 0, cx, cy - seekH, TRUE);
        if (g_hwndSeek)  MoveWindow(g_hwndSeek, 0, cy - seekH, cx, seekH, TRUE);
    }
    else {
        if (g_hwndList)  MoveWindow(g_hwndList, 0, 0, cx, cy, TRUE);
    }
}

// ----------------------------- Icon loader

static HICON LoadAppIcon(int cx, int cy) {
    wchar_t exePath[MAX_PATH] = {};
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    PathRemoveFileSpecW(exePath);
    std::wstring p = std::wstring(exePath) + L"\\Browse.ico";

    HICON h = (HICON)LoadImageW(NULL, p.c_str(), IMAGE_ICON, cx, cy, LR_LOADFROMFILE);
    if (!h) {
        h = (HICON)LoadImageW(NULL, L"Browse.ico", IMAGE_ICON, cx, cy, LR_LOADFROMFILE);
    }
    return h;
}

// ----------------------------- Network drive helpers

// ----------------------------- Network drive helpers (enhanced)

static std::wstring WinErrText(DWORD err) {
    wchar_t* msg = nullptr;
    DWORD n = FormatMessageW(
        FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        NULL, err, 0, (LPWSTR)&msg, 0, NULL);

    std::wstring s;
    if (n && msg) {
        s.assign(msg, msg + n);
        LocalFree(msg);
        s = Trim(s);
    }
    return s;
}

static bool RegReadStringValue(HKEY hKey, const wchar_t* valueName, std::wstring& out) {
    out.clear();
    DWORD type = 0, cb = 0;
    LONG rc = RegQueryValueExW(hKey, valueName, NULL, &type, NULL, &cb);
    if (rc != ERROR_SUCCESS) return false;
    if (type != REG_SZ && type != REG_EXPAND_SZ) return false;
    if (cb < sizeof(wchar_t)) return false;

    std::vector<wchar_t> buf(cb / sizeof(wchar_t) + 2, L'\0');
    rc = RegQueryValueExW(hKey, valueName, NULL, &type, (LPBYTE)buf.data(), &cb);
    if (rc != ERROR_SUCCESS) return false;

    buf.back() = L'\0';
    out = buf.data();
    out = Trim(out);
    return !out.empty();
}

// Reads persistent mapping (if present): HKCU\Network\<Letter>\RemotePath
static bool GetPersistentMappedRemotePath(wchar_t letter, std::wstring& outRemote) {
    outRemote.clear();
    letter = (wchar_t)towupper(letter);

    wchar_t subkey[64];
    swprintf_s(subkey, L"Network\\%c", letter);

    HKEY h = NULL;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, subkey, 0, KEY_READ, &h) != ERROR_SUCCESS) {
        return false;
    }

    std::wstring remote;
    bool ok = RegReadStringValue(h, L"RemotePath", remote);
    RegCloseKey(h);

    if (!ok) return false;
    outRemote = remote;
    return true;
}

// Enumerate currently-connected network drive letters (like "net use" Status=OK)
static DWORD GetConnectedNetDriveMask() {
    DWORD mask = 0;

    HANDLE hEnum = NULL;
    DWORD res = WNetOpenEnumW(RESOURCE_CONNECTED, RESOURCETYPE_DISK, 0, NULL, &hEnum);
    if (res != NO_ERROR) return 0;

    std::vector<BYTE> buf(16 * 1024);
    for (;;) {
        DWORD count = 0xFFFFFFFF;
        DWORD size = (DWORD)buf.size();
        res = WNetEnumResourceW(hEnum, &count, buf.data(), &size);
        if (res == ERROR_NO_MORE_ITEMS) break;
        if (res != NO_ERROR) break;

        NETRESOURCEW* nr = (NETRESOURCEW*)buf.data();
        for (DWORD i = 0; i < count; ++i) {
            if (nr[i].lpLocalName && nr[i].lpLocalName[0]) {
                wchar_t c = (wchar_t)towupper(nr[i].lpLocalName[0]);
                if (c >= L'A' && c <= L'Z' && nr[i].lpLocalName[1] == L':') {
                    mask |= (1u << (c - L'A'));
                }
            }
        }
    }

    WNetCloseEnum(hEnum);
    return mask;
}

static DWORD DisconnectDriveLetter(const std::wstring& localName /*"X:"*/) {
    // TRUE = force if open files, CONNECT_UPDATE_PROFILE removes persistent mapping
    return WNetCancelConnection2W(localName.c_str(), CONNECT_UPDATE_PROFILE, TRUE);
}

static DWORD ConnectDriveLetter(const std::wstring& localName /*"X:"*/,
    const std::wstring& remote /*"\\\\server\\share"*/) {

    NETRESOURCEW nr{};
    nr.dwType = RESOURCETYPE_DISK;
    nr.lpLocalName = (LPWSTR)localName.c_str();
    nr.lpRemoteName = (LPWSTR)remote.c_str();

    const wchar_t* user = g_cfg.netUsername.empty() ? NULL : g_cfg.netUsername.c_str();
    const wchar_t* pass = g_cfg.netPassword.empty() ? NULL : g_cfg.netPassword.c_str();

    DWORD flags = CONNECT_UPDATE_PROFILE;

    // If creds provided in INI -> non-interactive connect.
    if (user || pass) {
        return WNetAddConnection2W(&nr, pass, user, flags);
    }

    // Otherwise: allow Windows to prompt for creds if needed.
    wchar_t accessName[256] = {};
    DWORD accessSize = _countof(accessName);
    DWORD resultFlags = 0;
    return WNetUseConnectionW(
        g_hwndMain, &nr,
        NULL, NULL,
        CONNECT_INTERACTIVE | CONNECT_PROMPT | CONNECT_UPDATE_PROFILE,
        accessName, &accessSize, &resultFlags);
}

static bool GetSelectedDriveLetter(wchar_t& outLetter) {
    outLetter = 0;
    if (g_view != ViewKind::Drives) return false;

    int sel = ListView_GetNextItem(g_hwndList, -1, LVNI_SELECTED);
    if (sel < 0 || sel >= (int)g_rows.size()) return false;

    const Row& r = g_rows[sel];
    if (!r.full.empty()) outLetter = (wchar_t)towupper(r.full[0]);
    else if (!r.name.empty()) outLetter = (wchar_t)towupper(r.name[0]);

    return (outLetter >= L'A' && outLetter <= L'Z');
}

static void FixSelectedBrokenDrive() {
    wchar_t letter = 0;
    if (!GetSelectedDriveLetter(letter)) return;

    std::wstring remote;
    if (!GetPersistentMappedRemotePath(letter, remote)) {
        wchar_t buf[512];
        swprintf_s(buf,
            L"Cannot determine the original UNC share for %c:.\n"
            L"(No persistent mapping found in HKCU\\Network\\%c)\n\n"
            L"Map it again, then retry Fix.",
            letter, letter);
        MessageBoxW(g_hwndMain, buf, L"Fix Drive", MB_OK | MB_ICONERROR);
        return;
    }

    std::wstring localName;
    localName.push_back(letter);
    localName.push_back(L':');

    LogLine(L"FixDrive: %s -> %s (user set=%d pass set=%d)",
        localName.c_str(), remote.c_str(),
        g_cfg.netUsername.empty() ? 0 : 1,
        g_cfg.netPassword.empty() ? 0 : 1);

    // Disconnect (ignore common "already disconnected" errors)
    DWORD d = DisconnectDriveLetter(localName);
    LogLine(L"FixDrive: disconnect rc=%lu", d);

    // Reconnect
    DWORD c = ConnectDriveLetter(localName, remote);
    if (c != NO_ERROR) {
        std::wstring msg = L"Reconnect failed for ";
        msg += localName;
        msg += L" -> ";
        msg += remote;
        msg += L"\n\nError ";
        msg += std::to_wstring(c);
        std::wstring et = WinErrText(c);
        if (!et.empty()) { msg += L": "; msg += et; }

        MessageBoxW(g_hwndMain, msg.c_str(), L"Fix Drive", MB_OK | MB_ICONERROR);
        LogLine(L"FixDrive: reconnect FAILED rc=%lu", c);
        return;
    }

    LogLine(L"FixDrive: reconnect OK");
    ShowDrives(); // refresh + re-evaluate red/ok
}

// simple "map" prompt using existing single-line dialog style
static bool PromptSingleLine(const wchar_t* title, const wchar_t* label, const std::wstring& initial, std::wstring& out) {
    EnsureKwClassRegistered();
    g_kw = KwCtx();
    g_kw.accepted = false;
    g_kw.title = title ? title : L"Input";
    g_kw.label = label ? label : L"Input:";
    g_kw.initial = initial;

    MONITORINFO mi; mi.cbSize = sizeof(mi);
    HMONITOR hm = MonitorFromWindow(g_hwndMain, MONITOR_DEFAULTTONEAREST);
    GetMonitorInfoW(hm, &mi);
    RECT wa = mi.rcWork;

    int W = DpiScale(700), H = DpiScale(170);
    int X = wa.left + ((wa.right - wa.left) - W) / 2;
    int Y = wa.top + ((wa.bottom - wa.top) - H) / 2;

    HWND hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
        L"KwPromptClass",
        g_kw.title.c_str(),
        WS_POPUPWINDOW | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        X, Y, W, H, g_hwndMain, NULL, g_hInst, NULL);

    SetWindowPos(hwnd, HWND_TOPMOST, X, Y, W, H, SWP_SHOWWINDOW);
    SetForegroundWindow(hwnd);

    MSG msg;
    while (IsWindow(hwnd) && GetMessageW(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    if (g_kw.accepted) {
        out = g_kw.text;
        return true;
    }
    SetForegroundWindow(g_hwndMain);
    return false;
}

static wchar_t PickFreeDriveLetter() {
    DWORD mask = GetLogicalDrives();
    for (wchar_t c = L'Z'; c >= L'D'; --c) {
        int bit = (int)(c - L'A');
        if ((mask & (1u << bit)) == 0) return c;
    }
    return 0;
}

static bool ParseMapInput(const std::wstring& in, std::wstring& outLocal, std::wstring& outRemote) {
    outLocal.clear();
    outRemote.clear();

    std::wstring s = Trim(in);
    if (s.empty()) return false;

    // split "token rest"
    size_t sp = s.find_first_of(L" \t");
    std::wstring a = (sp == std::wstring::npos) ? s : Trim(s.substr(0, sp));
    std::wstring b = (sp == std::wstring::npos) ? L"" : Trim(s.substr(sp));

    // If first token looks like "X:" then it is local drive
    if (a.size() >= 2 && iswalpha(a[0]) && a[1] == L':') {
        wchar_t letter = (wchar_t)towupper(a[0]);
        outLocal.push_back(letter);
        outLocal.push_back(L':');
        outRemote = b;
    }
    else {
        outRemote = s;
    }

    if (outRemote.size() < 3 || !(outRemote[0] == L'\\' && outRemote[1] == L'\\')) {
        return false;
    }

    if (outLocal.empty()) {
        wchar_t freeL = PickFreeDriveLetter();
        if (!freeL) return false;
        outLocal.push_back(freeL);
        outLocal.push_back(L':');
    }
    return true;
}

static void MapNetworkDriveWithDefaults() {
    std::wstring input;
    if (!PromptSingleLine(
        L"Map Network Drive",
        L"Enter:  X: \\\\server\\share   (or just:  \\\\server\\share)",
        L"",
        input)) {
        return;
    }

    std::wstring localName, remote;
    if (!ParseMapInput(input, localName, remote)) {
        MessageBoxW(g_hwndMain,
            L"Invalid format.\n\nUse:\n  X: \\\\server\\share\nor:\n  \\\\server\\share",
            L"Map Network Drive", MB_OK | MB_ICONWARNING);
        return;
    }

    LogLine(L"MapDrive: %s -> %s (user set=%d pass set=%d)",
        localName.c_str(), remote.c_str(),
        g_cfg.netUsername.empty() ? 0 : 1,
        g_cfg.netPassword.empty() ? 0 : 1);

    DWORD rc = ConnectDriveLetter(localName, remote);
    if (rc != NO_ERROR) {
        std::wstring msg = L"Map failed for ";
        msg += localName;
        msg += L" -> ";
        msg += remote;
        msg += L"\n\nError ";
        msg += std::to_wstring(rc);
        std::wstring et = WinErrText(rc);
        if (!et.empty()) { msg += L": "; msg += et; }
        MessageBoxW(g_hwndMain, msg.c_str(), L"Map Network Drive", MB_OK | MB_ICONERROR);
        return;
    }

    if (g_view == ViewKind::Drives) {
        ShowDrives();
    }
}


static void ShowMapNetworkDriveDialog() {
    // now uses INI username/password if set (and prompts otherwise)
    MapNetworkDriveWithDefaults();
}

static void ShowDisconnectNetworkDriveDialog() {
    DWORD res = WNetDisconnectDialog(g_hwndMain, RESOURCETYPE_DISK);
    if (res != NO_ERROR && res != ERROR_CANCELLED) {
        wchar_t buf[256];
        swprintf_s(buf, L"Failed to disconnect network drive (error %lu).", res);
        MessageBoxW(g_hwndMain, buf, L"Disconnect Network Drive", MB_OK | MB_ICONERROR);
    }
}

// ----------------------------- List context menu

static void ShowListContextMenu(POINT ptScreen) {
    if (!g_hwndList) return;

    POINT ptClient = ptScreen;
    ScreenToClient(g_hwndList, &ptClient);

    LVHITTESTINFO hti{};
    hti.pt = ptClient;
    int idx = ListView_HitTest(g_hwndList, &hti);
    bool onItem = (idx >= 0) && (hti.flags & LVHT_ONITEM);

    // Special-case: broken mapped drive in Drives view => only "Fix"
    if (onItem && g_view == ViewKind::Drives) {
        const Row& r = g_rows[idx];
        if (r.isBrokenNetDrive) {
            HMENU hMenu = CreatePopupMenu();
            if (!hMenu) return;

            AppendMenuW(hMenu, MF_STRING, ID_CTX_FIXDRIVE, L"&Fix");

            int cmd = TrackPopupMenu(
                hMenu,
                TPM_RIGHTBUTTON | TPM_RETURNCMD,
                ptScreen.x, ptScreen.y,
                0, g_hwndMain, NULL);

            DestroyMenu(hMenu);

            if (cmd == ID_CTX_FIXDRIVE) {
                FixSelectedBrokenDrive();
            }
            return;
        }
    }

    if (onItem) {
        if (!(ListView_GetItemState(g_hwndList, idx, LVIS_SELECTED) & LVIS_SELECTED)) {
            ListView_SetItemState(g_hwndList, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
            ListView_SetItemState(g_hwndList, idx,
                LVIS_SELECTED | LVIS_FOCUSED,
                LVIS_SELECTED | LVIS_FOCUSED);
        }
    }

    HMENU hMenu = CreatePopupMenu();
    if (!hMenu) return;

    UINT pasteFlags = MF_GRAYED;
    if (!(g_clipMode == ClipMode::None || g_clipFiles.empty())) {
        pasteFlags = 0;
    }
    else if (IsClipboardFormatAvailable(CF_HDROP)) {
        pasteFlags = 0;
    }

    if (onItem) {
        const Row& r = g_rows[idx];
        AppendMenuW(hMenu, MF_STRING, ID_CTX_OPEN, L"&Open");
        if (!r.isDir && IsVideoFile(r.full)) {
            AppendMenuW(hMenu, MF_STRING, ID_CTX_PLAY, L"&Play video");
        }
        AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenuW(hMenu, MF_STRING, ID_CTX_RENAME, L"Rena&me");
        AppendMenuW(hMenu, MF_STRING, ID_CTX_CUT, L"Cu&t");
        AppendMenuW(hMenu, MF_STRING, ID_CTX_COPY, L"&Copy");
        AppendMenuW(hMenu, MF_STRING | pasteFlags, ID_CTX_PASTE, L"&Paste");
        AppendMenuW(hMenu, MF_STRING, ID_CTX_DELETE, L"&Delete");
        AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
    }
    else {
        AppendMenuW(hMenu, MF_STRING | pasteFlags, ID_CTX_PASTE, L"&Paste");
        AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
    }

    AppendMenuW(hMenu, MF_STRING, ID_CTX_MAPDRIVE, L"Map Network &Drive...");
    AppendMenuW(hMenu, MF_STRING, ID_CTX_DISCONNECT, L"&Disconnect Network Drive...");

    int cmd = TrackPopupMenu(
        hMenu,
        TPM_RIGHTBUTTON | TPM_RETURNCMD,
        ptScreen.x, ptScreen.y,
        0, g_hwndMain, NULL);
    DestroyMenu(hMenu);

    switch (cmd) {
    case ID_CTX_OPEN:
        ActivateSelection();
        break;
    case ID_CTX_PLAY:
        PlaySelectedVideos();
        break;
    case ID_CTX_RENAME:
        Browser_RenameSelected();
        break;
    case ID_CTX_CUT:
        Browser_CopySelectedToClipboard(ClipMode::Move);
        break;
    case ID_CTX_COPY:
        Browser_CopySelectedToClipboard(ClipMode::Copy);
        break;
    case ID_CTX_PASTE:
        Browser_PasteClipboardIntoCurrent();
        break;
    case ID_CTX_DELETE:
        Browser_DeleteSelected();
        break;
    case ID_CTX_MAPDRIVE:
        ShowMapNetworkDriveDialog();
        break;
    case ID_CTX_DISCONNECT:
        ShowDisconnectNetworkDriveDialog();
        break;
    case ID_CTX_FIXDRIVE:
        FixSelectedBrokenDrive();
        break;

    default:
        break;
    }
}

// ----------------------------- Window proc

LRESULT CALLBACK WndProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    switch (m) {
    case WM_CREATE: {
        g_hwndMain = h;
        INITCOMMONCONTROLSEX icc; icc.dwSize = sizeof(icc);
        icc.dwICC = ICC_LISTVIEW_CLASSES | ICC_BAR_CLASSES;
        InitCommonControlsEx(&icc);

        InitializeCriticalSection(&g_metaLock);

        g_hwndList = CreateWindowExW(
            WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"",
            WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS,
            0, 0, 100, 100, h, (HMENU)1001, g_hInst, NULL);
        ListView_SetExtendedListViewStyle(
            g_hwndList,
            LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER |
            LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
        LV_ResetColumns();
        SetWindowSubclass(g_hwndList, ListSubclass, 1, 0);

        g_hwndVideo = CreateWindowExW(
            0, L"STATIC", L"",
            WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN,
            0, 0, 100, 100, h, (HMENU)1002, g_hInst, NULL);
        ShowWindow(g_hwndVideo, SW_HIDE);
        SetWindowSubclass(g_hwndVideo, VideoSubclass, 2, 0);

        g_hwndSeek = CreateWindowExW(
            0, TRACKBAR_CLASSW, L"",
            WS_CHILD | TBS_HORZ | TBS_AUTOTICKS,
            0, 0, 100, 30, h, (HMENU)1003, g_hInst, NULL);
        ShowWindow(g_hwndSeek, SW_HIDE);
        SetWindowSubclass(g_hwndSeek, SeekSubclass, 3, 0);

        // Initial view: optional start folder from command line, otherwise drives
        if (!g_initialPath.empty()) {
            DWORD attrs = GetFileAttributesW(g_initialPath.c_str());
            if (attrs != INVALID_FILE_ATTRIBUTES) {
                if (attrs & FILE_ATTRIBUTE_DIRECTORY) {
                    ShowFolder(g_initialPath);
                }
                else {
                    std::wstring folder = g_initialPath;
                    PathRemoveFileSpecW(&folder[0]);
                    folder = folder.c_str();
                    ShowFolder(folder);
                }
            }
            else {
                ShowDrives();
            }
        }
        else {
            ShowDrives();
        }
        // Make sure the title reflects the initial view (e.g. "C:\temp\")
        SetTitleFolderOrDrives();
        return 0;
    }

    case WM_SIZE:
        OnSize(LOWORD(l), HIWORD(l));
        return 0;

    case WM_SETFOCUS:
        if (g_inPlayback) SetFocus(g_hwndVideo);
        else SetFocus(g_hwndList);
        return 0;

    case WM_NOTIFY: {
        LPNMHDR nm = (LPNMHDR)l;
        if (nm->hwndFrom == g_hwndList) {
            if (nm->code == NM_CUSTOMDRAW) {
                return HandleListCustomDraw((NMLVCUSTOMDRAW*)l);
            }
            if (nm->code == LVN_ITEMACTIVATE) {
                ActivateSelection();
                return 0;
            }
            if (nm->code == LVN_COLUMNCLICK) {
                LPNMLISTVIEW p = reinterpret_cast<LPNMLISTVIEW>(l);
                if (p->iSubItem == g_sortCol) g_sortAsc = !g_sortAsc;
                else { g_sortCol = p->iSubItem; g_sortAsc = true; }
                SendMessageW(g_hwndList, WM_SETREDRAW, FALSE, 0);
                SortRows(g_sortCol, g_sortAsc);
                SendMessageW(g_hwndList, WM_SETREDRAW, TRUE, 0);
                InvalidateRect(g_hwndList, NULL, TRUE);
                return 0;
            }
        }
        break;
    }

    case WM_HSCROLL:
        if ((HWND)l == g_hwndSeek && g_inPlayback && g_mp) {
            int code = LOWORD(w);
            if (code == TB_THUMBTRACK) {
                g_userDragging = true;
            }
            else if (code == TB_ENDTRACK || code == TB_THUMBPOSITION) {
                g_userDragging = false;
                LRESULT pos = SendMessageW(g_hwndSeek, TBM_GETPOS, 0, 0);
                libvlc_media_player_set_time(g_mp, (libvlc_time_t)pos);
            }
            return 0;
        }
        break;

    case WM_TIMER:
        if (w == kTimerPlaybackUI && g_inPlayback && g_mp) {
            libvlc_time_t len = libvlc_media_player_get_length(g_mp);
            libvlc_time_t cur = libvlc_media_player_get_time(g_mp);
            if (len != g_lastLenForRange && len > 0) {
                g_lastLenForRange = len;
                libvlc_time_t range = len; if (range > INT_MAX) range = INT_MAX;
                SendMessageW(g_hwndSeek, TBM_SETRANGEMIN, TRUE, 0);
                SendMessageW(g_hwndSeek, TBM_SETRANGEMAX, TRUE, (LPARAM)range);
            }
            if (!g_userDragging) {
                libvlc_time_t p = cur; if (p > INT_MAX) p = INT_MAX;
                SendMessageW(g_hwndSeek, TBM_SETPOS, TRUE, (LPARAM)p);
            }
            SetTitlePlaying();
            return 0;
        }
        break;

    case WM_KEYDOWN:
        if (w == VK_F1) { ShowHelp(); return 0; }
        if (g_inPlayback) {
            SendMessageW(g_hwndVideo, WM_KEYDOWN, w, l);
            return 0;
        }
        break;

    case WM_CONTEXTMENU:
        if ((HWND)w == g_hwndList) {
            POINT pt;
            pt.x = GET_X_LPARAM(l);
            pt.y = GET_Y_LPARAM(l);

            if (pt.x == -1 && pt.y == -1) {
                int sel = ListView_GetNextItem(
                    g_hwndList, -1, LVNI_FOCUSED | LVNI_SELECTED);
                RECT rc;
                if (sel >= 0) {
                    ListView_GetItemRect(g_hwndList, sel, &rc, LVIR_BOUNDS);
                    pt.x = rc.left + 10;
                    pt.y = rc.top + 10;
                }
                else {
                    GetClientRect(g_hwndList, &rc);
                    pt.x = rc.left + 10;
                    pt.y = rc.top + 10;
                }
                ClientToScreen(g_hwndList, &pt);
            }

            ShowListContextMenu(pt);
            return 0;
        }
        break;

    case WM_APP + 1:
        if (g_inPlayback && g_playlistIndex + 1 < g_playlist.size())
            NextInPlaylist();
        else if (g_inPlayback)
            ExitPlayback();
        return 0;

    case WM_APP_META: {
        MetaResult* r = (MetaResult*)l;
        if (r) {
            if (r->gen == g_metaGen.load(std::memory_order_relaxed)) {
                for (int i = 0; i < (int)g_rows.size(); ++i) {
                    if (_wcsicmp(g_rows[i].full.c_str(), r->path.c_str()) == 0) {
                        Row& it = g_rows[i];
                        it.vW = r->w;
                        it.vH = r->h;
                        it.vDur100ns = r->dur;
                        if (!it.isDir && IsVideoFile(it.full)) {
                            if (it.vW > 0 && it.vH > 0) {
                                wchar_t buf[64];
                                swprintf_s(buf, L"%dx%d", it.vW, it.vH);
                                ListView_SetItemText(g_hwndList, i, 4, buf);
                            }
                            if (it.vDur100ns > 0) {
                                std::wstring ds = FormatDuration100ns(it.vDur100ns);
                                ListView_SetItemText(g_hwndList, i, 5,
                                    const_cast<wchar_t*>(ds.c_str()));
                            }
                        }
                        break;
                    }
                }
            }
            delete r;
        }
        return 0;
    }

    case WM_CLOSE:
        if (g_loadingFolder) {
            MessageBoxW(h, L"Loading folder... please wait.", L"Browse", MB_OK);
            return 0;
        }
        DestroyWindow(h);
        return 0;

    case WM_DESTROY:
        KillTimer(h, kTimerPlaybackUI);

        CancelMetaWorkAndClearTodo();
        if (g_metaThread) {
            WaitForSingleObject(g_metaThread, 200);
            CloseHandle(g_metaThread);
            g_metaThread = NULL;
        }

        DeleteCriticalSection(&g_metaLock);

        if (g_mp) { libvlc_media_player_stop(g_mp); libvlc_media_player_release(g_mp); g_mp = NULL; }
        if (g_vlc) { libvlc_release(g_vlc); g_vlc = NULL; }
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(h, m, w, l);
}

// ----------------------------- Entry

int APIENTRY wWinMain(HINSTANCE hInst, HINSTANCE, LPWSTR, int nShow) {
    g_hInst = hInst;

    // parse optional command-line start folder
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv) {
        if (argc >= 2) {
            g_initialPath = argv[1];
        }
        LocalFree(argv);
    }

    HMODULE u = GetModuleHandleW(L"user32.dll");
    if (u) {
        typedef BOOL(WINAPI* SetProcessDPIAware_t)();
        SetProcessDPIAware_t setAw =
            (SetProcessDPIAware_t)GetProcAddress(u, "SetProcessDPIAware");
        if (setAw) setAw();
    }

    LoadConfigFromIni();
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    int bigW = GetSystemMetrics(SM_CXICON), bigH = GetSystemMetrics(SM_CYICON);
    int smW = GetSystemMetrics(SM_CXSMICON), smH = GetSystemMetrics(SM_CYSMICON);
    HICON hBig = LoadAppIcon(bigW, bigH);
    HICON hSm = LoadAppIcon(smW, smH);

    WNDCLASSEXW wc; ZeroMemory(&wc, sizeof(wc));
    wc.cbSize = sizeof(wc);
    wc.hInstance = hInst;
    wc.lpszClassName = L"BrowseWindowClass";
    wc.lpfnWndProc = WndProc;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hIcon = hBig ? hBig : LoadIcon(NULL, IDI_APPLICATION);
    wc.hIconSm = hSm ? hSm : wc.hIcon;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    RegisterClassExW(&wc);

    g_hwndMain = CreateWindowExW(
        0, wc.lpszClassName, L"Browse ",
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        CW_USEDEFAULT, CW_USEDEFAULT, 1500, 700,
        NULL, NULL, hInst, NULL);

    ShowWindow(g_hwndMain, nShow);
    UpdateWindow(g_hwndMain);

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    CoUninitialize();
    return 0;
}
