// MediaExplorer - Win32 file browser + libVLC player with recursive search
// Build: VS2022, x64, C++17  (fast folder load + async video metadata fill + cut/paste UI progress)
//
// Key speed-ups:
//  1) FindFirstFileExW(..., FindExInfoBasic, ..., FIND_FIRST_EX_LARGE_FETCH)
//  2) ListView redraw suspension during bulk insert
//  3) Deferred video metadata (resolution/duration) with fast cached try,
//     then background worker fills remaining cells; cancelled on navigation.
// New in this file:
//  4) Ctrl+X removes selected files from view and fills app clipboard
//  5) Ctrl+V shows sub-modal progress window "copy/move <file>..." + Cancel
//     - cancel immediately aborts the whole batch and drops the window
//     - clipboard is cleared when the window closes
//  6) Ctrl+P during playback: show ffprobe-based video properties
//  7) Ctrl+Up / Ctrl+Down: reorder single selected row in list
//  8) Ctrl+Plus: combine selected files via video_combine.exe in background threads
//     with per-task log windows; app cannot exit until all combines finish.

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
#include <commctrl.h>
#include <shlwapi.h>
#include <shobjidl_core.h>
#include <propsys.h>
#include <propkey.h>
#include <wrl/client.h>

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
#include <shlobj.h>     // SHCreateDirectoryExW
#include <cstdarg>
#include <io.h> // for _unlink


#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "User32.lib")
#pragma comment(lib, "Gdi32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Propsys.lib")
#pragma comment(lib, "Uuid.lib")

#include <vlc/vlc.h>

#ifndef FIND_FIRST_EX_LARGE_FETCH
// Older SDKs may miss this flag; define to 0 (ignored) to stay compatible.
#  define FIND_FIRST_EX_LARGE_FETCH 0x00000002
#endif

using Microsoft::WRL::ComPtr;

static void LogLine(const wchar_t* fmt, ...); // forward declaration

// ----------------------------- Globals
HINSTANCE g_hInst = NULL;
HWND g_hwndMain = NULL, g_hwndList = NULL, g_hwndVideo = NULL, g_hwndSeek = NULL;

enum class ViewKind { Drives, Folder, Search };
ViewKind g_view = ViewKind::Drives;
std::wstring g_folder; // valid in Folder view, ends with '\'

struct Row {
    std::wstring name;     // display (for Search, full path; for Folder, file name)
    std::wstring full;     // absolute path (dir ends with '\')
    bool         isDir;
    ULONGLONG    size;
    FILETIME     modified;
    // video props
    int          vW, vH;
    ULONGLONG    vDur100ns;
    Row() : isDir(false), size(0), vW(0), vH(0), vDur100ns(0) { modified.dwLowDateTime = modified.dwHighDateTime = 0; }
};
std::vector<Row> g_rows;

// sorting
int  g_sortCol = 0;      // 0=Name,1=Type,2=Size,3=Modified,4=Resolution,5=Duration
bool g_sortAsc = true;

// VLC
libvlc_instance_t* g_vlc = NULL;
libvlc_media_player_t* g_mp = NULL;
bool                      g_inPlayback = false;
std::vector<std::wstring> g_playlist;
size_t                    g_playlistIndex = 0;
bool                      g_userDragging = false;
libvlc_time_t             g_lastLenForRange = -1;

// ----------------------------- Configuration (mediaexplorer.ini)

struct AppConfig {
    std::wstring upscaleDirectory;  // e.g. w:\upscale\autosubmit (may be empty)
    bool ffmpegAvailable = false;  // if true, ffmpeg-based tools are enabled & shown in help
    bool ffprobeAvailable = false;  // if true, ffprobe-based info is enabled & shown in help
//    bool videoCombineAvailable = false;   // if true, Ctrl+Plus / video_combine.exe is enabled

    bool        loggingEnabled = false; // master on/off
    std::wstring loggingPath;          // folder from INI
    std::wstring logFile;             // full path to mediaexplorer.log
};

AppConfig g_cfg;



// fullscreen (app-managed)
bool g_fullscreen = false;
WINDOWPLACEMENT g_wpPrev; // set .length before use

// timers
const UINT_PTR kTimerPlaybackUI = 1;

// post-playback actions
enum class ActionType { DeleteFile, RenameFile, CopyToPath };
struct PostAction { ActionType type; std::wstring src; std::wstring param; };
std::vector<PostAction> g_post;

// filename clipboard for browser
enum class ClipMode { None, Copy, Move };
ClipMode g_clipMode = ClipMode::None;
std::vector<std::wstring> g_clipFiles; // absolute file paths

// ----------------------------- Search state
struct SearchState {
    bool active;
    ViewKind originView;
    std::wstring originFolder;               // empty if origin was Drives
    std::vector<std::wstring> termsLower;    // intersection terms

    // selection-aware explicit scope
    bool useExplicitScope;
    std::vector<std::wstring> explicitFolders; // each ends with '\'
    std::vector<std::wstring> explicitFiles;   // absolute file paths

    SearchState()
        : active(false),
        originView(ViewKind::Drives),
        useExplicitScope(false) {
    }
} g_search;

// ----------------------------- Async metadata fill
constexpr UINT WM_APP_META = WM_APP + 100;
constexpr UINT WM_APP_COMBINE_OUTPUT = WM_APP + 200;
constexpr UINT WM_APP_COMBINE_DONE = WM_APP + 201;
// New for ffmpeg tools
constexpr UINT WM_APP_FFMPEG_OUTPUT = WM_APP + 300;
constexpr UINT WM_APP_FFMPEG_DONE = WM_APP + 301;

struct MetaResult {
    std::wstring path;
    int          w, h;
    ULONGLONG    dur;
    uint32_t     gen;
};

std::atomic<uint32_t> g_metaGen{ 0 };
CRITICAL_SECTION      g_metaLock;          // protects g_metaTodoPaths
std::vector<std::wstring> g_metaTodoPaths; // paths that still need deep props
HANDLE                g_metaThread = NULL;

// ----------------------------- Combine tasks (video_combine in background)

struct CombineTask {
    HANDLE hThread;
    HANDLE hProcess;
    HWND   hwnd;      // log window
    HWND   hEdit;     // multiline read-only edit inside log window
    std::wstring workingDir;    // dir where inputs are copied
    std::vector<std::wstring> srcFiles;      // original source file paths
    std::wstring combinedFull;  // final combined video path
    std::wstring title;         // short description (e.g., output file name)
    bool   running;

    CombineTask() :
        hThread(NULL),
        hProcess(NULL),
        hwnd(NULL),
        hEdit(NULL),
        running(false) {
    }
};

CRITICAL_SECTION g_combineLock;
std::vector<CombineTask*> g_combineTasks;

// forward declarations for combine-related helpers
static void EnsureCombineLogClass();
static HWND CreateCombineLogWindow(CombineTask* task);
static void PostCombineOutput(CombineTask* task, const std::wstring& text);
static DWORD WINAPI CombineThreadProc(LPVOID param);

static bool HasRunningCombineTasks() {
    EnterCriticalSection(&g_combineLock);
    bool any = false;
    for (CombineTask* t : g_combineTasks) {
        if (t && t->running) { any = true; break; }
    }
    LeaveCriticalSection(&g_combineLock);
    return any;
}

// ----------------------------- FFmpeg processing tasks (trim/flip in background)

enum class FfmpegOpKind { TrimFront, TrimEnd, HFlip };

struct FfmpegTask {
    HANDLE hThread = NULL;
    HANDLE hProcess = NULL;
    HWND   hwnd = NULL;        // log window
    HWND   hEdit = NULL;       // multiline read-only edit in log window
    std::wstring sourceFull;   // original video path
    std::wstring workingDir;   // ...video_process
    std::wstring inputCopy;    // workingDir + base.ext (copied original)
    std::wstring outputTemp;   // workingDir + base_<op>.ext  (ffmpeg output)
    std::wstring finalWorking; // after rename, path of final processed file in workingDir
    std::wstring title;        // short title, e.g. "Trim front: file.mp4"
    FfmpegOpKind kind;
    libvlc_time_t refMs = 0;   // time in ms when user invoked operation
    bool running = false;
    bool done = false;
    DWORD exitCode = 0;
};

CRITICAL_SECTION g_ffLock;
std::vector<FfmpegTask*> g_ffTasks;

static bool HasRunningFfmpegTasks() {
    EnterCriticalSection(&g_ffLock);
    bool any = false;
    for (FfmpegTask* t : g_ffTasks) {
        if (t && t->running) { any = true; break; }
    }
    LeaveCriticalSection(&g_ffLock);
    return any;
}

// ----------------------------- FFmpeg task log window + helpers

static LRESULT CALLBACK FfmpegLogProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    FfmpegTask* task = reinterpret_cast<FfmpegTask*>(GetWindowLongPtrW(h, GWLP_USERDATA));
    switch (m) {
    case WM_CREATE: {
        LPCREATESTRUCT pcs = (LPCREATESTRUCT)l;
        task = (FfmpegTask*)pcs->lpCreateParams;
        SetWindowLongPtrW(h, GWLP_USERDATA, (LONG_PTR)task);

        HFONT hf = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        RECT rc; GetClientRect(h, &rc);

        HWND hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
            4, 4, rc.right - 8, rc.bottom - 8, h, (HMENU)101, g_hInst, NULL);
        SendMessageW(hEdit, WM_SETFONT, (WPARAM)hf, TRUE);

        if (task) task->hEdit = hEdit;
        return 0;
    }
    case WM_SIZE: {
        if (task && task->hEdit) {
            RECT rc; GetClientRect(h, &rc);
            MoveWindow(task->hEdit, 4, 4, rc.right - 8, rc.bottom - 8, TRUE);
        }
        return 0;
    }
    case WM_CLOSE:
        ShowWindow(h, SW_HIDE); // hide, don't destroy; keep log available
        return 0;
    }
    return DefWindowProcW(h, m, w, l);
}

static void EnsureFfmpegLogClass() {
    static bool reg = false;
    if (!reg) {
        WNDCLASSW wc{}; wc.lpfnWndProc = FfmpegLogProc; wc.hInstance = g_hInst;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"FfmpegLogClass";
        RegisterClassW(&wc);
        reg = true;
    }
}

static HWND CreateFfmpegLogWindow(FfmpegTask* task) {
    RECT r; SystemParametersInfoW(SPI_GETWORKAREA, 0, &r, 0);
    int W = 640, H = 480;
    int X = r.left + ((r.right - r.left) - W) / 2;
    int Y = r.top + ((r.bottom - r.top) - H) / 2;

    std::wstring title = L"FFmpeg task: ";
    title += task ? task->title : L"(unknown)";

    HWND hwnd = CreateWindowExW(WS_EX_TOOLWINDOW, L"FfmpegLogClass", title.c_str(),
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        X, Y, W, H, g_hwndMain, NULL, g_hInst, task);
    return hwnd;
}

static void PostFfmpegOutput(FfmpegTask* task, const std::wstring& text) {
    if (!task) return;
    std::wstring* p = new std::wstring(text);
    PostMessageW(g_hwndMain, WM_APP_FFMPEG_OUTPUT, (WPARAM)task, (LPARAM)p);
}

static DWORD WINAPI FfmpegThreadProc(LPVOID param) {
    FfmpegTask* task = (FfmpegTask*)param;
    if (!task) return 0;

    LogLine(L"FFmpegTask start: src=\"%s\" inputCopy=\"%s\" outputTemp=\"%s\" kind=%d refMs=%lld",
        task->sourceFull.c_str(), task->inputCopy.c_str(),
        task->outputTemp.c_str(), (int)task->kind, (long long)task->refMs);

    // 1) Create working directory
    if (!CreateDirectoryW(task->workingDir.c_str(), NULL)) {
        DWORD e = GetLastError();
        if (e != ERROR_ALREADY_EXISTS) {
            std::wstring msg = L"ERROR: Failed to create working directory:\r\n";
            msg += task->workingDir;
            msg += L"\r\n";
            PostFfmpegOutput(task, msg);
            task->exitCode = 1;
            task->running = false;
            PostMessageW(g_hwndMain, WM_APP_FFMPEG_DONE, (WPARAM)task, (LPARAM)task->exitCode);
            return 0;
        }
    }

    // Paths should already be filled in, but ensure they’re non-empty.
    if (task->inputCopy.empty() || task->outputTemp.empty()) {
        PostFfmpegOutput(task, L"ERROR: task paths are not initialized.\r\n");
        task->exitCode = 2;
        task->running = false;
        PostMessageW(g_hwndMain, WM_APP_FFMPEG_DONE, (WPARAM)task, (LPARAM)task->exitCode);
        return 0;
    }

    // 2) Copy input into working directory
    {
        std::wstring msg = L"Copying input to working directory:\r\n  ";
        msg += task->inputCopy;
        msg += L"\r\n";
        PostFfmpegOutput(task, msg);

        if (!CopyFileW(task->sourceFull.c_str(), task->inputCopy.c_str(), FALSE)) {
            std::wstring err = L"ERROR: Failed to copy file:\r\n  ";
            err += task->sourceFull;
            err += L"\r\n";
            PostFfmpegOutput(task, err);
            task->exitCode = 3;
            task->running = false;
            PostMessageW(g_hwndMain, WM_APP_FFMPEG_DONE, (WPARAM)task, (LPARAM)task->exitCode);
            return 0;
        }
    }

    // 3) Build ffmpeg command line
    double seconds = (double)task->refMs / 1000.0;
    wchar_t secBuf[64];
    swprintf_s(secBuf, L"%.3f", seconds);

    std::wstring cmd = L"ffmpeg -y ";
    switch (task->kind) {
    case FfmpegOpKind::TrimFront:
        // Keep from refMs -> end
        cmd += L"-ss ";
        cmd += secBuf;
        cmd += L" -i \"";
        cmd += task->inputCopy;
        cmd += L"\" -c copy \"";
        cmd += task->outputTemp;
        cmd += L"\"";
        break;
    case FfmpegOpKind::TrimEnd:
        // Keep from 0 -> refMs
        cmd += L"-i \"";
        cmd += task->inputCopy;
        cmd += L"\" -t ";
        cmd += secBuf;
        cmd += L" -c copy \"";
        cmd += task->outputTemp;
        cmd += L"\"";
        break;
    case FfmpegOpKind::HFlip:
        cmd += L"-i \"";
        cmd += task->inputCopy;
        cmd += L"\" -vf hflip -c:a copy \"";
        cmd += task->outputTemp;
        cmd += L"\"";
        break;
    }

    PostFfmpegOutput(task, L"Running command:\r\n");
    PostFfmpegOutput(task, cmd + L"\r\n\r\n");

    SECURITY_ATTRIBUTES sa{};
    sa.nLength = sizeof(sa);
    sa.bInheritHandle = TRUE;
    sa.lpSecurityDescriptor = NULL;

    HANDLE hRead = NULL, hWrite = NULL;
    if (!CreatePipe(&hRead, &hWrite, &sa, 0)) {
        PostFfmpegOutput(task, L"ERROR: Failed to create pipe.\r\n");
        task->exitCode = 4;
        task->running = false;
        PostMessageW(g_hwndMain, WM_APP_FFMPEG_DONE, (WPARAM)task, (LPARAM)task->exitCode);
        return 0;
    }
    SetHandleInformation(hRead, HANDLE_FLAG_INHERIT, 0);

    STARTUPINFOW si{};
    si.cb = sizeof(si);
    si.dwFlags |= STARTF_USESTDHANDLES;
    si.hStdOutput = hWrite;
    si.hStdError = hWrite;

    PROCESS_INFORMATION pi{};
    std::vector<wchar_t> cmdBuf(cmd.size() + 1);
    wcscpy_s(cmdBuf.data(), cmdBuf.size(), cmd.c_str());

    BOOL ok = CreateProcessW(NULL, cmdBuf.data(),
        NULL, NULL, TRUE,
        CREATE_NO_WINDOW,
        NULL, NULL,
        &si, &pi);
    CloseHandle(hWrite);
    hWrite = NULL;

    if (!ok) {
        PostFfmpegOutput(task, L"ERROR: Failed to start ffmpeg.\r\n");
        CloseHandle(hRead);
        task->exitCode = 5;
        task->running = false;
        PostMessageW(g_hwndMain, WM_APP_FFMPEG_DONE, (WPARAM)task, (LPARAM)task->exitCode);
        return 0;
    }

    task->hProcess = pi.hProcess;
    CloseHandle(pi.hThread);

    // 4) Read ffmpeg stdout/stderr
    char buf[4096];
    DWORD bytes = 0;
    std::string accum;

    while (ReadFile(hRead, buf, sizeof(buf), &bytes, NULL) && bytes > 0) {
        accum.append(buf, buf + bytes);
        size_t pos = 0;
        while (true) {
            size_t nl = accum.find('\n', pos);
            if (nl == std::string::npos) {
                accum.erase(0, pos);
                break;
            }
            std::string line = accum.substr(pos, nl - pos + 1);
            pos = nl + 1;

            int n = MultiByteToWideChar(CP_ACP, 0, line.c_str(), (int)line.size(), NULL, 0);
            if (n <= 0) continue;
            std::wstring wline(n, L'\0');
            MultiByteToWideChar(CP_ACP, 0, line.c_str(), (int)line.size(), &wline[0], n);
            PostFfmpegOutput(task, wline);
        }
    }
    CloseHandle(hRead);

    WaitForSingleObject(task->hProcess, INFINITE);
    DWORD exitCode = 0;
    GetExitCodeProcess(task->hProcess, &exitCode);
    CloseHandle(task->hProcess);
    task->hProcess = NULL;

    wchar_t doneMsg[128];
    swprintf_s(doneMsg, L"\r\n[ffmpeg exited with code %lu]\r\n", exitCode);
    PostFfmpegOutput(task, doneMsg);

    task->exitCode = exitCode;

    // On success, rename outputTemp -> inputCopy (so finalWorking is the "base.ext" in video_process)
    if (exitCode == 0) {
        // Delete original copy, then rename
        DeleteFileW(task->inputCopy.c_str());
        MoveFileExW(task->outputTemp.c_str(), task->inputCopy.c_str(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_COPY_ALLOWED);
        task->finalWorking = task->inputCopy;
    }

    task->running = false;
    task->done = true;
    PostMessageW(g_hwndMain, WM_APP_FFMPEG_DONE, (WPARAM)task, (LPARAM)exitCode);
    LogLine(L"FFmpegTask done: src=\"%s\" exitCode=%lu finalWorking=\"%s\"",
        task->sourceFull.c_str(), exitCode, task->finalWorking.c_str());
    return 0;
}

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
static void CollectSelection(std::vector<std::wstring>& outFolders, std::vector<std::wstring>& outFiles) {
    outFolders.clear(); outFiles.clear();
    int idx = -1;
    while ((idx = ListView_GetNextItem(g_hwndList, idx, LVNI_SELECTED)) != -1) {
        if (idx < 0 || idx >= (int)g_rows.size()) continue;
        const Row& r = g_rows[idx];
        if (r.isDir) {
            outFolders.push_back(EnsureSlash(r.full)); // works for drives & folders
        }
        else {
            outFiles.push_back(r.full);
        }
    }
}


// --- Search progress title + gentle UI pumping during long recursion ---
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
    std::wstring t = L"Media Explorer (libVLC) - searching ";
    t += EnsureSlash(folder);
    SetWindowTextW(g_hwndMain, t.c_str());
    PumpMessagesThrottled(50);
}

static inline std::wstring ParentDir(std::wstring p) {
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

// ----------------------------- Logging helpers

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

    // Create directory tree (best-effort)
    int rc = SHCreateDirectoryExW(NULL, folder.c_str(), NULL);
    if (rc != ERROR_SUCCESS && rc != ERROR_ALREADY_EXISTS && rc != ERROR_FILE_EXISTS) {
        // Cannot create folder -> disable logging
        g_cfg.loggingEnabled = false;
        return;
    }

    g_cfg.loggingPath = folder;
    g_cfg.logFile = folder + L"mediaexplorer.log";
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
    wchar_t buf[64]; swprintf_s(buf, L"%.2f %s", v, u[i]); return buf;
}
static std::wstring FormatFileTime(const FILETIME& ft) {
    SYSTEMTIME utc, loc; FileTimeToSystemTime(&ft, &utc);
    SystemTimeToTzSpecificLocalTime(NULL, &utc, &loc);
    wchar_t buf[64]; swprintf_s(buf, L"%04u-%02u-%02u %02u:%02u",
        loc.wYear, loc.wMonth, loc.wDay, loc.wHour, loc.wMinute); return buf;
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
    size_t dot = p.find_last_of(L'.'); if (dot == std::wstring::npos) return L"";
    std::wstring e = p.substr(dot); std::transform(e.begin(), e.end(), e.begin(), ::towlower); return e;
}
static bool IsVideoFile(const std::wstring& path) {
    static const wchar_t* exts[] = {
        L".mp4", L".mkv", L".mov", L".avi", L".wmv", L".m4v", L".ts", L".m2ts", L".webm", L".flv", L".rm",
    };
    std::wstring e = ExtLower(path);
    for (size_t i = 0; i < _countof(exts); ++i) if (e == exts[i]) return true;
    return false;
}

// Fast cached attempt (no I/O if system cache has props)
static bool GetVideoPropsFastCached(const std::wstring& path, int& outW, int& outH, ULONGLONG& outDur100ns) {
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

// Full property read (may hit disk); used by worker
static bool GetVideoProps(const std::wstring& path, int& outW, int& outH, ULONGLONG& outDur100ns) {
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

// Title
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
        swprintf_s(buf, L"(Play List %zu of %zu) ", g_playlistIndex + 1, g_playlist.size());
        left = buf;
    }

    std::wstring t = left;
    t += base; t += L"  ";
    t += FormatHMSms(cur); t += L" / "; t += FormatHMSms(len);
    SetWindowTextW(g_hwndMain, t.c_str());
}
static std::wstring JoinTermsForTitle() {
    if (!g_search.active || g_search.termsLower.empty()) return L"";
    std::wstring s = L"\""; s += g_search.termsLower[0]; s += L"\"";
    for (size_t i = 1; i < g_search.termsLower.size(); ++i) {
        s += L" & \""; s += g_search.termsLower[i]; s += L"\"";
    }
    return s;
}
static void SetTitleFolderOrDrives() {
    std::wstring t = L"Media Explorer (libVLC) - ";
    if (g_view == ViewKind::Drives) t += L"[Drives]";
    else if (g_view == ViewKind::Folder) t += EnsureSlash(g_folder);
    else t += L"Search - " + JoinTermsForTitle();
    SetWindowTextW(g_hwndMain, t.c_str());
}

// ----------------------------- ffprobe helpers (video properties during playback)

// Run a wide command line through ffprobe and collect stdout lines
static bool RunFfprobeCommand(const std::wstring& cmdLine, std::vector<std::string>& outLines) {
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

// Query width, height, video codec, audio codec for a given file via ffprobe.
// Query width, height, video codec, audio codec for a given file via ffprobe.
// Uses key=value output for robust parsing.
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

    // ---------------- Video stream: width, height, codec_name ----------------
    //
    // We ask ffprobe to print ONLY:
    //   codec_name=...
    //   width=...
    //   height=...
    //
    // Example lines:
    //   codec_name=h264
    //   width=576
    //   height=768
    //
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

            // codec_name=...
            const char* kCodec = "codec_name=";
            const size_t codecLen = sizeof("codec_name=") - 1;
            if (line.size() >= codecLen && line.compare(0, codecLen, kCodec) == 0) {
                codecV = line.substr(codecLen);
                continue;
            }

            // width=...
            const char* kWidth = "width=";
            const size_t widthLen = sizeof("width=") - 1;
            if (line.size() >= widthLen && line.compare(0, widthLen, kWidth) == 0) {
                wTmp = std::strtol(line.c_str() + widthLen, nullptr, 10);
                continue;
            }

            // height=...
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
            outVideoCodec.assign(codecV.begin(), codecV.end()); // ASCII-safe
        }

        gotV = (wTmp > 0 || hTmp > 0 || !codecV.empty());
    }

    // ---------------- Audio stream: codec_name ----------------
    //
    // Example line:
    //   codec_name=aac
    //
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

// Show MessageBox with media properties for the currently playing item
static void ShowCurrentVideoProperties() {
    if (!g_inPlayback || g_playlist.empty()) {
        MessageBoxW(g_hwndMain, L"No video is currently playing.", L"Video properties",
            MB_OK);
        return;
    }

    const std::wstring& full = g_playlist[g_playlistIndex];

    // 1) Start with resolution from Shell (same as file list) so we always have
    //    something sensible even if ffprobe parsing fails.
    int wShell = 0, hShell = 0;
    ULONGLONG durDummy = 0;
    GetVideoPropsFastCached(full, wShell, hShell, durDummy);  // ignores if it fails

    int w = wShell;
    int h = hShell;
    std::wstring vCodec, aCodec;

    // Pause playback BEFORE doing ffprobe and BEFORE showing the dialog
    bool wasPlaying = (g_mp && libvlc_media_player_is_playing(g_mp) > 0);
    if (g_mp && wasPlaying) {
        libvlc_media_player_set_pause(g_mp, 1);
    }

    bool okFF = false;
    if (g_cfg.ffprobeAvailable) {
        LogLine(L"ffprobe: querying \"%s\"", full.c_str());
        okFF = GetMediaInfoFromFfprobe(full, w, h, vCodec, aCodec);
        LogLine(L"ffprobe: \"%s\" result ok=%d w=%d h=%d vCodec=\"%s\" aCodec=\"%s\"",
            full.c_str(), okFF ? 1 : 0, w, h,
            vCodec.c_str(), aCodec.c_str());
    }

    // If ffprobe didn't give us good values, fall back to Shell result
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
        msg += L"\nNote: ffprobe.exe did not return information "
            L"(not found in PATH or error running command).";
    }
    else if (!g_cfg.ffprobeAvailable) {
        msg += L"\nNote: ffprobe-based details are disabled in mediaexplorer.ini.";
    }

    // Show dialog WHILE paused
    MessageBoxW(g_hwndMain, msg.c_str(), L"Video properties", MB_OK);

    // Resume playback only AFTER dialog closes
    if (g_mp && wasPlaying) {
        libvlc_media_player_set_pause(g_mp, 0);
    }
}

static void LoadConfigFromIni() {
    wchar_t exePath[MAX_PATH] = {};
    if (!GetModuleFileNameW(NULL, exePath, MAX_PATH)) return;
    PathRemoveFileSpecW(exePath);

    std::wstring iniPath = exePath;
    iniPath += L"\\mediaexplorer.ini";

    std::wifstream in(iniPath);
    if (!in) {
        // No .ini file -> all defaults (ffmpeg/ffprobe disabled, no upscaleDirectory)
        return;
    }

    std::wstring line;
    while (std::getline(in, line)) {
        line = Trim(line);
        if (line.empty()) continue;
        if (line[0] == L';' || line[0] == L'#') continue;
        if (line.front() == L'[' && line.back() == L']') continue; // ignore sections

        // Strip inline comments after ';'
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
    }
    // after the while(...) loop
    InitLoggingFromConfig();
    if (g_cfg.loggingEnabled) 
    {
        LogLine(L"Config: upscale=\"%s\" ffmpeg=%d ffprobe=%d loggingPath=\"%s\"",
            g_cfg.upscaleDirectory.c_str(),
            g_cfg.ffmpegAvailable ? 1 : 0,
            g_cfg.ffprobeAvailable ? 1 : 0,
            g_cfg.loggingPath.c_str());
    }
}


// Help
// Help
static void ShowHelp() {
    // If a libVLC media player exists and is currently playing,
    // pause it while the help window is visible.
    bool wasPlaying = false;
    if (g_mp) {
        wasPlaying = (libvlc_media_player_is_playing(g_mp) > 0);
        if (wasPlaying) {
            libvlc_media_player_set_pause(g_mp, 1);
        }
    }

    std::wstring msg;
    msg += L"Media Explorer - Help\n\n";

    msg += L"CONFIGURATION (mediaexplorer.ini)\n"
        L"  upscaledirectory = w:\\upscale\\autosubmit\n"
        L"  ffmpegAvailable  = 0|1  (enable FFmpeg tools: trim / flip)\n"
        L"  ffprobeAvailable = 0|1  (enable ffprobe-based details)\n\n";

    msg += L"FILE BROWSER (list)\n"
        L"  Enter / Double-click : Open folder / Play selected video(s)\n"
        L"  Left / Backspace     : Up one folder (from root -> drives) / Exit search\n"
        L"  Click column header  : Sort (folders always first)\n"
        L"  Ctrl+A               : Select all videos in current view\n"
        L"  Ctrl+P               : Play selected videos\n"
        L"  Ctrl+F               : Search (recursive). In Search view: refine (AND/intersection)\n"
        L"  Ctrl+Up/Down         : Move selected row up/down (single selection)\n";

    if (g_cfg.ffmpegAvailable) {
        msg += L"  Ctrl+Plus            : Combine selected files into one video (background)\n";
    }

    msg += L"  Ctrl+C / Ctrl+X / Ctrl+V : Copy / Cut / Paste files\n"
        L"  Del                  : Delete selected files\n"
        L"  F1                   : Help\n\n";

    msg += L"PLAYBACK\n"
        L"  Enter                : Toggle fullscreen\n"
        L"  Esc                  : Exit playback (applies queued actions & FFmpeg tasks)\n"
        L"  Space / Tab          : Pause / Resume\n"
        L"  Left / Right         : Seek -/+10s (hold Shift: -/+60s)\n"
        L"  Ctrl+Left / Ctrl+Right : Previous / Next in playlist\n"
        L"  Up / Down            : Volume +/-5 (0..200)\n"
        L"  Del                  : Remove current & delete on exit\n"
        L"  Ctrl+R               : Pause -> Save As (rename queued until exit)\n"
        L"  Ctrl+C               : Pause -> Save As (copy queued until exit; shown in title during copy)\n"
        L"  Ctrl+G               : Pause -> Playlist chooser (jump with arrows)\n";

    if (g_cfg.ffprobeAvailable) {
        msg += L"  Ctrl+P               : Show video properties (ffprobe + shell properties)\n";
    }
    else {
        msg += L"  Ctrl+P               : Show basic video properties (shell); ffprobe disabled in config\n";
    }

    // Video tools menu (Ctrl+V)
    if (!g_cfg.upscaleDirectory.empty() || g_cfg.ffmpegAvailable) {
        msg += L"  Ctrl+V               : Video tools menu:\n";
        if (!g_cfg.upscaleDirectory.empty()) {
            msg += L"                           Submit for upscaling (copy to upscaleDirectory after playback)\n";
        }
        if (g_cfg.ffmpegAvailable) {
            msg += L"                           Trim front to current time (FFmpeg)\n"
                L"                           Trim end at current time (FFmpeg)\n"
                L"                           Horizontal flip (FFmpeg)\n";
        }
        msg += L"\n"
            L"  At end of playback, if FFmpeg tasks are still running,\n"
            L"  the title bar shows \"waiting on N task(s)\" until they all complete.\n";
    }

    MessageBoxW(g_hwndMain, msg.c_str(), L"Media Explorer - Help", MB_OK);

    // Resume only if it was actually playing before F1 was pressed
    if (g_mp && wasPlaying) {
        libvlc_media_player_set_pause(g_mp, 0);
    }
}

// ----------------------------- ListView helpers
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
    it.mask = LVIF_TEXT | LVIF_PARAM; it.iItem = rowIndex;
    it.pszText = const_cast<wchar_t*>(r.name.c_str()); it.lParam = rowIndex;
    ListView_InsertItem(g_hwndList, &it);

    const wchar_t* type = r.isDir ? L"Folder" : L"Video";
    ListView_SetItemText(g_hwndList, rowIndex, 1, const_cast<wchar_t*>(type));

    if (!r.isDir) {
        std::wstring s = FormatSize(r.size);
        ListView_SetItemText(g_hwndList, rowIndex, 2, const_cast<wchar_t*>(s.c_str()));
    }
    if (r.modified.dwLowDateTime || r.modified.dwHighDateTime) {
        std::wstring m = FormatFileTime(r.modified);
        ListView_SetItemText(g_hwndList, rowIndex, 3, const_cast<wchar_t*>(m.c_str()));
    }
    if (!r.isDir && (r.vW > 0 || r.vH > 0)) {
        wchar_t buf[64]; swprintf_s(buf, L"%dx%d", r.vW, r.vH);
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

// ----------------------------- Sorting (dirs first)
static void SortRows(int col, bool asc) {
    g_sortCol = col; g_sortAsc = asc;
    std::sort(g_rows.begin(), g_rows.end(),
        [col, asc](const Row& A, const Row& B) {
            if (A.isDir != B.isDir) return A.isDir && !B.isDir; // dirs first
            switch (col) {
            case 0: return asc ? (_wcsicmp(A.name.c_str(), B.name.c_str()) < 0) : (_wcsicmp(A.name.c_str(), B.name.c_str()) > 0);
            case 1: {
                int ta = A.isDir ? 0 : 1, tb = B.isDir ? 0 : 1;
                if (ta != tb) return ta < tb;
                return asc ? (_wcsicmp(A.name.c_str(), B.name.c_str()) < 0) : (_wcsicmp(A.name.c_str(), B.name.c_str()) > 0);
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
        GetVideoProps(path, w, h, d); // heavy; OK in worker
        MetaResult* r = new MetaResult{ path, w, h, d, myGen };
        PostMessageW(g_hwndMain, WM_APP_META, 0, (LPARAM)r);
    }

    CoUninitialize();
    return 0;
}
static void StartMetaWorker() {
    if (g_metaThread) { CloseHandle(g_metaThread); g_metaThread = NULL; } // let old thread exit on gen bump
    g_metaThread = CreateThread(NULL, 0, MetaThreadProc, NULL, 0, NULL);
}
static void CancelMetaWorkAndClearTodo() {
    g_metaGen.fetch_add(1, std::memory_order_relaxed);
    EnterCriticalSection(&g_metaLock);
    g_metaTodoPaths.clear();
    LeaveCriticalSection(&g_metaLock);
}

// After (re)building g_rows + ListView, queue any videos still missing props:
static void QueueMissingPropsAndKickWorker() {
    EnterCriticalSection(&g_metaLock);
    for (const auto& r : g_rows) {
        if (!r.isDir && r.vW == 0 && r.vH == 0 && r.vDur100ns == 0) {
            g_metaTodoPaths.push_back(r.full);
        }
    }
    LeaveCriticalSection(&g_metaLock);
    if (!g_metaTodoPaths.empty()) StartMetaWorker();
}

// ----------------------------- Populate views
static void ShowDrives() {
    CancelMetaWorkAndClearTodo();

    g_view = ViewKind::Drives; g_folder.clear(); g_rows.clear();

    SendMessageW(g_hwndList, WM_SETREDRAW, FALSE, 0);
    LV_ResetColumns();

    DWORD mask = GetLogicalDrives();
    for (int i = 0; i < 26; ++i) {
        if (!(mask & (1u << i))) continue;
        wchar_t root[4] = { wchar_t(L'A' + i), L':', L'\\', 0 };
        Row r; r.name = root; r.full = root; r.isDir = true;
        g_rows.push_back(r);
    }
    SortRows(0, true);
    SendMessageW(g_hwndList, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(g_hwndList, NULL, TRUE);

    SetTitleFolderOrDrives();
}

static void ShowFolder(std::wstring abs) {
    CancelMetaWorkAndClearTodo();

    if (abs.size() == 2 && abs[1] == L':') abs += L'\\';
    abs = EnsureSlash(abs);
    g_view = ViewKind::Folder; g_folder = abs; g_rows.clear();

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
        SetTitleFolderOrDrives();
        return;
    }

    std::vector<Row> dirs, vids;
    do {
        if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;

        Row r; r.name = fd.cFileName; r.full = abs + fd.cFileName;
        r.isDir = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        r.modified = fd.ftLastWriteTime;

        if (r.isDir) {
            r.full += L'\\';
            dirs.push_back(r);
        }
        else if (IsVideoFile(r.full)) {
            ULARGE_INTEGER uli; uli.HighPart = fd.nFileSizeHigh; uli.LowPart = fd.nFileSizeLow;
            r.size = uli.QuadPart;

            // FAST cached try first (cheap)
            if (!GetVideoPropsFastCached(r.full, r.vW, r.vH, r.vDur100ns)) {
                r.vW = r.vH = 0; r.vDur100ns = 0; // mark for async
            }
            vids.push_back(r);
        }
    } while (FindNextFileW(h, &fd));
    FindClose(h);

    g_rows.reserve(dirs.size() + vids.size());
    g_rows.insert(g_rows.end(), dirs.begin(), dirs.end());
    g_rows.insert(g_rows.end(), vids.begin(), vids.end());

    SortRows(g_sortCol, g_sortAsc);

    SendMessageW(g_hwndList, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(g_hwndList, NULL, TRUE);

    // Queue remaining metadata and kick worker
    QueueMissingPropsAndKickWorker();

    SetTitleFolderOrDrives();
}

// ----------------------------- Search (recursive, case-insensitive, AND terms)
static bool NameContainsAllTerms(const std::wstring& full, const std::vector<std::wstring>& termsLower) {
    const wchar_t* base = wcsrchr(full.c_str(), L'\\'); base = base ? base + 1 : full.c_str();
    std::wstring bl = ToLower(base);
    for (size_t i = 0; i < termsLower.size(); ++i) if (bl.find(termsLower[i]) == std::wstring::npos) return false;
    return true;
}
static void SearchRecurseFolder(const std::wstring& folder,
    const std::vector<std::wstring>& terms,
    std::vector<Row>& out) {
    SetTitleSearchingFolder(folder);

    std::wstring pat = EnsureSlash(folder) + L"*";
    WIN32_FIND_DATAW fd; ZeroMemory(&fd, sizeof(fd));
    HANDLE h = FindFirstFileExW(pat.c_str(), FindExInfoBasic, &fd, FindExSearchNameMatch, NULL, FIND_FIRST_EX_LARGE_FETCH);
    if (h == INVALID_HANDLE_VALUE) return;

    do {
        if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;

        bool isDir = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        std::wstring full = EnsureSlash(folder) + fd.cFileName;

        if (isDir) {
            if (fd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) continue; // avoid loops
            SearchRecurseFolder(full, terms, out);
        }
        else if (IsVideoFile(full)) {
            if (NameContainsAllTerms(full, terms)) {
                Row r;
                r.name = full;          // display full path in Search view
                r.full = full;
                r.isDir = false;
                r.modified = fd.ftLastWriteTime;
                ULARGE_INTEGER uli; uli.HighPart = fd.nFileSizeHigh; uli.LowPart = fd.nFileSizeLow;
                r.size = uli.QuadPart;

                // FAST cached only here; deep props deferred to worker
                GetVideoPropsFastCached(r.full, r.vW, r.vH, r.vDur100ns);

                out.push_back(r);
            }
        }
    } while (FindNextFileW(h, &fd));

    FindClose(h);
}
static void RunSearchFromOrigin(std::vector<Row>& outResults) {
    outResults.clear();

    // If user explicitly selected scope (files/folders/drives), honor that.
    if (g_search.useExplicitScope) {
        // 1) Selected files: test each one directly.
        for (const auto& file : g_search.explicitFiles) {
            if (!IsVideoFile(file)) continue; // only index video files
            if (!NameContainsAllTerms(file, g_search.termsLower)) continue;

            WIN32_FILE_ATTRIBUTE_DATA fad{};
            if (GetFileAttributesExW(file.c_str(), GetFileExInfoStandard, &fad) &&
                (fad.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) {

                Row r;
                r.name = file;              // Show full path in Search view
                r.full = file;
                r.isDir = false;
                r.modified = fad.ftLastWriteTime;

                ULARGE_INTEGER uli{};
                uli.HighPart = fad.nFileSizeHigh;
                uli.LowPart = fad.nFileSizeLow;
                r.size = uli.QuadPart;

                GetVideoPropsFastCached(r.full, r.vW, r.vH, r.vDur100ns); // cheap, deep fill is async later
                outResults.push_back(std::move(r));
            }
        }

        // 2) Selected folders/drives: recurse each.
        for (const auto& folder : g_search.explicitFolders) {
            SetTitleSearchingFolder(folder);
            SearchRecurseFolder(folder, g_search.termsLower, outResults);
        }
        return;
    }

    // Original behavior (no explicit selection): search from origin
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
    g_rows = results; // copy

    SendMessageW(g_hwndList, WM_SETREDRAW, FALSE, 0);
    LV_ResetColumns();
    SortRows(g_sortCol, g_sortAsc);
    SendMessageW(g_hwndList, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(g_hwndList, NULL, TRUE);

    // Title includes terms; also show count
    std::wstring t = L"Media Explorer (libVLC) - Search - " + JoinTermsForTitle();
    wchar_t buf[64]; swprintf_s(buf, L" - %zu file(s)", g_rows.size());
    t += buf;
    SetWindowTextW(g_hwndMain, t.c_str());

    // Queue remaining metadata and kick worker
    QueueMissingPropsAndKickWorker();
}
static void ExitSearchToOrigin() {
    if (!g_search.active) return;
    if (g_search.originView == ViewKind::Drives) ShowDrives();
    else ShowFolder(g_search.originFolder);
    g_search = SearchState(); // reset
}

// ----------------------------- File operations (browser)
static void Browser_CopySelectedToClipboard(ClipMode mode) {
    g_clipFiles.clear();
    g_clipMode = ClipMode::None;
    if (g_view == ViewKind::Drives) return;

    // Collect selected files
    std::vector<int> selectedFileIdx;
    int idx = -1; bool any = false;
    while ((idx = ListView_GetNextItem(g_hwndList, idx, LVNI_SELECTED)) != -1) {
        if (idx < 0 || idx >= (int)g_rows.size()) continue;
        const Row& r = g_rows[idx];
        if (!r.isDir) {
            g_clipFiles.push_back(r.full);
            selectedFileIdx.push_back(idx);
            any = true;
        }
    }
    if (!any) return;

    g_clipMode = mode;

    // If this is a "cut" (Ctrl+X), remove items from current view immediately.
    if (mode == ClipMode::Move) {
        SendMessageW(g_hwndList, WM_SETREDRAW, FALSE, 0);

        // erase from both the vector and the UI, highest index first
        std::sort(selectedFileIdx.begin(), selectedFileIdx.end());
        for (int i = (int)selectedFileIdx.size() - 1; i >= 0; --i) {
            int rIdx = selectedFileIdx[i];
            if (rIdx >= 0 && rIdx < (int)g_rows.size() && !g_rows[rIdx].isDir) {
                g_rows.erase(g_rows.begin() + rIdx);
                ListView_DeleteItem(g_hwndList, rIdx);
            }
        }

        SendMessageW(g_hwndList, WM_SETREDRAW, TRUE, 0);
        InvalidateRect(g_hwndList, NULL, TRUE);
    }
}

static std::wstring UniqueName(const std::wstring& folder, const std::wstring& base, const std::wstring& ext) {
    std::wstring target = folder + base + ext;
    if (!PathFileExistsW(target.c_str())) return target;
    for (int i = 1; i < 10000; ++i) {
        wchar_t buf[32]; swprintf_s(buf, L" (%d)", i);
        std::wstring t = folder + base + buf + ext;
        if (!PathFileExistsW(t.c_str())) return t;
    }
    return target;
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
    BOOL* pCancelFlag{ nullptr };  // points to local BOOL in paste routine (for CopyFileEx)
} g_op;

static LRESULT CALLBACK OpProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    switch (m) {
    case WM_CREATE: {
        HFONT hf = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        RECT rc{}; GetClientRect(h, &rc);
        int margin = DpiScale(12);
        int btnW = DpiScale(100), btnH = DpiScale(28);

        g_op.hText = CreateWindowExW(WS_EX_TRANSPARENT, L"STATIC", L"",
            WS_CHILD | WS_VISIBLE | SS_LEFT, margin, margin, rc.right - 2 * margin, DpiScale(32),
            h, (HMENU)101, g_hInst, NULL);
        SendMessageW(g_op.hText, WM_SETFONT, (WPARAM)hf, TRUE);

        g_op.hCancel = CreateWindowExW(0, L"BUTTON", L"Cancel",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            rc.right - margin - btnW, rc.bottom - margin - btnH, btnW, btnH,
            h, (HMENU)IDCANCEL, g_hInst, NULL);
        SendMessageW(g_op.hCancel, WM_SETFONT, (WPARAM)hf, TRUE);
        return 0;
    }
    case WM_SIZE: {
        RECT rc{}; GetClientRect(h, &rc);
        int margin = DpiScale(12);
        int btnW = DpiScale(100), btnH = DpiScale(28);
        if (g_op.hText)   MoveWindow(g_op.hText, margin, margin, rc.right - 2 * margin, DpiScale(32), TRUE);
        if (g_op.hCancel) MoveWindow(g_op.hCancel, rc.right - margin - btnW, rc.bottom - margin - btnH, btnW, btnH, TRUE);
        return 0;
    }
    case WM_COMMAND:
        if (LOWORD(w) == IDCANCEL) {
            if (g_op.pCancelFlag) *g_op.pCancelFlag = TRUE;
            g_op.cancel.store(true, std::memory_order_relaxed);
            DestroyWindow(h); // Drop window immediately
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
        WNDCLASSW wc{}; wc.lpfnWndProc = OpProc; wc.hInstance = g_hInst;
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

    HWND h = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST, L"OpProgressClass", title,
        WS_POPUPWINDOW | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        X, Y, W, H, g_hwndMain, NULL, g_hInst, NULL);
    g_op.hwnd = h;
    return h;
}
static DWORD CALLBACK CopyProgressThunk(LARGE_INTEGER, LARGE_INTEGER, LARGE_INTEGER, LARGE_INTEGER,
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

// Run the copy/move with UI and cancellation
static void RunClipboardOperationWithUI(const std::wstring& dstFolder) {
    if (g_clipMode == ClipMode::None || g_clipFiles.empty()) return;

    const bool   isCopy = (g_clipMode == ClipMode::Copy);
    const size_t total = g_clipFiles.size();

    auto makeCaption = [&](size_t currentIndex) -> std::wstring {
        if (total <= 1) return std::wstring(isCopy ? L"Copying..." : L"Moving...");
        wchar_t buf[256];
        swprintf_s(buf, L"%s... %zu file of %zu files",
            isCopy ? L"Copying" : L"Moving",
            currentIndex, total);
        return buf;
        };

    std::wstring initialCap = makeCaption(total > 1 ? 1 : 0);
    HWND hw = CreateOpWindow(initialCap.c_str());
    if (!hw) return;

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

    for (size_t i = 0; i < total; ++i) {
        if (g_op.cancel.load(std::memory_order_relaxed)) break;

        const std::wstring& src = g_clipFiles[i];
        const wchar_t* base = wcsrchr(src.c_str(), L'\\'); base = base ? base + 1 : src.c_str();

        wchar_t fname[_MAX_FNAME] = {}, ext[_MAX_EXT] = {};
        _wsplitpath_s(base, NULL, 0, NULL, 0, fname, _MAX_FNAME, ext, _MAX_EXT);

        std::wstring dst = UniqueName(dstFolder, fname, ext);

        if (total > 1) updateCaption(i + 1);

        std::wstring line = (isCopy ? L"Copying " : L"Moving ");
        line += base;
        line += L"...";
        setStatusText(line);
        PumpMessagesThrottled(10);

        BOOL ok = FALSE;
        uiCancel = FALSE;

        if (isCopy) {
            ok = CopyFileExW(src.c_str(), dst.c_str(),
                CopyProgressThunk, &g_op, &uiCancel, 0);
        }
        else {
            if (SameVolume(src, dst)) {
                ok = MoveFileExW(src.c_str(), dst.c_str(), MOVEFILE_REPLACE_EXISTING);
            }
            else {
                ok = CopyFileExW(src.c_str(), dst.c_str(),
                    CopyProgressThunk, &g_op, &uiCancel, 0);
                if (ok) {
                    if (!DeleteFileW(src.c_str())) {
                        MoveFileExW(src.c_str(), NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
                    }
                }
                else {
                    DeleteFileW(dst.c_str());
                }
            }
        }

        if (g_op.cancel.load(std::memory_order_relaxed)) break;

        setStatusText(line + L" Done");
        PumpMessagesThrottled(10);

        if (total > 1 && (i + 1) < total) updateCaption(i + 2);
    }

    if (IsWindow(g_op.hwnd)) DestroyWindow(g_op.hwnd);

    g_clipFiles.clear();
    g_clipMode = ClipMode::None;

    if (g_view == ViewKind::Folder) ShowFolder(g_folder);
}

static void Browser_PasteClipboardIntoCurrent() {
    if ((g_view != ViewKind::Folder && g_view != ViewKind::Search) || g_clipMode == ClipMode::None || g_clipFiles.empty()) return;

    std::wstring dstFolder;
    if (g_view == ViewKind::Folder) dstFolder = g_folder;
    else if (g_view == ViewKind::Search && g_search.originView == ViewKind::Folder) dstFolder = g_search.originFolder;
    else return;

    RunClipboardOperationWithUI(dstFolder);
}

static void Browser_DeleteSelected() {
    if (g_view == ViewKind::Drives) return;
    if (MessageBoxW(g_hwndMain, L"Delete selected files permanently?", L"Confirm Delete",
        MB_YESNO | MB_DEFBUTTON2) != IDYES) return;

    int idx = -1;
    while ((idx = ListView_GetNextItem(g_hwndList, idx, LVNI_SELECTED)) != -1) {
        if (idx < 0 || idx >= (int)g_rows.size()) continue;
        const Row& r = g_rows[idx];
        if (!r.isDir && PathFileExistsW(r.full.c_str())) {
            if (!DeleteFileW(r.full.c_str())) MoveFileExW(r.full.c_str(), NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
        }
    }
    if (g_view == ViewKind::Folder) ShowFolder(g_folder);
}

// Move a single selected row up or down in the current list.
// direction = -1 (up) or +1 (down). Only works when exactly one row is selected.
static void Browser_MoveSelectedRow(int direction) {
    if (g_view == ViewKind::Drives) return;
    if (direction != -1 && direction != +1) return;
    if (g_rows.empty()) return;

    int sel = ListView_GetNextItem(g_hwndList, -1, LVNI_SELECTED);
    if (sel < 0) return;
    if (ListView_GetNextItem(g_hwndList, sel, LVNI_SELECTED) != -1) {
        // More than one selected – ambiguous order, ignore
        return;
    }

    int target = sel + direction;
    if (target < 0 || target >= (int)g_rows.size()) return;

    std::swap(g_rows[sel], g_rows[target]);
    LV_Rebuild();

    ListView_SetItemState(g_hwndList, target,
        LVIS_SELECTED | LVIS_FOCUSED,
        LVIS_SELECTED | LVIS_FOCUSED);
    ListView_EnsureVisible(g_hwndList, target, FALSE);
}

// Prompt for combined output filename, forcing it into the current folder (g_folder).
static bool PromptCombinedOutputName(const std::wstring& defaultName, std::wstring& outFullPath) {
    if (g_view != ViewKind::Folder) return false;
    if (g_folder.empty()) return false;

    ComPtr<IFileSaveDialog> dlg;
    if (FAILED(CoCreateInstance(CLSID_FileSaveDialog, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dlg))))
        return false;

    ComPtr<IShellItem> initFolder;
    if (SUCCEEDED(SHCreateItemFromParsingName(g_folder.c_str(), NULL, IID_PPV_ARGS(&initFolder)))) {
        dlg->SetFolder(initFolder.Get());
    }

    dlg->SetFileName(defaultName.c_str());

    COMDLG_FILTERSPEC spec[] = {
        { L"Video Files", L"*.mp4;*.mkv;*.mov;*.avi;*.wmv;*.ts;*.m2ts;*.webm;*.flv;*.m4v" },
        { L"All Files",   L"*.*" }
    };
    dlg->SetFileTypes(2, spec);
    dlg->SetTitle(L"Combined video filename");
    dlg->SetOptions(FOS_OVERWRITEPROMPT | FOS_FORCEFILESYSTEM);

    if (FAILED(dlg->Show(g_hwndMain))) return false;

    ComPtr<IShellItem> it;
    if (FAILED(dlg->GetResult(&it))) return false;

    PWSTR psz = NULL;
    if (FAILED(it->GetDisplayName(SIGDN_FILESYSPATH, &psz))) return false;
    std::wstring chosen(psz);
    CoTaskMemFree(psz);

    const wchar_t* base = wcsrchr(chosen.c_str(), L'\\');
    base = base ? base + 1 : chosen.c_str();

    outFullPath = EnsureSlash(g_folder);
    outFullPath += base;
    return true;
}

// Combine selected files using external video_combine.exe in background thread.
static void Browser_CombineSelected() {
    if (!g_cfg.ffmpegAvailable) {
        MessageBoxW(g_hwndMain,
            L"video_combine is disabled in mediaexplorer.ini.\n"
            L"Set ffmpegAvailable = 1 to enable combining videos.",
            L"Combine videos", MB_OK);
        return;
    }
    if (g_view != ViewKind::Folder) return;
    if (g_rows.empty()) return;

    std::vector<int> selIdx;
    int idx = -1;
    while ((idx = ListView_GetNextItem(g_hwndList, idx, LVNI_SELECTED)) != -1) {
        if (idx < 0 || idx >= (int)g_rows.size()) continue;
        if (g_rows[idx].isDir) continue;
        selIdx.push_back(idx);
    }
    if (selIdx.size() <= 1) {
        // Ignore if not more than one file selected
        return;
    }
    std::sort(selIdx.begin(), selIdx.end());

    std::vector<std::wstring> srcFiles;
    srcFiles.reserve(selIdx.size());
    for (int i : selIdx) {
        srcFiles.push_back(g_rows[i].full);
    }

    const std::wstring& firstFull = srcFiles.front();
    const wchar_t* baseFirst = wcsrchr(firstFull.c_str(), L'\\');
    baseFirst = baseFirst ? baseFirst + 1 : firstFull.c_str();

    wchar_t fname[_MAX_FNAME] = {}, ext[_MAX_EXT] = {};
    _wsplitpath_s(baseFirst, NULL, 0, NULL, 0, fname, _MAX_FNAME, ext, _MAX_EXT);

    std::wstring stem = fname;
    std::wstring extension = (*ext ? ext : L".mp4");
    std::wstring defaultName = stem + extension;

    std::wstring combinedFull;
    if (!PromptCombinedOutputName(defaultName, combinedFull)) return;

    std::wstring folderWithSlash = EnsureSlash(g_folder);
    const wchar_t* baseOut = wcsrchr(combinedFull.c_str(), L'\\');
    std::wstring baseOutName = baseOut ? baseOut + 1 : combinedFull;
    size_t dotPos = baseOutName.find_last_of(L'.');
    std::wstring outStem = (dotPos == std::wstring::npos) ? baseOutName : baseOutName.substr(0, dotPos);

    std::wstring copyDir = folderWithSlash + outStem + L"\\";

    CombineTask* task = new CombineTask();
    task->workingDir = copyDir;
    task->srcFiles = srcFiles;
    task->combinedFull = combinedFull;
    task->title = baseOutName;
    task->running = true;

    EnsureCombineLogClass();
    HWND logWnd = CreateCombineLogWindow(task);
    if (!logWnd) {
        delete task;
        MessageBoxW(g_hwndMain, L"Failed to create log window for video combine.", L"Combine videos",
            MB_OK);
        return;
    }
    task->hwnd = logWnd;

    EnterCriticalSection(&g_combineLock);
    g_combineTasks.push_back(task);
    LeaveCriticalSection(&g_combineLock);

    HANDLE hThread = CreateThread(NULL, 0, CombineThreadProc, task, 0, NULL);
    if (!hThread) {
        EnterCriticalSection(&g_combineLock);
        auto it = std::find(g_combineTasks.begin(), g_combineTasks.end(), task);
        if (it != g_combineTasks.end()) g_combineTasks.erase(it);
        LeaveCriticalSection(&g_combineLock);

        if (IsWindow(task->hwnd)) DestroyWindow(task->hwnd);
        delete task;

        MessageBoxW(g_hwndMain, L"Failed to start background thread for video combine.",
            L"Combine videos", MB_OK);
        return;
    }

    task->hThread = hThread;

    std::wstring startMsg = L"Starting combine for ";
    wchar_t buf[64]; swprintf_s(buf, L"%zu", task->srcFiles.size());
    startMsg += buf;
    startMsg += L" file(s)...\r\nWorking directory: ";
    startMsg += task->workingDir;
    startMsg += L"\r\n";
    PostCombineOutput(task, startMsg);
}

// ----------------------------- Dialog helpers (playback)
static bool PromptSaveAsFrom(const std::wstring& seedPath, std::wstring& outPath, const wchar_t* titleText) {
    ComPtr<IFileSaveDialog> dlg;
    if (FAILED(CoCreateInstance(CLSID_FileSaveDialog, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dlg))))
        return false;

    wchar_t dir[MAX_PATH] = {}; wcscpy_s(dir, seedPath.c_str()); PathRemoveFileSpecW(dir);
    ComPtr<IShellItem> initFolder;
    if (SUCCEEDED(SHCreateItemFromParsingName(dir, NULL, IID_PPV_ARGS(&initFolder))))
        dlg->SetFolder(initFolder.Get());

    const wchar_t* base = wcsrchr(seedPath.c_str(), L'\\'); base = base ? base + 1 : seedPath.c_str();
    dlg->SetFileName(base);

    COMDLG_FILTERSPEC spec[] = { { L"All Files", L"*.*" } };
    dlg->SetFileTypes(1, spec);
    dlg->SetTitle(titleText ? titleText : L"Save As");
    dlg->SetOptions(FOS_OVERWRITEPROMPT | FOS_FORCEFILESYSTEM);

    if (FAILED(dlg->Show(g_hwndMain))) return false;

    ComPtr<IShellItem> it; if (FAILED(dlg->GetResult(&it))) return false;
    PWSTR psz = NULL; if (FAILED(it->GetDisplayName(SIGDN_FILESYSPATH, &psz))) return false;
    outPath.assign(psz); CoTaskMemFree(psz); return true;
}

// ----------------------------- Keyword prompt (Ctrl+F)
struct KwCtx { HWND hwnd, hEdit, hOK, hCancel; bool accepted; std::wstring text; };
static KwCtx g_kw = { 0,0,0,0,false,L"" };

static LRESULT CALLBACK KwEditSub(HWND h, UINT m, WPARAM w, LPARAM l, UINT_PTR, DWORD_PTR) {
    if (m == WM_KEYDOWN) {
        if (w == VK_RETURN) { PostMessageW(GetParent(h), WM_COMMAND, MAKELONG(IDOK, BN_CLICKED), (LPARAM)g_kw.hOK); return 0; }
        if (w == VK_ESCAPE) { PostMessageW(GetParent(h), WM_COMMAND, MAKELONG(IDCANCEL, BN_CLICKED), (LPARAM)g_kw.hCancel); return 0; }
    }
    return DefSubclassProc(h, m, w, l);
}

// ----------------------------- Video tools menu (Ctrl+V during playback)

struct VideoToolsCtx {
    HWND hwnd = NULL;
    HWND btn1 = NULL; // upscale
    HWND btn2 = NULL; // trim front
    HWND btn3 = NULL; // trim end
    HWND btn4 = NULL; // hflip
    bool accepted = false;
    int  choice = 0;  // 1..4
    bool canUpscale = false;
    bool canFfmpeg = false;
};

static VideoToolsCtx g_vtools;

static LRESULT CALLBACK VideoToolsProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    switch (m) {
    case WM_CREATE: {
        HFONT hf = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        RECT rc; GetClientRect(h, &rc);
        int margin = DpiScale(12);
        int btnW = DpiScale(260);
        int btnH = DpiScale(28);
        int y = margin;

        g_vtools.hwnd = h;

        HWND hLbl = CreateWindowExW(0, L"STATIC", L"Video tools (Ctrl+V):",
            WS_CHILD | WS_VISIBLE, margin, y, rc.right - 2 * margin, DpiScale(20),
            h, NULL, g_hInst, NULL);
        SendMessageW(hLbl, WM_SETFONT, (WPARAM)hf, TRUE);
        y += DpiScale(28);

        if (g_vtools.canUpscale) {
            g_vtools.btn1 = CreateWindowExW(0, L"BUTTON", L"Submit for upscaling",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_LEFT,
                margin, y, btnW, btnH, h, (HMENU)4001, g_hInst, NULL);
            SendMessageW(g_vtools.btn1, WM_SETFONT, (WPARAM)hf, TRUE);
            y += btnH + DpiScale(6);
        }
        if (g_vtools.canFfmpeg) {
            g_vtools.btn2 = CreateWindowExW(0, L"BUTTON", L"Trim front to current time (ffmpeg)",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_LEFT,
                margin, y, btnW, btnH, h, (HMENU)4002, g_hInst, NULL);
            SendMessageW(g_vtools.btn2, WM_SETFONT, (WPARAM)hf, TRUE);
            y += btnH + DpiScale(6);

            g_vtools.btn3 = CreateWindowExW(0, L"BUTTON", L"Trim end at current time (ffmpeg)",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_LEFT,
                margin, y, btnW, btnH, h, (HMENU)4003, g_hInst, NULL);
            SendMessageW(g_vtools.btn3, WM_SETFONT, (WPARAM)hf, TRUE);
            y += btnH + DpiScale(6);

            g_vtools.btn4 = CreateWindowExW(0, L"BUTTON", L"Horizontal flip (ffmpeg)",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_LEFT,
                margin, y, btnW, btnH, h, (HMENU)4004, g_hInst, NULL);
            SendMessageW(g_vtools.btn4, WM_SETFONT, (WPARAM)hf, TRUE);
            y += btnH + DpiScale(6);
        }

        return 0;
    }
    case WM_COMMAND: {
        int id = LOWORD(w);
        if (id >= 4001 && id <= 4004) {
            g_vtools.accepted = true;
            if (id == 4001) g_vtools.choice = 1;
            else if (id == 4002) g_vtools.choice = 2;
            else if (id == 4003) g_vtools.choice = 3;
            else if (id == 4004) g_vtools.choice = 4;
            DestroyWindow(h);
            return 0;
        }
        break;
    }
    case WM_KEYDOWN:
        if (w == VK_ESCAPE) {
            g_vtools.accepted = false;
            DestroyWindow(h);
            return 0;
        }
        if (w == '1' && g_vtools.canUpscale) {
            SendMessageW(h, WM_COMMAND, 4001, 0);
            return 0;
        }
        if (w == '2' && g_vtools.canFfmpeg) {
            SendMessageW(h, WM_COMMAND, 4002, 0);
            return 0;
        }
        if (w == '3' && g_vtools.canFfmpeg) {
            SendMessageW(h, WM_COMMAND, 4003, 0);
            return 0;
        }
        if (w == '4' && g_vtools.canFfmpeg) {
            SendMessageW(h, WM_COMMAND, 4004, 0);
            return 0;
        }
        break;
    case WM_CLOSE:
        g_vtools.accepted = false;
        DestroyWindow(h);
        return 0;
    case WM_DESTROY:
        g_vtools.hwnd = NULL;
        return 0;
    }
    return DefWindowProcW(h, m, w, l);
}

static int PromptVideoToolsChoice(bool canUpscale, bool canFfmpeg) {
    if (!canUpscale && !canFfmpeg) {
        MessageBoxW(g_hwndMain,
            L"No video tools are available.\n\n"
            L"- Configure upscaleDirectory and/or\n"
            L"- Set ffmpegAvailable=1 in mediaexplorer.ini.",
            L"Video tools", MB_OK);
        return 0;
    }

    static bool reg = false;
    if (!reg) {
        WNDCLASSW wc{}; wc.lpfnWndProc = VideoToolsProc; wc.hInstance = g_hInst;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"VideoToolsClass";
        RegisterClassW(&wc);
        reg = true;
    }

    g_vtools = VideoToolsCtx();
    g_vtools.canUpscale = canUpscale;
    g_vtools.canFfmpeg = canFfmpeg;

    MONITORINFO mi{ sizeof(mi) };
    HMONITOR hm = MonitorFromWindow(g_hwndMain, MONITOR_DEFAULTTONEAREST);
    GetMonitorInfoW(hm, &mi);
    RECT wa = mi.rcWork;

    int W = DpiScale(420), H = DpiScale(220);
    int X = wa.left + ((wa.right - wa.left) - W) / 2;
    int Y = wa.top + ((wa.bottom - wa.top) - H) / 2;

    HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST, L"VideoToolsClass", L"Video tools",
        WS_POPUPWINDOW | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        X, Y, W, H, g_hwndMain, NULL, g_hInst, NULL);

    SetWindowPos(hwnd, HWND_TOPMOST, X, Y, W, H, SWP_SHOWWINDOW);
    SetForegroundWindow(hwnd);

    MSG msg;
    while (IsWindow(hwnd) && GetMessageW(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    if (g_vtools.accepted) return g_vtools.choice;
    SetForegroundWindow(g_hwndMain);
    return 0;
}


static LRESULT CALLBACK KwProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    switch (m) {
    case WM_CREATE: {
        HFONT hf = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        RECT rc; GetClientRect(h, &rc);
        int margin = DpiScale(12);
        int btnW = DpiScale(90), btnH = DpiScale(28);
        int labelH = DpiScale(20);
        int editH = DpiScale(24);

        HWND hLbl = CreateWindowExW(0, L"STATIC", L"Search keyword (case-insensitive):",
            WS_CHILD | WS_VISIBLE, margin, margin, rc.right - 2 * margin, labelH, h, NULL, g_hInst, NULL);
        SendMessageW(hLbl, WM_SETFONT, (WPARAM)hf, TRUE);

        g_kw.hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL, margin, margin + labelH + DpiScale(6),
            rc.right - 2 * margin - (btnW + DpiScale(10)), editH, h, (HMENU)201, g_hInst, NULL);
        SendMessageW(g_kw.hEdit, WM_SETFONT, (WPARAM)hf, TRUE);
        SetWindowSubclass(g_kw.hEdit, KwEditSub, 11, 0);

        int btnY = rc.bottom - margin - btnH;
        g_kw.hOK = CreateWindowExW(0, L"BUTTON", L"OK",
            WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
            rc.right - margin - btnW - (btnW + DpiScale(10)), btnY, btnW, btnH, h, (HMENU)IDOK, g_hInst, NULL);
        SendMessageW(g_kw.hOK, WM_SETFONT, (WPARAM)hf, TRUE);

        g_kw.hCancel = CreateWindowExW(0, L"BUTTON", L"Cancel",
            WS_CHILD | WS_VISIBLE,
            rc.right - margin - btnW, btnY, btnW, btnH, h, (HMENU)IDCANCEL, g_hInst, NULL);
        SendMessageW(g_kw.hCancel, WM_SETFONT, (WPARAM)hf, TRUE);

        SetFocus(g_kw.hEdit);
        return 0;
    }
    case WM_COMMAND:
        if (LOWORD(w) == IDOK) {
            {
                int len = GetWindowTextLengthW(g_kw.hEdit);
                std::wstring t(len, L'\0');
                GetWindowTextW(g_kw.hEdit, &t[0], len + 1);
                g_kw.text = t;
            }
            g_kw.accepted = !g_kw.text.empty();
            DestroyWindow(h);
            return 0;
        }
        if (LOWORD(w) == IDCANCEL) { g_kw.accepted = false; DestroyWindow(h); return 0; }
        break;
    case WM_CLOSE: g_kw.accepted = false; DestroyWindow(h); return 0;
    }
    return DefWindowProcW(h, m, w, l);
}
static bool PromptKeyword(std::wstring& out) {
    static bool reg = false;
    if (!reg) {
        WNDCLASSW wc; ZeroMemory(&wc, sizeof(wc));
        wc.lpfnWndProc = KwProc; wc.hInstance = g_hInst;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"KwPromptClass";
        wc.style = CS_DBLCLKS;
        RegisterClassW(&wc); reg = true;
    }
    g_kw = KwCtx(); g_kw.accepted = false;

    MONITORINFO mi; mi.cbSize = sizeof(mi);
    HMONITOR hm = MonitorFromWindow(g_hwndMain, MONITOR_DEFAULTTONEAREST);
    GetMonitorInfoW(hm, &mi);
    RECT wa = mi.rcWork;

    int W = DpiScale(600), H = DpiScale(160);
    int X = wa.left + ((wa.right - wa.left) - W) / 2;
    int Y = wa.top + ((wa.bottom - wa.top) - H) / 2;

    HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_TOPMOST, L"KwPromptClass", L"Search",
        WS_POPUPWINDOW | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        X, Y, W, H, g_hwndMain, NULL, g_hInst, NULL);

    SetWindowPos(hwnd, HWND_TOPMOST, X, Y, W, H, SWP_SHOWWINDOW);
    SetForegroundWindow(hwnd);

    MSG msg;
    while (IsWindow(hwnd) && GetMessageW(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    if (g_kw.accepted) { out = g_kw.text; return true; }
    SetForegroundWindow(g_hwndMain);
    return false;
}

// ----------------------------- Playback
static void ApplyPostActionsAndRefresh() {
    for (size_t i = 0; i < g_post.size(); ++i) {
        const PostAction& a = g_post[i];
        switch (a.type) {
        case ActionType::DeleteFile: {
            BOOL ok = DeleteFileW(a.src.c_str());
            if (!ok) {
                DWORD err = GetLastError();
                MoveFileExW(a.src.c_str(), NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
                LogLine(L"PostAction DeleteFile: src=\"%s\" FAILED err=%lu (queued delete on reboot)",
                    a.src.c_str(), err);
            }
            else {
                LogLine(L"PostAction DeleteFile: src=\"%s\" OK", a.src.c_str());
            }
            break;
        }
        case ActionType::RenameFile: {
            BOOL ok = MoveFileExW(a.src.c_str(), a.param.c_str(),
                MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING);
            DWORD err = ok ? 0 : GetLastError();
            LogLine(L"PostAction RenameFile: src=\"%s\" dst=\"%s\" %s err=%lu",
                a.src.c_str(), a.param.c_str(), ok ? L"OK" : L"FAILED", err);
            break;
        }
        case ActionType::CopyToPath: {
            std::wstring t = L"Media Explorer (libVLC) - copying file...";
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
    else if (g_view == ViewKind::Drives) ShowDrives();
    else ShowFolder(g_folder);
}

static void PlayIndex(size_t idx) {
    if (!g_vlc) {
        const char* args[] = { "--avcodec-hw=d3d11va", "--no-video-title-show" };
        g_vlc = libvlc_new((int)(sizeof(args) / sizeof(args[0])), args);
        g_mp = libvlc_media_player_new(g_vlc);
        libvlc_media_player_set_hwnd(g_mp, g_hwndVideo);
        libvlc_video_set_scale(g_mp, 0.f);       // auto fill
        libvlc_video_set_aspect_ratio(g_mp, NULL);

        libvlc_event_manager_t* em = libvlc_media_player_event_manager(g_mp);
        libvlc_event_attach(em, libvlc_MediaPlayerEndReached,
            [](const libvlc_event_t*, void*) { PostMessageW(g_hwndMain, WM_APP + 1, 0, 0); }, NULL);
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

static void FinalizeAllFfmpegTasks() {
    // Move processed files up a directory and then delete video_process dirs.
    std::vector<std::wstring> dirsToDelete;

    EnterCriticalSection(&g_ffLock);
    for (FfmpegTask* t : g_ffTasks) {
        if (!t) continue;

        // If success and we have a finalWorking file, move it up one directory.
        if (t->exitCode == 0 && !t->finalWorking.empty()) {
            std::wstring src = t->finalWorking;

            // Parent directory is one level up from workingDir
            std::wstring parent = t->workingDir;
            if (!parent.empty() && (parent.back() == L'\\' || parent.back() == L'/')) parent.pop_back();
            PathRemoveFileSpecW(&parent[0]);
            parent = parent.c_str();
            parent = EnsureSlash(parent);

            const wchar_t* base = wcsrchr(src.c_str(), L'\\');
            base = base ? base + 1 : src.c_str();

            wchar_t fname[_MAX_FNAME] = {}, ext[_MAX_EXT] = {};
            _wsplitpath_s(base, NULL, 0, NULL, 0, fname, _MAX_FNAME, ext, _MAX_EXT);

            std::wstring dst = UniqueName(parent, fname, ext);
            MoveFileExW(src.c_str(), dst.c_str(),
                MOVEFILE_COPY_ALLOWED | MOVEFILE_REPLACE_EXISTING);
        }

        // Collect workingDir for later deletion
        if (!t->workingDir.empty())
            dirsToDelete.push_back(t->workingDir);
    }
    LeaveCriticalSection(&g_ffLock);

    // Delete video_process directories (best-effort).
    std::sort(dirsToDelete.begin(), dirsToDelete.end());
    dirsToDelete.erase(std::unique(dirsToDelete.begin(), dirsToDelete.end()), dirsToDelete.end());

    for (const auto& dir : dirsToDelete) {
        std::wstring pattern = EnsureSlash(dir) + L"*";
        WIN32_FIND_DATAW fd{};
        HANDLE h = FindFirstFileW(pattern.c_str(), &fd);
        if (h != INVALID_HANDLE_VALUE) {
            do {
                if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
                std::wstring full = EnsureSlash(dir) + fd.cFileName;
                if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
                    // best effort, skip nested dirs for now
                    continue;
                }
                DeleteFileW(full.c_str());
            } while (FindNextFileW(h, &fd));
            FindClose(h);
        }
        RemoveDirectoryW(dir.c_str());
    }

    // Clean up task objects
    EnterCriticalSection(&g_ffLock);
    for (FfmpegTask* t : g_ffTasks) {
        if (!t) continue;
        if (t->hProcess) CloseHandle(t->hProcess);
        if (t->hThread) CloseHandle(t->hThread);
        if (t->hwnd && IsWindow(t->hwnd)) DestroyWindow(t->hwnd);
        delete t;
    }
    g_ffTasks.clear();
    LeaveCriticalSection(&g_ffLock);
}

static void WaitForFfmpegTasksAndFinalize() {
    // If nothing ever scheduled, nothing to do
    EnterCriticalSection(&g_ffLock);
    bool anyTasks = !g_ffTasks.empty();
    LeaveCriticalSection(&g_ffLock);
    if (!anyTasks) return;

    auto countRunning = []() -> int {
        int c = 0;
        EnterCriticalSection(&g_ffLock);
        for (FfmpegTask* t : g_ffTasks) {
            if (t && t->running) ++c;
        }
        LeaveCriticalSection(&g_ffLock);
        return c;
        };

    int remaining = countRunning();
    if (remaining > 0) {
        // Wait while updating title "waiting on N tasks"
        for (;;) {
            remaining = countRunning();
            if (remaining <= 0) break;

            wchar_t buf[256];
            swprintf_s(buf, L"Media Explorer (libVLC) - waiting on %d FFmpeg task(s)...", remaining);
            SetWindowTextW(g_hwndMain, buf);

            MSG msg;
            while (PeekMessageW(&msg, NULL, 0, 0, PM_REMOVE)) {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
            Sleep(50);
        }
    }

    // All tasks done; finalize (move outputs, delete video_process dirs, close windows)
    FinalizeAllFfmpegTasks();
}

static void ExitPlayback() {
    LogLine(L"ExitPlayback called: inPlayback=%d", g_inPlayback ? 1 : 0);
    if (!g_inPlayback) return;

    if (g_fullscreen) ToggleFullscreen();
    KillTimer(g_hwndMain, kTimerPlaybackUI);
    if (g_mp) libvlc_media_player_stop(g_mp);

    // Wait for any pending FFmpeg tasks and finalize them
    WaitForFfmpegTasksAndFinalize();

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

    if (g_view == ViewKind::Drives || r.isDir) {
        if (g_view == ViewKind::Search) return; // only files in Search
        ShowFolder(r.full);
    }
    else {
        PlaySelectedVideos();
    }
}

static void NavigateBack() {
    if (g_view == ViewKind::Search) { ExitSearchToOrigin(); return; }
    if (g_view == ViewKind::Drives) return;
    if (IsDriveRoot(g_folder)) { ShowDrives(); return; }
    std::wstring parent = ParentDir(g_folder);
    if (parent.empty()) ShowDrives(); else ShowFolder(parent);
}

// ----------------------------- Playlist chooser (Ctrl+G)
struct PickerCtx { HWND hwnd, hList; };
static PickerCtx g_pick = { 0,0 };

static LRESULT CALLBACK PickerProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    switch (m) {
    case WM_CREATE: {
        g_pick.hList = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
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
        if (HIWORD(w) == LBN_DBLCLK && (HWND)l == g_pick.hList) { DestroyWindow(h); return 0; }
        break;
    case WM_KEYDOWN:
        if (w == VK_RETURN || w == VK_ESCAPE) { DestroyWindow(h); return 0; }
        break;
    case WM_CLOSE: DestroyWindow(h); return 0;
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
        wc.lpfnWndProc = PickerProc; wc.hInstance = g_hInst;
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
    HWND hwnd = CreateWindowExW(WS_EX_TOOLWINDOW, L"PlaylistPickerClass", L"Playlist",
        WS_POPUPWINDOW | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        X, Y, W, H, g_hwndMain, NULL, g_hInst, NULL);
    g_pick.hwnd = hwnd;

    MSG msg;
    while (IsWindow(hwnd) && GetMessageW(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}

// ----------------------------- Combine log window + thread

static LRESULT CALLBACK CombineLogProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    CombineTask* task = reinterpret_cast<CombineTask*>(GetWindowLongPtrW(h, GWLP_USERDATA));
    switch (m) {
    case WM_CREATE: {
        LPCREATESTRUCT pcs = (LPCREATESTRUCT)l;
        task = (CombineTask*)pcs->lpCreateParams;
        SetWindowLongPtrW(h, GWLP_USERDATA, (LONG_PTR)task);

        HFONT hf = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        RECT rc; GetClientRect(h, &rc);

        HWND hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL,
            4, 4, rc.right - 8, rc.bottom - 8, h, (HMENU)101, g_hInst, NULL);
        SendMessageW(hEdit, WM_SETFONT, (WPARAM)hf, TRUE);

        if (task) task->hEdit = hEdit;
        return 0;
    }
    case WM_SIZE: {
        if (task && task->hEdit) {
            RECT rc; GetClientRect(h, &rc);
            MoveWindow(task->hEdit, 4, 4, rc.right - 8, rc.bottom - 8, TRUE);
        }
        return 0;
    }
    case WM_CLOSE:
        ShowWindow(h, SW_HIDE); // hide, don't destroy; preserve log
        return 0;
    }
    return DefWindowProcW(h, m, w, l);
}

static void EnsureCombineLogClass() {
    static bool reg = false;
    if (!reg) {
        WNDCLASSW wc{}; wc.lpfnWndProc = CombineLogProc; wc.hInstance = g_hInst;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"CombineLogClass";
        RegisterClassW(&wc);
        reg = true;
    }
}

static HWND CreateCombineLogWindow(CombineTask* task) {
    RECT r; SystemParametersInfoW(SPI_GETWORKAREA, 0, &r, 0);
    int W = 640, H = 480;
    int X = r.left + ((r.right - r.left) - W) / 2;
    int Y = r.top + ((r.bottom - r.top) - H) / 2;

    std::wstring title = L"Combine: ";
    title += task ? task->title : L"(unknown)";

    HWND hwnd = CreateWindowExW(WS_EX_TOOLWINDOW, L"CombineLogClass", title.c_str(),
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        X, Y, W, H, g_hwndMain, NULL, g_hInst, task);
    return hwnd;
}

static void PostCombineOutput(CombineTask* task, const std::wstring& text) {
    if (!task) return;
    std::wstring* p = new std::wstring(text);
    PostMessageW(g_hwndMain, WM_APP_COMBINE_OUTPUT, (WPARAM)task, (LPARAM)p);
}

// -------- Embedded video_combine logic (internal, ffmpeg-based) --------

// Narrow/wide helpers using the ANSI code page (matches old console argv behavior).
static std::string NarrowFromWideACP(const std::wstring& ws) {
    if (ws.empty()) return std::string();
    int len = WideCharToMultiByte(CP_ACP, 0, ws.c_str(), -1, NULL, 0, NULL, NULL);
    if (len <= 0) return std::string();
    std::string s(len - 1, '\0');
    WideCharToMultiByte(CP_ACP, 0, ws.c_str(), -1, &s[0], len, NULL, NULL);
    return s;
}

static std::wstring WideFromNarrowACP(const std::string& s) {
    if (s.empty()) return std::wstring();
    int len = MultiByteToWideChar(CP_ACP, 0, s.c_str(), -1, NULL, 0);
    if (len <= 0) return std::wstring();
    std::wstring ws(len - 1, L'\0');
    MultiByteToWideChar(CP_ACP, 0, s.c_str(), -1, &ws[0], len);
    return ws;
}

// Log an ANSI command line into the combine log window
static void VC_LogCmd(CombineTask* task, const char* cmd) {
    if (!task || !cmd) return;
    int n = MultiByteToWideChar(CP_ACP, 0, cmd, -1, NULL, 0);
    if (n <= 0) return;
    std::wstring w(n - 1, L'\0');
    MultiByteToWideChar(CP_ACP, 0, cmd, -1, &w[0], n);
    PostCombineOutput(task, w + L"\r\n");
}

static void VC_LogMsg(CombineTask* task, const std::wstring& msg) {
    if (!task) return;
    PostCombineOutput(task, msg);
}

// Update combine window title with a status suffix, e.g. "(ffmpeg in progress...)"
static void VC_UpdateCombineWindowTitle(CombineTask* task, const wchar_t* statusSuffix) {
    if (!task || !task->hwnd) return;
    std::wstring title = L"Combine: ";
    title += task->title;
    if (statusSuffix && *statusSuffix) {
        title += L" ";
        title += statusSuffix;
    }
    SetWindowTextW(task->hwnd, title.c_str());
}

// Run a command in a hidden console and wait for it to finish.
// Returns the process exit code, or (DWORD)-1 on failure to spawn.
static DWORD RunHiddenCommandAnsi(const char* cmdAnsi) {
    if (!cmdAnsi || !*cmdAnsi) return (DWORD)-1;

    std::wstring w = WideFromNarrowACP(cmdAnsi);
    if (w.empty()) return (DWORD)-1;

    std::vector<wchar_t> cmdBuf(w.size() + 1);
    wcscpy_s(cmdBuf.data(), cmdBuf.size(), w.c_str());

    STARTUPINFOW si{};
    si.cb = sizeof(si);
    PROCESS_INFORMATION pi{};
    DWORD flags = CREATE_NO_WINDOW;

    BOOL ok = CreateProcessW(
        NULL,               // use command line for program+args
        cmdBuf.data(),      // modifiable buffer
        NULL, NULL,
        FALSE,
        flags,
        NULL, NULL,
        &si, &pi);
    if (!ok) {
        return (DWORD)-1;
    }

    WaitForSingleObject(pi.hProcess, INFINITE);
    DWORD exitCode = 0;
    GetExitCodeProcess(pi.hProcess, &exitCode);
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return exitCode;
}

// Same, but execute via "cmd.exe /C <command>" for shell built-ins like copy.
static DWORD RunHiddenCommandAnsiViaCmd(const char* innerCmdAnsi) {
    if (!innerCmdAnsi || !*innerCmdAnsi) return (DWORD)-1;
    char full[9000];
    sprintf(full, "cmd.exe /C %s", innerCmdAnsi);
    return RunHiddenCommandAnsi(full);
}


// From video_combine.cpp: prevent duplicate entries in list.
static bool nodup_add(std::vector<std::string>& list, const std::string& new_file) {
    bool retval = false;
    int limit = (int)list.size();
    for (int index = 0; index < limit; ++index) {
        if (list[index] == new_file) {
            retval = true;
            break;
        }
    }
    if (!retval) list.push_back(new_file);
    return retval;
}

// ---- Direct ports of video_combine.cpp functions (same method, in-process) ----

// Stage 1: convert each source to an .mpg using ffmpeg -qscale:v 1
static bool convert2mpg(const std::vector<std::string>& srcfn,
    std::vector<std::string>& mpgfn,
    CombineTask* task)
{
    bool retval = true;
    char drive[260];
    char path[260];
    char fn[260];
    int  limit = (int)srcfn.size();
    char mpgfile[260];
    char cmd[8192];

    for (int index = 0; index < limit; ++index) {
        _splitpath(srcfn[index].c_str(), drive, path, fn, NULL);
        sprintf(mpgfile, "%s%s%s.mpg", drive, path, fn);
        mpgfn.push_back(mpgfile);

        // ffmpeg -i "src" -qscale:v 1 "dst.mpg"
        sprintf(cmd, "ffmpeg -i \"%s\" -qscale:v 1 \"%s\"",
            srcfn[index].c_str(), mpgfile);
        VC_LogCmd(task, cmd);

        DWORD exitCode = RunHiddenCommandAnsi(cmd);
        if (exitCode == 0) {
            // success
        }
        else {
            char fail[9000];
            sprintf(fail, "%s failed (exit=%lu)", cmd, (unsigned long)exitCode);
            VC_LogCmd(task, fail);
            retval = false;
            break;
        }
    }
    return retval;
}

// Stage 2: copy /B file1.mpg+file2.mpg+... combined.mpg
static bool combinempg(const std::vector<std::string>& mpgfn,
    std::string& combinefile,
    CombineTask* task)
{
    bool retval = true;
    char cmd[8192];
    int  limit = (int)mpgfn.size();

    strcpy(cmd, "copy /B ");

    for (int index = 0; index < limit; ++index) {
        strcat(cmd, "\"");
        strcat(cmd, mpgfn[index].c_str());
        strcat(cmd, "\" ");

        if (index < limit - 1)
            strcat(cmd, "+ ");
    }

    strcat(cmd, " \"");
    strcat(cmd, combinefile.c_str());
    strcat(cmd, "\"");

    VC_LogCmd(task, cmd);

    DWORD exitCode = RunHiddenCommandAnsiViaCmd(cmd);
    if (exitCode == 0) {
        // Delete intermediate .mpg files just like original code
        for (int index = 0; index < limit; ++index) {
            _unlink(mpgfn[index].c_str());
        }
    }
    else {
        char fail[9000];
        sprintf(fail, "%s failed (exit=%lu)", cmd, (unsigned long)exitCode);
        VC_LogCmd(task, fail);
        retval = false;
    }
    return retval;
}

// Stage 3: convert combined.mpg back to final format using ffmpeg -qscale:v 2
static bool convertback(const std::string& combinefile,
    const std::string& finalfile,
    CombineTask* task)
{
    char cmd[8192];
    bool retval = true;

    // ffmpeg -i "combined.mpg" -qscale:v 2 "final.ext"
    sprintf(cmd, "ffmpeg -i \"%s\" -qscale:v 2 \"%s\"",
        combinefile.c_str(), finalfile.c_str());
    VC_LogCmd(task, cmd);

    DWORD exitCode = RunHiddenCommandAnsi(cmd);
    if (exitCode == 0) {
        // success
    }
    else {
        char fail[9000];
        sprintf(fail, "%s failed (exit=%lu)", cmd, (unsigned long)exitCode);
        VC_LogCmd(task, fail);
        retval = false;
    }
    return retval;
}

// Full video_combine "combine_videos" logic as a function:
// uses the 3 stages above, same algorithm as original tool.
static bool VC_CombineVideos(CombineTask* task,
    const std::vector<std::string>& srcfn,
    std::string& finalfile)
{
    bool retval = false;
    std::vector<std::string> mpgfn;
    std::string combinefile;
    char drive[260];
    char path[260];
    char fn[260];

    int limit = (int)srcfn.size();
    if (limit == 0) return false;

    std::wstring wFinal = WideFromNarrowACP(finalfile);
    VC_LogMsg(task, L"Start Combining Video " + wFinal + L"\r\n");

    if (convert2mpg(srcfn, mpgfn, task)) {
        _splitpath(finalfile.c_str(), drive, path, fn, NULL);
        combinefile = drive;
        combinefile += path;
        combinefile += fn;
        combinefile += ".mpg";

        if (combinempg(mpgfn, combinefile, task)) {
            if (convertback(combinefile, finalfile, task)) {
                VC_LogMsg(task, L"Video combined successful for " + wFinal + L"\r\n");
                retval = true;
                _unlink(combinefile.c_str());
            }
        }
    }

    if (!retval) {
        VC_LogMsg(task, L"Video combined failed for " + WideFromNarrowACP(finalfile) + L"\r\n");
    }

    return retval;
}

// Run the embedded video_combine algorithm using the copied .mp4/.mkv files
// and the intended user-chosen output name (combinedFull).
static bool RunEmbeddedVideoCombine(CombineTask* task,
    const std::vector<std::wstring>& copiedFiles,
    const std::wstring& combinedFull)
{
    // Build ANSI source list with duplicate check...
    std::vector<std::string> srcAnsi;
    srcAnsi.reserve(copiedFiles.size());

    for (const auto& wf : copiedFiles) {
        std::string a = NarrowFromWideACP(wf);
        if (a.empty()) continue;
        if (nodup_add(srcAnsi, a)) {
            VC_LogMsg(task, L"Error Duplicate File in Combine filelist:\r\n  " + wf + L"\r\n");
            return false;
        }
    }
    if (srcAnsi.empty()) {
        VC_LogMsg(task, L"No source files to combine.\r\n");
        return false;
    }

    // Emulate video_combine's main() behavior...
    std::string destAnsi = NarrowFromWideACP(combinedFull);
    char drive[260], path[260], fn[260], ext[260], lastfn[260];
    _splitpath(destAnsi.c_str(), drive, path, fn, ext);
    sprintf(lastfn, "%s%s%s_combined%s", drive, path, fn, ext);
    std::string finalfile(lastfn);

    // Show ffmpeg progress in the combine window title while we work
    VC_UpdateCombineWindowTitle(task, L"(ffmpeg in progress...)");

    bool ok = VC_CombineVideos(task, srcAnsi, finalfile);

    // Update task->combinedFull to the actual final path (with _combined suffix)
    if (ok) {
        task->combinedFull = WideFromNarrowACP(finalfile);
    }

    VC_UpdateCombineWindowTitle(task, ok ? L"(done)" : L"(failed)");
    return ok;
}

// Delete the combine working directory and its files (shallow) on success.
static void DeleteCombineWorkingDirIfExists(CombineTask* task) {
    if (!task) return;
    if (task->workingDir.empty()) return;

    std::wstring dir = task->workingDir;

    // Ensure trailing slash for building file paths
    std::wstring pattern = EnsureSlash(dir) + L"*";
    WIN32_FIND_DATAW fd{};
    HANDLE hFind = FindFirstFileW(pattern.c_str(), &fd);
    if (hFind != INVALID_HANDLE_VALUE) {
        do {
            if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0)
                continue;

            std::wstring full = EnsureSlash(dir) + fd.cFileName;

            if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
                // Working dir should be flat; skip any subdirs just in case.
                continue;
            }
            DeleteFileW(full.c_str());
        } while (FindNextFileW(hFind, &fd));
        FindClose(hFind);
    }

    if (!RemoveDirectoryW(dir.c_str())) {
        DWORD err = GetLastError();
        wchar_t buf[512];
        swprintf_s(buf,
            L"Warning: failed to remove combine working directory \"%s\" (err=%lu)\r\n",
            dir.c_str(), err);
        PostCombineOutput(task, buf);
        LogLine(L"Failed to remove combine working directory \"%s\" err=%lu",
            dir.c_str(), err);
    }
    else {
        LogLine(L"Removed combine working directory \"%s\"", dir.c_str());
    }
}

static DWORD WINAPI CombineThreadProc(LPVOID param) {
    CombineTask* task = (CombineTask*)param;
    if (!task) return 0;

    // Create working directory
    if (!CreateDirectoryW(task->workingDir.c_str(), NULL)) {
        DWORD e = GetLastError();
        if (e != ERROR_ALREADY_EXISTS) {
            std::wstring msg = L"ERROR: Failed to create working directory:\r\n";
            msg += task->workingDir;
            msg += L"\r\n";
            PostCombineOutput(task, msg);
            PostMessageW(g_hwndMain, WM_APP_COMBINE_DONE, (WPARAM)task, (LPARAM)1);
            return 0;
        }
    }

    // Copy files into working directory (same as before)
    {
        wchar_t buf[256];
        swprintf_s(buf, L"Copying %zu file(s)...\r\n", task->srcFiles.size());
        PostCombineOutput(task, buf);
    }

    std::vector<std::wstring> copiedFiles;
    copiedFiles.reserve(task->srcFiles.size());

    for (size_t i = 0; i < task->srcFiles.size(); ++i) {
        const std::wstring& src = task->srcFiles[i];
        const wchar_t* base = wcsrchr(src.c_str(), L'\\');
        base = base ? base + 1 : src.c_str();

        std::wstring dst = task->workingDir;
        dst += base;

        std::wstring line = L"  -> ";
        line += dst;
        line += L"\r\n";
        PostCombineOutput(task, line);

        if (!CopyFileW(src.c_str(), dst.c_str(), FALSE)) {
            std::wstring err = L"ERROR: Failed to copy file:\r\n";
            err += src;
            err += L"\r\n";
            PostCombineOutput(task, err);
            PostMessageW(g_hwndMain, WM_APP_COMBINE_DONE, (WPARAM)task, (LPARAM)2);
            return 0;
        }
        copiedFiles.push_back(dst);
    }

    PostCombineOutput(task, L"All files copied. Combining via internal ffmpeg pipeline...\r\n");

    // --- NEW: run embedded video_combine algorithm instead of video_combine.exe ---
    bool ok = RunEmbeddedVideoCombine(task, copiedFiles, task->combinedFull);

    DWORD exitCode = ok ? 0 : 1;

    // On success, remove the working directory that held the copied parts
    if (ok) {
        DeleteCombineWorkingDirIfExists(task);
    }

    wchar_t doneMsg[256];
    swprintf_s(doneMsg, L"\r\n[internal video combine %s with code %lu]\r\n",
        ok ? L"succeeded" : L"failed", exitCode);
    PostCombineOutput(task, doneMsg);

    // Notify main window that the combine task is done (same message ID as before)
    PostMessageW(g_hwndMain, WM_APP_COMBINE_DONE, (WPARAM)task, (LPARAM)exitCode);
    return 0;
}

// ----------------------------- Subclasses
static LRESULT CALLBACK ListSubclass(HWND h, UINT m, WPARAM w, LPARAM l,
    UINT_PTR, DWORD_PTR) {
    if (m == WM_GETDLGCODE) return DLGC_WANTALLKEYS;
    if (m == WM_KEYDOWN) {
        bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;

        // Row reordering: Ctrl+Up / Ctrl+Down
        if (ctrl && (w == VK_UP || w == VK_DOWN)) {
            Browser_MoveSelectedRow(w == VK_UP ? -1 : +1);
            return 0;
        }

        // Combine selected: Ctrl+'+' (main plus key or numpad +)
        if (ctrl && (w == VK_OEM_PLUS || w == VK_ADD)) {
            Browser_CombineSelected();
            return 0;
        }

        switch (w) {
        case VK_RETURN: ActivateSelection(); return 0;
        case VK_LEFT:
        case VK_BACK:   NavigateBack(); return 0;
        case VK_F1:     ShowHelp(); return 0;

        case 'A': if (ctrl) {
            for (int i = 0; i < (int)g_rows.size(); ++i) {
                if (!g_rows[i].isDir) {
                    ListView_SetItemState(g_hwndList, i, LVIS_SELECTED, LVIS_SELECTED);
                }
                else {
                    ListView_SetItemState(g_hwndList, i, 0, LVIS_SELECTED);
                }
            }
            return 0;
        }
        case 'P': if (ctrl) { PlaySelectedVideos(); return 0; }

        case 'F': if (ctrl) {
            std::wstring kw;
            if (!PromptKeyword(kw)) return 0;
            kw = ToLower(kw);
            if (kw.empty()) return 0;

            if (g_view != ViewKind::Search) {
                g_search.active = true;
                g_search.originView = g_view;
                g_search.originFolder = (g_view == ViewKind::Folder ? g_folder : L"");
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
                    if (NameContainsAllTerms(g_rows[i].full, g_search.termsLower)) filtered.push_back(g_rows[i]);
                }
                ShowSearchResults(filtered);
            }
            return 0;
        }

        case 'C': if (ctrl) { Browser_CopySelectedToClipboard(ClipMode::Copy); return 0; }
        case 'X': if (ctrl) { Browser_CopySelectedToClipboard(ClipMode::Move); return 0; }
        case 'V': if (ctrl) { Browser_PasteClipboardIntoCurrent(); return 0; }
        case VK_DELETE: Browser_DeleteSelected(); return 0;
        }
    }
    return DefSubclassProc(h, m, w, l);
}

// ----------------------------- Post-playback upscale submit

static void ScheduleUpscaleForCurrentVideo() {
    if (!g_inPlayback || g_playlist.empty()) return;
    if (g_cfg.upscaleDirectory.empty()) {
        MessageBoxW(g_hwndMain,
            L"Upscale directory is not configured.\n"
            L"Set 'upscaleDirectory = ...' in mediaexplorer.ini.",
            L"Submit for upscaling", MB_OK);
        return;
    }

    const std::wstring& cur = g_playlist[g_playlistIndex];

    // Build target path in configured upscaleDirectory
    const wchar_t* base = wcsrchr(cur.c_str(), L'\\');
    base = base ? base + 1 : cur.c_str();

    wchar_t fname[_MAX_FNAME] = {}, ext[_MAX_EXT] = {};
    _wsplitpath_s(base, NULL, 0, NULL, 0, fname, _MAX_FNAME, ext, _MAX_EXT);

    std::wstring dst = UniqueName(g_cfg.upscaleDirectory, fname, ext);

    LogLine(L"Submit for upscaling queued: src=\"%s\" dst=\"%s\"",
        cur.c_str(), dst.c_str());

    // Queue as a post-playback copy
    g_post.push_back({ ActionType::CopyToPath, cur, dst });

    MessageBoxW(g_hwndMain,
        L"Video will be copied to upscaleDirectory at the end of playback.",
        L"Submit for upscaling", MB_OK);
}

// ----------------------------- FFmpeg task scheduler

static void ScheduleFfmpegTask(FfmpegOpKind kind) {
    if (!g_cfg.ffmpegAvailable) {
        MessageBoxW(g_hwndMain,
            L"ffmpegAvailable is not enabled in mediaexplorer.ini.\n"
            L"Set ffmpegAvailable = 1 to use FFmpeg tools.",
            L"FFmpeg tools", MB_OK);
        return;
    }
    if (!g_inPlayback || g_playlist.empty() || !g_mp) return;

    const std::wstring& cur = g_playlist[g_playlistIndex];

    // Get current playback time in ms
    libvlc_time_t refMs = libvlc_media_player_get_time(g_mp);
    if (refMs < 0) refMs = 0;

    // Determine folder + base + ext
    std::wstring folder = cur;
    PathRemoveFileSpecW(&folder[0]);
    folder = folder.c_str(); // shrink
    folder = EnsureSlash(folder);

    std::wstring workingDir = folder + L"video_process\\";

    const wchar_t* baseName = wcsrchr(cur.c_str(), L'\\');
    baseName = baseName ? baseName + 1 : cur.c_str();

    wchar_t fname[_MAX_FNAME] = {}, ext[_MAX_EXT] = {};
    _wsplitpath_s(baseName, NULL, 0, NULL, 0, fname, _MAX_FNAME, ext, _MAX_EXT);

    std::wstring inputCopy = workingDir;
    inputCopy += fname;
    inputCopy += ext;

    std::wstring outputTemp = workingDir;
    outputTemp += fname;
    switch (kind) {
    case FfmpegOpKind::TrimFront: outputTemp += L"_trimfront"; break;
    case FfmpegOpKind::TrimEnd:   outputTemp += L"_trimend";   break;
    case FfmpegOpKind::HFlip:     outputTemp += L"_hflip";     break;
    }
    outputTemp += ext;

    FfmpegTask* task = new FfmpegTask();
    task->sourceFull = cur;
    task->workingDir = workingDir;
    task->inputCopy = inputCopy;
    task->outputTemp = outputTemp;
    task->refMs = refMs;
    task->kind = kind;
    task->running = true;

    LogLine(L"FFmpegTask scheduled: kind=%d src=\"%s\" refMs=%lld workingDir=\"%s\"",
        (int)kind, cur.c_str(), (long long)refMs, workingDir.c_str());

    // Title for the log window
    switch (kind) {
    case FfmpegOpKind::TrimFront: task->title = L"Trim front: "; break;
    case FfmpegOpKind::TrimEnd:   task->title = L"Trim end: ";   break;
    case FfmpegOpKind::HFlip:     task->title = L"Horizontal flip: "; break;
    }
    task->title += baseName;

    EnsureFfmpegLogClass();
    HWND logWnd = CreateFfmpegLogWindow(task);
    if (!logWnd) {
        delete task;
        MessageBoxW(g_hwndMain, L"Failed to create FFmpeg task log window.",
            L"FFmpeg tools", MB_OK);
        return;
    }
    task->hwnd = logWnd;

    EnterCriticalSection(&g_ffLock);
    g_ffTasks.push_back(task);
    LeaveCriticalSection(&g_ffLock);

    HANDLE hThread = CreateThread(NULL, 0, FfmpegThreadProc, task, 0, NULL);
    if (!hThread) {
        EnterCriticalSection(&g_ffLock);
        auto it = std::find(g_ffTasks.begin(), g_ffTasks.end(), task);
        if (it != g_ffTasks.end()) g_ffTasks.erase(it);
        LeaveCriticalSection(&g_ffLock);

        if (IsWindow(task->hwnd)) DestroyWindow(task->hwnd);
        delete task;

        MessageBoxW(g_hwndMain, L"Failed to start background thread for FFmpeg task.",
            L"FFmpeg tools", MB_OK);
        return;
    }

    task->hThread = hThread;
}

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

        case 'V':
            if (ctrl) {
                // Pause while we show the tools menu
                bool wasPlaying = (libvlc_media_player_is_playing(g_mp) > 0);
                if (wasPlaying) libvlc_media_player_set_pause(g_mp, 1);

                bool canUpscale = !g_cfg.upscaleDirectory.empty();
                bool canFfmpeg = g_cfg.ffmpegAvailable;
                int choice = PromptVideoToolsChoice(canUpscale, canFfmpeg);

                if (choice == 1 && canUpscale) {
                    ScheduleUpscaleForCurrentVideo();
                }
                else if (choice == 2 && canFfmpeg) {
                    ScheduleFfmpegTask(FfmpegOpKind::TrimFront);
                }
                else if (choice == 3 && canFfmpeg) {
                    ScheduleFfmpegTask(FfmpegOpKind::TrimEnd);
                }
                else if (choice == 4 && canFfmpeg) {
                    ScheduleFfmpegTask(FfmpegOpKind::HFlip);
                }

                if (wasPlaying) libvlc_media_player_set_pause(g_mp, 0);
                return 0;
            }
            break;


        case VK_DELETE: {
            if (!g_playlist.empty()) {
                std::wstring doomed = g_playlist[g_playlistIndex];
                std::vector<std::wstring> np;
                np.reserve(g_playlist.size());
                for (size_t i = 0; i < g_playlist.size(); ++i) if (i != g_playlistIndex) np.push_back(g_playlist[i]);
                g_playlist.swap(np);
                g_post.push_back({ ActionType::DeleteFile, doomed, L"" });
                if (g_playlist.empty()) ExitPlayback();
                else if (g_playlistIndex >= g_playlist.size()) PlayIndex(g_playlist.size() - 1);
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
            if (ctrl) { if (w == VK_RIGHT) NextInPlaylist(); else PrevInPlaylist(); }
            else {
                libvlc_time_t cur = libvlc_media_player_get_time(g_mp);
                libvlc_time_t len = libvlc_media_player_get_length(g_mp);
                libvlc_time_t step = shift ? 60000 : 10000; // 60s / 10s
                if (w == VK_RIGHT) cur += step; else cur = (cur > step ? cur - step : 0);
                if (len > 0 && cur > len) cur = len;
                libvlc_media_player_set_time(g_mp, cur);
            }
            return 0;
        }
        }
    }
    return DefSubclassProc(h, m, w, l);
}

static LRESULT CALLBACK SeekSubclass(HWND h, UINT m, WPARAM w, LPARAM l,
    UINT_PTR, DWORD_PTR) {
    if (m == WM_KEYDOWN) {
        if (w == VK_F1) { ShowHelp(); return 0; }
        if (w == VK_ESCAPE) { ExitPlayback(); return 0; }
        if (w == VK_RETURN) { ToggleFullscreen(); return 0; }
        if (w == VK_LEFT || w == VK_RIGHT || w == VK_UP || w == VK_DOWN ||
            w == VK_SPACE || w == VK_TAB || w == VK_DELETE) {
            SendMessageW(g_hwndVideo, WM_KEYDOWN, w, l);
            return 0;
        }
        bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        if (ctrl && (w == 'R' || w == 'r')) { SendMessageW(g_hwndVideo, WM_KEYDOWN, 'R', 0); return 0; }
        if (ctrl && (w == 'C' || w == 'c')) { SendMessageW(g_hwndVideo, WM_KEYDOWN, 'C', 0); return 0; }
        if (ctrl && (w == 'G' || w == 'g')) { ShowPlaylistChooser(); return 0; }
        if (ctrl && (w == 'P' || w == 'p')) { SendMessageW(g_hwndVideo, WM_KEYDOWN, 'P', 0); return 0; }
        if (ctrl && (w == 'V' || w == 'v')) { SendMessageW(g_hwndVideo, WM_KEYDOWN, 'V', 0); return 0; }

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

// ------------- Icon loader
static HICON LoadAppIcon(int cx, int cy) {
    wchar_t exePath[MAX_PATH] = {};
    GetModuleFileNameW(NULL, exePath, MAX_PATH);
    PathRemoveFileSpecW(exePath);
    std::wstring p = std::wstring(exePath) + L"\\MediaExplorer.ico";

    HICON h = (HICON)LoadImageW(NULL, p.c_str(), IMAGE_ICON, cx, cy, LR_LOADFROMFILE);
    if (!h) {
        h = (HICON)LoadImageW(NULL, L"MediaExplorer.ico", IMAGE_ICON, cx, cy, LR_LOADFROMFILE);
    }
    return h;
}

// ----------------------------- Window proc
LRESULT CALLBACK WndProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    switch (m) {
    case WM_CREATE: {
        INITCOMMONCONTROLSEX icc; icc.dwSize = sizeof(icc);
        icc.dwICC = ICC_LISTVIEW_CLASSES | ICC_BAR_CLASSES;
        InitCommonControlsEx(&icc);

        InitializeCriticalSection(&g_metaLock);
        InitializeCriticalSection(&g_combineLock);
        InitializeCriticalSection(&g_ffLock);   // NEW

        g_hwndList = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"",
            WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS,
            0, 0, 100, 100, h, (HMENU)1001, g_hInst, NULL);
        ListView_SetExtendedListViewStyle(g_hwndList,
            LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
        LV_ResetColumns();
        SetWindowSubclass(g_hwndList, ListSubclass, 1, 0);

        g_hwndVideo = CreateWindowExW(0, L"STATIC", L"",
            WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN,
            0, 0, 100, 100, h, (HMENU)1002, g_hInst, NULL);
        ShowWindow(g_hwndVideo, SW_HIDE);
        SetWindowSubclass(g_hwndVideo, VideoSubclass, 2, 0);

        g_hwndSeek = CreateWindowExW(0, TRACKBAR_CLASSW, L"",
            WS_CHILD | TBS_HORZ | TBS_AUTOTICKS,
            0, 0, 100, 30, h, (HMENU)1003, g_hInst, NULL);
        ShowWindow(g_hwndSeek, SW_HIDE);
        SetWindowSubclass(g_hwndSeek, SeekSubclass, 3, 0);

        ShowDrives();
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
            if (nm->code == NM_DBLCLK || nm->code == LVN_ITEMACTIVATE) {
                ActivateSelection(); return 0;
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
        if (g_inPlayback) { SendMessageW(g_hwndVideo, WM_KEYDOWN, w, l); return 0; }
        break;

    case WM_APP + 1: // VLC: end reached
        if (g_inPlayback && g_playlistIndex + 1 < g_playlist.size()) NextInPlaylist();
        else if (g_inPlayback) ExitPlayback();
        return 0;

    case WM_APP_META: {
        MetaResult* r = (MetaResult*)l;
        if (r) {
            if (r->gen == g_metaGen.load(std::memory_order_relaxed)) {
                for (int i = 0; i < (int)g_rows.size(); ++i) {
                    if (_wcsicmp(g_rows[i].full.c_str(), r->path.c_str()) == 0) {
                        Row& it = g_rows[i];
                        it.vW = r->w; it.vH = r->h; it.vDur100ns = r->dur;
                        if (!it.isDir) {
                            if (it.vW > 0 && it.vH > 0) {
                                wchar_t buf[64]; swprintf_s(buf, L"%dx%d", it.vW, it.vH);
                                ListView_SetItemText(g_hwndList, i, 4, buf);
                            }
                            if (it.vDur100ns > 0) {
                                std::wstring ds = FormatDuration100ns(it.vDur100ns);
                                ListView_SetItemText(g_hwndList, i, 5, const_cast<wchar_t*>(ds.c_str()));
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

    case WM_APP_FFMPEG_OUTPUT: {
        FfmpegTask* task = (FfmpegTask*)w;
        std::wstring* p = (std::wstring*)l;
        if (task && task->hEdit && p) {
            HWND edit = task->hEdit;
            SendMessageW(edit, EM_SETSEL, (WPARAM)-1, (LPARAM)-1);
            SendMessageW(edit, EM_REPLACESEL, FALSE, (LPARAM)p->c_str());
            SendMessageW(edit, EM_SCROLLCARET, 0, 0);
        }
        delete p;
        return 0;
    }

    case WM_APP_FFMPEG_DONE: {
        FfmpegTask* task = (FfmpegTask*)w;
        DWORD exitCode = (DWORD)l;
        UNREFERENCED_PARAMETER(exitCode);

        EnterCriticalSection(&g_ffLock);
        for (FfmpegTask* t : g_ffTasks) {
            if (t == task) {
                t->running = false;
                t->done = true;
                t->exitCode = exitCode;
                break;
            }
        }
        LeaveCriticalSection(&g_ffLock);
        return 0;
    }

    case WM_APP_COMBINE_OUTPUT: {
        CombineTask* task = (CombineTask*)w;
        std::wstring* p = (std::wstring*)l;
        if (task && task->hEdit && p) {
            HWND edit = task->hEdit;
            SendMessageW(edit, EM_SETSEL, (WPARAM)-1, (LPARAM)-1);
            SendMessageW(edit, EM_REPLACESEL, FALSE, (LPARAM)p->c_str());
            SendMessageW(edit, EM_SCROLLCARET, 0, 0);
        }
        if (p && !p->empty()) {
            // Mirror to main log file (if loggingEnabled)
            LogLine(L"[Combine] %s", p->c_str());
        }
        delete p;
        return 0;
    }

    case WM_APP_COMBINE_DONE: {
        CombineTask* task = (CombineTask*)w;
        DWORD exitCode = (DWORD)l;
        UNREFERENCED_PARAMETER(exitCode);

        EnterCriticalSection(&g_combineLock);
        for (CombineTask* t : g_combineTasks) {
            if (t == task) {
                t->running = false;
                break;
            }
        }
        LeaveCriticalSection(&g_combineLock);

        if (task) {
            if (task->hProcess) { CloseHandle(task->hProcess); task->hProcess = NULL; }
            if (task->hThread) { CloseHandle(task->hThread); task->hThread = NULL; }
        }

        if (g_view == ViewKind::Folder) {
            ShowFolder(g_folder);
        }
        return 0;
    }

    case WM_CLOSE:
        if (HasRunningCombineTasks() || HasRunningFfmpegTasks()) {
            MessageBoxW(h,
                L"Background video operations are still running.\n"
                L"Please wait for them to finish before exiting Media Explorer.",
                L"Background tasks in progress",
                MB_OK);
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
        DeleteCriticalSection(&g_combineLock);
        DeleteCriticalSection(&g_ffLock);       // NEW

        // Cleanup any remaining FFmpeg tasks (should be none because ExitPlayback waits,
        // but this is a safety net)
        EnterCriticalSection(&g_ffLock);
        for (FfmpegTask* t : g_ffTasks) {
            if (!t) continue;
            if (t->hProcess) CloseHandle(t->hProcess);
            if (t->hThread)  CloseHandle(t->hThread);
            if (t->hwnd && IsWindow(t->hwnd)) DestroyWindow(t->hwnd);
            delete t;
        }
        g_ffTasks.clear();
        LeaveCriticalSection(&g_ffLock);


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

    HMODULE u = GetModuleHandleW(L"user32.dll");
    if (u) {
        typedef BOOL(WINAPI* SetProcessDPIAware_t)();
        SetProcessDPIAware_t setAw = (SetProcessDPIAware_t)GetProcAddress(u, "SetProcessDPIAware");
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
    wc.lpszClassName = L"MediaExplorerWindowClass";
    wc.lpfnWndProc = WndProc;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hIcon = hBig ? hBig : LoadIcon(NULL, IDI_APPLICATION);
    wc.hIconSm = hSm ? hSm : wc.hIcon;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    RegisterClassExW(&wc);

    g_hwndMain = CreateWindowExW(0, wc.lpszClassName, L"Media Explorer (libVLC)",
        WS_OVERLAPPEDWINDOW | WS_VISIBLE, CW_USEDEFAULT, CW_USEDEFAULT, 1500, 700,
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
