#include <windows.h>
#include <commctrl.h>   // InitCommonControlsEx
#include <shellapi.h>   // Shell_NotifyIcon, IsUserAnAdmin
#include <shlobj.h>     // SHGetFolderPathW
#include <shlwapi.h>    // PathCombine, PathFileExists
#include <strsafe.h>
#include <string>
#include "resource.h"

// ------------------ Globals ------------------
HINSTANCE g_hInst = nullptr;
HWND g_hDlg = nullptr;
bool g_enableAutoInput = true;

// Cached strings for Alt+1..9
std::wstring g_hotStrings[9];

// Config paths
std::wstring g_configDir;   // %APPDATA%\AutoInputTool
std::wstring g_configPath;  // %APPDATA%\AutoInputTool\config.ini

// Tray icon
NOTIFYICONDATAW g_nid = {};

// Single instance mutex
HANDLE g_hMutex = nullptr;
const wchar_t* kMutexName = L"Global\\AutoInputTool_SingleInstance";

// ------------------ Helpers ------------------

std::wstring GetExePath()
{
    wchar_t buf[MAX_PATH] = {0};
    GetModuleFileNameW(nullptr, buf, MAX_PATH);
    return std::wstring(buf);
}

bool EnsureDirectoryExists(const std::wstring& dir)
{
    if (PathFileExistsW(dir.c_str())) return true;
    return CreateDirectoryW(dir.c_str(), nullptr) || PathFileExistsW(dir.c_str());
}

bool InitConfigPaths()
{
    wchar_t dir[MAX_PATH] = {0};
    if (FAILED(SHGetFolderPathW(NULL, CSIDL_APPDATA, NULL, 0, dir))) {
        return false;
    }

    wchar_t fullDir[MAX_PATH] = {0};
    PathCombineW(fullDir, dir, L"AutoInputTool");
    g_configDir = fullDir;

    wchar_t cfg[MAX_PATH] = {0};
    PathCombineW(cfg, fullDir, L"config.ini");
    g_configPath = cfg;

    return EnsureDirectoryExists(g_configDir);
}

void LoadConfig()
{
    g_enableAutoInput = true;
    for (int i = 0; i < 9; ++i) g_hotStrings[i].clear();

    if (!PathFileExistsW(g_configPath.c_str())) return;

    // Read hotkeys
    wchar_t buffer[4096];
    for (int i = 0; i < 9; ++i) {
        wchar_t key[16];
        StringCchPrintfW(key, 16, L"Alt%d", i + 1);
        DWORD n = GetPrivateProfileStringW(L"Hotkeys", key, L"", buffer, 4096, g_configPath.c_str());
        g_hotStrings[i] = std::wstring(buffer, n);
    }

    // Read settings
    int enable = GetPrivateProfileIntW(L"Settings", L"EnableAutoInput", 1, g_configPath.c_str());
    g_enableAutoInput = (enable != 0);

    int startup = GetPrivateProfileIntW(L"Settings", L"Startup", 0, g_configPath.c_str());
    if (g_hDlg) {
        CheckDlgButton(g_hDlg, IDC_CHK_ENABLE, g_enableAutoInput ? BST_CHECKED : BST_UNCHECKED);
        CheckDlgButton(g_hDlg, IDC_CHK_STARTUP, startup ? BST_CHECKED : BST_UNCHECKED);
    }
}

void SaveConfig()
{
    // Gather text from UI
    wchar_t buffer[4096];

    for (int i = 0; i < 9; ++i) {
        int ctrlId = IDC_EDIT_ALT1 + (i * 2);
        GetDlgItemTextW(g_hDlg, ctrlId, buffer, 4096);
        g_hotStrings[i] = buffer;

        wchar_t key[16];
        StringCchPrintfW(key, 16, L"Alt%d", i + 1);
        WritePrivateProfileStringW(L"Hotkeys", key, buffer, g_configPath.c_str());
    }

    // Settings
    g_enableAutoInput = (IsDlgButtonChecked(g_hDlg, IDC_CHK_ENABLE) == BST_CHECKED);
    WritePrivateProfileStringW(L"Settings", L"EnableAutoInput", g_enableAutoInput ? L"1" : L"0", g_configPath.c_str());

    bool startupChecked = (IsDlgButtonChecked(g_hDlg, IDC_CHK_STARTUP) == BST_CHECKED);
    WritePrivateProfileStringW(L"Settings", L"Startup", startupChecked ? L"1" : L"0", g_configPath.c_str());

    // Apply startup registry
    HKEY hKey = nullptr;
    if (RegOpenKeyExW(HKEY_CURRENT_USER,
                      L"Software\\Microsoft\\Windows\\CurrentVersion\\Run",
                      0, KEY_SET_VALUE | KEY_QUERY_VALUE, &hKey) == ERROR_SUCCESS)
    {
        if (startupChecked) {
            std::wstring exe = GetExePath();
            std::wstring quoted = L"\"" + exe + L"\"";
            RegSetValueExW(hKey, L"AutoInputTool", 0, REG_SZ,
                           reinterpret_cast<const BYTE*>(quoted.c_str()),
                           static_cast<DWORD>((quoted.size() + 1) * sizeof(wchar_t)));
        } else {
            RegDeleteValueW(hKey, L"AutoInputTool");
        }
        RegCloseKey(hKey);
    }
}

// ---- Input helpers ----

// 主路径：把文本写入剪贴板，然后模拟 Ctrl+V 粘贴（对绝大多数输入框最稳定）
bool PasteTextViaClipboard(const std::wstring& text)
{
    if (text.empty()) return false;

    if (!OpenClipboard(nullptr)) return false;
    EmptyClipboard();

    size_t bytes = (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (!hMem) { CloseClipboard(); return false; }

    void* pMem = GlobalLock(hMem);
    if (!pMem) { GlobalFree(hMem); CloseClipboard(); return false; }
    memcpy(pMem, text.c_str(), bytes);
    GlobalUnlock(hMem);

    if (!SetClipboardData(CF_UNICODETEXT, hMem)) {
        GlobalFree(hMem);
        CloseClipboard();
        return false;
    }
    // hMem ownership transferred to the system
    CloseClipboard();

    // 释放可能残留的 Alt 键（防止组合键状态影响粘贴）
    INPUT relAlt = {};
    relAlt.type = INPUT_KEYBOARD;
    relAlt.ki.wVk = VK_MENU;
    relAlt.ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(1, &relAlt, sizeof(INPUT));

    // 模拟 Ctrl+V
    INPUT in[4] = {};
    in[0].type = INPUT_KEYBOARD; in[0].ki.wVk = VK_CONTROL; // Ctrl down
    in[1].type = INPUT_KEYBOARD; in[1].ki.wVk = 'V';        // V down
    in[2].type = INPUT_KEYBOARD; in[2].ki.wVk = 'V';        // V up
    in[2].ki.dwFlags = KEYEVENTF_KEYUP;
    in[3].type = INPUT_KEYBOARD; in[3].ki.wVk = VK_CONTROL; // Ctrl up
    in[3].ki.dwFlags = KEYEVENTF_KEYUP;

    return SendInput(4, in, sizeof(INPUT)) == 4;
}

// 后备路径：Unicode 逐字符注入
void SendUnicodeText(const std::wstring& text)
{
    for (wchar_t ch : text) {
        INPUT ip = {};
        ip.type = INPUT_KEYBOARD;
        ip.ki.wVk = 0;
        ip.ki.wScan = ch;
        ip.ki.dwFlags = KEYEVENTF_UNICODE;
        SendInput(1, &ip, sizeof(INPUT));

        ip.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
        SendInput(1, &ip, sizeof(INPUT));
    }
}

// 综合模拟：优先剪贴板粘贴，失败则逐字符注入
void SimulateTextInput(const std::wstring& text)
{
    if (text.empty()) return;

    // 如果焦点在我们自己的窗口，尽量让用户可见或不干扰
    HWND fg = GetForegroundWindow();
    if (fg == g_hDlg) {
        // 不强制切换焦点，直接执行粘贴，用户可以在外部窗口按热键
    }

    if (!PasteTextViaClipboard(text)) {
        SendUnicodeText(text);
    }
}

// ---- Hotkeys ----

bool RegisterHotkeys(HWND hWnd)
{
    bool ok = true;
    ok &= (RegisterHotKey(hWnd, HK_ALT1, MOD_ALT, '1') != 0);
    ok &= (RegisterHotKey(hWnd, HK_ALT2, MOD_ALT, '2') != 0);
    ok &= (RegisterHotKey(hWnd, HK_ALT3, MOD_ALT, '3') != 0);
    ok &= (RegisterHotKey(hWnd, HK_ALT4, MOD_ALT, '4') != 0);
    ok &= (RegisterHotKey(hWnd, HK_ALT5, MOD_ALT, '5') != 0);
    ok &= (RegisterHotKey(hWnd, HK_ALT6, MOD_ALT, '6') != 0);
    ok &= (RegisterHotKey(hWnd, HK_ALT7, MOD_ALT, '7') != 0);
    ok &= (RegisterHotKey(hWnd, HK_ALT8, MOD_ALT, '8') != 0);
    ok &= (RegisterHotKey(hWnd, HK_ALT9, MOD_ALT, '9') != 0);
    return ok;
}

void UnregisterHotkeys(HWND hWnd)
{
    UnregisterHotKey(hWnd, HK_ALT1);
    UnregisterHotKey(hWnd, HK_ALT2);
    UnregisterHotKey(hWnd, HK_ALT3);
    UnregisterHotKey(hWnd, HK_ALT4);
    UnregisterHotKey(hWnd, HK_ALT5);
    UnregisterHotKey(hWnd, HK_ALT6);
    UnregisterHotKey(hWnd, HK_ALT7);
    UnregisterHotKey(hWnd, HK_ALT8);
    UnregisterHotKey(hWnd, HK_ALT9);
}

void PopulateUiFromConfig()
{
    for (int i = 0; i < 9; ++i) {
        int ctrlId = IDC_EDIT_ALT1 + (i * 2);
        SetDlgItemTextW(g_hDlg, ctrlId, g_hotStrings[i].c_str());
    }
    CheckDlgButton(g_hDlg, IDC_CHK_ENABLE, g_enableAutoInput ? BST_CHECKED : BST_UNCHECKED);
}

// ------------------ Tray helpers ------------------

void AddTrayIcon(HWND hDlg)
{
    ZeroMemory(&g_nid, sizeof(g_nid));
    g_nid.cbSize = sizeof(g_nid);
    g_nid.hWnd = hDlg;
    g_nid.uID = 1;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon = LoadIconW(g_hInst, MAKEINTRESOURCE(IDI_APPICON));
    wcscpy_s(g_nid.szTip, L"自动输入程序");
    Shell_NotifyIconW(NIM_ADD, &g_nid);
}

void DeleteTrayIcon()
{
    if (g_nid.hWnd) {
        Shell_NotifyIconW(NIM_DELETE, &g_nid);
        ZeroMemory(&g_nid, sizeof(g_nid));
    }
}

void ShowTrayMenu(HWND hDlg)
{
    HMENU hMenu = CreatePopupMenu();
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_OPEN, L"打开");
    AppendMenuW(hMenu, MF_STRING, ID_TRAY_EXIT, L"退出");

    POINT pt;
    GetCursorPos(&pt);
    SetForegroundWindow(hDlg);
    TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_BOTTOMALIGN, pt.x, pt.y, 0, hDlg, NULL);
    DestroyMenu(hMenu);
}

void ShowMainWindow(HWND hDlg)
{
    ShowWindow(hDlg, SW_SHOW);
    SetForegroundWindow(hDlg);
}

// ------------------ Admin + single instance ------------------

bool EnsureSingleInstance()
{
    g_hMutex = CreateMutexW(nullptr, FALSE, kMutexName);
    if (!g_hMutex) return true; // 如果创建失败，继续运行
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        // 已有实例在运行，直接退出当前进程
        return false;
    }
    return true;
}

bool RelaunchAsAdminIfNeeded()
{
    // 使用 IsUserAnAdmin 简单判断管理员
    if (IsUserAnAdmin()) return true;

    std::wstring exe = GetExePath();
    HINSTANCE res = ShellExecuteW(nullptr, L"runas", exe.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    if ((INT_PTR)res <= 32) {
        // 提升失败，仍以普通权限继续运行
        return true;
    }
    // 已成功启动管理员实例，当前进程退出
    return false;
}

// ------------------ Dialog proc ------------------

INT_PTR CALLBACK MainDlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_INITDIALOG: {
        g_hDlg = hDlg;
        g_hInst = (HINSTANCE)GetWindowLongPtrW(hDlg, GWLP_HINSTANCE);

        // 标题栏图标
        HICON hIcon = LoadIconW(g_hInst, MAKEINTRESOURCE(IDI_APPICON));
        if (hIcon) {
            SendMessageW(hDlg, WM_SETICON, ICON_BIG, (LPARAM)hIcon);
            SendMessageW(hDlg, WM_SETICON, ICON_SMALL, (LPARAM)hIcon);
        }

        INITCOMMONCONTROLSEX icc{ sizeof(icc), ICC_STANDARD_CLASSES };
        InitCommonControlsEx(&icc);

        if (!InitConfigPaths()) {
            // 保持静默，不弹窗
        }
        LoadConfig();
        PopulateUiFromConfig();

        RegisterHotkeys(hDlg);
        AddTrayIcon(hDlg);
		ShowWindow(hDlg, SW_HIDE);
        return TRUE;
    }
    case WM_COMMAND: {
        switch (LOWORD(wParam)) {
        case IDC_BTN_SAVE: {
            SaveConfig();
            // 不弹窗提示
            return TRUE;
        }
        case IDC_CHK_ENABLE: {
            g_enableAutoInput = (IsDlgButtonChecked(hDlg, IDC_CHK_ENABLE) == BST_CHECKED);
            return TRUE;
        }
        case ID_TRAY_OPEN: {
            ShowMainWindow(hDlg);
            return TRUE;
        }
        case ID_TRAY_EXIT: {
            DeleteTrayIcon();
            EndDialog(hDlg, 0);
            return TRUE;
        }
        }
        break;
    }
    case WM_HOTKEY: {
        if (!g_enableAutoInput) break;

        int idx = -1;
        switch (wParam) {
        case HK_ALT1: idx = 0; break;
        case HK_ALT2: idx = 1; break;
        case HK_ALT3: idx = 2; break;
        case HK_ALT4: idx = 3; break;
        case HK_ALT5: idx = 4; break;
        case HK_ALT6: idx = 5; break;
        case HK_ALT7: idx = 6; break;
        case HK_ALT8: idx = 7; break;
        case HK_ALT9: idx = 8; break;
        default: break;
        }
        if (idx >= 0) {
            SimulateTextInput(g_hotStrings[idx]);
        }
        return TRUE;
    }
    case WM_TRAYICON: {
        if (lParam == WM_LBUTTONUP) {
            ShowMainWindow(hDlg);
        } else if (lParam == WM_RBUTTONUP) {
            ShowTrayMenu(hDlg);
        }
        return TRUE;
    }
    case WM_CLOSE:
        // 关闭按钮最小化到托盘
        ShowWindow(hDlg, SW_HIDE);
        return TRUE;
    case WM_DESTROY:
        UnregisterHotkeys(hDlg);
        DeleteTrayIcon();
        if (g_hMutex) CloseHandle(g_hMutex);
        return TRUE;
    }
    return FALSE;
}

// ------------------ Entry point ------------------

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int)
{
    if (!EnsureSingleInstance()) return 0;
    if (!RelaunchAsAdminIfNeeded()) return 0;

    g_hInst = hInstance;

    HWND hDlg = CreateDialogParamW(hInstance, MAKEINTRESOURCE(IDD_MAIN), nullptr, MainDlgProc, 0);
    if (!hDlg) return 0;

    // 启动后立即隐藏窗口
    ShowWindow(hDlg, SW_HIDE);

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        if (!IsDialogMessageW(hDlg, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    return (int)msg.wParam;
}
