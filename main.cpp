#define _CRT_SECURE_NO_WARNINGS
#define _ITERATOR_DEBUG_LEVEL 0
#define _HAS_ITERATOR_DEBUGGING 0

#include <Windows.h>
#include <commdlg.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <cstdlib>
#include <cwchar>
#include <shlwapi.h>  // 包含PathFileExistsW

#pragma comment(lib, "Shlwapi.lib")  // 链接Shlwapi库

// LibreOffice转换封装类
class OfficeConverter {
public:
    static bool ConvertFile(const std::wstring& inputPath,
        const std::wstring& outputPath,
        const std::wstring& format)
    {
        // 获取LibreOffice安装路径
        std::wstring sofficePath = GetLibreOfficePath();
        if (sofficePath.empty()) {
            return false;
        }

        // 构建转换命令（兼容5.1.1版本）
        std::wstring cmd = L"\"" + sofficePath + L"\" -headless -convert-to "
            + format + L" \"" + inputPath + L"\" -outdir \""
            + outputPath.substr(0, outputPath.find_last_of(L'\\')) + L"\"";

        // 执行转换命令
        int ret = _wsystem(cmd.c_str());
        return (ret == 0);
    }

    static std::wstring GetLibreOfficePath() {
        // 尝试从注册表获取安装路径
        HKEY hKey;
        wchar_t path[MAX_PATH] = { 0 };
        DWORD bufSize = sizeof(path);

        // 检查64位系统上的注册表路径
        if (RegOpenKeyExW(HKEY_LOCAL_MACHINE,
            L"SOFTWARE\\Wow6432Node\\LibreOffice\\UNO\\InstallPath",
            0, KEY_READ, &hKey) == ERROR_SUCCESS)
        {
            if (RegQueryValueExW(hKey, L"", NULL, NULL, (LPBYTE)path, &bufSize) == ERROR_SUCCESS) {
                RegCloseKey(hKey);
                std::wstring fullPath = std::wstring(path) + L"\\program\\soffice.exe";
                if (PathFileExistsW(fullPath.c_str())) {
                    return fullPath;
                }
            }
            RegCloseKey(hKey);
        }

        // 检查32位系统上的注册表路径
        if (RegOpenKeyExW(HKEY_LOCAL_MACHINE,
            L"SOFTWARE\\LibreOffice\\UNO\\InstallPath",
            0, KEY_READ, &hKey) == ERROR_SUCCESS)
        {
            bufSize = sizeof(path);
            if (RegQueryValueExW(hKey, L"", NULL, NULL, (LPBYTE)path, &bufSize) == ERROR_SUCCESS) {
                RegCloseKey(hKey);
                std::wstring fullPath = std::wstring(path) + L"\\program\\soffice.exe";
                if (PathFileExistsW(fullPath.c_str())) {
                    return fullPath;
                }
            }
            RegCloseKey(hKey);
        }

        // 检查默认安装路径
        const wchar_t* defaultPaths[] = {
            L"C:\\Program Files\\LibreOffice 5\\program\\soffice.exe",
            L"C:\\Program Files (x86)\\LibreOffice 5\\program\\soffice.exe",
            L"C:\\Program Files\\LibreOffice\\program\\soffice.exe",
            L"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe"
        };

        for (int i = 0; i < sizeof(defaultPaths) / sizeof(defaultPaths[0]); ++i) {
            if (PathFileExistsW(defaultPaths[i])) {
                return defaultPaths[i];
            }
        }

        return L"";
    }
};

// GUI主窗口类
class ConverterWindow {
private:
    HWND hwnd;
    HWND hInputEdit, hOutputEdit, hFormatCombo;
    HWND hConvertBtn, hInputBtn, hOutputBtn;
    HWND hStatusText;

    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
        ConverterWindow* pThis = nullptr;

        if (msg == WM_NCCREATE) {
            CREATESTRUCT* pCreate = reinterpret_cast<CREATESTRUCT*>(lParam);
            pThis = reinterpret_cast<ConverterWindow*>(pCreate->lpCreateParams);
            SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pThis));
            pThis->hwnd = hwnd;
        }
        else {
            pThis = reinterpret_cast<ConverterWindow*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
        }

        if (pThis) {
            return pThis->HandleMessage(msg, wParam, lParam);
        }

        return DefWindowProc(hwnd, msg, wParam, lParam);
    }

    LRESULT HandleMessage(UINT msg, WPARAM wParam, LPARAM lParam) {
        switch (msg) {
        case WM_COMMAND: {
            if (reinterpret_cast<HWND>(lParam) == hConvertBtn) {
                OnConvert();
            }
            else if (reinterpret_cast<HWND>(lParam) == hInputBtn) {
                OnSelectInput();
            }
            else if (reinterpret_cast<HWND>(lParam) == hOutputBtn) {
                OnSelectOutput();
            }
            break;
        }
        case WM_CLOSE:
            DestroyWindow(hwnd);
            break;
        case WM_DESTROY:
            PostQuitMessage(0);
            break;
        default:
            return DefWindowProc(hwnd, msg, wParam, lParam);
        }
        return 0;
    }

    void OnSelectInput() {
        OPENFILENAME ofn;
        wchar_t szFile[MAX_PATH] = L"";

        ZeroMemory(&ofn, sizeof(ofn));
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner = hwnd;
        ofn.lpstrFile = szFile;
        ofn.nMaxFile = sizeof(szFile) / sizeof(wchar_t);
        ofn.lpstrFilter = L"所有支持的文件\0*.pdf;*.doc;*.docx;*.odt;*.ppt;*.pptx;*.odp;*.xls;*.xlsx;*.ods\0"
            L"PDF文件\0*.pdf\0"
            L"Word文档\0*.doc;*.docx;*.odt\0"
            L"演示文稿\0*.ppt;*.pptx;*.odp\0"
            L"电子表格\0*.xls;*.xlsx;*.ods\0"
            L"所有文件\0*.*\0";
        ofn.nFilterIndex = 1;
        ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_EXPLORER;

        if (GetOpenFileNameW(&ofn)) {
            SetWindowTextW(hInputEdit, szFile);

            // 自动生成输出文件名
            std::wstring output = szFile;
            size_t pos = output.find_last_of(L'.');
            if (pos != std::wstring::npos) {
                output = output.substr(0, pos);
            }

            // 获取当前选择的格式
            wchar_t format[10] = { 0 };
            GetWindowTextW(hFormatCombo, format, 10);
            if (wcslen(format) > 0) {
                output += L".";
                output += format;
                SetWindowTextW(hOutputEdit, output.c_str());
            }
        }
    }

    void OnSelectOutput() {
        wchar_t szFile[MAX_PATH] = L"";
        GetWindowTextW(hOutputEdit, szFile, MAX_PATH);

        OPENFILENAME ofn;
        ZeroMemory(&ofn, sizeof(ofn));
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner = hwnd;
        ofn.lpstrFile = szFile;
        ofn.nMaxFile = sizeof(szFile) / sizeof(wchar_t);
        ofn.lpstrFilter = L"Word文档\0*.docx\0"
            L"PDF文件\0*.pdf\0"
            L"OpenDocument文本\0*.odt\0"
            L"OpenDocument演示\0*.odp\0"
            L"OpenDocument表格\0*.ods\0";
        ofn.nFilterIndex = 1;
        ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT | OFN_EXPLORER;

        if (GetSaveFileNameW(&ofn)) {
            SetWindowTextW(hOutputEdit, szFile);

            // 根据文件扩展名设置格式
            std::wstring output = szFile;
            size_t pos = output.find_last_of(L'.');
            if (pos != std::wstring::npos) {
                std::wstring ext = output.substr(pos + 1);
                if (ext == L"docx") {
                    SendMessageW(hFormatCombo, CB_SETCURSEL, 0, 0);
                }
                else if (ext == L"pdf") {
                    SendMessageW(hFormatCombo, CB_SETCURSEL, 1, 0);
                }
                else if (ext == L"odt") {
                    SendMessageW(hFormatCombo, CB_SETCURSEL, 2, 0);
                }
                else if (ext == L"odp") {
                    SendMessageW(hFormatCombo, CB_SETCURSEL, 3, 0);
                }
                else if (ext == L"ods") {
                    SendMessageW(hFormatCombo, CB_SETCURSEL, 4, 0);
                }
            }
        }
    }

    void OnConvert() {
        wchar_t inputPath[MAX_PATH] = { 0 };
        wchar_t outputPath[MAX_PATH] = { 0 };
        wchar_t format[10] = { 0 };
        GetWindowTextW(hInputEdit, inputPath, MAX_PATH);
        GetWindowTextW(hOutputEdit, outputPath, MAX_PATH);
        GetWindowTextW(hFormatCombo, format, 10);

        if (lstrlenW(inputPath) == 0 || lstrlenW(outputPath) == 0 || lstrlenW(format) == 0) {
            MessageBoxW(hwnd, L"请选择输入文件、输出文件和转换格式", L"错误", MB_ICONERROR);
            return;
        }

        // 显示转换状态
        SetWindowTextW(hStatusText, L"转换中，请稍候...");
        UpdateWindow(hStatusText);

        // 执行转换
        bool success = OfficeConverter::ConvertFile(inputPath, outputPath, format);

        if (success) {
            SetWindowTextW(hStatusText, L"转换成功!");
            MessageBoxW(hwnd, L"文件转换完成!", L"成功", MB_ICONINFORMATION);
        }
        else {
            SetWindowTextW(hStatusText, L"转换失败!");

            // 检查LibreOffice是否安装
            if (OfficeConverter::GetLibreOfficePath().empty()) {
                MessageBoxW(hwnd,
                    L"未找到LibreOffice 5.1.1安装路径\n请确保已正确安装LibreOffice",
                    L"错误",
                    MB_ICONERROR);
            }
            else {
                MessageBoxW(hwnd,
                    L"文件转换失败，可能原因:\n"
                    L"1. 文件格式不支持\n"
                    L"2. 文件被其他程序占用\n"
                    L"3. 输出路径无写入权限",
                    L"错误",
                    MB_ICONERROR);
            }
        }
    }

public:
    ConverterWindow(HINSTANCE hInstance) : hwnd(nullptr) {
        // 注册窗口类
        WNDCLASSW wc = {};
        wc.lpfnWndProc = WndProc;
        wc.hInstance = hInstance;
        wc.lpszClassName = L"FileConverterWindow";
        wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        RegisterClassW(&wc);

        // 创建窗口
        hwnd = CreateWindowW(
            wc.lpszClassName,
            L"办公文件格式转换器 (兼容LibreOffice 5.1.1)",
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
            CW_USEDEFAULT, CW_USEDEFAULT, 550, 300,
            nullptr, nullptr, hInstance, this
        );

        if (!hwnd) return;

        // 创建控件
        CreateWindowW(L"static", L"输入文件:", WS_VISIBLE | WS_CHILD,
            20, 20, 80, 25, hwnd, nullptr, hInstance, nullptr);
        hInputEdit = CreateWindowW(L"edit", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
            100, 20, 300, 25, hwnd, nullptr, hInstance, nullptr);
        hInputBtn = CreateWindowW(L"button", L"浏览...", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            410, 20, 80, 25, hwnd, (HMENU)1, hInstance, nullptr);

        CreateWindowW(L"static", L"输出文件:", WS_VISIBLE | WS_CHILD,
            20, 60, 80, 25, hwnd, nullptr, hInstance, nullptr);
        hOutputEdit = CreateWindowW(L"edit", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
            100, 60, 300, 25, hwnd, nullptr, hInstance, nullptr);
        hOutputBtn = CreateWindowW(L"button", L"浏览...", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            410, 60, 80, 25, hwnd, (HMENU)2, hInstance, nullptr);

        CreateWindowW(L"static", L"转换格式:", WS_VISIBLE | WS_CHILD,
            20, 100, 80, 25, hwnd, nullptr, hInstance, nullptr);
        hFormatCombo = CreateWindowW(L"combobox", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | CBS_DROPDOWNLIST,
            100, 100, 150, 150, hwnd, (HMENU)3, hInstance, nullptr);

        // 添加格式选项
        SendMessageW(hFormatCombo, CB_ADDSTRING, 0, (LPARAM)L"docx");
        SendMessageW(hFormatCombo, CB_ADDSTRING, 0, (LPARAM)L"pdf");
        SendMessageW(hFormatCombo, CB_ADDSTRING, 0, (LPARAM)L"odt");
        SendMessageW(hFormatCombo, CB_ADDSTRING, 0, (LPARAM)L"odp");
        SendMessageW(hFormatCombo, CB_ADDSTRING, 0, (LPARAM)L"ods");
        SendMessageW(hFormatCombo, CB_SETCURSEL, 0, 0); // 默认选择docx

        hConvertBtn = CreateWindowW(L"button", L"开始转换", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
            200, 140, 150, 40, hwnd, (HMENU)4, hInstance, nullptr);

        hStatusText = CreateWindowW(L"static", L"准备就绪", WS_VISIBLE | WS_CHILD | SS_LEFT,
            20, 190, 500, 30, hwnd, (HMENU)5, hInstance, nullptr);

        ShowWindow(hwnd, SW_SHOW);
        UpdateWindow(hwnd);
    }

    void Run() {
        MSG msg = {};
        while (GetMessage(&msg, nullptr, 0, 0)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
};

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    ConverterWindow window(hInstance);
    window.Run();
    return 0;
}
