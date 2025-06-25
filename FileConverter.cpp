锘?/ FileConverterFixed.cpp : 瀹氫箟搴旂敤绋嬪簭鐨勫叆鍙ｇ偣銆?
//

#include "framework.h"
#include "FileConverterFixed.h"

#define MAX_LOADSTRING 100

// 鍏ㄥ眬鍙橀噺:
HINSTANCE hInst;                                // 褰撳墠瀹炰緥
WCHAR szTitle[MAX_LOADSTRING];                  // 鏍囬鏍忔枃鏈?
WCHAR szWindowClass[MAX_LOADSTRING];            // 涓荤獥鍙ｇ被鍚?

// 姝や唬鐮佹ā鍧椾腑鍖呭惈鐨勫嚱鏁扮殑鍓嶅悜澹版槑:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: 鍦ㄦ澶勬斁缃唬鐮併€?

    // 鍒濆鍖栧叏灞€瀛楃涓?
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_FILECONVERTERFIXED, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // 鎵ц搴旂敤绋嬪簭鍒濆鍖?
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_FILECONVERTERFIXED));

    MSG msg;

    // 涓绘秷鎭惊鐜?
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int) msg.wParam;
}



//
//  鍑芥暟: MyRegisterClass()
//
//  鐩爣: 娉ㄥ唽绐楀彛绫汇€?
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_FILECONVERTERFIXED));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_FILECONVERTERFIXED);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   鍑芥暟: InitInstance(HINSTANCE, int)
//
//   鐩爣: 淇濆瓨瀹炰緥鍙ユ焺骞跺垱寤轰富绐楀彛
//
//   娉ㄩ噴:
//
//        鍦ㄦ鍑芥暟涓紝鎴戜滑鍦ㄥ叏灞€鍙橀噺涓繚瀛樺疄渚嬪彞鏌勫苟
//        鍒涘缓鍜屾樉绀轰富绋嬪簭绐楀彛銆?
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // 灏嗗疄渚嬪彞鏌勫瓨鍌ㄥ湪鍏ㄥ眬鍙橀噺涓?

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  鍑芥暟: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  鐩爣: 澶勭悊涓荤獥鍙ｇ殑娑堟伅銆?
//
//  WM_COMMAND  - 澶勭悊搴旂敤绋嬪簭鑿滃崟
//  WM_PAINT    - 缁樺埗涓荤獥鍙?
//  WM_DESTROY  - 鍙戦€侀€€鍑烘秷鎭苟杩斿洖
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // 鍒嗘瀽鑿滃崟閫夋嫨:
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: 鍦ㄦ澶勬坊鍔犱娇鐢?hdc 鐨勪换浣曠粯鍥句唬鐮?..
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// 鈥滃叧浜庘€濇鐨勬秷鎭鐞嗙▼搴忋€?
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
