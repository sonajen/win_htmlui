////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  main.cpp                                                                                      //
//  --------                                                                                      //
//                                                                                                //
//  Author: Steve Evans, Sonajen Oy                                                               //
//                                                                                                //
//  This file defines the entry point for the application and functions to manage the main window //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

#include "windows.h"
#include "atlbase.h"
#include "resource.h"
#include "string"
#include "InternetExplorer.h"
#include "Edge.h"

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//                                         Configuration                                          //
//                                         -------------                                          //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

#define CONFIG_APP_WINDOW_WIDTH  800
#define CONFIG_APP_WINDOW_HEIGHT 600
#define CONFIG_APP_URL           L"https://www.whatismybrowser.com/"

#define CONFIG_APP_CLASS L"SonajenHTMLUIClass"
#define CONFIG_APP_MUTEX L"SonajenHTMLUIMutex"

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//                                        Global Variables                                        //
//                                        ----------------                                        //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

CComModule _Module;
CWebBrowser *g_pBrowser;

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//                                      Function Prototypes                                       //
//                                      -------------------                                       //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

static HWND CreateMainWindow (HINSTANCE,int,int,HANDLE &);
static void ShowMainWindow (HWND);
static void AdjustWindowSize (HWND,int,int);
LRESULT CALLBACK WndProc (HWND,UINT,WPARAM,LPARAM);

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : ::_tWinMain                                                                        //
//                                                                                                //
//  Access   : Public                                                                             //
//                                                                                                //
//  Arguments: [IN] hInstance     - current application instance handle                           //
//             [IN] hPrevInstance - previous application instance handle                          //
//             [IN] lpCmdLine     - command line arguments                                        //
//             {IN] nCmdShow      - indicates how the window will be shown                        //
//                                                                                                //
//  Returned : the exit code of the process                                                       //
//                                                                                                //
//  Main entry point                                                                              //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

int APIENTRY 
_tWinMain (HINSTANCE hInstance,HINSTANCE hPrevInstance,LPTSTR lpCmdLine,int nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);
    LPWSTR *arrArgList;
    HWND hWnd;
    HANDLE hMutex;
    wchar_t strTitle[64];
    int nArgs,L1;
    enum EBrowserType { eInternetExplorer, eEdge } eBrowserType;

    try {

        // Initialize COM

        ::CoInitializeEx (NULL,COINIT_APARTMENTTHREADED);

        // Get the browser type (default to IE)

        eBrowserType = eInternetExplorer;
        arrArgList   = ::CommandLineToArgvW (::GetCommandLine (),&nArgs);

        for (L1 = 0; L1 < nArgs; L1++) {

            if (::_wcsicmp (arrArgList[L1],L"/edge") == 0) {
               eBrowserType = eEdge;
               break;
            }
        }

        // Create the main window

        hWnd = ::CreateMainWindow (hInstance,CONFIG_APP_WINDOW_WIDTH,CONFIG_APP_WINDOW_HEIGHT,
                                   hMutex);
        if (hWnd == NULL)
           return (0);

        // Set the title

        ::LoadStringW (NULL,IDS_APP_TITLE,strTitle,sizeof (strTitle) / sizeof (wchar_t));
        ::SetWindowText (hWnd,strTitle);

        // Create the browser object

        if (eBrowserType == eInternetExplorer)
           g_pBrowser = new CInternetExplorer;
        else
           g_pBrowser = new CEdge;

        // Create the browser window

        if (g_pBrowser->CreateBrowserWindow (hWnd) == NULL)
           throw 1;

        // Navigate to the URL

        g_pBrowser->Navigate (CONFIG_APP_URL);

        // Show the main window

        ::ShowMainWindow (hWnd);

        // Run the message loop

        throw g_pBrowser->RunMessageLoop (hInstance);
    }

    catch (int nExitCode) {

        // Release the mutex

        ::CloseHandle (hMutex);

        // Free the browser

        delete g_pBrowser;

        // Exit COM

        ::CoUninitialize ();

        // Return the exit code

        return (nExitCode);
    }

    catch (...) {}

    // Something went seriously wrong

    return (-1);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : ::CreateMainWindow                                                                 //
//                                                                                                //
//  Access   : Public                                                                             //
//                                                                                                //
//  Arguments: [IN]  hInstance - current application instance handle                              //
//             [IN]  nWidth    - the width of the window in pixels                                //
//             [IN]  nHeight   - the height of the window in pixels                               //
//             [OUT] hMutex    - the mutex handle of the created window                           //
//                                                                                                //
//  Returned : a window handle, or NULL for failure                                               //
//                                                                                                //
//  Create the main window                                                                        //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

HWND
CreateMainWindow (HINSTANCE hInstance,int nWidth,int nHeight,HANDLE &hMutex)
{
    WNDCLASSEX wndClass;
    HWND hWnd;
    RECT screenRect;
    DWORD dwStyle;
    int nX,nY;

    // Check if another instance of the browser is running and restore the window in that case

    hMutex = ::CreateMutexW (NULL,FALSE,CONFIG_APP_MUTEX);
    if (hMutex == NULL || ::GetLastError () == ERROR_ALREADY_EXISTS) {

       if (hMutex != NULL)
          ::CloseHandle (hMutex);

       if ((hWnd = ::FindWindowExW (NULL,NULL,CONFIG_APP_CLASS,NULL)) != NULL)
          ::SendMessage (hWnd,WM_RESTORE_WINDOW,0,0);

       return (NULL);
    }

    // Register the window class

    wndClass.cbSize        = sizeof (WNDCLASSEX);
    wndClass.style         = CS_HREDRAW | CS_VREDRAW;
    wndClass.lpfnWndProc   = ::WndProc;
    wndClass.cbClsExtra    = 0;
    wndClass.cbWndExtra    = 0;
    wndClass.hInstance     = hInstance;
    wndClass.hIcon         = ::LoadIcon (hInstance,MAKEINTRESOURCE (IDI_ICON));
    wndClass.hCursor       = ::LoadCursor (NULL,IDC_ARROW);
    wndClass.hbrBackground = NULL;
    wndClass.lpszMenuName  = NULL;
    wndClass.lpszClassName = CONFIG_APP_CLASS;
    wndClass.hIconSm       = ::LoadIcon (wndClass.hInstance,MAKEINTRESOURCE (IDI_ICON));

    if (::RegisterClassEx (&wndClass) == FALSE)
       return (NULL);

    // Set the window styles and adjust to the desired client rectangle

    RECT clientRect = { 1, 1, nWidth, nHeight };
    dwStyle = WS_SYSMENU | WS_MINIMIZEBOX | WS_CAPTION;

    ::AdjustWindowRectEx (&clientRect,dwStyle,FALSE,0);
    nWidth  = clientRect.right - clientRect.left + 1;
    nHeight = clientRect.bottom - clientRect.top + 1;

    // Create the application window

    hWnd = ::CreateWindowEx (0,CONFIG_APP_CLASS,L"",dwStyle,0,0,nWidth,nHeight,NULL,NULL,
                             hInstance,NULL);
    if (hWnd == NULL)
       return (NULL);

    // Set the position of the window using the screen area and adjusting for the task bar

    ::SystemParametersInfo (SPI_GETWORKAREA,0,&screenRect,0);
    if ((nX = (screenRect.right - screenRect.left - nWidth) / 2) < 0)
       nX = 0;

    if ((nY = (screenRect.bottom - screenRect.top - nHeight) / 2) < 0)
       nY = 0;

    ::SetWindowPos (hWnd,HWND_TOP,nX,nY,nWidth,nHeight,SWP_FRAMECHANGED);

    // Return the window handle

    return (hWnd);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : ::ShowMainWindow                                                                   //
//                                                                                                //
//  Access   : Private                                                                            //
//                                                                                                //
//  Arguments: [IN] hWnd - the window handle                                                      //
//                                                                                                //
//  Show the main window                                                                          //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

void
ShowMainWindow (HWND hWnd)
{
    HWND hWndTop;

    // Show the window. Resort to thread trickery to force the window to be the foreground window 

    hWndTop = ::GetForegroundWindow ();
    ::AttachThreadInput (::GetWindowThreadProcessId (hWndTop,NULL),::GetCurrentThreadId (),TRUE);

    ::SetForegroundWindow (hWnd);
    ::ShowWindow (hWnd,SW_SHOWNORMAL);
    ::UpdateWindow (hWnd);
    ::SetFocus (hWnd);

    hWndTop = ::GetForegroundWindow ();
    ::AttachThreadInput (::GetWindowThreadProcessId (hWndTop,NULL),::GetCurrentThreadId (),FALSE);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : ::AdjustWindowSize                                                                 //
//                                                                                                //
//  Access   : Private                                                                            //
//                                                                                                //
//  Arguments: [IN] hWnd    - the window handle                                                   //
//             [IN] nWidth  - the current window width                                            //
//             [IN] nHeight - the current window height                                           //
//                                                                                                //
//  Ensure the window has the correct dimensions                                                  //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

void
AdjustWindowSize (HWND hWnd,int nWidth,int nHeight)
{
    RECT screenRect;
    DWORD dwStyle;
    int nX,nY;

    // Do nothing if zero width and height

    if (nWidth == 0 && nHeight == 0)
       return;

    // If the dimensions are different to the expected window size, the ensure the window size is
    // correct. Windows seems to have an issue where if the resolution is set to something smaller
    // than the window, then resetting to a higher resolution causes a maximization effect. We
    // solve that here

    if (nWidth != CONFIG_APP_WINDOW_WIDTH || nHeight != CONFIG_APP_WINDOW_HEIGHT) {

       // Adjust the window rectangle

       RECT clientRect = { 1, 1, CONFIG_APP_WINDOW_WIDTH, CONFIG_APP_WINDOW_HEIGHT };
       dwStyle = WS_SYSMENU | WS_MINIMIZEBOX | WS_CAPTION;

       ::AdjustWindowRectEx (&clientRect,dwStyle,FALSE,0);
       nWidth  = clientRect.right - clientRect.left + 1;
       nHeight = clientRect.bottom - clientRect.top + 1;

       // Check the top left is not off the screen

       ::SystemParametersInfo (SPI_GETWORKAREA,0,&screenRect,0);
       nX = (screenRect.right - screenRect.left - nWidth) / 2;
       if (nX < 0)
          nX = 0;

       nY = (screenRect.bottom - screenRect.top - nHeight) / 2;
       if (nY < 0)
          nY = 0;

       // Set the new window position and dimensions

       ::SetWindowPos (hWnd,0,nX,nY,nWidth,nHeight,SWP_FRAMECHANGED | SWP_NOZORDER);
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : ::WndProc                                                                          //
//                                                                                                //
//  Access   : Private                                                                            //
//                                                                                                //
//  Arguments: [IN] hWnd   - the window handle                                                    //
//             [IN] uMsg   - the message to process                                               //
//             [IN] wParam - additional message information                                       //
//             {IN] lParam - additional message information                                       //
//                                                                                                //
//  Returned : the result of processing the message                                               //
//                                                                                                //
//  Windows message handler                                                                       //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

LRESULT CALLBACK
WndProc (HWND hWnd,UINT uMsg,WPARAM wParam,LPARAM lParam)
{

    switch (uMsg) {

        case WM_DESTROY:

             ::PostQuitMessage (0);
             break;

        case WM_DISPLAYCHANGE:
        case WM_SIZE:

             ::AdjustWindowSize (hWnd,lParam & 0xFFFF,(int) lParam >> 16);
             break;

        case WM_RESTORE_WINDOW:

             g_pBrowser->Refresh ();
             ::ShowWindow (hWnd,SW_SHOWNORMAL);
             ::SetForegroundWindow (hWnd);
             break;

        default:

             return (::DefWindowProc (hWnd,uMsg,wParam,lParam));
    }

    return (0);
}