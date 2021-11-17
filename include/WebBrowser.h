#ifndef WEBBROWSER_H
#define WEBBROWSER_H

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  WebBrowser.h                                                                                  //
//  ------------                                                                                  //
//                                                                                                //
//  Author: Steve Evans, Sonajen Oy                                                               //
//                                                                                                //
//  This file declares the CWebBrowser class                                                      //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

#include "windows.h"
#include "string"

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//                                             Macros                                             //
//                                             ------                                             //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

#define WM_RESTORE_WINDOW (WM_USER + 1)

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//                                        Class Definition                                        //
//                                        ----------------                                        //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

class CWebBrowser {

      public:

         // Create the browser window
         virtual bool CreateBrowserWindow (HWND hWndParent) = 0;

         // Navigate to a URL
         virtual bool Navigate (const std::wstring strURL) = 0;

         // Refresh the current page
         virtual void Refresh () = 0;

         // Run the message loop
         virtual int RunMessageLoop (HINSTANCE hInstance) = 0;
};

#endif // WEBBROWSER_H
