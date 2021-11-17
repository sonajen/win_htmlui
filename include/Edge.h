#ifndef EDGE_H
#define EDGE_H

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Edge.h                                                                                        //
//  ------                                                                                        //
//                                                                                                //
//  Author: Steve Evans, Sonajen Oy                                                               //
//                                                                                                //
//  This file declares the CEdge class                                                            //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

#include "wrl.h"
#include "wil/com.h"
#include "WebBrowser.h"
#include "WebView2.h"

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//                                        Class Definition                                        //
//                                        ----------------                                        //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

class CEdge : public CWebBrowser {

      public:

         // Construction and destruction
         CEdge () {}
         virtual ~CEdge () {}

         // Create the browser window
         virtual bool CreateBrowserWindow (HWND hWndParent);

         // Navigate to a URL
         virtual bool Navigate (const std::wstring strURL);

         // Refresh the current page
         virtual void Refresh ();

         // Run the message loop
         virtual int RunMessageLoop (HINSTANCE hInstance);

      protected:

         // The parent window
         HWND m_hWndParent;

         // The URL prior to window creation
         std::wstring m_strURL;

         // The web view controller
         wil::com_ptr<ICoreWebView2Controller> m_webviewController;

         // The web view window
         wil::com_ptr<ICoreWebView2> m_webviewWindow;

         // Create environment completed callback
         HRESULT OnCreateEnvironmentCompleted (HRESULT hr,ICoreWebView2Environment *pEnv);

         // Create WebView2 controller callback
         HRESULT OnCreateCoreWebView2ControllerCompleted (HRESULT result,
                                                          ICoreWebView2Controller *pObjontroller);
};

#endif // EDGE_H
