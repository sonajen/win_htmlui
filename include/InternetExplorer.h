#ifndef INTERNETEXPLORER_H
#define INTERNETEXPLORER_H

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  InternetExplorer.h                                                                            //
//  ------------------                                                                            //
//                                                                                                //
//  Author: Steve Evans, Sonajen Oy                                                               //
//                                                                                                //
//  This file declares the CInternetExplorer class                                                //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

#include "atlbase.h"
#include "atlwin.h"
#include "atlhost.h"
#include "dispex.h"
#include "exdispid.h"
#include "WebBrowser.h"

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

class CInternetExplorer : public CWebBrowser,
                          public CWindowImpl<CInternetExplorer>,
                          public IDispEventImpl<1,CInternetExplorer,&DIID_DWebBrowserEvents2,
                                                &LIBID_SHDocVw,1,0> {

      public:

         // Construction and destruction
         CInternetExplorer () {}
         virtual ~CInternetExplorer () {}

         // Create the browser window
         virtual bool CreateBrowserWindow (HWND hWndParent);

         // Navigate to a URL
         virtual bool Navigate (const std::wstring strURL);

         // Refresh the current page
         virtual void Refresh ();

         // Run the message loop
         virtual int RunMessageLoop (HINSTANCE hInstance);

         // Message handlers
         BEGIN_MSG_MAP(CInternetExplorer)
            MESSAGE_HANDLER (WM_DESTROY,OnDestroy)
            MESSAGE_HANDLER (WM_ERASEBKGND,OnEraseBkgnd)
         END_MSG_MAP()

         BEGIN_SINK_MAP (CInternetExplorer)
            SINK_ENTRY_EX (1,DIID_DWebBrowserEvents2,DISPID_DOCUMENTCOMPLETE,OnDocumentComplete)
         END_SINK_MAP()

      protected:

         // The web browser object
         CComPtr<IWebBrowser2> m_pBrowser;

         // Set Internet Explorer version 11
         void SetVersion11 ();

         // Message handlers
         LRESULT OnDestroy (UINT uMsg,WPARAM wParam,LPARAM lParam,BOOL &bHandled);
         LRESULT OnEraseBkgnd (UINT uMsg,WPARAM wParam,LPARAM lParam,BOOL &bHandled);

         // Event handlers
	     void __stdcall OnDocumentComplete (IDispatch *pDisp,VARIANT *pURL);
};

#endif // INTERNETEXPLORER_H
