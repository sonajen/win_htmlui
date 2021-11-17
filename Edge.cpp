////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Edge.cpp                                                                                      //
//  --------                                                                                      //
//                                                                                                //
//  Author: Steve Evans, Sonajen Oy                                                               //
//                                                                                                //
//  This file defines the CEdge class                                                             //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

#include "Edge.h"

using namespace Microsoft::WRL;

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CEdge::CreateBrowserWindow                                                         //
//                                                                                                //
//  Access   : Public                                                                             //
//                                                                                                //
//  Arguments: [IN] hWndParent - the parent window handle                                         //
//                                                                                                //
//  Returned : true if the windows was created, else false                                        //
//                                                                                                //
//  Create the browser window                                                                     //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

bool
CEdge::CreateBrowserWindow (HWND hWndParent)
{
    HRESULT hr;

    // Store the parent window handle

    m_hWndParent = hWndParent;

    // Create the environment which, if it succeeds, will then create the WebView controller and
    // window

    auto fnCallback { Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>
                               (this,&CEdge::OnCreateEnvironmentCompleted) };

    hr = ::CreateCoreWebView2EnvironmentWithOptions (nullptr,nullptr,nullptr,fnCallback.Get ());
    return (SUCCEEDED (hr) == TRUE);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CEdge::Navigate                                                                    //
//                                                                                                //
//  Access   : Public                                                                             //
//                                                                                                //
//  Arguments: [IN] pStrURL - the URL                                                             //
//                                                                                                //
//  Returned : true if the operation completed successfully, else false                           //
//                                                                                                //
//  Navigate to a URL                                                                             //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

bool
CEdge::Navigate (const std::wstring strURL)
{
    HRESULT hr;

    // Validate arguments

    if (strURL.length () == 0)
       return (false);

    // If the WebView window has not been created yet, then save the URL for later, otherwise
    // navigate immediately

    if (m_webviewWindow == nullptr) {
       m_strURL = strURL;
       return (true);
    }

    hr = m_webviewWindow->Navigate (strURL.c_str ());
    return (SUCCEEDED (hr) == TRUE);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CEdge::Refresh                                                                     //
//                                                                                                //
//  Access   : Public                                                                             //
//                                                                                                //
//  Arguments: None                                                                               //
//                                                                                                //
//  Refresh the current page                                                                      //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

void
CEdge::Refresh ()
{

    // Reload the URL

    if (m_webviewWindow != nullptr)
       m_webviewWindow->Reload ();
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CEdge::RunMessageLoop                                                              //
//                                                                                                //
//  Access   : Public                                                                             //
//                                                                                                //
//  Arguments: [IN] hInstance - the application instance handle                                   //
//                                                                                                //
//  Run the message loop                                                                          //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

int
CEdge::RunMessageLoop (HINSTANCE hInstance)
{
    MSG msg;

    // Enter the message loop

    while (::GetMessage (&msg,NULL,0,0) != FALSE) {
       ::TranslateMessage (&msg);
       ::DispatchMessage (&msg);
    }

    // Return the exit code

    return ((int) msg.wParam);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CEdge::OnCreateEnvironmentCompleted                                                //
//                                                                                                //
//  Access   : Protected                                                                          //
//                                                                                                //
//  Arguments: [IN] hr      - the result of creating the web environment                          //
//             [IN] pObjEnv - the environment                                                     //
//                                                                                                //
//  Returned : the result of creating the environment or WebView controller                       //
//                                                                                                //
//  CreateCoreWebView2EnvironmentWithOptions completed callback                                   //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

HRESULT
CEdge::OnCreateEnvironmentCompleted (HRESULT hr,ICoreWebView2Environment *pObjEnv)
{

    // Do nothing if we had a prior failure

    if (FAILED (hr) == TRUE)
       return (hr);

    // Create the WebView controller

    auto fnCallback { Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>
                               (this,&CEdge::OnCreateCoreWebView2ControllerCompleted) };
   
    return (pObjEnv->CreateCoreWebView2Controller (m_hWndParent,fnCallback.Get ()));
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CEdge::OnCreateEnvironmentCompleted                                                //
//                                                                                                //
//  Access   : Protected                                                                          //
//                                                                                                //
//  Arguments: [IN] hr             - the result of creating the web controller                    //
//             [IN] pObjController - the WebView controller                                       //
//                                                                                                //
//  Returned : the result of creating the WebView controller (if it failed), page navigation, or  //
//             S_OK on success                                                                    //
//                                                                                                //
//  CreateCoreWebView2Controller completed callback                                               //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

HRESULT
CEdge::OnCreateCoreWebView2ControllerCompleted (HRESULT hr,
                                                ICoreWebView2Controller *pObjController)
{
    RECT rect;

    // Do nothing if we had a prior failure

    if (FAILED (hr) == TRUE)
       return (hr);

    // Get the controller and window

    m_webviewController = pObjController;
    m_webviewController->get_CoreWebView2 (&m_webviewWindow);

    // Resize the WebView to the bounds of the parent window

    ::GetClientRect (m_hWndParent,&rect);
    m_webviewController->put_Bounds (rect);

    // If a URL has already been set, then navigate to it now

    if (m_strURL.length () != 0) 
       return (m_webviewWindow->Navigate (m_strURL.c_str ()));

    return (S_OK);
}