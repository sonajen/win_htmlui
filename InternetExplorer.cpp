////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  InternetExplorer.cpp                                                                          //
//  --------------------                                                                          //
//                                                                                                //
//  Author: Steve Evans, Sonajen Oy                                                               //
//                                                                                                //
//  This file defines the CInternetExplorer class                                                 //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

#include "activscp.h"
#include "comdef.h"
#include "resource.h"
#include "InternetExplorer.h"

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CInternetExplorer::CreateBrowserWindow                                             //
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
CInternetExplorer::CreateBrowserWindow (HWND hWndParent)
{
    CAxWindow *pAxWindow;
    IUnknown *pUnk,*pSinkSite;
    CComPtr<IAxWinHostWindow> spAxWindow;
    RECT rect;
    HRESULT hr;

    // Ensure the window has not been created already
    
    if (m_hWnd != NULL)
       return (false);

    // Set Internet Explorer version 11

    SetVersion11 ();

    // Create the window if necessary

    ::GetClientRect (hWndParent,&rect);
    if (Create (hWndParent,rect,L"",WS_CHILD,0) == NULL)
       return (false);

    // Create the host window

    pAxWindow = new CAxWindow;
    if (pAxWindow->Create (m_hWnd,rect,0,WS_CHILD | WS_VISIBLE,0) == FALSE)
       return (false);

    hr = pAxWindow->QueryHost (&spAxWindow);
    if (SUCCEEDED (hr) == FALSE)
       return (false);

    // Create the browser control

    pUnk = NULL;
    pSinkSite = (IUnknown *) (IDispEventImpl<1,CInternetExplorer,&DIID_DWebBrowserEvents2,
                                             &LIBID_SHDocVw,1,0> *) this;

    hr = spAxWindow->CreateControlEx (_T ("Shell.Explorer"),*pAxWindow,NULL,&pUnk,
                                      DIID_DWebBrowserEvents2,pSinkSite);
    if (SUCCEEDED (hr) == FALSE)
       return (false);

    pAxWindow->QueryControl (IID_IWebBrowser2,(void **) &m_pBrowser);

    // Disable script error popups which show up on some web sites when using the browser control

    m_pBrowser->put_Silent (VARIANT_TRUE);

    // The browser window was created successfully

    return (true);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CInternetExplorer::Navigate                                                        //
//                                                                                                //
//  Access   : Public                                                                             //
//                                                                                                //
//  Arguments: [IN] strURL - the URL                                                              //
//                                                                                                //
//  Returned : true if the operation completed successfully, else false                           //
//                                                                                                //
//  Navigate to a URL                                                                             //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

bool
CInternetExplorer::Navigate (const std::wstring strURL)
{
    CComVariant vtNull;
    CComBSTR bstrURL;
    HRESULT hr;

    // Check the browser control has been created

    if (m_pBrowser == NULL)
       return (false);

    // Validate arguments

    if (strURL.length () == 0)
       return (false);

    // Attempt to navigate to the URL

    bstrURL = strURL.c_str ();
    hr = m_pBrowser->Navigate (bstrURL,&vtNull,&vtNull,&vtNull,&vtNull);
    if (SUCCEEDED (hr) == TRUE) {
       Sleep (100);
       return (true);
    } else
       return (false); 
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CInternetExplorer::Refresh                                                         //
//                                                                                                //
//  Access   : Public                                                                             //
//                                                                                                //
//  Arguments: None                                                                               //
//                                                                                                //
//  Refresh the current page                                                                      //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

void
CInternetExplorer::Refresh ()
{
    VARIANT vtValue;
    HRESULT hr;
    int L1;

    ::VariantInit (&vtValue);
    vtValue.vt   = VT_I4;
    vtValue.lVal = REFRESH_COMPLETELY;

    for (L1 = 0; L1 < 10; L1++) {

        hr = m_pBrowser->Refresh2 (&vtValue);
        if (SUCCEEDED (hr) == TRUE)
           break;

        ::Sleep (1000);
    }

    ::VariantClear (&vtValue);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CInternetExplorer::RunMessageLoop                                                  //
//                                                                                                //
//  Access   : Public                                                                             //
//                                                                                                //
//  Arguments: [IN] hInstance - the application instance handle                                   //
//                                                                                                //
//  Returned : the exit code                                                                      //
//                                                                                                //
//  Run the message loop                                                                          //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

int
CInternetExplorer::RunMessageLoop (HINSTANCE hInstance)
{
    IOleInPlaceActiveObject *pOleInPlaceActiveObject;
    HACCEL hAccelTable;
    HRESULT hr;
    MSG msg;

    // Load accelerators

    hAccelTable = ::LoadAccelerators (hInstance,MAKEINTRESOURCE (IDR_ACCELERATOR));

    // Get the in place active object to handle key down events

    pOleInPlaceActiveObject = NULL;
    m_pBrowser->QueryInterface (IID_IOleInPlaceActiveObject,(void **) &pOleInPlaceActiveObject);

    // Enter the message loop

    while (::GetMessage (&msg,NULL,0,0) != FALSE) {

       if (msg.message == WM_KEYDOWN && pOleInPlaceActiveObject != NULL) {
          hr = pOleInPlaceActiveObject->TranslateAccelerator (&msg);
          if (hr == S_OK)
             continue;
       }

       if (::TranslateAccelerator (msg.hwnd,hAccelTable,&msg) == FALSE) {
          ::TranslateMessage (&msg);
          ::DispatchMessage (&msg);
       }
    }

    // Free the in place active object

    pOleInPlaceActiveObject->Release ();

    // Return the exit code

    return ((int) msg.wParam);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CInternetExplorer::SetVersion11                                                    //
//                                                                                                //
//  Access   : Protected                                                                          //
//                                                                                                //
//  Arguments: None                                                                               //
//                                                                                                //
//  Set Internet Explorer version 11                                                              //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

void
CInternetExplorer::SetVersion11 ()
{
    HKEY hKey;
    std::wstring strValue;
    wchar_t strFileName[1024],strExecutableName[_MAX_FNAME],strExtension[_MAX_EXT];
    DWORD dwLen,dwValue;

    // Get the executable name

    dwLen = ::GetModuleFileNameW (NULL,strFileName,sizeof (strFileName) / sizeof (wchar_t));
    strFileName[dwLen] = L'\0';
    ::_wsplitpath_s (strFileName,NULL,0,NULL,0,strExecutableName,_MAX_FNAME,strExtension,_MAX_EXT);

    // Open the registry key

    hKey = NULL;
    strValue = L"SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\" \
               L"FEATURE_BROWSER_EMULATION";

    if (::RegOpenKeyExW (HKEY_CURRENT_USER,strValue.c_str (),0,KEY_WRITE | KEY_WOW64_64KEY,
                         &hKey) == ERROR_SUCCESS) {

       // Set the entry for Internet Explorer 11 emulation

       strValue = std::wstring (strExecutableName) + strExtension;
       dwValue  = 11001;
       ::RegSetValueExW (hKey,strValue.c_str (),NULL,REG_DWORD,(BYTE *) &dwValue,sizeof (DWORD));

       // Close the key

       ::RegCloseKey (hKey);
    }
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CInternetExplorer::OnDestroy                                                       //
//                                                                                                //
//  Access   : Protected                                                                          //
//                                                                                                //
//  Arguments: [IN]  uMsg     - the message sent to the window                                    //
//             [IN]  wParam   - additional message specific information                           //
//             [IN]  lParam   - additional message specific information                           //
//             [OUT] bHandled - TRUE if the message was handled, else FALSE                       //
//                                                                                                //
//  Returned : result of processing the message (0 = successful)                                  //
//                                                                                                //
//  WM_DESTROY message handler                                                                    //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

LRESULT
CInternetExplorer::OnDestroy (UINT uMsg,WPARAM wParam,LPARAM lParam,BOOL &bHandled)
{

    // Pass the message onto the parent window

    GetParent ().SendMessage (WM_DESTROY,0,0);
    return (0);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CInternetExplorer::OnEraseBkgnd                                                    //
//                                                                                                //
//  Access   : Protected                                                                          //
//                                                                                                //
//  Arguments: [IN]  uMsg     - the message sent to the window                                    //
//             [IN]  wParam   - additional message specific information                           //
//             [IN]  lParam   - additional message specific information                           //
//             [OUT] bHandled - TRUE if the message was handled, else FALSE                       //
//                                                                                                //
//  Returned : result of processing the message (0 = successful)                                  //
//                                                                                                //
//  WM_ERASEBKGND message handler                                                                 //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

LRESULT
CInternetExplorer::OnEraseBkgnd (UINT uMsg,WPARAM wParam,LPARAM lParam,BOOL &bHandled)
{

    // Ignore this message to prevent screen flicker

    return (0);
}

////////////////////////////////////////////////////////////////////////////////////////////////////
//                                                                                                //
//  Function : CInternetExplorer::OnDocumentComplete                                              //
//                                                                                                //
//  Access   : Protected                                                                          //
//                                                                                                //
//  Arguments: [IN] pDisp - a pointer to the IDispatch interface of the window or frame in which  //
//                          the document is loaded                                                //
//             [IN] pURL  - a pointer that specifies the URL, Universal Naming Convention file    //
//                          name, or a pointer to an item identifier list of the loaded document  //
//                                                                                                //
//  DWebBrowserEvents2::DocumentComplete event handler                                            //
//                                                                                                //
////////////////////////////////////////////////////////////////////////////////////////////////////

void 
CInternetExplorer::OnDocumentComplete (IDispatch *pDisp,VARIANT *pURL)
{

    // Show the window

    ShowWindow (SW_SHOW);
}
