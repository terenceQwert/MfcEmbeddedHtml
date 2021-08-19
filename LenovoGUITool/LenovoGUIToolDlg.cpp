
// LenovoGUIToolDlg.cpp : implementation file
//

#include "pch.h"
#include "framework.h"
#include "LenovoGUITool.h"
#include "LenovoGUIToolDlg.h"
#include "afxdialogex.h"

///
/// reference https://cpp.hotexamples.com/examples/-/-/GetElementInterface/cpp-getelementinterface-function-examples.html
/// 
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CLenovoGUIToolDlg dialog

BEGIN_DHTML_EVENT_MAP(CLenovoGUIToolDlg)
	DHTML_EVENT_ONCLICK(_T("ButtonOK"), OnButtonOK)
	DHTML_EVENT_ONCLICK(_T("ButtonCancel"), OnButtonCancel)
	DHTML_EVENT_ONCLICK(_T("CheckLink"), OnCheckClick)
END_DHTML_EVENT_MAP()


BEGIN_DISPATCH_MAP(CLenovoGUIToolDlg, CDHtmlDialog)
	DISP_FUNCTION(CLenovoGUIToolDlg, "SayHello", func, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()


void 
CLenovoGUIToolDlg::func()
{
	MessageBox(_T("SayHello func", _T("Hello~~~")));
}

CLenovoGUIToolDlg::CLenovoGUIToolDlg(CWnd* pParent /*=nullptr*/)
	: CDHtmlDialog(IDD_LENOVOGUITOOL_DIALOG, IDR_HTML_LENOVOGUITOOL_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CLenovoGUIToolDlg::DoDataExchange(CDataExchange* pDX)
{
	CDHtmlDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CLenovoGUIToolDlg, CDHtmlDialog)
END_MESSAGE_MAP()


// CLenovoGUIToolDlg message handlers

BOOL CLenovoGUIToolDlg::OnInitDialog()
{
	CDHtmlDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	m_LinkActive = true;
	// TODO: Add extra initialization here
	EnableAutomation();
	SetExternalDispatch(GetIDispatch(TRUE));	//	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CLenovoGUIToolDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDHtmlDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CLenovoGUIToolDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

HRESULT CLenovoGUIToolDlg::OnButtonOK(IHTMLElement* /*pElement*/)
{

	/// <summary>
	///  Get comboBox data
	/// </summary>
	/// <param name=""></param>
	/// <returns></returns>
	CComPtr<IHTMLElement> spPoint;
	HRESULT hr = S_OK;
	BSTR	bStr;
	CString csAttr = _T("value");
	CString csId = _T("cars");
	hr = GetElementInterface(csId, &spPoint);
	if (S_OK == hr)
	{
		bStr = csAttr.AllocSysString();
		VARIANT	v;
		spPoint->getAttribute(bStr, 0, &v);
		CString csShow(v.bstrVal);
		MessageBox(csShow);
		::VariantClear(&v);
		::SysFreeString(bStr);
	}
	OnOK();
	return S_OK;
}

HRESULT CLenovoGUIToolDlg::OnButtonCancel(IHTMLElement* /*pElement*/)
{
	OnCancel();
	return S_OK;
}

HRESULT CLenovoGUIToolDlg::OnCheckClick(IHTMLElement* pElement)
{
	m_LinkActive = !m_LinkActive;

	IHTMLElement* pLinkElement = NULL;
	if (GetElement(_T("LinkCP"), &pLinkElement) == S_OK && pLinkElement != NULL)
	{
		if (m_LinkActive)
			pLinkElement->put_outerHTML(_bstr_t("<a ID=LinkCP target=_blank href='http://www.codeproject.com'>here</a>"));
		else
			pLinkElement->put_outerHTML(_bstr_t("<font ID=LinkCP color='#COCOCO'>here</font>"));
	}




	return S_OK;
}

void 
CLenovoGUIToolDlg::OnBnClickedButton1()
{
	IHTMLDocument2 *pDoc = NULL;
	CDHtmlDialog::GetDHtmlDocument(&pDoc);
	DISPPARAMS param = { 0 };
	VARIANT VtRet;
	calljsfunction(pDoc, _T("Cppcalljsfunc"), param, &VtRet, NULL, NULL);
	
}

HRESULT CLenovoGUIToolDlg::calljsfunction(
	IHTMLDocument2* pDoc2,
	CString strfunctionname,
	DISPPARAMS  Dispparams,
	VARIANT* varresult,
	EXCEPINFO* exceptinfo, UINT* nargerr) 
{
	IDispatch* pdispscript = NULL;
	HRESULT hResult;
	hResult = pDoc2->get_scripts((IHTMLElementCollection**)&pdispscript);
	if (FAILED(hResult)) {
		return S_FALSE;
	} DISPID DISPID;
	CComBSTR objbstrvalue = strfunctionname;
	BSTR Bstrvalue = objbstrvalue.Copy();
	OLECHAR* pszfunct = Bstrvalue;
	hResult = pdispscript->GetIDsOfNames(IID_NULL, &pszfunct, 1, LOCALE_SYSTEM_DEFAULT, &DISPID);

	if (FAILED(hResult)) {
		pdispscript->Release();
		return hResult;
	}
//	varresult - &GT; 
	varresult->vt = VT_VARIANT;
	hResult = pdispscript->Invoke(DISPID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,&Dispparams, 
		varresult, exceptinfo, nargerr);

	pdispscript->Release();
	return hResult;
}