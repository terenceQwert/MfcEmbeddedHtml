
// LenovoGUIToolDlg.h : header file
//

#pragma once


// CLenovoGUIToolDlg dialog
class CLenovoGUIToolDlg : public CDHtmlDialog
{
// Construction
public:
	CLenovoGUIToolDlg(CWnd* pParent = nullptr);	// standard constructor

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_LENOVOGUITOOL_DIALOG, IDH = IDR_HTML_LENOVOGUITOOL_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

	HRESULT OnButtonOK(IHTMLElement *pElement);
	HRESULT OnButtonCancel(IHTMLElement *pElement);
	HRESULT OnCheckClick(IHTMLElement* pElement);

// Implementation
protected:
	HICON m_hIcon;
	bool  m_LinkActive;
	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
	DECLARE_DHTML_EVENT_MAP()
	DECLARE_DISPATCH_MAP()

	HRESULT calljsfunction(IHTMLDocument2* pDoc2,
		CString strfunctionname,
		DISPPARAMS dispparams,
		VARIANT* VarResult,
		EXCEPINFO* exceptInfo,
		UINT* nargErr);

		void func();
public:
	afx_msg void OnBnClickedButton1();

	/// <summary>
	/// Override this fucntion to skip web acxtiveX security checking.
	/// </summary>
	/// <returns>always return true</returns>
	virtual BOOL CanAccessExternal() { return TRUE; }
};

