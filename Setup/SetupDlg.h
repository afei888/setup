// SetupDlg.h : header file
//

#pragma once

#include "BasicExcel.h"
#include "afxcmn.h"
#include "afxwin.h"
using namespace YExcel;
// CSetupDlg dialog
enum {RUNNUMBER,PAGE,CHAPTER,LEFT,TOP,RIGHT,BOTTOM,TYPE,JOINPREV,NAME};
class CSetupDlg : public CDialog
{
// Construction
public:
	CSetupDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_SETUP_DIALOG };
	enum {SELECT_VEDIO_PATH = 0x01,
		SELECT_EXECL = 0x02,
		SELECT_DESTINITION = 0x04,
		PAGE_MAX = 0x10,
		RECITER_MAX = 0x20,
		LANGUAGE_MAX = 0x40};
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	int		m_Ready;
	BasicExcel m_Setupxls;
	BasicExcel m_execl;
	BasicExcel m_pagexls;
	CString m_execl_name;
	CString m_str_videopath;
	CString m_str_ReciterName;
	CString m_str_LanguageName;
	CString m_strPageCnt;
	int m_LanguageCnt;
	int m_ReciterCnt;
	BOOL	m_SetupIsRead;
	CString m_str_DestPath;
	afx_msg void OnBnClickedButtonexeclselect();
	afx_msg void OnBnClickedButtonvideopathselect();
	afx_msg void OnBnClickedButtondestpathdest();
	afx_msg void OnBnClickedButtonmake();
	CString m_strLanguageStartCol;
	CProgressCtrl m_progressctrl;
	CFile Logfile,NarrationLog;
	//CComboBox m_language;
	CString m_strNarration;
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();
	virtual void OnFinalRelease();
	CString m_narration_title;
	CString m_setup;
	CString m_1xls;
	virtual BOOL DestroyWindow();
protected:
	virtual void PostNcDestroy();
	virtual void OnOK();
public:
	afx_msg void OnBnClickedButtondelete();
	CComboBox m_ComboxNarrList;
	afx_msg void OnBnClickedButtondeletexls();
};
