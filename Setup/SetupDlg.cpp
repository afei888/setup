// SetupDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Setup.h"
#include "SetupDlg.h"
#include "bookdef.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CSetupDlg dialog




CSetupDlg::CSetupDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CSetupDlg::IDD, pParent)
	, m_execl_name(_T(""))
	, m_str_videopath(_T(""))
	, m_strPageCnt(_T(""))
	, m_LanguageCnt(0)
	, m_ReciterCnt(0)
	, m_str_DestPath(_T(""))
	, m_strLanguageStartCol(_T(""))
	, m_strNarration(_T(""))
	, m_narration_title(_T(""))
	, m_setup(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_Ready = 0;
	m_str_ReciterName = _T("RECITER");
	m_str_LanguageName =  _T("LANGUAGE");
	m_SetupIsRead = false;
}

void CSetupDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	DDX_Text(pDX, IDC_EDIT2_execl_path, m_execl_name);
	DDX_Text(pDX, IDC_EDIT_path, m_str_videopath);
	DDX_Text(pDX, IDC_EDIT_page_col, m_strPageCnt);
	DDX_Text(pDX, IDC_EDIT_languagecnt, m_LanguageCnt);
	DDX_Text(pDX, IDC_EDIT_recitercnt, m_ReciterCnt);
	DDX_Text(pDX, IDC_EDIT_pathdest, m_str_DestPath);
	DDX_Text(pDX, IDC_EDIT_languagestart_col, m_strLanguageStartCol);
	DDX_Control(pDX, IDC_PROGRESS1, m_progressctrl);
	//DDX_Control(pDX, IDC_COMBO_language, m_language);
	DDX_Text(pDX, IDC_EDIT_narration, m_strNarration);
	DDX_Text(pDX, IDC_EDIT_narrtitle, m_narration_title);
	DDX_Control(pDX, IDC_COMBO1, m_ComboxNarrList);
}

BEGIN_MESSAGE_MAP(CSetupDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON_execlselect, &CSetupDlg::OnBnClickedButtonexeclselect)
	ON_BN_CLICKED(IDC_BUTTON_pathselect, &CSetupDlg::OnBnClickedButtonvideopathselect)
	ON_BN_CLICKED(IDC_BUTTON_pathdest, &CSetupDlg::OnBnClickedButtondestpathdest)
	ON_BN_CLICKED(IDC_BUTTON_make, &CSetupDlg::OnBnClickedButtonmake)
	ON_BN_CLICKED(IDC_BUTTON1, &CSetupDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CSetupDlg::OnBnClickedButton2)
	ON_BN_CLICKED(IDC_BUTTON_delete, &CSetupDlg::OnBnClickedButtondelete)
	ON_BN_CLICKED(IDC_BUTTON_delete_xls, &CSetupDlg::OnBnClickedButtondeletexls)
END_MESSAGE_MAP()


// CSetupDlg message handlers

BOOL CSetupDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CSetupDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CSetupDlg::OnPaint()
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
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CSetupDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CSetupDlg::OnBnClickedButtonexeclselect()
{
	// TODO: 在此添加控件通知处理程序代码
	CFileDialog dlgFile(TRUE);
	CString fileName;
	const int c_cMaxFiles = 100;
	const int c_cbBuffSize = (c_cMaxFiles * (MAX_PATH + 1)) + 1;
	dlgFile.GetOFN().lpstrFile = fileName.GetBuffer(c_cbBuffSize);
	dlgFile.GetOFN().nMaxFile = c_cMaxFiles;

	dlgFile.DoModal();
	m_execl_name = dlgFile.GetFolderPath() + _T("\\") + dlgFile.GetFileName();        //获取所选择的文件名称

	//if(!m_execl.Load(m_execl_name.GetBuffer()))
	//{
	//	AfxMessageBox(_T("Open execl file failed")); 
	//	m_Ready &= ~SELECT_EXECL;
	//}
	//else
	//{
	//	m_Ready |= SELECT_EXECL;
	//	GetDlgItem(IDC_EDIT2_execl_path)->SetWindowTextW(m_execl_name);
	//}
	m_Ready |= SELECT_EXECL;
	GetDlgItem(IDC_EDIT2_execl_path)->SetWindowTextW(m_execl_name);
	fileName.ReleaseBuffer();
}

void CSetupDlg::OnBnClickedButtonvideopathselect()
{
	// TODO: 在此添加控件通知处理程序代码
  char szPath[MAX_PATH];     //存放选择的目录路径 
    CString str;

    ZeroMemory(szPath, sizeof(szPath));   

    BROWSEINFO bi;   
    bi.hwndOwner = m_hWnd;   
    bi.pidlRoot = NULL;   
    bi.pszDisplayName = (LPWSTR)szPath;   
	bi.lpszTitle = _T("Please select folder of the voice:");   
    bi.ulFlags = 0;   
    bi.lpfn = NULL;   
    bi.lParam = 0;   
    bi.iImage = 0;   
    //弹出选择目录对话框
    LPITEMIDLIST lp = SHBrowseForFolder(&bi);   

    if(lp && SHGetPathFromIDList(lp, (LPWSTR)szPath))   
    {
        m_str_videopath.Format(_T("%s\\"),  szPath);
		GetDlgItem(IDC_EDIT_path)->SetWindowTextW(m_str_videopath);
		m_Ready |= SELECT_VEDIO_PATH;
    }
    else   
	{
		AfxMessageBox(_T("Invalid folder,please select again"));
		m_Ready &= ~SELECT_VEDIO_PATH;
	}
}
BOOL FindSpecFile(CString path)
{
	CFileFind dirfind;
	BOOL isfind = false;

	isfind = dirfind.FindFile(path,0);
	if(!isfind)
	{
		 return false;
	}
	else 
	{
		return true;
	}
}

void CSetupDlg::OnBnClickedButtondestpathdest()
{
	// TODO: 在此添加控件通知处理程序代码
	 char szPath[MAX_PATH];     //存放选择的目录路径 
    CString str;

    ZeroMemory(szPath, sizeof(szPath));   

    BROWSEINFO bi;   
    bi.hwndOwner = m_hWnd;   
    bi.pidlRoot = NULL;   
    bi.pszDisplayName = (LPWSTR)szPath;   
	bi.lpszTitle = _T("Please select Destination folder:");   
    bi.ulFlags = 0;   
    bi.lpfn = NULL;   
    bi.lParam = 0;   
    bi.iImage = 0;   
    //弹出选择目录对话框
    LPITEMIDLIST lp = SHBrowseForFolder(&bi);   

    if(lp && SHGetPathFromIDList(lp, (LPWSTR)szPath))   
    {
        m_str_DestPath.Format(_T("%s"),  szPath);


		m_setup = m_str_DestPath + _T("\\Setup.xls");
		m_1xls = m_str_DestPath + _T("\\1.xls");
		CFileStatus filestatus;
		if(CFile::GetStatus(m_1xls, filestatus))
		{
			CopyFile(m_1xls,m_setup, TRUE);
			BasicExcel newexcel;
			newexcel.Load(m_setup);
			newexcel.GetWorksheet(0)->Rename(_T(SHEET_SETUP));
			newexcel.Save();
			newexcel.Close();
		}
		else
		{
			AfxMessageBox(_T("无法找到【1.xls】文件，请在目的目录里面建立一个【1.xls】文件后再试")); 
			m_setup = _T("");
			m_1xls = _T("");
			return;
		}
		//if(!FindSpecFile(m_setup))
		//{
		//	BasicExcel newexcel;
		//	newexcel.New(3);
		//	newexcel.GetWorksheet(0)->Rename(_T(SHEET_SETUP));
		//	newexcel.SaveAs(m_setup.GetBuffer());
		//	newexcel.Close();
		//}
		if(!m_Setupxls.Load(m_setup.GetBuffer()))
		{
			AfxMessageBox(_T("Open setup execl file failed")); 
		}
		else
		{
			GetDlgItem(IDC_EDIT_pathdest)->SetWindowTextW(m_str_DestPath);
			m_Ready |= SELECT_DESTINITION;

			BasicExcelWorksheet* psetupSheet = m_Setupxls.GetWorksheet(_T(SHEET_SETUP));
			int NarrationCount = 0;
			if(psetupSheet->Cell(SETUP_NARRATION_ROW,0)->Type() == BasicExcelCell::UNDEFINED)//第一次
			{
				NarrationCount = 0;
			}
			else
			{
				NarrationCount = psetupSheet->Cell(SETUP_NARRATION_ROW,0)->GetInteger();
			}
			CString str;
			for(int i = 0;i < NarrationCount;i++)
			{
				str = psetupSheet->Cell(SETUP_NARRATION_ROW,i + 1)->GetWString();
				m_ComboxNarrList.AddString(str);
			}
			
		}
    }
    else   
	{
		AfxMessageBox(_T("Invalid folder,please select5 again")); 
		m_Ready &= ~SELECT_DESTINITION;
	}
}
int	CalcColFromChar(CString str)
{
	wchar_t col;

	col = str.GetAt(0);
	 
	if(col >= 'A' && col <= 'Z')
		return (col - 'A');
	else
		return 0;
}
void GenerateFolder(CString & destpath,int pageno,CString & str_reciter,CString & str_language,CString &xls)
{
	CString strpage,strxls;

	strpage.Format(_T("\\%d"),pageno);
	strxls.Format(_T("%d.xls"),pageno);
	if(!FindSpecFile(destpath + strpage + _T("\\*.*")))
	{
		CreateDirectory(destpath + strpage,NULL);
	}
	if(!FindSpecFile(destpath + strpage + _T("\\") + str_reciter + _T("\\*.*")))
	{
		CreateDirectory(destpath + strpage + _T("\\") + str_reciter,NULL);
	}
	if(!FindSpecFile(destpath + strpage + _T("\\") + str_language + _T("\\*.*")))
	{
		CreateDirectory(destpath + strpage + _T("\\") + str_language,NULL);
	}
	CString xlsname = destpath + strpage + _T("\\") + strxls;
	CFileStatus filestatus;
	if(CFile::GetStatus(xls, filestatus))
	{
		if(!CopyFile(xls,xlsname, FALSE))
		{
			::AfxMessageBox(_T("拷贝文件失败,请关闭已经打开setup.xls，重试")); 
			return;
		}
		else
		{
			BasicExcel newexcel;
			newexcel.Load(xlsname);
			newexcel.GetWorksheet(0)->Rename(_T(SHEET_AUDIO));
			newexcel.GetWorksheet(1)->Rename(_T(SHEET_NARRATION));
			newexcel.GetWorksheet(2)->Rename(_T(SHEET_AUDIO_BACK));
			newexcel.GetWorksheet(3)->Rename(_T(SHEET_NARRATION_BACK));
			newexcel.Save();
			newexcel.Close();
		}
	}
	else
	{
		::AfxMessageBox(_T("无法找到【1.xls】文件，请在目的目录里面建立一个【1.xls】文件后再试")); 
		return;
	}
	//if(!FindSpecFile(xlsname))
	//{
	//	BasicExcel newexcel;
	//	newexcel.New(3);
	//	newexcel.GetWorksheet(0)->Rename(_T(SHEET_AUDIO));
	//	newexcel.GetWorksheet(1)->Rename(_T(SHEET_NARRATION));
	//	newexcel.SaveAs(xlsname.GetBuffer());
	//	newexcel.Close();
	//}

}

BasicExcelWorksheet * GetPageExeclSheet(LPVOID lpParam,int pageno,int serial)
{

	CSetupDlg *pDlg = (CSetupDlg*)lpParam;
	CString strpage,str;
	CString logstr,pagefolder,xlsfile;

	BasicExcelWorksheet *pAudioSheet = NULL;


	strpage.Format(_T("%d.xls"),pageno);
	pagefolder.Format(_T("\\%d\\"),pageno);
	xlsfile = pDlg->m_str_DestPath + pagefolder + strpage;
	if(FindSpecFile(pDlg->m_str_DestPath + pagefolder + _T("*.*")))
	{
		if(!FindSpecFile(xlsfile))
		{
			logstr.Format(_T("wainning(%d): Can not find 【%s】 in【page-%d】.\r\n"),serial,strpage,pageno);
			pDlg->Logfile.Write(logstr.GetBuffer(),logstr.GetLength() * 2);
			pDlg->Logfile.Flush();
			logstr.ReleaseBuffer();	
		}
	}
	else
	{
		logstr.Format(_T("wainning(%d): Can not find【page-%d】folder.\r\n"),serial,pageno);
		pDlg->Logfile.Write(logstr.GetBuffer(),logstr.GetLength() * 2);
		pDlg->Logfile.Flush();
		logstr.ReleaseBuffer();	
	}
	str.Format(_T("正在处理第%d页，并写入文件...请等待"),pageno);
	pDlg->SetWindowTextW(str);
	pDlg->m_pagexls.Save();
	pDlg->m_pagexls.Close();	
	if(!pDlg->m_pagexls.Load(xlsfile.GetBuffer()))
	{
		logstr.Format(_T("wainning(%d): Can not load 【%s】 in【page-%d】 folder.\r\n"),serial,strpage,pageno);
		pDlg->Logfile.Write(logstr.GetBuffer(),logstr.GetLength() * 2);
		pDlg->Logfile.Flush();
		logstr.Empty();	
		pDlg->m_pagexls.Close();
	}
	else
	{
		pAudioSheet = pDlg->m_pagexls.GetWorksheet(_T(SHEET_AUDIO_BACK));//
		//pAudioSheet->Cell(0,NAME + 1)->EraseContents();
		//pAudioSheet->Cell(0,NAME + 1)->SetWString(strpage.GetBuffer());
	}
	return pAudioSheet;
}

DWORD WINAPI MakeFunction( LPVOID lpParam ) 
{ 
	int rowmax = 0,colmax = 0,row = 0,col = 0;
	int pagenumber = 0;
	CString strReciterFolder,strLanguageFolder,strPage,strFileName;
	CString strReciter,strLanguage,str;
	CFileStatus filestatus;

	CSetupDlg *pDlg = (CSetupDlg*)lpParam;
	pDlg->GetDlgItem(IDC_BUTTON_make)->EnableWindow(false);//
	pDlg->GetDlgItem(IDC_BUTTON_delete_xls)->EnableWindow(false);//
	str.Format(_T("正在读取EXECL文件内容...请稍等"));
	pDlg->SetWindowTextW(str);
	if(!pDlg->m_execl.Load(pDlg->m_execl_name.GetBuffer()))
	{
		::AfxMessageBox(_T("打开execl文件失败，如果其他应用程序已经打开该文件，请关闭后再试")); 
		pDlg->Logfile.Flush();
		pDlg->Logfile.Close();
		pDlg->GetDlgItem(IDC_BUTTON_make)->EnableWindow(true);
		return 0;
	}
	
	pDlg->GetDlgItem(IDOK)->EnableWindow(false);
	pDlg->GetDlgItem(IDC_BUTTON1)->EnableWindow(false);
	BasicExcelWorksheet* pSheet = pDlg->m_execl.GetWorksheet(0);
	ASSERT(pSheet);
	rowmax = pSheet->GetTotalRows();

	BasicExcelWorksheet* psetupSheet = pDlg->m_Setupxls.GetWorksheet(_T(SHEET_SETUP));
	ASSERT(psetupSheet);
	int index = 0;
	psetupSheet->Cell(SETUP_AUDIO_ROW,index++)->Set(2);
	psetupSheet->Cell(SETUP_AUDIO_ROW,index++)->SetWString(pDlg->m_str_ReciterName);//
	psetupSheet->Cell(SETUP_AUDIO_ROW,index++)->Set(pDlg->m_ReciterCnt);
	//m_str_ReciterName
	//m_str_LanguageName
	//pDlg->UpdateData(true);
	pDlg->m_progressctrl.SetRange(0,rowmax);
	pDlg->m_progressctrl.SetStep(1);
	int languagestart = CalcColFromChar(pDlg->m_strLanguageStartCol);
	int pagestart = CalcColFromChar(pDlg->m_strPageCnt);
	int pagelast = 0;
	int fristrow = 0;
	CString strline;
	int Setup[255];
	int serial = 0;
	memset(Setup,0,sizeof(Setup));//0 = none... mean this col is empty 
	//for test
	int file = 0;
	int languageindex = 0;
	int mustmakedir = false;
	int sentenchindex = 0;
	int FileNameColIndex = 0;
	BasicExcelWorksheet *pAudioSheet = NULL;

	for(row = 0;row < rowmax;row++)
	{
		pDlg->m_progressctrl.SetPos(row);
		colmax =  pSheet->GetTotalCols();
		int type = pSheet->Cell(row,pagestart)->Type();
		if(colmax <= 1 || type == BasicExcelCell::UNDEFINED || type == BasicExcelCell::STRING
			|| type == BasicExcelCell::FORMULA || type == BasicExcelCell::WSTRING)
			continue;
		if(!fristrow)
		{
			fristrow = 0xff;
			if(colmax < pDlg->m_ReciterCnt + pDlg->m_LanguageCnt)
			{
				strline.Format(_T("ERROR:frist rows < (reciter + language) ,in execl file. \r\n"),row + 1,col + 1);
				pDlg->Logfile.Write(strline.GetBuffer(),strline.GetLength() * 2);
				pDlg->Logfile.Flush();
				strline.ReleaseBuffer();
				pDlg->Logfile.Close();
				pDlg->GetDlgItem(IDOK)->EnableWindow(true);
				pDlg->GetDlgItem(IDC_BUTTON1)->EnableWindow(true);
				return 0 ;
			}
		}
		pagenumber = pSheet->Cell(row,pagestart)->GetInteger();
	
		str.Format(_T("正在处理第%d页，并拷贝文件...请等待"),pagenumber);
		pDlg->SetWindowTextW(str);
		strPage.Format(_T("\\%d\\"),pagenumber);
		if(pagelast != pagenumber)
		{
			GenerateFolder(pDlg->m_str_DestPath,pagenumber,pDlg->m_str_ReciterName,pDlg->m_str_LanguageName,pDlg->m_1xls);
		}
		strReciterFolder = pDlg->m_str_DestPath + strPage + pDlg->m_str_ReciterName + _T("\\");
		strLanguageFolder = pDlg->m_str_DestPath + strPage + pDlg->m_str_LanguageName + _T("\\");
		// 
		if(pagelast != pagenumber)
		{
			pAudioSheet =  GetPageExeclSheet((LPVOID)pDlg,pagenumber,serial++);
			sentenchindex = 0;
		}
		FileNameColIndex = 0;
		for(col = 0;col < colmax;col++)
		{
			int type = pSheet->Cell(row,col)->Type();
			if(type == BasicExcelCell::WSTRING || type == BasicExcelCell::STRING)
			{
				if(type == BasicExcelCell::WSTRING)
					strFileName = pSheet->Cell(row,col)->GetWString();
				else
				{
					//strFileName = pSheet->Cell(row,col)->GetString();
					const char *filename = pSheet->Cell(row,col)->GetString();
					int nLen = strlen(filename) + 1;
					int nwLen = MultiByteToWideChar(CP_ACP, 0, filename, nLen, NULL, 0);

					TCHAR lpszFile[256];
					MultiByteToWideChar(CP_ACP, 0, filename, nLen, lpszFile, nwLen);
					strFileName = lpszFile;
				}

				//--------------------------------for test-----------------------------------------------------
				/*strline.Format(_T("【%d行】【%d列】	= 【%s】 \r\n"),row + 1,col + 1,strFileName.GetBuffer());
				pDlg->Logfile.Write(strline.GetBuffer(),strline.GetLength() * 2);
				pDlg->Logfile.Flush();
				strline.ReleaseBuffer();*/
				//---------------------------------------------------------------------------------------------
				//if(strFileName.IsEmpty())
				//{
				//	if(pagelast != pagenumber)
				//		mustmakedir = true;
				//	else
				//		continue;
				//		strline.Format(_T("【%d行】【%d列】empty \r\n"),row + 1,col + 1);
				//		pDlg->Logfile.Write(strline.GetBuffer(),strline.GetLength() * 2);
				//		pDlg->Logfile.Flush();
				//		strline.ReleaseBuffer();
				//}
				if(strFileName.Find(_T("/"),0) >= 0)
				{
					strline.Format(_T("【%d行】【%d列】find \"/\",ignore \r\n"),row + 1,col + 1);
					pDlg->Logfile.Write(strline.GetBuffer(),strline.GetLength() * 2);
					pDlg->Logfile.Flush();
					strline.ReleaseBuffer();	
					if(fristrow == 0xffff && Setup[col] == 0xffff)
					{				
						if(pagelast != pagenumber)
							mustmakedir = true;
						else
							continue;
					}
					else
						continue;
				}
			}		
			else
			{
				strline.Format(_T("【%d行】【%d列】should not empty \r\n"),row + 1,col + 1);
				pDlg->Logfile.Write(strline.GetBuffer(),strline.GetLength() * 2);
				pDlg->Logfile.Flush();
				strline.ReleaseBuffer();
				if(fristrow == 0xffff && Setup[col] == 0xffff)
				{
					if(pagelast != pagenumber)
						mustmakedir = true;;
				}
				else
					continue;
			}
			if(col < pDlg->m_ReciterCnt)//reciter process
			{
				if(fristrow == 0xff)
				{
					Setup[col] = 0xffff;//0xffff = have... mean this col is not empty 
				}
				//FileNameColIndex++;
				strReciter.Format(_T("reciter-%d"),col + 1);
				if(pagelast != pagenumber || mustmakedir)
				{
					if(fristrow == 0xff)
					{
						psetupSheet->Cell(SETUP_AUDIO_ROW,index)->EraseContents();
						psetupSheet->Cell(SETUP_AUDIO_ROW,index)->SetWString(strReciter);//
						psetupSheet->Cell(SETUP_AUDIO_ROW_NAME,index)->EraseContents();
						psetupSheet->Cell(SETUP_AUDIO_ROW_NAME,index++)->SetWString(strReciter);//
					}
					CreateDirectory(strReciterFolder + strReciter + _T("\\"),NULL);
				}
				if(mustmakedir)
				{
					mustmakedir = false;
					continue;
				}
				if(CFile::GetStatus(pDlg->m_str_videopath + strFileName, filestatus))
				//if(FindSpecFile(pDlg->m_str_videopath + strFileName))
				{
					if(!CopyFile(pDlg->m_str_videopath + strFileName,strReciterFolder + strReciter + _T("\\") + strFileName, TRUE))
					{
						strline.Format(_T("wainning(%d): Copy File 【%s】 to %s was failed.error=%d\r\n"),serial++,strFileName,strPage + pDlg->m_str_ReciterName + _T("\\") + strReciter,GetLastError());
						pDlg->Logfile.Write(strline.GetBuffer(),strline.GetLength() * 2);
						pDlg->Logfile.Flush();
						strline.ReleaseBuffer();
					}
					else
					{
				/*		strline.Format(_T("wainning(%d): Copy File 【%s】 to %s was successful\r\n"),serial++,strFileName,strPage + pDlg->m_str_ReciterName + _T("\\") + strReciter);
						pDlg->Logfile.Write(strline.GetBuffer(),strline.GetLength() * 2);
						pDlg->Logfile.Flush();
						strline.ReleaseBuffer();*/
					}
				}
				else
				{
					strline.Format(_T("wainning(%d): File 【%s】 is not existing.\r\n"),serial++,strFileName);
					pDlg->Logfile.Write(strline.GetBuffer(),strline.GetLength() * 2);
					pDlg->Logfile.Flush();
					strline.ReleaseBuffer();
					//--------------------------------------for test
			/*		CString strtestfile;
					strtestfile.Format(_T("number%d.mp3"),file++);
					if(file >= 9)
						file = 0;
					if(CFile::GetStatus(pDlg->m_str_videopath + strtestfile, filestatus))
						CopyFile(pDlg->m_str_videopath + strtestfile, strReciterFolder + strReciter + _T("\\") + strtestfile, TRUE);*/
					//--------------------------------------for test
				}
			}
			else if (col >= languagestart && languageindex < pDlg-> m_LanguageCnt)// language process
			{
				strLanguage.Format(_T("language-%d"),languageindex + 1);
				languageindex++;
				//FileNameColIndex++;
				if(fristrow == 0xff)
				{
					Setup[col] = 0xffff;//0xffff = have... mean this col is not empty 
				}
				if(pagelast != pagenumber || mustmakedir)
				{
					if(col == languagestart && fristrow == 0xff)
					{
						psetupSheet->Cell(SETUP_AUDIO_ROW,index)->EraseContents();
						psetupSheet->Cell(SETUP_AUDIO_ROW,index++)->SetWString(pDlg->m_str_LanguageName);//
						psetupSheet->Cell(SETUP_AUDIO_ROW,index)->EraseContents();
						psetupSheet->Cell(SETUP_AUDIO_ROW,index++)->Set(pDlg->m_LanguageCnt);//
					}
					if(fristrow == 0xff)
					{
						psetupSheet->Cell(SETUP_AUDIO_ROW,index)->EraseContents();
						psetupSheet->Cell(SETUP_AUDIO_ROW,index)->SetWString(strLanguage);//
						psetupSheet->Cell(SETUP_AUDIO_ROW_NAME,index)->EraseContents();
						psetupSheet->Cell(SETUP_AUDIO_ROW_NAME,index++)->SetWString(strLanguage);//
					}
					CreateDirectory(strLanguageFolder + strLanguage + _T("\\"),NULL);
				}
				if(mustmakedir)
				{
					mustmakedir = false;
					continue;
				}
				if(CFile::GetStatus(pDlg->m_str_videopath + strFileName, filestatus))
				//if(FindSpecFile(pDlg->m_str_videopath + strFileName))
				{
					if(!CopyFile(pDlg->m_str_videopath + strFileName, strLanguageFolder + strLanguage + _T("\\") + strFileName, TRUE))
					{
						strline.Format(_T("wainning(%d): Copy File 【%s】 to %s was failed.error=%d\r\n"),serial++,strFileName,strPage + pDlg->m_str_LanguageName + _T("\\") + strLanguage,GetLastError());
						pDlg->Logfile.Write(strline.GetBuffer(),strline.GetLength() * 2);
						pDlg->Logfile.Flush();
						strline.ReleaseBuffer();
					}
					else
					{
						/*strline.Format(_T("wainning(%d): Copy File 【%s】 to %s was successful \r\n"),serial++,strFileName,strPage + pDlg->m_str_LanguageName + _T("\\") + strLanguage);
						pDlg->Logfile.Write(strline.GetBuffer(),strline.GetLength() * 2);
						pDlg->Logfile.Flush();
						strline.ReleaseBuffer();*/
					}
				}
				else
				{
					strline.Format(_T("wainning(%d): File 【%s】 of copy to 【%s】is not existing.\r\n"),serial++,strFileName,strLanguageFolder + strLanguage + _T("\\") + strFileName);
					pDlg->Logfile.Write(strline.GetBuffer(),strline.GetLength() * 2);
					pDlg->Logfile.Flush();
					strline.ReleaseBuffer();
					//--------------------------------------for test
					/*CString strtestfile;
					strtestfile.Format(_T("number%d.mp3"),file++);
					if(file >= 9)
						file = 0;
					if(CFile::GetStatus(pDlg->m_str_videopath + strtestfile, filestatus))
						CopyFile(pDlg->m_str_videopath + strtestfile, strLanguageFolder + strLanguage + _T("\\") + strtestfile, TRUE);*/
					//--------------------------------------for test
				}
			} 
			else
			{
				strline.Format(_T("【%d行】【%d列】ignore 【%s】\r\n"),row + 1,col + 1,strFileName);
				pDlg->Logfile.Write(strline.GetBuffer(),strline.GetLength() * 2);
				pDlg->Logfile.Flush();
				strline.ReleaseBuffer();
			}
			if(pAudioSheet != NULL)
			{
				pAudioSheet->Cell(sentenchindex,NAME + FileNameColIndex)->EraseContents();
				pAudioSheet->Cell(sentenchindex,NAME + FileNameColIndex++)->SetWString(strFileName.GetBuffer());
			}
		}
		pagelast = pagenumber;
		languageindex = 0;
		sentenchindex++;
		if(fristrow == 0xff)
			fristrow = 0xffff;
	}
	pDlg->Logfile.Flush();
	pDlg->Logfile.Close();
	psetupSheet->Cell(SETUP_EDITER_MAX,2)->Set(pagenumber);//
	pDlg->m_Setupxls.Save();
	pDlg->m_pagexls.Save();
	pDlg->m_pagexls.Close();
	pDlg->SetWindowTextW(_T("生成完成"));
//	pDlg->GetDlgItem(IDC_BUTTON_make)->SetWindowTextW(_T("生成完成"));
	pDlg->GetDlgItem(IDC_BUTTON_make)->EnableWindow(true);
	pDlg->GetDlgItem(IDC_BUTTON_delete_xls)->EnableWindow(true);//
	pDlg->GetDlgItem(IDOK)->EnableWindow(true);
	pDlg->GetDlgItem(IDC_BUTTON1)->EnableWindow(true);
	str.Format(_T("execl文件处理完毕"));
	pDlg->SetWindowTextW(str);
    return 0; 
} 

void CSetupDlg::OnBnClickedButtonmake()
{
	// TODO: 在此添加控件通知处理程序代码

	HANDLE  hThread;
	DWORD   dwThreadId;

	if(!(m_Ready & SELECT_VEDIO_PATH))
	{
		AfxMessageBox(_T("没有选择源语音目录，请选择")); 
		return;
	}
	if(!(m_Ready & SELECT_EXECL))
	{
		AfxMessageBox(_T("没有选择EXECL文件，请选择")); 
		return;
	}
	if(!(m_Ready & SELECT_DESTINITION))
	{
		AfxMessageBox(_T("没有选择目的文件夹目录，请选择")); 
		return;
	}
	UpdateData(true);//
	if(m_LanguageCnt < 0 || m_LanguageCnt > 255)
	{
		AfxMessageBox(_T("无效的语言个数，请重新输入")); 
		return;
	}
	if(m_ReciterCnt < 0 || m_ReciterCnt > 255)
	{
		AfxMessageBox(_T("无效的朗读者个数，请重新输入")); 
		return;
	}
	if(CalcColFromChar(m_strLanguageStartCol) == 0)
	{
		AfxMessageBox(_T("无效的语言起始列数，请重新输入")); 
		return;
	}
	if(CalcColFromChar(m_strPageCnt) == 0)
	{
		AfxMessageBox(_T("无效的页码起始列数，请重新输入")); 
		return;
	}
	if(CalcColFromChar(m_strLanguageStartCol) < m_ReciterCnt)
	{
		AfxMessageBox(_T("语言起始列数必须大于朗读者个数，请重新输入")); 
		return;
	}
	//if(m_PageCnt < 0 || m_PageCnt > 255)
	//{
	//	AfxMessageBox(_T("Invalid page number path,please input again")); 
	//	return;
	//}
	CFileException e;
	TCHAR* pszFileName = _T("log.txt");
	if(!Logfile.Open(pszFileName, CFile::modeCreate | CFile::modeReadWrite, &e))
	{
	   AfxMessageBox(_T("File could not be opened %d\n"), e.m_cause);
	   return;
	}
	CTime   t = CTime::GetCurrentTime(); 
	CString title;
	unsigned short int feff = 0xfeff;   
	Logfile.Write(&feff,sizeof(short int)); 

	title.Format(_T("make vedio log file  %s build time at %d 月 %d 日 %d 时 %d 分 %d 秒\r\n"),m_execl_name,t.GetMonth(),t.GetDay(),t.GetHour(),t.GetMinute(),t.GetSecond());
	Logfile.Write(title.GetBuffer(),title.GetLength() * 2);
	hThread = CreateThread( 
            NULL,                   // default security attributes
            0,                      // use default stack size  
            MakeFunction,       // thread function name
            (LPVOID)this,          // argument to thread function 
            0,                      // use default creation flags 
            &dwThreadId);   // returns the thread identifier 

     if (hThread == NULL) 
     {
           AfxMessageBox(_T("CreateThread"));
     }

}
BOOL ReadLine(wchar_t *filebuff,int &pageno,int &sentence,CString &narr,int filelength)
{
	static int seek = 0;
	wchar_t *buffer = filebuff + seek;
	if(seek >= filelength || seek < 0)
	{
		seek = 0;
		return false;
	}
	CString strpage,strsentence;

	while(*buffer != 0x007c && *buffer != 0x000a)
	{
		strpage.AppendChar(*buffer);
		buffer++;
	}
	if(strpage.IsEmpty())
		return false;
	else
		pageno = _wtoi(strpage);
	buffer++;
	while(*buffer != 0x007c && *buffer != 0x000a)
	{
		strsentence.AppendChar(*buffer);
		buffer++;
	}	
	buffer++;
	if(strsentence.IsEmpty())
		return false;
	else
		sentence = _wtoi(strsentence);
	narr.Empty();
	narr.ReleaseBuffer();
	while(*buffer != 0x000a)
	{
		narr.AppendChar(*buffer);
		buffer++;
	}	
	seek = (int)(buffer - filebuff + 1);
	return true;
}
int CountPagemax(wchar_t * filebuff,int length)
{
	int pagemax = 0;
	int count7c = 0;
	while(true)
	{
		if(*filebuff == 0x000a)
			pagemax++;
		else if(*filebuff == 0x007c)
			count7c++;
		length--;
		if(length == 0)
			break;
		filebuff++;
	}
	if(count7c > pagemax && pagemax > 0)
		return pagemax;
	else
		return 0;
}
DWORD WINAPI MakeNarration( LPVOID lpParam ) 
{ 

	CSetupDlg *pDlg = (CSetupDlg*)lpParam;
	CString strpage,str;

	str.Format(_T("写入歌词信息..."));
	pDlg->SetWindowTextW(str);

	CFile narrationfile;
	if(!narrationfile.Open(pDlg->m_strNarration,CFile::typeBinary | CFile::modeRead,NULL))
	{
		::AfxMessageBox(_T("Narration file can not open"));
		 return 0;
	}
	int bufferlength = (int)(narrationfile.GetLength());
	char *preadbuff = new char[bufferlength];
	narrationfile.Read(preadbuff,bufferlength);
	bufferlength = MultiByteToWideChar(CP_UTF8, 0, preadbuff, -1, NULL, 0);//CP_ACP是本地机器Local属性

	//可以修改成指定语言GB2312 936. Shift-JIS 932 UTF-8 65001.
	wchar_t *FileBuffer;
	FileBuffer = new wchar_t[bufferlength];
	MultiByteToWideChar(CP_UTF8, 0, preadbuff, -1, FileBuffer, bufferlength);
	int pagemax = 0;
	pagemax = CountPagemax(FileBuffer,bufferlength);
	if(pagemax == 0)
	{
		pDlg->NarrationLog.Flush();
		pDlg->NarrationLog.Close();
		delete [] FileBuffer;
		delete [] preadbuff;
		pDlg->m_Setupxls.Save();
		//pDlg->m_Setupxls.Close();
		narrationfile.Close();
		::AfxMessageBox(_T("歌词文件格式不对，请重新选择"));
		return 0;
	}
	pDlg->m_ComboxNarrList.AddString(pDlg->m_narration_title);
	BasicExcelWorksheet* psetupSheet = pDlg->m_Setupxls.GetWorksheet(_T(SHEET_SETUP));
	pDlg->GetDlgItem(IDC_BUTTON1)->EnableWindow(false);
	pDlg->GetDlgItem(IDOK)->EnableWindow(false);
	pDlg->GetDlgItem(IDC_BUTTON_make)->EnableWindow(false);
	int NarrationCount = 0;
	if(psetupSheet->Cell(SETUP_NARRATION_ROW,0)->Type() == BasicExcelCell::UNDEFINED)//第一次
	{
		NarrationCount = 1;
		psetupSheet->Cell(SETUP_NARRATION_ROW,0)->SetInteger(NarrationCount);
	}
	else
	{
		NarrationCount = psetupSheet->Cell(SETUP_NARRATION_ROW,0)->GetInteger() + 1;
		//psetupSheet->Cell(SETUP_NARRATION_ROW,0)->EraseContents();
		psetupSheet->Cell(SETUP_NARRATION_ROW,0)->SetInteger(NarrationCount);
	}
	//psetupSheet->Cell(SETUP_NARRATION_ROW,NarrationCount)->EraseContents();
	//psetupSheet->Cell(SETUP_NARRATION_CODEPAGE,NarrationCount)->EraseContents();
	psetupSheet->Cell(SETUP_NARRATION_ROW,NarrationCount)->SetWString(pDlg->m_narration_title.GetBuffer());//write language name
	psetupSheet->Cell(SETUP_NARRATION_CODEPAGE,NarrationCount)->SetWString(_T("utf8"));//write codepage
	pDlg->m_Setupxls.Save();
//	pDlg->m_Setupxls.Close();

	CString logstr,pagefolder,xlsfile;
	int serial = 0;


	pDlg->m_progressctrl.SetRange(0,pagemax);
	pDlg->m_progressctrl.SetStep(1);

	int narrPageNo = 0,narrChapterNo = 1,pageNoLast = 0xffff;
	int narrSentenceNO = 0,line = 0;
	CString strnarr;
	BasicExcelWorksheet* pNarrSheet,*pAudioSheet;
	//ReadLine(wchar_t *filebuff,int &pageno,int &sentence,CString &narr,int filelength)
	int sentencemax = sizeof(PageNumber)/sizeof(int);
	int pageindex = 0;
	int sentenchindex = 0;
	while(true)
	{
		pDlg->m_progressctrl.StepIt();
		if(!ReadLine(FileBuffer,narrChapterNo,narrSentenceNO,strnarr,bufferlength))
			break;
		if(narrChapterNo <= 0 || narrSentenceNO <= 0)
			continue;
		if(pageindex >= sentencemax)
			break;
		narrPageNo = PageNumber[pageindex++];
		if(pageNoLast != narrPageNo)
		{
			strpage.Format(_T("%d.xls"),narrPageNo);
			pagefolder.Format(_T("\\%d\\"),narrPageNo);
			xlsfile = pDlg->m_str_DestPath + pagefolder + strpage;
			if(FindSpecFile(pDlg->m_str_DestPath + pagefolder + _T("*.*")))
			{
				if(!FindSpecFile(xlsfile))
				{
					logstr.Format(_T("wainning(%d): Can not find 【%s】 in【page-%d】.\r\n"),serial++,strpage,narrPageNo);
					pDlg->NarrationLog.Write(logstr.GetBuffer(),logstr.GetLength() * 2);
					pDlg->NarrationLog.Flush();
					logstr.ReleaseBuffer();	
					continue;
				}
			}
			else
			{
				logstr.Format(_T("wainning(%d): Can not find【page-%d】folder.\r\n"),serial++,narrPageNo);
				pDlg->NarrationLog.Write(logstr.GetBuffer(),logstr.GetLength() * 2);
				pDlg->NarrationLog.Flush();
				logstr.ReleaseBuffer();	
				continue;
			}
			str.Format(_T("正在处理第%d页，并写入文件...请等待"),narrPageNo);
			pDlg->SetWindowTextW(str);
			if(pageNoLast != 0xffff)
			{
				pDlg->m_pagexls.Save();
				pDlg->m_pagexls.Close();	
			}
			if(!pDlg->m_pagexls.Load(xlsfile.GetBuffer()))
			{
				logstr.Format(_T("wainning(%d): Can not load 【%s】 in【page-%d】 folder.\r\n"),serial++,strpage,narrPageNo);
				pDlg->NarrationLog.Write(logstr.GetBuffer(),logstr.GetLength() * 2);
				pDlg->NarrationLog.Flush();
				logstr.Empty();	
				pDlg->m_pagexls.Close();
				pNarrSheet = NULL;
				pAudioSheet = NULL;
				continue;
			}
			else
			{
				pAudioSheet = pNarrSheet = NULL;
				pNarrSheet = pDlg->m_pagexls.GetWorksheet(_T(SHEET_NARRATION_BACK));//
				pAudioSheet = pDlg->m_pagexls.GetWorksheet(_T(SHEET_AUDIO_BACK));//
				if(pNarrSheet != NULL && pAudioSheet != NULL)
				{
					pageNoLast = narrPageNo;
					sentenchindex = 0;
				}
			}
		}
		if(pNarrSheet != NULL && pAudioSheet != NULL)
		{
			pNarrSheet->Cell(sentenchindex,NarrationCount - 1)->SetWString(strnarr.GetBuffer());
			pAudioSheet->Cell(sentenchindex++,CHAPTER)->SetInteger(narrChapterNo);
			
			logstr.Format(_T("***:page=%d sentence=%d.\r\n"),narrPageNo,narrSentenceNO);
			pDlg->NarrationLog.Write(logstr.GetBuffer(),logstr.GetLength() * 2);
			pDlg->NarrationLog.Flush();
			logstr.ReleaseBuffer();	
		}
		else
		{
			logstr.Format(_T("wainning(%d): Get work sheet failed in【page-%d】 folder.\r\n"),serial++,narrPageNo);
			pDlg->NarrationLog.Write(logstr.GetBuffer(),logstr.GetLength() * 2);
			pDlg->NarrationLog.Flush();
			logstr.ReleaseBuffer();	
		}
	}
	ReadLine(FileBuffer,narrPageNo,narrSentenceNO,strnarr,-1);//clear static seek variable
	pDlg->m_pagexls.Save();
	pDlg->m_pagexls.Close();
	pDlg->m_Setupxls.Save();
	//pDlg->m_Setupxls.Close();
	pDlg->NarrationLog.Flush();
	pDlg->NarrationLog.Close();
	delete [] FileBuffer;
	delete [] preadbuff;
	pDlg->GetDlgItem(IDC_BUTTON1)->EnableWindow(true);
	pDlg->GetDlgItem(IDOK)->EnableWindow(true);
	pDlg->GetDlgItem(IDC_BUTTON_make)->EnableWindow(true);
	pDlg->m_ComboxNarrList.SelectString(0,pDlg->m_narration_title);
	pDlg->m_narration_title = _T("");
	pDlg->GetDlgItem(IDC_EDIT_narrtitle)->SetWindowTextW(pDlg->m_narration_title);//
	str.Format(_T("歌词添加完毕."));
	pDlg->SetWindowTextW(str);
	return 0;
}
void CSetupDlg::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码

	HANDLE  hThread;
	DWORD   dwThreadId;


	UpdateData(true);//
	if(m_strNarration.IsEmpty())
	{
		AfxMessageBox(_T("please select a narration file"));
		return;
	}
	if(m_narration_title.IsEmpty())
	{
		AfxMessageBox(_T("please input narration name in edit box"));
		return;
	}
	if(!(m_Ready & SELECT_DESTINITION))
	{
		 AfxMessageBox(_T("Please Select destination path"));
		 return;
	}
	BasicExcelWorksheet* psetupSheet = m_Setupxls.GetWorksheet(_T(SHEET_SETUP));
	if(psetupSheet == NULL)
	{
		 AfxMessageBox(_T("Setup.xls File Loading failed"));
		 return;
	}
	CFileException e;
	TCHAR* pszFileName = _T("narrationlog.txt");
	if(!NarrationLog.Open(pszFileName, CFile::modeCreate | CFile::modeReadWrite, &e))
	{
	   AfxMessageBox(_T("File could not be opened %d\n"), e.m_cause);
	   return;
	}
	CTime   t = CTime::GetCurrentTime(); 
	CString title;
	unsigned short int feff = 0xfeff;   
	NarrationLog.Write(&feff,sizeof(short int)); 

	title.Format(_T("make narration log file %s build time at %d 月 %d 日 %d 时 %d 分 %d 秒\r\n"),m_strNarration,t.GetMonth(),t.GetDay(),t.GetHour(),t.GetMinute(),t.GetSecond());
	NarrationLog.Write(title.GetBuffer(),title.GetLength() * 2);
	hThread = CreateThread( 
            NULL,                   // default security attributes
            0,                      // use default stack size  
            MakeNarration,       // thread function name
            (LPVOID)this,          // argument to thread function 
            0,                      // use default creation flags 
            &dwThreadId);   // returns the thread identifier 

     if (hThread == NULL) 
     {
           AfxMessageBox(_T("CreateThread"));
	 }

}

void CSetupDlg::OnBnClickedButton2()
{
	// TODO: 在此添加控件通知处理程序代码
	CFileDialog dlgFile(TRUE);
	CString fileName;
	const int c_cMaxFiles = 100;
	const int c_cbBuffSize = (c_cMaxFiles * (MAX_PATH + 1)) + 1;
	dlgFile.GetOFN().lpstrFile = fileName.GetBuffer(c_cbBuffSize);
	dlgFile.GetOFN().nMaxFile = c_cMaxFiles;

	dlgFile.DoModal();
	m_strNarration = dlgFile.GetFolderPath() + _T("\\") + dlgFile.GetFileName();        //获取所选择的文件名称
	m_narration_title = _T("");
	GetDlgItem(IDC_EDIT_narration)->SetWindowTextW(m_strNarration);//
	GetDlgItem(IDC_EDIT_narrtitle)->SetWindowTextW(m_narration_title);//
	fileName.ReleaseBuffer();
}

void CSetupDlg::OnFinalRelease()
{
	// TODO: 在此添加专用代码和/或调用基类

	CDialog::OnFinalRelease();
}

BOOL CSetupDlg::DestroyWindow()
{
	// TODO: 在此添加专用代码和/或调用基类
	m_Setupxls.Save();
	m_Setupxls.Close();
	return CDialog::DestroyWindow();
}

void CSetupDlg::PostNcDestroy()
{
	// TODO: 在此添加专用代码和/或调用基类

	CDialog::PostNcDestroy();
}

void CSetupDlg::OnOK()
{
	// TODO: 在此添加专用代码和/或调用基类

	CDialog::OnOK();
}

void CSetupDlg::OnBnClickedButtondelete()
{
	// TODO: 在此添加控件通知处理程序代码
}

void CSetupDlg::OnBnClickedButtondeletexls()
{
	// TODO: 在此添加控件通知处理程序代码
	if(!(m_Ready & SELECT_DESTINITION))
	{
		AfxMessageBox(_T("没有选择目的文件夹目录，请选择")); 
		return;
	}
	m_setup = m_str_DestPath + _T("\\Setup.xls");
	m_1xls = m_str_DestPath + _T("\\1.xls");
	CFileStatus filestatus;
	if(CFile::GetStatus(m_1xls, filestatus))
	{
		m_Setupxls.Save();
		m_Setupxls.Close();
		if(!CopyFile(m_1xls,m_setup, FALSE))
		{
			::AfxMessageBox(_T("拷贝文件失败,请关闭已经打开1.xls，重试")); 
			return;
		}
		else
		{
			BasicExcel newexcel;
			newexcel.Load(m_setup);
			newexcel.GetWorksheet(0)->Rename(_T(SHEET_SETUP));
			newexcel.Save();
			newexcel.Close();
		}
		if(!m_Setupxls.Load(m_setup.GetBuffer()))
		{
			AfxMessageBox(_T("Open setup execl file failed")); 
		}
		else
		{
			m_ComboxNarrList.Clear();
			AfxMessageBox(_T("复位Setup成功")); 
		}
	}
	else
	{
		AfxMessageBox(_T("无法找到【1.xls】文件，请在目的目录里面建立一个【1.xls】文件后再试")); 
		m_setup = _T("");
		m_1xls = _T("");
		return;
	}

}
