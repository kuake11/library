
// DBTest_ADoDlg.cpp: 实现文件
//

#include "stdafx.h"
#include "DBTest_ADo.h"
#include "DBTest_ADoDlg.h"
#include "afxdialogex.h"
#include <time.h>
#include<iostream>
#include<sstream>
#define _CRT_SECURE_NO_DEPRECATE；
#define _CRT_SECURE_NO_WARNINGS；
#pragma warning(disable:4996)；

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnEnChangeEdit1();
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CDBTestADoDlg 对话框



CDBTestADODlg::CDBTestADODlg(CWnd* pParent /*=nullptr*/)////有修改
	: CDialogEx(IDD_DBTEST_ADO_DIALOG, pParent)
	, m_strName(_T(""))
	, m_strStuNum(_T(""))
	, m_strSearchway1(_T(""))
	, m_strSearchway2(_T(""))
	, m_strSearch1(_T(""))
	, m_strSearch2(_T(""))
	, m_Age(0)
	, m_strCmd(_T(""))
	, m_strCmd1(_T(""))
	, m_strCmd2(_T(""))
	, m_strBookName(_T(""))
	, m_strBookAuthorName(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CDBTestADODlg::DoDataExchange(CDataExchange* pDX)//文本框数据（一定记得改掉）
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_ACCESSList);
	DDX_Text(pDX, IDC_EDIT_STUNAME, m_strName);
	DDX_Text(pDX, IDC_EDIT2, m_strStuNum);
	DDX_Text(pDX, IDC_SearchWay1, m_strSearchway1);
	DDX_Text(pDX, IDC_SearchWay2, m_strSearchway2);
	DDX_Text(pDX, IDC_EDIT1, m_strSearch1);
	DDX_Text(pDX, IDC_EDIT5, m_strSearch2);
	DDX_Text(pDX, IDC_EDIT_BOOKNAME, m_strBookName);
	DDX_Text(pDX, IDC_EDIT4, m_strBookAuthorName);
	DDX_Text(pDX, IDC_EDIT6, stu_name_bor);
	DDX_Text(pDX, IDC_EDIT7, stu_num_bor);
	DDX_Text(pDX, IDC_EDIT8, book_num_bor);
}

BEGIN_MESSAGE_MAP(CDBTestADODlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTONRead_book, &CDBTestADODlg::OnBnClickedButtonread_book)
	ON_CBN_SELCHANGE(IDC_COMBO1, &CDBTestADODlg::OnCbnSelchangeCombo1)
	ON_BN_CLICKED(IDC_BUTTON_CHECKBOOK, &CDBTestADODlg::OnBnClickedButtonCheckbook)
	ON_BN_CLICKED(IDC_BUTTON_CHECK, &CDBTestADODlg::OnBnClickedButtonCheck)
	ON_BN_CLICKED(IDC_BUTTON_ADDSTU, &CDBTestADODlg::OnBnClickedButtonAddstu)
	ON_BN_CLICKED(IDC_BUTTON_ADDBOOK, &CDBTestADODlg::OnBnClickedButtonAddbook)
	ON_BN_CLICKED(IDC_BUTTON_BORROW, &CDBTestADODlg::OnBnClickedButtonBorrow)
	ON_BN_CLICKED(IDC_BUTTON1, &CDBTestADODlg::OnBnClickedButton1)
	ON_LBN_SELCHANGE(IDC_LIST1, &CDBTestADODlg::OnLbnSelchangeList1)
	ON_BN_CLICKED(IDC_BUTTONRead_book2, &CDBTestADODlg::OnBnClickedButtonreadbook2)
	ON_EN_CHANGE(IDC_EDIT2, &CDBTestADODlg::OnEnChangeEdit2)
	ON_BN_CLICKED(IDC_BUTTON_MODIFYSTU, &CDBTestADODlg::OnBnClickedButtonModifystu)
	ON_BN_CLICKED(IDC_BUTTON_MODIFYBOOK, &CDBTestADODlg::OnBnClickedButtonModifybook)
END_MESSAGE_MAP()


// CDBTestADoDlg 消息处理程序
CString gettime()
{
	CString time_return;
	time_t now = time(NULL);
	tm* tm_t = localtime(&now);
	std::stringstream ss;
	ss << tm_t->tm_year + 1900 << "/" << tm_t->tm_mon + 1 << "/" << tm_t->tm_mday;
	time_return = ss.str().c_str();
	return time_return;
}

BOOL CDBTestADODlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != nullptr)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	//窗体数据初始化
	m_strName = "";
	m_Age = 69;
	m_strCmd = "select * from 图书";
	m_strCmd1 = "select * from 借阅信息";
	m_strCmd2 = "select * from 学生";
	m_strSearchway1 = " ";
	UpdateData(false);
	m_ACCESSList.AddString("请进行操作");
	//使用ADO创建数据库记录集
	m_pRecordset_stu.CreateInstance(_uuidof(Recordset));
	m_pRecordset_book.CreateInstance(_uuidof(Recordset));
	m_pRecordset_mess.CreateInstance(_uuidof(Recordset));
	try
	{
		m_pRecordset_stu->Open("SELECT * FROM 学生",       //利用该SQL语言选择数据到智能指针,DemoTable为表名
			theApp.m_pConnection.GetInterfacePtr(),         //m_pConnection连接了数据库文件
			adOpenDynamic,
			adLockOptimistic,
			adCmdText);
		m_pRecordset_book->Open("SELECT * FROM 图书",       //利用该SQL语言选择数据到智能指针,DemoTable为表名
			theApp.m_pConnection.GetInterfacePtr(),         //m_pConnection连接了数据库文件
			adOpenDynamic,
			adLockOptimistic,
			adCmdText);
		m_pRecordset_mess->Open("SELECT * FROM 借阅信息",       //利用该SQL语言选择数据到智能指针,DemoTable为表名
			theApp.m_pConnection.GetInterfacePtr(),         //m_pConnection连接了数据库文件
			adOpenDynamic,
			adLockOptimistic,
			adCmdText);

	}
	catch (_com_error & e)
	{
		dump_com_error(e);
	}

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CDBTestADODlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CDBTestADODlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CDBTestADODlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

//打印调用ADO控件时产生的详细错误信息  
void CDBTestADODlg::dump_com_error(_com_error &e)
{
	CString ErrorStr;
	_bstr_t bstrSource(e.Source());
	_bstr_t bstrDescription(e.Description());
	ErrorStr.Format("/n/tADO Error/n/tCode = %08lx/n/tCode meaning = %s/n/tSource = %s/n/tDescription = %s/n/n",
		e.Error(), e.ErrorMessage(), (LPCTSTR)bstrSource, (LPCTSTR)bstrDescription);
	//在调试窗口中打印错误信息,在Release版中可用DBGView查看错误信息  
	::OutputDebugString((LPCTSTR)ErrorStr);
#ifdef _DEBUG  
	AfxMessageBox(ErrorStr, MB_OK | MB_ICONERROR);
#endif    
}

void CDBTestADODlg::OnBnClickedButtonread_book()//读取所有图书信息
{
	// TODO: 在此添加控件通知处理程序代码
	_variant_t var;
	CString strNumber, strBookName, strAuthorName, strState;
	// 清空列表框  
	m_ACCESSList.ResetContent();
	strNumber = strBookName = strAuthorName = strState = "";
	try
	{
		if (!m_pRecordset_book->BOF)
			m_pRecordset_book->MoveFirst();
		else
		{
			AfxMessageBox("表内数据为空");

			return;
		}
		// 读入库中各字段并加入列表框中  
		while (!m_pRecordset_book->adoEOF)
		{
			var = m_pRecordset_book->GetCollect("图书编号");
			if (var.vt != VT_NULL)
				strNumber = (LPCSTR)_bstr_t(var);
			var = m_pRecordset_book->GetCollect("书名");
			if (var.vt != VT_NULL)
				strBookName = (LPCSTR)_bstr_t(var);
			var = m_pRecordset_book->GetCollect("作者");
			if (var.vt != VT_NULL)
				strAuthorName = (LPCSTR)_bstr_t(var);
			var = m_pRecordset_book->GetCollect("状态");
			if (var.vt != VT_NULL)
				strState = (LPCSTR)_bstr_t(var);
			m_ACCESSList.AddString(strNumber + " --> " + strBookName + " --> " + strAuthorName + " --> " + strState);

			m_pRecordset_book->MoveNext();
		}
	}
	catch (_com_error& e)
	{
		dump_com_error(e);
	}
}



void CDBTestADODlg::OnBnClickedButton2()
{
	// TODO: 在此添加控件通知处理程序代码
}


void CDBTestADODlg::OnEnChangeEditname()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}




void CDBTestADODlg::OnCbnSelchangeCombo1()
{
	// TODO: 在此添加控件通知处理程序代码
}





void CDBTestADODlg::OnBnClickedButtonCheckbook()//找某本图书信息
{
//	CString m_strSearchway1;//查询某本图书信息——方式
//	CString m_strSearch1;//查询某本图书信息——信息
	// TODO: 在此添加控件通知处理程序代码
	UpdateData(true);
	_variant_t var;
	int flag_book = 0;
	CString strNumber, strBookName, strAuthorName, strState;
	CString  searchBookID, searchNAME,searchAU;
	CString temp;
	searchNAME = _T("书名");
	searchBookID = _T("图书编号");
	searchAU = _T("作者");

	m_ACCESSList.ResetContent();
	strNumber = strBookName = strAuthorName = strState = "";
	if (m_strSearchway1 == searchNAME)//姓名查找
	{
		try
		{
			if (!m_pRecordset_book->BOF)
				m_pRecordset_book->MoveFirst();
			else
			{
				AfxMessageBox("表内数据为空");
				return;
			}
			// 读入库中各字段并加入列表框中  
			while (!m_pRecordset_book->adoEOF)
			{
				var = m_pRecordset_book->GetCollect("书名");
				if(var.bstrVal== m_strSearch1)
				{
					flag_book = 1;
					var = m_pRecordset_book->GetCollect("图书编号");
					if (var.vt != VT_NULL)
						strNumber = (LPCSTR)_bstr_t(var);
					var = m_pRecordset_book->GetCollect("书名");
					if (var.vt != VT_NULL)
						strBookName = (LPCSTR)_bstr_t(var);
					var = m_pRecordset_book->GetCollect("作者");
					if (var.vt != VT_NULL)
						strAuthorName = (LPCSTR)_bstr_t(var);
					var = m_pRecordset_book->GetCollect("状态");
					if (var.vt != VT_NULL)
						strState = (LPCSTR)_bstr_t(var);
					m_ACCESSList.AddString(strNumber + " --> " + strBookName + " --> " + strAuthorName + " --> " + strState);
				}

				m_pRecordset_book->MoveNext();
			}
		}
		catch (_com_error& e)
		{
			dump_com_error(e);
		}
	}
	else if(m_strSearchway1 == searchAU)//作者查找
	{
		try
		{
			if (!m_pRecordset_book->BOF)
				m_pRecordset_book->MoveFirst();
			else
			{
				AfxMessageBox("表内数据为空");
				return;
			}
			// 读入库中各字段并加入列表框中  
			while (!m_pRecordset_book->adoEOF)
			{
				var = m_pRecordset_book->GetCollect("作者");
				if (var.bstrVal == m_strSearch1)
				{
					flag_book = 1;
					var = m_pRecordset_book->GetCollect("图书编号");
					if (var.vt != VT_NULL)
						strNumber = (LPCSTR)_bstr_t(var);
					var = m_pRecordset_book->GetCollect("书名");
					if (var.vt != VT_NULL)
						strBookName = (LPCSTR)_bstr_t(var);
					var = m_pRecordset_book->GetCollect("作者");
					if (var.vt != VT_NULL)
						strAuthorName = (LPCSTR)_bstr_t(var);
					var = m_pRecordset_book->GetCollect("状态");
					if (var.vt != VT_NULL)
						strState = (LPCSTR)_bstr_t(var);
					m_ACCESSList.AddString(strNumber + " --> " + strBookName + " --> " + strAuthorName + " --> " + strState);
				}

				m_pRecordset_book->MoveNext();
			}
		}
		catch (_com_error& e)
		{
			dump_com_error(e);
		}
	}
	else if (m_strSearchway1 == searchBookID)//标号查找
	{
		try
		{
			if (!m_pRecordset_book->BOF)
				m_pRecordset_book->MoveFirst();
			else
			{
				AfxMessageBox("表内数据为空");
				return;
			}
			// 读入库中各字段并加入列表框中  
			while (!m_pRecordset_book->adoEOF)
			{
				flag_book = 1;

				var = m_pRecordset_book->GetCollect("图书编号");
				temp = (LPCSTR)_bstr_t(var);
				if (temp == m_strSearch1)
				{
					var = m_pRecordset_book->GetCollect("图书编号");
					if (var.vt != VT_NULL)
						strNumber = (LPCSTR)_bstr_t(var);
					var = m_pRecordset_book->GetCollect("书名");
					if (var.vt != VT_NULL)
						strBookName = (LPCSTR)_bstr_t(var);
					var = m_pRecordset_book->GetCollect("作者");
					if (var.vt != VT_NULL)
						strAuthorName = (LPCSTR)_bstr_t(var);
					var = m_pRecordset_book->GetCollect("状态");
					if (var.vt != VT_NULL)
						strState = (LPCSTR)_bstr_t(var);
					m_ACCESSList.AddString(strNumber + " --> " + strBookName + " --> " + strAuthorName + " --> " + strState);
				}

				m_pRecordset_book->MoveNext();
			}
		}
		catch (_com_error& e)
		{
			dump_com_error(e);
		}
	}
	if (flag_book == 0)
	{
		m_ACCESSList.AddString("没有本书，查询失败！");
		return;
	}

}


void CDBTestADODlg::OnBnClickedButtonCheck()
{

	// TODO: 在此添加控件通知处理程序代码
	_variant_t var;
	_variant_t var_mess;
	CString searchStuID,searchStuID1, searchBookID, searchID,searchnow,borrowtime,returntime, searchStuname;
	int stuID[20];
	for (int i = 0; i < 20; i++) {
		stuID[i] = 0;
	}
	int searchif=0,size1 =0;//flag 找到，统计数目；
	// 清空列表框  
	m_ACCESSList.ResetContent();
	searchStuID = searchStuID1 = searchBookID = searchID = searchnow = borrowtime = returntime = searchStuname = " ";

	//更新窗体数据
	UpdateData(true);
	//执行SQL
	try
	{
		if ((m_strSearchway2 == "姓名"))
		{
			if (!m_pRecordset_stu->BOF)
				m_pRecordset_stu->MoveFirst();
			else
			{
				AfxMessageBox("表内数据为空");
				return;
			}
			// 读入库中各字段并加入列表框中  
			while (!m_pRecordset_stu->adoEOF)
			{
				var = m_pRecordset_stu->GetCollect("姓名");
				if (var.bstrVal == m_strSearch2)
				{
					var=  m_pRecordset_stu->GetCollect("学号");
					if (!m_pRecordset_mess->BOF)
						m_pRecordset_mess->MoveFirst();
					else
					{
						AfxMessageBox("表内数据为空");
						return;
					}
					while (!m_pRecordset_mess->adoEOF)
					{
						var_mess = m_pRecordset_mess->GetCollect("学号");
						if (var_mess.bstrVal == var.bstrVal)
						{
							searchif = 1;
							var_mess = m_pRecordset_mess->GetCollect("图书编号");
							if (var_mess.vt != VT_NULL)
								searchBookID = (LPCSTR)_bstr_t(var_mess);

							var_mess = m_pRecordset_mess->GetCollect("序号");
							if (var_mess.vt != VT_NULL)
								searchID = (LPCSTR)_bstr_t(var_mess);
							var_mess = m_pRecordset_mess->GetCollect("学号");
							if (var_mess.vt != VT_NULL)
								searchStuID = (LPCSTR)_bstr_t(var_mess);
							var_mess = m_pRecordset_mess->GetCollect("借阅日期");
							if (var_mess.vt != VT_NULL)
								borrowtime = (LPCSTR)_bstr_t(var_mess);
							var_mess = m_pRecordset_mess->GetCollect("归还日期");
							if (var_mess.vt != VT_NULL)
								returntime = (LPCSTR)_bstr_t(var_mess);
							else
								returntime = (LPCSTR)_bstr_t("         ");

							var_mess = m_pRecordset_mess->GetCollect("归还状态");
							if (var_mess.vt != VT_NULL)
								searchnow = (LPCSTR)_bstr_t(var_mess);
							m_ACCESSList.AddString(searchID + " --> " + searchStuID + " --> " + searchBookID + " --> " + borrowtime + " --> " + returntime + " --> " + searchnow);
						}
						m_pRecordset_mess->MoveNext();
					}
				}

				m_pRecordset_stu->MoveNext();
			}
		}

		else {
	
			if (!m_pRecordset_mess->BOF)
				m_pRecordset_mess->MoveFirst();
			else
			{
				AfxMessageBox("表内数据为空");
				return;
			}
			// 读入库中各字段并加入列表框中  
			while (!m_pRecordset_mess->adoEOF)
			{
				if (m_strSearchway2 == "学号")
				{
					var = m_pRecordset_mess->GetCollect("学号");
					if (var.vt != VT_NULL) {
						searchStuID = (LPCSTR)_bstr_t(var);
						if (searchStuID == m_strSearch2) {
							searchif = 1;
							var = m_pRecordset_mess->GetCollect("图书编号");
							if (var.vt != VT_NULL)
								searchBookID = (LPCSTR)_bstr_t(var);

							var = m_pRecordset_mess->GetCollect("序号");
							if (var.vt != VT_NULL)
								searchID = (LPCSTR)_bstr_t(var);

							var = m_pRecordset_mess->GetCollect("借阅日期");
							if (var.vt != VT_NULL)
								borrowtime = (LPCSTR)_bstr_t(var);

							var = m_pRecordset_mess->GetCollect("归还日期");
							if (var.vt != VT_NULL)
								returntime = (LPCSTR)_bstr_t(var);
							else
								returntime = (LPCSTR)_bstr_t("         ");

							var = m_pRecordset_mess->GetCollect("归还状态");
							if (var.vt != VT_NULL)
								searchnow = (LPCSTR)_bstr_t(var);
							m_ACCESSList.AddString(searchID + " --> " + searchStuID + " --> " + searchBookID + " --> " + borrowtime + " --> " + returntime + " --> " + searchnow);
						}
					}
				}
				else if (m_strSearchway2 == "图书编号")
				{
					var = m_pRecordset_mess->GetCollect("图书编号");
					if (var.vt != VT_NULL) {
						searchBookID = (LPCSTR)_bstr_t(var);
						if (searchBookID == m_strSearch2) {
							searchif = 1;
							var = m_pRecordset_mess->GetCollect("学号");
							if (var.vt != VT_NULL)
								searchStuID = (LPCSTR)_bstr_t(var);

							var = m_pRecordset_mess->GetCollect("序号");
							if (var.vt != VT_NULL)
								searchID = (LPCSTR)_bstr_t(var);

							var = m_pRecordset_mess->GetCollect("借阅日期");
							if (var.vt != VT_NULL)
								borrowtime = (LPCSTR)_bstr_t(var);

							var = m_pRecordset_mess->GetCollect("归还日期");
							if (var.vt != VT_NULL)
								returntime = (LPCSTR)_bstr_t(var);
							else
								returntime = (LPCSTR)_bstr_t("         ");

							var = m_pRecordset_mess->GetCollect("归还状态");
							if (var.vt != VT_NULL)
								searchnow = (LPCSTR)_bstr_t(var);
							m_ACCESSList.AddString(searchID + " --> " + searchStuID + " --> " + searchBookID + " --> " + borrowtime + " --> " + returntime + " --> " + searchnow);
						}
					}
				}
				else {
					searchif = 2;
				}
				m_pRecordset_mess->MoveNext();
			}
		}
	}
	catch (_com_error& e)
	{
	dump_com_error(e);
	}
	if (searchif == 0) 
		m_ACCESSList.AddString("未能查询到对应借阅信息，请检查输入姓名/学号/图书编号是否正确");
	
	if (searchif == 2) 
		m_ACCESSList.AddString("未能查询到对应借阅信息，请检查查询条件是否为姓名/学号/图书编号");

}



void CDBTestADODlg::OnBnClickedButtonAddstu()//添加学生信息
{
	//CString m_strName;//添加/修改学生信息——姓名
	//CString m_strStuNum;//添加/修改学生信息——学号
	m_ACCESSList.ResetContent();
	UINT bor_book_num = 0;//已经借阅的本数
	_variant_t var_name,var_num;
	CString temp_num;
	int is_stu = 0;//学生有吗？有的话，请点击修改按钮，本程序结束，没有的话，添加
	try
	{
		//更新数据
		UpdateData(true);
		//写入各字段值  
		if (!m_pRecordset_stu->BOF)
			m_pRecordset_stu->MoveFirst();
		else
		{
			AfxMessageBox("表内数据为空");
			return;
		}
		while (!m_pRecordset_stu->adoEOF)
		{
			var_name = m_pRecordset_stu->GetCollect("姓名");
			var_num = m_pRecordset_stu->GetCollect("学号");
			temp_num= (LPCSTR)_bstr_t(var_num);
			if (temp_num==m_strStuNum)
			{
				is_stu = 1;
				break;
			}
			m_pRecordset_stu->MoveNext();
		}
		if (is_stu == 1)
		{
			m_ACCESSList.AddString("已有该生信息，请点击修改按钮！");
			return;
		}
		else
		{
			if (!m_pRecordset_stu->adoEOF)
				m_pRecordset_stu->MoveLast();
			m_pRecordset_stu->AddNew();
			m_pRecordset_stu->PutCollect("姓名", _variant_t(m_strName));
			m_pRecordset_stu->PutCollect("学号", _variant_t(m_strStuNum));
			m_pRecordset_stu->PutCollect("已借阅本数", _variant_t(bor_book_num));
			m_pRecordset_stu->Update();
			m_ACCESSList.AddString("修改成功！");
		}
	}
	catch (_com_error& e)
	{
		dump_com_error(e);
	}


}


void CDBTestADODlg::OnBnClickedButtonAddbook()
{
	// TODO: 在此添加控件通知处理程序代码
	//CString m_strBookName;//添加/修改图书信息--书名
	//CString m_strBookAuthorName;//添加修改图书信息——作者名字
	//插入记录
	m_ACCESSList.ResetContent();
	CString state;
	state = _T("在架");
	UINT num;
	_variant_t var_name, var_num, listnum;
	int is_stu = 0;//图书有吗？有的话，请点击修改按钮，本程序结束，没有的话，添加
	try
	{
		//更新数据
		UpdateData(true);
		//写入各字段值  
		if (!m_pRecordset_book->BOF)
			m_pRecordset_book->MoveFirst();
		else
		{
			AfxMessageBox("表内数据为空");
			return;
		}
		while (!m_pRecordset_book->adoEOF)
		{
			var_name = m_pRecordset_book->GetCollect("书名");
			var_num = m_pRecordset_book->GetCollect("作者");
			if (var_name.bstrVal == m_strBookName || var_num.bstrVal == m_strBookAuthorName)
			{
				is_stu = 1;
				break;
			}
			listnum = m_pRecordset_book->GetCollect("图书编号");
			m_pRecordset_book->MoveNext();
			
		}
		if (is_stu == 1)
		{
			m_ACCESSList.AddString("已有该书籍信息，请点击修改按钮！");
			return;
		}
		else
		{
			num = listnum.intVal;
			num++;
			if (!m_pRecordset_book->adoEOF)
				m_pRecordset_book->MoveLast();
			m_pRecordset_book->AddNew();
			m_pRecordset_book->PutCollect("图书编号", _variant_t(num));
			m_pRecordset_book->PutCollect("书名", _variant_t(m_strBookName));
			m_pRecordset_book->PutCollect("作者", _variant_t(m_strBookAuthorName));
			m_pRecordset_book->PutCollect("状态", _variant_t(state));
			m_pRecordset_book->Update();
			m_ACCESSList.AddString("修改成功！");
		}
	}
	catch (_com_error& e)
	{
		dump_com_error(e);
	}
}


void CDBTestADODlg::OnBnClickedButtonAdd()
{
	// TODO: 在此添加控件通知处理程序代码
}


void CDBTestADODlg::OnBnClickedButtonBorrow()
{
	UpdateData(true);
	int flag_stu = 0;
	int book_state = 1;//书籍是否在架；
	int flag_book = 0;
	UINT bor_book_num;
	UINT mess_num;
	CString time_now;
	CString give_state = "未归还";
	//学生状态
	_variant_t var_name, var_num;
	CString var_name_str;
	m_ACCESSList.ResetContent();

	try
	{
		if (!m_pRecordset_stu->BOF)
			m_pRecordset_stu->MoveFirst();
		else
		{
			AfxMessageBox("表内数据为空");
			return;
		}
		// 读入库中各字段并加入列表框中  
		while (!m_pRecordset_stu->adoEOF)
		{
			var_name = m_pRecordset_stu->GetCollect("姓名");
			var_num = m_pRecordset_stu->GetCollect("学号");
			if (var_name.bstrVal == stu_name_bor && var_num.intVal == stu_num_bor)
			{
				flag_stu = 1;
				break;
			}
			m_pRecordset_stu->MoveNext();
		}
	}
	catch (_com_error& e)
	{
		dump_com_error(e);
	}
	if (flag_stu == 0)
	{
		m_ACCESSList.AddString("没有该学生的信息，借阅失败！");
		return;
	}

	_variant_t book_bor;
	_variant_t var_book_state;
	CString book_state_bor;
	book_state_bor = _T("外借");
	try
	{
		if (!m_pRecordset_book->BOF)
			m_pRecordset_book->MoveFirst();
		else
		{
			AfxMessageBox("表内数据为空");
			return;
		}
		// 读入库中各字段并加入列表框中  
		while (!m_pRecordset_book->adoEOF)
		{
			book_bor = m_pRecordset_book->GetCollect("图书编号");
			if (book_bor.intVal == book_num_bor)
			{
				flag_book = 1;
				var_book_state = m_pRecordset_book->GetCollect("状态");
				if (var_book_state.bstrVal == book_state_bor)
					book_state = 0;
				break;
			}
			m_pRecordset_book->MoveNext();
		}
	}
	catch (_com_error& e)
	{
		dump_com_error(e);
	}
	if (flag_book == 0)
	{
		m_ACCESSList.AddString("没有本书，借阅失败！");
		return;
	}
	if (!book_state)
	{
		m_ACCESSList.AddString("书籍不在架，借阅失败！");
		return;
	}

	m_ACCESSList.AddString("借阅成功！");

	try
	{

		m_pRecordset_book->PutCollect("状态", _variant_t(book_state_bor));
		bor_book_num = m_pRecordset_stu->GetCollect("已借阅本数").intVal;
		bor_book_num++;
		m_pRecordset_stu->PutCollect("已借阅本数", _variant_t(bor_book_num));
		m_pRecordset_stu->Update();
		m_pRecordset_book->Update();
		while (!m_pRecordset_mess->adoEOF)
		{
			mess_num = m_pRecordset_mess->GetCollect("序号").intVal;
			m_pRecordset_mess->MoveNext();
		}
		mess_num++;
		m_pRecordset_mess->AddNew();
		m_pRecordset_mess->PutCollect("序号", _variant_t(mess_num));
		m_pRecordset_mess->PutCollect("学号", _variant_t(stu_num_bor));

		m_pRecordset_mess->PutCollect("图书编号", _variant_t(book_num_bor));
		time_now = gettime();
		m_pRecordset_mess->PutCollect("借阅日期", _variant_t(time_now));
		m_pRecordset_mess->PutCollect("归还状态", _variant_t(give_state));
		m_pRecordset_mess->Update();

	}
	catch (_com_error& e)
	{
		dump_com_error(e);
	}
}


void CDBTestADODlg::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	// TODO: 在此添加控件通知处理程序代码
	UpdateData(true);
	int flag_stu = 0;
	int book_state = 1;//书籍是否在架；
	int flag_book = 0;
	UINT bor_book_num;
	UINT mess_num;
	CString time_now;
	CString temp_num;
	CString give_state = "已归还";
	_variant_t var_name, var_num, var_book;
	CString var_name_str;
	m_ACCESSList.ResetContent();

	try
	{
		if (!m_pRecordset_stu->BOF)
			m_pRecordset_stu->MoveFirst();
		else
		{
			AfxMessageBox("表内数据为空");
			return;
		}
		// 读入库中各字段并加入列表框中  
		while (!m_pRecordset_stu->adoEOF)
		{
			var_name = m_pRecordset_stu->GetCollect("姓名");
			var_num = m_pRecordset_stu->GetCollect("学号");
			//temp_num = (LPCSTR)_bstr_t(var_num);
			if (var_name.bstrVal == stu_name_bor && var_num.intVal == stu_num_bor)
			{
				flag_stu = 1;
				break;
			}
			m_pRecordset_stu->MoveNext();
		}
	}
	catch (_com_error& e)
	{
		dump_com_error(e);
	}
	if (flag_stu == 0)
	{
		m_ACCESSList.AddString("没有本书！");
		return;
	}

	_variant_t book_bor;
	_variant_t var_book_state;
	CString book_state_bor;
	book_state_bor = _T("在架");
	try
	{
		if (!m_pRecordset_book->BOF)
			m_pRecordset_book->MoveFirst();
		else
		{
			AfxMessageBox("表内数据为空");
			return;
		}
		// 读入库中各字段并加入列表框中  
		while (!m_pRecordset_book->adoEOF)
		{
			book_bor = m_pRecordset_book->GetCollect("图书编号");
			if (book_bor.intVal == book_num_bor)
			{
				flag_book = 1;
				var_book_state = m_pRecordset_book->GetCollect("状态");
				if (var_book_state.bstrVal == book_state_bor)
					book_state = 0;
				break;
			}
			m_pRecordset_book->MoveNext();
		}
	}
	catch (_com_error& e)
	{
		dump_com_error(e);
	}
	if (flag_book == 0)
	{
		m_ACCESSList.AddString("没有本书，归还失败！");
		return;
	}
	if (!book_state)
	{
		m_ACCESSList.AddString("书籍在架，归还失败！");
		return;
	}

	m_ACCESSList.AddString("归还成功！");

	try
	{

		m_pRecordset_book->PutCollect("状态", _variant_t(book_state_bor));
		bor_book_num = m_pRecordset_stu->GetCollect("已借阅本数").intVal;
		bor_book_num--;
		m_pRecordset_stu->PutCollect("已借阅本数", _variant_t(bor_book_num));
		m_pRecordset_stu->Update();
		m_pRecordset_book->Update();
		if (!m_pRecordset_mess->BOF)
			m_pRecordset_mess->MoveFirst();
		else
		{
			AfxMessageBox("表内数据为空");
			return;
		}
		while (!m_pRecordset_mess->adoEOF)
		{
			var_num = m_pRecordset_mess->GetCollect("学号");
			var_book = m_pRecordset_mess->GetCollect("图书编号");
			if (var_num.intVal == stu_num_bor && var_book.intVal == book_num_bor)
			{
				break;
			}
			m_pRecordset_mess->MoveNext();
		}
		time_now = gettime();
		m_pRecordset_mess->PutCollect("归还日期", _variant_t(time_now));
		m_pRecordset_mess->PutCollect("归还状态", _variant_t(give_state));
		m_pRecordset_mess->Update();

	}
	catch (_com_error& e)
	{
		dump_com_error(e);
	}
}


void CDBTestADODlg::OnBnClickedButtonAddstu2()
{
	// TODO: 在此添加控件通知处理程序代码
}


void CDBTestADODlg::OnBnClickedButtonModify()
{
	// TODO: 在此添加控件通知处理程序代码
}


void CDBTestADODlg::OnLbnSelchangeList1()
{
	// TODO: 在此添加控件通知处理程序代码
}


void CDBTestADODlg::OnBnClickedButtonreadbook2()//读取所有学生信息
{
	// TODO: 在此添加控件通知处理程序代码
	UpdateData(true);
	_variant_t var;
	CString strName, strName1, strAge;
	// 清空列表框  
	m_ACCESSList.ResetContent();
	strName = strAge = strName1 = "";
	try
	{
		if (!m_pRecordset_stu->BOF)
			m_pRecordset_stu->MoveFirst();
		else
		{
			AfxMessageBox("表内数据为空");
			return;
		}
		// 读入库中各字段并加入列表框中  
		while (!m_pRecordset_stu->adoEOF)
		{
			var = m_pRecordset_stu->GetCollect("姓名");
			if (var.vt != VT_NULL)
				strName = (LPCSTR)_bstr_t(var);
			var = m_pRecordset_stu->GetCollect("学号");
			if (var.vt != VT_NULL)
				strAge = (LPCSTR)_bstr_t(var);
			var = m_pRecordset_stu->GetCollect("已借阅本数");
			if (var.vt != VT_NULL)
				strName1 = (LPCSTR)_bstr_t(var);
			m_ACCESSList.AddString(strName + " --> " + strAge + " --> " + strName1);

			m_pRecordset_stu->MoveNext();
		}
	}
	catch (_com_error& e)
	{
		dump_com_error(e);
	}
}





void CDBTestADODlg::OnEnChangeEdit2()
{
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}


void CDBTestADODlg::OnBnClickedButtonModifystu()
{
	// TODO: 在此添加控件通知处理程序代码
		//CString m_strName;//添加/修改学生信息——姓名
	//CString m_strStuNum;//添加/修改学生信息——学号
	m_ACCESSList.ResetContent();
	UINT bor_book_num = 0;//已经借阅的本数
	_variant_t var_name, var_num;
	CString temp_num;
	int is_stu = 0;//学生有吗？有的话，请点击修改按钮，没有的话，添加
	try
	{
		//更新数据
		UpdateData(true);
		//写入各字段值  
		if (!m_pRecordset_stu->BOF)
			m_pRecordset_stu->MoveFirst();
		else
		{
			AfxMessageBox("表内数据为空");
			return;
		}
		while (!m_pRecordset_stu->adoEOF)
		{
			var_name = m_pRecordset_stu->GetCollect("姓名");
			var_num = m_pRecordset_stu->GetCollect("学号");
			temp_num = (LPCSTR)_bstr_t(var_num);
			if ( temp_num == m_strStuNum)
			{
				is_stu = 1;
				break;
			}
			m_pRecordset_stu->MoveNext();
		}
		if (is_stu == 0)
		{
			m_ACCESSList.AddString("没有该生信息，请添加！");
			return;
		}

		else
		{
			m_pRecordset_stu->PutCollect("姓名", _variant_t(m_strName));
			m_pRecordset_stu->PutCollect("学号", _variant_t(m_strStuNum));
			m_pRecordset_stu->Update();
			m_ACCESSList.AddString("修改成功！");
		}
	}
	catch (_com_error& e)
	{
		dump_com_error(e);
	}
}


void CDBTestADODlg::OnBnClickedButtonModifybook()
{
	//CString m_strBookName;//添加/修改图书信息--书名
	//CString m_strBookAuthorName;//添加修改图书信息——作者名字
	//插入记录
	m_ACCESSList.ResetContent();
	CString state;
	state = _T("在架");
	UINT num;
	_variant_t var_name, var_num, listnum;
	int is_stu = 0;//图书有吗？有的话，请点击修改按钮，本程序结束，没有的话，添加
	try
	{
		//更新数据
		UpdateData(true);
		//写入各字段值  
		if (!m_pRecordset_book->BOF)
			m_pRecordset_book->MoveFirst();
		else
		{
			AfxMessageBox("表内数据为空");
			return;
		}
		while (!m_pRecordset_book->adoEOF)
		{
			var_name = m_pRecordset_book->GetCollect("书名");
			var_num = m_pRecordset_book->GetCollect("作者");
			if (var_name.bstrVal == m_strBookName || var_num.bstrVal == m_strBookAuthorName)
			{
				is_stu = 1;
				break;
			}
			m_pRecordset_book->MoveNext();
		}
		if (is_stu == 0)
		{
			m_ACCESSList.AddString("没有该书信息，请添加！");
			return;
		}

		else
		{
			m_pRecordset_book->PutCollect("书名", _variant_t(m_strBookName));
			m_pRecordset_book->PutCollect("作者", _variant_t(m_strBookAuthorName));
			m_pRecordset_book->Update();
			m_ACCESSList.AddString("修改成功！");
		}
	}
	catch (_com_error& e)
	{
		dump_com_error(e);
	}
}
