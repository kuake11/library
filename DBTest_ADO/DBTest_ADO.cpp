
// DBTest_ADo.cpp: 定义应用程序的类行为。
//

#include "stdafx.h"
#include "DBTest_ADo.h"
#include "DBTest_ADoDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CDBTestADoApp

BEGIN_MESSAGE_MAP(CDBTestADOApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CDBTestADoApp 构造

CDBTestADOApp::CDBTestADOApp()
{
	// 支持重新启动管理器
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;

	// TODO: 在此处添加构造代码，
	// 将所有重要的初始化放置在 InitInstance 中
}


// 唯一的 CDBTestADoApp 对象

CDBTestADOApp theApp;


// CDBTestADoApp 初始化

BOOL CDBTestADOApp::InitInstance()
{
	//初始化COM,创建ADO连接等操作  
	AfxOleInit();
	m_pConnection.CreateInstance(__uuidof(Connection));
	//在ADO操作中建议语句中要常用try...catch()来捕获错误信息，  
	try
	{
		//Jet引擎打开数据库，Demo.mdb为数据库名，仅能打开低版本的ACCESS文件，且需要项目配置为x86
		//m_pConnection->Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Demo.mdb", "", "", adModeUnknown);

		//ACE引擎打开数据库，Demo.accdb为数据库名，mdb和accdb类型的ACCESS文件都能打开。
		//如果操作系统为64位系统，需要下载安装AccessDatabaseEngine_X64.exe，
		//下载地址为：https://www.microsoft.com/en-us/download/details.aspx?id=13255，选择64位程序，
		//需要项目配置为x64。
		m_pConnection->Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=book-manage.accdb", "", "", adModeUnknown);

		//获取数据库的表格名称
		_RecordsetPtr pRstSchema = NULL;
		pRstSchema = m_pConnection->OpenSchema(adSchemaTables);
		while (!(pRstSchema->adoEOF)) 
		{
			_bstr_t table_name = pRstSchema->Fields->GetItem("TABLE_NAME")->Value;
			printf("Table Name: %s\n", (LPCSTR)table_name);
			_bstr_t table_type = pRstSchema->Fields->GetItem("TABLE_TYPE")->Value;
			printf("Table type: %s\n\n", (LPCSTR)table_type);
			pRstSchema->MoveNext();
		}

	}
	catch (_com_error &e)
	{
		//调用在dump_com_error中打印错误信息的静态函数  
		CDBTestADODlg::dump_com_error(e);
		return FALSE;
	}


	// 如果一个运行在 Windows XP 上的应用程序清单指定要
	// 使用 ComCtl32.dll 版本 6 或更高版本来启用可视化方式，
	//则需要 InitCommonControlsEx()。  否则，将无法创建窗口。
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// 将它设置为包括所有要在应用程序中使用的
	// 公共控件类。
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();


	AfxEnableControlContainer();

	// 创建 shell 管理器，以防对话框包含
	// 任何 shell 树视图控件或 shell 列表视图控件。
	CShellManager *pShellManager = new CShellManager;

	// 激活“Windows Native”视觉管理器，以便在 MFC 控件中启用主题
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));

	// 标准初始化
	// 如果未使用这些功能并希望减小
	// 最终可执行文件的大小，则应移除下列
	// 不需要的特定初始化例程
	// 更改用于存储设置的注册表项
	// TODO: 应适当修改该字符串，
	// 例如修改为公司或组织名
	SetRegistryKey(_T("应用程序向导生成的本地应用程序"));

	CDBTestADODlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: 在此放置处理何时用
		//  “确定”来关闭对话框的代码
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: 在此放置处理何时用
		//  “取消”来关闭对话框的代码
	}
	else if (nResponse == -1)
	{
		TRACE(traceAppMsg, 0, "警告: 对话框创建失败，应用程序将意外终止。\n");
		TRACE(traceAppMsg, 0, "警告: 如果您在对话框上使用 MFC 控件，则无法 #define _AFX_NO_MFC_CONTROLS_IN_DIALOGS。\n");
	}

	// 删除上面创建的 shell 管理器。
	if (pShellManager != nullptr)
	{
		delete pShellManager;
	}

#if !defined(_AFXDLL) && !defined(_AFX_NO_MFC_CONTROLS_IN_DIALOGS)
	ControlBarCleanUp();
#endif

	// 由于对话框已关闭，所以将返回 FALSE 以便退出应用程序，
	//  而不是启动应用程序的消息泵。
	return FALSE;
}

//  CDBTestADoApp 退出时

BOOL CDBTestADOApp::ExitInstance()
{
	//退出应用时，关闭数据库连接
	if (m_pConnection->State) m_pConnection->Close();
	m_pConnection = NULL;
	return true;
}

