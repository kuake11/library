
// DBTest_ADo.h: PROJECT_NAME 应用程序的主头文件
//

#pragma once

#ifndef __AFXWIN_H__
	#error "在包含此文件之前包含“stdafx.h”以生成 PCH 文件"
#endif

#include "resource.h"		// 主符号


// CDBTestADoApp:
// 有关此类的实现，请参阅 DBTest_ADo.cpp
//

class CDBTestADOApp : public CWinApp
{
public:
	CDBTestADOApp();

// 重写
public:
	virtual BOOL InitInstance();
	virtual BOOL ExitInstance();
	_ConnectionPtr m_pConnection;  //用于建立与数据库的连接

// 实现

	DECLARE_MESSAGE_MAP()
};

extern CDBTestADOApp theApp;
