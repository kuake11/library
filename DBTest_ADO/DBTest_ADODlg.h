
// DBTest_ADoDlg.h: 头文件
//

#pragma once


// CDBTestADoDlg 对话框
class CDBTestADODlg : public CDialogEx
{
// 构造
public:
	CDBTestADODlg(CWnd* pParent = nullptr);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_DBTEST_ADO_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	afx_msg void OnBnClickedButtonread_book();
	CListBox m_ACCESSList;
	CString m_strName;//添加/修改学生信息——姓名
	CString m_strStuNum;//添加/修改学生信息——学号
	CString m_strSearchway1;//查询某本图书信息——方式
	CString m_strSearch1;//查询某本图书信息——信息
	CString m_strSearchway2;
	CString m_strSearch2;
	CString m_strBookName;//添加/修改图书信息--书名
	CString m_strBookAuthorName;//添加修改图书信息——作者名字
	UINT m_Age;	
	CString m_strCmd;
	CString m_strCmd1;
	CString m_strCmd2;
	CString stu_name_bor;//借阅功能
	UINT stu_num_bor;//借阅功能
	UINT book_num_bor;//借阅功能

public:
	static void dump_com_error(_com_error &e);  //数据库操作的错误信息
	_RecordsetPtr m_pRecordset_stu;                 //智能指针，student
	_RecordsetPtr m_pRecordset_book;//智能指針，book
	_RecordsetPtr m_pRecordset;
	_RecordsetPtr m_pRecordset_mess;//智能指針，信息
	afx_msg void OnBnClickedButton2();
	afx_msg void OnEnChangeEditname();
	afx_msg void OnEnChangeEdit1();
	afx_msg void OnCbnSelchangeCombo1();
	afx_msg void OnEnChangeEdit7();
	afx_msg void OnBnClickedButtonCheckbook();
	afx_msg void OnBnClickedButtonCheck();
	afx_msg void OnBnClickedButtonAddstu();
	afx_msg void OnBnClickedButtonAddbook();
	afx_msg void OnBnClickedButtonAdd();
	afx_msg void OnBnClickedButtonBorrow();
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButtonAddstu2();
	afx_msg void OnBnClickedButtonModify();
	afx_msg void OnLbnSelchangeList1();
	afx_msg void OnBnClickedButtonreadbook2();
	afx_msg void OnEnChangeEdit2();
	afx_msg void OnBnClickedButtonModifystu();
	afx_msg void OnBnClickedButtonModifybook();
};
