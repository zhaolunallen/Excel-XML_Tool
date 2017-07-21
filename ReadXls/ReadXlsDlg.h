
// ReadXlsDlg.h : 头文件
//

#pragma once
//#import "msxml6.dll"
#include "afxwin.h"
#include "Excel.h"
#include <map>
#include <vector>
#include <string>
// #include <msxml6.h>
// #include <comutil.h>
#include <locale>
#include "tinystr.h"
#include "tinyxml.h"
// 
// #pragma comment(lib, "comsuppwd.lib") 

using namespace std;
//using namespace MSXML2;

class CReadXlsDlgAutoProxy;
struct CSheetSize
{
	int row;
	int col;
};

struct CExcelSheet
{
	vector<vector<CString>> vecSheet;
	CString strSheetName;
	int BoHaveTitle = false;
	int BoIsSubTable = false;
};

struct CCharSheet
{
	vector<vector<char*>> vecSheet;
	char* strSheetName;
	int BoHaveTitle = false;
	int BoIsSubTable = false;
};

// CReadXlsDlg 对话框
class CReadXlsDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CReadXlsDlg);
	friend class CReadXlsDlgAutoProxy;

// 构造
public:
	CReadXlsDlg(CWnd* pParent = NULL);	// 标准构造函数
	virtual ~CReadXlsDlg();

// 对话框数据
	enum { IDD = IDD_READXLS_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	CReadXlsDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	DECLARE_MESSAGE_MAP()


	char* Unicode2Utf8(const char* unicode);
	char* Unicode2Ansi(const char* AnsiStr);
	char* Ansi2Unicode(const char* str);
	char* Utf82Unicode(const char* str);
	char* ConvertAnsiToUtf8(const char* str);
	char* ConvertUtf8ToAnsi(const char* str);
	CString __CStringConvertAnsiToUtf8(CString& strSrc);
	
	char* __CString2Constchar(CString& strSrc);
	CString  __Constchar2CString(const char* strSrc);
	void __CString2ConstcharWithAnsiToUtf8(CString& strSrc, char * outbuf);
	void readExcelAsData(CString& strInputFile,CString& strOutputPath);
	void readExcelAsData3(CString& strInputFile, CString& strOutputPath);
	void readExcelAsData4(CString& strInputFile, CString& strOutputPath);
	void __CreateXmlFile(CString& strPath);
	void __CreateXLsFile(CString& strPath);
	void __WriteItem(CStdioFile& obFile, size_t nID, CString strParentID);
	void __WriteItem3(CStdioFile& obFile, size_t nID, CString strParentID, CString strParentAttri);
	void __CreateXmlFile2(CString& strPath);
	void __CreateXmlFile3(CString& strPath);
	bool isSummaryTableMode();
	CString __GetFileName(CString strPath);
public:
	afx_msg void OnBnClickedToxml();
	// 退出程序
	CButton m_Exit;
	CButton m_btnToXml;
	afx_msg void OnBnClickedExit();
	CButton m_btnOpenFile;
	CButton m_btnSavePath;
	afx_msg void OnBnClickedOpenfile();
	CEdit m_editInputPath;
	CEdit m_editOutputPath;
	afx_msg void OnBnClickedSavepath();
	afx_msg void OnBnClickedToxml2();
	//afx_msg void OnBnClickedButton2();

private:
	char* m_oldLocale;
	Excel m_obExcel;
	CString m_strInputName;
	vector<CString> m_vecSheetName;
	vector<char *> m_vecCharSheetName;
	map<CString, vector<vector<CString>>> m_mapSheetList;
	map<char*, vector<vector<char* >>> m_mapCharSheetList;
	map<CString, CSheetSize> m_mapSheetSize;
	TiXmlDocument m_obXml;
	CString m_stroutputName;
	map<CString, CExcelSheet> m_XlsToXml_mapSheetList;
	map<char*, CCharSheet> m_XlsToXml_mapCharSheetList;

public:
	afx_msg void OnBnClickedButton3();
};
