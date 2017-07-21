
// DlgProxy.cpp : 实现文件
//

#include "stdafx.h"
#include "ReadXls.h"
#include "DlgProxy.h"
#include "ReadXlsDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CReadXlsDlgAutoProxy

IMPLEMENT_DYNCREATE(CReadXlsDlgAutoProxy, CCmdTarget)

CReadXlsDlgAutoProxy::CReadXlsDlgAutoProxy()
{
	EnableAutomation();
	
	// 为使应用程序在自动化对象处于活动状态时一直保持 
	//	运行，构造函数调用 AfxOleLockApp。
	AfxOleLockApp();

	// 通过应用程序的主窗口指针
	//  来访问对话框。  设置代理的内部指针
	//  指向对话框，并设置对话框的后向指针指向
	//  该代理。
	ASSERT_VALID(AfxGetApp()->m_pMainWnd);
	if (AfxGetApp()->m_pMainWnd)
	{
		ASSERT_KINDOF(CReadXlsDlg, AfxGetApp()->m_pMainWnd);
		if (AfxGetApp()->m_pMainWnd->IsKindOf(RUNTIME_CLASS(CReadXlsDlg)))
		{
			m_pDialog = reinterpret_cast<CReadXlsDlg*>(AfxGetApp()->m_pMainWnd);
			m_pDialog->m_pAutoProxy = this;
		}
	}
}

CReadXlsDlgAutoProxy::~CReadXlsDlgAutoProxy()
{
	// 为了在用 OLE 自动化创建所有对象后终止应用程序，
	//	析构函数调用 AfxOleUnlockApp。
	//  除了做其他事情外，这还将销毁主对话框
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CReadXlsDlgAutoProxy::OnFinalRelease()
{
	// 释放了对自动化对象的最后一个引用后，将调用
	// OnFinalRelease。  基类将自动
	// 删除该对象。  在调用该基类之前，请添加您的
	// 对象所需的附加清理代码。

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CReadXlsDlgAutoProxy, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CReadXlsDlgAutoProxy, CCmdTarget)
END_DISPATCH_MAP()

// 注意: 我们添加了对 IID_IReadXls 的支持
//  以支持来自 VBA 的类型安全绑定。  此 IID 必须同附加到 .IDL 文件中的
//  调度接口的 GUID 匹配。

// {6B3EFF69-A4AD-4306-BD9C-D7863A4202B2}
static const IID IID_IReadXls =
{ 0x6B3EFF69, 0xA4AD, 0x4306, { 0xBD, 0x9C, 0xD7, 0x86, 0x3A, 0x42, 0x2, 0xB2 } };

BEGIN_INTERFACE_MAP(CReadXlsDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CReadXlsDlgAutoProxy, IID_IReadXls, Dispatch)
END_INTERFACE_MAP()

// IMPLEMENT_OLECREATE2 宏在此项目的 StdAfx.h 中定义
// {174D6E32-4900-420B-9006-1C6D299D4ABC}
IMPLEMENT_OLECREATE2(CReadXlsDlgAutoProxy, "ReadXls.Application", 0x174d6e32, 0x4900, 0x420b, 0x90, 0x6, 0x1c, 0x6d, 0x29, 0x9d, 0x4a, 0xbc)


// CReadXlsDlgAutoProxy 消息处理程序
