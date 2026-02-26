#include "pch.h"
#include "framework.h"
#include "MarkdownEditor.h"
#include "MarkdownEditorDlg.h"
#include <shellapi.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

BEGIN_MESSAGE_MAP(CMarkdownEditorApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()

CMarkdownEditorApp::CMarkdownEditorApp()
{
	m_dwRestartManagerSupportFlags = AFX_RESTART_MANAGER_SUPPORT_RESTART;
}

CMarkdownEditorApp theApp;

BOOL CMarkdownEditorApp::InitInstance()
{
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	CWinApp::InitInstance();

	// Initialize RichEdit control.
	// We use RICHEDIT50W (msftedit.dll) for better Unicode/emoji support.
	// AfxInitRichEdit2() should load it, but do an explicit LoadLibrary so dialog creation won't fail silently.
	::LoadLibraryW(L"Msftedit.dll");
	if (!AfxInitRichEdit2())
	{
		AfxMessageBox(_T("RichEdit 初始化失败（AfxInitRichEdit2）。\r\n请检查系统组件 Msftedit.dll / Riched20.dll 是否存在。"),
			MB_OK | MB_ICONERROR);
		return FALSE;
	}

	AfxEnableControlContainer();

	CShellManager *pShellManager = new CShellManager;

	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerWindows));

	SetRegistryKey(_T("MyTypora Markdown Editor"));

	CMarkdownEditorDlg dlg;
	{
		int argc = 0;
		LPWSTR* argv = ::CommandLineToArgvW(::GetCommandLineW(), &argc);
		if (argv != nullptr)
		{
			for (int i = 1; i < argc; ++i)
			{
				CString arg(argv[i]);
				arg.Trim();
				if (arg.IsEmpty())
					continue;
				if (arg[0] == _T('-') || arg[0] == _T('/'))
					continue;
				dlg.SetStartupOpenFilePath(arg);
				break;
			}
			::LocalFree(argv);
		}
	}
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == -1)
	{
		DWORD err = ::GetLastError();
		CString msg;
		msg.Format(_T("程序启动失败：主对话框创建失败（DoModal 返回 -1）。\r\nGetLastError=0x%08X"), err);
		AfxMessageBox(msg, MB_OK | MB_ICONERROR);
	}

	if (pShellManager != nullptr)
	{
		delete pShellManager;
	}

#if !defined(_AFXDLL) && !defined(_AFX_NO_MFC_CONTROLS_IN_DIALOGS)
	ControlBarCleanUp();
#endif

	return FALSE;
}

