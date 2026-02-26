#pragma once

#ifndef __AFXWIN_H__
	#error "include 'pch.h' before including this file for PCH"
#endif

#include "resource.h"

class CMarkdownEditorApp : public CWinApp
{
public:
	CMarkdownEditorApp();

public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()
};

extern CMarkdownEditorApp theApp;
