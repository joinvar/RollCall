
// RollCall.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CRollCallApp: 
// �йش����ʵ�֣������ RollCall.cpp
//

class CRollCallApp : public CWinApp
{
public:
	CRollCallApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CRollCallApp theApp;