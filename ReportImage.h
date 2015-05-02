#ifndef REPORTIMAGE_H
#define REPORTIMAGE_H

#include "afxwin.h"

class CReportImage :
	public CStatic
{
public:
	CReportImage(void);
	~CReportImage(void);
	DECLARE_MESSAGE_MAP()
	afx_msg void OnPaint();
};

#endif