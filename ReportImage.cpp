#include "StdAfx.h"
#include ".\reportimage.h"

CReportImage::CReportImage(void)
{
}

CReportImage::~CReportImage(void)
{
}
BEGIN_MESSAGE_MAP(CReportImage, CStatic)
	ON_WM_PAINT()
END_MESSAGE_MAP()

void CReportImage::OnPaint()
{
	CPaintDC dc(this); // device context for painting
	HBITMAP hbmp = GetBitmap();
	CDC memDC;
	memDC.CreateCompatibleDC(&dc);
	memDC.SelectObject(hbmp);
	CRect rect;
	GetClientRect(&rect);
	int y;
	
	BITMAP bm;
	::GetObject(hbmp, sizeof(BITMAP), &bm);

	float fx = 1;
	if(rect.Width() < bm.bmWidth)
	{
		fx = (float)(bm.bmWidth/rect.Width());
	}
	y = (int)((rect.Height()-bm.bmHeight*fx)/2);
	dc.SetStretchBltMode(HALFTONE);
	dc.StretchBlt(rect.left, y, rect.Width(),(int)( bm.bmHeight*fx), &memDC, 0, 0, bm.bmWidth, bm.bmHeight, SRCCOPY);
	memDC.DeleteDC();
}
