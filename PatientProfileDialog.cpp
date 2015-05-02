#include "StdAfx.h"
#include ".\patientprofiledialog.h"

static int _OnPaneSelectCallbackFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPatientProfileDialog*)pWnd)->OnPaneSelectCallback(nOldItem, nNewItem);
	return 0;
}

static int _OnDocNoCheckValueFnc(CWnd *pWnd){
	return ((CPatientProfileDialog*) pWnd)->OnDocNoCheckValue();
}
CPatientProfileDialog::CPatientProfileDialog(CWnd* pParent, long nDocumentNo):
	CGuiDialog(pParent)
{
	m_nDocumentNo = nDocumentNo;
	m_nDlgWidth = 960;
	m_nDlgHeight = 620;

}

CPatientProfileDialog::~CPatientProfileDialog(void)
{
}


void CPatientProfileDialog::OnCreateComponents(){
	m_wndPane.Create(this, 0, 0, 150, 610);
	m_wndDrugLogs.Create(this, CRect(155, 0, 810, 610));
	m_wndDocNoLabel.Create(this, _T("Doc No"), 155, 615, 235, 640);
	m_wndDocNo.Create(this,240, 615, 320, 640); 
	
}

void CPatientProfileDialog::OnDoDataExchange(CDataExchange *pDX){
	DDX_Text(pDX, m_wndDocNo.GetDlgCtrlID(), m_nDocumentNo);
}

void CPatientProfileDialog::OnInitializeComponents()
{
	m_wndPane.AddCaption(_T("Drug Logs"));
	m_wndPane.Add(_T("7 Day"), _T(""), NULL);
	m_wndPane.Add(_T("15 Day"), _T(""), NULL);
	m_wndPane.Add(_T("30 Day"), _T(""), NULL);
	m_wndPane.Add(_T("60 Day"), _T(""), NULL);
	m_wndPane.Add(_T("All"), _T(""), NULL);
	m_wndPane.AddCaption(_T("Treat Information"));
	m_wndPane.Add(_T("Test"), _T(""), NULL);
	m_wndPane.Add(_T("PACS"), _T(""), NULL);
	m_wndPane.Add(_T("Surgery"), _T(""), NULL);
	m_wndDrugLogs.OnInitDialog();
}

void CPatientProfileDialog::OnSetWindowEvents()
{

	m_wndPane.SetCallback(_OnPaneSelectCallbackFnc);
	m_wndPane.SetCurSel(1);
	m_wndDrugLogs.m_nDocumentNo = m_nDocumentNo;
	m_wndDrugLogs.CreateColumns(_T("DRUGLOG"), 7);
	m_wndDrugLogs.Refresh();
	m_wndDocNo.SetEvent(WE_CHECKVALUE, _OnDocNoCheckValueFnc);

}

int CPatientProfileDialog::OnDocNoCheckValue(){
	m_wndPane.SetCurSel(1);
	m_wndDrugLogs.m_nDocumentNo = m_nDocumentNo;
	m_wndDrugLogs.CreateColumns(_T("DRUGLOG"), 7);
	m_wndDrugLogs.Refresh();
	return 0;
}

void CPatientProfileDialog::OnPaneSelectCallback(int nOldItem, int nNewItem)
{
	if(nNewItem < 0)
		return;

	m_wndDrugLogs.m_nDocumentNo = m_nDocumentNo;
	switch(nNewItem)
	{
	case 1:
		m_wndDrugLogs.CreateColumns(_T("DRUGLOG"), 7);
		break;
	case 2:
		m_wndDrugLogs.CreateColumns(_T("DRUGLOG"), 15);
		break;
	case 3:
		m_wndDrugLogs.CreateColumns(_T("DRUGLOG"), 30);
		break;
	case 4:
		m_wndDrugLogs.CreateColumns(_T("DRUGLOG"), 60);
		break;
	case 5:
		m_wndDrugLogs.CreateColumns(_T("DRUGLOG"), 0);
		break;
	case 7:
		m_wndDrugLogs.CreateColumns(_T("TEST"));
		break;
	case 8:
		m_wndDrugLogs.CreateColumns(_T("PACS"));
		break;
	case 9:
		m_wndDrugLogs.CreateColumns(_T("SURGERY"));
		break;
	}
	
	m_wndDrugLogs.m_nDocumentNo = m_nDocumentNo;
	m_wndDrugLogs.Refresh();
//	m_wndDrugLogs.ShowWindow(SW_SHOW);

}