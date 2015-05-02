#include "stdafx.h"
#include "PMAdvanceAmount.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include <typeinfo.h>
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CPMAdvanceAmount *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMAdvanceAmount* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CPMAdvanceAmount *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMAdvanceAmount *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMAdvanceAmount *)pWnd)->OnToDateCheckValue();
} 
static void _OnSourceSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMAdvanceAmount* )pWnd)->OnSourceSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSourceSelendokFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnSourceSelendok();
}
/*static void _OnSourceSetfocusFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnSourceKillfocus();
}*/
/*static void _OnSourceKillfocusFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnSourceKillfocus();
}*/
static long _OnSourceLoadDataFnc(CWnd *pWnd){
	return ((CPMAdvanceAmount *)pWnd)->OnSourceLoadData();
}
/*static void _OnSourceAddNewFnc(CWnd *pWnd){
	((CPMAdvanceAmount *)pWnd)->OnSourceAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMAdvanceAmount *pVw = (CPMAdvanceAmount *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMAdvanceAmount *pVw = (CPMAdvanceAmount *)pWnd;
	pVw->OnExportSelect();
} 
CPMAdvanceAmount::CPMAdvanceAmount(CWnd *pParent){

	m_nDlgWidth = 500;
	m_nDlgHeight = 160;
	SetDefaultValues();
}
CPMAdvanceAmount::~CPMAdvanceAmount(){
}
void CPMAdvanceAmount::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 490, 120);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 245, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 250, 30, 330, 55);
	m_wndReportPeriod.Create(this,335, 30, 485, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 245, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 60, 330, 85);
	m_wndToDate.Create(this,335, 60, 485, 85); 
	m_wndSourceLabel.Create(this, _T("Source"), 10, 90, 90, 115);
	m_wndSource.Create(this,95, 90, 485, 115); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 300, 125, 410, 150);
	m_wndExport.Create(this, _T("&Export"), 415, 125, 490, 150);

}
void CPMAdvanceAmount::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndSource.SetCheckValue(true);
	m_wndSource.LimitText(1024);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndReportPeriod.InsertColumn(1, _T("Name"), CFMT_TEXT, 100);

	m_wndSource.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndSource.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CPMAdvanceAmount::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndYear.SetEvent(WE_CHANGE, _OnYearChangeFnc);
	//m_wndYear.SetEvent(WE_SETFOCUS, _OnYearSetfocusFnc);
	//m_wndYear.SetEvent(WE_KILLFOCUS, _OnYearKillfocusFnc);
	m_wndYear.SetEvent(WE_CHECKVALUE, _OnYearCheckValueFnc);
	m_wndReportPeriod.SetEvent(WE_SELENDOK, _OnReportPeriodSelendokFnc);
	//m_wndReportPeriod.SetEvent(WE_SETFOCUS, _OnReportPeriodSetfocusFnc);
	//m_wndReportPeriod.SetEvent(WE_KILLFOCUS, _OnReportPeriodKillfocusFnc);
	m_wndReportPeriod.SetEvent(WE_SELCHANGE, _OnReportPeriodSelectChangeFnc);
	m_wndReportPeriod.SetEvent(WE_LOADDATA, _OnReportPeriodLoadDataFnc);
	//m_wndReportPeriod.SetEvent(WE_ADDNEW, _OnReportPeriodAddNewFnc);
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndSource.SetEvent(WE_SELENDOK, _OnSourceSelendokFnc);
	//m_wndSource.SetEvent(WE_SETFOCUS, _OnSourceSetfocusFnc);
	//m_wndSource.SetEvent(WE_KILLFOCUS, _OnSourceKillfocusFnc);
	m_wndSource.SetEvent(WE_SELCHANGE, _OnSourceSelectChangeFnc);
	m_wndSource.SetEvent(WE_LOADDATA, _OnSourceLoadDataFnc);
	//m_wndSource.SetEvent(WE_ADDNEW, _OnSourceAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	CDate dte;
	dte.ParseDate(pMF->GetSysDate());
	m_nYear = dte.GetYear();
	UpdateData(false);

}
void CPMAdvanceAmount::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndSource.GetDlgCtrlID(), m_szSourceKey);

}
void CPMAdvanceAmount::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szSourceKey.Empty();

}
int CPMAdvanceAmount::SetMode(int nMode){
 		int nOldMode = GetMode();
 		CGuiView::SetMode(nMode);
 		CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
 		CString szSQL;
 		CRecord rs(&pMF->m_db);
  		switch(nMode){
 		case VM_ADD: 
 			EnableControls(TRUE);
 			EnableButtons(TRUE, 3, 4, -1);
 			SetDefaultValues();
 			break;
 		case VM_EDIT: 
 			EnableControls(TRUE);
 			EnableButtons(TRUE, 3, 4, -1);
 			break;
 		case VM_VIEW: 
 			EnableControls(FALSE);
 			EnableButtons(FALSE, 3, 4, -1);
 			break;
 		case VM_NONE: 
 			EnableControls(FALSE);
 			EnableButtons(TRUE, 0, 6, -1);
 			SetDefaultValues();
 			break;
 		};
 		UpdateData(FALSE);
 		return nOldMode;
}
/*void CPMAdvanceAmount::OnYearChange(){
	
} */
/*void CPMAdvanceAmount::OnYearSetfocus(){
	
} */
/*void CPMAdvanceAmount::OnYearKillfocus(){
	
} */
int CPMAdvanceAmount::OnYearCheckValue(){
	return 0;
} 
void CPMAdvanceAmount::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMAdvanceAmount::OnReportPeriodSelendok(){
	 
}
/*void CPMAdvanceAmount::OnReportPeriodSetfocus(){
	
}*/
/*void CPMAdvanceAmount::OnReportPeriodKillfocus(){
	
}*/
long CPMAdvanceAmount::OnReportPeriodLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadReportPeriodList(&m_wndReportPeriod, m_szReportPeriodKey);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && !m_szReportPeriodKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szReportPeriodKey
};
	m_wndReportPeriod.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndReportPeriod.AddItems(
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CPMAdvanceAmount::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CPMAdvanceAmount::OnFromDateChange(){
	
} */
/*void CPMAdvanceAmount::OnFromDateSetfocus(){
	
} */
/*void CPMAdvanceAmount::OnFromDateKillfocus(){
	
} */
int CPMAdvanceAmount::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMAdvanceAmount::OnToDateChange(){
	
} */
/*void CPMAdvanceAmount::OnToDateSetfocus(){
	
} */
/*void CPMAdvanceAmount::OnToDateKillfocus(){
	
} */
int CPMAdvanceAmount::OnToDateCheckValue(){
	return 0;
} 
void CPMAdvanceAmount::OnSourceSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMAdvanceAmount::OnSourceSelendok(){
	 
}
/*void CPMAdvanceAmount::OnSourceSetfocus(){
	
}*/
/*void CPMAdvanceAmount::OnSourceKillfocus(){
	
}*/
long CPMAdvanceAmount::OnSourceLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadProductResourceList(&m_wndSource, m_szSourceKey);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndSource.IsSearchKey() && !m_szSourceKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szSourceKey
};
	m_wndSource.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndSource.AddItems(
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CPMAdvanceAmount::OnSourceAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMAdvanceAmount::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CReport rpt;
	CString szSQL, tmpStr;
	CRecord rs(&pMF->m_db);
	CReportSection *rptHeader=NULL, *rptSectio=NULL, *rptFooter=NULL;
	if (!rpt.Init(_T("Reports\\HMS\\PM_TEST.RPT")))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	if (!m_szReportPeriodKey.IsEmpty())
		tmpStr.Format(_T("%s"), m_wndReportPeriod.GetCurrent(1));
	else
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	if (!m_szSourceKey.IsEmpty())
		rptHeader->SetValue(_T("Source"), m_wndSource.GetCurrent(1));
	while (rs.IsEOF())
	{
		
		rs.MoveNext();
	}
	rpt.PrintPreview();
} 
void CPMAdvanceAmount::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 

CString CPMAdvanceAmount::GetQueryString(){
	CString szSQL;

	return szSQL;
}