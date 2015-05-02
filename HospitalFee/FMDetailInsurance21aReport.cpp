#include "stdafx.h"
#include "FMDetailInsurance21aReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMDetailInsurance21aReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMDetailInsurance21aReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMDetailInsurance21aReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMDetailInsurance21aReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMDetailInsurance21aReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMDetailInsurance21aReport* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMDetailInsurance21aReport *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnClerkAddNew();
}*/
static void _OnDeptSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMDetailInsurance21aReport* )pWnd)->OnDeptSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDeptSelendokFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnDeptSelendok();
}
/*static void _OnDeptSetfocusFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnDeptKillfocus();
}*/
/*static void _OnDeptKillfocusFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnDeptKillfocus();
}*/
static long _OnDeptLoadDataFnc(CWnd *pWnd){
	return ((CFMDetailInsurance21aReport *)pWnd)->OnDeptLoadData();
}
/*static void _OnDeptAddNewFnc(CWnd *pWnd){
	((CFMDetailInsurance21aReport *)pWnd)->OnDeptAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMDetailInsurance21aReport *pVw = (CFMDetailInsurance21aReport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMDetailInsurance21aReport *pVw = (CFMDetailInsurance21aReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMDetailInsurance21aReport *pVw = (CFMDetailInsurance21aReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CFMDetailInsurance21aReport *pVw = (CFMDetailInsurance21aReport *)pWnd;
	pVw->OnCloseSelect();
} 
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	 ((CFMDetailInsurance21aReport*)pWnd)->OnOutPatientSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	 ((CFMDetailInsurance21aReport*)pWnd)->OnInPatientSelect();
}
static int _OnAddCFMDetailInsurance21aReportFnc(CWnd *pWnd){
	 return ((CFMDetailInsurance21aReport*)pWnd)->OnAddCFMDetailInsurance21aReport();
} 
static int _OnEditCFMDetailInsurance21aReportFnc(CWnd *pWnd){
	 return ((CFMDetailInsurance21aReport*)pWnd)->OnEditCFMDetailInsurance21aReport();
} 
static int _OnDeleteCFMDetailInsurance21aReportFnc(CWnd *pWnd){
	 return ((CFMDetailInsurance21aReport*)pWnd)->OnDeleteCFMDetailInsurance21aReport();
} 
static int _OnSaveCFMDetailInsurance21aReportFnc(CWnd *pWnd){
	 return ((CFMDetailInsurance21aReport*)pWnd)->OnSaveCFMDetailInsurance21aReport();
} 
static int _OnCancelCFMDetailInsurance21aReportFnc(CWnd *pWnd){
	 return ((CFMDetailInsurance21aReport*)pWnd)->OnCancelCFMDetailInsurance21aReport();
} 
CFMDetailInsurance21aReport::CFMDetailInsurance21aReport(CWnd *pParent)
{
	m_nDlgWidth = 440;
	m_nDlgHeight = 190;
	SetDefaultValues();
}
CFMDetailInsurance21aReport::~CFMDetailInsurance21aReport()
{
}
void CFMDetailInsurance21aReport::OnCreateComponents()
{
	m_wndReportFilter.Create(this, _T("Report Parameters"), 5, 5, 430, 150);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 85, 55);
	m_wndYear.Create(this,90, 30, 210, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Month"), 215, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 85, 85);
	m_wndFromDate.Create(this,90, 60, 210, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 215, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	//m_wndPrint.Create(this, _T("&Print Preview"), 265, 155, 345, 180);
	m_wndExport.Create(this, _T("&Export"), 350, 155, 430, 180);
	m_wndDrugPTTT.Create(this, _T("Kh\xF4ng l\x1EA5y VTTH trong PTTT v\xE0o \x62\xE1o \x63\xE1o"), 125, 90, 425, 115);
	m_wndOutPatient.Create(this, _T("OutPatient"), 125, 120, 270, 145);
	m_wndInPatient.Create(this, _T("InPatient"), 275, 120, 425, 145);
}
void CFMDetailInsurance21aReport::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(FALSE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);
	m_wndDept.LimitText(81);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	//m_wndClerk.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	//m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

	//m_wndDept.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	//m_wndDept.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);
}
void CFMDetailInsurance21aReport::OnSetWindowEvents()
{
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
	m_wndClerk.SetEvent(WE_SELENDOK, _OnClerkSelendokFnc);
	//m_wndClerk.SetEvent(WE_SETFOCUS, _OnClerkSetfocusFnc);
	//m_wndClerk.SetEvent(WE_KILLFOCUS, _OnClerkKillfocusFnc);
	m_wndClerk.SetEvent(WE_SELCHANGE, _OnClerkSelectChangeFnc);
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	//m_wndClerk.SetEvent(WE_ADDNEW, _OnClerkAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndDept.SetEvent(WE_SELENDOK, _OnDeptSelendokFnc);
	//m_wndDept.SetEvent(WE_SETFOCUS, _OnDeptSetfocusFnc);
	//m_wndDept.SetEvent(WE_KILLFOCUS, _OnDeptKillfocusFnc);
	m_wndDept.SetEvent(WE_SELCHANGE, _OnDeptSelectChangeFnc);
	m_wndDept.SetEvent(WE_LOADDATA, _OnDeptLoadDataFnc);
	//m_wndDept.SetEvent(WE_ADDNEW, _OnDeptAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = pMF->GetSysDate() + _T("00:00");
	m_szToDate = pMF->GetSysDate() + _T("23:59");
	SetMode(VM_EDIT);

}
void CFMDetailInsurance21aReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	//DDX_TextEx(pDX, m_wndDept.GetDlgCtrlID(), m_szDeptKey);
	DDX_Check(pDX, m_wndDrugPTTT.GetDlgCtrlID(), m_bCheckDrug);
	DDX_Check(pDX, m_wndOutPatient.GetDlgCtrlID(), m_bOutPatient);
	DDX_Check(pDX, m_wndInPatient.GetDlgCtrlID(), m_bInPatient);

}
void CFMDetailInsurance21aReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMDetailInsurance21aReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMDetailInsurance21aReport::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDeptKey.Empty();
	m_szClerkKey.Empty();
	m_bOutPatient=TRUE;
	m_bInPatient=FALSE;

}
int CFMDetailInsurance21aReport::SetMode(int nMode){
 		int nOldMode = GetMode(); 
		CGuiView::SetMode(nMode); 
 		CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
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
 			EnableButtons(TRUE, 0, 1, 2, 3, 4, -1);
			m_wndDept.EnableWindow(FALSE);
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
/*void CFMDetailInsurance21aReport::OnYearChange(){
	
} */
/*void CFMDetailInsurance21aReport::OnYearSetfocus(){
	
} */
/*void CFMDetailInsurance21aReport::OnYearKillfocus(){
	
} */
int CFMDetailInsurance21aReport::OnYearCheckValue()
{
	UpdateData(TRUE);
	if (m_nYear > 0)
	{
		CDateTime dt;
		CDate date;
		CString szTemp;

		dt.ParseDateTime(m_szFromDate);
		date = dt.GetDate();
		if (date.GetYear() != 1752)
		{
			dt.SetDate(m_nYear, date.GetMonth(), date.GetDay());
			m_szFromDate = dt.GetDateTime();
			szTemp.Format(_T("%.2d/%.2d/%.4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndFromDate.SetWindowText(szTemp);
		}
		dt.ParseDateTime(m_szToDate);
		date = dt.GetDate();
		if (date.GetYear() != 1752)
		{
			dt.SetDate(m_nYear, date.GetMonth(), date.GetDay());
			m_szToDate = dt.GetDateTime();
			szTemp.Format(_T("%.2d/%.2d/%.4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndToDate.SetWindowText(szTemp);
		}
	}
	return 0;
} 
void CFMDetailInsurance21aReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMDetailInsurance21aReport::OnReportPeriodSelendok()
{	
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	CString tmpStr;
	CDate dte, date;

	UpdateData(true);

	date.ParseDate(pMF->GetSysDate());
	int nYear = date.GetYear();
	int nMonth = ToInt(m_szReportPeriodKey);

	if (nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/01 00:00"), nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), nYear, nMonth, dte.GetMonthLastDay());
	}
	if (nMonth == 13)
	{
		m_szFromDate.Format(_T("%.4d/01/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/03/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 14)
	{
		m_szFromDate.Format(_T("%.4d/04/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/06/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 15)
	{
		m_szFromDate.Format(_T("%.4d/07/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/09/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 16)
	{
		m_szFromDate.Format(_T("%.4d/10/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/10/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 17)
	{
		m_szFromDate.Format(_T("%.4d/01/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/12/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}

	UpdateData(false);

}
/*void CFMDetailInsurance21aReport::OnReportPeriodSetfocus(){
	
}*/
/*void CFMDetailInsurance21aReport::OnReportPeriodKillfocus(){
	
}*/
long CFMDetailInsurance21aReport::OnReportPeriodLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0)
	{
		szWhere.Format(_T(" WHERE hpr_idx=%ld "), ToInt(m_szReportPeriodKey));
	}
	m_wndReportPeriod.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT * FROM hms_period_report %s ORDER BY hpr_idx "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndReportPeriod.AddItems(
			rs.GetValue(_T("hpr_idx")), 
			rs.GetValue(_T("hpr_name")),			
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMDetailInsurance21aReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMDetailInsurance21aReport::OnFromDateChange(){
	
} */
/*void CFMDetailInsurance21aReport::OnFromDateSetfocus(){
	
} */
/*void CFMDetailInsurance21aReport::OnFromDateKillfocus(){
	
} */
int CFMDetailInsurance21aReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMDetailInsurance21aReport::OnToDateChange(){
	
} */
/*void CFMDetailInsurance21aReport::OnToDateSetfocus(){
	
} */
/*void CFMDetailInsurance21aReport::OnToDateKillfocus(){
	
} */
int CFMDetailInsurance21aReport::OnToDateCheckValue()
{
	return 0;
} 
void CFMDetailInsurance21aReport::OnClerkSelectChange(int nOldItemSel, int nNewItemSel)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMDetailInsurance21aReport::OnClerkSelendok(){
	 
}
/*void CFMDetailInsurance21aReport::OnClerkSetfocus(){
	
}*/
/*void CFMDetailInsurance21aReport::OnClerkKillfocus(){
	
}*/
long CFMDetailInsurance21aReport::OnClerkLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	}
	m_wndClerk.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as userid, su_name as name FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CFMDetailInsurance21aReport::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMDetailInsurance21aReport::OnDeptSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMDetailInsurance21aReport::OnDeptSelendok(){
	 
}
/*void CFMDetailInsurance21aReport::OnDeptSetfocus(){
	
}*/
/*void CFMDetailInsurance21aReport::OnDeptKillfocus(){
	
}*/
long CFMDetailInsurance21aReport::OnDeptLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();
	if(m_wndDept.IsSearchKey() && !m_szDeptKey.IsEmpty())
	{
		szWhere.Format(_T(" AND sd_id='%s' "), m_szDeptKey);
	}
	m_wndDept.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT sd_id AS id, sd_name AS name FROM sys_dept ")\
		         _T("WHERE sd_type='DT' %s ORDER BY sd_id"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDept.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMDetailInsurance21aReport::OnDeptAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMDetailInsurance21aReport::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_bPreview=true;
	//PrintTonghopsudungdichvukythuatBHYT();
} 
void CFMDetailInsurance21aReport::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();	
	m_bPreview=false;
	//PrintTonghopsudungdichvukythuatBHYT();
} 
void CFMDetailInsurance21aReport::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	ExportTonghopsudungdichvukythuatBHYT();
} 
void CFMDetailInsurance21aReport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//OnCancel();
	
} 
void CFMDetailInsurance21aReport::OnOutPatientSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bInPatient = FALSE;
	UpdateData(FALSE);
}
void CFMDetailInsurance21aReport::OnInPatientSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bOutPatient = FALSE;
	UpdateData(FALSE);
}
int CFMDetailInsurance21aReport::OnAddCFMDetailInsurance21aReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMDetailInsurance21aReport::OnEditCFMDetailInsurance21aReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMDetailInsurance21aReport::OnDeleteCFMDetailInsurance21aReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
 	CString szSQL; 
 	if(ShowMessage(1, MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2) == IDNO) 
 		return -1; 
 	szSQL.Format(_T("DELETE FROM  WHERE  AND") ); 
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret >= 0){ 
 		SetMode(VM_NONE); 
 		OnCancelCFMDetailInsurance21aReport(); 
 	}
	return 0;
}
int CFMDetailInsurance21aReport::OnSaveCFMDetailInsurance21aReport(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT) 
 		return -1; 
 	if(!IsValidateData()) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
 	CString szSQL; 
 	if(GetMode() == VM_ADD){ 
 		//szSQL = m_tblTbl.GetInsertSQL(); 
 	} 
 	else if(GetMode() == VM_EDIT){ 
 		CString szWhere; 
 		szWhere.Format(_T(" WHERE")); 
 		//szSQL = m_tblTbl.GetUpdateSQL(_T("createdby,createddate")); 
 		szSQL += szWhere; 
 	} 
 _fmsg(_T("%s"), szSQL); 
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret > 0) 
 	{ 
 		//OnCFMDetailInsurance21aReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMDetailInsurance21aReport::OnCancelCFMDetailInsurance21aReport(){
 	if(GetMode() == VM_EDIT) 
 	{ 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 		SetMode(VM_NONE); 
 	} 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
 	return 0;
} 
int CFMDetailInsurance21aReport::OnCFMDetailInsurance21aReportListLoadData(){
	return 0;
}

//Bao cao tong hop su dung dich vu ky thuat theo quu ( mau 21/BHYT)
void CFMDetailInsurance21aReport::ExportTonghopsudungdichvukythuatBHYT()
{	
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString tmpStr, szSQL, szWhere, m_szGroups, szDeptidCLS, szDeptOperation, szDeptDrug;
	CString szDept;
	szWhere.Empty();
	m_szGroups.Empty();

	UpdateData(true);
	BeginWaitCursor();

	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	//xls.SetRowHeight(8, 50);
	xls.SetColumnWidth(0, 8);
	xls.SetColumnWidth(1, 50);
	xls.SetColumnWidth(2, 8);
	xls.SetColumnWidth(3, 12);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 22);
	xls.SetColumnWidth(6, 10);

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	xls.SetCellText(0, 1, tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	xls.SetCellText(0, 2, tmpStr);

	if (m_bInPatient && !m_szDeptKey.IsEmpty())
	{
		StringUpper(m_wndDept.GetCurrent(1), tmpStr);
		xls.SetCellText(0, 3, tmpStr, true);
	}
	
	//xls.SetCellMergedColumns(0, 5, 12);	
	//StringUpper(_T("General statistics and technical services for patients using medical insurance"), tmpStr);
	xls.SetCellText(0, 5, _T("\x42\xC1O \x43\xC1O \x43HI TI\x1EBET S\x1EEC \x44\x1EE4NG \x44\x1ECA\x43H V\x1EE4 K\x1EF8 THU\x1EACT TH\x45O QU\xDD (M\x1EAAU \x32\x31/\x42HYT)"),
					FMT_TEXT, true, 13);
	//xls.SetCellMergedColumns(0, 6, 12);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), 
				  CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(0, 6, tmpStr, FMT_TEXT, true, 12, 0);
	

	int nRow = 8;
	
	
	xls.SetCellText(0, nRow, _T("STT"), FMT_TEXT | FMT_CENTER, true);	
	TranslateString(_T("Desc"), tmpStr);
	xls.SetCellText(1, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
	TranslateString(_T("Qty"), tmpStr);
	xls.SetCellText(2, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
	TranslateString(_T("Price"), tmpStr);	
	xls.SetCellText(3, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
	TranslateString(_T("Amount"), tmpStr);
	xls.SetCellText(4, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);


	//Page Header
	//Report Detail
	int nIndex = 1;
	int nSubIndex = 1;
	CString szOldLine, szNewLine, szAmount;
	CString szDate, szMoney, szTemp;

	long double grpCost = 0;
	long double ttlCost = 0;
	long double totalCost = 0;
	double cost = 0, costword = 0;
	grpCost = ttlCost = cost = 0;

	CString szStatus;

	//szStatus.Empty();
	//szStatus.Format(_T(" AND (hd_status='T' or hcr_status='T') "));
	
	if (m_bInPatient == 1)
	{
		//szDeptidCLS.Format(_T(" AND pcmso_depttype ='I' "));
		//szDeptOperation.Format(_T(" AND ho_depttype='I' "));
		//szDeptDrug.Format(_T(" AND hpo_depttype='I' "));
		szWhere.AppendFormat(_T(" and fi.hfe_class not in('E') "));
	}
	else if (m_bOutPatient==1)	
	{
		//szDeptidCLS.Format(_T(" AND pcmso_depttype ='E' "));	
		//szDeptOperation.Format(_T(" AND ho_depttype='E' "));
		//szDeptDrug.Format(_T(" AND hpo_depttype='E' "));
		//szDept.Format(_T(" AND hfi_deptid='KB' "));
		szWhere.AppendFormat(_T(" and fi.hfe_class in('E') "));
	}
		
	if (m_bInPatient==1 && m_bOutPatient==1)	
	{
		//szDeptidCLS.Empty();
		//szDeptOperation.Empty();
		//szDeptDrug.Empty();
		szDept.Empty();
	}
	
	m_szDrugPTTT.Format(_T(" and hpo_ordertype in('P','D','M','E') "));
	if (m_bCheckDrug == 1)
	{
		m_szDrugPTTT.Format(_T(" and hpo_ordertype in('P','D','E') "));
	}


/*Phi kham benh
-----------------------------------------------------------------------------------------------------------------------------*/
if (m_bInPatient != 1)
{
	szSQL.Format(_T(" select he_docno as docno,") \
				_T("        he_examtype as examtype,") \
				_T("        hfl_name as examname,") \
				_T("        hfl_unit as unit,") \
				_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) as fullname,") \
				_T("        hms_getage(Cast_Timestamp2Date(hd_admitdate), hp_birthdate) as age,") \
				_T("        hc_cardno as cardno,") \
				_T("        hc_expdate as expdate,") \
				_T("        sum(hfe_quantity) as soluong,") \
				_T("        hms_fee_view.hfe_insprice as insprice") \
				_T(" from hms_exam") \
				_T(" left join hms_fee_view on(hfe_docno=he_docno and hfe_orderid=he_receptidx)") \
				_T(" left join hms_fee_invoice fi on(fi.hfe_docno=he_docno and fi.hfe_invoiceno=hms_fee_view.hfe_invoiceno)") \
				_T(" left join hms_fee_list on(hfl_feeid=he_examtype)") \
				_T(" left join hms_doc on(hd_docno=he_docno)") \
				_T(" left join hms_patient on(hp_patientno=hd_patientno)") \
				_T(" left join hms_object on(ho_id=hd_object)") \
				_T(" left join hms_card on(hc_patientno=hd_patientno and hc_idx=hd_cardidx)") \
				_T(" where hd_status='T' and hms_fee_view.hfe_discount > 0 and ho_type in('I','C')") \
				_T("       and fi.hfe_status='P' and hfl_inspaid='Y' and hms_fee_view.hfe_insprice > 0") \
				_T("       and fi.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
				_T(" group by he_examtype, he_docno, hp_surname, hp_midname, hp_firstname, hfl_unit,") \
				_T("          hfl_name, hd_admitdate, hp_birthdate, hfe_insprice, hc_cardno, hc_expdate") \
				_T(" order by he_examtype, he_docno, examname"), m_szFromDate, m_szToDate, szWhere);


	
	_fmsg(_T("%s"), szSQL);
	nIndex = 1;
	grpCost = ttlCost = cost = 0;
	CString szExamType;
	long double nItemCost = 0;
	int nCount, nItemCount = 0, nQuantity = 0;

	rs.ExecSQL(szSQL);

	if (!rs.IsEOF())
	{	
		++nRow;		
		TranslateString(_T("Nh\xF3m ph\xED kh\xE1m"), tmpStr);
		xls.SetCellText(0, nRow,_T("E"), FMT_TEXT, true, 12);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true, 12);

		while (!rs.IsEOF())
		{	
			if (szExamType.IsEmpty() || szExamType != rs.GetValue(_T("examtype")))
			{
				nRow++;		
				tmpStr.Format(_T("%d"), nIndex++);
				xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER, true, 10);
				rs.GetValue(_T("examname"), tmpStr);
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true, 10);
				xls.SetCellText(3, nRow, rs.GetValue(_T("insprice")), FMT_NUMBER1, true, 10);
				
				if (nQuantity > 0)
				{
					tmpStr.Format(_T("%d"), nQuantity);
					xls.SetCellText(2, nRow - nItemCount - 1, tmpStr, FMT_INTEGER, true, 10);
				}

				if (grpCost > 0)
				{
					tmpStr.Format(_T("%.2lf"), grpCost);
					xls.SetCellText(4, nRow - nItemCount - 1, tmpStr, FMT_NUMBER1, true, 10);
					ttlCost += grpCost;
				}
				szExamType = rs.GetValue(_T("examtype"));
				grpCost = 0;
				nQuantity = 0;
				nItemCount = 0;
				nSubIndex = 1;
			}

			nRow++;
			tmpStr.Format(_T("%d"), nSubIndex);

			xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER, false, 10);
			szTemp.Format(_T("[%s]%s - %s"), rs.GetValue(_T("docno")), rs.GetValue(_T("fullname")), 
						  rs.GetValue(_T("age")));

			xls.SetCellText(1, nRow, szTemp, FMT_TEXT, false, 10);
			xls.SetCellText(2, nRow, rs.GetValue(_T("soluong")), FMT_INTEGER, false, 10);

			rs.GetValue(_T("insprice"), cost);
			rs.GetValue(_T("soluong"), nCount);

			tmpStr.Format(_T("%.2f"), cost * nCount);
			xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, false, 10);

			rs.GetValue(_T("cardno"), szTemp);
			xls.SetCellText(5, nRow, szTemp, FMT_TEXT, false, 10);

			szTemp = CDate::Convert(rs.GetValue(_T("expdate")), yyyymmdd, ddmmyyyy);
			xls.SetCellText(6, nRow, szTemp, FMT_DATE, false, 10);

			grpCost += cost * nCount;			
			nQuantity += nCount;
			nItemCount++;
			nSubIndex++;
			
			rs.MoveNext();
		}		
		

		if (nQuantity > 0)
		{
			tmpStr.Format(_T("%d"), nQuantity);
			xls.SetCellText(2, nRow - nItemCount, tmpStr, FMT_INTEGER, true, 10);
		}
		if (grpCost > 0)
		{
			tmpStr.Format(_T("%.2lf"), grpCost);
			xls.SetCellText(4, nRow - nItemCount, tmpStr, FMT_NUMBER1, true, 10);
			ttlCost += grpCost;
		}	
		
		if (ttlCost > 0)
		{
			nRow++;
			TranslateString(_T("Total Group"), tmpStr);
			xls.SetCellText(1, nRow, tmpStr + _T("(E)"), FMT_TEXT, TRUE);
			tmpStr.Format(_T("%.2lf"), ttlCost);
			xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
			totalCost += ttlCost;
		}
	}	
}

	
//------------------------------------------------------------------------------------------------

//if (!m_szDeptKey.IsEmpty())
//		szDept.Format(_T(" AND pcmso_deptid='%s' "), m_szDeptKey);

szSQL.Format(_T(" select hfg_name as groupname,") \
			_T("        groupid,") \
			_T("        itemid,") \
			_T("        pcmsname,") \
			_T("        unit,") \
			_T("        sum(soluong) as soluong,") \
			_T("        insprice,") \
			_T("        sum(soluong*insprice) as suminsprice") \
			_T(" from") \
			_T(" (") \
			_T("   select hfe_invoiceno as invoiceno,") \
			_T("          hpc_docno as docno,") \
			_T("          hfl_groupid as groupid,") \
			_T("          hpcl_itemid as itemid,") \
			_T("          hfl_name as pcmsname,") \
			_T("          hfl_unit as unit,") \
			_T("          sum(hfe_quantity) as soluong,") \
			_T("          hfe_insprice as insprice") \
			_T("   from hms_testorder") \
			_T("   left join hms_testorderline on(hpcl_orderid=hpc_orderid and hpcl_docno=hpc_docno)") \
			_T("   left join hms_fee_list on(hfl_feeid=hpcl_itemid)") \
			_T("   left join hms_fee_view on(hfe_orderid=hpcl_orderid and hfe_orderline=hpcl_orderlineid)") \
			_T("   left join hms_doc on(hd_docno=hpc_docno)") \
			_T("   left join hms_object on(ho_id=hd_object)") \
			_T("   where hfl_inspaid='Y' and hfe_discount > 0 and hd_status='T'") \
			_T("         and substr(hfl_groupid,1,2)='B1' and ho_type in('I','C')") \
			_T("         and hfl_hitech='N'") \
			_T("   group by hpc_docno, hfl_groupid, hpcl_itemid, hfl_name, hfl_unit, hfe_insprice, hfe_invoiceno") \
			_T(" ) tbl") \
			_T(" left join hms_fee_invoice fi on(hfe_docno=docno and hfe_invoiceno=invoiceno)") \
			_T(" left join hms_fee_group on(hfg_id=groupid)") \
			_T(" where insprice > 0 and fi.hfe_status='P'") \
			_T("       and fi.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
			_T(" group by groupid, hfg_name, itemid, pcmsname, unit, insprice") \
			_T(" order by groupid, itemid"),
			m_szFromDate, m_szToDate, szWhere);
	

	//_fmsg(_T("%s"), szSQL);

	nIndex = 1;
	grpCost = ttlCost = cost = ttlCost = 0;
	CRecord srs(&pMF->m_db);

	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{	
		++nRow;
		TranslateString(_T("Test"), tmpStr);
		xls.SetCellText(0, nRow, _T("A."), FMT_TEXT, TRUE, 12);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE, 12);

		while (!rs.IsEOF())
		{
			rs.GetValue(_T("groupname"), szNewLine);

			if (szNewLine != szOldLine && !szNewLine.IsEmpty())
			{
				CString szField;
				if (grpCost > 0)
				{
					nRow++;
					TranslateString(_T("C\x1ED9ng:"), tmpStr);
					xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE);
					tmpStr.Format(_T("%.2lf"), grpCost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
					ttlCost += grpCost;
					nIndex = 1;
					grpCost = 0;
				}
				nRow++;
				xls.SetCellText(1, nRow, szNewLine, FMT_TEXT, TRUE, 10);
				szOldLine = szNewLine;
			}

			nRow++;			
			tmpStr.Format(_T("%d"), nIndex++);
			xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER, true);

			rs.GetValue(_T("pcmsname"), tmpStr);
			xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true);

			rs.GetValue(_T("soluong"), tmpStr);					
			xls.SetCellText(2, nRow, tmpStr, FMT_INTEGER, true);

			rs.GetValue(_T("insprice"), cost);		
			tmpStr.Format(_T("%.2f"), cost);
			xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1, true);

			rs.GetValue(_T("suminsprice"), cost);
			grpCost += cost;
			tmpStr.Format(_T("%.2f"), cost);
			xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
			
			szSQL.Format(_T(" select fi.hfe_invoiceno,") \
						_T("        hpc_docno as docno,") \
						_T("        hpcl_itemid as itemid,") \
						_T("        hfl_groupid as groupid,") \
						_T("        hfl_name as pcmsname,") \
						_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) as fullname,") \
						_T("        hms_getage(Cast_Timestamp2Date(hd_admitdate), hp_birthdate) as age,") \
						_T("        hc_cardno as cardno,") \
						_T("        hc_expdate as expdate,") \
						_T("        hfl_unit as unit,") \
						_T("        sum(fe.hfe_quantity) as soluong,") \
						_T("        fe.hfe_insprice as insprice,") \
						_T("        sum(fe.hfe_quantity*fe.hfe_insprice) as amount") \
						_T(" from hms_testorder") \
						_T(" left join hms_testorderline on(hpc_orderid=hpcl_orderid and hpcl_docno=hpc_docno)") \
						_T(" left join hms_fee_list on(hfl_feeid=hpcl_itemid)") \
						_T(" left join hms_doc on(hd_docno=hpc_docno)") \
						_T(" left join hms_object on(ho_id=hd_object)") \
						_T(" left join hms_patient on(hp_patientno=hd_patientno)") \
						_T(" left join hms_card on(hc_patientno=hd_patientno and hc_idx=hd_cardidx)") \
						_T(" left join hms_fee_view fe on(fe.hfe_orderid=hpcl_orderid and fe.hfe_orderline=hpcl_orderlineid)") \
						_T(" left join hms_fee_invoice fi on(fi.hfe_docno=hpc_docno and fi.hfe_invoiceno=fe.hfe_invoiceno)") \
						_T(" where fe.hfe_discount > 0") \
						_T("       and substr(hfl_groupid, 1, 2)='B1'") \
						_T("       and ho_type in('I','C')") \
						_T("       and hfl_hitech='N'") \
						_T("       and hfl_inspaid='Y'") \
						_T("       and fi.hfe_status='P'") \
						_T("       and hpcl_itemid='%s'") \
						_T("       and fi.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
						_T(" group by hpc_docno, hfl_groupid, hpcl_itemid, hfl_name, hp_surname, hp_midname, hp_firstname,") \
						_T("          hd_admitdate, hp_birthdate, hc_cardno, hc_expdate, hfl_unit, fe.hfe_insprice, fi.hfe_invoiceno") \
						_T(" order by hpc_docno"),
						rs.GetValue(_T("itemid")), m_szFromDate, m_szToDate, szWhere);

			int ret = srs.ExecSQL(szSQL);
			//_fmsg(_T("%s"), szSQL);
			//return;
			nSubIndex = 1;

			while(!srs.IsEOF())
			{
				nRow++;
				tmpStr.Format(_T("%d"), nSubIndex);
				xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);
				szTemp.Format(_T("[%s]%s - %s"), srs.GetValue(_T("docno")), srs.GetValue(_T("fullname")),
							  srs.GetValue(_T("age")));
				xls.SetCellText(1, nRow, szTemp, FMT_TEXT);
				xls.SetCellText(2, nRow, srs.GetValue(_T("soluong")), FMT_INTEGER);
				xls.SetCellText(4, nRow, srs.GetValue(_T("amount")), FMT_NUMBER1);
				srs.GetValue(_T("cardno"), szTemp);
				xls.SetCellText(5, nRow, szTemp, FMT_TEXT);
				szTemp = CDate::Convert(srs.GetValue(_T("expdate")), yyyymmdd, ddmmyyyy);
				xls.SetCellText(6, nRow, szTemp, FMT_DATE);
				nSubIndex++;
				srs.MoveNext();
			}

			rs.MoveNext();
		}		
		
		if (grpCost > 0)
		{
			nRow++;
			TranslateString(_T("C\x1ED9ng:"), tmpStr);
			xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE);
			tmpStr.Format(_T("%.2lf"), grpCost);
			xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
			ttlCost += grpCost;		
		}

		if (ttlCost > 0)
		{
			nRow++;
			TranslateString(_T("Total Group"), tmpStr);
			xls.SetCellText(1, nRow, tmpStr + _T("(A)"),FMT_TEXT,TRUE);
			tmpStr.Format(_T("%.2lf"), ttlCost);
			xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1,TRUE);
			totalCost+=ttlCost;
		}
	}

	/*xls.Save(_T("Exports\\Baocao21achitietBHYT.xls"));
	EndWaitCursor();
	return;*/

/*--------------------------------------------------------------------------------------------------------------------------------------------
	NHOM PHAU THUAT, THU THUAT
--------------------------------------------------------------------------------------------------------------------------------------------*/

		szSQL.Format(_T(" select hfg_name as operatype,") \
					_T("        ho_itemid as itemid,") \
					_T("        hfl_name as itemname,") \
					_T("        sum(ho_qty) as qty,") \
					_T("        fe.hfe_insprice as insprice,") \
					_T("        sum(ho_qty*fe.hfe_insprice) as amount") \
					_T(" from hms_patient") \
					_T(" left join hms_doc on(hd_patientno=hp_patientno)") \
					_T(" left join hms_object ob on(ob.ho_id=hd_object)") \
					_T(" left join hms_operation op on(op.ho_docno=hd_docno)") \
					_T(" left join hms_fee_list on(hfl_feeid=ho_itemid)") \
					_T(" left join hms_fee_group on(hfg_id=hfl_groupid)") \
					_T(" left join hms_fee_view fe on(fe.hfe_docno=ho_docno and fe.hfe_itemid=ho_itemid)") \
					_T(" left join hms_fee_invoice fi on(fi.hfe_docno=ho_docno and fi.hfe_invoiceno=fe.hfe_invoiceno)") \
					_T(" where fi.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
					_T("       and hd_status='T' and hfl_inspaid='Y' and fi.hfe_status='P' and hfl_hitech='N'") \
					_T("       and ob.ho_type in('I','C') and fe.hfe_discount > 0 %s") \
					_T(" group by hfg_name, ho_itemid, hfl_name, fe.hfe_insprice") \
					_T(" order by hfg_name, hfl_name"), 
					m_szFromDate, m_szToDate, szWhere);


	nIndex = 1;
	grpCost = ttlCost = cost = 0;
	rs.ExecSQL(szSQL);

	if (!rs.IsEOF())
	{
		++nRow;
		TranslateString(_T("Surgery - procedures"), tmpStr);
		xls.SetCellText(0, nRow, _T("D"), FMT_TEXT, TRUE, 12);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE, 12);		
		nIndex = 1;

		while (!rs.IsEOF())
		{
			rs.GetValue(_T("operatype"), szNewLine);

			if (szNewLine != szOldLine && !szNewLine.IsEmpty())
			{
				CString szField;
				if (grpCost > 0)
				{	
					nRow++;
					TranslateString(_T("C\x1ED9ng:"), tmpStr);
					xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE);
					tmpStr.Format(_T("%.2lf"), grpCost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
					ttlCost += grpCost;
					nIndex = 1;
					grpCost = 0;				
				}
				nRow++;
				xls.SetCellText(1, nRow, szNewLine, FMT_TEXT, TRUE);
				szOldLine = szNewLine;
			}

			nRow++;
			tmpStr.Format(_T("%d"), nIndex++);
			xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER, true);

			rs.GetValue(_T("itemname"), tmpStr);
			xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true);

			rs.GetValue(_T("qty"), tmpStr);
			xls.SetCellText(2, nRow, tmpStr, FMT_INTEGER, true);

			rs.GetValue(_T("price"), cost);
			tmpStr.Format(_T("%.2f"), cost);
			xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1, true);

			rs.GetValue(_T("amount"), cost);
			grpCost += cost;
			tmpStr.Format(_T("%.2f"), cost);
			xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);

			szSQL.Format(_T(" select ho_docno as docno,") \
						_T("         trim(hp_surname||' '||hp_midname||' '||hp_firstname) as fullname,") \
						_T("         hms_getage(Cast_Timestamp2Date(hd_admitdate), hp_birthdate) as age,") \
						_T("         hc_cardno as cardno,") \
						_T(" 	     hc_expdate as expdate,") \
						_T(" 	     hfl_name as name,") \
						_T(" 		 fe.hfe_insprice as price,") \
						_T(" 		 sum(ho_qty) as qty,") \
						_T(" 		 sum(fe.hfe_insprice*ho_qty) as cost") \
						_T(" from hms_patient") \
						_T(" left join hms_doc on(hd_patientno=hp_patientno)") \
						_T(" left join hms_object ob on(ob.ho_id=hd_object)") \
						_T(" left join hms_card on(hc_idx=hd_cardidx and hc_patientno=hd_patientno)") \
						_T(" left join hms_operation op on(op.ho_docno=hd_docno)") \
						_T(" left join hms_fee_list on(hfl_feeid=ho_itemid)") \
						_T(" left join hms_fee_group on(hfg_id=hfl_groupid)") \
						_T(" left join hms_fee_view fe on(fe.hfe_docno=ho_docno and fe.hfe_itemid=ho_itemid)") \
						_T(" left join hms_fee_invoice fi on(fi.hfe_docno=ho_docno and fi.hfe_invoiceno=fe.hfe_invoiceno)") \
						_T(" where fi.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
						_T("       and hd_status='T' and hfl_inspaid='Y' and fi.hfe_status='P' and hfl_hitech='N'") \
						_T("       and ob.ho_type in('I','C') and fe.hfe_discount > 0 and ho_itemid='%s' %s") \
						_T(" group by ho_docno, hp_surname, hp_midname, hp_firstname, hd_admitdate, hp_birthdate,") \
						_T("          hc_cardno, hc_expdate, hfl_name, fe.hfe_insprice") \
						_T(" order by docno"),
						m_szFromDate, m_szToDate, rs.GetValue(_T("itemid")), szWhere);

			srs.ExecSQL(szSQL);

			nSubIndex = 1;

			while (!srs.IsEOF())
			{
				nRow++;
				tmpStr.Format(_T("%d"), nSubIndex);
				xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);
				szTemp.Format(_T("[%s]%s - %s"), srs.GetValue(_T("docno")), srs.GetValue(_T("fullname")),
							  srs.GetValue(_T("age")));
				xls.SetCellText(1, nRow, szTemp, FMT_TEXT);
				xls.SetCellText(2, nRow, srs.GetValue(_T("qty")), FMT_INTEGER);
				xls.SetCellText(4, nRow, srs.GetValue(_T("cost")), FMT_NUMBER1);
				srs.GetValue(_T("cardno"), szTemp);
				xls.SetCellText(5, nRow, szTemp, FMT_TEXT);
				szTemp = CDate::Convert(srs.GetValue(_T("expdate")), yyyymmdd, ddmmyyyy);
				xls.SetCellText(6, nRow, szTemp, FMT_DATE);
				nSubIndex++;
				srs.MoveNext();
			}


			rs.MoveNext();
		}		
		
		
		if (grpCost > 0)
		{
			nRow++;
			TranslateString(_T("C\x1ED9ng:"), tmpStr);
			xls.SetCellText(1, nRow, tmpStr,FMT_TEXT, TRUE);
			tmpStr.Format(_T("%.2lf"),grpCost);
			xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
			ttlCost += grpCost;
		
		}

		if (ttlCost > 0)
		{
			nRow++;
			TranslateString(_T("Total Group"), tmpStr);
			xls.SetCellText(1, nRow, tmpStr + _T("(D)"),FMT_TEXT, TRUE);
			tmpStr.Format(_T("%.2lf"),ttlCost);
			xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
			totalCost+=ttlCost;
		}
	}


	if (totalCost > 0)
	{
		nRow++;
		TranslateString(_T("Total Group"), tmpStr);
		xls.SetCellText(1, nRow, tmpStr + + _T("(A+B+C+D+E+F)"),FMT_TEXT, TRUE,12);
		tmpStr.Format(_T("%.2lf"),totalCost);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE, 12);		
	}
	//Page Footer
	//Report Footer
	nRow++;
	szMoney.Format(_T("%.0f"), floor(totalCost + 0.5));
	MoneyToString(szMoney, tmpStr);
	xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(_T(" Ng\xE0y %s th\xE1ng %s n\x103m %s"), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	xls.SetCellText(3, nRow+3, szDate,FMT_TEXT);


	xls.Save(_T("Exports\\Baocao21achitietBHYT.xls"));
	EndWaitCursor();
}
