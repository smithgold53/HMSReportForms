#include "stdafx.h"
#include "PACSFunctionalProbeListReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CPACSFunctionalProbeListReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSFunctionalProbeListReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CPACSFunctionalProbeListReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSFunctionalProbeListReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSFunctionalProbeListReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CPACSFunctionalProbeListReport *pVw = (CPACSFunctionalProbeListReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPACSFunctionalProbeListReport *pVw = (CPACSFunctionalProbeListReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSFunctionalProbeListReport* )pWnd)->OnGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnGroupSelendokFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnGroupSelendok();
}
/*static void _OnGroupSetfocusFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnGroupKillfocus();
}*/
/*static void _OnGroupKillfocusFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnGroupKillfocus();
}*/
static long _OnGroupLoadDataFnc(CWnd *pWnd){
	return ((CPACSFunctionalProbeListReport *)pWnd)->OnGroupLoadData();
}
/*static void _OnGroupAddNewFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnGroupAddNew();
}*/
static void _OnPractitionerSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSFunctionalProbeListReport* )pWnd)->OnPractitionerSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnPractitionerSelendokFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnPractitionerSelendok();
}
/*static void _OnPractitionerSetfocusFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnPractitionerKillfocus();
}*/
/*static void _OnPractitionerKillfocusFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnPractitionerKillfocus();
}*/
static long _OnPractitionerLoadDataFnc(CWnd *pWnd){
	return ((CPACSFunctionalProbeListReport *)pWnd)->OnPractitionerLoadData();
}
/*static void _OnPractitionerAddNewFnc(CWnd *pWnd){
	((CPACSFunctionalProbeListReport *)pWnd)->OnPractitionerAddNew();
}*/
static int _OnAddPACSFunctionalProbeListReportFnc(CWnd *pWnd){
	 return ((CPACSFunctionalProbeListReport*)pWnd)->OnAddPACSFunctionalProbeListReport();
} 
static int _OnEditPACSFunctionalProbeListReportFnc(CWnd *pWnd){
	 return ((CPACSFunctionalProbeListReport*)pWnd)->OnEditPACSFunctionalProbeListReport();
} 
static int _OnDeletePACSFunctionalProbeListReportFnc(CWnd *pWnd){
	 return ((CPACSFunctionalProbeListReport*)pWnd)->OnDeletePACSFunctionalProbeListReport();
} 
static int _OnSavePACSFunctionalProbeListReportFnc(CWnd *pWnd){
	 return ((CPACSFunctionalProbeListReport*)pWnd)->OnSavePACSFunctionalProbeListReport();
} 
static int _OnCancelPACSFunctionalProbeListReportFnc(CWnd *pWnd){
	 return ((CPACSFunctionalProbeListReport*)pWnd)->OnCancelPACSFunctionalProbeListReport();
} 
CPACSFunctionalProbeListReport::CPACSFunctionalProbeListReport(CWnd* pParent)
{
	m_nDlgWidth = 475;
	m_nDlgHeight = 190;
	SetDefaultValues();
}
CPACSFunctionalProbeListReport::~CPACSFunctionalProbeListReport()
{
}
void CPACSFunctionalProbeListReport::OnCreateComponents()
{
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 465, 150);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 115, 55);
	m_wndYear.Create(this,120, 30, 240, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 245, 30, 335, 55);
	m_wndReportPeriod.Create(this,340, 30, 460, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 115, 85);
	m_wndFromDate.Create(this,120, 60, 240, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 245, 60, 335, 85);
	m_wndToDate.Create(this,340, 60, 460, 85); 
	
	m_wndGroupLabel.Create(this, _T("Group"), 10, 90, 115, 115);
	m_wndGroup.Create(this,120, 90, 460, 115); 
	m_wndPractitionerLabel.Create(this, _T("Practitioner"), 10, 120, 115, 145);
	m_wndPractitioner.Create(this,120, 120, 460, 145);

	//m_wndPrint.Create(this, _T("&Print"), 260, 155, 360, 180);
	m_wndExport.Create(this, _T("&Export"), 365, 155, 465, 180);
}
void CPACSFunctionalProbeListReport::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(FALSE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndGroup.SetCheckValue(true);
	m_wndGroup.LimitText(162);
	//m_wndPractitioner.SetCheckValue(true);
	m_wndPractitioner.LimitText(65);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndGroup.InsertColumn(0, _T("ID"), CFMT_TEXT, 60);
	m_wndGroup.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);


	m_wndPractitioner.InsertColumn(0, _T("ID"), CFMT_TEXT, 75);
	m_wndPractitioner.InsertColumn(1, _T("Name"), CFMT_TEXT, 235);

}
void CPACSFunctionalProbeListReport::OnSetWindowEvents(){
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
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndGroup.SetEvent(WE_SELENDOK, _OnGroupSelendokFnc);
	//m_wndGroup.SetEvent(WE_SETFOCUS, _OnGroupSetfocusFnc);
	//m_wndGroup.SetEvent(WE_KILLFOCUS, _OnGroupKillfocusFnc);
	m_wndGroup.SetEvent(WE_SELCHANGE, _OnGroupSelectChangeFnc);
	m_wndGroup.SetEvent(WE_LOADDATA, _OnGroupLoadDataFnc);
	//m_wndGroup.SetEvent(WE_ADDNEW, _OnGroupAddNewFnc);
	m_wndPractitioner.SetEvent(WE_SELENDOK, _OnPractitionerSelendokFnc);
	//m_wndPractitioner.SetEvent(WE_SETFOCUS, _OnPractitionerSetfocusFnc);
	//m_wndPractitioner.SetEvent(WE_KILLFOCUS, _OnPractitionerKillfocusFnc);
	m_wndPractitioner.SetEvent(WE_SELCHANGE, _OnPractitionerSelectChangeFnc);
	m_wndPractitioner.SetEvent(WE_LOADDATA, _OnPractitionerLoadDataFnc);
	//m_wndPractitioner.SetEvent(WE_ADDNEW, _OnPractitionerAddNewFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL;
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	m_szGroupKey = _T("B3200");

	UpdateData(false);
}
void CPACSFunctionalProbeListReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndGroup.GetDlgCtrlID(), m_szGroupKey);
	DDX_TextEx(pDX, m_wndPractitioner.GetDlgCtrlID(), m_szPractitionerKey);

}
void CPACSFunctionalProbeListReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPACSFunctionalProbeListReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CPACSFunctionalProbeListReport::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szGroupKey.Empty();
	m_szPractitionerKey.Empty();

}
int CPACSFunctionalProbeListReport::SetMode(int nMode){
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
/*void CPACSFunctionalProbeListReport::OnYearChange(){
	
} */
/*void CPACSFunctionalProbeListReport::OnYearSetfocus(){
	
} */
/*void CPACSFunctionalProbeListReport::OnYearKillfocus(){
	
} */
int CPACSFunctionalProbeListReport::OnYearCheckValue()
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
	UpdateData(FALSE);
	return 0;
} 
void CPACSFunctionalProbeListReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSFunctionalProbeListReport::OnReportPeriodSelendok()
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
/*void CPACSFunctionalProbeListReport::OnReportPeriodSetfocus(){
	
}*/
/*void CPACSFunctionalProbeListReport::OnReportPeriodKillfocus(){
	
}*/
long CPACSFunctionalProbeListReport::OnReportPeriodLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0)
	{
		szWhere.Format(_T(" WHERE hpr_idx=%d "), ToInt(m_szReportPeriodKey));
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
/*void CPACSFunctionalProbeListReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CPACSFunctionalProbeListReport::OnFromDateChange(){
	
} */
/*void CPACSFunctionalProbeListReport::OnFromDateSetfocus(){
	
} */
/*void CPACSFunctionalProbeListReport::OnFromDateKillfocus(){
	
} */
int CPACSFunctionalProbeListReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CPACSFunctionalProbeListReport::OnToDateChange(){
	
} */
/*void CPACSFunctionalProbeListReport::OnToDateSetfocus(){
	
} */
/*void CPACSFunctionalProbeListReport::OnToDateKillfocus(){
	
} */
int CPACSFunctionalProbeListReport::OnToDateCheckValue(){
	return 0;
}
void CPACSFunctionalProbeListReport::OnGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSFunctionalProbeListReport::OnGroupSelendok(){
	 
}
/*void CPACSFunctionalProbeListReport::OnGroupSetfocus(){
	
}*/
/*void CPACSFunctionalProbeListReport::OnGroupKillfocus(){
	
}*/
long CPACSFunctionalProbeListReport::OnGroupLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndGroup.IsSearchKey() && !m_szGroupKey.IsEmpty())
	{
		szWhere.Format(_T(" AND hfg_id='%s' "), m_szGroupKey);
	}

	m_wndGroup.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT hfg_id as id, hfg_name as name ") \
		         _T("FROM hms_fee_group ") \
				 _T("WHERE substr(hfg_id,1,2)='B3' AND hfg_id not in('B3100') ") \
				 _T("ORDER BY hfg_id"));

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndGroup.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPACSFunctionalProbeListReport::OnGroupAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CPACSFunctionalProbeListReport::OnPractitionerSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSFunctionalProbeListReport::OnPractitionerSelendok(){
	 
}
/*void CPACSFunctionalProbeListReport::OnPractitionerSetfocus(){
	
}*/
/*void CPACSFunctionalProbeListReport::OnPractitionerKillfocus(){
	
}*/
long CPACSFunctionalProbeListReport::OnPractitionerLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndPractitioner.IsSearchKey() && !m_szPractitionerKey.IsEmpty())
	{
		szWhere.Format(_T(" AND su_userid='%s' "), m_szPractitionerKey);
	}

	m_wndPractitioner.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as id, su_name as name ") \
		         _T("FROM sys_user ") \
				 _T("WHERE su_deptid in('A3','C7','C8') ") \
				 _T("ORDER BY su_userid"));

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndPractitioner.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPACSFunctionalProbeListReport::OnPractitionerAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CPACSFunctionalProbeListReport::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSFunctionalProbeListReport::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	CString szTemp, tmpStr;


	BeginWaitCursor();

	UpdateData(TRUE);

	szSQL = GetQueryString();

	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}

	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 9);
	xls.SetColumnWidth(2, 18);
	xls.SetColumnWidth(3, 5);
	xls.SetColumnWidth(4, 5);
	xls.SetColumnWidth(5, 21);
	xls.SetColumnWidth(6, 10);
	xls.SetColumnWidth(7, 16);
	xls.SetColumnWidth(8, 14);

	xls.SetColumnWidth(9, 13);
	xls.SetColumnWidth(10, 11);
	xls.SetColumnWidth(11, 7);
	xls.SetColumnWidth(12, 12);
	xls.SetColumnWidth(13, 7);
	xls.SetColumnWidth(14, 8);
	xls.SetColumnWidth(15, 11);

	int nRow = 1;
	int nCol = 0;

	xls.SetRowHeight(6, 54);

	xls.SetCellMergedColumns(0, 1, 4);
	xls.SetCellMergedColumns(0, 2, 4);

	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	
	xls.SetCellMergedColumns(nCol, nRow + 3, 16);
	xls.SetCellMergedColumns(nCol, nRow + 4, 16);

	xls.SetCellText(nCol, nRow + 3, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N L\xC0M \x43\x1EACN L\xC2M S\xC0NG"),
		            FMT_TEXT | FMT_CENTER, true, 16, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);

	TranslateString(_T("Index"), tmpStr);
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Document No"), tmpStr);
	xls.SetCellText(nCol + 1, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Patient Name"), tmpStr);
	xls.SetCellText(nCol + 2, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Age"), tmpStr);
	xls.SetCellText(nCol + 3, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Sex"), tmpStr);
	xls.SetCellText(nCol + 4, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Address"), tmpStr);
	xls.SetCellText(nCol + 5, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Object"), tmpStr);
	xls.SetCellText(nCol + 6, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 7, nRow + 5, _T("T\xEAn k\x1EF9 thu\x1EADt"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Result"), tmpStr);
	xls.SetCellText(nCol + 8, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 9, nRow + 5, _T("\x46orm \x43LS"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 10, nRow + 5, _T("Nh\xF3m \x43LS"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 11, nRow + 5, _T("Ph\xF2ng th\x1EF1\x63 hi\x1EC7n"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 12, nRow + 5, _T("\x42S th\x1EF1\x63 hi\x1EC7n"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 13, nRow + 5, _T("Ph\xF2ng \x63h\x1EC9 \x111\x1ECBnh"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 14, nRow + 5, _T("Kho\x61 \x63h\x1EC9 \x111\x1ECBnh"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 15, nRow + 5, _T("\x42S \x63h\x1EC9 \x111\x1ECBnh"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	int nIndex = 1;
	nRow = 6;
	CString szOldItemID, szNewItemID;

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("itemid"), szNewItemID);

		/*if (!szNewItemID.IsEmpty() && szOldItemID != szNewItemID)
		{
			++nRow;
			rs.GetValue(_T("itemname"), tmpStr);
			xls.SetCellMergedColumns(nCol, nRow, 16);
			xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_WRAPING, true);

			szOldItemID = szNewItemID;
			nIndex = 1;
		}*/

		++nRow;

		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("age"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("sex"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("address"), tmpStr);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("objectname"), tmpStr);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("itemname"), tmpStr);
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("pacsresult"), tmpStr);
		xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("formcls"), tmpStr);
		xls.SetCellText(nCol + 9, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("groupname"), tmpStr);
		xls.SetCellText(nCol + 10, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("proomname"), tmpStr);
		xls.SetCellText(nCol + 11, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("practitioner"), tmpStr);
		xls.SetCellText(nCol + 12, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("hpc_roomid"), tmpStr);
		xls.SetCellText(nCol + 13, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("hpc_deptid"), tmpStr);
		xls.SetCellText(nCol + 14, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("doctorname"), tmpStr);
		xls.SetCellText(nCol + 15, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.MoveNext();
	}

	xls.Save(_T("Exports\\DSBenhNhanCLS.xls"));
	EndWaitCursor();
} 
int CPACSFunctionalProbeListReport::OnAddPACSFunctionalProbeListReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CPACSFunctionalProbeListReport::OnEditPACSFunctionalProbeListReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CPACSFunctionalProbeListReport::OnDeletePACSFunctionalProbeListReport(){
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
 		OnCancelPACSFunctionalProbeListReport(); 
 	}
	return 0;
}
int CPACSFunctionalProbeListReport::OnSavePACSFunctionalProbeListReport(){
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
 		//OnPACSFunctionalProbeListReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CPACSFunctionalProbeListReport::OnCancelPACSFunctionalProbeListReport(){
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
int CPACSFunctionalProbeListReport::OnPACSFunctionalProbeListReportListLoadData(){
	return 0;
}

CString CPACSFunctionalProbeListReport::GetQueryString()
{
	CString szSQL, szWhere;
	szSQL.Empty();

	if (!m_szGroupKey.IsEmpty())
		szWhere.AppendFormat(_T(" and hpc_groupid='%s' "), m_szGroupKey);

	if (!m_szPractitionerKey.IsEmpty())
		szWhere.AppendFormat(_T(" and hpc_practitioner='%s' "), m_szPractitionerKey);

	szSQL.Format(_T(" select hd_docno as docno,") \
				_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("        hms_getaddress(hp_provid, hp_distid, hp_villid) as address,") \
				_T("        hms_getage(hd_admitdate, hp_birthdate) as age,") \
				_T("        (select distinct ss_desc from sys_sel where ss_id='sys_sex' and ss_code=hp_sex) as sex,") \
				_T("        ho_desc as objectname,") \
				_T("        hpcl_itemid as itemid,") \
				_T("        hfl_name as itemname,") \
				_T("        hpr_desc as pacsresult,") \
				_T("        hpcl_result as formcls,") \
				_T("        hfg_name as groupname,") \
				_T("        hpc_pdeptid,") \
				_T("        hrl_name as proomname,") \
				_T("        (select distinct su_name from sys_user where su_userid=hpc_practitioner) as practitioner,") \
				_T("        hpc_deptid,") \
				_T("        hpc_roomid,") \
				_T("        (select distinct su_name from sys_user where su_userid=hpc_doctor) as doctorname,") \
				_T("        case when ho_type in('I','C') then 1") \
				_T("                  when ho_type in('S') then 2") \
				_T("        else 3 end as objectidx ") \
				_T(" from hms_patient") \
				_T(" left join hms_doc on(hd_patientno=hp_patientno)") \
				_T(" left join hms_object on(ho_id=hd_object)") \
				_T(" left join hms_pacsorder on(hpc_docno=hd_docno)") \
				_T(" left join hms_pacsorderline on(hpcl_orderid=hpc_orderid and hpcl_docno=hpc_docno)") \
				_T(" left join hms_pacs_result on(hpr_itemid=hpcl_itemid and hpr_orderid=hpc_orderid") \
				_T("                              and hpcl_docno=hpr_docno and hpr_name='Conclusion')") \
				_T(" left join hms_fee_list on(hfl_feeid=hpcl_itemid)") \
				_T(" left join hms_fee_group on(hfg_id=hpc_groupid)") \
				_T(" left join hms_roomlist on(hrl_id=hpcl_proomid and hrl_deptid='C7')") \
				_T(" where hpcl_status='T' and hpcl_date between cast_string2timestamp('%s')") \
				_T(" 		  and cast_string2timestamp('%s') and hpc_groupid not in('B3100') %s") \
				_T(" order by hpcl_date, hpc_practitioner"),
				m_szFromDate, m_szToDate, szWhere);


	return szSQL;
}