#include "stdafx.h"
#include "FMInsuranceReport21a.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMInsuranceReport21a *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceReport21a *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceReport21a *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMInsuranceReport21a *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMInsuranceReport21a* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd)
{
	((CFMInsuranceReport21a *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceReport21a *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceReport21a *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceReport21a *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMInsuranceReport21a *)pWnd)->OnReportPeriodAddNew();
}*/
static long _OnPatientLineLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceReport21a *)pWnd)->OnPatientLineLoadData();
}
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMInsuranceReport21a *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceReport21a *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceReport21a *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMInsuranceReport21a *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMInsuranceReport21a *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceReport21a *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceReport21a *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMInsuranceReport21a *)pWnd)->OnToDateCheckValue();
} 
static void _OnDrugPTTTSelectFnc(CWnd *pWnd){
	 ((CFMInsuranceReport21a*)pWnd)->OnDrugPTTTSelect();
}
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	 ((CFMInsuranceReport21a*)pWnd)->OnOutPatientSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	 ((CFMInsuranceReport21a*)pWnd)->OnInPatientSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMInsuranceReport21a *pVw = (CFMInsuranceReport21a *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMInsuranceReport21a *pVw = (CFMInsuranceReport21a *)pWnd;
	pVw->OnExportSelect();
}
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceReport21a*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CFMInsuranceReport21a*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMInsuranceReport21a*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMInsuranceReport21a*)pWnd)->OnListDeleteItem();
}
static int _OnAddFMInsuranceReport21aFnc(CWnd *pWnd){
	 return ((CFMInsuranceReport21a*)pWnd)->OnAddFMInsuranceReport21a();
} 
static int _OnEditFMInsuranceReport21aFnc(CWnd *pWnd){
	 return ((CFMInsuranceReport21a*)pWnd)->OnEditFMInsuranceReport21a();
} 
static int _OnDeleteFMInsuranceReport21aFnc(CWnd *pWnd){
	 return ((CFMInsuranceReport21a*)pWnd)->OnDeleteFMInsuranceReport21a();
} 
static int _OnSaveFMInsuranceReport21aFnc(CWnd *pWnd){
	 return ((CFMInsuranceReport21a*)pWnd)->OnSaveFMInsuranceReport21a();
} 
static int _OnCancelFMInsuranceReport21aFnc(CWnd *pWnd){
	 return ((CFMInsuranceReport21a*)pWnd)->OnCancelFMInsuranceReport21a();
}

static int _OnListCheckAllFnc(CWnd *pWnd){
	return ((CFMInsuranceReport21a*)pWnd)->OnListCheckAll();
}
static int _OnListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMInsuranceReport21a*)pWnd)->OnListUnCheckAll();
}

CFMInsuranceReport21a::CFMInsuranceReport21a(CWnd *pParent)
{
	m_nDlgWidth = 540;
	m_nDlgHeight = 505;
	SetDefaultValues();
}
CFMInsuranceReport21a::~CFMInsuranceReport21a(){
}
void CFMInsuranceReport21a::OnCreateComponents()
{
	m_wndReportFilter.Create(this, _T("Report Parameters"), 5, 5, 575, 545);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 120, 570, 510);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 270, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Month"), 275, 30, 365, 55);
	m_wndReportPeriod.Create(this,370, 30, 570, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 270, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 275, 60, 365, 85);
	m_wndToDate.Create(this,370, 60, 570, 85); 
	m_wndPatientLineLabel.Create(this, _T("Patient Line"), 10, 90, 90, 115);
	m_wndPatientLine.Create(this,95, 90, 570, 115); 
	m_wndDrugPTTT.Create(this, _T("Without Operation Material"), 205, 515, 385, 540);
	m_wndOutPatient.Create(this, _T("OutPatient"), 390, 515, 475, 540);
	m_wndInPatient.Create(this, _T("InPatient"), 480, 515, 570, 540);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 370, 550, 470, 575);
	m_wndExport.Create(this, _T("&Export"), 475, 550, 575, 575);
	m_wndList.Create(this,15, 145, 565, 505); 

}
void CFMInsuranceReport21a::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);

	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 180);

	m_wndPatientLine.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndPatientLine.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndList.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);
	m_wndList.SetCheckBox(TRUE);
}
void CFMInsuranceReport21a::OnSetWindowEvents()
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
	m_wndPatientLine.SetEvent(WE_LOADDATA, _OnPatientLineLoadDataFnc);
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndDrugPTTT.SetEvent(WE_CLICK, _OnDrugPTTTSelectFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);

	m_wndList.SetWindowText(_T("Dept"));
	m_wndList.AddEvent(1, _T("Check All"), _OnListCheckAllFnc);
	m_wndList.AddEvent(2, _T("UnCheck All"), _OnListUnCheckAllFnc);
	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

	OnListLoadData();
}
void CFMInsuranceReport21a::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndPatientLine.GetDlgCtrlID(), m_szPatientLineKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_Check(pDX, m_wndDrugPTTT.GetDlgCtrlID(), m_bCheckDrug);
	DDX_Check(pDX, m_wndOutPatient.GetDlgCtrlID(), m_bOutPatient);
	DDX_Check(pDX, m_wndInPatient.GetDlgCtrlID(), m_bInPatient);
	
}
void CFMInsuranceReport21a::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMInsuranceReport21a::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMInsuranceReport21a::SetDefaultValues(){

	m_nYear = 0;
	m_szReportPeriodKey.Empty();
	m_szPatientLineKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_bCheckDrug=FALSE;
	m_bOutPatient=TRUE;
	m_bInPatient=FALSE;

}
int CFMInsuranceReport21a::SetMode(int nMode){
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
/*void CFMInsuranceReport21a::OnYearChange(){
	
} */
/*void CFMInsuranceReport21a::OnYearSetfocus(){
	
} */
/*void CFMInsuranceReport21a::OnYearKillfocus(){
	
} */
int CFMInsuranceReport21a::OnYearCheckValue()
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
void CFMInsuranceReport21a::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMInsuranceReport21a::OnReportPeriodSelendok()
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
/*void CFMInsuranceReport21a::OnReportPeriodSetfocus(){
	
}*/
/*void CFMInsuranceReport21a::OnReportPeriodKillfocus(){
	
}*/
long CFMInsuranceReport21a::OnReportPeriodLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0)
	{
		szWhere.Format(_T(" WHERE hpr_idx = %ld"), ToInt(m_szReportPeriodKey));
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
/*void CFMInsuranceReport21a::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
long CFMInsuranceReport21a::OnPatientLineLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndPatientLine.IsSearchKey() && ToInt(m_szPatientLineKey) > 0)
	{
		szWhere.Format(_T(" AND ss_code = %d"), ToInt(m_szPatientLineKey));
	}
	m_wndPatientLine.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT ss_code code, ss_desc descr FROM sys_sel WHERE ss_id = 'hms_patient_line' ORDER BY ss_code"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndPatientLine.AddItems(
			rs.GetValue(_T("code")), 
			rs.GetValue(_T("descr")),			
			NULL);
		rs.MoveNext();
	}
	return nCount;	
}
/*void CFMInsuranceReport21a::OnFromDateChange(){
	
} */
/*void CFMInsuranceReport21a::OnFromDateSetfocus(){
	
} */
/*void CFMInsuranceReport21a::OnFromDateKillfocus(){
	
} */
int CFMInsuranceReport21a::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMInsuranceReport21a::OnToDateChange(){
	
} */
/*void CFMInsuranceReport21a::OnToDateSetfocus(){
	
} */
/*void CFMInsuranceReport21a::OnToDateKillfocus(){
	
} */
int CFMInsuranceReport21a::OnToDateCheckValue()
{
	return 0;
} 
void CFMInsuranceReport21a::OnDrugPTTTSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMInsuranceReport21a::OnOutPatientSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bInPatient = FALSE;
	OnListLoadData();
	UpdateData(FALSE);
}
void CFMInsuranceReport21a::OnInPatientSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bOutPatient = FALSE;
	OnListLoadData();
	UpdateData(FALSE);
}
void CFMInsuranceReport21a::OnListDblClick(){
	
} 
void CFMInsuranceReport21a::OnListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMInsuranceReport21a::OnListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMInsuranceReport21a::OnListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	m_wndList.BeginLoad(); 
	m_wndList.DeleteAllItems(); 
	int nCount = 0;

	if (m_bOutPatient)
		szWhere.Format(_T(" AND sd_type IN('KB') "));

	if (m_bInPatient)
		szWhere.Format(_T(" AND sd_type IN('DT') "));

	if (m_bOutPatient && m_bInPatient)
		szWhere.Format(_T(" AND sd_type IN('KB','DT') "));

	if (!m_bOutPatient && !m_bInPatient)
		szWhere.Format(_T(" AND 0=1 "));

	szSQL.Format(_T("SELECT sd_id AS ID, sd_name AS Name ") \
		         _T("FROM sys_dept ") \
				 _T("WHERE 1=1 %s ") \
				 _T("ORDER BY sd_id"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF())
	{ 
		m_wndList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
}
void CFMInsuranceReport21a::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString tmpStr, szSQL, szDate, szMoney;
	CString szTemp;

	if (!rpt.Init(_T("Reports/HMS/HF_THONGKETONGHOPDICHVUKYTHUATTHEOQUY.RPT")) )	
	{
		return;
	}

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);

	rpt.GetReportHeader()->SetValue(_T("ObjectGroup"), _T(""));
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), 
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDate::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	//Page Header
	//Report Detail
	int nIndex = 1;
	CString szOldLine, szNewLine, szAmount;
	CReportSection* rptDetail;

	long double grpCost = 0, ttlCost = 0, totalCost = 0;
	double cost = 0;
	
	CStringArray arrFeeGrp, arrTitle;

	arrFeeGrp.Add(_T("Test"));
	arrTitle.Add(_T("A"));

	arrFeeGrp.Add(_T("Diagnostic image"));
	arrTitle.Add(_T("B"));

	arrFeeGrp.Add(_T("Visiting the function rolls"));
	arrTitle.Add(_T("C"));

	arrFeeGrp.Add(_T("Surgery - procedures"));
	arrTitle.Add(_T("D"));

	arrFeeGrp.Add(_T("Consumables"));
	arrTitle.Add(_T("E"));

	arrFeeGrp.Add(_T("High-tech services"));
	arrTitle.Add(_T("F"));

	if (m_bOutPatient)
		arrFeeGrp.Add(_T("Exam Fee"));
	else
		arrFeeGrp.Add(_T("Bed Fee"));
	arrTitle.Add(_T("G"));
		
	nIndex = 1;
	int nCount = 0, nTemp = 0;

	BeginWaitCursor();

	szNewLine.Empty();
	szOldLine.Empty();


	for (int i = 0; i < arrFeeGrp.GetSize(); i++)
	{
		if (m_bInPatient)
		{
			if (i == 6)
				szSQL = GetQueryString(i + 1);
			else if (i == 4)
				szSQL = GetQueryString(i * 2);
			else
				szSQL = GetQueryString(i);
		}
		else
			szSQL = GetQueryString(i);
		//szSQL = GetQueryString(i);
		rs.ExecSQL(szSQL);
		nCount = rs.GetRecordCount();

		szNewLine.Empty();
		szOldLine.Empty();

		//Phi kham
		if (i == 6)
		{
			if (!rs.IsEOF())
			{
				grpCost = ttlCost = cost = ttlCost = 0;
				nIndex = 1;

				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(false);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetItalic(false);

				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));

				TranslateString(arrFeeGrp[i], tmpStr);
				rptDetail->SetValue(_T("GroupName"), tmpStr);
				rptDetail->SetValue(_T("Group"), arrTitle[i]);		
				
				if (m_bOutPatient)
				{
					while (!rs.IsEOF())
					{			
						rptDetail = rpt.AddDetail();
						tmpStr.Format(_T("%d"), nIndex++);
						rptDetail->SetValue(_T("1"), tmpStr);

						rs.GetValue(_T("examname"), tmpStr);
						rptDetail->SetValue(_T("2"), tmpStr);

						rs.GetValue(_T("soluong"), tmpStr);
						rptDetail->SetValue(_T("3"), tmpStr);

						rs.GetValue(_T("insprice"), cost);
						FormatCurrency(cost, tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);

						rs.GetValue(_T("suminsprice"), cost);
						grpCost += cost;
						FormatCurrency(cost, tmpStr);
						rptDetail->SetValue(_T("5"), tmpStr);
						
						rs.MoveNext();
					}

					ttlCost += grpCost;
					
					if (ttlCost > 0)
					{
						CString szField, szAmount;
						TranslateString(_T("Total Amount"), szAmount);
						rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
						rptDetail->SetValue(_T("TotalGroup"), szAmount + _T("(G)"));						
						FormatCurrency(ttlCost, tmpStr);
						rptDetail->SetValue(_T("s5"), tmpStr);
						totalCost+=ttlCost;
					}
				}
				else
				{
					CString szOldRoom, szNewRoom;
					//CString szOldLine, szNewLine;
					int nQty = 0, nTotalQty = 0;

					while (!rs.IsEOF())
					{
						rs.GetValue(_T("deptname"), szNewLine);
						if (szNewLine != szOldLine && !szNewLine.IsEmpty())
						{
							if (grpCost > 0)
							{	
								rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
								rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng kho\x61:"));
								FormatCurrency(grpCost, tmpStr);
								rptDetail->SetValue(_T("s5"), tmpStr);
								ttlCost += grpCost;
								nIndex = 1;
								grpCost = 0;
							}
							//nRow++;
							rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
							rptDetail->SetValue(_T("GroupName"), szNewLine);
							szOldLine = szNewLine;
							szTemp.Empty();
						}

						rs.GetValue(_T("roomname"), szNewRoom);
						if (szNewRoom != szOldRoom && !szNewRoom.IsEmpty())
						{
							//nRow++;
							//xls.SetCellText(1, nRow, szNewRoom, FMT_TEXT, true);
							rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
							rptDetail->SetValue(_T("GroupName"), szNewRoom);
							szOldRoom = szNewRoom;
							szTemp.Empty();
						}

						//nRow++;
						rptDetail = rpt.AddDetail();
						tmpStr.Format(_T("%d"), nIndex++);
						rptDetail->SetValue(_T("1"), tmpStr);

						rs.GetValue(_T("bedname"), tmpStr);
						rptDetail->SetValue(_T("2"), tmpStr);

						rs.GetValue(_T("treatqty"), nQty);
						nTotalQty += nQty;
						tmpStr.Format(_T("%d"), nQty);
						rptDetail->SetValue(_T("3"), tmpStr);

						rs.GetValue(_T("insprice"), cost);
						FormatCurrency(cost, tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);

						rs.GetValue(_T("total"), cost);
						//nLineCost += cost;
						grpCost += cost;
						FormatCurrency(cost, tmpStr);
						rptDetail->SetValue(_T("5"), tmpStr);
						
						rs.MoveNext();
					}

					if (grpCost > 0)
					{
						//nRow++;
						//TranslateString(_T("C\x1ED9ng:"), tmpStr);
						rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
						rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng kho\x61:"));
						FormatCurrency(grpCost, tmpStr);
						rptDetail->SetValue(_T("s5"), tmpStr);
						ttlCost += grpCost;
					}

					if (ttlCost > 0)
					{
						CString szField, szAmount;
						TranslateString(_T("Total Amount"), szAmount);
						rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
						rptDetail->SetValue(_T("TotalGroup"), szAmount + _T("(G)"));						
						FormatCurrency(ttlCost, tmpStr);
						rptDetail->SetValue(_T("s5"), tmpStr);
						totalCost += ttlCost;
					}
				}
			}
		}

		//Phi xet nghiem
		if (i == 0)
		{
			if (!rs.IsEOF())
			{	
				grpCost = ttlCost = cost = ttlCost = 0;
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(false);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetItalic(false);
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				TranslateString(arrFeeGrp[i], tmpStr);
				rptDetail->SetValue(_T("GroupName"), tmpStr);
				rptDetail->SetValue(_T("Group"), arrTitle[i]);		
				
				while (!rs.IsEOF())
				{
					rs.GetValue(_T("groupname"), szNewLine);
					if (szNewLine != szOldLine && !szNewLine.IsEmpty())
					{
						CString szField;
						if (grpCost > 0)
						{				
							TranslateString(_T("Total Group"), tmpStr);
							rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
							FormatCurrency(grpCost, tmpStr);
							rptDetail->SetValue(_T("s5"), tmpStr);
							ttlCost += grpCost;
							nIndex = 1;
							grpCost = 0;				
						}
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(10);
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(true);
						rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));				
						rptDetail->SetValue(_T("GroupName"), szNewLine);
						szOldLine = szNewLine;
					}

					rptDetail = rpt.AddDetail();

					tmpStr.Format(_T("%d"), nIndex++);
					rptDetail->SetValue(_T("1"), tmpStr);

					rs.GetValue(_T("testname"), tmpStr);
					rptDetail->SetValue(_T("2"), tmpStr);

					rs.GetValue(_T("qty"), tmpStr);					
					rptDetail->SetValue(_T("3"), tmpStr);

					rs.GetValue(_T("insprice"), cost);		
					FormatCurrency(cost, tmpStr);
					rptDetail->SetValue(_T("4"), tmpStr);

					rs.GetValue(_T("suminsprice"), cost);
					grpCost += cost;
					FormatCurrency(cost, tmpStr);
					rptDetail->SetValue(_T("5"), tmpStr);
					
					rs.MoveNext();
				}
				
				if (grpCost > 0)
				{
					CString szField, szAmount;		
					TranslateString(_T("Total Group"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));		
					FormatCurrency(grpCost, tmpStr);
					rptDetail->SetValue(_T("s5"), tmpStr);	
					ttlCost += grpCost;		
				}

				if (ttlCost > 0)
				{
					CString szField, szAmount;
					TranslateString(_T("Total Amount"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
					rptDetail->SetValue(_T("TotalGroup"), szAmount + _T("(A)"));						
					FormatCurrency(ttlCost, tmpStr);
					rptDetail->SetValue(_T("S5"), tmpStr);	
					totalCost+=ttlCost;
				}
			}
		}


		//Phi chan doan hinh anh
		if (i == 1)
		{
			if (!rs.IsEOF())
			{	
				grpCost = ttlCost = cost = ttlCost = 0;
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(false);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetItalic(false);
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				TranslateString(arrFeeGrp[i], tmpStr);
				rptDetail->SetValue(_T("GroupName"), tmpStr);
				rptDetail->SetValue(_T("Group"), arrTitle[i]);		
				
				while (!rs.IsEOF())
				{
					rs.GetValue(_T("groupname"), szNewLine);
					if(szNewLine != szOldLine && !szNewLine.IsEmpty())
					{
						CString szField;
						if(grpCost > 0)
						{				
							TranslateString(_T("Total Group"), tmpStr);
							rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
							FormatCurrency(grpCost, tmpStr);
							rptDetail->SetValue(_T("s5"), tmpStr);
							ttlCost += grpCost;
							nIndex = 1;
							grpCost = 0;				
						}
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(10);
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(true);
						rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));				
						rptDetail->SetValue(_T("GroupName"), szNewLine);
						szOldLine = szNewLine;
					}

					rptDetail = rpt.AddDetail();

					tmpStr.Format(_T("%d"), nIndex++);
					rptDetail->SetValue(_T("1"), tmpStr);

					rs.GetValue(_T("testname"), tmpStr);
					rptDetail->SetValue(_T("2"), tmpStr);

					rs.GetValue(_T("qty"), tmpStr);					
					rptDetail->SetValue(_T("3"), tmpStr);

					rs.GetValue(_T("insprice"), cost);		
					FormatCurrency(cost, tmpStr);
					rptDetail->SetValue(_T("4"), tmpStr);

					rs.GetValue(_T("suminsprice"), cost);
					grpCost += cost;
					FormatCurrency(cost, tmpStr);
					rptDetail->SetValue(_T("5"), tmpStr);
					
					rs.MoveNext();
				}
				
				if (grpCost > 0)
				{
					CString szField, szAmount;		
					TranslateString(_T("Total Group"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));		
					FormatCurrency(grpCost, tmpStr);
					rptDetail->SetValue(_T("s5"), tmpStr);	
					ttlCost += grpCost;
				}

				if (ttlCost > 0)
				{
					CString szField, szAmount;
					TranslateString(_T("Total Amount"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
					rptDetail->SetValue(_T("TotalGroup"), szAmount + _T("(B)"));						
					FormatCurrency(ttlCost, tmpStr);
					rptDetail->SetValue(_T("S5"), tmpStr);	
					totalCost += ttlCost;
				}
			}
		}

		//Phi tham do chuc nang
		if (i == 2)
		{
			if (!rs.IsEOF())
			{	
				grpCost = ttlCost = cost = ttlCost = 0;
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(false);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetItalic(false);
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				TranslateString(arrFeeGrp[i], tmpStr);
				rptDetail->SetValue(_T("GroupName"), tmpStr);
				rptDetail->SetValue(_T("Group"), arrTitle[i]);		
				
				while (!rs.IsEOF())
				{
					rs.GetValue(_T("groupname"), szNewLine);
					if(szNewLine != szOldLine && !szNewLine.IsEmpty())
					{
						CString szField;
						if(grpCost > 0)
						{				
							TranslateString(_T("Total Group"), tmpStr);
							rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
							FormatCurrency(grpCost, tmpStr);
							rptDetail->SetValue(_T("s5"), tmpStr);
							ttlCost += grpCost;
							nIndex = 1;
							grpCost = 0;				
						}
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(10);
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(true);
						rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));				
						rptDetail->SetValue(_T("GroupName"), szNewLine);
						szOldLine = szNewLine;
					}

					rptDetail = rpt.AddDetail();

					tmpStr.Format(_T("%d"), nIndex++);
					rptDetail->SetValue(_T("1"), tmpStr);

					rs.GetValue(_T("testname"), tmpStr);
					rptDetail->SetValue(_T("2"), tmpStr);

					rs.GetValue(_T("qty"), tmpStr);					
					rptDetail->SetValue(_T("3"), tmpStr);

					rs.GetValue(_T("insprice"), cost);		
					FormatCurrency(cost, tmpStr);
					rptDetail->SetValue(_T("4"), tmpStr);

					rs.GetValue(_T("suminsprice"), cost);
					grpCost += cost;
					FormatCurrency(cost, tmpStr);
					rptDetail->SetValue(_T("5"), tmpStr);
					
					rs.MoveNext();
				}
				
				if (grpCost > 0)
				{
					CString szField, szAmount;		
					TranslateString(_T("Total Group"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));		
					FormatCurrency(grpCost, tmpStr);
					rptDetail->SetValue(_T("s5"), tmpStr);	
					ttlCost += grpCost;		
				}

				if (ttlCost > 0)
				{
					CString szField, szAmount;
					TranslateString(_T("Total Amount"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
					rptDetail->SetValue(_T("TotalGroup"), szAmount + _T("(C)"));					
					FormatCurrency(ttlCost, tmpStr);
					rptDetail->SetValue(_T("S5"), tmpStr);	
					totalCost+=ttlCost;
				}
			}
		}


		//Phi phau thuat thu thuat
		if (i == 3)
		{
			if (!rs.IsEOF())
			{
				grpCost = ttlCost = cost = ttlCost = 0;
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(false);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetItalic(false);
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				TranslateString(arrFeeGrp[i], tmpStr);
				rptDetail->SetValue(_T("GroupName"), tmpStr);
				rptDetail->SetValue(_T("Group"), arrTitle[i]);
				nIndex = 1;
				while (!rs.IsEOF())
				{
					rs.GetValue(_T("operatype"), szNewLine);
					if(szNewLine != szOldLine && !szNewLine.IsEmpty())
					{
						CString szField;
						if (grpCost > 0)
						{				
							TranslateString(_T("Total Group"), tmpStr);
							rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
							FormatCurrency(grpCost, tmpStr);
							rptDetail->SetValue(_T("s5"), tmpStr);
							ttlCost += grpCost;
							nIndex = 1;
							grpCost = 0;				
						}
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(10);
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(true);
						rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));				
						rptDetail->SetValue(_T("GroupName"), szNewLine);
						szOldLine = szNewLine;
					}

					rptDetail = rpt.AddDetail();

					tmpStr.Format(_T("%d"), nIndex++);
					rptDetail->SetValue(_T("1"), tmpStr);

					rs.GetValue(_T("itemname"), tmpStr);
					rptDetail->SetValue(_T("2"), tmpStr);

					rs.GetValue(_T("qty"), tmpStr);
					rptDetail->SetValue(_T("3"), tmpStr);

					rs.GetValue(_T("insprice"), cost);
					FormatCurrency(cost, tmpStr);
					rptDetail->SetValue(_T("4"), tmpStr);

					rs.GetValue(_T("amount"), cost);
					grpCost += cost;
					FormatCurrency(cost, tmpStr);
					rptDetail->SetValue(_T("5"), tmpStr);		
					rs.MoveNext();
				}
				
				if (grpCost > 0)
				{
					CString szField, szAmount;		
					TranslateString(_T("Total Amount"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));		
					FormatCurrency(grpCost, tmpStr);
					rptDetail->SetValue(_T("s5"), tmpStr);
					ttlCost += grpCost;
				}

				if (ttlCost > 0)
				{
					CString szField, szAmount;
					TranslateString(_T("Total Amount"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
					rptDetail->SetValue(_T("TotalGroup"), szAmount + _T("(D)"));						
					FormatCurrency(ttlCost, tmpStr);
					rptDetail->SetValue(_T("s5"), tmpStr);
					totalCost += ttlCost;
				}
			}
		}


		//Phi vat tu y te tieu hao
		if (i == 4)
		{
			if (!rs.IsEOF())
			{
				grpCost = ttlCost = cost = ttlCost = 0;
				nIndex = 1;
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(false);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetItalic(false);

				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));

				TranslateString(arrFeeGrp[i], tmpStr);
				rptDetail->SetValue(_T("GroupName"), tmpStr);
				rptDetail->SetValue(_T("Group"), arrTitle[i]);
						
				while (!rs.IsEOF())
				{
					rs.GetValue(_T("typeid"), szNewLine);
					if (szNewLine != szOldLine && !szNewLine.IsEmpty())
					{
						CString szField;
						if(grpCost > 0)
						{
							TranslateString(_T("Total Group"), szAmount);
							rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
							rptDetail->SetValue(_T("TotalGroup"), szAmount);						
							FormatCurrency(grpCost, tmpStr);
							rptDetail->SetValue(_T("s5"), tmpStr);
							ttlCost += grpCost;						
							grpCost = 0;								
						}

						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(10);
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(true);	
						rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
						rptDetail->SetValue(_T("GroupName"), szNewLine);
						szOldLine = szNewLine;
						nIndex = 1;
					}

					rptDetail = rpt.AddDetail();

					tmpStr.Format(_T("%d"), nIndex++);
					rptDetail->SetValue(_T("1"), tmpStr);

					rs.GetValue(_T("drugname"), tmpStr);
					rptDetail->SetValue(_T("2"), tmpStr);
					
					rs.GetValue(_T("qty"), tmpStr);
					rptDetail->SetValue(_T("3"), tmpStr);

					rs.GetValue(_T("price"), cost);
					FormatCurrency(cost, tmpStr);
					rptDetail->SetValue(_T("4"), tmpStr);
					
					rs.GetValue(_T("amount"), cost);	
					grpCost += cost;
					FormatCurrency(cost, tmpStr);
					rptDetail->SetValue(_T("5"), tmpStr);

					rs.MoveNext();
				}
				
				if (grpCost > 0)
				{	
					TranslateString(_T("Total Group"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));	
					rptDetail->SetValue(_T("TotalGroup"), szAmount);
					FormatCurrency(grpCost, tmpStr);
					rptDetail->SetValue(_T("s5"), tmpStr);
					ttlCost += grpCost;
				}
				if (ttlCost > 0)
				{				
					TranslateString(_T("Total Amount"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
					rptDetail->SetValue(_T("TotalGroup"), szAmount +_T("(E)"));
					FormatCurrency(ttlCost, tmpStr);
					rptDetail->SetValue(_T("s5"), tmpStr);
					totalCost += ttlCost;
				}
			}
		}


		//Phi dich vu ky thuat cao
		if (i == 5)
		{
			if (!rs.IsEOF())
			{
				grpCost = ttlCost = cost = ttlCost = 0;
				nIndex = 1;
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(false);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetFaceSize(12);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetBold(true);
				rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetItalic(false);
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				TranslateString(arrFeeGrp[i], tmpStr);
				rptDetail->SetValue(_T("GroupName"), tmpStr);
				rptDetail->SetValue(_T("Group"), arrTitle[i]);
				
				while (!rs.IsEOF())
				{
					rs.GetValue(_T("groupname"), szNewLine);

					if(szNewLine != szOldLine && !szNewLine.IsEmpty())
					{
						CString szField;
						if(grpCost > 0)
						{				
							TranslateString(_T("Total Group"), tmpStr);
							rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
							FormatCurrency(grpCost, tmpStr);
							rptDetail->SetValue(_T("s5"), tmpStr);
							ttlCost += grpCost;
							nIndex = 1;
							grpCost = 0;				
						}
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(10);
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
						rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(true);
						rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));				
						rptDetail->SetValue(_T("GroupName"), szNewLine);
						szOldLine = szNewLine;
					}

					rptDetail = rpt.AddDetail();

					tmpStr.Format(_T("%d"), nIndex++);
					rptDetail->SetValue(_T("1"), tmpStr);

					rs.GetValue(_T("itemname"), tmpStr);
					rptDetail->SetValue(_T("2"), tmpStr);

					rs.GetValue(_T("qty"), tmpStr);					
					rptDetail->SetValue(_T("3"), tmpStr);

					rs.GetValue(_T("insprice"), cost);		
					FormatCurrency(cost, tmpStr);
					rptDetail->SetValue(_T("4"), tmpStr);

					rs.GetValue(_T("amount"), cost);
					grpCost += cost;
					FormatCurrency(cost, tmpStr);
					rptDetail->SetValue(_T("5"), tmpStr);
					
					rs.MoveNext();
				}
				
				if (grpCost > 0)
				{
					CString szField, szAmount;		
					TranslateString(_T("Total Amount"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));		
					FormatCurrency(grpCost, tmpStr);
					rptDetail->SetValue(_T("s5"), tmpStr);

					ttlCost += grpCost;
				}
				if (ttlCost > 0)
				{
					CString szField, szAmount;
					TranslateString(_T("Total Amount"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
					rptDetail->SetValue(_T("TotalGroup"), szAmount + _T("(F)"));				
					FormatCurrency(ttlCost, tmpStr);
					rptDetail->SetValue(_T("s5"), tmpStr);
					totalCost+=ttlCost;
				}

			}
		}
	}

	if (totalCost > 0)
	{
		CString szField, szAmount;
		TranslateString(_T("Total Amount"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), szAmount + _T("(A+B+C+D+E+F+G)"));						
		FormatCurrency(totalCost, tmpStr);
		rptDetail->SetValue(_T("s5"), tmpStr);
	}
	
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
		          tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);		
	MoneyToString(szMoney, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	EndWaitCursor();
	rpt.PrintPreview();
} 
void CFMInsuranceReport21a::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);

	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, szOldLine, szNewLine;
	CString tmpStr, szTemp;

	int nCol = 0, nRow = 0, nIndex = 1, nCount = 0, nTemp = 0;
	double cost = 0;
	long double grpCost = 0, ttlCost = 0, totalCost = 0;

	BeginWaitCursor();

	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 8);
	xls.SetColumnWidth(1, 50);
	xls.SetColumnWidth(2, 8);
	xls.SetColumnWidth(3, 12);
	xls.SetColumnWidth(4, 15);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellMergedColumns(nCol, nRow + 1, 2);

	xls.SetCellMergedColumns(nCol + 2, nRow, 5);
	xls.SetCellMergedColumns(nCol + 2, nRow + 1, 5);

	xls.SetCellMergedColumns(nCol, nRow + 2, 5);
	xls.SetCellMergedColumns(nCol, nRow + 3, 5);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 2, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 2, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	TranslateString(_T("General statistics and technical services for patients using medical insurance"), szTemp);
	StringUpper(szTemp, tmpStr);
	xls.SetCellText(nCol, nRow + 2, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);
	
	//Column Header
	CStringArray arrCol, arrFeeGrp, arrTitle;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn \x64\x1ECB\x63h v\x1EE5 k\x1EF9 thu\x1EADt"));
	arrCol.Add(_T("S\x1ED1 l\x1B0\x1EE3ng"));
	arrCol.Add(_T("\x110\x1A1n gi\xE1"));
	arrCol.Add(_T("Th\xE0nh ti\x1EC1n"));

	nRow = 4;

	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol + i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10); 
	}

	nRow = 5;

	arrFeeGrp.Add(_T("Test"));
	arrTitle.Add(_T("A"));

	arrFeeGrp.Add(_T("Diagnostic image"));
	arrTitle.Add(_T("B"));

	arrFeeGrp.Add(_T("Visiting the function rolls"));
	arrTitle.Add(_T("C"));

	arrFeeGrp.Add(_T("Surgery - procedures"));
	arrTitle.Add(_T("D"));

	arrFeeGrp.Add(_T("Consumables"));
	arrTitle.Add(_T("E"));

	arrFeeGrp.Add(_T("High-tech services"));
	arrTitle.Add(_T("F"));

	if (m_bOutPatient)
		arrFeeGrp.Add(_T("Exam Fee"));
	else
		arrFeeGrp.Add(_T("Bed Fee"));
	arrTitle.Add(_T("G"));


	CString szDate, szMoney;
	
	for (int i = 0; i < arrFeeGrp.GetSize(); i++)
	{

		if (m_bInPatient)
		{
			if (i == 6)
				szSQL = GetQueryString(i + 1);
			else if (i == 4)
				szSQL = GetQueryString(i * 2);
			else
				szSQL = GetQueryString(i);
		}
		else
			szSQL = GetQueryString(i);

		rs.ExecSQL(szSQL);
		nCount = rs.GetRecordCount();

		szNewLine.Empty();
		szOldLine.Empty();

		//Phi kham - Phi giuong
		if (i == 6)
		{
			nIndex = 1;
			grpCost = ttlCost = cost = 0;
			
			if (!rs.IsEOF())
			{	
				nRow++;
				TranslateString(arrFeeGrp[i], tmpStr);
				xls.SetCellText(0, nRow, arrTitle[i], FMT_TEXT, TRUE, 11);
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE, 11);

				if (m_bOutPatient)
				{
					while (!rs.IsEOF())
					{			
						nRow++;	
						tmpStr.Format(_T("%d"), nIndex++);
						xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);

						rs.GetValue(_T("examname"), tmpStr);
						xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);

						rs.GetValue(_T("qty"), tmpStr);					
						xls.SetCellText(2, nRow, tmpStr, FMT_INTEGER);

						rs.GetValue(_T("insprice"), cost);
						tmpStr.Format(_T("%.2f"), cost);
						xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);

						rs.GetValue(_T("suminsprice"), cost);
						grpCost += cost;
						tmpStr.Format(_T("%.2f"), cost);
						xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
						
						rs.MoveNext();
					}		
					
					ttlCost += grpCost;

					if (ttlCost > 0)
					{
						nRow++;
						TranslateString(_T("Total Group"), tmpStr);
						xls.SetCellText(1, nRow, tmpStr + _T("(G)"), FMT_TEXT, TRUE);
						tmpStr.Format(_T("%.2lf"), ttlCost);
						xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
						totalCost += ttlCost;
					}
				}
				else
				{
					CString szOldRoom, szNewRoom;
					//CString szOldLine, szNewLine;
					int nQty = 0, nTotalQty = 0;

					while (!rs.IsEOF())
					{
						rs.GetValue(_T("deptname"), szNewLine);
						if (szNewLine != szOldLine && !szNewLine.IsEmpty())
						{
							if (grpCost > 0)
							{	
								nRow++;
								xls.SetCellText(1, nRow, _T("\x43\x1ED9ng kho\x61:"), FMT_TEXT, true, 11);
								tmpStr.Format(_T("%.2Lf"), grpCost);
								xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true, 11);
								ttlCost += grpCost;
								nIndex = 1;
								grpCost = 0;
							}
							nRow++;
							xls.SetCellText(1, nRow, szNewLine, FMT_TEXT, true, 11);
							szOldLine = szNewLine;
							szTemp.Empty();
						}

						rs.GetValue(_T("roomname"), szNewRoom);
						if (szNewRoom != szOldRoom && !szNewRoom.IsEmpty())
						{
							nRow++;
							xls.SetCellText(1, nRow, szNewRoom, FMT_TEXT, true);
							szOldRoom = szNewRoom;
							szTemp.Empty();
						}

						nRow++;
						tmpStr.Format(_T("%d"), nIndex++);
						xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);

						rs.GetValue(_T("bedname"), tmpStr);
						xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);

						rs.GetValue(_T("treatqty"), nQty);
						nTotalQty += nQty;
						tmpStr.Format(_T("%d"), nQty);
						xls.SetCellText(2, nRow, tmpStr, FMT_INTEGER);

						rs.GetValue(_T("insprice"), cost);
						tmpStr.Format(_T("%.2f"), cost);
						xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);

						rs.GetValue(_T("total"), cost);
						//nLineCost += cost;
						grpCost += cost;
						tmpStr.Format(_T("%.2f"), cost);
						xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
						
						rs.MoveNext();
					}

					if (grpCost > 0)
					{
						nRow++;
						//TranslateString(_T("C\x1ED9ng:"), tmpStr);
						xls.SetCellText(1, nRow, _T("\x43\x1ED9ng kho\x61:"), FMT_TEXT | FMT_WRAPING, true, 11);
						tmpStr.Format(_T("%.2Lf"), grpCost);
						xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true, 11);
						ttlCost += grpCost;
					}

					if (ttlCost > 0)
					{
						nRow++;
						TranslateString(_T("Total Group"), tmpStr);
						xls.SetCellText(1, nRow, tmpStr + _T("(G)"), FMT_TEXT, true);
						tmpStr.Format(_T("%.2Lf"), ttlCost);
						xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
						totalCost += ttlCost;
					}
				}
			}
		}

		//Phi xet nghiem
		if (i == 0)
		{
			nIndex = 1;
			grpCost = ttlCost = cost = ttlCost = 0;

			if (!rs.IsEOF())
			{	
				nRow++;
				TranslateString(arrFeeGrp[i], tmpStr);

				xls.SetCellText(0, nRow, arrTitle[i], FMT_TEXT, TRUE, 11);
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE, 11);

				while (!rs.IsEOF())
				{
					rs.GetValue(_T("groupname"), szNewLine);

					if(szNewLine != szOldLine && !szNewLine.IsEmpty())
					{
						CString szField;
						if (grpCost > 0)
						{	
							nRow++;
							TranslateString(_T("C\x1ED9ng:"), tmpStr);
							xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE);
							tmpStr.Format(_T("%.2f"), grpCost);
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
					xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);

					rs.GetValue(_T("testname"), tmpStr);
					xls.SetCellText(1,nRow,tmpStr, FMT_TEXT);

					rs.GetValue(_T("qty"), tmpStr);					
					xls.SetCellText(2,nRow,tmpStr, FMT_INTEGER);

					rs.GetValue(_T("insprice"), cost);		
					tmpStr.Format(_T("%.2f"), cost);
					xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);

					rs.GetValue(_T("suminsprice"), cost);
					grpCost += cost;
					tmpStr.Format(_T("%.2f"), cost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
					
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
					xls.SetCellText(1, nRow, tmpStr + _T("(A)"), FMT_TEXT, TRUE);
					tmpStr.Format(_T("%.2lf"), ttlCost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
					totalCost += ttlCost;
				}
			}
		}

		//Phi chan doan hinh anh
		if (i == 1)
		{
			nIndex = 1;
			grpCost = ttlCost = cost = ttlCost = 0;

			if (!rs.IsEOF())
			{	
				nRow++;
				TranslateString(arrFeeGrp[i], tmpStr);

				xls.SetCellText(0, nRow, arrTitle[i], FMT_TEXT, TRUE, 11);
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE, 11);

				while (!rs.IsEOF())
				{
					rs.GetValue(_T("groupname"), szNewLine);

					if(szNewLine != szOldLine && !szNewLine.IsEmpty())
					{
						CString szField;
						if (grpCost > 0)
						{	
							nRow++;
							TranslateString(_T("C\x1ED9ng:"), tmpStr);
							xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE);
							tmpStr.Format(_T("%.2f"), grpCost);
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
					xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);

					rs.GetValue(_T("testname"), tmpStr);
					xls.SetCellText(1,nRow,tmpStr, FMT_TEXT);

					rs.GetValue(_T("qty"), tmpStr);					
					xls.SetCellText(2,nRow,tmpStr, FMT_INTEGER);

					rs.GetValue(_T("insprice"), cost);		
					tmpStr.Format(_T("%.2f"), cost);
					xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);

					rs.GetValue(_T("suminsprice"), cost);
					grpCost += cost;
					tmpStr.Format(_T("%.2f"), cost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
					
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
					xls.SetCellText(1, nRow, tmpStr + _T("(B)"), FMT_TEXT, TRUE);
					tmpStr.Format(_T("%.2lf"), ttlCost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
					totalCost += ttlCost;
				}
			}
		}


		//Phi tham do chuc nang
		if (i == 2)
		{
			nIndex = 1;
			grpCost = ttlCost = cost = ttlCost = 0;

			if (!rs.IsEOF())
			{	
				nRow++;
				TranslateString(arrFeeGrp[i], tmpStr);

				xls.SetCellText(0, nRow, arrTitle[i], FMT_TEXT, TRUE, 11);
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE, 11);

				while (!rs.IsEOF())
				{
					rs.GetValue(_T("groupname"), szNewLine);

					if(szNewLine != szOldLine && !szNewLine.IsEmpty())
					{
						CString szField;
						if (grpCost > 0)
						{	
							nRow++;
							TranslateString(_T("C\x1ED9ng:"), tmpStr);
							xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE);
							tmpStr.Format(_T("%.2f"), grpCost);
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
					xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);

					rs.GetValue(_T("testname"), tmpStr);
					xls.SetCellText(1,nRow,tmpStr, FMT_TEXT);

					rs.GetValue(_T("qty"), tmpStr);					
					xls.SetCellText(2,nRow,tmpStr, FMT_INTEGER);

					rs.GetValue(_T("insprice"), cost);		
					tmpStr.Format(_T("%.2f"), cost);
					xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);

					rs.GetValue(_T("suminsprice"), cost);
					grpCost += cost;
					tmpStr.Format(_T("%.2f"), cost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
					
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
					xls.SetCellText(1, nRow, tmpStr + _T("(C)"), FMT_TEXT, TRUE);
					tmpStr.Format(_T("%.2lf"), ttlCost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
					totalCost += ttlCost;
				}
			}
		}

		//Phi phau thuat thu thuat
		if (i == 3)
		{
			nIndex = 1;
			grpCost = ttlCost = cost = 0;

			if (!rs.IsEOF())
			{
				nRow++;		
				TranslateString(arrFeeGrp[i], tmpStr);
				xls.SetCellText(0, nRow, arrTitle[i], FMT_TEXT, TRUE, 11);
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE, 11);		
				nIndex = 1;	
				while (!rs.IsEOF())
				{
					rs.GetValue(_T("operatype"), szNewLine);
					if(szNewLine != szOldLine && !szNewLine.IsEmpty())
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
					xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);

					rs.GetValue(_T("itemname"), tmpStr);
					xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);

					rs.GetValue(_T("qty"), tmpStr);
					xls.SetCellText(2, nRow, tmpStr, FMT_INTEGER);

					rs.GetValue(_T("insprice"), cost);
					tmpStr.Format(_T("%.2f"), cost);
					xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);

					rs.GetValue(_T("amount"), cost);
					grpCost += cost;
					tmpStr.Format(_T("%.2f"), cost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
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

				if(ttlCost > 0)
				{
					nRow++;
					TranslateString(_T("Total Group"), tmpStr);
					xls.SetCellText(1, nRow, tmpStr + _T("(D)"), FMT_TEXT, TRUE);
					tmpStr.Format(_T("%.2lf"), ttlCost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
					totalCost += ttlCost;
				}
			}
		}

		//Phi vat tu y te tieu hao
		if (i == 4)
		{
			nIndex = 1;
			grpCost = ttlCost = cost = 0;

			if (!rs.IsEOF())
			{
				nRow++;
				TranslateString(arrFeeGrp[i], tmpStr);
				xls.SetCellText(0, nRow, arrTitle[i], FMT_TEXT, TRUE, 11);
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE, 11);		
				nIndex = 1;	
				while (!rs.IsEOF())
				{
					rs.GetValue(_T("typeid"), szNewLine);
					if(szNewLine != szOldLine && !szNewLine.IsEmpty())
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
					xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);

					rs.GetValue(_T("drugname"), tmpStr);
					xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);

					rs.GetValue(_T("qty"), tmpStr);
					xls.SetCellText(2, nRow, tmpStr, FMT_INTEGER);

					rs.GetValue(_T("price"), cost);
					tmpStr.Format(_T("%.2f"), cost);
					xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);

					rs.GetValue(_T("amount"), cost);
					grpCost += cost;
					tmpStr.Format(_T("%.2f"), cost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
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
					xls.SetCellText(1, nRow, tmpStr + _T("(E)"), FMT_TEXT, TRUE);
					tmpStr.Format(_T("%.2lf"), ttlCost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
					totalCost += ttlCost;
				}
			}
		}

		//Phi dich vu ky thuat cao
		if (i == 5)
		{
			nIndex = 1;
			grpCost = ttlCost = cost = 0;
			
			if (!rs.IsEOF())
			{
				nRow++;		
				TranslateString(arrFeeGrp[i], tmpStr);
				xls.SetCellText(0, nRow, arrTitle[i], FMT_TEXT, TRUE, 11);
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, TRUE, 11);		
				nIndex = 1;

				while (!rs.IsEOF())
				{
					rs.GetValue(_T("groupname"), szNewLine);
					if(szNewLine != szOldLine && !szNewLine.IsEmpty())
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
					xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);

					rs.GetValue(_T("itemname"), tmpStr);
					xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);

					rs.GetValue(_T("qty"), tmpStr);
					xls.SetCellText(2,nRow, tmpStr,FMT_INTEGER);

					rs.GetValue(_T("insprice"), cost);
					tmpStr.Format(_T("%.2f"), cost);
					xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);

					rs.GetValue(_T("amount"), cost);
					grpCost += cost;
					tmpStr.Format(_T("%.2f"), cost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
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
					xls.SetCellText(1, nRow, tmpStr + _T("(F)"), FMT_TEXT, TRUE);
					tmpStr.Format(_T("%.2lf"), ttlCost);
					xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE);
					totalCost += ttlCost;
				}
			}
		}
	}

	if (totalCost > 0)
	{
		nRow++;
		TranslateString(_T("Total Group"), tmpStr);
		xls.SetCellText(1, nRow, tmpStr + _T("(A+B+C+D+E+F+G)"), FMT_TEXT, TRUE, 11);
		tmpStr.Format(_T("%.2lf"), totalCost);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, TRUE, 11);		
	}

	nRow++;
	szMoney.Format(_T("%.0f"), floor(totalCost + 0.5));
	MoneyToString(szMoney, tmpStr);
	xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(_T("Ng\xE0y %s th\xE1ng %s n\x103m %s"),
		          tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	xls.SetCellText(3, nRow + 3, szDate, FMT_TEXT);

	EndWaitCursor();
	xls.Save(_T("Exports\\Thong Ke Dich Vu Ky Thuat Theo Quy.xls"));

} 
int CFMInsuranceReport21a::OnAddFMInsuranceReport21a(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMInsuranceReport21a::OnEditFMInsuranceReport21a(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMInsuranceReport21a::OnDeleteFMInsuranceReport21a(){
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
 		OnCancelFMInsuranceReport21a(); 
 	}
	return 0;
}
int CFMInsuranceReport21a::OnSaveFMInsuranceReport21a(){
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
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret > 0) 
 	{ 
 		//OnFMInsuranceReport21aListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMInsuranceReport21a::OnCancelFMInsuranceReport21a(){
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
int CFMInsuranceReport21a::OnFMInsuranceReport21aListLoadData(){
	return 0;
}

CString CFMInsuranceReport21a::GetQueryString(int nID)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere, szHospitalID, szDept, szStatus, szLineWhere;
	szSQL.Format(_T("SELECT sc_id hid FROM sys_company"));
	rs.ExecSQL(szSQL);
	rs.GetValue(_T("hid"), szHospitalID);
	if (m_bCheckDrug)
		m_szDrugPTTT.Format(_T(" and hpo_ordertype not in('M') "));
	//else
		//m_szDrugPTTT.Format(_T(" and hpo_ordertype in('P','D','M','E') "));

	if (m_bInPatient == 1)
	{
		szWhere.AppendFormat(_T(" and fi.hfe_class not in('E','S') "));
		szStatus.Format(_T(" and hd_isinpatient='Y' "));
	}
	else
	{
		szWhere.AppendFormat(_T(" and fi.hfe_class in('E') "));
		szStatus.Format(_T(" and hd_status='T' "));
	}
	if (!m_szPatientLineKey.IsEmpty())
		if (m_szPatientLineKey == _T("3"))
			szLineWhere.Format(_T(" AND (hd_insline = 'Y' OR hd_object = 12)"));
		else
		{
			szLineWhere.Format(_T(" AND hd_insline = 'N' AND hd_object <> 12"));
			if (m_szPatientLineKey == _T("2"))
				szLineWhere.AppendFormat(_T(" AND substr(hd_cardno, 16, 20) <> '%s'"), szHospitalID);
			else
				szLineWhere.AppendFormat(_T(" AND instr(substr(hd_cardno, 16, 20), '%s') > 0"), szHospitalID);
		}
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (!szDept.IsEmpty())
				szDept += _T(",");
			szDept.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
		}
	}

	if (!szDept.IsEmpty())
		szWhere.AppendFormat(_T(" and fi.hfe_deptid in(%s) "), szDept);
	
	switch (nID)
	{
	case 6:
		szSQL.Format(_T(" select hfl_name as examname,") \
					_T("        hfl_unit as unit,") \
					_T("        sum(qty) as qty,") \
					_T("        insprice,") \
					_T("        sum(qty*insprice) as suminsprice") \
					_T(" from") \
					_T(" (") \
					_T("   select hfe_invoiceno as invoiceno,") \
					_T("          he_docno as docno,") \
					_T("          he_examtype as examtype,") \
					_T("          sum(hfe_quantity) as qty,") \
					_T("          hfe_insprice as insprice") \
					_T("   from hms_exam") \
					_T("   left join hms_fee_view on(hfe_docno=he_docno and hfe_orderid=he_receptidx and hfe_type='E')") \
					_T("   left join hms_doc on(hd_docno=he_docno)") \
					_T("   left join hms_patient on(hp_patientno=hd_patientno)") \
					_T("   left join hms_object on(ho_id=hd_object)") \
					_T("   where hd_status='T' and hfe_discount > 0 and ho_type in('I','C') %s") \
					_T("   group by he_docno, he_examtype, hfe_insprice, hfe_invoiceno") \
					_T(" ) tbl") \
					_T(" left join hms_fee_invoice fi on(hfe_docno=docno and hfe_invoiceno=invoiceno)") \
					_T(" left join hms_fee_list on(hfl_feeid=examtype)") \
					_T(" where hfl_inspaid='Y' and insprice > 0 and fi.hfe_status='P'") \
					_T("       and fi.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
					_T(" group by hfl_name, hfl_unit, insprice") \
					_T(" order by hfl_name"),
					szLineWhere, m_szFromDate, m_szToDate, szWhere);_debug(_T("6"));
		break;
	case 0:
		szSQL.Format(_T(" select hfg_name as groupname,") \
					_T("        groupid,") \
					_T("        testname,") \
					_T("        unit,") \
					_T("        sum(qty) as qty,") \
					_T("        insprice,") \
					_T("        sum(qty*insprice) as suminsprice") \
					_T(" from") \
					_T(" (") \
					_T("   select fe.hfe_invoiceno as invoiceno,") \
					_T("          fe.hfe_docno as docno,") \
					_T("          fe.hfe_group as groupid,") \
					_T("          fe.hfe_desc as testname,") \
					_T("          fe.hfe_unit as unit,") \
					_T("          sum(hfe_quantity) as qty,") \
					_T("          hfe_insprice as insprice") \
					_T("   from hms_fee_view fe") \
					_T("   left join hms_fee_invoice fi on(fi.hfe_docno=fe.hfe_docno and fe.hfe_invoiceno=fi.hfe_invoiceno)") \
					_T("   left join hms_doc on(hd_docno=fe.hfe_docno)") \
					_T("   left join hms_object on(ho_id=hd_object)") \
					_T("   where fe.hfe_type IN ('P', 'T') AND fe.hfe_discount > 0 %s %s") \
					_T("         and substr(hfe_group,1,2)='B1' and ho_type in('I','C')") \
					_T("         and hfe_hitech='N' and fi.hfe_status='P'") \
					_T("	     and fi.hfe_date between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') %s") \
					_T("   group by fe.hfe_docno, fe.hfe_group, fe.hfe_desc, fe.hfe_unit, hfe_insprice, fe.hfe_invoiceno") \
					_T(" ) tbl") \
					_T(" left join hms_fee_group on(hfg_id=groupid)") \
					_T(" where insprice > 0 ") \
					_T(" group by groupid, hfg_name, testname, unit, insprice") \
					_T(" order by groupid"),
					szLineWhere, szStatus, m_szFromDate, m_szToDate, szWhere);_debug(_T("0"));
		break;
	case 1:
		szSQL.Format(_T(" select hfg_name as groupname,") \
					_T("        groupid,") \
					_T("        testname,") \
					_T("        unit,") \
					_T("        sum(qty) as qty,") \
					_T("        insprice,") \
					_T("        sum(qty*insprice) as suminsprice") \
					_T(" from") \
					_T(" (") \
					_T("   select hfe_invoiceno as invoiceno,") \
					_T("          hpc_docno as docno,") \
					_T("          hfl_groupid as groupid,") \
					_T("          hfl_name as testname,") \
					_T("          hfl_unit as unit,") \
					_T("          sum(hfe_quantity) as qty,") \
					_T("          hfe_insprice as insprice") \
					_T("   from hms_pacsorder") \
					_T("   left join hms_pacsorderline on(hpcl_orderid=hpc_orderid and hpcl_docno=hpc_docno)") \
					_T("   left join hms_fee_list on(hfl_groupid=hpc_groupid and hfl_feeid=hpcl_itemid)") \
					_T("   left join hms_fee_view on(hfe_orderid=hpcl_orderid and hfe_orderline=hpcl_orderlineid") \
					_T("                             and hfe_docno=hpcl_docno and hfe_type='P')") \
					_T("   left join hms_doc on(hd_docno=hpc_docno)") \
					_T("   left join hms_clinical_record on(hcr_docno=hd_docno)") \
					_T("   left join hms_object on(ho_id=hd_object)") \
					_T("   where hfe_discount > 0") \
					_T("         and substr(hfl_groupid,1,2)='B2' and ho_type in('I','C')") \
					_T("         and hfe_hitech='N' %s %s") \
					_T("   group by hpc_docno, hfl_groupid, hfl_name, hfl_unit, hfe_insprice, hfe_invoiceno") \
					_T(" ) tbl") \
					_T(" left join hms_fee_invoice fi on(hfe_docno=docno and hfe_invoiceno=invoiceno)") \
					_T(" left join hms_fee_group on(hfg_id=groupid)") \
					_T(" where insprice > 0 and fi.hfe_status='P'") \
					_T("       and fi.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
					_T(" group by groupid, hfg_name, testname, unit, insprice") \
					_T(" order by groupid"),
					szLineWhere, szStatus, m_szFromDate, m_szToDate, szWhere);_debug(_T("1"));
		break;
	case 2:
		szSQL.Format(_T(" select hfg_name as groupname,") \
					_T("        groupid,") \
					_T("        testname,") \
					_T("        unit,") \
					_T("        sum(qty) as qty,") \
					_T("        insprice,") \
					_T("        sum(qty*insprice) as suminsprice") \
					_T(" from") \
					_T(" (") \
					_T("   select hfe_invoiceno as invoiceno,") \
					_T("          hpc_docno as docno,") \
					_T("          hfl_groupid as groupid,") \
					_T("          hfl_name as testname,") \
					_T("          hfl_unit as unit,") \
					_T("          sum(hfe_quantity) as qty,") \
					_T("          hfe_insprice as insprice") \
					_T("   from hms_pacsorder") \
					_T("   left join hms_pacsorderline on(hpcl_orderid=hpc_orderid and hpcl_docno=hpc_docno)") \
					_T("   left join hms_fee_list on(hfl_groupid=hpc_groupid and hfl_feeid=hpcl_itemid)") \
					_T("   left join hms_fee_view on(hfe_orderid=hpcl_orderid and hfe_orderline=hpcl_orderlineid") \
					_T("                             and hfe_docno=hpcl_docno and hfe_type='P')") \
					_T("   left join hms_doc on(hd_docno=hpc_docno)") \
					_T("   left join hms_clinical_record on(hcr_docno=hd_docno)") \
					_T("   left join hms_object on(ho_id=hd_object)") \
					_T("   where hfe_discount > 0") \
					_T("         and substr(hfl_groupid,1,2)='B3' and ho_type in('I','C')") \
					_T("         and hfe_hitech='N' %s %s") \
					_T("   group by hpc_docno, hfl_groupid, hfl_name, hfl_unit, hfe_insprice, hfe_invoiceno") \
					_T(" ) tbl") \
					_T(" left join hms_fee_invoice fi on(hfe_docno=docno and hfe_invoiceno=invoiceno)") \
					_T(" left join hms_fee_group on(hfg_id=groupid)") \
					_T(" where insprice > 0 and fi.hfe_status='P'") \
					_T("       and fi.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
					_T(" group by groupid, hfg_name, testname, unit, insprice") \
					_T(" order by groupid"),
					szLineWhere, szStatus, m_szFromDate, m_szToDate, szWhere);_debug(_T("2"));
		break;
	case 3:
		szSQL.Format(_T(" select hfg_name as operatype,") \
					_T("        hfl_name as itemname,") \
					_T("        sum(fe.hfe_quantity) as qty,") \
					_T("        fe.hfe_insprice as insprice,") \
					_T("        sum(fe.hfe_quantity*fe.hfe_insprice) as amount") \
					_T(" from hms_patient") \
					_T(" left join hms_doc on(hd_patientno=hp_patientno)") \
					_T(" left join hms_clinical_record on(hcr_docno=hd_docno)") \
					_T(" left join hms_object ob on(ob.ho_id=hd_object)") \
					_T(" left join hms_operation op on(op.ho_docno=hd_docno)") \
					_T(" left join hms_fee_list on(hfl_feeid=ho_itemid)") \
					_T(" left join hms_fee_group on(hfg_id=hfl_groupid)") \
					_T(" left join hms_fee_view fe on(fe.hfe_docno=ho_docno and fe.hfe_orderid=ho_idx") \
					_T("                              and fe.hfe_type='O')") \
					_T(" left join hms_fee_invoice fi on(fi.hfe_docno=fe.hfe_docno and fi.hfe_invoiceno=fe.hfe_invoiceno)") \
					_T(" where fi.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s %s %s") \
					_T("       and fi.hfe_status='P' and hfe_hitech='N' and ob.ho_type in('I','C') and fe.hfe_discount > 0") \
					_T(" group by hfg_name, hfl_name, fe.hfe_insprice") \
					_T(" order by hfg_name, hfl_name"),
					m_szFromDate, m_szToDate, szStatus, szWhere, szLineWhere);_debug(_T("3"));
		break;
	case 4:
		szSQL.Format(_T(" select product_typename as typeid,") \
					_T("        product_code as code,") \
					_T("        product_name as drugname,") \
					_T("        product_uomname as unit,") \
					_T("        product_countryname as countryname,") \
					_T("        fe.hfe_insprice as price,") \
					_T("        sum(fe.hfe_quantity) as qty,") \
					_T("        sum(fe.hfe_quantity*fe.hfe_insprice) as amount") \
					_T(" from hms_pharmaorder") \
					_T(" left join hms_pharmaorderline on(hpol_orderid=hpo_orderid)") \
					_T(" left join hms_doc on(hd_docno=hpo_docno)") \
					_T(" left join m_productitem_view on(hpol_product_item_id=product_item_id)") \
					_T(" left join hms_object on(hd_object=ho_id)") \
					_T(" left join hms_fee_view fe on(hpol_orderid=fe.hfe_orderid and hpol_orderlineid=fe.hfe_orderline") \
					_T("                              and fe.hfe_docno=hpol_docno)") \
					_T(" left join hms_fee_invoice fi on(fi.hfe_docno=fe.hfe_docno and fi.hfe_invoiceno=fe.hfe_invoiceno)") \
					_T(" where ho_type in('I','C') and fe.hfe_discount > 0 %s %s") \
					_T("       and product_producttype='A9000' and fi.hfe_status='P'") \
					_T("       and fi.hfe_date between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') %s %s") \
					_T(" group by product_typename, product_code, product_name, product_uomname, product_countryname, fe.hfe_insprice") \
					_T(" order by product_typename"),
					szLineWhere, szStatus, m_szFromDate, m_szToDate, szWhere, m_szDrugPTTT);_debug(_T("4"));
		break;
	case 5:
		szSQL.Format(_T(" select hfg_name as groupname,") \
					_T("        hfl_groupid as groupid,") \
					_T("        hfl_name as itemname,") \
					_T("        hfl_unit as unit,") \
					_T("        sum(qty) as qty,") \
					_T("        insprice,") \
					_T("        sum(qty*insprice) as amount") \
					_T(" from") \
					_T(" (") \
					_T("   select fe.hfe_invoiceno as invoiceno,") \
					_T("          hpc_docno as docno,") \
					_T("          hpcl_itemid as itemid,") \
					_T("          hfe_quantity as qty,") \
					_T("          hfe_insprice as insprice") \
					_T("   from hms_testorder") \
					_T("   left join hms_testorderline on(hpc_orderid=hpcl_orderid)") \
					_T("   left join hms_fee_view fe on(hfe_orderid=hpcl_orderid and hfe_orderline=hpcl_orderlineid") \
					_T("                             and fe.hfe_docno=hpcl_docno and fe.hfe_type='T')") \
					_T("   left join hms_fee_invoice fi on(fi.hfe_docno=fe.hfe_docno and fi.hfe_invoiceno=fe.hfe_invoiceno)") \
					_T("   where 1=1 and fe.hfe_discount > 0 and hfe_hitech='Y' and fi.hfe_status='P' %s") \
					_T("	and fi.hfe_date between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')") \
					_T("   union all") \
					_T("   select fe.hfe_invoiceno,") \
					_T("          hpc_docno as docno,") \
					_T("          hpcl_itemid as itemid,") \
					_T("          hfe_quantity as qty,") \
					_T("          hfe_insprice as insprice") \
					_T("   from hms_pacsorder") \
					_T("   left join hms_pacsorderline on(hpc_orderid=hpcl_orderid)") \
					_T("   left join hms_fee_view fe on(hfe_orderid=hpcl_orderid and hfe_orderline=hpcl_orderlineid") \
					_T("                             and fe.hfe_docno=hpcl_docno and fe.hfe_type='P')") \
					_T("   left join hms_fee_invoice fi on(fi.hfe_docno=fe.hfe_docno and fi.hfe_invoiceno=fe.hfe_invoiceno)") \
					_T("   where 1=1 and fe.hfe_discount > 0 and hfe_hitech='Y' and fi.hfe_status='P' %s") \
					_T("	and fi.hfe_date between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')") \
					_T("   union all") \
					_T("   select fe.hfe_invoiceno,") \
					_T("          ho_docno as docno,") \
					_T("          ho_itemid as itemid,") \
					_T("          ho_qty as qty,") \
					_T("          hfe_insprice as insprice") \
					_T("   from hms_operation") \
					_T("   left join hms_fee_view fe on(hfe_orderid=ho_idx and fe.hfe_type='O')") \
					_T("   left join hms_fee_invoice fi on(fi.hfe_docno=fe.hfe_docno and fi.hfe_invoiceno=fe.hfe_invoiceno)") \
					_T("   where 1=1 and fe.hfe_discount > 0 and hfe_hitech='Y' and fi.hfe_status='P' %s") \
					_T("	and fi.hfe_date between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')") \
					_T(" ) tbl") \
					_T(" left join hms_doc on(hd_docno=docno)") \
					_T(" left join hms_clinical_record on(hcr_docno=hd_docno)") \
					_T(" left join hms_object on(ho_id=hd_object)") \
					_T(" left join hms_fee_list on(hfl_feeid=itemid)") \
					_T(" left join hms_fee_group on(hfg_id=hfl_groupid)") \
					_T(" where ho_type in('I','C') %s") \
					_T("       and insprice > 0  %s") \
					_T(" group by hfl_groupid, hfg_name, hfl_name, hfl_unit, insprice") \
					_T(" order by hfl_groupid"),
					szWhere, m_szFromDate, m_szToDate, szWhere, m_szFromDate, m_szToDate, szWhere, m_szFromDate, m_szToDate, szStatus, szLineWhere);_debug(_T("5"));
	
		break;
	case 7:
		szSQL.Format(_T(" select hfg_name as groupName,") \
					_T("        sd_name as deptname,") \
					_T("        hrl_name as roomname,") \
					_T("        bedname,") \
					_T("        insprice,") \
					_T("        sum(treatqty) as treatqty,") \
					_T("        sum(treatqty*insprice) as total") \
					_T(" from") \
					_T(" (") \
					_T("  select hb_docno as docno,") \
					_T(" 	      fe.hfe_invoiceno as invoiceno,") \
					_T(" 	      hb_roomid as roomid,") \
					_T(" 	      hb_deptid as deptid,") \
					_T(" 	      hbl_name as bedname,") \
					_T(" 	      fe.hfe_group as groupid,") \
					_T(" 	      fe.hfe_insprice as insprice,") \
					_T(" 	      hb_treatqty as treatqty") \
					_T("  from hms_bed be") \
					_T("  left join hms_bedlist on(hbl_roomid=hb_roomid and hbl_deptid=hb_deptid and hbl_id=hb_bedid)") \
					_T("  left join hms_fee_view fe on(hfe_docno=hb_docno and hfe_orderid=hb_idx") \
					_T("                               and hfe_type='B')") \
					_T("  left join hms_doc on(hd_docno=hb_docno)") \
					_T("  left join hms_clinical_record on(hcr_docno=hd_docno)") \
					_T("  where fe.hfe_discount > 0") \
					_T("        and hcr_status='T'") \
					_T("        and (hb_dynprice<>'Y' or hb_dynprice is null) %s") \
					_T("  union all") \
					_T("  select hbi_docno,") \
					_T(" 	      fe.hfe_invoiceno,") \
					_T(" 	      hbi_roomid as roomid,") \
					_T(" 	      hbi_deptid as deptid,") \
					_T(" 	      hdbl_name as bedname,") \
					_T(" 	      fe.hfe_group,") \
					_T(" 	      fe.hfe_insprice as insprice,") \
					_T(" 	      hbi_qty as treatqty") \
					_T("  from hms_bed_items be") \
					_T("  left join hms_dynbedlist on(hdbl_deptid=hbi_deptid and hdbl_idx=hbi_priceid)") \
					_T("  left join hms_fee_view fe on(hfe_docno=hbi_docno and hfe_orderid=hbi_idx") \
					_T("                               and hfe_type='B')") \
					_T("  left join hms_doc on(hd_docno=hbi_docno)") \
					_T("  left join hms_clinical_record on(hcr_docno=hd_docno)") \
					_T("  where fe.hfe_discount > 0") \
					_T("        and hcr_status='T' %s") \
					_T(" ) tbl") \
					_T(" left join hms_fee_invoice fi on(hfe_docno=docno and hfe_invoiceno=invoiceno)") \
					_T(" left join hms_fee_group on(hfg_id=groupid)") \
					_T(" left join sys_dept on(sd_id=deptid)") \
					_T(" left join hms_roomlist on(hrl_deptid=deptid and hrl_id=roomid)") \
					_T(" where insprice > 0 and fi.hfe_status='P'") \
					_T("       and hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
					_T(" group by hfg_name, sd_name, hrl_name, bedname, insprice") \
					_T(" order by deptname, roomname, bedname"),
					szLineWhere, szLineWhere, m_szFromDate, m_szToDate, szWhere);_debug(_T("7"));
		break;
	case 8:
		szSQL.Format(_T(" select hfg_name as typeid,") \
					_T("        mp_code as code,") \
					_T("        mp_name as drugname,") \
					_T("        adu_name as unit,") \
					_T("        adc_name as countryname,") \
					_T("        fe.hfe_insprice as price,") \
					_T("        sum(fe.hfe_quantity) as qty,") \
					_T("        sum(fe.hfe_quantity*fe.hfe_insprice) as amount") \
					_T(" from hms_ipharmaorder") \
					_T(" left join hms_ipharmaorderline on(hpol_orderid=hpo_orderid)") \
					_T(" left join hms_doc on(hd_docno=hpo_docno)") \
					_T(" left join hms_clinical_record on(hcr_docno=hd_docno)") \
					_T(" left join m_product_item on(hpol_product_item_id=mpi_product_item_id)") \
					_T(" left join m_product on(mp_product_id=mpi_product_id)") \
					_T(" left join hms_object on(hd_object=ho_id)") \
					_T(" left join hms_fee_view fe on(hpol_orderid=fe.hfe_orderid and hpol_orderlineid=fe.hfe_orderline") \
					_T("                              and fe.hfe_docno=hpol_docno)") \
					_T(" left join hms_fee_invoice fi on(fi.hfe_docno=fe.hfe_docno and fi.hfe_invoiceno=fe.hfe_invoiceno)") \
					_T(" left join hms_fee_group on(hfg_id=mp_producttype)") \
					_T(" left join ad_uom on(adu_uom_id=mp_uom_id)") \
					_T(" left join ad_country on(mp_country_id=adc_country_id)") \
					_T(" where ho_type in('I','C') and fe.hfe_discount > 0 %s %s") \
					_T("       and mp_producttype='A9000' and fi.hfe_status='P'") \
					_T("       and fi.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s %s") \
					_T(" group by hfg_name, mp_code, mp_name, adu_name, adc_name, fe.hfe_insprice") \
					_T(" order by hfg_name"),
					szLineWhere, szStatus, m_szFromDate, m_szToDate, szWhere, m_szDrugPTTT);_debug(_T("8"));
		break;
	}
	return szSQL;
}

int CFMInsuranceReport21a::OnListCheckAll()
{
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (!m_wndList.GetCheck(i))
		{
			m_wndList.SetCheck(i, TRUE);
		}
	}
	return 0;
}

int CFMInsuranceReport21a::OnListUnCheckAll()
{
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			m_wndList.SetCheck(i, FALSE);
		}
	}
	return 0;
}