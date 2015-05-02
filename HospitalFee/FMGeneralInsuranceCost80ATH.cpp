#include "stdafx.h"
#include "FMGeneralInsuranceCost80ATH.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost80ATH *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMGeneralInsuranceCost80ATH* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost80ATH *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost80ATH *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost80ATH *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMGeneralInsuranceCost80ATH* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost80ATH *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMGeneralInsuranceCost80ATH *pVw = (CFMGeneralInsuranceCost80ATH *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMGeneralInsuranceCost80ATH *pVw = (CFMGeneralInsuranceCost80ATH *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost80ATH*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost80ATH*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMGeneralInsuranceCost80ATH*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost80ATH*)pWnd)->OnDeptListDeleteItem();
} 
static int _OnAddFMGeneralInsuranceCost80ATHFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost80ATH*)pWnd)->OnAddFMGeneralInsuranceCost80ATH();
} 
static int _OnEditFMGeneralInsuranceCost80ATHFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost80ATH*)pWnd)->OnEditFMGeneralInsuranceCost80ATH();
} 
static int _OnDeleteFMGeneralInsuranceCost80ATHFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost80ATH*)pWnd)->OnDeleteFMGeneralInsuranceCost80ATH();
} 
static int _OnSaveFMGeneralInsuranceCost80ATHFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost80ATH*)pWnd)->OnSaveFMGeneralInsuranceCost80ATH();
} 
static int _OnCancelFMGeneralInsuranceCost80ATHFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost80ATH*)pWnd)->OnCancelFMGeneralInsuranceCost80ATH();
}

static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost80ATH*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost80ATH*)pWnd)->OnDeptListUnCheckAll();
}

CFMGeneralInsuranceCost80ATH::CFMGeneralInsuranceCost80ATH(CWnd* pParent)
{
	m_nDlgWidth = 450;
	m_nDlgHeight = 525;
	SetDefaultValues();
}
CFMGeneralInsuranceCost80ATH::~CFMGeneralInsuranceCost80ATH(){
}
void CFMGeneralInsuranceCost80ATH::OnCreateComponents()
{
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 440, 485);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 120, 435, 480);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 310, 55);
	m_wndReportPeriod.Create(this,315, 30, 435, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 310, 85);
	m_wndToDate.Create(this,315, 60, 435, 85); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 90, 115);
	m_wndClerk.Create(this,95, 90, 435, 115); 
	m_wndOutpatientTreatment.Create(this, _T("Outpatient Treatment"), 5, 490, 215, 515);
	m_wndPrint.Create(this, _T("&Print"), 235, 490, 335, 515);
	m_wndExport.Create(this, _T("&Export"), 340, 490, 440, 515);
	m_wndDeptList.Create(this,15, 145, 430, 475); 
}
void CFMGeneralInsuranceCost80ATH::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(75);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 240);


	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.SetCheckBox(TRUE);

}
void CFMGeneralInsuranceCost80ATH::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndDeptList.SetEvent(WE_SELCHANGE, _OnDeptListSelectChangeFnc);
	m_wndDeptList.SetEvent(WE_LOADDATA, _OnDeptListLoadDataFnc);
	m_wndDeptList.SetEvent(WE_DBLCLICK, _OnDeptListDblClickFnc);

	m_wndDeptList.AddEvent(1, _T("Check All"), _OnDeptListCheckAllFnc);
	m_wndDeptList.AddEvent(2, _T("UnCheck All"), _OnDeptListUnCheckAllFnc);
	
	CString szSQL;
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);
	OnDeptListLoadData();

}
void CFMGeneralInsuranceCost80ATH::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndOutpatientTreatment.GetDlgCtrlID(), m_bOutpatientTreatment);

}
void CFMGeneralInsuranceCost80ATH::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMGeneralInsuranceCost80ATH::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMGeneralInsuranceCost80ATH::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bOutpatientTreatment = FALSE;

}
int CFMGeneralInsuranceCost80ATH::SetMode(int nMode){
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
/*void CFMGeneralInsuranceCost80ATH::OnYearChange(){
	
} */
/*void CFMGeneralInsuranceCost80ATH::OnYearSetfocus(){
	
} */
/*void CFMGeneralInsuranceCost80ATH::OnYearKillfocus(){
	
} */
int CFMGeneralInsuranceCost80ATH::OnYearCheckValue()
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
void CFMGeneralInsuranceCost80ATH::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMGeneralInsuranceCost80ATH::OnReportPeriodSelendok()
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
/*void CFMGeneralInsuranceCost80ATH::OnReportPeriodSetfocus(){
	
}*/
/*void CFMGeneralInsuranceCost80ATH::OnReportPeriodKillfocus(){
	
}*/
long CFMGeneralInsuranceCost80ATH::OnReportPeriodLoadData()
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
/*void CFMGeneralInsuranceCost80ATH::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMGeneralInsuranceCost80ATH::OnFromDateChange(){
	
} */
/*void CFMGeneralInsuranceCost80ATH::OnFromDateSetfocus(){
	
} */
/*void CFMGeneralInsuranceCost80ATH::OnFromDateKillfocus(){
	
} */
int CFMGeneralInsuranceCost80ATH::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMGeneralInsuranceCost80ATH::OnToDateChange(){
	
} */
/*void CFMGeneralInsuranceCost80ATH::OnToDateSetfocus(){
	
} */
/*void CFMGeneralInsuranceCost80ATH::OnToDateKillfocus(){
	
} */
int CFMGeneralInsuranceCost80ATH::OnToDateCheckValue(){
	return 0;
} 
void CFMGeneralInsuranceCost80ATH::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMGeneralInsuranceCost80ATH::OnClerkSelendok(){
	 
}
/*void CFMGeneralInsuranceCost80ATH::OnClerkSetfocus(){
	
}*/
/*void CFMGeneralInsuranceCost80ATH::OnClerkKillfocus(){
	
}*/
long CFMGeneralInsuranceCost80ATH::OnClerkLoadData()
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
	szSQL.Format(_T("SELECT su_userid as id, su_name as name FROM sys_user WHERE su_groupid in('A','F') %s ORDER BY su_userid"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){
		m_wndClerk.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMGeneralInsuranceCost80ATH::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMGeneralInsuranceCost80ATH::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;

	if (!rpt.Init(_T("Reports/HMS/HF_BANGKETHANHTOAN_MAU80ATH.RPT")))
		return;

	TCHAR *lszLine[] ={_T("\x110\xFAng tuy\x1EBFn"), _T("Tr\xE1i tuy\x1EBFn")};
	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_ICONERROR | MB_OK);
		return;
	}
	
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);

	rpt.GetReportHeader()->SetValue(_T("ObjectGroup"), _T(""));
	rpt.GetReportHeader()->SetValue(_T("ReportPeriod"), m_szClerkKey);
	rs.GetValue(_T("invoiceno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	//Page Header
	//Report Detail
	int nIndex = 1;
	CString szOldLine, szNewLine, szOldGroup, szNewGroup, szLineName;
	CString szNewOfLine, szOldOfLine,szSumOfline;
	CReportSection* rptDetail;
	long double grpCost[23], grpLine[23], ttlCost[23], grpOfLine[23];
	double cost = 0;
	for (int i =0; i < 23; i++)
	{
		grpCost[i] = 0;
		grpLine[i] = 0;
		ttlCost[i] = 0;
		grpOfLine[i] =0;
	}
	while(!rs.IsEOF())
	{		
		rs.GetValue(_T("linename"), szNewGroup);
		if(szNewGroup != szOldGroup && !szNewGroup.IsEmpty())
		{
			if (grpCost[15] > 0)
			{	
				CString szField;
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(9);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				tmpStr.Format(_T("\x43\x1ED9ng(%s=(I+II)):"), szOldGroup);
				rptDetail->SetValue(_T("TotalGroup"), tmpStr);
				for (int i = 3; i < 18; i++)
				{
					szField.Format(_T("S%d"), i);
					FormatCurrency(grpCost[i], tmpStr);
					rptDetail->SetValue(szField, tmpStr);
					ttlCost[i] += grpCost[i];
					grpCost[i] = 0;
				}
			}
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetFaceSize(9);
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetItalic(true);
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetBold(true);			
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rs.GetValue(_T("insline"), tmpStr);
			rptDetail->SetValue(_T("InsuranceLine"),szNewGroup +_T(". ") + pMF->GetSelectionString(_T("hms_insline"), tmpStr));
			szOldGroup = szNewGroup;			
			nIndex=1;
		}

		rs.GetValue(_T("hdline"), szNewOfLine);
		if (szNewOfLine == _T("0") )
		{
			tmpStr.Format(_T("%s"), lszLine[0]);
			szSumOfline = _T("I");
		}
		else
		{
			tmpStr.Format(_T("%s"), lszLine[1]);
			szSumOfline = _T("II");
		}
			
	
		rptDetail = rpt.AddDetail();		
		rptDetail->SetValue(_T("1"), szSumOfline);
		rptDetail->SetValue(_T("2"), tmpStr);
		
		rs.GetValue(_T("total"),cost);
		grpCost[3] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("testfee"), cost);
		grpCost[4] += cost;
		grpLine[4] += cost;		
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("pacsfee"), cost);
		grpCost[5] += cost;
		grpLine[5] += cost;
		grpOfLine[5] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("drugfee"), cost);
		grpCost[6] += cost;
		grpLine[6] += cost;
		grpOfLine[6] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("bloodfee"), cost);
		grpCost[7] += cost;
		grpLine[7] += cost;
		grpOfLine[7] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);				

		rs.GetValue(_T("normtechfee"), cost);
		grpCost[8] += cost;
		grpLine[8] += cost;
		grpOfLine[8] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		
		rs.GetValue(_T("materialfee"), cost);
		grpCost[9] += cost;
		grpLine[9] += cost;
		grpOfLine[9] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("replace_materialfee"), cost);
		grpCost[10] += cost;
		grpLine[10] += cost;
		grpOfLine[10] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.GetValue(_T("hitechfee"), cost);
		grpCost[11] += cost;
		grpLine[11] += cost;
		grpOfLine[11] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);		

		rs.GetValue(_T("drugfeek"), cost);
		grpCost[12] += cost;
		grpLine[12] += cost;
		grpOfLine[12] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);
		
		rs.GetValue(_T("bedfee"), cost);
		grpCost[13] += cost;
		grpLine[13] += cost;
		grpOfLine[13] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);

		rs.GetValue(_T("otherfee"), cost);;
		grpCost[14] += cost;
		grpLine[14] += cost;
		grpOfLine[14] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("14"), tmpStr);		
		
		rs.GetValue(_T("cost"), cost);
		grpCost[15] += cost;
		grpLine[15] += cost;
		grpOfLine[15] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("15"), tmpStr);

		rs.GetValue(_T("patpaid"), cost);
		grpCost[16] += cost;
		grpLine[16] += cost;
		grpOfLine[16] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("16"), tmpStr);

		rs.GetValue(_T("inspaid"), cost);
		grpCost[17] += cost;
		grpLine[17] += cost;
		grpOfLine[17] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("17"), tmpStr);
		
		rs.MoveNext();
	}

	for (int i = 3; i < 18; i++)
	{
		ttlCost[i] += grpCost[i];
	}
	
	
	if (grpCost[15] > 0)
	{
		CString szField, szAmount;
		TranslateString(_T("Total Line"), szAmount);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(9);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		tmpStr.Format(_T("\x43\x1ED9ng(%s=(I+II)):"), szOldGroup);		
		rptDetail->SetValue(_T("TotalGroup"), tmpStr);
		for (int i =3; i < 18; i++)
		{
			szField.Format(_T("S%d"), i);
			FormatCurrency(grpCost[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);
		}
	
	}
	if (ttlCost[15] > 0)
	{
		CString szField, szAmount;
		TranslateString(_T("Total Amount:"), szAmount);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(9);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(false);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		for (int i =3; i < 18; i++)
		{
			
			szField.Format(_T("S%d"), i);
			FormatCurrency(ttlCost[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);		
		}

	}
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
	EndWaitCursor();
} 
void CFMGeneralInsuranceCost80ATH::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString tmpStr, szSQL;
	TCHAR *lszLine[] ={_T("\x110\xFAng tuy\x1EBFn"), _T("Tr\xE1i tuy\x1EBFn")};

	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);	

	CellFormat df(&xls), cf(&xls), hf(&xls);
	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	df.SetFontSize(11);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING);
	hf.SetFontSize(11);
	cf.SetFontName(_T("Segoe UI"));
	cf.SetFontSize(11);

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 23);
	xls.SetColumnWidth(2, 9);
	xls.SetColumnWidth(3, 12);
	xls.SetColumnWidth(4, 12);
	xls.SetColumnWidth(5, 12);
	xls.SetColumnWidth(6, 12);
	xls.SetColumnWidth(7, 12);
	xls.SetColumnWidth(8, 12);
	xls.SetColumnWidth(9, 12);
	xls.SetColumnWidth(10, 12);
	xls.SetColumnWidth(11, 12);
	xls.SetColumnWidth(12, 12);
	xls.SetColumnWidth(13, 12);
	xls.SetColumnWidth(14, 12);
	xls.SetColumnWidth(15, 12);
	xls.SetColumnWidth(16, 12);
	xls.SetColumnWidth(17, 12);
	
	xls.SetCellMergedColumns(0, 0, 4);
	xls.SetCellMergedColumns(0, 1, 4);

	xls.SetCellMergedColumns(11, 0, 6);
	xls.SetCellMergedColumns(11, 1, 6);

	xls.SetCellText(0, 0, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(11, 0, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(11, 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedColumns(0, 2, 17);
	xls.SetCellMergedColumns(0, 3, 17);

	xls.SetCellText(0, 2, _T("\x44\x41NH S\xC1\x43H \x110\x1EC0 NGH\x1ECA TH\x41NH TO\xC1N \x43HI PH\xCD KH\xC1M \x43H\x1EEE\x41 \x42\x1EC6NH N\x1ED8ITR\xDA"),
		            FMT_TEXT | FMT_CENTER, true, 16);
	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(0, 3, tmpStr, &df);

	int nRow = 4;

	xls.SetCellText(0, nRow, _T("STT"), &hf);

	TranslateString(_T("KCB Trong k\x1EF3"), tmpStr);
	xls.SetCellText(1, nRow, tmpStr, &hf);

	xls.SetCellText(2, nRow, _T("S\x1ED1 l\x1B0\x1EE3ng"), &hf);
	
	xls.SetCellMergedRows(0, nRow, 2);
	xls.SetCellMergedRows(1, nRow, 2);
	xls.SetCellMergedRows(2, nRow, 2);
		
	//xls.SetCellMergedRows(12, nRow, 2);
	xls.SetCellMergedRows(14, nRow, 2);
	xls.SetCellMergedRows(15, nRow, 2);
	xls.SetCellMergedRows(16, nRow, 2);
	xls.SetCellMergedRows(17, nRow, 2);
	
	xls.SetCellMergedColumns(3, nRow, 11);

	nRow = 5;

	TranslateString(_T("Medical expenses incurred at medical care facilities"), tmpStr);
	xls.SetCellText(3, nRow - 1, tmpStr, &hf);
	xls.SetCellText(3, nRow, _T("XN,CDCN"), &hf);	
	xls.SetCellText(4, nRow, _T("CDHA"), &hf);

	TranslateString(_T("Drug"), tmpStr);
	xls.SetCellText(5, nRow, tmpStr, &hf);

	TranslateString(_T("Blood"), tmpStr);
	xls.SetCellText(6, nRow, tmpStr, &hf);

	TranslateString(_T("TT, PT"), tmpStr);
	xls.SetCellText(7, nRow, tmpStr, &hf);

	xls.SetCellText(8, nRow, _T("VTYT"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(9, nRow, _T("VTYTTT"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(10, nRow, _T("DVKTC"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(11, nRow, _T("Thu\x1ED1\x63 K"), &hf);

	TranslateString(_T("Examination Fee"), tmpStr);
	xls.SetCellText(12, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(13, nRow, _T("Ph\xED V\x43"), &hf);

	TranslateString(_T("Total Cost"), tmpStr);

	xls.SetCellText(14, nRow - 1, tmpStr, &hf);

	TranslateString(_T("Difference Payment"), tmpStr);
	xls.SetCellText(15, nRow - 1, tmpStr, &hf);

	TranslateString(_T("Insurance Paid"), tmpStr);
	xls.SetCellText(16, nRow - 1, tmpStr, &hf);

	TranslateString(_T("Ngo\xE0i qu\x1EF9"), tmpStr);
	xls.SetCellText(17, nRow - 1, tmpStr, &hf);

	szSQL = GetQueryString();

	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return ;
	}
	long nIndex = 1;
	CString szOldLine, szNewLine, szOldGroup, szNewGroup, szLineName;
	CString szNewOfLine, szOldOfLine,szSumOfline;

	long double grpCost[17], grpLine[17], ttlCost[17], grpOfLine[17];
	double cost = 0;
	nRow++;

	for (int i =0; i < 17; i++)
	{
		tmpStr.Format(_T("(%d)"), i + 1);
		xls.SetCellText(i, nRow, tmpStr, FMT_TEXT | FMT_CENTER, TRUE);
		grpCost[i] = 0;
		grpOfLine[i] = 0;
		grpLine[i] = 0;
		ttlCost[i] = 0;	
	}	

	while (!rs.IsEOF())
	{		
		rs.GetValue(_T("linename"), szNewLine);
		if (szNewLine != szOldLine && !szNewLine.IsEmpty())
		{
			CString szField,szAmount;
			//Tong tuyen (A,B,C)
			if (grpLine[14] > 0)
			{
				nRow++;
				xls.SetCellMergedColumns(1, nRow, 2);
				tmpStr.Format(_T("\x43\x1ED9ng(%s %s=(I+II)): "), szOldLine, szOldGroup);				
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true);

				for (int i = 2; i < 17; i++)
				{					
					tmpStr.Format(_T("%.2lf"), grpLine[i]);
					xls.SetCellText(i, nRow, tmpStr, FMT_NUMBER1, true);
					ttlCost[i] += grpLine[i];
					grpLine[i] = 0;
				}
			}

			nRow++;
			xls.SetCellText(0, nRow , szNewLine, FMT_TEXT, true);
			rs.GetValue(_T("insline"), szLineName);
			tmpStr.Format(_T("%s"), pMF->GetSelectionString(_T("hms_insline"), szLineName));
			xls.SetCellMergedColumns(1, nRow, 17);
			xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true);
			szOldLine = szNewLine;
			nIndex = 1;
		}
		
		nRow++;

		rs.GetValue(_T("hdline"), szNewGroup);
		if (szNewGroup == _T("0"))
		{
			tmpStr.Format(_T("%s"), lszLine[0]);
			szSumOfline = _T("I");
		}
		else
		{
			tmpStr.Format(_T("%s"), lszLine[1]);
			szSumOfline = _T("II");
		}
		
		xls.SetCellText(0, nRow, szSumOfline, FMT_TEXT);

		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("total"), tmpStr);
		xls.SetCellText(2, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("testfee"), cost);		
		grpLine[3] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("pacsfee"), cost);		
		grpLine[4] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("drugfee"), cost);		
		grpLine[5] += cost;			
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("bloodfee"), cost);		
		grpLine[6] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("normtechfee"), cost);
		grpLine[7] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("materialfee"), cost);
		grpLine[8] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("replace_materialfee"), cost);
		grpLine[9] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(9, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("hitechfee"), cost);		
		grpLine[10] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(10, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("drugfeek"), cost);
		grpLine[11] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("bedfee"), cost);	
		grpLine[12] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(12, nRow, tmpStr, FMT_NUMBER1);
		

		/*rs.GetValue(_T("otherfee"), cost);		
		grpLine[12] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(12, nRow, tmpStr, FMT_NUMBER1);	*/

		rs.GetValue(_T("cost"), cost);
		grpLine[14] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(14, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("patpaid"), cost);
		grpLine[15] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(15, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("inspaid"), cost);
		grpLine[16] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(16, nRow, tmpStr, FMT_NUMBER1);
		rs.MoveNext();
	}

	for (int i = 0; i < 17; i++)
	{
		ttlCost[i] += grpLine[i];
	}
	
	if (grpLine[14] > 0)
	{
		nRow++;
		xls.SetCellMergedColumns(1, nRow, 2);
		tmpStr.Format(_T("\x43\x1ED9ng(%s=(I+II)): "), szNewLine);	
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true);
		for (int i = 2; i < 17; i++)
		{			
			tmpStr.Format(_T("%.2lf"), grpLine[i]);
			xls.SetCellText(i, nRow , tmpStr, FMT_NUMBER1, true);	
			grpLine[i] = 0;
		}
	}

	if (ttlCost[14] > 0)
	{
		nRow++;
		CString szField, szAmount;
		TranslateString(_T("Total Amount"), szAmount);
		tmpStr.Format(_T("%s (A+B+C):"), szAmount);
		xls.SetCellMergedColumns(1, nRow, 2);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true);
		for (int i = 2; i < 17; i++)
		{			
			tmpStr.Format(_T("%.2lf"), ttlCost[i]);
			xls.SetCellText(i, nRow, tmpStr, FMT_NUMBER1, true);
		}
	}	
	xls.Save(_T("Exports\\ChiPhiKhamBenhBHYT(80aTH-BHYT).xls"));
	EndWaitCursor();
} 
void CFMGeneralInsuranceCost80ATH::OnDeptListDblClick(){
	
} 
void CFMGeneralInsuranceCost80ATH::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMGeneralInsuranceCost80ATH::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMGeneralInsuranceCost80ATH::OnDeptListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;

	m_wndDeptList.BeginLoad();
	int nCount = 0;

	szSQL.Format(_T("SELECT sd_id as id, sd_name as name ") \
		         _T("FROM sys_dept ") \
				 _T("WHERE 1=1 AND sd_type='DT' ORDER BY id "));

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDeptList.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	m_wndDeptList.EndLoad();
	return nCount;
}
int CFMGeneralInsuranceCost80ATH::OnAddFMGeneralInsuranceCost80ATH(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMGeneralInsuranceCost80ATH::OnEditFMGeneralInsuranceCost80ATH(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMGeneralInsuranceCost80ATH::OnDeleteFMGeneralInsuranceCost80ATH(){
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
 		OnCancelFMGeneralInsuranceCost80ATH();
 	}
	return 0;
}
int CFMGeneralInsuranceCost80ATH::OnSaveFMGeneralInsuranceCost80ATH(){
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
 		//OnFMGeneralInsuranceCost80ATHListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CFMGeneralInsuranceCost80ATH::OnCancelFMGeneralInsuranceCost80ATH(){
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
int CFMGeneralInsuranceCost80ATH::OnFMGeneralInsuranceCost80ATHListLoadData(){
	return 0;
}

int CFMGeneralInsuranceCost80ATH::OnDeptListCheckAll()
{
	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (!m_wndDeptList.GetCheck(i))
		{
			m_wndDeptList.SetCheck(i, TRUE);
		}
	}
	return 0;
}

int CFMGeneralInsuranceCost80ATH::OnDeptListUnCheckAll()
{
	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			m_wndDeptList.SetCheck(i, FALSE);
		}
	}
	return 0;
}

CString CFMGeneralInsuranceCost80ATH::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere, szSubWhere, szMaterialID, szGroup;
	CString szDepts;

	szWhere.Empty();
	szSubWhere.Empty();
	szDepts.Empty();
	szMaterialID.Empty();
	szSQL.Format(_T("SELECT hsd_id id FROM hms_surgery_drugtype WHERE hsd_type = 'M'"));
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		if (!szMaterialID.IsEmpty())
			szMaterialID += _T(", ");
		szMaterialID.AppendFormat(_T("'%s'"),rs.GetValue(_T("id")));
		rs.MoveNext();
	}
	if (!szMaterialID.IsEmpty())
		szMaterialID += _T(", ");
	szMaterialID += _T("'A1500', 'A1600'");
	szWhere.Format(_T(" AND fi.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND fi.hfe_staff='%s' "), m_szClerkKey);

	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			if (!szDepts.IsEmpty())
				szDepts += _T(",");
			szDepts.AppendFormat(_T("'%s'"), m_wndDeptList.GetItemText(i, 0));
		}
	}

	if (!szDepts.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_deptid IN(%s) "), szDepts);
	}
	if (m_bOutpatientTreatment)
	{
		szWhere.AppendFormat(_T(" AND nvl(htr_outpatient, 'X') = 'Y'"));
		szGroup = _T("D0000");
	}
	else
	{
		szWhere.AppendFormat(_T(" AND nvl(htr_outpatient, 'X') <> 'Y' AND fe.hfe_group <> 'D0000'"));
		szGroup = _T("C0000");
	}
	szSQL.Format(_T(" SELECT SUM(total) AS total,") \
				_T("        hdline,") \
				_T("        linename,") \
				_T("        insline,") \
				_T("        SUM(drugfee) AS drugfee,") \
				_T("        SUM(drugfeek) AS drugfeek,") \
				_T("        SUM(bloodfee) AS bloodfee,") \
				_T("        SUM(testfee) AS testfee,") \
				_T("        SUM(pacsfee) AS pacsfee,") \
				_T("        SUM(normtechfee) AS normtechfee,") \
				_T("        SUM(hitechfee) AS hitechfee,") \
				_T("        SUM(materialfee) AS materialfee,") \
				_T("        SUM(replace_materialfee) AS replace_materialfee,") \
				_T("        SUM(bedfee) AS bedfee,") \
				_T("        SUM(transportfee) AS transportfee,") \
				_T("        SUM(cost) AS cost,") \
				_T("        SUM(otherfee) AS otherfee,") \
				_T("        SUM(inspaid) AS inspaid,") \
				_T("        SUM(patpaid) AS patpaid") \
				_T(" FROM") \
				_T(" (") \
				_T("  SELECT hfe_docno,") \
				_T("         1 AS total,") \
				_T("         CASE WHEN hd_insline ='Y' THEN 1 ELSE 0 END AS hdline,") \
				_T("         hc_line AS insline,") \
				_T("         (SELECT distinct ss_default FROM sys_sel WHERE ss_id='hms_insline' AND ss_code=CAST(hc_line AS NVARCHAR2(15))) AS linename,") \
				_T("         round(SUM(drugfee)) AS drugfee,") \
				_T("         round(SUM(drugfeek)) AS drugfeek,") \
				_T("         round(SUM(bloodfee)) AS bloodfee,") \
				_T("         round(SUM(testfee)) AS testfee,") \
				_T("         round(SUM(pacsfee)) AS pacsfee,") \
				_T("         round(SUM(normtechfee)) AS normtechfee,") \
				_T("         round(SUM(hitechfee)) AS hitechfee,") \
				_T("         round(SUM(materialfee)) AS materialfee,") \
				_T("         round(SUM(replace_materialfee)) AS replace_materialfee,") \
				_T("         round(SUM(bedfee)) AS bedfee,") \
				_T("         round(SUM(transportfee)) AS transportfee,") \
				_T("         round(SUM(inspaid)) AS cost,") \
				_T("         round(SUM(otherfee)) AS otherfee,") \
				_T("         round(SUM(discount)) AS inspaid,") \
				_T("         round(SUM(inspaid-discount)) AS patpaid") \
				_T("  FROM") \
				_T("  (") \
				_T("   SELECT fe.hfe_docno,") \
				_T("          CASE WHEN SUBSTR(hfe_group, 1, 2) IN('A1','A6') AND hfe_group NOT IN ('A1400', 'A1500', 'A1600') THEN fe.hfe_inspaid ELSE 0 END AS drugfee,") \
				_T("          CASE WHEN hfe_group='A1400' THEN fe.hfe_inspaid ELSE 0 END AS drugfeek,") \
				_T("          CASE WHEN NVL(hfl_bloodfee, 'N') = 'Y' THEN fe.hfe_inspaid ELSE 0 END AS bloodfee,") \
				_T("          CASE WHEN SUBSTR(hfe_group, 1, 2)='B1' AND fe.hfe_hitech ='N' AND NVL(hfl_bloodfee, 'N') <> 'Y' THEN fe.hfe_inspaid ELSE 0 END AS testfee,") \
				_T("          CASE WHEN SUBSTR(hfe_group, 1, 2) IN('B2','B3') AND hfe_hitech='N' THEN fe.hfe_inspaid ELSE 0 END AS pacsfee,") \
				_T("          CASE WHEN SUBSTR(hfe_group, 1, 2) IN('B4','B5') AND hfe_hitech='N' THEN fe.hfe_inspaid ELSE 0 END AS normtechfee,") \
				_T("          CASE WHEN hfe_hitech='Y' THEN fe.hfe_inspaid ELSE 0 END AS hitechfee,") \
				_T("          CASE WHEN hfe_group IN (%s) AND fe.hfe_feegroup NOT IN ('OPT_L', 'OPT') THEN fe.hfe_inspaid ELSE 0 END AS materialfee,") \
				_T("          CASE WHEN hfe_group IN (%s) AND fe.hfe_feegroup IN ('OPT_L', 'OPT') THEN fe.hfe_inspaid ELSE 0 END AS replace_materialfee,") \
				_T("          CASE WHEN hfe_group='%s' THEN fe.hfe_inspaid ELSE 0 END AS bedfee,") \
				_T("          CASE WHEN hfe_group='E0000' THEN fe.hfe_inspaid ELSE 0 END AS transportfee,") \
				_T("          CASE WHEN hfe_group='F0000' THEN fe.hfe_inspaid ELSE 0 END AS otherfee,") \
				_T("          fe.hfe_cost AS cost,") \
				_T("          CASE WHEN fe.hfe_feegroup NOT IN ('OPT_L', 'HITECH_L') THEN fe.hfe_inspaid ELSE 0 END AS inspaid,") \
				_T("          CASE WHEN fe.hfe_feegroup NOT IN ('OPT_L', 'HITECH_L') THEN fe.hfe_discount ELSE 0 END AS discount,") \
				_T("          fe.hfe_patpaid AS patpaid,") \
				_T("          fe.hfe_diffpaid AS difpaid") \
				_T("   FROM hms_fee_invoice fi") \
				_T("   LEFT JOIN hms_clinical_record ON (hcr_docno = fi.hfe_docno AND hcrf_invoicefee = fi.hfe_invoiceno)") \
				_T("   LEFT JOIN hms_treatment_record ON (hcr_docno = htr_docno AND hcr_refidx = htr_idx)") \
				_T("   LEFT JOIN hms_fee_view fe ON(fe.hfe_invoiceno=fi.hfe_invoiceno AND fe.hfe_docno=fi.hfe_docno)") \
				_T("   LEFT JOIN hms_fee_list ON (fe.hfe_itemid = hfl_feeid)") \
				_T("   WHERE fi.hfe_class IN('A', 'I') AND fe.hfe_discount > 0 AND fi.hfe_status='P'") \
				_T("   AND hcr_status = 'T' AND fe.hfe_type <> 'V' AND fe.hfe_category <> 'Y' %s") \
				_T("   UNION ALL") \
				_T("   SELECT fe.hfe_docno,") \
				_T("		  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,") \
				_T("          fe.hfe_cost AS cost,") \
				_T("          CASE WHEN fe.hfe_feegroup IN ('OPT_L', 'HITECH_L') THEN fe.hfe_inspaid ELSE 0 END AS inspaid,") \
				_T("          CASE WHEN fe.hfe_feegroup IN ('OPT_L', 'HITECH_L') THEN fe.hfe_discount ELSE 0 END AS discount,") \
				_T("          fe.hfe_patpaid AS patpaid,") \
				_T("          fe.hfe_diffpaid AS difpaid") \
				_T("   FROM hms_fee_invoice fi") \
				_T("   LEFT JOIN hms_clinical_record ON (hcr_docno = fi.hfe_docno AND hcrf_invoicefee = fi.hfe_invoiceno)") \
				_T("   LEFT JOIN hms_treatment_record ON (hcr_docno = htr_docno AND hcr_refidx = htr_idx)") \
				_T("   LEFT JOIN hms_fee_view fe ON(fe.hfe_invoiceno=fi.hfe_invoiceno AND fe.hfe_docno=fi.hfe_docno)") \
				_T("   LEFT JOIN hms_fee_list ON (fe.hfe_itemid = hfl_feeid)") \
				_T("   WHERE fi.hfe_class IN('A', 'I') AND fe.hfe_discount > 0 AND fi.hfe_status='P'") \
				_T("   AND hcr_status = 'T' AND fe.hfe_type = 'V' AND fe.hfe_category <> 'Y' %s") \
				_T("  ) tbl") \
				_T("  LEFT JOIN hms_doc ON(hfe_docno=hd_docno)") \
				_T("  LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx)") \
				_T("  LEFT JOIN hms_object ON(ho_id=hd_object)") \
				_T("  WHERE ho_type IN('I','C') AND LENGTH(hc_cardno) > 1") \
				_T("  GROUP BY hfe_docno, hc_line, hd_insline") \
				_T("  HAVING SUM(cost) > 0") \
				_T(" ) tblMain") \
				_T(" GROUP BY insline, linename, hdline") \
				_T(" ORDER BY insline, hdline"),
				szMaterialID, szMaterialID, szGroup, szWhere, szWhere);

	return szSQL;
}