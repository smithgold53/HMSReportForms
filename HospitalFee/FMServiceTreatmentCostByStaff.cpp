#include "stdafx.h"
#include "FMServiceTreatmentCostByStaff.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMServiceTreatmentCostByStaff *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMServiceTreatmentCostByStaff* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceTreatmentCostByStaff *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMServiceTreatmentCostByStaff *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMServiceTreatmentCostByStaff *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMServiceTreatmentCostByStaff* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceTreatmentCostByStaff *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMServiceTreatmentCostByStaff *pVw = (CFMServiceTreatmentCostByStaff *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMServiceTreatmentCostByStaff *pVw = (CFMServiceTreatmentCostByStaff *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMServiceTreatmentCostByStaff*)pWnd)->OnLockedSelect();
}
static long _OnObjectListLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceTreatmentCostByStaff*)pWnd)->OnObjectListLoadData();
} 
static void _OnObjectListDblClickFnc(CWnd *pWnd){
	((CFMServiceTreatmentCostByStaff*)pWnd)->OnObjectListDblClick();
}
static void _OnObjectListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMServiceTreatmentCostByStaff*)pWnd)->OnObjectListSelectChange(nOldItem, nNewItem);
} 
static int _OnObjectListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMServiceTreatmentCostByStaff*)pWnd)->OnObjectListDeleteItem();
}
static int _OnAddFMServiceTreatmentCostByStaffFnc(CWnd *pWnd){
	 return ((CFMServiceTreatmentCostByStaff*)pWnd)->OnAddFMServiceTreatmentCostByStaff();
}
static int _OnEditFMServiceTreatmentCostByStaffFnc(CWnd *pWnd){
	 return ((CFMServiceTreatmentCostByStaff*)pWnd)->OnEditFMServiceTreatmentCostByStaff();
}
static int _OnDeleteFMServiceTreatmentCostByStaffFnc(CWnd *pWnd){
	 return ((CFMServiceTreatmentCostByStaff*)pWnd)->OnDeleteFMServiceTreatmentCostByStaff();
} 
static int _OnSaveFMServiceTreatmentCostByStaffFnc(CWnd *pWnd){
	 return ((CFMServiceTreatmentCostByStaff*)pWnd)->OnSaveFMServiceTreatmentCostByStaff();
} 
static int _OnCancelFMServiceTreatmentCostByStaffFnc(CWnd *pWnd){
	 return ((CFMServiceTreatmentCostByStaff*)pWnd)->OnCancelFMServiceTreatmentCostByStaff();
}

static int _OnObjectListCheckAllFnc(CWnd *pWnd){
	return ((CFMServiceTreatmentCostByStaff*)pWnd)->OnObjectListCheckAll();
}
static int _OnObjectListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMServiceTreatmentCostByStaff*)pWnd)->OnObjectListUnCheckAll();
}

CFMServiceTreatmentCostByStaff::CFMServiceTreatmentCostByStaff(CWnd* pParent)
{
	m_nDlgWidth = 450;
	m_nDlgHeight = 480;
	SetDefaultValues();
}
CFMServiceTreatmentCostByStaff::~CFMServiceTreatmentCostByStaff(){
}
void CFMServiceTreatmentCostByStaff::OnCreateComponents()
{
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 440, 440);
	m_wndObjectInfo.Create(this, _T("Object"), 10, 120, 435, 405);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 310, 55);
	m_wndReportPeriod.Create(this,315, 30, 435, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDate.Create(this,315, 60, 435, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 310, 85);
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 90, 115);
	m_wndClerk.Create(this,95, 90, 435, 115); 
	m_wndPrint.Create(this, _T("&Print"), 235, 445, 335, 470);
	m_wndExport.Create(this, _T("&Export"), 340, 445, 440, 470);
	m_wndLocked.Create(this, _T("Locked"), 335, 410, 435, 435);
	m_wndObjectList.Create(this,15, 145, 430, 400);

}
void CFMServiceTreatmentCostByStaff::OnInitializeComponents()
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
	m_wndClerk.LimitText(65);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 240);


	m_wndObjectList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndObjectList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndObjectList.InsertColumn(2, _T("Type"), CFMT_TEXT, 0);
	m_wndObjectList.SetCheckBox(TRUE);

}
void CFMServiceTreatmentCostByStaff::OnSetWindowEvents()
{
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
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);
	m_wndObjectList.SetEvent(WE_SELCHANGE, _OnObjectListSelectChangeFnc);
	m_wndObjectList.SetEvent(WE_LOADDATA, _OnObjectListLoadDataFnc);
	m_wndObjectList.SetEvent(WE_DBLCLICK, _OnObjectListDblClickFnc);

	m_wndObjectList.AddEvent(1, _T("Check All"), _OnObjectListCheckAllFnc);
	m_wndObjectList.AddEvent(2, _T("UnCheck All"), _OnObjectListUnCheckAllFnc);

	CString szSQL;
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);

	OnObjectListLoadData();

	CString tmpStr;
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		tmpStr = m_wndObjectList.GetItemText(i, 2);

		if (tmpStr == _T("S"))
			m_wndObjectList.SetCheck(i, TRUE);
	}

}
void CFMServiceTreatmentCostByStaff::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);

}
void CFMServiceTreatmentCostByStaff::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMServiceTreatmentCostByStaff::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMServiceTreatmentCostByStaff::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bLocked=FALSE;

}
int CFMServiceTreatmentCostByStaff::SetMode(int nMode){
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
/*void CFMServiceTreatmentCostByStaff::OnYearChange(){
	
} */
/*void CFMServiceTreatmentCostByStaff::OnYearSetfocus(){
	
} */
/*void CFMServiceTreatmentCostByStaff::OnYearKillfocus(){
	
} */
int CFMServiceTreatmentCostByStaff::OnYearCheckValue()
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
void CFMServiceTreatmentCostByStaff::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMServiceTreatmentCostByStaff::OnReportPeriodSelendok()
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
/*void CFMServiceTreatmentCostByStaff::OnReportPeriodSetfocus(){
	
}*/
/*void CFMServiceTreatmentCostByStaff::OnReportPeriodKillfocus(){
	
}*/
long CFMServiceTreatmentCostByStaff::OnReportPeriodLoadData()
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
/*void CFMServiceTreatmentCostByStaff::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMServiceTreatmentCostByStaff::OnFromDateChange(){
	
} */
/*void CFMServiceTreatmentCostByStaff::OnFromDateSetfocus(){
	
} */
/*void CFMServiceTreatmentCostByStaff::OnFromDateKillfocus(){
	
} */
int CFMServiceTreatmentCostByStaff::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMServiceTreatmentCostByStaff::OnToDateChange(){
	
} */
/*void CFMServiceTreatmentCostByStaff::OnToDateSetfocus(){
	
} */
/*void CFMServiceTreatmentCostByStaff::OnToDateKillfocus(){
	
} */
int CFMServiceTreatmentCostByStaff::OnToDateCheckValue(){
	return 0;
} 
void CFMServiceTreatmentCostByStaff::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMServiceTreatmentCostByStaff::OnClerkSelendok(){
	 
}
/*void CFMServiceTreatmentCostByStaff::OnClerkSetfocus(){
	
}*/
/*void CFMServiceTreatmentCostByStaff::OnClerkKillfocus(){
	
}*/
long CFMServiceTreatmentCostByStaff::OnClerkLoadData()
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
/*void CFMServiceTreatmentCostByStaff::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMServiceTreatmentCostByStaff::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);

	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp;
	CString szSysDate;

	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/HF_BANGKETHANHTOAN_MAU6A.RPT")))
		return;

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);

	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	CReportSection *rptDetail;
	CString szOldLine, szNewLine;
	long double nGroupTotal[23];
	long double nTotal[23];
	double nCost;
	int nIndex = 1;

	for (int i = 0; i < 23; i++)
	{
		nGroupTotal[i] = 0;
		nTotal[i] = 0;
	}

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("recvdate"), szNewLine);

		if (!szNewLine.IsEmpty() && CompareDate(szNewLine, szOldLine) != 0)
		{
			if (nGroupTotal[21] + nGroupTotal[22] > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				//TranslateString(_T("\x43\x1ED9ng kho\x61"), tmpStr);
				rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng ng\xE0y"));

				for (int i = 6; i < 23; i++)
				{
					FormatCurrency(nGroupTotal[i], tmpStr);
					szTemp.Format(_T("s%d"), i + 1);
					rptDetail->SetValue(szTemp, tmpStr);
					nTotal[i] += nGroupTotal[i];
					nGroupTotal[i] = 0;
				}
			}

			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));

			szTemp = CDate::Convert(szNewLine, yyyymmdd, ddmmyyyy);

			tmpStr.Format(_T("Ng\xE0y %s"), szTemp);

			rptDetail->SetValue(_T("GroupName"), tmpStr);
			//rs.GetValue(_T("objid"), szObject);
			szOldLine = szNewLine;
			nIndex = 1;
		}

		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("username"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		//rs.GetValue(_T("recordno"), tmpStr);
		//rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("recvdate"), tmpStr);
		szTemp = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("4"), szTemp);

		rs.GetValue(_T("sn"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("deptid"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("phigiuong"), nCost);
		nGroupTotal[6] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("ptttktc"), nCost);
		nGroupTotal[7] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("boiduongpt"), nCost);
		nGroupTotal[8] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("thxn"), nCost);
		nGroupTotal[9] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.GetValue(_T("oxy"), nCost);
		nGroupTotal[10] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.GetValue(_T("dcvtch"), nCost);
		nGroupTotal[11] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);

		rs.GetValue(_T("thuocch"), nCost);
		nGroupTotal[12] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);

		rs.GetValue(_T("diennuoc"), nCost);
		nGroupTotal[13] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("14"), tmpStr);

		rs.GetValue(_T("quantrang"), nCost);
		nGroupTotal[14] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("15"), tmpStr);

		rs.GetValue(_T("thuockd"), nCost);
		nGroupTotal[15] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("16"), tmpStr);

		rs.GetValue(_T("dcvttb"), nCost);
		nGroupTotal[16] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("17"), tmpStr);

		rs.GetValue(_T("tienan"), nCost);
		nGroupTotal[17] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("18"), tmpStr);

		rs.GetValue(_T("monhanh"), nCost);
		nGroupTotal[18] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("19"), tmpStr);

		rs.GetValue(_T("total"), nCost);
		nGroupTotal[19] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("20"), tmpStr);

		rs.GetValue(_T("trichtg"), nCost);
		nGroupTotal[20] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("21"), tmpStr);

		rs.GetValue(_T("thu"), nCost);
		nGroupTotal[21] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("22"), tmpStr);

		rs.GetValue(_T("chi"), nCost);
		nGroupTotal[22] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("23"), tmpStr);

		rs.MoveNext();
	}

	if (nGroupTotal[21] + nGroupTotal[22] > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		//TranslateString(_T("\x43\x1ED9ng kho\x61"), tmpStr);
		rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng ng\xE0y"));

		for (int i = 6; i < 23; i++)
		{
			FormatCurrency(nGroupTotal[i], tmpStr);
			szTemp.Format(_T("s%d"), i + 1);
			rptDetail->SetValue(szTemp, tmpStr);
			nTotal[i] += nGroupTotal[i];
			nGroupTotal[i] = 0;
		}
	}

	if (nTotal[21] + nTotal[22] > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		//TranslateString(_T("T\x1ED5ng \x63\x1ED9ng"), tmpStr);
		rptDetail->SetValue(_T("TotalGroup"), _T("T\x1ED5ng \x63\x1ED9ng"));
		for (int i = 6; i < 23; i++)
		{
			FormatCurrency(nTotal[i], tmpStr);
			szTemp.Format(_T("s%d"), i + 1);
			rptDetail->SetValue(szTemp, tmpStr);
		}
	}

	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);

	EndWaitCursor();
	rpt.PrintPreview();
} 
void CFMServiceTreatmentCostByStaff::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	CString tmpStr, szTemp;

	UpdateData(TRUE);
	BeginWaitCursor();

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_ICONERROR | MB_OK);
		return;
	}

	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 3);
	xls.SetColumnWidth(1, 16);
	xls.SetColumnWidth(2, 7);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 3);
	xls.SetColumnWidth(5, 6);

	xls.SetColumnWidth(6, 11);
	xls.SetColumnWidth(7, 11);
	xls.SetColumnWidth(8, 11);
	xls.SetColumnWidth(9, 11);
	xls.SetColumnWidth(10, 11);
	xls.SetColumnWidth(11, 11);
	xls.SetColumnWidth(12, 11);
	xls.SetColumnWidth(13, 11);
	xls.SetColumnWidth(14, 11);
	xls.SetColumnWidth(15, 11);
	xls.SetColumnWidth(16, 11);
	xls.SetColumnWidth(17, 11);
	xls.SetColumnWidth(18, 11);
	xls.SetColumnWidth(19, 11);
	xls.SetColumnWidth(20, 11);
	xls.SetColumnWidth(21, 11);
	xls.SetColumnWidth(22, 11);

	int nRow = 1;
	int nCol = 0;

	xls.SetRowHeight(6, 22);
	xls.SetRowHeight(7, 52);

	xls.SetCellMergedColumns(0, 1, 4);
	xls.SetCellMergedColumns(0, 2, 4);

	xls.SetCellMergedColumns(19, 1, 4);
	xls.SetCellMergedColumns(19, 2, 4);

	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(19, 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(19, 2, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol, nRow + 3, 23);
	xls.SetCellMergedColumns(nCol, nRow + 4, 23);

	xls.SetCellText(nCol, nRow + 3, _T("\x42\x1EA2NG T\x1ED4NG H\x1EE2P TH\x41NH TO\xC1N \x43HI PH\xCD \x110I\x1EC0U TR\x1ECA"),
					FMT_TEXT | FMT_CENTER, true, 16, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);

	TranslateString(_T("Index"), tmpStr);
	xls.SetCellMergedRows(nCol, nRow + 5, 2);
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Full Name"), tmpStr);
	xls.SetCellMergedRows(nCol + 1, nRow + 5, 2);
	xls.SetCellText(nCol + 1, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Admission No"), tmpStr);
	xls.SetCellMergedRows(nCol + 2, nRow + 5, 2);
	xls.SetCellText(nCol + 2, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 3, nRow + 5, 2);
	xls.SetCellText(nCol + 3, nRow + 5, _T("Ng\xE0y th\x61nh to\xE1n"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	//TranslateString(_T("Treat Days"), tmpStr);
	xls.SetCellMergedRows(nCol + 4, nRow + 5, 2);
	xls.SetCellText(nCol + 4, nRow + 5, _T("SN"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Dept"), tmpStr);
	xls.SetCellMergedRows(nCol + 5, nRow + 5, 2);
	xls.SetCellText(nCol + 5, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedColumns(nCol + 6, nRow + 5, 11);
	xls.SetCellText(nCol + 6, nRow + 5, _T("\x43\xE1\x63 kho\x1EA3n ti\x1EC1n \x64\x1ECB\x63h v\x1EE5"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 6, nRow + 6, _T("Ti\x1EC1n gi\x1B0\x1EDDng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 7, nRow + 6, _T("Th\x1EE7 thu\x1EADt Ph\x1EABu thu\x1EADt KT\x43"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 8, nRow + 6, _T("\x42\x1ED3i \x64\x1B0\x1EE1ng PTT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 9, nRow + 6, _T("T\x1ED5ng h\x1EE3p \x58N"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 10, nRow + 6, _T("O\x78y"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 11, nRow + 6, _T("\x44\x43-VT \x63\x1EED\x61 h\xE0ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 12, nRow + 6, _T("Thu\x1ED1\x63 \x63\x1EED\x61 h\xE0ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 13, nRow + 6, _T("\x110i\x1EC7n n\x1B0\x1EDB\x63"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 14, nRow + 6, _T("Qu\xE2n tr\x61ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 15, nRow + 6, _T("Thu\x1ED1\x63 kho\x61 \x64\x1B0\x1EE3\x63"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 16, nRow + 6, _T("\x44\x43-VT kho\x61 T\x42"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 17, nRow + 5, 2);
	xls.SetCellText(nCol + 17, nRow + 5, _T("Ti\x1EC1n \x103n"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 18, nRow + 5, 2);
	xls.SetCellText(nCol + 18, nRow + 5, _T("M\x1ED5 nh\x61nh"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 19, nRow + 5, 2);
	xls.SetCellText(nCol + 19, nRow + 5, _T("T\x1ED5ng thu vi\x1EC7n ph\xED"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 20, nRow + 5, 2);
	xls.SetCellText(nCol + 20, nRow + 5, _T("Tr\xED\x63h t\x1EA1m g\x1EEDi"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedColumns(nCol + 21, nRow + 5, 2);
	xls.SetCellText(nCol + 21, nRow + 5, _T("\x43h\xEAnh l\x1EC7\x63h"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 21, nRow + 6, _T("Thu"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
	xls.SetCellText(nCol + 22, nRow + 6, _T("\x43hi"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	int nIndex = 1;
	nRow = 7;
	CString szOldLine, szNewLine;
	long double nGroupTotal[23];
	long double nTotal[23];
	double nCost;

	for (int i = 0; i < 23; i++)
	{
		nGroupTotal[i] = 0;
		nTotal[i] = 0;
	}

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("recvdate"), szNewLine);

		if (!szNewLine.IsEmpty() && CompareDate(szNewLine, szOldLine) != 0)
		{
			if (nGroupTotal[21] + nGroupTotal[22] > 0)
			{
				nRow++;
				xls.SetCellText(nCol + 1, nRow, _T("\x43\x1ED9ng ng\xE0y"), FMT_TEXT | FMT_WRAPING, true);
				for (int i = 6; i < 23; i++)
				{
					tmpStr.Format(_T("%.2lf"), nGroupTotal[i]);
					xls.SetCellText(nCol + i, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
					nTotal[i] += nGroupTotal[i];
					nGroupTotal[i] = 0;
				}
			}

			nRow++;

			szTemp = CDate::Convert(szNewLine, yyyymmdd, ddmmyyyy);
		
			tmpStr.Format(_T("Ng\xE0y %s"), szTemp);
			xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING, true); 
			szOldLine = szNewLine;
			nIndex = 1;
		}

		nRow++;

		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("username"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		/*rs.GetValue(_T("recordno"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);*/

		rs.GetValue(_T("recvdate"), tmpStr);
		szTemp = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		xls.SetCellText(nCol + 3, nRow, szTemp, FMT_DATE | FMT_WRAPING);

		rs.GetValue(_T("sn"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("deptid"), tmpStr);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("phigiuong"), nCost);
		nGroupTotal[6] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("ptttktc"), nCost);
		nGroupTotal[7] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("boiduongpt"), nCost);
		nGroupTotal[8] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("thxn"), nCost);
		nGroupTotal[9] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 9, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("oxy"), nCost);
		nGroupTotal[10] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 10, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("dcvtch"), nCost);
		nGroupTotal[11] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 11, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("thuocch"), nCost);
		nGroupTotal[12] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 12, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("diennuoc"), nCost);
		nGroupTotal[13] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 13, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("quantrang"), nCost);
		nGroupTotal[14] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 14, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("thuockd"), nCost);
		nGroupTotal[15] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 15, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("dcvttb"), nCost);
		nGroupTotal[16] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 16, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("tienan"), nCost);
		nGroupTotal[17] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 17, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("monhanh"), nCost);
		nGroupTotal[18] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 18, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("total"), nCost);
		nGroupTotal[19] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 19, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("trichtg"), nCost);
		nGroupTotal[20] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 20, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("thu"), nCost);
		nGroupTotal[21] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 21, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("chi"), nCost);
		nGroupTotal[22] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 22, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.MoveNext();
	}

	if (nGroupTotal[21] + nGroupTotal[22] > 0)
	{
		nRow++;
		xls.SetCellText(nCol + 1, nRow, _T("\x43\x1ED9ng ng\xE0y"), FMT_TEXT | FMT_WRAPING, true);
		for (int i = 6; i < 23; i++)
		{
			tmpStr.Format(_T("%.2lf"), nGroupTotal[i]);
			xls.SetCellText(nCol + i, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
			nTotal[i] += nGroupTotal[i];
			nGroupTotal[i] = 0;
		}
	}

	if (nTotal[21] + nTotal[22] > 0)
	{
		nRow++;
		xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_WRAPING, true);
		for (int i = 6; i < 23; i++)
		{
			tmpStr.Format(_T("%.2lf"), nTotal[i]);
			xls.SetCellText(nCol + i, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
		}
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\THChiPhiDTTheoNguoiThu.xls"));
} 
void CFMServiceTreatmentCostByStaff::OnLockedSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMServiceTreatmentCostByStaff::OnObjectListDblClick(){
	
} 
void CFMServiceTreatmentCostByStaff::OnObjectListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMServiceTreatmentCostByStaff::OnObjectListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMServiceTreatmentCostByStaff::OnObjectListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	CString szWhere;
	m_wndObjectList.BeginLoad(); 
	int nCount = 0;

	szWhere.Format(_T(" AND ho_type IN('S') "));

	szSQL.Format(_T("SELECT ho_id AS ID, ") \
		         _T("ho_desc AS Name, ") \
				 _T("ho_type AS Type ") \
		         _T("FROM hms_object ") \
				 _T("WHERE 1=1 %s") \
				 _T("ORDER BY cast(ho_id as integer)"),
				 szWhere);

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObjectList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")),
			rs.GetValue(_T("Type")), NULL);
		rs.MoveNext();
	}

	m_wndObjectList.EndLoad(); 
	return nCount;
}
int CFMServiceTreatmentCostByStaff::OnAddFMServiceTreatmentCostByStaff(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMServiceTreatmentCostByStaff::OnEditFMServiceTreatmentCostByStaff(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMServiceTreatmentCostByStaff::OnDeleteFMServiceTreatmentCostByStaff(){
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
 		OnCancelFMServiceTreatmentCostByStaff();
 	}
	return 0;
}
int CFMServiceTreatmentCostByStaff::OnSaveFMServiceTreatmentCostByStaff(){
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
 		//OnFMServiceTreatmentCostByStaffListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CFMServiceTreatmentCostByStaff::OnCancelFMServiceTreatmentCostByStaff(){
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
int CFMServiceTreatmentCostByStaff::OnFMServiceTreatmentCostByStaffListLoadData(){
	return 0;
}

CString CFMServiceTreatmentCostByStaff::GetQueryString()
{
	CString szSQL, szWhere;
	CString szObjects;
	CString szCondition;

	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (m_wndObjectList.GetCheck(i))
		{
			if (!szObjects.IsEmpty())
				szObjects += _T(",");
			szObjects.AppendFormat(_T("'%s'"), m_wndObjectList.GetItemText(i, 0));
		}
	}

	if (!szObjects.IsEmpty())
	{
		szCondition.AppendFormat(_T(" and hd_object in(%s) "), szObjects);
	}

	if (!m_szClerkKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and fi.hfe_staff='%s' "), m_szClerkKey);
	}

	if (!m_bLocked)
	{
		szWhere.AppendFormat(_T(" and fi.hfe_locked<>'Y' "));
		szWhere.AppendFormat(_T(" and fi.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}
	else
	{
		szWhere.AppendFormat(_T(" and fi.hfe_locked='Y' "));
		szWhere.AppendFormat(_T(" and fi.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}

	szSQL.Format(_T(" select su_userid as userid,") \
				_T("        su_name as username,") \
				_T("        recvdate,") \
				_T("        1 as sn,") \
				_T("        su_deptid as deptid,") \
				_T("        sum(phigiuong) as phigiuong,") \
				_T("        sum(ptttktc) as ptttktc,") \
				_T("        sum(boiduongpt) as boiduongpt,") \
				_T("        sum(thxn) as thxn,") \
				_T("        sum(oxy) as oxy,") \
				_T("        sum(dcvtch) as dcvtch,") \
				_T("        sum(thuocch) as thuocch,") \
				_T("        sum(diennuoc) as diennuoc,") \
				_T("        sum(quantrang) as quantrang,") \
				_T("        sum(thuockd) as thuockd,") \
				_T("        sum(dcvttb) as dcvttb,") \
				_T("        sum(tienan) as tienan,") \
				_T("        sum(monhanh) as monhanh,") \
				_T("        sum(trichtg) as trichtg,") \
				_T("        sum(total) as total,") \
				_T("        sum(thu) as thu,") \
				_T("        sum(chi) as chi") \
				_T(" from") \
				_T(" (") \
				_T("   select docno,") \
				_T("          staff,") \
				_T("          recvdate,") \
				_T("          sum(phigiuong) as phigiuong,") \
				_T("          sum(ptttktc) as ptttktc,") \
				_T("          sum(boiduongpt) as boiduongpt,") \
				_T("          sum(thxn) as thxn,") \
				_T("          sum(oxy) as oxy,") \
				_T("          sum(dcvtch) as dcvtch,") \
				_T("          sum(thuocch) as thuocch,") \
				_T("          sum(diennuoc) as diennuoc,") \
				_T("          sum(quantrang) as quantrang,") \
				_T("          sum(thuockd) as thuockd,") \
				_T("          sum(dcvttb) as dcvttb,") \
				_T("          sum(tienan) as tienan,") \
				_T("          sum(monhanh) as monhanh,") \
				_T("          trichtg,") \
				_T("          sum(phigiuong+ptttktc+boiduongpt+thxn+oxy+dcvtch+thuocch+") \
				_T("              diennuoc+quantrang+thuockd+dcvttb+tienan+monhanh) as total,") \
				_T("          case when trichtg-sum(phigiuong+ptttktc+boiduongpt+thxn+oxy+dcvtch+thuocch+") \
				_T("              diennuoc+quantrang+thuockd+dcvttb+tienan+monhanh) < 0 then sum(phigiuong+ptttktc+boiduongpt+thxn+oxy+dcvtch+thuocch+") \
				_T("              diennuoc+quantrang+thuockd+dcvttb+tienan+monhanh)-trichtg else 0 end as thu,") \
				_T("          case when trichtg-sum(phigiuong+ptttktc+boiduongpt+thxn+oxy+dcvtch+thuocch+") \
				_T("              diennuoc+quantrang+thuockd+dcvttb+tienan+monhanh) > 0 then trichtg-sum(phigiuong+ptttktc+boiduongpt+thxn+oxy+dcvtch+thuocch+") \
				_T("              diennuoc+quantrang+thuockd+dcvttb+tienan+monhanh) else 0 end as chi") \
				_T("   from") \
				_T("   (") \
				_T("     select fi.hfe_docno as docno,") \
				_T("            fi.hfe_invoiceno as invoiceno,") \
				_T("            Cast_Timestamp2Date(fi.hfe_date) as recvdate,") \
				_T("            fi.hfe_deptid as deptid,") \
				_T("            fi.hfe_staff as staff,") \
				_T("            case when fe.hfe_group='C0000' then fe.hfe_patpaid else 0 end as phigiuong,") \
				_T("            case when (substr(fe.hfe_group,1,2) in('B4','B5') and fe.hfe_hitech='N') or fe.hfe_hitech='Y' ") \
				_T("            then fe.hfe_patpaid else 0 end as ptttktc,") \
				_T("            0 as boiduongpt,") \
				_T("            case when substr(fe.hfe_group,1,2) in('B1','B2','B3') and fe.hfe_hitech='N' then fe.hfe_patpaid else 0 end as thxn,") \
				_T("            0 as oxy,") \
				_T("            case when hpo_storage_id in(15) and fe.hfe_type in('M') then fe.hfe_patpaid else 0 end as dcvtch,") \
				_T("            0 as thuocch,") \
				_T("            0 as diennuoc,") \
				_T("            0 as quantrang,") \
				_T("            0 as thuockd,") \
				_T("            case when hpo_storage_id in(5,6) and fe.hfe_type in('M') then fe.hfe_patpaid else 0 end as dcvttb,") \
				_T("            case when substr(fe.hfe_group,1,2) in('FF') then fe.hfe_patpaid else 0 end as tienan,") \
				_T("            0 as monhanh,") \
				_T("            fi.hfe_deposit as trichtg") \
				_T("     from hms_fee fe") \
				_T("     left join hms_fee_invoice fi on(fi.hfe_invoiceno=fe.hfe_invoiceno and fe.hfe_docno=fi.hfe_docno)") \
				_T("     left join hms_ipharmaorder on(hpo_docno=fe.hfe_docno and hpo_orderid=fe.hfe_orderid ") \
				_T("                                   and hpo_ordertype in('D','M'))") \
				_T("     where fi.hfe_class in('I', 'A') and fi.hfe_status='P' %s") \
				_T("   ) Tbl") \
				_T("   left join hms_doc on(hd_docno=docno)") \
				_T("   left join hms_object on(ho_id=hd_object)") \
				_T("   where 1=1 %s") \
				_T("   group by docno, staff, recvdate, trichtg") \
				_T("   order by recvdate, docno") \
				_T(" ) Tbl1") \
				_T(" left join sys_user on(su_userid=staff)") \
				_T(" where trichtg+total > 0") \
				_T(" group by su_userid, su_name, recvdate, su_deptid") \
				_T(" order by recvdate, su_userid"),
				szWhere, szCondition);

	//MessageBox(szSQL);

	return szSQL;
}

int CFMServiceTreatmentCostByStaff::OnObjectListCheckAll()
{
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (!m_wndObjectList.GetCheck(i))
		{
			m_wndObjectList.SetCheck(i, TRUE);
		}
	}
	return 0;
}

int CFMServiceTreatmentCostByStaff::OnObjectListUnCheckAll()
{
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (m_wndObjectList.GetCheck(i))
		{
			m_wndObjectList.SetCheck(i, FALSE);
		}
	}
	return 0;
}