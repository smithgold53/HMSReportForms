#include "stdafx.h"
#include "FMDischargedPaidPatientList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMDischargedPaidPatientList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMDischargedPaidPatientList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMDischargedPaidPatientList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMDischargedPaidPatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMDischargedPaidPatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMDischargedPaidPatientList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMDischargedPaidPatientList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnClerkAddNew();
}*/
static void _OnStaffSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMDischargedPaidPatientList* )pWnd)->OnStaffSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStaffSelendokFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnStaffSelendok();
}
/*static void _OnStaffSetfocusFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnStaffKillfocus();
}*/
/*static void _OnStaffKillfocusFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnStaffKillfocus();
}*/
static long _OnStaffLoadDataFnc(CWnd *pWnd){
	return ((CFMDischargedPaidPatientList *)pWnd)->OnStaffLoadData();
}
/*static void _OnStaffAddNewFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList *)pWnd)->OnStaffAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMDischargedPaidPatientList *pVw = (CFMDischargedPaidPatientList *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMDischargedPaidPatientList *pVw = (CFMDischargedPaidPatientList *)pWnd;
	pVw->OnExportSelect();
}
/* 
static long _OnObjectListLoadDataFnc(CWnd *pWnd){
	return ((CFMDischargedPaidPatientList*)pWnd)->OnObjectListLoadData();
} 
static void _OnObjectListDblClickFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList*)pWnd)->OnObjectListDblClick();
} 
static void _OnObjectListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMDischargedPaidPatientList*)pWnd)->OnObjectListSelectChange(nOldItem, nNewItem);
} 
static int _OnObjectListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMDischargedPaidPatientList*)pWnd)->OnObjectListDeleteItem();
} */
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMDischargedPaidPatientList*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMDischargedPaidPatientList*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMDischargedPaidPatientList*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMDischargedPaidPatientList*)pWnd)->OnDeptListDeleteItem();
} 
static int _OnAddFMDischargedPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMDischargedPaidPatientList*)pWnd)->OnAddFMDischargedPaidPatientList();
} 
static int _OnEditFMDischargedPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMDischargedPaidPatientList*)pWnd)->OnEditFMDischargedPaidPatientList();
} 
static int _OnDeleteFMDischargedPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMDischargedPaidPatientList*)pWnd)->OnDeleteFMDischargedPaidPatientList();
} 
static int _OnSaveFMDischargedPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMDischargedPaidPatientList*)pWnd)->OnSaveFMDischargedPaidPatientList();
} 
static int _OnCancelFMDischargedPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMDischargedPaidPatientList*)pWnd)->OnCancelFMDischargedPaidPatientList();
}
/*
static int _OnObjectListCheckAllFnc(CWnd *pWnd){
	return ((CFMDischargedPaidPatientList*)pWnd)->OnObjectListCheckAll();
}
static int _OnObjectListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMDischargedPaidPatientList*)pWnd)->OnObjectListUnCheckAll();
}*/
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMDischargedPaidPatientList*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMDischargedPaidPatientList*)pWnd)->OnDeptListUnCheckAll();
}

CFMDischargedPaidPatientList::CFMDischargedPaidPatientList(CWnd* pParent)
{
	m_nDlgWidth = 450;
	m_nDlgHeight = 555;
	SetDefaultValues();
}
CFMDischargedPaidPatientList::~CFMDischargedPaidPatientList(){
}
void CFMDischargedPaidPatientList::OnCreateComponents()
{
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 550);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndStaffLabel.Create(this, _T("Staff"), 10, 90, 90, 115);
	m_wndStaff.Create(this,95, 90, 215, 115); 
	m_wndPrint.Create(this, _T("&Print"), 265, 555, 345, 580);
	m_wndExport.Create(this, _T("&Export"), 350, 555, 430, 580);
	m_wndDeptList.Create(this,10, 120, 425, 340); 
	
}
void CFMDischargedPaidPatientList::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	//m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(75);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 240);

	m_wndStaff.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndStaff.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);
	/*m_wndObjectList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndObjectList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndObjectList.InsertColumn(2, _T("Type"), CFMT_TEXT, 0);
	m_wndObjectList.SetCheckBox(TRUE);*/

	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.SetCheckBox(TRUE);

}
void CFMDischargedPaidPatientList::OnSetWindowEvents()
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
	/*m_wndObjectList.SetEvent(WE_SELCHANGE, _OnObjectListSelectChangeFnc);
	m_wndObjectList.SetEvent(WE_LOADDATA, _OnObjectListLoadDataFnc);
	m_wndObjectList.SetEvent(WE_DBLCLICK, _OnObjectListDblClickFnc);*/

	/*m_wndObjectList.AddEvent(1, _T("Check All"), _OnObjectListCheckAllFnc);
	m_wndObjectList.AddEvent(2, _T("UnCheck All"), _OnObjectListUnCheckAllFnc);*/
	m_wndStaff.SetEvent(WE_SELENDOK, _OnStaffSelendokFnc);
	//m_wndStaff.SetEvent(WE_SETFOCUS, _OnStaffSetfocusFnc);
	//m_wndStaff.SetEvent(WE_KILLFOCUS, _OnStaffKillfocusFnc);
	m_wndStaff.SetEvent(WE_SELCHANGE, _OnStaffSelectChangeFnc);
	m_wndStaff.SetEvent(WE_LOADDATA, _OnStaffLoadDataFnc);
	//m_wndStaff.SetEvent(WE_ADDNEW, _OnStaffAddNewFnc);

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
//	OnObjectListLoadData();
	OnDeptListLoadData();

	/*CString tmpStr;
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		tmpStr = m_wndObjectList.GetItemText(i, 2);

		if (tmpStr == _T("S"))
			m_wndObjectList.SetCheck(i, TRUE);
	}*/
}
void CFMDischargedPaidPatientList::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	//DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_TextEx(pDX, m_wndStaff.GetDlgCtrlID(), m_szStaffKey);

}
void CFMDischargedPaidPatientList::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMDischargedPaidPatientList::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMDischargedPaidPatientList::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();

}
int CFMDischargedPaidPatientList::SetMode(int nMode){
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
/*void CFMDischargedPaidPatientList::OnYearChange(){
	
} */
/*void CFMDischargedPaidPatientList::OnYearSetfocus(){
	
} */
/*void CFMDischargedPaidPatientList::OnYearKillfocus(){
	
} */
int CFMDischargedPaidPatientList::OnYearCheckValue()
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
void CFMDischargedPaidPatientList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMDischargedPaidPatientList::OnReportPeriodSelendok()
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
/*void CFMDischargedPaidPatientList::OnReportPeriodSetfocus(){
	
}*/
/*void CFMDischargedPaidPatientList::OnReportPeriodKillfocus(){
	
}*/
long CFMDischargedPaidPatientList::OnReportPeriodLoadData()
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
/*void CFMDischargedPaidPatientList::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMDischargedPaidPatientList::OnFromDateChange(){
	
} */
/*void CFMDischargedPaidPatientList::OnFromDateSetfocus(){
	
} */
/*void CFMDischargedPaidPatientList::OnFromDateKillfocus(){
	
} */
int CFMDischargedPaidPatientList::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMDischargedPaidPatientList::OnToDateChange(){
	
} */
/*void CFMDischargedPaidPatientList::OnToDateSetfocus(){
	
} */
/*void CFMDischargedPaidPatientList::OnToDateKillfocus(){
	
} */
int CFMDischargedPaidPatientList::OnToDateCheckValue(){
	return 0;
} 
void CFMDischargedPaidPatientList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMDischargedPaidPatientList::OnClerkSelendok(){
	 
}
/*void CFMDischargedPaidPatientList::OnClerkSetfocus(){
	
}*/
/*void CFMDischargedPaidPatientList::OnClerkKillfocus(){
	
}*/
long CFMDischargedPaidPatientList::OnClerkLoadData()
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
/*void CFMDischargedPaidPatientList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMDischargedPaidPatientList::OnStaffSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMDischargedPaidPatientList::OnStaffSelendok(){
	 
}
/*void CFMDischargedPaidPatientList::OnStaffSetfocus(){
	
}*/
/*void CFMDischargedPaidPatientList::OnStaffKillfocus(){
	
}*/
long CFMDischargedPaidPatientList::OnStaffLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStaff.IsSearchKey() && !m_szStaffKey.IsEmpty()){
		szWhere.Format(_T(" and su_userid='%s' "), m_szStaffKey);
	}
	m_wndStaff.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as userid, su_name as username FROM sys_user WHERE su_deptid = 'PTC' %s ORDER BY userid "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStaff.AddItems(
			rs.GetValue(_T("userid")), 
			rs.GetValue(_T("username")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMDischargedPaidPatientList::OnStaffAddNew(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} */
void CFMDischargedPaidPatientList::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CReportSection *rptDetail;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp, szSysDate, szOldLine, szNewLine, szObjectNames;
	bool bCheckServ = false, bCheckIns = false;
	long double nGroupTotal[12];
	long double nTotal[12];
	double nCost =0;
	int nIndex = 1;
	CStringArray strArr;
	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}
	
	if (!rpt.Init(_T("Reports/HMS/HF_DANHSACHBENHNHANRAVIENDATHANHTOAN.RPT")))
		return;
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	rpt.GetReportHeader()->SetValue(_T("Department"), pMF->GetCurrentDepartmentName());
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm),
	CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	strArr.Add(_T("\x64\x1ECB\x63h v\x1EE5"));
	strArr.Add(_T("\x42\x110 - \x43S - \x42H"));
	for(int i = 0; i < 12; i++)
	{
		nGroupTotal[i] = 0;
		nTotal[i] = 0;
	}
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("deptid"), szNewLine);
		if (!szNewLine.IsEmpty() && szNewLine != szOldLine)
		{
			if (nGroupTotal[9] > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				//TranslateString(_T("\x43\x1ED9ng kho\x61"), tmpStr);
				rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng kho\x61"));
				for(int i = 8; i < 12; i++)
				{
					FormatCurrency(nGroupTotal[i], tmpStr);
					szTemp.Format(_T("s%d"), i);
					rptDetail->SetValue(szTemp, tmpStr);
					nTotal[i] += nGroupTotal[i];
					nGroupTotal[i] = 0;
				}
			}
			szOldLine = szNewLine;
		}
		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("recordno"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		
		rs.GetValue(_T("ngayvao"), tmpStr);
		szTemp = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("5"), szTemp);
		
		rs.GetValue(_T("ngayra"), tmpStr);
		szTemp = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("6"), szTemp);

		if (nGroupTotal[8] == 0)
		{
			rs.GetValue(_T("deptid"), tmpStr);
			rptDetail->SetValue(_T("7"), tmpStr);
		}
		rs.GetValue(_T("deposit"), nCost);
		nGroupTotal[8] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("feeamt"), nCost);
		nGroupTotal[9] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("thu"), nCost);
		nGroupTotal[10] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);
		
		rs.GetValue(_T("chi"), nCost);
		nGroupTotal[11] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);
		
		rs.MoveNext();
	}

	if (nGroupTotal[9] > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		//TranslateString(_T("\x43\x1ED9ng kho\x61"), tmpStr);
		rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng kho\x61"));
		for(int i = 8; i < 12; i++)
		{
			FormatCurrency(nGroupTotal[i], tmpStr);
			szTemp.Format(_T("s%d"), i);
			rptDetail->SetValue(szTemp, tmpStr);
			nTotal[i] += nGroupTotal[i];
			nGroupTotal[i] = 0;
		}
	}

	if (nTotal[9] > 0)
	{
		for(int i = 8; i < 12; i++)
		{
			FormatCurrency(nTotal[i], tmpStr);
			szTemp.Format(_T("ss%d"), i);
			rpt.GetReportFooter()->SetValue(szTemp, tmpStr);
		}
	}

	CString szMoney;
	FormatCurrency(nTotal[9], szMoney);
	rpt.GetReportFooter()->SetValue(_T("Total"), szMoney);
	szMoney.Replace(_T("."), _T(""));
	MoneyToString(szMoney, szTemp);
	rpt.GetReportFooter()->SetValue(_T("SumInWords"), szTemp + _T(" \x111\x1ED3ng"));

	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
		          szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);

	EndWaitCursor();
	rpt.PrintPreview();
} 
void CFMDischargedPaidPatientList::OnExportSelect()
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

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 18);
	xls.SetColumnWidth(2, 7);
	xls.SetColumnWidth(3, 20);
	xls.SetColumnWidth(4, 11);
	xls.SetColumnWidth(5, 11);
	xls.SetColumnWidth(6, 11);
	xls.SetColumnWidth(7, 11);
	xls.SetColumnWidth(8, 11);
	xls.SetColumnWidth(9, 11);

	int nRow = 1;
	int nCol = 0;

	xls.SetRowHeight(6, 45);

	xls.SetCellMergedColumns(0, 1, 4);
	xls.SetCellMergedColumns(0, 2, 4);

	xls.SetCellMergedColumns(6, 1, 4);
	xls.SetCellMergedColumns(6, 2, 4);

	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(6, 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(6, 2, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol, nRow + 3, 10);
	xls.SetCellMergedColumns(nCol, nRow + 4, 10);

	xls.SetCellText(nCol, nRow + 3, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N T\x1EA0M G\x1EECI \x110\xC3 R\x41 VI\x1EC6N \x43H\x1AF\x41 TH\x41NH TO\xC1N"),
					FMT_TEXT | FMT_CENTER, true, 16, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);

	TranslateString(_T("Index"), tmpStr);
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Patient Name"), tmpStr);
	xls.SetCellText(nCol + 1, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Document No"), tmpStr);
	xls.SetCellText(nCol + 2, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Admission No"), tmpStr);
	xls.SetCellText(nCol + 3, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Dept"), tmpStr);
	xls.SetCellText(nCol + 4, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 5, nRow + 5, _T("Ti\x1EC1n t\x1EA1m g\x1EEDi"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Total Policy"), tmpStr);
	xls.SetCellText(nCol + 6, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 7, nRow + 5, _T("Ti\x1EC1n \x63hi ph\xED \x111\x1B0\x1EE3\x63 gi\x1EA3m"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Hospital Fee"), tmpStr);
	xls.SetCellText(nCol + 8, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 9, nRow + 5, _T("Thu"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 10, nRow + 5, _T("\x43hi"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	int nIndex = 1;
	nRow = 7;
	CString szOldLine, szNewLine;
	long double nGroupTotal[10];
	long double nTotal[10];
	double nCost;
	for(int i = 0; i < 10; i++)
	{
		nGroupTotal[i] = 0;
		nTotal[i] = 0;
	}
	while (!rs.IsEOF())
	{

		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("recordno"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("deptid"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("deposit"), nCost);
		nGroupTotal[4] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("chedo"), nCost);
		nGroupTotal[5] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("miengiam"), nCost);
		nGroupTotal[6] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("feeamt"), nCost);
		nGroupTotal[7] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("income"), nCost);
		nGroupTotal[8] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 9, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("outlay"), nCost);
		nGroupTotal[9] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 10, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.MoveNext();
		nRow++;
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\DanhSachBenhRaVaoVien.xls"));
} 
/*void CFMDischargedPaidPatientList::OnObjectListDblClick(){
	
} 
void CFMDischargedPaidPatientList::OnObjectListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMDischargedPaidPatientList::OnObjectListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMDischargedPaidPatientList::OnObjectListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	CString szWhere;
	m_wndObjectList.BeginLoad(); 
	int nCount = 0;

	//szWhere.Format(_T(" AND ho_type IN('S') "));

	szSQL.Format(_T("SELECT ho_id AS ID, ") \
		         _T("ho_desc AS Name, ") \
				 _T("ho_type AS Type ") \
		         _T("FROM hms_object ") \
				 _T("WHERE 1=1 %s ") \
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
}*/
void CFMDischargedPaidPatientList::OnDeptListDblClick(){
	
} 
void CFMDischargedPaidPatientList::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMDischargedPaidPatientList::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMDischargedPaidPatientList::OnDeptListLoadData()
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
int CFMDischargedPaidPatientList::OnAddFMDischargedPaidPatientList(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMDischargedPaidPatientList::OnEditFMDischargedPaidPatientList(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMDischargedPaidPatientList::OnDeleteFMDischargedPaidPatientList(){
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
 		OnCancelFMDischargedPaidPatientList();
 	}
	return 0;
}
int CFMDischargedPaidPatientList::OnSaveFMDischargedPaidPatientList(){
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
 		//OnFMDischargedPaidPatientListListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CFMDischargedPaidPatientList::OnCancelFMDischargedPaidPatientList(){
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

int CFMDischargedPaidPatientList::OnDeptListCheckAll()
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

int CFMDischargedPaidPatientList::OnDeptListUnCheckAll()
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


/*int CFMDischargedPaidPatientList::OnObjectListCheckAll()
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

int CFMDischargedPaidPatientList::OnObjectListUnCheckAll()
{
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (m_wndObjectList.GetCheck(i))
		{
			m_wndObjectList.SetCheck(i, FALSE);
		}
	}
	return 0;
}*/

CString CFMDischargedPaidPatientList::GetQueryString(){
	CString szSQL, szWhere, szObjects, szDepts, szWhere1;
	

	/*for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (m_wndObjectList.GetCheck(i))
		{
			if (!szObjects.IsEmpty())
				szObjects += _T(",");
			szObjects.AppendFormat(_T("'%s'"), m_wndObjectList.GetItemText(i, 0));
		}
	}*/

	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			if (!szDepts.IsEmpty())
				szDepts += _T(",");
			szDepts.AppendFormat(_T("'%s'"), m_wndDeptList.GetItemText(i, 0));
		}
	}

	//if (!szObjects.IsEmpty())
	//{
	//	szWhere.AppendFormat(_T(" and hfe_objectid in(%s) "), szObjects);
	//}

	if(!m_szStaffKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and hfe_approveby = '%s'"), m_szStaffKey);
	}
	if (!szDepts.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" and hcr_dischargedept in(%s) "), szDepts);
	}
	if(!m_szStaffKey.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" and i.hfe_staff = '%s'"), m_szStaffKey);
	}
	/*szSQL.Format(_T("SELECT    Get_patientname(docno) pname, docno,") \
				_T("          deptid, ") \
				_T("          recordno,") \
				_T("          ngayvao, ") \
				_T("          ngayra, ") \
				_T("          deposit, ") \
				_T("          feeamt, ") \
				_T("          CASE WHEN feeamt > deposit THEN feeamt - deposit ") \
				_T("          ELSE 0 ") \
				_T("          END AS thu, ") \
				_T("          CASE WHEN feeamt < deposit THEN deposit - feeamt ") \
				_T("          ELSE 0 ") \
				_T("          END AS chi ") \
				_T("FROM      (SELECT docno, ") \
				_T("                  deptid, ngayvao, ngayra,") \
				_T("                  recordno,") \
				_T("                  refidx, ") \
				_T("                  Sum(hfe_amount) deposit ") \
				_T("           FROM   (SELECT    d.hfe_docno docno, ") \
				_T("                             hcr_dischargedept deptid, ") \
				_T("                             hcr_recordno recordno, ") \
				_T("                             d.hfe_refidx refidx, ") \
				_T("                             d.hfe_amount, i.hfe_date as ngaythu,") \
				_T("							 a.hfe_approveby,") \
				_T("                             hcr_admitdate AS ngayvao, ") \
				_T("                             hcr_dischargedate AS ngayra ") \
				_T("                   FROM      hms_fee_deposit d ") \
				_T("				   LEFT JOIN hms_fee_invoice i ON (i.hfe_invoiceno = d.hfe_refidx) ") \
				_T("                   LEFT JOIN hms_doc dc ON (dc.hd_docno = d.hfe_docno) ") \
				_T("				   LEFT JOIN hms_fee_approve a ON ( a.hfe_docno = d.hfe_docno ) ") \
				_T("                   LEFT JOIN hms_clinical_record ON ( hcr_docno = d.hfe_docno ) ") \
				_T("                   WHERE     1 = 1 ") \
				_T("                             AND hcr_status = 'T' AND d.hfe_status NOT IN ('C', 'R') ") \
				_T("                             AND a.hfe_status IN ('P')  ")\
				_T("                             AND dc.hd_object = '7') ") \
				_T("           WHERE  ngaythu BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s') %s") \
				_T("           GROUP  BY docno, ") \
				_T("                     deptid, ") \
				_T("                     ngayvao, ") \
				_T("                     ngayra, ") \
				_T("                     recordno, ") \
				_T("                     refidx) ") \
				_T("LEFT JOIN (SELECT hfe_docno, ") \
				_T("                  Sum(hfe_cost) feeamt ") \
				_T("           FROM   hms_fee ") \
				_T("           WHERE  hfe_status IN ('P')") \
				_T("                  AND hfe_category = 'N' AND hfe_isapprove = 'O'") \
				_T("                  AND hfe_object = '7' ") \
				_T("           GROUP  BY hfe_docno) f ON ( f.hfe_docno = docno ) ") \
				_T("ORDER     BY deptid, ") \
				_T("             docno "), m_szFromDate, m_szToDate, szWhere);*/

	szSQL.Format(_T("SELECT    Get_patientname(docno) pname, ") \
				_T("          docno, ") \
				_T("          deptid, ") \
				_T("          recordno, ") \
				_T("          ngayvao, ") \
				_T("          ngayra, ") \
				_T("          deposit, ") \
				_T("          CASE WHEN feeamt IS NOT NULL THEN feeamt ") \
				_T("          ELSE 0 ") \
				_T("          END AS feeamt, ") \
				_T("          CASE WHEN feeamt > deposit THEN feeamt - deposit ") \
				_T("          ELSE 0 ") \
				_T("          END AS thu, ") \
				_T("          CASE WHEN feeamt < deposit THEN deposit - feeamt ") \
				_T("          ELSE ") \
				_T("              CASE WHEN feeamt IS NULL THEN deposit ") \
				_T("              ELSE 0 ") \
				_T("              END ") \
				_T("          END AS chi ") \
				_T("FROM      (SELECT docno, ") \
				_T("                  deptid, ") \
				_T("                  ngayvao, ") \
				_T("                  ngayra, ") \
				_T("                  recordno, ") \
				_T("                  refidx, ") \
				_T("                  Sum(hfe_amount) deposit ") \
				_T("           FROM   (SELECT    d.hfe_docno docno, ") \
				_T("                             hcr_dischargedept deptid, ") \
				_T("                             hcr_recordno recordno, ") \
				_T("                             d.hfe_refidx refidx, ") \
				_T("                             d.hfe_amount, ") \
				_T("                             i.hfe_date AS ngaythu, ") \
				_T("                             a.hfe_approveby, ") \
				_T("                             d.hfe_date AS ngayTU, ") \
				_T("                             hcr_admitdate AS ngayvao, ") \
				_T("                             hcr_dischargedate AS ngayra ") \
				_T("                   FROM      hms_fee_deposit d ") \
				_T("                   LEFT JOIN hms_fee_invoice i ON ( i.hfe_invoiceno = d.hfe_refidx ) ") \
				_T("                   LEFT JOIN hms_doc dc ON ( dc.hd_docno = d.hfe_docno ) ") \
				_T("                   LEFT JOIN hms_fee_approve a ON ( a.hfe_docno = d.hfe_docno ) ") \
				_T("                   LEFT JOIN hms_clinical_record ON ( hcr_docno = d.hfe_docno ) ") \
				_T("                   WHERE     1 = 1 ") \
				_T("                             AND hcr_status = 'T' ") \
				_T("                             AND d.hfe_status NOT IN ( 'C' ) ") \
				_T("                             AND a.hfe_status IN ( 'P', 'A' ) ") \
				_T("                             AND d.hfe_objectid = '7' %s) ") \
				_T("           WHERE  ngayTU BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s') ") \
				_T("                   OR ngaythu BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s') %s") \
				_T("           GROUP  BY docno, ") \
				_T("                     deptid, ") \
				_T("                     ngayvao, ") \
				_T("                     ngayra, ") \
				_T("                     recordno, ") \
				_T("                     refidx) ") \
				_T("LEFT JOIN (SELECT hfe_docno, ") \
				_T("                  Sum(hfe_cost) feeamt ") \
				_T("           FROM   hms_fee ") \
				_T("           WHERE  hfe_status IN ( 'P', 'A' ) ") \
				_T("                  AND hfe_category = 'N' ") \
				_T("                  AND hfe_isapprove = 'O' ") \
				_T("                  AND hfe_object = '7' ") \
				_T("           GROUP  BY hfe_docno) f ON ( f.hfe_docno = docno ) ") \
				_T("ORDER     BY deptid, ") \
				_T("             docno "), szWhere1, m_szFromDate, m_szToDate, m_szFromDate, m_szToDate, szWhere);


	_fmsg(_T("%s"), szSQL);
	return szSQL;
}
