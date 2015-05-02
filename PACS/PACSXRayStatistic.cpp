#include "stdafx.h"
#include "PACSXRayStatistic.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "StringToken.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CPACSXRayStatistic *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSXRayStatistic* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CPACSXRayStatistic *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSXRayStatistic *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSXRayStatistic *)pWnd)->OnToDateCheckValue();
} 
static void _OnDepartmentSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSXRayStatistic* )pWnd)->OnDepartmentSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDepartmentSelendokFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnDepartmentSelendok();
}
/*static void _OnDepartmentSetfocusFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnDepartmentKillfocus();
}*/
/*static void _OnDepartmentKillfocusFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnDepartmentKillfocus();
}*/
static long _OnDepartmentLoadDataFnc(CWnd *pWnd){
	return ((CPACSXRayStatistic *)pWnd)->OnDepartmentLoadData();
}
/*static void _OnDepartmentAddNewFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnDepartmentAddNew();
}*/
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSXRayStatistic* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CPACSXRayStatistic *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CPACSXRayStatistic *)pWnd)->OnObjectAddNew();
}*/
static void _OnExportSelectFnc(CWnd *pWnd){
	CPACSXRayStatistic *pVw = (CPACSXRayStatistic *)pWnd;
	pVw->OnExportSelect();
}
static long _OnRoomListLoadDataFnc(CWnd *pWnd){
	return ((CPACSXRayStatistic*)pWnd)->OnRoomListLoadData();
} 
static void _OnRoomListDblClickFnc(CWnd *pWnd){
	((CPACSXRayStatistic*)pWnd)->OnRoomListDblClick();
} 
static void _OnRoomListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPACSXRayStatistic*)pWnd)->OnRoomListSelectChange(nOldItem, nNewItem);
}
static int _OnRoomListDeleteItemFnc(CWnd *pWnd){
	return ((CPACSXRayStatistic*)pWnd)->OnRoomListDeleteItem();
} 
static int _OnListCheckAllFnc(CWnd *pWnd){
	return ((CPACSXRayStatistic*)pWnd)->OnListCheckAll();
}
static int _OnListUnCheckAllFnc(CWnd *pWnd){
	return ((CPACSXRayStatistic*)pWnd)->OnListUnCheckAll();
}
CPACSXRayStatistic::CPACSXRayStatistic(CWnd *pParent){

	m_nDlgWidth = 475;
	m_nDlgHeight = 509;
	SetDefaultValues();
}
CPACSXRayStatistic::~CPACSXRayStatistic(){
}
void CPACSXRayStatistic::OnCreateComponents(){
	m_wndReportFilter.Create(this, _T("Report Parameters"), 5, 5, 465, 500);
	m_wndRoomInfo.Create(this, _T("Room List"), 10, 150, 460, 465);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 115, 55);
	m_wndYear.Create(this,120, 30, 240, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 245, 30, 335, 55);
	m_wndReportPeriod.Create(this,340, 30, 460, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 115, 85);
	m_wndFromDate.Create(this,120, 60, 240, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 245, 60, 335, 85);
	m_wndToDate.Create(this,340, 60, 460, 85); 
	m_wndDepartmentLabel.Create(this, _T("Department"), 10, 90, 115, 115);
	m_wndDepartment.Create(this,120, 90, 460, 115); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 120, 115, 145);
	m_wndObject.SetCheckBox(true);
	m_wndObject.Create(this,120, 120, 460, 145);
	m_wndExport.Create(this, _T("&Export"), 365, 470, 465, 495);
	m_wndRoomList.Create(this,15, 175, 455, 460);
	m_wndRoomList.SetCheckBox(true);
}
void CPACSXRayStatistic::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	//m_wndDepartment.SetCheckValue(true);
	m_wndDepartment.LimitText(35);
	//m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(35);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

	m_wndDepartment.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndDepartment.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);


	m_wndObject.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 50);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);

	m_wndRoomList.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 80);
	m_wndRoomList.InsertColumn(1, _T("Name"), CFMT_TEXT, 320);
}
void CPACSXRayStatistic::OnSetWindowEvents(){
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
	m_wndDepartment.SetEvent(WE_SELENDOK, _OnDepartmentSelendokFnc);
	//m_wndDepartment.SetEvent(WE_SETFOCUS, _OnDepartmentSetfocusFnc);
	//m_wndDepartment.SetEvent(WE_KILLFOCUS, _OnDepartmentKillfocusFnc);
	m_wndDepartment.SetEvent(WE_SELCHANGE, _OnDepartmentSelectChangeFnc);
	m_wndDepartment.SetEvent(WE_LOADDATA, _OnDepartmentLoadDataFnc);
	//m_wndDepartment.SetEvent(WE_ADDNEW, _OnDepartmentAddNewFnc);
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndRoomList.SetEvent(WE_SELCHANGE, _OnRoomListSelectChangeFnc);
	m_wndRoomList.SetEvent(WE_LOADDATA, _OnRoomListLoadDataFnc);
	m_wndRoomList.SetEvent(WE_DBLCLICK, _OnRoomListDblClickFnc);
	m_wndRoomList.AddEvent(1, _T("Check All"), _OnListCheckAllFnc);
	m_wndRoomList.AddEvent(2, _T("UnCheck All"), _OnListUnCheckAllFnc);
	CString szSQL;
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	m_szDepartmentKey = pMF->GetDepartmentID();
	UpdateData(false);
	OnRoomListLoadData();
}
void CPACSXRayStatistic::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDepartment.GetDlgCtrlID(), m_szDepartmentKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);

}
void CPACSXRayStatistic::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDepartmentKey.Empty();
	m_szObjectKey.Empty();

}
int CPACSXRayStatistic::SetMode(int nMode){
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
/*void CPACSXRayStatistic::OnYearChange(){
	
} */
/*void CPACSXRayStatistic::OnYearSetfocus(){
	
} */
/*void CPACSXRayStatistic::OnYearKillfocus(){
	
} */
int CPACSXRayStatistic::OnYearCheckValue(){
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
void CPACSXRayStatistic::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSXRayStatistic::OnReportPeriodSelendok(){
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
/*void CPACSXRayStatistic::OnReportPeriodSetfocus(){
	
}*/
/*void CPACSXRayStatistic::OnReportPeriodKillfocus(){
	
}*/
long CPACSXRayStatistic::OnReportPeriodLoadData(){
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
/*void CPACSXRayStatistic::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CPACSXRayStatistic::OnFromDateChange(){
	
} */
/*void CPACSXRayStatistic::OnFromDateSetfocus(){
	
} */
/*void CPACSXRayStatistic::OnFromDateKillfocus(){
	
} */
int CPACSXRayStatistic::OnFromDateCheckValue(){
	return 0;
} 
/*void CPACSXRayStatistic::OnToDateChange(){
	
} */
/*void CPACSXRayStatistic::OnToDateSetfocus(){
	
} */
/*void CPACSXRayStatistic::OnToDateKillfocus(){
	
} */
int CPACSXRayStatistic::OnToDateCheckValue(){
	return 0;
} 
void CPACSXRayStatistic::OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSXRayStatistic::OnDepartmentSelendok(){
	 
}
/*void CPACSXRayStatistic::OnDepartmentSetfocus(){
	
}*/
/*void CPACSXRayStatistic::OnDepartmentKillfocus(){
	
}*/
long CPACSXRayStatistic::OnDepartmentLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if (m_wndDepartment.IsSearchKey() && !m_szDepartmentKey.IsEmpty())
	{
		szWhere.Format(_T(" AND sd_id='%s' "), m_szDepartmentKey);
	}

	m_wndDepartment.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT sd_id as id, sd_name as name FROM sys_dept ") \
		         _T("WHERE sd_type IN ('KB', 'DT') ORDER BY sd_id"));
	nCount = rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{ 
		m_wndDepartment.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPACSXRayStatistic::OnDepartmentAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CPACSXRayStatistic::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSXRayStatistic::OnObjectSelendok(){
	 
}
/*void CPACSXRayStatistic::OnObjectSetfocus(){
	
}*/
/*void CPACSXRayStatistic::OnObjectKillfocus(){
	
}*/
long CPACSXRayStatistic::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty()){
		szWhere.Format(_T(" and ho_id='%s' "), m_szObjectKey);
	}
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT ho_id as id, ho_desc as name FROM hms_object WHERE 1=1 %s ORDER BY cast(ho_id as integer) "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPACSXRayStatistic::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CPACSXRayStatistic::OnRoomListDblClick(){
	
} 
void CPACSXRayStatistic::OnRoomListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CPACSXRayStatistic::OnRoomListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CPACSXRayStatistic::OnRoomListLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndRoomList.BeginLoad(); 
	m_wndRoomList.DeleteAllItems(); 
	int nCount = 0;

	szSQL.Format(_T("SELECT hrl_id as ID, hrl_name as Name ") \
		_T("FROM hms_roomlist ") \
		_T("WHERE hrl_deptid='%s' ") \
		_T("ORDER BY hrl_id"), m_szDepartmentKey);

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndRoomList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndRoomList.EndLoad(); 
	return nCount;
}
int CPACSXRayStatistic::OnListCheckAll()
{
	for (int i = 0; i < m_wndRoomList.GetItemCount(); i++)
	{
		if (!m_wndRoomList.GetCheck(i))
			m_wndRoomList.SetCheck(i, TRUE);
	}
	return 0;
}
int CPACSXRayStatistic::OnListUnCheckAll()
{
	for (int i = 0; i < m_wndRoomList.GetItemCount(); i++)
	{
		if (m_wndRoomList.GetCheck(i))
			m_wndRoomList.SetCheck(i, FALSE);
	}
	return 0;
}
void CPACSXRayStatistic::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	OnInitDataString();
	CRecord rs(&pMF->m_db), crs(&pMF->m_db);
	CExcel xls;
	CString szSQL, tmpStr, szOldVendor, szNewVendor, szTemp;
	CStringToken token(m_szProductIDStr);
	double nTmp = 0;
	int nIdx = 1, nCol = 0, nRow = 0, i = 0, nOldCol = 0;
	long double nTotalAmt = 0;
	szSQL = GetQueryString();
	int nRet = rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 6);
	xls.SetCellMergedColumns(0, 3, 6);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr = _T("TH\x1ED0NG K\xCA T\x1ED4NG H\x1EE2P \x42\x1EC6NH NH\xC2N \x58 QU\x41NG");
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	nRow = 4;
	//Fixed col
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(6, 10);
	xls.SetCellMergedRows(0, nRow, 3);
	xls.SetCellMergedRows(1, nRow, 3);
	xls.SetCellMergedColumns(2, nRow, 4);
	xls.SetCellMergedColumns(2, nRow+1, 2);
	xls.SetCellMergedColumns(4, nRow+1, 2);
	xls.SetCellMergedRows(6, nRow, 3);
	xls.SetCellMergedRows(7, nRow, 3);
	xls.SetCellMergedRows(8, nRow, 3);
	xls.SetCellText(0, nRow, _T("STT"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	xls.SetCellText(1, nRow, _T("H\x1ECD v\xE0 t\xEAn"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	xls.SetCellText(2, nRow, _T("\x110\x1ED1i t\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	xls.SetCellText(2, nRow+1, _T("N\x1ED9i tr\xFA"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	xls.SetCellText(2, nRow+2, _T("K"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	xls.SetCellText(3, nRow+2, _T("SL"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	xls.SetCellText(4, nRow+1, _T("Ngo\x1EA1i tr\xFA"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	xls.SetCellText(4, nRow+2, _T("K"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	xls.SetCellText(5, nRow+2, _T("SL"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	xls.SetCellText(6, nRow, _T("S\x1ED1 h\x1ED3 s\x1A1"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	xls.SetCellText(7, nRow, _T("\x44i\x1EC7n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	xls.SetCellText(8, nRow, _T("Ghi \x63h\xFA"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	//Dynamic Col Material
	nCol = 9;
	szSQL = GetQueryString1();
	crs.ExecSQL(szSQL);
	while (!crs.IsEOF())
	{
		crs.GetValue(_T("vendor"), szNewVendor);
		if (szNewVendor != szOldVendor)
		{
			xls.SetCellText(nCol, nRow, szNewVendor, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
			szOldVendor = szNewVendor;
			if (i > 1)
				xls.SetCellMergedColumns(nOldCol, nRow, i);
			i = 0;
			nOldCol = nCol;
			//_msg(_T("%d"), i);
		}
		xls.SetCellMergedRows(nCol, nRow+1, 2);
		xls.SetCellText(nCol++, nRow+1, crs.GetValue(_T("product_name")), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
		i++;
		crs.MoveNext();
	}
	//Dynamic Col Xquang
	szSQL.Format(_T(" SELECT ss_desc ") \
				_T(" FROM   sys_sel ") \
				_T(" WHERE  ss_id = 'hms_fee_category' ") \
				_T("    AND ss_vndesc IN ( 'XQuang' ) ") \
				_T(" ORDER  BY Cast(ss_code AS INTEGER)"));
	crs.ExecSQL(szSQL);
	i = 0;
	xls.SetCellText(nCol, nRow, _T("\x43\xE1\x63 t\x1EA1ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	nOldCol = nCol;
	while (!crs.IsEOF())
	{
		xls.SetCellText(nCol++, nRow+1, crs.GetValue(_T("ss_desc")), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
		i++;
		crs.MoveNext();
	}
	xls.SetCellMergedColumns(nOldCol, nRow, i);
	xls.SetCellText(nOldCol + 8, nRow, _T("T\x1ED5ng ti\x1EC1n"), 4098, true);
	nCol = 0;
	nRow = 8;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("patient_name")), FMT_TEXT);
		rs.GetValue(_T("patient_class"), tmpStr);
		if (tmpStr == _T("I"))
		{
			xls.SetCellText(2, nRow, rs.GetValue(_T("dept_id")), FMT_TEXT);
			xls.SetCellText(3, nRow, _T("1"), FMT_TEXT);
		}
		else
		{
			xls.SetCellText(4, nRow, rs.GetValue(_T("dept_id")), FMT_TEXT);
			xls.SetCellText(5, nRow, _T("1"), FMT_TEXT);
		}
		xls.SetCellText(6, nRow, rs.GetValue(_T("doc_no")), FMT_TEXT);
		xls.SetCellText(7, nRow, rs.GetValue(_T("obj_name")), FMT_TEXT);
		xls.SetCellText(8, nRow, rs.GetValue(_T("position")), FMT_TEXT);
		nCol = 9;
		for (int i = 0; i < token.GetSize(); i++)
		{
			token.GetAt(i, szTemp);
			tmpStr.Format(_T("PID_%s"), szTemp);
			xls.SetCellText(nCol++, nRow, rs.GetValue(tmpStr), FMT_NUMBER1);
		}
		xls.SetCellText(nCol++, nRow, rs.GetValue(_T("heart_lung")), FMT_NUMBER1);
		xls.SetCellText(nCol++, nRow, rs.GetValue(_T("spine")), FMT_NUMBER1);
		xls.SetCellText(nCol++, nRow, rs.GetValue(_T("belly")), FMT_NUMBER1);
		xls.SetCellText(nCol++, nRow, rs.GetValue(_T("bone")), FMT_NUMBER1);
		xls.SetCellText(nCol++, nRow, rs.GetValue(_T("intestine")), FMT_NUMBER1);
		xls.SetCellText(nCol++, nRow, rs.GetValue(_T("uiv")), FMT_NUMBER1);
		xls.SetCellText(nCol++, nRow, rs.GetValue(_T("gullet")), FMT_NUMBER1);
		xls.SetCellText(nCol++, nRow, rs.GetValue(_T("stomach")), FMT_NUMBER1);
		rs.GetValue(_T("amount"), nTmp);
		nTotalAmt += nTmp;
		xls.SetCellText(nCol, nRow, double2str(nTmp), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nTotalAmt > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
		tmpStr.Format(_T("%.2f"), nTotalAmt);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_NUMBER1);
	}
	xls.Save(_T("Exports\\Thong ke BN lam XQuang.xls"));

} 

CString CPACSXRayStatistic::GetQueryString(){
	CString szSQL, szWhere, szWhere1, szWhere2, szField, szField1, szSumField, tmpStr, szTemp, szRoomID, szObjects;
	CStringToken token(m_szProductIDStr);
	for (int i = 0; i < token.GetSize(); i++)
	{
		token.GetAt(i, szTemp);
		tmpStr.Format(_T("PID_%s"), szTemp);
		if (!szSumField.IsEmpty())
			szSumField += _T(", ");
		szSumField.AppendFormat(_T("SUM(%s) %s"),tmpStr, tmpStr);
		if (!szField.IsEmpty())
			szField += _T(", ");
		szField.AppendFormat(_T("0 %s"), tmpStr);
		if (!szField1.IsEmpty())
			szField1 += _T(", ");
		szField1.AppendFormat(_T("CASE WHEN product_id = %s THEN hpol_qtyissue ELSE 0 END %s"), szTemp,  tmpStr);
	}
	szWhere1.AppendFormat(_T(" AND hpc_performdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')")
		, m_szFromDate, m_szToDate);
	szWhere2.AppendFormat(_T(" AND   hpo_approvedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')")
		, m_szFromDate, m_szToDate);
	for (int i = 0; i < m_wndRoomList.GetItemCount(); i++)
	{
		if (m_wndRoomList.GetCheck(i))
		{
			if (!szRoomID.IsEmpty())
				szRoomID += _T(",");
			szRoomID.AppendFormat(_T("%d"), str2int(m_wndRoomList.GetItemText(i, 0)));
		}
	}
	if (!szRoomID.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" and hpcl_proomid in(%s) "), szRoomID);
		szWhere2.AppendFormat(_T(" and hpo_roomid in(%s) "), szRoomID);
	}
	for (int i = 0; i < m_wndObject.GetItemCount(); i++)
	{
		if (m_wndObject.GetCheck(i))
		{
			m_wndObject.SetCurSel(i);
			if (!szObjects.IsEmpty())
				szObjects += _T(", ");
			szObjects += m_wndObject.GetCurrent(0);
		}
	}
	if (!m_szObjectKey.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" AND hpc_object IN(%s)"), m_szObjectKey);
		szWhere2.AppendFormat(_T(" AND hpo_object_id IN(%s)"), m_szObjectKey);
	}
	if (!m_szDepartmentKey.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" AND hpc_deptid = '%s'"), m_szDepartmentKey);
		//szWhere2.AppendFormat(_T(" AND hpo_deptid = '%s'"), m_szDepartmentKey);
	}
	szSQL.Format(_T(" SELECT get_patientname(doc_no) patient_name, ") \
				_T("        patient_class, ") \
				_T("		doc_no,") \
				_T("        CASE WHEN dept_id = 'C1.1' THEN 'P '||substr(hrl_name, 1, 2) ELSE dept_id END dept_id, ") \
				_T("		(select ss_vndesc from sys_sel where ss_id = 'hms_rank' and ss_code = hd_rank) position,") \
				_T("        get_objectname(obj) obj_name, ") \
				_T("		%s,") \
				_T("        SUM(heart_lung) heart_lung, ") \
				_T("        SUM(spine) spine, ") \
				_T("        SUM(belly) belly, ") \
				_T("        SUM(bone) bone, ") \
				_T("        SUM(intestine) intestine, ") \
				_T("        SUM(uiv) uiv, ") \
				_T("        SUM(gullet) gullet, ") \
				_T("        SUM(stomach) stomach, ") \
				_T("        SUM(amount) amount") \
				_T(" FROM   (SELECT    hpc_docno  doc_no, ") \
				_T("                   hpc_deptid dept_id, ") \
				_T("				   hpc_roomid room_id,") \
				_T("                   hpc_class  patient_class, ") \
				_T("				   hpc_orderid porder,") \
				_T("                   hpc_object obj, ") \
				_T("                   %s, ") \
				_T("                   CASE WHEN hfl_categoryid = 6 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END        heart_lung, ") \
				_T("                   CASE WHEN hfl_categoryid = 7 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END        spine, ") \
				_T("                   CASE WHEN hfl_categoryid = 8 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END        belly, ") \
				_T("                   CASE WHEN hfl_categoryid = 9 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END        bone, ") \
				_T("                   CASE WHEN hfl_categoryid = 10 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END        intestine, ") \
				_T("                   CASE WHEN hfl_categoryid = 11 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END        uiv, ") \
				_T("                   CASE WHEN hfl_categoryid = 12 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END        gullet, ") \
				_T("                   CASE WHEN hfl_categoryid = 13 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END        stomach, ") \
				_T("                   hfe_cost   amount ") \
				_T("         FROM      hms_pacsorder ") \
				_T("         LEFT JOIN hms_pacsorderline ON ( hpc_orderid = hpcl_orderid ) ") \
				_T("         LEFT JOIN hms_fee_list ON ( hfl_feeid = hpcl_itemid ) ") \
				_T("         LEFT JOIN hms_fee ON ( hfe_docno = hpc_docno ") \
				_T("                                AND hfe_orderid = hpc_orderid ") \
				_T("                                AND hfe_orderline = hpcl_orderlineid ) ") \
				_T("         WHERE     hpc_status NOT IN ( 'O', 'S' ) ") \
				_T("		%s") \
				_T("		 AND hfl_categoryid IN (6, 7, 8, 9, 10, 11, 12, 13)") \
				_T("         UNION ALL ") \
				_T("         SELECT  DISTINCT  hpc_docno  doc_no, ") \
				_T("                   hpc_deptid, ") \
				_T("				   hpc_roomid,") \
				_T("                   hpc_class  patient_class, ") \
				_T("				   hpc_orderid porder,") \
				_T("                   hpc_object obj, ") \
				_T("                   %s, ") \
				_T("                   0          heart_lung, ") \
				_T("                   0          spine, ") \
				_T("                   0          belly, ") \
				_T("                   0          bone, ") \
				_T("                   0          intestine, ") \
				_T("                   0          uiv, ") \
				_T("                   0          gullet, ") \
				_T("                   0          stomach, ") \
				_T("                   hpol_unitprice*hpol_qtyissue amount ") \
				_T("         FROM      hms_ipharmaorder ") \
				_T("		 LEFT JOIN hms_ipharmaorderline ON(hpol_orderid = hpo_orderid) ") \
				_T("         LEFT JOIN m_productitem_view ON(product_item_id  = hpol_product_item_id) ") \
				_T("         LEFT JOIN hms_pacsorder ON ( hpc_orderid = hpo_reference_id ) ") \
				_T("         LEFT JOIN hms_pacsorderline  ON ( hpc_orderid = hpcl_orderid ) ") \
				_T("         WHERE     hpc_status NOT IN ( 'O', 'S' ) AND hpo_orderstatus = 'A' AND substr(product_code, 1, 3) = 'PHI' AND hpo_deptid = 'C8' ") \
				_T("		%s)") \
				_T(" LEFT JOIN hms_doc ON (hd_docno = doc_no)") \
				_T(" LEFT JOIN hms_roomlist ON (hrl_deptid = dept_id AND hrl_id = room_id)") \
				_T(" GROUP  BY doc_no, ") \
				_T("           patient_class, ") \
				_T("           obj, ") \
				_T("		   hd_rank,") \
				_T("           dept_id, hrl_name") \
				_T(" ORDER BY doc_no "), szSumField, szField, szWhere1, szField1, szWhere2 );
	return szSQL;
}

CString CPACSXRayStatistic::GetQueryString1(){
	CString szSQL, szWhere, szSubWhere, szRoomID;
	for (int i = 0; i < m_wndRoomList.GetItemCount(); i++)
	{
		if (m_wndRoomList.GetCheck(i))
		{
			if (!szRoomID.IsEmpty())
				szRoomID += _T(",");
			szRoomID.AppendFormat(_T("%d"), str2int(m_wndRoomList.GetItemText(i, 0)));
		}
	}
	if (!szRoomID.IsEmpty())
		szSubWhere.AppendFormat(_T(" AND hpo_roomid in(%s) "), szRoomID);
// 	if (!m_szDepartmentKey.IsEmpty())
// 		szSubWhere.AppendFormat(_T(" AND hpo_deptid = '%s'"), m_szDepartmentKey);
	if (!m_szObjectKey.IsEmpty())
		szSubWhere.AppendFormat(_T(" AND hpc_object = %s"), m_szObjectKey);
	szWhere.Format(_T("  AND   hpo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s"), 
				  m_szFromDate, m_szToDate, szSubWhere);
	szSQL.Format(_T(" SELECT   distinct product_manufacturename vendor, ") \
				_T("           product_id, ") \
				_T("           product_name ") \
				_T(" FROM      hms_ipharmaorder ") \
				_T(" LEFT JOIN hms_ipharmaorderline ON(hpol_orderid = hpo_orderid) ") \
				_T(" LEFT JOIN m_productitem_view ON(product_item_id  = hpol_product_item_id) ") \
				_T(" LEFT JOIN hms_pacsorder ON ( hpc_orderid = hpo_reference_id ) ") \
				_T(" LEFT JOIN hms_pacsorderline  ON ( hpc_orderid = hpcl_orderid ) ") \
				_T(" WHERE     hpo_orderstatus = 'A' AND substr(product_code, 1, 3) = 'PHI' AND hpo_deptid = 'C8' %s") \
				_T(" ORDER BY trim(vendor), product_id"), szWhere);
	return szSQL;
}

void CPACSXRayStatistic::OnInitDataString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString1();
	rs.ExecSQL(szSQL);
	m_szProductIDStr.Empty();
	m_szProductNameStr.Empty();
	while (!rs.IsEOF())
	{
		if (!m_szProductIDStr.IsEmpty())
			m_szProductIDStr += _T(",");
		m_szProductIDStr += rs.GetValue(_T("product_id"));
		if (!m_szProductNameStr.IsEmpty())
			m_szProductNameStr += _T(",");
		m_szProductNameStr += rs.GetValue(_T("product_name"));
		rs.MoveNext();
	}

}