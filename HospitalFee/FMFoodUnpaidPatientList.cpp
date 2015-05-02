#include "stdafx.h"
#include "FMFoodUnpaidPatientList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMFoodUnpaidPatientList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMFoodUnpaidPatientList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMFoodUnpaidPatientList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMFoodUnpaidPatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMFoodUnpaidPatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMFoodUnpaidPatientList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMFoodUnpaidPatientList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMFoodUnpaidPatientList *pVw = (CFMFoodUnpaidPatientList *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMFoodUnpaidPatientList *pVw = (CFMFoodUnpaidPatientList *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnObjectListLoadDataFnc(CWnd *pWnd){
	return ((CFMFoodUnpaidPatientList*)pWnd)->OnObjectListLoadData();
} 
static void _OnObjectListDblClickFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList*)pWnd)->OnObjectListDblClick();
} 
static void _OnObjectListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMFoodUnpaidPatientList*)pWnd)->OnObjectListSelectChange(nOldItem, nNewItem);
} 
static int _OnObjectListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMFoodUnpaidPatientList*)pWnd)->OnObjectListDeleteItem();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMFoodUnpaidPatientList*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMFoodUnpaidPatientList*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMFoodUnpaidPatientList*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMFoodUnpaidPatientList*)pWnd)->OnDeptListDeleteItem();
}
static void _OnDischargeSelectFnc(CWnd *pWnd){
	  ((CFMFoodUnpaidPatientList*)pWnd)->OnDischargeSelect();
}
static void _OnIndischargeSelectFnc(CWnd *pWnd){
	  ((CFMFoodUnpaidPatientList*)pWnd)->OnIndischargeSelect();
}
static int _OnAddFMFoodUnpaidPatientListFnc(CWnd *pWnd){
	 return ((CFMFoodUnpaidPatientList*)pWnd)->OnAddFMFoodUnpaidPatientList();
} 
static int _OnEditFMFoodUnpaidPatientListFnc(CWnd *pWnd){
	 return ((CFMFoodUnpaidPatientList*)pWnd)->OnEditFMFoodUnpaidPatientList();
} 
static int _OnDeleteFMFoodUnpaidPatientListFnc(CWnd *pWnd){
	 return ((CFMFoodUnpaidPatientList*)pWnd)->OnDeleteFMFoodUnpaidPatientList();
} 
static int _OnSaveFMFoodUnpaidPatientListFnc(CWnd *pWnd){
	 return ((CFMFoodUnpaidPatientList*)pWnd)->OnSaveFMFoodUnpaidPatientList();
} 
static int _OnCancelFMFoodUnpaidPatientListFnc(CWnd *pWnd){
	 return ((CFMFoodUnpaidPatientList*)pWnd)->OnCancelFMFoodUnpaidPatientList();
}

static int _OnObjectListCheckAllFnc(CWnd *pWnd){
	return ((CFMFoodUnpaidPatientList*)pWnd)->OnObjectListCheckAll();
}
static int _OnObjectListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMFoodUnpaidPatientList*)pWnd)->OnObjectListUnCheckAll();
}
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMFoodUnpaidPatientList*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMFoodUnpaidPatientList*)pWnd)->OnDeptListUnCheckAll();
}

CFMFoodUnpaidPatientList::CFMFoodUnpaidPatientList(CWnd* pParent)
{
	m_nDlgWidth = 450;
	m_nDlgHeight = 560;
	SetDefaultValues();
}
CFMFoodUnpaidPatientList::~CFMFoodUnpaidPatientList(){
}
void CFMFoodUnpaidPatientList::OnCreateComponents()
{
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 440, 520);
	m_wndObjectInfo.Create(this, _T("Object"), 10, 90, 435, 235);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 240, 435, 485);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 310, 55);
	m_wndReportPeriod.Create(this,315, 30, 435, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDate.Create(this,315, 60, 435, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 310, 85);
	m_wndPrint.Create(this, _T("&Print"), 235, 525, 335, 550);
	m_wndExport.Create(this, _T("&Export"), 340, 525, 440, 550);
	m_wndObjectList.Create(this,15, 115, 430, 230); 
	m_wndDeptList.Create(this,15, 265, 430, 480);


	m_wndDischarge.Create(this, _T("Discharge"), 345, 490, 435, 515);
	//m_wndDischarge.Create(this, _T("Discharge"), 250, 490, 340, 515);
	//m_wndIndischarge.Create(this, _T("Indischarge"), 345, 490, 435, 515);

}
void CFMFoodUnpaidPatientList::OnInitializeComponents()
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


	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.InsertColumn(2, _T("Type"), CFMT_TEXT, 0);
	m_wndDeptList.SetCheckBox(TRUE);

}
void CFMFoodUnpaidPatientList::OnSetWindowEvents()
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
	m_wndObjectList.SetEvent(WE_SELCHANGE, _OnObjectListSelectChangeFnc);
	m_wndObjectList.SetEvent(WE_LOADDATA, _OnObjectListLoadDataFnc);
	m_wndObjectList.SetEvent(WE_DBLCLICK, _OnObjectListDblClickFnc);

	m_wndObjectList.AddEvent(1, _T("Check All"), _OnObjectListCheckAllFnc);
	m_wndObjectList.AddEvent(2, _T("UnCheck All"), _OnObjectListUnCheckAllFnc);

	m_wndDeptList.SetEvent(WE_SELCHANGE, _OnDeptListSelectChangeFnc);
	m_wndDeptList.SetEvent(WE_LOADDATA, _OnDeptListLoadDataFnc);
	m_wndDeptList.SetEvent(WE_DBLCLICK, _OnDeptListDblClickFnc);

	m_wndDeptList.AddEvent(1, _T("Check All"), _OnDeptListCheckAllFnc);
	m_wndDeptList.AddEvent(2, _T("UnCheck All"), _OnDeptListUnCheckAllFnc);

	m_wndDischarge.SetEvent(WE_CLICK, _OnDischargeSelectFnc);
	m_wndIndischarge.SetEvent(WE_CLICK, _OnIndischargeSelectFnc);

	CString szSQL;
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);
	OnObjectListLoadData();
	OnDeptListLoadData();

	/*CString tmpStr;
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		tmpStr = m_wndObjectList.GetItemText(i, 2);

		if (tmpStr == _T("S"))
			m_wndObjectList.SetCheck(i, TRUE);
	}*/
	
}
void CFMFoodUnpaidPatientList::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	//DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndDischarge.GetDlgCtrlID(), m_bDischarge);
	//DDX_Radio(pDX, m_wndIndischarge.GetDlgCtrlID(), m_nIndischarge);
}
void CFMFoodUnpaidPatientList::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMFoodUnpaidPatientList::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMFoodUnpaidPatientList::SetDefaultValues()
{

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bDischarge=TRUE;
	m_nIndischarge=1;

}
int CFMFoodUnpaidPatientList::SetMode(int nMode){
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
/*void CFMFoodUnpaidPatientList::OnYearChange(){
	
} */
/*void CFMFoodUnpaidPatientList::OnYearSetfocus(){
	
} */
/*void CFMFoodUnpaidPatientList::OnYearKillfocus(){
	
} */
int CFMFoodUnpaidPatientList::OnYearCheckValue()
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
void CFMFoodUnpaidPatientList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMFoodUnpaidPatientList::OnReportPeriodSelendok()
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
/*void CFMFoodUnpaidPatientList::OnReportPeriodSetfocus(){
	
}*/
/*void CFMFoodUnpaidPatientList::OnReportPeriodKillfocus(){
	
}*/
long CFMFoodUnpaidPatientList::OnReportPeriodLoadData()
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
/*void CFMFoodUnpaidPatientList::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMFoodUnpaidPatientList::OnFromDateChange(){
	
} */
/*void CFMFoodUnpaidPatientList::OnFromDateSetfocus(){
	
} */
/*void CFMFoodUnpaidPatientList::OnFromDateKillfocus(){
	
} */
int CFMFoodUnpaidPatientList::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMFoodUnpaidPatientList::OnToDateChange(){
	
} */
/*void CFMFoodUnpaidPatientList::OnToDateSetfocus(){
	
} */
/*void CFMFoodUnpaidPatientList::OnToDateKillfocus(){
	
} */
int CFMFoodUnpaidPatientList::OnToDateCheckValue(){
	return 0;
} 
void CFMFoodUnpaidPatientList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMFoodUnpaidPatientList::OnClerkSelendok(){
	 
}
/*void CFMFoodUnpaidPatientList::OnClerkSetfocus(){
	
}*/
/*void CFMFoodUnpaidPatientList::OnClerkKillfocus(){
	
}*/
long CFMFoodUnpaidPatientList::OnClerkLoadData()
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
/*void CFMFoodUnpaidPatientList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMFoodUnpaidPatientList::OnPrintSelect()
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

	if (m_bDischarge)
	{
		if (!rpt.Init(_T("Reports/HMS/HF_DANHSACHANCUABENHNHANDARAVIENCHUATHANHTOAN.RPT")))
			return;
	}
	else
	{
		if (!rpt.Init(_T("Reports/HMS/HF_DANHSACHANCUABENHNHANCHUATHANHTOAN.RPT")))
			return;
	}

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
	long double nGroupTotal = 0;
	long double nTotal = 0;
	double nCost;
	int nIndex = 1;
	int nCount = 0;

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("deptid"), szNewLine);

		if (!szNewLine.IsEmpty() && szNewLine != szOldLine)
		{
			if (nGroupTotal > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				TranslateString(_T("Total Amount :"), tmpStr);

				szTemp.Format(_T("%s %d"), tmpStr, nCount);

				rptDetail->SetValue(_T("TotalGroup"), szTemp);
				FormatCurrency(nGroupTotal, tmpStr);
				rptDetail->SetValue(_T("TotalAmount"), tmpStr);
				nTotal += nGroupTotal;
				nGroupTotal = 0;
			}

			szOldLine = szNewLine;
			nCount = 0;
		}

		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("recordno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		if (nGroupTotal == 0)
		{
			rs.GetValue(_T("deptname"), tmpStr);
			rptDetail->SetValue(_T("4"), tmpStr);
		}

		rs.GetValue(_T("admitdate"), tmpStr);
		szTemp = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("5"), szTemp);

		rs.GetValue(_T("tienan"), nCost);
		nGroupTotal += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		nCount++;
		rs.MoveNext();
	}

	if (nGroupTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		TranslateString(_T("Total Amount :"), tmpStr);

		szTemp.Format(_T("%s %d"), tmpStr, nCount);
		rptDetail->SetValue(_T("TotalGroup"), szTemp);

		FormatCurrency(nGroupTotal, tmpStr);
		rptDetail->SetValue(_T("TotalAmount"), tmpStr);
		nTotal += nGroupTotal;
		nGroupTotal = 0;
	}

	if (nTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));

		TranslateString(_T("Total Amount :"), tmpStr);
		szTemp.Format(_T("%s %d"), tmpStr, nIndex - 1);

		rptDetail->SetValue(_T("TotalGroup"), szTemp);

		FormatCurrency(nTotal, tmpStr);
		rptDetail->SetValue(_T("TotalAmount"), tmpStr);
	}

	tmpStr = pMF->m_UserInfo.su_name;
	StringUpper(tmpStr, szTemp);
	rpt.GetReportFooter()->SetValue(_T("ReceiverBy"), szTemp);

	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);

	EndWaitCursor();
	rpt.PrintPreview();
} 
void CFMFoodUnpaidPatientList::OnExportSelect()
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
	xls.SetColumnWidth(2, 9);
	xls.SetColumnWidth(3, 25);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(5, 15);

	int nRow = 1;
	int nCol = 0;

	xls.SetRowHeight(6, 45);

	xls.SetCellMergedColumns(0, 1, 3);
	xls.SetCellMergedColumns(0, 2, 3);

	xls.SetCellMergedColumns(3, 1, 3);
	xls.SetCellMergedColumns(3, 2, 3);

	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(3, 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(3, 2, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol, nRow + 3, 6);
	xls.SetCellMergedColumns(nCol, nRow + 4, 6);

	if (m_bDischarge)
		tmpStr.Format(_T("\x110\xC3 R\x41 VI\x1EC6N"));

	szTemp.Format(_T("\x44\x41NH S\xC1\x43H \x102N \x43\x1EE6\x41 \x42\x1EC6NH NH\xC2N %s \x43H\x1AF\x41 TH\x41NH TO\xC1N"),
		          tmpStr);

	xls.SetCellText(nCol, nRow + 3, szTemp, FMT_TEXT | FMT_CENTER, true, 14, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);

	TranslateString(_T("Index"), tmpStr);
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Patient Name"), tmpStr);
	xls.SetCellText(nCol + 1, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Admission No"), tmpStr);
	xls.SetCellText(nCol + 2, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Dept"), tmpStr);
	xls.SetCellText(nCol + 3, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Admissdion Date"), tmpStr);
	xls.SetCellText(nCol + 4, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Amount"), tmpStr);
	xls.SetCellText(nCol + 5, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	int nIndex = 1;
	nRow = 6;
	CString szOldLine, szNewLine;
	long double nGroupTotal = 0;
	long double nTotal = 0;
	double nCost;
	int nCount = 0;

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("deptid"), szNewLine);

		if (!szNewLine.IsEmpty() && szNewLine != szOldLine)
		{
			if (nGroupTotal > 0)
			{
				nRow++;
				TranslateString(_T("Total Amount :"), tmpStr);

				szTemp.Format(_T("%s %d"), tmpStr, nCount);

				xls.SetCellText(nCol + 1, nRow, szTemp, FMT_TEXT | FMT_WRAPING, true);
				tmpStr.Format(_T("%.2lf"), nGroupTotal);
				xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
				nTotal += nGroupTotal;
				nGroupTotal = 0;
			}

			szOldLine = szNewLine;
			nCount = 0;
		}

		nRow++;

		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("recordno"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		if (nGroupTotal == 0)
		{
			rs.GetValue(_T("deptname"), tmpStr);
			xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);
		}

		rs.GetValue(_T("admitdate"), tmpStr);
		szTemp = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		xls.SetCellText(nCol + 4, nRow, szTemp, FMT_DATE | FMT_WRAPING);

		rs.GetValue(_T("tienan"), nCost);
		nGroupTotal += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		nCount++;
		rs.MoveNext();
	}

	if (nGroupTotal > 0)
	{
		nRow++;

		TranslateString(_T("Total Amount :"), tmpStr);

		szTemp.Format(_T("%s %d"), tmpStr, nCount);
		xls.SetCellText(nCol + 1, nRow, szTemp, FMT_TEXT | FMT_WRAPING, true);
		
		tmpStr.Format(_T("%.2lf"), nGroupTotal);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
		nTotal += nGroupTotal;
		nGroupTotal = 0;

	}

	if (nTotal > 0)
	{
		nRow++;
		TranslateString(_T("Total Amount :"), tmpStr);

		szTemp.Format(_T("%s %d"), tmpStr, nIndex - 1);
		xls.SetCellText(nCol + 1, nRow, szTemp, FMT_TEXT | FMT_WRAPING, true);
		
		tmpStr.Format(_T("%.2lf"), nTotal);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\DSAnBNChuaTT.xls"));
} 
void CFMFoodUnpaidPatientList::OnObjectListDblClick(){
	
} 
void CFMFoodUnpaidPatientList::OnObjectListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMFoodUnpaidPatientList::OnObjectListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMFoodUnpaidPatientList::OnObjectListLoadData()
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
void CFMFoodUnpaidPatientList::OnDeptListDblClick(){
	
} 
void CFMFoodUnpaidPatientList::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMFoodUnpaidPatientList::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMFoodUnpaidPatientList::OnDeptListLoadData()
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
void CFMFoodUnpaidPatientList::OnDischargeSelect(){
	
}
void CFMFoodUnpaidPatientList::OnIndischargeSelect(){
	
}
int CFMFoodUnpaidPatientList::OnAddFMFoodUnpaidPatientList(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMFoodUnpaidPatientList::OnEditFMFoodUnpaidPatientList(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMFoodUnpaidPatientList::OnDeleteFMFoodUnpaidPatientList(){
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
 		OnCancelFMFoodUnpaidPatientList();
 	}
	return 0;
}
int CFMFoodUnpaidPatientList::OnSaveFMFoodUnpaidPatientList(){
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
 		//OnFMFoodUnpaidPatientListListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CFMFoodUnpaidPatientList::OnCancelFMFoodUnpaidPatientList(){
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
int CFMFoodUnpaidPatientList::OnFMFoodUnpaidPatientListListLoadData(){
	return 0;
}

int CFMFoodUnpaidPatientList::OnDeptListCheckAll()
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

int CFMFoodUnpaidPatientList::OnDeptListUnCheckAll()
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


int CFMFoodUnpaidPatientList::OnObjectListCheckAll()
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

int CFMFoodUnpaidPatientList::OnObjectListUnCheckAll()
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

CString CFMFoodUnpaidPatientList::GetQueryString()
{
	CString szSQL;
	CString szWhere;
	CString szCondition;
	CString szDepts, szObjects;

	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (m_wndObjectList.GetCheck(i))
		{
			if (!szObjects.IsEmpty())
				szObjects += _T(",");
			szObjects.AppendFormat(_T("'%s'"), m_wndObjectList.GetItemText(i, 0));
		}
	}

	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			if (!szDepts.IsEmpty())
				szDepts += _T(",");
			szDepts.AppendFormat(_T("'%s'"), m_wndDeptList.GetItemText(i, 0));
		}
	}

	/*if (!m_szClerkKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and fi.hfe_staff='%s' "), m_szClerkKey);
	}*/

	if (!szObjects.IsEmpty())
	{
		szWhere.Format(_T(" and hd_object in(%s) "), szObjects);
	}

	if (m_bDischarge)
	{
		szWhere.AppendFormat(_T(" and hcr_status='T' "));
	}


	if (!szDepts.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and deptid in(%s) "), szDepts);
	}


	szSQL.Format(_T(" select hd_docno as docno,") \
				_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("        hcr_recordno as recordno,") \
				_T("        deptid,") \
				_T("        get_departmentname(deptid) as deptname,") \
				_T("        Cast_Timestamp2Date(hcr_admitdate) as admitdate,") \
				_T("        sum(tienan) as tienan") \
				_T(" from") \
				_T(" (") \
				_T("   select hfe_docno as docno,") \
				_T("          hfe_deptid as deptid,") \
				_T("          hfe_patdebt as tienan") \
				_T("   from hms_fee") \
				_T("   where hfe_class in('I') and hfe_status<>'P' and substr(hfe_group,1,2)='FF'") \
				_T(" ) Tbl") \
				_T(" left join hms_doc on(hd_docno=docno)") \
				_T(" left join hms_object on(ho_id=hd_object)") \
				_T(" left join hms_clinical_record on(hcr_docno=hd_docno)") \
				_T(" left join hms_patient on(hp_patientno=hd_patientno)") \
				_T(" where hcr_admitdate between cast_string2timestamp('%s')") \
				_T("       and cast_string2timestamp('%s') %s") \
				_T(" group by deptid, hd_docno, hcr_recordno, hp_surname, hp_midname, hp_firstname, hcr_admitdate") \
				_T(" order by deptid, hcr_admitdate"),
				m_szFromDate, m_szToDate, szWhere);

	return szSQL;
}
