#include "stdafx.h"
#include "FMExpenseInsuranceTreatmentCost.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMExpenseInsuranceTreatmentCost* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMExpenseInsuranceTreatmentCost* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMExpenseInsuranceTreatmentCost *pVw = (CFMExpenseInsuranceTreatmentCost *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMExpenseInsuranceTreatmentCost *pVw = (CFMExpenseInsuranceTreatmentCost *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnObjectListLoadDataFnc(CWnd *pWnd){
	return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnObjectListLoadData();
} 
static void _OnObjectListDblClickFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnObjectListDblClick();
} 
static void _OnObjectListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnObjectListSelectChange(nOldItem, nNewItem);
} 
static int _OnObjectListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnObjectListDeleteItem();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnDeptListDeleteItem();
} 
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnLockedSelect();
}
static int _OnAddFMExpenseInsuranceTreatmentCostFnc(CWnd *pWnd){
	 return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnAddFMExpenseInsuranceTreatmentCost();
} 
static int _OnEditFMExpenseInsuranceTreatmentCostFnc(CWnd *pWnd){
	 return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnEditFMExpenseInsuranceTreatmentCost();
} 
static int _OnDeleteFMExpenseInsuranceTreatmentCostFnc(CWnd *pWnd){
	 return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnDeleteFMExpenseInsuranceTreatmentCost();
} 
static int _OnSaveFMExpenseInsuranceTreatmentCostFnc(CWnd *pWnd){
	 return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnSaveFMExpenseInsuranceTreatmentCost();
} 
static int _OnCancelFMExpenseInsuranceTreatmentCostFnc(CWnd *pWnd){
	 return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnCancelFMExpenseInsuranceTreatmentCost();
}

static int _OnObjectListCheckAllFnc(CWnd *pWnd){
	return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnObjectListCheckAll();
}
static int _OnObjectListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnObjectListUnCheckAll();
}
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMExpenseInsuranceTreatmentCost*)pWnd)->OnDeptListUnCheckAll();
}

CFMExpenseInsuranceTreatmentCost::CFMExpenseInsuranceTreatmentCost(CWnd* pParent)
{
	m_nDlgWidth = 450;
	m_nDlgHeight = 590;
	SetDefaultValues();
}
CFMExpenseInsuranceTreatmentCost::~CFMExpenseInsuranceTreatmentCost(){
}
void CFMExpenseInsuranceTreatmentCost::OnCreateComponents()
{
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 440, 550);
	m_wndObjectInfo.Create(this, _T("Object"), 10, 120, 435, 265);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 270, 435, 515);
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
	m_wndPrint.Create(this, _T("&Print"), 235, 555, 335, 580);
	m_wndExport.Create(this, _T("&Export"), 340, 555, 440, 580);
	m_wndObjectList.Create(this,15, 145, 430, 260); 
	m_wndDeptList.Create(this,15, 295, 430, 510); 
	m_wndLocked.Create(this, _T("Locked"), 335, 520, 435, 545);

}
void CFMExpenseInsuranceTreatmentCost::OnInitializeComponents()
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


	m_wndObjectList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndObjectList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndObjectList.InsertColumn(2, _T("Type"), CFMT_TEXT, 0);
	m_wndObjectList.SetCheckBox(TRUE);


	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.SetCheckBox(TRUE);
}
void CFMExpenseInsuranceTreatmentCost::OnSetWindowEvents()
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

	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);
	
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

	CString tmpStr;
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		tmpStr = m_wndObjectList.GetItemText(i, 2);
		
		if (tmpStr != _T("S"))
			m_wndObjectList.SetCheck(i, TRUE);
	}
}
void CFMExpenseInsuranceTreatmentCost::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);

}
void CFMExpenseInsuranceTreatmentCost::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMExpenseInsuranceTreatmentCost::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMExpenseInsuranceTreatmentCost::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bLocked=FALSE;

}
int CFMExpenseInsuranceTreatmentCost::SetMode(int nMode){
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
/*void CFMExpenseInsuranceTreatmentCost::OnYearChange(){
	
} */
/*void CFMExpenseInsuranceTreatmentCost::OnYearSetfocus(){
	
} */
/*void CFMExpenseInsuranceTreatmentCost::OnYearKillfocus(){
	
} */
int CFMExpenseInsuranceTreatmentCost::OnYearCheckValue()
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
void CFMExpenseInsuranceTreatmentCost::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMExpenseInsuranceTreatmentCost::OnReportPeriodSelendok()
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
/*void CFMExpenseInsuranceTreatmentCost::OnReportPeriodSetfocus(){
	
}*/
/*void CFMExpenseInsuranceTreatmentCost::OnReportPeriodKillfocus(){
	
}*/
long CFMExpenseInsuranceTreatmentCost::OnReportPeriodLoadData()
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
/*void CFMExpenseInsuranceTreatmentCost::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMExpenseInsuranceTreatmentCost::OnFromDateChange(){
	
} */
/*void CFMExpenseInsuranceTreatmentCost::OnFromDateSetfocus(){
	
} */
/*void CFMExpenseInsuranceTreatmentCost::OnFromDateKillfocus(){
	
} */
int CFMExpenseInsuranceTreatmentCost::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMExpenseInsuranceTreatmentCost::OnToDateChange(){
	
} */
/*void CFMExpenseInsuranceTreatmentCost::OnToDateSetfocus(){
	
} */
/*void CFMExpenseInsuranceTreatmentCost::OnToDateKillfocus(){
	
} */
int CFMExpenseInsuranceTreatmentCost::OnToDateCheckValue(){
	return 0;
} 
void CFMExpenseInsuranceTreatmentCost::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMExpenseInsuranceTreatmentCost::OnClerkSelendok(){
	 
}
/*void CFMExpenseInsuranceTreatmentCost::OnClerkSetfocus(){
	
}*/
/*void CFMExpenseInsuranceTreatmentCost::OnClerkKillfocus(){
	
}*/
long CFMExpenseInsuranceTreatmentCost::OnClerkLoadData()
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
/*void CFMExpenseInsuranceTreatmentCost::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMExpenseInsuranceTreatmentCost::OnPrintSelect()
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

	if (!rpt.Init(_T("Reports/HMS/HF_BANGKETHANHTOAN_MAU9A.RPT")))
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
	long double nGroupTotal[20];
	long double nTotal[20];
	double nCost;
	int nIndex = 1;

	for (int i = 0; i < 20; i++)
	{
		nGroupTotal[i] = 0;
		nTotal[i] = 0;
	}

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("deptid"), szNewLine);

		if (!szNewLine.IsEmpty() && szNewLine != szOldLine)
		{
			if (nGroupTotal[19] > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				//TranslateString(_T("\x43\x1ED9ng kho\x61"), tmpStr);
				rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng kho\x61"));

				for (int i = 7; i < 20; i++)
				{
					FormatCurrency(nGroupTotal[i], tmpStr);
					szTemp.Format(_T("s%d"), i + 1);
					rptDetail->SetValue(szTemp, tmpStr);
					nTotal[i] += nGroupTotal[i];
					nGroupTotal[i] = 0;
				}
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rs.GetValue(_T("deptname"), tmpStr);
			rptDetail->SetValue(_T("GroupName"), tmpStr);
			//rs.GetValue(_T("objid"), szObject);
			szOldLine = szNewLine;
		}

		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("objname"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("recordno"), tmpStr);
		rptDetail->SetValue(_T("4"), szTemp);

		rs.GetValue(_T("workplace"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("deptid"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("ngaydt"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("trichtgui"), nCost);
		nGroupTotal[7] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("vienphi"), nCost);
		nGroupTotal[8] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("bugao"), nCost);
		nGroupTotal[9] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.GetValue(_T("xnktc"), nCost);
		nGroupTotal[10] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.GetValue(_T("dc"), nCost);
		nGroupTotal[11] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);

		rs.GetValue(_T("cong"), nCost);
		nGroupTotal[12] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);

		rs.GetValue(_T("conlai"), nCost);
		nGroupTotal[13] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("14"), tmpStr);

		rs.GetValue(_T("chitra"), nCost);
		nGroupTotal[14] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("15"), tmpStr);

		rs.GetValue(_T("chiluong"), nCost);
		nGroupTotal[15] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("16"), tmpStr);

		rs.GetValue(_T("temtauxe"), nCost);
		nGroupTotal[16] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("17"), tmpStr);

		rs.GetValue(_T("letet"), nCost);
		nGroupTotal[17] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("18"), tmpStr);

		rs.GetValue(_T("chikhac"), nCost);
		nGroupTotal[18] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("19"), tmpStr);

		rs.GetValue(_T("congchi"), nCost);
		nGroupTotal[19] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("20"), tmpStr);

		rs.MoveNext();
	}

	if (nGroupTotal[19] > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		//TranslateString(_T("\x43\x1ED9ng kho\x61"), tmpStr);
		rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng kho\x61"));

		for (int i = 7; i < 20; i++)
		{
			FormatCurrency(nGroupTotal[i], tmpStr);
			szTemp.Format(_T("s%d"), i + 1);
			rptDetail->SetValue(szTemp, tmpStr);
			nTotal[i] += nGroupTotal[i];
			nGroupTotal[i] = 0;
		}
	}

	long double nTemp;

	if (nTotal[19] > 0)
	{
		m_nTotal = floor(nTotal[19] + 0.5);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		//TranslateString(_T("T\x1ED5ng \x63\x1ED9ng"), tmpStr);
		rptDetail->SetValue(_T("TotalGroup"), _T("T\x1ED5ng \x63\x1ED9ng"));
		for (int i = 7; i < 20; i++)
		{
			nTemp = floor(nTotal[i] + 0.5);
			FormatCurrency(nTemp, tmpStr);
			szTemp.Format(_T("s%d"), i + 1);
			rptDetail->SetValue(szTemp, tmpStr);
		}
	}

	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);

	EndWaitCursor();
	rpt.PrintPreview();

	OnPrintExpenseInvoice();
} 
void CFMExpenseInsuranceTreatmentCost::OnExportSelect()
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
	xls.SetColumnWidth(2, 11);
	xls.SetColumnWidth(3, 8);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 7);
	xls.SetColumnWidth(6, 5);

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

	int nRow = 1;
	int nCol = 0;

	xls.SetRowHeight(6, 30);
	xls.SetRowHeight(7, 30);

	xls.SetCellMergedColumns(0, 1, 4);
	xls.SetCellMergedColumns(0, 2, 4);

	xls.SetCellMergedColumns(16, 1, 4);
	xls.SetCellMergedColumns(16, 2, 4);

	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(16, 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(16, 2, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol, nRow + 3, 20);
	xls.SetCellMergedColumns(nCol, nRow + 4, 20);

	xls.SetCellText(nCol, nRow + 3, _T("\x42\x1EA2NG K\xCA TH\x41NH TO\xC1N \x43HI \x43\x1EE6\x41 \x42\x1EC6NH NH\xC2N \x42\x110 - \x43S - \x42H"),
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

	TranslateString(_T("Rank"), tmpStr);
	xls.SetCellMergedRows(nCol + 2, nRow + 5, 2);
	xls.SetCellText(nCol + 2, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Admission No"), tmpStr);
	xls.SetCellMergedRows(nCol + 3, nRow + 5, 2);
	xls.SetCellText(nCol + 3, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Unit"), tmpStr);
	xls.SetCellMergedRows(nCol + 4, nRow + 5, 2);
	xls.SetCellText(nCol + 4, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Dept"), tmpStr);
	xls.SetCellMergedRows(nCol + 5, nRow + 5, 2);
	xls.SetCellText(nCol + 5, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Days"), tmpStr);
	xls.SetCellMergedRows(nCol + 6, nRow + 5, 2);
	xls.SetCellText(nCol + 6, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 7, nRow + 5, 2);
	xls.SetCellText(nCol + 7, nRow + 5, _T("Tr\xED\x63h t\x1EA1m g\x1EEDi"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedColumns(nCol + 8, nRow + 5, 5);
	xls.SetCellText(nCol + 8, nRow + 5, _T("Thu vi\x1EC7n ph\xED"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 8, nRow + 6, _T("Vi\x1EC7n ph\xED"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 9, nRow + 6, _T("\x42\xF9 g\x1EA1o"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 10, nRow + 6, _T("\x58N-KT\x43"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 11, nRow + 6, _T("\x44\x43"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 12, nRow + 6, _T("\x43\x1ED9ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 13, nRow + 5, 2);
	xls.SetCellText(nCol + 13, nRow + 5, _T("\x43\xF2n l\x1EA1i"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 14, nRow + 5, 2);
	xls.SetCellText(nCol + 14, nRow + 5, _T("\x43hi tr\x1EA3 \x63hi ph\xED"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 15, nRow + 5, 2);
	xls.SetCellText(nCol + 15, nRow + 5, _T("\x43hi l\x1B0\x1A1ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 16, nRow + 5, 2);
	xls.SetCellText(nCol + 16, nRow + 5, _T("T\x65m t\xE0u \x78\x65"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 17, nRow + 5, 2);
	xls.SetCellText(nCol + 17, nRow + 5, _T("L\x1EC5 T\x1EBFt"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 18, nRow + 5, 2);
	xls.SetCellText(nCol + 18, nRow + 5, _T("\x43hi kh\xE1\x63"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 19, nRow + 5, 2);
	xls.SetCellText(nCol + 19, nRow + 5, _T("\x43\x1ED9ng \x63hi"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	int nIndex = 1;
	nRow = 7;
	CString szOldLine, szNewLine;
	long double nGroupTotal[20];
	long double nTotal[20];
	double nCost;

	for (int i = 0; i < 20; i++)
	{
		nGroupTotal[i] = 0;
		nTotal[i] = 0;
	}

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("deptid"), szNewLine);

		if (!szNewLine.IsEmpty() && szNewLine != szOldLine)
		{
			if (nGroupTotal[19] > 0)
			{
				nRow++;
				xls.SetCellText(nCol + 1, nRow, _T("\x43\x1ED9ng kho\x61"), FMT_TEXT | FMT_WRAPING, true);
				for (int i = 7; i < 20; i++)
				{
					tmpStr.Format(_T("%.2lf"), nGroupTotal[i]);
					xls.SetCellText(nCol + i, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
					nTotal[i] += nGroupTotal[i];
					nGroupTotal[i] = 0;
				}
			}

			nRow++;
			rs.GetValue(_T("deptname"), tmpStr);
			xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING, true); 
			szOldLine = szNewLine;
		}

		nRow++;

		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("objname"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("recordno"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("workplace"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("deptid"), tmpStr);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("ngaydt"), tmpStr);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("trichtgui"), nCost);
		nGroupTotal[7] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("vienphi"), nCost);
		nGroupTotal[8] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("bugao"), nCost);
		nGroupTotal[9] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 9, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("xnktc"), nCost);
		nGroupTotal[10] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 10, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("dc"), nCost);
		nGroupTotal[11] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 11, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("cong"), nCost);
		nGroupTotal[12] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 12, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("conlai"), nCost);
		nGroupTotal[13] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 13, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("chitra"), nCost);
		nGroupTotal[14] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 14, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("chiluong"), nCost);
		nGroupTotal[15] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 15, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("temtauxe"), nCost);
		nGroupTotal[16] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 16, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("letet"), nCost);
		nGroupTotal[17] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 17, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("chikhac"), nCost);
		nGroupTotal[18] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 18, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("congchi"), nCost);
		nGroupTotal[19] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 19, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.MoveNext();
	}

	if (nGroupTotal[19] > 0)
	{
		nRow++;
		xls.SetCellText(nCol + 1, nRow, _T("\x43\x1ED9ng kho\x61"), FMT_TEXT | FMT_WRAPING, true);
		for (int i = 7; i < 20; i++)
		{
			tmpStr.Format(_T("%.2lf"), nGroupTotal[i]);
			xls.SetCellText(nCol + i, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
			nTotal[i] += nGroupTotal[i];
			nGroupTotal[i] = 0;
		}
	}

	if (nTotal[19] > 0)
	{
		nRow++;
		xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_WRAPING, true);
		for (int i = 7; i < 20; i++)
		{
			tmpStr.Format(_T("%.2lf"), nTotal[i]);
			xls.SetCellText(nCol + i, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
		}
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\BangKeTTChi.xls"));
} 
void CFMExpenseInsuranceTreatmentCost::OnObjectListDblClick(){
	
} 
void CFMExpenseInsuranceTreatmentCost::OnObjectListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMExpenseInsuranceTreatmentCost::OnObjectListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMExpenseInsuranceTreatmentCost::OnObjectListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	CString szWhere;
	m_wndObjectList.BeginLoad(); 
	int nCount = 0;

	szWhere.Format(_T(" AND ho_type NOT IN('S') "));

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
}
void CFMExpenseInsuranceTreatmentCost::OnDeptListDblClick(){
	
} 
void CFMExpenseInsuranceTreatmentCost::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMExpenseInsuranceTreatmentCost::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMExpenseInsuranceTreatmentCost::OnDeptListLoadData()
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
void CFMExpenseInsuranceTreatmentCost::OnLockedSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
int CFMExpenseInsuranceTreatmentCost::OnAddFMExpenseInsuranceTreatmentCost(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMExpenseInsuranceTreatmentCost::OnEditFMExpenseInsuranceTreatmentCost(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMExpenseInsuranceTreatmentCost::OnDeleteFMExpenseInsuranceTreatmentCost(){
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
 		OnCancelFMExpenseInsuranceTreatmentCost();
 	}
	return 0;
}
int CFMExpenseInsuranceTreatmentCost::OnSaveFMExpenseInsuranceTreatmentCost(){
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
 		//OnFMExpenseInsuranceTreatmentCostListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CFMExpenseInsuranceTreatmentCost::OnCancelFMExpenseInsuranceTreatmentCost(){
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
int CFMExpenseInsuranceTreatmentCost::OnFMExpenseInsuranceTreatmentCostListLoadData(){
	return 0;
}

int CFMExpenseInsuranceTreatmentCost::OnDeptListCheckAll()
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

int CFMExpenseInsuranceTreatmentCost::OnDeptListUnCheckAll()
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


int CFMExpenseInsuranceTreatmentCost::OnObjectListCheckAll()
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

int CFMExpenseInsuranceTreatmentCost::OnObjectListUnCheckAll()
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

CString CFMExpenseInsuranceTreatmentCost::GetQueryString()
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

	if (!m_szClerkKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and fi.hfe_staff='%s' "), m_szClerkKey);
	}

	if (!szObjects.IsEmpty())
	{
		szCondition.Format(_T(" and hd_object in(%s) "), szObjects);
	}

	if (!szDepts.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and fi.hfe_deptid in(%s) "), szDepts);
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

	szSQL.Format(_T(" select trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("        sys_sel_getname('hms_rank', hp_rank) as rank,") \
				_T("        get_objectname(ho_id) as objname,") \
				_T("        hcr_recordno as recordno,") \
				_T("        hp_workplace as workplace,") \
				_T("        deptid,") \
				_T("        hcr_sumtreat as ngaydt,") \
				_T("        get_departmentname(deptid) as deptname,") \
				_T("        trichtgui,") \
				_T("        vienphi,") \
				_T("        bugao,") \
				_T("        xnktc,") \
				_T("        dc,") \
				_T("        cong,") \
				_T("        conlai,") \
				_T("        chitra,") \
				_T("        chiluong,") \
				_T("        temtauxe,") \
				_T("        letet,") \
				_T("        chikhac,") \
				_T("        conlai+chiluong+temtauxe+letet+chikhac as congchi") \
				_T(" from") \
				_T(" (") \
				_T("  select docno,") \
				_T("         recvdate,") \
				_T("         invoiceno,") \
				_T("         deptid,") \
				_T("         trichtgui as trichtgui,") \
				_T("         sum(vienphi) as vienphi,") \
				_T("         sum(bugao) as bugao,") \
				_T("         sum(xnktc) as xnktc,") \
				_T("         sum(dc) as dc,") \
				_T("         sum(vienphi+bugao+xnktc+dc) as cong,") \
				_T("         trichtgui-sum(vienphi+bugao+xnktc+dc) as conlai,") \
				_T("         sum(chitra) as chitra,") \
				_T("         (select sum(hfe_salary_amount)") \
				_T("          from hms_fee_sold") \
				_T("          where hfe_invoiceno=invoiceno and hfe_docno=docno) as chiluong,") \
				_T("         (select sum(hfe_train_amount+hfe_stamp_amount)") \
				_T("          from hms_fee_sold") \
				_T("          where hfe_invoiceno=invoiceno and hfe_docno=docno) as temtauxe,") \
				_T("         (select sum(hfe_holiday_amount)") \
				_T("          from hms_fee_sold") \
				_T("          where hfe_invoiceno=invoiceno and hfe_docno=docno) as letet,") \
				_T("         (select sum(hfe_other_amount)") \
				_T("          from hms_fee_sold") \
				_T("          where hfe_invoiceno=invoiceno and hfe_docno=docno) as chikhac") \
				_T("  from") \
				_T("  (") \
				_T("    select fi.hfe_docno as docno,") \
				_T("           fi.hfe_invoiceno as invoiceno,") \
				_T("           Cast_Timestamp2Date(fi.hfe_date) as recvdate,") \
				_T("           fi.hfe_deptid as deptid,") \
				_T("           fi.hfe_deposit as trichtgui,") \
				_T("           fe.hfe_patpaid as vienphi,") \
				_T("           0 as bugao,") \
				_T("           case when substr(fe.hfe_group,1,2) in('B1','B2','B3') and fe.hfe_hitech='Y'") \
				_T("           then fe.hfe_patpaid else 0 end as xnktc,") \
				_T("           0 as dc,") \
				_T("           0 as chitra,") \
				_T("           0 as chiluong,") \
				_T("           0 as temtauxe,") \
				_T("           0 as letet,") \
				_T("           0 as chikhac") \
				_T("    from hms_fee fe") \
				_T("    left join hms_fee_invoice fi on(fi.hfe_invoiceno=fe.hfe_invoiceno and fe.hfe_docno=fi.hfe_docno)") \
				_T("    where fi.hfe_class in('I') and fi.hfe_status='P' %s") \
				_T("  ) Tbl") \
				_T("  group by docno, recvdate, invoiceno, deptid, trichtgui") \
				_T(" ) Tbl1") \
				_T(" left join hms_doc on(hd_docno=docno)") \
				_T(" left join hms_patient on(hp_patientno=hd_patientno)") \
				_T(" left join hms_clinical_record on(hcr_docno=hd_docno)") \
				_T(" left join hms_object on(ho_id=hd_object)") \
				_T(" where conlai+chiluong+temtauxe+letet+chikhac>0 %s") \
				_T(" order by deptid, pname"),
				szWhere, szCondition);_fmsg(_T("%s"), szSQL);

	return szSQL;
}


int CFMExpenseInsuranceTreatmentCost::OnPrintExpenseInvoice()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	UpdateData(true);

	CReport rpt;
	CRecord rs(&pMF->m_db);
	CRecord rss(&pMF->m_db);
	CString szSQL, tmpStr, szSysDate, szTemp;

	int nIdx = 0;
	double nAmount = 0;
	long double nTotal = 0;

	if (m_szClerkKey.IsEmpty())
		return -1;

	if (!rpt.Init(_T("Reports/HMS/HF_PHIEUCHI30BB.rpt"), false, false, 1))
		return 0;

	CReportSection *rptDetail;

	rptDetail = rpt.AddDetail();

	rptDetail->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rptDetail->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);

	tmpStr.Format(rptDetail->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rptDetail->SetValue(_T("ReportDate"), tmpStr);

	if (!m_szClerkKey.IsEmpty())
	{
		szSQL.Format(_T(" SELECT su_name AS staff,") \
				_T("        sd_name AS deptname") \
				_T(" FROM sys_user") \
				_T(" LEFT JOIN sys_dept ON(sd_id=su_deptid)") \
				_T(" WHERE su_userid='%s'"), m_szClerkKey);

		rs.ExecSQL(szSQL);

		if (!rs.IsEOF())
		{
			rs.GetValue(_T("deptname"), tmpStr);
			rptDetail->SetValue(_T("address"), tmpStr);

			rs.GetValue(_T("staff"), tmpStr);
			StringUpper(tmpStr, szTemp);
			rptDetail->SetValue(_T("ReceiveBy"), szTemp);
		}
	}
	/*else
	{
		rptDetail->SetValue(_T("address"), pMF->GetDepartmentName(pMF->GetCurrentDepartmentID()));
		tmpStr = pMF->GetCurrentUser();
	}*/

	/*while (!rs.IsEOF())
	{
		rs.GetValue(_T("amount"), nAmount);
		nTotal += nAmount;
		rs.MoveNext();
	}*/

	//rptDetail->SetValue(_T("patientname"), tmpStr);

	rptDetail->SetValue(_T("patientname"), _T("H\x1ED8 \x42\x110 - \x43S - \x42H (T\x1ED4NG H\x1EE2P \x43HI)"));
	rptDetail->SetValue(_T("reason"), _T("\x43hi ti\x1EC1n \x62\x1EC7nh nh\xE2n th\x61nh to\xE1n vi\x1EC7n ph\xED"));
	
	//nTotal += 0.5;
	//long double nTemp = floor(nTotal);

	FormatCurrency(m_nTotal, tmpStr);
	rptDetail->SetValue(_T("TotalAmount"), tmpStr);

	szTemp.Format(_T("%.0Lf"), m_nTotal);
	MoneyToString(szTemp, tmpStr);

	rptDetail->SetValue(_T("SumInWord"), tmpStr);

	CDateTime dt;
	dt.ParseDateTime(m_szToDate);
	CDate date = dt.GetDate();
	szSysDate = date.GetDate();

	tmpStr.Format(rptDetail->GetValue(_T("PrintDate")),
		          szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));

	rptDetail->SetValue(_T("PrintDate"), tmpStr);

	rpt.PrintPreview();

	return 1;
}