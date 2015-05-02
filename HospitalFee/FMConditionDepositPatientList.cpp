#include "stdafx.h"
#include "FMConditionDepositPatientList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMConditionDepositPatientList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMConditionDepositPatientList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMConditionDepositPatientList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMConditionDepositPatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMConditionDepositPatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMConditionDepositPatientList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMConditionDepositPatientList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMConditionDepositPatientList *pVw = (CFMConditionDepositPatientList *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMConditionDepositPatientList *pVw = (CFMConditionDepositPatientList *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnObjectListLoadDataFnc(CWnd *pWnd){
	return ((CFMConditionDepositPatientList*)pWnd)->OnObjectListLoadData();
} 
static void _OnObjectListDblClickFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList*)pWnd)->OnObjectListDblClick();
} 
static void _OnObjectListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMConditionDepositPatientList*)pWnd)->OnObjectListSelectChange(nOldItem, nNewItem);
} 
static int _OnObjectListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMConditionDepositPatientList*)pWnd)->OnObjectListDeleteItem();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMConditionDepositPatientList*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMConditionDepositPatientList*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMConditionDepositPatientList*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMConditionDepositPatientList*)pWnd)->OnDeptListDeleteItem();
} 
static void _OnDischargeSelectFnc(CWnd *pWnd){
	 ((CFMConditionDepositPatientList*)pWnd)->OnDischargeSelect();
}
static void _OnIndischargeSelectFnc(CWnd *pWnd){
	 ((CFMConditionDepositPatientList*)pWnd)->OnIndischargeSelect();
}
static void _OnPaidedSelectFnc(CWnd *pWnd){
	 ((CFMConditionDepositPatientList*)pWnd)->OnPaidedSelect();
}
static void _OnUnpaidedSelectFnc(CWnd *pWnd){
	 ((CFMConditionDepositPatientList*)pWnd)->OnUnpaidedSelect();
}
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMConditionDepositPatientList*)pWnd)->OnLockedSelect();
}
static int _OnAddFMConditionDepositPatientListFnc(CWnd *pWnd){
	 return ((CFMConditionDepositPatientList*)pWnd)->OnAddFMConditionDepositPatientList();
} 
static int _OnEditFMConditionDepositPatientListFnc(CWnd *pWnd){
	 return ((CFMConditionDepositPatientList*)pWnd)->OnEditFMConditionDepositPatientList();
} 
static int _OnDeleteFMConditionDepositPatientListFnc(CWnd *pWnd){
	 return ((CFMConditionDepositPatientList*)pWnd)->OnDeleteFMConditionDepositPatientList();
} 
static int _OnSaveFMConditionDepositPatientListFnc(CWnd *pWnd){
	 return ((CFMConditionDepositPatientList*)pWnd)->OnSaveFMConditionDepositPatientList();
} 
static int _OnCancelFMConditionDepositPatientListFnc(CWnd *pWnd){
	 return ((CFMConditionDepositPatientList*)pWnd)->OnCancelFMConditionDepositPatientList();
}

static int _OnObjectListCheckAllFnc(CWnd *pWnd){
	return ((CFMConditionDepositPatientList*)pWnd)->OnObjectListCheckAll();
}
static int _OnObjectListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMConditionDepositPatientList*)pWnd)->OnObjectListUnCheckAll();
}
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMConditionDepositPatientList*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMConditionDepositPatientList*)pWnd)->OnDeptListUnCheckAll();
}

CFMConditionDepositPatientList::CFMConditionDepositPatientList(CWnd* pParent)
{
	m_nDlgWidth = 450;
	m_nDlgHeight = 590;
	SetDefaultValues();
}
CFMConditionDepositPatientList::~CFMConditionDepositPatientList(){
}
void CFMConditionDepositPatientList::OnCreateComponents()
{
	/*m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 440, 575);
	m_wndObjectInfo.Create(this, _T("Object"), 10, 120, 435, 275);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 280, 435, 540);
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
	m_wndPrint.Create(this, _T("&Print"), 235, 580, 335, 605);
	m_wndExport.Create(this, _T("&Export"), 340, 580, 440, 605);
	m_wndObjectList.Create(this,15, 145, 430, 270); 
	m_wndDeptList.Create(this,15, 305, 430, 535); 
	m_wndDischarge.Create(this, _T("Discharge"), 230, 545, 330, 570);
	m_wndIndischarge.Create(this, _T("Indischarge"), 335, 545, 435, 570);
	m_wndPaided.Create(this, _T("Paided"), 10, 545, 115, 570);
	m_wndUnpaided.Create(this, _T("Unpaided"), 120, 545, 225, 570);*/

	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 440, 550);
	m_wndObjectInfo.Create(this, _T("Object"), 10, 120, 435, 255);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 260, 435, 485);
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
	m_wndPrint.Create(this, _T("&Print"), 235, 555, 335, 580);
	m_wndExport.Create(this, _T("&Export"), 340, 555, 440, 580);
	m_wndObjectList.Create(this,15, 145, 430, 250); 
	m_wndDeptList.Create(this,15, 285, 430, 480); 
	m_wndDischarge.Create(this, _T("Discharge"), 240, 490, 330, 515);
	m_wndIndischarge.Create(this, _T("Indischarge"), 335, 490, 435, 515);
	m_wndPaided.Create(this, _T("Paided"), 10, 490, 115, 515);
	m_wndUnpaided.Create(this, _T("Unpaided"), 120, 490, 235, 515);
	m_wndLocked.Create(this, _T("Locked"), 335, 520, 435, 545);
}
void CFMConditionDepositPatientList::OnInitializeComponents()
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
void CFMConditionDepositPatientList::OnSetWindowEvents()
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
	m_wndPaided.SetEvent(WE_CLICK, _OnPaidedSelectFnc);
	m_wndUnpaided.SetEvent(WE_CLICK, _OnUnpaidedSelectFnc);
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

	/*CString tmpStr;
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		tmpStr = m_wndObjectList.GetItemText(i, 2);

		if (tmpStr == _T("S"))
			m_wndObjectList.SetCheck(i, TRUE);
	}*/
}
void CFMConditionDepositPatientList::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndDischarge.GetDlgCtrlID(), m_bDischarge);
	DDX_Check(pDX, m_wndIndischarge.GetDlgCtrlID(), m_bIndischarge);
	DDX_Check(pDX, m_wndPaided.GetDlgCtrlID(), m_bPaided);
	DDX_Check(pDX, m_wndUnpaided.GetDlgCtrlID(), m_bUnpaided);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);

}
void CFMConditionDepositPatientList::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMConditionDepositPatientList::GetScreenToData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMConditionDepositPatientList::SetDefaultValues()
{
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bDischarge=FALSE;
	m_bIndischarge=FALSE;
	m_bPaided=FALSE;
	m_bUnpaided=TRUE;
	m_bLocked=FALSE;

}
int CFMConditionDepositPatientList::SetMode(int nMode)
{
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
/*void CFMConditionDepositPatientList::OnYearChange(){
	
} */
/*void CFMConditionDepositPatientList::OnYearSetfocus(){
	
} */
/*void CFMConditionDepositPatientList::OnYearKillfocus(){
	
} */
int CFMConditionDepositPatientList::OnYearCheckValue()
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
void CFMConditionDepositPatientList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMConditionDepositPatientList::OnReportPeriodSelendok()
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
/*void CFMConditionDepositPatientList::OnReportPeriodSetfocus(){
	
}*/
/*void CFMConditionDepositPatientList::OnReportPeriodKillfocus(){
	
}*/
long CFMConditionDepositPatientList::OnReportPeriodLoadData()
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
/*void CFMConditionDepositPatientList::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMConditionDepositPatientList::OnFromDateChange(){
	
} */
/*void CFMConditionDepositPatientList::OnFromDateSetfocus(){
	
} */
/*void CFMConditionDepositPatientList::OnFromDateKillfocus(){
	
} */
int CFMConditionDepositPatientList::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMConditionDepositPatientList::OnToDateChange(){
	
} */
/*void CFMConditionDepositPatientList::OnToDateSetfocus(){
	
} */
/*void CFMConditionDepositPatientList::OnToDateKillfocus(){
	
} */
int CFMConditionDepositPatientList::OnToDateCheckValue(){
	return 0;
} 
void CFMConditionDepositPatientList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMConditionDepositPatientList::OnClerkSelendok(){
	 
}
/*void CFMConditionDepositPatientList::OnClerkSetfocus(){
	
}*/
/*void CFMConditionDepositPatientList::OnClerkKillfocus(){
	
}*/
long CFMConditionDepositPatientList::OnClerkLoadData()
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
/*void CFMConditionDepositPatientList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMConditionDepositPatientList::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);

	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp;
	CString szSysDate;
	CString szObjectNames;

	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);


	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}


	if (m_bUnpaided)
	{
		if (!rpt.Init(_T("Reports/HMS/HF_DANHSACHBENHNHANTAMGUITIENVIENPHICHUATHANHTOAN.RPT")))
			return;
	}
	else if (m_bDischarge)
	{
		if (!rpt.Init(_T("Reports/HMS/HF_DANHSACHBENHNHANTAMGUICHUARAVIEN.RPT")))
			return;
	}
	else
	{
		if (!rpt.Init(_T("Reports/HMS/HF_DANHSACHBENHNHANTAMGUITIENVIENPHI.RPT")))
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

	szObjectNames.Empty();
	CStringArray strArr;
	bool bCheckServ = false, bCheckIns = false;


	strArr.Add(_T("\x64\x1ECB\x63h v\x1EE5"));
	strArr.Add(_T("\x42\x110 - \x43S - \x42H"));


	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (m_wndObjectList.GetCheck(i))
		{
			tmpStr = m_wndObjectList.GetItemText(i, 2);

			if (tmpStr == _T("S"))
				bCheckServ = true;
			else
				bCheckIns = true;
		}
	}

	if (bCheckServ || bCheckIns)
	{
		if (bCheckServ)
			szObjectNames.AppendFormat(_T("%s"), strArr[0]);

		if (bCheckIns)
		{
			if (!szObjectNames.IsEmpty())
				szObjectNames += _T(", ");
			szObjectNames.AppendFormat(_T("%s"), strArr[1]);
		}

		TranslateString(_T("Object"), szTemp);
		tmpStr.Format(_T("%s %s"), szTemp, szObjectNames);
		rpt.GetReportHeader()->SetValue(_T("Object"), tmpStr);
	}

	if (!bCheckIns && !bCheckServ)
		rpt.GetReportHeader()->SetValue(_T("Object"), _T("T\x1EA5t \x63\x1EA3 \x63\xE1\x63 \x111\x1ED1i t\x1B0\x1EE3ng"));


	while (!rs.IsEOF())
	{
		rs.GetValue(_T("deptid"), szNewLine);

		if (!szNewLine.IsEmpty() && szNewLine != szOldLine)
		{
			if (nGroupTotal > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				//TranslateString(_T("\x43\x1ED9ng kho\x61"), tmpStr);
				rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng kho\x61"));
				FormatCurrency(nGroupTotal, tmpStr);
				rptDetail->SetValue(_T("TotalAmount"), tmpStr);
				nTotal += nGroupTotal;
				nGroupTotal = 0;
			}
			/*rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rs.GetValue(_T("objname"), tmpStr);
			rptDetail->SetValue(_T("GroupName"), tmpStr);
			rs.GetValue(_T("objid"), szObject);*/
			szOldLine = szNewLine;
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

		rs.GetValue(_T("receivedate"), tmpStr);
		szTemp = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("5"), szTemp);

		rs.GetValue(_T("amount"), nCost);
		nGroupTotal += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.MoveNext();
	}

	if (nGroupTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		//TranslateString(_T("\x43\x1ED9ng kho\x61"), tmpStr);
		rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng kho\x61"));
		FormatCurrency(nGroupTotal, tmpStr);
		rptDetail->SetValue(_T("TotalAmount"), tmpStr);
		nTotal += nGroupTotal;
		nGroupTotal = 0;
	}

	if (nTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		//TranslateString(_T("T\x1ED5ng \x63\x1ED9ng"), tmpStr);
		rptDetail->SetValue(_T("TotalGroup"), _T("T\x1ED5ng \x63\x1ED9ng"));
		FormatCurrency(nTotal, tmpStr);
		rptDetail->SetValue(_T("TotalAmount"), tmpStr);
	}

	if (!m_szClerkKey.IsEmpty())
	{
		tmpStr = m_wndClerk.GetCurrent(1);
		StringUpper(tmpStr, szTemp);
		rpt.GetReportFooter()->SetValue(_T("ReceiverBy"), szTemp);
	}


	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
		          szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);

	EndWaitCursor();
	rpt.PrintPreview();
} 
void CFMConditionDepositPatientList::OnExportSelect()
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
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 20);
	xls.SetColumnWidth(4, 12);
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

	CString szStr1, szStr2;

	if (m_bPaided)
		szStr1 = _T("\x110\xC3 TH\x41NH TO\xC1N");
	else
		szStr1 = _T("\x43H\x1AF\x41 TH\x41NH TO\xC1N");

	if (m_bDischarge)
		szStr2 = _T("\x110\xC3 R\x41 VI\x1EC6N");
	else if (m_bIndischarge)
		szStr2 = _T("\x43H\x1AF\x41 R\x41 VI\x1EC6N");

	tmpStr.Format(_T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N T\x1EA0M G\x1EECI %s %s"),
		          szStr2, szStr1);

	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, true, 14, 0);

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

	xls.SetCellText(nCol + 4, nRow + 5, _T("Ng\xE0y n\x1ED9p"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 5, nRow + 5, _T("Ti\x1EC1n t\x1EA1m g\x1EEDi"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	int nIndex = 1;
	nRow = 6;
	CString szOldLine, szNewLine;
	long double nGroupTotal = 0;
	long double nTotal = 0;
	double nCost;


	while (!rs.IsEOF())
	{
		rs.GetValue(_T("deptid"), szNewLine);

		if (!szNewLine.IsEmpty() && szNewLine != szOldLine)
		{
			if (nGroupTotal > 0)
			{
				nRow++;
				xls.SetCellText(nCol + 1, nRow, _T("\x43\x1ED9ng kho\x61"), FMT_TEXT | FMT_WRAPING, true);
				tmpStr.Format(_T("%.2lf"), nGroupTotal);
				xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
				nTotal += nGroupTotal;
				nGroupTotal = 0;
			}
			szOldLine = szNewLine;
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

		rs.GetValue(_T("receivedate"), tmpStr);
		szTemp = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		xls.SetCellText(nCol + 4, nRow, szTemp, FMT_DATE | FMT_WRAPING);

		rs.GetValue(_T("amount"), nCost);
		nGroupTotal += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.MoveNext();
	}

	if (nGroupTotal > 0)
	{
		nRow++;
		xls.SetCellText(nCol + 1, nRow, _T("\x43\x1ED9ng kho\x61"), FMT_TEXT | FMT_WRAPING, true);
		
		tmpStr.Format(_T("%.2lf"), nGroupTotal);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
		nTotal += nGroupTotal;
		nGroupTotal = 0;

	}

	if (nTotal > 0)
	{
		nRow++;
		xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_WRAPING, true);
		
		tmpStr.Format(_T("%.2lf"), nTotal);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\DSBNTamGuiNoiTru.xls"));
} 
void CFMConditionDepositPatientList::OnObjectListDblClick(){
	
} 
void CFMConditionDepositPatientList::OnObjectListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMConditionDepositPatientList::OnObjectListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMConditionDepositPatientList::OnObjectListLoadData()
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
}
void CFMConditionDepositPatientList::OnDeptListDblClick(){
	
} 
void CFMConditionDepositPatientList::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMConditionDepositPatientList::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMConditionDepositPatientList::OnDeptListLoadData()
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
void CFMConditionDepositPatientList::OnDischargeSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bIndischarge = FALSE;
	UpdateData(FALSE);
}
void CFMConditionDepositPatientList::OnIndischargeSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bDischarge = FALSE;
	UpdateData(FALSE);
}
void CFMConditionDepositPatientList::OnPaidedSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bUnpaided = FALSE;
	UpdateData(FALSE);
}
void CFMConditionDepositPatientList::OnUnpaidedSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bPaided = FALSE;
	UpdateData(FALSE);
}
void CFMConditionDepositPatientList::OnLockedSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
int CFMConditionDepositPatientList::OnAddFMConditionDepositPatientList(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMConditionDepositPatientList::OnEditFMConditionDepositPatientList(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMConditionDepositPatientList::OnDeleteFMConditionDepositPatientList(){
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
 		OnCancelFMConditionDepositPatientList();
 	}
	return 0;
}
int CFMConditionDepositPatientList::OnSaveFMConditionDepositPatientList(){
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
 		//OnFMConditionDepositPatientListListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CFMConditionDepositPatientList::OnCancelFMConditionDepositPatientList(){
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
int CFMConditionDepositPatientList::OnFMConditionDepositPatientListListLoadData(){
	return 0;
}

CString CFMConditionDepositPatientList::GetQueryString()
{
	CString szSQL;
	CString szWhere;
	CString szDepts, szObjects;
	CString szField;

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
		szWhere.AppendFormat(_T(" and hfe_staff='%s' "), m_szClerkKey);
	}

	if (!szObjects.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and hfe_objectid in(%s) "), szObjects);
	}

	if (!szDepts.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and hfe_deptid in(%s) "), szDepts);
	}

	if (m_bPaided)
	{
		szWhere.AppendFormat(_T(" and hfe_status='P' "));
		szField.Format(_T("hfe_patpaid"));
	}
	else
	{
		szWhere.AppendFormat(_T(" and hfe_status<>'P' "));
		szField.Format(_T("hfe_amount"));
	}

	if (m_bDischarge)
		szWhere.AppendFormat(_T(" and hcr_status='T' "));
	else if (m_bIndischarge)
		szWhere.AppendFormat(_T(" and hcr_status<>'T' "));

	if (!m_bLocked)
	{
		szWhere.AppendFormat(_T(" and hfe_locked<>'Y' "));
		szWhere.AppendFormat(_T(" and hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}
	else
	{
		szWhere.AppendFormat(_T(" and hfe_locked='Y' "));
		szWhere.AppendFormat(_T(" and hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}

	szSQL.Format(_T(" select trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("        hcr_recordno as recordno,") \
				_T("        hfe_deptid as deptid,") \
				_T("        get_departmentname(hfe_deptid) as deptname,") \
				_T("        Cast_Timestamp2Date(hfe_date) as receivedate,") \
				_T("        sum(%s) as amount") \
				_T(" from hms_patient") \
				_T(" left join hms_doc on(hd_patientno=hp_patientno)") \
				_T(" left join hms_clinical_record on(hcr_docno=hd_docno)") \
				_T(" left join hms_fee_deposit on(hfe_docno=hd_docno)") \
				_T(" left join hms_object on(ho_id=hfe_objectid)") \
				_T(" where hfe_class in('I') %s") \
				_T(" group by hfe_docno, hcr_recordno, hp_surname, hp_midname,") \
				_T("          hp_firstname, hfe_deptid, Cast_Timestamp2Date(hfe_date)") \
				_T(" having sum(%s) > 0") \
				_T(" order by deptid, pname"),
				szField, szWhere, szField);

	return szSQL;
}

int CFMConditionDepositPatientList::OnDeptListCheckAll()
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

int CFMConditionDepositPatientList::OnDeptListUnCheckAll()
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


int CFMConditionDepositPatientList::OnObjectListCheckAll()
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

int CFMConditionDepositPatientList::OnObjectListUnCheckAll()
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