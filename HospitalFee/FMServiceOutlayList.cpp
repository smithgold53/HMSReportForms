#include "stdafx.h"
#include "FMServiceOutlayList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMServiceOutlayList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMServiceOutlayList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceOutlayList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMServiceOutlayList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMServiceOutlayList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMServiceOutlayList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceOutlayList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnClerkAddNew();
}*/
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMServiceOutlayList* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceOutlayList *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CFMServiceOutlayList *)pWnd)->OnObjectAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMServiceOutlayList *pVw = (CFMServiceOutlayList *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMServiceOutlayList *pVw = (CFMServiceOutlayList *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnApprovalSelectFnc(CWnd *pWnd){
	 ((CFMServiceOutlayList*)pWnd)->OnApprovalSelect();
}
static long _OnPaymentListLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceOutlayList*)pWnd)->OnPaymentListLoadData();
} 
static void _OnPaymentListDblClickFnc(CWnd *pWnd){
	((CFMServiceOutlayList*)pWnd)->OnPaymentListDblClick();
} 
static void _OnPaymentListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMServiceOutlayList*)pWnd)->OnPaymentListSelectChange(nOldItem, nNewItem);
} 
static int _OnPaymentListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMServiceOutlayList*)pWnd)->OnPaymentListDeleteItem();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceOutlayList*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMServiceOutlayList*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMServiceOutlayList*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMServiceOutlayList*)pWnd)->OnDeptListDeleteItem();
} 
CFMServiceOutlayList::CFMServiceOutlayList(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CFMServiceOutlayList::~CFMServiceOutlayList(){
}
void CFMServiceOutlayList::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 550);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 335, 425, 545);
	m_wndPayment.Create(this, _T("Payment"), 10, 120, 425, 330);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 90, 115);
	m_wndClerk.Create(this,95, 90, 215, 115); 
	m_wndObjectLabel.Create(this, _T("Object"), 220, 90, 300, 115);
	m_wndObject.Create(this,305, 90, 425, 115); 
	m_wndPrint.Create(this, _T("&Print"), 265, 555, 345, 580);
	m_wndExport.Create(this, _T("&Export"), 350, 555, 430, 580);
	m_wndApproval.Create(this, _T("Approval"), 5, 555, 85, 580);
	m_wndPaymentList.Create(this,15, 145, 420, 325); 
	m_wndDeptList.Create(this,15, 358, 420, 538); 

}
void CFMServiceOutlayList::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(true);

	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 40);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.SetCheckBox(TRUE);

	m_wndPaymentList.InsertColumn(0, _T("ID"), CFMT_NUMBER, 60);
	m_wndPaymentList.InsertColumn(1, _T("Name"), CFMT_TEXT, 125);
	m_wndPaymentList.InsertColumn(2, _T("Locked"), CFMT_TEXT, 80);
	m_wndPaymentList.SetCheckBox(true);
}

void CFMServiceOutlayList::OnSetWindowEvents(){
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
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndApproval.SetEvent(WE_CLICK, _OnApprovalSelectFnc);
	m_wndPaymentList.SetEvent(WE_SELCHANGE, _OnPaymentListSelectChangeFnc);
	m_wndPaymentList.SetEvent(WE_LOADDATA, _OnPaymentListLoadDataFnc);
	m_wndPaymentList.SetEvent(WE_DBLCLICK, _OnPaymentListDblClickFnc);

	m_wndDeptList.SetEvent(WE_SELCHANGE, _OnDeptListSelectChangeFnc);
	m_wndDeptList.SetEvent(WE_LOADDATA, _OnDeptListLoadDataFnc);
	m_wndDeptList.SetEvent(WE_DBLCLICK, _OnDeptListDblClickFnc);

	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);
	OnDeptListLoadData();
	OnPaymentListLoadData();
}
void CFMServiceOutlayList::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_Check(pDX, m_wndApproval.GetDlgCtrlID(), m_bApproval);

}
void CFMServiceOutlayList::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
}

int CFMServiceOutlayList::SetMode(int nMode){
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

/*void CFMServiceOutlayList::OnYearChange(){
	
} */
/*void CFMServiceOutlayList::OnYearSetfocus(){
	
} */
/*void CFMServiceOutlayList::OnYearKillfocus(){
	
} */
int CFMServiceOutlayList::OnYearCheckValue(){
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
			szTemp.Format(_T("%.2d/2d/4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndFromDate.SetWindowText(szTemp);
		}
		dt.ParseDateTime(m_szToDate);
		date = dt.GetDate();
		if (date.GetYear() != 1752)
		{
			dt.SetDate(m_nYear, date.GetMonth(), date.GetDay());
			m_szToDate = dt.GetDateTime();
			szTemp.Format(_T("%.2d/2d/4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndToDate.SetWindowText(szTemp);
		}
	}
	UpdateData(FALSE);
	OnPaymentListLoadData();
	return 0;
}
 
void CFMServiceOutlayList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
void CFMServiceOutlayList::OnReportPeriodSelendok(){
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	CString tmpStr;
	CDate dte, date;

	UpdateData(true);

	date.ParseDate(pMF->GetSysDate());
	int nYear = date.GetYear();
	int nMonth = ToInt(m_szReportPeriodKey);

	if (nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/1 00:00"), nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), nYear, nMonth, dte.GetMonthLastDay());
	}

	if (nMonth == 13)
	{
		m_szFromDate.Format(_T("%.4d/1/1 00:00"), nYear);
		tmpStr.Format(_T("%.4d/3/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/3/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 14)
	{
		m_szFromDate.Format(_T("%.4d/4/1 00:00"), nYear);
		tmpStr.Format(_T("%.4d/6/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/6/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 15)
	{
		m_szFromDate.Format(_T("%.4d/7/1 00:00"), nYear);
		tmpStr.Format(_T("%.4d/9/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/9/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 16)
	{
		m_szFromDate.Format(_T("%.4d/10/1 00:00"), nYear);
		tmpStr.Format(_T("%.4d/10/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 17)
	{
		m_szFromDate.Format(_T("%.4d/1/1 00:00"), nYear);
		tmpStr.Format(_T("%.4d/12/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}

	UpdateData(false); 
	OnPaymentListLoadData();
}

/*void CFMServiceOutlayList::OnReportPeriodSetfocus(){
	
}*/
/*void CFMServiceOutlayList::OnReportPeriodKillfocus(){
	
}*/
long CFMServiceOutlayList::OnReportPeriodLoadData(){
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

/*void CFMServiceOutlayList::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMServiceOutlayList::OnFromDateChange(){
	
} */
/*void CFMServiceOutlayList::OnFromDateSetfocus(){
	
} */
/*void CFMServiceOutlayList::OnFromDateKillfocus(){
	
} */
int CFMServiceOutlayList::OnFromDateCheckValue(){
	OnPaymentListLoadData();
	return 0;
}
 
/*void CFMServiceOutlayList::OnToDateChange(){
	
} */
/*void CFMServiceOutlayList::OnToDateSetfocus(){
	
} */
/*void CFMServiceOutlayList::OnToDateKillfocus(){
	
} */
int CFMServiceOutlayList::OnToDateCheckValue(){
	OnPaymentListLoadData();
	return 0;
}
 
void CFMServiceOutlayList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMServiceOutlayList::OnClerkSelendok(){
	 
}
/*void CFMServiceOutlayList::OnClerkSetfocus(){
	
}*/
/*void CFMServiceOutlayList::OnClerkKillfocus(){
	
}*/
long CFMServiceOutlayList::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND su_deptid = 'PTC'"));
	return pMF->LoadUserList(&m_wndClerk, m_szClerkKey, szFilter);
}

/*void CFMServiceOutlayList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMServiceOutlayList::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMServiceOutlayList::OnObjectSelendok(){
	 
}
/*void CFMServiceOutlayList::OnObjectSetfocus(){
	
}*/
/*void CFMServiceOutlayList::OnObjectKillfocus(){
	
}*/
long CFMServiceOutlayList::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT ho_id as idx, ho_desc descr FROM hms_object ORDER BY cast(ho_id as integer)"));
	long nCount = rs.ExecSQL(szSQL);
	m_wndObject.DeleteAllItems();
	while (!rs.IsEOF())
	{
		m_wndObject.AddItems(
			rs.GetValue(_T("idx")),
			rs.GetValue(_T("descr")),
			NULL);
		rs.MoveNext();
	}
	return nCount;
}

/*void CFMServiceOutlayList::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMServiceOutlayList::OnPrintSelect(){
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

	if (!rpt.Init(_T("Reports/HMS/HF_BANGKECHIBNDICHVU.RPT")))
		return;

	StringUpper(pMF->m_szHealthService, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);

	StringUpper(pMF->m_szHospitalName, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	CReportSection *rptDetail;
	CString szOldLine, szNewLine;
	CStringArray arrField;
	long double nTotal[6];
	double nCost;
	int nIndex = 1;

	for (int i = 0; i < 6; i++)
	{
		nTotal[i] = 0;
	}
	arrField.Add(_T("hitechreturn"));
	arrField.Add(_T("eamount"));
	arrField.Add(_T("iamount"));
	arrField.Add(_T("deposit"));
	arrField.Add(_T("refund"));
	arrField.Add(_T("totalamount"));
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("recordno"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		//rs.GetValue(_T("deptid"), tmpStr);
		//rptDetail->SetValue(_T("5"), tmpStr);

		for (int j = 0; j < 6; j++)
		{
			rs.GetValue(arrField.GetAt(j), nCost);
			nTotal[j] += nCost;
			FormatCurrency(nCost, tmpStr);
			szTemp.Format(_T("%d"), j+5);
			rptDetail->SetValue(szTemp, tmpStr);	
		}		

		rs.MoveNext();
	}

	if (nTotal[5] > 0)
	{
		for (int i = 0; i < 6; i++)
		{
			FormatCurrency(nTotal[i], tmpStr);
			szTemp.Format(_T("s%d"), i + 5);
			rpt.GetReportFooter()->SetValue(szTemp, tmpStr);
		}
	}

	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("User"), pMF->GetDoctorName(pMF->GetCurrentUser()));
	EndWaitCursor();
	rpt.PrintPreview();
}
 
void CFMServiceOutlayList::OnExportSelect(){
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
	CStringArray arrCol, arrField;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 18);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 5);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 15);
	xls.SetColumnWidth(7, 15);
	xls.SetColumnWidth(8, 15);
	xls.SetColumnWidth(9, 15);

	int nRow = 1;
	int nCol = 0;

	xls.SetRowHeight(6, 22);
	xls.SetRowHeight(7, 40);

	xls.SetCellMergedColumns(0, 1, 4);
	xls.SetCellMergedColumns(0, 2, 4);


	xls.SetCellText(0, 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol, nRow + 3, 18);
	xls.SetCellMergedColumns(nCol, nRow + 4, 18);

	xls.SetCellText(nCol, nRow + 3, _T("\x42\x1EA2NG K\xCA CHI \x42\x1EC6NH NH\xC2N \x44\x1ECA\x43H V\x1EE4"),
					FMT_TEXT | FMT_CENTER, true, 16, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);
	
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("H\x1ECD v\xE0 t\xEAn"));
	arrCol.Add(_T("S\x1ED1 \x42\x41"));
	arrCol.Add(_T("S\x1ED1 h\x1ED3 s\x1A1"));
	arrCol.Add(_T("Kho\x61 \x111i\x1EC1u tr\x1ECB"));
	arrCol.Add(_T("\x43\xE1\x63 kho\x1EA3n \x63hi tr\x1EA3 m\xE1y \x43N\x43"));
	arrCol.Add(_T("\x43\xE1\x63 kho\x1EA3n ph\x1EA3i chi"));
	arrCol.Add(_T("Vi\x1EC7n ph\xED ph\x1EA3i thu"));
	arrCol.Add(_T("Tr\xED\x63h t\x1EA1m g\x1EEDi"));
	arrCol.Add(_T("\x43hi tr\x1EA3 l\x1EA1i"));
	arrCol.Add(_T("T\x1ED5ng chi"));
	nRow = 6;
	for (int i = 0; i < arrCol.GetCount(); i++){
		xls.SetCellText(nCol + i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER|FMT_WRAPING, true, 10);
	}

	int nIndex = 1;
	nRow = 7;
	long double nTotal[6];
	double nCost;

	for (int i = 0; i < 6; i++)
	{
		nTotal[i] = 0;
	}
	arrField.Add(_T("hitechreturn"));
	arrField.Add(_T("eamount"));
	arrField.Add(_T("iamount"));
	arrField.Add(_T("deposit"));
	arrField.Add(_T("refund"));
	arrField.Add(_T("totalamount"));
	while (!rs.IsEOF())
	{
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("recordno"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("deptid"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		for (int j = 0; j < arrField.GetCount(); j++)
		{
			rs.GetValue(arrField.GetAt(j), nCost);
			nTotal[j] += nCost;
			tmpStr.Format(_T("%.2f"), nCost);
			xls.SetCellText(nCol + j + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);
		}
		nRow++;
		rs.MoveNext();
	}
	if (nTotal[5] > 0)
	{
		xls.SetCellText(nCol + 4, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 10);
		for (int i = 0; i < 6; i++)
		{
			tmpStr.Format(_T("%.2f"), nTotal[i]);
			xls.SetCellText(nCol + i + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);
		}
	}
	xls.Save(_T("Exports\\BANG KE CHI BENH NHAN DICH VU.xls"));
}
 
	void CFMServiceOutlayList::OnApprovalSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMServiceOutlayList::OnPaymentListDblClick(){
	
} 
void CFMServiceOutlayList::OnPaymentListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMServiceOutlayList::OnPaymentListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMServiceOutlayList::OnPaymentListLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND fac_user_id = '%s' AND fac_locked = 'Y'") \
				   _T("	AND fac_updateddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
					, pMF->GetCurrentUser(), m_szFromDate, m_szToDate);
	szSQL.Format(_T("SELECT fac_cash_id as idx, fac_invoiceno descr, fac_locked stt FROM fa_cash WHERE fac_invoicetype = 'P' %s ORDER BY fac_cash_id"), szWhere);
	long nCount = rs.ExecSQL(szSQL);
	m_wndPaymentList.BeginLoad();
	while (!rs.IsEOF())
	{
		m_wndPaymentList.AddItems(
			rs.GetValue(_T("idx")),
			rs.GetValue(_T("descr")),
			rs.GetValue(_T("stt")),
			NULL);
		rs.MoveNext();
	}
	m_wndPaymentList.EndLoad();
	return nCount;
}

void CFMServiceOutlayList::OnDeptListDblClick(){
	
}
 
void CFMServiceOutlayList::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
int CFMServiceOutlayList::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
}
 
long CFMServiceOutlayList::OnDeptListLoadData(){
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


CString CFMServiceOutlayList::GetQueryString(){
	CString szSQL, szWhere;
	CString szDepts, szObjects, szWhere1, szRefWhere, szRefWhere0, szPayments;
	szSQL.Empty();
	/*if (m_bReportByPayment)
	{
		if (m_szPaymentListKey.IsEmpty())
		{
			AfxMessageBox(_T("Select a payment!"));
			m_wndPaymentList.SetFocus();
			return szSQL;
		}
		szWhere.Format(_T(" AND hfe_cash_id = %s AND hfe_status = 'P'"), m_szPaymentListKey);
		szRefWhere.Format(_T(" AND (r.hfe_cash_id = %s AND r.hfe_status = 'P')"), m_szPaymentListKey);
		szRefWhere0.Format(_T(" AND rf.hfe_cash_id = %s AND rf.hfe_status = 'P'"), m_szPaymentListKey);
	}*/
	//else
	for (int i = 0; i < m_wndPaymentList.GetItemCount();i++)
	{
		if (m_wndPaymentList.GetCheck(i))
		{
			if (!szPayments.IsEmpty())
				szPayments += _T(", ");
			szPayments += m_wndPaymentList.GetItemText(i, 0);
		}
	}
	if (!szPayments.IsEmpty())
	{
		szWhere.Format(_T(" AND hfe_cash_id IN (%s) AND hfe_status = 'P'"), szPayments);
		szRefWhere.Format(_T(" AND (r.hfe_cash_id IN (%s) AND r.hfe_status = 'P')"), szPayments);
		szRefWhere0.Format(_T(" AND rf.hfe_cash_id IN (%s) AND rf.hfe_status = 'P'"), szPayments);
	}
	else
	{
		for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
		{
			if (m_wndDeptList.GetCheck(i))
			{
				if (!szDepts.IsEmpty())
					szDepts += _T(",");
				szDepts.AppendFormat(_T("'%s'"), m_wndDeptList.GetItemText(i, 0));
			}
		}
		for (int i = 0; i < m_wndObject.GetItemCount(); i++)
		{
			if (m_wndObject.GetCheck(i))
			{
				m_wndObject.SetCurSel(i);
				if (!szObjects.IsEmpty())
					szObjects += _T(", ");
				szObjects.AppendFormat(_T("%s"), m_wndObject.GetCurrent(0));
			}
		}
		szRefWhere0.Format(_T(" AND rf.hfe_status = 'P' AND rf.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
		szWhere.Format(_T(" AND hfe_status = 'P' AND hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
		szRefWhere.Format(_T(" AND r.hfe_status = 'P' AND r.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
		if (!m_szClerkKey.IsEmpty())
		{
			szRefWhere0.AppendFormat(_T(" AND rf.hfe_staff = '%s'"), m_szClerkKey);
			szWhere.AppendFormat(_T(" AND hfe_staff = '%s'"), m_szClerkKey);
			szRefWhere.AppendFormat(_T(" AND r.hfe_staff = '%s'"), m_szClerkKey);
		}
		if (!szDepts.IsEmpty())
		{
			szRefWhere0.AppendFormat(_T(" AND rf.hfe_deptid IN (%s)"), szDepts);
			szWhere.AppendFormat(_T(" AND hfe_deptid IN (%s)"), szDepts);
			szRefWhere.AppendFormat(_T(" AND r.hfe_deptid IN (%s)"), szDepts);
		}
		if (!szObjects.IsEmpty())
		{
			szRefWhere0.AppendFormat(_T(" AND rf.hfe_objectid IN (%s)"), szObjects);
			szWhere.AppendFormat(_T(" AND hfe_objectid IN (%s)"), szObjects);
			szRefWhere.AppendFormat(_T(" AND r.hfe_objectid IN (%s)"), szObjects);
		}
	}
	szSQL.Format(_T(" SELECT    Get_patientname(docno) pname, ") \
				_T("           hcr_recordno recordno, ") \
				_T("           docno, ") \
				_T("           deptid, ") \
				_T("		   hfe_date,") \
				_T("		   SUM(hitechreturn) hitechreturn,") \
				_T("           SUM(eamount) eamount, ") \
				_T("           SUM(iamount) iamount, ") \
				_T("           SUM(deposit) deposit, ") \
				_T("           SUM(refund)   AS refund, ") \
				_T("		   SUM(hitechreturn+eamount+refund) AS totalamount")
				_T(" FROM      (SELECT    rf.hfe_docno AS docno, ") \
				_T("                      rf.hfe_deptid AS deptid, ") \
				_T("                      rf.hfe_date hfe_date, ") \
				_T("                      SUM(CASE WHEN NVL(hfl_hitechmachine, 'X') = 'Y' THEN hfe_cost ELSE 0 END) AS hitechreturn, ") \
				_T("                      SUM(CASE WHEN NVL(hfl_hitechmachine, 'X') <> 'Y' THEN hfe_cost ELSE 0 END) AS eamount, ") \
				_T("                      0 AS iamount, ") \
				_T("                      0 AS deposit, ") \
				_T("                      0 AS refund") \
				_T("            FROM      hms_fee_refund rf ") \
				_T("            LEFT JOIN hms_fee_refundline rfl ON ( rf.hfe_docno = rfl.hfe_docno ") \
				_T("                                                  AND rf.hfe_invoiceno = rfl.hfe_invoiceno ) ") \
				_T("            LEFT JOIN hms_fee_list ON ( hfe_itemid = hfl_feeid ) ") \
				_T("            WHERE     rf.hfe_class IN ('E', 'X') ") \
				_T("			%s") \
				_T("            GROUP     BY rf.hfe_docno, ") \
				_T("                         rf.hfe_deptid, ") \
				_T("                         rf.hfe_date ") \
				_T("			UNION ALL") \
				_T("			SELECT r.hfe_docno  AS docno, ") \
				_T("                   r.hfe_deptid AS deptid, ") \
				_T("                   r.hfe_date, ") \
				_T("				   0,") \
				_T("                   0		  AS eamount, ") \
				_T("                   0          AS iamount, ") \
				_T("                   0          AS deposit, ") \
				_T("                   r.hfe_amount AS refund ") \
				_T("            FROM   hms_fee_refund r") \
				_T("			LEFT JOIN hms_fee_deposit d ON (r.hfe_refidx = d.hfe_invoiceno)") \
				_T("            WHERE  r.hfe_class IN ( 'A', 'I' ) AND (d.hfe_status = 'R' OR r.hfe_type = 'E') ") \
				_T("			%s") \
				_T("            UNION ALL ") \
				_T("            SELECT i.hfe_docno   AS docno, ") \
				_T("                   i.hfe_deptid  AS deptid, ") \
				_T("                   i.hfe_date, ") \
				_T("				   0,") \
				_T("                   0, ") \
				_T("                   i.hfe_amount, ") \
				_T("                   i.hfe_deposit, ") \
				_T("                   r.hfe_amount ") \
				_T("            FROM   hms_fee_refund r") \
				_T("			LEFT JOIN hms_fee_invoice i ON (i.hfe_docno = r.hfe_docno AND i.hfe_invoiceno = r.hfe_refidx)") \
				_T("            WHERE  r.hfe_class IN( 'A', 'I' ) AND r.hfe_type = 'F' AND i.hfe_refund > 0 ") \
				_T("			%s) ") \
				_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = docno ) ") \
				_T(" LEFT JOIN hms_doc ON ( hd_docno = docno ) ") \
				_T(" WHERE     1 = 1") \
				_T(" GROUP     BY docno, ") \
				_T("              hcr_recordno, ") \
				_T("              deptid, ") \
				_T("              hfe_date ") \
				_T(" HAVING sum(hitechreturn+eamount+refund) > 0") \
				_T(" ORDER     BY hfe_date "), szRefWhere0, szRefWhere, szRefWhere);
	return szSQL;
}
