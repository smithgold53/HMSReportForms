#include "stdafx.h"
#include "FMFeeInvoiceCancelReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMFeeInvoiceCancelReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMFeeInvoiceCancelReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMFeeInvoiceCancelReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMFeeInvoiceCancelReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMFeeInvoiceCancelReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMFeeInvoiceCancelReport* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMFeeInvoiceCancelReport *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnClerkAddNew();
}*/
/*static void _OnDocNoChangeFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnDocNoChange();
} */
/*static void _OnDocNoSetfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnDocNoSetfocus();} */ 
/*static void _OnDocNoKillfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnDocNoKillfocus();
} */
static int _OnDocNoCheckValueFnc(CWnd *pWnd){
	return ((CFMFeeInvoiceCancelReport *)pWnd)->OnDocNoCheckValue();
} 
/*static void _OnInvoiceNoChangeFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnInvoiceNoChange();
} */
/*static void _OnInvoiceNoSetfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnInvoiceNoSetfocus();} */ 
/*static void _OnInvoiceNoKillfocusFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport *)pWnd)->OnInvoiceNoKillfocus();
} */
static int _OnInvoiceNoCheckValueFnc(CWnd *pWnd){
	return ((CFMFeeInvoiceCancelReport *)pWnd)->OnInvoiceNoCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMFeeInvoiceCancelReport *pVw = (CFMFeeInvoiceCancelReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMFeeInvoiceCancelReport *pVw = (CFMFeeInvoiceCancelReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnInsuranceSelectFnc(CWnd *pWnd){
	 ((CFMFeeInvoiceCancelReport*)pWnd)->OnInsuranceSelect();
}
static void _OnServiceSelectFnc(CWnd *pWnd){
	 ((CFMFeeInvoiceCancelReport*)pWnd)->OnServiceSelect();
}
static void _OnPoliciesSelectFnc(CWnd *pWnd){
	 ((CFMFeeInvoiceCancelReport*)pWnd)->OnPoliciesSelect();
}
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CFMFeeInvoiceCancelReport*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CFMFeeInvoiceCancelReport*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMFeeInvoiceCancelReport*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMFeeInvoiceCancelReport*)pWnd)->OnListDeleteItem();
} 
static int _OnAddFMFeeInvoiceCancelReportFnc(CWnd *pWnd){
	 return ((CFMFeeInvoiceCancelReport*)pWnd)->OnAddFMFeeInvoiceCancelReport();
} 
static int _OnEditFMFeeInvoiceCancelReportFnc(CWnd *pWnd){
	 return ((CFMFeeInvoiceCancelReport*)pWnd)->OnEditFMFeeInvoiceCancelReport();
} 
static int _OnDeleteFMFeeInvoiceCancelReportFnc(CWnd *pWnd){
	 return ((CFMFeeInvoiceCancelReport*)pWnd)->OnDeleteFMFeeInvoiceCancelReport();
} 
static int _OnSaveFMFeeInvoiceCancelReportFnc(CWnd *pWnd){
	 return ((CFMFeeInvoiceCancelReport*)pWnd)->OnSaveFMFeeInvoiceCancelReport();
} 
static int _OnCancelFMFeeInvoiceCancelReportFnc(CWnd *pWnd){
	 return ((CFMFeeInvoiceCancelReport*)pWnd)->OnCancelFMFeeInvoiceCancelReport();
}

static int _OnListCheckAllFnc(CWnd *pWnd){
	return ((CFMFeeInvoiceCancelReport*)pWnd)->OnListCheckAll();
}
static int _OnListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMFeeInvoiceCancelReport*)pWnd)->OnListUnCheckAll();
}

CFMFeeInvoiceCancelReport::CFMFeeInvoiceCancelReport(CWnd* pParent)
{
	m_nDlgWidth = 475;
	m_nDlgHeight = 525;
	SetDefaultValues();
}
CFMFeeInvoiceCancelReport::~CFMFeeInvoiceCancelReport(){
}
void CFMFeeInvoiceCancelReport::OnCreateComponents()
{
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 485);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 150, 425, 450);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 90, 115);
	m_wndClerk.Create(this,95, 90, 425, 115); 
	m_wndDocNoLabel.Create(this, _T("Doc No"), 10, 120, 90, 145);
	m_wndDocNo.Create(this,95, 120, 215, 145); 
	m_wndInvoiceNoLabel.Create(this, _T("Invoice No"), 220, 120, 300, 145);
	m_wndInvoiceNo.Create(this,305, 120, 425, 145); 
	m_wndInsurance.Create(this, _T("Insurance"), 175, 455, 255, 480);
	m_wndService.Create(this, _T("Service"), 260, 455, 340, 480);
	m_wndPolicies.Create(this, _T("Policies"), 345, 455, 425, 480);
	m_wndExport.Create(this, _T("&Export"), 330, 490, 430, 515);
	m_wndList.Create(this,15, 175, 420, 445); 

}
void CFMFeeInvoiceCancelReport::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(FALSE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(65);
	m_wndDocNo.SetLimitText(16);
	//m_wndDocNo.SetCheckValue(true);
	m_wndInvoiceNo.SetLimitText(16);
	//m_wndInvoiceNo.SetCheckValue(true);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 75);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);

	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);

	m_wndList.SetCheckBox(TRUE);
	m_wndList.ModifyStyle(0, LVS_NOSORTHEADER);
}
void CFMFeeInvoiceCancelReport::OnSetWindowEvents()
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
	//m_wndDocNo.SetEvent(WE_CHANGE, _OnDocNoChangeFnc);
	//m_wndDocNo.SetEvent(WE_SETFOCUS, _OnDocNoSetfocusFnc);
	//m_wndDocNo.SetEvent(WE_KILLFOCUS, _OnDocNoKillfocusFnc);
	m_wndDocNo.SetEvent(WE_CHECKVALUE, _OnDocNoCheckValueFnc);
	//m_wndInvoiceNo.SetEvent(WE_CHANGE, _OnInvoiceNoChangeFnc);
	//m_wndInvoiceNo.SetEvent(WE_SETFOCUS, _OnInvoiceNoSetfocusFnc);
	//m_wndInvoiceNo.SetEvent(WE_KILLFOCUS, _OnInvoiceNoKillfocusFnc);
	m_wndInvoiceNo.SetEvent(WE_CHECKVALUE, _OnInvoiceNoCheckValueFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndInsurance.SetEvent(WE_CLICK, _OnInsuranceSelectFnc);
	m_wndService.SetEvent(WE_CLICK, _OnServiceSelectFnc);
	m_wndPolicies.SetEvent(WE_CLICK, _OnPoliciesSelectFnc);
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);

	m_wndList.SetWindowText(_T("Dept"));
	m_wndList.AddEvent(1, _T("Check All"), _OnListCheckAllFnc);
	m_wndList.AddEvent(2, _T("UnCheck All"), _OnListUnCheckAllFnc);

	//SetWindowFont(&m_wndDocNo, GetFaceName(), GetFaceSize() + 2, true);
	m_wndDocNo.SetTextColor(RGB(255, 0, 0));

	//SetWindowFont(&m_wndInvoiceNo, GetFaceName(), GetFaceSize() + 2, true);
	m_wndInvoiceNo.SetTextColor(RGB(0, 0, 255));

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL;
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);
	OnListLoadData();
}
void CFMFeeInvoiceCancelReport::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Text(pDX, m_wndDocNo.GetDlgCtrlID(), m_nDocNo);
	DDX_Text(pDX, m_wndInvoiceNo.GetDlgCtrlID(), m_nInvoiceNo);
	DDX_Check(pDX, m_wndInsurance.GetDlgCtrlID(), m_bInsurance);
	DDX_Check(pDX, m_wndService.GetDlgCtrlID(), m_bService);
	DDX_Check(pDX, m_wndPolicies.GetDlgCtrlID(), m_bPolicies);

}
void CFMFeeInvoiceCancelReport::GetDataToScreen()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMFeeInvoiceCancelReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMFeeInvoiceCancelReport::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_nDocNo=0;
	m_nInvoiceNo=0;
	m_bInsurance=TRUE;
	m_bService=FALSE;
	m_bPolicies=FALSE;

}
int CFMFeeInvoiceCancelReport::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, 1, -1);
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
/*void CFMFeeInvoiceCancelReport::OnYearChange(){
	
} */
/*void CFMFeeInvoiceCancelReport::OnYearSetfocus(){
	
} */
/*void CFMFeeInvoiceCancelReport::OnYearKillfocus(){
	
} */
int CFMFeeInvoiceCancelReport::OnYearCheckValue()
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
void CFMFeeInvoiceCancelReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMFeeInvoiceCancelReport::OnReportPeriodSelendok()
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
/*void CFMFeeInvoiceCancelReport::OnReportPeriodSetfocus(){
	
}*/
/*void CFMFeeInvoiceCancelReport::OnReportPeriodKillfocus(){
	
}*/
long CFMFeeInvoiceCancelReport::OnReportPeriodLoadData()
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
/*void CFMFeeInvoiceCancelReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMFeeInvoiceCancelReport::OnFromDateChange(){
	
} */
/*void CFMFeeInvoiceCancelReport::OnFromDateSetfocus(){
	
} */
/*void CFMFeeInvoiceCancelReport::OnFromDateKillfocus(){
	
} */
int CFMFeeInvoiceCancelReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMFeeInvoiceCancelReport::OnToDateChange(){
	
} */
/*void CFMFeeInvoiceCancelReport::OnToDateSetfocus(){
	
} */
/*void CFMFeeInvoiceCancelReport::OnToDateKillfocus(){
	
} */
int CFMFeeInvoiceCancelReport::OnToDateCheckValue(){
	return 0;
} 
void CFMFeeInvoiceCancelReport::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMFeeInvoiceCancelReport::OnClerkSelendok(){
	 
}
/*void CFMFeeInvoiceCancelReport::OnClerkSetfocus(){
	
}*/
/*void CFMFeeInvoiceCancelReport::OnClerkKillfocus(){
	
}*/
long CFMFeeInvoiceCancelReport::OnClerkLoadData()
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
	szSQL.Format(_T("SELECT su_userid as userid, su_name as name ") \
		         _T("FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);

	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("userid")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMFeeInvoiceCancelReport::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMFeeInvoiceCancelReport::OnDocNoChange(){
	
} */
/*void CFMFeeInvoiceCancelReport::OnDocNoSetfocus(){
	
} */
/*void CFMFeeInvoiceCancelReport::OnDocNoKillfocus(){
	
} */
int CFMFeeInvoiceCancelReport::OnDocNoCheckValue(){
	return 0;
} 
/*void CFMFeeInvoiceCancelReport::OnInvoiceNoChange(){
	
} */
/*void CFMFeeInvoiceCancelReport::OnInvoiceNoSetfocus(){
	
} */
/*void CFMFeeInvoiceCancelReport::OnInvoiceNoKillfocus(){
	
} */
int CFMFeeInvoiceCancelReport::OnInvoiceNoCheckValue(){
	return 0;
} 
void CFMFeeInvoiceCancelReport::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMFeeInvoiceCancelReport::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szWhere;
	CString tmpStr, szTemp;
	CRecord rs(&pMF->m_db);
	UpdateData(TRUE);

	BeginWaitCursor();

	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 17);
	xls.SetColumnWidth(2, 18);
	xls.SetColumnWidth(3, 17);
	xls.SetColumnWidth(4, 18);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 10);
	xls.SetColumnWidth(7, 12);
	xls.SetColumnWidth(8, 15);
	xls.SetColumnWidth(9, 13);
	xls.SetColumnWidth(10, 15);


	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 4);
	xls.SetCellMergedColumns(nCol, nRow + 2, 4);

	xls.SetCellMergedColumns(nCol + 6, nRow + 1, 5);
	xls.SetCellMergedColumns(nCol + 6, nRow + 2, 5);

	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellText(nCol, nRow + 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 12);

	xls.SetCellText(nCol + 6, nRow + 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellText(nCol + 6, nRow + 2, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true, 12);

	xls.SetCellMergedColumns(nCol, nRow + 3, 11);
	xls.SetCellText(nCol, nRow + 3, _T("\x44\x41NH S\xC1\x43H H\x1EE6Y H\xD3\x41 \x110\x1A0N"), FMT_TEXT | FMT_CENTER, true, 14);
	xls.SetCellMergedColumns(nCol, nRow + 4, 11);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12);

	xls.SetCellText(nCol, nRow + 5, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 5, _T("Ng\x1B0\x1EDDi t\x1EA1o"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 5, _T("Ng\xE0y t\x1EA1o"), FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellText(nCol + 3, nRow + 5, _T("Ng\x1B0\x1EDDi h\x1EE7y"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 5, _T("Ng\xE0y h\x1EE7y"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol + 5, nRow + 5, _T("M\xE3 \x62\x1EC7nh nh\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 6, nRow + 5, _T("S\x1ED1 h\x1ED3 s\x1A1"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 7, nRow + 5, _T("S\x1ED1 h\xF3\x61 \x111\x1A1n"), FMT_TEXT |FMT_CENTER, true, 11);	

	xls.SetCellText(nCol + 8, nRow + 5, _T("T\x1ED5ng \x63hi ph\xED"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 9, nRow + 5, _T("Mi\x1EC5n gi\x1EA3m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 10, nRow + 5, _T("\x42\x1EC7nh nh\xE2n tr\x1EA3"), FMT_TEXT |FMT_CENTER, true, 11);	

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	//_fmsg(_T("%s"), szSQL);
	int nIndex = 1;

	while(!rs.IsEOF())
	{
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow + 6, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("tennguoitao"),tmpStr);
		xls.SetCellText(nCol + 1, nRow + 6, tmpStr, FMT_TEXT);
 
		rs.GetValue(_T("ngaytao"), tmpStr);
		szTemp = CDateTime::Convert(tmpStr, yyyymmdd | hhmm, ddmmyyyy | hhmm);
		xls.SetCellText(nCol + 2, nRow + 6, szTemp, FMT_DATETIME);

		rs.GetValue(_T("tennguoicapnhat"), tmpStr);
		xls.SetCellText(nCol + 3, nRow + 6, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("ngaycapnhat"), tmpStr);
		szTemp = CDateTime::Convert(tmpStr, yyyymmdd | hhmm, ddmmyyyy | hhmm);
		xls.SetCellText(nCol + 4, nRow + 6, szTemp, FMT_DATETIME);
		
		rs.GetValue(_T("mabenhnhan"), tmpStr);
		xls.SetCellText(nCol + 5, nRow + 6, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("sohoso"), tmpStr);
		xls.SetCellText(nCol + 6, nRow + 6, tmpStr, FMT_INTEGER);
		
		rs.GetValue(_T("sohoadon"), tmpStr);
		xls.SetCellText(nCol + 7, nRow + 6, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("tongchiphi"), tmpStr);
		xls.SetCellText(nCol + 8, nRow + 6, tmpStr, FMT_CURRENCY1);

		rs.GetValue(_T("miengiam"), tmpStr);
		xls.SetCellText(nCol + 9, nRow + 6, tmpStr, FMT_CURRENCY1);

		rs.GetValue(_T("benhnhantra"), tmpStr);
		xls.SetCellText(nCol + 10, nRow + 6, tmpStr, FMT_CURRENCY1);

		nRow++;
		rs.MoveNext();
			
	}

	xls.Save(_T("Exports\\DSHUYHOADON.xls"));
	EndWaitCursor();
}

void CFMFeeInvoiceCancelReport::OnInsuranceSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
	void CFMFeeInvoiceCancelReport::OnServiceSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
	void CFMFeeInvoiceCancelReport::OnPoliciesSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
	void CFMFeeInvoiceCancelReport::OnListDblClick(){
	
} 
void CFMFeeInvoiceCancelReport::OnListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMFeeInvoiceCancelReport::OnListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMFeeInvoiceCancelReport::OnListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	m_wndList.BeginLoad(); 
	m_wndList.DeleteAllItems(); 
	int nCount = 0;

	szSQL.Format(_T("SELECT sd_id AS ID, sd_name AS Name ") \
		         _T("FROM sys_dept ") \
				 _T("WHERE sd_type IN('KB') ") \
				 _T("ORDER BY sd_id"));
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
int CFMFeeInvoiceCancelReport::OnAddFMFeeInvoiceCancelReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMFeeInvoiceCancelReport::OnEditFMFeeInvoiceCancelReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMFeeInvoiceCancelReport::OnDeleteFMFeeInvoiceCancelReport(){
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
 		OnCancelFMFeeInvoiceCancelReport(); 
 	}
	return 0;
}
int CFMFeeInvoiceCancelReport::OnSaveFMFeeInvoiceCancelReport(){
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
 		//OnFMFeeInvoiceCancelReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMFeeInvoiceCancelReport::OnCancelFMFeeInvoiceCancelReport(){
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
int CFMFeeInvoiceCancelReport::OnFMFeeInvoiceCancelReportListLoadData(){
	return 0;
}

int CFMFeeInvoiceCancelReport::OnListCheckAll()
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

int CFMFeeInvoiceCancelReport::OnListUnCheckAll()
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

CString CFMFeeInvoiceCancelReport::GetQueryString()
{
	CString szSQL, szWhere;
	CString szDeptID, szObject;

	szWhere.Empty();
	szDeptID.Empty();
	szObject.Empty();

	if (m_bInsurance)
		szObject.AppendFormat(_T(" or ho_type in('I','C') "));
	if (m_bService)
		szObject.AppendFormat(_T(" or ho_type in('S') "));
	if (m_bPolicies)
		szObject.AppendFormat(_T(" or ho_type not in('I','C','S') "));

	if (!m_szClerkKey.IsEmpty())
		szWhere.Format(_T(" and hfe_staff='%s' "), m_szClerkKey);

	if (m_nDocNo > 0)
		szWhere.AppendFormat(_T(" and hfe_docno=%ld "), m_nDocNo);

	if (m_nInvoiceNo > 0)
		szWhere.AppendFormat(_T(" and hfe_invoiceno=%ld "), m_nInvoiceNo);

	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (!szDeptID.IsEmpty())
				szDeptID += _T(",");
			szDeptID.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
		}
	}
	
	if (!szDeptID.IsEmpty())
		szWhere.AppendFormat(_T(" and hfe_deptid in(%s) "), szDeptID);


	szSQL.Format(_T(" select hfe_createdby as nguoitao,") \
				_T("        (select distinct su_name from sys_user where su_userid=hfe_createdby) as tennguoitao,") \
				_T("        hfe_createddate as ngaytao,") \
				_T("        hfe_updatedby as capnhatboi,") \
				_T("        (select distinct su_name from sys_user where su_userid=hfe_updatedby) as tennguoicapnhat,") \
				_T("        hfe_updateddate as ngaycapnhat,") \
				_T("        hfe_patientno as mabenhnhan,") \
				_T("        hfe_docno as sohoso,") \
				_T("        hfe_invoiceno as sohoadon,") \
				_T("        hfe_amount as tongchiphi,") \
				_T("        hfe_discount as miengiam,") \
				_T("        hfe_patpaid as benhnhantra") \
				_T(" from hms_fee_invoice") \
				_T(" left join hms_doc on(hd_docno=hfe_docno)") \
				_T(" left join hms_object on(ho_id=hd_object)") \
				_T(" where hfe_updateddate between cast_string2timestamp('%s')") \
				_T("       and cast_string2timestamp('%s') and hfe_status='C'") \
				_T("       and (1=1 %s) %s") \
				_T(" order by hfe_updateddate, hfe_docno"),
				m_szFromDate, m_szToDate, szObject, szWhere);

	return szSQL;
}