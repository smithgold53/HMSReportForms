#include "stdafx.h"
#include "FMServiceIncomeList.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMServiceIncomeList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMServiceIncomeList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceIncomeList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMServiceIncomeList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMServiceIncomeList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMServiceIncomeList *)pWnd)->OnToDateCheckValue();
} 
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceIncomeList *)pWnd)->OnClerkLoadData();
}
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceIncomeList *)pWnd)->OnObjectLoadData();
}
static long _OnReceiptListLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceIncomeList *)pWnd)->OnReceiptListLoadData();
}
static void _OnReportByReceiptSelectFnc(CWnd *pWnd){
	CFMServiceIncomeList *pVw = (CFMServiceIncomeList *)pWnd;
	pVw->OnReportByReceiptSelect();
} 

static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMServiceIncomeList *pVw = (CFMServiceIncomeList *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMServiceIncomeList *pVw = (CFMServiceIncomeList *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMServiceIncomeList*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMServiceIncomeList*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMServiceIncomeList*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMServiceIncomeList*)pWnd)->OnDeptListDeleteItem();
} 
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMServiceIncomeList*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUncheckAllFnc(CWnd *pWnd){
	return ((CFMServiceIncomeList*)pWnd)->OnDeptListUncheckAll();
}
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMServiceIncomeList*)pWnd)->OnLockedSelect();
}
CFMServiceIncomeList::CFMServiceIncomeList(CWnd *pParent){

	m_nDlgWidth = 585;
	m_nDlgHeight = 585;
	SetDefaultValues();
}
CFMServiceIncomeList::~CFMServiceIncomeList(){
}
void CFMServiceIncomeList::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 550);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 150, 425, 545);
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
	m_wndObject.SetCheckBox(true);
	m_wndObjectLabel.Create(this, _T("Object"), 220, 90, 300, 115);
	m_wndObject.Create(this,305, 90, 425, 115); 
	m_wndReceiptListLabel.Create(this, _T("Receipt List"), 10, 120, 90, 145);
	m_wndReceiptList.Create(this,95, 120, 215, 145); 
	m_wndReportByReceipt.Create(this, _T("Report by Receipt"), 220, 120, 400, 145);
	m_wndPrint.Create(this, _T("&Print"), 265, 555, 345, 580);
	m_wndExport.Create(this, _T("&Export"), 350, 555, 430, 580);
	m_wndDeptList.Create(this,15, 175, 420, 540); 
	m_wndApproval.Create(this, _T("Approval"), 5, 555, 85, 580);

}
void CFMServiceIncomeList::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);

	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.SetCheckBox(TRUE);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

	m_wndObject.Format(2, 1, 12);
	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 40);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

	m_wndReceiptList.InsertColumn(0, _T("ID"), CFMT_NUMBER, 40);
	m_wndReceiptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 125);

}

void CFMServiceIncomeList::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	m_wndReceiptList.SetEvent(WE_LOADDATA, _OnReceiptListLoadDataFnc);
	m_wndReportByReceipt.SetEvent(WE_CLICK, _OnReportByReceiptSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndDeptList.SetEvent(WE_SELCHANGE, _OnDeptListSelectChangeFnc);
	m_wndDeptList.SetEvent(WE_LOADDATA, _OnDeptListLoadDataFnc);
	m_wndDeptList.SetEvent(WE_DBLCLICK, _OnDeptListDblClickFnc);
	m_wndDeptList.AddEvent(1, _T("Check All"), _OnDeptListCheckAllFnc);
	m_wndDeptList.AddEvent(2, _T("Uncheck All"), _OnDeptListUncheckAllFnc);
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);
	OnDeptListLoadData();
	m_wndReceiptList.EnableWindow(FALSE);

}
void CFMServiceIncomeList::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_TextEx(pDX, m_wndReceiptList.GetDlgCtrlID(), m_szReceiptListKey);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);
	DDX_Check(pDX, m_wndApproval.GetDlgCtrlID(), m_bApproval);
	DDX_Check(pDX, m_wndReportByReceipt.GetDlgCtrlID(), m_bReportbyReceipt);

}
void CFMServiceIncomeList::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_bLocked=FALSE;
	m_bReportbyReceipt=FALSE;

}

int CFMServiceIncomeList::SetMode(int nMode){
 		int nOldMode = GetMode();
 		CGuiView::SetMode(nMode);
 		CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
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

/*void CFMServiceIncomeList::OnYearChange(){
	
} */
/*void CFMServiceIncomeList::OnYearSetfocus(){
	
} */
/*void CFMServiceIncomeList::OnYearKillfocus(){
	
} */
int CFMServiceIncomeList::OnYearCheckValue(){
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
			szTemp.Format(_T("%.2d/.2d/.4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndFromDate.SetWindowText(szTemp);
		}
		dt.ParseDateTime(m_szToDate);
		date = dt.GetDate();
		if (date.GetYear() != 1752)
		{
			dt.SetDate(m_nYear, date.GetMonth(), date.GetDay());
			m_szToDate = dt.GetDateTime();
			szTemp.Format(_T("%.2d/.2d/.4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndToDate.SetWindowText(szTemp);
		}
	}
	UpdateData(FALSE);
	return 0;
}
 
void CFMServiceIncomeList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
void CFMServiceIncomeList::OnReportPeriodSelendok(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
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
		m_szFromDate.Format(_T("%.4d/0/1 00:00"), nYear);
		tmpStr.Format(_T("%.4d/10/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/2/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 17)
	{
		m_szFromDate.Format(_T("%.4d/1/1 00:00"), nYear);
		tmpStr.Format(_T("%.4d/12/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/2/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}

	UpdateData(false); 
}

/*void CFMServiceIncomeList::OnReportPeriodSetfocus(){
	
}*/
/*void CFMServiceIncomeList::OnReportPeriodKillfocus(){
	
}*/
long CFMServiceIncomeList::OnReportPeriodLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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

/*void CFMServiceIncomeList::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CFMServiceIncomeList::OnFromDateChange(){
	
} */
/*void CFMServiceIncomeList::OnFromDateSetfocus(){
	
} */
/*void CFMServiceIncomeList::OnFromDateKillfocus(){
	
} */
int CFMServiceIncomeList::OnFromDateCheckValue(){
	return 0;
}
 
/*void CFMServiceIncomeList::OnToDateChange(){
	
} */
/*void CFMServiceIncomeList::OnToDateSetfocus(){
	
} */
/*void CFMServiceIncomeList::OnToDateKillfocus(){
	
} */
int CFMServiceIncomeList::OnToDateCheckValue(){
	return 0;
}
long CFMServiceIncomeList::OnClerkLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND su_deptid = 'PTC'"));
	return pMF->LoadUserList(&m_wndClerk, m_szClerkKey, szFilter);
}
long CFMServiceIncomeList::OnObjectLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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

long CFMServiceIncomeList::OnReceiptListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND fac_user_id = '%s'"), pMF->GetCurrentUser());
	szSQL.Format(_T("SELECT fac_cash_id as idx, fac_invoiceno descr FROM fa_cash WHERE fac_invoicetype = 'R' %s ORDER BY fac_cash_id"), szWhere);
	long nCount = rs.ExecSQL(szSQL);
	m_wndReceiptList.DeleteAllItems();
	while (!rs.IsEOF())
	{
		m_wndReceiptList.AddItems(
			rs.GetValue(_T("idx")),
			rs.GetValue(_T("descr")),
			NULL);
		rs.MoveNext();
	}
	return nCount;
}

void CFMServiceIncomeList::OnReportByReceiptSelect(){
	UpdateData(true);
	if (m_bReportbyReceipt)
	{
		EnableControls(FALSE);
		m_wndReceiptList.EnableWindow(TRUE);
	}
	else
	{
		EnableControls(TRUE);
		m_wndReceiptList.EnableWindow(FALSE);
	}

}

void CFMServiceIncomeList::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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

	if (!rpt.Init(_T("Reports/HMS/HF_BANGKETHUBNDICHVU.RPT")))
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
	arrField.Add(_T("eamount"));
	arrField.Add(_T("deposit"));
	arrField.Add(_T("iamount"));
	arrField.Add(_T("inpatient_deposit_paid"));
	arrField.Add(_T("inpatient_payment"));
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

	EndWaitCursor();
	rpt.PrintPreview();
}
 
void CFMServiceIncomeList::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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

	xls.SetCellText(nCol, nRow + 3, _T("\x42\x1EA2NG K\xCA THU \x42\x1EC6NH NH\xC2N \x44\x1ECA\x43H V\x1EE4"),
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
	arrCol.Add(_T("\x43\xE1\x63 kho\x1EA3n ph\x1EA3i thu"));
	arrCol.Add(_T("\x42N t\x1EA1m g\x1EEDi"));
	arrCol.Add(_T("Vi\x1EC7n ph\xED ph\x1EA3i thu"));
	arrCol.Add(_T("Tr\xED\x63h t\x1EA1m g\x1EEDi"));
	arrCol.Add(_T("Thu th\xEAm"));
	arrCol.Add(_T("T\x1ED5ng thu"));
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
	arrField.Add(_T("eamount"));
	arrField.Add(_T("deposit"));
	arrField.Add(_T("iamount"));
	arrField.Add(_T("inpatient_deposit_paid"));
	arrField.Add(_T("inpatient_payment"));
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
	xls.Save(_T("Exports\\BANG KE THU BENH NHAN DICH VU.xls"));
}
void CFMServiceIncomeList::OnDeptListDblClick(){
	
}
 
void CFMServiceIncomeList::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
int CFMServiceIncomeList::OnDeptListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
}
 
long CFMServiceIncomeList::OnDeptListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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

void CFMServiceIncomeList::OnLockedSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}

CString CFMServiceIncomeList::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szWhere, szDepts, szInnerWhere, szWhere1, szObjects;
	szSQL.Empty();
	if (m_bReportbyReceipt)
	{
		if (m_szReceiptListKey.IsEmpty())
		{
			AfxMessageBox(_T("Select a receipt!"));
			m_wndReceiptList.SetFocus();
			return szSQL;
		}
		szWhere.Format(_T(" AND hfe_cash_id = %s AND hfe_status = 'P'"), m_szReceiptListKey);
		szWhere1.Format(_T(" AND hfe_cash_id = %s AND hfe_status IN ('P', 'R')"), m_szReceiptListKey);
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
		szWhere.Format(_T(" AND hfe_cash_id > 0 AND hfe_status = 'P' AND hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
		szWhere1.Format(_T(" AND hfe_cash_id > 0 AND hfe_status IN ('P', 'R') AND hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
		if (!szDepts.IsEmpty())
			szWhere.AppendFormat(_T(" AND hfe_deptid IN (%s)"), szDepts);
		if (!m_szClerkKey.IsEmpty())
			szWhere.AppendFormat(_T(" AND hfe_staff = '%s'"), m_szClerkKey);
		if (!szObjects.IsEmpty())
			szWhere.AppendFormat(_T(" AND hfe_objectid IN (%s)"), szObjects);	
	}

	szSQL.Format(_T(" SELECT docno, ") \
				_T("		get_patientname(docno) pname,") \
				_T("		hcr_recordno recordno,") \
				_T("        deptid, ") \
				_T("		hfe_date,") \
				_T("        SUM(patpaid)               AS eamount, ") \
				_T("        SUM(deposit)               AS deposit, ") \
				_T("        SUM(inpatient_cost)        AS iamount, ") \
				_T("        SUM(inpatient_depositpaid) AS inpatient_deposit_paid, ") \
				_T("        SUM(inpatient_payment)     AS inpatient_payment, ") \
				_T("		SUM(patpaid+inpatient_payment+deposit) AS totalamount") \
				_T(" FROM   (SELECT hfe_docno   AS docno, ") \
				_T("                hfe_deptid  AS deptid, ") \
				_T("				hfe_date,") \
				_T("                hfe_patpaid AS patpaid, ") \
				_T("                hfe_deposit AS deposit, ") \
				_T("                0           AS inpatient_cost, ") \
				_T("                0           inpatient_depositpaid, ") \
				_T("                0           inpatient_payment ") \
				_T("         FROM   hms_fee_invoice ") \
				_T("         WHERE  hfe_class IN ('E', 'X') %s") \
				_T("         UNION ALL ") \
				_T("         SELECT hfe_docno, ") \
				_T("                hfe_deptid, ") \
				_T("				hfe_date,") \
				_T("                0, ") \
				_T("                hfe_amount, ") \
				_T("                0 AS inpatient_cost, ") \
				_T("                0 inpatient_depositpaid, ") \
				_T("                0 inpatient_payment ") \
				_T("         FROM   hms_fee_deposit ") \
				_T("         WHERE  hfe_class IN( 'I', 'A' ) %s") \
				_T("         UNION ALL ") \
				_T("         SELECT hfe_docno   AS docno, ") \
				_T("                hfe_deptid  AS deptid, ") \
				_T("				hfe_date,") \
				_T("                0, ") \
				_T("                0, ") \
				_T("                hfe_amount  AS inpatient_cost, ") \
				_T("                CASE WHEN hfe_deposit > hfe_patpaid THEN hfe_patpaid ") \
				_T("                ELSE hfe_deposit ") \
				_T("                END         AS inpatient_depositpaid, ") \
				_T("                hfe_payment AS inpatient_payment ") \
				_T("         FROM   hms_fee_invoice ") \
				_T("         WHERE  hfe_class IN( 'A', 'I' ) %s) ") \
				_T(" LEFT JOIN hms_clinical_record ON (hcr_docno = docno)")\
				_T(" LEFT JOIN hms_doc ON (hd_docno = docno) WHERE 1=1") \
				_T(" GROUP  BY docno, ") \
				_T("		   hcr_recordno,") \
				_T("           deptid,") \
				_T("		   hfe_date ") \
				_T(" HAVING SUM(patpaid+inpatient_payment+deposit) > 0") \
				_T(" ORDER BY hfe_date"), szWhere, szWhere1, szWhere);_fmsg(_T("%s"), szSQL);


	return szSQL;
}

int CFMServiceIncomeList::OnDeptListCheckAll()
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

int CFMServiceIncomeList::OnDeptListUncheckAll()
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