#include "stdafx.h"
#include "FMInsuranceOutlayList.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMInsuranceOutlayList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMInsuranceOutlayList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceOutlayList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMInsuranceOutlayList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMInsuranceOutlayList *)pWnd)->OnToDateCheckValue();
} 
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceOutlayList *)pWnd)->OnClerkLoadData();
}
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceOutlayList *)pWnd)->OnObjectLoadData();
}
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMInsuranceOutlayList *pVw = (CFMInsuranceOutlayList *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMInsuranceOutlayList *pVw = (CFMInsuranceOutlayList *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceOutlayList*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMInsuranceOutlayList*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMInsuranceOutlayList*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMInsuranceOutlayList*)pWnd)->OnDeptListDeleteItem();
} 
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMInsuranceOutlayList*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUncheckAllFnc(CWnd *pWnd){
	return ((CFMInsuranceOutlayList*)pWnd)->OnDeptListUncheckAll();
}
static int _OnListLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceOutlayList*) pWnd)->OnListLoadData();
}
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMInsuranceOutlayList*)pWnd)->OnLockedSelect();
}
CFMInsuranceOutlayList::CFMInsuranceOutlayList(CWnd *pParent){

	m_nDlgWidth = 585;
	m_nDlgHeight = 585;
	SetDefaultValues();
}
CFMInsuranceOutlayList::~CFMInsuranceOutlayList(){
}
void CFMInsuranceOutlayList::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 550);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 340, 425, 545);
	m_wndPaymentSheet.Create(this, _T("Payment Sheet"), 10, 120, 425, 335);
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
	m_wndPrint.Create(this, _T("&Print"), 265, 555, 345, 580);
	m_wndExport.Create(this, _T("&Export"), 350, 555, 430, 580);
	m_wndDeptList.Create(this,15, 365, 420, 540); 
	m_wndList.Create(this,15, 145, 420, 330); 

}
void CFMInsuranceOutlayList::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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

	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 60);
	m_wndList.InsertColumn(1, _T("Name"), CFMT_TEXT, 125);
	m_wndList.InsertColumn(2, _T("Locked"), CFMT_TEXT, 80);
	m_wndList.SetCheckBox(true);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

	m_wndObject.Format(2, 1, 12);
	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 40);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);
}

void CFMInsuranceOutlayList::OnSetWindowEvents(){
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
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndDeptList.SetEvent(WE_SELCHANGE, _OnDeptListSelectChangeFnc);
	m_wndDeptList.SetEvent(WE_LOADDATA, _OnDeptListLoadDataFnc);
	m_wndDeptList.SetEvent(WE_DBLCLICK, _OnDeptListDblClickFnc);
	m_wndDeptList.AddEvent(1, _T("Check All"), _OnDeptListCheckAllFnc);
	m_wndDeptList.AddEvent(2, _T("Uncheck All"), _OnDeptListUncheckAllFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);
	OnDeptListLoadData();

}
void CFMInsuranceOutlayList::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);
	DDX_Check(pDX, m_wndApproval.GetDlgCtrlID(), m_bApproval);

}
void CFMInsuranceOutlayList::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_bLocked=FALSE;

}

int CFMInsuranceOutlayList::SetMode(int nMode){
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

/*void CFMInsuranceOutlayList::OnYearChange(){
	
} */
/*void CFMInsuranceOutlayList::OnYearSetfocus(){
	
} */
/*void CFMInsuranceOutlayList::OnYearKillfocus(){
	
} */
int CFMInsuranceOutlayList::OnYearCheckValue(){
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
 
void CFMInsuranceOutlayList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
void CFMInsuranceOutlayList::OnReportPeriodSelendok(){
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
	OnListLoadData();
}

/*void CFMInsuranceOutlayList::OnReportPeriodSetfocus(){
	
}*/
/*void CFMInsuranceOutlayList::OnReportPeriodKillfocus(){
	
}*/
long CFMInsuranceOutlayList::OnReportPeriodLoadData(){
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

/*void CFMInsuranceOutlayList::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMInsuranceOutlayList::OnFromDateChange(){
	
} */
/*void CFMInsuranceOutlayList::OnFromDateSetfocus(){
	
} */
/*void CFMInsuranceOutlayList::OnFromDateKillfocus(){
	
} */
int CFMInsuranceOutlayList::OnFromDateCheckValue(){
	OnListLoadData();
	return 0;
}
 
/*void CFMInsuranceOutlayList::OnToDateChange(){
	
} */
/*void CFMInsuranceOutlayList::OnToDateSetfocus(){
	
} */
/*void CFMInsuranceOutlayList::OnToDateKillfocus(){
	
} */
int CFMInsuranceOutlayList::OnToDateCheckValue(){
	OnListLoadData();
	return 0;
}
long CFMInsuranceOutlayList::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND su_deptid = 'PTC'"));
	return pMF->LoadUserList(&m_wndClerk, m_szClerkKey, szFilter);
}
long CFMInsuranceOutlayList::OnObjectLoadData(){
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
void CFMInsuranceOutlayList::OnPrintSelect(){
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

	if (!rpt.Init(_T("Reports/HMS/HF_BANGKECHIBNBH.RPT")))
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
	long double nTotal[14];
	double nCost;
	int nIndex = 1;

	for (int i = 0; i < 14; i++)
	{
		nTotal[i] = 0;
	}
	arrField.Add(_T("deposit_extraction"));
	arrField.Add(_T("food_fee"));
	arrField.Add(_T(""));
	arrField.Add(_T("ins_payment_amount"));
	arrField.Add(_T("pol_payment_amount"));
	arrField.Add(_T("total_amount"));
	arrField.Add(_T("remaining_amount"));
	arrField.Add(_T("return_amount"));
	arrField.Add(_T("salary_amount"));
	arrField.Add(_T("train_amount"));
	arrField.Add(_T("holiday_amount"));
	arrField.Add(_T("stamp_amount"));
	arrField.Add(_T("other_amount"));
	arrField.Add(_T("total_payment_amount"));
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("patient_name"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("rank"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("doc_no"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("record_no"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("dept_id"), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("treatment_days"), tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		for (int j = 0; j < 14; j++)
		{
			rs.GetValue(arrField.GetAt(j), nCost);
			nTotal[j] += nCost;
			FormatCurrency(nCost, tmpStr);
			szTemp.Format(_T("%d"), j+11);
			rptDetail->SetValue(szTemp, tmpStr);	
		}		

		rs.MoveNext();
	}

	if (nTotal[5] > 0)
	{
		for (int i = 0; i < 14; i++)
		{
			FormatCurrency(nTotal[i], tmpStr);
			szTemp.Format(_T("s%d"), i + 11);
			rpt.GetReportFooter()->SetValue(szTemp, tmpStr);
		}
	}

	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	szTemp.Format(_T("%2.f"), nTotal[13]);
	MoneyToString(szTemp, tmpStr);
	tmpStr += _T(" \x111\x1ED3ng.");
	rpt.GetReportFooter()->SetValue(_T("SumInWords"), tmpStr);
	EndWaitCursor();
	rpt.PrintPreview();
}
 
void CFMInsuranceOutlayList::OnExportSelect(){
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
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(5, 35);
	xls.SetColumnWidth(6, 15);
	xls.SetColumnWidth(7, 8);
	xls.SetColumnWidth(8, 15);
	xls.SetColumnWidth(9, 10);
	xls.SetColumnWidth(11, 15);
	xls.SetColumnWidth(12, 15);
	xls.SetColumnWidth(13, 15);
	xls.SetColumnWidth(14, 15);
	xls.SetColumnWidth(15, 15);
	xls.SetColumnWidth(16, 10);
	xls.SetColumnWidth(17, 10);
	xls.SetColumnWidth(18, 10);
	xls.SetColumnWidth(19, 10);
	xls.SetColumnWidth(20, 10);
	xls.SetColumnWidth(21, 15);

	int nRow = 1;
	int nCol = 0;

	xls.SetCellMergedColumns(0, 1, 4);
	xls.SetCellMergedColumns(0, 2, 4);


	xls.SetCellText(0, 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol, nRow + 3, 18);
	xls.SetCellMergedColumns(nCol, nRow + 4, 18);

	xls.SetCellText(nCol, nRow + 3, _T("\x42\x1EA2NG K\xCA TH\x41NH TO\xC1N CHI \x42\x1EC6NH NH\xC2N \x42\x1ED8 \x110\x1ED8I - \x42H"),
					FMT_TEXT | FMT_CENTER, true, 16, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);
	
	nRow = 6;
	xls.SetCellMergedColumns(9, nRow, 5);
	xls.SetCellText(0, nRow, _T("STT"), 528386, true);
	xls.SetCellText(1, nRow, _T("H\x1ECD v\xE0 t\xEAn"), 528386, true);
	xls.SetCellText(2, nRow, _T("\x43\x1EA5p \x62\x1EAD\x63"), 528386, true);
	xls.SetCellText(3, nRow, _T("S\x1ED1 HS"), 528386, true);
	xls.SetCellText(4, nRow, _T("S\x1ED1 \x42\x41"), 528386, true);
	xls.SetCellText(5, nRow, _T("\x110\x1ECB\x61 \x63h\x1EC9"), 528386, true);
	xls.SetCellText(6, nRow, _T("Kho\x61"), 528386, true);
	xls.SetCellText(7, nRow, _T("S\x1ED1 ng\xE0y"), 528386, true);
	xls.SetCellText(8, nRow, _T("Tr\xED\x63h TG"), 528386, true);
	xls.SetCellText(9, nRow, _T("Vi\x1EC7n ph\xED ph\x1EA3i thu"), 528386, true);
	xls.SetCellText(9, nRow + 1, _T("Ti\x1EC1n \x103n"), 528386, true);
	xls.SetCellText(10, nRow + 1, _T("\x42\xF9 g\x1EA1o"), 528386, true);
	xls.SetCellText(11, nRow + 1, _T("BHYT"), 528386, true);
	xls.SetCellText(12, nRow + 1, _T("CS"), 528386, true);
	xls.SetCellText(13, nRow + 1, _T("\x43\x1ED9ng"), 528386, true);
	xls.SetCellText(14, nRow, _T("\x43\xF2n l\x1EA1i"), 528386, true);
	xls.SetCellText(15, nRow, _T("\x43hi tr\x1EA3 l\x1EA1i"), 528386, true);
	xls.SetCellText(16, nRow, _T("L\x1B0\x1A1ng"), 528386, true);
	xls.SetCellText(17, nRow, _T("T\xE0u \x78\x65"), 528386, true);
	xls.SetCellText(18, nRow, _T("L\x1EC5 t\x1EBFt"), 528386, true);
	xls.SetCellText(19, nRow, _T("T\x65m th\x1B0"), 528386, true);
	xls.SetCellText(20, nRow, _T("\x43hi kh\xE1\x63"), 528386, true);
	xls.SetCellText(21, nRow, _T("\x43\x1ED9ng \x63hi"), 528386, true);
	int nIndex = 1;
	nRow = 8;
	long double nTotal[14];
	double nCost;

	for (int i = 0; i < 14; i++)
	{
		nTotal[i] = 0;
	}
	arrField.Add(_T("deposit_extraction"));
	arrField.Add(_T("food_fee"));
	arrField.Add(_T(""));
	arrField.Add(_T("ins_payment_amount"));
	arrField.Add(_T("pol_payment_payment"));
	arrField.Add(_T("total_amount"));
	arrField.Add(_T("remaining_amount"));
	arrField.Add(_T("return_amount"));
	arrField.Add(_T("salary_amount"));
	arrField.Add(_T("train_amount"));
	arrField.Add(_T("holiday_amount"));
	arrField.Add(_T("stamp_amount"));
	arrField.Add(_T("other_amount"));
	arrField.Add(_T("total_payment_amount"));

	while (!rs.IsEOF())
	{
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("patient_name"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("rank"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("doc_no"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("record_no"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("unit"), tmpStr);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("dept_id"), tmpStr);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("treatment_days"), tmpStr);
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_TEXT);

		for (int j = 0; j < arrField.GetCount(); j++)
		{
			rs.GetValue(arrField.GetAt(j), nCost);
			nTotal[j] += nCost;
			tmpStr.Format(_T("%.2f"), nCost);
			xls.SetCellText(nCol + j + 8, nRow, tmpStr, FMT_NUMBER1);
		}
		nRow++;
		rs.MoveNext();
	}
	if (nTotal[13] > 0)
	{
		xls.SetCellText(nCol + 10, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 10);
		for (int i = 0; i < 14; i++)
		{
			tmpStr.Format(_T("%.2f"), nTotal[i]);
			xls.SetCellText(nCol + i + 8, nRow, tmpStr, FMT_NUMBER1);
		}
	}
	xls.Save(_T("Exports\\BANG KE CHI BENH NHAN BH.xls"));
}

int CFMInsuranceOutlayList::OnListLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND fac_user_id = '%s' AND fac_locked = 'Y' ") \
				   _T(" AND fac_updateddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
					, pMF->GetCurrentUser(), m_szFromDate, m_szToDate);
	szSQL.Format(_T("SELECT fac_cash_id as idx, fac_invoiceno descr, fac_locked stt FROM fa_cash WHERE fac_invoicetype = 'P' %s ORDER BY fac_cash_id"), szWhere);
	long nCount = rs.ExecSQL(szSQL);
	m_wndList.BeginLoad();
	while (!rs.IsEOF())
	{
		m_wndList.AddItems(
			rs.GetValue(_T("idx")),
			rs.GetValue(_T("descr")),
			rs.GetValue(_T("stt")),
			NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad();
	return nCount;	
}

void CFMInsuranceOutlayList::OnDeptListDblClick(){
	
}
 
void CFMInsuranceOutlayList::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
int CFMInsuranceOutlayList::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
}
 
long CFMInsuranceOutlayList::OnDeptListLoadData(){
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

void CFMInsuranceOutlayList::OnLockedSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}

CString CFMInsuranceOutlayList::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szWhere;
	CString szDepts, szObjects, szPaymentStr;
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
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (!szPaymentStr.IsEmpty())
				szPaymentStr += _T(", ");
			szPaymentStr.AppendFormat(_T("%s"), m_wndList.GetItemText(i, 0));
		}
	}
	if (!szPaymentStr.IsEmpty())
	{
		szWhere.Format(_T(" AND rf.hfe_objectid <> 7 AND rf.hfe_cash_id IN (%s) AND rf.hfe_status = 'P'"), szPaymentStr);
	}

	else
	{
		szWhere.Format(_T(" AND rf.hfe_objectid <> 7 AND rf.hfe_cash_id > 0 AND rf.hfe_status = 'P' ") \
					_T(" AND rf.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
		if (!szDepts.IsEmpty())
			szWhere.AppendFormat(_T(" AND rf.hfe_deptid IN (%s)"), szDepts);
		if (!m_szClerkKey.IsEmpty())
			szWhere.AppendFormat(_T(" AND rf.hfe_staff = '%s'"), m_szClerkKey);
		if (!szObjects.IsEmpty())
			szWhere.AppendFormat(_T(" AND rf.hfe_objectid IN (%s)"), szObjects);
	}
	//Edited 1st
	//Added deposit refund fee
	szSQL.Format(_T(" SELECT    Get_patientname(doc_no) patient_name, ") \
				_T("           hfe_date, ") \
				_T("           CASE WHEN Coalesce(hp_rank, 0) > 0 THEN Get_syssel_desc('hms_rank', hp_rank) ") \
				_T("           ELSE Get_objectname(hd_object) ") \
				_T("           END rank, ") \
				_T("           doc_no, ") \
				_T("           hcr_recordno record_no, ") \
				_T("           CASE WHEN hp_workplace IS NOT NULL THEN hp_workplace ") \
				_T("           ELSE Cast(Hms_getaddress(hp_provid, hp_distid, hp_villid) AS NVARCHAR2(256)) ") \
				_T("           END unit, ") \
				_T("           dept_id, ") \
				_T("           hcr_sumtreat treatment_days, ") \
				_T("           ROUND(SUM(deposit_extraction)) deposit_extraction, ") \
				_T("           ROUND(SUM(food_fee)) food_fee, ") \
				_T("           ROUND(SUM(pol_payment_amount)) pol_payment_amount, ") \
				_T("           ROUND(SUM(ins_payment_amount)) ins_payment_amount, ") \
				_T("           ROUND(SUM(total_amount)) total_amount, ") \
				_T("           ROUND(SUM(deposit_extraction - total_amount)) remaining_amount, ") \
				_T("           ROUND(SUM(return_amount)) return_amount, ") \
				_T("           ROUND(SUM(salary_amount)) salary_amount, ") \
				_T("           ROUND(SUM(train_amount)) train_amount, ") \
				_T("           ROUND(SUM(holiday_amount)) holiday_amount, ") \
				_T("           ROUND(SUM(stamp_amount)) stamp_amount, ") \
				_T("           ROUND(SUM(other_amount)) other_amount, ") \
				_T("           ROUND(SUM(deposit_extraction + return_amount + salary_amount ") \
				_T("               + train_amount + holiday_amount + stamp_amount ") \
				_T("               + other_amount - total_amount))     total_payment_amount ") \
				_T(" FROM      (SELECT    rf.hfe_docno doc_no, ") \
				_T("					  rf.hfe_date hfe_date,") \
				_T("					  rf.hfe_deptid dept_id,") \
				_T("                      coalesce(hfe_deposit, 0) deposit_extraction, ") \
				_T("                      0 food_fee, ") \
				_T("                      0 pol_payment_amount, ") \
				_T("                      0 ins_payment_amount, ") \
				_T("                      0 total_amount, ") \
				_T("					  0 return_amount,") \
				_T("                      Coalesce(hfe_salary_amount, 0)  salary_amount, ") \
				_T("                      Coalesce(hfe_train_amount, 0)   train_amount, ") \
				_T("                      Coalesce(hfe_holiday_amount, 0) holiday_amount, ") \
				_T("                      Coalesce(hfe_stamp_amount, 0)   stamp_amount, ") \
				_T("                      Coalesce(hfe_other_amount, 0)   other_amount ") \
				_T("            FROM      hms_fee_refund rf ") \
				_T("			LEFT JOIN hms_fee_invoice i ON (i.hfe_docno = rf.hfe_docno AND i.hfe_invoiceno = rf.hfe_refidx)") \
				_T("            LEFT JOIN hms_fee_sold fs ON ( i.hfe_docno = fs.hfe_docno ") \
				_T("                                           AND i.hfe_invoiceno = fs.hfe_invoiceno ) ") \
				_T("            WHERE     rf.hfe_class IN ( 'A', 'I' ) ") \
				_T("			%s") \
				_T("            UNION ALL ") \
				_T("            SELECT    rf.hfe_docno                 doc_no, ") \
				_T("					  rf.hfe_date,") \
				_T("					  rf.hfe_deptid,") \
				_T("                      0 deposit_extraction, ") \
				_T("                      CASE WHEN f.hfe_type = 'F' THEN f.hfe_cost ") \
				_T("                      ELSE 0 ") \
				_T("                      END food_fee, ") \
				_T("                      CASE WHEN ( f.hfe_type <> 'F' ") \
				_T("                             AND i.hfe_objectid IN ( 1, 3, 8 ) ) THEN f.hfe_patpaid ELSE 0") \
				_T("                      END pol_payment_amount, ") \
				_T("                      CASE WHEN ( f.hfe_type <> 'F' ") \
				_T("                             AND i.hfe_objectid NOT IN ( 1, 3, 8 ) ) THEN f.hfe_patpaid ELSE 0") \
				_T("                      END ins_payment_amount, ") \
				_T("                      f.hfe_patpaid total_amount,") \
				_T("					  0,") \
				_T("                      0 salary_amount, ") \
				_T("                      0 train_amount, ") \
				_T("                      0 holiday_amount, ") \
				_T("                      0 stamp_amount, ") \
				_T("                      0                           other_amount ") \
				_T("            FROM      hms_fee_refund rf ") \
				_T("			LEFT JOIN hms_fee_invoice i ON (i.hfe_docno = rf.hfe_docno AND i.hfe_invoiceno = rf.hfe_refidx)") \
				_T("            LEFT JOIN hms_fee f ON ( i.hfe_docno = f.hfe_docno AND i.hfe_invoiceno = f.hfe_invoiceno ) ") \
				_T("            WHERE     f.hfe_status = 'P' ") \
				_T("                  AND rf.hfe_class IN ('A', 'I') ") \
				_T("				  AND f.hfe_category <> 'Y' AND f.hfe_feegroup NOT IN ('OPT_L', 'HITECH_L') ") \
				_T("			%s") \
				_T("            UNION ALL ") \
				_T("            SELECT    rf.hfe_docno                 doc_no, ") \
				_T("					  rf.hfe_date,") \
				_T("					  rf.hfe_deptid,") \
				_T("                      0                           deposit_extraction, ") \
				_T("                      CASE WHEN f.hfe_type = 'F' THEN f.hfe_cost ") \
				_T("                      ELSE 0 ") \
				_T("                      END                         food_fee, ") \
				_T("                      CASE WHEN ( f.hfe_type <> 'F' ") \
				_T("                             AND i.hfe_objectid IN ( 1, 3, 8 ) ) THEN f.hfe_patpaid ELSE 0") \
				_T("                      END                         pol_payment_amount, ") \
				_T("                      CASE WHEN ( f.hfe_type <> 'F' ") \
				_T("                             AND i.hfe_objectid NOT IN ( 1, 3, 8 ) ) THEN f.hfe_patpaid ELSE 0") \
				_T("                      END                         ins_payment_amount, ") \
				_T("                      f.hfe_patpaid total_amount,") \
				_T("					  0,") \
				_T("                      0 salary_amount, ") \
				_T("                      0 train_amount, ") \
				_T("                      0 holiday_amount, ") \
				_T("                      0 stamp_amount, ") \
				_T("                      0 other_amount ") \
				_T("            FROM      hms_fee_refund rf ") \
				_T("			LEFT JOIN hms_fee_invoice i ON (i.hfe_docno = rf.hfe_docno AND i.hfe_invoiceno = rf.hfe_refidx)") \
				_T("            LEFT JOIN hms_fee f ON ( i.hfe_docno = f.hfe_docno AND i.hfe_invoiceno = f.hfe_invoiceno ) ") \
				_T("            WHERE     f.hfe_status = 'P' ") \
				_T("                  AND rf.hfe_class IN ('A', 'I') ") \
				_T("				  AND f.hfe_category <> 'Y' AND f.hfe_type = 'V' ") \
				_T("			%s") \
				_T("            UNION ALL ") \
				_T("            SELECT    rf.hfe_docno                 doc_no, ") \
				_T("					  rf.hfe_date,") \
				_T("					  rf.hfe_deptid,") \
				_T("                      0 deposit_extraction, ") \
				_T("                      0 food_fee, ") \
				_T("                      0 pol_payment_amount, ") \
				_T("                      0 ins_payment_amount, ") \
				_T("                      0 total_amount,") \
				_T("					  rf.hfe_amount,") \
				_T("                      0 salary_amount, ") \
				_T("                      0 train_amount, ") \
				_T("                      0 holiday_amount, ") \
				_T("                      0 stamp_amount, ") \
				_T("                      0 other_amount ") \
				_T("            FROM      hms_fee_refund rf ") \
				_T("			LEFT JOIN hms_fee_deposit d ON (d.hfe_docno = rf.hfe_docno AND d.hfe_invoiceno = rf.hfe_refidx)") \
				_T("            WHERE     rf.hfe_class IN ('A', 'I') AND d.hfe_status = 'R' ") \
				_T("			%s)") \
				_T(" LEFT JOIN hms_doc ON (hd_docno = doc_no) ") \
				_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = doc_no ) ") \
				_T(" LEFT JOIN hms_patient ON ( hp_patientno = hd_patientno ) ") \
				_T(" GROUP     BY doc_no, ") \
				_T("              hfe_date, ") \
				_T("              hp_rank, ") \
				_T("              hcr_recordno, ") \
				_T("              hp_workplace, ") \
				_T("              dept_id, ") \
				_T("              hp_provid, ") \
				_T("              hp_distid, ") \
				_T("              hp_villid, ") \
				_T("              hd_object, ") \
				_T("              hcr_sumtreat ") \
				_T(" HAVING    SUM(deposit_extraction + return_amount + salary_amount ") \
				_T("               + train_amount + holiday_amount + stamp_amount ") \
				_T("               + other_amount - total_amount) > 0 ") \
				_T(" ORDER BY hfe_date"), szWhere, szWhere, szWhere, szWhere);
_fmsg(_T("%s"), szSQL);


	return szSQL;
}

int CFMInsuranceOutlayList::OnDeptListCheckAll()
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

int CFMInsuranceOutlayList::OnDeptListUncheckAll()
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