#include "stdafx.h"
#include "FMGeneralInsuranceOutlayList.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceOutlayList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMGeneralInsuranceOutlayList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceOutlayList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceOutlayList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceOutlayList *)pWnd)->OnToDateCheckValue();
} 
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceOutlayList *)pWnd)->OnClerkLoadData();
}
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceOutlayList *)pWnd)->OnObjectLoadData();
}
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMGeneralInsuranceOutlayList *pVw = (CFMGeneralInsuranceOutlayList *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMGeneralInsuranceOutlayList *pVw = (CFMGeneralInsuranceOutlayList *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceOutlayList*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMGeneralInsuranceOutlayList*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMGeneralInsuranceOutlayList*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceOutlayList*)pWnd)->OnDeptListDeleteItem();
} 
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceOutlayList*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUncheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceOutlayList*)pWnd)->OnDeptListUncheckAll();
}
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMGeneralInsuranceOutlayList*)pWnd)->OnLockedSelect();
}
CFMGeneralInsuranceOutlayList::CFMGeneralInsuranceOutlayList(CWnd *pParent){

	m_nDlgWidth = 585;
	m_nDlgHeight = 585;
	SetDefaultValues();
}
CFMGeneralInsuranceOutlayList::~CFMGeneralInsuranceOutlayList(){
}
void CFMGeneralInsuranceOutlayList::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 550);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 120, 425, 545);
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
	m_wndDeptList.Create(this,15, 145, 420, 540); 
	m_wndApproval.Create(this, _T("Approval"), 5, 555, 85, 580);

}
void CFMGeneralInsuranceOutlayList::OnInitializeComponents(){
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
}

void CFMGeneralInsuranceOutlayList::OnSetWindowEvents(){
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

}
void CFMGeneralInsuranceOutlayList::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);
	DDX_Check(pDX, m_wndApproval.GetDlgCtrlID(), m_bApproval);

}
void CFMGeneralInsuranceOutlayList::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_bLocked=FALSE;

}

int CFMGeneralInsuranceOutlayList::SetMode(int nMode){
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

/*void CFMGeneralInsuranceOutlayList::OnYearChange(){
	
} */
/*void CFMGeneralInsuranceOutlayList::OnYearSetfocus(){
	
} */
/*void CFMGeneralInsuranceOutlayList::OnYearKillfocus(){
	
} */
int CFMGeneralInsuranceOutlayList::OnYearCheckValue(){
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
 
void CFMGeneralInsuranceOutlayList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
void CFMGeneralInsuranceOutlayList::OnReportPeriodSelendok(){
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
}

/*void CFMGeneralInsuranceOutlayList::OnReportPeriodSetfocus(){
	
}*/
/*void CFMGeneralInsuranceOutlayList::OnReportPeriodKillfocus(){
	
}*/
long CFMGeneralInsuranceOutlayList::OnReportPeriodLoadData(){
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

/*void CFMGeneralInsuranceOutlayList::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CFMGeneralInsuranceOutlayList::OnFromDateChange(){
	
} */
/*void CFMGeneralInsuranceOutlayList::OnFromDateSetfocus(){
	
} */
/*void CFMGeneralInsuranceOutlayList::OnFromDateKillfocus(){
	
} */
int CFMGeneralInsuranceOutlayList::OnFromDateCheckValue(){
	return 0;
}
 
/*void CFMGeneralInsuranceOutlayList::OnToDateChange(){
	
} */
/*void CFMGeneralInsuranceOutlayList::OnToDateSetfocus(){
	
} */
/*void CFMGeneralInsuranceOutlayList::OnToDateKillfocus(){
	
} */
int CFMGeneralInsuranceOutlayList::OnToDateCheckValue(){
	return 0;
}
long CFMGeneralInsuranceOutlayList::OnClerkLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND su_deptid = 'PTC'"));
	return pMF->LoadUserList(&m_wndClerk, m_szClerkKey, szFilter);
}
long CFMGeneralInsuranceOutlayList::OnObjectLoadData(){
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
void CFMGeneralInsuranceOutlayList::OnPrintSelect(){
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

	if (!rpt.Init(_T("Reports/HMS/HF_BANGKETHCHIBH.RPT")))
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
	long double nTotal[13];
	double nCost;
	int nIndex = 1;

	for (int i = 0; i < 13; i++)
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
	arrField.Add(_T("stamp_train_amount"));
	arrField.Add(_T("holiday_amount"));
	arrField.Add(_T("other_amount"));
	arrField.Add(_T("total_payment_amount"));
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();

		rs.GetValue(_T("hfe_date"), tmpStr);
		rptDetail->SetValue(_T("1"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

		for (int j = 0; j < 13; j++)
		{
			rs.GetValue(arrField.GetAt(j), nCost);
			nTotal[j] += nCost;
			FormatCurrency(nCost, tmpStr);
			szTemp.Format(_T("%d"), j+2);
			rptDetail->SetValue(szTemp, tmpStr);	
		}		

		rs.MoveNext();
	}

	if (nTotal[5] > 0)
	{
		for (int i = 0; i < 13; i++)
		{
			FormatCurrency(nTotal[i], tmpStr);
			szTemp.Format(_T("s%d"), i + 2);
			rpt.GetReportFooter()->SetValue(szTemp, tmpStr);
		}
	}

	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);

	EndWaitCursor();
	rpt.PrintPreview();
}
 
void CFMGeneralInsuranceOutlayList::OnExportSelect(){
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

	xls.SetColumnWidth(0, 10);
	xls.SetColumnWidth(1, 15);
	xls.SetColumnWidth(2, 15);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 15);
	xls.SetColumnWidth(7, 15);
	xls.SetColumnWidth(8, 15);
	xls.SetColumnWidth(9, 15);
	xls.SetColumnWidth(11, 15);
	xls.SetColumnWidth(12, 15);
	xls.SetColumnWidth(13, 15);


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
	xls.SetCellMergedColumns(2, nRow, 5);
	xls.SetCellText(0, nRow, _T("Ng\xE0y"), 528386, true);
	xls.SetCellText(1, nRow, _T("Tr\xED\x63h TG"), 528386, true);
	xls.SetCellText(2, nRow, _T("Vi\x1EC7n ph\xED ph\x1EA3i thu"), 528386, true);
	xls.SetCellText(2, nRow + 1, _T("Ti\x1EC1n \x103n"), 528386, true);
	xls.SetCellText(3, nRow + 1, _T("\x42\xF9 g\x1EA1o"), 528386, true);
	xls.SetCellText(4, nRow + 1, _T("BHYT"), 528386, true);
	xls.SetCellText(5, nRow + 1, _T("CS"), 528386, true);
	xls.SetCellText(6, nRow + 1, _T("\x43\x1ED9ng"), 528386, true);
	xls.SetCellText(7, nRow, _T("\x43\xF2n l\x1EA1i"), 528386, true);
	xls.SetCellText(8, nRow, _T("\x43hi tr\x1EA3 l\x1EA1i"), 528386, true);
	xls.SetCellText(9, nRow, _T("L\x1B0\x1A1ng"), 528386, true);
	xls.SetCellText(10, nRow, _T("T\x65m T\xE0u \x78\x65"), 528386, true);
	xls.SetCellText(11, nRow, _T("L\x1EC5 t\x1EBFt"), 528386, true);
	xls.SetCellText(12, nRow, _T("\x43hi kh\xE1\x63"), 528386, true);
	xls.SetCellText(13, nRow, _T("\x43\x1ED9ng \x63hi"), 528386, true);
	int nIndex = 1;
	nRow = 8;
	long double nTotal[13];
	double nCost;

	for (int i = 0; i < 13; i++)
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
	arrField.Add(_T("stamp_train_amount"));
	arrField.Add(_T("holiday_amount"));
	arrField.Add(_T("other_amount"));
	arrField.Add(_T("total_payment_amount"));

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("hfe_date"), tmpStr);
		xls.SetCellText(nCol, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), FMT_TEXT);

		for (int j = 0; j < 13; j++)
		{
			rs.GetValue(arrField.GetAt(j), nCost);
			nTotal[j] += nCost;
			tmpStr.Format(_T("%.2f"), nCost);
			xls.SetCellText(nCol + j + 1, nRow, tmpStr, FMT_NUMBER1);
		}
		nRow++;
		rs.MoveNext();
	}
	if (nTotal[12] > 0)
	{
		xls.SetCellText(nCol, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 10);
		for (int i = 0; i < 13; i++)
		{
			tmpStr.Format(_T("%.2f"), nTotal[i]);
			xls.SetCellText(nCol + i + 1, nRow, tmpStr, FMT_NUMBER1, true);
		}
	}
	xls.Save(_T("Exports\\BANG KE CHI BENH NHAN BH.xls"));
}
void CFMGeneralInsuranceOutlayList::OnDeptListDblClick(){
	
}
 
void CFMGeneralInsuranceOutlayList::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
int CFMGeneralInsuranceOutlayList::OnDeptListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
}
 
long CFMGeneralInsuranceOutlayList::OnDeptListLoadData(){
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

void CFMGeneralInsuranceOutlayList::OnLockedSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}

CString CFMGeneralInsuranceOutlayList::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szWhere;
	CString szDepts, szObjects;
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
	szWhere.Format(_T(" AND rf.hfe_cash_id > 0 AND rf.hfe_status = 'P' AND rf.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!szDepts.IsEmpty())
		szWhere.AppendFormat(_T(" AND rf.hfe_deptid IN (%s)"), szDepts);
	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND rf.hfe_staff = '%s'"), m_szClerkKey);
	if (!szObjects.IsEmpty())
		szWhere.AppendFormat(_T(" AND rf.hfe_objectid IN (%s)"), szObjects);
	szSQL.Format(_T(" SELECT hfe_date, ") \
				_T("        SUM(deposit_extraction)                deposit_extraction, ") \
				_T("        SUM(food_fee)                          food_fee, ") \
				_T("        SUM(pol_payment_amount)                pol_payment_amount, ") \
				_T("        SUM(ins_payment_amount)                ins_payment_amount, ") \
				_T("        SUM(total_amount)                      total_amount, ") \
				_T("        SUM(deposit_extraction - total_amount) remaining_amount, ") \
				_T("        SUM(return_amount)                     return_amount, ") \
				_T("        SUM(salary_amount)                     salary_amount, ") \
				_T("        SUM(stamp_amount+train_amount)         stamp_train_amount, ") \
				_T("        SUM(holiday_amount)                    holiday_amount, ") \
				_T("        SUM(other_amount)                      other_amount, ") \
				_T("        SUM(total_payment_amount)              total_payment_amount ") \
				_T(" FROM   (SELECT hfe_date, ") \
				_T("                SUM(deposit_extraction)                deposit_extraction, ") \
				_T("                SUM(food_fee)                          food_fee, ") \
				_T("                SUM(pol_payment_amount)                pol_payment_amount, ") \
				_T("                SUM(ins_payment_amount)                ins_payment_amount, ") \
				_T("                SUM(total_amount)                      total_amount, ") \
				_T("                SUM(deposit_extraction - total_amount) remaining_amount, ") \
				_T("                CASE WHEN SUM(deposit_extraction) > SUM(total_amount) THEN SUM(deposit_extraction - total_amount) ") \
				_T("                ELSE 0 ") \
				_T("                END                                    return_amount, ") \
				_T("                SUM(salary_amount)                     salary_amount, ") \
				_T("                SUM(train_amount)                      train_amount, ") \
				_T("                SUM(holiday_amount)                    holiday_amount, ") \
				_T("                SUM(stamp_amount)                      stamp_amount, ") \
				_T("                SUM(other_amount)                      other_amount, ") \
				_T("                SUM(deposit_extraction + salary_amount ") \
				_T("                    + train_amount + holiday_amount + stamp_amount ") \
				_T("                    + other_amount - total_amount)     total_payment_amount ") \
				_T("         FROM   (SELECT    rf.hfe_docno                     doc_no, ") \
				_T("                           Trunc(rf.hfe_date)               hfe_date, ") \
				_T("                           coalesce(hfe_deposit, 0)                     deposit_extraction, ") \
				_T("                           0                               food_fee, ") \
				_T("                           0                               pol_payment_amount, ") \
				_T("                           0                               ins_payment_amount, ") \
				_T("                           0                               total_amount, ") \
				_T("						   coalesce(rf.hfe_amount, 0) return_amount,") \
				_T("                           Coalesce(hfe_salary_amount, 0)  salary_amount, ") \
				_T("                           Coalesce(hfe_train_amount, 0)   train_amount, ") \
				_T("                           Coalesce(hfe_holiday_amount, 0) holiday_amount, ") \
				_T("                           Coalesce(hfe_stamp_amount, 0)   stamp_amount, ") \
				_T("                           Coalesce(hfe_other_amount, 0)   other_amount ") \
				_T("            FROM      hms_fee_refund rf ") \
				_T("			LEFT JOIN hms_fee_invoice i ON (i.hfe_docno = rf.hfe_docno AND i.hfe_invoiceno = rf.hfe_refidx)") \
				_T("            LEFT JOIN hms_fee_sold fs ON ( i.hfe_docno = fs.hfe_docno ") \
				_T("                                           AND i.hfe_invoiceno = fs.hfe_invoiceno ) ") \
				_T("                 WHERE     rf.hfe_class IN ( 'A', 'I' ) ") \
				_T("				 %s") \
				_T("                 UNION ALL ") \
				_T("                 SELECT    rf.hfe_docno doc_no, ") \
				_T("                           Trunc(rf.hfe_date), ") \
				_T("                           0           deposit_extraction, ") \
				_T("                           CASE WHEN f.hfe_type = 'F' THEN f.hfe_cost ") \
				_T("                           ELSE 0 ") \
				_T("                           END         food_fee, ") \
				_T("                           CASE WHEN ( f.hfe_type <> 'F' ") \
				_T("                                  AND i.hfe_objectid IN ( 1, 3, 8 ) ) THEN f.hfe_patpaid ") \
				_T("                           ELSE 0 ") \
				_T("                           END         pol_payment_amount, ") \
				_T("                           CASE WHEN ( f.hfe_type <> 'F' ") \
				_T("                                  AND i.hfe_objectid NOT IN ( 1, 3, 8 ) ) THEN  f.hfe_patpaid ELSE 0 ") \
				_T("                           END         ins_payment_amount, ") \
				_T("                           f.hfe_patpaid total_amount, ") \
				_T("						   0,") \
				_T("                           0           salary_amount, ") \
				_T("                           0           train_amount, ") \
				_T("                           0           holiday_amount, ") \
				_T("                           0           stamp_amount, ") \
				_T("                           0           other_amount ") \
				_T("            FROM      hms_fee_refund rf ") \
				_T("			LEFT JOIN hms_fee_invoice i ON (i.hfe_docno = rf.hfe_docno AND i.hfe_invoiceno = rf.hfe_refidx)") \
				_T("            LEFT JOIN hms_fee f ON ( i.hfe_docno = f.hfe_docno AND i.hfe_invoiceno = f.hfe_invoiceno ) ") \
				_T("                 WHERE     f.hfe_status = 'P' ") \
				_T("                       AND rf.hfe_class IN ( 'A', 'I' ) ") \
				_T("                       AND f.hfe_category <> 'Y' ") \
				_T("				 %s)") \
				_T("         GROUP  BY doc_no, ") \
				_T("                   hfe_date ") \
				_T("         HAVING SUM(deposit_extraction + salary_amount ") \
				_T("                    + train_amount + holiday_amount + stamp_amount ") \
				_T("                    + other_amount - total_amount) > 0) ") \
				_T(" GROUP  BY hfe_date ") \
				_T(" ORDER  BY hfe_date "), szWhere, szWhere);
_fmsg(_T("%s"), szSQL);


	return szSQL;
}

int CFMGeneralInsuranceOutlayList::OnDeptListCheckAll()
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

int CFMGeneralInsuranceOutlayList::OnDeptListUncheckAll()
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