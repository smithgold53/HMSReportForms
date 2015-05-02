#include "stdafx.h"
#include "FMUndischargedFoodSummary.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMUndischargedFoodSummary *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMUndischargedFoodSummary* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMUndischargedFoodSummary *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMUndischargedFoodSummary *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMUndischargedFoodSummary *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMUndischargedFoodSummary *pVw = (CFMUndischargedFoodSummary *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMUndischargedFoodSummary *pVw = (CFMUndischargedFoodSummary *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnObjectListLoadDataFnc(CWnd *pWnd){
	return ((CFMUndischargedFoodSummary*)pWnd)->OnObjectListLoadData();
} 
static void _OnObjectListDblClickFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary*)pWnd)->OnObjectListDblClick();
} 
static void _OnObjectListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMUndischargedFoodSummary*)pWnd)->OnObjectListSelectChange(nOldItem, nNewItem);
} 
static int _OnObjectListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMUndischargedFoodSummary*)pWnd)->OnObjectListDeleteItem();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMUndischargedFoodSummary*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMUndischargedFoodSummary*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMUndischargedFoodSummary*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMUndischargedFoodSummary*)pWnd)->OnDeptListDeleteItem();
} 
static int _OnAddFMFoodSummaryByDayFnc(CWnd *pWnd){
	 return ((CFMUndischargedFoodSummary*)pWnd)->OnAddFMFoodSummaryByDay();
} 
static int _OnEditFMFoodSummaryByDayFnc(CWnd *pWnd){
	 return ((CFMUndischargedFoodSummary*)pWnd)->OnEditFMFoodSummaryByDay();
} 
static int _OnDeleteFMFoodSummaryByDayFnc(CWnd *pWnd){
	 return ((CFMUndischargedFoodSummary*)pWnd)->OnDeleteFMFoodSummaryByDay();
} 
static int _OnSaveFMFoodSummaryByDayFnc(CWnd *pWnd){
	 return ((CFMUndischargedFoodSummary*)pWnd)->OnSaveFMFoodSummaryByDay();
} 
static int _OnCancelFMFoodSummaryByDayFnc(CWnd *pWnd){
	 return ((CFMUndischargedFoodSummary*)pWnd)->OnCancelFMFoodSummaryByDay();
}

static int _OnObjectListCheckAllFnc(CWnd *pWnd){
	return ((CFMUndischargedFoodSummary*)pWnd)->OnObjectListCheckAll();
}
static int _OnObjectListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMUndischargedFoodSummary*)pWnd)->OnObjectListUnCheckAll();
}
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMUndischargedFoodSummary*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMUndischargedFoodSummary*)pWnd)->OnDeptListUnCheckAll();
}

CFMUndischargedFoodSummary::CFMUndischargedFoodSummary(CWnd* pParent)
{
	m_nDlgWidth = 450;
	m_nDlgHeight = 530;
	SetDefaultValues();
}
CFMUndischargedFoodSummary::~CFMUndischargedFoodSummary(){
}
void CFMUndischargedFoodSummary::OnCreateComponents()
{
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 440, 490);
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
	m_wndPrint.Create(this, _T("&Print"), 235, 495, 335, 520);
	m_wndExport.Create(this, _T("&Export"), 340, 495, 440, 520);
	m_wndObjectList.Create(this,15, 115, 430, 230); 
	m_wndDeptList.Create(this,15, 265, 430, 480); 
}
void CFMUndischargedFoodSummary::OnInitializeComponents()
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


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndObjectList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndObjectList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndObjectList.SetCheckBox(TRUE);


	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.SetCheckBox(TRUE);
}
void CFMUndischargedFoodSummary::OnSetWindowEvents()
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
}
void CFMUndischargedFoodSummary::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CFMUndischargedFoodSummary::GetDataToScreen()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMUndischargedFoodSummary::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMUndischargedFoodSummary::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CFMUndischargedFoodSummary::SetMode(int nMode){
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
/*void CFMUndischargedFoodSummary::OnYearChange(){
	
} */
/*void CFMUndischargedFoodSummary::OnYearSetfocus(){
	
} */
/*void CFMUndischargedFoodSummary::OnYearKillfocus(){
	
} */
int CFMUndischargedFoodSummary::OnYearCheckValue()
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
void CFMUndischargedFoodSummary::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMUndischargedFoodSummary::OnReportPeriodSelendok()
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
/*void CFMUndischargedFoodSummary::OnReportPeriodSetfocus(){
	
}*/
/*void CFMUndischargedFoodSummary::OnReportPeriodKillfocus(){
	
}*/
long CFMUndischargedFoodSummary::OnReportPeriodLoadData()
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
/*void CFMUndischargedFoodSummary::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMUndischargedFoodSummary::OnFromDateChange(){
	
} */
/*void CFMUndischargedFoodSummary::OnFromDateSetfocus(){
	
} */
/*void CFMUndischargedFoodSummary::OnFromDateKillfocus(){
	
} */
int CFMUndischargedFoodSummary::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMUndischargedFoodSummary::OnToDateChange(){
	
} */
/*void CFMUndischargedFoodSummary::OnToDateSetfocus(){
	
} */
/*void CFMUndischargedFoodSummary::OnToDateKillfocus(){
	
} */
int CFMUndischargedFoodSummary::OnToDateCheckValue(){
	return 0;
} 
void CFMUndischargedFoodSummary::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL;
	double nTmp = 0;
	int nIdx = 1, j = 0;
	CString szSQL, tmpStr, szReportName, szDate, szPos;
	long double nTotal[2];
	for (int i = 0; i < 2; i++)
	{
		nTotal[i] = 0;
	}
	szReportName = _T("Reports\\HMS\\HF_DANHSACHBENHNHANDACHAMANCHUARAVIEN.RPT");
	if (!rpt.Init(szReportName))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF()){
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
		CDate::Convert(m_szToDate), yyyymmdd, ddmmyyyy);
		rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	}
	while (!rs.IsEOF())
	{
		//Group Seperation
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("dept_id")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("doc_no")));
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("record_no")));
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("rank")));
			rs.GetValue(_T("admit_date"), tmpStr);
			rptDetail->SetValue(_T("6"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
			rs.GetValue(_T("quantity"), nTmp);
			nTotal[0]+= nTmp;
			rptDetail->SetValue(_T("7"), double2str(nTmp));
			rs.GetValue(_T("amount"), nTmp);
			nTotal[1]+= nTmp;
			rptDetail->SetValue(_T("8"), double2str(nTmp));
		}
		rs.MoveNext();
	}
	tmpStr = pMF->GetSysDate();
	/*szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);*/
	rpt.PrintPreview();
}
void CFMUndischargedFoodSummary::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CExcel xls;
	CString szSQL, tmpStr, szOldLine, szNewLine;
	double nTmp = 0, nGroupQty = 0, nGroupAmt = 0;
	long double nTtlQty = 0, nTtlAmt = 0;
	int nIdx = 1, nRow = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(1, 30);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 7);
	xls.SetCellMergedColumns(0, 3, 7);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr = _T("\x44S \x42N \x43H\x1EA4M \x102N \x43H\x1AF\x41 R\x41 VI\x1EC6N");
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	xls.SetCellText(0, 5, _T("STT"), 4098, true);
	xls.SetCellText(1, 5, _T("H\x1ECD t\xEAn"), 4098, true);
	xls.SetCellText(2, 5, _T("S\x1ED1 HS"), 4098, true);
	xls.SetCellText(3, 5, _T("S\x1ED1 BA"), 4098, true);
	xls.SetCellText(4, 5, _T("\x43\x1EA5p \x62\x1EAD\x63"), 4098, true);
	xls.SetCellText(5, 5, _T("Ng\xE0y v\xE0o"), 4098, true);
	xls.SetCellText(6, 5, _T("SL"), 4098, true);
	xls.SetCellText(7, 5, _T("Th\xE0nh ti\x1EC1n"), 4098, true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("dept_id"), szNewLine);
		if (szNewLine != szOldLine)
		{
			if (nGroupAmt > 0)
			{
				xls.SetCellText(1, nRow, _T("\x43\x1ED9ng"), FMT_TEXT, true);
				xls.SetCellText(6, nRow, double2str(nGroupQty), FMT_NUMBER1, true);
				xls.SetCellText(7, nRow, double2str(nGroupAmt), FMT_NUMBER1, true);
				nTtlQty += nGroupQty;
				nGroupQty = 0;
				nTtlAmt += nGroupAmt;
				nGroupAmt = 0;
				nRow++;
			}
			xls.SetCellText(1, nRow, rs.GetValue(_T("dept_name")), FMT_TEXT, true);
			szOldLine = szNewLine;
			nRow++;
		}
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("patient_name")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("doc_no")), FMT_TEXT);
		xls.SetCellText(3, nRow, rs.GetValue(_T("record_no")), FMT_TEXT);
		xls.SetCellText(4, nRow, rs.GetValue(_T("rank")), FMT_TEXT);
		tmpStr = CDate::Convert(rs.GetValue(_T("admit_date")), yyyymmdd, ddmmyyyy);
		xls.SetCellText(5, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("quantity"), nTmp);
		nGroupQty += nTmp;
		xls.SetCellText(6, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("amount"), nTmp);
		nGroupAmt += nTmp;
		xls.SetCellText(7, nRow, double2str(nTmp), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nGroupAmt > 0)
	{
		xls.SetCellText(1, nRow, _T("\x43\x1ED9ng"), FMT_TEXT, true);
		xls.SetCellText(6, nRow, double2str(nGroupQty), FMT_NUMBER1, true);
		xls.SetCellText(7, nRow, double2str(nGroupAmt), FMT_NUMBER1, true);
		nTtlQty += nGroupQty;
		nTtlAmt += nGroupAmt;
		nRow++;
	}
	if (nTtlAmt > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng \x43\x1ED9ng"), FMT_TEXT, true);
		xls.SetCellText(6, nRow, double2str(nTtlQty), FMT_NUMBER1, true);
		xls.SetCellText(7, nRow, double2str(nTtlAmt), FMT_NUMBER1, true);
		nRow++;
	}
	xls.Save(_T("Exports\\DS BN cham an chua ra vien.xls"));
} 
void CFMUndischargedFoodSummary::OnObjectListDblClick()
{
	
} 
void CFMUndischargedFoodSummary::OnObjectListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMUndischargedFoodSummary::OnObjectListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMUndischargedFoodSummary::OnObjectListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	CString szWhere;
	m_wndObjectList.BeginLoad(); 
	int nCount = 0;

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
void CFMUndischargedFoodSummary::OnDeptListDblClick(){
	
} 
void CFMUndischargedFoodSummary::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMUndischargedFoodSummary::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMUndischargedFoodSummary::OnDeptListLoadData()
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
int CFMUndischargedFoodSummary::OnAddFMFoodSummaryByDay(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMUndischargedFoodSummary::OnEditFMFoodSummaryByDay(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMUndischargedFoodSummary::OnDeleteFMFoodSummaryByDay(){
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
 		OnCancelFMFoodSummaryByDay();
 	}
	return 0;
}
int CFMUndischargedFoodSummary::OnSaveFMFoodSummaryByDay(){
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
 		//OnFMFoodSummaryByDayListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}

int CFMUndischargedFoodSummary::OnCancelFMFoodSummaryByDay(){
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
int CFMUndischargedFoodSummary::OnFMFoodSummaryByDayListLoadData(){
	return 0;
}

int CFMUndischargedFoodSummary::OnDeptListCheckAll()
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

int CFMUndischargedFoodSummary::OnDeptListUnCheckAll()
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


int CFMUndischargedFoodSummary::OnObjectListCheckAll()
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

int CFMUndischargedFoodSummary::OnObjectListUnCheckAll()
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

CString CFMUndischargedFoodSummary::GetQueryString()
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

	szWhere.Format(_T(" AND hfo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!szObjects.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and hd_object in(%s) "), szObjects);
	}

	if (!szDepts.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and hfo_deptid in(%s) "), szDepts);
	}

	
	szSQL.Format(_T(" SELECT    hfo_deptid                          dept_id, ") \
	_T("		   get_departmentname(hfo_deptid) dept_name,") \
	_T("           hfo_docno                           doc_no, ") \
	_T("		   get_patientname(hfo_docno) patient_name,") \
	_T("           hcr_recordno                        record_no, ") \
	_T("           (SELECT ss_vndesc ") \
	_T("            FROM   sys_sel ") \
	_T("            WHERE  ss_id = 'hms_rank' ") \
	_T("               AND ss_code = hd_rank)          rank, ") \
	_T("           hcr_admitdate                       admit_date, ") \
	_T("           SUM(hfol_qtyissue)                  quantity, ") \
	_T("           SUM(hfol_qtyissue * hfol_unitprice) amount ") \
	_T(" FROM      hms_feefood ") \
	_T(" LEFT JOIN hms_feefoodline ON ( hfo_orderid = hfol_orderlineid ) ") \
	_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hfo_docno ) ") \
	_T(" LEFT JOIN hms_doc ON ( hd_docno = hcr_docno ) ") \
	_T(" LEFT JOIN sys_dept ON ( sd_id = hfo_deptid ) ") \
	_T(" WHERE     hcr_status <> 'T' AND hfo_orderstatus = 'A' %s") \
	_T(" GROUP     BY hfo_deptid, ") \
	_T("              hfo_docno, ") \
	_T("              hcr_recordno, ") \
	_T("              hcr_admitdate, ") \
	_T("              hd_rank, ") \
	_T("              sd_index ") \
	_T(" ORDER     BY sd_index, ") \
	_T("              hfo_deptid, ") \
	_T("              hfo_docno"), szWhere);


	return szSQL;
}