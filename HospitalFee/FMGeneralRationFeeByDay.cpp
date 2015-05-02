#include "stdafx.h"
#include "FMGeneralRationFeeByDay.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralRationFeeByDay *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMGeneralRationFeeByDay* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralRationFeeByDay *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralRationFeeByDay *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralRationFeeByDay *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMGeneralRationFeeByDay *pVw = (CFMGeneralRationFeeByDay *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMGeneralRationFeeByDay *pVw = (CFMGeneralRationFeeByDay *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnObjectListLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralRationFeeByDay*)pWnd)->OnObjectListLoadData();
} 
static void _OnObjectListDblClickFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay*)pWnd)->OnObjectListDblClick();
} 
static void _OnObjectListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMGeneralRationFeeByDay*)pWnd)->OnObjectListSelectChange(nOldItem, nNewItem);
} 
static int _OnObjectListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMGeneralRationFeeByDay*)pWnd)->OnObjectListDeleteItem();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralRationFeeByDay*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMGeneralRationFeeByDay*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMGeneralRationFeeByDay*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMGeneralRationFeeByDay*)pWnd)->OnDeptListDeleteItem();
} 
static int _OnAddFMGeneralRationFeeByDayFnc(CWnd *pWnd){
	 return ((CFMGeneralRationFeeByDay*)pWnd)->OnAddFMGeneralRationFeeByDay();
} 
static int _OnEditFMGeneralRationFeeByDayFnc(CWnd *pWnd){
	 return ((CFMGeneralRationFeeByDay*)pWnd)->OnEditFMGeneralRationFeeByDay();
} 
static int _OnDeleteFMGeneralRationFeeByDayFnc(CWnd *pWnd){
	 return ((CFMGeneralRationFeeByDay*)pWnd)->OnDeleteFMGeneralRationFeeByDay();
} 
static int _OnSaveFMGeneralRationFeeByDayFnc(CWnd *pWnd){
	 return ((CFMGeneralRationFeeByDay*)pWnd)->OnSaveFMGeneralRationFeeByDay();
} 
static int _OnCancelFMGeneralRationFeeByDayFnc(CWnd *pWnd){
	 return ((CFMGeneralRationFeeByDay*)pWnd)->OnCancelFMGeneralRationFeeByDay();
}

static int _OnObjectListCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralRationFeeByDay*)pWnd)->OnObjectListCheckAll();
}
static int _OnObjectListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralRationFeeByDay*)pWnd)->OnObjectListUnCheckAll();
}
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralRationFeeByDay*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralRationFeeByDay*)pWnd)->OnDeptListUnCheckAll();
}

CFMGeneralRationFeeByDay::CFMGeneralRationFeeByDay(CWnd* pParent)
{
	m_nDlgWidth = 450;
	m_nDlgHeight = 530;
	SetDefaultValues();
}
CFMGeneralRationFeeByDay::~CFMGeneralRationFeeByDay()
{
}
void CFMGeneralRationFeeByDay::OnCreateComponents()
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
void CFMGeneralRationFeeByDay::OnInitializeComponents()
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
void CFMGeneralRationFeeByDay::OnSetWindowEvents()
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
void CFMGeneralRationFeeByDay::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CFMGeneralRationFeeByDay::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMGeneralRationFeeByDay::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMGeneralRationFeeByDay::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CFMGeneralRationFeeByDay::SetMode(int nMode){
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
/*void CFMGeneralRationFeeByDay::OnYearChange(){
	
} */
/*void CFMGeneralRationFeeByDay::OnYearSetfocus(){
	
} */
/*void CFMGeneralRationFeeByDay::OnYearKillfocus(){
	
} */
int CFMGeneralRationFeeByDay::OnYearCheckValue()
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
void CFMGeneralRationFeeByDay::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMGeneralRationFeeByDay::OnReportPeriodSelendok()
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
/*void CFMGeneralRationFeeByDay::OnReportPeriodSetfocus(){
	
}*/
/*void CFMGeneralRationFeeByDay::OnReportPeriodKillfocus(){
	
}*/
long CFMGeneralRationFeeByDay::OnReportPeriodLoadData()
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
/*void CFMGeneralRationFeeByDay::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMGeneralRationFeeByDay::OnFromDateChange(){
	
} */
/*void CFMGeneralRationFeeByDay::OnFromDateSetfocus(){
	
} */
/*void CFMGeneralRationFeeByDay::OnFromDateKillfocus(){
	
} */
int CFMGeneralRationFeeByDay::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMGeneralRationFeeByDay::OnToDateChange(){
	
} */
/*void CFMGeneralRationFeeByDay::OnToDateSetfocus(){
	
} */
/*void CFMGeneralRationFeeByDay::OnToDateKillfocus(){
	
} */
int CFMGeneralRationFeeByDay::OnToDateCheckValue(){
	return 0;
} 
void CFMGeneralRationFeeByDay::OnPrintSelect()
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

	if (!rpt.Init(_T("Reports/HMS/HF_TONGHOPXUATANTHEONGAY_TONGHOP.RPT")))
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

	long double nTotal = 0;
	double nCost = 0, nQtyOrder = 0;
	int nIndex = 1;

	while (!rs.IsEOF())
	{

		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("deptname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("qtyorder"), nCost);
		nQtyOrder += nCost;
		rptDetail->SetValue(_T("3"), double2str(nCost));

		rptDetail->SetValue(_T("4"), rs.GetValue(_T("price")));

		rs.GetValue(_T("amount"), nCost);
		nTotal += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.MoveNext();
	}

	if (nTotal > 0)
	{
		rpt.GetReportFooter()->SetValue(_T("TotalGroup"), _T("T\x1ED4NG:"));

		FormatCurrency(nTotal, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	}

	
	tmpStr = pMF->m_UserInfo.su_name;
	StringUpper(tmpStr, szTemp);
	rpt.GetReportFooter()->SetValue(_T("ReceiverBy"), szTemp);
	

	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
		          szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);

	EndWaitCursor();
	rpt.PrintPreview();
} 
void CFMGeneralRationFeeByDay::OnExportSelect()
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
	xls.SetColumnWidth(1, 60);
	xls.SetColumnWidth(2, 20);

	int nRow = 1;
	int nCol = 0;

	xls.SetRowHeight(6, 45);

	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 2);

	xls.SetCellMergedColumns(2, 1, 3);
	xls.SetCellMergedColumns(2, 2, 3);

	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(2, 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(2, 2, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol, nRow + 3, 3);
	xls.SetCellMergedColumns(nCol, nRow + 4, 3);

	xls.SetCellText(nCol, nRow + 3, _T("T\x1ED4NG H\x1EE2P \x58U\x1EA4T \x102N TH\x45O NG\xC0Y"),
		            FMT_TEXT | FMT_CENTER, true, 14, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);

	TranslateString(_T("Index"), tmpStr);
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Dept"), tmpStr);
	xls.SetCellText(nCol + 1, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
	xls.SetCellText(nCol + 2, nRow + 5, _T("S\x1ED1 \x78u\x1EA5t"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
	xls.SetCellText(nCol + 3, nRow + 5, _T("\x110\x1A1n gi\xE1"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
	xls.SetCellText(nCol + 4, nRow + 5, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	int nIndex = 1;
	nRow = 6;

	long double nTotal = 0;
	double nCost;
	int nCount = 0;

	while (!rs.IsEOF())
	{
		nRow++;

		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("deptname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("qtyorder"), nCost);
		xls.SetCellText(nCol + 2, nRow, double2str(nCost), FMT_NUMBER1 | FMT_WRAPING);
		xls.SetCellText(nCol + 3, nRow, rs.GetValue(_T("price")), FMT_NUMBER1 | FMT_WRAPING);
		rs.GetValue(_T("amount"), nCost);
		nTotal += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.MoveNext();
	}

	if (nTotal > 0)
	{
		nRow++;
		xls.SetCellText(nCol + 1, nRow, _T("T\x1ED4NG:"), FMT_TEXT | FMT_WRAPING, true);
		
		tmpStr.Format(_T("%.2lf"), nTotal);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\THTienAnTheoNgay.xls"));
} 
void CFMGeneralRationFeeByDay::OnObjectListDblClick(){
	
} 
void CFMGeneralRationFeeByDay::OnObjectListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMGeneralRationFeeByDay::OnObjectListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMGeneralRationFeeByDay::OnObjectListLoadData()
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
void CFMGeneralRationFeeByDay::OnDeptListDblClick(){
	
} 
void CFMGeneralRationFeeByDay::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMGeneralRationFeeByDay::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMGeneralRationFeeByDay::OnDeptListLoadData()
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
int CFMGeneralRationFeeByDay::OnAddFMGeneralRationFeeByDay(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMGeneralRationFeeByDay::OnEditFMGeneralRationFeeByDay(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMGeneralRationFeeByDay::OnDeleteFMGeneralRationFeeByDay(){
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
 		OnCancelFMGeneralRationFeeByDay();
 	}
	return 0;
}
int CFMGeneralRationFeeByDay::OnSaveFMGeneralRationFeeByDay(){
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
 		//OnFMGeneralRationFeeByDayListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CFMGeneralRationFeeByDay::OnCancelFMGeneralRationFeeByDay(){
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
int CFMGeneralRationFeeByDay::OnFMGeneralRationFeeByDayListLoadData(){
	return 0;
}

int CFMGeneralRationFeeByDay::OnDeptListCheckAll()
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

int CFMGeneralRationFeeByDay::OnDeptListUnCheckAll()
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


int CFMGeneralRationFeeByDay::OnObjectListCheckAll()
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

int CFMGeneralRationFeeByDay::OnObjectListUnCheckAll()
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

CString CFMGeneralRationFeeByDay::GetQueryString()
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


	if (!szObjects.IsEmpty())
	{
		szWhere.Format(_T(" and hd_object in(%s) "), szObjects);
	}

	if (!szDepts.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and hfo_deptid in(%s) "), szDepts);
	}


	szSQL.Format(_T(" SELECT hfo_deptid as deptid,") \
		        _T("        upper(get_departmentname(hfo_deptid)) as deptname,") \
				_T("		sum(hfol_qtyorder) qtyorder,") \
				_T("		hfol_unitprice price,") \
				_T("        sum(hfol_qtyorder*hfol_unitprice) as amount") \
				_T(" FROM hms_feefood") \
				_T(" LEFT JOIN hms_feefoodline ON(hfol_orderid=hfo_orderid)") \
				_T(" LEFT JOIN hms_doc ON(hd_docno=hfo_docno)") \
				_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
				_T(" LEFT JOIN hms_object ON(ho_id=hd_object)") \
				_T(" WHERE hfo_orderdate BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
				_T(" GROUP BY hfo_deptid, hfol_unitprice") \
				_T(" ORDER BY hfo_deptid"), m_szFromDate, m_szToDate, szWhere);

	return szSQL;
}