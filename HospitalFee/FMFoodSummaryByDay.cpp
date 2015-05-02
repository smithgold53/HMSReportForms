#include "stdafx.h"
#include "FMFoodSummaryByDay.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMFoodSummaryByDay *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMFoodSummaryByDay* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMFoodSummaryByDay *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMFoodSummaryByDay *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMFoodSummaryByDay *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMFoodSummaryByDay *pVw = (CFMFoodSummaryByDay *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMFoodSummaryByDay *pVw = (CFMFoodSummaryByDay *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnObjectListLoadDataFnc(CWnd *pWnd){
	return ((CFMFoodSummaryByDay*)pWnd)->OnObjectListLoadData();
} 
static void _OnObjectListDblClickFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay*)pWnd)->OnObjectListDblClick();
} 
static void _OnObjectListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMFoodSummaryByDay*)pWnd)->OnObjectListSelectChange(nOldItem, nNewItem);
} 
static int _OnObjectListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMFoodSummaryByDay*)pWnd)->OnObjectListDeleteItem();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMFoodSummaryByDay*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMFoodSummaryByDay*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMFoodSummaryByDay*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMFoodSummaryByDay*)pWnd)->OnDeptListDeleteItem();
} 
static int _OnAddFMFoodSummaryByDayFnc(CWnd *pWnd){
	 return ((CFMFoodSummaryByDay*)pWnd)->OnAddFMFoodSummaryByDay();
} 
static int _OnEditFMFoodSummaryByDayFnc(CWnd *pWnd){
	 return ((CFMFoodSummaryByDay*)pWnd)->OnEditFMFoodSummaryByDay();
} 
static int _OnDeleteFMFoodSummaryByDayFnc(CWnd *pWnd){
	 return ((CFMFoodSummaryByDay*)pWnd)->OnDeleteFMFoodSummaryByDay();
} 
static int _OnSaveFMFoodSummaryByDayFnc(CWnd *pWnd){
	 return ((CFMFoodSummaryByDay*)pWnd)->OnSaveFMFoodSummaryByDay();
} 
static int _OnCancelFMFoodSummaryByDayFnc(CWnd *pWnd){
	 return ((CFMFoodSummaryByDay*)pWnd)->OnCancelFMFoodSummaryByDay();
}

static int _OnObjectListCheckAllFnc(CWnd *pWnd){
	return ((CFMFoodSummaryByDay*)pWnd)->OnObjectListCheckAll();
}
static int _OnObjectListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMFoodSummaryByDay*)pWnd)->OnObjectListUnCheckAll();
}
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMFoodSummaryByDay*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMFoodSummaryByDay*)pWnd)->OnDeptListUnCheckAll();
}

CFMFoodSummaryByDay::CFMFoodSummaryByDay(CWnd* pParent)
{
	m_nDlgWidth = 450;
	m_nDlgHeight = 530;
	SetDefaultValues();
}
CFMFoodSummaryByDay::~CFMFoodSummaryByDay(){
}
void CFMFoodSummaryByDay::OnCreateComponents()
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
void CFMFoodSummaryByDay::OnInitializeComponents()
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
void CFMFoodSummaryByDay::OnSetWindowEvents()
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
void CFMFoodSummaryByDay::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CFMFoodSummaryByDay::GetDataToScreen()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMFoodSummaryByDay::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMFoodSummaryByDay::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CFMFoodSummaryByDay::SetMode(int nMode){
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
/*void CFMFoodSummaryByDay::OnYearChange(){
	
} */
/*void CFMFoodSummaryByDay::OnYearSetfocus(){
	
} */
/*void CFMFoodSummaryByDay::OnYearKillfocus(){
	
} */
int CFMFoodSummaryByDay::OnYearCheckValue()
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
void CFMFoodSummaryByDay::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMFoodSummaryByDay::OnReportPeriodSelendok()
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
/*void CFMFoodSummaryByDay::OnReportPeriodSetfocus(){
	
}*/
/*void CFMFoodSummaryByDay::OnReportPeriodKillfocus(){
	
}*/
long CFMFoodSummaryByDay::OnReportPeriodLoadData()
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
/*void CFMFoodSummaryByDay::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMFoodSummaryByDay::OnFromDateChange(){
	
} */
/*void CFMFoodSummaryByDay::OnFromDateSetfocus(){
	
} */
/*void CFMFoodSummaryByDay::OnFromDateKillfocus(){
	
} */
int CFMFoodSummaryByDay::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMFoodSummaryByDay::OnToDateChange(){
	
} */
/*void CFMFoodSummaryByDay::OnToDateSetfocus(){
	
} */
/*void CFMFoodSummaryByDay::OnToDateKillfocus(){
	
} */
int CFMFoodSummaryByDay::OnToDateCheckValue(){
	return 0;
} 
void CFMFoodSummaryByDay::OnPrintSelect()
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

	if (!rpt.Init(_T("Reports/HMS/HF_BANGTONGHOPCHAMANTHEONGAY.RPT")))
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

	long nTotal[4];
	long nTemp;
	int nIndex = 1;

	for (int i = 0; i < 4; i++)
		nTotal[i] = 0;

	while (!rs.IsEOF())
	{

		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("deptname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("sang"), nTemp);
		nTotal[0] += nTemp;
		tmpStr.Format(_T("%ld"), nTemp);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("trua"), nTemp);
		nTotal[1] += nTemp;
		tmpStr.Format(_T("%ld"), nTemp);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("chieu"), nTemp);
		nTotal[2] += nTemp;
		tmpStr.Format(_T("%ld"), nTemp);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("tong"), nTemp);
		nTotal[3] += nTemp;
		tmpStr.Format(_T("%ld"), nTemp);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.MoveNext();
	}

	if (nTotal[3] > 0)
	{
		//rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rpt.GetReportFooter()->SetValue(_T("TotalGroup"), _T("T\x1ED4NG:"));
		for (int i = 0; i < 4; i++)
		{
			szTemp.Format(_T("s%d"), i + 3);
			tmpStr.Format(_T("%ld"), nTotal[i]);
			rpt.GetReportFooter()->SetValue(szTemp, tmpStr);
		}
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
void CFMFoodSummaryByDay::OnExportSelect()
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
	xls.SetColumnWidth(1, 40);
	xls.SetColumnWidth(2, 12);
	xls.SetColumnWidth(3, 12);
	xls.SetColumnWidth(4, 12);

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

	xls.SetCellMergedColumns(nCol, nRow + 3, 5);
	xls.SetCellMergedColumns(nCol, nRow + 4, 5);

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

	xls.SetCellText(nCol + 2, nRow + 5, _T("S\xE1ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 3, nRow + 5, _T("Tr\x1B0\x61"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 4, nRow + 5, _T("\x43hi\x1EC1u"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 5, nRow + 5, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	int nIndex = 1;
	nRow = 6;

	long nTotal[4];
	long nTemp;

	for (int i = 0; i < 4; i++)
		nTotal[i] = 0;

	while (!rs.IsEOF())
	{
		nRow++;

		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("deptname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("sang"), nTemp);
		nTotal[0] += nTemp;
		tmpStr.Format(_T("%ld"), nTemp);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("trua"), nTemp);
		nTotal[1] += nTemp;
		tmpStr.Format(_T("%ld"), nTemp);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("chieu"), nTemp);
		nTotal[2] += nTemp;
		tmpStr.Format(_T("%ld"), nTemp);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("tong"), nTemp);
		nTotal[3] += nTemp;
		tmpStr.Format(_T("%ld"), nTemp);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.MoveNext();
	}

	if (nTotal[3] > 0)
	{
		nRow++;
		xls.SetCellText(nCol + 1, nRow, _T("T\x1ED4NG:"), FMT_TEXT | FMT_WRAPING, true);
		
		for (int i = 0; i < 4; i++)
		{
			tmpStr.Format(_T("%ld"), nTotal[i]);
			xls.SetCellText(nCol + 2 + i, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
		}
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\THXuatAnTheoNgay.xls"));
} 
void CFMFoodSummaryByDay::OnObjectListDblClick()
{
	
} 
void CFMFoodSummaryByDay::OnObjectListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMFoodSummaryByDay::OnObjectListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMFoodSummaryByDay::OnObjectListLoadData()
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
void CFMFoodSummaryByDay::OnDeptListDblClick(){
	
} 
void CFMFoodSummaryByDay::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMFoodSummaryByDay::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMFoodSummaryByDay::OnDeptListLoadData()
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
int CFMFoodSummaryByDay::OnAddFMFoodSummaryByDay(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMFoodSummaryByDay::OnEditFMFoodSummaryByDay(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMFoodSummaryByDay::OnDeleteFMFoodSummaryByDay(){
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
int CFMFoodSummaryByDay::OnSaveFMFoodSummaryByDay(){
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

int CFMFoodSummaryByDay::OnCancelFMFoodSummaryByDay(){
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
int CFMFoodSummaryByDay::OnFMFoodSummaryByDayListLoadData(){
	return 0;
}

int CFMFoodSummaryByDay::OnDeptListCheckAll()
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

int CFMFoodSummaryByDay::OnDeptListUnCheckAll()
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


int CFMFoodSummaryByDay::OnObjectListCheckAll()
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

int CFMFoodSummaryByDay::OnObjectListUnCheckAll()
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

CString CFMFoodSummaryByDay::GetQueryString()
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


	szSQL.Format(_T(" SELECT deptid, ") \
		_T("        deptname, ") \
		_T("        SUM(sang)                AS sang, ") \
		_T("        SUM(trua)                AS trua, ") \
		_T("        SUM(chieu)               AS chieu, ") \
		_T("        SUM(sang + trua + chieu) AS tong ") \
		_T(" FROM   (SELECT    hfo_deptid                            AS deptid, ") \
		_T("                   Upper(Get_departmentname(hfo_deptid)) AS deptname, ") \
		_T("                   CASE WHEN hfol_type = 'S' THEN SUM(hfol_qtyissue) ") \
		_T("                   ELSE 0 ") \
		_T("                   END                                   AS sang, ") \
		_T("                   CASE WHEN hfol_type = 'T' THEN SUM(hfol_qtyissue) ") \
		_T("                   ELSE 0 ") \
		_T("                   END                                   AS trua, ") \
		_T("                   CASE WHEN hfol_type = 'C' THEN SUM(hfol_qtyissue) ") \
		_T("                   ELSE 0 ") \
		_T("                   END                                   AS chieu, ") \
		_T("                   SUM(hfol_qtyissue)                    AS tong ") \
		_T("         FROM      hms_feefood ") \
		_T("         LEFT JOIN hms_feefoodline ON( hfol_orderid = hfo_orderid ) ") \
		_T("         LEFT JOIN hms_doc ON( hd_docno = hfo_docno ) ") \
		_T("         LEFT JOIN hms_object ON( ho_id = hd_object ) ") \
		_T("         WHERE     hfo_orderstatus = 'A' %s ") \
		_T("         GROUP     BY hfo_deptid, ") \
		_T("                      hfol_type ") \
		_T("         HAVING    SUM(hfol_qtyissue) > 0 ") \
		_T("         ORDER     BY hfo_deptid) Tbl ") \
		_T(" GROUP  BY deptid, ") \
		_T("           deptname "), szWhere);


	return szSQL;
}