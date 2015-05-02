#include "stdafx.h"
#include "FMOutpatientCostSynthesisbyDay.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMOutpatientCostSynthesisbyDay* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMOutpatientCostSynthesisbyDay *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMOutpatientCostSynthesisbyDay *pVw = (CFMOutpatientCostSynthesisbyDay *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMOutpatientCostSynthesisbyDay *pVw = (CFMOutpatientCostSynthesisbyDay *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddFMOutpatientCostSynthesisbyDayFnc(CWnd *pWnd){
	 return ((CFMOutpatientCostSynthesisbyDay*)pWnd)->OnAddFMOutpatientCostSynthesisbyDay();
} 
static int _OnEditFMOutpatientCostSynthesisbyDayFnc(CWnd *pWnd){
	 return ((CFMOutpatientCostSynthesisbyDay*)pWnd)->OnEditFMOutpatientCostSynthesisbyDay();
} 
static int _OnDeleteFMOutpatientCostSynthesisbyDayFnc(CWnd *pWnd){
	 return ((CFMOutpatientCostSynthesisbyDay*)pWnd)->OnDeleteFMOutpatientCostSynthesisbyDay();
} 
static int _OnSaveFMOutpatientCostSynthesisbyDayFnc(CWnd *pWnd){
	 return ((CFMOutpatientCostSynthesisbyDay*)pWnd)->OnSaveFMOutpatientCostSynthesisbyDay();
} 
static int _OnCancelFMOutpatientCostSynthesisbyDayFnc(CWnd *pWnd){
	 return ((CFMOutpatientCostSynthesisbyDay*)pWnd)->OnCancelFMOutpatientCostSynthesisbyDay();
} 
CFMOutpatientCostSynthesisbyDay::CFMOutpatientCostSynthesisbyDay(CWnd *pParent){

	m_nDlgWidth = 440;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CFMOutpatientCostSynthesisbyDay::~CFMOutpatientCostSynthesisbyDay(){
}
void CFMOutpatientCostSynthesisbyDay::OnCreateComponents()
{
	m_wndOutpatientCostSynthesis.Create(this, _T("Outpatient Cost Synthesis"), 5, 5, 430, 90);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 240, 95, 350, 120);
	m_wndExport.Create(this, _T("&Export"), 355, 95, 430, 120);

}
void CFMOutpatientCostSynthesisbyDay::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

}
void CFMOutpatientCostSynthesisbyDay::OnSetWindowEvents(){
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
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	//AddEvent(1, _T("Add	Ctrl+A"), _OnAddFMOutpatientCostSynthesisbyDayFnc, 0, 'A', VK_CONTROL);
	//AddEvent(2, _T("Edit	Ctrl+E"), _OnEditFMOutpatientCostSynthesisbyDayFnc, 0, 'E', VK_CONTROL);
	//AddEvent(3, _T("Delete	Ctrl+D"), _OnDeleteFMOutpatientCostSynthesisbyDayFnc, 0, 'D', VK_CONTROL);
	//AddEvent(4, _T("Save	Ctrl+S"), _OnSaveFMOutpatientCostSynthesisbyDayFnc, 0, 'S', VK_CONTROL);
	//AddEvent(0, _T("-"));
	//AddEvent(5, _T("Cancel	Ctrl+T"), _OnCancelFMOutpatientCostSynthesisbyDayFnc, 0, 'T', VK_CONTROL);
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);

}
void CFMOutpatientCostSynthesisbyDay::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CFMOutpatientCostSynthesisbyDay::GetDataToScreen()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMOutpatientCostSynthesisbyDay::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMOutpatientCostSynthesisbyDay::SetDefaultValues()
{

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CFMOutpatientCostSynthesisbyDay::SetMode(int nMode){
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
/*void CFMOutpatientCostSynthesisbyDay::OnYearChange(){
	
} */
/*void CFMOutpatientCostSynthesisbyDay::OnYearSetfocus(){
	
} */
/*void CFMOutpatientCostSynthesisbyDay::OnYearKillfocus(){
	
} */
int CFMOutpatientCostSynthesisbyDay::OnYearCheckValue()
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
void CFMOutpatientCostSynthesisbyDay::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMOutpatientCostSynthesisbyDay::OnReportPeriodSelendok()
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
/*void CFMOutpatientCostSynthesisbyDay::OnReportPeriodSetfocus(){
	
}*/
/*void CFMOutpatientCostSynthesisbyDay::OnReportPeriodKillfocus(){
	
}*/
long CFMOutpatientCostSynthesisbyDay::OnReportPeriodLoadData()
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
	szSQL.Format(_T("SELECT * FROM hms_period_report %s ORDER BY hpr_idx"), szWhere);
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
/*void CFMOutpatientCostSynthesisbyDay::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMOutpatientCostSynthesisbyDay::OnFromDateChange(){
	
} */
/*void CFMOutpatientCostSynthesisbyDay::OnFromDateSetfocus(){
	
} */
/*void CFMOutpatientCostSynthesisbyDay::OnFromDateKillfocus(){
	
} */
int CFMOutpatientCostSynthesisbyDay::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMOutpatientCostSynthesisbyDay::OnToDateChange(){
	
} */
/*void CFMOutpatientCostSynthesisbyDay::OnToDateSetfocus(){
	
} */
/*void CFMOutpatientCostSynthesisbyDay::OnToDateKillfocus(){
	
} */
int CFMOutpatientCostSynthesisbyDay::OnToDateCheckValue(){
	return 0;
} 
void CFMOutpatientCostSynthesisbyDay::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);

	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szSysDate, szData, szTemp;
	int nIdx = 0;
	double nAmount = 0;

	if (!rpt.Init(_T("Reports/HMS/HF_THCHIPHIKCBNGOAITRUTHEONGAY.rpt")))
		return;

	BeginWaitCursor();

	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	
	long double nTotal[24];
	for (int i = 0; i < 24; i++)
		nTotal[i] = 0;

	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection *rptDetail;
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();

		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("c1"), tmpStr);
		rptDetail->SetValue(_T("2"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

		for (int i = 0; i < 24; i++)
		{
			szData.Format(_T("c%d"), i + 2);
			szTemp.Format(_T("%d"), i + 3);
			rs.GetValue(szData, nAmount);
			nTotal[i] += nAmount;
			FormatCurrency(nAmount, tmpStr);
			rptDetail->SetValue(szTemp, tmpStr);
		}
		rs.MoveNext();
	}
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	for (int i = 0; i < 24; i++)
	{
		szTemp.Format(_T("S%d"), i+3);
		FormatCurrency(nTotal[i], tmpStr);
		rptDetail->SetValue(szTemp, tmpStr);
	}
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
		          szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);

	EndWaitCursor();
	rpt.PrintPreview();	
} 
void CFMOutpatientCostSynthesisbyDay::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	UpdateData(true);

	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp, szData;
	int nCol = 0, nRow = 0, nIdx = 0;
	double nAmount = 0;

	BeginWaitCursor();

	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}

	long double nTotal[24];
	for (int i = 0; i < 24; i++)
	{
		nTotal[i] = 0;
	}

	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	CellFormat df(&xls), hf(&xls);

	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER);

	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(3, 13);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 15);
	xls.SetColumnWidth(7, 10);
	xls.SetColumnWidth(8, 10);
	xls.SetColumnWidth(9, 10);
	xls.SetColumnWidth(10, 10);
	xls.SetColumnWidth(11, 10);
	xls.SetColumnWidth(12, 10);
	xls.SetColumnWidth(13, 10);
	xls.SetColumnWidth(14, 10);
	xls.SetColumnWidth(15, 10);
	xls.SetColumnWidth(16, 10);
	xls.SetColumnWidth(17, 10);
	xls.SetColumnWidth(18, 10);
	xls.SetColumnWidth(19, 10);
	xls.SetColumnWidth(20, 10);
	xls.SetColumnWidth(21, 10);
	xls.SetColumnWidth(22, 10);
	xls.SetColumnWidth(23, 10);
	xls.SetColumnWidth(24, 10);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellMergedColumns(nCol, nRow + 1, 4);

	xls.SetCellMergedColumns(nCol + 20, nRow, 5);
	xls.SetCellMergedColumns(nCol + 20, nRow + 1, 5);

	xls.SetCellMergedColumns(nCol, nRow + 2, 25);
	xls.SetCellMergedColumns(nCol, nRow + 3, 25);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 20, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 20, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	TranslateString(_T("Outpatient Cost Synthesis by Day"), szTemp);
	StringUpper(szTemp, tmpStr);
	xls.SetCellText(nCol, nRow + 2, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, &df);
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("Ng\xE0y kh\xE1m"));
	arrCol.Add(_T("S\x1ED1 BN"));
	arrCol.Add(_T("\x42N tr\x1EA3"));
	arrCol.Add(_T("\x42H th\x61nh to\xE1n"));
	arrCol.Add(_T("T\x1ED5ng \x63\x1ED9ng"));
	arrCol.Add(_T("Thu\x1ED1\x63-H\x43"));
	arrCol.Add(_T("H\xF3\x61 nghi\x1EC7m"));
	arrCol.Add(_T("Sinh h\xF3\x61"));
	arrCol.Add(_T("Sinh v\x1EADt"));
	arrCol.Add(_T("Vi sinh"));
	arrCol.Add(_T("X.Quang"));
	arrCol.Add(_T("\x43h\x1EE9\x63 ph\x1EADn"));
	arrCol.Add(_T("Ph\xF3ng \x78\x1EA1"));
	arrCol.Add(_T("L\xFD li\x1EC7u"));
	arrCol.Add(_T("GP \x62\x1EC7nh"));
	arrCol.Add(_T("TNT"));
	arrCol.Add(_T("Si\xEAu \xE2m m\x1EAFt"));
	arrCol.Add(_T("Soi \x64\x1EA1 \x64\xE0y"));
	arrCol.Add(_T("M\xE1u tr\x1EA3 NS"));
	arrCol.Add(_T("CP khoa r\x103ng"));
	arrCol.Add(_T("Mi\x1EC5n \x64\x1ECB\x63h"));
	arrCol.Add(_T("\x43\xF4ng kh\xE1m"));
	arrCol.Add(_T("VTTH"));
	arrCol.Add(_T("Th\x1EE7 thu\x1EADt"));
	arrCol.Add(_T("\x44VKT \x63\x61o"));

	nRow = 4;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol + i, nRow, arrCol.GetAt(i), &hf); 
	}

	nRow = 5;
	while (!rs.IsEOF())
	{
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("c1"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, CDateTime::Convert(tmpStr, yyyymmdd, ddmmyyyy), FMT_DATE);

		for (int i = 0; i < 24; i++)
		{
			szData.Format(_T("c%d"), i + 2);
			rs.GetValue(szData, nAmount);
			nTotal[i] += nAmount;
			//FormatCurrency(nAmount, tmpStr);
			tmpStr.Format(_T("%.2f"), nAmount);
			xls.SetCellText(nCol + 2 + i, nRow, tmpStr, FMT_NUMBER1);
		}

		nRow++;
		rs.MoveNext();
	}
	xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT , true, 10);

	for (int i = 0; i < 24; i++)
	{
		//FormatCurrency(nTotal[i], tmpStr);
		tmpStr.Format(_T("%.2lf"), nTotal[i]);
		xls.SetCellText(nCol + 2 + i, nRow, tmpStr, FMT_NUMBER1, true);
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\Tong Hop Chi Phi KCB Ngoai Tru theo Ngay.xls"));
} 
int CFMOutpatientCostSynthesisbyDay::OnAddFMOutpatientCostSynthesisbyDay(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMOutpatientCostSynthesisbyDay::OnEditFMOutpatientCostSynthesisbyDay(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMOutpatientCostSynthesisbyDay::OnDeleteFMOutpatientCostSynthesisbyDay(){
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
 		OnCancelFMOutpatientCostSynthesisbyDay(); 
 	}
	return 0;
}
int CFMOutpatientCostSynthesisbyDay::OnSaveFMOutpatientCostSynthesisbyDay(){
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
 		//OnFMOutpatientCostSynthesisbyDayListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMOutpatientCostSynthesisbyDay::OnCancelFMOutpatientCostSynthesisbyDay(){
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
int CFMOutpatientCostSynthesisbyDay::OnFMOutpatientCostSynthesisbyDayListLoadData(){
	return 0;
}
CString CFMOutpatientCostSynthesisbyDay::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	szWhere.Empty();
	
	szSQL.Format(_T(" SELECT recvdate as c1,") \
				_T("        count(docno) as c2,") \
				_T("        round(sum(bntra)) as c3,") \
				_T("        round(sum(bhtra)) as c4,") \
				_T("        round(sum(tongcf)) as c5,") \
				_T("        round(sum(phithuoc)) as c6,") \
				_T("        round(sum(hoanghiem)) as c7,") \
				_T("        round(sum(sinhhoa)) as c8,") \
				_T("        0 as c9,") \
				_T("        round(sum(visinh)) as c10,") \
				_T("        round(sum(xquang)) as c11,") \
				_T("        round(sum(chucphan)) as c12,") \
				_T("        round(sum(phongxa)) as c13,") \
				_T("        round(sum(lylieu)) as c14,") \
				_T("        round(sum(giaiphaubenh)) as c15,") \
				_T("        round(sum(thannhantao)) as c16,") \
				_T("        round(sum(sieuammat)) as c17,") \
				_T("        round(sum(soidaday)) as c18,") \
				_T("        round(sum(mau)) as c19,") \
				_T("        round(sum(cfkhoarang)) as c20,") \
				_T("        round(sum(miendich)) as c21,") \
				_T("        round(sum(congkham)) as c22,") \
				_T("        round(sum(vtth)) as c23,") \
				_T("        round(sum(thuthuat)) as c24,") \
				_T("        round(sum(dvktcao)) as c25") \
				_T(" FROM") \
				_T(" (") \
				_T("   SELECT recvdate,") \
				_T("          docno,") \
				_T("          sum(tongcf-bhtra) as bntra,") \
				_T("          sum(bhtra) as bhtra,") \
				_T("          sum(tongcf) as tongcf,") \
				_T("          sum(phithuoc) as phithuoc,") \
				_T("          sum(hoanghiem) as hoanghiem,") \
				_T("          sum(sinhhoa) as sinhhoa,") \
				_T("          sum(visinh) as visinh,") \
				_T("          sum(xquang) as xquang,") \
				_T("          sum(chucphan) as chucphan,") \
				_T("          sum(phongxa) as phongxa,") \
				_T("          sum(lylieu) as lylieu,") \
				_T("          sum(giaiphaubenh) as giaiphaubenh,") \
				_T("          sum(thannhantao) as thannhantao,") \
				_T("          sum(sieuammat) as sieuammat,") \
				_T("          sum(soidaday) as soidaday,") \
				_T("          sum(mau) as mau,") \
				_T("          sum(cfkhoarang) as cfkhoarang,") \
				_T("          sum(miendich) as miendich,") \
				_T("          sum(congkham) as congkham,") \
				_T("          sum(thuthuat) as thuthuat,") \
				_T("          sum(vtth) as vtth,") \
				_T("          sum(dvktcao) AS dvktcao") \
				_T("   FROM") \
				_T("   (") \
				_T("    SELECT Cast_Timestamp2Date(fi.hfe_date) as recvdate,") \
				_T("           fe.hfe_docno as docno,") \
				_T("           fi.hfe_invoiceno as invoiceno,") \
				_T("           case when fe.hfe_patpaid > 0 then fe.hfe_patpaid else 0 end as bntra,") \
				_T("           fe.hfe_inspaid as bhtra,") \
				_T("           fe.hfe_cost as tongcf,") \
				_T("           case when substr(hfe_group,1,2) in('A1','A3') then fe.hfe_inspaid else 0 end as phithuoc,") \
				_T("           case when hfe_group='B1100' then fe.hfe_inspaid else 0 end as hoanghiem,") \
				_T("           case when hfe_group in('B1A00','B1200','B1300','B1600','B1900') then fe.hfe_inspaid else 0 end as sinhhoa,") \
				_T("           case when hfe_group='B1500' then fe.hfe_inspaid else 0 end as visinh,") \
				_T("           case when hfe_group in('B2100','B2200','B2300') then fe.hfe_inspaid else 0 end as xquang,") \
				_T("           case when hfe_group in('B2400','B3300','B3400','B3500') then fe.hfe_inspaid else 0 end as chucphan,") \
				_T("           case when hfe_group='B1700' then fe.hfe_inspaid else 0 end as phongxa,") \
				_T("           case when hfe_group='B5000' and fe.hfe_deptid='C6' then fe.hfe_inspaid else 0 end as lylieu,") \
				_T("           case when hfe_group='B1E00' then fe.hfe_inspaid else 0 end as giaiphaubenh,") \
				_T("           case when substr(hfe_group,6,2) in('F2') then fe.hfe_inspaid else 0 end as thannhantao,") \
				_T("           case when substr(hfe_group,6,2) in('D3') then fe.hfe_inspaid else 0 end as sieuammat,") \
				_T("           case when hfe_group='B3100' then fe.hfe_inspaid else 0 end as soidaday,") \
				_T("           case when hfe_group='B1G00' then fe.hfe_inspaid else 0 end as mau,") \
				_T("           case when hfe_group='B5000' and fe.hfe_deptid='B10' then fe.hfe_inspaid else 0 end as cfkhoarang,") \
				_T("           case when hfe_group='B1400' then fe.hfe_inspaid else 0 end as miendich,") \
				_T("           case when hfe_group='D0000' then fe.hfe_inspaid else 0 end as congkham,") \
				_T("           case when hfe_group='A9000' then fe.hfe_inspaid else 0 end as vtth,") \
				_T("           case when hfe_group='B5000' and fe.hfe_deptid not in('C6','B10') then fe.hfe_inspaid else 0 end as thuthuat,") \
				_T("           case when hfe_hitech='Y' then fe.hfe_inspaid else 0 end as dvktcao") \
				_T("     FROM hms_fee_invoice fi") \
				_T("     LEFT JOIN hms_fee_view fe ON(fe.hfe_docno=fi.hfe_docno AND fe.hfe_invoiceno=fi.hfe_invoiceno)") \
				_T("     LEFT JOIN hms_doc ON(hd_docno=fi.hfe_docno)") \
				_T("     LEFT JOIN hms_object ON(ho_id=hd_object)") \
				_T("     WHERE fi.hfe_date BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
				_T("           AND fe.hfe_inspaid > 0 AND fi.hfe_class in('E')") \
				_T("  ) feeTbl2") \
				_T("   GROUP BY recvdate, docno") \
				_T(" ) feeTbl3") \
				_T(" GROUP BY recvdate") \
				_T(" ORDER BY recvdate"),
				m_szFromDate, m_szToDate,
				szWhere);
	return szSQL;
}