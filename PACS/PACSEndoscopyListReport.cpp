#include "stdafx.h"
#include "PACSEndoscopyListReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CPACSEndoscopyListReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSEndoscopyListReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CPACSEndoscopyListReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSEndoscopyListReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPACSEndoscopyListReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSEndoscopyListReport *)pWnd)->OnToDateCheckValue();
} 
static int _OnPerformRoomLoadDataFnc(CWnd *pWnd){
	return ((CPACSEndoscopyListReport *)pWnd)->OnPerformRoomLoadData();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPACSEndoscopyListReport *pVw = (CPACSEndoscopyListReport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPACSEndoscopyListReport *pVw = (CPACSEndoscopyListReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnInsuranceSelectFnc(CWnd *pWnd){
	  ((CPACSEndoscopyListReport*)pWnd)->OnInsuranceSelect();
}
static void _OnServiceSelectFnc(CWnd *pWnd){
	  ((CPACSEndoscopyListReport*)pWnd)->OnServiceSelect();
}
static void _OnPolicySelectFnc(CWnd *pWnd){
	  ((CPACSEndoscopyListReport*)pWnd)->OnPolicySelect();
}
static int _OnAddPACSEndoscopyListReportFnc(CWnd *pWnd){
	 return ((CPACSEndoscopyListReport*)pWnd)->OnAddPACSEndoscopyListReport();
} 
static int _OnEditPACSEndoscopyListReportFnc(CWnd *pWnd){
	 return ((CPACSEndoscopyListReport*)pWnd)->OnEditPACSEndoscopyListReport();
} 
static int _OnDeletePACSEndoscopyListReportFnc(CWnd *pWnd){
	 return ((CPACSEndoscopyListReport*)pWnd)->OnDeletePACSEndoscopyListReport();
} 
static int _OnSavePACSEndoscopyListReportFnc(CWnd *pWnd){
	 return ((CPACSEndoscopyListReport*)pWnd)->OnSavePACSEndoscopyListReport();
} 
static int _OnCancelPACSEndoscopyListReportFnc(CWnd *pWnd){
	 return ((CPACSEndoscopyListReport*)pWnd)->OnCancelPACSEndoscopyListReport();
} 
CPACSEndoscopyListReport::CPACSEndoscopyListReport(CWnd* pParent)
{
	m_nDlgWidth = 460;
	m_nDlgHeight = 160;
	SetDefaultValues();
}
CPACSEndoscopyListReport::~CPACSEndoscopyListReport()
{
}
void CPACSEndoscopyListReport::OnCreateComponents()
{
	m_wndReportFilter.Create(this, _T("Report Parameters"), 5, 5, 460, 150);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 100, 55);
	m_wndYear.Create(this,105, 30, 230, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Month"), 235, 30, 325, 55);
	m_wndReportPeriod.Create(this,330, 30, 455, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 100, 85);
	m_wndFromDate.Create(this,105, 60, 230, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 235, 60, 325, 85);
	m_wndToDate.Create(this,330, 60, 455, 85); 
	m_wndPerformRoom.SetCheckBox(true);
	m_wndPerformRoomLabel.Create(this, _T("Perform Room"), 10, 90, 100, 115);
	m_wndPerformRoom.Create(this,105, 90, 455, 115); 
	m_wndInsurance.Create(this, _T("Insurance"), 180, 120, 260, 145);
	m_wndService.Create(this, _T("Service"), 265, 120, 345, 145);
	m_wndPolicy.Create(this, _T("Policies"), 350, 120, 455, 145);
	m_wndExport.Create(this, _T("&Export"), 360, 155, 460, 180);
}
void CPACSEndoscopyListReport::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(FALSE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);

	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(15);

	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

	m_wndPerformRoom.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 50);
	m_wndPerformRoom.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

}
void CPACSEndoscopyListReport::OnSetWindowEvents()
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
	m_wndPerformRoom.SetEvent(WE_LOADDATA, _OnPerformRoomLoadDataFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndInsurance.SetEvent(WE_CLICK, _OnInsuranceSelectFnc);
	m_wndService.SetEvent(WE_CLICK, _OnServiceSelectFnc);
	m_wndPolicy.SetEvent(WE_CLICK, _OnPolicySelectFnc);
	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL;
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);
}
void CPACSEndoscopyListReport::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndPerformRoom.GetDlgCtrlID(), m_szPerformRoomKey);
	DDX_Radio(pDX, m_wndInsurance.GetDlgCtrlID(), m_nInsurance);
	DDX_Radio(pDX, m_wndService.GetDlgCtrlID(), m_nService);
	DDX_Radio(pDX, m_wndPolicy.GetDlgCtrlID(), m_nPolicy);
}
void CPACSEndoscopyListReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPACSEndoscopyListReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CPACSEndoscopyListReport::SetDefaultValues()
{
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szPerformRoomKey.Empty();
	m_nInsurance=0;
	m_nService=1;
	m_nPolicy=1;
}
int CPACSEndoscopyListReport::SetMode(int nMode){
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
/*void CPACSEndoscopyListReport::OnYearChange(){
	
} */
/*void CPACSEndoscopyListReport::OnYearSetfocus(){
	
} */
/*void CPACSEndoscopyListReport::OnYearKillfocus(){
	
} */
int CPACSEndoscopyListReport::OnYearCheckValue()
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
void CPACSEndoscopyListReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSEndoscopyListReport::OnReportPeriodSelendok()
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
/*void CPACSEndoscopyListReport::OnReportPeriodSetfocus(){
	
}*/
/*void CPACSEndoscopyListReport::OnReportPeriodKillfocus(){
	
}*/
long CPACSEndoscopyListReport::OnReportPeriodLoadData()
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
/*void CPACSEndoscopyListReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CPACSEndoscopyListReport::OnFromDateChange(){
	
} */
/*void CPACSEndoscopyListReport::OnFromDateSetfocus(){
	
} */
/*void CPACSEndoscopyListReport::OnFromDateKillfocus(){
	
} */
int CPACSEndoscopyListReport::OnFromDateCheckValue()
{
	return 0;
} 
/*void CPACSEndoscopyListReport::OnToDateChange(){
	
} */
/*void CPACSEndoscopyListReport::OnToDateSetfocus(){
	
} */
/*void CPACSEndoscopyListReport::OnToDateKillfocus(){
	
} */
int CPACSEndoscopyListReport::OnToDateCheckValue()
{
	return 0;
} 

int CPACSEndoscopyListReport::OnPerformRoomLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndPerformRoom.IsSearchKey() && !m_szPerformRoomKey.IsEmpty()){
		szWhere.Format(_T(" and id='%s' "), m_szPerformRoomKey);
	}
	m_wndPerformRoom.DeleteAllItems(); 
	int nCount = 0;

	szSQL.Format(_T("SELECT hrl_id as id, hrl_name as name FROM hms_roomlist WHERE hrl_deptid = 'A3' AND hrl_section = 'HA' %s ORDER BY id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndPerformRoom.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}

void CPACSEndoscopyListReport::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSEndoscopyListReport::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	CString szTemp, tmpStr;


	BeginWaitCursor();

	UpdateData(TRUE);

	szSQL = GetQueryString();

	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}

	CString szTitle;
	if (m_nInsurance == 0)
		szTitle = _T("\x44\x41NH S\xC1\x43H \x42\x1ED2I \x44\x1AF\x1EE0NG TH\x1EE6 THU\x1EACT \x42\x1EA2O HI\x1EC2M Y T\x1EBE");
	else if (m_nService == 0)
		szTitle = _T("\x44\x41NH S\xC1\x43H \x42\x1ED2I \x44\x1AF\x1EE0NG TH\x1EE6 THU\x1EACT \x44\x1ECA\x43H V\x1EE4");
	else if (m_nPolicy == 0)
		szTitle = _T("\x44\x41NH S\xC1\x43H \x42\x1ED2I \x44\x1AF\x1EE0NG TH\x1EE6 THU\x1EACT \x42\x110\x43S");

	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 3);
	xls.SetColumnWidth(1, 28);

	if (m_nInsurance == 0)
		xls.SetColumnWidth(2, 21);
	else
		xls.SetColumnWidth(2, 12);

	if (m_nService == 0)
		xls.SetColumnWidth(3, 6);
	else
		xls.SetColumnWidth(3, 12);
	

	if (m_nService == 0)
		xls.SetColumnWidth(4, 18);
	else
		xls.SetColumnWidth(4, 7);

	xls.SetColumnWidth(5, 18);
	xls.SetColumnWidth(6, 18);
	xls.SetColumnWidth(7, 18);

	xls.SetColumnWidth(8, 12);

	int nRow = 1;
	int nCol = 0;

	xls.SetRowHeight(6, 30);
	xls.SetRowHeight(7, 40);

	xls.SetCellMergedColumns(0, 1, 4);
	xls.SetCellMergedColumns(0, 2, 4);

	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	if (m_nService == 0)
	{
		xls.SetCellMergedColumns(nCol, nRow + 3, 7);
		xls.SetCellMergedColumns(nCol, nRow + 4, 7);
	}
	else
	{
		xls.SetCellMergedColumns(nCol, nRow + 3, 9);
		xls.SetCellMergedColumns(nCol, nRow + 4, 9);
	}

	xls.SetCellText(nCol, nRow + 3, szTitle, FMT_TEXT | FMT_CENTER, true, 16, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);

	TranslateString(_T("Index"), tmpStr);
	xls.SetCellMergedRows(nCol, nRow + 5, 2);
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Patient Name"), tmpStr);
	xls.SetCellMergedRows(nCol + 1, nRow + 5, 2);
	xls.SetCellText(nCol + 1, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	if (m_nInsurance == 0)
	{
		xls.SetCellMergedRows(nCol + 2, nRow + 5, 2);
		xls.SetCellText(nCol + 2, nRow + 5, _T("S\x1ED1 \x62\x1EA3o hi\x1EC3m"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	}
	else if (m_nService == 0)
	{
		TranslateString(_T("Operation"), tmpStr);
		xls.SetCellMergedRows(nCol + 2, nRow + 5, 2);
		xls.SetCellText(nCol + 2, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	}
	else if (m_nPolicy == 0)
	{
		TranslateString(_T("Object"), tmpStr);
		xls.SetCellMergedRows(nCol + 2, nRow + 5, 2);
		xls.SetCellText(nCol + 2, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
	}

	if (m_nService == 0)
	{
		xls.SetCellMergedRows(nCol + 3, nRow + 5, 2);
		xls.SetCellText(nCol + 3, nRow + 5, _T("Ph\xE2n lo\x1EA1i TT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

		xls.SetCellMergedColumns(nCol + 4, nRow + 5, 3);
		xls.SetCellText(nCol + 4, nRow + 5, _T("Th\x1EE7 thu\x1EADt vi\xEAn"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

		xls.SetCellText(nCol + 4, nRow + 6, _T("\x43h\xEDnh"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
		xls.SetCellText(nCol + 5, nRow + 6, _T("Ph\x1EE5"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
		xls.SetCellText(nCol + 6, nRow + 6, _T("Gi\xFAp vi\x1EC7\x63"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
	}
	else
	{
		TranslateString(_T("Operation"), tmpStr);
		xls.SetCellMergedRows(nCol + 3, nRow + 5, 2);
		xls.SetCellText(nCol + 3, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

		xls.SetCellMergedRows(nCol + 4, nRow + 5, 2);
		xls.SetCellText(nCol + 4, nRow + 5, _T("Ph\xE2n lo\x1EA1i TT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

		xls.SetCellMergedColumns(nCol + 5, nRow + 5, 3);
		xls.SetCellText(nCol + 5, nRow + 5, _T("Th\x1EE7 thu\x1EADt vi\xEAn"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

		xls.SetCellText(nCol + 5, nRow + 6, _T("\x43h\xEDnh"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
		xls.SetCellText(nCol + 6, nRow + 6, _T("Ph\x1EE5"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
		xls.SetCellText(nCol + 7, nRow + 6, _T("Gi\xFAp vi\x1EC7\x63"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
	}

	if (m_nInsurance == 0 || m_nPolicy == 0)
	{
		TranslateString(_T("Money"), tmpStr);
		xls.SetCellMergedRows(nCol + 8, nRow + 5, 2);
		xls.SetCellText(nCol + 8, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
	}

	nRow = 7;
	CStringArray arrItem;
	CArray<long, long> arrTotal;
	CString szOldItemID, szNewItemID, szOldEmergency, szNewEmergency;
	long nIndex = 0;


	while (!rs.IsEOF())
	{
		rs.GetValue(_T("itemid"), szNewItemID);

		if (szNewItemID != szOldItemID)
		{
			nRow++;

			rs.GetValue(_T("itemname"), tmpStr);

			xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING, true);

			szOldItemID = szNewItemID;

			szOldEmergency.Empty();
			szNewEmergency.Empty();
		}

		rs.GetValue(_T("hpc_emergency"), szNewEmergency);

		if (szNewEmergency != szOldEmergency)
		{
			rs.GetValue(_T("itemname"), tmpStr);

			if (szNewEmergency == _T("Y"))
			{
				nRow++;
				xls.SetCellText(nCol + 1, nRow, _T("\x43\x1EA5p \x63\x1EE9u"), FMT_TEXT | FMT_WRAPING, true);

				tmpStr.AppendFormat(_T(" \x63\x1EA5p \x63\x1EE9u"));
			}

			if (tmpStr.GetLength() > 0)
				arrItem.Add(tmpStr);

			if (nIndex > 0)
				arrTotal.Add(nIndex);

			szOldEmergency = szNewEmergency;
			nIndex = 0;
		}

		nRow++;
		tmpStr.Format(_T("%d"), ++nIndex);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		if (m_nInsurance == 0)
		{
			rs.GetValue(_T("cardno"), tmpStr);
			xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);
		}
		else if (m_nService == 0)
		{
			rs.GetValue(_T("itemname"), tmpStr);

			if (szNewEmergency == _T("Y"))
				szTemp = GetItemAbbreviation(tmpStr, _T("\x63\x63"));
			else
				szTemp = GetItemAbbreviation(tmpStr);

			xls.SetCellText(nCol + 2, nRow, szTemp, FMT_TEXT | FMT_WRAPING);
		}
		else if (m_nPolicy == 0)
		{
			rs.GetValue(_T("rank"), tmpStr);
			xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);
		}

		if (m_nService == 0)
		{
			xls.SetCellText(nCol + 3, nRow, _T("\x31"), FMT_INTEGER | FMT_WRAPING);
			rs.GetValue(_T("doctorname"), tmpStr);
			xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);
		}
		else
		{
			rs.GetValue(_T("itemname"), tmpStr);

			if (szNewEmergency == _T("Y"))
				szTemp = GetItemAbbreviation(tmpStr, _T("\x63\x63"));
			else
				szTemp = GetItemAbbreviation(tmpStr);

			xls.SetCellText(nCol + 3, nRow, szTemp, FMT_TEXT | FMT_WRAPING);

			xls.SetCellText(nCol + 4, nRow, _T("\x31"), FMT_INTEGER | FMT_WRAPING);
			rs.GetValue(_T("doctorname"), tmpStr);
			xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);
		}

		rs.MoveNext();
	}

	if (nIndex > 0)
	{
		arrTotal.Add(nIndex);
	}

	nRow += 2;

	TranslateString(_T("Operation"), tmpStr);
	xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING, true);

	xls.SetCellText(nCol + 2, nRow, _T("S\x1ED1 \x63\x61 soi"), FMT_TEXT | FMT_WRAPING, true);

	xls.SetCellText(nCol + 3, nRow, _T("Gi\xE1 ti\x1EC1n"), FMT_TEXT | FMT_WRAPING, true);

	TranslateString(_T("Money"), tmpStr);
	xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_WRAPING, true);


	for (int i = 0; i < arrItem.GetSize(); i++)
	{
		nRow++;
		tmpStr.Format(_T("%s"), arrItem[i]);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		tmpStr.Format(_T("%d"), arrTotal[i]);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);
	}

	CString szFileName;

	if (m_nInsurance == 0)
		szFileName.Format(_T("Exports\\DSBoiDuongTTBH.xls"));
	else if (m_nService == 0)
		szFileName.Format(_T("Exports\\DSBoiDuongTTDV.xls"));
	else if (m_nPolicy == 0)
		szFileName.Format(_T("Exports\\DSBoiDuongTTBDCS.xls"));

	xls.Save(szFileName);
	EndWaitCursor();
} 
void CPACSEndoscopyListReport::OnInsuranceSelect(){
	
}
void CPACSEndoscopyListReport::OnServiceSelect(){
	
}
void CPACSEndoscopyListReport::OnPolicySelect(){
	
}
int CPACSEndoscopyListReport::OnAddPACSEndoscopyListReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CPACSEndoscopyListReport::OnEditPACSEndoscopyListReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CPACSEndoscopyListReport::OnDeletePACSEndoscopyListReport(){
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
 		OnCancelPACSEndoscopyListReport(); 
 	}
	return 0;
}
int CPACSEndoscopyListReport::OnSavePACSEndoscopyListReport(){
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
 		//OnPACSEndoscopyListReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CPACSEndoscopyListReport::OnCancelPACSEndoscopyListReport(){
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
int CPACSEndoscopyListReport::OnPACSEndoscopyListReportListLoadData(){
	return 0;
}

CString CPACSEndoscopyListReport::GetItemAbbreviation(CString szItemName, CString szExtraWord)
{
	CStringArray arr;
	CString szTemp;
	CString szValue;
	int nIndex = 0;
	szValue.Empty();

	for (int i = 0; i < szItemName.GetLength(); i++)
	{
		if (szItemName[i] == _T(' '))
		{
			szTemp = szItemName.Mid(nIndex, i - nIndex);
			arr.Add(szTemp);
			nIndex = i + 1;
		}
	}

	if (szItemName.GetLength() > 0)
	{
		szTemp = szItemName.Mid(nIndex, szItemName.GetLength() - nIndex);
		arr.Add(szTemp);
	}

	szTemp.Empty();
	LPTSTR str = new TCHAR[1];

	for (int i = 0; i < arr.GetSize(); i++)
	{
		szTemp = arr[i];

		if (szTemp.Left(1) != _T("\x110") && szTemp.Left(1) != _T("\x111"))
		{
			UnMarkString(szTemp.Left(1), str);
			CString tmpStr = CString(str);
			szValue.AppendFormat(_T("%s"), tmpStr.Left(1));
		}
		else
			szValue.AppendFormat(_T("%s"), szTemp.Left(1));
	}

	delete[] str;

	if (szExtraWord.IsEmpty())
		return szValue;
	else
		return szValue + _T(" ") + szExtraWord;
}

CString CPACSEndoscopyListReport::GetQueryString()
{
	CString szSQL, szWhere, szPerformRoom;
	for (int i = 0; i < m_wndPerformRoom.GetItemCount(); i++)
	{
		if (m_wndPerformRoom.GetCheck(i))
		{
			m_wndPerformRoom.SetCurSel(i);
			if (!szPerformRoom.IsEmpty())
				szPerformRoom += _T(", ");
			szPerformRoom += m_wndPerformRoom.GetCurrent(0);
		}
	}
	if (m_nInsurance == 0)
		szWhere.Format(_T(" and ho_type in('I','C') "));
	else if (m_nService == 0)
		szWhere.Format(_T(" and ho_type in('S') "));
	else if (m_nPolicy == 0)
		szWhere.Format(_T(" and ho_type not in('I','C','S') "));
	if (!szPerformRoom.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpcl_perform_deptid = 'A3' AND hpcl_proomid IN (%s)"), szPerformRoom);
	szSQL.Format(_T(" select hd_docno as docno,") \
				_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("        (select distinct ss_desc from sys_sel where ss_id='hms_rank' and ss_code=cast(hp_rank as nvarchar2(15))) as rank,") \
				_T("        hc_cardno as cardno,") \
				_T("        (select distinct su_name from sys_user where su_userid=hpc_practitioner) as doctorname,") \
				_T("        hpcl_itemid as itemid,") \
				_T("        hfl_name as itemname,") \
				_T("        case when ho_type in('I','C') then 1") \
				_T("             when ho_type in('S') then 2") \
				_T("             else 3 end as objectidx,") \
				_T("        case when hpc_deptid='C1.3' then 'Y' else 'N' end as hpc_emergency,") \
                _T("        case when hpc_deptid='C1.3' then 2 else 1 end as emergencyidx") \
				_T(" from hms_patient") \
				_T(" left join hms_doc on(hd_patientno=hp_patientno)") \
				_T(" left join hms_pacsorder on(hpc_docno=hd_docno)") \
				_T(" left join hms_pacsorderline on(hpcl_docno=hpc_docno and hpcl_orderid=hpc_orderid)") \
				_T(" left join hms_fee_list on(hfl_feeid=hpcl_itemid)") \
				_T(" left join hms_object on(ho_id=hd_object)") \
				_T(" left join hms_card on(hc_patientno=hd_patientno and hc_idx=hd_cardidx)") \
				_T(" where hpcl_status='T' and hpcl_date between cast_string2timestamp('%s')") \
				_T("       and cast_string2timestamp('%s') and hpc_groupid='B3100' %s") \
				_T(" order by objectidx, itemid, emergencyidx, docno"),
				m_szFromDate, m_szToDate, szWhere);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}