#include "stdafx.h"
#include "HMSInOutPatientTrack.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CHMSInOutPatientTrack *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CHMSInOutPatientTrack* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CHMSInOutPatientTrack *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CHMSInOutPatientTrack *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CHMSInOutPatientTrack *)pWnd)->OnToDateCheckValue();
} 
static void _OnDeptSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CHMSInOutPatientTrack* )pWnd)->OnDeptSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDeptSelendokFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnDeptSelendok();
}
/*static void _OnDeptSetfocusFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnDeptKillfocus();
}*/
/*static void _OnDeptKillfocusFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnDeptKillfocus();
}*/
static long _OnDeptLoadDataFnc(CWnd *pWnd){
	return ((CHMSInOutPatientTrack *)pWnd)->OnDeptLoadData();
}
/*static void _OnDeptAddNewFnc(CWnd *pWnd){
	((CHMSInOutPatientTrack *)pWnd)->OnDeptAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CHMSInOutPatientTrack *pVw = (CHMSInOutPatientTrack *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CHMSInOutPatientTrack *pVw = (CHMSInOutPatientTrack *)pWnd;
	pVw->OnPrintSelect();
} 
CHMSInOutPatientTrack::CHMSInOutPatientTrack(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CHMSInOutPatientTrack::~CHMSInOutPatientTrack(){
}
void CHMSInOutPatientTrack::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 510, 120);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 110, 55);
	m_wndYear.Create(this,115, 30, 255, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 260, 30, 360, 55);
	m_wndReportPeriod.Create(this,365, 30, 505, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 110, 85);
	m_wndFromDate.Create(this,115, 60, 255, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 260, 60, 360, 85);
	m_wndToDate.Create(this,365, 60, 505, 85); 
	m_wndDeptLabel.Create(this, _T("Dept"), 10, 90, 110, 115);
	m_wndDept.Create(this,115, 90, 505, 115); 
	m_wndPrintPreview.Create(this, _T("&Print"), 345, 125, 425, 150);
	m_wndPrint.Create(this, _T("&Export"), 430, 125, 510, 150);

}
void CHMSInOutPatientTrack::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndDept.SetCheckValue(true);
	m_wndDept.LimitText(1024);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Name"), CFMT_TEXT, 100);

	m_wndDept.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndDept.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);

}
void CHMSInOutPatientTrack::OnSetWindowEvents(){
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
	m_wndDept.SetEvent(WE_SELENDOK, _OnDeptSelendokFnc);
	//m_wndDept.SetEvent(WE_SETFOCUS, _OnDeptSetfocusFnc);
	//m_wndDept.SetEvent(WE_KILLFOCUS, _OnDeptKillfocusFnc);
	m_wndDept.SetEvent(WE_SELCHANGE, _OnDeptSelectChangeFnc);
	m_wndDept.SetEvent(WE_LOADDATA, _OnDeptLoadDataFnc);
	//m_wndDept.SetEvent(WE_ADDNEW, _OnDeptAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);	

}
void CHMSInOutPatientTrack::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDept.GetDlgCtrlID(), m_szDeptKey);

}
void CHMSInOutPatientTrack::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDeptKey.Empty();

}
int CHMSInOutPatientTrack::SetMode(int nMode){
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
/*void CHMSInOutPatientTrack::OnYearChange(){
	
} */
/*void CHMSInOutPatientTrack::OnYearSetfocus(){
	
} */
/*void CHMSInOutPatientTrack::OnYearKillfocus(){
	
} */
int CHMSInOutPatientTrack::OnYearCheckValue(){
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
void CHMSInOutPatientTrack::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CHMSInOutPatientTrack::OnReportPeriodSelendok(){
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
/*void CHMSInOutPatientTrack::OnReportPeriodSetfocus(){
	
}*/
/*void CHMSInOutPatientTrack::OnReportPeriodKillfocus(){
	
}*/
long CHMSInOutPatientTrack::OnReportPeriodLoadData(){
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
/*void CHMSInOutPatientTrack::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CHMSInOutPatientTrack::OnFromDateChange(){
	
} */
/*void CHMSInOutPatientTrack::OnFromDateSetfocus(){
	
} */
/*void CHMSInOutPatientTrack::OnFromDateKillfocus(){
	
} */
int CHMSInOutPatientTrack::OnFromDateCheckValue(){
	return 0;
} 
/*void CHMSInOutPatientTrack::OnToDateChange(){
	
} */
/*void CHMSInOutPatientTrack::OnToDateSetfocus(){
	
} */
/*void CHMSInOutPatientTrack::OnToDateKillfocus(){
	
} */
int CHMSInOutPatientTrack::OnToDateCheckValue(){
	return 0;
} 
void CHMSInOutPatientTrack::OnDeptSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CHMSInOutPatientTrack::OnDeptSelendok(){
	 
}
/*void CHMSInOutPatientTrack::OnDeptSetfocus(){
	
}*/
/*void CHMSInOutPatientTrack::OnDeptKillfocus(){
	
}*/
long CHMSInOutPatientTrack::OnDeptLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDept.IsSearchKey() && !m_szDeptKey.IsEmpty()){
		szWhere.Format(_T(" and sd_id='%s' "), m_szDeptKey);
	}
	m_wndDept.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT sd_id as id, sd_name as name FROM sys_dept WHERE 1=1 %s ORDER BY id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDept.AddItems(
			rs.GetValue(_T("id")),
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CHMSInOutPatientTrack::OnDeptAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CHMSInOutPatientTrack::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CHMSInOutPatientTrack::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	if (m_szDeptKey.IsEmpty())
		return;
	CRecord rs(&pMF->m_db);
	CExcel xls;
	CDate dte;
	//init with m_szFromDate
	dte.ParseDate(m_szFromDate);
	CString szSQL, tmpStr;
	int nTmp = 0, nPrev_Qty = 0, nRemain = 0, nRemain2 = 0;
	int nDataCol[14];
	int nRow = 0;
	szSQL = GetQuery_PrevQty();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 7);
	xls.SetCellMergedColumns(0, 3, 7);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr = _T("TH\x45O \x44\xD5I \x42N R\x41 V\xC0O VI\x1EC6N");
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	xls.SetColumnWidth(0, 6);
	xls.SetColumnWidth(1, 6);
	xls.SetColumnWidth(2, 6);
	xls.SetColumnWidth(3, 6);
	xls.SetColumnWidth(4, 6);
	xls.SetColumnWidth(5, 6);
	xls.SetColumnWidth(6, 6);
	xls.SetColumnWidth(7, 6);
	xls.SetColumnWidth(8, 6);
	xls.SetColumnWidth(9, 6);
	xls.SetColumnWidth(10, 6);
	xls.SetColumnWidth(11, 6);
	xls.SetColumnWidth(12, 6);
	xls.SetColumnWidth(13, 6);
	xls.SetColumnWidth(14, 6);
	xls.SetColumnWidth(15, 6);
	xls.SetColumnWidth(16, 6);
	xls.SetColumnWidth(17, 6);
	xls.SetColumnWidth(18, 6);
	xls.SetColumnWidth(19, 6);
	xls.SetColumnWidth(20, 6);
	xls.SetColumnWidth(21, 6);
	xls.SetColumnWidth(22, 6);
	xls.SetColumnWidth(23, 6);
	xls.SetColumnWidth(24, 6);
	xls.SetColumnWidth(25, 6);

	xls.SetCellMergedRows(0, 4, 2);
	xls.SetCellMergedColumns(1, 4, 4);
	xls.SetCellMergedColumns(5, 4, 5);
	xls.SetCellMergedRows(10, 4, 2);
	xls.SetCellMergedColumns(11, 4, 4);
	xls.SetCellMergedRows(15, 4, 2);
	xls.SetCellMergedColumns(16, 4, 3);
	xls.SetCellMergedRows(19, 4, 2);
	xls.SetCellMergedRows(20, 4, 2);
	xls.SetCellMergedRows(21, 4, 2);
	xls.SetCellMergedRows(22, 4, 2);
	xls.SetCellMergedRows(23, 4, 2);
	xls.SetCellMergedRows(24, 4, 2);
	xls.SetCellMergedRows(25, 4, 2);
	xls.SetCellText(0, 4, _T("Ng\xE0y"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(1, 4, _T("T\x103ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(1, 5, _T("S\x1ED1 \x63\x169"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(2, 5, _T("V\xE0o"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(3, 5, _T("\x43K \x111\x1EBFn"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(4, 5, _T("\x43\x1ED9ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(5, 4, _T("Gi\x1EA3m"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(5, 5, _T("V\x1EC1 \x111\x1A1n v\x1ECB"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(6, 5, _T("\x43K \x111i"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(7, 5, _T("\x43V"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(8, 5, _T("T\x1EED vong"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(9, 5, _T("\x43\x1ED9ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(10, 4, _T("\x43\xF2n l\x1EA1i"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(11, 4, _T("Qu\xE2n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(11, 5, _T("T\x1B0\x1EDBng \x63\x169"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(12, 5, _T("\x33//-\x34//"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(13, 5, _T("\x32//-\x43.s\x129"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(14, 5, _T("\x42HQ t\x1EA1i ng\x169"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(15, 4, _T("\x42HQ\x110"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(16, 4, _T("H\x1B0u"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(16, 5, _T("T\x1B0\x1EDBng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(17, 5, _T("\x33//-\x34//"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(18, 5, _T("\x32//-\x43.s\x129"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(19, 4, _T("\x42\x1EA1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(20, 4, _T("T\x45<\x36T"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(21, 4, _T("\x43.s\xE1\x63h"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(22, 4, _T("\x44\x1ECB\x63h v\x1EE5"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(23, 4, _T("\x42HYT kh\xE1\x63"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(24, 4, _T("\x42HTN"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	xls.SetCellText(25, 4, _T("\x43\x1ED9ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, 4098);
	nRow = 6;
	for (int i = 0; i < 14; i++)
		nDataCol[i] = 0;
	xls.SetCellText(0, nRow, m_szFromDate.Left(2), 4098);
	rs.GetValue(_T("prev_qty"), nPrev_Qty);
	xls.SetCellText(1, nRow, int2str(nPrev_Qty), FMT_NUMBER1);
	szSQL = GetQuery_VarianceNumber();
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{	
		rs.GetValue(_T("admit_date"), tmpStr);
		while (dte.Compare(tmpStr) < 0)
		{
			xls.SetCellText(0, nRow, int2str(dte.GetDay()), 4098);
			xls.SetCellText(1, nRow, int2str(nPrev_Qty), FMT_NUMBER1);
			xls.SetCellText(4, nRow, _T("0"), FMT_TEXT | FMT_RIGHT);
			xls.SetCellText(9, nRow, _T("0"), FMT_TEXT | FMT_RIGHT);
			xls.SetCellText(10, nRow, int2str(nPrev_Qty), FMT_NUMBER1, true);
			nRow++;
			dte++;//Tang ngay den khi = ngay ra vao 
		}
		rs.GetValue(_T("in_qty"), nDataCol[0]);
		rs.GetValue(_T("movein_qty"), nDataCol[1]);
		nDataCol[2] = nDataCol[0] + nDataCol[1];
		rs.GetValue(_T("return_unit"), nDataCol[3]);
		rs.GetValue(_T("moveout_qty"), nDataCol[4]);
		rs.GetValue(_T("transfer_qty"), nDataCol[5]);
		rs.GetValue(_T("dead"), nDataCol[6]);
		nDataCol[7] = nDataCol[3] + nDataCol[4] + nDataCol[5] + nDataCol[6];
		nRemain = nPrev_Qty + nDataCol[2] - nDataCol[7];
		xls.SetCellText(0, nRow, int2str(dte.GetDay()), 4098);
		xls.SetCellText(1, nRow, int2str(nPrev_Qty), FMT_NUMBER1);
		xls.SetCellText(2, nRow, int2str(nDataCol[0]), FMT_NUMBER1);
		xls.SetCellText(3, nRow, int2str(nDataCol[1]), FMT_NUMBER1);
		xls.SetCellText(4, nRow, int2str(nDataCol[2]), FMT_NUMBER1);
		xls.SetCellText(5, nRow, int2str(nDataCol[3]), FMT_NUMBER1);
		xls.SetCellText(6, nRow, int2str(nDataCol[4]), FMT_NUMBER1);
		xls.SetCellText(7, nRow, int2str(nDataCol[5]), FMT_NUMBER1);
		xls.SetCellText(8, nRow, int2str(nDataCol[6]), FMT_NUMBER1);
		xls.SetCellText(9, nRow, int2str(nDataCol[7]), FMT_NUMBER1);
		xls.SetCellText(10, nRow, int2str(nRemain), FMT_NUMBER1, true);
		nPrev_Qty = nRemain;
		nRow++;
		dte++;
		rs.MoveNext();
	}
	//Reinit variable
	for (int i = 0; i < 14; i++)
		nDataCol[i] = 0;
	dte.ParseDate(m_szFromDate);
	dte++;
	nRow = 6;
	szSQL = GetQuery_PrevQtybyObj();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;
	for (int i = 0; i < 14; i++)
	{
		tmpStr.Format(_T("c%d"), i+1);
		rs.GetValue(tmpStr, nDataCol[i]);
		xls.SetCellText(i+11, nRow, int2str(nDataCol[i]), FMT_NUMBER1);
	}
	rs.GetValue(_T("total"), nRemain2);
	xls.SetCellText(25, nRow, rs.GetValue(_T("total")), FMT_NUMBER1, true);
	nRow++;
	szSQL = GetQuery_VariancebyObj();
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("admit_date"), tmpStr);
		while (dte.Compare(tmpStr) < 0)
		{
			for (int i = 0; i < 14; i++)
				xls.SetCellText(i+11, nRow, int2str(nDataCol[i]), FMT_NUMBER1);
			xls.SetCellText(25, nRow, int2str(nRemain2), FMT_NUMBER1, true);
			nRow++;
			dte++;//Tang ngay den khi = ngay ra vao 
		}
		for (int i = 0; i < 14; i++)
		{
			tmpStr.Format(_T("c%d"), i+1);
			rs.GetValue(tmpStr, nTmp);
			nDataCol[i] += nTmp;
			xls.SetCellText(i+11, nRow, int2str(nDataCol[i]), FMT_NUMBER1);
		}
		rs.GetValue(_T("total"), nTmp);
		nRemain2 += nTmp;
		xls.SetCellText(25, nRow, int2str(nRemain2), FMT_NUMBER1, true);
		nRow++;
		dte++;
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\Theo doi bn ra vao vien.xls"));	
} 

CString CHMSInOutPatientTrack::GetQuery_PrevQty(){
	CString szSQL;
	//Lay Bn dang dieu tri truoc nut trai thoi diem chay bc
	//Lay Bn vao truoc nut trai, ket thuc sau nut trai
	szSQL.Format(_T(" SELECT SUM(prev_qty) prev_qty") \
	_T(" FROM  (SELECT Count(DISTINCT htr_docno) prev_qty ") \
	_T("        FROM   hms_treatment_record ") \
	_T("        WHERE  htr_status = 'I' AND htr_admitdate < Cast_string2timestamp('%s') AND htr_deptid = '%s'") \
	_T("        UNION ALL ") \
	_T("        SELECT Count(distinct htr_docno) ") \
	_T("        FROM   hms_treatment_record ") \
	_T("        WHERE  htr_status = 'T' AND htr_admitdate < Cast_string2timestamp('%s') ") \
	_T("		AND htr_dischargedate >= cast_string2timestamp('%s') AND htr_deptid = '%s')")
	, m_szFromDate, m_szDeptKey, m_szFromDate, m_szFromDate, m_szDeptKey);

	return szSQL;
}

CString CHMSInOutPatientTrack::GetQuery_PrevQtybyObj(){
	CString szSQL, szFromDate;
	szFromDate.Format(_T("%s 23:59:00"), m_szFromDate.Left(10));
	//Lay Bn dang dieu tri truoc nut phai thoi diem chay bc
	//Lay Bn vao truoc nut phai, ket thuc sau nut phai
	szSQL.Format(_T(" SELECT SUM(c1)  c1, ") \
	_T("        SUM(c2)  c2, ") \
	_T("        SUM(c3)  c3, ") \
	_T("        SUM(c4)  c4, ") \
	_T("        SUM(c5)  c5, ") \
	_T("        SUM(c6)  c6, ") \
	_T("        SUM(c7)  c7, ") \
	_T("        SUM(c8)  c8, ") \
	_T("        SUM(c9)  c9, ") \
	_T("        SUM(c10) c10, ") \
	_T("        SUM(c11) c11, ") \
	_T("        SUM(c12) c12, ") \
	_T("        SUM(c13) c13, ") \
	_T("        SUM(c14) c14, ") \
	_T("        SUM(total) total") \
	_T(" FROM   (    SELECT CASE WHEN hd_object = 1 AND hd_rank > 15 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c1, ") \
	_T("                CASE WHEN hd_object = 1 AND hd_rank IN ( 13, 14 ) AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c2, ") \
	_T("                CASE WHEN hd_object = 1 AND hd_rank < 13 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c3, ") \
	_T("                CASE WHEN hd_object = 2 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c4, ") \
	_T("                CASE WHEN hd_object IN ( 10, 11 ) AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c5, ") \
	_T("                CASE WHEN hd_rank > 15 AND hp_occupation = 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c6, ") \
	_T("                CASE WHEN hd_rank IN ( 13, 14 ) AND hp_occupation = 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c7, ") \
	_T("                CASE WHEN hd_rank < 13 AND hp_occupation = 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c8, ") \
	_T("                CASE WHEN hd_object = 8 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c9, ") \
	_T("                CASE WHEN hd_object = 6 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c10, ") \
	_T("                CASE WHEN hd_object = 3 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c11, ") \
	_T("                CASE WHEN hd_object = 7 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c12, ") \
	_T("                CASE WHEN hd_object = 4 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c13, ") \
	_T("                CASE WHEN hd_object = 5 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c14, ") \
	_T("                1   total ") \
	_T("         FROM   (SELECT    distinct htr_docno, ") \
	_T("                                    hd_object, ") \
	_T("                                    hp_occupation, ") \
	_T("                                    hd_rank ") \
	_T("                 FROM      hms_treatment_record ") \
	_T("                 LEFT JOIN hms_doc ON ( htr_docno = hd_docno ) ") \
	_T("                 LEFT JOIN hms_patient ON ( hd_patientno = hp_patientno ) ") \
	_T("                 WHERE     htr_status = 'I' AND htr_admitdate < Cast_string2timestamp('%s') ") \
	_T("				 AND htr_deptid = '%s'") \
	_T("        UNION ALL ") \
	_T("        SELECT distinct htr_docno, hd_object, hp_occupation, hd_rank ") \
	_T("        FROM   hms_treatment_record ") \
	_T("        LEFT JOIN hms_doc ON ( htr_docno = hd_docno ) ") \
	_T("        LEFT JOIN hms_patient ON ( hd_patientno = hp_patientno ) ") \
	_T("        WHERE  htr_status = 'T' AND htr_admitdate < Cast_string2timestamp('%s') ") \
	_T("		AND htr_dischargedate >= cast_string2timestamp('%s')") \
	_T("		AND htr_deptid = '%s')) "), szFromDate, m_szDeptKey,
	szFromDate, szFromDate, m_szDeptKey);

	return szSQL;
}

CString CHMSInOutPatientTrack::GetQuery_VarianceNumber(){
	//Lay BN ra vao khoa
	//Lay BN chuyen kham giua cac khoa
	CString szSQL, szWhere, szWhere1, szWhere2, szWhere3;
	szWhere.Format(_T(" AND hcr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
				   _T("	AND hcr_admitdept = '%s' "), m_szFromDate, m_szToDate, m_szDeptKey);
	szWhere1.Format(_T(" AND hcr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
					_T(" AND hcr_dischargedept = '%s'"), m_szFromDate, m_szToDate, m_szDeptKey);
	szWhere2.Format(_T(" AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
					_T(" AND htr_deptid = '%s'"), m_szFromDate, m_szToDate, m_szDeptKey);
	szWhere3.Format(_T(" AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
					_T(" AND htr_deptid = '%s'"), m_szFromDate, m_szToDate, m_szDeptKey);
	szSQL.Format(_T(" SELECT admit_date, ") \
	_T("        SUM(in_qty)       in_qty, ") \
	_T("        SUM(movein_qty)   movein_qty, ") \
	_T("        SUM(return_unit)  return_unit, ") \
	_T("        SUM(moveout_qty)  moveout_qty, ") \
	_T("        SUM(transfer_qty) transfer_qty, ") \
	_T("        SUM(dead)         dead ") \
	_T(" FROM   (SELECT Trunc(hcr_admitdate) admit_date, ") \
	_T("                1                    in_qty, ") \
	_T("                0                    movein_qty, ") \
	_T("                0                    return_unit, ") \
	_T("                0                    moveout_qty, ") \
	_T("                0                    transfer_qty, ") \
	_T("                0                    dead ") \
	_T("         FROM   hms_clinical_record ") \
	_T("         WHERE  hcr_status <> 'A' %s ") \
	_T("         UNION ALL ") \
	_T("         SELECT Trunc(hcr_dischargedate), ") \
	_T("                0, ") \
	_T("                0, ") \
	_T("                CASE WHEN hcr_suggestion = 'T' AND hcr_result NOT IN ( 5, 6 ) THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END, ") \
	_T("                0, ") \
	_T("                CASE WHEN hcr_suggestion = 'F' THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END, ") \
	_T("                CASE WHEN hcr_suggestion = 'T' AND hcr_result IN ( 5, 6 ) THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END ") \
	_T("         FROM   hms_clinical_record ") \
	_T("         WHERE  hcr_status = 'T' %s ") \
	_T("         UNION ALL ") \
	_T("         SELECT admit_date, ") \
	_T("                0, ") \
	_T("                1, ") \
	_T("                0, ") \
	_T("                0, ") \
	_T("                0, ") \0
	_T("                0 ") \
	_T("         FROM   (SELECT Trunc(htr_admitdate) admit_date, ") \
	_T("                        htr_docno ") \
	_T("                 FROM   hms_treatment_record ") \
	_T("                 WHERE  htr_status <> 'A' %s AND htr_idx > 1) ") \
	_T("         UNION ALL ") \
	_T("         SELECT discharge_date, ") \
	_T("                0, ") \
	_T("                0, ") \
	_T("                0, ") \
	_T("                1, ") \
	_T("                0, ") \
	_T("                0 ") \
	_T("         FROM  (SELECT Trunc(htr_dischargedate) discharge_date, ") \
	_T("                       htr_docno ") \
	_T("                FROM   hms_treatment_record ") \
	_T("				LEFT JOIN hms_clinical_record ON (hcr_docno = htr_docno)") \
	_T("                WHERE  htr_status = 'T' %s AND htr_idx < hcr_refidx)) ") \
	_T(" GROUP BY admit_date") \
	_T(" ORDER BY admit_date"), szWhere, szWhere1, szWhere2, szWhere3);

	return szSQL;
}

CString CHMSInOutPatientTrack::GetQuery_VariancebyObj(){
	//Lay BN ra vao khoa theo doi tuong
	//Lay BN chuyen kham giua cac khoa theo doi tuong
	CString szSQL, szWhere, szWhere1, szWhere2, szWhere3, szFromDate;
	CDate dte;
	dte.ParseDate(m_szFromDate);
	dte++;
	szFromDate = dte.GetDate();
	szWhere.Format(_T(" AND hcr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
				   _T(" AND hcr_admitdept = '%s'"), szFromDate, m_szToDate, m_szDeptKey);
	szWhere1.Format(_T(" AND hcr_dischargedater BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
					_T(" AND hcr_dischargedept = '%s'"), szFromDate, m_szToDate, m_szDeptKey);
	szWhere2.Format(_T(" AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
					_T(" AND htr_deptid = '%s'"), szFromDate, m_szToDate, m_szDeptKey);
	szWhere3.Format(_T(" AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
					_T(" AND htr_deptid = '%s'"), szFromDate, m_szToDate, m_szDeptKey);
	szSQL.Format(_T(" SELECT admit_date, ") \
	_T("        SUM(c1)    c1, ") \
	_T("        SUM(c2)    c2, ") \
	_T("        SUM(c3)    c3, ") \
	_T("        SUM(c4)    c4, ") \
	_T("        SUM(c5)    c5, ") \
	_T("        SUM(c6)    c6, ") \
	_T("        SUM(c7)    c7, ") \
	_T("        SUM(c8)    c8, ") \
	_T("        SUM(c9)    c9, ") \
	_T("        SUM(c10)   c10, ") \
	_T("        SUM(c11)   c11, ") \
	_T("        SUM(c12)   c12, ") \
	_T("        SUM(c13)   c13, ") \
	_T("        SUM(c14)   c14, ") \
	_T("        SUM(total) total ") \
	_T(" FROM   (SELECT    Trunc(hcr_admitdate) admit_date, ") \
	_T("                   CASE WHEN hd_object = 1 AND hd_rank > 15 AND hp_occupation <> 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c1, ") \
	_T("                   CASE WHEN hd_object = 1 AND hd_rank IN ( 13, 14 ) AND hp_occupation <> 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c2, ") \
	_T("                   CASE WHEN hd_object = 1 AND hd_rank < 13 AND hp_occupation <> 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c3, ") \
	_T("                   CASE WHEN hd_object = 2 AND hp_occupation <> 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c4, ") \
	_T("                   CASE WHEN hd_object IN ( 10, 11 ) AND hp_occupation <> 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c5, ") \
	_T("                   CASE WHEN hd_rank > 15 AND hp_occupation = 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c6, ") \
	_T("                   CASE WHEN hd_rank IN ( 13, 14 ) AND hp_occupation = 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c7, ") \
	_T("                   CASE WHEN hd_rank < 13 AND hp_occupation = 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c8, ") \
	_T("                   CASE WHEN hd_object = 8 AND hp_occupation <> 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c9, ") \
	_T("                   CASE WHEN hd_object = 6 AND hp_occupation <> 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c10, ") \
	_T("                   CASE WHEN hd_object = 3 AND hp_occupation <> 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c11, ") \
	_T("                   CASE WHEN hd_object = 7 AND hp_occupation <> 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c12, ") \
	_T("                   CASE WHEN hd_object = 4 AND hp_occupation <> 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c13, ") \
	_T("                   CASE WHEN hd_object = 5 AND hp_occupation <> 24 THEN 1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END                  c14, ") \
	_T("                   1                    total ") \
	_T("         FROM      hms_clinical_record ") \
	_T("         LEFT JOIN hms_doc ON ( hd_docno = hcr_docno ) ") \
	_T("         LEFT JOIN hms_patient ON ( hp_patientno = hd_patientno ) ") \
	_T("         WHERE     hcr_status <> 'A' %s ") \
	_T("         UNION ALL ") \
	_T("         SELECT    Trunc(hcr_dischargedate), ") \
	_T("                   CASE WHEN hd_object = 1 AND hd_rank > 15 AND hp_occupation <> 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c1, ") \
	_T("                   CASE WHEN hd_object = 1 AND hd_rank IN ( 13, 14 ) AND hp_occupation <> 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c2, ") \
	_T("                   CASE WHEN hd_object = 1 AND hd_rank < 13 AND hp_occupation <> 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c3, ") \
	_T("                   CASE WHEN hd_object = 2 AND hp_occupation <> 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c4, ") \
	_T("                   CASE WHEN hd_object IN ( 10, 11 ) AND hp_occupation <> 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c5, ") \
	_T("                   CASE WHEN hd_rank > 15 AND hp_occupation = 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c6, ") \
	_T("                   CASE WHEN hd_rank IN ( 13, 14 ) AND hp_occupation = 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c7, ") \
	_T("                   CASE WHEN hd_rank < 13 AND hp_occupation = 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c8, ") \
	_T("                   CASE WHEN hd_object = 8 AND hp_occupation <> 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c9, ") \
	_T("                   CASE WHEN hd_object = 6 AND hp_occupation <> 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c10, ") \
	_T("                   CASE WHEN hd_object = 3 AND hp_occupation <> 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c11, ") \
	_T("                   CASE WHEN hd_object = 7 AND hp_occupation <> 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c12, ") \
	_T("                   CASE WHEN hd_object = 4 AND hp_occupation <> 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c13, ") \
	_T("                   CASE WHEN hd_object = 5 AND hp_occupation <> 24 THEN -1 ") \
	_T("                   ELSE 0 ") \
	_T("                   END c14, ") \
	_T("                   -1  total ") \
	_T("         FROM      hms_clinical_record ") \
	_T("         LEFT JOIN hms_doc ON ( hd_docno = hcr_docno ) ") \
	_T("         LEFT JOIN hms_patient ON ( hp_patientno = hd_patientno ) ") \
	_T("         WHERE     hcr_status = 'T' %s ") \
	_T("         UNION ALL ") \
	_T("         SELECT admit_date, ") \
	_T("                CASE WHEN hd_object = 1 AND hd_rank > 15 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c1, ") \
	_T("                CASE WHEN hd_object = 1 AND hd_rank IN ( 13, 14 ) AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c2, ") \
	_T("                CASE WHEN hd_object = 1 AND hd_rank < 13 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c3, ") \
	_T("                CASE WHEN hd_object = 2 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c4, ") \
	_T("                CASE WHEN hd_object IN ( 10, 11 ) AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c5, ") \
	_T("                CASE WHEN hd_rank > 15 AND hp_occupation = 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c6, ") \
	_T("                CASE WHEN hd_rank IN ( 13, 14 ) AND hp_occupation = 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c7, ") \
	_T("                CASE WHEN hd_rank < 13 AND hp_occupation = 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c8, ") \
	_T("                CASE WHEN hd_object = 8 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c9, ") \
	_T("                CASE WHEN hd_object = 6 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c10, ") \
	_T("                CASE WHEN hd_object = 3 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c11, ") \
	_T("                CASE WHEN hd_object = 7 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c12, ") \
	_T("                CASE WHEN hd_object = 4 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c13, ") \
	_T("                CASE WHEN hd_object = 5 AND hp_occupation <> 24 THEN 1 ") \
	_T("                ELSE 0 ") \
	_T("                END c14, ") \
	_T("                1   total ") \
	_T("         FROM   (SELECT    Trunc(htr_admitdate) admit_date, ") \
	_T("                           hd_object, ") \
	_T("                           hd_rank, ") \
	_T("                           hp_occupation ") \
	_T("                 FROM      hms_treatment_record ") \
	_T("                 LEFT JOIN hms_doc ON ( hd_docno = htr_docno ) ") \
	_T("                 LEFT JOIN hms_patient ON ( hp_patientno = hd_patientno ) ") \
	_T("                 WHERE     htr_status <> 'A' %s AND htr_idx > 1) ") \
	_T("         UNION ALL ") \
	_T("         SELECT discharge_date, ") \
	_T("                CASE WHEN hd_object = 1 AND hd_rank > 15 AND hp_occupation <> 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c1, ") \
	_T("                CASE WHEN hd_object = 1 AND hd_rank IN ( 13, 14 ) AND hp_occupation <> 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c2, ") \
	_T("                CASE WHEN hd_object = 1 AND hd_rank < 13 AND hp_occupation <> 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c3, ") \
	_T("                CASE WHEN hd_object = 2 AND hp_occupation <> 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c4, ") \
	_T("                CASE WHEN hd_object IN ( 10, 11 ) AND hp_occupation <> 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c5, ") \
	_T("                CASE WHEN hd_rank > 15 AND hp_occupation = 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c6, ") \
	_T("                CASE WHEN hd_rank IN ( 13, 14 ) AND hp_occupation = 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c7, ") \
	_T("                CASE WHEN hd_rank < 13 AND hp_occupation = 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c8, ") \
	_T("                CASE WHEN hd_object = 8 AND hp_occupation <> 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c9, ") \
	_T("                CASE WHEN hd_object = 6 AND hp_occupation <> 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c10, ") \
	_T("                CASE WHEN hd_object = 3 AND hp_occupation <> 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c11, ") \
	_T("                CASE WHEN hd_object = 7 AND hp_occupation <> 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c12, ") \
	_T("                CASE WHEN hd_object = 4 AND hp_occupation <> 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c13, ") \
	_T("                CASE WHEN hd_object = 5 AND hp_occupation <> 24 THEN -1 ") \
	_T("                ELSE 0 ") \
	_T("                END c14, ") \
	_T("                -1  total ") \
	_T("         FROM  (SELECT     Trunc(htr_dischargedate) discharge_date, ") \
	_T("                           hd_object, ") \
	_T("                           hd_rank, ") \
	_T("                           hp_occupation ") \
	_T("                FROM      hms_treatment_record ") \
	_T("				LEFT JOIN hms_clinical_record ON (hcr_docno = htr_docno)") \
	_T("                LEFT JOIN hms_doc ON ( hd_docno = htr_docno ) ") \
	_T("                LEFT JOIN hms_patient ON ( hp_patientno = hd_patientno ) ") \
	_T("                WHERE     htr_status = 'T' %s AND htr_idx < hcr_refidx)) ") \
	_T(" GROUP BY admit_date ") \
	_T(" ORDER BY admit_date"), szWhere, szWhere1, szWhere2, szWhere3);

	return szSQL;
}