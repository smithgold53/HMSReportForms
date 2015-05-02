#include "stdafx.h"
#include "EMGeneralSoldierExamination.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMGeneralSoldierExamination *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMGeneralSoldierExamination* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMGeneralSoldierExamination *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMGeneralSoldierExamination *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMGeneralSoldierExamination *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMGeneralSoldierExamination *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMGeneralSoldierExamination *pVw = (CEMGeneralSoldierExamination *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMGeneralSoldierExamination *pVw = (CEMGeneralSoldierExamination *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddEMGeneralSoldierExaminationFnc(CWnd *pWnd){
	 return ((CEMGeneralSoldierExamination*)pWnd)->OnAddEMGeneralSoldierExamination();
} 
static int _OnEditEMGeneralSoldierExaminationFnc(CWnd *pWnd){
	 return ((CEMGeneralSoldierExamination*)pWnd)->OnEditEMGeneralSoldierExamination();
} 
static int _OnDeleteEMGeneralSoldierExaminationFnc(CWnd *pWnd){
	 return ((CEMGeneralSoldierExamination*)pWnd)->OnDeleteEMGeneralSoldierExamination();
} 
static int _OnSaveEMGeneralSoldierExaminationFnc(CWnd *pWnd){
	 return ((CEMGeneralSoldierExamination*)pWnd)->OnSaveEMGeneralSoldierExamination();
} 
static int _OnCancelEMGeneralSoldierExaminationFnc(CWnd *pWnd){
	 return ((CEMGeneralSoldierExamination*)pWnd)->OnCancelEMGeneralSoldierExamination();
} 
CEMGeneralSoldierExamination::CEMGeneralSoldierExamination(CWnd* pParent)
{
	m_nDlgWidth = 475;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CEMGeneralSoldierExamination::~CEMGeneralSoldierExamination(){
}
void CEMGeneralSoldierExamination::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 465, 90);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 115, 55);
	m_wndYear.Create(this,120, 30, 240, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 245, 30, 335, 55);
	m_wndReportPeriod.Create(this,340, 30, 460, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 115, 85);
	m_wndFromDate.Create(this,120, 60, 240, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 245, 60, 335, 85);
	m_wndToDate.Create(this,340, 60, 460, 85); 
	//m_wndPrint.Create(this, _T("&Print"), 260, 95, 360, 120);
	m_wndExport.Create(this, _T("&Export"), 365, 95, 465, 120);
}
void CEMGeneralSoldierExamination::OnInitializeComponents(){
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

}
void CEMGeneralSoldierExamination::OnSetWindowEvents()
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

	CString szSQL;
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);

}
void CEMGeneralSoldierExamination::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CEMGeneralSoldierExamination::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMGeneralSoldierExamination::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMGeneralSoldierExamination::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMGeneralSoldierExamination::SetMode(int nMode){
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
/*void CEMGeneralSoldierExamination::OnYearChange(){
	
} */
/*void CEMGeneralSoldierExamination::OnYearSetfocus(){
	
} */
/*void CEMGeneralSoldierExamination::OnYearKillfocus(){
	
} */
int CEMGeneralSoldierExamination::OnYearCheckValue()
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
void CEMGeneralSoldierExamination::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMGeneralSoldierExamination::OnReportPeriodSelendok()
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
/*void CEMGeneralSoldierExamination::OnReportPeriodSetfocus(){
	
}*/
/*void CEMGeneralSoldierExamination::OnReportPeriodKillfocus(){
	
}*/
long CEMGeneralSoldierExamination::OnReportPeriodLoadData()
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
/*void CEMGeneralSoldierExamination::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMGeneralSoldierExamination::OnFromDateChange(){
	
} */
/*void CEMGeneralSoldierExamination::OnFromDateSetfocus(){
	
} */
/*void CEMGeneralSoldierExamination::OnFromDateKillfocus(){
	
} */
int CEMGeneralSoldierExamination::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMGeneralSoldierExamination::OnToDateChange(){
	
} */
/*void CEMGeneralSoldierExamination::OnToDateSetfocus(){
	
} */
/*void CEMGeneralSoldierExamination::OnToDateKillfocus(){
	
} */
int CEMGeneralSoldierExamination::OnToDateCheckValue(){
	return 0;
} 
void CEMGeneralSoldierExamination::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMGeneralSoldierExamination::OnExportSelect()
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

	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 23);

	xls.SetColumnWidth(2, 7);
	xls.SetColumnWidth(3, 10);

	xls.SetColumnWidth(4, 7);
	xls.SetColumnWidth(5, 10);

	xls.SetColumnWidth(6, 7);
	xls.SetColumnWidth(7, 10);

	xls.SetColumnWidth(8, 7);
	xls.SetColumnWidth(9, 10);

	xls.SetColumnWidth(10, 7);
	xls.SetColumnWidth(11, 10);

	xls.SetColumnWidth(12, 7);
	xls.SetColumnWidth(13, 10);

	int nRow = 1;
	int nCol = 0;

	xls.SetRowHeight(6, 20);
	xls.SetRowHeight(7, 36);
	xls.SetRowHeight(8, 36);

	xls.SetCellMergedColumns(0, 1, 4);
	xls.SetCellMergedColumns(0, 2, 4);

	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol, nRow + 3, 16);
	xls.SetCellMergedColumns(nCol, nRow + 4, 16);

	xls.SetCellText(nCol, nRow + 3, _T("T\xECnh h\xECnh kh\xE1m, \x63h\x1EEF\x61 \x62\x1EC7nh \x63\x1EE7\x61 qu\xE2n nh\xE2n th\x61m gi\x61 \x110\x1EC1 \xE1n th\xED \x111i\x1EC3m \x62\x1EA3o hi\x1EC3m y t\x1EBF qu\xE2n nh\xE2n"),
					FMT_TEXT | FMT_CENTER, true, 16, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);

	xls.SetCellMergedRows(nCol, nRow + 5, 3);
	xls.SetCellText(nCol, nRow + 5, _T("TT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedRows(nCol + 1, nRow + 5, 3);
	xls.SetCellText(nCol + 1, nRow + 5, _T("\x110\x1A1n v\x1ECB th\x61m gi\x61 \x110\x1EC1 \xE1n"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedColumns(nCol + 2, nRow + 5, 6);
	xls.SetCellText(nCol + 2, nRow + 5, _T("K\x43\x42 ngo\x1EA1i tr\xFA"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedColumns(nCol + 2, nRow + 6, 2);
	xls.SetCellText(nCol + 2, nRow + 6, _T("\x43hi ph\xED t\x1EEB NSQP"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	xls.SetCellText(nCol + 2,  nRow + 7, _T("S\x1ED1 l\x1B0\x1EE3t"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);
	xls.SetCellText(nCol + 3,  nRow + 7, _T("S\x1ED1 ti\x1EC1n (\x111\x1ED3ng)"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	xls.SetCellMergedColumns(nCol + 4, nRow + 6, 2);
	xls.SetCellText(nCol + 4, nRow + 6, _T("\x43hi ph\xED t\x1EEB qu\x1EF9 \x42HYT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	xls.SetCellText(nCol + 4,  nRow + 7, _T("S\x1ED1 l\x1B0\x1EE3t"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);
	xls.SetCellText(nCol + 5,  nRow + 7, _T("S\x1ED1 ti\x1EC1n (\x111\x1ED3ng)"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	xls.SetCellMergedColumns(nCol + 6, nRow + 6, 2);
	xls.SetCellText(nCol + 6, nRow + 6, _T("\x43\x1ED9ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	xls.SetCellText(nCol + 6,  nRow + 7, _T("S\x1ED1 l\x1B0\x1EE3t"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);
	xls.SetCellText(nCol + 7,  nRow + 7, _T("S\x1ED1 ti\x1EC1n (\x111\x1ED3ng)"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	xls.SetCellMergedColumns(nCol + 8, nRow + 5, 6);
	xls.SetCellText(nCol + 8, nRow + 5, _T("K\x43\x42 n\x1ED9i tr\xFA"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellMergedColumns(nCol + 8, nRow + 6, 2);
	xls.SetCellText(nCol + 8, nRow + 6, _T("\x43hi ph\xED t\x1EEB NSQP"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	xls.SetCellText(nCol + 8,  nRow + 7, _T("S\x1ED1 l\x1B0\x1EE3t"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);
	xls.SetCellText(nCol + 9,  nRow + 7, _T("S\x1ED1 ti\x1EC1n (\x111\x1ED3ng)"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	xls.SetCellMergedColumns(nCol + 10, nRow + 6, 2);
	xls.SetCellText(nCol + 10, nRow + 6, _T("\x43hi ph\xED t\x1EEB qu\x1EF9 \x42HYT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	xls.SetCellText(nCol + 10,  nRow + 7, _T("S\x1ED1 l\x1B0\x1EE3t"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);
	xls.SetCellText(nCol + 11,  nRow + 7, _T("S\x1ED1 ti\x1EC1n (\x111\x1ED3ng)"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	xls.SetCellMergedColumns(nCol + 12, nRow + 6, 2);
	xls.SetCellText(nCol + 12, nRow + 6, _T("\x43\x1ED9ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	xls.SetCellText(nCol + 12,  nRow + 7, _T("S\x1ED1 l\x1B0\x1EE3t"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);
	xls.SetCellText(nCol + 13,  nRow + 7, _T("S\x1ED1 ti\x1EC1n (\x111\x1ED3ng)"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	xls.SetCellMergedColumns(nCol + 14, nRow + 5, 2);
	xls.SetCellMergedRows(nCol + 14, nRow + 5, 2);

	xls.SetCellText(nCol + 14, nRow + 5, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);
	xls.SetCellText(nCol + 14, nRow + 7, _T("S\x1ED1 l\x1B0\x1EE3t"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);
	xls.SetCellText(nCol + 15, nRow + 7, _T("S\x1ED1 ti\x1EC1n (\x111\x1ED3ng)"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, false, 11);

	int nIndex = 1;
	nRow = 8;

	while (!rs.IsEOF())
	{
		nRow++;

		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("workplace"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("ngt"), tmpStr);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("nt"), tmpStr);
		xls.SetCellText(nCol + 12, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("tong"), tmpStr);
		xls.SetCellText(nCol + 14, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.MoveNext();
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\THKCBQuanNhan.xls"));
} 
int CEMGeneralSoldierExamination::OnAddEMGeneralSoldierExamination(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CEMGeneralSoldierExamination::OnEditEMGeneralSoldierExamination(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMGeneralSoldierExamination::OnDeleteEMGeneralSoldierExamination(){
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
 		OnCancelEMGeneralSoldierExamination();
 	}
	return 0;
}
int CEMGeneralSoldierExamination::OnSaveEMGeneralSoldierExamination(){
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
 		//OnEMGeneralSoldierExaminationListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CEMGeneralSoldierExamination::OnCancelEMGeneralSoldierExamination(){
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
int CEMGeneralSoldierExamination::OnEMGeneralSoldierExaminationListLoadData(){
	return 0;
}

CString CEMGeneralSoldierExamination::GetQueryString()
{
	CString szSQL;

	szSQL.Format(_T(" select workplace,") \
				_T("        sum(ngt) as ngt,") \
				_T("        sum(nt) as nt,") \
				_T("        sum(ngt+nt) as tong") \
				_T(" from") \
				_T(" (") \
				_T("  select case when hd_suggestion not in('D','C') then 1 else 0 end as ngt,") \
				_T("         case when hd_suggestion in('D','C') and length(hcr_recordno) > 0 then 1 else 0 end as nt,") \
				_T("         trim(hp_workplace) as workplace") \
				_T("  from hms_patient") \
				_T("  left join hms_doc on(hd_patientno=hp_patientno)") \
				_T("  left join hms_clinical_record on(hcr_docno=hd_docno)") \
				_T("  where length(hp_workplace) > 0") \
				_T("        and substr(hd_cardno,1,3)='QN9'") \
				_T("        and hd_status='T'") \
				_T("        and hd_enddate between cast_string2timestamp('%s')") \
				_T("        and cast_string2timestamp('%s')") \
				_T(" ) Tbl") \
				_T(" group by workplace") \
				_T(" order by workplace"),
				m_szFromDate, m_szToDate);

	return szSQL;
}
