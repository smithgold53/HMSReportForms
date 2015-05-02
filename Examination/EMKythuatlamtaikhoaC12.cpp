#include "stdafx.h"
#include "EMKythuatlamtaikhoaC12.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMKythuatlamtaikhoaC12  *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMKythuatlamtaikhoaC12 * )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMKythuatlamtaikhoaC12  *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMKythuatlamtaikhoaC12  *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMKythuatlamtaikhoaC12  *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMKythuatlamtaikhoaC12  *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMKythuatlamtaikhoaC12  *pVw = (CEMKythuatlamtaikhoaC12  *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMKythuatlamtaikhoaC12  *pVw = (CEMKythuatlamtaikhoaC12  *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMKythuatlamtaikhoaC12  *pVw = (CEMKythuatlamtaikhoaC12  *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMKythuatlamtaikhoaC12 *)pWnd)->OnAddEMStatisticsGeneralPatientReport();
} 
static int _OnEditEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMKythuatlamtaikhoaC12 *)pWnd)->OnEditEMStatisticsGeneralPatientReport();
} 
static int _OnDeleteEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMKythuatlamtaikhoaC12 *)pWnd)->OnDeleteEMStatisticsGeneralPatientReport();
} 
static int _OnSaveEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMKythuatlamtaikhoaC12 *)pWnd)->OnSaveEMStatisticsGeneralPatientReport();
} 
static int _OnCancelEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMKythuatlamtaikhoaC12 *)pWnd)->OnCancelEMStatisticsGeneralPatientReport();
} 
CEMKythuatlamtaikhoaC12 ::CEMKythuatlamtaikhoaC12 (CWnd *pParent)
{
	m_szTitle = _T("Patients General Statistics");
	m_nDlgWidth = 420;
	m_nDlgHeight = 125;
	SetDefaultValues();
}
CEMKythuatlamtaikhoaC12 ::~CEMKythuatlamtaikhoaC12 ()
{
}
void CEMKythuatlamtaikhoaC12 ::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 205, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 215, 30, 295, 55);
	m_wndReportPeriod.Create(this,300, 30, 410, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 415, 90);
	m_wndFromDate.Create(this,95, 60, 205, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 215, 60, 295, 85);
	m_wndToDate.Create(this,300, 60, 410, 85); 
	//m_wndPrint.Create(this, _T("&Print"), 250, 95, 330, 120);
	m_wndExport.Create(this, _T("&Export"), 335, 95, 415, 120);

}
void CEMKythuatlamtaikhoaC12 ::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(false);
	m_wndReportPeriod.SetCheckValue(false);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(false);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(false);
	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);
}
void CEMKythuatlamtaikhoaC12 ::OnSetWindowEvents(){
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
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	SetMode(VM_EDIT);
}
void CEMKythuatlamtaikhoaC12 ::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
}
void CEMKythuatlamtaikhoaC12 ::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CEMKythuatlamtaikhoaC12 ::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMKythuatlamtaikhoaC12 ::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMKythuatlamtaikhoaC12 ::SetMode(int nMode){
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
 			EnableButtons(TRUE,0,1,2,3,-1); 
 			break; 
 		case VM_VIEW: 
 			EnableControls(TRUE); 
 			EnableButtons(FALSE, 3, 4, -1); 
 			break; 
 		case VM_NONE: 
 			EnableControls(TRUE); 
 			EnableButtons(TRUE, 0, 6, -1); 
 			SetDefaultValues(); 
 			break; 
 		}; 
 		UpdateData(FALSE); 
 		return nOldMode; 
}
/*void CEMKythuatlamtaikhoaC12 ::OnYearChange(){
	
} */
/*void CEMKythuatlamtaikhoaC12 ::OnYearSetfocus(){
	
} */
/*void CEMKythuatlamtaikhoaC12 ::OnYearKillfocus(){
	
} */
int CEMKythuatlamtaikhoaC12 ::OnYearCheckValue(){
	return 0;
} 
void CEMKythuatlamtaikhoaC12 ::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMKythuatlamtaikhoaC12 ::OnReportPeriodSelendok(){
	CString tmpStr;
	CDate dte;
	UpdateData(true);
	int nMonth = ToInt(m_szReportPeriodKey);
	if(nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/01"), m_nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), m_nYear, nMonth, dte.GetMonthLastDay());
	}
	if(nMonth == 13){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/03/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59" ), m_nYear,  dte.GetMonthLastDay());
	}
	if(nMonth == 14){
		m_szFromDate.Format(_T("%.4d/04/01"), m_nYear);
		tmpStr.Format(_T("%.4d/06/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 15){
		m_szFromDate.Format(_T("%.4d/07/01"), m_nYear);
		tmpStr.Format(_T("%.4d/09/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 16){
		m_szFromDate.Format(_T("%.4d/10/01"), m_nYear);
		tmpStr.Format(_T("%.4d/10/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 17){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/12/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());		
	}	
	UpdateData(false);	 
	 
}
/*void CEMKythuatlamtaikhoaC12 ::OnReportPeriodSetfocus(){
	
}*/
/*void CEMKythuatlamtaikhoaC12 ::OnReportPeriodKillfocus(){
	
}*/
long CEMKythuatlamtaikhoaC12 ::OnReportPeriodLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
		CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0)
	{
		szWhere.Format(_T(" WHERE hpr_idx=%ld"), ToInt(m_szReportPeriodKey));
	};
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
/*void CEMKythuatlamtaikhoaC12 ::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMKythuatlamtaikhoaC12 ::OnFromDateChange(){
	
} */
/*void CEMKythuatlamtaikhoaC12 ::OnFromDateSetfocus(){
	
} */
/*void CEMKythuatlamtaikhoaC12 ::OnFromDateKillfocus(){
	
} */
int CEMKythuatlamtaikhoaC12 ::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMKythuatlamtaikhoaC12 ::OnToDateChange(){
	
} */
/*void CEMKythuatlamtaikhoaC12 ::OnToDateSetfocus(){
	
} */
/*void CEMKythuatlamtaikhoaC12 ::OnToDateKillfocus(){
	
} */
int CEMKythuatlamtaikhoaC12 ::OnToDateCheckValue(){
	return 0;
} 
void CEMKythuatlamtaikhoaC12 ::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
} 
void CEMKythuatlamtaikhoaC12 ::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, szTemp, tmpStr;
	CExcel xls;
	BeginWaitCursor();
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 15);
	for(int i = 2; i<=20; i++)
	{
		xls.SetColumnWidth(i, 12);
	}
	int nRow = 0, nCol = 0;
	
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellMergedColumns(nCol, nRow + 1, 4);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(nCol, nRow + 2, 14);
	xls.SetCellMergedColumns(nCol, nRow + 3, 14);
	xls.SetCellText(nCol, nRow + 2, _T("K\x1EF8 THU\x1EACT L\xC0M T\x1EA0I KHO\x41"), FMT_TEXT | FMT_CENTER, true, 13);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y: %s \x110\x1EBFn ng\xE0y: %s"), 
		          CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss),
				  CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellMergedRows(nCol, nRow + 4, 2);
	xls.SetCellText(nCol, nRow + 4, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedRows(nCol + 1, nRow + 4, 2);
	xls.SetCellText(nCol + 1, nRow + 4, _T("Ng\xE0y"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedColumns(nCol + 2, nRow + 4, 4);
	xls.SetCellText(nCol + 2, nRow + 4, _T("\x110I\x1EC6N TIM"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 5, _T("Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 3, nRow + 5, _T("H\x1B0u"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 5, _T("\x42HYT Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 5, nRow + 5, _T("T\x1ED5ng s\x1ED1"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedColumns(nCol + 6, nRow + 4, 4);
	xls.SetCellText(nCol + 6, nRow + 4, _T("SI\xCAU \xC2M"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 6, nRow + 5, _T("Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 7, nRow + 5, _T("H\x1B0u"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 8, nRow + 5, _T("\x42HYT Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 9, nRow + 5, _T("T\x1ED5ng s\x1ED1"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedColumns(nCol + 10, nRow + 4, 4);
	xls.SetCellText(nCol + 10, nRow + 4, _T("HUY\x1EBET H\x1ECC\x43"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 10, nRow + 5, _T("Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 11, nRow + 5, _T("H\x1B0u"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 12, nRow + 5, _T("\x42HYT Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 13, nRow + 5, _T("T\x1ED5ng s\x1ED1"), FMT_TEXT | FMT_CENTER, true, 11);

	nRow += 6;
	int nIndex = 1;

	int nTotal[15];
	for(int i = 0; i<=13; i++)
		{
			nTotal[i] = 0;
		}
	while (!rs.IsEOF())
	{
		szTemp.Format(_T("%d"), nIndex);
		xls.SetCellText(nCol, nRow, szTemp, FMT_INTEGER);

		szTemp = rs.GetValue(_T("hpc_orderdate"));
		xls.SetCellText(nCol + 1, nRow, CDate::Convert(szTemp, yyyymmdd, ddmmyyyy), FMT_TEXT);

		szTemp = rs.GetValue(_T("dientimquan"));
		nTotal[2] += ToInt(szTemp);
		xls.SetCellText(nCol + 2, nRow, szTemp, FMT_NUMBER1);
		szTemp = rs.GetValue(_T("dientimhuu"));
		nTotal[3] += ToInt(szTemp);
		xls.SetCellText(nCol + 3, nRow, szTemp, FMT_NUMBER1);
		szTemp = rs.GetValue(_T("dientimbhytquan"));
		nTotal[4] += ToInt(szTemp);
		xls.SetCellText(nCol + 4, nRow, szTemp, FMT_NUMBER1);
		szTemp = rs.GetValue(_T("tongdientim"));
		nTotal[5] += ToInt(szTemp);
		xls.SetCellText(nCol + 5, nRow, szTemp, FMT_NUMBER1);


		szTemp = rs.GetValue(_T("sieuamquan"));
		nTotal[6] += ToInt(szTemp);
		xls.SetCellText(nCol + 6, nRow, szTemp, FMT_NUMBER1);
		szTemp = rs.GetValue(_T("sieuamhuu"));
		nTotal[7] += ToInt(szTemp);
		xls.SetCellText(nCol + 7, nRow, szTemp, FMT_NUMBER1);
		szTemp = rs.GetValue(_T("sieuambhytquan"));
		nTotal[8] += ToInt(szTemp);
		xls.SetCellText(nCol + 8, nRow, szTemp, FMT_NUMBER1);
		szTemp = rs.GetValue(_T("tongsieuam"));
		nTotal[9] += ToInt(szTemp);
		xls.SetCellText(nCol + 9, nRow, szTemp, FMT_NUMBER1);

		szTemp = rs.GetValue(_T("huyethocquan"));
		nTotal[10] += ToInt(szTemp);
		xls.SetCellText(nCol + 10, nRow, szTemp, FMT_NUMBER1);
		szTemp = rs.GetValue(_T("huyethochuu"));
		nTotal[11] += ToInt(szTemp);
		xls.SetCellText(nCol + 11, nRow, szTemp, FMT_NUMBER1);
		szTemp = rs.GetValue(_T("huyethocbhytquan"));
		nTotal[12] += ToInt(szTemp);
		xls.SetCellText(nCol + 12, nRow, szTemp, FMT_NUMBER1);
		szTemp = rs.GetValue(_T("tonghuyethoc"));
		nTotal[13] += ToInt(szTemp);
		xls.SetCellText(nCol + 13, nRow, szTemp, FMT_NUMBER1);

		nIndex++;
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 11);
	 for( int i = 2; i<= 13; i++)
	 {
		 tmpStr.Format(_T("%d"), nTotal[i]);
		 xls.SetCellText(i, nRow, tmpStr, FMT_NUMBER1, true, 11);
	 }
	
	EndWaitCursor();
	xls.Save(_T("Exports\\Ky thuat lam tai khoa.xls"));

} 
void CEMKythuatlamtaikhoaC12 ::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CEMKythuatlamtaikhoaC12 ::OnAddEMStatisticsGeneralPatientReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMKythuatlamtaikhoaC12 ::OnEditEMStatisticsGeneralPatientReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMKythuatlamtaikhoaC12 ::OnDeleteEMStatisticsGeneralPatientReport(){
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
 		OnCancelEMStatisticsGeneralPatientReport(); 
 	}
	return 0;
}
int CEMKythuatlamtaikhoaC12 ::OnSaveEMStatisticsGeneralPatientReport(){
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
 		//OnEMStatisticsGeneralPatientReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMKythuatlamtaikhoaC12 ::OnCancelEMStatisticsGeneralPatientReport(){
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
int CEMKythuatlamtaikhoaC12 ::OnEMStatisticsGeneralPatientReportListLoadData()
{
	return 0;
}
CString CEMKythuatlamtaikhoaC12 ::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Empty();
		szSQL.Format(_T("   select  hpc_orderdate,") \
					_T("              sum(dientimquan) as dientimquan, sum(dientimhuu) as dientimhuu,   sum(dientimbhytquan) as dientimbhytquan,") \
					_T("              sum(dientimquan+dientimhuu+dientimbhytquan) as tongdientim,") \
					_T("              sum(sieuamquan) as sieuamquan, sum(sieuamhuu) as sieuamhuu,   sum(sieuambhytquan) as sieuambhytquan, ") \
					_T("              sum(sieuamquan+sieuamhuu+sieuambhytquan) as tongsieuam,") \
					_T("              sum(huyethocquan) as huyethocquan, sum(huyethochuu) as huyethochuu,  sum(huyethocbhytquan) as huyethocbhytquan, ") \
					_T("              sum(huyethocquan+huyethochuu+huyethocbhytquan) as tonghuyethoc, ") \
					_T("              sum(sinhhoaquan) as sinhhoaquan,  sum(sinhhoahuu) as sinhhoahuu, sum(sinhhoabhytquan) as sinhhoabhytquan,") \
					_T("              sum(sinhhoaquan+sinhhoahuu+sinhhoabhytquan) as tongsinhhoa") \
					_T("              from") \
					_T("              (") \
					_T("                SELECT  Cast_Timestamp2Date(hpc_orderdate) as hpc_orderdate ,") \
					_T("                ") \
					_T("                             CASE WHEN Substr(hpc_groupid, 1, 3)='B12' and hd_object = '1' THEN 1 ELSE 0 END AS sinhhoaquan,") \
					_T("                             CASE WHEN Substr(hpc_groupid, 1, 3)='B11' and hd_object = '1' THEN 1 ELSE 0 END AS huyethocquan,") \
					_T("                             CASE WHEN hpc_groupid='B3200' and hd_object = '1' THEN 1 ELSE 0 END AS sieuamquan, ") \
					_T("                             CASE WHEN hpc_groupid='B3300' and hd_object = '1' THEN 1 ELSE 0 END AS dientimquan, ") \
					_T("                             CASE WHEN Substr(hpc_groupid, 1, 3)='B12' and hd_object = '2' THEN 1 ELSE 0 END AS sinhhoahuu,") \
					_T("                             CASE WHEN Substr(hpc_groupid, 1, 3)='B11' and hd_object = '2' THEN 1 ELSE 0 END AS huyethochuu,") \
					_T("                             CASE WHEN hpc_groupid='B3200' and hd_object = '2' THEN 1 ELSE 0 END AS sieuamhuu, ") \
					_T("                             CASE WHEN hpc_groupid='B3300' and hd_object = '2' THEN 1 ELSE 0 END AS dientimhuu,                            ") \
					_T("                             CASE WHEN Substr(hpc_groupid, 1, 3)='B12' and hd_object = '11' THEN 1 ELSE 0 END AS sinhhoabhytquan,") \
					_T("                             CASE WHEN Substr(hpc_groupid, 1, 3)='B11' and hd_object = '11' THEN 1 ELSE 0 END AS huyethocbhytquan,") \
					_T("                             CASE WHEN hpc_groupid='B3200' and hd_object = '11' THEN 1 ELSE 0 END AS sieuambhytquan, ") \
					_T("                             CASE WHEN hpc_groupid='B3300' and hd_object = '11' THEN 1 ELSE 0 END AS dientimbhytquan") \
					_T("                       FROM   ") \
					_T("                       (") \
					_T("                         SELECT hpc_roomid, hpc_docno, hpc_groupid, hpc_orderdate, hd_object") \
					_T("                           FROM   hms_testorder") \
					_T("                           LEFT JOIN hms_exam ON (he_docno = hpc_docno)") \
					_T("                           LEFT JOIN hms_doc ON (he_docno = hd_docno)") \
					_T("                           WHERE  hpc_status<>'O' AND Cast_Timestamp2Date(hpc_orderdate) BETWEEN Cast_string2date('%s') AND Cast_string2date('%s')") \
					_T("                                and he_deptid  = 'C1.2'") \
					_T("                           GROUP  BY hpc_roomid,hpc_groupid,hpc_docno, hpc_orderdate, hd_object") \
					_T("                          UNION ALL") \
					_T("                          SELECT hpc_roomid, hpc_docno, hpc_groupid, hpc_orderdate, hd_object") \
					_T("                          FROM   hms_pacsorder") \
					_T("                          LEFT JOIN hms_exam ON (he_docno = hpc_docno)") \
					_T("                          LEFT JOIN hms_doc ON (he_docno = hd_docno)") \
					_T("                          WHERE  hpc_status<>'O' AND Cast_Timestamp2Date(hpc_orderdate) BETWEEN Cast_string2date('%s') AND Cast_string2date('%s')") \
					_T("                                and he_deptid  = 'C1.2'") \
					_T("                           GROUP  BY hpc_roomid, hpc_groupid, hpc_docno, hpc_orderdate, hd_object") \
					_T("                         ) tmp") \
					_T("                 ) total") \
					_T("     GROUP  BY total.hpc_orderdate") \
					_T("     ORDER  BY total.hpc_orderdate"), m_szFromDate, m_szToDate, m_szFromDate, m_szToDate);
	return szSQL;
}
