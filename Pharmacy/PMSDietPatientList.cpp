#include "stdafx.h"
#include "PMSDietPatientList.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CPMSDietPatientList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSDietPatientList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CPMSDietPatientList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSDietPatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSDietPatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSDietPatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSDietPatientList *pVw = (CPMSDietPatientList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSDietPatientList *pVw = (CPMSDietPatientList *)pWnd;
	pVw->OnExportSelect();
} 
CPMSDietPatientList::CPMSDietPatientList(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CPMSDietPatientList::~CPMSDietPatientList(){
}
void CPMSDietPatientList::OnCreateComponents(){
	m_wndDietPatientList.Create(this, _T("Diet Patient List"), 5, 5, 430, 90);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 305, 95, 425, 120);
	m_wndExport.Create(this, _T("&Export"), 220, 95, 300, 120);

}
void CPMSDietPatientList::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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

void CPMSDietPatientList::OnSetWindowEvents(){
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
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	SetMode(VM_EDIT);

}
void CPMSDietPatientList::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CPMSDietPatientList::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}

int CPMSDietPatientList::SetMode(int nMode){
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

/*void CPMSDietPatientList::OnYearChange(){
	
} */
/*void CPMSDietPatientList::OnYearSetfocus(){
	
} */
/*void CPMSDietPatientList::OnYearKillfocus(){
	
} */
int CPMSDietPatientList::OnYearCheckValue(){
	return 0;
}
 
void CPMSDietPatientList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
void CPMSDietPatientList::OnReportPeriodSelendok(){
	CString tmpStr;
	CDate dte;
	UpdateData(true);
	int nMonth = ToInt(m_szReportPeriodKey);
	if(nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/1"), m_nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), m_nYear, nMonth, dte.GetMonthLastDay());
	}
	if(nMonth == 13){
		m_szFromDate.Format(_T("%.4d/1/1"), m_nYear);
		tmpStr.Format(_T("%.4d/3/1"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/3/%.2d 23:59" ), m_nYear,  dte.GetMonthLastDay());
	}
	if(nMonth == 14){
		m_szFromDate.Format(_T("%.4d/4/1"), m_nYear);
		tmpStr.Format(_T("%.4d/6/1"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/6/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}

	if(nMonth == 15){
		m_szFromDate.Format(_T("%.4d/7/1"), m_nYear);
		tmpStr.Format(_T("%.4d/9/1"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/9/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 16){
		m_szFromDate.Format(_T("%.4d/10/1"), m_nYear);
		tmpStr.Format(_T("%.4d/10/1"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 17){
		m_szFromDate.Format(_T("%.4d/1/1"), m_nYear);
		tmpStr.Format(_T("%.4d/12/1"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());		
	}	
	UpdateData(false);
}

/*void CPMSDietPatientList::OnReportPeriodSetfocus(){
	
}*/
/*void CPMSDietPatientList::OnReportPeriodKillfocus(){
	
}*/
long CPMSDietPatientList::OnReportPeriodLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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

/*void CPMSDietPatientList::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CPMSDietPatientList::OnFromDateChange(){
	
} */
/*void CPMSDietPatientList::OnFromDateSetfocus(){
	
} */
/*void CPMSDietPatientList::OnFromDateKillfocus(){
	
} */
int CPMSDietPatientList::OnFromDateCheckValue(){
	return 0;
}
 
/*void CPMSDietPatientList::OnToDateChange(){
	
} */
/*void CPMSDietPatientList::OnToDateSetfocus(){
	
} */
/*void CPMSDietPatientList::OnToDateKillfocus(){
	
} */
int CPMSDietPatientList::OnToDateCheckValue(){
	return 0;
}
 
void CPMSDietPatientList::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSDietPatientList::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CString szSQL, tmpStr;
	CStringArray arrColTit;
	CRecord rs(&pMF->m_db);
	BeginWaitCursor();
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);

	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	//xls.SetColumnWidth(0, 7);
	//xls.SetColumnWidth(1, 25);
	//xls.SetColumnWidth(2, 6);
	//xls.SetColumnWidth(3, 10);
	//xls.SetColumnWidth(4, 6);
	//xls.SetColumnWidth(5, 6);
	//xls.SetColumnWidth(6, 10);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 6);
	xls.SetCellText(nCol, nRow + 3, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N \x43\xD3 \x102N"), FMT_TEXT | FMT_CENTER, true, 14);
	xls.SetCellMergedColumns(nCol, nRow + 4, 6);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	//xls.SetCellText(nCol, nRow + 5, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);

	//xls.SetCellText(nCol + 2, nRow + 5, _T("H\x1ECD v\xE0 t\xEAn"), FMT_TEXT | FMT_CENTER, true, 11);
	//xls.SetCellText(nCol + 3, nRow + 5, _T("\x43\x1EA5p \x62\x1EAD\x63"), FMT_TEXT | FMT_CENTER, true, 11);
	//xls.SetCellText(nCol + 4, nRow + 5, _T("\x43h\x1EE9\x63 v\x1EE5"), FMT_TEXT | FMT_CENTER, true, 11);
	//xls.SetCellText(nCol + 5, nRow + 5, _T("\x110\x1A1n v\x1ECB"), FMT_TEXT | FMT_CENTER, true, 11);
	//xls.SetCellText(nCol + 6, nRow + 5, _T("\x110\x1ED1i t\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER, true, 11);
	//xls.SetCellText(nCol + 7, nRow + 5, _T("M\xE3 I\x43\x44\x31\x30"), FMT_TEXT | FMT_CENTER, true, 11);
	//xls.SetCellText(nCol + 8, nRow + 5, _T("\x42\x1EC7nh \x63h\xEDnh"), FMT_TEXT | FMT_CENTER, true, 11);
	arrColTit.Add(_T("Index"));
	arrColTit.Add(_T("Name"));
	arrColTit.Add(_T("Record No"));
	arrColTit.Add(_T("Inward Date"));
	arrColTit.Add(_T("Noon"));
	arrColTit.Add(_T("Afternoon"));
	arrColTit.Add(_T("Food Money"));
	for (int i = 0; i < arrColTit.GetCount(); i++)
	{
		TranslateString(arrColTit.GetAt(i), tmpStr);
		xls.SetCellText(nCol + i, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	}
	nRow = 6;
	int nIndex = 1;
	while(!rs.IsEOF())
	{
		xls.SetCellText(nCol, nRow, int2str(nIndex), FMT_NUMBER1);
		xls.SetCellText(nCol + 1, nRow, rs.GetValue(_T("patname")), FMT_TEXT);
		xls.SetCellText(nCol + 2, nRow, rs.GetValue(_T("recordno")), FMT_TEXT);
		rs.GetValue(_T("inwardte"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), FMT_TEXT);
		xls.SetCellText(nCol + 6, nRow, rs.GetValue(_T("amount")), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\DANH SACH BENH NHAN CO AN.xls"));
}
 
CString CPMSDietPatientList::GetQueryString(){
	CString szSQL, szWhere; 
	szWhere.Format(_T(" AND hfo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szSQL.Format(_T(" SELECT") \
				_T("   Get_patientname(hfo_docno) AS patname,") \
				_T("   hcr_numinward              AS recordno,") \
				_T("   hcr_admitdate              AS inwardte,") \
				_T("   hfo_amount                 AS amount") \
				_T(" FROM   hms_feefood") \
				_T(" LEFT JOIN hms_clinical_record ON (hfo_docno=hcr_docno)") \
				_T(" WHERE hfo_orderstatus <> 'O' %s") \
				_T(" ORDER  BY patname "), szWhere);

	return szSQL;
}