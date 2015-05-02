#include "stdafx.h"
#include "FMInsurancePatientTechDetail.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMInsurancePatientTechDetail *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMInsurancePatientTechDetail* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMInsurancePatientTechDetail *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMInsurancePatientTechDetail *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMInsurancePatientTechDetail *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMInsurancePatientTechDetail *pVw = (CFMInsurancePatientTechDetail *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMInsurancePatientTechDetail *pVw = (CFMInsurancePatientTechDetail *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMInsurancePatientTechDetail*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMInsurancePatientTechDetail*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMInsurancePatientTechDetail*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMInsurancePatientTechDetail*)pWnd)->OnDeptListDeleteItem();
} 
CFMInsurancePatientTechDetail::CFMInsurancePatientTechDetail(CWnd *pWnd){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CFMInsurancePatientTechDetail::~CFMInsurancePatientTechDetail(){
}
void CFMInsurancePatientTechDetail::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 550);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 90, 425, 545);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndPrint.Create(this, _T("&Print"), 265, 555, 345, 580);
	m_wndExport.Create(this, _T("&Export"), 350, 555, 430, 580);
	m_wndDeptList.Create(this,15, 115, 420, 540); 

}
void CFMInsurancePatientTechDetail::OnInitializeComponents(){
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


	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.SetCheckBox(TRUE);

}

void CFMInsurancePatientTechDetail::OnSetWindowEvents(){
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
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);
	OnDeptListLoadData();

}
void CFMInsurancePatientTechDetail::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CFMInsurancePatientTechDetail::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}

int CFMInsurancePatientTechDetail::SetMode(int nMode){
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

/*void CFMInsurancePatientTechDetail::OnYearChange(){
	
} */
/*void CFMInsurancePatientTechDetail::OnYearSetfocus(){
	
} */
/*void CFMInsurancePatientTechDetail::OnYearKillfocus(){
	
} */
int CFMInsurancePatientTechDetail::OnYearCheckValue(){
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
			szTemp.Format(_T("%.2d/2d/4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndFromDate.SetWindowText(szTemp);
		}
		dt.ParseDateTime(m_szToDate);
		date = dt.GetDate();
		if (date.GetYear() != 1752)
		{
			dt.SetDate(m_nYear, date.GetMonth(), date.GetDay());
			m_szToDate = dt.GetDateTime();
			szTemp.Format(_T("%.2d/2d/4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndToDate.SetWindowText(szTemp);
		}
	}
	UpdateData(FALSE);
	return 0;
}
 
void CFMInsurancePatientTechDetail::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
void CFMInsurancePatientTechDetail::OnReportPeriodSelendok(){
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
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

/*void CFMInsurancePatientTechDetail::OnReportPeriodSetfocus(){
	
}*/
/*void CFMInsurancePatientTechDetail::OnReportPeriodKillfocus(){
	
}*/
long CFMInsurancePatientTechDetail::OnReportPeriodLoadData(){
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

/*void CFMInsurancePatientTechDetail::OnReportPeriodAddNew(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMInsurancePatientTechDetail::OnFromDateChange(){
	
} */
/*void CFMInsurancePatientTechDetail::OnFromDateSetfocus(){
	
} */
/*void CFMInsurancePatientTechDetail::OnFromDateKillfocus(){
	
} */
int CFMInsurancePatientTechDetail::OnFromDateCheckValue(){
	return 0;
}
 
/*void CFMInsurancePatientTechDetail::OnToDateChange(){
	
} */
/*void CFMInsurancePatientTechDetail::OnToDateSetfocus(){
	
} */
/*void CFMInsurancePatientTechDetail::OnToDateKillfocus(){
	
} */
int CFMInsurancePatientTechDetail::OnToDateCheckValue(){
	return 0;
}
 
void CFMInsurancePatientTechDetail::OnPrintSelect(){
}
 
void CFMInsurancePatientTechDetail::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CExcel xls;
	CString szSQL, tmpStr;
	double nTmp = 0;
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

	xls.SetColumnWidth(0, 8);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(6, 10);
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

	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 17);
	xls.SetCellMergedColumns(0, 3, 17);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr = _T("\x42\x1EA2NG T\x1ED4NG H\x1EE2P \x43HI TI\x1EBET \x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N N\x1ED8P \x44VKT \x42N\x42HYT");
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	xls.SetCellText(0, 5, _T("STT"), 4098, true);
	xls.SetCellText(1, 5, _T("H\x1ECD t\xEAn"), 4098, true);
	xls.SetCellText(2, 5, _T("S\x1ED1 h\x1ED3 s\x1A1"), 4098, true);
	xls.SetCellText(3, 5, _T("S\x42\x41"), 4098, true);
	xls.SetCellText(4, 5, _T("Kho\x61"), 4098, true);
	xls.SetCellText(5, 5, _T("Thu\x1ED1\x63"), 4098, true);
	xls.SetCellText(6, 5, _T("\x44\x1EE5ng \x63\x1EE5"), 4098, true);
	xls.SetCellText(7, 5, _T("Ph\xF3ng \x78\x1EA1"), 4098, true);
	xls.SetCellText(8, 5, _T("\x44M"), 4098, true);
	xls.SetCellText(9, 5, _T("\x43/L PT"), 4098, true);
	xls.SetCellText(10, 5, _T("\x43ity"), 4098, true);
	xls.SetCellText(11, 5, _T("MRI"), 4098, true);
	xls.SetCellText(12, 5, _T("PET"), 4098, true);
	xls.SetCellText(13, 5, _T("\x58\x1EA1 ph\x1EABu"), 4098, true);
	xls.SetCellText(14, 5, _T("Ph\xE1 s\x1ECFi"), 4098, true);
	xls.SetCellText(15, 5, _T("\x58N"), 4098, true);
	xls.SetCellText(16, 5, _T("% \x63\xF9ng \x63hi"), 4098, true);
	xls.SetCellText(17, 5, _T("Kh\xE1\x63"), 4098, true);
	xls.SetCellText(18, 5, _T("\x43\x1ED9ng"), 4098, true);

	nRow = 6;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("patient_name")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("doc_no")), FMT_TEXT);
		xls.SetCellText(3, nRow, rs.GetValue(_T("record_no")), FMT_TEXT);
		xls.SetCellText(4, nRow, rs.GetValue(_T("dept_id")), FMT_TEXT);
		rs.GetValue(_T("drug"), nTmp);
		xls.SetCellText(5, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("material"), nTmp);
		xls.SetCellText(6, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("radio_active"), nTmp);
		xls.SetCellText(7, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("artery"), nTmp);
		xls.SetCellText(8, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("operation"), nTmp);
		xls.SetCellText(9, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("city"), nTmp);
		xls.SetCellText(10, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("mri"), nTmp);
		xls.SetCellText(11, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("pet"), nTmp);
		xls.SetCellText(12, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("radio_opt"), nTmp);
		xls.SetCellText(13, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("gravel_break"), nTmp);
		xls.SetCellText(14, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("tes"), nTmp);
		xls.SetCellText(15, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("co_paid_amount"), nTmp);
		xls.SetCellText(16, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("other"), nTmp);
		xls.SetCellText(17, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("total"), nTmp);
		xls.SetCellText(18, nRow, double2str(nTmp), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\AutoExcel.xls"));
}
 
void CFMInsurancePatientTechDetail::OnDeptListDblClick(){
	
}
 
void CFMInsurancePatientTechDetail::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
int CFMInsurancePatientTechDetail::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
}
 
long CFMInsurancePatientTechDetail::OnDeptListLoadData(){
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


CString CFMInsurancePatientTechDetail::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szWhere, szDepts, szObjects;
	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			if (!szDepts.IsEmpty())
				szDepts += _T(",");
			szDepts.AppendFormat(_T("'%s'"), m_wndDeptList.GetItemText(i, 0));
		}
	}
	szWhere.Format(_T(" AND i.hfe_class IN ('A', 'I') AND i.hfe_status = 'P' AND ho_type IN ('I', 'C')") \
				   _T(" AND f.hfe_category <> 'Y'") \
				   _T(" AND i.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!szDepts.IsEmpty())
		szWhere.AppendFormat(_T(" AND i.hfe_deptid IN (%s)"), szDepts);
	
	szSQL.Format(_T(" SELECT    Get_patientname(doc_no) patient_name, ") \
	_T("           doc_no, ") \
	_T("           hcr_recordno            record_no, ") \
	_T("           dept_id, ") \
	_T("           SUM(drug)               drug, ") \
	_T("           SUM(material)           material, ") \
	_T("           SUM(radio_active)       radio_active, ") \
	_T("           SUM(artery)             artery, ") \
	_T("           SUM(operation)          operation, ") \
	_T("           SUM(city)               city, ") \
	_T("           SUM(MRI)                MRI, ") \
	_T("           SUM(PET)                PET, ") \
	_T("           SUM(radio_opt)          radio_opt, ") \
	_T("           SUM(gravel_break)       gravel_break, ") \
	_T("           SUM(TES)                TES, ") \
	_T("           SUM(co_paid_amount)     co_paid_amount, ") \
	_T("           SUM(other)              other, ") \
	_T("           SUM(total)              total ") \
	_T(" FROM      (SELECT    i.hfe_docno                    doc_no, ") \
	_T("                      i.hfe_deptid                   dept_id, ") \
	_T("                      CASE WHEN product_org_id = 'PM' THEN CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_diffcost") \
	_T("						ELSE f.hfe_patpaid END") \
	_T("                      ELSE 0 ") \
	_T("                      END                            drug, ") \
	_T("                      CASE WHEN product_org_id = 'MA' THEN CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_diffcost") \
	_T("						ELSE f.hfe_patpaid END") \
	_T("                      ELSE 0 ") \
	_T("                      END                            material, ") \
	_T("                      0                              radio_active, ") \
	_T("                      0                              artery, ") \
	_T("                      0                              operation, ") \
	_T("                      0                              city, ") \
	_T("                      0                              MRI, ") \
	_T("                      0                              PET, ") \
	_T("                      0                              radio_opt, ") \
	_T("                      0                              gravel_break, ") \
	_T("                      0                              TES, ") \
	_T("                      CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_patpaid - f.hfe_diffcost ELSE 0 END co_paid_amount, ") \
	_T("                      0                              other, ") \
	_T("                      f.hfe_patpaid total ") \
	_T("            FROM      hms_fee_invoice i ") \
	_T("			LEFT JOIN hms_object ON (i.hfe_objectid = ho_id)") \
	_T("            LEFT JOIN hms_fee f ON ( i.hfe_invoiceno = f.hfe_invoiceno AND i.hfe_docno = f.hfe_docno ) ") \
	_T("            LEFT JOIN m_productitem_view ON ( f.hfe_itemid = product_item_id ) ") \
	_T("            WHERE     f.hfe_type IN ( 'D', 'M' ) AND f.hfe_patpaid > 0 ") \
	_T("			%s") \
	_T("            UNION ALL ") \
	_T("            SELECT    i.hfe_docno                    doc_no, ") \
	_T("                      i.hfe_deptid                   dept_id, ") \
	_T("                      0                              drug, ") \
	_T("                      0                              material, ") \
	_T("                      CASE WHEN hfl_categoryid = 5 THEN  CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_diffcost") \
	_T("						ELSE f.hfe_patpaid END") \
	_T("                      ELSE 0 ") \
	_T("                      END radio_active, ") \
	_T("                      CASE WHEN hfl_categoryid = 1 THEN  CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_diffcost") \
	_T("						ELSE f.hfe_patpaid END") \
	_T("                      ELSE 0 ") \
	_T("                      END artery, ") \
	_T("                      0                              operation, ") \
	_T("                      0                              city, ") \
	_T("                      0                              MRI, ") \
	_T("                      CASE WHEN hfl_categoryid = 2 THEN CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_diffcost") \
	_T("						ELSE f.hfe_patpaid END") \
	_T("                      ELSE 0 ") \
	_T("                      END PET, ") \
	_T("                      CASE WHEN hfl_categoryid = 3 THEN CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_diffcost") \
	_T("						ELSE f.hfe_patpaid END") \
	_T("                      ELSE 0 ") \
	_T("                      END radio_opt, ") \
	_T("                      CASE WHEN hfl_categoryid = 4 THEN CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_diffcost") \
	_T("						ELSE f.hfe_patpaid END") \
	_T("                      ELSE 0 ") \
	_T("                      END gravel_break, ") \
	_T("                      0                              TES, ") \
	_T("                      CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_patpaid - f.hfe_diffcost ELSE 0 END co_paid_amount, ") \
	_T("                      0                              other, ") \
	_T("                      f.hfe_patpaid total ") \
	_T("            FROM      hms_fee_invoice i ") \
	_T("			LEFT JOIN hms_object ON (i.hfe_objectid = ho_id)") \
	_T("            LEFT JOIN hms_fee f ON ( i.hfe_invoiceno = f.hfe_invoiceno AND i.hfe_docno = f.hfe_docno ) ") \
	_T("            LEFT JOIN hms_fee_list ON ( f.hfe_itemid = hfl_feeid ) ") \
	_T("            WHERE     f.hfe_type IN ('O', 'P', 'T' ) ") \
	_T("			%s") \
	_T("			AND f.hfe_patpaid > 0 AND Coalesce(hfl_categoryid, 0) > 0 AND Coalesce(hfl_categoryid, 0) < 6 ") \
	_T("            UNION ALL ") \
	_T("            SELECT    i.hfe_docno                    doc_no, ") \
	_T("                      i.hfe_deptid                   dept_id, ") \
	_T("                      0                              drug, ") \
	_T("                      0                              material, ") \
	_T("                      0                              radio_active, ") \
	_T("                      0                              artery, ") \
	_T("                      CASE WHEN f.hfe_type = 'O' THEN f.hfe_diffcost ") \
	_T("                      ELSE 0 ") \
	_T("                      END                            operation, ") \
	_T("                      CASE WHEN f.hfe_group = 'B2200' THEN CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_diffcost ") \
	_T("						ELSE f.hfe_patpaid END") \
	_T("                      ELSE 0 ") \
	_T("                      END                            city, ") \
	_T("                      CASE WHEN f.hfe_group = 'B2300' THEN CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_diffcost ") \
	_T("						ELSE f.hfe_patpaid END") \
	_T("                      ELSE 0 ") \
	_T("                      END                            MRI, ") \
	_T("                      0                              PET, ") \
	_T("                      0                              radio_opt, ") \
	_T("                      0                              gravel_break, ") \
	_T("                      CASE WHEN f.hfe_type = 'T' THEN CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_diffcost ") \
	_T("						ELSE f.hfe_patpaid END") \
	_T("                      ELSE 0 ") \
	_T("                      END                            TES, ") \
	_T("                      CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_patpaid - f.hfe_diffcost ELSE 0 END co_paid_amount, ") \
	_T("                      CASE WHEN ( f.hfe_group NOT IN ( 'B2200', 'B2300' ) AND f.hfe_type <> 'T' ) THEN CASE WHEN f.hfe_inspaid > 0 THEN f.hfe_diffcost ") \
	_T("						ELSE f.hfe_patpaid END") \
	_T("                      ELSE 0 ") \
	_T("                      END                            other, ") \
	_T("                      f.hfe_patpaid total ") \
	_T("            FROM      hms_fee_invoice i ") \
	_T("			LEFT JOIN hms_object ON (i.hfe_objectid = ho_id)") \
	_T("            LEFT JOIN hms_fee f ON ( i.hfe_invoiceno = f.hfe_invoiceno AND i.hfe_docno = f.hfe_docno ) ") \
	_T("            LEFT JOIN hms_fee_list ON ( f.hfe_itemid = hfl_feeid ) ") \
	_T("            WHERE     f.hfe_type IN ('O', 'P', 'T') ") \
	_T("			%s ") \
	_T("			AND f.hfe_patpaid > 0 AND Coalesce(hfl_categoryid, 0) = 0) ") \
	_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = doc_no ) ") \
	_T(" GROUP     BY doc_no, hcr_recordno, dept_id") \
	_T(" ORDER BY  dept_id, doc_no"), szWhere, szWhere, szWhere);
_fmsg(_T("%s"), szSQL);

	return szSQL;
}

