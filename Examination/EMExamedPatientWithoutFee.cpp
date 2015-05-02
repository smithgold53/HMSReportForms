#include "stdafx.h"
#include "EMExamedPatientWithoutFee.h"
#include "HMSMainFrame.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMExamedPatientWithoutFee *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMExamedPatientWithoutFee* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMExamedPatientWithoutFee *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExamedPatientWithoutFee *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMExamedPatientWithoutFee *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExamedPatientWithoutFee *)pWnd)->OnToDateCheckValue();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMExamedPatientWithoutFee *pVw = (CEMExamedPatientWithoutFee *)pWnd;
	pVw->OnExportSelect();
} 
CEMExamedPatientWithoutFee::CEMExamedPatientWithoutFee(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CEMExamedPatientWithoutFee::~CEMExamedPatientWithoutFee(){
}
void CEMExamedPatientWithoutFee::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 90);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndExport.Create(this, _T("&Export"), 350, 95, 430, 120);

}
void CEMExamedPatientWithoutFee::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);
}
void CEMExamedPatientWithoutFee::OnSetWindowEvents(){
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
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CEMExamedPatientWithoutFee::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CEMExamedPatientWithoutFee::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMExamedPatientWithoutFee::SetMode(int nMode){
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
/*void CEMExamedPatientWithoutFee::OnYearChange(){
	
} */
/*void CEMExamedPatientWithoutFee::OnYearSetfocus(){
	
} */
/*void CEMExamedPatientWithoutFee::OnYearKillfocus(){
	
} */
int CEMExamedPatientWithoutFee::OnYearCheckValue(){
	return 0;
} 
void CEMExamedPatientWithoutFee::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMExamedPatientWithoutFee::OnReportPeriodSelendok(){
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
/*void CEMExamedPatientWithoutFee::OnReportPeriodSetfocus(){
	
}*/
/*void CEMExamedPatientWithoutFee::OnReportPeriodKillfocus(){
	
}*/
long CEMExamedPatientWithoutFee::OnReportPeriodLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0){
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
/*void CEMExamedPatientWithoutFee::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMExamedPatientWithoutFee::OnFromDateChange(){
	
} */
/*void CEMExamedPatientWithoutFee::OnFromDateSetfocus(){
	
} */
/*void CEMExamedPatientWithoutFee::OnFromDateKillfocus(){
	
} */
int CEMExamedPatientWithoutFee::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMExamedPatientWithoutFee::OnToDateChange(){
	
} */
/*void CEMExamedPatientWithoutFee::OnToDateSetfocus(){
	
} */
/*void CEMExamedPatientWithoutFee::OnToDateKillfocus(){
	
} */
int CEMExamedPatientWithoutFee::OnToDateCheckValue(){
	return 0;
} 
void CEMExamedPatientWithoutFee::OnExportSelect(){
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
	xls.SetColumnWidth(1, 20);
	xls.SetColumnWidth(6, 25);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 7);
	xls.SetCellMergedColumns(0, 3, 7);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr = _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N \x110\xC3 KH\xC1M \x43H\x1AF\x41 THU TI\x1EC0N");
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	xls.SetCellText(0, 5, _T("STT"), 4098, true);
	xls.SetCellText(1, 5, _T("T\xEAn \x42N"), 4098, true);
	xls.SetCellText(2, 5, _T("S\x1ED1 HS"), 4098, true);
	xls.SetCellText(3, 5, _T("Tr\x1EA1ng th\xE1i"), 4098, true);
	xls.SetCellText(4, 5, _T("Ng\xE0y kh\xE1m"), 4098, true);
	xls.SetCellText(5, 5, _T("Ph\xF2ng"), 4098, true);
	xls.SetCellText(6, 5, _T("Ki\x1EC3u kh\xE1m"), 4098, true);
	xls.SetCellText(7, 5, _T("Ph\xED kh\xE1m"), 4098, true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("patient_name")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("doc_no")), FMT_TEXT);
		xls.SetCellText(3, nRow, rs.GetValue(_T("status")), FMT_TEXT);
		tmpStr = CDate::Convert(rs.GetValue(_T("exam_date")), yyyymmdd, ddmmyyyy);
		xls.SetCellText(4, nRow, tmpStr, FMT_TEXT);
		xls.SetCellText(5, nRow, rs.GetValue(_T("room")), FMT_TEXT);
		xls.SetCellText(6, nRow, rs.GetValue(_T("descr")), FMT_TEXT);
		xls.SetCellText(7, nRow, rs.GetValue(_T("exam_fee")), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\Danh sach bn da kham chua thu phi.xls"));
	
} 

CString CEMExamedPatientWithoutFee::GetQueryString(){
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szSQL.Format(_T(" SELECT    Get_patientname(he_docno)                    patient_name, ") \
				_T("           he_docno                                     doc_no, ") \
				_T("           Get_syssel_desc('hms_doc_status', he_status) status, ") \
				_T("           he_examdate exam_date, ") \
				_T("           hfe_desc descr, ") \
				_T("           he_deptid ") \
				_T("           ||'|' ") \
				_T("           ||he_roomid room, ") \
				_T("           hfe_cost exam_fee") \
				_T(" FROM      hms_exam ") \
				_T(" LEFT JOIN hms_fee ON ( hfe_docno = he_docno ") \
				_T("                        AND hfe_orderid = he_receptidx ) ") \
				_T(" WHERE     hfe_refstatus = 'O' ") \
				_T("       AND he_hasfee = 'Y' ") \
				_T("       AND hfe_object = 7 %s"), szWhere);
	return szSQL;
}
