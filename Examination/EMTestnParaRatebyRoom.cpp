#include "stdafx.h"
#include "EMTestnParaRatebyRoom.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMTestnParaRatebyRoom *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMTestnParaRatebyRoom* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMTestnParaRatebyRoom *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMTestnParaRatebyRoom *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMTestnParaRatebyRoom *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMTestnParaRatebyRoom *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMTestnParaRatebyRoom *pVw = (CEMTestnParaRatebyRoom *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMTestnParaRatebyRoom *pVw = (CEMTestnParaRatebyRoom *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMTestnParaRatebyRoom *pVw = (CEMTestnParaRatebyRoom *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMTestnParaRatebyRoom*)pWnd)->OnAddEMStatisticsGeneralPatientReport();
} 
static int _OnEditEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMTestnParaRatebyRoom*)pWnd)->OnEditEMStatisticsGeneralPatientReport();
} 
static int _OnDeleteEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMTestnParaRatebyRoom*)pWnd)->OnDeleteEMStatisticsGeneralPatientReport();
} 
static int _OnSaveEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMTestnParaRatebyRoom*)pWnd)->OnSaveEMStatisticsGeneralPatientReport();
} 
static int _OnCancelEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMTestnParaRatebyRoom*)pWnd)->OnCancelEMStatisticsGeneralPatientReport();
} 
CEMTestnParaRatebyRoom::CEMTestnParaRatebyRoom(CWnd *pParent)
{
	m_szTitle = _T("Patients General Statistics");
	m_nDlgWidth = 420;
	m_nDlgHeight = 125;
	SetDefaultValues();
}
CEMTestnParaRatebyRoom::~CEMTestnParaRatebyRoom()
{
}
void CEMTestnParaRatebyRoom::OnCreateComponents(){
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
void CEMTestnParaRatebyRoom::OnInitializeComponents(){
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
void CEMTestnParaRatebyRoom::OnSetWindowEvents(){
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
void CEMTestnParaRatebyRoom::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
}
void CEMTestnParaRatebyRoom::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CEMTestnParaRatebyRoom::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMTestnParaRatebyRoom::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMTestnParaRatebyRoom::SetMode(int nMode){
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
/*void CEMTestnParaRatebyRoom::OnYearChange(){
	
} */
/*void CEMTestnParaRatebyRoom::OnYearSetfocus(){
	
} */
/*void CEMTestnParaRatebyRoom::OnYearKillfocus(){
	
} */
int CEMTestnParaRatebyRoom::OnYearCheckValue(){
	return 0;
} 
void CEMTestnParaRatebyRoom::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMTestnParaRatebyRoom::OnReportPeriodSelendok(){
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
/*void CEMTestnParaRatebyRoom::OnReportPeriodSetfocus(){
	
}*/
/*void CEMTestnParaRatebyRoom::OnReportPeriodKillfocus(){
	
}*/
long CEMTestnParaRatebyRoom::OnReportPeriodLoadData()
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
/*void CEMTestnParaRatebyRoom::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMTestnParaRatebyRoom::OnFromDateChange(){
	
} */
/*void CEMTestnParaRatebyRoom::OnFromDateSetfocus(){
	
} */
/*void CEMTestnParaRatebyRoom::OnFromDateKillfocus(){
	
} */
int CEMTestnParaRatebyRoom::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMTestnParaRatebyRoom::OnToDateChange(){
	
} */
/*void CEMTestnParaRatebyRoom::OnToDateSetfocus(){
	
} */
/*void CEMTestnParaRatebyRoom::OnToDateKillfocus(){
	
} */
int CEMTestnParaRatebyRoom::OnToDateCheckValue(){
	return 0;
} 
void CEMTestnParaRatebyRoom::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	if(!rpt.Init(_T("Reports/HMS/HE_THONGKEBENHNHANCAPTUONG.RPT")) )
	return ;
	BeginWaitCursor();
	UpdateData(true);
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	rs.GetValue(_T("tenkhoa"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);

	int nIndex = 1;
	CReportSection* rptDetail;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%ld"), nIndex ++);
		rptDetail->SetValue(_T("0"), tmpStr);
		
		rs.GetValue(_T("docno"),tmpStr);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("tenBN"),tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("tencapbac"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("chucvu"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("donvi"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		
		rs.GetValue(_T("tendoituong"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("icd10"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		
		rs.GetValue(_T("benhchinh"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.MoveNext();	
	}	
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();
	
} 
void CEMTestnParaRatebyRoom::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CString szSQL, tmpStr;
	int nTmp = 0, nQty = 0;
	double nRate = 0;
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

	xls.SetColumnWidth(0, 30);
	//xls.SetColumnWidth(1, 10);
	//xls.SetColumnWidth(2, 25);
	//xls.SetColumnWidth(3, 12);
	//xls.SetColumnWidth(4, 23);
	//xls.SetColumnWidth(5, 23);
	//xls.SetColumnWidth(6, 12);
	//xls.SetColumnWidth(7, 12);
	//xls.SetColumnWidth(8, 66);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 6);
	xls.SetCellText(nCol, nRow + 3, _T("TH\x1ED0NG K\xCA T\x1EF6 L\x1EC6 \x58\xC9T NGHI\x1EC6M \x43\x1A0 \x42\x1EA2N"), FMT_TEXT | FMT_CENTER, true, 14);
	xls.SetCellMergedColumns(nCol, nRow + 4, 6);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol, nRow + 5, _T("Ph\xF2ng kh\xE1m"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);
	xls.SetCellText(nCol + 1, nRow + 5, _T("T\x1ED5ng s\x1ED1 kh\xE1m"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);
	xls.SetCellText(nCol + 2, nRow + 5, _T("Sinh h\xF3\x61"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);
	xls.SetCellText(nCol + 3, nRow + 5, _T("Huy\x1EBFt h\x1ECD\x63"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);
	xls.SetCellText(nCol + 4, nRow + 5, _T("\x43h\x1EA9n \x111o\xE1n \x63h\x1EE9\x63 n\x103ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);
	xls.SetCellText(nCol + 5, nRow + 5, _T("\x58 Qu\x61ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);

	nRow = 6;
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("roomname"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr);
		rs.GetValue(_T("examed"), nTmp);
		xls.SetCellText(nCol + 1, nRow, int2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("sh"), nQty);
		rs.GetValue(_T("shr"), nRate);
		tmpStr.Format(_T("%d (%.2f )"),nQty, nRate);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
		rs.GetValue(_T("hh"), nQty);
		rs.GetValue(_T("hhr"), nRate);
		tmpStr.Format(_T("%d (%.2f )"),nQty, nRate);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
		rs.GetValue(_T("td"), nQty);
		rs.GetValue(_T("tdr"), nRate);
		tmpStr.Format(_T("%d (%.2f )"),nQty, nRate);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
		rs.GetValue(_T("xq"), nQty);
		rs.GetValue(_T("xqr"), nRate);
		tmpStr.Format(_T("%d (%.2f )"),nQty, nRate);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
		nRow++;
		rs.MoveNext();		
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\THONG KE TY LE XET NGHIEM CO BAN.xls"));
} 
void CEMTestnParaRatebyRoom::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CEMTestnParaRatebyRoom::OnAddEMStatisticsGeneralPatientReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMTestnParaRatebyRoom::OnEditEMStatisticsGeneralPatientReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMTestnParaRatebyRoom::OnDeleteEMStatisticsGeneralPatientReport(){
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
int CEMTestnParaRatebyRoom::OnSaveEMStatisticsGeneralPatientReport(){
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
int CEMTestnParaRatebyRoom::OnCancelEMStatisticsGeneralPatientReport(){
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
int CEMTestnParaRatebyRoom::OnEMStatisticsGeneralPatientReportListLoadData(){
	return 0;
}
CString CEMTestnParaRatebyRoom::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString szSQL, szWhere, szInnerWhere;
	szInnerWhere.Format(_T(" AND hpc_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	szInnerWhere.AppendFormat(_T(" AND hpc_orderdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szWhere.Format(_T(" AND he_examdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szWhere.AppendFormat(_T(" AND he_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	szSQL.Format(_T(" SELECT") \
				_T("   roomname,") \
				_T("   sh,") \
				_T("   hh,") \
				_T("   xq,") \
				_T("   td,") \
				_T("   examed,") \
				_T("   sh*100/examed AS shr,") \
				_T("   hh*100/examed AS hhr,") \
				_T("   xq*100/examed AS xqr,") \
				_T("   td*100/examed AS tdr") \
				_T(" FROM   (SELECT") \
				_T("           he_roomid,") \
				_T("           (SELECT") \
				_T("              hrl_name") \
				_T("            FROM   hms_roomlist") \
				_T("            WHERE  hrl_id=he_roomid") \
				_T("                   AND hrl_deptid='%s') AS roomname,") \
				_T("           Count(he_docno)                AS examed") \
				_T("         FROM   hms_exam") \
				_T("         WHERE  he_status IN ('P', 'T') %s") \
				_T("         GROUP  BY he_roomid,") \
				_T("                   he_deptid") \
				_T("         ORDER  BY he_roomid) tbl2") \
				_T("        LEFT JOIN (SELECT") \
				_T("                     hpc_roomid,") \
				_T("                     Sum(sh) AS sh,") \
				_T("                     Sum(hh) AS hh,") \
				_T("                     Sum(xq) AS xq,") \
				_T("                     Sum(td) AS td") \
				_T("                   FROM   (SELECT") \
				_T("                             hpc_roomid,") \
				_T("                             CASE WHEN Substr(hpc_groupid, 1, 3)='B12' THEN 1") \
				_T("                             ELSE 0") \
				_T("                             END AS sh,") \
				_T("                             CASE WHEN Substr(hpc_groupid, 1, 3)='B11' THEN 1") \
				_T("                             ELSE 0") \
				_T("                             END AS hh,") \
				_T("                             CASE WHEN Substr(hpc_groupid, 1, 3)='B21' THEN 1") \
				_T("                             ELSE 0") \
				_T("                             END AS xq,") \
				_T("                             CASE WHEN hpc_groupid='B3' THEN 1") \
				_T("                             ELSE 0") \
				_T("                             END AS td") \
				_T("                           FROM   (SELECT") \
				_T("                                     hpc_roomid,") \
				_T("                                     hpc_docno,") \
				_T("                                     hpc_groupid") \
				_T("                                   FROM   hms_testorder") \
				_T("                                   WHERE  hpc_status<>'O' %s") \
				_T("                                   GROUP  BY hpc_roomid,") \
				_T("                                             hpc_groupid,") \
				_T("                                             hpc_docno") \
				_T("                                   UNION ALL") \
				_T("                                   SELECT") \
				_T("                                     hpc_roomid,") \
				_T("                                     hpc_docno,") \
				_T("                                     Substr(hpc_groupid, 1, 2)") \
				_T("                                   FROM   hms_pacsorder") \
				_T("                                   WHERE  hpc_status<>'O' %s") \
				_T("                                          AND Substr(hpc_groupid, 1, 2)='B3'") \
				_T("                                   GROUP  BY hpc_roomid,") \
				_T("                                             Substr(hpc_groupid, 1, 2),") \
				_T("                                             hpc_docno") \
				_T("                                   UNION ALL") \
				_T("                                   SELECT") \
				_T("                                     hpc_roomid,") \
				_T("                                     hpc_docno,") \
				_T("                                     hpc_groupid") \
				_T("                                   FROM   hms_pacsorder") \
				_T("                                   WHERE  hpc_status<>'O' %s") \
				_T("                                          AND Substr(hpc_groupid, 1, 2)<>'B3'") \
				_T("                                   GROUP  BY hpc_roomid,") \
				_T("                                             hpc_groupid,") \
				_T("                                             hpc_docno)tmp) tbl0") \
				_T("                   GROUP  BY hpc_roomid") \
				_T("                   ORDER  BY hpc_roomid) tbl1 ON (hpc_roomid=he_roomid) ") \
				_T(" ORDER BY cast(substr(roomname, 1, 2) as integer)"),pMF->GetCurrentDepartmentID(), szWhere, szInnerWhere, szInnerWhere, szInnerWhere);
	return szSQL;
}
