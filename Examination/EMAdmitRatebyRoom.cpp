#include "stdafx.h"
#include "EMAdmitRatebyRoom.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMAdmitRatebyRoom *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMAdmitRatebyRoom* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMAdmitRatebyRoom *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMAdmitRatebyRoom *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMAdmitRatebyRoom *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMAdmitRatebyRoom *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMAdmitRatebyRoom *pVw = (CEMAdmitRatebyRoom *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMAdmitRatebyRoom *pVw = (CEMAdmitRatebyRoom *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMAdmitRatebyRoom *pVw = (CEMAdmitRatebyRoom *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMAdmitRatebyRoom*)pWnd)->OnAddEMStatisticsGeneralPatientReport();
} 
static int _OnEditEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMAdmitRatebyRoom*)pWnd)->OnEditEMStatisticsGeneralPatientReport();
} 
static int _OnDeleteEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMAdmitRatebyRoom*)pWnd)->OnDeleteEMStatisticsGeneralPatientReport();
} 
static int _OnSaveEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMAdmitRatebyRoom*)pWnd)->OnSaveEMStatisticsGeneralPatientReport();
} 
static int _OnCancelEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMAdmitRatebyRoom*)pWnd)->OnCancelEMStatisticsGeneralPatientReport();
} 
CEMAdmitRatebyRoom::CEMAdmitRatebyRoom(CWnd *pParent)
{
	m_szTitle = _T("Patients General Statistics");
	m_nDlgWidth = 420;
	m_nDlgHeight = 125;
	SetDefaultValues();
}
CEMAdmitRatebyRoom::~CEMAdmitRatebyRoom()
{
}
void CEMAdmitRatebyRoom::OnCreateComponents(){
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
void CEMAdmitRatebyRoom::OnInitializeComponents(){
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
void CEMAdmitRatebyRoom::OnSetWindowEvents(){
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
void CEMAdmitRatebyRoom::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
}
void CEMAdmitRatebyRoom::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CEMAdmitRatebyRoom::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMAdmitRatebyRoom::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMAdmitRatebyRoom::SetMode(int nMode){
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
/*void CEMAdmitRatebyRoom::OnYearChange(){
	
} */
/*void CEMAdmitRatebyRoom::OnYearSetfocus(){
	
} */
/*void CEMAdmitRatebyRoom::OnYearKillfocus(){
	
} */
int CEMAdmitRatebyRoom::OnYearCheckValue(){
	return 0;
} 
void CEMAdmitRatebyRoom::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMAdmitRatebyRoom::OnReportPeriodSelendok(){
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
/*void CEMAdmitRatebyRoom::OnReportPeriodSetfocus(){
	
}*/
/*void CEMAdmitRatebyRoom::OnReportPeriodKillfocus(){
	
}*/
long CEMAdmitRatebyRoom::OnReportPeriodLoadData()
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
/*void CEMAdmitRatebyRoom::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMAdmitRatebyRoom::OnFromDateChange(){
	
} */
/*void CEMAdmitRatebyRoom::OnFromDateSetfocus(){
	
} */
/*void CEMAdmitRatebyRoom::OnFromDateKillfocus(){
	
} */
int CEMAdmitRatebyRoom::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMAdmitRatebyRoom::OnToDateChange(){
	
} */
/*void CEMAdmitRatebyRoom::OnToDateSetfocus(){
	
} */
/*void CEMAdmitRatebyRoom::OnToDateKillfocus(){
	
} */
int CEMAdmitRatebyRoom::OnToDateCheckValue(){
	return 0;
} 
void CEMAdmitRatebyRoom::OnPrintSelect(){
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
void CEMAdmitRatebyRoom::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CString szSQL, tmpStr;
	int nTmp = 0;
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
	xls.SetColumnWidth(1, 15);
	xls.SetColumnWidth(2, 15);
	xls.SetColumnWidth(3, 15);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 1);
	xls.SetCellMergedColumns(nCol, nRow + 2, 1);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 4);
	xls.SetCellText(nCol, nRow + 3, _T("TH\x1ED0NG K\xCA T\x1EF6 L\x1EC6 \x42\x1EC6NH NH\xC2N NH\x1EACP VI\x1EC6N"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 4);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol, nRow + 5, _T("Ph\xF2ng kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 5, _T("T\x1ED5ng s\x1ED1 kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 5, _T("S\x1ED1 nh\x1EADp vi\x1EC7n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 3, nRow + 5, _T("T\x1EF7 l\x1EC7 %"), FMT_TEXT | FMT_CENTER, true, 11);
	nRow = 6;
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("nme"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr);
		rs.GetValue(_T("examed"), nTmp);
		FormatCurrency(nTmp, tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
		rs.GetValue(_T("inward"), nTmp);
		FormatCurrency(nTmp, tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
		rs.GetValue(_T("rate"), nRate);
		tmpStr.Format(_T("%.2f"), nRate);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
		nRow++;
		rs.MoveNext();
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\THONG KE TY LE BENH NHAN NHAP VIEN.xls"));
} 
void CEMAdmitRatebyRoom::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CEMAdmitRatebyRoom::OnAddEMStatisticsGeneralPatientReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMAdmitRatebyRoom::OnEditEMStatisticsGeneralPatientReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMAdmitRatebyRoom::OnDeleteEMStatisticsGeneralPatientReport(){
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
int CEMAdmitRatebyRoom::OnSaveEMStatisticsGeneralPatientReport(){
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
int CEMAdmitRatebyRoom::OnCancelEMStatisticsGeneralPatientReport(){
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
int CEMAdmitRatebyRoom::OnEMStatisticsGeneralPatientReportListLoadData(){
	return 0;
}
CString CEMAdmitRatebyRoom::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND he_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	szWhere.AppendFormat(_T(" AND hd_enddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szSQL.Format(_T(" SELECT hrl_name as nme,he_roomid, ") \
				_T("   Sum(inward) AS inward,") \
				_T("   Sum(examed) AS examed,") \
				_T("   sum(inward)/sum(examed) as rate") \
				_T(" FROM   (SELECT") \
				_T("           he_roomid,") \
				_T("           CASE WHEN hd_suggestion='D'") \
				_T("                AND he_doctor=hd_doctor THEN 1") \
				_T("           END AS inward,") \
				_T("           1   AS examed") \
				_T("         FROM   hms_exam") \
				_T("         LEFT JOIN hms_doc ON (he_docno=hd_docno)") \
				_T("         WHERE  he_status IN ('P', 'T') %s") \
				_T("        )") \
				_T(" LEFT JOIN hms_roomlist ON (he_roomid = hrl_id AND hrl_deptid = '%s')") \
				_T(" GROUP  BY hrl_name, he_roomid ") \
				_T(" ORDER  BY cast(substr(hrl_name, 1, 2) as integer) "), szWhere, pMF->GetCurrentDepartmentID());
	return szSQL;
}
