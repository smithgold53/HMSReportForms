#include "stdafx.h"
#include "EMExaminationActivitiesReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMExaminationActivitiesReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMExaminationActivitiesReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMExaminationActivitiesReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExaminationActivitiesReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMExaminationActivitiesReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExaminationActivitiesReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMExaminationActivitiesReport *pVw = (CEMExaminationActivitiesReport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMExaminationActivitiesReport *pVw = (CEMExaminationActivitiesReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMExaminationActivitiesReport *pVw = (CEMExaminationActivitiesReport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddCEMExaminationActivitiesReportFnc(CWnd *pWnd){
	 return ((CEMExaminationActivitiesReport*)pWnd)->OnAddCEMExaminationActivitiesReport();
} 
static int _OnEditCEMExaminationActivitiesReportFnc(CWnd *pWnd){
	 return ((CEMExaminationActivitiesReport*)pWnd)->OnEditCEMExaminationActivitiesReport();
} 
static int _OnDeleteCEMExaminationActivitiesReportFnc(CWnd *pWnd){
	 return ((CEMExaminationActivitiesReport*)pWnd)->OnDeleteCEMExaminationActivitiesReport();
} 
static int _OnSaveCEMExaminationActivitiesReportFnc(CWnd *pWnd){
	 return ((CEMExaminationActivitiesReport*)pWnd)->OnSaveCEMExaminationActivitiesReport();
} 
static int _OnCancelCEMExaminationActivitiesReportFnc(CWnd *pWnd){
	 return ((CEMExaminationActivitiesReport*)pWnd)->OnCancelCEMExaminationActivitiesReport();
} 
CEMExaminationActivitiesReport::CEMExaminationActivitiesReport(CWnd *pParent)
	{
	m_nDlgWidth = 370;
	m_nDlgHeight = 125;
	SetDefaultValues();
	m_szTitle = _T("Examination Activities Report");
	m_bPreview = true;
}
CEMExaminationActivitiesReport::~CEMExaminationActivitiesReport(){
}
void CEMExaminationActivitiesReport::OnCreateComponents(){
	m_wndReportConditio.Create(this, _T("Report Condition"), 5, 5, 360, 90);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 180, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 185, 30, 265, 55);
	m_wndReportPeriod.Create(this,270, 30, 355, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 180, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 185, 60, 265, 85);
	m_wndToDate.Create(this,270, 60, 355, 85); 
	m_wndPrintPreview.Create(this, _T("&Print"), 205, 95, 280, 120);
	m_wndPrint.Create(this, _T("&Export"), 285, 95, 360, 120);
}
void CEMExaminationActivitiesReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(16);
//	//m_wndYear.SetCheckValue(true);
//	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);
}
void CEMExaminationActivitiesReport::OnSetWindowEvents(){
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
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString tmpStr;
	tmpStr = pMF->GetSysDate();
	m_nYear = ToInt(tmpStr.Left(4));
	m_szFromDate = m_szToDate = tmpStr;
	m_szReportPeriodKey.Format(_T("%d"), ToInt(tmpStr.Mid(5, 2)));
	SetMode(VM_EDIT);

}
void CEMExaminationActivitiesReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CEMExaminationActivitiesReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMExaminationActivitiesReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMExaminationActivitiesReport::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMExaminationActivitiesReport::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, 1, 2, -1); 
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
/*void CEMExaminationActivitiesReport::OnYearChange(){
	
} */
/*void CEMExaminationActivitiesReport::OnYearSetfocus(){
	
} */
/*void CEMExaminationActivitiesReport::OnYearKillfocus(){
	
} */
int CEMExaminationActivitiesReport::OnYearCheckValue(){
	return 0;
} 
void CEMExaminationActivitiesReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMExaminationActivitiesReport::OnReportPeriodSelendok(){
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
/*void CEMExaminationActivitiesReport::OnReportPeriodSetfocus(){
	
}*/
/*void CEMExaminationActivitiesReport::OnReportPeriodKillfocus(){
	
}*/
long CEMExaminationActivitiesReport::OnReportPeriodLoadData(){
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
/*void CEMExaminationActivitiesReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMExaminationActivitiesReport::OnFromDateChange(){
	
} */
/*void CEMExaminationActivitiesReport::OnFromDateSetfocus(){
	
} */
/*void CEMExaminationActivitiesReport::OnFromDateKillfocus(){
	
} */
int CEMExaminationActivitiesReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMExaminationActivitiesReport::OnToDateChange(){
	
} */
/*void CEMExaminationActivitiesReport::OnToDateSetfocus(){
	
} */
/*void CEMExaminationActivitiesReport::OnToDateKillfocus(){
	
} */
int CEMExaminationActivitiesReport::OnToDateCheckValue(){
	return 0;
} 
void CEMExaminationActivitiesReport::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	m_bPreview = true;
	PrintExaminationActivitiesReport(m_szFromDate, m_szToDate, m_bPreview);
	UpdateData(false);
} 
void CEMExaminationActivitiesReport::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);	
	ExportExaminationActivitiesReport();
	UpdateData(false);
} 
void CEMExaminationActivitiesReport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CEMExaminationActivitiesReport::OnAddCEMExaminationActivitiesReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMExaminationActivitiesReport::OnEditCEMExaminationActivitiesReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMExaminationActivitiesReport::OnDeleteCEMExaminationActivitiesReport(){
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
 		OnCancelCEMExaminationActivitiesReport(); 
 	}
	return 0;
}
int CEMExaminationActivitiesReport::OnSaveCEMExaminationActivitiesReport(){
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
 		//OnCEMExaminationActivitiesReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMExaminationActivitiesReport::OnCancelCEMExaminationActivitiesReport(){
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
int CEMExaminationActivitiesReport::OnCEMExaminationActivitiesReportListLoadData(){
	return 0;
}
void CEMExaminationActivitiesReport::PrintExaminationActivitiesReport(CString szFromDate, CString szToDate, bool bPreview)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szFromDateLabel, szToDateLabel, szWhere, szAmount;
	CString szDate, szSysDate;	
	CReport rpt; 

	if(!rpt.Init(_T("Reports/HMS/HE_EXAMINATIONACTIVITIESREPORT.RPT")))
		return ;
	BeginWaitCursor();
	szSysDate = pMF->GetSysDate(); 
	szFromDate = m_szFromDate;
	szToDate = m_szToDate;
	szSQL = GetQueryString();
	
	_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);

	//Report Header
	tmpStr = rpt.GetReportHeader()->GetValue(_T("ReportDate"));
	szDate.Format(tmpStr, CDate::Convert(szFromDate, yyyymmdd, ddmmyyyy),CDate::Convert(szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);

	//Report Detail
	CReportSection* rptDetail = rpt.GetDetail(); 
	int nIndex = 1;
	long nTotal[10];
	for (int i =0; i < 9; i++)
		nTotal[i] = 0;

	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();	
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("Index"), tmpStr);

		rs.GetValue(_T("roomname"), tmpStr);
		rptDetail->SetValue(_T("roomname"), tmpStr);

		rs.GetValue(_T("totalExam"), tmpStr);
		nTotal[0] += ToInt(tmpStr);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("insexam"), tmpStr);
		nTotal[1] += ToInt(tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("servexam"), tmpStr);
		nTotal[2] += ToInt(tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("discountexam"), tmpStr);
		nTotal[3] += ToInt(tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("emergencyexam"), tmpStr);
		nTotal[4] += ToInt(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("admission"), tmpStr);
		nTotal[5] += ToInt(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("transfer"), tmpStr);
		nTotal[6] += ToInt(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("outpatient"), tmpStr);
		nTotal[7] += ToInt(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("treatdays"), tmpStr);
		nTotal[8] += ToInt(tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.MoveNext();
	}
	rptDetail = rpt.AddDetail();
	rptDetail->GetItem(_T("roomname"))->SetBold(true);
	TranslateString(_T("Total"),tmpStr);
	rptDetail->SetValue(_T("roomname"),tmpStr);

	for (int i =0; i < 9; i++)
	{
		tmpStr.Format(_T("%d"), i+1);
		szAmount.Format(_T("%ld"),nTotal[i]);
		rptDetail->GetItem(tmpStr)->SetBold(true);
		rptDetail->SetValue(tmpStr, szAmount);
	}

	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	
	rpt.PrintPreview();
	EndWaitCursor();
}
void CEMExaminationActivitiesReport::ExportExaminationActivitiesReport()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szFromDateLabel, szToDateLabel, szWhere, szAmount;
	CString szDate, szSysDate;	
	CReport rpt; 
	
	szSysDate = pMF->GetSysDateTime(); 
	BeginWaitCursor();
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);

	_fmsg(_T("%s"), szSQL);

	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 35);
	xls.SetColumnWidth(2, 14);
	xls.SetColumnWidth(3, 14);
	xls.SetColumnWidth(4, 14);
	xls.SetColumnWidth(5, 14);
	xls.SetColumnWidth(6, 14);
	xls.SetColumnWidth(7, 14);
	xls.SetColumnWidth(8, 14);
	xls.SetColumnWidth(9, 14);
	xls.SetColumnWidth(10, 14);
	xls.SetColumnWidth(11, 14);
	xls.SetColumnWidth(12, 14);
		
	int nCol = 0;
	int nRow = 0;
	int age;
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellMergedColumns(nCol, nRow + 1, 2);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellMergedColumns(nCol, nCol + 3, 12);	
	xls.SetCellText(nCol, nCol + 3, _T("\x42\xC1O \x43\xC1O HO\x1EA0T \x110\x1ED8NG \x110I\x1EC0U TR\x1ECA"), FMT_TEXT | FMT_CENTER, true, 16);
	xls.SetCellMergedColumns(nCol, nCol + 4, 12);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(nCol, nCol + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12);
	
	age=0;

	xls.SetCellMergedRows(nCol, nRow + 6, 2);
	xls.SetCellText(nCol, nRow + 6, _T("STT"), FMT_TEXT | FMT_CENTER, true);	

	xls.SetCellMergedRows(nCol + 1, nRow + 6, 2);
	TranslateString(_T("Exam Room"), tmpStr);
	xls.SetCellText(nCol + 1, nRow + 6, tmpStr, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedRows(nCol + 2, nRow + 6, 2);
	TranslateString(_T("Total"), tmpStr);	
	xls.SetCellText(nCol + 2, nRow + 6, tmpStr, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol + 3, nRow + 6, 3);
	xls.SetCellText(nCol + 3, nRow + 6, _T("S\x1ED1 l\x1EA7n kh\xE1m"), FMT_TEXT | FMT_CENTER, true);
	TranslateString(_T("Insurance"), tmpStr);
	xls.SetCellText(nCol + 3, nRow + 7, tmpStr, FMT_TEXT | FMT_CENTER, true);

	TranslateString(_T("Service"), tmpStr);
	xls.SetCellText(nCol + 4, nRow + 7, tmpStr, FMT_TEXT | FMT_CENTER, true);	

	TranslateString(_T("UnCollected"), tmpStr);
	xls.SetCellText(nCol + 5, nRow + 7, tmpStr, FMT_TEXT | FMT_CENTER, true);	

	xls.SetCellMergedRows(nCol + 6, nRow + 6, 2);
	TranslateString(_T("Emergency"), tmpStr);
	xls.SetCellText(nCol + 6, nRow + 6, tmpStr, FMT_TEXT | FMT_CENTER, true);	
	
	xls.SetCellMergedRows(nCol + 7, nRow + 6, 2);
	TranslateString(_T("Admission"), tmpStr);
	xls.SetCellText(nCol + 7, nRow + 6, tmpStr, FMT_TEXT | FMT_CENTER, true);	

	xls.SetCellMergedRows(nCol + 7, nRow + 6, 2);
	TranslateString(_T("Transfer Hospital"), tmpStr);
	xls.SetCellText(nCol + 8, nRow + 6, tmpStr, FMT_TEXT | FMT_CENTER, true);	

	xls.SetCellMergedColumns(nCol + 9, nRow + 6, 2);
	xls.SetCellText(nCol + 9, nRow + 6, _T("\x110i\x1EC1u tr\x1ECB ngo\x1EA1i tr\xFA"), FMT_TEXT | FMT_CENTER, true);
	TranslateString(_T("Outpatient"), tmpStr);
	xls.SetCellText(nCol + 9, nRow + 7, tmpStr, FMT_TEXT | FMT_CENTER, true);

	TranslateString(_T("Treat Days"), tmpStr);
	xls.SetCellText(nCol + 10, nRow + 7, tmpStr, FMT_TEXT | FMT_CENTER, true);	
	
	int nIndex = 1, SongayDT=0;
	int i=0,nTotal[11];
	for (i=0; i<11; i++)
	{
		nTotal[i]=0;
	}
	while(!rs.IsEOF())
	{
		nRow++;		
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow + 7, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("roomname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow + 7, tmpStr, FMT_TEXT);

		rs.GetValue(_T("totalexam"), tmpStr);
		nTotal[2] += ToInt(tmpStr);
		xls.SetCellText(nCol + 2, nRow + 7, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("insexam"), tmpStr);
		nTotal[3] += ToInt(tmpStr);
		xls.SetCellText(nCol + 3, nRow + 7, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("servexam"), tmpStr);
		nTotal[4] += ToInt(tmpStr);
		xls.SetCellText(nCol + 4, nRow + 7, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("discountexam"), tmpStr);
		nTotal[5] += ToInt(tmpStr);
		xls.SetCellText(nCol + 5, nRow + 7, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("emergencyexam"), tmpStr);
		nTotal[6] += ToInt(tmpStr);
		xls.SetCellText(nCol + 6, nRow + 7, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("admission"), tmpStr);		
		nTotal[7] += ToInt(tmpStr);
		tmpStr.Format(_T("%ld"), SongayDT);
		xls.SetCellText(nCol + 7, nRow + 7, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("Transfer"), tmpStr);
		nTotal[8] += ToInt(tmpStr);
		xls.SetCellText(nCol + 8, nRow + 7, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("outpatient"), tmpStr);
		nTotal[9] += ToInt(tmpStr);
		xls.SetCellText(nCol + 9, nRow + 7, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("treatdays"), tmpStr);
		nTotal[10] += ToInt(tmpStr);		
		xls.SetCellText(nCol + 10, nRow + 7, tmpStr, FMT_NUMBER1);		

		rs.MoveNext();
	}
	nRow++;
	xls.SetCellMergedColumns(nCol, nRow + 7, 2);
	TranslateString(_T("Total"), tmpStr);
	xls.SetCellText(nCol, nRow + 7, tmpStr, FMT_TEXT | FMT_CENTER, true);
	for (int i = 2; i <= 10; i++)
	{		
		tmpStr.Format(_T("%ld"), nTotal[i]);		
		xls.SetCellText(i, nRow + 7, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\BAO CAO HOAT DONG KHAM BENH.xls"));
	EndWaitCursor();
} // chua cho hien thi m_szDept o CMaiFrame

CString CEMExaminationActivitiesReport::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere, m_szDept;
	szWhere.Empty();
	szSQL.Format(_T(" SELECT 	deptid, roomid, hrl_name as roomname, ") \
				_T(" 	sum(insuranceexam+serviceexam+discountexam) as totalexam, ") \
				_T(" 	sum(insuranceexam) as insexam,") \
				_T(" 	sum(serviceexam) as servexam,") \
				_T(" 	sum(discountexam) as discountexam,") \
				_T(" 	sum(emergencyexam) as emergencyexam,") \
				_T(" 	sum(admission) as admission,") \
				_T(" 	sum(transfer) as transfer,") \
				_T(" 	sum(outpatient) as outpatient,") \
				_T(" 	sum(treatdays) as treatdays") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT	hd_admitdept as deptid, he_roomid as roomid,") \
				_T(" 	case when ho_type in('I','C') and hd_status = 'T' and (he_examdate) between cast_string2timestamp('%s') and cast_string2timestamp('%s') then 1 else 0 end as insuranceexam,") \
				_T(" 	case when ho_type in('S') and hd_status = 'T' and (he_examdate) between cast_string2timestamp('%s') and cast_string2timestamp('%s') then 1 else 0 end as serviceexam,") \
				_T(" 	case when ho_type in('D') and hd_status = 'T' and (he_examdate) between cast_string2timestamp('%s') and cast_string2timestamp('%s') then 1 else 0 end as discountexam,") \
				_T(" 	case when he_emergency='Y' and hd_status = 'T' and (he_examdate) between cast_string2timestamp('%s') and cast_string2timestamp('%s') then 1 else 0 end as emergencyexam,") \
				_T(" 	case when hd_suggestion='D' and hd_status='T' and (hd_enddate) between cast_string2timestamp('%s') and cast_string2timestamp('%s') then 1 else 0 end as admission,") \
				_T(" 	case when hd_suggestion='F' and hd_status='T' and (hd_enddate) between cast_string2timestamp('%s') and cast_string2timestamp('%s') then 1 else 0 end as transfer,") \
				_T(" 	case when hd_suggestion='E' and hd_status='T' and (hd_enddate) between cast_string2timestamp('%s') and cast_string2timestamp('%s') then 1 else 0 end as outpatient,") \
				_T(" 	case when hd_suggestion='E' and hd_status='T' and (hd_enddate) between cast_string2timestamp('%s') and cast_string2timestamp('%s') then hop_treatdays else 0 end as treatdays ") \
				_T(" FROM hms_doc ") \
				_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno) ") \
				_T(" LEFT JOIN hms_exam ON (he_docno = hd_docno)") \
				_T(" LEFT JOIN hms_object ON(hd_object=ho_id)") \
				_T(" LEFT JOIN hms_outpatient ON(hop_docno=hd_docno)") \
				_T(" WHERE 	 hd_admitdept='%s' ") \
				_T(" ) tbl") \
				_T(" LEFT JOIN hms_roomlist ON(deptid=hrl_deptid and roomid=hrl_id)") \
				_T(" WHERE hrl_id > 0 ") \
				_T(" GROUP BY deptid, roomid, hrl_name ") \
				_T(" ORDER BY deptid, roomid"), m_szFromDate, m_szToDate,m_szFromDate, m_szToDate,m_szFromDate, m_szToDate,m_szFromDate, m_szToDate,
				m_szFromDate, m_szToDate,m_szFromDate, m_szToDate,m_szFromDate, m_szToDate,m_szFromDate, m_szToDate, pMF->m_szDept); 
	//m_szFromDate, m_szToDate,m_szFromDate, m_szToDate,m_szFromDate, m_szToDate,m_szFromDate, m_szToDate, m_szDept); 
	return szSQL;

}

