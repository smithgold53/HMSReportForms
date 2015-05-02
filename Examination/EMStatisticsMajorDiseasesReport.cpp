#include "stdafx.h"
#include "EMStatisticsMajorDiseasesReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMStatisticsMajorDiseasesReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMStatisticsMajorDiseasesReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMStatisticsMajorDiseasesReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMStatisticsMajorDiseasesReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMStatisticsMajorDiseasesReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnExamRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMStatisticsMajorDiseasesReport* )pWnd)->OnExamRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnExamRoomSelendokFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnExamRoomSelendok();
}
/*static void _OnExamRoomSetfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnExamRoomKillfocus();
}*/
/*static void _OnExamRoomKillfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnExamRoomKillfocus();
}*/
static long _OnExamRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMStatisticsMajorDiseasesReport *)pWnd)->OnExamRoomLoadData();
}
/*static void _OnExamRoomAddNewFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReport *)pWnd)->OnExamRoomAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMStatisticsMajorDiseasesReport *pVw = (CEMStatisticsMajorDiseasesReport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMStatisticsMajorDiseasesReport *pVw = (CEMStatisticsMajorDiseasesReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMStatisticsMajorDiseasesReport *pVw = (CEMStatisticsMajorDiseasesReport *)pWnd;
	pVw->OnCloseSelect();
}
static void _OnSelectAllDiseasesSelectFnc(CWnd *pWnd){
	 ((CEMStatisticsMajorDiseasesReport*)pWnd)->OnSelectAllDiseasesSelect();
}
static int _OnAddEMStatisticsMajorDiseasesReportFnc(CWnd *pWnd){
	 return ((CEMStatisticsMajorDiseasesReport*)pWnd)->OnAddEMStatisticsMajorDiseasesReport();
} 
static int _OnEditEMStatisticsMajorDiseasesReportFnc(CWnd *pWnd){
	 return ((CEMStatisticsMajorDiseasesReport*)pWnd)->OnEditEMStatisticsMajorDiseasesReport();
} 
static int _OnDeleteEMStatisticsMajorDiseasesReportFnc(CWnd *pWnd){
	 return ((CEMStatisticsMajorDiseasesReport*)pWnd)->OnDeleteEMStatisticsMajorDiseasesReport();
} 
static int _OnSaveEMStatisticsMajorDiseasesReportFnc(CWnd *pWnd){
	 return ((CEMStatisticsMajorDiseasesReport*)pWnd)->OnSaveEMStatisticsMajorDiseasesReport();
} 
static int _OnCancelEMStatisticsMajorDiseasesReportFnc(CWnd *pWnd){
	 return ((CEMStatisticsMajorDiseasesReport*)pWnd)->OnCancelEMStatisticsMajorDiseasesReport();
} 
CEMStatisticsMajorDiseasesReport::CEMStatisticsMajorDiseasesReport(CWnd *pParent)
{
	m_nDlgWidth = 440;
	m_nDlgHeight = 190;
	m_szTitle = _T("Diseases Suffered Report");
	SetDefaultValues();
}
CEMStatisticsMajorDiseasesReport::~CEMStatisticsMajorDiseasesReport()
{
}
void CEMStatisticsMajorDiseasesReport::OnCreateComponents()
{
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 150);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndExamRoomLabel.Create(this, _T("Exam Room"), 10, 90, 90, 115);
	m_wndExamRoom.Create(this,95, 90, 425, 115); 
	m_wndSelectAllDiseases.Create(this, _T("Select All Diseases"), 295, 120, 425, 145);
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 240, 155, 350, 180);
	m_wndExport.Create(this, _T("&Export"), 355, 155, 430, 180);
}
void CEMStatisticsMajorDiseasesReport::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	////m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	//m_wndExamRoom.SetCheckValue(true);
	m_wndExamRoom.LimitText(81);
	//m_wndExamRoom.SetCheckBox(TRUE);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndExamRoom.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 50);
	m_wndExamRoom.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

}
void CEMStatisticsMajorDiseasesReport::OnSetWindowEvents()
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
	m_wndExamRoom.SetEvent(WE_SELENDOK, _OnExamRoomSelendokFnc);
	//m_wndExamRoom.SetEvent(WE_SETFOCUS, _OnExamRoomSetfocusFnc);
	//m_wndExamRoom.SetEvent(WE_KILLFOCUS, _OnExamRoomKillfocusFnc);
	m_wndExamRoom.SetEvent(WE_SELCHANGE, _OnExamRoomSelectChangeFnc);
	m_wndExamRoom.SetEvent(WE_LOADDATA, _OnExamRoomLoadDataFnc);
	//m_wndExamRoom.SetEvent(WE_ADDNEW, _OnExamRoomAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	m_wndSelectAllDiseases.SetEvent(WE_CLICK, _OnSelectAllDiseasesSelectFnc);
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");	
	SetMode(VM_EDIT);
}
void CEMStatisticsMajorDiseasesReport::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndExamRoom.GetDlgCtrlID(), m_szExamRoomKey);
	DDX_Check(pDX, m_wndSelectAllDiseases.GetDlgCtrlID(), m_bSelectAllDiseases);
}
void CEMStatisticsMajorDiseasesReport::GetDataToScreen()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMStatisticsMajorDiseasesReport::GetScreenToData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMStatisticsMajorDiseasesReport::SetDefaultValues()
{
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szExamRoomKey.Empty();
	m_bSelectAllDiseases=FALSE;
}
int CEMStatisticsMajorDiseasesReport::SetMode(int nMode)
{
 		int nOldMode = GetMode(); 
 		CGuiView::SetMode(nMode); 
 		CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 		CString szSQL, tmpStr; 
 		CRecord rs(&pMF->m_db);
		tmpStr = pMF->GetSysDate();
  		switch(nMode)
		{ 
 		case VM_ADD: 
 			EnableControls(TRUE); 
 			EnableButtons(TRUE, 0, 1, 2, -1);
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
/*void CEMStatisticsMajorDiseasesReport::OnYearChange(){
	
} */
/*void CEMStatisticsMajorDiseasesReport::OnYearSetfocus(){
	
} */
/*void CEMStatisticsMajorDiseasesReport::OnYearKillfocus(){
	
} */
int CEMStatisticsMajorDiseasesReport::OnYearCheckValue(){
	return 0;
} 
void CEMStatisticsMajorDiseasesReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMStatisticsMajorDiseasesReport::OnReportPeriodSelendok()
{
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
	if(nMonth == 13)
	{
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/03/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59" ), m_nYear,  dte.GetMonthLastDay());
	}
	if(nMonth == 14)
	{
		m_szFromDate.Format(_T("%.4d/04/01"), m_nYear);
		tmpStr.Format(_T("%.4d/06/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 15)
	{
		m_szFromDate.Format(_T("%.4d/07/01"), m_nYear);
		tmpStr.Format(_T("%.4d/09/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 16)
	{
		m_szFromDate.Format(_T("%.4d/10/01"), m_nYear);
		tmpStr.Format(_T("%.4d/10/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 17)
	{
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/12/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	UpdateData(false);
}
/*void CEMStatisticsMajorDiseasesReport::OnReportPeriodSetfocus(){
	
}*/
/*void CEMStatisticsMajorDiseasesReport::OnReportPeriodKillfocus(){
	
}*/
long CEMStatisticsMajorDiseasesReport::OnReportPeriodLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0)
	{
		szWhere.Format(_T(" WHERE hpr_idx=%ld "), ToInt(m_szReportPeriodKey));
	}
	m_wndReportPeriod.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT * FROM hms_period_report %s ORDER BY hpr_idx "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF())
	{ 
		m_wndReportPeriod.AddItems(
			rs.GetValue(_T("hpr_idx")), 
			rs.GetValue(_T("hpr_name")),			
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMStatisticsMajorDiseasesReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMStatisticsMajorDiseasesReport::OnFromDateChange(){
	
} */
/*void CEMStatisticsMajorDiseasesReport::OnFromDateSetfocus(){
	
} */
/*void CEMStatisticsMajorDiseasesReport::OnFromDateKillfocus(){
	
} */
int CEMStatisticsMajorDiseasesReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMStatisticsMajorDiseasesReport::OnToDateChange(){
	
} */
/*void CEMStatisticsMajorDiseasesReport::OnToDateSetfocus(){
	
} */
/*void CEMStatisticsMajorDiseasesReport::OnToDateKillfocus(){
	
} */
int CEMStatisticsMajorDiseasesReport::OnToDateCheckValue(){
	return 0;
} 
void CEMStatisticsMajorDiseasesReport::OnExamRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMStatisticsMajorDiseasesReport::OnExamRoomSelendok(){
	 
}
/*void CEMStatisticsMajorDiseasesReport::OnExamRoomSetfocus(){
	
}*/
/*void CEMStatisticsMajorDiseasesReport::OnExamRoomKillfocus(){
	
}*/
long CEMStatisticsMajorDiseasesReport::OnExamRoomLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();
	if(m_wndExamRoom.IsSearchKey() && !m_szExamRoomKey.IsEmpty())
	{
		szWhere.Format(_T("and hrl_id = %d"), ToInt(m_szExamRoomKey));
	}
	m_wndExamRoom.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("select hrl_name as name, hrl_id as id from hms_roomlist where hrl_deptid = '%s' %s Order by id"), pMF->m_szDept, szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndExamRoom.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMStatisticsMajorDiseasesReport::OnExamRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMStatisticsMajorDiseasesReport::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString szSQL, szTemp, tmpStr;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	if(!rpt.Init(_T("Reports/HMS/HE_THONGKEMATBENHCHINH.RPT")) )
		return ;

	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	int nIndex = 1;
	CReportSection* rptDetail;
	double  nTotal = 0;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%d"), nIndex ++);
		rptDetail->SetValue(_T("1"), tmpStr);

		szTemp = rs.GetValue(_T("hi_icd"));
		rptDetail->SetValue(_T("2"), szTemp);

		szTemp = rs.GetValue(_T("ICDName"));
		rptDetail->SetValue(_T("3"), szTemp);

		szTemp = rs.GetValue(_T("total"));
		nTotal += ToInt(szTemp);
		rptDetail->SetValue(_T("4"), szTemp);
		rs.MoveNext();
	}	
	szTemp.Format(_T("s%d"), nTotal);
	rpt.GetReportFooter()->SetValue(_T("s4"), nTotal);
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
	EndWaitCursor();
		
} 
void CEMStatisticsMajorDiseasesReport::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, szTemp, tmpStr;
	CExcel xls;
	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 78);
	xls.SetColumnWidth(3, 12);

	int nRow = 0, nCol = 0;
	xls.SetCellMergedColumns(nCol, nRow + 2, 4);

	xls.SetCellText(nCol+2, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol+2, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(nCol, nRow + 2, 4);
	xls.SetCellMergedColumns(nCol, nRow + 3, 4);
	xls.SetCellText(nCol, nRow + 2, _T("TH\x1ED0NG K\xCA M\x1EB6T \x42\x1EC6NH \x43H\xCDNH"), FMT_TEXT | FMT_CENTER, true, 13);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), 
		          CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss),
				  CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellText(nCol, nRow + 4, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol + 1, nRow + 4, _T("M\xE3 \x62\x1EC7nh"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol + 2, nRow + 4, _T("T\xEAn m\x1EB7t \x62\x1EC7nh"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol + 3, nRow + 4, _T("S\x1ED1 l\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER, true, 11);

	nRow += 5;
	int nIndex = 1;
	long double nTotal = 0.0;
	while (!rs.IsEOF())
	{
		szTemp.Format(_T("%d"), nIndex);
		xls.SetCellText(nCol, nRow, szTemp, FMT_INTEGER);

		szTemp = rs.GetValue(_T("hi_icd"));
		xls.SetCellText(nCol + 1, nRow, szTemp, FMT_TEXT);

		szTemp = rs.GetValue(_T("ICDName"));
		xls.SetCellText(nCol + 2, nRow, szTemp, FMT_TEXT);

		szTemp = rs.GetValue(_T("total"));
		nTotal += ToDouble(szTemp);
		xls.SetCellText(nCol + 3, nRow, szTemp, FMT_INTEGER);
		nIndex++;
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellText(nCol + 3, nRow, double2str(nTotal), FMT_NUMBER1, true, 12 );
	EndWaitCursor();
	xls.Save(_T("Exports\\THONG KE MAT BENH CHINH.XLS"));
} 
void CEMStatisticsMajorDiseasesReport::OnCloseSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CEMStatisticsMajorDiseasesReport::OnSelectAllDiseasesSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
int CEMStatisticsMajorDiseasesReport::OnAddEMStatisticsMajorDiseasesReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMStatisticsMajorDiseasesReport::OnEditEMStatisticsMajorDiseasesReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMStatisticsMajorDiseasesReport::OnDeleteEMStatisticsMajorDiseasesReport(){
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
 		OnCancelEMStatisticsMajorDiseasesReport(); 
 	}
	return 0;
}
int CEMStatisticsMajorDiseasesReport::OnSaveEMStatisticsMajorDiseasesReport(){
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
 		//OnEMStatisticsMajorDiseasesReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMStatisticsMajorDiseasesReport::OnCancelEMStatisticsMajorDiseasesReport(){
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
int CEMStatisticsMajorDiseasesReport::OnEMStatisticsMajorDiseasesReportListLoadData(){
	return 0;
}

CString CEMStatisticsMajorDiseasesReport::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere, szCondition;
	szWhere.Empty();

	if (!m_szExamRoomKey.IsEmpty())
	{
		szWhere.Format(_T(" and he_roomid = %d "), ToInt(m_szExamRoomKey));
	}
	if (m_bSelectAllDiseases)
		szCondition.Empty();
	//else
		//szCondition.Format(_T(" LIMIT 10 "));

	szSQL.Format(_T(" SELECT  hi_icd,") \
				_T(" 	hi_name as ICDName,") \
				_T(" 	count(hd_docno) as total") \
				_T("  FROM ") \
				_T("  (	  ") \
				_T("   SELECT ") \
				_T(" 	distinct hd_docno,") \
				_T(" 	max(he_examdate) as he_examedate,") \
				_T(" 	hd_icd") \
				_T("   FROM hms_doc") \
				_T("   LEFT JOIN hms_exam ON (he_docno = hd_docno)") \
				_T("   WHERE hd_status = 'T' AND length(hd_icd) > 0 AND (he_examdate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') AND hd_admitdept = '%s' %s") \
				_T("   GROUP BY hd_docno, hd_icd") \
				_T("  ) Tbl") \
				_T("  LEFT JOIN hms_icd ON(hi_icd = hd_icd)") \
				_T("  GROUP BY hi_icd, hi_name") \
				_T("  ORDER BY hi_icd, ICDName DESC %s "), m_szFromDate, m_szToDate, pMF->m_szDept, szWhere, szCondition);
	return szSQL;
}