#include "stdafx.h"
#include "EMAgeDiseaseStatisticsReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMAgeDiseaseStatisticsReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMAgeDiseaseStatisticsReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMAgeDiseaseStatisticsReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMAgeDiseaseStatisticsReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMAgeDiseaseStatisticsReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnExamRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMAgeDiseaseStatisticsReport* )pWnd)->OnExamRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnExamRoomSelendokFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnExamRoomSelendok();
}
/*static void _OnExamRoomSetfocusFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnExamRoomKillfocus();
}*/
/*static void _OnExamRoomKillfocusFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnExamRoomKillfocus();
}*/
static long _OnExamRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMAgeDiseaseStatisticsReport *)pWnd)->OnExamRoomLoadData();
}
/*static void _OnExamRoomAddNewFnc(CWnd *pWnd){
	((CEMAgeDiseaseStatisticsReport *)pWnd)->OnExamRoomAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMAgeDiseaseStatisticsReport *pVw = (CEMAgeDiseaseStatisticsReport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMAgeDiseaseStatisticsReport *pVw = (CEMAgeDiseaseStatisticsReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMAgeDiseaseStatisticsReport *pVw = (CEMAgeDiseaseStatisticsReport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddEMAgeDiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMAgeDiseaseStatisticsReport*)pWnd)->OnAddEMAgeDiseaseStatisticsReport();
} 
static int _OnEditEMAgeDiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMAgeDiseaseStatisticsReport*)pWnd)->OnEditEMAgeDiseaseStatisticsReport();
} 
static int _OnDeleteEMAgeDiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMAgeDiseaseStatisticsReport*)pWnd)->OnDeleteEMAgeDiseaseStatisticsReport();
} 
static int _OnSaveEMAgeDiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMAgeDiseaseStatisticsReport*)pWnd)->OnSaveEMAgeDiseaseStatisticsReport();
} 
static int _OnCancelEMAgeDiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMAgeDiseaseStatisticsReport*)pWnd)->OnCancelEMAgeDiseaseStatisticsReport();
} 
CEMAgeDiseaseStatisticsReport::CEMAgeDiseaseStatisticsReport(CWnd *pParent)
{
	m_nDlgWidth = 440;
	m_nDlgHeight = 155;
	m_szTitle = _T("Statistic ICD - Age");
	SetDefaultValues();
}
CEMAgeDiseaseStatisticsReport::~CEMAgeDiseaseStatisticsReport(){

}
void CEMAgeDiseaseStatisticsReport::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 120);
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
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 240, 125, 350, 150);
	m_wndExport.Create(this, _T("&Export"), 355, 125, 430, 150);
}
void CEMAgeDiseaseStatisticsReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	////m_wndYear.SetCheckValue(true);
	////m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	//m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	//m_wndToDate.SetCheckValue(true);
	//m_wndExamRoom.SetCheckValue(true);
	m_wndExamRoom.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

	m_wndExamRoom.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndExamRoom.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);
}
void CEMAgeDiseaseStatisticsReport::OnSetWindowEvents(){
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
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	SetMode(VM_EDIT);
}
void CEMAgeDiseaseStatisticsReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndExamRoom.GetDlgCtrlID(), m_szExamRoomKey);
}
void CEMAgeDiseaseStatisticsReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMAgeDiseaseStatisticsReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CEMAgeDiseaseStatisticsReport::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szExamRoomKey.Empty();

}
int CEMAgeDiseaseStatisticsReport::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0,1,2,3,-1); 
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
/*void CEMAgeDiseaseStatisticsReport::OnYearChange(){
	
} */
/*void CEMAgeDiseaseStatisticsReport::OnYearSetfocus(){
	
} */
/*void CEMAgeDiseaseStatisticsReport::OnYearKillfocus(){
	
} */
int CEMAgeDiseaseStatisticsReport::OnYearCheckValue(){
	return 0;
} 
void CEMAgeDiseaseStatisticsReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMAgeDiseaseStatisticsReport::OnReportPeriodSelendok(){
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
/*void CEMAgeDiseaseStatisticsReport::OnReportPeriodSetfocus(){
	
}*/
/*void CEMAgeDiseaseStatisticsReport::OnReportPeriodKillfocus(){
	
}*/
long CEMAgeDiseaseStatisticsReport::OnReportPeriodLoadData(){
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
/*void CEMAgeDiseaseStatisticsReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMAgeDiseaseStatisticsReport::OnFromDateChange(){
	
} */
/*void CEMAgeDiseaseStatisticsReport::OnFromDateSetfocus(){
	
} */
/*void CEMAgeDiseaseStatisticsReport::OnFromDateKillfocus(){
	
} */
int CEMAgeDiseaseStatisticsReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMAgeDiseaseStatisticsReport::OnToDateChange(){
	
} */
/*void CEMAgeDiseaseStatisticsReport::OnToDateSetfocus(){
	
} */
/*void CEMAgeDiseaseStatisticsReport::OnToDateKillfocus(){
	
} */
int CEMAgeDiseaseStatisticsReport::OnToDateCheckValue(){
	return 0;
} 
void CEMAgeDiseaseStatisticsReport::OnExamRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMAgeDiseaseStatisticsReport::OnExamRoomSelendok(){
	 
}
/*void CEMAgeDiseaseStatisticsReport::OnExamRoomSetfocus(){
	
}*/
/*void CEMAgeDiseaseStatisticsReport::OnExamRoomKillfocus(){
	
}*/
long CEMAgeDiseaseStatisticsReport::OnExamRoomLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndExamRoom.IsSearchKey() && !m_szExamRoomKey.IsEmpty()){
		szWhere.Format(_T("and hrl_id = %d"), ToInt(m_szExamRoomKey));
	};
	m_wndExamRoom.DeleteAllItems(); 
	szSQL.Format(_T("select hrl_name as name, hrl_id as id from hms_roomlist where hrl_deptid = '%s' %s Order by id"), pMF->m_szDept, szWhere);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndExamRoom.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMAgeDiseaseStatisticsReport::OnExamRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMAgeDiseaseStatisticsReport::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);

	if(!rpt.Init(_T("Reports/HMS/HE_THONGKEMATBENHNHOMTUOI.RPT")) )
	return ;
	UpdateData(true);
	BeginWaitCursor();
	szSQL = GetQueryString();
	
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	rs.GetValue(_T("tenkhoa"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("Department"), tmpStr);

	rs.GetValue(_T("tenkhoa"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);

	int nIndex = 1;
	int nTotal[5];
	for(int i = 1; i<=4; i++)
	{
		 nTotal[i] = 0;
	}
		CReportSection* rptDetail;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%ld"), nIndex ++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("ICD"),tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("ICDName"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("TuoiNhohon50"), tmpStr);
		nTotal[1] += ToInt(tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		
		rs.GetValue(_T("Tuoi5060"), tmpStr);
		nTotal[2] += ToInt(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		
		rs.GetValue(_T("Tuoi6070"), tmpStr);
		nTotal[3] += ToInt(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("Tuoilonhon70"), tmpStr);
		nTotal[4] += ToInt(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.MoveNext();	
	}	
		for (int i = 1; i <= 4; i++)
		{

			CString  szField;
			szField.Format(_T("s%d"), i+3);
			tmpStr.Format(_T("%d"), nTotal[i]);
			rpt.GetReportFooter()->SetValue(szField, tmpStr);		
		}
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();
} 
void CEMAgeDiseaseStatisticsReport::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szSQL1, szWhere;
	CRecord rs(&pMF->m_db);
	CRecord rs0(&pMF->m_db);
	UpdateData(true);
	BeginWaitCursor();
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 58);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 13);
	xls.SetColumnWidth(5, 13);
	xls.SetColumnWidth(6, 10);
	int nCol = 0;
	int nRow = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);

	xls.SetCellMergedColumns(nCol, nRow , 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	rs.GetValue(_T("tenkhoa"), tmpStr);
	xls.SetCellText(nCol, nRow + 2, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 7);
	xls.SetCellText(nCol, nRow + 3, _T("TH\x1ED0NG K\xCA M\x1EB6T \x42\x1EC6NH – NH\xD3M TU\x1ED4I"), FMT_TEXT | FMT_CENTER, true, 14);
	xls.SetCellMergedColumns(nCol, nRow + 4, 7);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol, nRow + 5, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 5, _T("M\xE3 I\x43\x44"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 5, _T("M\x1EB7t \x62\x1EC7nh"), FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellText(nCol + 3, nRow + 5, _T("< \x35\x30 tu\x1ED5i"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 5, _T("\x35\x30 – \x36\x30 tu\x1ED5i"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 5, nRow + 5, _T("\x36\x30 – \x37\x30 tu\x1ED5i"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 6, nRow + 5, _T("> \x37\x30 tu\x1ED5i"), FMT_TEXT | FMT_CENTER, true, 11);
	int nIndex = 1;
	int nTotal[5];
	for (int i = 1; i<= 4; i++)
	{
		nTotal[i] = 0;
	}
	while(!rs.IsEOF())
	{
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow + 6, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("ICD"),tmpStr);
		xls.SetCellText(nCol + 1, nRow + 6, tmpStr, FMT_TEXT);

		rs.GetValue(_T("ICDName"),tmpStr);
		xls.SetCellText(nCol + 2, nRow + 6, tmpStr, FMT_TEXT);

		rs.GetValue(_T("TuoiNhohon50"), tmpStr);
		nTotal[1] += ToInt(tmpStr);
		xls.SetCellText(nCol + 3, nRow + 6, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("Tuoi5060"),tmpStr);
		nTotal[2] += ToInt(tmpStr);
		xls.SetCellText(nCol + 4, nRow + 6, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("Tuoi6070"), tmpStr);
		nTotal[3] += ToInt(tmpStr);
		xls.SetCellText(nCol + 5, nRow + 6, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("Tuoilonhon70"), tmpStr);
		nTotal[4] += ToInt(tmpStr);
		xls.SetCellText(nCol + 6, nRow + 6, tmpStr, FMT_NUMBER1);
		nRow++;
		rs.MoveNext();		
	}
	xls.SetCellMergedColumns(nCol, nRow + 6, 3);
	xls.SetCellText(nCol, nRow + 6, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 11);
	 for( int i = 1; i<= 4; i++)
	 {
		 tmpStr.Format(_T("%d"), nTotal[i]);
		 xls.SetCellText(i+2, nRow + 6, tmpStr, FMT_NUMBER1, true, 11);
	 }
	CString szDate;
	EndWaitCursor();
	xls.Save(_T("Exports\\THONG KE MAT BENH - NHOM TUOI.xls"));
	
} 
void CEMAgeDiseaseStatisticsReport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CEMAgeDiseaseStatisticsReport::OnAddEMAgeDiseaseStatisticsReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMAgeDiseaseStatisticsReport::OnEditEMAgeDiseaseStatisticsReport()
{
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMAgeDiseaseStatisticsReport::OnDeleteEMAgeDiseaseStatisticsReport(){
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
 		OnCancelEMAgeDiseaseStatisticsReport(); 
 	}
	return 0;
}
int CEMAgeDiseaseStatisticsReport::OnSaveEMAgeDiseaseStatisticsReport(){
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
 		//OnEMAgeDiseaseStatisticsReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMAgeDiseaseStatisticsReport::OnCancelEMAgeDiseaseStatisticsReport(){
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
int CEMAgeDiseaseStatisticsReport::OnEMAgeDiseaseStatisticsReportListLoadData(){
	return 0;
}
CString CEMAgeDiseaseStatisticsReport:: GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	szWhere.Empty();
	if (!m_szExamRoomKey.IsEmpty())
	{
		szWhere.Format(_T(" and he_roomid = %d "), ToInt(m_szExamRoomKey));
	}
	szSQL.Format(_T(" SELECT  ") \
				_T("       ICD,") \
				_T("       ICDName, ") \
				_T("       sum(TuoiNhohon50) as TuoiNhohon50,") \
				_T("       sum(Tuoi5060) as Tuoi5060,") \
				_T("       sum(Tuoi6070) as Tuoi6070,") \
				_T("       sum(Tuoilonhon70) as Tuoilonhon70 	") \
				_T(" FROM ") \
				_T(" (") \
				_T("  SELECT		") \
				_T("     icd as ICD,") \
				_T("     hi_name AS ICDName,") \
				_T("     docno,") \
				_T("     case when ageno < 50 then 1 else 0 end as TuoiNhohon50,") \
				_T("     case when ageno >= 50 and ageno <= 60 then 1 else 0 end as Tuoi5060,") \
				_T("     case when ageno >60 and ageno <=70 then 1 else 0 end as Tuoi6070,") \
				_T("     case when ageno > 70 then 1 else 0 end as Tuoilonhon70	") \
				_T("  FROM    (") \
				_T("   SELECT  docno, icd, age, ageno, MAX(he_examdate) AS endexamdate ") \
				_T("   FROM(") \
				_T("     SELECT  hms_getage(hd_admitdate, hp_birthdate) as age,") \
				_T("             hd_docno as docno,") \
				_T("			 hd_admitdept,") \
				_T("             floor(GetDaysBetween(hd_admitdate, hp_birthdate)/365) as ageno,") \
				_T("             hd_icd as icd") \
				_T("     FROM hms_patient") \
				_T("     LEFT JOIN hms_doc ON hp_patientno = hd_patientno) tbl") \
				_T("   LEFT JOIN hms_exam ON he_docno = docno") \
				_T("   WHERE he_status IN ('T', 'P')  AND LENGTH(icd) > 0 AND hd_admitdept = '%s' %s") \
				_T("   GROUP BY docno, icd, age, ageno") \
				_T("      ) tbl1    ") \
				_T("  LEFT JOIN HMS_ICD ON(HI_ICD = icd)") \
				_T("  WHERE length(icd) > 0 AND (endexamdate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
				_T("  ) tbl2") \
				_T(" GROUP BY ICD, ICDNAME") \
				_T(" ORDER BY ICD, ICDNAME"), pMF->m_szDept, szWhere, m_szFromDate, m_szToDate);
	return szSQL;
}