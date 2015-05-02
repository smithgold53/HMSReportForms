#include "stdafx.h"
#include "EMObjectDiseaseStatisticsReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMObjectDiseaseStatisticsReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){ 
	((CEMObjectDiseaseStatisticsReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMObjectDiseaseStatisticsReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMObjectDiseaseStatisticsReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMObjectDiseaseStatisticsReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnExamRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMObjectDiseaseStatisticsReport* )pWnd)->OnExamRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnExamRoomSelendokFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnExamRoomSelendok();
}
/*static void _OnExamRoomSetfocusFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnExamRoomKillfocus();
}*/
/*static void _OnExamRoomKillfocusFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnExamRoomKillfocus();
}*/
static long _OnExamRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMObjectDiseaseStatisticsReport *)pWnd)->OnExamRoomLoadData();
}
/*static void _OnExamRoomAddNewFnc(CWnd *pWnd){
	((CEMObjectDiseaseStatisticsReport *)pWnd)->OnExamRoomAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMObjectDiseaseStatisticsReport *pVw = (CEMObjectDiseaseStatisticsReport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMObjectDiseaseStatisticsReport *pVw = (CEMObjectDiseaseStatisticsReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMObjectDiseaseStatisticsReport *pVw = (CEMObjectDiseaseStatisticsReport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddEMObjectDiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMObjectDiseaseStatisticsReport*)pWnd)->OnAddEMObjectDiseaseStatisticsReport();
} 
static int _OnEditEMObjectDiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMObjectDiseaseStatisticsReport*)pWnd)->OnEditEMObjectDiseaseStatisticsReport();
} 
static int _OnDeleteEMObjectDiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMObjectDiseaseStatisticsReport*)pWnd)->OnDeleteEMObjectDiseaseStatisticsReport();
} 
static int _OnSaveEMObjectDiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMObjectDiseaseStatisticsReport*)pWnd)->OnSaveEMObjectDiseaseStatisticsReport();
} 
static int _OnCancelEMObjectDiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMObjectDiseaseStatisticsReport*)pWnd)->OnCancelEMObjectDiseaseStatisticsReport();
} 
CEMObjectDiseaseStatisticsReport::CEMObjectDiseaseStatisticsReport(CWnd *pParent)
{
	m_nDlgWidth = 445;
	m_nDlgHeight = 155;
	m_szTitle = _T("Statistic ICD - Object");
	SetDefaultValues();
}
CEMObjectDiseaseStatisticsReport::~CEMObjectDiseaseStatisticsReport(){
}
void CEMObjectDiseaseStatisticsReport::OnCreateComponents(){
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
void CEMObjectDiseaseStatisticsReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	//m_wndExamRoom.SetCheckValue(true);
	m_wndExamRoom.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);


	m_wndExamRoom.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndExamRoom.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

}
void CEMObjectDiseaseStatisticsReport::OnSetWindowEvents(){
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
void CEMObjectDiseaseStatisticsReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndExamRoom.GetDlgCtrlID(), m_szExamRoomKey);
}
void CEMObjectDiseaseStatisticsReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMObjectDiseaseStatisticsReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMObjectDiseaseStatisticsReport::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szExamRoomKey.Empty();

}
int CEMObjectDiseaseStatisticsReport::SetMode(int nMode){
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
/*void CEMObjectDiseaseStatisticsReport::OnYearChange(){
	
} */
/*void CEMObjectDiseaseStatisticsReport::OnYearSetfocus(){
	
} */
/*void CEMObjectDiseaseStatisticsReport::OnYearKillfocus(){
	
} */
int CEMObjectDiseaseStatisticsReport::OnYearCheckValue(){
	return 0;
} 
void CEMObjectDiseaseStatisticsReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMObjectDiseaseStatisticsReport::OnReportPeriodSelendok(){
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
/*void CEMObjectDiseaseStatisticsReport::OnReportPeriodSetfocus(){
	
}*/
/*void CEMObjectDiseaseStatisticsReport::OnReportPeriodKillfocus(){
	
}*/
long CEMObjectDiseaseStatisticsReport::OnReportPeriodLoadData(){
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
/*void CEMObjectDiseaseStatisticsReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMObjectDiseaseStatisticsReport::OnFromDateChange(){
	
} */
/*void CEMObjectDiseaseStatisticsReport::OnFromDateSetfocus(){
	
} */
/*void CEMObjectDiseaseStatisticsReport::OnFromDateKillfocus(){
	
} */
int CEMObjectDiseaseStatisticsReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMObjectDiseaseStatisticsReport::OnToDateChange(){
	
} */
/*void CEMObjectDiseaseStatisticsReport::OnToDateSetfocus(){
	
} */
/*void CEMObjectDiseaseStatisticsReport::OnToDateKillfocus(){
	
} */
int CEMObjectDiseaseStatisticsReport::OnToDateCheckValue(){
	return 0;
} 
void CEMObjectDiseaseStatisticsReport::OnExamRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMObjectDiseaseStatisticsReport::OnExamRoomSelendok(){
	 
}
/*void CEMObjectDiseaseStatisticsReport::OnExamRoomSetfocus(){
	
}*/
/*void CEMObjectDiseaseStatisticsReport::OnExamRoomKillfocus(){
	
}*/
long CEMObjectDiseaseStatisticsReport::OnExamRoomLoadData(){
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
/*void CEMObjectDiseaseStatisticsReport::OnExamRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMObjectDiseaseStatisticsReport::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);

	if(!rpt.Init(_T("Reports/HMS/HE_THONGKEMATBENHTHEODOITUONGC12.RPT")) )
		return;
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

	rs.GetValue(_T("tenkhoa"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);

	int nIndex = 1;
	long nTotal[11];
	for(int i = 1; i<= 10; i++)
	{
		 nTotal[i] = 0;
	}
	CReportSection* rptDetail = NULL;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%ld"), nIndex ++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("ICD"),tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("ICDName"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("c1"), tmpStr);
		nTotal[1] += ToLong(tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		
		rs.GetValue(_T("c2"), tmpStr);
		nTotal[2] += ToLong(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		
		rs.GetValue(_T("c3"), tmpStr);
		nTotal[3] += ToLong(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("c4"), tmpStr);
		nTotal[4] += ToLong(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("c5"), tmpStr);
		nTotal[5] += ToLong(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("c6"), tmpStr);
		nTotal[6] += ToLong(tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("c7"), tmpStr);
		nTotal[7] += ToLong(tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.GetValue(_T("c8"), tmpStr);
		nTotal[8] += ToLong(tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.GetValue(_T("c9"), tmpStr);
		nTotal[9] += ToLong(tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);

		rs.GetValue(_T("c10"), tmpStr);
		nTotal[10] += ToLong(tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);

		rs.MoveNext();
			
	}	
	for (int i = 1; i <= 10; i++)
	{
		CString  szField;
		szField.Format(_T("s%d"), i+3);
		tmpStr.Format(_T("%ld"), nTotal[i]);
		rpt.GetReportFooter()->SetValue(szField, tmpStr);	
	}
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();
	
} 
void CEMObjectDiseaseStatisticsReport::OnExportSelect(){
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
	xls.SetColumnWidth(2, 54);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(5, 10);
	xls.SetColumnWidth(6, 10);
	xls.SetColumnWidth(7, 10);
	xls.SetColumnWidth(8, 10);
	xls.SetColumnWidth(9, 10);
	xls.SetColumnWidth(10, 10);
	xls.SetColumnWidth(11, 10);
	xls.SetColumnWidth(12, 13);
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
	xls.SetCellText(nCol, nRow + 2, tmpStr, FMT_TEXT, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 13);
	xls.SetCellText(nCol, nRow + 3, _T("TH\x1ED0NG K\xCA M\x1EB6T \x42\x1EC6NH TH\x45O \x110\x1ED0I T\x1AF\x1EE2NG"), FMT_TEXT | FMT_CENTER, true, 14);
	xls.SetCellMergedColumns(nCol, nRow + 4, 13);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol, nRow + 5, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 5, _T("ICD10"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 5, _T("M\x1EB7t \x62\x1EC7nh"), FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellText(nCol + 3, nRow + 5, _T("Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 5, _T("\x42HYTQ"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 5, nRow + 5, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 11);
	int nIndex = 1;
	long nTotal[4];
	for (int i = 1; i<=4; i++)
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

		rs.GetValue(_T("c1"), tmpStr);
		nTotal[1] += ToLong(tmpStr);
		xls.SetCellText(nCol + 3, nRow + 6, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("c2"),tmpStr);
		nTotal[2] += ToLong(tmpStr);
		xls.SetCellText(nCol + 4, nRow + 6, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("c10"), tmpStr);
		nTotal[3] += ToLong(tmpStr);
		xls.SetCellText(nCol + 5, nRow + 6, tmpStr, FMT_NUMBER1);

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
	xls.Save(_T("Exports\\THONG KE MAT BENH - DOI TUONG TAI C1-2.xls"));
	
} 
void CEMObjectDiseaseStatisticsReport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
	
} 
int CEMObjectDiseaseStatisticsReport::OnAddEMObjectDiseaseStatisticsReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMObjectDiseaseStatisticsReport::OnEditEMObjectDiseaseStatisticsReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMObjectDiseaseStatisticsReport::OnDeleteEMObjectDiseaseStatisticsReport(){
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
 		OnCancelEMObjectDiseaseStatisticsReport(); 
 	}
	return 0;
}
int CEMObjectDiseaseStatisticsReport::OnSaveEMObjectDiseaseStatisticsReport(){
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
 		//OnEMObjectDiseaseStatisticsReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMObjectDiseaseStatisticsReport::OnCancelEMObjectDiseaseStatisticsReport(){
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
int CEMObjectDiseaseStatisticsReport::OnEMObjectDiseaseStatisticsReportListLoadData(){
	return 0;
}
CString CEMObjectDiseaseStatisticsReport::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	if (!m_szExamRoomKey.IsEmpty())
	{
		szWhere.Format(_T(" and he_roomid = %d "), ToInt(m_szExamRoomKey));
	}
	szSQL.Format(_T(" SELECT  ") \
				_T(" 	ICD,") \
				_T(" 	ICDName, ") \
				_T(" 	sum(c1) as c1,") \
				_T(" 	sum(c2) as c2,") \
				_T(" 	sum(c3) as c3,") \
				_T(" 	sum(c4) as c4,") \
				_T(" 	sum(c5) as c5,") \
				_T(" 	sum(c6) as c6,") \
				_T(" 	sum(c7) as c7,") \
				_T(" 	sum(c8) as c8,") \
				_T(" 	sum(c9) as c9,") \
				_T(" 	sum(c10) as c10") \
				_T(" FROM ") \
				_T(" (") \
				_T(" SELECT		") \
				_T(" 		hi_icd as ICD,") \
				_T(" 		hi_name as ICDName,") \
				_T(" 		case when hd_object = '1' then 1 else 0 end as c1, ") \
				_T(" 		case when hd_object = '2' then 1 else 0 end as c2, ") \
				_T(" 		case when hd_object = '5' then 1 else 0 end as c3, ") \
				_T(" 		case when hd_object = '10' then 1 else 0 end as c4, ") \
				_T(" 		case when hd_object = '4' then 1 else 0 end as c5, ") \
				_T(" 		case when hd_object = '3' then 1 else 0 end as c6, ") \
				_T(" 		case when hd_object in('6', '9') then 1 else 0 end as c7, ") \
				_T(" 		case when hd_object = '8' then 1 else 0 end as c8, ") \
				_T(" 		case when hd_object = '7' then 1 else 0 end as c9, ") \
				_T(" 		1 as c10 ") \
				_T("  FROM    (") \
				_T(" 	SELECT") \
				_T(" 	distinct hd_docno, hd_object, ") \
				_T(" 	max(he_examdate) as endexamdate,") \
				_T(" 	hd_icd") \
				_T(" 	FROM hms_doc") \
				_T(" 	LEFT JOIN hms_exam ON he_patientno = hd_patientno and he_docno = hd_docno") \
				_T(" 	WHERE hd_status = 'T' AND length(hd_icd) > 0 %s AND hd_admitdept = '%s'") \
				_T(" 	GROUP BY hd_docno, hd_icd, hd_object") \
				_T("      ) tbl    ") \
				_T("  LEFT JOIN hms_icd ON(hi_icd = hd_icd)") \
				_T("  WHERE length(hd_icd) > 0 AND (endexamdate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
				_T("  GROUP BY hi_icd, hd_object, hd_docno,  hi_name") \
				_T("  ) tbl1") \
				_T("  GROUP BY ICD, ICDName") \
				_T("  ORDER BY ICD, ICDName"),szWhere, pMF->m_szDept, m_szFromDate, m_szToDate );
	return szSQL;
}