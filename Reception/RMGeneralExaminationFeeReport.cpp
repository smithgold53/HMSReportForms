#include "stdafx.h"
#include "RMGeneralExaminationFeeReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CRMGeneralExaminationFeeReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRMGeneralExaminationFeeReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CRMGeneralExaminationFeeReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CRMGeneralExaminationFeeReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CRMGeneralExaminationFeeReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRMGeneralExaminationFeeReport* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CRMGeneralExaminationFeeReport *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CRMGeneralExaminationFeeReport *)pWnd)->OnObjectAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CRMGeneralExaminationFeeReport *pVw = (CRMGeneralExaminationFeeReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CRMGeneralExaminationFeeReport *pVw = (CRMGeneralExaminationFeeReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CRMGeneralExaminationFeeReport *pVw = (CRMGeneralExaminationFeeReport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddRMGeneralExaminationFeeReportFnc(CWnd *pWnd){
	 return ((CRMGeneralExaminationFeeReport*)pWnd)->OnAddRMGeneralExaminationFeeReport();
} 
static int _OnEditRMGeneralExaminationFeeReportFnc(CWnd *pWnd){
	 return ((CRMGeneralExaminationFeeReport*)pWnd)->OnEditRMGeneralExaminationFeeReport();
} 
static int _OnDeleteRMGeneralExaminationFeeReportFnc(CWnd *pWnd){
	 return ((CRMGeneralExaminationFeeReport*)pWnd)->OnDeleteRMGeneralExaminationFeeReport();
} 
static int _OnSaveRMGeneralExaminationFeeReportFnc(CWnd *pWnd){
	 return ((CRMGeneralExaminationFeeReport*)pWnd)->OnSaveRMGeneralExaminationFeeReport();
} 
static int _OnCancelRMGeneralExaminationFeeReportFnc(CWnd *pWnd){
	 return ((CRMGeneralExaminationFeeReport*)pWnd)->OnCancelRMGeneralExaminationFeeReport();
} 
CRMGeneralExaminationFeeReport::CRMGeneralExaminationFeeReport(CWnd *pParent){

	m_nDlgWidth = 420;
	m_nDlgHeight = 155;
	m_szTitle = _T("Examination Fee Report");
	SetDefaultValues();
}
CRMGeneralExaminationFeeReport::~CRMGeneralExaminationFeeReport(){
}
void CRMGeneralExaminationFeeReport::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 490, 120);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 245, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 250, 30, 330, 55);
	m_wndReportPeriod.Create(this,335, 30, 485, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 245, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 60, 330, 85);
	m_wndToDate.Create(this,335, 60, 485, 85); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 90, 90, 115);
	m_wndObject.Create(this,95, 90, 485, 115); 
	m_wndPrint.Create(this, _T("&Print"), 325, 125, 405, 150);
	m_wndExport.Create(this, _T("&Export"), 410, 125, 490, 150);
}
void CRMGeneralExaminationFeeReport::OnInitializeComponents(){
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

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 80);
	m_wndObject.InsertColumn(1, _T("Description"), CFMT_TEXT, 210);
}
void CRMGeneralExaminationFeeReport::OnSetWindowEvents(){
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
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T(" 00:00");
	m_szToDate += _T(" 23:59");
	UpdateData(false);
}
void CRMGeneralExaminationFeeReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CRMGeneralExaminationFeeReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CRMGeneralExaminationFeeReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CRMGeneralExaminationFeeReport::SetDefaultValues(){
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szObjectKey.Empty();
}
int CRMGeneralExaminationFeeReport::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0,1,2,3, 4, -1); 
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
/*void CRMGeneralExaminationFeeReport::OnYearChange(){
	
} */
/*void CRMGeneralExaminationFeeReport::OnYearSetfocus(){
	
} */
/*void CRMGeneralExaminationFeeReport::OnYearKillfocus(){
	
} */
int CRMGeneralExaminationFeeReport::OnYearCheckValue(){
	return 0;
} 
void CRMGeneralExaminationFeeReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CRMGeneralExaminationFeeReport::OnReportPeriodSelendok()
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
/*void CRMGeneralExaminationFeeReport::OnReportPeriodSetfocus(){
	
}*/
/*void CRMGeneralExaminationFeeReport::OnReportPeriodKillfocus(){
	
}*/
long CRMGeneralExaminationFeeReport::OnReportPeriodLoadData(){
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
/*void CRMGeneralExaminationFeeReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CRMGeneralExaminationFeeReport::OnFromDateChange(){
	
} */
/*void CRMGeneralExaminationFeeReport::OnFromDateSetfocus(){
	
} */
/*void CRMGeneralExaminationFeeReport::OnFromDateKillfocus(){
	
} */
int CRMGeneralExaminationFeeReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CRMGeneralExaminationFeeReport::OnToDateChange(){
	
} */
/*void CRMGeneralExaminationFeeReport::OnToDateSetfocus(){
	
} */
/*void CRMGeneralExaminationFeeReport::OnToDateKillfocus(){
	
} */
int CRMGeneralExaminationFeeReport::OnToDateCheckValue(){
	return 0;
} 
void CRMGeneralExaminationFeeReport::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CRMGeneralExaminationFeeReport::OnObjectSelendok(){
	 
}
/*void CRMGeneralExaminationFeeReport::OnObjectSetfocus(){
	
}*/
/*void CRMGeneralExaminationFeeReport::OnObjectKillfocus(){
	
}*/
long CRMGeneralExaminationFeeReport::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty()){
		szWhere.Format(_T(" where ho_id='%s'"), m_szObjectKey);
	};
	m_wndObject.DeleteAllItems(); 
	szSQL.Format(_T(" select ho_id as id, ho_desc as name from hms_object %s order by cast(ho_id as integer)"), szWhere);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CrptThongkeTienPhiKham::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CRMGeneralExaminationFeeReport::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, tmpStr, szFromDateLabel, szToDateLabel, szWhere;
	CString szDate, szSysDate;
	BeginWaitCursor();
	szSysDate = pMF->GetSysDate(); 
	CRecord rs(&pMF->m_db);
	CReport rpt;
	UpdateData(true);
	if(!rpt.Init(_T("Reports/HMS/RM_BAOCAOTHUTIENPHIKHAM.RPT")) )
	return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return ;
	}
	_fmsg(_T("%s"), szSQL);
	
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	CReportSection* rptDetail;
	int nIndex = 1;
	int nPatient = 0;
	int grpPatient = 0;
	long double grpCost = 0;
	double Cost = 0;

	CString szNewRoomName, szOldRoomName;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();	
		rs.GetValue(_T("roomname"), szNewRoomName);
		if(szOldRoomName != szNewRoomName)
		{
			tmpStr.Format(_T("%d"), nIndex++);
			rptDetail->SetValue(_T("1"), tmpStr);
			rptDetail->SetValue(_T("2"), szNewRoomName);
			szOldRoomName = szNewRoomName;
		}
		else
		{
			rptDetail->SetValue(_T("1"), _T(""));
			rptDetail->SetValue(_T("2"), _T("-"));
		}
		rs.GetValue(_T("examtype"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("totalexam"), nPatient);
		grpPatient += nPatient;
		FormatCurrency(nPatient,tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("money"), Cost);
		grpCost += Cost;
		FormatCurrency(Cost,tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.MoveNext();
	}
			tmpStr.Format(_T("%d"), grpPatient);
			FormatCurrency(grpPatient,tmpStr);
			rpt.GetReportFooter()->SetValue(_T("s4"), tmpStr);		

			tmpStr.Format(_T("%.0f"), grpCost);
			FormatCurrency(grpCost,tmpStr);
			rpt.GetReportFooter()->SetValue(_T("s5"), tmpStr);		

	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();
	
} 
void CRMGeneralExaminationFeeReport::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CString szDate, szSysDate;
	szSysDate = pMF->GetSysDate(); 
	CRecord rs(&pMF->m_db);
	UpdateData(true);
	BeginWaitCursor();
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 28);
	xls.SetColumnWidth(2, 42);
	xls.SetColumnWidth(3, 13);
	xls.SetColumnWidth(4, 14);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedColumns(nCol, nRow + 3, 5);
	xls.SetCellText(nCol, nRow + 3, _T("\x42\xC1O \x43\xC1O THU TI\x1EC0N PH\xCD KH\xC1M"), FMT_TEXT | FMT_CENTER, true, 14);
	xls.SetCellMergedColumns(nCol, nRow + 4, 5);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12);

	xls.SetCellText(nCol, nRow + 5, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 5, _T("Ph\xF2ng kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 5, _T("Lo\x1EA1i kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellText(nCol + 3, nRow + 5, _T("T\x1ED5ng s\x1ED1 \x42N"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 5, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 11);
		
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
		// _T(" 	AND he_deptid='%s' ") \
		//_T(" ORDER BY roomid, examtype, roomname") , pMF->m_szDept, m_szFromDate, m_szToDate);
	if(rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return ;
	}
	_fmsg(_T("%s"), szSQL);
	int nIndex = 1;
	int nPatient = 0;

	long double nAmount = 0.0;
	CString szOldRoomName, szNewRoomName;
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("roomname"), szNewRoomName);
		if(szOldRoomName != szNewRoomName)
		{
			tmpStr.Format(_T("%d"), nIndex++);
			xls.SetCellText(nCol, nRow + 6, tmpStr, FMT_INTEGER);
			xls.SetCellText(nCol + 1, nRow + 6, szNewRoomName, FMT_TEXT);
			szOldRoomName = szNewRoomName;
		}
		else
		{
			xls.SetCellText(nCol, nRow + 6, _T(""), FMT_TEXT);
			xls.SetCellText(nCol + 1, nRow + 6, _T("-"), FMT_TEXT);
		} 
		rs.GetValue(_T("examtype"),tmpStr);
		xls.SetCellText(nCol + 2, nRow + 6, tmpStr, FMT_TEXT);

		rs.GetValue(_T("totalexam"),tmpStr);
		nPatient += ToInt(tmpStr);
		xls.SetCellText(nCol + 3, nRow + 6, tmpStr, FMT_INTEGER);
		
		rs.GetValue(_T("money"),tmpStr);
		nAmount += ToDouble(tmpStr);
		xls.SetCellText(nCol + 4, nRow + 6, tmpStr, FMT_NUMBER1);
		rs.MoveNext();
		nRow++;		
	}
	xls.SetCellMergedColumns(nCol, nRow + 6, 3);
	xls.SetCellText(nCol, nRow + 6, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellText(nCol + 3, nRow + 6, int2str(nPatient), FMT_INTEGER, true, 12);
	xls.SetCellText(nCol + 4, nRow + 6, double2str(nAmount), FMT_NUMBER1, true, 12);
	/*nRow++;
	rs.GetValue(_T("szSysDate"), szDate);
	szDate.Format(rs.GetValue(_T("szSysDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	xls.SetCellText(nCol + 3, nRow + 6, szDate, FMT_TEXT | FMT_CENTER, true, 12);
	nRow++;

	xls.SetCellText(nCol + 3, nRow + 6, _T("Ng\x1B0\x1EDDi l\x1EADp \x62i\x1EC3u"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellMergedRows(nCol + 3, nRow + 6, 3);
	xls.SetCellText(nCol + 3, nRow + 6, _T("H\x1ECD v\xE0 t\xEAn:\x2026\x2026\x2026\x2026\x2026\x2026\x2026\x2026\x2026\x2026.."), FMT_TEXT | FMT_CENTER, true, 11);*/

	EndWaitCursor();
	xls.Save(_T("Exports\\BAO CAO THU TIEN PHI KHAM.xls"));
} 
void CRMGeneralExaminationFeeReport::OnCloseSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
	
} 
int CRMGeneralExaminationFeeReport::OnAddRMGeneralExaminationFeeReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CRMGeneralExaminationFeeReport::OnEditRMGeneralExaminationFeeReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CRMGeneralExaminationFeeReport::OnDeleteRMGeneralExaminationFeeReport(){
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
 		OnCancelRMGeneralExaminationFeeReport(); 
 	}
	return 0;
}
int CRMGeneralExaminationFeeReport::OnSaveRMGeneralExaminationFeeReport(){
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
 		//OnRMGeneralExaminationFeeReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CRMGeneralExaminationFeeReport::OnCancelRMGeneralExaminationFeeReport(){
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
int CRMGeneralExaminationFeeReport::OnRMGeneralExaminationFeeReportListLoadData(){
	return 0;
}
CString CRMGeneralExaminationFeeReport::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	UpdateData(true);
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND i.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	//szWhere.AppendFormat(_T(" AND i.hfe_staff = '%s'"), pMF->GetCurrentUser());
	/*if (!m_wndObject.GetCurrent(0).IsEmpty())
	{
		szWhere.AppendFormat(_T(" and hd_object = '%s' "), m_wndObject.GetCurrent(0));
	}*/
	szSQL.Format(_T(" SELECT") \
				_T("   (SELECT") \
				_T("      hrl_name") \
				_T("    FROM   hms_roomlist") \
				_T("    WHERE  hrl_id=he_roomid") \
				_T("           AND he_deptid=hrl_deptid) AS roomname,") \
				_T("   Sum(cost)                         AS money,") \
				_T("   Sum(examed)                       AS totalexam,") \
				_T("   (SELECT") \
				_T("      hfl_name") \
				_T("    FROM   hms_fee_list") \
				_T("    WHERE  hfl_feeid=hfe_itemid)     AS examtype") \
				_T(" FROM   (SELECT") \
				_T("           f.hfe_cost AS cost,") \
				_T("           f.hfe_orderid,") \
				_T("           1          AS examed,") \
				_T("           hfe_itemid") \
				_T("         FROM   hms_fee_invoice i") \
				_T("                LEFT JOIN hms_fee f ON (i.hfe_invoiceno=f.hfe_invoiceno)") \
				_T("				LEFT JOIN hms_doc ON (hd_docno = i.hfe_docno)") \
				_T("         WHERE  i.hfe_status='P'") \
				_T("                AND f.hfe_itemid='D0000001' AND hd_object = 7 ") \
				_T("		 %s)") \
				_T("        LEFT JOIN hms_exam ON (he_receptidx=hfe_orderid)") \
				_T(" GROUP  BY he_roomid,") \
				_T("           hfe_itemid,") \
				_T("           he_deptid ") \
				_T(" ORDER BY he_roomid") , szWhere);
	//QueryOpener(szSQL);
	return szSQL;
}