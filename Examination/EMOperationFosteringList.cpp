#include "stdafx.h"
#include "EMOperationFosteringList.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMOperationFosteringList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMOperationFosteringList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMOperationFosteringList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMOperationFosteringList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMOperationFosteringList *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMOperationFosteringList *pVw = (CEMOperationFosteringList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMOperationFosteringList *pVw = (CEMOperationFosteringList *)pWnd;
	pVw->OnExportSelect();
} 
CEMOperationFosteringList::CEMOperationFosteringList(CWnd *pParent){

	m_nDlgWidth = 500;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CEMOperationFosteringList::~CEMOperationFosteringList(){
}
void CEMOperationFosteringList::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 490, 90);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 245, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 250, 30, 330, 55);
	m_wndReportPeriod.Create(this,335, 30, 485, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 245, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 60, 330, 85);
	m_wndToDate.Create(this,335, 60, 485, 85); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 300, 95, 410, 120);
	m_wndExport.Create(this, _T("&Export"), 415, 95, 490, 120);

}
void CEMOperationFosteringList::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), FMT_INTEGER, 30);
	m_wndReportPeriod.InsertColumn(1, _T("Name"), FMT_TEXT, 80);

}

void CEMOperationFosteringList::OnSetWindowEvents(){
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
	CDate dte;
	dte.ParseDate(pMF->GetSysDate());
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	m_nYear = dte.GetYear();
	m_szReportPeriodKey.Format(_T("%d"), dte.GetMonth());
	UpdateData(false);

}
void CEMOperationFosteringList::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CEMOperationFosteringList::SetDefaultValues(){

	m_nYear = 0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}

int CEMOperationFosteringList::SetMode(int nMode){
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

/*void CEMOperationFosteringList::OnYearChange(){
	
} */
/*void CEMOperationFosteringList::OnYearSetfocus(){
	
} */
/*void CEMOperationFosteringList::OnYearKillfocus(){
	
} */
int CEMOperationFosteringList::OnYearCheckValue(){
	return 0;
}
 
void CEMOperationFosteringList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
 
void CEMOperationFosteringList::OnReportPeriodSelendok(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	pMF->OnReportPeriodSelendok(&m_szFromDate, &m_szToDate, m_szReportPeriodKey);
	UpdateData(false);	 
}

/*void CEMOperationFosteringList::OnReportPeriodSetfocus(){
	
}*/
/*void CEMOperationFosteringList::OnReportPeriodKillfocus(){
	
}*/
long CEMOperationFosteringList::OnReportPeriodLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadReportPeriodList(&m_wndReportPeriod, m_szReportPeriodKey);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && !m_szReportPeriodKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szReportPeriodKey
};
	m_wndReportPeriod.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndReportPeriod.AddItems(
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}

/*void CEMOperationFosteringList::OnReportPeriodAddNew(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMOperationFosteringList::OnFromDateChange(){
	
} */
/*void CEMOperationFosteringList::OnFromDateSetfocus(){
	
} */
/*void CEMOperationFosteringList::OnFromDateKillfocus(){
	
} */
int CEMOperationFosteringList::OnFromDateCheckValue(){
	return 0;
}
 
/*void CEMOperationFosteringList::OnToDateChange(){
	
} */
/*void CEMOperationFosteringList::OnToDateSetfocus(){
	
} */
/*void CEMOperationFosteringList::OnToDateKillfocus(){
	
} */
int CEMOperationFosteringList::OnToDateCheckValue(){
	return 0;
}
 
void CEMOperationFosteringList::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CReportSection* rptDetail = NULL, *rptFooter = NULL;
	CString szSQL, tmpStr, tmpStr2, szDate;
	double nTmp = 0, nAmt = 0, nType1 = 0, nType3 = 0, nOldDoc = 0, nNewDoc = 0;
	int nCount=0;
	CRecord rs(&pMF->m_db);
	int nIdx = 1;
	if (!rpt.Init(_T("Reports\\HMS\\BAO CAO BDPT C15.RPT")))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rs.GetValue(_T("ho_recordno"), tmpStr);
		if (tmpStr == _T("YAG.X") || tmpStr == _T("YA"))
			nType1++;
		else
			nType3++;
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("patnme")));
		rs.GetValue(_T("doc_no"), nNewDoc);
		if (nNewDoc != nOldDoc)
		{
			nCount++;
			nOldDoc = nNewDoc;
		}
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("doc_no")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("itemnme")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("itemtype")));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("thefirst")));
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("thesecond")));
		rptDetail->SetValue(_T("8"), rs.GetValue(_T("assist")));
		rs.GetValue(_T("amount"), nTmp);
		nAmt += nTmp;
		rptDetail->SetValue(_T("9"), double2str(nTmp));
		rs.MoveNext();
	}
	
	szDate = pMF->GetSysDate();
	rptFooter = rpt.GetReportFooter();
	tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5, 2), szDate.Left(4));
	rptFooter->SetValue(_T("PrintDate"), tmpStr);
	rptFooter->SetValue(_T("10"), double2str(nAmt));
	rptFooter->SetValue(_T("11"), double2str(nType1));
	rptFooter->SetValue(_T("12"), _T("85500"));
	rptFooter->SetValue(_T("13"), double2str(nType1*85500));
	rptFooter->SetValue(_T("14"), double2str(nType3));
	rptFooter->SetValue(_T("15"), _T("28500"));
	rptFooter->SetValue(_T("16"), double2str(nType3*28500));
	rptFooter->SetValue(_T("TotalPatient"), int2str(nCount));
	rptFooter->SetValue(_T("TotalItem"), double2str(nType1+nType3));
	rptFooter->SetValue(_T("TotalAmount"), double2str(nAmt));
	tmpStr2.Format(_T("%f"), nAmt);
	MoneyToString(tmpStr2, tmpStr);
	tmpStr += _T(" \x111\x1ED3ng.");
	rptFooter->SetValue(_T("SumInWord"), tmpStr);
	rpt.PrintPreview();
}
 
void CEMOperationFosteringList::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0, nCount = 0, nOldDoc = 0, nNewDoc = 0;
	int nIdx = 1;
	double nTmp = 0, nAmt = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 3, 8);
	xls.SetCellMergedColumns(nCol, nRow + 4, 8);
	xls.SetCellText(nCol, nRow, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 2, pMF->GetDepartmentName(pMF->GetCurrentUser()), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 3, _T("\x44\x41NH S\xC1\x43H \x110\x1EC0 NGH\x1ECA \x42\x1ED2I \x44\x1AF\x1EE0NG PTTT"), FMT_TEXT | FMT_CENTER, true, 12);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, false, 10);
	//Column
	xls.SetCellMergedRows(nCol, nRow + 5, 2);
	xls.SetCellMergedRows(nCol + 1, nRow + 5, 2);
	xls.SetCellMergedRows(nCol + 2, nRow + 5, 2);
	xls.SetCellMergedRows(nCol + 3, nRow + 5, 2);
	xls.SetCellMergedRows(nCol + 4, nRow + 5, 2);
	xls.SetCellMergedRows(nCol + 8, nRow + 5, 2);
	xls.SetCellMergedColumns(nCol + 5, nRow + 5, 3);
	xls.SetCellText(nCol, nRow + 5, _T("STT"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 1, nRow + 5, _T("H\x1ECD v\xE0 t\xEAn"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 2, nRow + 5, _T("T\xEAn PTTT"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 3, nRow + 5, _T("S\x1ED1 HS"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 4, nRow + 5, _T("Lo\x1EA1i PTTT"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 5, nRow + 5, _T("K\xEDp th\x1EF1\x63 hi\x1EC7n"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 5, nRow + 6, _T("\x43h\xEDnh"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 6, nRow + 6, _T("Ph\x1EE5"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 7, nRow + 6, _T("Gi\xFAp vi\x1EC7\x63"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 8, nRow + 5, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true);
	nRow = 7;
	while (!rs.IsEOF())
	{
		xls.SetCellText(nCol, nRow, int2str(nIdx++));
		xls.SetCellText(nCol + 1, nRow, rs.GetValue(_T("patnme")));
		rs.GetValue(_T("doc_no"), nNewDoc);
		if (nNewDoc != nOldDoc)
		{
			nCount++;
			nOldDoc = nNewDoc;
		}
		xls.SetCellText(nCol + 2, nRow, rs.GetValue(_T("doc_no")));
		xls.SetCellText(nCol + 3, nRow, rs.GetValue(_T("itemnme")));
		xls.SetCellText(nCol + 4, nRow, rs.GetValue(_T("itemtype")));
		xls.SetCellText(nCol + 5, nRow, rs.GetValue(_T("thefirst")));
		xls.SetCellText(nCol + 6, nRow, rs.GetValue(_T("thesecond")));
		xls.SetCellText(nCol + 7, nRow, rs.GetValue(_T("assist")));
		rs.GetValue(_T("amount"), nTmp);
		nAmt += nTmp;
		xls.SetCellText(nCol + 8, nRow, double2str(nTmp));
		nRow++;
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\Danh sach boi duong PTTT.xls"));
}	
 
CString CEMOperationFosteringList::GetQueryString(){
	CString szSQL, szWhere;
	szWhere.AppendFormat(_T(" AND ho_performdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szSQL.Format(_T(" SELECT ho_docno doc_no,") \
				_T("   Get_patientname(ho_docno)     AS patnme, ho_recordno,") \
				_T("   case when ho_recordno IN  ('YAG.X', 'YA') then 'U m\xE1u' else 'U s\x1EAF\x63 t\x1ED1' end AS itemnme,") \
				_T("   case when ho_recordno IN  ('YAG.X', 'YA') then 'Lo\x1EA1i 1' else 'Lo\x1EA1i 3' end AS itemtype,") \
				_T("   case when ho_recordno IN  ('YAG.X', 'YA') then '85500' else '28500' end AS amount,") \
				_T("   Get_username(ho_practitioner) AS thefirst,") \
				_T("   Get_username(ho_anesthetist)  AS thesecond,") \
				_T("   Get_username(ho_assistant)    AS assist") \
				_T(" FROM   hms_operation") \
				_T(" WHERE  ho_payment='P'") \
				_T(" AND ho_pdeptid = 'C15'") \
				_T(" AND ho_practitioner IS NOT NULL %s") \
				_T(" ORDER BY patnme, doc_no"), szWhere);
	return szSQL;
}