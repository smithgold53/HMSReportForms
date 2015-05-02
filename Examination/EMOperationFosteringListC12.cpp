#include "stdafx.h"
#include "EMOperationFosteringListC12.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMOperationFosteringListC12 *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMOperationFosteringListC12* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMOperationFosteringListC12 *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMOperationFosteringListC12 *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMOperationFosteringListC12 *)pWnd)->OnToDateCheckValue();
}
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMOperationFosteringListC12 *)pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CEMOperationFosteringListC12 *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CEMOperationFosteringListC12 *)pWnd)->OnObjectAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMOperationFosteringListC12 *pVw = (CEMOperationFosteringListC12 *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMOperationFosteringListC12 *pVw = (CEMOperationFosteringListC12 *)pWnd;
	pVw->OnExportSelect();
} 
CEMOperationFosteringListC12::CEMOperationFosteringListC12(CWnd *pParent){

	m_nDlgWidth = 605;
	m_nDlgHeight = 405;
	SetDefaultValues();
}
CEMOperationFosteringListC12::~CEMOperationFosteringListC12(){
}
void CEMOperationFosteringListC12::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 1, 1, 431, 116);
	m_wndYearLabel.Create(this, _T("Year"), 5, 25, 85, 50);
	m_wndYear.Create(this,90, 25, 210, 50); 
	m_wndReportPeriodLabel.Create(this, _T("Period Report"), 220, 25, 300, 50);
	m_wndReportPeriod.Create(this,305, 25, 425, 50); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 5, 55, 85, 80);
	m_wndFromDate.Create(this,90, 55, 210, 80); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 55, 300, 80);
	m_wndToDate.Create(this,305, 55, 425, 80); 
	m_wndObjectLabel.Create(this, _T("Object"), 5, 85, 85, 110);
	m_wndObject.SetCheckBox(true);
	m_wndObject.Create(this,90, 85, 425, 110); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 220, 120, 320, 145);
	m_wndExport.Create(this, _T("&Export"), 325, 120, 425, 145);
}
void CEMOperationFosteringListC12::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(1024);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), FMT_INTEGER, 30);
	m_wndReportPeriod.InsertColumn(1, _T("Name"), FMT_TEXT, 80);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 50);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);
}

void CEMOperationFosteringListC12::OnSetWindowEvents(){
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
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
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
void CEMOperationFosteringListC12::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);

}
void CEMOperationFosteringListC12::SetDefaultValues(){

	m_nYear = 0;
	m_szReportPeriodKey.Empty();
	m_szObjectKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
}

int CEMOperationFosteringListC12::SetMode(int nMode){
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

/*void CEMOperationFosteringListC12::OnYearChange(){
	
} */
/*void CEMOperationFosteringListC12::OnYearSetfocus(){
	
} */
/*void CEMOperationFosteringListC12::OnYearKillfocus(){
	
} */
int CEMOperationFosteringListC12::OnYearCheckValue(){
	return 0;
}
 
void CEMOperationFosteringListC12::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
 
void CEMOperationFosteringListC12::OnReportPeriodSelendok(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	pMF->OnReportPeriodSelendok(&m_szFromDate, &m_szToDate, m_szReportPeriodKey);
	UpdateData(false);		 
}

/*void CEMOperationFosteringListC12::OnReportPeriodSetfocus(){
	
}*/
/*void CEMOperationFosteringListC12::OnReportPeriodKillfocus(){
	
}*/
long CEMOperationFosteringListC12::OnReportPeriodLoadData(){
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

/*void CEMOperationFosteringListC12::OnReportPeriodAddNew(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMOperationFosteringListC12::OnFromDateChange(){
	
} */
/*void CEMOperationFosteringListC12::OnFromDateSetfocus(){
	
} */
/*void CEMOperationFosteringListC12::OnFromDateKillfocus(){
	
} */
int CEMOperationFosteringListC12::OnFromDateCheckValue(){
	return 0;
}
 
/*void CEMOperationFosteringListC12::OnToDateChange(){
	
} */
/*void CEMOperationFosteringListC12::OnToDateSetfocus(){
	
} */
/*void CEMOperationFosteringListC12::OnToDateKillfocus(){
	
} */
int CEMOperationFosteringListC12::OnToDateCheckValue(){
	return 0;
}
void CEMOperationFosteringListC12::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	
} 
void CEMOperationFosteringListC12::OnObjectSelendok(){
	 
}
/*void CEMOperationFosteringListC12::OnObjectSetfocus(){
	
}*/
/*void CEMOperationFosteringListC12::OnObjectKillfocus(){
	
}*/
long CEMOperationFosteringListC12::OnObjectLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty()){
	};
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT ho_id as id, ho_desc as name FROM hms_object"));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CEMOperationFosteringListC12::OnObjectAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	
} */ 
void CEMOperationFosteringListC12::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CReportSection* rptDetail = NULL, *rptFooter = NULL;
	CString szSQL, tmpStr, tmpStr2, szDate, szObjects;
	double nTmp = 0, nAmt = 0, nType3 = 0, nOldDoc = 0, nNewDoc = 0;
	int nCount=0;
	CRecord rs(&pMF->m_db);
	int nIdx = 1;
	if (!rpt.Init(_T("Reports\\HMS\\BAO CAO BDPT C12.RPT")))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	for (int i = 0; i < m_wndObject.GetItemCount(); i++)
	{
		if (m_wndObject.GetCheck(i))
		{
			m_wndObject.SetCurSel(i);
			if (!szObjects.IsEmpty())
				szObjects += _T(", ");
			szObjects += m_wndObject.GetCurrent(1);
		}
	}
	if (!szObjects.IsEmpty())
		rpt.GetReportHeader()->SetValue(_T("Object"), szObjects);
	else
		rpt.GetReportHeader()->SetValue(_T("Object"), _T("To\xE0n \x62\x1ED9"));
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("patname")));
		rs.GetValue(_T("doc_no"), nNewDoc);
		if (nNewDoc != nOldDoc)
		{
			nCount++;
			nOldDoc = nNewDoc;
		}
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("doc_no")));
		if(rs.GetValue(_T("pobject")) == 'I')
		{
			rptDetail->SetValue(_T("11"), rs.GetValue(_T("cardno")));
		}
		else if(rs.GetValue(_T("pobject")) == 'D')
		{
			rptDetail->SetValue(_T("11"), rs.GetValue(_T("rank")));
		}
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("ptname")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("ptlevel")));
		nType3++;
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("thefirst")));
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("thesecond")));
		rptDetail->SetValue(_T("8"), rs.GetValue(_T("assist")));
		rs.GetValue(_T("ptcost"), nTmp);
		nAmt += nTmp;
		rptDetail->SetValue(_T("9"), double2str(nTmp));
		rs.MoveNext();
	}
	
	szDate = pMF->GetSysDate();
	rptFooter = rpt.GetReportFooter();
	tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5, 2), szDate.Left(4));
	rptFooter->SetValue(_T("PrintDate"), tmpStr);
	rptFooter->SetValue(_T("10"), double2str(nAmt));
	rptFooter->SetValue(_T("14"), double2str(nType3));
	rptFooter->SetValue(_T("15"), _T("24000"));
	rptFooter->SetValue(_T("16"), double2str(nType3*24000));
	rptFooter->SetValue(_T("TotalPatient"), int2str(nCount));
	rptFooter->SetValue(_T("TotalItem"), double2str(nType3));
	rptFooter->SetValue(_T("TotalAmount"), double2str(nAmt));
	tmpStr2.Format(_T("%f"), nAmt);
	MoneyToString(tmpStr2, tmpStr);
	tmpStr += _T(" \x111\x1ED3ng.");
	rptFooter->SetValue(_T("SumInWord"), tmpStr);
	rpt.PrintPreview();
}
 
void CEMOperationFosteringListC12::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp, szWhere, szObjects;
	UpdateData(TRUE);
	BeginWaitCursor();
	szWhere.Empty();
	int nRow =0, nCol = 0, nCount = 0, nOldDoc = 0, nNewDoc = 0;
	int nIdx = 1;
	double nTmp = 0, nAmt = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;
	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 8);	xls.SetColumnWidth(1, 16);	xls.SetColumnWidth(2, 14);	xls.SetColumnWidth(3, 14);	xls.SetColumnWidth(4, 36);	xls.SetColumnWidth(5, 8);	xls.SetColumnWidth(6, 8);	xls.SetColumnWidth(7, 8);	xls.SetColumnWidth(8, 8);	xls.SetColumnWidth(9, 22);	xls.SetColumnWidth(10, 8);	xls.SetColumnWidth(11, 8);	xls.SetColumnWidth(12, 15);

	xls.SetCellText(nCol, nRow, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol+1, nRow + 2, _T("C1.2"), FMT_TEXT | FMT_CENTER, true, 11);
	for (int i = 0; i < m_wndObject.GetItemCount(); i++)
	{
		if (m_wndObject.GetCheck(i))
		{
			m_wndObject.SetCurSel(i);
			if (!szObjects.IsEmpty())
				szObjects += _T(", ");
			szObjects += m_wndObject.GetCurrent(1);
		}
	}
	if (!szObjects.IsEmpty())
		tmpStr.Format(_T("\x42\x1EC7nh nh\xE2n %s"), szObjects);
	else
		tmpStr.Format(_T("\x42\x1EC7nh nh\xE2n"));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(0, 2, _T("KHO\x41:"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);
	xls.SetCellText(0, 3, _T("\x44\x41NH S\xC1\x43H \x110\x1EC0 NGH\x1ECA \x42\x1ED2I \x44\x1AF\x1EE0NG PH\x1EAAU THU\x1EACT, TH\x1EE6 THU\x1EACT"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 12);
	xls.SetCellText(0, 6, _T("S\x1ED1 TT"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(1, 6, _T("H\x1ECD v\xE0 t\xEAn \x62\x1EC7nh nh\xE2n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(2, 6, _T("S\x1ED1 h\x1ED3 s\x1A1"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(3, 6, _T("SH, s\x1ED1 th\x1EBB, \x63\x1EA5p \x62\x1EAD\x63"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(4, 6, _T("T\xEAn ph\x1EABu thu\x1EADt, th\x1EE7 thu\x1EADt"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(5, 6, _T("Ph\xE2n lo\x1EA1i ph\x1EABu thu\x1EADt, th\x1EE7 thu\x1EADt"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(5, 8, _T("\x110\x1EB7\x63 \x62i\x1EC7t"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(6, 8, _T("Lo\x1EA1i \x31"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(7, 8, _T("Lo\x1EA1i \x32"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(8, 8, _T("Lo\x1EA1i \x33"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(9, 0, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(9, 1, _T("\x110\x1ED9\x63 l\x1EADp - T\x1EF1 \x64o - H\x1EA1nh ph\xFA\x63"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(9, 6, _T("K\xEDp th\x1EF1\x63 hi\x1EC7n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(9, 8, _T("\x43h\xEDnh"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(10, 8, _T("Ph\x1EE5"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(11, 8, _T("Gi\xFAp vi\x1EC7\x63"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);
	xls.SetCellText(12, 6, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);

	xls.SetMerge(0, 0, 0, 2);	xls.SetMerge(1, 1, 0, 2);	xls.SetMerge(3, 3, 0, 12);	xls.SetMerge(4, 4, 0, 12);	xls.SetMerge(5, 5, 0, 12);	xls.SetMerge(6, 9, 0, 0);	xls.SetMerge(2, 2, 1, 2);	xls.SetMerge(6, 9, 1, 1);	xls.SetMerge(174, 174, 1, 11);	xls.SetMerge(6, 9, 2, 2);	xls.SetMerge(6, 9, 3, 3);	xls.SetMerge(6, 9, 4, 4);	xls.SetMerge(6, 7, 5, 8);	xls.SetMerge(8, 9, 5, 5);	xls.SetMerge(8, 9, 6, 6);	xls.SetMerge(8, 9, 7, 7);	xls.SetMerge(8, 9, 8, 8);	xls.SetMerge(0, 0, 9, 12);	xls.SetMerge(1, 1, 9, 12);	xls.SetMerge(6, 7, 9, 11);	xls.SetMerge(8, 9, 9, 9);	xls.SetMerge(8, 9, 10, 10);	xls.SetMerge(8, 9, 11, 11);	xls.SetMerge(6, 9, 12, 12);

	nRow = 10;
	nCol = 0;
	while (!rs.IsEOF())
	{
		xls.SetCellText(nCol, nRow, int2str(nIdx++));
		xls.SetCellText(nCol + 1, nRow, rs.GetValue(_T("patname")));
		rs.GetValue(_T("doc_no"), nNewDoc);
		if (nNewDoc != nOldDoc)
		{
			nCount++;
			nOldDoc = nNewDoc;
		}
		xls.SetCellText(nCol + 2, nRow, rs.GetValue(_T("doc_no")));
		if(rs.GetValue(_T("pobject")) == 'I')
		{
			xls.SetCellText(nCol + 3, nRow, rs.GetValue(_T("cardno")));
		}
		else if(rs.GetValue(_T("pobject")) == 'D')
		{
			xls.SetCellText(nCol + 3, nRow, rs.GetValue(_T("rank")));
		}
		xls.SetCellText(nCol + 4, nRow, rs.GetValue(_T("ptname")));
		xls.SetCellText(nCol + 8, nRow, rs.GetValue(_T("ptlevel")));
		xls.SetCellText(nCol + 9, nRow, rs.GetValue(_T("thefirst")));
		rs.GetValue(_T("ptcost"), nTmp);
		nAmt += nTmp;
		xls.SetCellText(nCol + 12, nRow, double2str(nTmp), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nAmt > 0)
	{
		xls.SetCellMergedColumns(nCol+1, nRow, 11);
		xls.SetCellText(nCol+1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT, true);
		tmpStr.Format(_T("%.2f"), nAmt);
		xls.SetCellText(nCol + 12, nRow, tmpStr, FMT_NUMBER1);
		nRow++;
	}
	xls.Save(_T("Exports\\Danh sach boi duong PTTT.xls"));
}	
 
CString CEMOperationFosteringListC12::GetQueryString(){
	CString szSQL, szWhere, szObjects;
	szWhere.Format(_T(" AND hpc_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	for (int i = 0; i < m_wndObject.GetItemCount(); i++)
	{
		if (m_wndObject.GetCheck(i))
		{
			m_wndObject.SetCurSel(i);
			if (!szObjects.IsEmpty())
				szObjects += _T(", ");
			szObjects += m_wndObject.GetCurrent(0);
		}
	}
	if(!szObjects.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpc_object IN(%s)"), szObjects);
	szSQL.Format(_T(" SELECT hpc_docno AS doc_no,") \
		_T("   Get_patientname(hpc_docno) AS patname,") \
		_T("   hd_cardno cardno,") \
		_T("   ho_type pobject,") \
		_T("   (SELECT ss_vndesc FROM sys_sel WHERE ss_id = 'hms_rank' AND ss_code = hp_rank) rank,") \
		_T("   hfl_name AS ptname,") \
		_T("   'X' AS ptlevel,") \
		_T("   '24000' AS ptcost,") \
		_T("   Get_username(hpc_practitioner) AS thefirst") \
		_T(" FROM hms_pacsorder") \
		_T(" LEFT JOIN hms_fee ON(hfe_docno = hpc_docno AND hfe_group = hpc_groupid)") \
		_T(" LEFT JOIN hms_fee_list ON(hfl_feeid = hfe_itemid)") \
		_T(" LEFT JOIN hms_doc ON(hd_docno = hfe_docno)") \
		_T(" LEFT JOIN hms_patient ON(hp_patientno = hd_patientno)") \
		_T(" LEFT JOIN hms_object ON(ho_id =  hpc_object)") \
		_T(" WHERE hpc_payment     ='P'") \
		_T(" AND hpc_deptid       = 'C1.2'") \
		_T(" AND hpc_groupid = 'B3200'") \
		_T(" AND hpc_practitioner IS NOT NULL") \
		_T(" AND hfl_feeid NOT IN('B320002253', 'B320002254', 'B320002255', 'B320002256', 'B320002257', 'B320002263') %s ") \
		_T(" ORDER BY patname, doc_no"), szWhere);
	return szSQL;
}