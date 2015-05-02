#include "stdafx.h"
#include "MAInwardExportSheetList.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CMAInwardExportSheetList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CMAInwardExportSheetList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CMAInwardExportSheetList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CMAInwardExportSheetList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CMAInwardExportSheetList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CMAInwardExportSheetList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CMAInwardExportSheetList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CMAInwardExportSheetList *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CMAInwardExportSheetList* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CMAInwardExportSheetList *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CMAInwardExportSheetList *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CMAInwardExportSheetList *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CMAInwardExportSheetList *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CMAInwardExportSheetList *)pWnd)->OnStockAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CMAInwardExportSheetList *pVw = (CMAInwardExportSheetList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CMAInwardExportSheetList *pVw = (CMAInwardExportSheetList *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddPMSExportSheetListFnc(CWnd *pWnd){
	 return ((CMAInwardExportSheetList*)pWnd)->OnAddPMSExportSheetList();
} 
static int _OnEditPMSExportSheetListFnc(CWnd *pWnd){
	 return ((CMAInwardExportSheetList*)pWnd)->OnEditPMSExportSheetList();
} 
static int _OnDeletePMSExportSheetListFnc(CWnd *pWnd){
	 return ((CMAInwardExportSheetList*)pWnd)->OnDeletePMSExportSheetList();
} 
static int _OnSavePMSExportSheetListFnc(CWnd *pWnd){
	 return ((CMAInwardExportSheetList*)pWnd)->OnSavePMSExportSheetList();
} 
static int _OnCancelPMSExportSheetListFnc(CWnd *pWnd){
	 return ((CMAInwardExportSheetList*)pWnd)->OnCancelPMSExportSheetList();
} 
CMAInwardExportSheetList::CMAInwardExportSheetList(CWnd *pParent){

	m_nDlgWidth = 440;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CMAInwardExportSheetList::~CMAInwardExportSheetList(){
}
void CMAInwardExportSheetList::OnCreateComponents(){
	m_wndExportSheetList.Create(this, _T("Export Sheet List"), 5, 5, 575, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 90, 85);
	m_wndStock.Create(this,95, 60, 570, 85); 
	m_wndByServPrice.Create(this, _T("By Serv Price"), 5, 95, 100, 120);
	m_wndByServPrice.EnableWindow(FALSE);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 385, 95, 495, 120);
	m_wndExport.Create(this, _T("&Export"), 500, 95, 575, 120);

}
void CMAInwardExportSheetList::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);


	m_wndStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStock.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

}
void CMAInwardExportSheetList::OnSetWindowEvents(){
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndStock.SetEvent(WE_SELENDOK, _OnStockSelendokFnc);
	//m_wndStock.SetEvent(WE_SETFOCUS, _OnStockSetfocusFnc);
	//m_wndStock.SetEvent(WE_KILLFOCUS, _OnStockKillfocusFnc);
	m_wndStock.SetEvent(WE_SELCHANGE, _OnStockSelectChangeFnc);
	m_wndStock.SetEvent(WE_LOADDATA, _OnStockLoadDataFnc);
	//m_wndStock.SetEvent(WE_ADDNEW, _OnStockAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CMAInwardExportSheetList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_Check(pDX, m_wndByServPrice.GetDlgCtrlID(), m_nByServPrice);
}
void CMAInwardExportSheetList::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CMAInwardExportSheetList::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CMAInwardExportSheetList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_nByServPrice = 0;
}
int CMAInwardExportSheetList::SetMode(int nMode){
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
/*void CMAInwardExportSheetList::OnFromDateChange(){
	
} */
/*void CMAInwardExportSheetList::OnFromDateSetfocus(){
	
} */
/*void CMAInwardExportSheetList::OnFromDateKillfocus(){
	
} */
int CMAInwardExportSheetList::OnFromDateCheckValue(){
	return 0;
} 
/*void CMAInwardExportSheetList::OnToDateChange(){
	
} */
/*void CMAInwardExportSheetList::OnToDateSetfocus(){
	
} */
/*void CMAInwardExportSheetList::OnToDateKillfocus(){
	
} */
int CMAInwardExportSheetList::OnToDateCheckValue(){
	return 0;
} 
void CMAInwardExportSheetList::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CMAInwardExportSheetList::OnStockSelendok(){
	 
}
/*void CMAInwardExportSheetList::OnStockSetfocus(){
	
}*/
/*void CMAInwardExportSheetList::OnStockKillfocus(){
	
}*/

long CMAInwardExportSheetList::OnStockLoadData()
{
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageList(&m_wndStock, m_szStockKey);
}
/*void CMAInwardExportSheetList::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CMAInwardExportSheetList::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString szSQL, tmpStr, szSysDate, szRptName, szOldType, szNewType, szOldDept, szNewDept;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	int nIdx = 0;
	double nAmount = 0;
	long double nGrpTotal = 0, nTypeTotal = 0;
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	szRptName.Format(_T("Reports/HMS/PMS_BANGKEPHIEULINH.RPT"));
	if (!rpt.Init(szRptName))
		return;
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	if (!m_szStockKey.IsEmpty())
		rpt.GetReportHeader()->SetValue(_T("StockName"), m_wndStock.GetCurrent(1));
	CReportSection *rptDetail;
	szNewType = _T("XXX");
	szNewDept = _T("XX");
	while (!rs.IsEOF()){
		rs.GetValue(_T("doc_type"), szNewType);
		rs.GetValue(_T("department_to"), szNewDept);
		if (szOldType != szNewType)
		{
			if (nGrpTotal > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->GetItem(_T("TotalGroup"))->SetItalic(true);
				rptDetail->GetItem(_T("TotalGroup"))->SetBold(false);
				rptDetail->GetItem(_T("s7"))->SetItalic(true);
				rptDetail->GetItem(_T("s7"))->SetBold(false);
				tmpStr.Format(_T("%.2f"), nGrpTotal);
				rptDetail->SetValue(_T("s7"), tmpStr);
				nTypeTotal += nGrpTotal;
				nGrpTotal = 0;
			}
			if (nTypeTotal > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				tmpStr.Format(_T("%.2f"), nTypeTotal);
				rptDetail->SetValue(_T("s7"), tmpStr);
				nTypeTotal = 0;
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("doc_typename")));
			szOldType = szNewType;
			szOldDept.Empty();
		}
		if (szOldDept != szNewDept)
		{
			if (nGrpTotal > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->GetItem(_T("TotalGroup"))->SetItalic(true);
				rptDetail->GetItem(_T("TotalGroup"))->SetBold(false);
				rptDetail->GetItem(_T("s7"))->SetItalic(true);
				rptDetail->GetItem(_T("s7"))->SetBold(false);
				tmpStr.Format(_T("%.2f"), nGrpTotal);
				rptDetail->SetValue(_T("s7"), tmpStr);
				nTypeTotal += nGrpTotal;
				nGrpTotal = 0;
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rptDetail->GetItem(_T("GroupName"))->SetItalic(true);
			rptDetail->GetItem(_T("GroupName"))->SetBold(false);
			rptDetail->SetValue(_T("GroupName"), pMF->GetDepartmentName(rs.GetValue(_T("department_to"))));
			szOldDept = szNewDept;
		}
		rptDetail = rpt.AddDetail();
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("order_no"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("approval_date"), tmpStr);
		rptDetail->SetValue(_T("3"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		tmpStr.Empty();
		rs.GetValue(_T("descr"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("doc_typename"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rs.GetValue(_T("amount"), nAmount);
		nGrpTotal += nAmount;
		FormatCurrency(nAmount, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rs.MoveNext();
	}
	if (nGrpTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->GetItem(_T("TotalGroup"))->SetItalic(true);
		rptDetail->GetItem(_T("TotalGroup"))->SetBold(false);
		rptDetail->GetItem(_T("s7"))->SetItalic(true);
		rptDetail->GetItem(_T("s7"))->SetBold(false);
		tmpStr.Format(_T("%.2f"), nGrpTotal);
		rptDetail->SetValue(_T("s7"), tmpStr);
		nTypeTotal += nGrpTotal;
	}
	if (nTypeTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		tmpStr.Format(_T("%.2f"), nTypeTotal);
		rptDetail->SetValue(_T("s7"), tmpStr);
	}
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CMAInwardExportSheetList::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szOldType, szNewType;
	int nCol = 0, nRow = 0, nIdx = 0;
	double nAmount = 0;
	long double nGrpTotal = 0;
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 12);
	xls.SetColumnWidth(3, 36);
	xls.SetColumnWidth(5, 25);
	xls.SetColumnWidth(6, 10);
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 7);
	xls.SetCellMergedColumns(nCol, nRow + 3, 7);
	xls.SetCellMergedColumns(nCol, nRow + 4, 4);
	xls.SetCellText(nCol, nRow, _T("HEALTHSERVICE"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, _T("HOSPITALNAME"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, _T("\x42\x1EA2NG K\xCA PHI\x1EBEU \x58U\x1EA4T"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	tmpStr.Format(_T("Kho : %s"), m_wndStock.GetCurrent(1));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT, false, 11);

	//Column Heading
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("S\x1ED1 phi\x1EBFu"));
	arrCol.Add(_T("Ng\xE0y \x78u\x1EA5t"));
	arrCol.Add(_T("\x110\x1A1n v\x1ECB nh\x1EADn h\xE0ng"));
	arrCol.Add(_T("H\x1ECD v\xE0 t\xEAn"));
	arrCol.Add(_T("S\x1ED1 ti\x1EC1n"));
	nRow = 5;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10); 
	}
	nRow = 6;
	while (!rs.IsEOF()){
		rs.GetValue(_T("doc_typename"), szNewType);
		if (szOldType != szNewType)
		{
			if (nGrpTotal > 0)
			{
				xls.SetCellText(4, nRow, _T("T\x1ED5ng nh\xF3m phi\x1EBFu"), FMT_TEXT, true);
				tmpStr.Format(_T("%.2f"), nGrpTotal);
				xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);
				nRow++;
				nGrpTotal = 0;
			}
			xls.SetCellText(0, nRow, rs.GetValue(_T("doc_typename")), FMT_TEXT, true);
			szOldType = szNewType;
			nRow++;
		}
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		rs.GetValue(_T("order_no"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_CENTER);
		rs.GetValue(_T("approval_date"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, CDate::Convert(tmpStr, yyyymmdd , ddmmyyyy), FMT_TEXT);
		rs.GetValue(_T("descr"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("doc_typename"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("amount"), nAmount);
		nGrpTotal += nAmount;
		xls.SetCellText(nCol + 5, nRow, double2str(nAmount), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nGrpTotal > 0)
	{
		xls.SetCellText(4, nRow, _T("T\x1ED5ng nh\xF3m phi\x1EBFu"), FMT_TEXT, true);
		tmpStr.Format(_T("%.2f"), nGrpTotal);
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);
		nRow++;
	}
	xls.Save(_T("Exports\\Bang Ke Phieu Linh Vat Tu.xls"));
	
} 
int CMAInwardExportSheetList::OnAddPMSExportSheetList(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CMAInwardExportSheetList::OnEditPMSExportSheetList(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CMAInwardExportSheetList::OnDeletePMSExportSheetList(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
 	CString szSQL; 
 	if(ShowMessage(1, MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2) == IDNO) 
 		return -1; 
 	szSQL.Format(_T("DELETE FROM  WHERE  AND") ); 
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret >= 0){ 
 		SetMode(VM_NONE); 
 		OnCancelPMSExportSheetList(); 
 	}
	return 0;
}
int CMAInwardExportSheetList::OnSavePMSExportSheetList(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT) 
 		return -1; 
 	if(!IsValidateData()) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
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
 		//OnPMSExportSheetListListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CMAInwardExportSheetList::OnCancelPMSExportSheetList(){
 	if(GetMode() == VM_EDIT) 
 	{ 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 		SetMode(VM_NONE); 
 	} 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
 	return 0;
} 
int CMAInwardExportSheetList::OnPMSExportSheetListListLoadData(){
	return 0;
}

CString CMAInwardExportSheetList::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
	CString szSQL, szWhere, szIssueSQL, szReturnSQL, szCondition;
	//szCondition = _T("product_taxprice");
	//if (m_nByServPrice == 1)
	szCondition = _T("sale_price");
	szWhere.Format(_T(" AND approval_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')")
		, m_szFromDate, m_szToDate);
	if (!m_szStockKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND storage_id = %s"), m_szStockKey);
	szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	//Cac tac vu xuat#
	szIssueSQL.Format(_T(" SELECT    1 transaction_type, ") \
					_T("           mt_orderno order_no, ") \
					_T("           mt_approveddate approval_date, ") \
					_T("           mt_doctype doc_type, ") \
					_T("           (mt_department_to_id) department_to, ") \
					_T("           (mt_department_to_id) descr, ") \
					_T("           mtl_qtyorder qty, ") \
					_T("           DECODE(0, mtl_saleprice, mtl_taxprice, mtl_saleprice) sale_price, ") \
					_T("           mt_storage_id storage_id, ") \
					_T("           mtl_product_item_id item_id ") \
					_T(" FROM      m_transaction ") \
					_T(" LEFT JOIN m_transactionline ON ( mt_transaction_id = mtl_transaction_id ) ") \
					_T(" WHERE     mt_doctype NOT IN ( 'DRO', 'SRO', 'CRO', 'PRO', 'BIO' ) ") \
					_T(" AND mtl_qtydelivery > 0 AND mt_status = 'A' "));
	//Ban le
	szIssueSQL += _T(" UNION ALL ");
	szIssueSQL.AppendFormat(_T(" SELECT    1, ") \
						_T("           so_orderno, ") \
						_T("           so_approveddate, ") \
						_T("           so_doctype, ") \
						_T("           CASE WHEN Nvl(so_partner_type_id, 'W') = 'I' THEN so_partneraddress ") \
						_T("           ELSE cast('X' as nvarchar2(1))") \
						_T("           END, ") \
						_T("           CASE WHEN Nvl(so_partner_type_id, 'W') = 'I' THEN '['||so_partneraddress||']'||Get_patientname(so_docno) ") \
						_T("           ELSE Get_patientname(so_docno) ") \
						_T("           END, ") \
						_T("           sol_qtyorder, ") \
						_T("           sol_unitprice, ") \
						_T("           so_storage_id, ") \
						_T("           sol_product_item_id ") \
						_T(" FROM      sale_order ") \
						_T(" LEFT JOIN sale_orderline ON ( so_sale_order_id = sol_sale_order_id ) ") \
						_T(" WHERE     so_status = 'A'"));
	//Don noi tru
	szIssueSQL += _T(" UNION ALL ");
	szIssueSQL.AppendFormat(_T(" SELECT    1, ") \
						_T("           cast(hpo_orderid as nvarchar2(10)), ") \
						_T("           hpo_approvedate, ") \
						_T("           hpo_doctype, ") \
						_T("		   hpo_deptid,") \
						_T("           get_patientname(hpo_docno), ") \
						_T("           hpol_qtyorder, ") \
						_T("           hpol_unitprice, ") \
						_T("           hpo_storage_id, ") \
						_T("           hpol_product_item_id ") \
						_T(" FROM      hms_ipharmaorder ") \
						_T(" LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
						_T(" WHERE     hpo_ordertype <> 'C' AND Coalesce(hpo_transaction_id, 0) = 0 AND hpo_orderstatus = 'A' "));
	//Ky gui
	szIssueSQL += _T(" UNION ALL ");
	szIssueSQL.AppendFormat(_T(" SELECT    1, ") \
				_T("           mt_orderno, ") \
				_T("           mt_approveddate, ") \
				_T("           mt_doctype, ") \
				_T("		   NVL(ho_deptid, mt_department_to_id),") \
				_T("           '['||NVL(ho_deptid, mt_department_to_id)||']'||get_patientname(hpo_docno), ") \
				_T("           hpol_qtyorder, ") \
				_T("           hpol_unitprice, ") \
				_T("           mt_storage_id, ") \
				_T("           hpol_product_item_id ") \
				_T(" FROM      purchase_orderline2 ") \
				_T(" LEFT JOIN m_transaction ON ( mt_transaction_id = pol_transaction_id ) ") \
				_T(" LEFT JOIN hms_ipharmaorderline ON ( pol_orderid = hpol_orderid AND pol_product_id = hpol_product_id ) ") \
				_T(" LEFT JOIN hms_ipharmaorder ON (hpo_orderid = hpol_orderid)") \
				_T(" LEFT JOIN hms_operation ON (ho_idx = hpo_reference_id)") \
				_T(" WHERE     mt_status = 'A' "));
	//Tra lai
	szReturnSQL.Format(_T(" SELECT    16, ") \
					_T("           mt_orderno, ") \
					_T("           mt_approveddate, ") \
					_T("           mt_doctype, ") \
					_T("		   mt_department_id,") \
					_T("           (mt_department_id), ") \
					_T("           hpol_qtyreverse, ") \
					_T("		   hpol_unitprice, ") \
					_T("           mt_storage_to_id, ") \
					_T("           hpol_product_item_id ") \
					_T(" FROM      hms_ipharmaorderline ") \
					_T(" LEFT JOIN m_transaction ON ( mt_transaction_id = hpol_retorder_id) ") \
					_T(" WHERE     mt_doctype = 'DRO' AND mt_status = 'A' ") \
					_T(" UNION ALL") \
					_T(" SELECT    16, ") \
					_T("           mt_orderno, ") \
					_T("           mt_approveddate, ") \
					_T("           mt_doctype, ") \
					_T("		   mt_department_id,") \
					_T("           (mt_department_id), ") \
					_T("           mtl_qtyorder, ") \
					_T("		   DECODE(0, mtl_saleprice, mtl_taxprice, mtl_saleprice) END, ") \
					_T("           mt_storage_to_id, ") \
					_T("           mtl_product_item_id ") \
					_T(" FROM      m_transaction ") \
					_T(" LEFT JOIN m_transactionline ON ( mt_transaction_id = mtl_transaction_id ) ") \
					_T(" WHERE     mt_doctype IN ('SRO', 'CRO', 'PRO') AND mt_status = 'A' "));
	//Tra lai ban le
	szReturnSQL += _T(" UNION ALL ");
	szReturnSQL.AppendFormat(_T(" SELECT    16, ") \
					_T("           hpo_documentno, ") \
					_T("           hpo_approveddate, ") \
					_T("           hpo_doctype, ") \
					_T("           CASE WHEN Nvl(so_partner_type_id, 'W') = 'I' THEN so_partneraddress ") \
					_T("           ELSE cast('X' as nvarchar2(1))") \
					_T("           END, ") \
					_T("           CASE WHEN Nvl(so_partner_type_id, 'W') = 'W' THEN Get_patientname(hpo_docno) ") \
					_T("           ELSE '['||so_partneraddress||']'||Get_patientname(hpo_docno) ") \
					_T("           END, ") \
					_T("           hpol_qtyreturn, ") \
					_T("           hpol_unitprice, ") \
					_T("           hpo_storage_id, ") \
					_T("           hpol_product_item_id ") \
					_T(" FROM      hms_pharmareturnorder ") \
					_T(" LEFT JOIN hms_pharmareturnorder_line ON ( hpo_pharmareturnorder_id = hpol_pharmareturnorder_Id ) ") \
					_T(" LEFT JOIN sale_order ON ( so_sale_order_id = hpo_orderid ) ") \
					_T(" WHERE     hpo_doctype = 'RRO' AND hpo_status = 'A' "));

	szSQL.Format(_T(" SELECT order_no, approval_date, doc_type, get_doctype(doc_type) doc_typename, ") \
				 _T("		 department_to,") \
				 _T("		 descr,") \
				 _T("		 SUM(qty* %s) amount ") \
				 _T(" FROM (%s UNION ALL %s)") \
				 _T(" LEFT JOIN m_productitem_view ON (item_id = product_item_id)") \
				 _T(" WHERE 1=1 %s") \
				 _T(" GROUP BY transaction_type, order_no, approval_date, doc_type, department_to, descr") \
				 _T(" ORDER BY transaction_type, doc_type, department_to"), szCondition, szIssueSQL, szReturnSQL, szWhere);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}

BEGIN_MESSAGE_MAP(CMAInwardExportSheetList, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CMAInwardExportSheetList::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
