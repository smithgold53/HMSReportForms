#include "stdafx.h"
#include "PMSInwardExportSheetList.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSInwardExportSheetList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSInwardExportSheetList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSInwardExportSheetList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSInwardExportSheetList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSInwardExportSheetList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSInwardExportSheetList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSInwardExportSheetList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSInwardExportSheetList *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSInwardExportSheetList* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CPMSInwardExportSheetList *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CPMSInwardExportSheetList *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CPMSInwardExportSheetList *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CPMSInwardExportSheetList *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CPMSInwardExportSheetList *)pWnd)->OnStockAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSInwardExportSheetList *pVw = (CPMSInwardExportSheetList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSInwardExportSheetList *pVw = (CPMSInwardExportSheetList *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddPMSExportSheetListFnc(CWnd *pWnd){
	 return ((CPMSInwardExportSheetList*)pWnd)->OnAddPMSExportSheetList();
} 
static int _OnEditPMSExportSheetListFnc(CWnd *pWnd){
	 return ((CPMSInwardExportSheetList*)pWnd)->OnEditPMSExportSheetList();
} 
static int _OnDeletePMSExportSheetListFnc(CWnd *pWnd){
	 return ((CPMSInwardExportSheetList*)pWnd)->OnDeletePMSExportSheetList();
} 
static int _OnSavePMSExportSheetListFnc(CWnd *pWnd){
	 return ((CPMSInwardExportSheetList*)pWnd)->OnSavePMSExportSheetList();
} 
static int _OnCancelPMSExportSheetListFnc(CWnd *pWnd){
	 return ((CPMSInwardExportSheetList*)pWnd)->OnCancelPMSExportSheetList();
} 
CPMSInwardExportSheetList::CPMSInwardExportSheetList(CWnd *pParent){

	m_nDlgWidth = 440;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CPMSInwardExportSheetList::~CPMSInwardExportSheetList(){
}
void CPMSInwardExportSheetList::OnCreateComponents(){
	m_wndExportSheetList.Create(this, _T("Export Sheet List"), 5, 5, 575, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 90, 85);
	m_wndStock.Create(this,95, 60, 570, 85); 
	m_wndByServPrice.Create(this, _T("By Serv Price"), 5, 95, 100, 120);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 385, 95, 495, 120);
	m_wndExport.Create(this, _T("&Export"), 500, 95, 575, 120);

}
void CPMSInwardExportSheetList::OnInitializeComponents(){
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
void CPMSInwardExportSheetList::OnSetWindowEvents(){
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
void CPMSInwardExportSheetList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_Check(pDX, m_wndByServPrice.GetDlgCtrlID(), m_nByServPrice);
}
void CPMSInwardExportSheetList::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPMSInwardExportSheetList::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CPMSInwardExportSheetList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_nByServPrice = 0;
}
int CPMSInwardExportSheetList::SetMode(int nMode){
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
/*void CPMSInwardExportSheetList::OnFromDateChange(){
	
} */
/*void CPMSInwardExportSheetList::OnFromDateSetfocus(){
	
} */
/*void CPMSInwardExportSheetList::OnFromDateKillfocus(){
	
} */
int CPMSInwardExportSheetList::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSInwardExportSheetList::OnToDateChange(){
	
} */
/*void CPMSInwardExportSheetList::OnToDateSetfocus(){
	
} */
/*void CPMSInwardExportSheetList::OnToDateKillfocus(){
	
} */
int CPMSInwardExportSheetList::OnToDateCheckValue(){
	return 0;
} 
void CPMSInwardExportSheetList::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSInwardExportSheetList::OnStockSelendok(){
	 
}
/*void CPMSInwardExportSheetList::OnStockSetfocus(){
	
}*/
/*void CPMSInwardExportSheetList::OnStockKillfocus(){
	
}*/

long CPMSInwardExportSheetList::OnStockLoadData()
{
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageList(&m_wndStock, m_szStockKey);
}
/*void CPMSInwardExportSheetList::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSInwardExportSheetList::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString szSQL, tmpStr, szSysDate, szPrice, szRptName, szOldType, szNewType;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	int nIdx = 0;
	double nAmount = 0;
	long double nGrpTotal = 0;
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
	szPrice = _T("vatamount");
	if (m_nByServPrice)
		szPrice = _T("servamount");
	while (!rs.IsEOF()){
		rs.GetValue(_T("mt_doctype"), szNewType);
		if (szOldType != szNewType)
		{
			if (nGrpTotal > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				tmpStr.Format(_T("%.2f"), nGrpTotal);
				rptDetail->SetValue(_T("s7"), tmpStr);
				nGrpTotal = 0;
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("doctype")));
			szOldType = szNewType;
		}
		rptDetail = rpt.AddDetail();
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("orderno"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("expdate"), tmpStr);
		rptDetail->SetValue(_T("3"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		tmpStr.Empty();
		rs.GetValue(_T("deptname"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("doctype"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rs.GetValue(szPrice, nAmount);
		nGrpTotal += nAmount;
		FormatCurrency(nAmount, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rs.MoveNext();
	}
	if (nGrpTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		tmpStr.Format(_T("%.2f"), nGrpTotal);
		rptDetail->SetValue(_T("s7"), tmpStr);
	}
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CPMSInwardExportSheetList::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szPrice, szOldType, szNewType;
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
	szPrice = _T("vatamount");
	if (m_nByServPrice)
		szPrice = _T("servamount");
	while (!rs.IsEOF()){
		rs.GetValue(_T("mt_doctype"), szNewType);
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
			xls.SetCellText(0, nRow, rs.GetValue(_T("doctype")), FMT_TEXT, true);
			szOldType = szNewType;
			nRow++;
		}
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		rs.GetValue(_T("orderno"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_CENTER);
		rs.GetValue(_T("expdate"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, CDate::Convert(tmpStr, yyyymmdd , ddmmyyyy), FMT_TEXT);
		rs.GetValue(_T("deptname"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("doctype"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(szPrice, nAmount);
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
	xls.Save(_T("Exports\\Bang Ke Phieu Linh.xls"));
	
} 
int CPMSInwardExportSheetList::OnAddPMSExportSheetList(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CPMSInwardExportSheetList::OnEditPMSExportSheetList(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CPMSInwardExportSheetList::OnDeletePMSExportSheetList(){
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
int CPMSInwardExportSheetList::OnSavePMSExportSheetList(){
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
int CPMSInwardExportSheetList::OnCancelPMSExportSheetList(){
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
int CPMSInwardExportSheetList::OnPMSExportSheetListListLoadData(){
	return 0;
}

CString CPMSInwardExportSheetList::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
	CString szSQL, szWhere, szTransactionWhere, szSaleWhere, szWhere1;
	szWhere.Format(_T(" AND hpo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szWhere1.Format(_T(" AND hpo_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szSaleWhere.Format(_T(" AND so_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szTransactionWhere.Format(_T(" AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szStockKey.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" AND hpo_storage_id = %d"), ToInt(m_szStockKey));
		szWhere.AppendFormat(_T(" AND hpo_storage_id = %d"), ToInt(m_szStockKey));
		szSaleWhere.AppendFormat(_T(" AND so_storage_id = %d"), str2int(m_szStockKey));
		szTransactionWhere.AppendFormat(_T(" AND mt_storage_to_id = %d"), ToInt(m_szStockKey));
	}
	szSQL.Format(_T(" SELECT idx, orderno, ") \
				_T("        expdate, ") \
				_T("		mt_doctype, ") \
				_T("		Get_doctype(mt_doctype) AS doctype, ") \
				_T("        deptname, ") \
				_T("        SUM(vatamount)  AS vatamount, ") \
				_T("        SUM(servamount) AS servamount ") \
				_T(" FROM   (") \
				_T("         SELECT    '1' as idx, nvl(mt_orderno, hpo_orderid) AS orderno, ") \
				_T("                   hpo_approvedate AS expdate, ") \
				_T("                   NVL(mt_doctype, hpo_doctype) mt_doctype, ") \
				_T("                   Get_departmentname(hpo_deptid) AS deptname, ") \
				_T("                   hpol_qtyorder * product_taxprice AS vatamount, ") \
				_T("                   hpol_qtyorder * mtl_saleprice   AS servamount ") \
				_T("         FROM      hms_ipharmaorder ") \
				_T("         LEFT JOIN m_transaction ON ( mt_transaction_id = hpo_transaction_id ) ") \
				_T("         LEFT JOIN m_transactionline ON ( mt_transaction_id = mtl_transaction_id ) ") \
				_T("		 LEFT JOIN hms_ipharmaorderline ON (hpol_orderid = hpo_orderid AND hpol_product_item_id = mtl_product_item_id)") \
				_T("         LEFT JOIN m_productitem_view ON ( product_item_id = mtl_product_item_id ) ") \
				_T("         WHERE  hpo_orderstatus = 'A' AND hpo_ordertype IN ('D', 'M', 'B') %s ") \
				_T("		 UNION ALL") \
				_T("         SELECT    '2', mt_orderno AS orderno, ") \
				_T("                   mt_approveddate AS expdate, ") \
				_T("                   mt_doctype , ") \
				_T("                   Get_departmentname(mt_department_id)   AS deptname, ") \
				_T("                   hpol_qtyreverse * product_taxprice AS vatamount, ") \
				_T("                   hpol_qtyreverse * hpol_unitprice   AS servamount ") \
				_T("		 FROM hms_ipharmaorder") \
				_T("         LEFT JOIN hms_ipharmaorderline ON (hpo_orderid = hpol_orderid)") \
				_T("         LEFT JOIN m_transaction ON ( mt_transaction_id = hpol_retorder_id ) ") \
				_T("         LEFT JOIN m_productitem_view ON ( product_item_id = hpol_product_item_id ) ") \
				_T("         WHERE hpol_qtyreverse > 0 AND mt_doctype = 'DRO' AND mt_status = 'A' %s") \
				_T("		 UNION ALL") \
				_T("         SELECT    '3', so_orderno AS orderno, ") \
				_T("                   so_approveddate AS expdate, ") \
				_T("                   so_doctype , ") \
				_T("                   so_partnername AS deptname, ") \
				_T("                   sol_qtyorder * product_taxprice AS vatamount, ") \
				_T("                   sol_qtyorder * sol_unitprice   AS servamount ") \
				_T("		 FROM sale_order") \
				_T("         LEFT JOIN sale_orderline ON (so_sale_order_id = sol_sale_order_id)") \
				_T("         LEFT JOIN m_productitem_view ON ( product_item_id = sol_product_item_id ) ") \
				_T("         WHERE so_status = 'A' AND sol_product_id > 0 AND so_payment = 'P' %s") \
				_T("		 UNION ALL ") \
				_T("		 SELECT	   '4', cast(hpo_pharmareturnorder_id as nvarchar2(10)),") \
				_T("				   hpo_approveddate,") \
				_T("				   hpo_doctype,") \
				_T("				   get_patientname(hpo_docno),") \
				_T("				   hpol_qtyreturn * product_taxprice,") \
				_T("				   hpol_qtyreturn * hpol_unitprice") \
				_T("		 FROM hms_pharmareturnorder ") \
				_T("		 LEFT JOIN hms_pharmareturnorder_line ON (hpo_pharmareturnorder_id = hpol_pharmareturnorder_id)") \
				_T("		 LEFT JOIN m_productitem_view ON (product_item_id = hpol_product_item_id)") \
				_T("		 WHERE hpo_status = 'A' AND hpo_doctype = 'RRO' %s") \
				_T(")") \
				_T(" GROUP  BY idx, ") \
				_T("		   orderno, ") \
				_T("           expdate, ") \
				_T("           mt_doctype, ") \
				_T("           deptname") \
				_T(" ORDER BY idx, expdate"), szWhere, szTransactionWhere, szSaleWhere, szWhere1);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}

BEGIN_MESSAGE_MAP(CPMSInwardExportSheetList, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSInwardExportSheetList::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
