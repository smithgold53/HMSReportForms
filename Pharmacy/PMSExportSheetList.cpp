#include "stdafx.h"
#include "PMSExportSheetList.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSExportSheetList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSExportSheetList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSExportSheetList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSExportSheetList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSExportSheetList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSExportSheetList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSExportSheetList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSExportSheetList *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSExportSheetList* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CPMSExportSheetList *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CPMSExportSheetList *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CPMSExportSheetList *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CPMSExportSheetList *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CPMSExportSheetList *)pWnd)->OnStockAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSExportSheetList *pVw = (CPMSExportSheetList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSExportSheetList *pVw = (CPMSExportSheetList *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddPMSExportSheetListFnc(CWnd *pWnd){
	 return ((CPMSExportSheetList*)pWnd)->OnAddPMSExportSheetList();
} 
static int _OnEditPMSExportSheetListFnc(CWnd *pWnd){
	 return ((CPMSExportSheetList*)pWnd)->OnEditPMSExportSheetList();
} 
static int _OnDeletePMSExportSheetListFnc(CWnd *pWnd){
	 return ((CPMSExportSheetList*)pWnd)->OnDeletePMSExportSheetList();
} 
static int _OnSavePMSExportSheetListFnc(CWnd *pWnd){
	 return ((CPMSExportSheetList*)pWnd)->OnSavePMSExportSheetList();
} 
static int _OnCancelPMSExportSheetListFnc(CWnd *pWnd){
	 return ((CPMSExportSheetList*)pWnd)->OnCancelPMSExportSheetList();
} 
CPMSExportSheetList::CPMSExportSheetList(CWnd *pParent){

	m_nDlgWidth = 440;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CPMSExportSheetList::~CPMSExportSheetList(){
}
void CPMSExportSheetList::OnCreateComponents(){
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
void CPMSExportSheetList::OnInitializeComponents(){
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
void CPMSExportSheetList::OnSetWindowEvents(){
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
void CPMSExportSheetList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_Check(pDX, m_wndByServPrice.GetDlgCtrlID(), m_nByServPrice);
}
void CPMSExportSheetList::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPMSExportSheetList::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CPMSExportSheetList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_nByServPrice = 0;
}
int CPMSExportSheetList::SetMode(int nMode){
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
/*void CPMSExportSheetList::OnFromDateChange(){
	
} */
/*void CPMSExportSheetList::OnFromDateSetfocus(){
	
} */
/*void CPMSExportSheetList::OnFromDateKillfocus(){
	
} */
int CPMSExportSheetList::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSExportSheetList::OnToDateChange(){
	
} */
/*void CPMSExportSheetList::OnToDateSetfocus(){
	
} */
/*void CPMSExportSheetList::OnToDateKillfocus(){
	
} */
int CPMSExportSheetList::OnToDateCheckValue(){
	return 0;
} 
void CPMSExportSheetList::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSExportSheetList::OnStockSelendok(){
	 
}
/*void CPMSExportSheetList::OnStockSetfocus(){
	
}*/
/*void CPMSExportSheetList::OnStockKillfocus(){
	
}*/

long CPMSExportSheetList::OnStockLoadData()
{
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageList(&m_wndStock, m_szStockKey);
}
/*void CPMSExportSheetList::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSExportSheetList::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString szSQL, tmpStr, szSysDate, szPrice, szRptName;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	int nIdx = 0;
	double nAmount = 0, nTotal = 0;
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	szRptName.Format(_T("Reports/HMS/PM_BANGKEPHIEUXUAT.RPT"));
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
		rptDetail = rpt.AddDetail();
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("orderid"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("expdate"), tmpStr);
		rptDetail->SetValue(_T("3"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		tmpStr.Empty();
		rs.GetValue(_T("patname"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("id"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rs.GetValue(_T("doctype"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rs.GetValue(szPrice, nAmount);
		nTotal += nAmount;
		rptDetail->SetValue(_T("7"), double2str(nAmount));
		rs.MoveNext();
	}
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	tmpStr.Format(_T("%s"), _T("T\x1ED5ng \x63\x1ED9ng"));
	rptDetail->SetValue(_T("TotalAmount"), tmpStr);
	tmpStr.Format(_T("%f"), nTotal);
	rptDetail->SetValue(_T("s7"), tmpStr);
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CPMSExportSheetList::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szPrice;
	int nCol = 0, nRow = 0, nIdx = 0;
	double nAmount = 0, nTotal = 0;
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
	arrCol.Add(_T("M\xE3 s\x1ED1"));
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
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		rs.GetValue(_T("orderid"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_CENTER);
		rs.GetValue(_T("expdate"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, CDate::Convert(tmpStr, yyyymmdd , ddmmyyyy), FMT_TEXT);
		rs.GetValue(_T("patname"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("id"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_INTEGER | FMT_CENTER);
		rs.GetValue(_T("doctype"), tmpStr);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(szPrice, nAmount);
		nTotal += nAmount;
		xls.SetCellText(nCol + 6, nRow, double2str(nAmount), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellText(nCol + 5, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT, true, 10);
	xls.SetCellText(nCol + 6, nRow, double2str(nTotal), FMT_NUMBER1, true, 10);
	xls.Save(_T("Exports\\Bang Ke Phieu Xuat.xls"));
	
} 
int CPMSExportSheetList::OnAddPMSExportSheetList(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CPMSExportSheetList::OnEditPMSExportSheetList(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CPMSExportSheetList::OnDeletePMSExportSheetList(){
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
int CPMSExportSheetList::OnSavePMSExportSheetList(){
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
int CPMSExportSheetList::OnCancelPMSExportSheetList(){
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
int CPMSExportSheetList::OnPMSExportSheetListListLoadData(){
	return 0;
}

CString CPMSExportSheetList::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	szWhere.Format(_T("AND (hfe_createddate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szStockKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_storage_id = %d"), ToInt(m_szStockKey));
	szSQL.Format(_T(" SELECT orderid, ") \
				_T("        docno, ") \
				_T("        expdate, ") \
				_T("        doctype, ") \
				_T("        id, ") \
				_T("        patname, ") \
				_T("        SUM(vatamount)  AS vatamount, ") \
				_T("        SUM(servamount) AS servamount ") \
				_T(" FROM   (SELECT    Cast(hpo_orderid AS NVARCHAR2(15)) AS orderid, ") \
				_T("                   hpo_docno                          AS docno, ") \
				_T("                   hfe_createddate                    AS expdate, ") \
				_T("                   Get_doctype(hpo_doctype)           AS doctype, ") \
				_T("                   CASE WHEN hpo_doctype = 'PPO' THEN Get_patientname(hpo_docno) ") \
				_T("                                                 ||'- ' ") \
				_T("                                                 ||Hms_getagebydoc(hpo_docno) ") \
				_T("                                                 ||N' Tu\x1ED5i' ") \
				_T("                   ELSE Cast('' AS NVARCHAR2(1)) ") \
				_T("                   END                                AS patname, ") \
				_T("                   CASE WHEN hpo_doctype = 'PPO' THEN hpo_docno ") \
				_T("                   ELSE 0 ") \
				_T("                   END                                AS id, ") \
				_T("                   hpol_qtyissue * product_taxprice   AS vatamount, ") \
				_T("                   hpol_qtyissue * hpol_unitprice     AS servamount ") \
				_T("         FROM      hms_pharmaorder ") \
				_T("         LEFT JOIN hms_pharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T("         LEFT JOIN hms_fee_invoice ON ( hfe_docno = hpo_docno AND hfe_invoiceno = hpo_invoiceno ) ") \
				_T("         LEFT JOIN m_productitem_view ON ( product_item_id = hpol_product_item_id ) ") \
				_T("         WHERE     hpo_ordertype = 'P' ") \
				_T("               AND hfe_refstatus = 'P' ") \
				_T("               AND hfe_status = 'P' %s)") \
				_T(" GROUP  BY orderid, ") \
				_T("           docno, ") \
				_T("           expdate, ") \
				_T("           doctype, ") \
				_T("           id, ") \
				_T("           patname"), szWhere);
	return szSQL;

}
BEGIN_MESSAGE_MAP(CPMSExportSheetList, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSExportSheetList::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
