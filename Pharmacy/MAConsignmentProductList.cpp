#include "stdafx.h"
#include "MAConsignmentProductList.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CMAConsignmentProductList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CMAConsignmentProductList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CMAConsignmentProductList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CMAConsignmentProductList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CMAConsignmentProductList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CMAConsignmentProductList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CMAConsignmentProductList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CMAConsignmentProductList *)pWnd)->OnToDateCheckValue();
} 
static void _OnSupplierSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CMAConsignmentProductList* )pWnd)->OnSupplierSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSupplierSelendokFnc(CWnd *pWnd){
	((CMAConsignmentProductList *)pWnd)->OnSupplierSelendok();
}
/*static void _OnSupplierSetfocusFnc(CWnd *pWnd){
	((CMAConsignmentProductList *)pWnd)->OnSupplierKillfocus();
}*/
/*static void _OnSupplierKillfocusFnc(CWnd *pWnd){
	((CMAConsignmentProductList *)pWnd)->OnSupplierKillfocus();
}*/
static long _OnSupplierLoadDataFnc(CWnd *pWnd){
	return ((CMAConsignmentProductList *)pWnd)->OnSupplierLoadData();
}
/*static void _OnSupplierAddNewFnc(CWnd *pWnd){
	((CMAConsignmentProductList *)pWnd)->OnSupplierAddNew();
}*/
static void _OnNotImportedSelectFnc(CWnd *pWnd){
	 ((CMAConsignmentProductList*)pWnd)->OnNotImportedSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CMAConsignmentProductList *pVw = (CMAConsignmentProductList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CMAConsignmentProductList *pVw = (CMAConsignmentProductList *)pWnd;
	pVw->OnExportSelect();
} 
CMAConsignmentProductList::CMAConsignmentProductList(CWnd *pWnd){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CMAConsignmentProductList::~CMAConsignmentProductList(){
}
void CMAConsignmentProductList::OnCreateComponents(){
	m_wndConsignmentPatientList.Create(this, _T("Consignment Patient List"), 5, 5, 470, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 235, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 240, 30, 320, 55);
	m_wndToDate.Create(this,325, 30, 465, 55); 
	m_wndSupplierLabel.Create(this, _T("Supplier"), 10, 60, 90, 85);
	m_wndSupplier.Create(this,95, 60, 465, 85); 
	m_wndNotImported.Create(this, _T("Not Imported"), 5, 95, 130, 120);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 265, 95, 365, 120);
	m_wndExport.Create(this, _T("&Export"), 370, 95, 470, 120);

}
void CMAConsignmentProductList::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndSupplier.SetCheckValue(true);
	m_wndSupplier.LimitText(35);


	m_wndSupplier.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndSupplier.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);

}

void CMAConsignmentProductList::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndSupplier.SetEvent(WE_SELENDOK, _OnSupplierSelendokFnc);
	//m_wndSupplier.SetEvent(WE_SETFOCUS, _OnSupplierSetfocusFnc);
	//m_wndSupplier.SetEvent(WE_KILLFOCUS, _OnSupplierKillfocusFnc);
	m_wndSupplier.SetEvent(WE_SELCHANGE, _OnSupplierSelectChangeFnc);
	m_wndSupplier.SetEvent(WE_LOADDATA, _OnSupplierLoadDataFnc);
	//m_wndSupplier.SetEvent(WE_ADDNEW, _OnSupplierAddNewFnc);
	m_wndNotImported.SetEvent(WE_CLICK, _OnNotImportedSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	m_wndNotImported.ShowWindow(SW_HIDE);
	UpdateData(false);

}
void CMAConsignmentProductList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szSupplierKey);
	DDX_Check(pDX, m_wndNotImported.GetDlgCtrlID(), m_bNotImported);

}
void CMAConsignmentProductList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szSupplierKey.Empty();
	m_bNotImported = FALSE;

}

int CMAConsignmentProductList::SetMode(int nMode){
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

/*void CMAConsignmentProductList::OnFromDateChange(){
	
} */
/*void CMAConsignmentProductList::OnFromDateSetfocus(){
	
} */
/*void CMAConsignmentProductList::OnFromDateKillfocus(){
	
} */
int CMAConsignmentProductList::OnFromDateCheckValue(){
	return 0;
}
 
/*void CMAConsignmentProductList::OnToDateChange(){
	
} */
/*void CMAConsignmentProductList::OnToDateSetfocus(){
	
} */
/*void CMAConsignmentProductList::OnToDateKillfocus(){
	
} */
int CMAConsignmentProductList::OnToDateCheckValue(){
	return 0;
}
 
void CMAConsignmentProductList::OnSupplierSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
void CMAConsignmentProductList::OnSupplierSelendok(){
	 
}

/*void CMAConsignmentProductList::OnSupplierSetfocus(){
	
}*/
/*void CMAConsignmentProductList::OnSupplierKillfocus(){
	
}*/
long CMAConsignmentProductList::OnSupplierLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadPartnerList(&m_wndSupplier, m_szSupplierKey);
}

/*void CMAConsignmentProductList::OnSupplierAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
	void CMAConsignmentProductList::OnNotImportedSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
void CMAConsignmentProductList::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CReportSection *rptDetail = NULL;
	CString szSQL, tmpStr, szDate, szReportName;
	int nIdx = 1;
	szReportName = _T("Reports\\HMS\\MA_BAOCAOVATUCHUAHOADON.RPT");
	if (!rpt.Init(szReportName))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("patientname")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("docno")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("objectname")));
		rs.GetValue(_T("approvedate"), tmpStr);
		rptDetail->SetValue(_T("5"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rs.GetValue(_T("orderdate"), tmpStr);
		rptDetail->SetValue(_T("6"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("partnername")));
		rs.MoveNext();
	}
	szDate = pMF->GetSysDate(); 
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szDate.Right(2),szDate.Mid(5,2),szDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}
 
void CMAConsignmentProductList::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr;
	CStringArray arrCol;
	int nIdx = 1, nRow = 0;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(5, 10);
	xls.SetColumnWidth(6, 40);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 7);
	xls.SetCellMergedColumns(0, 3, 7);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	xls.SetCellText(0, 2, _T("\x42\xC1O \x43\xC1O V\x1EACT T\x1AF K\xDD G\x1EECI \x43H\x1AF\x41 \x43\xD3 H\xD3\x41 \x110\x1A0N"), 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098);
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("V\x1EADt t\x1B0"));
	arrCol.Add(_T("\x110\x1A1n v\x1ECB t\xEDnh"));
	arrCol.Add(_T("S\x1ED1 l\x1B0\x1EE3ng"));
	arrCol.Add(_T("Ng\xE0y \x64uy\x1EC7t"));
	arrCol.Add(_T("Ng\xE0y s\x1EED \x64\x1EE5ng"));
	arrCol.Add(_T("\x43\xF4ng ty \x63ung \x63\x1EA5p"));
	for (int i = 0; i < arrCol.GetCount(); i++)
		xls.SetCellText(i, 4, arrCol.GetAt(i), 4098, true);
	nRow = 5;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_INTEGER);
		xls.SetCellText(1, nRow, rs.GetValue(_T("productname")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("uoname")), 4098);
		xls.SetCellText(3, nRow, rs.GetValue(_T("qtyorder")), FMT_NUMBER1);
		rs.GetValue(_T("approvedate"), tmpStr);
		xls.SetCellText(4, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), 4098);
		rs.GetValue(_T("orderdate"), tmpStr);
		xls.SetCellText(5, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), 4098);
		xls.SetCellText(6, nRow, rs.GetValue(_T("partnername")), FMT_TEXT); 
		nRow++;
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\Bao cao vat tu ky gui chua co hoa don.xls"));
}
 
CString CMAConsignmentProductList::GetQueryString(){
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND hpo_approvedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szWhere.AppendFormat(_T(" AND NVL(mt_status, 'X') <> 'A'"));
	if (!m_szSupplierKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND product_bpartnerid = '%s'"), m_szSupplierKey);
	szSQL.Format(_T(" SELECT    product_name productname, ") \
				_T("			product_uomname uomname,") \
				_T("			sum(hpol_qtyissue) qtyorder,") \
				_T("			hpo_approvedate approvedate,") \
				_T("			hpo_orderdate orderdate,") \
				_T("			get_partnername(product_bpartnerid) partnername") \
				_T(" FROM      hms_ipharmaorder ") \
				_T(" LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T(" LEFT JOIN purchase_orderline2 ON (pol_orderid = hpol_orderid AND hpol_product_id = pol_product_id)") \
				_T(" LEFT JOIN m_transaction ON (mt_transaction_id = pol_transaction_id)") \
				_T(" LEFT JOIN m_productitem_view ON ( hpol_product_item_id = product_item_id ) ") \
				_T(" WHERE     hpo_storage_id = 12 ") \
				_T("       AND hpo_orderstatus = 'A' ") \
				_T("		%s") \
				_T(" GROUP BY product_name, product_uomname, hpo_approvedate, hpo_orderdate, product_bpartnerid") \
				_T(" ORDER BY productname, approvedate, partnername"), szWhere);
	return szSQL;
}