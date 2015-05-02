#include "stdafx.h"
#include "PMDetailPurchaseOrder.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMDetailPurchaseOrder *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMDetailPurchaseOrder *)pWnd)->OnToDateCheckValue();
} 
static void _OnSupplierSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMDetailPurchaseOrder* )pWnd)->OnSupplierSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSupplierSelendokFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnSupplierSelendok();
}
/*static void _OnSupplierSetfocusFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnSupplierKillfocus();
}*/
/*static void _OnSupplierKillfocusFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnSupplierKillfocus();
}*/
static long _OnSupplierLoadDataFnc(CWnd *pWnd){
	return ((CPMDetailPurchaseOrder *)pWnd)->OnSupplierLoadData();
}
/*static void _OnSupplierAddNewFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnSupplierAddNew();
}*/
static void _OnTypeSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMDetailPurchaseOrder* )pWnd)->OnTypeSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnTypeSelendokFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnTypeSelendok();
}
/*static void _OnTypeSetfocusFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnTypeKillfocus();
}*/
/*static void _OnTypeKillfocusFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnTypeKillfocus();
}*/
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CPMDetailPurchaseOrder *)pWnd)->OnTypeLoadData();
}
/*static void _OnTypeAddNewFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnTypeAddNew();
}*/
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMDetailPurchaseOrder* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMDetailPurchaseOrder *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnStorageAddNew();
}*/
static void _OnItemSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMDetailPurchaseOrder* )pWnd)->OnItemSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnItemSelendokFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnItemSelendok();
}
/*static void _OnItemSetfocusFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnItemKillfocus();
}*/
/*static void _OnItemKillfocusFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnItemKillfocus();
}*/
static long _OnItemLoadDataFnc(CWnd *pWnd){
	return ((CPMDetailPurchaseOrder *)pWnd)->OnItemLoadData();
}
/*static void _OnItemAddNewFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder *)pWnd)->OnItemAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMDetailPurchaseOrder *pVw = (CPMDetailPurchaseOrder *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMDetailPurchaseOrder *pVw = (CPMDetailPurchaseOrder *)pWnd;
	pVw->OnExportSelect();
} 
CPMDetailPurchaseOrder::CPMDetailPurchaseOrder(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CPMDetailPurchaseOrder::~CPMDetailPurchaseOrder(){
}
void CPMDetailPurchaseOrder::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 150);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 285, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 290, 30, 370, 55);
	m_wndToDate.Create(this,375, 30, 570, 55); 
	m_wndSupplierLabel.Create(this, _T("Supplier"), 10, 60, 90, 85);
	m_wndSupplier.Create(this,95, 60, 570, 85); 
	m_wndTypeLabel.Create(this, _T("Type"), 10, 90, 90, 115);
	m_wndType.Create(this,95, 90, 285, 115); 
	m_wndStorageLabel.Create(this, _T("Storage"), 290, 90, 370, 115);
	m_wndStorage.Create(this,375, 90, 570, 115); 
	m_wndItemLabel.Create(this, _T("Item"), 10, 120, 90, 145);
	m_wndItem.Create(this,95, 120, 570, 145); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 350, 155, 460, 180);
	m_wndExport.Create(this, _T("&Export"), 465, 155, 575, 180);

}
void CPMDetailPurchaseOrder::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndSupplier.SetCheckValue(true);
	m_wndSupplier.LimitText(1024);
	m_wndType.SetCheckValue(true);
	m_wndType.LimitText(1024);
	m_wndStorage.SetCheckValue(true);
	m_wndStorage.LimitText(35);


	m_wndSupplier.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndSupplier.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);


	m_wndType.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndType.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);


	m_wndStorage.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndStorage.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

	m_wndItem.Format(3, 2);
	m_wndItem.InsertColumn(0, _T(""), CFMT_TEXT, 0);
	m_wndItem.InsertColumn(1, _T("ID"), CFMT_TEXT, 80);
	m_wndItem.InsertColumn(2, _T("Name"), CFMT_TEXT, 350);

}

void CPMDetailPurchaseOrder::OnSetWindowEvents(){
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
	m_wndType.SetEvent(WE_SELENDOK, _OnTypeSelendokFnc);
	//m_wndType.SetEvent(WE_SETFOCUS, _OnTypeSetfocusFnc);
	//m_wndType.SetEvent(WE_KILLFOCUS, _OnTypeKillfocusFnc);
	m_wndType.SetEvent(WE_SELCHANGE, _OnTypeSelectChangeFnc);
	m_wndType.SetEvent(WE_LOADDATA, _OnTypeLoadDataFnc);
	//m_wndType.SetEvent(WE_ADDNEW, _OnTypeAddNewFnc);
	m_wndStorage.SetEvent(WE_SELENDOK, _OnStorageSelendokFnc);
	//m_wndStorage.SetEvent(WE_SETFOCUS, _OnStorageSetfocusFnc);
	//m_wndStorage.SetEvent(WE_KILLFOCUS, _OnStorageKillfocusFnc);
	m_wndStorage.SetEvent(WE_SELCHANGE, _OnStorageSelectChangeFnc);
	m_wndStorage.SetEvent(WE_LOADDATA, _OnStorageLoadDataFnc);
	//m_wndStorage.SetEvent(WE_ADDNEW, _OnStorageAddNewFnc);
	m_wndItem.SetEvent(WE_SELENDOK, _OnItemSelendokFnc);
	//m_wndItem.SetEvent(WE_SETFOCUS, _OnItemSetfocusFnc);
	//m_wndItem.SetEvent(WE_KILLFOCUS, _OnItemKillfocusFnc);
	m_wndItem.SetEvent(WE_SELCHANGE, _OnItemSelectChangeFnc);
	m_wndItem.SetEvent(WE_LOADDATA, _OnItemLoadDataFnc);
	//m_wndItem.SetEvent(WE_ADDNEW, _OnItemAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CPMDetailPurchaseOrder::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szSupplierKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_TextEx(pDX, m_wndStorage.GetDlgCtrlID(), m_szStorageKey);
	DDX_TextEx(pDX, m_wndItem.GetDlgCtrlID(), m_szItemKey);

}
void CPMDetailPurchaseOrder::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szSupplierKey.Empty();
	m_szTypeKey.Empty();
	m_szStorageKey.Empty();

}

int CPMDetailPurchaseOrder::SetMode(int nMode){
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

/*void CPMDetailPurchaseOrder::OnFromDateChange(){
	
} */
/*void CPMDetailPurchaseOrder::OnFromDateSetfocus(){
	
} */
/*void CPMDetailPurchaseOrder::OnFromDateKillfocus(){
	
} */
int CPMDetailPurchaseOrder::OnFromDateCheckValue(){
	return 0;
}
 
/*void CPMDetailPurchaseOrder::OnToDateChange(){
	
} */
/*void CPMDetailPurchaseOrder::OnToDateSetfocus(){
	
} */
/*void CPMDetailPurchaseOrder::OnToDateKillfocus(){
	
} */
int CPMDetailPurchaseOrder::OnToDateCheckValue(){
	return 0;
}
 
void CPMDetailPurchaseOrder::OnSupplierSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
void CPMDetailPurchaseOrder::OnSupplierSelendok(){
	 
}

/*void CPMDetailPurchaseOrder::OnSupplierSetfocus(){
	
}*/
/*void CPMDetailPurchaseOrder::OnSupplierKillfocus(){
	
}*/
long CPMDetailPurchaseOrder::OnSupplierLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadPartnerList(&m_wndSupplier, m_szSupplierKey);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndSupplier.IsSearchKey() && !m_szSupplierKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szSupplierKey
};
	m_wndSupplier.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndSupplier.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}

/*void CPMDetailPurchaseOrder::OnSupplierAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMDetailPurchaseOrder::OnTypeSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
void CPMDetailPurchaseOrder::OnTypeSelendok(){
	 
}

/*void CPMDetailPurchaseOrder::OnTypeSetfocus(){
	
}*/
/*void CPMDetailPurchaseOrder::OnTypeKillfocus(){
	
}*/
long CPMDetailPurchaseOrder::OnTypeLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadProductTypeList(&m_wndType, m_szTypeKey);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndType.IsSearchKey() && !m_szTypeKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szTypeKey
};
	m_wndType.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndType.AddItems( o
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}

/*void CPMDetailPurchaseOrder::OnTypeAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMDetailPurchaseOrder::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
void CPMDetailPurchaseOrder::OnStorageSelendok(){
}

/*void CPMDetailPurchaseOrder::OnStorageSetfocus(){
	
}*/
/*void CPMDetailPurchaseOrder::OnStorageKillfocus(){
	
}*/
long CPMDetailPurchaseOrder::OnStorageLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageList(&m_wndStorage, m_szStorageKey);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStorage.IsSearchKey() && !m_szStorageKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szStorageKey
};
	m_wndStorage.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStorage.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}

/*void CPMDetailPurchaseOrder::OnStorageAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMDetailPurchaseOrder::OnItemSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMDetailPurchaseOrder::OnItemSelendok(){
	 
}
/*void CPMDetailPurchaseOrder::OnItemSetfocus(){
	
}*/
/*void CPMDetailPurchaseOrder::OnItemKillfocus(){
	
}*/
long CPMDetailPurchaseOrder::OnItemLoadData(){	
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	UpdateData(true);

	int nCount = 0;

	szWhere.AppendFormat(_T(" AND mp_org_id = '%s'"), pMF->GetModuleID());
	szSQL.Format(_T(" SELECT  mp_code AS code,") \
			_T("   mp_name AS name,") \
			_T("   mp_product_id as id") \
			_T(" FROM m_product ") \
			_T(" WHERE  1=1 %s") \
			_T(" ORDER  BY name "), szWhere);


	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndItem.AddItems(
			rs.GetValue(_T("id")),
			rs.GetValue(_T("code")), 
			rs.GetValue(_T("Name")), 
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPMDetailPurchaseOrder::OnItemAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMDetailPurchaseOrder::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	/*Declaration Section*/
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL;
	double nTmp = 0;
	int nIdx = 1;
	CString szSQL, tmpStr, szReportName, szDate;
	long double nTotal = 0;
	szReportName = _T("Reports\\HMS\\PMS_BANGKENHAPKHOCHITIET.RPT");
	if (!rpt.Init(szReportName))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
		CDate::Convert(m_szToDate), yyyymmdd, ddmmyyyy);
		rptHeader->SetValue(_T("ReportDate"), tmpStr);
		tmpStr = _T("T\x1EA5t \x63\x1EA3");
		if (!m_szStorageKey.IsEmpty())
			rptHeader->SetValue(_T("StockName"), m_wndStorage.GetCurrent(1));
		else
			rptHeader->SetValue(_T("StockName"), tmpStr);
		if (!m_szItemKey.IsEmpty())
			rptHeader->SetValue(_T("ItemName"), m_wndItem.GetCurrent(2));
		else
			rptHeader->SetValue(_T("ItemName"), tmpStr);
	}
	while (!rs.IsEOF())
	{
		/*Group Seperation*/
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("name")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("partner_name")));
			rs.GetValue(_T("import_date"), tmpStr);
			rptDetail->SetValue(_T("4"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("order_no")));
			rptDetail->SetValue(_T("6"), rs.GetValue(_T("qty")));
			rptDetail->SetValue(_T("7"), rs.GetValue(_T("unit_price")));
			rs.GetValue(_T("amount"), nTmp);
			nTotal+= nTmp;
			rptDetail->SetValue(_T("8"), double2str(nTmp));
		}
		rs.MoveNext();
	}
	if (nTotal > 0)
	{
		tmpStr.Format(_T("%f"), nTotal);
		rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	}
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
	
}
 
void CPMDetailPurchaseOrder::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CExcel xls;
	CString szSQL, tmpStr;
	double nTmp = 0;
	long double nTotal = 0;
	int nIdx = 1, nRow = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 8);
	xls.SetColumnWidth(1, 20);
	xls.SetColumnWidth(2, 20);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(6, 10);
	xls.SetColumnWidth(7, 15);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 7);
	xls.SetCellMergedColumns(0, 3, 7);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr = _T("\x42\x1EA2NG K\xCA NH\x1EACP KHO \x43HI TI\x1EBET");
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	xls.SetCellText(0, 4, _T("STT"), 4098, true);
	xls.SetCellText(1, 4, _T("T\xEAn h\xE0ng"), 4098, true);
	xls.SetCellText(2, 4, _T("T\xEAn \x63\xF4ng ty"), 4098, true);
	xls.SetCellText(3, 4, _T("Ng\xE0y nh\x1EADp"), 4098, true);
	xls.SetCellText(4, 4, _T("S\x1ED1 nh\x1EADp l\x1EC7nh"), 4098, true);
	xls.SetCellText(5, 4, _T("SL"), 4098, true);
	xls.SetCellText(6, 4, _T("\x110\x1A1n gi\xE1"), 4098, true);
	xls.SetCellText(7, 4, _T("Th\xE0nh ti\x1EC1n"), 4098, true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("name")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("partner_name")), FMT_TEXT);
		tmpStr = CDate::Convert(rs.GetValue(_T("import_date")), yyyymmdd, ddmmyyyy);
		xls.SetCellText(3, nRow, tmpStr, FMT_TEXT);
		xls.SetCellText(4, nRow, rs.GetValue(_T("order_no")), FMT_TEXT);
		xls.SetCellText(5, nRow, rs.GetValue(_T("qty")), FMT_NUMBER1);
		xls.SetCellText(6, nRow, rs.GetValue(_T("unit_price")), FMT_NUMBER1);
		rs.GetValue(_T("amount"), nTmp);
		nTotal += nTmp;
		xls.SetCellText(7, nRow, double2str(nTmp), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nTotal > 0)
	{
		xls.SetCellText(6, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
		tmpStr.Format(_T("%f"), nTotal);
		xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\AutoExcel.xls"));
	
}
 
CString CPMDetailPurchaseOrder::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND po_status = 'A' AND po_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
		, m_szFromDate, m_szToDate);
	if (!m_szStorageKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND po_storage_id IN (%s)"), m_szStorageKey);
	if (!m_szSupplierKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND po_partner_id = '%s'"), m_szSupplierKey);
	szWhere.AppendFormat(_T(" AND po_org_id = '%s'"), pMF->GetModuleID());
	if (!m_szTypeKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND mp_producttype = '%s'"), m_szTypeKey);
	if (!m_szItemKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND pol_product_id = %s"), m_szItemKey);
	szSQL.Format(_T(" SELECT po_approveddate import_date,") \
				_T("   get_partnername(po_partner_id) partner_name,") \
				_T("   po_orderno order_no,") \
				_T("   mp_code AS id,") \
				_T("   mp_name AS name,") \
				_T("   mpi_unitprice AS unit_price,") \
				_T("   sum(pol_qtyorder) AS qty,") \
				_T("   sum(pol_qtyorder*mpi_unitprice) AS amount") \
				_T(" FROM   purchase_order") \
				_T(" LEFT JOIN purchase_orderline ON (po_purchase_order_id = pol_purchase_order_id)") \
				_T(" LEFT JOIN m_product_item ON(mpi_product_item_id=pol_product_item_id)") \
				_T(" LEFT JOIN m_product ON(mp_product_id=mpi_product_id)") \
				_T(" WHERE 1=1 %s") \
				_T(" GROUP BY mp_code, mp_name, mpi_unitprice, po_partner_id, po_approveddate, po_orderno") \
				_T(" ORDER BY name, partner_name"), szWhere);
	return szSQL;
}