#include "stdafx.h"
#include "PMDrugListwithoutInvoice.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include ".\pmdruglistwithoutinvoice.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMDrugListwithoutInvoice *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMDrugListwithoutInvoice *)pWnd)->OnToDateCheckValue();
} 
static void _OnSupplierSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMDrugListwithoutInvoice* )pWnd)->OnSupplierSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSupplierSelendokFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnSupplierSelendok();
}
/*static void _OnSupplierSetfocusFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnSupplierKillfocus();
}*/
/*static void _OnSupplierKillfocusFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnSupplierKillfocus();
}*/
static long _OnSupplierLoadDataFnc(CWnd *pWnd){
	return ((CPMDrugListwithoutInvoice *)pWnd)->OnSupplierLoadData();
}
/*static void _OnSupplierAddNewFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnSupplierAddNew();
}*/
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMDrugListwithoutInvoice* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CPMDrugListwithoutInvoice *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnStockAddNew();
}*/
static void _OnTypeSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMDrugListwithoutInvoice* )pWnd)->OnTypeSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnTypeSelendokFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnTypeSelendok();
}
/*static void _OnTypeSetfocusFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnTypeKillfocus();
}*/
/*static void _OnTypeKillfocusFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnTypeKillfocus();
}*/
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CPMDrugListwithoutInvoice *)pWnd)->OnTypeLoadData();
}
/*static void _OnTypeAddNewFnc(CWnd *pWnd){
	((CPMDrugListwithoutInvoice *)pWnd)->OnTypeAddNew();
}*/

static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMDrugListwithoutInvoice *pVw = (CPMDrugListwithoutInvoice *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CPMDrugListwithoutInvoice *pVw = (CPMDrugListwithoutInvoice *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMDrugListwithoutInvoice *pVw = (CPMDrugListwithoutInvoice *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CPMDrugListwithoutInvoice *pVw = (CPMDrugListwithoutInvoice *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddPMSPurchaseOrderReportDialogFnc(CWnd *pWnd){
	 return ((CPMDrugListwithoutInvoice*)pWnd)->OnAddPMSPurchaseOrderReportDialog();
} 
static int _OnEditPMSPurchaseOrderReportDialogFnc(CWnd *pWnd){
	 return ((CPMDrugListwithoutInvoice*)pWnd)->OnEditPMSPurchaseOrderReportDialog();
} 
static int _OnDeletePMSPurchaseOrderReportDialogFnc(CWnd *pWnd){
	 return ((CPMDrugListwithoutInvoice*)pWnd)->OnDeletePMSPurchaseOrderReportDialog();
} 
static int _OnSavePMSPurchaseOrderReportDialogFnc(CWnd *pWnd){
	 return ((CPMDrugListwithoutInvoice*)pWnd)->OnSavePMSPurchaseOrderReportDialog();
} 
static int _OnCancelPMSPurchaseOrderReportDialogFnc(CWnd *pWnd){
	 return ((CPMDrugListwithoutInvoice*)pWnd)->OnCancelPMSPurchaseOrderReportDialog();
} 
CPMDrugListwithoutInvoice::CPMDrugListwithoutInvoice(CWnd *pParent)
{
	m_nDlgWidth = 585;
	m_nDlgHeight = 185;
//	m_szTitle.Format(_T("Tong hop bien ban kiem ke phieu nhap kho"));
	SetDefaultValues();
}
CPMDrugListwithoutInvoice::~CPMDrugListwithoutInvoice(){

}
void CPMDrugListwithoutInvoice::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 150);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 110, 55);
	m_wndFromDate.Create(this,115, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 395, 55);
	m_wndToDate.Create(this,400, 30, 570, 55); 
	m_wndSupplierLabel.Create(this, _T("Supplier"), 10, 60, 110, 85);
	m_wndSupplier.Create(this,115, 60, 570, 85); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 90, 110, 115);
	m_wndStock.Create(this,115, 90, 570, 115); 
	m_wndStock.SetCheckBox(true);
	m_wndTypeLabel.Create(this, _T("Type"), 10, 120, 110, 145);
	m_wndType.Create(this,115, 120, 570, 145); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 385, 155, 495, 180);
	m_wndExport.Create(this, _T("&Export"), 500, 155, 575, 180);

}
void CPMDrugListwithoutInvoice::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
//	m_wndSupplier.SetCheckValue(true);
	m_wndSupplier.LimitText(35);
//	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);


	m_wndSupplier.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndSupplier.InsertColumn(1, _T("Description"), CFMT_TEXT, 320);


	m_wndStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStock.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndType.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndType.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);


}
void CPMDrugListwithoutInvoice::OnSetWindowEvents(){
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
	m_wndStock.SetEvent(WE_SELENDOK, _OnStockSelendokFnc);
	//m_wndStock.SetEvent(WE_SETFOCUS, _OnStockSetfocusFnc);
	//m_wndStock.SetEvent(WE_KILLFOCUS, _OnStockKillfocusFnc);
	m_wndStock.SetEvent(WE_SELCHANGE, _OnStockSelectChangeFnc);
	m_wndStock.SetEvent(WE_LOADDATA, _OnStockLoadDataFnc);
	//m_wndStock.SetEvent(WE_ADDNEW, _OnStockAddNewFnc);

	m_wndType.SetEvent(WE_SELENDOK, _OnTypeSelendokFnc);
	//m_wndType.SetEvent(WE_SETFOCUS, _OnTypeSetfocusFnc);
	//m_wndType.SetEvent(WE_KILLFOCUS, _OnTypeKillfocusFnc);
	m_wndType.SetEvent(WE_SELCHANGE, _OnTypeSelectChangeFnc);
	m_wndType.SetEvent(WE_LOADDATA, _OnTypeLoadDataFnc);
	//m_wndType.SetEvent(WE_ADDNEW, _OnTypeAddNewFnc);


	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();		
	m_szFromDate =  pMF->GetSysDate()+ _T("00:00");	
	m_szToDate =pMF->GetSysDate() + _T("23:59");
	SetMode(VM_EDIT);

}
void CPMDrugListwithoutInvoice::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szSupplierKey);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
}
void CPMDrugListwithoutInvoice::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CPMDrugListwithoutInvoice::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CPMDrugListwithoutInvoice::SetDefaultValues(){
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szSupplierKey.Empty();
	m_szStockKey.Empty();
}
int CPMDrugListwithoutInvoice::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, 1, 2, 3, -1); 
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
/*void CPMDrugListwithoutInvoice::OnFromDateChange(){
	
} */
/*void CPMDrugListwithoutInvoice::OnFromDateSetfocus(){
	
} */
/*void CPMDrugListwithoutInvoice::OnFromDateKillfocus(){
	
} */
int CPMDrugListwithoutInvoice::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMDrugListwithoutInvoice::OnToDateChange(){
	
} */
/*void CPMDrugListwithoutInvoice::OnToDateSetfocus(){
	
} */
/*void CPMDrugListwithoutInvoice::OnToDateKillfocus(){
	
} */
int CPMDrugListwithoutInvoice::OnToDateCheckValue(){
	return 0;
} 
void CPMDrugListwithoutInvoice::OnSupplierSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMDrugListwithoutInvoice::OnSupplierSelendok(){
	 
}
/*void CPMDrugListwithoutInvoice::OnSupplierSetfocus(){
	
}*/
/*void CPMDrugListwithoutInvoice::OnSupplierKillfocus(){
	
}*/
long CPMDrugListwithoutInvoice::OnSupplierLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndSupplier.IsSearchKey() && !m_szSupplierKey.IsEmpty()){
		szWhere.Format(_T(" AND adp_partner_id='%s' "), m_szSupplierKey);
	};
	m_wndSupplier.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT adp_partner_id as id, adp_name as name FROM ad_partner WHERE 1=1 %s ORDER BY adp_partner_id"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndSupplier.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CPMDrugListwithoutInvoice::OnSupplierAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMDrugListwithoutInvoice::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMDrugListwithoutInvoice::OnStockSelendok(){
	 
}
/*void CPMDrugListwithoutInvoice::OnStockSetfocus(){
	
}*/
/*void CPMDrugListwithoutInvoice::OnStockKillfocus(){
	
}*/
long CPMDrugListwithoutInvoice::OnStockLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageList(&m_wndStock, m_szStockKey);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStock.IsSearchKey() && ToInt(m_szStockKey) > 0){
		szWhere.Format(_T(" AND msl_storage_id=%d"), ToInt(m_szStockKey));
	};
	m_wndStock.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT msl_storage_id as id, msl_name as name FROM m_storagelist WHERE msl_type='A' %s ORDER BY msl_storage_id"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStock.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CPMDrugListwithoutInvoice::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */

void CPMDrugListwithoutInvoice::OnTypeSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMDrugListwithoutInvoice::OnTypeSelendok(){
	 
}
/*void CPMDrugListwithoutInvoice::OnTypeSetfocus(){
	
}*/
/*void CPMDrugListwithoutInvoice::OnTypeKillfocus(){
	
}*/
long CPMDrugListwithoutInvoice::OnTypeLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndType.IsSearchKey() && !m_szTypeKey.IsEmpty()){
		szWhere.Format(_T(" and mpt_product_type_id='%s' "), m_szTypeKey);
	};
	//szWhere.AppendFormat(_T(" AND hfg_moduleid = '%s'"), pMF->GetModuleID());
	m_wndType.DeleteAllItems(); 
	int nCount = 0;

	szSQL.Format(_T("SELECT mpt_product_type_id as id, mpt_name as name FROM m_product_type WHERe substr(mpt_product_type_id,1, 1)='A' %s ORDER BY mpt_product_type_id "), szWhere);

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndType.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPMDrugListwithoutInvoice::OnTypeAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */

void CPMDrugListwithoutInvoice::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);	

} 
void CPMDrugListwithoutInvoice::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
	UpdateData(true);
	CReport rpt;
	CString tmpStr, szSQL;
	CString szOldLine, szNewLine, szAmount, szName, szOffice, szUserName, szRptName;
	CRecord rs(&pMF->m_db);
	CRecord mrs(&pMF->m_db);
	CReportSection* rptDetail;
	long double nSubAmount=0, nVatAmount=0, nTotalAmount=0, nTotalPayment=0;
	double cost=0.0;
	long double grpCost[15], ttCost[15];
	CString szInvoiceNo;
	CString szField;
	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		//QueryOpener(szSQL);
		return;
	}
	for (int i=0;i<15; i++)
	{
		grpCost[i]=0;
		ttCost[i]=0;
	}
	szRptName.Format(_T("Reports/HMS/PMS_DANHSACHPHIEUNHAPKHOCHUACOHOADON.RPT"));
	if(!rpt.Init(szRptName,TRUE) )
		return;

	rpt.GetReportHeader()->SetValue(_T("HEAlTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	//Page Header
	//Report Detail
	int nIndex=1;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("invoiceno"), szNewLine);
		if(szNewLine != szOldLine && !szNewLine.IsEmpty())
		{					
			if(grpCost[12] > 0)
			{
				TranslateString(_T("Total Amount"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("TotalGroup"), szAmount);												
				for(i=10;i<15; i++){
					szField.Format(_T("s%d"), i);
					FormatCurrency(grpCost[i], tmpStr);
					rptDetail->SetValue(szField, tmpStr);
					ttCost[i] += grpCost[i];
					grpCost[i]=0;
				}
			}

			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rptDetail->SetValue(_T("InvoiceNo"),szNewLine );	
			rs.GetValue(_T("Impdate"), tmpStr);
			rptDetail->SetValue(_T("InvoiceDate"),CDate::Convert(tmpStr,yyyymmdd,ddmmyyyy) );					
			rs.GetValue(_T("suppliers"), tmpStr);
			rptDetail->SetValue(_T("Supplier"),tmpStr );	
			szInvoiceNo.AppendFormat(_T("%s, "),szNewLine);;
			szOldLine = szNewLine;
			nIndex=1;
		}

		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("name"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("id"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("qty"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		//rptDetail->SetValue(_T("8"), tmpStr);
		//rs.GetValue(_T("unitprice"), cost);	

		rs.GetValue(_T("expdate"), tmpStr);
		rptDetail->SetValue(_T("5"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

		rs.GetValue(_T("lotno"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("unitprice"), cost);
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("subamount"), cost);
		grpCost[10] += cost;								
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.GetValue(_T("Vat"), tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.GetValue(_T("amountvat"), cost);
		grpCost[12] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);	

		rs.GetValue(_T("costamount"), cost);
		grpCost[13] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);

		rs.GetValue(_T("country"), tmpStr);
		rptDetail->SetValue(_T("14"), tmpStr);

		rs.GetValue(_T("orgcountry"), tmpStr);				
		rptDetail->SetValue(_T("16"), tmpStr);

		rs.GetValue(_T("barcode"), tmpStr);
		rptDetail->SetValue(_T("15"), tmpStr);			

		rs.MoveNext();
	}
	//Page Footer
	//Report Footer
	for(i=10;i<15;i++)
	{
		ttCost[i]+=grpCost[i];
	}

	if(grpCost[12] > 0)
	{
		TranslateString(_T("Total Amount"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), szAmount);									
		for(int i=10;i<15; i++)
		{
			szField.Format(_T("s%d"), i);
			FormatCurrency(grpCost[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);
		}						
	}
	
	if(ttCost[12] > 0)
	{
		TranslateString(_T("Total Sum"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), szAmount);			
		for(int i=10;i<15; i++)
		{
			szField.Format(_T("s%d"), i);
			FormatCurrency(ttCost[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);
		}
	}
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), szInvoiceNo);
	szAmount.Format(_T("%.2f"), ttCost[13]);
	MoneyToString(szAmount, tmpStr);	
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr + _T(" \x111\x1ED3ng."));
	CString szSysDate, szDate;
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	
	rpt.PrintPreview();
	EndWaitCursor();

} 
void CPMDrugListwithoutInvoice::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 	
	UpdateData(true);
	CString tmpStr, szSQL, szTemp;
	CString szOldLine, szNewLine, szAmount, szName, szOffice, szUserName;
	CRecord rs(&pMF->m_db);
	CRecord mrs(&pMF->m_db);
	
	long double nSubAmount=0, nVatAmount=0, nTotalAmount=0, nTotalPayment=0;
	double cost=0.0;
	long double grpCost[15], ttCost[15];
	//int i;
	CString szInvoiceNo;
	CString szField;
	
	for (int i = 0; i < 15; i++)
	{
		grpCost[i]=0;
		ttCost[i]=0;
	}
	
	BeginWaitCursor();
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	CExcel xls;	
	int nRow = 6;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	CellFormat df(&xls), cf(&xls), hf(&xls);
	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER);
	cf.SetFontName(_T("Segoe UI"));

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(2, 25);
	xls.SetColumnWidth(3, 6);
	xls.SetColumnWidth(4, 11);
	xls.SetColumnWidth(5, 7);
	xls.SetColumnWidth(6, 14);
	xls.SetColumnWidth(7, 20);
	xls.SetColumnWidth(8, 11);
	xls.SetColumnWidth(9, 10);
	xls.SetColumnWidth(10, 12);
	xls.SetColumnWidth(11, 15);
	xls.SetColumnWidth(12, 15);
	
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellText(0, 0, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(0, 2, 13);
	TranslateString(_T("Purchase Order Control Record"), szTemp);
	StringUpper(szTemp, tmpStr);
	xls.SetCellText(0, 2, tmpStr, FMT_TEXT | FMT_CENTER, true, 12);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellMergedColumns(0, 3, 13);
	xls.SetCellText(0, 3, tmpStr, &df);
	nRow = 3;
	if (!m_szStockKey.IsEmpty())
	{
		nRow++;
		xls.SetCellMergedColumns(0, nRow, 13);
		tmpStr.Format(_T("Kho: %s"), m_szStock);
		xls.SetCellText(0, nRow, tmpStr, &cf);	
	}
	nRow++;
	xls.SetCellText(0, nRow, _T("STT"), &hf);	
	TranslateString(_T("Drug Name"), tmpStr);
	xls.SetCellText(1, nRow, tmpStr, &hf);	
	TranslateString(_T("Unit"), tmpStr);
	xls.SetCellText(2, nRow, tmpStr, &hf);	
	TranslateString(_T("Qty"), tmpStr);
	xls.SetCellText(3, nRow, tmpStr, &hf);	
	TranslateString(_T("Price"), tmpStr);
	xls.SetCellText(4, nRow, tmpStr, &hf);
	//TranslateString(_T("Sub Total"), tmpStr);
	//xls.SetCellText(5, nRow, tmpStr, &hf);	
	TranslateString(_T("VAT(%)"), tmpStr);
	xls.SetCellText(5, nRow, tmpStr, &hf);
	TranslateString(_T("S\x1ED1 ti\x1EC1n VAT(%)"), tmpStr);
	xls.SetCellText(6, nRow, tmpStr, &hf);
	TranslateString(_T("T\x1ED5ng ti\x1EC1n"), tmpStr);
	xls.SetCellText(7, nRow, tmpStr, &hf);
	TranslateString(_T("Loot"), tmpStr);
	xls.SetCellText(8, nRow, tmpStr, &hf);
	TranslateString(_T("Exp Date"), tmpStr);
	xls.SetCellText(9, nRow, tmpStr, &hf);
	TranslateString(_T("Country"), tmpStr);
	xls.SetCellText(10, nRow, tmpStr, &hf);
	TranslateString(_T("H\xE3ng SX"), tmpStr);
	xls.SetCellText(11, nRow, tmpStr, &hf);
	TranslateString(_T("S\x1ED1 \x110K"), tmpStr);
	xls.SetCellText(12, nRow, tmpStr, &hf);	
	int nIndex = 1;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("invoiceno"), szNewLine);
		if(szNewLine != szOldLine && !szNewLine.IsEmpty())
		{					
			if(grpCost[7] > 0)
			{
				nRow++;
				xls.SetCellMergedColumns(0, nRow, 6);
				TranslateString(_T("C\x1ED9ng:"), tmpStr);
				xls.SetCellText(0, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
				for (int i = 6; i < 12; i++)
				{
				
					tmpStr.Format(_T("%.2f"), grpCost[i]);
					xls.SetCellText(i, nRow, tmpStr, FMT_NUMBER1, true);
					ttCost[i] += grpCost[i];
					grpCost[i] = 0;
				}
			}

			nRow++;
			/*xls.SetCellMergedColumns(0,nRow,2);
			xls.SetCellMergedColumns(2,nRow,3);
			xls.SetCellMergedColumns(5,nRow,5);			*/		
			xls.SetCellText(0, nRow, _T("S\x1ED1 HD:") + szNewLine, FMT_TEXT, true);	
			rs.GetValue(_T("Impdate"), tmpStr);
			xls.SetCellText(2,nRow, _T("Ng\xE0y nh\x1EADp: ")+ CDate::Convert(tmpStr,yyyymmdd,ddmmyyyy), FMT_TEXT, true);
			rs.GetValue(_T("suppliers"), tmpStr);
			xls.SetCellText(5, nRow, _T("Nh\xE0 \x63ung \x63\x1EA5p:") + tmpStr, FMT_TEXT, true);					
			szInvoiceNo.AppendFormat(_T("%s, "), szNewLine);
			szOldLine = szNewLine;
			nIndex = 1;
		}

		nRow++;
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("name"), tmpStr);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("unit"), tmpStr);
		xls.SetCellText(2, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("qty"), tmpStr);
		xls.SetCellText(3, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("unitprice"), cost);				
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);

		/*rs.GetValue(_T("subamount"), cost);
		grpCost[5] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(5,nRow,tmpStr, FMT_NUMBER1);*/
		rs.GetValue(_T("Vat"), cost);				
		tmpStr.Format(_T("%.2f"), cost);				
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("amountvat"), cost);
		grpCost[6] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("costamount"), cost);
		grpCost[7] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1);
					
		rs.GetValue(_T("lotno"), tmpStr);
		xls.SetCellText(8, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("expdate"), tmpStr);
		if (tmpStr.Left(4) != _T("1752"))
			xls.SetCellText(9, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), FMT_DATE);
		
		rs.GetValue(_T("country"), tmpStr);
		xls.SetCellText(10, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("orgcountry"), tmpStr);
		xls.SetCellText(11, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("barcode"), tmpStr);
		xls.SetCellText(12, nRow, tmpStr, FMT_TEXT);

		rs.MoveNext();
	}
	//Page Footer
	//Report Footer
	for(int i = 6; i < 12; i++)
	{
		ttCost[i] += grpCost[i];
	}

	if (grpCost[6] > 0)
	{
		nRow++;
		xls.SetCellMergedColumns(0, nRow, 6);
		TranslateString(_T("C\x1ED9ng:"), tmpStr);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);									
		for (int i = 6; i < 12; i++)
		{				
			tmpStr.Format(_T("%.2f"),grpCost[i]);
			xls.SetCellText(i, nRow, tmpStr, FMT_NUMBER1, true);
		}						
	}
	
	if(ttCost[6] > 0)
	{
		nRow++;
		xls.SetCellMergedColumns(0, nRow, 6);
		TranslateString(_T("Total Amount:"), tmpStr);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
		for (int i = 6; i < 12; i++)
		{
			tmpStr.Format(_T("%.2f"), ttCost[i]);
			xls.SetCellText(i, nRow, tmpStr, FMT_NUMBER1, true);		
		}						
	}
	xls.Save(_T("Exports\\Danh sach phieu nhap kho khong co hoa don.xls"));	
	EndWaitCursor();
}
	
void CPMDrugListwithoutInvoice::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
	
} 
int CPMDrugListwithoutInvoice::OnAddPMSPurchaseOrderReportDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CPMDrugListwithoutInvoice::OnEditPMSPurchaseOrderReportDialog(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CPMDrugListwithoutInvoice::OnDeletePMSPurchaseOrderReportDialog(){
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
 		OnCancelPMSPurchaseOrderReportDialog(); 
 	}
	return 0;
}
int CPMDrugListwithoutInvoice::OnSavePMSPurchaseOrderReportDialog(){
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
 		//OnPMSPurchaseOrderReportDialogListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CPMDrugListwithoutInvoice::OnCancelPMSPurchaseOrderReportDialog(){
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
int CPMDrugListwithoutInvoice::OnPMSPurchaseOrderReportDialogListLoadData(){
	return 0;
}

CString CPMDrugListwithoutInvoice::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szWhere;
	if(!m_szSupplierKey.IsEmpty())
		szWhere.AppendFormat(_T("and po_partner_id='%s' "), m_szSupplierKey);
	
	CString szStock;
	szStock.Empty();
	m_szStock.Empty();
	for (int i=0 ; i<= m_wndStock.GetCount(); i++)
	{
		if(m_wndStock.GetCheck(i))
		{
			m_wndStock.SetCurSel(i);
			if(!szStock.IsEmpty())
			{
				szStock += _T(",");
				m_szStock += _T(", ");
			}
			szStock.AppendFormat(_T("%s"), m_wndStock.GetCurrent(0));				
			m_szStock.AppendFormat(_T("%s"), m_wndStock.GetCurrent(1));				
		}		
	}
	
	CString szType;
	szType.Empty();
	for (int i=0 ; i<= m_wndType.GetCount(); i++)
	{
		if(m_wndType.GetCheck(i))
		{
			m_wndType.SetCurSel(i);
			if(!szType.IsEmpty())
			{
				szType += _T(", ");
			}
			szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));				
		}		
	}
	

	if(!szStock.IsEmpty())
		szWhere.AppendFormat(_T("and po_storage_id in(%s)"), szStock);

	if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and mp_producttype in(%s)"), szType);	
	szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	szSQL.Format(_T(" SELECT po_orderno                                           AS invoiceno,") \
	_T("   product_code                                                           AS id,") \
	_T("   get_taxrate(po_tax_id)                                                 AS Vat,") \
	_T("   po_orderdate                                                           AS Impdate,") \
	_T("   product_name                                                           AS name,") \
	_T("   get_uomname(product_uomid)                                             AS unit,") \
	_T("   product_expdate                                                        AS expdate,") \
	_T("   product_lotno                                                          AS lotno,") \
	_T("   product_unitprice                                                      AS unitprice,") \
	_T("   product_taxprice                                                       AS vatprice,") \
	_T("   sum(pol_qtyorder)                                                      AS qty,") \
	_T("   sum(pol_qtyorder)*product_unitprice                                    AS subamount,") \
	_T("   sum(pol_qtyorder)*product_taxprice-sum(pol_qtyorder)*product_unitprice AS amountvat,") \
	_T("   sum(pol_qtyorder)*product_taxprice                                     AS costamount,") \
	_T("   get_countryname(product_countryid)                                     AS country,") \
	_T("   get_manufacturename(product_manufactureid)                             AS orgcountry,") \
	_T("   get_partnername(po_partner_id)                                         AS suppliers") \
	_T(" FROM   purchase_order") \
	_T("        LEFT JOIN purchase_orderline ON(po_purchase_order_id=pol_purchase_order_id)") \
	_T("        LEFT JOIN m_productitem_view ON(product_item_id=pol_product_item_id)") \
	_T(" WHERE  po_orderdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
	_T("        AND po_status<>'O' AND po_invoiceno IS NULL %s") \
	_T(" GROUP  BY product_code,") \
	_T("           po_orderdate,") \
	_T("           po_tax_id,") \
	_T("           po_orderno,") \
	_T("           product_name,") \
	_T("           product_uomid,") \
	_T("           product_expdate,") \
	_T("           product_lotno,") \
	_T("           product_unitprice,") \
	_T("           product_taxprice,") \
	_T("           product_countryid,") \
	_T("           product_manufactureid,") \
	_T("           po_partner_id") \
	_T(" ORDER  BY suppliers,") \
	_T("           Impdate,") \
	_T("           invoiceno "), m_szFromDate, m_szToDate, szWhere);
	return szSQL;
}BEGIN_MESSAGE_MAP(CPMDrugListwithoutInvoice, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMDrugListwithoutInvoice::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
