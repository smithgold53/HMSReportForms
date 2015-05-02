#include "stdafx.h"
#include "MADetailPurchaseOrder.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CMADetailPurchaseOrder *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CMADetailPurchaseOrder *)pWnd)->OnToDateCheckValue();
} 
static void _OnSupplierSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CMADetailPurchaseOrder* )pWnd)->OnSupplierSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSupplierSelendokFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnSupplierSelendok();
}
/*static void _OnSupplierSetfocusFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnSupplierKillfocus();
}*/
/*static void _OnSupplierKillfocusFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnSupplierKillfocus();
}*/
static long _OnSupplierLoadDataFnc(CWnd *pWnd){
	return ((CMADetailPurchaseOrder *)pWnd)->OnSupplierLoadData();
}
/*static void _OnSupplierAddNewFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnSupplierAddNew();
}*/
static void _OnStockListSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CMADetailPurchaseOrder* )pWnd)->OnStockListSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockListSelendokFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnStockListSelendok();
}
/*static void _OnStockListSetfocusFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnStockListKillfocus();
}*/
/*static void _OnStockListKillfocusFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnStockListKillfocus();
}*/
static long _OnStockListLoadDataFnc(CWnd *pWnd){
	return ((CMADetailPurchaseOrder *)pWnd)->OnStockListLoadData();
}
/*static void _OnStockListAddNewFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnStockListAddNew();
}*/
static void _OnTypeSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CMADetailPurchaseOrder* )pWnd)->OnTypeSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnTypeSelendokFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnTypeSelendok();
}
/*static void _OnTypeSetfocusFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnTypeKillfocus();
}*/
/*static void _OnTypeKillfocusFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnTypeKillfocus();
}*/
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CMADetailPurchaseOrder *)pWnd)->OnTypeLoadData();
}
/*static void _OnTypeAddNewFnc(CWnd *pWnd){
	((CMADetailPurchaseOrder *)pWnd)->OnTypeAddNew();
}*/

static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CMADetailPurchaseOrder *pVw = (CMADetailPurchaseOrder *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CMADetailPurchaseOrder *pVw = (CMADetailPurchaseOrder *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CMADetailPurchaseOrder *pVw = (CMADetailPurchaseOrder *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CMADetailPurchaseOrder *pVw = (CMADetailPurchaseOrder *)pWnd;
	pVw->OnCloseSelect();
}
CMADetailPurchaseOrder::CMADetailPurchaseOrder(CWnd *pParent, int bDetail)
	{
	m_bDetail = bDetail;
	m_nDlgWidth = 585;
	m_nDlgHeight = 185;
//	m_szTitle.Format(_T("Tong hop bien ban kiem ke phieu nhap kho"));
	SetDefaultValues();
}
CMADetailPurchaseOrder::~CMADetailPurchaseOrder(){

}
void CMADetailPurchaseOrder::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 545);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 110, 55);
	m_wndFromDate.Create(this,115, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 395, 55);
	m_wndToDate.Create(this,400, 30, 570, 55); 
	m_wndSupplierLabel.Create(this, _T("Supplier"), 10, 60, 110, 85);
	m_wndSupplier.Create(this,115, 60, 570, 85); 
	m_wndTypeLabel.Create(this, _T("Type"), 10, 90, 110, 115);
	m_wndType.Create(this,115, 90, 570, 115);
	//m_wndGroupByDrug.Create(this, _T("Group by Drug"), 5, 550, 115, 575);
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 385, 550, 495, 575);
	m_wndExport.Create(this, _T("&Export"), 500, 550, 575, 575);
	m_wndStockList.Create(this,10, 120, 570, 540); 
	m_wndStockList.SetCheckBox(true);

}
void CMADetailPurchaseOrder::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
//	m_wndSupplier.SetCheckValue(true);
	m_wndSupplier.LimitText(35);
//	m_wndStockList.SetCheckValue(true);

	m_wndSupplier.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndSupplier.InsertColumn(1, _T("Description"), CFMT_TEXT, 320);


	m_wndStockList.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStockList.InsertColumn(1, _T("Description"), CFMT_TEXT, 450);

	m_wndType.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndType.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);


}
void CMADetailPurchaseOrder::OnSetWindowEvents(){
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
	m_wndStockList.SetEvent(WE_SELENDOK, _OnStockListSelendokFnc);
	//m_wndStockList.SetEvent(WE_SETFOCUS, _OnStockListSetfocusFnc);
	//m_wndStockList.SetEvent(WE_KILLFOCUS, _OnStockListKillfocusFnc);
	m_wndStockList.SetEvent(WE_SELCHANGE, _OnStockListSelectChangeFnc);
	m_wndStockList.SetEvent(WE_LOADDATA, _OnStockListLoadDataFnc);
	//m_wndStockList.SetEvent(WE_ADDNEW, _OnStockListAddNewFnc);

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
	OnStockListLoadData();
	SetMode(VM_EDIT);

}
void CMADetailPurchaseOrder::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szSupplierKey);
	DDX_TextEx(pDX, m_wndStockList.GetDlgCtrlID(), m_szStockKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	//DDX_Check(pDX, m_wndGroupByDrug.GetDlgCtrlID(), m_bGroupbyDrug);
}
void CMADetailPurchaseOrder::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CMADetailPurchaseOrder::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CMADetailPurchaseOrder::SetDefaultValues(){
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szSupplierKey.Empty();
	m_szStockKey.Empty();
}
int CMADetailPurchaseOrder::SetMode(int nMode){
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
/*void CMADetailPurchaseOrder::OnFromDateChange(){
	
} */
/*void CMADetailPurchaseOrder::OnFromDateSetfocus(){
	
} */
/*void CMADetailPurchaseOrder::OnFromDateKillfocus(){
	
} */
int CMADetailPurchaseOrder::OnFromDateCheckValue(){
	return 0;
} 
/*void CMADetailPurchaseOrder::OnToDateChange(){
	
} */
/*void CMADetailPurchaseOrder::OnToDateSetfocus(){
	
} */
/*void CMADetailPurchaseOrder::OnToDateKillfocus(){
	
} */
int CMADetailPurchaseOrder::OnToDateCheckValue(){
	return 0;
} 
void CMADetailPurchaseOrder::OnSupplierSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CMADetailPurchaseOrder::OnSupplierSelendok(){
	 
}
/*void CMADetailPurchaseOrder::OnSupplierSetfocus(){
	
}*/
/*void CMADetailPurchaseOrder::OnSupplierKillfocus(){
	
}*/
long CMADetailPurchaseOrder::OnSupplierLoadData(){
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
/*void CMADetailPurchaseOrder::OnSupplierAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CMADetailPurchaseOrder::OnStockListSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CMADetailPurchaseOrder::OnStockListSelendok(){
	 
}
/*void CMADetailPurchaseOrder::OnStockListSetfocus(){
	
}*/
/*void CMADetailPurchaseOrder::OnStockListKillfocus(){
	
}*/
long CMADetailPurchaseOrder::OnStockListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND msl_type = 'A'"));
	return pMF->LoadStorageIntoListCtrl(&m_wndStockList, szFilter);
	//CRecord rs(&pMF->m_db);
	//CString szSQL, szWhere;
	//int nCount = 0;
	//szSQL.Format(_T("SELECT msl_storage_id as id, msl_name as name FROM m_storagelist WHERE msl_type='A' %s ORDER BY msl_storage_id"), szWhere);
	//nCount = rs.ExecSQL(szSQL);
	//m_wndStockList.BeginLoad();
	//while(!rs.IsEOF()){ 
	//	m_wndStockList.AddItems(
	//		rs.GetValue(_T("id")), 
	//		rs.GetValue(_T("name")), NULL);
	//	rs.MoveNext();
	//}
	//m_wndStockList.EndLoad();
	//return nCount;

}
/*void CMADetailPurchaseOrder::OnStockListAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */

void CMADetailPurchaseOrder::OnTypeSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CMADetailPurchaseOrder::OnTypeSelendok(){
	 
}
/*void CMADetailPurchaseOrder::OnTypeSetfocus(){
	
}*/
/*void CMADetailPurchaseOrder::OnTypeKillfocus(){
	
}*/
long CMADetailPurchaseOrder::OnTypeLoadData(){
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
/*void CMADetailPurchaseOrder::OnTypeAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */

void CMADetailPurchaseOrder::OnPrintPreviewSelect(){

}
void CMADetailPurchaseOrder::OnPrintSelect(){

}
void CMADetailPurchaseOrder::OnExportSelect(){

}
void CMADetailPurchaseOrder::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 

void CMADetailPurchaseOrder::OnPrint(bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
	CReport rpt;
	CString tmpStr, szSQL,szWhere;
	CRecord rs(&pMF->m_db);
	CReportSection* rptDetail;
	long double nSubAmount=0, nVatAmount=0, nTotalAmount=0, nTotalPayment=0;
	UpdateData(TRUE);
	if(!m_bDetail)
	{
		if(!rpt.Init(_T("Reports/HMS/PMS_GENERALPURCHASEORDERREPORT.RPT"),true) )
			return;
		rpt.GetReportHeader()->SetValue(_T("HEAlTHSERVICE"), pMF->m_szHealthService);
		rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
		tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
			CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
		rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
		//Page Header
		//Report Detail
		
		CString m_szStock,szStockName;
		for (int i=0 ; i< m_wndStockList.GetItemCount(); i++)
		{
			if(m_wndStockList.GetCheck(i))
			{
				m_wndStockList.SetCurSel(i);
				if(!m_szStock.IsEmpty())
				{
					m_szStock += _T(",");
					szStockName +=_T(", ");

				}
				m_szStock.AppendFormat(_T("%s"), m_wndStockList.GetItemText(i, 0));				
				szStockName.AppendFormat(_T("%s"), m_wndStockList.GetItemText(i, 1));
			}		
		}
		
		if(!m_szSupplierKey.IsEmpty())
			szWhere.AppendFormat(_T("and pmpo_suppid='%s' "), m_szSupplierKey);
		if(!m_szStockKey.IsEmpty())
			szWhere.AppendFormat(_T("and pmpo_stockid in(%s)"), m_szStock);

		szSQL.Format(_T(" SELECT 	pmpo_invoiceno as invoiceno,") \
			_T(" 	pmpo_date as invoicedate,") \
			_T(" 	(select distinct pms_name from pms_suppliers where pms_id=pmpo_suppid) as supplier,") \
			_T(" 	(select distinct pmsl_name from pms_stocklist where pmsl_stockid=pmpo_stockid) as stockname,") \
			_T(" 	pmpo_tax1  as vat,") \
			_T(" 	pmpo_subtotal as subamount,") \
			_T(" 	pmpo_tax1amt as vatamount,") \
			_T(" 	pmpo_total as amount, ") \
			_T(" 	pmpo_payment as payment") \
			_T(" FROM pms_purchaseinvoice ") \
			_T(" WHERE pmpo_status <>'O' ")
			_T(" AND (pmpo_date) between date('%s') and date('%s') %s "),
			m_szFromDate, m_szToDate, szWhere);
		rs.ExecSQL(szSQL);
		int nIndex=1;
		while (!rs.IsEOF())
		{
			rptDetail = rpt.AddDetail();
			tmpStr.Format(_T("%d"), nIndex++);
			rptDetail->SetValue(_T("1"), tmpStr);
			rs.GetValue(_T("invoiceno"), tmpStr);
			rptDetail->SetValue(_T("2"), tmpStr);
			rs.GetValue(_T("invoicedate"), tmpStr);
			rptDetail->SetValue(_T("3"), tmpStr);
			rs.GetValue(_T("supplier"), tmpStr);
			rptDetail->SetValue(_T("4"), tmpStr);
			rs.GetValue(_T("stockname"), tmpStr);
			rptDetail->SetValue(_T("5"), tmpStr);
			rs.GetValue(_T("subamount"), tmpStr);
			nSubAmount+= ToDouble(tmpStr);
			rptDetail->SetValue(_T("6"), tmpStr);
			rs.GetValue(_T("vat"), tmpStr);
			rptDetail->SetValue(_T("7"), tmpStr);
			rs.GetValue(_T("vatamount"), tmpStr);
			nVatAmount+= ToDouble(tmpStr);
			rptDetail->SetValue(_T("8"), tmpStr);
			rs.GetValue(_T("amount"), tmpStr);
			nTotalAmount+= ToDouble(tmpStr);
			rptDetail->SetValue(_T("9"), tmpStr);
			nTotalPayment+= ToDouble(tmpStr);
			rptDetail->SetValue(_T("10"), tmpStr);
			rs.MoveNext();
		}


		
		//Page Footer
		//Report Footer
		rpt.GetReportFooter()->SetValue(_T("s6"), ToString(nSubAmount));
		rpt.GetReportFooter()->SetValue(_T("s7"), _T(""));
		rpt.GetReportFooter()->SetValue(_T("s8"), ToString(nVatAmount));
		rpt.GetReportFooter()->SetValue(_T("s9"), ToString(nTotalAmount));
		rpt.GetReportFooter()->SetValue(_T("s10"), ToString(nTotalPayment));

		CString szSysDate, szDate;
		szSysDate = pMF->GetSysDate(); 
		szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
		rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);

		if(bPreview)
			rpt.PrintPreview();
		else
			rpt.Print();
	}
	else
	{
		CString szInvoiceNo, szWhere;
		CRecord prs(&pMF->m_db);

		CString m_szStock,szStockName;
		for (int i=0 ; i< m_wndStockList.GetItemCount(); i++)
		{
			if(m_wndStockList.GetCheck(i))
			{
				m_wndStockList.SetCurSel(i);
				if(!m_szStock.IsEmpty())
				{
					m_szStock += _T(",");
					szStockName +=_T(", ");

				}
				m_szStock.AppendFormat(_T("%s"), m_wndStockList.GetItemText(i, 0));				
				szStockName.AppendFormat(_T("%s"), m_wndStockList.GetItemText(i, 1));
			}		
		}
		
		if(!m_szSupplierKey.IsEmpty())
			szWhere.AppendFormat(_T("and pmpo_suppid='%s' "), m_szSupplierKey);
		if(!m_szStockKey.IsEmpty())
			szWhere.AppendFormat(_T("and pmpo_stockid in(%s)"), m_szStock);

		szSQL.Format(_T(" SELECT 	pmpo_invoiceno as invoiceno, pmpo_invoiceno||' ['||to_char(pmpo_date, 'dd/mm/yyyy')||']' as invoicename,") \
			_T(" 	(select distinct pms_name from pms_suppliers where pms_id=pmpo_suppid) as supplier,") \
			_T(" 	(select distinct pmsl_name from pms_stocklist where pmsl_stockid=pmpo_stockid) as stockname, ") \
			_T("	pmpo_total as amount ") \
			_T(" FROM pms_purchaseinvoice") \
			_T(" WHERE pmpo_date between date('%s') and date('%s') and pmpo_total > 0 %s ") \
			_T(" ORDER BY pmpo_date"), m_szFromDate, m_szToDate, szWhere);
		prs.ExecSQL(szSQL);
		if(!rpt.Init(_T("Reports/HMS/PMS_DETAILPURCHASEORDERS.RPT"),true) )
			return;
		rpt.GetReportHeader()->SetValue(_T("HEAlTHSERVICE"), pMF->m_szHealthService);
		rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
		tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
			CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
		rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

		while(!prs.IsEOF())
		{
			prs.GetValue(_T("invoiceno"), szInvoiceNo);
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			prs.GetValue(_T("invoicename"), tmpStr);
			rptDetail->SetValue(_T("Invoiceno"), tmpStr);
			prs.GetValue(_T("supplier"), tmpStr);
			rptDetail->SetValue(_T("Supplier"), tmpStr);
			prs.GetValue(_T("stockname"), tmpStr);
			rptDetail->SetValue(_T("Stockname"), tmpStr);
			prs.GetValue(_T("amount"), tmpStr);
			nTotalAmount += ToDouble(tmpStr);
			rptDetail->SetValue(_T("Amount"), tmpStr);

			szSQL.Format(_T(" SELECT 	pmi_id as id,") \
				_T(" 	pmi_name as name,") \
				_T(" 	pmi_unit as unit,") \
				_T(" 	pmsi_expdate as expdate,") \
				_T(" 	pmsi_lotno as lotno,") \
				_T(" 	pmsi_unitprice as unitprice,") \
				_T(" 	pmsi_vatprice as vatprice,") \
				_T(" 	pmpol_qty as qty,") \
				_T(" 	pmpol_qty*pmsi_unitprice as subamount,") \
				_T(" 	pmpol_qty*pmsi_vatprice as amount, ") \
				_T("	(select distinct sc_name from sys_country where sc_id=pmsi_countryid) as orgcountry ") \
				_T(" FROM pms_purchaseinvoice_line") \
				_T(" LEFT JOIN pms_stockitems ON(pmsi_id=pmpol_sitemid)") \
				_T(" LEFT JOIN pms_items ON(pmi_id=pmsi_itemid)") \
				_T(" WHERE pmpol_id='%s'") \
				_T(" ORDER BY pmpol_lineidx"), szInvoiceNo);

			//Page Header
			//Report Detail
			rs.ExecSQL(szSQL);
			int nIndex=1;
			while (!rs.IsEOF())
			{
				rptDetail = rpt.AddDetail();
				tmpStr.Format(_T("%d"), nIndex++);
				rptDetail->SetValue(_T("1"), tmpStr);
				rs.GetValue(_T("name"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rs.GetValue(_T("id"), tmpStr);
				rptDetail->SetValue(_T("3"), tmpStr);
				rs.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				rs.GetValue(_T("expdate"), tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				rs.GetValue(_T("lotno"), tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);
				rs.GetValue(_T("qty"), tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);
				rs.GetValue(_T("vatprice"), tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				rs.GetValue(_T("amount"), tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				rs.GetValue(_T("orgcountry"), tmpStr);
				rptDetail->SetValue(_T("11"), tmpStr);

				rs.MoveNext();
			}
			prs.MoveNext();
		}
		//Page Footer
		//Report Footer
		rpt.GetReportFooter()->SetValue(_T("TotalAmount"), ToString(nTotalAmount));
		CString szSysDate, szDate;
		szSysDate = pMF->GetSysDate(); 
		szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
		rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);

		if(bPreview)
			rpt.PrintPreview();
		else
			rpt.Print();
	}
}

void CMADetailPurchaseOrder::OnPrintTonghopchitietphieunhapkho(bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
	CReport rpt;
	CString tmpStr, szSQL,szWhere, szRptName;
	CRecord rs(&pMF->m_db);
	CReportSection* rptDetail;
	long double nSubAmount=0, nVatAmount=0, nTotalAmount=0, nTotalAmount2 = 0, nTotalPayment=0;
	long nCost = 0;
	UpdateData(TRUE);
	
	CString szInvoiceNo;
	CRecord prs(&pMF->m_db);

	CString szStockName, szStockID;
	for (int i=0 ; i< m_wndStockList.GetItemCount(); i++)
	{
		if(m_wndStockList.GetCheck(i))
		{
			m_wndStockList.SetCurSel(i);
			if(!szStockName.IsEmpty())
			{
				szStockID += _T(", ");
				szStockName +=_T(", ");

			}
			szStockID.AppendFormat(_T("%s"), m_wndStockList.GetItemText(i, 0));
			szStockName.AppendFormat(_T("%s"), m_wndStockList.GetItemText(i, 1));
		}		
	}
	if (pMF->GetModuleID() == _T("PM"))
		szRptName.Format(_T("Reports/HMS/PMS_DETAILPURCHASEORDERS.RPT"));
	else
		szRptName.Format(_T("Reports/HMS/MA_DETAILPURCHASEORDERS.RPT"));
	if(!rpt.Init(szRptName,true) )
		return;
	rpt.GetReportHeader()->SetValue(_T("HEAlTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("StockName"), m_szStockKey);

	szSQL = GetQueryTonghopchitietphieunhapkho();
	int nCount = prs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		//QueryOpener(szSQL);
		return;
	}

	while(!prs.IsEOF())
	{
		if (!m_bGroupbyDrug)
		{
			prs.GetValue(_T("orderid"), szInvoiceNo);
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			tmpStr.Format(_T("%s -%s"), prs.GetValue(_T("ordername")), prs.GetValue(_T("invoicename")));
			rptDetail->SetValue(_T("Invoiceno"), tmpStr);
			prs.GetValue(_T("supplier"), tmpStr);
			rptDetail->SetValue(_T("Supplier"), tmpStr);
			prs.GetValue(_T("stockname"), tmpStr);
			rptDetail->SetValue(_T("Stockname"), tmpStr);
			prs.GetValue(_T("amount"), nCost);
			nTotalAmount += nCost;
			FormatCurrency(nCost, tmpStr);
			szSQL.Format(_T(" SELECT") \
						_T("   mp_code                  AS id,") \
						_T("   mp_name                         AS name,") \
						_T("   get_uomname(mp_uom_id)          AS unit,") \
						_T("   (mpi_expdate)                   AS expdate,") \
						_T("   mpi_lotno                       AS lotno,") \
						_T("   mpi_unitprice                   AS unitprice,") \
						_T("   pol_qtyorder                    AS qty,") \
						_T("   pol_qtyorder*mpi_unitprice      AS amount,") \
						_T("   get_countryname((mp_country_id)) AS orgcountry") \
						_T(" FROM   purchase_orderline") \
						_T(" LEFT JOIN m_product_item ON(mpi_product_item_id=pol_product_item_id)") \
						_T(" LEFT JOIN m_product ON(mp_product_id=mpi_product_id)") \
						_T(" WHERE pol_purchase_order_id = '%s'") \
						_T(" ORDER BY pol_line"), szInvoiceNo);
		}
		else
		{
			szWhere.Format(_T(" AND po_status = 'A' AND po_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
				, m_szFromDate, m_szToDate);
			if (!szStockID.IsEmpty())
				szWhere.AppendFormat(_T(" AND po_storage_id IN (%s)"), szStockID);
			if (!m_szSupplierKey.IsEmpty())
				szWhere.AppendFormat(_T(" AND po_partner_id = '%s'"), m_szSupplierKey);
			szWhere.AppendFormat(_T(" AND po_org_id = '%s'"), pMF->GetModuleID());
			if (!m_szTypeKey.IsEmpty())
				szWhere.AppendFormat(_T(" AND mp_producttype = '%s'"), m_szTypeKey);
			szSQL.Format(_T(" SELECT") \
						_T("   mp_code                  AS id,") \
						_T("   mp_name                         AS name,") \
						_T("   get_uomname(mp_uom_id)          AS unit,") \
						_T("   (mpi_expdate)                   AS expdate,") \
						_T("   mpi_lotno                       AS lotno,") \
						_T("   mpi_unitprice                   AS unitprice,") \
						_T("   sum(pol_qtyorder)                    AS qty,") \
						_T("   sum(pol_qtyorder*mpi_unitprice)      AS amount,") \
						_T("   get_countryname((mp_country_id)) AS orgcountry") \
						_T(" FROM   purchase_order") \
						_T(" LEFT JOIN purchase_orderline ON (po_purchase_order_id = pol_purchase_order_id)") \
						_T(" LEFT JOIN m_product_item ON(mpi_product_item_id=pol_product_item_id)") \
						_T(" LEFT JOIN m_product ON(mp_product_id=mpi_product_id)") \
						_T(" WHERE 1=1 %s") \
						_T(" GROUP BY mp_code, mp_name, mp_uom_id, mpi_expdate, mpi_lotno, mpi_unitprice, mp_country_id") \
						_T(" ORDER BY mp_name"), szWhere);
		}

		//Page Header
		//Report Detail
		//szSQL.Format(_T("select pmpil_id from pms_purchaseinvoice_line where pmpil_id='%s' "), szInvoiceNo);

		rs.ExecSQL(szSQL);
		int nIndex=1;
		while (!rs.IsEOF())
		{
			rptDetail = rpt.AddDetail();
			tmpStr.Format(_T("%d"), nIndex++);
			rptDetail->SetValue(_T("1"), tmpStr);
			rs.GetValue(_T("name"), tmpStr);
			rptDetail->SetValue(_T("2"), tmpStr);
			rs.GetValue(_T("id"), tmpStr);
			rptDetail->SetValue(_T("3"), tmpStr);
			rs.GetValue(_T("unit"), tmpStr);
			rptDetail->SetValue(_T("4"), tmpStr);
			rs.GetValue(_T("expdate"), tmpStr);
			rptDetail->SetValue(_T("5"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
			rs.GetValue(_T("lotno"), tmpStr);
			rptDetail->SetValue(_T("6"), tmpStr);
			rs.GetValue(_T("qty"), nCost);
			tmpStr.Format(_T("%ld"), nCost);
			rptDetail->SetValue(_T("7"), tmpStr);
			rptDetail->SetValue(_T("8"), tmpStr);
			rs.GetValue(_T("unitprice"), nCost);
			tmpStr.Format(_T("%ld"), nCost);
			rptDetail->SetValue(_T("9"), tmpStr);
			rs.GetValue(_T("amount"), nCost);
			nTotalAmount2 += nCost;
			tmpStr.Format(_T("%ld"), nCost);
			rptDetail->SetValue(_T("10"), tmpStr);
			rs.GetValue(_T("orgcountry"), tmpStr);
			rptDetail->SetValue(_T("11"), tmpStr);

			rs.MoveNext();
		}
		if (m_bGroupbyDrug)
			break;
		prs.MoveNext();
	}
	//Page Footer
	//Report Footer
	if (!m_bGroupbyDrug)
		tmpStr.Format(_T("%f"), nTotalAmount);
	else
		tmpStr.Format(_T("%f"), nTotalAmount2);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	CString szSysDate, szDate;
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print();
}

void CMADetailPurchaseOrder::OnExportTonghopchitietphieunhapkho(bool bPreview)
{
	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
	UpdateData(true);
	CString tmpStr, szSQL, szTemp, szWhere, szStockID;
	CRecord rs(&pMF->m_db);
	CReportSection* rptDetail = NULL;
	long double nSubAmount=0, nVatAmount=0, nTotalAmount=0, nTotalAmount2 = 0, nTotalPayment=0;
	double nCost = 0;
	int nIndex = 1;
	szSQL = GetQueryTonghopchitietphieunhapkho();
	BeginWaitCursor();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);	
		return;
	}
	CExcel xls;	
	CStringArray arrCol;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);		
	CellFormat df(&xls), cf(&xls), hf(&xls);
	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER);
	cf.SetFontName(_T("Segoe UI"));
	xls.SetColumnWidth(0,8);
	xls.SetColumnWidth(1,25);
	xls.SetColumnWidth(2,12);
	xls.SetColumnWidth(3,12);
	xls.SetColumnWidth(4,12);
	xls.SetColumnWidth(5,12);
	xls.SetColumnWidth(6,12);
	xls.SetColumnWidth(7,12);
	xls.SetColumnWidth(8,20);
	
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellText(0, 0, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(0, 2, 10);
	TranslateString(_T("Detail Purchase Order"), szTemp);
	StringUpper(szTemp, tmpStr);
	xls.SetCellText(0, 2, tmpStr, FMT_TEXT | FMT_CENTER, true, 12);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellMergedColumns(0, 3, 10);
	xls.SetCellText(0, 3, tmpStr, &df);
	int nRow = 3;
	if (!m_szStock.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("T\x1EEB kho: %s"), m_szStock);
		xls.SetCellMergedColumns(0, nRow, 10);
		xls.SetCellText(0, nRow, tmpStr, &cf);
	}
	nRow++;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn h\xE0ng"));
	arrCol.Add(_T("\x110\x1A1n v\x1ECB"));
	arrCol.Add(_T("N\x1B0\x1EDB\x63 S\x58"));
	arrCol.Add(_T("S\x1ED1 H\x110"));
	arrCol.Add(_T("Ng\xE0y nh\x1EADp"));
	arrCol.Add(_T("S\x1ED1 nh\x1EADp l\x1EC7nh"));
	arrCol.Add(_T("S\x1ED1 l\x1B0\x1EE3ng"));
	arrCol.Add(_T("\x110\x1A1n gi\xE1"));
	arrCol.Add(_T("Th\xE0nh ti\x1EC1n"));
	for (int i = 0; i < arrCol.GetCount(); i++)
		xls.SetCellText(i, nRow, arrCol.GetAt(i), &hf);
	for (int i = 0; i < m_wndStockList.GetItemCount(); i++){
		if (m_wndStockList.GetCheck(i))
		{
			m_wndStockList.SetCheck(i);
			if (!szStockID.IsEmpty())
				szStockID += _T(", ");
			szStockID.AppendFormat(_T("%s"), m_wndStockList.GetItemText(i, 0));
		}
	}
	CString szInvoiceNo;
	while(!rs.IsEOF())
	{
		//Page Header
		//Report Detail
		nRow ++;
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("Name"), tmpStr);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("Uom_name"), tmpStr);
		xls.SetCellText(2, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("Country_name"), tmpStr);
		xls.SetCellText(3, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("Invoice_no"), tmpStr);
		xls.SetCellText(4, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("Import_date"), tmpStr);
		xls.SetCellText(5, nRow, CDate::Convert(tmpStr, yyyymmdd,ddmmyyyy), FMT_TEXT);
		rs.GetValue(_T("Order_no"), tmpStr);
		xls.SetCellText(6, nRow, tmpStr, FMT_TEXT);				
		rs.GetValue(_T("qty"), tmpStr);
		xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("tax_price"), tmpStr);
		xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("total_amount"), nCost);
		nTotalAmount += nCost;
		xls.SetCellText(9, nRow, double2str(nCost), FMT_NUMBER1);
		rs.MoveNext();
	}
	nRow++;
	xls.SetCellMergedColumns(0, nRow, 9);
	TranslateString(_T("Total Amount"), tmpStr);
	xls.SetCellText(0, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
	tmpStr.Format(_T("%f"), nTotalAmount);
	xls.SetCellText(9, nRow, tmpStr, FMT_NUMBER1, true);
	EndWaitCursor();
	xls.Save(_T("Exports\\Chi Tiet Phieu Nhap Kho.xls"));
}


CString CMADetailPurchaseOrder::GetQueryTonghopchitietphieunhapkho(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szWhere;
	CString szStock;
	UpdateData(true);
	m_szStock.Empty();
	for (int i=0 ; i< m_wndStockList.GetItemCount(); i++)
	{
		if(m_wndStockList.GetCheck(i))
		{
			m_wndStockList.SetCurSel(i);
			if(!szStock.IsEmpty())
			{
				szStock += _T(", ");
				m_szStock += _T(", ");
			}	
			szStock.AppendFormat(_T("%s"), m_wndStockList.GetItemText(i, 0));				
			m_szStock.AppendFormat(_T("%s"), m_wndStockList.GetItemText(i, 1));				
		}		
	}
	szWhere.Format(_T(" AND po_status = 'A' AND po_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
		, m_szFromDate, m_szToDate);
	if(!m_szSupplierKey.IsEmpty())
		szWhere.AppendFormat(_T("and po_partner_id='%s' "), m_szSupplierKey);
	if(!szStock.IsEmpty())
		szWhere.AppendFormat(_T("and po_storage_id in(%s)"), szStock);
	if (!m_szTypeKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND product_producttype = '%s'"), m_szTypeKey);
	szWhere.AppendFormat(_T(" AND po_org_id = '%s'"), pMF->GetModuleID());
	szSQL.Format(_T(" SELECT po_approveddate import_date,") \
				_T("   get_partnername(po_partner_id) partner_name,") \
				_T("   po_orderno order_no,") \
				_T("   po_invoiceno invoice_no,") \
				_T("   product_uomname uom_name,") \
				_T("   product_countryname country_name,") \
				_T("   product_code AS code,") \
				_T("   product_name AS name,") \
				_T("   pol_unitprice+pol_unitprice*pol_tax_id/100 AS tax_price,") \
				_T("   SUM(pol_qtyorder) AS qty,") \
				_T("   SUM(pol_totalamount) AS total_amount") \
				_T(" FROM   purchase_order") \
				_T(" LEFT JOIN purchase_orderline ON (po_purchase_order_id = pol_purchase_order_id)") \
				_T(" LEFT JOIN m_product_view ON(product_id=pol_product_id)") \
				_T(" WHERE 1=1 ") \
				_T("	%s") \
				_T(" GROUP BY product_code, product_name, pol_unitprice, pol_tax_id, po_partner_id, po_approveddate, ") \
				_T("		  po_orderno, product_uomname, product_countryname, po_invoiceno") \
				_T(" ORDER BY name, partner_name"), szWhere);
	return szSQL;
}

BEGIN_MESSAGE_MAP(CMADetailPurchaseOrder, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CMADetailPurchaseOrder::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
