#include "stdafx.h"
#include "PMReportConditionDialog.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include ".\pmreportconditiondialog.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSReportConditionDialog *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSefocusFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSReportConditionDialog *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSReportConditionDialog* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CPMSReportConditionDialog *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnStockAddNew();
}*/
static void _OnTypeSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSReportConditionDialog* )pWnd)->OnTypeSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnTypeSelendokFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnTypeSelendok();
}
/*static void _OnTypeSetfocusFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnTypeKillfocus();
}*/
/*static void _OnTypeKillfocusFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnTypeKillfocus();
}*/
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CPMSReportConditionDialog *)pWnd)->OnTypeLoadData();
}
/*static void _OnTypeAddNewFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnTypeAddNew();
}*/
static long _OnSupplierLoadDataFnc(CWnd *pWnd){
	return ((CPMSReportConditionDialog *)pWnd)->OnSupplierLoadData();
}
static void _OnGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSReportConditionDialog* )pWnd)->OnGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnGroupSelendokFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnGroupSelendok();
}
/*static void _OnGroupSetfocusFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnGroupKillfocus();
}*/
/*static void _OnGroupKillfocusFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnGroupKillfocus();
}*/
static long _OnGroupLoadDataFnc(CWnd *pWnd){
	return ((CPMSReportConditionDialog *)pWnd)->OnGroupLoadData();
}
/*static void _OnGroupAddNewFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnGroupAddNew();
}*/
static void _OnSourceSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSReportConditionDialog* )pWnd)->OnSourceSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSourceSelendokFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnSourceSelendok();
}
/*static void _OnSourceSetfocusFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnSourceKillfocus();
}*/
/*static void _OnSourceKillfocusFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnSourceKillfocus();
}*/
static long _OnSourceLoadDataFnc(CWnd *pWnd){
	return ((CPMSReportConditionDialog *)pWnd)->OnSourceLoadData();
}
/*static void _OnSourceAddNewFnc(CWnd *pWnd){
	((CPMSReportConditionDialog *)pWnd)->OnSourceAddNew();
}*/
static void _OnSortByTypeSelectFnc(CWnd *pWnd){
	CPMSReportConditionDialog *pVw = (CPMSReportConditionDialog *)pWnd;
	pVw->OnSortByTypeSelect();
} 
static void _OnCatSelectFnc(CWnd *pWnd){
	CPMSReportConditionDialog *pVw = (CPMSReportConditionDialog *)pWnd;
	pVw->OnCatSelect();
} 
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSReportConditionDialog *pVw = (CPMSReportConditionDialog *)pWnd;
	pVw->OnPrintPreviewSelect();
}
static void _OnPrintPreviewInsSelectFnc(CWnd *pWnd){
	((CPMSReportConditionDialog*)pWnd)->OnPrintPreviewInsSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CPMSReportConditionDialog *pVw = (CPMSReportConditionDialog *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSReportConditionDialog *pVw = (CPMSReportConditionDialog *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CPMSReportConditionDialog *pVw = (CPMSReportConditionDialog *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddPMSReportConditionDialogFnc(CWnd *pWnd){
	 return ((CPMSReportConditionDialog*)pWnd)->OnAddPMSReportConditionDialog();
} 
static int _OnEditPMSReportConditionDialogFnc(CWnd *pWnd){
	 return ((CPMSReportConditionDialog*)pWnd)->OnEditPMSReportConditionDialog();
} 
static int _OnDeletePMSReportConditionDialogFnc(CWnd *pWnd){
	 return ((CPMSReportConditionDialog*)pWnd)->OnDeletePMSReportConditionDialog();
} 
static int _OnSavePMSReportConditionDialogFnc(CWnd *pWnd){
	 return ((CPMSReportConditionDialog*)pWnd)->OnSavePMSReportConditionDialog();
} 
static int _OnCancelPMSReportConditionDialogFnc(CWnd *pWnd){
	 return ((CPMSReportConditionDialog*)pWnd)->OnCancelPMSReportConditionDialog();
} 
CPMSReportConditionDialog::CPMSReportConditionDialog(CWnd *pParent, int nType)
{
	m_nDlgWidth = 645;
	m_nDlgHeight = 485;
	m_nType = nType;
	m_szSort = _T("N");
	SetDefaultValues();
}
CPMSReportConditionDialog::~CPMSReportConditionDialog(){
}
void CPMSReportConditionDialog::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 180);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 110, 55);
	m_wndFromDate.Create(this,115, 30, 300, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 305, 30, 385, 55);
	m_wndToDate.Create(this,390, 30, 570, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 110, 85);
	m_wndStock.Create(this,115, 60, 570, 85); 
	if (m_nType == INSTOCK || m_nType == IMPORT_EXPORT_INSTOCK_DETAILFORMONEY)
		m_wndStock.SetCheckBox(true);
	if (m_nType == INSTOCK)
	{
		m_wndTypeLabel.Create(this, _T("Type"), 10, 90, 110, 115);
		m_wndType.Create(this,115, 90, 300, 115); 
		m_wndSupplierLabel.Create(this, _T("Supplier"), 305, 90, 385, 115);
		m_wndSupplier.Create(this,390, 90, 570, 115); 
	}
	else
	{
		m_wndTypeLabel.Create(this, _T("Type"), 10, 90, 110, 115);
		m_wndType.Create(this,115, 90, 570, 115); 
	}
	m_wndGroupLabel.Create(this, _T("Group"), 10, 120, 110, 145);
	m_wndGroup.Create(this,115, 120, 570, 145); 
	m_wndSourceLabel.Create(this, _T("Source"), 10, 150, 110, 175);
	m_wndSource.Create(this,115, 150, 570, 175); 
	if (m_nType == INSTOCK || m_nType == IMPORT_EXPORT_INSTOCK_DETAILFORMONEY)
	{
		if (m_nType == INSTOCK)
		{
			m_wndCheck.Create(this, _T("No show Qty"), 20, 185, 165, 210);
			m_wndServPrice.Create(this, _T("Service Price"), 175, 185, 265, 210);
		}
		m_wndCat.Create(this, _T("Sort by Category"), 290, 185, 415, 210);
		m_wndSortByType.Create(this, _T("Sort by Type"), 430, 185, 570, 210);
	}
	if (m_nType == IMPORT_EXPORT_INSTOCK_DETAILFORMONEY)
	{
		m_wndBillPrice.Create(this, _T("Bill Price"), 5, 215, 85, 240);
		m_wndPrintPreviewIns.Create(this, _T("Print Preview (Ins)"), 265, 215, 375, 240);
	}
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 380, 215, 490, 240);
	m_wndExport.Create(this, _T("&Export"), 495, 215, 575, 240);

}
void CPMSReportConditionDialog::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);
	

	m_wndStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStock.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);


	m_wndType.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 50);
	m_wndType.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndSupplier.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndSupplier.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);

	m_wndGroup.InsertColumn(0, _T(""), CFMT_TEXT, 0);
	m_wndGroup.InsertColumn(1, _T("Code"), CFMT_TEXT, 50);
	m_wndGroup.InsertColumn(2, _T("Name"), CFMT_TEXT, 400);
	m_wndGroup.Format(3, 2);

	m_wndSource.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndSource.InsertColumn(1, _T("Description"), CFMT_TEXT, 125);
	
}
void CPMSReportConditionDialog::OnSetWindowEvents(){
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
	m_wndType.SetEvent(WE_SELENDOK, _OnTypeSelendokFnc);
	//m_wndType.SetEvent(WE_SETFOCUS, _OnTypeSetfocusFnc);
	//m_wndType.SetEvent(WE_KILLFOCUS, _OnTypeKillfocusFnc);
	m_wndType.SetEvent(WE_SELCHANGE, _OnTypeSelectChangeFnc);
	m_wndType.SetEvent(WE_LOADDATA, _OnTypeLoadDataFnc);
	//m_wndType.SetEvent(WE_ADDNEW, _OnTypeAddNewFnc);
	m_wndSupplier.SetEvent(WE_LOADDATA, _OnSupplierLoadDataFnc);
	m_wndGroup.SetEvent(WE_SELENDOK, _OnGroupSelendokFnc);
	//m_wndGroup.SetEvent(WE_SETFOCUS, _OnGroupSetfocusFnc);
	//m_wndGroup.SetEvent(WE_KILLFOCUS, _OnGroupKillfocusFnc);
	m_wndGroup.SetEvent(WE_SELCHANGE, _OnGroupSelectChangeFnc);
	m_wndGroup.SetEvent(WE_LOADDATA, _OnGroupLoadDataFnc);
	//m_wndGroup.SetEvent(WE_ADDNEW, _OnGroupAddNewFnc);
	m_wndSource.SetEvent(WE_SELENDOK, _OnSourceSelendokFnc);
	//m_wndSource.SetEvent(WE_SETFOCUS, _OnSourceSetfocusFnc);
	//m_wndSource.SetEvent(WE_KILLFOCUS, _OnSourceKillfocusFnc);
	m_wndSource.SetEvent(WE_SELCHANGE, _OnSourceSelectChangeFnc);
	m_wndSource.SetEvent(WE_LOADDATA, _OnSourceLoadDataFnc);
	//m_wndSource.SetEvent(WE_ADDNEW, _OnSourceAddNewFnc);
	m_wndSortByType.SetEvent(WE_CLICK, _OnSortByTypeSelectFnc);
	m_wndCat.SetEvent(WE_CLICK, _OnCatSelectFnc);
	m_wndPrintPreviewIns.SetEvent(WE_CLICK, _OnPrintPreviewInsSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_szFromDate =  pMF->GetSysDate()+ _T("00:00");	
	if (m_nType == INSTOCK)
		m_wndFromDate.EnableWindow(FALSE);
	m_szToDate =pMF->GetSysDate() + _T("23:59");
	m_wndStock.EnableWindow(true);
	UpdateData(false);
}
void CPMSReportConditionDialog::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	if (m_nType == INSTOCK)
		DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szSupplierKey);
	DDX_TextEx(pDX, m_wndGroup.GetDlgCtrlID(), m_szGroupKey);
	DDX_TextEx(pDX, m_wndSource.GetDlgCtrlID(), m_szSourceKey);
	if (m_nType == INSTOCK) DDX_Check(pDX, m_wndCheck.GetDlgCtrlID(), m_bCheck);
	if (m_nType == INSTOCK){
		DDX_Check(pDX, m_wndServPrice.GetDlgCtrlID(), m_bServPrice);
	}
	if (m_nType == IMPORT_EXPORT_INSTOCK_DETAILFORMONEY)
		DDX_Check(pDX, m_wndBillPrice.GetDlgCtrlID(), m_bBillPrice);
}
void CPMSReportConditionDialog::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPMSReportConditionDialog::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CPMSReportConditionDialog::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_szTypeKey.Empty();
	m_szGroupKey.Empty();
	m_szSourceKey.Empty();
	m_bBillPrice = FALSE;

}
int CPMSReportConditionDialog::SetMode(int nMode){
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
 			EnableButtons(FALSE, 0, 1, 2, 3, -1); 
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
/*void CPMSReportConditionDialog::OnFromDateChange(){
	
} */
/*void CPMSReportConditionDialog::OnFromDateSetfocus(){
	
} */
/*void CPMSReportConditionDialog::OnFromDateKillfocus(){
	
} */
int CPMSReportConditionDialog::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSReportConditionDialog::OnToDateChange(){
	
} */
/*void CPMSReportConditionDialog::OnToDateSetfocus(){
	
} */
/*void CPMSReportConditionDialog::OnToDateKillfocus(){
	
} */
int CPMSReportConditionDialog::OnToDateCheckValue(){
	return 0;
} 
void CPMSReportConditionDialog::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSReportConditionDialog::OnStockSelendok(){
	 
}
/*void CPMSReportConditionDialog::OnStockSetfocus(){
	
}*/
/*void CPMSReportConditionDialog::OnStockKillfocus(){
	
}*/
long CPMSReportConditionDialog::OnStockLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageList(&m_wndStock, m_szStockKey);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStock.IsSearchKey() && ToInt(m_szStockKey) > 0){
		szWhere.Format(_T(" AND msl_storage_id=%d"), ToInt(m_szStockKey));
	};

	//szWhere.AppendFormat(_T(" and msl_storage_id in(%s) "), pMF->m_szStorages);

	//szWhere.AppendFormat(_T(" and pmssl_storage_id in(%s) "), pMF->m_szStorages);
	m_wndStock.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT msl_storage_id as id, msl_name as name FROM m_storagelist WHERE 1=1 %s ORDER BY msl_type, msl_storage_id"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStock.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}

/*void CPMSReportConditionDialog::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSReportConditionDialog::OnTypeSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSReportConditionDialog::OnTypeSelendok(){
	 
}
/*void CPMSReportConditionDialog::OnTypeSetfocus(){
	
}*/
/*void CPMSReportConditionDialog::OnTypeKillfocus(){
	
}*/
long CPMSReportConditionDialog::OnTypeLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//return pMF->LoadProductTypeList(&m_wndType, m_szTypeKey);
	CRecord rs(&pMF->m_db);
	CString szSQL;
	int nCount = 0;
	szSQL.Format(_T(" SELECT DISTINCT product_type_id, ") \
	_T("                 product_type_name ") \
	_T(" FROM   (SELECT CASE WHEN mpt_product_type_id IN ( 'A1100', 'A1130', 'A1140', 'A1160', ") \
	_T("                                              'A1170', 'A1180', 'A1190', 'A1400', 'A6000' ) THEN N'A0000' ") \
	_T("                ELSE mpt_product_type_id ") \
	_T("                END product_type_id, ") \
	_T("                CASE WHEN mpt_product_type_id IN ( 'A1100', 'A1130', 'A1140', 'A1160', ") \
	_T("                                              'A1170', 'A1180', 'A1190', 'A1400', 'A6000' ) THEN N'Thu\x1ED1\x63' ") \
	_T("                ELSE mpt_name ") \
	_T("                END product_type_name ") \
	_T("         FROM   m_product_type ") \
	_T("         WHERE  mpt_org_id = '%s' ") \
	_T("            AND mpt_isactive = 'Y' AND mpt_product_type_id <> 'A1000') ") \
	_T(" ORDER BY product_type_id"), pMF->GetModuleID());
	nCount = rs.ExecSQL(szSQL);
	m_wndType.DeleteAllItems();
	while (!rs.IsEOF())
	{
		m_wndType.AddItems(
			rs.GetValue(_T("product_type_id")),
			rs.GetValue(_T("product_type_name")),
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPMSReportConditionDialog::OnTypeAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
long CPMSReportConditionDialog::OnSupplierLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadPartnerList(&m_wndSupplier, m_szSupplierKey);
}
void CPMSReportConditionDialog::OnGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSReportConditionDialog::OnGroupSelendok(){
	 
}
/*void CPMSReportConditionDialog::OnGroupSetfocus(){
	
}*/
/*void CPMSReportConditionDialog::OnGroupKillfocus(){
	
}*/
long CPMSReportConditionDialog::OnGroupLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND mpc_level = 1"));
	return pMF->LoadProductCategoryList(&m_wndGroup, m_szGroupKey, szFilter);

}
/*void CPMSReportConditionDialog::OnGroupAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSReportConditionDialog::OnSourceSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSReportConditionDialog::OnSourceSelendok(){
	 
}
/*void CPMSReportConditionDialog::OnSourceSetfocus(){
	
}*/
/*void CPMSReportConditionDialog::OnSourceKillfocus(){
	
}*/
long CPMSReportConditionDialog::OnSourceLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndSource.IsSearchKey() && str2int(m_szSourceKey) > 0){
		szWhere.Format(_T(" and mpr_product_resource_id=%d"), str2int(m_szSourceKey));
	};
	szWhere.AppendFormat(_T(" AND mpr_org_id = '%s'"), pMF->GetModuleID());
	m_wndSource.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT * FROM m_product_resource WHERE mpr_isactive='Y' %s ORDER BY mpr_product_resource_id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndSource.AddItems(
			rs.GetValue(_T("mpr_product_resource_id")), 
			rs.GetValue(_T("mpr_name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
void CPMSReportConditionDialog::OnSortByTypeSelect(){
	m_szSort = _T("T");
	_debug(_T("%s"), m_szSort);
}
void CPMSReportConditionDialog::OnCatSelect(){
	m_szSort = _T("C");
	_debug(_T("%s"), m_szSort);
}
/*void CPMSReportConditionDialog::OnSourceAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSReportConditionDialog::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_szStockName = m_wndStock.GetCurrent(1);
	switch(m_nType)
	{
		case IMPORT_EXPORT:
			OnPrintImportExportReport(true);
			break;
		case IMPORT_EXPORT_INSTOCK:
			OnPrintImportExportAndInstockReport(true);
			break;
		case IMPORT_EXPORT_INSTOCK_DETAILFORMONEY:
			OnPrintImportExportAndInstockForDetailMoneyReport(true);
			break;
		case INSTOCK:
			OnPrintInstockReport(true);
			break;
		case IMPORT_EXPORT_INSTOCK_DETAIL_PROFESSIONAL:
			OnPrintTonghopnhapxuattonchitiettheonghiepvu(true);
			break;
		case THUOC_HET_HAN:
			OnPrint_Danhsachthuochethan(true);
			break;		
	}
} 
void CPMSReportConditionDialog::OnPrintPreviewInsSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString tmpStr, szDate, szSysDate, szSQL, szWhere, szOrderby, szPrice, szCondition, szRptName, szStockSel;
	//szRptName.Format(_T("Reports/HMS/PMS_SUMMARYIMPORTEXPORTINSTOCKFORMONEYREPORT.RPT"));
	if (pMF->GetModuleID() == _T("PM"))
		szRptName.Format(_T("Reports/HMS/PM_BAOCAOXUATNHAPCHUNG_BH.RPT"));
	else
		szRptName.Format(_T("Reports/HMS/MA_BAOCAOXUATNHAPCHUNG_BH.RPT"));
	if(!rpt.Init(szRptName,true) )
	{
		return;
	}
	
	BeginWaitCursor();
	CString szType, szTypeName, szStock, szStockName;
	for (int i=0 ; i<= m_wndType.GetCount(); i++)
	 {
		 if(m_wndType.GetCheck(i)){
			 m_wndType.SetCurSel(i);
		 	if(!szType.IsEmpty())
					szType += _T(",");						
					szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));			
			if(!szTypeName.IsEmpty())
					szTypeName += _T(", ");						
			szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		 }
	 }
	 for (int i=0 ; i<= m_wndStock.GetCount(); i++)
	 {
		 if(m_wndStock.GetCheck(i)){
			 m_wndStock.SetCurSel(i);
		 	if(!szStock.IsEmpty())
				szStock += _T(",");						
			szStock.AppendFormat(_T("%s"), m_wndStock.GetCurrent(0));			
			if(!szStockName.IsEmpty())
				szStockName += _T(", ");						
			szStockName.AppendFormat(_T("%s"), m_wndStock.GetCurrent(1));
		 }
	 }	
	if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and product_groupid in(%s)"), szType);	
	else
		szTypeName.Format(_T("TÊt c¶ thuèc vµ vËt t­"));

	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(product_rootid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}
	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and product_resourceid=%d "), str2int(m_szSourceKey));
	}
	szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	if (!szStock.IsEmpty())
		szStockSel.Format(_T(" AND stockid IN (%s)"), szStock);
	szOrderby.Format(_T("ORDER BY name"));
	if (m_szSort == _T("C"))
		szOrderby.Format(_T("ORDER BY idx1,name, unit"));
	else if (m_szSort == _T("T"))
		szOrderby.Format(_T("ORDER BY idx2,name, unit"));
	//szPrice = _T("product_taxprice as price");
	//szCondition = _T("product_taxprice");
	//if (m_bServPrice)
	//{
	//	szPrice = _T("product_saleprice1 as price");
	//	szCondition = _T("product_saleprice1");
	//}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
		CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	if (!m_szStockKey.IsEmpty())
		rpt.GetReportHeader()->SetValue(_T("StockName"), szStockName);
	else
		rpt.GetReportHeader()->SetValue(_T("StockName"), _T("To\xE0n \x62\x1ED9"));
	if (!m_szTypeKey.IsEmpty())
		rpt.GetReportHeader()->SetValue(_T("Type"), m_wndType.GetCurrent(1));
	if (!m_szGroupKey.IsEmpty())
		rpt.GetReportHeader()->SetValue(_T("Group"), m_wndGroup.GetCurrent(1));
	else
		rpt.GetReportHeader()->SetValue(_T("Group"), _T("To\xE0n \x62\x1ED9"));
	if (m_szSourceKey.IsEmpty())
		rpt.GetReportHeader()->SetValue(_T("Original"), m_wndSource.GetCurrent(1));

	//Page Header
	//Report Detail
	bool bHavePrice = false;
	szPrice = _T("product_taxprice");
	szSQL.Format(_T("SELECT distinct msl_category FROM m_storagelist WHERE msl_storage_id IN (%s)"), szStock);
	rs.ExecSQL(szSQL);
	if (rs.GetRecordCount() == 1)
	{
		rs.GetValue(_T("msl_category"), tmpStr);
		if (tmpStr == _T("I"))
			szPrice = _T("product_saleprice2");
		else if (tmpStr == _T("S"))
			szPrice = _T("product_saleprice1");
		else if (tmpStr == _T("P"))
			szPrice = _T("product_saleprice3");
	}
	//	if(rpt.GetDetail()->GetItem(_T("5")) != NULL)
	bHavePrice = true;
	if(bHavePrice)
	{
		szSQL.Format(_T(" SELECT * FROM") \
					_T(" ( SELECT  product_rootid as idx1, product_rootname as groupname, product_code as id,        ") \
					_T("		   product_groupid as idx2, product_groupname as typename, product_classname as generic,") \
					_T("           product_name as name, get_uomname(product_purchase_uomid) as unit,       ") \
					_T("           %s as price,         ") \
					_T("           sum(periodqty) as periodqty,    ") \
					_T("           sum(periodqty*%s) as periodamount,   ") \
					_T("           sum(impqty) as impqty,") \
					_T("           sum(impqty*%s) as impamount,         ") \
					_T("           sum(expqty) as expqty,") \
					_T("           sum(expqty) *%s as expamount,        ") \
					_T("           sum(periodqty+impqty-expqty) as instockqty,     ") \
					_T("           sum((periodqty+impqty-expqty))*%s as instockamount ") \
					_T("   FROM ( ") \
					_T("       SELECT * FROM (    ") \
					_T("           SELECT impstockid as stockid, impdate as iodate, sitemid, 0 as periodqty, impqty, 0 as expqty") \
					_T("           FROM miv        ") \
					_T("           UNION ALL") \
					_T("           SELECT expstockid, expdate, sitemid, 0, 0, expqty") \
					_T("           FROM mev ") \
					_T("           ) iotbl ") \
					_T("       WHERE iodate between cast_string2timestamp( '%s' )and cast_string2timestamp ('%s' ) %s") \
					_T("       UNION ALL ") \
					_T("       SELECT stockid, iodate, sitemid, coalesce(impqty-expqty, 0), 0, 0") \
					_T("       FROM (      ") \
					_T("           SELECT impstockid as stockid, impdate as iodate, sitemid,  0 as periodqty, impqty, 0 as expqty") \
					_T("           FROM miv") \
					_T("           UNION ALL       ") \
					_T("           SELECT expstockid, expdate, sitemid, 0, 0, expqty") \
					_T("           FROM mev ") \
					_T("           ) ptbl ") \
					_T("       WHERE iodate < cast_string2timestamp('%s') %s") \
					_T("       ) instock ") \
					_T("     LEFT JOIN m_productitem_view ON(product_item_id=sitemid)") \
					_T("     WHERE sitemid > 0 %s") \
					_T("     GROUP BY product_rootid, product_rootname, product_groupid, product_groupname, product_code,") \
					_T("			  product_name, product_purchase_uomid, product_classname, %s ") \
					_T(" ) tbl ") \
					_T(" WHERE (periodqty + impqty + expqty > 0)  ") \
					_T(" %s"), szPrice, szPrice, szPrice, szPrice, szPrice, 
					m_szFromDate, m_szToDate, szStockSel, m_szFromDate, szStockSel, szWhere, szPrice, szOrderby);
		_fmsg(_T("%s"), szSQL);
	}
	else
	{
		szSQL.Format(_T("SELECT * FROM ") \
			_T(" ( SELECT hfg_index as idx, hfg_name as TypeName, pmi_typeid, pmi_id as id, ") \
			_T(" 	pmi_name as name, ") \
			_T(" 	pmg_name as genericname, ") \
			_T(" 	pmi_unit as unit, ") \
			_T(" 	pmsi_vatprice as price, ") \
			_T(" 	sum(periodqty) as periodqty,") \
			_T(" 	sum(periodqty*pmsi_vatprice) as periodamount,") \
			_T(" 	sum(impqty) as impqty, ") \
			_T(" 	sum(impqty*pmsi_vatprice) as impamount,") \
			_T(" 	sum(expqty) as expqty,") \
			_T(" 	sum(expqty) *pmsi_vatprice as expamount,") \
			_T(" 	sum(periodqty+impqty-expqty) as instockqty,") \
			_T(" 	sum((periodqty+impqty-expqty))*pmsi_vatprice as instockamount") \
			_T(" FROM") \
			_T(" (") \
			_T(" SELECT * FROM") \
			_T(" (") \
			_T(" 	SELECT impstockid as stockid, impdate as iodate, sitemid, 0 as periodqty, impqty, 0 as expqty") \
			_T(" 	FROM pmsv_import") \
			_T(" 	UNION ALL") \
			_T(" 	SELECT impstockid, expdate, sitemid, 0, 0, expqty") \
			_T(" 	FROM pmsv_export") \
			_T(" ) as iotbl") \
			_T(" WHERE 	stockid=%d ") \
			_T(" 	AND iodate between timestamp '%s' and timestamp '%s' ") \
			_T(" UNION ALL") \
			_T(" SELECT stockid, iodate, sitemid, coalesce(impqty-expqty, 0), 0, 0 FROM") \
			_T(" (") \
			_T(" 	SELECT impstockid as stockid, impdate as iodate, sitemid,  0 as periodqty, impqty, 0 as expqty") \
			_T(" 	FROM pmsv_import") \
			_T(" 	UNION ALL") \
			_T(" 	SELECT impstockid, expdate, sitemid, 0, 0, expqty") \
			_T(" 	FROM pmsv_export") \
			_T(" ) as ptbl") \
			_T(" WHERE 	stockid=%d ") \
			_T(" 	and iodate < timestamp '%s' ") \
			_T(" ) as instock") \
			_T(" LEFT JOIN pms_stockitems ON(pmsi_id=sitemid)") \
			_T(" LEFT JOIN pms_items ON(pmsi_itemid=pmi_id)") \
			_T(" LEFT JOIN pms_generic ON(pmg_id=pmi_genericid) ") \
			_T(" LEFT JOIN hms_fee_group ON(hfg_id = pmi_typeid) ") \
			_T(" WHERE sitemid > 0 %s ") \
			_T(" GROUP BY idx, name,pmi_typeid,typename,genericname,id,unit, price") \
			_T(" ) as tbl") \
			_T(" WHERE length(id) > 0 and (impqty + expqty + instockqty > 0) ") \
			_T(" ORDER BY idx, typename,name,genericname, unit"), ToInt(m_szStockKey), m_szFromDate, m_szToDate, ToInt(m_szStockKey), m_szFromDate, szWhere);		}
	CReportSection* rptDetail;
	CString szOldLine, szNewLine;
	
	long double grpCost[23];
	long double ttlCost[23];
	double cost = 0;
	int nItem = 1;
	rs.ExecSQL(szSQL);

	if(rs.IsEOF()){
	//QueryOpener(szSQL);
	}

	long nTotalPeriod=0;
	long nTotalImport=0;
	long nTotalExport=0;
	long nTotalInStock=0;
	
	for (int i =0; i < 16; i++)
	{
		grpCost[i] = 0;
		ttlCost[i] = 0;
	}
	
	while(!rs.IsEOF())
	{
		if (m_szSort == _T("C"))
		{
			rs.GetValue(_T("idx1"), szNewLine);
			if(szNewLine != szOldLine && !szNewLine.IsEmpty()){
			CString szField;
			if((grpCost[7] + grpCost[9] + grpCost[11] + grpCost[13]) > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				for (int i =7; i < 16; i+=2)
				{
					szField.Format(_T("s%d"), i);
					tmpStr.Format(_T("%f"),grpCost[i]);
					rptDetail->SetValue(szField, tmpStr);
					ttlCost[i] += grpCost[i];						
					grpCost[i] = 0;				
				}
			}

			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			rs.GetValue(_T("groupname"), tmpStr);
				
			rptDetail->SetValue(_T("GroupName"),tmpStr );
			szOldLine = szNewLine;
			nItem=1;
			}
		}
		else if (m_szSort == _T("T"))
		{
			rs.GetValue(_T("idx2"), szNewLine);
			if(szNewLine != szOldLine && !szNewLine.IsEmpty()){
			CString szField;
			if((grpCost[7] + grpCost[9] + grpCost[11] + grpCost[13]) > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				for (int i =7; i < 16; i+=2)
				{
					szField.Format(_T("s%d"), i);
					tmpStr.Format(_T("%f"),grpCost[i]);
					rptDetail->SetValue(szField, tmpStr);
					ttlCost[i] += grpCost[i];						
					grpCost[i] = 0;				
				}
			}

			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			rs.GetValue(_T("typename"), tmpStr);
				
			rptDetail->SetValue(_T("GroupName"),tmpStr );
			szOldLine = szNewLine;
			nItem=1;
			}
		}
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nItem++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("name"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("generic"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("price"), cost);
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);			
		rs.GetValue(_T("periodqty"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rs.GetValue(_T("periodamount"), cost);	
		grpCost[7] += cost;
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("7"), cost);				

		rs.GetValue(_T("impqty"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.GetValue(_T("impamount"), cost);		
		grpCost[9] += cost;
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("9"), cost);


		rs.GetValue(_T("expqty"), tmpStr);				
		rptDetail->SetValue(_T("10"), tmpStr);
		rs.GetValue(_T("expamount"), cost);		
		grpCost[11] += cost;				
		rptDetail->SetValue(_T("11"), cost);
		
		rs.GetValue(_T("instockqty"), tmpStr);				
		rptDetail->SetValue(_T("12"), tmpStr);
		rs.GetValue(_T("instockamount"), cost);
		grpCost[13] += cost;				
		rptDetail->SetValue(_T("13"), cost);

		rs.MoveNext();

	}
	for (int i =0; i < 16; i++)
	{
			ttlCost[i] += grpCost[i];
	}
	if (m_szSort != _T("N"))
	if((grpCost[7] + grpCost[9] + grpCost[11] + grpCost[13]) > 0){
		CString szField;
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		for (int i =7; i < 16; i+=2)
		{
			szField.Format(_T("s%d"), i);
			tmpStr.Format(_T("%f"),grpCost[i]);
			rptDetail->SetValue(szField, tmpStr);
		}	
	}
	if((ttlCost[7] + ttlCost[9] + ttlCost[11] + ttlCost[13]) > 0){
		CString szField, szAmount;
		TranslateString(_T("Total Amount"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		for (int i =7; i < 16; i+=2)
		{
			
			szField.Format(_T("s%d"), i);
			tmpStr.Format(_T("%f"),ttlCost[i]);
			rptDetail->SetValue(szField, tmpStr);
			ttlCost[i] += grpCost[i];
		}
		
	}
	
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	//rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
	rpt.PrintPreview();
	EndWaitCursor();
}
void CPMSReportConditionDialog::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_szStockName = m_wndStock.GetCurrent(1);
	switch(m_nType)
	{
		case IMPORT_EXPORT:
			OnPrintImportExportReport(true);
			break;
		case IMPORT_EXPORT_INSTOCK:
			OnPrintImportExportAndInstockReport(true);
			break;
		case IMPORT_EXPORT_INSTOCK_DETAILFORMONEY:
			OnPrintImportExportAndInstockForDetailMoneyReport(true);
			break;
		case INSTOCK:
			OnPrintInstockReport(true);
			break;
		case IMPORT_EXPORT_INSTOCK_DETAIL_PROFESSIONAL:
			OnPrintTonghopnhapxuattonchitiettheonghiepvu(true);
			break;
		case THUOC_HET_HAN:
			OnPrint_Danhsachthuochethan(true);
			break;		
	}
} 
void CPMSReportConditionDialog::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
	UpdateData(TRUE);
	m_szStockName = m_wndStock.GetCurrent(1);

	switch(m_nType){
		case IMPORT_EXPORT:
			OnExcell_ImportExportReport();
			break;
		case IMPORT_EXPORT_INSTOCK:			
			OnExcell_ImportExportAndInstockReport();
			break;
		case IMPORT_EXPORT_INSTOCK_DETAILFORMONEY:
			OnExcell_ImportExportAndInstockForDetailMoneyReport();
			break;		
		case INSTOCK:
			OnExportInstockReport(true);
			break;
		case IMPORT_EXPORT_INSTOCK_DETAIL_PROFESSIONAL:
			OnExcell_Baocaotonghopnhapxuattonchitiettheonghiepvu();
			break;
		case THUOC_HET_HAN:
			OnExport_Danhsachthuochethan(true);
			break;		
	}
} 
void CPMSReportConditionDialog::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CPMSReportConditionDialog::OnAddPMSReportConditionDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CPMSReportConditionDialog::OnEditPMSReportConditionDialog(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CPMSReportConditionDialog::OnDeletePMSReportConditionDialog(){
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
 		OnCancelPMSReportConditionDialog(); 
 	}
	return 0;
}
int CPMSReportConditionDialog::OnSavePMSReportConditionDialog(){
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
	//_fmsg(_T("%s"), szSQL); 
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret > 0) 
 	{ 
 		//OnPMSReportConditionDialogListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 	
}
int CPMSReportConditionDialog::OnCancelPMSReportConditionDialog(){
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
int CPMSReportConditionDialog::OnPMSReportConditionDialogListLoadData(){
	return 0;
}
void CPMSReportConditionDialog::OnPrintImportExportReport(bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
    CString tmpStr, szDate, szSysDate, szSQL, szWhere, szTypeName, szType;
	
	if(!rpt.Init(_T("Reports/HMS/PMS_GENERALIMPORTEXPORTREPORT.RPT"),true) )
	{
		return;
	}

	BeginWaitCursor();

	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
			CDate::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	rpt.GetReportHeader()->SetValue(_T("StockName"), m_szStockName);
	rpt.GetReportHeader()->SetValue(_T("Type"), m_wndType.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Group"), m_wndGroup.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Original"), m_wndSource.GetCurrent(1));
			//Page Header
			//Report Detail
			
	 for (int i=0 ; i<= m_wndType.GetCount(); i++)
	 {
		 if(m_wndType.GetCheck(i)){
			 m_wndType.SetCurSel(i);
		 	if(!szType.IsEmpty())
					szType += _T(",");						
					szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));			
			if(!szTypeName.IsEmpty())
					szTypeName += _T(", ");						
			szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		 }
	 }
	
	 if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and pmi_typeid in(%s)"), szType);	
	else
		szTypeName.Format(_T("TÊt c¶ thuèc vµ vËt t­"));
	
	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(pmi_groupid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}

	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and pmsi_originalid=%d "), str2int(m_szSourceKey));
	}



	bool bHavePrice = false;
	if(rpt.GetDetail()->GetItem(_T("5")) != NULL)
		bHavePrice = true;

	if(bHavePrice)
	{
			szSQL.Format(_T("SELECT hfg_index as idx ,hfg_name, * FROM ") \
				_T("( SELECT ")\
				_T(" (SELECT pmi_typeid,") \
				_T("	pmi_id as id, ") \
				_T(" 	pmi_name as name, ") \
				_T(" 	pmi_unit as unit, ") \
				_T(" 	pmsi_vatprice as price,") \
				_T(" 	sum(impqty) as impqty, ") \
				_T(" 	sum(impqty*pmsi_vatprice) as impamount,") \
				_T(" 	sum(expqty) as expqty,") \
				_T(" 	sum(expqty*pmsi_vatprice) as expamount") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT impstockid as stockid, impdate as iodate, sitemid, impqty, 0 as expqty") \
				_T(" FROM pmsv_import") \
				_T(" UNION ALL") \
				_T(" SELECT impstockid, expdate, sitemid, 0, expqty") \
				_T(" FROM pmsv_export") \
				_T(" ) as instock") \
				_T(" LEFT JOIN pms_stockitems ON(pmsi_id=sitemid)") \
				_T(" LEFT JOIN pms_items ON(pmsi_itemid=pmi_id)") \
				_T(" WHERE 	stockid=%d ") \
				_T(" 	AND iodate between timestamp '%s' and timestamp '%s' ") \
				_T(" 	and sitemid > 0 %s") \
				_T(" GROUP BY pmi_typeid,pmi_groupid,id, name, unit, price") \
				_T(" ) as tbl LEFT JOIN hms_fee_group ON(hfg_id=pmi_typeid) ") \
				_T(" WHERE pmi_moduleid = '%s' AND length(id) > 0 and impqty+expqty > 0 ") \
				_T(" ORDER BY idx, pmi_typeid,name, unit"),pMF->GetModuleID(), ToInt(m_szStockKey), m_szFromDate, m_szToDate, szWhere);
	}
	else
	{
		szSQL.Format(_T("SELECT hfg_index as idx ,hfg_name , * FROM ") \
				_T("( SELECT ")\
				_T(" (SELECT pmi_typeid,") \
				_T("	pmi_id as id, ") \
				_T(" 	pmi_name as name, ") \
				_T(" 	pmi_unit as unit, ") \
				_T(" 	sum(impqty) as impqty, ") \
				_T(" 	sum(impqty*pmsi_vatprice) as impamount,") \
				_T(" 	sum(expqty) as expqty,") \
				_T(" 	sum(expqty*pmsi_vatprice) as expamount") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT impstockid as stockid, impdate as iodate, sitemid, impqty, 0 as expqty") \
				_T(" FROM pmsv_import") \
				_T(" UNION ALL") \
				_T(" SELECT impstockid, expdate, sitemid, 0, expqty") \
				_T(" FROM pmsv_export") \
				_T(" ) as instock") \
				_T(" LEFT JOIN pms_stockitems ON(pmsi_id=sitemid)") \
				_T(" LEFT JOIN pms_items ON(pmsi_itemid=pmi_id)") \
				_T(" WHERE 	stockid=%d ") \
				_T(" 	AND iodate between timestamp '%s' and timestamp '%s' ") \
				_T(" 	and sitemid > 0 %s") \
				_T(" GROUP BY pmi_typeid,pmi_groupid,id, name, unit") \
				_T(" ) as tbl LEFT JOIN hms_fee_group ON(hfg_id=pmi_typeid) ") \
				_T(" WHERE pmi_moduleid = '%s' AND length(id) > 0 and impqty+expqty > 0 ") \
				_T(" ORDER BY idx, pmi_typeid,name, unit"), pMF->GetModuleID(), ToInt(m_szStockKey), m_szFromDate, m_szToDate, szWhere);
	}
			//_fmsg(_T("%s"), szSQL);
			CReportSection* rptDetail;
			CString szField, szAmount;
			CRecord rs(&pMF->m_db);
			CString szOldLine, szNewLine, szOldGroup, szNewGroup;
			
			long double grpCost[23];
			long double ttlCost[23];
			double cost = 0;
			int nItem = 1;
			long nSum1, nSum2, nSum3, nSum4;
			nSum1 = nSum2 = nSum3 = nSum4 = 0;
			for (int i =0;i < 10; i++)
			{
				grpCost[i] = 0;
				ttlCost[i] = 0;
			}
			rs.ExecSQL(szSQL);
			if(rs.IsEOF()){
				ShowMessage(150, MB_ICONSTOP);
				return ;
			}
			
			while(!rs.IsEOF())
			{

				rs.GetValue(_T("pmi_typeid"), szNewLine);
				if(szNewLine != szOldLine && !szNewLine.IsEmpty()){				
				if((grpCost[7] + grpCost[9]) > 0)
				{
					TranslateString(_T("Total Group"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
					rptDetail->SetValue(_T("TotalGroup"), szAmount);				
					for (int i =7; i < 10; i++)
					{
						szField.Format(_T("s%d"), i);
						FormatCurrency(grpCost[i], tmpStr);
						rptDetail->SetValue(szField, tmpStr);
						ttlCost[i] += grpCost[i];						
						grpCost[i] = 0;				
					}
				}

				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				rs.GetValue(_T("pmi_typeid"), tmpStr);
				rptDetail->SetValue(_T("GroupName"),tmpStr );
				szOldLine = szNewLine;
				
			}
					

				rptDetail = rpt.AddDetail();
				tmpStr.Format(_T("%d"), nItem++);
				rptDetail->SetValue(_T("1"), tmpStr);
				rs.GetValue(_T("id"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rs.GetValue(_T("name"), tmpStr);
				rptDetail->SetValue(_T("3"), tmpStr);
				rs.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);				
				rs.GetValue(_T("price"), cost);
				FormatCurrency(cost, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				rs.GetValue(_T("impqty"), tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);				
				
				rs.GetValue(_T("impamount"), cost);				
				grpCost[7] += cost;
				FormatCurrency(cost, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);
				rs.GetValue(_T("expqty"), tmpStr);				
				rptDetail->SetValue(_T("8"), tmpStr);
				rs.GetValue(_T("expamount"), cost);
				grpCost[9] += cost;
				FormatCurrency(cost, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				rs.MoveNext();
			}

				for (int i =0; i < 10; i++)
			{
					ttlCost[i] += grpCost[i];
			}
			if((grpCost[7] + grpCost[9] ) > 0){				
				TranslateString(_T("Total Group"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("TotalGroup"), szAmount);				
				for (int i =7; i < 10; i++)
				{
					szField.Format(_T("s%d"), i);
					FormatCurrency(grpCost[i], tmpStr);
					rptDetail->SetValue(szField, tmpStr);
				}
			
			}
			if((ttlCost[7] + ttlCost[9] ) > 0){				
				TranslateString(_T("Total Amount"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("TotalGroup"), szAmount);
				for (int i =7; i < 10; i++)
				{
					
					szField.Format(_T("s%d"), i);
					FormatCurrency(ttlCost[i], tmpStr);
					rptDetail->SetValue(szField, tmpStr);					
				}
			}
		
			//Page Footer
			//Report Footer
			tmpStr.Format(_T("%ld"), nSum1);
			rpt.GetReportFooter()->SetValue(_T("Sum1"), tmpStr);
			tmpStr.Format(_T("%ld"), nSum2);
			rpt.GetReportFooter()->SetValue(_T("Sum2"), tmpStr);

			szSysDate = pMF->GetSysDate(); 
			szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
			rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
			
			//	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
			if(bPreview)
				rpt.PrintPreview();
			else
				rpt.Print();

	EndWaitCursor();
}
/*------------------------------------------------------------------------------------------------------------------------------
	Print Bao cao tong hop ton kho
------------------------------------------------------------------------------------------------------------------------------*/
void CPMSReportConditionDialog::OnPrintImportExportAndInstockReport(bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL, szWhere;
	CString szTypeName, szType, szReportName;
	szReportName.Format(_T("Reports/HMS/%sS_GENERALIMPORTEXPORTINSTOCKREPORT.RPT"), pMF->GetModuleID());

	if(!rpt.Init(szReportName),true)
	{
		return;
	}
		
	BeginWaitCursor();
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
	CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	rpt.GetReportHeader()->SetValue(_T("StockName"), m_szStockName);
	rpt.GetReportHeader()->SetValue(_T("Type"), m_wndType.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Group"), m_wndGroup.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Original"), m_wndSource.GetCurrent(1));
		
	 for (int i = 0 ; i <= m_wndType.GetCount(); i++)
	 {
		 if(m_wndType.GetCheck(i)){
			 m_wndType.SetCurSel(i);
		 	if(!szType.IsEmpty())
					szType += _T(",");						
					szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));			
			if(!szTypeName.IsEmpty())
					szTypeName += _T(", ");						
			szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		 }
	 }
	
	 if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and pmi_typeid in(%s)"), szType);	
	else
		szTypeName.Format(_T("TÊt c¶ thuèc vµ vËt t­"));
	
	if(!m_szGroupKey.IsEmpty())
	{
			szWhere.AppendFormat(_T(" and substr(pmi_groupid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}
	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and pmsi_originalid=%d "), str2int(m_szSourceKey));
	}
	bool bHavePrice = false;
	if(rpt.GetDetail()->GetItem(_T("5")) != NULL)
		bHavePrice = true;
	if(bHavePrice)
	{

			szSQL.Format(_T("SELECT hfg_index as idx, hfg_name as typename, * FROM") \
			_T(" (") \
			_T("  SELECT pmi_typeid, pmi_id as id, ") \
			_T(" 				 	pmi_name as name, ") \
			_T(" 				 	pmi_unit as unit, ") \
			_T(" 				 	pmsi_vatprice as price, ") \
			_T(" 				 	sum(periodqty) as periodqty,") \
			_T(" 				 	sum(periodqty*pmsi_vatprice) as periodamount,") \
			_T(" 				 	sum(impqty) as impqty, ") \
			_T(" 				 	sum(impqty*pmsi_vatprice) as impamount,") \
			_T(" 				 	sum(expqty) as expqty,") \
			_T(" 				 	sum(expqty*pmsi_vatprice) as expamount,") \
			_T(" 				 	sum(periodqty+impqty-expqty) as instockqty,") \
			_T(" 				 	sum((periodqty+impqty-expqty)*pmsi_vatprice) as instockamount") \
			_T(" 				 FROM") \
			_T(" 				 (") \
			_T(" 				 SELECT * FROM") \
			_T(" 				 (") \
			_T(" 				 	SELECT impstockid as stockid, impdate as iodate, sitemid, 0 as periodqty, impqty, 0 as expqty") \
			_T(" 				 	FROM pmsv_import") \
			_T(" 				 	UNION ALL") \
			_T(" 				 	SELECT impstockid, expdate, sitemid, 0, 0, expqty") \
			_T(" 				 	FROM pmsv_export") \
			_T(" 				 ) as iotbl") \
			_T(" 				 WHERE 	stockid=%d ") \
			_T(" 				 	AND iodate between timestamp '%s' and timestamp '%s' ") \
			_T(" 				 UNION ALL") \
			_T(" 				 SELECT stockid, iodate, sitemid, coalesce(impqty-expqty, 0), 0, 0 FROM") \
			_T(" 				 (") \
			_T(" 				 	SELECT impstockid as stockid, impdate as iodate, sitemid,  0 as periodqty, impqty, 0 as expqty") \
			_T(" 				 	FROM pmsv_import") \
			_T(" 				 	UNION ALL") \
			_T(" 				 	SELECT impstockid, expdate, sitemid, 0, 0, expqty") \
			_T(" 				 	FROM pmsv_export") \
			_T(" 				 ) as ptbl") \
			_T(" 				 WHERE 	stockid=%d ") \
			_T(" 				 	AND iodate < timestamp '%s' ") \
			_T(" 				 ) as instock") \
			_T(" 				 LEFT JOIN pms_stockitems ON(pmsi_id=sitemid)") \
			_T(" 				 LEFT JOIN pms_items ON(pmsi_itemid=pmi_id)") \
			_T(" 				WHERE sitemid > 0 %s ") \
			_T(" 				 GROUP BY pmi_typeid,id, name, unit, price") \
			_T(" ) as tbl LEFT JOIN hms_fee_group ON(hfg_id=pmi_typeid)") \
			_T(" WHERE length(id) > 0 and (impqty + expqty + instockqty > 0)") \
			_T(" ORDER BY idx, pmi_typeid, name, unit") , ToInt(m_szStockKey), m_szFromDate, m_szToDate, ToInt(m_szStockKey), m_szFromDate, szWhere);
	}
	else
	{

		szSQL.Format(_T("SELECT hfg_index as idx , hfg_name as pmi_typename, * FROM") \
			_T(" (") \
			_T("  SELECT pmi_typeid, pmi_id as id, ") \
			_T(" 				 	pmi_name as name, ") \
			_T(" 				 	pmi_unit as unit, ") \
			_T(" 				 	sum(periodqty) as periodqty,") \
			_T(" 				 	sum(periodqty*pmsi_vatprice) as periodamount,") \
			_T(" 				 	sum(impqty) as impqty, ") \
			_T(" 				 	sum(impqty*pmsi_vatprice) as impamount,") \
			_T(" 				 	sum(expqty) as expqty,") \
			_T(" 				 	sum(expqty*pmsi_vatprice) as expamount,") \
			_T(" 				 	sum(periodqty+impqty-expqty) as instockqty,") \
			_T(" 				 	sum((periodqty+impqty-expqty)*pmsi_vatprice) as instockamount") \
			_T(" 				 FROM") \
			_T(" 				 (") \
			_T(" 				 SELECT * FROM") \
			_T(" 				 (") \
			_T(" 				 	SELECT impstockid as stockid, impdate as iodate, sitemid, 0 as periodqty, impqty, 0 as expqty") \
			_T(" 				 	FROM pmsv_import") \
			_T(" 				 	UNION ALL") \
			_T(" 				 	SELECT impstockid, expdate, sitemid, 0, 0, expqty") \
			_T(" 				 	FROM pmsv_export") \
			_T(" 				 ) as iotbl") \
			_T(" 				 WHERE 	stockid=%d ") \
			_T(" 				 	AND iodate between timestamp '%s' and timestamp '%s' ") \
			_T(" 				 UNION ALL") \
			_T(" 				 SELECT stockid, iodate, sitemid, coalesce(impqty-expqty, 0), 0, 0 FROM") \
			_T(" 				 (") \
			_T(" 				 	SELECT impstockid as stockid, impdate as iodate, sitemid,  0 as periodqty, impqty, 0 as expqty") \
			_T(" 				 	FROM pmsv_import") \
			_T(" 				 	UNION ALL") \
			_T(" 				 	SELECT impstockid, expdate, sitemid, 0, 0, expqty") \
			_T(" 				 	FROM pmsv_export") \
			_T(" 				 ) as ptbl") \
			_T(" 				 WHERE 	stockid=%d ") \
			_T(" 				 	AND iodate < timestamp '%s' ") \
			_T(" 				 ) as instock") \
			_T(" 				 LEFT JOIN pms_stockitems ON(pmsi_id=sitemid)") \
			_T(" 				 LEFT JOIN pms_items ON(pmsi_itemid=pmi_id)") \
			_T(" 				WHERE sitemid > 0 %s ") \
			_T(" 				 GROUP BY pmi_typeid,id, name, unit") \
			_T(" ) as tbl LEFT JOIN hms_fee_group ON(hfg_id=pmi_typeid)") \
			_T(" WHERE length(id) > 0 and (impqty + expqty + instockqty > 0)") \
			_T(" ORDER BY idx, pmi_typeid,name, unit") , ToInt(m_szStockKey), m_szFromDate, m_szToDate, ToInt(m_szStockKey), m_szFromDate, szWhere);

	}
//_fmsg(_T("%s"), szSQL);
		
			CReportSection* rptDetail;
			CString szOldLine, szNewLine;
			CRecord rs(&pMF->m_db);
			long double grpCost[23];
			long double ttlCost[23];
			double cost = 0;
			int nItem = 1;

			for (int i =0; i < 23; i++)
			{
				grpCost[i] = 0;
				ttlCost[i] = 0;
			}
			

			rs.ExecSQL(szSQL);
			if(rs.IsEOF()){
				ShowMessage(150, MB_ICONSTOP);
				return ;
			}
			
			
			while(!rs.IsEOF())
			{
				rs.GetValue(_T("pmi_typeid"), szNewLine);
				if(szNewLine != szOldLine && !szNewLine.IsEmpty()){				
					rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
					rs.GetValue(_T("pmi_typename"), tmpStr);
					rptDetail->SetValue(_T("GroupName"),tmpStr );
					szOldLine = szNewLine;
				}

				rptDetail = rpt.AddDetail();
				tmpStr.Format(_T("%d"), nItem++);
				rptDetail->SetValue(_T("1"), tmpStr);
				rs.GetValue(_T("id"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rs.GetValue(_T("name"), tmpStr);
				rptDetail->SetValue(_T("3"), tmpStr);
				rs.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				rs.GetValue(_T("price"), cost);
				FormatCurrency(cost, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);

				rs.GetValue(_T("periodqty"), tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				rs.GetValue(_T("impqty"), tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				rs.GetValue(_T("expqty"), tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);
				rs.GetValue(_T("instockqty"), tmpStr);	
				rptDetail->SetValue(_T("9"), tmpStr);
				rs.MoveNext();

			}
			
			szSysDate = pMF->GetSysDate(); 
			szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
			rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
			//	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
			if(bPreview)
				rpt.PrintPreview();
			else
				rpt.Print();

			EndWaitCursor();

}

/*------------------------------------------------------------------------------------------------------------------------------
	Export To Excell Bao cao tong hop nhap xuat ton kho
------------------------------------------------------------------------------------------------------------------------------*/
void CPMSReportConditionDialog::OnExcell_ImportExportAndInstockReport(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr,szObjects,m_szFacList,tmpStr2, szDate, szSysDate, szSQL,szName,szAmount,szFromDate,szToDate, szWhere;
	
	CString szType, szTypeName;
	for (int i=0 ; i<= m_wndType.GetCount(); i++)
	 {
		 if(m_wndType.GetCheck(i)){
			 m_wndType.SetCurSel(i);
		 	if(!szType.IsEmpty())
					szType += _T(",");						
					szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));			
			if(!szTypeName.IsEmpty())
					szTypeName += _T(", ");						
			szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		 }
	 }
	
	 if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and pmi_typeid in(%s)"), szType);	
	else
		szTypeName.Format(_T("TÊt c¶ thuèc vµ vËt t­"));
	
	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(pmi_groupid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}

	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and pmsi_originalid=%d "), str2int(m_szSourceKey));
	}
	
szSQL.Format(_T("SELECT  hfg_index as idx,hfg_name as pmi_typename, * FROM") \
			_T(" (") \
			_T("  SELECT pmi_typeid, pmi_id as id, ") \
			_T(" 				 	pmi_name as name, ") \
			_T(" 				 	pmi_unit as unit, ") \
			_T(" 				 	pmsi_vatprice as price, ") \
			_T(" 				 	sum(periodqty) as periodqty,") \
			_T(" 				 	sum(periodqty*pmsi_vatprice) as periodamount,") \
			_T(" 				 	sum(impqty) as impqty, ") \
			_T(" 				 	sum(impqty*pmsi_vatprice) as impamount,") \
			_T(" 				 	sum(expqty) as expqty,") \
			_T(" 				 	sum(expqty*pmsi_vatprice) as expamount,") \
			_T(" 				 	sum(periodqty+impqty-expqty) as instockqty,") \
			_T(" 				 	sum((periodqty+impqty-expqty)*pmsi_vatprice) as instockamount") \
			_T(" 				 FROM") \
			_T(" 				 (") \
			_T(" 				 SELECT * FROM") \
			_T(" 				 (") \
			_T(" 				 	SELECT impstockid as stockid, impdate as iodate, sitemid, 0 as periodqty, impqty, 0 as expqty") \
			_T(" 				 	FROM pmsv_import") \
			_T(" 				 	UNION ALL") \
			_T(" 				 	SELECT impstockid, expdate, sitemid, 0, 0, expqty") \
			_T(" 				 	FROM pmsv_export") \
			_T(" 				 ) as iotbl") \
			_T(" 				 WHERE 	stockid=%d ") \
			_T(" 				 	AND iodate between timestamp '%s' and timestamp '%s' ") \
			_T(" 				 UNION ALL") \
			_T(" 				 SELECT stockid, iodate, sitemid, coalesce(impqty-expqty, 0), 0, 0 FROM") \
			_T(" 				 (") \
			_T(" 				 	SELECT impstockid as stockid, impdate as iodate, sitemid,  0 as periodqty, impqty, 0 as expqty") \
			_T(" 				 	FROM pmsv_import") \
			_T(" 				 	UNION ALL") \
			_T(" 				 	SELECT impstockid, expdate, sitemid, 0, 0, expqty") \
			_T(" 				 	FROM pmsv_export") \
			_T(" 				 ) as ptbl") \
			_T(" 				 WHERE 	stockid=%d ") \
			_T(" 				 	AND iodate < timestamp '%s' ") \
			_T(" 				 ) as instock") \
			_T(" 				 LEFT JOIN pms_stockitems ON(pmsi_id=sitemid)") \
			_T(" 				 LEFT JOIN pms_items ON(pmsi_itemid=pmi_id)") \
			_T(" 				WHERE sitemid > 0 %s ") \
			_T(" 				 GROUP BY pmi_typeid,id, name, unit, price") \
			_T(" ) as tbl") \
			_T(" LEFT JOIN hms_fee_group ON(hfg_id=pmi_typeid) ") \
			_T(" WHERE length(id) > 0 and (impqty + expqty + instockqty > 0)") \
			_T(" ORDER BY idx, pmi_typeid, name, unit") , ToInt(m_szStockKey), m_szFromDate, m_szToDate, ToInt(m_szStockKey), m_szFromDate, szWhere);

//_fmsg(_T("%s"), szSQL);
		
	CReportSection* rptDetail=NULL;
	CString szOldLine, szNewLine;
	CRecord rs(&pMF->m_db);	
	double cost = 0;
	int nItem = 1;
				

	rs.ExecSQL(szSQL);
	if(rs.IsEOF()){
		ShowMessage(150, MB_ICONSTOP);
		return ;
	}
	
			
	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 35);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 15);
	xls.SetColumnWidth(7, 15);
	xls.SetColumnWidth(8, 15);
	xls.SetColumnWidth(9, 15);
	xls.SetColumnWidth(10, 15);
	xls.SetColumnWidth(11, 15);
	xls.SetColumnWidth(12, 15);



	xls.SetCellMergedColumns(0,3,13);	
	xls.SetCellText(0, 3, _T("T\x1ED4NG H\x1EE2P NH\x1EACP XU\x1EA4T T\x1ED2N KHO"),FMT_TEXT|FMT_CENTER,true,14,0);

	xls.SetCellMergedColumns(0,4,13);
	tmpStr.Format(_T("T\x1EEB ngày %s \x111\x1EBFn ngày %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(0, 4, tmpStr,FMT_TEXT|FMT_CENTER,true,12,0);

	

	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
				CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));	
	

	int nRow = 5;
		tmpStr.Format(_T("T\x1EEB kho: %s"), m_wndStock.GetCurrent(1));
		xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);

		if(!m_szTypeKey.IsEmpty())
		{
			tmpStr.Format(_T("Lo\x1EA1i: %s"), m_wndType.GetCurrent(1));
			xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
		}
		if(!m_szGroupKey.IsEmpty())
		{
			tmpStr.Format(_T("Nhóm: %s"), m_wndGroup.GetCurrent(1));
			xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
		}
		if(!m_szSourceKey.IsEmpty())
		{
			tmpStr.Format(_T("Ngu\x1ED3n: %s"), m_wndSource.GetCurrent(1));
			xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
		}

	
	xls.SetCellText(0, nRow, _T("STT"));
	
	
	xls.SetCellText(1, nRow, _T("Mã thu\x1ED1\x63"), FMT_TEXT|FMT_CENTER, true);
	
	
	xls.SetCellText(2, nRow, _T("Tên thu\x1ED1\x63/ HL"), FMT_TEXT|FMT_CENTER, true);
	
	
	xls.SetCellText(3, nRow, _T("\x110\x1A1n v\x1ECB"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(4, nRow, _T("\x110\x1A1n giá"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(5, nRow, _T("T\x1ED3n \x111\x1EA7u k\x1EF3"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(6, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(7, nRow, _T("Nh\x1EADp trong k\x1EF3"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(8, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(9, nRow, _T("Xu\x1EA5t trong k\x1EF3"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(10, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(11, nRow, _T("T\x1ED3n \x63u\x1ED1i k\x1EF3"), FMT_TEXT|FMT_CENTER, true);	
	
	xls.SetCellText(12, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT|FMT_CENTER, true);
	
			
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("pmi_typeid"), szNewLine);
		if(szNewLine != szOldLine && !szNewLine.IsEmpty()){				
			nRow ++;
			CString szField, szAmount;			
			tmpStr.Format(_T("%s"), szNewLine);
			xls.SetCellMergedColumns(0,nRow,13);
			rs.GetValue(_T("pmi_typename"), tmpStr);
			xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);
			szOldLine = szNewLine;
			nItem=1;
		}

		nRow ++;		
		tmpStr.Format(_T("%d"), nItem++);
		xls.SetCellText(0,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("id"), tmpStr);
		xls.SetCellText(1,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("name"), tmpStr);
		xls.SetCellText(2,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("unit"), tmpStr);		
		xls.SetCellText(3,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("price"), tmpStr);
		xls.SetCellText(4,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("periodqty"), tmpStr);
		xls.SetCellText(5,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("periodamount"), tmpStr);
		xls.SetCellText(6,nRow,tmpStr, FMT_NUMBER1);				
		rs.GetValue(_T("impqty"), tmpStr);				
		xls.SetCellText(7,nRow,tmpStr, FMT_NUMBER1);				
		rs.GetValue(_T("impamount"), tmpStr);				
		xls.SetCellText(8,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("expqty"), tmpStr);				
		xls.SetCellText(9,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("expamount"), tmpStr);				
		xls.SetCellText(10,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("instockqty"), tmpStr);				
		xls.SetCellText(11,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("instockamount"), tmpStr);				
		xls.SetCellText(12,nRow,tmpStr, FMT_NUMBER1);
		rs.MoveNext();
	}
	
	xls.Save(_T("Exports\\Tonghopnhapxuattonkho.xls"));
}



/*------------------------------------------------------------------------------------------------------------------------------
	Export To Excell Bao cao tong hop nhap xuat ton kho
------------------------------------------------------------------------------------------------------------------------------*/
void CPMSReportConditionDialog::OnExcell_ImportExportReport(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr,szObjects,m_szFacList,tmpStr2, szDate, szSysDate, szSQL,szName,szAmount,szFromDate,szToDate, szWhere;
	BeginWaitCursor();
	CString szType, szTypeName;
	for (int i=0 ; i<= m_wndType.GetCount(); i++)
	 {
		 if(m_wndType.GetCheck(i)){
			 m_wndType.SetCurSel(i);
		 	if(!szType.IsEmpty())
					szType += _T(",");						
					szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));			
			if(!szTypeName.IsEmpty())
					szTypeName += _T(", ");						
			szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		 }
	 }
	
	 if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and pmi_typeid in(%s)"), szType);	
	else
		szTypeName.Format(_T("TÊt c¶ thuèc vµ vËt t­"));

	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(pmi_groupid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}

	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and pmsi_originalid=%d "), str2int(m_szSourceKey));
	}
	
szSQL.Format(_T("SELECT  hfg_index as idx,hfg_name as typename, * FROM") \
			_T(" (") \
			_T("  SELECT pmi_typeid, pmi_id as id, ") \
			_T(" 				 	pmi_name as name, ") \
			_T(" 				 	pmi_unit as unit, ") \
			_T(" 				 	pmsi_vatprice as price, ") \
			_T(" 				 	sum(periodqty) as periodqty,") \
			_T(" 				 	sum(periodqty*pmsi_vatprice) as periodamount,") \
			_T(" 				 	sum(impqty) as impqty, ") \
			_T(" 				 	sum(impqty*pmsi_vatprice) as impamount,") \
			_T(" 				 	sum(expqty) as expqty,") \
			_T(" 				 	sum(expqty*pmsi_vatprice) as expamount,") \
			_T(" 				 	sum(periodqty+impqty-expqty) as instockqty,") \
			_T(" 				 	sum((periodqty+impqty-expqty)*pmsi_vatprice) as instockamount") \
			_T(" 				 FROM") \
			_T(" 				 (") \
			_T(" 				 SELECT * FROM") \
			_T(" 				 (") \
			_T(" 				 	SELECT impstockid as stockid, impdate as iodate, sitemid, 0 as periodqty, impqty, 0 as expqty") \
			_T(" 				 	FROM pmsv_import") \
			_T(" 				 	UNION ALL") \
			_T(" 				 	SELECT impstockid, expdate, sitemid, 0, 0, expqty") \
			_T(" 				 	FROM pmsv_export") \
			_T(" 				 ) as iotbl") \
			_T(" 				 WHERE 	stockid=%d ") \
			_T(" 				 	AND iodate between timestamp '%s' and timestamp '%s' ") \
			_T(" 				 UNION ALL") \
			_T(" 				 SELECT stockid, iodate, sitemid, coalesce(impqty-expqty, 0), 0, 0 FROM") \
			_T(" 				 (") \
			_T(" 				 	SELECT impstockid as stockid, impdate as iodate, sitemid,  0 as periodqty, impqty, 0 as expqty") \
			_T(" 				 	FROM pmsv_import") \
			_T(" 				 	UNION ALL") \
			_T(" 				 	SELECT impstockid, expdate, sitemid, 0, 0, expqty") \
			_T(" 				 	FROM pmsv_export") \
			_T(" 				 ) as ptbl") \
			_T(" 				 WHERE 	stockid=%d ") \
			_T(" 				 	AND iodate < timestamp '%s' ") \
			_T(" 				 ) as instock") \
			_T(" 				 LEFT JOIN pms_stockitems ON(pmsi_id=sitemid)") \
			_T(" 				 LEFT JOIN pms_items ON(pmsi_itemid=pmi_id)") \
			_T(" 				WHERE sitemid > 0 %s ") \
			_T(" 				 GROUP BY pmi_typeid,id, name, unit, price") \
			_T(" ) as tbl") \
			_T(" LEFT JOIN hms_fee_group ON(hfg_id=pmi_typeid) ") \
			_T(" WHERE length(id) > 0 and (impqty + expqty + instockqty > 0)") \
			_T(" ORDER BY idx, pmi_typeid, name, unit") , ToInt(m_szStockKey), m_szFromDate, m_szToDate, ToInt(m_szStockKey), m_szFromDate, szWhere);

//_fmsg(_T("%s"), szSQL);
		
	CReportSection* rptDetail = NULL;
	CString szOldLine, szNewLine;
	CRecord rs(&pMF->m_db);	
	double cost = 0;
	int nItem = 1;
				

	rs.ExecSQL(szSQL);
	if(rs.IsEOF()){
		ShowMessage(150, MB_ICONSTOP);
		return ;
	}
	
			
	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 35);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 15);
	xls.SetColumnWidth(7, 15);
	xls.SetColumnWidth(8, 15);
	


	xls.SetCellMergedColumns(0,3,13);	
	xls.SetCellText(0, 3, _T("T\x1ED4NG H\x1EE2P NH\x1EACP XU\x1EA4T T\x1ED2N KHO"),FMT_TEXT|FMT_CENTER,true,14,0);

	xls.SetCellMergedColumns(0,4,13);
	tmpStr.Format(_T("T\x1EEB ngày %s \x111\x1EBFn ngày %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(0, 4, tmpStr,FMT_TEXT|FMT_CENTER,true,12,0);

	

	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
				CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));	
	

	int nRow = 5;
	tmpStr.Format(_T("T\x1EEB kho: %s"), m_wndStock.GetCurrent(1));
	xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);

	if(!m_szTypeKey.IsEmpty())
	{
		tmpStr.Format(_T("Lo\x1EA1i: %s"), m_wndType.GetCurrent(1));
		xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
	}
	if(!m_szGroupKey.IsEmpty())
	{
		tmpStr.Format(_T("Nhóm: %s"), m_wndGroup.GetCurrent(1));
		xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
	}
	if(!m_szSourceKey.IsEmpty())
	{
		tmpStr.Format(_T("Ngu\x1ED3n: %s"), m_wndSource.GetCurrent(1));
		xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
	}

	
	xls.SetCellText(0, nRow, _T("STT"));
	
	
	xls.SetCellText(1, nRow, _T("Mã thu\x1ED1\x63"), FMT_TEXT|FMT_CENTER, true);
	
	
	xls.SetCellText(2, nRow, _T("Tên thu\x1ED1\x63/ HL"), FMT_TEXT|FMT_CENTER, true);
	
	
	xls.SetCellText(3, nRow, _T("\x110\x1A1n v\x1ECB"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(4, nRow, _T("\x110\x1A1n giá"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(5, nRow, _T("Nh\x1EADp"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(6, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(7, nRow, _T("Xu\x1EA5t"), FMT_TEXT|FMT_CENTER, true);
	
	xls.SetCellText(8, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT|FMT_CENTER, true);
	
			
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("pmi_typeid"), szNewLine);
		if(szNewLine != szOldLine && !szNewLine.IsEmpty()){				
			nRow ++;
			CString szField, szAmount;			
			tmpStr.Format(_T("%s"), szNewLine);
			xls.SetCellMergedColumns(0,nRow,13);
			rs.GetValue(_T("catname"), tmpStr);
			xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);
			szOldLine = szNewLine;
			nItem=1;
		}

		nRow ++;		
		tmpStr.Format(_T("%d"), nItem++);
		xls.SetCellText(0,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("id"), tmpStr);
		xls.SetCellText(1,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("name"), tmpStr);
		xls.SetCellText(2,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("unit"), tmpStr);		
		xls.SetCellText(3,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("price"), tmpStr);
		xls.SetCellText(4,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("impqty"), tmpStr);				
		xls.SetCellText(5,nRow,tmpStr, FMT_NUMBER1);				
		rs.GetValue(_T("impamount"), tmpStr);				
		xls.SetCellText(6,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("expqty"), tmpStr);				
		xls.SetCellText(7,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("expamount"), tmpStr);				
		xls.SetCellText(8,nRow,tmpStr, FMT_NUMBER1);		
		rs.MoveNext();
	}
	
	xls.Save(_T("Exports\\Tonghopnhapxuat.xls"));
	EndWaitCursor();
}

void CPMSReportConditionDialog::OnPrintImportExportAndInstockForDetailMoneyReport(bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL, szWhere, szOrderby, szPrice, szCondition, szRptName, szStockSel;
	//szRptName.Format(_T("Reports/HMS/PMS_SUMMARYIMPORTEXPORTINSTOCKFORMONEYREPORT.RPT"));
	if (pMF->GetModuleID() == _T("PM"))
		szRptName.Format(_T("Reports/HMS/PM_BAOCAOXUATNHAPCHUNG.RPT"));
	else
		szRptName.Format(_T("Reports/HMS/MA_BAOCAOXUATNHAPCHUNG.RPT"));
	if(!rpt.Init(szRptName,true) )
	{
		return;
	}
	
	BeginWaitCursor();
	CString szType, szTypeName, szStock, szStockName;
	 for (int i=0 ; i<= m_wndStock.GetCount(); i++)
	 {
		 if(m_wndStock.GetCheck(i)){
			 m_wndStock.SetCurSel(i);
		 	if(!szStock.IsEmpty())
					szStock += _T(",");						
					szStock.AppendFormat(_T("%s"), m_wndStock.GetCurrent(0));			
			if(!szStockName.IsEmpty())
					szStockName += _T(", ");						
			szStockName.AppendFormat(_T("%s"), m_wndStock.GetCurrent(1));
		 }
	 }	
	if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and product_groupid in('%s')"), m_szTypeKey);	
	else
		szTypeName.Format(_T("TÊt c¶ thuèc vµ vËt t­"));

	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(product_rootid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}
	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and product_resourceid=%d "), str2int(m_szSourceKey));
	}
	szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	if (!szStock.IsEmpty())
		szStockSel.Format(_T(" AND stockid IN (%s)"), szStock);
	szOrderby.Format(_T("ORDER BY name"));
	if (m_szSort == _T("C"))
		szOrderby.Format(_T("ORDER BY idx1,name, unit"));
	else if (m_szSort == _T("T"))
		szOrderby.Format(_T("ORDER BY idx2,name, unit"));

	szCondition = _T("product_taxprice");
	if (m_bBillPrice)
		szCondition = _T("product_unitprice");
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
		CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	if (!m_szStockKey.IsEmpty())
		rpt.GetReportHeader()->SetValue(_T("StockName"), szStockName);
	else
		rpt.GetReportHeader()->SetValue(_T("StockName"), _T("To\xE0n \x62\x1ED9"));
	if (!m_szTypeKey.IsEmpty())
		rpt.GetReportHeader()->SetValue(_T("Type"), m_wndType.GetCurrent(1));
	if (!m_szGroupKey.IsEmpty())
		rpt.GetReportHeader()->SetValue(_T("Group"), m_wndGroup.GetCurrent(1));
	else
		rpt.GetReportHeader()->SetValue(_T("Group"), _T("To\xE0n \x62\x1ED9"));
	if (m_szSourceKey.IsEmpty())
		rpt.GetReportHeader()->SetValue(_T("Original"), m_wndSource.GetCurrent(1));

	//Page Header
	//Report Detail
	szSQL.Format(_T(" SELECT * FROM") \
				_T(" ( SELECT  product_rootid as idx1, product_rootname as groupname, product_code as id,        ") \
				_T("		   product_groupid as idx2, product_groupname as typename, product_classname as generic,") \
				_T("           product_name as name, get_uomname(product_purchase_uomid) as unit,       ") \
				_T("           %s as price,         ") \
				_T("           sum(periodqty) as periodqty,    ") \
				_T("           sum(periodqty*%s) as periodamount,   ") \
				_T("           sum(impqty) as impqty,") \
				_T("           sum(impqty*%s) as impamount,         ") \
				_T("           sum(expqty) as expqty,") \
				_T("           sum(expqty) *product_saleprice1 as expamount,        ") \
				_T("		   sum(expqty) *(product_saleprice1-%s) as realprofit,") \
				_T("           sum(periodqty+impqty-expqty) as instockqty,     ") \
				_T("           sum((periodqty+impqty-expqty))*%s as instockamount ") \
				_T("   FROM ( ") \
				_T("       SELECT * FROM (    ") \
				_T("           SELECT impstockid as stockid, impdate as iodate, sitemid, 0 as periodqty, impqty, 0 as expqty") \
				_T("           FROM miv        ") \
				_T("           UNION ALL") \
				_T("           SELECT expstockid, expdate, sitemid, 0, 0, expqty") \
				_T("           FROM mev ") \
				_T("           ) iotbl ") \
				_T("       WHERE iodate between cast_string2timestamp( '%s' )and cast_string2timestamp ('%s' ) %s") \
				_T("       UNION ALL ") \
				_T("       SELECT stockid, iodate, sitemid, coalesce(impqty-expqty, 0), 0, 0") \
				_T("       FROM (      ") \
				_T("           SELECT impstockid as stockid, impdate as iodate, sitemid,  0 as periodqty, impqty, 0 as expqty") \
				_T("           FROM miv") \
				_T("           UNION ALL       ") \
				_T("           SELECT expstockid, expdate, sitemid, 0, 0, expqty") \
				_T("           FROM mev ") \
				_T("           ) ptbl ") \
				_T("       WHERE iodate < cast_string2timestamp( '%s'  ) %s") \
				_T("       ) instock ") \
				_T("     LEFT JOIN m_productitem_view ON(product_item_id=sitemid)") \
				_T("     WHERE sitemid > 0 %s") \
				 _T("     GROUP BY product_rootid, product_rootname, product_groupid, product_groupname, product_code,") \
				_T("			  product_name, product_purchase_uomid, product_classname, %s, product_saleprice1 ") \
				_T(" ) tbl ") \
				_T(" WHERE (periodqty + impqty + expqty > 0)  ") \
				_T(" %s"), szCondition, szCondition, szCondition, szCondition, szCondition,
				m_szFromDate, m_szToDate, szStockSel, m_szFromDate, szStockSel, szWhere, szCondition, szOrderby);
	CReportSection* rptDetail;
	CString szOldLine, szNewLine;
	CRecord rs(&pMF->m_db);
	long double grpCost[23];
	long double ttlCost[23];
	double cost = 0;
	int nItem = 1;
	rs.ExecSQL(szSQL);

	if(rs.IsEOF()){
	//QueryOpener(szSQL);
	}

	long nTotalPeriod=0;
	long nTotalImport=0;
	long nTotalExport=0;
	long nTotalInStock=0;
	
	for (int i =0; i < 16; i++)
	{
		grpCost[i] = 0;
		ttlCost[i] = 0;
	}
	
	while(!rs.IsEOF())
	{
		if (m_szSort == _T("C"))
		{
			rs.GetValue(_T("idx1"), szNewLine);
			if(szNewLine != szOldLine && !szNewLine.IsEmpty()){
			CString szField;
			if((grpCost[7] + grpCost[9] + grpCost[11] + grpCost[13]) > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				for (int i =7; i < 16; i+=2)
				{
					szField.Format(_T("s%d"), i);
					tmpStr.Format(_T("%f"),grpCost[i]);
					rptDetail->SetValue(szField, tmpStr);
					ttlCost[i] += grpCost[i];						
					grpCost[i] = 0;				
				}
			}

			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			rs.GetValue(_T("groupname"), tmpStr);
				
			rptDetail->SetValue(_T("GroupName"),tmpStr );
			szOldLine = szNewLine;
			nItem=1;
			}
		}
		else if (m_szSort == _T("T"))
		{
			rs.GetValue(_T("idx2"), szNewLine);
			if(szNewLine != szOldLine && !szNewLine.IsEmpty()){
			CString szField;
			if((grpCost[7] + grpCost[9] + grpCost[11] + grpCost[13]) > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				for (int i =7; i < 16; i+=2)
				{
					szField.Format(_T("s%d"), i);
					tmpStr.Format(_T("%f"),grpCost[i]);
					rptDetail->SetValue(szField, tmpStr);
					ttlCost[i] += grpCost[i];						
					grpCost[i] = 0;				
				}
			}

			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			rs.GetValue(_T("typename"), tmpStr);
				
			rptDetail->SetValue(_T("GroupName"),tmpStr );
			szOldLine = szNewLine;
			nItem=1;
			}
		}
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nItem++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("name"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("generic"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("price"), cost);
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);			
		rs.GetValue(_T("periodqty"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rs.GetValue(_T("periodamount"), cost);	
		grpCost[7] += cost;
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("7"), cost);				

		rs.GetValue(_T("impqty"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.GetValue(_T("impamount"), cost);		
		grpCost[9] += cost;
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("9"), cost);


		rs.GetValue(_T("expqty"), tmpStr);				
		rptDetail->SetValue(_T("10"), tmpStr);
		rs.GetValue(_T("expamount"), cost);		
		grpCost[11] += cost;				
		rptDetail->SetValue(_T("11"), cost);
		
		rs.GetValue(_T("instockqty"), tmpStr);				
		rptDetail->SetValue(_T("12"), tmpStr);
		rs.GetValue(_T("instockamount"), cost);
		grpCost[13] += cost;				
		rptDetail->SetValue(_T("13"), cost);

		rs.GetValue(_T("realprofit"), cost);
		grpCost[15] += cost;				
		rptDetail->SetValue(_T("15"), cost);
		rs.MoveNext();

	}
	
	for (int i =0; i < 16; i++)
	{
			ttlCost[i] += grpCost[i];
	}
	if((grpCost[7] + grpCost[9] + grpCost[11] + grpCost[13]) > 0){
		CString szField;
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		for (int i =7; i < 16; i+=2)
		{
			szField.Format(_T("s%d"), i);
			tmpStr.Format(_T("%f"),grpCost[i]);
			rptDetail->SetValue(szField, tmpStr);
		}	
	}
	if((ttlCost[7] + ttlCost[9] + ttlCost[11] + ttlCost[13]) > 0){
		CString szField, szAmount;
		TranslateString(_T("Total Amount"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		for (int i =7; i < 16; i+=2)
		{
			
			szField.Format(_T("s%d"), i);
			tmpStr.Format(_T("%f"),ttlCost[i]);
			rptDetail->SetValue(szField, tmpStr);
			ttlCost[i] += grpCost[i];
		}
		
	}
	
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	//	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print();
	EndWaitCursor();
}
/*--------------------------------------------------------------------------------------------------------------------------------
Export To Excell Tong hop nhap xuat ton kho (chi tiet)
---------------------------------------------------------------------------------------------------------------------------------*/
void CPMSReportConditionDialog::OnExcell_ImportExportAndInstockForDetailMoneyReport(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CString tmpStr, szOrderby, szWhere, szSQL, szTemp, szCondition, szPrice, szStockSel;
	CString szType, szTypeName;
	for (int i=0 ; i<= m_wndType.GetCount(); i++)
	 {
		 if(m_wndType.GetCheck(i)){
			 m_wndType.SetCurSel(i);
		 	if(!szType.IsEmpty())
					szType += _T(",");						
					szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));			
			if(!szTypeName.IsEmpty())
					szTypeName += _T(", ");						
			szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		 }
	 }
	
	 if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and product_producttype in(%s)"), szType);	
	else
		szTypeName.Format(_T("TÊt c¶ thuèc vµ vËt t­"));

	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(product_rootid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}

	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and product_resourceid=%d "), str2int(m_szSourceKey));
	}
	szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	if (!m_szStockKey.IsEmpty())
		szStockSel.Format(_T(" AND stockid = %d"), ToInt(m_szStockKey));
	szOrderby.Format(_T("ORDER BY name"));
	if (m_szSort == _T("C"))
		szOrderby.Format(_T("ORDER BY idx1,name, unit"));
	else if (m_szSort == _T("T"))
		szOrderby.Format(_T("ORDER BY idx2,name, unit"));
	//szPrice = _T("product_taxprice as price");
	//szCondition = _T("product_taxprice");
	//if (m_bServPrice)
	//{
	//	szPrice = _T("product_saleprice1 as price");
	//	szCondition = _T("product_saleprice1");	
	//}
		szSQL.Format(_T(" SELECT * FROM") \
					_T(" ( SELECT  product_rootid as idx1, product_rootname as groupname, product_code as id,        ") \
					_T("		   product_producttype as idx2, product_typename as typename, product_classname as generic, ") \
					_T("           product_name as name, get_uomname(product_purchase_uomid) as unit,       ") \
					_T("           product_taxprice as price,         ") \
					_T("           sum(periodqty) as periodqty,    ") \
					_T("           sum(periodqty*product_taxprice) as periodamount,   ") \
					_T("           sum(impqty) as impqty,") \
					_T("           sum(impqty*product_taxprice) as impamount,         ") \
					_T("           sum(expqty) as expqty,") \
					_T("           sum(expqty) *product_saleprice1 as expamount,        ") \
					_T("		   sum(expqty) *(product_saleprice1-product_taxprice) as realprofit,") \
					_T("           sum(periodqty+impqty-expqty) as instockqty,     ") \
					_T("           sum((periodqty+impqty-expqty))*product_taxprice as instockamount ") \
					_T("   FROM ( ") \
					_T("       SELECT * FROM (    ") \
					_T("           SELECT impstockid as stockid, impdate as iodate, sitemid, 0 as periodqty, impqty, 0 as expqty") \
					_T("           FROM miv        ") \
					_T("           UNION ALL") \
					_T("           SELECT expstockid, expdate, sitemid, 0, 0, expqty") \
					_T("           FROM mev ") \
					_T("           ) iotbl ") \
					_T("       WHERE iodate between cast_string2timestamp( '%s' )and cast_string2timestamp ('%s' ) %s") \
					_T("       UNION ALL ") \
					_T("       SELECT stockid, iodate, sitemid, coalesce(impqty-expqty, 0), 0, 0") \
					_T("       FROM (      ") \
					_T("           SELECT impstockid as stockid, impdate as iodate, sitemid,  0 as periodqty, impqty, 0 as expqty") \
					_T("           FROM miv") \
					_T("           UNION ALL       ") \
					_T("           SELECT expstockid, expdate, sitemid, 0, 0, expqty") \
					_T("           FROM mev ") \
					_T("           ) ptbl ") \
					_T("       WHERE iodate < cast_string2timestamp( '%s'  ) %s") \
					_T("       ) instock ") \
					_T("     LEFT JOIN m_productitem_view ON(product_item_id=sitemid)") \
					_T("     WHERE sitemid > 0 %s") \
					_T("     GROUP BY product_rootid, product_rootname, product_producttype, product_typename, product_code, product_name, product_purchase_uomid, product_classname, product_taxprice, product_saleprice1 ") \
					_T(" ) tbl ") \
					_T(" WHERE (periodqty + impqty + expqty > 0)  ") \
					_T(" %s"), m_szFromDate, m_szToDate, szStockSel, m_szFromDate, szStockSel, szWhere, szOrderby);
	CString szOldLine, szNewLine;
	CRecord rs(&pMF->m_db);	
	double nCost = 0;
	long double nTotalPeriod = 0, nTotalImport = 0, nTotalExport = 0, nTotalInstock = 0, nTotalProfit = 0;
	long double nGrpPeriod = 0, nGrpImport = 0, nGrpExport = 0, nGrpInstock = 0, nGrpProfit = 0;
	int nItem = 1;
	BeginWaitCursor();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	CellFormat df(&xls), cf(&xls), hf(&xls);
	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER);
	cf.SetFontName(_T("Segoe UI"));

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 35);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 15);
	xls.SetColumnWidth(7, 15);
	xls.SetColumnWidth(8, 15);
	xls.SetColumnWidth(9, 15);
	xls.SetColumnWidth(10, 15);
	xls.SetColumnWidth(11, 15);
	xls.SetColumnWidth(12, 15);

	xls.SetCellMergedColumns(0, 0, 3);
	xls.SetCellMergedColumns(0, 1, 3);
	xls.SetCellText(0, 0, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(0, 2, 12);
	TranslateString(_T("General Import Export Instock For Money Report"), szTemp);
	StringUpper(szTemp, tmpStr);
	xls.SetCellText(0, 2, tmpStr,FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(0,3,12);
	tmpStr.Format(_T("T\x1EEB ngày %s \x111\x1EBFn ngày %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(0, 3, tmpStr, &df);

	int nRow = 3;
	if (!m_szStockKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("T\x1EEB kho: %s"), m_wndStock.GetCurrent(1));
		xls.SetCellMergedColumns(0, nRow, 12);
		xls.SetCellText(0, nRow, tmpStr, &cf);
	}
	if(!m_szTypeKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("Lo\x1EA1i: %s"), szTypeName);
		xls.SetCellMergedColumns(0, nRow, 12);
		xls.SetCellText(0, nRow, tmpStr, &cf);
	}
	if(!m_szGroupKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("Nhóm: %s"), m_wndGroup.GetCurrent(1));
		xls.SetCellMergedColumns(0, nRow, 12);
		xls.SetCellText(0, nRow, tmpStr, &cf);
	}
	if(!m_szSourceKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("Ngu\x1ED3n: %s"), m_wndSource.GetCurrent(1));
		xls.SetCellMergedColumns(0, nRow, 12);
		xls.SetCellText(0, nRow, tmpStr, &cf);
	}
	nRow++;
	xls.SetCellText(0, nRow, _T("STT"));	
	xls.SetCellText(1, nRow, _T("Ho\x1EA1t \x63h\x1EA5t"), FMT_TEXT|FMT_CENTER, true);	
	xls.SetCellText(2, nRow, _T("Tên thu\x1ED1\x63/ HL"), FMT_TEXT|FMT_CENTER, true);	
	xls.SetCellText(3, nRow, _T("\x110\x1A1n v\x1ECB"), FMT_TEXT|FMT_CENTER, true);
	xls.SetCellText(4, nRow, _T("\x110\x1A1n giá"), FMT_TEXT|FMT_CENTER, true);	
	xls.SetCellText(5, nRow, _T("T\x1ED3n \x111\x1EA7u k\x1EF3"), FMT_TEXT|FMT_CENTER, true);	
	xls.SetCellText(6, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT|FMT_CENTER, true);	
	xls.SetCellText(7, nRow, _T("Nh\x1EADp trong k\x1EF3"), FMT_TEXT|FMT_CENTER, true);	
	xls.SetCellText(8, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT|FMT_CENTER, true);	
	xls.SetCellText(9, nRow, _T("Xu\x1EA5t trong k\x1EF3"), FMT_TEXT|FMT_CENTER, true);	
	xls.SetCellText(10, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT|FMT_CENTER, true);	
	xls.SetCellText(11, nRow, _T("Th\x1EF1\x63 l\xE3i"), FMT_TEXT|FMT_CENTER, true);	
	xls.SetCellText(12, nRow, _T("T\x1ED3n \x63u\x1ED1i k\x1EF3"), FMT_TEXT|FMT_CENTER, true);		
	xls.SetCellText(13, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT|FMT_CENTER, true);
	nRow++;
			
	while(!rs.IsEOF())
	{
		if (m_szSort == _T("C"))
		{
			rs.GetValue(_T("idx1"), szNewLine);
			if(szNewLine != szOldLine && !szNewLine.IsEmpty()){	
				if ((nGrpExport + nGrpImport + nGrpInstock + nGrpPeriod)>0)
				{
					TranslateString(_T("Total Group"), tmpStr);
					xls.SetCellMergedColumns(0, nRow, 5);
					xls.SetCellText(0, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
					tmpStr.Format(_T("%f"), nGrpPeriod);
					xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER1, true);
					tmpStr.Format(_T("%f"), nGrpImport);
					xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1, true);
					tmpStr.Format(_T("%f"), nGrpExport);
					xls.SetCellText(10, nRow, tmpStr, FMT_NUMBER1, true);
					tmpStr.Format(_T("%f"), nGrpProfit);
					xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1, true);
					tmpStr.Format(_T("%f"), nGrpInstock);
					xls.SetCellText(13, nRow, tmpStr, FMT_NUMBER1, true);
					nTotalExport += nGrpExport;nGrpExport = 0;
					nTotalImport += nGrpImport;nGrpImport = 0;
					nTotalInstock += nGrpInstock;nGrpInstock = 0;
					nTotalProfit += nGrpProfit;nGrpProfit = 0;
					nTotalPeriod += nGrpPeriod;nGrpPeriod = 0;
					nRow++;
				}
				tmpStr.Format(_T("%s"), szNewLine);
				xls.SetCellMergedColumns(0,nRow,13);
				tmpStr.Format(_T("%s. %s"), szNewLine, rs.GetValue(_T("groupname")));
				xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);
				szOldLine = szNewLine;
				nItem=1;
				nRow++;
			}
		}	
		else if (m_szSort == _T("T"))
		{
			rs.GetValue(_T("idx2"), szNewLine);
			if(szNewLine != szOldLine && !szNewLine.IsEmpty()){	
				if ((nGrpExport + nGrpImport + nGrpInstock + nGrpPeriod)>0)
				{
					TranslateString(_T("Total Group"), tmpStr);
					xls.SetCellMergedColumns(0, nRow, 5);
					xls.SetCellText(0, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
					tmpStr.Format(_T("%f"), nGrpPeriod);
					xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER1, true);
					tmpStr.Format(_T("%f"), nGrpImport);
					xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1, true);
					tmpStr.Format(_T("%f"), nGrpExport);
					xls.SetCellText(10, nRow, tmpStr, FMT_NUMBER1, true);
					tmpStr.Format(_T("%f"), nGrpProfit);
					xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1, true);
					tmpStr.Format(_T("%f"), nGrpInstock);
					xls.SetCellText(13, nRow, tmpStr, FMT_NUMBER1, true);
					nTotalExport += nGrpExport;nGrpExport = 0;
					nTotalImport += nGrpImport;nGrpImport = 0;
					nTotalInstock += nGrpInstock;nGrpInstock = 0;
					nTotalProfit += nGrpProfit;nGrpProfit = 0;
					nTotalPeriod += nGrpPeriod;nGrpPeriod = 0;
					nRow++;
				}
				tmpStr.Format(_T("%s"), szNewLine);
				xls.SetCellMergedColumns(0,nRow,13);
				tmpStr.Format(_T("%s. %s"), szNewLine, rs.GetValue(_T("typename")));
				xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);
				szOldLine = szNewLine;
				nItem=1;
				nRow++;
			}
		}	
		tmpStr.Format(_T("%d"), nItem);
		xls.SetCellText(0,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("name"), tmpStr);
		xls.SetCellText(1,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("generic"), tmpStr);
		xls.SetCellText(2,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("unit"), tmpStr);		
		xls.SetCellText(3,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("price"), tmpStr);
		xls.SetCellText(4,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("periodqty"), tmpStr);
		xls.SetCellText(5,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("periodamount"), nCost);
		xls.SetCellText(6,nRow,double2str(nCost), FMT_NUMBER1);		
		nGrpPeriod += nCost;
		rs.GetValue(_T("impqty"), tmpStr);				
		xls.SetCellText(7,nRow,tmpStr, FMT_NUMBER1);				
		rs.GetValue(_T("impamount"), nCost);		
		xls.SetCellText(8,nRow,double2str(nCost), FMT_NUMBER1);
		nGrpImport += nCost;
		rs.GetValue(_T("expqty"), tmpStr);				
		xls.SetCellText(9,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("expamount"), nCost);				
		xls.SetCellText(10,nRow,double2str(nCost), FMT_NUMBER1);
		nGrpExport += nCost;
		rs.GetValue(_T("realprofit"), nCost);
		nGrpProfit += nCost;
		xls.SetCellText(11, nRow, double2str(nCost), FMT_NUMBER1);
		rs.GetValue(_T("instockqty"), tmpStr);				
		xls.SetCellText(12,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("instockamount"), nCost);				
		xls.SetCellText(13,nRow,double2str(nCost), FMT_NUMBER1);
		nGrpInstock += nCost;
		nRow++;
		nItem++;
		rs.MoveNext();
	}
	if (m_bCheck)	
		if ((nGrpExport + nGrpImport + nGrpInstock + nGrpPeriod)>0)
		{
			TranslateString(_T("Total Group"), tmpStr);
			xls.SetCellMergedColumns(0, nRow, 5);
			xls.SetCellText(0, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
			tmpStr.Format(_T("%f"), nGrpPeriod);
			xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER1, true);
			tmpStr.Format(_T("%f"), nGrpImport);
			xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1, true);
			tmpStr.Format(_T("%f"), nGrpExport);
			xls.SetCellText(10, nRow, tmpStr, FMT_NUMBER1, true);
			tmpStr.Format(_T("%f"), nGrpProfit);
			xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1, true);
			tmpStr.Format(_T("%f"), nGrpInstock);
			xls.SetCellText(13, nRow, tmpStr, FMT_NUMBER1, true);
			nRow++;
		}
	nTotalExport += nGrpExport;nGrpExport = 0;
	nTotalImport += nGrpImport;nGrpImport = 0;
	nTotalInstock += nGrpInstock;nGrpInstock = 0;
	nTotalProfit += nGrpProfit;nGrpProfit = 0;
	nTotalPeriod += nGrpPeriod;nGrpPeriod = 0;
	if ((nTotalPeriod+ nTotalExport+ nTotalImport+ nTotalInstock)>0)
	{
		TranslateString(_T("Total Amount"), tmpStr);
		xls.SetCellMergedColumns(0, nRow, 5);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
		tmpStr.Format(_T("%f"), nTotalPeriod);
		xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER1, true);
		tmpStr.Format(_T("%f"), nTotalImport);
		xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1, true);
		tmpStr.Format(_T("%f"), nTotalExport);
		xls.SetCellText(10, nRow, tmpStr, FMT_NUMBER1, true);
		tmpStr.Format(_T("%f"), nTotalProfit);
		xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1, true);
		tmpStr.Format(_T("%f"), nTotalInstock);
		xls.SetCellText(13, nRow, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\TONGHOPNHAPXUATTONKHOTHEOTONGTIEN.XLS"));
	EndWaitCursor();
}

/*--------------------------------------------------------------------------------------------------------------------------------
Print Bien ban kiem ke kho 
---------------------------------------------------------------------------------------------------------------------------------*/
void CPMSReportConditionDialog::OnPrintInstockReport(bool bPreview){	
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	UpdateData(true);
	if (m_szStockKey.IsEmpty())
	{
		ShowMessageBox(_T("Choose a stock first!"), MB_ICONSTOP);
		m_wndStock.SetFocus();
		return;
	}
	CReport rpt;
	CReportSection* rptDetail;
	CString szOldLine, szNewLine;
	CRecord rs(&pMF->m_db);
	CRecord mrs(&pMF->m_db);
	int nOldLocation = 0, nNewLocation = 0;
	long double grpCost = 0, ChestGrp = 0;
	long double ttlCost = 0;
	double cost = 0;
	long nItem = 1;
	CString tmpStr, szDate, szSysDate, szSQL, szWhere, szMemory, szName, szAmount, tmp, szPrice, szCondition, szOrderby, szRptName;
	if (pMF->GetModuleID() == _T("PM"))
		tmpStr.Format(_T("PMS_BIENBANKIEMKETHUOC.RPT"));
	else
		tmpStr.Format(_T("MA_BIENBANKIEMKETHUOC.RPT"));
	szRptName.Format(_T("Reports/HMS/%s"), tmpStr);
	if(!rpt.Init(szRptName) ){
			return;
	}
	

	BeginWaitCursor();
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy ),
	CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);	
	rpt.GetReportHeader()->SetValue(_T("Type"), m_wndType.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Group"), m_wndGroup.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Original"), m_wndSource.GetCurrent(1));

	//Page Header
	//Report Detail
	szSQL.Format(_T("SELECT * FROM m_member ORDER BY mm_id"));
	mrs.ExecSQL(szSQL);
	while(!mrs.IsEOF()){
		int n;
		mrs.GetValue(_T("mm_id"), n);
		szName.Format(_T("Member%d"), n);
		mrs.GetValue(_T("mm_name"), tmpStr);
		if (!tmpStr.IsEmpty())
		{
			tmp.Format(_T("%d. %s"),n, tmpStr);
			rpt.GetReportHeader()->SetValue(szName, tmp);
			szName.Format(_T("Title%d"), n);
			mrs.GetValue(_T("mm_title"), tmpStr);
			rpt.GetReportHeader()->SetValue(szName, tmpStr);
		}
		mrs.MoveNext();
	}
	
	CString m_szStock;
	m_szStock = m_szStockKey;
	//m_szStockName.Empty();
	//for (int i=0 ; i<= m_wndStock.GetCount(); i++)
	//{
	//	 if(m_wndStock.GetCheck(i))
	//	 {
	//		 m_wndStock.SetCurSel(i);
	//		 if(!m_szStock.IsEmpty())
	//		 {
	//			m_szStock += _T(",");
	//			m_szStockName +=_T(", ");

	//		 }
	//		 m_szStock.AppendFormat(_T("%s"), m_wndStock.GetCurrent(0));				
	//		 m_szStockName.AppendFormat(_T("%s"), m_wndStock.GetCurrent(1));
	//	}		
	// }
	CString szType, szTypeName;
	for (int i=0 ; i<= m_wndType.GetCount(); i++)
	 {
		 if(m_wndType.GetCheck(i)){
			 m_wndType.SetCurSel(i);
		 	if(!szType.IsEmpty())
					szType += _T(",");						
					szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));			
			if(!szTypeName.IsEmpty())
					szTypeName += _T(", ");						
			szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		 }
	 }
	
	 
	if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and mp_producttype in(%s)"), szType);	
	else
		szTypeName.Format(_T("TÊt c¶ thuèc vµ vËt t­"));

	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(product_categoryid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}

	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and product_resourceid=%d "), str2int(m_szSourceKey));
	}
	if (!m_szSupplierKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND product_bpartnerid = '%s'"), m_szSupplierKey);
	szWhere.AppendFormat(_T(" and product_org_id = '%s'"), pMF->GetModuleID());
	szPrice = _T("product_taxprice as price");
	szCondition = _T("product_taxprice");
	if (m_bServPrice)
	{
		szPrice = _T("product_saleprice1 as price");
		szCondition = _T("product_saleprice1");
	}
	szOrderby.Format(_T("ORDER BY location, name"));
	if (m_szSort == _T("C"))
		szOrderby.Format(_T("ORDER BY location, idx1,name, unit"));
	else if (m_szSort == _T("T"))
		szOrderby.Format(_T("ORDER BY location, idx2,name, unit"));
	szSQL.Format(_T(" SELECT") \
	_T("   *") \
	_T(" FROM   (SELECT") \
	_T("           product_categoryid AS idx1,") \
	_T("		   product_producttype AS idx2,") \
	_T("		   product_typename AS typename,") \
	_T("           product_categoryname AS catname,") \
	_T("		   product_code as id,") \
	_T("           product_name AS name,") \
	_T("           product_uomname AS unit,") \
	_T("           product_lotno                                                       AS lotno,") \
	_T("           product_expdate                                                     AS exdate,") \
	_T("           product_countryname AS countryid,") \
	_T("		   get_locationname(%s, mpl_location_id) locationame,") \
	_T("		   coalesce(mpl_location_id, 999) location,") \
	_T("           %s,") \
	_T("           sum(periodqty)                                    AS instockqty,") \
	_T("           sum((periodqty)*%s)                     AS instockamount") \
	_T("         FROM   (SELECT") \
	_T("                   stockid,") \
	_T("                   iodate,") \
	_T("                   sitemid,") \
	_T("                   COALESCE(impqty-expqty, 0) periodqty") \
	_T("                 FROM   (SELECT") \
	_T("                           impstockid AS stockid,") \
	_T("                           impdate    AS iodate,") \
	_T("                           sitemid,") \
	_T("                           0          AS periodqty,") \
	_T("                           impqty,") \
	_T("                           0          AS expqty,") \
	_T("                           0") \
	_T("                         FROM   miv") \
	_T("                         UNION ALL") \
	_T("                         SELECT") \
	_T("                           expstockid,") \
	_T("                           expdate,") \
	_T("                           sitemid,") \
	_T("                           0,") \
	_T("                           0,") \
	_T("                           expqty,") \
	_T("                           0") \
	_T("                         FROM   mev) ptbl") \
	_T("                 WHERE  stockid IN(%s)") \
	_T("                        AND iodate<Cast_string2timestamp ('%s')) instock") \
	_T("                LEFT JOIN m_productitem_view ON(product_item_id=sitemid)") \
	_T("				LEFT JOIN m_product_location ON (mpl_storage_id = %s AND mpl_product_id = product_id)") \
	_T("         WHERE  sitemid>0 %s") \
	_T("         GROUP  BY product_categoryid,") \
	_T("                   product_categoryname,") \
	_T("				   product_producttype,") \
	_T("				   product_typename,") \
	_T("                   product_name,") \
	_T("				   product_code,") \
	_T("                   product_uomname,") \
	_T("                   product_lotno,") \
	_T("                   product_expdate,") \
	_T("                   product_countryname,") \
	_T("				   mpl_location_id,") \
	_T("                   product_taxprice, %s) tbl") \
	_T(" WHERE  instockqty>0 AND id IS NOT NULL %s"), m_szStock, szPrice, szCondition, m_szStock, 
	m_szToDate, m_szStock, szWhere, szCondition, szOrderby);
	_fmsg(_T("%s"), szSQL);
		rs.ExecSQL(szSQL);
		if(rs.IsEOF()){
			ShowMessage(150, MB_ICONSTOP);
			return ;
		}
	
		grpCost = 0;
		ttlCost = 0;
		long nTotalOrder = 0;
		
		while(!rs.IsEOF())
		{
			rs.GetValue(_T("location"), nNewLocation);
			if (nNewLocation != nOldLocation)
			{
				if (grpCost > 0)
				{
					if (m_szSort != _T("N"))
					{
						TranslateString(_T("Total Group"), szAmount);
						rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
						rptDetail->SetValue(_T("TotalGroup"), szAmount);						
						FormatCurrency(grpCost, tmpStr);
						rptDetail->SetValue(_T("s9"), tmpStr);
						
					}
					ChestGrp += grpCost;
					grpCost = 0;
				}
				if(ChestGrp > 0)
				{

					TranslateString(_T("Total Chest"), szAmount);
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
					rptDetail->SetValue(_T("TotalGroup"), szAmount);						
					FormatCurrency(ChestGrp, tmpStr);
					rptDetail->SetValue(_T("s9"), tmpStr);
					ttlCost += ChestGrp;
					ChestGrp = 0;
					nTotalOrder += nItem - 1;
					nItem=1;
				}
				if (nOldLocation > 0  && rptDetail)
				{
					rptDetail->SetPageBreak();
				}
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				rs.GetValue(_T("locationame"), tmpStr);
				rptDetail->SetValue(_T("GroupName"),tmpStr );	
				nOldLocation = nNewLocation;

			}

			if (m_szSort == _T("C"))
			{
				rs.GetValue(_T("idx1"), szNewLine);
				if(szNewLine != szOldLine && !szNewLine.IsEmpty())
				{	
					if(grpCost > 0)
					{
						TranslateString(_T("Total Group"), szAmount);
						rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
						rptDetail->SetValue(_T("TotalGroup"), szAmount);						
						FormatCurrency(grpCost, tmpStr);
						rptDetail->SetValue(_T("s9"), tmpStr);
						ChestGrp += grpCost;
						grpCost = 0;
						nTotalOrder += nItem - 1;
						nItem=1;
					}
					rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
					rs.GetValue(_T("catname"), tmpStr);
					rptDetail->SetValue(_T("GroupName"),tmpStr );
					szOldLine = szNewLine;

				}
			}
			else if (m_szSort == _T("T"))
			{
				rs.GetValue(_T("idx2"), szNewLine);
				if(szNewLine != szOldLine && !szNewLine.IsEmpty())
				{	
					if(grpCost > 0)
					{
						TranslateString(_T("Total Group"), szAmount);
						rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
						rptDetail->SetValue(_T("TotalGroup"), szAmount);						
						FormatCurrency(grpCost, tmpStr);
						rptDetail->SetValue(_T("s9"), tmpStr);
						ChestGrp += grpCost;
						grpCost = 0;
						nTotalOrder += nItem - 1;
						nItem=1;
					}
					rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
					rs.GetValue(_T("typename"), tmpStr);
					rptDetail->SetValue(_T("GroupName"),tmpStr );
					szOldLine = szNewLine;
					
				}
			}
			else
				nTotalOrder = nItem;
			rptDetail = rpt.AddDetail();
			tmpStr.Format(_T("%d"), nItem++);
			rptDetail->SetValue(_T("1"), tmpStr);
			rs.GetValue(_T("id"), tmpStr);
			rptDetail->SetValue(_T("2"), tmpStr);
			rs.GetValue(_T("name"), tmpStr);
			rptDetail->SetValue(_T("3"), tmpStr);
			rs.GetValue(_T("unit"), tmpStr);
			rptDetail->SetValue(_T("4"), tmpStr);
			rs.GetValue(_T("price"), cost);
			FormatCurrency(cost, tmpStr);
			rptDetail->SetValue(_T("5"), tmpStr);
			rs.GetValue(_T("countryid"), tmpStr);
			rptDetail->SetValue(_T("6"), tmpStr);
			rs.GetValue(_T("exdate"), tmpStr);
			rptDetail->SetValue(_T("7"), CDate::Convert(tmpStr,yyyymmdd,ddmmyyyy));
			if (!m_bCheck)
			{
				rs.GetValue(_T("instockqty"), tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);
				rs.GetValue(_T("instockamount"), cost);
				grpCost += cost;
				FormatCurrency(cost, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
			}
			rs.GetValue(_T("lotno"), tmpStr);
			rptDetail->SetValue(_T("10"), tmpStr);
			rs.GetValue(_T("country"), tmpStr);
			rptDetail->SetValue(_T("11"), tmpStr);
			rs.GetValue(_T("ortherqty"), tmpStr);
			rptDetail->SetValue(_T("12"),tmpStr);
			rs.MoveNext();
		}
		if (m_szSort != _T("N"))
			nTotalOrder += nItem - 1;
		
		if (grpCost > 0)
		{
			if (m_szSort != _T("N"))
			{
				TranslateString(_T("Total Group"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				rptDetail->SetValue(_T("TotalGroup"), szAmount);						
				FormatCurrency(grpCost, tmpStr);
				rptDetail->SetValue(_T("s9"), tmpStr);
				
			}
			ChestGrp += grpCost;
			grpCost = 0;
		}
		if(ChestGrp > 0)
		{

			TranslateString(_T("Total Chest"), szAmount);
			rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
			rptDetail->SetValue(_T("TotalGroup"), szAmount);						
			FormatCurrency(ChestGrp, tmpStr);
			rptDetail->SetValue(_T("s9"), tmpStr);
			ttlCost += ChestGrp;
			ChestGrp = 0;
			nTotalOrder += nItem - 1;
		}
		if(ttlCost > 0)
		{				
			TranslateString(_T("Total Amount"), szAmount);
			rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
			rptDetail->SetValue(_T("TotalGroup"), szAmount);
			FormatCurrency(ttlCost, tmpStr);
			rptDetail->SetValue(_T("s9"), tmpStr);				
		}
		
		szSysDate = pMF->GetSysDate();
		rpt.GetReportFooter()->SetValue(_T("TotalItem"), int2str(nTotalOrder));
		szMemory.Format(_T("%.2f"), ttlCost);
		rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
		MoneyToString(szMemory, tmpStr);
		if (!tmpStr.IsEmpty())
		{
			tmpStr += _T(" \x111\x1ED3ng");
			rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
		}

		rpt.GetReportHeader()->SetValue(_T("StockName"), m_szStockName);

		//if (pMF->m_szModuleID == _T("IN"))
		//{
		//	tmpStr.Format(_T("%ld"), nTotalOrder);
		//	rpt.GetReportFooter()->SetValue(_T("TotalOrder"), tmpStr);
		//}
		
		szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
		rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);			
		if(bPreview)
			rpt.PrintPreview();
		else
			rpt.Print();

	EndWaitCursor();

}
/*--------------------------------------------------------------------------------------------------------------------------------
Export To Excell Bien ban kiem ke kho 
---------------------------------------------------------------------------------------------------------------------------------*/
void CPMSReportConditionDialog::OnExportInstockReport(bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();	
	UpdateData(true);
	if (m_szStockKey.IsEmpty())
	{
		ShowMessageBox(_T("Choose a stock first!"), MB_ICONSTOP);
		m_wndStock.SetFocus();
		return;
	}
	CString szOldLine, szNewLine;
	CRecord rs(&pMF->m_db);
	CRecord mrs(&pMF->m_db);
	long double grpCost;
	long double ttlCost;
	long qty=0;
	double cost = 0;
	int nItem = 1;
	CString tmpStr, szDate, szSysDate, szSQL,szName,szAmount,szFromDate,szToDate, szWhere, szType, szTypeName, szPrice, szCondition, szOrderBy;
	BeginWaitCursor();
	CString m_szStock;
	m_szStockName.Empty();
	m_szStock = m_szStockKey;

	 for (int i=0 ; i<= m_wndType.GetCount(); i++)
	 {
		 if(m_wndType.GetCheck(i))
		 {
			m_wndType.SetCurSel(i);
		 	if(!szType.IsEmpty())
					szType += _T(", ");						
					szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));			
			if(!szTypeName.IsEmpty())
					szTypeName += _T(", ");						
			szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		 }
	 }
	
	if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and mp_producttype in(%s)"), szType);
	
	
	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(mp_product_category_id, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}
	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and mpi_resource_id=%d "), str2int(m_szSourceKey));
	}
	if (!m_szSupplierKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND product_bpartnerid = '%s'"), m_szSupplierKey);
	szWhere.AppendFormat(_T(" and product_org_id = '%s'"), pMF->GetModuleID());
	szPrice = _T("product_taxprice as price");
	szCondition = _T("product_taxprice");
	if (m_bServPrice)
	{
		szPrice = _T("product_saleprice1 as price");
		szCondition = _T("product_saleprice1");
	}
	szOrderBy.Format(_T("ORDER BY name"));
	if (m_szSort == _T("C"))
		szOrderBy.Format(_T("ORDER BY idx1,name, unit"));
	else if (m_szSort == _T("T"))
		szOrderBy.Format(_T("ORDER BY idx2,name, unit"));
	szSQL.Format(_T(" SELECT") \
	_T("   *") \
	_T(" FROM   (SELECT") \
	_T("           product_categoryid                                         AS idx1,") \
	_T("           product_categoryname                                                        AS catname,") \
	_T("		   product_producttype AS idx2,") \
	_T("		   product_typename AS typename,") \
	_T("		   product_code as id,") \
	_T("           product_name                                                         AS name,") \
	_T("           Get_uomname(product_purchase_uomid)                                          AS unit,") \
	_T("           product_lotno                                                       AS lotno,") \
	_T("           product_expdate                                                     AS exdate,") \
	_T("           Get_countryname((SELECT") \
	_T("                              DISTINCT") \
	_T("                              adm_country_id") \
	_T("                            FROM   ad_manufacture") \
	_T("                            WHERE  adm_manufacture_id=product_manufactureid)) AS countryid,") \
	_T("           %s                                                    ,") \
	_T("           Sum(periodqty)                                    AS instockqty,") \
	_T("           Sum((periodqty)*%s)                     AS instockamount") \
	_T("         FROM   (SELECT") \
	_T("                   stockid,") \
	_T("                   iodate,") \
	_T("                   sitemid,") \
	_T("                   COALESCE(impqty-expqty, 0) periodqty") \
	_T("                 FROM   (SELECT") \
	_T("                           impstockid AS stockid,") \
	_T("                           impdate    AS iodate,") \
	_T("                           sitemid,") \
	_T("                           0          AS periodqty,") \
	_T("                           impqty,") \
	_T("                           0          AS expqty,") \
	_T("                           0") \
	_T("                         FROM   miv") \
	_T("                         UNION ALL") \
	_T("                         SELECT") \
	_T("                           expstockid,") \
	_T("                           expdate,") \
	_T("                           sitemid,") \
	_T("                           0,") \
	_T("                           0,") \
	_T("                           expqty,") \
	_T("                           0") \
	_T("                         FROM   mev) ptbl") \
	_T("                 WHERE  stockid IN(%s)") \
	_T("                        AND iodate<Cast_string2timestamp ('%s')) instock") \
	_T("		 LEFT JOIN m_productitem_view ON (sitemid = product_item_id)") \
	_T("         WHERE  sitemid>0 %s") \
	_T("         GROUP  BY product_categoryid,") \
	_T("                   product_categoryname,") \
	_T("				   product_producttype,") \
	_T("				   product_typename,") \
	_T("                   product_name,") \
	_T("				   product_code,") \
	_T("                   product_purchase_uomid,") \
	_T("                   product_lotno,") \
	_T("                   product_expdate,") \
	_T("                   product_manufactureid,") \
	_T("                   product_taxprice, %s) tbl") \
	_T(" WHERE  instockqty>0 AND id IS NOT NULL %s"), szPrice, szCondition, m_szStock, m_szFromDate, szWhere, szCondition, szOrderBy);
	//QueryOpener(szSQL);
	BeginWaitCursor();

	rs.ExecSQL(szSQL);
	if(rs.IsEOF()){
		ShowMessage(150, MB_ICONSTOP);
		return ;
	}
			

	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	CellFormat df(&xls), cf(&xls);
	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT| FMT_CENTER);
	cf.SetFontName(_T("Segoe UI"));
	cf.SetCellStyle(FMT_TEXT);
	xls.SetColumnWidth(0,5);
	xls.SetColumnWidth(1,30);
	xls.SetColumnWidth(2,20);
	xls.SetColumnWidth(3,10);
	xls.SetColumnWidth(4,10);
	xls.SetColumnWidth(5,10);
	xls.SetColumnWidth(6,12);
	xls.SetColumnWidth(7,12);
	xls.SetColumnWidth(8,15);
	xls.SetColumnWidth(9,12);
	xls.SetColumnWidth(10,12);
	xls.SetColumnWidth(11,12);
	xls.SetColumnWidth(12,15);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellText(0, 0, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(0,2,15);	
	xls.SetCellText(0, 2, _T("\x42\xC1O \x43\xC1O KI\x1EC2M K\xCA KHO THU\x1ED0\x43 V\xC2T T\x1AF"),FMT_TEXT|FMT_CENTER,true, 12);
	xls.SetCellMergedColumns(0,3,15);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(0, 3, tmpStr, &df);
	

	int nRow = 4;
	tmpStr.Format(_T("T\x1EEB kho: %s"), m_szStockName);
	xls.SetCellMergedColumns(0, nRow, 15);
	xls.SetCellText(0, nRow, tmpStr,&cf);

	if(!m_szTypeKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("Lo\x1EA1i: %s"), m_wndType.GetCurrent(1));
		xls.SetCellMergedColumns(0, nRow, 14);
		xls.SetCellText(0, nRow, tmpStr,&cf);
	}
	if(!m_szGroupKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("Nh\xF3m: %s"), m_wndGroup.GetCurrent(1));
		xls.SetCellMergedColumns(0, nRow, 14);
		xls.SetCellText(0, nRow, tmpStr,&cf);
	}
	if(!m_szSourceKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("Ngu\x1ED3n thu\x1ED1\x63: %s"), m_wndSource.GetCurrent(1));
		xls.SetCellMergedColumns(0, nRow, 14);
		xls.SetCellText(0, nRow, tmpStr,&cf);
	}
	nRow++;
	xls.SetCellText(0, nRow, _T("STT"),FMT_TEXT,true);
	xls.SetCellText(1, nRow, _T("T\xEAn thu\x1ED1\x63-H\xE0m l\x1B0\x1EE3ng"),FMT_TEXT,true);
	xls.SetCellText(2, nRow, _T("T\xEAn ho\x1EA1t \x63h\x1EA5t"),FMT_TEXT,true);
	xls.SetCellText(3, nRow, _T("\x110\x1A1n v\x1ECB"),FMT_TEXT,true);
	xls.SetCellText(4, nRow, _T("H\xE3ng SX"),FMT_TEXT,true);
	xls.SetCellText(5, nRow, _T("H\x1EA1n \x64\xF9ng"),FMT_TEXT,true);
	xls.SetCellText(6, nRow, _T("\x110\x1A1n gi\xE1"),FMT_TEXT,true);
	xls.SetCellText(7, nRow, _T("S\x1ED1 l\x1B0\x1EE3ng"),FMT_TEXT,true);
	xls.SetCellText(8, nRow, _T("Th\xE0nh ti\x1EC1n"),FMT_TEXT,true);	
	xls.SetCellText(9, nRow, _T("SL Th\x1EEB\x61"),FMT_TEXT,true);
	xls.SetCellText(10, nRow, _T("SL Thi\x1EBFu"),FMT_TEXT,true);
	xls.SetCellText(11, nRow, _T("V\x1EE1 h\x1ECFng"),FMT_TEXT,true);
	xls.SetCellText(12, nRow, _T("Ghi \x63h\xFA"),FMT_TEXT,true);
	xls.SetCellText(13, nRow, _T("S\x1ED1 l\xF4"),FMT_TEXT,true);
	xls.SetCellText(14, nRow, _T("N\x1B0\x1EDB\x63 SX"),FMT_TEXT,true);

	grpCost = 0;
	ttlCost = 0;
			
	while(!rs.IsEOF())
	{
		if (m_szSort == _T("C"))
		{
			rs.GetValue(_T("idx1"), szNewLine);
			if(szNewLine != szOldLine && !szNewLine.IsEmpty()){	
				if(grpCost > 0)
				{			
					CString szField, szAmount;
					nRow ++;
					TranslateString(_T("\x43\x1ED9ng:"), szAmount);												
					xls.SetCellMergedColumns(1,nRow,7);
					xls.SetCellText(1, nRow, szAmount, FMT_TEXT,true,11,0);
					tmpStr.Format(_T("%.2f"), grpCost);
					xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1,true,11,0);
					ttlCost += grpCost;
					grpCost = 0;								
					nItem=1;
				}
				nRow ++;
				CString szField, szAmount;			
				xls.SetCellMergedColumns(0,nRow,12);
				rs.GetValue(_T("catname"), tmpStr);
				xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);
				szOldLine = szNewLine;
				nItem=1;
			}
		}
		else if (m_szSort == _T("T"))
		{
			rs.GetValue(_T("idx2"), szNewLine);
			if(szNewLine != szOldLine && !szNewLine.IsEmpty()){	
				if(grpCost > 0)
				{			
					CString szField, szAmount;
					nRow ++;
					TranslateString(_T("\x43\x1ED9ng:"), szAmount);												
					xls.SetCellMergedColumns(1,nRow,7);
					xls.SetCellText(1, nRow, szAmount, FMT_TEXT,true,11,0);
					tmpStr.Format(_T("%.2f"), grpCost);
					xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1,true,11,0);
					ttlCost += grpCost;
					grpCost = 0;								
					nItem=1;
				}
				nRow ++;
				CString szField, szAmount;			
				xls.SetCellMergedColumns(0,nRow,12);
				rs.GetValue(_T("typename"), tmpStr);
				xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);
				szOldLine = szNewLine;
				nItem=1;
			}
		}
		nRow ++;		
		tmpStr.Format(_T("%d"), nItem++);
		xls.SetCellText(0,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("name"), tmpStr);
		xls.SetCellText(1,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("id"), tmpStr);
		xls.SetCellText(2,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("unit"), tmpStr);		
		xls.SetCellText(3,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("countryid"), tmpStr);
		xls.SetCellText(4,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("exdate"), tmpStr);
		xls.SetCellText(5,nRow,CDate::Convert(tmpStr,yyyymmdd,ddmmyyyy), FMT_TEXT);
		rs.GetValue(_T("price"), tmpStr);				
		xls.SetCellText(6,nRow,tmpStr, FMT_NUMBER1);
		if (!m_bCheck)
		{
			rs.GetValue(_T("instockqty"), qty);
			tmpStr.Format(_T("%ld"), qty);
			xls.SetCellText(7,nRow,tmpStr, FMT_NUMBER1);				
			rs.GetValue(_T("instockamount"), cost);				
			tmpStr.Format(_T("%.2f"), cost);
			grpCost+=cost;
			xls.SetCellText(8,nRow,tmpStr, FMT_NUMBER1);
		}
		rs.GetValue(_T("ortherqty"), qty);
		tmpStr.Format(_T("%ld"), qty);
		xls.SetCellText(11,nRow,tmpStr, FMT_NUMBER1);				
		rs.GetValue(_T("lotno"), tmpStr);		
		xls.SetCellText(13,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("country"), tmpStr);
		xls.SetCellText(14,nRow,tmpStr, FMT_TEXT);

		rs.MoveNext();
}
	ttlCost += grpCost;
	if(grpCost > 0){
		nRow++;
		TranslateString(_T("\x43\x1ED9ng:"), tmpStr);		
		xls.SetCellMergedColumns(1,nRow,7);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT,true,11,0);
		tmpStr.Format(_T("%.2f"), grpCost);
		xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1,true,11,0);		
	}
	if(ttlCost > 0){
		nRow++;
		TranslateString(_T("Total Amount"), tmpStr);		
		xls.SetCellMergedColumns(1,nRow,7);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT,true,11,0);
		tmpStr.Format(_T("%.2f"), ttlCost);
		xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1,true,11,0);		
	}
	xls.Save(_T("Exports\\Bangkiemkethuocvattu.xls"));
	EndWaitCursor();
}

void CPMSReportConditionDialog::OnPrintTonghopnhapxuattonchitiettheonghiepvu(bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL, szWhere;
	CString szType, szTypeName;

	BeginWaitCursor();
	if(!rpt.Init(_T("Reports/HMS/PMS_CHITIETTONGHOPNHAPXUATTONKHOTHEONGHIEPVU.RPT"),true) )
	{
		return;
	}

	 for (int i=0 ; i<= m_wndType.GetCount(); i++)
	 {
		 if(m_wndType.GetCheck(i)){
			 m_wndType.SetCurSel(i);
		 	if(!szType.IsEmpty())
					szType += _T(",");						
					szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));			
			if(!szTypeName.IsEmpty())
					szTypeName += _T(", ");						
			szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		 }
	 }
	
	if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and pmi_typeid in(%s)"), szType);
	

	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(pmi_groupid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}

	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and pmsi_originalid=%d "), str2int(m_szSourceKey));
	}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
	CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	rpt.GetReportHeader()->SetValue(_T("StockName"), m_szStockName);
	rpt.GetReportHeader()->SetValue(_T("Type"), m_wndType.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Group"), m_wndGroup.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Original"), m_wndSource.GetCurrent(1));


	//Page Header
	//Report Detail


szSQL.Format(_T(" SELECT * FROM  ( SELECT 	") \
	_T(" 			hfg_index as idx, hfg_name as typename,pmi_typeid,pmi_id as id,  	") \
	_T(" 			pmi_name as name,  	pmi_unit as unit,  	pmsi_vatprice as price,  	") \
	_T("			sum(periodqty) as periodqty, sum(periodqty) * pmsi_vatprice as amountperiodqty,sum(iHoadonP) as iHoadonP,")\
	_T(" 			sum(itutrucZ) as itutrucZ,sum(ikhoaX) as ikhoaX, sum(iKhoA) as iKhoA,") \
	_T(" 			sum(iPhieulinhY) as iPhieulinhY,sum(iDieuchinhW) as iDieuchinhW,") \
	_T(" 			sum(iHoadonP + itutrucZ + iKhoaX + iKhoA + iPhieulinhY + iDieuchinhW) as itongnhap, ") \
	_T(" 			sum(iHoadonP + itutrucZ + iKhoaX + iKhoA + iPhieulinhY + iDieuchinhW) * pmsi_vatprice as itongtiennhap , ") \
	_T(" 			sum(eXaH) as eXaH, sum(eHangngayB) as eHangngayB, sum(eTutruC) as eTutruC,") \
	_T(" 			sum(ethanhlyD) as ethanhlyD,sum(ehethanE) as ehethanE,") \
	_T(" 			sum(eVohongF) as eVohongF,sum(edieuchinhG) as edieuchinhG,sum(eKhacO) as eKhacO,") \
	_T(" 			sum(eKhoaX) as eKhoaX,sum(eKhoA) as eKhoA,sum(ephieulinhY) as ephieulinhY,sum(eTutrucZ) as eTutrucZ,sum(eDonthuocP) as eDonthuocP,") \
	_T(" 			sum(eXaH + eHangngayB+eTutruC + ethanhlyD +ehethanE+eVohongF+edieuchinhG + eKhacO +") \
	_T(" 			eKhoaX+eKhoA +ephieulinhY+eTutrucZ+eDonthuocP) as eTongxuat,") \
	_T(" 			sum(eXaH + eHangngayB+eTutruC + ethanhlyD +ehethanE+eVohongF+edieuchinhG + eKhacO +") \
	_T(" 			eKhoaX + eKhoA +ephieulinhY+eTutrucZ+eDonthuocP)* pmsi_vatprice as eTongtienxuat,") \
	_T(" 			sum(impqty) as impqty, sum(expqty) as expqty,")\
	_T("			sum(periodqty+impqty-expqty)  as instockqty,") \
	_T("			sum(periodqty+impqty-expqty) *  pmsi_vatprice as amountinstockqty") \
	_T(" 		FROM ( SELECT * FROM ( 	SELECT impstockid as stockid, impdate as iodate, sitemid, ") \
	_T(" 						case when iotype='P' then impqty else 0 end as iHoadonP,") \
	_T(" 						case when iotype='Z' then impqty else 0 end as itutrucZ,") \
	_T(" 						case when iotype='X' then impqty else 0 end as ikhoaX,") \
	_T(" 						case when iotype='A' then impqty else 0 end as iKhoA,") \
	_T(" 						case when iotype='Y' then impqty else 0 end as iPhieulinhY,") \
	_T(" 						case when iotype='W' then impqty else 0 end as iDieuchinhW,") \
	_T(" 						0 as eXaH, 0 as eHangngayB , 0 as eTutruC, 0 as ethanhlyD, ") \
	_T(" 						0 as ehethanE, 0 as eVohongF, 0 as edieuchinhG, 0 as eKhacO,") \
	_T(" 						0 as eKhoaX,0 as eKhoA,0 as ephieulinhY,0 as eTutrucZ,0 as eDonthuocP, ") \
	_T("						0 as periodqty, impqty, 0 as expqty") \
	_T(" 					FROM pmsv_import") \
	_T(" 					UNION ALL") \
	_T(" 					SELECT impstockid, expdate, sitemid, 0,0, 0, 0,0,0,") \
	_T(" 						case when iotype='H' then expqty else 0 end as eXaH,") \
	_T(" 						case when iotype='B' then expqty else 0 end as eHangngayB,") \
	_T(" 						case when iotype in('C','M') then expqty else 0 end as eTutruC,") \
	_T(" 						case when iotype='D' then expqty else 0 end  as ethanhlyD,") \
	_T(" 						case when iotype='E' then expqty else 0 end as ehethanE,") \
	_T(" 						case when iotype='F' then expqty else 0 end as eVohongF,") \
	_T(" 						case when iotype='G' then expqty else 0 end as edieuchinhG,") \
	_T(" 						case when iotype='O' then expqty else 0 end as eKhacO,") \
	_T(" 						case when iotype='X' then expqty else 0 end as ekhoaX,") \
	_T(" 						case when iotype='A' then expqty else 0 end as eKhoA,") \
	_T(" 						case when iotype='I' then expqty else 0 end as ephieulinhY,") \
	_T(" 						case when iotype='Z' then expqty else 0 end as eTutrucZ,") \
	_T(" 						case when iotype='P' then expqty else 0 end as eDonthuocP,") \
	_T(" 						0, 0, expqty") \
	_T(" 					FROM pmsv_export ") \
	_T(" 		) as iotbl WHERE stockid=%d  AND iodate between timestamp '%s' and timestamp '%s' ") \
	_T(" 		UNION ALL") \
	_T(" 		SELECT stockid, iodate, sitemid, 0,0, 0, 0,0,0,0, 0,0,0,0,0,0,0,0,0,0,0,0,coalesce(impqty-expqty, 0), 0, 0") \
	_T(" 		FROM 	(SELECT impstockid as stockid, impdate as iodate, sitemid, ") \
	_T(" 						case when iotype='P' then impqty else 0 end as iHoadonP,") \
	_T(" 						case when iotype='Z' then impqty else 0 end as itutrucZ,") \
	_T(" 						case when iotype='X' then impqty else 0 end as ikhoaX,") \
	_T(" 						case when iotype='A' then impqty else 0 end as iKhoA,") \
	_T(" 						case when iotype='Y' then impqty else 0 end as iPhieulinhY,") \
	_T(" 						case when iotype='W' then impqty else 0 end as iDieuchinhW,") \
	_T(" 						0 as eXaH, 0 as eHangngayB , 0 as eTutruC, 0 as ethanhlyD, ") \
	_T(" 						0 as ehethanE, 0 as eVohongF, 0 as edieuchinhG, 0 as eKhacO,") \
	_T(" 						0 as eKhoaX,0 as eKhoA,0 as ephieulinhY,0 as eTutrucZ,0 as eDonthuocP , ") \
	_T("						0 as periodqty,  impqty, 0 as expqty") \
	_T(" 					FROM pmsv_import") \
	_T(" 					UNION ALL") \
	_T(" 					SELECT impstockid, expdate, sitemid, 0,0, 0, 0,0,0,") \
	_T(" 						case when iotype='H' then expqty else 0 end as eXaH,") \
	_T(" 						case when iotype='B' then expqty else 0 end as eHangngayB,") \
	_T(" 						case when iotype in('C','M') then expqty else 0 end as eTutruC,") \
	_T(" 						case when iotype='D' then expqty else 0 end as ethanhlyD,") \
	_T(" 						case when iotype='E' then expqty else 0 end as ehethanE,") \
	_T(" 						case when iotype='F' then expqty else 0 end as eVohongF,") \
	_T(" 						case when iotype='G' then expqty else 0 end as edieuchinhG,") \
	_T(" 						case when iotype='O' then expqty else 0 end as eKhacO,") \
	_T(" 						case when iotype='X' then expqty else 0 end as eKhoaX,") \
	_T(" 						case when iotype='A' then expqty else 0 end as eKhoA,") \
	_T(" 						case when iotype='I' then expqty else 0 end as ephieulinhY,") \
	_T(" 						case when iotype='Z' then expqty else 0 end as eTutrucZ,") \
	_T(" 						case when iotype='P' then expqty else 0 end as eDonthuocP,") \
	_T(" 						0, 0, expqty") \
	_T(" 					FROM pmsv_export ") \
	_T(" 		) as ptbl WHERE 	stockid=%d  	AND iodate < timestamp '%s'  ") \
	_T(" 	) as instock ") \
	_T(" 	LEFT JOIN pms_stockitems ON(pmsi_id=sitemid) ") \
	_T(" 	LEFT JOIN pms_items ON(pmsi_itemid=pmi_id) ") \
	_T("	LEFT JOIN hms_fee_group ON(hfg_id = pmi_typeid) ") \
	_T(" WHERE sitemid > 0  %s ") \
	_T(" 		GROUP BY idx, typename,pmi_typeid,id, name, unit, price") \
	_T(" ) as tbl  ") \
	_T(" WHERE length(id) > 0  and (periodqty+instockqty+eTongxuat) > 0 ") \
	_T(" ORDER BY idx,typename,name, unit"), ToInt(m_szStockKey), m_szFromDate, m_szToDate, ToInt(m_szStockKey), m_szFromDate, szWhere);


	CReportSection* rptDetail;
	CString szOldLine, szNewLine;
	CRecord rs(&pMF->m_db);
	long double grpCost[30];
	long double ttlCost[30];
	double cost = 0;
	int nItem = 1;
	//_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);

	if(rs.IsEOF()){
		ShowMessage(150, MB_ICONSTOP);
		return ;
	}

	long nTotalPeriod=0;
	long nTotalImport=0;
	long nTotalExport=0;
	long nTotalInStock=0;
			
	for (int i =0; i < 31; i++)
	{
		grpCost[i] = 0;
		ttlCost[i] = 0;
	}
	
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("pmi_typeid"), szNewLine);
		if(szNewLine != szOldLine && !szNewLine.IsEmpty())
		{
			CString szField;
			if((grpCost[7] + grpCost[15] + grpCost[28] + grpCost[30]) > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				for (int i =7; i < 31; i++)
				{
					szField.Format(_T("s%d"), i);						
					FormatCurrency(grpCost[i], tmpStr);
					rptDetail->SetValue(szField, tmpStr);
					ttlCost[i] += grpCost[i];						
					grpCost[i] = 0;				
				}
			}

			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			rs.GetValue(_T("catname"), tmpStr);
			rptDetail->SetValue(_T("GroupName"),tmpStr );
			szOldLine = szNewLine;
			nItem=1;
		}

		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nItem++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("id"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("name"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("price"), tmpStr);				
		rptDetail->SetValue(_T("5"), tmpStr);
		// Ton dau ky
		rs.GetValue(_T("periodqty"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rs.GetValue(_T("amountperiodqty"), cost);	
		grpCost[7] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);				

		//Nhap trong ky
		rs.GetValue(_T("iHoadonP"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.GetValue(_T("itutrucZ"), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);								
		rs.GetValue(_T("iKhoaX"), tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);
		rs.GetValue(_T("iKhoA"), tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);
		rs.GetValue(_T("iPhieulinhY"), tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);				
		rs.GetValue(_T("iDieuchinhW"), tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);
		rs.GetValue(_T("itongnhap"), tmpStr);
		rptDetail->SetValue(_T("14"), tmpStr);
		rs.GetValue(_T("itongtiennhap"), cost);
		grpCost[15] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("15"), tmpStr);
		// Xuat trong ky
		rs.GetValue(_T("eXaH"), tmpStr);
		rptDetail->SetValue(_T("16"), tmpStr);
		rs.GetValue(_T("eHangngayB"), tmpStr);
		rptDetail->SetValue(_T("17"), tmpStr);
		rs.GetValue(_T("eTutruC"), tmpStr);
		rptDetail->SetValue(_T("18"), tmpStr);
		rs.GetValue(_T("ethanhlyD"), tmpStr);
		rptDetail->SetValue(_T("19"), tmpStr);
		rs.GetValue(_T("ehethanE"), tmpStr);
		rptDetail->SetValue(_T("20"), tmpStr);
		rs.GetValue(_T("eVohongF"), tmpStr);
		rptDetail->SetValue(_T("21"), tmpStr);
		rs.GetValue(_T("edieuchinhG"), tmpStr);
		rptDetail->SetValue(_T("22"), tmpStr);
		rs.GetValue(_T("eKhacO"), tmpStr);
		rptDetail->SetValue(_T("23"), tmpStr);
		rs.GetValue(_T("eKhoA"), tmpStr);
		rptDetail->SetValue(_T("24"), tmpStr);
		rs.GetValue(_T("ephieulinhY"), tmpStr);
		rptDetail->SetValue(_T("25"), tmpStr);				
		rs.GetValue(_T("eDonthuocP"), tmpStr);
		rptDetail->SetValue(_T("26"), tmpStr);
		rs.GetValue(_T("etongxuat"), tmpStr);
		rptDetail->SetValue(_T("27"), tmpStr);
		rs.GetValue(_T("etongtienxuat"), cost);
		grpCost[28] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("28"), tmpStr);
		//Ton cuoi ky
			
		rs.GetValue(_T("instockqty"), tmpStr);				
		rptDetail->SetValue(_T("29"), tmpStr);
		rs.GetValue(_T("amountinstockqty"), cost);
		grpCost[30] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("30"), tmpStr);
		rs.MoveNext();

	}
			
	for (int i =0; i < 31; i++)
	{
		ttlCost[i] += grpCost[i];
	}
	if((grpCost[7] + grpCost[15] + grpCost[28] + grpCost[30]) > 0){
		CString szField;
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
		for (int i =7; i < 31; i++)
		{
			szField.Format(_T("s%d"), i);
			FormatCurrency(grpCost[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);
		}			
	}
	if((ttlCost[7] + ttlCost[15] + ttlCost[28] + ttlCost[30]) > 0){
		CString szField, szAmount;
		TranslateString(_T("Total Amount"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		for (int i =7; i < 31; i++)
		{
			
			szField.Format(_T("s%d"), i);
			FormatCurrency(ttlCost[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);					
		}
	}
			
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	//	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print();
	EndWaitCursor();
}
/*--------------------------------------------------------------------------------------------------------------------------------
Export To Excell Tong hop nhap xuat ton kho chi tiet theo nghiep vu
---------------------------------------------------------------------------------------------------------------------------------*/
void CPMSReportConditionDialog::OnExcell_Baocaotonghopnhapxuattonchitiettheonghiepvu(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL, szWhere;
	CString szType, szTypeName;
	BeginWaitCursor();
	 for (int i=0 ; i<= m_wndType.GetCount(); i++)
	 {
		 if(m_wndType.GetCheck(i)){
			 m_wndType.SetCurSel(i);
		 	if(!szType.IsEmpty())
					szType += _T(",");						
					szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));			
			if(!szTypeName.IsEmpty())
					szTypeName += _T(", ");						
			szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		 }
	 }
	
	if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and pmi_typeid in(%s)"), szType);

	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(pmi_groupid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}

	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and pmsi_originalid=%d "), str2int(m_szSourceKey));
	}	
			szSQL.Format(_T(" SELECT * FROM  ( SELECT 	") \
			_T(" 			pmi_typeid, hfg_name as typename, hfg_index as idx,pmi_id as id,  	") \
			_T(" 			pmi_name as name,pmg_name as namegeneric,  	pmi_unit as unit, 	pmsi_vatprice as price,  	") \
			_T(" 			(select distinct distinct sc_name from sys_country where sc_id=pmi_countryid) as orgcountry, ") \
			_T("			sum(periodqty) as periodqty, sum(periodqty) * pmsi_vatprice as amountperiodqty,sum(iHoadonP) as iHoadonP,")\
			_T(" 			sum(itutrucZ) as itutrucZ,sum(ikhoaX) as ikhoaX, sum(iKhoA) as iKhoA,") \
			_T(" 			sum(iPhieulinhY) as iPhieulinhY,sum(iDieuchinhW) as iDieuchinhW,") \
			_T(" 			sum(iHoadonP + itutrucZ + iKhoaX + iKhoA + iPhieulinhY + iDieuchinhW) as itongnhap, ") \
			_T(" 			sum(iHoadonP + itutrucZ + iKhoaX + iKhoA + iPhieulinhY + iDieuchinhW) * pmsi_vatprice as itongtiennhap , ") \
			_T(" 			sum(eXaH) as eXaH, sum(eHangngayB) as eHangngayB, sum(eTutruC) as eTutruC,") \
			_T(" 			sum(ethanhlyD) as ethanhlyD,sum(ehethanE) as ehethanE,") \
			_T(" 			sum(eVohongF) as eVohongF,sum(edieuchinhG) as edieuchinhG,sum(eKhacO) as eKhacO,") \
			_T(" 			sum(eKhoaX) as eKhoaX,sum(eKhoA) as eKhoA,sum(ephieulinhY) as ephieulinhY,sum(eTutrucZ) as eTutrucZ,sum(eDonthuocP) as eDonthuocP,") \
			_T(" 			sum(eXaH + eHangngayB+eTutruC + ethanhlyD +ehethanE+eVohongF+edieuchinhG + eKhacO +") \
			_T(" 			eKhoaX+eKhoA +ephieulinhY+eTutrucZ+eDonthuocP) as eTongxuat,") \
			_T(" 			sum(eXaH + eHangngayB+eTutruC + ethanhlyD +ehethanE+eVohongF+edieuchinhG + eKhacO +") \
			_T(" 			eKhoaX + eKhoA +ephieulinhY+eTutrucZ+eDonthuocP)* pmsi_vatprice as eTongtienxuat,") \
			_T(" 			sum(impqty) as impqty, sum(expqty) as expqty,")\
			_T("			sum(periodqty+impqty-expqty)  as instockqty,") \
			_T("			sum(periodqty+impqty-expqty) *  pmsi_vatprice as amountinstockqty") \
			_T(" 		FROM ( SELECT * FROM ( 	SELECT impstockid as stockid, impdate as iodate, sitemid, ") \
			_T(" 						case when iotype='P' then impqty else 0 end as iHoadonP,") \
			_T(" 						case when iotype='Z' then impqty else 0 end as itutrucZ,") \
			_T(" 						case when iotype='X' then impqty else 0 end as ikhoaX,") \
			_T(" 						case when iotype='A' then impqty else 0 end as iKhoA,") \
			_T(" 						case when iotype='Y' then impqty else 0 end as iPhieulinhY,") \
			_T(" 						case when iotype='W' then impqty else 0 end as iDieuchinhW,") \
			_T(" 						0 as eXaH, 0 as eHangngayB , 0 as eTutruC, 0 as ethanhlyD, ") \
			_T(" 						0 as ehethanE, 0 as eVohongF, 0 as edieuchinhG, 0 as eKhacO,") \
			_T(" 						0 as eKhoaX,0 as eKhoA,0 as ephieulinhY,0 as eTutrucZ,0 as eDonthuocP, ") \
			_T("						0 as periodqty, impqty, 0 as expqty") \
			_T(" 					FROM pmsv_import") \
			_T(" 					UNION ALL") \
			_T(" 					SELECT impstockid, expdate, sitemid, 0,0, 0, 0,0,0,") \
			_T(" 						case when iotype='H' then expqty else 0 end as eXaH,") \
			_T(" 						case when iotype='B' then expqty else 0 end as eHangngayB,") \
			_T(" 						case when iotype in('C','M') then expqty else 0 end as eTutruC,") \
			_T(" 						case when iotype='D' then expqty else 0 end  as ethanhlyD,") \
			_T(" 						case when iotype='E' then expqty else 0 end as ehethanE,") \
			_T(" 						case when iotype='F' then expqty else 0 end as eVohongF,") \
			_T(" 						case when iotype='G' then expqty else 0 end as edieuchinhG,") \
			_T(" 						case when iotype='O' then expqty else 0 end as eKhacO,") \
			_T(" 						case when iotype='X' then expqty else 0 end as ekhoaX,") \
			_T(" 						case when iotype='A' then expqty else 0 end as eKhoA,") \
			_T(" 						case when iotype='I' then expqty else 0 end as ephieulinhY,") \
			_T(" 						case when iotype='Z' then expqty else 0 end as eTutrucZ,") \
			_T(" 						case when iotype='P' then expqty else 0 end as eDonthuocP,") \
			_T(" 						0, 0, expqty") \
			_T(" 					FROM pmsv_export ") \
			_T(" 		) as iotbl WHERE stockid=%d  AND iodate between timestamp '%s' and timestamp '%s' ") \
			_T(" 		UNION ALL") \
			_T(" 		SELECT stockid, iodate, sitemid, 0,0, 0, 0,0,0,0, 0,0,0,0,0,0,0,0,0,0,0,0,coalesce(impqty-expqty, 0), 0, 0") \
			_T(" 		FROM 	(SELECT impstockid as stockid, impdate as iodate, sitemid, ") \
			_T(" 						case when iotype='P' then impqty else 0 end as iHoadonP,") \
			_T(" 						case when iotype='Z' then impqty else 0 end as itutrucZ,") \
			_T(" 						case when iotype='X' then impqty else 0 end as ikhoaX,") \
			_T(" 						case when iotype='A' then impqty else 0 end as iKhoA,") \
			_T(" 						case when iotype='Y' then impqty else 0 end as iPhieulinhY,") \
			_T(" 						case when iotype='W' then impqty else 0 end as iDieuchinhW,") \
			_T(" 						0 as eXaH, 0 as eHangngayB , 0 as eTutruC, 0 as ethanhlyD, ") \
			_T(" 						0 as ehethanE, 0 as eVohongF, 0 as edieuchinhG, 0 as eKhacO,") \
			_T(" 						0 as eKhoaX,0 as eKhoA,0 as ephieulinhY,0 as eTutrucZ,0 as eDonthuocP , ") \
			_T("						0 as periodqty,  impqty, 0 as expqty") \
			_T(" 					FROM pmsv_import") \
			_T(" 					UNION ALL") \
			_T(" 					SELECT impstockid, expdate, sitemid, 0,0, 0, 0,0,0,") \
			_T(" 						case when iotype='H' then expqty else 0 end as eXaH,") \
			_T(" 						case when iotype='B' then expqty else 0 end as eHangngayB,") \
			_T(" 						case when iotype in('C','M') then expqty else 0 end as eTutruC,") \
			_T(" 						case when iotype='D' then expqty else 0 end as ethanhlyD,") \
			_T(" 						case when iotype='E' then expqty else 0 end as ehethanE,") \
			_T(" 						case when iotype='F' then expqty else 0 end as eVohongF,") \
			_T(" 						case when iotype='G' then expqty else 0 end as edieuchinhG,") \
			_T(" 						case when iotype='O' then expqty else 0 end as eKhacO,") \
			_T(" 						case when iotype='X' then expqty else 0 end as eKhoaX,") \
			_T(" 						case when iotype='A' then expqty else 0 end as eKhoA,") \
			_T(" 						case when iotype='I' then expqty else 0 end as ephieulinhY,") \
			_T(" 						case when iotype='Z' then expqty else 0 end as eTutrucZ,") \
			_T(" 						case when iotype='P' then expqty else 0 end as eDonthuocP,") \
			_T(" 						0, 0, expqty") \
			_T(" 					FROM pmsv_export ") \
			_T(" 		) as ptbl WHERE 	stockid=%d  	AND iodate < timestamp '%s'  ") \
			_T(" 	) as instock ") \
			_T(" 	LEFT JOIN pms_stockitems ON(pmsi_id=sitemid) ") \
			_T(" 	LEFT JOIN pms_items ON(pmsi_itemid=pmi_id) ") \
			_T("	LEFT JOIN pms_generic ON(pmg_id=pmi_genericid) ") \
			_T("	LEFT JOIN hms_fee_group ON(hfg_id = pmi_typeid) ") \
			_T(" WHERE sitemid > 0  %s ") \
			_T(" 		GROUP BY typename, idx, pmi_typeid,id, name,namegeneric,orgcountry, unit, price") \
			_T(" ) as tbl  ") \
			_T(" WHERE length(id) > 0  and (periodqty+instockqty+eTongxuat) > 0 ") \
			_T(" ORDER BY idx, pmi_typeid, name,namegeneric, unit"), ToInt(m_szStockKey), m_szFromDate, m_szToDate, ToInt(m_szStockKey), m_szFromDate, szWhere);

	CReportSection* rptDetail = NULL;
	CString szOldLine, szNewLine,szAmount;
	CRecord rs(&pMF->m_db);
	long double grpCost[32];
	long double ttlCost[32];
	double cost = 0;
	int nItem = 1;
	//_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);
	if(rs.IsEOF()){
		
	}

	for (int i =0; i < 32; i++)
	{
		grpCost[i] = 0;
		ttlCost[i] = 0;
	}

		CExcel xls;	
		xls.CreateSheet(1);
		xls.SetWorksheet(0);
		xls.SetColumnWidth(1,25);
		xls.SetColumnWidth(2,25);
		xls.SetRowHeight(8,30);
		xls.SetRowHeight(9,30);

		xls.SetCellMergedColumns(0,3,12);	
		xls.SetCellText(0, 3, _T("\x43HI TI\x1EBET NH\x1EACP \x58U\x1EA4T T\x1ED2N KHO\x41 TH\x45O NGHI\x1EC6P V\x1EE4"),FMT_TEXT,true,16,0);
		xls.SetCellMergedColumns(0,4,12);
		tmpStr.Format(_T("T\x1EEB ngy\x61\x66 %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
		xls.SetCellText(0, 4, tmpStr,FMT_TEXT,true,12,0);
		int nRow = 5;
		tmpStr.Format(_T("T\x1EEB kho: %s"), m_wndStock.GetCurrent(1));
		xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);

		if(!m_szTypeKey.IsEmpty())
		{
			tmpStr.Format(_T("Lo\x1EA1i: %s"), m_wndType.GetCurrent(1));
			xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
		}
		if(!m_szGroupKey.IsEmpty())
		{
			tmpStr.Format(_T("Nh\xF3m: %s"), m_wndGroup.GetCurrent(1));
			xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
		}
		if(!m_szSourceKey.IsEmpty())
		{
			tmpStr.Format(_T("Ngu\x1ED3n thu\x1ED1\x63: %s"), m_wndSource.GetCurrent(1));
			xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
		}

		nRow+=3;

		xls.SetCellText(0, nRow, _T("STT"),FMT_TEXT,TRUE);		
		xls.SetCellText(1, nRow, _T("Generic"),FMT_TEXT,TRUE);
		xls.SetCellText(2, nRow, _T("Tên thu\x1ED1\x63/ Hàm l\x1B0\x1EE3ng"),FMT_TEXT,TRUE);
		xls.SetCellText(3, nRow, _T("\x110\x1A1n v\x1ECB"),FMT_TEXT,TRUE);
		xls.SetCellText(4, nRow, _T("\x110\x1A1n giá"),FMT_TEXT,TRUE);
		xls.SetCellText(5, nRow, _T("N\x1B0\x1EDB\x63 S\x58"),FMT_TEXT,TRUE);

		//Ton dau ky
		xls.SetCellMergedColumns(6,nRow-1,2);
		xls.SetCellText(6, nRow-1, _T("T\x1ED3n \x111\x1EA7u k\x1EF3"),FMT_TEXT,TRUE);		
		xls.SetCellText(6, nRow, _T("S\x1ED1 l\x1B0\x1EE3ng"),FMT_TEXT,TRUE);
		xls.SetCellText(7, nRow, _T("S\x1ED1 ti\x1EC1n"),FMT_TEXT,TRUE);
		//Nhap trong ky
		xls.SetCellMergedColumns(8,nRow-1,8);
		xls.SetCellText(8, nRow-1, _T("Nh\x1EADp trong k\x1EF3"),FMT_TEXT,TRUE);		
		xls.SetCellText(8, nRow, _T("Nhà \x63ung \x63\x1EA5p"),FMT_TEXT,TRUE);
		xls.SetCellText(9, nRow, _T("HT T\x1EE7 tr\x1EF1\x63"),FMT_TEXT,TRUE);
		xls.SetCellText(10, nRow, _T("HT T\x1EEB kho\x61"),FMT_TEXT,TRUE);
		xls.SetCellText(11, nRow, _T("HT T\x1EEB kho"),FMT_TEXT,TRUE);
		xls.SetCellText(12, nRow, _T("HT Phi\x1EBFu l\x129nh"),FMT_TEXT,TRUE);
		xls.SetCellText(13, nRow, _T("\x110i\x1EC1u \x63h\x1EC9nh"),FMT_TEXT,TRUE);
		xls.SetCellText(14, nRow, _T("T\x1ED5ng"),FMT_TEXT,TRUE);
		xls.SetCellText(15, nRow, _T("S\x1ED1 ti\x1EC1n"),FMT_TEXT,TRUE);
		//Xuat trong ky
		xls.SetCellMergedColumns(16,nRow-1,14);
		xls.SetCellText(16, nRow-1, _T("Xu\x1EA5t trong k\x1EF3"),FMT_TEXT,TRUE);
		xls.SetCellText(16, nRow, _T("Xã ph\x1B0\x1EDDng"),FMT_TEXT,TRUE);
		xls.SetCellText(17, nRow, _T("\x44\x1EF1 trù"),FMT_TEXT,TRUE);
		xls.SetCellText(18, nRow, _T("T\x1EE7 tr\x1EF1\x63"),FMT_TEXT,TRUE);
		xls.SetCellText(19, nRow, _T("Th\x61nh lý"),FMT_TEXT,TRUE);
		xls.SetCellText(20, nRow, _T("H\x1EBFt h\x1EA1n"),FMT_TEXT,TRUE);
		xls.SetCellText(21, nRow, _T("H\x1ECFng v\x1EE1"),FMT_TEXT,TRUE);
		xls.SetCellText(22, nRow, _T("\x110i\x1EC1u \x63h\x1EC9nh"),FMT_TEXT,TRUE);	
		xls.SetCellText(23, nRow, _T("Xu\x1EA5t khá\x63"),FMT_TEXT,TRUE);
		xls.SetCellText(24, nRow, _T("Ho\xE0n tr\x1EA3 t\x1EEB kho"),FMT_TEXT,TRUE);
		xls.SetCellText(25, nRow, _T("L\x1B0u \x63huy\x1EC3n"),FMT_TEXT,TRUE);
		xls.SetCellText(26, nRow, _T("Phi\x1EBFu l\x129nh"),FMT_TEXT,TRUE);
		xls.SetCellText(27, nRow, _T("\x110\x1A1n thu\x1ED1\x63"),FMT_TEXT,TRUE);			
		xls.SetCellText(28, nRow, _T("T\x1ED5ng"),FMT_TEXT,TRUE);
		xls.SetCellText(29, nRow, _T("S\x1ED1 ti\x1EC1n"),FMT_TEXT,TRUE);	
		//Ton cuoi ky
		xls.SetCellMergedColumns(30,nRow-1,2);
		xls.SetCellText(30, nRow-1, _T("T\x1ED3n \x63u\x1ED1i k\x1EF3"),FMT_TEXT,TRUE);
		xls.SetCellText(30, nRow, _T("S\x1ED1 l\x1B0\x1EE3ng"),FMT_TEXT,TRUE);		
		xls.SetCellText(31, nRow, _T("S\x1ED1 ti\x1EC1n"),FMT_TEXT,TRUE);

			
		while(!rs.IsEOF())
		{
			rs.GetValue(_T("pmi_typeid"), szNewLine);
			if(szNewLine != szOldLine && !szNewLine.IsEmpty()){
			CString szField;
			if((grpCost[7] + grpCost[15] + grpCost[29] + grpCost[31]) > 0)
			{					
				nRow ++;
				TranslateString(_T("Total Group"), szAmount);												
				xls.SetCellMergedColumns(0,nRow,6);
				xls.SetCellText(0, nRow, szAmount, FMT_TEXT,true,11,0);
				tmpStr.Format(_T("%.2f"), grpCost[7]);
				xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1,true,11,0);
				tmpStr.Format(_T("%.2f"), grpCost[15]);
				xls.SetCellText(15, nRow, tmpStr, FMT_NUMBER1,true,11,0);
				tmpStr.Format(_T("%.2f"), grpCost[29]);
				xls.SetCellText(29, nRow, tmpStr, FMT_NUMBER1,true,11,0);
				tmpStr.Format(_T("%.2f"), grpCost[31]);
				xls.SetCellText(31, nRow, tmpStr, FMT_NUMBER1,true,11,0);
				for(int i=0; i<32;i++)
				{
					ttlCost[i] += grpCost[i];						
					grpCost[i]= 0;
				}
			}
			
			nRow ++;			
			rs.GetValue(_T("catname"), tmpStr);
			xls.SetCellMergedColumns(0,nRow,6);
			xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);
			szOldLine = szNewLine;
			nItem=1;
			}

		nRow ++;			
		tmpStr.Format(_T("%d"), nItem++);
		xls.SetCellText(0, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("namegeneric"), tmpStr);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("name"), tmpStr);
		xls.SetCellText(2, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("unit"), tmpStr);
		xls.SetCellText(3, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("price"), tmpStr);				
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);	
		rs.GetValue(_T("orgcountry"), tmpStr);
		xls.SetCellText(5, nRow, tmpStr, FMT_TEXT);

		//Ton dau ky
		rs.GetValue(_T("periodqty"), tmpStr);			
		xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER1);			
		rs.GetValue(_T("amountperiodqty"), tmpStr);
		grpCost[7]+=ToFloat(tmpStr);
		xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1);
		//Nhap trong ky
		rs.GetValue(_T("iHoadonP"), tmpStr);
		xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("itutrucZ"), tmpStr);
		xls.SetCellText(9, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("iKhoaX"), tmpStr);
		xls.SetCellText(10, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("iKhoA"), tmpStr);
		xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("iPhieulinhY"), tmpStr);
		xls.SetCellText(12, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("iDieuchinhW"), tmpStr);
		xls.SetCellText(13, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("iTongnhap"), tmpStr);
		xls.SetCellText(14, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("itongtiennhap"), tmpStr);
		grpCost[15]+=ToFloat(tmpStr);
		xls.SetCellText(15, nRow, tmpStr, FMT_NUMBER1);			

		// Xuat trong ky
		rs.GetValue(_T("eXaH"), tmpStr);
		xls.SetCellText(16, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("eHangngayB"), tmpStr);
		xls.SetCellText(17, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("eTutruC"), tmpStr);
		xls.SetCellText(18, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("ethanhlyD"), tmpStr);
		xls.SetCellText(19, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("ehethanE"), tmpStr);
		xls.SetCellText(20, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("eVohongF"), tmpStr);
		xls.SetCellText(21, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("edieuchinhG"), tmpStr);
		xls.SetCellText(22, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("eKhacO"), tmpStr);
		xls.SetCellText(23, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("eKhoaX"), tmpStr);
		xls.SetCellText(24, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("eKhoA"), tmpStr);
		xls.SetCellText(25, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("ephieulinhY"), tmpStr);			
		xls.SetCellText(26, nRow, tmpStr, FMT_NUMBER1);			
		rs.GetValue(_T("eDonthuocP"), tmpStr);
		xls.SetCellText(27, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("etongxuat"), tmpStr);						
		xls.SetCellText(28, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("etongtienxuat"), tmpStr);
		grpCost[29]+=ToFloat(tmpStr);			
		xls.SetCellText(29, nRow, tmpStr, FMT_NUMBER1);	
		// Tong xuat
		rs.GetValue(_T("instockqty"), tmpStr);				
		xls.SetCellText(30, nRow, tmpStr, FMT_NUMBER1);			
		rs.GetValue(_T("amountinstockqty"), tmpStr);
		grpCost[31]+=ToFloat(tmpStr);
		xls.SetCellText(31, nRow, tmpStr, FMT_NUMBER1);
		rs.MoveNext();
	}
		
		for (int i =0; i < 32; i++)
		{
				ttlCost[i] += grpCost[i];
		}

		if((grpCost[7] + grpCost[15] + grpCost[29] + grpCost[31]) > 0){
			nRow ++;
			TranslateString(_T("Total Group"), szAmount);												
			xls.SetCellMergedColumns(0,nRow,6);
			xls.SetCellText(0, nRow, szAmount, FMT_TEXT,true,11,0);
			tmpStr.Format(_T("%.2f"), grpCost[7]);
			xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1,true,11,0);
			tmpStr.Format(_T("%.2f"), grpCost[15]);
			xls.SetCellText(15, nRow, tmpStr, FMT_NUMBER1,true,11,0);
			tmpStr.Format(_T("%.2f"), grpCost[29]);
			xls.SetCellText(29, nRow, tmpStr, FMT_NUMBER1,true,11,0);
			tmpStr.Format(_T("%.2f"), grpCost[31]);
			xls.SetCellText(31, nRow, tmpStr, FMT_NUMBER1,true,11,0);			
		}
		if((ttlCost[7] + ttlCost[15] + ttlCost[29] + ttlCost[31]) > 0)
		{
			nRow ++;
			TranslateString(_T("Total Amount"), szAmount);												
			xls.SetCellMergedColumns(0,nRow,6);
			xls.SetCellText(0, nRow, szAmount, FMT_TEXT,true,11,0);
			tmpStr.Format(_T("%.2f"), ttlCost[7]);
			xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1,true,11,0);
			tmpStr.Format(_T("%.2f"), ttlCost[15]);
			xls.SetCellText(15, nRow, tmpStr, FMT_NUMBER1,true,11,0);
			tmpStr.Format(_T("%.2f"), ttlCost[29]);
			xls.SetCellText(29, nRow, tmpStr, FMT_NUMBER1,true,11,0);
			tmpStr.Format(_T("%.2f"), ttlCost[31]);
			xls.SetCellText(31, nRow, tmpStr, FMT_NUMBER1,true,11,0);		
		}
		
		xls.Save(_T("Exports\\CHITIETNHAPXUATTONKHOTHEONGHIEPVU.XLS"));
		EndWaitCursor();
}
/*--------------------------------------------------------------------------------------------------------------------------------
Print Bien danh sach thuoc het han 
---------------------------------------------------------------------------------------------------------------------------------*/
void CPMSReportConditionDialog::OnPrint_Danhsachthuochethan(bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CReportSection* rptDetail;
	CString szOldLine, szNewLine,szMemory;
	CRecord rs(&pMF->m_db);	
	long double grpCost;
	long double ttlCost;
	double cost = 0;
	int nItem = 1,nTotalItem=0;
	BeginWaitCursor();
	CString tmpStr,tmp, szDate, szSysDate, szSQL,szName,szAmount, szWhere;

	if(!rpt.Init(_T("Reports/HMS/PMS_DANHSACHTHUOCHETHAN.RPT"),true) ){
			return;
	}
				
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
	CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	rpt.GetReportHeader()->SetValue(_T("StockName"), m_szStockName);
	rpt.GetReportHeader()->SetValue(_T("Type"), m_wndType.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Group"), m_wndGroup.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Original"), m_wndSource.GetCurrent(1));
			//Page Header
			//Report Detail
		
	if(!m_szTypeKey.IsEmpty())
	{
		szWhere.Format(_T(" and substr(pmi_typeid, 1, length(rtrim('%s','0'))) =rtrim('%s','0')   "), m_szTypeKey, m_szTypeKey);
	}
	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(pmi_groupid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}

	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and pmsi_originalid=%d "), str2int(m_szSourceKey));
	}


	szSQL.Format(_T(" SELECT (SELECT distinct hfg_name as name FROM hms_fee_group where hfg_id= pmi_typeid) as pmi_typeid,pmi_id as id,  ") \
				_T(" 	pmi_name as name,  pmi_unit as unit, (SELECT distinct pmm_name from pms_mfg where pmm_id=pmsi_mfgid) as contruyid,") \
				_T(" 	pmsi_expdate as exdate, pmsi_vatprice as price, sum(pmsl_qty) as instockqty, sum(pmsl_qty) * pmsi_vatprice as instockamount ") \
				_T(" FROM pms_stockline") \
				_T(" LEFT JOIN pms_stockitems ON(pmsl_sitemid=pmsi_id)") \
				_T(" LEFT JOIN pms_items ON(pmsi_itemid=pmi_id) ") \
				_T(" WHERE pmsl_stockid=%d and pmsl_qty >0 and date(pmsi_expdate) < date('%s') %s AND pmsl_moduleid = '%s'") \
				_T(" GROUP BY pmi_typeid, id,name, unit,contruyid,exdate,price") \
				_T(" ORDER BY pmi_typeid,name, unit") , ToInt(m_szStockKey), m_szToDate, szWhere, pMF->GetModuleID());

			//_fmsg(_T("%s"), szSQL);
		
			rs.ExecSQL(szSQL);
			if(rs.IsEOF()){
				ShowMessage(150, MB_ICONSTOP);
				return ;
			}
		
			grpCost = 0;
			ttlCost = 0;			
			
			while(!rs.IsEOF())
			{
				rs.GetValue(_T("pmi_typeid"), szNewLine);
				if(szNewLine != szOldLine && !szNewLine.IsEmpty()){	
					if(grpCost > 0)
					{
						TranslateString(_T("Total Group"), szAmount);
						rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
						rptDetail->SetValue(_T("TotalGroup"), szAmount);						
						FormatCurrency(grpCost, tmpStr);
						rptDetail->SetValue(_T("s8"), tmpStr);
						ttlCost += grpCost;						
						grpCost = 0;
						nItem=1;
					}
					rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
					rs.GetValue(_T("pmi_typeid"), tmpStr);
					rptDetail->SetValue(_T("GroupName"),tmpStr );
					szOldLine = szNewLine;
					nItem=1;
				}

				rptDetail = rpt.AddDetail();
				tmpStr.Format(_T("%d"), nItem++);
				rptDetail->SetValue(_T("1"), tmpStr);
				rs.GetValue(_T("id"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rs.GetValue(_T("name"), tmpStr);
				rptDetail->SetValue(_T("3"), tmpStr);
				rs.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				rs.GetValue(_T("contruyid"), tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				rs.GetValue(_T("price"), tmpStr);				
				rptDetail->SetValue(_T("6"), tmpStr);
				rs.GetValue(_T("instockqty"), tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);												
				rs.GetValue(_T("instockamount"), cost);		
				grpCost += cost;
				FormatCurrency(cost, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);				
				rs.GetValue(_T("exdate"), tmpStr);
				rptDetail->SetValue(_T("9"), CDate::Convert(tmpStr,yyyymmdd,ddmmyyyy));				
				nTotalItem++;
				rs.MoveNext();
			}

			ttlCost += grpCost;
			
			if(grpCost > 0){	
				TranslateString(_T("Total Group"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));	
				rptDetail->SetValue(_T("TotalGroup"), szAmount);
				FormatCurrency(grpCost, tmpStr);
				rptDetail->SetValue(_T("s8"), tmpStr);			
			}
			if(ttlCost > 0){				
				TranslateString(_T("Total Amount"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				rptDetail->SetValue(_T("TotalGroup"), szAmount);
				FormatCurrency(ttlCost, tmpStr);
				rptDetail->SetValue(_T("s8"), tmpStr);				
			}
			szSysDate = pMF->GetSysDate(); 

			tmpStr.Format(_T("%d"), nTotalItem);
			rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
			szMemory.Format(_T("%.2f"), ttlCost );
			MoneyToString(szMemory,tmpStr);
			rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
			szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
			rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);			
			if(bPreview)
				rpt.PrintPreview();
			else
				rpt.Print();
	EndWaitCursor();
}
/*--------------------------------------------------------------------------------------------------------------------------------
Export To Excell Danh sach thuoc het han
---------------------------------------------------------------------------------------------------------------------------------*/
void CPMSReportConditionDialog::OnExport_Danhsachthuochethan(bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();	
	
	CString szOldLine, szNewLine;
	CRecord rs(&pMF->m_db);
	
	long double grpCost;
	long double ttlCost;
	long qty=0;
	double cost = 0;
	int nItem = 1;
	CString tmpStr, szDate, szSysDate, szSQL,szName,szAmount,szFromDate,szToDate, szWhere;	
		
	BeginWaitCursor();
	if(!m_szTypeKey.IsEmpty())
	{
		szWhere.Format(_T(" and substr(pmi_typeid, 1, length(rtrim('%s','0'))) =rtrim('%s','0')   "), m_szTypeKey, m_szTypeKey);
	}
	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(pmi_groupid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}
	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and pmsi_originalid=%d "), str2int(m_szSourceKey));
	}

	szSQL.Format(_T(" SELECT (SELECT distinct hfg_index as idx,hfg_name as name FROM hms_fee_group where hfg_id= pmi_typeid) as pmi_typeid,pmi_id as id,  ") \
				_T(" 	pmi_name as name,  pmi_unit as unit, (SELECT distinct pmm_name from pms_mfg where pmm_id=pmsi_mfgid) as contruyid,") \
				_T(" 	pmsi_expdate as exdate, pmsi_vatprice as price, sum(pmsl_qty) as instockqty, sum(pmsl_qty) * pmsi_vatprice as instockamount ") \
				_T(" FROM pms_stockline") \
				_T(" LEFT JOIN pms_stockitems ON(pmsl_sitemid=pmsi_id)") \
				_T(" LEFT JOIN pms_items ON(pmsi_itemid=pmi_id) ") \
				_T(" WHERE pmsl_stockid=%d and pmsl_qty >0 and pmsi_expdate < date('%s') %s AND pmsl_moduleid = '%s'") \
				_T(" GROUP BY pmi_typeid, id,name, unit,contruyid,exdate,price") \
				_T(" ORDER BY idx, pmi_typeid,name, unit") , ToInt(m_szStockKey), m_szToDate, szWhere, pMF->GetModuleID());

			//_fmsg(_T("%s"), szSQL);
		
			rs.ExecSQL(szSQL);
			if(rs.IsEOF()){
				ShowMessage(150, MB_ICONSTOP);
				return ;
			}
			

	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetCellMergedColumns(0,3,12);	
	xls.SetCellText(0, 3, _T("BAO DANH SACH THUOC HET HAN"),FMT_TEXT,true,18,0);
	xls.SetCellMergedColumns(0,4,12);
	tmpStr.Format(_T("Tu ngay %s Den ngay %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(0, 5, tmpStr,FMT_TEXT,true,12,0);


		int nRow = 6;
		tmpStr.Format(_T("Tu kho: %s"), m_wndStock.GetCurrent(1));
		xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);

		if(!m_szTypeKey.IsEmpty())
		{
			tmpStr.Format(_T("Loai: %s"), m_wndType.GetCurrent(1));
			xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
		}
		if(!m_szGroupKey.IsEmpty())
		{
			tmpStr.Format(_T("Nhom: %s"), m_wndGroup.GetCurrent(1));
			xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
		}
		if(!m_szSourceKey.IsEmpty())
		{
			tmpStr.Format(_T("Nguon mua: %s"), m_wndSource.GetCurrent(1));
			xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
		}

	xls.SetCellText(0, nRow, _T("STT"));
	xls.SetCellText(1, nRow, _T("Ma thuoc"));
	xls.SetCellText(2, nRow, _T("Ten thuoc / Ham luong"));
	xls.SetCellText(3, nRow, _T("Don vi"));
	xls.SetCellText(4, nRow, _T("Nuoc SX"));	
	xls.SetCellText(5, nRow, _T("Don gia"));
	xls.SetCellText(6, nRow, _T("So luong"));
	xls.SetCellText(7, nRow, _T("Thanh tien"));	
	xls.SetCellText(8, nRow, _T("Han SD"));
	

	grpCost = 0;
	ttlCost = 0;
			
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("pmi_typeid"), szNewLine);
		if(szNewLine != szOldLine && !szNewLine.IsEmpty()){	
			if(grpCost > 0)
			{			
				CString szField, szAmount;
				nRow ++;
				TranslateString(_T("Total Group"), szAmount);												
				xls.SetCellMergedColumns(0,nRow,7);
				xls.SetCellText(0, nRow, szAmount, FMT_TEXT,true,11,0);
				tmpStr.Format(_T("%.2f"), grpCost);
				xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1,true,11,0);
				ttlCost += grpCost;
				grpCost = 0;								
				nItem=1;
			}
			nRow ++;
			CString szField, szAmount;			
			tmpStr.Format(_T("%s"), szNewLine);
			xls.SetCellMergedColumns(0,nRow,7);
			xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);
			szOldLine = szNewLine;
			nItem=1;
		}

		nRow ++;		
		tmpStr.Format(_T("%d"), nItem++);
		xls.SetCellText(0,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("id"), tmpStr);
		xls.SetCellText(1,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("name"), tmpStr);
		xls.SetCellText(2,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("unit"), tmpStr);		
		xls.SetCellText(3,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("contruyid"), tmpStr);
		xls.SetCellText(4,nRow,tmpStr, FMT_TEXT);		
		rs.GetValue(_T("price"), tmpStr);				
		xls.SetCellText(5,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("instockqty"), qty);
		tmpStr.Format(_T("%ld"), qty);
		xls.SetCellText(6,nRow,tmpStr, FMT_NUMBER1);				
		rs.GetValue(_T("instockamount"), cost);				
		tmpStr.Format(_T("%.2f"), cost);
		grpCost+=cost;
		xls.SetCellText(7,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("exdate"), tmpStr);
		xls.SetCellText(8,nRow,CDate::Convert(tmpStr,yyyymmdd,ddmmyyyy), FMT_TEXT);
		rs.MoveNext();
}
	ttlCost += grpCost;
	if(grpCost > 0){
		nRow++;
		TranslateString(_T("Total Group"), tmpStr);		
		xls.SetCellMergedColumns(0,nRow,7);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);
		tmpStr.Format(_T("%.2f"), grpCost);
		xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1,true,11,0);		
	}
	if(ttlCost > 0){
		nRow++;
		TranslateString(_T("Total Amount"), tmpStr);		
		xls.SetCellMergedColumns(0,nRow,7);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);
		tmpStr.Format(_T("%.2f"), ttlCost);
		xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1,true,11,0);		
	}
	xls.Save(_T("Exports\\Danhsachthuochethan.xls"));
	EndWaitCursor();
}
void CPMSReportConditionDialog::OnPrintBaocaosudungthuoctheoquy(bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL,tmpStr2,szObjects,m_szFacList, szWhere;
	CString  m_szServiceFlag,m_szInsFlag,m_szChildFlag,m_szFreeFlag,m_szGroups;
	
	if(!rpt.Init(_T("Reports/HMS/PMS_THONGKETONGHOPSUDUNGTHUOC.RPT"),true) )
	{
		return;
	}
	
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
		CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	rpt.GetReportHeader()->SetValue(_T("StockName"), m_szStockName);
	rpt.GetReportHeader()->SetValue(_T("Type"), m_wndType.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Group"), m_wndGroup.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Original"), m_wndSource.GetCurrent(1));

	if (!m_szStockKey.IsEmpty())	szWhere.Format(_T(" and hpo_stockid=%d "),ToInt(m_szStockKey));


	CString szType, szTypeName;
	for (int i=0 ; i<= m_wndType.GetCount(); i++)
	 {
		 if(m_wndType.GetCheck(i)){
			 m_wndType.SetCurSel(i);
		 	if(!szType.IsEmpty())
					szType += _T(",");						
					szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));			
			if(!szTypeName.IsEmpty())
					szTypeName += _T(", ");						
			szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		 }
	 }
	
	 if(!m_szTypeKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and pmi_typeid in(%s)"), szType);	
	else
		szTypeName.Format(_T("TÊt c¶ thuèc vµ vËt t­"));
	 
	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(pmi_groupid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}
	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and pmsi_originalid=%d "), str2int(m_szSourceKey));
	}

	//Page Header
	//Report Detail

	szSQL.Format(_T(" SELECT hfg_index as idx,hfg_name as typename,pmi_typeid,pmi_id as id, ") \
	_T("		 	pmi_name as name,") \
	_T("		 	pms_generic.pmg_name as genericname,") \
	_T("			pms_group.pmg_id || '> ' || pms_group.pmg_name as groupname, ") \
	_T(" 		 	pmi_unit as unit,") \
	_T("			(SELECT distinct  pmm_name from pms_mfg where pmm_id=pmsi_mfgid) as contruyid,") \
	_T(" 		 	pmsi_vatprice as price,	") \
	_T("			sum(hpol_issueqty) as qty, ")\
	_T(" 		 	sum(hpol_issueqty*pmsi_vatprice) as amount") \
	_T(" FROM hms_patient") \
	_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
	_T(" LEFT JOIN hms_pharmacyorder ON(hpo_docno=hd_docno)") \
	_T(" LEFT JOIN hms_pharmacyorder_line ON (hpol_orderid=hpo_orderid)") \
	_T(" LEFT JOIN pms_stockitems ON(pmsi_id=hpol_sitemid)") \
	_T(" LEFT JOIN pms_items ON(pmsi_itemid=pmi_id)") \
	_T(" LEFT JOIN pms_generic ON(pms_generic.pmg_id=pmi_genericid) ") \
	_T(" LEFT JOIN pms_group ON(pms_group.pmg_id=pmi_groupid) ") \
	_T(" LEFT JOIN hms_fee_group ON(hfg_id = pmi_typeid) ") \
	_T(" LEFT JOIN hms_object ON(hd_object=ho_id)") \
	_T(" WHERE hpo_issuedate between timestamp '%s' and timestamp '%s' AND ho_type in('I','C') and hpo_type in('P','D','M','E') and hpo_status IN('A','I','D') %s") \
	_T(" GROUP BY idx,typename,genericname,groupname,pmi_typeid,pmi_id, pmi_name ,pmi_unit ,pmsi_vatprice, contruyid") \
	_T(" ORDER BY idx,groupname,name,unit"), m_szFromDate, m_szToDate, szWhere );			

	CReportSection* rptDetail;
	CString szOldLine, szNewLine, szAmount;
	CRecord rs(&pMF->m_db);
	long double grpCost=0.0;
	long double ttlCost=0.0;
	double cost = 0;
	int nItem = 1;

	rs.ExecSQL(szSQL);

	if(rs.IsEOF()){
		ShowMessage(150, MB_ICONSTOP);
		return ;
	}

					
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("groupname"), szNewLine);
		if(szNewLine != szOldLine && !szNewLine.IsEmpty()){
			CString szField;
			if(grpCost > 0)
			{
				TranslateString(_T("Total Group"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("TotalGroup"), szAmount);						
				FormatCurrency(grpCost, tmpStr);
				rptDetail->SetValue(_T("s8"), tmpStr);
				ttlCost += grpCost;						
				grpCost = 0;								
			}

			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));			
			rptDetail->SetValue(_T("GroupName"),szNewLine );
			szOldLine = szNewLine;
			nItem=1;
		}

		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nItem++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("genericname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("name"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("contruyid"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rs.GetValue(_T("price"), cost);
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("qty"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rs.GetValue(_T("amount"), cost);	
		grpCost += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.MoveNext();
	}
						
	ttlCost += grpCost;
			
	if(grpCost > 0){	
		TranslateString(_T("Total Group"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));	
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		FormatCurrency(grpCost, tmpStr);
		rptDetail->SetValue(_T("s8"), tmpStr);			
	}
	if(ttlCost > 0){				
		TranslateString(_T("Total Amount"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		FormatCurrency(ttlCost, tmpStr);
		rptDetail->SetValue(_T("s8"), tmpStr);
		ttlCost += grpCost;	
	}
			
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	//	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print();
	EndWaitCursor();
}BEGIN_MESSAGE_MAP(CPMSReportConditionDialog, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSReportConditionDialog::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
