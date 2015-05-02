#include "stdafx.h"
#include "PMDetailPurchaseOrder.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include "ReportDocument.h"

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
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMDetailPurchaseOrder *pVw = (CPMDetailPurchaseOrder *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMDetailPurchaseOrder *pVw = (CPMDetailPurchaseOrder *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnItemListLoadDataFnc(CWnd *pWnd){
	return ((CPMDetailPurchaseOrder*)pWnd)->OnItemListLoadData();
} 
static void _OnItemListDblClickFnc(CWnd *pWnd){
	((CPMDetailPurchaseOrder*)pWnd)->OnItemListDblClick();
} 
static void _OnItemListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPMDetailPurchaseOrder*)pWnd)->OnItemListSelectChange(nOldItem, nNewItem);
} 
static int _OnItemListDeleteItemFnc(CWnd *pWnd){
	 return ((CPMDetailPurchaseOrder*)pWnd)->OnItemListDeleteItem();
} 
CPMDetailPurchaseOrder::CPMDetailPurchaseOrder(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CPMDetailPurchaseOrder::~CPMDetailPurchaseOrder(){
}
void CPMDetailPurchaseOrder::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 545);
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
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 350, 550, 460, 575);
	m_wndExport.Create(this, _T("&Export"), 465, 550, 575, 575);
	m_wndItemList.Create(this,10, 120, 570, 540); 

}
void CPMDetailPurchaseOrder::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
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


	m_wndItemList.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndItemList.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);

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
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndItemList.SetEvent(WE_SELCHANGE, _OnItemListSelectChangeFnc);
	m_wndItemList.SetEvent(WE_LOADDATA, _OnItemListLoadDataFnc);
	m_wndItemList.SetEvent(WE_DBLCLICK, _OnItemListDblClickFnc);
	m_wndItemList.AddEvent(1, _T("Delete"), _OnItemListDeleteItemFnc, 0, VK_DELETE, 0);
	SetMode(VM_NONE);

}
void CPMDetailPurchaseOrder::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szSupplierKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_TextEx(pDX, m_wndStorage.GetDlgCtrlID(), m_szStorageKey);

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
void CPMDetailPurchaseOrder::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMDetailPurchaseOrder::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMDetailPurchaseOrder::OnItemListDblClick(){
	
} 
void CPMDetailPurchaseOrder::OnItemListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CPMDetailPurchaseOrder::OnItemListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
} 
long CPMDetailPurchaseOrder::OnItemListLoadData(){
/*
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndItemList.BeginLoad(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndItemList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndItemList.EndLoad(); 
	return nCount;
*/
	return 0;
}

CString CPMDetailPurchaseOrder::GetQueryString(){
	CString szSQL;

	return szSQL;
}