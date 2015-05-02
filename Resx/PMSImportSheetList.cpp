#include "PMSImportSheetList.h"
#include "MainFrm.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSImportSheetList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSImportSheetList *)pWnd)->OnToDateCheckValue();
} 
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSImportSheetList* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMSImportSheetList *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnStorageAddNew();
}*/
static void _OnItemGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSImportSheetList* )pWnd)->OnItemGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnItemGroupSelendokFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnItemGroupSelendok();
}
/*static void _OnItemGroupSetfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnItemGroupKillfocus();
}*/
/*static void _OnItemGroupKillfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnItemGroupKillfocus();
}*/
static long _OnItemGroupLoadDataFnc(CWnd *pWnd){
	return ((CPMSImportSheetList *)pWnd)->OnItemGroupLoadData();
}
/*static void _OnItemGroupAddNewFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnItemGroupAddNew();
}*/
static void _OnInvoiceImportSelectFnc(CWnd *pWnd){
	 ((CPMSImportSheetList*)pWnd)->OnInvoiceImportSelect();
}
static void _OnNonInvoiceImportSelectFnc(CWnd *pWnd){
	 ((CPMSImportSheetList*)pWnd)->OnNonInvoiceImportSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSImportSheetList *pVw = (CPMSImportSheetList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSImportSheetList *pVw = (CPMSImportSheetList *)pWnd;
	pVw->OnExportSelect();
} 
CPMSImportSheetList::CPMSImportSheetList(){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CPMSImportSheetList::~CPMSImportSheetList(){
}
void CPMSImportSheetList::OnCreateComponents(){
	m_wndGeneralDepartmentUsage.Create(this, _T("Import Sheet List"), 5, 5, 575, 120);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 285, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 290, 30, 370, 55);
	m_wndToDate.Create(this,375, 30, 570, 55); 
	m_wndStorageLabel.Create(this, _T("Storage"), 10, 60, 90, 85);
	m_wndStorage.Create(this,95, 60, 570, 85); 
	m_wndItemGroupLabel.Create(this, _T("Item Group"), 10, 90, 90, 115);
	m_wndItemGroup.Create(this,95, 90, 570, 115); 
	m_wndInvoiceImport.Create(this, _T("Invoice Import"), 5, 125, 140, 150);
	m_wndNonInvoiceImport.Create(this, _T("NonInvoice Import"), 145, 125, 280, 150);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 380, 125, 490, 150);
	m_wndExport.Create(this, _T("&Export"), 495, 125, 575, 150);

}
void CPMSImportSheetList::OnInitializeComponents(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndStorage.SetCheckValue(true);
	m_wndStorage.LimitText(1024);
	m_wndItemGroup.SetCheckValue(true);
	m_wndItemGroup.LimitText(1024);





}
void CPMSImportSheetList::OnSetWindowEvents(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndStorage.SetEvent(WE_SELENDOK, _OnStorageSelendokFnc);
	//m_wndStorage.SetEvent(WE_SETFOCUS, _OnStorageSetfocusFnc);
	//m_wndStorage.SetEvent(WE_KILLFOCUS, _OnStorageKillfocusFnc);
	m_wndStorage.SetEvent(WE_SELCHANGE, _OnStorageSelectChangeFnc);
	m_wndStorage.SetEvent(WE_LOADDATA, _OnStorageLoadDataFnc);
	//m_wndStorage.SetEvent(WE_ADDNEW, _OnStorageAddNewFnc);
	m_wndItemGroup.SetEvent(WE_SELENDOK, _OnItemGroupSelendokFnc);
	//m_wndItemGroup.SetEvent(WE_SETFOCUS, _OnItemGroupSetfocusFnc);
	//m_wndItemGroup.SetEvent(WE_KILLFOCUS, _OnItemGroupKillfocusFnc);
	m_wndItemGroup.SetEvent(WE_SELCHANGE, _OnItemGroupSelectChangeFnc);
	m_wndItemGroup.SetEvent(WE_LOADDATA, _OnItemGroupLoadDataFnc);
	//m_wndItemGroup.SetEvent(WE_ADDNEW, _OnItemGroupAddNewFnc);
	m_wndInvoiceImport.SetEvent(WE_CLICK, _OnInvoiceImportSelectFnc);
	m_wndNonInvoiceImport.SetEvent(WE_CLICK, _OnNonInvoiceImportSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	SetMode(VM_NONE);

}
void CPMSImportSheetList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStorage.GetDlgCtrlID(), m_szStorageKey);
	DDX_TextEx(pDX, m_wndItemGroup.GetDlgCtrlID(), m_szItemGroupKey);
	DDX_Check(pDX, m_wndInvoiceImport.GetDlgCtrlID(), m_bInvoiceImport);
	DDX_Check(pDX, m_wndNonInvoiceImport.GetDlgCtrlID(), m_bNonInvoiceImport);

}
void CPMSImportSheetList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStorageKey.Empty();
	m_szItemGroupKey.Empty();
	m_bInvoiceImport=FALSE;
	m_bNonInvoiceImport=FALSE;

}
int CPMSImportSheetList::SetMode(int nMode){
 		int nOldMode = GetMode();
 		CGuiView::SetMode(nMode);
 		CMainFrame *pMF = (CMainFrame *) AfxGetMainWnd();
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
/*void CPMSImportSheetList::OnFromDateChange(){
	
} */
/*void CPMSImportSheetList::OnFromDateSetfocus(){
	
} */
/*void CPMSImportSheetList::OnFromDateKillfocus(){
	
} */
int CPMSImportSheetList::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSImportSheetList::OnToDateChange(){
	
} */
/*void CPMSImportSheetList::OnToDateSetfocus(){
	
} */
/*void CPMSImportSheetList::OnToDateKillfocus(){
	
} */
int CPMSImportSheetList::OnToDateCheckValue(){
	return 0;
} 
void CPMSImportSheetList::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
void CPMSImportSheetList::OnStorageSelendok(){
	 
}
/*void CPMSImportSheetList::OnStorageSetfocus(){
	
}*/
/*void CPMSImportSheetList::OnStorageKillfocus(){
	
}*/
long CPMSImportSheetList::OnStorageLoadData(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
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
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CPMSImportSheetList::OnStorageAddNew(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} */
void CPMSImportSheetList::OnItemGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
void CPMSImportSheetList::OnItemGroupSelendok(){
	 
}
/*void CPMSImportSheetList::OnItemGroupSetfocus(){
	
}*/
/*void CPMSImportSheetList::OnItemGroupKillfocus(){
	
}*/
long CPMSImportSheetList::OnItemGroupLoadData(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndItemGroup.IsSearchKey() && !m_szItemGroupKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szItemGroupKey
};
	m_wndItemGroup.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndItemGroup.AddItems(
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CPMSImportSheetList::OnItemGroupAddNew(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} */
	void CPMSImportSheetList::OnInvoiceImportSelect(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
}
	void CPMSImportSheetList::OnNonInvoiceImportSelect(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
}
void CPMSImportSheetList::OnPrintPreviewSelect(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
void CPMSImportSheetList::OnExportSelect(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
