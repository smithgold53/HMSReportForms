#include "PMSUnpaidOrderList.h"
#include "MainFrm.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSUnpaidOrderList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSUnpaidOrderList *)pWnd)->OnToDateCheckValue();
} 
static void _OnItemGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSUnpaidOrderList* )pWnd)->OnItemGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnItemGroupSelendokFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnItemGroupSelendok();
}
/*static void _OnItemGroupSetfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnItemGroupKillfocus();
}*/
/*static void _OnItemGroupKillfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnItemGroupKillfocus();
}*/
static long _OnItemGroupLoadDataFnc(CWnd *pWnd){
	return ((CPMSUnpaidOrderList *)pWnd)->OnItemGroupLoadData();
}
/*static void _OnItemGroupAddNewFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnItemGroupAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSUnpaidOrderList *pVw = (CPMSUnpaidOrderList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSUnpaidOrderList *pVw = (CPMSUnpaidOrderList *)pWnd;
	pVw->OnExportSelect();
} 
CPMSUnpaidOrderList::CPMSUnpaidOrderList(){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CPMSUnpaidOrderList::~CPMSUnpaidOrderList(){
}
void CPMSUnpaidOrderList::OnCreateComponents(){
	m_wndGeneralDepartmentUsage.Create(this, _T("Import Sheet List"), 5, 5, 575, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 270, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 275, 30, 355, 55);
	m_wndToDate.Create(this,360, 30, 570, 55); 
	m_wndItemGroupLabel.Create(this, _T("Item Group"), 10, 60, 90, 85);
	m_wndItemGroup.Create(this,95, 60, 570, 85); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 380, 94, 490, 119);
	m_wndExport.Create(this, _T("&Export"), 495, 95, 575, 120);

}
void CPMSUnpaidOrderList::OnInitializeComponents(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndItemGroup.SetCheckValue(true);
	m_wndItemGroup.LimitText(1024);



}
void CPMSUnpaidOrderList::OnSetWindowEvents(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndItemGroup.SetEvent(WE_SELENDOK, _OnItemGroupSelendokFnc);
	//m_wndItemGroup.SetEvent(WE_SETFOCUS, _OnItemGroupSetfocusFnc);
	//m_wndItemGroup.SetEvent(WE_KILLFOCUS, _OnItemGroupKillfocusFnc);
	m_wndItemGroup.SetEvent(WE_SELCHANGE, _OnItemGroupSelectChangeFnc);
	m_wndItemGroup.SetEvent(WE_LOADDATA, _OnItemGroupLoadDataFnc);
	//m_wndItemGroup.SetEvent(WE_ADDNEW, _OnItemGroupAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	SetMode(VM_NONE);

}
void CPMSUnpaidOrderList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndItemGroup.GetDlgCtrlID(), m_szItemGroupKey);

}
void CPMSUnpaidOrderList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szItemGroupKey.Empty();

}
int CPMSUnpaidOrderList::SetMode(int nMode){
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
/*void CPMSUnpaidOrderList::OnFromDateChange(){
	
} */
/*void CPMSUnpaidOrderList::OnFromDateSetfocus(){
	
} */
/*void CPMSUnpaidOrderList::OnFromDateKillfocus(){
	
} */
int CPMSUnpaidOrderList::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSUnpaidOrderList::OnToDateChange(){
	
} */
/*void CPMSUnpaidOrderList::OnToDateSetfocus(){
	
} */
/*void CPMSUnpaidOrderList::OnToDateKillfocus(){
	
} */
int CPMSUnpaidOrderList::OnToDateCheckValue(){
	return 0;
} 
void CPMSUnpaidOrderList::OnItemGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
void CPMSUnpaidOrderList::OnItemGroupSelendok(){
	 
}
/*void CPMSUnpaidOrderList::OnItemGroupSetfocus(){
	
}*/
/*void CPMSUnpaidOrderList::OnItemGroupKillfocus(){
	
}*/
long CPMSUnpaidOrderList::OnItemGroupLoadData(){
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
/*void CPMSUnpaidOrderList::OnItemGroupAddNew(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} */
void CPMSUnpaidOrderList::OnPrintPreviewSelect(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
void CPMSUnpaidOrderList::OnExportSelect(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
