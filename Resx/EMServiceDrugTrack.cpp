#include "EMServiceDrugTrack.h"
#include "MainFrm.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMServiceDrugTrack *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMServiceDrugTrack *)pWnd)->OnToDateCheckValue();
} 
static void _OnDrugNameSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMServiceDrugTrack* )pWnd)->OnDrugNameSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDrugNameSelendokFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnDrugNameSelendok();
}
/*static void _OnDrugNameSetfocusFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnDrugNameKillfocus();
}*/
/*static void _OnDrugNameKillfocusFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnDrugNameKillfocus();
}*/
static long _OnDrugNameLoadDataFnc(CWnd *pWnd){
	return ((CEMServiceDrugTrack *)pWnd)->OnDrugNameLoadData();
}
/*static void _OnDrugNameAddNewFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnDrugNameAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMServiceDrugTrack *pVw = (CEMServiceDrugTrack *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMServiceDrugTrack *pVw = (CEMServiceDrugTrack *)pWnd;
	pVw->OnExportSelect();
} 
CEMServiceDrugTrack::CEMServiceDrugTrack(){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CEMServiceDrugTrack::~CEMServiceDrugTrack(){
}
void CEMServiceDrugTrack::OnCreateComponents(){
	m_wndServiceDrugTrack.Create(this, _T("Service Drug Track"), 5, 5, 490, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndDrugNameLabel.Create(this, _T("Drug Name"), 10, 60, 90, 85);
	m_wndDrugName.Create(this,95, 60, 485, 85); 
	m_wndPrint.Create(this, _T("&Print"), 325, 100, 405, 125);
	m_wndExport.Create(this, _T("&Export"), 410, 100, 490, 125);

}
void CEMServiceDrugTrack::OnInitializeComponents(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndDrugName.SetCheckValue(true);
	m_wndDrugName.LimitText(1024);



}
void CEMServiceDrugTrack::OnSetWindowEvents(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndDrugName.SetEvent(WE_SELENDOK, _OnDrugNameSelendokFnc);
	//m_wndDrugName.SetEvent(WE_SETFOCUS, _OnDrugNameSetfocusFnc);
	//m_wndDrugName.SetEvent(WE_KILLFOCUS, _OnDrugNameKillfocusFnc);
	m_wndDrugName.SetEvent(WE_SELCHANGE, _OnDrugNameSelectChangeFnc);
	m_wndDrugName.SetEvent(WE_LOADDATA, _OnDrugNameLoadDataFnc);
	//m_wndDrugName.SetEvent(WE_ADDNEW, _OnDrugNameAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	SetMode(VM_NONE);

}
void CEMServiceDrugTrack::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDrugName.GetDlgCtrlID(), m_szDrugNameKey);

}
void CEMServiceDrugTrack::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDrugNameKey.Empty();

}
int CEMServiceDrugTrack::SetMode(int nMode){
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
/*void CEMServiceDrugTrack::OnFromDateChange(){
	
} */
/*void CEMServiceDrugTrack::OnFromDateSetfocus(){
	
} */
/*void CEMServiceDrugTrack::OnFromDateKillfocus(){
	
} */
int CEMServiceDrugTrack::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMServiceDrugTrack::OnToDateChange(){
	
} */
/*void CEMServiceDrugTrack::OnToDateSetfocus(){
	
} */
/*void CEMServiceDrugTrack::OnToDateKillfocus(){
	
} */
int CEMServiceDrugTrack::OnToDateCheckValue(){
	return 0;
} 
void CEMServiceDrugTrack::OnDrugNameSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
void CEMServiceDrugTrack::OnDrugNameSelendok(){
	 
}
/*void CEMServiceDrugTrack::OnDrugNameSetfocus(){
	
}*/
/*void CEMServiceDrugTrack::OnDrugNameKillfocus(){
	
}*/
long CEMServiceDrugTrack::OnDrugNameLoadData(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDrugName.IsSearchKey() && !m_szDrugNameKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szDrugNameKey
};
	m_wndDrugName.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDrugName.AddItems(
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CEMServiceDrugTrack::OnDrugNameAddNew(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} */
void CEMServiceDrugTrack::OnPrintSelect(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
void CEMServiceDrugTrack::OnExportSelect(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
