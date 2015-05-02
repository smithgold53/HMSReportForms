#include "stdafx.h"
#include "PMDonthuoctonghop.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include ".\pmdonthuoctonghop.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((rptDonthuoctonghop *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSefocusFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((rptDonthuoctonghop *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((rptDonthuoctonghop* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((rptDonthuoctonghop *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnStockAddNew();
}*/
static void _OnTypeSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((rptDonthuoctonghop* )pWnd)->OnTypeSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnTypeSelendokFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnTypeSelendok();
}
/*static void _OnTypeSetfocusFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnTypeKillfocus();
}*/
/*static void _OnTypeKillfocusFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnTypeKillfocus();
}*/
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((rptDonthuoctonghop *)pWnd)->OnTypeLoadData();
}
/*static void _OnTypeAddNewFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnTypeAddNew();
}*/
static void _OnGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((rptDonthuoctonghop* )pWnd)->OnGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnGroupSelendokFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnGroupSelendok();
}
/*static void _OnGroupSetfocusFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnGroupKillfocus();
}*/
/*static void _OnGroupKillfocusFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnGroupKillfocus();
}*/
static long _OnGroupLoadDataFnc(CWnd *pWnd){
	return ((rptDonthuoctonghop *)pWnd)->OnGroupLoadData();
}
/*static void _OnGroupAddNewFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnGroupAddNew();
}*/
static void _OnSourceSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((rptDonthuoctonghop* )pWnd)->OnSourceSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSourceSelendokFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnSourceSelendok();
}
/*static void _OnSourceSetfocusFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnSourceKillfocus();
}*/
/*static void _OnSourceKillfocusFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnSourceKillfocus();
}*/
static long _OnSourceLoadDataFnc(CWnd *pWnd){
	return ((rptDonthuoctonghop *)pWnd)->OnSourceLoadData();
}
/*static void _OnSourceAddNewFnc(CWnd *pWnd){
	((rptDonthuoctonghop *)pWnd)->OnSourceAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	rptDonthuoctonghop *pVw = (rptDonthuoctonghop *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	rptDonthuoctonghop *pVw = (rptDonthuoctonghop *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	rptDonthuoctonghop *pVw = (rptDonthuoctonghop *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	rptDonthuoctonghop *pVw = (rptDonthuoctonghop *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddrptDonthuoctonghopFnc(CWnd *pWnd){
	 return ((rptDonthuoctonghop*)pWnd)->OnAddrptDonthuoctonghop();
} 
static int _OnEditrptDonthuoctonghopFnc(CWnd *pWnd){
	 return ((rptDonthuoctonghop*)pWnd)->OnEditrptDonthuoctonghop();
} 
static int _OnDeleterptDonthuoctonghopFnc(CWnd *pWnd){
	 return ((rptDonthuoctonghop*)pWnd)->OnDeleterptDonthuoctonghop();
} 
static int _OnSaverptDonthuoctonghopFnc(CWnd *pWnd){
	 return ((rptDonthuoctonghop*)pWnd)->OnSaverptDonthuoctonghop();
} 
static int _OnCancelrptDonthuoctonghopFnc(CWnd *pWnd){
	 return ((rptDonthuoctonghop*)pWnd)->OnCancelrptDonthuoctonghop();
} 
rptDonthuoctonghop::rptDonthuoctonghop(CWnd *pParent)
	{

	m_nDlgWidth = 565;
	m_nDlgHeight = 220;
	SetDefaultValues();
}
rptDonthuoctonghop::~rptDonthuoctonghop(){
}
void rptDonthuoctonghop::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 180);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 110, 55);
	m_wndFromDate.Create(this,115, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 110, 85);
	m_wndStock.Create(this,115, 60, 570, 85); 
	m_wndTypeLabel.Create(this, _T("Type"), 10, 90, 110, 115);
	m_wndType.Create(this,115, 90, 570, 115); 
	m_wndGroupLabel.Create(this, _T("Group"), 10, 120, 110, 145);
	m_wndGroup.Create(this,115, 120, 570, 145); 
	m_wndSourceLabel.Create(this, _T("Source"), 10, 150, 110, 175);
	m_wndSource.Create(this,115, 150, 570, 175); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 380, 185, 490, 210);
	m_wndExport.Create(this, _T("&Export"), 495, 185, 575, 210);
}
void rptDonthuoctonghop::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);
	

	m_wndStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStock.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);


	m_wndType.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 50);
	m_wndType.InsertColumn(1, _T("Description"), CFMT_TEXT, 300);

	m_wndGroup.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndGroup.InsertColumn(1, _T("Description"), CFMT_TEXT, 500);


	m_wndSource.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndSource.InsertColumn(1, _T("Description"), CFMT_TEXT, 125);

}
void rptDonthuoctonghop::OnSetWindowEvents(){
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
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_szFromDate =  pMF->GetSysDate()+ _T("00:00");	
	m_szToDate =pMF->GetSysDate() + _T("23:59");
	SetMode(VM_EDIT);

}
void rptDonthuoctonghop::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_TextEx(pDX, m_wndGroup.GetDlgCtrlID(), m_szGroupKey);
	DDX_TextEx(pDX, m_wndSource.GetDlgCtrlID(), m_szSourceKey);
}
void rptDonthuoctonghop::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void rptDonthuoctonghop::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void rptDonthuoctonghop::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_szTypeKey.Empty();
	m_szGroupKey.Empty();
	m_szSourceKey.Empty();

}
int rptDonthuoctonghop::SetMode(int nMode){
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
/*void rptDonthuoctonghop::OnFromDateChange(){
	
} */
/*void rptDonthuoctonghop::OnFromDateSetfocus(){
	
} */
/*void rptDonthuoctonghop::OnFromDateKillfocus(){
	
} */
int rptDonthuoctonghop::OnFromDateCheckValue(){
	return 0;
} 
/*void rptDonthuoctonghop::OnToDateChange(){
	
} */
/*void rptDonthuoctonghop::OnToDateSetfocus(){
	
} */
/*void rptDonthuoctonghop::OnToDateKillfocus(){
	
} */
int rptDonthuoctonghop::OnToDateCheckValue(){
	return 0;
} 
void rptDonthuoctonghop::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void rptDonthuoctonghop::OnStockSelendok(){
	 
}
/*void rptDonthuoctonghop::OnStockSetfocus(){
	
}*/
/*void rptDonthuoctonghop::OnStockKillfocus(){
	
}*/
long rptDonthuoctonghop::OnStockLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageList(&m_wndStock, m_szStockKey);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStock.IsSearchKey() && ToInt(m_szStockKey) > 0){
		szWhere.Format(_T(" AND msl_storage_id=%d"), ToInt(m_szStockKey));
	};
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
/*void rptDonthuoctonghop::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void rptDonthuoctonghop::OnTypeSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void rptDonthuoctonghop::OnTypeSelendok(){
	 
}
/*void rptDonthuoctonghop::OnTypeSetfocus(){
	
}*/
/*void rptDonthuoctonghop::OnTypeKillfocus(){
	
}*/
long rptDonthuoctonghop::OnTypeLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndType.IsSearchKey() && !m_szTypeKey.IsEmpty()){
		//szWhere.Format(_T(" WHERE pmdt_id='%s' "), m_szTypeKey);
	};
	//szWhere.Format(_T(" AND hfg_moduleid = '%s'"), pMF->GetModuleID());
	m_wndType.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT hfg_id as id, hfg_name as name FROM hms_fee_group WHERE substr(hfg_id,1, 1)='A' %s ORDER BY hfg_index, hfg_id "), szWhere);

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndType.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), 			
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void rptDonthuoctonghop::OnTypeAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void rptDonthuoctonghop::OnGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void rptDonthuoctonghop::OnGroupSelendok(){
	 
}
/*void rptDonthuoctonghop::OnGroupSetfocus(){
	
}*/
/*void rptDonthuoctonghop::OnGroupKillfocus(){
	
}*/
long rptDonthuoctonghop::OnGroupLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndGroup.IsSearchKey() && !m_szGroupKey.IsEmpty()){
		szWhere.Format(_T(" AND pmg_id='%s' "), m_szGroupKey);
	};
	m_wndGroup.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT pmg_id as id, pmg_name as name FROM pms_group WHERE pmg_moduleid = '%s' %s ORDER BY pmg_id "), pMF->GetModuleID(), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndGroup.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void rptDonthuoctonghop::OnGroupAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void rptDonthuoctonghop::OnSourceSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void rptDonthuoctonghop::OnSourceSelendok(){
	 
}
/*void rptDonthuoctonghop::OnSourceSetfocus(){
	
}*/
/*void rptDonthuoctonghop::OnSourceKillfocus(){
	
}*/
long rptDonthuoctonghop::OnSourceLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndSource.IsSearchKey() && str2int(m_szSourceKey) > 0){
		szWhere.Format(_T(" and pmo_id=%d"), str2int(m_szSourceKey));
	};
	szWhere.AppendFormat(_T(" AND pmo_moduleid = '%s'"), pMF->GetModuleID());
	m_wndSource.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT * FROM pms_original WHERE pmo_active='Y' %s ORDER BY pmo_id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndSource.AddItems(
			rs.GetValue(_T("pmo_id")), 
			rs.GetValue(_T("pmo_name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void rptDonthuoctonghop::OnSourceAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void rptDonthuoctonghop::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_szStockName = m_wndStock.GetCurrent(1);
	OnPrint(true);

} 
void rptDonthuoctonghop::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_szStockName = m_wndStock.GetCurrent(1);
	OnPrint(false);
} 
void rptDonthuoctonghop::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_szStockName = m_wndStock.GetCurrent(1);
	OnExportToXLS();
} 
void rptDonthuoctonghop::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int rptDonthuoctonghop::OnAddrptDonthuoctonghop(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int rptDonthuoctonghop::OnEditrptDonthuoctonghop(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int rptDonthuoctonghop::OnDeleterptDonthuoctonghop(){
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
 		OnCancelrptDonthuoctonghop(); 
 	}
	return 0;
}
int rptDonthuoctonghop::OnSaverptDonthuoctonghop(){
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
 		//OnrptDonthuoctonghopListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int rptDonthuoctonghop::OnCancelrptDonthuoctonghop(){
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
int rptDonthuoctonghop::OnrptDonthuoctonghopListLoadData(){
	return 0;
}
/*------------------------------------------------------------------------------------------------------------------------------
	Print Bao cao don thuoc tong hop
------------------------------------------------------------------------------------------------------------------------------*/
void rptDonthuoctonghop::OnPrint(bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL, szWhere, szTemp;

	if(!rpt.Init(_T("Reports/HMS/PMS_DONTHUOCTONGHOP.RPT"),TRUE) )
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

			//Page Header
			//Report Detail

	CString szType, szTypeName;
	for (int i=0 ; i<= m_wndType.GetCount(); i++)
	{
		 if(m_wndType.GetCheck(i))
		 {
			 m_wndType.SetCurSel(i);
			 if(!szType.IsEmpty())
			 {
				szType += _T(",");
				szTypeName +=_T(", ");

			 }
			 szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));				
			 szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		}		
	 }
	

	if(!szType.IsEmpty())	
		szWhere.AppendFormat(_T(" and pmi_typeid in(%s)"), szType);	
	else
		szTypeName.Format(_T("T\x1EA5t \x63\x1EA3 lo\x1EA1i thu\x1ED1\x63"));

	if(!m_szGroupKey.IsEmpty())
	{
			szWhere.AppendFormat(_T(" and substr(pmi_groupid, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}
	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and pmsi_originalid=%d "), str2int(m_szSourceKey));
	}
	if (!m_szStockKey.IsEmpty())
	{
		szWhere.AppendFormat(_T (" and expstockid =%d"), ToInt(m_szStockKey));
	}
	szSQL.Format(_T("  SELECT  ") \
				_T(" 	mp_name as name, ") \
				_T(" 	mpc_name as genericname, ") \
				_T(" 	get_uomname(mp_uom_id) as unit, ") \
				_T(" 	mpi_taxprice as price, ") \
				_T(" 	sum(expqty) as qty,") \
				_T(" 	sum(expqty*mpi_taxprice) as amount") \
				_T(" FROM mev") \
				_T(" LEFT JOIN m_product_item ON(mpi_product_item_id=sitemid)") \
				_T(" LEFT JOIN m_product ON(mpi_product_id=mp_product_id)") \
				_T(" LEFT JOIN m_product_class ON(mpc_product_class_id=mp_product_class_id) ") \
				_T(" WHERE sitemid > 0 AND iotype = 'PPO' and expdate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s ") \
				_T(" GROUP BY mp_name,mpc_name,mp_uom_id,mpi_taxprice") \
				_T(" ORDER BY name,unit"), m_szFromDate, m_szToDate, szWhere);	
		
		//_msg(_T("%s"), szSQL);
	
		CReportSection* rptDetail;
		CString szOldLine, szNewLine;
		CRecord rs(&pMF->m_db);
		long double grpCost=0.0;
		long double ttlCost=0.0;
		double ttcost = 0;
		int nItem = 1;
		

		rs.ExecSQL(szSQL);
		if(rs.IsEOF()){
			ShowMessage(150, MB_ICONSTOP);
			return ;
		}
		
		
		while(!rs.IsEOF())
		{
			rptDetail = rpt.AddDetail();
			tmpStr.Format(_T("%d"), nItem++);
			rptDetail->SetValue(_T("1"), tmpStr);
			rs.GetValue(_T("genericname"), tmpStr);
			rptDetail->SetValue(_T("2"), tmpStr);
			rs.GetValue(_T("name"), tmpStr);
			rptDetail->SetValue(_T("3"), tmpStr);
			rs.GetValue(_T("unit"), tmpStr);
			rptDetail->SetValue(_T("4"), tmpStr);
			rs.GetValue(_T("price"), tmpStr);
			rptDetail->SetValue(_T("5"), tmpStr);
			rs.GetValue(_T("qty"), tmpStr);
			rptDetail->SetValue(_T("6"), tmpStr);
			rs.GetValue(_T("amount"), ttcost);
			ttlCost += ttcost;
			FormatCurrency(ttcost, tmpStr);
			rptDetail->SetValue(_T("7"), tmpStr);
			rs.MoveNext();

		}
		if (ttlCost >0)
		{
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			tmpStr.Format(_T("T\x1ED5ng ti\x1EC1n:"));
			rptDetail->SetValue(_T("1"),tmpStr );				
			FormatCurrency(ttlCost, tmpStr);
			rptDetail->SetValue(_T("5"), tmpStr);			
		}
		szSysDate = pMF->GetSysDate(); 
		szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
		rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
		//	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
		if(bPreview)
			rpt.PrintPreview();
		else
			rpt.Print();

}

/*------------------------------------------------------------------------------------------------------------------------------
	Export To Excell Bao cao tong hop nhap xuat ton kho
------------------------------------------------------------------------------------------------------------------------------*/
void rptDonthuoctonghop::OnExportToXLS(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CString tmpStr,szSQL, szWhere, szTemp;

	CString szType, szTypeName;
	for (int i=0 ; i<= m_wndType.GetCount(); i++)
	{
		 if(m_wndType.GetCheck(i))
		 {
			 m_wndType.SetCurSel(i);
			 if(!szType.IsEmpty())
			 {
				szType += _T(",");
				szTypeName +=_T(", ");

			 }
			 szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));				
			 szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		}		
	 }
	

	if(!szType.IsEmpty())	
		szWhere.AppendFormat(_T(" and mp_producttype in(%s)"), szType);	
	else
		szTypeName.Format(_T("T\x1EA5t \x63\x1EA3 lo\x1EA1i thu\x1ED1\x63"));

	if(!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and substr(mp_product_class_id, 1, %d)= '%s' "), m_szGroupKey.GetLength(), m_szGroupKey);
	}

	if(str2int(m_szSourceKey) > 0)
	{
		szWhere.AppendFormat(_T(" and mpi_resource_id=%d "), str2int(m_szSourceKey));
	}
	if (!m_szStockKey.IsEmpty())
		szWhere.AppendFormat(_T(" and expstockid = %d"), ToInt(m_szStockKey));
	szSQL.Format(_T("  SELECT  ") \
				_T(" 	mp_name as name, ") \
				_T(" 	mpc_name as genericname, ") \
				_T(" 	get_uomname(mp_uom_id) as unit, ") \
				_T(" 	mpi_taxprice as price, ") \
				_T(" 	sum(expqty) as qty,") \
				_T(" 	sum(expqty*mpi_taxprice) as amount") \
				_T(" FROM mev") \
				_T(" LEFT JOIN m_product_item ON(mpi_product_item_id=sitemid)") \
				_T(" LEFT JOIN m_product ON(mpi_product_id=mp_product_id)") \
				_T(" LEFT JOIN m_product_class ON(mpc_product_class_id=mp_product_class_id) ") \
				_T(" WHERE sitemid > 0 AND iotype = 'PPO' and expdate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s ") \
				_T(" GROUP BY mp_name,mpc_name,mp_uom_id,mpi_taxprice") \
				_T(" ORDER BY name,unit"), m_szFromDate, m_szToDate, szWhere);	
	CReportSection* rptDetail=NULL;
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
	CellFormat df(&xls), cf(&xls), hf(&xls);
	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER);
	cf.SetFontName(_T("Segoe UI"));

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 35);
	xls.SetColumnWidth(2, 35);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(5, 10);
	xls.SetColumnWidth(6, 15);	
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellText(0, 0, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(0, 2, 7);
	TranslateString(_T("General Prescription"), szTemp);
	StringUpper(szTemp, tmpStr);
	xls.SetCellText(0, 2, tmpStr,FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(0, 3, 7);
	tmpStr.Format(_T("T\x1EEB ngày %s \x111\x1EBFn ngày %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(0, 3, tmpStr, &df);

	int nRow = 3;
	if (!m_szStockKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("T\x1EEB kho: %s"), m_wndStock.GetCurrent(1));
		xls.SetCellMergedColumns(0, nRow, 7);
		xls.SetCellText(0, nRow, tmpStr, &cf);
	}
	if(!m_szTypeKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("Lo\x1EA1i: %s"), m_wndType.GetCurrent(1));
		xls.SetCellMergedColumns(0, nRow, 7);
		xls.SetCellText(0, nRow, tmpStr, &cf);
	}
	if(!m_szGroupKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("Nhóm: %s"), m_wndGroup.GetCurrent(1));
		xls.SetCellMergedColumns(0, nRow, 7);
		xls.SetCellText(0, nRow, tmpStr, &cf);
	}
	if(!m_szSourceKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("Ngu\x1ED3n: %s"), m_wndSource.GetCurrent(1));
		xls.SetCellMergedColumns(0, nRow, 7);
		xls.SetCellText(0, nRow, tmpStr, &cf);
	}
	nRow++;
	xls.SetCellText(0, nRow, _T("STT"), &hf);	
	xls.SetCellText(1, nRow, _T("Tên thu\x1ED1\x63/ HL"), &hf);	
	xls.SetCellText(2, nRow, _T("T\xEAn ho\x1EA1t \x63h\x1EA5t"), &hf);	
	xls.SetCellText(3, nRow, _T("\x110\x1A1n v\x1ECB"), &hf);	
	xls.SetCellText(4, nRow, _T("\x110\x1A1n giá"), &hf);	
	xls.SetCellText(5, nRow, _T("S\x1ED1 l\x1B0\x1EE3ng"), &hf);	
	xls.SetCellText(6, nRow, _T("Th\xE0nh ti\x1EC1n"), &hf);	
		
	double Amount=0.0, ttCost =0.0;		
	while(!rs.IsEOF())
	{
		//rs.GetValue(_T("expinvoice"), szNewLine);
		//if(szNewLine != szOldLine && !szNewLine.IsEmpty())
		//{
		//	nRow++;
		//	CString szField, szAmount;
		//	xls.SetCellText(1, nRow, szNewLine, FMT_TEXT,true);			
		//	rs.GetValue(_T("iotype"), tmpStr);
		//	xls.SetCellText(2,nRow,pMF->GetSelectionString(_T("pms_export_type"), tmpStr),FMT_TEXT,true);
		//	rs.GetValue(_T("expdate"), tmpStr);
		//	xls.SetCellText(3, nRow, CDate::Convert(tmpStr,yyyymmdd,ddmmyyyy), FMT_DATE,true);
		//	rs.GetValue(_T("deptid"), tmpStr);
		//	xls.SetCellText(4,nRow,tmpStr,FMT_TEXT, true);			

		//	if (grpCost >0)
		//	{
		//		tmpStr.Format(_T("%.2f"), grpCost);
		//		xls.SetCellText(6,nRow-nItem,tmpStr,FMT_NUMBER1,true);
		//		ttCost += grpCost;	
		//	}

		//	szOldLine = szNewLine;
		//	nItem=1;
		//	grpCost= 0;
		//}

		nRow ++;
		tmpStr.Format(_T("%d"), nItem++);
		xls.SetCellText(0,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("name"), tmpStr);
		xls.SetCellText(1,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("genericname"), tmpStr);
		xls.SetCellText(2,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("unit"), tmpStr);		
		xls.SetCellText(3,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("price"), tmpStr);
		xls.SetCellText(4,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("qty"), tmpStr);
		xls.SetCellText(5,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("amount"), Amount);		
		ttCost +=Amount;
		tmpStr.Format(_T("%.2f"), Amount);
		xls.SetCellText(6,nRow,tmpStr, FMT_NUMBER1);		
		rs.MoveNext();
	}

	//if (grpCost >0)
	//{	nRow++;
	//	tmpStr.Format(_T("%.2f"), grpCost);
	//	xls.SetCellText(6,nRow-nItem,tmpStr,FMT_NUMBER1,true);
	//	ttCost += grpCost;	
	//}
	
	if (ttCost >0)	{
		nRow++;
		xls.SetCellText(1,nRow,_T("T\x1ED5ng ti\x1EC1n:"), FMT_TEXT,true);
		tmpStr.Format(_T("%.2f"), ttCost);
		xls.SetCellText(6,nRow,tmpStr,FMT_NUMBER1,true);
	}
	
	xls.Save(_T("Exports\\Don Thuoc Tong Hop.XLS"));
}



BEGIN_MESSAGE_MAP(rptDonthuoctonghop, CGuiView)
	ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void rptDonthuoctonghop::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
