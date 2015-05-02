#include "stdafx.h"
#include "PMBaocaoxuatthuocchokhoatheodoituong.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include ".\pmbaocaoxuatthuocchokhoatheodoituong.h"

static int _OnCheckAllFnc(CWnd *pWnd){
	 return ((CrptBaocaoxuatthuocchokhoatheodoituong*)pWnd)->OnCheckAll();
} 
static int _OnUnCheckAllFnc(CWnd *pWnd){
	 return ((CrptBaocaoxuatthuocchokhoatheodoituong*)pWnd)->OnUnCheckAll();
} 
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CrptBaocaoxuatthuocchokhoatheodoituong* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnStockAddNew();
}*/
static void _OnTypeSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CrptBaocaoxuatthuocchokhoatheodoituong* )pWnd)->OnTypeSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnTypeSelendokFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnTypeSelendok();
}
/*static void _OnTypeSetfocusFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnTypeKillfocus();
}*/
/*static void _OnTypeKillfocusFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnTypeKillfocus();
}*/
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnTypeLoadData();
}
/*static void _OnTypeAddNewFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnTypeAddNew();
}*/

static void _OnOriginalSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CrptBaocaoxuatthuocchokhoatheodoituong* )pWnd)->OnOriginalSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnOriginalSelendokFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnOriginalSelendok();
}
/*static void _OnOriginalSetfocusFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnOriginalKillfocus();
}*/
/*static void _OnOriginalKillfocusFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnOriginalKillfocus();
}*/
static long _OnOriginalLoadDataFnc(CWnd *pWnd){
	return ((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnOriginalLoadData();
}
/*static void _OnOriginalAddNewFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd)->OnOriginalAddNew();
}*/


static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CrptBaocaoxuatthuocchokhoatheodoituong *pVw = (CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CrptBaocaoxuatthuocchokhoatheodoituong *pVw = (CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CrptBaocaoxuatthuocchokhoatheodoituong *pVw = (CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CrptBaocaoxuatthuocchokhoatheodoituong *pVw = (CrptBaocaoxuatthuocchokhoatheodoituong *)pWnd;
	pVw->OnCloseSelect();
} 
static long _OnListFacLoadDataFnc(CWnd *pWnd){
	return ((CrptBaocaoxuatthuocchokhoatheodoituong*)pWnd)->OnListFacLoadData();
} 
static void _OnListFacDblClickFnc(CWnd *pWnd){
	((CrptBaocaoxuatthuocchokhoatheodoituong*)pWnd)->OnListFacDblClick();
} 
static void _OnListFacSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CrptBaocaoxuatthuocchokhoatheodoituong*)pWnd)->OnListFacSelectChange(nOldItem, nNewItem);
} 
static int _OnListFacDeleteItemFnc(CWnd *pWnd){
	 return ((CrptBaocaoxuatthuocchokhoatheodoituong*)pWnd)->OnListFacDeleteItem();
} 
static int _OnAddrptBaocaoxuatthuocchokhoatheodoituongFnc(CWnd *pWnd){
	 return ((CrptBaocaoxuatthuocchokhoatheodoituong*)pWnd)->OnAddrptBaocaoxuatthuocchokhoatheodoituong();
} 
static int _OnEditrptBaocaoxuatthuocchokhoatheodoituongFnc(CWnd *pWnd){
	 return ((CrptBaocaoxuatthuocchokhoatheodoituong*)pWnd)->OnEditrptBaocaoxuatthuocchokhoatheodoituong();
} 
static int _OnDeleterptBaocaoxuatthuocchokhoatheodoituongFnc(CWnd *pWnd){
	 return ((CrptBaocaoxuatthuocchokhoatheodoituong*)pWnd)->OnDeleterptBaocaoxuatthuocchokhoatheodoituong();
} 
static int _OnSaverptBaocaoxuatthuocchokhoatheodoituongFnc(CWnd *pWnd){
	 return ((CrptBaocaoxuatthuocchokhoatheodoituong*)pWnd)->OnSaverptBaocaoxuatthuocchokhoatheodoituong();
} 
static int _OnCancelrptBaocaoxuatthuocchokhoatheodoituongFnc(CWnd *pWnd){
	 return ((CrptBaocaoxuatthuocchokhoatheodoituong*)pWnd)->OnCancelrptBaocaoxuatthuocchokhoatheodoituong();
} 
CrptBaocaoxuatthuocchokhoatheodoituong::CrptBaocaoxuatthuocchokhoatheodoituong(CWnd *pParent)
	{

	m_nDlgWidth = 585;
	m_nDlgHeight = 590;
//	m_szTitle=_T("Bao cao xuat thuoc cho khoa");
	SetDefaultValues();
}
CrptBaocaoxuatthuocchokhoatheodoituong::~CrptBaocaoxuatthuocchokhoatheodoituong(){
}
void CrptBaocaoxuatthuocchokhoatheodoituong::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 550);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 90, 85);
	m_wndStock.Create(this,95, 60, 570, 85); 
	m_wndTypeLabel.Create(this, _T("Type"), 10, 90, 90, 115);
	m_wndType.Create(this,95, 90, 570, 115); 
	m_wndOriginalLabel.Create(this, _T("Original"), 10, 120, 90, 145);
	m_wndOriginal.Create(this,95, 120, 570, 145); 
	m_wndCheck.Create(this, _T("Sort by group"), 5, 555, 160, 580);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 380, 555, 490, 580);
	m_wndExport.Create(this, _T("&Export"), 495, 555, 575, 580);
	m_wndListFac.Create(this,10, 150, 570, 545); 

}
void CrptBaocaoxuatthuocchokhoatheodoituong::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	//m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	//m_wndToDate.SetCheckValue(true);
	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);
	m_wndType.SetCheckValue(true);
	m_wndType.LimitText(35);


	m_wndStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStock.InsertColumn(1, _T("Description"), CFMT_TEXT, 300);


	m_wndType.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndType.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);


	m_wndOriginal.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndOriginal.InsertColumn(1, _T("Description"), CFMT_TEXT, 120);


	m_wndListFac.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndListFac.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndListFac.SetCheckBox(true);
	

}
void CrptBaocaoxuatthuocchokhoatheodoituong::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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
	m_wndOriginal.SetEvent(WE_SELENDOK, _OnOriginalSelendokFnc);
	//m_wndOriginal.SetEvent(WE_SETFOCUS, _OnOriginalSetfocusFnc);
	//m_wndOriginal.SetEvent(WE_KILLFOCUS, _OnOriginalKillfocusFnc);
	m_wndOriginal.SetEvent(WE_SELCHANGE, _OnOriginalSelectChangeFnc);
	m_wndOriginal.SetEvent(WE_LOADDATA, _OnOriginalLoadDataFnc);
	//m_wndOriginal.SetEvent(WE_ADDNEW, _OnOriginalAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	m_wndListFac.SetEvent(WE_SELCHANGE, _OnListFacSelectChangeFnc);
	m_wndListFac.SetEvent(WE_LOADDATA, _OnListFacLoadDataFnc);
	m_wndListFac.SetEvent(WE_DBLCLICK, _OnListFacDblClickFnc);	
	m_wndListFac.AddEvent(1, _T("Check All"), _OnCheckAllFnc, 0, VK_DELETE, 0);
	m_wndListFac.AddEvent(2, _T("UnCheck All"), _OnUnCheckAllFnc, 0, VK_DELETE, 0);

	m_szFromDate =  pMF->GetSysDate()+ _T("00:00");	
	m_szToDate =pMF->GetSysDate() + _T("23:59");
	OnListFacLoadData();
	SetMode(VM_EDIT);

}
void CrptBaocaoxuatthuocchokhoatheodoituong::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_TextEx(pDX, m_wndOriginal.GetDlgCtrlID(), m_szOriginalKey);
	DDX_Check(pDX, m_wndCheck.GetDlgCtrlID(), m_bCheck);

}
void CrptBaocaoxuatthuocchokhoatheodoituong::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CrptBaocaoxuatthuocchokhoatheodoituong::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CrptBaocaoxuatthuocchokhoatheodoituong::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_szTypeKey.Empty();
	m_szOriginalKey.Empty();

}
int CrptBaocaoxuatthuocchokhoatheodoituong::SetMode(int nMode){
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
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnFromDateChange(){
	
} */
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnFromDateSetfocus(){
	
} */
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnFromDateKillfocus(){
	
} */
int CrptBaocaoxuatthuocchokhoatheodoituong::OnFromDateCheckValue(){
	return 0;
} 
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnToDateChange(){
	
} */
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnToDateSetfocus(){
	
} */
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnToDateKillfocus(){
	
} */
int CrptBaocaoxuatthuocchokhoatheodoituong::OnToDateCheckValue(){
	return 0;
} 
void CrptBaocaoxuatthuocchokhoatheodoituong::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CrptBaocaoxuatthuocchokhoatheodoituong::OnStockSelendok(){
	 
}
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnStockSetfocus(){
	
}*/
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnStockKillfocus(){
	
}*/
long CrptBaocaoxuatthuocchokhoatheodoituong::OnStockLoadData(){
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
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CrptBaocaoxuatthuocchokhoatheodoituong::OnTypeSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CrptBaocaoxuatthuocchokhoatheodoituong::OnTypeSelendok(){
	 
}
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnTypeSetfocus(){
	
}*/
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnTypeKillfocus(){
	
}*/
long CrptBaocaoxuatthuocchokhoatheodoituong::OnTypeLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndType.IsSearchKey() && !m_szTypeKey.IsEmpty()){
		szWhere.Format(_T(" and ss_code='%s' "), m_szTypeKey);
	};
	m_wndType.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("select ss_code as id, ss_desc as name from sys_sel where ss_id='pms_export_type' and ss_code in('B','C')  %s"), szWhere);

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndType.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnTypeAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CrptBaocaoxuatthuocchokhoatheodoituong::OnOriginalSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CrptBaocaoxuatthuocchokhoatheodoituong::OnOriginalSelendok(){
	 
}
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnOriginalSetfocus(){
	
}*/
/*void CrptBaocaoxuatthuocchokhoatheodoituong::OnOriginalKillfocus(){
	
}*/
long CrptBaocaoxuatthuocchokhoatheodoituong::OnOriginalLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndOriginal.IsSearchKey() && !m_szOriginalKey.IsEmpty()){
		szWhere.Format(_T(" and mpr_product_resource_id=%d"), str2int(m_szOriginalKey));
	};
	szWhere.AppendFormat(_T(" AND mpr_org_id = '%s'"), pMF->GetModuleID());
	m_wndOriginal.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT * FROM m_product_resource WHERE mpr_isactive='Y' %s ORDER BY mpr_product_resource_id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndOriginal.AddItems(
			rs.GetValue(_T("mpr_product_resource_id")), 
			rs.GetValue(_T("mpr_name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*
void CrptBaocaoxuatthuocchokhoatheodoituong::OnOriginalAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
*/
void CrptBaocaoxuatthuocchokhoatheodoituong::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//UpdateData(true);
	OnPrintBaocaoxuathuocchokhoatheodoituong();
	//UpdateData(false);
} 
void CrptBaocaoxuatthuocchokhoatheodoituong::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//UpdateData(true);
	OnExportBaocaoxuathuocchokhoatheodoituong();
	//UpdateData(false);
} 
void CrptBaocaoxuatthuocchokhoatheodoituong::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//UpdateData(true);
	OnExportBaocaoxuathuocchokhoatheodoituong();
	//UpdateData(false);
} 
void CrptBaocaoxuatthuocchokhoatheodoituong::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

} 

int CrptBaocaoxuatthuocchokhoatheodoituong::OnCheckAll(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	for (int i=0; i< m_wndListFac.GetItemCount(); i++){
		m_wndListFac.SetCheck(i, true);
	}
	 return 0;
} 
int CrptBaocaoxuatthuocchokhoatheodoituong::OnUnCheckAll(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	for (int i=0; i< m_wndListFac.GetItemCount(); i++){
		m_wndListFac.SetCheck(i, false);
	}
	 return 0;
} 

void CrptBaocaoxuatthuocchokhoatheodoituong::OnListFacDblClick(){
	
} 
void CrptBaocaoxuatthuocchokhoatheodoituong::OnListFacSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CrptBaocaoxuatthuocchokhoatheodoituong::OnListFacDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
} 

long CrptBaocaoxuatthuocchokhoatheodoituong::OnListFacLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndListFac.BeginLoad(); 
	m_wndListFac.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" select sd_id as id, sd_name as name from sys_dept order by  sd_index,sd_type "));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndListFac.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndListFac.EndLoad(); 
	return nCount;
}
int CrptBaocaoxuatthuocchokhoatheodoituong::OnAddrptBaocaoxuatthuocchokhoatheodoituong(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CrptBaocaoxuatthuocchokhoatheodoituong::OnEditrptBaocaoxuatthuocchokhoatheodoituong(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CrptBaocaoxuatthuocchokhoatheodoituong::OnDeleterptBaocaoxuatthuocchokhoatheodoituong(){
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
 		OnCancelrptBaocaoxuatthuocchokhoatheodoituong(); 
 	}
	return 0;
}
int CrptBaocaoxuatthuocchokhoatheodoituong::OnSaverptBaocaoxuatthuocchokhoatheodoituong(){
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
 		//OnrptBaocaoxuatthuocchokhoatheodoituongListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CrptBaocaoxuatthuocchokhoatheodoituong::OnCancelrptBaocaoxuatthuocchokhoatheodoituong(){
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
int CrptBaocaoxuatthuocchokhoatheodoituong::OnrptBaocaoxuatthuocchokhoatheodoituongListLoadData(){
	return 0;
}

void CrptBaocaoxuatthuocchokhoatheodoituong::OnPrintBaocaoxuathuocchokhoatheodoituong(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL,tmpStr2,szObjects,m_szFacList, szWhere, szDeptID, szOrderby;
	CString  m_szServiceFlag,m_szInsFlag,m_szChildFlag,m_szFreeFlag,m_szGroups, szCaseDept, szSumDept, szSumTotal, szTotalAmount,szSumQty;
	CString szType, szTypeName;
	
	if(!rpt.Init(_T("Reports/HMS/PM_BAOCAOTINHHINHSUDUNGTAIDONVI.RPT"),true) )
	{
		return;
	}
	szWhere.Empty();
	
	if (!m_szStockKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and expstockid =%d "),ToLong(m_szStockKey));
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
	{
		if (m_szTypeKey == _T("B"))
			szWhere.AppendFormat(_T(" and iotype = 'DMO'"));
		else if (m_szTypeKey == _T("C"))
			szWhere.AppendFormat(_T(" and iotype = 'CSO'"));
	}	
	
	if(!m_szOriginalKey.IsEmpty())
		szWhere.AppendFormat(_T(" and product_resourceid=%d"), str2int(m_szOriginalKey));
	//szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	
	for (int i=0; i< m_wndListFac.GetItemCount(); i++)
	{
		if(m_wndListFac.GetCheck(i))
		{
			szDeptID.Format(_T("\"ID%s\""),m_wndListFac.GetItemText(i, 0));

			if(!m_szGroups.IsEmpty())
				m_szGroups += _T(",");						
			m_szGroups.AppendFormat(_T("'%s'"), m_wndListFac.GetItemText(i, 0));
			if (!szCaseDept.IsEmpty())
				szCaseDept +=_T(",");
			szCaseDept.AppendFormat(_T(" case when deptid='%s' then expqty else 0 end as %s "),m_wndListFac.GetItemText(i, 0),szDeptID);				
					
			if(!szSumDept.IsEmpty())
				szSumDept += _T(",");						
			szSumDept.AppendFormat(_T(" sum(%s) as %s"), szDeptID,szDeptID);

			if(!szSumTotal.IsEmpty())
				szSumTotal += _T("+");						
				szSumTotal.AppendFormat(_T("%s"), szDeptID);					

			if (!m_szFacList.IsEmpty())
				m_szFacList+= _T(", ");
			m_szFacList.AppendFormat(_T("%s"), m_wndListFac.GetItemText(i,1));
			}			
	}
	if (m_szFacList.IsEmpty())
	{
		AfxMessageBox(_T("Choose a department!"), MB_ICONASTERISK);
		return;
	}
	if (!szSumTotal.IsEmpty())
	{
		szTotalAmount.Format(_T(" sum(%s) * price as totalAmount "), szSumTotal);
		szSumQty.Format(_T(" sum(%s) as expqty "), szSumTotal);
	}
	if (m_bCheck)
		szOrderby.Format(_T("ORDER BY idx,groupname,name"));
	else
		szOrderby.Format(_T("ORDER BY name"));
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
		CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	if (!m_szStockKey.IsEmpty())
		rpt.GetReportHeader()->SetValue(_T("StockName"), m_wndStock.GetCurrent(1));
	else
		rpt.GetReportHeader()->SetValue(_T("StockName"), _T("To\xE0n \x62\x1ED9"));
	if (!m_szFacList.IsEmpty())
		if (m_szFacList.Find(',') > 0)
			rpt.GetReportHeader()->SetValue(_T("Department"), _T("Nhi\x1EC1u kho\x61"));
		else
			rpt.GetReportHeader()->SetValue(_T("Department"), m_szFacList);
	else
		rpt.GetReportHeader()->SetValue(_T("Department"), _T("To\xE0n \x62\x1ED9"));

	//rpt.GetReportHeader()->SetValue(_T("Department"), m_szFacList);
	rpt.GetReportHeader()->SetValue(_T("InsFlag"),m_szInsFlag);
	rpt.GetReportHeader()->SetValue(_T("ChildFlag"),m_szChildFlag);
	rpt.GetReportHeader()->SetValue(_T("FreeFlag"),m_szFreeFlag);
	rpt.GetReportHeader()->SetValue(_T("ServiceFlag"),m_szServiceFlag);
	//Page Header
	//Report Detail
		

	szSQL.Format(_T(" SELECT idx,groupname,id, name,unit,price, lot, expdte, %s,%s, %s FROM (") \
	_T(" 	SELECT product_categoryid as idx, product_categoryname as groupname,product_code as id, product_name as name,product_uomname as unit,") \
	_T("	product_saleprice1 as price, product_lotno as lot, product_expdate as expdte,") \
	_T("	%s") \
	_T(" 	FROM mev ") \
	_T(" 	LEFT JOIN m_productitem_view ON(product_item_id=sitemid)") \
	_T(" 	WHERE deptid in(%s) and expdate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s ") \
	_T(" 	) tbl") \
	_T(" GROUP BY idx,groupname,id, name,unit,price, lot, expdte") \
	_T(" %s"),szSumDept,szSumQty, szTotalAmount, szCaseDept, m_szGroups, m_szFromDate, m_szToDate,szWhere, szOrderby);
	CReportSection* rptDetail;
	CString szOldLine, szNewLine, szAmount;
	CRecord rs(&pMF->m_db);
	long double grpCost=0.0;
	long double ttlCost=0.0;
	double cost = 0;
	int nItem = 1;
	//_msg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);

	if(rs.IsEOF()){
		ShowMessage(150, MB_ICONSTOP);
		return ;
	}

	while(!rs.IsEOF())
	{
		if (m_bCheck)
		{
			rs.GetValue(_T("idx"), szNewLine);
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
				tmpStr.Format(_T("%s. %s"), szNewLine, rs.GetValue(_T("groupname")));
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
		//rs.GetValue(_T("name"), tmpStr);
		//rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("lot"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("expdte"), tmpStr);
		rptDetail->SetValue(_T("5"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rs.GetValue(_T("expqty"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rs.GetValue(_T("price"), cost);
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("7"), double2str(cost));
		rs.GetValue(_T("totalAmount"), cost);	
		grpCost += cost;
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("8"), double2str(cost));
		rs.MoveNext();
	}
						
	ttlCost += grpCost;
	if (m_bCheck)
	{
		if(grpCost > 0){	
			TranslateString(_T("Total Group"), szAmount);
			rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));	
			rptDetail->SetValue(_T("TotalGroup"), szAmount);
			FormatCurrency(grpCost, tmpStr);
			rptDetail->SetValue(_T("8"), tmpStr);			
		}
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
		
	rpt.PrintPreview();
			
}

void CrptBaocaoxuatthuocchokhoatheodoituong::OnExportBaocaoxuathuocchokhoatheodoituong(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	UpdateData(true);
	CString tmpStr, szSQL, m_szFacList, szWhere, szDeptID, szTemp;
	CString  m_szGroups, szCaseDept, szSumDept, szSumTotal, szTotalAmount,szSumQty;	
	CString szType, szTypeName;		
	int m_nCountList=0;
	if (!m_szStockKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and expstockid =%d "),ToLong(m_szStockKey));
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
	{
		if (m_szTypeKey == _T("B"))
			szWhere.AppendFormat(_T(" and iotype = 'DMO'"));
		else if (m_szTypeKey == _T("C"))
			szWhere.AppendFormat(_T(" and iotype = 'CSO'"));
	}

	
	if(!m_szOriginalKey.IsEmpty())
		szWhere.AppendFormat(_T(" and product_resourceid=%d"), str2int(m_szOriginalKey));
	//szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());

	for (int i=0; i< m_wndListFac.GetItemCount(); i++){
		if(m_wndListFac.GetCheck(i))
		{
			szDeptID.Format(_T("\"%s\""),m_wndListFac.GetItemText(i, 0));
			if(!m_szGroups.IsEmpty())
				m_szGroups += _T(",");						
			m_szGroups.AppendFormat(_T("'%s'"), m_wndListFac.GetItemText(i, 0));
			if (!szCaseDept.IsEmpty())
				szCaseDept +=_T(",");
			szCaseDept.AppendFormat(_T(" case when deptid='%s' then expqty else 0 end as %s "),m_wndListFac.GetItemText(i, 0),szDeptID);				
					
			if(!szSumDept.IsEmpty())
				szSumDept += _T(",");						
			szSumDept.AppendFormat(_T(" sum(%s) as %s"), szDeptID,szDeptID);

			if(!szSumTotal.IsEmpty())
				szSumTotal += _T("+");						
			szSumTotal.AppendFormat(_T("%s"), szDeptID);					

			if (!m_szFacList.IsEmpty())
				m_szFacList+= _T(", ");
			m_szFacList.AppendFormat(_T("%s"), m_wndListFac.GetItemText(i,1));

			m_nCountList++;
		}			
	}

	szTotalAmount.Format(_T(" sum(%s) * price as totalAmount "), szSumTotal);
	szSumQty.Format(_T(" sum(%s) as expqty "), szSumTotal);
	
	//Page Header
	//Report Detail
			
	szSQL.Format(_T(" SELECT idx,groupname,id, name,unit,price, %s,%s, %s FROM (") \
	_T(" 	SELECT product_categoryid as idx, product_categoryname as groupname,product_code as id, product_name as name,product_uomname as unit,") \
	_T("	product_saleprice1 as price, ") \
	_T("	%s") \
	_T(" 	FROM mev ") \
	_T(" 	LEFT JOIN m_productitem_view ON(product_item_id=sitemid)") \
	_T(" 	WHERE deptid in(%s) and expdate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s ") \
	_T(" 	) tbl") \
	_T(" GROUP BY idx,groupname,id, name,unit,price") \
	_T(" ORDER BY idx,groupname,name"),szSumDept,szSumQty, szTotalAmount, szCaseDept, m_szGroups, m_szFromDate, m_szToDate,szWhere );
	CString szOldLine, szNewLine;
	CString szOldGroupName, szNewGroupName;
	CRecord rs(&pMF->m_db);
	long double grpCost=0.0, ttlCost = 0.0;
	double cost = 0;
	long qty=0;
	int nItem = 1;

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
	xls.SetColumnWidth(2, 30);
	xls.SetColumnWidth(3, 8);
	xls.SetColumnWidth(6, 10);
	xls.SetColumnWidth(7, 10);
	xls.SetColumnWidth(8, 10);
	xls.SetColumnWidth(9, 10);
	xls.SetColumnWidth(10, 20);
	xls.SetCellMergedColumns(0, 0, 3);
	xls.SetCellMergedColumns(0, 1, 3);
	xls.SetCellText(0, 0, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(0, 2, 11);
	TranslateString(_T("Department Drug Export Report"), szTemp);
	StringUpper(szTemp, tmpStr);
	xls.SetCellText(0, 2, tmpStr,FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(0,3,11);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(0, 3, tmpStr, &df);
	int nRow = 3;
	if (!m_szStockKey.IsEmpty())
	{
		nRow++;
		TranslateString(_T("From Stock"), tmpStr);
		xls.SetCellMergedColumns(0, nRow, 11);
		xls.SetCellText(0, nRow, tmpStr + _T(":") + m_wndStock.GetCurrent(1), &cf);
	}
	if (!m_szFacList.IsEmpty())
	{
		nRow++;
		TranslateString(_T("Dept"), tmpStr);
		xls.SetCellMergedColumns(0, nRow, 11);
		xls.SetCellText(0, nRow, tmpStr + _T(":") +m_szFacList, &cf);
	}
	if (!m_szTypeKey.IsEmpty())
	{
		nRow++;
		TranslateString(_T("Type"), tmpStr);
		xls.SetCellMergedColumns(0, nRow, 11);
		xls.SetCellText(0, nRow, tmpStr + _T(":")+ pMF->GetSelectionString(_T("pms_export_type"), m_szTypeKey), &cf);
	}
	nRow++;
	xls.SetCellText(0, nRow, _T("STT"), &hf);
	TranslateString(_T("ID"), tmpStr);
	xls.SetCellText(1, nRow, tmpStr, &hf);
	TranslateString(_T("Drug Name"), tmpStr);
	xls.SetCellText(2, nRow, tmpStr, &hf);
	TranslateString(_T("Unit"), tmpStr);
	xls.SetCellText(3, nRow, tmpStr, &hf);
	TranslateString(_T("Country"), tmpStr);
	xls.SetCellText(4, nRow, tmpStr + _T("SX"), &hf);
	TranslateString(_T("Price"), tmpStr);
	xls.SetCellText(5, nRow, tmpStr,  &hf);
	int j=6;
	CString tmpStr1;
	for (i=0;i< m_wndListFac.GetItemCount(); i++)
	{
		if(m_wndListFac.GetCheck(i))
		{				
			tmpStr1.Format(_T("%s"), m_wndListFac.GetItemText(i,0));			
			xls.SetCellText(j,nRow, tmpStr1,&hf);
			j++;
		}    
	}
	TranslateString(_T("Total Export"), tmpStr);
	xls.SetCellText(5+m_nCountList+1, nRow, tmpStr,&hf);
	TranslateString(_T("Total Amount"), tmpStr);
	xls.SetCellText(5+ m_nCountList+2, nRow, tmpStr,&hf);	
	
	nRow++;
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	while(!rs.IsEOF())
	{
		if (m_bCheck)
		{
			rs.GetValue(_T("idx"), szNewLine);
			if(szNewLine != szOldLine && !szNewLine.IsEmpty()){	
				if(grpCost > 0)
				{			
					CString szField, szAmount;
					TranslateString(_T("\x43\x1ED9ng:"), szAmount);
					xls.SetCellMergedColumns(1,nRow,5+m_nCountList);
					xls.SetCellText(1, nRow, szAmount, FMT_TEXT,true,11,0);
					tmpStr.Format(_T("%.2f"), grpCost);
					xls.SetCellText(5+m_nCountList+2, nRow, tmpStr, FMT_NUMBER1,true,11,0);
					nRow++;
					ttlCost += grpCost;
					grpCost = 0;
					nItem=1;
				}			
				xls.SetCellMergedColumns(1,nRow,5+m_nCountList +2);	
				tmpStr.Format(_T("%s) %s"), szNewLine, rs.GetValue(_T("groupname")));
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT,true,10,0);
				nRow++;
				szOldLine = szNewLine;
				nItem=1;
			}
		}
	
		tmpStr.Format(_T("%d"), nItem++);
		xls.SetCellText(0,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("id"), tmpStr);
		xls.SetCellText(1,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("name"), tmpStr);
		xls.SetCellText(2,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("unit"), tmpStr);		
		xls.SetCellText(3,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("Countryid"), tmpStr);
		xls.SetCellText(4,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("price"), tmpStr);
		xls.SetCellText(5,nRow,tmpStr, FMT_NUMBER1);		
		j=6;
		for (int i=0;i< m_wndListFac.GetItemCount(); i++)
		{
			if(m_wndListFac.GetCheck(i))
			{					
				tmpStr1.Format(_T("%s"), m_wndListFac.GetItemText(i,0));
				rs.GetValue(tmpStr1, tmpStr);
				xls.SetCellText(j,nRow, tmpStr,FMT_NUMBER1);
				j++;
			}    
		}
		
		rs.GetValue(_T("expqty"), qty);
		tmpStr.Format(_T("%ld"), qty);
		xls.SetCellText(5+ m_nCountList+1,nRow,tmpStr, FMT_NUMBER1);				
		rs.GetValue(_T("totalAmount"), cost);				
		tmpStr.Format(_T("%.2f"), cost);
		grpCost+=cost;
		xls.SetCellText(5+ m_nCountList+2,nRow,tmpStr, FMT_NUMBER1);		
		nRow++;
		rs.MoveNext();
	}
	ttlCost += grpCost;
	if (m_bCheck)
	{
		if(grpCost > 0){
			TranslateString(_T("\x43\x1ED9ng:"), tmpStr);		
			xls.SetCellMergedColumns(1,nRow,5+m_nCountList+1);
			xls.SetCellText(1, nRow, tmpStr, FMT_TEXT,true,11,0);
			tmpStr.Format(_T("%.2f"), grpCost);
			xls.SetCellText(5+m_nCountList +2, nRow, tmpStr, FMT_NUMBER1,true,11,0);		
		}
	}
	if(ttlCost > 0){
		TranslateString(_T("T\x1ED5ng ti\x1EC1n:"), tmpStr);		
		xls.SetCellMergedColumns(1,nRow,5+m_nCountList+1);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT,true,11,0);
		tmpStr.Format(_T("%.2f"), ttlCost);
		xls.SetCellText(5+m_nCountList +2, nRow, tmpStr, FMT_NUMBER1,true,11,0);		
	}
	xls.Save(_T("Exports\\Baocaoxuatthuocchokhoa.xls"));
}
BEGIN_MESSAGE_MAP(CrptBaocaoxuatthuocchokhoatheodoituong, CGuiView)
	ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CrptBaocaoxuatthuocchokhoatheodoituong::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
