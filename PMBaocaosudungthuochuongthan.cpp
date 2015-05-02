#include "stdafx.h"
#include "PMBaocaosudungthuochuongthan.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CrptBaocaosudungthuochuongthan *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CrptBaocaosudungthuochuongthan *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CrptBaocaosudungthuochuongthan *)pWnd)->OnFromDateKillfocus();
} */
static int _OnListSelectAllFnc(CWnd *pWnd){
	 return ((CrptBaocaosudungthuochuongthan*)pWnd)->OnListSelectAll();
} 

static int _OnListUnselectAllFnc(CWnd *pWnd){
	 return ((CrptBaocaosudungthuochuongthan*)pWnd)->OnListUnselectAll();
}
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CrptBaocaosudungthuochuongthan *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CrptBaocaosudungthuochuongthan *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CrptBaocaosudungthuochuongthan *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CrptBaocaosudungthuochuongthan *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CrptBaocaosudungthuochuongthan *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CrptBaocaosudungthuochuongthan* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CrptBaocaosudungthuochuongthan *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CrptBaocaosudungthuochuongthan *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CrptBaocaosudungthuochuongthan *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CrptBaocaosudungthuochuongthan *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CrptBaocaosudungthuochuongthan *)pWnd)->OnStockAddNew();
}*/
static void _OnInsuranceSelectFnc(CWnd *pWnd){
	 ((CrptBaocaosudungthuochuongthan*)pWnd)->OnInsuranceSelect();
}
static void _OnDiscountSelectFnc(CWnd *pWnd){
	 ((CrptBaocaosudungthuochuongthan*)pWnd)->OnDiscountSelect();
}
static void _OnServiceSelectFnc(CWnd *pWnd){
	 ((CrptBaocaosudungthuochuongthan*)pWnd)->OnServiceSelect();
}
static void _OnOtherSelectFnc(CWnd *pWnd){
	 ((CrptBaocaosudungthuochuongthan*)pWnd)->OnOtherSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CrptBaocaosudungthuochuongthan *pVw = (CrptBaocaosudungthuochuongthan *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CrptBaocaosudungthuochuongthan *pVw = (CrptBaocaosudungthuochuongthan *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CrptBaocaosudungthuochuongthan *pVw = (CrptBaocaosudungthuochuongthan *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CrptBaocaosudungthuochuongthan *pVw = (CrptBaocaosudungthuochuongthan *)pWnd;
	pVw->OnCloseSelect();
} 
static long _OnListFacLoadDataFnc(CWnd *pWnd){
	return ((CrptBaocaosudungthuochuongthan*)pWnd)->OnListFacLoadData();
} 
static void _OnListFacDblClickFnc(CWnd *pWnd){
	((CrptBaocaosudungthuochuongthan*)pWnd)->OnListFacDblClick();
} 
static void _OnListFacSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CrptBaocaosudungthuochuongthan*)pWnd)->OnListFacSelectChange(nOldItem, nNewItem);
} 
static int _OnListFacDeleteItemFnc(CWnd *pWnd){
	 return ((CrptBaocaosudungthuochuongthan*)pWnd)->OnListFacDeleteItem();
} 
static int _OnAddrptBaocaosudungthuochuongthanFnc(CWnd *pWnd){
	 return ((CrptBaocaosudungthuochuongthan*)pWnd)->OnAddrptBaocaosudungthuochuongthan();
} 
static int _OnEditrptBaocaosudungthuochuongthanFnc(CWnd *pWnd){
	 return ((CrptBaocaosudungthuochuongthan*)pWnd)->OnEditrptBaocaosudungthuochuongthan();
} 
static int _OnDeleterptBaocaosudungthuochuongthanFnc(CWnd *pWnd){
	 return ((CrptBaocaosudungthuochuongthan*)pWnd)->OnDeleterptBaocaosudungthuochuongthan();
} 
static int _OnSaverptBaocaosudungthuochuongthanFnc(CWnd *pWnd){
	 return ((CrptBaocaosudungthuochuongthan*)pWnd)->OnSaverptBaocaosudungthuochuongthan();
} 
static int _OnCancelrptBaocaosudungthuochuongthanFnc(CWnd *pWnd){
	 return ((CrptBaocaosudungthuochuongthan*)pWnd)->OnCancelrptBaocaosudungthuochuongthan();
} 
CrptBaocaosudungthuochuongthan::CrptBaocaosudungthuochuongthan(CWnd *pParent)
{
	m_nDlgWidth = 455;
	m_nDlgHeight = 505;
	m_szTitle = _T("Danh sach benh nhan su dung thuoc gay nghien - huong than");
	SetDefaultValues();
}
CrptBaocaosudungthuochuongthan::~CrptBaocaosudungthuochuongthan(){
}
void CrptBaocaosudungthuochuongthan::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 447, 465);
	m_wndObject.Create(this, _T("Object"), 13, 91, 440, 153);
	m_wndFromDateLabel.Create(this, _T("From Date"), 13, 30, 115, 55);
	m_wndFromDate.Create(this,120, 30, 235, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 240, 30, 320, 55);
	m_wndToDate.Create(this,325, 30, 440, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 13, 60, 115, 85);
	m_wndStock.Create(this,120, 60, 440, 85); 
	m_wndInsurance.Create(this, _T("Insurance"), 27, 119, 123, 144);
	m_wndDiscount.Create(this, _T("Discount"), 128, 119, 224, 144);
	m_wndService.Create(this, _T("Service"), 229, 119, 325, 144);
	m_wndOther.Create(this, _T("Other"), 330, 119, 426, 144);
	//m_wndPrintPreview.Create(this, _T("Print Pre&view"), 110, 470, 190, 495);
	//m_wndPrint.Create(this, _T("&Print"), 195, 470, 275, 495);
	//m_wndExport.Create(this, _T("&Export"), 280, 470, 360, 495);
	m_wndExport.Create(this, _T("&Export"), 365, 470, 445, 495);
	m_wndListFac.Create(this,13, 158, 438, 458); 

}
void CrptBaocaosudungthuochuongthan::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(FALSE);
	//EnableButtons(TRUE, 0, -1);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);


	m_wndStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStock.InsertColumn(1, _T("Description"), CFMT_TEXT, 220);

	m_wndListFac.SetCheckBox(true);
	m_wndListFac.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndListFac.InsertColumn(1, _T("Drug Name"), CFMT_TEXT, 250);
	m_wndListFac.InsertColumn(2, _T("Unit"), CFMT_TEXT, 60);

	m_szFromDate=pMF->GetSysDate() + _T("00:00");
	m_szToDate = pMF->GetSysDateTime();
	OnListFacLoadData();
	OnListSelectAll();
	SetMode(VM_EDIT);


}
void CrptBaocaosudungthuochuongthan::OnSetWindowEvents(){
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
	m_wndInsurance.SetEvent(WE_CLICK, _OnInsuranceSelectFnc);
	m_wndDiscount.SetEvent(WE_CLICK, _OnDiscountSelectFnc);
	m_wndService.SetEvent(WE_CLICK, _OnServiceSelectFnc);
	m_wndOther.SetEvent(WE_CLICK, _OnOtherSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	m_wndListFac.SetEvent(WE_SELCHANGE, _OnListFacSelectChangeFnc);
	m_wndListFac.SetEvent(WE_LOADDATA, _OnListFacLoadDataFnc);
	m_wndListFac.SetEvent(WE_DBLCLICK, _OnListFacDblClickFnc);

	m_wndListFac.AddEvent(1, _T("Select All"), _OnListSelectAllFnc);
	m_wndListFac.AddEvent(2, _T("Unselect All"), _OnListUnselectAllFnc);

}
void CrptBaocaosudungthuochuongthan::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_Check(pDX, m_wndInsurance.GetDlgCtrlID(), m_bInsurance);
	DDX_Check(pDX, m_wndDiscount.GetDlgCtrlID(), m_bDiscount);
	DDX_Check(pDX, m_wndService.GetDlgCtrlID(), m_bService);
	DDX_Check(pDX, m_wndOther.GetDlgCtrlID(), m_bOther);

}
void CrptBaocaosudungthuochuongthan::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CrptBaocaosudungthuochuongthan::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CrptBaocaosudungthuochuongthan::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_bInsurance=TRUE;
	m_bDiscount=TRUE;
	m_bService=TRUE;
	m_bOther=TRUE;

}
int CrptBaocaosudungthuochuongthan::SetMode(int nMode){
 		int nOldMode = GetMode(); 
 		CGuiView::SetMode(nMode); 
 		CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
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
 			EnableButtons(TRUE,0,1, 2, 3, 4,5, -1); 
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
/*void CrptBaocaosudungthuochuongthan::OnFromDateChange(){
	
} */
/*void CrptBaocaosudungthuochuongthan::OnFromDateSetfocus(){
	
} */
/*void CrptBaocaosudungthuochuongthan::OnFromDateKillfocus(){
	
} */
int CrptBaocaosudungthuochuongthan::OnFromDateCheckValue(){
	return 0;
} 
/*void CrptBaocaosudungthuochuongthan::OnToDateChange(){
	
} */
/*void CrptBaocaosudungthuochuongthan::OnToDateSetfocus(){
	
} */
/*void CrptBaocaosudungthuochuongthan::OnToDateKillfocus(){
	
} */
int CrptBaocaosudungthuochuongthan::OnToDateCheckValue(){
	return 0;
} 
void CrptBaocaosudungthuochuongthan::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CrptBaocaosudungthuochuongthan::OnStockSelendok(){
	 
}
/*void CrptBaocaosudungthuochuongthan::OnStockSetfocus(){
	
}*/
/*void CrptBaocaosudungthuochuongthan::OnStockKillfocus(){
	
}*/
long CrptBaocaosudungthuochuongthan::OnStockLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStock.IsSearchKey() && ToInt(m_szStockKey) > 0){
		szWhere.Format(_T(" AND pmsl_stockid=%d"), ToInt(m_szStockKey));
	};
	szWhere.AppendFormat(_T(" and pmsl_stockid in(%s) "), pMF->m_szStocks);
	m_wndStock.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT pmsl_stockid as id, pmsl_name as name FROM pms_stocklist where 1=1 %s ORDER BY pmsl_stockid"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStock.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CrptBaocaosudungthuochuongthan::OnStockAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
	void CrptBaocaosudungthuochuongthan::OnInsuranceSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
	void CrptBaocaosudungthuochuongthan::OnDiscountSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
	void CrptBaocaosudungthuochuongthan::OnServiceSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
	void CrptBaocaosudungthuochuongthan::OnOtherSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CrptBaocaosudungthuochuongthan::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	OnPrintBaocaodanhsachbenhnhansudungthuochuongthan(true);
	UpdateData(false);
} 
void CrptBaocaosudungthuochuongthan::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	OnPrintBaocaodanhsachbenhnhansudungthuochuongthan(true);
	UpdateData(false);
} 
void CrptBaocaosudungthuochuongthan::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	OnExportBaocaodanhsachbenhnhansudungthuochuongthan();
	UpdateData(false);
} 
void CrptBaocaosudungthuochuongthan::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	
} 
void CrptBaocaosudungthuochuongthan::OnListFacDblClick(){
	
} 
void CrptBaocaosudungthuochuongthan::OnListFacSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CrptBaocaosudungthuochuongthan::OnListFacDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CrptBaocaosudungthuochuongthan::OnListFacLoadData(){

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndListFac.BeginLoad(); 
	m_wndListFac.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" SELECT pmi_id as id, pmi_name as name, pmi_unit as unit FROM pms_items WHERE pmi_typeid in('A1130','A1140') "));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndListFac.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), 
			rs.GetValue(_T("Unit")), NULL);
		rs.MoveNext();
	}
	m_wndListFac.EndLoad(); 
	return nCount;
}
int CrptBaocaosudungthuochuongthan::OnListSelectAll(){
	for (int i =0; i < m_wndListFac.GetItemCount(); i++){
		m_wndListFac.SetCheck(i, true);
	}
	return i;
}

int CrptBaocaosudungthuochuongthan::OnListUnselectAll(){
	for (int i =0; i < m_wndListFac.GetItemCount(); i++){
		m_wndListFac.SetCheck(i, false);
	}
	return i;
}
int CrptBaocaosudungthuochuongthan::OnAddrptBaocaosudungthuochuongthan(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CrptBaocaosudungthuochuongthan::OnEditrptBaocaosudungthuochuongthan(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CrptBaocaosudungthuochuongthan::OnDeleterptBaocaosudungthuochuongthan(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
 	CString szSQL; 
 	if(ShowMessage(1, MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2) == IDNO) 
 		return -1; 
 	szSQL.Format(_T("DELETE FROM  WHERE  AND") ); 
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret >= 0){ 
 		SetMode(VM_NONE); 
 		OnCancelrptBaocaosudungthuochuongthan(); 
 	}
	return 0;
}
int CrptBaocaosudungthuochuongthan::OnSaverptBaocaosudungthuochuongthan(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT) 
 		return -1; 
 	if(!IsValidateData()) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
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
 		//OnrptBaocaosudungthuochuongthanListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CrptBaocaosudungthuochuongthan::OnCancelrptBaocaosudungthuochuongthan(){
 	if(GetMode() == VM_EDIT) 
 	{ 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 		SetMode(VM_NONE); 
 	} 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
 	return 0;
} 
int CrptBaocaosudungthuochuongthan::OnrptBaocaosudungthuochuongthanListLoadData(){
	return 0;
}
void CrptBaocaosudungthuochuongthan::OnPrintBaocaodanhsachbenhnhansudungthuochuongthan(bool m_nPreview){
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL, szObjects, szWhere, tmpStr2;
	CString szCaseDrugID,m_szFacList,m_szGroups,szSumDept;
	if(!rpt.Init(_T("Reports/HMS/PMS_PRESCRIPTIONLISTREPORT.RPT"),true) )
		return;
	UpdateData(TRUE);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy),
		CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);


	rpt.GetReportHeader()->SetValue(_T("StockName"), m_wndStock.GetCurrent(1));
	tmpStr2 = _T("' '");
	if(m_bInsurance)
	{
		TranslateString(_T("Insurance"), tmpStr);
		szObjects.Format(_T("%s"), tmpStr);
		tmpStr2 += _T(",'I','C'");
	}
	if(m_bDiscount)
	{
		TranslateString(_T("Discount"), tmpStr);
		if(!szObjects.IsEmpty())
			szObjects += _T("; ");
		szObjects.AppendFormat(_T("%s"), tmpStr);
		tmpStr2 += _T(",'D','Z'");
	}

	if(m_bService)
	{
		TranslateString(_T("Service"), tmpStr);
		if(!szObjects.IsEmpty())
			szObjects += _T("; ");
		szObjects.AppendFormat(_T("%s"), tmpStr);
		tmpStr2 += _T(",'S','X'");
	}
	if(m_bOther)
	{
		TranslateString(_T("Other"), tmpStr);
		if(!szObjects.IsEmpty())
			szObjects += _T("; ");
		szObjects.AppendFormat(_T("%s"), tmpStr);
		tmpStr2 += _T(",'Y'");
	}
	szWhere.Format(_T(" AND ho_type in(%s) "), tmpStr2);
	tmpStr.Empty();
	
	for (int i=0; i< m_wndListFac.GetItemCount(); i++){
				if(m_wndListFac.GetCheck(i)){
					if(!m_szGroups.IsEmpty())
						m_szGroups += _T(",");						
					m_szGroups.AppendFormat(_T("'%s'"), m_wndListFac.GetItemText(i, 0));
				if (!szCaseDrugID.IsEmpty())
					szCaseDrugID +=_T(",");
					szCaseDrugID.AppendFormat(_T(" case when pmi_id='%s' then hpol_issueqty else 0 end as %s "),m_wndListFac.GetItemText(i, 0),m_wndListFac.GetItemText(i, 0));				

				if (!m_szFacList.IsEmpty())
					m_szFacList+= _T(", ");
					m_szFacList.AppendFormat(_T("%s"), m_wndListFac.GetItemText(i,1));

				if(!szSumDept.IsEmpty())
					szSumDept += _T(",");						
					szSumDept.AppendFormat(_T(" sum(%s) as %s"), m_wndListFac.GetItemText(i, 0),m_wndListFac.GetItemText(i, 0));
				}				
			
		}

	rpt.GetReportHeader()->SetValue(_T("Object"), szObjects);
	//Page Header
	//Report Detail
	szSQL.Format(_T("SELECT docno,pname,age,CardNo,DeptID, ") \
	_T(" %s") \
	_T(" FROM ( ") \
	_T("	SELECT distinct hd_docno as docno,") \
	_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
	_T(" 	hms_getage(date(hd_admitdate), hp_birthdate) as age, 	") \
	_T(" 	hd_cardno as cardNo,hpo_deptid as DeptId,	") \
	_T(" %s")	\
	_T(" FROM hms_patient") \
	_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
	_T(" LEFT JOIN hms_pharmacyorder ON(hpo_docno=hd_docno)") \
	_T(" LEFT JOIN hms_pharmacyorder_line ON(hpol_orderid=hpo_orderid) ") \
	_T(" LEFT JOIN pms_stockitems ON(pmsi_id=hpol_sitemid)") \
	_T(" LEFT JOIN pms_items ON(pmsi_itemid=pmi_id)") \
	_T(" LEFT JOIN hms_object ON(hd_object=ho_id)") \
	_T(" WHERE 	hpo_status IN('A','I')	and hpo_stockid=%d") \
	_T(" 	and (hpo_issuedate) between ('%s') and ('%s')") \
	_T(" 	and hpo_type in('P','D','M','E') and pmi_typeid ='A1140'") \
	_T(" ) as TblTemp ") \
	_T(" GROUP BY  docno,pname,age,CardNo,DeptID") \
	_T(" ORDER BY docno") ,szSumDept, szCaseDrugID,ToInt(m_szStockKey), m_szFromDate, m_szToDate);

	CReportSection* rptDetail;
	CReportSection* rptGroup;
	CRecord rs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	//_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);
	long nItem = 1, nTotalOrder=0;
	long nOrderID;
	double nTotalAmount=0,cost, nAmount;

	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nItem++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("pname"), tmpStr);		
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("age"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rs.GetValue(_T("cardno"), tmpStr);		
		rptDetail->SetValue(_T("6"), tmpStr);
		rs.MoveNext();
	}
	
	//Page Footer
	//Report Footer

	tmpStr.Format(_T("%ld"), nTotalOrder);
	rpt.GetReportFooter()->SetValue(_T("TotalOrder"), tmpStr);
	tmpStr.Format(_T("%.2f"), nTotalAmount);
	FormatCurrency(nTotalAmount, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	//	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
	if(m_bPreview)
		rpt.PrintPreview();
	else
		rpt.Print();

}
void CrptBaocaosudungthuochuongthan::OnExportBaocaodanhsachbenhnhansudungthuochuongthan(){
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();	
	int m_nCountList=0;
	CString tmpStr, szDate, szSysDate, szSQL, szObjects, szWhere, tmpStr2;
	CString szCaseDrugID,m_szFacList,m_szGroups,szSumDept;
	
	tmpStr2 = _T("' '");
	if(m_bInsurance)
	{
		TranslateString(_T("Insurance"), tmpStr);
		szObjects.Format(_T("%s"), tmpStr);
		tmpStr2 += _T(",'I','C'");
	}
	if(m_bDiscount)
	{
		TranslateString(_T("Discount"), tmpStr);
		if(!szObjects.IsEmpty())
			szObjects += _T("; ");
		szObjects.AppendFormat(_T("%s"), tmpStr);
		tmpStr2 += _T(",'D','Z'");
	}

	if(m_bService)
	{
		TranslateString(_T("Service"), tmpStr);
		if(!szObjects.IsEmpty())
			szObjects += _T("; ");
		szObjects.AppendFormat(_T("%s"), tmpStr);
		tmpStr2 += _T(",'S','X'");
	}
	if(m_bOther)
	{
		TranslateString(_T("Other"), tmpStr);
		if(!szObjects.IsEmpty())
			szObjects += _T("; ");
		szObjects.AppendFormat(_T("%s"), tmpStr);
		tmpStr2 += _T(",'Y'");
	}
	szWhere.Format(_T(" AND ho_type in(%s) "), tmpStr2);
	tmpStr.Empty();
	
	for (int i=0; i< m_wndListFac.GetItemCount(); i++){
		if(m_wndListFac.GetCheck(i)){
				if(!m_szGroups.IsEmpty())
						m_szGroups += _T(",");						
					m_szGroups.AppendFormat(_T("'%s'"), m_wndListFac.GetItemText(i, 0));
				if (!szCaseDrugID.IsEmpty())
					szCaseDrugID +=_T(",");
					szCaseDrugID.AppendFormat(_T(" case when pmi_id='%s' then hpol_issueqty else 0 end as %s "),m_wndListFac.GetItemText(i, 0),m_wndListFac.GetItemText(i, 0));				

				if (!m_szFacList.IsEmpty())
					m_szFacList+= _T(", ");
					m_szFacList.AppendFormat(_T("%s"), m_wndListFac.GetItemText(i,1));

				if(!szSumDept.IsEmpty())
					szSumDept += _T(",");						
					szSumDept.AppendFormat(_T(" sum(%s) as %s"), m_wndListFac.GetItemText(i, 0),m_wndListFac.GetItemText(i, 0));
			m_nCountList++;
		}				
			
	}

	
	//Page Header
	//Report Detail
	szSQL.Format(_T("SELECT docno,pname,age,CardNo,DeptID,icd,diagnostic, ") \
	_T(" %s") \
	_T(" FROM ( ") \
	_T("	SELECT distinct hd_docno as docno,") \
	_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
	_T(" 	hms_getage(date(hd_admitdate), hp_birthdate) as age, 	") \
	_T(" 	hd_cardno as cardNo,hpo_deptid as DeptId,hd_icd as icd, hd_diagnostic as diagnostic,	") \
	_T(" %s")	\
	_T(" FROM hms_patient") \
	_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
	_T(" LEFT JOIN hms_pharmacyorder ON(hpo_docno=hd_docno)") \
	_T(" LEFT JOIN hms_pharmacyorder_line ON(hpol_orderid=hpo_orderid) ") \
	_T(" LEFT JOIN pms_stockitems ON(pmsi_id=hpol_sitemid)") \
	_T(" LEFT JOIN pms_items ON(pmsi_itemid=pmi_id)") \
	_T(" LEFT JOIN hms_object ON(hd_object=ho_id)") \
	_T(" WHERE 	hpo_status IN('A','I')	and hpo_stockid=%d") \
	_T(" 	and (hpo_issuedate) between ('%s') and ('%s')") \
	_T(" 	and hpo_type in('P','D','M','E') and pmi_typeid in('A1130','A1140') %s ") \
	_T(" ) as TblTemp ") \
	_T(" GROUP BY  docno,pname,age,CardNo,DeptID, icd, diagnostic") \
	_T(" ORDER BY docno,DeptID") ,szSumDept, szCaseDrugID,ToInt(m_szStockKey), m_szFromDate, m_szToDate, szWhere);

	
	CString szOldLine, szNewLine;
	CRecord rs(&pMF->m_db);
	long double grpCost=0.0;
	long double ttlCost=0.0;
	double cost = 0;
	long qty=0;
	int nItem = 1;
	_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);

	if(rs.IsEOF()){
		ShowMessageBox(_T("No data"), MB_OK);
		return ;
	}

					
	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);	
	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname,FMT_TEXT);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name,FMT_TEXT);
	xls.SetCellMergedColumns(0,4,7+m_nCountList);	
	TranslateString(_T("DANH SACH BENH NHAN SU DUNG THUOC HUONG THAN"),tmpStr);
	xls.SetCellText(0, 4, tmpStr,FMT_TEXT,true,14,0);
	xls.SetCellMergedColumns(0,5,7+m_nCountList);
	tmpStr.Format(_T("Tu ngay %s Den ngay %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(0, 5, tmpStr,FMT_TEXT,true,12,0);	

	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
				CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	TranslateString(_T("From Stock"), tmpStr);
	xls.SetCellText(0, 7,tmpStr + _T(":") + m_wndStock.GetCurrent(1));
	TranslateString(_T("Ojbect"), tmpStr);
	xls.SetCellText(0, 8,tmpStr + _T(":") + szObjects);

	int nRow = 10;
	xls.SetCellText(0, nRow, _T("STT"),FMT_TEXT,true,11,0);
	TranslateString(_T("Document No"), tmpStr);
	xls.SetCellText(1, nRow, tmpStr,FMT_TEXT,true,11,0);
	TranslateString(_T("Patient Name"), tmpStr);
	xls.SetCellText(2, nRow, tmpStr,FMT_TEXT,true,11,0);
	TranslateString(_T("Age"), tmpStr);
	xls.SetCellText(3, nRow, tmpStr,FMT_TEXT,true,11,0);
	TranslateString(_T("Insurance"), tmpStr);
	xls.SetCellText(4, nRow, tmpStr ,FMT_TEXT,true,11,0);
	TranslateString(_T("Dept"), tmpStr);
	xls.SetCellText(5, nRow, tmpStr,FMT_TEXT,true,11,0);
	TranslateString(_T("ICD"), tmpStr);
	xls.SetCellText(6, nRow, tmpStr ,FMT_TEXT,true,11,0);
	TranslateString(_T("Diagnostic"), tmpStr);
	xls.SetCellText(7, nRow, tmpStr,FMT_TEXT,true,11,0);
	int j=8;
	CString tmpStr1;
	for (i=0;i< m_wndListFac.GetItemCount(); i++)
	{
		if(m_wndListFac.GetCheck(i))
		{				
			tmpStr1.Format(_T("%s"), m_wndListFac.GetItemText(i,1));			
			xls.SetCellText(j,nRow, tmpStr1,FMT_TEXT,true,11,0);
			j++;
		}    
	}	
	
			
	while(!rs.IsEOF())
	{
		nRow ++;		
		tmpStr.Format(_T("%d"), nItem++);
		xls.SetCellText(0,nRow,tmpStr, FMT_NUMBER1);		
		rs.GetValue(_T("Docno"), tmpStr);
		xls.SetCellText(1,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("Pname"), tmpStr);
		xls.SetCellText(2,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("Age"), tmpStr);		
		xls.SetCellText(3,nRow,tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("CardNo"), tmpStr);
		xls.SetCellText(4,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("DeptID"), tmpStr);
		xls.SetCellText(5,nRow,tmpStr, FMT_TEXT);		
		rs.GetValue(_T("ICD"), tmpStr);
		xls.SetCellText(6,nRow,tmpStr, FMT_TEXT);
		rs.GetValue(_T("Diagnostic"), tmpStr);
		xls.SetCellText(7,nRow,tmpStr, FMT_TEXT);		
		j=8;
		for (i=0;i< m_wndListFac.GetItemCount(); i++)
		{
			if(m_wndListFac.GetCheck(i))
			{					
				tmpStr1.Format(_T("%s"), m_wndListFac.GetItemText(i,0));
				rs.GetValue(tmpStr1, tmpStr);
				xls.SetCellText(j,nRow, tmpStr,FMT_NUMBER1);
				j++;
			}    
		}	
		
		rs.MoveNext();
	}
	
	xls.Save(_T("Exports\\Baocaoxuatthuocchokhoa.xls"));

}