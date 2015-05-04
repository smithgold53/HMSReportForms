#include "stdafx.h"
#include "PMStockCardReportDialog.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include ".\pmstockcardreportdialog.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSStockCardReportDialog *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSStockCardReportDialog *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSStockCardReportDialog* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CPMSStockCardReportDialog *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnStockAddNew();
}*/
/*static void _OnNameChangeFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnNameChange();
} */
/*static void _OnNameSetfocusFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnNameSetfocus();} */ 
/*static void _OnNameKillfocusFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog *)pWnd)->OnNameKillfocus();
} */
static int _OnNameCheckValueFnc(CWnd *pWnd){
	return ((CPMSStockCardReportDialog *)pWnd)->OnNameCheckValue();
} 


static int _OnListSelectAllFnc(CWnd *pWnd){
	return ((CPMSStockCardReportDialog*)pWnd)->OnListSelectAll();
} 
static int _OnListUnselectAllFnc(CWnd *pWnd){
	return ((CPMSStockCardReportDialog*)pWnd)->OnListUnselectAll();
} 
static void _OnByProductLineSelectFnc(CWnd *pWnd){
	return ((CPMSStockCardReportDialog*)pWnd)->OnByProductLineSelect();
}
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CPMSStockCardReportDialog*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CPMSStockCardReportDialog*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPMSStockCardReportDialog*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CPMSStockCardReportDialog*)pWnd)->OnListDeleteItem();
} 
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSStockCardReportDialog *pVw = (CPMSStockCardReportDialog *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CPMSStockCardReportDialog *pVw = (CPMSStockCardReportDialog *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSStockCardReportDialog *pVw = (CPMSStockCardReportDialog *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CPMSStockCardReportDialog *pVw = (CPMSStockCardReportDialog *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddPMSStockCardReportDialogFnc(CWnd *pWnd){
	 return ((CPMSStockCardReportDialog*)pWnd)->OnAddPMSStockCardReportDialog();
} 
static int _OnEditPMSStockCardReportDialogFnc(CWnd *pWnd){
	 return ((CPMSStockCardReportDialog*)pWnd)->OnEditPMSStockCardReportDialog();
} 
static int _OnDeletePMSStockCardReportDialogFnc(CWnd *pWnd){
	 return ((CPMSStockCardReportDialog*)pWnd)->OnDeletePMSStockCardReportDialog();
} 
static int _OnSavePMSStockCardReportDialogFnc(CWnd *pWnd){
	 return ((CPMSStockCardReportDialog*)pWnd)->OnSavePMSStockCardReportDialog();
} 
static int _OnCancelPMSStockCardReportDialogFnc(CWnd *pWnd){
	 return ((CPMSStockCardReportDialog*)pWnd)->OnCancelPMSStockCardReportDialog();
} 
CPMSStockCardReportDialog::CPMSStockCardReportDialog(CWnd *pParent, int bDetail)
	{
	m_nPrintID = bDetail;
	m_nDlgWidth = 585;
	m_nDlgHeight = 590;
	SetDefaultValues();
}
CPMSStockCardReportDialog::~CPMSStockCardReportDialog(){
}
void CPMSStockCardReportDialog::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 550);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 110, 55);
	m_wndFromDate.Create(this,115, 30, 285, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 290, 30, 390, 55);
	m_wndToDate.Create(this,395, 30, 570, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 110, 85);
	m_wndStock.Create(this,115, 60, 570, 85); 
	m_wndNameLabel.Create(this, _T("Name /Cnt"), 10, 90, 110, 115);
	m_wndName.Create(this,115, 90, 570, 115); 
	m_wndByProductLine.Create(this, _T("By Product Line"), 5, 555, 175, 580);
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 385, 555, 495, 580);
	m_wndExport.Create(this, _T("&Export"), 500, 555, 575, 580);
	m_wndList.Create(this,10, 120, 570, 545); 
}
void CPMSStockCardReportDialog::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);
	m_wndName.SetLimitText(35);
	m_wndName.SetCheckValue(true);
	m_wndName.SetNotEmpty(false);

	m_wndStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStock.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndList.SetCheckBox(TRUE);
	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndList.InsertColumn(1, _T("Name /Cnt"), CFMT_TEXT, 150);
	m_wndList.InsertColumn(2, _T("Unit"), CFMT_TEXT, 80);
	m_wndList.InsertColumn(3, _T("Generic"), CFMT_TEXT, 100);
	m_wndList.InsertColumn(4, _T("Org Name"), CFMT_TEXT, 80);
	m_wndList.InsertColumn(5, _T("Product ID"), CFMT_TEXT, 0);

}
void CPMSStockCardReportDialog::OnSetWindowEvents(){
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
	//m_wndName.SetEvent(WE_CHANGE, _OnNameChangeFnc);
	//m_wndName.SetEvent(WE_SETFOCUS, _OnNameSetfocusFnc);
	//m_wndName.SetEvent(WE_KILLFOCUS, _OnNameKillfocusFnc);
	m_wndName.SetEvent(WE_CHECKVALUE, _OnNameCheckValueFnc);
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);
	m_wndList.SetWindowText(_T("Drug List"));
	m_wndList.AddEvent(1, _T("Select All"), _OnListSelectAllFnc);
	m_wndList.AddEvent(2, _T("Unselect All"), _OnListUnselectAllFnc);
	m_wndByProductLine.SetEvent(WE_CLICK, _OnByProductLineSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_szFromDate.Format(_T("%s 00:00"), pMF->GetSysDate());
	m_szToDate = pMF->GetSysDate() + _T("23:59");
	SetMode(VM_EDIT);

}
void CPMSStockCardReportDialog::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_Text(pDX, m_wndName.GetDlgCtrlID(), m_szName);
	DDX_Check(pDX, m_wndByProductLine.GetDlgCtrlID(), m_bByProductLine);

}
void CPMSStockCardReportDialog::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPMSStockCardReportDialog::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CPMSStockCardReportDialog::SetDefaultValues()
{

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_szName.Empty();
	m_bByProductLine = FALSE;

}
int CPMSStockCardReportDialog::SetMode(int nMode)
{
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
		if (m_nPrintID > 1)
			m_wndExport.EnableWindow(FALSE);
 		UpdateData(FALSE); 
 		return nOldMode; 
}
/*void CPMSStockCardReportDialog::OnFromDateChange(){
	
} */
/*void CPMSStockCardReportDialog::OnFromDateSetfocus(){
	
} */
/*void CPMSStockCardReportDialog::OnFromDateKillfocus(){
	
} */
int CPMSStockCardReportDialog::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSStockCardReportDialog::OnToDateChange(){
	
} */
/*void CPMSStockCardReportDialog::OnToDateSetfocus(){
	
} */
/*void CPMSStockCardReportDialog::OnToDateKillfocus(){
	
} */
int CPMSStockCardReportDialog::OnToDateCheckValue(){
	return 0;
} 
void CPMSStockCardReportDialog::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSStockCardReportDialog::OnStockSelendok(){
	UpdateData(true);
	OnListLoadData();
	 
}
/*void CPMSStockCardReportDialog::OnStockSetfocus(){
	
}*/
/*void CPMSStockCardReportDialog::OnStockKillfocus(){
	
}*/
long CPMSStockCardReportDialog::OnStockLoadData()
{
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageList(&m_wndStock, m_szStockKey);
}
/*void CPMSStockCardReportDialog::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CPMSStockCardReportDialog::OnNameChange(){
	
} */
/*void CPMSStockCardReportDialog::OnNameSetfocus(){
	
} */
/*void CPMSStockCardReportDialog::OnNameKillfocus(){
	
} */
int CPMSStockCardReportDialog::OnNameCheckValue(){
	OnListLoadData();
	return 1;
}


int CPMSStockCardReportDialog::OnListSelectAll(){
	for (int i =0; i < m_wndList.GetItemCount(); i++)
	{
			m_wndList.SetCheck(i, TRUE);
	}
	return 0;
} 

int CPMSStockCardReportDialog::OnListUnselectAll(){
	for (int i =0; i < m_wndList.GetItemCount(); i++)
	{
		m_wndList.SetCheck(i, FALSE);
	}
	return 0;
} 
void CPMSStockCardReportDialog::OnListDblClick(){
	UpdateData(TRUE);
	int nSel = m_wndList.GetCurSel();
	if(nSel < 0) return;
	for (int i=0; i< m_wndList.GetItemCount(); i++)
	{
		m_wndList.SetCheck(i,false);
	}
	m_wndList.SetCheck(nSel, !m_wndList.GetCheck(nSel));
} 
void CPMSStockCardReportDialog::OnListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CPMSStockCardReportDialog::OnListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
} 
long CPMSStockCardReportDialog::OnListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	UpdateData(true);

	m_wndList.BeginLoad(); 
	int nCount = 0;
	if (m_nPrintID == 1 || m_nPrintID == 2);
		//szWhere.Format(_T(" AND msl_qtyonhand > 0"));
	else
		szWhere.Format(_T(" AND msl_qtyordered > 0"));
	
	szWhere.Empty();

	if(!m_szName.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and lower(product_name) like(lower('%s%s%s')) "), _T("%"), m_szName, _T("%"));
	}
	szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	if (!m_bByProductLine)
		szSQL.Format(_T(" SELECT DISTINCT product_code AS code, ") \
					_T("                  product_name AS name, ") \
					_T("                  product_uomname AS unit, ") \
					_T("                  product_classname AS genericname, ") \
					_T("                  product_countryname AS orgname, ") \
					_T("                  product_id AS id ") \
					_T(" FROM m_storageline ") \
					_T(" LEFT JOIN m_productitem_view ON(msl_product_id = product_id) ") \
					_T(" WHERE msl_storage_id = %s ") \
					_T(" AND msl_product_item_id > 0 ") \
					_T(" %s ") \
					_T(" ORDER BY name,unit "), m_szStockKey, szWhere);
	else
		szSQL.Format(_T(" SELECT product_code AS code, ") \
					_T("         product_name AS name, ") \
					_T("         product_uomname AS unit, ") \
					_T("         product_classname AS genericname, ") \
					_T("         product_countryname AS orgname, ") \
					_T("         product_item_id AS id ") \
					_T(" FROM m_storageline ") \
					_T(" LEFT JOIN m_productitem_view ON(msl_product_item_id = product_item_id) ") \
					_T(" WHERE msl_storage_id = %s ") \
					_T(" AND msl_product_item_id > 0 ") \
					_T(" %s ") \
					_T(" ORDER BY name,unit "), m_szStockKey, szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndList.AddItems(
			rs.GetValue(_T("code")), 
			rs.GetValue(_T("Name")), 
			rs.GetValue(_T("Unit")), 
			rs.GetValue(_T("GenericName")), 
			rs.GetValue(_T("OrgName")),
			rs.GetValue(_T("id")),
			NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
}

void CPMSStockCardReportDialog::OnByProductLineSelect(){
	OnListLoadData();
}

void CPMSStockCardReportDialog::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	switch (m_nPrintID)
	{
		case 1:
			OnPrint(true);
			break;
		case 2:
			OnOrderPrint(true);
			break;
		case 3:
			OnPrint2(true);
			break;
		case 4:
			OnOrderPrint2(true);
			break;
	}
} 
void CPMSStockCardReportDialog::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	switch (m_nPrintID)
	{
		case 1:
			OnPrint();
			break;
		case 2:
			OnOrderPrint();
			break;
	}

	
} 
void CPMSStockCardReportDialog::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CExcel xls;
	CString tmpStr;
	int nSheet = 0;
	xls.CreateSheet(1);
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (nSheet > 0)
				xls.AddSheet(_T(""));
			xls.SetWorksheet(nSheet);
			tmpStr = m_wndList.GetItemText(i, 1);
			xls.SetSheetName(tmpStr);
			xls.SetCellText(0, 0, int2str(i));
			nSheet++;
		}
	}
	xls.Save(_T("Exports\\multi_sheet.xls"));
} 
void CPMSStockCardReportDialog::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CPMSStockCardReportDialog::OnAddPMSStockCardReportDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CPMSStockCardReportDialog::OnEditPMSStockCardReportDialog(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CPMSStockCardReportDialog::OnDeletePMSStockCardReportDialog(){
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
 		OnCancelPMSStockCardReportDialog(); 
 	}
	return 0;
}
int CPMSStockCardReportDialog::OnSavePMSStockCardReportDialog(){
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
 		//OnPMSStockCardReportDialogListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CPMSStockCardReportDialog::OnCancelPMSStockCardReportDialog(){
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
int CPMSStockCardReportDialog::OnPMSStockCardReportDialogListLoadData(){
	return 0;
}

void CPMSStockCardReportDialog::OnPrint(bool bPreview)
{
	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
	CReport rpt;
	CString tmpStr, szType, szCategory, szSQL, szWhere, szSysDate, szDate, szRptName, szGroupBy, szDesc;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	double	nOnhand = 0, nTemp = 0;
	szRptName.Format(_T("Reports/HMS/PMS_STOCKCARD.RPT"));
	if(!rpt.Init(szRptName,true) )
		return;
	UpdateData(true);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	rpt.GetReportHeader()->SetValue(_T("StockName"), m_wndStock.GetCurrent((1)));
	//Page Header
	//Report Detail

	//A: Dieu chuyen kho
	//B: Cap phat hang ngay
	//C: Cap phat tu truc
	//D: Xuat thanh ly
	//E: Xuat het han
	//F: Xuat hong vo		
	//G: Xuat dieu chinh
	//H: Xuat tuyen xa
	//O: Xuat khac
	//X: Phieu hoan tra tu kho
	//Y: Phieu hoan tra tu khoa(Phieu linh)
	//Z: Phieu hoan tra tu tu truc


	CReportSection* rptDetail=NULL, *rptHeader = NULL, *rptFooter = NULL;
	CString szCode;
	int nOrderID = 0;
	for (int i =0; i < m_wndList.GetItemCount(); i++)
	{
		if(m_wndList.GetCheck(i)){			
			if(rptFooter)	
			{
				rptFooter->SetPageBreak();
			}
			szCode = m_wndList.GetItemText(i, 0);
			tmpStr = m_wndList.GetItemText(i, 5);
			nOrderID = ToInt(tmpStr);
			rptHeader = rpt.AddDetail(rpt.GetGroupHeader(1));
			rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
			rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
			szSysDate = pMF->GetSysDate(); 
			szDate.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
			rptHeader->SetValue(_T("ReportDate"), szDate);
			rptHeader->SetValue(_T("StockName"), m_wndStock.GetCurrent((1)));

			rptHeader->SetValue(_T("DrugName"), m_wndList.GetItemText(i, 1));
			
			rptHeader->SetValue(_T("ItemID"), szCode);
			rptHeader->SetValue(_T("Unit"), m_wndList.GetItemText(i, 2));
			rptHeader->SetValue(_T("GenericName"), m_wndList.GetItemText(i, 3));
			rptHeader->SetValue(_T("OrgName"), m_wndList.GetItemText(i, 4));
			//tinh ton dau ky

			if (m_bByProductLine)
				szWhere.Format(_T(" AND product_item_id = %d"), nOrderID);
			else
				szWhere.Format(_T(" AND product_id = %d"), nOrderID);
			
			szSQL.Format(_T(" SELECT sum(impqty-expqty) as onhand") \
				_T(" FROM") \
				_T(" (") \
				_T(" 	SELECT impstockid, impdate as iodate, sitemid, impqty, 0 as expqty") \
				_T(" 	FROM MIV") \
				_T(" 	UNION ALL") \
				_T(" 	SELECT expstockid, expdate, sitemid, 0, expqty") \
				_T(" 	FROM MEV") \
				_T(" ) tbl") \
				_T(" LEFT JOIN m_productitem_view ON(sitemid=product_item_id)") \
				_T(" WHERE impstockid=%d") \
				_T(" 	AND iodate < cast_string2timestamp('%s') ") \
				_T(" 	%s"), ToInt(m_szStockKey), m_szFromDate, szWhere);
			rs.ExecSQL(szSQL);
			if(!rs.IsEOF())
				rs.GetValue(_T("onhand"), nOnhand);
			/*else
				nOnhand = 0;*/
			szSysDate = pMF->GetSysDate(); 
			szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
			rptHeader->SetValue(_T("PeriodQuantity"), ToString(nOnhand));
			//line nhap xuat
			if (m_bByProductLine)
				szGroupBy = _T(", product_item_id");
			szSQL.Format(_T(" SELECT impstockid,") \
						_T("   expstockid,") \
						_T("   iodate AS invoicedate,") \
						_T("   Cast_Timestamp2Date(orderdate) AS orderdate,") \
						_T("   invoiceno,") \
						_T("   invoiceid,") \
						_T("   SUM(qty)     AS qty,") \
						_T("   product_id       AS itemid,") \
						_T("   deptid,") \
						_T("   iotype,") \
						_T("   get_doctype(iotype) as iotypename,") \
						_T("   category") \
						_T(" FROM") \
						_T("   (SELECT impstockid,") \
						_T("	 expstockid,") \
						_T("	 invoiceid,") \
						_T("     impdate    AS iodate,") \
						_T("     impinvoice AS invoiceno,") \
						_T("	 orderdate,") \
						_T("     sitemid,") \
						_T("     impqty AS qty,") \
						_T("	 deptid,") \
						_T("     iotype,") \
						_T("     CAST('I' AS NVARCHAR2(1)) as category") \
						_T("   FROM MIV") \
						_T("   UNION ALL") \
						_T("   SELECT expstockid,") \
						_T("	 impstockid,") \
						_T("	 invoiceid,") \
						_T("     expdate,") \
						_T("     expinvoice,") \
						_T("	 orderdate,") \
						_T("     sitemid,") \
						_T("     expqty,") \
						_T("	 deptid,") \
						_T("     iotype,") \
						_T("     category") \
						_T("   FROM MEV") \
						_T("   ) iotbl") \
						_T(" LEFT JOIN m_productitem_view ON(product_item_id=sitemid)") \
						_T(" WHERE qty > 0") \
						_T(" AND impstockid=%d") \
						_T(" AND iodate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s:59') ") \
						_T(" %s") \
						_T(" GROUP BY impstockid, expstockid, iodate, orderdate, invoiceno, ") \
						_T("		  product_id, iotype, category, deptid, invoiceid %s") \
						_T(" ORDER BY invoicedate"), ToInt(m_szStockKey), m_szFromDate, m_szToDate.Left(16), szWhere, szGroupBy);
			int nIdx = 0;
			long nTotalImport =0, nTotalExport=0;
			rs.ExecSQL(szSQL);
			if (rs.IsEOF())
			{
				rptFooter = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptFooter->SetValue(_T("s6"), _T("0"));
				rptFooter->SetValue(_T("s7"), _T("0"));
				FormatCurrency(nOnhand, tmpStr);
				rptFooter->SetValue(_T("s8"), tmpStr);
				/*rptFooter = rpt.GetReportFooter();
				rptFooter->SetValue(_T("PrintDate"), szDate);
				tmpStr = pMF->GetUserName(pMF->GetCurrentUser());
				rptFooter->SetValue(_T("username"), tmpStr);*/
				//if(bPreview)
				//	rpt.PrintPreview();
				//else
				//	rpt.Print();
				//return;
			}
			while(!rs.IsEOF()){
				nIdx++;
				rptDetail = rpt.AddDetail();
				tmpStr.Format(_T("%d"), nIdx);
				rptDetail->SetValue(_T("1"), tmpStr);
				rs.GetValue(_T("iotype"), szType);
				rs.GetValue(_T("category"), szCategory);
				rs.GetValue(_T("invoiceno"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rs.GetValue(_T("orderdate"), tmpStr);
				rptDetail->SetValue(_T("3"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));	
				rs.GetValue(_T("invoicedate"), tmpStr);
				rptDetail->SetValue(_T("5"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
				if(szCategory == _T("I"))
				{					
					//DC
					if(szType == _T("MOV") || szType == _T("MRO"))
					{
						_debug(_T("DC"));
						rs.GetValue(_T("impstockid"), tmpStr);
						szSQL.Format(_T("select msl_name as name from m_storagelist where msl_storage_id= %d"), str2int(tmpStr));
						_debug(_T("%s"), szSQL);
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("name"), tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);
					}
					else if (szType == _T("DRO"))
					{
						CString szName;
						rs.GetValue(_T("invoiceid"), tmpStr);
						szSQL.Format(_T("SELECT mt_department_id dept_id FROM m_transaction WHERE mt_transaction_id = %s"), tmpStr);
						rs2.ExecSQL(szSQL);
						szName.Format(_T("[%s] %s"),rs2.GetValue(_T("dept_id")), rs.GetValue(_T("iotypename")));
						rptDetail->SetValue(_T("4"), szName);
					}
					else if (szType == _T("RRO"))
					{
						//type I: [Type-Dept] Patient;type W: [Type] Patient
						rs.GetValue(_T("invoiceno"), tmpStr);
						szSQL.Format(_T(" SELECT case when NVL(so_partner_type_id, 'W') = 'I' ") \
									 _T("		 then '['||get_doctype('RRO')||'-'||so_partneraddress||'] '||get_patientname(so_docno) ") \
									 _T("		 else '['||get_doctype('RRO')||'] '||get_patientname(so_docno) end descr") \
									 _T(" FROM sale_order WHERE so_sale_order_id = %s"), tmpStr);
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("descr"), szDesc);
						rptDetail->SetValue(_T("4"), szDesc);
					}
					else if (szType == _T("CRO"))
					{
						//Kho tra
						rs.GetValue(_T("expstockid"), tmpStr);
						szSQL.Format(_T("select msl_name as name from m_storagelist where msl_storage_id= %d"), str2int(tmpStr));
						rs2.ExecSQL(szSQL);
						tmpStr.Format(_T("[%s] %s"), rs2.GetValue(_T("name")), rs.GetValue(_T("iotypename")));
						rptDetail->SetValue(_T("4"), tmpStr);
					}
					else if(szType != _T("POO"))
					{
						_debug(_T("HT"));
						rptDetail->SetValue(_T("4"), rs.GetValue(_T("iotypename")));
					}
					//Nhap
					else
					{
						_debug(_T("Nhap"));
						rs.GetValue(_T("invoiceid"), tmpStr);
						szSQL.Format(_T("select get_partnername(po_partner_id) as name from purchase_order where po_purchase_order_id=%s "), tmpStr);
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("name"), tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);
					
					}
					rs.GetValue(_T("qty"), nTemp);
					nOnhand += nTemp;
					nTotalImport += nTemp;
					FormatCurrency(nTemp, tmpStr);
					rptDetail->SetValue(_T("6"), tmpStr);
				}
				else
				{
					//DT
					if(szType == _T("DPO") || szType == _T("DDO"))
					{
						_debug(_T("DT"));
						rs.GetValue(_T("deptid"), tmpStr);
						szSQL.Format(_T("select sd_name as name from sys_dept where sd_id='%s'"), tmpStr);
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("name"), tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);
					}
					//DC
					else if(szType == _T("MOV") || szType == _T("MRO"))
					{
						_debug(_T("DC"));
						rs.GetValue(_T("expstockid"), tmpStr);
						szSQL.Format(_T("select msl_name as name from m_storagelist where msl_storage_id= %d"), str2int(tmpStr));
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("name"), tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);
					}
					//BN
					else if(szType == _T("PPO") || szType == _T("M"))
					{
						CString szName;
						rs.GetValue(_T("invoiceno"), tmpStr);
						szSQL.Format(_T("select trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname from hms_patient left join hms_doc on(hp_patientno=hd_patientno) where hd_docno=%ld"), ToLong(tmpStr));
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("pname"), tmpStr);
						szName.Format(_T("[%s] %s"), rs.GetValue(_T("deptid")), tmpStr);
						rptDetail->SetValue(_T("4"), szName);						
					}
					else if (szType == _T("SOO"))
					{
						//type I: [Type-Dept] Patient;type W: [Type] Patient
						rs.GetValue(_T("invoiceid"), tmpStr);
						szSQL.Format(_T(" SELECT case when NVL(so_partner_type_id, 'W') = 'I' ") \
									 _T("		 then '['||get_syssel_desc('sale_order_type', 'I')||'-'||so_partneraddress||'] '||get_patientname(so_docno) ") \
									 _T("		 else '['||get_syssel_desc('sale_order_type', 'W')||'] '||get_patientname(so_docno) end descr") \
									 _T(" FROM sale_order WHERE so_sale_order_id = %s"), tmpStr);
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("descr"), szDesc);
						rptDetail->SetValue(_T("4"), szDesc);
					}
					else if (szType == _T("CON"))
					{
						rs.GetValue(_T("invoiceid"), tmpStr);
						szSQL.Format(_T("SELECT '['||mt_department_to_id||'] '||mt_description descr FROM m_transaction WHERE mt_transaction_id = %s"), tmpStr);
						rs2.ExecSQL(szSQL);
						rptDetail->SetValue(_T("4"), rs2.GetValue(_T("descr")));
					}
					else
					{
						tmpStr.Format(_T("[%s] %s"), rs.GetValue(_T("deptid")), rs.GetValue(_T("iotypename")));
						rptDetail->SetValue(_T("4"), tmpStr);
					}
 
					rs.GetValue(_T("qty"), nTemp);
					nOnhand -= nTemp;
					nTotalExport += nTemp;
					FormatCurrency(nTemp, tmpStr);
					rptDetail->SetValue(_T("7"), tmpStr); 
				}
				FormatCurrency(nOnhand, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);
				rptDetail->SetValue(_T("9"), _T(""));
				rs.MoveNext();
			}
			if(nTotalImport + nTotalExport > 0){
				CString szLabel;
				TranslateString(_T("Total"), szLabel);
				rptFooter = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptFooter->GetItem(_T("s6"))->SetBold(true);
				rptFooter->GetItem(_T("s7"))->SetBold(true);
				rptFooter->GetItem(_T("s8"))->SetBold(true);
				FormatCurrency(nTotalImport, tmpStr);
				rptFooter->SetValue(_T("s6"), tmpStr);
				FormatCurrency(nTotalExport, tmpStr);
				rptFooter->SetValue(_T("s7"), tmpStr);
				FormatCurrency(nOnhand, tmpStr);
				rptFooter->SetValue(_T("s8"), tmpStr);
				
			}
		}
	}
	//Page Footer
	//Report Footer
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	tmpStr = pMF->GetUserName(pMF->GetCurrentUser());
	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print();
	
}

void CPMSStockCardReportDialog::OnPrint2(bool bPreview)
{
	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
	UpdateData(true);
	CReport rpt;
	CString tmpStr, szType, szCategory, szSQL, szWhere, szSysDate, szDate, szRptName, szGroupBy;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	double	nOnhand = 0, nTemp = 0;
	szRptName.Format(_T("Reports/HMS/PMS_STOCKCARD.RPT"));
	if(!rpt.Init(szRptName,true) )
		return;
	
	//Page Header
	//Report Detail

	//A: Dieu chuyen kho
	//B: Cap phat hang ngay
	//C: Cap phat tu truc
	//D: Xuat thanh ly
	//E: Xuat het han
	//F: Xuat hong vo		
	//G: Xuat dieu chinh
	//H: Xuat tuyen xa
	//O: Xuat khac
	//X: Phieu hoan tra tu kho
	//Y: Phieu hoan tra tu khoa(Phieu linh)
	//Z: Phieu hoan tra tu tu truc


	CReportSection* rptDetail=NULL, *rptHeader = NULL, *rptFooter = NULL;
	CString szCode;
	int nOrderID = 0;
	
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	for (int i =0; i < m_wndList.GetItemCount(); i++)
	{
		if(m_wndList.GetCheck(i)){			
			if(rptDetail)	
			{
				rptDetail->SetPageBreak();
			}

			szCode = m_wndList.GetItemText(i, 0);
			tmpStr = m_wndList.GetItemText(i, 5);
			nOrderID = ToInt(tmpStr);
			rptHeader = rpt.AddDetail(rpt.GetGroupHeader(1));
			rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
			rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
			szDate.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
			rptHeader->SetValue(_T("ReportDate"), szDate);
			rptHeader->SetValue(_T("StockName"), m_wndStock.GetCurrent((1)));

			rptHeader->SetValue(_T("DrugName"), m_wndList.GetItemText(i, 1));
			
			rptHeader->SetValue(_T("ItemID"), szCode);
			rptHeader->SetValue(_T("Unit"), m_wndList.GetItemText(i, 2));
			rptHeader->SetValue(_T("GenericName"), m_wndList.GetItemText(i, 3));
			rptHeader->SetValue(_T("OrgName"), m_wndList.GetItemText(i, 4));
			//tinh ton dau ky
			if (m_bByProductLine)
				szWhere.Format(_T(" AND product_item_id = %d"), nOrderID);
			else
				szWhere.Format(_T(" AND product_id = %d"), nOrderID);
			szSQL.Format(_T(" SELECT sum(import_qty-export_qty) as onhand") \
				_T(" FROM") \
				_T(" (") \
				_T(" 	SELECT *") \
				_T(" 	FROM v_import") \
				_T(" 	UNION ALL") \
				_T(" 	SELECT *") \
				_T(" 	FROM v_export") \
				_T(" ) tbl") \
				_T(" LEFT JOIN m_productitem_view ON(sitemid=product_item_id)") \
				_T(" WHERE storage_id=%d ") \
				_T(" 	AND approved_date < cast_string2timestamp('%s') ") \
				_T(" 	%s"), ToInt(m_szStockKey), m_szFromDate, szWhere);
			rs.ExecSQL(szSQL);

			if(!rs.IsEOF())
				rs.GetValue(_T("onhand"), nOnhand);
			/*else
				nOnhand = 0;*/
			rptHeader->SetValue(_T("PeriodQuantity"), ToString(nOnhand));
			//line nhap xuat
			szSQL.Format(_T(" SELECT storage_id,") \
						_T("   approved_date,") \
						_T("   cast_timestamp2date(order_date) order_date,") \
						_T("   order_no,") \
						_T("   order_id,") \
						_T("   sum(import_qty) import_qty,") \
						_T("   sum(export_qty) export_qty,") \
						_T("   descr") \
						_T(" FROM") \
						_T("   (SELECT *") \
						_T("   FROM v_import") \
						_T("   UNION ALL") \
						_T("   SELECT *") \
						_T("   FROM v_export_detail") \
						_T("   ) iotbl") \
						_T(" LEFT JOIN m_productitem_view ON(product_item_id=sitemid)") \
						_T(" WHERE storage_id=%d") \
						_T(" AND approved_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s:59')") \
						_T(" %s") \
						_T(" GROUP BY storage_id, approved_date, order_date, order_no, order_id, descr") \
						_T(" ORDER BY approved_date"), ToInt(m_szStockKey), m_szFromDate, m_szToDate.Left(16), szWhere);
			int nIdx = 0;
			long nTotalImport =0, nTotalExport=0;
			rs.ExecSQL(szSQL);
			if (rs.IsEOF())
			{
				rptDetail = rpt.AddDetail();
				rptDetail->SetValue(_T("6"), _T("0"));
				rptDetail->SetValue(_T("7"), _T("0"));
				FormatCurrency(nOnhand, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);
				//rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
				//tmpStr = pMF->GetUserName(pMF->GetCurrentUser());
				//rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
				//if(bPreview)
				//	rpt.PrintPreview();
				//else
				//	rpt.Print();
				//return;
			}
			while(!rs.IsEOF()){
				nIdx++;
				rptDetail = rpt.AddDetail();
				tmpStr.Format(_T("%d"), nIdx);
				rptDetail->SetValue(_T("1"), tmpStr);
				rs.GetValue(_T("order_no"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rs.GetValue(_T("order_date"), tmpStr);
				rptDetail->SetValue(_T("3"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));	
				rs.GetValue(_T("approved_date"), tmpStr);
				rptDetail->SetValue(_T("5"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
				rptDetail->SetValue(_T("4"), rs.GetValue(_T("descr")));
				rs.GetValue(_T("import_qty"), nTemp);
				nTotalImport += nTemp;
				nOnhand += nTemp;
				FormatCurrency(nTemp, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);
				rs.GetValue(_T("export_qty"), nTemp);
				nTotalExport += nTemp;
				nOnhand -= nTemp;
				FormatCurrency(nTemp, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);
				FormatCurrency(nOnhand, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);
				rptDetail->SetValue(_T("9"), _T(""));
				rs.MoveNext();
			}
			if(nTotalImport + nTotalExport > 0){
				CString szLabel;
				TranslateString(_T("Total"), szLabel);
				rptFooter = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptFooter->GetItem(_T("s6"))->SetBold(true);
				rptFooter->GetItem(_T("s7"))->SetBold(true);
				rptFooter->GetItem(_T("s8"))->SetBold(true);
				FormatCurrency(nTotalImport, tmpStr);
				rptFooter->SetValue(_T("s6"), tmpStr);
				FormatCurrency(nTotalExport, tmpStr);
				rptFooter->SetValue(_T("s7"), tmpStr);
				FormatCurrency(nOnhand, tmpStr);
				rptFooter->SetValue(_T("s8"), tmpStr);
				
			}
		}

	}
	//Page Footer
	//Report Footer
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Mid(8, 2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	tmpStr = pMF->GetUserName(pMF->GetCurrentUser());
	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print();
	
}

void CPMSStockCardReportDialog::OnOrderPrint(bool bPreview)
{
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL, szObjects, szWhere, tmpStr2;
	if(!rpt.Init(_T("Reports/HMS/PMS_DANHSACHTHUOCYEUCAUCHUADUYET.RPT"),true) )
		return;
	UpdateData(TRUE);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);

	CString szCode;
	int nOrderID = 0;
	for (int i =0; i < m_wndList.GetItemCount(); i++)
	{
		if(m_wndList.GetCheck(i)){		
			szCode = m_wndList.GetItemText(i, 0);
			rpt.GetReportHeader()->SetValue(_T("Code"),szCode);
			rpt.GetReportHeader()->SetValue(_T("DrugName"),m_wndList.GetItemText(i, 1));
			rpt.GetReportHeader()->SetValue(_T("Unit"),m_wndList.GetItemText(i, 2));
			rpt.GetReportHeader()->SetValue(_T("GenericName"),m_wndList.GetItemText(i, 3));
			rpt.GetReportHeader()->SetValue(_T("OrgName"),m_wndList.GetItemText(i, 4));
			tmpStr = m_wndList.GetItemText(i, 5);
			nOrderID = ToInt(tmpStr);
		}
	}
	rpt.GetReportHeader()->SetValue(_T("StockName"), m_wndStock.GetCurrent(1));

	
	tmpStr.Empty();	
		
	rpt.GetReportHeader()->SetValue(_T("Object"), szObjects);
	szWhere.Format(_T(" WHERE itemid = %d"), nOrderID);
	szWhere.AppendFormat(_T(" AND orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szStockKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND stockid = %d"), ToInt(m_szStockKey));
	szSQL.Format(_T(" SELECT orderno,") \
				_T("   docno,") \
				_T("   pname,") \
				_T("   (orderdate) AS orderdate,") \
				_T("   deptid,") \
				_T("   get_storagename(stockid) stockname,") \
				_T("   SUM(orderqty) AS qty") \
				_T(" FROM ") \
				_T(" m_orderexport_view") \
				_T(" %s") \
				_T(" GROUP BY orderno,") \
				_T("   docno,") \
				_T("   pname,") \
				_T("   orderdate,") \
				_T("   deptid,") \
				_T("   stockid") \
				_T(" ORDER BY orderdate"), szWhere);
	//QueryOpener(szSQL);
	CReportSection* rptDetail;
	CReportSection* rptGroup=NULL;
	CRecord rs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	rs.ExecSQL(szSQL);
	int nItem = 0, nTotalOrder=0;
	long TotalQty=0,m_nQty;
	double nTotalAmount=0, nAmount=0;
	while(!rs.IsEOF())
	{	
		rptDetail = rpt.AddDetail();
		nItem++;
		tmpStr.Format(_T("%d"), nItem);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("orderdate"), tmpStr);	
		rptDetail->SetValue(_T("2"), CDate::Convert(tmpStr,yyyymmdd,ddmmyyyy));

		//rs.GetValue(_T("orderid"), tmpStr);	
		//if (!tmpStr.IsEmpty()) 
		//	rs.GetValue(_T("orderid"), tmpStr);	
		//else
		rs.GetValue(_T("orderno"), tmpStr);	

		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("docno"), tmpStr2);
		if (str2int(tmpStr2) > 0)
			tmpStr.Format(_T("[%s] %s"), tmpStr2, rs.GetValue(_T("pname")));
		else
			rs.GetValue(_T("pname"), tmpStr);
		if (tmpStr.GetLength()==1)
			tmpStr= pMF->GetSelectionString(_T("pms_export_type"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("deptid"), tmpStr);
		rs.GetValue(_T("pname"), tmpStr2);
		if (tmpStr2 == _T("Phi\x1EBFu l\x129nh thu\x1ED1\x63, v\x1EADt t\x1B0"))
			tmpStr = rs.GetValue(_T("stockname"));
		rptDetail->SetValue(_T("5"), tmpStr);
		rs.GetValue(_T("qty"), m_nQty);
		TotalQty+=m_nQty;
		rptDetail->SetValue(_T("6"), m_nQty);
		rs.MoveNext();
	}
	if (TotalQty>0)
	{		
		rptDetail = rpt.AddDetail();
		rptDetail->GetItem(_T("4"))->SetBold(true);
		rptDetail->GetItem(_T("6"))->SetBold(true);	
		TranslateString(_T("Total"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rptDetail->SetValue(_T("6"), TotalQty);
	}
	//Page Footer
	//Report Footer
	tmpStr.Format(_T("%d"), nTotalOrder);
	rpt.GetReportFooter()->SetValue(_T("TotalOrder"), tmpStr);
	tmpStr.Format(_T("%.2f"), nTotalAmount);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	//	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print();

	
}
void CPMSStockCardReportDialog::OnOrderPrint2(bool bPreview)
{
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL, szObjects, szWhere, tmpStr2;
	if(!rpt.Init(_T("Reports/HMS/PMS_DANHSACHTHUOCYEUCAUCHUADUYET.RPT"),true) )
		return;
	UpdateData(TRUE);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);

	CString szCode;
	int nOrderID = 0;
	for (int i =0; i < m_wndList.GetItemCount(); i++)
	{
		if(m_wndList.GetCheck(i)){		
			szCode = m_wndList.GetItemText(i, 0);
			rpt.GetReportHeader()->SetValue(_T("Code"),szCode);
			rpt.GetReportHeader()->SetValue(_T("DrugName"),m_wndList.GetItemText(i, 1));
			rpt.GetReportHeader()->SetValue(_T("Unit"),m_wndList.GetItemText(i, 2));
			rpt.GetReportHeader()->SetValue(_T("GenericName"),m_wndList.GetItemText(i, 3));
			rpt.GetReportHeader()->SetValue(_T("OrgName"),m_wndList.GetItemText(i, 4));
			tmpStr = m_wndList.GetItemText(i, 5);
			nOrderID = ToInt(tmpStr);
		}
	}
	rpt.GetReportHeader()->SetValue(_T("StockName"), m_wndStock.GetCurrent(1));

	
	tmpStr.Empty();	
		
	rpt.GetReportHeader()->SetValue(_T("Object"), szObjects);
	szWhere.Format(_T(" WHERE product_id = %d"), nOrderID);
	szWhere.AppendFormat(_T(" AND order_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szStockKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND storage_id = %d"), ToInt(m_szStockKey));
	szSQL.Format(_T(" SELECT order_no,") \
				_T("   order_date,") \
				_T("   descr,") \
				_T("   sum(export_qty) export_qty") \
				_T(" FROM v_export_pending") \
				_T(" LEFT JOIN m_productitem_view ON (sitemid = product_item_id)") \
				_T(" %s") \
				_T(" GROUP BY order_id, order_no,") \
				_T("   order_date,") \
				_T("   product_id, descr") \
				_T(" ORDER BY order_date"), szWhere);
	//QueryOpener(szSQL);
	CReportSection* rptDetail;
	CReportSection* rptGroup=NULL;
	CRecord rs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	rs.ExecSQL(szSQL);
	int nItem = 0, nTotalOrder=0;
	long TotalQty=0,m_nQty;
	double nTotalAmount=0, nAmount=0;
	while(!rs.IsEOF())
	{	
		rptDetail = rpt.AddDetail();
		nItem++;
		tmpStr.Format(_T("%d"), nItem);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("order_date"), tmpStr);	
		rptDetail->SetValue(_T("2"), CDate::Convert(tmpStr,yyyymmdd,ddmmyyyy));
		rs.GetValue(_T("order_no"), tmpStr);	
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("descr"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("export_qty"), m_nQty);
		TotalQty+=m_nQty;
		rptDetail->SetValue(_T("6"), m_nQty);
		rs.MoveNext();
	}
	if (TotalQty>0)
	{		
		rptDetail = rpt.AddDetail();
		rptDetail->GetItem(_T("4"))->SetBold(true);
		rptDetail->GetItem(_T("6"))->SetBold(true);	
		TranslateString(_T("Total"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rptDetail->SetValue(_T("6"), TotalQty);
	}
	//Page Footer
	//Report Footer
	tmpStr.Format(_T("%d"), nTotalOrder);
	rpt.GetReportFooter()->SetValue(_T("TotalOrder"), tmpStr);
	tmpStr.Format(_T("%.2f"), nTotalAmount);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	//	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print();

	
}

CString CPMSStockCardReportDialog::GetQueryString_FP(long nID){
	CString szSQL;
	
	return szSQL;
}

CString CPMSStockCardReportDialog::GetQueryString_IO(long nID){
	CString szSQL;

	return szSQL;
}

BEGIN_MESSAGE_MAP(CPMSStockCardReportDialog, CGuiView)
	ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSStockCardReportDialog::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}

