#include "stdafx.h"
#include "PMReportConditionForm.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnFromDateKillfocus();
} */
/*static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMReportConditionForm *)pWnd)->OnFromDateCheckValue();
}*/ 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMReportConditionForm *)pWnd)->OnToDateCheckValue();
} 
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMReportConditionForm* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMReportConditionForm *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnStorageAddNew();
}*/
static long _OnSupplierLoadDataFnc(CWnd *pWnd){
	return ((CPMReportConditionForm *)pWnd)->OnSupplierLoadData();
}
static void _OnTypeSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMReportConditionForm* )pWnd)->OnTypeSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnTypeSelendokFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnTypeSelendok();
}
/*static void _OnTypeSetfocusFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnTypeKillfocus();
}*/
/*static void _OnTypeKillfocusFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnTypeKillfocus();
}*/
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CPMReportConditionForm *)pWnd)->OnTypeLoadData();
}
/*static void _OnTypeAddNewFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnTypeAddNew();
}*/
static void _OnCategorySelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMReportConditionForm* )pWnd)->OnCategorySelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnCategorySelendokFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnCategorySelendok();
}
/*static void _OnCategorySetfocusFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnCategoryKillfocus();
}*/
/*static void _OnCategoryKillfocusFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnCategoryKillfocus();
}*/
static long _OnCategoryLoadDataFnc(CWnd *pWnd){
	return ((CPMReportConditionForm *)pWnd)->OnCategoryLoadData();
}
/*static void _OnCategoryAddNewFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnCategoryAddNew();
}*/
static void _OnSourceSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMReportConditionForm* )pWnd)->OnSourceSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSourceSelendokFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnSourceSelendok();
}
/*static void _OnSourceSetfocusFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnSourceKillfocus();
}*/
/*static void _OnSourceKillfocusFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnSourceKillfocus();
}*/
static long _OnSourceLoadDataFnc(CWnd *pWnd){
	return ((CPMReportConditionForm *)pWnd)->OnSourceLoadData();
}
/*static void _OnSourceAddNewFnc(CWnd *pWnd){
	((CPMReportConditionForm *)pWnd)->OnSourceAddNew();
}*/
static void _OnChestInventorySelectFnc(CWnd *pWnd){
	 ((CPMReportConditionForm*)pWnd)->OnChestInventorySelect();
}
static void _OnTypeSortSelectFnc(CWnd *pWnd){
	 ((CPMReportConditionForm*)pWnd)->OnTypeSortSelect();
}
static void _OnCategorySortSelectFnc(CWnd *pWnd){
	 ((CPMReportConditionForm*)pWnd)->OnCategorySortSelect();
}
static void _OnSupplierSortSelectFnc(CWnd *pWnd){
	 ((CPMReportConditionForm*)pWnd)->OnSupplierSortSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMReportConditionForm *pVw = (CPMReportConditionForm *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMReportConditionForm *pVw = (CPMReportConditionForm *)pWnd;
	pVw->OnExportSelect();
} 
CPMReportConditionForm::CPMReportConditionForm(CWnd *pParent){

	m_nDlgWidth = 585;
	m_nDlgHeight = 245;
	SetDefaultValues();
}
CPMReportConditionForm::~CPMReportConditionForm(){
}
void CPMReportConditionForm::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 180);
	m_wndToDateLabel.Create(this, _T("To Date"), 10, 30, 100, 55);
	m_wndToDate.Create(this,105, 30, 245, 55); 
	m_wndStorage.SetCheckBox(true);
	m_wndStorageLabel.Create(this, _T("Storage"), 250, 30, 325, 55);
	m_wndStorage.Create(this,330, 30, 570, 55); 
	m_wndSupplierLabel.Create(this, _T("Supplier"), 10, 60, 100, 85);
	m_wndSupplier.Create(this,105, 60, 570, 85); 
	m_wndTypeLabel.Create(this, _T("Type"), 10, 90, 100, 115);
	m_wndType.Create(this,105, 90, 570, 115); 
	m_wndCategoryLabel.Create(this, _T("Category"), 10, 120, 100, 145);
	m_wndCategory.Create(this,105, 120, 570, 145); 
	m_wndSourceLabel.Create(this, _T("Source"), 10, 150, 100, 175);
	m_wndSource.Create(this,105, 150, 570, 175); 
	m_wndChestInventory.Create(this, _T("Chest Inventory"), 5, 185, 140, 210);
	m_wndTypeSort.Create(this, _T("Type Sort"), 145, 185, 280, 210);
	m_wndCategorySort.Create(this, _T("Category Sort"), 285, 185, 420, 210);
	m_wndSupplierSort.Create(this, _T("Supplier Sort"), 425, 185, 560, 210);
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 380, 215, 490, 240);
	m_wndExport.Create(this, _T("&Export"), 495, 215, 575, 240);

}
void CPMReportConditionForm::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndStorage.SetCheckValue(true);
	m_wndStorage.LimitText(1024);
	m_wndType.SetCheckValue(true);
	m_wndType.LimitText(1024);
	m_wndCategory.SetCheckValue(true);
	m_wndCategory.LimitText(1024);
	m_wndSource.SetCheckValue(true);
	m_wndSource.LimitText(1024);

	m_wndStorage.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStorage.InsertColumn(1, _T("Name"), CFMT_TEXT, 400);

	m_wndSupplier.InsertColumn(0, _T("ID"), CFMT_NUMBER, 80);
	m_wndSupplier.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);

	m_wndType.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndType.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);

	m_wndCategory.Format(3, 2);
	m_wndCategory.InsertColumn(0, _T(""), CFMT_NUMBER, 0);
	m_wndCategory.InsertColumn(1, _T("ID"), CFMT_TEXT, 50);
	m_wndCategory.InsertColumn(2, _T("Name"), CFMT_TEXT, 350);
	
	m_wndSource.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndSource.InsertColumn(1, _T("Name"), CFMT_TEXT, 400);
	
}
void CPMReportConditionForm::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	//m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
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
	m_wndSupplier.SetEvent(WE_LOADDATA, _OnSupplierLoadDataFnc);
	m_wndType.SetEvent(WE_SELENDOK, _OnTypeSelendokFnc);
	//m_wndType.SetEvent(WE_SETFOCUS, _OnTypeSetfocusFnc);
	//m_wndType.SetEvent(WE_KILLFOCUS, _OnTypeKillfocusFnc);
	m_wndType.SetEvent(WE_SELCHANGE, _OnTypeSelectChangeFnc);
	m_wndType.SetEvent(WE_LOADDATA, _OnTypeLoadDataFnc);
	//m_wndType.SetEvent(WE_ADDNEW, _OnTypeAddNewFnc);
	m_wndCategory.SetEvent(WE_SELENDOK, _OnCategorySelendokFnc);
	//m_wndCategory.SetEvent(WE_SETFOCUS, _OnCategorySetfocusFnc);
	//m_wndCategory.SetEvent(WE_KILLFOCUS, _OnCategoryKillfocusFnc);
	m_wndCategory.SetEvent(WE_SELCHANGE, _OnCategorySelectChangeFnc);
	m_wndCategory.SetEvent(WE_LOADDATA, _OnCategoryLoadDataFnc);
	//m_wndCategory.SetEvent(WE_ADDNEW, _OnCategoryAddNewFnc);
	m_wndSource.SetEvent(WE_SELENDOK, _OnSourceSelendokFnc);
	//m_wndSource.SetEvent(WE_SETFOCUS, _OnSourceSetfocusFnc);
	//m_wndSource.SetEvent(WE_KILLFOCUS, _OnSourceKillfocusFnc);
	m_wndSource.SetEvent(WE_SELCHANGE, _OnSourceSelectChangeFnc);
	m_wndSource.SetEvent(WE_LOADDATA, _OnSourceLoadDataFnc);
	//m_wndSource.SetEvent(WE_ADDNEW, _OnSourceAddNewFnc);
	m_wndChestInventory.SetEvent(WE_CLICK, _OnChestInventorySelectFnc);
	m_wndTypeSort.SetEvent(WE_CLICK, _OnTypeSortSelectFnc);
	m_wndCategorySort.SetEvent(WE_CLICK, _OnCategorySortSelectFnc);
	m_wndSupplierSort.SetEvent(WE_CLICK, _OnSupplierSortSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	//m_wndFromDate.EnableWindow(FALSE);
	m_szToDate =pMF->GetSysDate() + _T("23:59");
	UpdateData(false);

}
void CPMReportConditionForm::OnDoDataExchange(CDataExchange* pDX){
	//DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szSupplierKey);
	DDX_TextEx(pDX, m_wndStorage.GetDlgCtrlID(), m_szStorageKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_TextEx(pDX, m_wndCategory.GetDlgCtrlID(), m_szCategoryKey);
	DDX_TextEx(pDX, m_wndSource.GetDlgCtrlID(), m_szSourceKey);
	DDX_Check(pDX, m_wndChestInventory.GetDlgCtrlID(), m_bChestInventory);
	DDX_Check(pDX, m_wndTypeSort.GetDlgCtrlID(), m_bTypeSort);
	DDX_Check(pDX, m_wndCategorySort.GetDlgCtrlID(), m_bCategorySort);
	DDX_Check(pDX, m_wndSupplierSort.GetDlgCtrlID(), m_bSupplierSort);

}
void CPMReportConditionForm::SetDefaultValues(){

	//m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStorageKey.Empty();
	m_szTypeKey.Empty();
	m_szCategoryKey.Empty();
	m_szSourceKey.Empty();
	m_bChestInventory=FALSE;
	m_bTypeSort=FALSE;
	m_bCategorySort=FALSE;
	m_bSupplierSort=FALSE;
}
int CPMReportConditionForm::SetMode(int nMode){
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
/*void CPMReportConditionForm::OnFromDateChange(){
	
} */
/*void CPMReportConditionForm::OnFromDateSetfocus(){
	
} */
/*void CPMReportConditionForm::OnFromDateKillfocus(){
	
} */
/*int CPMReportConditionForm::OnFromDateCheckValue(){
	return 0;
}*/ 
/*void CPMReportConditionForm::OnToDateChange(){
	
} */
/*void CPMReportConditionForm::OnToDateSetfocus(){
	
} */
/*void CPMReportConditionForm::OnToDateKillfocus(){
	
} */
int CPMReportConditionForm::OnToDateCheckValue(){
	return 0;
} 
void CPMReportConditionForm::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMReportConditionForm::OnStorageSelendok(){
	 
}
/*void CPMReportConditionForm::OnStorageSetfocus(){
	
}*/
/*void CPMReportConditionForm::OnStorageKillfocus(){
	
}*/
long CPMReportConditionForm::OnStorageLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageList(&m_wndStorage, m_szStorageKey);
}
/*void CPMReportConditionForm::OnStorageAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
long CPMReportConditionForm::OnSupplierLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*)AfxGetMainWnd();
	return pMF->LoadPartnerList(&m_wndSupplier, m_szSupplierKey);
}
void CPMReportConditionForm::OnTypeSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

} 
void CPMReportConditionForm::OnTypeSelendok(){
	 
}
/*void CPMReportConditionForm::OnTypeSetfocus(){
	
}*/
/*void CPMReportConditionForm::OnTypeKillfocus(){
	
}*/
long CPMReportConditionForm::OnTypeLoadData(){
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
/*void CPMReportConditionForm::OnTypeAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMReportConditionForm::OnCategorySelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMReportConditionForm::OnCategorySelendok(){
	 
}
/*void CPMReportConditionForm::OnCategorySetfocus(){
	
}*/
/*void CPMReportConditionForm::OnCategoryKillfocus(){
	
}*/
long CPMReportConditionForm::OnCategoryLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND mpc_level = 1"));
	return pMF->LoadProductCategoryList(&m_wndCategory, m_szCategoryKey, szFilter);
}
/*void CPMReportConditionForm::OnCategoryAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMReportConditionForm::OnSourceSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMReportConditionForm::OnSourceSelendok(){
	 
}
/*void CPMReportConditionForm::OnSourceSetfocus(){
	
}*/
/*void CPMReportConditionForm::OnSourceKillfocus(){
	
}*/
long CPMReportConditionForm::OnSourceLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadProductResourceList(&m_wndSource, m_szSourceKey);
}
/*void CPMReportConditionForm::OnSourceAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMReportConditionForm::OnChestInventorySelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndSupplierSort.SetCheck(false);
}
void CPMReportConditionForm::OnTypeSortSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndCategorySort.SetCheck(false);
}
void CPMReportConditionForm::OnCategorySortSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndTypeSort.SetCheck(false);	
}
void CPMReportConditionForm::OnSupplierSortSelect(){
	m_wndChestInventory.SetCheck(false);
}
void CPMReportConditionForm::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CGuiMenu menu(this);
	CString tmpStr;
	
	menu.CreatePopupMenu();
	TranslateString(_T("Print Format 1"), tmpStr);
	menu.AppendMenu(MF_BYPOSITION, 1, tmpStr);
	menu.AppendMenu(MF_SEPARATOR);
	TranslateString(_T("Print Format 2"), tmpStr);
	menu.AppendMenu(MF_BYPOSITION, 2, tmpStr);
	CRect rect;
	m_wndPrintPreview.GetWindowRect(&rect);
	int nPos = menu.TrackPopupMenu(TPM_RETURNCMD|TPM_RIGHTALIGN|TPM_BOTTOMALIGN|TPM_NONOTIFY, rect.right, rect.top, this);
	switch (nPos)
	{
		case 1:
			PrintFormat1();
			break;
		case 2:
			PrintFormat2();
			break;
	}
}

void CPMReportConditionForm::PrintFormat1(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CReportSection *rptDetail = NULL, *rptHeader = NULL, *rptFooter = NULL;
	CRecord rs(&pMF->m_db), mrs(&pMF->m_db);
	CString szSQL, szName, tmpStr, szPreviousID, szCurrentID, szPreviousGroup, szCurrentGroup, szCurrentDate, szSumInWord;
	double nCost = 0;
	long double nGroupAmount = 0, nTtlAmount = 0, nChestAmount = 0, nSupplierAmount = 0;
	int n = 0, nIdx = 1;
	if (pMF->GetModuleID() == _T("PM"))
		tmpStr.Format(_T("PMS_BIENBANKIEMKETHUOC.RPT"));
	else
		tmpStr.Format(_T("MA_BIENBANKIEMKETHUOC.RPT"));
	szName.Format(_T("Reports/HMS/%s"), tmpStr);
	if(!rpt.Init(szName)){
			return;
	}
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);

	//Member
	szSQL.Format(_T("SELECT mm_id id, mm_title title, mm_name name FROM m_member ORDER BY mm_id"));
	mrs.ExecSQL(szSQL);
	while(!mrs.IsEOF()){
		tmpStr.Format(_T("Title%d"), nIdx);
		rptHeader->SetValue(tmpStr, mrs.GetValue(_T("title")));
		tmpStr.Format(_T("Member%d"), nIdx);
		rptHeader->SetValue(tmpStr, mrs.GetValue(_T("name")));
		nIdx++;
		mrs.MoveNext();
	}
	nIdx = 1;
	tmpStr.Format(_T("%s"), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	rptHeader->SetValue(_T("StockName"), m_szStockName);
	while (!rs.IsEOF())
	{
		//_debug(_T("Chest:%f|Group:%f|ChestNewID:%s|ChestOldID:%s|GroupNewID:%s|GroupOldID:%s "), nChestAmount, nGroupAmount, szCurrentID, szPreviousID, szCurrentGroup, szPreviousGroup);
		if ((m_bTypeSort || m_bCategorySort) && (!m_bChestInventory && !m_bSupplierSort))
		{
			rs.GetValue(_T("groupid"), szCurrentGroup);
			if (szCurrentGroup != szPreviousGroup)
			{
				if (nGroupAmount > 0)
				{
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
					tmpStr.Format(_T("%f"), nGroupAmount);
					rptDetail->SetValue(_T("s9"), tmpStr);	
					nGroupAmount = 0;
				}
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				TranslateString(rs.GetValue(_T("groupname")), tmpStr);
				rptDetail->SetValue(_T("GroupName"), tmpStr);
				szPreviousGroup = szCurrentGroup;
			}
		}
		if (m_bChestInventory)
		{
			rs.GetValue(_T("location"), szCurrentID);
			if (szCurrentID != szPreviousID)
			{
				if (nChestAmount > 0)
				{
					if ((m_bCategorySort || m_bTypeSort) && nGroupAmount > 0)
					{
						rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
						tmpStr.Format(_T("%f"), nGroupAmount);
						rptDetail->SetValue(_T("s9"), tmpStr);	
						nGroupAmount = 0;
					}
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
					tmpStr.Format(_T("%f"), nChestAmount);
					rptDetail->SetValue(_T("s9"), tmpStr);	
					nChestAmount = 0;
				}
				if (str2int(szPreviousID) > 0 && szCurrentID != szPreviousID)
					rptDetail->SetPageBreak();
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("locationame")));
				szPreviousID = szCurrentID;
			}
		}
		if (m_bSupplierSort)
		{
			rs.GetValue(_T("product_bpartnerid"), szCurrentID);
			if (szCurrentID != szPreviousID)
			{
				if (nSupplierAmount > 0)
				{
					if ((m_bCategorySort || m_bTypeSort) && nGroupAmount > 0)
					{
						rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
						tmpStr.Format(_T("%f"), nGroupAmount);
						rptDetail->SetValue(_T("s9"), tmpStr);	
						nGroupAmount = 0;
					}
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
					tmpStr.Format(_T("%f"), nSupplierAmount);
					rptDetail->SetValue(_T("s9"), tmpStr);	
					nSupplierAmount = 0;
				}
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("partnername")));
				szPreviousID = szCurrentID;
			}
		}
		if (m_bTypeSort || m_bCategorySort)
		{
			rs.GetValue(_T("groupid"), szCurrentGroup);
			if (szCurrentGroup != szPreviousGroup)
			{
				if (nGroupAmount > 0)
				{
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
					tmpStr.Format(_T("%f"), nGroupAmount);
					rptDetail->SetValue(_T("s9"), tmpStr);	
					nGroupAmount = 0;
				}
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("groupname")));
				szPreviousGroup = szCurrentGroup;
			}
		}
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("name")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("unit")));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("country")));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("lotno")));
		tmpStr = CDate::Convert(rs.GetValue(_T("expdate")), yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("7"), tmpStr.IsEmpty()?_T("NA"):tmpStr);
		rptDetail->SetValue(_T("8"), rs.GetValue(_T("instockqty")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("price")));
		rs.GetValue(_T("instockamount"), nCost);
		nGroupAmount += nCost;
		nChestAmount += nCost;
		nSupplierAmount += nCost;
		nTtlAmount += nCost;
		rptDetail->SetValue(_T("9"), double2str(nCost));
		rs.MoveNext();
	}
	if (m_bTypeSort || m_bCategorySort)
	{
		rs.GetValue(_T("groupid"), szCurrentGroup);
		if (szCurrentGroup != szPreviousGroup)
		{
			//rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			//rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("groupname")));
			if (nGroupAmount > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				tmpStr.Format(_T("%f"), nGroupAmount);
				rptDetail->SetValue(_T("s9"), tmpStr);	
			}
			szPreviousGroup = szCurrentGroup;
		}
	}
	if (m_bSupplierSort)
	{
		if (nSupplierAmount > 0)
		{
			rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
			tmpStr.Format(_T("%f"), nSupplierAmount);
			rptDetail->SetValue(_T("s9"), tmpStr);	
		}
	}
	if (m_bChestInventory)
	{
		rs.GetValue(_T("location"), szCurrentID);
		if (szCurrentID != szPreviousID)
		{
			//rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			//rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("locationame")));
			if (nChestAmount > 0)
			{
				//if (m_bTypeSort || m_bCategorySort)
				//{
				//	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				//	tmpStr.Format(_T("%f"), nGroupAmount);
				//	rptDetail->SetValue(_T("s9"), tmpStr);
				//}
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				tmpStr.Format(_T("%f"), nChestAmount);
				rptDetail->SetValue(_T("s9"), tmpStr);
				
			}
			szPreviousID = szCurrentID;
		}
	}
	if (nTtlAmount > 0)
	{
		rptFooter = rpt.GetReportFooter();
		nIdx--;
		rptFooter->SetValue(_T("TotalOrder"), int2str(nIdx));
		tmpStr.Format(_T("%f"), nTtlAmount);
		rptFooter->SetValue(_T("TotalAmount"), tmpStr);
		MoneyToString(tmpStr, szSumInWord);
		rptFooter->SetValue(_T("SumInWord"), szSumInWord);
	}
	szCurrentDate = pMF->GetSysDate();
	tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), szCurrentDate.Mid(8, 2), szCurrentDate.Mid(5, 2), szCurrentDate.Left(4));
	rptFooter->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}

void CPMReportConditionForm::PrintFormat2(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString szSQL, tmpStr, szName, szCurrentGroup, szPreviousGroup, szCurrentID, szPreviousID, szSumInWord, szCurrentDate;
	CReportSection *rptDetail = NULL, *rptHeader = NULL, *rptFooter = NULL;
	CRecord rs(&pMF->m_db), mrs(&pMF->m_db);
	double nCost = 0;
	long double nGroupAmount = 0, nTtlAmount = 0, nChestAmount = 0, nSupplierAmount = 0;
	int n = 0, nIdx = 1;
	tmpStr.Format(_T("PMS_BIENBANKIEMKETHUOC_BB.RPT"));
	szName.Format(_T("Reports/HMS/%s"), tmpStr);
	if(!rpt.Init(szName)){
			return;
	}
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	nIdx = 1;
	tmpStr.Format(_T("%s"), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	rptHeader->SetValue(_T("StockName"), m_szStockName);
	while (!rs.IsEOF())
	{
		//_debug(_T("Chest:%f|Group:%f|ChestNewID:%s|ChestOldID:%s|GroupNewID:%s|GroupOldID:%s "), nChestAmount, nGroupAmount, szCurrentID, szPreviousID, szCurrentGroup, szPreviousGroup);
		if ((m_bTypeSort || m_bCategorySort) && (!m_bChestInventory && !m_bSupplierSort))
		{
			rs.GetValue(_T("groupid"), szCurrentGroup);
			if (szCurrentGroup != szPreviousGroup)
			{
				if (nGroupAmount > 0)
				{
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
					tmpStr.Format(_T("%f"), nGroupAmount);
					rptDetail->SetValue(_T("s9"), tmpStr);	
					nGroupAmount = 0;
				}
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("groupname")));
				szPreviousGroup = szCurrentGroup;
			}
		}
		if (m_bChestInventory)
		{
			rs.GetValue(_T("location"), szCurrentID);
			if (szCurrentID != szPreviousID)
			{
				if (nChestAmount > 0)
				{
					if ((m_bCategorySort || m_bTypeSort) && nGroupAmount > 0)
					{
						rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
						tmpStr.Format(_T("%f"), nGroupAmount);
						rptDetail->SetValue(_T("s9"), tmpStr);	
						nGroupAmount = 0;
					}
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
					tmpStr.Format(_T("%f"), nChestAmount);
					rptDetail->SetValue(_T("s9"), tmpStr);	
					nChestAmount = 0;
				}
				if (str2int(szPreviousID) > 0 && szCurrentID != szPreviousID)
					rptDetail->SetPageBreak();
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("locationame")));
				szPreviousID = szCurrentID;
			}
		}
		if (m_bSupplierSort)
		{
			rs.GetValue(_T("product_bpartnerid"), szCurrentID);
			if (szCurrentID != szPreviousID)
			{
				if (nSupplierAmount > 0)
				{
					if ((m_bCategorySort || m_bTypeSort) && nGroupAmount > 0)
					{
						rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
						tmpStr.Format(_T("%f"), nGroupAmount);
						rptDetail->SetValue(_T("s9"), tmpStr);	
						nGroupAmount = 0;
					}
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
					tmpStr.Format(_T("%f"), nSupplierAmount);
					rptDetail->SetValue(_T("s9"), tmpStr);	
					nSupplierAmount = 0;
				}
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("partnername")));
				szPreviousID = szCurrentID;
			}
		}
		if (m_bTypeSort || m_bCategorySort)
		{
			rs.GetValue(_T("groupid"), szCurrentGroup);
			if (szCurrentGroup != szPreviousGroup)
			{
				if (nGroupAmount > 0)
				{
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
					tmpStr.Format(_T("%f"), nGroupAmount);
					rptDetail->SetValue(_T("s9"), tmpStr);	
					nGroupAmount = 0;
				}
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("groupname")));
				szPreviousGroup = szCurrentGroup;
			}
		}
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("name")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("unit")));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("country")));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("lotno")));
		tmpStr = CDate::Convert(rs.GetValue(_T("expdate")), yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("7"), tmpStr.IsEmpty()?_T("NA"):tmpStr);
		rptDetail->SetValue(_T("8"), rs.GetValue(_T("instockqty")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("price")));
		rs.GetValue(_T("instockamount"), nCost);
		nGroupAmount += nCost;
		nChestAmount += nCost;
		nSupplierAmount += nCost;
		nTtlAmount += nCost;
		rptDetail->SetValue(_T("9"), double2str(nCost));
		rs.MoveNext();
	}
	if (m_bTypeSort || m_bCategorySort)
	{
		rs.GetValue(_T("groupid"), szCurrentGroup);
		if (szCurrentGroup != szPreviousGroup)
		{
			//rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			//rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("groupname")));
			if (nGroupAmount > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				tmpStr.Format(_T("%f"), nGroupAmount);
				rptDetail->SetValue(_T("s9"), tmpStr);	
			}
			szPreviousGroup = szCurrentGroup;
		}
	}
	if (m_bSupplierSort)
	{
		if (nSupplierAmount > 0)
		{
			rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
			tmpStr.Format(_T("%f"), nSupplierAmount);
			rptDetail->SetValue(_T("s9"), tmpStr);	
		}
	}
	if (m_bChestInventory)
	{
		rs.GetValue(_T("location"), szCurrentID);
		if (szCurrentID != szPreviousID)
		{
			//rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			//rptDetail->SetValue(_T("GroupName"), rs.GetValue(_T("locationame")));
			if (nChestAmount > 0)
			{
				//if (m_bTypeSort || m_bCategorySort)
				//{
				//	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				//	tmpStr.Format(_T("%f"), nGroupAmount);
				//	rptDetail->SetValue(_T("s9"), tmpStr);
				//}
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				tmpStr.Format(_T("%f"), nChestAmount);
				rptDetail->SetValue(_T("s9"), tmpStr);
				
			}
			szPreviousID = szCurrentID;
		}
	}
	if (nTtlAmount > 0)
	{
		rptFooter = rpt.GetReportFooter();
		nIdx--;
		rptFooter->SetValue(_T("TotalOrder"), int2str(nIdx));
		tmpStr.Format(_T("%f"), nTtlAmount);
		rptFooter->SetValue(_T("TotalAmount"), tmpStr);
		MoneyToString(tmpStr, szSumInWord);
		rptFooter->SetValue(_T("SumInWord"), szSumInWord);
	}
	szCurrentDate = pMF->GetSysDate();
	tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), szCurrentDate.Mid(8, 2), szCurrentDate.Mid(5, 2), szCurrentDate.Left(4));
	rptFooter->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}

void CPMReportConditionForm::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szPreviousID, szCurrentID, szPreviousGroup, szCurrentGroup;
	CStringArray arrCol;
	int nCol = 0, nRow = 0, nIdx = 1;
	double nCost = 0;
	long double nChestAmount = 0, nGroupAmount = 0, nTtlAmount = 0, nSupplierAmount = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 30);
	xls.SetColumnWidth(2, 5);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(5, 10);
	xls.SetColumnWidth(6, 10);
	xls.SetColumnWidth(7, 15);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 8);
	xls.SetCellMergedColumns(0, 3, 8);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	xls.SetCellText(0, 2, _T("\x42I\xCAN \x42\x1EA2N KI\x1EC2M K\xCA THU\x1ED0\x43 H\xD3\x41 \x43H\x1EA4T V\x1EACT T\x1AF"), 4098, true);
	tmpStr.Format(_T("\x110\x1EBFn ng\xE0y %s"), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098);
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn thu\x1ED1\x63/ h\xE0m l\x1B0\x1EE3ng"));
	arrCol.Add(_T("\x110VT"));
	arrCol.Add(_T("N\x1B0\x1EDB\x63 S\x58"));
	arrCol.Add(_T("H\x1EA1n \x64\xF9ng"));
	arrCol.Add(_T("SL"));
	arrCol.Add(_T("\x110\x1A1n gi\xE1"));
	arrCol.Add(_T("Th\xE0nh ti\x1EC1n"));
	for (int i = 0; i < arrCol.GetCount(); i ++)
	{
		xls.SetCellText(nCol + i,  4, arrCol.GetAt(i), 4098, true);
	}
	nRow = 5;
	while (!rs.IsEOF())
	{
		if ((m_bTypeSort || m_bCategorySort) && (!m_bChestInventory && !m_bSupplierSort))
		{
			rs.GetValue(_T("groupid"), szCurrentGroup);
			if (szCurrentGroup != szPreviousGroup)
			{
				if (nGroupAmount > 0)
				{
					xls.SetCellText(6, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
					tmpStr.Format(_T("%f"), nGroupAmount);
					xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1, true);
					nRow++;
					nGroupAmount = 0;
				}
				xls.SetCellText(0, nRow, rs.GetValue(_T("groupname")), FMT_TEXT, true);
				nRow++;
				szPreviousGroup = szCurrentGroup;
			}
		}
		if (m_bChestInventory)
		{
			rs.GetValue(_T("location"), szCurrentID);
			if (szCurrentID != szPreviousID)
			{
				if (nChestAmount > 0)
				{
					if ((m_bCategorySort || m_bTypeSort) && nGroupAmount > 0)
					{
						xls.SetCellText(6, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
						tmpStr.Format(_T("%f"), nGroupAmount);
						xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1, true);
						nRow++;
						nGroupAmount = 0;
					}
					xls.SetCellText(6, nRow, _T("T\x1ED5ng t\x1EE7"), 4098, true);
					tmpStr.Format(_T("%f"), nChestAmount);
					xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1, true);
					nRow++;
					nChestAmount = 0;
				}
				xls.SetCellText(0, nRow, rs.GetValue(_T("locationame")), FMT_TEXT, true);
				szPreviousID = szCurrentID;
				nRow++;
			}
		}
		if (m_bSupplierSort)
		{
			rs.GetValue(_T("product_bpartnerid"), szCurrentID);
			if (szCurrentID != szPreviousID)
			{
				if (nSupplierAmount > 0)
				{
					if ((m_bCategorySort || m_bTypeSort) && nGroupAmount > 0)
					{
						xls.SetCellText(6, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
						tmpStr.Format(_T("%f"), nGroupAmount);
						xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1, true);
						nRow++;
						nGroupAmount = 0;
					}
					xls.SetCellText(6, nRow, _T("T\x1ED5ng t\x1EE7"), 4098, true);
					tmpStr.Format(_T("%f"), nSupplierAmount);
					xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1, true);
					nRow++;
					nSupplierAmount = 0;
				}
				xls.SetCellText(0, nRow, rs.GetValue(_T("partnername")), FMT_TEXT, true);
				szPreviousID = szCurrentID;
				nRow++;
			}
		}
		if (m_bTypeSort || m_bCategorySort)
		{
			rs.GetValue(_T("groupid"), szCurrentGroup);
			if (szCurrentGroup != szPreviousGroup)
			{
				if (nGroupAmount > 0)
				{
					xls.SetCellText(6, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
					tmpStr.Format(_T("%f"), nGroupAmount);
					xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1, true);
					nRow++;
					nGroupAmount = 0;
				}
				xls.SetCellText(0, nRow, rs.GetValue(_T("groupname")), FMT_TEXT, true);
				nRow++;
				szPreviousGroup = szCurrentGroup;
			}
		}
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("name")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("unit")), FMT_TEXT);
		xls.SetCellText(3, nRow, rs.GetValue(_T("country")), FMT_TEXT);
		tmpStr = CDate::Convert(rs.GetValue(_T("expdate")), yyyymmdd, ddmmyyyy);
		xls.SetCellText(4, nRow, tmpStr.IsEmpty()?_T("NA"):tmpStr, FMT_TEXT);
		xls.SetCellText(5, nRow, rs.GetValue(_T("instockqty")), FMT_NUMBER1);
		xls.SetCellText(6, nRow, rs.GetValue(_T("price")), FMT_NUMBER1);
		rs.GetValue(_T("instockamount"), nCost);
		nGroupAmount += nCost;
		nChestAmount += nCost;
		nSupplierAmount += nCost;
		nTtlAmount += nCost;
		xls.SetCellText(7, nRow, double2str(nCost), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (m_bTypeSort || m_bCategorySort)
	{
		if (nGroupAmount > 0)
		{
			xls.SetCellText(6, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
			tmpStr.Format(_T("%f"), nGroupAmount);
			xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1, true);
			nRow++;
			nGroupAmount = 0;
		}
	}
	if (m_bSupplierSort)
	{
		if (nSupplierAmount > 0)
		{
			xls.SetCellText(6, nRow, _T("T\x1ED5ng t\x1EE7"), 4098, true);
			tmpStr.Format(_T("%f"), nSupplierAmount);
			xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1, true);
			nRow++;
			nSupplierAmount = 0;
		}
	}
	if (m_bChestInventory)
	{
		if (nChestAmount > 0)
		{
			xls.SetCellText(6, nRow, _T("T\x1ED5ng t\x1EE7"), 4098, true);
			tmpStr.Format(_T("%f"), nChestAmount);
			xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1, true);
			nRow++;
			nChestAmount = 0;
		}
	}
	if (nTtlAmount > 0)
	{
		xls.SetCellText(6, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
		tmpStr.Format(_T("%f"), nTtlAmount);
		xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\Bien ban kiem kho kho.xls"));
} 
CString CPMReportConditionForm::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szWhere, szInnerWhere, szStockID, szField, szGroupBy, szOrderBy;
	m_szStockName.Empty();
	for (int i = 0; i < m_wndStorage.GetItemCount(); i++)
	{
		if (m_wndStorage.GetCheck(i))
		{
			m_wndStorage.SetCurSel(i);
			if (!szStockID.IsEmpty())
				szStockID.AppendFormat(_T(", "));
			szStockID.AppendFormat(_T("%s"), m_wndStorage.GetCurrent(0));
			if (!m_szStockName.IsEmpty())
				m_szStockName.AppendFormat(_T(", "));
			m_szStockName.AppendFormat(_T("%s"), m_wndStorage.GetCurrent(1));
		}
	}
	szInnerWhere.Format(_T(" AND io_date < cast_string2timestamp('%s')"), m_szToDate);
	if (!szStockID.IsEmpty())
		szInnerWhere.AppendFormat(_T(" AND storage_id IN (%s)"), szStockID);
	if (!m_szSupplierKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND product_bpartnerid = '%s'"), m_szSupplierKey);
	if (!m_szTypeKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND product_groupid = '%s'"), m_szTypeKey);
	if (!m_szCategoryKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND product_rootid = '%s'"), m_szCategoryKey);
	//if (!m_szSourceKey.IsEmpty())
	//	szWhere.AppendFormat(_T(" AND product_resourceid = %s"), m_szSourceKey);
	szWhere.AppendFormat(_T(" and product_org_id = '%s'"), pMF->GetModuleID());
	szField.Empty();
	szGroupBy.Empty();
	szOrderBy.Empty();
	if (m_bChestInventory)
	{
		if (!szField.IsEmpty())
			szField.AppendFormat(_T(", "));
		szField.AppendFormat(_T("nvl(get_locationname(storage_id, mpl_location_id), 'Thu\x1ED1\x63 \x63h\x1B0\x61 thi\x1EBFt l\x1EADp t\x1EE7') locationame, coalesce(mpl_location_id, 999) location"));
		if (!szGroupBy.IsEmpty())
			szGroupBy.AppendFormat(_T(", "));
		szGroupBy.AppendFormat(_T("storage_id, mpl_location_id"));
		if (!szOrderBy.IsEmpty())
			szOrderBy.AppendFormat(_T(", "));
		szOrderBy.AppendFormat(_T("storage_id, mpl_location_id"));
	}
	if (m_bSupplierSort)
	{
		if (!szField.IsEmpty())
			szField.AppendFormat(_T(", "));
		szField.AppendFormat(_T("product_bpartnerid, get_partnername(product_bpartnerid) partnername"));
		if (!szGroupBy.IsEmpty())
			szGroupBy.AppendFormat(_T(", "));
		szGroupBy.AppendFormat(_T("product_bpartnerid"));
		if (!szOrderBy.IsEmpty())
			szOrderBy.AppendFormat(_T(", "));
		szOrderBy.AppendFormat(_T("product_bpartnerid"));
	}
	if (m_bTypeSort)
	{
		if (!szField.IsEmpty())
			szField.AppendFormat(_T(", "));
		//szField.AppendFormat(_T("product_producttype AS groupid, product_typename AS groupname"));
		szField.AppendFormat(_T("product_groupid AS groupid, product_groupname AS groupname"));
		if (!szGroupBy.IsEmpty())
			szGroupBy.AppendFormat(_T(", "));
		szGroupBy.AppendFormat(_T("product_groupid, product_groupname"));
		if (!szOrderBy.IsEmpty())
			szOrderBy.AppendFormat(_T(", "));
		szOrderBy.AppendFormat(_T("product_groupid"));

	}
	if (m_bCategorySort)
	{
		if (!szField.IsEmpty())
			szField.AppendFormat(_T(", "));
		szField.AppendFormat(_T("product_rootid groupid, product_rootname groupname"));
		if (!szGroupBy.IsEmpty())
			szGroupBy.AppendFormat(_T(", "));
		szGroupBy.AppendFormat(_T("product_rootid, product_rootname"));
		if (!szOrderBy.IsEmpty())
			szOrderBy.AppendFormat(_T(", "));
		szOrderBy.AppendFormat(_T("product_rootid"));
	}
	if (!szField.IsEmpty())
		szField.AppendFormat(_T(", "));
	if (!szGroupBy.IsEmpty())
		szGroupBy.AppendFormat(_T(", "));
	if (!szOrderBy.IsEmpty())
		szOrderBy.AppendFormat(_T(", "));
	szSQL.Format(_T("SELECT %s") \
	_T("		   product_code as id,") \
	_T("           product_name AS name,") \
	_T("           product_uomname AS unit,") \
	_T("           product_lotno AS lotno,") \
	_T("           CASE WHEN product_hasexp = 'Y' THEN product_expdate ELSE NULL END AS expdate,") \
	_T("           product_countryname AS country,") \
	_T("           CASE WHEN NVL(msl_category, 'Z') = 'S' THEN product_saleprice1 ELSE product_taxprice END as price,") \
	_T("           SUM(period_qty) AS instockqty,") \
	_T("           CASE WHEN NVL(msl_category, 'Z') = 'S' THEN sum(period_qty*product_saleprice1) ELSE sum(period_qty*product_taxprice) END AS instockamount") \
	_T("         FROM   (SELECT") \
	_T("                   storage_id,") \
	_T("                   item_id,") \
	_T("                   COALESCE(imp_qty-exp_qty, 0) period_qty") \
	_T("                 FROM   (SELECT") \
	_T("                           impstockid AS storage_id,") \
	_T("                           impdate    AS io_date,") \
	_T("                           sitemid item_id,") \
	_T("                           impqty imp_qty,") \
	_T("                           0          AS exp_qty") \
	_T("                         FROM   miv") \
	_T("                         UNION ALL") \
	_T("                         SELECT") \
	_T("                           expstockid,") \
	_T("                           expdate,") \
	_T("                           sitemid,") \
	_T("                           0,") \
	_T("                           expqty") \
	_T("                         FROM   mev) ptbl") \
	_T("                 WHERE 1=1 ") \
	_T("					%s) instock") \
	_T("                 LEFT JOIN m_productitem_view ON(product_item_id=item_id)") \
	_T("				 LEFT JOIN m_product_location ON (mpl_storage_id = storage_id AND mpl_product_id = product_id)") \
	_T("				 LEFT JOIN m_storagelist ON (msl_storage_id = storage_id)") \
	_T("         WHERE item_id>0 ") \
	_T("				%s") \
	_T("         GROUP  BY %s ") \
	_T("				   product_name,") \
	_T("				   product_code,") \
	_T("                   product_uomname,") \
	_T("                   product_lotno,") \
	_T("                   product_expdate,") \
	_T("                   product_countryname,") \
	_T("                   product_taxprice,") \
	_T("				   product_saleprice1,") \
	_T("				   msl_category,") \
	_T("				   product_hasexp") \
	_T("		HAVING sum(period_qty) > 0 ") \
	_T("		ORDER BY %s ") \
	_T("		product_code, product_expdate, product_taxprice"), szField, szInnerWhere, szWhere, szGroupBy, szOrderBy);
	_msg(_T("%s"), szSQL);
	return szSQL;
}