#include "stdafx.h"
#include "FMSudungthuocBHYT20a_Y2015.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMSudungthuocBHYT20a_Y2015* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnStockAddNew();
}*/
static void _OnTypeSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMSudungthuocBHYT20a_Y2015* )pWnd)->OnTypeSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnTypeSelendokFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnTypeSelendok();
}
/*static void _OnTypeSetfocusFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnTypeKillfocus();
}*/
/*static void _OnTypeKillfocusFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnTypeKillfocus();
}*/
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnTypeLoadData();
}
/*static void _OnTypeAddNewFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnTypeAddNew();
}*/
static void _OnGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMSudungthuocBHYT20a_Y2015* )pWnd)->OnGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnGroupSelendokFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnGroupSelendok();
}
/*static void _OnGroupSetfocusFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnGroupKillfocus();
}*/
/*static void _OnGroupKillfocusFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnGroupKillfocus();
}*/
static long _OnGroupLoadDataFnc(CWnd *pWnd){
	return ((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnGroupLoadData();
}
/*static void _OnGroupAddNewFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnGroupAddNew();
}*/
static void _OnSourceSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMSudungthuocBHYT20a_Y2015* )pWnd)->OnSourceSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSourceSelendokFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnSourceSelendok();
}
/*static void _OnSourceSetfocusFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnSourceKillfocus();
}*/
/*static void _OnSourceKillfocusFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnSourceKillfocus();
}*/
static long _OnSourceLoadDataFnc(CWnd *pWnd){
	return ((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnSourceLoadData();
}
/*static void _OnSourceAddNewFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015 *)pWnd)->OnSourceAddNew();
}*/
static void _OnDrugPTTTSelectFnc(CWnd *pWnd){
	 ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnDrugPTTTSelect();
}
static void _OnExamSelectFnc(CWnd *pWnd){
	  ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnExamSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	  ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnInPatientSelect();
}
static void _OnAllHospitalsSelectFnc(CWnd *pWnd){
	  ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnAllHospitalsSelect();
}
static void _OnOrderBySelectFnc(CWnd *pWnd){
	 ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnOrderBySelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMSudungthuocBHYT20a_Y2015 *pVw = (CFMSudungthuocBHYT20a_Y2015 *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMSudungthuocBHYT20a_Y2015 *pVw = (CFMSudungthuocBHYT20a_Y2015 *)pWnd;
	pVw->OnExportSelect();
}
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnListDeleteItem();
}
static int _OnAddFMSudungthuocBHYT20aFnc(CWnd *pWnd){
	 return ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnAddFMSudungthuocBHYT20a();
} 
static int _OnEditFMSudungthuocBHYT20aFnc(CWnd *pWnd){
	 return ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnEditFMSudungthuocBHYT20a();
} 
static int _OnDeleteFMSudungthuocBHYT20aFnc(CWnd *pWnd){
	 return ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnDeleteFMSudungthuocBHYT20a();
} 
static int _OnSaveFMSudungthuocBHYT20aFnc(CWnd *pWnd){
	 return ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnSaveFMSudungthuocBHYT20a();
} 
static int _OnCancelFMSudungthuocBHYT20aFnc(CWnd *pWnd){
	 return ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnCancelFMSudungthuocBHYT20a();
}

static int _OnListCheckAllFnc(CWnd *pWnd){
	return ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnListCheckAll();
}
static int _OnListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMSudungthuocBHYT20a_Y2015*)pWnd)->OnListUnCheckAll();
}

CFMSudungthuocBHYT20a_Y2015::CFMSudungthuocBHYT20a_Y2015(CWnd *pParent)
{
	m_nDlgWidth = 460;
	m_nDlgHeight = 495;
	SetDefaultValues();
}
CFMSudungthuocBHYT20a_Y2015::~CFMSudungthuocBHYT20a_Y2015()
{
}
void CFMSudungthuocBHYT20a_Y2015::OnCreateComponents()
{
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 150, 425, 485);
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 490);
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 30, 300, 55);
	m_wndToDate.Create(this,305, 30, 425, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 90, 85);
	m_wndStock.Create(this,95, 60, 425, 85); 
	m_wndTypeLabel.Create(this, _T("Type"), 10, 90, 90, 115);
	m_wndType.Create(this,95, 90, 425, 115); 
	m_wndGroupLabel.Create(this, _T("Group"), 10, 120, 90, 145);
	m_wndGroup.Create(this,95, 120, 425, 145); 
	m_wndList.Create(this,15, 175, 420, 480); 
	m_wndDrugPTTT.Create(this, _T("Without Operation Material"), 5, 495, 190, 520);
	m_wndByDischargedDate.Create(this, _T("By Discharged Date"), 195, 495, 380, 520);
	m_wndOnlyCommander.Create(this, _T("Only Commander"), 5, 525, 190, 550);
	m_wndOrderBy.Create(this, _T("Sort by Category"), 195, 525, 380, 550);
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 225, 555, 325, 580);
	m_wndPrintPreview.EnableWindow(FALSE);
	m_wndExport.Create(this, _T("&Export"), 330, 555, 430, 580);
}

void CFMSudungthuocBHYT20a_Y2015::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);

	//m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(128);

	//m_wndType.SetCheckValue(true);
	m_wndType.LimitText(254);

	//m_wndGroup.SetCheckValue(true);
	m_wndGroup.LimitText(81);

	m_wndStock.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	m_wndStock.InsertColumn(1, _T("Name"), CFMT_TEXT, 380);


	m_wndType.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	m_wndType.InsertColumn(1, _T("Description"), CFMT_TEXT, 380);


	m_wndGroup.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndGroup.InsertColumn(1, _T("Description"), CFMT_TEXT, 400);
	m_wndGroup.InsertColumn(2, _T("CategoryID"), CFMT_TEXT, 0);


	m_wndSource.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndSource.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);


	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndList.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);
	m_wndList.SetCheckBox(TRUE);
}
void CFMSudungthuocBHYT20a_Y2015::OnSetWindowEvents()
{
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
	m_wndDrugPTTT.SetEvent(WE_CLICK, _OnDrugPTTTSelectFnc);
	m_wndExam.SetEvent(WE_CLICK, _OnExamSelectFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	m_wndAllHospitals.SetEvent(WE_CLICK, _OnAllHospitalsSelectFnc);
	m_wndOrderBy.SetEvent(WE_CLICK, _OnOrderBySelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);

	m_wndList.SetWindowText(_T("Dept"));
	m_wndList.AddEvent(1, _T("Check All"), _OnListCheckAllFnc);
	m_wndList.AddEvent(2, _T("UnCheck All"), _OnListUnCheckAllFnc);

	m_mapLine[_T("I")] = _T("\x110\xFAng tuy\x1EBFn \x111\x103ng k\xFD t\x1EA1i \x62\x1EC7nh vi\x1EC7n");
	m_mapLine[_T("II")] = _T("\x110\xFAng tuy\x1EBFn \x111\x103ng k\xFD t\x1EA1i ngo\x1EA1i t\x1EC9nh");
	m_mapLine[_T("III")] = _T("\x42\x1EC7nh nh\xE2n tr\xE1i tuy\x1EBFn");

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szFromDate =  pMF->GetSysDate()+ _T("00:00");	
	m_szToDate =pMF->GetSysDate() + _T("23:59");
	UpdateData(false);

	OnListLoadData();
}
void CFMSudungthuocBHYT20a_Y2015::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_TextEx(pDX, m_wndGroup.GetDlgCtrlID(), m_szGroupKey);
	DDX_Check(pDX, m_wndDrugPTTT.GetDlgCtrlID(), m_bDrugPTTT);
	DDX_Radio(pDX, m_wndAllHospitals.GetDlgCtrlID(), m_nAllHospitals);
	DDX_Check(pDX, m_wndOrderBy.GetDlgCtrlID(), m_bOrderBy);
	DDX_Check(pDX, m_wndByDischargedDate.GetDlgCtrlID(), m_bByDischargedDate);
	DDX_Check(pDX, m_wndOnlyCommander.GetDlgCtrlID(), m_bOnlyCommander);

}
void CFMSudungthuocBHYT20a_Y2015::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMSudungthuocBHYT20a_Y2015::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMSudungthuocBHYT20a_Y2015::SetDefaultValues()
{
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_szTypeKey.Empty();
	m_szGroupKey.Empty();
	m_szSourceKey.Empty();

	m_bDrugPTTT=FALSE;
	m_nExam=0;
	m_nInPatient=1;
	m_nAllHospitals=1;
	m_bOrderBy=FALSE;
	m_bByDischargedDate = FALSE;
}
int CFMSudungthuocBHYT20a_Y2015::SetMode(int nMode)
{
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
/*void CFMSudungthuocBHYT20a_Y2015::OnFromDateChange(){
	
} */
/*void CFMSudungthuocBHYT20a_Y2015::OnFromDateSetfocus(){
	
} */
/*void CFMSudungthuocBHYT20a_Y2015::OnFromDateKillfocus(){
	
} */
int CFMSudungthuocBHYT20a_Y2015::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMSudungthuocBHYT20a_Y2015::OnToDateChange(){
	
} */
/*void CFMSudungthuocBHYT20a_Y2015::OnToDateSetfocus(){
	
} */
/*void CFMSudungthuocBHYT20a_Y2015::OnToDateKillfocus(){
	
} */
int CFMSudungthuocBHYT20a_Y2015::OnToDateCheckValue(){
	return 0;
} 
void CFMSudungthuocBHYT20a_Y2015::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMSudungthuocBHYT20a_Y2015::OnStockSelendok(){
	 
}
/*void CFMSudungthuocBHYT20a_Y2015::OnStockSetfocus(){
	
}*/
/*void CFMSudungthuocBHYT20a_Y2015::OnStockKillfocus(){
	
}*/

long CFMSudungthuocBHYT20a_Y2015::OnStockLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndStock.IsSearchKey() && ToInt(m_szStockKey) > 0)
	{
		szWhere.Format(_T(" AND msl_storage_id=%d "), ToInt(m_szStockKey));
	}
	CString szStorages;
	//szStorages = pMF->m_szStoragePerm;
	if (!szStorages.IsEmpty())
		szWhere.AppendFormat(_T(" AND msl_storage_id in(%s) "), szStorages);
	//szWhere.AppendFormat(_T(" AND msl_storage_id in(%s) "), pMF->m_szStorages);

	m_wndStock.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT msl_storage_id as id, ") \
		         _T("msl_name as name ") \
		         _T("FROM m_storagelist ") \
				 _T("WHERE msl_org_id in('GL','PM','WM') and msl_isactive='Y' %s ") \
				 _T("ORDER BY msl_storage_id"),
				 szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStock.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMSudungthuocBHYT20a_Y2015::OnStockAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMSudungthuocBHYT20a_Y2015::OnTypeSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMSudungthuocBHYT20a_Y2015::OnTypeSelendok(){
	 
}
/*void CFMSudungthuocBHYT20a_Y2015::OnTypeSetfocus(){
	
}*/
/*void CFMSudungthuocBHYT20a_Y2015::OnTypeKillfocus(){
	
}*/
long CFMSudungthuocBHYT20a_Y2015::OnTypeLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if (m_wndType.IsSearchKey() && !m_szTypeKey.IsEmpty())
	{
		szWhere.Format(_T(" and mpt_product_type_id='%s' "), m_szTypeKey);
	}

	//szWhere.AppendFormat(_T(" AND hfg_moduleid = '%s'"), pMF->GetModuleID());

	m_wndType.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" select mpt_product_type_id as id,") \
				_T("        mpt_name as name") \
				_T(" from m_product_type") \
				_T(" where substr(mpt_product_type_id, 1, 1)='A'") \
				_T(" order by id"), szWhere);

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
/*void CFMSudungthuocBHYT20a_Y2015::OnTypeAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMSudungthuocBHYT20a_Y2015::OnGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMSudungthuocBHYT20a_Y2015::OnGroupSelendok(){
	 
}
/*void CFMSudungthuocBHYT20a_Y2015::OnGroupSetfocus(){
	
}*/
/*void CFMSudungthuocBHYT20a_Y2015::OnGroupKillfocus(){
	
}*/
long CFMSudungthuocBHYT20a_Y2015::OnGroupLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if (m_wndGroup.IsSearchKey() && !m_szGroupKey.IsEmpty())
	{
		szWhere.Format(_T(" AND pmg_id='%s' "), m_szGroupKey);
	}
	m_wndGroup.DeleteAllItems(); 
	int nCount = 0;

	szSQL.Format(_T(" SELECT mpc_product_category_id as categoryid,") \
				_T("        mpc_code as id,") \
				_T("        mpc_name as name") \
				_T(" FROM m_product_category") \
				_T(" WHERE mpc_level > 0") \
				_T(" CONNECT BY PRIOR mpc_product_category_id=mpc_parent_id") \
				_T(" ORDER SIBLINGS BY mpc_product_category_id"), szWhere);

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF())
	{ 
		m_wndGroup.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")),
			rs.GetValue(_T("categoryid")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CFMSudungthuocBHYT20a_Y2015::OnGroupAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMSudungthuocBHYT20a_Y2015::OnSourceSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMSudungthuocBHYT20a_Y2015::OnSourceSelendok(){
	 
}
/*void CFMSudungthuocBHYT20a_Y2015::OnSourceSetfocus(){
	
}*/
/*void CFMSudungthuocBHYT20a_Y2015::OnSourceKillfocus(){
	
}*/
long CFMSudungthuocBHYT20a_Y2015::OnSourceLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndSource.IsSearchKey() && !m_szSourceKey.IsEmpty()){
	};
	m_wndSource.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndSource.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CFMSudungthuocBHYT20a_Y2015::OnSourceAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
	void CFMSudungthuocBHYT20a_Y2015::OnDrugPTTTSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMSudungthuocBHYT20a_Y2015::OnExamSelect()
{
	UpdateData(TRUE);
	OnListLoadData();
}
void CFMSudungthuocBHYT20a_Y2015::OnInPatientSelect()
{
	UpdateData(TRUE);
	OnListLoadData();
}
void CFMSudungthuocBHYT20a_Y2015::OnAllHospitalsSelect()
{
	UpdateData(TRUE);
	OnListLoadData();
}
void CFMSudungthuocBHYT20a_Y2015::OnOrderBySelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}

void CFMSudungthuocBHYT20a_Y2015::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere, tmpStr, szDate, szSysDate;
	szSQL = GetQueryString();
	BeginWaitCursor();

	int nCount = rs.ExecSQL(szSQL);

	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/PMS_THONGKETONGHOPSUDUNGTHUOC.RPT"), true))
	{
		return;
	}

	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);

	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm ),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	rpt.GetReportHeader()->SetValue(_T("StockName"), m_szStockName);
	rpt.GetReportHeader()->SetValue(_T("Type"), m_wndType.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Group"), m_wndGroup.GetCurrent(1));
	//rpt.GetReportHeader()->SetValue(_T("Original"), m_wndSource.GetCurrent(1));
	CReportSection* rptDetail;
	CString szOldLine, szNewLine, szAmount, szOrderBy;

	long double grpCost = 0, ttlCost = 0;
	double cost = 0;
	int nItem = 1;

	if (m_bOrderBy)
		szOrderBy.Format(_T("groupname"));
	//else
	//	szOrderBy.Format(_T("typeName"));

	while(!rs.IsEOF())
	{
		rs.GetValue(szOrderBy, szNewLine);

		if(szNewLine != szOldLine && !szNewLine.IsEmpty())
		{
			if(grpCost > 0)
			{
				TranslateString(_T("Total Group"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("TotalGroup"), szAmount);						
				FormatCurrency(grpCost, tmpStr);
				rptDetail->SetValue(_T("s9"), tmpStr);
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

		rs.GetValue(_T("drugname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("lvl"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("dosage"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("serial"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("qty"), tmpStr);
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("price"), cost);
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("amount"), cost);
		grpCost += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		/*rs.GetValue(_T("serial"), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);*/

		rs.MoveNext();
	}

	if (grpCost > 0)
	{
		TranslateString(_T("Total Group"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));	
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		FormatCurrency(grpCost, tmpStr);
		rptDetail->SetValue(_T("s9"), tmpStr);
		ttlCost += grpCost;
	}

	if (ttlCost > 0)
	{
		TranslateString(_T("Total Amount"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		FormatCurrency(ttlCost, tmpStr);
		rptDetail->SetValue(_T("s9"), tmpStr);
		ttlCost += grpCost;
	}

	szSysDate = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
		          szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();

}

void CFMSudungthuocBHYT20a_Y2015::OnPrintDetail(){

}

void CFMSudungthuocBHYT20a_Y2015::OnExportSelect()
{
	CGuiMenu menu(this);
	CString tmpStr;
	
	menu.CreatePopupMenu();
	TranslateString(_T("Export 20A Format"), tmpStr);
	menu.AppendMenu(MF_BYPOSITION, 1, tmpStr);
	menu.AppendMenu(MF_SEPARATOR);
	TranslateString(_T("Export Detail Format"), tmpStr);
	menu.AppendMenu(MF_BYPOSITION, 2, tmpStr);
	CRect rect;
	m_wndExport.GetWindowRect(&rect);
	int nPos = menu.TrackPopupMenu(TPM_RETURNCMD|TPM_RIGHTALIGN|TPM_BOTTOMALIGN|TPM_NONOTIFY, rect.right, rect.top, this);
	switch (nPos)
	{
		case 1:
			OnExport();
			break;
		case 2:
			OnExportDetail();
			break;
	}
}

void CFMSudungthuocBHYT20a_Y2015::OnExport(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp;
	CString szOldGroup, szNewGroup, szAmount, szOrderBy, szOldLine, szNewLine;
	int nCol = 0, nRow = 0, nItem = 0, nQty = 0;
	double nCost = 0;
	long double nGrpCost = 0, nLineCost = 0, nTotalCost = 0;

	BeginWaitCursor();

	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}


	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(2, 25);
	xls.SetColumnWidth(3, 25);
	xls.SetColumnWidth(4, 25);
	xls.SetColumnWidth(6, 10);
	xls.SetColumnWidth(11, 15);

	//xls.SetRowHeight(5, 40);

	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);

	xls.SetCellMergedColumns(nCol + 5, nRow, 4);
	xls.SetCellMergedColumns(nCol + 5, nRow + 1, 4);

	xls.SetCellMergedColumns(nCol, nRow + 2, 9);
	xls.SetCellMergedColumns(nCol, nRow + 3, 9);

	/*xls.SetCellText(nCol, nRow, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);*/
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol + 5, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 5, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol, nRow + 2, _T("\x42\xC1O \x43\xC1O S\x1EEC \x44\x1EE4NG THU\x1ED0\x43 TH\x45O QU\xDD (M\x1EAAU \x42HYT 20A)"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);
	nRow = 4;

	if (!m_szTypeKey.IsEmpty())
	{
		tmpStr.Format(_T("Lo\x1EA1i: %s"), m_wndType.GetCurrent(1));
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT, true, 10);
		nRow++;
	}
	if (!m_szGroupKey.IsEmpty())
	{
		tmpStr.Format(_T("Nh\xF3m: %s"), m_wndGroup.GetCurrent(1));
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT, true, 10);
		nRow++;
	}
	/*if(!m_szSourceKey.IsEmpty())
	{
		tmpStr.Format(_T("Nguon mua: %s"), m_wndSource.GetCurrent(1));
		xls.SetCellText(0, nRow++, tmpStr,FMT_TEXT,true,12,0);
	}*/	
	/*xls.SetCellText(0, 7,_T("Kho thuoc") + m_szStockName);
	xls.SetCellText(0, 8, _T("Khoa phong") +m_szFacList);
	xls.SetCellText(2, 9,_T("BHYT:") + m_szInsFlag);
	xls.SetCellText(4, 9,_T("Tre em:")+m_szChildFlag);
	xls.SetCellText(5, 9,_T("Vien phi:") +m_szFreeFlag);
	xls.SetCellText(6, 9,_T("Dich vu:")+m_szServiceFlag);*/
	xls.SetCellMergedRows(0, nRow, 2);
	xls.SetCellMergedRows(1, nRow, 2);
	xls.SetCellMergedRows(2, nRow, 2);
	xls.SetCellMergedRows(3, nRow, 2);
	xls.SetCellMergedRows(4, nRow, 2);
	xls.SetCellMergedRows(5, nRow, 2);
	xls.SetCellMergedRows(6, nRow, 2);
	xls.SetCellMergedRows(7, nRow, 2);
	xls.SetCellMergedRows(10, nRow, 2);
	xls.SetCellMergedRows(11, nRow, 2);
	xls.SetCellMergedColumns(8, nRow, 2);
	xls.SetCellText(0, nRow, _T("STT"), 4098, true);
	xls.SetCellText(1, nRow, _T("STT th\x65o \x44MT \x63\x1EE7\x61 \x42YT"), 4098, true);
	xls.SetCellText(2, nRow, _T("T\xEAn ho\x1EA1t \x63h\x1EA5t"), 4098, true);
	xls.SetCellText(3, nRow, _T("T\xEAn thu\x1ED1\x63 th\xE0nh ph\x1EA9m"), 4098, true);
	xls.SetCellText(4, nRow, _T("\x110\x1B0\x1EDDng \x64\xF9ng, \x64\x1EA1ng \x62\xE0o \x63h\x1EBF"), 4098, true);
	xls.SetCellText(5, nRow, _T("H\xE0m l\x1B0\x1EE3ng/n\x1ED3ng \x111\x1ED9"), 4098, true);
	xls.SetCellText(6, nRow, _T("S\x1ED1 \x111\x103ng k\xFD ho\x1EB7\x63 s\x1ED1 GPNK"), 4098, true);
	xls.SetCellText(7, nRow, _T("\x110\x1A1n v\x1ECB"), 4098, true);
	xls.SetCellText(8, nRow, _T("S\x1ED1 l\x1B0\x1EE3ng"), 4098, true);
	xls.SetCellText(8, nRow + 1, _T("Ngo\x1EA1i tr\xFA"), 4098, true);
	xls.SetCellText(9, nRow + 1, _T("N\x1ED9i tr\xFA"), 4098, true);
	xls.SetCellText(10, nRow, _T("\x110\x1A1n gi\xE1(\x111\x1ED3ng)"), 4098, true);
	xls.SetCellText(11, nRow, _T("Th\xE0nh ti\x1EC1n"), 4098, true);
	nRow = nRow + 2;
	if (m_bOrderBy)
		szOrderBy = _T("groupname");
	//else
	//	szOrderBy = _T("typename");

	while(!rs.IsEOF())
	{
		rs.GetValue(_T("patient_line"), szNewLine);
		rs.GetValue(szOrderBy, szNewGroup);
		if (szNewLine != szOldLine)
		{
			if (nGrpCost > 0)
			{
				TranslateString(_T("\x43\x1ED9ng:"), szAmount);												
				xls.SetCellMergedColumns(1, nRow, 10);
				xls.SetCellText(1, nRow, szAmount, FMT_TEXT, true, 11, 0);
				tmpStr.Format(_T("%.2lf"), nGrpCost);
				xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1, true, 11, 0);
				nLineCost += nGrpCost;
				nGrpCost = 0;	
				nRow++;
			}
			if (nLineCost > 0)
			{
				TranslateString(_T("\x43\x1ED9ng:"), szAmount);
				tmpStr.Format(_T("%s %s"), szAmount, szOldLine);
				xls.SetCellMergedColumns(1, nRow, 10);
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true, 11, 0);
				tmpStr.Format(_T("%.2lf"), nLineCost);
				xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1, true, 11, 0);
				nTotalCost += nLineCost;
				nLineCost = 0;	
				nRow++;
			}
			xls.SetCellMergedColumns(0, nRow, 12);	
			m_mapLine.Lookup(szNewLine, szTemp);
			tmpStr.Format(_T("%s. %s"), szNewLine, szTemp);
			xls.SetCellText(0, nRow, tmpStr, FMT_TEXT, true, 11);
			szOldLine = szNewLine;
			szOldGroup.Empty();
			nItem = 1;
			nRow++;
		}
		if (szNewGroup != szOldGroup && !szNewGroup.IsEmpty())
		{	
			if (nGrpCost > 0)
			{			
				TranslateString(_T("\x43\x1ED9ng:"), szAmount);												
				xls.SetCellMergedColumns(1, nRow, 10);
				xls.SetCellText(1, nRow, szAmount, FMT_TEXT, true, 11, 0);
				tmpStr.Format(_T("%.2lf"), nGrpCost);
				xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1, true, 11, 0);
				nTotalCost += nGrpCost;
				nGrpCost = 0;								
				nRow++;
			}

			xls.SetCellMergedColumns(0, nRow, 12);			
			xls.SetCellText(0, nRow, szNewGroup, FMT_TEXT, true, 11);
			szOldGroup = szNewGroup;
			nRow++;
		}
		tmpStr.Format(_T("%d"), nItem++);
		xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);
		
		rs.GetValue(_T("genericname"), tmpStr);
		xls.SetCellText(2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("drugname"), tmpStr);
		xls.SetCellText(3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("specification"), tmpStr);
		xls.SetCellText(4, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("dosage"), tmpStr);
		xls.SetCellText(5, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("serial"), tmpStr);
		xls.SetCellText(6, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("unit"), tmpStr);		
		xls.SetCellText(7, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("exam_qty"), tmpStr);
		xls.SetCellText(8, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("inward_qty"), tmpStr);
		xls.SetCellText(9, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("price"), tmpStr);				
		xls.SetCellText(10, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("amount"), nCost);				
		tmpStr.Format(_T("%.2f"), nCost);
		nGrpCost += nCost;
		xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);
		nRow++;
		rs.MoveNext();
	}
	
	if (nGrpCost > 0)
	{
		TranslateString(_T("\x43\x1ED9ng:"), tmpStr);		
		xls.SetCellMergedColumns(1, nRow, 10);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true, 11 ,0);
		tmpStr.Format(_T("%.2lf"), nGrpCost);
		xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1, true, 11, 0);
		nLineCost += nGrpCost;
		nRow++;
	}

	if (nLineCost > 0)
	{
		TranslateString(_T("\x43\x1ED9ng:"), szAmount);		
		xls.SetCellMergedColumns(1, nRow, 10);
		tmpStr.Format(_T("%s %s"), szAmount, szNewLine);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true, 11 ,0);
		tmpStr.Format(_T("%.2lf"), nGrpCost);
		xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1, true, 11, 0);
		nTotalCost += nLineCost;
		nRow++;
	}
	if (nTotalCost > 0)
	{
		TranslateString(_T("Total Amount"), tmpStr);		
		xls.SetCellMergedColumns(0, nRow, 10);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT, true, 11);
		tmpStr.Format(_T("%.2lf"), nTotalCost);
		xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1, true, 11);
	}

	xls.Save(_T("Exports\\SudungthuocBHYT20a.xls"));
	EndWaitCursor();

}

void CFMSudungthuocBHYT20a_Y2015::OnExportDetail(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CExcel xls;
	CString szSQL, tmpStr, szOldType, szNewType, szOldLine, szNewLine;
	CStringArray arrCol;
	double nTmp = 0, nTypeAmt = 0, nTypeEQty = 0, nTypeIQty = 0, nLineAmt = 0, nLineEQty = 0, nLineIQty = 0;
	long double nTotal  = 0;
	int nIdx = 1, nRow = 0, nIdx1 = 1, nOldLine = 0;
	szSQL = GetQueryString1();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(1, 20);
	xls.SetColumnWidth(2, 35);
	xls.SetColumnWidth(4, 20);
	xls.SetColumnWidth(7, 15);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 4);
	xls.SetCellMergedColumns(0, 3, 4);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr = _T("\x43HI TI\x1EBET S\x1EEC \x44\x1EE4NG THU\x1ED0\x43 \x42HYT \x32\x30\x41");
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	xls.SetCellMergedColumns(6, nRow, 2);
	xls.SetCellText(0, 4, _T("STT"), 4098, true);
	xls.SetCellText(1, 4, _T("T\xEAn ho\x1EA1t \x63h\x1EA5t"), 4098, true);
	xls.SetCellText(2, 4, _T("T\xEAn thu\x1ED1\x63 - h\xE0m l\x1B0\x1EE3ng"), 4098, true);
	xls.SetCellText(3, 4, _T("\x110\x1A1n v\x1ECB"), 4098, true);
	xls.SetCellText(4, 4, _T("N\x1B0\x1EDB\x63 S\x58"), 4098, true);
	xls.SetCellText(5, 4, _T("\x110\x1A1n gi\xE1"), 4098, true);
	xls.SetCellText(6, 4, _T("S\x1ED1 l\x1B0\x1EE3ng"), 4098, true);
	xls.SetCellText(6, 5, _T("Ngo\x1EA1i tr\xFA"), 4098, true);
	xls.SetCellText(7, 5, _T("N\x1ED9i tr\xFA"), 4098, true);
	xls.SetCellText(8, 4, _T("Th\xE0nh ti\x1EC1n"), 4098, true);
	xls.SetCellText(9, 4, _T("S\x44K"), 4098, true);
	xls.SetCellText(10, 4, _T("\x42\xE1\x63 s\x1EF9"), 4098, true);

	nRow = 6;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("product_producttype"), szNewType);
		if (szNewType != szOldType)
		{
			if (nLineEQty + nLineIQty > 0)
			{
				xls.SetCellText(6, nOldLine, double2str(nLineEQty), FMT_NUMBER1, true);
				xls.SetCellText(7, nOldLine, double2str(nLineIQty), FMT_NUMBER1, true);
				xls.SetCellText(8, nOldLine, double2str(nLineAmt), FMT_NUMBER1, true);
				nTypeAmt += nLineAmt;
				nTypeEQty += nLineEQty;
				nTypeIQty += nLineIQty;
				nLineAmt = 0;
				nLineEQty = nLineIQty = 0;
			}
			if (nTypeEQty + nTypeIQty > 0)
			{
				xls.SetCellText(5, nRow, _T("\x43\x1ED9ng"), 4098, true, 11);
				xls.SetCellText(6, nRow, double2str(nTypeEQty), FMT_NUMBER1, true, 11);
				xls.SetCellText(7, nRow, double2str(nTypeIQty), FMT_NUMBER1, true, 11);
				xls.SetCellText(8, nRow, double2str(nTypeAmt), FMT_NUMBER1, true, 11);
				nTotal += nTypeAmt;
				nTypeAmt = 0;
				nTypeEQty = nTypeIQty = 0;
				nRow++;
			}
			xls.SetCellText(0, nRow, rs.GetValue(_T("product_typename")), FMT_TEXT, true, 11);
			szOldType = szNewType;
			szOldLine.Empty();
			nRow++;
		}
		rs.GetValue(_T("product_id"), szNewLine);
		if (szNewLine != szOldLine)
		{
			if (nLineEQty + nLineIQty > 0)
			{
				xls.SetCellText(6, nOldLine, double2str(nLineEQty), FMT_NUMBER1, true);
				xls.SetCellText(7, nOldLine, double2str(nLineIQty), FMT_NUMBER1, true);
				xls.SetCellText(8, nOldLine, double2str(nLineAmt), FMT_NUMBER1, true);
				nTypeAmt += nLineAmt;
				nTypeEQty += nLineEQty;
				nTypeIQty += nLineIQty;
				nLineAmt = 0;
				nLineEQty = nLineIQty = 0;
			}
			xls.SetCellText(0, nRow, int2str(nIdx1++), FMT_TEXT | FMT_RIGHT, true);
			xls.SetCellText(1, nRow, rs.GetValue(_T("product_classname")), FMT_TEXT, true);
			xls.SetCellText(2, nRow, rs.GetValue(_T("product_name")), FMT_TEXT, true);
			xls.SetCellText(3, nRow, rs.GetValue(_T("product_uomname")), FMT_TEXT, true);
			xls.SetCellText(4, nRow, rs.GetValue(_T("product_countryname")), FMT_TEXT, true);
			xls.SetCellText(5, nRow, rs.GetValue(_T("ins_price")), FMT_NUMBER1, true);
			nOldLine = nRow;
			szOldLine = szNewLine;
			nIdx = 1;
			nRow++;
		}
		xls.SetCellText(1, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		tmpStr.Format(_T("[%s] %s - %s"), rs.GetValue(_T("doc_no")), rs.GetValue(_T("patient_name")), rs.GetValue(_T("age")));
		xls.SetCellText(2, nRow, tmpStr, FMT_TEXT);
		xls.SetCellText(4, nRow, rs.GetValue(_T("card_no")), FMT_TEXT);
		rs.GetValue(_T("exam_qty"), nTmp);
		nLineEQty += nTmp;
		xls.SetCellText(6, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("inward_qty"), nTmp);
		nLineIQty += nTmp;
		xls.SetCellText(7, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("amount"), nTmp);
		nLineAmt += nTmp;
		xls.SetCellText(8, nRow, double2str(nTmp), FMT_NUMBER1);
		xls.SetCellText(10, nRow, rs.GetValue(_T("doctor")), FMT_TEXT);
		nRow++;
		rs.MoveNext();
	}
	if (nLineEQty + nLineIQty > 0)
	{
		xls.SetCellText(6, nOldLine, double2str(nLineEQty), FMT_NUMBER1, true);
		xls.SetCellText(7, nOldLine, double2str(nLineIQty), FMT_NUMBER1, true);
		xls.SetCellText(8, nOldLine, double2str(nLineAmt), FMT_NUMBER1, true);
		nTypeAmt += nLineAmt;
		nTypeEQty += nLineEQty;
		nTypeIQty += nLineIQty;
	}
	if (nTypeEQty + nTypeIQty > 0)
	{
		xls.SetCellText(5, nRow, _T("\x43\x1ED9ng"), 4098, true, 11);
		xls.SetCellText(6, nRow, double2str(nTypeEQty), FMT_NUMBER1, true, 11);
		xls.SetCellText(7, nRow, double2str(nTypeIQty), FMT_NUMBER1, true, 11);
		xls.SetCellText(8, nRow, double2str(nTypeAmt), FMT_NUMBER1, true, 11);
		nTotal += nTypeAmt;
		nRow++;
	}
	if (nTotal > 0)
	{
		xls.SetCellText(5, nRow, _T("T\x1ED5ng \x43\x1ED9ng"), 4098, true, 12);
		tmpStr.Format(_T("%.2f"), nTotal);
		xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1, true, 12);
	}
	xls.Save(_T("Exports\\Chitiet20A.xls"));	
}

void CFMSudungthuocBHYT20a_Y2015::OnListDblClick(){
	
} 
void CFMSudungthuocBHYT20a_Y2015::OnListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMSudungthuocBHYT20a_Y2015::OnListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMSudungthuocBHYT20a_Y2015::OnListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	m_wndList.BeginLoad(); 
	m_wndList.DeleteAllItems(); 
	int nCount = 0;

	if (m_nExam == 0)
		szWhere.Format(_T(" AND sd_type IN('KB') "));

	if (m_nInPatient == 0)
		szWhere.Format(_T(" AND sd_type IN('DT') "));

	if (m_nAllHospitals == 0)
		szWhere.Format(_T(" AND sd_type IN('KB','DT') "));

	szSQL.Format(_T("SELECT sd_id AS ID, sd_name AS Name ") \
		         _T("FROM sys_dept ") \
				 _T("WHERE 1=1 %s ") \
				 _T("ORDER BY sd_id"), szWhere);
	//MessageBox(szSQL);

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF())
	{ 
		m_wndList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
}
int CFMSudungthuocBHYT20a_Y2015::OnAddFMSudungthuocBHYT20a(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMSudungthuocBHYT20a_Y2015::OnEditFMSudungthuocBHYT20a(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMSudungthuocBHYT20a_Y2015::OnDeleteFMSudungthuocBHYT20a(){
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
 		OnCancelFMSudungthuocBHYT20a(); 
 	}
	return 0;
}
int CFMSudungthuocBHYT20a_Y2015::OnSaveFMSudungthuocBHYT20a(){
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
 		//OnFMSudungthuocBHYT20aListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMSudungthuocBHYT20a_Y2015::OnCancelFMSudungthuocBHYT20a(){
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
int CFMSudungthuocBHYT20a_Y2015::OnFMSudungthuocBHYT20aListLoadData(){
	return 0;
}

CString CFMSudungthuocBHYT20a_Y2015::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString szSQL, szWhere, szOrderBy;
	CString szDept, szFeeCondition;

	if (!m_szStockKey.IsEmpty())
		szWhere.AppendFormat(_T(" and storage_id=%d "), ToInt(m_szStockKey));


	if (!m_szTypeKey.IsEmpty())
		szWhere.AppendFormat(_T(" and product_producttype='%s'"), m_szTypeKey);


	if (!m_szGroupKey.IsEmpty())
		szWhere.AppendFormat(_T(" and product_categoryid='%s'"), m_wndGroup.GetCurrent(2));


	if (m_bDrugPTTT)
		szWhere.AppendFormat(_T(" and fee_type <> 'PT'"));
	//else
		//szWhere.AppendFormat(_T(" and hpo_ordertype in('P','D','M','E') "));

	if (m_bOrderBy) 
		szOrderBy.Format(_T("groupname,"));
	//else
	//	szOrderBy.Format(_T("typename"));

	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (!szDept.IsEmpty())
				szDept += _T(",");
			szDept.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
		}
	}

	if (!szDept.IsEmpty())
		szWhere.AppendFormat(_T(" and dept_id in(%s) "), szDept);
	szFeeCondition.Format(_T(" AND invoice_status = 'P' ") \
						_T("  AND invoice_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss')"), 
						m_szFromDate, m_szToDate);
	if (m_bByDischargedDate)
		szFeeCondition.Format(_T(" AND fee_status IN ('Y', 'P')") \
								_T(" AND fee_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss')"),
								m_szFromDate, m_szToDate);
	if (m_bOnlyCommander)
		szWhere.AppendFormat(_T(" AND doc_rank IN (15, 16, 17, 18, 21, 22, 23, 24)"));
	szSQL.Format(_T(" SELECT patient_line, ") \
				_T("        idx, ") \
				_T("        typename, ") \
				_T("        code, ") \
				_T("        dosage, ") \
				_T("        serial, ") \
				_T("        drugname, ") \
				_T("        lvl, ") \
				_T("        genericname, ") \
				_T("        groupname, ") \
				_T("        unit, ") \
				_T("        countryname, ") \
				_T("        specification, ") \
				_T("        price, ") \
				_T("        SUM(inward_qty) inward_qty, ") \
				_T("        SUM(exam_qty) exam_qty, ") \
				_T("        SUM(amount) amount ") \
				_T(" FROM   (SELECT    Hms_getpatientline(doc_no) patient_line, ") \
				_T("                   hfg_index AS idx, ") \
				_T("                   Get_producttypename(product_producttype) AS typename, ") \
				_T("                   product_code AS code, ") \
				_T("                   product_name2 AS dosage, ") \
				_T("                   product_serial AS serial, ") \
				_T("                   product_name AS drugname, ") \
				_T("                   product_priority AS lvl, ") \
				_T("                   product_classname AS genericname, ") \
				_T("                   product_category ") \
				_T("                   ||'> ' ") \
				_T("                   || product_category AS groupname, ") \
				_T("                   product_uomname AS unit, ") \
				_T("                   product_countryname AS countryname, ") \
				_T("                   product_specification specification, ") \
				_T("                   ins_price AS price, ") \
				_T("                   inward_qty AS inward_qty, ") \
				_T("                   exam_qty exam_qty, ") \
				_T("                   qty * ins_price AS amount ") \
				_T("         FROM      (SELECT    hpo_docno doc_no, ") \
				_T("                              fi.hfe_deptid dept_id, ") \
				_T("                              hpo_storage_id storage_id, ") \
				_T("                              hpol_product_id, ") \
				_T("                              CASE WHEN fi.hfe_class = 'I' THEN hcr_status ") \
				_T("                              ELSE hd_status ") \
				_T("                              END doc_status, ") \
				_T("                              CASE WHEN fi.hfe_class = 'I' THEN hcrf_acceptedfee ") \
				_T("                              ELSE hdf_acceptedfee ") \
				_T("                              END fee_status, ") \
				_T("                              CASE WHEN fi.hfe_class = 'I' THEN hcr_dischargedate ") \
				_T("                              ELSE fi.hfe_date ") \
				_T("                              END fee_date, ") \
				_T("                              fi.hfe_date invoice_date, ") \
				_T("                              fi.hfe_status invoice_status, ") \
				_T("                              hd_rank doc_rank, ") \
				_T("                              CASE WHEN fi.hfe_class = 'I' THEN hfe_quantity ") \
				_T("                              ELSE 0 ") \
				_T("                              END inward_qty, ") \
				_T("                              CASE WHEN fi.hfe_class <> 'I' THEN hfe_quantity ") \
				_T("                              ELSE 0 ") \
				_T("                              END exam_qty, ") \
				_T("                              Nvl(hpo_feetype, 'X') fee_type, ") \
				_T("                              hfe_quantity qty, ") \
				_T("                              fe.hfe_insprice ins_price ") \
				_T("                    FROM      hms_pharmaorder_view ") \
				_T("                    LEFT JOIN hms_pharmaorderline_view2 ON( hpol_orderid = hpo_orderid ) ") \
				_T("                    LEFT JOIN hms_doc ON( hd_docno = hpo_docno ) ") \
				_T("                    LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno ) ") \
				_T("                    LEFT JOIN hms_treatment_record ON ( htr_docno = hcr_docno AND hcr_refidx = htr_idx ) ") \
				_T("                    LEFT JOIN hms_fee_view fe ON( hpol_orderid = fe.hfe_orderid AND hpol_orderlineid = fe.hfe_orderline AND fe.hfe_docno = hpol_docno AND hfe_type IN ( 'D', 'M' ) ) ") \
				_T("                    LEFT JOIN hms_fee_invoice fi ON( fi.hfe_docno = hd_docno AND fi.hfe_invoiceno = fe.hfe_invoiceno ) ") \
				_T("                    LEFT JOIN hms_object ON( fi.hfe_objectid = ho_id ) ") \
				_T("                    WHERE     ho_type IN ( 'I', 'C' ) AND fe.hfe_discount > 0 AND Nvl(htr_outpatient, 'X') <> 'Y' ") \
				_T("						AND fe.hfe_category <> 'Y' AND fe.hfe_status <> 'C') ") \
				_T("         LEFT JOIN m_product_view ON ( product_id = hpol_product_id ) ") \
				_T("         LEFT JOIN hms_fee_group ON( hfg_id = product_producttype ) ") \
				_T("         WHERE     product_org_id IN ( 'PM' ) AND product_producttype NOT IN ( 'A1500', 'A1600' ) AND doc_status = 'T' ") \
				_T("			%s %s) ") \
				_T(" GROUP  BY patient_line,idx,typename,code,dosage,serial,drugname,lvl,genericname,groupname,unit,countryname,specification,price ") \
				_T(" ORDER  BY patient_line,idx,drugname,unit"), szWhere, szFeeCondition);
	
_fmsg(_T("%s"), szSQL);
	return szSQL;
}

CString CFMSudungthuocBHYT20a_Y2015::GetQueryString1(){
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString szSQL, szWhere, szOrderBy;
	CString szDept, szFeeCondition;

	if (!m_szStockKey.IsEmpty())
		szWhere.AppendFormat(_T(" and storage_id=%d "), ToInt(m_szStockKey));


	if (!m_szTypeKey.IsEmpty())
		szWhere.AppendFormat(_T(" and product_producttype='%s'"), m_szTypeKey);


	if (!m_szGroupKey.IsEmpty())
		szWhere.AppendFormat(_T(" and product_categoryid='%s'"), m_wndGroup.GetCurrent(2));


	if (m_bDrugPTTT)
		szWhere.AppendFormat(_T(" and fee_type <> 'PT'"));

	if (m_bOrderBy) 
		szOrderBy.Format(_T("groupname"));
	//else
	//	szOrderBy.Format(_T("typename"));

	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (!szDept.IsEmpty())
				szDept += _T(",");
			szDept.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
		}
	}

	if (!szDept.IsEmpty())
		szWhere.AppendFormat(_T(" and dept_id in(%s) "), szDept);
	szFeeCondition.Format(_T(" AND invoice_status = 'P' ") \
						_T("  AND invoice_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss')"), 
						m_szFromDate, m_szToDate);
	if (m_bByDischargedDate)
		szFeeCondition.Format(_T(" AND fee_status IN ('Y', 'P') ") \
								_T("AND end_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss')"), 
								m_szFromDate, m_szToDate);
	if (m_bOnlyCommander)
		szWhere.AppendFormat(_T(" AND doc_rank IN (15, 16, 17, 18, 21, 22, 23, 24)"));
	szSQL.Format(_T(" SELECT    product_producttype, ") \
				_T("           Get_username(doctor) doctor, ") \
				_T("           product_typename, ") \
				_T("           product_id, ") \
				_T("           product_classname, ") \
				_T("           product_name, ") \
				_T("           product_uomname, ") \
				_T("           product_countryname, ") \
				_T("           doc_no, ") \
				_T("           Get_patientname(doc_no) patient_name, ") \
				_T("           Hms_getagebydoc(doc_no) age, ") \
				_T("           card_no, ") \
				_T("           ins_price, ") \
				_T("           SUM(inward_qty) inward_qty, ") \
				_T("           SUM(exam_qty) exam_qty, ") \
				_T("           SUM(qty) qty, ") \
				_T("           SUM(qty * ins_price) amount ") \
				_T(" FROM      (SELECT    hfe_itemid, ") \
				_T("                      fi.hfe_docno doc_no, ") \
				_T("                      hd_cardno card_no, ") \
				_T("                      hfe_insprice ins_price, ") \
				_T("                      hpo_doctor doctor, ") \
				_T("                      hpo_storage_id storage_id, ") \
				_T("                      CASE WHEN fi.hfe_class = 'I' THEN fe.hfe_quantity ") \
				_T("                      ELSE 0 ") \
				_T("                      END inward_qty, ") \
				_T("                      CASE WHEN fi.hfe_class IN ( 'E', 'S' ) THEN fe.hfe_quantity ") \
				_T("                      ELSE 0 ") \
				_T("                      END exam_qty, ") \
				_T("                      hfe_quantity qty, ") \
				_T("                      CASE WHEN fi.hfe_class = 'I' THEN hcrf_acceptedfee ") \
				_T("                      ELSE hdf_acceptedfee ") \
				_T("                      END fee_status, ") \
				_T("                      fi.hfe_status invoice_status, ") \
				_T("                      fi.hfe_deptid dept_id, ") \
				_T("                      fi.hfe_date invoice_date, ") \
				_T("                      Nvl(hpo_feetype, 'X') fee_type, ") \
				_T("                      hd_rank doc_rank, ") \
				_T("                      CASE WHEN fi.hfe_class = 'I' THEN hcr_status ") \
				_T("                      ELSE hd_status ") \
				_T("                      END doc_status, ") \
				_T("                      CASE WHEN fi.hfe_class = 'I' THEN hcr_dischargedate ") \
				_T("                      ELSE fi.hfe_date ") \
				_T("                      END end_date ") \
				_T("            FROM      hms_fee fe ") \
				_T("            LEFT JOIN hms_fee_invoice fi ON ( fi.hfe_docno = fe.hfe_docno AND fi.hfe_invoiceno = fe.hfe_invoiceno ) ") \
				_T("            LEFT JOIN hms_doc ON ( fi.hfe_docno = hd_docno ) ") \
				_T("            LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno ) ") \
				_T("            LEFT JOIN hms_treatment_record ON ( hcr_docno = htr_docno AND hcr_refidx = htr_idx ) ") \
				_T("            LEFT JOIN hms_object ON ( fi.hfe_objectid = ho_id ) ") \
				_T("            LEFT JOIN hms_pharmaorder_view ON ( hpo_orderid = fe.hfe_orderid AND hpo_docno = fe.hfe_docno ) ") \
				_T("            WHERE     ho_type IN ( 'I', 'C' ) AND fe.hfe_type IN ( 'D', 'M' ) AND fe.hfe_status <> 'C' ") \
				_T("				AND fe.hfe_discount > 0 AND Nvl(htr_outpatient, 'X') <> 'Y') tbl0 ") \
				_T(" LEFT JOIN m_productitem_view ON ( hfe_itemid = product_item_id ) ") \
				_T(" WHERE     product_org_id IN ( 'PM' ) AND product_producttype NOT IN ( 'A1500', 'A1600' ) ") \
				_T("	AND fee_status IN ('Y', 'P') AND doc_status = 'T' ") \
				_T("    %s %s")
				_T(" GROUP     BY product_producttype,product_typename,product_id,product_classname,product_countryname,product_name,doc_no,card_no,ins_price,product_uomname,doctor ") \
				_T(" ORDER     BY product_producttype,product_name,doc_no"), szWhere, szFeeCondition);

_fmsg(_T("%s"), szSQL);
	return szSQL;
}

int CFMSudungthuocBHYT20a_Y2015::OnListCheckAll()
{
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (!m_wndList.GetCheck(i))
		{
			m_wndList.SetCheck(i, TRUE);
		}
	}
	return 0;
}

int CFMSudungthuocBHYT20a_Y2015::OnListUnCheckAll()
{
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			m_wndList.SetCheck(i, FALSE);
		}
	}
	return 0;
}
