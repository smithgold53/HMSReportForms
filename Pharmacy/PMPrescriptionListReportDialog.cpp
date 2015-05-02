#include "stdafx.h"
#include "PMPrescriptionListReportDialog.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include "MainFrame_E10.h"
#include ".\pmprescriptionlistreportdialog.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSPrescriptionListReportDialog *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSPrescriptionListReportDialog *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSPrescriptionListReportDialog* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CPMSPrescriptionListReportDialog *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnStockAddNew();
}*/
static void _OnDoctorSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSPrescriptionListReportDialog* )pWnd)->OnDoctorSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDoctorSelendokFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnDoctorSelendok();
}
/*static void _OnDoctorSetfocusFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnDoctorKillfocus();
}*/
/*static void _OnDoctorKillfocusFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnDoctorKillfocus();
}*/
static long _OnDoctorLoadDataFnc(CWnd *pWnd){
	return ((CPMSPrescriptionListReportDialog *)pWnd)->OnDoctorLoadData();
}
/*static void _OnDoctorAddNewFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnDoctorAddNew();
}*/
static long _OnApprovedByLoadDataFnc(CWnd *pWnd){
	return ((CPMSPrescriptionListReportDialog *)pWnd)->OnApprovedByLoadData();
}
static int _OnMinTotalAmtCheckValueFnc(CWnd *pWnd){
	return ((CPMSPrescriptionListReportDialog *)pWnd)->OnMinTotalAmtCheckValue();
}
static void _OnCheckAllSelectFnc(CWnd *pWnd){
	 ((CPMSPrescriptionListReportDialog*)pWnd)->OnCheckAllSelect();
}
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	 ((CPMSPrescriptionListReportDialog*)pWnd)->OnOutPatientSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	 ((CPMSPrescriptionListReportDialog*)pWnd)->OnInPatientSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSPrescriptionListReportDialog *pVw = (CPMSPrescriptionListReportDialog *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSPrescriptionListReportDialog *pVw = (CPMSPrescriptionListReportDialog *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CPMSPrescriptionListReportDialog *pVw = (CPMSPrescriptionListReportDialog *)pWnd;
	pVw->OnCloseSelect();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSPrescriptionListReportDialog* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CPMSPrescriptionListReportDialog *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog *)pWnd)->OnObjectAddNew();
}*/
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CPMSPrescriptionListReportDialog*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CPMSPrescriptionListReportDialog*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPMSPrescriptionListReportDialog*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListCheckC11RoomFnc(CWnd *pWnd){
	return ((CPMSPrescriptionListReportDialog*)pWnd)->OnListCheckC11Room();
} 
static int _OnListCheckC12RoomFnc(CWnd *pWnd){
	return ((CPMSPrescriptionListReportDialog*)pWnd)->OnListCheckC12Room();
} 
static int _OnListCheckC13RoomFnc(CWnd *pWnd){
	return ((CPMSPrescriptionListReportDialog*)pWnd)->OnListCheckC13Room();
} 
static int _OnAddPMSPrescriptionListReportDialogFnc(CWnd *pWnd){
	 return ((CPMSPrescriptionListReportDialog*)pWnd)->OnAddPMSPrescriptionListReportDialog();
} 
static int _OnEditPMSPrescriptionListReportDialogFnc(CWnd *pWnd){
	 return ((CPMSPrescriptionListReportDialog*)pWnd)->OnEditPMSPrescriptionListReportDialog();
}
static int _OnListUncheckAllFnc(CWnd *pWnd){
	return ((CPMSPrescriptionListReportDialog*)pWnd)->OnListUncheckAll();
} 
static int _OnDeletePMSPrescriptionListReportDialogFnc(CWnd *pWnd){
	 return ((CPMSPrescriptionListReportDialog*)pWnd)->OnDeletePMSPrescriptionListReportDialog();
} 
static int _OnSavePMSPrescriptionListReportDialogFnc(CWnd *pWnd){
	 return ((CPMSPrescriptionListReportDialog*)pWnd)->OnSavePMSPrescriptionListReportDialog();
} 
static int _OnCancelPMSPrescriptionListReportDialogFnc(CWnd *pWnd){
	 return ((CPMSPrescriptionListReportDialog*)pWnd)->OnCancelPMSPrescriptionListReportDialog();
} 
CPMSPrescriptionListReportDialog::CPMSPrescriptionListReportDialog(CWnd *pParent)
{
	m_nDlgWidth = 590;
	m_nDlgHeight = 495;
	SetDefaultValues();
}
CPMSPrescriptionListReportDialog::~CPMSPrescriptionListReportDialog(){
}
void CPMSPrescriptionListReportDialog::OnCreateComponents(){
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 110, 55);
	m_wndFromDate.Create(this,115, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 395, 55);
	m_wndToDate.Create(this,400, 30, 570, 55); 
	m_wndItemType.Create(this, _T("Item Type"), 295, 185, 575, 240);
	m_wndDeptType.Create(this, _T("Dept Type"), 5, 185, 290, 240);
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 110, 85);
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 180);
	m_wndStock.Create(this,115, 60, 570, 85); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 89, 110, 114);
	m_wndObject.Create(this,115, 89, 570, 114); 
	m_wndDoctorLabel.Create(this, _T("Doctor"), 10, 120, 110, 145);
	m_wndDoctor.Create(this,115, 120, 290, 145); 
	m_wndApprovedByLabel.Create(this, _T("Approved By"), 295, 120, 395, 145);
	m_wndApprovedBy.Create(this,400, 120, 570, 145); 
	m_wndMinTotalAmtLabel.Create(this, _T("Min Total Amt"), 10, 150, 110, 175);
	m_wndMinTotalAmt.Create(this,115, 150, 290, 175); 
	//m_wndServicePrice.Create(this, _T("Serv Price"), 295, 150, 395, 175);
	m_wndCheckAll.Create(this, _T("Check All"), 10, 210, 95, 235);
	m_wndOutPatient.Create(this, _T("OutPatient"), 100, 210, 170, 235);
	m_wndInPatient.Create(this, _T("In Patient"), 175, 210, 285, 235);
	m_wndOrientalMedicine.Create(this, _T("Oriental Medicine"), 300, 210, 430, 235);
	m_wndOccidentMedicine.Create(this, _T("Occident Medicine"), 435, 210, 570, 235);
	m_wndList.Create(this,5, 245, 575, 550); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 370, 555, 480, 580);
	m_wndExport.Create(this, _T("&Export"), 485, 555, 575, 580);

}
void CPMSPrescriptionListReportDialog::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
//	m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
//	m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
//	m_wndToDate.SetCheckValue(true);
	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);
	m_wndDoctor.SetCheckValue(true);
	m_wndDoctor.LimitText(35);
	m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(35);

	m_wndMinTotalAmt.SetCurrencyFormat(true);

	m_wndStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 60);
	m_wndStock.InsertColumn(1, _T("Description"), CFMT_TEXT, 220);

	m_wndDoctor.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDoctor.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);
	
	m_wndList.SetCheckBox(true);
	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 60);
	m_wndList.InsertColumn(1, _T("Desc"), CFMT_TEXT, 390);
	m_wndList.InsertColumn(2, _T("DeptID"), CFMT_TEXT, 0);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 20);
	m_wndObject.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndApprovedBy.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndApprovedBy.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

	m_szFromDate = pMF->GetSysDate() + _T("00:00");
	m_szToDate = pMF->GetSysDate() + _T("23:59");
	SetMode(VM_EDIT);

}
void CPMSPrescriptionListReportDialog::OnSetWindowEvents(){
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
	m_wndDoctor.SetEvent(WE_SELENDOK, _OnDoctorSelendokFnc);
	//m_wndDoctor.SetEvent(WE_SETFOCUS, _OnDoctorSetfocusFnc);
	//m_wndDoctor.SetEvent(WE_KILLFOCUS, _OnDoctorKillfocusFnc);
	m_wndDoctor.SetEvent(WE_SELCHANGE, _OnDoctorSelectChangeFnc);
	m_wndDoctor.SetEvent(WE_LOADDATA, _OnDoctorLoadDataFnc);
	//m_wndDoctor.SetEvent(WE_ADDNEW, _OnDoctorAddNewFnc);
	m_wndMinTotalAmt.SetEvent(WE_CHECKVALUE, _OnMinTotalAmtCheckValueFnc);
	m_wndCheckAll.SetEvent(WE_CLICK, _OnCheckAllSelectFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndApprovedBy.SetEvent(WE_LOADDATA, _OnApprovedByLoadDataFnc);
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);
	m_wndList.AddEvent(1, _T("Check C1.1 Room"), _OnListCheckC11RoomFnc, 0, 0, 0);
	m_wndList.AddEvent(2, _T("Check C1.2 Room"), _OnListCheckC12RoomFnc, 0, 0, 0);
	m_wndList.AddEvent(3, _T("Check C1.3 Room"), _OnListCheckC13RoomFnc, 0, 0, 0);
	m_wndList.AddEvent(4, _T("Uncheck All"), _OnListUncheckAllFnc, 0, 0, 0);
	m_wndCheckAll.SetCheck(true);
	m_szCondition = _T("A");
}
void CPMSPrescriptionListReportDialog::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_TextEx(pDX, m_wndDoctor.GetDlgCtrlID(), m_szDoctorKey);
	DDX_TextEx(pDX, m_wndApprovedBy.GetDlgCtrlID(), m_szApprovedByKey);
	DDX_Text(pDX, m_wndMinTotalAmt.GetDlgCtrlID(), m_nMinTotalAmt);
	//DDX_Check(pDX, m_wndServicePrice.GetDlgCtrlID(), m_bServPrice);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_Check(pDX, m_wndOrientalMedicine.GetDlgCtrlID(), m_bOriental);
	DDX_Check(pDX, m_wndOccidentMedicine.GetDlgCtrlID(), m_bOccident);
}
void CPMSPrescriptionListReportDialog::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPMSPrescriptionListReportDialog::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CPMSPrescriptionListReportDialog::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_szDoctorKey.Empty();
	m_szObjectKey.Empty();
	m_bOriental = FALSE;
	m_bOccident = FALSE;

}
int CPMSPrescriptionListReportDialog::SetMode(int nMode){
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
/*void CPMSPrescriptionListReportDialog::OnFromDateChange(){
	
} */
/*void CPMSPrescriptionListReportDialog::OnFromDateSetfocus(){
	
} */
/*void CPMSPrescriptionListReportDialog::OnFromDateKillfocus(){
	
} */
int CPMSPrescriptionListReportDialog::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSPrescriptionListReportDialog::OnToDateChange(){
	
} */
/*void CPMSPrescriptionListReportDialog::OnToDateSetfocus(){
	
} */
/*void CPMSPrescriptionListReportDialog::OnToDateKillfocus(){
	
} */
int CPMSPrescriptionListReportDialog::OnToDateCheckValue(){
	return 0;
} 
void CPMSPrescriptionListReportDialog::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSPrescriptionListReportDialog::OnStockSelendok(){
	 
}
/*void CPMSPrescriptionListReportDialog::OnStockSetfocus(){
	
}*/
/*void CPMSPrescriptionListReportDialog::OnStockKillfocus(){
	
}*/
long CPMSPrescriptionListReportDialog::OnStockLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageList(&m_wndStock, m_szStockKey);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStock.IsSearchKey() && ToInt(m_szStockKey) > 0){
		szWhere.Format(_T(" AND msl_storage_id=%d"), ToInt(m_szStockKey));
	};
	m_wndStock.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT msl_storage_id as id, msl_name as name FROM m_storagelist WHERE 1=1 %s ORDER BY msl_storage_id"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStock.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPMSPrescriptionListReportDialog::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSPrescriptionListReportDialog::OnDoctorSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSPrescriptionListReportDialog::OnDoctorSelendok(){
	 
}
/*void CPMSPrescriptionListReportDialog::OnDoctorSetfocus(){
	
}*/
/*void CPMSPrescriptionListReportDialog::OnDoctorKillfocus(){
	
}*/
long CPMSPrescriptionListReportDialog::OnDoctorLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();		

	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDoctor.IsSearchKey() && !m_szDoctorKey.IsEmpty()){
		szWhere.Format(_T(" WHERE su_userid='%s' "), m_szDoctorKey);
	};
	m_wndDoctor.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" SELECT su_userid as id, su_name as name FROM sys_user %s"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDoctor.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CPMSPrescriptionListReportDialog::OnDoctorAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
long CPMSPrescriptionListReportDialog::OnApprovedByLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();		
	CString szFilter;
	szFilter.Format(_T(" AND su_deptid = 'KD'"));
	return pMF->LoadUserList(&m_wndApprovedBy, m_szApprovedByKey, szFilter);

}
int CPMSPrescriptionListReportDialog::OnMinTotalAmtCheckValue(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return 0;
}

void CPMSPrescriptionListReportDialog::OnListDblClick(){
	
} 
void CPMSPrescriptionListReportDialog::OnListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CPMSPrescriptionListReportDialog::OnListCheckC11Room(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	if (m_szCondition == _T("I"))
		return 0;
	CString tmpStr;

	for (int i = 0 ; i<m_wndList.GetItemCount(); i++)
	{
		tmpStr = m_wndList.GetItemText(i, 2);
		if (tmpStr == _T("C1.1"))
			m_wndList.SetCheck(i);
	}
	return 0;
} 
int CPMSPrescriptionListReportDialog::OnListCheckC12Room(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	if (m_szCondition == _T("I"))
		return 0;
	CString tmpStr;
	for (int i = 0 ; i<m_wndList.GetItemCount(); i++)
	{
		tmpStr = m_wndList.GetItemText(i, 2);
		if (tmpStr == _T("C1.2"))
			m_wndList.SetCheck(i);
	}
	return 0;
} 
int CPMSPrescriptionListReportDialog::OnListCheckC13Room(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	if (m_szCondition == _T("I"))
		return 0;
	CString tmpStr;
	for (int i = 0 ; i < m_wndList.GetItemCount(); i++)
	{
		tmpStr = m_wndList.GetItemText(i, 2);
		if (tmpStr == _T("C1.3"))
			m_wndList.SetCheck(i);
	}
	return 0;
} 
int CPMSPrescriptionListReportDialog::OnListUncheckAll(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	if (m_szCondition == _T("I"))
		return 0;
	for (int i = 0 ; i < m_wndList.GetItemCount(); i++)
	{
		m_wndList.SetCheck(i, false);
	}
	return 0;
} 
long CPMSPrescriptionListReportDialog::OnListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndList.BeginLoad(); 
	m_wndList.DeleteAllItems(); 
	int nCount = 0;
	if (m_szCondition == _T("O"))	
		szSQL.Format(_T(" SELECT hrl_id as id, hrl_name as name, hrl_deptid as deptid FROM hms_roomlist WHERE hrl_deptid IN ('C1.1', 'C1.2', 'C1.3') ORDER BY hrl_deptid, hrl_id"));	
	if (m_szCondition == _T("I")) 
		szSQL.Format(_T(" SELECT sd_id as id, sd_name as name, 0 as deptid FROM sys_dept WHERE sd_type ='DT' ORDER BY sd_index"));
	
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndList.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), 
			rs.GetValue(_T("deptid")), NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
}

void CPMSPrescriptionListReportDialog::OnCheckAllSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_szCondition = _T("A");
	OnListLoadData();
}
void CPMSPrescriptionListReportDialog::OnOutPatientSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_szCondition = _T("O");
	OnListLoadData();
}
void CPMSPrescriptionListReportDialog::OnInPatientSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_szCondition = _T("I");
	OnListLoadData();
}
void CPMSPrescriptionListReportDialog::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CGuiMenu menu(this);
	CString tmpStr;
	
	menu.CreatePopupMenu();
	TranslateString(_T("Print Synthesis Format"), tmpStr);
	menu.AppendMenu(MF_BYPOSITION, 1, tmpStr);
	menu.AppendMenu(MF_SEPARATOR);
	TranslateString(_T("Print Detail Format"), tmpStr);
	menu.AppendMenu(MF_BYPOSITION, 2, tmpStr);
	CRect rect;
	m_wndPrintPreview.GetWindowRect(&rect);
	int nPos = menu.TrackPopupMenu(TPM_RETURNCMD|TPM_RIGHTALIGN|TPM_BOTTOMALIGN|TPM_NONOTIFY, rect.right, rect.top, this);
	switch (nPos)
	{
		case 1:
			OnPrintSynthesis();
			break;
		case 2:
			OnPrintPreview();
			break;
	}
} 
void CPMSPrescriptionListReportDialog::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString tmpStr;
	CGuiMenu menu(this);
	menu.CreatePopupMenu();
	TranslateString(_T("Export Synthesis Format"), tmpStr);
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
			OnExportSynthesis();
			break;
		case 2:
			OnExport();
			break;
	}
} 
void CPMSPrescriptionListReportDialog::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

} 
void CPMSPrescriptionListReportDialog::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSPrescriptionListReportDialog::OnObjectSelendok(){
	 
}
/*void CPMSPrescriptionListReportDialog::OnObjectSetfocus(){
	
}*/
/*void CPMSPrescriptionListReportDialog::OnObjectKillfocus(){
	
}*/
long CPMSPrescriptionListReportDialog::OnObjectLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty()){
	};
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT ho_type as id, ho_desc as name FROM hms_object"));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPMSPrescriptionListReportDialog::OnObjectAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
int CPMSPrescriptionListReportDialog::OnAddPMSPrescriptionListReportDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CPMSPrescriptionListReportDialog::OnEditPMSPrescriptionListReportDialog(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CPMSPrescriptionListReportDialog::OnDeletePMSPrescriptionListReportDialog(){
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
 		OnCancelPMSPrescriptionListReportDialog(); 
 	}
	return 0;
}
int CPMSPrescriptionListReportDialog::OnSavePMSPrescriptionListReportDialog(){
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
 		//OnPMSPrescriptionListReportDialogListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CPMSPrescriptionListReportDialog::OnCancelPMSPrescriptionListReportDialog(){
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
int CPMSPrescriptionListReportDialog::OnPMSPrescriptionListReportDialogListLoadData(){
	return 0;
}


void CPMSPrescriptionListReportDialog::OnPrintPreview()
{
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL, szObjectID, szObjectName, szWhere, szWhere1, 
		szWhere2, tmpStr2,szListID, szRoomID, szRoomName;
	if(!rpt.Init(_T("Reports/HMS/PMS_PRESCRIPTIONLISTREPORT.RPT"),true) )
	{
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);


	rpt.GetReportHeader()->SetValue(_T("StockName"), m_wndStock.GetCurrent(1));
	//rpt.GetReportHeader()->SetValue(_T("Doctor"), m_wndDoctor.GetCurrent(1));
	

	tmpStr2 = _T("' '");
	//for(int i=0; i < m_wndObject.GetCount(); i++)
	//{
	//	//if(m_wndObject.GetCheck(i))
	//	{
	//		m_wndObject.SetCurSel(i);
	//		if (!szObjectID.IsEmpty())
	//			szObjectID += _T(",");
	//		szObjectID.AppendFormat(_T("'%s'"), m_wndObject.GetCurrent(0));

	//		if (!szObjectName.IsEmpty())
	//			szObjectName += _T(",");
	//		szObjectName.AppendFormat(_T("%s"), m_wndObject.GetCurrent(1));
	//	}
	//}
	szWhere.Format(_T(" AND hpo_orderstatus = 'A' AND hpo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				   _T(" AND hpo_ordertype IN ('P', 'D', 'M', 'E')"), m_szFromDate, m_szToDate);
	if (!m_szObjectKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND ho_type in('%s') "), m_szObjectKey);
	
	CString szRoomC11, szRoomC12, szRoomC13, szDept, szProductType, szNewOrderID, szOldOrderID;
	szProductType.Empty();
	for(int i=0; i < m_wndList.GetItemCount(); i++)
	{
		if(m_wndList.GetCheck(i))
		{
			m_wndList.SetCurSel(i);
			if (!szListID.IsEmpty())
				szListID += _T(",");
			szListID.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i,0));
			szDept = m_wndList.GetItemText(i, 2);
			if (szDept == _T("C1.1"))
			{
				if (!szRoomC11.IsEmpty())
					szRoomC11 += _T(", ");
				szRoomC11.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
			}
			else if (szDept == _T("C1.2"))
			{
				if (!szRoomC12.IsEmpty())
					szRoomC12 += _T(", ");
				szRoomC12.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
			}
			else if (szDept == _T("C1.3"))
			{
				if (!szRoomC13.IsEmpty())
					szRoomC13 += _T(", ");
				szRoomC13.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
			}
			if (!szRoomName.IsEmpty())
				szRoomName += _T(", ");
			szRoomName.AppendFormat(_T("%s"), m_wndList.GetItemText(i,1));
		}
	}
	BeginWaitCursor();
	CString szDeptType, szWhereRoomID, szTable, szLineTable;
	szDeptType.Empty();
	szWhereRoomID.Empty();
	if (m_szCondition == _T("A"))
		szDeptType.IsEmpty();
	else
		if (m_szCondition == _T("O"))
		{
			szDeptType.Format(_T(" and hpo_depttype ='KB' "));
			if (!szRoomC11.IsEmpty())				
				szWhereRoomID.AppendFormat(_T(" and ((hpo_roomid in (%s) and hpo_deptid= 'C1.1') "), szRoomC11);
			if (!szRoomC12.IsEmpty())
				if (!szWhereRoomID.IsEmpty())
					szWhereRoomID.AppendFormat(_T(" or (hpo_roomid in (%s) and hpo_deptid= 'C1.2') "), szRoomC12);
				else
					szWhereRoomID.AppendFormat(_T(" and ((hpo_roomid in (%s) and hpo_deptid= 'C1.2') "), szRoomC12);
			if (!szRoomC13.IsEmpty())
				if (!szWhereRoomID.IsEmpty())
					szWhereRoomID.AppendFormat(_T(" or (hpo_roomid in (%s) and hpo_deptid= 'C1.3')"), szRoomC13);
				else
					szWhereRoomID.AppendFormat(_T(" and ((hpo_roomid in (%s) and hpo_deptid= 'C1.3')"), szRoomC13);
			if (!szWhereRoomID.IsEmpty())
				szWhereRoomID.AppendFormat(_T(")"));
		}
		else
		{
			szDeptType.Format(_T(" and hpo_depttype ='DT' "));
			if(!szListID.IsEmpty())
				szWhereRoomID.Format(_T(" and hpo_deptid in(%s) ") , szListID);
		}
	if (!szDeptType.IsEmpty())
		szWhere.AppendFormat(_T("%s"), szDeptType);

	if(!m_szDoctorKey.IsEmpty())
		szWhere.AppendFormat(_T(" and hpo_doctor='%s' "), m_szDoctorKey);

	if (!m_szStockKey.IsEmpty())
		szWhere.AppendFormat(_T(" and hpo_storage_id=%d"), ToInt(m_szStockKey));
	if (m_nMinTotalAmt > 0)
		szWhere2.Format(_T(" AND taxamount >= %ld"), m_nMinTotalAmt);
	if (!m_szApprovedByKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_approveby = '%s'"), m_szApprovedByKey);
	//szPrice = _T("product_taxprice");
	//if (m_bServPrice)
	//szPrice = _T("hpol_unitprice");
	if (m_bOriental)
		szProductType = _T(" product_producttype = 'A1301'");
	if (m_bOccident)
	{
		if (!szProductType.IsEmpty())
			szProductType += _T(" OR");
		szProductType += _T(" product_producttype <> 'A1301'");
	}
	if (!szProductType.IsEmpty())
		szWhere.AppendFormat(_T(" AND (%s)"), szProductType);
	rpt.GetReportHeader()->SetValue(_T("Object"), szObjectName);
	//Page Header
	//Report Detail
	szSQL.Format(_T(" SELECT orderid, ") \
	_T("        docno, ") \
	_T("        get_patientname(docno) pname, ") \
	_T("        Hms_getagebydoc(docno) age, ") \
	_T("        Get_username(hpo_doctor) doctor, ") \
	_T("        NAME, ") \
	_T("        Get_uomname(product_purchase_uomid) unit, ") \
	_T("        price, ") \
	_T("        SUM(qty) qty, ") \
	_T("        SUM(amount) amount ") \
	_T(" FROM   (SELECT    hpol_orderid orderid, ") \
	_T("                   hpol_orderlineid, ") \
	_T("                   hpol_qtyissue * hpol_unitprice amount, ") \
	_T("                   SUM(hpol_qtyissue * hpol_unitprice) ") \
	_T("                     OVER ( ") \
	_T("                       PARTITION BY (hpol_orderid)) taxamount, ") \
	_T("                   hpo_docno docno, ") \
	_T("                   Hms_getage(( hd_admitdate ), hp_birthdate) age, ") \
	_T("                   hpo_doctor, ") \
	_T("                   product_id, ") \
	_T("                   product_name NAME, ") \
	_T("                   product_purchase_uomid, ") \
	_T("                   hpol_unitprice price, ") \
	_T("                   hpol_qtyissue qty ") \
	_T("         FROM      hms_pharmaorder ") \
	_T("         LEFT JOIN hms_pharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
	_T("         LEFT JOIN m_productitem_view ON ( product_item_id = hpol_product_item_id ) ") \
	_T("         LEFT JOIN hms_doc ON ( hpo_docno = hd_docno ) ") \
	_T("         LEFT JOIN hms_patient ON ( hd_patientno = hp_patientno ) ") \
	_T("         LEFT JOIN hms_object ON ( ho_id = hpo_object_id ) ") \
	_T("		 WHERE 1=1 %s %s) ") \
	_T(" WHERE 1=1 %s") \
	_T(" GROUP  BY orderid,docno,hpo_doctor,NAME,price,product_purchase_uomid,product_id ") \
	_T(" ORDER BY orderid, docno, hpo_doctor, product_id"), szWhere, szWhereRoomID, szWhere2);

	//rpt.GetReportHeader()->SetValue(_T("RoomName"), szRoomName);

	CReportSection* rptDetail = NULL;
	CReportSection* rptGroup = NULL;
	CRecord rs(&pMF->m_db);
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		//_msg(_T("%s"), szSQL);
		//QueryOpener(szSQL);
		return;
	}
	long nItem = 1, nTotalOrder=0, nItemCount = 0;
	long nOrderID;
	double nTotalAmount=0,cost = 0, nAmount = 0;
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("orderid"), szNewOrderID);
		if (szNewOrderID != szOldOrderID)
		{
			if (nAmount > 0)
			{			
				rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));	
				if (rptGroup)
				{
					rptGroup->GetItem(_T("PName"))->SetTextAlign(1);
					rptGroup->GetItem(_T("PName"))->SetItalic(true);
					rptGroup->GetItem(_T("Amount"))->SetItalic(true);
					TranslateString(_T("Total Amount"), tmpStr);
					rptGroup->SetValue(_T("PName"), tmpStr);
					FormatCurrency(nAmount, tmpStr);		
					rptGroup->SetValue(_T("Amount"), tmpStr);
				}
				nTotalAmount += nAmount;
				nAmount = 0;
			}
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));
			if (rptGroup)
			{
				rs.GetValue(_T("orderid"), nOrderID);
				tmpStr.Format(_T("%ld"), nOrderID);
				rptGroup->SetValue(_T("OrderID"), tmpStr);
				TranslateString(_T("Age"), tmpStr2);
				tmpStr.Format(_T("[%s]  -  %s   -   %s:%s  -  %s"), rs.GetValue(_T("docno")), rs.GetValue(_T("pname")),tmpStr2,rs.GetValue(_T("Age")), rs.GetValue(_T("doctor")));
				rptGroup->SetValue(_T("PName"), tmpStr);
				rs.GetValue(_T("doctor"), tmpStr);
				rptGroup->SetValue(_T("Amount"), tmpStr);
			}
			nItemCount += nItem - 1;
			nTotalOrder++;
			szOldOrderID = szNewOrderID;
			nItem = 1;
		}
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			tmpStr.Format(_T("%ld"), nItem++);			
			rptDetail->SetValue(_T("1"), tmpStr);
			rs.GetValue(_T("name"), tmpStr);
			rptDetail->SetValue(_T("2"), tmpStr);
			rs.GetValue(_T("unit"), tmpStr);
			rptDetail->SetValue(_T("3"), tmpStr);			
			rs.GetValue(_T("price"), tmpStr);			
			rptDetail->SetValue(_T("4"), tmpStr);
			rs.GetValue(_T("qty"), tmpStr);
			rptDetail->SetValue(_T("5"), tmpStr);
			rs.GetValue(_T("amount"), cost);
			FormatCurrency(cost, tmpStr);			
			rptDetail->SetValue(_T("6"), tmpStr);				
			nAmount +=cost;
		}
		rs.MoveNext();
	}
	if (nAmount > 0)
	{			
		rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));	
		if (rptGroup)
		{
			rptGroup->GetItem(_T("PName"))->SetTextAlign(1);
			rptGroup->GetItem(_T("PName"))->SetItalic(true);
			rptGroup->GetItem(_T("Amount"))->SetItalic(true);
			TranslateString(_T("Total Amount"), tmpStr);
			rptGroup->SetValue(_T("PName"), tmpStr);
			FormatCurrency(nAmount, tmpStr);		
			rptGroup->SetValue(_T("Amount"), tmpStr);
		}
		nTotalAmount += nAmount;
		nAmount = 0;
	}
	nItemCount += nItem - 1;
	//Page Footer
	//Report Footer
	
	tmpStr.Format(_T("%ld"), nTotalOrder);
	rpt.GetReportFooter()->SetValue(_T("TotalOrder"), tmpStr);
	tmpStr.Format(_T("%.2f"), nTotalAmount);
	FormatCurrency(nTotalAmount, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	tmpStr.Format(_T("%ld"), nItemCount);
	rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	//rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
	rpt.PrintPreview();
	EndWaitCursor();


}

void CPMSPrescriptionListReportDialog::OnPrintSynthesis(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL;
	double nTmp = 0;
	int nIdx = 1, j = 0, nTotalOrder = 0;
	CString szSQL, tmpStr, szReportName, szDate, szPos;
	long double nTotalAmount = 0;
	szReportName = _T("Reports\\HMS\\PMS_PRESCRIPTIONLISTREPORT_TH.RPT");
	if (!rpt.Init(szReportName))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF()){
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
		CDate::Convert(m_szToDate), yyyymmdd, ddmmyyyy);
		rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	}
	while (!rs.IsEOF())
	{
		//Group Seperation
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("1"), rs.GetValue(_T("orderid")));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("pname")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("age")));
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("docno")));
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("doctor")));
			rs.GetValue(_T("amount"), nTmp);
			nTotalAmount+= nTmp;
			rptDetail->SetValue(_T("6"), double2str(nTmp));
		}
		nTotalOrder++;
		rs.MoveNext();
	}
	rptHeader = rpt.GetReportFooter();
	if (rptHeader)
	{
		if (nTotalOrder > 0)
			rptHeader->SetValue(_T("TotalOrder"), int2str(nTotalOrder));
		if (nTotalAmount > 0)
		{
			tmpStr.Format(_T("%.2f"), nTotalAmount);
			rptHeader->SetValue(_T("TotalAmount"), tmpStr);
		}
		tmpStr = pMF->GetSysDate();
		szDate.Format(rptHeader->GetValue(_T("PrintDate")), tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4));
		rptHeader->SetValue(_T("PrintDate"), szDate);
	}
		
	rpt.PrintPreview();
}

void CPMSPrescriptionListReportDialog::OnExportSynthesis(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CStringArray arrCol;
	CExcel xls;
	CString szSQL, tmpStr;
	double nTmp = 0;
	long double nTotalAmount = 0;
	int nIdx = 1, nRow = 0, nTotalOrder = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 10);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(2, 5);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 25);
	xls.SetColumnWidth(5, 15);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 6);
	xls.SetCellMergedColumns(0, 3, 6);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	xls.SetCellText(0, 2, _T("\x42\x1EA2NG K\xCA \x110\x1A0N T\x1ED4NG H\x1EE2P"), 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	arrCol.Add(_T("S\x1ED1 phi\x1EBFu"));
	arrCol.Add(_T("T\xEAn \x42N"));
	arrCol.Add(_T("Tu\x1ED5i"));
	arrCol.Add(_T("S\x1ED1 h\x1ED3 s\x1A1"));
	arrCol.Add(_T("T\xEAn \x62\xE1\x63 s\x1EF9"));
	arrCol.Add(_T("Th\xE0nh ti\x1EC1n"));
	for (int i = 0; i < arrCol.GetCount(); i++)
		xls.SetCellText(i, 4, arrCol.GetAt(i), 4098, true);
	nRow = 5;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, rs.GetValue(_T("orderid")), FMT_TEXT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("pname")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("age")), FMT_TEXT);
		xls.SetCellText(3, nRow, rs.GetValue(_T("docno")), FMT_TEXT);
		xls.SetCellText(4, nRow, rs.GetValue(_T("doctor")), FMT_TEXT);
		rs.GetValue(_T("amount"), nTmp);
		nTotalAmount += nTmp;
		xls.SetCellText(5, nRow, double2str(nTmp), FMT_NUMBER1);
		nRow++;
		nTotalOrder++;
		rs.MoveNext();
	}
	if (nTotalOrder > 0)
	{
		xls.SetCellText(4, nRow, _T("S\x1ED1 \x111\x1A1n:"), 4098, true);
		xls.SetCellText(5, nRow, int2str(nTotalOrder), FMT_NUMBER1, true);
		nRow++;
		xls.SetCellText(4, nRow, _T("T\x1ED5ng s\x1ED1 ti\x1EC1n:"), 4098, true);
		tmpStr.Format(_T("%f"), nTotalAmount);
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\Bang ke don thuoc_TH.xls"));
}

void CPMSPrescriptionListReportDialog::OnExport()
{
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	UpdateData(true);
	CString tmpStr, szSQL, szObjects, szWhere, szWhere1, szProductType, tmpStr2, szObjectID, 
		szObjectName,szRoomName,szRoomID, szListID, szTemp, szOldOrderID, szNewOrderID, szDept;
	CExcel xls;
	double nTemp = 0;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	CellFormat df(&xls), cf(&xls), hf(&xls);
	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER);
	cf.SetFontName(_T("Segoe UI"));

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 30);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 13);
	xls.SetColumnWidth(5, 13);
	xls.SetColumnWidth(6, 15);
	int nCol = 0;
	int nRow = 0;
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellText(0, 0, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(0, 2, 6);
	TranslateString(_T("Prescription List Report"), szTemp);
	StringUpper(szTemp, tmpStr);
	xls.SetCellText(0, 2, tmpStr, FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(0, 3, 6);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow+3, tmpStr, &df);
	nRow = 4;
	if (!m_szStockKey.IsEmpty())
	{
		xls.SetCellMergedColumns(0, nRow, 6);
		xls.SetCellText(nCol, nRow++, m_wndStock.GetCurrent(1), &cf);
	}
	if (!m_szDoctorKey.IsEmpty())
	{
		xls.SetCellMergedColumns(0, nRow, 6);
		xls.SetCellText(nCol, nRow++, m_wndDoctor.GetCurrent(1), &cf);
	}
	tmpStr2 = _T("' '");
	
	//for(int i=0; i < m_wndObject.GetCount(); i++)
	//{
	//	if(m_wndObject.GetCheck(i))
	//	{
	//		m_wndObject.SetCurSel(i);
	//		if (!szObjectID.IsEmpty())
	//			szObjectID += _T(",");
	//		szObjectID.AppendFormat(_T("'%s'"), m_wndObject.GetCurrent(0));

	//		if (!szObjectName.IsEmpty())
	//			szObjectName += _T(",");
	//		szObjectName.AppendFormat(_T("%s"), m_wndObject.GetCurrent(1));
	//	}
	//}
	szWhere.Format(_T(" AND hpo_orderstatus = 'A' AND hpo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				   _T(" AND hpo_ordertype IN ('P', 'D', 'M', 'E')"), m_szFromDate, m_szToDate);
	if (!m_szObjectKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND ho_type in('%s') "), m_szObjectKey);	
	CString szDeptType, szWhereRoomID, szRoomC11, szRoomC12, szRoomC13;	
	for(int i=0; i < m_wndList.GetItemCount(); i++)
	{
		if(m_wndList.GetCheck(i))
		{
			m_wndList.SetCurSel(i);
			if (!szListID.IsEmpty())
				szListID += _T(",");
			szListID.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i,0));
			szDept = m_wndList.GetItemText(i, 2);
			if (szDept == _T("C1.1"))
			{
				if (!szRoomC11.IsEmpty())
					szRoomC11 += _T(", ");
				szRoomC11.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
			}
			else if (szDept == _T("C1.2"))
			{
				if (!szRoomC12.IsEmpty())
					szRoomC12 += _T(", ");
				szRoomC12.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
			}
			else if (szDept == _T("C1.3"))
			{
				if (!szRoomC13.IsEmpty())
					szRoomC13 += _T(", ");
				szRoomC13.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
			}
			if (!szRoomName.IsEmpty())
				szRoomName += _T(", ");
			szRoomName.AppendFormat(_T("%s"), m_wndList.GetItemText(i,1));
		}
	}
	if (!szRoomName.IsEmpty())
	{
		tmpStr.Format(_T("Khoa Ph\xF2ng: %s"), szRoomName);
		xls.SetCellMergedColumns(0, nRow, 6);
		xls.SetCellText(0, nRow++, tmpStr, &cf);
	}
	xls.SetCellText(0, nRow, _T("STT"), &hf);
	xls.SetCellText(1, nRow, _T("M\xE3 thu\x1ED1\x63"), &hf);
	xls.SetCellText(2, nRow, _T("T\xEAn thu\x1ED1\x63/ H\xE0m l\x1B0\x1EE3ng"), &hf);
	xls.SetCellText(3, nRow, _T("\x110\x1A1n v\x1ECB"), &hf);
	xls.SetCellText(4, nRow, _T("\x110\x1A1n gi\xE1"), &hf);		
	xls.SetCellText(5, nRow, _T("SL"), &hf);
	xls.SetCellText(6, nRow, _T("Th\xE0nh ti\x1EC1n"), &hf);
	nRow++;
	BeginWaitCursor();

	szDeptType.Empty();
	szWhereRoomID.Empty();
	if (m_szCondition == _T("A"))
		szDeptType.IsEmpty();
	else
		if (m_szCondition == _T("O"))
		{
			szDeptType.Format(_T(" and hpo_depttype ='KB' "));
			if (!szRoomC11.IsEmpty())				
				szWhereRoomID.AppendFormat(_T(" and ((hpo_roomid in (%s) and hpo_deptid= 'C1.1') "), szRoomC11);
			if (!szRoomC12.IsEmpty())
				if (!szWhereRoomID.IsEmpty())
					szWhereRoomID.AppendFormat(_T(" or (hpo_roomid in (%s) and hpo_deptid= 'C1.2') "), szRoomC12);
				else
					szWhereRoomID.AppendFormat(_T(" and ((hpo_roomid in (%s) and hpo_deptid= 'C1.2') "), szRoomC12);
			if (!szRoomC13.IsEmpty())
				if (!szWhereRoomID.IsEmpty())
					szWhereRoomID.AppendFormat(_T(" or (hpo_roomid in (%s) and hpo_deptid= 'C1.3')"), szRoomC13);
				else
					szWhereRoomID.AppendFormat(_T(" and ((hpo_roomid in (%s) and hpo_deptid= 'C1.3')"), szRoomC13);
			if (!szWhereRoomID.IsEmpty())
				szWhereRoomID.AppendFormat(_T(")"));
		}
		else
		{
			szDeptType.Format(_T(" and hpo_depttype ='DT' "));
			if(!szListID.IsEmpty())
				szWhereRoomID.Format(_T(" and hpo_deptid in(%s) ") , szListID);
		}
	if (!szDeptType.IsEmpty())
		szWhere.AppendFormat(_T("%s"), szDeptType);

	if(!m_szDoctorKey.IsEmpty())
		szWhere.AppendFormat(_T(" and hpo_doctor='%s' "), m_szDoctorKey);
	if (!m_szStockKey.IsEmpty())
		szWhere.AppendFormat(_T(" and hpo_storage_id=%d"), ToInt(m_szStockKey));
	if (m_nMinTotalAmt > 0)
		szWhere1.AppendFormat(_T(" AND taxamount >= %ld"), m_nMinTotalAmt);
	if (!m_szApprovedByKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_approveby = '%s'"), m_szApprovedByKey);
	//szPrice = _T("product_taxprice");
	//if (m_bServPrice)
	//szPrice = _T("hpol_unitprice");
	if (m_bOriental)
		szProductType = _T(" product_producttype = 'A1301'");
	if (m_bOccident)
	{
		if (!szProductType.IsEmpty())
			szProductType += _T(" OR");
		szProductType += _T(" product_producttype <> 'A1301'");
	}
	if (!szProductType.IsEmpty())
		szWhere1.Format(_T(" AND (%s)"), szProductType);
	//Page Header
	//Report Detail
	szSQL.Format(_T(" SELECT orderid, ") \
	_T("		product_code id,") \
	_T("        docno, ") \
	_T("        Get_patientname(docno) pname, ") \
	_T("        Hms_getagebydoc(docno) age, ") \
	_T("        Get_username(hpo_doctor) doctor, ") \
	_T("        NAME, ") \
	_T("        Get_uomname(product_purchase_uomid) unit, ") \
	_T("        price, ") \
	_T("        SUM(qty) qty, ") \
	_T("        SUM(amount) amount ") \
	_T(" FROM   (SELECT    hpol_orderid orderid, ") \
	_T("                   hpol_qtyissue * hpol_unitprice amount, ") \
	_T("                   SUM(hpol_qtyissue * hpol_unitprice) ") \
	_T("                     OVER ( ") \
	_T("                       PARTITION BY (hpol_orderid)) taxamount, ") \
	_T("                   hpo_docno docno, ") \
	_T("                   hpo_doctor, ") \
	_T("                   product_id, ") \
	_T("				   product_code,") \
	_T("                   product_name NAME, ") \
	_T("                   product_purchase_uomid, ") \
	_T("                   hpol_unitprice price, ") \
	_T("                   hpol_qtyissue qty ") \
	_T("         FROM      hms_pharmaorder ") \
	_T("         LEFT JOIN hms_pharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
	_T("         LEFT JOIN m_productitem_view ON ( product_item_id = hpol_product_item_id ) ") \
	_T("         LEFT JOIN hms_doc ON ( hpo_docno = hd_docno ) ") \
	_T("         LEFT JOIN hms_patient ON ( hd_patientno = hp_patientno ) ") \
	_T("         LEFT JOIN hms_object ON ( ho_id = hpo_object_id ) ") \
	_T("		 WHERE 1=1 %s %s) ") \
	_T(" WHERE 1=1 %s") \
	_T(" GROUP  BY orderid,docno,hpo_doctor,NAME,price,product_purchase_uomid,product_id, product_code ") \
	_T(" ORDER BY orderid, docno, hpo_doctor, product_id, product_code"), szWhere, szWhereRoomID, szWhere1);

	//QueryOpener(szSQL);

	CRecord rs(&pMF->m_db);
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	int nItem = 1;
	long nItemCount = 0;
	long nOrderID = 0, nTotalOrder=0;
	long double nTotalAmount=0, nAmount = 0;

	while(!rs.IsEOF())
	{	
		rs.GetValue(_T("orderid"), szNewOrderID);
		if (szNewOrderID != szOldOrderID)
		{
			if (nAmount > 0)
			{
				TranslateString(_T("Total Amount"), tmpStr);
				//xls.SetCellMergedColumns(1,nRow,3);
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
				tmpStr.Format(_T("%.2f"), nAmount);
				xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);	
				nTotalAmount += nAmount;
				nAmount = 0;
				nRow++;		
			}
			//xls.SetCellMergedColumns(1,nRow,4);		
			rs.GetValue(_T("orderid"), nOrderID);
			tmpStr.Format(_T("%ld"), nOrderID);
			xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER,true,10,0);
			TranslateString(_T("Age"), tmpStr2);
			tmpStr.Format(_T("[%s]  -  %s   -   %s:%s  -  %s"), rs.GetValue(_T("docno")), rs.GetValue(_T("pname")),
				tmpStr2,rs.GetValue(_T("Age")), rs.GetValue(_T("doctor")));
			xls.SetCellText(1, nRow, tmpStr, FMT_TEXT,true,10,0);
			rs.GetValue(_T("doctor"), tmpStr);
			xls.SetCellText(5, nRow, tmpStr, FMT_TEXT,true,10,0);

			szOldOrderID = szNewOrderID;
			nTotalOrder++;
			nItemCount += nItem - 1;
			nItem = 1;
			nRow++;
		}
		tmpStr.Format(_T("%ld"), nItem++);			
		xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);
		rs.GetValue(_T("id"), tmpStr);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("name"), tmpStr);
		xls.SetCellText(2, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("unit"), tmpStr);
		xls.SetCellText(3, nRow, tmpStr, FMT_TEXT);			
		rs.GetValue(_T("price"), tmpStr);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("qty"), tmpStr);
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("amount"), nTemp);
		nAmount += nTemp;
		xls.SetCellText(6, nRow, double2str(nTemp), FMT_NUMBER1);	
		nRow++;
		rs.MoveNext();
	}
	nItemCount += nItem - 1;
	if (nAmount > 0)
	{
		TranslateString(_T("Total Amount"), tmpStr);
		xls.SetCellMergedColumns(1,nRow,3);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
		tmpStr.Format(_T("%.2f"), nAmount);
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);	
		nTotalAmount += nAmount;
		nRow++;		
	}
	//Page Footer
	//Report Footer
	TranslateString(_T("Total"), tmpStr);
	xls.SetCellMergedColumns(1,nRow,3);
	xls.SetCellText(1, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
	tmpStr.Format(_T("%ld"), nTotalOrder);
	xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);
	nRow++;
	//Tong so thuoc trong don
	TranslateString(_T("Total Item"), tmpStr);
	xls.SetCellMergedColumns(1,nRow,3);
	xls.SetCellText(1, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
	tmpStr.Format(_T("%ld"), nItemCount);
	xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);
	nRow++;
	//Tong tien
	TranslateString(_T("Total Amount"), tmpStr);
	xls.SetCellMergedColumns(1,nRow,3);
	xls.SetCellText(1, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
	tmpStr.Format(_T("%.2f"), nTotalAmount);
	xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);

	xls.Save(_T("Exports\\Bang Ke Don Thuoc.xls"));
	EndWaitCursor();

}

CString CPMSPrescriptionListReportDialog::GetQueryString(){
	CString szSQL, szDeptType, szWhereRoomID, szWhere, szListID, szRoomID, szRoomName, szHavingClause, szProductType;
	szDeptType.Empty();
	szWhereRoomID.Empty();
	if (!m_szObjectKey.IsEmpty())
		szWhere.Format(_T(" AND ho_type in('%s') "), m_szObjectKey);	
	
	CString szRoomC11, szRoomC12, szRoomC13, szDept;
	for(int i=0; i < m_wndList.GetItemCount(); i++)
	{
		if(m_wndList.GetCheck(i))
		{
			m_wndList.SetCurSel(i);
			if (!szListID.IsEmpty())
				szListID += _T(",");
			szListID.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i,0));
			szDept = m_wndList.GetItemText(i, 2);
			if (szDept == _T("C1.1"))
			{
				if (!szRoomC11.IsEmpty())
					szRoomC11 += _T(", ");
				szRoomC11.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
			}
			else if (szDept == _T("C1.2"))
			{
				if (!szRoomC12.IsEmpty())
					szRoomC12 += _T(", ");
				szRoomC12.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
			}
			else if (szDept == _T("C1.3"))
			{
				if (!szRoomC13.IsEmpty())
					szRoomC13 += _T(", ");
				szRoomC13.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
			}
			if (!szRoomName.IsEmpty())
				szRoomName += _T(", ");
			szRoomName.AppendFormat(_T("%s"), m_wndList.GetItemText(i,1));
		}
	}
	if (m_szCondition == _T("A"))
		szDeptType.IsEmpty();
	else
		if (m_szCondition == _T("O"))
		{
			szDeptType.Format(_T(" and hpo_depttype ='KB' "));
			if (!szRoomC11.IsEmpty())				
				szWhereRoomID.AppendFormat(_T(" and ((hpo_roomid in (%s) and hpo_deptid= 'C1.1') "), szRoomC11);
			if (!szRoomC12.IsEmpty())
				if (!szWhereRoomID.IsEmpty())
					szWhereRoomID.AppendFormat(_T(" or (hpo_roomid in (%s) and hpo_deptid= 'C1.2') "), szRoomC12);
				else
					szWhereRoomID.AppendFormat(_T(" and ((hpo_roomid in (%s) and hpo_deptid= 'C1.2') "), szRoomC12);
			if (!szRoomC13.IsEmpty())
				if (!szWhereRoomID.IsEmpty())
					szWhereRoomID.AppendFormat(_T(" or (hpo_roomid in (%s) and hpo_deptid= 'C1.3')"), szRoomC13);
				else
					szWhereRoomID.AppendFormat(_T(" and ((hpo_roomid in (%s) and hpo_deptid= 'C1.3')"), szRoomC13);
			if (!szWhereRoomID.IsEmpty())
				szWhereRoomID.AppendFormat(_T(")"));
		}
		else
		{
			szDeptType.Format(_T(" and hpo_depttype ='DT' "));
			if(!szListID.IsEmpty())
				szWhereRoomID.Format(_T(" and hpo_deptid in(%s) ") , szListID);
		}
	if (!szDeptType.IsEmpty())
		szWhere.AppendFormat(_T("%s"), szDeptType);

	if(!m_szDoctorKey.IsEmpty())
		szWhere.AppendFormat(_T(" and hpo_doctor='%s' "), m_szDoctorKey);
	if (!m_szStockKey.IsEmpty())
		szWhere.AppendFormat(_T(" and hpo_storage_id=%d"), ToInt(m_szStockKey));
	//szPrice = _T("product_taxprice");
	//if (m_bServPrice)
	//szPrice = _T("hpol_unitprice");
	if (m_nMinTotalAmt > 0)
		szHavingClause.Format(_T(" HAVING SUM(hpol_qtyissue*hpol_unitprice) > %ld"), m_nMinTotalAmt);
	if (!m_szApprovedByKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_approveby = '%s'"), m_szApprovedByKey);
	if (m_bOriental)
		szProductType = _T(" product_producttype = 'A1301'");
	if (m_bOccident)
	{
		if (!szProductType.IsEmpty())
			szProductType += _T(" OR");
		szProductType += _T(" product_producttype <> 'A1301'");
	}
	if (!szProductType.IsEmpty())
		szWhere.AppendFormat(_T(" AND (%s)"), szProductType);	
	szSQL.Format(_T(" SELECT hpo_orderid as orderid,") \
		_T(" 	hpo_docno as docno,") \
		_T("	get_username(hpo_doctor) as doctor,") \
		_T(" 	get_patientname(hpo_docno) as pname,") \
		_T(" 	hms_getagebydoc(hpo_docno) as age, ") \
		_T("	sum(hpol_qtyissue*hpol_unitprice) amount ")
		_T(" FROM hms_pharmaorder ") \
		_T(" LEFT JOIN hms_pharmaorderline ON (hpo_orderid = hpol_orderid)") \
		_T(" LEFT JOIN m_productitem_view ON (product_item_id = hpol_product_item_id)") \
		_T(" LEFT JOIN hms_doc ON (hpo_docno=hd_docno)") \
		_T(" LEFT JOIN hms_object ON (hpo_object_id = ho_id)") \
		_T(" WHERE 	hpo_orderstatus IN('A','I')") \
		_T("	and hpo_approvedate between cast_string2timestamp('%s') and cast_string2timestamp('%s') ")
		_T(" 	and hpo_ordertype in('P','D','M','E') %s %s ") \
		_T(" GROUP BY hpo_orderid, hpo_docno, hpo_doctor %s") \
		_T(" ORDER BY hpo_orderid"), m_szFromDate, m_szToDate, szWhere, szWhereRoomID, szHavingClause);

	return szSQL;
}

BEGIN_MESSAGE_MAP(CPMSPrescriptionListReportDialog, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSPrescriptionListReportDialog::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
