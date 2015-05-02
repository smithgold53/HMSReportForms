#include "stdafx.h"
#include "EMPrescriptionListReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMPrescriptionListReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMPrescriptionListReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMPrescriptionListReport* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CEMPrescriptionListReport *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnStockAddNew();
}*/
static void _OnDoctorSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMPrescriptionListReport* )pWnd)->OnDoctorSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDoctorSelendokFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnDoctorSelendok();
}
/*static void _OnDoctorSetfocusFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnDoctorKillfocus();
}*/
/*static void _OnDoctorKillfocusFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnDoctorKillfocus();
}*/
static long _OnDoctorLoadDataFnc(CWnd *pWnd){
	return ((CEMPrescriptionListReport *)pWnd)->OnDoctorLoadData();
}
/*static void _OnDoctorAddNewFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnDoctorAddNew();
}*/
static long _OnApprovedByLoadDataFnc(CWnd *pWnd){
	return ((CEMPrescriptionListReport *)pWnd)->OnApprovedByLoadData();
}
static int _OnMinTotalAmtCheckValueFnc(CWnd *pWnd){
	return ((CEMPrescriptionListReport *)pWnd)->OnMinTotalAmtCheckValue();
}
static void _OnCheckAllSelectFnc(CWnd *pWnd){
	 ((CEMPrescriptionListReport*)pWnd)->OnCheckAllSelect();
}
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	 ((CEMPrescriptionListReport*)pWnd)->OnOutPatientSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	 ((CEMPrescriptionListReport*)pWnd)->OnInPatientSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMPrescriptionListReport *pVw = (CEMPrescriptionListReport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMPrescriptionListReport *pVw = (CEMPrescriptionListReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMPrescriptionListReport *pVw = (CEMPrescriptionListReport *)pWnd;
	pVw->OnCloseSelect();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMPrescriptionListReport* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CEMPrescriptionListReport *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CEMPrescriptionListReport *)pWnd)->OnObjectAddNew();
}*/
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CEMPrescriptionListReport*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CEMPrescriptionListReport*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CEMPrescriptionListReport*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListCheckC11RoomFnc(CWnd *pWnd){
	return ((CEMPrescriptionListReport*)pWnd)->OnListCheckC11Room();
} 
static int _OnListCheckC12RoomFnc(CWnd *pWnd){
	return ((CEMPrescriptionListReport*)pWnd)->OnListCheckC12Room();
} 
static int _OnListCheckC13RoomFnc(CWnd *pWnd){
	return ((CEMPrescriptionListReport*)pWnd)->OnListCheckC13Room();
} 
static int _OnAddPMSPrescriptionListReportDialogFnc(CWnd *pWnd){
	 return ((CEMPrescriptionListReport*)pWnd)->OnAddPMSPrescriptionListReportDialog();
} 
static int _OnEditPMSPrescriptionListReportDialogFnc(CWnd *pWnd){
	 return ((CEMPrescriptionListReport*)pWnd)->OnEditPMSPrescriptionListReportDialog();
}
static int _OnListUncheckAllFnc(CWnd *pWnd){
	return ((CEMPrescriptionListReport*)pWnd)->OnListUncheckAll();
} 
static int _OnDeletePMSPrescriptionListReportDialogFnc(CWnd *pWnd){
	 return ((CEMPrescriptionListReport*)pWnd)->OnDeletePMSPrescriptionListReportDialog();
} 
static int _OnSavePMSPrescriptionListReportDialogFnc(CWnd *pWnd){
	 return ((CEMPrescriptionListReport*)pWnd)->OnSavePMSPrescriptionListReportDialog();
} 
static int _OnCancelPMSPrescriptionListReportDialogFnc(CWnd *pWnd){
	 return ((CEMPrescriptionListReport*)pWnd)->OnCancelPMSPrescriptionListReportDialog();
} 
CEMPrescriptionListReport::CEMPrescriptionListReport(CWnd *pParent)
{
	m_nDlgWidth = 590;
	m_nDlgHeight = 495;
	SetDefaultValues();
}
CEMPrescriptionListReport::~CEMPrescriptionListReport(){
}
void CEMPrescriptionListReport::OnCreateComponents(){
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 110, 55);
	m_wndFromDate.Create(this,115, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 395, 55);
	m_wndToDate.Create(this,400, 30, 570, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 110, 85);
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 550);
	m_wndStock.Create(this,115, 60, 570, 85); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 89, 110, 114);
	m_wndObject.Create(this,115, 89, 570, 114); 
	m_wndDoctorLabel.Create(this, _T("Doctor"), 10, 120, 110, 145);
	m_wndDoctor.Create(this,115, 120, 290, 145); 
	m_wndApprovedByLabel.Create(this, _T("Approved By"), 295, 120, 395, 145);
	m_wndApprovedBy.Create(this,400, 120, 570, 145); 
	m_wndMinTotalAmtLabel.Create(this, _T("Min Total Amt"), 10, 150, 110, 175);
	m_wndMinTotalAmt.Create(this,115, 150, 290, 175); 
	m_wndServicePrice.Create(this, _T("Serv Price"), 10, 180, 110, 205);
	m_wndCheckAll.Create(this, _T("Check All"), 260, 180, 360, 205);
	m_wndOutPatient.Create(this, _T("OutPatient"), 365, 180, 465, 205);
	m_wndInPatient.Create(this, _T("In Patient"), 470, 180, 570, 205);
	m_wndList.Create(this,10, 210, 570, 545); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 370, 555, 480, 580);
	m_wndExport.Create(this, _T("&Export"), 485, 555, 575, 580);

}
void CEMPrescriptionListReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
void CEMPrescriptionListReport::OnSetWindowEvents(){
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
void CEMPrescriptionListReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_TextEx(pDX, m_wndDoctor.GetDlgCtrlID(), m_szDoctorKey);
	DDX_TextEx(pDX, m_wndApprovedBy.GetDlgCtrlID(), m_szApprovedByKey);
	DDX_Text(pDX, m_wndMinTotalAmt.GetDlgCtrlID(), m_nMinTotalAmt);
	DDX_Check(pDX, m_wndServicePrice.GetDlgCtrlID(), m_bServPrice);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
}
void CEMPrescriptionListReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMPrescriptionListReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMPrescriptionListReport::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_szDoctorKey.Empty();
	m_szObjectKey.Empty();

}
int CEMPrescriptionListReport::SetMode(int nMode){
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
/*void CEMPrescriptionListReport::OnFromDateChange(){
	
} */
/*void CEMPrescriptionListReport::OnFromDateSetfocus(){
	
} */
/*void CEMPrescriptionListReport::OnFromDateKillfocus(){
	
} */
int CEMPrescriptionListReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMPrescriptionListReport::OnToDateChange(){
	
} */
/*void CEMPrescriptionListReport::OnToDateSetfocus(){
	
} */
/*void CEMPrescriptionListReport::OnToDateKillfocus(){
	
} */
int CEMPrescriptionListReport::OnToDateCheckValue(){
	return 0;
} 
void CEMPrescriptionListReport::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMPrescriptionListReport::OnStockSelendok(){
	 
}
/*void CEMPrescriptionListReport::OnStockSetfocus(){
	
}*/
/*void CEMPrescriptionListReport::OnStockKillfocus(){
	
}*/
long CEMPrescriptionListReport::OnStockLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
/*void CEMPrescriptionListReport::OnStockAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMPrescriptionListReport::OnDoctorSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMPrescriptionListReport::OnDoctorSelendok(){
	 
}
/*void CEMPrescriptionListReport::OnDoctorSetfocus(){
	
}*/
/*void CEMPrescriptionListReport::OnDoctorKillfocus(){
	
}*/
long CEMPrescriptionListReport::OnDoctorLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();		

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
/*void CEMPrescriptionListReport::OnDoctorAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
long CEMPrescriptionListReport::OnApprovedByLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();		
	CString szFilter;
	szFilter.Format(_T(" AND su_deptid = 'KD'"));
	return pMF->LoadUserList(&m_wndApprovedBy, m_szApprovedByKey, szFilter);

}
int CEMPrescriptionListReport::OnMinTotalAmtCheckValue(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	return 0;
}

void CEMPrescriptionListReport::OnListDblClick(){
	
} 
void CEMPrescriptionListReport::OnListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CEMPrescriptionListReport::OnListCheckC11Room(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
int CEMPrescriptionListReport::OnListCheckC12Room(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
int CEMPrescriptionListReport::OnListCheckC13Room(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
int CEMPrescriptionListReport::OnListUncheckAll(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	if (m_szCondition == _T("I"))
		return 0;
	for (int i = 0 ; i < m_wndList.GetItemCount(); i++)
	{
		m_wndList.SetCheck(i, false);
	}
	return 0;
} 
long CEMPrescriptionListReport::OnListLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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

void CEMPrescriptionListReport::OnCheckAllSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szCondition = _T("A");
	OnListLoadData();
}
void CEMPrescriptionListReport::OnOutPatientSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szCondition = _T("O");
	OnListLoadData();
}
void CEMPrescriptionListReport::OnInPatientSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szCondition = _T("I");
	OnListLoadData();
}
void CEMPrescriptionListReport::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	OnPrintPreview();
} 
void CEMPrescriptionListReport::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	OnExport();
} 
void CEMPrescriptionListReport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

} 
void CEMPrescriptionListReport::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMPrescriptionListReport::OnObjectSelendok(){
	 
}
/*void CEMPrescriptionListReport::OnObjectSetfocus(){
	
}*/
/*void CEMPrescriptionListReport::OnObjectKillfocus(){
	
}*/
long CEMPrescriptionListReport::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

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
/*void CEMPrescriptionListReport::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
int CEMPrescriptionListReport::OnAddPMSPrescriptionListReportDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMPrescriptionListReport::OnEditPMSPrescriptionListReportDialog(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMPrescriptionListReport::OnDeletePMSPrescriptionListReportDialog(){
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
 		OnCancelPMSPrescriptionListReportDialog(); 
 	}
	return 0;
}
int CEMPrescriptionListReport::OnSavePMSPrescriptionListReportDialog(){
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
 		//OnPMSPrescriptionListReportDialogListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMPrescriptionListReport::OnCancelPMSPrescriptionListReportDialog(){
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
int CEMPrescriptionListReport::OnPMSPrescriptionListReportDialogListLoadData(){
	return 0;
}


void CEMPrescriptionListReport::OnPrintPreview()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL, szObjectID, szObjectName, szWhere, tmpStr2,szListID, szRoomID, szRoomName;
	if(!rpt.Init(_T("Reports/HMS/PM_BANGKETHUOCVTC13.RPT"),true) )
	{
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);


	rpt.GetReportHeader()->SetValue(_T("StockName"), m_wndStock.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Doctor"), m_wndDoctor.GetCurrent(1));
	

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
		if (m_bServPrice)
			szWhere.AppendFormat(_T(" AND hpo_amount >= %ld"), m_nMinTotalAmt);
		else
			szWhere.AppendFormat(_T(" AND taxamount >=%ld"), m_nMinTotalAmt);
	if (!m_szApprovedByKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_approveby = '%s'"), m_szApprovedByKey);
	rpt.GetReportHeader()->SetValue(_T("Object"), szObjectName);
	//Page Header
	//Report Detail
	szSQL.Format(_T("SELECT hpo_orderid                                AS orderid, ") \
				_T("       hd_docno                                   AS docno, ") \
				_T("       Get_username(hpo_doctor)                   AS doctor, ") \
				_T("       Trim(hp_surname ") \
				_T("            || ' ' ") \
				_T("            || hp_midname ") \
				_T("            || ' ' ") \
				_T("            || hp_firstname)                      AS pname, ") \
				_T("       Hms_getage(( hd_admitdate ), hp_birthdate) AS age, ") \
				_T("       ho_type                                    AS objecttype, ") \
				_T("       taxamount ") \
				_T("FROM   hms_pharmaorder ") \
				_T("       LEFT JOIN hms_doc ") \
				_T("              ON( hpo_docno = hd_docno ) ") \
				_T("       LEFT JOIN hms_patient ") \
				_T("              ON ( hd_patientno = hp_patientno ) ") \
				_T("       LEFT JOIN hms_object ") \
				_T("              ON( hd_object = ho_id ) ") \
				_T("       LEFT JOIN (SELECT hpol_orderid                          AS hpo_orderid, ") \
				_T("                         Sum(hpol_qtyissue * product_taxprice) AS taxamount ") \
				_T("                  FROM   hms_pharmaorderline ") \
				_T("                         LEFT JOIN m_productitem_view ") \
				_T("                                ON ( product_item_id = hpol_product_item_id ) ") \
				_T("                  GROUP  BY hpol_orderid) USING (hpo_orderid) ") \
				_T("WHERE  hpo_orderstatus IN( 'A', 'I' ) ") \
				_T("       AND hpo_approvedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s') ") \
				_T("       AND hpo_ordertype IN( 'P', 'D', 'M', 'E' ) %s %s") \
				_T("ORDER  BY hpo_approvedate"), m_szFromDate, m_szToDate, szWhere, szWhereRoomID);
	rpt.GetReportHeader()->SetValue(_T("RoomName"), szRoomName);

	CReportSection* rptDetail;
	CReportSection* rptGroup;
	CRecord rs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		//QueryOpener(szSQL);
		return;
	}
	long nItem = 1, nTotalOrder=0;
	long nOrderID = 0;
	double nTotalAmount=0,cost=0, nAmount;
	CString szPrice;
	szPrice = _T("product_taxprice");
	if (m_bServPrice)
		szPrice = _T("hpol_unitprice");

	while(!rs.IsEOF())
	{
		rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));
		rs.GetValue(_T("orderid"), nOrderID);
		tmpStr.Format(_T("%ld"), nOrderID);
		rptGroup->SetValue(_T("OrderID"), tmpStr);
		TranslateString(_T("Age"), tmpStr2);
		tmpStr.Format(_T("[%s]  -  %s   -   %s:%s  -  %s"), rs.GetValue(_T("docno")), rs.GetValue(_T("pname")),tmpStr2,rs.GetValue(_T("Age")), rs.GetValue(_T("doctor")));
		rptGroup->SetValue(_T("PName"), tmpStr);
		rs.GetValue(_T("doctor"), tmpStr);
		rptGroup->SetValue(_T("Amount"), tmpStr);
		
		szSQL.Format(_T(" SELECT") \
					_T("   product_id,") \
					_T("   hpol_line,") \
					_T("   product_name                          AS name,") \
					_T("   Get_uomname(product_purchase_uomid)           AS unit,") \
					_T("   %s                    AS price,") \
					_T("   Sum(hpol_qtyissue)               AS qty,") \
					_T("   Sum(hpol_qtyissue)*%s AS amount") \
					_T(" FROM   hms_pharmaorderline") \
					_T("        LEFT JOIN m_productitem_view ON(product_item_id=hpol_product_item_id)") \
					_T(" WHERE  hpol_orderid=%ld") \
					_T(" GROUP  BY product_id,") \
					_T("           product_name,") \
					_T("           product_purchase_uomid,") \
					_T("           %s,") \
					_T("           hpol_line "), szPrice, szPrice, nOrderID, szPrice);
		//_fmsg(_T("%s"), szSQL);
		rsl.ExecSQL(szSQL);
		nItem = 1;
		nAmount  =0;
		while(!rsl.IsEOF())
		{
			rptDetail = rpt.AddDetail();
			tmpStr.Format(_T("%ld"), nItem++);			
			rptDetail->SetValue(_T("1"), tmpStr);
			rsl.GetValue(_T("name"), tmpStr);
			rptDetail->SetValue(_T("2"), tmpStr);
			rsl.GetValue(_T("unit"), tmpStr);
			rptDetail->SetValue(_T("3"), tmpStr);			
			/*rsl.GetValue(_T("price"), tmpStr);			
			rptDetail->SetValue(_T("4"), tmpStr);*/
			rsl.GetValue(_T("qty"), tmpStr);
			rptDetail->SetValue(_T("4"), tmpStr);
			/*rsl.GetValue(_T("amount"), cost);
			FormatCurrency(cost, tmpStr);			
			rptDetail->SetValue(_T("6"), tmpStr);
			nAmount +=cost;*/
			rsl.MoveNext();
		}		 
		/*nTotalAmount += nAmount;
		rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));		
		rptGroup->GetItem(_T("PName"))->SetTextAlign(1);
		rptGroup->GetItem(_T("PName"))->SetItalic(true);
		rptGroup->GetItem(_T("Amount"))->SetItalic(true);
		TranslateString(_T("Total Amount"), tmpStr);
		rptGroup->SetValue(_T("PName"), tmpStr);
		FormatCurrency(nAmount, tmpStr);		
		rptGroup->SetValue(_T("Amount"), tmpStr);*/
		nTotalOrder ++;
		rs.MoveNext();
	}
	
	//Page Footer
	//Report Footer

	tmpStr.Format(_T("%ld"), nTotalOrder);
	rpt.GetReportFooter()->SetValue(_T("TotalOrder"), tmpStr);
	/*tmpStr.Format(_T("%.2f"), nTotalAmount);
	FormatCurrency(nTotalAmount, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);*/
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	//	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
	
	rpt.PrintPreview();
	EndWaitCursor();

}
void CEMPrescriptionListReport::OnExport()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	UpdateData(true);
	CString tmpStr, szSQL, szObjects, szWhere, tmpStr2, szObjectID, szObjectName,szRoomName,szRoomID, szListID, szTemp;
	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	CellFormat df(&xls), cf(&xls), hf(&xls);
	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER);
	cf.SetFontName(_T("Segoe UI"));

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 30);
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
	if (!m_szObjectKey.IsEmpty())
		szWhere.Format(_T(" AND ho_type in('%s') "), m_szObjectKey);	
	
	for(int i=0; i < m_wndList.GetItemCount(); i++)
	{
		if(m_wndList.GetCheck(i))
		{
			m_wndList.SetCurSel(i);
			if (!szListID.IsEmpty())
				szListID += _T(",");
			szListID.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i,0));

			if (!szRoomID.IsEmpty())
				szRoomID += _T(",");

			szRoomID.AppendFormat(_T("%s"), m_wndList.GetItemText(i,0));
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
	BeginWaitCursor();
	CString szDeptType, szWhereRoomID;
	szDeptType.Empty();
	szWhereRoomID.Empty();
	if (m_szCondition == _T("A"))
		szDeptType.IsEmpty();
	else
		if (m_szCondition == _T("O"))
		{
			szDeptType.Format(_T(" and hpo_depttype ='KB' "));
			if (!szListID.IsEmpty())				
				szWhereRoomID.Format(_T(" and hpo_roomid in(%s) "), szRoomID);
		}
		else
		{
			szDeptType.Format(_T(" and hpo_depttype ='DT' "));
			if(!szRoomID.IsEmpty())
				szWhereRoomID.Format(_T(" and hpo_deptid in(%s) ") , szListID);
		}
	if (!szDeptType.IsEmpty())
		szWhere.AppendFormat(_T("%s"), szDeptType);

	if(!m_szDoctorKey.IsEmpty())
		szWhere.AppendFormat(_T(" and hpo_doctor='%s' "), m_szDoctorKey);
	if (!m_szStockKey.IsEmpty())
		szWhere.AppendFormat(_T(" and hpo_storage_id=%d"), ToInt(m_szStockKey));
	if (m_nMinTotalAmt > 0)
		szWhere.AppendFormat(_T(" AND hpo_amount >= %ld"), m_nMinTotalAmt);
	if (!m_szApprovedByKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_approveby = '%s'"), m_szApprovedByKey);
	//Page Header
	//Report Detail
	szSQL.Format(_T(" SELECT 	hpo_orderid as orderid,") \
		_T(" 	hd_docno as docno,") \
		_T("	(select distinct su_name from sys_user where su_userid = hpo_doctor) as doctor,") \
		_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
		_T(" 	hms_getage((hd_admitdate), hp_birthdate) as age, ") \
		_T("	(select distinct su_name from sys_user where su_userid=hpo_doctor) as doctor, ")
		_T("	ho_type as objecttype ")
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_object ON(hd_object=ho_id) ")
		_T(" LEFT JOIN hms_pharmaorder ON(hpo_docno=hd_docno)") \
		_T(" WHERE 	hpo_orderstatus IN('A','I')") \
		_T("	and hpo_approvedate between cast_string2timestamp('%s') and cast_string2timestamp('%s') ")
		_T(" 	and hpo_ordertype in('P','D','M','E') %s %s ") \
		_T(" ORDER BY hpo_approvedate"), m_szFromDate, m_szToDate, szWhere, szWhereRoomID);
	//QueryOpener(szSQL);
	CRecord rs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	int nItem = 1;
	long nOrderID, nTotalOrder=0;
	long double nTotalAmount=0, nAmount;
	CString szPrice;
	szPrice = _T("product_taxprice");
	if (m_bServPrice)
		szPrice = _T("product_saleprice1");
	while(!rs.IsEOF())
	{	
		nRow++;
		nTotalOrder++;
		xls.SetCellMergedColumns(1,nRow,4);		
		rs.GetValue(_T("orderid"), nOrderID);
		tmpStr.Format(_T("%ld"), nOrderID);
		xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER,true,10,0);
		TranslateString(_T("Age"), tmpStr2);
		tmpStr.Format(_T("[%s]  -  %s   -   %s:%s  -  %s"), rs.GetValue(_T("docno")), rs.GetValue(_T("pname")),tmpStr2,rs.GetValue(_T("Age")), rs.GetValue(_T("doctor")));
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT,true,10,0);
		rs.GetValue(_T("doctor"), tmpStr);
		xls.SetCellText(5, nRow, tmpStr, FMT_TEXT,true,10,0);
		szSQL.Format(_T(" SELECT product_code as id, ") \
			_T(" 	product_name as name,") \
			_T(" 	get_uomname(product_purchase_uomid) as unit,") \
			_T(" 	%s as price,") \
			_T(" 	sum(hpol_qtyissue) as qty,") \
			_T(" 	sum(hpol_qtyissue)*%s as amount") \
			_T(" FROM hms_pharmaorderline ") \
			_T(" LEFT JOIN m_productitem_view ON(product_item_id=hpol_product_item_id)") \
			_T(" WHERE hpol_orderid=%ld") \
			_T(" GROUP BY product_code, product_name, product_purchase_uomid, %s"), szPrice, szPrice, nOrderID, szPrice);
		//QueryOpener(szSQL);
		rsl.ExecSQL(szSQL);
		nItem = 1;
		nAmount  =0;
		while(!rsl.IsEOF())
		{
			nRow++;
			tmpStr.Format(_T("%ld"), nItem++);			
			xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);
			rsl.GetValue(_T("id"), tmpStr);
			xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
			rsl.GetValue(_T("name"), tmpStr);
			xls.SetCellText(2, nRow, tmpStr, FMT_TEXT);
			rsl.GetValue(_T("unit"), tmpStr);
			xls.SetCellText(3, nRow, tmpStr, FMT_TEXT);			
			rsl.GetValue(_T("price"), tmpStr);
			xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
			rsl.GetValue(_T("qty"), tmpStr);
			xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1);
			rsl.GetValue(_T("amount"), tmpStr);
			nAmount += ToFloat(tmpStr);
			xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER1);
			rsl.MoveNext();
		}

		nTotalAmount += nAmount;
		nRow++;		
		TranslateString(_T("Total Amount"), tmpStr);
		xls.SetCellMergedColumns(1,nRow,3);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
		tmpStr.Format(_T("%.2f"), nAmount);
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);		
		rs.MoveNext();
	}
	
	//Page Footer
	//Report Footer
	nRow++;
	TranslateString(_T("Total"), tmpStr);
	xls.SetCellMergedColumns(1,nRow,3);
	xls.SetCellText(1, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
	tmpStr.Format(_T("%ld"), nTotalOrder);
	xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);
	nRow++;
	TranslateString(_T("Total Amount"), tmpStr);
	xls.SetCellMergedColumns(1,nRow,3);
	xls.SetCellText(1, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
	tmpStr.Format(_T("%.2f"), nTotalAmount);
	xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);


	xls.Save(_T("Exports\\Bang Ke Don Thuoc C1.3.xls"));
	EndWaitCursor();
}BEGIN_MESSAGE_MAP(CEMPrescriptionListReport, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CEMPrescriptionListReport::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
