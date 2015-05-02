#include "stdafx.h"
#include "PACSSettlementChemistryMaterial.h"
#include "HMSMainFrame.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSSettlementChemistryMaterial *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSSettlementChemistryMaterial *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSSettlementChemistryMaterial* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CPACSSettlementChemistryMaterial *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnStockAddNew();
}*/
static void _OnInPatientSelectFnc(CWnd *pWnd){
	  ((CPACSSettlementChemistryMaterial*)pWnd)->OnInPatientSelect();
}
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	  ((CPACSSettlementChemistryMaterial*)pWnd)->OnOutPatientSelect();
}
static void _OnExportSelectFnc(CWnd *pWnd){
	CPACSSettlementChemistryMaterial *pVw = (CPACSSettlementChemistryMaterial *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddPACSSettlementChemistryMaterialFnc(CWnd *pWnd){
	 return ((CPACSSettlementChemistryMaterial*)pWnd)->OnAddPACSSettlementChemistryMaterial();
} 
static int _OnEditPACSSettlementChemistryMaterialFnc(CWnd *pWnd){
	 return ((CPACSSettlementChemistryMaterial*)pWnd)->OnEditPACSSettlementChemistryMaterial();
} 
static int _OnDeletePACSSettlementChemistryMaterialFnc(CWnd *pWnd){
	 return ((CPACSSettlementChemistryMaterial*)pWnd)->OnDeletePACSSettlementChemistryMaterial();
} 
static int _OnSavePACSSettlementChemistryMaterialFnc(CWnd *pWnd){
	 return ((CPACSSettlementChemistryMaterial*)pWnd)->OnSavePACSSettlementChemistryMaterial();
} 
static int _OnCancelPACSSettlementChemistryMaterialFnc(CWnd *pWnd){
	 return ((CPACSSettlementChemistryMaterial*)pWnd)->OnCancelPACSSettlementChemistryMaterial();
} 
CPACSSettlementChemistryMaterial::CPACSSettlementChemistryMaterial(CWnd *pParent){

	m_nDlgWidth = 420;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CPACSSettlementChemistryMaterial::~CPACSSettlementChemistryMaterial(){
}
void CPACSSettlementChemistryMaterial::OnCreateComponents(){
	m_wndReportParameter.Create(this, _T("Report Parameters"), 5, 5, 410, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 205, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 210, 30, 290, 55);
	m_wndToDate.Create(this,295, 30, 405, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 90, 85);
	m_wndStock.Create(this,95, 60, 205, 85); 
	m_wndInPatient.Create(this, _T("InPatient"), 210, 60, 290, 85);
	m_wndOutPatient.Create(this, _T("OutPatient"), 296, 60, 376, 85);
	m_wndExport.Create(this, _T("Export"), 330, 95, 410, 120);

}
void CPACSSettlementChemistryMaterial::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);


	m_wndStock.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndStock.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CPACSSettlementChemistryMaterial::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	m_wndStock.SetCheckBox(true);
	SetMode(VM_EDIT);
}
void CPACSSettlementChemistryMaterial::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_Radio(pDX, m_wndInPatient.GetDlgCtrlID(), m_nInPatient);
	DDX_Radio(pDX, m_wndOutPatient.GetDlgCtrlID(), m_nOutPatient);

}
void CPACSSettlementChemistryMaterial::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPACSSettlementChemistryMaterial::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CPACSSettlementChemistryMaterial::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_nInPatient=0;
	m_nOutPatient=0;

}
int CPACSSettlementChemistryMaterial::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, -1);
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
/*void CPACSSettlementChemistryMaterial::OnFromDateChange(){
	
} */
/*void CPACSSettlementChemistryMaterial::OnFromDateSetfocus(){
	
} */
/*void CPACSSettlementChemistryMaterial::OnFromDateKillfocus(){
	
} */
int CPACSSettlementChemistryMaterial::OnFromDateCheckValue(){
	return 0;
} 
/*void CPACSSettlementChemistryMaterial::OnToDateChange(){
	
} */
/*void CPACSSettlementChemistryMaterial::OnToDateSetfocus(){
	
} */
/*void CPACSSettlementChemistryMaterial::OnToDateKillfocus(){
	
} */
int CPACSSettlementChemistryMaterial::OnToDateCheckValue(){
	return 0;
} 
void CPACSSettlementChemistryMaterial::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSSettlementChemistryMaterial::OnStockSelendok(){
	 
}
/*void CPACSSettlementChemistryMaterial::OnStockSetfocus(){
	
}*/
/*void CPACSSettlementChemistryMaterial::OnStockKillfocus(){
	
}*/
long CPACSSettlementChemistryMaterial::OnStockLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStock.IsSearchKey() && !m_szStockKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szStockKey
};
	m_wndStock.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStock.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CPACSSettlementChemistryMaterial::OnStockAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CPACSSettlementChemistryMaterial::OnInPatientSelect(){
	
}
void CPACSSettlementChemistryMaterial::OnOutPatientSelect(){
	
}
void CPACSSettlementChemistryMaterial::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
} 
int CPACSSettlementChemistryMaterial::OnAddPACSSettlementChemistryMaterial(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CPACSSettlementChemistryMaterial::OnEditPACSSettlementChemistryMaterial(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CPACSSettlementChemistryMaterial::OnDeletePACSSettlementChemistryMaterial(){
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
 		OnCancelPACSSettlementChemistryMaterial();
 	}
	return 0;
}
int CPACSSettlementChemistryMaterial::OnSavePACSSettlementChemistryMaterial(){
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
 		//OnPACSSettlementChemistryMaterialListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CPACSSettlementChemistryMaterial::OnCancelPACSSettlementChemistryMaterial(){
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
int CPACSSettlementChemistryMaterial::OnPACSSettlementChemistryMaterialListLoadData(){
	return 0;
}