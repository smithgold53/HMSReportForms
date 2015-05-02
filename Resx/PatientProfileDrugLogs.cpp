#include "PatientProfileDrugLogs.h"
#include "MainFrm.h"
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CPatientProfileDrugLogs*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CPatientProfileDrugLogs*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPatientProfileDrugLogs*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CPatientProfileDrugLogs*)pWnd)->OnListDeleteItem();
} 
static int _OnAddPatientProfileDrugLogsFnc(CWnd *pWnd){
	 return ((CPatientProfileDrugLogs*)pWnd)->OnAddPatientProfileDrugLogs();
} 
static int _OnEditPatientProfileDrugLogsFnc(CWnd *pWnd){
	 return ((CPatientProfileDrugLogs*)pWnd)->OnEditPatientProfileDrugLogs();
} 
static int _OnDeletePatientProfileDrugLogsFnc(CWnd *pWnd){
	 return ((CPatientProfileDrugLogs*)pWnd)->OnDeletePatientProfileDrugLogs();
} 
static int _OnSavePatientProfileDrugLogsFnc(CWnd *pWnd){
	 return ((CPatientProfileDrugLogs*)pWnd)->OnSavePatientProfileDrugLogs();
} 
static int _OnCancelPatientProfileDrugLogsFnc(CWnd *pWnd){
	 return ((CPatientProfileDrugLogs*)pWnd)->OnCancelPatientProfileDrugLogs();
} 
CPatientProfileDrugLogs::CPatientProfileDrugLogs(){

	m_nDlgWidth = 765;
	m_nDlgHeight = 585;
	SetDefaultValues();
}
CPatientProfileDrugLogs::~CPatientProfileDrugLogs(){
}
void CPatientProfileDrugLogs::OnCreateComponents(){
	m_wndDrugLogsInformation.Create(this, _T("DrugLogsInformation"), 5, 5, 760, 580);
	m_wndList.Create(this,10, 30, 755, 570); 

}
void CPatientProfileDrugLogs::OnInitializeComponents(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();


	m_wndList.InsertColumn(0, _T("Idx"), CFMT_NUMBER, 35);
	m_wndList.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

}
void CPatientProfileDrugLogs::OnSetWindowEvents(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);
	m_wndList.AddEvent(1, _T("Delete"), _OnListDeleteItemFnc, 0, VK_DELETE, 0);
	AddEvent(1, _T("Add	Ctrl+A"), _OnAddPatientProfileDrugLogsFnc, 0, 'A', VK_CONTROL);
	AddEvent(2, _T("Edit	Ctrl+E"), _OnEditPatientProfileDrugLogsFnc, 0, 'E', VK_CONTROL);
	AddEvent(3, _T("Delete	Ctrl+D"), _OnDeletePatientProfileDrugLogsFnc, 0, 'D', VK_CONTROL);
	AddEvent(4, _T("Save	Ctrl+S"), _OnSavePatientProfileDrugLogsFnc, 0, 'S', VK_CONTROL);
	AddEvent(0, _T("-"));
	AddEvent(5, _T("Cancel	Ctrl+T"), _OnCancelPatientProfileDrugLogsFnc, 0, 'T', VK_CONTROL);
	SetMode(VM_NONE);

}
void CPatientProfileDrugLogs::OnDoDataExchange(CDataExchange* pDX){

}
void CPatientProfileDrugLogs::GetDataToScreen(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPatientProfileDrugLogs::GetScreenToData(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();

}
void CPatientProfileDrugLogs::SetDefaultValues(){


}
int CPatientProfileDrugLogs::SetMode(int nMode){
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
void CPatientProfileDrugLogs::OnListDblClick(){
	
} 
void CPatientProfileDrugLogs::OnListSelectChange(int nOldItem, int nNewItem){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
int CPatientProfileDrugLogs::OnListDeleteItem(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CPatientProfileDrugLogs::OnListLoadData(){
/*
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndList.BeginLoad(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndList.AddItems(
			rs.GetValue(_T("Idx")), 
			rs.GetValue(_T("Description")), NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
*/
	return 0;
}
int CPatientProfileDrugLogs::OnAddPatientProfileDrugLogs(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CMainFrame *pMF = (CMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CPatientProfileDrugLogs::OnEditPatientProfileDrugLogs(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CMainFrame *pMF = (CMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CPatientProfileDrugLogs::OnDeletePatientProfileDrugLogs(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CMainFrame *pMF = (CMainFrame *)AfxGetMainWnd();
 	CString szSQL;
 	if(ShowMessage(1, MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2) == IDNO)
 		return -1;
 	szSQL.Format(_T("DELETE FROM  WHERE  AND") );
 	int ret = pMF->ExecSQL(szSQL);
 	if(ret >= 0){
 		SetMode(VM_NONE);
 		OnCancelPatientProfileDrugLogs();
 	}
	return 0;
}
int CPatientProfileDrugLogs::OnSavePatientProfileDrugLogs(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT)
 		return -1;
 	if(!IsValidateData())
 		return -1;
 	CMainFrame *pMF = (CMainFrame *)AfxGetMainWnd();
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
 		//OnPatientProfileDrugLogsListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CPatientProfileDrugLogs::OnCancelPatientProfileDrugLogs(){
 	if(GetMode() == VM_EDIT)
 	{
 		SetMode(VM_VIEW);
 	} 
 	else 
 	{
 		SetMode(VM_NONE);
 	} 
 	CMainFrame *pMF = (CMainFrame *)AfxGetMainWnd();
 	return 0;
} 
int CPatientProfileDrugLogs::OnPatientProfileDrugLogsListLoadData(){
	return 0;
}
