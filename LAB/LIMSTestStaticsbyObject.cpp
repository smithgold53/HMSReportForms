#include "stdafx.h"
#include "LIMSTestStaticsbyObject.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
static void _OnDepartmentSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CLIMSTestStaticsbyObject* )pWnd)->OnDepartmentSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDepartmentSelendokFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnDepartmentSelendok();
}
/*static void _OnDepartmentSetfocusFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnDepartmentKillfocus();
}*/
/*static void _OnDepartmentKillfocusFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnDepartmentKillfocus();
}*/
static long _OnDepartmentLoadDataFnc(CWnd *pWnd){
	return ((CLIMSTestStaticsbyObject *)pWnd)->OnDepartmentLoadData();
}
/*static void _OnDepartmentAddNewFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnDepartmentAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CLIMSTestStaticsbyObject *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CLIMSTestStaticsbyObject *)pWnd)->OnToDateCheckValue();
} 
static void _OnMachineNoSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CLIMSTestStaticsbyObject* )pWnd)->OnMachineNoSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnMachineNoSelendokFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnMachineNoSelendok();
}
/*static void _OnMachineNoSetfocusFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnMachineNoKillfocus();
}*/
/*static void _OnMachineNoKillfocusFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnMachineNoKillfocus();
}*/
static long _OnMachineNoLoadDataFnc(CWnd *pWnd){
	return ((CLIMSTestStaticsbyObject *)pWnd)->OnMachineNoLoadData();
}
/*static void _OnMachineNoAddNewFnc(CWnd *pWnd){
	((CLIMSTestStaticsbyObject *)pWnd)->OnMachineNoAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CLIMSTestStaticsbyObject *pVw = (CLIMSTestStaticsbyObject *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CLIMSTestStaticsbyObject *pVw = (CLIMSTestStaticsbyObject *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddLIMSTestStaticsbyObjectFnc(CWnd *pWnd){
	 return ((CLIMSTestStaticsbyObject*)pWnd)->OnAddLIMSTestStaticsbyObject();
} 
static int _OnEditLIMSTestStaticsbyObjectFnc(CWnd *pWnd){
	 return ((CLIMSTestStaticsbyObject*)pWnd)->OnEditLIMSTestStaticsbyObject();
} 
static int _OnDeleteLIMSTestStaticsbyObjectFnc(CWnd *pWnd){
	 return ((CLIMSTestStaticsbyObject*)pWnd)->OnDeleteLIMSTestStaticsbyObject();
} 
static int _OnSaveLIMSTestStaticsbyObjectFnc(CWnd *pWnd){
	 return ((CLIMSTestStaticsbyObject*)pWnd)->OnSaveLIMSTestStaticsbyObject();
} 
static int _OnCancelLIMSTestStaticsbyObjectFnc(CWnd *pWnd){
	 return ((CLIMSTestStaticsbyObject*)pWnd)->OnCancelLIMSTestStaticsbyObject();
} 
CLIMSTestStaticsbyObject::CLIMSTestStaticsbyObject(CWnd* pParent){

	m_nDlgWidth = 400;
	m_nDlgHeight = 160;
	SetDefaultValues();
}
CLIMSTestStaticsbyObject::~CLIMSTestStaticsbyObject(){
}
void CLIMSTestStaticsbyObject::OnCreateComponents(){
	m_wndReportParameters.Create(this, _T("Report Parameters"), 5, 5, 390, 120);
	m_wndDepartmentLabel.Create(this, _T("Department"), 10, 30, 90, 55);
	m_wndDepartment.Create(this,95, 30, 385, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 195, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 200, 60, 280, 85);
	m_wndToDate.Create(this,285, 60, 385, 85); 
	m_wndMachineNoLabel.Create(this, _T("Machine No"), 10, 90, 90, 115);
	m_wndMachineNo.Create(this,95, 90, 385, 115); 
	m_wndPrint.Create(this, _T("&Print"), 235, 125, 310, 150);
	m_wndExport.Create(this, _T("&Export"), 315, 125, 390, 150);

}
void CLIMSTestStaticsbyObject::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(FALSE);
	//EnableButtons(TRUE, 0, -1);
	m_wndDepartment.SetCheckValue(true);
	m_wndDepartment.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	m_wndMachineNo.SetCheckValue(true);
	m_wndMachineNo.LimitText(35);


	m_wndDepartment.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndDepartment.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndMachineNo.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndMachineNo.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

}
void CLIMSTestStaticsbyObject::OnSetWindowEvents(){
	m_wndDepartment.SetEvent(WE_SELENDOK, _OnDepartmentSelendokFnc);
	//m_wndDepartment.SetEvent(WE_SETFOCUS, _OnDepartmentSetfocusFnc);
	//m_wndDepartment.SetEvent(WE_KILLFOCUS, _OnDepartmentKillfocusFnc);
	m_wndDepartment.SetEvent(WE_SELCHANGE, _OnDepartmentSelectChangeFnc);
	m_wndDepartment.SetEvent(WE_LOADDATA, _OnDepartmentLoadDataFnc);
	//m_wndDepartment.SetEvent(WE_ADDNEW, _OnDepartmentAddNewFnc);
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndMachineNo.SetEvent(WE_SELENDOK, _OnMachineNoSelendokFnc);
	//m_wndMachineNo.SetEvent(WE_SETFOCUS, _OnMachineNoSetfocusFnc);
	//m_wndMachineNo.SetEvent(WE_KILLFOCUS, _OnMachineNoKillfocusFnc);
	m_wndMachineNo.SetEvent(WE_SELCHANGE, _OnMachineNoSelectChangeFnc);
	m_wndMachineNo.SetEvent(WE_LOADDATA, _OnMachineNoLoadDataFnc);
	//m_wndMachineNo.SetEvent(WE_ADDNEW, _OnMachineNoAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	/*AddEvent(1, _T("Add	Ctrl+A"), _OnAddLIMSTestStaticsbyObjectFnc, 0, 'A', VK_CONTROL);
	AddEvent(2, _T("Edit	Ctrl+E"), _OnEditLIMSTestStaticsbyObjectFnc, 0, 'E', VK_CONTROL);
	AddEvent(3, _T("Delete	Ctrl+D"), _OnDeleteLIMSTestStaticsbyObjectFnc, 0, 'D', VK_CONTROL);
	AddEvent(4, _T("Save	Ctrl+S"), _OnSaveLIMSTestStaticsbyObjectFnc, 0, 'S', VK_CONTROL);
	AddEvent(0, _T("-"));
	AddEvent(5, _T("Cancel	Ctrl+T"), _OnCancelLIMSTestStaticsbyObjectFnc, 0, 'T', VK_CONTROL);*/
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	UpdateData(false);
}
void CLIMSTestStaticsbyObject::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndDepartment.GetDlgCtrlID(), m_szDepartmentKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndMachineNo.GetDlgCtrlID(), m_szMachineNoKey);

}
void CLIMSTestStaticsbyObject::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CLIMSTestStaticsbyObject::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CLIMSTestStaticsbyObject::SetDefaultValues(){

	m_szDepartmentKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szMachineNoKey.Empty();

}
int CLIMSTestStaticsbyObject::SetMode(int nMode){
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
void CLIMSTestStaticsbyObject::OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CLIMSTestStaticsbyObject::OnDepartmentSelendok(){
	 
}
/*void CLIMSTestStaticsbyObject::OnDepartmentSetfocus(){
	
}*/
/*void CLIMSTestStaticsbyObject::OnDepartmentKillfocus(){
	
}*/
long CLIMSTestStaticsbyObject::OnDepartmentLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDepartment.IsSearchKey() && !m_szDepartmentKey.IsEmpty()){
	};
	szSQL.Format(_T("SELECT sd_id as id, sd_name as description FROM sys_dept WHERE sd_id IN ('C2', 'C3') ORDER BY sd_id"));
	m_wndDepartment.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDepartment.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CLIMSTestStaticsbyObject::OnDepartmentAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CLIMSTestStaticsbyObject::OnFromDateChange(){
	
} */
/*void CLIMSTestStaticsbyObject::OnFromDateSetfocus(){
	
} */
/*void CLIMSTestStaticsbyObject::OnFromDateKillfocus(){
	
} */
int CLIMSTestStaticsbyObject::OnFromDateCheckValue(){
	return 0;
} 
/*void CLIMSTestStaticsbyObject::OnToDateChange(){
	
} */
/*void CLIMSTestStaticsbyObject::OnToDateSetfocus(){
	
} */
/*void CLIMSTestStaticsbyObject::OnToDateKillfocus(){
	
} */
int CLIMSTestStaticsbyObject::OnToDateCheckValue(){
	return 0;
} 
void CLIMSTestStaticsbyObject::OnMachineNoSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CLIMSTestStaticsbyObject::OnMachineNoSelendok(){
	 
}
/*void CLIMSTestStaticsbyObject::OnMachineNoSetfocus(){
	
}*/
/*void CLIMSTestStaticsbyObject::OnMachineNoKillfocus(){
	
}*/
long CLIMSTestStaticsbyObject::OnMachineNoLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndMachineNo.IsSearchKey() && !m_szMachineNoKey.IsEmpty()){
	};
	m_wndMachineNo.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndMachineNo.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CLIMSTestStaticsbyObject::OnMachineNoAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CLIMSTestStaticsbyObject::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CLIMSTestStaticsbyObject::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CLIMSTestStaticsbyObject::OnAddLIMSTestStaticsbyObject(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CLIMSTestStaticsbyObject::OnEditLIMSTestStaticsbyObject(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CLIMSTestStaticsbyObject::OnDeleteLIMSTestStaticsbyObject(){
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
 		OnCancelLIMSTestStaticsbyObject(); 
 	}
	return 0;
}
int CLIMSTestStaticsbyObject::OnSaveLIMSTestStaticsbyObject(){
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
 		//OnLIMSTestStaticsbyObjectListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CLIMSTestStaticsbyObject::OnCancelLIMSTestStaticsbyObject(){
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
int CLIMSTestStaticsbyObject::OnLIMSTestStaticsbyObjectListLoadData(){
	return 0;
}
CString CLIMSTestStaticsbyObject::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND (limso_performdate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szMachineNoKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND limso_instid= %d"), ToInt(m_szMachineNoKey));
	szSQL.Format(_T(" SELECT SUM(tongso) AS tongso,") \
				_T("          SUM(quan)        AS quan,") \
				_T("          SUM(testquan)    AS testquan,") \
				_T("          SUM(bhquan)      AS bhquan,") \
				_T("          SUM(testbhquan)      AS testbhquan,") \
				_T("          SUM(bhquandoi)   AS bhquandoi,") \
				_T("          SUM(testbhquandoi)   AS testbhquandoi,") \
				_T("          SUM(bhthannhan)  AS bhthannhan,") \
				_T("          SUM(testbhthannhan)  AS testbhthannhan,") \
				_T("          SUM(bhkhac)      AS bhkhac,") \
				_T("          SUM(testbhkhac)      AS testbhkhac,") \
				_T("          SUM(treem)       AS treem,") \
				_T("          SUM(testtreem)       AS testtreem,") \
				_T("          SUM(treemcapthe) AS treemcapthe,") \
				_T("          SUM(testtreemcapthe) AS testtreemcapthe,") \
				_T("          SUM(chinhsach)   AS chinhsach,") \
				_T("          SUM(testchinhsach)   AS testchinhsach,") \
				_T("          SUM(dichvu)      AS dichvu,") \
				_T("          SUM(testdichvu)      AS testdichvu,") \
				_T("          SUM(ban)         AS ban,") \
				_T("          SUM(testban)         AS testban") \
				_T(" FROM") \
				_T("  (SELECT limso_docno AS docno,") \
				_T("    1                 AS tongso,") \
				_T("    CASE WHEN limsp_object = '1' THEN 1 ELSE 0 END AS quan,") \
				_T("    CASE WHEN limsp_object = '1' THEN count(*) ELSE 0 END AS testquan,") \
				_T("    CASE WHEN limsp_object = '2' THEN 1 ELSE 0 END AS bhquan,") \
				_T("    CASE WHEN limsp_object = '2' THEN count(*) ELSE 0 END AS testbhquan,") \
				_T("    CASE WHEN limsp_object = '3' THEN 1 ELSE 0 END AS chinhsach,") \
				_T("    CASE WHEN limsp_object = '3' THEN count(*) ELSE 0 END AS testchinhsach,") \
				_T("    CASE WHEN limsp_object IN('4') THEN 1 ELSE 0 END AS bhkhac,") \
				_T("    CASE WHEN limsp_object IN('4') THEN count(*) ELSE 0 END AS testbhkhac,") \
				_T("    CASE WHEN limsp_object IN('5') THEN 1 ELSE 0 END AS bhthannhan,") \
				_T("    CASE WHEN limsp_object IN('5') THEN count(*) ELSE 0 END AS testbhthannhan,") \
				_T("    CASE WHEN limsp_object IN('6') THEN 1 ELSE 0 END AS treem,") \
				_T("    CASE WHEN limsp_object IN('6') THEN count(*) ELSE 0 END AS testtreem,") \
				_T("    CASE WHEN limsp_object NOT IN('1','2','3','4','5','6','8','9','10') THEN 1 ELSE 0 END AS dichvu,") \
				_T("    CASE WHEN limsp_object NOT IN('1','2','3','4','5','6','8','9','10') THEN count(*) ELSE 0 END AS testdichvu,") \
				_T("    CASE WHEN limsp_object IN('8') THEN 1 ELSE 0 END AS ban,") \
				_T("    CASE WHEN limsp_object IN('8') THEN count(*) ELSE 0 END AS testban,") \
				_T("    CASE WHEN limsp_object IN('9') THEN 1 ELSE 0 END AS treemcapthe,") \
				_T("    CASE WHEN limsp_object IN('9') THEN count(*) ELSE 0 END AS testtreemcapthe,") \
				_T("    CASE WHEN limsp_object = '10' THEN 1 ELSE 0 END AS bhquandoi,") \
				_T("    CASE WHEN limsp_object = '10' THEN count(*) ELSE 0 END AS testbhquandoi") \
				_T("  FROM lims_patient") \
				_T("  LEFT JOIN lims_order ON(limso_docno    =limsp_docno)") \
				_T("  LEFT JOIN lims_order_line ON (limso_oid= limsol_oid)") \
				_T("  WHERE limso_status IN('P','U') %s") \
				_T("  GROUP BY limso_docno, limso_oid, limsp_object") \
				_T(" ) tbl"), szWhere);
	return szSQL;
}
