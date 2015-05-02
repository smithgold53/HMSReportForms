#include "stdafx.h"
#include "PACSPatientListUseContrast.h"
#include "HMSMainFrame.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPACSPatientListUseContrast *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPACSPatientListUseContrast *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPACSPatientListUseContrast *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSPatientListUseContrast *)pWnd)->OnFromDateCheckValue();
} 
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPACSPatientListUseContrast *)pWnd)->OnStorageLoadData();
}
static void _OnOrderNoSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSPatientListUseContrast* )pWnd)->OnOrderNoSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnOrderNoSelendokFnc(CWnd *pWnd){
	((CPACSPatientListUseContrast *)pWnd)->OnOrderNoSelendok();
}
/*static void _OnOrderNoSetfocusFnc(CWnd *pWnd){
	((CPACSPatientListUseContrast *)pWnd)->OnOrderNoKillfocus();
}*/
/*static void _OnOrderNoKillfocusFnc(CWnd *pWnd){
	((CPACSPatientListUseContrast *)pWnd)->OnOrderNoKillfocus();
}*/
static long _OnOrderNoLoadDataFnc(CWnd *pWnd){
	return ((CPACSPatientListUseContrast *)pWnd)->OnOrderNoLoadData();
}
/*static void _OnOrderNoAddNewFnc(CWnd *pWnd){
	((CPACSPatientListUseContrast *)pWnd)->OnOrderNoAddNew();
}*/
static void _OnDrugSelectFnc(CWnd *pWnd){
	return ((CPACSPatientListUseContrast *)pWnd)->OnDrugSelect();
}
static void _OnMaterialSelectFnc(CWnd *pWnd){
	return ((CPACSPatientListUseContrast *)pWnd)->OnMaterialSelect();
}
static void _OnExportSelectFnc(CWnd *pWnd){
	CPACSPatientListUseContrast *pVw = (CPACSPatientListUseContrast *)pWnd;
	pVw->OnExportSelect();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPACSPatientListUseContrast *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPACSPatientListUseContrast *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPACSPatientListUseContrast *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSPatientListUseContrast *)pWnd)->OnToDateCheckValue();
} 
static int _OnAddPACSPatientListUseContrastFnc(CWnd *pWnd){
	 return ((CPACSPatientListUseContrast*)pWnd)->OnAddPACSPatientListUseContrast();
} 
static int _OnEditPACSPatientListUseContrastFnc(CWnd *pWnd){
	 return ((CPACSPatientListUseContrast*)pWnd)->OnEditPACSPatientListUseContrast();
} 
static int _OnDeletePACSPatientListUseContrastFnc(CWnd *pWnd){
	 return ((CPACSPatientListUseContrast*)pWnd)->OnDeletePACSPatientListUseContrast();
} 
static int _OnSavePACSPatientListUseContrastFnc(CWnd *pWnd){
	 return ((CPACSPatientListUseContrast*)pWnd)->OnSavePACSPatientListUseContrast();
} 
static int _OnCancelPACSPatientListUseContrastFnc(CWnd *pWnd){
	 return ((CPACSPatientListUseContrast*)pWnd)->OnCancelPACSPatientListUseContrast();
} 
CPACSPatientListUseContrast::CPACSPatientListUseContrast(CWnd *pParent){
	m_szType = _T("PM");
	m_nDlgWidth = 360;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CPACSPatientListUseContrast::~CPACSPatientListUseContrast(){
}
void CPACSPatientListUseContrast::OnCreateComponents(){
	m_wndReportParameters.Create(this, _T("Report Parameters"), 5, 5, 410, 120);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 205, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 210, 30, 290, 55);
	m_wndToDate.Create(this,295, 30, 405, 55); 
	m_wndStorageLabel.Create(this, _T("Storage"), 10, 60, 90, 85);
	m_wndStorage.Create(this,95, 60, 405, 85); 
	m_wndOrderNoLabel.Create(this, _T("Order No"), 10, 90, 90, 115);
	m_wndOrderNo.Create(this,95, 90, 405, 115); 
	m_wndExport.Create(this, _T("Export"), 330, 125, 410, 150);
	m_wndDrug.Create(this, _T("Drug"), 5, 125, 85, 150);
	m_wndMaterial.Create(this, _T("Material"), 88, 125, 168, 150);
	
}
void CPACSPatientListUseContrast::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndOrderNo.SetCheckValue(true);
	m_wndOrderNo.LimitText(35);
	m_wndToDate.SetCheckValue(true);

	m_wndStorage.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndStorage.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

	m_wndOrderNo.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndOrderNo.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CPACSPatientListUseContrast::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	m_wndStorage.SetEvent(WE_LOADDATA, _OnStorageLoadDataFnc);
	m_wndOrderNo.SetEvent(WE_SELENDOK, _OnOrderNoSelendokFnc);
	//m_wndOrderNo.SetEvent(WE_SETFOCUS, _OnOrderNoSetfocusFnc);
	//m_wndOrderNo.SetEvent(WE_KILLFOCUS, _OnOrderNoKillfocusFnc);
	m_wndOrderNo.SetEvent(WE_SELCHANGE, _OnOrderNoSelectChangeFnc);
	m_wndOrderNo.SetEvent(WE_LOADDATA, _OnOrderNoLoadDataFnc);
	//m_wndOrderNo.SetEvent(WE_ADDNEW, _OnOrderNoAddNewFnc);
	m_wndDrug.SetEvent(WE_CLICK, _OnDrugSelectFnc);
	m_wndMaterial.SetEvent(WE_CLICK, _OnMaterialSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	OnStorageLoadData();
	SetMode(VM_EDIT);

}
void CPACSPatientListUseContrast::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStorage.GetDlgCtrlID(), m_szStorageKey);
	DDX_TextEx(pDX, m_wndOrderNo.GetDlgCtrlID(), m_szOrderNoKey);
	DDX_Radio(pDX, m_wndDrug.GetDlgCtrlID(), m_nDrug);
	DDX_Radio(pDX, m_wndMaterial.GetDlgCtrlID(), m_nMaterial);

}
void CPACSPatientListUseContrast::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPACSPatientListUseContrast::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CPACSPatientListUseContrast::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szOrderNoKey.Empty();
	m_szStorageKey.Empty();
	m_szToDate.Empty();
	m_nDrug = 0;
	m_nMaterial = 1;
}
int CPACSPatientListUseContrast::SetMode(int nMode){
 		int nOldMode = GetMode();
 		CGuiView::SetMode(nMode);
 		CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 		CString szSQL;
 		CRecord rs(&pMF->m_db);
  		switch(nMode){
 		case VM_ADD: 
 			EnableControls(TRUE);
 			EnableButtons(TRUE, 0, 1, -1);
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
/*void CPACSPatientListUseContrast::OnFromDateChange(){
	
} */
/*void CPACSPatientListUseContrast::OnFromDateSetfocus(){
	
} */
/*void CPACSPatientListUseContrast::OnFromDateKillfocus(){
	
} */
int CPACSPatientListUseContrast::OnFromDateCheckValue(){
	return 0;
} 
/*void CPACSPatientListUseContrast::OnToDateChange(){
	
} */
/*void CPACSPatientListUseContrast::OnToDateSetfocus(){
	
} */
/*void CPACSPatientListUseContrast::OnToDateKillfocus(){
	
} */
int CPACSPatientListUseContrast::OnToDateCheckValue(){
	return 0;
} 

long CPACSPatientListUseContrast::OnStorageLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	int nCount = 0;
	m_wndStorage.DeleteAllItems();
	szSQL.Format(_T("SELECT msl_storage_id id, msl_name name FROM m_storagelist WHERE msl_dept_id = 'C8' ORDER BY msl_storage_id"));
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		m_wndStorage.AddItems(
			rs.GetValue(_T("id")),
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}

void CPACSPatientListUseContrast::OnOrderNoSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSPatientListUseContrast::OnOrderNoSelendok(){
	 
}
/*void CPACSPatientListUseContrast::OnOrderNoSetfocus(){
	
}*/
/*void CPACSPatientListUseContrast::OnOrderNoKillfocus(){
	
}*/
long CPACSPatientListUseContrast::OnOrderNoLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	UpdateData(TRUE);
	if(m_wndOrderNo.IsSearchKey() && !m_szOrderNoKey.IsEmpty()){
		szWhere.Format(_T(" and mt_transaction_id = %s "), m_szOrderNoKey);
	}
	m_wndOrderNo.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT mt_transaction_id as id, mt_orderno as orderno FROM m_transaction ") \
				_T(" WHERE mt_doctype = 'CSO' AND mt_department_to_id = 'C8' AND mt_status <> 'C' ") \
				_T(" AND mt_orderdate BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') %s"),m_szFromDate, m_szToDate, szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndOrderNo.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("orderno")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPACSPatientListUseContrast::OnOrderNoAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */

void CPACSPatientListUseContrast::OnDrugSelect(){
	m_szType = _T("PM");
}

void CPACSPatientListUseContrast::OnMaterialSelect(){
	m_szType = _T("MA");
}


void CPACSPatientListUseContrast::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
		if (nCount < 0)
			_msg(_T("%s"), szSQL);
		else
			_fmsg(_T("%s"), szSQL);
		return;
} 
int CPACSPatientListUseContrast::OnAddPACSPatientListUseContrast(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CPACSPatientListUseContrast::OnEditPACSPatientListUseContrast(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CPACSPatientListUseContrast::OnDeletePACSPatientListUseContrast(){
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
 		OnCancelPACSPatientListUseContrast();
 	}
	return 0;
}
int CPACSPatientListUseContrast::OnSavePACSPatientListUseContrast(){
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
 		//OnPACSPatientListUseContrastListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CPACSPatientListUseContrast::OnCancelPACSPatientListUseContrast(){
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
int CPACSPatientListUseContrast::OnPACSPatientListUseContrastListLoadData(){
	return 0;
}

CString CPACSPatientListUseContrast::GetQueryString(){
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szOrderNoKey.IsEmpty())
		szWhere.Format(_T("  AND hpo_orderid IN (SELECT DISTINCT hmt_orderid ") \
					   _T("  FROM   hms_medical_transaction_view ") \
					   _T("  WHERE  hmt_reftransaction_id = %s)"), m_szOrderNoKey);
	if (!m_szStorageKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_storage_id = %s"), m_szStorageKey);
	szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), m_szType);
	szSQL.Format(_T(" SELECT    hmt_product_id, ") \
				_T("           (SELECT hrl_name ") \
				_T("            FROM   hms_roomlist ") \
				_T("            WHERE  hrl_deptid = hpo_deptid AND hrl_id = hpo_roomid) AS tenphong, ") \
				_T("           hpo_deptid, ") \
				_T("           hpo_roomid, ") \
				_T("           SUM(hmt_qtyissue) qty, ") \
				_T("           SUM(CASE WHEN hpo_object_id = 1 THEN hmt_qtyissue ") \
				_T("               ELSE 0 ") \
				_T("               END) sol_qty, ") \
				_T("           SUM(CASE WHEN hpo_object_id = 2 THEN hmt_qtyissue ") \
				_T("               ELSE 0 ") \
				_T("               END) solins_qty, ") \
				_T("           SUM(CASE WHEN hpo_object_id = 4 THEN hmt_qtyissue ") \
				_T("               ELSE 0 ") \
				_T("               END) otherins_qty, ") \
				_T("           SUM(CASE WHEN hpo_object_id = 3 THEN hmt_qtyissue ") \
				_T("               ELSE 0 ") \
				_T("               END) pol_qty, ") \
				_T("           SUM(CASE WHEN hpo_object_id = 7 THEN hmt_qtyissue ") \
				_T("               ELSE 0 ") \
				_T("               END) service_qty, ") \
				_T("           SUM(CASE WHEN hpo_object_id NOT IN ( 1, 2, 3, 4, 7 ) THEN hmt_qtyissue ") \
				_T("               ELSE 0 ") \
				_T("               END) other_qty, ") \
				_T("           hpo_doctor, ") \
				_T("           Get_patientname(hpc_docno) patientname, ") \
				_T("           product_name tenthuoc, ") \
				_T("           CASE WHEN hpc_class = 'E' THEN he_diagnostic ") \
				_T("           ELSE htr_maindisease ") \
				_T("           END diagnostic, ") \
				_T("           hfl_name assignment, ") \
				_T("           CASE WHEN hpc_class = 'E' THEN hpc_deptid ") \
				_T("                                     || '/P' ") \
				_T("                                     || hpc_roomid ") \
				_T("           ELSE hpc_deptid ") \
				_T("           END dept, ") \
				_T("           hpo_orderdate order_date ") \
				_T(" FROM      hms_ipharmaorder ") \
				_T(" LEFT JOIN hms_medical_transaction_view ON ( hpo_orderid = hmt_orderid ) ") \
				_T(" LEFT JOIN m_transaction ON ( mt_transaction_id = hmt_reftransaction_id ) ") \
				_T(" LEFT JOIN hms_fee_list ON ( hfl_feeid = hpo_refitem_id ) ") \
				_T(" LEFT JOIN hms_pacsorder ON ( hpo_reference_id = hpc_orderid AND hpc_docno = hpo_docno ) ") \
				_T(" LEFT JOIN m_product_view ON ( product_id = hmt_product_id ) ") \
				_T(" LEFT JOIN hms_treatment_record ON ( htr_docno = hpc_docno AND htr_idx = hpc_refidx AND htr_deptid = hpc_deptid ) ") \
				_T(" LEFT JOIN hms_exam ON ( he_docno = hpc_docno AND he_roomid = hpc_roomid AND he_deptid = hpc_deptid ) ") \
				_T(" WHERE     hpo_deptid = 'C8' ") \
				_T(" %s") \
				_T(" GROUP     BY hmt_product_id,hpo_deptid,hpo_roomid,hpo_doctor,hpo_orderdate,hpc_docno,product_name,") \
				_T(" he_diagnostic,htr_maindisease,hfl_name,hpc_deptid,hpc_roomid,hpc_class ") \
				_T(" ORDER     BY product_name,hpo_roomid,hpc_docno "), szWhere);

	return szSQL;
}