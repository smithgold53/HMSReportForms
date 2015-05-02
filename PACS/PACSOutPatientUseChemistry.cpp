#include "stdafx.h"
#include "PACSOutPatientUseChemistry.h"
#include "HMSMainFrame.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPACSOutPatientUseChemistry *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPACSOutPatientUseChemistry *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPACSOutPatientUseChemistry *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSOutPatientUseChemistry *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPACSOutPatientUseChemistry *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPACSOutPatientUseChemistry *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPACSOutPatientUseChemistry *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSOutPatientUseChemistry *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSOutPatientUseChemistry* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CPACSOutPatientUseChemistry *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CPACSOutPatientUseChemistry *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CPACSOutPatientUseChemistry *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CPACSOutPatientUseChemistry *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CPACSOutPatientUseChemistry *)pWnd)->OnStockAddNew();
}*/
static void _OnExportSelectFnc(CWnd *pWnd){
	CPACSOutPatientUseChemistry *pVw = (CPACSOutPatientUseChemistry *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnDrugSelectFnc(CWnd *pWnd){
	  ((CPACSOutPatientUseChemistry*)pWnd)->OnDrugSelect();
}
static void _OnMaterialSelectFnc(CWnd *pWnd){
	  ((CPACSOutPatientUseChemistry*)pWnd)->OnMaterialSelect();
}
static int _OnAddPACSOutPatientUseChemistryFnc(CWnd *pWnd){
	 return ((CPACSOutPatientUseChemistry*)pWnd)->OnAddPACSOutPatientUseChemistry();
} 
static int _OnEditPACSOutPatientUseChemistryFnc(CWnd *pWnd){
	 return ((CPACSOutPatientUseChemistry*)pWnd)->OnEditPACSOutPatientUseChemistry();
} 
static int _OnDeletePACSOutPatientUseChemistryFnc(CWnd *pWnd){
	 return ((CPACSOutPatientUseChemistry*)pWnd)->OnDeletePACSOutPatientUseChemistry();
} 
static int _OnSavePACSOutPatientUseChemistryFnc(CWnd *pWnd){
	 return ((CPACSOutPatientUseChemistry*)pWnd)->OnSavePACSOutPatientUseChemistry();
} 
static int _OnCancelPACSOutPatientUseChemistryFnc(CWnd *pWnd){
	 return ((CPACSOutPatientUseChemistry*)pWnd)->OnCancelPACSOutPatientUseChemistry();
} 
CPACSOutPatientUseChemistry::CPACSOutPatientUseChemistry(CWnd *pParent){

	m_nDlgWidth = 440;
	m_nDlgHeight = 160;
	SetDefaultValues();
}
CPACSOutPatientUseChemistry::~CPACSOutPatientUseChemistry(){
}
void CPACSOutPatientUseChemistry::OnCreateComponents(){
	m_wndReportParameters.Create(this, _T("Report Parameters"), 5, 5, 430, 120);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 30, 300, 55);
	m_wndToDate.Create(this,305, 30, 425, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 90, 85);
	m_wndStock.Create(this,95, 60, 425, 85); 
	m_wndExport.Create(this, _T("Export"), 350, 125, 430, 150);
	m_wndDrug.Create(this, _T("Drug"), 280, 90, 340, 115);
	m_wndMaterial.Create(this, _T("Material"), 345, 90, 425, 115);

}
void CPACSOutPatientUseChemistry::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);


	m_wndStock.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 50);
	m_wndStock.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CPACSOutPatientUseChemistry::OnSetWindowEvents(){
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
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndDrug.SetEvent(WE_CLICK, _OnDrugSelectFnc);
	m_wndMaterial.SetEvent(WE_CLICK, _OnMaterialSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	m_wndStock.SetCheckBox(true);
	SetMode(VM_EDIT);

}
void CPACSOutPatientUseChemistry::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_Radio(pDX, m_wndDrug.GetDlgCtrlID(), m_nDrug);
	DDX_Radio(pDX, m_wndMaterial.GetDlgCtrlID(), m_nMaterial);

}
void CPACSOutPatientUseChemistry::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPACSOutPatientUseChemistry::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CPACSOutPatientUseChemistry::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_nDrug=1;
	m_nMaterial=0;

}
int CPACSOutPatientUseChemistry::SetMode(int nMode){
 		int nOldMode = GetMode();
 		CGuiView::SetMode(nMode);
 		CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 		CString szSQL;
 		CRecord rs(&pMF->m_db);
  		switch(nMode){
 		case VM_ADD: 
 			EnableControls(TRUE);
 			EnableButtons(TRUE, 0, -1);
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
/*void CPACSOutPatientUseChemistry::OnFromDateChange(){
	
} */
/*void CPACSOutPatientUseChemistry::OnFromDateSetfocus(){
	
} */
/*void CPACSOutPatientUseChemistry::OnFromDateKillfocus(){
	
} */
int CPACSOutPatientUseChemistry::OnFromDateCheckValue(){
	return 0;
} 
/*void CPACSOutPatientUseChemistry::OnToDateChange(){
	
} */
/*void CPACSOutPatientUseChemistry::OnToDateSetfocus(){
	
} */
/*void CPACSOutPatientUseChemistry::OnToDateKillfocus(){
	
} */
int CPACSOutPatientUseChemistry::OnToDateCheckValue(){
	return 0;
} 
void CPACSOutPatientUseChemistry::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSOutPatientUseChemistry::OnStockSelendok(){
	 
}
/*void CPACSOutPatientUseChemistry::OnStockSetfocus(){
	
}*/
/*void CPACSOutPatientUseChemistry::OnStockKillfocus(){
	
}*/
long CPACSOutPatientUseChemistry::OnStockLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStock.IsSearchKey() && !m_szStockKey.IsEmpty()){
		szWhere.Format(_T(" and msl_storage_id ='%s' "), m_szStockKey);
	}
	m_wndStock.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT msl_storage_id as id, msl_name as name FROM m_storagelist WHERE msl_type = 'C' AND msl_dept_id = 'C8' %s ORDER BY id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStock.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPACSOutPatientUseChemistry::OnStockAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CPACSOutPatientUseChemistry::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	{
		if (m_wndStock.GetCheck(i))
		{
			if (!szStock.IsEmpty())
				szStock += _T(", ");
			szStock.AppendFormat(_T("'%s'"), m_wndStock.GetCheck(i, 0));
		}
	}
	szWhere1.Format(_T(" AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szWhere.Format(_T(" AND hpo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!szStock.IsEmpty()){
		szWhere1.AppendFormat(_T(" AND mt_storage_to_id IN (%s)"), szStock);
		szWhere.AppendFormat(_T(" AND hpo_storage_id IN (%s)"), szStock);
					_T("               Sum(sol_qty)      sol_qty, ") \
					_T("               Sum(solins_qty)   solins_qty, ") \
					_T("               Sum(otherins_qty) otherins_qty, ") \
					_T("               Sum(pol_qty)      pol_qty, ") \
					_T("               Sum(service_qty)  service_qty, ") \
					_T("               Sum(other_qty)    other_qty, ") \
					_T("               Sum(total_qty)    total_qty ") \
					_T("		FROM") \
					_T("		 (SELECT hmt_suppproduct_id,") \
					_T("                       CASE ") \
					_T("                         WHEN hmt_objectid = 1 THEN hmt_qtyissue ") \
					_T("                         ELSE 0 ") \
					_T("                       END           sol_qty, ") \
					_T("                       CASE ") \
					_T("                         WHEN hmt_objectid = 2 THEN hmt_qtyissue ") \
					_T("                         ELSE 0 ") \
					_T("                       END           solins_qty, ") \
					_T("                       CASE ") \
					_T("                         WHEN hmt_objectid = 4 THEN hmt_qtyissue ") \
					_T("                         ELSE 0 ") \
					_T("                       END           otherins_qty, ") \
					_T("                       CASE ") \
					_T("                         WHEN hmt_objectid = 3 THEN hmt_qtyissue ") \
					_T("                         ELSE 0 ") \
					_T("                       END           pol_qty, ") \
					_T("                       CASE ") \
					_T("                         WHEN hmt_objectid = 7 THEN hmt_qtyissue ") \
					_T("                         ELSE 0 ") \
					_T("                       END           service_qty, ") \
					_T("                       CASE ") \
					_T("                         WHEN hmt_objectid NOT IN ( 1, 2, 3, 4, 7 ) THEN hmt_qtyissue ") \
					_T("                         ELSE 0 ") \
					_T("                       END           other_qty, ") \
					_T("                       hmt_qtyissue total_qty ") \
					_T("         FROM   m_transaction ") \
					_T("		 LEFT JOIN hms_medical_transaction_view ON (mt_transaction_id = hmt_reftransaction_id)") \
					_T("		 LEFT JOIN hms_ipharmaorder ON (hpo_orderid = hmt_orderid)") \
					_T("         LEFT JOIN hms_pacsorder ON ( hpc_docno = hpo_docno AND hpc_orderid = hpo_reference_id ) ") \
					_T("         WHERE  mt_department_to_id = 'C8' AND mt_status = 'A' AND hpc_class = 'E'") \
					_T("                       %s ") \
					_T("         AND mt_doctype = 'CSO') ") \
					_T(" LEFT JOIN m_product_view ON (product_id = hmt_suppproduct_id)") \
					_T(" WHERE product_org_id = '%s'") \
					_T(" GROUP  BY product_id, ") \
					_T("           product_name, ") \
					_T("           product_uomname "), szWhere1, m_szType);
					_T("        product_name, ") \
					_T("        product_uomname, ") \
					_T("        Sum(sol_qty)      sol_qty, ") \
					_T("        Sum(solins_qty)   solins_qty, ") \
					_T("        Sum(otherins_qty) otherins_qty, ") \
					_T("        Sum(pol_qty)      pol_qty, ") \
					_T("        Sum(service_qty)  service_qty, ") \
					_T("        Sum(other_qty)    other_qty, ") \
					_T("        Sum(total_qty)    total_qty ") \
					_T(" FROM   (SELECT hpol_product_id, ") \
					_T("                CASE ") \
					_T("                  WHEN hpo_object_id = 1 THEN hpol_qtyissue ") \
					_T("                  ELSE 0 ") \
					_T("                END           sol_qty, ") \
					_T("                CASE ") \
					_T("                  WHEN hpo_object_id = 2 THEN hpol_qtyissue ") \
					_T("                  ELSE 0 ") \
					_T("                END           solins_qty, ") \
					_T("                CASE ") \
					_T("                  WHEN hpo_object_id = 4 THEN hpol_qtyissue ") \
					_T("                  ELSE 0 ") \
					_T("                END           otherins_qty, ") \
					_T("                CASE ") \
					_T("                  WHEN hpo_object_id = 3 THEN hpol_qtyissue ") \
					_T("                  ELSE 0 ") \
					_T("                END           pol_qty, ") \
					_T("                CASE ") \
					_T("                  WHEN hpo_object_id = 7 THEN hpol_qtyissue ") \
					_T("                  ELSE 0 ") \
					_T("                END           service_qty, ") \
					_T("                CASE ") \
					_T("                  WHEN hpo_object_id NOT IN ( 1, 2, 3, 4, 7 ) THEN hpol_qtyissue ") \
					_T("                  ELSE 0 ") \
					_T("                END           other_qty, ") \
					_T("                hpol_qtyissue total_qty ") \
					_T("         FROM   hms_ipharmaorder ") \
					_T("                LEFT JOIN hms_ipharmaorderline ") \
					_T("                       ON ( hpo_orderid = hpol_orderid ) ") \
					_T("                LEFT JOIN hms_pacsorder ") \
					_T("                       ON ( hpc_docno = hpo_docno ") \
					_T("                            AND hpc_orderid = hpo_reference_id ) ") \
					_T("         WHERE  hpo_deptid = 'C8' %s ") \
					_T("         AND hpc_class = 'E'  ") \
					_T("         AND hpo_orderstatus = 'A' ") \
					_T("         AND hpo_ordertype = 'C') tbl0 ") \
					_T(" LEFT JOIN m_product_view ON (product_id = hpol_product_id)") \
					_T(" WHERE product_org_id = '%s'") \
					_T(" GROUP  BY product_id, ") \
					_T("           product_name, ") \
					_T("           product_uomname "), szWhere, m_szType);

		if (nCount < 0)
			_msg(_T("%s"), szSQL);
		else
			_fmsg(_T("%s"), szSQL);
		return;
	
} 
void CPACSOutPatientUseChemistry::OnDrugSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	m_szType.Format(_T("PM"));
}
void CPACSOutPatientUseChemistry::OnMaterialSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	m_szType.Format(_T("MA"));
}
int CPACSOutPatientUseChemistry::OnAddPACSOutPatientUseChemistry(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CPACSOutPatientUseChemistry::OnEditPACSOutPatientUseChemistry(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CPACSOutPatientUseChemistry::OnDeletePACSOutPatientUseChemistry(){
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
 		OnCancelPACSOutPatientUseChemistry();
 	}
	return 0;
}
int CPACSOutPatientUseChemistry::OnSavePACSOutPatientUseChemistry(){
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
 		//OnPACSOutPatientUseChemistryListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CPACSOutPatientUseChemistry::OnCancelPACSOutPatientUseChemistry(){
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
int CPACSOutPatientUseChemistry::OnPACSOutPatientUseChemistryListLoadData(){
	return 0;
}