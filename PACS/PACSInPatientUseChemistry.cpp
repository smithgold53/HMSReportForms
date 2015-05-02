#include "stdafx.h"
#include "PACSInPatientUseChemistry.h"
#include "HMSMainFrame.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPACSInPatientUseChemistry *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPACSInPatientUseChemistry *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPACSInPatientUseChemistry *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSInPatientUseChemistry *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPACSInPatientUseChemistry *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPACSInPatientUseChemistry *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPACSInPatientUseChemistry *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSInPatientUseChemistry *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSInPatientUseChemistry* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CPACSInPatientUseChemistry *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CPACSInPatientUseChemistry *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CPACSInPatientUseChemistry *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CPACSInPatientUseChemistry *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CPACSInPatientUseChemistry *)pWnd)->OnStockAddNew();
}*/
static void _OnExportSelectFnc(CWnd *pWnd){
	CPACSInPatientUseChemistry *pVw = (CPACSInPatientUseChemistry *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnDrugSelectFnc(CWnd *pWnd){
	  ((CPACSInPatientUseChemistry*)pWnd)->OnDrugSelect();
}
static void _OnMaterialSelectFnc(CWnd *pWnd){
	  ((CPACSInPatientUseChemistry*)pWnd)->OnMaterialSelect();
}
static int _OnAddPACSInPatientUseChemistryFnc(CWnd *pWnd){
	 return ((CPACSInPatientUseChemistry*)pWnd)->OnAddPACSInPatientUseChemistry();
} 
static int _OnEditPACSInPatientUseChemistryFnc(CWnd *pWnd){
	 return ((CPACSInPatientUseChemistry*)pWnd)->OnEditPACSInPatientUseChemistry();
} 
static int _OnDeletePACSInPatientUseChemistryFnc(CWnd *pWnd){
	 return ((CPACSInPatientUseChemistry*)pWnd)->OnDeletePACSInPatientUseChemistry();
} 
static int _OnSavePACSInPatientUseChemistryFnc(CWnd *pWnd){
	 return ((CPACSInPatientUseChemistry*)pWnd)->OnSavePACSInPatientUseChemistry();
} 
static int _OnCancelPACSInPatientUseChemistryFnc(CWnd *pWnd){
	 return ((CPACSInPatientUseChemistry*)pWnd)->OnCancelPACSInPatientUseChemistry();
} 
CPACSInPatientUseChemistry::CPACSInPatientUseChemistry(CWnd *pParent){
	m_szType = _T("MA");
	m_nDlgWidth = 440;
	m_nDlgHeight = 160;
	SetDefaultValues();
}
CPACSInPatientUseChemistry::~CPACSInPatientUseChemistry(){
}
void CPACSInPatientUseChemistry::OnCreateComponents(){
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
void CPACSInPatientUseChemistry::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);

	m_wndStock.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 50);
	m_wndStock.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CPACSInPatientUseChemistry::OnSetWindowEvents(){
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
void CPACSInPatientUseChemistry::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_Radio(pDX, m_wndDrug.GetDlgCtrlID(), m_nDrug);
	DDX_Radio(pDX, m_wndMaterial.GetDlgCtrlID(), m_nMaterial);

}
void CPACSInPatientUseChemistry::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPACSInPatientUseChemistry::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CPACSInPatientUseChemistry::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_nDrug=1;
	m_nMaterial=0;

}
int CPACSInPatientUseChemistry::SetMode(int nMode){
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
/*void CPACSInPatientUseChemistry::OnFromDateChange(){
	
} */
/*void CPACSInPatientUseChemistry::OnFromDateSetfocus(){
	
} */
/*void CPACSInPatientUseChemistry::OnFromDateKillfocus(){
	
} */
int CPACSInPatientUseChemistry::OnFromDateCheckValue(){
	return 0;
} 
/*void CPACSInPatientUseChemistry::OnToDateChange(){
	
} */
/*void CPACSInPatientUseChemistry::OnToDateSetfocus(){
	
} */
/*void CPACSInPatientUseChemistry::OnToDateKillfocus(){
	
} */
int CPACSInPatientUseChemistry::OnToDateCheckValue(){
	return 0;
} 
void CPACSInPatientUseChemistry::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSInPatientUseChemistry::OnStockSelendok(){
	 
}
/*void CPACSInPatientUseChemistry::OnStockSetfocus(){
	
}*/
/*void CPACSInPatientUseChemistry::OnStockKillfocus(){
	
}*/
long CPACSInPatientUseChemistry::OnStockLoadData(){
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
/*void CPACSInPatientUseChemistry::OnStockAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CPACSInPatientUseChemistry::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	{
		if (m_wndStock.GetCheck(i))
		{
			if (!szStock.IsEmpty())
				szStock += _T(", ");
			szStock.AppendFormat(_T("'%s'"), m_wndStock.GetCheck(i, 0));
		}
	}
	if (!szStock.IsEmpty()){
		szWhere.AppendFormat(_T(" AND hpo_storage_id IN (%s)"), szStock);
					_T("        product_name, ") \
					_T("        product_uomname, ") \
					_T("        Sum(sol_qty)      sol_qty, ") \
					_T("        Sum(solins_qty)   solins_qty, ") \
					_T("        Sum(otherins_qty) otherins_qty, ") \
					_T("        Sum(pol_qty)      pol_qty, ") \
					_T("        Sum(service_qty)  service_qty, ") \
					_T("        Sum(other_qty)    other_qty, ") \
					_T("        Sum(total_qty)    total_qty ") \
					_T(" FROM  (SELECT hmt_suppproduct_id, ") \
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
					_T("         WHERE  mt_department_to_id = 'C8' AND mt_status = 'A' AND hpc_class = 'I'") \
					_T("                       %s ") \
					_T("					   AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
					_T("                       AND mt_doctype = 'CSO') ") \
					_T("        GROUP  BY hmt_suppproduct_id ") \
					_T("        UNION ALL ") \
					_T("        SELECT mtl_product_id, ") \
					_T("               Sum(mtl_qtysold), ") \
					_T("               Sum(mtl_qtysoldins), ") \
					_T("               Sum(mtl_qtyotherins), ") \
					_T("               Sum(mtl_qtypolicy), ") \
					_T("               Sum(mtl_qtyservice), ") \
					_T("               Sum(mtl_qtyother), ") \
					_T("               Sum(mtl_qtydelivery) ") \
					_T("        FROM   m_transaction ") \
					_T("               LEFT JOIN m_transactionline ") \
					_T("                      ON ( mt_transaction_id = mtl_transaction_id ) ") \
					_T("        WHERE  mt_doctype IN ('DDO', 'CSO') AND mtl_qtydelivery > 0") \
					_T("               AND mt_department_to_id = 'C8' AND mtl_org_id = 'PACS'") \
					_T("               %s ") \
					_T("               AND mt_status = 'A' AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
					_T("        GROUP  BY mtl_product_id) tbl ") \
					_T("       LEFT JOIN m_product_view ") \
					_T("              ON ( product_id = hmt_suppproduct_id ) ") \
					_T(" WHERE  product_org_id = '%s' ") \
					_T(" GROUP  BY product_id, ") \
					_T("           product_name, ") \
					_T("           product_uomname "), szWhere1, m_szFromDate, m_szToDate, szWhere1, m_szFromDate, m_szToDate, m_szType);
					_T("        product_name, ") \
					_T("        product_uomname, ") \
					_T("        Sum(sol_qty)      sol_qty, ") \
					_T("        Sum(solins_qty)   solins_qty, ") \
					_T("        Sum(otherins_qty) otherins_qty, ") \
					_T("        Sum(pol_qty)      pol_qty, ") \
					_T("        Sum(service_qty)  service_qty, ") \
					_T("        Sum(other_qty)    other_qty, ") \
					_T("        Sum(total_qty)    total_qty ") \
					_T(" FROM  (SELECT hpol_product_id, ") \
					_T("               Sum(sol_qty)      sol_qty, ") \
					_T("               Sum(solins_qty)   solins_qty, ") \
					_T("               Sum(otherins_qty) otherins_qty, ") \
					_T("               Sum(pol_qty)      pol_qty, ") \
					_T("               Sum(service_qty)  service_qty, ") \
					_T("               Sum(other_qty)    other_qty, ") \
					_T("               Sum(total_qty)    total_qty ") \
					_T("        FROM   (SELECT hpol_product_id, ") \
					_T("                       CASE ") \
					_T("                         WHEN hpo_object_id = 1 THEN hpol_qtyissue ") \
					_T("                         ELSE 0 ") \
					_T("                       END           sol_qty, ") \
					_T("                       CASE ") \
					_T("                         WHEN hpo_object_id = 2 THEN hpol_qtyissue ") \
					_T("                         ELSE 0 ") \
					_T("                       END           solins_qty, ") \
					_T("                       CASE ") \
					_T("                         WHEN hpo_object_id = 4 THEN hpol_qtyissue ") \
					_T("                         ELSE 0 ") \
					_T("                       END           otherins_qty, ") \
					_T("                       CASE ") \
					_T("                         WHEN hpo_object_id = 3 THEN hpol_qtyissue ") \
					_T("                         ELSE 0 ") \
					_T("                       END           pol_qty, ") \
					_T("                       CASE ") \
					_T("                         WHEN hpo_object_id = 7 THEN hpol_qtyissue ") \
					_T("                         ELSE 0 ") \
					_T("                       END           service_qty, ") \
					_T("                       CASE ") \
					_T("                         WHEN hpo_object_id NOT IN ( 1, 2, 3, 4, 7 ) THEN hpol_qtyissue ") \
					_T("                         ELSE 0 ") \
					_T("                       END           other_qty, ") \
					_T("                       hpol_qtyissue total_qty ") \
					_T("                FROM   hms_ipharmaorder ") \
					_T("                LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
					_T("                LEFT JOIN hms_pacsorder ON ( hpc_docno = hpo_docno AND hpc_orderid = hpo_reference_id ) ") \
					_T("                WHERE  hpo_deptid = 'C8' AND hpo_orderstatus = 'A' AND hpc_class = 'I'") \
					_T("                       %s ") \
					_T("					   AND hpo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
					_T("                       AND hpo_ordertype = 'C') tbl0 ") \
					_T("        GROUP  BY hpol_product_id ") \
					_T("        UNION ALL ") \
					_T("        SELECT mtl_product_id, ") \
					_T("               Sum(mtl_qtysold), ") \
					_T("               Sum(mtl_qtysoldins), ") \
					_T("               Sum(mtl_qtyotherins), ") \
					_T("               Sum(mtl_qtypolicy), ") \
					_T("               Sum(mtl_qtyservice), ") \
					_T("               Sum(mtl_qtyother), ") \
					_T("               Sum(mtl_qtydelivery) ") \
					_T("        FROM   m_transaction ") \
					_T("               LEFT JOIN m_transactionline ") \
					_T("                      ON ( mt_transaction_id = mtl_transaction_id ) ") \
					_T("        WHERE  mt_doctype IN ('DDO', 'CSO') ") \
					_T("               AND mt_department_to_id = 'C8' AND mtl_org_id = 'PACS'") \
					_T("               %s ") \
					_T("               AND mt_status = 'A' AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
					_T("        GROUP  BY mtl_product_id) tbl ") \
					_T("       LEFT JOIN m_product_view ") \
					_T("              ON ( product_id = hpol_product_id ) ") \
					_T(" WHERE  product_org_id = '%s' ") \
					_T(" GROUP  BY product_id, ") \
					_T("           product_name, ") \
					_T("           product_uomname "), szWhere, m_szFromDate, m_szToDate, szWhere1, m_szFromDate, m_szToDate, m_szType);
	
		if (nCount < 0)
			_msg(_T("%s"), szSQL);
		else
			_fmsg(_T("%s"), szSQL);
		return;
} 
void CPACSInPatientUseChemistry::OnDrugSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	m_szType.Format(_T("PM"));
}
void CPACSInPatientUseChemistry::OnMaterialSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	m_szType.Format(_T("MA"));
}
int CPACSInPatientUseChemistry::OnAddPACSInPatientUseChemistry(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CPACSInPatientUseChemistry::OnEditPACSInPatientUseChemistry(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CPACSInPatientUseChemistry::OnDeletePACSInPatientUseChemistry(){
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
 		OnCancelPACSInPatientUseChemistry();
 	}
	return 0;
}
int CPACSInPatientUseChemistry::OnSavePACSInPatientUseChemistry(){
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
 		//OnPACSInPatientUseChemistryListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CPACSInPatientUseChemistry::OnCancelPACSInPatientUseChemistry(){
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
int CPACSInPatientUseChemistry::OnPACSInPatientUseChemistryListLoadData(){
	return 0;
}