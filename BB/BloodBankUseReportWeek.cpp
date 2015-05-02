#include "stdafx.h"
#include "BloodBankUseReportWeek.h"
#include "HMSMainFrame.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CBloodBankUseReportWeek *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CBloodBankUseReportWeek *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CBloodBankUseReportWeek *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CBloodBankUseReportWeek *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CBloodBankUseReportWeek *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CBloodBankUseReportWeek *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CBloodBankUseReportWeek *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CBloodBankUseReportWeek *)pWnd)->OnToDateCheckValue();
} 
static void _OnDepartmentSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CBloodBankUseReportWeek* )pWnd)->OnDepartmentSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDepartmentSelendokFnc(CWnd *pWnd){
	((CBloodBankUseReportWeek *)pWnd)->OnDepartmentSelendok();
}
/*static void _OnDepartmentSetfocusFnc(CWnd *pWnd){
	((CBloodBankUseReportWeek *)pWnd)->OnDepartmentKillfocus();
}*/
/*static void _OnDepartmentKillfocusFnc(CWnd *pWnd){
	((CBloodBankUseReportWeek *)pWnd)->OnDepartmentKillfocus();
}*/
static long _OnDepartmentLoadDataFnc(CWnd *pWnd){
	return ((CBloodBankUseReportWeek *)pWnd)->OnDepartmentLoadData();
}
/*static void _OnDepartmentAddNewFnc(CWnd *pWnd){
	((CBloodBankUseReportWeek *)pWnd)->OnDepartmentAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CBloodBankUseReportWeek *pVw = (CBloodBankUseReportWeek *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CBloodBankUseReportWeek *pVw = (CBloodBankUseReportWeek *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddBloodBankUseReportWeekFnc(CWnd *pWnd){
	 return ((CBloodBankUseReportWeek*)pWnd)->OnAddBloodBankUseReportWeek();
} 
static int _OnEditBloodBankUseReportWeekFnc(CWnd *pWnd){
	 return ((CBloodBankUseReportWeek*)pWnd)->OnEditBloodBankUseReportWeek();
} 
static int _OnDeleteBloodBankUseReportWeekFnc(CWnd *pWnd){
	 return ((CBloodBankUseReportWeek*)pWnd)->OnDeleteBloodBankUseReportWeek();
} 
static int _OnSaveBloodBankUseReportWeekFnc(CWnd *pWnd){
	 return ((CBloodBankUseReportWeek*)pWnd)->OnSaveBloodBankUseReportWeek();
} 
static int _OnCancelBloodBankUseReportWeekFnc(CWnd *pWnd){
	 return ((CBloodBankUseReportWeek*)pWnd)->OnCancelBloodBankUseReportWeek();
} 
CBloodBankUseReportWeek::CBloodBankUseReportWeek(CWnd *pWnd){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CBloodBankUseReportWeek::~CBloodBankUseReportWeek(){
}
void CBloodBankUseReportWeek::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("ReportCondition"), 5, 5, 470, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 235, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 240, 30, 320, 55);
	m_wndToDate.Create(this,325, 30, 465, 55); 
	m_wndDepartmentLabel.Create(this, _T("Department"), 10, 60, 90, 85);
	m_wndDepartment.Create(this,95, 60, 465, 85); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 265, 95, 365, 120);
	m_wndPrintPreview.EnableWindow(FALSE);
	m_wndExport.Create(this, _T("&Export"), 370, 95, 470, 120);

}
void CBloodBankUseReportWeek::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndDepartment.SetCheckValue(true);
	m_wndDepartment.LimitText(1024);

	m_wndDepartment.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndDepartment.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);



}
void CBloodBankUseReportWeek::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndDepartment.SetEvent(WE_SELENDOK, _OnDepartmentSelendokFnc);
	//m_wndDepartment.SetEvent(WE_SETFOCUS, _OnDepartmentSetfocusFnc);
	//m_wndDepartment.SetEvent(WE_KILLFOCUS, _OnDepartmentKillfocusFnc);
	m_wndDepartment.SetEvent(WE_SELCHANGE, _OnDepartmentSelectChangeFnc);
	m_wndDepartment.SetEvent(WE_LOADDATA, _OnDepartmentLoadDataFnc);
	//m_wndDepartment.SetEvent(WE_ADDNEW, _OnDepartmentAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CBloodBankUseReportWeek::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDepartment.GetDlgCtrlID(), m_szDepartmentKey);

}
void CBloodBankUseReportWeek::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CBloodBankUseReportWeek::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CBloodBankUseReportWeek::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDepartmentKey.Empty();

}
int CBloodBankUseReportWeek::SetMode(int nMode){
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
/*void CBloodBankUseReportWeek::OnFromDateChange(){
	
} */
/*void CBloodBankUseReportWeek::OnFromDateSetfocus(){
	
} */
/*void CBloodBankUseReportWeek::OnFromDateKillfocus(){
	
} */
int CBloodBankUseReportWeek::OnFromDateCheckValue(){
	return 0;
} 
/*void CBloodBankUseReportWeek::OnToDateChange(){
	
} */
/*void CBloodBankUseReportWeek::OnToDateSetfocus(){
	
} */
/*void CBloodBankUseReportWeek::OnToDateKillfocus(){
	
} */
int CBloodBankUseReportWeek::OnToDateCheckValue(){
	return 0;
} 
void CBloodBankUseReportWeek::OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CBloodBankUseReportWeek::OnDepartmentSelendok(){
	 
}
/*void CBloodBankUseReportWeek::OnDepartmentSetfocus(){
	
}*/
/*void CBloodBankUseReportWeek::OnDepartmentKillfocus(){
	
}*/
long CBloodBankUseReportWeek::OnDepartmentLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDepartment.IsSearchKey() && !m_szDepartmentKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szDepartmentKey);
	}
	m_wndDepartment.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("select sd_id as id, sd_name as name from sys_dept where sd_type in('KB','DT') order by id"));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDepartment.AddItems(
			rs.GetValue(_T("id")),
			rs.GetValue(_T("name")),NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CBloodBankUseReportWeek::OnDepartmentAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CBloodBankUseReportWeek::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CBloodBankUseReportWeek::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	for (int i = 1; i<= 15; i++)
	{
		nTotal[i] = 0;
	}
		if (szOldGroup != szNewGroup)
		{
			xls.SetCellText(0, nRow, rs.GetValue(_T("type")), FMT_TEXT, true);
			/*xls.SetCellMergedColumns(nCol+1, nRow+1 , 1);
			xls.SetCellText(nCol, nRow, _T("\x43\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 11);

			for( int i = 1; i<= 15; i++)
			{
				tmpStr.Format(_T("%d"), nTotal[i]);
				xls.SetCellText(i , nRow , tmpStr, FMT_NUMBER1 , true, 11);
			}*/
			szOldGroup = szNewGroup;
			nRow++;
		}
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng"), FMT_TEXT | FMT_CENTER, true, 11);

	for( int i = 1; i<= 15; i++)
	{
		tmpStr.Format(_T("%d"), nTotal[i]);
		xls.SetCellText(i , nRow , tmpStr, FMT_NUMBER1 , true, 11);
	}
} 
int CBloodBankUseReportWeek::OnAddBloodBankUseReportWeek(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CBloodBankUseReportWeek::OnEditBloodBankUseReportWeek(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CBloodBankUseReportWeek::OnDeleteBloodBankUseReportWeek(){
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
 		OnCancelBloodBankUseReportWeek();
 	}
	return 0;
}
int CBloodBankUseReportWeek::OnSaveBloodBankUseReportWeek(){
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
 		//OnBloodBankUseReportWeekListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CBloodBankUseReportWeek::OnCancelBloodBankUseReportWeek(){
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
int CBloodBankUseReportWeek::OnBloodBankUseReportWeekListLoadData(){
	return 0;
}
CString CBloodBankUseReportWeek::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere, szDept;
	if (!m_szDepartmentKey.IsEmpty())
		szDept.AppendFormat(_T(" AND hpo_deptid ='%s'"), m_szDepartmentKey);
	szSQL.Format(_T(" SELECT 1 id,'Kh\x1ED1i n\x1ED9i' type,doituong,sum(songuoik) songuoik,sum(m250) m250,sum(m350) m350,sum(slm) slm,") \
		_T("        sum(songuoih) songuoih,sum(h150) h150,sum(h250) h250,sum(slh) slh,") \
		_T("        sum(songuoit) songuoit,sum(tpool) tpool,sum(tmay) tmay,sum(slt) slt,sum(songuoik+songuoih+songuoit) tongsn,") \
		_T("        sum(m250+m350+h150+h250+tpool+tmay) tongsl,") \
		_T("        sum(slm+slh+slt) tongslml") \
		_T(" FROM(SELECT get_objectname(doituong) doituong,") \
		_T("        0 songuoik,") \
		_T("        sum(m250+k250) m250,sum(m350+k350) m350,sum(m250+k250)*250+sum(m350+k350)*350 slm,") \
		_T("        0 songuoih,") \
		_T("        sum(h150) h150,sum(h250) h250,sum(h150)*150+sum(h250)*250 slh,") \
		_T("        0 songuoit,") \
		_T("        sum(tpool) tpool,sum(tmay) tmay,sum(tpool)*150+sum(tmay)*250 slt") \
		_T(" FROM(SELECT doituong,") \
		_T("        case when product_id=6966 then 1 else 0 end m250,") \
		_T("        case when product_id=6970 then 1 else 0 end m350,") \
		_T("        case when product_id=6972 then 1 else 0 end k250,") \
		_T("        case when product_id=6973 then 1 else 0 end k350,") \
		_T("        case when product_id=6976 then 1 else 0 end h150,") \
		_T("        case when product_id=6975 then 1 else 0 end h250,") \
		_T("        case when product_id=6978 then 1 else 0 end tpool,") \
		_T("        case when product_id=6977 then 1 else 0 end tmay") \
		_T(" FROM(SELECT hpol_docno,") \
		_T("        product_id,") \
		_T("        hpo_object_id doituong") \
		_T(" FROM hms_ipharmaorder") \
		_T(" LEFT JOIN hms_ipharmaorderline ON(hpol_orderid=hpo_orderid)") \
		_T(" LEFT JOIN m_productitem_view ON(product_item_id=hpol_product_item_id)") \
		_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = hpol_product_item_id)") \
		_T(" WHERE hpo_org_id='BB' and substr(hpo_deptid,1,1)='A' and hpo_orderstatus='A' and hpo_orderdate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s))") \
		_T(" GROUP BY doituong") \
		_T(" UNION ALL") \
		_T(" SELECT get_objectname(doituong) doituong,sum(songuoik) songuoik,0 m250,0 m350,0 slm,sum(songuoih) songuoih,0 h150,0 h250,0 slh,sum(songuoit) songuoit,0 tpool,0 tmay,0 slt") \
		_T(" FROM(SELECT distinct hpol_docno,1 songuoik,0 songuoih,0 songuoit,hpo_object_id doituong") \
		_T(" FROM hms_ipharmaorder") \
		_T(" LEFT JOIN hms_ipharmaorderline ON(hpol_orderid=hpo_orderid)") \
		_T(" LEFT JOIN m_productitem_view ON(product_item_id=hpol_product_item_id)") \
		_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = hpol_product_item_id)") \
		_T(" WHERE hpo_org_id='BB' and product_id in(6966,6970,6972,6973) and substr(hpo_deptid,1,1)='A' and hpo_orderstatus='A' and hpo_orderdate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
		_T(" UNION ALL") \
		_T(" SELECT distinct hpol_docno,0 songuoik,1 songuoih,0 songuoit,hpo_object_id doituong") \
		_T(" FROM hms_ipharmaorder") \
		_T(" LEFT JOIN hms_ipharmaorderline ON(hpol_orderid=hpo_orderid)") \
		_T(" LEFT JOIN m_productitem_view ON(product_item_id=hpol_product_item_id)") \
		_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = hpol_product_item_id)") \
		_T(" WHERE hpo_org_id='BB' and product_id in(6976,6975) and substr(hpo_deptid,1,1)='A' and hpo_orderstatus='A' and hpo_orderdate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
		_T(" UNION ALL") \
		_T(" SELECT distinct hpol_docno,0 songuoik,0 songuoih,1 songuoit,hpo_object_id doituong") \
		_T(" FROM hms_ipharmaorder") \
		_T(" LEFT JOIN hms_ipharmaorderline ON(hpol_orderid=hpo_orderid)") \
		_T(" LEFT JOIN m_productitem_view ON(product_item_id=hpol_product_item_id)") \
		_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = hpol_product_item_id)") \
		_T(" WHERE hpo_org_id='BB' and product_id in(6978,6977) and substr(hpo_deptid,1,1)='A' and hpo_orderstatus='A' and hpo_orderdate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s)") \
		_T(" GROUP BY doituong)") \
		_T(" GROUP BY doituong") \
		_T(" UNION ALL") \
		_T(" SELECT 2 id,'Kh\x1ED1i ngo\x1EA1i' type,doituong,sum(songuoik) songuoik,sum(m250) m250,sum(m350) m350,sum(slm) slm,") \
		_T("        sum(songuoih) songuoih,sum(h150) h150,sum(h250) h250,sum(slh) slh,") \
		_T("        sum(songuoit) songuoit,sum(tpool) tpool,sum(tmay) tmay,sum(slt) slt,sum(songuoik+songuoih+songuoit) tongsn,") \
		_T("        sum(m250+m350+h150+h250+tpool+tmay) tongsl,") \
		_T("        sum(slm+slh+slt) tongslml") \
		_T(" FROM(SELECT get_objectname(doituong) doituong,") \
		_T("        0 songuoik,") \
		_T("        sum(m250+k250) m250,sum(m350+k350) m350,sum(m250+k250)*250+sum(m350+k350)*350 slm,") \
		_T("        0 songuoih,") \
		_T("        sum(h150) h150,sum(h250) h250,sum(h150)*150+sum(h250)*250 slh,") \
		_T("        0 songuoit,") \
		_T("        sum(tpool) tpool,sum(tmay) tmay,sum(tpool)*150+sum(tmay)*250 slt") \
		_T(" FROM(SELECT doituong,") \
		_T("        case when product_id=6966 then 1 else 0 end m250,") \
		_T("        case when product_id=6970 then 1 else 0 end m350,") \
		_T("        case when product_id=6972 then 1 else 0 end k250,") \
		_T("        case when product_id=6973 then 1 else 0 end k350,") \
		_T("        case when product_id=6976 then 1 else 0 end h150,") \
		_T("        case when product_id=6975 then 1 else 0 end h250,") \
		_T("        case when product_id=6978 then 1 else 0 end tpool,") \
		_T("        case when product_id=6977 then 1 else 0 end tmay") \
		_T(" FROM(SELECT hpol_docno,") \
		_T("        product_id,") \
		_T("        hpo_object_id doituong") \
		_T(" FROM hms_ipharmaorder") \
		_T(" LEFT JOIN hms_ipharmaorderline ON(hpol_orderid=hpo_orderid)") \
		_T(" LEFT JOIN m_productitem_view ON(product_item_id=hpol_product_item_id)") \
		_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = hpol_product_item_id)") \
		_T(" WHERE hpo_org_id='BB' and substr(hpo_deptid,1,1) IN('B','C') and hpo_orderstatus='A' and hpo_orderdate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s))") \
		_T(" GROUP BY doituong") \
		_T(" UNION ALL") \
		_T(" SELECT get_objectname(doituong) doituong,sum(songuoik) songuoik,0 m250,0 m350,0 slm,sum(songuoih) songuoih,0 h150,0 h250,0 slh,sum(songuoit) songuoit,0 tpool,0 tmay,0 slt") \
		_T(" FROM(SELECT distinct hpol_docno,1 songuoik,0 songuoih,0 songuoit,hpo_object_id doituong") \
		_T(" FROM hms_ipharmaorder") \
		_T(" LEFT JOIN hms_ipharmaorderline ON(hpol_orderid=hpo_orderid)") \
		_T(" LEFT JOIN m_productitem_view ON(product_item_id=hpol_product_item_id)") \
		_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = hpol_product_item_id)") \
		_T(" WHERE hpo_org_id='BB' and product_id in(6966,6970,6972,6973) and substr(hpo_deptid,1,1) IN('B','C') and hpo_orderstatus='A' and hpo_orderdate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
		_T(" UNION ALL") \
		_T(" SELECT distinct hpol_docno,0 songuoik,1 songuoih,0 songuoit,hpo_object_id doituong") \
		_T(" FROM hms_ipharmaorder") \
		_T(" LEFT JOIN hms_ipharmaorderline ON(hpol_orderid=hpo_orderid)") \
		_T(" LEFT JOIN m_productitem_view ON(product_item_id=hpol_product_item_id)") \
		_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = hpol_product_item_id)") \
		_T(" WHERE hpo_org_id='BB' and product_id in(6976,6975) and substr(hpo_deptid,1,1) IN('B','C') and hpo_orderstatus='A' and hpo_orderdate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
		_T(" UNION ALL") \
		_T(" SELECT distinct hpol_docno,0 songuoik,0 songuoih,1 songuoit,hpo_object_id doituong") \
		_T(" FROM hms_ipharmaorder") \
		_T(" LEFT JOIN hms_ipharmaorderline ON(hpol_orderid=hpo_orderid)") \
		_T(" LEFT JOIN m_productitem_view ON(product_item_id=hpol_product_item_id)") \
		_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = hpol_product_item_id)") \
		_T(" WHERE hpo_org_id='BB' and product_id in(6978,6977) and substr(hpo_deptid,1,1) IN('B','C') and hpo_orderstatus='A' and hpo_orderdate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s)") \
		_T(" GROUP BY doituong)") \
		_T(" GROUP BY doituong"),m_szFromDate,m_szToDate,szDept,
								m_szFromDate,m_szToDate,szDept,
								m_szFromDate,m_szToDate,szDept,
								m_szFromDate,m_szToDate,szDept,
								m_szFromDate,m_szToDate,szDept,
								m_szFromDate,m_szToDate,szDept,
								m_szFromDate,m_szToDate,szDept,
								m_szFromDate,m_szToDate,szDept);
	//QueryOpener(szSQL);
	return szSQL;
}