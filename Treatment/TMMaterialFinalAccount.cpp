#include "stdafx.h"
#include "TMMaterialFinalAccount.h"
#include "HMSMainFrame.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CTMMaterialFinalAccount *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CTMMaterialFinalAccount *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CTMMaterialFinalAccount *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CTMMaterialFinalAccount *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CTMMaterialFinalAccount *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CTMMaterialFinalAccount *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CTMMaterialFinalAccount *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CTMMaterialFinalAccount *)pWnd)->OnToDateCheckValue();
} 
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMMaterialFinalAccount* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CTMMaterialFinalAccount *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CTMMaterialFinalAccount *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CTMMaterialFinalAccount *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CTMMaterialFinalAccount *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CTMMaterialFinalAccount *)pWnd)->OnStorageAddNew();
}*/
static void _OnExportSelectFnc(CWnd *pWnd){
	CTMMaterialFinalAccount *pVw = (CTMMaterialFinalAccount *)pWnd;
	pVw->OnExportSelect();
} 
CTMMaterialFinalAccount::CTMMaterialFinalAccount(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CTMMaterialFinalAccount::~CTMMaterialFinalAccount(){
}
void CTMMaterialFinalAccount::OnCreateComponents(){
	m_wndMaterialFinalAccount.Create(this, _T("Material Final Account"), 5, 5, 445, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 225, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 230, 30, 310, 55);
	m_wndToDate.Create(this,315, 30, 440, 55);
	m_wndStorage.SetCheckBox(true);
	m_wndStorageLabel.Create(this, _T("Storage"), 10, 60, 90, 85);
	m_wndStorage.Create(this,95, 60, 440, 85);
	m_wndDrug.Create(this, _T("Drug"), 5, 95, 85, 120);
	m_wndMaterial.Create(this, _T("Material"), 90, 95, 170, 120);
	m_wndExport.Create(this, _T("&Export"), 365, 95, 445, 120);

}
void CTMMaterialFinalAccount::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndStorage.SetCheckValue(true);
	m_wndStorage.LimitText(1024);

	m_wndStorage.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 50);
	m_wndStorage.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);

}
void CTMMaterialFinalAccount::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndStorage.SetEvent(WE_SELENDOK, _OnStorageSelendokFnc);
	//m_wndStorage.SetEvent(WE_SETFOCUS, _OnStorageSetfocusFnc);
	//m_wndStorage.SetEvent(WE_KILLFOCUS, _OnStorageKillfocusFnc);
	m_wndStorage.SetEvent(WE_SELCHANGE, _OnStorageSelectChangeFnc);
	m_wndStorage.SetEvent(WE_LOADDATA, _OnStorageLoadDataFnc);
	//m_wndStorage.SetEvent(WE_ADDNEW, _OnStorageAddNewFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CTMMaterialFinalAccount::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStorage.GetDlgCtrlID(), m_szStorageKey);
	DDX_Check(pDX, m_wndDrug.GetDlgCtrlID(), m_bDrug);
	DDX_Check(pDX, m_wndMaterial.GetDlgCtrlID(), m_bMaterial);

}
void CTMMaterialFinalAccount::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStorageKey.Empty();
	m_bDrug = false;
	m_bMaterial = false;

}
int CTMMaterialFinalAccount::SetMode(int nMode){
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
/*void CTMMaterialFinalAccount::OnFromDateChange(){
	
} */
/*void CTMMaterialFinalAccount::OnFromDateSetfocus(){
	
} */
/*void CTMMaterialFinalAccount::OnFromDateKillfocus(){
	
} */
int CTMMaterialFinalAccount::OnFromDateCheckValue(){
	return 0;
} 
/*void CTMMaterialFinalAccount::OnToDateChange(){
	
} */
/*void CTMMaterialFinalAccount::OnToDateSetfocus(){
	
} */
/*void CTMMaterialFinalAccount::OnToDateKillfocus(){
	
} */
int CTMMaterialFinalAccount::OnToDateCheckValue(){
	return 0;
} 
void CTMMaterialFinalAccount::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMMaterialFinalAccount::OnStorageSelendok(){
	 
}
/*void CTMMaterialFinalAccount::OnStorageSetfocus(){
	
}*/
/*void CTMMaterialFinalAccount::OnStorageKillfocus(){
	
}*/
long CTMMaterialFinalAccount::OnStorageLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStorage.IsSearchKey() && !m_szStorageKey.IsEmpty()){
		szWhere.Format(_T(" and msl_storage_id='%s' "), m_szStorageKey);
	}
	m_wndStorage.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT msl_storage_id as id, msl_name as name FROM m_storagelist ") \
				 _T("WHERE msl_isactive = 'Y' AND msl_type = 'C' AND msl_category IN ('I', 'S', 'P') %s ORDER BY id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStorage.AddItems(
			rs.GetValue(_T("id")),
			rs.GetValue(_T("name")),
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CTMMaterialFinalAccount::OnStorageAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CTMMaterialFinalAccount::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CExcel xls;
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 30);
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	nRow = 7;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, rs.GetValue(_T("product_name")), FMT_TEXT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("product_uomname")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("solqty_fp")), FMT_NUMBER1);
		xls.SetCellText(3, nRow, rs.GetValue(_T("insqty_fp")), FMT_NUMBER1);
		xls.SetCellText(4, nRow, rs.GetValue(_T("serqty_fp")), FMT_NUMBER1);
		xls.SetCellText(5, nRow, rs.GetValue(_T("totalqty_fp")), FMT_NUMBER1);
		xls.SetCellText(6, nRow, rs.GetValue(_T("totalamount_fp")), FMT_NUMBER1);
		xls.SetCellText(7, nRow, rs.GetValue(_T("solqty_import")), FMT_NUMBER1);
		xls.SetCellText(8, nRow, rs.GetValue(_T("insqty_import")), FMT_NUMBER1);
		xls.SetCellText(9, nRow, rs.GetValue(_T("serqty_import")), FMT_NUMBER1);
		xls.SetCellText(10, nRow, rs.GetValue(_T("totalqty_import")), FMT_NUMBER1);
		xls.SetCellText(11, nRow, rs.GetValue(_T("totalamount_import")), FMT_NUMBER1);
		xls.SetCellText(12, nRow, rs.GetValue(_T("solqty_export")), FMT_NUMBER1);
		xls.SetCellText(13, nRow, rs.GetValue(_T("inssolqty_export")), FMT_NUMBER1);
		xls.SetCellText(14, nRow, rs.GetValue(_T("insothqty_export")), FMT_NUMBER1);
		xls.SetCellText(15, nRow, rs.GetValue(_T("serqty_export")), FMT_NUMBER1);
		xls.SetCellText(16, nRow, rs.GetValue(_T("totalqty_export")), FMT_NUMBER1);
		xls.SetCellText(17, nRow, rs.GetValue(_T("totalamount_export")), FMT_NUMBER1);
		xls.SetCellText(18, nRow, rs.GetValue(_T("solqty_adjustimport")), FMT_NUMBER1);
		xls.SetCellText(19, nRow, rs.GetValue(_T("insqty_adjustimport")), FMT_NUMBER1);
		xls.SetCellText(20, nRow, rs.GetValue(_T("serqty_adjustimport")), FMT_NUMBER1);
		xls.SetCellText(21, nRow, rs.GetValue(_T("solqty_adjustexport")), FMT_NUMBER1);
		xls.SetCellText(22, nRow, rs.GetValue(_T("insqty_adjustexport")), FMT_NUMBER1);
		xls.SetCellText(23, nRow, rs.GetValue(_T("serqty_adjustexport")), FMT_NUMBER1);
		xls.SetCellText(24, nRow, rs.GetValue(_T("solqty_lp")), FMT_NUMBER1);
		xls.SetCellText(25, nRow, rs.GetValue(_T("insqty_lp")), FMT_NUMBER1);
		xls.SetCellText(26, nRow, rs.GetValue(_T("serqty_lp")), FMT_NUMBER1);
		xls.SetCellText(27, nRow, rs.GetValue(_T("totalqty_lp")), FMT_NUMBER1);
		xls.SetCellText(28, nRow, rs.GetValue(_T("totalamount_lp")), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\Quyet Toan Hoa Chat Vat Tu cho Xet nghiem.xls"));
		
} 

CString CTMMaterialFinalAccount::GetQueryString(){
	CString szSQL, szWhere, szFpWhere, szItemType, szStorages;
	szFpWhere.Format(_T(" AND approval_date < to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')")
		, m_szFromDate);
	szWhere.Format(_T(" AND approval_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')")
		, m_szFromDate, m_szToDate);
	for (int i = 0; i < m_wndStorage.GetItemCount(); i++)
	{
		if (m_wndStorage.GetCheck(i))
		{
			m_wndStorage.SetCurSel(i);
			if (!szStorages.IsEmpty())
				szStorages += _T(", ");
			szStorages += m_wndStorage.GetCurrent(0);
		}
	}
	if (!szStorages.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND storage_id IN (%s)"), szStorages);
		szFpWhere.AppendFormat(_T(" AND storage_id IN (%s)"), szStorages);
	}
	if (m_bDrug)
	{
		if (!szItemType.IsEmpty())
			szItemType += _T(", ");
		szItemType += _T("'PM'");
	}
	if (m_bDrug)
	{
		if (!szItemType.IsEmpty())
			szItemType += _T(", ");
		szItemType += _T("'MA'");
	}
	if (!szItemType.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND product_org_id IN (%s)"), szItemType);
		szFpWhere.AppendFormat(_T(" AND product_org_id IN (%s)"), szItemType);
	}
	szSQL.Format(_T(" SELECT    product_id, ") \
				_T("           product_name, ") \
				_T("		   product_uomname,") \
				_T("           solqty_fp, ") \
				_T("           insqty_fp, ") \
				_T("           serqty_fp, ") \
				_T("           totalqty_fp, ") \
				_T("           totalamount_fp, ") \
				_T("           solqty_import, ") \
				_T("           insqty_import, ") \
				_T("           serqty_import, ") \
				_T("           solqty_export, ") \
				_T("           inssolqty_export, ") \
				_T("           insothqty_export, ") \
				_T("           serqty_export, ") \
				_T("           solqty_adjustimport, ") \
				_T("           insqty_adjustimport, ") \
				_T("           serqty_adjustimport, ") \
				_T("           solqty_adjustexport, ") \
				_T("           insqty_adjustexport, ") \
				_T("           serqty_adjustexport, ") \
				_T("           totalqty_import, ") \
				_T("           totalamount_import, ") \
				_T("           totalqty_export, ") \
				_T("           totalamount_export, ") \
				_T("           Coalesce(solqty_fp, 0) ") \
				_T("           + Coalesce( solqty_import, 0) - Coalesce(solqty_export, 0) + Coalesce(solqty_adjustimport, 0) - Coalesce(solqty_adjustexport, 0) solqty_lp, ") \
				_T("           Coalesce(serqty_fp, 0) ") \
				_T("           + Coalesce( serqty_import, 0) - Coalesce(serqty_export, 0) + Coalesce(serqty_adjustimport, 0) - Coalesce(serqty_adjustexport, 0) serqty_lp, ") \
				_T("           Coalesce(insqty_fp, 0) ") \
				_T("           + Coalesce( insqty_import, 0) - Coalesce(inssolqty_export, 0) - Coalesce(insothqty_export, 0) + Coalesce(insqty_adjustimport, 0) - Coalesce(insqty_adjustexport, 0) insqty_lp, ") \
				_T("           Coalesce(totalqty_fp, 0) ") \
				_T("           + Coalesce( totalqty_import, 0) - Coalesce(totalqty_export, 0) + Coalesce(totalqty_adjustimport, 0) - Coalesce(totalqty_adjustexport, 0) totalqty_lp, ") \
				_T("           Coalesce(totalamount_fp, 0) ") \
				_T("           + Coalesce( totalamount_import, 0) - Coalesce(totalamount_export, 0) + Coalesce(totalamount_adjustimport, 0) - Coalesce(totalamount_adjustexport, 0) totalamount_lp ") \
				_T(" FROM      (SELECT    product_id, ") \
				_T("                      product_name, ") \
				_T("					  product_uomname,") \
				_T("                      SUM(CASE WHEN msl_category = 'P' THEN import_qty - export_qty ") \
				_T("                          ELSE 0 ") \
				_T("                          END) solqty_fp, ") \
				_T("                      SUM(CASE WHEN msl_category = 'I' THEN import_qty - export_qty ") \
				_T("                          ELSE 0 ") \
				_T("                          END) insqty_fp, ") \
				_T("                      SUM(CASE WHEN msl_category = 'S' THEN import_qty - export_qty ") \
				_T("                          ELSE 0 ") \
				_T("                          END) serqty_fp, ") \
				_T("                      SUM(import_qty - export_qty) totalqty_fp, ") \
				_T("                      SUM(( import_qty - export_qty ) * Decode(0, unit_price, Decode(msl_category, 'I', product_taxprice, ") \
				_T("                                                                                                   'S', product_saleprice1, ") \
				_T("                                                                                                   'P', product_saleprice3), ") \
				_T("                                                                  unit_price)) totalamount_fp ") \
				_T("            FROM      (SELECT storage_id, ") \
				_T("                              product_item_id sitem_id, ") \
				_T("                              qtyimport import_qty, ") \
				_T("                              0 export_qty, ") \
				_T("                              unit_price, ") \
				_T("                              approved_date approval_date ") \
				_T("                       FROM   m_import_view ") \
				_T("                       UNION ALL ") \
				_T("                       SELECT storage_id, ") \
				_T("                              product_item_id sitem_id, ") \
				_T("                              0 import_qty, ") \
				_T("                              qtyexport export_qty, ") \
				_T("                              unit_price, ") \
				_T("                              approved_date approval_date ") \
				_T("                       FROM   m_export_view) ") \
				_T("            LEFT JOIN m_productitem_view ON ( sitem_id = product_item_id ) ") \
				_T("            LEFT JOIN m_storagelist ON ( msl_storage_id = storage_id ) ") \
				_T("            WHERE     msl_type = 'C' AND msl_category IN ( 'I', 'S', 'P' ) %s ") \
				_T("            GROUP     BY product_id,product_name, product_uomname) tbl_fp ") \
				_T(" FULL JOIN (SELECT    product_id, ") \
				_T("                      SUM(CASE WHEN msl_category = 'P' THEN import_qty ") \
				_T("                          ELSE 0 ") \
				_T("                          END) solqty_import, ") \
				_T("                      SUM(CASE WHEN msl_category = 'I' THEN import_qty ") \
				_T("                          ELSE 0 ") \
				_T("                          END) insqty_import, ") \
				_T("                      SUM(CASE WHEN msl_category = 'S' THEN import_qty ") \
				_T("                          ELSE 0 ") \
				_T("                          END) serqty_import, ") \
				_T("                      SUM(CASE WHEN msl_category = 'P' THEN export_qty ") \
				_T("                          ELSE 0 ") \
				_T("                          END) solqty_export, ") \
				_T("                      SUM(CASE WHEN msl_category = 'I' AND object_id = 2 THEN export_qty ") \
				_T("                          ELSE 0 ") \
				_T("                          END) inssolqty_export, ") \
				_T("                      SUM(CASE WHEN msl_category = 'I' AND object_id IN ( 4, 5, 10, 11, 12 ) THEN export_qty ") \
				_T("                          ELSE 0 ") \
				_T("                          END) insothqty_export, ") \
				_T("                      SUM(CASE WHEN msl_category = 'S' THEN export_qty ") \
				_T("                          ELSE 0 ") \
				_T("                          END) serqty_export, ") \
				_T("                      SUM(CASE WHEN msl_category = 'P' THEN adjust_qtyimport ") \
				_T("                          ELSE 0 ") \
				_T("                          END) solqty_adjustimport, ") \
				_T("                      SUM(CASE WHEN msl_category = 'I' THEN adjust_qtyimport ") \
				_T("                          ELSE 0 ") \
				_T("                          END) insqty_adjustimport, ") \
				_T("                      SUM(CASE WHEN msl_category = 'S' THEN adjust_qtyimport ") \
				_T("                          ELSE 0 ") \
				_T("                          END) serqty_adjustimport, ") \
				_T("                      SUM(CASE WHEN msl_category = 'P' THEN adjust_qtyexport ") \
				_T("                          ELSE 0 ") \
				_T("                          END) solqty_adjustexport, ") \
				_T("                      SUM(CASE WHEN msl_category = 'I' THEN adjust_qtyexport ") \
				_T("                          ELSE 0 ") \
				_T("                          END) insqty_adjustexport, ") \
				_T("                      SUM(CASE WHEN msl_category = 'S' THEN adjust_qtyexport ") \
				_T("                          ELSE 0 ") \
				_T("                          END) serqty_adjustexport, ") \
				_T("                      SUM(import_qty) totalqty_import, ") \
				_T("                      SUM(import_qty * Decode(0, unit_price, Decode(msl_category, 'I', product_taxprice, ") \
				_T("                                                                                  'S', product_saleprice1, ") \
				_T("                                                                                  'P', product_saleprice3), ") \
				_T("                                                 unit_price)) totalamount_import, ") \
				_T("                      SUM(export_qty) totalqty_export, ") \
				_T("                      SUM(export_qty * Decode(0, unit_price, Decode(msl_category, 'I', product_taxprice, ") \
				_T("                                                                                  'S', product_saleprice1, ") \
				_T("                                                                                  'P', product_saleprice3), ") \
				_T("                                                 unit_price)) totalamount_export, ") \
				_T("                      SUM(adjust_qtyimport) totalqty_adjustimport, ") \
				_T("                      SUM(adjust_qtyimport * Decode(msl_category, 'I', product_taxprice, ") \
				_T("                                                                  'S', product_saleprice1, ") \
				_T("                                                                  'P', product_saleprice3)) totalamount_adjustimport, ") \
				_T("                      SUM(adjust_qtyexport) totalqty_adjustexport, ") \
				_T("                      SUM(adjust_qtyexport * Decode(msl_category, 'I', product_taxprice, ") \
				_T("                                                                  'S', product_saleprice1, ") \
				_T("                                                                  'P', product_saleprice3)) totalamount_adjustexport ") \
				_T("            FROM      (SELECT storage_id, ") \
				_T("                              product_item_id sitem_id, ") \
				_T("                              CASE WHEN iotype = 'CSO' THEN qtyimport ") \
				_T("                              ELSE 0 ") \
				_T("                              END import_qty, ") \
				_T("                              0 export_qty, ") \
				_T("                              CASE WHEN iotype = 'ADO' THEN qtyimport ") \
				_T("                              ELSE 0 ") \
				_T("                              END adjust_qtyimport, ") \
				_T("                              0 adjust_qtyexport, ") \
				_T("                              unit_price, ") \
				_T("                              approved_date approval_date, ") \
				_T("                              0 object_id ") \
				_T("                       FROM   m_import_view ") \
				_T("                       WHERE  iotype IN ( 'CSO', 'ADO' ) ") \
				_T("                       UNION ALL ") \
				_T("                       SELECT storage_id, ") \
				_T("                              product_item_id sitem_id, ") \
				_T("                              0 import_qty, ") \
				_T("                              0, ") \
				_T("                              0, ") \
				_T("                              CASE WHEN iotype = 'ADO' THEN qtyexport ") \
				_T("                              ELSE 0 ") \
				_T("                              END, ") \
				_T("                              unit_price, ") \
				_T("                              approved_date approval_date, ") \
				_T("                              0 ") \
				_T("                       FROM   m_export_view ") \
				_T("                       WHERE  iotype = 'ADO' ") \
				_T("                       UNION ALL ") \
				_T("                       SELECT    hpo_storage_id, ") \
				_T("                                 hpol_product_item_id, ") \
				_T("                                 0, ") \
				_T("                                 hpol_qtyissue, ") \
				_T("                                 0, ") \
				_T("                                 0, ") \
				_T("                                 hpol_unitprice, ") \
				_T("                                 hpo_approvedate, ") \
				_T("                                 hpo_object_id ") \
				_T("                       FROM      hms_pharmaorder_view ") \
				_T("                       LEFT JOIN hms_pharmaorderline_view2 ON ( hpo_orderid = hpol_orderid ) ") \
				_T("                       WHERE     hpo_ordertype = 'C') ") \
				_T("            LEFT JOIN m_productitem_view ON ( sitem_id = product_item_id ) ") \
				_T("            LEFT JOIN m_storagelist ON ( msl_storage_id = storage_id ) ") \
				_T("            WHERE     msl_type = 'C' AND msl_category IN ( 'I', 'S', 'P' ) %s ") \
				_T("            GROUP     BY product_id) tbl_io USING (product_id) "), szFpWhere, szWhere);
	return szSQL;
}