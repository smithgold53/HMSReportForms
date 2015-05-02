#include "stdafx.h"
#include "PMSDrugDetailBook.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include ".\pmsdrugdetailbook.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSDrugDetailBook *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSDrugDetailBook *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSDrugDetailBook* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CPMSDrugDetailBook *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnStockAddNew();
}*/
/*static void _OnNameChangeFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnNameChange();
} */
/*static void _OnNameSetfocusFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnNameSetfocus();} */ 
/*static void _OnNameKillfocusFnc(CWnd *pWnd){
	((CPMSDrugDetailBook *)pWnd)->OnNameKillfocus();
} */
static int _OnNameCheckValueFnc(CWnd *pWnd){
	return ((CPMSDrugDetailBook *)pWnd)->OnNameCheckValue();
} 


static int _OnListSelectAllFnc(CWnd *pWnd){
	return ((CPMSDrugDetailBook*)pWnd)->OnListSelectAll();
} 
static int _OnListUnselectAllFnc(CWnd *pWnd){
	return ((CPMSDrugDetailBook*)pWnd)->OnListUnselectAll();
} 
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CPMSDrugDetailBook*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CPMSDrugDetailBook*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPMSDrugDetailBook*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CPMSDrugDetailBook*)pWnd)->OnListDeleteItem();
} 
static void _OnImportDiarySelectFnc(CWnd *pWnd){
	((CPMSDrugDetailBook*)pWnd)->OnImportDiarySelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSDrugDetailBook *pVw = (CPMSDrugDetailBook *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CPMSDrugDetailBook *pVw = (CPMSDrugDetailBook *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSDrugDetailBook *pVw = (CPMSDrugDetailBook *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CPMSDrugDetailBook *pVw = (CPMSDrugDetailBook *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddPMSDrugDetailBookFnc(CWnd *pWnd){
	 return ((CPMSDrugDetailBook*)pWnd)->OnAddPMSDrugDetailBook();
} 
static int _OnEditPMSDrugDetailBookFnc(CWnd *pWnd){
	 return ((CPMSDrugDetailBook*)pWnd)->OnEditPMSDrugDetailBook();
} 
static int _OnDeletePMSDrugDetailBookFnc(CWnd *pWnd){
	 return ((CPMSDrugDetailBook*)pWnd)->OnDeletePMSDrugDetailBook();
} 
static int _OnSavePMSDrugDetailBookFnc(CWnd *pWnd){
	 return ((CPMSDrugDetailBook*)pWnd)->OnSavePMSDrugDetailBook();
} 
static int _OnCancelPMSDrugDetailBookFnc(CWnd *pWnd){
	 return ((CPMSDrugDetailBook*)pWnd)->OnCancelPMSDrugDetailBook();
} 
CPMSDrugDetailBook::CPMSDrugDetailBook(CWnd* pParent){

	m_nDlgWidth = 585;
	m_nDlgHeight = 590;
	SetDefaultValues();
}
CPMSDrugDetailBook::~CPMSDrugDetailBook(){
}
void CPMSDrugDetailBook::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 550);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 110, 55);
	m_wndFromDate.Create(this,115, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 395, 55);
	m_wndToDate.Create(this,400, 30, 570, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 110, 85);
	m_wndStock.Create(this,115, 60, 570, 85); 
	m_wndNameLabel.Create(this, _T("Name /Cnt"), 10, 90, 110, 115);
	m_wndName.Create(this,115, 90, 570, 115); 
	m_wndList.Create(this,10, 120, 570, 545); 
	m_wndImportDiary.Create(this, _T("Import Diary"), 5, 555, 305, 580);
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 385, 555, 495, 580);
	m_wndExport.Create(this, _T("&Export"), 500, 555, 575, 580);

}
void CPMSDrugDetailBook::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);
	m_wndName.SetLimitText(35);
	m_wndName.SetCheckValue(true);
	m_wndName.SetNotEmpty(false);

	m_wndStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStock.InsertColumn(1, _T("Description"), CFMT_TEXT, 220);

	m_wndList.SetCheckBox(TRUE);
	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndList.InsertColumn(1, _T("Name /Cnt"), CFMT_TEXT, 150);
	m_wndList.InsertColumn(2, _T("Unit"), CFMT_TEXT, 80);
	m_wndList.InsertColumn(3, _T("Generic"), CFMT_TEXT, 100);
	m_wndList.InsertColumn(4, _T("Org Name"), CFMT_TEXT, 80);
	m_wndList.InsertColumn(5, _T("Product ID"), CFMT_TEXT, 0);
}
void CPMSDrugDetailBook::OnSetWindowEvents(){
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
	//m_wndName.SetEvent(WE_CHANGE, _OnNameChangeFnc);
	//m_wndName.SetEvent(WE_SETFOCUS, _OnNameSetfocusFnc);
	//m_wndName.SetEvent(WE_KILLFOCUS, _OnNameKillfocusFnc);
	m_wndName.SetEvent(WE_CHECKVALUE, _OnNameCheckValueFnc);
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);
	m_wndList.SetWindowText(_T("Drug List"));
	m_wndList.AddEvent(1, _T("Select All"), _OnListSelectAllFnc);
	m_wndList.AddEvent(2, _T("Unselect All"), _OnListUnselectAllFnc);
	m_wndImportDiary.SetEvent(WE_CLICK, _OnImportDiarySelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_szFromDate.Format(_T("%s 00:00"), pMF->GetSysDate());
	m_szToDate = pMF->GetSysDate() + _T("23:59");
	SetMode(VM_EDIT);

}
void CPMSDrugDetailBook::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_Text(pDX, m_wndName.GetDlgCtrlID(), m_szName);
	DDX_Check(pDX, m_wndImportDiary.GetDlgCtrlID(), m_bImportDiary);

}
void CPMSDrugDetailBook::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPMSDrugDetailBook::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CPMSDrugDetailBook::SetDefaultValues()
{

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_szName.Empty();

}
int CPMSDrugDetailBook::SetMode(int nMode)
{
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
/*void CPMSDrugDetailBook::OnFromDateChange(){
	
} */
/*void CPMSDrugDetailBook::OnFromDateSetfocus(){
	
} */
/*void CPMSDrugDetailBook::OnFromDateKillfocus(){
	
} */
int CPMSDrugDetailBook::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSDrugDetailBook::OnToDateChange(){
	
} */
/*void CPMSDrugDetailBook::OnToDateSetfocus(){
	
} */
/*void CPMSDrugDetailBook::OnToDateKillfocus(){
	
} */
int CPMSDrugDetailBook::OnToDateCheckValue(){
	return 0;
} 
void CPMSDrugDetailBook::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSDrugDetailBook::OnStockSelendok(){
	UpdateData(true);
	OnListLoadData();
	 
}
/*void CPMSDrugDetailBook::OnStockSetfocus(){
	
}*/
/*void CPMSDrugDetailBook::OnStockKillfocus(){
	
}*/
long CPMSDrugDetailBook::OnStockLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageList(&m_wndStock, m_szStockKey);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStock.IsSearchKey() && ToInt(m_szStockKey) > 0){
		szWhere.Format(_T(" AND msl_storage_id=%d"), ToInt(m_szStockKey));
	};

	//szWhere.AppendFormat(_T(" and msl_storage_id in(%s) "), pMF->m_szStorages);

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
/*void CPMSDrugDetailBook::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CPMSDrugDetailBook::OnNameChange(){
	
} */
/*void CPMSDrugDetailBook::OnNameSetfocus(){
	
} */
/*void CPMSDrugDetailBook::OnNameKillfocus(){
	
} */
int CPMSDrugDetailBook::OnNameCheckValue(){
	OnListLoadData();
	return 1;
}

void CPMSDrugDetailBook::OnImportDiarySelect(){
	
}

int CPMSDrugDetailBook::OnListSelectAll(){
	for (int i =0; i < m_wndList.GetItemCount(); i++)
	{
			m_wndList.SetCheck(i, TRUE);
	}
	return 0;
} 

int CPMSDrugDetailBook::OnListUnselectAll(){
	for (int i =0; i < m_wndList.GetItemCount(); i++)
	{
		m_wndList.SetCheck(i, FALSE);
	}
	return 0;
} 
void CPMSDrugDetailBook::OnListDblClick(){
	UpdateData(TRUE);
	int nSel = m_wndList.GetCurSel();
	if(nSel < 0) return;
	for (int i=0; i< m_wndList.GetItemCount(); i++)
	{
		m_wndList.SetCheck(i,false);
	}
	m_wndList.SetCheck(nSel, !m_wndList.GetCheck(nSel));
} 
void CPMSDrugDetailBook::OnListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CPMSDrugDetailBook::OnListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
} 
long CPMSDrugDetailBook::OnListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	UpdateData(true);

	m_wndList.BeginLoad(); 
	int nCount = 0;
	if(!m_szName.IsEmpty())
	{
		szWhere.Format(_T(" and lower(mp_name) like(lower('%s%s%s')) "), _T("%"), m_szName, _T("%"));
	}
	szWhere.AppendFormat(_T(" AND mp_org_id = '%s'"), pMF->GetModuleID());
	szSQL.Format(_T(" SELECT") \
	_T("   DISTINCT") \
	_T("   mp_code                                                         AS id,") \
	_T("   mp_name                                                         AS name,") \
	_T("   Get_uomname(mp_uom_id)                                          AS unit,") \
	_T("   Get_productclassname(mp_product_class_id)                       AS genericname,") \
	_T("   Get_countryname((SELECT") \
	_T("                      DISTINCT") \
	_T("                      adm_country_id") \
	_T("                    FROM   ad_manufacture") \
	_T("                    WHERE  adm_manufacture_id=mp_manufacture_id)) AS orgname,") \
	_T("   mp_product_id as productid") \
	_T(" FROM   m_storageline") \
	_T("        LEFT JOIN m_product ON(msl_product_id=mp_product_id)") \
	_T(" WHERE  msl_storage_id=%d") \
	_T("        AND msl_product_item_id>0") \
	_T("        AND msl_qtyonhand>0 %s") \
	_T(" ORDER  BY name,") \
	_T("           unit "), ToInt(m_szStockKey), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndList.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("Name")), 
			rs.GetValue(_T("Unit")), 
			rs.GetValue(_T("GenericName")), 
			rs.GetValue(_T("OrgName")), 
			rs.GetValue(_T("ProductID")),
			NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
}
void CPMSDrugDetailBook::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CString szSQL, szWhere, tmpStr, szTemp, szCode, szType, szCategory, szName, szPrice, szUnit, szSysDate, szRptName, szID;
	int nIdx = 0;
	double nTtlImpQty =0, nTtlExpQty=0, nTtlOnhandQty = 0;
	double nGrpImpQty =0, nGrpExpQty=0, nGrpOnhandQty = 0;
	double nTtlImpAmt = 0, nTtlExpAmt = 0, nTtlOnhandAmt = 0;
	double nGrpImpAmt = 0, nGrpExpAmt = 0, nGrpOnhandAmt = 0;
	double nAmount = 0, nTotal = 0, nOnhand = 0, nPrice = 0;
	szRptName.Format(_T("Reports/HMS/PMS_SOCHITIETVATTU.rpt"));
	if (!rpt.Init(szRptName))
		return;
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection *rptDetail;
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if(m_wndList.GetCheck(i)){
			szCode = m_wndList.GetItemText(i, 0);
			szID = m_wndList.GetItemText(i, 5);
			szSQL.Format(_T(" SELECT sum(impqty-expqty) as onhand, mp_name as name,") \
				_T(" mpi_unitprice as price, get_uomname(mp_uom_id) as unit, sum(impqty-expqty)*mpi_unitprice as total ") \
				_T(" FROM") \
				_T(" (") \
				_T(" 	SELECT impstockid, impdate as iodate, sitemid, impqty, 0 as expqty") \
				_T(" 	FROM miv") \
				_T(" 	UNION ALL") \
				_T(" 	SELECT expstockid, expdate, sitemid, 0, expqty") \
				_T(" 	FROM mev") \
				_T(" ) tbl") \
				_T(" LEFT JOIN m_product_item ON(sitemid=mpi_product_item_id)") \
				_T(" LEFT JOIN m_product ON(mp_product_id=mpi_product_id)") \
				_T(" WHERE 	impstockid=%d ") \
				_T(" 	AND iodate < cast_string2timestamp('%s') ") \
				_T(" 	AND mpi_product_id=%s GROUP BY mp_name, mpi_unitprice, mp_uom_id") , ToInt(m_szStockKey), m_szFromDate, szID);
			rs.ExecSQL(szSQL);
			if(!rs.IsEOF())
			{
				rs.GetValue(_T("onhand"), nOnhand);
				rs.GetValue(_T("name"), szName);
				rs.GetValue(_T("price"), nPrice);
				FormatCurrency(nPrice, szPrice);
				rs.GetValue(_T("unit"), szUnit);
				rs.GetValue(_T("total"), nTotal);
			}
			else
				nOnhand = 0;
			if (m_bImportDiary)
				szWhere.Format(_T(" AND iotype = 'POO'"));
			szSQL.Format(_T(" SELECT get_storagename(impstockid) as stockname,") \
									_T("   get_departmentname(deptid) as dept,") \
									_T("   Cast_Timestamp2Date(iodate) AS invoicedate,") \
									_T("   invoiceno,") \
									_T("   invoiceid,") \
									_T("   SUM(qty)     AS qty,") \
									_T("   product_code       AS itemid,") \
									_T("   product_name       AS name,") \
									_T("   product_uomname       AS unit,") \
									_T("   product_unitprice       AS price,") \
									_T("   product_lotno   AS lotno,") \
									_T("   product_expdate AS expdate,") \
									_T("   iotype,") \
									_T("   category") \
									_T(" FROM") \
									_T("   (SELECT impstockid,") \
									_T("     deptid,") \
									_T("     impdate    AS iodate,") \
									_T("     impinvoice AS invoiceno,") \
									_T("	 invoiceid,	") \
									_T("     sitemid,") \
									_T("     impqty AS qty,") \
									_T("     iotype,") \
									_T("     CAST('I' AS NVARCHAR2(1)) as category") \
									_T("   FROM MIV") \
									_T("   UNION ALL") \
									_T("   SELECT expstockid,") \
									_T("     deptid,") \
									_T("     expdate,") \
									_T("     expinvoice,") \
									_T("	 invoiceid,") \
									_T("     sitemid,") \
									_T("     expqty,") \
									_T("     iotype,") \
									_T("     category") \
									_T("   FROM MEV") \
									_T("   ) iotbl") \
									_T(" LEFT JOIN m_productitem_view ON(product_item_id=sitemid)") \
									_T(" WHERE impstockid=%d") \
									_T(" AND iodate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') AND product_id='%s' %s") \
									_T(" GROUP BY impstockid, deptid, iodate, invoiceno, invoiceid, product_code, product_name, iotype, category, product_lotno, product_expdate, product_uomname, product_unitprice") \
									_T(" ORDER BY iodate"), ToInt(m_szStockKey), m_szFromDate, m_szToDate, szID, szWhere);

			//QueryOpener(szSQL);
			rs.ExecSQL(szSQL);
			if (rs.IsEOF())
				continue;
			if (nOnhand == 0)
			{
				rs.GetValue(_T("name"), szName);
				rs.GetValue(_T("unit"), szUnit);
				rs.GetValue(_T("price"), nPrice);
				FormatCurrency(nPrice, szPrice);
			}
			tmpStr.Format(_T("M\xE3 h\xE0ng: %s \t T\xEAn h\xE0ng: %s"), szCode, szName);
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			rptDetail->SetValue(_T("GroupName"), tmpStr);
			if (nOnhand != 0 && !m_bImportDiary)
			{
				rptDetail = rpt.AddDetail();
				rptDetail->SetValue(_T("3"), _T("S\x1ED1 \x64\x1B0 \x111\x1EA7u k\x1EF3"));
				rptDetail->SetValue(_T("5"), szUnit);	
				rptDetail->SetValue(_T("6"), szPrice);	
				FormatCurrency(nOnhand, tmpStr);	
				rptDetail->SetValue(_T("13"), tmpStr);
				FormatCurrency(nTotal, tmpStr);
				rptDetail->SetValue(_T("14"), tmpStr);
			}
			while (!rs.IsEOF()){
				//data from sql
				rptDetail = rpt.AddDetail();
				nIdx++;
				rptDetail->SetValue(_T("1"), int2str(nIdx));
				rs.GetValue(_T("invoicedate"), tmpStr);
				rptDetail->SetValue(_T("3"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
				tmpStr.Empty();
				rs.GetValue(_T("iotype"), szType);
				rs.GetValue(_T("category"), szCategory);
				rs.GetValue(_T("invoiceno"), tmpStr);
				//if(szType == _T("P") || szType == _T("I"))
				//if (szType == _T("POO"))
				//{					
				//	szSQL.Format(_T("select po_invoiceno as invoiceno from purchase_order where po_purchase_order_id = '%s'"), tmpStr);
				//	rs2.ExecSQL(szSQL);					
				//	if (!rs2.IsEOF()) rs2.GetValue(_T("invoiceno"), tmpStr);			
				//}

				if(szCategory == _T("I"))
				{
					rptDetail->SetValue(_T("2"), tmpStr);
					//if(szType == _T("A") || szType == _T("C"))
					if (szType == _T("MOV"))
					{
						rptDetail->SetValue(_T("4"), rs.GetValue(_T("stockname")));
					}
					else if(szType == _T("POO"))
					{
						szSQL.Format(_T("SELECT get_partnername(po_partner_id) as pname FROM purchase_order WHERE po_purchase_order_id = '%s'"), rs.GetValue(_T("invoiceid")));
						rs2.ExecSQL(szSQL);
						rptDetail->SetValue(_T("4"), rs2.GetValue(_T("pname")));
					}	
					else
					{
						rptDetail->SetValue(_T("4"), rs.GetValue(_T("dept")));
					}	
					rptDetail->SetValue(_T("5"), szUnit);
					rptDetail->SetValue(_T("6"), szPrice);
					rs.GetValue(_T("qty"), tmpStr);
					nAmount = CalcAmt(tmpStr, nPrice);
					rptDetail->SetValue(_T("7"), tmpStr);
					nOnhand += ToLong(tmpStr);
					nGrpImpQty += ToLong(tmpStr);
					if (!m_bImportDiary)
					{
						FormatCurrency(nAmount, tmpStr);
						rptDetail->SetValue(_T("8"), tmpStr);
					}
				}
				else
				{
					rptDetail->SetValue(_T("2"), tmpStr);
					tmpStr.Empty();
					if(szType == _T("DDO") || szType == _T("DPO"))
					{
						rptDetail->SetValue(_T("4"), rs.GetValue(_T("dept")));
					}
					else if(szType == _T("MOV") || szType == _T("MRO"))
					{
						rptDetail->SetValue(_T("4"), rs.GetValue(_T("stockname")));
					}
					else if(szType == _T("PPO"))
					{
						CString szName;
						rs.GetValue(_T("invoiceno"), tmpStr);
						szSQL.Format(_T("select trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname from hms_patient left join hms_doc on(hp_patientno=hd_patientno) where hd_docno=%ld"), ToLong(tmpStr));
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("pname"), tmpStr);
						szName.Format(_T("[%s] %s"), rs.GetValue(_T("dept")), tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);
					}
					else if (szType == _T("SOO"))
					{
						rptDetail->SetValue(_T("4"), _T("\x110\x1A1n \x62\xE1n l\x1EBB")); 
					}
					else
					{
						szSQL.Format(_T("SELECT get_doctype('%s') as doctype FROM dual"), szType);
						rs2.ExecSQL(szSQL);
						rptDetail->SetValue(_T("4"), rs2.GetValue(_T("doctype")));
					}
					rptDetail->SetValue(_T("5"), szUnit);
					rptDetail->SetValue(_T("6"), szPrice);
					rs.GetValue(_T("qty"), tmpStr);
					nAmount = CalcAmt(tmpStr, nPrice);
					rptDetail->SetValue(_T("9"), tmpStr);
					nOnhand -= ToLong(tmpStr);
					nGrpExpQty += ToLong(tmpStr);
					FormatCurrency(nAmount, tmpStr);
					rptDetail->SetValue(_T("10"), tmpStr);
				}
				if (!m_bImportDiary)
				{
					FormatCurrency(nOnhand, tmpStr);
					rptDetail->SetValue(_T("13"), tmpStr);
					nAmount = nOnhand*nPrice;
					FormatCurrency(nAmount, tmpStr);
					rptDetail->SetValue(_T("14"), tmpStr);
					//rs.GetValue(_T("lotno"), tmpStr);
					//rs.GetValue(_T("expdate"), tmpStr);
				}
				rs.MoveNext();
			}
			if(nGrpImpQty + nGrpExpQty > 0){
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				CString szLabel;
				/*TranslateString(_T("Sum"), szLabel);
				xls.SetCellText(2, nRow, szLabel, FMT_TEXT | FMT_CENTER, true, 9);*/
				FormatCurrency(nGrpImpQty, tmpStr);
				rptDetail->SetValue(_T("s7"), tmpStr);
				if (!m_bImportDiary)
				{
					nAmount = nGrpImpQty*nPrice;
					FormatCurrency(nAmount, tmpStr);
					rptDetail->SetValue(_T("s8"), tmpStr);
					nTtlImpAmt += nAmount;
					FormatCurrency(nGrpExpQty, tmpStr);
					rptDetail->SetValue(_T("s9"), tmpStr);
					nAmount = nGrpExpQty*nPrice;
					FormatCurrency(nAmount, tmpStr);
					rptDetail->SetValue(_T("s10"), tmpStr);
					nTtlExpAmt += nAmount;
					FormatCurrency(nOnhand, tmpStr);
					rptDetail->SetValue(_T("s13"), tmpStr);
					nAmount = nOnhand*nPrice;
					FormatCurrency(nAmount, tmpStr);
					rptDetail->SetValue(_T("s14"), tmpStr);
				}
				nTtlOnhandAmt += nAmount;
				nTtlExpQty += nGrpExpQty;
				nTtlImpQty += nGrpImpQty;
				nTtlOnhandQty += nOnhand;
				nGrpExpQty = nGrpImpQty = nGrpOnhandQty = 0;
				nGrpExpAmt = nGrpExpAmt = nGrpOnhandAmt = 0;
			}
		}
	}
	if(nTtlImpQty + nTtlExpQty > 0){
		CString szLabel;
		//TranslateString(_T("Total"), szLabel);
		//xls.SetCellText(2, nRow, szLabel, FMT_TEXT | FMT_CENTER, true, 9);
		FormatCurrency(nTtlImpQty, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("sum7"), tmpStr);
		if (!m_bImportDiary)
		{
			FormatCurrency(nTtlImpAmt, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("sum8"), tmpStr);
			FormatCurrency(nTtlExpQty, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("sum9"), tmpStr);
			FormatCurrency(nTtlExpAmt, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("sum10"), tmpStr);
			FormatCurrency(nTtlOnhandQty, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("sum13"), tmpStr);
			FormatCurrency(nTtlOnhandAmt, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("sum14"), tmpStr);
		}
	}
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}
void CPMSDrugDetailBook::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSDrugDetailBook::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CString szSQL, tmpStr, szTemp, szCode, szType, szCategory, szName, szPrice, szUnit;
	double nTtlImpQty =0, nTtlExpQty=0, nTtlOnhandQty = 0;
	double nGrpImpQty =0, nGrpExpQty=0, nGrpOnhandQty = 0;
	double nTtlImpAmt = 0, nTtlExpAmt = 0, nTtlOnhandAmt = 0;
	double nGrpImpAmt = 0, nGrpExpAmt = 0, nGrpOnhandAmt = 0;
	int nCol = 0, nRow = 0;
	double nAmount = 0, nTotal = 0, nOnhand = 0, nPrice = 0;
	if (!xls.Load(_T("Exports/Sochitietvattu.xls")))
	{
		ShowMessageBox(_T("Missing file: SoChiTietVatTu.xls"));
		return;
	}
	//xls.CreateSheet(1);
	xls.SetWorksheet(0);
	CellFormat df(&xls), cf(&xls), hf(&xls);
	df.SetItalic(true);
	df.SetFontSize(10);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER);
	hf.SetFontSize(10);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 3, 14);
	xls.SetCellText(nCol, nRow, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, &df);	
	nRow = 4;
	if (!m_szStockKey.IsEmpty())
	{
		tmpStr.Format(_T("M\xE3 kho: %s"), m_szStockKey);
		xls.SetCellMergedColumns(0, nRow, 16);
		xls.SetCellText(0, nRow, tmpStr, &cf);
		nRow++;
		tmpStr.Format(_T("T\xEAn kho: %s"), m_wndStock.GetCurrent(1));
		xls.SetCellMergedColumns(0, nRow, 16);
		xls.SetCellText(0, nRow, tmpStr, &cf);
		nRow++;
	}
	//#1
	//tinh ton dau ky
	nRow = 9;
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if(m_wndList.GetCheck(i)){
			szCode = m_wndList.GetItemText(i, 0);
			szSQL.Format(_T(" SELECT sum(impqty-expqty) as onhand, pmi_name as name,") \
				_T(" pmsi_unitprice as price, pmi_unit as unit, sum(impqty-expqty)*pmsi_unitprice as total ") \
				_T(" FROM") \
				_T(" (") \
				_T(" 	SELECT impstockid, impdate as iodate, sitemid, impqty, 0 as expqty, moduleid") \
				_T(" 	FROM pmsv_import") \
				_T(" 	UNION ALL") \
				_T(" 	SELECT expstockid, expdate, sitemid, 0, expqty, moduleid") \
				_T(" 	FROM pmsv_export") \
				_T(" ) tbl") \
				_T(" LEFT JOIN pms_stockitems ON(sitemid=pmsi_id)") \
				_T(" LEFT JOIN pms_items ON(pmi_id=pmsi_itemid)") \
				_T(" WHERE moduleid = '%s' AND impstockid=%d ") \
				_T(" 	AND iodate < cast_string2timestamp('%s') ") \
				_T(" 	AND pmsi_itemid='%s' GROUP BY pmi_name, pmsi_unitprice, pmi_unit"),pMF->GetModuleID(), ToInt(m_szStockKey), m_szFromDate, szCode);
			rs.ExecSQL(szSQL);
			if(!rs.IsEOF())	
			{
				rs.GetValue(_T("onhand"), nOnhand);
				rs.GetValue(_T("name"), szName);
				rs.GetValue(_T("price"), nPrice);
				FormatCurrency(nPrice, szPrice);
				rs.GetValue(_T("unit"), szUnit);
				rs.GetValue(_T("total"), nTotal);
			}
			else
				nOnhand = 0;
			//#2
			szSQL.Format(_T(" SELECT expstockid,") \
									_T("   deptid,") \
									_T("   iodate AS invoicedate,") \
									_T("   invoiceno,") \
									_T("   SUM(qty)     AS qty,") \
									_T("   pmi_id       AS itemid,") \
									_T("   pmi_name       AS name,") \
									_T("   pmi_unit       AS unit,") \
									_T("   pmi_unitprice       AS price,") \
									_T("   pmsi_lotno   AS lotno,") \
									_T("   pmsi_expdate AS expdate,") \
									_T("   iotype,") \
									_T("   category") \
									_T(" FROM") \
									_T("   (SELECT impstockid,") \
									_T("     expstockid,") \
									_T("     deptid,") \
									_T("     impdate    AS iodate,") \
									_T("     impinvoice AS invoiceno,") \
									_T("     sitemid,") \
									_T("     impqty AS qty,") \
									_T("     iotype,") \
									_T("     CAST('I' AS NVARCHAR2(1)) as category,") \
									_T("     moduleid") \
									_T("   FROM pmsv_import") \
									_T("   UNION ALL") \
									_T("   SELECT expstockid,") \
									_T("     impstockid,") \
									_T("     deptid,") \
									_T("     expdate,") \
									_T("     expinvoice,") \
									_T("     sitemid,") \
									_T("     expqty,") \
									_T("     iotype,") \
									_T("     category,") \
									_T("     moduleid") \
									_T("   FROM pmsv_export") \
									_T("   ) iotbl") \
									_T(" LEFT JOIN pms_stockitems ON(pmsi_id=sitemid)") \
									_T(" LEFT JOIN pms_items ON(pmi_id     =pmsi_itemid)") \
									_T(" WHERE moduleid = '%s' AND impstockid=%d") \
									_T(" AND iodate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') AND pmi_id='%s'") \
									_T(" GROUP BY expstockid, deptid, iodate, invoiceno, pmi_id, pmi_name, iotype, category, pmsi_lotno, pmsi_expdate, pmi_unit, pmi_unitprice") \
									_T(" ORDER BY iodate"),pMF->GetModuleID(), ToInt(m_szStockKey), m_szFromDate, m_szToDate, szCode);
			rs.ExecSQL(szSQL);
			//ma hang+ten
			xls.SetCellMergedColumns(0, nRow, 16);
			if (nOnhand == 0)
			{
				rs.GetValue(_T("name"), szName);
				rs.GetValue(_T("unit"), szUnit);
				rs.GetValue(_T("price"), nPrice);
				FormatCurrency(nPrice, szPrice);
			}
			tmpStr.Format(_T("M\xE3 h\xE0ng: %s \t T\xEAn h\xE0ng: %s"), szCode, szName);
			xls.SetCellText(0, nRow, tmpStr, FMT_TEXT, true, 9);
			nRow++;
			if (nOnhand != 0)
			{
				xls.SetCellText(2, nRow, _T("S\x1ED1 \x64\x1B0 \x111\x1EA7u k\x1EF3"), FMT_TEXT, true, 9);
				xls.SetCellText(4, nRow, szUnit, FMT_TEXT | FMT_CENTER, false, 9);			
				xls.SetCellText(5, nRow, szPrice, FMT_TEXT | FMT_RIGHT, true, 9);
				FormatCurrency(nOnhand, tmpStr);
				xls.SetCellText(12, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
				FormatCurrency(nTotal, tmpStr);
				xls.SetCellText(13, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
			}
			while (!rs.IsEOF()){
				//data from sql
				rs.GetValue(_T("invoicedate"), tmpStr);
				xls.SetCellText(1, nRow, CDateTime::Convert(tmpStr, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), FMT_DATETIME | FMT_CENTER, false, 9);
				tmpStr.Empty();
				rs.GetValue(_T("iotype"), szType);
				rs.GetValue(_T("category"), szCategory);
				rs.GetValue(_T("invoiceno"), tmpStr);
				if(szType == _T("P") || szType == _T("I"))
				{					
					szSQL.Format(_T("select pmpi_invoiceno as invoiceno from pms_purchaseinvoice where pmpi_id = '%s'"), tmpStr);
					rs2.ExecSQL(szSQL);					
					if (!rs2.IsEOF()) rs2.GetValue(_T("invoiceno"), tmpStr);			
				}

				if(szCategory == _T("I"))
				{
					xls.SetCellText(0, nRow, tmpStr, FMT_TEXT | FMT_CENTER, false, 9);
					if(szType == _T("A") || szType == _T("C"))
					{
						rs.GetValue(_T("expstockid"), tmpStr);
						szSQL.Format(_T("select msl_name as name from m_storagelist where msl_storage_id= %d"), ToInt(tmpStr));
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("name"), tmpStr);
						xls.SetCellText(2, nRow, tmpStr, FMT_TEXT, false, 9);
					}
					else if(szType != _T("P"))
					{
						xls.SetCellText(2, nRow, pMF->GetSelectionString(_T("pms_export_type"), szType), FMT_TEXT, false, 9);
					}	
					else
					{
						rs.GetValue(_T("deptid"), tmpStr);
						szSQL.Format(_T("select pms_name as name from pms_suppliers where pms_id='%s' "), tmpStr);
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("name"), tmpStr);
						xls.SetCellText(2, nRow, tmpStr, FMT_TEXT, false, 9);
					}	
					xls.SetCellText(4, nRow, szUnit, FMT_TEXT | FMT_CENTER, false, 9);
					xls.SetCellText(5, nRow, szPrice, FMT_TEXT | FMT_RIGHT, false, 9);
					rs.GetValue(_T("qty"), tmpStr);
					nAmount = CalcAmt(tmpStr, nPrice);
					xls.SetCellText(6, nRow, tmpStr, FMT_INTEGER, false, 9);
					nOnhand += ToLong(tmpStr);
					nGrpImpQty += ToLong(tmpStr);
					FormatCurrency(nAmount, tmpStr);
					xls.SetCellText(7, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, false, 9);
				}
				else
				{
					xls.SetCellText(0, nRow, tmpStr, FMT_TEXT | FMT_CENTER, false, 9);
					tmpStr.Empty();
					if(szType == _T("B"))
					{
						rs.GetValue(_T("deptid"), tmpStr);
						szSQL.Format(_T("select sd_name as name from sys_dept where sd_id='%s'"), tmpStr);
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("name"), tmpStr);
						xls.SetCellText(2, nRow, tmpStr, FMT_TEXT, false, 9);
					}
					else if(szType == _T("A") || szType == _T("C"))
					{
						rs.GetValue(_T("expstockid"), tmpStr);
						szSQL.Format(_T("select msl_name as name from m_storagelist where msl_storage_id= %d"), ToInt(tmpStr));
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("name"), tmpStr);
						xls.SetCellText(2, nRow, tmpStr, FMT_TEXT, false, 9);
					}
					else if(szType == _T("P") || szType == _T("M"))
					{
						CString szName;
						rs.GetValue(_T("invoiceno"), tmpStr);
						szSQL.Format(_T("select trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname from hms_patient left join hms_doc on(hp_patientno=hd_patientno) where hd_docno=%ld"), ToLong(tmpStr));
						rs2.ExecSQL(szSQL);
						rs2.GetValue(_T("pname"), tmpStr);
						szName.Format(_T("[%s] %s"), rs.GetValue(_T("deptid")), tmpStr);
						xls.SetCellText(2, nRow, tmpStr, FMT_TEXT, false, 9);
					}
					else
					{
						xls.SetCellText(2, nRow, tmpStr, FMT_TEXT, false, 9);
					}
					xls.SetCellText(4, nRow, szUnit, FMT_TEXT | FMT_CENTER, false, 9);
					xls.SetCellText(5, nRow, szPrice, FMT_TEXT | FMT_RIGHT, false, 9);	
					rs.GetValue(_T("qty"), tmpStr);
					nAmount = CalcAmt(tmpStr, nPrice);
					xls.SetCellText(8, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, false, 9);
					nOnhand -= ToLong(tmpStr);
					nGrpExpQty += ToLong(tmpStr);
					FormatCurrency(nAmount, tmpStr);
					xls.SetCellText(9, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, false, 9);
				}
				FormatCurrency(nOnhand, tmpStr);
				xls.SetCellText(12, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, false, 9);
				nAmount = nOnhand*nPrice;
				FormatCurrency(nAmount, tmpStr);
				xls.SetCellText(13, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, false, 9);
				rs.GetValue(_T("lotno"), tmpStr);
				rs.GetValue(_T("expdate"), tmpStr);
				nRow++;
				rs.MoveNext();
			}
			if(nGrpImpQty + nGrpExpQty > 0){
				CString szLabel;
				TranslateString(_T("Sum"), szLabel);
				xls.SetCellText(2, nRow, szLabel, FMT_TEXT | FMT_CENTER, true, 9);
				FormatCurrency(nGrpImpQty, tmpStr);
				xls.SetCellText(6, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
				nAmount = nGrpImpQty*nPrice;
				FormatCurrency(nAmount, tmpStr);
				xls.SetCellText(7, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
				nTtlImpAmt += nAmount;
				FormatCurrency(nGrpExpQty, tmpStr);
				xls.SetCellText(8, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
				nAmount = nGrpExpQty*nPrice;
				FormatCurrency(nAmount, tmpStr);
				xls.SetCellText(9, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
				nTtlExpAmt += nAmount;
				FormatCurrency(nOnhand, tmpStr);
				xls.SetCellText(12, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
				nAmount = nOnhand*nPrice;
				FormatCurrency(nAmount, tmpStr);
				xls.SetCellText(13, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
				nTtlOnhandAmt += nAmount;
				nTtlExpQty += nGrpExpQty;
				nTtlImpQty += nGrpImpQty;
				nTtlOnhandQty += nOnhand;
				nGrpExpQty = nGrpImpQty = nGrpOnhandQty = 0;
				nGrpExpAmt = nGrpExpAmt = nGrpOnhandAmt = 0;
				nRow++;
			}
		}		
	}
	if(nTtlImpQty + nTtlExpQty > 0){
		CString szLabel;
		TranslateString(_T("Total"), szLabel);
		xls.SetCellText(2, nRow, szLabel, FMT_TEXT | FMT_CENTER, true, 9);
		FormatCurrency(nTtlImpQty, tmpStr);
		xls.SetCellText(6, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
		FormatCurrency(nTtlImpAmt, tmpStr);
		xls.SetCellText(7, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
		FormatCurrency(nTtlExpQty, tmpStr);
		xls.SetCellText(8, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
		FormatCurrency(nTtlExpAmt, tmpStr);
		xls.SetCellText(9, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
		FormatCurrency(nTtlOnhandQty, tmpStr);
		xls.SetCellText(12, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);
		FormatCurrency(nTtlOnhandAmt, tmpStr);
		xls.SetCellText(13, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 9);		
	}
	//FormatCurrency(nTotal, tmpStr);
	//xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT , true, 10);
	//xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 10);
	xls.Save(_T("Exports\\So Chi Tiet Vat Tu.xls"));
} 
void CPMSDrugDetailBook::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CPMSDrugDetailBook::OnAddPMSDrugDetailBook(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CPMSDrugDetailBook::OnEditPMSDrugDetailBook(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CPMSDrugDetailBook::OnDeletePMSDrugDetailBook(){
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
 		OnCancelPMSDrugDetailBook(); 
 	}
	return 0;
}
int CPMSDrugDetailBook::OnSavePMSDrugDetailBook(){
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
 		//OnPMSDrugDetailBookListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CPMSDrugDetailBook::OnCancelPMSDrugDetailBook(){
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
int CPMSDrugDetailBook::OnPMSDrugDetailBookListLoadData(){
	return 0;
}
double CPMSDrugDetailBook::CalcAmt(CString szQty, double nPrice){
	double nRet = 0;
	nRet = ToInt(szQty)*nPrice;
	return nRet;
}BEGIN_MESSAGE_MAP(CPMSDrugDetailBook, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSDrugDetailBook::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
