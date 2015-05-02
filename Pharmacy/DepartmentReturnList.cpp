#include "stdafx.h"
#include "DepartmentReturnList.h"
#include "MainFrame_E10.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CDepartmentReturnList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CDepartmentReturnList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CDepartmentReturnList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CDepartmentReturnList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CDepartmentReturnList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CDepartmentReturnList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CDepartmentReturnList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CDepartmentReturnList *)pWnd)->OnToDateCheckValue();
} 
static void _OnOutsideListSelectFnc(CWnd *pWnd){
	 ((CDepartmentReturnList*)pWnd)->OnOutsideListSelect();
}
static void _OnExportbyPatSelectFnc(CWnd *pWnd){
	CDepartmentReturnList *pVw = (CDepartmentReturnList *)pWnd;
	pVw->OnExportbyPatSelect();
} 
static void _OnExportbyDrugSelectFnc(CWnd *pWnd){
	CDepartmentReturnList *pVw = (CDepartmentReturnList *)pWnd;
	pVw->OnExportbyDrugSelect();
} 
static long _OnListFacLoadDataFnc(CWnd *pWnd){
	return ((CDepartmentReturnList*)pWnd)->OnListFacLoadData();
} 
static void _OnListFacDblClickFnc(CWnd *pWnd){
	((CDepartmentReturnList*)pWnd)->OnListFacDblClick();
} 
static void _OnListFacSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CDepartmentReturnList*)pWnd)->OnListFacSelectChange(nOldItem, nNewItem);
} 
static int _OnListFacDeleteItemFnc(CWnd *pWnd){
	 return ((CDepartmentReturnList*)pWnd)->OnListFacDeleteItem();
} 
static int _OnListFacCheckAllFnc(CWnd *pWnd){
	return ((CDepartmentReturnList*)pWnd)->OnListFacCheckAll();
}
static int _OnListFacUncheckAllFnc(CWnd *pWnd){
	return ((CDepartmentReturnList*)pWnd)->OnListFacUncheckAll();
}
static long _OnListStockLoadDataFnc(CWnd *pWnd){
	return ((CDepartmentReturnList*)pWnd)->OnListStockLoadData();
} 
static void _OnListStockDblClickFnc(CWnd *pWnd){
	((CDepartmentReturnList*)pWnd)->OnListStockDblClick();
} 
static void _OnListStockSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CDepartmentReturnList*)pWnd)->OnListStockSelectChange(nOldItem, nNewItem);
} 
static int _OnListStockDeleteItemFnc(CWnd *pWnd){
	 return ((CDepartmentReturnList*)pWnd)->OnListStockDeleteItem();
} 
static int _OnListStockCheckAllFnc(CWnd *pWnd){
	return ((CDepartmentReturnList*)pWnd)->OnListStockCheckAll();
}
static int _OnListStockUncheckAllFnc(CWnd *pWnd){
	return ((CDepartmentReturnList*)pWnd)->OnListStockUncheckAll();
}
CDepartmentReturnList::CDepartmentReturnList(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CDepartmentReturnList::~CDepartmentReturnList(){
}
void CDepartmentReturnList::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 550);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndOutsideList.Create(this, _T("Outside List"), 5, 555, 160, 580);
	m_wndExportbyPat.Create(this, _T("Export by Patient"), 350, 555, 460, 580);
	m_wndExportbyDrug.Create(this, _T("Export by Drug"), 465, 555, 575, 580);
	m_wndListFac.Create(this,10, 305, 570, 545); 
	m_wndListStock.Create(this,10, 60, 570, 300); 

}
void CDepartmentReturnList::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);

	m_wndListStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndListStock.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);
	m_wndListStock.SetCheckBox(true);

	m_wndListFac.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndListFac.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndListFac.SetCheckBox(true);

}
void CDepartmentReturnList::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndOutsideList.SetEvent(WE_CLICK, _OnOutsideListSelectFnc);
	m_wndExportbyPat.SetEvent(WE_CLICK, _OnExportbyPatSelectFnc);
	m_wndExportbyDrug.SetEvent(WE_CLICK, _OnExportbyDrugSelectFnc);
	m_wndListFac.SetEvent(WE_SELCHANGE, _OnListFacSelectChangeFnc);
	m_wndListFac.SetEvent(WE_LOADDATA, _OnListFacLoadDataFnc);
	m_wndListFac.SetEvent(WE_DBLCLICK, _OnListFacDblClickFnc);
	m_wndListFac.AddEvent(1, _T("Check All"), _OnListFacCheckAllFnc);
	m_wndListFac.AddEvent(2, _T("Uncheck All"), _OnListFacUncheckAllFnc);
	m_wndListStock.SetEvent(WE_SELCHANGE, _OnListStockSelectChangeFnc);
	m_wndListStock.SetEvent(WE_LOADDATA, _OnListStockLoadDataFnc);
	m_wndListStock.SetEvent(WE_DBLCLICK, _OnListStockDblClickFnc);
	m_wndListStock.AddEvent(1, _T("Check All"), _OnListStockCheckAllFnc);
	m_wndListStock.AddEvent(2, _T("Uncheck All"), _OnListStockUncheckAllFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	OnListFacLoadData();
	OnListStockLoadData();
	UpdateData(FALSE);

}
void CDepartmentReturnList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_Check(pDX, m_wndOutsideList.GetDlgCtrlID(), m_bOutsideList);

}
void CDepartmentReturnList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_bOutsideList=FALSE;

}
int CDepartmentReturnList::SetMode(int nMode){
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
/*void CDepartmentReturnList::OnFromDateChange(){
	
} */
/*void CDepartmentReturnList::OnFromDateSetfocus(){
	
} */
/*void CDepartmentReturnList::OnFromDateKillfocus(){
	
} */
int CDepartmentReturnList::OnFromDateCheckValue(){
	return 0;
} 
/*void CDepartmentReturnList::OnToDateChange(){
	
} */
/*void CDepartmentReturnList::OnToDateSetfocus(){
	
} */
/*void CDepartmentReturnList::OnToDateKillfocus(){
	
} */
int CDepartmentReturnList::OnToDateCheckValue(){
	return 0;
} 
void CDepartmentReturnList::OnOutsideListSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
void CDepartmentReturnList::OnExportbyPatSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere, StockStr, StockNameStr, FacStr, FacNameStr, szOldDept, szNewDept, tmpStr, szOldDoc, szNewDoc;
	CStringArray arrCol;
	int nCol = 0, nRow = 0, nIdx = 1, nOldRow = 0;
	double nTemp = 0;
	long double nDeptAmount = 0, nTtlAmount = 0;
	double nDeptQty = 0, nTtlQty = 0; 
	if (!m_bOutsideList)
		szWhere.Format(_T(" AND mt_createddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	for (int i = 0; i < m_wndListStock.GetItemCount(); i++)
	{
		if (m_wndListStock.GetCheck(i))
		{
			if (!StockStr.IsEmpty())
			{
				StockStr += _T(", ");
				StockNameStr += _T(", ");
			}
			StockNameStr.AppendFormat(_T("%s"), m_wndListStock.GetItemText(i, 1));
			StockStr.AppendFormat(_T("%s"), m_wndListStock.GetItemText(i, 0));
		}
	}
	for (int i = 0; i < m_wndListFac.GetItemCount(); i++)
	{
		if (m_wndListFac.GetCheck(i))
		{
			if (!FacStr.IsEmpty())
			{
				FacStr += _T(", ");
				FacNameStr += _T(", ");
			}
			FacNameStr.AppendFormat(_T("%s"), m_wndListFac.GetItemText(i, 1));
			FacStr.AppendFormat(_T("'%s'"), m_wndListFac.GetItemText(i, 0));
		}
	}
	if (!StockStr.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_storage_id IN (%s)"), StockStr);
	if (!FacStr.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_deptid IN (%s)"), FacStr);
	//if (m_bOutsideList)
	//	szWhere.AppendFormat(_T(" AND hpol_qtyreverse > 0"));
	//else
	//	szWhere.AppendFormat(_T(" AND hpol_processed <> 'P' AND hpol_qtyreverse > 0"));
	if (m_bOutsideList)
		szWhere.AppendFormat(_T(" AND (hpol_processed <> 'P' AND hpol_qtyreverse > 0)"));
	else
		szWhere.AppendFormat(_T(" AND mt_status = 'A'"));
	szSQL.Format(_T(" SELECT    Get_patientname(hpo_docno)     AS patname, ") \
				_T("           hpo_docno                      AS docno, ") \
				_T("           product_name                   AS prodname, ") \
				_T("           sum(hpol_qtyreverse)                AS qty, ") \
				_T("		   hpol_unitprice price,") \
				_T("		   sum(hpol_qtyreverse)*hpol_unitprice amount,") \
				_T("           Get_departmentname(hpo_deptid) AS dept ") \
				_T(" FROM      hms_ipharmaorder ") \
				_T(" LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T(" LEFT JOIN m_transaction ON ( mt_transaction_id = hpol_retorder_id ) ") \
				_T(" LEFT JOIN m_product_view ON ( product_id = hpol_product_id ) ") \
				_T(" WHERE 1=1 %s") \
				_T(" GROUP BY hpo_docno, product_name, hpo_deptid, hpol_unitprice") \
				_T(" ORDER BY hpo_deptid, hpo_docno, product_name"), szWhere);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."), MB_ICONASTERISK);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 40);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(4, 15);
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);
	xls.SetCellMergedColumns(nCol, nRow + 3, 3);
	xls.SetCellText(nCol, nRow, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol, nRow + 2, _T("\x44\x41NH S\xC1\x43H THU\x1ED0\x43 TR\x1EA2 L\x1EA0I T\x1EEA KHO\x41"), FMT_TEXT | FMT_CENTER, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), m_szFromDate, m_szToDate);
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER);
	nRow = 3;
	if (!StockNameStr.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("T\x1EEB kho: %s"), StockNameStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
	}
	if (!FacNameStr.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("T\x1EEB khoa: %s"), FacNameStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
	}
	nRow++;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn thu\x1ED1\x63"));
	arrCol.Add(_T("S\x1ED1 l\x1B0\x1EE3ng"));
	arrCol.Add(_T("\x110\x1A1n gi\xE1"));
	arrCol.Add(_T("Th\xE0nh ti\x1EC1n"));
	nRow++;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol + i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true);
	}
	nRow++;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("dept"), szNewDept);
		rs.GetValue(_T("docno"), szNewDoc);
		if (szNewDept != szOldDept)
		{
			if (nDeptQty > 0)
			{
				xls.SetCellText(2, nOldRow, double2str(nDeptQty), FMT_NUMBER1, true);
				tmpStr.Format(_T("%f"), nDeptAmount);
				xls.SetCellText(4, nOldRow, tmpStr, FMT_NUMBER1, true);
				nTtlAmount += nDeptAmount;
				nTtlQty += nDeptQty;
				nDeptAmount = 0;
				nDeptQty = 0;
			}
			xls.SetCellText(nCol, nRow, szNewDept, FMT_TEXT, true);
			szOldDept = szNewDept;
			nOldRow = nRow;
			nRow++;
		}
		if (szNewDoc != szOldDoc)
		{
			xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
			tmpStr.Format(_T("[%s] %s"), rs.GetValue(_T("docno")), rs.GetValue(_T("patname")));
			xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT);
			szOldDoc = szNewDoc;
			nRow++;
		}
		
		xls.SetCellText(nCol + 1, nRow, _T("- ")+rs.GetValue(_T("prodname")), FMT_TEXT); 
		rs.GetValue(_T("qty"), nTemp);
		nDeptQty += nTemp;
		xls.SetCellText(nCol + 2, nRow, double2str(nTemp), FMT_NUMBER1);
		xls.SetCellText(3, nRow, rs.GetValue(_T("price")), FMT_NUMBER1);
		rs.GetValue(_T("amount"), nTemp);
		nDeptAmount += nTemp;
		xls.SetCellText(4, nRow, double2str(nTemp), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nDeptQty > 0)
	{
		xls.SetCellText(2, nOldRow, double2str(nDeptQty), FMT_NUMBER1, true);
		tmpStr.Format(_T("%f"), nDeptAmount);
		xls.SetCellText(4, nOldRow, tmpStr, FMT_NUMBER1, true);
		nTtlAmount += nDeptAmount;
		nTtlQty += nDeptQty;
	}
	if (nTtlQty > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
		xls.SetCellText(2, nRow, double2str(nTtlQty), FMT_NUMBER1, true);
		tmpStr.Format(_T("%f"), nTtlAmount);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
		
	}
	xls.Save(_T("Exports\\Danh sach thuoc tra lai tu khoa theo BN.xls"));
} 
void CDepartmentReturnList::OnExportbyDrugSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere, StockStr, StockNameStr, FacStr, FacNameStr, szOldDept, szNewDept, tmpStr;
	CStringArray arrCol;
	int nCol = 0, nRow = 0, nIdx = 1;
	int nOldRow = 0;
	double nGrpQty = 0, nTtlQty = 0;
	double nCost = 0;
	long double nGrpAmount = 0, nTtlAmount = 0;
	if (!m_bOutsideList)
		szWhere.Format(_T(" AND mt_createddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	for (int i = 0; i < m_wndListStock.GetItemCount(); i++)
	{
		if (m_wndListStock.GetCheck(i))
		{
			if (!StockStr.IsEmpty())
			{
				StockStr += _T(", ");
				StockNameStr += _T(", ");
			}
			StockNameStr.AppendFormat(_T("%s"), m_wndListStock.GetItemText(i, 1));
			StockStr.AppendFormat(_T("%s"), m_wndListStock.GetItemText(i, 0));
		}
	}
	for (int i = 0; i < m_wndListFac.GetItemCount(); i++)
	{
		if (m_wndListFac.GetCheck(i))
		{
			if (!FacStr.IsEmpty())
			{
				FacStr += _T(", ");
				FacNameStr += _T(", ");
			}
			FacNameStr.AppendFormat(_T("%s"), m_wndListFac.GetItemText(i, 1));
			FacStr.AppendFormat(_T("'%s'"), m_wndListFac.GetItemText(i, 0));
		}
	}
	if (!StockStr.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_storage_id IN (%s)"), StockStr);
	if (!FacStr.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_deptid IN (%s)"), FacStr);
	if (m_bOutsideList)
		szWhere.AppendFormat(_T(" AND (hpol_processed <> 'P' AND hpol_qtyreverse > 0)"));
	else 
		szWhere.AppendFormat(_T(" AND mt_status = 'A'"));
	szSQL.Format(_T(" SELECT   product_name                   AS prodname, ") \
				_T("           sum(hpol_qtyreverse)                AS qty, ") \
				_T("		   hpol_unitprice price,") \
				_T("		   sum(hpol_qtyreverse)*hpol_unitprice amount,") \
				_T("           Get_departmentname(hpo_deptid) AS dept ") \
				_T(" FROM hms_ipharmaorder ") \
				_T(" LEFT JOIN hms_ipharmaorderline ON (hpo_orderid = hpol_orderid)") \
				_T(" LEFT JOIN m_transaction ON ( mt_transaction_id = hpol_retorder_id ) ") \
				_T(" LEFT JOIN m_product_view ON ( product_id = hpol_product_id ) ") \
				_T(" WHERE 1=1 %s") \
				_T(" GROUP BY product_name, hpo_deptid, hpol_unitprice") \
				_T(" ORDER BY hpo_deptid, product_name"), szWhere);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."), MB_ICONASTERISK);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 40);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(4, 15);
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);
	xls.SetCellMergedColumns(nCol, nRow + 3, 3);
	xls.SetCellText(nCol, nRow, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol, nRow + 2, _T("\x44\x41NH S\xC1\x43H THU\x1ED0\x43 TR\x1EA2 L\x1EA0I T\x1EEA KHO\x41"), FMT_TEXT | FMT_CENTER, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), m_szFromDate, m_szToDate);
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER);
	nRow = 3;
	if (!StockNameStr.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("T\x1EEB kho: %s"), StockNameStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
	}
	if (!FacNameStr.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("T\x1EEB khoa: %s"), FacNameStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
	}
	nRow++;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn thu\x1ED1\x63"));
	arrCol.Add(_T("S\x1ED1 l\x1B0\x1EE3ng"));
	arrCol.Add(_T("\x110\x1A1n gi\xE1"));
	arrCol.Add(_T("Th\xE0nh ti\x1EC1n"));
	nRow++;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol + i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true);
	}
	nRow++;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("dept"), szNewDept);
		if (szNewDept != szOldDept)
		{
			if (nGrpQty > 0)
			{
				xls.SetCellText(nCol + 2, nOldRow, double2str(nGrpQty), FMT_NUMBER1, true); 
				tmpStr.Format(_T("%f"), nGrpAmount);
				xls.SetCellText(nCol + 4, nOldRow, tmpStr, FMT_NUMBER1, true);
				nTtlAmount += nGrpAmount;
				nTtlQty += nGrpQty;
				nGrpQty = 0;
				nGrpAmount = 0;
			}
			xls.SetCellText(nCol, nRow, szNewDept, FMT_TEXT, true);
			szOldDept = szNewDept;
			nOldRow = nRow;
			nIdx = 1;
			nRow++;
		}		
		xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(nCol + 1, nRow, _T("- ")+rs.GetValue(_T("prodname")), FMT_TEXT); 
		nGrpQty += str2double(rs.GetValue(_T("qty")));
		xls.SetCellText(nCol + 2, nRow, rs.GetValue(_T("qty")), FMT_NUMBER1);
		xls.SetCellText(nCol + 3, nRow, rs.GetValue(_T("price")), FMT_NUMBER1);
		rs.GetValue(_T("amount"), nCost);
		nGrpAmount += nCost;
		xls.SetCellText(nCol + 4, nRow, double2str(nCost), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nGrpQty > 0)
	{
		xls.SetCellText(nCol + 2, nOldRow, double2str(nGrpQty), FMT_NUMBER1, true);
		tmpStr.Format(_T("%f"), nGrpAmount);
		xls.SetCellText(nCol + 4, nOldRow, tmpStr, FMT_NUMBER1, true); 
		nTtlQty += nGrpQty;
		nTtlAmount += nGrpAmount;
	}
	if (nTtlQty > 0)
	{
		xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
		xls.SetCellText(nCol + 2, nRow, double2str(nTtlQty), FMT_NUMBER1, true); 
		tmpStr.Format(_T("%f"), nTtlAmount);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1, true); 
	}
	xls.Save(_T("Exports\\Danh sach thuoc tra lai tu khoa.xls"));
	
} 
void CDepartmentReturnList::OnListFacDblClick(){
	
} 
void CDepartmentReturnList::OnListFacSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CDepartmentReturnList::OnListFacDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
} 
int CDepartmentReturnList::OnListFacCheckAll(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	for (int i=0; i< m_wndListFac.GetItemCount(); i++){
		m_wndListFac.SetCheck(i, true);
	}
	return 0;	
}

int CDepartmentReturnList::OnListFacUncheckAll(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	for (int i=0; i< m_wndListFac.GetItemCount(); i++){
		m_wndListFac.SetCheck(i, false);
	}
	return 0;	
}

long CDepartmentReturnList::OnListFacLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndListFac.BeginLoad(); 
	m_wndListFac.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" select sd_id as id, sd_name as name from sys_dept where sd_type = 'DT' order by  sd_index,sd_type "));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndListFac.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndListFac.EndLoad(); 
	return nCount;
}
void CDepartmentReturnList::OnListStockDblClick(){
	
} 
void CDepartmentReturnList::OnListStockSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CDepartmentReturnList::OnListStockDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
} 
long CDepartmentReturnList::OnListStockLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageIntoListCtrl(&m_wndListStock);
/*
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndListStock.BeginLoad(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndListStock.AddItems(
		rs.MoveNext();
	}
	m_wndListStock.EndLoad(); 
	return nCount;
*/
	return 0;
}

int CDepartmentReturnList::OnListStockCheckAll(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	for (int i=0; i< m_wndListStock.GetItemCount(); i++){
		m_wndListStock.SetCheck(i, true);
	}
	 return 0;	
}

int CDepartmentReturnList::OnListStockUncheckAll(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	for (int i=0; i< m_wndListStock.GetItemCount(); i++){
		m_wndListStock.SetCheck(i, false);
	}
	 return 0;	
}