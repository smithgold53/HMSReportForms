#include "stdafx.h"
#include "EMCabinetUsage.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMCabinetUsage *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMCabinetUsage *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMCabinetUsage *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMCabinetUsage *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMCabinetUsage *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMCabinetUsage *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMCabinetUsage *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMCabinetUsage *)pWnd)->OnToDateCheckValue();
} 
static void _OnOrderListSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMCabinetUsage* )pWnd)->OnOrderListSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnOrderListSelendokFnc(CWnd *pWnd){
	((CEMCabinetUsage *)pWnd)->OnOrderListSelendok();
}
/*static void _OnOrderListSetfocusFnc(CWnd *pWnd){
	((CEMCabinetUsage *)pWnd)->OnOrderListKillfocus();
}*/
/*static void _OnOrderListKillfocusFnc(CWnd *pWnd){
	((CEMCabinetUsage *)pWnd)->OnOrderListKillfocus();
}*/
static long _OnOrderListLoadDataFnc(CWnd *pWnd){
	return ((CEMCabinetUsage *)pWnd)->OnOrderListLoadData();
}
/*static void _OnOrderListAddNewFnc(CWnd *pWnd){
	((CEMCabinetUsage *)pWnd)->OnOrderListAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMCabinetUsage *pVw = (CEMCabinetUsage *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMCabinetUsage *pVw = (CEMCabinetUsage *)pWnd;
	pVw->OnExportSelect();
} 
CEMCabinetUsage::CEMCabinetUsage(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CEMCabinetUsage::~CEMCabinetUsage(){
}
void CEMCabinetUsage::OnCreateComponents(){
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 105, 55);
	m_wndFromDate.Create(this,110, 30, 260, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 265, 30, 345, 55);
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 500, 90);
	m_wndToDate.Create(this,350, 30, 495, 55); 
	m_wndOrderListLabel.Create(this, _T("Order List"), 10, 60, 105, 85);
	m_wndOrderList.Create(this,110, 60, 260, 85); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 295, 95, 405, 120);
	m_wndExport.Create(this, _T("&Export"), 410, 95, 500, 120);

}
void CEMCabinetUsage::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndOrderList.SetCheckValue(true);
	m_wndOrderList.LimitText(35);


	m_wndOrderList.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndOrderList.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CEMCabinetUsage::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndOrderList.SetEvent(WE_SELENDOK, _OnOrderListSelendokFnc);
	//m_wndOrderList.SetEvent(WE_SETFOCUS, _OnOrderListSetfocusFnc);
	//m_wndOrderList.SetEvent(WE_KILLFOCUS, _OnOrderListKillfocusFnc);
	m_wndOrderList.SetEvent(WE_SELCHANGE, _OnOrderListSelectChangeFnc);
	m_wndOrderList.SetEvent(WE_LOADDATA, _OnOrderListLoadDataFnc);
	//m_wndOrderList.SetEvent(WE_ADDNEW, _OnOrderListAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	UpdateData(FALSE);

}
void CEMCabinetUsage::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndOrderList.GetDlgCtrlID(), m_szOrderListKey);

}
void CEMCabinetUsage::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szOrderListKey.Empty();

}
int CEMCabinetUsage::SetMode(int nMode){
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
/*void CEMCabinetUsage::OnFromDateChange(){
	
} */
/*void CEMCabinetUsage::OnFromDateSetfocus(){
	
} */
/*void CEMCabinetUsage::OnFromDateKillfocus(){
	
} */
int CEMCabinetUsage::OnFromDateCheckValue(){
	OnOrderListLoadData();
	return 0;
} 
/*void CEMCabinetUsage::OnToDateChange(){
	
} */
/*void CEMCabinetUsage::OnToDateSetfocus(){
	
} */
/*void CEMCabinetUsage::OnToDateKillfocus(){
	
} */
int CEMCabinetUsage::OnToDateCheckValue(){
	OnOrderListLoadData();
	return 0;
} 
void CEMCabinetUsage::OnOrderListSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMCabinetUsage::OnOrderListSelendok(){
	 
}
/*void CEMCabinetUsage::OnOrderListSetfocus(){
	
}*/
/*void CEMCabinetUsage::OnOrderListKillfocus(){
	
}*/
long CEMCabinetUsage::OnOrderListLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndOrderList.IsSearchKey() && !m_szOrderListKey.IsEmpty()){
		szWhere.Format(_T(" and id='%s' "), m_szOrderListKey);
	}
	szWhere.AppendFormat(_T(" AND mt_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), 
						m_szFromDate, m_szToDate);
	m_wndOrderList.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT mt_transaction_id as id, mt_orderno as name FROM m_transaction WHERE mt_doctype = 'CSO' AND mt_department_to_id = 'C1.3' %s ORDER BY id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndOrderList.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMCabinetUsage::OnOrderListAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMCabinetUsage::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	/*Declaration Section*/
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL;
	double nTmp = 0;
	int nIdx = 1, nOrderCount = 0;
	CString szSQL, tmpStr, szReportName, szDate, szOldLev1, szNewLev1;
	long double nTotal = 0, nGroupTotal1 = 0;
	szReportName = _T("Reports/HMS/PM_BANGKETHUOCVTC13.RPT");
	if (!rpt.Init(szReportName))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
		CDate::Convert(m_szToDate), yyyymmdd, ddmmyyyy);
		rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	}
	while (!rs.IsEOF())
	{
		/*Group Seperation*/
		rs.GetValue(_T("order_id"), szNewLev1);
		if (szOldLev1 != szNewLev1)
		{
			if (nGroupTotal1 > 0)
			{
				tmpStr.Format(_T("%f"), nGroupTotal1);
				rptGroup = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptGroup->SetValue(_T("s6"), tmpStr);
				nTotal+= nGroupTotal1;
				nGroupTotal1 = 0;
			}
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));
			if (rptGroup)
			{
				rptGroup->SetValue(_T("OrderID"), szNewLev1);
				tmpStr.Format(_T("[%s] %s - Tu\x1ED5i: %s"), rs.GetValue(_T("doc_no")), rs.GetValue(_T("patient_name")), rs.GetValue(_T("age")));
				rptGroup->SetValue(_T("PName"), tmpStr);
			}
			nIdx = 1;
			nOrderCount++;
			szOldLev1 = szNewLev1;
		}
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("product_name")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("product_uomname")));
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("qty_issue")));
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("unit_price")));
			rs.GetValue(_T("amount"), nTmp);
			nGroupTotal1+= nTmp;
			rptDetail->SetValue(_T("6"), double2str(nTmp));
		}
		rs.MoveNext();
	}
	if (nGroupTotal1 > 0)
	{
		tmpStr.Format(_T("%f"), nGroupTotal1);
		rptGroup = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptGroup->SetValue(_T("s6"), tmpStr);
		nTotal+= nGroupTotal1;
	}
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.GetReportFooter()->SetValue(_T("TotalOrder"), int2str(nOrderCount));
	FormatCurrency(nTotal, tmpStr);
	tmpStr += _T(" \x111\x1ED3ng.");
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	rpt.PrintPreview();

} 
void CEMCabinetUsage::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CExcel xls;
	CString szSQL, tmpStr, szOldLine, szNewLine;
	CStringArray arrCol;
	double nTmp = 0, nGrpCost = 0;
	long double nTtlCost = 0;
	int nIdx = 1, nRow = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(1, 30);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(5, 15);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 5);
	xls.SetCellMergedColumns(0, 3, 5);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr = _T("\x42\x1EA2NG K\xCA THU\x1ED0\x43, V\x1EACT T\x1AF \x43\x31.\x33");
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn thu\x1ED1\x63"));
	arrCol.Add(_T("\x110VT"));
	arrCol.Add(_T("\x110\x1A1n gi\xE1"));
	arrCol.Add(_T("S\x1ED1 l\x1B0\x1EE3ng"));
	arrCol.Add(_T("Th\xE0nh ti\x1EC1n"));
	for (int i = 0; i < arrCol.GetCount(); i++)
		xls.SetCellText(i, 5, arrCol.GetAt(i), 4098, true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("order_id"), szNewLine);
		if (szNewLine != szOldLine)
		{
			if (nGrpCost > 0)
			{
				xls.SetCellMergedColumns(0, nRow, 5);
				xls.SetCellText(0, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
				xls.SetCellText(5, nRow, double2str(nGrpCost), FMT_NUMBER1, true);
				nTtlCost += nGrpCost;
				nGrpCost = 0;
				nRow++;
			}
			xls.SetCellMergedColumns(1, nRow, 5);
			xls.SetCellText(0, nRow, rs.GetValue(_T("order_id")), FMT_TEXT, true);
			tmpStr.Format(_T("[%s] %s - Tu\x1ED5i: %s"), rs.GetValue(_T("doc_no")), rs.GetValue(_T("patient_name")), rs.GetValue(_T("age")));
			xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true);
			nRow++;
			nIdx = 1;
			szOldLine = szNewLine;
		}
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);	
		xls.SetCellText(1, nRow, rs.GetValue(_T("product_name")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("product_uomname")), FMT_TEXT);
		xls.SetCellText(3, nRow, rs.GetValue(_T("unit_price")), FMT_NUMBER1);
		xls.SetCellText(4, nRow, rs.GetValue(_T("qty_issue")), FMT_NUMBER1);
		rs.GetValue(_T("amount"), nTmp);
		nGrpCost += nTmp;
		xls.SetCellText(5, nRow, double2str(nTmp), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nGrpCost > 0)
	{
		xls.SetCellMergedColumns(0, nRow, 5);
		xls.SetCellText(0, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
		xls.SetCellText(5, nRow, double2str(nGrpCost), FMT_NUMBER1, true);
		nTtlCost += nGrpCost;
		nRow++;
	}
	if (nTtlCost > 0)
	{
		xls.SetCellMergedColumns(0, nRow, 5);
		xls.SetCellText(0, nRow, _T("T\x1ED5ng"), 4098, true);
		tmpStr.Format(_T("%f"), nTtlCost);
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\Bang ke thuoc vat tu C1.3.xls"));
	
} 
CString CEMCabinetUsage::GetQueryString(){
	CString szSQL;
	szSQL.Format(_T(" SELECT    hpo_orderid                         order_id, ") \
				_T("           hpo_docno                           doc_no, ") \
				_T("           Get_patientname(hpo_docno)          patient_name, ") \
				_T("           Hms_getagebydoc(hpo_docno)          age, ") \
				_T("		   hpol_line,") \
				_T("           product_id, ") \
				_T("           product_name, ") \
				_T("           product_uomname, ") \
				_T("           hpol_unitprice                      unit_price, ") \
				_T("           SUM(hpol_qtyissue)                  qty_issue, ") \
				_T("           SUM(hpol_qtyissue * hpol_unitprice) amount ") \
				_T(" FROM      hms_pharmaorder ") \
				_T(" LEFT JOIN hms_pharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T(" LEFT JOIN m_productitem_view ON ( product_item_id = hpol_product_item_id ) ") \
				_T(" WHERE     hpo_orderid IN (SELECT DISTINCT hmt_orderid ") \
				_T("                           FROM   hms_medical_transaction_view ") \
				_T("                           WHERE  hmt_reftransaction_id = %ld)") \
				_T(" GROUP     BY hpo_orderid, ") \
				_T("              hpo_docno, hpol_line,") \
				_T("              product_id, ") \
				_T("              product_name, ") \
				_T("              product_uomname, ") \
				_T("              hpol_unitprice") \
				_T(" ORDER BY hpo_orderid, hpo_docno, hpol_line"), str2long(m_szOrderListKey));
	return szSQL;
}