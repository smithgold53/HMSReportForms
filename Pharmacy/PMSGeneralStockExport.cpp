#include "stdafx.h"
#include "PMSGeneralStockExport.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include ".\pmsgeneralstockexport.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSGeneralStockExport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSGeneralStockExport *)pWnd)->OnToDateCheckValue();
} 
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSGeneralStockExport* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMSGeneralStockExport *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnStorageAddNew();
}*/
static void _OnItemGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSGeneralStockExport* )pWnd)->OnItemGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnItemGroupSelendokFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnItemGroupSelendok();
}
/*static void _OnItemGroupSetfocusFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnItemGroupKillfocus();
}*/
/*static void _OnItemGroupKillfocusFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnItemGroupKillfocus();
}*/
static long _OnItemGroupLoadDataFnc(CWnd *pWnd){
	return ((CPMSGeneralStockExport *)pWnd)->OnItemGroupLoadData();
}
/*static void _OnItemGroupAddNewFnc(CWnd *pWnd){
	((CPMSGeneralStockExport *)pWnd)->OnItemGroupAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSGeneralStockExport *pVw = (CPMSGeneralStockExport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSGeneralStockExport *pVw = (CPMSGeneralStockExport *)pWnd;
	pVw->OnExportSelect();
} 
CPMSGeneralStockExport::CPMSGeneralStockExport(CWnd *pParent){

	m_nDlgWidth = 545;
	m_nDlgHeight = 155;
	SetDefaultValues();
}
CPMSGeneralStockExport::~CPMSGeneralStockExport(){
}
void CPMSGeneralStockExport::OnCreateComponents(){
	m_wndGeneralDepartmentUsage.Create(this, _T("Import Sheet List"), 5, 5, 575, 120);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndStorage.SetCheckBox(true);
	m_wndStorageLabel.Create(this, _T("Storage"), 10, 60, 90, 85);
	m_wndStorage.Create(this,95, 60, 570, 85); 
	m_wndItemGroupLabel.Create(this, _T("Item Group"), 10, 90, 90, 115);
	m_wndItemGroup.Create(this,95, 90, 570, 115); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 380, 125, 490, 150);
	m_wndExport.Create(this, _T("&Export"), 495, 125, 575, 150);

}
void CPMSGeneralStockExport::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndStorage.LimitText(35);
	m_wndItemGroup.LimitText(35);


	m_wndStorage.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStorage.InsertColumn(1, _T("Name"), CFMT_TEXT, 380);
	
	m_wndItemGroup.Format(3, 2);
	m_wndItemGroup.InsertColumn(0, _T(""), CFMT_TEXT, 0);
	m_wndItemGroup.InsertColumn(1, _T("Code"), CFMT_TEXT, 60);
	m_wndItemGroup.InsertColumn(2, _T("Name"), CFMT_TEXT, 400);

}
void CPMSGeneralStockExport::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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
	m_wndItemGroup.SetEvent(WE_SELENDOK, _OnItemGroupSelendokFnc);
	//m_wndItemGroup.SetEvent(WE_SETFOCUS, _OnItemGroupSetfocusFnc);
	//m_wndItemGroup.SetEvent(WE_KILLFOCUS, _OnItemGroupKillfocusFnc);
	m_wndItemGroup.SetEvent(WE_SELCHANGE, _OnItemGroupSelectChangeFnc);
	m_wndItemGroup.SetEvent(WE_LOADDATA, _OnItemGroupLoadDataFnc);
	//m_wndItemGroup.SetEvent(WE_ADDNEW, _OnItemGroupAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T(" 00:00");
	m_szToDate += _T(" 23:59");
	UpdateData(false);

}
void CPMSGeneralStockExport::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStorage.GetDlgCtrlID(), m_szStorageKey);
	DDX_TextEx(pDX, m_wndItemGroup.GetDlgCtrlID(), m_szItemGroupKey);

}
void CPMSGeneralStockExport::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStorageKey.Empty();
	m_szItemGroupKey.Empty();

}
int CPMSGeneralStockExport::SetMode(int nMode){
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
/*void CPMSGeneralStockExport::OnFromDateChange(){
	
} */
/*void CPMSGeneralStockExport::OnFromDateSetfocus(){
	
} */
/*void CPMSGeneralStockExport::OnFromDateKillfocus(){
	
} */
int CPMSGeneralStockExport::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSGeneralStockExport::OnToDateChange(){
	
} */
/*void CPMSGeneralStockExport::OnToDateSetfocus(){
	
} */
/*void CPMSGeneralStockExport::OnToDateKillfocus(){
	
} */
int CPMSGeneralStockExport::OnToDateCheckValue(){
	return 0;
} 
void CPMSGeneralStockExport::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSGeneralStockExport::OnStorageSelendok(){
	 
}
/*void CPMSGeneralStockExport::OnStorageSetfocus(){
	
}*/
/*void CPMSGeneralStockExport::OnStorageKillfocus(){
	
}*/
long CPMSGeneralStockExport::OnStorageLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageList(&m_wndStorage, m_szStorageKey);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStorage.IsSearchKey() && !m_szStorageKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szStorageKey
};
	m_wndStorage.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStorage.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CPMSGeneralStockExport::OnStorageAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSGeneralStockExport::OnItemGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSGeneralStockExport::OnItemGroupSelendok(){
	 
}
/*void CPMSGeneralStockExport::OnItemGroupSetfocus(){
	
}*/
/*void CPMSGeneralStockExport::OnItemGroupKillfocus(){
	
}*/
long CPMSGeneralStockExport::OnItemGroupLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadProductCategoryList(&m_wndItemGroup, m_szItemGroupKey);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndItemGroup.IsSearchKey() && !m_szItemGroupKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szItemGroupKey
};
	m_wndItemGroup.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndItemGroup.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CPMSGeneralStockExport::OnItemGroupAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSGeneralStockExport::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CReportSection* rptDetail, *rptHeader, *rptGroup;
	CString szSQL, tmpStr;
	CRecord rs(&pMF->m_db);
	CStringArray arrDat;
	CString szCurDte, szOldCat, szNewCat;
	long double nGrpAmt = 0;
	int nIdx = 1, j = 0;
	if (!rpt.Init(_T("Reports\\HMS\\PM_BAOCAOTONGKETXUATKHO.RPT")))
	{
		return;
	}
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONSTOP | MB_OK);
		return;
	}
	//Header Section
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	tmpStr = _T("To\xE0n \x62\x1ED9");
	if (!m_szStorageKey.IsEmpty())
		rptHeader->SetValue(_T("StockName"), m_wndStorage.GetCurrent(1));
	else
		rptHeader->SetValue(_T("StockName"), tmpStr);
	if (!m_szItemGroupKey.IsEmpty())
		rptHeader->SetValue(_T("Type"), m_wndItemGroup.GetCurrent(1));
	else
		rptHeader->SetValue(_T("Type"), tmpStr);
	tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
	CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	//Detail
	while (!rs.IsEOF())
	{
		/*rs.GetValue(_T("product_categoryname"), szNewCat);
		if (szNewCat != szOldCat && !szNewCat.IsEmpty())
		{
			if (nGrpAmt > 0)
			{
				rptGroup = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptGroup->SetValue(_T("s6"), double2str(nGrpAmt));
			}
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(0));
			rptGroup->SetValue(_T("GroupName"), szNewCat);
			nGrpAmt = 0;
			szOldCat = szNewCat;
			
		}*/
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		//tmpStr.Format(_T("%s %s"), rs.GetValue(_T("pname")), rs.GetValue(_T("pname2")));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("pname")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("uom")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("price")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("qty")));
		rs.GetValue(_T("amt"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		nGrpAmt += str2double(tmpStr);
		rs.MoveNext();
	}
	//Footer
	if (nGrpAmt > 0)
	{
		rptGroup = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptGroup->SetValue(_T("s6"), double2str(nGrpAmt));
	}
	szCurDte = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szCurDte.Right(2), szCurDte.Mid(5, 2), szCurDte.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CPMSGeneralStockExport::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr;
	CString szOldCat, szNewCat;
	CRecord rs(&pMF->m_db);
	int nIdx = 1, j = 0;
	long double nGrpAmt = 0;

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONASTERISK | MB_OK);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	//Header Section
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 30);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 7);
	xls.SetCellText(nCol, nRow + 3, _T("\x42\xC1O \x43\xC1O T\x1ED4NG K\x1EBET \x58U\x1EA4T KHO	"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 7);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	//Col Head
	nRow = 5;
	xls.SetCellMergedRows(nCol, nRow, 2);
	xls.SetCellMergedRows(nCol+1, nRow, 2);
	xls.SetCellMergedRows(nCol+2, nRow, 2);
	xls.SetCellMergedRows(nCol+3, nRow, 2);
	xls.SetCellMergedColumns(nCol+4, nRow, 2);
	xls.SetCellMergedRows(nCol+6, nRow, 2);
	xls.SetCellText(nCol, nRow, _T("STT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+1, nRow, _T("T\xEAn thu\x1ED1\x63, h\xE0m l\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+2, nRow, _T("\x110VT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+3, nRow, _T("\x110\x1A1n gi\xE1"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+4, nRow, _T("T\x1ED5ng s\x1ED1 \x78u\x1EA5t kho"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+4, nRow+1, _T("S\x1ED1 l\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+5, nRow+1, _T("Th\xE0nh ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+6, nRow, _T("Ghi \x63h\xFA"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	nRow = 7;
	while (!rs.IsEOF())
	{
		/*rs.GetValue(_T("product_categoryname"), szNewCat);
		if (szNewCat != szOldCat && !szNewCat.IsEmpty())
		{
			if (nGrpAmt > 0)
			{
				xls.SetCellText(nCol+3, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT);
				xls.SetCellText(nCol+5, nRow, ToString(nGrpAmt), FMT_NUMBER1);
				nRow++;
			}
			xls.SetCellMergedColumns(nCol, nRow, 7);
			xls.SetCellText(nCol, nRow, szNewCat, FMT_TEXT, true);
			nRow++;
			nGrpAmt = 0;
			szOldCat = szNewCat;
		}*/
		xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, rs.GetValue(_T("pname")), FMT_TEXT);
		xls.SetCellText(nCol+2, nRow, rs.GetValue(_T("uom")), FMT_TEXT);
		xls.SetCellText(nCol+3, nRow, rs.GetValue(_T("price")), FMT_NUMBER1);
		xls.SetCellText(nCol+4, nRow, rs.GetValue(_T("qty")), FMT_NUMBER1);
		rs.GetValue(_T("amt"), tmpStr);
		xls.SetCellText(nCol+5, nRow, tmpStr, FMT_NUMBER1);
		nGrpAmt += str2double(tmpStr);
		nRow++;
		rs.MoveNext();
	}
	if (nGrpAmt > 0)
	{
		xls.SetCellText(nCol+3, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT);
		xls.SetCellText(nCol+5, nRow, ToString(nGrpAmt), FMT_NUMBER1);
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\BAOCAOTONGKETXUATKHO.xls"));
} 

CString CPMSGeneralStockExport::GetQueryString(){
	CMainFrame_E10 * pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szWhere, szStock;
	szStock.Empty();
	szWhere.Format(_T(" AND expdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	for (int i = 0; i < m_wndStorage.GetItemCount(); i++)
	{
		if (m_wndStorage.GetCheck(i))
		{
			m_wndStorage.SetCurSel(i);
			if (!szStock.IsEmpty())
				szStock.AppendFormat(_T(", "));
			szStock.AppendFormat(_T("%s"), m_wndStorage.GetCurrent(0));
		}
	}
	if (!szStock.IsEmpty())
		szWhere.AppendFormat(_T(" AND expstockid IN (%s)"), szStock);
	if (!m_szItemGroupKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND product_categoryid = %d"), ToInt(m_szItemGroupKey));
	szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	szSQL.Format(_T(" SELECT    product_name AS pname, ") \
				_T("           product_uomname AS uom, ") \
				_T("           unit_price AS price, ") \
				_T("           SUM(expqty) AS qty, ") \
				_T("           SUM(expqty) * unit_price AS amt ") \
				_T(" FROM      (SELECT sitemid, ") \
				_T("                   expqty, ") \
				_T("                   unit_price, ") \
				_T("				   expdate,") \
				_T("				   expstockid") \
				_T("            FROM   mev ") \
				_T("            WHERE  iotype <> 'MOV'") \
				_T("            UNION ALL ") \
				_T("            SELECT sitemid, ") \
				_T("                   expqty, ") \
				_T("                   unit_price, ") \
				_T("				   expdate,") \
				_T("				   expstockid") \
				_T("            FROM   mev ") \
				_T("            WHERE  iotype = 'MOV' AND impstockid NOT IN ( %s )) tbl ") \
				_T(" LEFT JOIN m_productitem_view ON ( product_item_id = sitemid ) ") \
				_T(" WHERE 1=1 %s") \
				_T(" GROUP     BY product_name,product_uomname,unit_price ") \
				_T(" ORDER     BY pname "), szStock, szWhere);

_fmsg(_T("%s"), szSQL);
	return szSQL;
}
BEGIN_MESSAGE_MAP(CPMSGeneralStockExport, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSGeneralStockExport::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
