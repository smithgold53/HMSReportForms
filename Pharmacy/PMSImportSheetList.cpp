#include "stdafx.h"
#include "PMSImportSheetList.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include ".\pmsimportsheetlist.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSImportSheetList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSImportSheetList *)pWnd)->OnToDateCheckValue();
} 
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSImportSheetList* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMSImportSheetList *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnStorageAddNew();
}*/
/*static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CPMSImportSheetList*)pWnd)->OnTypeLoadData();
}*/
static long _OnSupplierLoadDataFnc(CWnd *pWnd){
	return ((CPMSImportSheetList *)pWnd)->OnSupplierLoadData();
}
static void _OnItemGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSImportSheetList* )pWnd)->OnItemGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnItemGroupSelendokFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnItemGroupSelendok();
}
/*static void _OnItemGroupSetfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnItemGroupKillfocus();
}*/
/*static void _OnItemGroupKillfocusFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnItemGroupKillfocus();
}*/
static long _OnItemGroupLoadDataFnc(CWnd *pWnd){
	return ((CPMSImportSheetList *)pWnd)->OnItemGroupLoadData();
}
/*static void _OnItemGroupAddNewFnc(CWnd *pWnd){
	((CPMSImportSheetList *)pWnd)->OnItemGroupAddNew();
}*/
static void _OnNonInvoiceImportSelectFnc(CWnd *pWnd){
	((CPMSImportSheetList*)pWnd)->OnNonInvoiceImportSelect();
}
static void _OnCompanyImportSelectFnc(CWnd *pWnd){
	((CPMSImportSheetList*)pWnd)->OnCompanyImportSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSImportSheetList *pVw = (CPMSImportSheetList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSImportSheetList *pVw = (CPMSImportSheetList *)pWnd;
	pVw->OnExportSelect();
} 
CPMSImportSheetList::CPMSImportSheetList(CWnd *pParent){

	m_nDlgWidth = 545;
	m_nDlgHeight = 155;
	SetDefaultValues();
}
CPMSImportSheetList::~CPMSImportSheetList(){
}
void CPMSImportSheetList::OnCreateComponents(){
	m_wndGeneralDepartmentUsage.Create(this, _T("Import Sheet List"), 5, 5, 575, 120);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 285, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 290, 30, 370, 55);
	m_wndToDate.Create(this,375, 30, 570, 55); 
	m_wndStorage.SetCheckBox(true);
	m_wndStorageLabel.Create(this, _T("Storage"), 10, 60, 90, 85);
	m_wndStorage.Create(this,95, 60, 285, 85); 
	//m_wndTypeLabel.Create(this, _T("Type"), 290, 60, 370, 85);
	//m_wndType.Create(this,375, 60, 570, 85);
	m_wndItemGroupLabel.Create(this, _T("Item Group"), 290, 60, 370, 85);
	m_wndItemGroup.Create(this,375, 60, 570, 85);
	m_wndSupplierLabel.Create(this, _T("Supplier"), 10, 90, 90, 115);
	m_wndSupplier.Create(this,95, 90, 570, 115);
	m_wndNonInvoiceImport.Create(this, _T("NonInvoice Import"), 5, 125, 155, 150);
	m_wndCompanyImport.Create(this, _T("Nh\x1EADp t\x1EEB kho \x37\x30\x38"), 160, 125, 310, 150);
	m_wndGroupByType.Create(this, _T("Group by Type"), 315, 125, 425, 150);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 380, 155, 490, 180);
	m_wndExport.Create(this, _T("&Export"), 495, 155, 575, 180);
}
void CPMSImportSheetList::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndStorage.LimitText(35);
	m_wndItemGroup.LimitText(35);

	m_wndStorage.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStorage.InsertColumn(1, _T("Name"), CFMT_TEXT, 170);
	
// 	m_wndType.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
// 	m_wndType.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);

	m_wndItemGroup.InsertColumn(1, _T("ID"), CFMT_TEXT, 50);
	m_wndItemGroup.InsertColumn(2, _T("Name"), CFMT_TEXT, 250);

	m_wndSupplier.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndSupplier.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);
}
void CPMSImportSheetList::OnSetWindowEvents(){
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
	//m_wndType.SetEvent(WE_LOADDATA, _OnTypeLoadDataFnc);
	m_wndItemGroup.SetEvent(WE_SELENDOK, _OnItemGroupSelendokFnc);
	//m_wndItemGroup.SetEvent(WE_SETFOCUS, _OnItemGroupSetfocusFnc);
	//m_wndItemGroup.SetEvent(WE_KILLFOCUS, _OnItemGroupKillfocusFnc);
	m_wndItemGroup.SetEvent(WE_SELCHANGE, _OnItemGroupSelectChangeFnc);
	m_wndItemGroup.SetEvent(WE_LOADDATA, _OnItemGroupLoadDataFnc);
	//m_wndItemGroup.SetEvent(WE_ADDNEW, _OnItemGroupAddNewFnc);
	m_wndSupplier.SetEvent(WE_LOADDATA, _OnSupplierLoadDataFnc);
	m_wndNonInvoiceImport.SetEvent(WE_CLICK, _OnNonInvoiceImportSelectFnc);
	m_wndCompanyImport.SetEvent(WE_CLICK, _OnCompanyImportSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T(" 00:00");
	m_szToDate += _T(" 23:59");
	UpdateData(false);

}
void CPMSImportSheetList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStorage.GetDlgCtrlID(), m_szStorageKey);
	//DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_TextEx(pDX, m_wndItemGroup.GetDlgCtrlID(), m_szItemGroupKey);
	DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szSupplierKey);
	DDX_Check(pDX, m_wndNonInvoiceImport.GetDlgCtrlID(), m_bNonInvoice);
	DDX_Check(pDX, m_wndCompanyImport.GetDlgCtrlID(), m_bCompany);
	DDX_Check(pDX, m_wndGroupByType.GetDlgCtrlID(), m_bGroupbyType);
}
void CPMSImportSheetList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStorageKey.Empty();
	//m_szTypeKey.Empty();
	m_szItemGroupKey.Empty();
	m_szSupplierKey.Empty();
	m_bNonInvoice = FALSE;
	m_bCompany = FALSE;
	m_bGroupbyType = FALSE;

}
int CPMSImportSheetList::SetMode(int nMode){
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
/*void CPMSImportSheetList::OnFromDateChange(){
	
} */
/*void CPMSImportSheetList::OnFromDateSetfocus(){
	
} */
/*void CPMSImportSheetList::OnFromDateKillfocus(){
	
} */
int CPMSImportSheetList::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSImportSheetList::OnToDateChange(){
	
} */
/*void CPMSImportSheetList::OnToDateSetfocus(){
	
} */
/*void CPMSImportSheetList::OnToDateKillfocus(){
	
} */
int CPMSImportSheetList::OnToDateCheckValue(){
	return 0;
} 
void CPMSImportSheetList::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSImportSheetList::OnStorageSelendok(){
	 
}
/*void CPMSImportSheetList::OnStorageSetfocus(){
	
}*/
/*void CPMSImportSheetList::OnStorageKillfocus(){
	
}*/
long CPMSImportSheetList::OnStorageLoadData(){
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
void CPMSImportSheetList::OnItemGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSImportSheetList::OnItemGroupSelendok(){
	 
}
/*void CPMSImportSheetList::OnItemGroupSetfocus(){
	
}*/
/*void CPMSImportSheetList::OnItemGroupKillfocus(){
	
}*/
long CPMSImportSheetList::OnItemGroupLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	int nCount = 0;
	szSQL.Format(_T(" SELECT DISTINCT product_type_id, ") \
		_T("                 product_type_name ") \
		_T(" FROM   (SELECT CASE WHEN mpt_product_type_id IN ( 'A1100', 'A1130', 'A1140', 'A1160', ") \
		_T("                                              'A1170', 'A1180', 'A1190', 'A1400', 'A6000' ) THEN N'A0000' ") \
		_T("                ELSE mpt_product_type_id ") \
		_T("                END product_type_id, ") \
		_T("                CASE WHEN mpt_product_type_id IN ( 'A1100', 'A1130', 'A1140', 'A1160', ") \
		_T("                                              'A1170', 'A1180', 'A1190', 'A1400', 'A6000' ) THEN N'Thu\x1ED1\x63' ") \
		_T("                ELSE mpt_name ") \
		_T("                END product_type_name ") \
		_T("         FROM   m_product_type ") \
		_T("         WHERE  mpt_org_id = '%s' ") \
		_T("            AND mpt_isactive = 'Y' AND mpt_product_type_id <> 'A1000') ") \
		_T(" ORDER BY product_type_id"), pMF->GetModuleID());
	nCount = rs.ExecSQL(szSQL);
	m_wndItemGroup.DeleteAllItems();
	while (!rs.IsEOF())
	{
		m_wndItemGroup.AddItems(
			rs.GetValue(_T("product_type_id")),
			rs.GetValue(_T("product_type_name")),
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPMSImportSheetList::OnItemGroupAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
long CPMSImportSheetList::OnSupplierLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadPartnerList(&m_wndSupplier, m_szSupplierKey);
}
void CPMSImportSheetList::OnNonInvoiceImportSelect(){
	m_wndCompanyImport.SetCheck(false);
}
void CPMSImportSheetList::OnCompanyImportSelect(){
	m_wndNonInvoiceImport.SetCheck(false);
}
void CPMSImportSheetList::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CReportSection* rptDetail = NULL, *rptHeader = NULL, *rptGroup = NULL, *rptFooter = NULL;
	CString szSQL, tmpStr, szField, szTemp;
	CRecord rs(&pMF->m_db);
	CStringArray arrDat;
	CString szCurDte, szOldGroup, szNewGroup, szStorages;
	double nTmp = 0;
	long double nGrpAmt = 0, nTtlAmt = 0;
	int nIdx = 1, j = 0;
	if (!rpt.Init(_T("Reports\\HMS\\PM_BANGKEPHIEUNHAP.RPT")))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONSTOP | MB_OK);
		return;
	}
	//Header Section
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	rptHeader->SetValue(_T("DeptName"), pMF->GetDepartmentName(pMF->GetDepartmentID()));
	tmpStr = _T("To\xE0n \x62\x1ED9");
	for (int i = 0; i < m_wndStorage.GetItemCount(); i++)
	{
		if (m_wndStorage.GetCheck(i))
		{
			m_wndStorage.SetCurSel(i);
			if (!szStorages.IsEmpty())
				szStorages += _T(", ");
			szStorages += m_wndStorage.GetCurrent(1);
		}
	}
	if (!szStorages.IsEmpty())
		rptHeader->SetValue(_T("StockName"), szStorages);
	else
		rptHeader->SetValue(_T("StockName"), tmpStr);
	if (!m_szItemGroupKey.IsEmpty())
		rptHeader->SetValue(_T("Type"), m_wndItemGroup.GetCurrent(1));
	else
		rptHeader->SetValue(_T("Type"), tmpStr);
// 	if (!m_szTypeKey.IsEmpty())
// 		rptHeader->SetValue(_T("Type"), m_wndType.GetCurrent(1));
// 	else
// 		rptHeader->SetValue(_T("Type"), tmpStr);
	tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
	CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	//Detail
	while (!rs.IsEOF())
	{
		if (m_bGroupbyType)
			tmpStr = _T("product_groupid");
		else
			tmpStr = _T("dept");
		rs.GetValue(tmpStr, szNewGroup);
		if (szNewGroup != szOldGroup && !szNewGroup.IsEmpty())
		{
			if (nGrpAmt > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				tmpStr.Format(rptDetail->GetValue(_T("TotalGroup")), szOldGroup);
				rptDetail->SetValue(_T("TotalGroup"), tmpStr);
				rptDetail->SetValue(_T("s5"), ToString(nGrpAmt));
				nTtlAmt += nGrpAmt;
				nGrpAmt = 0;
			}
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));
			if (rptGroup)
			{
				if (m_bGroupbyType)
					tmpStr = rs.GetValue(_T("product_groupname"));
				else
					tmpStr = szNewGroup;
				rptGroup->SetValue(_T("GroupName"), tmpStr);
			}
				
			szOldGroup = szNewGroup;
		}
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("ordno")));
		rs.GetValue(_T("impdte"), tmpStr);
		rptDetail->SetValue(_T("3"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("dept")));
		rs.GetValue(_T("ttlamt"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		nGrpAmt += str2double(tmpStr);
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("invno")));
		rs.MoveNext();
	}
	if (nGrpAmt > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		tmpStr.Format(rptDetail->GetValue(_T("TotalGroup")), szOldGroup);
		rptDetail->SetValue(_T("TotalGroup"), tmpStr);
		tmpStr.Format(_T("%f"), nGrpAmt);
		rptDetail->SetValue(_T("s5"), tmpStr);
		nTtlAmt += nGrpAmt;
		nGrpAmt = 0;
	}
	if (nTtlAmt > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		tmpStr.Format(rptDetail->GetValue(_T("TotalGroup")), szOldGroup);
		rptDetail->SetValue(_T("TotalGroup"), tmpStr);
		tmpStr.Format(_T("%f"), nTtlAmt);
		rptDetail->SetValue(_T("s5"), tmpStr);
	}
	//Footer
	rptFooter = rpt.GetReportFooter();
	if(!m_bNonInvoice)
	{
		szSQL = GetQueryString1();
		rs.ExecSQL(szSQL);
		if (!rs.IsEOF())
		{
			//l1..l4
			int nIdx = 1;
			for (int i = 0; i < rs.GetFieldCount(); i++)
			{
				szTemp = rs.GetFieldName(i);
				szField.Format(_T("%s%s"), szTemp.Left(1), szTemp.Right(szTemp.GetLength()-1).MakeLower());
				rs.GetValue(szField, nTmp);
				if (nTmp > 0)
				{
					tmpStr.Format(_T("l%d"), nIdx);
					TranslateString(szField, szTemp);
					rptFooter->SetValue(tmpStr, szTemp);
					tmpStr.Format(_T("s%d"), nIdx);
					FormatCurrency(nTmp, szTemp);
					szTemp += _T(" \x111\x1ED3ng.");
					rptFooter->SetValue(tmpStr, szTemp);
					nIdx++;
				}
			}
		}
	}
// 	if(!rs.IsEOF())
// 	{	
// 		for(int i = 0; i < 4; i++)
// 		{
// 			tmpStr.Format(_T("s%d"), i);
// 			rs.GetValue(tmpStr, nTmp);
// 			if(nTmp > 0)
// 			{
// 				FormatCurrency(nTmp, szTemp);
// 				szTemp += _T(" \x111\x1ED3ng.");
// 				rptFooter->SetValue(tmpStr, szTemp);
// 			}
// 			else
// 			{
// 				szTemp = _T("0 \x111\x1ED3ng.");
// 				rptFooter->SetValue(tmpStr, szTemp);
// 			}
// 		}
// 	}
	szCurDte = pMF->GetSysDate();
	tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), szCurDte.Right(2), szCurDte.Mid(5, 2), szCurDte.Left(4));
	rptFooter->SetValue(_T("PrintDate"), tmpStr);
	rptFooter->SetValue(_T("Username"), pMF->GetCurrentUser());
	rpt.PrintPreview();
} 
void CPMSImportSheetList::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr;
	CString szOldGroup, szNewGroup;
	CRecord rs(&pMF->m_db);
	int nIdx = 1, j = 0;
	long double nGrpAmt = 0, nTtlAmt = 0;

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
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 40);
	xls.SetColumnWidth(4, 30);
	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 7);
	xls.SetCellText(nCol, nRow + 3, _T("\x42\x1EA2NG K\xCA PHI\x1EBEU NH\x1EACP"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 7);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	//Col Head
	nRow = 5;
	xls.SetCellMergedRows(nCol, nRow, 2);
	xls.SetCellMergedColumns(nCol+1, nRow, 2);
	xls.SetCellMergedRows(nCol+3, nRow, 2);
	xls.SetCellMergedRows(nCol+4, nRow, 2);
	xls.SetCellMergedRows(nCol+5, nRow, 2);
	xls.SetCellMergedRows(nCol+6, nRow, 2);
	xls.SetCellText(nCol, nRow, _T("STT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+1, nRow, _T("Phi\x1EBFu nh\x1EADp"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+1, nRow+1, _T("S\x1ED1 \x63h\x1EE9ng t\x1EEB"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+2, nRow+1, _T("Ng\xE0y"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+3, nRow, _T("Ngu\x1ED3n h\xE0ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+4, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+5, nRow, _T("N\x1EE3"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+6, nRow, _T("S\x1ED1 h\xF3\x61 \x111\x1A1n"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	nRow = 7;
	while (!rs.IsEOF())
	{
		if (m_bGroupbyType)
			tmpStr = _T("product_groupid");
		else
			tmpStr = _T("dept");
		rs.GetValue(tmpStr, szNewGroup);
		if (szNewGroup != szOldGroup && !szNewGroup.IsEmpty())
		{
			if (nGrpAmt > 0)
			{
				xls.SetCellText(nCol+3, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT, true);
				xls.SetCellText(nCol+4, nRow, ToString(nGrpAmt), FMT_NUMBER1, true);
				nTtlAmt += nGrpAmt;
				nGrpAmt = 0;
				nRow++;
			}
			xls.SetCellMergedColumns(0, nRow, 7);
			if (m_bGroupbyType)
				rs.GetValue(_T("product_groupname"), tmpStr);
			else
				tmpStr = szNewGroup;
			xls.SetCellText(0, nRow, tmpStr, FMT_TEXT, true);
			szOldGroup = szNewGroup;
		}
		xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, rs.GetValue(_T("ordno")), FMT_TEXT);
		rs.GetValue(_T("impdte"), tmpStr);
		xls.SetCellText(nCol+2, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		xls.SetCellText(nCol+3, nRow, rs.GetValue(_T("dept")), FMT_TEXT);
		rs.GetValue(_T("ttlamt"), tmpStr);
		xls.SetCellText(nCol+4, nRow, tmpStr, FMT_NUMBER1);
		nGrpAmt += str2double(tmpStr);
		xls.SetCellText(nCol+6, nRow, rs.GetValue(_T("invno")), FMT_TEXT);
		nRow++;
		rs.MoveNext();
	}
	if (nGrpAmt > 0)
	{
		xls.SetCellText(nCol+3, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT, true);
		tmpStr.Format(_T("%f"), nGrpAmt);
		xls.SetCellText(nCol+4, nRow, tmpStr, FMT_NUMBER1, true);
		nTtlAmt += nGrpAmt;
		nRow++;
	}
	if (nTtlAmt > 0)
	{
		xls.SetCellText(nCol+3, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT, true);
		tmpStr.Format(_T("%f"), nTtlAmt);
		xls.SetCellText(nCol+4, nRow, tmpStr, FMT_NUMBER1, true);
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\TONGKETSUDUNGTHUOCTAIDONVI.xls"));
} 

CString CPMSImportSheetList::GetQueryString(){
	CMainFrame_E10 * pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szTransactionWhere, szPurchaseWhere, szStorages, szCondition;
	szTransactionWhere.Format(_T(" AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szPurchaseWhere.Format(_T(" AND po_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if(!m_szSupplierKey.IsEmpty())
	{
		szTransactionWhere.AppendFormat(_T(" AND mt_partner_id = '%s'"), m_szSupplierKey);
		szPurchaseWhere.AppendFormat(_T(" AND po_partner_id = '%s'"), m_szSupplierKey);
	}
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
		szTransactionWhere.AppendFormat(_T(" AND mt_storage_to_id IN (%s)"), szStorages);
		szPurchaseWhere.AppendFormat(_T(" AND po_storage_id IN (%s)"), szStorages);
	}
	//if (!m_szItemGroupKey.IsEmpty())
	//	szWhere.AppendFormat(_T(" AND product_categoryid = %d"), ToInt(m_szItemGroupKey));
	szTransactionWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	szPurchaseWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	if (!m_szItemGroupKey.IsEmpty())
	{
		szTransactionWhere.AppendFormat(_T(" AND product_groupid = '%s'"), m_szItemGroupKey);
		szPurchaseWhere.AppendFormat(_T(" AND product_groupid = '%s'"), m_szItemGroupKey);
	}
	//Neu nhap k thanh toan
	if (m_bNonInvoice)
	{
		szTransactionWhere.AppendFormat(_T(" AND mt_doctype = 'DRO'"));
		szPurchaseWhere.AppendFormat(_T(" AND 1 = 0"));
	}
	else if (m_bCompany)//Nhap cong ty
	{
		szTransactionWhere.AppendFormat(_T(" AND 1 = 0"));
		szPurchaseWhere.AppendFormat(_T(" AND po_partner_id = '708'"));
	}
	else //Nhap thanh toan
	{
		szTransactionWhere.AppendFormat(_T(" AND (mt_doctype = 'MOV' AND mt_storage_id IN (3, 14, 15))"));
		szPurchaseWhere.AppendFormat(_T(""));
	}
	if (m_bGroupbyType)
		szCondition = _T(" product_groupid, product_groupname,");
	szSQL.Format(_T(" SELECT doctype, ") \
				_T("        ordno, ") \
				_T("        invno, ") \
				_T("        dept, ") \
				_T("		%s") \
				_T("        impdte, ") \
				_T("        sum(ttlamt) ttlamt ") \
				_T(" FROM   (SELECT    mt_doctype                           doctype, ") \
				_T("                   mt_orderno                           AS ordno, ") \
				_T("                   Cast (0 AS NVARCHAR2(1))             AS invno, ") \
				_T("                   Get_departmentname(mt_department_id) AS dept, ") \
				_T("                   Cast_timestamp2date(mt_approveddate) AS impdte, ") \
				_T("				   product_groupid,") \
				_T("				   product_groupname,") \
				_T("                   mtl_ttlamount                   AS ttlamt ") \
				_T("         FROM      m_transaction ") \
				_T("         LEFT JOIN m_transactionline ON ( mt_transaction_id = mtl_transaction_id ) ") \
				_T("         LEFT JOIN m_productitem_view ON ( mtl_product_item_id = product_item_id ) ") \
				_T("         WHERE     mt_status = 'A' %s ") \
				_T("         UNION ALL ") \
				_T("         SELECT    po_doctype, ") \
				_T("                   po_orderno, ") \
				_T("                   po_invoiceno, ") \
				_T("                   Get_partnername(po_partner_id), ") \
				_T("                   Cast_timestamp2date(po_approveddate), ") \
				_T("				   product_groupid,") \
				_T("				   product_groupname,") \
				_T("                   pol_totalamount ") \
				_T("         FROM      purchase_order ") \
				_T("         LEFT JOIN purchase_orderline ON ( po_purchase_order_id = pol_purchase_order_id ) ") \
				_T("         LEFT JOIN m_productitem_view ON ( product_item_id = pol_product_item_id ) ") \
				_T("         WHERE     po_status = 'A' %s) ") \
				_T(" GROUP BY doctype, ") \
				_T("          ordno, ") \
				_T("          invno, ") \
				_T("          dept, ") \
				_T("		  %s") \
				_T("		  impdte") \
				_T(" ORDER  BY doctype, ") \
				_T("		   %s") \
				_T("           dept "), szCondition, szTransactionWhere, szPurchaseWhere, szCondition, szCondition);
	//_msg(_T("%s"), szSQL);
	return szSQL;
}

CString CPMSImportSheetList::GetQueryString1(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szTransactionWhere, szPurchaseWhere, szStorages;
	szTransactionWhere.Format(_T(" AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szPurchaseWhere.Format(_T(" AND po_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if(!m_szSupplierKey.IsEmpty())
	{
		szTransactionWhere.AppendFormat(_T(" AND mt_partner_id = '%s'"), m_szSupplierKey);
		szPurchaseWhere.AppendFormat(_T(" AND po_partner_id = '%s'"), m_szSupplierKey);
	}
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
		szTransactionWhere.AppendFormat(_T(" AND mt_storage_to_id IN (%s)"), szStorages);
		szPurchaseWhere.AppendFormat(_T(" AND po_storage_id IN (%s)"), szStorages);
	}
	//if (!m_szItemGroupKey.IsEmpty())
	//	szWhere.AppendFormat(_T(" AND product_categoryid = %d"), ToInt(m_szItemGroupKey));
	szTransactionWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	szPurchaseWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	if (!m_szItemGroupKey.IsEmpty())
	{
		szTransactionWhere.AppendFormat(_T(" AND product_groupid = '%s'"), m_szItemGroupKey);
		szPurchaseWhere.AppendFormat(_T(" AND product_groupid = '%s'"), m_szItemGroupKey);
	}
	//Neu nhap k thanh toan
	if (m_bNonInvoice)
	{
		szTransactionWhere.AppendFormat(_T(" AND mt_doctype = 'DRO'"));
		szPurchaseWhere.AppendFormat(_T(" AND 1 = 0"));
	}
	else if (m_bCompany)//Nhap cong ty
	{
		szTransactionWhere.AppendFormat(_T(" AND 1 = 0"));
		szPurchaseWhere.AppendFormat(_T(" AND po_partner_id = '708'"));
	}
	else //Nhap thanh toan
	{
		szTransactionWhere.AppendFormat(_T(" AND (mt_doctype = 'MOV' AND mt_storage_id IN (3, 14, 15))"));
		szPurchaseWhere.AppendFormat(_T(""));
	}
// 	szSQL.Format(_T("SELECT sum(amt) amt, sum(non_invoice) noninvoice, sum(invoice) invoice FROM (") \
// 		_T(" SELECT CASE WHEN po_partner_id = '708' THEN po_totalamount ELSE 0 END amt, ") \
// 		_T(" 0 non_invoice, 0 invoice FROM purchase_order ") \
// 		_T(" LEFT JOIN purchase_orderline ON (pol_purchase_order_id = po_purchase_order_id)") \
// 		_T(" LEFT JOIN m_productitem_view ON (product_item_id = pol_product_item_id) WHERE 1=1 %s") \
// 		_T(" UNION ALL") \
// 		_T(" SELECT 0, CASE WHEN mt_doctype = 'DRO' THEN mt_totalamount ELSE 0 END,") \
// 		_T(" CASE WHEN (mt_doctype = 'MOV' AND mt_storage_id IN (3, 14, 15)) THEN mt_totalamount ELSE 0 END FROM m_transaction") \
// 		_T(" LEFT JOIN m_transactionline ON (mt_transaction_id = mtl_transaction_id)") \
// 		_T(" LEFT JOIN m_productitem_view ON (product_item_id = mtl_product_item_id)") \
// 		_T(" WHERE 1=1 %s) tbl0"), szPurchaseWhere, szTransactionWhere);

	szSQL.Format(_T(" SELECT SUM(kho708)kho708,") \
		_T("   SUM(caccongty) caccongty,") \
		_T("   SUM(muale) muale") \
		_T(" FROM") \
		_T("   ( SELECT") \
		_T("   CASE") \
		_T("       WHEN mt_department_id = '708'") \
		_T("       THEN mtl_ttlamount") \
		_T("       ELSE 0") \
		_T("     END kho708,") \
		_T("     CASE") \
		_T("       WHEN mt_department_id NOT IN ('708', 'ML')") \
		_T("       THEN mtl_ttlamount") \
		_T("       ELSE 0") \
		_T("     END caccongty,") \
		_T("     CASE") \
		_T("       WHEN mt_department_id = 'ML'") \
		_T("       THEN mtl_ttlamount") \
		_T("       ELSE 0") \
		_T("     END muale") \
		_T("   FROM m_transaction") \
		_T("   LEFT JOIN m_transactionline") \
		_T("   ON ( mt_transaction_id = mtl_transaction_id )") \
		_T("   LEFT JOIN m_productitem_view") \
		_T("   ON ( mtl_product_item_id = product_item_id )") \
		_T("   WHERE mt_status          = 'A' %s ") \
		_T("   UNION ALL") \
		_T("   SELECT") \
		_T("     CASE") \
		_T("       WHEN po_partner_id = '708'") \
		_T("       THEN pol_totalamount") \
		_T("       ELSE 0") \
		_T("     END kho708,") \
		_T("     CASE") \
		_T("       WHEN po_partner_id NOT IN ('708', 'ML')") \
		_T("       THEN pol_totalamount") \
		_T("       ELSE 0") \
		_T("     END caccongty,") \
		_T("     CASE") \
		_T("       WHEN po_partner_id = 'ML'") \
		_T("       THEN pol_totalamount") \
		_T("       ELSE 0") \
		_T("     END muale") \
		_T("   FROM purchase_order") \
		_T("   LEFT JOIN purchase_orderline") \
		_T("   ON ( po_purchase_order_id = pol_purchase_order_id )") \
		_T("   LEFT JOIN m_productitem_view") \
		_T("   ON ( product_item_id = pol_product_item_id )") \
		_T("   WHERE po_status      = 'A' %s") \
		_T("   ) tbl"), szTransactionWhere, szPurchaseWhere);
	//_msg(_T("%s"), szSQL);
	return szSQL;

}

BEGIN_MESSAGE_MAP(CPMSImportSheetList, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSImportSheetList::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
