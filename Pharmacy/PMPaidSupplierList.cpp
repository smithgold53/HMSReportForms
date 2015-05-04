#include "stdafx.h"
#include "PMPaidSupplierList.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include <typeinfo.h>
#include ".\pmpaidsupplierlist.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMPaidSupplierList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMPaidSupplierList *)pWnd)->OnToDateCheckValue();
} 
static long _OnSupplierLoadDataFnc(CWnd *pWnd){
	return ((CPMPaidSupplierList *)pWnd)->OnSupplierLoadData();
}
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMPaidSupplierList* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMPaidSupplierList *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnStorageAddNew();
}*/
static void _OnItemGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMPaidSupplierList* )pWnd)->OnItemGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnItemGroupSelendokFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnItemGroupSelendok();
}
/*static void _OnItemGroupSetfocusFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnItemGroupKillfocus();
}*/
/*static void _OnItemGroupKillfocusFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnItemGroupKillfocus();
}*/
static long _OnItemGroupLoadDataFnc(CWnd *pWnd){
	return ((CPMPaidSupplierList *)pWnd)->OnItemGroupLoadData();
}
/*static void _OnItemGroupAddNewFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnItemGroupAddNew();
}*/
static void _OnSourceSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMPaidSupplierList* )pWnd)->OnSourceSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSourceSelendokFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnSourceSelendok();
}
/*static void _OnSourceSetfocusFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnSourceKillfocus();
}*/
/*static void _OnSourceKillfocusFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnSourceKillfocus();
}*/
static long _OnSourceLoadDataFnc(CWnd *pWnd){
	return ((CPMPaidSupplierList *)pWnd)->OnSourceLoadData();
}
/*static void _OnSourceAddNewFnc(CWnd *pWnd){
	((CPMPaidSupplierList *)pWnd)->OnSourceAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMPaidSupplierList *pVw = (CPMPaidSupplierList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMPaidSupplierList *pVw = (CPMPaidSupplierList *)pWnd;
	pVw->OnExportSelect();
} 
CPMPaidSupplierList::CPMPaidSupplierList(CWnd *pParent){
	
	m_nDlgWidth = 585;
	m_nDlgHeight = 359;
	SetDefaultValues();
}
CPMPaidSupplierList::~CPMPaidSupplierList(){
}
void CPMPaidSupplierList::OnCreateComponents(){
	m_wndGeneralDepartmentUsage.Create(this, _T("Paid Supplier List"), 5, 5, 575, 150);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndSupplierLabel.Create(this, _T("Supplier"), 10, 60, 90, 85);
	m_wndSupplier.Create(this,95, 60, 570, 85);
	m_wndStorageLabel.Create(this, _T("Storage"), 10, 90, 90, 115);
	m_wndStorage.SetCheckBox(true);
	m_wndStorage.Create(this,95, 90, 570, 115);
	m_wndItemGroupLabel.Create(this, _T("Item Group"), 10, 120, 90, 145);
	m_wndItemGroup.Create(this,95, 120, 290, 145); 
	m_wndSourceLabel.Create(this, _T("Source"), 295, 120, 375, 145);
	m_wndSource.Create(this,380, 120, 570, 145); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 395, 155, 490, 180);
	m_wndExport.Create(this, _T("&Export"), 495, 155, 575, 180);

}
void CPMPaidSupplierList::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	//m_wndStorage.SetCheckValue(true);
	m_wndStorage.LimitText(1024);
	//m_wndItemGroup.SetCheckValue(true);
	m_wndItemGroup.LimitText(1024);
	//m_wndSource.SetCheckValue(true);
	m_wndSource.LimitText(1024);

	m_wndSupplier.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndSupplier.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);

	m_wndStorage.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 100);
	m_wndStorage.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);

	m_wndItemGroup.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndItemGroup.InsertColumn(1, _T("Code"), CFMT_TEXT, 150);

	m_wndSource.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndSource.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);
}
void CPMPaidSupplierList::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndSupplier.SetEvent(WE_LOADDATA, _OnSupplierLoadDataFnc);
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
	m_wndSource.SetEvent(WE_SELENDOK, _OnSourceSelendokFnc);
	//m_wndSource.SetEvent(WE_SETFOCUS, _OnSourceSetfocusFnc);
	//m_wndSource.SetEvent(WE_KILLFOCUS, _OnSourceKillfocusFnc);
	m_wndSource.SetEvent(WE_SELCHANGE, _OnSourceSelectChangeFnc);
	m_wndSource.SetEvent(WE_LOADDATA, _OnSourceLoadDataFnc);
	//m_wndSource.SetEvent(WE_ADDNEW, _OnSourceAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T(" 00:00");
	m_szToDate += _T(" 23:59");
	UpdateData(false);

}
void CPMPaidSupplierList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndItemGroup.GetDlgCtrlID(), m_szItemGroupKey);
	DDX_TextEx(pDX, m_wndSource.GetDlgCtrlID(), m_szSourceKey);
	DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szSupplierKey);
	DDX_TextEx(pDX, m_wndStorage.GetDlgCtrlID(), m_szStorageKey);

}
void CPMPaidSupplierList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szItemGroupKey.Empty();
	m_szSourceKey.Empty();
	m_szSupplierKey.Empty();
	m_szStorageKey.Empty();

}
int CPMPaidSupplierList::SetMode(int nMode){
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
/*void CPMPaidSupplierList::OnFromDateChange(){
	
} */
/*void CPMPaidSupplierList::OnFromDateSetfocus(){
	
} */
/*void CPMPaidSupplierList::OnFromDateKillfocus(){
	
} */
int CPMPaidSupplierList::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMPaidSupplierList::OnToDateChange(){
	
} */
/*void CPMPaidSupplierList::OnToDateSetfocus(){
	
} */
/*void CPMPaidSupplierList::OnToDateKillfocus(){
	
} */
int CPMPaidSupplierList::OnToDateCheckValue(){
	return 0;
} 
long CPMPaidSupplierList::OnSupplierLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadPartnerList(&m_wndSupplier, m_szSupplierKey);
}
void CPMPaidSupplierList::OnItemGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
void CPMPaidSupplierList::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMPaidSupplierList::OnStorageSelendok(){
	 
}
long CPMPaidSupplierList::OnStorageLoadData(){
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
/*void CPMPaidSupplierList::OnStorageAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMPaidSupplierList::OnItemGroupSelendok(){
	 
}
/*void CPMPaidSupplierList::OnItemGroupSetfocus(){
	
}*/
/*void CPMPaidSupplierList::OnItemGroupKillfocus(){
	
}*/
long CPMPaidSupplierList::OnItemGroupLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//return pMF->LoadProductTypeList(&m_wndType, m_szTypeKey);
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
/*void CPMPaidSupplierList::OnItemGroupAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMPaidSupplierList::OnSourceSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMPaidSupplierList::OnSourceSelendok(){
	 
}
/*void CPMPaidSupplierList::OnSourceSetfocus(){
	
}*/
/*void CPMPaidSupplierList::OnSourceKillfocus(){
	
}*/
long CPMPaidSupplierList::OnSourceLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadPaymentResourceList(&m_wndSource, m_szSourceKey);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndSource.IsSearchKey() && !m_szSourceKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szSourceKey
};
	m_wndSource.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndSource.AddItems(
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CPMPaidSupplierList::OnSourceAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMPaidSupplierList::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CReport rpt;
	UpdateData(true);
	CString szSQL, tmpStr, tmpStr2;
	CString szMoney, szDate, szStorages;
	CRecord rs(&pMF->m_db);
	CReportSection *rptHeader, *rptDetail, *rptFooter;
	int nIdx = 0;
	double nAmt = 0;
	long double nTotalAmount = 0;
	if (m_szSourceKey.IsEmpty())
	{
		AfxMessageBox(_T("Choose a source!"));
		m_wndSource.SetFocus();
		return;
	}
	if (!rpt.Init(_T("Reports\\HMS\\PM_BANGKETHANHTOANTRANOCACCONGTY.RPT")))
		return;
	szSQL = GetQueryString();
	//QueryOpener(szSQL);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	rptHeader->SetValue(_T("Department"), pMF->GetDepartmentName(pMF->GetDepartmentID()));
	tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
		CDate::Convert(m_szToDate), yyyymmdd, ddmmyyyy);
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
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
		rptHeader->SetValue(_T("Storage"), szStorages);
	else
		rptHeader->SetValue(_T("Storage"), _T("To\xE0n \x62\x1ED9"));
	if (!m_szSourceKey.IsEmpty())
		rptHeader->SetValue(_T("Source"), m_wndSource.GetCurrent(1));
	if (!m_szItemGroupKey.IsEmpty())
		rptHeader->SetValue(_T("ItemType"), m_wndItemGroup.GetCurrent(1));
	else
		rptHeader->SetValue(_T("ItemType"), _T("To\xE0n \x62\x1ED9"));
	if (m_szFromDate.Left(10) != m_szToDate.Left(10))
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	else
		tmpStr.Format(_T("Ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		nIdx++;
		tmpStr.IsEmpty();
		rptDetail->SetValue(_T("1"), int2str(nIdx));
		tmpStr.Format(rptDetail->GetValue(_T("2")), rs.GetValue(_T("nme")), rs.GetValue(_T("addr")), rs.GetValue(_T("contr")), rs.GetValue(_T("recno")));
		rptDetail->SetValue(_T("2"), tmpStr);
		if(!rs.GetValue(_T("acc")).IsEmpty())
		{
			tmpStr2.Format(_T("TK %s"), rs.GetValue(_T("acc")));
		}
		else
		{
			tmpStr2.Format(_T(""));
		}
		tmpStr.Format(rptDetail->GetValue(_T("3")), tmpStr2, rs.GetValue(_T("banknme")), rs.GetValue(_T("bankbranch")), rs.GetValue(_T("code")));
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("total"), nAmt);
		nTotalAmount+=nAmt;
		rptDetail->SetValue(_T("4"), double2str(nAmt));
		rs.MoveNext();
	}
	rptFooter = rpt.GetReportFooter();
	if (!m_szSourceKey.IsEmpty())
		rptFooter->SetValue(_T("S1"), m_wndSource.GetCurrent(1));
	rptFooter->SetValue(_T("s4"), ToString(nTotalAmount));
	FormatCurrency(nTotalAmount, szMoney);
	tmpStr.Format(rptFooter->GetValue(_T("Statistic")), nIdx, szMoney);
	rptFooter->SetValue(_T("Statistic"), tmpStr);
	rptFooter->SetValue(_T("Amount"), szMoney);
	szDate = pMF->GetSysDate();
	tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5,2), szDate.Left(4));
	rptFooter->SetValue(_T("PrintDate"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("Username"), pMF->GetCurrentUser());
	rpt.PrintPreview();
} 
void CPMPaidSupplierList::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr, szTemp;
	CRecord rs(&pMF->m_db);
	int nIdx = 0;
	double nTmp = 0;
	long double nTotalAmt = 0;
	if (m_szSourceKey.IsEmpty())
	{
		AfxMessageBox(_T("Choose a source!"));
		m_wndSource.SetFocus();
		return;
	}
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
	xls.SetColumnWidth(1, 40);
	xls.SetColumnWidth(2, 40);
	xls.SetColumnWidth(3, 15);
	//xls.SetColumnWidth(4, 15);
	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 7);
	xls.SetCellText(nCol, nRow + 3, _T("\x42\x1EA2NG K\xCA TH\x41NH TO\xC1N TR\x1EA2 N\x1EE2 \x43\xC1\x43 \x43\xD4NG TY"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 7);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	//Col Head
	nRow = 5;
	xls.SetCellText(nCol, nRow, _T("STT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+1, nRow, _T("\x43\xF4ng ty - \x111\x1ECB\x61 \x63h\x1EC9"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol+2, nRow, _T("T\xE0i kho\x1EA3n ng\xE2n h\xE0ng"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol+3, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		nIdx++;
		tmpStr.IsEmpty();
		xls.SetCellText(nCol, nRow, int2str(nIdx), FMT_TEXT | FMT_RIGHT);
		tmpStr = rs.GetValue(_T("nme"));
		if (!rs.GetValue(_T("addr")).IsEmpty())
			tmpStr.AppendFormat(_T(" - %s"), rs.GetValue(_T("addr")));
		if (!rs.GetValue(_T("contr")).IsEmpty())
			tmpStr.AppendFormat(_T(" - %s"), rs.GetValue(_T("contr")));
		if (!rs.GetValue(_T("recno")).IsEmpty())
			tmpStr.AppendFormat(_T(" - %s"), rs.GetValue(_T("recno")));
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);
		if (!rs.GetValue(_T("acc")).IsEmpty())
			tmpStr.Format(_T("TK %s"), rs.GetValue(_T("acc")));
		if (!rs.GetValue(_T("banknme")).IsEmpty())
			tmpStr.AppendFormat(_T(" - %s"), rs.GetValue(_T("banknme")));
		if (!rs.GetValue(_T("bankbranch")).IsEmpty())
			tmpStr.AppendFormat(_T(" - \x43hi nh\xE1nh %s"), rs.GetValue(_T("bankbranch")));
		if (!rs.GetValue(_T("code")).IsEmpty())
			tmpStr.AppendFormat(_T(" - %s"), rs.GetValue(_T("code")));
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);
		rs.GetValue(_T("total"), nTmp);
		nTotalAmt += nTmp;
		xls.SetCellText(nCol + 3, nRow, double2str(nTmp), FMT_NUMBER1 | FMT_WRAPING);
		nRow++;
		rs.MoveNext();
	}
	if (nTotalAmt > 0)
	{
		FormatCurrency(nTotalAmt, szTemp);
		tmpStr.Format(_T("K\x1EBFt to\xE1n \x62\x1EA3ng n\xE0y g\x1ED3m %d \x63\xF4ng ty v\x1EDBi t\x1ED5ng ti\x1EC1n l\xE0 %s"), nIdx, szTemp);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\BANG KE THANH TOAN TRA NO CAC CONG TY.xls"));
	
} 

CString CPMPaidSupplierList::GetQueryString(){
	CString szSQL, szWhere, szStorages;
	szWhere.Format(_T(" AND pp_createddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
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
	if(!szStorages.IsEmpty())
		szWhere.AppendFormat(_T(" AND po_storage_id IN(%s)"), szStorages);
	if(!m_szSupplierKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND po_partner_id = '%s'"), m_szSupplierKey);
	if (!m_szItemGroupKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND product_groupid = '%s'"), m_szItemGroupKey);
	if (!m_szSourceKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND pp_payment_resource_id = %s"), m_szSourceKey);
	szSQL.Format(_T(" SELECT") \
				_T("   po_partner_id,") \
				_T("   adp_name            AS nme,") \
				_T("   adp_bankaccount     AS acc,") \
				_T("   adb_name            AS banknme,") \
				_T("   adp_address         AS addr,") \
				_T("   adp_branch          AS bankbranch,") \
				_T("   adp_fax             AS contr,") \
				_T("   adp_email           AS recno,") \
				_T("   adp_description     AS code,") \
				_T("   Sum(pol_totalamount) AS total") \
				_T(" FROM   purchase_order") \
				_T(" LEFT JOIN purchase_orderline ON (po_purchase_order_id = pol_purchase_order_id)") \
				_T(" LEFT JOIN m_productitem_view ON (product_item_id = pol_product_item_id)") \
				_T(" LEFT JOIN purchase_payment ON (po_refpayment_id=pp_purchase_payment_id)") \
				_T(" LEFT JOIN ad_partner ON (adp_partner_id=po_partner_id)") \
				_T(" LEFT JOIN ad_bank ON (adp_bank_id=adb_bank_id)") \
				_T(" WHERE 1=1 %s") \
				_T(" GROUP  BY po_partner_id,") \
				_T("           adp_name,") \
				_T("           adp_bankaccount,") \
				_T("           adb_name,") \
				_T("           adp_address,") \
				_T("           adp_branch,") \
				_T("           adp_fax,") \
				_T("		   adp_email,") \
				_T("		   adp_description") \
				_T(" ORDER  BY po_partner_id "), szWhere);
	//_msg(_T("%s"), szSQL);
	return szSQL;
}BEGIN_MESSAGE_MAP(CPMPaidSupplierList, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMPaidSupplierList::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
