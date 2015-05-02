#include "stdafx.h"
#include "PMSUnpaidOrderList.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include <typeinfo.h>
#include ".\pmssupplierpaymentlist.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSUnpaidOrderList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSUnpaidOrderList *)pWnd)->OnToDateCheckValue();
} 
static long _OnSupplierLoadDataFnc(CWnd *pWnd){
	return ((CPMSUnpaidOrderList*)pWnd)->OnSupplierLoadData();
}
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSUnpaidOrderList* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMSUnpaidOrderList *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnStorageAddNew();
}*/
static void _OnItemGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSUnpaidOrderList* )pWnd)->OnItemGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnItemGroupSelendokFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnItemGroupSelendok();
}
/*static void _OnItemGroupSetfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnItemGroupKillfocus();
}*/
/*static void _OnItemGroupKillfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnItemGroupKillfocus();
}*/
static long _OnItemGroupLoadDataFnc(CWnd *pWnd){
	return ((CPMSUnpaidOrderList *)pWnd)->OnItemGroupLoadData();
}
/*static void _OnItemGroupAddNewFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnItemGroupAddNew();
}*/
static void _OnSourceSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSUnpaidOrderList* )pWnd)->OnSourceSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSourceSelendokFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnSourceSelendok();
}
/*static void _OnSourceSetfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnSourceKillfocus();
}*/
/*static void _OnSourceKillfocusFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnSourceKillfocus();
}*/
static long _OnSourceLoadDataFnc(CWnd *pWnd){
	return ((CPMSUnpaidOrderList *)pWnd)->OnSourceLoadData();
}
/*static void _OnSourceAddNewFnc(CWnd *pWnd){
	((CPMSUnpaidOrderList *)pWnd)->OnSourceAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSUnpaidOrderList *pVw = (CPMSUnpaidOrderList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSUnpaidOrderList *pVw = (CPMSUnpaidOrderList *)pWnd;
	pVw->OnExportSelect();
} 
CPMSUnpaidOrderList::CPMSUnpaidOrderList(CWnd *pParent){

	m_nDlgWidth = 585;
	m_nDlgHeight = 359;
	SetDefaultValues();
}
CPMSUnpaidOrderList::~CPMSUnpaidOrderList(){
}
void CPMSUnpaidOrderList::OnCreateComponents(){
	m_wndGeneralDepartmentUsage.Create(this, _T("Unpaid Order List"), 5, 5, 575, 150);
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
	m_wndSource.EnableWindow(FALSE);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 395, 155, 490, 180);
	m_wndExport.Create(this, _T("&Export"), 495, 155, 575, 180);

}
void CPMSUnpaidOrderList::OnInitializeComponents(){
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

	m_wndStorage.InsertColumn(0, _T("ID"), CFMT_TEXT |CFMT_RIGHT, 100);
	m_wndStorage.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);

	m_wndItemGroup.InsertColumn(0, _T("ID"), CFMT_NUMBER, 80);
	m_wndItemGroup.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

	m_wndSource.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndSource.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);
}
void CPMSUnpaidOrderList::OnSetWindowEvents(){
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
	m_wndSource.SetEvent(WE_LOADDATA, _OnSourceLoadDataFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T(" 00:00");
	m_szToDate += _T(" 23:59");
	UpdateData(false);

}
void CPMSUnpaidOrderList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szSupplierKey);
	DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szStorageKey);
	DDX_TextEx(pDX, m_wndItemGroup.GetDlgCtrlID(), m_szItemGroupKey);

}
void CPMSUnpaidOrderList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szItemGroupKey.Empty();
	m_szSourceKey.Empty();
	m_szSupplierKey.Empty();
	m_szStorageKey.Empty();

}
int CPMSUnpaidOrderList::SetMode(int nMode){
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
/*void CPMSUnpaidOrderList::OnFromDateChange(){
	
} */
/*void CPMSUnpaidOrderList::OnFromDateSetfocus(){
	
} */
/*void CPMSUnpaidOrderList::OnFromDateKillfocus(){
	
} */
int CPMSUnpaidOrderList::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSUnpaidOrderList::OnToDateChange(){
	
} */
/*void CPMSUnpaidOrderList::OnToDateSetfocus(){
	
} */
/*void CPMSUnpaidOrderList::OnToDateKillfocus(){
	
} */
int CPMSUnpaidOrderList::OnToDateCheckValue(){
	return 0;
} 
long CPMSUnpaidOrderList::OnSupplierLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadPartnerList(&m_wndSupplier, m_szSupplierKey);
}
void CPMSUnpaidOrderList::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSUnpaidOrderList::OnStorageSelendok(){
	 
}
long CPMSUnpaidOrderList::OnStorageLoadData(){
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
/*void CPMSUnpaidOrderList::OnStorageAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSUnpaidOrderList::OnItemGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSUnpaidOrderList::OnItemGroupSelendok(){
	 
}
/*void CPMSUnpaidOrderList::OnItemGroupSetfocus(){
	
}*/
/*void CPMSUnpaidOrderList::OnItemGroupKillfocus(){
	
}*/
long CPMSUnpaidOrderList::OnItemGroupLoadData(){
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
/*void CPMSUnpaidOrderList::OnItemGroupAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSUnpaidOrderList::OnSourceSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSUnpaidOrderList::OnSourceSelendok(){
	 
}
/*void CPMSUnpaidOrderList::OnSourceSetfocus(){
	
}*/
/*void CPMSUnpaidOrderList::OnSourceKillfocus(){
	
}*/
long CPMSUnpaidOrderList::OnSourceLoadData(){
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
/*void CPMSUnpaidOrderList::OnSourceAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSUnpaidOrderList::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString szSQL, tmpStr, szOrderID, szNewLine, szOldLine, szDate, szOldID, szNewID, szTemp, szField, szStorages;
	float nTmp = 0, nAmt = 0, nGroupTotal = 0, nTotal = 0;
	int nIdx = 0, nIndex = 0;
	CRecord rs(&pMF->m_db);
	CReportSection* rptHeader, * rptDetail;
	if (!rpt.Init(_T("Reports/HMS/PM_BANGKEHANGMUACHUATRATIEN.RPT")))
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
	if (!m_szItemGroupKey.IsEmpty())
		rptHeader->SetValue(_T("ItemType"), m_wndItemGroup.GetCurrent(1));
	else
		rptHeader->SetValue(_T("ItemType"), _T("To\xE0n \x62\x1ED9"));
	if (!m_szSupplierKey.IsEmpty())
		rptHeader->SetValue(_T("Supplier"), m_wndSupplier.GetCurrent(1));
	else
		rptHeader->SetValue(_T("Supplier"), _T("To\xE0n \x62\x1ED9"));
	if (m_szFromDate.Left(10) != m_szToDate.Left(10))
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	else
		tmpStr.Format(_T("Ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	szSQL = GetQueryString1();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		_fmsg(_T("%s"), szSQL);
		return;
	}
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("partnerid"), szNewID);
		rs.GetValue(_T("partnername"), szNewLine);
		if (szNewLine != szOldLine)
		{
			nIndex++;
			nIdx = 0;
			if (nGroupTotal > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("tongso"), _T("\x43\x1ED9ng"));
				tmpStr.Format(_T("%.2f"), nGroupTotal);
				rptDetail->SetValue(_T("s4"), tmpStr);
				nTotal += nGroupTotal;
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("tongso"), _T("N\x1EE3 \x63\xE1\x63 th\xE1ng tr\x1B0\x1EDB\x63"));
				nAmt = GetDebtAmount(szOldID, m_szFromDate);
				tmpStr.Format(_T("%.2f"), nAmt);
				rptDetail->SetValue(_T("s4"), tmpStr);
				nGroupTotal += nAmt;
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("tongso"), _T("T\x1ED5ng \x43\x1ED9ng"));
				tmpStr.Format(_T("%.2f"), nGroupTotal);
				rptDetail->SetValue(_T("s4"), tmpStr);
				nGroupTotal = 0;
			}
			szOldID = szNewID;
			szOldLine = szNewLine;
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			tmpStr.Format(rptDetail->GetValue(_T("GroupName")), nIndex, szNewLine);
			rptDetail->SetValue(_T("GroupName"), tmpStr);
		}
		rptDetail = rpt.AddDetail();
		nIdx++;
		rptDetail->SetValue(_T("1"), int2str(nIdx));
		if (!szOrderID.IsEmpty())
			szOrderID += _T(", ");
		szOrderID += rs.GetValue(_T("id"));
		rs.GetValue(_T("invoicedate"), tmpStr);
		rptDetail->SetValue(_T("2"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("invoiceno")));
		rs.GetValue(_T("amount"), nAmt);
		nGroupTotal += nAmt;
		tmpStr.Format(_T("%.2f"), nAmt);
		rptDetail->SetValue(_T("4"), tmpStr);
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("importmonth")));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("orderno")));
		rptDetail->SetValue(_T("7"), _T(""));
		rptDetail->SetValue(_T("8"), _T(""));
		rs.MoveNext();	
	}
	if (nGroupTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("tongso"), _T("\x43\x1ED9ng"));
		tmpStr.Format(_T("%.2f"), nGroupTotal);
		rptDetail->SetValue(_T("s4"), tmpStr);
		nTotal += nGroupTotal;
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("tongso"), _T("N\x1EE3 \x63\xE1\x63 th\xE1ng tr\x1B0\x1EDB\x63"));
		nAmt = GetDebtAmount(szOldID, m_szFromDate);
		tmpStr.Format(_T("%.2f"), nAmt);
		rptDetail->SetValue(_T("s4"), tmpStr);
		nGroupTotal += nAmt;
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("tongso"), _T("T\x1ED5ng \x43\x1ED9ng"));
		tmpStr.Format(_T("%.2f"), nGroupTotal);
		rptDetail->SetValue(_T("s4"), tmpStr);
	}
/*	if (nTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("s4"), ToString(nTotal));
	}*/	

	rptDetail = rpt.GetReportFooter();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("summary")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rptDetail->SetValue(_T("summary"), tmpStr);
	
	FormatCurrency(nTotal, szTemp);
	szTemp += _T(" \x111\x1ED3ng.");
	rptDetail->SetValue(_T("s7"), szTemp);
	szSQL = GetQueryString2();
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
				rptDetail->SetValue(tmpStr, szTemp);
				tmpStr.Format(_T("s%d"), nIdx);
				FormatCurrency(nTmp, szTemp);
				szTemp += _T(" \x111\x1ED3ng.");
				rptDetail->SetValue(tmpStr, szTemp);
				nIdx++;
			}
		}
	}
	szDate = pMF->GetSysDate();
	tmpStr.Format(rptDetail->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5, 2), szDate.Left(4));
	rptDetail->SetValue(_T("PrintDate"), tmpStr);
	rptDetail->SetValue(_T("Username"), pMF->GetCurrentUser());
	rpt.PrintPreview();
} 
void CPMSUnpaidOrderList::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr, szTemp, szOrderID;
	CString szOldLine, szNewLine;
	CRecord rs(&pMF->m_db);
	int nIdx = 0, nIndex = 1;
	double nTmp = 0, nGrpAmt = 0;
	long double nTotalAmt = 0;
	szSQL = GetQueryString1();
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
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(2, 12);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 12);
	xls.SetColumnWidth(5, 15);
	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 7);
	xls.SetCellText(nCol, nRow + 3, _T("\x44\x41NH S\xC1\x43H H\xC0NG MU\x41 \x43H\x1AF\x41 TR\x1EA2 TI\x1EC0N"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 7);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	//Col Head
	nRow = 5;
	xls.SetCellText(nCol, nRow, _T("STT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+1, nRow, _T("Ng\xE0y h\xF3\x61 \x111\x1A1n"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol+2, nRow, _T("S\x1ED1 h\xF3\x61 \x111\x1A1n"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol+3, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT | FMT_CENTER , true);
	xls.SetCellText(nCol+4, nRow, _T("Ng\xE0y nh\x1EADp"), FMT_TEXT | FMT_CENTER , true);
	xls.SetCellText(nCol+5, nRow, _T("S\x1ED1 phi\x1EBFu nh\x1EADp"), FMT_TEXT | FMT_CENTER , true);
	xls.SetCellText(nCol+6, nRow, _T("K\xFD nh\x1EADn"), FMT_TEXT | FMT_CENTER , true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("partnername"), szNewLine);
		if (szNewLine != szOldLine)
		{
			if (nGrpAmt > 0)
			{
				xls.SetCellMergedColumns(nCol, nRow, 5);
				xls.SetCellText(nCol, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
				tmpStr.Format(_T("%.2f"), nGrpAmt);
				xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1);
				nRow++;
				nTotalAmt += nGrpAmt;
				nGrpAmt = 0;
			}
			xls.SetCellMergedColumns(nCol, nRow, 8);
			xls.SetCellText(nCol, nRow, szNewLine, FMT_TEXT | FMT_CENTER, true);
			szOldLine = szNewLine;
			nIdx++;
			nRow++;
			nIndex = 1;
		}
		xls.SetCellText(nCol, nRow, int2str(nIndex), FMT_TEXT | FMT_RIGHT);
		rs.GetValue(_T("invoicedate"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), FMT_TEXT | FMT_CENTER);
		xls.SetCellText(nCol + 2, nRow, rs.GetValue(_T("invoiceno")), FMT_TEXT);
		rs.GetValue(_T("amount"), nTmp);
		nGrpAmt += nTmp;
		xls.SetCellText(nCol + 3, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("importmonth"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), FMT_TEXT | FMT_CENTER);
		xls.SetCellText(nCol + 5, nRow, rs.GetValue(_T("orderno")), FMT_TEXT);
		nRow++;
		rs.MoveNext();
	}
	if (nGrpAmt > 0)
	{
		xls.SetCellMergedColumns(nCol, nRow, 5);
		xls.SetCellText(nCol, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT, true);
		tmpStr.Format(_T("%.2f"), nGrpAmt);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1);
		nRow++;
		nTotalAmt += nGrpAmt;
	}
	if (nTotalAmt > 0)
	{
		szTemp.Format(_T("%.2f"), nTotalAmt);
		tmpStr.Format(_T("K\x1EBFt to\xE1n \x62\x1EA3ng n\xE0y g\x1ED3m %d \x63\xF4ng ty v\x1EDBi t\x1ED5ng ti\x1EC1n l\xE0 %s"), nIdx, szTemp);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
	}
	szSQL = GetQueryString2();
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{
		nRow++;
		rs.GetValue(_T("thuoc"), nTmp);
		szTemp.Format(_T("%ld"), nTmp);
		if (szTemp.IsEmpty())
			szTemp = _T("0");
		tmpStr.Format(_T("Thu\x1ED1\x63: %s \x111\x1ED3ng"), szTemp);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_NUMBER1);
		nRow++;
		rs.GetValue(_T("hoachat"), nTmp);
		szTemp.Format(_T("%ld"), nTmp);
		if (szTemp.IsEmpty())
			szTemp = _T("0");
		tmpStr.Format(_T("H\xF3\x61 \x63h\x1EA5t: %s \x111\x1ED3ng"), szTemp);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_NUMBER1);
		nRow++;
		rs.GetValue(_T("bbg"), nTmp);
		szTemp.Format(_T("%ld"), nTmp);
		if (szTemp.IsEmpty())
			szTemp = _T("0");
		tmpStr.Format(_T("\x42\xF4ng \x62\x103ng g\x1EA1\x63: %s \x111\x1ED3ng"), szTemp);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_NUMBER1);
		nRow++;
		rs.GetValue(_T("duoclieu"), nTmp);
		szTemp.Format(_T("%ld"), nTmp);
		if (szTemp.IsEmpty())
			szTemp = _T("0");
		tmpStr.Format(_T("\x44\x1B0\x1EE3\x63 li\x1EC7u: %s \x111\x1ED3ng"), szTemp);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_NUMBER1);
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\BANG KE THANH TOAN NO CAC CONG TY.xls"));
} 

CString CPMSUnpaidOrderList::GetQueryString1(){
	CString szSQL, szWhere, szStorages;
	szWhere.Format(_T("AND po_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
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
	szSQL.Format(_T(" SELECT") \
				_T("   po_purchase_order_id as id,") \
				_T("   po_invoiceno                   AS invoiceno,") \
				_T("   Cast_Timestamp2Date(po_invoicedate)                 AS invoicedate,") \
				_T("   po_orderno                     AS orderno,") \
				_T("   extract(month from po_approveddate)                AS importmonth,") \
				_T("   sum(pol_totalamount) AS amount,") \
				_T("   po_partner_id				  AS partnerid,") \
				_T("   Get_partnername(po_partner_id) AS partnername") \
				_T(" FROM   purchase_order") \
				_T(" LEFT JOIN purchase_orderline ON (po_purchase_order_id = pol_purchase_order_id)") \
				_T(" LEFT JOIN m_productitem_view ON (pol_product_item_id = product_item_id)") \
				_T(" WHERE  po_payment = 'N' AND po_status = 'A' AND product_org_id = 'PM' %s") \
				_T(" GROUP BY po_purchase_order_id, po_invoiceno, po_invoicedate, po_orderno, po_approveddate, po_partner_id") \
				_T(" ORDER  BY po_partner_id "), szWhere);
	return szSQL;
}

CString CPMSUnpaidOrderList::GetQueryString2(){
	CString szSQL, szWhere, szStorages;
	szWhere.Format(_T("AND po_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
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
	szSQL.Format(_T(" SELECT SUM(thuoc) drug, ") \
				_T("        SUM(hoachat) chemicals, ") \
				_T("        SUM(bbg) cotton, ") \
				_T("        SUM(duoclieu) herb,") \
				_T("        SUM(tpdd) compound,") \
				_T("        SUM(ycu) instrument") \
				_T(" FROM   (SELECT CASE WHEN (instr(product_producttype, 'A11') > 0 OR product_producttype IN ('A1400', 'A6000'))") \
				_T("                     THEN pol_totalamount ") \
				_T("                ELSE 0 ") \
				_T("                END AS thuoc, ") \
				_T("                CASE WHEN instr(product_producttype, 'A12') > 0 THEN pol_totalamount ") \
				_T("                ELSE 0 ") \
				_T("                END AS hoachat, ") \
				_T("                CASE WHEN product_producttype = 'A1500' THEN pol_totalamount ") \
				_T("                ELSE 0 ") \
				_T("                END AS bbg, ") \
				_T("                CASE WHEN product_producttype = 'A1302' THEN pol_totalamount ") \
				_T("                ELSE 0 ") \
				_T("                END AS duoclieu, ") \
				_T("				CASE WHEN product_producttype = 'A1301' THEN pol_totalamount") \
				_T("				ELSE 0") \
				_T("				END AS tpdd,") \
				_T("				CASE WHEN product_producttype = 'A1600' THEN pol_totalamount") \
				_T("				ELSE 0") \
				_T("				END AS ycu") \
				_T("         FROM   purchase_orderline, ") \
				_T("                m_productitem_view, ") \
				_T("                pms_drugtype, ") \
				_T("				purchase_order") \
				_T("         WHERE  pol_product_item_id = product_item_id ") \
				_T("            AND Instr(pmdt_desc, product_producttype) > 0 AND po_purchase_order_id = pol_purchase_order_id ") \
				_T("			AND po_payment = 'N' AND po_status = 'A' %s)"), szWhere);
	return szSQL;
}
BEGIN_MESSAGE_MAP(CPMSUnpaidOrderList, CGuiView)
	ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSUnpaidOrderList::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}

float CPMSUnpaidOrderList::GetDebtAmount(CString szPartnerID, CString szDate){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere, tmpStr, szStorages;
	float nTtlAmt = 0;
	szWhere.Format(_T(" AND po_approveddate < cast_string2timestamp('%s') AND po_partner_id = '%s'"), szDate, szPartnerID);
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
	if (!m_szItemGroupKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND product_groupid = '%s'"), m_szItemGroupKey);
	szSQL.Format(_T("SELECT sum(pol_totalamount) ttlamt ") \
		_T(" FROM purchase_order ") \
		_T(" LEFT JOIN purchase_orderline ON (po_purchase_order_id = pol_purchase_order_id) ") \
		_T(" LEFT JOIN m_productitem_view ON (product_item_id = pol_product_item_id) ") \
		_T(" WHERE po_status = 'A' AND po_payment = 'N' %s"), szWhere);
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
		rs.GetValue(_T("ttlamt"), nTtlAmt);
	return nTtlAmt;	
}