#include "stdafx.h"
#include "PMSSupplierPaymentList.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include <typeinfo.h>
#include ".\pmssupplierpaymentlist.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSSupplierPaymentList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSSupplierPaymentList *)pWnd)->OnToDateCheckValue();
} 
static long _OnSupplierLoadDataFnc(CWnd *pWnd){
	return ((CPMSSupplierPaymentList *)pWnd)->OnSupplierLoadData();
}
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSSupplierPaymentList* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMSSupplierPaymentList *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnStorageAddNew();
}*/
static void _OnItemGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSSupplierPaymentList* )pWnd)->OnItemGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnItemGroupSelendokFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnItemGroupSelendok();
}
/*static void _OnItemGroupSetfocusFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnItemGroupKillfocus();
}*/
/*static void _OnItemGroupKillfocusFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnItemGroupKillfocus();
}*/
static long _OnItemGroupLoadDataFnc(CWnd *pWnd){
	return ((CPMSSupplierPaymentList *)pWnd)->OnItemGroupLoadData();
}
/*static void _OnItemGroupAddNewFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnItemGroupAddNew();
}*/
static void _OnSourceSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSSupplierPaymentList* )pWnd)->OnSourceSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSourceSelendokFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnSourceSelendok();
}
/*static void _OnSourceSetfocusFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnSourceKillfocus();
}*/
/*static void _OnSourceKillfocusFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnSourceKillfocus();
}*/
static long _OnSourceLoadDataFnc(CWnd *pWnd){
	return ((CPMSSupplierPaymentList *)pWnd)->OnSourceLoadData();
}
/*static void _OnSourceAddNewFnc(CWnd *pWnd){
	((CPMSSupplierPaymentList *)pWnd)->OnSourceAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSSupplierPaymentList *pVw = (CPMSSupplierPaymentList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSSupplierPaymentList *pVw = (CPMSSupplierPaymentList *)pWnd;
	pVw->OnExportSelect();
} 
CPMSSupplierPaymentList::CPMSSupplierPaymentList(CWnd *pParent){

	m_nDlgWidth = 585;
	m_nDlgHeight = 359;
	SetDefaultValues();
}
CPMSSupplierPaymentList::~CPMSSupplierPaymentList(){
}
void CPMSSupplierPaymentList::OnCreateComponents(){
	m_wndGeneralDepartmentUsage.Create(this, _T("Import Sheet List"), 5, 5, 575, 150);
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
void CPMSSupplierPaymentList::OnInitializeComponents(){
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

	m_wndSource.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndSource.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

	m_wndItemGroup.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndItemGroup.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);
}
void CPMSSupplierPaymentList::OnSetWindowEvents(){
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
void CPMSSupplierPaymentList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szSupplierKey);
	DDX_TextEx(pDX, m_wndStorage.GetDlgCtrlID(), m_szStorageKey);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndItemGroup.GetDlgCtrlID(), m_szItemGroupKey);
	DDX_TextEx(pDX, m_wndSource.GetDlgCtrlID(), m_szSourceKey);

}
void CPMSSupplierPaymentList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szItemGroupKey.Empty();
	m_szSourceKey.Empty();
	m_szSupplierKey.Empty();
	m_szStorageKey.Empty();

}
int CPMSSupplierPaymentList::SetMode(int nMode){
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
/*void CPMSSupplierPaymentList::OnFromDateChange(){
	
} */
/*void CPMSSupplierPaymentList::OnFromDateSetfocus(){
	
} */
/*void CPMSSupplierPaymentList::OnFromDateKillfocus(){
	
} */
int CPMSSupplierPaymentList::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSSupplierPaymentList::OnToDateChange(){
	
} */
/*void CPMSSupplierPaymentList::OnToDateSetfocus(){
	
} */
/*void CPMSSupplierPaymentList::OnToDateKillfocus(){
	
} */
int CPMSSupplierPaymentList::OnToDateCheckValue(){
	return 0;
} 
long CPMSSupplierPaymentList::OnSupplierLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadPartnerList(&m_wndSupplier, m_szSupplierKey);
}
void CPMSSupplierPaymentList::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSSupplierPaymentList::OnStorageSelendok(){
	 
}
long CPMSSupplierPaymentList::OnStorageLoadData(){
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
/*void CPMSSupplierPaymentList::OnStorageAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSSupplierPaymentList::OnItemGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSSupplierPaymentList::OnItemGroupSelendok(){
	 
}
/*void CPMSSupplierPaymentList::OnItemGroupSetfocus(){
	
}*/
/*void CPMSSupplierPaymentList::OnItemGroupKillfocus(){
	
}*/
long CPMSSupplierPaymentList::OnItemGroupLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
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
	int nCount = rs.ExecSQL(szSQL);
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
/*void CPMSSupplierPaymentList::OnItemGroupAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSSupplierPaymentList::OnSourceSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSSupplierPaymentList::OnSourceSelendok(){
	 
}
/*void CPMSSupplierPaymentList::OnSourceSetfocus(){
	
}*/
/*void CPMSSupplierPaymentList::OnSourceKillfocus(){
	
}*/
long CPMSSupplierPaymentList::OnSourceLoadData(){
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
/*void CPMSSupplierPaymentList::OnSourceAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSSupplierPaymentList::OnPrintPreviewSelect(){
	/*
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString szSQL, tmpStr, szOrderID, szNewLine, szOldLine, szDate;
	double nAmt = 0, nGroupTotal = 0;
	long double nTotal = 0;
	int nIdx = 0, nIndex = 0;
	CRecord rs(&pMF->m_db);
	CReportSection* rptHeader, * rptDetail;
	if (!rpt.Init(_T("Reports/HMS/PM_BANGKETHANHTOANCONGNO.RPT")))
		return;

	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	if (!m_szItemGroupKey.IsEmpty())
		rptHeader->SetValue(_T("Type"), m_wndItemGroup.GetCurrent(1));
	if (!m_szSourceKey.IsEmpty())
		rptHeader->SetValue(_T("Source"), m_wndSource.GetCurrent(1));
	if (m_szFromDate.Left(10) != m_szToDate.Left(10))
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	else
		tmpStr.Format(_T("Ngày %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy));
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
		rs.GetValue(_T("partnername"), szNewLine);
		if (szNewLine != szOldLine)
		{
			nIndex++;
			nIdx = 0;
			if (nGroupTotal > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("s6"), double2str(nGroupTotal));
				nTotal += nGroupTotal;
				nGroupTotal = 0;
			}
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
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("invoiceno")));
		rs.GetValue(_T("invoicedate"), tmpStr);
		rptDetail->SetValue(_T("3"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("orderno")));
		rs.GetValue(_T("importdate"), tmpStr);
		rptDetail->SetValue(_T("5"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rs.GetValue(_T("amount"), nAmt);
		nGroupTotal += nAmt;
		rptDetail->SetValue(_T("6"), double2str(nAmt));
		rptDetail->SetValue(_T("7"), _T(""));
		rptDetail->SetValue(_T("8"), _T(""));
		rs.MoveNext();	
	}
	if (nGroupTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("s6"), double2str(nGroupTotal));
		nTotal += nGroupTotal;
		nGroupTotal = 0;
	}
	if (nTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("s6"), ToString(nTotal));
	}	
	rptDetail = rpt.GetReportFooter();
	rptDetail->SetValue(_T("TotalItem"), int2str(nIndex));
	tmpStr.Format(_T("%f"), nTotal);
	rptDetail->SetValue(_T("TotalAmount"), tmpStr);
	if (!szOrderID.IsEmpty())
	{
		szSQL = GetQueryString2();
		rs.ExecSQL(szSQL);
		if (!rs.IsEOF())
		{
			rptDetail->SetValue(_T("s1"), rs.GetValue(_T("thuoc")));
			rptDetail->SetValue(_T("s2"), rs.GetValue(_T("hoachat")));
			rptDetail->SetValue(_T("s3"), rs.GetValue(_T("bbg")));
			rptDetail->SetValue(_T("s4"), rs.GetValue(_T("duoclieu")));
		}
	}
	szDate = pMF->GetSysDate();
	tmpStr.Format(rptDetail->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5, 2), szDate.Left(4));
	rptDetail->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
	*/
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL;
	double nTmp = 0;
	int nIdx = 1, nIdx2 = 0;
	CString szSQL, tmpStr, szReportName, szDate, szOldLev1, szNewLev1, szField, szTemp, szStorages;
	long double nTotal = 0, nGroupTotal1 = 0;
	if (m_szSourceKey.IsEmpty())
	{
		AfxMessageBox(_T("Choose a source!"));
		m_wndSource.SetFocus();
		return;
	}
	if (!rpt.Init(_T("Reports/HMS/PM_BANGKETHANHTOANCONGNO.RPT")))
		return;
	szSQL = GetQueryString1();
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
		rptHeader->SetValue(_T("DeptName"), pMF->GetDepartmentName(pMF->GetDepartmentID()));
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
		if (!m_szSourceKey.IsEmpty())
			rptHeader->SetValue(_T("Source"), m_wndSource.GetCurrent(1));
		else
			rptHeader->SetValue(_T("Source"), _T("To\xE0n \x62\x1ED9"));
	}
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("partnername"), szNewLev1);
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
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(0));
			if (rptGroup)
				rptGroup->SetValue(_T("GroupName"), szNewLev1);
			nIdx2++;
			szOldLev1 = szNewLev1;
		}
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("invoiceno")));
			rs.GetValue(_T("invoicedate"), tmpStr);
			rptDetail->SetValue(_T("3"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));		
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("orderno")));
			rs.GetValue(_T("importdate"), tmpStr);
			rptDetail->SetValue(_T("5"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));		
			rs.GetValue(_T("amount"), nTmp);
			nGroupTotal1+= nTmp;
			rptDetail->SetValue(_T("6"), double2str(nTmp));
		}
		rs.MoveNext();
	}
	if (nGroupTotal1 > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("s6"), double2str(nGroupTotal1));
		nTotal += nGroupTotal1;
	}
	if (nTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("s6"), ToString(nTotal));
	}	
	rptDetail = rpt.GetReportFooter();
	rptDetail->SetValue(_T("TotalItem"), int2str(nIdx2++));
	tmpStr.Format(_T("%f"), nTotal);
	rptDetail->SetValue(_T("TotalAmount"), tmpStr);
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
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.GetReportFooter()->SetValue(_T("Username"), pMF->GetCurrentUser());
	rpt.PrintPreview();
} 
void CPMSSupplierPaymentList::OnExportSelect(){
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
	xls.SetCellText(nCol, nRow + 3, _T("\x44\x41NH S\xC1\x43H TH\x41NH TO\xC1N N\x1EE2 \x43\xC1\x43 \x43\xD4NG TY"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 7);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	//Col Head
	nRow = 5;
	xls.SetCellText(nCol, nRow, _T("STT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+1, nRow, _T("S\x1ED1 h\xF3\x61 \x111\x1A1n"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol+2, nRow, _T("Ng\xE0y h\xF3\x61 \x111\x1A1n"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol+3, nRow, _T("S\x1ED1 phi\x1EBFu nh\x1EADp"), FMT_TEXT | FMT_CENTER , true);
	xls.SetCellText(nCol+4, nRow, _T("Ng\xE0y nh\x1EADp"), FMT_TEXT | FMT_CENTER , true);
	xls.SetCellText(nCol+5, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT | FMT_CENTER , true);
	xls.SetCellText(nCol+6, nRow, _T("S\x1ED1 UNC"), FMT_TEXT | FMT_CENTER , true);
	xls.SetCellText(nCol+7, nRow, _T("K\xFD nh\x1EADn"), FMT_TEXT | FMT_CENTER , true);
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
		if (!szOrderID.IsEmpty())
			szOrderID += _T(", ");
		szOrderID += rs.GetValue(_T("id"));
		xls.SetCellText(nCol + 1, nRow, rs.GetValue(_T("invoiceno")), FMT_TEXT);
		rs.GetValue(_T("invoicedate"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), FMT_TEXT | FMT_CENTER);
		xls.SetCellText(nCol + 3, nRow, rs.GetValue(_T("orderno")), FMT_TEXT);
		rs.GetValue(_T("importdate"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), FMT_TEXT | FMT_CENTER);
		rs.GetValue(_T("amount"), nTmp);
		nGrpAmt += nTmp;
		xls.SetCellText(nCol + 5, nRow, double2str(nTmp), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nGrpAmt > 0)
	{
		xls.SetCellMergedColumns(nCol, nRow, 5);
		xls.SetCellText(nCol, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT, true);
		tmpStr.Format(_T("%.2f"), nGrpAmt);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1);
		nRow++;
		nTotalAmt += nGrpAmt;
	}
	if (nTotalAmt > 0)
	{
		szTemp.Format(_T("%f"), nTotalAmt);
		tmpStr.Format(_T("K\x1EBFt to\xE1n \x62\x1EA3ng n\xE0y g\x1ED3m %d \x63\xF4ng ty v\x1EDBi t\x1ED5ng ti\x1EC1n l\xE0 %s"), nIdx, szTemp);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
	}
	//Detail by group

	//szSQL = GetQueryString2();
	//rs.ExecSQL(szSQL);
	//if (!rs.IsEOF())
	//{
	//	nRow++;
	//	rs.GetValue(_T("thuoc"), nTmp);
	//	szTemp.Format(_T("%ld"), nTmp);
	//	if (szTemp.IsEmpty())
	//		szTemp = _T("0");
	//	tmpStr.Format(_T("Thu\x1ED1\x63: %s \x111\x1ED3ng"), szTemp);
	//	xls.SetCellText(nCol, nRow, tmpStr, FMT_NUMBER1);
	//	nRow++;
	//	rs.GetValue(_T("hoachat"), nTmp);
	//	szTemp.Format(_T("%ld"), nTmp);
	//	if (szTemp.IsEmpty())
	//		szTemp = _T("0");
	//	tmpStr.Format(_T("H\xF3\x61 \x63h\x1EA5t: %s \x111\x1ED3ng"), szTemp);
	//	xls.SetCellText(nCol, nRow, tmpStr, FMT_NUMBER1);
	//	nRow++;
	//	rs.GetValue(_T("bbg"), nTmp);
	//	szTemp.Format(_T("%ld"), nTmp);
	//	if (szTemp.IsEmpty())
	//		szTemp = _T("0");
	//	tmpStr.Format(_T("\x42\xF4ng \x62\x103ng g\x1EA1\x63: %s \x111\x1ED3ng"), szTemp);
	//	xls.SetCellText(nCol, nRow, tmpStr, FMT_NUMBER1);
	//	nRow++;
	//	rs.GetValue(_T("duoclieu"), nTmp);
	//	szTemp.Format(_T("%ld"), nTmp);
	//	if (szTemp.IsEmpty())
	//		szTemp = _T("0");
	//	tmpStr.Format(_T("\x44\x1B0\x1EE3\x63 li\x1EC7u: %s \x111\x1ED3ng"), szTemp);
	//	xls.SetCellText(nCol, nRow, tmpStr, FMT_NUMBER1);
	//}

	EndWaitCursor();
	xls.Save(_T("Exports\\BANG KE THANH TOAN NO CAC CONG TY.xls"));
} 

CString CPMSSupplierPaymentList::GetQueryString1(){
	CString szSQL, szWhere, szStorages;
	szWhere.Format(_T("AND pp_createddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if(!m_szSupplierKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND po_partner_id = '%s'"), m_szSupplierKey);
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
	if (!m_szSourceKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND pp_payment_resource_id = %s"), m_szSourceKey);
	szSQL.Format(_T(" SELECT po_purchase_order_id as id,") \
				_T("   po_invoiceno AS invoiceno,") \
				_T("   Cast_Timestamp2Date(po_invoicedate) AS invoicedate,") \
				_T("   po_orderno AS orderno,") \
				_T("   Cast_Timestamp2Date(po_approveddate) AS importdate,") \
				_T("   sum(pol_totalamount)  AS amount,") \
				_T("   Get_partnername(po_partner_id) AS partnername") \
				_T(" FROM   purchase_order") \
				_T(" LEFT JOIN purchase_orderline ON (po_purchase_order_id = pol_purchase_order_id)") \
				_T(" LEFT JOIN m_productitem_view ON (product_item_id = pol_product_item_id)") \
				_T(" LEFT JOIN purchase_payment ON (po_refpayment_id=pp_purchase_payment_id)") \
				_T(" WHERE  po_payment = 'Y' %s") \
				_T(" GROUP BY po_purchase_order_id, po_invoiceno, po_invoicedate, ") \
				_T("		  po_orderno, po_approveddate, po_partner_id") \
				_T(" ORDER  BY po_partner_id "), szWhere);
	return szSQL;
}

CString CPMSSupplierPaymentList::GetQueryString2(){
	CString szSQL, szWhere, szStorages;
	szWhere.Format(_T("AND pp_createddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if(!m_szSupplierKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND po_partner_id = '%s'"), m_szSupplierKey);
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
	if (!m_szSourceKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND pp_payment_resource_id = %s"), m_szSourceKey);
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
				_T("				purchase_order,") \
				_T("				purchase_payment") \
				_T("         WHERE  pol_product_item_id = product_item_id ") \
				_T("			AND po_purchase_order_id = pol_purchase_order_id") \
				_T("			AND po_refpayment_id=pp_purchase_payment_id AND po_payment = 'Y' %s) "), szWhere);
	return szSQL;
}
BEGIN_MESSAGE_MAP(CPMSSupplierPaymentList, CGuiView)
	ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSSupplierPaymentList::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
