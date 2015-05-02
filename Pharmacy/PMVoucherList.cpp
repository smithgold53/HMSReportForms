#include "stdafx.h"
#include "PMVoucherList.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include <typeinfo.h>
#include ".\pmvoucherlist.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CPMVoucherList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMVoucherList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CPMVoucherList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMVoucherList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMVoucherList *)pWnd)->OnToDateCheckValue();
} 
static long _OnSupplierLoadDataFnc(CWnd *pWnd){
	return ((CPMVoucherList *)pWnd)->OnSupplierLoadData();
}
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMVoucherList* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMVoucherList *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnStorageAddNew();
}*/
static void _OnSourceSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMVoucherList* )pWnd)->OnSourceSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSourceSelendokFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnSourceSelendok();
}
/*static void _OnSourceSetfocusFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnSourceKillfocus();
}*/
/*static void _OnSourceKillfocusFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnSourceKillfocus();
}*/
static long _OnSourceLoadDataFnc(CWnd *pWnd){
	return ((CPMVoucherList *)pWnd)->OnSourceLoadData();
}
/*static void _OnSourceAddNewFnc(CWnd *pWnd){
	((CPMVoucherList *)pWnd)->OnSourceAddNew();
}*/
static long _OnItemGroupLoadDataFnc(CWnd *pWnd){
	return ((CPMVoucherList *)pWnd)->OnItemGroupLoadData();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMVoucherList *pVw = (CPMVoucherList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMVoucherList *pVw = (CPMVoucherList *)pWnd;
	pVw->OnExportSelect();
} 
CPMVoucherList::CPMVoucherList(CWnd *pParent){

	m_nDlgWidth = 605;
	m_nDlgHeight = 332;
	SetDefaultValues();
}
CPMVoucherList::~CPMVoucherList(){
}
void CPMVoucherList::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 180);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 290, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Period"), 295, 30, 375, 55);
	m_wndReportPeriod.Create(this,380, 30, 570, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 290, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 60, 375, 85);
	m_wndToDate.Create(this,380, 60, 570, 85); 
	m_wndSupplierLabel.Create(this, _T("Supplier"), 10, 90, 90, 115);
	m_wndSupplier.Create(this,95, 90, 570, 115); 
	m_wndStorageLabel.Create(this, _T("Storage"), 10, 120, 90, 145);
	m_wndStorage.SetCheckBox(true);
	m_wndStorage.Create(this,95, 120, 570, 145); 
	m_wndItemGroupLabel.Create(this, _T("Item Group"), 10, 150, 90, 175);
	m_wndItemGroup.Create(this,95, 150, 290, 175); 
	m_wndSourceLabel.Create(this, _T("Source"), 295, 150, 375, 175);
	m_wndSource.Create(this,380, 150, 570, 175); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 396, 185, 491, 210);
	m_wndExport.Create(this, _T("&Export"), 495, 185, 575, 210);
}
void CPMVoucherList::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	//m_wndSource.SetCheckValue(true);
	m_wndSource.LimitText(1024);
	//m_wndStorage.SetCheckValue(true);
	m_wndStorage.LimitText(1024);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndReportPeriod.InsertColumn(1, _T("Name"), CFMT_TEXT, 100);

	m_wndSupplier.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndSupplier.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);

	m_wndStorage.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 100);
	m_wndStorage.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);

	m_wndSource.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndSource.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

	m_wndItemGroup.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndItemGroup.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);
}
void CPMVoucherList::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndYear.SetEvent(WE_CHANGE, _OnYearChangeFnc);
	//m_wndYear.SetEvent(WE_SETFOCUS, _OnYearSetfocusFnc);
	//m_wndYear.SetEvent(WE_KILLFOCUS, _OnYearKillfocusFnc);
	m_wndYear.SetEvent(WE_CHECKVALUE, _OnYearCheckValueFnc);
	m_wndReportPeriod.SetEvent(WE_SELENDOK, _OnReportPeriodSelendokFnc);
	//m_wndReportPeriod.SetEvent(WE_SETFOCUS, _OnReportPeriodSetfocusFnc);
	//m_wndReportPeriod.SetEvent(WE_KILLFOCUS, _OnReportPeriodKillfocusFnc);
	m_wndReportPeriod.SetEvent(WE_SELCHANGE, _OnReportPeriodSelectChangeFnc);
	m_wndReportPeriod.SetEvent(WE_LOADDATA, _OnReportPeriodLoadDataFnc);
	//m_wndReportPeriod.SetEvent(WE_ADDNEW, _OnReportPeriodAddNewFnc);
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
	m_wndSource.SetEvent(WE_SELENDOK, _OnSourceSelendokFnc);
	//m_wndSource.SetEvent(WE_SETFOCUS, _OnSourceSetfocusFnc);
	//m_wndSource.SetEvent(WE_KILLFOCUS, _OnSourceKillfocusFnc);
	m_wndSource.SetEvent(WE_SELCHANGE, _OnSourceSelectChangeFnc);
	m_wndSource.SetEvent(WE_LOADDATA, _OnSourceLoadDataFnc);
	//m_wndSource.SetEvent(WE_ADDNEW, _OnSourceAddNewFnc);
	m_wndItemGroup.SetEvent(WE_LOADDATA, _OnItemGroupLoadDataFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	CDate dte;
	dte.ParseDate(pMF->GetSysDate());
	m_nYear = dte.GetYear();
	m_szReportPeriodKey = int2str(dte.GetMonth());
	UpdateData(false);
		
}
void CPMVoucherList::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndSupplier.GetDlgCtrlID(), m_szSupplierKey);
	DDX_TextEx(pDX, m_wndStorage.GetDlgCtrlID(), m_szStorageKey);
	DDX_TextEx(pDX, m_wndSource.GetDlgCtrlID(), m_szSourceKey);
	DDX_TextEx(pDX, m_wndItemGroup.GetDlgCtrlID(), m_szItemGroupKey);
}
void CPMVoucherList::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szSourceKey.Empty();
	m_szStorageKey.Empty();
	m_szItemGroupKey.Empty();
	m_szSupplierKey.Empty();

}
int CPMVoucherList::SetMode(int nMode){
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
/*void CPMVoucherList::OnYearChange(){
	
} */
/*void CPMVoucherList::OnYearSetfocus(){
	
} */
/*void CPMVoucherList::OnYearKillfocus(){
	
} */
int CPMVoucherList::OnYearCheckValue(){
	return 0;
} 
void CPMVoucherList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
} 
void CPMVoucherList::OnReportPeriodSelendok(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString tmpStr;
	CDate dte;
	UpdateData(true);
	int nMonth = ToInt(m_szReportPeriodKey);
	if(nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/01"), m_nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), m_nYear, nMonth, dte.GetMonthLastDay());
	}
	if(nMonth == 13){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/03/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59" ), m_nYear,  dte.GetMonthLastDay());
	}
	if(nMonth == 14){
		m_szFromDate.Format(_T("%.4d/04/01"), m_nYear);
		tmpStr.Format(_T("%.4d/06/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 15){
		m_szFromDate.Format(_T("%.4d/07/01"), m_nYear);
		tmpStr.Format(_T("%.4d/09/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 16){
		m_szFromDate.Format(_T("%.4d/10/01"), m_nYear);
		tmpStr.Format(_T("%.4d/10/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 17){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/12/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());		
	}	
	UpdateData(false);	 
}
/*void CPMVoucherList::OnReportPeriodSetfocus(){
	
}*/
/*void CPMVoucherList::OnReportPeriodKillfocus(){
	
}*/
long CPMVoucherList::OnReportPeriodLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadReportPeriodList(&m_wndReportPeriod, m_szReportPeriodKey);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && !m_szReportPeriodKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szReportPeriodKey
};
	m_wndReportPeriod.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndReportPeriod.AddItems(
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CPMVoucherList::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CPMVoucherList::OnFromDateChange(){
	
} */
/*void CPMVoucherList::OnFromDateSetfocus(){
	
} */
/*void CPMVoucherList::OnFromDateKillfocus(){
	
} */
int CPMVoucherList::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMVoucherList::OnToDateChange(){
	
} */
/*void CPMVoucherList::OnToDateSetfocus(){
	
} */
/*void CPMVoucherList::OnToDateKillfocus(){
	
} */
int CPMVoucherList::OnToDateCheckValue(){
	return 0;
} 
long CPMVoucherList::OnSupplierLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadPartnerList(&m_wndSupplier, m_szSupplierKey);
}
void CPMVoucherList::OnSourceSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMVoucherList::OnSourceSelendok(){
	 
}
/*void CPMVoucherList::OnSourceSetfocus(){
	
}*/
/*void CPMVoucherList::OnSourceKillfocus(){
	
}*/
long CPMVoucherList::OnSourceLoadData(){
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
/*void CPMVoucherList::OnSourceAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMVoucherList::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMVoucherList::OnStorageSelendok(){
	 
}
/*void CPMVoucherList::OnStorageSetfocus(){
	
}*/
/*void CPMVoucherList::OnStorageKillfocus(){
	
}*/
long CPMVoucherList::OnStorageLoadData(){
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
/*void CPMVoucherList::OnStorageAddNew(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} */
long CPMVoucherList::OnItemGroupLoadData(){
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
void CPMVoucherList::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString szSQL, tmpStr, szMoney, szDate, szOldLev1, szNewLev1, szStorages;
	double nAmt = 0;
	long double nTotal = 0, nGroupTotal1 = 0;
	int nIdx = 0, nIdx2 = 0;
	CRecord rs(&pMF->m_db);
	CReportSection *rptHeader, *rptDetail, *rptFooter, *rptGroup;
	if (m_szSourceKey.IsEmpty())
	{
		AfxMessageBox(_T("Choose a source!"));
		m_wndSource.SetFocus();
		return;
	}
	if (!rpt.Init(_T("Reports\\HMS\\PM_BANGKECHUNGTUCHITIEU.RPT")))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		_fmsg(_T("%s"), szSQL);
		return;
	}
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	rptHeader->SetValue(_T("Department"), rs.GetValue(_T("dept")));

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
	else
		rptHeader->SetValue(_T("Source"), _T("To\xE0n \x62\x1ED9"));
	if (!m_szItemGroupKey.IsEmpty())
		rptHeader->SetValue(_T("ItemType"), m_wndItemGroup.GetCurrent(1));
	else
		rptHeader->SetValue(_T("ItemType"), _T("To\xE0n \x62\x1ED9"));
	if (!m_szSupplierKey.IsEmpty())
		rptHeader->SetValue(_T("Supplier"), m_wndSupplier.GetCurrent(1));
	else
		rptHeader->SetValue(_T("Supplier"), _T("To\xE0n \x62\x1ED9"));
	if (!m_szReportPeriodKey.IsEmpty())
		tmpStr.Format(_T("%s n\x103m %d"), m_wndReportPeriod.GetCurrent(1), m_nYear);
	rptHeader->SetValue(_T("Period"), tmpStr);
	tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("partnername"), szNewLev1);
		if(szNewLev1 != szOldLev1)
		{
			if(nGroupTotal1 > 0)
			{
				tmpStr.Format(_T("%f"), nGroupTotal1);
				rptGroup = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptGroup->SetValue(_T("s5"), tmpStr);
				nTotal += nGroupTotal1;
				nGroupTotal1 = 0;
			}
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(0));
			if(rptGroup)
				rptGroup->SetValue(_T("GroupName"), szNewLev1);
			nIdx2++;
			szOldLev1 = szNewLev1;
		}
		rptDetail = rpt.AddDetail();
		nIdx++;
		rptDetail->SetValue(_T("1"), int2str(nIdx));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("invoiceno")));
		rs.GetValue(_T("invoicedate"), tmpStr);
		rptDetail->SetValue(_T("3"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		//rs.GetValue(_T("importdate"), tmpStr);
		//rptDetail->SetValue(_T("4"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("productname")));
		rs.GetValue(_T("total"), nAmt);
		nGroupTotal1 += nAmt;
		rptDetail->SetValue(_T("5"), double2str(nAmt));
		rs.MoveNext();
	}
	if (nGroupTotal1 > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("s5"), double2str(nGroupTotal1));
		nTotal += nGroupTotal1;
	}
	if (nTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("s5"), ToString(nTotal));
	}
	rptFooter = rpt.GetReportFooter();
	FormatCurrency(nTotal, szMoney);
	tmpStr.Format(rptFooter->GetValue(_T("Statistic")), nIdx2, szMoney);
	rptFooter->SetValue(_T("Statistic"), tmpStr);
	szDate = pMF->GetSysDate();
	tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5, 2), szDate.Left(4));
	rptFooter->SetValue(_T("PrintDate"), tmpStr);
	rptFooter->SetValue(_T("Username"), pMF->GetCurrentUser());
	rpt.PrintPreview();
} 
void CPMVoucherList::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr, szTemp;
	CRecord rs(&pMF->m_db);
	int nIdx = 1;
	double nTmp = 0;
	long double nTotalAmt = 0;
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
	xls.SetColumnWidth(1, 15);
	xls.SetColumnWidth(2, 12);
	xls.SetColumnWidth(3, 40);
	xls.SetColumnWidth(4, 15);
	//xls.SetColumnWidth(5, 15);
	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 7);
	xls.SetCellText(nCol, nRow + 3, _T("\x42\x1EA2NG K\xCA \x43H\x1EE8NG T\x1EEA \x43HI TI\xCAU"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 7);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	//Col Head
	nRow = 5;
	xls.SetCellText(nCol, nRow, _T("STT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true);
	xls.SetCellText(nCol+1, nRow, _T("S\x1ED1 h\xF3\x61 \x111\x1A1n"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol+2, nRow, _T("Ng\xE0y h\xF3\x61 \x111\x1A1n"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol+3, nRow, _T("N\x1ED9i \x64ung th\x61nh to\xE1n"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol+4, nRow, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT | FMT_CENTER , true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		nIdx++;
		xls.SetCellText(nCol, nRow, int2str(nIdx), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(nCol + 1, nRow, rs.GetValue(_T("invoiceno")), FMT_TEXT);
		rs.GetValue(_T("invoicedate"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), FMT_TEXT | FMT_CENTER);
		xls.SetCellText(nCol + 3, nRow, rs.GetValue(_T("productname")), FMT_TEXT);
		rs.GetValue(_T("total"), nTmp);
		nTotalAmt += nTmp;
		xls.SetCellText(nCol + 4, nRow, double2str(nTmp), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nTotalAmt > 0)
	{
		xls.SetCellMergedColumns(nCol, nRow, 4);
		xls.SetCellText(nCol, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
		tmpStr.Format(_T("%.2f"), nTotalAmt);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1);
		nRow++;
		FormatCurrency(nTotalAmt, szTemp);
		tmpStr.Format(_T("K\x1EBFt to\xE1n \x62\x1EA3ng n\xE0y g\x1ED3m %d \x63\xF4ng ty v\x1EDBi t\x1ED5ng ti\x1EC1n l\xE0 %s"), nIdx, szTemp);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\BANG KE CHUNG TU CHI TIEU.xls"));
} 

CString CPMVoucherList::GetQueryString(){
	CString szSQL, szWhere, szStorages;
	szWhere.Format(_T(" AND po_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
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
				_T("   DISTINCT") \
				_T("   po_purchase_order_id as orderid,") \
				_T("   po_invoiceno as invoiceno,") \
				_T("   po_partner_id as id,") \
				_T("   get_partnername(po_partner_id) as partnername,") \
				_T("   po_invoicedate as invoicedate,") \
				_T("   po_approveddate                AS importdate,") \
				_T("   First_value(product_name)") \
				_T("     OVER (") \
				_T("       partition BY po_purchase_order_id") \
				_T("       ORDER BY product_name ASC) AS productname,") \
				_T("   (select get_departmentname(msl_dept_id) from m_storagelist where msl_storage_id = po_storage_id) as dept,") \
				_T("   po_totalamount as total") \
				_T(" FROM   purchase_order") \
				_T(" LEFT JOIN purchase_orderline ON (po_purchase_order_id=pol_purchase_order_id)") \
				_T(" LEFT JOIN m_productitem_view ON (product_item_id=pol_product_item_id)") \
				_T(" LEFT JOIN purchase_payment ON (po_refpayment_id=pp_purchase_payment_id)") \
				_T(" WHERE  po_status='A' AND product_org_id = 'PM' %s") \
				_T(" ORDER  BY po_partner_id"), szWhere);
	return szSQL;
}BEGIN_MESSAGE_MAP(CPMVoucherList, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMVoucherList::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
