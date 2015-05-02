#include "stdafx.h"
#include "FMUndischargedAdvancePatientList.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMUndischargedAdvancePatientList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMUndischargedAdvancePatientList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMUndischargedAdvancePatientList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMUndischargedAdvancePatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMUndischargedAdvancePatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMUndischargedAdvancePatientList *pVw = (CFMUndischargedAdvancePatientList *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMUndischargedAdvancePatientList *pVw = (CFMUndischargedAdvancePatientList *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnObjectListLoadDataFnc(CWnd *pWnd){
	return ((CFMUndischargedAdvancePatientList*)pWnd)->OnObjectListLoadData();
} 
static void _OnObjectListDblClickFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList*)pWnd)->OnObjectListDblClick();
} 
static void _OnObjectListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMUndischargedAdvancePatientList*)pWnd)->OnObjectListSelectChange(nOldItem, nNewItem);
} 
static int _OnObjectListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMUndischargedAdvancePatientList*)pWnd)->OnObjectListDeleteItem();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMUndischargedAdvancePatientList*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMUndischargedAdvancePatientList*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMUndischargedAdvancePatientList*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMUndischargedAdvancePatientList*)pWnd)->OnDeptListDeleteItem();
} 
CFMUndischargedAdvancePatientList::CFMUndischargedAdvancePatientList(CWnd *pParent){

	m_nDlgWidth = 585;
	m_nDlgHeight = 585;
	SetDefaultValues();
}
CFMUndischargedAdvancePatientList::~CFMUndischargedAdvancePatientList(){
}
void CFMUndischargedAdvancePatientList::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 545);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 280, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 285, 30, 375, 55);
	m_wndReportPeriod.Create(this,380, 30, 570, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 280, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 285, 60, 375, 85);
	m_wndToDate.Create(this,380, 60, 570, 85); 
	m_wndPrint.Create(this, _T("&Print"), 410, 550, 490, 575);
	m_wndExport.Create(this, _T("&Export"), 495, 550, 575, 575);
	m_wndObjectList.Create(this,10, 90, 570, 310); 
	m_wndDeptList.Create(this,10, 315, 570, 540); 

}
void CFMUndischargedAdvancePatientList::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);







}
void CFMUndischargedAdvancePatientList::OnSetWindowEvents(){
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
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndObjectList.SetEvent(WE_SELCHANGE, _OnObjectListSelectChangeFnc);
	m_wndObjectList.SetEvent(WE_LOADDATA, _OnObjectListLoadDataFnc);
	m_wndObjectList.SetEvent(WE_DBLCLICK, _OnObjectListDblClickFnc);
	m_wndObjectList.AddEvent(1, _T("Delete"), _OnObjectListDeleteItemFnc, 0, VK_DELETE, 0);
	m_wndDeptList.SetEvent(WE_SELCHANGE, _OnDeptListSelectChangeFnc);
	m_wndDeptList.SetEvent(WE_LOADDATA, _OnDeptListLoadDataFnc);
	m_wndDeptList.SetEvent(WE_DBLCLICK, _OnDeptListDblClickFnc);
	m_wndDeptList.AddEvent(1, _T("Delete"), _OnDeptListDeleteItemFnc, 0, VK_DELETE, 0);
	SetMode(VM_NONE);

}
void CFMUndischargedAdvancePatientList::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CFMUndischargedAdvancePatientList::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CFMUndischargedAdvancePatientList::SetMode(int nMode){
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
/*void CFMUndischargedAdvancePatientList::OnYearChange(){
	
} */
/*void CFMUndischargedAdvancePatientList::OnYearSetfocus(){
	
} */
/*void CFMUndischargedAdvancePatientList::OnYearKillfocus(){
	
} */
int CFMUndischargedAdvancePatientList::OnYearCheckValue(){
	return 0;
} 
void CFMUndischargedAdvancePatientList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CFMUndischargedAdvancePatientList::OnReportPeriodSelendok(){
	 
}
/*void CFMUndischargedAdvancePatientList::OnReportPeriodSetfocus(){
	
}*/
/*void CFMUndischargedAdvancePatientList::OnReportPeriodKillfocus(){
	
}*/
long CFMUndischargedAdvancePatientList::OnReportPeriodLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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
/*void CFMUndischargedAdvancePatientList::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CFMUndischargedAdvancePatientList::OnFromDateChange(){
	
} */
/*void CFMUndischargedAdvancePatientList::OnFromDateSetfocus(){
	
} */
/*void CFMUndischargedAdvancePatientList::OnFromDateKillfocus(){
	
} */
int CFMUndischargedAdvancePatientList::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMUndischargedAdvancePatientList::OnToDateChange(){
	
} */
/*void CFMUndischargedAdvancePatientList::OnToDateSetfocus(){
	
} */
/*void CFMUndischargedAdvancePatientList::OnToDateKillfocus(){
	
} */
int CFMUndischargedAdvancePatientList::OnToDateCheckValue(){
	return 0;
} 
void CFMUndischargedAdvancePatientList::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CFMUndischargedAdvancePatientList::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CFMUndischargedAdvancePatientList::OnObjectListDblClick(){
	
} 
void CFMUndischargedAdvancePatientList::OnObjectListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CFMUndischargedAdvancePatientList::OnObjectListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
} 
long CFMUndischargedAdvancePatientList::OnObjectListLoadData(){
/*
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndObjectList.BeginLoad(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObjectList.AddItems(
		rs.MoveNext();
	}
	m_wndObjectList.EndLoad(); 
	return nCount;
*/
	return 0;
}
void CFMUndischargedAdvancePatientList::OnDeptListDblClick(){
	
} 
void CFMUndischargedAdvancePatientList::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CFMUndischargedAdvancePatientList::OnDeptListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
} 
long CFMUndischargedAdvancePatientList::OnDeptListLoadData(){
/*
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndDeptList.BeginLoad(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDeptList.AddItems(
		rs.MoveNext();
	}
	m_wndDeptList.EndLoad(); 
	return nCount;
*/
	return 0;
}
