#include "FMDischargeDepositUnpaidReport.h"
#include "MainFrm.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMDischargeDepositUnpaidReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMDischargeDepositUnpaidReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMDischargeDepositUnpaidReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMDischargeDepositUnpaidReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMDischargeDepositUnpaidReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMDischargeDepositUnpaidReport *pVw = (CFMDischargeDepositUnpaidReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMDischargeDepositUnpaidReport *pVw = (CFMDischargeDepositUnpaidReport *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnObjectListLoadDataFnc(CWnd *pWnd){
	return ((CFMDischargeDepositUnpaidReport*)pWnd)->OnObjectListLoadData();
} 
static void _OnObjectListDblClickFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport*)pWnd)->OnObjectListDblClick();
} 
static void _OnObjectListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMDischargeDepositUnpaidReport*)pWnd)->OnObjectListSelectChange(nOldItem, nNewItem);
} 
static int _OnObjectListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMDischargeDepositUnpaidReport*)pWnd)->OnObjectListDeleteItem();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMDischargeDepositUnpaidReport*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMDischargeDepositUnpaidReport*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMDischargeDepositUnpaidReport*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMDischargeDepositUnpaidReport*)pWnd)->OnDeptListDeleteItem();
} 
CFMDischargeDepositUnpaidReport::CFMDischargeDepositUnpaidReport(){

	m_nDlgWidth = 585;
	m_nDlgHeight = 585;
	SetDefaultValues();
}
CFMDischargeDepositUnpaidReport::~CFMDischargeDepositUnpaidReport(){
}
void CFMDischargeDepositUnpaidReport::OnCreateComponents(){
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
void CFMDischargeDepositUnpaidReport::OnInitializeComponents(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);







}
void CFMDischargeDepositUnpaidReport::OnSetWindowEvents(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
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
void CFMDischargeDepositUnpaidReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CFMDischargeDepositUnpaidReport::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CFMDischargeDepositUnpaidReport::SetMode(int nMode){
 		int nOldMode = GetMode();
 		CGuiView::SetMode(nMode);
 		CMainFrame *pMF = (CMainFrame *) AfxGetMainWnd();
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
/*void CFMDischargeDepositUnpaidReport::OnYearChange(){
	
} */
/*void CFMDischargeDepositUnpaidReport::OnYearSetfocus(){
	
} */
/*void CFMDischargeDepositUnpaidReport::OnYearKillfocus(){
	
} */
int CFMDischargeDepositUnpaidReport::OnYearCheckValue(){
	return 0;
} 
void CFMDischargeDepositUnpaidReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
void CFMDischargeDepositUnpaidReport::OnReportPeriodSelendok(){
	 
}
/*void CFMDischargeDepositUnpaidReport::OnReportPeriodSetfocus(){
	
}*/
/*void CFMDischargeDepositUnpaidReport::OnReportPeriodKillfocus(){
	
}*/
long CFMDischargeDepositUnpaidReport::OnReportPeriodLoadData(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
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
/*void CFMDischargeDepositUnpaidReport::OnReportPeriodAddNew(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMDischargeDepositUnpaidReport::OnFromDateChange(){
	
} */
/*void CFMDischargeDepositUnpaidReport::OnFromDateSetfocus(){
	
} */
/*void CFMDischargeDepositUnpaidReport::OnFromDateKillfocus(){
	
} */
int CFMDischargeDepositUnpaidReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMDischargeDepositUnpaidReport::OnToDateChange(){
	
} */
/*void CFMDischargeDepositUnpaidReport::OnToDateSetfocus(){
	
} */
/*void CFMDischargeDepositUnpaidReport::OnToDateKillfocus(){
	
} */
int CFMDischargeDepositUnpaidReport::OnToDateCheckValue(){
	return 0;
} 
void CFMDischargeDepositUnpaidReport::OnPrintSelect(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
void CFMDischargeDepositUnpaidReport::OnExportSelect(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
void CFMDischargeDepositUnpaidReport::OnObjectListDblClick(){
	
} 
void CFMDischargeDepositUnpaidReport::OnObjectListSelectChange(int nOldItem, int nNewItem){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
int CFMDischargeDepositUnpaidReport::OnObjectListDeleteItem(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMDischargeDepositUnpaidReport::OnObjectListLoadData(){
/*
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
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
void CFMDischargeDepositUnpaidReport::OnDeptListDblClick(){
	
} 
void CFMDischargeDepositUnpaidReport::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} 
int CFMDischargeDepositUnpaidReport::OnDeptListDeleteItem(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMDischargeDepositUnpaidReport::OnDeptListLoadData(){
/*
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
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
