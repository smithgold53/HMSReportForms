#include "stdafx.h"
#include "PMBangKeNoPhaiThu.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BangKeNoPhaiThu* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BangKeNoPhaiThu* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnStockAddNew();
}*/
static void _OnTypeSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BangKeNoPhaiThu* )pWnd)->OnTypeSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnTypeSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnTypeSelendok();
}
/*static void _OnTypeSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnTypeKillfocus();
}*/
/*static void _OnTypeKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnTypeKillfocus();
}*/
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnTypeLoadData();
}
/*static void _OnTypeAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnTypeAddNew();
}*/
static void _OnSuppelierSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BangKeNoPhaiThu* )pWnd)->OnSuppelierSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSuppelierSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnSuppelierSelendok();
}
/*static void _OnSuppelierSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnSuppelierKillfocus();
}*/
/*static void _OnSuppelierKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnSuppelierKillfocus();
}*/
static long _OnSuppelierLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnSuppelierLoadData();
}
/*static void _OnSuppelierAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnSuppelierAddNew();
}*/
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BangKeNoPhaiThu* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BangKeNoPhaiThu *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CRptPMS108_BangKeNoPhaiThu *pVw = (CRptPMS108_BangKeNoPhaiThu *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CRptPMS108_BangKeNoPhaiThu *pVw = (CRptPMS108_BangKeNoPhaiThu *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CRptPMS108_BangKeNoPhaiThu *pVw = (CRptPMS108_BangKeNoPhaiThu *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddRptPMS108_BangKeNoPhaiThuFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangKeNoPhaiThu*)pWnd)->OnAddRptPMS108_BangKeNoPhaiThu();
} 
static int _OnEditRptPMS108_BangKeNoPhaiThuFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangKeNoPhaiThu*)pWnd)->OnEditRptPMS108_BangKeNoPhaiThu();
} 
static int _OnDeleteRptPMS108_BangKeNoPhaiThuFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangKeNoPhaiThu*)pWnd)->OnDeleteRptPMS108_BangKeNoPhaiThu();
} 
static int _OnSaveRptPMS108_BangKeNoPhaiThuFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangKeNoPhaiThu*)pWnd)->OnSaveRptPMS108_BangKeNoPhaiThu();
} 
static int _OnCancelRptPMS108_BangKeNoPhaiThuFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangKeNoPhaiThu*)pWnd)->OnCancelRptPMS108_BangKeNoPhaiThu();
} 
CRptPMS108_BangKeNoPhaiThu::CRptPMS108_BangKeNoPhaiThu(CWnd *pParent)
	{

	m_nDlgWidth = 420;
	m_nDlgHeight = 245;
	m_szTitle = _T("\x42\x1EA3ng k\xEA n\x1EE3 ph\x1EA3i tr\x1EA3");
	SetDefaultValues();
}
CRptPMS108_BangKeNoPhaiThu::~CRptPMS108_BangKeNoPhaiThu(){
}
void CRptPMS108_BangKeNoPhaiThu::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 415, 210);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 205, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 215, 30, 295, 55);
	m_wndReportPeriod.Create(this,300, 30, 410, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 205, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 215, 60, 295, 85);
	m_wndToDate.Create(this,300, 60, 410, 85); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 90, 90, 115);
	m_wndStock.Create(this,95, 90, 410, 115); 
	m_wndTypeLabel.Create(this, _T("Type"), 10, 120, 90, 145);
	m_wndType.Create(this,95, 120, 410, 145); 
	m_wndSupplierLabel.Create(this, _T("Supplier"), 10, 150, 90, 175);
	m_wndSuppelier.Create(this,95, 150, 410, 175); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 180, 90, 205);
	m_wndClerk.Create(this,95, 180, 410, 205); 
	m_wndPrint.Create(this, _T("&Print"), 250, 215, 330, 240);
	m_wndExport.Create(this, _T("&Export"), 335, 215, 415, 240);

}
void CRptPMS108_BangKeNoPhaiThu::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	//m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);
	//m_wndType.SetCheckValue(true);
	m_wndType.LimitText(35);
	//m_wndSuppelier.SetCheckValue(true);
	m_wndSuppelier.LimitText(35);
	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 80);

	m_wndStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStock.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

	m_wndType.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndType.InsertColumn(1, _T("Description"), CFMT_TEXT, 220);

	m_wndSuppelier.InsertColumn(0, _T("ID"), CFMT_NUMBER, 80);
	m_wndSuppelier.InsertColumn(1, _T("Description"), CFMT_TEXT, 240);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_NUMBER, 80);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 240);
}
void CRptPMS108_BangKeNoPhaiThu::OnSetWindowEvents(){
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
	m_wndStock.SetEvent(WE_SELENDOK, _OnStockSelendokFnc);
	//m_wndStock.SetEvent(WE_SETFOCUS, _OnStockSetfocusFnc);
	//m_wndStock.SetEvent(WE_KILLFOCUS, _OnStockKillfocusFnc);
	m_wndStock.SetEvent(WE_SELCHANGE, _OnStockSelectChangeFnc);
	m_wndStock.SetEvent(WE_LOADDATA, _OnStockLoadDataFnc);
	//m_wndStock.SetEvent(WE_ADDNEW, _OnStockAddNewFnc);
	m_wndType.SetEvent(WE_SELENDOK, _OnTypeSelendokFnc);
	//m_wndType.SetEvent(WE_SETFOCUS, _OnTypeSetfocusFnc);
	//m_wndType.SetEvent(WE_KILLFOCUS, _OnTypeKillfocusFnc);
	m_wndType.SetEvent(WE_SELCHANGE, _OnTypeSelectChangeFnc);
	m_wndType.SetEvent(WE_LOADDATA, _OnTypeLoadDataFnc);
	//m_wndType.SetEvent(WE_ADDNEW, _OnTypeAddNewFnc);
	m_wndSuppelier.SetEvent(WE_SELENDOK, _OnSuppelierSelendokFnc);
	//m_wndSuppelier.SetEvent(WE_SETFOCUS, _OnSuppelierSetfocusFnc);
	//m_wndSuppelier.SetEvent(WE_KILLFOCUS, _OnSuppelierKillfocusFnc);
	m_wndSuppelier.SetEvent(WE_SELCHANGE, _OnSuppelierSelectChangeFnc);
	m_wndSuppelier.SetEvent(WE_LOADDATA, _OnSuppelierLoadDataFnc);
	//m_wndSuppelier.SetEvent(WE_ADDNEW, _OnSuppelierAddNewFnc);
	m_wndClerk.SetEvent(WE_SELENDOK, _OnClerkSelendokFnc);
	//m_wndClerk.SetEvent(WE_SETFOCUS, _OnClerkSetfocusFnc);
	//m_wndClerk.SetEvent(WE_KILLFOCUS, _OnClerkKillfocusFnc);
	m_wndClerk.SetEvent(WE_SELCHANGE, _OnClerkSelectChangeFnc);
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	//m_wndClerk.SetEvent(WE_ADDNEW, _OnClerkAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);

	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate =  pMF->GetSysDate()+ _T("00:00");	
	m_szToDate = pMF->GetSysDate() + _T("23:59");
	SetMode(VM_EDIT);

}
void CRptPMS108_BangKeNoPhaiThu::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_TextEx(pDX, m_wndSuppelier.GetDlgCtrlID(), m_szSuppelierKey);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);

}
void CRptPMS108_BangKeNoPhaiThu::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CRptPMS108_BangKeNoPhaiThu::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CRptPMS108_BangKeNoPhaiThu::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_szTypeKey.Empty();
	m_szSuppelierKey.Empty();
	m_szClerkKey.Empty();

}
int CRptPMS108_BangKeNoPhaiThu::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, 1, 2, 3, -1); 
 			break; 
 		case VM_VIEW: 
 			EnableControls(TRUE); 
 			EnableButtons(FALSE, 3, 4, -1); 
 			break; 
 		case VM_NONE: 
 			EnableControls(TRUE); 
 			EnableButtons(TRUE, 0, 6, -1); 
 			SetDefaultValues(); 
 			break; 
 		}; 
 		UpdateData(FALSE); 
 		return nOldMode; 
}
/*void CRptPMS108_BangKeNoPhaiThu::OnYearChange(){
	
} */
/*void CRptPMS108_BangKeNoPhaiThu::OnYearSetfocus(){
	
} */
/*void CRptPMS108_BangKeNoPhaiThu::OnYearKillfocus(){
	
} */
int CRptPMS108_BangKeNoPhaiThu::OnYearCheckValue(){
	return 0;
} 
void CRptPMS108_BangKeNoPhaiThu::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangKeNoPhaiThu::OnReportPeriodSelendok(){
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
/*void CRptPMS108_BangKeNoPhaiThu::OnReportPeriodSetfocus(){
	
}*/
/*void CRptPMS108_BangKeNoPhaiThu::OnReportPeriodKillfocus(){
	
}*/
long CRptPMS108_BangKeNoPhaiThu::OnReportPeriodLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0){
		szWhere.Format(_T(" WHERE hpr_idx=%ld"), ToInt(m_szReportPeriodKey));
	};
	m_wndReportPeriod.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT * FROM hms_period_report %s ORDER BY hpr_idx "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndReportPeriod.AddItems(
			rs.GetValue(_T("hpr_idx")), 
			rs.GetValue(_T("hpr_name")),			
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CRptPMS108_BangKeNoPhaiThu::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CRptPMS108_BangKeNoPhaiThu::OnFromDateChange(){
	
} */
/*void CRptPMS108_BangKeNoPhaiThu::OnFromDateSetfocus(){
	
} */
/*void CRptPMS108_BangKeNoPhaiThu::OnFromDateKillfocus(){
	
} */
int CRptPMS108_BangKeNoPhaiThu::OnFromDateCheckValue(){
	return 0;
} 
/*void CRptPMS108_BangKeNoPhaiThu::OnToDateChange(){
	
} */
/*void CRptPMS108_BangKeNoPhaiThu::OnToDateSetfocus(){
	
} */
/*void CRptPMS108_BangKeNoPhaiThu::OnToDateKillfocus(){
	
} */
int CRptPMS108_BangKeNoPhaiThu::OnToDateCheckValue(){
	return 0;
} 
void CRptPMS108_BangKeNoPhaiThu::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangKeNoPhaiThu::OnStockSelendok(){
	 
}
/*void CRptPMS108_BangKeNoPhaiThu::OnStockSetfocus(){
	
}*/
/*void CRptPMS108_BangKeNoPhaiThu::OnStockKillfocus(){
	
}*/
long CRptPMS108_BangKeNoPhaiThu::OnStockLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szSQL.Format(_T("SELECT pmsl_stockid as id, pmsl_name as description FROM pms_stocklist ORDER BY pmsl_stockid"));
	if(m_wndStock.IsSearchKey() && !m_szStockKey.IsEmpty()){
	};
	m_wndStock.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStock.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CRptPMS108_BangKeNoPhaiThu::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BangKeNoPhaiThu::OnTypeSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangKeNoPhaiThu::OnTypeSelendok(){
	 
}
/*void CRptPMS108_BangKeNoPhaiThu::OnTypeSetfocus(){
	
}*/
/*void CRptPMS108_BangKeNoPhaiThu::OnTypeKillfocus(){
	
}*/
long CRptPMS108_BangKeNoPhaiThu::OnTypeLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szSQL.Format(_T("SELECT pmdt_id as id, pmdt_name as description FROM pms_drugtype ORDER BY pmdt_index"));
	if(m_wndType.IsSearchKey() && !m_szTypeKey.IsEmpty()){
	};
	m_wndType.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndType.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CRptPMS108_BangKeNoPhaiThu::OnTypeAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BangKeNoPhaiThu::OnSuppelierSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangKeNoPhaiThu::OnSuppelierSelendok(){
	 
}
/*void CRptPMS108_BangKeNoPhaiThu::OnSuppelierSetfocus(){
	
}*/
/*void CRptPMS108_BangKeNoPhaiThu::OnSuppelierKillfocus(){
	
}*/
long CRptPMS108_BangKeNoPhaiThu::OnSuppelierLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szSQL.Format(_T("SELECT pms_id as id, pms_name as name FROM pms_suppliers %s ORDER BY pms_id"), szWhere);
	if(m_wndSuppelier.IsSearchKey() && !m_szSuppelierKey.IsEmpty()){
	};
	m_wndSuppelier.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndSuppelier.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CRptPMS108_BangKeNoPhaiThu::OnSuppelierAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BangKeNoPhaiThu::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangKeNoPhaiThu::OnClerkSelendok(){
	 
}
/*void CRptPMS108_BangKeNoPhaiThu::OnClerkSetfocus(){
	
}*/
/*void CRptPMS108_BangKeNoPhaiThu::OnClerkKillfocus(){
	
}*/
long CRptPMS108_BangKeNoPhaiThu::OnClerkLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty()){
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	};
	m_wndClerk.DeleteAllItems(); 
	szSQL.Format(_T("SELECT su_userid as userid, su_name as name FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("userid")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CRptPMS108_BangKeNoPhaiThu::OnClerkAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BangKeNoPhaiThu::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangKeNoPhaiThu::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangKeNoPhaiThu::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CRptPMS108_BangKeNoPhaiThu::OnAddRptPMS108_BangKeNoPhaiThu(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CRptPMS108_BangKeNoPhaiThu::OnEditRptPMS108_BangKeNoPhaiThu(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CRptPMS108_BangKeNoPhaiThu::OnDeleteRptPMS108_BangKeNoPhaiThu(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
 	CString szSQL; 
 	if(ShowMessage(1, MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2) == IDNO) 
 		return -1; 
 	szSQL.Format(_T("DELETE FROM  WHERE  AND") ); 
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret >= 0){ 
 		SetMode(VM_NONE); 
 		OnCancelRptPMS108_BangKeNoPhaiThu(); 
 	}
	return 0;
}
int CRptPMS108_BangKeNoPhaiThu::OnSaveRptPMS108_BangKeNoPhaiThu(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT) 
 		return -1; 
 	if(!IsValidateData()) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
 	CString szSQL; 
 	if(GetMode() == VM_ADD){ 
 		//szSQL = m_tblTbl.GetInsertSQL(); 
 	} 
 	else if(GetMode() == VM_EDIT){ 
 		CString szWhere; 
 		szWhere.Format(_T(" WHERE")); 
 		//szSQL = m_tblTbl.GetUpdateSQL(_T("createdby,createddate")); 
 		szSQL += szWhere; 
 	} 
 _fmsg(_T("%s"), szSQL); 
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret > 0) 
 	{ 
 		//OnRptPMS108_BangKeNoPhaiThuListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CRptPMS108_BangKeNoPhaiThu::OnCancelRptPMS108_BangKeNoPhaiThu(){
 	if(GetMode() == VM_EDIT) 
 	{ 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 		SetMode(VM_NONE); 
 	} 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
 	return 0;
} 
int CRptPMS108_BangKeNoPhaiThu::OnRptPMS108_BangKeNoPhaiThuListLoadData(){
	return 0;
}
CString CRptPMS108_BangKeNoPhaiThu::GetQueryString(){
	CString szSQL;
	return szSQL;
}