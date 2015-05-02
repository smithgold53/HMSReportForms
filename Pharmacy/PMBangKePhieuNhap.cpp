#include "stdafx.h"
#include "PMBangKePhieuNhap.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKePhieuNhap *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BangKePhieuNhap* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKePhieuNhap *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKePhieuNhap *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKePhieuNhap *)pWnd)->OnToDateCheckValue();
} 
static void _OnSuppelierSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BangKePhieuNhap* )pWnd)->OnSuppelierSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSuppelierSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnSuppelierSelendok();
}
/*static void _OnSuppelierSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnSuppelierKillfocus();
}*/
/*static void _OnSuppelierKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnSuppelierKillfocus();
}*/
static long _OnSuppelierLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKePhieuNhap *)pWnd)->OnSuppelierLoadData();
}
/*static void _OnSuppelierAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnSuppelierAddNew();
}*/
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BangKePhieuNhap* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKePhieuNhap *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnStockAddNew();
}*/
static void _OnTypeSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BangKePhieuNhap* )pWnd)->OnTypeSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnTypeSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnTypeSelendok();
}
/*static void _OnTypeSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnTypeKillfocus();
}*/
/*static void _OnTypeKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnTypeKillfocus();
}*/
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKePhieuNhap *)pWnd)->OnTypeLoadData();
}
/*static void _OnTypeAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnTypeAddNew();
}*/
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BangKePhieuNhap* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BangKePhieuNhap *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BangKePhieuNhap *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CRptPMS108_BangKePhieuNhap *pVw = (CRptPMS108_BangKePhieuNhap *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CRptPMS108_BangKePhieuNhap *pVw = (CRptPMS108_BangKePhieuNhap *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CRptPMS108_BangKePhieuNhap *pVw = (CRptPMS108_BangKePhieuNhap *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddRptPMS108_BangKePhieuNhapFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangKePhieuNhap*)pWnd)->OnAddRptPMS108_BangKePhieuNhap();
} 
static int _OnEditRptPMS108_BangKePhieuNhapFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangKePhieuNhap*)pWnd)->OnEditRptPMS108_BangKePhieuNhap();
} 
static int _OnDeleteRptPMS108_BangKePhieuNhapFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangKePhieuNhap*)pWnd)->OnDeleteRptPMS108_BangKePhieuNhap();
} 
static int _OnSaveRptPMS108_BangKePhieuNhapFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangKePhieuNhap*)pWnd)->OnSaveRptPMS108_BangKePhieuNhap();
} 
static int _OnCancelRptPMS108_BangKePhieuNhapFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangKePhieuNhap*)pWnd)->OnCancelRptPMS108_BangKePhieuNhap();
} 
CRptPMS108_BangKePhieuNhap::CRptPMS108_BangKePhieuNhap(CWnd *pParent)
	{
	m_nDlgWidth = 440;
	m_nDlgHeight = 250;
	SetDefaultValues();
}
CRptPMS108_BangKePhieuNhap::~CRptPMS108_BangKePhieuNhap(){
}
void CRptPMS108_BangKePhieuNhap::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 100, 55);
	m_wndYear.Create(this,105, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 230, 30, 310, 55);
	m_wndReportPeriod.Create(this,315, 30, 425, 55); 
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 210);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 100, 85);
	m_wndFromDate.Create(this,105, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 230, 60, 310, 85);
	m_wndToDate.Create(this,315, 60, 425, 85); 
	m_wndSupplierLabel.Create(this, _T("Supplier"), 10, 90, 100, 115);
	m_wndSuppelier.Create(this,105, 90, 425, 115); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 120, 100, 145);
	m_wndStock.Create(this,105, 120, 425, 145); 
	m_wndTypeLabel.Create(this, _T("Type"), 10, 150, 100, 175);
	m_wndType.Create(this,105, 150, 425, 175); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 180, 100, 205);
	m_wndClerk.Create(this,105, 180, 425, 205); 
	m_wndPrint.Create(this, _T("&Print"), 265, 220, 345, 245);
	m_wndExport.Create(this, _T("&Export"), 350, 220, 430, 245);
}
void CRptPMS108_BangKePhieuNhap::OnInitializeComponents(){
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
//  //m_wndSuppelier.SetCheckValue(true);
	m_wndSuppelier.LimitText(35);
//	//m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);
//	//m_wndType.SetCheckValue(true);
	m_wndType.LimitText(35);
//	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 80);

	m_wndSuppelier.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndSuppelier.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

	m_wndStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStock.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

	m_wndType.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndType.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

}
void CRptPMS108_BangKePhieuNhap::OnSetWindowEvents(){
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
	m_wndSuppelier.SetEvent(WE_SELENDOK, _OnSuppelierSelendokFnc);
	//m_wndSuppelier.SetEvent(WE_SETFOCUS, _OnSuppelierSetfocusFnc);
	//m_wndSuppelier.SetEvent(WE_KILLFOCUS, _OnSuppelierKillfocusFnc);
	m_wndSuppelier.SetEvent(WE_SELCHANGE, _OnSuppelierSelectChangeFnc);
	m_wndSuppelier.SetEvent(WE_LOADDATA, _OnSuppelierLoadDataFnc);
	//m_wndSuppelier.SetEvent(WE_ADDNEW, _OnSuppelierAddNewFnc);
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
void CRptPMS108_BangKePhieuNhap::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndSuppelier.GetDlgCtrlID(), m_szSuppelierKey);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);

}
void CRptPMS108_BangKePhieuNhap::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CRptPMS108_BangKePhieuNhap::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CRptPMS108_BangKePhieuNhap::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szSuppelierKey.Empty();
	m_szStockKey.Empty();
	m_szTypeKey.Empty();
	m_szClerkKey.Empty();

}
int CRptPMS108_BangKePhieuNhap::SetMode(int nMode){
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
/*void CRptPMS108_BangKePhieuNhap::OnYearChange(){
	
} */
/*void CRptPMS108_BangKePhieuNhap::OnYearSetfocus(){
	
} */
/*void CRptPMS108_BangKePhieuNhap::OnYearKillfocus(){
	
} */
int CRptPMS108_BangKePhieuNhap::OnYearCheckValue(){
	return 0;
} 
void CRptPMS108_BangKePhieuNhap::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangKePhieuNhap::OnReportPeriodSelendok(){
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
/*void CRptPMS108_BangKePhieuNhap::OnReportPeriodSetfocus(){
	
}*/
/*void CRptPMS108_BangKePhieuNhap::OnReportPeriodKillfocus(){
	
}*/
long CRptPMS108_BangKePhieuNhap::OnReportPeriodLoadData(){
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
/*void CRptPMS108_BangKePhieuNhap::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CRptPMS108_BangKePhieuNhap::OnFromDateChange(){
	
} */
/*void CRptPMS108_BangKePhieuNhap::OnFromDateSetfocus(){
	
} */
/*void CRptPMS108_BangKePhieuNhap::OnFromDateKillfocus(){
	
} */
int CRptPMS108_BangKePhieuNhap::OnFromDateCheckValue(){
	return 0;
} 
/*void CRptPMS108_BangKePhieuNhap::OnToDateChange(){
	
} */
/*void CRptPMS108_BangKePhieuNhap::OnToDateSetfocus(){
	
} */
/*void CRptPMS108_BangKePhieuNhap::OnToDateKillfocus(){
	
} */
int CRptPMS108_BangKePhieuNhap::OnToDateCheckValue(){
	return 0;
} 
void CRptPMS108_BangKePhieuNhap::OnSuppelierSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangKePhieuNhap::OnSuppelierSelendok(){
	 
}
/*void CRptPMS108_BangKePhieuNhap::OnSuppelierSetfocus(){
	
}*/
/*void CRptPMS108_BangKePhieuNhap::OnSuppelierKillfocus(){
	
}*/
long CRptPMS108_BangKePhieuNhap::OnSuppelierLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndSuppelier.IsSearchKey() && !m_szSuppelierKey.IsEmpty()){
		szWhere.Format(_T(" WHERE pms_id='%s' "), m_szSuppelierKey);
	};
	m_wndSuppelier.DeleteAllItems(); 
		szSQL.Format(_T("SELECT pms_id as id, pms_name as name FROM pms_suppliers %s ORDER BY pms_id"), szWhere);
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
/*void CRptPMS108_BangKePhieuNhap::OnSuppelierAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BangKePhieuNhap::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangKePhieuNhap::OnStockSelendok(){
	 
}
/*void CRptPMS108_BangKePhieuNhap::OnStockSetfocus(){
	
}*/
/*void CRptPMS108_BangKePhieuNhap::OnStockKillfocus(){
	
}*/
long CRptPMS108_BangKePhieuNhap::OnStockLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStock.IsSearchKey() && ToInt(m_szStockKey) > 0){
		szWhere.Format(_T(" AND pmsl_stockid=%d"), ToInt(m_szStockKey));
	};
	m_wndStock.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT pmsl_stockid as id, pmsl_name as name FROM pms_stocklist WHERE pmsl_type='A' %s ORDER BY pmsl_stockid"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStock.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CRptPMS108_BangKePhieuNhap::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BangKePhieuNhap::OnTypeSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangKePhieuNhap::OnTypeSelendok(){
	 
}
/*void CRptPMS108_BangKePhieuNhap::OnTypeSetfocus(){
	
}*/
/*void CRptPMS108_BangKePhieuNhap::OnTypeKillfocus(){
	
}*/
long CRptPMS108_BangKePhieuNhap::OnTypeLoadData(){
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
/*void CRptPMS108_BangKePhieuNhap::OnTypeAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BangKePhieuNhap::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangKePhieuNhap::OnClerkSelendok(){
	 
}
/*void CRptPMS108_BangKePhieuNhap::OnClerkSetfocus(){
	
}*/
/*void CRptPMS108_BangKePhieuNhap::OnClerkKillfocus(){
	
}*/
long CRptPMS108_BangKePhieuNhap::OnClerkLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szSQL.Format(_T("SELECT su_userid as userid, su_name as name FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty()){
	};
	m_wndClerk.DeleteAllItems(); 
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
/*void CRptPMS108_BangKePhieuNhap::OnClerkAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BangKePhieuNhap::OnPrintSelect(){
//CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
//	CReport rpt;
//	CString tmpStr, szSQL,szWhere;
//	CRecord rs(&pMF->m_db);
//	CReportSection* rptDetail;
//	long double nSubAmount=0, nVatAmount=0, nTotalAmount=0, nTotalPayment=0;
//	UpdateData(TRUE);
//	
//	if(!rpt.Init(_T("Reports/HMS/PMS_GENERALPURCHASEORDERREPORT.RPT")) )
//			return;
//		rpt.GetReportHeader()->SetValue(_T("HEAlTHSERVICE"), pMF->m_szHealthService);
//		rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
//		tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
//			CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
//		rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
//		//Page Header
//		//Report Detail
//		
//		if(!m_szSuppelierKey.IsEmpty())
//			szWhere.AppendFormat(_T("and pmpo_suppid='%s' "), m_szSuppelierKey);
//		if(!m_szStockKey.IsEmpty())
//			szWhere.AppendFormat(_T("and pmpo_stockid = %d"), ToInt(m_szStockKey));
//
//		szSQL.Format(_T(" SELECT 	pmpo_invoiceno as invoiceno,") \
//					_T(" 	pmpo_date as invoicedate,") \
//					_T(" 	(select distinct pms_name from pms_suppliers where pms_id=pmpo_suppid) as supplier,") \
//					_T(" 	(select distinct pmsl_name from pms_stocklist where pmsl_stockid=pmpo_stockid) as stockname,") \
//					_T(" 	pmpo_tax1  as vat,") \
//					_T(" 	pmpo_subtotal as subamount,") \
//					_T(" 	pmpo_tax1amt as vatamount,") \
//					_T(" 	pmpo_total as amount, ") \
//					_T(" 	pmpo_payment as payment") \
//					_T(" FROM pms_purchaseinvoice ") \
//					_T(" WHERE pmpo_status <>'O' ")
//					_T(" AND (pmpo_impdate) between ('%s') and ('%s') %s "),
//					m_szFromDate, m_szToDate, szWhere);
//		rs.ExecSQL(szSQL);
//		int nIndex=1;
//		while (!rs.IsEOF())
//		{
//			rptDetail = rpt.AddDetail();
//			tmpStr.Format(_T("%d"), nIndex++);
//			rptDetail->SetValue(_T("1"), tmpStr);
//
//			rs.GetValue(_T("invoiceno"), tmpStr);
//			rptDetail->SetValue(_T("3"), tmpStr);
//
//			rs.GetValue(_T("invoicedate"), tmpStr);
//			rptDetail->SetValue(_T("4"), tmpStr);
//
//			rs.GetValue(_T("supplier"), tmpStr);
//			rptDetail->SetValue(_T("5"), tmpStr);
//
//			rs.GetValue(_T("stockname"), tmpStr);
//			rptDetail->SetValue(_T("6"), tmpStr);
//
//			rs.GetValue(_T("subamount"), tmpStr);
//			nSubAmount+= ToDouble(tmpStr);
//			rptDetail->SetValue(_T("7"), tmpStr);
//
//			rs.GetValue(_T("vat"), tmpStr);
//			rptDetail->SetValue(_T("8"), tmpStr);
//
//			rs.GetValue(_T("vatamount"), tmpStr);
//			nVatAmount+= ToDouble(tmpStr);
//			rptDetail->SetValue(_T("9"), tmpStr);
//
//			rs.GetValue(_T("amount"), tmpStr);
//			nTotalAmount+= ToDouble(tmpStr);
//			rptDetail->SetValue(_T("10"), tmpStr);
//
//			nTotalPayment+= ToDouble(tmpStr);
//			rptDetail->SetValue(_T("11"), tmpStr);
//			rs.MoveNext();
//		}
//		//Page Footer
//		//Report Footer
//		rpt.GetReportFooter()->SetValue(_T("s6"), ToString(nSubAmount));
//
//		rpt.GetReportFooter()->SetValue(_T("s7"), _T(""));
//
//		rpt.GetReportFooter()->SetValue(_T("s8"), ToString(nVatAmount));
//
//		rpt.GetReportFooter()->SetValue(_T("s9"), ToString(nTotalAmount));
//
//		rpt.GetReportFooter()->SetValue(_T("s10"), ToString(nTotalPayment));
//
//		CString szSysDate, szDate;
//		szSysDate = pMF->GetSysDate(); 
//		szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
//		rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
//		rpt.PrintPreview();	
} 
void CRptPMS108_BangKePhieuNhap::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	/*CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
	CReport rpt;
	CString tmpStr, szSQL, szWhere, szFromDateToDate;
	CRecord rs(&pMF->m_db);
	CReportSection* rptDetail;

	long double nSubAmount=0, nVatAmount=0, nTotalAmount=0, nTotalPayment=0;
	double cost=0;;
	UpdateData(TRUE);
	
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);	
	xls.SetColumnWidth(0,8);
	xls.SetColumnWidth(1,15);
	xls.SetColumnWidth(2,12);
	xls.SetColumnWidth(3,12);
	xls.SetColumnWidth(4,12);
	xls.SetColumnWidth(5,12);

	int nCol = 0;
	int nRow = 0;
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellMergedColumns(nCol, nRow + 1, 4);
	xls.SetCellText(nCol, nRow, pMF->m_szHealthService, FMT_TEXT|FMT_CENTER, true);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHospitalName, FMT_TEXT|FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol, nRow + 3, 7);
	TranslateString(_T("\x42\x1EA2NG K\xCA PHI\x1EBEU NH\x1EACP"), tmpStr);
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, true, 18, 0);

	xls.SetCellMergedColumns(nCol, nRow + 4, 7);
	TranslateString(_T("From Date"),tmpStr);
	szFromDateToDate.Format(_T("%s: %s"), tmpStr, CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy));
	TranslateString(_T("To Date"),tmpStr);
	szFromDateToDate.AppendFormat(_T(" %s: %s"), tmpStr, CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));	
	xls.SetCellText(nCol, nRow + 4, szFromDateToDate, FMT_TEXT | FMT_CENTER, true, 12);
	
	xls.SetCellMergedRows(nCol, nRow + 4, 2);
	xls.SetCellText(nCol, nRow + 4, _T("STT"), FMT_TEXT, true);	
	
	xls.SetCellMergedColumns(nCol + 1, nRow + 4, 2); 
	TranslateString(_T("Phi\x1EBFu nh\x1EADp"), tmpStr);	
	xls.SetCellText(nCol + 1, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true);

	TranslateString(_T("InvoinceNo"), tmpStr);	
	xls.SetCellText(nCol + 1, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER, true);	
	TranslateString(_T("Invoince Date"), tmpStr);
	xls.SetCellText(nCol + 2, nRow + 5, tmpStr, FMT_TEXT, true);	

	xls.SetCellMergedRows(nCol + 3, nRow + 4, 2);
	TranslateString(_T("Ngu\x1ED3n h\xE0ng"), tmpStr);
	xls.SetCellText(nCol + 3, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true);	

	TranslateString(_T("Money"), tmpStr);
	xls.SetCellText(nCol + 4, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true);

	TranslateString(_T("Remain Amount"), tmpStr);
	xls.SetCellText(nCol + 5, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true);

	TranslateString(_T("Note"), tmpStr);
	xls.SetCellText(nCol + 6, nRow + 4, tmpStr, FMT_TEXT, true);

	if(!m_szSuppelierKey.IsEmpty())
			szWhere.AppendFormat(_T("and pmpo_suppid='%s' "), m_szSuppelierKey);

	if(!m_szStockKey.IsEmpty())
			szWhere.AppendFormat(_T("and pmpo_stockid = %d"), ToInt(m_szStockKey));

szSQL.Format(_T(" SELECT 	pmpo_invoiceno as invoiceno,") \
			_T(" 	pmpo_date as invoicedate,") \
			_T(" 	(select distinct pms_name from pms_suppliers where pms_id=pmpo_suppid) as supplier,") \
			_T(" 	(select distinct pmsl_name from pms_stocklist where pmsl_stockid=pmpo_stockid) as stockname,") \
			_T(" 	pmpo_tax1  as vat,") \
			_T(" 	pmpo_subtotal as subamount,") \
			_T(" 	pmpo_tax1amt as vatamount,") \
			_T(" 	pmpo_total as amount, ") \
			_T(" 	pmpo_payment as payment") \
			_T(" FROM pms_purchaseinvoice ") \
			_T(" WHERE pmpo_status <>'O' ")
			_T(" AND (pmpo_impdate) between ('%s') and ('%s') %s "), m_szFromDate, m_szToDate, szWhere);
		
		BeginWaitCursor();
		rs.ExecSQL(szSQL);

		int nIndex = 1;
		while (!rs.IsEOF())
		{
			nRow ++;
			tmpStr.Format(_T("%d"), nIndex++);
			xls.SetCellText(0, nRow, tmpStr, FMT_NUMBER1);
			rs.GetValue(_T("invoiceno"), tmpStr);
			xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
			rs.GetValue(_T("invoicedate"), tmpStr);
			xls.SetCellText(2, nRow, CDate::Convert(tmpStr,yyyymmdd,ddmmyyyy), FMT_TEXT);		
			rs.GetValue(_T("supplier"), tmpStr);
			xls.SetCellText(3, nRow, tmpStr, FMT_TEXT);
			rs.GetValue(_T("Stockname"), tmpStr);
			xls.SetCellText(4, nRow, tmpStr, FMT_TEXT);

			rs.GetValue(_T("subamount"), cost);			
			tmpStr.Format(_T("%.2f"),cost);
			nSubAmount+=cost;
			xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER2);
			
			rs.GetValue(_T("vat"), tmpStr);
			xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER2);

			rs.GetValue(_T("vatamount"), cost);
			nTotalPayment+=cost;
			tmpStr.Format(_T("%.2f"),cost);
			xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER2);
			
			rs.GetValue(_T("amount"), cost);
			nTotalAmount+= cost;
			tmpStr.Format(_T("%.2f"),cost);
			xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER2);			
			rs.MoveNext();
		}
		EndWaitCursor();*/
		//xls.Save(_T("Exports\\Bangtonghopphieunhapkho.xls"));

	
} 
void CRptPMS108_BangKePhieuNhap::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CRptPMS108_BangKePhieuNhap::OnAddRptPMS108_BangKePhieuNhap(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CRptPMS108_BangKePhieuNhap::OnEditRptPMS108_BangKePhieuNhap(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CRptPMS108_BangKePhieuNhap::OnDeleteRptPMS108_BangKePhieuNhap(){
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
 		OnCancelRptPMS108_BangKePhieuNhap(); 
 	}
	return 0;
}
int CRptPMS108_BangKePhieuNhap::OnSaveRptPMS108_BangKePhieuNhap(){
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
 		//OnRptPMS108_BangKePhieuNhapListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CRptPMS108_BangKePhieuNhap::OnCancelRptPMS108_BangKePhieuNhap(){
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
int CRptPMS108_BangKePhieuNhap::OnRptPMS108_BangKePhieuNhapListLoadData(){
	return 0;
}
CString CRptPMS108_BangKePhieuNhap::GetQueryString(){
	CString szSQL;
	return szSQL;
}