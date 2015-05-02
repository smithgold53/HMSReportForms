#include "stdafx.h"
#include "PMBaoCaoKiemKeKho.h"
#include "MainFrame_E10.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BaoCaoKiemKeKho* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BaoCaoKiemKeKho* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnStockAddNew();
}*/
static void _OnTypeSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BaoCaoKiemKeKho* )pWnd)->OnTypeSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnTypeSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnTypeSelendok();
}
/*static void _OnTypeSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnTypeKillfocus();
}*/
/*static void _OnTypeKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnTypeKillfocus();
}*/
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnTypeLoadData();
}
/*static void _OnTypeAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnTypeAddNew();
}*/
static void _OnGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BaoCaoKiemKeKho* )pWnd)->OnGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnGroupSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnGroupSelendok();
}
/*static void _OnGroupSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnGroupKillfocus();
}*/
/*static void _OnGroupKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnGroupKillfocus();
}*/
static long _OnGroupLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnGroupLoadData();
}
/*static void _OnGroupAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnGroupAddNew();
}*/
static void _OnSourceSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BaoCaoKiemKeKho* )pWnd)->OnSourceSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSourceSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnSourceSelendok();
}
/*static void _OnSourceSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnSourceKillfocus();
}*/
/*static void _OnSourceKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnSourceKillfocus();
}*/
static long _OnSourceLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnSourceLoadData();
}
/*static void _OnSourceAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BaoCaoKiemKeKho *)pWnd)->OnSourceAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CRptPMS108_BaoCaoKiemKeKho *pVw = (CRptPMS108_BaoCaoKiemKeKho *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CRptPMS108_BaoCaoKiemKeKho *pVw = (CRptPMS108_BaoCaoKiemKeKho *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CRptPMS108_BaoCaoKiemKeKho *pVw = (CRptPMS108_BaoCaoKiemKeKho *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CRptPMS108_BaoCaoKiemKeKho *pVw = (CRptPMS108_BaoCaoKiemKeKho *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddRptPMS108_BaoCaoKiemKeKhoFnc(CWnd *pWnd){
	 return ((CRptPMS108_BaoCaoKiemKeKho*)pWnd)->OnAddRptPMS108_BaoCaoKiemKeKho();
} 
static int _OnEditRptPMS108_BaoCaoKiemKeKhoFnc(CWnd *pWnd){
	 return ((CRptPMS108_BaoCaoKiemKeKho*)pWnd)->OnEditRptPMS108_BaoCaoKiemKeKho();
} 
static int _OnDeleteRptPMS108_BaoCaoKiemKeKhoFnc(CWnd *pWnd){
	 return ((CRptPMS108_BaoCaoKiemKeKho*)pWnd)->OnDeleteRptPMS108_BaoCaoKiemKeKho();
} 
static int _OnSaveRptPMS108_BaoCaoKiemKeKhoFnc(CWnd *pWnd){
	 return ((CRptPMS108_BaoCaoKiemKeKho*)pWnd)->OnSaveRptPMS108_BaoCaoKiemKeKho();
} 
static int _OnCancelRptPMS108_BaoCaoKiemKeKhoFnc(CWnd *pWnd){
	 return ((CRptPMS108_BaoCaoKiemKeKho*)pWnd)->OnCancelRptPMS108_BaoCaoKiemKeKho();
} 
CRptPMS108_BaoCaoKiemKeKho::CRptPMS108_BaoCaoKiemKeKho(CWnd *pParent)
{

	m_nDlgWidth = 445;
	m_nDlgHeight = 245;
	SetDefaultValues();
}
CRptPMS108_BaoCaoKiemKeKho::~CRptPMS108_BaoCaoKiemKeKho(){
}
void CRptPMS108_BaoCaoKiemKeKho::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 110, 55);
	m_wndYear.Create(this,115, 30, 230, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 235, 30, 315, 55);
	m_wndReportPeriod.Create(this,320, 30, 435, 55); 
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 440, 210);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 110, 85);
	m_wndFromDate.Create(this,115, 60, 230, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 235, 60, 315, 85);
	m_wndToDate.Create(this,320, 60, 435, 85); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 90, 110, 115);
	m_wndStock.Create(this,115, 90, 435, 115); 
	m_wndTypeLabel.Create(this, _T("Type"), 10, 120, 110, 145);
	m_wndType.Create(this,115, 120, 435, 145); 
	m_wndGroupLabel.Create(this, _T("Group"), 10, 150, 110, 175);
	m_wndGroup.Create(this,115, 150, 435, 175); 
	m_wndSourceLabel.Create(this, _T("Source"), 10, 180, 110, 205);
	m_wndSource.Create(this,115, 180, 435, 205); 
	//m_wndPrintPreview.Create(this, _T("Print Pre&view"), 105, 215, 185, 240);
	m_wndPrint.Create(this, _T("&Print"), 275, 215, 355, 240);
	m_wndExport.Create(this, _T("&Export"), 360, 215, 440, 240);

}
void CRptPMS108_BaoCaoKiemKeKho::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	//m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	//m_wndToDate.SetCheckValue(true);
	//m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);
	//m_wndType.SetCheckValue(true);
	m_wndType.LimitText(35);
	//m_wndGroup.SetCheckValue(true);
	m_wndGroup.LimitText(35);
	//m_wndSource.SetCheckValue(true);
	m_wndSource.LimitText(35);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 80);


	m_wndStock.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStock.InsertColumn(1, _T("Description"), CFMT_TEXT, 220);


	m_wndType.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndType.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);


	m_wndGroup.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	m_wndGroup.InsertColumn(1, _T("Description"), CFMT_TEXT, 240);


	m_wndSource.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndSource.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

}
void CRptPMS108_BaoCaoKiemKeKho::OnSetWindowEvents(){
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
	m_wndGroup.SetEvent(WE_SELENDOK, _OnGroupSelendokFnc);
	//m_wndGroup.SetEvent(WE_SETFOCUS, _OnGroupSetfocusFnc);
	//m_wndGroup.SetEvent(WE_KILLFOCUS, _OnGroupKillfocusFnc);
	m_wndGroup.SetEvent(WE_SELCHANGE, _OnGroupSelectChangeFnc);
	m_wndGroup.SetEvent(WE_LOADDATA, _OnGroupLoadDataFnc);
	//m_wndGroup.SetEvent(WE_ADDNEW, _OnGroupAddNewFnc);
	m_wndSource.SetEvent(WE_SELENDOK, _OnSourceSelendokFnc);
	//m_wndSource.SetEvent(WE_SETFOCUS, _OnSourceSetfocusFnc);
	//m_wndSource.SetEvent(WE_KILLFOCUS, _OnSourceKillfocusFnc);
	m_wndSource.SetEvent(WE_SELCHANGE, _OnSourceSelectChangeFnc);
	m_wndSource.SetEvent(WE_LOADDATA, _OnSourceLoadDataFnc);
	//m_wndSource.SetEvent(WE_ADDNEW, _OnSourceAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
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
void CRptPMS108_BaoCaoKiemKeKho::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_TextEx(pDX, m_wndGroup.GetDlgCtrlID(), m_szGroupKey);
	DDX_TextEx(pDX, m_wndSource.GetDlgCtrlID(), m_szSourceKey);

}
void CRptPMS108_BaoCaoKiemKeKho::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CRptPMS108_BaoCaoKiemKeKho::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CRptPMS108_BaoCaoKiemKeKho::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_szTypeKey.Empty();
	m_szGroupKey.Empty();
	m_szSourceKey.Empty();

}
int CRptPMS108_BaoCaoKiemKeKho::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, 1, -1); 
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
/*void CRptPMS108_BaoCaoKiemKeKho::OnYearChange(){
	
} */
/*void CRptPMS108_BaoCaoKiemKeKho::OnYearSetfocus(){
	
} */
/*void CRptPMS108_BaoCaoKiemKeKho::OnYearKillfocus(){
	
} */
int CRptPMS108_BaoCaoKiemKeKho::OnYearCheckValue(){
	return 0;
} 
void CRptPMS108_BaoCaoKiemKeKho::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BaoCaoKiemKeKho::OnReportPeriodSelendok(){
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
/*void CRptPMS108_BaoCaoKiemKeKho::OnReportPeriodSetfocus(){
	
}*/
/*void CRptPMS108_BaoCaoKiemKeKho::OnReportPeriodKillfocus(){
	
}*/
long CRptPMS108_BaoCaoKiemKeKho::OnReportPeriodLoadData(){
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
/*void CRptPMS108_BaoCaoKiemKeKho::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CRptPMS108_BaoCaoKiemKeKho::OnFromDateChange(){
	
} */
/*void CRptPMS108_BaoCaoKiemKeKho::OnFromDateSetfocus(){
	
} */
/*void CRptPMS108_BaoCaoKiemKeKho::OnFromDateKillfocus(){
	
} */
int CRptPMS108_BaoCaoKiemKeKho::OnFromDateCheckValue(){
	return 0;
} 
/*void CRptPMS108_BaoCaoKiemKeKho::OnToDateChange(){
	
} */
/*void CRptPMS108_BaoCaoKiemKeKho::OnToDateSetfocus(){
	
} */
/*void CRptPMS108_BaoCaoKiemKeKho::OnToDateKillfocus(){
	
} */
int CRptPMS108_BaoCaoKiemKeKho::OnToDateCheckValue(){
	return 0;
} 
void CRptPMS108_BaoCaoKiemKeKho::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BaoCaoKiemKeKho::OnStockSelendok(){
	 
}
/*void CRptPMS108_BaoCaoKiemKeKho::OnStockSetfocus(){
	
}*/
/*void CRptPMS108_BaoCaoKiemKeKho::OnStockKillfocus(){
	
}*/
long CRptPMS108_BaoCaoKiemKeKho::OnStockLoadData(){
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
/*void CRptPMS108_BaoCaoKiemKeKho::OnStockAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BaoCaoKiemKeKho::OnTypeSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BaoCaoKiemKeKho::OnTypeSelendok(){
	 
}
/*void CRptPMS108_BaoCaoKiemKeKho::OnTypeSetfocus(){
	
}*/
/*void CRptPMS108_BaoCaoKiemKeKho::OnTypeKillfocus(){
	
}*/
long CRptPMS108_BaoCaoKiemKeKho::OnTypeLoadData(){
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
/*void CRptPMS108_BaoCaoKiemKeKho::OnTypeAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BaoCaoKiemKeKho::OnGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BaoCaoKiemKeKho::OnGroupSelendok(){
	 
}
/*void CRptPMS108_BaoCaoKiemKeKho::OnGroupSetfocus(){
	
}*/
/*void CRptPMS108_BaoCaoKiemKeKho::OnGroupKillfocus(){
	
}*/
long CRptPMS108_BaoCaoKiemKeKho::OnGroupLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szSQL.Format(_T("SELECT pmg_id as id, pmg_name as description FROM pms_group ORDER BY pmg_id"));
	if(m_wndGroup.IsSearchKey() && !m_szGroupKey.IsEmpty()){
	};
	m_wndGroup.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndGroup.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CRptPMS108_BaoCaoKiemKeKho::OnGroupAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BaoCaoKiemKeKho::OnSourceSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BaoCaoKiemKeKho::OnSourceSelendok(){
	 
}
/*void CRptPMS108_BaoCaoKiemKeKho::OnSourceSetfocus(){
	
}*/
/*void CRptPMS108_BaoCaoKiemKeKho::OnSourceKillfocus(){
	
}*/
long CRptPMS108_BaoCaoKiemKeKho::OnSourceLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szSQL.Format(_T("SELECT pmo_id as id, pmo_name as description FROM pms_original ORDER BY pmo_id"));
	if(m_wndSource.IsSearchKey() && !m_szSourceKey.IsEmpty()){
	};
	m_wndSource.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndSource.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CRptPMS108_BaoCaoKiemKeKho::OnSourceAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BaoCaoKiemKeKho::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BaoCaoKiemKeKho::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BaoCaoKiemKeKho::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BaoCaoKiemKeKho::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CRptPMS108_BaoCaoKiemKeKho::OnAddRptPMS108_BaoCaoKiemKeKho(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CRptPMS108_BaoCaoKiemKeKho::OnEditRptPMS108_BaoCaoKiemKeKho(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CRptPMS108_BaoCaoKiemKeKho::OnDeleteRptPMS108_BaoCaoKiemKeKho(){
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
 		OnCancelRptPMS108_BaoCaoKiemKeKho(); 
 	}
	return 0;
}
int CRptPMS108_BaoCaoKiemKeKho::OnSaveRptPMS108_BaoCaoKiemKeKho(){
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
 		//OnRptPMS108_BaoCaoKiemKeKhoListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CRptPMS108_BaoCaoKiemKeKho::OnCancelRptPMS108_BaoCaoKiemKeKho(){
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
int CRptPMS108_BaoCaoKiemKeKho::OnRptPMS108_BaoCaoKiemKeKhoListLoadData(){
	return 0;
}
CString CRptPMS108_BaoCaoKiemKeKho::GetQueryString(){
	CString szSQL;
	return szSQL;
}