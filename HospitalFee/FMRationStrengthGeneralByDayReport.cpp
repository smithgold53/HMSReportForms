#include "stdafx.h"
#include "FMRationStrengthGeneralByDayReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMRationStrengthGeneralByDayReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMRationStrengthGeneralByDayReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMRationStrengthGeneralByDayReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMRationStrengthGeneralByDayReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMRationStrengthGeneralByDayReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnDepartmentSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMRationStrengthGeneralByDayReport* )pWnd)->OnDepartmentSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDepartmentSelendokFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnDepartmentSelendok();
}
/*static void _OnDepartmentSetfocusFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnDepartmentKillfocus();
}*/
/*static void _OnDepartmentKillfocusFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnDepartmentKillfocus();
}*/
static long _OnDepartmentLoadDataFnc(CWnd *pWnd){
	return ((CFMRationStrengthGeneralByDayReport *)pWnd)->OnDepartmentLoadData();
}
/*static void _OnDepartmentAddNewFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnDepartmentAddNew();
}*/
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMRationStrengthGeneralByDayReport* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMRationStrengthGeneralByDayReport *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMRationStrengthGeneralByDayReport *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMRationStrengthGeneralByDayReport *pVw = (CFMRationStrengthGeneralByDayReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMRationStrengthGeneralByDayReport *pVw = (CFMRationStrengthGeneralByDayReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CFMRationStrengthGeneralByDayReport *pVw = (CFMRationStrengthGeneralByDayReport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddFMRationStrengthGeneralByDayReportFnc(CWnd *pWnd){
	 return ((CFMRationStrengthGeneralByDayReport*)pWnd)->OnAddFMRationStrengthGeneralByDayReport();
} 
static int _OnEditFMRationStrengthGeneralByDayReportFnc(CWnd *pWnd){
	 return ((CFMRationStrengthGeneralByDayReport*)pWnd)->OnEditFMRationStrengthGeneralByDayReport();
} 
static int _OnDeleteFMRationStrengthGeneralByDayReportFnc(CWnd *pWnd){
	 return ((CFMRationStrengthGeneralByDayReport*)pWnd)->OnDeleteFMRationStrengthGeneralByDayReport();
} 
static int _OnSaveFMRationStrengthGeneralByDayReportFnc(CWnd *pWnd){
	 return ((CFMRationStrengthGeneralByDayReport*)pWnd)->OnSaveFMRationStrengthGeneralByDayReport();
} 
static int _OnCancelFMRationStrengthGeneralByDayReportFnc(CWnd *pWnd){
	 return ((CFMRationStrengthGeneralByDayReport*)pWnd)->OnCancelFMRationStrengthGeneralByDayReport();
} 
CFMRationStrengthGeneralByDayReport::CFMRationStrengthGeneralByDayReport(CWnd *pParent)
	{
	m_nDlgWidth = 420;
	m_nDlgHeight = 185;
	m_szTitle = _T("T\x1ED5ng h\x1EE3p \x78u\x1EA5t \x103n th\x65o ng\xE0y");
	SetDefaultValues();
}
CFMRationStrengthGeneralByDayReport::~CFMRationStrengthGeneralByDayReport(){
}
void CFMRationStrengthGeneralByDayReport::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 205, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 215, 30, 295, 55);
	m_wndReportPeriod.Create(this,300, 30, 410, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 205, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 215, 60, 295, 85);
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 415, 148);
	m_wndToDate.Create(this,300, 60, 410, 85); 
	m_wndDepartmentLabel.Create(this, _T("Department"), 10, 90, 90, 115);
	m_wndDepartment.Create(this,95, 90, 410, 115); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 120, 90, 145);
	m_wndClerk.Create(this,95, 120, 410, 145); 
	m_wndPrint.Create(this, _T("&Print"), 250, 155, 330, 180);
	m_wndExport.Create(this, _T("&Export"), 335, 155, 415, 180);
}
void CFMRationStrengthGeneralByDayReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0,1,2,3, -1);
	m_wndYear.SetLimitText(16);
	////m_wndYear.SetCheckValue(true);
	////m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	//m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	//m_wndToDate.SetCheckValue(true);
	//m_wndDepartment.SetCheckValue(true);
	m_wndDepartment.LimitText(35);
	////m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

	m_wndDepartment.InsertColumn(0, _T("ID"), CFMT_NUMBER, 100);
	m_wndDepartment.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_NUMBER, 100);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

}
void CFMRationStrengthGeneralByDayReport::OnSetWindowEvents(){
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
	m_wndDepartment.SetEvent(WE_SELENDOK, _OnDepartmentSelendokFnc);
	//m_wndDepartment.SetEvent(WE_SETFOCUS, _OnDepartmentSetfocusFnc);
	//m_wndDepartment.SetEvent(WE_KILLFOCUS, _OnDepartmentKillfocusFnc);
	m_wndDepartment.SetEvent(WE_SELCHANGE, _OnDepartmentSelectChangeFnc);
	m_wndDepartment.SetEvent(WE_LOADDATA, _OnDepartmentLoadDataFnc);
	//m_wndDepartment.SetEvent(WE_ADDNEW, _OnDepartmentAddNewFnc);
	m_wndClerk.SetEvent(WE_SELENDOK, _OnClerkSelendokFnc);
	//m_wndClerk.SetEvent(WE_SETFOCUS, _OnClerkSetfocusFnc);
	//m_wndClerk.SetEvent(WE_KILLFOCUS, _OnClerkKillfocusFnc);
	m_wndClerk.SetEvent(WE_SELCHANGE, _OnClerkSelectChangeFnc);
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	//m_wndClerk.SetEvent(WE_ADDNEW, _OnClerkAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDateTime();
	SetMode(VM_EDIT);
}
void CFMRationStrengthGeneralByDayReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDepartment.GetDlgCtrlID(), m_szDepartmentKey);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);

}
void CFMRationStrengthGeneralByDayReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMRationStrengthGeneralByDayReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMRationStrengthGeneralByDayReport::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDepartmentKey.Empty();
	m_szClerkKey.Empty();

}
int CFMRationStrengthGeneralByDayReport::SetMode(int nMode){
 		int nOldMode = GetMode(); 
 		CGuiView::SetMode(nMode); 
 		CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
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
 			EnableButtons(TRUE,0, 1, 2, 3, 4, -1); 
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
/*void CFMRationStrengthGeneralByDayReport::OnYearChange(){
	
} */
/*void CFMRationStrengthGeneralByDayReport::OnYearSetfocus(){
	
} */
/*void CFMRationStrengthGeneralByDayReport::OnYearKillfocus(){
	
} */
int CFMRationStrengthGeneralByDayReport::OnYearCheckValue(){
	return 0;
} 
void CFMRationStrengthGeneralByDayReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMRationStrengthGeneralByDayReport::OnReportPeriodSelendok(){
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
/*void CFMRationStrengthGeneralByDayReport::OnReportPeriodSetfocus(){
	
}*/
/*void CFMRationStrengthGeneralByDayReport::OnReportPeriodKillfocus(){
	
}*/
long CFMRationStrengthGeneralByDayReport::OnReportPeriodLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
/*void CFMRationStrengthGeneralByDayReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMRationStrengthGeneralByDayReport::OnFromDateChange(){
	
} */
/*void CFMRationStrengthGeneralByDayReport::OnFromDateSetfocus(){
	
} */
/*void CFMRationStrengthGeneralByDayReport::OnFromDateKillfocus(){
	
} */
int CFMRationStrengthGeneralByDayReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMRationStrengthGeneralByDayReport::OnToDateChange(){
	
} */
/*void CFMRationStrengthGeneralByDayReport::OnToDateSetfocus(){
	
} */
/*void CFMRationStrengthGeneralByDayReport::OnToDateKillfocus(){
	
} */
int CFMRationStrengthGeneralByDayReport::OnToDateCheckValue(){
	return 0;
} 
void CFMRationStrengthGeneralByDayReport::OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMRationStrengthGeneralByDayReport::OnDepartmentSelendok(){
	 
}
/*void CFMRationStrengthGeneralByDayReport::OnDepartmentSetfocus(){
	
}*/
/*void CFMRationStrengthGeneralByDayReport::OnDepartmentKillfocus(){
	
}*/
long CFMRationStrengthGeneralByDayReport::OnDepartmentLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDepartment.IsSearchKey() && !m_szDepartmentKey.IsEmpty()){
		szWhere.Format(_T(" and sd_id = '%s' "), m_szDepartmentKey);
	};
	szSQL.Format(_T(" SELECT sd_id as id, sd_name as name from sys_dept where sd_type in('KB','DT') %s"), szWhere);
	m_wndDepartment.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDepartment.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMRationStrengthGeneralByDayReport::OnDepartmentAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMRationStrengthGeneralByDayReport::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMRationStrengthGeneralByDayReport::OnClerkSelendok(){
	 
}
/*void CFMRationStrengthGeneralByDayReport::OnClerkSetfocus(){
	
}*/
/*void CFMRationStrengthGeneralByDayReport::OnClerkKillfocus(){
	
}*/
long CFMRationStrengthGeneralByDayReport::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	};
	m_wndClerk.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as userid, su_name as name FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);
		
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("userid")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CFMRationStrengthGeneralByDayReport::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMRationStrengthGeneralByDayReport::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	if(!rpt.Init(_T("Reports/HMS/HF_BAOCAOXUATANTHEONGAY.RPT")) )
	return ;
	UpdateData(true);
// viet cau lenh

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm,ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	int nIndex = 1;
	CReportSection* rptDetail;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%ld"), nIndex ++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("DeptName"),tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("SumMoney"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.MoveNext();	
	}
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
} 
void CFMRationStrengthGeneralByDayReport::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	UpdateData(true);
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 41);
	xls.SetColumnWidth(2, 12);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 2, _T("Ph\xF2ng t\xE0i \x63h\xEDnh"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedColumns(nCol, nRow + 4, 3);
	xls.SetCellText(nCol, nRow + 4, _T("T\x1ED4NG H\x1EE2P \x58U\x1EA4T \x102N TH\x45O NG\xC0Y"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedColumns(nCol, nRow + 5, 3);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol, nRow + 7, _T("STT"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 1, nRow + 7, _T("Kho\x61"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 2, nRow + 7, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 10);

	int nIndex = 1;
	while(!rs.IsEOF())
	{
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow + 6, tmpStr, FMT_INTEGER);

		rs.GetValue(_T(""),tmpStr);
		xls.SetCellText(nCol + 1, nRow + 6, tmpStr, FMT_TEXT);
 
		rs.GetValue(_T(""),tmpStr);
		xls.SetCellText(nCol + 2, nRow + 6, tmpStr, FMT_TEXT);

		nRow++;
		rs.MoveNext();
			
	}
	xls.SetCellMergedColumns(nCol, nRow + 6, 2);
	xls.SetCellText(nCol, nRow + 6, _T("T\x1ED5ng"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.Save(_T("Exports\\TONG HOP XUAT AN THEO NGAY.xls"));
	
} 
void CFMRationStrengthGeneralByDayReport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
	
} 
int CFMRationStrengthGeneralByDayReport::OnAddFMRationStrengthGeneralByDayReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMRationStrengthGeneralByDayReport::OnEditFMRationStrengthGeneralByDayReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMRationStrengthGeneralByDayReport::OnDeleteFMRationStrengthGeneralByDayReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
 	CString szSQL; 
 	if(ShowMessage(1, MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2) == IDNO) 
 		return -1; 
 	szSQL.Format(_T("DELETE FROM  WHERE  AND") ); 
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret >= 0){ 
 		SetMode(VM_NONE); 
 		OnCancelFMRationStrengthGeneralByDayReport(); 
 	}
	return 0;
}
int CFMRationStrengthGeneralByDayReport::OnSaveFMRationStrengthGeneralByDayReport(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT) 
 		return -1; 
 	if(!IsValidateData()) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
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
 		//OnFMRationStrengthGeneralByDayReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMRationStrengthGeneralByDayReport::OnCancelFMRationStrengthGeneralByDayReport(){
 	if(GetMode() == VM_EDIT) 
 	{ 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 		SetMode(VM_NONE); 
 	} 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
 	return 0;
} 
int CFMRationStrengthGeneralByDayReport::OnFMRationStrengthGeneralByDayReportListLoadData(){
	return 0;
}
