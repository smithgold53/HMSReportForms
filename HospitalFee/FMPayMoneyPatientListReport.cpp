#include "stdafx.h"
#include "FMPayMoneyPatientListReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMPayMoneyPatientListReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMPayMoneyPatientListReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMPayMoneyPatientListReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMPayMoneyPatientListReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMPayMoneyPatientListReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnDepartmentSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMPayMoneyPatientListReport* )pWnd)->OnDepartmentSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDepartmentSelendokFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnDepartmentSelendok();
}
/*static void _OnDepartmentSetfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnDepartmentKillfocus();
}*/
/*static void _OnDepartmentKillfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnDepartmentKillfocus();
}*/
static long _OnDepartmentLoadDataFnc(CWnd *pWnd){
	return ((CFMPayMoneyPatientListReport *)pWnd)->OnDepartmentLoadData();
}
/*static void _OnDepartmentAddNewFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnDepartmentAddNew();
}*/
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMPayMoneyPatientListReport* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CFMPayMoneyPatientListReport *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnObjectAddNew();
}*/
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMPayMoneyPatientListReport* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMPayMoneyPatientListReport *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMPayMoneyPatientListReport *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMPayMoneyPatientListReport *pVw = (CFMPayMoneyPatientListReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMPayMoneyPatientListReport *pVw = (CFMPayMoneyPatientListReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CFMPayMoneyPatientListReport *pVw = (CFMPayMoneyPatientListReport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddFMPayMoneyPatientListReportFnc(CWnd *pWnd){
	 return ((CFMPayMoneyPatientListReport*)pWnd)->OnAddFMPayMoneyPatientListReport();
} 
static int _OnEditFMPayMoneyPatientListReportFnc(CWnd *pWnd){
	 return ((CFMPayMoneyPatientListReport*)pWnd)->OnEditFMPayMoneyPatientListReport();
} 
static int _OnDeleteFMPayMoneyPatientListReportFnc(CWnd *pWnd){
	 return ((CFMPayMoneyPatientListReport*)pWnd)->OnDeleteFMPayMoneyPatientListReport();
} 
static int _OnSaveFMPayMoneyPatientListReportFnc(CWnd *pWnd){
	 return ((CFMPayMoneyPatientListReport*)pWnd)->OnSaveFMPayMoneyPatientListReport();
} 
static int _OnCancelFMPayMoneyPatientListReportFnc(CWnd *pWnd){
	 return ((CFMPayMoneyPatientListReport*)pWnd)->OnCancelFMPayMoneyPatientListReport();
} 
CFMPayMoneyPatientListReport::CFMPayMoneyPatientListReport(CWnd *pParent)
	{
	m_nDlgWidth = 420;
	m_nDlgHeight = 215;
	m_szTitle = _T("\x44\x61nh s\xE1\x63h \x62\x1EC7nh nh\xE2n n\x1ED9p ti\x1EC1n");
	SetDefaultValues();
}
CFMPayMoneyPatientListReport::~CFMPayMoneyPatientListReport(){
}
void CFMPayMoneyPatientListReport::OnCreateComponents()
{
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 205, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 215, 30, 295, 55);
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 415, 180);
	m_wndReportPeriod.Create(this,300, 30, 410, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 205, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 215, 60, 295, 85);
	m_wndToDate.Create(this,300, 60, 410, 85); 
	m_wndDepartmentLabel.Create(this, _T("Department"), 10, 90, 90, 115);
	m_wndDepartment.Create(this,95, 90, 410, 115); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 120, 90, 145);
	m_wndObject.Create(this,95, 120, 410, 145); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 150, 90, 175);
	m_wndClerk.Create(this,95, 150, 410, 175); 
	m_wndPrint.Create(this, _T("&Print"), 250, 185, 330, 210);
	m_wndExport.Create(this, _T("&Export"), 335, 184, 415, 210);
}
void CFMPayMoneyPatientListReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(false);
	m_wndReportPeriod.SetCheckValue(false);
	m_wndReportPeriod.LimitText(35);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(false);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(false);
	m_wndDepartment.SetCheckValue(false);
	m_wndDepartment.LimitText(35);
	m_wndObject.SetCheckValue(false);
	m_wndObject.LimitText(35);
	m_wndClerk.SetCheckValue(false);
	m_wndClerk.LimitText(35);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 80);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndDepartment.InsertColumn(0, _T("ID"), CFMT_NUMBER, 80);
	m_wndDepartment.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 80);
	m_wndObject.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_NUMBER, 80);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

}
void CFMPayMoneyPatientListReport::OnSetWindowEvents(){
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
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
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
void CFMPayMoneyPatientListReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDepartment.GetDlgCtrlID(), m_szDepartmentKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);

}
void CFMPayMoneyPatientListReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMPayMoneyPatientListReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMPayMoneyPatientListReport::SetDefaultValues(){
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDepartmentKey.Empty();
	m_szObjectKey.Empty();
	m_szClerkKey.Empty();

}
int CFMPayMoneyPatientListReport::SetMode(int nMode){
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
/*void CFMPayMoneyPatientListReport::OnYearChange(){
	
} */
/*void CFMPayMoneyPatientListReport::OnYearSetfocus(){
	
} */
/*void CFMPayMoneyPatientListReport::OnYearKillfocus(){
	
} */
int CFMPayMoneyPatientListReport::OnYearCheckValue(){
	return 0;
} 
void CFMPayMoneyPatientListReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMPayMoneyPatientListReport::OnReportPeriodSelendok(){
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
/*void CFMPayMoneyPatientListReport::OnReportPeriodSetfocus(){
	
}*/
/*void CFMPayMoneyPatientListReport::OnReportPeriodKillfocus(){
	
}*/
long CFMPayMoneyPatientListReport::OnReportPeriodLoadData(){
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
/*void CFMPayMoneyPatientListReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMPayMoneyPatientListReport::OnFromDateChange(){
	
} */
/*void CFMPayMoneyPatientListReport::OnFromDateSetfocus(){
	
} */
/*void CFMPayMoneyPatientListReport::OnFromDateKillfocus(){
	
} */
int CFMPayMoneyPatientListReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMPayMoneyPatientListReport::OnToDateChange(){
	
} */
/*void CFMPayMoneyPatientListReport::OnToDateSetfocus(){
	
} */
/*void CFMPayMoneyPatientListReport::OnToDateKillfocus(){
	
} */
int CFMPayMoneyPatientListReport::OnToDateCheckValue(){
	return 0;
} 
void CFMPayMoneyPatientListReport::OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMPayMoneyPatientListReport::OnDepartmentSelendok(){
	 
}
/*void CFMPayMoneyPatientListReport::OnDepartmentSetfocus(){
	
}*/
/*void CFMPayMoneyPatientListReport::OnDepartmentKillfocus(){
	
}*/
long CFMPayMoneyPatientListReport::OnDepartmentLoadData(){
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
/*void CFMPayMoneyPatientListReport::OnDepartmentAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMPayMoneyPatientListReport::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMPayMoneyPatientListReport::OnObjectSelendok(){
	 
}
/*void CFMPayMoneyPatientListReport::OnObjectSetfocus(){
	
}*/
/*void CFMPayMoneyPatientListReport::OnObjectKillfocus(){
	
}*/
long CFMPayMoneyPatientListReport::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && str2int(m_szObjectKey) > 0){
		szWhere.Format(_T(" WHERE ho_id = %d"), str2int(m_szObjectKey));
	};
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT ho_id as id, ho_desc as name FROM hms_object %s ORDER BY ho_id"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMPayMoneyPatientListReport::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMPayMoneyPatientListReport::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMPayMoneyPatientListReport::OnClerkSelendok(){
	 
}
/*void CFMPayMoneyPatientListReport::OnClerkSetfocus(){
	
}*/
/*void CFMPayMoneyPatientListReport::OnClerkKillfocus(){
	
}*/
long CFMPayMoneyPatientListReport::OnClerkLoadData(){
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
/*void CFMPayMoneyPatientListReport::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMPayMoneyPatientListReport::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	if(!rpt.Init(_T("Reports/HMS/HF_DANHSACHBNNOPTIEN.RPT")) )
	return ;
	UpdateData(true);
	szSQL = GetQueryString();
	int nRes = rs.ExecSQL(szSQL);
	if (nRes < 0)
	{
		//QueryOpener(szSQL);
		return;
	}
	if (rs.IsEOF())
	{
		MessageBox(_T("No Data!"), 0, MB_ICONASTERISK);
		//QueryOpener(szSQL);
		return;
	}

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm,ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	long double nTotal[8];
	for (int i = 1; i<= 7; i++)
	{
		nTotal[i] = 0.0;
	}
	int nIndex = 1;
		CReportSection* rptDetail;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%ld"), nIndex ++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T(""),tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		nTotal[1] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		nTotal[2] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		
		rs.GetValue(_T(""), tmpStr);
		nTotal[3] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		
		rs.GetValue(_T(""), tmpStr);
		nTotal[4] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		nTotal[5] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		nTotal[6] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.MoveNext();	
	}
	for (int i = 1; i<= 7; i++)
	{
		CString  szField;
		szField.Format(_T("s%d"), i);
		tmpStr.Format(_T("%.0f"), nTotal[i]);
		rpt.GetReportFooter()->SetValue(szField, tmpStr);		
	}
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
} 
void CFMPayMoneyPatientListReport::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	UpdateData(true);
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 26);
	xls.SetColumnWidth(2, 11);
	xls.SetColumnWidth(3, 11);
	xls.SetColumnWidth(4, 11);
	xls.SetColumnWidth(5, 11);
	xls.SetColumnWidth(6, 11);
	xls.SetColumnWidth(7, 11);
	xls.SetColumnWidth(8, 11);
	xls.SetColumnWidth(9, 11);
	xls.SetColumnWidth(10, 11);
	xls.SetColumnWidth(11, 11);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 2, _T("Ph\xF2ng t\xE0i \x63h\xEDnh"), FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellMergedColumns(nCol + 4, nRow + 1, 6);
	xls.SetCellText(nCol + 4, nRow + 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellMergedColumns(nCol + 4, nRow + 2, 6);
	xls.SetCellText(nCol + 4, nRow + 2, _T("\x110\x1ED9\x63 l\x1EADp - T\x1EF1 \x64o - H\x1EA1nh ph\xFA\x63"), FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellMergedColumns(nCol, nRow + 3, 9);
	xls.SetCellText(nCol, nRow + 3, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N N\x1ED8P TI\x1EC0N"), FMT_TEXT | FMT_CENTER, true, 14);

	xls.SetCellMergedColumns(nCol, nRow + 4, 9);
	xls.SetCellText(nCol, nRow + 4, _T("\x110\x1ED1i t\x1B0\x1EE3ng: %s"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedColumns(nCol, nRow + 5, 9);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellMergedRows(nCol, nRow + 7, 2);
	xls.SetCellText(nCol, nRow + 7, _T("STT"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedRows(nCol + 1, nRow + 7, 2);
	xls.SetCellText(nCol + 1, nRow + 7, _T("H\x1ECD v\xE0 t\xEAn"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol + 2, nRow + 7, 3);
	xls.SetCellText(nCol + 2, nRow + 7, _T("Thu"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 2, nRow + 8, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 3, nRow + 8, _T("\x42\x1ED3i \x64\x1B0\x1EE1ng"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 4, nRow + 8, _T("\x43\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol + 5, nRow + 7, 3);
	xls.SetCellText(nCol + 5, nRow + 7, _T("\x43hi tr\x1EA3 l\x1EA1i \x42N"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 5, nRow + 8, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 6, nRow + 8, _T("\x42\x1ED3i \x64\x1B0\x1EE1ng"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 7, nRow + 8, _T("\x43\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 10);

	int nIndex = 1;
	long double nTotal[8];
	for(int i = 1 ; i<= 7; i++)
	{
		nTotal[i] = 0.0;
	}
	while(!rs.IsEOF())
	{
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow + 6, tmpStr, FMT_INTEGER);

		rs.GetValue(_T(""),tmpStr);
		xls.SetCellText(nCol + 1, nRow + 6, tmpStr, FMT_TEXT);
 
		rs.GetValue(_T(""),tmpStr);
		nTotal[1] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 2, nRow + 6, tmpStr, FMT_TEXT);

		rs.GetValue(_T(""), tmpStr);
		nTotal[2] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 3, nRow + 6, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T(""),tmpStr);
		nTotal[3] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 4, nRow + 6, tmpStr, FMT_CURRENCY1);
		
		rs.GetValue(_T(""), tmpStr);
		nTotal[4] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 5, nRow + 6, tmpStr, FMT_CURRENCY1);

		rs.GetValue(_T(""), tmpStr);
		nTotal[5] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 6, nRow + 6, tmpStr, FMT_CURRENCY1);
		
		rs.GetValue(_T(""), tmpStr);
		nTotal[6] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 7, nRow + 6, tmpStr, FMT_CURRENCY1);

		nRow++;
		rs.MoveNext();
			
	}
	xls.SetCellMergedColumns(nCol, nRow + 6, 6);
	xls.SetCellText(nCol, nRow + 6, _T("T\x1ED5ng"), FMT_TEXT | FMT_CENTER, true, 11);
	for (int i = 1; i<= 7; i++)
	{
		tmpStr.Format(_T("%.0f"), nTotal[i]);
		xls.SetCellText(nCol, nRow + 6, tmpStr, FMT_CURRENCY1, true, 11);
	}
	CString szDate;
	xls.Save(_T("Exports\\DANH SACH BENH NHAN NOP TIEN.xls"));
	
} 
void CFMPayMoneyPatientListReport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
	
} 
int CFMPayMoneyPatientListReport::OnAddFMPayMoneyPatientListReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMPayMoneyPatientListReport::OnEditFMPayMoneyPatientListReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMPayMoneyPatientListReport::OnDeleteFMPayMoneyPatientListReport(){
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
 		OnCancelFMPayMoneyPatientListReport(); 
 	}
	return 0;
}
int CFMPayMoneyPatientListReport::OnSaveFMPayMoneyPatientListReport(){
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
 		//OnFMPayMoneyPatientListReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMPayMoneyPatientListReport::OnCancelFMPayMoneyPatientListReport(){
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
int CFMPayMoneyPatientListReport::OnFMPayMoneyPatientListReportListLoadData(){
	return 0;
}
CString CFMPayMoneyPatientListReport::GetQueryString(){
	CString szSQL, szWhere;
	if (!m_szObjectKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hfe_objectid = %s"), m_szObjectKey);
	}
	if (!m_szClerkKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hfe_staff = '%s'"), m_szClerkKey);
	}
	if (!m_szDepartmentKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hfe_deptid = '%s'"), m_szDepartmentKey);
	}
	szSQL.Format(_T(" SELECT 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) AS Name, ") \
					_T(" 	hcr_recordno AS Record,") \
					_T("    (select distinct sd_name from sys_dept where sd_id = hcr_admitdept) AS Dept,") \
					_T(" 	(hfe_date) AS RecvDate,") \
					_T(" 	hfe_deposit AS AdvPay") \
					_T(" FROM hms_clinical_record ") \
					_T(" LEFT JOIN hms_fee_invoice ON (hfe_docno = hcr_docno)") \
					_T(" LEFT JOIN hms_patient ON (hfe_patientno = hp_patientno)") \
					_T(" WHERE hcr_status = 'A'") \
					_T(" AND hfe_type = 'A' AND hfe_deposit >0") \
					_T(" AND (hfe_date) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
					_T(" ORDER BY hcr_docno"), m_szFromDate, m_szToDate, szWhere);
	return szSQL;
}
