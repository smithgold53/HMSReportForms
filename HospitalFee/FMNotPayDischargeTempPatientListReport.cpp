#include "stdafx.h"
#include "FMNotPayDischargeTempPatientListReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMNotPayDischargeTempPatientListReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMNotPayDischargeTempPatientListReport* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnClerkAddNew();
}*/
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMNotPayDischargeTempPatientListReport* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnObjectAddNew();
}*/
static void _OnDepartmentSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMNotPayDischargeTempPatientListReport* )pWnd)->OnDepartmentSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDepartmentSelendokFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnDepartmentSelendok();
}
/*static void _OnDepartmentSetfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnDepartmentKillfocus();
}*/
/*static void _OnDepartmentKillfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnDepartmentKillfocus();
}*/
static long _OnDepartmentLoadDataFnc(CWnd *pWnd){
	return ((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnDepartmentLoadData();
}
/*static void _OnDepartmentAddNewFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnDepartmentAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMNotPayDischargeTempPatientListReport *pVw = (CFMNotPayDischargeTempPatientListReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMNotPayDischargeTempPatientListReport *pVw = (CFMNotPayDischargeTempPatientListReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CFMNotPayDischargeTempPatientListReport *pVw = (CFMNotPayDischargeTempPatientListReport *)pWnd;
	pVw->OnCloseSelect();
} 
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMNotPayDischargeTempPatientListReport *)pWnd)->OnYearCheckValue();
} 
static int _OnAddFMNotPayDischargeTempPatientListReportFnc(CWnd *pWnd){
	 return ((CFMNotPayDischargeTempPatientListReport*)pWnd)->OnAddFMNotPayDischargeTempPatientListReport();
} 
static int _OnEditFMNotPayDischargeTempPatientListReportFnc(CWnd *pWnd){
	 return ((CFMNotPayDischargeTempPatientListReport*)pWnd)->OnEditFMNotPayDischargeTempPatientListReport();
} 
static int _OnDeleteFMNotPayDischargeTempPatientListReportFnc(CWnd *pWnd){
	 return ((CFMNotPayDischargeTempPatientListReport*)pWnd)->OnDeleteFMNotPayDischargeTempPatientListReport();
} 
static int _OnSaveFMNotPayDischargeTempPatientListReportFnc(CWnd *pWnd){
	 return ((CFMNotPayDischargeTempPatientListReport*)pWnd)->OnSaveFMNotPayDischargeTempPatientListReport();
} 
static int _OnCancelFMNotPayDischargeTempPatientListReportFnc(CWnd *pWnd){
	 return ((CFMNotPayDischargeTempPatientListReport*)pWnd)->OnCancelFMNotPayDischargeTempPatientListReport();
} 
CFMNotPayDischargeTempPatientListReport::CFMNotPayDischargeTempPatientListReport(CWnd *pParent)
	{
	m_nDlgWidth = 445;
	m_nDlgHeight = 215;
	m_szTitle = _T("\x44\x61nh s\xE1\x63h \x42N t\x1EA1m g\x1EEDi \x111\xE3 r\x61 vi\x1EC7n \x63h\x1B0\x61 th\x61nh to\xE1n");
	SetDefaultValues();
}
CFMNotPayDischargeTempPatientListReport::~CFMNotPayDischargeTempPatientListReport(){
}
void CFMNotPayDischargeTempPatientListReport::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 100, 55);
	m_wndYear.Create(this,105, 30, 220, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 225, 30, 315, 55);
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 440, 180);
	m_wndReportPeriod.Create(this,320, 30, 435, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 100, 85);
	m_wndFromDate.Create(this,105, 60, 220, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 225, 60, 315, 85);
	m_wndToDate.Create(this,320, 60, 435, 85); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 100, 115);
	m_wndClerk.Create(this,105, 90, 435, 115); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 120, 100, 145);
	m_wndObject.Create(this,105, 120, 435, 145); 
	m_wndDepartmentLabel.Create(this, _T("Department"), 10, 150, 100, 175);
	m_wndDepartment.Create(this,105, 150, 435, 175); 
	m_wndPrint.Create(this, _T("&Print"), 265, 185, 350, 210);
	m_wndExport.Create(this, _T("&Export"), 355, 185, 440, 210);
}
void CFMNotPayDischargeTempPatientListReport::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	////m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
//	m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	//m_wndFromDate.SetCheckValue(true);
//	m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	//m_wndToDate.SetCheckValue(true);
	////m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);
	//m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(35);
	//m_wndDepartment.SetCheckValue(true);
	m_wndDepartment.LimitText(35);
	m_wndYear.SetLimitText(16);
	////m_wndYear.SetCheckValue(true);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_NUMBER, 100);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 100);
	m_wndObject.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndDepartment.InsertColumn(0, _T("ID"), CFMT_NUMBER, 100);
	m_wndDepartment.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);
}
void CFMNotPayDischargeTempPatientListReport::OnSetWindowEvents()
{
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
	m_wndClerk.SetEvent(WE_SELENDOK, _OnClerkSelendokFnc);
	//m_wndClerk.SetEvent(WE_SETFOCUS, _OnClerkSetfocusFnc);
	//m_wndClerk.SetEvent(WE_KILLFOCUS, _OnClerkKillfocusFnc);
	m_wndClerk.SetEvent(WE_SELCHANGE, _OnClerkSelectChangeFnc);
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	//m_wndClerk.SetEvent(WE_ADDNEW, _OnClerkAddNewFnc);
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndDepartment.SetEvent(WE_SELENDOK, _OnDepartmentSelendokFnc);
	//m_wndDepartment.SetEvent(WE_SETFOCUS, _OnDepartmentSetfocusFnc);
	//m_wndDepartment.SetEvent(WE_KILLFOCUS, _OnDepartmentKillfocusFnc);
	m_wndDepartment.SetEvent(WE_SELCHANGE, _OnDepartmentSelectChangeFnc);
	m_wndDepartment.SetEvent(WE_LOADDATA, _OnDepartmentLoadDataFnc);
	//m_wndDepartment.SetEvent(WE_ADDNEW, _OnDepartmentAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	//m_wndYear.SetEvent(WE_CHANGE, _OnYearChangeFnc);
	//m_wndYear.SetEvent(WE_SETFOCUS, _OnYearSetfocusFnc);
	//m_wndYear.SetEvent(WE_KILLFOCUS, _OnYearKillfocusFnc);
	m_wndYear.SetEvent(WE_CHECKVALUE, _OnYearCheckValueFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = pMF->GetSysDate() + _T("00:00");
	m_szToDate = pMF->GetSysDate() + _T("23:59");
	SetMode(VM_EDIT);
}
void CFMNotPayDischargeTempPatientListReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_TextEx(pDX, m_wndDepartment.GetDlgCtrlID(), m_szDepartmentKey);
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
}
void CFMNotPayDischargeTempPatientListReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CFMNotPayDischargeTempPatientListReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CFMNotPayDischargeTempPatientListReport::SetDefaultValues()
{
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_szObjectKey.Empty();
	m_szDepartmentKey.Empty();
	m_nYear=0;

}
int CFMNotPayDischargeTempPatientListReport::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, 1,2,3, -1); 
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
void CFMNotPayDischargeTempPatientListReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMNotPayDischargeTempPatientListReport::OnReportPeriodSelendok(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString tmpStr;
	CDate dte;
	UpdateData(true);
	int nYear;
	tmpStr = pMF->GetSysDate();
	nYear = ToInt(tmpStr.Left(4));
	int nMonth = ToInt(m_szReportPeriodKey);
	if(nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/01"), nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), nYear, nMonth, dte.GetMonthLastDay());
	}
	if(nMonth == 13){
		m_szFromDate.Format(_T("%.4d/01/01"), nYear);
		tmpStr.Format(_T("%.4d/03/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59" ), nYear,  dte.GetMonthLastDay());
	}
	if(nMonth == 14){
		m_szFromDate.Format(_T("%.4d/04/01"), nYear);
		tmpStr.Format(_T("%.4d/06/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}

	if(nMonth == 15){
		m_szFromDate.Format(_T("%.4d/07/01"), nYear);
		tmpStr.Format(_T("%.4d/09/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 16){
		m_szFromDate.Format(_T("%.4d/10/01"), nYear);
		tmpStr.Format(_T("%.4d/10/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 17){
		m_szFromDate.Format(_T("%.4d/01/01"), nYear);
		tmpStr.Format(_T("%.4d/12/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());		
	}	
	UpdateData(false);	 
}
/*void CFMNotPayDischargeTempPatientListReport::OnReportPeriodSetfocus(){
	
}*/
/*void CFMNotPayDischargeTempPatientListReport::OnReportPeriodKillfocus(){
	
}*/
long CFMNotPayDischargeTempPatientListReport::OnReportPeriodLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0)
	{
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
/*void CFMNotPayDischargeTempPatientListReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMNotPayDischargeTempPatientListReport::OnFromDateChange(){
	
} */
/*void CFMNotPayDischargeTempPatientListReport::OnFromDateSetfocus(){
	
} */
/*void CFMNotPayDischargeTempPatientListReport::OnFromDateKillfocus(){
	
} */
int CFMNotPayDischargeTempPatientListReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMNotPayDischargeTempPatientListReport::OnToDateChange(){
	
} */
/*void CFMNotPayDischargeTempPatientListReport::OnToDateSetfocus(){
	
} */
/*void CFMNotPayDischargeTempPatientListReport::OnToDateKillfocus(){
	
} */
int CFMNotPayDischargeTempPatientListReport::OnToDateCheckValue(){
	return 0;
} 
void CFMNotPayDischargeTempPatientListReport::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMNotPayDischargeTempPatientListReport::OnClerkSelendok(){
	 
}
/*void CFMNotPayDischargeTempPatientListReport::OnClerkSetfocus(){
	
}*/
/*void CFMNotPayDischargeTempPatientListReport::OnClerkKillfocus(){
	
}*/
long CFMNotPayDischargeTempPatientListReport::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty()){
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
/*void CFMNotPayDischargeTempPatientListReport::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMNotPayDischargeTempPatientListReport::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMNotPayDischargeTempPatientListReport::OnObjectSelendok(){
	 
}
/*void CFMNotPayDischargeTempPatientListReport::OnObjectSetfocus(){
	
}*/
/*void CFMNotPayDischargeTempPatientListReport::OnObjectKillfocus(){
	
}*/
long CFMNotPayDischargeTempPatientListReport::OnObjectLoadData(){
CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty()){
		szWhere.Format(_T(" where ho_id =%d "), ToInt(m_szObjectKey));
	};
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" select ho_type as id, ho_desc as name from hms_object %s order by cast(ho_id as integer)"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	return nCount;


}
/*void CFMNotPayDischargeTempPatientListReport::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMNotPayDischargeTempPatientListReport::OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMNotPayDischargeTempPatientListReport::OnDepartmentSelendok(){
	 
}
/*void CFMNotPayDischargeTempPatientListReport::OnDepartmentSetfocus(){
	
}*/
/*void CFMNotPayDischargeTempPatientListReport::OnDepartmentKillfocus(){
	
}*/
long CFMNotPayDischargeTempPatientListReport::OnDepartmentLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDepartment.IsSearchKey() && !m_szDepartmentKey.IsEmpty()){
		szWhere.Format(_T(" and sd_id ='%s' "), m_szDepartmentKey);
	};
	m_wndDepartment.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT sd_id as id, sd_name as name FROM sys_dept WHERE sd_type='DT' %s ORDER BY sd_id"), szWhere);
		
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDepartment.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CFMNotPayDischargeTempPatientListReport::OnDepartmentAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMNotPayDischargeTempPatientListReport::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere, tmpStr;
	if (!rpt.Init(_T("Reports/HMS/HF_DSBNTAMGUIDARAVIENCHUATHANHTOANTHEOKHOA1.RPT")))
		return;
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportObject")), m_wndObject.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("ReportObject"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("pMF->m_CompanyInfo.sc_name"), tmpStr);
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
	CReportSection *rptDetail;
	int index = 0;
	int i = 0;
	long double groupCost[7];
	long double ttlCost[7];
	for(i = 0; i <= 6 ; i++)
	{
		groupCost[i] = 0;
		ttlCost[i] = 0;
	}
	
	CString szOldDept, szNewDept;
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("Dept"), szNewDept);
		if (szOldDept!= szNewDept)
		{
			if(groupCost[i] > 0)
			{
				//FormatCurrency(groupCost[i], tmpStr);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				rptDetail->SetValue(_T("DeptSum"), tmpStr);
				ttlCost[i] += groupCost[i];
				groupCost[i] = 0;
			}
			szOldDept = szNewDept;
		}
		rptDetail = rpt.AddDetail();
		index++;
		tmpStr.Format(_T("%d"), index); 
		rptDetail->SetValue(_T("1"), tmpStr);
// ten 
		rs.GetValue(_T("patientname"), tmpStr);
		rptDetail->SetValue(_T("patientname"), tmpStr);
// so BA
		rs.GetValue(_T("Record"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
// Khoa
		rs.GetValue(_T("Dept"), szNewDept);
		rptDetail->SetValue(_T("4"), szNewDept);
// tien tam gui
		rs.GetValue(_T("tamgui"), tmpStr);
		groupCost[0] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
// tien che do
		rs.GetValue(_T("chedo"), tmpStr);
		groupCost[1] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
// tien chi phi duoc giam
		rs.GetValue(_T(""), tmpStr);
		groupCost[2] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
// vien phi phai thu
		rs.GetValue(_T(""), tmpStr);
		groupCost[3] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
// Thu
		rs.GetValue(_T(""), tmpStr);
		groupCost[4] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);
// chi
		rs.GetValue(_T("AdvPay"), tmpStr);
		groupCost[5] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);
		
		rs.MoveNext();
	}
	if(groupCost[i]> 0)
	{
		//FormatCurrency(groupCost[i], tmpStr);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));

		rptDetail->SetValue(_T("S%d"), tmpStr);
		ttlCost[i] += groupCost[i];
		groupCost[i] = 0;
	}
	if(ttlCost[i] >0)
	{
		//FormatCurrency(ttlCost[i], tmpStr);
		rpt.GetReportFooter()->SetValue(_T("l%d"), tmpStr);

	}
	tmpStr.Format(_T("%s"),pMF->GetSysDate());
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4)); 
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}
void CFMNotPayDischargeTempPatientListReport::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CRecord ors(&pMF->m_db);
	CExcel xls;
	CString szSQL, szWhere, szTemp, tmpStr;
	int nCol, nRow;
	nCol = nRow = 0;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	UpdateData(true);

	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 24);
	xls.SetColumnWidth(2, 8);
	xls.SetColumnWidth(3, 24);
	xls.SetColumnWidth(4, 17);
	xls.SetColumnWidth(5, 17);
	xls.SetColumnWidth(6, 17);
	xls.SetColumnWidth(7, 17);
	xls.SetColumnWidth(8, 17);
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow+1, 3);
	xls.SetCellMergedColumns(nCol, nRow+2, 10);
	xls.SetCellMergedColumns(nCol, nRow+3, 10);
	xls.SetCellMergedColumns(nCol, nRow+4, 6);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow+1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow+2, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N T\x1EA0M G\x1EECI \x110\xC3 R\x41 VI\x1EC6N \x43H\x1AF\x41 TH\x41NH TO\xC1N"), FMT_TEXT | FMT_CENTER, true, 12);
	szTemp.Format(_T("\x110\x1ED1i t\x1B0\x1EE3ng %s"), m_wndObject.GetCurrent(1));
	xls.SetCellText(nCol, nRow+3, szTemp, FMT_TEXT | FMT_CENTER);
	szTemp.Format(_T("T\xEDnh \x111\x1EBFn ng\xE0y: %s \x110\x1EBFn ng\xE0y: %s"), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(nCol, nRow+4, szTemp, FMT_TEXT | FMT_CENTER);

	xls.SetCellText(nCol, nRow+5, _T("STT"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 1, nRow + 5, _T("H\x1ECD v\xE0 t\xEAn"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 2, nRow + 5, _T("S\x1ED1 BA"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 3, nRow + 5, _T("Kho\x61"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 4, nRow + 5, _T("Ti\x1EC1n t\x1EA1m g\x1EEDi"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 5, nRow + 5, _T("Ti\x1EC1n \x63h\x1EBF \x111\x1ED9"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 6, nRow + 5, _T("\x43.Ph\xED \x111\x1B0\x1EE3\x63 gi\x1EA3m"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 7, nRow + 5, _T("Vi\x1EC7n ph\xED ph\x1EA3i thu"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 8, nRow + 5, _T("Thu"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 9, nRow + 5, _T("\x43hi"), FMT_TEXT | FMT_CENTER, true, 10);


	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	int index= 0;
	long double grpCost=0, ttlCost = 0;
	CString szOldLine, szNewLine;
	while (!rs.IsEOF())
	{
		index++;
		szTemp.Format(_T("%d"), index);
		xls.SetCellText(nCol, nRow+6, szTemp, FMT_INTEGER);
		
		rs.GetValue(_T("Name"), szTemp);
		xls.SetCellText(nCol+1, nRow+6, szTemp, FMT_TEXT);

		rs.GetValue(_T("Record"), szTemp);
		xls.SetCellText(nCol+2, nRow+6, szTemp, FMT_INTEGER);

		rs.GetValue(_T("Dept"), szNewLine);
		xls.SetCellText(nCol+3, nRow+6, szNewLine, FMT_TEXT);

		rs.GetValue(_T(""), tmpStr);
		xls.SetCellText(nCol+4, nRow+6, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("AdvPay"), szTemp);
		grpCost += ToDouble(szTemp);
		xls.SetCellText(nCol+5, nRow+6, szTemp, FMT_NUMBER1);
		nRow++;
		szOldLine = szNewLine;
		rs.MoveNext();
		rs.GetValue(_T("Dept"), szNewLine);
		if (szNewLine!=szOldLine)
		{
			xls.SetCellText(nCol+1, nRow+6, _T("\x43\x1ED9ng Kho\x61"), FMT_TEXT|FMT_CENTER, true);
			szTemp.Format(_T("%0.2f"), grpCost);
			xls.SetCellText(nCol+5, nRow+6, szTemp, FMT_NUMBER1, true);
			ttlCost += grpCost;
			grpCost = 0;
			nRow++;
		}
	}
	xls.SetCellText(nCol+1, nRow+6, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT|FMT_CENTER, true);
	szTemp.Format(_T("%0.2f"), ttlCost);
	xls.SetCellText(nCol+5, nRow+6, szTemp, FMT_NUMBER1, true);
	xls.Save(_T("Exports\\DANH SACH BENH NHAN TAM GUI CHUA RA VIEN.xls"));
	
} 
void CFMNotPayDischargeTempPatientListReport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
/*void CFMNotPayDischargeTempPatientListReport::OnYearChange(){
	
} */
/*void CFMNotPayDischargeTempPatientListReport::OnYearSetfocus(){
	
} */
/*void CFMNotPayDischargeTempPatientListReport::OnYearKillfocus(){
	
} */
int CFMNotPayDischargeTempPatientListReport::OnYearCheckValue(){
	return 0;
} 
int CFMNotPayDischargeTempPatientListReport::OnAddFMNotPayDischargeTempPatientListReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMNotPayDischargeTempPatientListReport::OnEditFMNotPayDischargeTempPatientListReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMNotPayDischargeTempPatientListReport::OnDeleteFMNotPayDischargeTempPatientListReport(){
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
 		OnCancelFMNotPayDischargeTempPatientListReport(); 
 	}
	return 0;
}
int CFMNotPayDischargeTempPatientListReport::OnSaveFMNotPayDischargeTempPatientListReport(){
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
 		//OnFMNotPayDischargeTempPatientListReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMNotPayDischargeTempPatientListReport::OnCancelFMNotPayDischargeTempPatientListReport(){
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
int CFMNotPayDischargeTempPatientListReport::OnFMNotPayDischargeTempPatientListReportListLoadData(){
	return 0;
}

CString CFMNotPayDischargeTempPatientListReport::GetQueryString(){
	CString szSQL, szWhere;
	if (!m_szObjectKey.IsEmpty()){
		szWhere.AppendFormat(_T(" AND hfe_objectid = %s"), m_szObjectKey);
	}
	if (!m_szClerkKey.IsEmpty()){
		szWhere.AppendFormat(_T(" AND hfe_staff = '%s'"), m_szClerkKey);
	}
	if (!m_szDepartmentKey.IsEmpty()){
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
