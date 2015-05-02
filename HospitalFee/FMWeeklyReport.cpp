#include "stdafx.h"
#include "FMWeeklyReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

typedef struct tagCol
{
	CString szGroupID;
	CString szName;
	int nType;
	int nIndex;
}GROUPCOL;

const int MAX_COL = 11;

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMWeeklyReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMWeeklyReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMWeeklyReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMWeeklyReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMWeeklyReport *)pWnd)->OnToDateCheckValue();
}
/*static void _OnTotalCostChangeFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnTotalCostChange();
} */
/*static void _OnTotalCostSetfocusFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnTotalCostSetfocus();} */ 
/*static void _OnTotalCostKillfocusFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnTotalCostKillfocus();
} */
static int _OnTotalCostCheckValueFnc(CWnd *pWnd){
	return ((CFMWeeklyReport *)pWnd)->OnTotalCostCheckValue();
}
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMWeeklyReport* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMWeeklyReport *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMWeeklyReport *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMWeeklyReport *pVw = (CFMWeeklyReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMWeeklyReport *pVw = (CFMWeeklyReport *)pWnd;
	pVw->OnExportSelect();
}
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	  ((CFMWeeklyReport*)pWnd)->OnOutPatientSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	  ((CFMWeeklyReport*)pWnd)->OnInPatientSelect();
}
static int _OnAddFMWeeklyReportFnc(CWnd *pWnd){
	 return ((CFMWeeklyReport*)pWnd)->OnAddFMWeeklyReport();
} 
static int _OnEditFMWeeklyReportFnc(CWnd *pWnd){
	 return ((CFMWeeklyReport*)pWnd)->OnEditFMWeeklyReport();
} 
static int _OnDeleteFMWeeklyReportFnc(CWnd *pWnd){
	 return ((CFMWeeklyReport*)pWnd)->OnDeleteFMWeeklyReport();
} 
static int _OnSaveFMWeeklyReportFnc(CWnd *pWnd){
	 return ((CFMWeeklyReport*)pWnd)->OnSaveFMWeeklyReport();
} 
static int _OnCancelFMWeeklyReportFnc(CWnd *pWnd){
	 return ((CFMWeeklyReport*)pWnd)->OnCancelFMWeeklyReport();
} 
CFMWeeklyReport::CFMWeeklyReport(CWnd *pParent)
{
	m_nDlgWidth = 475;
	m_nDlgHeight = 220;
	SetDefaultValues();
}
CFMWeeklyReport::~CFMWeeklyReport()
{
}
void CFMWeeklyReport::OnCreateComponents()
{
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 435, 180);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 95, 55);
	m_wndYear.Create(this,100, 30, 220, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 225, 30, 305, 55);
	m_wndReportPeriod.Create(this,310, 30, 430, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 95, 85);
	m_wndFromDate.Create(this,100, 60, 220, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 225, 60, 305, 85);
	m_wndToDate.Create(this,310, 60, 430, 85); 
	m_wndTotalCostLabel.Create(this, _T("Min Total Cost"), 10, 90, 95, 115);
	m_wndTotalCost.Create(this,100, 90, 430, 115); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 120, 95, 145);
	m_wndClerk.Create(this,100, 120, 430, 145); 
	m_wndExport.Create(this, _T("&Export"), 335, 185, 435, 210);
	m_wndOutPatient.Create(this, _T("OutPatient"), 265, 150, 345, 175);
	m_wndInPatient.Create(this, _T("InPatient"), 350, 150, 430, 175);
}
void CFMWeeklyReport::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndTotalCost.SetLimitText(16);
	m_wndTotalCost.SetCheckValue(true);
	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(65);

	m_wndTotalCost.SetCurrencyFormat(TRUE);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 280);

}
void CFMWeeklyReport::OnSetWindowEvents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
	//m_wndTotalCost.SetEvent(WE_CHANGE, _OnTotalCostChangeFnc);
	//m_wndTotalCost.SetEvent(WE_SETFOCUS, _OnTotalCostSetfocusFnc);
	//m_wndTotalCost.SetEvent(WE_KILLFOCUS, _OnTotalCostKillfocusFnc);
	m_wndTotalCost.SetEvent(WE_CHECKVALUE, _OnTotalCostCheckValueFnc);
	m_wndClerk.SetEvent(WE_SELENDOK, _OnClerkSelendokFnc);
	//m_wndClerk.SetEvent(WE_SETFOCUS, _OnClerkSetfocusFnc);
	//m_wndClerk.SetEvent(WE_KILLFOCUS, _OnClerkKillfocusFnc);
	m_wndClerk.SetEvent(WE_SELCHANGE, _OnClerkSelectChangeFnc);
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	//m_wndClerk.SetEvent(WE_ADDNEW, _OnClerkAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	
	CString szSQL;
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);
}
void CFMWeeklyReport::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_Text(pDX, m_wndTotalCost.GetDlgCtrlID(), m_nTotalCost);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Radio(pDX, m_wndOutPatient.GetDlgCtrlID(), m_nOutPatient);
	DDX_Radio(pDX, m_wndInPatient.GetDlgCtrlID(), m_nInPatient);
}
void CFMWeeklyReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMWeeklyReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMWeeklyReport::SetDefaultValues()
{
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_nTotalCost=1700000;
	m_szClerkKey.Empty();
	m_nOutPatient=0;
	m_nInPatient=1;
}
int CFMWeeklyReport::SetMode(int nMode)
{
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
 			EnableButtons(TRUE, 0, 1, -1);
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
/*void CFMWeeklyReport::OnYearChange(){
	
} */
/*void CFMWeeklyReport::OnYearSetfocus(){
	
} */
/*void CFMWeeklyReport::OnYearKillfocus(){
	
} */
int CFMWeeklyReport::OnYearCheckValue()
{
	UpdateData(TRUE);
	if (m_nYear > 0)
	{
		CDateTime dt;
		CDate date;
		CString szTemp;

		dt.ParseDateTime(m_szFromDate);
		date = dt.GetDate();
		if (date.GetYear() != 1752)
		{
			dt.SetDate(m_nYear, date.GetMonth(), date.GetDay());
			m_szFromDate = dt.GetDateTime();
			szTemp.Format(_T("%.2d/%.2d/%.4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndFromDate.SetWindowText(szTemp);
		}
		dt.ParseDateTime(m_szToDate);
		date = dt.GetDate();
		if (date.GetYear() != 1752)
		{
			dt.SetDate(m_nYear, date.GetMonth(), date.GetDay());
			m_szToDate = dt.GetDateTime();
			szTemp.Format(_T("%.2d/%.2d/%.4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndToDate.SetWindowText(szTemp);
		}
	}
	UpdateData(FALSE);
	return 0;
} 
void CFMWeeklyReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMWeeklyReport::OnReportPeriodSelendok()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	CString tmpStr;
	CDate dte, date;

	UpdateData(true);

	date.ParseDate(pMF->GetSysDate());
	int nYear = date.GetYear();
	int nMonth = ToInt(m_szReportPeriodKey);

	if (nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/01 00:00"), nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), nYear, nMonth, dte.GetMonthLastDay());
	}
	if (nMonth == 13)
	{
		m_szFromDate.Format(_T("%.4d/01/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/03/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 14)
	{
		m_szFromDate.Format(_T("%.4d/04/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/06/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 15)
	{
		m_szFromDate.Format(_T("%.4d/07/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/09/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 16)
	{
		m_szFromDate.Format(_T("%.4d/10/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/10/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 17)
	{
		m_szFromDate.Format(_T("%.4d/01/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/12/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}

	UpdateData(false); 
}
/*void CFMWeeklyReport::OnReportPeriodSetfocus(){
	
}*/
/*void CFMWeeklyReport::OnReportPeriodKillfocus(){
	
}*/
long CFMWeeklyReport::OnReportPeriodLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0)
	{
		szWhere.Format(_T(" WHERE hpr_idx=%d "), ToInt(m_szReportPeriodKey));
	}
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
/*void CFMWeeklyReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMWeeklyReport::OnFromDateChange(){
	
} */
/*void CFMWeeklyReport::OnFromDateSetfocus(){
	
} */
/*void CFMWeeklyReport::OnFromDateKillfocus(){
	
} */
int CFMWeeklyReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMWeeklyReport::OnToDateChange(){
	
} */
/*void CFMWeeklyReport::OnToDateSetfocus(){
	
} */
/*void CFMWeeklyReport::OnToDateKillfocus(){
	
} */
int CFMWeeklyReport::OnToDateCheckValue(){
	return 0;
}
/*void CFMWeeklyReport::OnTotalCostChange(){
	
} */
/*void CFMWeeklyReport::OnTotalCostSetfocus(){
	
} */
/*void CFMWeeklyReport::OnTotalCostKillfocus(){
	
} */
int CFMWeeklyReport::OnTotalCostCheckValue()
{
	UpdateData(TRUE);
	if (m_nTotalCost < 0)
	{
		return -1;
	}
	return 0;
}
void CFMWeeklyReport::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMWeeklyReport::OnClerkSelendok(){
	 
}
/*void CFMWeeklyReport::OnClerkSetfocus(){
	
}*/
/*void CFMWeeklyReport::OnClerkKillfocus(){
	
}*/
long CFMWeeklyReport::OnClerkLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	}
	m_wndClerk.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as id, su_name as name FROM sys_user WHERE su_groupid in('A','F') %s ORDER BY su_userid"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){
		m_wndClerk.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMWeeklyReport::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMWeeklyReport::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMWeeklyReport::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CRecord rss(&pMF->m_db);

	UpdateData(TRUE);

	BeginWaitCursor();
	CString szSQL;
	CString szTemp, tmpStr;
	CString szStr;

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}

	CStringArray strArr;
	GROUPCOL grpCol;
	CArray<GROUPCOL, GROUPCOL&> arrCol;
	CArray<long double, long double> arrTotal;

	int nIndex = MAX_COL + 1;
	int nType = 0, nCount = 0;
	int n1 = 0, n2 = 0, n3 = 0;

	bool bCheck = false;
	strArr.RemoveAll();
	arrCol.RemoveAll();

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("groupid"), tmpStr);

		if (bCheck)
			bCheck = false;

		for (int i = 0; i < strArr.GetSize(); i++)
		{
			if (tmpStr == strArr[i])
			{
				bCheck = true;
				break;
			}
		}

		if (!bCheck)
		{
			if (tmpStr.Left(2) == _T("B1") ||
				tmpStr.Left(2) == _T("B2") ||
				tmpStr.Left(2) == _T("B3"))
				strArr.Add(tmpStr);
		}
		rs.MoveNext();
	}

	rs.MoveFirst();

	for (int i = 0; i < strArr.GetSize(); i++)
	{
		if (!szTemp.IsEmpty())
			szTemp += _T(",");
		szTemp.AppendFormat(_T("'%s'"), strArr[i]);
	}

	szSQL.Format(_T("SELECT hfg_id AS ID, hfg_name AS Name ") \
		         _T("FROM hms_fee_group ") \
				 _T("WHERE hfg_id IN(%s) ") \
				 _T("ORDER BY hfg_index,hfg_id"), szTemp);
	//MessageBox(szSQL);
	rss.ExecSQL(szSQL);

	while (!rss.IsEOF())
	{
		rss.GetValue(_T("ID"), tmpStr);
		grpCol.szGroupID = tmpStr;

		if (tmpStr.Left(2) == _T("B1"))
		{
			nType = 1;
			n1++;
		}
		else if (tmpStr.Left(2) == _T("B2"))
		{
			nType = 2;
			n2++;
		}
		else if (tmpStr.Left(2) == _T("B3"))
		{
			nType = 3;
			n3++;
		}

		rss.GetValue(_T("Name"), tmpStr);
		grpCol.szName = tmpStr;
		grpCol.nIndex = nIndex;
		grpCol.nType = nType;
		arrCol.Add(grpCol);

		nIndex++;
		rss.MoveNext();
	}

	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 8);
	xls.SetColumnWidth(2, 14);
	xls.SetColumnWidth(3, 5);

	xls.SetColumnWidth(4, 4);
	xls.SetColumnWidth(5, 16);

	xls.SetColumnWidth(6, 6);

	xls.SetColumnWidth(7, 9);
	xls.SetColumnWidth(8, 12);

	xls.SetColumnWidth(9, 10);
	xls.SetColumnWidth(10, 6);
	xls.SetColumnWidth(11, 9);

	for (int i = 0; i < arrCol.GetSize(); i++)
	{
		grpCol = arrCol[i];
		xls.SetColumnWidth(grpCol.nIndex, 9);
	}

	int nRow = 1;
	int nCol = 0;

	xls.SetColumnWidth(nCol + MAX_COL + (int)arrCol.GetSize() + 1, 10);
	xls.SetColumnWidth(nCol + MAX_COL + (int)arrCol.GetSize() + 2, 10);
	xls.SetColumnWidth(nCol + MAX_COL + (int)arrCol.GetSize() + 3, 10);
	xls.SetColumnWidth(nCol + MAX_COL + (int)arrCol.GetSize() + 4, 10);
	xls.SetColumnWidth(nCol + MAX_COL + (int)arrCol.GetSize() + 5, 10);

	xls.SetRowHeight(6, 30);
	xls.SetRowHeight(7, 50);

	xls.SetCellMergedColumns(0, 1, 4);
	xls.SetCellMergedColumns(0, 2, 4);
	xls.SetCellMergedColumns(nCol + MAX_COL + (int)arrCol.GetSize() + 1, 1, 5);
	xls.SetCellMergedColumns(nCol + MAX_COL + (int)arrCol.GetSize() + 1, 2, 5);

	xls.SetCellMergedColumns(nCol, nRow + 3, MAX_COL + (int)arrCol.GetSize() + 6);
	xls.SetCellMergedColumns(nCol, nRow + 4, MAX_COL + (int)arrCol.GetSize() + 6);

	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol + MAX_COL + (int)arrCol.GetSize() + 1, 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + MAX_COL + (int)arrCol.GetSize() + 1, 2, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol, nRow + 3, _T("\x42\xC1O \x43\xC1O THU VI\x1EC6N PH\xCD TU\x1EA6N"),
					FMT_TEXT | FMT_CENTER, true, 16, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);

	TranslateString(_T("Index"), tmpStr);
	xls.SetCellMergedRows(nCol, nRow + 5, 2);
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	TranslateString(_T("Doc No"), tmpStr);
	xls.SetCellMergedRows(nCol + 1, nRow + 5, 2);
	xls.SetCellText(nCol + 1, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	TranslateString(_T("Full Name"), tmpStr);
	xls.SetCellMergedRows(nCol + 2, nRow + 5, 2);
	xls.SetCellText(nCol + 2, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	TranslateString(_T("Birth year"), tmpStr);
	xls.SetCellMergedRows(nCol + 3, nRow + 5, 2);
	xls.SetCellText(nCol + 3, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	TranslateString(_T("Sex"), tmpStr);
	xls.SetCellMergedRows(nCol + 4, nRow + 5, 2);
	xls.SetCellText(nCol + 4, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	xls.SetCellMergedRows(nCol + 5, nRow + 5, 2);
	xls.SetCellText(nCol + 5, nRow + 5, _T("Th\x1EBB \x42HYT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	TranslateString(_T("Object"), tmpStr);
	xls.SetCellMergedRows(nCol + 6, nRow + 5, 2);
	xls.SetCellText(nCol + 6, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	TranslateString(_T("Exam Room"), tmpStr);
	xls.SetCellMergedRows(nCol + 7, nRow + 5, 2);
	xls.SetCellText(nCol + 7, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	xls.SetCellMergedRows(nCol + 8, nRow + 5, 2);
	xls.SetCellText(nCol + 8, nRow + 5, _T("M\xE3 \x62\x1EC7nh"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	TranslateString(_T("Total Cost"), tmpStr);
	xls.SetCellMergedRows(nCol + 9, nRow + 5, 2);
	xls.SetCellText(nCol + 9, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	xls.SetCellMergedRows(nCol + 10, nRow + 5, 2);
	xls.SetCellText(nCol + 10, nRow + 5, _T("Ti\x1EC1n kh\xE1m"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	xls.SetCellMergedRows(nCol + 11, nRow + 5, 2);
	xls.SetCellText(nCol + 11, nRow + 5, _T("Ti\x1EC1n thu\x1ED1\x63"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	nType = -1;
	nCount = 0;

	if (n1 > 0)
	{
		xls.SetCellMergedColumns(nCol + MAX_COL + 1, nRow + 5, n1);
		xls.SetCellText(nCol + MAX_COL + 1, nRow + 5, _T("\x58\xE9t nghi\x1EC7m"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);
	}
	if (n2 > 0)
	{
		xls.SetCellMergedColumns(nCol + MAX_COL + n1 + 1, nRow + 5, n2);
		xls.SetCellText(nCol + MAX_COL + n1 + 1, nRow + 5, _T("\x43h\x1EA9n \x111o\xE1n h\xECnh \x1EA3nh"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);
	}
	if (n3 > 0)
	{
		xls.SetCellMergedColumns(nCol + MAX_COL + n1 + n2 + 1, nRow + 5, n3);
		xls.SetCellText(nCol + MAX_COL + n1 + n2 + 1, nRow + 5, _T("Th\x103m \x64\xF2 \x63h\x1EE9\x63 n\x103ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);
	}

	for (int i = 0; i < arrCol.GetSize(); i++)
	{
		grpCol = arrCol[i];
		xls.SetCellText(nCol + grpCol.nIndex, nRow + 6, grpCol.szName, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);
		nCount++;
	}

	xls.SetCellMergedRows(nCol + MAX_COL + (int)arrCol.GetSize() + 1, nRow + 5, 2);
	xls.SetCellText(nCol + MAX_COL + (int)arrCol.GetSize() + 1, nRow + 5, _T("PT, TT"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	xls.SetCellMergedRows(nCol + MAX_COL + (int)arrCol.GetSize() + 2, nRow + 5, 2);
	xls.SetCellText(nCol + MAX_COL + (int)arrCol.GetSize() + 2, nRow + 5, _T("\x44VKT \x63\x61o"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	xls.SetCellMergedRows(nCol + MAX_COL + (int)arrCol.GetSize() + 3, nRow + 5, 2);
	xls.SetCellText(nCol + MAX_COL + (int)arrCol.GetSize() + 3, nRow + 5, _T("\x43P v\x1EADn \x63huy\x1EC3n"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	xls.SetCellMergedRows(nCol + MAX_COL + (int)arrCol.GetSize() + 4, nRow + 5, 2);
	xls.SetCellText(nCol + MAX_COL + (int)arrCol.GetSize() + 4, nRow + 5, _T("\x42H tr\x1EA3"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);
	
	xls.SetCellMergedRows(nCol + MAX_COL + (int)arrCol.GetSize() + 5, nRow + 5, 2);
	xls.SetCellText(nCol + MAX_COL + (int)arrCol.GetSize() + 5, nRow + 5, _T("\x42N tr\x1EA3"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10);

	szTemp.Empty();
	arrTotal.RemoveAll();
	nIndex = 1;
	nRow += 7;
	double nCost;
	long double nTemp;

	for (int i = 0; i < arrCol.GetSize() + MAX_COL + 6; i++)
		arrTotal.Add(0);

	while (!rs.IsEOF())
	{
		if (szTemp != rs.GetValue(_T("docno")))
		{
			if (nIndex > 1)
			{
				for (int i = 0; i < arrTotal.GetSize(); i++)
				{
					nTemp = arrTotal[i];
					if (nTemp > 0)
					{
						tmpStr.Format(_T("%.2lf"), nTemp);
						xls.SetCellText(nCol + i, nRow - 1, tmpStr, FMT_NUMBER1 | FMT_WRAPING);
						arrTotal.SetAt(i, 0);
					}
				}
			}
			tmpStr.Format(_T("%d"), nIndex++);
			xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

			rs.GetValue(_T("docno"), tmpStr);
			xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

			szStr.Empty();

			szSQL.Format(_T(" select distinct hrl_name||'['||he_deptid||']' AS itemname, he_receptidx") \
						_T(" from hms_exam") \
						_T(" left join hms_roomlist on(hrl_id=he_roomid and hrl_deptid=he_deptid)") \
						_T(" where he_docno=%s") \
						_T(" order by he_receptidx"), tmpStr);
			rss.ExecSQL(szSQL);

			while (!rss.IsEOF())
			{
				if (!szStr.IsEmpty())
					szStr += _T(",");
				szStr.AppendFormat(_T("%s"), rss.GetValue(_T("itemname")));
				rss.MoveNext();
			}

			xls.SetCellText(nCol + 7, nRow, szStr, FMT_TEXT | FMT_WRAPING);

			szStr.Empty();

			szSQL.Format(_T(" select distinct hi_name||'['||hi_icd||']' AS itemname, he_receptidx") \
						_T(" from hms_exam") \
						_T(" left join hms_icd on(hi_icd=he_icd10)") \
						_T(" where he_docno=%s and length(he_icd10)>0") \
						_T(" order by he_receptidx"), tmpStr);
			rss.ExecSQL(szSQL);

			while (!rss.IsEOF())
			{
				if (!szStr.IsEmpty())
					szStr += _T(",");
				szStr.AppendFormat(_T("%s"), rss.GetValue(_T("itemname")));
				rss.MoveNext();
			}

			xls.SetCellText(nCol + 8, nRow, szStr, FMT_TEXT | FMT_WRAPING);

			rs.GetValue(_T("pname"), tmpStr);
			xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

			rs.GetValue(_T("birthyear"), tmpStr);
			xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

			rs.GetValue(_T("sex"), tmpStr);
			xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

			rs.GetValue(_T("cardno"), tmpStr);
			xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

			rs.GetValue(_T("obj"), tmpStr);
			rs.GetValue(_T("hdline"), tmpStr);
			//xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);
			if (tmpStr == _T("0"))
			{
				xls.SetCellText(nCol + 6, nRow, _T("\x110\xFAng tuy\x1EBFn"), FMT_TEXT | FMT_WRAPING);
			}
			else
			{
				xls.SetCellText(nCol + 6, nRow, _T("Tr\xE1i tuy\x1EBFn"), FMT_TEXT | FMT_WRAPING);
			}

			rs.GetValue(_T("ttlinspaid"), tmpStr);
			xls.SetCellText(nCol + 9, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

			rs.GetValue(_T("ttlpatpaid"), tmpStr);
			xls.SetCellText(nCol + MAX_COL + (int)arrCol.GetSize() + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

			rs.GetValue(_T("ttldiscount"), tmpStr);
			xls.SetCellText(nCol + MAX_COL + (int)arrCol.GetSize() + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

			nRow++;
			rs.GetValue(_T("docno"), szTemp);
		}

		rs.GetValue(_T("groupid"), tmpStr);
		rs.GetValue(_T("hitech"), szStr);
		rs.GetValue(_T("inspaid"), nCost);

		if (tmpStr == _T("D0000"))
		{
			nTemp = arrTotal[MAX_COL - 1];
			nTemp += nCost;
			arrTotal.SetAt(MAX_COL - 1, nTemp);
		}
		if ((tmpStr.Left(2) == _T("A1") || tmpStr.Left(2) == _T("A6"))
			&& tmpStr != _T("A1400"))
		{
			nTemp = arrTotal[MAX_COL];
			nTemp += nCost;
			arrTotal.SetAt(MAX_COL, nTemp);
		}

		for (int i = 0; i < arrCol.GetSize(); i++)
		{
			grpCol = arrCol[i];

			if (tmpStr == grpCol.szGroupID)
			{
				nTemp = arrTotal[grpCol.nIndex];
				nTemp += nCost;
				arrTotal.SetAt(grpCol.nIndex, nTemp);
				break;
			}
		}

		if (tmpStr.Left(2) == _T("B4") || tmpStr.Left(2) == _T("B5"))
		{
			nTemp = arrTotal[MAX_COL + (int)arrCol.GetSize() + 1];
			nTemp += nCost;
			arrTotal.SetAt(MAX_COL + (int)arrCol.GetSize() + 1, nTemp);
		}

		if (szStr == _T("Y"))
		{
			nTemp = arrTotal[MAX_COL + (int)arrCol.GetSize() + 2];
			nTemp += nCost;
			arrTotal.SetAt(MAX_COL + (int)arrCol.GetSize() + 2, nTemp);
		}

		rs.MoveNext();
	}

	if (nIndex > 1)
	{
		for (int i = 0; i < arrTotal.GetSize(); i++)
		{
			nTemp = arrTotal[i];
			if (nTemp > 0)
			{
				tmpStr.Format(_T("%.2lf"), nTemp);
				xls.SetCellText(nCol + i, nRow - 1, tmpStr, FMT_NUMBER1 | FMT_WRAPING);
				arrTotal.SetAt(i, 0);
			}
		}
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\BaoCaoThuVPTuan.xls"));
}
void CFMWeeklyReport::OnOutPatientSelect()
{
	
}
void CFMWeeklyReport::OnInPatientSelect()
{
	
}
int CFMWeeklyReport::OnAddFMWeeklyReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMWeeklyReport::OnEditFMWeeklyReport(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMWeeklyReport::OnDeleteFMWeeklyReport(){
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
 		OnCancelFMWeeklyReport();
 	}
	return 0;
}
int CFMWeeklyReport::OnSaveFMWeeklyReport(){
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
 		//OnFMWeeklyReportListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CFMWeeklyReport::OnCancelFMWeeklyReport(){
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
int CFMWeeklyReport::OnFMWeeklyReportListLoadData(){
	return 0;
}

CString CFMWeeklyReport::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Empty();

	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" and fe.hfe_staff='%s' "), m_szClerkKey);

	if (m_nOutPatient == 0)
	{
		szWhere.AppendFormat(_T(" and fi.hfe_class in('E') and hd_status='T' "));
	}
	else
	{
		szWhere.AppendFormat(_T(" and fi.hfe_class not in('E','S') and hcr_status='T' "));
	}
		
	szSQL.Format(_T(" select docno,") \
				_T("        max(invoiceno) as invoiceno,") \
				_T("        pname,") \
				_T("        birthyear,") \
				_T("        sex,") \
				_T("        cardno,") \
				_T("        obj,") \
				_T("        hdline,") \
				_T("        sum(ttlinspaid) as ttlinspaid,") \
				_T("        sum(ttldiscount) as ttldiscount,") \
				_T("        sum(ttlpatpaid) as ttlpatpaid,") \
				_T("        groupid,") \
				_T("        sum(inspaid) as inspaid,") \
				_T("        hitech") \
				_T(" from") \
				_T(" (") \
				_T("  select hd_docno as docno,") \
				_T("         fe.hfe_invoiceno as invoiceno,") \
				_T("         trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("         extract(year from hp_birthdate) as birthyear,") \
				_T("         (select distinct ss_desc from sys_sel where ss_id='sys_sex' and ss_code=hp_sex) as sex,") \
				_T("         case when length(hc_cardno) > 15 then substr(hc_cardno,1,15) else hc_cardno end as cardno,") \
				_T("         ho_id||'-'||ho_desc as obj,") \
				_T("         case when hd_insline='Y' then 1 else 0 end as hdline,") \
				_T("         fe.hfe_inspaid as ttlinspaid,") \
				_T("         fe.hfe_discount as ttldiscount,") \
				_T("         fe.hfe_inspaid-fe.hfe_discount as ttlpatpaid,") \
				_T("         hfe_group as groupid,") \
				_T("         fi.hfe_inspaid as inspaid,") \
				_T("         fi.hfe_hitech as hitech") \
				_T("  from hms_patient") \
				_T("  left join hms_doc on(hd_patientno=hp_patientno)") \
				_T("  left join hms_fee_view fi on(hfe_docno=hd_docno)") \
				_T("  left join hms_fee_invoice fe on(fe.hfe_docno=fi.hfe_docno and fe.hfe_invoiceno=fi.hfe_invoiceno)") \
				_T("  left join hms_card on(hc_patientno=hp_patientno and hc_idx=hd_cardidx)") \
				_T("  left join hms_object on(hd_object=ho_id)") \
				_T("  where fe.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T("        and fe.hfe_amount >= %f and ho_type in('I','C') and fi.hfe_discount > 0") \
				_T("        and length(hc_cardno) > 1 and fe.hfe_status='P' %s") \
				_T(" ) tbl") \
				_T(" group by docno, pname, groupid, pname, birthyear, sex, cardno, obj, hitech, hdline") \
				_T(" order by docno"),
				m_szFromDate, m_szToDate, m_nTotalCost, szWhere);

	//MessageBox(szSQL);

	return szSQL;
}