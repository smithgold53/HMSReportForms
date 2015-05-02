#include "stdafx.h"
#include "FMTempSendPatientListByDay.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMTempSendPatientListByDay *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMTempSendPatientListByDay* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMTempSendPatientListByDay *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMTempSendPatientListByDay *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMTempSendPatientListByDay *)pWnd)->OnToDateCheckValue();
} 
static void _OnStaffSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMTempSendPatientListByDay* )pWnd)->OnStaffSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStaffSelendokFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnStaffSelendok();
}
/*static void _OnStaffSetfocusFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnStaffKillfocus();
}*/
/*static void _OnStaffKillfocusFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnStaffKillfocus();
}*/
static long _OnStaffLoadDataFnc(CWnd *pWnd){
	return ((CFMTempSendPatientListByDay *)pWnd)->OnStaffLoadData();
}
/*static void _OnStaffAddNewFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay *)pWnd)->OnStaffAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMTempSendPatientListByDay *pVw = (CFMTempSendPatientListByDay *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMTempSendPatientListByDay *pVw = (CFMTempSendPatientListByDay *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMTempSendPatientListByDay*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMTempSendPatientListByDay*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMTempSendPatientListByDay*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMTempSendPatientListByDay*)pWnd)->OnDeptListDeleteItem();
} 
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMTempSendPatientListByDay*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMTempSendPatientListByDay*)pWnd)->OnDeptListUnCheckAll();
}
CFMTempSendPatientListByDay::CFMTempSendPatientListByDay(CWnd *pParent){

	m_nDlgWidth = 585;
	m_nDlgHeight = 585;
	SetDefaultValues();
}
CFMTempSendPatientListByDay::~CFMTempSendPatientListByDay(){
}
void CFMTempSendPatientListByDay::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 345);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndPrint.Create(this, _T("&Print"), 265, 350, 345, 375);
	m_wndExport.Create(this, _T("&Export"), 350, 350, 430, 375);
	m_wndDeptList.Create(this,10, 122, 425, 342); 
	m_wndStaffLabel.Create(this, _T("Staff"), 10, 90, 90, 115);
	m_wndStaff.Create(this,95, 90, 215, 115); 

}
void CFMTempSendPatientListByDay::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndStaff.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	m_wndStaff.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.SetCheckBox(TRUE);
}

void CFMTempSendPatientListByDay::OnSetWindowEvents(){
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
	m_wndStaff.SetEvent(WE_SELENDOK, _OnStaffSelendokFnc);
	//m_wndStaff.SetEvent(WE_SETFOCUS, _OnStaffSetfocusFnc);
	//m_wndStaff.SetEvent(WE_KILLFOCUS, _OnStaffKillfocusFnc);
	m_wndStaff.SetEvent(WE_SELCHANGE, _OnStaffSelectChangeFnc);
	m_wndStaff.SetEvent(WE_LOADDATA, _OnStaffLoadDataFnc);
	//m_wndStaff.SetEvent(WE_ADDNEW, _OnStaffAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);

	m_wndDeptList.SetEvent(WE_SELCHANGE, _OnDeptListSelectChangeFnc);
	m_wndDeptList.SetEvent(WE_LOADDATA, _OnDeptListLoadDataFnc);
	m_wndDeptList.SetEvent(WE_DBLCLICK, _OnDeptListDblClickFnc);

	m_wndDeptList.AddEvent(1, _T("Check All"), _OnDeptListCheckAllFnc);
	m_wndDeptList.AddEvent(2, _T("UnCheck All"), _OnDeptListUnCheckAllFnc);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);
	OnDeptListLoadData();

}
void CFMTempSendPatientListByDay::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStaff.GetDlgCtrlID(), m_szStaffKey);

}
void CFMTempSendPatientListByDay::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStaffKey.Empty();

}

int CFMTempSendPatientListByDay::SetMode(int nMode){
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

/*void CFMTempSendPatientListByDay::OnYearChange(){
	
} */
/*void CFMTempSendPatientListByDay::OnYearSetfocus(){
	
} */
/*void CFMTempSendPatientListByDay::OnYearKillfocus(){
	
} */
int CFMTempSendPatientListByDay::OnYearCheckValue(){
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
			szTemp.Format(_T("%.2d/.2d/.4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndFromDate.SetWindowText(szTemp);
		}
		dt.ParseDateTime(m_szToDate);
		date = dt.GetDate();
		if (date.GetYear() != 1752)
		{
			dt.SetDate(m_nYear, date.GetMonth(), date.GetDay());
			m_szToDate = dt.GetDateTime();
			szTemp.Format(_T("%.2d/.2d/.4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndToDate.SetWindowText(szTemp);
		}
	}

	UpdateData(FALSE);
	return 0;
}
 
void CFMTempSendPatientListByDay::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
void CFMTempSendPatientListByDay::OnReportPeriodSelendok(){
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

/*void CFMTempSendPatientListByDay::OnReportPeriodSetfocus(){
	
}*/
/*void CFMTempSendPatientListByDay::OnReportPeriodKillfocus(){
	
}*/
long CFMTempSendPatientListByDay::OnReportPeriodLoadData(){
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

/*void CFMTempSendPatientListByDay::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMTempSendPatientListByDay::OnFromDateChange(){
	
} */
/*void CFMTempSendPatientListByDay::OnFromDateSetfocus(){
	
} */
/*void CFMTempSendPatientListByDay::OnFromDateKillfocus(){
	
} */
int CFMTempSendPatientListByDay::OnFromDateCheckValue(){
	return 0;
}
 
/*void CFMTempSendPatientListByDay::OnToDateChange(){
	
} */
/*void CFMTempSendPatientListByDay::OnToDateSetfocus(){
	
} */
/*void CFMTempSendPatientListByDay::OnToDateKillfocus(){
	
} */
int CFMTempSendPatientListByDay::OnToDateCheckValue(){
	return 0;
}
void CFMTempSendPatientListByDay::OnStaffSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMTempSendPatientListByDay::OnStaffSelendok(){
	 
}
/*void CFMTempSendPatientListByDay::OnStaffSetfocus(){
	
}*/
/*void CFMTempSendPatientListByDay::OnStaffKillfocus(){
	
}*/
long CFMTempSendPatientListByDay::OnStaffLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStaff.IsSearchKey() && !m_szStaffKey.IsEmpty()){
		szWhere.Format(_T(" and su_userid='%s' "), m_szStaffKey);
	}
	m_wndStaff.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as userid, su_name as username FROM sys_user WHERE su_deptid = 'PTC' %s ORDER BY userid "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStaff.AddItems(
			rs.GetValue(_T("userid")), 
			rs.GetValue(_T("username")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMTempSendPatientListByDay::OnStaffAddNew(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} */
void CFMTempSendPatientListByDay::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
	UpdateData(TRUE);

	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp;
	CString szSysDate;

	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/HF_DANHSACHBENHNHANTAMGUI_1.RPT")))
		return;

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);

	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), pMF->GetCurrentDepartmentName());
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	CReportSection *rptDetail;
	CString szOldLine, szNewLine;
	CString szObjectNames;

	long double nGroupTotal = 0, nTotal = 0;
	double nCost;
	int nIndex = 1;

	szObjectNames.Empty();
	CStringArray strArr;
	bool bCheckServ = false, bCheckIns = false;

	strArr.Add(_T("\x64\x1ECB\x63h v\x1EE5"));
	strArr.Add(_T("\x42\x110 - \x43S - \x42H"));

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("receiptdate"), tmpStr);
		szNewLine = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		if (!szNewLine.IsEmpty() && szNewLine != szOldLine)
		{
			if (nGroupTotal > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng ng\xE0y :"));
				FormatCurrency(nGroupTotal, tmpStr);
				rptDetail->SetValue(_T("TotalAmount"), tmpStr);
				nTotal += nGroupTotal;
				nGroupTotal = 0;
			}
			szOldLine = szNewLine;
		}
		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("recordno"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("dept"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		
		if (nGroupTotal == 0)
		{
			rs.GetValue(_T("receiptdate"), tmpStr);
			rptDetail->SetValue(_T("6"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		}
		
		rs.GetValue(_T("deposit"), nCost);
		nGroupTotal += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		
		rs.MoveNext();
	}

	if (nGroupTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng kho\x61"));
		FormatCurrency(nGroupTotal, tmpStr);
		rptDetail->SetValue(_T("TotalAmount"), tmpStr);
		nTotal += nGroupTotal;
		nGroupTotal = 0;
	}

	if (nTotal > 0)
	{
		FormatCurrency(nTotal, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("Total"), tmpStr + _T(" \x111\x1ED3ng"));
		CString szMoney;
		tmpStr.Replace(_T("."), _T(""));
		MoneyToString(tmpStr, szMoney);
		rpt.GetReportFooter()->SetValue(_T("SumInWords"), szMoney + _T(" \x111\x1ED3ng"));
	}

	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
		          szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);

	EndWaitCursor();
	rpt.PrintPreview();
}
 
void CFMTempSendPatientListByDay::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	CString tmpStr, szTemp;

	UpdateData(TRUE);
	BeginWaitCursor();

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_ICONERROR | MB_OK);
		return;
	}

	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 18);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 20);
	xls.SetColumnWidth(4, 11);
	xls.SetColumnWidth(5, 15);

	int nRow = 1;
	int nCol = 0;
	int nIndex = 1;
	CString szOldLine, szNewLine;
	long double nGroupTotal = 0, nTotal = 0;
	double nCost = 0;

	xls.SetRowHeight(6, 45);

	xls.SetCellMergedColumns(0, 1, 4);
	xls.SetCellMergedColumns(0, 2, 4);

	xls.SetCellMergedColumns(6, 1, 4);
	xls.SetCellMergedColumns(6, 2, 4);

	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(6, 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(6, 2, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol, nRow + 3, 6);
	xls.SetCellMergedColumns(nCol, nRow + 4, 6);

	xls.SetCellText(nCol, nRow + 3, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N T\x1EA0M G\x1EECI "),
					FMT_TEXT | FMT_CENTER, true, 16, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);

	TranslateString(_T("Index"), tmpStr);
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Patient Name"), tmpStr);
	xls.SetCellText(nCol + 1, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Document No"), tmpStr);
	xls.SetCellText(nCol + 2, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Admission No"), tmpStr);
	xls.SetCellText(nCol + 3, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Dept"), tmpStr);
	xls.SetCellText(nCol + 4, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 5, nRow + 5, _T("Ng\xE0y n\x1ED9p ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
	xls.SetCellText(nCol + 6, nRow + 5, _T("S\x1ED1 ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
	nRow = 7;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("receiptdate"), tmpStr);
		szNewLine = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		if (!szNewLine.IsEmpty() && szNewLine != szOldLine)
		{
			if (nGroupTotal > 0)
			{
				nRow++;
				xls.SetCellText(nCol + 1, nRow, _T("\x43\x1ED9ng ng\xE0y"), FMT_TEXT | FMT_WRAPING, true);
				tmpStr.Format(_T("%.2lf"), nGroupTotal);
				xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
				nTotal += nGroupTotal;
				nGroupTotal = 0;
			}
			szOldLine = szNewLine;
		}
		nRow++;

		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("recordno"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);
			
		rs.GetValue(_T("dept"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		if (nGroupTotal == 0)
		{
			tmpStr = CDate::Convert(rs.GetValue(_T("receiptdate")), yyyymmdd, ddmmyyyy);
			xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);
		}
		rs.GetValue(_T("deposit"), nCost);
		nGroupTotal += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);
		rs.MoveNext();
		
	}
	
	if (nGroupTotal > 0)
	{
		nRow++;
		xls.SetCellText(nCol + 1, nRow, _T("\x43\x1ED9ng kho\x61"), FMT_TEXT | FMT_WRAPING, true);
		tmpStr.Format(_T("%.2lf"), nGroupTotal);
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
		nTotal += nGroupTotal;
		nGroupTotal = 0;
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\DSBNTamGui.xls"));
}

void CFMTempSendPatientListByDay::OnDeptListDblClick(){
	
}
 
void CFMTempSendPatientListByDay::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
int CFMTempSendPatientListByDay::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
}
 
long CFMTempSendPatientListByDay::OnDeptListLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;

	m_wndDeptList.BeginLoad();
	int nCount = 0;

	szSQL.Format(_T("SELECT sd_id as id, sd_name as name ") \
		         _T("FROM sys_dept ") \
				 _T("WHERE 1=1 AND sd_type='DT' ORDER BY id "));

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDeptList.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	m_wndDeptList.EndLoad();
	return nCount;
}

CString CFMTempSendPatientListByDay::GetQueryString(){
	CString szSQL, szWhere, szObjects, szDepts;

	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			if (!szDepts.IsEmpty())
				szDepts += _T(",");
			szDepts.AppendFormat(_T("'%s'"), m_wndDeptList.GetItemText(i, 0));
		}
	}
	if(!m_szStaffKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and hfe_staff = '%s' "), m_szStaffKey);
	}

	if (!szDepts.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and hfe_deptid in(%s) "), szDepts);
	}
	szSQL.Format(_T(" SELECT    Get_patientname(hfe_docno)     pname, ") \
				_T("			hfe_docno as docno, ") \
				_T("            Get_departmentname(hfe_deptid) dept, ") \
				_T("           hcr_recordno                     recordno, ") \
				_T("           fac_posteddate                       receiptdate, ") \
				_T("           hfe_amount                     deposit ") \
				_T(" FROM      hms_fee_deposit") \
				_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hfe_docno ) ") \
				_T(" LEFT JOIN fa_cash ON(hfe_cash_id = fac_cash_id)")
				_T(" WHERE     fac_posteddate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s') %s ") \
				_T("       AND hfe_status = 'P' AND hfe_objectid = '7' AND fac_posted = 'Y'") \
				_T(" ORDER     BY fac_posteddate "), m_szFromDate, m_szToDate, szWhere);

	return szSQL;
}

int CFMTempSendPatientListByDay::OnDeptListCheckAll()
{
	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (!m_wndDeptList.GetCheck(i))
		{
			m_wndDeptList.SetCheck(i, TRUE);
		}
	}
	return 0;
}

int CFMTempSendPatientListByDay::OnDeptListUnCheckAll()
{
	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			m_wndDeptList.SetCheck(i, FALSE);
		}
	}
	return 0;
}
