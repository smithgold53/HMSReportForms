#include "stdafx.h"
#include "FMOutpatientInsuranceReport21a1.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMOutpatientInsuranceReport21a1* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnToDateCheckValue();
} 
static void _OnPatientLineSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMOutpatientInsuranceReport21a1* )pWnd)->OnPatientLineSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnPatientLineSelendokFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnPatientLineSelendok();
}
/*static void _OnPatientLineSetfocusFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnPatientLineKillfocus();
}*/
/*static void _OnPatientLineKillfocusFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnPatientLineKillfocus();
}*/
static long _OnPatientLineLoadDataFnc(CWnd *pWnd){
	return ((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnPatientLineLoadData();
}
/*static void _OnPatientLineAddNewFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1 *)pWnd)->OnPatientLineAddNew();
}*/
static void _OnDrugPTTTSelectFnc(CWnd *pWnd){
	 ((CFMOutpatientInsuranceReport21a1*)pWnd)->OnDrugPTTTSelect();
}
static void _OnByDischargedDateSelectFnc(CWnd *pWnd){
	 ((CFMOutpatientInsuranceReport21a1*)pWnd)->OnByDischargedDateSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMOutpatientInsuranceReport21a1 *pVw = (CFMOutpatientInsuranceReport21a1 *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMOutpatientInsuranceReport21a1 *pVw = (CFMOutpatientInsuranceReport21a1 *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CFMOutpatientInsuranceReport21a1*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CFMOutpatientInsuranceReport21a1*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMOutpatientInsuranceReport21a1*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMOutpatientInsuranceReport21a1*)pWnd)->OnListDeleteItem();
} 
CFMOutpatientInsuranceReport21a1::CFMOutpatientInsuranceReport21a1(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CFMOutpatientInsuranceReport21a1::~CFMOutpatientInsuranceReport21a1(){
}
void CFMOutpatientInsuranceReport21a1::OnCreateComponents(){
	m_wndReportFilter.Create(this, _T("Report Parameters"), 5, 5, 440, 520);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 120, 435, 515);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Month"), 220, 30, 310, 55);
	m_wndReportPeriod.Create(this,315, 30, 435, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 310, 85);
	m_wndToDate.Create(this,315, 60, 435, 85); 
	m_wndPatientLineLabel.Create(this, _T("Patient Line"), 10, 90, 90, 115);
	m_wndPatientLine.Create(this,95, 90, 435, 115); 
	m_wndDrugPTTT.Create(this, _T("Without Operation Material"), 10, 525, 190, 550);
	m_wndByDischargedDate.Create(this, _T("By Discharged Date"), 195, 525, 380, 550);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 235, 555, 335, 580);
	m_wndExport.Create(this, _T("&Export"), 340, 555, 440, 580);
	m_wndList.Create(this,15, 145, 430, 510); 

}
void CFMOutpatientInsuranceReport21a1::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);

	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 180);

	m_wndPatientLine.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndPatientLine.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndList.InsertColumn(1, _T("Name"), CFMT_TEXT, 350);
	m_wndList.SetCheckBox(TRUE);
}

void CFMOutpatientInsuranceReport21a1::OnSetWindowEvents(){
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
	m_wndPatientLine.SetEvent(WE_SELENDOK, _OnPatientLineSelendokFnc);
	//m_wndPatientLine.SetEvent(WE_SETFOCUS, _OnPatientLineSetfocusFnc);
	//m_wndPatientLine.SetEvent(WE_KILLFOCUS, _OnPatientLineKillfocusFnc);
	m_wndPatientLine.SetEvent(WE_SELCHANGE, _OnPatientLineSelectChangeFnc);
	m_wndPatientLine.SetEvent(WE_LOADDATA, _OnPatientLineLoadDataFnc);
	//m_wndPatientLine.SetEvent(WE_ADDNEW, _OnPatientLineAddNewFnc);
	m_wndDrugPTTT.SetEvent(WE_CLICK, _OnDrugPTTTSelectFnc);
	m_wndByDischargedDate.SetEvent(WE_CLICK, _OnByDischargedDateSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);
	m_wndList.AddEvent(1, _T("Delete"), _OnListDeleteItemFnc, 0, VK_DELETE, 0);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	m_mapChapter[_T("A")] = _T("\x58\xE9t nghi\x1EC7m");
	m_mapChapter[_T("B")] = _T("\x43h\x1EA9n \x111o\xE1n h\xECnh \x1EA3nh");
	m_mapChapter[_T("C")] = _T("Th\x103m \x64\xF2 \x63h\x1EE9\x63 n\x103ng");
	m_mapChapter[_T("D")] = _T("Ph\x1EABu thu\xE2t- th\x1EE7 thu\x1EADt");
	m_mapChapter[_T("E")] = _T("V\x1EADt t\x1B0 ti\xEAu h\x61o");
	m_mapChapter[_T("F")] = _T("\x44V k\x1EF9 thu\x1EADt \x63\x61o");

	m_mapLine[_T("I")] = _T("\x110\xFAng tuy\x1EBFn \x111\x103ng k\xFD t\x1EA1i \x62\x1EC7nh vi\x1EC7n");
	m_mapLine[_T("II")] = _T("\x110\xFAng tuy\x1EBFn \x111\x103ng k\xFD t\x1EA1i ngo\x1EA1i t\x1EC9nh");
	m_mapLine[_T("III")] = _T("\x42\x1EC7nh nh\xE2n tr\xE1i tuy\x1EBFn");
	UpdateData(false);
	OnListLoadData();

}
void CFMOutpatientInsuranceReport21a1::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndPatientLine.GetDlgCtrlID(), m_szPatientLineKey);
	DDX_Check(pDX, m_wndDrugPTTT.GetDlgCtrlID(), m_bCheckDrug);
	DDX_Check(pDX, m_wndByDischargedDate.GetDlgCtrlID(), m_bByDischargedDate);

}
void CFMOutpatientInsuranceReport21a1::SetDefaultValues(){

	m_nYear = 0;
	m_szReportPeriodKey.Empty();
	m_szPatientLineKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_bCheckDrug=FALSE;
	m_bByDischargedDate = FALSE;

}

int CFMOutpatientInsuranceReport21a1::SetMode(int nMode){
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

/*void CFMOutpatientInsuranceReport21a1::OnYearChange(){
	
} */
/*void CFMOutpatientInsuranceReport21a1::OnYearSetfocus(){
	
} */
/*void CFMOutpatientInsuranceReport21a1::OnYearKillfocus(){
	
} */
int CFMOutpatientInsuranceReport21a1::OnYearCheckValue(){
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
 
void CFMOutpatientInsuranceReport21a1::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
void CFMOutpatientInsuranceReport21a1::OnReportPeriodSelendok(){
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	CString tmpStr;
	CDate dte, date;

	UpdateData(true);

	date.ParseDate(pMF->GetSysDate());
	int nYear = date.GetYear();
	int nMonth = ToInt(m_szReportPeriodKey);

	if (nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/1 00:00"), nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), nYear, nMonth, dte.GetMonthLastDay());
	}
	if (nMonth == 13)
	{
		m_szFromDate.Format(_T("%.4d/1/1 00:00"), nYear);
		tmpStr.Format(_T("%.4d/3/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/3/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 14)
	{
		m_szFromDate.Format(_T("%.4d/4/1 00:00"), nYear);
		tmpStr.Format(_T("%.4d/6/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/6/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 15)
	{
		m_szFromDate.Format(_T("%.4d/7/1 00:00"), nYear);
		tmpStr.Format(_T("%.4d/9/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/9/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 16)
	{
		m_szFromDate.Format(_T("%.4d/10/1 00:00"), nYear);
		tmpStr.Format(_T("%.4d/12/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 17)
	{
		m_szFromDate.Format(_T("%.4d/1/1 00:00"), nYear);
		tmpStr.Format(_T("%.4d/12/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	UpdateData(false);	 
}

/*void CFMOutpatientInsuranceReport21a1::OnReportPeriodSetfocus(){
	
}*/
/*void CFMOutpatientInsuranceReport21a1::OnReportPeriodKillfocus(){
	
}*/
long CFMOutpatientInsuranceReport21a1::OnReportPeriodLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0)
	{
		szWhere.Format(_T(" WHERE hpr_idx = %ld"), ToInt(m_szReportPeriodKey));
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

/*void CFMOutpatientInsuranceReport21a1::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMOutpatientInsuranceReport21a1::OnFromDateChange(){
	
} */
/*void CFMOutpatientInsuranceReport21a1::OnFromDateSetfocus(){
	
} */
/*void CFMOutpatientInsuranceReport21a1::OnFromDateKillfocus(){
	
} */
int CFMOutpatientInsuranceReport21a1::OnFromDateCheckValue(){
	return 0;
}
 
/*void CFMOutpatientInsuranceReport21a1::OnToDateChange(){
	
} */
/*void CFMOutpatientInsuranceReport21a1::OnToDateSetfocus(){
	
} */
/*void CFMOutpatientInsuranceReport21a1::OnToDateKillfocus(){
	
} */
int CFMOutpatientInsuranceReport21a1::OnToDateCheckValue(){
	return 0;
}
 
void CFMOutpatientInsuranceReport21a1::OnPatientLineSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMOutpatientInsuranceReport21a1::OnPatientLineSelendok(){
	 
}
/*void CFMOutpatientInsuranceReport21a1::OnPatientLineSetfocus(){
	
}*/
/*void CFMOutpatientInsuranceReport21a1::OnPatientLineKillfocus(){
	
}*/
long CFMOutpatientInsuranceReport21a1::OnPatientLineLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndPatientLine.IsSearchKey() && ToInt(m_szPatientLineKey) > 0)
	{
		szWhere.Format(_T(" AND ss_code = %d"), ToInt(m_szPatientLineKey));
	}
	m_wndPatientLine.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT ss_code code, ss_desc descr FROM sys_sel WHERE ss_id = 'hms_patient_line' ORDER BY ss_code"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndPatientLine.AddItems(
			rs.GetValue(_T("code")), 
			rs.GetValue(_T("descr")),			
			NULL);
		rs.MoveNext();
	}
	return nCount;	
}

/*void CFMOutpatientInsuranceReport21a1::OnPatientLineAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMOutpatientInsuranceReport21a1::OnDrugPTTTSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}

void CFMOutpatientInsuranceReport21a1::OnByDischargedDateSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMOutpatientInsuranceReport21a1::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString tmpStr, szSQL, szDate, szMoney, szOldGroup, szNewGroup, szOldChap, szNewChap;
	CString szTemp;

	if (!rpt.Init(_T("Reports/MS/F_THONGKETONGHOPDICHVUKYTHUATTHEOQUY.RPT")) )	
	{
		return;
	}
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);

	rpt.GetReportHeader()->SetValue(_T("ObjectGroup"), _T(""));
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), 
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDate::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	//Page Header
	//Report Detail
	int nIdx = 1;
	CString szOldLine, szNewLine, szAmount;
	CReportSection* rptDetail = NULL, *rptGroup = NULL;

	long double nGroupCost = 0, nChapCost = 0, nTotalCost = 0;
	double nCost = 0;
	m_mapChapter[_T("G")] = _T("Ph\xED kh\xE1m");
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("chapter"), szNewChap);
		rs.GetValue(_T("group_id"), szNewGroup);
		if (szNewChap != szOldChap)
		{
			if (nGroupCost > 0)
			{
				rptGroup = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptGroup->SetValue(_T("TotalGroup"), _T("T\x1ED5ng nh\xF3m"));
				tmpStr.Format(_T("%f"), nGroupCost);
				rptGroup->SetValue(_T("s5"), tmpStr);
				nChapCost += nGroupCost;
				nGroupCost = 0;
			}
			if (nChapCost > 0)
			{
				rptGroup = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptGroup->SetValue(_T("TotalGroup"), _T("T\x1ED5ng ") + szOldChap);
				tmpStr.Format(_T("%f"), nChapCost);
				rptGroup->SetValue(_T("s5"), tmpStr);
				nTotalCost += nChapCost;
				nChapCost = 0;
			}
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(0));
			rptGroup->SetValue(_T("Group"), szNewChap);
			m_mapChapter.Lookup(szNewChap, tmpStr);
			rptGroup->SetValue(_T("GroupName"), tmpStr);
			szOldChap = szNewChap;
			szOldGroup.Empty();
		}
		if (szNewGroup != szOldGroup)
		{
			if (nGroupCost > 0)
			{
				rptGroup = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptGroup->SetValue(_T("TotalGroup"), _T("T\x1ED5ng nh\xF3m"));
				tmpStr.Format(_T("%f"), nGroupCost);
				rptGroup->SetValue(_T("s5"), tmpStr);
				nChapCost += nGroupCost;
				nGroupCost = 0;
			}
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(0));
			rptGroup->SetValue(_T("GroupName"), rs.GetValue(_T("group_name")));
			szOldGroup = szNewGroup;
		}
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("item_name")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("qty")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("unit_price")));
		rs.GetValue(_T("amount"), nCost);
		nGroupCost += nCost;
		rptDetail->SetValue(_T("5"), double2str(nCost));
		rs.MoveNext();
	}
	if (nGroupCost > 0)
	{
		rptGroup = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptGroup->SetValue(_T("TotalGroup"), _T("T\x1ED5ng nh\xF3m"));
		tmpStr.Format(_T("%f"), nGroupCost);
		rptGroup->SetValue(_T("s5"), tmpStr);
		nChapCost += nGroupCost;
	}
	if (nChapCost > 0)
	{
		rptGroup = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptGroup->SetValue(_T("TotalGroup"), _T("T\x1ED5ng ") + szOldChap);
		tmpStr.Format(_T("%f"), nChapCost);
		rptGroup->SetValue(_T("s5"), tmpStr);
		nTotalCost += nChapCost;
	}
	if (nTotalCost > 0)
	{
		rptGroup = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptGroup->SetValue(_T("TotalGroup"), _T("T\x1ED5ng \x63\x1ED9ng"));
		tmpStr.Format(_T("%f"), nTotalCost);
		rptGroup->SetValue(_T("s5"), tmpStr);
	}
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
		          tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);		
	MoneyToString(szMoney, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	EndWaitCursor();
	rpt.PrintPreview();
}
 
void CFMOutpatientInsuranceReport21a1::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CGuiMenu menu(this);
	CString tmpStr;
	CRect rect;
	menu.CreatePopupMenu();
	m_wndExport.GetWindowRect(&rect);
	TranslateString(_T("Export 21A Format"), tmpStr);
	menu.AppendMenu(MF_BYPOSITION, 1, tmpStr);
	TranslateString(_T("Export Detail Format"), tmpStr);
	menu.AppendMenu(MF_SEPARATOR);
	menu.AppendMenu(MF_BYPOSITION, 2, tmpStr);
	int nPos = menu.TrackPopupMenu(TPM_RETURNCMD|TPM_RIGHTALIGN|TPM_BOTTOMALIGN|TPM_NONOTIFY, rect.right, rect.top, this);
	switch (nPos)
	{
		case 1:
			OnExport();
			break;
		case 2:
			OnExportDetail();
			break;
	}
}
 
void CFMOutpatientInsuranceReport21a1::OnExport(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);

	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, szOldGroup, szNewGroup, szOldChap, szNewChap, szOldLine, szNewLine, szDate, szMoney;
	CString tmpStr, szTemp;

	int nCol = 0, nRow = 0, nIdx = 1;
	double nCost = 0;
	long double nGroupCost = 0, nChapCost = 0, nLineCost = 0, nTotalCost = 0;

	BeginWaitCursor();
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}

	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 8);
	xls.SetColumnWidth(1, 50);
	xls.SetColumnWidth(2, 8);
	xls.SetColumnWidth(3, 12);
	xls.SetColumnWidth(4, 15);
	
	//Header
	//xls.SetCellMergedColumns(nCol, nRow, 2);
	//xls.SetCellMergedColumns(nCol, nRow + 1, 2);

	//xls.SetCellMergedColumns(nCol + 2, nRow, 5);
	//xls.SetCellMergedColumns(nCol + 2, nRow + 1, 5);

	//xls.SetCellMergedColumns(nCol, nRow + 2, 5);
	//xls.SetCellMergedColumns(nCol, nRow + 3, 5);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 2, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 2, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	TranslateString(_T("General statistics and technical services for patients using medical insurance"), szTemp);
	StringUpper(szTemp, tmpStr);
	xls.SetCellText(nCol, nRow + 2, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);
	
	//Column Header
	CStringArray arrCol, arrFeeGrp, arrTitle;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn \x64\x1ECB\x63h v\x1EE5 k\x1EF9 thu\x1EADt"));
	arrCol.Add(_T("S\x1ED1 l\x1B0\x1EE3ng"));
	arrCol.Add(_T("\x110\x1A1n gi\xE1"));
	arrCol.Add(_T("Th\xE0nh ti\x1EC1n"));

	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol + i, nRow + 4, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10); 
	}
	m_mapChapter[_T("G")] = _T("Ph\xED kh\xE1m");
	nRow = 5;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("patient_line"), szNewLine);
		rs.GetValue(_T("chapter"), szNewChap);
		rs.GetValue(_T("group_id"), szNewGroup);
		if (szNewLine != szOldLine)
		{
			if (nGroupCost > 0)
			{
				xls.SetCellText(1, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
				tmpStr.Format(_T("%f"), nGroupCost);
				xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
				nChapCost += nGroupCost;
				nGroupCost = 0;
				nRow++;
			}
			if (nChapCost > 0)
			{
				xls.SetCellText(1, nRow, _T("T\x1ED5ng ") + szOldChap, 4098, true);
				tmpStr.Format(_T("%f"), nChapCost);
				xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
				nLineCost += nChapCost;
				nChapCost = 0;
				nRow++;
			}
			if (nLineCost > 0)
			{
				xls.SetCellText(1, nRow, _T("T\x1ED5ng ") + szOldLine, 4098, true);
				tmpStr.Format(_T("%f"), nLineCost);
				xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
				nTotalCost += nLineCost;
				nLineCost = 0;
				nRow++;
			}
			m_mapLine.Lookup(szNewLine, szTemp);
			//xls.SetCellMergedColumns(0, nRow, 5);
			tmpStr.Format(_T("%s. %s"), szNewLine, szTemp);
			xls.SetCellText(0, nRow, tmpStr, FMT_TEXT, true);
			szOldLine = szNewLine;
			szOldChap.Empty();
			szOldGroup.Empty();
			nRow++;

		}
		if (szNewChap != szOldChap)
		{
			if (nGroupCost > 0)
			{
				xls.SetCellText(1, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
				tmpStr.Format(_T("%f"), nGroupCost);
				xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
				nChapCost += nGroupCost;
				nGroupCost = 0;
				nRow++;
			}
			if (nChapCost > 0)
			{
				xls.SetCellText(1, nRow, _T("T\x1ED5ng ") + szOldChap, 4098, true);
				tmpStr.Format(_T("%f"), nChapCost);
				xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
				nLineCost += nChapCost;
				nChapCost = 0;
				nRow++;
			}
			m_mapChapter.Lookup(szNewChap, tmpStr);
			//xls.SetCellMergedColumns(0, nRow, 5);
			xls.SetCellText(0, nRow, tmpStr, FMT_TEXT, true);
			szOldChap = szNewChap;
			szOldGroup.Empty();
			nRow++;
		}
		if (szNewGroup != szOldGroup)
		{
			if (nGroupCost > 0){
				xls.SetCellText(1, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
				tmpStr.Format(_T("%f"), nGroupCost);
				xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
				nChapCost += nGroupCost;
				nGroupCost = 0;
				nRow++;
			}
			//xls.SetCellMergedColumns(1, nRow, 4);
			xls.SetCellText(1, nRow, rs.GetValue(_T("group_name")), FMT_TEXT, true);
			szOldGroup = szNewGroup;
			nRow++;
		}
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("item_name")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("qty")), FMT_NUMBER1);
		xls.SetCellText(3, nRow, rs.GetValue(_T("unit_price")), FMT_NUMBER1);
		rs.GetValue(_T("amount"), nCost);
		nGroupCost += nCost;
		xls.SetCellText(4, nRow, double2str(nCost), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nGroupCost > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
		tmpStr.Format(_T("%f"), nGroupCost);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
		nChapCost += nGroupCost;
		nRow++;
	}
	if (nChapCost > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng ") + szOldChap, 4098, true);
		tmpStr.Format(_T("%f"), nChapCost);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
		nLineCost += nChapCost;
		nRow++;
	}
	if (nLineCost > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng ") + szOldLine, 4098, true);
		tmpStr.Format(_T("%f"), nLineCost);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
		nTotalCost += nLineCost;
		nRow++;
	}
	if (nTotalCost > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
		tmpStr.Format(_T("%f"), nTotalCost);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
		nRow++;		
	}
	szMoney.Format(_T("%.0f"), floor(nTotalCost + 0.5));
	MoneyToString(szMoney, tmpStr);
	xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(_T("Ng\xE0y %s th\xE1ng %s n\x103m %s"),
		          tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	xls.SetCellText(3, nRow + 3, szDate, FMT_TEXT);

	EndWaitCursor();
	xls.Save(_T("Exports\\Thong Ke Dich Vu Ky Thuat Theo Quy.xls"));

}

void CFMOutpatientInsuranceReport21a1::OnExportDetail(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);

	CExcel xls;
	CellFormat xlsFormat(&xls);
	xlsFormat.SetCellStyle(FMT_NUMBER1);
	CRecord rs(&pMF->m_db);
	CString szSQL, szOldGroup, szNewGroup, szOldChap, szNewChap, szOldLine, szNewLine, szDate, szMoney;
	CString tmpStr, szTemp;

	int nCol = 0, nRow = 0, nIdx = 1, nIdx1 = 1, nSubIdx1 = 1, nOldLine = 0;
	double nCost = 0;
	long double nGroupCost = 0, nChapCost = 0, nLineCost = 0, nTotalCost = 0;

	BeginWaitCursor();
	szSQL = GetQueryString1();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}

	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 8);
	xls.SetColumnWidth(1, 50);
	xls.SetColumnWidth(2, 8);
	xls.SetColumnWidth(3, 12);
	xls.SetColumnWidth(4, 15);
	
	//Header
	//xls.SetCellMergedColumns(nCol, nRow, 2);
	//xls.SetCellMergedColumns(nCol, nRow + 1, 2);

	//xls.SetCellMergedColumns(nCol + 2, nRow, 5);
	//xls.SetCellMergedColumns(nCol + 2, nRow + 1, 5);

	//xls.SetCellMergedColumns(nCol, nRow + 2, 5);
	//xls.SetCellMergedColumns(nCol, nRow + 3, 5);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 2, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 2, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	TranslateString(_T("General statistics and technical services for patients using medical insurance"), szTemp);
	StringUpper(szTemp, tmpStr);
	xls.SetCellText(nCol, nRow + 2, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);
	
	//Column Header
	CStringArray arrCol, arrFeeGrp, arrTitle;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn \x64\x1ECB\x63h v\x1EE5 k\x1EF9 thu\x1EADt"));
	arrCol.Add(_T("S\x1ED1 l\x1B0\x1EE3ng"));
	arrCol.Add(_T("\x110\x1A1n gi\xE1"));
	arrCol.Add(_T("Th\xE0nh ti\x1EC1n"));

	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol + i, nRow + 4, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10); 
	}
	m_mapChapter[_T("G")] = _T("Ph\xED kh\xE1m");
	nRow = 5;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("chapter"), szNewChap);
		rs.GetValue(_T("group_id"), szNewGroup);
		rs.GetValue(_T("item_name"), szNewLine);
		if (szNewChap != szOldChap)
		{
			if (nLineCost > 0)
			{
				tmpStr.Format(_T("%f"), nLineCost);
				//xls.SetCellText(4, nOldLine, tmpStr, FMT_NUMBER1, true);
				xls.WriteNumber(4, nOldLine, nLineCost, &xlsFormat);
				nGroupCost += nLineCost;
				nLineCost = 0;
				nRow++;
			}
			if (nGroupCost > 0)
			{
				xls.SetCellText(1, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true, 11);
				tmpStr.Format(_T("%f"), nGroupCost);
				//xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true, 11);
				xls.WriteNumber(4, nRow, nGroupCost, &xlsFormat);
				nChapCost += nGroupCost;
				nGroupCost = 0;
				nRow++;
			}
			if (nChapCost > 0)
			{
				xls.SetCellText(1, nRow, _T("T\x1ED5ng ") + szOldChap, 4098, true, 12);
				tmpStr.Format(_T("%f"), nChapCost);
				//xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true, 12);
				xls.WriteNumber(4, nRow, nChapCost, &xlsFormat);
				nTotalCost += nChapCost;
				nChapCost = 0;
				nRow++;
			}
			m_mapChapter.Lookup(szNewChap, tmpStr);
			xls.SetCellText(0, nRow, szNewChap, FMT_TEXT | FMT_RIGHT, true, 12);
			xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true, 12);
			szOldChap = szNewChap;
			szOldGroup.Empty();
			szOldLine.Empty();
			nIdx1 = 1;
			nSubIdx1 = 1;
			nRow++;

		}
		if (szNewGroup != szOldGroup)
		{
			if (nLineCost > 0)
			{
				tmpStr.Format(_T("%f"), nLineCost);
				//xls.SetCellText(4, nOldLine, tmpStr, FMT_NUMBER1, true);
				xls.WriteNumber(4, nOldLine, nLineCost, &xlsFormat);
				nGroupCost += nLineCost;
				nLineCost = 0;
				nRow++;
			}
			if (nGroupCost > 0)
			{
				xls.SetCellText(1, nRow, _T("T\x1ED5ng  nh\xF3m"), 4098, true, 11);
				tmpStr.Format(_T("%f"), nGroupCost);
				//xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true, 11);
				xls.WriteNumber(4, nRow, nGroupCost, &xlsFormat);
				nChapCost += nGroupCost;
				nGroupCost = 0;
				nRow++;
			}
			xls.SetCellText(0, nRow, int2str(nIdx1++), FMT_TEXT | FMT_RIGHT, true, 11);
			xls.SetCellText(1, nRow, rs.GetValue(_T("group_name")), FMT_TEXT, true, 11);
			szOldGroup = szNewGroup;
			szOldLine.Empty();
			nSubIdx1 = 1;
			nRow++;
		}
		if (szNewLine != szOldLine)
		{
			if (nLineCost > 0){
				tmpStr.Format(_T("%f"), nLineCost);
				//xls.SetCellText(4, nOldLine, tmpStr, FMT_NUMBER1, true);
				xls.WriteNumber(4, nOldLine, nLineCost, &xlsFormat);
				nGroupCost += nLineCost;
				nLineCost = 0;
				nRow++;
			}
			tmpStr.Format(_T("%d.%d"), nIdx1-1, nSubIdx1++);
			xls.SetCellText(0, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true);
			xls.SetCellText(1, nRow, rs.GetValue(_T("item_name")), FMT_TEXT, true);
			szOldLine = szNewLine;
			nOldLine = nRow;
			nRow++;
		}
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		tmpStr.Format(_T("[%s] %s - %s | %s"), rs.GetValue(_T("doc_no")), rs.GetValue(_T("patient_name")), 
			rs.GetValue(_T("patient_age")), rs.GetValue(_T("hd_cardno")));
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
		//xls.SetCellText(2, nRow, rs.GetValue(_T("qty")), FMT_NUMBER1);
		rs.GetValue(_T("qty"), nCost);
		xls.WriteNumber(2, nRow, nCost, &xlsFormat);
		rs.GetValue(_T("unit_price"), nCost);
		//xls.SetCellText(3, nRow, rs.GetValue(_T("unit_price")), FMT_NUMBER1);
		xls.WriteNumber(3, nRow, nCost, &xlsFormat);
		rs.GetValue(_T("amount"), nCost);
		nGroupCost += nCost;
		//xls.SetCellText(4, nRow, double2str(nCost) , FMT_NUMBER1);
		xls.WriteNumber(4, nRow, nCost, &xlsFormat);
		xls.SetCellText(5, nRow, CDate::Convert(rs.GetValue(_T("invoice_date")), yyyymmdd, ddmmyyyy));
		xls.SetCellText(6, nRow, rs.GetValue(_T("dept_id")), FMT_TEXT);
		nRow++;
		rs.MoveNext();
	}
	if (nLineCost > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng"), 4098, true);
		tmpStr.Format(_T("%f"), nLineCost);
		//xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true);
		xls.WriteNumber(4, nRow, nLineCost, &xlsFormat);
		nGroupCost += nLineCost;
		nRow++;
	}
	if (nGroupCost > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng  nh\xF3m"), 4098, true, 11);
		tmpStr.Format(_T("%f"), nGroupCost);
		//xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true, 11);
		xls.WriteNumber(4, nRow, nGroupCost, &xlsFormat);
		nChapCost += nGroupCost;
		nRow++;
	}
	if (nChapCost > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng ") + szOldChap, 4098, true, 12);
		tmpStr.Format(_T("%f"), nChapCost);
		//xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true, 12);
		xls.WriteNumber(4, nRow, nChapCost, &xlsFormat);
		nTotalCost += nChapCost;
		nRow++;
	}
	if (nTotalCost > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true, 12);
		tmpStr.Format(_T("%f"), nTotalCost);
		//xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1, true, 12);
		xls.WriteNumber(4, nRow, nTotalCost, &xlsFormat);
		nRow++;		
	}
	szMoney.Format(_T("%.0f"), floor(nTotalCost + 0.5));
	MoneyToString(szMoney, tmpStr);
	xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(_T("Ng\xE0y %s th\xE1ng %s n\x103m %s"),
		          tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	xls.SetCellText(3, nRow + 3, szDate, FMT_TEXT);

	EndWaitCursor();
	xls.Save(_T("Exports\\Chi Tiet Thong Ke Dich Vu Ky Thuat Theo Quy.xls"));

}

void CFMOutpatientInsuranceReport21a1::OnListDblClick(){
	
}
 
void CFMOutpatientInsuranceReport21a1::OnListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
int CFMOutpatientInsuranceReport21a1::OnListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
}
 
long CFMOutpatientInsuranceReport21a1::OnListLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	m_wndList.BeginLoad(); 
	m_wndList.DeleteAllItems(); 
	int nCount = 0;

	szSQL.Format(_T("SELECT sd_id AS ID, sd_name AS Name ") \
		         _T("FROM sys_dept ") \
				 _T("WHERE 1=1 AND sd_type = 'DT'") \
				 _T("ORDER BY sd_id"));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF())
	{ 
		m_wndList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
}


int CFMOutpatientInsuranceReport21a1::OnListCheckAll(){
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (!m_wndList.GetCheck(i))
		{
			m_wndList.SetCheck(i, TRUE);
		}
	}
	return 0;
}


int CFMOutpatientInsuranceReport21a1::OnListUnCheckAll(){
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			m_wndList.SetCheck(i, FALSE);
		}
	}
	return 0;
}


CString CFMOutpatientInsuranceReport21a1::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szSubSQL, szWhere, szSubWhere, szHospitalID;
	CString szLineWhere, szDept, szTable, szFeeCondition;
	szHospitalID = pMF->m_CompanyInfo.sc_id;
	if (!m_szPatientLineKey.IsEmpty())
		if (m_szPatientLineKey == _T("3"))
			szLineWhere.Format(_T(" AND (hd_insline = 'Y' OR hd_object = 12)"));
		else
		{
			szLineWhere.Format(_T(" AND hd_insline = 'N' AND hd_object <> 12"));
			if (m_szPatientLineKey == _T("2"))
				szLineWhere.AppendFormat(_T(" AND substr(hd_cardno, 16, 20) <> '%s'"), szHospitalID);
			else
				szLineWhere.AppendFormat(_T(" AND instr(substr(hd_cardno, 16, 20), '%s') > 0"), szHospitalID);
		}
		
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (!szDept.IsEmpty())
				szDept += _T(",");
			szDept.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
		}
	}
	if (!szLineWhere.IsEmpty())
		szWhere += szLineWhere;
	if (!szDept.IsEmpty())
		szWhere.AppendFormat(_T(" AND i.hfe_deptid in(%s) "), szDept);
	szWhere.AppendFormat(_T(" AND i.hfe_class IN ('A', 'I')"));
	szTable.Format(_T(" LEFT JOIN hms_clinical_record ON (hcr_docno = hd_docno)") \
					_T(" LEFT JOIN hms_treatment_record ON (hcr_docno = htr_docno AND htr_treattime = f.hfe_treattime)"));
	szWhere.AppendFormat(_T(" AND nvl(htr_outpatient, 'X') = 'Y'"));
	if (m_bByDischargedDate)
	{
		szFeeCondition.Format(_T(" AND htrf_acceptedfee IN ('Y', 'P')") \
							_T(" AND htr_dischargedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"),
							m_szFromDate, m_szToDate);
	}		
	else
		szFeeCondition.Format(_T(" AND i.hfe_status = 'P'") \
			_T(" AND i.hfe_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')")
			, m_szFromDate, m_szToDate);
	szWhere.AppendFormat(_T(" AND ho_type IN ('I', 'C') AND f.hfe_discount > 0 %s"), szFeeCondition);
	szSubSQL.Format(_T("SELECT hms_getpatientline(hd_docno) patient_line,") \
		_T("				   Cast ('A' AS NVARCHAR2(1))       chapter, ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   hfe_insprice                     unit_price, ") \
		_T("                   f.hfe_inspaid amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         WHERE     (f.hfe_type = 'T' OR (f.hfe_type = 'P' AND Substr(hfe_group, 1, 2) = 'B1'))") \
		_T("               AND hfe_hitech = 'N' ") \
		_T("		 %s") \
		_T("         UNION ALL ") \
		_T("         SELECT	   hms_getpatientline(hd_docno) ,") \
		_T("				   Cast ('B' AS NVARCHAR2(1)), ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   hfe_insprice                     unit_price, ") \
		_T("                   f.hfe_inspaid amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         WHERE     f.hfe_type = 'P' ") \
		_T("               AND Substr(hfe_group, 1, 2) = 'B2' ") \
		_T("               AND hfe_hitech = 'N' ") \
		_T("		 %s") \
		_T("         UNION ALL ") \
		_T("         SELECT	   hms_getpatientline(hd_docno) ,") \
		_T("				   Cast ('C' AS NVARCHAR2(1)), ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   hfe_insprice                     unit_price, ") \
		_T("                   f.hfe_inspaid amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         WHERE     f.hfe_type = 'P' ") \
		_T("               AND Substr(hfe_group, 1, 2) = 'B3' ") \
		_T("               AND hfe_hitech = 'N' ") \
		_T("		 %s") \
		_T("         UNION ALL ") \
		_T("         SELECT	   hms_getpatientline(hd_docno) ,") \
		_T("				   Cast ('D' AS NVARCHAR2(1)), ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   hfe_insprice                     unit_price, ") \
		_T("                   f.hfe_inspaid amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         WHERE     f.hfe_type = 'O' ") \
		_T("               AND hfe_hitech = 'N' ") \
		_T("		 %s") \
		_T("         UNION ALL ") \
		_T("         SELECT		hms_getpatientline(hd_docno), ") \
		_T("				   Cast ('F' AS NVARCHAR2(1)), ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   hfe_insprice                     unit_price, ") \
		_T("                   f.hfe_inspaid amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         WHERE     hfe_hitech = 'Y' ") \
		_T("		 %s"),szTable, szWhere, szTable, szWhere, szTable, szWhere, szTable, szWhere, szTable, szWhere);
	if (m_bCheckDrug)
		szSubWhere = _T(" AND hpo_feetype <> 'PT'");
	szSubSQL.AppendFormat(
		_T("         UNION ALL ") \
		_T("         SELECT		hms_getpatientline(hd_docno) ,") \
		_T("				   Cast ('E' AS NVARCHAR2(1)), ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   CASE WHEN f.hfe_feegroup = 'OPT' THEN f.hfe_unitprice ELSE f.hfe_insprice END unit_price, ") \
		_T("                   CASE WHEN f.hfe_feegroup = 'OPT' THEN f.hfe_cost ELSE f.hfe_inspaid END amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         LEFT JOIN m_productitem_view ON ( Cast(hfe_itemid AS NVARCHAR2(15)) = product_item_id ) ") \
		_T("		 LEFT JOIN hms_pharmaorder_view ON (hpo_orderid = f.hfe_orderid)") \
		_T("         WHERE     f.hfe_type IN ('D', 'M') %s") \
		_T("               AND hfe_hitech = 'N' ") \
		_T("               AND (product_org_id = 'MA' OR (product_org_id = 'PM' AND product_producttype IN ('A1500', 'A1600')))") \
		_T("               AND hfe_category = 'N' ") \
		_T("		 %s"), szTable, szSubWhere, szWhere);

		szSubSQL.AppendFormat(
		_T("         UNION ALL ") \
		_T("         SELECT	   hms_getpatientline(hd_docno) ,") \
		_T("				   Cast ('G' AS NVARCHAR2(1)), ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   hfe_insprice                     unit_price, ") \
		_T("                   f.hfe_inspaid amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         WHERE     f.hfe_type = 'E' ") \
		_T("               AND hfe_hitech = 'N' ") \
		_T("		 %s"), szTable, szWhere);

	szSQL.Format(_T(" SELECT patient_line, chapter, group_id, group_name, item_name, sum(qty) qty, unit_price, sum(amount) amount ") \
				_T(" FROM   (") \
				_T(" %s) ") \
				_T(" GROUP BY patient_line, chapter, group_id, group_name, item_name, unit_price") \
				_T(" ORDER  BY patient_line, chapter, ") \
				_T("           group_id, ") \
				_T("           item_name "), szSubSQL);
_fmsg(_T("%s"), szSQL);
	return szSQL;	
}
CString CFMOutpatientInsuranceReport21a1::GetQueryString1(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szSubSQL, szWhere, szHospitalID, szTable;
	CString szDept, szLineWhere, szSubWhere, szFeeCondition;
	szHospitalID = pMF->m_CompanyInfo.sc_id;
	if (!m_szPatientLineKey.IsEmpty())
		if (m_szPatientLineKey == _T("3"))
			szLineWhere.Format(_T(" AND (hd_insline = 'Y' OR hd_object = 12)"));
		else
		{
			szLineWhere.Format(_T(" AND hd_insline = 'N' AND hd_object <> 12"));
			if (m_szPatientLineKey == _T("2"))
				szLineWhere.AppendFormat(_T(" AND substr(hd_cardno, 16, 20) <> '%s'"), szHospitalID);
			else
				szLineWhere.AppendFormat(_T(" AND instr(substr(hd_cardno, 16, 20), '%s') > 0"), szHospitalID);
		}
		
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (!szDept.IsEmpty())
				szDept += _T(",");
			szDept.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
		}
	}
	if (!szLineWhere.IsEmpty())
		szWhere += szLineWhere;
	if (!szDept.IsEmpty())
		szWhere.AppendFormat(_T(" AND i.hfe_deptid in(%s) "), szDept);
	szWhere.AppendFormat(_T(" AND i.hfe_class IN ('A', 'I')"));
	szTable.Format(_T(" LEFT JOIN hms_clinical_record ON (hcr_docno = hd_docno)") \
				   _T(" LEFT JOIN hms_treatment_record ON (hcr_docno = htr_docno AND htr_treattime = f.hfe_treattime)"));
	szWhere.AppendFormat(_T(" AND nvl(htr_outpatient, 'X') = 'Y'"));
	if (m_bByDischargedDate)
	{
		szFeeCondition.Format(_T(" AND htrf_acceptedfee IN ('Y', 'P')") \
							_T(" AND htr_dischargedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"),
							m_szFromDate, m_szToDate);
	}		
	else
		szFeeCondition.Format(_T(" AND i.hfe_status = 'P'") \
			_T(" AND i.hfe_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')")
			, m_szFromDate, m_szToDate);
	szWhere.AppendFormat(_T(" AND ho_type IN ('I', 'C') AND f.hfe_discount > 0 %s"), szFeeCondition);
	szSubSQL.Format(_T("SELECT hms_getpatientline(hd_docno) patient_line,") \
		_T("				   hd_patientno,") \
		_T("				   hd_cardno,") \
		_T("				   i.hfe_date invoice_date,") \
		_T("				   i.hfe_deptid dept_id,")
		_T("				   Cast ('A' AS NVARCHAR2(1))       chapter, ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("				   hd_docno doc_no,") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   hfe_insprice                     unit_price, ") \
		_T("                   f.hfe_inspaid amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         WHERE     (f.hfe_type = 'T' OR (f.hfe_type = 'P' AND Substr(hfe_group, 1, 2) = 'B1'))") \
		_T("               AND hfe_hitech = 'N' ") \
		_T("		 %s") \
		_T("         UNION ALL ") \
		_T("         SELECT	   hms_getpatientline(hd_docno) ,") \
		_T("				   hd_patientno,") \
		_T("				   hd_cardno,") \
		_T("				   i.hfe_date invoice_date,") \
		_T("				   i.hfe_deptid dept_id,")
		_T("				   Cast ('B' AS NVARCHAR2(1)), ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("				   hd_docno doc_no,") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   hfe_insprice                     unit_price, ") \
		_T("                   f.hfe_inspaid amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         WHERE     f.hfe_type = 'P' ") \
		_T("               AND Substr(hfe_group, 1, 2) = 'B2' ") \
		_T("               AND hfe_hitech = 'N' ") \
		_T("		 %s") \
		_T("         UNION ALL ") \
		_T("         SELECT	   hms_getpatientline(hd_docno) ,") \
		_T("				   hd_patientno,") \
		_T("				   hd_cardno,") \
		_T("				   i.hfe_date invoice_date,") \
		_T("				   i.hfe_deptid dept_id,")
		_T("				   Cast ('C' AS NVARCHAR2(1)), ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("				   hd_docno doc_no,") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   hfe_insprice                     unit_price, ") \
		_T("                   f.hfe_inspaid amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         WHERE     f.hfe_type = 'P' ") \
		_T("               AND Substr(hfe_group, 1, 2) = 'B3' ") \
		_T("               AND hfe_hitech = 'N' ") \
		_T("		 %s") \
		_T("         UNION ALL ") \
		_T("         SELECT	   hms_getpatientline(hd_docno) ,") \
		_T("				   hd_patientno,") \
		_T("				   hd_cardno,") \
		_T("				   i.hfe_date invoice_date,") \
		_T("				   i.hfe_deptid dept_id,")
		_T("				   Cast ('D' AS NVARCHAR2(1)), ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("				   hd_docno doc_no,") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   hfe_insprice                     unit_price, ") \
		_T("                   f.hfe_inspaid amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         WHERE     f.hfe_type = 'O' ") \
		_T("               AND hfe_hitech = 'N' ") \
		_T("		 %s") \
		_T("         UNION ALL ") \
		_T("         SELECT		hms_getpatientline(hd_docno), ") \
		_T("				   hd_patientno,") \
		_T("				   hd_cardno,") \
		_T("				   i.hfe_date invoice_date,") \
		_T("				   i.hfe_deptid dept_id,")
		_T("				   Cast ('F' AS NVARCHAR2(1)), ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("				   hd_docno doc_no,") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   hfe_insprice                     unit_price, ") \
		_T("                   f.hfe_inspaid amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         WHERE     hfe_hitech = 'Y' ") \
		_T("		 %s"),szTable, szWhere, szTable, szWhere, szTable, szWhere, szTable, szWhere, szTable, szWhere);
	if (m_bCheckDrug)
		szSubWhere = _T(" AND hpo_feetype <> 'PT'");
	szSubSQL.AppendFormat(
		_T("         UNION ALL ") \
		_T("         SELECT		hms_getpatientline(hd_docno) ,") \
		_T("				   hd_patientno,") \
		_T("				   hd_cardno,") \
		_T("				   i.hfe_date invoice_date,") \
		_T("				   i.hfe_deptid dept_id,")
		_T("				   Cast ('E' AS NVARCHAR2(1)), ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("				   hd_docno doc_no,") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   CASE WHEN f.hfe_feegroup = 'OPT' THEN f.hfe_unitprice ELSE f.hfe_insprice END unit_price, ") \
		_T("                   CASE WHEN f.hfe_feegroup = 'OPT' THEN f.hfe_cost ELSE f.hfe_inspaid END amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         LEFT JOIN m_productitem_view ON ( Cast(hfe_itemid AS NVARCHAR2(15)) = product_item_id ) ") \
		_T("		 LEFT JOIN hms_pharmaorder_view ON (hpo_orderid = f.hfe_orderid)") \
		_T("         WHERE     f.hfe_type IN ('D', 'M') ") \
		_T("		 %s") \
		_T("               AND hfe_hitech = 'N' ") \
		_T("               AND (product_org_id = 'MA' OR (product_org_id = 'PM' AND product_producttype IN ('A1500', 'A1600')))") \
		_T("               AND hfe_category = 'N' ") \
		_T("		 %s"), szTable, szSubWhere, szWhere);

		szSubSQL.AppendFormat(
		_T("         UNION ALL ") \
		_T("         SELECT	   hms_getpatientline(hd_docno) ,") \
		_T("				   hd_patientno,") \
		_T("				   hd_cardno,") \
		_T("				   i.hfe_date invoice_date,") \
		_T("				   i.hfe_deptid dept_id,")
		_T("				   Cast ('G' AS NVARCHAR2(1)), ") \
		_T("                   hfe_group                        group_id, ") \
		_T("                   Get_feegroupname(hfe_group)      group_name, ") \
		_T("				   hd_docno doc_no,") \
		_T("                   f.hfe_desc                       item_name, ") \
		_T("                   hfe_quantity                qty, ") \
		_T("                   hfe_insprice                     unit_price, ") \
		_T("                   f.hfe_inspaid amount ") \
		_T("         FROM      hms_fee f ") \
		_T("		 LEFT JOIN hms_doc ON (hd_docno = hfe_docno) ") \
		_T("		 %s") \
		_T("         LEFT JOIN hms_object ON ( hfe_object = ho_id ) ") \
		_T("         LEFT JOIN hms_fee_invoice i ON ( f.hfe_docno = i.hfe_docno ") \
		_T("                                          AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
		_T("         WHERE     f.hfe_type = 'E' ") \
		_T("               AND hfe_hitech = 'N' ") \
		_T("		 %s"), szTable, szWhere);

	szSQL.Format(_T(" SELECT chapter, group_id, group_name, item_name, doc_no, ") \
				_T(" hms_getagebydoc(doc_no) patient_age, hd_cardno, dept_id,") \
				_T(" invoice_date, get_patientname(doc_no) patient_name, sum(qty) qty, unit_price, sum(amount) amount ") \
				_T(" FROM   (") \
				_T(" %s) ") \
				_T(" LEFT JOIN hms_patient ON (hd_patientno = hp_patientno)") \
				_T(" GROUP BY patient_line, chapter, group_id, group_name, item_name, doc_no, hd_cardno, invoice_date, unit_price, dept_id") \
				_T(" ORDER  BY chapter, ") \
				_T("           group_id, ") \
				_T("           item_name, invoice_date, dept_id, doc_no"), szSubSQL);
_fmsg(_T("%s"), szSQL);
	return szSQL;		
}