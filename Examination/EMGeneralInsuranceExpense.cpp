#include "stdafx.h"
#include "EMGeneralInsuranceExpense.h"
#include "HMSMainFrame.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMGeneralInsuranceExpense *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMGeneralInsuranceExpense* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMGeneralInsuranceExpense *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMGeneralInsuranceExpense *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMGeneralInsuranceExpense *)pWnd)->OnToDateCheckValue();
} 
static void _OnRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMGeneralInsuranceExpense* )pWnd)->OnRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnRoomSelendokFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnRoomSelendok();
}
/*static void _OnRoomSetfocusFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnRoomKillfocus();
}*/
/*static void _OnRoomKillfocusFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnRoomKillfocus();
}*/
static long _OnRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMGeneralInsuranceExpense *)pWnd)->OnRoomLoadData();
}
/*static void _OnRoomAddNewFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnRoomAddNew();
}*/
static void _OnDoctorSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMGeneralInsuranceExpense* )pWnd)->OnDoctorSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDoctorSelendokFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnDoctorSelendok();
}
/*static void _OnDoctorSetfocusFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnDoctorKillfocus();
}*/
/*static void _OnDoctorKillfocusFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnDoctorKillfocus();
}*/
static long _OnDoctorLoadDataFnc(CWnd *pWnd){
	return ((CEMGeneralInsuranceExpense *)pWnd)->OnDoctorLoadData();
}
/*static void _OnDoctorAddNewFnc(CWnd *pWnd){
	((CEMGeneralInsuranceExpense *)pWnd)->OnDoctorAddNew();
}*/
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMGeneralInsuranceExpense *pVw = (CEMGeneralInsuranceExpense *)pWnd;
	pVw->OnExportSelect();
} 
CEMGeneralInsuranceExpense::CEMGeneralInsuranceExpense(CWnd *pWnd){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CEMGeneralInsuranceExpense::~CEMGeneralInsuranceExpense(){
}
void CEMGeneralInsuranceExpense::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 415, 150);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 205, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 215, 30, 295, 55);
	m_wndReportPeriod.Create(this,300, 30, 410, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 205, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 215, 60, 295, 85);
	m_wndToDate.Create(this,300, 60, 410, 85); 
	m_wndRoomLabel.Create(this, _T("Room"), 10, 90, 90, 115);
	m_wndRoom.Create(this,95, 90, 410, 115); 
	m_wndDoctorLabel.Create(this, _T("Doctor"), 10, 120, 90, 145);
	m_wndDoctor.Create(this,95, 120, 410, 145); 
	m_wndExport.Create(this, _T("&Export"), 335, 155, 415, 180);

}
void CEMGeneralInsuranceExpense::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndRoom.SetCheckValue(true);
	m_wndRoom.LimitText(35);
	m_wndDoctor.SetCheckValue(true);
	m_wndDoctor.LimitText(35);

	m_wndRoom.InsertColumn(0, _T("ID"), CFMT_TEXT, 30);
	m_wndRoom.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

	m_wndDoctor.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndDoctor.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

}
void CEMGeneralInsuranceExpense::OnSetWindowEvents(){
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
	m_wndRoom.SetEvent(WE_SELENDOK, _OnRoomSelendokFnc);
	//m_wndRoom.SetEvent(WE_SETFOCUS, _OnRoomSetfocusFnc);
	//m_wndRoom.SetEvent(WE_KILLFOCUS, _OnRoomKillfocusFnc);
	m_wndRoom.SetEvent(WE_SELCHANGE, _OnRoomSelectChangeFnc);
	m_wndRoom.SetEvent(WE_LOADDATA, _OnRoomLoadDataFnc);
	//m_wndRoom.SetEvent(WE_ADDNEW, _OnRoomAddNewFnc);
	m_wndDoctor.SetEvent(WE_SELENDOK, _OnDoctorSelendokFnc);
	//m_wndDoctor.SetEvent(WE_SETFOCUS, _OnDoctorSetfocusFnc);
	//m_wndDoctor.SetEvent(WE_KILLFOCUS, _OnDoctorKillfocusFnc);
	m_wndDoctor.SetEvent(WE_SELCHANGE, _OnDoctorSelectChangeFnc);
	m_wndDoctor.SetEvent(WE_LOADDATA, _OnDoctorLoadDataFnc);
	//m_wndDoctor.SetEvent(WE_ADDNEW, _OnDoctorAddNewFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), str2int(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CEMGeneralInsuranceExpense::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndRoom.GetDlgCtrlID(), m_szRoomKey);
	DDX_TextEx(pDX, m_wndDoctor.GetDlgCtrlID(), m_szDoctorKey);

}
void CEMGeneralInsuranceExpense::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szRoomKey.Empty();
	m_szDoctorKey.Empty();

}
int CEMGeneralInsuranceExpense::SetMode(int nMode){
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
/*void CEMGeneralInsuranceExpense::OnYearChange(){
	
} */
/*void CEMGeneralInsuranceExpense::OnYearSetfocus(){
	
} */
/*void CEMGeneralInsuranceExpense::OnYearKillfocus(){
	
} */
int CEMGeneralInsuranceExpense::OnYearCheckValue(){
	return 0;
} 
void CEMGeneralInsuranceExpense::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMGeneralInsuranceExpense::OnReportPeriodSelendok(){
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
/*void CEMGeneralInsuranceExpense::OnReportPeriodSetfocus(){
	
}*/
/*void CEMGeneralInsuranceExpense::OnReportPeriodKillfocus(){
	
}*/
long CEMGeneralInsuranceExpense::OnReportPeriodLoadData(){
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
/*void CEMGeneralInsuranceExpense::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMGeneralInsuranceExpense::OnFromDateChange(){
	
} */
/*void CEMGeneralInsuranceExpense::OnFromDateSetfocus(){
	
} */
/*void CEMGeneralInsuranceExpense::OnFromDateKillfocus(){
	
} */
int CEMGeneralInsuranceExpense::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMGeneralInsuranceExpense::OnToDateChange(){
	
} */
/*void CEMGeneralInsuranceExpense::OnToDateSetfocus(){
	
} */
/*void CEMGeneralInsuranceExpense::OnToDateKillfocus(){
	
} */
int CEMGeneralInsuranceExpense::OnToDateCheckValue(){
	return 0;
} 
void CEMGeneralInsuranceExpense::OnRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMGeneralInsuranceExpense::OnRoomSelendok(){
	 
}
/*void CEMGeneralInsuranceExpense::OnRoomSetfocus(){
	
}*/
/*void CEMGeneralInsuranceExpense::OnRoomKillfocus(){
	
}*/
long CEMGeneralInsuranceExpense::OnRoomLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndRoom.IsSearchKey() && !m_szRoomKey.IsEmpty()){
	 szWhere.Format(_T(" and hrl_id='%s' "), m_szRoomKey);
	}
	m_wndRoom.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT hrl_id as id, hrl_name as name FROM hms_roomlist WHERE 1=1 AND hrl_deptid = '%s' %s ORDER BY id "), pMF->GetCurrentDepartmentID(), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndRoom.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMGeneralInsuranceExpense::OnRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMGeneralInsuranceExpense::OnDoctorSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMGeneralInsuranceExpense::OnDoctorSelendok(){
	 
}
/*void CEMGeneralInsuranceExpense::OnDoctorSetfocus(){
	
}*/
/*void CEMGeneralInsuranceExpense::OnDoctorKillfocus(){
	
}*/
long CEMGeneralInsuranceExpense::OnDoctorLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDoctor.IsSearchKey() && !m_szDoctorKey.IsEmpty()){
		szWhere.Format(_T(" and id='%s' "), m_szDoctorKey);
	}
	m_wndDoctor.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as id, su_name as name FROM sys_user WHERE 1=1 AND su_groupid = 'D' AND su_deptid = '%s' %s ORDER BY id "), pMF->GetCurrentDepartmentID(), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDoctor.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMGeneralInsuranceExpense::OnDoctorAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMGeneralInsuranceExpense::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CExcel xls;
	CString szSQL, tmpStr;
	CStringArray arrField;
	int nCol = 0, nRow = 0, nOldRow = 0;
	double nTmp = 0;
	long double nTotal[15];
	CString szPreviousDept, szCurrentDept;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	for (int i = 0; i < 15; i++)
		nTotal[i] = 0;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(1, 25);
	for (int i = 2; i < 15; i++)
		xls.SetColumnWidth(i, 15);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 10);
	xls.SetCellMergedColumns(0, 3, 10);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	xls.SetCellText(0, 2, _T("Th\x1ED1ng k\xEA t\x1ED5ng \x63hi ph\xED \x63\x1EE7\x61 \x42N \x42HYT"), 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098);
	xls.SetCellMergedRows(0, 4, 2);
	xls.SetCellMergedRows(1, 4, 2);
	xls.SetCellText(0, 4, _T("Ph\xF2ng"), 4098, true);
	xls.SetCellText(1, 4, _T("\x42\xE1\x63 s\x1EF9"), 4098, true);
	xls.SetCellMergedColumns(2, 4, 5);
	xls.SetCellText(2, 4, _T("T\x1ED5ng s\x1ED1 \x62\x1EA3o hi\x1EC3m"), 4098, true);
	xls.SetCellText(2, 5, _T("S\x1ED1 \x42N"), 4098);
	xls.SetCellText(3, 5, _T("S\x1ED1 ti\x1EC1n XN"), 4098);
	xls.SetCellText(4, 5, _T("S\x1ED1 ti\x1EC1n thu\x1ED1\x63"), 4098);
	xls.SetCellText(5, 5, _T("T\x1ED5ng ti\x1EC1n"), 4098);
	xls.SetCellText(6, 5, _T("TB/BN"), 4098);
	xls.SetCellMergedColumns(7, 4, 5);
	xls.SetCellText(7, 4, _T("\x42H \x63huy\x1EC3n tuy\x1EBFn + tr\xE1i tuy\x1EBFn(tuy\x1EBFn \x32)"), 4098, true);
	xls.SetCellText(7, 5, _T("S\x1ED1 \x42N"), 4098);
	xls.SetCellText(8, 5, _T("S\x1ED1 ti\x1EC1n XN"), 4098);
	xls.SetCellText(9, 5, _T("S\x1ED1 ti\x1EC1n thu\x1ED1\x63"), 4098);
	xls.SetCellText(10, 5, _T("T\x1ED5ng ti\x1EC1n"), 4098);
	xls.SetCellText(11, 5, _T("TB/BN"), 4098);
	xls.SetCellMergedColumns(12, 4, 5);
	xls.SetCellText(12, 4, _T("\x42H t\x1EA1i \x63h\x1ED7(\x111\x103ng k\xFD t\x1EA1i vi\x1EC7n:tuy\x1EBFn \x31)"), 4098, true);
	xls.SetCellText(12, 5, _T("S\x1ED1 \x42N"), 4098);
	xls.SetCellText(13, 5, _T("S\x1ED1 ti\x1EC1n XN"), 4098);
	xls.SetCellText(14, 5, _T("S\x1ED1 ti\x1EC1n thu\x1ED1\x63"), 4098);
	xls.SetCellText(15, 5, _T("T\x1ED5ng ti\x1EC1n"), 4098);
	xls.SetCellText(16, 5, _T("TB/BN"), 4098);

	nRow = 6;
	nOldRow = 6;
	arrField.Add(_T("insqty"));
	arrField.Add(_T("testamt"));
	arrField.Add(_T("drugamt"));
	arrField.Add(_T("totalamt"));
	arrField.Add(_T("insrate"));
	arrField.Add(_T("line2qty"));
	arrField.Add(_T("line2testamt"));
	arrField.Add(_T("line2drugamt"));
	arrField.Add(_T("line2amt"));
	arrField.Add(_T("line2rate"));
	arrField.Add(_T("line1qty"));
	arrField.Add(_T("line1testamt"));
	arrField.Add(_T("line1drugamt"));
	arrField.Add(_T("line1qty"));
	arrField.Add(_T("line1rate"));

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("roomid"), szCurrentDept);
		if (szCurrentDept != szPreviousDept)
		{
			xls.SetCellMergedRows(0, nOldRow, nRow - nOldRow);
			xls.SetCellText(0, nRow, szCurrentDept, FMT_TEXT);
			nOldRow = nRow;
			szPreviousDept = szCurrentDept;
		}
		xls.SetCellText(1, nRow, rs.GetValue(_T("doctor")), FMT_TEXT);
		for (int i = 0; i < arrField.GetCount(); i++)
		{
			rs.GetValue(arrField.GetAt(i), nTmp);
			xls.SetCellText(nCol + i + 2, nRow, double2str(nTmp), FMT_NUMBER1);
			nTotal[i] += nTmp;
		}
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellText(1, nRow, _T("T\x1ED5ng"), 4098, true);
	for (int i = 0; i < 15;i++)
	{
		tmpStr.Format(_T("%f"), nTotal[i]);
		xls.SetCellText(nCol + i + 2, nRow, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\Thong ke tong chi phi cua BN BHYT.xls"));
} 
CString CEMGeneralInsuranceExpense::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szTestWhere, szPharmaWhere;
	szTestWhere.Format(_T(" AND hpc_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szPharmaWhere.Format(_T(" AND hpo_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szRoomKey.IsEmpty())
	{
		szTestWhere.AppendFormat(_T(" AND hpc_roomid = %s"), m_szRoomKey);
		szPharmaWhere.AppendFormat(_T(" AND hpo_roomid = %s"), m_szRoomKey);
	}
	if (!m_szDoctorKey.IsEmpty())
	{
		szTestWhere.AppendFormat(_T(" AND hpc_doctor = '%s'"), m_szDoctorKey);
		szPharmaWhere.AppendFormat(_T(" AND hpo_doctor = '%s'"), m_szDoctorKey);
	}
	szSQL.Format(_T(" SELECT hpc_roomid, hrl_name roomid, ") \
				_T("        get_username(hpc_doctor) doctor, ") \
				_T("        SUM(insqty)                                insqty, ") \
				_T("        SUM(testamount)                            testamt, ") \
				_T("        SUM(drugamount)                            drugamt, ") \
				_T("        SUM(testamount + drugamount)               totalamt, ") \
				_T("        SUM(testamount + drugamount) / SUM(insqty) insrate, ") \
				_T("        SUM(line2qty)                              line2qty, ") \
				_T("        SUM(line2testamount)                       line2testamt, ") \
				_T("        SUM(line2drugamount)                       line2drugamt, ") \
				_T("        SUM(line2testamount + line2drugamount)     line2amt, ") \
				_T("        CASE WHEN SUM(line2qty) > 0 THEN SUM(line2testamount + line2drugamount) / SUM(line2qty) ") \
				_T("        ELSE 0 ") \
				_T("        END                                        line2rate, ") \
				_T("        SUM(line1qty)                              line1qty, ") \
				_T("        SUM(line1testamount)                       line1testamt, ") \
				_T("        SUM(line1drugamount)                       line1drugamt, ") \
				_T("        SUM(line1testamount + line1drugamount)     line1amt, ") \
				_T("        CASE WHEN SUM(line1qty) > 0 THEN SUM(line1testamount + line1drugamount) / SUM(line1qty) ") \
				_T("        ELSE 0 ") \
				_T("        END                                        line1rate ") \
				_T(" FROM   (SELECT    hpc_roomid, ") \
				_T("                   hpc_doctor, ") \
				_T("                   1   insqty, ") \
				_T("                   testamount, ") \
				_T("                   drugamount, ") \
				_T("                   CASE WHEN hc_line <> 1 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END line2qty, ") \
				_T("                   CASE WHEN hc_line = 1 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END line1qty, ") \
				_T("                   CASE WHEN hc_line <> 1 THEN testamount ") \
				_T("                   ELSE 0 ") \
				_T("                   END line2testamount, ") \
				_T("                   CASE WHEN hc_line = 1 THEN testamount ") \
				_T("                   ELSE 0 ") \
				_T("                   END line1testamount, ") \
				_T("                   CASE WHEN hc_line <> 1 THEN drugamount ") \
				_T("                   ELSE 0 ") \
				_T("                   END line2drugamount, ") \
				_T("                   CASE WHEN hc_line = 1 THEN drugamount ") \
				_T("                   ELSE 0 ") \
				_T("                   END line1drugamount ") \
				_T("         FROM      (SELECT hpc_docno, ") \
				_T("                           hpc_doctor, ") \
				_T("                           hpc_roomid, ") \
				_T("                           SUM(testamount) testamount, ") \
				_T("                           SUM(drugamount) drugamount ") \
				_T("                    FROM   (SELECT    hpc_docno, ") \
				_T("                                      hpc_doctor, ") \
				_T("                                      hpc_roomid, ") \
				_T("                                      SUM(Coalesce(hfe_cost, 0)) testamount, ") \
				_T("                                      0                          drugamount ") \
				_T("                            FROM      hms_testorder ") \
				_T("                            LEFT JOIN hms_fee ON ( hfe_docno = hpc_docno ") \
				_T("                                                   AND hpc_orderid = hfe_orderid ) ") \
				_T("                            WHERE     hpc_status IN ( 'S', 'T' ) ") \
				_T("                                  AND hfe_class = 'E' ") \
				_T("                                  AND hpc_deptid = 'C1.1' ") \
				_T("                                  AND hpc_roomid > 0 %s") \
				_T("                            GROUP     BY hpc_docno, ") \
				_T("                                         hpc_doctor, ") \
				_T("                                         hpc_roomid ") \
				_T("                            UNION ALL ") \
				_T("                            SELECT    hpo_docno, ") \
				_T("                                      hpo_doctor, ") \
				_T("                                      hpo_roomid, ") \
				_T("                                      0, ") \
				_T("                                      SUM(hpol_unitprice * hpol_qtyorder) ") \
				_T("                            FROM      hms_pharmaorder ") \
				_T("                            LEFT JOIN hms_pharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T("                            WHERE     hpo_orderstatus IN ( 'S', 'T' ) ") \
				_T("                                  AND hfe_refstatus = 'P' ") \
				_T("                                  AND hpo_ordertype = 'P' ") \
				_T("                                  AND hpo_deptid = 'C1.1' ") \
				_T("                                  AND hpo_roomid > 0 %s") \
				_T("                            GROUP     BY hpo_docno, ") \
				_T("                                         hpo_doctor, ") \
				_T("                                         hpo_roomid) tbl0 ") \
				_T("                    GROUP  BY hpc_docno, ") \
				_T("                              hpc_doctor, ") \
				_T("                              hpc_roomid) tbldoc ") \
				_T("         LEFT JOIN hms_doc ON ( hd_docno = hpc_docno ) ") \
				_T("         LEFT JOIN hms_object ON ( ho_id = hd_object ) ") \
				_T("         LEFT JOIN hms_card ON ( hd_patientno = hc_patientno ") \
				_T("                                 AND hd_cardidx = hc_idx ) ") \
				_T("         WHERE     ho_type = 'I') tblfilter ") \
				_T(" LEFT JOIN hms_roomlist ON (hrl_deptid = '%s' AND hrl_id = hpc_roomid)") \
				_T(" GROUP  BY hpc_roomid, hrl_name,") \
				_T("           hpc_doctor ") \
				_T(" ORDER  BY cast(substr(hrl_name, 1, 2) as integer), ") \
				_T("           hpc_doctor "), szTestWhere, szPharmaWhere, pMF->GetCurrentDepartmentID());
	return szSQL;
}