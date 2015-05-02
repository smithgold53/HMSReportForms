#include "stdafx.h"
#include "EMGeneralTestIncome.h"
#include "HMSMainFrame.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMGeneralTestIncome *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMGeneralTestIncome* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMGeneralTestIncome *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMGeneralTestIncome *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMGeneralTestIncome *)pWnd)->OnToDateCheckValue();
} 
static void _OnRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMGeneralTestIncome* )pWnd)->OnRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnRoomSelendokFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnRoomSelendok();
}
/*static void _OnRoomSetfocusFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnRoomKillfocus();
}*/
/*static void _OnRoomKillfocusFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnRoomKillfocus();
}*/
static long _OnRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMGeneralTestIncome *)pWnd)->OnRoomLoadData();
}
/*static void _OnRoomAddNewFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnRoomAddNew();
}*/
static void _OnDoctorSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMGeneralTestIncome* )pWnd)->OnDoctorSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDoctorSelendokFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnDoctorSelendok();
}
/*static void _OnDoctorSetfocusFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnDoctorKillfocus();
}*/
/*static void _OnDoctorKillfocusFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnDoctorKillfocus();
}*/
static long _OnDoctorLoadDataFnc(CWnd *pWnd){
	return ((CEMGeneralTestIncome *)pWnd)->OnDoctorLoadData();
}
/*static void _OnDoctorAddNewFnc(CWnd *pWnd){
	((CEMGeneralTestIncome *)pWnd)->OnDoctorAddNew();
}*/
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMGeneralTestIncome *pVw = (CEMGeneralTestIncome *)pWnd;
	pVw->OnExportSelect();
} 
CEMGeneralTestIncome::CEMGeneralTestIncome(CWnd *pWnd){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CEMGeneralTestIncome::~CEMGeneralTestIncome(){
}
void CEMGeneralTestIncome::OnCreateComponents(){
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
void CEMGeneralTestIncome::OnInitializeComponents(){
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
void CEMGeneralTestIncome::OnSetWindowEvents(){
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
void CEMGeneralTestIncome::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndRoom.GetDlgCtrlID(), m_szRoomKey);
	DDX_TextEx(pDX, m_wndDoctor.GetDlgCtrlID(), m_szDoctorKey);

}
void CEMGeneralTestIncome::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szRoomKey.Empty();
	m_szDoctorKey.Empty();

}
int CEMGeneralTestIncome::SetMode(int nMode){
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
/*void CEMGeneralTestIncome::OnYearChange(){
	
} */
/*void CEMGeneralTestIncome::OnYearSetfocus(){
	
} */
/*void CEMGeneralTestIncome::OnYearKillfocus(){
	
} */
int CEMGeneralTestIncome::OnYearCheckValue(){
	return 0;
} 
void CEMGeneralTestIncome::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMGeneralTestIncome::OnReportPeriodSelendok(){
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
/*void CEMGeneralTestIncome::OnReportPeriodSetfocus(){
	
}*/
/*void CEMGeneralTestIncome::OnReportPeriodKillfocus(){
	
}*/
long CEMGeneralTestIncome::OnReportPeriodLoadData(){
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
/*void CEMGeneralTestIncome::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMGeneralTestIncome::OnFromDateChange(){
	
} */
/*void CEMGeneralTestIncome::OnFromDateSetfocus(){
	
} */
/*void CEMGeneralTestIncome::OnFromDateKillfocus(){
	
} */
int CEMGeneralTestIncome::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMGeneralTestIncome::OnToDateChange(){
	
} */
/*void CEMGeneralTestIncome::OnToDateSetfocus(){
	
} */
/*void CEMGeneralTestIncome::OnToDateKillfocus(){
	
} */
int CEMGeneralTestIncome::OnToDateCheckValue(){
	return 0;
} 
void CEMGeneralTestIncome::OnRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMGeneralTestIncome::OnRoomSelendok(){
	 
}
/*void CEMGeneralTestIncome::OnRoomSetfocus(){
	
}*/
/*void CEMGeneralTestIncome::OnRoomKillfocus(){
	
}*/
long CEMGeneralTestIncome::OnRoomLoadData(){
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
/*void CEMGeneralTestIncome::OnRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMGeneralTestIncome::OnDoctorSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMGeneralTestIncome::OnDoctorSelendok(){
	 
}
/*void CEMGeneralTestIncome::OnDoctorSetfocus(){
	
}*/
/*void CEMGeneralTestIncome::OnDoctorKillfocus(){
	
}*/
long CEMGeneralTestIncome::OnDoctorLoadData(){
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

/*void CEMGeneralTestIncome::OnDoctorAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */

void CEMGeneralTestIncome::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr;
	CStringArray arrField;
	int nCol = 0, nRow = 0, nOldRow = 0;
	double nTemp = 0;
	long double nTotal[18];
	CString szPreviousDept, szCurrentDept;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	for (int i = 0; i < 18; i++)
		nTotal[i] = 0;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(6, 15);
	xls.SetColumnWidth(7, 10);
	xls.SetColumnWidth(9, 15);
	xls.SetColumnWidth(10, 10);
	xls.SetColumnWidth(12, 15);
	xls.SetColumnWidth(13, 10);
	xls.SetColumnWidth(15, 15);
	xls.SetColumnWidth(16, 10);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 10);
	xls.SetCellMergedColumns(0, 3, 10);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	xls.SetCellText(0, 2, _T("Th\x1ED1ng k\xEA s\x1ED1 ti\x1EC1n \x42N \x78\xE9t nghi\x1EC7m"), 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098);
	xls.SetCellMergedRows(0, 4, 2);
	xls.SetCellMergedRows(1, 4, 2);
	xls.SetCellText(0, 4, _T("Ph\xF2ng"), 4098, true);
	xls.SetCellText(1, 4, _T("\x42\xE1\x63 s\x1EF9"), 4098, true);
	xls.SetCellMergedColumns(2, 4, 3);
	xls.SetCellText(2, 4, _T("T\x1ED5ng"), 4098, true);
	xls.SetCellText(2, 5, _T("S\x1ED1 \x42N"), 4098);
	xls.SetCellText(3, 5, _T("S\x1ED1 ti\x1EC1n"), 4098);
	xls.SetCellText(4, 5, _T("TB/BN"), 4098);
	xls.SetCellMergedColumns(5, 4, 3);
	xls.SetCellText(5, 4, _T("\x44\x1ECB\x63h v\x1EE5"), 4098, true);
	xls.SetCellText(5, 5, _T("S\x1ED1 \x42N"), 4098);
	xls.SetCellText(6, 5, _T("S\x1ED1 ti\x1EC1n"), 4098);
	xls.SetCellText(7, 5, _T("TB/BN"), 4098);
	xls.SetCellMergedColumns(8, 4, 3);
	xls.SetCellText(8, 4, _T("BH \x63huy\x1EC3n+tr\xE1i tuy\x1EBFn"), 4098, true);
	xls.SetCellText(8, 5, _T("S\x1ED1 \x42N"), 4098);
	xls.SetCellText(9, 5, _T("S\x1ED1 ti\x1EC1n"), 4098);
	xls.SetCellText(10, 5, _T("TB/BN"), 4098);
	xls.SetCellMergedColumns(11, 4, 3);
	xls.SetCellText(11, 4, _T("BH t\x1EA1i \x63h\x1ED7"), 4098, true);
	xls.SetCellText(11, 5, _T("S\x1ED1 \x42N"), 4098);
	xls.SetCellText(12, 5, _T("S\x1ED1 ti\x1EC1n"), 4098);
	xls.SetCellText(13, 5, _T("TB/BN"), 4098);
	xls.SetCellMergedColumns(14, 4, 3);
	xls.SetCellText(14, 4, _T("Qu\xE2n"), 4098, true);
	xls.SetCellText(14, 5, _T("S\x1ED1 \x42N"), 4098);
	xls.SetCellText(15, 5, _T("S\x1ED1 ti\x1EC1n"), 4098);
	xls.SetCellText(16, 5, _T("TB/BN"), 4098);
	xls.SetCellMergedColumns(17, 4, 3);
	xls.SetCellText(17, 4, _T("Kh\xE1\x63"), 4098, true);
	xls.SetCellText(17, 5, _T("S\x1ED1 \x42N"), 4098);
	xls.SetCellText(18, 5, _T("S\x1ED1 ti\x1EC1n"), 4098);
	xls.SetCellText(19, 5, _T("TB/BN"), 4098);

	nRow = 6;
	nOldRow = 6;
	arrField.Add(_T("patientqty"));
	arrField.Add(_T("amount"));
	arrField.Add(_T("patientrate"));
	arrField.Add(_T("serviceqty"));
	arrField.Add(_T("serviceamt"));
	arrField.Add(_T("servicerate"));
	arrField.Add(_T("line2qty"));
	arrField.Add(_T("line2amt"));
	arrField.Add(_T("line2rate"));
	arrField.Add(_T("line1qty"));
	arrField.Add(_T("line1amt"));
	arrField.Add(_T("line1rate"));
	arrField.Add(_T("solqty"));
	arrField.Add(_T("solamt"));
	arrField.Add(_T("solrate"));
	arrField.Add(_T("otherqty"));
	arrField.Add(_T("otheramt"));
	arrField.Add(_T("otherrate"));
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
			tmpStr = arrField.GetAt(i);
			rs.GetValue(tmpStr, nTemp);
			xls.SetCellText(nCol + i + 2, nRow, rs.GetValue(tmpStr), FMT_NUMBER1);
			if (tmpStr.Find(_T("rate")) == -1)
				nTotal[i] += nTemp;
		}
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellText(1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
	for (int i = 0; i < 18; i++)
	{
		tmpStr.Format(_T("%f"), nTotal[i]);
		xls.SetCellText(i+2, nRow, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\Thong ke so tien BN xet nghiem.xls"));
} 

CString CEMGeneralTestIncome::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szWhere, szWhere1;
	szWhere.Format(_T(" AND hpc_deptid = '%s' AND hpc_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
					, pMF->GetCurrentDepartmentID(), m_szFromDate, m_szToDate);
	szWhere1.Format(_T(" AND he_deptid = '%s' AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
					, pMF->GetCurrentDepartmentID(), m_szFromDate, m_szToDate);
	if (!m_szRoomKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hpc_roomid = %s"), m_szRoomKey);
		szWhere1.AppendFormat(_T(" AND he_roomid = %s"), m_szRoomKey);
	}
	if (!m_szDoctorKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hpc_doctor = '%s'"), m_szDoctorKey);
		szWhere1.AppendFormat(_T(" AND he_doctor = '%s'"), m_szDoctorKey);
	}
	szSQL.Format(_T(" SELECT hrl_name roomid, hpc_roomid,") \
				_T("        Get_username(hpc_doctor) doctor, ") \
				_T("        SUM(patientqty)          patientqty, ") \
				_T("        SUM(amount)              amount, ") \
				_T("        CASE WHEN SUM(patientqty) > 0 THEN SUM(amount) / SUM(patientqty) ") \
				_T("        ELSE 0 ") \
				_T("        END                      patientrate, ") \
				_T("        SUM(serviceqty)          serviceqty, ") \
				_T("        SUM(serviceamt)          serviceamt, ") \
				_T("        CASE WHEN SUM(serviceqty) > 0 THEN SUM(serviceamt) / SUM(serviceqty) ") \
				_T("        ELSE 0 ") \
				_T("        END                      servicerate, ") \
				_T("        SUM(line2qty)            line2qty, ") \
				_T("        SUM(line2amt)            line2amt, ") \
				_T("        CASE WHEN SUM (line2qty) > 0 THEN SUM(line2amt) / SUM(line2qty) ") \
				_T("        ELSE 0 ") \
				_T("        END                      line2rate, ") \
				_T("        SUM(line1qty)            line1qty, ") \
				_T("        SUM(line1amt)            line1amt, ") \
				_T("        CASE WHEN SUM(line1qty) > 0 THEN SUM(line1amt) / SUM(line1qty) ") \
				_T("        ELSE 0 ") \
				_T("        END                      line1rate, ") \
				_T("        SUM(solqty)              solqty, ") \
				_T("        SUM(solamt)              solamt, ") \
				_T("        CASE WHEN SUM(solqty) > 0 THEN SUM(solamt) / SUM(solqty) ") \
				_T("        ELSE 0 ") \
				_T("        END                      solrate, ") \
				_T("        SUM(otherqty)            otherqty, ") \
				_T("        SUM(otheramt)            otheramt, ") \
				_T("        CASE WHEN SUM(otherqty) > 0 THEN SUM(otheramt) / SUM(otherqty) ") \
				_T("        ELSE 0 ") \
				_T("        END                      otherrate ") \
				_T(" FROM   (SELECT    hpc_roomid, ") \
				_T("                   hpc_doctor, ") \
				_T("                   0 patientqty, ") \
				_T("                   CASE WHEN ho_type IN ('I', 'C') THEN SUM(hfe_quantity*hfe_insprice)") \
				_T("						WHEN ho_type IN ('D', 'P') THEN SUM(hfe_quantity*hfe_polprice)") \
				_T("						ELSE SUM(f.hfe_cost) END amount,") \
				_T("                   0 serviceqty, ") \
				_T("                   CASE WHEN hd_object = 7 THEN SUM(f.hfe_cost) ") \
				_T("                   ELSE 0 ") \
				_T("                   END                   serviceamt, ") \
				_T("                   0 line2qty, ") \
				_T("                   CASE WHEN ho_type IN ('I', 'C') AND hc_line <> 1 AND hd_object <> 12 THEN SUM(hfe_quantity*hfe_insprice) ") \
				_T("                   ELSE 0 ") \
				_T("                   END                   line2amt, ") \
				_T("                   0 line1qty, ") \
				_T("                   CASE WHEN ho_type IN ('I', 'C') AND hc_line = 1 AND hd_object <> 12 THEN SUM(hfe_quantity*hfe_insprice) ") \
				_T("                   ELSE 0 ") \
				_T("                   END                   line1amt, ") \
				_T("                   0 solqty, ") \
				_T("                   CASE WHEN hd_object = 1 THEN SUM(hfe_quantity*hfe_polprice) ") \
				_T("                   ELSE 0 ") \
				_T("                   END                   solamt, ") \
				_T("                   0 otherqty, ") \
				_T("                   CASE WHEN hd_object IN ( 3, 8 ) THEN SUM(hfe_quantity*hfe_polprice) ") \
				_T("						WHEN hd_object = 12 THEN SUM(hfe_quantity*hfe_insprice)") \
				_T("                   ELSE 0 ") \
				_T("                   END                   otheramt ") \
				_T("         FROM    hmsv_paraclinic ") \
				_T("         LEFT JOIN hms_fee f ON ( hfe_docno = hpc_docno AND hpc_orderid = hfe_orderid) ") \
				_T("         LEFT JOIN hms_fee_list ON ( hfl_feeid = hfe_itemid ) ") \
				_T("         LEFT JOIN hms_doc ON ( hd_docno = hpc_docno ) ") \
				_T("		 LEFT JOIN hms_object ON (hd_object = ho_id)") \
				_T("         LEFT JOIN hms_card ON ( hd_patientno = hc_patientno AND hc_idx = hd_cardidx ) ") \
				_T("         WHERE     hpc_status IN ( 'P', 'S', 'T' ) AND hfe_group NOT IN ('B2200', 'B2300')") \
				_T("			   AND hfe_class = 'E' ") \
				_T("               AND hfl_hitech <> 'Y' %s") \
				_T("         GROUP     BY hpc_doctor, ") \
				_T("                      hpc_roomid, ") \
				_T("                      hd_object, ") \
				_T("                      hc_line, ho_type ") \
				_T("		 UNION ALL") \
				_T("		 SELECT    he_roomid, ") \
				_T("                   he_doctor, ") \
				_T("                   count(*)                     patientqty, ") \
				_T("                   0 amount,") \
				_T("                   CASE WHEN hd_object = 7 THEN count(*) ") \
				_T("                   ELSE 0 ") \
				_T("                   END                   serviceqty, ") \
				_T("                   0 serviceamt, ") \
				_T("                   CASE WHEN ho_type IN ('I', 'C') AND hc_line <> 1 AND hd_object <> 12 THEN count(*) ") \
				_T("                   ELSE 0 ") \
				_T("                   END                   line2qty, ") \
				_T("                   0 line2amt, ") \
				_T("                   CASE WHEN ho_type IN ('I', 'C') AND hc_line = 1 AND hd_object <> 12 THEN count(*) ") \
				_T("                   ELSE 0 ") \
				_T("                   END                   line1qty, ") \
				_T("                   0 line1amt, ") \
				_T("                   CASE WHEN hd_object = 1 THEN count(*) ") \
				_T("                   ELSE 0 ") \
				_T("                   END                   solqty, ") \
				_T("                   0 solamt, ") \
				_T("                   CASE WHEN hd_object IN ( 3, 8, 12 ) THEN count(*) ") \
				_T("                   ELSE 0 ") \
				_T("                   END                   otherqty, ") \
				_T("                   0 otheramt ") \
				_T("         FROM    hms_exam ") \
				_T("         LEFT JOIN hms_doc ON ( hd_docno = he_docno ) ") \
				_T("		 LEFT JOIN hms_object ON (hd_object = ho_id)") \
				_T("         LEFT JOIN hms_card ON ( hd_patientno = hc_patientno AND hc_idx = hd_cardidx ) ") \
				_T("         WHERE     he_status IN ( 'P', 'T' ) %s") \
				_T("         GROUP     BY he_doctor, ") \
				_T("                      he_roomid, ") \
				_T("                      hd_object, ") \
				_T("                      hc_line, ho_type) tbl0 ") \
				_T(" LEFT JOIN hms_roomlist ON (hrl_deptid = '%s' AND hrl_id = hpc_roomid)") \
				_T(" GROUP  BY hrl_name, hpc_roomid,") \
				_T("           hpc_doctor ") \
				_T(" HAVING SUM(patientqty) > 0 ") \
				_T(" ORDER  BY cast(substr(hrl_name, 1, 2) as integer), ") \
				_T("           hpc_doctor "), szWhere, szWhere1, pMF->GetCurrentDepartmentID());
_fmsg(_T("%s"), szSQL);
	return szSQL;
}