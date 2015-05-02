#include "stdafx.h"
#include "EMInsuranceDrugAmount.h"
#include "HMSMainFrame.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMInsuranceDrugAmount *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMInsuranceDrugAmount* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMInsuranceDrugAmount *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMInsuranceDrugAmount *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMInsuranceDrugAmount *)pWnd)->OnToDateCheckValue();
} 
static void _OnRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMInsuranceDrugAmount* )pWnd)->OnRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnRoomSelendokFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnRoomSelendok();
}
/*static void _OnRoomSetfocusFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnRoomKillfocus();
}*/
/*static void _OnRoomKillfocusFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnRoomKillfocus();
}*/
static long _OnRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMInsuranceDrugAmount *)pWnd)->OnRoomLoadData();
}
/*static void _OnRoomAddNewFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnRoomAddNew();
}*/
static void _OnDoctorSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMInsuranceDrugAmount* )pWnd)->OnDoctorSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDoctorSelendokFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnDoctorSelendok();
}
/*static void _OnDoctorSetfocusFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnDoctorKillfocus();
}*/
/*static void _OnDoctorKillfocusFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnDoctorKillfocus();
}*/
static long _OnDoctorLoadDataFnc(CWnd *pWnd){
	return ((CEMInsuranceDrugAmount *)pWnd)->OnDoctorLoadData();
}
/*static void _OnDoctorAddNewFnc(CWnd *pWnd){
	((CEMInsuranceDrugAmount *)pWnd)->OnDoctorAddNew();
}*/
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMInsuranceDrugAmount *pVw = (CEMInsuranceDrugAmount *)pWnd;
	pVw->OnExportSelect();
} 
CEMInsuranceDrugAmount::CEMInsuranceDrugAmount(CWnd *pWnd){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CEMInsuranceDrugAmount::~CEMInsuranceDrugAmount(){
}
void CEMInsuranceDrugAmount::OnCreateComponents(){
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
void CEMInsuranceDrugAmount::OnInitializeComponents(){
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
void CEMInsuranceDrugAmount::OnSetWindowEvents(){
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
void CEMInsuranceDrugAmount::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndRoom.GetDlgCtrlID(), m_szRoomKey);
	DDX_TextEx(pDX, m_wndDoctor.GetDlgCtrlID(), m_szDoctorKey);

}
void CEMInsuranceDrugAmount::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szRoomKey.Empty();
	m_szDoctorKey.Empty();

}
int CEMInsuranceDrugAmount::SetMode(int nMode){
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
/*void CEMInsuranceDrugAmount::OnYearChange(){
	
} */
/*void CEMInsuranceDrugAmount::OnYearSetfocus(){
	
} */
/*void CEMInsuranceDrugAmount::OnYearKillfocus(){
	
} */
int CEMInsuranceDrugAmount::OnYearCheckValue(){
	return 0;
} 
void CEMInsuranceDrugAmount::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMInsuranceDrugAmount::OnReportPeriodSelendok(){
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
/*void CEMInsuranceDrugAmount::OnReportPeriodSetfocus(){
	
}*/
/*void CEMInsuranceDrugAmount::OnReportPeriodKillfocus(){
	
}*/
long CEMInsuranceDrugAmount::OnReportPeriodLoadData(){
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
/*void CEMInsuranceDrugAmount::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMInsuranceDrugAmount::OnFromDateChange(){
	
} */
/*void CEMInsuranceDrugAmount::OnFromDateSetfocus(){
	
} */
/*void CEMInsuranceDrugAmount::OnFromDateKillfocus(){
	
} */
int CEMInsuranceDrugAmount::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMInsuranceDrugAmount::OnToDateChange(){
	
} */
/*void CEMInsuranceDrugAmount::OnToDateSetfocus(){
	
} */
/*void CEMInsuranceDrugAmount::OnToDateKillfocus(){
	
} */
int CEMInsuranceDrugAmount::OnToDateCheckValue(){
	return 0;
} 
void CEMInsuranceDrugAmount::OnRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMInsuranceDrugAmount::OnRoomSelendok(){
	 
}
/*void CEMInsuranceDrugAmount::OnRoomSetfocus(){
	
}*/
/*void CEMInsuranceDrugAmount::OnRoomKillfocus(){
	
}*/
long CEMInsuranceDrugAmount::OnRoomLoadData(){
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
/*void CEMInsuranceDrugAmount::OnRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMInsuranceDrugAmount::OnDoctorSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMInsuranceDrugAmount::OnDoctorSelendok(){
	 
}
/*void CEMInsuranceDrugAmount::OnDoctorSetfocus(){
	
}*/
/*void CEMInsuranceDrugAmount::OnDoctorKillfocus(){
	
}*/
long CEMInsuranceDrugAmount::OnDoctorLoadData(){
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
/*void CEMInsuranceDrugAmount::OnDoctorAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMInsuranceDrugAmount::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CExcel xls;
	CString szSQL, tmpStr;
	CStringArray arrField;
	int nCol = 0, nRow = 0, nOldRow = 0;
	double nTmp = 0;
	long double nTotal[12];
	CString szPreviousDept, szCurrentDept;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	for (int i = 0; i < 12; i++)
	{
		nTotal[i] = 0;
	}
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
	xls.SetCellText(0, 2, _T("Th\x1ED1ng k\xEA s\x1ED1 ti\x1EC1n \x42N thu\x1ED1\x63 \x63\x1EE7\x61 \x42N \x42\x1EA3o hi\x1EC3m y t\x1EBF v\xE0 Qu\xE2n"), 4098, true);
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
	xls.SetCellText(5, 4, _T("BH \x63huy\x1EC3n+tr\xE1i tuy\x1EBFn"), 4098, true);
	xls.SetCellText(5, 5, _T("S\x1ED1 \x42N"), 4098);
	xls.SetCellText(6, 5, _T("S\x1ED1 ti\x1EC1n"), 4098);
	xls.SetCellText(7, 5, _T("TB/BN"), 4098);
	xls.SetCellMergedColumns(8, 4, 3);
	xls.SetCellText(8, 4, _T("BH t\x1EA1i \x63h\x1ED7"), 4098, true);
	xls.SetCellText(8, 5, _T("S\x1ED1 \x42N"), 4098);
	xls.SetCellText(9, 5, _T("S\x1ED1 ti\x1EC1n"), 4098);
	xls.SetCellText(10, 5, _T("TB/BN"), 4098);
	xls.SetCellMergedColumns(11, 4, 3);
	xls.SetCellText(11, 4, _T("Qu\xE2n"), 4098, true);
	xls.SetCellText(11, 5, _T("S\x1ED1 \x42N"), 4098);
	xls.SetCellText(12, 5, _T("S\x1ED1 ti\x1EC1n"), 4098);
	xls.SetCellText(13, 5, _T("TB/BN"), 4098);

	nRow = 6;
	nOldRow = 6;
	arrField.Add(_T("patientqty"));
	arrField.Add(_T("amt"));
	arrField.Add(_T("patientrate"));
	arrField.Add(_T("line2qty"));
	arrField.Add(_T("line2amt"));
	arrField.Add(_T("line2rate"));
	arrField.Add(_T("line1qty"));
	arrField.Add(_T("line1amt"));
	arrField.Add(_T("line1rate"));
	arrField.Add(_T("solqty"));
	arrField.Add(_T("solamt"));
	arrField.Add(_T("solrate"));

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
			if (arrField.GetAt(i).Find(_T("qty")) != -1 || arrField.GetAt(i).Find(_T("amt")) != -1)
			{
				nTotal[i] += nTmp;
			}
		}
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellText(1, nRow, _T("T\x1ED5ng"), FMT_TEXT | FMT_CENTER, true);
	for (int i = 0; i < 12; i++)
	{
		tmpStr.Format(_T("%f"), nTotal[i]);
		xls.SetCellText(nCol + i + 2, nRow, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\Thong ke so tien thuoc BN BH va Quan.xls"));
} 
CString CEMInsuranceDrugAmount::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND hpo_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_deptid = '%s' AND hpo_roomid = %s"), pMF->GetCurrentDepartmentID(), m_szRoomKey);
	if (!m_szDoctorKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_doctor = '%s'"), m_szDoctorKey);
	szSQL.Format(_T(" SELECT hrl_name roomid, hpo_roomid, ") \
				_T("        get_username(hpo_doctor) doctor, ") \
				_T("        SUM(patientqty) patientqty, ") \
				_T("        SUM(amount)     amt, ") \
				_T("		CASE WHEN SUM(amount) > 0 THEN SUM(amount)/SUM(patientqty) ELSE 0 END patientrate,") \
				_T("        SUM(line2qty)   line2qty, ") \
				_T("        SUM(line2amt)   line2amt, ") \
				_T("		CASE WHEN SUM (line2amt) > 0 THEN SUM(line2amt)/SUM(line2qty) ELSE 0 END line2rate,") \
				_T("        SUM(line1qty)   line1qty, ") \
				_T("        SUM(line1amt)   line1amt, ") \
				_T("		CASE WHEN SUM(line1amt) > 0 THEN SUM(line1amt)/SUM(line1qty) ELSE 0 END line1rate,") \
				_T("        SUM(solqty)     solqty, ") \
				_T("        SUM(solamt)     solamt, ") \
				_T("		CASE WHEN SUM(solqty) > 0 THEN SUM(solamt)/SUM(solqty) ELSE 0 END solrate") \
				_T(" FROM   (SELECT    hpo_roomid, ") \
				_T("                   hpo_doctor, ") \
				_T("                   CASE WHEN amount > 0 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END patientqty, ") \
				_T("                   amount, ") \
				_T("                   CASE WHEN hd_object = 7 ") \
				_T("                        AND amount > 0 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END serviceqty, ") \
				_T("                   CASE WHEN hd_object = 7 THEN amount ") \
				_T("                   ELSE 0 ") \
				_T("                   END serviceamt, ") \
				_T("                   CASE WHEN hc_line <> 1 ") \
				_T("                        AND amount > 0 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END line2qty, ") \
				_T("                   CASE WHEN hc_line <> 1 THEN amount ") \
				_T("                   ELSE 0 ") \
				_T("                   END line2amt, ") \
				_T("                   CASE WHEN hc_line = 1 ") \
				_T("                        AND amount > 0 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END line1qty, ") \
				_T("                   CASE WHEN hc_line = 1 THEN amount ") \
				_T("                   ELSE 0 ") \
				_T("                   END line1amt, ") \
				_T("                   CASE WHEN hd_object = 1 ") \
				_T("                        AND amount > 0 THEN 1 ") \
				_T("                   ELSE 0 ") \
				_T("                   END solqty, ") \
				_T("                   CASE WHEN hd_object = 1 THEN amount ") \
				_T("                   ELSE 0 ") \
				_T("                   END solamt ") \
				_T("         FROM      (SELECT    hd_patientno, hpo_docno, ") \
				_T("                              hpo_doctor, ") \
				_T("                              hpo_roomid, ") \
				_T("							  hd_cardidx,") \
				_T("							  hd_object,") \
				_T("                              SUM(hpol_unitprice*hpol_qtyorder) amount ") \
				_T("                    FROM      hms_pharmaorder ") \
				_T("                    LEFT JOIN hms_pharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T("					LEFT JOIN hms_doc ON ( hd_docno = hpo_docno ) ") \
				_T("					LEFT JOIN hms_object ON (hd_object = ho_id)") \
				_T("                    WHERE     hpo_orderstatus IN ( 'S', 'A' ) AND hpo_ordertype = 'P'") \
				_T("						  AND ho_type IN ('I', 'D')") \
				_T("                          AND hpo_deptid = 'C1.1' ") \
				_T("                          AND hpo_roomid > 0 %s") \
				_T("                    GROUP     BY hd_patientno, hpo_docno, ") \
				_T("                                 hpo_doctor, ") \
				_T("                                 hpo_roomid, hd_cardidx, hd_object) tblpharma ") \
				_T("         LEFT JOIN hms_card ON ( hd_patientno = hc_patientno ") \
				_T("                                 AND hc_idx = hd_cardidx )) tbl0 ") \
				_T(" LEFT JOIN hms_roomlist ON (hrl_deptid = '%s' AND hrl_id = hpo_roomid)") \
				_T(" GROUP  BY hrl_name, hpo_roomid,") \
				_T("           hpo_doctor ") \
				_T(" HAVING SUM(patientqty) > 0 ") \
				_T(" ORDER  BY cast(substr(hrl_name, 1, 2) as integer), ") \
				_T("           hpo_doctor "), szWhere, pMF->GetCurrentDepartmentID());_fmsg(_T("%s"), szSQL);
	return szSQL;
}