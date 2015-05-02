#include "stdafx.h"
#include "PACSImageFosteringList.h"
#include "HMSMainFrame.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CPACSImageFosteringList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSImageFosteringList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CPACSImageFosteringList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSImageFosteringList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSImageFosteringList *)pWnd)->OnToDateCheckValue();
} 
static void _OnGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSImageFosteringList* )pWnd)->OnGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnGroupSelendokFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnGroupSelendok();
}
/*static void _OnGroupSetfocusFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnGroupKillfocus();
}*/
/*static void _OnGroupKillfocusFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnGroupKillfocus();
}*/
static long _OnGroupLoadDataFnc(CWnd *pWnd){
	return ((CPACSImageFosteringList *)pWnd)->OnGroupLoadData();
}
/*static void _OnGroupAddNewFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnGroupAddNew();
}*/
static long _OnItemLoadDataFnc(CWnd *pWnd){
	return ((CPACSImageFosteringList *)pWnd)->OnItemLoadData();
}
static void _OnPractitionerSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSImageFosteringList* )pWnd)->OnPractitionerSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnPractitionerSelendokFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnPractitionerSelendok();
}
/*static void _OnPractitionerSetfocusFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnPractitionerKillfocus();
}*/
/*static void _OnPractitionerKillfocusFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnPractitionerKillfocus();
}*/
static long _OnPractitionerLoadDataFnc(CWnd *pWnd){
	return ((CPACSImageFosteringList *)pWnd)->OnPractitionerLoadData();
}
/*static void _OnPractitionerAddNewFnc(CWnd *pWnd){
	((CPACSImageFosteringList *)pWnd)->OnPractitionerAddNew();
}*/
static void _OnExportSelectFnc(CWnd *pWnd){
	CPACSImageFosteringList *pVw = (CPACSImageFosteringList *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnInsuranceSelectFnc(CWnd *pWnd){
	  ((CPACSImageFosteringList*)pWnd)->OnInsuranceSelect();
}
static void _OnServiceSelectFnc(CWnd *pWnd){
	  ((CPACSImageFosteringList*)pWnd)->OnServiceSelect();
}
static void _OnPolicySelectFnc(CWnd *pWnd){
	  ((CPACSImageFosteringList*)pWnd)->OnPolicySelect();
}
CPACSImageFosteringList::CPACSImageFosteringList(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CPACSImageFosteringList::~CPACSImageFosteringList(){
}
void CPACSImageFosteringList::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 465, 180);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 115, 55);
	m_wndYear.Create(this,120, 30, 240, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 245, 30, 335, 55);
	m_wndReportPeriod.Create(this,340, 30, 460, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 115, 85);
	m_wndFromDate.Create(this,120, 60, 240, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 245, 60, 335, 85);
	m_wndToDate.Create(this,340, 60, 460, 85); 
	m_wndGroupLabel.Create(this, _T("Group"), 10, 90, 115, 115);
	m_wndGroup.Create(this,120, 90, 460, 115); 
	m_wndItemLabel.Create(this, _T("Item"), 10, 120, 115, 145);
	m_wndItem.Create(this,120, 120, 460, 145); 
	m_wndPractitionerLabel.Create(this, _T("Practitioner"), 10, 150, 115, 175);
	m_wndPractitioner.Create(this,120, 150, 460, 175); 
	m_wndExport.Create(this, _T("&Export"), 365, 185, 465, 210);
	m_wndInsurance.Create(this, _T("Insurance"), 5, 185, 85, 210);
	m_wndService.Create(this, _T("Service"), 90, 185, 170, 210);
	m_wndPolicy.Create(this, _T("Policy"), 175, 185, 255, 210);

}
void CPACSImageFosteringList::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndGroup.SetCheckValue(true);
	m_wndGroup.LimitText(1024);
	m_wndPractitioner.SetCheckValue(true);
	m_wndPractitioner.LimitText(1024);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndGroup.InsertColumn(0, _T("ID"), CFMT_TEXT, 60);
	m_wndGroup.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndItem.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndItem.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);


	m_wndPractitioner.InsertColumn(0, _T("ID"), CFMT_TEXT, 75);
	m_wndPractitioner.InsertColumn(1, _T("Name"), CFMT_TEXT, 235);

}
void CPACSImageFosteringList::OnSetWindowEvents(){
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
	m_wndGroup.SetEvent(WE_SELENDOK, _OnGroupSelendokFnc);
	//m_wndGroup.SetEvent(WE_SETFOCUS, _OnGroupSetfocusFnc);
	//m_wndGroup.SetEvent(WE_KILLFOCUS, _OnGroupKillfocusFnc);
	m_wndGroup.SetEvent(WE_SELCHANGE, _OnGroupSelectChangeFnc);
	m_wndGroup.SetEvent(WE_LOADDATA, _OnGroupLoadDataFnc);
	m_wndItem.SetEvent(WE_LOADDATA, _OnItemLoadDataFnc);
	//m_wndGroup.SetEvent(WE_ADDNEW, _OnGroupAddNewFnc);
	m_wndPractitioner.SetEvent(WE_SELENDOK, _OnPractitionerSelendokFnc);
	//m_wndPractitioner.SetEvent(WE_SETFOCUS, _OnPractitionerSetfocusFnc);
	//m_wndPractitioner.SetEvent(WE_KILLFOCUS, _OnPractitionerKillfocusFnc);
	m_wndPractitioner.SetEvent(WE_SELCHANGE, _OnPractitionerSelectChangeFnc);
	m_wndPractitioner.SetEvent(WE_LOADDATA, _OnPractitionerLoadDataFnc);
	//m_wndPractitioner.SetEvent(WE_ADDNEW, _OnPractitionerAddNewFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndInsurance.SetEvent(WE_CLICK, _OnInsuranceSelectFnc);
	m_wndService.SetEvent(WE_CLICK, _OnServiceSelectFnc);
	m_wndPolicy.SetEvent(WE_CLICK, _OnPolicySelectFnc);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	m_szGroupKey = _T("B3200");
	UpdateData(false);

}
void CPACSImageFosteringList::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndGroup.GetDlgCtrlID(), m_szGroupKey);
	DDX_TextEx(pDX, m_wndItem.GetDlgCtrlID(), m_szItemKey);
	DDX_TextEx(pDX, m_wndPractitioner.GetDlgCtrlID(), m_szPractitionerKey);
	DDX_Radio(pDX, m_wndInsurance.GetDlgCtrlID(), m_nInsurance);
	DDX_Radio(pDX, m_wndService.GetDlgCtrlID(), m_nService);
	DDX_Radio(pDX, m_wndPolicy.GetDlgCtrlID(), m_nPolicy);

}
void CPACSImageFosteringList::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szGroupKey.Empty();
	m_szItemKey.Empty();
	m_szPractitionerKey.Empty();
	m_nInsurance=0;
	m_nService=1;
	m_nPolicy=1;

}
int CPACSImageFosteringList::SetMode(int nMode){
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
/*void CPACSImageFosteringList::OnYearChange(){
	
} */
/*void CPACSImageFosteringList::OnYearSetfocus(){
	
} */
/*void CPACSImageFosteringList::OnYearKillfocus(){
	
} */
int CPACSImageFosteringList::OnYearCheckValue(){
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
void CPACSImageFosteringList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSImageFosteringList::OnReportPeriodSelendok(){
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
/*void CPACSImageFosteringList::OnReportPeriodSetfocus(){
	
}*/
/*void CPACSImageFosteringList::OnReportPeriodKillfocus(){
	
}*/
long CPACSImageFosteringList::OnReportPeriodLoadData(){
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
/*void CPACSImageFosteringList::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CPACSImageFosteringList::OnFromDateChange(){
	
} */
/*void CPACSImageFosteringList::OnFromDateSetfocus(){
	
} */
/*void CPACSImageFosteringList::OnFromDateKillfocus(){
	
} */
int CPACSImageFosteringList::OnFromDateCheckValue(){
	return 0;
} 
/*void CPACSImageFosteringList::OnToDateChange(){
	
} */
/*void CPACSImageFosteringList::OnToDateSetfocus(){
	
} */
/*void CPACSImageFosteringList::OnToDateKillfocus(){
	
} */
int CPACSImageFosteringList::OnToDateCheckValue(){
	return 0;
} 
void CPACSImageFosteringList::OnGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSImageFosteringList::OnGroupSelendok(){
	OnItemLoadData();
}
/*void CPACSImageFosteringList::OnGroupSetfocus(){
	
}*/
/*void CPACSImageFosteringList::OnGroupKillfocus(){
	
}*/
long CPACSImageFosteringList::OnGroupLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndGroup.IsSearchKey() && !m_szGroupKey.IsEmpty())
	{
		szWhere.Format(_T(" AND hfg_id='%s' "), m_szGroupKey);
	}

	m_wndGroup.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT hfg_id as id, hfg_name as name ") \
		         _T("FROM hms_fee_group ") \
				 _T("WHERE substr(hfg_id,1,2)='B3' AND hfg_id not in('B3100') ") \
				 _T("ORDER BY hfg_id"));

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndGroup.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPACSImageFosteringList::OnGroupAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */

long CPACSImageFosteringList::OnItemLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndItem.IsSearchKey() && !m_szItemKey.IsEmpty())
	{
		szWhere.Format(_T(" AND hfg_id='%s' "), m_szItemKey);
	}
	if (!m_szGroupKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfl_groupid = '%s'"), m_szGroupKey);
	m_wndItem.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT hfl_feeid as id, hfl_name as name ") \
		         _T("FROM hms_fee_list ") \
				 _T("WHERE substr(hfl_groupid,1,2)='B3' AND hfl_groupid not in('B3100') %s") \
				 _T("ORDER BY hfl_feeid"), szWhere);

	nCount = rs.ExecSQL(szSQL);
	if (nCount < 0)
		_msg(_T("%s"), szSQL);
	while(!rs.IsEOF()){ 
		m_wndItem.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}

void CPACSImageFosteringList::OnPractitionerSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSImageFosteringList::OnPractitionerSelendok(){
	 
}
/*void CPACSImageFosteringList::OnPractitionerSetfocus(){
	
}*/
/*void CPACSImageFosteringList::OnPractitionerKillfocus(){
	
}*/
long CPACSImageFosteringList::OnPractitionerLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndPractitioner.IsSearchKey() && !m_szPractitionerKey.IsEmpty())
	{
		szWhere.Format(_T(" AND su_userid='%s' "), m_szPractitionerKey);
	}

	m_wndPractitioner.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as id, su_name as name ") \
		         _T("FROM sys_user ") \
				 _T("WHERE su_deptid in('A3','C7','C8') ") \
				 _T("ORDER BY su_userid"));

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndPractitioner.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPACSImageFosteringList::OnPractitionerAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CPACSImageFosteringList::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CExcel xls;
	CString szSQL, tmpStr, szOldLine, szNewLine;
	double nTmp = 0;
	int nIdx = 1, nGroupIdx = 1, nRow = 0;
	szSQL = GetQueryString();
	int nRet = rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		if (nRet < 0)
			_msg(_T("%s"), szSQL);
		else
			_fmsg(_T("%s"), szSQL);
		AfxMessageBox(_T("No Data."));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 8);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(3, 25);
	xls.SetColumnWidth(5, 20);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 5);
	xls.SetCellMergedColumns(0, 3, 5);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr = _T("\x44\x41NH S\xC1\x43H \x110\x1EC0 NGH\x1ECA \x42\x1ED2I \x44\x1AF\x1EE0NG PTTT");
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	xls.SetCellText(0, 4, _T("STT"), 4098, true);
	xls.SetCellText(1, 4, _T("BN"), 4098, true);
	xls.SetCellText(2, 4, _T("Qu\xE2n h\xE0m"), 4098, true);
	xls.SetCellText(3, 4, _T("PTTT"), 4098, true);
	xls.SetCellText(4, 4, _T("Ph\xE2n lo\x1EA1i"), 4098, true);
	xls.SetCellText(5, 4, _T("Th\x1EF1\x63 hi\x1EC7n"), 4098, true);
	xls.SetCellText(6, 4, _T("M\x1EABu"), 4098, true);
	xls.SetCellText(7, 4, _T("S\x1ED1 ti\x1EC1n"), 4098, true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("order_name"), szNewLine);
		if (szNewLine != szOldLine)
		{
			tmpStr.Format(_T("%d. %s"), nGroupIdx++, szNewLine);
			xls.SetCellText(3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING, true);
			szOldLine = szNewLine;
			nRow++;
		}
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		tmpStr.Format(_T("[%s] %s"), rs.GetValue(_T("doc_no")), rs.GetValue(_T("patient_name")));
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("rank")), FMT_TEXT);
		xls.SetCellText(3, nRow, rs.GetValue(_T("order_name")), FMT_TEXT);
		xls.SetCellText(5, nRow, rs.GetValue(_T("practitioner")), FMT_TEXT);
		xls.SetCellText(6, nRow, rs.GetValue(_T("result_id")), FMT_TEXT);
		nRow++;
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\Danh sach de nghi boi duong hinh anh.xls"));
} 
void CPACSImageFosteringList::OnInsuranceSelect(){
	
}
void CPACSImageFosteringList::OnServiceSelect(){
	
}
void CPACSImageFosteringList::OnPolicySelect(){
	
}


CString CPACSImageFosteringList::GetQueryString(){
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND hpc_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
		, m_szFromDate, m_szToDate);
	if (!m_szGroupKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpc_groupid = '%s'"), m_szGroupKey);
	if (!m_szItemKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpcl_itemid = '%s'"), m_szItemKey);
	if (m_nInsurance == 0)
		szWhere.AppendFormat(_T(" AND ho_type IN ('I', 'C')"));
	else if (m_nService == 0)
		szWhere.AppendFormat(_T(" AND ho_type = 'S'"));
	else if (m_nPolicy == 0)
		szWhere.AppendFormat(_T(" AND ho_type NOT IN ('I', 'C', 'S')"));
	
	szSQL.Format(_T(" SELECT    hpcl_docno doc_no, Get_patientname(hpcl_docno) patient_name, ") \
	_T("           Get_syssel_desc('hms_rank', hd_rank)     rank, ") \
	_T("           hfl_name                                 order_name, ") \
	_T("           substr(get_username(NVL(hpcl_practitioner, hpc_practitioner)), ") \
	_T("			instr(get_username(NVL(hpcl_practitioner, hpc_practitioner)), ' ', -1) + 1, ") \
	_T("			length(get_username(NVL(hpcl_practitioner, hpc_practitioner)))) practitioner, ") \
	_T("           hpcl_result                              result_id ") \
	_T(" FROM      hms_pacsorderline ") \
	_T(" LEFT JOIN hms_pacsorder ON ( hpc_orderid = hpcl_orderid ) ") \
	_T(" LEFT JOIN hms_doc ON ( hpcl_docno = hd_docno ) ") \
	_T(" LEFT JOIN hms_fee_list ON ( hpcl_itemid = hfl_feeid ) ") \
	_T(" LEFT JOIN hms_object ON (hpc_object = ho_id)") \
	_T(" WHERE     hfe_refstatus NOT IN ( 'C', 'R' ) AND hpcl_status NOT IN ( 'O', 'C' ) %s") \
	_T(" ORDER BY order_name, patient_name"), szWhere);
	
	return szSQL;
}