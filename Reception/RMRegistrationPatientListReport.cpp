#include "stdafx.h"
#include "RMRegistrationPatientListReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CRMRegistrationPatientListReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRMRegistrationPatientListReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CRMRegistrationPatientListReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CRMRegistrationPatientListReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CRMRegistrationPatientListReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnExamRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRMRegistrationPatientListReport* )pWnd)->OnExamRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnExamRoomSelendokFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnExamRoomSelendok();
}
/*static void _OnExamRoomSetfocusFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnExamRoomKillfocus();
}*/
/*static void _OnExamRoomKillfocusFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnExamRoomKillfocus();
}*/
static long _OnExamRoomLoadDataFnc(CWnd *pWnd){
	return ((CRMRegistrationPatientListReport *)pWnd)->OnExamRoomLoadData();
}
/*static void _OnExamRoomAddNewFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnExamRoomAddNew();
}*/
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRMRegistrationPatientListReport* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CRMRegistrationPatientListReport *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CRMRegistrationPatientListReport *)pWnd)->OnObjectAddNew();
}*/
static void _OnAllSelectFnc(CWnd *pWnd){
	 ((CRMRegistrationPatientListReport*)pWnd)->OnAllSelect();
}
static void _OnExaminatedSelectFnc(CWnd *pWnd){
	 ((CRMRegistrationPatientListReport*)pWnd)->OnExaminatedSelect();
}
static void _OnHasfeeSelectFnc(CWnd *pWnd){
	 ((CRMRegistrationPatientListReport*)pWnd)->OnHasfeeSelect();
}
static void _OnPrintSelectFnc(CWnd *pWnd){
	CRMRegistrationPatientListReport *pVw = (CRMRegistrationPatientListReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CRMRegistrationPatientListReport *pVw = (CRMRegistrationPatientListReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CRMRegistrationPatientListReport *pVw = (CRMRegistrationPatientListReport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddHRRegistrationPatientListReportFnc(CWnd *pWnd){
	 return ((CRMRegistrationPatientListReport*)pWnd)->OnAddHRRegistrationPatientListReport();
} 
static int _OnEditHRRegistrationPatientListReportFnc(CWnd *pWnd){
	 return ((CRMRegistrationPatientListReport*)pWnd)->OnEditHRRegistrationPatientListReport();
} 
static int _OnDeleteHRRegistrationPatientListReportFnc(CWnd *pWnd){
	 return ((CRMRegistrationPatientListReport*)pWnd)->OnDeleteHRRegistrationPatientListReport();
} 
static int _OnSaveHRRegistrationPatientListReportFnc(CWnd *pWnd){
	 return ((CRMRegistrationPatientListReport*)pWnd)->OnSaveHRRegistrationPatientListReport();
} 
static int _OnCancelHRRegistrationPatientListReportFnc(CWnd *pWnd){
	 return ((CRMRegistrationPatientListReport*)pWnd)->OnCancelHRRegistrationPatientListReport();
} 
CRMRegistrationPatientListReport::CRMRegistrationPatientListReport(CWnd *pParent){
	m_nDlgWidth = 945;
	m_nDlgHeight = 220;
	m_szTitle = _T("\x44\x61nh s\xE1\x63h \x62\x1EC7nh nh\xE2n ti\x1EBFp \x111\xF3n");
	SetDefaultValues();
}
CRMRegistrationPatientListReport::~CRMRegistrationPatientListReport(){
}
void CRMRegistrationPatientListReport::OnCreateComponents(){
	m_wndReceptionList.Create(this, _T("Report Condition"), 5, 5, 490, 180);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 245, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 250, 30, 330, 55);
	m_wndReportPeriod.Create(this,335, 30, 485, 55); 
	m_wndFromDate.EnableCalendar(true);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 245, 85); 
	m_wndToDate.EnableCalendar(true);
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 60, 330, 85);
	m_wndToDate.Create(this,335, 60, 485, 85); 
	m_wndExamRoomLabel.Create(this, _T("Exam Room"), 10, 90, 90, 115);
	m_wndExamRoom.Create(this,95, 90, 485, 115); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 120, 90, 145);
	m_wndObject.Create(this,95, 120, 485, 145); 
	m_wndAll.Create(this, _T("All"), 175, 150, 275, 175);
	m_wndExaminated.Create(this, _T("Examinated"), 280, 150, 380, 175);
	m_wndHasfee.Create(this, _T("Has fee"), 385, 150, 485, 175);
	m_wndPrint.Create(this, _T("&Print"), 325, 185, 405, 210);
	m_wndExport.Create(this, _T("&Export"), 410, 185, 490, 210);

}
void CRMRegistrationPatientListReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(35);
	////m_wndYear.SetCheckValue(true);
	m_wndYear.ModifyStyle(0, ES_NUMBER);
	
	////m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
//	m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
//	m_wndFromDate.SetCheckValue(true);
//	m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
//  m_wndToDate.SetCheckValue(true);
	//m_wndExamRoom.SetCheckValue(true);
	m_wndExamRoom.LimitText(65);
	//m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

	m_wndExamRoom.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndExamRoom.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndObject.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);
}
void CRMRegistrationPatientListReport::OnSetWindowEvents(){
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
	m_wndExamRoom.SetEvent(WE_SELENDOK, _OnExamRoomSelendokFnc);
	//m_wndExamRoom.SetEvent(WE_SETFOCUS, _OnExamRoomSetfocusFnc);
	//m_wndExamRoom.SetEvent(WE_KILLFOCUS, _OnExamRoomKillfocusFnc);
	m_wndExamRoom.SetEvent(WE_SELCHANGE, _OnExamRoomSelectChangeFnc);
	m_wndExamRoom.SetEvent(WE_LOADDATA, _OnExamRoomLoadDataFnc);
	//m_wndExamRoom.SetEvent(WE_ADDNEW, _OnExamRoomAddNewFnc);
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndAll.SetEvent(WE_CLICK, _OnAllSelectFnc);
	m_wndExaminated.SetEvent(WE_CLICK, _OnExaminatedSelectFnc);
	m_wndHasfee.SetEvent(WE_CLICK, _OnHasfeeSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	SetMode(GetMode());

}
void CRMRegistrationPatientListReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndExamRoom.GetDlgCtrlID(), m_szExamRoomKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_Radio(pDX, m_wndAll.GetDlgCtrlID(), m_nAll);
	DDX_Radio(pDX, m_wndExaminated.GetDlgCtrlID(), m_nExaminated);
	DDX_Radio(pDX, m_wndHasfee.GetDlgCtrlID(), m_nHasfee);

}
void CRMRegistrationPatientListReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CRMRegistrationPatientListReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CRMRegistrationPatientListReport::SetDefaultValues(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	//m_szFromDate.Empty();
	//m_szToDate.Empty();
	m_szFromDate = m_szToDate = pMF->GetSysDate(yyyymmdd);
	m_szToDate += _T("23:59");
	m_szExamRoomKey.Empty();
	m_szObjectKey.Empty();
	m_nAll = 0;
	m_nExaminated = 1;
	m_nHasfee = 1;
}
int CRMRegistrationPatientListReport::SetMode(int nMode){
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
 			EnableControls(TRUE); 
 			EnableButtons(FALSE, 3, 4, -1); 
 			break; 
 		case VM_NONE: 
 			EnableControls(TRUE); 
 			EnableButtons(TRUE, 0, 1, 2, -1); 
 			SetDefaultValues(); 
 			break; 
 		}; 
 		UpdateData(FALSE); 
 		return nOldMode; 
}
/*void CRMRegistrationPatientListReport::OnYearChange(){
	
} */
/*void CRMRegistrationPatientListReport::OnYearSetfocus(){
	
} */
/*void CRMRegistrationPatientListReport::OnYearKillfocus(){
	
} */
int CRMRegistrationPatientListReport::OnYearCheckValue(){
	return 0;
} 
void CRMRegistrationPatientListReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CRMRegistrationPatientListReport::OnReportPeriodSelendok(){
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
/*void CRMRegistrationPatientListReport::OnReportPeriodSetfocus(){
	
}*/
/*void CRMRegistrationPatientListReport::OnReportPeriodKillfocus(){
	
}*/
long CRMRegistrationPatientListReport::OnReportPeriodLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szSQL.Format(_T("SELECT hpr_idx, hpr_name FROM hms_period_report ORDER BY hpr_idx"));
	if(m_wndReportPeriod.IsSearchKey() && !m_szReportPeriodKey.IsEmpty())
	{
	};
	m_wndReportPeriod.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndReportPeriod.AddItems(
			rs.GetValue(_T("hpr_idx")), 
			rs.GetValue(_T("hpr_name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CRMRegistrationPatientListReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CRMRegistrationPatientListReport::OnFromDateChange(){
	
} */
/*void CRMRegistrationPatientListReport::OnFromDateSetfocus(){
	
} */
/*void CRMRegistrationPatientListReport::OnFromDateKillfocus(){
	
} */
int CRMRegistrationPatientListReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CRMRegistrationPatientListReport::OnToDateChange(){
	
} */
/*void CRMRegistrationPatientListReport::OnToDateSetfocus(){
	
} */
/*void CRMRegistrationPatientListReport::OnToDateKillfocus(){
	
} */
int CRMRegistrationPatientListReport::OnToDateCheckValue(){
	return 0;
}
void CRMRegistrationPatientListReport::OnAllSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CRMRegistrationPatientListReport::OnExaminatedSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CRMRegistrationPatientListReport::OnHasfeeSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CRMRegistrationPatientListReport::OnExamRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
} 
void CRMRegistrationPatientListReport::OnExamRoomSelendok(){
	 
}
/*void CRMRegistrationPatientListReport::OnExamRoomSetfocus(){
	
}*/
/*void CRMRegistrationPatientListReport::OnExamRoomKillfocus(){
	
}*/
long CRMRegistrationPatientListReport::OnExamRoomLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndExamRoom.IsSearchKey() && !m_szExamRoomKey.IsEmpty())
	{
		szWhere.Format(_T("and hrl_id = %d"), ToInt(m_szExamRoomKey)); 
	};
	m_wndExamRoom.DeleteAllItems(); 
	szSQL.Format(_T("select hrl_name as name, hrl_id as id from hms_roomlist left join sys_dept on (sd_id = hrl_deptid) where hrl_deptid = '%s' %s Order by id"), pMF->m_szDept,szWhere);
//_debug(_T("%s"), szSQL);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF())
	{ 
		m_wndExamRoom.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CRMRegistrationPatientListReport::OnExamRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */

void CRMRegistrationPatientListReport::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CRMRegistrationPatientListReport::OnObjectSelendok(){
	 
}
/*void CRMRegistrationPatientListReport::OnObjectSetfocus(){
	
}*/
/*void CRMRegistrationPatientListReport::OnObjectKillfocus(){
	
}*/
long CRMRegistrationPatientListReport::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szSQL.Format(_T("SELECT ho_id, ho_desc FROM hms_object ORDER BY ho_id"));
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty())
	{
	};
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("ho_id")), 
			rs.GetValue(_T("ho_desc")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CRMRegistrationPatientListReport::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CRMRegistrationPatientListReport::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString tmpStr;
	CString szSQL;	
	CString szWhere;
	int index = 0;
	if(!rpt.Init(_T("Reports/HMS/RM_DANHSACHBENHNHANTIEPDON.RPT")) )
		return ;
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	//Object
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportObject")), m_wndObject.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("ReportObject"), tmpStr);
	//pMF->m_CompanyInfo.sc_name
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), tmpStr);
	//HealthService
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HealthService"), tmpStr);

	//Page Header
	//Report Detail

	BeginWaitCursor();	
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);

	while (!rs.IsEOF()){
		CReportSection* rptDetail = rpt.AddDetail(); 
		index++;
		rptDetail->SetValue(_T("1"), index);

		rs.GetValue(_T("SoHs"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("SoPk"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		tmpStr = CDate::Convert(rs.GetValue(_T("NgayKham")), yyyymmdd,ddmmyyyy);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("HovaTen"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("Tuoi"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("Gioi"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("NgheNghiep"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("SoTheBHYT"), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("DiaChi"), tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.MoveNext();
	}
	EndWaitCursor();

	tmpStr.Format(_T("%s"),pMF->GetSysDate());
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4)); 
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CRMRegistrationPatientListReport::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	UpdateData(true);
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	//Column Width
	xls.SetColumnWidth(0, 6);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 14);
	xls.SetColumnWidth(3, 12);
	xls.SetColumnWidth(4, 22);
	xls.SetColumnWidth(5, 5);
	xls.SetColumnWidth(6, 5);
	xls.SetColumnWidth(7, 17);
	xls.SetColumnWidth(8, 25);
	xls.SetColumnWidth(9, 45);

	int nCol = 0;
	int nRow = 0;	
	
	//Report Header
	nRow++;
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 11);
	nRow++;
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);
	nRow++;
	xls.SetCellMergedColumns(nCol, nRow, 10);
	xls.SetCellText(nCol, nRow, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N TI\x1EBEP \x110\xD3N"), FMT_TEXT | FMT_CENTER, true, 14);
	nRow++;
	xls.SetCellMergedColumns(nCol, nRow, 10);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	nRow += 2;
	xls.SetCellText(nCol, nRow, _T("STT"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 1, nRow, _T("S\x1ED1 h\x1ED3 s\x1A1"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 2, nRow, _T("S\x1ED1 phi\x1EBFu kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 3, nRow, _T("Ng\xE0y kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 4, nRow, _T("H\x1ECD v\xE0 t\xEAn"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 5, nRow, _T("Tu\x1ED5i"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 6, nRow, _T("Gi\x1EDBi"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 7, nRow, _T("Ngh\x1EC1 nghi\x1EC7p"), FMT_TEXT |FMT_CENTER, true, 10);	

	xls.SetCellText(nCol + 8, nRow, _T("S\x1ED1 th\x1EBB \x42HYT"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 9, nRow, _T("\x110\x1ECB\x61 \x63h\x1EC9"), FMT_TEXT | FMT_CENTER, true, 10);	
	
	//SQL command
	BeginWaitCursor();
	
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	
	_fmsg(_T("%s"), szSQL);

	//Detail section
	int nIndex = 1;
	//BeginWaitCursor();
	nRow++;
	while(!rs.IsEOF())
	{
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("SoHs"),tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_INTEGER);
 
		rs.GetValue(_T("SoPk"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("NgayKham"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), FMT_DATE | FMT_CENTER);
		
		rs.GetValue(_T("HovaTen"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("Tuoi"), tmpStr);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("Gioi"), tmpStr);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("NgheNghiep"), tmpStr);
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_TEXT );

		rs.GetValue(_T("SoTheBHYT"), tmpStr);
		xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("DiaChi"), tmpStr);
		xls.SetCellText(nCol + 9, nRow, tmpStr, FMT_TEXT);

		nRow++;
		rs.MoveNext();	
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\DANH SACH BENH NHAN TIEP DON.xls"));
} 
void CRMRegistrationPatientListReport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
} 
int CRMRegistrationPatientListReport::OnAddHRRegistrationPatientListReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CRMRegistrationPatientListReport::OnEditHRRegistrationPatientListReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CRMRegistrationPatientListReport::OnDeleteHRRegistrationPatientListReport(){
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
 		OnCancelHRRegistrationPatientListReport(); 
 	}
	return 0;
}
int CRMRegistrationPatientListReport::OnSaveHRRegistrationPatientListReport(){
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
 		//OnHRRegistrationPatientListReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CRMRegistrationPatientListReport::OnCancelHRRegistrationPatientListReport(){
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
int CRMRegistrationPatientListReport::OnHRRegistrationPatientListReportListLoadData(){
	return 0;
}

CString CRMRegistrationPatientListReport::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Empty();
	if(m_nAll == 0)
		szWhere.Format(_T(" and he_status in('O','P','T')"));
	if(m_nExaminated == 0)
		szWhere.Format(_T(" and he_status in('P','T') "));
	if (m_nHasfee == 0)
		szWhere.Format(_T(" and hfe_status = 'P'"));
	if(str2int(m_szExamRoomKey) > 0)
	{
		szWhere.AppendFormat(_T(" and he_roomid= %d"), str2int(m_szExamRoomKey));
	}
	if(str2int(m_szObjectKey) > 0){
		szWhere.AppendFormat(_T(" and hd_object=%d "), str2int(m_szObjectKey));
	}
	szSQL.Format(_T(" SELECT	 distinct hd_docno AS SoHs,") \
					_T("         trim(he_roomid||'.'||he_receptno) AS SoPk,") \
					_T("         (hd_admitdate ) AS NgayKham,") \
					_T("         trim(hp_surname)||' '||trim(hp_midname)||' '||trim(hp_firstname) AS HovaTen,") \
					_T("       	hms_getage((hd_admitdate),hp_birthdate) AS Tuoi,") \
					_T("         (select distinct ss_desc FROM sys_sel WHERE ss_id = 'sys_sex' AND ss_code = hp_sex) AS Gioi,") \
					_T("         (select distinct ss_desc FROM sys_sel WHERE ss_id = 'sys_occupation' AND cast(ss_code AS integer) = hp_occupation) AS NgheNghiep,") \
					_T("         hd_cardno AS SoTheBHYT,") \
					_T("         hms_getaddress(hp_provid, hp_distid, hp_villid) AS DiaChi") \
					_T(" FROM hms_patient ") \
					_T(" LEFT JOIN hms_doc ON (hd_patientno = hp_patientno)") \
					_T(" LEFT JOIN hms_exam ON (he_docno = hd_docno)") \
					_T(" LEFT JOIN hms_fee ON (hfe_docno = hd_docno AND hfe_group = 'D0000')") \
					_T(" WHERE (he_examdate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
					_T(" AND he_deptid= '%s' ") \
					_T(" %s") \
					_T(" AND hd_admitdept = he_deptid") \
					_T(" ORDER BY hd_docno"), m_szFromDate, m_szToDate, pMF->m_szDept, szWhere);
	return szSQL;
}