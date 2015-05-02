#include "stdafx.h"
#include "FMServiceMedicalIncomingOutgoingReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((FMServiceMedicalIncomingOutgoingReport * )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((FMServiceMedicalIncomingOutgoingReport * )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnClerkAddNew();
}*/

static long _OnDeptLoadDataFnc(CWnd *pWnd){
	return ((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnDeptLoadData();
}

static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnObjectLoadData();
}
static void _OnPrintSelectFnc(CWnd *pWnd){
	FMServiceMedicalIncomingOutgoingReport *pVw = (FMServiceMedicalIncomingOutgoingReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	FMServiceMedicalIncomingOutgoingReport *pVw = (FMServiceMedicalIncomingOutgoingReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	FMServiceMedicalIncomingOutgoingReport *pVw = (FMServiceMedicalIncomingOutgoingReport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddHMSGeneralCostReportDialogFnc(CWnd *pWnd){
	 return ((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnAddHMSGeneralCostReportDialog();
} 
static int _OnEditHMSGeneralCostReportDialogFnc(CWnd *pWnd){
	 return ((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnEditHMSGeneralCostReportDialog();
} 
static int _OnDeleteHMSGeneralCostReportDialogFnc(CWnd *pWnd){
	 return ((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnDeleteHMSGeneralCostReportDialog();
} 
static int _OnSaveHMSGeneralCostReportDialogFnc(CWnd *pWnd){
	 return ((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnSaveHMSGeneralCostReportDialog();
} 
static int _OnCancelHMSGeneralCostReportDialogFnc(CWnd *pWnd){
	 return ((FMServiceMedicalIncomingOutgoingReport *)pWnd)->OnCancelHMSGeneralCostReportDialog();
} 
FMServiceMedicalIncomingOutgoingReport ::FMServiceMedicalIncomingOutgoingReport (CWnd *pParent)
	{
	m_nDlgWidth = 445;
	m_nDlgHeight = 215;
	m_szTitle = _T("\x42\xE1o \x63\xE1o thu \x63hi \x64\x1ECB\x63h v\x1EE5 y t\x1EBF");
	SetDefaultValues();
}
FMServiceMedicalIncomingOutgoingReport ::~FMServiceMedicalIncomingOutgoingReport (){
}
void FMServiceMedicalIncomingOutgoingReport ::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 100, 55);
	m_wndYear.Create(this,102, 30, 217, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 228, 30, 318, 55);
	m_wndReportFilter.Create(this, _T("Report Condition"), 5, 5, 440, 180);
	m_wndReportPeriod.Create(this,323, 30, 438, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 100, 85);
	m_wndFromDate.Create(this,102, 60, 217, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 228, 60, 318, 85);
	m_wndToDate.Create(this,323, 60, 438, 85); 
	m_wndDeptLabel.Create(this, _T("Dept"), 10, 90, 100, 115);
	m_wndDept.Create(this,102, 90, 437, 115); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 120, 100, 145);
	m_wndObject.Create(this,102, 120, 437, 145); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 150, 100, 175);
	m_wndClerk.Create(this,102, 150, 439, 175); 
	m_wndPrint.Create(this, _T("&Print"), 255, 185, 345, 210);
	m_wndExport.Create(this, _T("&Export"), 350, 185, 440, 210);

}
void FMServiceMedicalIncomingOutgoingReport ::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	//m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	//m_wndToDate.SetCheckValue(true);
	////m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);

	m_wndReportPeriod.Format(2, 1, 17);
	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 100);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

	m_wndDept.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndDept.InsertColumn(1,_T("Description"), CFMT_TEXT,200);
	
	m_wndObject.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndObject.InsertColumn(1,_T("Description"), CFMT_TEXT,200);

}
void FMServiceMedicalIncomingOutgoingReport ::OnSetWindowEvents(){
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
	m_wndClerk.SetEvent(WE_SELENDOK, _OnClerkSelendokFnc);
	//m_wndClerk.SetEvent(WE_SETFOCUS, _OnClerkSetfocusFnc);
	//m_wndClerk.SetEvent(WE_KILLFOCUS, _OnClerkKillfocusFnc);
	m_wndClerk.SetEvent(WE_SELCHANGE, _OnClerkSelectChangeFnc);
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	//m_wndClerk.SetEvent(WE_ADDNEW, _OnClerkAddNewFnc);

	m_wndDept.SetEvent(WE_LOADDATA, _OnDeptLoadDataFnc);

	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	SetMode(VM_EDIT);
}
void FMServiceMedicalIncomingOutgoingReport ::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_TextEx(pDX, m_wndDept.GetDlgCtrlID(), m_szDeptKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
}
void FMServiceMedicalIncomingOutgoingReport ::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void FMServiceMedicalIncomingOutgoingReport ::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void FMServiceMedicalIncomingOutgoingReport ::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();

}
int FMServiceMedicalIncomingOutgoingReport ::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, 1, 2, 3, -1); 
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
/*void FMServiceMedicalIncomingOutgoingReport ::OnYearChange(){
	
} */
/*void FMServiceMedicalIncomingOutgoingReport ::OnYearSetfocus(){
	
} */
/*void FMServiceMedicalIncomingOutgoingReport ::OnYearKillfocus(){
	
} */
int FMServiceMedicalIncomingOutgoingReport ::OnYearCheckValue(){
	return 0;
} 
void FMServiceMedicalIncomingOutgoingReport ::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void FMServiceMedicalIncomingOutgoingReport ::OnReportPeriodSelendok(){
	
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
/*void FMServiceMedicalIncomingOutgoingReport ::OnReportPeriodSetfocus(){
	
}*/
/*void FMServiceMedicalIncomingOutgoingReport ::OnReportPeriodKillfocus(){
	
}*/
long FMServiceMedicalIncomingOutgoingReport ::OnReportPeriodLoadData(){
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
/*void FMServiceMedicalIncomingOutgoingReport ::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void FMServiceMedicalIncomingOutgoingReport ::OnFromDateChange(){
	
} */
/*void FMServiceMedicalIncomingOutgoingReport ::OnFromDateSetfocus(){
	
} */
/*void FMServiceMedicalIncomingOutgoingReport ::OnFromDateKillfocus(){
	
} */
int FMServiceMedicalIncomingOutgoingReport ::OnFromDateCheckValue(){
	return 0;
} 
/*void FMServiceMedicalIncomingOutgoingReport ::OnToDateChange(){
	
} */
/*void FMServiceMedicalIncomingOutgoingReport ::OnToDateSetfocus(){
	
} */
/*void FMServiceMedicalIncomingOutgoingReport ::OnToDateKillfocus(){
	
} */
int FMServiceMedicalIncomingOutgoingReport ::OnToDateCheckValue(){
	return 0;
} 
void FMServiceMedicalIncomingOutgoingReport ::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void FMServiceMedicalIncomingOutgoingReport ::OnClerkSelendok(){
	 
}
/*void FMServiceMedicalIncomingOutgoingReport ::OnClerkSetfocus(){
	
}*/
/*void FMServiceMedicalIncomingOutgoingReport ::OnClerkKillfocus(){
	
}*/
long FMServiceMedicalIncomingOutgoingReport ::OnClerkLoadData(){
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
/*void FMServiceMedicalIncomingOutgoingReport ::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */

long FMServiceMedicalIncomingOutgoingReport ::OnDeptLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDept.IsSearchKey() && !m_szDeptKey.IsEmpty()){
		szWhere.Format(_T(" and sd_id ='%s' "), m_szDeptKey);
	};
	m_wndDept.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT sd_id as id, sd_name as name FROM sys_dept WHERE sd_type='DT' %s ORDER BY sd_id"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDept.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
long FMServiceMedicalIncomingOutgoingReport ::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty()){
		szWhere.Format(_T(" where ho_id = %d "), ToInt(m_szObjectKey));
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
void FMServiceMedicalIncomingOutgoingReport ::OnPrintSelect(){	
CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	if(!rpt.Init(_T("Reports/HMS/HF_BAOCAOTHUCHIDICHVUYTE.RPT")) )
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
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	long double nTotal[9];
	long cost = 0;
	for (int i = 1; i<= 8; i++)
	{
		nTotal[i] = 0.0;
	}
	int nIndex = 1;
		CReportSection* rptDetail;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("0"), tmpStr);
//khoa
		rs.GetValue(_T("c1"),tmpStr);
		rptDetail->SetValue(_T("Department"), tmpStr);
// noi tru	
		rs.GetValue(_T("c2"), cost);
		nTotal[1] += cost;
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("1"), tmpStr);
// ngoai tru
		rs.GetValue(_T("c3"), cost);
		nTotal[2] += cost;
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
//nhan vien
		rs.GetValue(_T("c4"), cost);
		nTotal[3] += cost;
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
// tong so tien
		rs.GetValue(_T("c5"), cost);
		nTotal[4] += cost;
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
// tong hop quan so noi tru
		rs.GetValue(_T("c6"), cost);
		nTotal[5] += cost;
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
//tong hop quan so nhan vien
		rs.GetValue(_T("c7"), cost);
		nTotal[6] += cost;
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
// tong
		rs.GetValue(_T("c8"), cost);
		nTotal[7] += cost;
		//FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.MoveNext();	
	}
	for (int i = 1; i<= 8; i++)
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
void FMServiceMedicalIncomingOutgoingReport ::OnExportSelect()
{	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	UpdateData(true);
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 30);
	xls.SetColumnWidth(2, 11);
	xls.SetColumnWidth(3, 11);
	xls.SetColumnWidth(4, 11);
	xls.SetColumnWidth(5, 11);
	xls.SetColumnWidth(6, 11);
	xls.SetColumnWidth(7, 11);
	xls.SetColumnWidth(8, 11);
	xls.SetColumnWidth(9, 11);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);

	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellMergedColumns(nCol + 4, nRow + 1, 6);
	xls.SetCellText(nCol + 4, nRow + 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellMergedColumns(nCol + 4, nRow + 2, 6);
	xls.SetCellText(nCol + 4, nRow + 2, _T("\x110\x1ED9\x63 l\x1EADp - T\x1EF1 \x64o - H\x1EA1nh ph\xFA\x63"), FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellMergedColumns(nCol, nRow + 3, 10);
	xls.SetCellText(nCol, nRow + 3, _T("\x42\xC1O \x43\xC1O THU \x43HI \x44\x1ECA\x43H V\x1EE4 Y T\x1EBE"), FMT_TEXT | FMT_CENTER, true, 14);

	xls.SetCellMergedColumns(nCol, nRow + 4, 10);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedRows(nCol, nRow + 5, 2);
	xls.SetCellText(nCol, nRow + 5, _T("STT"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 1, nRow + 5, 2);
	xls.SetCellText(nCol + 1, nRow + 5, _T("\x43\xE1\x63 kho\x61"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol + 2, nRow + 5, 4);
	xls.SetCellText(nCol + 2, nRow + 5, _T("T\x1ED5ng h\x1EE3p ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 2, nRow + 6, _T("N\x1ED9i tr\xFA"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 3, nRow + 6, _T("Ngo\x1EA1i tr\xFA"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 4, nRow + 6, _T("Nh\xE2n vi\xEAn"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 5, nRow + 6, _T("T\x1ED5ng"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol + 6, nRow + 5, 4);
	xls.SetCellText(nCol + 6, nRow + 5, _T("T\x1ED5ng h\x1EE3p qu\xE2n s\x1ED1"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 6, nRow + 6, _T("N\x1ED9i tr\xFA"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 7, nRow + 6, _T("Ngo\x1EA1i tr\xFA"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 8, nRow + 6, _T("Nh\xE2n vi\xEAn"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 9, nRow + 6, _T("T\x1ED5ng"), FMT_TEXT | FMT_CENTER, true, 10);
	
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
// khoa
		rs.GetValue(_T("c1"),tmpStr);
		xls.SetCellText(nCol + 1, nRow + 6, tmpStr, FMT_TEXT);
// tong hop tien noi tru
		rs.GetValue(_T("c2"),tmpStr);
		nTotal[1] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 2, nRow + 6, tmpStr, FMT_NUMBER1);
// ngoai tru
		rs.GetValue(_T("c3"), tmpStr);
		nTotal[2] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 3, nRow + 6, tmpStr, FMT_NUMBER1);
//nhan vien
		rs.GetValue(_T("c4"),tmpStr);
		nTotal[3] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 4, nRow + 6, tmpStr, FMT_NUMBER1);
// tong		
		rs.GetValue(_T("c5"), tmpStr);
		nTotal[4] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 5, nRow + 6, tmpStr, FMT_NUMBER1);
// noi tru
		rs.GetValue(_T("c6"), tmpStr);
		nTotal[5] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 6, nRow + 6, tmpStr, FMT_NUMBER1);
// ngoai tru		
		rs.GetValue(_T("c7"), tmpStr);
		nTotal[6] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 7, nRow + 6, tmpStr, FMT_NUMBER1);
// nhan vien
		rs.GetValue(_T("c8"), tmpStr);
		nTotal[7] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 8, nRow + 6, tmpStr, FMT_NUMBER1);
// tong
		rs.GetValue(_T("c9"), tmpStr);
		nTotal[8] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 9, nRow + 6, tmpStr, FMT_NUMBER1);

		nRow++;
		rs.MoveNext();
			
	}
	xls.SetCellMergedColumns(nCol, nRow + 6, 6);
	xls.SetCellText(nCol, nRow + 6, _T("T\x1ED5ng"), FMT_TEXT | FMT_CENTER, true, 11);
	for (int i = 1; i<= 7; i++)
	{
		tmpStr.Format(_T("%.0f"), nTotal[i]);
		xls.SetCellText(nCol, nRow + 6, tmpStr, FMT_NUMBER1, true, 11);
	}
	CString szDate;
	xls.Save(_T("Exports\\BAO CAO THU CHI DICH VU Y TE.xls"));	
} 
void FMServiceMedicalIncomingOutgoingReport ::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int FMServiceMedicalIncomingOutgoingReport ::OnAddHMSGeneralCostReportDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int FMServiceMedicalIncomingOutgoingReport ::OnEditHMSGeneralCostReportDialog(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int FMServiceMedicalIncomingOutgoingReport ::OnDeleteHMSGeneralCostReportDialog(){
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
 		OnCancelHMSGeneralCostReportDialog(); 
 	}
	return 0;
}
int FMServiceMedicalIncomingOutgoingReport ::OnSaveHMSGeneralCostReportDialog(){
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
 		//OnHMSGeneralCostReportDialogListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int FMServiceMedicalIncomingOutgoingReport ::OnCancelHMSGeneralCostReportDialog(){
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
int FMServiceMedicalIncomingOutgoingReport ::OnHMSGeneralCostReportDialogListLoadData(){
	return 0;
}
CString FMServiceMedicalIncomingOutgoingReport::GetQueryString()
{
	CString tmpStr, szSQL, szWhere,szUserReceiver,szFromDateToDate, szTmp, szObject;
	szWhere.Empty();	
	
	if (!m_wndClerk.GetCurrent(0).IsEmpty())
		szUserReceiver.Format(_T(" and hfi_receiver ='%s' "), m_wndClerk.GetCurrent(0));
	
	if (!m_szDeptKey.IsEmpty())
		szWhere.Format(_T(" and hfe_deptid ='%s' "), m_szDeptKey);

	if (!m_szObjectKey.IsEmpty())
		szObject.AppendFormat(_T(" and ho_type ='%s' "), m_szObjectKey);
	return szSQL;
}