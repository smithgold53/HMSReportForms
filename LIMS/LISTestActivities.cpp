#include "stdafx.h"
#include "LISTestActivities.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CLISTestActitivities *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CLISTestActitivities *)pWnd)->OnToDateCheckValue();
} 
static void _OnGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CLISTestActitivities* )pWnd)->OnGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnGroupSelendokFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnGroupSelendok();
}
/*static void _OnGroupSetfocusFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnGroupKillfocus();
}*/
/*static void _OnGroupKillfocusFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnGroupKillfocus();
}*/
static long _OnGroupLoadDataFnc(CWnd *pWnd){
	return ((CLISTestActitivities *)pWnd)->OnGroupLoadData();
}
/*static void _OnGroupAddNewFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnGroupAddNew();
}*/
static void _OnDeptSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CLISTestActitivities* )pWnd)->OnDeptSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDeptSelendokFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnDeptSelendok();
}
/*static void _OnDeptSetfocusFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnDeptKillfocus();
}*/
/*static void _OnDeptKillfocusFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnDeptKillfocus();
}*/
static long _OnDeptLoadDataFnc(CWnd *pWnd){
	return ((CLISTestActitivities *)pWnd)->OnDeptLoadData();
}
/*static void _OnDeptAddNewFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnDeptAddNew();
}*/
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CLISTestActitivities* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CLISTestActitivities *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnObjectAddNew();
}*/
static void _OnStatusSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CLISTestActitivities* )pWnd)->OnStatusSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStatusSelendokFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnStatusSelendok();
}
/*static void _OnStatusSetfocusFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnStatusKillfocus();
}*/
/*static void _OnStatusKillfocusFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnStatusKillfocus();
}*/
static long _OnStatusLoadDataFnc(CWnd *pWnd){
	return ((CLISTestActitivities *)pWnd)->OnStatusLoadData();
}
/*static void _OnStatusAddNewFnc(CWnd *pWnd){
	((CLISTestActitivities *)pWnd)->OnStatusAddNew();
}*/
static void _OnByPerformDateSelectFnc(CWnd *pWnd){
	  ((CLISTestActitivities*)pWnd)->OnByPerformDateSelect();
}
static void _OnByOrderDateSelectFnc(CWnd *pWnd){
	  ((CLISTestActitivities*)pWnd)->OnByOrderDateSelect();
}
static void _OnPrintSelectFnc(CWnd *pWnd){
	CLISTestActitivities *pVw = (CLISTestActitivities *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CLISTestActitivities *pVw = (CLISTestActitivities *)pWnd;
	pVw->OnExportSelect();
} 
CLISTestActitivities::CLISTestActitivities(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CLISTestActitivities::~CLISTestActitivities(){
}
void CLISTestActitivities::OnCreateComponents(){
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 30, 300, 55);
	m_wndToDate.Create(this,305, 30, 425, 55); 
	m_wndGroupLabel.Create(this, _T("Group"), 10, 60, 90, 85);
	m_wndGroup.Create(this,95, 60, 425, 85); 
	m_wndDeptLabel.Create(this, _T("Dept"), 10, 90, 90, 115);
	m_wndTestActivities.Create(this, _T("Test Activities"), 5, 5, 430, 150);
	m_wndDept.Create(this,95, 90, 425, 115); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 120, 90, 145);
	m_wndObject.Create(this,95, 120, 215, 145); 
	m_wndStatusLabel.Create(this, _T("Status"), 220, 120, 300, 145);
	m_wndStatus.Create(this,305, 120, 425, 145); 
	m_wndByPerformDate.Create(this, _T("By Perform Date"), 5, 155, 140, 180);
	m_wndByOrderDate.Create(this, _T("By Order Date"), 145, 155, 280, 180);
	m_wndPrint.Create(this, _T("&Print"), 270, 185, 350, 210);
	m_wndExport.Create(this, _T("&Export"), 355, 185, 435, 210);

}
void CLISTestActitivities::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	//m_wndGroup.SetCheckValue(true);
	m_wndGroup.LimitText(35);
	//m_wndDept.SetCheckValue(true);
	m_wndDept.LimitText(35);
	//m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(35);
	//m_wndStatus.SetCheckValue(true);
	m_wndStatus.LimitText(35);


	m_wndGroup.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndGroup.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);


	m_wndDept.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndDept.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);


	m_wndObject.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);


	m_wndStatus.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndStatus.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

}
void CLISTestActitivities::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
	//m_wndGroup.SetEvent(WE_ADDNEW, _OnGroupAddNewFnc);
	m_wndDept.SetEvent(WE_SELENDOK, _OnDeptSelendokFnc);
	//m_wndDept.SetEvent(WE_SETFOCUS, _OnDeptSetfocusFnc);
	//m_wndDept.SetEvent(WE_KILLFOCUS, _OnDeptKillfocusFnc);
	m_wndDept.SetEvent(WE_SELCHANGE, _OnDeptSelectChangeFnc);
	m_wndDept.SetEvent(WE_LOADDATA, _OnDeptLoadDataFnc);
	//m_wndDept.SetEvent(WE_ADDNEW, _OnDeptAddNewFnc);
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndStatus.SetEvent(WE_SELENDOK, _OnStatusSelendokFnc);
	//m_wndStatus.SetEvent(WE_SETFOCUS, _OnStatusSetfocusFnc);
	//m_wndStatus.SetEvent(WE_KILLFOCUS, _OnStatusKillfocusFnc);
	m_wndStatus.SetEvent(WE_SELCHANGE, _OnStatusSelectChangeFnc);
	m_wndStatus.SetEvent(WE_LOADDATA, _OnStatusLoadDataFnc);
	//m_wndStatus.SetEvent(WE_ADDNEW, _OnStatusAddNewFnc);
	m_wndByPerformDate.SetEvent(WE_CLICK, _OnByPerformDateSelectFnc);
	m_wndByOrderDate.SetEvent(WE_CLICK, _OnByOrderDateSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	m_nByPerformDate = 0;
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T(" SELECT sup_permid perm_id FROM sys_userperm WHERE sup_userid = '%s' AND sup_moduleid = '%s'") \
				 _T(" AND substr(sup_permid, 1, 2) = 'B1'"), pMF->GetCurrentUser(), pMF->GetModuleID());
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		if (!m_szPermGroupID.IsEmpty())
			m_szPermGroupID += _T(", ");
		m_szPermGroupID.AppendFormat(_T("'%s'"), rs.GetValue(_T("perm_id")));
		rs.MoveNext();
	}
	UpdateData(false);

}
void CLISTestActitivities::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndGroup.GetDlgCtrlID(), m_szGroupKey);
	DDX_TextEx(pDX, m_wndDept.GetDlgCtrlID(), m_szDeptKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_TextEx(pDX, m_wndStatus.GetDlgCtrlID(), m_szStatusKey);
	DDX_Radio(pDX, m_wndByPerformDate.GetDlgCtrlID(), m_nByPerformDate);
	DDX_Radio(pDX, m_wndByOrderDate.GetDlgCtrlID(), m_nByOrderDate);

}
void CLISTestActitivities::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szGroupKey.Empty();
	m_szDeptKey.Empty();
	m_szObjectKey.Empty();
	m_szStatusKey.Empty();
	m_nByPerformDate=0;
	m_nByOrderDate = 1;

}
int CLISTestActitivities::SetMode(int nMode){
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
/*void CLISTestActitivities::OnFromDateChange(){
	
} */
/*void CLISTestActitivities::OnFromDateSetfocus(){
	
} */
/*void CLISTestActitivities::OnFromDateKillfocus(){
	
} */
int CLISTestActitivities::OnFromDateCheckValue(){
	return 0;
} 
/*void CLISTestActitivities::OnToDateChange(){
	
} */
/*void CLISTestActitivities::OnToDateSetfocus(){
	
} */
/*void CLISTestActitivities::OnToDateKillfocus(){
	
} */
int CLISTestActitivities::OnToDateCheckValue(){
	return 0;
} 
void CLISTestActitivities::OnGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CLISTestActitivities::OnGroupSelendok(){
	 
}
/*void CLISTestActitivities::OnGroupSetfocus(){
	
}*/
/*void CLISTestActitivities::OnGroupKillfocus(){
	
}*/
long CLISTestActitivities::OnGroupLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndGroup.IsSearchKey() && !m_szGroupKey.IsEmpty()){
		szWhere.Format(_T(" and hfg_id='%s' "), m_szGroupKey);
	}
	if (!m_szPermGroupID.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfg_id IN (%s)"), m_szPermGroupID);
	m_wndGroup.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT hfg_id as id, hfg_name as name FROM hms_fee_group WHERE substr(hfg_id , 1, 2) = 'B1' %s ORDER BY id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndGroup.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CLISTestActitivities::OnGroupAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CLISTestActitivities::OnDeptSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CLISTestActitivities::OnDeptSelendok(){
	 
}
/*void CLISTestActitivities::OnDeptSetfocus(){
	
}*/
/*void CLISTestActitivities::OnDeptKillfocus(){
	
}*/
long CLISTestActitivities::OnDeptLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDept.IsSearchKey() && !m_szDeptKey.IsEmpty()){
		szWhere.Format(_T(" and sd_id='%s' "), m_szDeptKey);
	}
	m_wndDept.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT sd_id as id, sd_name as name FROM sys_dept WHERE 1=1 %s ORDER BY id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDept.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CLISTestActitivities::OnDeptAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CLISTestActitivities::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CLISTestActitivities::OnObjectSelendok(){
	 
}
/*void CLISTestActitivities::OnObjectSetfocus(){
	
}*/
/*void CLISTestActitivities::OnObjectKillfocus(){
	
}*/
long CLISTestActitivities::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty()){
		szWhere.Format(_T(" and ho_id='%s' "), m_szObjectKey);
	}
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT ho_id as id, ho_desc as name FROM hms_object WHERE 1=1 %s ORDER BY id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CLISTestActitivities::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CLISTestActitivities::OnStatusSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CLISTestActitivities::OnStatusSelendok(){
	 
}
/*void CLISTestActitivities::OnStatusSetfocus(){
	
}*/
/*void CLISTestActitivities::OnStatusKillfocus(){
	
}*/
long CLISTestActitivities::OnStatusLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStatus.IsSearchKey() && !m_szStatusKey.IsEmpty()){
		szWhere.Format(_T(" and ss_code = '%s' "), m_szStatusKey);
	}
	szWhere.AppendFormat(_T(" AND ss_code IN ('S', 'T')"));
	m_wndStatus.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT ss_code as id, ss_desc as name FROM sys_sel WHERE ss_id = 'hms_order_status' %s ORDER BY ss_index "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStatus.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CLISTestActitivities::OnStatusAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CLISTestActitivities::OnByPerformDateSelect(){
	
}
void CLISTestActitivities::OnByOrderDateSelect(){
	
}
void CLISTestActitivities::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	/*Declaration Section*/
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL;
	double nTmp = 0;
	int nIdx = 1;
	CString szSQL, tmpStr, szReportName, szDate;
	long double nTotal = 0;
	szReportName = _T("Reports\\HMS\\HOATDONGXETNGHIEM.RPT");
	if (!rpt.Init(szReportName))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
		CDate::Convert(m_szToDate), yyyymmdd, ddmmyyyy);
		rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	}
	while (!rs.IsEOF())
	{
		/*Group Seperation*/
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("test_name")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("doc_no")));
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("patient_name")));
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("age")));
			rptDetail->SetValue(_T("6"), rs.GetValue(_T("sex")));
			rptDetail->SetValue(_T("7"), rs.GetValue(_T("obj")));
			rptDetail->SetValue(_T("8"), rs.GetValue(_T("rank")));
			rptDetail->SetValue(_T("9"), rs.GetValue(_T("dept")));
			rptDetail->SetValue(_T("10"), rs.GetValue(_T("doctor")));
			rs.GetValue(_T("order_date"), tmpStr);
			rptDetail->SetValue(_T("11"), CDateTime::Convert(tmpStr, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
			rs.GetValue(_T("perform_date"), tmpStr);
			rptDetail->SetValue(_T("12"), CDateTime::Convert(tmpStr, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
			rptDetail->SetValue(_T("13"), rs.GetValue(_T("res")));
		}
		rs.MoveNext();
	}
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
	 
} 
void CLISTestActitivities::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CExcel xls;
	CString szSQL, tmpStr;
	double nTmp = 0;
	int nIdx = 1, nRow = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 8);
	xls.SetColumnWidth(1, 30);
	xls.SetColumnWidth(3, 25);
	xls.SetColumnWidth(6, 15);
	xls.SetColumnWidth(9, 25);
	xls.SetColumnWidth(10, 15);
	xls.SetColumnWidth(11, 15);
	xls.SetColumnWidth(12, 25);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 12);
	xls.SetCellMergedColumns(0, 3, 12);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr = _T("HO\x1EA0T \x110\x1ED8NG \x58\xC9T NGHI\x1EC6M");
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	xls.SetCellText(0, 4, _T("STT"), 4098, true);
	xls.SetCellText(1, 4, _T("T\xEAn \x58N"), 4098, true);
	xls.SetCellText(2, 4, _T("S\x1ED1 HS"), 4098, true);
	xls.SetCellText(3, 4, _T("T\xEAn \x42N"), 4098, true);
	xls.SetCellText(4, 4, _T("Tu\x1ED5i"), 4098, true);
	xls.SetCellText(5, 4, _T("Gi\x1EDBi"), 4098, true);
	xls.SetCellText(6, 4, _T("\x110\x1ED1i t\x1B0\x1EE3ng"), 4098, true);
	xls.SetCellText(7, 4, _T("\x43\x1EA5p \x62\x1EAD\x63"), 4098, true);
	xls.SetCellText(8, 4, _T("Kho\x61"), 4098, true);
	xls.SetCellText(9, 4, _T("\x42\xE1\x63 s\x129"), 4098, true);
	xls.SetCellText(10, 4, _T("Ng\xE0y y/\x63"), 4098, true);
	xls.SetCellText(11, 4, _T("Ng\xE0y th\x1EF1\x63 hi\x1EC7n"), 4098, true);
	xls.SetCellText(12, 4, _T("K\x1EBFt qu\x1EA3"), 4098, true);
	xls.SetCellText(13, 4, _T("Barcode"), 4098, true);
	xls.SetCellText(14, 4, _T("S\x1ED1 th\x1EBB"), 4098, true);
	xls.SetCellText(15, 4, _T("K\x1EBFt lu\x1EADn"), 4098, true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("test_name")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("doc_no")), FMT_TEXT);
		xls.SetCellText(3, nRow, rs.GetValue(_T("patient_name")), FMT_TEXT);
		xls.SetCellText(4, nRow, rs.GetValue(_T("age")), FMT_TEXT);
		xls.SetCellText(5, nRow, rs.GetValue(_T("sex")), FMT_TEXT);
		xls.SetCellText(6, nRow, rs.GetValue(_T("obj")), FMT_TEXT);
		xls.SetCellText(7, nRow, rs.GetValue(_T("rank")), FMT_TEXT);
		xls.SetCellText(8, nRow, rs.GetValue(_T("dept")), FMT_TEXT);
		xls.SetCellText(9, nRow, rs.GetValue(_T("doctor")), FMT_TEXT);
		tmpStr = CDateTime::Convert(rs.GetValue(_T("order_date")), yyyymmdd|hhmmss, ddmmyyyy|hhmmss);
		xls.SetCellText(10, nRow, tmpStr, FMT_TEXT);
		tmpStr = CDateTime::Convert(rs.GetValue(_T("perform_date")), yyyymmdd|hhmmss, ddmmyyyy|hhmmss);
		xls.SetCellText(11, nRow, tmpStr, FMT_TEXT);
		xls.SetCellText(12, nRow, rs.GetValue(_T("res")), FMT_TEXT);
		xls.SetCellText(13, nRow, rs.GetValue(_T("bar_code")), FMT_TEXT);
		xls.SetCellText(14, nRow, rs.GetValue(_T("card_no")), FMT_TEXT);
		xls.SetCellText(15, nRow, rs.GetValue(_T("conclusion")), FMT_TEXT);
		nRow++;
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\Hoat dong xet nghiem.xls"));
} 

CString CLISTestActitivities::GetQueryString(){
	CString szSQL, szWhere, tmpStr;
	if (m_nByPerformDate == 0)
		szWhere.AppendFormat(_T(" AND hpc_performdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
		_T(" AND hpc_status IN ('P', 'T')")
		, m_szFromDate, m_szToDate);
	else if (m_nByOrderDate == 0)
		szWhere.AppendFormat(_T(" AND hpc_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
		, m_szFromDate, m_szToDate);
	if (!m_szGroupKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpc_groupid = '%s'"), m_szGroupKey);
	else
		szWhere.AppendFormat(_T(" AND hpc_groupid IN (%s)"), m_szPermGroupID);
	if (!m_szDeptKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpc_deptid = '%s'"), m_szDeptKey);
	if (!m_szObjectKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpc_object = %s"), m_szObjectKey);
	tmpStr.Empty();
	if (!m_szStatusKey.IsEmpty())
		if (m_szStatusKey == _T("S"))
			tmpStr = _T("'S'");
		else
			tmpStr = _T("'P', 'T'");
	if (!tmpStr.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpc_status IN (%s)"), tmpStr);
	szSQL.Format(_T(" SELECT    hpc_docno                            doc_no, ") \
				_T("           Get_patientname(hpc_docno)           patient_name, ") \
				_T("           Hms_getagebydoc(hpc_docno)           age, ") \
				_T("           Get_syssel_desc('sys_sex', hp_sex)   sex, ") \
				_T("           (SELECT ho_desc ") \
				_T("            FROM   hms_object ") \
				_T("            WHERE  ho_id = hd_object)           obj, ") \
				_T("           Get_syssel_desc('hms_rank', hd_rank) rank, ") \
				_T("           hfl_name                             test_name, ") \
				_T("           hpc_deptid                           dept, ") \
				_T("           Get_username(hpc_doctor)             doctor, ") \
				_T("           hpc_orderdate                        order_date, ") \
				_T("           hpc_performdate                      perform_date, ") \
				_T("           hpcl_result                          res, ") \
				_T("		   hd_cardno card_no,") \
				_T("		   hd_conclusion conclusion,") \
				_T("		   hpc_sid bar_code") \
				_T(" FROM      hms_testorder ") \
				_T(" LEFT JOIN hms_doc ON ( hd_docno = hpc_docno ) ") \
				_T(" LEFT JOIN hms_patient ON ( hd_patientno = hp_patientno ) ") \
				_T(" LEFT JOIN hms_testorderline ON ( hpc_orderid = hpcl_orderid ) ") \
				_T(" LEFT JOIN hms_fee_list ON ( hfl_feeid = hpcl_itemid ) ") \
				_T(" WHERE     hpcl_hasfee = 'Y' %s") \
				_T(" ORDER BY hpc_orderdate, hpc_docno"), szWhere);

	return szSQL;
}
