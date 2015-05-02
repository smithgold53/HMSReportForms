#include "stdafx.h"
#include "LIMSPatientList.h"
#include "HMSMainFrame.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CLIMSPatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CLIMSPatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnDepartmentSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CLIMSPatientList* )pWnd)->OnDepartmentSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDepartmentSelendokFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnDepartmentSelendok();
}
/*static void _OnDepartmentSetfocusFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnDepartmentKillfocus();
}*/
/*static void _OnDepartmentKillfocusFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnDepartmentKillfocus();
}*/
static long _OnDepartmentLoadDataFnc(CWnd *pWnd){
	return ((CLIMSPatientList *)pWnd)->OnDepartmentLoadData();
}
/*static void _OnDepartmentAddNewFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnDepartmentAddNew();
}*/
static void _OnGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CLIMSPatientList* )pWnd)->OnGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnGroupSelendokFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnGroupSelendok();
}
/*static void _OnGroupSetfocusFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnGroupKillfocus();
}*/
/*static void _OnGroupKillfocusFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnGroupKillfocus();
}*/
static long _OnGroupLoadDataFnc(CWnd *pWnd){
	return ((CLIMSPatientList *)pWnd)->OnGroupLoadData();
}
/*static void _OnGroupAddNewFnc(CWnd *pWnd){
	((CLIMSPatientList *)pWnd)->OnGroupAddNew();
}*/
static void _OnOrderDateSelectFnc(CWnd *pWnd){
	 ((CLIMSPatientList*)pWnd)->OnOrderDateSelect();
}
static void _OnPerformDateSelectFnc(CWnd *pWnd){
	 ((CLIMSPatientList*)pWnd)->OnPerformDateSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CLIMSPatientList *pVw = (CLIMSPatientList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CLIMSPatientList *pVw = (CLIMSPatientList *)pWnd;
	pVw->OnExportSelect();
} 
CLIMSPatientList::CLIMSPatientList(CWnd *pParent){

	m_nDlgWidth = 585;
	m_nDlgHeight = 585;
	SetDefaultValues();
}
CLIMSPatientList::~CLIMSPatientList(){
}
void CLIMSPatientList::OnCreateComponents(){
	m_wndParaclinicalPatientList.Create(this, _T("Paraclinical Patient List"), 5, 5, 580, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 295, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 300, 30, 380, 55);
	m_wndToDate.Create(this,385, 30, 575, 55); 
	m_wndDepartmentLabel.Create(this, _T("Department"), 10, 60, 90, 85);
	m_wndDepartment.Create(this,95, 60, 295, 85); 
	m_wndGroupLabel.Create(this, _T("Group"), 300, 60, 380, 85);
	m_wndGroup.Create(this,385, 60, 575, 85); 
	m_wndOrderDate.Create(this, _T("Order Date"), 5, 95, 125, 120);
	m_wndPerformDate.Create(this, _T("Perform Date"), 130, 95, 250, 120);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 390, 95, 490, 120);
	m_wndExport.Create(this, _T("&Export"), 495, 95, 575, 120);

}
void CLIMSPatientList::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndDepartment.SetCheckValue(true);
	m_wndDepartment.LimitText(35);
	m_wndGroup.SetCheckValue(true);
	m_wndGroup.LimitText(35);


	m_wndDepartment.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndDepartment.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);


	m_wndGroup.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndGroup.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CLIMSPatientList::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndDepartment.SetEvent(WE_SELENDOK, _OnDepartmentSelendokFnc);
	//m_wndDepartment.SetEvent(WE_SETFOCUS, _OnDepartmentSetfocusFnc);
	//m_wndDepartment.SetEvent(WE_KILLFOCUS, _OnDepartmentKillfocusFnc);
	m_wndDepartment.SetEvent(WE_SELCHANGE, _OnDepartmentSelectChangeFnc);
	m_wndDepartment.SetEvent(WE_LOADDATA, _OnDepartmentLoadDataFnc);
	//m_wndDepartment.SetEvent(WE_ADDNEW, _OnDepartmentAddNewFnc);
	m_wndGroup.SetEvent(WE_SELENDOK, _OnGroupSelendokFnc);
	//m_wndGroup.SetEvent(WE_SETFOCUS, _OnGroupSetfocusFnc);
	//m_wndGroup.SetEvent(WE_KILLFOCUS, _OnGroupKillfocusFnc);
	m_wndGroup.SetEvent(WE_SELCHANGE, _OnGroupSelectChangeFnc);
	m_wndGroup.SetEvent(WE_LOADDATA, _OnGroupLoadDataFnc);
	//m_wndGroup.SetEvent(WE_ADDNEW, _OnGroupAddNewFnc);
	m_wndOrderDate.SetEvent(WE_CLICK, _OnOrderDateSelectFnc);
	m_wndPerformDate.SetEvent(WE_CLICK, _OnPerformDateSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(FALSE);

}
void CLIMSPatientList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDepartment.GetDlgCtrlID(), m_szDepartmentKey);
	DDX_TextEx(pDX, m_wndGroup.GetDlgCtrlID(), m_szGroupKey);
	DDX_Check(pDX, m_wndOrderDate.GetDlgCtrlID(), m_bOrderDate);
	DDX_Check(pDX, m_wndPerformDate.GetDlgCtrlID(), m_bPerformDate);

}
void CLIMSPatientList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDepartmentKey.Empty();
	m_szGroupKey.Empty();
	m_bOrderDate=TRUE;
	m_bPerformDate=FALSE;

}
int CLIMSPatientList::SetMode(int nMode){
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
/*void CLIMSPatientList::OnFromDateChange(){
	
} */
/*void CLIMSPatientList::OnFromDateSetfocus(){
	
} */
/*void CLIMSPatientList::OnFromDateKillfocus(){
	
} */
int CLIMSPatientList::OnFromDateCheckValue(){
	return 0;
} 
/*void CLIMSPatientList::OnToDateChange(){
	
} */
/*void CLIMSPatientList::OnToDateSetfocus(){
	
} */
/*void CLIMSPatientList::OnToDateKillfocus(){
	
} */
int CLIMSPatientList::OnToDateCheckValue(){
	return 0;
} 
void CLIMSPatientList::OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CLIMSPatientList::OnDepartmentSelendok(){
	 
}
/*void CLIMSPatientList::OnDepartmentSetfocus(){
	
}*/
/*void CLIMSPatientList::OnDepartmentKillfocus(){
	
}*/
long CLIMSPatientList::OnDepartmentLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDepartment.IsSearchKey() && !m_szDepartmentKey.IsEmpty()){
		szWhere.Format(_T(" and sd_id='%s' "), m_szDepartmentKey);
	}
	m_wndDepartment.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT sd_id as id, sd_name as name FROM sys_dept WHERE sd_type IN ('KB', 'DT') %s ORDER BY sd_id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDepartment.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CLIMSPatientList::OnDepartmentAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CLIMSPatientList::OnGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CLIMSPatientList::OnGroupSelendok(){
	 
}
/*void CLIMSPatientList::OnGroupSetfocus(){
	
}*/
/*void CLIMSPatientList::OnGroupKillfocus(){
	
}*/
long CLIMSPatientList::OnGroupLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndGroup.IsSearchKey() && !m_szGroupKey.IsEmpty()){
		szWhere.Format(_T(" and hfg_id='%s' "), m_szGroupKey);
	}
	m_wndGroup.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT hfg_id id, hfg_name name FROM hms_fee_group WHERE substr(hfg_id, 1, 2) IN('B1') %s ORDER BY hfg_index, hfg_id"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndGroup.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CLIMSPatientList::OnGroupAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CLIMSPatientList::OnOrderDateSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndPerformDate.SetCheck(FALSE);
}
void CLIMSPatientList::OnPerformDateSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndOrderDate.SetCheck(FALSE);
}
void CLIMSPatientList::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CReportSection *rptDetail = NULL, *rptGroup = NULL;
	CString szSQL, tmpStr, szOldGroup, szNewGroup, szOldDoc, szNewDoc, szPrintDate;
	double nAmt = 0;
	long double nPatAmt = 0, nGrpAmt = 0, nTtlAmt = 0;
	int nIdx = 1;
	if (!rpt.Init(_T("Reports\\HMS\\LIMS_DSBNLAMCLS.RPT")))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss),
		CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("hfg_id"), szNewGroup);
		if (szNewGroup != szOldGroup)
		{
			if (nPatAmt > 0)
			{
				rptGroup = rpt.AddDetail(rpt.GetGroupFooter(0));
				rptGroup->SetValue(_T("TotalAmount"), _T("T\x1ED5ng BN"));
				tmpStr.Format(_T("%f"), nPatAmt);
				rptGroup->SetValue(_T("s5"), tmpStr);
				nGrpAmt += nPatAmt;
				nPatAmt = 0;
			}
			if (nGrpAmt > 0)
			{
				rptGroup = rpt.AddDetail(rpt.GetGroupFooter(0));
				rptGroup->SetValue(_T("TotalAmount"), _T("T\x1ED5ng nh\xF3m"));
				tmpStr.Format(_T("%f"), nGrpAmt);
				rptGroup->SetValue(_T("s5"), tmpStr);
				nTtlAmt += nGrpAmt;
				nGrpAmt = 0;
			}
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));
			rptGroup->SetValue(_T("GroupName"), rs.GetValue(_T("hfg_name")));
			szOldGroup = szNewGroup;
			szOldDoc.Empty();
		}
		rs.GetValue(_T("hfe_docno"), szNewDoc);
		if (szNewDoc != szOldDoc)
		{
			if (nPatAmt > 0)
			{
				rptGroup = rpt.AddDetail(rpt.GetGroupFooter(0));
				rptGroup->SetValue(_T("TotalAmount"), _T("T\x1ED5ng BN"));
				tmpStr.Format(_T("%f"), nPatAmt);
				rptGroup->SetValue(_T("s5"), tmpStr);
				nGrpAmt += nPatAmt;
				nPatAmt = 0;
			}
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));
			tmpStr.Format(_T("%s - %s - %s"), rs.GetValue(_T("hfe_docno")), rs.GetValue(_T("patientname")), rs.GetValue(_T("hd_icd")));
			rptGroup->SetValue(_T("GroupName"), tmpStr);
			szOldDoc = szNewDoc;
		}
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("hfl_name")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("hfe_unitprice")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("qty")));
		rs.GetValue(_T("amount"), nAmt);
		nPatAmt += nAmt;
		rptDetail->SetValue(_T("5"), double2str(nAmt));
		rs.MoveNext();
	}
	if (nPatAmt > 0)
	{
		rptGroup = rpt.AddDetail(rpt.GetGroupFooter(0));
		rptGroup->SetValue(_T("TotalAmount"), _T("T\x1ED5ng BN"));
		tmpStr.Format(_T("%f"), nPatAmt);
		rptGroup->SetValue(_T("s5"), tmpStr);
		nGrpAmt += nPatAmt;
	}
	if (nGrpAmt > 0)
	{
		rptGroup = rpt.AddDetail(rpt.GetGroupFooter(0));
		rptGroup->SetValue(_T("TotalAmount"), _T("T\x1ED5ng nh\xF3m"));
		tmpStr.Format(_T("%f"), nGrpAmt);
		rptGroup->SetValue(_T("s5"), tmpStr);
		nTtlAmt += nGrpAmt;
	}
	szPrintDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szPrintDate.Mid(8, 2), szPrintDate.Mid(5, 2), szPrintDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CLIMSPatientList::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
} 

CString CLIMSPatientList::GetQueryString(){
	CString szSQL, szWhere;
	if (m_bOrderDate)
		szWhere.Format(_T(" AND hpc_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	else
		szWhere.Format(_T(" AND hpc_performdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szDepartmentKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hpc_deptid = '%s'"), m_szDepartmentKey);
	}
	if (!m_szGroupKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hfe_group = '%s'"), m_szGroupKey);
	}
	szSQL.Format(_T(" SELECT   hfg_id, ") \
				_T("           hfg_name, ") \
				_T("           hfe_docno, ") \
				_T("           Get_patientname(hfe_docno)        patientname, ") \
				_T("		   hd_icd,") \
				_T("           hfl_name, ") \
				_T("           hfe_unitprice, ") \
				_T("           SUM(hfe_quantity)                 qty, ") \
				_T("           SUM(hfe_unitprice * hfe_quantity) amount ") \
				_T(" FROM      hms_testorder ") \
				_T(" LEFT JOIN hms_doc ON (hd_docno = hpc_docno)") \
				_T(" LEFT JOIN hms_fee ON ( hfe_type = 'T' ") \
				_T("                        AND hpc_docno = hfe_docno ") \
				_T("                        AND hpc_orderid = hfe_orderid ) ") \
				_T(" LEFT JOIN hms_fee_list ON ( hfe_itemid = hfl_feeid ) ") \
				_T(" LEFT JOIN hms_fee_group ON ( hfg_id = hfe_group ) ") \
				_T(" WHERE 1=1 %s") \
				_T(" GROUP     BY hfl_name, ") \
				_T("              hfe_unitprice, ") \
				_T("              hfe_docno, ") \
				_T("			  hd_icd,") \
				_T("              hfg_id, ") \
				_T("              hfg_name") \
				_T(" ORDER BY hfg_id, hfe_docno"), szWhere);
	return szSQL;
}