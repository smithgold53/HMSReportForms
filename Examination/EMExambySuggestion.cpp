#include "stdafx.h"
#include "EMExambySuggestion.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMExambySuggestion *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){ 
	((CEMExambySuggestion *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMExambySuggestion *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExambySuggestion *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMExambySuggestion *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMExambySuggestion *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMExambySuggestion *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExambySuggestion *)pWnd)->OnToDateCheckValue();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMExambySuggestion* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CEMExambySuggestion *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CEMExambySuggestion *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CEMExambySuggestion *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CEMExambySuggestion *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CEMExambySuggestion *)pWnd)->OnObjectAddNew();
}*/

static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMExambySuggestion *pVw = (CEMExambySuggestion *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMExambySuggestion *pVw = (CEMExambySuggestion *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMExambySuggestion*)pWnd)->OnAddEMInwardListReport();
} 
static int _OnEditEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMExambySuggestion*)pWnd)->OnEditEMInwardListReport();
} 
static int _OnDeleteEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMExambySuggestion*)pWnd)->OnDeleteEMInwardListReport();
} 
static int _OnSaveEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMExambySuggestion*)pWnd)->OnSaveEMInwardListReport();
} 
static int _OnCancelEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMExambySuggestion*)pWnd)->OnCancelEMInwardListReport();
} 
CEMExambySuggestion::CEMExambySuggestion(CWnd *pParent){

	/*m_nDlgWidth = 500;
	m_nDlgHeight = 130;
	SetDefaultValues();*/

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CEMExambySuggestion::~CEMExambySuggestion(){
}
void CEMExambySuggestion::OnCreateComponents()
{
	/*m_wndInwardListReport.Create(this, _T("Inward List Report"), 5, 5, 490, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 60, 90, 85);
	m_wndObject.Create(this,95, 60, 245, 85); 
	m_wndSectionLabel.Create(this, _T("Section"), 250, 60, 330, 85);
	m_wndSection.Create(this,335, 60, 485, 85); 
	m_wndPrint.Create(this, _T("&Print"), 335, 95, 410, 120);
	m_wndExport.Create(this, _T("&Export"), 415, 95, 490, 120);*/

	m_wndInwardListReport.Create(this, _T("Exam by Suggestion"), 5, 5, 490, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 60, 90, 85);
	m_wndObject.Create(this,95, 60, 485, 85); 
	m_wndPrint.Create(this, _T("&Print"), 335, 95, 410, 120);
	m_wndExport.Create(this, _T("&Export"), 415, 95, 490, 120);
	m_wndExport.EnableWindow(FALSE);
}
void CEMExambySuggestion::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	//m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(35);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndObject.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);
}
void CEMExambySuggestion::OnSetWindowEvents(){
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CEMExambySuggestion::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
}
void CEMExambySuggestion::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMExambySuggestion::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMExambySuggestion::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szObjectKey.Empty();
	

}
int CEMExambySuggestion::SetMode(int nMode){
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
/*void CEMExambySuggestion::OnFromDateChange(){
	
} */
/*void CEMExambySuggestion::OnFromDateSetfocus(){
	
} */
/*void CEMExambySuggestion::OnFromDateKillfocus(){
	
} */
int CEMExambySuggestion::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMExambySuggestion::OnToDateChange(){
	
} */
/*void CEMExambySuggestion::OnToDateSetfocus(){
	
} */
/*void CEMExambySuggestion::OnToDateKillfocus(){
	
} */
int CEMExambySuggestion::OnToDateCheckValue(){
	return 0;
} 
void CEMExambySuggestion::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMExambySuggestion::OnObjectSelendok(){
	 
}
/*void CEMExambySuggestion::OnObjectSetfocus(){
	
}*/
/*void CEMExambySuggestion::OnObjectKillfocus(){
	
}*/
long CEMExambySuggestion::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty()){
		//szWhere.Format();
	};
	szSQL.Format(_T("SELECT ho_id as id, ho_desc as description FROM hms_object ORDER BY cast(id as integer)"));
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMExambySuggestion::OnObjectAddNew()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
} */

void CEMExambySuggestion::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szPos;
	CReportSection *rptDetail = NULL;
	double nLine_total = 0, nTmp = 0;
	int nCount = 0, nRet = 0;
	//QueryOpener(szSQL);
	if (!rpt.Init(_T("Reports\\HMS\\THONGKEBENHNHANTHEOHUONGDIEUTRIKHOAC13.RPT")))
		return;
	//So tu vong
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" WHERE hd_enddate BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s') ") \
				_T(" AND hd_suggestion IN ('X', 'Y') AND hd_status = 'T' AND hd_enddept = 'C1.3'") \
				_T(" ) tbl2") \
				_T(" HAVING SUM(c13) > 0"), m_szFromDate, m_szToDate);
	nRet = rs.ExecSQL(szSQL);
	if (nRet < 0)
		_msg(_T("%s"), szSQL);
	nCount = rs.GetRecordCount();
	
	if (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), _T("T\x1EED vong"));
		for (int i = 0; i < 12; i++)
		{
			tmpStr.Format(_T("c%d"), i+1);
			rs.GetValue(tmpStr, nTmp);
			nLine_total += nTmp;
			szPos.Format(_T("%d"), i+2);
			rptDetail->SetValue(szPos, double2str(nTmp));
		}
		rptDetail->SetValue(_T("14"), double2str(nLine_total));
	}
	nLine_total = 0;//reinit
	//So ke don
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hpo_object_id as policy,") \
				_T("         case when hpo_object_id=1 then 1 else 0 end as c1,") \
				_T("         case when hpo_object_id=2 then 1 else 0 end as c2,") \
				_T("         case when hpo_object_id=3 then 1 else 0 end as c3,") \
				_T("         case when hpo_object_id=4 then 1 else 0 end as c4,") \
				_T("         case when hpo_object_id=5 then 1 else 0 end as c5,") \
				_T("         case when hpo_object_id=6 then 1 else 0 end as c6,") \
				_T("         case when hpo_object_id=7 then 1 else 0 end as c7,") \
				_T("         case when hpo_object_id=8 then 1 else 0 end as c8,") \
				_T("         case when hpo_object_id=9 then 1 else 0 end as c9,") \
				_T("         case when hpo_object_id=10 then 1 else 0 end as c10,") \
				_T("         case when hpo_object_id=11 then 1 else 0 end as c11,") \
				_T("         case when hpo_object_id=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM (") \
				_T(" SELECT distinct hpo_docno, hpo_object_id") \
				_T(" FROM hms_pharmaorder") \
				_T(" LEFT JOIN m_storagelist ON (msl_storage_id = hpo_storage_id)") \
				_T(" WHERE hpo_orderdate BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s') AND msl_type <> 'C'") \
				_T(" AND hpo_orderstatus NOT IN ('O', 'C') AND hpo_deptid = 'C1.3' AND hpo_ordertype = 'P' AND NVL(hpo_feetype, 'P')<> 'PT') tbl1") \
				_T(" ) tbl2") \
				_T(" HAVING SUM(c13) > 0"), m_szFromDate, m_szToDate);
	nRet = rs.ExecSQL(szSQL);
	if (nRet < 0)
		_msg(_T("%s"), szSQL);
	nCount += rs.GetRecordCount();
	
	if (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), _T("S\x1ED1 k\xEA \x111\x1A1n"));
		for (int i = 0; i < 12; i++)
		{
			tmpStr.Format(_T("c%d"), i+1);
			rs.GetValue(tmpStr, nTmp);
			nLine_total += nTmp;
			szPos.Format(_T("%d"), i+2);
			rptDetail->SetValue(szPos, double2str(nTmp));
		}
		rptDetail->SetValue(_T("14"), double2str(nLine_total));
	}
	if (nCount <= 0)
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);	
	
	CString szSysDate;
	szSysDate = pMF->GetSysDate(); 
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CEMExambySuggestion::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));

} 
int CEMExambySuggestion::OnAddEMInwardListReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMExambySuggestion::OnEditEMInwardListReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMExambySuggestion::OnDeleteEMInwardListReport(){
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
 		OnCancelEMInwardListReport(); 
 	}
	return 0;
}
int CEMExambySuggestion::OnSaveEMInwardListReport(){
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
 		//OnEMInwardListReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMExambySuggestion::OnCancelEMInwardListReport(){
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
int CEMExambySuggestion::OnEMInwardListReportListLoadData(){
	return 0;
}

CString CEMExambySuggestion::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND he_examdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szObjectKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hd_object = '%s'"), m_szObjectKey);

	szSQL.Format(_T(" SELECT") \
				_T("  distinct he_docno  AS docno, he_doctor as doctor,") \
				_T("   he_examdate,") \
				_T("   Trim(hp_surname||' '||hp_midname||' '||hp_firstname)            AS pname,") \
				_T("   Extract(YEAR FROM hp_birthdate) AS yob,") \
				_T("   (SELECT DISTINCT ss_desc FROM   sys_sel WHERE  ss_id='hms_rank' AND ss_code=hp_rank)    AS rank,") \
				_T("   hp_workplace                    AS workplace,") \
				_T("   he_diagnostic                   AS diagno,") \
				_T("   CASE WHEN hd_suggestion IN ('C', 'D')") \
				_T("        AND Length(hd_indept)>0 THEN hd_indept") \
				_T("   END                             AS inward,") \
				_T("	CASE WHEN hd_suggestion in ('E', 'A') then 'X'  else '' end as home,") \
				_T("   hd_cardno                       AS cardno,") \
				_T("   get_objectname(hd_object) as obj") \
				_T(" FROM   hms_patient") \
				_T(" LEFT JOIN hms_doc ON (hd_patientno=hp_patientno)") \
				_T(" LEFT JOIN hms_exam ON (he_docno=hd_docno)") \
				_T(" WHERE  he_status IN ('P', 'T') AND hd_enddept = 'C1.3' %s") \
				_T(" ORDER  BY he_examdate "), szWhere); 


	/*szSQL.Format(_T(" SELECT") \
				_T("  distinct he_docno                        AS docno, he_doctor as doctor,") \
				_T("   he_examdate,") \
				_T("   Trim(hp_surname||' '||hp_midname||' '||hp_firstname)            AS pname,") \
				_T("   Extract(YEAR FROM hp_birthdate) AS yob,") \
				_T("   (SELECT DISTINCT ss_desc FROM   sys_sel WHERE  ss_id='hms_rank' AND ss_code=hp_rank)    AS rank,") \
				_T("   hp_workplace                    AS workplace,") \
				_T("   he_diagnostic                   AS diagno,") \
				_T("   CASE WHEN hd_suggestion='D'") \
				_T("        AND Length(hd_indept)>0 THEN hd_indept") \
				_T("   END                             AS inward,") \
				_T("	CASE WHEN hd_suggestion in ('E', 'A') then 'X'  else '' end as home,") \
				_T("   hd_cardno                       AS cardno,") \
				_T("   get_objectname(hd_object) as obj") \
				_T(" FROM   hms_patient") \
				_T("        LEFT JOIN hms_doc ON (hd_patientno=hp_patientno)") \
				_T("        LEFT JOIN hms_exam ON (he_docno=hd_docno)") \
				_T(" WHERE  he_status IN ('P', 'T') AND he_deptid in ('C1.3', 'C1.2') and hp_rank in ('13', '14', '15', '16', '17', '18') %s") \
				_T(" ORDER  BY he_examdate "), szWhere);*/


	//QueryOpener(szSQL);
	return szSQL;
}