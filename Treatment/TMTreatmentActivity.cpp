#include "stdafx.h"
#include "TMTreatmentActivity.h"
#include "ReportDocument.h"
#include "Excel.h"
#include "HMSMainFrame.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CTMTreatmentActivity *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMTreatmentActivity* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CTMTreatmentActivity *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CTMTreatmentActivity *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CTMTreatmentActivity *)pWnd)->OnToDateCheckValue();
} 
static void _OnDoctorSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMTreatmentActivity* )pWnd)->OnDoctorSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDoctorSelendokFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnDoctorSelendok();
}
/*static void _OnDoctorSetfocusFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnDoctorKillfocus();
}*/
/*static void _OnDoctorKillfocusFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnDoctorKillfocus();
}*/
static long _OnDoctorLoadDataFnc(CWnd *pWnd){
	return ((CTMTreatmentActivity *)pWnd)->OnDoctorLoadData();
}
/*static void _OnDoctorAddNewFnc(CWnd *pWnd){
	((CTMTreatmentActivity *)pWnd)->OnDoctorAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CTMTreatmentActivity *pVw = (CTMTreatmentActivity *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CTMTreatmentActivity *pVw = (CTMTreatmentActivity *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CTMTreatmentActivity *pVw = (CTMTreatmentActivity *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddCTMTreatmentActivityFnc(CWnd *pWnd){
	 return ((CTMTreatmentActivity*)pWnd)->OnAddCTMTreatmentActivity();
} 
static int _OnEditCTMTreatmentActivityFnc(CWnd *pWnd){
	 return ((CTMTreatmentActivity*)pWnd)->OnEditCTMTreatmentActivity();
} 
static int _OnDeleteCTMTreatmentActivityFnc(CWnd *pWnd){
	 return ((CTMTreatmentActivity*)pWnd)->OnDeleteCTMTreatmentActivity();
} 
static int _OnSaveCTMTreatmentActivityFnc(CWnd *pWnd){
	 return ((CTMTreatmentActivity*)pWnd)->OnSaveCTMTreatmentActivity();
} 
static int _OnCancelCTMTreatmentActivityFnc(CWnd *pWnd){
	 return ((CTMTreatmentActivity*)pWnd)->OnCancelCTMTreatmentActivity();
} 
CTMTreatmentActivity::CTMTreatmentActivity(CWnd *pParent)
	{

	m_nDlgWidth = 410;
	m_nDlgHeight = 185;
	SetDefaultValues();
	m_bPreview = true;
}
CTMTreatmentActivity::~CTMTreatmentActivity(){
}
void CTMTreatmentActivity::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 470, 120);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 110, 55);
	m_wndYear.Create(this,115, 30, 235, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 240, 30, 340, 55);
	m_wndReportPeriod.Create(this,345, 30, 465, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 59, 110, 84);
	m_wndFromDate.Create(this,115, 59, 235, 84); 
	m_wndToDateLabel.Create(this, _T("To Date"), 240, 60, 340, 85);
	m_wndToDate.Create(this,345, 60, 465, 85); 
	m_wndDoctorLabel.Create(this, _T("Type"), 10, 89, 110, 114);
	m_wndDoctor.Create(this,115, 89, 465, 114); 
	m_wndCheckBed.Create(this, _T("Only Bed Terminate"), 5, 125, 200, 150);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 285, 125, 385, 150);
	m_wndExport.Create(this, _T("&Export"), 390, 125, 470, 150);
}
void CTMTreatmentActivity::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(true);
//	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
//	m_wndDoctor.SetCheckValue(true);
	m_wndDoctor.LimitText(35);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 60);


	m_wndDoctor.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndDoctor.InsertColumn(1, _T("Description"), CFMT_TEXT, 110);

}
void CTMTreatmentActivity::OnSetWindowEvents(){
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
	m_wndDoctor.SetEvent(WE_SELENDOK, _OnDoctorSelendokFnc);
	//m_wndDoctor.SetEvent(WE_SETFOCUS, _OnDoctorSetfocusFnc);
	//m_wndDoctor.SetEvent(WE_KILLFOCUS, _OnDoctorKillfocusFnc);
	m_wndDoctor.SetEvent(WE_SELCHANGE, _OnDoctorSelectChangeFnc);
	m_wndDoctor.SetEvent(WE_LOADDATA, _OnDoctorLoadDataFnc);
	//m_wndDoctor.SetEvent(WE_ADDNEW, _OnDoctorAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString tmpStr;
	tmpStr = pMF->GetSysDate();
	m_nYear = ToInt(tmpStr.Left(4));
	m_szFromDate = m_szToDate = tmpStr;
	m_szToDate += _T("23:59");
	m_szReportPeriodKey.Format(_T("%d"), ToInt(tmpStr.Mid(5, 2)));
	UpdateData(false);
}
void CTMTreatmentActivity::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDoctor.GetDlgCtrlID(), m_szDoctorKey);
	DDX_Check(pDX, m_wndCheckBed.GetDlgCtrlID(), m_bCheckBed);
}
void CTMTreatmentActivity::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CTMTreatmentActivity::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CTMTreatmentActivity::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDoctorKey.Empty();

}
int CTMTreatmentActivity::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, 1, 2, -1); 
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
/*void CTMTreatmentActivity::OnYearChange(){
	
} */
/*void CTMTreatmentActivity::OnYearSetfocus(){
	
} */
/*void CTMTreatmentActivity::OnYearKillfocus(){
	
} */
int CTMTreatmentActivity::OnYearCheckValue(){
	return 0;
} 
void CTMTreatmentActivity::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMTreatmentActivity::OnReportPeriodSelendok(){
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
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59"), m_nYear,  dte.GetMonthLastDay());
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
/*void CTMTreatmentActivity::OnReportPeriodSetfocus(){
	
}*/
/*void CTMTreatmentActivity::OnReportPeriodKillfocus(){
	
}*/
long CTMTreatmentActivity::OnReportPeriodLoadData(){
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
/*void CTMTreatmentActivity::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CTMTreatmentActivity::OnFromDateChange(){
	
} */
/*void CTMTreatmentActivity::OnFromDateSetfocus(){
	
} */
/*void CTMTreatmentActivity::OnFromDateKillfocus(){
	
} */
int CTMTreatmentActivity::OnFromDateCheckValue(){
	return 0;
} 
/*void CTMTreatmentActivity::OnToDateChange(){
	
} */
/*void CTMTreatmentActivity::OnToDateSetfocus(){
	
} */
/*void CTMTreatmentActivity::OnToDateKillfocus(){
	
} */
int CTMTreatmentActivity::OnToDateCheckValue(){
	return 0;
} 
void CTMTreatmentActivity::OnDoctorSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMTreatmentActivity::OnDoctorSelendok(){
	 
}
/*void CTMTreatmentActivity::OnDoctorSetfocus(){
	
}*/
/*void CTMTreatmentActivity::OnDoctorKillfocus(){
	
}*/
long CTMTreatmentActivity::OnDoctorLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDoctor.IsSearchKey() && !m_szDoctorKey.IsEmpty()){
		szWhere.Format(_T(" and ss_code ='%s'"), m_szDoctorKey);
	};
	
	szSQL.Format(_T("SELECT ss_code as id, ss_desc as name FROM sys_sel WHERE ss_id='hms_suggestion' and ss_code in('I','O') %s "), szWhere);	
	m_wndDoctor.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDoctor.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CTMTreatmentActivity::OnDoctorAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CTMTreatmentActivity::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	m_bPreview = true;
	PrintHoatdongdieutri(m_szFromDate, m_szToDate, m_szDoctorKey, m_bPreview);
	UpdateData(false);
	
} 
void CTMTreatmentActivity::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	m_bPreview = false;
	ExportHoatdongdieutri(m_szFromDate, m_szToDate, m_szDoctorKey);
	UpdateData(false);
} 
void CTMTreatmentActivity::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

} 
int CTMTreatmentActivity::OnAddCTMTreatmentActivity(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0;
}
int CTMTreatmentActivity::OnEditCTMTreatmentActivity(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CTMTreatmentActivity::OnDeleteCTMTreatmentActivity(){
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
 		OnCancelCTMTreatmentActivity(); 
 	}
	return 0;
}
int CTMTreatmentActivity::OnSaveCTMTreatmentActivity(){
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
 		//OnCTMTreatmentActivityListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CTMTreatmentActivity::OnCancelCTMTreatmentActivity(){
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
int CTMTreatmentActivity::OnCTMTreatmentActivityListLoadData(){
	return 0;
}
void CTMTreatmentActivity::PrintHoatdongdieutri(CString szFromDate, CString szToDate, CString szDoctor, bool bPreview){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, tmpStr1, szFromDateLabel, szToDateLabel, szWhere,szAmount, szOutpatient, m_szStatus;
	CString szDate, szSysDate;
	szSysDate = pMF->GetSysDate(); 
	CReport rpt;

	szSQL = GetQueryString();

	_fmsg(_T("%s"), szSQL);
	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	if(!rpt.Init(_T("Reports/HMS/RPT_HOATDONGDIEUTRI.RPT"))) return ;
	//Report Header
	tmpStr = rpt.GetReportHeader()->GetValue(_T("ReportDate"));
	szDate.Format(tmpStr, CDate::Convert(szFromDate, yyyymmdd, ddmmyyyy),CDate::Convert(szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	//Report Detail
	CReportSection* rptDetail = rpt.GetDetail(); 
	int nIndex = 1, SongayDT=0;
	int i=0,nTotal[16];
	for (i=0;i<16;i++)
	{
		nTotal[i]=0;
	}
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();			
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("Deptname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("Totalbed"), tmpStr);
		nTotal[3] += ToInt(tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("oldpatient"), tmpStr);
		nTotal[4] += ToInt(tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("Admission"), tmpStr);
		nTotal[5] += ToInt(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rs.GetValue(_T("Tre15"), tmpStr);
		nTotal[6] += ToInt(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rs.GetValue(_T("Tre24Day"), tmpStr);
		nTotal[7] += ToInt(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rs.GetValue(_T("emergency"), tmpStr);
		nTotal[8] += ToInt(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.GetValue(_T("SongayDT"), tmpStr);
		rs.GetValue(_T("SongayDTO"), tmpStr1);
		if (m_bCheckBed) 
			SongayDT=(ToInt(tmpStr));
		else
			SongayDT=(ToInt(tmpStr) + ToInt(tmpStr1));
		nTotal[9] += SongayDT;
		tmpStr.Format(_T("%ld"), SongayDT);
		rptDetail->SetValue(_T("9"), tmpStr);				
		rs.GetValue(_T("dead"), tmpStr);
		nTotal[10] += ToInt(tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);		
		rs.GetValue(_T("dead15"), tmpStr);
		nTotal[11] += ToInt(tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);
		rs.GetValue(_T("dead24Day"), tmpStr);
		nTotal[12] += ToInt(tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);
		rs.GetValue(_T("dead24h"), tmpStr);
		nTotal[13] += ToInt(tmpStr);
		rptDetail->SetValue(_T("13"),tmpStr);
		rs.GetValue(_T("BHYT"), tmpStr);
		nTotal[14] += ToInt(tmpStr);
		rptDetail->SetValue(_T("14"), tmpStr);
		rs.GetValue(_T("remain"), tmpStr);		
		nTotal[15] += ToInt(tmpStr);
		rptDetail->SetValue(_T("15"), tmpStr);
		rs.MoveNext();
	}
	
	rptDetail = rpt.AddDetail();	
	TranslateString(_T("Total"),tmpStr);
	rptDetail->GetItem(_T("2"))->SetBold(true);
	rptDetail->SetValue(_T("2"),tmpStr);
	for (int i =3; i <= 15; i++){
		tmpStr.Format(_T("%d"), i);
		szAmount.Format(_T("%ld"),nTotal[i] );
		//rptDetail->GetItem(tmpStr)->SetBold(true);
		rptDetail->SetValue(tmpStr, szAmount);		
	}

	//Page Footer
	//Report Footer
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print();
}

void CTMTreatmentActivity::ExportHoatdongdieutri(CString szFromDate, CString szToDate, CString szDoctor)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, tmpStr1, szFromDateLabel, szToDateLabel, szWhere,szAmount, szOutpatient, m_szStatus;
	CString szDate, szSysDate;
	szSysDate = pMF->GetSysDate(); 
	CReport rpt; 
	
	szSQL = GetQueryString();

	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(1, 50);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 2);
	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname,4098,true,12,0);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name,4098,true,12,0);
	xls.SetCellMergedColumns(0,4,12);	
	xls.SetCellText(0, 4, _T("\x42\xC1O \x43\xC1O HO\x1EA0T \x110\x1ED8NG \x110I\x1EC0U TR\x1ECA"),4098, true,16,0);
	xls.SetCellMergedColumns(0,5,12);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 5, tmpStr,4098,true,12,0);
	
	
	int nRow = 7, age=0;
	xls.SetCellText(0, nRow, _T("STT"), 528386, true);	
	TranslateString(_T("Dept"), tmpStr);
	xls.SetCellText(1, nRow, tmpStr, 528386,true);
	TranslateString(_T("Bed"), tmpStr);	
	xls.SetCellText(2, nRow, tmpStr, 528386,true);	
	TranslateString(_T("Bn Dau ky"), tmpStr);
	xls.SetCellText(3, nRow, tmpStr, 528386,true);	
	TranslateString(_T("Total"), tmpStr);
	xls.SetCellText(4, nRow, tmpStr, 528386,true);	
	TranslateString(_T("Children < 6 Age"), tmpStr);
	xls.SetCellText(5, nRow, tmpStr, 528386,true);
	TranslateString(_T("Tr\x1EBB < 24 ng\xE0y"), tmpStr);
	xls.SetCellText(6, nRow, tmpStr, 528386,true);
	TranslateString(_T("Emergency"), tmpStr);
	xls.SetCellText(7, nRow, tmpStr, 528386,true);	
	TranslateString(_T("Ngay DT"), tmpStr);
	xls.SetCellText(8, nRow, tmpStr, 528386,true);
	TranslateString(_T("Tong so tu vong"), tmpStr);
	xls.SetCellText(9, nRow, tmpStr, 528386,true);	

	TranslateString(_T("Tre < 6 tuoi tu vong"), tmpStr);
	xls.SetCellText(10, nRow, tmpStr, 528386,true);

	TranslateString(_T("T\x1EED vong 24 ng\xE0y"), tmpStr);
	xls.SetCellText(11, nRow, tmpStr, 528386,true);	

	TranslateString(_T("Tu vong truoc 24 gio"), tmpStr);
	xls.SetCellText(12, nRow, tmpStr, 528386,true);	

	TranslateString(_T("BHYT"), tmpStr);
	xls.SetCellText(13, nRow, tmpStr, 528386,true);	
	
	TranslateString(_T("Ton cuoi ky"), tmpStr);
	xls.SetCellText(14, nRow, tmpStr, 528386,true);

	int nIndex = 1, SongayDT=0;
	int i=0,nTotal[15];
	for (i=0;i<15;i++)
	{
		nTotal[i]=0;
	}
	while(!rs.IsEOF())
	{
		nRow++;
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);
		rs.GetValue(_T("Deptname"), tmpStr);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("Totalbed"), tmpStr);
		nTotal[2] += ToInt(tmpStr);
		xls.SetCellText(2, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("oldpatient"), tmpStr);
		nTotal[3] += ToInt(tmpStr);
		xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("Admission"), tmpStr);
		nTotal[4] += ToInt(tmpStr);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("tre15"), tmpStr);
		nTotal[5] += ToInt(tmpStr);
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("Tre24Day"), tmpStr);
		nTotal[6] += ToInt(tmpStr);
		xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("emergency"), tmpStr);
		nTotal[7] += ToInt(tmpStr);
		xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("SongayDT"), tmpStr);
		rs.GetValue(_T("SongayDTO"), tmpStr1);
		if (m_bCheckBed) 
			SongayDT=(ToInt(tmpStr));
		else
			SongayDT=(ToInt(tmpStr) + ToInt(tmpStr1));
		nTotal[8] += SongayDT;
		tmpStr.Format(_T("%ld"), SongayDT);
		xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("dead"), tmpStr);
		nTotal[9] += ToInt(tmpStr);
		xls.SetCellText(9, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("dead15"), tmpStr);
		nTotal[10] += ToInt(tmpStr);
		xls.SetCellText(10, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("Dead24Day"), tmpStr);
		nTotal[11] += ToInt(tmpStr);
		xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("Dead24h"), tmpStr);
		nTotal[12] += ToInt(tmpStr);
		xls.SetCellText(12, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("BHYT"), tmpStr);
		nTotal[13] += ToInt(tmpStr);
		xls.SetCellText(13, nRow, tmpStr, FMT_NUMBER1);
		rs.GetValue(_T("remain"), tmpStr);		
		nTotal[14] += ToInt(tmpStr);
		xls.SetCellText(14, nRow, tmpStr, FMT_NUMBER1);
		rs.MoveNext();
	}
	
	nRow++;
	TranslateString(_T("Total"),tmpStr);
	xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true);
	for (int i =2; i <= 14; i++){		
		tmpStr.Format(_T("%d"),nTotal[i] );		
		xls.SetCellText(i, nRow, tmpStr, FMT_NUMBER1, true);
	}

	xls.Save(_T("Exports\\Bao cao hoat dong dieu tri.xls"));
	EndWaitCursor();
}

CString CTMTreatmentActivity::GetQueryString(){
	CString szSQL, szOutpatient;
	szOutpatient.Empty();
	szOutpatient.Format(_T(" (htr_outpatient ='N' or htr_outpatient is null) "));
	if (!m_szDoctorKey.IsEmpty()) 
	{
		if (m_szDoctorKey ==_T("I")) 
			szOutpatient.Format(_T(" (htr_outpatient ='N' or htr_outpatient is null) "));
		else
			szOutpatient.Format(_T(" htr_outpatient = 'Y' "));
	}

	
	szSQL.Format(_T(" SELECT deptid,") \
	_T(" 	sd_name as deptname,") \
	_T(" 	(select count(hbl_id) from hms_bedlist where hbl_deptid=deptid and hbl_active ='Y') as Totalbed , 	") \
	_T(" 	sum(oldpatient) as oldpatient,") \
	_T(" 	sum(admission) as admission,") \
	_T(" 	sum(emergency) as emergency,") \
	_T(" 	sum(Tre15) as Tre15,") \
	_T(" 	sum(dead) as dead,") \
	_T("	sum(dead15) as dead15, ") \
	_T("	sum(dead24h) as dead24h, ") \
	_T("	sum(BHYT) as BHYT,") \
	_T("	sum(Tre24Day) as Tre24Day, ") \
	_T("	sum(dead24Day) as dead24Day, ") \
	_T(" 	sum(discharge) as discharge,") \
	_T("	(select sum(hb_treatqty) from hms_bed left join hms_treatment_record on(htr_docno=hb_docno and hb_refidx=htr_idx) ") \
	_T("	where hb_deptid=deptid and htr_admitdate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and hb_status ='T') as SongayDT,") \
	_T("	(select sum(trunc(to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')) - trunc(hb_admitdate) +1) from hms_bed left join hms_treatment_record on(htr_docno=hb_docno and htr_idx = hb_refidx) ") \
	_T("	where hb_deptid=deptid and hb_admitdate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and hb_status = 'O') as SongayDTO, ") \
	_T(" 	sum(remain) as remain") \
	_T(" FROM") \
	_T(" (") \
	_T(" SELECT deptid, status, suggestion,") \
	_T(" 	case when admitdate < to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and (dischargedate >= to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') or status in('I')) then 1 else 0 end as oldpatient,") \
	_T(" 	case when admitdate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') then 1 else 0 end as admission,") \
	_T(" 	case when admitdate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and objecttype = 'C' then 1	else 0 end as tre15,") \
	_T(" 	case when emergency='Y' and admitdate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') then 1 else 0 end as emergency,") \
	_T(" 	case when dischargedate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and suggestion='D' and result in('5','6') then 1 else 0 end as dead,") \
	_T(" 	case when dischargedate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and suggestion='D' and result in('5','6') and objecttype = 'C' then 1 else 0 end as dead15,") \
	_T(" 	case when dischargedate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and suggestion='D' and result ='6' then 1 else 0 end as dead24h,") \
	_T(" 	case when dischargedate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and suggestion='D'  and result not in('5','6') then 1 else 0 end as discharge,") \
	_T(" 	case when admitdate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and ObjectType in('I','C','D') then 1 else 0 end as BHYT ,") \
	_T(" 	case when admitdate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and NumberDay <= 28 then 1 else 0 end as Tre24Day ,") \
	_T(" 	case when dischargedate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and suggestion='D' and result in('5','6') and NumberDay <= 28 then 1 else 0 end as dead24Day,") \
	_T(" 	case when dischargedate > to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') or status in('I') or (status = 'A' and treatidx > 1) then 1 else 0 end as remain") \
	_T(" FROM") \
	_T(" (") \
	_T(" 	SELECT 	htr_deptid as deptid,") \
	_T("		hms_getage(hd_admitdate, hp_birthdate) as age, ") \
	_T("		trunc(hcr_admitdate) - trunc(hp_birthdate) as NumberDay, ") \
	_T(" 		htr_status as status,") \
	_T("		ho_type as ObjectType, ") \
	_T(" 		htr_suggestion as suggestion,") \
	_T(" 		htr_emergency as emergency,") \
	_T(" 		htr_idx as treatidx,") \
	_T(" 		hcr_result as result,") \
	_T(" 		hcr_admitdate as admitdate,") \
	_T(" 		hcr_dischargedate as dischargedate") \
	_T(" 	FROM hms_clinical_record ") \
	_T(" 	LEFT JOIN hms_treatment_record ON(htr_docno=hcr_docno and hcr_refidx=htr_idx)") \
	_T(" 	LEFT JOIN hms_doc ON(hd_docno=htr_docno) ") \
	_T("	LEFT JOIN hms_patient ON(hp_patientno=hd_patientno) ") \
	_T(" 	LEFT JOIN hms_object ON(ho_id=hd_object) ") \
	_T(" 	WHERE %s ") \
	_T(" 		and (htr_admitdate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')") \
	_T(" 		or htr_dischargedate between to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')") \
	_T(" 		or (htr_admitdate <  to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and htr_status in('A','I') )") \
	_T(" 		or (htr_admitdate <  to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and htr_dischargedate > to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') and htr_status ='T') )") \
	_T(" ) tbl") \
	_T(" ) tbl2") \
	_T(" LEFT JOIN sys_dept ON(sd_id=deptid)") \
	_T(" WHERE length(sd_name) > 0 ") \
	_T(" GROUP BY deptid, sd_name") \
	_T(" ORDER BY deptid"),	
		m_szFromDate, m_szToDate,
		m_szToDate,
		m_szFromDate, m_szToDate,
		m_szFromDate, m_szFromDate,
		m_szFromDate, m_szToDate,
		m_szFromDate, m_szToDate,
		m_szFromDate, m_szToDate,
		m_szFromDate, m_szToDate,
		m_szFromDate, m_szToDate,
		m_szFromDate, m_szToDate,
		m_szFromDate, m_szToDate,
		m_szFromDate, m_szToDate,
		m_szFromDate, m_szToDate,
		m_szFromDate, m_szToDate,
		m_szToDate ,
		szOutpatient,
		m_szFromDate, m_szToDate,
		m_szFromDate, m_szToDate,
		m_szFromDate,
		m_szFromDate, m_szFromDate);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}