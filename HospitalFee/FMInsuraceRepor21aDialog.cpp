#include "stdafx.h"
#include "FMInsuraceRepor21aDialog.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CHMSInsuraceRepor21aDialog *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CHMSInsuraceRepor21aDialog* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CHMSInsuraceRepor21aDialog *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CHMSInsuraceRepor21aDialog *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CHMSInsuraceRepor21aDialog *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CHMSInsuraceRepor21aDialog* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CHMSInsuraceRepor21aDialog *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CHMSInsuraceRepor21aDialog *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CHMSInsuraceRepor21aDialog *pVw = (CHMSInsuraceRepor21aDialog *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CHMSInsuraceRepor21aDialog *pVw = (CHMSInsuraceRepor21aDialog *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CHMSInsuraceRepor21aDialog *pVw = (CHMSInsuraceRepor21aDialog *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CHMSInsuraceRepor21aDialog *pVw = (CHMSInsuraceRepor21aDialog *)pWnd;
	pVw->OnCloseSelect();
} 
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	 ((CHMSInsuraceRepor21aDialog*)pWnd)->OnOutPatientSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	 ((CHMSInsuraceRepor21aDialog*)pWnd)->OnInPatientSelect();
}
static int _OnAddHMSInsuraceRepor21aDialogFnc(CWnd *pWnd){
	 return ((CHMSInsuraceRepor21aDialog*)pWnd)->OnAddHMSInsuraceRepor21aDialog();
} 
static int _OnEditHMSInsuraceRepor21aDialogFnc(CWnd *pWnd){
	 return ((CHMSInsuraceRepor21aDialog*)pWnd)->OnEditHMSInsuraceRepor21aDialog();
} 
static int _OnDeleteHMSInsuraceRepor21aDialogFnc(CWnd *pWnd){
	 return ((CHMSInsuraceRepor21aDialog*)pWnd)->OnDeleteHMSInsuraceRepor21aDialog();
} 
static int _OnSaveHMSInsuraceRepor21aDialogFnc(CWnd *pWnd){
	 return ((CHMSInsuraceRepor21aDialog*)pWnd)->OnSaveHMSInsuraceRepor21aDialog();
} 
static int _OnCancelHMSInsuraceRepor21aDialogFnc(CWnd *pWnd){
	 return ((CHMSInsuraceRepor21aDialog*)pWnd)->OnCancelHMSInsuraceRepor21aDialog();
} 
CHMSInsuraceRepor21aDialog::CHMSInsuraceRepor21aDialog(CWnd *pParent)
	{

	m_nDlgWidth = 430;
	m_nDlgHeight = 225;
	SetDefaultValues();
}
CHMSInsuraceRepor21aDialog::~CHMSInsuraceRepor21aDialog(){
}
void CHMSInsuraceRepor21aDialog::OnCreateComponents(){
	m_wndReportFilter.Create(this, _T("Report Parameters"), 5, 5, 420, 150);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 110, 55);
	m_wndYear.Create(this,115, 30, 210, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Month"), 215, 30, 315, 55);
	m_wndReportPeriod.Create(this,320, 30, 415, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 110, 85);
	m_wndFromDate.Create(this,115, 60, 210, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 215, 60, 315, 85);
	m_wndToDate.Create(this,320, 60, 415, 85); 
	//m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 110, 115);	
	//m_wndClerk.Create(this,115, 90, 415, 115);
	m_wndDrugPTTT.Create(this, _T("Kh\xF4ng l\x1EA5y VTTH trong PTTT v\xE0o \x62\xE1o \x63\xE1o"), 115, 90, 415, 115);
	m_wndOutPatient.Create(this, _T("OutPatient"), 115, 120, 250, 145);
	m_wndInPatient.Create(this, _T("InPatient"), 255, 120, 385, 145);	
	m_wndPrint.Create(this, _T("&Print Preview"), 255, 155, 335, 180);
	m_wndExport.Create(this, _T("&Export"), 340, 155, 420, 180);
}
void CHMSInsuraceRepor21aDialog::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(FALSE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 60);


	/*m_wndClerk.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);*/
	
}
void CHMSInsuraceRepor21aDialog::OnSetWindowEvents(){
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
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	UpdateData(false);

}
void CHMSInsuraceRepor21aDialog::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_Check(pDX, m_wndDrugPTTT.GetDlgCtrlID(), m_bCheckDrug);
	DDX_Check(pDX, m_wndOutPatient.GetDlgCtrlID(), m_bOutPatient);
	DDX_Check(pDX, m_wndInPatient.GetDlgCtrlID(), m_bInPatient);

}
void CHMSInsuraceRepor21aDialog::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CHMSInsuraceRepor21aDialog::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CHMSInsuraceRepor21aDialog::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bOutPatient=FALSE;
	m_bInPatient=FALSE;

}
int CHMSInsuraceRepor21aDialog::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0,1,2,3, 4, -1); 
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
/*void CHMSInsuraceRepor21aDialog::OnYearChange(){
	
} */
/*void CHMSInsuraceRepor21aDialog::OnYearSetfocus(){
	
} */
/*void CHMSInsuraceRepor21aDialog::OnYearKillfocus(){
	
} */
int CHMSInsuraceRepor21aDialog::OnYearCheckValue(){
	return 0;
} 
void CHMSInsuraceRepor21aDialog::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CHMSInsuraceRepor21aDialog::OnReportPeriodSelendok(){
	
	CString tmpStr;
	CDate dte;
	UpdateData(true);
	int nMonth = ToInt(m_szReportPeriodKey);
	if(nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/01"), m_nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d"), m_nYear, nMonth, dte.GetMonthLastDay());
	}
	if(nMonth == 13){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/03/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d"), m_nYear,  dte.GetMonthLastDay());
	}
	if(nMonth == 14){
		m_szFromDate.Format(_T("%.4d/04/01"), m_nYear);
		tmpStr.Format(_T("%.4d/06/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d"), m_nYear, dte.GetMonthLastDay());
	}

	if(nMonth == 15){
		m_szFromDate.Format(_T("%.4d/07/01"), m_nYear);
		tmpStr.Format(_T("%.4d/09/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 16){
		m_szFromDate.Format(_T("%.4d/10/01"), m_nYear);
		tmpStr.Format(_T("%.4d/10/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 17){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/12/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d"), m_nYear, dte.GetMonthLastDay());		
	}	
	UpdateData(false);
}
/*void CHMSInsuraceRepor21aDialog::OnReportPeriodSetfocus(){
	
}*/
/*void CHMSInsuraceRepor21aDialog::OnReportPeriodKillfocus(){
	
}*/
long CHMSInsuraceRepor21aDialog::OnReportPeriodLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && !m_szReportPeriodKey.IsEmpty()){
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
/*void CHMSInsuraceRepor21aDialog::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CHMSInsuraceRepor21aDialog::OnFromDateChange(){
	
} */
/*void CHMSInsuraceRepor21aDialog::OnFromDateSetfocus(){
	
} */
/*void CHMSInsuraceRepor21aDialog::OnFromDateKillfocus(){
	
} */
int CHMSInsuraceRepor21aDialog::OnFromDateCheckValue(){
	return 0;
} 
/*void CHMSInsuraceRepor21aDialog::OnToDateChange(){
	
} */
/*void CHMSInsuraceRepor21aDialog::OnToDateSetfocus(){
	
} */
/*void CHMSInsuraceRepor21aDialog::OnToDateKillfocus(){
	
} */
int CHMSInsuraceRepor21aDialog::OnToDateCheckValue(){
	return 0;
} 
void CHMSInsuraceRepor21aDialog::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CHMSInsuraceRepor21aDialog::OnClerkSelendok(){
	 
}
/*void CHMSInsuraceRepor21aDialog::OnClerkSetfocus(){
	
}*/
/*void CHMSInsuraceRepor21aDialog::OnClerkKillfocus(){
	
}*/
long CHMSInsuraceRepor21aDialog::OnClerkLoadData(){
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
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CHMSInsuraceRepor21aDialog::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CHMSInsuraceRepor21aDialog::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_bPreview=true;
	PrintTonghopsudungdichvukythuatBHYT();
} 
void CHMSInsuraceRepor21aDialog::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();	
	m_bPreview=false;
	PrintTonghopsudungdichvukythuatBHYT();
} 
void CHMSInsuraceRepor21aDialog::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	ExportTonghopsudungdichvukythuatBHYT();
} 
void CHMSInsuraceRepor21aDialog::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
	
} 
	void CHMSInsuraceRepor21aDialog::OnOutPatientSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
	void CHMSInsuraceRepor21aDialog::OnInPatientSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
int CHMSInsuraceRepor21aDialog::OnAddHMSInsuraceRepor21aDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CHMSInsuraceRepor21aDialog::OnEditHMSInsuraceRepor21aDialog(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CHMSInsuraceRepor21aDialog::OnDeleteHMSInsuraceRepor21aDialog(){
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
 		OnCancelHMSInsuraceRepor21aDialog(); 
 	}
	return 0;
}
int CHMSInsuraceRepor21aDialog::OnSaveHMSInsuraceRepor21aDialog(){
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
 		//OnHMSInsuraceRepor21aDialogListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CHMSInsuraceRepor21aDialog::OnCancelHMSInsuraceRepor21aDialog(){
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
int CHMSInsuraceRepor21aDialog::OnHMSInsuraceRepor21aDialogListLoadData(){
	return 0;
}
//Bao cao tong hop su dung dich vu ky thuat theo quu ( mau 21/BHYT)
void CHMSInsuraceRepor21aDialog::PrintTonghopsudungdichvukythuatBHYT()
{	
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString tmpStr, szSQL, szDate, szMoney;
	if(!rpt.Init(_T("Reports/HMS/HF_THONGKETONGHOPDICHVUKYTHUATTHEOQUY.RPT")) )	
	{
		return;
	}	
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ObjectGroup"), _T(""));
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd,ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	//Page Header
	//Report Detail
	int nIndex = 1;
	CString szOldLine, szNewLine,szAmount;
	CReportSection* rptDetail;
	long double grpCost=0, ttlCost=0, totalCost=0;
	double cost=0;
	//nID = 1-PhiKB
	//nID = 2-Xet nghiem
	//nID = 3-CDHA
	//nID = 4-TDCN
	//nID = 5-PTTT
	//nID = 6-VTTH
	//nID = 7-Ky thuat cao
	CStringArray arrFeeGrp, arrTitle;
	arrFeeGrp.Add(_T("Test"));arrTitle.Add(_T("A"));
	arrFeeGrp.Add(_T("Diagnostic image"));arrTitle.Add(_T("B"));
	arrFeeGrp.Add(_T("Visiting the function rolls"));arrTitle.Add(_T("C"));
	arrFeeGrp.Add(_T("Surgery - procedures"));arrTitle.Add(_T("D"));
	arrFeeGrp.Add(_T("Consumables"));arrTitle.Add(_T("E"));
	arrFeeGrp.Add(_T("High-tech services"));arrTitle.Add(_T("F"));
	
	//if (m_bInPatient != 1)
	//{
	//	/*Phi kham benh
	//-----------------------------------------------------------------------------------------------------------------------------*/
	//	szSQL.Format(_T(" select hfl_name as examname,hfl_unit as unit, ") \
	//	_T(" 	sum(soluong) as soluong, ") \
	//	_T(" 	insprice,sum(soluong * insprice) as suminsprice") \
	//	_T(" from (") \
	//	_T(" 	select he_docno,he_examtype,hfe_invoiceno as invoiceno, ") \
	//	_T(" 		sum(1) as soluong,hfe_insprice as insprice") \
	//	_T(" 	from hms_exam	") \
	//	_T(" 	left join hms_doc on(hd_docno=he_docno) ") \
	//	_T(" 	where hd_object in('1','2') and he_payment ='P'") \
	//	_T(" 	group by he_docno,he_examtype,insprice, invoiceno") \
	//	_T(" ) as tb1 ") \
	//	_T(" left join hms_fee_invoice on(hfi_docno = he_docno and hfi_type ='P' and invoiceno = hfi_invoiceno) ") \
	//	_T(" left join hms_fee_list on (hfl_feeid=he_examtype)") \
	//	_T(" where insprice > 0 and date(hfi_recvdate) between '%s' and '%s'  and hfl_inspaid ='Y'") \
	//	_T(" group by examname,unit, insprice,examname,hfl_unit ") \
	//	_T(" order by examname;"),m_szFromDate,m_szToDate);
	//	
	//	rs.ExecSQL(szSQL);
	//	if (!rs.IsEOF())
	//	{
	//		grpCost=ttlCost=cost=ttlCost=0;
	//		nIndex=1;
	//		rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(12);
	//		rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
	//		rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(false);
	//		rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetFaceSize(12);
	//		rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetBold(true);
	//		rpt.GetGroupHeader(0)->GetItem(_T("Group"))->SetItalic(false);
	//		rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
	//		TranslateString(_T("Nh\xF3m ph\xED kh\xE1m \x62\x1EC7nh"), tmpStr);
	//		rptDetail->SetValue(_T("GroupName"),tmpStr );
	//		rptDetail->SetValue(_T("Group"),_T("E") );		
	//		
	//		while(!rs.IsEOF())
	//		{			
	//			rptDetail = rpt.AddDetail();
	//			tmpStr.Format(_T("%d"), nIndex++);
	//			rptDetail->SetValue(_T("1"), tmpStr);
	//			rs.GetValue(_T("examname"), tmpStr);
	//			rptDetail->SetValue(_T("2"), tmpStr);		
	//			rs.GetValue(_T("soluong"), tmpStr);					
	//			rptDetail->SetValue(_T("3"), tmpStr);
	//			rs.GetValue(_T("insprice"), cost);		
	//			FormatCurrency(cost, tmpStr);
	//			rptDetail->SetValue(_T("4"), tmpStr);		
	//			rs.GetValue(_T("suminsprice"), cost);
	//			grpCost += cost;
	//			FormatCurrency(cost, tmpStr);
	//			rptDetail->SetValue(_T("5"), tmpStr);
	//			
	//			rs.MoveNext();
	//		}
	//		ttlCost += grpCost;
	//		
	//		
	//		if(ttlCost > 0){
	//			CString szField, szAmount;
	//			TranslateString(_T("Total Amount"), szAmount);
	//			rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	//			rptDetail->SetValue(_T("TotalGroup"), szAmount + _T("(E)"));						
	//			FormatCurrency(ttlCost, tmpStr);
	//			rptDetail->SetValue(_T("s5"), tmpStr);
	//			totalCost+=ttlCost;
	//		}
	//	}

	//}		
	grpCost=ttlCost=cost=ttlCost=0;
	nIndex=1;
	int j = 0, nCount = 0, nTemp = 0;
	BeginWaitCursor();
	for (int i = 0; i < arrFeeGrp.GetCount(); i++)
	{
		grpCost=ttlCost=cost=ttlCost=0;
		nIndex=1;
		j = i+2;
		szSQL = GetQueryString(j);
		nCount = rs.ExecSQL(szSQL);
		if(!rs.IsEOF())
		{	
			rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(12);
			rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
			rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(false);
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			TranslateString(arrFeeGrp.GetAt(i), tmpStr);
			rptDetail->SetValue(_T("GroupName"), tmpStr);
			rptDetail->SetValue(_T("Group"), arrTitle.GetAt(i));		
			
			while(!rs.IsEOF())
			{
				rs.GetValue(_T("groupname"), szNewLine);
				if(szNewLine != szOldLine && !szNewLine.IsEmpty()){
					if(grpCost > 0)
					{				
						TranslateString(_T("Total Child Group"), tmpStr);
						rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
						rptDetail->SetValue(_T("TotalGroup"), tmpStr);
						FormatCurrency(grpCost, tmpStr);
						rptDetail->SetValue(_T("s5"), tmpStr);
						ttlCost += grpCost;
						nIndex=1;
						grpCost = 0;				
					}
					rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetFaceSize(10);
					rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetBold(true);
					rpt.GetGroupHeader(0)->GetItem(_T("GroupName"))->SetItalic(true);
					rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));				
					rptDetail->SetValue(_T("GroupName"),szNewLine );
					szOldLine = szNewLine;
				}

				rptDetail = rpt.AddDetail();
				tmpStr.Format(_T("%d"), nIndex++);
				rptDetail->SetValue(_T("1"), tmpStr);
				rs.GetValue(_T("pcmsname"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);		
				rs.GetValue(_T("soluong"), tmpStr);					
				rptDetail->SetValue(_T("3"), tmpStr);
				rs.GetValue(_T("insprice"), cost);		
				FormatCurrency(cost, tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);		
				rs.GetValue(_T("suminsprice"), cost);
				grpCost += cost;
				FormatCurrency(cost, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				
				rs.MoveNext();
			}
			
			
			if(grpCost > 0){	
				TranslateString(_T("Total Child Group"), tmpStr);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));	
				rptDetail->SetValue(_T("TotalGroup"), tmpStr);
				FormatCurrency(grpCost, tmpStr);
				rptDetail->SetValue(_T("s5"), tmpStr);	
				ttlCost += grpCost;		
			}

			if(ttlCost > 0){
				TranslateString(_T("Total Parent Group"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("TotalGroup"), szAmount + _T("(") + arrTitle.GetAt(i) + _T(")"));			
				FormatCurrency(ttlCost, tmpStr);
				rptDetail->SetValue(_T("S5"), tmpStr);	
				totalCost+=ttlCost;
			}
			nTemp += nCount;
		}
	}
	if (nTemp == 0)
	{
		ShowMessageBox(_T("No Data."), MB_ICONSTOP);
		return;
	}
	if(totalCost > 0){
		TranslateString(_T("Total Amount"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), szAmount + _T("(A+B+C+D+E+F)"));						
		FormatCurrency(totalCost, tmpStr);
		rptDetail->SetValue(_T("s5"), tmpStr);
	}
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);			
	MoneyToString(szMoney, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	EndWaitCursor();
	rpt.PrintPreview();	
}

//Bao cao tong hop su dung dich vu ky thuat theo quu ( mau 21/BHYT)
void CHMSInsuraceRepor21aDialog::ExportTonghopsudungdichvukythuatBHYT()
{	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szOldLine, szNewLine;
	int nCol = 0, nRow = 0, j = 0, nIdx = 0, nCount = 0, nTemp = 0;
	double nAmount = 0, nGrp = 0, nTotalGrp = 0, nTotal = 0;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	//xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 30);
	//xls.SetColumnWidth(3, 10);
	//xls.SetColumnWidth(4, 15);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 5);
	xls.SetCellMergedColumns(nCol, nRow + 3, 5);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, _T("\x42\x1EA2NG TH\x1ED0NG K\xCA \x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N N\x1ED8P TI\x1EC0N"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol, arrFeeGrp, arrTitle;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn \x64\x1ECB\x63h v\x1EE5 k\x1EF9 thu\x1EADt"));
	arrCol.Add(_T("S\x1ED1 l\x1B0\x1EE3ng"));
	arrCol.Add(_T("\x110\x1A1n gi\xE1"));
	arrCol.Add(_T("Th\xE0nh ti\x1EC1n"));
	nRow = 4;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10); 
	}
	nRow = 5;
	arrFeeGrp.Add(_T("Test"));arrTitle.Add(_T("A"));
	arrFeeGrp.Add(_T("Diagnostic image"));arrTitle.Add(_T("B"));
	arrFeeGrp.Add(_T("Visiting the function rolls"));arrTitle.Add(_T("C"));
	arrFeeGrp.Add(_T("Surgery - procedures"));arrTitle.Add(_T("D"));
	arrFeeGrp.Add(_T("Consumables"));arrTitle.Add(_T("E"));
	arrFeeGrp.Add(_T("High-tech services"));arrTitle.Add(_T("F"));
	for (int i = 0; i < arrFeeGrp.GetCount(); i++)
	{
		nIdx = 0;
		nAmount = nGrp = nTotalGrp = 0;
		j = i + 2;
		szSQL = GetQueryString(j);
		rs.ExecSQL(szSQL);
		nCount = rs.GetRecordCount();
		if (!rs.IsEOF())
		{
			xls.SetCellMergedColumns(nCol + 1, nRow, 4);
			xls.SetCellText(nCol, nRow, arrTitle.GetAt(i), FMT_TEXT | FMT_CENTER, true, 12);
			TranslateString(arrFeeGrp.GetAt(i), tmpStr);
			xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT, true, 12);
			nRow++;
			while (!rs.IsEOF())
			{
				rs.GetValue(_T("groupname"), szNewLine);
				if (szNewLine != szOldLine)
				{
					if (nGrp > 0)
					{
						xls.SetCellMergedColumns(nCol, nRow, 4);
						TranslateString(_T("Total Child Group"), tmpStr);
						xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);
						FormatCurrency(nGrp, tmpStr);
						xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 10);
						nTotalGrp += nGrp;
						nRow++;
					}
					xls.SetCellMergedColumns(nCol + 1, nRow, 4);
					xls.SetCellText(nCol + 1, nRow, szNewLine, FMT_TEXT, true, 10);
					szOldLine = szNewLine;
					nIdx = 0;
					nGrp = 0;
					nRow++;	
				}
				nIdx++;
				tmpStr.Format(_T("%d"), nIdx);
				xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
				rs.GetValue(_T("pcmsname"), tmpStr);
				xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT);
				rs.GetValue(_T("soluong"), tmpStr);
				xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER);
				nAmount = 0;
				rs.GetValue(_T("insprice"), nAmount);
				FormatCurrency(nAmount, tmpStr);
				xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
				nAmount = 0;
				rs.GetValue(_T("suminsprice"), nAmount);
				FormatCurrency(nAmount, tmpStr);
				xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
				nGrp += nAmount;
				nRow++;
				rs.MoveNext();
			}
			if (nGrp > 0)
			{
				xls.SetCellMergedColumns(nCol, nRow, 4);
				TranslateString(_T("Total Child Group"), tmpStr);
				xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);
				FormatCurrency(nGrp, tmpStr);
				xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 10);
				nTotalGrp += nGrp;
				nRow++;
			}
			if (nTotalGrp > 0)
			{
				xls.SetCellMergedColumns(nCol, nRow, 4);
				TranslateString(_T("Total Parent Group"), tmpStr);
				xls.SetCellText(nCol, nRow, tmpStr + _T("(") + arrTitle.GetAt(i) + _T(")"), FMT_TEXT | FMT_CENTER, true, 10);
				FormatCurrency(nTotalGrp, tmpStr);
				xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true);
				nTotal += nTotalGrp;
				nRow++;				
			}
			nTemp += nCount;	
		}
	}
	if (nTemp == 0)
	{
		ShowMessageBox(_T("No Data."), MB_ICONSTOP);
		return;
	}
	if (nTotal > 0)
	{
		xls.SetCellMergedColumns(nCol, nRow, 4);
		TranslateString(_T("Total Amount"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr + _T("(A+B+C+D+E+F)"), FMT_TEXT | FMT_CENTER, true, 10);
		FormatCurrency(nTotal, tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true);
		nRow++;		
	}
	xls.Save(_T("Exports\\Thong Ke Dich Vu Ky Thuat Theo Quy.xls"));
}

CString CHMSInsuraceRepor21aDialog::GetQueryString(int nID){
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString szSQL, szWhere, szDeptidCLS, szDeptOperation, szDeptDrug;
	szWhere.Format(_T(" AND trunc_date(hfe_date) BETWEEN cast_string2date('%s') AND cast_string2date('%s') AND hfe_type = 'I'"), m_szFromDate, m_szToDate);
	if (m_bCheckDrug)
		m_szDrugPTTT.Format(_T(" AND hpo_type IN ('P', 'D', 'E')"));
	else
		m_szDrugPTTT.Format(_T(" AND hpo_type IN ('P', 'D', 'M', 'E')"));
	if (m_bInPatient==1 && m_bOutPatient==1)	
	{
		szDeptidCLS.Empty();
		szDeptOperation.Empty();
		szDeptDrug.Empty();
	}
	else if (m_bInPatient==1)
	{
		szDeptidCLS.Format(_T(" AND hpc_depttype ='I' "));
		szDeptOperation.Format(_T(" AND ho_depttype='I' "));
		szDeptDrug.Format(_T(" AND hpo_depttype='I' "));
	}
	else
	{
		szDeptidCLS.Format(_T(" AND hpc_depttype ='E' "));	
		szDeptOperation.Format(_T(" AND ho_depttype='E' "));
		szDeptDrug.Format(_T(" AND hpo_depttype='E' "));
	}
	m_szDrugPTTT.Format(_T(" AND hpo_type in('P','D','M','E') "));
	if(m_bCheckDrug ==1)
	{
		m_szDrugPTTT.Format(_T(" AND hpo_type in('P','D','E') "));
	}
	//nID = 1-PhiKB
	//nID = 2-Xet nghiem
	//nID = 3-CDHA
	//nID = 4-TDCN
	//nID = 5-PTTT
	//nID = 6-VTTH
	//nID = 7-Ky thuat cao
	switch (nID){
	case 1:
		//szSQL.Format();
		break;
	case 2:
		szSQL.Format(_T(" SELECT") \
					_T("   (SELECT distinct hfg_name FROM hms_fee_group WHERE hfg_id=groupid") \
					_T("   ) AS groupname,") \
					_T("   groupid,") \
					_T("   pcmsname,") \
					_T("   unit,") \
					_T("   SUM(soluong) AS soluong,") \
					_T("   servprice,") \
					_T("   insprice,") \
					_T("   SUM(soluong * servprice) AS sumservprice,") \
					_T("   SUM(soluong * insprice)  AS suminsprice") \
					_T(" FROM") \
					_T("   (SELECT hpc_docno,") \
					_T("     hfl_groupid   AS groupid,") \
					_T("     hfl_name      AS pcmsname,") \
					_T("     hfl_unit      AS unit,") \
					_T("     SUM(1)        AS soluong,") \
					_T("     hfl_servprice AS servprice,") \
					_T("     hfl_insprice  AS insprice") \
					_T("   FROM hms_xtest") \
					_T("   LEFT JOIN hms_xtest_line ON(hpc_orderid=hpcl_orderid)") \
					_T("   LEFT JOIN hms_fee_list ON (hfl_groupid=hpc_groupid AND hfl_feeid  =hpcl_itemid)") \
					_T("   LEFT JOIN hms_doc ON(hd_docno                =hpc_docno)") \
					_T("   WHERE hpc_status        <> 'O'") \
					_T("   AND hfl_inspaid            ='Y'") \
					_T("   AND SUBSTR(hfl_groupid,1,2)='B1'") \
					_T("   AND hd_object             IN('2','4','5','9','10')") \
					_T("   AND hfl_hitech             ='N'") \
					_T("   AND hpc_payment          ='P'") \
					_T("   GROUP BY hpc_docno,") \
					_T("     hfl_groupid,") \
					_T("     hfl_name,") \
					_T("     hfl_unit,") \
					_T("     hfl_servprice,") \
					_T("     hfl_insprice") \
					_T("   ) tb1") \
					_T(" LEFT JOIN hms_xfee_invoice ON(hfe_docno      = hpc_docno)") \
					_T(" WHERE servprice   > 0 %s") \
					_T(" GROUP BY groupid,") \
					_T("   pcmsname,") \
					_T("   unit,") \
					_T("   servprice,") \
					_T("   insprice") \
					_T(" ORDER BY groupid"), szWhere);
		break;
	case 3:
		szSQL.Format(_T(" SELECT") \
					_T("   (SELECT distinct hfg_name FROM hms_fee_group WHERE hfg_id=groupid") \
					_T("   ) AS groupName,") \
					_T("   groupid,") \
					_T("   pcmsname,") \
					_T("   unit,") \
					_T("   SUM(soluong) AS soluong,") \
					_T("   servprice,") \
					_T("   insprice,") \
					_T("   SUM(soluong * servprice) AS sumservprice,") \
					_T("   SUM(soluong * insprice)  AS suminsprice") \
					_T(" FROM") \
					_T("   (SELECT hpc_docno,") \
					_T("     hfl_groupid   AS groupid,") \
					_T("     hfl_name      AS pcmsname,") \
					_T("     hfl_unit      AS unit,") \
					_T("     SUM(1)        AS soluong,") \
					_T("     hfl_servprice AS servprice,") \
					_T("     hfl_insprice  AS insprice") \
					_T("   FROM hms_xpacs") \
					_T("   LEFT JOIN hms_xpacs_line ON(hpc_orderid=hpcl_orderid)") \
					_T("   LEFT JOIN hms_fee_list ON (hfl_groupid=hpc_groupid AND hfl_feeid  =hpcl_itemid)") \
					_T("   LEFT JOIN hms_doc ON(hd_docno = hpc_docno)") \
					_T("   WHERE hpc_status        <> 'O'") \
					_T("   AND hfl_inspaid            ='Y'") \
					_T("   AND SUBSTR(hfl_groupid,1,2)='B2'") \
					_T("   AND hd_object             IN('2','4','5','9','10')") \
					_T("   AND hfl_hitech             ='N'") \
					_T("   AND hpc_payment          ='P'") \
					_T("   GROUP BY hpc_docno,") \
					_T("     hfl_groupid,") \
					_T("     hfl_name,") \
					_T("     hfl_unit,") \
					_T("     hfl_servprice,") \
					_T("     hfl_insprice") \
					_T("   ) tb1") \
					_T(" LEFT JOIN hms_xfee_invoice ON(hfe_docno = hpc_docno)") \
					_T(" WHERE servprice   > 0 %s") \
					_T(" GROUP BY groupid,") \
					_T("   pcmsname,") \
					_T("   unit,") \
					_T("   servprice,") \
					_T("   insprice") \
					_T(" ORDER BY groupid"), szWhere);
		break;
	case 4:
		szSQL.Format(_T(" SELECT") \
					_T("   (SELECT distinct hfg_name FROM hms_fee_group WHERE hfg_id=groupid") \
					_T("   ) AS groupName,") \
					_T("   groupid,") \
					_T("   pcmsname,") \
					_T("   unit,") \
					_T("   SUM(soluong) AS soluong,") \
					_T("   servprice,") \
					_T("   insprice,") \
					_T("   SUM(soluong * servprice) AS sumservprice,") \
					_T("   SUM(soluong * insprice)  AS suminsprice") \
					_T(" FROM") \
					_T("   (SELECT hpc_docno,") \
					_T("     hfl_groupid   AS groupid,") \
					_T("     hfl_name      AS pcmsname,") \
					_T("     hfl_unit      AS unit,") \
					_T("     SUM(1)        AS soluong,") \
					_T("     hfl_servprice AS servprice,") \
					_T("     hfl_insprice  AS insprice") \
					_T("   FROM hms_xpacs") \
					_T("   LEFT JOIN hms_xpacs_line ON(hpc_orderid=hpcl_orderid)") \
					_T("   LEFT JOIN hms_fee_list ON (hfl_groupid=hpc_groupid AND hfl_feeid  =hpcl_itemid)") \
					_T("   LEFT JOIN hms_doc ON(hd_docno                =hpc_docno)") \
					_T("   WHERE hpc_status        <> 'O'") \
					_T("   AND hfl_inspaid            ='Y'") \
					_T("   AND SUBSTR(hfl_groupid,1,2)='B3'") \
					_T("   AND hd_object             IN('2','4','5','9','10')") \
					_T("   AND hfl_hitech             ='N'") \
					_T("   AND hpc_payment          = 'P'") \
					_T("   GROUP BY hpc_docno,") \
					_T("     hfl_groupid,") \
					_T("     hfl_name,") \
					_T("     hfl_unit,") \
					_T("     hfl_servprice,") \
					_T("     hfl_insprice") \
					_T("   ) tb1") \
					_T(" LEFT JOIN hms_xfee_invoice ON(hfe_docno = hpc_docno)") \
					_T(" WHERE servprice   > 0 %s") \
					_T(" GROUP BY groupid,") \
					_T("   pcmsname,") \
					_T("   unit,") \
					_T("   servprice,") \
					_T("   insprice") \
					_T(" ORDER BY groupid"), szWhere);
		break;
	case 5:
		szWhere.Format(_T(" AND trunc_date(xi.hfe_date) BETWEEN cast_string2date('%s') AND cast_string2date('%s')"), m_szFromDate, m_szToDate);
		szSQL.Format(_T(" SELECT hfl_name AS pcmsname,") \
					_T("   (SELECT DISTINCT hfg_name FROM hms_fee_group WHERE hfg_id=hfl_groupid") \
					_T("   )                AS groupname,") \
					_T("   hfe_insprice     AS insprice,") \
					_T("   SUM(ho_qty)      AS soluong,") \
					_T("   SUM(xf.hfe_inspaid) AS suminsprice") \
					_T(" FROM hms_doc") \
					_T(" LEFT JOIN hms_operation ON (hd_docno = ho_docno)") \
					_T(" LEFT JOIN hms_xfee xf ON (hfe_rowid = hfe_refrowid)") \
					_T(" LEFT JOIN hms_fee_list ON (hfl_feeid=ho_itemid)") \
					_T(" LEFT JOIN hms_xfee_invoice xi ON(xi.hfe_docno = hd_docno)") \
					_T(" WHERE hd_object  IN('2','4','5','9','10')") \
					_T(" AND ho_payment  ='P'") \
					_T(" AND hfl_hitech  ='N'") \
					_T(" AND hfl_inspaid ='Y'") \
					_T(" AND trunc_date(xi.hfe_date) BETWEEN cast_string2date('%s') AND cast_string2date('%s') AND xi.hfe_type = 'I'") \
					_T(" GROUP BY hfl_groupid,") \
					_T("   hfl_name,") \
					_T("   hfe_insprice") \
					_T(" ORDER BY hfl_groupid"), m_szFromDate, m_szToDate);
		break;
	case 6:
		szSQL.Format(_T(" SELECT") \
					_T("   (SELECT distinct hfg_name AS name FROM hms_fee_group WHERE hfg_id= pmi_typeid") \
					_T("   )        AS groupname,") \
					_T("   pmi_id   AS id,") \
					_T("   pmi_name AS pcmsname,") \
					_T("   pmi_unit AS unit,") \
					_T("   (SELECT distinct pmm_name FROM pms_mfg WHERE pmm_id=pmsi_mfgid") \
					_T("   )                                AS countryid,") \
					_T("   pmsi_vatprice                    AS insprice,") \
					_T("   SUM(hpol_issueqty)               AS soluong,") \
					_T("   SUM(hpol_issueqty*pmsi_vatprice) AS suminsprice") \
					_T(" FROM hms_pharma_xorder") \
					_T(" LEFT JOIN hms_pharma_xorder_line ON (hpol_orderid=hpo_orderid)") \
					_T(" LEFT JOIN hms_doc ON(hd_docno=hpo_docno)") \
					_T(" LEFT JOIN hms_xfee_invoice ON(hfe_docno = hpo_docno)") \
					_T(" LEFT JOIN pms_stockitems ON(pmsi_id=hpol_sitemid)") \
					_T(" LEFT JOIN pms_items ON(pmsi_itemid=pmi_id)") \
					_T(" LEFT JOIN hms_object ON(hd_object      =ho_id)") \
					_T(" WHERE pmi_inspaid ='Y'") \
					_T(" AND ho_type    IN('I','C')") \
					_T(" AND pmi_typeid  ='A9000'") \
					_T(" AND hpo_type = 'P'") \
					_T(" AND hpo_status IN('A','I','P')") \
					_T(" AND hpo_payment ='P' %s %s") \
					_T(" GROUP BY pmi_typeid,") \
					_T("   pmi_id,") \
					_T("   pmi_name ,") \
					_T("   pmi_unit ,") \
					_T("   pmsi_vatprice,") \
					_T("   pmsi_mfgid") \
					_T(" ORDER BY pmi_typeid"), szWhere, m_szDrugPTTT);
		break;
	case 7:
		szSQL.Format(_T(" SELECT") \
					_T("   (SELECT distinct hfg_name FROM hms_fee_group WHERE hfg_id=hfl_groupid) AS groupname,") \
					_T("   hfl_groupid  AS groupid,") \
					_T("   hfl_name     AS pcmsname,") \
					_T("   hfl_unit     AS unit,") \
					_T("   SUM(soluong) AS soluong,") \
					_T("   insprice,") \
					_T("   SUM(soluong * insprice) AS suminsprice") \
					_T(" FROM") \
					_T("   (SELECT hfe_invoiceno as invoiceno,") \
					_T("     pcmso_docno    AS docno,") \
					_T("     pcmsol_itemid  AS itemid,") \
					_T("     pcmso_payment  AS payment,") \
					_T("     pcmso_depttype AS depttype,") \
					_T("     1              AS soluong,") \
					_T("     hfe_insprice   AS insprice") \
					_T("   FROM pcms_order") \
					_T("   LEFT JOIN pcms_order_line ON(pcmso_orderid   =pcmsol_orderid)") \
					_T("   WHERE 1 =1") \
					_T("   ") \
					_T("   UNION ALL") \
					_T("   SELECT hfe_invoiceno,") \
					_T("     ho_docno,") \
					_T("     ho_itemid,") \
					_T("     ho_payment,") \
					_T("     ho_depttype,") \
					_T("     ho_qty       AS qty,") \
					_T("     hfe_insprice AS price") \
					_T("   FROM hms_operation") \
					_T("   LEFT JOIN hms_xfee ON (hfe_rowid = hfe_refrowid)") \
					_T("   WHERE 1 =1") \
					_T("   %s") \
					_T("   ) tb1") \
					_T(" LEFT JOIN hms_xfee_invoice ON(hfe_docno = docno AND hfe_invoiceno=invoiceno) ") \
					_T(" LEFT JOIN hms_doc ON (hd_docno = docno) ") \
					_T(" LEFT JOIN hms_fee_list ON (hfl_feeid = itemid)") \
					_T(" WHERE hd_object IN('2','4','5','9','10')") \
					_T(" AND hfl_hitech   ='Y'") \
					_T(" AND hfl_inspaid  ='Y'") \
					_T(" AND payment      ='P'") \
					_T(" AND insprice     > 0 %s") \
					_T(" GROUP BY   hfl_groupid,") \
					_T("   hfl_name,") \
					_T("   hfl_unit,") \
					_T("   insprice") \
					_T(" ORDER BY groupid"), szDeptOperation, szWhere);
		break;
	}
	return szSQL;
}