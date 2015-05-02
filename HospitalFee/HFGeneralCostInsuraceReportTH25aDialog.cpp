#include "stdafx.h"
#include "HFGeneralCostInsuraceReportTH25aDialog.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CHMSGeneralCostInsuraceReportTH25aDialog* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CHMSGeneralCostInsuraceReportTH25aDialog* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CHMSGeneralCostInsuraceReportTH25aDialog *pVw = (CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CHMSGeneralCostInsuraceReportTH25aDialog *pVw = (CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CHMSGeneralCostInsuraceReportTH25aDialog *pVw = (CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CHMSGeneralCostInsuraceReportTH25aDialog *pVw = (CHMSGeneralCostInsuraceReportTH25aDialog *)pWnd;
	pVw->OnCloseSelect();
} 
static void _OnKhonglaybnccSelectFnc(CWnd *pWnd){
	 ((CHMSGeneralCostInsuraceReportTH25aDialog*)pWnd)->OnKhonglaybnccSelect();
}
static void _OnLaybnccSelectFnc(CWnd *pWnd){
	 ((CHMSGeneralCostInsuraceReportTH25aDialog*)pWnd)->OnLaybnccSelect();
}
static int _OnAddHMSGeneralCostInsuraceReportTH25aDialogFnc(CWnd *pWnd){
	 return ((CHMSGeneralCostInsuraceReportTH25aDialog*)pWnd)->OnAddHMSGeneralCostInsuraceReportTH25aDialog();
} 
static int _OnEditHMSGeneralCostInsuraceReportTH25aDialogFnc(CWnd *pWnd){
	 return ((CHMSGeneralCostInsuraceReportTH25aDialog*)pWnd)->OnEditHMSGeneralCostInsuraceReportTH25aDialog();
} 
static int _OnDeleteHMSGeneralCostInsuraceReportTH25aDialogFnc(CWnd *pWnd){
	 return ((CHMSGeneralCostInsuraceReportTH25aDialog*)pWnd)->OnDeleteHMSGeneralCostInsuraceReportTH25aDialog();
} 
static int _OnSaveHMSGeneralCostInsuraceReportTH25aDialogFnc(CWnd *pWnd){
	 return ((CHMSGeneralCostInsuraceReportTH25aDialog*)pWnd)->OnSaveHMSGeneralCostInsuraceReportTH25aDialog();
} 
static int _OnCancelHMSGeneralCostInsuraceReportTH25aDialogFnc(CWnd *pWnd){
	 return ((CHMSGeneralCostInsuraceReportTH25aDialog*)pWnd)->OnCancelHMSGeneralCostInsuraceReportTH25aDialog();
} 
CHMSGeneralCostInsuraceReportTH25aDialog::CHMSGeneralCostInsuraceReportTH25aDialog(CWnd *pParent)
	{

	m_nDlgWidth = 355;
	m_nDlgHeight = 215;
	m_szTitle = _T("General Examination Cost 25a Report(25a-TH/BHYT)");
	SetDefaultValues();
}
CHMSGeneralCostInsuraceReportTH25aDialog::~CHMSGeneralCostInsuraceReportTH25aDialog()
{
	if (m_rs)
		delete m_rs;
}
void CHMSGeneralCostInsuraceReportTH25aDialog::OnCreateComponents(){	
	m_wndReportFilter.Create(this, _T("Report Parameters"), 5, 5, 350, 180);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 175, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Month"), 180, 30, 260, 55);
	m_wndReportPeriod.Create(this,265, 30, 345, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 175, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 180, 60, 260, 85);
	m_wndToDate.Create(this,265, 60, 345, 85); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 90, 115);
	m_wndClerk.Create(this,95, 90, 345, 115); 
	m_wndKhonglaybncc.Create(this, _T("Khong lay benh nhan cap cuu vao bao cao"), 10, 120, 415, 145);
	m_wndLaybncc.Create(this, _T("Chi lay benh nhan cap cuu vao bao cao"), 10, 150, 315, 175);
	//m_wndPrintPreview.Create(this, _T("Print Pre&view"), 84, 185, 164, 210);
	m_wndPrint.Create(this, _T("&Print"), 185, 185, 265, 210);
	m_wndExport.Create(this, _T("&Export"), 270, 185, 350, 210);


}
void CHMSGeneralCostInsuraceReportTH25aDialog::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(FALSE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	//m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	//m_wndToDate.SetCheckValue(true);
	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);

	m_rs = new CRecord(&pMF->m_db);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 60);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 75);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

}
void CHMSGeneralCostInsuraceReportTH25aDialog::OnSetWindowEvents(){
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
	m_wndKhonglaybncc.SetEvent(WE_CLICK, _OnKhonglaybnccSelectFnc);
	m_wndLaybncc.SetEvent(WE_CLICK, _OnLaybnccSelectFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL;
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	SetMode(VM_EDIT);
}
void CHMSGeneralCostInsuraceReportTH25aDialog::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndKhonglaybncc.GetDlgCtrlID(), m_bKhonglaybncc);
	DDX_Check(pDX, m_wndLaybncc.GetDlgCtrlID(), m_bLaybncc);

}
void CHMSGeneralCostInsuraceReportTH25aDialog::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();	

}
void CHMSGeneralCostInsuraceReportTH25aDialog::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CHMSGeneralCostInsuraceReportTH25aDialog::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bKhonglaybncc=FALSE;
	m_bLaybncc=FALSE;

}
int CHMSGeneralCostInsuraceReportTH25aDialog::SetMode(int nMode){
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
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnYearChange(){
	
} */
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnYearSetfocus(){
	
} */
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnYearKillfocus(){
	
} */
int CHMSGeneralCostInsuraceReportTH25aDialog::OnYearCheckValue(){
	return 0;
} 
void CHMSGeneralCostInsuraceReportTH25aDialog::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CHMSGeneralCostInsuraceReportTH25aDialog::OnReportPeriodSelendok(){
	
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
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnReportPeriodSetfocus(){
	
}*/
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnReportPeriodKillfocus(){
	
}*/
long CHMSGeneralCostInsuraceReportTH25aDialog::OnReportPeriodLoadData(){
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
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnFromDateChange(){
	
} */
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnFromDateSetfocus(){
	
} */
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnFromDateKillfocus(){
	
} */
int CHMSGeneralCostInsuraceReportTH25aDialog::OnFromDateCheckValue(){
	return 0;
} 
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnToDateChange(){
	
} */
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnToDateSetfocus(){
	
} */
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnToDateKillfocus(){
	
} */
int CHMSGeneralCostInsuraceReportTH25aDialog::OnToDateCheckValue(){
	return 0;
} 
void CHMSGeneralCostInsuraceReportTH25aDialog::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CHMSGeneralCostInsuraceReportTH25aDialog::OnClerkSelendok(){
	 
}
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnClerkSetfocus(){
	
}*/
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnClerkKillfocus(){
	
}*/
long CHMSGeneralCostInsuraceReportTH25aDialog::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty()){
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	};
	m_wndClerk.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as id, su_name as name FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CHMSGeneralCostInsuraceReportTH25aDialog::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CHMSGeneralCostInsuraceReportTH25aDialog::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	PrintGeneralCost25AReport(m_szClerkKey,m_szFromDate,m_szToDate, true);		
	UpdateData(false);
} 
void CHMSGeneralCostInsuraceReportTH25aDialog::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();	
	UpdateData(true);
	PrintGeneralCost25AReport(m_szClerkKey,m_szFromDate,m_szToDate, true);		
	UpdateData(false);
} 
void CHMSGeneralCostInsuraceReportTH25aDialog::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();	
	UpdateData(true);
	ExportToExcell25AReport(m_szClerkKey,m_szFromDate,m_szToDate, true);	
	UpdateData(false);
} 
void CHMSGeneralCostInsuraceReportTH25aDialog::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
	void CHMSGeneralCostInsuraceReportTH25aDialog::OnKhonglaybnccSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	m_bLaybncc=false;
	UpdateData(false);
}
	void CHMSGeneralCostInsuraceReportTH25aDialog::OnLaybnccSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	m_bKhonglaybncc=false;	
	UpdateData(false);
}
int CHMSGeneralCostInsuraceReportTH25aDialog::OnAddHMSGeneralCostInsuraceReportTH25aDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CHMSGeneralCostInsuraceReportTH25aDialog::OnEditHMSGeneralCostInsuraceReportTH25aDialog(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CHMSGeneralCostInsuraceReportTH25aDialog::OnDeleteHMSGeneralCostInsuraceReportTH25aDialog(){
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
 		OnCancelHMSGeneralCostInsuraceReportTH25aDialog(); 
 	}
	return 0;
}
int CHMSGeneralCostInsuraceReportTH25aDialog::OnSaveHMSGeneralCostInsuraceReportTH25aDialog(){
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
 		//OnHMSGeneralCostInsuraceReportTH25aDialogListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CHMSGeneralCostInsuraceReportTH25aDialog::OnCancelHMSGeneralCostInsuraceReportTH25aDialog(){
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
int CHMSGeneralCostInsuraceReportTH25aDialog::OnHMSGeneralCostInsuraceReportTH25aDialogListLoadData(){
	return 0;
}
void CHMSGeneralCostInsuraceReportTH25aDialog::PrintGeneralCost25AReport(CString szReportPeriod, CString szFromDate, CString szToDate, bool bPreview)
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	TCHAR *lszLine[] ={_T("\x110\xFAng tuy\x1EBFn"), _T("Tr\xE1i tuy\x1EBFn")};
	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	if(rs.IsEOF()){
		ShowMessage(150, MB_ICONSTOP);
		return ;
	}
	if(!rpt.Init(_T("Reports/HMS/HF_25aTH.RPT")) )
		return;
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	rpt.GetReportHeader()->SetValue(_T("ObjectGroup"), _T(""));
	rpt.GetReportHeader()->SetValue(_T("ReportPeriod"), szReportPeriod);
	rs.GetValue(_T("invoiceno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(szToDate, yyyymmdd,ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	//Page Header
	//Report Detail
	int nIndex = 1;
	CString szOldLine, szNewLine, szOldGroup, szNewGroup, szLineName;
	CString szNewOfLine, szOldOfLine,szSumOfline;
	CReportSection* rptDetail;
	long double grpCost[23], grpLine[23], ttlCost[23], grpOfLine[23];
	double cost = 0;
	for (int i =0; i < 23; i++)
	{
		grpCost[i] = 0;
		grpLine[i] = 0;
		ttlCost[i] = 0;
		grpOfLine[i] =0;
	}
	while(!rs.IsEOF())
	{		
		rs.GetValue(_T("linename"), szNewGroup);
		if(szNewGroup != szOldGroup && !szNewGroup.IsEmpty()){
			if(grpCost[14] > 0)
			{	
				CString szField;
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(8);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				tmpStr.Format(_T("\x43\x1ED9ng(%s=(I+II)):"), szOldGroup);
				rptDetail->SetValue(_T("TotalGroup"), tmpStr );
				for (int i =3; i < 17; i++)
				{
					szField.Format(_T("S%d"), i);
					FormatCurrency(grpCost[i], tmpStr);
					rptDetail->SetValue(szField, tmpStr);
					ttlCost[i] += grpCost[i];
					grpCost[i] = 0;
				}
			}
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetFaceSize(10);
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetItalic(true);
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetBold(true);			
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rs.GetValue(_T("insline"), tmpStr);
			rptDetail->SetValue(_T("InsuranceLine"),szNewGroup +_T(". ") + pMF->GetSelectionString(_T("hms_insline"), tmpStr));
			szOldGroup = szNewGroup;			
			nIndex=1;
		}

		rs.GetValue(_T("hdline"), szNewOfLine);
		if (szNewOfLine == _T("0") )
		{
			tmpStr.Format(_T("%s"), lszLine[0]);
			szSumOfline = _T("I");
		}
		else
		{
			tmpStr.Format(_T("%s"), lszLine[1]);
			szSumOfline = _T("II");
		}
			
	
		rptDetail = rpt.AddDetail();		
		rptDetail->SetValue(_T("1"), szSumOfline);
		rptDetail->SetValue(_T("2"), tmpStr);
		
		rs.GetValue(_T("total"),cost);
		grpCost[3] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("testfee"), cost);
		grpCost[4] += cost;
		grpLine[4] += cost;		
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("pacsfee"), cost);
		grpCost[5] += cost;
		grpLine[5] += cost;
		grpOfLine[5] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("drugfee"), cost);
		grpCost[6] += cost;
		grpLine[6] += cost;
		grpOfLine[6] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("bloodfee"), cost);
		grpCost[7] += cost;
		grpLine[7] += cost;
		grpOfLine[7] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);				

		rs.GetValue(_T("normtechfee"), cost);
		grpCost[8] += cost;
		grpLine[8] += cost;
		grpOfLine[8] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		
		rs.GetValue(_T("materialfee"), cost);
		grpCost[9] += cost;
		grpLine[9] += cost;
		grpOfLine[9] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("hitechfee"), cost);
		grpCost[10] += cost;
		grpLine[10] += cost;
		grpOfLine[10] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);		

		rs.GetValue(_T("drugfeek"), cost);
		grpCost[11] += cost;
		grpLine[11] += cost;
		grpOfLine[11] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);
		
		rs.GetValue(_T("examfee"), cost);
		grpCost[12] += cost;
		grpLine[12] += cost;
		grpOfLine[12] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);

		rs.GetValue(_T("otherfee"), cost);;
		grpCost[13] += cost;
		grpLine[13] += cost;
		grpOfLine[13] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);		
		
		rs.GetValue(_T("cost"), cost);
		grpCost[14] += cost;
		grpLine[14] += cost;
		grpOfLine[14] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("14"), tmpStr);

		rs.GetValue(_T("patpaid"), cost);
		grpCost[15] += cost;
		grpLine[15] += cost;
		grpOfLine[15] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("15"), tmpStr);

		rs.GetValue(_T("inspaid"), cost);
		grpCost[16] += cost;
		grpLine[16] += cost;
		grpOfLine[16] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("16"), tmpStr);
		
		rs.MoveNext();
	}
	for (int i =3; i < 17; i++)
	{
			ttlCost[i] += grpCost[i];
	}
	
	
	if(grpCost[14] > 0){		
		CString szField, szAmount;
		TranslateString(_T("Total Line"), szAmount);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(8);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		tmpStr.Format(_T("\x43\x1ED9ng(%s=(I+II)):"), szOldGroup);		
		rptDetail->SetValue(_T("TotalGroup"), tmpStr );
		for (int i =3; i < 17; i++)
		{
			szField.Format(_T("S%d"), i);
			FormatCurrency(grpCost[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);
		}
	
	}
	if(ttlCost[14] > 0)
	{
		CString szField, szAmount;
		TranslateString(_T("Total Amount:"), szAmount);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(10);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(false);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		for (int i =3; i < 17; i++)
		{
			
			szField.Format(_T("S%d"), i);
			FormatCurrency(ttlCost[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);		
		}

	}
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	
	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print();

	EndWaitCursor();

}

void CHMSGeneralCostInsuraceReportTH25aDialog::ExportToExcell25AReport(CString szReportPeriod, CString szFromDate, CString szToDate, bool bPreview)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString tmpStr, szSQL;
	TCHAR *lszLine[] ={_T("\x110\xFAng tuy\x1EBFn"), _T("Tr\xE1i tuy\x1EBFn")};

	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);	

	CellFormat df(&xls), cf(&xls), hf(&xls);
	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	df.SetFontSize(11);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING);
	hf.SetFontSize(11);
	cf.SetFontName(_T("Segoe UI"));
	cf.SetFontSize(11);
	xls.SetColumnWidth(0,5);
	xls.SetColumnWidth(1,23);
	xls.SetColumnWidth(2,9);
	xls.SetColumnWidth(3,12);
	xls.SetColumnWidth(4,12);
	xls.SetColumnWidth(5,12);
	xls.SetColumnWidth(6,12);
	xls.SetColumnWidth(7,12);
	xls.SetColumnWidth(8,12);
	xls.SetColumnWidth(9,12);
	xls.SetColumnWidth(10,12);
	xls.SetColumnWidth(11,12);
	xls.SetColumnWidth(12,12);
	xls.SetColumnWidth(13,12);
	xls.SetColumnWidth(14,12);
	xls.SetColumnWidth(15,12);
	xls.SetColumnWidth(16,12);
	xls.SetColumnWidth(17,12);
	
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellText(0, 0, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellMergedColumns(0, 2, 16);
	xls.SetCellText(0, 2, _T("\x44\x41NH S\xC1\x43H \x110\x1EC0 NGH\x1ECA TH\x41NH TO\xC1N \x43HI PH\xCD KH\xC1M \x43H\x1EEE\x41 \x42\x1EC6NH NGO\x1EA0I TR\xDA"), FMT_TEXT | FMT_CENTER, true, 14);
	xls.SetCellMergedColumns(0, 3, 16);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(0, 3, tmpStr, &df);
	int nRow = 4;
	xls.SetCellText(0, nRow, _T("STT"), &hf);	
	TranslateString(_T("KCB Trong k\x1EF3"), tmpStr);
	xls.SetCellText(1, nRow, tmpStr, &hf);
	xls.SetCellText(2, nRow, _T("S\x1ED1 l\x1B0"), &hf);
	
	xls.SetCellMergedRows(0, nRow, 1);
	xls.SetCellMergedRows(1, nRow, 1);
	xls.SetCellMergedRows(2, nRow, 1);	
		
	//xls.SetCellMergedRows(12, nRow, 2);
	xls.SetCellMergedRows(13, nRow, 1);
	xls.SetCellMergedRows(14, nRow, 1);
	xls.SetCellMergedRows(15, nRow, 1);
	xls.SetCellMergedRows(16, nRow, 1);
	
	xls.SetCellMergedColumns(3, nRow, 8);
	nRow = 5;
	TranslateString(_T("Medical expenses incurred at medical care facilities"), tmpStr);
	xls.SetCellText(3, nRow-1, tmpStr, &hf);	
	xls.SetCellText(3, nRow, _T("XN,CDCN"), &hf);	
	xls.SetCellText(4, nRow, _T("CDHA"), &hf);

	TranslateString(_T("Drug"), tmpStr);
	xls.SetCellText(5, nRow, tmpStr, &hf);	
	TranslateString(_T("Blood"), tmpStr);
	xls.SetCellText(6, nRow, tmpStr, &hf);	
	TranslateString(_T("TT, PT"), tmpStr);
	xls.SetCellText(7, nRow, tmpStr, &hf);
	xls.SetCellText(8, nRow, _T("VTYT"), FMT_TEXT|FMT_CENTER,true, 11);	
	xls.SetCellText(9, nRow, _T("DVKTC"), FMT_TEXT|FMT_CENTER,true, 11);	
	xls.SetCellText(10, nRow, _T("Thu\x1ED1\x63 K"), &hf);
	TranslateString(_T("Examination Fee"), tmpStr);
	xls.SetCellText(11, nRow, tmpStr, FMT_TEXT|FMT_CENTER,true, 11);	;	
	xls.SetCellText(12, nRow, _T("Phi VC"), &hf);	
	TranslateString(_T("Total Cost"), tmpStr);
	xls.SetCellText(13, nRow -1, tmpStr, &hf);
	TranslateString(_T("Difference Payment"), tmpStr);
	xls.SetCellText(14, nRow - 1, tmpStr, &hf);	
	TranslateString(_T("Insurance Paid"), tmpStr);	
	xls.SetCellText(15, nRow - 1, tmpStr, &hf);
	TranslateString(_T("Ngo\xE0i qu\x1EF9"), tmpStr);	
	xls.SetCellText(16, nRow - 1, tmpStr, &hf);
	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	if(rs.IsEOF()){
		ShowMessage(150, MB_ICONSTOP);
		return ;
	}
	long nIndex = 1;
	CString szOldLine, szNewLine, szOldGroup, szNewGroup, szLineName;
	CString szNewOfLine, szOldOfLine,szSumOfline;

	double grpCost[17], grpLine[17], ttlCost[7];
	double cost = 0;
	nRow++;
	for (int i =0; i < 17; i++)
	{
		tmpStr.Format(_T("(%d)"), i+1);
		xls.SetCellText(i,nRow,tmpStr,FMT_TEXT|FMT_CENTER,TRUE);		
		grpLine[i] = 0;
		ttlCost[i] = 0;		
		grpCost[i] = 0;
	}	

	while(!rs.IsEOF())
	{		
		rs.GetValue(_T("linename"), szNewLine);
		if(szNewLine != szOldLine && !szNewLine.IsEmpty()){
			CString szField,szAmount;			
			//Tong tuyen (A,B,C)
			if(grpLine[13] > 0)
			{			
				nRow++;				
				xls.SetCellMergedColumns(1,nRow,2);
				tmpStr.Format(_T("\x43\x1ED9ng(%s %s=(I+II)): "),szOldLine , szOldGroup);				
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT,true);
				for (int i = 2; i < 17; i++)
				{					
					tmpStr.Format(_T("%.2f"),grpLine[i]);
					xls.SetCellText(i, nRow , tmpStr, FMT_NUMBER1,true);
					ttlCost[i] += grpLine[i];
					grpLine[i] = 0;
				}
			}		

			nRow++;			
			xls.SetCellText(0, nRow , szNewLine, FMT_TEXT,true);
			rs.GetValue(_T("insline"), szLineName);
			tmpStr.Format(_T("%s"), pMF->GetSelectionString(_T("hms_insline"), szLineName));
			xls.SetCellMergedColumns(1,nRow,16);
			xls.SetCellText(1, nRow , tmpStr, FMT_TEXT,true);
			szOldLine = szNewLine;			
			nIndex=1;
		}
		
		nRow++;

		rs.GetValue(_T("hdline"), szNewGroup);
		if (szNewGroup == _T("0") )
		{
			tmpStr.Format(_T("%s"), lszLine[0]);
			szSumOfline = _T("I");
		}
		else
		{
			tmpStr.Format(_T("%s"), lszLine[1]);
			szSumOfline = _T("II");
		}
		
		xls.SetCellText(0, nRow, szSumOfline, FMT_TEXT);		
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("total"), tmpStr);
		xls.SetCellText(2, nRow, tmpStr, FMT_INTEGER);
		rs.GetValue(_T("testfee"), cost);		
		grpLine[3] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("pacsfee"), cost);		
		grpLine[4] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("drugfee"), cost);		
		grpLine[5] += cost;			
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("bloodfee"), cost);		
		grpLine[6] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER1);


		rs.GetValue(_T("normtechfee"), cost);		
		grpLine[7] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("materialfee"), cost);		
		grpLine[8] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("hitechfee"), cost);		
		grpLine[9] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(9, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("drugfeek"), cost);		
		grpLine[10] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(10, nRow, tmpStr, FMT_NUMBER1);
		

		rs.GetValue(_T("examfee"), cost);		
		grpLine[11] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1);
		

		rs.GetValue(_T("otherfee"), cost);		
		grpLine[12] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(12, nRow, tmpStr, FMT_NUMBER1);		

		rs.GetValue(_T("cost"), cost);		
		grpLine[13] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(13, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("patpaid"), cost);		
		grpLine[14] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(14, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("inspaid"), cost);		
		grpLine[15] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(15, nRow, tmpStr, FMT_NUMBER1);		
		rs.MoveNext();
	}

	for (int i = 0; i < 17; i++)
	{
			ttlCost[i] += grpLine[i];
	}
	
	if(grpLine[13] > 0)
	{
		nRow++;
		xls.SetCellMergedColumns(1,nRow,2);
		tmpStr.Format(_T("\x43\x1ED9ng(%s=(I+II)): "),szNewLine);				
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT,true);
		for (int i = 2; i < 17; i++)
		{					
			tmpStr.Format(_T("%.2f"),grpLine[i]);
			xls.SetCellText(i, nRow , tmpStr, FMT_NUMBER1,true);				
			grpLine[i] = 0;					
		}
	}
	if(ttlCost[13] > 0)
	{
		nRow ++;
		CString szField, szAmount;
		TranslateString(_T("Total Amount"), szAmount);	
		tmpStr.Format(_T("%s (A+B+C):"),szAmount );
		xls.SetCellMergedColumns(1,nRow,2);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT,true);
		for (int i = 2; i < 17; i++)
		{			
			tmpStr.Format(_T("%.2f"),ttlCost[i]);
			xls.SetCellText(i, nRow, tmpStr, FMT_NUMBER1,true);
		}
	}	
	xls.Save(_T("Exports/CHIPHIKHAMBENHBHYT-25A-THBHYT.xls"));
	EndWaitCursor();
} 

CString CHMSGeneralCostInsuraceReportTH25aDialog::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString szSQL, szWhere, szUserReceiver;
	if (m_bLaybncc==1)
		szWhere.Format(_T(" AND hd_emergency = 'Y'"));
	else
		if (m_bKhonglaybncc==1) 
			szWhere.Format(_T(" AND hd_emergency <> 'Y' OR hd_emergency is null) "));
	
	if (!m_szClerkKey.IsEmpty())
		szUserReceiver.Format(_T(" and hfe_staff ='%s' "), m_szClerkKey);
	szSQL.Format(_T(" SELECT SUM(total) AS total,") \
				_T("   hdline,") \
				_T("   linename,") \
				_T("   insline,") \
				_T("   SUM(drugfee)     AS drugfee,") \
				_T("   SUM(drugfeek)    AS drugfeek,") \
				_T("   SUM(bloodfee)    AS bloodfee ,") \
				_T("   SUM(testfee)     AS testfee ,") \
				_T("   SUM(pacsfee)     AS pacsfee ,") \
				_T("   SUM(normtechfee) AS normtechfee ,") \
				_T("   SUM(hitechfee)   AS hitechfee ,") \
				_T("   SUM(materialfee) AS materialfee ,") \
				_T("   SUM(examfee)     AS examfee ,") \
				_T("   SUM(cost)        AS cost,") \
				_T("   SUM(otherfee)    AS otherfee ,") \
				_T("   SUM(inspaid)     AS inspaid ,") \
				_T("   SUM(patpaid)     AS patpaid") \
				_T(" FROM") \
				_T("   (SELECT hfe_docno,") \
				_T("     1 AS total,") \
				_T("     CASE WHEN hd_insline ='Y' THEN 1 ELSE 0 END     AS hdline,") \
				_T("     hc_line AS insline,") \
				_T("     (SELECT distinct ss_default FROM sys_sel WHERE ss_id='hms_insline' AND ss_code=CAST(hc_line AS NVARCHAR2(1)) )                 AS linename,") \
				_T("     SUM(drugfee)      AS drugfee,") \
				_T("     SUM(drugfeek)     AS drugfeek,") \
				_T("     SUM(bloodfee)     AS bloodfee ,") \
				_T("     SUM(testfee)      AS testfee ,") \
				_T("     SUM(pacsfee)      AS pacsfee ,") \
				_T("     SUM(normtechfee)  AS normtechfee ,") \
				_T("     SUM(hitechfee)    AS hitechfee ,") \
				_T("     SUM(materialfee)  AS materialfee ,") \
				_T("     SUM(examfee)      AS examfee ,") \
				_T("     SUM(cost-difpaid) AS cost,") \
				_T("     SUM(otherfee)     AS otherfee ,") \
				_T("     SUM(discount)     AS inspaid ,") \
				_T("     SUM(patpaid)      AS patpaid") \
				_T("   FROM") \
				_T("     (SELECT xi.hfe_docno,") \
				_T("       CASE WHEN SUBSTR(hfe_group, 1, 2)='A1' AND hfe_group              <>'A1400' THEN xf.hfe_cost-xf.hfe_diffcost ELSE 0 END AS drugfee,") \
				_T("       CASE WHEN hfe_group ='A1400' THEN xf.hfe_cost-xf.hfe_diffcost ELSE 0 END AS drugfeek,") \
				_T("       CASE WHEN SUBSTR(hfe_group, 1, 2)='A2' THEN xf.hfe_cost-xf.hfe_diffcost ELSE 0 END AS bloodfee,") \
				_T("       CASE WHEN SUBSTR(hfe_group, 1, 2)='B1' THEN xf.hfe_cost-xf.hfe_diffcost ELSE 0 END AS testfee,") \
				_T("       CASE WHEN SUBSTR(hfe_group, 1, 2) IN('B2','B3') AND hfe_hitech                ='N' THEN xf.hfe_cost-xf.hfe_diffcost ELSE 0 END AS pacsfee,") \
				_T("       CASE WHEN SUBSTR(hfe_group, 1, 2) IN('B4','B5') AND hfe_hitech                ='N' THEN xf.hfe_cost-xf.hfe_diffcost ELSE 0 END AS normtechfee,") \
				_T("       CASE WHEN hfe_hitech='Y' THEN xf.hfe_cost-xf.hfe_diffcost ELSE 0 END AS hitechfee,") \
				_T("       CASE WHEN SUBSTR(hfe_group, 1, 2)='A9' THEN xf.hfe_cost-xf.hfe_diffcost ELSE 0 END AS materialfee,") \
				_T("       CASE WHEN hfe_group='D0000' THEN xf.hfe_cost-xf.hfe_diffcost ELSE 0 END AS examfee,") \
				_T("       CASE WHEN hfe_group='F0000' THEN xf.hfe_cost-xf.hfe_diffcost ELSE 0 END AS otherfee,") \
				_T("       xf.hfe_cost     AS cost,") \
				_T("       xf.hfe_discount AS discount,") \
				_T("       xf.hfe_patpaid  AS patpaid,") \
				_T("       xf.hfe_diffpaid  AS difpaid,") \
				_T("       hfe_patdebt  AS patdebt") \
				_T("     FROM hms_xfee_invoice xi") \
				_T("     LEFT JOIN hms_xfee xf ON(xf.hfe_invoiceno=xi.hfe_invoiceno)") \
				_T("     WHERE xf.hfe_status='P'") \
				_T("	 AND (xi.hfe_date) BETWEEN cast_string2DATE('%s') AND cast_string2DATE('%s')") \
				_T("     AND xi.hfe_type     ='I'") \
				_T("     AND xi.hfe_class   IN('E')") \
				_T("     AND xf.hfe_discount > 0 %s") \
				_T("     ) tbl") \
				_T("   LEFT JOIN hms_doc ON(hfe_docno=hd_docno)") \
				_T("   LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx     =hd_cardidx)") \
				_T("   LEFT JOIN hms_object ON(ho_id              =hd_object)") \
				_T("   WHERE ho_type        IN('I','C')") \
				_T("   AND hd_status         ='T'") \
				_T("   AND LENGTH(hc_cardno) > 1 %s") \
				_T("   GROUP BY hfe_docno, hc_line, hd_insline") \
				_T("   HAVING SUM(cost) > 0") \
				_T("   ) tblMain") \
				_T(" GROUP BY insline, linename, hdline") \
				_T(" ORDER BY insline,") \
				_T("   hdline"), m_szFromDate, m_szToDate, szUserReceiver, szWhere);
	return szSQL;

}