#include "stdafx.h"
#include "TMDetailIOTransbookReport.h"
#include "ReportDocument.h"
#include "Excel.h"
#include "HMSMainFrame.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnYearKillfocus();
} */
static void _OnInsuranceSelectFnc(CWnd *pWnd){
	 ((CTMDetailIOTransbookreport*)pWnd)->OnInsuranceSelect();
}
static void _OnDiscountSelectFnc(CWnd *pWnd){
	 ((CTMDetailIOTransbookreport*)pWnd)->OnDiscountSelect();
}
static void _OnServiceSelectFnc(CWnd *pWnd){
	 ((CTMDetailIOTransbookreport*)pWnd)->OnServiceSelect();
}

static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CTMDetailIOTransbookreport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMDetailIOTransbookreport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CTMDetailIOTransbookreport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CTMDetailIOTransbookreport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CTMDetailIOTransbookreport *)pWnd)->OnToDateCheckValue();
} 
static void _OnStatusSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMDetailIOTransbookreport* )pWnd)->OnStatusSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStatusSelendokFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnStatusSelendok();
}
/*static void _OnStatusSetfocusFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnDoctorKillfocus();
}*/
/*static void _OnDoctorKillfocusFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnDoctorKillfocus();
}*/
static long _OnStatusLoadDataFnc(CWnd *pWnd){
	return ((CTMDetailIOTransbookreport *)pWnd)->OnStatusLoadData();
}
/*static void _OnStatusAddNewFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnDoctorAddNew();
}*/

static void _OnFloorSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMDetailIOTransbookreport* )pWnd)->OnFloorSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnFloorSelendokFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnFloorSelendok();
}
/*static void _OnFloorSetfocusFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnDoctorKillfocus();
}*/
/*static void _OnDoctorKillfocusFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnDoctorKillfocus();
}*/
static long _OnFloorLoadDataFnc(CWnd *pWnd){
	return ((CTMDetailIOTransbookreport *)pWnd)->OnFloorLoadData();
}
/*static void _OnFloorAddNewFnc(CWnd *pWnd){
	((CTMDetailIOTransbookreport *)pWnd)->OnDoctorAddNew();
}*/

static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CTMDetailIOTransbookreport *pVw = (CTMDetailIOTransbookreport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CTMDetailIOTransbookreport *pVw = (CTMDetailIOTransbookreport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CTMDetailIOTransbookreport *pVw = (CTMDetailIOTransbookreport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddCTMDetailIOTransbookreportFnc(CWnd *pWnd){
	 return ((CTMDetailIOTransbookreport*)pWnd)->OnAddCTMDetailIOTransbookreport();
} 
static int _OnEditCTMDetailIOTransbookreportFnc(CWnd *pWnd){
	 return ((CTMDetailIOTransbookreport*)pWnd)->OnEditCTMDetailIOTransbookreport();
} 
static int _OnDeleteCTMDetailIOTransbookreportFnc(CWnd *pWnd){
	 return ((CTMDetailIOTransbookreport*)pWnd)->OnDeleteCTMDetailIOTransbookreport();
} 
static int _OnSaveCTMDetailIOTransbookreportFnc(CWnd *pWnd){
	 return ((CTMDetailIOTransbookreport*)pWnd)->OnSaveCTMDetailIOTransbookreport();
} 
static int _OnCancelCTMDetailIOTransbookreportFnc(CWnd *pWnd){
	 return ((CTMDetailIOTransbookreport*)pWnd)->OnCancelCTMDetailIOTransbookreport();
}
static void _OnNoPatientDialysisToReportSelectFnc(CWnd *pWnd){
	 ((CTMDetailIOTransbookreport*)pWnd)->OnNoPatientDialysisToReportSelect();
}
static void _OnNoPatientTerminatedToReportSelectFnc(CWnd *pWnd){
	 ((CTMDetailIOTransbookreport*)pWnd)->OnNoPatientTerminatedToReportSelect();
}
CTMDetailIOTransbookreport::CTMDetailIOTransbookreport(CWnd *pParent)
	{

	m_nDlgWidth = 460;
	m_nDlgHeight = 340;
	SetDefaultValues();
	m_bPreview = true;
}
CTMDetailIOTransbookreport::~CTMDetailIOTransbookreport(){
}
void CTMDetailIOTransbookreport::OnCreateComponents()
{
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 110, 55);
	m_wndYear.Create(this,115, 30, 255, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 260, 30, 360, 55);
	m_wndReportPeriod.Create(this,365, 30, 505, 55); 
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 510, 150);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 110, 85);
	m_wndFromDate.Create(this,115, 60, 255, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 260, 60, 360, 85);
	m_wndToDate.Create(this,365, 60, 505, 85); 
	m_wndFloor.SetCheckBox(true);
	m_wndFloorLabel.Create(this, _T("Floor"), 10, 90, 110, 115);
	m_wndFloor.Create(this,115, 90, 505, 115); 
	m_wndStatusLabel.Create(this, _T("Status"), 10, 120, 110, 145);
	m_wndStatus.Create(this,115, 120, 505, 145); 
	m_wndNotPatientFloor.Create(this, _T("Not Patient Floor To Report"), 5, 155, 255, 180);
	m_wndNoPatientSurgeryToReport.Create(this, _T("No Patient Surgery To Report"), 260, 155, 510, 180);
	m_wndNoPatientDialysisToReport.Create(this, _T("No Patient Dialysis To Report"), 5, 185, 255, 210);
	m_wndNoPatientTerminatedToReport.Create(this, _T("No Patient Terminated To Report"), 260, 185, 510, 210);
	m_wndInsurance.Create(this, _T("Insurance"), 5, 215, 105, 240);
	m_wndDiscount.Create(this, _T("Discount"), 110, 215, 210, 240);
	m_wndService.Create(this, _T("Service"), 215, 215, 320, 240);
	m_wndPrintPreview.Create(this, _T("&Print"), 345, 215, 425, 240);
	m_wndPrint.Create(this, _T("&Export"), 430, 215, 510, 240);
}
void CTMDetailIOTransbookreport::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(true);
//	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
//	m_wndDoctor.SetCheckValue(true);
	m_wndStatus.LimitText(35);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 60);


	m_wndStatus.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndStatus.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	
	m_wndFloor.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 100);
	m_wndFloor.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);


}
void CTMDetailIOTransbookreport::OnSetWindowEvents(){
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

	m_wndFloor.SetEvent(WE_SELENDOK, _OnFloorSelendokFnc);
	//m_wndFloor.SetEvent(WE_SETFOCUS, _OnFloorSetfocusFnc);
	//m_wndFloor.SetEvent(WE_KILLFOCUS, _OnFloorKillfocusFnc);
	m_wndFloor.SetEvent(WE_SELCHANGE, _OnFloorSelectChangeFnc);
	m_wndFloor.SetEvent(WE_LOADDATA, _OnFloorLoadDataFnc);
	//m_wndFloor.SetEvent(WE_ADDNEW, _OnFloorAddNewFnc);

	m_wndStatus.SetEvent(WE_SELENDOK, _OnStatusSelendokFnc);
	//m_wndStatus.SetEvent(WE_SETFOCUS, _OnStatusSetfocusFnc);
	//m_wndStatus.SetEvent(WE_KILLFOCUS, _OnStatusKillfocusFnc);
	m_wndStatus.SetEvent(WE_SELCHANGE, _OnStatusSelectChangeFnc);
	m_wndStatus.SetEvent(WE_LOADDATA, _OnStatusLoadDataFnc);
	//m_wndStatus.SetEvent(WE_ADDNEW, _OnStatusAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);

	m_wndInsurance.SetEvent(WE_CLICK, _OnInsuranceSelectFnc);
	m_wndDiscount.SetEvent(WE_CLICK, _OnDiscountSelectFnc);
	m_wndService.SetEvent(WE_CLICK, _OnServiceSelectFnc);
	m_wndNoPatientDialysisToReport.SetEvent(WE_CLICK, _OnNoPatientDialysisToReportSelectFnc);
	m_wndNoPatientTerminatedToReport.SetEvent(WE_CLICK, _OnNoPatientTerminatedToReportSelectFnc);


	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString tmpStr;
	tmpStr = pMF->GetSysDate();
	m_nYear = ToInt(tmpStr.Left(4));
	m_szFromDate = tmpStr + _T("00:00");
	m_szToDate = tmpStr + _T("23:59");
	m_szReportPeriodKey.Format(_T("%d"), ToInt(tmpStr.Mid(5, 2)));
	UpdateData(FALSE);
}
void CTMDetailIOTransbookreport::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStatus.GetDlgCtrlID(), m_szStatusKey);
	DDX_TextEx(pDX, m_wndFloor.GetDlgCtrlID(), m_szFloorKey);
	DDX_Check(pDX, m_wndInsurance.GetDlgCtrlID(), m_bInsurance);
	DDX_Check(pDX, m_wndDiscount.GetDlgCtrlID(), m_bDiscount);
	DDX_Check(pDX, m_wndService.GetDlgCtrlID(), m_bService);
	DDX_Check(pDX, m_wndNotPatientFloor.GetDlgCtrlID(), m_bNotPatientFloor);
	DDX_Check(pDX, m_wndNoPatientSurgeryToReport.GetDlgCtrlID(), m_bNoPatientSurgeryToReport);
	DDX_Check(pDX, m_wndNoPatientDialysisToReport.GetDlgCtrlID(), m_bNoPatientDialysisToReport);
	DDX_Check(pDX, m_wndNoPatientTerminatedToReport.GetDlgCtrlID(), m_bNoPatientTerminatedToReport);
}
void CTMDetailIOTransbookreport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CTMDetailIOTransbookreport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CTMDetailIOTransbookreport::SetDefaultValues()
{
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStatusKey.Empty();
	m_szFloorKey.Empty();
	m_bInsurance=TRUE;
	m_bDiscount=TRUE;
	m_bService=TRUE;
	m_bNotPatientFloor=FALSE;
	m_bNoPatientSurgeryToReport=TRUE;
	m_bNoPatientDialysisToReport=FALSE;
	m_bNoPatientTerminatedToReport=FALSE;
}
int CTMDetailIOTransbookreport::SetMode(int nMode)
{
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
/*void CTMDetailIOTransbookreport::OnYearChange(){
	
} */
/*void CTMDetailIOTransbookreport::OnYearSetfocus(){
	
} */
/*void CTMDetailIOTransbookreport::OnYearKillfocus(){
	
} */
int CTMDetailIOTransbookreport::OnYearCheckValue(){
	return 0;
} 
void CTMDetailIOTransbookreport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMDetailIOTransbookreport::OnReportPeriodSelendok(){
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
/*void CTMDetailIOTransbookreport::OnReportPeriodSetfocus(){
	
}*/
/*void CTMDetailIOTransbookreport::OnReportPeriodKillfocus(){
	
}*/
long CTMDetailIOTransbookreport::OnReportPeriodLoadData(){
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
/*void CTMDetailIOTransbookreport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CTMDetailIOTransbookreport::OnFromDateChange(){
	
} */
/*void CTMDetailIOTransbookreport::OnFromDateSetfocus(){
	
} */
/*void CTMDetailIOTransbookreport::OnFromDateKillfocus(){
	
} */
int CTMDetailIOTransbookreport::OnFromDateCheckValue(){
	return 0;
} 
/*void CTMDetailIOTransbookreport::OnToDateChange(){
	
} */
/*void CTMDetailIOTransbookreport::OnToDateSetfocus(){
	
} */
/*void CTMDetailIOTransbookreport::OnToDateKillfocus(){
	
} */
int CTMDetailIOTransbookreport::OnToDateCheckValue(){
	return 0;
} 
void CTMDetailIOTransbookreport::OnStatusSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMDetailIOTransbookreport::OnStatusSelendok(){
	 
}
/*void CTMDetailIOTransbookreport::OnStatusSetfocus(){
	
}*/
/*void CTMDetailIOTransbookreport::OnStatusKillfocus(){
	
}*/
long CTMDetailIOTransbookreport::OnStatusLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();	

	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStatus.IsSearchKey() && !m_szStatusKey.IsEmpty()){
		szWhere.Format(_T(" AND ss_code ='%s' ") , m_szStatusKey);
	};
	szSQL.Format(_T(" SELECT ss_code as id,ss_desc as name FROM sys_sel WHERE ss_id ='hms_suggestion' and ss_code in('A','T','D','F','M')  %s"), szWhere);
	m_wndStatus.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStatus.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CTMDetailIOTransbookreport::OnStatusAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */

void CTMDetailIOTransbookreport::OnFloorSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMDetailIOTransbookreport::OnFloorSelendok(){
	 
}
/*void CTMDetailIOTransbookreport::OnFloorSetfocus(){
	
}*/
/*void CTMDetailIOTransbookreport::OnFloorKillfocus(){
	
}*/
long CTMDetailIOTransbookreport::OnFloorLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere, szGroupID;
	szWhere.Empty();
	
	szSQL.Format(_T(" SELECT sd_groupid FROM sys_dept WHERE sd_id = '%s' "), pMF->m_szDept);
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{
		rs.GetValue(_T("sd_groupid"), szGroupID);
	}

	if(m_wndFloor.IsSearchKey() && !m_szFloorKey.IsEmpty()){
		//szWhere.Format(_T(" AND sd_id ='%s' ") , m_szFloorKey);
	};
	szSQL.Format(_T(" SELECT sd_id as id,sd_name as name FROM sys_dept WHERE sd_type = 'DT' and sd_groupid ='%s' %s"), szGroupID,szWhere);
	m_wndFloor.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndFloor.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CTMDetailIOTransbookreport::OnFloorAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */

void CTMDetailIOTransbookreport::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	m_bPreview = true;
	OnReportInpatientpatientlistreport();
	//UpdateData(false);
	
} 
void CTMDetailIOTransbookreport::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	OnExport();
	//UpdateData(false);
} 
void CTMDetailIOTransbookreport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

} 
int CTMDetailIOTransbookreport::OnAddCTMDetailIOTransbookreport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0;
}
int CTMDetailIOTransbookreport::OnEditCTMDetailIOTransbookreport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CTMDetailIOTransbookreport::OnDeleteCTMDetailIOTransbookreport(){
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
 		OnCancelCTMDetailIOTransbookreport(); 
 	}
	return 0;
}
int CTMDetailIOTransbookreport::OnSaveCTMDetailIOTransbookreport(){
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
 		//OnCTMDetailIOTransbookreportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CTMDetailIOTransbookreport::OnCancelCTMDetailIOTransbookreport(){
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
int CTMDetailIOTransbookreport::OnCTMDetailIOTransbookreportListLoadData(){
	return 0;
}
void CTMDetailIOTransbookreport::OnInsuranceSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CTMDetailIOTransbookreport::OnDiscountSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CTMDetailIOTransbookreport::OnServiceSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CTMDetailIOTransbookreport::OnNoPatientDialysisToReportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CTMDetailIOTransbookreport::OnNoPatientTerminatedToReportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CTMDetailIOTransbookreport::OnReportInpatientpatientlistreport()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr,tmpStr2;
	CString szDate, szSysDate, szAge, szFloor,szFloorName;
	CString szFromDate, szToDate;
	CReport rpt; 
	//UpdateData(true);
	szSysDate = pMF->GetSysDate(); 
	szFromDate = m_szFromDate;
	szToDate = m_szToDate;
	
	if(!rpt.Init(_T("Reports/HMS/TM_SOVAOVIENRAVIENCHUYENVIENTHEOKHOA.RPT")))
	{
		return ;
	}
	
	UpdateData(true);
	BeginWaitCursor();
	//Report Detail	
	szSQL = GetQueryString();
_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	
	tmpStr = rpt.GetReportHeader()->GetValue(_T("ReportDate"));
	szDate.Format(tmpStr, CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	rpt.GetReportHeader()->SetValue(_T("Department"), szFloorName);
	
	CString szNewLine, szOldLine;
	int nIndex = 1 ;
	long docno = 0, age = 0, nTotalPatient = 0;
	long c[35];
	for (int i = 0; i < 35;i++)
	{
		c[i] = 0;
	}

	
	CReportSection* rptDetail = rpt.GetDetail(); 
	while(!rs.IsEOF())
	{ 			
		/*rs.GetValue(_T("deptname"), szNewLine);
		if(szNewLine != szOldLine && !szNewLine.IsEmpty())
		{
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rptDetail->SetValue(_T("GroupName"), szNewLine);
			szOldLine = szNewLine;
			nIndex=1;
		}*/
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%d"), nIndex);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("recordno"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("patientname"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("sex"), tmpStr);
		if(tmpStr == _T("M"))
		{
			rs.GetValue(_T("age"), tmpStr);
			rptDetail->SetValue(_T("5"), tmpStr);
			c[4]++;
		}
		else 
		{
			rs.GetValue(_T("age"), tmpStr);
			rptDetail->SetValue(_T("6"), tmpStr);
			c[5]++;
		}
		
		rs.GetValue(_T("cardno"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[6]++;
		rptDetail->SetValue(_T("7"), tmpStr);
		
		rs.GetValue(_T("tre12thang"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[7]++;
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("tre6tuoi"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[8]++;
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("tre15tuoi"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[9]++;
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.GetValue(_T("tre30ngay"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[10]++;
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.GetValue(_T("occupation"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[11]++;
		rptDetail->SetValue(_T("12"), tmpStr);

		rs.GetValue(_T("address"), tmpStr);
		if(!tmpStr.IsEmpty())
			c[12]++;
		rptDetail->SetValue(_T("13"), tmpStr);

		rs.GetValue(_T("transferplace"), tmpStr);
		if(!tmpStr.IsEmpty())
			c[13]++;
		rptDetail->SetValue(_T("14"), tmpStr);
		
		rs.GetValue(_T("deptname"), tmpStr);		
		rptDetail->SetValue(_T("15"), tmpStr);

		rs.GetValue(_T("admitdate"), tmpStr);
		tmpStr2.Format(_T("%s"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
		if(!tmpStr.IsEmpty()) 
			c[14]++;
		rptDetail->SetValue(_T("15"), tmpStr2);

		rs.GetValue(_T("dischargedate"), tmpStr);		
		if(!tmpStr.IsEmpty())
		{
			c[15]++;
			tmpStr2.Format(_T("%s"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
			rptDetail->SetValue(_T("16"), tmpStr2);
		}
		
		rs.GetValue(_T("transhospitall"), tmpStr);
		if(!tmpStr.IsEmpty()) 
		{
			c[16]++;
			tmpStr2.Format(_T("%s"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
			rptDetail->SetValue(_T("17"), tmpStr2);
		}
		
		rs.GetValue(_T("tuvongtruoc24"), tmpStr);		
		if(!tmpStr.IsEmpty())		
			c[17]++;

		rptDetail->SetValue(_T("18"), tmpStr);
		
		rs.GetValue(_T("tuvongsau24"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[18]++;
		rptDetail->SetValue(_T("19"), tmpStr);
		
		
		rs.GetValue(_T("transferdiagnostic"), tmpStr);
		if (!tmpStr.IsEmpty())
		{
			tmpStr.AppendFormat(_T("(%s)"), rs.GetValue(_T("icd")));
			c[19]++;
		}
		rptDetail->SetValue(_T("20"), tmpStr);

		rs.GetValue(_T("diagnostic"), tmpStr);
		if (!tmpStr.IsEmpty())
		{
			tmpStr.AppendFormat(_T("(%s)"), rs.GetValue(_T("icd")));
			c[20]++;
		}
		rptDetail->SetValue(_T("21"), tmpStr);

		rs.GetValue(_T("maindisease"), tmpStr);
		if (!tmpStr.IsEmpty())
			c[21]++;

		rptDetail->SetValue(_T("22"), tmpStr);

		rs.GetValue(_T("dischargedept"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[22]++;
		rptDetail->SetValue(_T("23"), tmpStr);

		rs.GetValue(_T("fromdept"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[23]++;
		rptDetail->SetValue(_T("24"), tmpStr);

		rs.GetValue(_T("todept"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[24]++;
		rptDetail->SetValue(_T("25"), tmpStr);

		rs.GetValue(_T("dischargefloor"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[25]++;
		rptDetail->SetValue(_T("26"), tmpStr);

		rs.GetValue(_T("khoi"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[26]++;
		rptDetail->SetValue(_T("27"), tmpStr);

		rs.GetValue(_T("khongthaydoi"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[27]++;
		rptDetail->SetValue(_T("28"), tmpStr);
		
		rs.GetValue(_T("dogiam"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[28]++;
		rptDetail->SetValue(_T("29"), tmpStr);
		
		rs.GetValue(_T("nang"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[29]++;
		rptDetail->SetValue(_T("30"), tmpStr);
		
		rs.GetValue(_T("xinve"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[30]++;
		rptDetail->SetValue(_T("31"), tmpStr);

		rs.GetValue(_T("tronvien"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[31]++;	
		rptDetail->SetValue(_T("32"), tmpStr);

		rs.GetValue(_T("treatqty"), tmpStr);
		if(ToInt(tmpStr) > 0)
		{
			c[32] += ToInt(tmpStr);
			rptDetail->SetValue(_T("33"), tmpStr);
		}
		else
		{
			tmpStr.Empty();
			rs.GetValue(_T("ngaydt"), tmpStr);
			if (ToInt(tmpStr) > 0)
			{
				c[33] += ToInt(tmpStr);
				rptDetail->SetValue(_T("34"), tmpStr);
			}
		}

		nIndex++;
		rs.MoveNext();
	}
	
	for(i = 4; i < 35; i++)
	{	
		CString szFied;
		szFied.Format(_T("s%d"), i+1);
		tmpStr.Format(_T("%ld"), c[i]);				
		rpt.GetReportFooter()->SetValue(szFied, tmpStr);
	}

	tmpStr.Format(_T("%d"), docno);
	//rpt.GetReportHeader()->SetValue(_T("ToNum"), tmpStr);
	if (rs.GetRecordCount() > 0)
		tmpStr.Format(_T("%ld"), rs.GetRecordCount());
	rpt.GetReportHeader()->SetValue(_T("TotalPatient"), tmpStr);	
	
	float congxuat=0;
	congxuat = (float) c[32] / nIndex;
	
	tmpStr.Format(_T("%.2f"), congxuat);
	rpt.GetReportHeader()->SetValue(_T("TreatNumber"), tmpStr);
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);

	EndWaitCursor();
	rpt.PrintPreview();

}
void CTMDetailIOTransbookreport::OnExport()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	CString szDate, szSysDate;
	CString szFromDate, szToDate,szAge, szFloorName, tmpStr2;
	CReport rpt; 

	szSysDate = pMF->GetSysDate(); 
	szFromDate = m_szFromDate;
	szToDate = m_szToDate;

	UpdateData(true);
	BeginWaitCursor();
	//Report Detail	
	
	szSQL = GetQueryString();
_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 11);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 20);
	xls.SetColumnWidth(4, 6);
	xls.SetColumnWidth(5, 6);
	xls.SetColumnWidth(6, 6);
	xls.SetColumnWidth(7, 6);
	xls.SetColumnWidth(8, 6);
	xls.SetColumnWidth(9, 6);
	xls.SetColumnWidth(10, 20);
	xls.SetColumnWidth(11, 15);
	xls.SetColumnWidth(12, 42);
	xls.SetColumnWidth(13, 20);
	xls.SetColumnWidth(14, 20);
	xls.SetColumnWidth(15, 16);
	xls.SetColumnWidth(16, 16);
	xls.SetColumnWidth(17, 16);
	xls.SetColumnWidth(18, 7);
	xls.SetColumnWidth(19, 7);
	xls.SetColumnWidth(20, 50);
	xls.SetColumnWidth(21, 50);
	xls.SetColumnWidth(22, 50);
	xls.SetColumnWidth(23, 7);
	xls.SetColumnWidth(24, 12);
	xls.SetColumnWidth(25, 12);
	xls.SetColumnWidth(26, 10);
	xls.SetColumnWidth(27, 11);
	xls.SetColumnWidth(28, 7);
	xls.SetColumnWidth(29, 10);
	xls.SetColumnWidth(30, 19);
	xls.SetColumnWidth(31, 19);
	int nRow = 0, nCol = 0;
	
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellMergedColumns(nCol, nRow + 1, 4);
	xls.SetCellMergedColumns(nCol, nRow + 2, 32);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT, true, 10);
	xls.SetCellText(nCol, nRow + 2, szFloorName, FMT_TEXT, true, 10);
	xls.SetCellMergedColumns(nCol, nRow + 4, 17);
	xls.SetCellText(nCol, nRow + 4, _T("S\x1ED4 V\xC0O VI\x1EC6N – R\x41 VI\x1EC6N – \x43HUY\x1EC2N VI\x1EC6N"), FMT_TEXT | FMT_CENTER, true, 16);
	if (m_bNoPatientTerminatedToReport)
		xls.SetCellText(nCol, nRow + 5, _T("(\x42\x1EC7nh nh\xE2n t\x1ED3n \x63u\x1ED1i k\x1EF3)"), FMT_TEXT, true, 10);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDate::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(nCol, nRow + 6, tmpStr, FMT_TEXT, true, 10);
	
	//xls.SetCellMergedColumns(nCol, nRow + 6, 32);
	
	xls.SetCellMergedColumns(nCol, nRow + 7, 3);
	if (rs.GetRecordCount() > 0)
	{
		tmpStr.Format(_T("%d"), rs.GetRecordCount());
		xls.SetCellText(nCol, nRow + 7, _T("T\x1ED5ng s\x1ED1 \x62\x1EC7nh nh\xE2n: ") +  tmpStr, FMT_TEXT, true, 10);
	}
	
	xls.SetCellMergedRows(nCol, nRow + 8, 3);
	xls.SetCellText(nCol, nRow + 8, _T("STT"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 1, nRow + 8, 3);
	TranslateString(_T("Record No"), tmpStr);
	xls.SetCellText(nCol + 1, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 2, nRow + 8, 3);
	TranslateString(_T("Document No"), tmpStr);
	xls.SetCellText(nCol + 2, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 3, nRow + 8, 3);
	TranslateString(_T("Patient Name"), tmpStr);	
	xls.SetCellText(nCol + 3, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol + 4, nRow + 8, 2);
	TranslateString(_T("Age"), tmpStr);
	xls.SetCellText(nCol + 4, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 4, nRow + 9, 2);
	TranslateString(_T("Male"), tmpStr);
	xls.SetCellText(nCol + 4, nRow + 9, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 5, nRow + 9, 2);
	TranslateString(_T("Female"), tmpStr);
	xls.SetCellText(nCol + 5, nRow + 9, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol + 6, nRow + 8, 4);
	xls.SetCellText(nCol + 6, nRow + 8, _T("Tr\x1EBB < \x31\x36T"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 6, nRow + 9, 2);
	xls.SetCellText(nCol + 6, nRow + 9, _T("<12th"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 7, nRow + 9, 2);
	xls.SetCellText(nCol + 7, nRow + 9, _T("1-6T"), FMT_TEXT | FMT_CENTER, true, 10);	

	xls.SetCellMergedRows(nCol + 8, nRow + 9, 2);
	xls.SetCellText(nCol + 8, nRow + 9, _T("6-15T"), FMT_TEXT | FMT_CENTER, true, 10);	

	xls.SetCellMergedRows(nCol + 9, nRow + 9, 2);
	xls.SetCellText(nCol + 9, nRow + 9, _T("<=30 Ng\xE0y"), FMT_TEXT | FMT_CENTER, true, 10);	

	xls.SetCellMergedRows(nCol + 10, nRow + 8, 3);
	xls.SetCellText(nCol + 10, nRow + 8, _T("S\x1ED1 \x42HYT"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 11, nRow + 8, 3);
	TranslateString(_T("Occupation"), tmpStr);
	xls.SetCellText(nCol + 11, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);	

	xls.SetCellMergedRows(nCol + 12, nRow + 8, 3);
	TranslateString(_T("Address"), tmpStr);
	xls.SetCellText(nCol + 12, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 13, nRow + 8, 3);
	TranslateString(_T("Instroduce Place"), tmpStr);
	xls.SetCellText(nCol + 13, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);	

	xls.SetCellMergedRows(nCol + 14, nRow + 8, 3);
	xls.SetCellText(nCol + 14, nRow + 8, _T("T\xEAn kho\x61 \x110T"), FMT_TEXT | FMT_CENTER, true, 10);	
	
	xls.SetCellMergedColumns(nCol + 15, nRow + 8, 5);
	xls.SetCellText(nCol + 15, nRow + 8, _T("\x43\xE1\x63h gi\x1EA3i quy\x1EBFt/\x43huy\x1EC3n"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 15, nRow + 9, 2);
	TranslateString(_T("Admit Date"), tmpStr);
	xls.SetCellText(nCol + 15, nRow + 9, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 16, nRow + 9, 2);
	TranslateString(_T("Discharge Date"), tmpStr);
	xls.SetCellText(nCol + 16, nRow + 9, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 17, nRow + 9, 2);
	TranslateString(_T("Hospital Transfer"), tmpStr);
	xls.SetCellText(nCol + 17, nRow + 9, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);	

	xls.SetCellMergedColumns(nCol + 18, nRow + 9, 2);
	TranslateString(_T("Death"), tmpStr);
	xls.SetCellText(nCol + 18, nRow + 9, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 18, nRow + 10, _T("< 24h"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 19, nRow + 10, _T("> 24h"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol + 20, nRow + 8, 3);
	TranslateString(_T("Diagnostic"), tmpStr);
	xls.SetCellText(nCol + 20, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);		
	
	xls.SetCellMergedRows(nCol + 20, nRow + 9, 2);
	xls.SetCellText(nCol + 20, nRow + 9, _T("Tuy\x1EBFn \x64\x1B0\x1EDBi"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 21, nRow + 9, 2);
	xls.SetCellText(nCol + 21, nRow + 9, _T("Kho\x61  K\x42"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 22, nRow + 9, 2);
	xls.SetCellText(nCol + 22, nRow + 9, _T("Kho\x61  \x110T"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol + 23, nRow + 8, 2);
	TranslateString(_T("Department Transfer"), tmpStr);
	xls.SetCellText(nCol + 23, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	TranslateString(_T("Department Transfer"), tmpStr);
	xls.SetCellText(nCol + 23, nRow + 10, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);
		
	TranslateString(_T("From Department"), tmpStr);
	xls.SetCellText(nCol + 24, nRow + 10, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);
	
	TranslateString(_T("To Department"), tmpStr);
	xls.SetCellText(nCol + 25, nRow + 10, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 26, nRow + 8, 3);
	TranslateString(_T("Floor Transfer"), tmpStr);
	xls.SetCellText(nCol + 26, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 27, nRow + 8, 3);
	TranslateString(_T("Good"), tmpStr);
	xls.SetCellText(nCol + 27, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 28, nRow + 8, 3);
	TranslateString(_T("No change"), tmpStr);
	xls.SetCellText(nCol + 28, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 29, nRow + 8, 3);
	TranslateString(_T("Help reduce"), tmpStr);
	xls.SetCellText(nCol + 29, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 30, nRow + 8, 3);
	TranslateString(_T("Heavy"), tmpStr);
	xls.SetCellText(nCol + 30, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 31, nRow + 8, 3);
	xls.SetCellText(nCol + 31, nRow + 8, _T("\x58in v\x1EC1"), FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedRows(nCol + 32, nRow + 8, 3);
	xls.SetCellText(nCol + 32, nRow + 8, _T("Tr\x1ED1n vi\x1EC7n"), FMT_TEXT | FMT_CENTER, true, 10);
	

	xls.SetCellMergedColumns(nCol + 33, nRow + 8, 3);
	TranslateString(_T("Ng\xE0y \x111i\x1EC1u tr\x1ECB"), tmpStr);
	xls.SetCellText(nCol + 33, nRow + 8, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);
	
	TranslateString(_T("Th\x1EF1\x63 t\x1EBF"), tmpStr);
	xls.SetCellText(nCol + 33, nRow + 10, tmpStr, FMT_TEXT | FMT_CENTER, true, 10);	
	xls.SetCellText(nCol + 34, nRow + 10, _T("T\x1EA1m t\xEDnh"), FMT_TEXT | FMT_CENTER, true, 10);	
	xls.SetCellText(nCol + 35, nRow + 10, _T("T\x1ED5ng"), FMT_TEXT | FMT_CENTER, true, 10);

	CString szNewLine, szOldLine;
	int nIndex = 1, nTreat=0 ;
	long docno = 0, age = 0, nTotalPatient = 0;
	long c[36];
	for (int i = 0; i < 35;i++)
	{
		c[i] = 0;
	}

	nRow += 11;

	while(!rs.IsEOF())
	{ 			
		/*rs.GetValue(_T("deptname"), szNewLine);
		if(szNewLine != szOldLine && !szNewLine.IsEmpty())
		{
			xls.SetCellMergedColumns(nCol, nRow, 35);
			xls.SetCellText(nCol, nRow, szNewLine, FMT_TEXT, true);
			szOldLine = szNewLine;
			nIndex=1;
			nRow++;
		}	*/
		tmpStr.Format(_T("%d"), nIndex);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		
		rs.GetValue(_T("recordno"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("patientname"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("age"), szAge);		

		rs.GetValue(_T("sex"), tmpStr);
		if(tmpStr == _T("M"))
		{
			xls.SetCellText(nCol + 4, nRow, szAge, FMT_TEXT);
			c[4]++;
		}
		else
		{
			xls.SetCellText(nCol + 5, nRow, szAge, FMT_TEXT);
			c[5]++;
		}
					
		rs.GetValue(_T("tre12thang"), tmpStr);
		if(!tmpStr.IsEmpty())
			c[6]++;
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("tre6tuoi"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[7]++;
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("tre15tuoi"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[8]++;
		xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("tre30Ngay"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[9]++;
		xls.SetCellText(nCol + 9, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("cardno"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[10]++;
		xls.SetCellText(nCol + 10, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("occupation"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[11]++;
		xls.SetCellText(nCol + 11, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("address"), tmpStr);
		if(!tmpStr.IsEmpty())
			c[12]++;
		xls.SetCellText(nCol + 12, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("transferplace"), tmpStr);
		if(!tmpStr.IsEmpty())
			c[13]++;
		xls.SetCellText(nCol + 13, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("deptname"), tmpStr);
		if(!tmpStr.IsEmpty())
			c[14]++;
		xls.SetCellText(nCol + 14, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("admitdate"), tmpStr);		
		if(!tmpStr.IsEmpty()) 
		{
			c[15]++;
			tmpStr2.Format(_T("%s"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
			xls.SetCellText(nCol + 15, nRow, tmpStr2, FMT_TEXT);
		}
		
		
		rs.GetValue(_T("dischargedate"), tmpStr);		
		if(!tmpStr.IsEmpty()) 
		{
			tmpStr2.Format(_T("%s"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
			xls.SetCellText(nCol + 16, nRow, tmpStr2, FMT_TEXT);
			c[16]++;
		}
		
		
		rs.GetValue(_T("transhospitall"), tmpStr);		
		if(!tmpStr.IsEmpty()) 
		{
			c[17]++;
			tmpStr2.Format(_T("%s"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
			xls.SetCellText(nCol + 17, nRow, tmpStr2, FMT_TEXT);		
		}
		
		rs.GetValue(_T("tuvongtruoc24"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[18]++;
		xls.SetCellText(nCol + 18, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("tuvongsau24"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[19]++;
		xls.SetCellText(nCol + 19, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("transferdiagnostic"), tmpStr);
		if (!tmpStr.IsEmpty())
		{
			tmpStr.AppendFormat(_T("(%s)"), rs.GetValue(_T("icd")));
			c[20]++;
		}
		xls.SetCellText(nCol + 20, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("diagnostic"), tmpStr);
		if (!tmpStr.IsEmpty())
		{
			tmpStr.AppendFormat(_T("(%s)"), rs.GetValue(_T("icd")));
			c[21]++;
		}
		xls.SetCellText(nCol + 21, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("maindisease"), tmpStr);
		if (!tmpStr.IsEmpty())
			c[22]++;

		xls.SetCellText(nCol + 22, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("dischargedept"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[23]++;
		xls.SetCellText(nCol + 23, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("fromdept"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[24]++;
		xls.SetCellText(nCol + 24, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("todept"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[25]++;
		xls.SetCellText(nCol + 25, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("dischargefloor"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[26]++;
		xls.SetCellText(nCol + 26, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("khoi"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[27]++;
		xls.SetCellText(nCol + 27, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("khongthaydoi"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[28]++;
		xls.SetCellText(nCol + 28, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("dogiam"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[29]++;
		xls.SetCellText(nCol + 29, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("nang"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[30]++;
		xls.SetCellText(nCol + 30, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("xinve"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[31]++;
		xls.SetCellText(nCol + 31, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("tronvien"), tmpStr);
		if(!tmpStr.IsEmpty()) 
			c[32]++;	
		xls.SetCellText(nCol + 32, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("treatqty"), tmpStr);
		c[33] += ToInt(tmpStr);
		nTreat = ToInt(tmpStr);
		xls.SetCellText(nCol + 33, nRow, tmpStr, FMT_INTEGER);
	
		rs.GetValue(_T("ngaydt"), tmpStr);
		c[34] += ToInt(tmpStr);
		xls.SetCellText(nCol + 34, nRow, tmpStr, FMT_INTEGER);
		
		c[35] = c[33]+c[34];

		nTreat += ToInt(tmpStr);
		
		tmpStr.Format(_T("%d"), nTreat);
		xls.SetCellText(nCol + 35, nRow, tmpStr, FMT_INTEGER);

		nIndex++;
		nRow++;
		rs.MoveNext();
	}

	TranslateString(_T("Total Sum"), tmpStr);
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);	


	for(i = 4; i < 36; i++)
	{			
		if (c[i] > 0)
		{
			tmpStr.Format(_T("%ld"), c[i]);				
			xls.SetCellText(i, nRow, tmpStr, FMT_INTEGER, true);
		}
	}
	

	float congxuat=0;
	congxuat = (float) c[33] / nIndex;
	
	tmpStr.Format(_T("C\xF4ng su\x1EA5t gi\x1B0\x1EDDng \x64\xF9ng: %.2f"), congxuat);
	xls.SetCellMergedColumns(5, 7, 10);
	xls.SetCellText(5, 7, tmpStr, FMT_TEXT,true);

	EndWaitCursor();
	xls.Save(_T("Exports\\SOVAOVIENRAVIENCHUYENVIENTHEOKHOA.XLS"));	
}

CString CTMDetailIOTransbookreport::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szWhereObjects, tmpStr2, szObjects, szWhere, szWhereDate, szFloor, szFloorName, tmpStr, szWhereFromDate, szDept, szCondition;
	szWhereObjects.Empty();
	tmpStr2 = _T("' '");
	if (m_bInsurance)
	{
		if (!szObjects.IsEmpty())
			szObjects += _T(",");
		szObjects += _T("'I', 'C'");
	}
	if (m_bDiscount)
	{
		if (!szObjects.IsEmpty())
			szObjects += _T(",");
		szObjects += _T("'P', 'D'");
	}
	if (m_bService)
	{
		if (!szObjects.IsEmpty())
			szObjects += _T(",");
		szObjects += _T("'S'");
	}

	szWhere.Empty();
	if (!m_szStatusKey.IsEmpty())
	{
		if(m_szStatusKey ==_T("A"))
		{
			szWhereDate.Format(_T(" AND htr_admitdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
			szWhere.Format(_T(" AND htr_status='I' AND htr_suggestion is null "));
		}
		else if(m_szStatusKey ==_T("D"))
		{
			szWhereDate.Format(_T(" AND htr_dischargedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
			szWhere.Format(_T(" and htr_suggestion='D' and htr_status='T' "));
		}
		else if(m_szStatusKey ==_T("T"))
		{
			szWhereDate.Format(_T(" AND htr_dischargedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
			szWhere.Format(_T(" and htr_suggestion='T' and htr_status='T' "));
		}
		else if (m_szStatusKey==_T("M"))
		{
			szWhereDate.Format(_T(" AND htr_dischargedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
			szWhere.Format(_T(" and htr_suggestion='M' "));
		}
		else if (m_szStatusKey==_T("F"))
		{
			szWhereDate.Format(_T(" AND htr_dischargedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
			szWhere.Format(_T(" and htr_suggestion='F' and htr_status='T' "));
		}
	}
	else
	{
		szWhereDate.Format(_T(" AND (htr_dischargedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
			_T(" OR htr_admitdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s'))"), m_szFromDate, m_szToDate, m_szFromDate, m_szToDate);
		szWhere.Format(_T(" and hcr_status IN ('I', 'T') "));
	}
	if (!szWhereDate.IsEmpty())
		szWhere.AppendFormat(_T("%s"), szWhereDate);
	for (int i=0 ; i<= m_wndFloor.GetCount(); i++)
	{
		if(m_wndFloor.GetCheck(i))
		{
			m_wndFloor.SetCurSel(i);
		 	if(!szFloor.IsEmpty())
				szFloor += _T(",");						
			szFloor.AppendFormat(_T("'%s'"), m_wndFloor.GetCurrent(0));			
			if(!szFloorName.IsEmpty())
				szFloorName += _T(", ");						
			szFloorName.AppendFormat(_T("%s"), m_wndFloor.GetCurrent(1));
		}
	}	

	if(!m_szFloorKey.IsEmpty())	
		szDept = szFloor;
	else
		szDept.Format(_T("'%s'"), pMF->m_szDept);
	if (!szDept.IsEmpty())
		szWhere.AppendFormat(_T(" AND htr_deptid IN (%s)"), szDept);
	if (m_bNotPatientFloor)
	{
		szWhere.AppendFormat(_T(" and htr_suggestion<>'F' "));	
	}

	if (m_bNoPatientSurgeryToReport)
	{
		if (m_bNoPatientDialysisToReport)
		{
			szWhere.AppendFormat(_T(" and NVL(htr_tdeptid, 'X') NOT IN ('PTGMHS', 'PTNT')"));
			szWhereFromDate.Format(_T(" WHERE NVL(FromDept, 'X') NOT IN ('PTGMHS', 'PTNT')"));
		}
		else
		{
			szWhere.AppendFormat(_T(" and NVL(htr_tdeptid, 'X')<>'PTGMHS' "));
			szWhereFromDate.Format(_T(" WHERE NVL(FromDept, 'X') <>'PTGMHS'"));
		}
	}
	
	if (m_bNoPatientDialysisToReport && !m_bNoPatientSurgeryToReport)
	{
		szWhere.AppendFormat(_T(" and NVL(htr_tdeptid, 'X')<>'PTNT'"));
		szWhereFromDate.Format(_T(" WHERE NVL(FromDept, 'X') <>'PTNT'"));
	}

	if (m_bNoPatientTerminatedToReport)
		szWhere.AppendFormat(_T(" and htr_status <>'T' "));
	else
		szWhere.AppendFormat(_T(" and htr_status IN ('I','T') "));
	if (!szObjects.IsEmpty())
		szWhere.AppendFormat(_T(" AND ho_type IN (%s)"), szObjects);
	szSQL.Format(_T(" SELECT * ") \
				_T(" FROM   (SELECT recordno, docno, idx, suggestion, patientname, sex, age, address, ") \
				_T("                occupation, cardno, transferplace, ") \
				_T("                transferdiagnostic, icd, diagnostic, maindisease, mainicd, ") \
				_T("                admitdate, deptname, tre12thang, tre15tuoi, tre6tuoi, ") \
				_T("                tre30ngay, dischargedept, dischargefloor, dischargedate, ") \
				_T("                ToDept, transhospitall, tuvongsau24, tuvongtruoc24, ") \
				_T("                khoi, dogiam, khongthaydoi, nang, xinve, tronvien, ") \
				_T("                treatqty, ngaydt, ") \
				_T("                (SELECT htr_deptid ") \
				_T("                 FROM   hms_treatment_record ") \
				_T("                 WHERE  htr_docno = docno ") \
				_T("                    AND htr_idx = idx - 1 ") \
				_T("                    AND htr_suggestion = 'M') AS FromDept ") \
				_T("         FROM   (SELECT    hcr_recordno                                    AS recordno, ") \
				_T("                           hd_docno                                        AS docno, ") \
				_T("                           htr_idx                                         AS idx, ") \
				_T("                           htr_suggestion                                  AS suggestion, ") \
				_T("                           Trim(hp_surname ") \
				_T("                                ||' ' ") \
				_T("                                ||hp_midname ") \
				_T("                                ||' ' ") \
				_T("                                ||hp_firstname)                            AS patientname, ") \
				_T("                           hp_sex                                          AS sex, ") \
				_T("                           Hms_getage(Trunc(hd_admitdate), hp_birthdate)   AS age, ") \
				_T("                           Hms_getaddress(hp_provid, hp_distid, hp_villid) AS address, ") \
				_T("                           (SELECT ss_desc ") \
				_T("                            FROM   sys_sel ") \
				_T("                            WHERE  ss_id = 'sys_occupation' ") \
				_T("                               AND Cast(ss_code AS INT) = hp_occupation)   AS occupation, ") \
				_T("                           hd_cardno                                       AS cardno, ") \
				_T("                           hd_transplace                                   AS transferplace, ") \
				_T("                           hd_transdiagn                                   AS transferdiagnostic, ") \
				_T("                           hd_icd                                          AS icd, ") \
				_T("                           hd_diagnostic                                   AS diagnostic, ") \
				_T("                           Trim(hcr_maindisease)                           AS maindisease, ") \
				_T("                           Trim(hcr_mainicd)                               AS mainicd, ") \
				_T("                           ( htr_admitdate )                               AS admitdate, ") \
				_T("                           sd_name                                         AS deptname, ") \
				_T("                           CASE WHEN Trunc(hcr_admitdate) - Trunc(hp_birthdate) <= 365 THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS tre12thang, ") \
				_T("                           CASE WHEN Trunc(hcr_admitdate) - Trunc(hp_birthdate) <= 365 * 15 THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS tre15tuoi, ") \
				_T("                           CASE WHEN Trunc(hcr_admitdate) - Trunc(hp_birthdate) <= 365 * 6 THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS tre6tuoi, ") \
				_T("                           CASE WHEN Trunc(hcr_admitdate) - Trunc(hp_birthdate) <= 30 THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS tre30ngay, ") \
				_T("                           CASE WHEN htr_idx > 1 ") \
				_T("                                AND htr_suggestion = 'M' THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS dischargedept, ") \
				_T("                           CASE WHEN htr_idx > 1 ") \
				_T("                                AND htr_suggestion = 'F' THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS dischargefloor, ") \
				_T("                           CASE WHEN htr_suggestion = 'D' THEN ( hcr_dischargedate ) ") \
				_T("                           ELSE NULL ") \
				_T("                           END                                             AS dischargedate, ") \
				_T("                           CASE WHEN htr_suggestion = 'M' THEN htr_tdeptid ") \
				_T("                           ELSE NULL ") \
				_T("                           END                                             AS ToDept, ") \
				_T("                           CASE WHEN htr_suggestion = 'T' THEN ( hcr_dischargedate ) ") \
				_T("                           ELSE NULL ") \
				_T("                           END                                             AS transhospitall, ") \
				_T("                           CASE WHEN htr_suggestion = 'D' ") \
				_T("                                AND hcr_result = '5' THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS tuvongsau24, ") \
				_T("                           CASE WHEN htr_suggestion = 'D' ") \
				_T("                                AND hcr_result = '6' THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS tuvongtruoc24, ") \
				_T("                           CASE WHEN htr_suggestion = 'D' ") \
				_T("                                AND hcr_result = '1' THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS khoi, ") \
				_T("                           CASE WHEN htr_suggestion = 'D' ") \
				_T("                                AND hcr_result = '2' THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS dogiam, ") \
				_T("                           CASE WHEN htr_suggestion = 'D' ") \
				_T("                                AND hcr_result = '3' THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS khongthaydoi, ") \
				_T("                           CASE WHEN htr_suggestion = 'D' ") \
				_T("                                AND hcr_result = '4' THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS nang, ") \
				_T("                           CASE WHEN htr_suggestion = 'D' ") \
				_T("                                AND hcr_result = '7' THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS xinve, ") \
				_T("                           CASE WHEN htr_suggestion = 'D' ") \
				_T("                                AND hcr_result = '8' THEN 'X' ") \
				_T("                           ELSE '' ") \
				_T("                           END                                             AS tronvien, ") \
				_T("                           CASE WHEN htr_status = 'T' THEN (SELECT SUM(hb_treatqty) ") \
				_T("                                                       FROM   hms_bed ") \
				_T("                                                       WHERE  hb_docno = hcr_docno) ") \
				_T("                           ELSE 0 ") \
				_T("                           END                                             AS treatqty, ") \
				_T("                           CASE WHEN htr_status = 'I' THEN ( trunc(cast_string2timestamp('%s')) - Trunc(hcr_admitdate) ) + 1 ") \
				_T("                           ELSE 0 ") \
				_T("                           END                                             AS ngaydt ") \
				_T("                 FROM      hms_patient ") \
				_T("                 LEFT JOIN hms_doc ON( hd_patientno = hp_patientno ) ") \
				_T("                 LEFT JOIN hms_treatment_record ON( htr_docno = hd_docno ) ") \
				_T("                 LEFT JOIN hms_clinical_record ON( hcr_docno = htr_docno ") \
				_T("                                                   AND htr_idx = hcr_refidx ) ") \
				_T("                 LEFT JOIN hms_object ON( hd_object = ho_id ) ") \
				_T("                 LEFT JOIN sys_dept ON( sd_id = htr_deptid ) ") \
				_T("                 WHERE     1=1 %s") \
				_T("                ) tblReport) tblMain ") \
				_T("				%s") \
				_T(" ORDER  BY docno, ") \
				_T("           idx "), m_szToDate, szWhere, szWhereFromDate) ;
_fmsg(_T("%s"), szSQL);
	return szSQL;
}