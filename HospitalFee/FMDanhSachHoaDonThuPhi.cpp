#include "stdafx.h"
#include "FMDanhSachHoaDonThuPhi.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnYearSetfocus();} */
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CRptDanhSachHoaDonThuPhi *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptDanhSachHoaDonThuPhi* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CRptDanhSachHoaDonThuPhi *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CRptDanhSachHoaDonThuPhi *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CRptDanhSachHoaDonThuPhi *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptDanhSachHoaDonThuPhi* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CRptDanhSachHoaDonThuPhi *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CRptDanhSachHoaDonThuPhi *)pWnd)->OnClerkAddNew();
}*/
static void _OnOtherFeeSelectFnc(CWnd *pWnd){
	 ((CRptDanhSachHoaDonThuPhi*)pWnd)->OnOtherFeeSelect();
}
static void _OnExamFeeSelectFnc(CWnd *pWnd){
	 ((CRptDanhSachHoaDonThuPhi*)pWnd)->OnExamFeeSelect();
}
static void _OnDrugFeeSelectFnc(CWnd *pWnd){
	 ((CRptDanhSachHoaDonThuPhi*)pWnd)->OnDrugFeeSelect();
}
static void _OnMaterialFeeSelectFnc(CWnd *pWnd){
	 ((CRptDanhSachHoaDonThuPhi*)pWnd)->OnMaterialFeeSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CRptDanhSachHoaDonThuPhi *pVw = (CRptDanhSachHoaDonThuPhi *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CRptDanhSachHoaDonThuPhi *pVw = (CRptDanhSachHoaDonThuPhi *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CRptDanhSachHoaDonThuPhi *pVw = (CRptDanhSachHoaDonThuPhi *)pWnd;
	pVw->OnCloseSelect();
}
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CRptDanhSachHoaDonThuPhi*)pWnd)->OnLockedSelect();
}
static int _OnAddRptDanhSachHoaDonThuPhiFnc(CWnd *pWnd){
	 return ((CRptDanhSachHoaDonThuPhi*)pWnd)->OnAddRptDanhSachHoaDonThuPhi();
} 
static int _OnEditRptDanhSachHoaDonThuPhiFnc(CWnd *pWnd){
	 return ((CRptDanhSachHoaDonThuPhi*)pWnd)->OnEditRptDanhSachHoaDonThuPhi();
} 
static int _OnDeleteRptDanhSachHoaDonThuPhiFnc(CWnd *pWnd){
	 return ((CRptDanhSachHoaDonThuPhi*)pWnd)->OnDeleteRptDanhSachHoaDonThuPhi();
} 
static int _OnSaveRptDanhSachHoaDonThuPhiFnc(CWnd *pWnd){
	 return ((CRptDanhSachHoaDonThuPhi*)pWnd)->OnSaveRptDanhSachHoaDonThuPhi();
} 
static int _OnCancelRptDanhSachHoaDonThuPhiFnc(CWnd *pWnd){
	 return ((CRptDanhSachHoaDonThuPhi*)pWnd)->OnCancelRptDanhSachHoaDonThuPhi();
} 
CRptDanhSachHoaDonThuPhi::CRptDanhSachHoaDonThuPhi(CWnd *pParent)
{
	m_nDlgWidth = 585;
	m_nDlgHeight = 220;
	m_szTitle = _T("\x44\x61nh s\xE1\x63h h\xF3\x61 \x111\x1A1n thu ph\xED");
	SetDefaultValues();
}
CRptDanhSachHoaDonThuPhi::~CRptDanhSachHoaDonThuPhi(){
}
void CRptDanhSachHoaDonThuPhi::OnCreateComponents()
{
	/*m_wndReportCondition.Create(this, _T("Charged Invoice List"), 5, 5, 575, 150);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 285, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 290, 30, 370, 55);
	m_wndReportPeriod.Create(this,375, 30, 570, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 285, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 290, 60, 370, 85);
	m_wndToDate.Create(this,375, 60, 570, 85); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 90, 115);
	m_wndClerk.Create(this,95, 90, 570, 115); 
	m_wndExamFee.Create(this, _T("Exam Fee"), 320, 120, 400, 145);
	m_wndDrugFee.Create(this, _T("Drug Fee"), 405, 120, 485, 145);
	m_wndOtherFee.Create(this, _T("Other Fee"), 490, 120, 570, 145);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 385, 155, 495, 180);
	m_wndExport.Create(this, _T("&Export"), 500, 155, 575, 180);*/

	m_wndReportCondition.Create(this, _T("Charged Invoice List"), 5, 5, 430, 180);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 90, 115);
	m_wndClerk.Create(this,95, 90, 425, 115); 
	m_wndOtherFee.Create(this, _T("Other Fee"), 325, 120, 425, 145);
	m_wndExamFee.Create(this, _T("Tooth Fee"), 10, 120, 120, 145);
	m_wndDrugFee.Create(this, _T("Drug Fee"), 125, 120, 205, 145);
	m_wndMaterialFee.Create(this, _T("Material Fee"), 210, 120, 320, 145);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 205, 185, 315, 210);
	m_wndExport.Create(this, _T("&Export"), 320, 185, 430, 210);
	m_wndLocked.Create(this, _T("Locked"), 325, 150, 425, 175);
}
void CRptDanhSachHoaDonThuPhi::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(false);
	m_wndReportPeriod.SetCheckValue(false);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(false);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(false);
	m_wndClerk.SetCheckValue(false);
	m_wndClerk.LimitText(65);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 340);
}
void CRptDanhSachHoaDonThuPhi::OnSetWindowEvents()
{
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
	m_wndOtherFee.SetEvent(WE_CLICK, _OnOtherFeeSelectFnc);
	m_wndExamFee.SetEvent(WE_CLICK, _OnExamFeeSelectFnc);
	m_wndDrugFee.SetEvent(WE_CLICK, _OnDrugFeeSelectFnc);
	m_wndMaterialFee.SetEvent(WE_CLICK, _OnMaterialFeeSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	SetMode(VM_EDIT);
}
void CRptDanhSachHoaDonThuPhi::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndOtherFee.GetDlgCtrlID(), m_bOtherFee);
	DDX_Check(pDX, m_wndExamFee.GetDlgCtrlID(), m_bExamFee);
	DDX_Check(pDX, m_wndDrugFee.GetDlgCtrlID(), m_bDrugFee);
	DDX_Check(pDX, m_wndMaterialFee.GetDlgCtrlID(), m_bMaterialFee);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);
}
void CRptDanhSachHoaDonThuPhi::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CRptDanhSachHoaDonThuPhi::GetScreenToData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CRptDanhSachHoaDonThuPhi::SetDefaultValues()
{
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bOtherFee = FALSE;
	m_bExamFee = FALSE;
	m_bDrugFee = FALSE;
	m_bMaterialFee=FALSE;
	m_bLocked = FALSE;
}
int CRptDanhSachHoaDonThuPhi::SetMode(int nMode)
{
 		int nOldMode = GetMode(); 
 		CGuiView::SetMode(nMode); 
 		CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 		CString szSQL; 
 		CRecord rs(&pMF->m_db); 
  		switch(nMode)
		{ 
 		case VM_ADD:
 			EnableControls(TRUE);
 			EnableButtons(TRUE, 3, 4, -1);
 			SetDefaultValues();
 			break;
 		case VM_EDIT:
 			EnableControls(TRUE); 
 			EnableButtons(TRUE, 0, 1, -1);
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
/*void CRptDanhSachHoaDonThuPhi::OnYearChange(){
	
} */
/*void CRptDanhSachHoaDonThuPhi::OnYearSetfocus(){
	
} */
/*void CRptDanhSachHoaDonThuPhi::OnYearKillfocus(){
	
} */
int CRptDanhSachHoaDonThuPhi::OnYearCheckValue()
{
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
void CRptDanhSachHoaDonThuPhi::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CRptDanhSachHoaDonThuPhi::OnReportPeriodSelendok()
{
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
/*void CRptDanhSachHoaDonThuPhi::OnReportPeriodSetfocus(){
	
}*/
/*void CRptDanhSachHoaDonThuPhi::OnReportPeriodKillfocus(){
	
}*/
long CRptDanhSachHoaDonThuPhi::OnReportPeriodLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
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
/*void CRptDanhSachHoaDonThuPhi::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CRptDanhSachHoaDonThuPhi::OnFromDateChange(){
	
} */
/*void CRptDanhSachHoaDonThuPhi::OnFromDateSetfocus(){
	
} */
/*void CRptDanhSachHoaDonThuPhi::OnFromDateKillfocus(){
	
} */
int CRptDanhSachHoaDonThuPhi::OnFromDateCheckValue(){
	return 0;
} 
/*void CRptDanhSachHoaDonThuPhi::OnToDateChange(){
	
} */
/*void CRptDanhSachHoaDonThuPhi::OnToDateSetfocus(){
	
} */
/*void CRptDanhSachHoaDonThuPhi::OnToDateKillfocus(){
	
} */
int CRptDanhSachHoaDonThuPhi::OnToDateCheckValue(){
	return 0;
} 
void CRptDanhSachHoaDonThuPhi::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CRptDanhSachHoaDonThuPhi::OnClerkSelendok(){
	 
}
/*void CRptDanhSachHoaDonThuPhi::OnClerkSetfocus(){
	
}*/
/*void CRptDanhSachHoaDonThuPhi::OnClerkKillfocus(){
	
}*/
long CRptDanhSachHoaDonThuPhi::OnClerkLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	}
	m_wndClerk.DeleteAllItems(); 
	szSQL.Format(_T("SELECT su_userid as userid, su_name as name FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("userid")),
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CRptDanhSachHoaDonThuPhi::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CRptDanhSachHoaDonThuPhi::OnOtherFeeSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bDrugFee = FALSE;
	m_bExamFee = FALSE;
	m_bMaterialFee = FALSE;
	UpdateData(FALSE);
}
void CRptDanhSachHoaDonThuPhi::OnExamFeeSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bDrugFee = FALSE;
	m_bOtherFee = FALSE;
	m_bMaterialFee = FALSE;
	UpdateData(FALSE);
}
void CRptDanhSachHoaDonThuPhi::OnDrugFeeSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bOtherFee = FALSE;
	m_bExamFee = FALSE;
	m_bMaterialFee = FALSE;
	UpdateData(FALSE);
}
void CRptDanhSachHoaDonThuPhi::OnMaterialFeeSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bDrugFee = FALSE;
	m_bExamFee = FALSE;
	m_bOtherFee = FALSE;
	UpdateData(FALSE);
}
void CRptDanhSachHoaDonThuPhi::OnLockedSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CRptDanhSachHoaDonThuPhi::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString tmpStr, szSQL;
	double nAmount = 0, nTotal = 0;
	int nIdx = 0;
	if(!rpt.Init(_T("Reports/HMS/HF_DANHSACHHOADONTHUPHI.RPT")))
		return;
	BeginWaitCursor();
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	CReportSection* rptDetail;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("invoiceno"),tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		tmpStr = CDate::Convert(rs.GetValue(_T("receiptdate")), yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rs.GetValue(_T("yob"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rs.GetValue(_T("sex"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rs.GetValue(_T("amount"), nAmount);
		nTotal += nAmount;
		FormatCurrency(nAmount, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.MoveNext();	
	}	
	FormatCurrency(nTotal, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("s8"), tmpStr);
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();
} 
void CRptDanhSachHoaDonThuPhi::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;

	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0;
	double nAmount = 0;
	long double nTotal = 0;
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 10);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 25);
	xls.SetColumnWidth(4, 8);
	xls.SetColumnWidth(5, 5);
	xls.SetColumnWidth(6, 15);

	xls.SetRowHeight(4, 40);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);

	xls.SetCellMergedColumns(nCol + 4, nRow, 4);
	xls.SetCellMergedColumns(nCol + 4, nRow + 1, 4);

	xls.SetCellMergedColumns(nCol, nRow + 2, 7);
	xls.SetCellMergedColumns(nCol, nRow + 3, 7);


	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 4, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 4, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol, nRow + 2, _T("\x44\x41NH S\xC1\x43H H\xD3\x41 \x110\x1A0N THU PH\xCD"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("S\x1ED1 \x63h\x1EE9ng t\x1EEB"));
	arrCol.Add(_T("Ng\xE0y thu"));
	arrCol.Add(_T("S\x1ED1 h\x1ED3 s\x1A1"));
	arrCol.Add(_T("T\xEAn \x62\x1EC7nh nh\xE2n"));
	arrCol.Add(_T("N\x103m sinh"));
	arrCol.Add(_T("Gi\x1EDBi"));
	arrCol.Add(_T("S\x1ED1 ti\x1EC1n"));
	nRow = 4;

	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol + i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10); 
	}

	nRow = 5;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("invoiceno"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("receiptdate"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), FMT_DATE | FMT_WRAPING);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("yob"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("sex"), tmpStr);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("amount"), nAmount);
		tmpStr.Format(_T("%.2f"), nAmount);
		nTotal += nAmount;
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		nRow++;
		rs.MoveNext();
	}
	
	xls.SetCellMergedColumns(nCol, nRow, 6);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 10);
	tmpStr.Format(_T("%.2Lf"), nTotal);
	xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
	xls.Save(_T("Exports\\Danh Sach Hoa Don Thu Phi.xls"));
	
} 
void CRptDanhSachHoaDonThuPhi::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
	
} 
int CRptDanhSachHoaDonThuPhi::OnAddRptDanhSachHoaDonThuPhi(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CRptDanhSachHoaDonThuPhi::OnEditRptDanhSachHoaDonThuPhi(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0; 
}
int CRptDanhSachHoaDonThuPhi::OnDeleteRptDanhSachHoaDonThuPhi(){
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
 		OnCancelRptDanhSachHoaDonThuPhi(); 
 	}
	return 0;
}
int CRptDanhSachHoaDonThuPhi::OnSaveRptDanhSachHoaDonThuPhi()
{
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT)
 		return -1;
 	if(!IsValidateData())
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
 	CString szSQL;
 	if(GetMode() == VM_ADD)
	{
 		//szSQL = m_tblTbl.GetInsertSQL();
 	} 
 	else if(GetMode() == VM_EDIT)
	{ 
 		CString szWhere;
 		szWhere.Format(_T(" WHERE"));
 		//szSQL = m_tblTbl.GetUpdateSQL(_T("createdby,createddate")); 
 		szSQL += szWhere; 
 	}
	_fmsg(_T("%s"), szSQL); 
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret > 0) 
 	{ 
 		//OnRptDanhSachHoaDonThuPhiListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CRptDanhSachHoaDonThuPhi::OnCancelRptDanhSachHoaDonThuPhi(){
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
int CRptDanhSachHoaDonThuPhi::OnRptDanhSachHoaDonThuPhiListLoadData(){
	return 0;
}
CString CRptDanhSachHoaDonThuPhi::GetQueryString()
{
	CString szSQL, szWhere;
	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" and tbl1.hfe_staff='%s' "), m_szClerkKey);
	
	if (m_bDrugFee)
	{
		szWhere.AppendFormat(_T(" and tbl2.hfe_itemid in('F0000001') "));
	}

	if (m_bExamFee)
		szWhere.AppendFormat(_T(" and tbl2.hfe_itemid in('F0000010') "));

	if (m_bMaterialFee)
		szWhere.AppendFormat(_T(" and tbl2.hfe_itemid in('F0000003') "));

	if (m_bOtherFee)
		szWhere.AppendFormat(_T(" and tbl2.hfe_group in('F0000') and tbl2.hfe_itemid not in('F0000001','F0000003','F0000010') "));

	if (m_bLocked)
	{
		szWhere.AppendFormat(_T(" and tbl1.hfe_locked='Y' "));
		szWhere.AppendFormat(_T(" and tbl1.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}
	else
	{
		szWhere.AppendFormat(_T(" and tbl1.hfe_locked<>'Y' "));
		szWhere.AppendFormat(_T(" and tbl1.hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}

	szSQL.Format(_T(" SELECT hd_docno as docno,") \
		        _T("        invoiceno,") \
				_T("        receiptdate,") \
				_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("        extract(year from hp_birthdate) as yob,") \
				_T("        (select distinct ss_desc from sys_sel where ss_id='sys_sex' and ss_code=hp_sex) as sex,") \
				_T("        sum(amount) as amount,") \
				_T("        sum(discount) as inspaid,") \
				_T("        sum(patpaid) as patpaid,") \
				_T("        receiveby") \
				_T(" FROM") \
				_T(" (") \
				_T("   SELECT tbl1.hfe_staff as receiveby,") \
				_T("          tbl1.hfe_date as receiptdate,") \
				_T("          tbl2.hfe_deptid,") \
				_T("          tbl1.hfe_class,") \
				_T("          tbl1.hfe_docno as docno,") \
				_T("          tbl1.hfe_invoiceno as invoiceno,") \
				_T("          tbl2.hfe_group as groupid,") \
				_T("          tbl2.hfe_quantity as qty,") \
				_T("          tbl2.hfe_cost as amount,") \
				_T("          tbl2.hfe_discount as discount,") \
				_T("          tbl2.hfe_patpaid as patpaid,") \
				_T("          tbl2.hfe_patdebt as patdebt") \
				_T("    FROM hms_fee_invoice tbl1") \
				_T("    LEFT JOIN hms_fee tbl2 ON(tbl2.hfe_docno=tbl1.hfe_docno AND tbl2.hfe_invoiceno=tbl1.hfe_invoiceno)") \
				_T("    WHERE tbl2.hfe_class in('E','X') AND tbl1.hfe_status IN('P','R')") \
				_T("          and tbl2.hfe_cost > 0 %s") \
				_T(" ) Tbl") \
				_T(" LEFT JOIN hms_doc ON(docno=hd_docno)") \
				_T(" LEFT JOIN hms_patient ON(hd_patientno=hp_patientno)") \
				_T(" LEFT JOIN hms_object ON(ho_id=hd_object)") \
				_T(" GROUP BY hd_docno, invoiceno, hp_surname, hp_midname, hp_firstname, hp_birthdate, hp_sex, receiveby, receiptdate") \
				_T(" ORDER BY docno, pname"),
				szWhere);

	return szSQL;
}
