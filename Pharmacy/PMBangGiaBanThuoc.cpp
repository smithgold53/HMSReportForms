#include "stdafx.h"
#include "PMBangGiaBanThuoc.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BangGiaBanThuoc* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptPMS108_BangGiaBanThuoc* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CRptPMS108_BangGiaBanThuoc *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CRptPMS108_BangGiaBanThuoc *pVw = (CRptPMS108_BangGiaBanThuoc *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CRptPMS108_BangGiaBanThuoc *pVw = (CRptPMS108_BangGiaBanThuoc *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CRptPMS108_BangGiaBanThuoc *pVw = (CRptPMS108_BangGiaBanThuoc *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddRptPMS108_BangGiaBanThuocFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangGiaBanThuoc*)pWnd)->OnAddRptPMS108_BangGiaBanThuoc();
} 
static int _OnEditRptPMS108_BangGiaBanThuocFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangGiaBanThuoc*)pWnd)->OnEditRptPMS108_BangGiaBanThuoc();
} 
static int _OnDeleteRptPMS108_BangGiaBanThuocFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangGiaBanThuoc*)pWnd)->OnDeleteRptPMS108_BangGiaBanThuoc();
} 
static int _OnSaveRptPMS108_BangGiaBanThuocFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangGiaBanThuoc*)pWnd)->OnSaveRptPMS108_BangGiaBanThuoc();
} 
static int _OnCancelRptPMS108_BangGiaBanThuocFnc(CWnd *pWnd){
	 return ((CRptPMS108_BangGiaBanThuoc*)pWnd)->OnCancelRptPMS108_BangGiaBanThuoc();
} 
CRptPMS108_BangGiaBanThuoc::CRptPMS108_BangGiaBanThuoc(CWnd *pParent)
	{
	m_nDlgWidth = 420;
	m_nDlgHeight = 155;
	m_szTitle = _T("Price sell drugs");
	SetDefaultValues();
}
CRptPMS108_BangGiaBanThuoc::~CRptPMS108_BangGiaBanThuoc(){
}
void CRptPMS108_BangGiaBanThuoc::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 205, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 215, 30, 295, 55);
	m_wndReportPeriod.Create(this,300, 30, 410, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 205, 85); 
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 415, 119);
	m_wndToDateLabel.Create(this, _T("To Date"), 215, 60, 295, 85);
	m_wndToDate.Create(this,300, 60, 410, 85); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 90, 115);
	m_wndClerk.Create(this,95, 90, 410, 115); 
	m_wndPrint.Create(this, _T("&Print"), 249, 125, 329, 150);
	m_wndExport.Create(this, _T("&Export"), 335, 125, 415, 150);
}
void CRptPMS108_BangGiaBanThuoc::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 80);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);
}
void CRptPMS108_BangGiaBanThuoc::OnSetWindowEvents(){
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
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate =  pMF->GetSysDate()+ _T("00:00");	
	m_szToDate = pMF->GetSysDate() + _T("23:59");
	SetMode(VM_EDIT);
}
void CRptPMS108_BangGiaBanThuoc::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
}
void CRptPMS108_BangGiaBanThuoc::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CRptPMS108_BangGiaBanThuoc::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
}
void CRptPMS108_BangGiaBanThuoc::SetDefaultValues(){
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();

}
int CRptPMS108_BangGiaBanThuoc::SetMode(int nMode){
 		int nOldMode = GetMode(); 
 		CGuiView::SetMode(nMode); 
 		CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
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
/*void CRptPMS108_BangGiaBanThuoc::OnYearChange(){
	
} */
/*void CRptPMS108_BangGiaBanThuoc::OnYearSetfocus(){
	
} */
/*void CRptPMS108_BangGiaBanThuoc::OnYearKillfocus(){
	
} */
int CRptPMS108_BangGiaBanThuoc::OnYearCheckValue(){
	return 0;
} 
void CRptPMS108_BangGiaBanThuoc::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
} 
void CRptPMS108_BangGiaBanThuoc::OnReportPeriodSelendok(){
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
/*void CRptPMS108_BangGiaBanThuoc::OnReportPeriodSetfocus(){
	
}*/
/*void CRptPMS108_BangGiaBanThuoc::OnReportPeriodKillfocus(){
	
}*/
long CRptPMS108_BangGiaBanThuoc::OnReportPeriodLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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
/*void CRptPMS108_BangGiaBanThuoc::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CRptPMS108_BangGiaBanThuoc::OnFromDateChange(){
	
} */
/*void CRptPMS108_BangGiaBanThuoc::OnFromDateSetfocus(){
	
} */
/*void CRptPMS108_BangGiaBanThuoc::OnFromDateKillfocus(){
	
} */
int CRptPMS108_BangGiaBanThuoc::OnFromDateCheckValue(){
	return 0;
} 
/*void CRptPMS108_BangGiaBanThuoc::OnToDateChange(){
	
} */
/*void CRptPMS108_BangGiaBanThuoc::OnToDateSetfocus(){
	
} */
/*void CRptPMS108_BangGiaBanThuoc::OnToDateKillfocus(){
	
} */
int CRptPMS108_BangGiaBanThuoc::OnToDateCheckValue(){
	return 0;
} 
void CRptPMS108_BangGiaBanThuoc::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CRptPMS108_BangGiaBanThuoc::OnClerkSelendok(){
	 
}
/*void CRptPMS108_BangGiaBanThuoc::OnClerkSetfocus(){
	
}*/
/*void CRptPMS108_BangGiaBanThuoc::OnClerkKillfocus(){
	
}*/
long CRptPMS108_BangGiaBanThuoc::OnClerkLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	};
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
/*void CRptPMS108_BangGiaBanThuoc::OnClerkAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CRptPMS108_BangGiaBanThuoc::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	if(!rpt.Init(_T("Reports/HMS/PMS_BANGGIABANTHUOC.RPT")) )
	return ;
	UpdateData(true);
// viet cau lenh
	_fmsg(_T("%s"), szSQL);
	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	StringUpper(pMF->m_szHealthService, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_szHospitalName, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm,ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), szDate);

	long double nTotal1, nTotal2;
	nTotal1 = nTotal2 = 0;
	int nIndex = 1;
		CReportSection* rptDetail;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%ld"), nIndex ++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T(""),tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		nTotal1 += ToDouble(tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		
		rs.GetValue(_T(""), tmpStr);
		nTotal2 += ToDouble(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		
		rs.GetValue(_T(""), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.MoveNext();	
	}
			tmpStr.Format(_T("%.0f"), nTotal1);
			rpt.GetReportFooter()->SetValue(_T("s4"), tmpStr);	

			tmpStr.Format(_T("%.0f"), nTotal2);
			rpt.GetReportFooter()->SetValue(_T("s5"), tmpStr);	
	EndWaitCursor();
	rpt.PrintPreview();
} 
void CRptPMS108_BangGiaBanThuoc::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString tmpStr, szSQL, szWhere, szFromDateToDate;
	CRecord rs(&pMF->m_db);
	UpdateData(TRUE);
	
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);	
	xls.SetColumnWidth(0,8);
	xls.SetColumnWidth(1,15);
	xls.SetColumnWidth(2,12);
	xls.SetColumnWidth(3,12);
	xls.SetColumnWidth(4,12);
	xls.SetColumnWidth(5,12);
	
	long double nTotal1, nTotal2;
	double cost = 0; 
	nTotal1 = nTotal2 = 0;
	int nCol = 0;
	int nRow = 0;
	
	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 11, 0);
	xls.SetCellText(nCol, nRow + 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 11, 0);
	xls.SetCellMergedColumns(nCol, nRow + 3, 7);
	TranslateString(_T("PRICE SELL DRUGS"), tmpStr); 
	xls.SetCellMergedColumns(nCol, nRow + 4, 7);
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, true, 18, 0);

	TranslateString(_T("From Date"),tmpStr);
	szFromDateToDate.Format(_T("%s: %s"), tmpStr, CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	TranslateString(_T("To Date"),tmpStr);
	szFromDateToDate.AppendFormat(_T(" %s: %s"), tmpStr, CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));	
	//xls.SetCellText(nCol, nRow + 4, szFromDateToDate,FMT_TEXT,true,12);

	xls.SetCellText(nCol, nRow + 6, _T("STT"), FMT_TEXT | FMT_CENTER, true);	

	TranslateString(_T("Drug Name/Cnt"), tmpStr);	
	xls.SetCellText(nCol + 1, nRow + 6, tmpStr, FMT_TEXT | FMT_CENTER, true);	

	TranslateString(_T("Unit"), tmpStr);
	xls.SetCellText(nCol + 2, nRow + 6, tmpStr, FMT_TEXT | FMT_CENTER, true);	

	TranslateString(_T("Unit Price"), tmpStr);
	xls.SetCellText(nCol + 3, nRow + 6, tmpStr, FMT_TEXT | FMT_CENTER, true);	

	TranslateString(_T("Sell Price"), tmpStr);
	xls.SetCellText(nCol + 4, nRow + 6, tmpStr, FMT_TEXT | FMT_CENTER, true);	

	TranslateString(_T("Percent"), tmpStr);
	xls.SetCellText(nCol + 5, nRow + 6, tmpStr, FMT_TEXT |FMT_CENTER, true);

	TranslateString(_T("Company Distribution"), tmpStr);
	xls.SetCellText(nCol + 6, nRow + 6, tmpStr, FMT_TEXT | FMT_CENTER, true);
// viet lai cau lenh
	szSQL = GetQueryString();
		BeginWaitCursor();
		rs.ExecSQL(szSQL);

		int nIndex = 1;
		while (!rs.IsEOF())
		{
			
			tmpStr.Format(_T("%d"), nIndex++);
			xls.SetCellText(nCol, nRow + 6, tmpStr, FMT_NUMBER1);
// ten thuoc - ham luong
			rs.GetValue(_T(""), tmpStr);
			xls.SetCellText(nCol + 1, nRow + 6, tmpStr, FMT_TEXT);
// Don vi
			rs.GetValue(_T(""), tmpStr);
			xls.SetCellText(nCol + 2, nRow + 6, tmpStr, FMT_TEXT | FMT_CENTER);		
// Gia nhap
			rs.GetValue(_T(""), tmpStr);
			tmpStr.Format(_T("%.2f"), cost);
			nTotal1 += cost;
			xls.SetCellText(nCol + 3, nRow + 6, tmpStr, FMT_NUMBER1);
// Gia ban
			rs.GetValue(_T(""), cost);			
			tmpStr.Format(_T("%.2f"), cost);
			nTotal2 += cost;
			xls.SetCellText(nCol + 4, nRow + 6, tmpStr, FMT_NUMBER2);
//Ty le		
			rs.GetValue(_T(""), tmpStr);
			xls.SetCellText(nCol + 5, nRow + 6, tmpStr, FMT_TEXT);
// cong ty phan phoi.
			rs.GetValue(_T(""), tmpStr);
			xls.SetCellText(nCol + 6, nRow + 6, tmpStr, FMT_TEXT);
			
			nRow ++;
			rs.MoveNext();
		}
		EndWaitCursor();
		xls.Save(_T("Exports\\Bang gia ban thuoc.xls"));
} 
void CRptPMS108_BangGiaBanThuoc::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CRptPMS108_BangGiaBanThuoc::OnAddRptPMS108_BangGiaBanThuoc(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CRptPMS108_BangGiaBanThuoc::OnEditRptPMS108_BangGiaBanThuoc(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CRptPMS108_BangGiaBanThuoc::OnDeleteRptPMS108_BangGiaBanThuoc(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
 	CString szSQL; 
 	if(ShowMessage(1, MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2) == IDNO) 
 		return -1; 
 	szSQL.Format(_T("DELETE FROM  WHERE  AND") ); 
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret >= 0){ 
 		SetMode(VM_NONE); 
 		OnCancelRptPMS108_BangGiaBanThuoc(); 
 	}
	return 0;
}
int CRptPMS108_BangGiaBanThuoc::OnSaveRptPMS108_BangGiaBanThuoc(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT) 
 		return -1; 
 	if(!IsValidateData()) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
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
 		//OnRptPMS108_BangGiaBanThuocListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CRptPMS108_BangGiaBanThuoc::OnCancelRptPMS108_BangGiaBanThuoc(){
 	if(GetMode() == VM_EDIT) 
 	{ 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 		SetMode(VM_NONE); 
 	} 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
 	return 0;
} 
int CRptPMS108_BangGiaBanThuoc::OnRptPMS108_BangGiaBanThuocListLoadData(){
	return 0;
}
CString CRptPMS108_BangGiaBanThuoc::GetQueryString(){
	CString szSQL, szWhere;
	szSQL.Format(_T("SELECT  pmpi_invoiceno as invoiceno,") \
				_T(" 		pmpi_date as invoicedate,") \
				_T(" 		(select distinct pms_name from pms_suppliers where pms_id=pmpi_suppid) as supplier,") \
				_T(" 		(select distinct pmsl_name from pms_stocklist where pmsl_stockid=pmpi_stockid) as stockname,") \
				_T(" 		pmpi_tax1  as vat,") \
				_T(" 		pmpi_subtotal as subamount,") \
				_T(" 		pmpi_tax1amt as vatamount,") \
				_T(" 		pmpi_total as amount, ") \
				_T(" 		pmpi_payment as payment") \
				_T(" FROM pms_purchaseinvoice ") \
				_T(" WHERE pmpi_status <>'O'") \
				_T(" AND (pmpi_impdate) between cast_string2timestamp('%s') and cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	return szSQL;
}