#include "stdafx.h"
#include "EMDanhsachBNVaoVien.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CRptDanhsachBNVaoVien *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptDanhsachBNVaoVien* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CRptDanhsachBNVaoVien *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CRptDanhsachBNVaoVien *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CRptDanhsachBNVaoVien *)pWnd)->OnToDateCheckValue();
} 
static void _OnDepartmentSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptDanhsachBNVaoVien* )pWnd)->OnDepartmentSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDepartmentSelendokFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnDepartmentSelendok();
}
/*static void _OnDepartmentSetfocusFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnDepartmentKillfocus();
}*/
/*static void _OnDepartmentKillfocusFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnDepartmentKillfocus();
}*/
static long _OnDepartmentLoadDataFnc(CWnd *pWnd){
	return ((CRptDanhsachBNVaoVien *)pWnd)->OnDepartmentLoadData();
}
/*static void _OnDepartmentAddNewFnc(CWnd *pWnd){
	((CRptDanhsachBNVaoVien *)pWnd)->OnDepartmentAddNew();
}*/
static void _OnMinitarySelectFnc(CWnd *pWnd){
	  ((CRptDanhsachBNVaoVien*)pWnd)->OnMinitarySelect();
}
static void _OnRetireSelectFnc(CWnd *pWnd){
	  ((CRptDanhsachBNVaoVien*)pWnd)->OnRetireSelect();
}
static void _OnPrintSelectFnc(CWnd *pWnd){
	CRptDanhsachBNVaoVien *pVw = (CRptDanhsachBNVaoVien *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CRptDanhsachBNVaoVien *pVw = (CRptDanhsachBNVaoVien *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CRptDanhsachBNVaoVien *pVw = (CRptDanhsachBNVaoVien *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddRptDanhsachBNVaoVienFnc(CWnd *pWnd){
	 return ((CRptDanhsachBNVaoVien*)pWnd)->OnAddRptDanhsachBNVaoVien();
} 
static int _OnEditRptDanhsachBNVaoVienFnc(CWnd *pWnd){
	 return ((CRptDanhsachBNVaoVien*)pWnd)->OnEditRptDanhsachBNVaoVien();
} 
static int _OnDeleteRptDanhsachBNVaoVienFnc(CWnd *pWnd){
	 return ((CRptDanhsachBNVaoVien*)pWnd)->OnDeleteRptDanhsachBNVaoVien();
} 
static int _OnSaveRptDanhsachBNVaoVienFnc(CWnd *pWnd){
	 return ((CRptDanhsachBNVaoVien*)pWnd)->OnSaveRptDanhsachBNVaoVien();
} 
static int _OnCancelRptDanhsachBNVaoVienFnc(CWnd *pWnd){
	 return ((CRptDanhsachBNVaoVien*)pWnd)->OnCancelRptDanhsachBNVaoVien();
} 
CRptDanhsachBNVaoVien::CRptDanhsachBNVaoVien(CWnd *pParent)
{

	m_nDlgWidth = 435;
	m_nDlgHeight = 185;
	SetDefaultValues();
}
CRptDanhsachBNVaoVien::~CRptDanhsachBNVaoVien(){
}
void CRptDanhsachBNVaoVien::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 210, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 420, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 210, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 425, 145);
	m_wndToDate.Create(this,305, 60, 420, 85); 
	m_wndDepartmentLabel.Create(this, _T("Department"), 10, 90, 90, 115);
	m_wndDepartment.Create(this,95, 91, 420, 116); 
	m_wndMinitary.Create(this, _T("Minitary"), 257, 120, 339, 145);
	m_wndRetire.Create(this, _T("Retire"), 344, 120, 420, 145);
	m_wndPrint.Create(this, _T("&Print"), 262, 155, 342, 180);
	m_wndExport.Create(this, _T("&Export"), 347, 155, 427, 180);
}
void CRptDanhsachBNVaoVien::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
	m_wndDepartment.SetCheckValue(true);
	m_wndDepartment.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

	m_wndDepartment.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndDepartment.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);
	
}
void CRptDanhsachBNVaoVien::OnSetWindowEvents(){
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
	m_wndDepartment.SetEvent(WE_SELENDOK, _OnDepartmentSelendokFnc);
	//m_wndDepartment.SetEvent(WE_SETFOCUS, _OnDepartmentSetfocusFnc);
	//m_wndDepartment.SetEvent(WE_KILLFOCUS, _OnDepartmentKillfocusFnc);
	m_wndDepartment.SetEvent(WE_SELCHANGE, _OnDepartmentSelectChangeFnc);
	m_wndDepartment.SetEvent(WE_LOADDATA, _OnDepartmentLoadDataFnc);
	//m_wndDepartment.SetEvent(WE_ADDNEW, _OnDepartmentAddNewFnc);
	m_wndMinitary.SetEvent(WE_CLICK, _OnMinitarySelectFnc);
	m_wndRetire.SetEvent(WE_CLICK, _OnRetireSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDateTime();
	SetMode(VM_EDIT);
}
void CRptDanhsachBNVaoVien::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDepartment.GetDlgCtrlID(), m_szDepartmentKey);
	DDX_Radio(pDX, m_wndMinitary.GetDlgCtrlID(), m_nMinitary);
	DDX_Radio(pDX, m_wndRetire.GetDlgCtrlID(), m_nRetire);

}
void CRptDanhsachBNVaoVien::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CRptDanhsachBNVaoVien::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CRptDanhsachBNVaoVien::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDepartmentKey.Empty();
	m_nMinitary=0;
	m_nRetire=0;

}
int CRptDanhsachBNVaoVien::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0,1,2,3,-1); 
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
/*void CRptDanhsachBNVaoVien::OnYearChange(){
	
} */
/*void CRptDanhsachBNVaoVien::OnYearSetfocus(){
	
} */
/*void CRptDanhsachBNVaoVien::OnYearKillfocus(){
	
} */
int CRptDanhsachBNVaoVien::OnYearCheckValue(){
	return 0;
} 
void CRptDanhsachBNVaoVien::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
} 
void CRptDanhsachBNVaoVien::OnReportPeriodSelendok(){
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
/*void CRptDanhsachBNVaoVien::OnReportPeriodSetfocus(){
	
}*/
/*void CRptDanhsachBNVaoVien::OnReportPeriodKillfocus(){
	
}*/
long CRptDanhsachBNVaoVien::OnReportPeriodLoadData(){
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
/*void CRptDanhsachBNVaoVien::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CRptDanhsachBNVaoVien::OnFromDateChange(){
	
} */
/*void CRptDanhsachBNVaoVien::OnFromDateSetfocus(){
	
} */
/*void CRptDanhsachBNVaoVien::OnFromDateKillfocus(){
	
} */
int CRptDanhsachBNVaoVien::OnFromDateCheckValue(){
	return 0;
} 
/*void CRptDanhsachBNVaoVien::OnToDateChange(){
	
} */
/*void CRptDanhsachBNVaoVien::OnToDateSetfocus(){
	
} */
/*void CRptDanhsachBNVaoVien::OnToDateKillfocus(){
	
} */
int CRptDanhsachBNVaoVien::OnToDateCheckValue(){
	return 0;
} 
void CRptDanhsachBNVaoVien::OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CRptDanhsachBNVaoVien::OnDepartmentSelendok(){
	 
}
/*void CRptDanhsachBNVaoVien::OnDepartmentSetfocus(){
	
}*/
/*void CRptDanhsachBNVaoVien::OnDepartmentKillfocus(){
	
}*/
long CRptDanhsachBNVaoVien::OnDepartmentLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDepartment.IsSearchKey() && !m_szDepartmentKey.IsEmpty()){
		szWhere.Format(_T(" AND sd_id = '%s'"), m_szDepartmentKey);
	};
	szSQL.Format(_T("select sd_id as id, sd_name as name from sys_dept where sd_type = 'DT' %s order by id"), szWhere);
	m_wndDepartment.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDepartment.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CRptDanhsachBNVaoVien::OnDepartmentAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CRptDanhsachBNVaoVien::OnMinitarySelect(){
	
}
void CRptDanhsachBNVaoVien::OnRetireSelect(){
	
}
void CRptDanhsachBNVaoVien::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	szSQL = GetQueryString();
	if (!rpt.Init(_T("Reports\\HMS\\HE_DANHSACHBENHNHANVAOVIEN.RPT")))
		return;
	int ret = rs.ExecSQL(szSQL);
	if (ret <=0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection *rptDetail = NULL;
	int nIdx=0;
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);
		//sovaovien
		rs.GetValue(_T("inwardno"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		//tenbn
		rs.GetValue(_T("tenbenhnhan"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		//namsinh
		rs.GetValue(_T("namsinh"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		//capbac
		rs.GetValue(_T("capbac"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		//dv
		rs.GetValue(_T("dv"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		//chandoan
		rs.GetValue(_T("chandoan"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		//doituong
		rs.GetValue(_T("doituong"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		//vaovien
		rs.GetValue(_T("vaokhoa"), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);
		rs.MoveNext();
	}
	CString szSysDate;
	szSysDate = pMF->GetSysDate(); 
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
	
} 
void CRptDanhsachBNVaoVien::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CExcel xls;
	CString szSQL, tmpStr;
	CRecord rs(&pMF->m_db);
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	int nCol = 0, nRow = 0;
	//Width
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 25);
	xls.SetColumnWidth(5, 11);
	xls.SetColumnWidth(6, 30);
	xls.SetColumnWidth(7, 45);
	xls.SetColumnWidth(8, 10);

	xls.SetCellMergedColumns(nCol, nRow , 3);
	xls.SetCellMergedColumns(nCol, nRow+1, 4);
	xls.SetCellMergedColumns(nCol, nRow+2, 10);
	xls.SetCellMergedColumns(nCol, nRow+3, 10);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow+1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow+2, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N V\xC0O VI\x1EC6N"), FMT_TEXT | FMT_CENTER, true, 11);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow+3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);
	//Column header
	nRow = 4;
	CStringArray arrCol;
	arrCol.Add(_T("Stt"));
	arrCol.Add(_T("Doc No"));
	arrCol.Add(_T("Inward No"));
	arrCol.Add(_T("T\xEAn \x62\x1EC7nh nh\xE2n"));
	arrCol.Add(_T("N\x103m sinh"));
	arrCol.Add(_T("\x43\x1EA5p \x62\x1EAD\x63"));
	arrCol.Add(_T("\x110v"));
	arrCol.Add(_T("\x43h\x1EA9n \x111o\xE1n"));
	arrCol.Add(_T("\x110\x1ED1i t\x1B0\x1EE3ng"));
	arrCol.Add(_T("V\xE0o kho\x61"));
	for (int i = 0;i<arrCol.GetCount();i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10);	
	}
	//get data
	szSQL = GetQueryString();
	int ret = rs.ExecSQL(szSQL);
	if (ret <=0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	int nIdx = 0;
	while (!rs.IsEOF())
	{
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow+1, tmpStr, FMT_INTEGER);
		//docno
		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol+1, nRow+1, tmpStr, FMT_INTEGER);
		//inwardno
		rs.GetValue(_T("inwardno"), tmpStr);
		xls.SetCellText(nCol+2, nRow+1, tmpStr, FMT_INTEGER);
		//tenbenhnhan
		rs.GetValue(_T("tenbenhnhan"), tmpStr);
		xls.SetCellText(nCol+3, nRow+1, tmpStr, FMT_TEXT);
		//namsinh
		rs.GetValue(_T("namsinh"), tmpStr);
		xls.SetCellText(nCol+4, nRow+1, tmpStr, FMT_INTEGER);
		//capbac
		rs.GetValue(_T("capbac"), tmpStr);
		xls.SetCellText(nCol+5, nRow+1, tmpStr, FMT_TEXT);
		//dv
		rs.GetValue(_T("dv"), tmpStr);
		xls.SetCellText(nCol+6, nRow+1, tmpStr, FMT_TEXT);
		//chandoan
		rs.GetValue(_T("chandoan"), tmpStr);
		xls.SetCellText(nCol+7, nRow+1, tmpStr, FMT_TEXT);
		//doituong
		rs.GetValue(_T("doituong"), tmpStr);
		xls.SetCellText(nCol+8, nRow+1, tmpStr, FMT_TEXT);
		//vaokhoa
		rs.GetValue(_T("vaokhoa"), tmpStr);
		xls.SetCellText(nCol+9, nRow+1, tmpStr, FMT_TEXT);
		nRow++;
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\Danh sach BN vao vien.xls"));
} 
void CRptDanhsachBNVaoVien::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
//	

} 
int CRptDanhsachBNVaoVien::OnAddRptDanhsachBNVaoVien(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CRptDanhsachBNVaoVien::OnEditRptDanhsachBNVaoVien(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CRptDanhsachBNVaoVien::OnDeleteRptDanhsachBNVaoVien(){
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
 		OnCancelRptDanhsachBNVaoVien(); 
 	}
	return 0;
}
int CRptDanhsachBNVaoVien::OnSaveRptDanhsachBNVaoVien(){
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
 		//OnRptDanhsachBNVaoVienListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CRptDanhsachBNVaoVien::OnCancelRptDanhsachBNVaoVien(){
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
int CRptDanhsachBNVaoVien::OnRptDanhsachBNVaoVienListLoadData(){
	return 0;
}

CString CRptDanhsachBNVaoVien::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	/*if (!m_szObjectKey.IsEmpty())
		szWhere.Format(_T(""), m_szObjectKey);*/
	szSQL.Format(_T(" SELECT  hcr_docno as docno,") \
				_T("         hcr_recordno as inwardno,") \
				_T("         trim(hp_surname||' '||hp_midname||' '||hp_firstname) as tenbenhnhan,") \
				_T("         extract (year from hp_birthdate) as namsinh,") \
				_T("         (select distinct ss_desc from sys_sel where ss_id = 'hms_rank' and ss_code = hp_rank) as capbac,") \
				_T("         hp_workplace as dv,") \
				_T("         hcr_admitdisease as chandoan,") \
				_T("         case when hd_object <> 2 then ") \
				_T("         (select distinct ho_desc from hms_object where ho_id = hd_object) ") \
				_T("         else (select distinct ss_desc from sys_sel where ss_id = 'hms_occupation' and ss_code = 8)") \
				_T("         end as doituong,") \
				_T("         hcr_admitdept as vaokhoa") \
				_T(" FROM hms_clinical_record") \
				_T(" LEFT JOIN hms_doc ON (hd_docno = hcr_docno)") \
				_T(" LEFT JOIN hms_patient ON (hcr_patientno = hp_patientno)") \
				_T(" WHERE hcr_recordno IS NOT NULL AND hd_admitdept = '%s'") \
				_T(" AND hcr_admitdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T(" ORDER BY inwardno"),pMF->m_szDept, m_szFromDate, m_szToDate, szWhere);
	return szSQL;
}
