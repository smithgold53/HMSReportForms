#include "stdafx.h"
#include "EMInwardStatisticForDepartmentReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMInwardStatisticForDepartmentReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMInwardStatisticForDepartmentReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMInwardStatisticForDepartmentReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMInwardStatisticForDepartmentReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMInwardStatisticForDepartmentReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnDepartmentSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMInwardStatisticForDepartmentReport* )pWnd)->OnDepartmentSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDepartmentSelendokFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnDepartmentSelendok();
}
/*static void _OnDepartmentSetfocusFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnDepartmentKillfocus();
}*/
/*static void _OnDepartmentKillfocusFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnDepartmentKillfocus();
}*/
static long _OnDepartmentLoadDataFnc(CWnd *pWnd){
	return ((CEMInwardStatisticForDepartmentReport *)pWnd)->OnDepartmentLoadData();
}
/*static void _OnDepartmentAddNewFnc(CWnd *pWnd){
	((CEMInwardStatisticForDepartmentReport *)pWnd)->OnDepartmentAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMInwardStatisticForDepartmentReport *pVw = (CEMInwardStatisticForDepartmentReport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMInwardStatisticForDepartmentReport *pVw = (CEMInwardStatisticForDepartmentReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMInwardStatisticForDepartmentReport *pVw = (CEMInwardStatisticForDepartmentReport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddEMInwardStatisticForDepartmentReportFnc(CWnd *pWnd){
	 return ((CEMInwardStatisticForDepartmentReport*)pWnd)->OnAddEMInwardStatisticForDepartmentReport();
} 
static int _OnEditEMInwardStatisticForDepartmentReportFnc(CWnd *pWnd){
	 return ((CEMInwardStatisticForDepartmentReport*)pWnd)->OnEditEMInwardStatisticForDepartmentReport();
} 
static int _OnDeleteEMInwardStatisticForDepartmentReportFnc(CWnd *pWnd){
	 return ((CEMInwardStatisticForDepartmentReport*)pWnd)->OnDeleteEMInwardStatisticForDepartmentReport();
} 
static int _OnSaveEMInwardStatisticForDepartmentReportFnc(CWnd *pWnd){
	 return ((CEMInwardStatisticForDepartmentReport*)pWnd)->OnSaveEMInwardStatisticForDepartmentReport();
} 
static int _OnCancelEMInwardStatisticForDepartmentReportFnc(CWnd *pWnd){
	 return ((CEMInwardStatisticForDepartmentReport*)pWnd)->OnCancelEMInwardStatisticForDepartmentReport();
} 
CEMInwardStatisticForDepartmentReport::CEMInwardStatisticForDepartmentReport(CWnd *pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 155;
	m_szTitle = _T("Th\x1ED1ng k\xEA \x62\x1EC7nh nh\xE2n v\xE0o vi\x1EC7n th\x65o kho\x61");
	SetDefaultValues();
}
CEMInwardStatisticForDepartmentReport::~CEMInwardStatisticForDepartmentReport(){
}
void CEMInwardStatisticForDepartmentReport::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 490, 120);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 245, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 250, 30, 330, 55);
	m_wndReportPeriod.Create(this,335, 30, 485, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 245, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 60, 330, 85);
	m_wndToDate.Create(this,335, 60, 485, 85); 
	m_wndDepartmentLabel.Create(this, _T("Department"), 10, 90, 90, 115);
	m_wndDepartment.Create(this,95, 90, 485, 115); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 300, 125, 410, 150);
	m_wndExport.Create(this, _T("&Export"), 415, 125, 490, 150);

}
void CEMInwardStatisticForDepartmentReport::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
//	EnableControls(TRUE);
//	EnableButtons(TRUE, 0,1,2,3, -1);
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(false);
	m_wndReportPeriod.SetCheckValue(false);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(false);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(false);
	m_wndDepartment.SetCheckValue(false);
	m_wndDepartment.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

	m_wndDepartment.InsertColumn(0, _T("ID"), CFMT_NUMBER, 100);
	m_wndDepartment.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);
}
void CEMInwardStatisticForDepartmentReport::OnSetWindowEvents(){
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
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	SetMode(VM_EDIT);
}
void CEMInwardStatisticForDepartmentReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDepartment.GetDlgCtrlID(), m_szDepartmentKey);

}
void CEMInwardStatisticForDepartmentReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMInwardStatisticForDepartmentReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMInwardStatisticForDepartmentReport::SetDefaultValues(){
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDepartmentKey.Empty();
}
int CEMInwardStatisticForDepartmentReport::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, 1, 2, 3, 4, -1); 
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
/*void CEMInwardStatisticForDepartmentReport::OnYearChange(){
	
} */
/*void CEMInwardStatisticForDepartmentReport::OnYearSetfocus(){
	
} */
/*void CEMInwardStatisticForDepartmentReport::OnYearKillfocus(){
	
} */
int CEMInwardStatisticForDepartmentReport::OnYearCheckValue(){
	return 0;
} 
void CEMInwardStatisticForDepartmentReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMInwardStatisticForDepartmentReport::OnReportPeriodSelendok(){
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
/*void CEMInwardStatisticForDepartmentReport::OnReportPeriodSetfocus(){
	
}*/
/*void CEMInwardStatisticForDepartmentReport::OnReportPeriodKillfocus(){
	
}*/
long CEMInwardStatisticForDepartmentReport::OnReportPeriodLoadData(){
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
/*void CEMInwardStatisticForDepartmentReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMInwardStatisticForDepartmentReport::OnFromDateChange(){
	
} */
/*void CEMInwardStatisticForDepartmentReport::OnFromDateSetfocus(){
	
} */
/*void CEMInwardStatisticForDepartmentReport::OnFromDateKillfocus(){
	
} */
int CEMInwardStatisticForDepartmentReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMInwardStatisticForDepartmentReport::OnToDateChange(){
	
} */
/*void CEMInwardStatisticForDepartmentReport::OnToDateSetfocus(){
	
} */
/*void CEMInwardStatisticForDepartmentReport::OnToDateKillfocus(){
	
} */
int CEMInwardStatisticForDepartmentReport::OnToDateCheckValue(){
	return 0;
} 
void CEMInwardStatisticForDepartmentReport::OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMInwardStatisticForDepartmentReport::OnDepartmentSelendok(){
	 
}
/*void CEMInwardStatisticForDepartmentReport::OnDepartmentSetfocus(){
	
}*/
/*void CEMInwardStatisticForDepartmentReport::OnDepartmentKillfocus(){
	
}*/
long CEMInwardStatisticForDepartmentReport::OnDepartmentLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDepartment.IsSearchKey() && !m_szDepartmentKey.IsEmpty()){
		szWhere.Format(_T("and sd_id = '%s'"), m_szDepartmentKey);
	};
	m_wndDepartment.DeleteAllItems(); 
	szSQL.Format(_T("select sd_id as id, sd_name as name from sys_dept where sd_type = 'DT' order by id"), szWhere);
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
/*void CEMInwardStatisticForDepartmentReport::OnDepartmentAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMInwardStatisticForDepartmentReport::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString tmpStr, szSQL;
	CRecord rs(&pMF->m_db);

	if(!rpt.Init(_T("Reports/HMS/HE_THONGKEBENHNHANVAOVIENTHEOKHOA.RPT")) )
		return ;
	
	BeginWaitCursor();
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	rs.GetValue(_T("tenkhoa"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("Department"), tmpStr);

	rs.GetValue(_T("tenkhoa"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);

	int nIndex = 1;
	long nTotal[14];
	for(int i = 1; i<= 13; i++)
	{
		 nTotal[i] = 0;
	}
	CReportSection* rptDetail;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%ld"), nIndex ++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("admdept"),tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("c1"), tmpStr);
		nTotal[1] += ToLong(tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		
		rs.GetValue(_T("c2"), tmpStr);
		nTotal[2] += ToLong(tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		
		rs.GetValue(_T("c3"), tmpStr);
		nTotal[3] += ToLong(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("c4"), tmpStr);
		nTotal[4] += ToLong(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("c5"), tmpStr);
		nTotal[5] += ToLong(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("c6"), tmpStr);
		nTotal[6] += ToLong(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("c7"), tmpStr);
		nTotal[7] += ToLong(tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("c8"), tmpStr);
		nTotal[8] += ToLong(tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.GetValue(_T("c9"), tmpStr);
		nTotal[9] += ToLong(tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.GetValue(_T("c10"), tmpStr);
		nTotal[10] += ToLong(tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);

		rs.GetValue(_T("c11"), tmpStr);
		nTotal[11] += ToLong(tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);

		rs.GetValue(_T("c12"), tmpStr);
		nTotal[12] += ToLong(tmpStr);
		rptDetail->SetValue(_T("14"), tmpStr);

		rs.GetValue(_T("c13"), tmpStr);
		nTotal[13] += ToLong(tmpStr);
		rptDetail->SetValue(_T("15"), tmpStr);
		rs.MoveNext();
			
	}	
		for (int i = 1; i <= 13; i++)
		{
			CString  szField;
			szField.Format(_T("s%d"), i+2);
			tmpStr.Format(_T("%ld"), nTotal[i]);
			rpt.GetReportFooter()->SetValue(szField, tmpStr);	
		}
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();
} 
void CEMInwardStatisticForDepartmentReport::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0, nQty = 0; 
	int nTtl[13];
	BeginWaitCursor();
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	
	xls.SetColumnWidth(0, 30);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 15);
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 11);
	xls.SetCellMergedColumns(nCol, nRow + 3, 11);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, _T("TK \x42\x1EC6NH NH\xC2N V\xC0O VI\x1EC6N TH\x45O KHO\x41"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol, arrVal;
	arrCol.Add(_T("Khoa/Ph\xF2ng"));
	arrCol.Add(_T("Qu\xE2n"));
	arrCol.Add(_T("BHYT Qu\xE2n"));
	arrCol.Add(_T("\x43h\xEDnh s\xE1\x63h"));
	arrCol.Add(_T("BHYT Kh\xE1\x63"));
	arrCol.Add(_T("BHYT th\xE2n nh\xE2n"));
	arrCol.Add(_T("<6 tu\x1ED5i"));
	arrCol.Add(_T("\x44\x1ECB\x63h v\x1EE5"));
	arrCol.Add(_T("\x44\xE2n"));
	arrCol.Add(_T("<6 tu\x1ED5i (\x63\xF3 th\x1EBB)"));
	arrCol.Add(_T("BHYT Qu\xE2n \x111\x1ED9i"));
	arrCol.Add(_T("\x42HYT Qu\xE2n nh\xE2n"));
	arrCol.Add(_T("\x42HYT (N\x1EE3 th\x1EBB)"));
	arrCol.Add(_T("T\x1ED5ng \x63\x1ED9ng"));
	nRow = 4;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10); 
	}
	for (int i = 0; i < 14; i++)
	{
		tmpStr.Format(_T("c%d"), i+1);
		arrVal.Add(tmpStr);
		nTtl[i] = 0;
	}
	nRow = 5;
	while (!rs.IsEOF()){
		rs.GetValue(_T("admdept"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
		for (int i = 0; i< arrVal.GetCount(); i++)
		{
			rs.GetValue(arrVal.GetAt(i), nQty);
			FormatCurrency(nQty, tmpStr);
			xls.SetCellText(nCol + i + 1, nRow, tmpStr, FMT_INTEGER);
			nTtl[i] += nQty;
		}
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng"), FMT_TEXT , true, 10);
	for (int i = 0; i< 14; i++)
	{
		FormatCurrency(nTtl[i], tmpStr);
		xls.SetCellText(nCol + i + 1, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 10);
	}
	xls.Save(_T("Exports\\Thong Ke Benh Nhan Vao Vien Theo Khoa.xls"));
} 
void CEMInwardStatisticForDepartmentReport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

} 
int CEMInwardStatisticForDepartmentReport::OnAddEMInwardStatisticForDepartmentReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMInwardStatisticForDepartmentReport::OnEditEMInwardStatisticForDepartmentReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMInwardStatisticForDepartmentReport::OnDeleteEMInwardStatisticForDepartmentReport(){
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
 		OnCancelEMInwardStatisticForDepartmentReport(); 
 	}
	return 0;
}
int CEMInwardStatisticForDepartmentReport::OnSaveEMInwardStatisticForDepartmentReport(){
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
 		//OnEMInwardStatisticForDepartmentReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMInwardStatisticForDepartmentReport::OnCancelEMInwardStatisticForDepartmentReport(){
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
int CEMInwardStatisticForDepartmentReport::OnEMInwardStatisticForDepartmentReportListLoadData(){
	return 0;
}
CString CEMInwardStatisticForDepartmentReport::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	szSQL.Format(_T(" SELECT admdept,") \
				_T("       sum(c1) as c1,") \
				_T("       sum(c2) as c2,") \
				_T("       sum(c3) as c3,") \
				_T("       sum(c4) as c4,") \
				_T("       sum(c5) as c5,") \
				_T("       sum(c6) as c6,") \
				_T("       sum(c7) as c7,") \
				_T("       sum(c8) as c8,") \
				_T("       sum(c9) as c9, ") \
				_T("       sum(c10) as c10, ") \
				_T("       sum(c11) as c11, ") \
				_T("       sum(c12) as c12, ") \
				_T("       sum(c1+c2+c3+c4+c5+c6+c7+c8+c9+c10+c11+c12) as c13") \
				_T(" FROM") \
				_T(" (") \
				_T("    SELECT hcr_admitdept as admdept,") \
				_T("           case when hd_object='1' then 1 else 0 end as c1,") \
				_T("           case when hd_object='2' then 1 else 0 end as c2,") \
				_T("           case when hd_object='3' then 1 else 0 end as c3,") \
				_T("           case when hd_object='4' then 1 else 0 end as c4,") \
				_T("           case when hd_object='5' then 1 else 0 end as c5,") \
				_T("           case when hd_object='6' then 1 else 0 end as c6,") \
				_T("           case when hd_object='7' then 1 else 0 end as c7,") \
				_T("           case when hd_object='8' then 1 else 0 end as c8,") \
				_T("           case when hd_object='9' then 1 else 0 end as c9,") \
				_T("           case when hd_object='10' then 1 else 0 end as c10,") \
				_T("           case when hd_object='11' then 1 else 0 end as c11,") \
				_T("           case when hd_object='12' then 1 else 0 end as c12") \
				_T("    FROM hms_doc") \
				_T("	LEFT JOIN hms_clinical_record ON (hcr_docno = hd_docno)") \
				_T("    WHERE hcr_numinward > 0 and hcr_admitdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') AND hd_enddept = '%s'") \
				_T(" ) tbl") \
				_T(" GROUP BY admdept") \
				_T(" ORDER BY admdept"), m_szFromDate, m_szToDate, pMF->GetCurrentDepartmentID());
	//QueryOpener(szSQL);
	return szSQL;
}