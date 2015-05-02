#include "stdafx.h"
#include "PMSoldierDrugUsage.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include "StringToken.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CPMSoldierDrugUsage *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSoldierDrugUsage* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CPMSoldierDrugUsage *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSoldierDrugUsage *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSoldierDrugUsage *)pWnd)->OnToDateCheckValue();
} 
static void _OnExaminationSelectFnc(CWnd *pWnd){
	 ((CPMSoldierDrugUsage*)pWnd)->OnExaminationSelect();
}
static void _OnOutpatientSelectFnc(CWnd *pWnd){
	 ((CPMSoldierDrugUsage*)pWnd)->OnOutpatientSelect();
}
static void _OnInwardSelectFnc(CWnd *pWnd){
	 ((CPMSoldierDrugUsage*)pWnd)->OnInwardSelect();
}
static void _OnPrintSelectFnc(CWnd *pWnd){
	CPMSoldierDrugUsage *pVw = (CPMSoldierDrugUsage *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSoldierDrugUsage *pVw = (CPMSoldierDrugUsage *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnObjectListLoadDataFnc(CWnd *pWnd){
	return ((CPMSoldierDrugUsage*)pWnd)->OnObjectListLoadData();
} 
static void _OnObjectListDblClickFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage*)pWnd)->OnObjectListDblClick();
} 
static void _OnObjectListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPMSoldierDrugUsage*)pWnd)->OnObjectListSelectChange(nOldItem, nNewItem);
} 
static int _OnObjectListDeleteItemFnc(CWnd *pWnd){
	 return ((CPMSoldierDrugUsage*)pWnd)->OnObjectListDeleteItem();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CPMSoldierDrugUsage*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CPMSoldierDrugUsage*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPMSoldierDrugUsage*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CPMSoldierDrugUsage*)pWnd)->OnDeptListDeleteItem();
} 
CPMSoldierDrugUsage::CPMSoldierDrugUsage(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CPMSoldierDrugUsage::~CPMSoldierDrugUsage(){
}
void CPMSoldierDrugUsage::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 545);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 280, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 285, 30, 375, 55);
	m_wndReportPeriod.Create(this,380, 30, 570, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 280, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 285, 60, 375, 85);
	m_wndToDate.Create(this,380, 60, 570, 85); 
	m_wndExamination.Create(this, _T("Examination"), 5, 550, 90, 575);
	m_wndOutpatient.Create(this, _T("Outpatient"), 95, 550, 195, 575);
	m_wndInward.Create(this, _T("Inward"), 200, 550, 300, 575);
	m_wndPrint.Create(this, _T("&Print"), 410, 550, 490, 575);
	m_wndExport.Create(this, _T("&Export"), 495, 550, 575, 575);
	m_wndObjectList.Create(this,10, 90, 570, 310); 
	m_wndDeptList.Create(this,10, 315, 570, 540); 

}
void CPMSoldierDrugUsage::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndObjectList.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndObjectList.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);
	m_wndObjectList.InsertColumn(2, _T("Type"), CFMT_TEXT, 0);
	m_wndObjectList.SetCheckBox(TRUE);

	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.SetCheckBox(TRUE);
}

void CPMSoldierDrugUsage::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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
	m_wndExamination.SetEvent(WE_CLICK, _OnExaminationSelectFnc);
	m_wndOutpatient.SetEvent(WE_CLICK, _OnOutpatientSelectFnc);
	m_wndInward.SetEvent(WE_CLICK, _OnInwardSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndObjectList.SetEvent(WE_SELCHANGE, _OnObjectListSelectChangeFnc);
	m_wndObjectList.SetEvent(WE_LOADDATA, _OnObjectListLoadDataFnc);
	m_wndObjectList.SetEvent(WE_DBLCLICK, _OnObjectListDblClickFnc);
	m_wndObjectList.AddEvent(1, _T("Delete"), _OnObjectListDeleteItemFnc, 0, VK_DELETE, 0);
	m_wndDeptList.SetEvent(WE_SELCHANGE, _OnDeptListSelectChangeFnc);
	m_wndDeptList.SetEvent(WE_LOADDATA, _OnDeptListLoadDataFnc);
	m_wndDeptList.SetEvent(WE_DBLCLICK, _OnDeptListDblClickFnc);
	m_wndDeptList.AddEvent(1, _T("Delete"), _OnDeptListDeleteItemFnc, 0, VK_DELETE, 0);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);
	OnObjectListLoadData();
	OnDeptListLoadData();

}
void CPMSoldierDrugUsage::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_Check(pDX, m_wndExamination.GetDlgCtrlID(), m_bExamination);
	DDX_Radio(pDX, m_wndOutpatient.GetDlgCtrlID(), m_nOutpatient);
	DDX_Radio(pDX, m_wndInward.GetDlgCtrlID(), m_nInward);

}
void CPMSoldierDrugUsage::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_nOutpatient = 1;
	m_nInward = 1;
}

int CPMSoldierDrugUsage::SetMode(int nMode){
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

/*void CPMSoldierDrugUsage::OnYearChange(){
	
} */
/*void CPMSoldierDrugUsage::OnYearSetfocus(){
	
} */
/*void CPMSoldierDrugUsage::OnYearKillfocus(){
	
} */
int CPMSoldierDrugUsage::OnYearCheckValue(){
	return 0;
}
 
void CPMSoldierDrugUsage::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
void CPMSoldierDrugUsage::OnReportPeriodSelendok(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	UpdateData(true);
	CDate dte;
	CString tmpStr;
	int nMonth = ToInt(m_szReportPeriodKey);
	int nYear = m_nYear;
	if(nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/1"), nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), nYear, nMonth, dte.GetMonthLastDay());
	}
	if(nMonth == 13){
		m_szFromDate.Format(_T("%.4d/1/1"), nYear);
		tmpStr.Format(_T("%.4d/3/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/3/%.2d 23:59"), nYear,  dte.GetMonthLastDay());
	}
	if(nMonth == 14){
		m_szFromDate.Format(_T("%.4d/4/1"), nYear);
		tmpStr.Format(_T("%.4d/6/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/6/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 15){
		m_szFromDate.Format(_T("%.4d/7/1"), nYear);
		tmpStr.Format(_T("%.4d/9/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/9/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 16){
		m_szFromDate.Format(_T("%.4d/10/1"), nYear);
		tmpStr.Format(_T("%.4d/10/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 17){
		m_szFromDate.Format(_T("%.4d/1/1"), nYear);
		tmpStr.Format(_T("%.4d/12/1"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());		
	}	
	UpdateData(false); 
}

/*void CPMSoldierDrugUsage::OnReportPeriodSetfocus(){
	
}*/
/*void CPMSoldierDrugUsage::OnReportPeriodKillfocus(){
	
}*/
long CPMSoldierDrugUsage::OnReportPeriodLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

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

/*void CPMSoldierDrugUsage::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CPMSoldierDrugUsage::OnFromDateChange(){
	
} */
/*void CPMSoldierDrugUsage::OnFromDateSetfocus(){
	
} */
/*void CPMSoldierDrugUsage::OnFromDateKillfocus(){
	
} */
int CPMSoldierDrugUsage::OnFromDateCheckValue(){
	return 0;
}
 
/*void CPMSoldierDrugUsage::OnToDateChange(){
	
} */
/*void CPMSoldierDrugUsage::OnToDateSetfocus(){
	
} */
/*void CPMSoldierDrugUsage::OnToDateKillfocus(){
	
} */
int CPMSoldierDrugUsage::OnToDateCheckValue(){
	return 0;
}
 
void CPMSoldierDrugUsage::OnExaminationSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	OnDeptListLoadData();
	if (m_bExamination)
	{
		m_wndInward.EnableWindow(FALSE);
		m_wndOutpatient.EnableWindow(FALSE);
	}
	else
	{
		m_wndInward.EnableWindow(TRUE);
		m_wndOutpatient.EnableWindow(TRUE);
	}
	m_nOutpatient = 1;
	m_nInward = 1;
	UpdateData(false);
}

void CPMSoldierDrugUsage::OnOutpatientSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}

void CPMSoldierDrugUsage::OnInwardSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}

void CPMSoldierDrugUsage::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	/*Declaration Section*/
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL;
	double nTmp = 0;
	int nIdx = 1;
	CString szSQL, tmpStr, szReportName, szDate, szDepts, szObjects;
	long double nTotal = 0;
	szReportName = _T("Reports\\HMS\\PM_QUYETTOANSUDUNGTHUOC-VTYT.RPT");
	if (!rpt.Init(szReportName))
		return;
	if (m_bExamination)
		szSQL = GetQueryString1();
	else
		szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
		CDate::Convert(m_szToDate), yyyymmdd, ddmmyyyy);
		rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
		for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
		{
			if (m_wndObjectList.GetCheck(i))
			{
				if (!szObjects.IsEmpty())
					szObjects += _T(", ");
				szObjects += m_wndObjectList.GetItemText(i, 1);
			}
		}
		for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
		{
			if (m_wndDeptList.GetCheck(i))
			{
				if (!szDepts.IsEmpty())
					szDepts += _T(", ");
				szDepts.AppendFormat(_T("'%s'"), m_wndDeptList.GetItemText(i, 1));
			}
		}
		CStringToken token(szDepts, _T(","));
		int nCount = token.GetSize();
		if (nCount == 0)
			tmpStr = _T("To\xE0n \x62\x1ED9");
		else if (nCount == 1)
			tmpStr = szDepts;
		else
			tmpStr = _T("Nhi\x1EC1u kho\x61");
		rptHeader->SetValue(_T("Object"), szObjects);
		rptHeader->SetValue(_T("Department"), tmpStr);
	}
	while (!rs.IsEOF())
	{
		/*Group Seperation*/
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("patient_name")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("rank")));
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("work_place")));
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("patient_id")));
			rptDetail->SetValue(_T("6"), rs.GetValue(_T("dept")));
			rptDetail->SetValue(_T("7"), rs.GetValue(_T("doctor_name")));
			rptDetail->SetValue(_T("8"), rs.GetValue(_T("diagnostic")));
			rptDetail->SetValue(_T("9"), rs.GetValue(_T("numbertr")));
			rs.GetValue(_T("amount"), nTmp);
			nTotal+= nTmp;
			rptDetail->SetValue(_T("10"), double2str(nTmp));
		}
		rs.MoveNext();
	}
	if (nTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
		tmpStr.Format(_T("%f"), nTotal);
		rptDetail->SetValue(_T("s9"), tmpStr);
	}
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), pMF->GetSysTime(), tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
}
 
void CPMSoldierDrugUsage::OnExportSelect(){
	CMainFrame_E10 *pMF	= (CMainFrame_E10*)	AfxGetMainWnd();
	UpdateData(true);
	CRecord	rs(&pMF->m_db);
	CExcel xls;
	CString	szSQL, tmpStr;
	double nTmp	= 0;
	long double nTotal = 0;
	int	nIdx = 1, nRow = 0;
	if (m_bExamination)
		szSQL = GetQueryString1();
	else
		szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 10);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 25);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 5);
	xls.SetColumnWidth(6, 10);
	xls.SetColumnWidth(7, 25);
	xls.SetColumnWidth(8, 10);
	xls.SetColumnWidth(9, 15);
	xls.SetCellMergedColumns(0,	0, 2);
	xls.SetCellMergedColumns(0,	1, 2);
	xls.SetCellMergedColumns(0,	2, 9);
	xls.SetCellMergedColumns(0,	3, 9);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098,	true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO:	Write Excel	Name
	tmpStr = _T("\x42\xC1O \x43\xC1O S\x1EEC \x44\x1EE4NG THU\x1ED0\x43 \x110\x1ED0I T\x1AF\x1EE2NG \x42\x1ED8 \x110\x1ED8I");
	xls.SetCellText(0, 2, tmpStr, 4098,	true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),	CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098,	true);
	//TODO:	Write Column Header
	xls.SetCellText(0, 5, _T("STT"), 4098, true);
	xls.SetCellText(1, 5, _T("H\x1ECD v\xE0 t\xEAn"), 4098, true);
	xls.SetCellText(2, 5, _T("\x43\x1EA5p \x62\x1EAD\x63"), 4098, true);
	xls.SetCellText(3, 5, _T("\x110\x1A1n v\x1ECB"), 4098, true);
	xls.SetCellText(4, 5, _T("S\x1ED1 \x42\x41"), 4098, true);
	xls.SetCellText(5, 5, _T("Kho\x61 \x111i\x1EC1u tr\x1ECB"), 4098, true);
	xls.SetCellText(6, 5, _T("\x42S\x110T"), 4098, true);
	xls.SetCellText(7, 5, _T("\x43h\x1EA9n \x111o\xE1n"), 4098, true);
	xls.SetCellText(8, 5, _T("Ng\xE0y \x110T"), 4098, true);
	xls.SetCellText(9, 5, _T("Th\xE0nh ti\x1EC1n"), 4098, true);
	xls.SetCellText(10, 5, _T("HS"), 4098, true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);	
		xls.SetCellText(1, nRow, rs.GetValue(_T("patient_name")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("rank")), FMT_TEXT);
		xls.SetCellText(3, nRow, rs.GetValue(_T("work_place")),	FMT_TEXT);
		xls.SetCellText(4, nRow, rs.GetValue(_T("patient_id")), FMT_TEXT);
		xls.SetCellText(5, nRow, rs.GetValue(_T("dept")), FMT_TEXT);
		xls.SetCellText(6, nRow, rs.GetValue(_T("doctor_name")), FMT_TEXT);
		xls.SetCellText(7, nRow, rs.GetValue(_T("diagnostic")),	FMT_TEXT);
		xls.SetCellText(8, nRow, rs.GetValue(_T("numbertr")), FMT_NUMBER1);
		rs.GetValue(_T("amount"), nTmp);
		nTotal += nTmp;
		xls.SetCellText(9, nRow, double2str(nTmp), FMT_NUMBER1);
		xls.SetCellText(10, nRow, rs.GetValue(_T("doc_no")), FMT_TEXT);
		nRow++;
		rs.MoveNext();
	}
	if (nTotal > 0)
	{
		tmpStr.Format(_T("%f"), nTotal);
		xls.SetCellText(8, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
		xls.SetCellText(9, nRow, tmpStr, FMT_NUMBER1);
	}
	xls.Save(_T("Exports\\Bao cao su dung thuoc dt bo doi.xls"));
}
 
void CPMSoldierDrugUsage::OnObjectListDblClick(){
	
}
 
void CPMSoldierDrugUsage::OnObjectListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
int CPMSoldierDrugUsage::OnObjectListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
}
 
long CPMSoldierDrugUsage::OnObjectListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	CString szWhere;
	m_wndObjectList.BeginLoad(); 
	int nCount = 0;

	//szWhere.Format(_T(" AND ho_type IN('S') "));

	szSQL.Format(_T("SELECT ho_id AS ID, ") \
		         _T("ho_desc AS Name, ") \
				 _T("ho_type AS Type ") \
		         _T("FROM hms_object ") \
				 _T("WHERE ho_id IN (1, 3, 8) %s ") \
				 _T("ORDER BY cast(ho_id as integer)"),
				 szWhere);

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObjectList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")),
			rs.GetValue(_T("Type")), NULL);
		rs.MoveNext();
	}
	m_wndObjectList.EndLoad(); 
	return nCount;
}

void CPMSoldierDrugUsage::OnDeptListDblClick(){
	
}
 
void CPMSoldierDrugUsage::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
int CPMSoldierDrugUsage::OnDeptListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
}
 
long CPMSoldierDrugUsage::OnDeptListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	
	m_wndDeptList.BeginLoad();
	int nCount = 0;
	szWhere = _T("'DT'");
	if (m_bExamination)
		szWhere = _T("'KB'");
	szSQL.Format(_T("SELECT sd_id as id, sd_name as name ") \
		         _T("FROM sys_dept ") \
				 _T("WHERE 1=1 AND sd_type = %s ORDER BY id "), szWhere);

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDeptList.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	m_wndDeptList.EndLoad();
	return nCount;
}

CString CPMSoldierDrugUsage::GetQueryString(){
	CString szSQL, szWhere, szObjects, szDepts;
	szWhere.Format(_T(" AND discharge_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') ") \
		_T(" AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')"), m_szFromDate, m_szToDate);
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (m_wndObjectList.GetCheck(i))
		{
			if (!szObjects.IsEmpty())
				szObjects += _T(", ");
			szObjects += m_wndObjectList.GetItemText(i, 0);
		}
	}
	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			if (!szDepts.IsEmpty())
				szDepts += _T(", ");
			szDepts.AppendFormat(_T("'%s'"), m_wndDeptList.GetItemText(i, 0));
		}
	}
	if (szObjects.IsEmpty())
		szObjects = _T("1, 3, 8");
	szWhere.AppendFormat(_T(" AND hfe_object IN (%s)"), szObjects);
	if (!szDepts.IsEmpty())
		szWhere.AppendFormat(_T(" AND discharge_dept IN (%s)"), szDepts);
	if (m_nOutpatient == 0)
		szWhere.AppendFormat(_T(" AND hd_outpatient = 'Y'"));
	else if (m_nInward == 0)
		szWhere.AppendFormat(_T(" AND NVL(hd_outpatient, 'X') <> 'Y'"));
	szSQL.Format(_T(" SELECT    Get_patientname(doc_no) patient_name, ") \
				_T("           Get_syssel_desc('hms_rank', Decode(0, doc_rank, hp_rank, ") \
				_T("                                                 hp_rank)) rank, ") \
				_T("           hp_workplace work_place, ") \
				_T("           record_no patient_id, ") \
				_T("           discharge_dept dept, ") \
				_T("           (SELECT Substr(su_name, Instr(su_name, ' ', -1) + 1) ") \
				_T("            FROM   sys_user ") \
				_T("            WHERE  su_userid = treat_doctor) doctor_name, ") \
				_T("           main_disease diagnostic, ") \
				_T("           sumtreat numbertr, ") \
				_T("           SUM(hfe_polprice * hfe_quantity) amount ") \
				_T(" FROM     (SELECT    hd_docno doc_no, ") \
				_T("                     hcrf_invoicefee invoice_fee, ") \
				_T("                     hcr_dischargedept discharge_dept, ") \
				_T("                     hd_outpatient, ") \
				_T("                     hd_patientno patient_no, ") \
				_T("                     hd_rank doc_rank, ") \
				_T("                     hcr_recordno record_no, ") \
				_T("                     hcr_treatdoctor treat_doctor, ") \
				_T("                     hcr_maindisease main_disease, ") \
				_T("                     hcr_sumtreat sumtreat, ") \
				_T("                     hcr_dischargedate discharge_date ") \
				_T("           FROM      hms_clinical_record ") \
				_T("           LEFT JOIN hms_doc ON ( hcr_docno = hd_docno ) ") \
				_T("           WHERE     hd_outpatient = 'N' AND hcr_status = 'T' AND hcrf_acceptedfee IN ( 'Y', 'P' ) ") \
				_T("           UNION ALL ") \
				_T("           SELECT    DISTINCT hd_docno, ") \
				_T("                              htrf_invoicefee, ") \
				_T("                              htr_tdeptid, ") \
				_T("                              hd_outpatient, ") \
				_T("                              hd_patientno, ") \
				_T("                              hd_rank, ") \
				_T("                              hcr_recordno, ") \
				_T("                              htr_doctor, ") \
				_T("                              htr_maindisease, ") \
				_T("                              htr_sumtreat, ") \
				_T("                              htr_dischargedate ") \
				_T("           FROM      hms_clinical_record ") \
				_T("           LEFT JOIN hms_treatment_record ON ( hcr_docno = htr_docno ) ") \
				_T("           LEFT JOIN hms_doc ON ( hcr_docno = hd_docno ) ") \
				_T("           WHERE     hd_outpatient = 'Y' AND htr_status = 'T' AND htrf_acceptedfee IN ( 'Y', 'P' )) ") \
				_T(" LEFT JOIN hms_patient ON ( hp_patientno = patient_no ) ") \
				_T(" LEFT JOIN hms_fee ON ( hfe_docno = doc_no ) ") \
				_T(" LEFT JOIN hms_fee_group ON ( hfg_id = hfe_group ) ") \
				_T(" LEFT JOIN sys_dept ON ( discharge_dept = sd_id ) ") \
				_T(" WHERE     hfe_category <> 'Y' AND hfe_inspaid > 0 AND hfg_type_id = 800 %s") \
				_T(" GROUP     BY doc_no,doc_rank,hp_rank,discharge_dept,sd_index,hp_workplace,record_no,treat_doctor,main_disease,sumtreat ") \
				_T(" ORDER     BY sd_index "), szWhere);

_fmsg(_T("%s"), szSQL);
	return szSQL;
}

CString CPMSoldierDrugUsage::GetQueryString1(){
	CString szSQL, szWhere, szObjects, szDepts;
	szWhere.Format(_T(" AND he_examdate BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') ") \
		_T(" AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')"), m_szFromDate, m_szToDate);
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (m_wndObjectList.GetCheck(i))
		{
			if (!szObjects.IsEmpty())
				szObjects += _T(", ");
			szObjects += m_wndObjectList.GetItemText(i, 0);
		}
	}
	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			if (!szDepts.IsEmpty())
				szDepts += _T(", ");
			szDepts.AppendFormat(_T("'%s'"), m_wndDeptList.GetItemText(i, 0));
		}
	}
	if (szObjects.IsEmpty())
		szObjects = _T("1, 3, 8");
	szWhere.AppendFormat(_T(" AND hpo_object_id IN (%s)"), szObjects);
	if (!szDepts.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_deptid IN (%s)"), szDepts);
	szSQL.Format(_T(" SELECT    Get_patientname(hd_docno)            patient_name, ") \
	_T("						hd_docno doc_no,") \
	_T("           get_syssel_desc('hms_rank', DECODE(0, hd_rank, hp_rank, hp_rank)) rank, ") \
	_T("           hp_workplace                         work_place, ") \
	_T("           hd_enddept                    dept, ") \
	_T("           (SELECT Substr(su_name, Instr(su_name, ' ', -1) + 1) ") \
	_T("            FROM   sys_user ") \
	_T("            WHERE  su_userid = hd_doctor) doctor_name, ") \
	_T("		   hd_docno patient_id,") \
	_T("           hd_conclusion                      diagnostic, ") \
	_T("           SUM(hfe_polprice * hfe_quantity)     amount ") \
	_T(" FROM      hms_doc") \
	_T(" LEFT JOIN hms_pharmaorder ON (hpo_docno = hd_docno)") \
	_T(" LEFT JOIN hms_fee ON ( hfe_docno = hd_docno AND hpo_orderid = hfe_orderid) ") \
	_T(" LEFT JOIN m_productitem_view ON (cast(product_item_id as nvarchar2(15))= hfe_itemid)") \
	_T(" LEFT JOIN hms_exam ON (hd_docno = he_docno AND he_deptid = hpo_deptid)") \
	_T(" LEFT JOIN hms_patient ON ( hp_patientno = hd_patientno ) ") \
	_T(" LEFT JOIN hms_fee_group ON ( hfg_id = hfe_group ) ") \
	_T(" LEFT JOIN sys_dept ON (sd_id = hd_enddept)") \
	_T(" LEFT JOIN m_storagelist ON (msl_storage_id = hpo_storage_id)") \
	_T(" WHERE     hd_suggestion NOT IN ('C', 'D') AND hfe_status <> 'C' AND msl_type = 'C'") \
	_T(" AND product_org_id = 'PM' AND hpo_orderstatus NOT IN ('O', 'C') %s ") \
	_T(" GROUP     BY hd_docno, ") \
	_T("              hd_rank, ") \
	_T("			  hp_rank,") \
	_T("              hd_enddept, ") \
	_T("			  sd_index,") \
	_T("              hp_workplace, ") \
	_T("              hd_doctor, ") \
	_T("			  hd_docno,") \
	_T("              hd_conclusion ") \
	_T(" ORDER BY sd_index"), szWhere);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}