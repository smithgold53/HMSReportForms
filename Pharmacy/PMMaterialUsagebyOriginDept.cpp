#include "stdafx.h"
#include "PMMaterialUsagebyOriginDept.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMMaterialUsagebyOriginDept* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept *)pWnd)->OnToDateCheckValue();
} 
static int _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept*) pWnd)->OnTypeLoadData();
}
static void _OnPrintSelectFnc(CWnd *pWnd){
	CPMMaterialUsagebyOriginDept *pVw = (CPMMaterialUsagebyOriginDept *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMMaterialUsagebyOriginDept *pVw = (CPMMaterialUsagebyOriginDept *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnStorageListLoadDataFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept*)pWnd)->OnStorageListLoadData();
}
static long _OnObjectListLoadDataFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept*)pWnd)->OnObjectListLoadData();
} 
static void _OnObjectListDblClickFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept*)pWnd)->OnObjectListDblClick();
} 
static void _OnObjectListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPMMaterialUsagebyOriginDept*)pWnd)->OnObjectListSelectChange(nOldItem, nNewItem);
} 
static int _OnObjectListDeleteItemFnc(CWnd *pWnd){
	 return ((CPMMaterialUsagebyOriginDept*)pWnd)->OnObjectListDeleteItem();
} 
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CPMMaterialUsagebyOriginDept*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPMMaterialUsagebyOriginDept*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CPMMaterialUsagebyOriginDept*)pWnd)->OnDeptListDeleteItem();
} 
static int _OnObjectListCheckAllFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept*)pWnd)->OnObjectListCheckAll();
}
static int _OnObjectListUnCheckAllFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept*)pWnd)->OnObjectListUnCheckAll();
}
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUnCheckAllFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept*)pWnd)->OnDeptListUnCheckAll();
}
static int _OnStorageListCheckAllFnc(CWnd* pWnd){
	return ((CPMMaterialUsagebyOriginDept*)pWnd)->OnStorageListCheckAll();
}
static int _OnStorageListUnCheckAllFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept*)pWnd)->OnStorageListUncheckAll();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	return ((CPMMaterialUsagebyOriginDept*)pWnd)->OnInPatientSelect();
}
CPMMaterialUsagebyOriginDept::CPMMaterialUsagebyOriginDept(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CPMMaterialUsagebyOriginDept::~CPMMaterialUsagebyOriginDept(){
}
void CPMMaterialUsagebyOriginDept::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 575, 545);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 280, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 285, 30, 375, 55);
	m_wndReportPeriod.Create(this,380, 30, 570, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 280, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 285, 60, 375, 85);
	m_wndToDate.Create(this,380, 60, 570, 85); 
	m_wndTypeLabel.Create(this, _T("Type"), 10, 90, 90, 115);
	m_wndType.Create(this,95, 90, 280, 115); 
	m_wndInPatient.Create(this, _T("In Patient"), 5, 550, 105, 575);
	m_wndPrint.Create(this, _T("&Print"), 410, 550, 490, 575);
	m_wndExport.Create(this, _T("&Export"), 495, 550, 575, 575);
	m_wndStorageList.Create(this,10, 120, 280, 320); 
	m_wndObjectList.Create(this,285, 120, 570, 320); 
	m_wndDeptList.Create(this,10, 325, 570, 540); 

}
void CPMMaterialUsagebyOriginDept::OnInitializeComponents(){
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

	m_wndType.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndType.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

	m_wndStorageList.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndStorageList.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);
	m_wndStorageList.SetCheckBox(TRUE);

	m_wndObjectList.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndObjectList.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);
	m_wndObjectList.InsertColumn(2, _T("Type"), CFMT_TEXT, 0);
	m_wndObjectList.SetCheckBox(TRUE);

	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.SetCheckBox(TRUE);

}

void CPMMaterialUsagebyOriginDept::OnSetWindowEvents(){
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
	m_wndType.SetEvent(WE_LOADDATA, _OnTypeLoadDataFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndStorageList.SetEvent(WE_LOADDATA, _OnStorageListLoadDataFnc);
	m_wndObjectList.SetEvent(WE_SELCHANGE, _OnObjectListSelectChangeFnc);
	m_wndObjectList.SetEvent(WE_LOADDATA, _OnObjectListLoadDataFnc);
	m_wndObjectList.SetEvent(WE_DBLCLICK, _OnObjectListDblClickFnc);

	m_wndStorageList.AddEvent(1, _T("Check All"), _OnStorageListCheckAllFnc);
	m_wndStorageList.AddEvent(2, _T("UnCheck All"), _OnStorageListUnCheckAllFnc);

	m_wndDeptList.SetEvent(WE_SELCHANGE, _OnDeptListSelectChangeFnc);
	m_wndDeptList.SetEvent(WE_LOADDATA, _OnDeptListLoadDataFnc);
	m_wndDeptList.SetEvent(WE_DBLCLICK, _OnDeptListDblClickFnc);

	m_wndObjectList.AddEvent(1, _T("Check All"), _OnObjectListCheckAllFnc);
	m_wndObjectList.AddEvent(2, _T("UnCheck All"), _OnObjectListUnCheckAllFnc);

	m_wndDeptList.AddEvent(1, _T("Check All"), _OnDeptListCheckAllFnc);
	m_wndDeptList.AddEvent(2, _T("UnCheck All"), _OnDeptListUnCheckAllFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);
	OnObjectListLoadData();
	OnDeptListLoadData();
	OnStorageListLoadData();

}
void CPMMaterialUsagebyOriginDept::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_Check(pDX, m_wndInPatient.GetDlgCtrlID(), m_bInPatient);

}
void CPMMaterialUsagebyOriginDept::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}

int CPMMaterialUsagebyOriginDept::SetMode(int nMode){
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

/*void CPMMaterialUsagebyOriginDept::OnYearChange(){
	
} */
/*void CPMMaterialUsagebyOriginDept::OnYearSetfocus(){
	
} */
/*void CPMMaterialUsagebyOriginDept::OnYearKillfocus(){
	
} */
int CPMMaterialUsagebyOriginDept::OnYearCheckValue(){
	return 0;
}
 
void CPMMaterialUsagebyOriginDept::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
void CPMMaterialUsagebyOriginDept::OnReportPeriodSelendok(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	UpdateData(true);
	CDate dte;
	CString tmpStr;
	int nMonth = ToInt(m_szReportPeriodKey);
	int nYear = m_nYear;
	if(nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/01"), nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), nYear, nMonth, dte.GetMonthLastDay());
	}
	if(nMonth == 13){
		m_szFromDate.Format(_T("%.4d/01/01"), nYear);
		tmpStr.Format(_T("%.4d/03/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59"), nYear,  dte.GetMonthLastDay());
	}
	if(nMonth == 14){
		m_szFromDate.Format(_T("%.4d/04/01"), nYear);
		tmpStr.Format(_T("%.4d/06/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 15){
		m_szFromDate.Format(_T("%.4d/07/01"), nYear);
		tmpStr.Format(_T("%.4d/09/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 16){
		m_szFromDate.Format(_T("%.4d/10/01"), nYear);
		tmpStr.Format(_T("%.4d/10/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 17){
		m_szFromDate.Format(_T("%.4d/01/01"), nYear);
		tmpStr.Format(_T("%.4d/12/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());		
	}	
	UpdateData(false); 
}

/*void CPMMaterialUsagebyOriginDept::OnReportPeriodSetfocus(){
	
}*/
/*void CPMMaterialUsagebyOriginDept::OnReportPeriodKillfocus(){
	
}*/
long CPMMaterialUsagebyOriginDept::OnReportPeriodLoadData(){
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

/*void CPMMaterialUsagebyOriginDept::OnReportPeriodAddNew(){
	CMainFrame *pMF = (CMainFrame*) AfxGetMainWnd();
	
} */
/*void CPMMaterialUsagebyOriginDept::OnFromDateChange(){
	
} */
/*void CPMMaterialUsagebyOriginDept::OnFromDateSetfocus(){
	
} */
/*void CPMMaterialUsagebyOriginDept::OnFromDateKillfocus(){
	
} */
int CPMMaterialUsagebyOriginDept::OnFromDateCheckValue(){
	return 0;
}
 
/*void CPMMaterialUsagebyOriginDept::OnToDateChange(){
	
} */
/*void CPMMaterialUsagebyOriginDept::OnToDateSetfocus(){
	
} */
/*void CPMMaterialUsagebyOriginDept::OnToDateKillfocus(){
	
} */
int CPMMaterialUsagebyOriginDept::OnToDateCheckValue(){
	return 0;
}
 
long CPMMaterialUsagebyOriginDept::OnTypeLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	int nCount = 0;
	szSQL.Format(_T(" SELECT DISTINCT product_type_id, ") \
	_T("                 product_type_name ") \
	_T(" FROM   (SELECT CASE WHEN mpt_product_type_id IN ( 'A1100', 'A1130', 'A1140', 'A1160', ") \
	_T("                                              'A1170', 'A1180', 'A1190', 'A1400', 'A6000' ) THEN N'A0000' ") \
	_T("                ELSE mpt_product_type_id ") \
	_T("                END product_type_id, ") \
	_T("                CASE WHEN mpt_product_type_id IN ( 'A1100', 'A1130', 'A1140', 'A1160', ") \
	_T("                                              'A1170', 'A1180', 'A1190', 'A1400', 'A6000' ) THEN N'Thu\x1ED1\x63' ") \
	_T("                ELSE mpt_name ") \
	_T("                END product_type_name ") \
	_T("         FROM   m_product_type ") \
	_T("         WHERE  mpt_org_id = '%s' ") \
	_T("            AND mpt_isactive = 'Y' AND mpt_product_type_id <> 'A1000') ") \
	_T(" ORDER BY product_type_id"), pMF->GetModuleID());
	nCount = rs.ExecSQL(szSQL);
	m_wndType.DeleteAllItems();
	while (!rs.IsEOF())
	{
		m_wndType.AddItems(
			rs.GetValue(_T("product_type_id")),
			rs.GetValue(_T("product_type_name")),
			NULL);
		rs.MoveNext();
	}
	return nCount;
	
}

void CPMMaterialUsagebyOriginDept::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CReportSection *rptDetail;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp, szSysDate, szTime, szOldDept, szNewDept, szObjectNames, szStorage;
	bool bCheckServ = false, bCheckIns = false;
	long nCost = 0;
	long double nGrpCost = 0, nTtlCost = 0;
	int nIndex = 1;
	CStringArray strArr;
	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/QUYETTOANSUDUNGTHUOC-VTYT.RPT")))
		return;

	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	rpt.GetReportHeader()->SetValue(_T("Department"), pMF->GetDepartmentName(pMF->GetDepartmentID()));
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm),
	CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	szObjectNames.Empty();

	strArr.Add(_T("\x64\x1ECB\x63h v\x1EE5"));
	strArr.Add(_T("\x42\x110 - \x43S - \x42H"));

	for (int i = 0; i < m_wndStorageList.GetItemCount(); i++)
	{
		if (m_wndStorageList.GetCheck(i))
		{
			if (!szStorage.IsEmpty())
				szStorage += _T(", ");
			szStorage += m_wndStorageList.GetItemText(i, 1);
		}
	}
	//if (szStorage.Find(',') == -1)
	rpt.GetReportHeader()->SetValue(_T("StockName"), szStorage);
	//else
		//rpt.GetReportHeader()->SetValue(_T("StockName"), _T("Nhi\x1EC1u kho"));
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (m_wndObjectList.GetCheck(i))
		{
			tmpStr = m_wndObjectList.GetItemText(i, 2);

			if (tmpStr == _T("S"))
				bCheckServ = true;
			else
				bCheckIns = true;
		}
	}

	if (bCheckServ || bCheckIns)
	{
		if (bCheckServ)
			szObjectNames.AppendFormat(_T("%s"), strArr[0]);

		if (bCheckIns)
		{
			if (!szObjectNames.IsEmpty())
				szObjectNames += _T(", ");
			szObjectNames.AppendFormat(_T("%s"), strArr[1]);
		}

		TranslateString(_T("Object"), szTemp);
		tmpStr.Format(_T("%s %s"), szTemp, szObjectNames);
		rpt.GetReportHeader()->SetValue(_T("Object"), tmpStr);
	}

	if (!bCheckIns && !bCheckServ)
		rpt.GetReportHeader()->SetValue(_T("Object"), _T("T\x1EA5t \x63\x1EA3 \x63\xE1\x63 \x111\x1ED1i t\x1B0\x1EE3ng"));

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("hpo_deptid"), szNewDept);
		if (szNewDept != szOldDept)
		{
			if (nGrpCost > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				tmpStr.Format(_T("%f"), nGrpCost);
				rptDetail->SetValue(_T("TotalAmount"), _T("\x43\x1ED9ng nh\xF3m"));
				rptDetail->SetValue(_T("s9"), tmpStr);
				nTtlCost += nGrpCost;
				nGrpCost = 0;
				nIndex = 1;
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));			
			rptDetail->SetValue(_T("GroupName"), pMF->GetDepartmentName(szNewDept));
			szOldDept = szNewDept;
		}
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIndex++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("pname")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("docno")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("age")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("cardno")));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("workplace")));
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("rank")));
		rptDetail->SetValue(_T("8"), rs.GetValue(_T("recordno")));
		rs.GetValue(_T("amt"), nCost);
		nGrpCost += nCost;
		rptDetail->SetValue(_T("9"), rs.GetValue(_T("amt")));
		rs.MoveNext();
	}
	if (nGrpCost > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
		tmpStr.Format(_T("%f"), nGrpCost);
		rptDetail->SetValue(_T("s9"), tmpStr);
		nTtlCost += nGrpCost;
		nGrpCost = 0;
	}
	if (nTtlCost > 0)
	{
		rptDetail= rpt.AddDetail(rpt.GetGroupFooter(0));
		tmpStr.Format(_T("%f"), nTtlCost);
		rptDetail->SetValue(_T("TotalAmount"), _T("T\x1ED5ng \x63\x1ED9ng"));
		rptDetail->SetValue(_T("s9"), tmpStr);
	}
	szSysDate = pMF->GetSysDate();
	szTime = pMF->GetSysTime();
	_debug(_T("%s || %s"), szTime, szTime.Right(8));
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szTime.Right(8),
		          szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	EndWaitCursor();
	rpt.PrintPreview();
}
 
void CPMMaterialUsagebyOriginDept::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	CString tmpStr, szTemp, szOldLine, szNewLine;
	long nCost = 0;
	long double nGroupTotal = 0, nTotal = 0;
	UpdateData(TRUE);
	BeginWaitCursor();

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_ICONERROR | MB_OK);
		return;
	}

	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 20);
	xls.SetColumnWidth(5, 30);
	xls.SetColumnWidth(6, 20);
	xls.SetColumnWidth(7, 15);
	xls.SetColumnWidth(8, 10);

	int nRow = 1;
	int nCol = 0;
	int nIndex = 1;
	xls.SetRowHeight(6, 45);

	xls.SetCellMergedColumns(0, 1, 4);
	xls.SetCellMergedColumns(0, 2, 4);

	xls.SetCellMergedColumns(6, 1, 4);
	xls.SetCellMergedColumns(6, 2, 4);

	xls.SetCellText(0, 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(6, 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(6, 2, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol, nRow + 3, 10);
	xls.SetCellMergedColumns(nCol, nRow + 4, 10);

	xls.SetCellText(nCol, nRow + 3, _T("T\xECnh h\xECnh s\x1EED \x64\x1EE5ng v\x1EADt t\x1B0 y t\x1EBF"), FMT_TEXT | FMT_CENTER, true, 16, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);

	TranslateString(_T("Index"), tmpStr);
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Patient Name"), tmpStr);
	xls.SetCellText(nCol + 1, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Doc No"), tmpStr);
	xls.SetCellText(nCol + 2, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Age"), tmpStr);
	xls.SetCellText(nCol + 3, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Card No"), tmpStr);
	xls.SetCellText(nCol + 4, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Workplace"), tmpStr);
	xls.SetCellText(nCol + 5, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Rank"), tmpStr);
	xls.SetCellText(nCol + 6, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Record No"), tmpStr);
	xls.SetCellText(nCol + 7, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Total"), tmpStr);
	xls.SetCellText(nCol + 8, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);
	nRow = 6;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("hpo_deptid"), szNewLine);
		if (!szNewLine.IsEmpty() && szNewLine != szOldLine)
		{
			if (nGroupTotal > 0)
			{
				nRow++;
				xls.SetCellText(nCol + 1, nRow, _T("\x43\x1ED9ng kho\x61"), FMT_TEXT, true);
				tmpStr.Format(_T("%.2lf"), nGroupTotal);
				xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_NUMBER1, true);
				nTotal += nGroupTotal;
				nGroupTotal = 0;
			}
			nRow++;
			xls.SetCellText(nCol + 1, nRow, pMF->GetDepartmentName(szNewLine), FMT_TEXT , true);
			szOldLine = szNewLine;
		}

		nRow++;

		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT);
		xls.SetCellText(nCol + 3, nRow, rs.GetValue(_T("age")), FMT_TEXT);
		xls.SetCellText(nCol + 4, nRow, rs.GetValue(_T("cardno")), FMT_TEXT);
		xls.SetCellText(nCol + 5, nRow, rs.GetValue(_T("workplace")), FMT_TEXT);
		xls.SetCellText(nCol + 6, nRow, rs.GetValue(_T("rank")), FMT_TEXT);
		xls.SetCellText(nCol + 7, nRow, rs.GetValue(_T("recordno")), FMT_TEXT);
		rs.GetValue(_T("amt"), nCost);
		nGroupTotal += nCost;
		tmpStr.Format(_T("%ld"), nCost);
		xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_NUMBER1);

		rs.MoveNext();
	}

	if (nGroupTotal > 0)
	{
		nRow++;
		xls.SetCellText(nCol + 1, nRow, _T("\x43\x1ED9ng kho\x61"), FMT_TEXT, true);
		tmpStr.Format(_T("%.2lf"), nGroupTotal);
		xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_NUMBER1, true);
		nTotal += nGroupTotal;
		nGroupTotal = 0;
	}

	if (nTotal > 0)
	{
		nRow++;
		xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT, true);
		tmpStr.Format(_T("%.2lf"), nTotal);
		xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_NUMBER1, true);
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\Tinh hinh su dung vat tu y te.xls"));
	EndWaitCursor();
	xls.Save(_T("Exports\\Tinh hinh su dung vat tu y te.xls"));
}
 
void CPMMaterialUsagebyOriginDept::OnObjectListDblClick(){
	
}
 
void CPMMaterialUsagebyOriginDept::OnObjectListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
int CPMMaterialUsagebyOriginDept::OnObjectListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
}

long CPMMaterialUsagebyOriginDept::OnStorageListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND msl_org_id = '%s'"), pMF->GetModuleID());
	return pMF->LoadStorageIntoListCtrl(&m_wndStorageList, szFilter);
}

long CPMMaterialUsagebyOriginDept::OnObjectListLoadData(){
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
				 _T("WHERE 1=1 %s ") \
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

void CPMMaterialUsagebyOriginDept::OnDeptListDblClick(){
	
}
 
void CPMMaterialUsagebyOriginDept::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
int CPMMaterialUsagebyOriginDept::OnDeptListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
}
 
long CPMMaterialUsagebyOriginDept::OnDeptListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;

	m_wndDeptList.BeginLoad();
	int nCount = 0;
	if (m_bInPatient)
		szWhere = _T(" AND sd_type = 'DT'");
	else
		szWhere = _T(" AND sd_type = 'KB'");
	szSQL.Format(_T("SELECT sd_id as id, sd_name as name ") \
		         _T("FROM sys_dept ") \
				 _T("WHERE 1=1 %s ORDER BY id "), szWhere);

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

CString CPMMaterialUsagebyOriginDept::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szWhere, szWhere1, szObjects, szDepts, szStorage, szDeptWhere;
	//if (!m_bInPatient)
	szWhere.AppendFormat(_T(" AND mt_status = 'A' AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szWhere1.AppendFormat(_T(" AND hpo_orderstatus = 'A' AND hpo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	for (int i = 0; i < m_wndStorageList.GetItemCount(); i++)
	{
		if (m_wndStorageList.GetCheck(i))
		{
			if (!szStorage.IsEmpty())
				szStorage += _T(", ");
			szStorage.AppendFormat(_T("'%s'"), m_wndStorageList.GetItemText(i, 0));
		}
	}
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (m_wndObjectList.GetCheck(i))
		{
			if (!szObjects.IsEmpty())
				szObjects += _T(",");
			szObjects.AppendFormat(_T("'%s'"), m_wndObjectList.GetItemText(i, 0));
		}
	}

	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			if (!szDepts.IsEmpty())
				szDepts += _T(",");
			szDepts.AppendFormat(_T("'%s'"), m_wndDeptList.GetItemText(i, 0));
		}
	}

	if (!szStorage.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND mt_storage_id IN (%s)"), szStorage);
		szWhere1.AppendFormat(_T(" AND hpo_storage_id IN (%s)"), szStorage);
	}
	if (!szObjects.IsEmpty())
	{
		szWhere.AppendFormat(_T(" and hpo_object_id in(%s) "), szObjects);
		szWhere1.AppendFormat(_T(" and hpo_object_id in(%s) "), szObjects);
	}

	if (!szDepts.IsEmpty())
		szDeptWhere.Format(_T(" AND hpo_deptid IN (%s)"), szDepts);
	if (!m_szTypeKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND product_groupid = '%s'"), m_szTypeKey);
		szWhere1.AppendFormat(_T(" AND product_groupid = '%s'"), m_szTypeKey);
	}
	szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	szWhere1.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	if (m_bInPatient)
		szSQL.Format(_T(" SELECT    hpo_docno                              docno, ") \
					_T("           Hms_getage(hd_admitdate, hp_birthdate) age, ") \
					_T("           hd_cardno                              cardno, ") \
					_T("           Get_patientname(hpo_docno)             pname, ") \
					_T("           hpo_deptid, ") \
					_T("           Get_syssel_desc('hms_rank', hp_rank)   rank, ") \
					_T("           hp_workplace                           workplace, ") \
					_T("           hcr_recordno                           recordno, ") \
					_T("           sum(amt) amt ") \
					_T(" FROM      (SELECT    hpo_docno, ") \
					_T("                      CASE WHEN hpo_deptid = 'B5' AND hpo_storage_id = 12 ") \
					_T("						THEN (select ho_deptid FROM hms_operation where ho_idx = hpo_reference_id) ELSE hpo_deptid END AS hpo_deptid, ") \
					_T("                      hpol_qtyorder * hpol_unitprice amt ") \
					_T("            FROM      hms_ipharmaorder ") \
					_T("            LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
					_T("            LEFT JOIN m_productitem_view ON ( hpol_product_item_id = product_item_id ) ") \
					_T("            WHERE     hpo_ordertype IN ('D', 'M', 'B') %s") \
					_T("            UNION ALL ") \
					_T("            SELECT    hpo_docno, ") \
					_T("                      hpo_deptid, ") \
					_T("                      hpol_qtyorder * hpol_unitprice amt ") \
					_T("            FROM      m_transaction ") \
					_T("            LEFT JOIN hms_medical_transaction_view ON ( mt_transaction_id = hmt_reftransaction_id ) ") \
					_T("            LEFT JOIN hms_ipharmaorder ON ( hpo_orderid = hmt_orderid ") \
					_T("                                            AND hpo_docno = hmt_docno ) ") \
					_T("            LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ") \
					_T("                                                AND hmt_product_id = hpol_product_id ) ") \
					_T("            LEFT JOIN m_productitem_view ON ( hpol_product_item_id = product_item_id ) ") \
					_T("            WHERE     mt_doctype IN ('CSO') %s") \
					_T("            UNION ALL ") \
					_T("            SELECT    cast(mt_partner_id as number(10)), ") \
					_T("                      mt_department_to_id, ") \
					_T("                      hpol_qtyissue * hpol_unitprice amt ") \
					_T("            FROM      m_transaction ") \
					_T("			LEFT JOIN purchase_orderline2 ON (mt_transaction_id = pol_transaction_id)") \
					_T("			LEFT JOIN hms_ipharmaorder ON (hpo_orderid = pol_orderid)")
					_T("			LEFT JOIN hms_ipharmaorderline ON (pol_orderid = hpol_orderid AND pol_product_id = hpol_product_id)") \
					_T("            LEFT JOIN m_productitem_view ON ( pol_product_item_id = product_item_id ) ") \
					_T("            WHERE     mt_doctype IN ('CON') %s)") \
					_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hpo_docno ) ") \
					_T(" LEFT JOIN hms_doc ON ( hcr_docno = hd_docno ) ") \
					_T(" LEFT JOIN hms_patient ON ( hp_patientno = hd_patientno ) ") \
					_T(" WHERE 1=1 %s") \
					_T(" GROUP BY hpo_docno, ") \
					_T("		  hpo_deptid, ") \
					_T("		  hd_admitdate, hp_birthdate, hp_rank, hd_cardno, hp_workplace, hcr_recordno") \
					_T(" ORDER     BY hpo_deptid, ") \
					_T("              hpo_docno "), szWhere1, szWhere, szWhere, szDeptWhere);
	else
		szSQL.Format(_T(" SELECT    hpo_docno                              docno, ") \
				_T("           Hms_getage(hd_admitdate, hp_birthdate) age, ") \
				_T("           hd_cardno                              cardno, ") \
				_T("           Get_patientname(hpo_docno)             pname, ") \
				_T("           hpo_deptid, ") \
				_T("           Get_syssel_desc('hms_rank', hp_rank)   rank, ") \
				_T("           hp_workplace                           workplace, ") \
				_T("           sum(amt) amt ") \
				_T(" FROM      (SELECT    hpo_docno, ") \
				_T("                      CASE WHEN hpo_deptid = 'B5' AND hpo_storage_id = 12 ") \
				_T("						THEN (select ho_deptid FROM hms_operation where ho_idx = hpo_reference_id) ELSE hpo_deptid END AS hpo_deptid, ") \
				_T("                      hpol_qtyorder * hpol_unitprice amt ") \
				_T("            FROM      hms_pharmaorder ") \
				_T("            LEFT JOIN hms_pharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T("            LEFT JOIN m_productitem_view ON ( hpol_product_item_id = product_item_id ) ") \
				_T("            WHERE 1=1 %s)") \
				_T(" LEFT JOIN hms_doc ON ( hpo_docno = hd_docno ) ") \
				_T(" LEFT JOIN hms_patient ON ( hp_patientno = hd_patientno ) ") \
				_T(" WHERE 1=1 %s") \
				_T(" GROUP BY hpo_docno, ") \
				_T("		  hpo_deptid, ") \
				_T("		  hd_admitdate, hp_birthdate, hp_rank, hd_cardno, hp_workplace") \
				_T(" ORDER     BY hpo_deptid, ") \
				_T("              hpo_docno "), szWhere1, szDeptWhere);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}

int CPMMaterialUsagebyOriginDept::OnDeptListCheckAll()
{
	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (!m_wndDeptList.GetCheck(i))
		{
			m_wndDeptList.SetCheck(i, TRUE);
		}
	}
	return 0;
}

int CPMMaterialUsagebyOriginDept::OnDeptListUnCheckAll()
{
	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			m_wndDeptList.SetCheck(i, FALSE);
		}
	}
	return 0;
}


int CPMMaterialUsagebyOriginDept::OnObjectListCheckAll()
{
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (!m_wndObjectList.GetCheck(i))
		{
			m_wndObjectList.SetCheck(i, TRUE);
		}
	}
	return 0;
}

int CPMMaterialUsagebyOriginDept::OnObjectListUnCheckAll()
{
	for (int i = 0; i < m_wndObjectList.GetItemCount(); i++)
	{
		if (m_wndObjectList.GetCheck(i))
		{
			m_wndObjectList.SetCheck(i, FALSE);
		}
	}
	return 0;
}

int CPMMaterialUsagebyOriginDept::OnStorageListCheckAll()
{
	for (int i = 0; i < m_wndStorageList.GetItemCount(); i++)
	{
		if (!m_wndStorageList.GetCheck(i))
		{
			m_wndStorageList.SetCheck(i, TRUE);
		}
	}
	return 0;
}

int CPMMaterialUsagebyOriginDept::OnStorageListUncheckAll()
{
	for (int i = 0; i < m_wndStorageList.GetItemCount(); i++)
	{
		if (m_wndStorageList.GetCheck(i))
		{
			m_wndStorageList.SetCheck(i, FALSE);
		}
	}
	return 0;
}

void CPMMaterialUsagebyOriginDept::OnInPatientSelect(){
	OnDeptListLoadData();
}