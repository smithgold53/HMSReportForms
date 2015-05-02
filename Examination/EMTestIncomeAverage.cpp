#include "stdafx.h"
#include "EMTestIncomeAverage.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMTestIncomeAverage *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMTestIncomeAverage* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMTestIncomeAverage *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMTestIncomeAverage *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMTestIncomeAverage *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMTestIncomeAverage *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMTestIncomeAverage *pVw = (CEMTestIncomeAverage *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMTestIncomeAverage *pVw = (CEMTestIncomeAverage *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMTestIncomeAverage *pVw = (CEMTestIncomeAverage *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMTestIncomeAverage*)pWnd)->OnAddEMStatisticsGeneralPatientReport();
} 
static int _OnEditEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMTestIncomeAverage*)pWnd)->OnEditEMStatisticsGeneralPatientReport();
} 
static int _OnDeleteEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMTestIncomeAverage*)pWnd)->OnDeleteEMStatisticsGeneralPatientReport();
} 
static int _OnSaveEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMTestIncomeAverage*)pWnd)->OnSaveEMStatisticsGeneralPatientReport();
} 
static int _OnCancelEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMTestIncomeAverage*)pWnd)->OnCancelEMStatisticsGeneralPatientReport();
} 
CEMTestIncomeAverage::CEMTestIncomeAverage(CWnd *pParent)
{
	m_szTitle = _T("Patients General Statistics");
	m_nDlgWidth = 420;
	m_nDlgHeight = 125;
	SetDefaultValues();
}
CEMTestIncomeAverage::~CEMTestIncomeAverage()
{
}
void CEMTestIncomeAverage::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 205, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 215, 30, 295, 55);
	m_wndReportPeriod.Create(this,300, 30, 410, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 415, 90);
	m_wndFromDate.Create(this,95, 60, 205, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 215, 60, 295, 85);
	m_wndToDate.Create(this,300, 60, 410, 85); 
	//m_wndPrint.Create(this, _T("&Print"), 250, 95, 330, 120);
	m_wndExport.Create(this, _T("&Export"), 335, 95, 415, 120);

}
void CEMTestIncomeAverage::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(false);
	m_wndReportPeriod.SetCheckValue(false);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(false);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(false);
	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);
}
void CEMTestIncomeAverage::OnSetWindowEvents(){
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
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
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
void CEMTestIncomeAverage::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
}
void CEMTestIncomeAverage::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CEMTestIncomeAverage::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMTestIncomeAverage::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMTestIncomeAverage::SetMode(int nMode){
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
 			EnableButtons(TRUE,0,1,2,3,-1); 
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
/*void CEMTestIncomeAverage::OnYearChange(){
	
} */
/*void CEMTestIncomeAverage::OnYearSetfocus(){
	
} */
/*void CEMTestIncomeAverage::OnYearKillfocus(){
	
} */
int CEMTestIncomeAverage::OnYearCheckValue(){
	return 0;
} 
void CEMTestIncomeAverage::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMTestIncomeAverage::OnReportPeriodSelendok(){
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
/*void CEMTestIncomeAverage::OnReportPeriodSetfocus(){
	
}*/
/*void CEMTestIncomeAverage::OnReportPeriodKillfocus(){
	
}*/
long CEMTestIncomeAverage::OnReportPeriodLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
		CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0)
	{
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
/*void CEMTestIncomeAverage::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMTestIncomeAverage::OnFromDateChange(){
	
} */
/*void CEMTestIncomeAverage::OnFromDateSetfocus(){
	
} */
/*void CEMTestIncomeAverage::OnFromDateKillfocus(){
	
} */
int CEMTestIncomeAverage::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMTestIncomeAverage::OnToDateChange(){
	
} */
/*void CEMTestIncomeAverage::OnToDateSetfocus(){
	
} */
/*void CEMTestIncomeAverage::OnToDateKillfocus(){
	
} */
int CEMTestIncomeAverage::OnToDateCheckValue(){
	return 0;
} 
void CEMTestIncomeAverage::OnPrintSelect(){
	
} 
void CEMTestIncomeAverage::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CString szSQL, tmpStr;
	CString szOldRoom, szNewRoom;
	int nTmp = 0;
	double nAve = 0;
	CRecord rs(&pMF->m_db), ors(&pMF->m_db);
	CStringArray arrObj;
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
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(5, 10);
	xls.SetColumnWidth(6, 10);
	xls.SetColumnWidth(7, 10);
	xls.SetColumnWidth(8, 10);
	xls.SetColumnWidth(9, 10);
	xls.SetColumnWidth(10, 10);
	xls.SetColumnWidth(11, 10);
	xls.SetColumnWidth(12, 10);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 1);
	xls.SetCellMergedColumns(nCol, nRow + 2, 1);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 10);
	xls.SetCellText(nCol, nRow + 3, _T("TH\x1ED0NG K\xCA TI\x1EC0N \x58\xC9T NGHI\x1EC6M TRUNG \x42\xCCNH TH\x45O \x42N"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 6);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	szSQL.Format(_T("SELECT ho_desc FROM hms_object ORDER BY cast(ho_id as integer)"));
	ors.ExecSQL(szSQL);
	while (!ors.IsEOF())
	{
		ors.GetValue(_T("ho_desc"), tmpStr);
		arrObj.Add(tmpStr);
		ors.MoveNext();
	}
	xls.SetCellText(nCol, nRow + 5, _T("Ph\xF2ng kh\xE1m"), 528386, true, 11);
	for (int i = 0; i < arrObj.GetCount(); i++)
		xls.SetCellText(nCol + i + 1, nRow + 5, arrObj.GetAt(i), 528386, true, 11);
	nRow = 5;
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("nme"), szNewRoom);
		if (szOldRoom != szNewRoom)
		{
			nRow++;
			szOldRoom = szNewRoom;
			xls.SetCellText(nCol, nRow, szNewRoom);
		}
		rs.GetValue(_T("obj"), nTmp);
		rs.GetValue(_T("ave"), nAve);
		FormatCurrency(nAve, tmpStr);
		xls.SetCellText(nCol + nTmp, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
		rs.MoveNext();		
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\THONG KE TIEN XET NGHIEM TRUNG BINH THEO BN.xls"));
} 
void CEMTestIncomeAverage::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CEMTestIncomeAverage::OnAddEMStatisticsGeneralPatientReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMTestIncomeAverage::OnEditEMStatisticsGeneralPatientReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMTestIncomeAverage::OnDeleteEMStatisticsGeneralPatientReport(){
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
 		OnCancelEMStatisticsGeneralPatientReport(); 
 	}
	return 0;
}
int CEMTestIncomeAverage::OnSaveEMStatisticsGeneralPatientReport(){
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
 		//OnEMStatisticsGeneralPatientReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMTestIncomeAverage::OnCancelEMStatisticsGeneralPatientReport(){
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
int CEMTestIncomeAverage::OnEMStatisticsGeneralPatientReportListLoadData()
{
	return 0;
}
CString CEMTestIncomeAverage::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString szSQL, szWhere, szWhere1;
	szWhere.Format(_T(" AND hpc_status IN ('P', 'S', 'T') AND hpc_deptid = '%s'") \
				   _T(" AND hpc_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"),
				   pMF->GetCurrentDepartmentID(), m_szFromDate, m_szToDate);
	szWhere1.Format(_T(" AND he_deptid = '%s' AND he_status IN ('P', 'T') AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
					, pMF->GetCurrentDepartmentID(), m_szFromDate, m_szToDate);
	szSQL.Format(_T(" SELECT hrl_name nme, ") \
				_T("         hd_object obj, ") \
				_T("         hfe_cost/ qty ave ") \
				_T(" FROM   (SELECT    hpc_roomid room_id, ") \
				_T("                   hd_object, ") \
				_T("				CASE WHEN hd_object = 7 THEN SUM(f.hfe_cost)") \
				_T("				ELSE CASE WHEN hd_object IN (1, 3, 8) THEN SUM(f.hfe_polprice*hfe_quantity)") \
				_T("				ELSE SUM(f.hfe_insprice*hfe_quantity) END END hfe_cost") \
				_T("         FROM      hmsv_paraclinic ") \
				_T("         LEFT JOIN hms_fee f ON (hpc_docno = hfe_docno AND hpc_orderid = hfe_orderid ) ") \
				_T("         LEFT JOIN hms_fee_list ON ( hfl_feeid = hfe_itemid ) ") \
				_T("         LEFT JOIN hms_doc ON ( hd_docno = hfe_docno ) ") \
				_T("		 LEFT JOIN hms_object ON (ho_id = hd_object)") \
				_T("         WHERE     1 = 1 AND hfe_group NOT IN ('B2200', 'B2300')") \
				_T("         AND hfe_class = 'E' AND hfl_hitech <> 'Y' %s") \
				_T("		 GROUP BY hpc_roomid, hd_object)") \
				_T(" LEFT JOIN (SELECT he_roomid room_id, hd_object, count(he_docno) qty ") \
				_T("			FROM hms_exam ") \
				_T("			LEFT JOIN hms_doc ON (he_docno = hd_docno) ") \
				_T("			WHERE 1 = 1 %s GROUP BY he_roomid, hd_object) USING (room_id, hd_object)") \
				_T(" LEFT JOIN hms_roomlist ON (hrl_deptid = '%s' AND room_id = hrl_id)") \
				_T(" ORDER  BY cast(substr(hrl_name, 1, 2) as integer)"), szWhere, szWhere1, pMF->GetCurrentDepartmentID());

_fmsg(_T("%s"), szSQL);
	return szSQL;
}
