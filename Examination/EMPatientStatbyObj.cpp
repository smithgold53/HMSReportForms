#include "stdafx.h"
#include "EMPatientStatbyObj.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMPatientStatbyObj *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMPatientStatbyObj* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMPatientStatbyObj *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMPatientStatbyObj *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyObj *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMPatientStatbyObj *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMPatientStatbyObj *pVw = (CEMPatientStatbyObj *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMPatientStatbyObj *pVw = (CEMPatientStatbyObj *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMPatientStatbyObj *pVw = (CEMPatientStatbyObj *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMPatientStatbyObj*)pWnd)->OnAddEMStatisticsGeneralPatientReport();
} 
static int _OnEditEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMPatientStatbyObj*)pWnd)->OnEditEMStatisticsGeneralPatientReport();
} 
static int _OnDeleteEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMPatientStatbyObj*)pWnd)->OnDeleteEMStatisticsGeneralPatientReport();
} 
static int _OnSaveEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMPatientStatbyObj*)pWnd)->OnSaveEMStatisticsGeneralPatientReport();
} 
static int _OnCancelEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMPatientStatbyObj*)pWnd)->OnCancelEMStatisticsGeneralPatientReport();
} 
CEMPatientStatbyObj::CEMPatientStatbyObj(CWnd *pParent)
{
	m_szTitle = _T("Patients General Statistics");
	m_nDlgWidth = 420;
	m_nDlgHeight = 125;
	SetDefaultValues();
}
CEMPatientStatbyObj::~CEMPatientStatbyObj()
{
}
void CEMPatientStatbyObj::OnCreateComponents(){
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
void CEMPatientStatbyObj::OnInitializeComponents(){
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
void CEMPatientStatbyObj::OnSetWindowEvents(){
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
void CEMPatientStatbyObj::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
}
void CEMPatientStatbyObj::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CEMPatientStatbyObj::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMPatientStatbyObj::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMPatientStatbyObj::SetMode(int nMode){
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
/*void CEMPatientStatbyObj::OnYearChange(){
	
} */
/*void CEMPatientStatbyObj::OnYearSetfocus(){
	
} */
/*void CEMPatientStatbyObj::OnYearKillfocus(){
	
} */
int CEMPatientStatbyObj::OnYearCheckValue(){
	return 0;
} 
void CEMPatientStatbyObj::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMPatientStatbyObj::OnReportPeriodSelendok(){
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
/*void CEMPatientStatbyObj::OnReportPeriodSetfocus(){
	
}*/
/*void CEMPatientStatbyObj::OnReportPeriodKillfocus(){
	
}*/
long CEMPatientStatbyObj::OnReportPeriodLoadData()
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
/*void CEMPatientStatbyObj::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMPatientStatbyObj::OnFromDateChange(){
	
} */
/*void CEMPatientStatbyObj::OnFromDateSetfocus(){
	
} */
/*void CEMPatientStatbyObj::OnFromDateKillfocus(){
	
} */
int CEMPatientStatbyObj::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMPatientStatbyObj::OnToDateChange(){
	
} */
/*void CEMPatientStatbyObj::OnToDateSetfocus(){
	
} */
/*void CEMPatientStatbyObj::OnToDateKillfocus(){
	
} */
int CEMPatientStatbyObj::OnToDateCheckValue(){
	return 0;
} 
void CEMPatientStatbyObj::OnPrintSelect(){}

void CEMPatientStatbyObj::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CString szSQL, tmpStr;
	CRecord rs(&pMF->m_db);
	CStringArray arrCol;
	int nTmp = 0;
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
	xls.SetColumnWidth(1, 20);


	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 1);
	xls.SetCellMergedColumns(nCol, nRow + 2, 1);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 2);
	xls.SetCellText(nCol, nRow + 3, _T("TH\x1ED0NG K\xCA \x42\x1EC6NH NH\xC2N TH\x45O \x43H\x1EBE \x110\x1ED8"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 2);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 5, _T("\x43h\x1EBF \x111\x1ED9 kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 5, _T("S\x1ED1 l\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER, true, 11);
	nRow=6;
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("obj"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("cnt"), nTmp);
		FormatCurrency(nTmp, tmpStr);
		xls.SetCellText(nCol+1, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
		nRow++;
		rs.MoveNext();
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\THONG KE BENH NHAN THEO CHE DO.xls"));
} 
void CEMPatientStatbyObj::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CEMPatientStatbyObj::OnAddEMStatisticsGeneralPatientReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMPatientStatbyObj::OnEditEMStatisticsGeneralPatientReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMPatientStatbyObj::OnDeleteEMStatisticsGeneralPatientReport(){
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
int CEMPatientStatbyObj::OnSaveEMStatisticsGeneralPatientReport(){
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
int CEMPatientStatbyObj::OnCancelEMStatisticsGeneralPatientReport(){
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
int CEMPatientStatbyObj::OnEMStatisticsGeneralPatientReportListLoadData(){
	return 0;
}
CString CEMPatientStatbyObj::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND hd_admitdept = '%s'"), pMF->GetCurrentDepartmentID());
	szWhere.AppendFormat(_T(" AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szSQL.Format(_T(" SELECT") \
				_T("   Get_objectname(hd_object) AS obj,") \
				_T("   Count(hd_docno)           AS cnt") \
				_T(" FROM   hms_doc") \
				_T(" LEFT JOIN hms_exam ON (hd_docno=he_docno)") \
				_T(" WHERE  hd_object>0") \
				_T("        AND he_status IN ('P', 'T') %s") \
				_T(" GROUP  BY hd_object") \
				_T(" ORDER  BY Cast(hd_object AS INTEGER)"), szWhere);
	return szSQL;
}
