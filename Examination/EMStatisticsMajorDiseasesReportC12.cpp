#include "stdafx.h"
#include "EMStatisticsMajorDiseasesReportC12.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMStatisticsMajorDiseasesReportC12* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnToDateCheckValue();
} 
static void _OnExamRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMStatisticsMajorDiseasesReportC12* )pWnd)->OnExamRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnExamRoomSelendokFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnExamRoomSelendok();
}
/*static void _OnExamRoomSetfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnExamRoomKillfocus();
}*/
/*static void _OnExamRoomKillfocusFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnExamRoomKillfocus();
}*/
static long _OnExamRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnExamRoomLoadData();
}
/*static void _OnExamRoomAddNewFnc(CWnd *pWnd){
	((CEMStatisticsMajorDiseasesReportC12 *)pWnd)->OnExamRoomAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMStatisticsMajorDiseasesReportC12 *pVw = (CEMStatisticsMajorDiseasesReportC12 *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMStatisticsMajorDiseasesReportC12 *pVw = (CEMStatisticsMajorDiseasesReportC12 *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMStatisticsMajorDiseasesReportC12 *pVw = (CEMStatisticsMajorDiseasesReportC12 *)pWnd;
	pVw->OnCloseSelect();
}
static void _OnSelectAllDiseasesSelectFnc(CWnd *pWnd){
	 ((CEMStatisticsMajorDiseasesReportC12*)pWnd)->OnSelectAllDiseasesSelect();
}
static int _OnAddEMStatisticsMajorDiseasesReportFnc(CWnd *pWnd){
	 return ((CEMStatisticsMajorDiseasesReportC12*)pWnd)->OnAddEMStatisticsMajorDiseasesReport();
} 
static int _OnEditEMStatisticsMajorDiseasesReportFnc(CWnd *pWnd){
	 return ((CEMStatisticsMajorDiseasesReportC12*)pWnd)->OnEditEMStatisticsMajorDiseasesReport();
} 
static int _OnDeleteEMStatisticsMajorDiseasesReportFnc(CWnd *pWnd){
	 return ((CEMStatisticsMajorDiseasesReportC12*)pWnd)->OnDeleteEMStatisticsMajorDiseasesReport();
} 
static int _OnSaveEMStatisticsMajorDiseasesReportFnc(CWnd *pWnd){
	 return ((CEMStatisticsMajorDiseasesReportC12*)pWnd)->OnSaveEMStatisticsMajorDiseasesReport();
} 
static int _OnCancelEMStatisticsMajorDiseasesReportFnc(CWnd *pWnd){
	 return ((CEMStatisticsMajorDiseasesReportC12*)pWnd)->OnCancelEMStatisticsMajorDiseasesReport();
} 
CEMStatisticsMajorDiseasesReportC12::CEMStatisticsMajorDiseasesReportC12(CWnd *pParent)
{
	m_nDlgWidth = 440;
	m_nDlgHeight = 190;
	m_szTitle = _T("Diseases Suffered Report");
	SetDefaultValues();
}
CEMStatisticsMajorDiseasesReportC12::~CEMStatisticsMajorDiseasesReportC12()
{
}
void CEMStatisticsMajorDiseasesReportC12::OnCreateComponents()
{
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 150);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndExamRoomLabel.Create(this, _T("Exam Room"), 10, 90, 90, 115);
	m_wndExamRoom.Create(this,95, 90, 425, 115); 
	m_wndSelectAllDiseases.Create(this, _T("\x35 nh\xF3m \x62\x1EC7nh"), 295, 120, 425, 145);
//	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 240, 155, 350, 180);
	m_wndExport.Create(this, _T("&Export"), 355, 155, 430, 180);
}
void CEMStatisticsMajorDiseasesReportC12::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	////m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	//m_wndExamRoom.SetCheckValue(true);
	m_wndExamRoom.LimitText(81);
	//m_wndExamRoom.SetCheckBox(TRUE);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndExamRoom.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 50);
	m_wndExamRoom.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

}
void CEMStatisticsMajorDiseasesReportC12::OnSetWindowEvents()
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
	m_wndExamRoom.SetEvent(WE_SELENDOK, _OnExamRoomSelendokFnc);
	//m_wndExamRoom.SetEvent(WE_SETFOCUS, _OnExamRoomSetfocusFnc);
	//m_wndExamRoom.SetEvent(WE_KILLFOCUS, _OnExamRoomKillfocusFnc);
	m_wndExamRoom.SetEvent(WE_SELCHANGE, _OnExamRoomSelectChangeFnc);
	m_wndExamRoom.SetEvent(WE_LOADDATA, _OnExamRoomLoadDataFnc);
	//m_wndExamRoom.SetEvent(WE_ADDNEW, _OnExamRoomAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	m_wndSelectAllDiseases.SetEvent(WE_CLICK, _OnSelectAllDiseasesSelectFnc);
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");	
	SetMode(VM_EDIT);
	//m_wndSelectAllDiseases.EnableWindow(false);
}
void CEMStatisticsMajorDiseasesReportC12::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndExamRoom.GetDlgCtrlID(), m_szExamRoomKey);
	DDX_Check(pDX, m_wndSelectAllDiseases.GetDlgCtrlID(), m_bSelectAllDiseases);
}
void CEMStatisticsMajorDiseasesReportC12::GetDataToScreen()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMStatisticsMajorDiseasesReportC12::GetScreenToData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMStatisticsMajorDiseasesReportC12::SetDefaultValues()
{
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szExamRoomKey.Empty();
	m_bSelectAllDiseases=FALSE;
}
int CEMStatisticsMajorDiseasesReportC12::SetMode(int nMode)
{
 		int nOldMode = GetMode(); 
 		CGuiView::SetMode(nMode); 
 		CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 		CString szSQL, tmpStr; 
 		CRecord rs(&pMF->m_db);
		tmpStr = pMF->GetSysDate();
  		switch(nMode)
		{ 
 		case VM_ADD: 
 			EnableControls(TRUE); 
 			EnableButtons(TRUE, 0, 1, 2, -1);
 			SetDefaultValues(); 
 			break; 
 		case VM_EDIT: 
 			EnableControls(TRUE); 
 			EnableButtons(TRUE, 0, 1, 2, -1);
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
/*void CEMStatisticsMajorDiseasesReportC12::OnYearChange(){
	
} */
/*void CEMStatisticsMajorDiseasesReportC12::OnYearSetfocus(){
	
} */
/*void CEMStatisticsMajorDiseasesReportC12::OnYearKillfocus(){
	
} */
int CEMStatisticsMajorDiseasesReportC12::OnYearCheckValue(){
	return 0;
} 
void CEMStatisticsMajorDiseasesReportC12::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMStatisticsMajorDiseasesReportC12::OnReportPeriodSelendok()
{
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
	if(nMonth == 13)
	{
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/03/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59" ), m_nYear,  dte.GetMonthLastDay());
	}
	if(nMonth == 14)
	{
		m_szFromDate.Format(_T("%.4d/04/01"), m_nYear);
		tmpStr.Format(_T("%.4d/06/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 15)
	{
		m_szFromDate.Format(_T("%.4d/07/01"), m_nYear);
		tmpStr.Format(_T("%.4d/09/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 16)
	{
		m_szFromDate.Format(_T("%.4d/10/01"), m_nYear);
		tmpStr.Format(_T("%.4d/10/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 17)
	{
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/12/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	UpdateData(false);
}
/*void CEMStatisticsMajorDiseasesReportC12::OnReportPeriodSetfocus(){
	
}*/
/*void CEMStatisticsMajorDiseasesReportC12::OnReportPeriodKillfocus(){
	
}*/
long CEMStatisticsMajorDiseasesReportC12::OnReportPeriodLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0)
	{
		szWhere.Format(_T(" WHERE hpr_idx=%ld "), ToInt(m_szReportPeriodKey));
	}
	m_wndReportPeriod.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT * FROM hms_period_report %s ORDER BY hpr_idx "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF())
	{ 
		m_wndReportPeriod.AddItems(
			rs.GetValue(_T("hpr_idx")), 
			rs.GetValue(_T("hpr_name")),			
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMStatisticsMajorDiseasesReportC12::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMStatisticsMajorDiseasesReportC12::OnFromDateChange(){
	
} */
/*void CEMStatisticsMajorDiseasesReportC12::OnFromDateSetfocus(){
	
} */
/*void CEMStatisticsMajorDiseasesReportC12::OnFromDateKillfocus(){
	
} */
int CEMStatisticsMajorDiseasesReportC12::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMStatisticsMajorDiseasesReportC12::OnToDateChange(){
	
} */
/*void CEMStatisticsMajorDiseasesReportC12::OnToDateSetfocus(){
	
} */
/*void CEMStatisticsMajorDiseasesReportC12::OnToDateKillfocus(){
	
} */
int CEMStatisticsMajorDiseasesReportC12::OnToDateCheckValue(){
	return 0;
} 
void CEMStatisticsMajorDiseasesReportC12::OnExamRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMStatisticsMajorDiseasesReportC12::OnExamRoomSelendok(){
	 
}
/*void CEMStatisticsMajorDiseasesReportC12::OnExamRoomSetfocus(){
	
}*/
/*void CEMStatisticsMajorDiseasesReportC12::OnExamRoomKillfocus(){
	
}*/
long CEMStatisticsMajorDiseasesReportC12::OnExamRoomLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();
	if(m_wndExamRoom.IsSearchKey() && !m_szExamRoomKey.IsEmpty())
	{
		szWhere.Format(_T("and hrl_id = %d"), ToInt(m_szExamRoomKey));
	}
	m_wndExamRoom.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("select hrl_name as name, hrl_id as id from hms_roomlist where hrl_deptid = '%s' %s Order by id"), pMF->m_szDept, szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndExamRoom.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMStatisticsMajorDiseasesReportC12::OnExamRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMStatisticsMajorDiseasesReportC12::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
		
} 
void CEMStatisticsMajorDiseasesReportC12::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CRecord rs1(&pMF->m_db);
	CString szSQL,szSQL1, szTemp, tmpStr;
	CExcel xls;

	if(m_bSelectAllDiseases)
	{
		szSQL.Format(_T(" SELECT  hi_icd, 'Nhom1' as nhom, ") \
					_T("         hi_name as ICDName,") \
					_T("         count(hd_docno) as total") \
					_T("  FROM ") \
					_T("  (	  ") \
					_T("   SELECT ") \
					_T("         distinct hd_docno,") \
					_T("         max(he_examdate) as he_examedate,") \
					_T("         hd_icd") \
					_T("   FROM hms_doc") \
					_T("   LEFT JOIN hms_exam ON (he_docno = hd_docno)") \
					_T("   WHERE hd_status = 'T' AND length(hd_icd) > 0 ") \
					_T("         AND (he_examdate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
					_T("         AND hd_admitdept = 'C1.2'") \
					_T("         and hd_icd in ('E10', 'O24.1', 'R30.2', 'R30.3', 'E11', 'P71.5', 'I10', 'I10.1', 'I10.2', 'I10.3', 'O10', '', 'I50', 'I50.3', 'I50.4', 'I21', 'I25.2')") \
					_T("   GROUP BY hd_docno, hd_icd") \
					_T("  ) Tbl") \
					_T("  LEFT JOIN hms_icd ON(hi_icd = hd_icd)") \
					_T("  GROUP BY hi_icd, hi_name") \
					_T(" union all") \
					_T(" SELECT  hi_icd, 'Nhom2' as nhom, ") \
					_T("         hi_name as ICDName,") \
					_T("         count(hd_docno) as total") \
					_T("  FROM ") \
					_T("  (	  ") \
					_T("   SELECT ") \
					_T("         distinct hd_docno,") \
					_T("         max(he_examdate) as he_examedate,") \
					_T("         hd_icd") \
					_T("   FROM hms_doc") \
					_T("   LEFT JOIN hms_exam ON (he_docno = hd_docno)") \
					_T("   WHERE hd_status = 'T' AND length(hd_icd) > 0 ") \
					_T("         AND (he_examdate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
					_T("         AND hd_admitdept = 'C1.2'") \
					_T("         AND hd_icd in ('M49.0', 'M90.0', 'A15', 'A15.0', 'A16.0', 'A16.10', 'A16.5', 'A18', 'A16.10', 'J20', 'J20.1', 'J20.2', 'J20.0', 'J20.4', 'J20.5', 'J40', 'J41', 'J41.0', 'J41.1', 'J41.8', 'J42', 'J02', 'J02.0', 'J02.8', 'J02.9', 'J31.2')") \
					_T("   GROUP BY hd_docno, hd_icd") \
					_T("  ) Tbl") \
					_T("  LEFT JOIN hms_icd ON(hi_icd = hd_icd)") \
					_T("  GROUP BY hi_icd, hi_name") \
					_T("  union all") \
					_T("  SELECT  hi_icd, 'Nhom3' as nhom, ") \
					_T("         hi_name as ICDName,") \
					_T("         count(hd_docno) as total") \
					_T("  FROM ") \
					_T("  (	  ") \
					_T("   SELECT ") \
					_T("         distinct hd_docno,") \
					_T("         max(he_examdate) as he_examedate,") \
					_T("         hd_icd") \
					_T("   FROM hms_doc") \
					_T("   LEFT JOIN hms_exam ON (he_docno = hd_docno)") \
					_T("   WHERE hd_status = 'T' AND length(hd_icd) > 0 ") \
					_T("         AND (he_examdate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
					_T("         AND hd_admitdept = 'C1.2'") \
					_T("         AND hd_icd in ('H60', 'H60.9', 'H60.2', 'H62.0', 'H62.4', 'H65.1', 'H66.0', 'H66.1', 'H66.3', 'H66.9', 'J00', 'J30', 'J30.0', 'J30.3', 'J31', 'J90', 'J90.1', 'J90.2', 'J12', 'J15', 'J15.1', 'J15.8', 'J15.9', 'J18.9', 'J18', 'J32', 'J32.0', 'J32.0', 'J32.2', 'J32.3', 'J01', 'J01.0', 'J01.1', 'K79.7', 'I84', 'I84.1', 'I84.2', 'I84.5', 'K51.2', 'D01.2', 'K80', 'K80.1', 'K80.2', 'D13.4', 'K59.0')") \
					_T("   GROUP BY hd_docno, hd_icd") \
					_T("  ) Tbl") \
					_T("  LEFT JOIN hms_icd ON(hi_icd = hd_icd)") \
					_T("  GROUP BY hi_icd, hi_name") \
					_T("  union all") \
					_T("   SELECT  hi_icd, 'Nhom4' as nhom, ") \
					_T("         hi_name as ICDName,") \
					_T("         count(hd_docno) as total") \
					_T("  FROM ") \
					_T("  (	  ") \
					_T("   SELECT ") \
					_T("         distinct hd_docno,") \
					_T("         max(he_examdate) as he_examedate,") \
					_T("         hd_icd") \
					_T("   FROM hms_doc") \
					_T("   LEFT JOIN hms_exam ON (he_docno = hd_docno)") \
					_T("   WHERE hd_status = 'T' AND length(hd_icd) > 0 ") \
					_T("         AND (he_examdate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
					_T("         AND hd_admitdept = 'C1.2'") \
					_T("         AND hd_icd in ('C09.2', 'C09.4', 'C09.5', 'C09.6', 'C09.7')") \
					_T("   GROUP BY hd_docno, hd_icd") \
					_T("  ) Tbl") \
					_T("  LEFT JOIN hms_icd ON(hi_icd = hd_icd)") \
					_T("  GROUP BY hi_icd, hi_name") \
					_T(" union all") \
					_T("    SELECT  hi_icd, 'Nhomkhac' as nhom, ") \
					_T("         hi_name as ICDName,") \
					_T("         count(hd_docno) as total") \
					_T("  FROM ") \
					_T("  (	  ") \
					_T("   SELECT ") \
					_T("         distinct hd_docno,") \
					_T("         max(he_examdate) as he_examedate,") \
					_T("         hd_icd") \
					_T("   FROM hms_doc") \
					_T("   LEFT JOIN hms_exam ON (he_docno = hd_docno)") \
					_T("   WHERE hd_status = 'T' AND length(hd_icd) > 0 ") \
					_T("         AND (he_examdate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
					_T("         AND hd_admitdept = 'C1.2'") \
					_T("         AND (hd_icd  not in ('E10', 'O24.1', 'R30.2', 'R30.3', 'E11', 'P71.5', 'I10', 'I10.1', 'I10.2', 'I10.3', 'O10', 'I50', 'I50.3', 'I50.4', 'I21', 'I25.2')") \
					_T("         and hd_icd  not in ('M49.0', 'M90.0', 'A15', 'A15.0', 'A16.0', 'A16.10', 'A16.5', 'A18', 'A16.10', 'J20', 'J20.1', 'J20.2', 'J20.0', 'J20.4', 'J20.5', 'J40', 'J41', 'J41.0', 'J41.1', 'J41.8', 'J42', 'J02', 'J02.0', 'J02.8', 'J02.9', 'J31.2')") \
					_T("         and hd_icd not in ('H60', 'H60.9', 'H60.2', 'H62.0', 'H62.4', 'H65.1', 'H66.0', 'H66.1', 'H66.3', 'H66.9', 'J00', 'J30', 'J30.0', 'J30.3', 'J31', 'J90', 'J90.1', 'J90.2', 'J12', 'J15', 'J15.1', 'J15.8', 'J15.9', 'J18.9', 'J18', 'J32', 'J32.0', 'J32.0', 'J32.2', 'J32.3', 'J01', 'J01.0', 'J01.1', 'K79.7', 'I84', 'I84.1', 'I84.2', 'I84.5', 'K51.2', 'D01.2', 'K80', 'K80.1', 'K80.2', 'D13.4', 'K59.0')") \
					_T("         and hd_icd  not in ('C09.2', 'C09.4', 'C09.5', 'C09.6', 'C09.7'))") \
					_T("   GROUP BY hd_docno, hd_icd") \
					_T("  ) Tbl") \
					_T("  LEFT JOIN hms_icd ON(hi_icd = hd_icd)") \
					_T("  GROUP BY hi_icd, hi_name") \
					_T("  ORDER BY nhom,  hi_icd"), m_szFromDate, m_szToDate, m_szFromDate, m_szToDate, m_szFromDate, m_szToDate,m_szFromDate, m_szToDate,m_szFromDate, m_szToDate);
		rs.ExecSQL(szSQL);
			xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 20);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 40);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 20);

	int nRow = 0, nCol = 0;
	xls.SetCellMergedColumns(nCol, nRow + 2, 4);

	xls.SetCellText(nCol+2, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol+2, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(nCol, nRow + 2, 4);
	xls.SetCellMergedColumns(nCol, nRow + 3, 4);
	xls.SetCellText(nCol, nRow + 2, _T("TH\x1ED0NG K\xCA M\x1EB6T \x42\x1EC6NH TH\x45O 5 NH\xD3M"), FMT_TEXT | FMT_CENTER, true, 13);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), 
		          CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss),
				  CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellText(nCol, nRow + 4, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 4, _T("M\xE3 \x62\x1EC7nh"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 4, _T("T\xEAn m\x1EB7t \x62\x1EC7nh"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 3, nRow + 4, _T("S\x1ED1 l\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER, true, 11); 

	//xls.SetCellText(0, 5, _T("Qu\xE2n:"), FMT_TEXT, true, 12);
	//rs1.GetValue(_T("quan"), tmpStr);
	//xls.SetCellText(1, 5, tmpStr, FMT_TEXT, true, 12);

	//rs1.GetValue(_T("huu"), tmpStr);
	//szTemp.Format(_T("H\x1B0u: %s"),tmpStr);
	//xls.SetCellText(2, 5, szTemp, FMT_TEXT, true, 12);

	//rs1.GetValue(_T("bhytquan"), tmpStr);
	//szTemp.Format(_T("\x42HYT Qu\xE2n: %s"),tmpStr);
	//xls.SetCellText(3, 5, szTemp, FMT_TEXT, true, 12);

	//rs1.GetValue(_T("tongdoituong"), tmpStr);
	//szTemp.Format(_T("T\x1ED5ng: %s"),tmpStr);
	//xls.SetCellText(4, 5, szTemp, FMT_TEXT, true, 12);

	//xls.SetCellText(0, 6, _T("VV qu\xE2n:"), FMT_TEXT, true, 12);
	//rs1.GetValue(_T("vaoquan"), tmpStr);
	//xls.SetCellText(1, 6, tmpStr, FMT_TEXT, true, 12);

	//rs1.GetValue(_T("vaobhytquan"), tmpStr);
	//szTemp.Format(_T("VV H\x1B0u: %s"),tmpStr);
	//xls.SetCellText(2, 6, szTemp, FMT_TEXT, true, 12);

	//rs1.GetValue(_T("vaoquancothe"), tmpStr);
	//szTemp.Format(_T("VV \x42HYT Qu\xE2n: %s"),tmpStr);
	//xls.SetCellText(3, 6, szTemp, FMT_TEXT, true, 12);

	//rs1.GetValue(_T("tongvao"), tmpStr);
	//szTemp.Format(_T("VV T\x1ED5ng: %s"),tmpStr);
	//xls.SetCellText(4, 6, szTemp, FMT_TEXT, true, 12);

	nRow += 6;
	int nIndex = 1;
	int nItem = 1;
	long double nTotal = 0.0;
	long double grpCost=0.0;
	long double ttlCost=0.0;

	CString szNhomName;
	double cost = 0;
	CString szNewLine, szOldLine;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("nhom"), szNewLine);
	
		if(szNewLine != szOldLine && !szNewLine.IsEmpty())
		{	
			if(nTotal > 0)
			{			
				CString szField, szAmount;
				nRow++;
				TranslateString(_T("T\x1ED5ng nh\xF3m:"), szAmount);												
				xls.SetCellMergedColumns(0,nRow,3);
				xls.SetCellText(0, nRow, szAmount, FMT_TEXT,true,11,0);
				tmpStr.Format(_T("%.2f"), nTotal);
				xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1,true,11,0);
				ttlCost += nTotal;
				nTotal = 0;								
				nItem=1;
			}
			nRow++;
			CString szField, szAmount;			
			xls.SetCellMergedColumns(0,nRow,4);	
		
			rs.GetValue(_T("nhom"), szNhomName);
			if(szNhomName == _T("Nhom1")) 
			{
				xls.SetCellText(0, nRow, _T("Tim m\x1EA1\x63h, \x62\x1EC7nh \x63huy\x1EC3n h\xF3\x61"), FMT_TEXT,true,11,0);
				
			}
			else if(szNhomName == _T("Nhom2"))  
			{
				xls.SetCellText(0, nRow, _T("\x42\x1EC7nh \x63\x1A1 qu\x61n h\xF4 h\x1EA5p"), FMT_TEXT,true,11,0);
					
			}
			else if(szNhomName == _T("Nhom3"))  xls.SetCellText(0, nRow, _T("\x42\x1EC7nh \x63\x1A1 qu\x61n ti\xEAu h\xF3\x61"), FMT_TEXT,true,11,0);
			else if(szNhomName == _T("Nhom4"))  xls.SetCellText(0, nRow, _T("Ung th\x1B0 \x63\xE1\x63 lo\x1EA1i"), FMT_TEXT,true,11,0);
			else if(szNhomName == _T("Nhomkhac"))  xls.SetCellText(0, nRow, _T("\x43\xE1\x63 \x62\x1EC7nh kh\xE1\x63"), FMT_TEXT,true,11,0);
	



			szOldLine = szNewLine;
			nItem=1;
		}
		nRow++;
		szTemp.Format(_T("%d"), nItem);
		xls.SetCellText(nCol, nRow, szTemp, FMT_INTEGER);
		
		szTemp = rs.GetValue(_T("icd"));
		xls.SetCellText(nCol + 1, nRow, szTemp, FMT_TEXT);

		szTemp = rs.GetValue(_T("ICDName"));
		xls.SetCellText(nCol + 2, nRow, szTemp, FMT_TEXT);

		szTemp = rs.GetValue(_T("total"));
		nTotal += ToDouble(szTemp);
		xls.SetCellText(nCol + 3, nRow, szTemp, FMT_INTEGER);
		nItem++;
		rs.MoveNext();
	}
	ttlCost += nTotal;
	if(nTotal > 0)
	{
		nRow++;
		TranslateString(_T("T\x1ED5ng nh\xF3m:"), tmpStr);		
		xls.SetCellMergedColumns(0,nRow,3);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);
		tmpStr.Format(_T("%.2f"), nTotal);
		xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1,true,11,0);		
	}
	if(ttlCost > 0)
	{
		nRow++;
		TranslateString(_T("Total"), tmpStr);		
		xls.SetCellMergedColumns(0,nRow,3);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);

		tmpStr.Format(_T("%.2f"), ttlCost);
		xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1,true,11,0);		
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\THONG KE THEO 5 NHOM BENH.XLS"));

	}
	else if(m_bSelectAllDiseases == false)
	{
	szSQL = GetQueryString();
	
	
	szSQL1.Format(_T(" SELECT  ") \
				_T("          sum(quan) as quan, sum(huu) as huu, sum(bhytquan) as bhytquan, sum(tongdoituong) as tongdoituong,") \
				_T("          sum(tongvao) as tongvao,") \
				_T("          sum(vaoquan) as vaoquan, sum(vaobhytquan) as vaobhytquan, sum(vaoquancothe) as vaoquancothe") \
				_T("          ") \
				_T("  FROM   (	SELECT     ") \
				_T("                distinct hd_docno,        ") \
				_T("                max(he_examdate) as he_examedate,") \
				_T("               hd_icd as icd, hd_object, ") \
				_T("                case when hd_object in ('1', '2', '11') then 1 else 0 end as tongdoituong, ") \
				_T("                case when hd_object = '1' then 1 else 0 end as quan,") \
				_T("                case when hd_object = '2' then 1 else 0 end as huu,") \
				_T("                case when hd_object = '11' then 1 else 0 end as bhytquan,") \
				_T("                case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('1', '2', '11') then 1 else 0 end as tongvao,") \
				_T("                case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('1') then 1 else 0 end as vaoquan,   ") \
				_T("                case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('2')  then 1 else 0 end as vaobhytquan,   ") \
				_T("                case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('11')then 1 else 0 end as vaoquancothe") \
				_T("            FROM hms_doc ") \
				_T("            LEFT JOIN hms_exam ON (he_docno = hd_docno) ") \
				_T("            WHERE hd_status = 'T' AND length(hd_icd) > 0    ") \
				_T("                  AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
				_T("            AND hd_admitdept = 'C1.2' ") \
				_T("            GROUP BY hd_docno, hd_icd, hd_object , hd_suggestion, hd_indept") \
				_T("      )") \
				_T("  LEFT JOIN hms_icd ON(hi_icd = icd)") \
				_T("  LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T("  LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") , m_szFromDate, m_szToDate);


	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	rs1.ExecSQL(szSQL1);
	_fmsg(_T("'%s'"), szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 20);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 40);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 20);

	int nRow = 0, nCol = 0;
	xls.SetCellMergedColumns(nCol, nRow + 2, 4);

	xls.SetCellText(nCol+2, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol+2, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(nCol, nRow + 2, 4);
	xls.SetCellMergedColumns(nCol, nRow + 3, 4);
	xls.SetCellText(nCol, nRow + 2, _T("TH\x1ED0NG K\xCA M\x1EB6T \x42\x1EC6NH TH\x45O NH\xD3M"), FMT_TEXT | FMT_CENTER, true, 13);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), 
		          CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss),
				  CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellText(nCol, nRow + 4, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 4, _T("M\xE3 \x62\x1EC7nh"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 4, _T("T\xEAn m\x1EB7t \x62\x1EC7nh"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 3, nRow + 4, _T("S\x1ED1 l\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER, true, 11); 

	xls.SetCellText(0, 5, _T("Qu\xE2n:"), FMT_TEXT, true, 12);
	rs1.GetValue(_T("quan"), tmpStr);
	xls.SetCellText(1, 5, tmpStr, FMT_TEXT, true, 12);

	rs1.GetValue(_T("huu"), tmpStr);
	szTemp.Format(_T("H\x1B0u: %s"),tmpStr);
	xls.SetCellText(2, 5, szTemp, FMT_TEXT, true, 12);

	rs1.GetValue(_T("bhytquan"), tmpStr);
	szTemp.Format(_T("\x42HYT Qu\xE2n: %s"),tmpStr);
	xls.SetCellText(3, 5, szTemp, FMT_TEXT, true, 12);

	rs1.GetValue(_T("tongdoituong"), tmpStr);
	szTemp.Format(_T("T\x1ED5ng: %s"),tmpStr);
	xls.SetCellText(4, 5, szTemp, FMT_TEXT, true, 12);

	xls.SetCellText(0, 6, _T("VV qu\xE2n:"), FMT_TEXT, true, 12);
	rs1.GetValue(_T("vaoquan"), tmpStr);
	xls.SetCellText(1, 6, tmpStr, FMT_TEXT, true, 12);

	rs1.GetValue(_T("vaobhytquan"), tmpStr);
	szTemp.Format(_T("VV H\x1B0u: %s"),tmpStr);
	xls.SetCellText(2, 6, szTemp, FMT_TEXT, true, 12);

	rs1.GetValue(_T("vaoquancothe"), tmpStr);
	szTemp.Format(_T("VV \x42HYT Qu\xE2n: %s"),tmpStr);
	xls.SetCellText(3, 6, szTemp, FMT_TEXT, true, 12);

	rs1.GetValue(_T("tongvao"), tmpStr);
	szTemp.Format(_T("VV T\x1ED5ng: %s"),tmpStr);
	xls.SetCellText(4, 6, szTemp, FMT_TEXT, true, 12);

	nRow += 6;
	int nIndex = 1;
	int nItem = 1;
	long double nTotal = 0.0;
	long double grpCost=0.0;
	long double ttlCost=0.0;

	CString szNameChapter;
	double cost = 0;
	CString szNewLine, szOldLine;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("hic_id"), szNewLine);
		rs.GetValue(_T("hic_name"), szNameChapter);
		if(szNewLine != szOldLine && !szNewLine.IsEmpty())
		{	
			if(nTotal > 0)
			{			
				CString szField, szAmount;
				nRow++;
				TranslateString(_T("T\x1ED5ng nh\xF3m:"), szAmount);												
				xls.SetCellMergedColumns(0,nRow,3);
				xls.SetCellText(0, nRow, szAmount, FMT_TEXT,true,11,0);
				tmpStr.Format(_T("%.2f"), nTotal);
				xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1,true,11,0);
				ttlCost += nTotal;
				nTotal = 0;								
				nItem=1;
			}
			nRow++;
			CString szField, szAmount;			
			xls.SetCellMergedColumns(0,nRow,4);	
			xls.SetCellText(0, nRow, szNameChapter, FMT_TEXT,true,11,0);
			szOldLine = szNewLine;
			nItem=1;
		}
		nRow++;
		szTemp.Format(_T("%d"), nItem);
		xls.SetCellText(nCol, nRow, szTemp, FMT_INTEGER);
		
		szTemp = rs.GetValue(_T("icd"));
		xls.SetCellText(nCol + 1, nRow, szTemp, FMT_TEXT);

		szTemp = rs.GetValue(_T("ICDName"));
		xls.SetCellText(nCol + 2, nRow, szTemp, FMT_TEXT);

		szTemp = rs.GetValue(_T("total"));
		nTotal += ToDouble(szTemp);
		xls.SetCellText(nCol + 3, nRow, szTemp, FMT_INTEGER);
		nItem++;
		rs.MoveNext();
	}
	ttlCost += nTotal;
	if(nTotal > 0)
	{
		nRow++;
		TranslateString(_T("T\x1ED5ng nh\xF3m:"), tmpStr);		
		xls.SetCellMergedColumns(0,nRow,3);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);
		tmpStr.Format(_T("%.2f"), nTotal);
		xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1,true,11,0);		
	}
	if(ttlCost > 0)
	{
		nRow++;
		TranslateString(_T("Total"), tmpStr);		
		xls.SetCellMergedColumns(0,nRow,3);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);

		tmpStr.Format(_T("%.2f"), ttlCost);
		xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1,true,11,0);		
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\THONG KE THEO NHOM BENH.XLS"));
	}
} 
void CEMStatisticsMajorDiseasesReportC12::OnCloseSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CEMStatisticsMajorDiseasesReportC12::OnSelectAllDiseasesSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
int CEMStatisticsMajorDiseasesReportC12::OnAddEMStatisticsMajorDiseasesReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMStatisticsMajorDiseasesReportC12::OnEditEMStatisticsMajorDiseasesReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMStatisticsMajorDiseasesReportC12::OnDeleteEMStatisticsMajorDiseasesReport(){
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
 		OnCancelEMStatisticsMajorDiseasesReport(); 
 	}
	return 0;
}
int CEMStatisticsMajorDiseasesReportC12::OnSaveEMStatisticsMajorDiseasesReport(){
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
 		//OnEMStatisticsMajorDiseasesReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMStatisticsMajorDiseasesReportC12::OnCancelEMStatisticsMajorDiseasesReport(){
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
int CEMStatisticsMajorDiseasesReportC12::OnEMStatisticsMajorDiseasesReportListLoadData(){
	return 0;
}

CString CEMStatisticsMajorDiseasesReportC12::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szSQL1, szWhere, szCondition;
	szWhere.Empty();

	if (!m_szExamRoomKey.IsEmpty())
	{
		szWhere.Format(_T(" and he_roomid = %d "), ToInt(m_szExamRoomKey));
	}

	szSQL.Format(_T(" SELECT  icd,        ") \
				_T("         hi_name as ICDName,    ") \
				_T("         count(hd_docno) as total,   hic_name,   hic_id, hig_name, hig_name") \
				_T(" FROM   (	SELECT      ") \
				_T("               distinct hd_docno,         ") \
				_T("               max(he_examdate) as he_examedate, ") \
				_T("               hd_icd as icd  ") \
				_T("           FROM hms_doc  ") \
				_T("           LEFT JOIN hms_exam ON (he_docno = hd_docno)  ") \
				_T("           WHERE hd_status = 'T' AND length(hd_icd) > 0     ") \
				_T("                 AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T("           AND hd_admitdept = 'C1.2'  ") \
				_T("           GROUP BY hd_docno, hd_icd  ") \
				_T("     ) ") \
				_T(" LEFT JOIN hms_icd ON(hi_icd = icd) ") \
				_T(" LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T("  LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") \
				_T(" GROUP BY icd, hi_name, hic_id, hig_name,hic_name ") \
				_T(" ORDER BY hic_name, hic_id") , m_szFromDate, m_szToDate, szWhere);

	return szSQL;
}