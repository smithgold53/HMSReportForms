#include "stdafx.h"
#include "EMObject11DiseaseStatisticsReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMObject11DiseaseStatisticsReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){ 
	((CEMObject11DiseaseStatisticsReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMObject11DiseaseStatisticsReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMObject11DiseaseStatisticsReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMObject11DiseaseStatisticsReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnExamRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMObject11DiseaseStatisticsReport* )pWnd)->OnExamRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnExamRoomSelendokFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnExamRoomSelendok();
}
/*static void _OnExamRoomSetfocusFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnExamRoomKillfocus();
}*/
/*static void _OnExamRoomKillfocusFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnExamRoomKillfocus();
}*/
static long _OnExamRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMObject11DiseaseStatisticsReport *)pWnd)->OnExamRoomLoadData();
}
/*static void _OnExamRoomAddNewFnc(CWnd *pWnd){
	((CEMObject11DiseaseStatisticsReport *)pWnd)->OnExamRoomAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMObject11DiseaseStatisticsReport *pVw = (CEMObject11DiseaseStatisticsReport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMObject11DiseaseStatisticsReport *pVw = (CEMObject11DiseaseStatisticsReport *)pWnd;
	pVw->OnExportSelect();
} 
 
static int _OnAddEMObject11DiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMObject11DiseaseStatisticsReport*)pWnd)->OnAddEMObject11DiseaseStatisticsReport();
} 
static int _OnEditEMObject11DiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMObject11DiseaseStatisticsReport*)pWnd)->OnEditEMObject11DiseaseStatisticsReport();
} 
static int _OnDeleteEMObject11DiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMObject11DiseaseStatisticsReport*)pWnd)->OnDeleteEMObject11DiseaseStatisticsReport();
} 
static int _OnSaveEMObject11DiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMObject11DiseaseStatisticsReport*)pWnd)->OnSaveEMObject11DiseaseStatisticsReport();
} 
static int _OnCancelEMObject11DiseaseStatisticsReportFnc(CWnd *pWnd){
	 return ((CEMObject11DiseaseStatisticsReport*)pWnd)->OnCancelEMObject11DiseaseStatisticsReport();
} 
CEMObject11DiseaseStatisticsReport::CEMObject11DiseaseStatisticsReport(CWnd *pParent)
{
	m_nDlgWidth = 445;
	m_nDlgHeight = 155;
	m_szTitle = _T("Statistic ICD - Object");
	SetDefaultValues();
}
CEMObject11DiseaseStatisticsReport::~CEMObject11DiseaseStatisticsReport(){
}
void CEMObject11DiseaseStatisticsReport::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 120);
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
//	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 240, 125, 350, 150);
	m_wndExport.Create(this, _T("&Export"), 355, 125, 430, 150);

}
void CEMObject11DiseaseStatisticsReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	//m_wndExamRoom.SetCheckValue(true);
	m_wndExamRoom.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);


	m_wndExamRoom.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndExamRoom.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

}
void CEMObject11DiseaseStatisticsReport::OnSetWindowEvents(){
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

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	SetMode(VM_EDIT);
}
void CEMObject11DiseaseStatisticsReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndExamRoom.GetDlgCtrlID(), m_szExamRoomKey);
}
void CEMObject11DiseaseStatisticsReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMObject11DiseaseStatisticsReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMObject11DiseaseStatisticsReport::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szExamRoomKey.Empty();

}
int CEMObject11DiseaseStatisticsReport::SetMode(int nMode){
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
/*void CEMObject11DiseaseStatisticsReport::OnYearChange(){
	
} */
/*void CEMObject11DiseaseStatisticsReport::OnYearSetfocus(){
	
} */
/*void CEMObject11DiseaseStatisticsReport::OnYearKillfocus(){
	
} */
int CEMObject11DiseaseStatisticsReport::OnYearCheckValue(){
	return 0;
} 
void CEMObject11DiseaseStatisticsReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMObject11DiseaseStatisticsReport::OnReportPeriodSelendok(){
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
/*void CEMObject11DiseaseStatisticsReport::OnReportPeriodSetfocus(){
	
}*/
/*void CEMObject11DiseaseStatisticsReport::OnReportPeriodKillfocus(){
	
}*/
long CEMObject11DiseaseStatisticsReport::OnReportPeriodLoadData(){
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
/*void CEMObject11DiseaseStatisticsReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMObject11DiseaseStatisticsReport::OnFromDateChange(){
	
} */
/*void CEMObject11DiseaseStatisticsReport::OnFromDateSetfocus(){
	
} */
/*void CEMObject11DiseaseStatisticsReport::OnFromDateKillfocus(){
	
} */
int CEMObject11DiseaseStatisticsReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMObject11DiseaseStatisticsReport::OnToDateChange(){
	
} */
/*void CEMObject11DiseaseStatisticsReport::OnToDateSetfocus(){
	
} */
/*void CEMObject11DiseaseStatisticsReport::OnToDateKillfocus(){
	
} */
int CEMObject11DiseaseStatisticsReport::OnToDateCheckValue(){
	return 0;
} 
void CEMObject11DiseaseStatisticsReport::OnExamRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMObject11DiseaseStatisticsReport::OnExamRoomSelendok(){
	 
}
/*void CEMObject11DiseaseStatisticsReport::OnExamRoomSetfocus(){
	
}*/
/*void CEMObject11DiseaseStatisticsReport::OnExamRoomKillfocus(){
	
}*/
long CEMObject11DiseaseStatisticsReport::OnExamRoomLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndExamRoom.IsSearchKey() && !m_szExamRoomKey.IsEmpty()){
		szWhere.Format(_T("and hrl_id = %d"), ToInt(m_szExamRoomKey));
	};
	m_wndExamRoom.DeleteAllItems(); 
	szSQL.Format(_T("select hrl_name as name, hrl_id as id from hms_roomlist where hrl_deptid = '%s' %s Order by id"), pMF->m_szDept, szWhere);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndExamRoom.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMObject11DiseaseStatisticsReport::OnExamRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMObject11DiseaseStatisticsReport::OnPrintPreviewSelect()
{
	
} 
void CEMObject11DiseaseStatisticsReport::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	BeginWaitCursor();
	UpdateData(true);
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 54);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(5, 10);
	xls.SetColumnWidth(6, 10);
	xls.SetColumnWidth(7, 10);
	xls.SetColumnWidth(8, 10);
	xls.SetColumnWidth(9, 10);
	xls.SetColumnWidth(10, 10);
	xls.SetColumnWidth(11, 10);
	xls.SetColumnWidth(12, 13);
	int nCol = 0;
	int nRow = 0;

if (!m_szExamRoomKey.IsEmpty())
	{
		szWhere.Format(_T(" and he_roomid = %d "), ToInt(m_szExamRoomKey));
	}

	szSQL.Format(_T(" SELECT  icd, 'nhom1' as phanloainhom,        ") \
				_T("        hi_name as ICDName,    ") \
				_T("        count(hd_docno) as total, sum(tongdoituong) as tongdoituong, sum(quan) as quan, sum(huu) as huu, sum(bhytquan) as bhytquan, ") \
				_T("        sum(tongvao) as tongvao, sum(vaoquan) as vaoquan, sum(vaobhytquan) as vaobhytquan, sum(vaoquancothe) as vaoquancothe") \
				_T(" FROM   (	SELECT      ") \
				_T("              distinct hd_docno,         ") \
				_T("              max(he_examdate) as he_examedate, ") \
				_T("              hd_icd as icd, ") \
				_T("             case when hd_object in ('1', '2', '11') then 1 else 0 end as tongdoituong, ") \
				_T("             case when hd_object = '1' then 1 else 0 end as quan,") \
				_T("             case when hd_object = '2' then 1 else 0 end as huu,") \
				_T("             case when hd_object = '11' then 1 else 0 end as bhytquan,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('1', '2', '11') then 1 else 0 end as tongvao,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('1') then 1 else 0 end as vaoquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('2')  then 1 else 0 end as vaobhytquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('11')then 1 else 0 end as vaoquancothe") \
				_T("          FROM hms_doc  ") \
				_T("          LEFT JOIN hms_exam ON (he_docno = hd_docno)  ") \
				_T("          WHERE hd_status = 'T' AND length(hd_icd) > 0     ") \
				_T("                AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T("          AND hd_admitdept = 'C1.2'  ") \
				_T("          AND hd_icd in ('E75.5','E78', 'E78.2', 'E78.5')") \
				_T("          GROUP BY hd_docno, hd_icd, hd_object, hd_suggestion, hd_indept") \
				_T("    ) ") \
				_T(" LEFT JOIN hms_icd ON(hi_icd = icd) ") \
				_T(" LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T(" LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") \
				_T(" GROUP BY icd, hi_name") \
				_T(" union all ") \
				_T(" SELECT  icd, 'nhom2' as phanloainhom,        ") \
				_T("        hi_name as ICDName,    ") \
				_T("        count(hd_docno) as total, sum(tongdoituong) as tongdoituong, sum(quan) as quan, sum(huu) as huu, sum(bhytquan) as bhytquan, ") \
				_T("        sum(tongvao) as tongvao, sum(vaoquan) as vaoquan, sum(vaobhytquan) as vaobhytquan, sum(vaoquancothe) as vaoquancothe") \
				_T(" FROM   (	SELECT      ") \
				_T("              distinct hd_docno,         ") \
				_T("              max(he_examdate) as he_examedate, ") \
				_T("              hd_icd as icd, ") \
				_T("             case when hd_object in ('1', '2', '11') then 1 else 0 end as tongdoituong, ") \
				_T("             case when hd_object = '1' then 1 else 0 end as quan,") \
				_T("             case when hd_object = '2' then 1 else 0 end as huu,") \
				_T("             case when hd_object = '11' then 1 else 0 end as bhytquan,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('1', '2', '11') then 1 else 0 end as tongvao,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('1') then 1 else 0 end as vaoquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('2')  then 1 else 0 end as vaobhytquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('11')then 1 else 0 end as vaoquancothe") \
				_T("          FROM hms_doc  ") \
				_T("          LEFT JOIN hms_exam ON (he_docno = hd_docno)  ") \
				_T("          WHERE hd_status = 'T' AND length(hd_icd) > 0     ") \
				_T("                AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T("          AND hd_admitdept = 'C1.2'  ") \
				_T("          AND hd_icd in ('I10','I10.1', 'I10.2', 'I10.3')") \
				_T("           GROUP BY hd_docno, hd_icd, hd_object, hd_suggestion, hd_indept") \
				_T("    ) ") \
				_T(" LEFT JOIN hms_icd ON(hi_icd = icd) ") \
				_T(" LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T(" LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") \
				_T(" GROUP BY icd, hi_name") \
				_T(" union all ") \
				_T(" SELECT  icd, 'nhom3' as phanloainhom,        ") \
				_T("        hi_name as ICDName,    ") \
				_T("        count(hd_docno) as total, sum(tongdoituong) as tongdoituong, sum(quan) as quan, sum(huu) as huu, sum(bhytquan) as bhytquan, ") \
				_T("        sum(tongvao) as tongvao, sum(vaoquan) as vaoquan, sum(vaobhytquan) as vaobhytquan, sum(vaoquancothe) as vaoquancothe") \
				_T(" FROM   (	SELECT      ") \
				_T("              distinct hd_docno,         ") \
				_T("              max(he_examdate) as he_examedate, ") \
				_T("              hd_icd as icd, ") \
				_T("             case when hd_object in ('1', '2', '11') then 1 else 0 end as tongdoituong, ") \
				_T("             case when hd_object = '1' then 1 else 0 end as quan,") \
				_T("             case when hd_object = '2' then 1 else 0 end as huu,") \
				_T("             case when hd_object = '11' then 1 else 0 end as bhytquan,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('1', '2', '11') then 1 else 0 end as tongvao,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('1') then 1 else 0 end as vaoquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('2')  then 1 else 0 end as vaobhytquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('11')then 1 else 0 end as vaoquancothe") \
				_T("          FROM hms_doc  ") \
				_T("          LEFT JOIN hms_exam ON (he_docno = hd_docno)  ") \
				_T("          WHERE hd_status = 'T' AND length(hd_icd) > 0     ") \
				_T("                AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T("          AND hd_admitdept = 'C1.2'  ") \
				_T("          AND hd_icd in ('I.21', 'I25', 'I25.2', 'I25.7', 'U14', 'U82.8', 'I49.7', 'U64.5', 'U64.9', 'I50', 'I50.0', 'I50.2', 'I50.3', 'I50.4', 'I50.9')") \
				_T("           GROUP BY hd_docno, hd_icd, hd_object, hd_suggestion, hd_indept") \
				_T("    ) ") \
				_T(" LEFT JOIN hms_icd ON(hi_icd = icd) ") \
				_T(" LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T(" LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") \
				_T(" GROUP BY icd, hi_name") \
				_T(" union all ") \
				_T(" SELECT  icd, 'nhom4' as phanloainhom,        ") \
				_T("        hi_name as ICDName,    ") \
				_T("        count(hd_docno) as total, sum(tongdoituong) as tongdoituong, sum(quan) as quan, sum(huu) as huu, sum(bhytquan) as bhytquan, ") \
				_T("        sum(tongvao) as tongvao, sum(vaoquan) as vaoquan, sum(vaobhytquan) as vaobhytquan, sum(vaoquancothe) as vaoquancothe") \
				_T(" FROM   (	SELECT      ") \
				_T("              distinct hd_docno,         ") \
				_T("              max(he_examdate) as he_examedate, ") \
				_T("              hd_icd as icd, ") \
				_T("             case when hd_object in ('1', '2', '11') then 1 else 0 end as tongdoituong, ") \
				_T("             case when hd_object = '1' then 1 else 0 end as quan,") \
				_T("             case when hd_object = '2' then 1 else 0 end as huu,") \
				_T("             case when hd_object = '11' then 1 else 0 end as bhytquan,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('1', '2', '11') then 1 else 0 end as tongvao,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('1') then 1 else 0 end as vaoquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('2')  then 1 else 0 end as vaobhytquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('11')then 1 else 0 end as vaoquancothe") \
				_T("          FROM hms_doc  ") \
				_T("          LEFT JOIN hms_exam ON (he_docno = hd_docno)  ") \
				_T("          WHERE hd_status = 'T' AND length(hd_icd) > 0     ") \
				_T("                AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T("          AND hd_admitdept = 'C1.2'  ") \
				_T("          AND hd_icd in ('M10')") \
				_T("           GROUP BY hd_docno, hd_icd, hd_object, hd_suggestion, hd_indept") \
				_T("    ) ") \
				_T(" LEFT JOIN hms_icd ON(hi_icd = icd) ") \
				_T(" LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T(" LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") \
				_T(" GROUP BY icd, hi_name") \
				_T(" union all") \
				_T(" SELECT  icd, 'nhom5' as phanloainhom,        ") \
				_T("        hi_name as ICDName,    ") \
				_T("        count(hd_docno) as total, sum(tongdoituong) as tongdoituong, sum(quan) as quan, sum(huu) as huu, sum(bhytquan) as bhytquan, ") \
				_T("        sum(tongvao) as tongvao, sum(vaoquan) as vaoquan, sum(vaobhytquan) as vaobhytquan, sum(vaoquancothe) as vaoquancothe") \
				_T(" FROM   (	SELECT      ") \
				_T("              distinct hd_docno,         ") \
				_T("              max(he_examdate) as he_examedate, ") \
				_T("              hd_icd as icd, ") \
				_T("             case when hd_object in ('1', '2', '11') then 1 else 0 end as tongdoituong, ") \
				_T("             case when hd_object = '1' then 1 else 0 end as quan,") \
				_T("             case when hd_object = '2' then 1 else 0 end as huu,") \
				_T("             case when hd_object = '11' then 1 else 0 end as bhytquan,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('1', '2', '11') then 1 else 0 end as tongvao,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('1') then 1 else 0 end as vaoquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('2')  then 1 else 0 end as vaobhytquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('11')then 1 else 0 end as vaoquancothe") \
				_T("          FROM hms_doc  ") \
				_T("          LEFT JOIN hms_exam ON (he_docno = hd_docno)  ") \
				_T("          WHERE hd_status = 'T' AND length(hd_icd) > 0     ") \
				_T("                AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T("          AND hd_admitdept = 'C1.2'  ") \
				_T("          AND hd_icd in ('R30.2', 'R30.3')") \
				_T("           GROUP BY hd_docno, hd_icd, hd_object, hd_suggestion, hd_indept") \
				_T("    ) ") \
				_T(" LEFT JOIN hms_icd ON(hi_icd = icd) ") \
				_T(" LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T(" LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") \
				_T(" GROUP BY icd, hi_name") \
				_T(" union all") \
				_T(" SELECT  icd, 'nhom6' as phanloainhom,        ") \
				_T("        hi_name as ICDName,    ") \
				_T("        count(hd_docno) as total, sum(tongdoituong) as tongdoituong, sum(quan) as quan, sum(huu) as huu, sum(bhytquan) as bhytquan, ") \
				_T("        sum(tongvao) as tongvao, sum(vaoquan) as vaoquan, sum(vaobhytquan) as vaobhytquan, sum(vaoquancothe) as vaoquancothe") \
				_T(" FROM   (	SELECT      ") \
				_T("              distinct hd_docno,         ") \
				_T("              max(he_examdate) as he_examedate, ") \
				_T("              hd_icd as icd, ") \
				_T("             case when hd_object in ('1', '2', '11') then 1 else 0 end as tongdoituong, ") \
				_T("             case when hd_object = '1' then 1 else 0 end as quan,") \
				_T("             case when hd_object = '2' then 1 else 0 end as huu,") \
				_T("             case when hd_object = '11' then 1 else 0 end as bhytquan,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('1', '2', '11') then 1 else 0 end as tongvao,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('1') then 1 else 0 end as vaoquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('2')  then 1 else 0 end as vaobhytquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('11')then 1 else 0 end as vaoquancothe") \
				_T("          FROM hms_doc  ") \
				_T("          LEFT JOIN hms_exam ON (he_docno = hd_docno)  ") \
				_T("          WHERE hd_status = 'T' AND length(hd_icd) > 0     ") \
				_T("                AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T("          AND hd_admitdept = 'C1.2'  ") \
				_T("          AND hd_icd in ('G81', 'I61', 'I63', 'I69.3', 'U02.8', 'U02.11')") \
				_T("         GROUP BY hd_docno, hd_icd, hd_object, hd_suggestion, hd_indept") \
				_T("    ) ") \
				_T(" LEFT JOIN hms_icd ON(hi_icd = icd) ") \
				_T(" LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T(" LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") \
				_T(" GROUP BY icd, hi_name") \
				_T(" union all") \
				_T(" SELECT  icd, 'nhom7' as phanloainhom,        ") \
				_T("        hi_name as ICDName,    ") \
				_T("        count(hd_docno) as total, sum(tongdoituong) as tongdoituong, sum(quan) as quan, sum(huu) as huu, sum(bhytquan) as bhytquan, ") \
				_T("        sum(tongvao) as tongvao, sum(vaoquan) as vaoquan, sum(vaobhytquan) as vaobhytquan, sum(vaoquancothe) as vaoquancothe") \
				_T(" FROM   (	SELECT      ") \
				_T("              distinct hd_docno,         ") \
				_T("              max(he_examdate) as he_examedate, ") \
				_T("              hd_icd as icd, ") \
				_T("             case when hd_object in ('1', '2', '11') then 1 else 0 end as tongdoituong, ") \
				_T("             case when hd_object = '1' then 1 else 0 end as quan,") \
				_T("             case when hd_object = '2' then 1 else 0 end as huu,") \
				_T("             case when hd_object = '11' then 1 else 0 end as bhytquan,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('1', '2', '11') then 1 else 0 end as tongvao,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('1') then 1 else 0 end as vaoquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('2')  then 1 else 0 end as vaobhytquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('11')then 1 else 0 end as vaoquancothe") \
				_T("          FROM hms_doc  ") \
				_T("          LEFT JOIN hms_exam ON (he_docno = hd_docno)  ") \
				_T("          WHERE hd_status = 'T' AND length(hd_icd) > 0     ") \
				_T("                AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T("          AND hd_admitdept = 'C1.2'  ") \
				_T("          AND(hd_icd in ('B15', 'B15.0', 'B15.9', 'B16', 'B16.0', 'B16.1', 'B16.2', 'B16.9', 'B17', 'B17.0', 'B17.1', 'B17.2', 'B17.8', 'B18', 'B18.0')") \
				_T("                 and hd_icd in ('B18.1', 'B18.2', 'B18.8', 'B18.9', 'B19', 'B19.0', 'B19.9', 'B25.1','K70.1', 'K71.2', 'K71.3', 'K71.5', 'K71.6', 'K73','K73.0', 'K73.2', 'K73.8', 'K73.9', 'K75')") \
				_T("                 and hd_icd in ('K75.2', 'K75.3', 'K75.8', 'K75.9', 'K76.10')") \
				_T("             )") \
				_T("         GROUP BY hd_docno, hd_icd, hd_object, hd_suggestion, hd_indept") \
				_T("    ) ") \
				_T(" LEFT JOIN hms_icd ON(hi_icd = icd) ") \
				_T(" LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T(" LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") \
				_T(" GROUP BY icd, hi_name") \
				_T(" union all") \
				_T(" SELECT  icd, 'nhom8' as phanloainhom,        ") \
				_T("        hi_name as ICDName,    ") \
				_T("        count(hd_docno) as total, sum(tongdoituong) as tongdoituong, sum(quan) as quan, sum(huu) as huu, sum(bhytquan) as bhytquan, ") \
				_T("        sum(tongvao) as tongvao, sum(vaoquan) as vaoquan, sum(vaobhytquan) as vaobhytquan, sum(vaoquancothe) as vaoquancothe") \
				_T(" FROM   (	SELECT      ") \
				_T("              distinct hd_docno,         ") \
				_T("              max(he_examdate) as he_examedate, ") \
				_T("              hd_icd as icd, ") \
				_T("             case when hd_object in ('1', '2', '11') then 1 else 0 end as tongdoituong, ") \
				_T("             case when hd_object = '1' then 1 else 0 end as quan,") \
				_T("             case when hd_object = '2' then 1 else 0 end as huu,") \
				_T("             case when hd_object = '11' then 1 else 0 end as bhytquan,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('1', '2', '11') then 1 else 0 end as tongvao,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('1') then 1 else 0 end as vaoquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('2')  then 1 else 0 end as vaobhytquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('11')then 1 else 0 end as vaoquancothe") \
				_T("          FROM hms_doc  ") \
				_T("          LEFT JOIN hms_exam ON (he_docno = hd_docno)  ") \
				_T("          WHERE hd_status = 'T' AND length(hd_icd) > 0     ") \
				_T("                AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T("          AND hd_admitdept = 'C1.2'  ") \
				_T("          AND hd_icd in ('U78', 'Z22.5')") \
				_T("         GROUP BY hd_docno, hd_icd, hd_object, hd_suggestion, hd_indept") \
				_T("    ) ") \
				_T(" LEFT JOIN hms_icd ON(hi_icd = icd) ") \
				_T(" LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T(" LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") \
				_T(" GROUP BY icd, hi_name") \
				_T(" union all") \
				_T(" SELECT  icd, 'nhom9' as phanloainhom,        ") \
				_T("        hi_name as ICDName,    ") \
				_T("        count(hd_docno) as total, sum(tongdoituong) as tongdoituong, sum(quan) as quan, sum(huu) as huu, sum(bhytquan) as bhytquan, ") \
				_T("        sum(tongvao) as tongvao, sum(vaoquan) as vaoquan, sum(vaobhytquan) as vaobhytquan, sum(vaoquancothe) as vaoquancothe") \
				_T(" FROM   (	SELECT      ") \
				_T("              distinct hd_docno,         ") \
				_T("              max(he_examdate) as he_examedate, ") \
				_T("              hd_icd as icd, ") \
				_T("             case when hd_object in ('1', '2', '11') then 1 else 0 end as tongdoituong, ") \
				_T("             case when hd_object = '1' then 1 else 0 end as quan,") \
				_T("             case when hd_object = '2' then 1 else 0 end as huu,") \
				_T("             case when hd_object = '11' then 1 else 0 end as bhytquan,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('1', '2', '11') then 1 else 0 end as tongvao,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('1') then 1 else 0 end as vaoquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('2')  then 1 else 0 end as vaobhytquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('11')then 1 else 0 end as vaoquancothe") \
				_T("          FROM hms_doc  ") \
				_T("          LEFT JOIN hms_exam ON (he_docno = hd_docno)  ") \
				_T("          WHERE hd_status = 'T' AND length(hd_icd) > 0     ") \
				_T("                AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T("          AND hd_admitdept = 'C1.2'  ") \
				_T("          AND hd_icd in ('K21', 'K21.0', 'K21.9', 'K22.6', 'K25', 'K25.1', 'K28', 'K29', 'K29.0', 'K29.1', 'K29.2', 'K29.3', 'K29.4', 'K29.5', 'K29.6', 'K29.7', 'K29.9', 'K26.1')") \
				_T("         GROUP BY hd_docno, hd_icd, hd_object, hd_suggestion, hd_indept") \
				_T("    ) ") \
				_T(" LEFT JOIN hms_icd ON(hi_icd = icd) ") \
				_T(" LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T(" LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") \
				_T(" GROUP BY icd, hi_name") \
				_T(" union all") \
				_T(" SELECT  icd, 'nhom99' as phanloainhom,        ") \
				_T("        hi_name as ICDName,    ") \
				_T("        count(hd_docno) as total, sum(tongdoituong) as tongdoituong, sum(quan) as quan, sum(huu) as huu, sum(bhytquan) as bhytquan, ") \
				_T("        sum(tongvao) as tongvao, sum(vaoquan) as vaoquan, sum(vaobhytquan) as vaobhytquan, sum(vaoquancothe) as vaoquancothe") \
				_T(" FROM   (	SELECT      ") \
				_T("              distinct hd_docno,         ") \
				_T("              max(he_examdate) as he_examedate, ") \
				_T("              hd_icd as icd, ") \
				_T("             case when hd_object in ('1', '2', '11') then 1 else 0 end as tongdoituong, ") \
				_T("             case when hd_object = '1' then 1 else 0 end as quan,") \
				_T("             case when hd_object = '2' then 1 else 0 end as huu,") \
				_T("             case when hd_object = '11' then 1 else 0 end as bhytquan,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('1', '2', '11') then 1 else 0 end as tongvao,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('1') then 1 else 0 end as vaoquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('2')  then 1 else 0 end as vaobhytquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('11')then 1 else 0 end as vaoquancothe") \
				_T("          FROM hms_doc  ") \
				_T("          LEFT JOIN hms_exam ON (he_docno = hd_docno)  ") \
				_T("          WHERE hd_status = 'T' AND length(hd_icd) > 0     ") \
				_T("                AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T("          AND hd_admitdept = 'C1.2'  ") \
				_T("          AND hd_icd in ('M47', 'M47.3', 'M47.4', 'M47.8', 'M50.3', 'M51.0', 'M51.2', 'M53.4', 'M54.6','U11.9')") \
				_T("         GROUP BY hd_docno, hd_icd, hd_object, hd_suggestion, hd_indept") \
				_T("    ) ") \
				_T(" LEFT JOIN hms_icd ON(hi_icd = icd) ") \
				_T(" LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T(" LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") \
				_T(" GROUP BY icd, hi_name") \
				_T(" union all") \
				_T(" SELECT  icd, 'nhom999' as phanloainhom,        ") \
				_T("        hi_name as ICDName,    ") \
				_T("        count(hd_docno) as total, sum(tongdoituong) as tongdoituong, sum(quan) as quan, sum(huu) as huu, sum(bhytquan) as bhytquan, ") \
				_T("        sum(tongvao) as tongvao, sum(vaoquan) as vaoquan, sum(vaobhytquan) as vaobhytquan, sum(vaoquancothe) as vaoquancothe") \
				_T(" FROM   (	SELECT      ") \
				_T("              distinct hd_docno,         ") \
				_T("              max(he_examdate) as he_examedate, ") \
				_T("              hd_icd as icd, ") \
				_T("             case when hd_object in ('1', '2', '11') then 1 else 0 end as tongdoituong, ") \
				_T("             case when hd_object = '1' then 1 else 0 end as quan,") \
				_T("             case when hd_object = '2' then 1 else 0 end as huu,") \
				_T("             case when hd_object = '11' then 1 else 0 end as bhytquan,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('1', '2', '11') then 1 else 0 end as tongvao,") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('1') then 1 else 0 end as vaoquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in('2')  then 1 else 0 end as vaobhytquan,   ") \
				_T("             case when hd_suggestion = 'D' and length(hd_indept) > 0 and hd_object in ('11')then 1 else 0 end as vaoquancothe") \
				_T("          FROM hms_doc  ") \
				_T("          LEFT JOIN hms_exam ON (he_docno = hd_docno)  ") \
				_T("          WHERE hd_status = 'T' AND length(hd_icd) > 0     ") \
				_T("                AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T("          AND hd_admitdept = 'C1.2'  ") \
				_T("          AND (hd_icd in ('C03', 'C09.2', 'C09.3', 'C18', 'C19.5', 'C49', 'C67',  'C73', 'C74', 'C75', 'C76', 'C77', 'C78', 'C78.0', 'C78.1', 'C78.2', 'C78.3', 'C78.4', 'C78.5')") \
				_T("                 and hd_icd in   ('C78.6', 'C78.7', 'C78.8', 'C79', 'C79.0', 'C79.1', 'C79.2','C79.3','C79.4', 'C79.5', 'C79.6', 'C79.7', 'C79.7', 'C80', 'C96', 'C96.7', 'C96.9', 'C97'))") \
				_T("         GROUP BY hd_docno, hd_icd, hd_object, hd_suggestion, hd_indept") \
				_T("    ) ") \
				_T(" LEFT JOIN hms_icd ON(hi_icd = icd) ") \
				_T(" LEFT JOIN hms_icdgroup ON (hi_groupid=hig_id)") \
				_T(" LEFT JOIN hms_icdchapter on(hic_id=hig_chapterid) 	") \
				_T(" GROUP BY icd, hi_name") \
				_T(" ORDER BY phanloainhom"), m_szFromDate, m_szToDate, szWhere,
											  m_szFromDate, m_szToDate, szWhere,
											  m_szFromDate, m_szToDate, szWhere,
											  m_szFromDate, m_szToDate, szWhere,
											  m_szFromDate, m_szToDate, szWhere,
											  m_szFromDate, m_szToDate, szWhere,
											  m_szFromDate, m_szToDate, szWhere,
											  m_szFromDate, m_szToDate, szWhere,
											  m_szFromDate, m_szToDate, szWhere,
											  m_szFromDate, m_szToDate, szWhere,
											  m_szFromDate, m_szToDate, szWhere);
	
     rs.ExecSQL(szSQL);
	xls.SetCellMergedColumns(nCol, nRow , 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(nCol, nRow + 3, 13);
	xls.SetCellText(nCol, nRow + 3, _T("TH\x1ED0NG K\xCA \x42\x1EC6NH T\x1EACT TH\x45O \x31\x31 M\x1EB6T \x42\x1EC6NH \x43H\xCDNH"), FMT_TEXT | FMT_CENTER, true, 14);
	xls.SetCellMergedColumns(nCol, nRow + 4, 13);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellMergedRows(nCol, nRow + 5, 2);
	xls.SetCellMergedRows(nCol + 1, nRow + 5, 2);
	xls.SetCellMergedRows(nCol + 2, nRow + 5, 2);
	xls.SetCellText(nCol, nRow + 5, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 5, _T("ICD10"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 5, _T("M\x1EB7t \x62\x1EC7nh"), FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellMergedColumns(nCol + 3, nRow + 5, 4);
	xls.SetCellText(nCol + 3, nRow + 5, _T("Kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 3, nRow + 6, _T("Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 6, _T("\x42HYTQ"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 5, nRow + 6, _T("\x42HYT qu\xE2n nh\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 6, nRow + 6, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedColumns(nCol + 7, nRow + 5, 4);
	xls.SetCellText(nCol + 7, nRow + 5, _T("V\xE0o vi\x1EC7n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 7, nRow + 6, _T("Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 8, nRow + 6, _T("\x42HYTQ"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 9, nRow + 6, _T("\x42HYT qu\xE2n nh\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 10, nRow + 6, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 11);

	int nIndex = 1, nItem = 0;
	long nTotal[10],ttlCost[10];

	CString c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11;
	CString szNewLine, szOldLine;
	c1 = _T("RL\x43H Lipit");
	c2 = _T("T\x103ng huy\x1EBFt \xE1p");
	c3 = _T("\x42\x1EC7nh l\xFD tim m\x1EA1\x63h");
	c4 = _T("G\xFAt");
	c5 = _T("\x110\xE1i th\xE1o \x111\x1B0\x1EDDng");
	c6 = _T("T\x61i \x62i\x1EBFn m\x1EA1\x63h m\xE1u n\xE3o");
	c7 = _T("Vi\xEAm g\x61n \x63\x1EA5p, m\x1EA1n, \x78\x1A1 g\x61n");
	c8 = _T("Nhi\x1EC5m vi r\xFAt");
	c9 = _T("Vi\xEAm lo\xE1t \x64\x1EA1 \x64\xE0y, h\xE0nh t\xE1 tr\xE0ng");
	c10 = _T("Tho\xE1i h\xF3\x61 \x63\x1ED9t s\x1ED1ng \x63\x1ED5, th\x1EAFt l\x1B0ng");
	c11 = _T("K \x63\xE1\x63 lo\x1EA1i");

	for (int i = 1; i<=10; i++)
	{
		nTotal[i] = 0;
		ttlCost[i] = 0;
	}
	nRow = 6;
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("phanloainhom"), szNewLine);

		if(szNewLine != szOldLine && !szNewLine.IsEmpty())
		{	
			if(nTotal[4] > 0)
			{			
				CString szField, szAmount;
				nRow++;
				TranslateString(_T("T\x1ED5ng nh\xF3m:"), szAmount);												
				xls.SetCellMergedColumns(0,nRow,3);
				xls.SetCellText(0, nRow, szAmount, FMT_TEXT,true,11,0);
				for(int i = 1; i<=10; i++)
				{
					tmpStr.Format(_T("%d"), nTotal[i]);
					xls.SetCellText(i+2, nRow, tmpStr, FMT_NUMBER1,true,11,0);
					ttlCost[i] += nTotal[i];
					nTotal[i] = 0;								
					//nIn=1;
				}
			}
			nRow++;
			//CString szField, szAmount;
			xls.SetCellMergedColumns(0,nRow,11);	
			//rs.GetValue(_T("phanloainhom"), szNewLine);
			
			if(szNewLine == _T("nhom1")) 
			{
				xls.SetCellText(0, nRow, _T("RL\x43H Lipit"), FMT_TEXT,true,11,0);
				
			}
			else if(szNewLine == _T("nhom2"))  
			{
				xls.SetCellText(0, nRow, _T("T\x103ng huy\x1EBFt \xE1p"), FMT_TEXT,true,11,0);
					
			}
			
			else if(szNewLine == _T("nhom3"))  xls.SetCellText(0, nRow, c3, FMT_TEXT,true,11,0);
			else if(szNewLine == _T("nhom4"))  xls.SetCellText(0, nRow, c4, FMT_TEXT,true,11,0);
			else if(szNewLine == _T("nhom5"))  xls.SetCellText(0, nRow, c5, FMT_TEXT,true,11,0);
			else if(szNewLine == _T("nhom6"))  xls.SetCellText(0, nRow, c6, FMT_TEXT,true,11,0);
			else if(szNewLine == _T("nhom7"))  xls.SetCellText(0, nRow, c7, FMT_TEXT,true,11,0);
			else if(szNewLine == _T("nhom8"))  xls.SetCellText(0, nRow, c8, FMT_TEXT,true,11,0);
			else if(szNewLine == _T("nhom9"))  xls.SetCellText(0, nRow, c9, FMT_TEXT,true,11,0);
			else if(szNewLine == _T("nhom99"))  xls.SetCellText(0, nRow, c10, FMT_TEXT,true,11,0);
			else if(szNewLine == _T("nhom999"))  xls.SetCellText(0, nRow, c11, FMT_TEXT,true,11,0);
			szOldLine = szNewLine;
			nIndex=1;
			
		}
		nRow++;
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("icd"),tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("ICDName"),tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("quan"), tmpStr);
		nTotal[1] += ToLong(tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("huu"),tmpStr);
		nTotal[2] += ToLong(tmpStr);
		xls.SetCellText(nCol + 4, nRow , tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("bhytquan"),tmpStr);
		nTotal[3] += ToLong(tmpStr);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("tongdoituong"),tmpStr);
		nTotal[4] += ToLong(tmpStr);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("vaoquan"), tmpStr);
		nTotal[5] += ToLong(tmpStr);
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("vaobhytquan"),tmpStr);
		nTotal[6] += ToLong(tmpStr);
		xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("vaoquancothe"),tmpStr);
		nTotal[7] += ToLong(tmpStr);
		xls.SetCellText(nCol + 9, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("tongvao"),tmpStr);
		nTotal[8] += ToLong(tmpStr);
		xls.SetCellText(nCol + 10, nRow, tmpStr, FMT_NUMBER1);

		
		rs.MoveNext();		
	}
	for ( int i = 1; i<=10; i++)
	{
		ttlCost[i] += nTotal[i];
	}
	if(nTotal[4] > 0)
			{			
				CString szField, szAmount;
				nRow++;
				TranslateString(_T("T\x1ED5ng nh\xF3m:"), szAmount);												
				xls.SetCellMergedColumns(0,nRow,3);
				xls.SetCellText(0, nRow, szAmount, FMT_TEXT,true,11,0);
				for(int i = 1; i<=10; i++)
				{
					tmpStr.Format(_T("%d"), nTotal[i]);
					xls.SetCellText(i+2, nRow, tmpStr, FMT_NUMBER1,true,11,0);
					
				}
			}
	/*if(ttlCost[4] > 0)
	{
		nRow++;
		TranslateString(_T("Total"), tmpStr);		
		xls.SetCellMergedColumns(0,nRow,3);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true,11,0);

		for(int i = 1; i<=10; i++)
		{
			tmpStr.Format(_T("%d"), ttlCost[i]);
			xls.SetCellText(i+2, nRow, tmpStr, FMT_NUMBER1,true,11,0);		
		}
	}*/

	xls.Save(_T("Exports\\THONG KE MAT BENH - DOI TUONG THEO 11 MAT BENH.xls"));
	EndWaitCursor();
	
} 

int CEMObject11DiseaseStatisticsReport::OnAddEMObject11DiseaseStatisticsReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMObject11DiseaseStatisticsReport::OnEditEMObject11DiseaseStatisticsReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMObject11DiseaseStatisticsReport::OnDeleteEMObject11DiseaseStatisticsReport(){
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
 		OnCancelEMObject11DiseaseStatisticsReport(); 
 	}
	return 0;
}
int CEMObject11DiseaseStatisticsReport::OnSaveEMObject11DiseaseStatisticsReport(){
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
 		//OnEMObject11DiseaseStatisticsReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMObject11DiseaseStatisticsReport::OnCancelEMObject11DiseaseStatisticsReport()
{
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
int CEMObject11DiseaseStatisticsReport::OnEMObject11DiseaseStatisticsReportListLoadData()
{
	return 0;
}