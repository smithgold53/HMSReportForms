#include "stdafx.h"
#include "EMExaminationTreatmentActReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMExaminationTreatmentActReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMExaminationTreatmentActReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMExaminationTreatmentActReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExaminationTreatmentActReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExaminationTreatmentActReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMExaminationTreatmentActReport* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CEMExaminationTreatmentActReport *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CEMExaminationTreatmentActReport *)pWnd)->OnObjectAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMExaminationTreatmentActReport *pVw = (CEMExaminationTreatmentActReport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMExaminationTreatmentActReport *pVw = (CEMExaminationTreatmentActReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMExaminationTreatmentActReport *pVw = (CEMExaminationTreatmentActReport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddEMExaminationTreatmentActReportFnc(CWnd *pWnd){
	 return ((CEMExaminationTreatmentActReport*)pWnd)->OnAddEMExaminationTreatmentActReport();
} 
static int _OnEditEMExaminationTreatmentActReportFnc(CWnd *pWnd){
	 return ((CEMExaminationTreatmentActReport*)pWnd)->OnEditEMExaminationTreatmentActReport();
} 
static int _OnDeleteEMExaminationTreatmentActReportFnc(CWnd *pWnd){
	 return ((CEMExaminationTreatmentActReport*)pWnd)->OnDeleteEMExaminationTreatmentActReport();
} 
static int _OnSaveEMExaminationTreatmentActReportFnc(CWnd *pWnd){
	 return ((CEMExaminationTreatmentActReport*)pWnd)->OnSaveEMExaminationTreatmentActReport();
} 
static int _OnCancelEMExaminationTreatmentActReportFnc(CWnd *pWnd){
	 return ((CEMExaminationTreatmentActReport*)pWnd)->OnCancelEMExaminationTreatmentActReport();
} 
CEMExaminationTreatmentActReport::CEMExaminationTreatmentActReport(CWnd *pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 155;
	m_szTitle = _T("Kh\xE1m v\xE0 nh\x1EADn \x111i\x1EC1u tr\x1ECB n\x1ED9i tr\xFA \x43\x42\x43\x43");
	SetDefaultValues();
}
CEMExaminationTreatmentActReport::~CEMExaminationTreatmentActReport(){
}
void CEMExaminationTreatmentActReport::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 490, 120);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 245, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 250, 30, 330, 55);
	m_wndReportPeriod.Create(this,335, 30, 485, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 245, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 60, 330, 85);
	m_wndToDate.Create(this,335, 60, 485, 85); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 90, 90, 115);
	m_wndObject.Create(this,95, 90, 485, 115); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 300, 125, 410, 150);
	m_wndExport.Create(this, _T("&Export"), 415, 125, 490, 150);
}
void CEMExaminationTreatmentActReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(35);
//	//m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(false);
	m_wndReportPeriod.LimitText(35);
//	m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(false);
//	m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(false);
	m_wndObject.SetCheckValue(false);
	m_wndObject.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndObject.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);
}
void CEMExaminationTreatmentActReport::OnSetWindowEvents(){
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
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = pMF->GetSysDate() + _T("00:00");
	m_szToDate = pMF->GetSysDate() + _T("23:59");
	SetMode(VM_EDIT);

}
void CEMExaminationTreatmentActReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);

}
void CEMExaminationTreatmentActReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CEMExaminationTreatmentActReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CEMExaminationTreatmentActReport::SetDefaultValues(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szObjectKey.Empty();
}
int CEMExaminationTreatmentActReport::SetMode(int nMode){
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
 			EnableButtons(TRUE,0,1,2,3, 4, -1); 
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
/*void CEMExaminationTreatmentActReport::OnYearChange(){
	
} */
/*void CEMExaminationTreatmentActReport::OnYearSetfocus(){
	
} */
/*void CEMExaminationTreatmentActReport::OnYearKillfocus(){
	
} */
int CEMExaminationTreatmentActReport::OnYearCheckValue(){
	return 0;
} 
void CEMExaminationTreatmentActReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
} 
void CEMExaminationTreatmentActReport::OnReportPeriodSelendok(){
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
/*void CEMExaminationTreatmentActReport::OnReportPeriodSetfocus(){
	
}*/
/*void CEMExaminationTreatmentActReport::OnReportPeriodKillfocus(){
	
}*/
long CEMExaminationTreatmentActReport::OnReportPeriodLoadData(){
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
/*void CEMExaminationTreatmentActReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMExaminationTreatmentActReport::OnFromDateChange(){
	
} */
/*void CEMExaminationTreatmentActReport::OnFromDateSetfocus(){
	
} */
/*void CEMExaminationTreatmentActReport::OnFromDateKillfocus(){
	
} */
int CEMExaminationTreatmentActReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMExaminationTreatmentActReport::OnToDateChange(){
	
} */
/*void CEMExaminationTreatmentActReport::OnToDateSetfocus(){
	
} */
/*void CEMExaminationTreatmentActReport::OnToDateKillfocus(){
	
} */
int CEMExaminationTreatmentActReport::OnToDateCheckValue(){
	return 0;
} 
void CEMExaminationTreatmentActReport::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMExaminationTreatmentActReport::OnObjectSelendok(){
	 
}
/*void CEMExaminationTreatmentActReport::OnObjectSetfocus(){
	
}*/
/*void CEMExaminationTreatmentActReport::OnObjectKillfocus(){
	
}*/
long CEMExaminationTreatmentActReport::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szSQL.Format(_T("SELECT ho_id, ho_desc FROM hms_object ORDER BY ho_id"));
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty()){
	};
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("ho_id")), 
			rs.GetValue(_T("ho_desc")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMExaminationTreatmentActReport::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMExaminationTreatmentActReport::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString tmpStr;
	CString szSQL;
	CString szWhere;
	UpdateData(true);
	if(!rpt.Init(_T("Reports/HMS/HE_BAOCAOTHONGKETINHHINHKHAMVANHANDIEUTRINOITRUCBCC.RPT")) )
		return ;
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("Object")), m_wndObject.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Object"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HealthService"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), tmpStr);
	BeginWaitCursor();
	szSQL = GetQueryString();
	//QueryOpener(szSQL);
//QueryOpener(szSQL);
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	
	int nTotal[13];
	for(int i = 0 ; i<13; i++)
	{
		nTotal[i] = 0;
	}
	CReportSection* rptDetail;
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail(); 
		rs.GetValue(_T("rank"), tmpStr);
		rptDetail->SetValue(_T("rank"), tmpStr);
// C1 - 2: 
// Kham quan
		rs.GetValue(_T("c1"), tmpStr);
		nTotal[0] += ToInt(tmpStr);
		rptDetail->SetValue(_T("1"), tmpStr);
//kham huu
		rs.GetValue(_T("c2"), tmpStr);
		nTotal[1] += ToInt(tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
// cong
		rs.GetValue(_T("c3"), tmpStr);
		nTotal[2] += ToInt(tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
// dieu tri
		rs.GetValue(_T("c4"), tmpStr);
		nTotal[3] += ToInt(tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
// gui kham C1 - 1
// khoa noi
		rs.GetValue(_T("c5"), tmpStr);
		nTotal[4] += ToInt(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
// khoa ngoai
		rs.GetValue(_T("c6"), tmpStr);
		nTotal[5] += ToInt(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
// khoa mat
		rs.GetValue(_T("c7"), tmpStr);
		nTotal[6] += ToInt(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
// TMH
		rs.GetValue(_T("c8"), tmpStr);
		nTotal[7] += ToInt(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
//RHM
		rs.GetValue(_T("c9"), tmpStr);
		nTotal[8] += ToInt(tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);
// San
		rs.GetValue(_T("c10"), tmpStr);
		nTotal[9] += ToInt(tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);
// dong y
		rs.GetValue(_T("c11"), tmpStr);
		nTotal[10] += ToInt(tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);
// cong 
		rs.GetValue(_T("c12"), tmpStr);
		nTotal[11] += ToInt(tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);
		
//don thuoc
		rs.GetValue(_T("c13"), tmpStr);
		nTotal[12] += ToInt(tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);
		rs.MoveNext();
	}
CString  szField;
	for( int i = 1; i<=13; i++)
	{
		szField.Format(_T("s%d"), i);
		tmpStr.Format(_T("%d"), nTotal[i-1]);
		rpt.GetReportFooter()->SetValue(szField, tmpStr);
	}
	EndWaitCursor();
	tmpStr.Format(_T("%s"), pMF->GetSysDate());
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4)); 
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CEMExaminationTreatmentActReport::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CString tmpStr, szSQL;
	CRecord rs(&pMF->m_db);
	CExcel xls;	
	if (!xls.Load(_T("Exports/KhamDieuTriCBCC.xls")))
		return;
	xls.SetWorksheet(0);
	int nCol = 0, nRow = 0;
	long double nTotal[13];
	for (int i=0; i< 13; i++)
	{
		nTotal[i] = 0;
	}
	BeginWaitCursor();
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	CStringArray arrCol;
	//c1..c13
	for (int i = 0; i  <13; i++)
	{
		tmpStr.Format(_T("c%d"), i+1);
		arrCol.Add(tmpStr);
	}
	int nIndex = 0;
	nRow = 3;
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);
	nRow = 7;
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("rank"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
		for (int i = 0; i < 13; i++)
		{
			rs.GetValue(arrCol[i],tmpStr);
			nTotal[i] += ToDouble(tmpStr);
			xls.SetCellText(nCol + i + 1, nRow, tmpStr, FMT_NUMBER1);
		}
		nRow++;
		rs.MoveNext();		
	}
	//Total
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng \x63\x1ED9ng:"), FMT_TEXT | FMT_CENTER, true, 12);
	for (int i=0; i<13; i++)
	{
		tmpStr.Format(_T("%.0f"), nTotal[i]);
		xls.SetCellText(nCol + i + 1, nRow, tmpStr, FMT_NUMBER1);
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\BAO CAO KHAM VA DTNT CBCC.xls"));
} 
void CEMExaminationTreatmentActReport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CEMExaminationTreatmentActReport::OnAddEMExaminationTreatmentActReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMExaminationTreatmentActReport::OnEditEMExaminationTreatmentActReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMExaminationTreatmentActReport::OnDeleteEMExaminationTreatmentActReport(){
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
 		OnCancelEMExaminationTreatmentActReport(); 
 	}
	return 0;
}
int CEMExaminationTreatmentActReport::OnSaveEMExaminationTreatmentActReport(){
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
 		//OnEMExaminationTreatmentActReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMExaminationTreatmentActReport::OnCancelEMExaminationTreatmentActReport(){
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
int CEMExaminationTreatmentActReport::OnEMExaminationTreatmentActReportListLoadData(){
	return 0;
}
CString CEMExaminationTreatmentActReport::GetQueryString()
{
	CString szSQL, szStt;
	bool bIgnoreDocStt = false;
	if (bIgnoreDocStt)
		szStt.Empty();
	else
		szStt.Format(_T(" AND hd_status IN('P','T')"));

	szSQL.Format(_T(" SELECT (select distinct ss_desc from sys_sel where ss_id = 'hms_rank' and ss_code = hp_rank) as rank,") \
				_T("   c1,") \
				_T("   c2,") \
				_T("   c3,") \
				_T("   c4,") \
				_T("   c5,") \
				_T("   c6,") \
				_T("   c7,") \
				_T("   c8,") \
				_T("   c9,") \
				_T("   c10,") \
				_T("   c11,") \
				_T("   (c5+c6+c7+c8+c9+c10+c11) AS c12,") \
				_T("   c13") \
				_T(" FROM") \
				_T("   (SELECT hp_rank,") \
				_T("     SUM(c1)    AS c1,") \
				_T("     SUM(c2)    AS c2,") \
				_T("     SUM(c1+c2) AS c3,") \
				_T("     SUM(c5)    AS c5,") \
				_T("     SUM(c6)    AS c6,") \
				_T("     SUM(c7)    AS c7,") \
				_T("     SUM(c8)    AS c8,") \
				_T("     SUM(c9)    AS c9,") \
				_T("     SUM(c10)   AS c10,") \
				_T("     SUM(c11)   AS c11") \
				_T("   FROM") \
				_T("     (SELECT hp_rank,") \
				_T("       CASE") \
				_T("         WHEN hd_object = 1") \
				_T("         AND he_deptid ='C1.2'") \
				_T("         THEN 1") \
				_T("         ELSE 0") \
				_T("       END AS c1,") \
				_T("       CASE") \
				_T("         WHEN hd_object = 2") \
				_T("         AND he_deptid    ='C1.2'") \
				_T("         THEN 1") \
				_T("         ELSE 0") \
				_T("       END AS c2,") \
				_T("       CASE") \
				_T("         WHEN hrl_type = 2") \
				_T("         AND he_deptid       <> hd_admitdept") \
				_T("         THEN 1") \
				_T("         ELSE 0") \
				_T("       END AS c5,") \
				_T("       CASE") \
				_T("         WHEN hrl_type = 3") \
				_T("         AND he_deptid       <> hd_admitdept") \
				_T("         THEN 1") \
				_T("         ELSE 0") \
				_T("       END AS c6,") \
				_T("       CASE") \
				_T("         WHEN hrl_type = 0 and hrl_id = 12") \
				_T("         AND he_deptid       <> hd_admitdept") \
				_T("         THEN 1") \
				_T("         ELSE 0") \
				_T("       END AS c7,") \
				_T("       CASE") \
				_T("         WHEN hrl_type = 0 and hrl_id = 4") \
				_T("         AND he_deptid       <> hd_admitdept") \
				_T("         THEN 1") \
				_T("         ELSE 0") \
				_T("       END AS c8,") \
				_T("       CASE") \
				_T("         WHEN  hrl_id = 2") \
				_T("         AND he_deptid <> hd_admitdept") \
				_T("         THEN 1") \
				_T("         ELSE 0") \
				_T("       END AS c9,") \
				_T("       CASE") \
				_T("         WHEN hrl_type = 0 and hrl_id = 1") \
				_T("         AND he_deptid   <> hd_admitdept") \
				_T("         THEN 1") \
				_T("         ELSE 0") \
				_T("       END AS c10,") \
				_T("       CASE") \
				_T("         WHEN hrl_type = 8") \
				_T("         AND he_deptid       <> hd_admitdept") \
				_T("         THEN 1") \
				_T("         ELSE 0") \
				_T("       END AS c11") \
				_T("     FROM hms_doc") \
				_T("     LEFT JOIN hms_patient") \
				_T("     ON (hd_patientno = hp_patientno)") \
				_T("     LEFT JOIN hms_exam") \
				_T("     ON(he_docno=hd_docno)") \
				_T("     LEFT JOIN hms_roomlist") \
				_T("     ON(hrl_deptid=he_deptid AND hrl_id = he_roomid)") \
				_T("     WHERE (he_examdate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				_T("     AND he_status IN('T','P')  and hp_rank in ('18','17','16','15','14', '13')") \
				_T("     %s") \
				_T("     AND hp_rank        > 0") \
				_T("     AND hd_admitdept   ='C1.2'") \
				_T("     ) tbl") \
				_T("   GROUP BY hp_rank") \
				_T("   ) tbl2") \
				_T(" LEFT JOIN") \
				_T("   (SELECT hp_rank,") \
				_T("     COUNT(*) AS c4") \
				_T("   FROM hms_doc") \
				_T("   LEFT JOIN hms_patient ON (hd_patientno = hp_patientno)") \
				_T("   LEFT JOIN hms_clinical_record ON (hcr_docno = hd_docno)") \
				_T("   WHERE hd_enddept='C1.2' and hp_rank in ('18','17','16','15','14', '13')") \
				_T("   AND hcr_admitdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				_T("   AND hcr_numinward > 0") \
				_T("   GROUP BY hp_rank") \
				_T("   ) tbl3 USING(hp_rank)") \
				_T(" LEFT JOIN") \
				_T("   (SELECT hp_rank,") \
				_T("     COUNT(hpo_docno) AS c13") \
				_T("   FROM hms_doc") \
				_T("   LEFT JOIN hms_patient ON (hd_patientno = hp_patientno)") \
				_T("   LEFT JOIN hms_pharmaorder") \
				_T("   ON(hd_docno =hpo_docno)") \
				_T("   WHERE hd_admitdept='C1.2' and hp_rank in ('18','17','16','15','14', '13')") \
				_T("   AND (hpo_orderdate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				_T("   AND hpo_orderstatus IN('S', 'A')") \
				_T("   GROUP BY hp_rank") \
				_T("   ) tbl4 USING(hp_rank)") \
				_T(" ORDER BY hp_rank"), m_szFromDate, m_szToDate, szStt,m_szFromDate, m_szToDate, m_szFromDate, m_szToDate);
	return szSQL;
}