#include "stdafx.h"
#include "EMTestMoneybyPatient.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMTestMoneybyPatient *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMTestMoneybyPatient* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMTestMoneybyPatient *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMTestMoneybyPatient *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMTestMoneybyPatient *)pWnd)->OnToDateCheckValue();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMTestMoneybyPatient* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CEMTestMoneybyPatient *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnObjectAddNew();
}*/
static void _OnRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMTestMoneybyPatient* )pWnd)->OnRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnRoomSelendokFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnRoomSelendok();
}
/*static void _OnRoomSetfocusFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnRoomKillfocus();
}*/
/*static void _OnRoomKillfocusFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnRoomKillfocus();
}*/
static long _OnRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMTestMoneybyPatient *)pWnd)->OnRoomLoadData();
}
/*static void _OnRoomAddNewFnc(CWnd *pWnd){
	((CEMTestMoneybyPatient *)pWnd)->OnRoomAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMTestMoneybyPatient *pVw = (CEMTestMoneybyPatient *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMTestMoneybyPatient *pVw = (CEMTestMoneybyPatient *)pWnd;
	pVw->OnExportSelect();
} 
CEMTestMoneybyPatient::CEMTestMoneybyPatient(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CEMTestMoneybyPatient::~CEMTestMoneybyPatient(){
}
void CEMTestMoneybyPatient::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 245, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 250, 30, 330, 55);
	m_wndReportPeriod.Create(this,335, 30, 485, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 245, 85); 
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 490, 150);
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 60, 330, 85);
	m_wndToDate.Create(this,335, 60, 485, 85); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 90, 90, 115);
	m_wndObject.Create(this,95, 90, 485, 115); 
	m_wndRoomLabel.Create(this, _T("Room"), 10, 120, 90, 145);
	m_wndRoom.Create(this,95, 120, 485, 145); 
	//m_wndPrintPreview.Create(this, _T("Print Pre&view"), 300, 155, 410, 180);
	m_wndExport.Create(this, _T("&Export"), 415, 155, 490, 180);

}
void CEMTestMoneybyPatient::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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

	m_wndRoom.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndRoom.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);
}

void CEMTestMoneybyPatient::OnSetWindowEvents(){
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
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndRoom.SetEvent(WE_SELENDOK, _OnRoomSelendokFnc);
	//m_wndRoom.SetEvent(WE_SETFOCUS, _OnRoomSetfocusFnc);
	//m_wndRoom.SetEvent(WE_KILLFOCUS, _OnRoomKillfocusFnc);
	m_wndRoom.SetEvent(WE_SELCHANGE, _OnRoomSelectChangeFnc);
	m_wndRoom.SetEvent(WE_LOADDATA, _OnRoomLoadDataFnc);
	//m_wndRoom.SetEvent(WE_ADDNEW, _OnRoomAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = pMF->GetSysDate() + _T("00:00");
	m_szToDate = pMF->GetSysDate() + _T("23:59");
	UpdateData(false);

}
void CEMTestMoneybyPatient::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_TextEx(pDX, m_wndRoom.GetDlgCtrlID(), m_szRoomKey);

}
void CEMTestMoneybyPatient::SetDefaultValues(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szObjectKey.Empty();
}

int CEMTestMoneybyPatient::SetMode(int nMode){
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

/*void CEMTestMoneybyPatient::OnYearChange(){
	
} */
/*void CEMTestMoneybyPatient::OnYearSetfocus(){
	
} */
/*void CEMTestMoneybyPatient::OnYearKillfocus(){
	
} */
int CEMTestMoneybyPatient::OnYearCheckValue(){
	return 0;
}
 
void CEMTestMoneybyPatient::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
}
 
void CEMTestMoneybyPatient::OnReportPeriodSelendok(){
	CString tmpStr;
	CDate dte;
	UpdateData(true);
	int nMonth = ToInt(m_szReportPeriodKey);
	if(nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/1"), m_nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), m_nYear, nMonth, dte.GetMonthLastDay());
	}
	if(nMonth == 13){
		m_szFromDate.Format(_T("%.4d/1/1"), m_nYear);
		tmpStr.Format(_T("%.4d/3/1"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/3/%.2d 23:59" ), m_nYear,  dte.GetMonthLastDay());
	}
	if(nMonth == 14){
		m_szFromDate.Format(_T("%.4d/4/1"), m_nYear);
		tmpStr.Format(_T("%.4d/6/1"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/6/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}

	if(nMonth == 15){
		m_szFromDate.Format(_T("%.4d/7/1"), m_nYear);
		tmpStr.Format(_T("%.4d/9/1"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/9/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 16){
		m_szFromDate.Format(_T("%.4d/10/1"), m_nYear);
		tmpStr.Format(_T("%.4d/10/1"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 17){
		m_szFromDate.Format(_T("%.4d/1/1"), m_nYear);
		tmpStr.Format(_T("%.4d/12/1"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());		
	}	
	UpdateData(false);
}

/*void CEMTestMoneybyPatient::OnReportPeriodSetfocus(){
	
}*/
/*void CEMTestMoneybyPatient::OnReportPeriodKillfocus(){
	
}*/
long CEMTestMoneybyPatient::OnReportPeriodLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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

/*void CEMTestMoneybyPatient::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CEMTestMoneybyPatient::OnFromDateChange(){
	
} */
/*void CEMTestMoneybyPatient::OnFromDateSetfocus(){
	
} */
/*void CEMTestMoneybyPatient::OnFromDateKillfocus(){
	
} */
int CEMTestMoneybyPatient::OnFromDateCheckValue(){
	return 0;
}
 
/*void CEMTestMoneybyPatient::OnToDateChange(){
	
} */
/*void CEMTestMoneybyPatient::OnToDateSetfocus(){
	
} */
/*void CEMTestMoneybyPatient::OnToDateKillfocus(){
	
} */
int CEMTestMoneybyPatient::OnToDateCheckValue(){
	return 0;
}
 
void CEMTestMoneybyPatient::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
void CEMTestMoneybyPatient::OnObjectSelendok(){
	 
}

/*void CEMTestMoneybyPatient::OnObjectSetfocus(){
	
}*/
/*void CEMTestMoneybyPatient::OnObjectKillfocus(){
	
}*/
long CEMTestMoneybyPatient::OnObjectLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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

/*void CEMTestMoneybyPatient::OnObjectAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CEMTestMoneybyPatient::OnRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CEMTestMoneybyPatient::OnRoomSelendok(){
	 
}
/*void CEMTestMoneybyPatient::OnRoomSetfocus(){
	
}*/
/*void CEMTestMoneybyPatient::OnRoomKillfocus(){
	
}*/
long CEMTestMoneybyPatient::OnRoomLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();
	if(m_wndRoom.IsSearchKey() && !m_szRoomKey.IsEmpty())
	{
		szWhere.Format(_T("and hrl_id = %d"), ToInt(m_szRoomKey));
	}
	m_wndRoom.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("select hrl_name as name, hrl_id as id from hms_roomlist where hrl_deptid = 'C1.1' %s Order by id"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndRoom.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMTestMoneybyPatient::OnRoomAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CEMTestMoneybyPatient::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString tmpStr;
	CString szSQL;
	CString szWhere;
	UpdateData(true);
	if(!rpt.Init(_T("Reports/MS/E_BAOCAOTHONGKETINHHINHKHAMVANHANDIEUTRINOITRUCBCC.RPT")) )
		return ;
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("Object")), m_wndObject.GetCurrent(1));
	rpt.GetReportHeader()->SetValue(_T("Object"), tmpStr);
	StringUpper(pMF->m_szHealthService, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HealthService"), tmpStr);
	StringUpper(pMF->m_szHospitalName, tmpStr);
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
 
void CEMTestMoneybyPatient::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CString tmpStr, szSQL;
	CRecord rs(&pMF->m_db);
	CExcel xls;	
	xls.CreateSheet(1);
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
	int nIndex = 0;
	xls.SetColumnWidth(0, 30);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 15);
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 5);
	xls.SetCellMergedColumns(nCol, nRow + 3, 5);
	xls.SetCellText(nCol, nRow, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol, nRow + 2, _T("TH\x1ED0NG K\xCA TI\x1EC0N \x58\xC9T NGHI\x1EC6M TH\x45O \x42\x1EC6NH NH\xC2N"), FMT_TEXT | FMT_CENTER, true, 12);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), m_szFromDate, m_szToDate);
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 10);
	CStringArray arrCol;
	arrCol.Add(_T("Lo\x1EA1i \x78\xE9t nghi\x1EC7m"));
	arrCol.Add(_T("S\x1ED1 \x78\xE9t nghi\x1EC7m"));
	arrCol.Add(_T("T\x1ED5ng s\x1ED1 BN"));
	arrCol.Add(_T("S\x1ED1 ti\x1EC1n"));
	arrCol.Add(_T("TB 1 BN"));
	for (int i = 0; i < arrCol.GetCount(); i++)
		xls.SetCellText(nCol + i, nRow + 4, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	nRow = 5;
	while(!rs.IsEOF())
	{
		xls.SetCellText(nCol, nRow, rs.GetValue(_T("groupname")), FMT_TEXT);
		xls.SetCellText(nCol + 1, nRow, rs.GetValue(_T("testnum")), FMT_NUMBER1);
		xls.SetCellText(nCol + 2, nRow, rs.GetValue(_T("patientnum")), FMT_NUMBER1);
		xls.SetCellText(nCol + 3, nRow, rs.GetValue(_T("amount")), FMT_NUMBER1);
		xls.SetCellText(nCol + 4, nRow, rs.GetValue(_T("avg")), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();		
	}
	//Total
	//xls.SetCellText(nCol, nRow, _T("T\x1ED5ng \x63\x1ED9ng:"), FMT_TEXT | FMT_CENTER, true, 12);
	EndWaitCursor();
	xls.Save(_T("Exports\\TIEN KHAM THEO BN.xls"));
}
 
CString CEMTestMoneybyPatient::GetQueryString(){
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND hpc_deptid = 'C1.1' AND hpc_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				   _T(" AND hpc_status IN ('P', 'S', 'T')"), m_szFromDate, m_szToDate);
	if (!m_szRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpc_roomid = %d"), ToInt(m_szRoomKey));
	if (!m_szObjectKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hd_object = %d"), ToInt(m_szObjectKey));
	szSQL.Format(_T(" SELECT    hfg_name groupname, ") \
				_T("           Count(DISTINCT hpc_orderid) testnum, ") \
				_T("           Count(DISTINCT hpc_docno) patientnum, ") \
				_T("           SUM(hfe_cost) amount, ") \
				_T("           SUM(hfe_cost) / Count(DISTINCT hpc_docno) avg") \
				_T(" FROM      (SELECT    hpc_groupid, ") \
				_T("                      hpc_docno, ") \
				_T("                      hpc_orderid, ") \
				_T("                      CASE WHEN ho_type IN ( 'I', 'C' ) THEN f.hfe_insprice * hfe_quantity ") \
				_T("                        WHEN ho_type IN ( 'D', 'P' ) THEN f.hfe_polprice * hfe_quantity ") \
				_T("                      ELSE f.hfe_cost ") \
				_T("                      END hfe_cost ") \
				_T("            FROM      hms_testorder ") \
				_T("            LEFT JOIN hms_fee f ON ( hpc_docno = hfe_docno ") \
				_T("                                     AND hpc_orderid = hfe_orderid ) ") \
				_T("            LEFT JOIN hms_doc ON ( hd_docno = hfe_docno ) ") \
				_T("            LEFT JOIN hms_object ON ( ho_id = hd_object ) ") \
				_T("            WHERE     hfe_type = 'T' %s") \
				_T("            UNION ALL ") \
				_T("            SELECT    hpc_groupid, ") \
				_T("                      hpc_docno, ") \
				_T("                      hpc_orderid, ") \
				_T("                      CASE WHEN ho_type IN ( 'I', 'C' ) THEN f.hfe_insprice * hfe_quantity ") \
				_T("                        WHEN ho_type IN ( 'D', 'P' ) THEN f.hfe_polprice * hfe_quantity ") \
				_T("                      ELSE f.hfe_cost ") \
				_T("                      END hfe_cost ") \
				_T("            FROM      hms_pacsorder ") \
				_T("            LEFT JOIN hms_fee f ON ( hpc_docno = hfe_docno ") \
				_T("                                     AND hpc_orderid = hfe_orderid ) ") \
				_T("            LEFT JOIN hms_doc ON ( hd_docno = hfe_docno ) ") \
				_T("            LEFT JOIN hms_object ON ( ho_id = hd_object ) ") \
				_T("            WHERE     hfe_type = 'P' %s") \
				_T("                  AND hfe_group NOT IN ( 'B2200', 'B2300' )) ") \
				_T(" LEFT JOIN hms_fee_group ON ( hfg_id = hpc_groupid ) ") \
				_T(" GROUP     BY hfg_name ") \
				_T(" ORDER     BY hfg_name "), szWhere, szWhere);

	_fmsg(_T("%s"), szSQL);
	return szSQL;
}