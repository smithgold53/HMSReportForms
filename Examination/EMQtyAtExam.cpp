#include "stdafx.h"
#include "EMQtyAtExam.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMQtyAtExam *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMQtyAtExam* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMQtyAtExam *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMQtyAtExam *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMQtyAtExam *)pWnd)->OnToDateCheckValue();
} 
static void _OnRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMQtyAtExam* )pWnd)->OnRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnRoomSelendokFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnRoomSelendok();
}
/*static void _OnRoomSetfocusFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnRoomKillfocus();
}*/
/*static void _OnRoomKillfocusFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnRoomKillfocus();
}*/
static long _OnRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMQtyAtExam *)pWnd)->OnRoomLoadData();
}
/*static void _OnRoomAddNewFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnRoomAddNew();
}*/
static void _OnDoctorSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMQtyAtExam* )pWnd)->OnDoctorSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDoctorSelendokFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnDoctorSelendok();
}
/*static void _OnDoctorSetfocusFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnDoctorKillfocus();
}*/
/*static void _OnDoctorKillfocusFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnDoctorKillfocus();
}*/
static long _OnDoctorLoadDataFnc(CWnd *pWnd){
	return ((CEMQtyAtExam *)pWnd)->OnDoctorLoadData();
}
/*static void _OnDoctorAddNewFnc(CWnd *pWnd){
	((CEMQtyAtExam *)pWnd)->OnDoctorAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMQtyAtExam *pVw = (CEMQtyAtExam *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMQtyAtExam *pVw = (CEMQtyAtExam *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddEMQtyAtExamFnc(CWnd *pWnd){
	 return ((CEMQtyAtExam*)pWnd)->OnAddEMQtyAtExam();
} 
static int _OnEditEMQtyAtExamFnc(CWnd *pWnd){
	 return ((CEMQtyAtExam*)pWnd)->OnEditEMQtyAtExam();
} 
static int _OnDeleteEMQtyAtExamFnc(CWnd *pWnd){
	 return ((CEMQtyAtExam*)pWnd)->OnDeleteEMQtyAtExam();
} 
static int _OnSaveEMQtyAtExamFnc(CWnd *pWnd){
	 return ((CEMQtyAtExam*)pWnd)->OnSaveEMQtyAtExam();
} 
static int _OnCancelEMQtyAtExamFnc(CWnd *pWnd){
	 return ((CEMQtyAtExam*)pWnd)->OnCancelEMQtyAtExam();
} 
CEMQtyAtExam::CEMQtyAtExam(CWnd *pParent)
{
	m_nDlgWidth = 435;
	m_nDlgHeight = 185;
	SetDefaultValues();
}
CEMQtyAtExam::~CEMQtyAtExam(){
}
void CEMQtyAtExam::OnCreateComponents()
{
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 210, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 420, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 210, 85); 
	m_wndReportCondition.Create(this, _T("Report Condition"), 4, 3, 424, 149);
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 420, 85); 
	m_wndRoomLabel.Create(this, _T("Room"), 10, 90, 90, 115);
	m_wndRoom.Create(this,95, 90, 420, 115); 
	m_wndDoctorLabel.Create(this, _T("Doctor"), 10, 120, 90, 145);
	m_wndDoctor.Create(this,95, 120, 420, 145); 
	m_wndPrint.Create(this, _T("&Print"), 245, 155, 330, 180);
	m_wndExport.Create(this, _T("&Export"), 335, 155, 420, 180);

}
void CEMQtyAtExam::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
//	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndRoom.SetCheckValue(true);
	m_wndRoom.LimitText(1024);
	m_wndDoctor.SetCheckValue(true);
	m_wndDoctor.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Name"), CFMT_TEXT, 80);

	m_wndRoom.InsertColumn(0, _T("ID"), CFMT_TEXT|CFMT_RIGHT, 80);
	m_wndRoom.SetCheckBox(true);
	m_wndRoom.InsertColumn(1, _T("Name"), CFMT_TEXT, 220);

	m_wndDoctor.InsertColumn(0, _T("ID"), CFMT_TEXT|FMT_RIGHT, 80);
	m_wndDoctor.SetCheckBox(true);
	m_wndDoctor.InsertColumn(1, _T("Name"), CFMT_TEXT, 220);

}
void CEMQtyAtExam::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
	m_wndRoom.SetEvent(WE_SELENDOK, _OnRoomSelendokFnc);
	//m_wndRoom.SetEvent(WE_SETFOCUS, _OnRoomSetfocusFnc);
	//m_wndRoom.SetEvent(WE_KILLFOCUS, _OnRoomKillfocusFnc);
	m_wndRoom.SetEvent(WE_SELCHANGE, _OnRoomSelectChangeFnc);
	m_wndRoom.SetEvent(WE_LOADDATA, _OnRoomLoadDataFnc);
	//m_wndRoom.SetEvent(WE_ADDNEW, _OnRoomAddNewFnc);
	m_wndDoctor.SetEvent(WE_SELENDOK, _OnDoctorSelendokFnc);
	//m_wndDoctor.SetEvent(WE_SETFOCUS, _OnDoctorSetfocusFnc);
	//m_wndDoctor.SetEvent(WE_KILLFOCUS, _OnDoctorKillfocusFnc);
	m_wndDoctor.SetEvent(WE_SELCHANGE, _OnDoctorSelectChangeFnc);
	m_wndDoctor.SetEvent(WE_LOADDATA, _OnDoctorLoadDataFnc);
	//m_wndDoctor.SetEvent(WE_ADDNEW, _OnDoctorAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	/*AddEvent(1, _T("Add	Ctrl+A"), _OnAddEMQtyAtExamFnc, 0, 'A', VK_CONTROL);
	AddEvent(2, _T("Edit	Ctrl+E"), _OnEditEMQtyAtExamFnc, 0, 'E', VK_CONTROL);
	AddEvent(3, _T("Delete	Ctrl+D"), _OnDeleteEMQtyAtExamFnc, 0, 'D', VK_CONTROL);
	AddEvent(4, _T("Save	Ctrl+S"), _OnSaveEMQtyAtExamFnc, 0, 'S', VK_CONTROL);
	AddEvent(0, _T("-"));
	AddEvent(5, _T("Cancel	Ctrl+T"), _OnCancelEMQtyAtExamFnc, 0, 'T', VK_CONTROL);*/
//	SetMode(VM_NONE);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);
	SetMode(VM_EDIT);
	
}
void CEMQtyAtExam::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndRoom.GetDlgCtrlID(), m_szRoomKey);
	DDX_TextEx(pDX, m_wndDoctor.GetDlgCtrlID(), m_szDoctorKey);

}
void CEMQtyAtExam::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMQtyAtExam::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMQtyAtExam::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szRoomKey.Empty();
	m_szDoctorKey.Empty();

}
int CEMQtyAtExam::SetMode(int nMode)
{
 		int nOldMode = GetMode();
 		CGuiView::SetMode(nMode);
 		CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 		CString szSQL;
 		CRecord rs(&pMF->m_db);
  		switch(nMode){
 		case VM_ADD: 
 			EnableControls(TRUE);
 			EnableButtons(TRUE, 3,4, -1);
 			SetDefaultValues();
 			break;
 		case VM_EDIT: 
 			EnableControls(TRUE);
 			EnableButtons(TRUE,0,1,2,3, 4, -1); 
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
/*void CEMQtyAtExam::OnYearChange(){
	
} */
/*void CEMQtyAtExam::OnYearSetfocus(){
	
} */
/*void CEMQtyAtExam::OnYearKillfocus(){
	
} */
int CEMQtyAtExam::OnYearCheckValue(){
	return 0;
} 
void CEMQtyAtExam::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMQtyAtExam::OnReportPeriodSelendok()
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
/*void CEMQtyAtExam::OnReportPeriodSetfocus(){
	
}*/
/*void CEMQtyAtExam::OnReportPeriodKillfocus(){
	
}*/
long CEMQtyAtExam::OnReportPeriodLoadData(){
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
/*void CEMQtyAtExam::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMQtyAtExam::OnFromDateChange(){
	
} */
/*void CEMQtyAtExam::OnFromDateSetfocus(){
	
} */
/*void CEMQtyAtExam::OnFromDateKillfocus(){
	
} */
int CEMQtyAtExam::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMQtyAtExam::OnToDateChange(){
	
} */
/*void CEMQtyAtExam::OnToDateSetfocus(){
	
} */
/*void CEMQtyAtExam::OnToDateKillfocus(){
	
} */
int CEMQtyAtExam::OnToDateCheckValue(){
	return 0;
} 
void CEMQtyAtExam::OnRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMQtyAtExam::OnRoomSelendok(){
	 
}
/*void CEMQtyAtExam::OnRoomSetfocus(){
	
}*/
/*void CEMQtyAtExam::OnRoomKillfocus(){
	
}*/
long CEMQtyAtExam::OnRoomLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();
	if(m_wndRoom.IsSearchKey() && !m_szRoomKey.IsEmpty())
	{
		szWhere.Format(_T("and hrl_id = %d"), ToInt(m_szRoomKey));
	}
	m_wndRoom.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("select hrl_name as name, hrl_id as id from hms_roomlist where hrl_deptid = 'C1.2' %s Order by id"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndRoom.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CEMQtyAtExam::OnRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMQtyAtExam::OnDoctorSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMQtyAtExam::OnDoctorSelendok(){
	 
}
/*void CEMQtyAtExam::OnDoctorSetfocus(){
	
}*/
/*void CEMQtyAtExam::OnDoctorKillfocus(){
	
}*/
long CEMQtyAtExam::OnDoctorLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDoctor.IsSearchKey() && !m_szDoctorKey.IsEmpty())
	{
		szWhere.Format(_T(" and su_userid='%s' "), m_szDoctorKey);
	}
	m_wndDoctor.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("select su_userid as id, su_name as name from sys_user  where  (su_deptid = 'C1.2' or su_hms_xdept = 'C1.2') and su_groupid in('D', 'P') %s order by su_name "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF())
	{ 
	m_wndDoctor.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CEMQtyAtExam::OnDoctorAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMQtyAtExam::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString szSQL, szTemp, tmpStr;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
		
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	if(!rpt.Init(_T("Reports/HMS/HE_SOLUONGKHAMTAICACPHONG.RPT")) )
		return ;

	rpt.GetReportHeader()->SetValue(_T("HeathService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	int nIndex = 1;
	CReportSection* rptDetail;
	int  nTotal1 = 0, nTotal2 = 0, nTotal3 = 0;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		szTemp = rs.GetValue(_T("doctorname"));
		rptDetail->SetValue(_T("2"), szTemp);

		szTemp = rs.GetValue(_T("bhquan"));
		nTotal1 += ToInt(szTemp);
		rptDetail->SetValue(_T("3"), szTemp);

		szTemp = rs.GetValue(_T("bhhuu"));
		nTotal2 += ToInt(szTemp);
		rptDetail->SetValue(_T("4"), szTemp);

		szTemp = rs.GetValue(_T("tongso"));
		nTotal3 += ToInt(szTemp);
		rptDetail->SetValue(_T("5"), szTemp);

		rs.MoveNext();
	}	
	rpt.GetReportFooter()->SetValue(_T("s3"), nTotal1);
	rpt.GetReportFooter()->SetValue(_T("s4"), nTotal2);
	rpt.GetReportFooter()->SetValue(_T("s5"), nTotal3);

	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
	EndWaitCursor();
		
	
} 
void CEMQtyAtExam::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, szTemp, tmpStr;
	CExcel xls;
	BeginWaitCursor();
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 30);
	xls.SetColumnWidth(2, 15 );
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 15);

	int nRow = 0, nCol = 0;
	
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellMergedColumns(nCol, nRow + 1, 4);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(nCol, nRow + 2, 4);
	xls.SetCellMergedColumns(nCol, nRow + 3, 4);
	xls.SetCellText(nCol, nRow + 2, _T("KH\xC1M T\x1EA0I \x43\xC1\x43 PH\xD2NG TH\x45O \x42\xC1\x43 S\x1EF8"), FMT_TEXT | FMT_CENTER, true, 13);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y: %s \x110\x1EBFn ng\xE0y: %s"), 
		          CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss),
				  CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellText(nCol, nRow + 4, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 4, _T("\x42\xE1\x63 s\x1EF9"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 4, _T("Ph\xF2ng kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 3, nRow + 4, _T("Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 4, _T("H\x1B0u"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 5, nRow + 4, _T("T\x1ED5ng"), FMT_TEXT | FMT_CENTER, true, 11);
	
	nRow += 5;
	int nIndex = 1;
	int nTotal1 = 0,nTotal2 = 0, nTotal3 = 0 ;
	while (!rs.IsEOF())
	{
		szTemp.Format(_T("%d"), nIndex);
		xls.SetCellText(nCol, nRow, szTemp, FMT_INTEGER);

		szTemp = rs.GetValue(_T("doctorname"));
		xls.SetCellText(nCol + 1, nRow, szTemp, FMT_TEXT);

		szTemp = rs.GetValue(_T("bhquan"));
		nTotal1 += ToInt(szTemp);
		xls.SetCellText(nCol + 3, nRow, szTemp, FMT_NUMBER1);

		szTemp = rs.GetValue(_T("bhhuu"));
		nTotal2 += ToInt(szTemp);
		xls.SetCellText(nCol + 4, nRow, szTemp, FMT_NUMBER1);

		szTemp = rs.GetValue(_T("tongso"));
		nTotal3 += ToInt(szTemp);
		xls.SetCellText(nCol + 5, nRow, szTemp, FMT_NUMBER1);

		nIndex++;
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellText(nCol + 3, nRow, int2str(nTotal1), FMT_NUMBER1, true, 12 );
	xls.SetCellText(nCol + 4, nRow, int2str(nTotal2), FMT_NUMBER1, true, 12 );
	xls.SetCellText(nCol + 5, nRow, int2str(nTotal3), FMT_NUMBER1, true, 12 );
	EndWaitCursor();
	xls.Save(_T("Exports\\Tinh hinh kham benh theo bac sy.xls"));
} 
int CEMQtyAtExam::OnAddEMQtyAtExam(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CEMQtyAtExam::OnEditEMQtyAtExam(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMQtyAtExam::OnDeleteEMQtyAtExam(){
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
 		OnCancelEMQtyAtExam();
 	}
	return 0;
}
int CEMQtyAtExam::OnSaveEMQtyAtExam(){
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
 		//OnEMQtyAtExamListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CEMQtyAtExam::OnCancelEMQtyAtExam(){
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
int CEMQtyAtExam::OnEMQtyAtExamListLoadData(){
	return 0;
}
CString CEMQtyAtExam::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere, szWhere1;
	CString szDoctor,szRoom;
	
	
	for (int i = 0; i < m_wndDoctor.GetCount(); i++)
	{
		if (m_wndDoctor.GetCheck(i))
		{
			if (!szDoctor.IsEmpty())
				szDoctor += _T(", ");
			szDoctor.AppendFormat(_T("'%s'"), m_wndDoctor.GetCheck(i, 0));
		}
	}
	if (!szDoctor.IsEmpty())
		szWhere.AppendFormat(_T(" AND he_doctor in (%s) "), szDoctor);


	for (int i = 0; i < m_wndRoom.GetCount(); i++)
	{
		if (m_wndRoom.GetCheck(i))
		{
			if (!szRoom.IsEmpty())
				szRoom += _T(", ");
			szRoom.AppendFormat(_T("'%s'"), m_wndRoom.GetCheck(i, 0));
		}
	}
	if (!szRoom.IsEmpty())
		szWhere.AppendFormat(_T(" AND he_roomid in (%s) "), szRoom);

		szSQL.Format(_T(" select doctor, doctorname,") \
					_T("       sum(bhhuu) as bhhuu, sum(bhquan) as bhquan,") \
					_T("       sum(bhhuu) + sum(bhquan) as tongso") \
					_T(" from") \
					_T(" (") \
					_T(" select  distinct hd_docno, ") \
					_T("          he_doctor as doctor, he_status, ") \
					_T("          (select su_name from sys_user where he_doctor = su_userid ") \
					_T("		  and  (su_deptid = he_deptid or su_hms_xdept = he_deptid) and su_groupid in('P', 'D') and he_deptid = 'C1.2') as doctorname, ") \
					_T("          hd_object as obj, ") \
					_T("         case when hd_object in ('2') and substr(hd_cardno, 1, 2) = 'HT' then 1 else 0 end as bhhuu,") \
					_T("          case when hd_object in ('1') then 1 else 0 end as bhquan, ") \
					_T("          he_roomid,") \
					_T("         (select hrl_name from hms_roomlist where hrl_id = he_roomid and he_deptid = hrl_deptid) as roomname") \
					_T(" from hms_exam") \
					_T(" left join hms_doc on (he_docno = hd_docno)") \
					_T(" where  1= 1") \
					_T("       and he_status = 'T' ") \
					_T("       and he_deptid = 'C1.2'") \
					_T("      AND he_examdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s') %s") \
					_T(" )") \
					_T(" where   length(doctorname)>0") \
					_T(" group by doctor, doctorname") \
					_T(" order by doctorname"),m_szFromDate, m_szToDate, szWhere);
	
	return szSQL;
}