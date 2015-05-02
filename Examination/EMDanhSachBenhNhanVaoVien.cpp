#include "stdafx.h"
#include "EMDanhSachBenhNhanVaoVien.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMDanhSachBenhNhanVaoVien* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnToDateCheckValue();
} 
static void _OnRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMDanhSachBenhNhanVaoVien* )pWnd)->OnRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnRoomSelendokFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnRoomSelendok();
}
/*static void _OnRoomSetfocusFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnRoomKillfocus();
}*/
/*static void _OnRoomKillfocusFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnRoomKillfocus();
}*/
static long _OnRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnRoomLoadData();
}
/*static void _OnRoomAddNewFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnRoomAddNew();
}*/
static void _OnDoctorSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMDanhSachBenhNhanVaoVien* )pWnd)->OnDoctorSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDoctorSelendokFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnDoctorSelendok();
}
/*static void _OnDoctorSetfocusFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnDoctorKillfocus();
}*/
/*static void _OnDoctorKillfocusFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnDoctorKillfocus();
}*/
static long _OnDoctorLoadDataFnc(CWnd *pWnd){
	return ((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnDoctorLoadData();
}
/*static void _OnDoctorAddNewFnc(CWnd *pWnd){
	((CEMDanhSachBenhNhanVaoVien *)pWnd)->OnDoctorAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMDanhSachBenhNhanVaoVien *pVw = (CEMDanhSachBenhNhanVaoVien *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMDanhSachBenhNhanVaoVien *pVw = (CEMDanhSachBenhNhanVaoVien *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddEMDanhSachBenhNhanVaoVienFnc(CWnd *pWnd){
	 return ((CEMDanhSachBenhNhanVaoVien*)pWnd)->OnAddEMDanhSachBenhNhanVaoVien();
} 
static int _OnEditEMDanhSachBenhNhanVaoVienFnc(CWnd *pWnd){
	 return ((CEMDanhSachBenhNhanVaoVien*)pWnd)->OnEditEMDanhSachBenhNhanVaoVien();
} 
static int _OnDeleteEMDanhSachBenhNhanVaoVienFnc(CWnd *pWnd){
	 return ((CEMDanhSachBenhNhanVaoVien*)pWnd)->OnDeleteEMDanhSachBenhNhanVaoVien();
} 
static int _OnSaveEMDanhSachBenhNhanVaoVienFnc(CWnd *pWnd){
	 return ((CEMDanhSachBenhNhanVaoVien*)pWnd)->OnSaveEMDanhSachBenhNhanVaoVien();
} 
static int _OnCancelEMDanhSachBenhNhanVaoVienFnc(CWnd *pWnd){
	 return ((CEMDanhSachBenhNhanVaoVien*)pWnd)->OnCancelEMDanhSachBenhNhanVaoVien();
} 
CEMDanhSachBenhNhanVaoVien::CEMDanhSachBenhNhanVaoVien(CWnd *pParent)
{
	m_nDlgWidth = 435;
	m_nDlgHeight = 185;
	SetDefaultValues();
}
CEMDanhSachBenhNhanVaoVien::~CEMDanhSachBenhNhanVaoVien(){
}
void CEMDanhSachBenhNhanVaoVien::OnCreateComponents()
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
	m_wndObjectLabel.Create(this, _T("\x110\x1ED1i t\x1B0\x1EE3ng"), 10, 120, 90, 145);
	m_wndObject.Create(this,95, 120, 420, 145); 
	//m_wndPrint.Create(this, _T("&Print"), 245, 155, 330, 180);
	m_wndExport.Create(this, _T("&Export"), 335, 155, 420, 180);

}
void CEMDanhSachBenhNhanVaoVien::OnInitializeComponents()
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
	m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Name"), CFMT_TEXT, 80);

	m_wndRoom.InsertColumn(0, _T("ID"), CFMT_TEXT|CFMT_RIGHT, 80);
	m_wndRoom.SetCheckBox(true);
	m_wndRoom.InsertColumn(1, _T("Name"), CFMT_TEXT, 220);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_TEXT|FMT_RIGHT, 80);
	m_wndObject.SetCheckBox(true);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 220);

}
void CEMDanhSachBenhNhanVaoVien::OnSetWindowEvents(){
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
	m_wndObject.SetEvent(WE_SELENDOK, _OnDoctorSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnDoctorSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnDoctorKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnDoctorSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnDoctorLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnDoctorAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	/*AddEvent(1, _T("Add	Ctrl+A"), _OnAddEMDanhSachBenhNhanVaoVienFnc, 0, 'A', VK_CONTROL);
	AddEvent(2, _T("Edit	Ctrl+E"), _OnEditEMDanhSachBenhNhanVaoVienFnc, 0, 'E', VK_CONTROL);
	AddEvent(3, _T("Delete	Ctrl+D"), _OnDeleteEMDanhSachBenhNhanVaoVienFnc, 0, 'D', VK_CONTROL);
	AddEvent(4, _T("Save	Ctrl+S"), _OnSaveEMDanhSachBenhNhanVaoVienFnc, 0, 'S', VK_CONTROL);
	AddEvent(0, _T("-"));
	AddEvent(5, _T("Cancel	Ctrl+T"), _OnCancelEMDanhSachBenhNhanVaoVienFnc, 0, 'T', VK_CONTROL);*/
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
void CEMDanhSachBenhNhanVaoVien::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndRoom.GetDlgCtrlID(), m_szRoomKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
}
void CEMDanhSachBenhNhanVaoVien::GetDataToScreen()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMDanhSachBenhNhanVaoVien::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMDanhSachBenhNhanVaoVien::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szRoomKey.Empty();
	m_szObjectKey.Empty();

}
int CEMDanhSachBenhNhanVaoVien::SetMode(int nMode)
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
/*void CEMDanhSachBenhNhanVaoVien::OnYearChange(){
	
} */
/*void CEMDanhSachBenhNhanVaoVien::OnYearSetfocus(){
	
} */
/*void CEMDanhSachBenhNhanVaoVien::OnYearKillfocus(){
	
} */
int CEMDanhSachBenhNhanVaoVien::OnYearCheckValue(){
	return 0;
} 
void CEMDanhSachBenhNhanVaoVien::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMDanhSachBenhNhanVaoVien::OnReportPeriodSelendok()
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
/*void CEMDanhSachBenhNhanVaoVien::OnReportPeriodSetfocus(){
	
}*/
/*void CEMDanhSachBenhNhanVaoVien::OnReportPeriodKillfocus(){
	
}*/
long CEMDanhSachBenhNhanVaoVien::OnReportPeriodLoadData(){
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
/*void CEMDanhSachBenhNhanVaoVien::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMDanhSachBenhNhanVaoVien::OnFromDateChange(){
	
} */
/*void CEMDanhSachBenhNhanVaoVien::OnFromDateSetfocus(){
	
} */
/*void CEMDanhSachBenhNhanVaoVien::OnFromDateKillfocus(){
	
} */
int CEMDanhSachBenhNhanVaoVien::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMDanhSachBenhNhanVaoVien::OnToDateChange(){
	
} */
/*void CEMDanhSachBenhNhanVaoVien::OnToDateSetfocus(){
	
} */
/*void CEMDanhSachBenhNhanVaoVien::OnToDateKillfocus(){
	
} */
int CEMDanhSachBenhNhanVaoVien::OnToDateCheckValue(){
	return 0;
} 
void CEMDanhSachBenhNhanVaoVien::OnRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMDanhSachBenhNhanVaoVien::OnRoomSelendok(){
	 
}
/*void CEMDanhSachBenhNhanVaoVien::OnRoomSetfocus(){
	
}*/
/*void CEMDanhSachBenhNhanVaoVien::OnRoomKillfocus(){
	
}*/
long CEMDanhSachBenhNhanVaoVien::OnRoomLoadData()
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
/*void CEMDanhSachBenhNhanVaoVien::OnRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMDanhSachBenhNhanVaoVien::OnDoctorSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMDanhSachBenhNhanVaoVien::OnDoctorSelendok(){
	 
}
/*void CEMDanhSachBenhNhanVaoVien::OnDoctorSetfocus(){
	
}*/
/*void CEMDanhSachBenhNhanVaoVien::OnDoctorKillfocus(){
	
}*/
long CEMDanhSachBenhNhanVaoVien::OnDoctorLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty())
	{
	}
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("select ho_id as id, ho_desc as name from hms_object where ho_id in ('1', '2', '11') "));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF())
	{ 
	m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CEMDanhSachBenhNhanVaoVien::OnDoctorAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMDanhSachBenhNhanVaoVien::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
} 
void CEMDanhSachBenhNhanVaoVien::OnExportSelect(){
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
	xls.SetColumnWidth(1, 20);
	xls.SetColumnWidth(2, 10 );
	xls.SetColumnWidth(3, 13);
	xls.SetColumnWidth(4, 20);
	xls.SetColumnWidth(5, 18);
	xls.SetColumnWidth(6, 10);
	xls.SetColumnWidth(7, 20);
	xls.SetColumnWidth(8, 15);
	xls.SetColumnWidth(9, 40);

	int nRow = 0, nCol = 0;
	
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellMergedColumns(nCol, nRow + 1, 4);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(nCol, nRow + 2, 4);
	xls.SetCellMergedColumns(nCol, nRow + 3, 4);
	xls.SetCellText(nCol, nRow + 2, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N V\xC0O VI\x1EC6N"), FMT_TEXT | FMT_CENTER, true, 13);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y: %s \x110\x1EBFn ng\xE0y: %s"), 
		          CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss),
				  CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellText(nCol, nRow + 6, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 6, _T("T\xEAn \x62\x1EC7nh nh\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 6, _T("N\x103m sinh"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 3, nRow + 6, _T("\x43\x1EA5p \x62\x1EAD\x63"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 6, _T("\x110\x1A1n v\x1ECB"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 5, nRow + 6, _T("\x110\x1ED1i t\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 6, nRow + 6, _T("M\xE3 I\x43\x44"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 7, nRow + 6, _T("\x43h\x1EA9n \x111o\xE1n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 8, nRow + 6, _T("Ng\xE0y kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 9, nRow + 6, _T("V\xE0o kho\x61"), FMT_TEXT | FMT_CENTER, true, 11);

	int nIndex = 1;
	int nIdx = 0;
	int c1 = 0, c2 = 0, c3 = 0, c4 = 0, c5 = 0;
	nRow = 7;
	while (!rs.IsEOF())
	{
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		
		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("yob"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER);
		
		rs.GetValue(_T("rank"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("workplace"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("obj"), tmpStr);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("hd_icd"), tmpStr);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("diagno"), tmpStr);
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("he_examdate"), tmpStr);
		xls.SetCellText(nCol + 8, nRow, CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy), FMT_TEXT);
		
		rs.GetValue(_T("deptname"), tmpStr);
		xls.SetCellText(nCol + 9, nRow, tmpStr, FMT_TEXT);

		
		rs.GetValue(_T("hd_object"), tmpStr);
		if(tmpStr == _T('1')) 
			{
				c1++;
			}
		else if(tmpStr == _T('2')) 
			{
				c2++;
			}
		else if(tmpStr == _T("11"))
			{
				c3++;
			}

		nRow++;
		rs.MoveNext();
	}

	szTemp.Format(_T("V\xE0o vi\x1EC7n Qu\xE2n: %d"),c1++);
	xls.SetCellText(1, 4, szTemp, FMT_TEXT, true, 12);

	xls.SetCellMergedColumns(2, 4, 2);
	szTemp.Format(_T("V\xE0o vi\x1EC7n H\x1B0u: %d"),c2++);
	xls.SetCellText(2, 4, szTemp, FMT_TEXT, true, 12);

	xls.SetCellMergedColumns(4, 4, 2);
	szTemp.Format(_T("VV \x42HYT Qu\xE2n nh\xE2n: %d"),c3++);
	xls.SetCellText(4, 4, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("T\x1ED5ng s\x1ED1: %d"),nIdx++);
	xls.SetCellText(6, 4, szTemp, FMT_TEXT, true, 12);

	EndWaitCursor();
	xls.Save(_T("Exports\\Danh Sach Benh Nhan Vao Vien Theo Khoa.xls"));
} 
int CEMDanhSachBenhNhanVaoVien::OnAddEMDanhSachBenhNhanVaoVien(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CEMDanhSachBenhNhanVaoVien::OnEditEMDanhSachBenhNhanVaoVien(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMDanhSachBenhNhanVaoVien::OnDeleteEMDanhSachBenhNhanVaoVien(){
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
 		OnCancelEMDanhSachBenhNhanVaoVien();
 	}
	return 0;
}
int CEMDanhSachBenhNhanVaoVien::OnSaveEMDanhSachBenhNhanVaoVien(){
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
 		//OnEMDanhSachBenhNhanVaoVienListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CEMDanhSachBenhNhanVaoVien::OnCancelEMDanhSachBenhNhanVaoVien(){
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
int CEMDanhSachBenhNhanVaoVien::OnEMDanhSachBenhNhanVaoVienListLoadData(){
	return 0;
}
CString CEMDanhSachBenhNhanVaoVien::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere, szWhere1;
	CString szObject,szRoom;
	
	
	for (int i = 0; i < m_wndObject.GetCount(); i++)
	{
		if (m_wndObject.GetCheck(i))
		{
			if (!szObject.IsEmpty())
				szObject += _T(", ");
			szObject.AppendFormat(_T("'%s'"), m_wndObject.GetCheck(i, 0));
		}
	}
	if (!szObject.IsEmpty())
		szWhere.AppendFormat(_T(" AND hd_object in (%s) "), szObject);


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

		szSQL.Format(_T(" SELECT  distinct he_docno as docno, he_examdate,") \
					_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,   ") \
					_T("        extract(year from hp_birthdate) as yob, get_objectname(hd_object) as obj,hd_object,    ") \
					_T("        (select distinct ss_desc from sys_sel where ss_id = 'hms_rank' and ss_code = hp_rank) as rank, ") \
					_T("        hp_workplace as workplace,        ") \
					_T("        hd_diagnostic as diagno, ") \
					_T("        hd_icd, hd_indept, (select sd_name from sys_dept where sd_id = hd_indept) as deptname") \
					_T(" FROM hms_patient ") \
					_T(" LEFT JOIN hms_doc ON (hd_patientno = hp_patientno) ") \
					_T(" LEFT JOIN hms_exam ON (he_docno = hd_docno)") \
					_T(" LEFT JOIN hms_pharmaorder ON (hpo_docno = he_docno) ") \
					_T(" WHERE he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s ") \
					_T("       AND he_status <> 'O'") \
					_T("       AND he_deptid = 'C1.2'") \
					_T("       AND hd_suggestion = 'D'") \
					_T("       AND length(hd_indept) > 0 ") \
					_T(" order by docno, he_examdate"), m_szFromDate, m_szToDate, szWhere);

	return szSQL;
}