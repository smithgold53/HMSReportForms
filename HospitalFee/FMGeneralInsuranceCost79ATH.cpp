#include "stdafx.h"
#include "FMGeneralInsuranceCost79ATH.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMGeneralInsuranceCost79ATH* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMGeneralInsuranceCost79ATH* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH *)pWnd)->OnClerkAddNew();
}*/
static void _OnWithoutEmergencySelectFnc(CWnd *pWnd){
	 ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnWithoutEmergencySelect();
}
static void _OnOnlyEmergencySelectFnc(CWnd *pWnd){
	 ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnOnlyEmergencySelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMGeneralInsuranceCost79ATH *pVw = (CFMGeneralInsuranceCost79ATH *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMGeneralInsuranceCost79ATH *pVw = (CFMGeneralInsuranceCost79ATH *)pWnd;
	pVw->OnExportSelect();
}
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMGeneralInsuranceCost79ATH*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnListDeleteItem();
}
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMGeneralInsuranceCost79ATH*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnDeptListDeleteItem();
} 
static long _OnRoomListLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnRoomListLoadData();
} 
static void _OnRoomListDblClickFnc(CWnd *pWnd){
	((CFMGeneralInsuranceCost79ATH*)pWnd)->OnRoomListDblClick();
} 
static void _OnRoomListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMGeneralInsuranceCost79ATH*)pWnd)->OnRoomListSelectChange(nOldItem, nNewItem);
} 
static int _OnRoomListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnRoomListDeleteItem();
}
static int _OnAddFMGeneralInsuranceCost79ATHFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnAddFMGeneralInsuranceCost79ATH();
} 
static int _OnEditFMGeneralInsuranceCost79ATHFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnEditFMGeneralInsuranceCost79ATH();
} 
static int _OnDeleteFMGeneralInsuranceCost79ATHFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnDeleteFMGeneralInsuranceCost79ATH();
} 
static int _OnSaveFMGeneralInsuranceCost79ATHFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnSaveFMGeneralInsuranceCost79ATH();
} 
static int _OnCancelFMGeneralInsuranceCost79ATHFnc(CWnd *pWnd){
	 return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnCancelFMGeneralInsuranceCost79ATH();
}
static int _OnListCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnListCheckAll();
}
static int _OnListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnListUnCheckAll();
}
static int _OnRoomListCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnRoomListCheckAll();
}
static int _OnRoomListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnRoomListUnCheckAll();
}
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMGeneralInsuranceCost79ATH*)pWnd)->OnDeptListUnCheckAll();
}
CFMGeneralInsuranceCost79ATH::CFMGeneralInsuranceCost79ATH(CWnd *pParent)
{
	m_nDlgWidth = 450;
	m_nDlgHeight = 575;
	SetDefaultValues();
}
CFMGeneralInsuranceCost79ATH::~CFMGeneralInsuranceCost79ATH(){
}
void CFMGeneralInsuranceCost79ATH::OnCreateComponents()
{
	/*m_wndReportFilter.Create(this, _T("Report Parameters"), 5, 5, 490, 180);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 245, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Month"), 250, 30, 330, 55);
	m_wndReportPeriod.Create(this,335, 30, 485, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 245, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 251, 60, 331, 85);
	m_wndToDate.Create(this,335, 60, 485, 85); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 90, 115);
	m_wndClerk.Create(this,95, 90, 485, 115); 
	m_wndWithoutEmergency.Create(this, _T("Without Emergency"), 95, 120, 485, 145);
	m_wndOnlyEmergency.Create(this, _T("Only Emergency"), 95, 150, 485, 175);
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 300, 185, 410, 210);
	m_wndExport.Create(this, _T("&Export"), 415, 185, 490, 210);*/

	m_wndReportFilter.Create(this, _T("Report Parameters"), 5, 5, 440, 535);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 120, 435, 230);
	m_wndExamRoomInfo.Create(this, _T("Exam Room"), 10, 235, 435, 470);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 310, 55);
	m_wndReportPeriod.Create(this,315, 30, 435, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 310, 85);
	m_wndToDate.Create(this,315, 60, 435, 85); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 90, 115);
	m_wndClerk.Create(this,95, 90, 435, 115); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 235, 540, 335, 565);
	m_wndExport.Create(this, _T("&Export"), 340, 540, 440, 565);
	m_wndWithoutEmergency.Create(this, _T("Without Emergency"), 170, 475, 435, 500);
	m_wndOnlyEmergency.Create(this, _T("Only Emergency"), 170, 505, 435, 530);
	m_wndDeptList.Create(this,15, 145, 430, 225); 
	m_wndRoomList.Create(this,15, 260, 430, 465); 
}
void CFMGeneralInsuranceCost79ATH::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(65);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 270);

	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.SetCheckBox(TRUE);

	m_wndRoomList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndRoomList.InsertColumn(1, _T("Name"), CFMT_TEXT, 220);
	m_wndRoomList.InsertColumn(2, _T("Dept"), CFMT_TEXT, 80);
	m_wndRoomList.SetCheckBox(TRUE);
}
void CFMGeneralInsuranceCost79ATH::OnSetWindowEvents()
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
	m_wndClerk.SetEvent(WE_SELENDOK, _OnClerkSelendokFnc);
	//m_wndClerk.SetEvent(WE_SETFOCUS, _OnClerkSetfocusFnc);
	//m_wndClerk.SetEvent(WE_KILLFOCUS, _OnClerkKillfocusFnc);
	m_wndClerk.SetEvent(WE_SELCHANGE, _OnClerkSelectChangeFnc);
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	//m_wndClerk.SetEvent(WE_ADDNEW, _OnClerkAddNewFnc);
	m_wndWithoutEmergency.SetEvent(WE_CLICK, _OnWithoutEmergencySelectFnc);
	m_wndOnlyEmergency.SetEvent(WE_CLICK, _OnOnlyEmergencySelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);

	m_wndList.SetWindowText(_T("Dept"));
	m_wndList.AddEvent(1, _T("Check All"), _OnListCheckAllFnc);
	m_wndList.AddEvent(2, _T("UnCheck All"), _OnListUnCheckAllFnc);

	m_wndDeptList.SetEvent(WE_SELCHANGE, _OnDeptListSelectChangeFnc);
	m_wndDeptList.SetEvent(WE_LOADDATA, _OnDeptListLoadDataFnc);
	m_wndDeptList.SetEvent(WE_DBLCLICK, _OnDeptListDblClickFnc);

	m_wndRoomList.SetEvent(WE_SELCHANGE, _OnRoomListSelectChangeFnc);
	m_wndRoomList.SetEvent(WE_LOADDATA, _OnRoomListLoadDataFnc);
	m_wndRoomList.SetEvent(WE_DBLCLICK, _OnRoomListDblClickFnc);

	m_wndRoomList.SetWindowText(_T("Room List"));
	m_wndRoomList.AddEvent(1, _T("Check All"), _OnRoomListCheckAllFnc);
	m_wndRoomList.AddEvent(2, _T("UnCheck All"), _OnRoomListUnCheckAllFnc);

	m_wndDeptList.SetWindowText(_T("Dept List"));
	m_wndDeptList.AddEvent(1, _T("Check All"), _OnDeptListCheckAllFnc);
	m_wndDeptList.AddEvent(2, _T("UnCheck All"), _OnDeptListUnCheckAllFnc);
	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL;
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

	OnDeptListLoadData();
	OnRoomListLoadData();
}
void CFMGeneralInsuranceCost79ATH::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndWithoutEmergency.GetDlgCtrlID(), m_bWithoutEmergency);
	DDX_Check(pDX, m_wndOnlyEmergency.GetDlgCtrlID(), m_bOnlyEmergency);

}
void CFMGeneralInsuranceCost79ATH::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMGeneralInsuranceCost79ATH::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMGeneralInsuranceCost79ATH::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bWithoutEmergency=FALSE;
	m_bOnlyEmergency=FALSE;

}
int CFMGeneralInsuranceCost79ATH::SetMode(int nMode){
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
/*void CFMGeneralInsuranceCost79ATH::OnYearChange(){
	
} */
/*void CFMGeneralInsuranceCost79ATH::OnYearSetfocus(){
	
} */
/*void CFMGeneralInsuranceCost79ATH::OnYearKillfocus(){
	
} */
int CFMGeneralInsuranceCost79ATH::OnYearCheckValue()
{
	UpdateData(TRUE);
	if (m_nYear > 0)
	{
		CDateTime dt;
		CDate date;
		CString szTemp;

		dt.ParseDateTime(m_szFromDate);
		date = dt.GetDate();
		if (date.GetYear() != 1752)
		{
			dt.SetDate(m_nYear, date.GetMonth(), date.GetDay());
			m_szFromDate = dt.GetDateTime();
			szTemp.Format(_T("%.2d/%.2d/%.4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndFromDate.SetWindowText(szTemp);
		}
		dt.ParseDateTime(m_szToDate);
		date = dt.GetDate();
		if (date.GetYear() != 1752)
		{
			dt.SetDate(m_nYear, date.GetMonth(), date.GetDay());
			m_szToDate = dt.GetDateTime();
			szTemp.Format(_T("%.2d/%.2d/%.4d %.2d:%.2d"), dt.GetDate().GetDay(), dt.GetDate().GetMonth(), 
						  dt.GetDate().GetYear(), dt.GetTime().GetHour(), dt.GetTime().GetMinute());
			m_wndToDate.SetWindowText(szTemp);
		}
	}

	UpdateData(FALSE);
	return 0;
} 
void CFMGeneralInsuranceCost79ATH::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMGeneralInsuranceCost79ATH::OnReportPeriodSelendok()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	CString tmpStr;
	CDate dte, date;

	UpdateData(true);

	date.ParseDate(pMF->GetSysDate());
	int nYear = date.GetYear();
	int nMonth = ToInt(m_szReportPeriodKey);

	if (nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/01 00:00"), nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), nYear, nMonth, dte.GetMonthLastDay());
	}
	if (nMonth == 13)
	{
		m_szFromDate.Format(_T("%.4d/01/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/03/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 14)
	{
		m_szFromDate.Format(_T("%.4d/04/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/06/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 15)
	{
		m_szFromDate.Format(_T("%.4d/07/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/09/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 16)
	{
		m_szFromDate.Format(_T("%.4d/10/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/10/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}
	if (nMonth == 17)
	{
		m_szFromDate.Format(_T("%.4d/01/01 00:00"), nYear);
		tmpStr.Format(_T("%.4d/12/01"), nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), nYear, dte.GetMonthLastDay());
	}

	UpdateData(false);	 
}
/*void CFMGeneralInsuranceCost79ATH::OnReportPeriodSetfocus(){
	
}*/
/*void CFMGeneralInsuranceCost79ATH::OnReportPeriodKillfocus(){
	
}*/
long CFMGeneralInsuranceCost79ATH::OnReportPeriodLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
/*void CFMGeneralInsuranceCost79ATH::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMGeneralInsuranceCost79ATH::OnFromDateChange(){
	
} */
/*void CFMGeneralInsuranceCost79ATH::OnFromDateSetfocus(){
	
} */
/*void CFMGeneralInsuranceCost79ATH::OnFromDateKillfocus(){
	
} */
int CFMGeneralInsuranceCost79ATH::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMGeneralInsuranceCost79ATH::OnToDateChange(){
	
} */
/*void CFMGeneralInsuranceCost79ATH::OnToDateSetfocus(){
	
} */
/*void CFMGeneralInsuranceCost79ATH::OnToDateKillfocus(){
	
} */
int CFMGeneralInsuranceCost79ATH::OnToDateCheckValue(){
	return 0;
} 
void CFMGeneralInsuranceCost79ATH::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMGeneralInsuranceCost79ATH::OnClerkSelendok(){
	 
}
/*void CFMGeneralInsuranceCost79ATH::OnClerkSetfocus(){
	
}*/
/*void CFMGeneralInsuranceCost79ATH::OnClerkKillfocus(){
	
}*/
long CFMGeneralInsuranceCost79ATH::OnClerkLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if (m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	}
	m_wndClerk.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as id, su_name as name FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMGeneralInsuranceCost79ATH::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMGeneralInsuranceCost79ATH::OnWithoutEmergencySelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMGeneralInsuranceCost79ATH::OnOnlyEmergencySelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMGeneralInsuranceCost79ATH::OnListDblClick(){
	
} 
void CFMGeneralInsuranceCost79ATH::OnListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMGeneralInsuranceCost79ATH::OnListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMGeneralInsuranceCost79ATH::OnListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	m_wndList.BeginLoad(); 
	m_wndList.DeleteAllItems(); 
	int nCount = 0;

	szSQL.Format(_T("SELECT sd_id AS ID, sd_name AS Name ") \
		         _T("FROM sys_dept ") \
				 _T("WHERE sd_type IN('KB') ") \
				 _T("ORDER BY sd_id"));
	nCount = rs.ExecSQL(szSQL);

	while(!rs.IsEOF())
	{ 
		m_wndList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
}
void CFMGeneralInsuranceCost79ATH::OnDeptListDblClick(){
	
} 
void CFMGeneralInsuranceCost79ATH::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMGeneralInsuranceCost79ATH::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMGeneralInsuranceCost79ATH::OnDeptListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndDeptList.BeginLoad(); 
	int nCount = 0;

	szSQL.Format(_T("SELECT sd_id as ID, sd_name as Name ") \
		         _T("FROM sys_dept ") \
				 _T("WHERE sd_type='KB' ") \
				 _T("ORDER BY id "));

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDeptList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndDeptList.EndLoad(); 
	return nCount;
}
void CFMGeneralInsuranceCost79ATH::OnRoomListDblClick(){
	
} 
void CFMGeneralInsuranceCost79ATH::OnRoomListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMGeneralInsuranceCost79ATH::OnRoomListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMGeneralInsuranceCost79ATH::OnRoomListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndRoomList.BeginLoad(); 
	int nCount = 0;

	szSQL.Format(_T(" SELECT hrl_id AS ID,") \
				_T("        hrl_name AS Name, ") \
				_T("        hrl_deptid AS DeptID") \
				_T(" FROM hms_roomlist") \
				_T(" LEFT JOIN sys_dept ON(sd_id=hrl_deptid)") \
				_T(" WHERE sd_type='KB'") \
				_T(" ORDER BY hrl_deptid, hrl_id"));

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndRoomList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), 
			rs.GetValue(_T("DeptID")), NULL);
		rs.MoveNext();
	}
	m_wndRoomList.EndLoad(); 
	return nCount;
}
void CFMGeneralInsuranceCost79ATH::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;

	if (!rpt.Init(_T("Reports/HMS/HF_25aTH.RPT")))
		return;


	TCHAR *lszLine[] ={_T("\x110\xFAng tuy\x1EBFn"), _T("Tr\xE1i tuy\x1EBFn")};
	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	rpt.GetReportHeader()->SetValue(_T("ObjectGroup"), _T(""));
	rpt.GetReportHeader()->SetValue(_T("ReportPeriod"), m_szClerkKey);
	rs.GetValue(_T("invoiceno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	//Page Header
	//Report Detail
	int nIndex = 1;
	CString szOldLine, szNewLine, szOldGroup, szNewGroup, szLineName;
	CString szNewOfLine, szOldOfLine,szSumOfline;
	CReportSection* rptDetail;
	long double grpCost[23], grpLine[23], ttlCost[23], grpOfLine[23];
	double cost = 0;
	for (int i =0; i < 23; i++)
	{
		grpCost[i] = 0;
		grpLine[i] = 0;
		ttlCost[i] = 0;
		grpOfLine[i] =0;
	}
	while(!rs.IsEOF())
	{		
		rs.GetValue(_T("linename"), szNewGroup);
		if(szNewGroup != szOldGroup && !szNewGroup.IsEmpty())
		{
			if (grpCost[14] > 0)
			{	
				CString szField;
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(9);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				tmpStr.Format(_T("\x43\x1ED9ng(%s=(I+II)):"), szOldGroup);
				rptDetail->SetValue(_T("TotalGroup"), tmpStr);
				for (int i =3; i < 17; i++)
				{
					szField.Format(_T("S%d"), i);
					FormatCurrency(grpCost[i], tmpStr);
					rptDetail->SetValue(szField, tmpStr);
					ttlCost[i] += grpCost[i];
					grpCost[i] = 0;
				}
			}
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetFaceSize(9);
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetItalic(true);
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetBold(true);			
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rs.GetValue(_T("insline"), tmpStr);
			rptDetail->SetValue(_T("InsuranceLine"),szNewGroup +_T(". ") + pMF->GetSelectionString(_T("hms_insline"), tmpStr));
			szOldGroup = szNewGroup;			
			nIndex=1;
		}

		rs.GetValue(_T("hdline"), szNewOfLine);
		if (szNewOfLine == _T("0") )
		{
			tmpStr.Format(_T("%s"), lszLine[0]);
			szSumOfline = _T("I");
		}
		else
		{
			tmpStr.Format(_T("%s"), lszLine[1]);
			szSumOfline = _T("II");
		}
			
	
		rptDetail = rpt.AddDetail();		
		rptDetail->SetValue(_T("1"), szSumOfline);
		rptDetail->SetValue(_T("2"), tmpStr);
		
		rs.GetValue(_T("total"),cost);
		grpCost[3] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("testfee"), cost);
		grpCost[4] += cost;
		grpLine[4] += cost;		
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("pacsfee"), cost);
		grpCost[5] += cost;
		grpLine[5] += cost;
		grpOfLine[5] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("drugfee"), cost);
		grpCost[6] += cost;
		grpLine[6] += cost;
		grpOfLine[6] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("bloodfee"), cost);
		grpCost[7] += cost;
		grpLine[7] += cost;
		grpOfLine[7] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);				

		rs.GetValue(_T("normtechfee"), cost);
		grpCost[8] += cost;
		grpLine[8] += cost;
		grpOfLine[8] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		
		rs.GetValue(_T("materialfee"), cost);
		grpCost[9] += cost;
		grpLine[9] += cost;
		grpOfLine[9] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("hitechfee"), cost);
		grpCost[10] += cost;
		grpLine[10] += cost;
		grpOfLine[10] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);		

		rs.GetValue(_T("drugfeek"), cost);
		grpCost[11] += cost;
		grpLine[11] += cost;
		grpOfLine[11] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);
		
		rs.GetValue(_T("examfee"), cost);
		grpCost[12] += cost;
		grpLine[12] += cost;
		grpOfLine[12] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);

		rs.GetValue(_T("otherfee"), cost);;
		grpCost[13] += cost;
		grpLine[13] += cost;
		grpOfLine[13] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);		
		
		rs.GetValue(_T("cost"), cost);
		grpCost[14] += cost;
		grpLine[14] += cost;
		grpOfLine[14] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("14"), tmpStr);

		rs.GetValue(_T("patpaid"), cost);
		grpCost[15] += cost;
		grpLine[15] += cost;
		grpOfLine[15] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("15"), tmpStr);

		rs.GetValue(_T("inspaid"), cost);
		grpCost[16] += cost;
		grpLine[16] += cost;
		grpOfLine[16] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("16"), tmpStr);
		
		rs.MoveNext();
	}

	for (int i =3; i < 17; i++)
	{
			ttlCost[i] += grpCost[i];
	}
	
	
	if (grpCost[14] > 0)
	{
		CString szField, szAmount;
		TranslateString(_T("Total Line"), szAmount);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(9);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		tmpStr.Format(_T("\x43\x1ED9ng(%s=(I+II)):"), szOldGroup);		
		rptDetail->SetValue(_T("TotalGroup"), tmpStr );
		for (int i =3; i < 17; i++)
		{
			szField.Format(_T("S%d"), i);
			FormatCurrency(grpCost[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);
		}
	
	}
	if (ttlCost[14] > 0)
	{
		CString szField, szAmount;
		TranslateString(_T("Total Amount:"), szAmount);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(9);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(false);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		for (int i =3; i < 17; i++)
		{
			
			szField.Format(_T("S%d"), i);
			FormatCurrency(ttlCost[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);		
		}

	}
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
	EndWaitCursor();
	
} 
void CFMGeneralInsuranceCost79ATH::OnExportSelect()
{
	//CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//CExcel xls;
	//int nCol = 0, nRow = 0;
	//CString tmpStr;
	//xls.CreateSheet(1);
	//xls.SetWorksheet(0);
	//xls.SetCellMergedRows(0, 0, 2);
	//xls.SetCellMergedRows(1, 0, 2);
	//CellFormat cf(&xls);
	//cf.SetFontName(_T("Tahoma"));
	//cf.SetFontSize(20);
	//cf.SetColor(XLCOLOR_BLUE);
	//cf.SetFillPattern(XLFILLPATTERN_GRAY25);
	////cf.SetAlignV(XLALIGNV_CENTER);
	//cf.SetCellStyle(FMT_TEXT|FMT_CENTER|FMT_VCENTER);
	//xls.Save(_T("Exports\\test.xls"));
	//return;
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString tmpStr, szSQL;
	TCHAR *lszLine[] ={_T("\x110\xFAng tuy\x1EBFn"), _T("Tr\xE1i tuy\x1EBFn")};

	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);	

	CellFormat df(&xls), cf(&xls), hf(&xls);
	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	df.SetFontSize(11);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING);
	hf.SetFontSize(11);
	cf.SetFontName(_T("Segoe UI"));
	cf.SetFontSize(11);

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 23);
	xls.SetColumnWidth(2, 9);
	xls.SetColumnWidth(3, 12);
	xls.SetColumnWidth(4, 12);
	xls.SetColumnWidth(5, 12);
	xls.SetColumnWidth(6, 12);
	xls.SetColumnWidth(7, 12);
	xls.SetColumnWidth(8, 12);
	xls.SetColumnWidth(9, 12);
	xls.SetColumnWidth(10, 12);
	xls.SetColumnWidth(11, 12);
	xls.SetColumnWidth(12, 12);
	xls.SetColumnWidth(13, 12);
	xls.SetColumnWidth(14, 12);
	xls.SetColumnWidth(15, 12);
	xls.SetColumnWidth(16, 12);
	//xls.SetColumnWidth(17, 12);
	
	xls.SetCellMergedColumns(0, 0, 4);
	xls.SetCellMergedColumns(0, 1, 4);

	xls.SetCellMergedColumns(11, 0, 6);
	xls.SetCellMergedColumns(11, 1, 6);

	xls.SetCellText(0, 0, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(11, 0, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(11, 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedColumns(0, 2, 17);
	xls.SetCellMergedColumns(0, 3, 17);

	xls.SetCellText(0, 2, _T("\x44\x41NH S\xC1\x43H \x110\x1EC0 NGH\x1ECA TH\x41NH TO\xC1N \x43HI PH\xCD KH\xC1M \x43H\x1EEE\x41 \x42\x1EC6NH NGO\x1EA0I TR\xDA"),
		            FMT_TEXT | FMT_CENTER, true, 16);
	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(0, 3, tmpStr, &df);

	int nRow = 4;

	xls.SetCellText(0, nRow, _T("STT"), &hf);

	TranslateString(_T("KCB Trong k\x1EF3"), tmpStr);
	xls.SetCellText(1, nRow, tmpStr, &hf);

	xls.SetCellText(2, nRow, _T("S\x1ED1 l\x1B0\x1EE3ng"), &hf);
	
	xls.SetCellMergedRows(0, nRow, 2);
	xls.SetCellMergedRows(1, nRow, 2);
	xls.SetCellMergedRows(2, nRow, 2);
		
	//xls.SetCellMergedRows(12, nRow, 2);
	xls.SetCellMergedRows(13, nRow, 2);
	xls.SetCellMergedRows(14, nRow, 2);
	xls.SetCellMergedRows(15, nRow, 2);
	xls.SetCellMergedRows(16, nRow, 2);
	
	xls.SetCellMergedColumns(3, nRow, 10);

	nRow = 5;

	TranslateString(_T("Medical expenses incurred at medical care facilities"), tmpStr);
	xls.SetCellText(3, nRow - 1, tmpStr, &hf);
	xls.SetCellText(3, nRow, _T("XN,CDCN"), &hf);	
	xls.SetCellText(4, nRow, _T("CDHA"), &hf);

	TranslateString(_T("Drug"), tmpStr);
	xls.SetCellText(5, nRow, tmpStr, &hf);

	TranslateString(_T("Blood"), tmpStr);
	xls.SetCellText(6, nRow, tmpStr, &hf);

	TranslateString(_T("TT, PT"), tmpStr);
	xls.SetCellText(7, nRow, tmpStr, &hf);

	xls.SetCellText(8, nRow, _T("VTYT"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(9, nRow, _T("DVKTC"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(10, nRow, _T("Thu\x1ED1\x63 K"), &hf);

	TranslateString(_T("Examination Fee"), tmpStr);
	xls.SetCellText(11, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(12, nRow, _T("Ph\xED V\x43"), &hf);

	TranslateString(_T("Total Cost"), tmpStr);

	xls.SetCellText(13, nRow - 1, tmpStr, &hf);

	TranslateString(_T("Difference Payment"), tmpStr);
	xls.SetCellText(14, nRow - 1, tmpStr, &hf);

	TranslateString(_T("Insurance Paid"), tmpStr);
	xls.SetCellText(15, nRow - 1, tmpStr, &hf);

	TranslateString(_T("Ngo\xE0i qu\x1EF9"), tmpStr);
	xls.SetCellText(16, nRow - 1, tmpStr, &hf);

	szSQL = GetQueryString();

	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return ;
	}
	long nIndex = 1;
	CString szOldLine, szNewLine, szOldGroup, szNewGroup, szLineName;
	CString szNewOfLine, szOldOfLine,szSumOfline;

	long double grpCost[17], grpLine[17], ttlCost[17], grpOfLine[17];
	double cost = 0;
	nRow++;

	for (int i =0; i < 17; i++)
	{
		tmpStr.Format(_T("(%d)"), i + 1);
		xls.SetCellText(i, nRow, tmpStr, FMT_TEXT | FMT_CENTER, TRUE);
		grpCost[i] = 0;
		grpOfLine[i] = 0;
		grpLine[i] = 0;
		ttlCost[i] = 0;		
	}	

	while (!rs.IsEOF())
	{		
		rs.GetValue(_T("linename"), szNewLine);
		if (szNewLine != szOldLine && !szNewLine.IsEmpty())
		{
			CString szField,szAmount;
			//Tong tuyen (A,B,C)
			if (grpLine[13] > 0)
			{
				nRow++;
				xls.SetCellMergedColumns(1, nRow, 2);
				tmpStr.Format(_T("\x43\x1ED9ng(%s %s=(I+II)): "), szOldLine, szOldGroup);				
				xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true);

				for (int i = 2; i < 17; i++)
				{					
					tmpStr.Format(_T("%.2lf"), grpLine[i]);
					xls.SetCellText(i, nRow, tmpStr, FMT_NUMBER1, true);
					ttlCost[i] += grpLine[i];
					grpLine[i] = 0;
				}
			}

			nRow++;
			xls.SetCellText(0, nRow , szNewLine, FMT_TEXT, true);
			rs.GetValue(_T("insline"), szLineName);
			tmpStr.Format(_T("%s"), pMF->GetSelectionString(_T("hms_insline"), szLineName));
			xls.SetCellMergedColumns(1, nRow, 16);
			xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true);
			szOldLine = szNewLine;
			nIndex = 1;
		}
		
		nRow++;

		rs.GetValue(_T("hdline"), szNewGroup);
		if (szNewGroup == _T("0"))
		{
			tmpStr.Format(_T("%s"), lszLine[0]);
			szSumOfline = _T("I");
		}
		else
		{
			tmpStr.Format(_T("%s"), lszLine[1]);
			szSumOfline = _T("II");
		}
		
		xls.SetCellText(0, nRow, szSumOfline, FMT_TEXT);

		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("total"), tmpStr);
		xls.SetCellText(2, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("testfee"), cost);		
		grpLine[3] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("pacsfee"), cost);		
		grpLine[4] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("drugfee"), cost);		
		grpLine[5] += cost;			
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("bloodfee"), cost);		
		grpLine[6] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("normtechfee"), cost);
		grpLine[7] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("materialfee"), cost);
		grpLine[8] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("hitechfee"), cost);		
		grpLine[9] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(9, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("drugfeek"), cost);
		grpLine[10] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(10, nRow, tmpStr, FMT_NUMBER1);
		

		rs.GetValue(_T("examfee"), cost);		
		grpLine[11] += cost;		
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(11, nRow, tmpStr, FMT_NUMBER1);
		

		/*rs.GetValue(_T("otherfee"), cost);		
		grpLine[12] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(12, nRow, tmpStr, FMT_NUMBER1);	*/

		rs.GetValue(_T("cost"), cost);
		grpLine[13] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(13, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("patpaid"), cost);
		grpLine[14] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(14, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("inspaid"), cost);
		grpLine[15] += cost;
		tmpStr.Format(_T("%.2f"), cost);
		xls.SetCellText(15, nRow, tmpStr, FMT_NUMBER1);
		rs.MoveNext();
	}

	for (int i = 0; i < 17; i++)
	{
		ttlCost[i] += grpLine[i];
	}
	
	if (grpLine[13] > 0)
	{
		nRow++;
		xls.SetCellMergedColumns(1, nRow, 2);
		tmpStr.Format(_T("\x43\x1ED9ng(%s=(I+II)): "), szNewLine);	
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true);
		for (int i = 2; i < 17; i++)
		{			
			tmpStr.Format(_T("%.2lf"), grpLine[i]);
			xls.SetCellText(i, nRow , tmpStr, FMT_NUMBER1, true);	
			grpLine[i] = 0;
		}
	}

	if (ttlCost[13] > 0)
	{
		nRow++;
		CString szField, szAmount;
		TranslateString(_T("Total Amount"), szAmount);
		tmpStr.Format(_T("%s (A+B+C):"), szAmount);
		xls.SetCellMergedColumns(1, nRow, 2);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT, true);
		for (int i = 2; i < 17; i++)
		{			
			tmpStr.Format(_T("%.2lf"), ttlCost[i]);
			xls.SetCellText(i, nRow, tmpStr, FMT_NUMBER1, true);
		}
	}	
	xls.Save(_T("Exports\\ChiPhiKhamBenhBHYT(79aTH-BHYT).xls"));
	EndWaitCursor();
	
} 
int CFMGeneralInsuranceCost79ATH::OnAddFMGeneralInsuranceCost79ATH()
{
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMGeneralInsuranceCost79ATH::OnEditFMGeneralInsuranceCost79ATH(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMGeneralInsuranceCost79ATH::OnDeleteFMGeneralInsuranceCost79ATH(){
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
 		OnCancelFMGeneralInsuranceCost79ATH(); 
 	}
	return 0;
}
int CFMGeneralInsuranceCost79ATH::OnSaveFMGeneralInsuranceCost79ATH(){
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
 		//OnFMGeneralInsuranceCost79ATHListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMGeneralInsuranceCost79ATH::OnCancelFMGeneralInsuranceCost79ATH(){
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
int CFMGeneralInsuranceCost79ATH::OnFMGeneralInsuranceCost79ATHListLoadData(){
	return 0;
}
CString CFMGeneralInsuranceCost79ATH::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	/*CString szSQL, szWhere, szSubWhere;
	CString szDept;

	szWhere.Empty();
	szSubWhere.Empty();

	if (m_bWithoutEmergency)
		szWhere.Format(_T(" AND hd_emergency<>'Y' "));
	if (m_bOnlyEmergency)
		szWhere.Format(_T(" AND hd_emergency='Y' "));

	if (!m_szClerkKey.IsEmpty())
		szSubWhere.Format(_T(" AND fi.hfe_staff='%s' "), m_szClerkKey);

	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (!szDept.IsEmpty())
				szDept += _T(",");
			szDept.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
		}
	}

	if (!szDept.IsEmpty())
		szSubWhere.AppendFormat(_T(" AND fi.hfe_deptid IN(%s) "), szDept);*/

	CString szSQL, szWhere, szSubWhere;
	CString szRooms, szDeptID, szDepts;
	CString szRoomCondition;

	szWhere.Empty();
	szDepts.Empty();
	szRooms.Empty();
	szDeptID.Empty();
	szSubWhere.Empty();

	if (m_bWithoutEmergency)
		szWhere.AppendFormat(_T(" AND hd_emergency<>'Y' "));
	if (m_bOnlyEmergency)
		szWhere.AppendFormat(_T(" AND hd_emergency='Y' "));

	if (!m_szClerkKey.IsEmpty())
		szSubWhere.AppendFormat(_T(" AND fi.hfe_staff='%s' "), m_szClerkKey);

	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			if (!szDepts.IsEmpty())
				szDepts += _T(",");
			szDepts.AppendFormat(_T("'%s'"), m_wndDeptList.GetItemText(i, 0));
		}
	}

	for (int i = 0; i < m_wndRoomList.GetItemCount(); i++)
	{
		if (m_wndRoomList.GetCheck(i))
		{
			if (!szRooms.IsEmpty())
				szRooms += _T(",");

			if (szDeptID.IsEmpty())
				szDeptID = m_wndRoomList.GetItemText(i, 2);

			szRooms.AppendFormat(_T("%s"), m_wndRoomList.GetItemText(i, 0));
		}
	}

	if (!szRooms.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND (select he_roomid from hms_exam where he_docno=hfe_docno and upper(he_deptid)=upper('%s') and he_status='T' ") \
			                 _T("and he_receptidx=(select max(he_receptidx) from hms_exam where he_docno=hfe_docno)) IN(%s) "),
							 szDeptID, szRooms);
	}

	if (!szDepts.IsEmpty() || !szDeptID.IsEmpty())
	{
		if (szDepts.IsEmpty())
			szDepts = _T("''");

		szSubWhere.AppendFormat(_T(" AND (fi.hfe_deptid IN(%s) OR fi.hfe_deptid IN('%s')) "),
			                    szDepts, szDeptID);
	}

	/*szSQL.Format(_T(" SELECT SUM(total) AS total,") \
				_T("        hdline,") \
				_T("        linename,") \
				_T("        insline,") \
				_T("        SUM(drugfee) AS drugfee,") \
				_T("        SUM(drugfeek) AS drugfeek,") \
				_T("        SUM(bloodfee) AS bloodfee,") \
				_T("        SUM(testfee) AS testfee,") \
				_T("        SUM(pacsfee) AS pacsfee,") \
				_T("        SUM(normtechfee) AS normtechfee,") \
				_T("        SUM(hitechfee) AS hitechfee,") \
				_T("        SUM(materialfee) AS materialfee,") \
				_T("        SUM(examfee) AS examfee,") \
				_T("        SUM(cost) AS cost,") \
				_T("        SUM(otherfee) AS otherfee,") \
				_T("        SUM(inspaid) AS inspaid,") \
				_T("        SUM(patpaid) AS patpaid") \
				_T(" FROM") \
				_T(" (") \
				_T("  SELECT hfe_docno,") \
				_T("         1 AS total,") \
				_T("      CASE WHEN hd_insline ='Y' THEN 1 ELSE 0 END AS hdline,") \
				_T("         hc_line AS insline,") \
				_T("         (SELECT distinct ss_default FROM sys_sel WHERE ss_id='hms_insline' AND ss_code=CAST(hc_line AS NVARCHAR2(15)) ) AS linename,") \
				_T("         SUM(drugfee) AS drugfee,") \
				_T("         SUM(drugfeek) AS drugfeek,") \
				_T("         SUM(bloodfee) AS bloodfee,") \
				_T("         SUM(testfee) AS testfee,") \
				_T("         SUM(pacsfee) AS pacsfee,") \
				_T("         SUM(normtechfee) AS normtechfee,") \
				_T("         SUM(hitechfee) AS hitechfee,") \
				_T("         SUM(materialfee) AS materialfee,") \
				_T("         SUM(examfee) AS examfee,") \
				_T("         SUM(cost) AS cost,") \
				_T("         SUM(otherfee) AS otherfee,") \
				_T("         SUM(discount) AS inspaid,") \
				_T("         SUM(patpaid) AS patpaid") \
				_T("  FROM") \
				_T("  (") \
				_T("   SELECT fe.hfe_docno,") \
				_T("          CASE WHEN SUBSTR(hfe_groupid, 1, 2) IN('A1','A3') AND hfe_groupid<>'A1400' THEN fe.hfe_cost ELSE 0 END AS drugfee,") \
				_T("          CASE WHEN hfe_groupid='A1400' THEN fe.hfe_cost ELSE 0 END AS drugfeek,") \
				_T("          CASE WHEN SUBSTR(hfe_groupid, 1, 2)='A2' THEN fe.hfe_cost ELSE 0 END AS bloodfee,") \
				_T("          CASE WHEN SUBSTR(hfe_groupid, 1, 2)='B1' THEN fe.hfe_cost ELSE 0 END AS testfee,") \
				_T("          CASE WHEN SUBSTR(hfe_groupid, 1, 2) IN('B2','B3') AND hfe_hitech='N' THEN fe.hfe_cost ELSE 0 END AS pacsfee,") \
				_T("          CASE WHEN SUBSTR(hfe_groupid, 1, 2) IN('B4','B5') AND hfe_hitech='N' THEN fe.hfe_cost ELSE 0 END AS normtechfee,") \
				_T("          CASE WHEN hfe_hitech='Y' THEN fe.hfe_cost ELSE 0 END AS hitechfee,") \
				_T("          CASE WHEN hfe_groupid='A9000' THEN fe.hfe_cost ELSE 0 END AS materialfee,") \
				_T("          CASE WHEN hfe_groupid='D0000' THEN fe.hfe_cost ELSE 0 END AS examfee,") \
				_T("          CASE WHEN hfe_groupid='F0000' THEN fe.hfe_cost ELSE 0 END AS otherfee,") \
				_T("          fe.hfe_cost AS cost,") \
				_T("          fe.hfe_discount AS discount,") \
				_T("          fe.hfe_patpaid AS patpaid,") \
				_T("          fe.hfe_diffpaid AS difpaid") \
				_T("   FROM hms_fee_invoice fi") \
				_T("   LEFT JOIN hms_fee_invoiceline fe ON(fe.hfe_invoiceno=fi.hfe_invoiceno AND fe.hfe_docno=fi.hfe_docno)") \
				_T("   WHERE fi.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				_T("         AND fi.hfe_class IN('E') AND fe.hfe_discount > 0 %s") \
				_T("  ) tbl") \
				_T("  LEFT JOIN hms_doc ON(hfe_docno=hd_docno)") \
				_T("  LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx)") \
				_T("  LEFT JOIN hms_object ON(ho_id=hd_object)") \
				_T("  WHERE ho_type IN('I','C') AND hd_status='T' AND LENGTH(hc_cardno) > 1 %s") \
				_T("  GROUP BY hfe_docno, hc_line, hd_insline") \
				_T("  HAVING SUM(cost) > 0") \
				_T(" ) tblMain") \
				_T(" GROUP BY insline, linename, hdline") \
				_T(" ORDER BY insline, hdline"),
				m_szFromDate, m_szToDate, szUserReceiver, szWhere);*/
	szSQL.Format(_T(" SELECT SUM(total) AS total,") \
				_T("        hdline,") \
				_T("        linename,") \
				_T("        insline,") \
				_T("        SUM(drugfee) AS drugfee,") \
				_T("        SUM(drugfeek) AS drugfeek,") \
				_T("        SUM(bloodfee) AS bloodfee,") \
				_T("        SUM(testfee) AS testfee,") \
				_T("        SUM(pacsfee) AS pacsfee,") \
				_T("        SUM(normtechfee) AS normtechfee,") \
				_T("        SUM(hitechfee) AS hitechfee,") \
				_T("        SUM(materialfee) AS materialfee,") \
				_T("        SUM(examfee) AS examfee,") \
				_T("        SUM(cost) AS cost,") \
				_T("        SUM(otherfee) AS otherfee,") \
				_T("        SUM(inspaid) AS inspaid,") \
				_T("        SUM(patpaid) AS patpaid") \
				_T(" FROM") \
				_T(" (") \
				_T("  SELECT hfe_docno,") \
				_T("         1 AS total,") \
				_T("         CASE WHEN hd_insline ='Y' THEN 1 ELSE 0 END AS hdline,") \
				_T("         hc_line AS insline,") \
				_T("         (SELECT distinct ss_default FROM sys_sel WHERE ss_id='hms_insline' AND ss_code=CAST(hc_line AS NVARCHAR2(15))) AS linename,") \
				_T("         round(SUM(drugfee)) AS drugfee,") \
				_T("         round(SUM(drugfeek)) AS drugfeek,") \
				_T("         round(SUM(bloodfee)) AS bloodfee,") \
				_T("         round(SUM(testfee)) AS testfee,") \
				_T("         round(SUM(pacsfee)) AS pacsfee,") \
				_T("         round(SUM(normtechfee)) AS normtechfee,") \
				_T("         round(SUM(hitechfee)) AS hitechfee,") \
				_T("         round(SUM(materialfee)) AS materialfee,") \
				_T("         round(SUM(examfee)) AS examfee,") \
				_T("         round(SUM(inspaid)) AS cost,") \
				_T("         round(SUM(otherfee)) AS otherfee,") \
				_T("         round(SUM(discount)) AS inspaid,") \
				_T("         round(SUM(inspaid-discount)) AS patpaid") \
				_T("  FROM") \
				_T("  (") \
				_T("   SELECT fe.hfe_docno,") \
				_T("          CASE WHEN SUBSTR(hfe_group, 1, 2) IN('A1','A6') AND hfe_group<>'A1400' THEN fe.hfe_inspaid ELSE 0 END AS drugfee,") \
				_T("          CASE WHEN hfe_group='A1400' THEN fe.hfe_inspaid ELSE 0 END AS drugfeek,") \
				_T("          CASE WHEN NVL(hfl_bloodfee, 'X') = 'Y' THEN fe.hfe_inspaid ELSE 0 END AS bloodfee,") \
				_T("          CASE WHEN SUBSTR(hfe_group, 1, 2)='B1' AND hfe_hitech = 'N' AND NVL(hfl_bloodfee, 'X') <> 'Y' THEN fe.hfe_inspaid ELSE 0 END AS testfee,") \
				_T("          CASE WHEN SUBSTR(hfe_group, 1, 2) IN('B2','B3') AND hfe_hitech='N' THEN fe.hfe_inspaid ELSE 0 END AS pacsfee,") \
				_T("          CASE WHEN SUBSTR(hfe_group, 1, 2) IN('B4','B5') AND hfe_hitech='N' THEN fe.hfe_inspaid ELSE 0 END AS normtechfee,") \
				_T("          CASE WHEN hfe_hitech='Y' THEN fe.hfe_inspaid ELSE 0 END AS hitechfee,") \
				_T("          CASE WHEN hfe_group='A9000' THEN fe.hfe_inspaid ELSE 0 END AS materialfee,") \
				_T("          CASE WHEN hfe_group='D0000' THEN fe.hfe_inspaid ELSE 0 END AS examfee,") \
				_T("          CASE WHEN hfe_group='F0000' THEN fe.hfe_inspaid ELSE 0 END AS otherfee,") \
				_T("          fe.hfe_cost AS cost,") \
				_T("          fe.hfe_inspaid AS inspaid,") \
				_T("          fe.hfe_discount AS discount,") \
				_T("          fe.hfe_patpaid AS patpaid,") \
				_T("          fe.hfe_diffpaid AS difpaid") \
				_T("   FROM hms_fee_invoice fi") \
				_T("   LEFT JOIN hms_fee_view fe ON(fe.hfe_invoiceno=fi.hfe_invoiceno AND fe.hfe_docno=fi.hfe_docno)") \
				_T("   LEFT JOIN hms_fee_list ON (fe.hfe_itemid = hfl_feeid)") \
				_T("   WHERE fi.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				_T("         AND fi.hfe_class IN('E') AND fe.hfe_discount > 0 AND fi.hfe_status='P' %s") \
				_T("  ) tbl") \
				_T("  LEFT JOIN hms_doc ON(hfe_docno=hd_docno)") \
				_T("  LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx)") \
				_T("  LEFT JOIN hms_object ON(ho_id=hd_object)") \
				_T("  WHERE ho_type IN('I','C') AND hd_status='T' AND LENGTH(hc_cardno) > 1 %s") \
				_T("  GROUP BY hfe_docno, hc_line, hd_insline") \
				_T("  HAVING SUM(cost) > 0") \
				_T(" ) tblMain") \
				_T(" GROUP BY insline, linename, hdline") \
				_T(" ORDER BY insline, hdline"),
				m_szFromDate, m_szToDate,
				szSubWhere, szWhere);
	_fmsg(_T("%s"), szSQL);
	return szSQL;
}

int CFMGeneralInsuranceCost79ATH::OnListCheckAll()
{
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (!m_wndList.GetCheck(i))
		{
			m_wndList.SetCheck(i, TRUE);
		}
	}
	return 0;
}

int CFMGeneralInsuranceCost79ATH::OnListUnCheckAll()
{
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			m_wndList.SetCheck(i, FALSE);
		}
	}
	return 0;
}

int CFMGeneralInsuranceCost79ATH::OnDeptListCheckAll()
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

int CFMGeneralInsuranceCost79ATH::OnDeptListUnCheckAll()
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


int CFMGeneralInsuranceCost79ATH::OnRoomListCheckAll()
{
	for (int i = 0; i < m_wndRoomList.GetItemCount(); i++)
	{
		if (!m_wndRoomList.GetCheck(i))
		{
			m_wndRoomList.SetCheck(i, TRUE);
		}
	}
	return 0;
}

int CFMGeneralInsuranceCost79ATH::OnRoomListUnCheckAll()
{
	for (int i = 0; i < m_wndRoomList.GetItemCount(); i++)
	{
		if (m_wndRoomList.GetCheck(i))
		{
			m_wndRoomList.SetCheck(i, FALSE);
		}
	}
	return 0;
}

