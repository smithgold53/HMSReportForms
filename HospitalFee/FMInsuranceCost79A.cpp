#include "stdafx.h"
#include "FMInsuranceCost79A.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMInsuranceCost79A* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMInsuranceCost79A* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnClerkAddNew();
}*/
static void _OnDeptSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMInsuranceCost79A* )pWnd)->OnDeptSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDeptSelendokFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnDeptSelendok();
}
/*static void _OnDeptSetfocusFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnDeptKillfocus();
}*/
/*static void _OnDeptKillfocusFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnDeptKillfocus();
}*/
static long _OnDeptLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A *)pWnd)->OnDeptLoadData();
}
/*static void _OnDeptAddNewFnc(CWnd *pWnd){
	((CFMInsuranceCost79A *)pWnd)->OnDeptAddNew();
}*/
static void _OnWithoutEmergencySelectFnc(CWnd *pWnd){
	 ((CFMInsuranceCost79A*)pWnd)->OnWithoutEmergencySelect();
}
static void _OnOnlyEmergencySelectFnc(CWnd *pWnd){
	 ((CFMInsuranceCost79A*)pWnd)->OnOnlyEmergencySelect();
}
static void _OnUnapprovedSelectFnc(CWnd *pWnd){
	((CFMInsuranceCost79A*)pWnd)->OnUnapprovedSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMInsuranceCost79A *pVw = (CFMInsuranceCost79A *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMInsuranceCost79A *pVw = (CFMInsuranceCost79A *)pWnd;
	pVw->OnExportSelect();
}
static void _OnExportForInsSelectFnc(CWnd *pWnd){
	CFMInsuranceCost79A *pVw = (CFMInsuranceCost79A *)pWnd;
	pVw->OnExportForInsSelect();
}
//static long _OnListLoadDataFnc(CWnd *pWnd){
//	return ((CFMInsuranceCost79A*)pWnd)->OnListLoadData();
//} 
//static void _OnListDblClickFnc(CWnd *pWnd){
//	((CFMInsuranceCost79A*)pWnd)->OnListDblClick();
//} 
//static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
//	((CFMInsuranceCost79A*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
//} 
//static int _OnListDeleteItemFnc(CWnd *pWnd){
//	 return ((CFMInsuranceCost79A*)pWnd)->OnListDeleteItem();
//}
static long _OnCardListLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A*)pWnd)->OnCardListLoadData();
} 
static void _OnCardListDblClickFnc(CWnd *pWnd){
	((CFMInsuranceCost79A*)pWnd)->OnCardListDblClick();
} 
static void _OnCardListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMInsuranceCost79A*)pWnd)->OnCardListSelectChange(nOldItem, nNewItem);
} 
static int _OnCardListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMInsuranceCost79A*)pWnd)->OnCardListDeleteItem();
} 

static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A*)pWnd)->OnDeptListLoadData();
} 
static void _OnDeptListDblClickFnc(CWnd *pWnd){
	((CFMInsuranceCost79A*)pWnd)->OnDeptListDblClick();
} 
static void _OnDeptListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMInsuranceCost79A*)pWnd)->OnDeptListSelectChange(nOldItem, nNewItem);
} 
static int _OnDeptListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMInsuranceCost79A*)pWnd)->OnDeptListDeleteItem();
}
static long _OnRoomListLoadDataFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A*)pWnd)->OnRoomListLoadData();
} 
static void _OnRoomListDblClickFnc(CWnd *pWnd){
	((CFMInsuranceCost79A*)pWnd)->OnRoomListDblClick();
} 
static void _OnRoomListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMInsuranceCost79A*)pWnd)->OnRoomListSelectChange(nOldItem, nNewItem);
} 
static int _OnRoomListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMInsuranceCost79A*)pWnd)->OnRoomListDeleteItem();
}
static int _OnAddFMInsuranceCost79AFnc(CWnd *pWnd){
	 return ((CFMInsuranceCost79A*)pWnd)->OnAddFMInsuranceCost79A();
} 
static int _OnEditFMInsuranceCost79AFnc(CWnd *pWnd){
	 return ((CFMInsuranceCost79A*)pWnd)->OnEditFMInsuranceCost79A();
} 
static int _OnDeleteFMInsuranceCost79AFnc(CWnd *pWnd){
	 return ((CFMInsuranceCost79A*)pWnd)->OnDeleteFMInsuranceCost79A();
} 
static int _OnSaveFMInsuranceCost79AFnc(CWnd *pWnd){
	 return ((CFMInsuranceCost79A*)pWnd)->OnSaveFMInsuranceCost79A();
} 
static int _OnCancelFMInsuranceCost79AFnc(CWnd *pWnd){
	 return ((CFMInsuranceCost79A*)pWnd)->OnCancelFMInsuranceCost79A();
}
static int _OnRoomListCheckAllFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A*)pWnd)->OnRoomListCheckAll();
}
static int _OnRoomListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A*)pWnd)->OnRoomListUnCheckAll();
}
static int _OnDeptListCheckAllFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A*)pWnd)->OnDeptListCheckAll();
}
static int _OnDeptListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A*)pWnd)->OnDeptListUnCheckAll();
}
static int _OnCardListCheckAllFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A*)pWnd)->OnCardListCheckAll();
}
static int _OnCardListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMInsuranceCost79A*)pWnd)->OnCardListUnCheckAll();
}
CFMInsuranceCost79A::CFMInsuranceCost79A(CWnd *pParent)
{
	m_nDlgWidth = 450;
	m_nDlgHeight = 585;
	SetDefaultValues();
}
CFMInsuranceCost79A::~CFMInsuranceCost79A()
{
}
void CFMInsuranceCost79A::OnCreateComponents()
{
	m_wndReportFilter.Create(this, _T("Report Parameters"), 5, 5, 440, 545);
	m_wndInsCardInfo.Create(this, _T("Insurance Card Information"), 10, 120, 435, 225);
	m_wndDeptInfo.Create(this, _T("Dept"), 10, 230, 435, 340);
	m_wndExamRoomInfo.Create(this, _T("Exam Room"), 10, 345, 435, 450);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 310, 55);
	m_wndReportPeriod.Create(this,315, 30, 435, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 310, 85);
	m_wndToDate.Create(this,315, 60, 435, 85); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 90, 90, 115);
	m_wndClerk.Create(this,95, 90, 215, 115); 
	m_wndNumberOfDocLabel.Create(this, _T("Number of Doc"), 220, 90, 310, 115);
	m_wndNumberOfDoc.Create(this,315, 90, 435, 115); 
	m_wndWithoutEmergency.Create(this, _T("Without Emergency"), 10, 455, 245, 480);
	m_wndOnlyEmergency.Create(this, _T("Only Emergency"), 250, 455, 435, 480);
	m_wndOnlyCommander.Create(this, _T("Only General"), 10, 485, 125, 510);
	m_wndUnapproved.Create(this, _T("Unapproved"), 130, 485, 245, 510);
	m_wndOutpatientTreatment.Create(this, _T("Outpatient Treatment"), 250, 485, 435, 510);
	m_wndByDischargedDate.Create(this, _T("By Discharged Date"), 10, 515, 125, 540);
	m_wndSoldierPatient.Create(this, _T("Soldier Patient"), 130, 515, 245, 540);
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 145, 550, 245, 575);
	m_wndExport.Create(this, _T("&Export"), 250, 550, 335, 575);
	m_wndExportForIns.Create(this, _T("Export for Ins"), 340, 550, 440, 575);
	m_wndCardList.Create(this,15, 145, 430, 220); 
	m_wndDeptList.Create(this,15, 255, 430, 335); 
	m_wndRoomList.Create(this,15, 370, 430, 445); 
}
void CFMInsuranceCost79A::OnInitializeComponents()
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

	//m_wndDept.SetCheckValue(true);
	m_wndDept.LimitText(254);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 270);

	/*m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndList.InsertColumn(1, _T("Name"), CFMT_TEXT, 280);*/
	m_wndDept.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	m_wndDept.InsertColumn(1, _T("Name"), CFMT_TEXT, 280);


	m_wndCardList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndCardList.InsertColumn(1, _T("Desc"), CFMT_TEXT, 300);
	m_wndCardList.SetCheckBox(TRUE);

	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
	m_wndDeptList.SetCheckBox(TRUE);

	m_wndRoomList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndRoomList.InsertColumn(1, _T("Name"), CFMT_TEXT, 220);
	m_wndRoomList.InsertColumn(2, _T("Dept"), CFMT_TEXT, 80);
	m_wndRoomList.SetCheckBox(TRUE);

}
void CFMInsuranceCost79A::OnSetWindowEvents()
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
	m_wndDept.SetEvent(WE_SELENDOK, _OnDeptSelendokFnc);
	//m_wndDept.SetEvent(WE_SETFOCUS, _OnDeptSetfocusFnc);
	//m_wndDept.SetEvent(WE_KILLFOCUS, _OnDeptKillfocusFnc);
	m_wndDept.SetEvent(WE_SELCHANGE, _OnDeptSelectChangeFnc);
	m_wndDept.SetEvent(WE_LOADDATA, _OnDeptLoadDataFnc);
	//m_wndDept.SetEvent(WE_ADDNEW, _OnDeptAddNewFnc);
	m_wndWithoutEmergency.SetEvent(WE_CLICK, _OnWithoutEmergencySelectFnc);
	m_wndOnlyEmergency.SetEvent(WE_CLICK, _OnOnlyEmergencySelectFnc);
	m_wndUnapproved.SetEvent(WE_CLICK, _OnUnapprovedSelectFnc);

	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndExportForIns.SetEvent(WE_CLICK, _OnExportForInsSelectFnc);
	m_wndCardList.SetEvent(WE_SELCHANGE, _OnCardListSelectChangeFnc);
	m_wndCardList.SetEvent(WE_LOADDATA, _OnCardListLoadDataFnc);
	m_wndCardList.SetEvent(WE_DBLCLICK, _OnCardListDblClickFnc);
	m_wndCardList.SetWindowText(_T(""));
	m_wndCardList.AddEvent(1, _T("Check All"), _OnCardListCheckAllFnc);
	m_wndCardList.AddEvent(2, _T("UnCheck All"), _OnCardListUnCheckAllFnc);

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
	OnCardListLoadData();
}
void CFMInsuranceCost79A::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Text(pDX, m_wndNumberOfDoc.GetDlgCtrlID(), m_nNumberofDoc);
	//DDX_TextEx(pDX, m_wndDept.GetDlgCtrlID(), m_szDeptKey);
	DDX_Check(pDX, m_wndWithoutEmergency.GetDlgCtrlID(), m_bWithoutEmergency);
	DDX_Check(pDX, m_wndOnlyEmergency.GetDlgCtrlID(), m_bOnlyEmergency);
	DDX_Check(pDX, m_wndUnapproved.GetDlgCtrlID(), m_bUnapproved);
	DDX_Check(pDX, m_wndOnlyCommander.GetDlgCtrlID(), m_bOnlyCommander);
	DDX_Check(pDX, m_wndOutpatientTreatment.GetDlgCtrlID(), m_bOutpatientTreament);
	DDX_Check(pDX, m_wndByDischargedDate.GetDlgCtrlID(), m_bByDischargedDate);
	DDX_Check(pDX, m_wndSoldierPatient.GetDlgCtrlID(), m_bSoldierPatient);

}
void CFMInsuranceCost79A::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMInsuranceCost79A::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMInsuranceCost79A::SetDefaultValues()
{
	m_nYear = 0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_szDeptKey.Empty();
	m_bWithoutEmergency=FALSE;
	m_bOnlyEmergency=FALSE;
	m_bUnapproved=FALSE;
	m_bIsInsPurpose = FALSE;
	m_bByDischargedDate = FALSE;
	m_bShowDiff = FALSE;
	m_bSoldierPatient = FALSE;
}
int CFMInsuranceCost79A::SetMode(int nMode)
{
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
/*void CFMInsuranceCost79A::OnYearChange(){
	
} */
/*void CFMInsuranceCost79A::OnYearSetfocus(){
	
} */
/*void CFMInsuranceCost79A::OnYearKillfocus(){
	
} */
int CFMInsuranceCost79A::OnYearCheckValue()
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
void CFMInsuranceCost79A::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMInsuranceCost79A::OnReportPeriodSelendok()
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
/*void CFMInsuranceCost79A::OnReportPeriodSetfocus(){
	
}*/
/*void CFMInsuranceCost79A::OnReportPeriodKillfocus(){
	
}*/
long CFMInsuranceCost79A::OnReportPeriodLoadData()
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
/*void CFMInsuranceCost79A::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CFMInsuranceCost79A::OnFromDateChange(){
	
} */
/*void CFMInsuranceCost79A::OnFromDateSetfocus(){
	
} */
/*void CFMInsuranceCost79A::OnFromDateKillfocus(){
	
} */
int CFMInsuranceCost79A::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMInsuranceCost79A::OnToDateChange(){
	
} */
/*void CFMInsuranceCost79A::OnToDateSetfocus(){
	
} */
/*void CFMInsuranceCost79A::OnToDateKillfocus(){
	
} */
int CFMInsuranceCost79A::OnToDateCheckValue(){
	return 0;
} 
void CFMInsuranceCost79A::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMInsuranceCost79A::OnClerkSelendok(){
	 
}
/*void CFMInsuranceCost79A::OnClerkSetfocus(){
	
}*/
/*void CFMInsuranceCost79A::OnClerkKillfocus(){
	
}*/
long CFMInsuranceCost79A::OnClerkLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	}
	m_wndClerk.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as id, su_name as name FROM sys_user WHERE su_groupid in('A','F') %s ORDER BY su_userid"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){
		m_wndClerk.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMInsuranceCost79A::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMInsuranceCost79A::OnDeptSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMInsuranceCost79A::OnDeptSelendok()
{
	UpdateData(TRUE);
	OnRoomListLoadData();
}
/*void CFMInsuranceCost79A::OnDeptSetfocus(){
	
}*/
/*void CFMInsuranceCost79A::OnDeptKillfocus(){
	
}*/
long CFMInsuranceCost79A::OnDeptLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if (m_wndDept.IsSearchKey() && !m_szDeptKey.IsEmpty())
	{
	  szWhere.Format(_T(" and sd_id='%s' "), m_szDeptKey);
	}
	m_wndDept.DeleteAllItems(); 
	int nCount = 0;

	szSQL.Format(_T("SELECT sd_id as id, sd_name as name ") \
		         _T("FROM sys_dept ") \
				 _T("WHERE 1=1 AND sd_type='KB' %s ORDER BY id "), szWhere);

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDept.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMInsuranceCost79A::OnDeptAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMInsuranceCost79A::OnWithoutEmergencySelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMInsuranceCost79A::OnOnlyEmergencySelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}

void CFMInsuranceCost79A::OnUnapprovedSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	if (m_bUnapproved)
		m_wndClerk.EnableWindow(FALSE);
	else
		m_wndClerk.EnableWindow(TRUE);
}

void CFMInsuranceCost79A::OnCardListDblClick(){
	
} 
void CFMInsuranceCost79A::OnCardListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMInsuranceCost79A::OnCardListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMInsuranceCost79A::OnCardListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	CString tmpStr, szTemp;

	m_wndCardList.BeginLoad(); 
	int nCount = 0;

	szSQL.Format(_T(" select hig_note as itemid, hig_name as itemname") \
				_T(" from hms_insurance_group") \
				_T(" order by hig_group, hig_idx"));

	nCount = rs.ExecSQL(szSQL);

	while (!rs.IsEOF())
	{ 
		m_wndCardList.AddItems(
			rs.GetValue(_T("itemid")), 
			rs.GetValue(_T("itemname")), NULL);
		rs.MoveNext();
	}

	m_wndCardList.EndLoad(); 
	return nCount;
}

void CFMInsuranceCost79A::OnDeptListDblClick(){
	
} 
void CFMInsuranceCost79A::OnDeptListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMInsuranceCost79A::OnDeptListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMInsuranceCost79A::OnDeptListLoadData()
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

void CFMInsuranceCost79A::OnRoomListDblClick(){
	
} 
void CFMInsuranceCost79A::OnRoomListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMInsuranceCost79A::OnRoomListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMInsuranceCost79A::OnRoomListLoadData()
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

void CFMInsuranceCost79A::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;

	TCHAR *lszLine[] ={_T("a) \x42\x1EC7nh nh\xE2n \x111\xFAng tuy\x1EBFn."),
		               _T("b) \x42\x1EC7nh nh\xE2n tr\xE1i tuy\x1EBFn.")};
	
	szSQL = GetQueryString();

	//MessageBox(szSQL);

	BeginWaitCursor();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		if (nCount < 0)
			_msg(_T("%s"), szSQL);
		else
			_fmsg(_T("%s"), szSQL);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/HF_GENERALEXAMINATIONCOST79AREPORT.RPT")))
		return;
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);

	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);

	rpt.GetReportHeader()->SetValue(_T("ObjectGroup"), _T(""));
	rpt.GetReportHeader()->SetValue(_T("ReportPeriod"), m_szClerkKey);

	rs.GetValue(_T("invoiceno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	int nIndex = 1;

	CString szOldLine, szNewLine, szOldGroup, szNewGroup, szLineName;
	CString szNewOfLine, szOldOfLine, szSumOfline;

	CReportSection* rptDetail;
	long double grpCost[24], grpLine[24], ttlCost[24], grpOfLine[24];
	double cost = 0;

	for (int i = 0; i < 24; i++)
	{
		grpCost[i] = 0;
		grpLine[i] = 0;
		ttlCost[i] = 0;
		grpOfLine[i] =0;
	}

	while(!rs.IsEOF())
	{
		rs.GetValue(_T("insline"), szNewLine);

		if (szNewLine != szOldLine && !szNewLine.IsEmpty())
		{
			CString szField, szAmount;

			/*if (grpOfLine[20] > 0)
			{
				TranslateString(_T("\x43\x1ED9ng"), szAmount);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(9);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				tmpStr.Format(_T("%s (%s) "), szAmount, szSumOfline);
				rptDetail->SetValue(_T("TotalGroup"), tmpStr);
				for (int i = 10; i < 24; i++)
				{
					szField.Format(_T("S%d"), i);
					FormatCurrency(grpOfLine[i], tmpStr);
					rptDetail->SetValue(szField, tmpStr);					
					grpOfLine[i] = 0;
				}
			}*/

			if (grpCost[9] > 0)
			{					
				TranslateString(_T("Total Group"), szAmount);

				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(11);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				tmpStr.Format(_T("%s (%s)"), szAmount, szOldGroup);

				rptDetail->SetValue(_T("TotalGroup"), tmpStr);

				for (int i = 9; i < 24; i++)
				{
					szField.Format(_T("S%d"), i);
					FormatCurrency(grpCost[i], tmpStr);
					//_msg(_T("%s"), tmpStr);
					rptDetail->SetValue(szField, tmpStr);
					//grpLine[i] = 0;
					grpCost[i] = 0;
				}
			}

			if (grpLine[9] > 0)
			{
				TranslateString(_T("Total Line"), szAmount);

				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(10);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(false);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				tmpStr.Format(_T("%s (%s)"), szAmount, szLineName);
				rptDetail->SetValue(_T("TotalGroup"), tmpStr);
				for (int i = 9; i < 24; i++)
				{
					szField.Format(_T("S%d"), i);
					FormatCurrency(grpLine[i], tmpStr);
					rptDetail->SetValue(szField, tmpStr);		
					ttlCost[i] += grpLine[i];
					grpLine[i] = 0;
				}
			}

			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetFaceSize(11);
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetBold(true);
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetItalic(false);
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));

			rs.GetValue(_T("linename"), szLineName);
			rptDetail->SetValue(_T("InsuranceLine"), szLineName + _T(". ") + pMF->GetSelectionString(_T("hms_insline"), szNewLine));
			szOldLine = szNewLine;
			szOldOfLine.Empty();
			szOldGroup.Empty();

			nIndex = 1;
		}

		rs.GetValue(_T("insgroup"), szNewGroup);

		if (szNewGroup != szOldGroup && !szNewGroup.IsEmpty())
		{
			CString szField, szAmount;
			/*if (grpOfLine[8] > 0)
			{
				TranslateString(_T("\x43\x1ED9ng"), szAmount);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(9);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				tmpStr.Format(_T("%s (%s) "),szAmount , szSumOfline);
				rptDetail->SetValue(_T("TotalGroup"), tmpStr );
				for (int i =10; i < 24; i++)
				{
					szField.Format(_T("S%d"), i);
					FormatCurrency(grpOfLine[i], tmpStr);
					rptDetail->SetValue(szField, tmpStr);					
					grpOfLine[i] = 0;
				}
			}*/
			if (grpCost[9] > 0)
			{
				TranslateString(_T("Total Group"), szAmount);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(10);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));

				tmpStr.Format(_T("%s (%s)"), szAmount, szOldGroup);

				rptDetail->SetValue(_T("TotalGroup"), tmpStr);
				for (int i = 9; i < 24; i++)
				{
					szField.Format(_T("S%d"), i);
					FormatCurrency(grpCost[i], tmpStr);
					rptDetail->SetValue(szField, tmpStr);				
					grpCost[i] = 0;
				}
			}
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetFaceSize(10);
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetItalic(true);
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetBold(true);			
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rptDetail->SetValue(_T("InsuranceLine"), szNewGroup + _T(". ") + pMF->GetSelectionString(_T("hms_insurance_group"), szNewGroup));
			szOldGroup = szNewGroup;
			szOldOfLine.Empty();
			nIndex = 1;
		}

		/*rs.GetValue(_T("hdline"), szNewOfLine);

		if(szNewOfLine != szOldOfLine && !szNewOfLine.IsEmpty())
		{
			CString szField, szAmount;
			if (grpOfLine[9] > 0)
			{
				TranslateString(_T("\x43\x1ED9ng"), szAmount);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(9);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
				rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				tmpStr.Format(_T("%s (%s) "),szAmount , szSumOfline);
				rptDetail->SetValue(_T("TotalGroup"), tmpStr );
				for (int i = 9; i < 24; i++)
				{
					szField.Format(_T("S%d"), i);
					FormatCurrency(grpOfLine[i], tmpStr);
					rptDetail->SetValue(szField, tmpStr);					
					grpOfLine[i] = 0;
				}
			}
			
			if (szNewOfLine == _T("0") )
			{
				tmpStr.Format(_T("%s"), lszLine[0]);
				szSumOfline = _T("a");
			}
			else
			{
				tmpStr.Format(_T("%s"), lszLine[1]);
				szSumOfline = _T("b");
			}
			
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetFaceSize(9);
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetItalic(true);
			rpt.GetGroupHeader(1)->GetItem(_T("InsuranceLine"))->SetBold(true);			
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rptDetail->SetValue(_T("InsuranceLine"),tmpStr);
			szOldOfLine = szNewOfLine;			
			nIndex = 1;
		}*/

		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("mbirthyear"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("fbirthyear"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		/*rs.GetValue(_T("sex"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);*/

		rs.GetValue(_T("cardno"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("regplace"), tmpStr);
		tmpStr.Insert(2, _T("-"));
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("icd10"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("examdate"), tmpStr);
		//tmpStr = CDate::Convert(rs.GetValue(_T("examdate")), yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("8"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

		/*rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);*/

		rs.GetValue(_T("cost"), cost);
		grpCost[9] += cost;
		grpLine[9] += cost;
		//grpOfLine[9] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);
		
		rs.GetValue(_T("testfee"), cost);
		grpCost[10] += cost;
		grpLine[10] += cost;
		//grpOfLine[10] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.GetValue(_T("pacsfee"), cost);
		grpCost[11] += cost;
		grpLine[11] += cost;
		//grpOfLine[11] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.GetValue(_T("drugfee"), cost);
		grpCost[12] += cost;
		grpLine[12] += cost;
		//grpOfLine[12] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);

		rs.GetValue(_T("bloodfee"), cost);
		grpCost[13] += cost;
		grpLine[13] += cost;
		//grpOfLine[13] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);				

		rs.GetValue(_T("normtechfee"), cost);
		grpCost[14] += cost;
		grpLine[14] += cost;
		//grpOfLine[14] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("14"), tmpStr);
		
		rs.GetValue(_T("materialfee"), cost);
		grpCost[15] += cost;
		grpLine[15] += cost;
		//grpOfLine[15] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("15"), tmpStr);

		rs.GetValue(_T("hitechfee"), cost);
		grpCost[17] += cost;
		grpLine[17] += cost;
		//grpOfLine[17] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("17"), tmpStr);		

		rs.GetValue(_T("drugfeek"), cost);
		grpCost[18] += cost;
		grpLine[18] += cost;
		//grpOfLine[18] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("18"), tmpStr);
		
		rs.GetValue(_T("examfee"), cost);
		grpCost[19] += cost;
		grpLine[19] += cost;
		//grpOfLine[19] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("19"), tmpStr);

		/*rs.GetValue(_T("otherfee"), cost);;
		grpCost[19] += cost;
		grpLine[19] += cost;
		grpOfLine[19] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("19"), tmpStr);*/		
		

		rs.GetValue(_T("patpaid"), cost);
		grpCost[21] += cost;
		grpLine[21] += cost;
		//grpOfLine[21] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("21"), tmpStr);

		rs.GetValue(_T("inspaid"), cost);
		grpCost[22] += cost;
		grpLine[22] += cost;
		//grpOfLine[22] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("22"), tmpStr);
		
		rs.MoveNext();
	}

	for (int i = 9; i < 24; i++)
	{
		ttlCost[i] += grpLine[i];
	}
	
	/*if (grpOfLine[20] > 0)
	{
		CString szField,szAmount;
		TranslateString(_T("\x43\x1ED9ng"), szAmount);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(9);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		tmpStr.Format(_T("%s (%s)"),szAmount , szSumOfline);
		rptDetail->SetValue(_T("TotalGroup"), tmpStr );
		for (int i =10; i < 24; i++)
		{
			szField.Format(_T("S%d"), i);
			FormatCurrency(grpLine[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);
		}	
	}*/

	if (grpCost[9] > 0)
	{
		CString szField,szAmount;
		TranslateString(_T("Total Group"), szAmount);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(10);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(true);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		tmpStr.Format(_T("%s (%s)"), szAmount, szOldGroup);
		rptDetail->SetValue(_T("TotalGroup"), tmpStr);
		for (int i = 9; i < 24; i++)
		{
			szField.Format(_T("S%d"), i);
			FormatCurrency(grpCost[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);
		}
	}

	if (grpLine[9] > 0)
	{		
		CString szField, szAmount;
		TranslateString(_T("Total Line"), szAmount);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(11);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(false);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		tmpStr.Format(_T("%s (%s)"), szAmount, szLineName);
		rptDetail->SetValue(_T("TotalGroup"), tmpStr);
		for (int i = 9; i < 24; i++)
		{
			szField.Format(_T("S%d"), i);
			FormatCurrency(grpCost[i], tmpStr);
			rptDetail->SetValue(szField, tmpStr);
		}
	
	}

	if (ttlCost[9] > 0)
	{
		CString szField, szAmount;

		TranslateString(_T("Total Amount:"), szAmount);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetFaceSize(12);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetBold(true);
		rpt.GetGroupFooter(1)->GetItem(_T("TotalGroup"))->SetItalic(false);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		for (int i = 9; i < 24; i++)
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
	EndWaitCursor();	
	rpt.PrintPreview();
	
} 
void CFMInsuranceCost79A::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	CString tmpStr, szTemp, szFromDateToDate;
	CString szDate;
	CString szMonth, szYear;

	BeginWaitCursor();
	UpdateData(TRUE);
	szSQL = GetQueryString();

	_fmsg(_T("%s"), szSQL);
	int nCount = rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		if (nCount < 0)
			_msg(_T("%s"), szSQL);
		else
			_fmsg(_T("%s"), szSQL);
		return;
	}

	CDate dt1, dt2;
	dt1.ParseDate(m_szFromDate);
	dt2.ParseDate(m_szToDate);

	if (dt1.GetYear() == dt2.GetYear())
		szYear.Format(_T("%d"), dt1.GetYear());
	else
		szYear.Format(_T("%d"), dt2.GetYear());

	if (dt1.GetMonth() == dt2.GetMonth())
		szMonth.Format(_T("%d"), dt1.GetMonth());
	else
		szMonth.Format(_T("%d"), dt2.GetMonth());

	CExcel xls;

	xls.CreateSheet(1);
	xls.SetWorksheet(0);	

	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 17);
	xls.SetColumnWidth(2, 7);
	xls.SetColumnWidth(3, 5);
	xls.SetColumnWidth(4, 16);
	xls.SetColumnWidth(5, 8);

	xls.SetColumnWidth(6, 24);

	xls.SetColumnWidth(7, 10);
	xls.SetColumnWidth(8, 10);

	xls.SetColumnWidth(9, 6);

	xls.SetColumnWidth(10, 8);
	xls.SetColumnWidth(11, 8);
	xls.SetColumnWidth(12, 8);
	xls.SetColumnWidth(13, 8);
	xls.SetColumnWidth(14, 8);
	xls.SetColumnWidth(15, 8);
	xls.SetColumnWidth(16, 8);
	xls.SetColumnWidth(17, 8);
	xls.SetColumnWidth(18, 8);
	xls.SetColumnWidth(19, 8);
	xls.SetColumnWidth(20, 8);
	xls.SetColumnWidth(21, 8);
	xls.SetColumnWidth(22, 8);
	xls.SetColumnWidth(23, 8);
	xls.SetColumnWidth(24, 8);

	xls.SetColumnWidth(25, 6);

	xls.SetColumnWidth(26, 9);

	xls.SetColumnWidth(27, 6);
	xls.SetColumnWidth(28, 5);

	xls.SetColumnWidth(29, 6);
	xls.SetColumnWidth(30, 7);
	xls.SetColumnWidth(31, 10);
	xls.SetColumnWidth(32, 10);
	xls.SetColumnWidth(33, 44);

	xls.SetColumnWidth(34, 8);
	xls.SetColumnWidth(35, 9);
	xls.SetColumnWidth(36, 6);
	xls.SetColumnWidth(37, 8);
	xls.SetColumnWidth(38, 8);
	xls.SetColumnWidth(39, 6);
	xls.SetColumnWidth(40, 8);
	xls.SetColumnWidth(41, 8);
	xls.SetColumnWidth(42, 7);
	xls.SetColumnWidth(43, 5);
	xls.SetColumnWidth(46, 8);
	xls.SetColumnWidth(47, 8);
	xls.SetColumnWidth(48, 8);
	xls.SetCellMergedColumns(0, 1, 5);
	xls.SetCellMergedColumns(0, 2, 5);

	xls.SetCellMergedColumns(37, 1, 7);
	xls.SetCellMergedColumns(37, 2, 7);

	xls.SetCellMergedColumns(0, 3, 44);
	xls.SetCellMergedColumns(0, 4, 44);

	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(37, 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(37, 2, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(0, 3, _T("\x44\x41NH S\xC1\x43H \x110\x1EC0 NGH\x1ECA TH\x41NH TO\xC1N \x43HI PH\xCD KH\xC1M \x43H\x1EEE\x41 \x42\x1EC6NH NGO\x1EA0I TR\xDA"), FMT_TEXT|FMT_CENTER, true, 18);
	
	TranslateString(_T("From Date"), tmpStr);
	szFromDateToDate.Format(_T("%s: %s"), tmpStr,
		                    CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	TranslateString(_T("To Date"), tmpStr);
	szFromDateToDate.AppendFormat(_T(" %s: %s"), tmpStr,
		                          CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));	
	xls.SetCellText(0, 4, szFromDateToDate, FMT_TEXT | FMT_CENTER, true, 12);

	int nRow = 6, nCol = 0;

	xls.SetCellText(0, nRow, _T("stt"), FMT_TEXT);
	xls.SetCellText(1, nRow, _T("hoten"), FMT_TEXT);	
	xls.SetCellText(2, nRow, _T("namsinh"), FMT_TEXT);
	xls.SetCellText(3, nRow, _T("gioitinh"), FMT_TEXT);
	xls.SetCellText(4, nRow, _T("mathe"), FMT_TEXT);

	xls.SetCellText(5, nRow, _T("ma_dkbd"), FMT_TEXT);
	
	xls.SetCellText(6, nRow, _T("mabenh"), FMT_TEXT);

	xls.SetCellText(7, nRow, _T("ngay_vao"), FMT_TEXT);

	xls.SetCellText(8, nRow, _T("ngay_ra"), FMT_TEXT);
	xls.SetCellText(9, nRow, _T("ngaydtr"), FMT_TEXT);
	
	xls.SetCellText(10, nRow, _T("t_tongchi"), FMT_TEXT);
	xls.SetCellText(11, nRow, _T("t_xn"), FMT_TEXT);	
	xls.SetCellText(12, nRow, _T("t_cdha"), FMT_TEXT);
	xls.SetCellText(13, nRow, _T("t_thuoc"), FMT_TEXT);
	xls.SetCellText(14, nRow, _T("t_mau"), FMT_TEXT);	
	xls.SetCellText(15, nRow, _T("t_pttt"), FMT_TEXT);
	xls.SetCellText(16, nRow, _T("t_vtytth"), FMT_TEXT);
	xls.SetCellText(17, nRow, _T("t_vtyttt"), FMT_TEXT);

	xls.SetCellText(18, nRow, _T("t_dvktc"), FMT_TEXT);	
	xls.SetCellText(19, nRow, _T("t_ktg"), FMT_TEXT);
	xls.SetCellText(20, nRow, _T("t_kham"), FMT_TEXT);
	xls.SetCellText(21, nRow, _T("t_vchuyen"), FMT_TEXT);	
	
	xls.SetCellText(22, nRow, _T("t_bnct"), FMT_TEXT);	
	xls.SetCellText(23, nRow, _T("t_bhtt"), FMT_TEXT);
	xls.SetCellText(24, nRow, _T("t_ngoaids"), FMT_TEXT);

	xls.SetCellText(25, nRow, _T("lydo_vv"), FMT_TEXT);
	xls.SetCellText(26, nRow, _T("benhkhac"), FMT_TEXT);
	xls.SetCellText(27, nRow, _T("noikcb"), FMT_TEXT);
	xls.SetCellText(28, nRow, _T("khoa"), FMT_TEXT);
	xls.SetCellText(29, nRow, _T("nam_qt"), FMT_TEXT);
	xls.SetCellText(30, nRow, _T("thang_qt"), FMT_TEXT);
	

	xls.SetCellText(31, nRow, _T("gt_tu"), FMT_TEXT);
	xls.SetCellText(32, nRow, _T("gt_den"), FMT_TEXT);
	xls.SetCellText(33, nRow, _T("diachi"), FMT_TEXT);

	xls.SetCellText(34, nRow, _T("giamdinh"), FMT_TEXT);
	xls.SetCellText(35, nRow, _T("t_xuattoan"), FMT_TEXT);

	xls.SetCellText(36, nRow, _T("lydo_xt"), FMT_TEXT);
	xls.SetCellText(37, nRow, _T("t_datuyen"), FMT_TEXT);
	xls.SetCellText(38, nRow, _T("t_vuottran"), FMT_TEXT);

	xls.SetCellText(39, nRow, _T("loaikcb"), FMT_TEXT);
	xls.SetCellText(40, nRow, _T("noi_ttoan"), FMT_TEXT);
	xls.SetCellText(41, nRow, _T("sophieu"), FMT_TEXT);
	xls.SetCellText(42, nRow, _T("maicd10"), FMT_TEXT);
	xls.SetCellText(43, nRow, _T("tuyen"), FMT_TEXT);
	xls.SetCellText(44, nRow, _T("ngay_tt"), FMT_TEXT);
	xls.SetCellText(45, nRow, _T("phong_kham"), FMT_TEXT);
	nCol = 46;
	if (m_bOnlyCommander)
	{
		xls.SetCellText(nCol, nRow, _T("cap_bac"), FMT_TEXT);
		nCol++;
		xls.SetCellText(nCol, nRow, _T("chuc_vu"), FMT_TEXT);
		nCol++;
	}
	if (m_bIsInsPurpose)
	{
		xls.SetCellText(nCol, nRow, _T("nhom"), FMT_TEXT);
		nCol++;
	}
	if (m_bShowDiff)
	{
		xls.SetCellText(nCol, nRow, _T("chenh"), FMT_TEXT);
		nCol++;
	}
	if (m_bSoldierPatient)
	{
		xls.SetCellText(nCol, nRow, _T("ky_hieu"), FMT_TEXT);
		nCol++;
		xls.SetCellText(nCol, nRow, _T("noi_lam_viec"), FMT_TEXT);
	}
	int nIndex = 1, nTemp;
	double nCost = 0;
	long double nTotal[14];
	CellFormat xlsFormat(&xls);
	xlsFormat.SetCellStyle(FMT_INTEGER);

	CString szNewLine, szOldLine;
	CString szNewGroup, szOldGroup;
	CDateTime dt;
	CDate date;
	CString str;
	nCol = 0;//reinit for use
	int nIncrement = 0;
	for (int i = 0; i < 14; i++)
		nTotal[i] = 0;
	while (!rs.IsEOF())
	{
		/*if (nRow == 50)
			break;*/

		if (!m_bIsInsPurpose)
			rs.GetValue(_T("insline"), szNewLine);

		if (szOldLine != szNewLine && !szNewLine.IsEmpty())
		{
			nRow++;
			xls.SetCellMergedColumns(nCol + 1, nRow, 9);
			rs.GetValue(_T("linename"), tmpStr);
			xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT, true);
			szTemp.Format(_T("%s"), pMF->GetSelectionString(_T("hms_insline"), szNewLine));
			xls.SetCellText(nCol + 1, nRow, szTemp, FMT_TEXT, true);
			szOldLine = szNewLine;
			szOldGroup.Empty();
		}
		if (!m_bIsInsPurpose)
			rs.GetValue(_T("insgroup"), szNewGroup);

		if (szOldGroup != szNewGroup && !szNewGroup.IsEmpty())
		{
			nRow++;
			xls.SetCellText(nCol, nRow, szNewGroup, FMT_TEXT, true);
			xls.SetCellMergedColumns(nCol + 1, nRow, 9);
			tmpStr.Format(_T("%s"), pMF->GetSelectionString(_T("hms_insurance_group"), szNewGroup));
			xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT, true);
			szOldGroup = szNewGroup;
			nIndex = 1;
		}

		nRow++;
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("sex"), tmpStr);
		if (tmpStr == _T("F"))
		{
			rs.GetValue(_T("fbirthyear"), tmpStr);
			xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER);

			xls.WriteNumber(nRow, nCol + 3, 2, &xlsFormat);
		}
		else
		{
			rs.GetValue(_T("mbirthyear"), tmpStr);
			xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER);

			xls.WriteNumber(nRow, nCol + 3, 1, &xlsFormat);
		}
		
		rs.GetValue(_T("cardno"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("regplace"), tmpStr);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("icd10"), tmpStr);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_TEXT);

		tmpStr.Empty();
		rs.GetValue(_T("examdate"), tmpStr);
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_TEXT);

		tmpStr.Empty();
		//szDate.Empty();
		rs.GetValue(_T("enddate"), tmpStr);
		if (!tmpStr.IsEmpty())
		{
			//_msg(_T("%s - %s"), str, rs.GetValue(_T("docno")));
			//szDate = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
			//MessageBox(szDate);
			xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_TEXT);
		}

		tmpStr.Empty();
		rs.GetValue(_T("totaltreat"), tmpStr);
		if (!tmpStr.IsEmpty())
			xls.SetCellText(nCol + 9, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("cost"), nCost);
		nTotal[0] += nCost;
		xls.WriteNumber(nRow, nCol + 10, nCost, &xlsFormat);
		
		rs.GetValue(_T("testfee"), nCost);
		nTotal[1] += nCost;
		xls.WriteNumber(nRow, nCol + 11, nCost, &xlsFormat);

		rs.GetValue(_T("pacsfee"), nCost);
		nTotal[2] += nCost;
		xls.WriteNumber(nRow, nCol + 12, nCost, &xlsFormat);

		rs.GetValue(_T("drugfee"), nCost);
		nTotal[3] += nCost;
		xls.WriteNumber(nRow, nCol + 13, nCost, &xlsFormat);

		rs.GetValue(_T("bloodfee"), nCost);
		nTotal[4] += nCost;
		xls.WriteNumber(nRow, nCol + 14, nCost, &xlsFormat);

		rs.GetValue(_T("normtechfee"), nCost);
		nTotal[5] += nCost;
		xls.WriteNumber(nRow, nCol + 15, nCost, &xlsFormat);

		rs.GetValue(_T("materialfee"), nCost);
		nTotal[6] += nCost;
		xls.WriteNumber(nRow, nCol + 16, nCost, &xlsFormat);

		rs.GetValue(_T("hitechfee"), nCost);
		nTotal[8] += nCost;
		xls.WriteNumber(nRow, nCol + 18, nCost, &xlsFormat);

		rs.GetValue(_T("drugfeek"), nCost);
		nTotal[9] += nCost;
		xls.WriteNumber(nRow, nCol + 19, nCost, &xlsFormat);

		rs.GetValue(_T("examfee"), nCost);
		nTotal[10] += nCost;
		xls.WriteNumber(nRow, nCol + 20, nCost, &xlsFormat);

		//rs.GetValue(_T("transportfee"), nCost);
		//xls.WriteNumber(nRow, nCol + 19, nCost, &xlsFormat);

		rs.GetValue(_T("patpaid"), nCost);
		nTotal[12] += nCost;
		xls.WriteNumber(nRow, nCol + 22, nCost, &xlsFormat);

		rs.GetValue(_T("inspaid"), nCost);
		nTotal[13] += nCost;
		xls.WriteNumber(nRow, nCol + 23, nCost, &xlsFormat);

		//rs.GetValue(_T("ratefee"), nCost);
		//xls.WriteNumber(nRow, nCol + 23, nCost, &xlsFormat);

		rs.GetValue(_T("emergency"), tmpStr);
		if (tmpStr == _T("Y"))
			nTemp = 2;
		else
		{
			rs.GetValue(_T("hdline"), tmpStr);
			if (tmpStr == _T("1"))
			{
				nTemp = 0;
			}
			else
			{
				nTemp = 1;
			}
		}
		
		xls.WriteNumber(nRow, nCol + 25, nTemp, &xlsFormat);

		rs.GetValue(_T("reldisease"), tmpStr);
		xls.SetCellText(nCol + 26, nRow, tmpStr, FMT_TEXT);

		tmpStr.Format(_T("%s"), pMF->m_CompanyInfo.sc_id);
		xls.SetCellText(nCol + 27, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("deptid"), tmpStr);
		xls.SetCellText(nCol + 28, nRow, tmpStr, FMT_TEXT);

		nCost = ToDouble(szYear);
		xls.WriteNumber(nRow, nCol + 29, nCost, &xlsFormat);

		nCost = ToDouble(szMonth);
		xls.WriteNumber(nRow, nCol + 30, nCost, &xlsFormat);
		
		tmpStr.Empty();
		szDate.Empty();
		rs.GetValue(_T("regdate"), tmpStr);
		szDate = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		xls.SetCellText(nCol + 31, nRow, szDate, FMT_TEXT);


		tmpStr.Empty();
		szDate.Empty();
		rs.GetValue(_T("expdate"), tmpStr);
		szDate = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		xls.SetCellText(nCol + 32, nRow, szDate, FMT_TEXT);

		rs.GetValue(_T("address"), tmpStr);
		xls.SetCellText(nCol + 33, nRow, tmpStr, FMT_TEXT);

		xls.WriteNumber(nRow, nCol + 35, 0, &xlsFormat);

		xls.WriteNumber(nRow, nCol + 37, 0, &xlsFormat);
		xls.WriteNumber(nRow, nCol + 38, 0, &xlsFormat);

		xls.SetCellText(nCol + 39, nRow, _T("NGO\x41I"), FMT_TEXT);
		xls.SetCellText(nCol + 40, nRow, _T("\x43SK\x43\x42"), FMT_TEXT);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 41, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("hd_icd"), tmpStr);
		xls.SetCellText(nCol + 42, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("insline"), tmpStr);
		xls.SetCellText(nCol + 43, nRow, tmpStr, FMT_INTEGER);

		tmpStr.Empty();
		szDate.Empty();
		rs.GetValue(_T("invoicedate"), tmpStr);
		szDate = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		xls.SetCellText(nCol + 44, nRow, szDate, FMT_TEXT);

		rs.GetValue(_T("he_roomid"), tmpStr);
		xls.SetCellText(nCol + 45, nRow, tmpStr, FMT_INTEGER);
		//xls.SetCellText(nCol + 40, nRow, _T("\x30\x31"), FMT_TEXT);
		nIncrement = 1;
		if (m_bOnlyCommander)
		{
			rs.GetValue(_T("generalrank"), tmpStr);
			xls.SetCellText(nCol + 45 + nIncrement, nRow, tmpStr, FMT_TEXT);
			nIncrement++;
			rs.GetValue(_T("position"), tmpStr);
			xls.SetCellText(nCol + 45 + nIncrement, nRow, tmpStr, FMT_TEXT);
			nIncrement++;
		}
		if (m_bIsInsPurpose)
		{
			xls.SetCellText(nCol + 45 + nIncrement, nRow, rs.GetValue(_T("insgroup")), FMT_TEXT);
			nIncrement++;
		}
		if (m_bShowDiff)
		{
			xls.SetCellText(nCol + 45 + nIncrement, nRow, rs.GetValue(_T("diffpaid")), FMT_TEXT);
			nIncrement++;
		}
		if (m_bSoldierPatient)
		{
			xls.SetCellText(nCol + 45 + nIncrement, nRow, rs.GetValue(_T("soldier_rank")), FMT_TEXT);
			nIncrement++;
			xls.SetCellText(nCol + 45 + nIncrement, nRow, rs.GetValue(_T("work_place")), FMT_TEXT);
		}
		rs.MoveNext();
	}
	nRow++;
	xls.SetCellText(0, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
	for (int i = 0; i < 14; i++)
	{
		tmpStr.Format(_T("%f"), nTotal[i]);
		xls.SetCellText(i+10, nRow, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\ChiphikhambenhBHYT_79a.xls"));
	EndWaitCursor();
}

void CFMInsuranceCost79A::OnExportForInsSelect(){
	m_bIsInsPurpose = true;
	OnExportSelect();
	m_bIsInsPurpose = false;

}

int CFMInsuranceCost79A::OnAddFMInsuranceCost79A()
{
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMInsuranceCost79A::OnEditFMInsuranceCost79A()
{
 	if(GetMode() != VM_VIEW) 
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMInsuranceCost79A::OnDeleteFMInsuranceCost79A(){
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
 		OnCancelFMInsuranceCost79A(); 
 	}
	return 0;
}
int CFMInsuranceCost79A::OnSaveFMInsuranceCost79A(){
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
 		//OnFMInsuranceCost79AListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMInsuranceCost79A::OnCancelFMInsuranceCost79A()
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
int CFMInsuranceCost79A::OnFMInsuranceCost79AListLoadData()
{
	return 0;
}
CString CFMInsuranceCost79A::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	CString szSQL, szWhere, szSubWhere, szSubSQL1, szSubSQL2,  szOrderBy, szObjectCondition;
	CString szRooms, szDeptID, szDepts, szCards;
	CString tmpStr, szTemp;

	int nIndex = 0;

	szWhere.Empty();

	szCards.Empty();
	szDepts.Empty();
	szRooms.Empty();
	szDeptID.Empty();
	szSubWhere.Empty();

	if (m_bWithoutEmergency)
		szWhere.AppendFormat(_T(" AND hd_emergency<>'Y' "));
	if (m_bOnlyEmergency)
		szWhere.AppendFormat(_T(" AND hd_emergency='Y' "));

	if (!m_szClerkKey.IsEmpty() && (!m_bUnapproved))
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

	
	for (int i = 0; i < m_wndCardList.GetItemCount(); i++)
	{
		if (m_wndCardList.GetCheck(i))
		{
			tmpStr = m_wndCardList.GetItemText(i, 0);
			nIndex = 0;

			if (tmpStr.Find(_T(',')) > -1)
			{
				for (int i = 0; i < tmpStr.GetLength(); i++)
				{
					if (tmpStr[i] == _T(','))
					{
						szTemp = tmpStr.Mid(nIndex, i - nIndex);
						nIndex = i + 1;
						if (!szCards.IsEmpty())
							szCards += _T(",");
						szCards.AppendFormat(_T("'%s'"), szTemp);
					}
				}

				szTemp = tmpStr.Right(tmpStr.GetLength() - nIndex);
			}
			else
			{
				szTemp = tmpStr;
			}

			if (!szCards.IsEmpty())
				szCards += _T(",");

			szCards.AppendFormat(_T("'%s'"), szTemp);
		}
	}

	if (!szCards.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND substr(hd_cardno,1,2) IN(%s) "), szCards);
		//MessageBox(szCards);
	}

	if (!szRooms.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND (select he_roomid from hms_exam where he_docno=docno and upper(he_deptid)=upper(deptid) and he_status='T' ") \
			                 _T("and he_receptidx=(select max(he_receptidx) from hms_exam where he_docno=docno)) IN(%s) "),
							 szRooms);
	}

	if (!szDepts.IsEmpty() || !szDeptID.IsEmpty())
	{
		if (szDepts.IsEmpty())
			szDepts = _T("''");
		if (!m_bUnapproved)
			szSubWhere.AppendFormat(_T(" AND (fi.hfe_deptid IN(%s)) "),
			                    szDepts, szDeptID);
		else
			szWhere.AppendFormat(_T(" AND hd_enddept IN (%s)"), szDepts);
	}
	if (m_nNumberofDoc > 0)
		szWhere.AppendFormat(_T(" AND hms_getnumberofdocument(hd_docno, '%s', '%s') > %d"), m_szFromDate, m_szToDate, m_nNumberofDoc);
	if (m_bOnlyCommander)
	{
		szWhere.AppendFormat(_T(" AND hd_rank IN (15, 16, 17, 18, 21, 22, 23, 24)"));
	}
	if (m_bOutpatientTreament)
		szWhere.AppendFormat(_T(" AND hd_outpatient = 'Y'"));
	szOrderBy = _T(" insline, insgroup, hdline, examdate, docno, cardno");
	if (m_bIsInsPurpose)
		szOrderBy = _T(" to_date(enddate, 'dd/mm/yyyy'), deptid, hp_surname, hp_midname, hp_firstname");
	if (m_bOnlyCommander || (szCards == _T("'QN'")))
		m_bShowDiff = true;
	else
		m_bShowDiff = false;
	szObjectCondition = _T(" AND ho_type IN ('I', 'C') AND length(hc_cardno) > 0");
	if (m_bSoldierPatient)
		szObjectCondition = _T(" AND hd_object = 1");
	szWhere += szObjectCondition;
	if (m_bShowDiff)
	{
		//phi chenh cho doi tuong cap tuong , QN
		//th da thu phi
		szSubSQL1.Format(_T(" UNION ALL ") \
				_T("  SELECT fi.hfe_deptid as deptid,") \
				_T("         fi.hfe_class,") \
				_T("         fi.hfe_docno as docno,") \
				_T("         fi.hfe_invoiceno as invoiceno,") \
				_T("		 fi.hfe_date as invoicedate,") \
				_T("         fe.hfe_group AS groupid,") \
				_T("         0,0,0,0,0,0,0,0,0,0,0,0,0,0,") \
				_T("         DECODE('P', fe.hfe_status, fe.hfe_patpaid, fe.hfe_patdebt) AS diffpaid") \
				_T("   FROM hms_fee_invoice fi ") \
				_T("   LEFT JOIN hms_fee_view fe ON(fe.hfe_invoiceno=fi.hfe_invoiceno AND fe.hfe_docno=fi.hfe_docno)") \
				_T("   LEFT JOIN hms_fee_list ON (hfl_feeid = hfe_itemid)	") \
				_T("   WHERE fi.hfe_class IN('E') AND fe.hfe_discount = 0 AND fi.hfe_status='P' ") \
				_T("   AND fi.hfe_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') ") \
				_T("   %s"), m_szFromDate, m_szToDate, szSubWhere);
		//phi chenh cho doi tuong cap tuong , QN
		//th da xac nhan phi o khoa
		szSubSQL2.Format(_T(" UNION ALL ") \
					_T("  SELECT hfe_docno as docno,") \
					_T("         0,0,0,0,0,0,0,0,0,0,0,0,0,0,") \
					_T("         DECODE('P', fe.hfe_status, fe.hfe_patpaid, fe.hfe_patdebt) AS difpaid") \
					_T("   FROM hms_fee_view fe ") \
					_T("   LEFT JOIN hms_fee_list ON (hfe_itemid = hfl_feeid)") \
					_T("   WHERE hfe_class IN('E') AND hfe_discount = 0 AND hfe_status NOT IN ('C', 'P')"));
	}
	szSQL.Format(_T("SELECT hd_docno AS docno,") \
				_T("        max(invoiceno) AS invoiceno,") \
				_T("		max(invoicedate) AS invoicedate,") \
				_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) AS pname,") \
				_T("        case when hp_sex='M' then extract(year from hp_birthdate) else 0 end AS mbirthyear,") \
				_T("        case when hp_sex='F' then extract(year from hp_birthdate) else 0 end AS fbirthyear,") \
				_T("        hp_sex AS sex,") \
				_T("        hms_getaddress(hp_provid, hp_distid, hp_villid) as address,") \
				_T("        hd_reldisease as reldisease,") \
				_T("        CASE WHEN LENGTH(hd_cardno) > 15 THEN SUBSTR(hd_cardno,1,15) ELSE hd_cardno END AS cardno,") \
				_T("        hd_insregdate AS regdate,") \
				_T("        hd_insexpdate AS expdate,") \
				_T("        CASE WHEN hd_insline='Y' THEN 1 ELSE 0 END AS hdline,") \
				_T("        hc_regcode AS regplace,") \
				_T("        hc_line AS insline,") \
				_T("        (SELECT distinct ss_default FROM sys_sel WHERE ss_id='hms_insline' AND ss_code=CAST(hc_line AS NVARCHAR2(15))) AS linename,") \
				_T("        (SELECT distinct hig_group FROM hms_insurance_group WHERE hig_idx= hc_groupid) AS insgroup,") \
				_T("        TO_CHAR(hd_admitdate, 'DD/MM/YYYY') AS examdate,") \
				_T("        TO_CHAR(hd_enddate, 'DD/MM/YYYY') AS enddate,") \
				_T("        Cast_Timestamp2Date(hd_enddate)-Cast_Timestamp2Date(hd_admitdate)+1 AS totaltreat,") \
				_T("        hd_icd,") \
				_T("        CASE WHEN length(hd_icd) > 0 THEN hi_name||'['||hd_icd||']' ELSE NULL END AS icd10,") \
				_T("        case when hd_admitdept='C1.3' then 'Y' else 'N' end AS emergency,") \
				_T("        NVL(lower(deptid), lower(hd_admitdept)) AS deptid,") \
				_T("		(select ss_vndesc from sys_sel where ss_id = 'hms_rank' and ss_code = hd_rank) soldier_rank,") \
				_T("		get_syssel_desc('hms_rank', hd_rank) generalrank,") \
				_T("		get_syssel_desc('hms_position', hp_position) position,") \
				_T("		hp_workplace work_place,") \
				_T("        round(SUM(drugfee)) AS drugfee,") \
				_T("        round(SUM(drugfeek)) AS drugfeek,") \
				_T("        round(SUM(bloodfee)) AS bloodfee,") \
				_T("        round(SUM(testfee)) AS testfee,") \
				_T("        round(SUM(pacsfee)) AS pacsfee,") \
				_T("        round(SUM(normtechfee)) AS normtechfee,") \
				_T("        round(SUM(hitechfee)) AS hitechfee,") \
				_T("        round(SUM(materialfee)) AS materialfee,") \
				_T("        round(SUM(examfee)) AS examfee,") \
				_T("        round(SUM(inspaid)) AS cost,") \
				_T("        round(SUM(otherfee)) AS otherfee,") \
				_T("        round(SUM(discount)) AS inspaid,") \
				_T("        round(SUM(inspaid-discount)) AS patpaid,") \
				_T("		round(SUM(diffpaid)) AS diffpaid") \
				_T(" FROM") \
				_T(" (") \
				_T("  SELECT fi.hfe_deptid as deptid,") \
				_T("         fi.hfe_class,") \
				_T("         fi.hfe_docno as docno,") \
				_T("         fi.hfe_invoiceno as invoiceno,") \
				_T("		 fi.hfe_date as invoicedate,") \
				_T("         fe.hfe_group AS groupid,") \
				_T("         CASE WHEN SUBSTR(fe.hfe_group,1,2) IN('A1','A6') AND fe.hfe_group<>'A1400' THEN fe.hfe_inspaid ELSE 0 END AS drugfee,") \
				_T("         CASE WHEN fe.hfe_group='A1400' THEN fe.hfe_inspaid ELSE 0 END AS drugfeek,") \
				_T("         CASE WHEN NVL(hfl_bloodfee, 'X') = 'Y' THEN fe.hfe_inspaid ELSE 0 END AS bloodfee,") \
				_T("         CASE WHEN SUBSTR(fe.hfe_group, 1, 2)='B1' AND hfe_hitech = 'N' AND NVL(hfl_bloodfee, 'X') <> 'Y' THEN fe.hfe_inspaid ELSE 0 END AS testfee,") \
				_T("         CASE WHEN SUBSTR(fe.hfe_group, 1, 2) IN('B2','B3') AND fe.hfe_hitech ='N' THEN fe.hfe_inspaid ELSE 0 END AS pacsfee,") \
				_T("         CASE WHEN SUBSTR(fe.hfe_group, 1, 2) IN('B4','B5') AND fe.hfe_hitech ='N' THEN fe.hfe_inspaid ELSE 0 END AS normtechfee,") \
				_T("         CASE WHEN fe.hfe_hitech='Y' THEN fe.hfe_inspaid ELSE 0 END AS hitechfee,") \
				_T("         CASE WHEN fe.hfe_group='A9000' THEN fe.hfe_inspaid ELSE 0 END AS materialfee,") \
				_T("         CASE WHEN fe.hfe_group='D0000' THEN fe.hfe_inspaid ELSE 0 END AS examfee,") \
				_T("         CASE WHEN fe.hfe_group='F0000' THEN fe.hfe_inspaid ELSE 0 END AS otherfee,") \
				_T("         fe.hfe_cost AS cost,") \
				_T("         fe.hfe_inspaid AS inspaid,") \
				_T("         fe.hfe_discount AS discount,") \
				_T("         fe.hfe_patpaid AS patpaid,") \
				_T("         fe.hfe_diffpaid AS diffpaid") \
				_T("   FROM hms_fee_invoice fi ") \
				_T("   LEFT JOIN hms_fee_view fe ON(fe.hfe_invoiceno=fi.hfe_invoiceno AND fe.hfe_docno=fi.hfe_docno)") \
				_T("   LEFT JOIN hms_fee_list ON (hfl_feeid = hfe_itemid)	") \
				_T("   WHERE fi.hfe_class IN('E') AND fe.hfe_discount > 0 AND fi.hfe_status='P' ") \
				_T("   AND fi.hfe_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') ") \
				_T(" %s %s") \
				_T(" ) tbl") \
				_T(" LEFT JOIN hms_doc ON(docno=hd_docno)") \
				_T(" LEFT JOIN hms_patient ON(hd_patientno=hp_patientno)") \
				_T(" LEFT JOIN hms_card ON(hc_patientno=hp_patientno AND hc_idx=hd_cardidx)") \
				_T(" LEFT JOIN hms_object ON(ho_id=hd_object)") \
				_T(" LEFT JOIN hms_icd ON(hd_icd=hi_icd)") \
				_T(" WHERE hd_status='T' ") \
				_T(" %s ") \
				_T(" GROUP BY hd_docno, hp_surname, hp_midname, hp_firstname, hp_birthdate, hp_sex, hd_cardno, hd_insline, hc_groupid, hc_line, hd_insline, hc_regcode,") \
				_T("          hd_icd, hd_admitdate, hd_enddate, hp_provid, hp_distid, hp_villid, hd_insregdate, hd_insexpdate, hd_reldisease, hi_name, hd_emergency, deptid, hd_admitdept,") \
				_T("		  hd_rank, hp_position, hp_workplace") \
				_T(" HAVING SUM(cost) > 0") \
				_T(" ORDER BY ") \
				_T(" %s") , m_szFromDate, m_szToDate, szSubWhere, szSubSQL1, szWhere, szOrderBy);
	if (m_bUnapproved)
		szSQL.Format(_T(" SELECT hd_docno AS docno,") \
					_T("		he_roomid,") \
					_T("		hd_enddept deptid,") \
					_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) AS pname,") \
					_T("        case when hp_sex='M' then extract(year from hp_birthdate) else 0 end AS mbirthyear,") \
					_T("        case when hp_sex='F' then extract(year from hp_birthdate) else 0 end AS fbirthyear,") \
					_T("        hp_sex AS sex,") \
					_T("        hms_getaddress(hp_provid, hp_distid, hp_villid) as address,") \
					_T("        hd_reldisease as reldisease,") \
					_T("        CASE WHEN LENGTH(hd_cardno) > 15 THEN SUBSTR(hd_cardno,1,15) ELSE hd_cardno END AS cardno,") \
					_T("        hd_insregdate AS regdate,") \
					_T("        hd_insexpdate AS expdate,") \
					_T("        CASE WHEN hd_insline='Y' THEN 1 ELSE 0 END AS hdline,") \
					_T("        hc_regcode AS regplace,") \
					_T("        hc_line AS insline,") \
					_T("        (SELECT distinct ss_default FROM sys_sel WHERE ss_id='hms_insline' AND ss_code=CAST(hc_line AS NVARCHAR2(15))) AS linename,") \
					_T("        (SELECT distinct hig_group FROM hms_insurance_group WHERE hig_idx= hc_groupid) AS insgroup,") \
					_T("        to_char(hd_admitdate, 'DD/MM/YYYY') AS examdate,") \
					_T("        to_char(hd_enddate, 'DD/MM/YYYY') AS enddate,") \
					_T("        Cast_Timestamp2Date(hd_enddate)-Cast_Timestamp2Date(hd_admitdate)+1 AS totaltreat,") \
					_T("        hd_icd,") \
					_T("        CASE WHEN length(hd_icd) > 0 THEN hi_name||'['||hd_icd||']' ELSE NULL END AS icd10,") \
					_T("        case when hd_admitdept='C1.3' then 'Y' else 'N' end AS emergency,") \
					_T("		get_syssel_desc('hms_rank', hd_rank) generalrank,") \
					_T("		get_syssel_desc('hms_position', hp_position) position,") \
					_T("		(select ss_vndesc from sys_sel where ss_id = 'hms_rank' and ss_code = hd_rank) soldier_rank,") \
					_T("		hp_workplace work_place,") \
					_T("        round(SUM(drugfee)) AS drugfee,") \
					_T("        round(SUM(drugfeek)) AS drugfeek,") \
					_T("        round(SUM(bloodfee)) AS bloodfee,") \
					_T("        round(SUM(testfee)) AS testfee,") \
					_T("        round(SUM(pacsfee)) AS pacsfee,") \
					_T("        round(SUM(normtechfee)) AS normtechfee,") \
					_T("        round(SUM(hitechfee)) AS hitechfee,") \
					_T("        round(SUM(materialfee)) AS materialfee,") \
					_T("        round(SUM(examfee)) AS examfee,") \
					_T("        round(SUM(inspaid)) AS cost,") \
					_T("        round(SUM(otherfee)) AS otherfee,") \
					_T("        round(SUM(discount)) AS inspaid,") \
					_T("        round(SUM(inspaid-discount)) AS patpaid,") \
					_T("		round(SUM(diffpaid)) AS diffpaid") \
					_T(" FROM") \
					_T(" (") \
					_T("  SELECT hfe_docno as docno,") \
					_T("         CASE WHEN SUBSTR(hfe_group,1,2) IN('A1','A6') AND hfe_group<>'A1400' THEN hfe_inspaid ELSE 0 END AS drugfee,") \
					_T("         CASE WHEN hfe_group='A1400' THEN hfe_inspaid ELSE 0 END AS drugfeek,") \
					_T("         CASE WHEN NVL(hfl_bloodfee, 'X') = 'Y' THEN hfe_inspaid ELSE 0 END AS bloodfee,") \
					_T("         CASE WHEN SUBSTR(hfe_group, 1, 2)='B1' AND hfe_hitech = 'N' AND NVL(hfl_bloodfee, 'X') <> 'Y' THEN hfe_inspaid ELSE 0 END AS testfee,") \
					_T("         CASE WHEN SUBSTR(hfe_group, 1, 2) IN('B2','B3') AND hfe_hitech ='N' THEN hfe_inspaid ELSE 0 END AS pacsfee,") \
					_T("         CASE WHEN SUBSTR(hfe_group, 1, 2) IN('B4','B5') AND hfe_hitech ='N' THEN hfe_inspaid ELSE 0 END AS normtechfee,") \
					_T("         CASE WHEN hfe_hitech='Y' THEN hfe_inspaid ELSE 0 END AS hitechfee,") \
					_T("         CASE WHEN hfe_group='A9000' THEN hfe_inspaid ELSE 0 END AS materialfee,") \
					_T("         CASE WHEN hfe_group='D0000' THEN hfe_inspaid ELSE 0 END AS examfee,") \
					_T("         CASE WHEN hfe_group='F0000' THEN hfe_inspaid ELSE 0 END AS otherfee,") \
					_T("         hfe_cost AS cost,") \
					_T("         hfe_inspaid AS inspaid,") \
					_T("         hfe_discount AS discount,") \
					_T("         hfe_patpaid AS patpaid,") \
					_T("         hfe_diffpaid AS diffpaid") \
					_T("   FROM hms_fee_view fe ") \
					_T("   LEFT JOIN hms_fee_list ON (hfe_itemid = hfl_feeid)") \
					_T("   WHERE hfe_class IN('E') AND hfe_discount > 0 AND hfe_status NOT IN ('C', 'P')") \
					_T(" %s") \
					_T(" ) tbl") \
					_T(" LEFT JOIN hms_doc ON(docno=hd_docno)") \
					_T(" LEFT JOIN hms_exam ON (hd_docno = he_docno AND hd_doctor = he_doctor)") \
					_T(" LEFT JOIN hms_patient ON(hd_patientno=hp_patientno)") \
					_T(" LEFT JOIN hms_card ON(hc_patientno=hp_patientno AND hc_idx=hd_cardidx)") \
					_T(" LEFT JOIN hms_object ON(ho_id=hd_object)") \
					_T(" LEFT JOIN hms_icd ON(hd_icd=hi_icd)") \
					_T(" WHERE hd_status='T' AND hd_suggestion NOT IN ('C', 'D') ") \
					_T(" AND hd_enddate BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') ") \
					_T(" %s %s") \
					_T(" GROUP BY hd_docno, hp_surname, hp_midname, hp_firstname, hp_birthdate, hp_sex, hd_cardno, hd_insline, hc_groupid, hc_line, hd_insline, hc_regcode,") \
					_T("          hd_icd, hd_admitdate, hd_enddate, hp_provid, hp_distid, hp_villid, hd_insregdate, hd_insexpdate, hd_reldisease, hi_name, hd_emergency, hd_enddept, hd_admitdept, he_roomid,") \
					_T("		  hd_rank, hp_position, hp_workplace") \
					_T(" HAVING SUM(cost) > 0") \
					_T(" ORDER BY ") \
					_T(" %s"),szSubSQL2, m_szFromDate, m_szToDate, szSubWhere, szWhere, szOrderBy);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}


int CFMInsuranceCost79A::OnCardListCheckAll()
{
	for (int i = 0; i < m_wndCardList.GetItemCount(); i++)
	{
		if (!m_wndCardList.GetCheck(i))
		{
			m_wndCardList.SetCheck(i, TRUE);
		}
	}
	return 0;
}

int CFMInsuranceCost79A::OnCardListUnCheckAll()
{
	for (int i = 0; i < m_wndCardList.GetItemCount(); i++)
	{
		if (m_wndCardList.GetCheck(i))
		{
			m_wndCardList.SetCheck(i, FALSE);
		}
	}
	return 0;
}


int CFMInsuranceCost79A::OnDeptListCheckAll()
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

int CFMInsuranceCost79A::OnDeptListUnCheckAll()
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


int CFMInsuranceCost79A::OnRoomListCheckAll()
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

int CFMInsuranceCost79A::OnRoomListUnCheckAll()
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
