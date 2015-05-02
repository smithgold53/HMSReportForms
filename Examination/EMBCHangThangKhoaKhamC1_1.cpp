#include "stdafx.h"
#include "EMBCHangThangKhoaKhamC1_1.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptBCHangThangKhoaKhamC1_1* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnReportPeriodLoadData();//
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnToDateCheckValue();
} 
static void _OnExamRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CRptBCHangThangKhoaKhamC1_1* )pWnd)->OnExamRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnExamRoomSelendokFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnExamRoomSelendok();
}
/*static void _OnExamRoomSetfocusFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnExamRoomKillfocus();
}*/
/*static void _OnExamRoomKillfocusFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnExamRoomKillfocus();
}*/
static long _OnExamRoomLoadDataFnc(CWnd *pWnd){
	return ((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnExamRoomLoadData();
}
/*static void _OnExamRoomAddNewFnc(CWnd *pWnd){
	((CRptBCHangThangKhoaKhamC1_1 *)pWnd)->OnExamRoomAddNew();
}*/
static void _OnPatientSelectFnc(CWnd *pWnd){
	 ((CRptBCHangThangKhoaKhamC1_1*)pWnd)->OnPatientSelect();
}
static void _OnTurnSelectFnc(CWnd *pWnd){
	 ((CRptBCHangThangKhoaKhamC1_1*)pWnd)->OnTurnSelect();
}
static void _OnDrugOrderSelectFnc(CWnd *pWnd){
	 ((CRptBCHangThangKhoaKhamC1_1*)pWnd)->OnDrugOrderSelect();
}
static void _OnGoDeptSelectFnc(CWnd *pWnd){
	 ((CRptBCHangThangKhoaKhamC1_1*)pWnd)->OnGoDeptSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CRptBCHangThangKhoaKhamC1_1 *pVw = (CRptBCHangThangKhoaKhamC1_1 *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CRptBCHangThangKhoaKhamC1_1 *pVw = (CRptBCHangThangKhoaKhamC1_1 *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CRptBCHangThangKhoaKhamC1_1 *pVw = (CRptBCHangThangKhoaKhamC1_1 *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddRptBCHangThangKhoaKhamC1_1Fnc(CWnd *pWnd){
	 return ((CRptBCHangThangKhoaKhamC1_1*)pWnd)->OnAddRptBCHangThangKhoaKhamC1_1();
} 
static int _OnEditRptBCHangThangKhoaKhamC1_1Fnc(CWnd *pWnd){
	 return ((CRptBCHangThangKhoaKhamC1_1*)pWnd)->OnEditRptBCHangThangKhoaKhamC1_1();
} 
static int _OnDeleteRptBCHangThangKhoaKhamC1_1Fnc(CWnd *pWnd){
	 return ((CRptBCHangThangKhoaKhamC1_1*)pWnd)->OnDeleteRptBCHangThangKhoaKhamC1_1();
} 
static int _OnSaveRptBCHangThangKhoaKhamC1_1Fnc(CWnd *pWnd){
	 return ((CRptBCHangThangKhoaKhamC1_1*)pWnd)->OnSaveRptBCHangThangKhoaKhamC1_1();
} 
static int _OnCancelRptBCHangThangKhoaKhamC1_1Fnc(CWnd *pWnd){
	 return ((CRptBCHangThangKhoaKhamC1_1*)pWnd)->OnCancelRptBCHangThangKhoaKhamC1_1();
} 
CRptBCHangThangKhoaKhamC1_1::CRptBCHangThangKhoaKhamC1_1(CWnd *pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 185;
	m_szTitle = _T("\x42\xE1o \x63\xE1o h\xE0ng th\xE1ng kho\x61 kh\xE1m \x43\x31-\x31 ");
	SetDefaultValues();
}
CRptBCHangThangKhoaKhamC1_1::~CRptBCHangThangKhoaKhamC1_1(){
}
void CRptBCHangThangKhoaKhamC1_1::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 490, 150);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 245, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 250, 30, 330, 55);
	m_wndReportPeriod.Create(this,335, 30, 485, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 245, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 60, 330, 85);
	m_wndToDate.Create(this,335, 60, 485, 85); 
	m_wndExamRoomLabel.Create(this, _T("Exam Room"), 10, 90, 90, 115);
	m_wndExamRoom.Create(this,95, 90, 485, 115); 
	m_wndPatient.Create(this, _T("Patient"), 95, 120, 185, 145);
	m_wndTurn.Create(this, _T("Turn"), 190, 120, 280, 145);
	m_wndDrugOrder.Create(this, _T("Drug Order"), 285, 120, 375, 145);
	m_wndGoDept.Create(this, _T("Go Dept"), 395, 120, 485, 145);
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 300, 155, 410, 180);
	m_wndExport.Create(this, _T("&Export"), 415, 155, 490, 180);

}
void CRptBCHangThangKhoaKhamC1_1::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	//m_wndExamRoom.SetCheckValue(true);
	m_wndExamRoom.LimitText(35);
	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);
	m_wndExamRoom.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndExamRoom.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);
	m_bPatient = TRUE;
	m_bTurn = TRUE;
	m_bDrugOrder = TRUE;
	m_bGoDept = TRUE;

}
void CRptBCHangThangKhoaKhamC1_1::OnSetWindowEvents(){
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
	m_wndPatient.SetEvent(WE_CLICK, _OnPatientSelectFnc);
	m_wndTurn.SetEvent(WE_CLICK, _OnTurnSelectFnc);
	m_wndDrugOrder.SetEvent(WE_CLICK, _OnDrugOrderSelectFnc);
	m_wndGoDept.SetEvent(WE_CLICK, _OnGoDeptSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
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
void CRptBCHangThangKhoaKhamC1_1::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndExamRoom.GetDlgCtrlID(), m_szExamRoomKey);
	DDX_Check(pDX, m_wndPatient.GetDlgCtrlID(), m_bPatient);
	DDX_Check(pDX, m_wndTurn.GetDlgCtrlID(), m_bTurn);
	DDX_Check(pDX, m_wndDrugOrder.GetDlgCtrlID(), m_bDrugOrder);
	DDX_Check(pDX, m_wndGoDept.GetDlgCtrlID(), m_bGoDept);

}
void CRptBCHangThangKhoaKhamC1_1::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CRptBCHangThangKhoaKhamC1_1::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CRptBCHangThangKhoaKhamC1_1::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szExamRoomKey.Empty();
	m_bPatient=	TRUE;
	m_bTurn=FALSE;
	m_bDrugOrder=FALSE;
	m_bGoDept=FALSE;

}
int CRptBCHangThangKhoaKhamC1_1::SetMode(int nMode){
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
/*void CRptBCHangThangKhoaKhamC1_1::OnYearChange(){
	
} */
/*void CRptBCHangThangKhoaKhamC1_1::OnYearSetfocus(){
	
} */
/*void CRptBCHangThangKhoaKhamC1_1::OnYearKillfocus(){
	
} */
int CRptBCHangThangKhoaKhamC1_1::OnYearCheckValue(){
	return 0;
} 
void CRptBCHangThangKhoaKhamC1_1::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CRptBCHangThangKhoaKhamC1_1::OnReportPeriodSelendok(){
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
/*void CRptBCHangThangKhoaKhamC1_1::OnReportPeriodSetfocus(){
	
}*/
/*void CRptBCHangThangKhoaKhamC1_1::OnReportPeriodKillfocus(){
	
}*/
long CRptBCHangThangKhoaKhamC1_1::OnReportPeriodLoadData(){
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
/*void CRptBCHangThangKhoaKhamC1_1::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CRptBCHangThangKhoaKhamC1_1::OnFromDateChange(){
	
} */
/*void CRptBCHangThangKhoaKhamC1_1::OnFromDateSetfocus(){
	
} */
/*void CRptBCHangThangKhoaKhamC1_1::OnFromDateKillfocus(){
	
} */
int CRptBCHangThangKhoaKhamC1_1::OnFromDateCheckValue(){
	return 0;
} 
/*void CRptBCHangThangKhoaKhamC1_1::OnToDateChange(){
	
} */
/*void CRptBCHangThangKhoaKhamC1_1::OnToDateSetfocus(){
	
} */
/*void CRptBCHangThangKhoaKhamC1_1::OnToDateKillfocus(){
	
} */
int CRptBCHangThangKhoaKhamC1_1::OnToDateCheckValue(){
	return 0;
} 
void CRptBCHangThangKhoaKhamC1_1::OnExamRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CRptBCHangThangKhoaKhamC1_1::OnExamRoomSelendok(){
	 
}
/*void CRptBCHangThangKhoaKhamC1_1::OnExamRoomSetfocus(){
	
}*/
/*void CRptBCHangThangKhoaKhamC1_1::OnExamRoomKillfocus(){
	
}*/
long CRptBCHangThangKhoaKhamC1_1::OnExamRoomLoadData(){
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
/*void CRptBCHangThangKhoaKhamC1_1::OnExamRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CRptBCHangThangKhoaKhamC1_1::OnPatientSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CRptBCHangThangKhoaKhamC1_1::OnTurnSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CRptBCHangThangKhoaKhamC1_1::OnDrugOrderSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CRptBCHangThangKhoaKhamC1_1::OnGoDeptSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CRptBCHangThangKhoaKhamC1_1::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString tmpStr, szSQL;
	CReportSection *rptDetail = NULL;
	CReportSection *rptSector = NULL;
	CRecord rs(&pMF->m_db);
	CString szDept = pMF->m_szDept;
	//if (szDept != _T("C1.1"))
	//{
	//	AfxMessageBox(_T("This report can only run by C1.1 user!"), MB_ICONASTERISK);
	//	return;
	//}
	int nOldGroup=-1, nNewGroup = -1, nIndex = 0;
	int nNewTit = 0, nOldTit = 0;
	long tc[13], sc[13], c[13];	
	CString szCol, szData;
	if(!rpt.Init(_T("Reports/HMS/HE_BAOCAOTONGHOPTHANGPHONGKHAM_C1_1.RPT")) )
		return ;
	CReportSection *pTemp = rpt.GetGroupHeader(1);
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	BeginWaitCursor();
	szSQL = GetQueryString();
	//QueryOpener(szSQL);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	for (int i =0; i <= 12; i++)
	{
		tc[i] = sc[i] = c[i] = 0;
	}
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("title"), nNewTit);
		if (nNewTit != nOldTit)
		{
			if (nOldTit != 0)
			{
				rptSector = rpt.AddDetail(rpt.GetGroupFooter(1));
				for (int i =1; i <= 12; i++){
					szCol.Format(_T("s%d"), i+2);
					szData.Format(_T("%ld"), tc[i-1]);
					rptSector->SetValue(szCol, szData);
					tc[i-1] = sc[i-1] = c[i-1] = 0;
				}	
			}
			//1:So kham,2:So luot,3:So ke don,4:Vao vien
			switch (nNewTit){
			case 1:
				tmpStr.Format(_T("S\x1ED1 Ng\x1B0\x1EDDi Kh\xE1m"));
				break;
			case 2:
				tmpStr.Format(_T("S\x1ED1 L\x1B0\x1EE3t Kh\xE1m"));
				break;
			case 3:
				tmpStr.Format(_T("S\x1ED1 K\xEA \x110\x1A1n"));
				break;
			case 4:
				tmpStr.Format(_T("S\x1ED1 V\xE0o Vi\x1EC7n"));
				break;
			}
			rptSector = rpt.AddDetail(pTemp);
			rptSector->SetValue(_T("Group"), tmpStr);
			nOldTit = nNewTit;
		}
		nNewGroup = ToInt(rs.GetValue(_T("extype")));
		//extype = 2->Noi,3->Ngoai,>3->Chuyen Khoa
		if(nNewGroup == 2 && nOldGroup != nNewGroup)
		{
			//if(sc[10] > 0){
			//	rptSector = rpt.AddDetail(pTemp);
			//	rptSector->SetValue(_T("Group"), _T("\x43\x1ED9ng"));
			//	for (int i =1; i <= 10; i++){
			//		szCol.Format(_T("%d"), i+2);
			//		szData.Format(_T("%ld"), sc[i-1]);
			//		rptDetail->SetValue(szCol, szData);
			//		sc[i-1] = 0;
			//	}
			//}
			rptSector = rpt.AddDetail(pTemp);
			rptSector->SetValue(_T("Group"), _T("- N\x1ED9i"));
			nOldGroup = nNewGroup;
			nIndex = 0;	
		}
		if(nNewGroup == 3 && nOldGroup != nNewGroup)
		{
			rptSector = rpt.AddDetail(pTemp);
			rptSector->SetValue(_T("Group"), _T("- Ngo\x1EA1i"));
			nOldGroup = nNewGroup;
			nIndex = 0;	
		}
		if(nNewGroup > 3 && nOldGroup == 3)
		{	
			rptSector = rpt.AddDetail(rpt.GetGroupFooter(1));
			for (int i =1; i <= 12; i++)
			{
				szCol.Format(_T("s%d"), i+2);
				szData.Format(_T("%ld"), tc[i-1]);
				rptSector->SetValue(szCol, szData);
				tc[i-1] = sc[i-1] = c[i-1] = 0;
			}	
			rptSector = rpt.AddDetail(pTemp);
			rptSector->SetValue(_T("Group"), _T("- \x43huy\xEAn kho\x61"));
			nOldGroup = nNewGroup;				
		}
		rptDetail = rpt.AddDetail();
		nIndex++;
		tmpStr.Format(_T("%d"), nIndex);
		rptDetail->SetValue(_T("0"), tmpStr);
		rs.GetValue(_T("roomid"), tmpStr);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("roomname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		for (int i =1; i <= 12; i++)
		{
			tmpStr.Format(_T("%d"), i+2);
			szCol.Format(_T("c%d"), i);
			szData = rs.GetValue(szCol);
			c[i-1] = ToInt(szData);
			tc[i-1] += c[i-1];
			sc[i-1] += c[i-1];
			rptDetail->SetValue(tmpStr, szData);
		}
		rs.MoveNext();
	}
	rptSector = rpt.AddDetail(rpt.GetGroupFooter(1));
	for (int i =1; i <= 12; i++)
	{
		szCol.Format(_T("s%d"), i+2);
		szData.Format(_T("%ld"), tc[i-1]);
		rptSector->SetValue(szCol, szData);
		tc[i-1] = sc[i-1] = c[i-1] = 0;
	}	
	tmpStr = pMF->GetSysDate();
	CString szDate;
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();
} 
void CRptBCHangThangKhoaKhamC1_1::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CString tmpStr, szSQL, szCol, szData;
	int nOldTit = 0, nNewTit = 0, nOldGroup = -1, nNewGroup = -1;
	CString szDept = pMF->m_szDept;
	if (szDept != _T("C1.1"))
	{
		AfxMessageBox(_T("This report can only run by C1.1 user!"), MB_ICONASTERISK);
		return;
	}
	CRecord rs(&pMF->m_db);
	BeginWaitCursor();
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	CellFormat tf(&xls), of(&xls);
	tf.SetColor(XLCOLOR_BLUE);
	tf.SetBold(true);
	of.SetColor(XLCOLOR_SKYBLUE);
	of.SetBold(true);
	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 30);
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
	xls.SetColumnWidth(13, 10);
	xls.SetColumnWidth(14, 13);
	int nCol = 0, nRow = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	xls.SetCellMergedColumns(nCol, nRow , 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	rs.GetValue(_T("tenkhoa"), tmpStr);
	xls.SetCellText(nCol, nRow + 2, tmpStr, FMT_TEXT, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 14);
	xls.SetCellText(nCol, nRow + 3, _T("\x42\xC1O \x43\xC1O T\x1ED4NG H\x1EE2P TH\xC1NG PH\xD2NG KH\xC1M \x43\x31-\x31"), FMT_TEXT | FMT_CENTER, true, 14);
	xls.SetCellMergedColumns(nCol, nRow + 4, 14);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol, nRow + 5, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 5, _T("M\xE3 ph\xF2ng"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 5, _T("Ph\xF2ng kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 3, nRow + 5, _T("Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 5, _T("BH Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 5, nRow + 5, _T("\x42\x1EA1n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 6, nRow + 5, _T("Tr\x1EBB < \x36 tu\x1ED5i"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 7, nRow + 5, _T("\x43h\xEDnh s\xE1\x63h"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 8, nRow + 5, _T("\x44\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 9, nRow + 5, _T("BHYT Kh\xE1\x63"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 10, nRow + 5, _T("BH TNSQ"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 11, nRow + 5, _T("BH QÐ"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 12, nRow + 5, _T("BH Qu\xE2n nh\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 13, nRow + 5, _T("\x42H Y t\x1EBF( n\x1EE3 th\x1EBB)"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 14, nRow + 5, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 11);
	int nIndex = 1;
	long c[13], tc[13];

	for (int i = 0; i<=12; i++)
	{
		c[i] = 0;
		tc[i] = 0;
	}
	nRow = 5;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("title"), nNewTit);
		if (nNewTit != nOldTit)
		{
			if (nOldTit != 0)
			{
				nRow++;
				for (int i =1; i <= 12; i++){
					szData.Format(_T("%ld"), tc[i-1]);
					xls.SetCellMergedColumns(0, nRow, 3);
					xls.SetCellText(0, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
					xls.SetCellText(i+2, nRow, szData, FMT_INTEGER, true);
					tc[i-1] = c[i-1] = 0;
				}	
			}
			//1:So kham,2:So luot,3:So ke don,4:Vao vien
			switch (nNewTit){
			case 1:
				tmpStr.Format(_T("S\x1ED1 Ng\x1B0\x1EDDi Kh\xE1m"));
				break;
			case 2:
				tmpStr.Format(_T("S\x1ED1 L\x1B0\x1EE3t Kh\xE1m"));
				break;
			case 3:
				tmpStr.Format(_T("S\x1ED1 K\xEA \x110\x1A1n"));
				break;
			case 4:
				tmpStr.Format(_T("S\x1ED1 V\xE0o Vi\x1EC7n"));
				break;
			}
			nRow++;
			xls.SetCellMergedColumns(1, nRow, 14);
			xls.SetCellText(1, nRow, tmpStr, &tf);
			nOldTit = nNewTit;
		}
		int check=-1;
		check = ToInt(rs.GetValue(_T("roomid")));
		nNewGroup = ToInt(rs.GetValue(_T("extype")));
		//extype = 2->Noi,3->Ngoai,>3->Chuyen Khoa
		if(nNewGroup == 2&& nOldGroup != nNewGroup){
			//if(sc[10] > 0){
			//	rptSector = rpt.AddDetail(pTemp);
			//	rptSector->SetValue(_T("Group"), _T("\x43\x1ED9ng"));
			//	for (int i =1; i <= 10; i++){
			//		szCol.Format(_T("%d"), i+2);
			//		szData.Format(_T("%ld"), sc[i-1]);
			//		rptDetail->SetValue(szCol, szData);
			//		sc[i-1] = 0;
			//	}
			//}
			//rptSector = rpt.AddDetail(pTemp);
			//rptSector->SetValue(_T("Group"), _T("N\x1ED9i"));
			nRow++;
			xls.SetCellMergedColumns(1, nRow, 14);
			xls.SetCellText(1, nRow, _T("N\x1ED9i"), &of);
			nOldGroup = nNewGroup;
			nIndex = 0;	
		}
		if(nNewGroup == 3 && nOldGroup != nNewGroup){
			nRow++;
			xls.SetCellText(1, nRow, _T("Ngo\x1EA1i"), &of);
			nOldGroup = nNewGroup;
			nIndex = 0;	
		}
		if((nNewGroup > 3) && nOldGroup <= 3){	
			nRow++;
			for (int i =1; i <= 12; i++){
				szData.Format(_T("%ld"), tc[i-1]);
				xls.SetCellMergedColumns(0, nRow, 3);
				xls.SetCellText(0, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
				xls.SetCellText(i+2, nRow, szData, FMT_INTEGER, true);
				tc[i-1] = c[i-1] = 0;
			}
			nRow++;
			xls.SetCellText(1, nRow, _T("\x43huy\xEAn kho\x61"), &of);
			nOldGroup = nNewGroup;				
		}
		nRow++;
		nIndex++;
		tmpStr.Format(_T("%d"), nIndex);
		xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);
		rs.GetValue(_T("roomid"), tmpStr);
		xls.SetCellText(1, nRow, tmpStr, FMT_INTEGER);
		rs.GetValue(_T("roomname"), tmpStr);
		xls.SetCellText(2, nRow, tmpStr, FMT_TEXT);
		for (int i =1; i <= 12; i++){
			tmpStr.Format(_T("%d"), i+2);
			szCol.Format(_T("c%d"), i);
			szData = rs.GetValue(szCol);
			c[i-1] = ToInt(szData);
			tc[i-1] += c[i-1];
			xls.SetCellText(i+2, nRow, szData, FMT_INTEGER);
		}
		rs.MoveNext();
	}	
	nRow++;
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
	for( int i = 1; i<= 12; i++)
	{
		tmpStr.Format(_T("%d"), tc[i-1]);
		xls.SetCellText(i+2, nRow, tmpStr, FMT_NUMBER1, true);
	}
	CString szDate;
	EndWaitCursor();
	xls.Save(_T("Exports\\BAO CAO HANG THANG KHOA KHAM C1 - 1.xls"));
} 
void CRptBCHangThangKhoaKhamC1_1::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CRptBCHangThangKhoaKhamC1_1::OnAddRptBCHangThangKhoaKhamC1_1(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CRptBCHangThangKhoaKhamC1_1::OnEditRptBCHangThangKhoaKhamC1_1(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CRptBCHangThangKhoaKhamC1_1::OnDeleteRptBCHangThangKhoaKhamC1_1(){
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
 		OnCancelRptBCHangThangKhoaKhamC1_1(); 
 	}
	return 0;
}
int CRptBCHangThangKhoaKhamC1_1::OnSaveRptBCHangThangKhoaKhamC1_1(){
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
 		//OnRptBCHangThangKhoaKhamC1_1ListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CRptBCHangThangKhoaKhamC1_1::OnCancelRptBCHangThangKhoaKhamC1_1(){
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
int CRptBCHangThangKhoaKhamC1_1::OnRptBCHangThangKhoaKhamC1_1ListLoadData(){
	return 0;
}
CString CRptBCHangThangKhoaKhamC1_1::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szSubSQL, szWhere;
	UpdateData(true);
	szWhere.Format(_T(" AND he_examdate BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s') AND he_deptid = '%s'"), m_szFromDate, m_szToDate, pMF->GetCurrentDepartmentID());
	if (m_bPatient){
		if (!szSubSQL.IsEmpty())
			szSubSQL.AppendFormat(_T("UNION ALL"));
		szSubSQL.AppendFormat(_T(" SELECT 1 as title, hd_docno,") \
					_T("	 he_roomid,") \
					_T("     case when hd_object=1 then 1 else 0 end as c1,") \
					_T("     case when hd_object=2 then 1 else 0 end as c2,") \
					_T("     case when hd_object=8 then 1 else 0 end as c3,") \
					_T("     case when hd_object in(6,9) then 1 else 0 end as c4,") \
					_T("     case when hd_object=3 then 1 else 0 end as c5,") \
					_T("     case when hd_object=7 then 1 else 0 end as c6,") \
					_T("     case when hd_object=4 then 1 else 0 end as c7,") \
					_T("     case when hd_object=5 then 1 else 0 end as c8,") \
					_T("     case when hd_object=10 then 1 else 0 end as c9,") \
					_T("     case when hd_object=11 then 1 else 0 end as c10,") \
					_T("	 case when hd_object=12 then 1 else 0 end as c11,") \
					_T("     1 as c12 ") \
					_T(" FROM     (SELECT he_docno          docno, ") \
					_T("                  Min(he_receptidx) receptidx ") \
					_T("           FROM   hms_exam ") \
					_T("           WHERE  he_status IN ( 'T', 'P' ) ") \
					_T("              AND he_deptid = 'C1.1' ") \
					_T("           GROUP  BY he_docno) ") \
					_T(" LEFT JOIN hms_exam ON( he_docno = docno ") \
					_T("                        AND he_receptidx = receptidx ) ") \
					_T(" LEFT JOIN hms_doc ON ( he_docno = hd_docno ) ") \
					_T(" WHERE 1=1 ") \
					_T(" %s"), szWhere);

	}
	if (m_bTurn){
		if (!szSubSQL.IsEmpty())
			szSubSQL.AppendFormat(_T("UNION ALL"));
		szSubSQL.AppendFormat(_T(" SELECT 2 as title, hd_docno, he_roomid, ") \
					_T(" 				case when hd_object=1 then 1 else 0 end as c1,") \
					_T(" 				case when hd_object=2 then 1 else 0 end as c2,") \
					_T(" 				case when hd_object=8 then 1 else 0 end as c3,") \
					_T(" 				case when hd_object in(6,9) then 1 else 0 end as c4,") \
					_T(" 				case when hd_object=3 then 1 else 0 end as c5,") \
					_T(" 				case when hd_object=7 then 1 else 0 end as c6,") \
					_T(" 				case when hd_object=4 then 1 else 0 end as c7,") \
					_T(" 				case when hd_object=5 then 1 else 0 end as c8,") \
					_T(" 				case when hd_object=10 then 1 else 0 end as c9,") \
					_T("				case when hd_object=11 then 1 else 0 end as c10,") \
					_T("				case when hd_object=12 then 1 else 0 end as c11,") \
					_T(" 				1 as c12 ") \
					_T("		FROM hms_doc") \
					_T("		LEFT JOIN hms_exam ON (he_docno = hd_docno)") \
					_T("		WHERE he_status IN ('P', 'T') ") \
					_T("		%s"), szWhere);
	}
	//QueryOpener(szSubSQL);
	//1 BN cap 2 don -->luot cap/BN cap
	if (m_bDrugOrder){
		szWhere.Empty();
		if (!szSubSQL.IsEmpty())
			szSubSQL.AppendFormat(_T("UNION ALL"));
		if (!m_szExamRoomKey.IsEmpty())
			szWhere.Format(_T(" AND hpo_roomid = '%s'"), m_szExamRoomKey);
		szSubSQL.AppendFormat(_T(" SELECT 3 as title, hd_docno, hpo_roomid as he_roomid, ") \
						_T(" 				case when hd_object=1 then 1 else 0 end as c1,") \
						_T(" 				case when hd_object=2 then 1 else 0 end as c2,") \
						_T(" 				case when hd_object=8 then 1 else 0 end as c3,") \
						_T(" 				case when hd_object in(6,9) then 1 else 0 end as c4,") \
						_T(" 				case when hd_object=3 then 1 else 0 end as c5,") \
						_T(" 				case when hd_object=7 then 1 else 0 end as c6,") \
						_T(" 				case when hd_object=4 then 1 else 0 end as c7,") \
						_T(" 				case when hd_object=5 then 1 else 0 end as c8,") \
						_T(" 				case when hd_object =10 then 1 else 0 end as c9,") \
						_T(" 				case when hd_object =11 then 1 else 0 end as c10,") \
						_T("				case when hd_object=12 then 1 else 0 end as c11,") \
						_T(" 				1 as c12 ") \
						_T(" 		FROM hms_doc") \
						_T(" 		LEFT JOIN hms_pharmaorder ON(hpo_docno=hd_docno)") \
						_T(" 		WHERE   hpo_orderdate BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
						_T(" 		AND hpo_orderstatus IN ('A', 'S')") \
						_T("		AND hpo_deptid = '%s' ") \
						_T("		%s"), m_szFromDate, m_szToDate, pMF->GetCurrentDepartmentID(), szWhere);
	}
#ifdef COP
	_msg(_T("%s"), szSubSQL);
#endif
	if (m_bGoDept){
		if (!szSubSQL.IsEmpty())
			szSubSQL.AppendFormat(_T("UNION ALL"));
		szSubSQL.AppendFormat(_T(" SELECT 4 as title, hd_docno, he_roomid,") \
					_T(" 			case when hd_object=1 then 1 else 0 end as c1,") \
					_T(" 			case when hd_object=2 then 1 else 0 end as c2,") \
					_T(" 			case when hd_object=8 then 1 else 0 end as c3,") \
					_T(" 			case when hd_object in(6, 9) then 1 else 0 end as c4,") \
					_T(" 			case when hd_object=3 then 1 else 0 end as c5,") \
					_T(" 			case when hd_object=7 then 1 else 0 end as c6,") \
					_T(" 			case when hd_object=4 then 1 else 0 end as c7,") \
					_T(" 			case when hd_object=5 then 1 else 0 end as c8,") \
					_T(" 			case when hd_object=10 then 1 else 0 end as c9,") \
					_T(" 			case when hd_object=11 then 1 else 0 end as c10,") \
					_T("			case when hd_object=12 then 1 else 0 end as c11,") \
					_T(" 			1 as c12 ") \
					_T(" 	FROM hms_doc") \
					_T("	LEFT JOIN hms_exam ON (he_docno = hd_docno AND he_doctor = hd_doctor)") \
					_T("	LEFT JOIN hms_clinical_record ON (hd_docno = hcr_docno)") \
					_T(" 	WHERE hcr_admitdate BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
					_T("    and hcr_numinward > 0 ") \
					_T("	AND hd_enddept = '%s'"), m_szFromDate, m_szToDate, pMF->GetCurrentDepartmentID());
	}
	szSQL.Format(_T(" SELECT title, hrl_type as extype, hrl_id as roomid, hrl_name as roomname,") \
					_T(" 			sum(c1) as c1,") \
					_T(" 			sum(c2) as c2,") \
					_T(" 			sum(c3) as c3,") \
					_T(" 			sum(c4) as c4,") \
					_T(" 			sum(c5) as c5,") \
					_T(" 			sum(c6) as c6,") \
					_T(" 			sum(c7) as c7,") \
					_T(" 			sum(c8) as c8,") \
					_T(" 			sum(c9) as c9,") \
					_T(" 			sum(c10) as c10,") \
					_T(" 			sum(c11) as c11,") \
					_T("			sum(c12) as c12") \
					_T(" 	FROM") \
					_T(" 	(%s) tbl") \
					_T("	LEFT JOIN hms_roomlist ON(hrl_deptid='%s' and hrl_id=he_roomid)") \
					_T(" 	GROUP BY hrl_type, hrl_id, hrl_name, title") \
					_T(" 	ORDER BY title, cast(extype as integer), hrl_id"), szSubSQL, pMF->GetCurrentDepartmentID());
_fmsg(_T("%s"), szSQL);
	return szSQL;
}