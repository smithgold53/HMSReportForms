#include "stdafx.h"
#include "FMDanhsachBNNopTien_VP.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CrptDanhsachBNNopTien_VP *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CrptDanhsachBNNopTien_VP* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CrptDanhsachBNNopTien_VP *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CrptDanhsachBNNopTien_VP *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CrptDanhsachBNNopTien_VP *)pWnd)->OnToDateCheckValue();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CrptDanhsachBNNopTien_VP* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CrptDanhsachBNNopTien_VP *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnObjectAddNew();
}*/
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CrptDanhsachBNNopTien_VP* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CrptDanhsachBNNopTien_VP *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CrptDanhsachBNNopTien_VP *)pWnd)->OnClerkAddNew();
}*/
static void _OnOrderCashVoucherSelectFnc(CWnd *pWnd){
	 ((CrptDanhsachBNNopTien_VP*)pWnd)->OnOrderCashVoucherSelect();
}
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CrptDanhsachBNNopTien_VP*)pWnd)->OnLockedSelect();
}
static void _OnPrintSelectFnc(CWnd *pWnd){
	CrptDanhsachBNNopTien_VP *pVw = (CrptDanhsachBNNopTien_VP *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CrptDanhsachBNNopTien_VP *pVw = (CrptDanhsachBNNopTien_VP *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CrptDanhsachBNNopTien_VP *pVw = (CrptDanhsachBNNopTien_VP *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddrptDanhsachBNNopTien_VPFnc(CWnd *pWnd){
	 return ((CrptDanhsachBNNopTien_VP*)pWnd)->OnAddrptDanhsachBNNopTien_VP();
} 
static int _OnEditrptDanhsachBNNopTien_VPFnc(CWnd *pWnd){
	 return ((CrptDanhsachBNNopTien_VP*)pWnd)->OnEditrptDanhsachBNNopTien_VP();
} 
static int _OnDeleterptDanhsachBNNopTien_VPFnc(CWnd *pWnd){
	 return ((CrptDanhsachBNNopTien_VP*)pWnd)->OnDeleterptDanhsachBNNopTien_VP();
} 
static int _OnSaverptDanhsachBNNopTien_VPFnc(CWnd *pWnd){
	 return ((CrptDanhsachBNNopTien_VP*)pWnd)->OnSaverptDanhsachBNNopTien_VP();
} 
static int _OnCancelrptDanhsachBNNopTien_VPFnc(CWnd *pWnd){
	 return ((CrptDanhsachBNNopTien_VP*)pWnd)->OnCancelrptDanhsachBNNopTien_VP();
} 
CrptDanhsachBNNopTien_VP::CrptDanhsachBNNopTien_VP(CWnd *pParent)
	{
	m_nDlgWidth = 420;
	m_nDlgHeight = 215;
	m_szTitle = _T("\x44\x61nh s\xE1\x63h \x62\x1EC7nh nh\xE2n n\x1ED9p ti\x1EC1n");
	SetDefaultValues();
}
CrptDanhsachBNNopTien_VP::~CrptDanhsachBNNopTien_VP(){
}
void CrptDanhsachBNNopTien_VP::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 415, 180);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 205, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 215, 30, 295, 55);
	m_wndReportPeriod.Create(this,300, 30, 410, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 205, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 215, 60, 295, 85);
	m_wndToDate.Create(this,300, 60, 410, 85); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 90, 90, 115);
	m_wndObject.Create(this,95, 90, 410, 115); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 120, 90, 145);
	m_wndClerk.Create(this,95, 120, 410, 145); 
	m_wndOrderCashVoucher.Create(this, _T("Order cash voucher "), 145, 150, 310, 175);
	m_wndLocked.Create(this, _T("Locked"), 315, 150, 410, 175);
	m_wndPrint.Create(this, _T("&Print"), 250, 185, 330, 210);
	m_wndExport.Create(this, _T("&Export"), 335, 185, 415, 210);
}
void CrptDanhsachBNNopTien_VP::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(false);
	m_wndReportPeriod.SetCheckValue(false);
	m_wndReportPeriod.LimitText(35);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(false);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(false);
	m_wndObject.SetCheckValue(false);
	m_wndObject.LimitText(35);
	m_wndClerk.SetCheckValue(false);
	m_wndClerk.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndObject.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

}
void CrptDanhsachBNNopTien_VP::OnSetWindowEvents(){
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
	m_wndClerk.SetEvent(WE_SELENDOK, _OnClerkSelendokFnc);
	//m_wndClerk.SetEvent(WE_SETFOCUS, _OnClerkSetfocusFnc);
	//m_wndClerk.SetEvent(WE_KILLFOCUS, _OnClerkKillfocusFnc);
	m_wndClerk.SetEvent(WE_SELCHANGE, _OnClerkSelectChangeFnc);
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	//m_wndClerk.SetEvent(WE_ADDNEW, _OnClerkAddNewFnc);
	m_wndOrderCashVoucher.SetEvent(WE_CLICK, _OnOrderCashVoucherSelectFnc);
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDateTime();
	SetMode(VM_EDIT);

}
void CrptDanhsachBNNopTien_VP::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndOrderCashVoucher.GetDlgCtrlID(), m_bOrderCashVoucher);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);

}
void CrptDanhsachBNNopTien_VP::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CrptDanhsachBNNopTien_VP::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CrptDanhsachBNNopTien_VP::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szObjectKey.Empty();
	m_szClerkKey.Empty();
	m_bOrderCashVoucher=FALSE;
	m_bLocked=FALSE;

}
int CrptDanhsachBNNopTien_VP::SetMode(int nMode){
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
/*void CrptDanhsachBNNopTien_VP::OnYearChange(){
	
} */
/*void CrptDanhsachBNNopTien_VP::OnYearSetfocus(){
	
} */
/*void CrptDanhsachBNNopTien_VP::OnYearKillfocus(){
	
} */
int CrptDanhsachBNNopTien_VP::OnYearCheckValue(){
	return 0;
} 
void CrptDanhsachBNNopTien_VP::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
} 
void CrptDanhsachBNNopTien_VP::OnReportPeriodSelendok(){
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
/*void CrptDanhsachBNNopTien_VP::OnReportPeriodSetfocus(){
	
}*/
/*void CrptDanhsachBNNopTien_VP::OnReportPeriodKillfocus(){
	
}*/
long CrptDanhsachBNNopTien_VP::OnReportPeriodLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
		CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0){
		szWhere.Format(_T(" WHERE hpr_idx = %ld"), ToInt(m_szReportPeriodKey));
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
/*void CrptDanhsachBNNopTien_VP::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CrptDanhsachBNNopTien_VP::OnFromDateChange(){
	
} */
/*void CrptDanhsachBNNopTien_VP::OnFromDateSetfocus(){
	
} */
/*void CrptDanhsachBNNopTien_VP::OnFromDateKillfocus(){
	
} */
int CrptDanhsachBNNopTien_VP::OnFromDateCheckValue(){
	return 0;
} 
/*void CrptDanhsachBNNopTien_VP::OnToDateChange(){
	
} */
/*void CrptDanhsachBNNopTien_VP::OnToDateSetfocus(){
	
} */
/*void CrptDanhsachBNNopTien_VP::OnToDateKillfocus(){
	
} */
int CrptDanhsachBNNopTien_VP::OnToDateCheckValue(){
	return 0;
} 
void CrptDanhsachBNNopTien_VP::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CrptDanhsachBNNopTien_VP::OnObjectSelendok(){
	 
}
/*void CrptDanhsachBNNopTien_VP::OnObjectSetfocus(){
	
}*/
/*void CrptDanhsachBNNopTien_VP::OnObjectKillfocus(){
	
}*/
long CrptDanhsachBNNopTien_VP::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty())
	{
		szWhere.Format(_T(" where ho_id='%s'"), m_szObjectKey);	
	};
	m_wndObject.DeleteAllItems(); 
	szSQL.Format(_T(" select ho_type as id, ho_desc as name from hms_object %s order by cast(ho_id as integer)"), szWhere);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CrptDanhsachBNNopTien_VP::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CrptDanhsachBNNopTien_VP::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CrptDanhsachBNNopTien_VP::OnClerkSelendok(){
	 
}
/*void CrptDanhsachBNNopTien_VP::OnClerkSetfocus(){
	
}*/
/*void CrptDanhsachBNNopTien_VP::OnClerkKillfocus(){
	
}*/
long CrptDanhsachBNNopTien_VP::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	};
	m_wndClerk.DeleteAllItems(); 
	szSQL.Format(_T("SELECT su_userid as userid, su_name as name FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("userid")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CrptDanhsachBNNopTien_VP::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
	void CrptDanhsachBNNopTien_VP::OnOrderCashVoucherSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
	void CrptDanhsachBNNopTien_VP::OnLockedSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CrptDanhsachBNNopTien_VP::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	if(!rpt.Init(_T("Reports/HMS/HF_DANHSACHBENHNHANNOPTIEN.RPT")) )
		return ;
	BeginWaitCursor();
	UpdateData(true);

	szSQL = GetQueryString();
	int nRes = rs.ExecSQL(szSQL);
	if (nRes < 0)
	{
		MessageBox(_T("Wrong SQL"), 0, MB_OK);
		return;
	}
	if (rs.IsEOF())
	{
		MessageBox(_T("No Data!"), 0, MB_ICONASTERISK);
		return;
	}

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	int nIndex = 1;
	CReportSection* rptDetail;
	double nTotal = 0; 
	double nGrpTotal = 0;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%d"), nIndex ++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		
		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		tmpStr.Format(_T("%s- %s"), rs.GetValue(_T("hfe_deptid")), rs.GetValue(_T("roomid")));
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("amount"), nTotal);
		nGrpTotal += nTotal;
		FormatCurrency(nTotal, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rs.MoveNext();	
	}	
	tmpStr.Format(_T("%.0f"), nGrpTotal);
	FormatCurrency(nGrpTotal, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("SumTotalAmount"), tmpStr);
	CString szMoney;
	MoneyToString(tmpStr, szMoney);
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), szMoney);
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();
	
} 
void CrptDanhsachBNNopTien_VP::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0, nIdx = 0;
	double nAmount = 0, nTotal = 0;
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 15);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 5);
	xls.SetCellMergedColumns(nCol, nRow + 3, 5);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, _T("\x42\x1EA2NG TH\x1ED0NG K\xCA \x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N N\x1ED8P TI\x1EC0N"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn \x62\x1EC7nh nh\xE2n"));
	arrCol.Add(_T("S\x1ED1 h\x1ED3 s\x1A1"));
	arrCol.Add(_T("Ph\xF2ng kh\xE1m"));
	arrCol.Add(_T("S\x1ED1 ti\x1EC1n"));
	nRow = 4;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10); 
	}
	nRow = 5;
	while (!rs.IsEOF()){
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER);
		tmpStr.Format(_T("%s- %s"), rs.GetValue(_T("hfe_deptid")), rs.GetValue(_T("roomid")));
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("amount"), nAmount);
		nTotal += nAmount;
		FormatCurrency(nAmount, tmpStr);
		xls.SetCellText(nCol+4, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
		nRow++;
		rs.MoveNext();
	}
	FormatCurrency(nTotal, tmpStr);
	xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT , true, 10);
	xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_RIGHT, true, 10);
	xls.Save(_T("Exports\\Bao Cao Su Dung Thuoc Dich Vu Ngoai Tru.xls"));
	
} 
void CrptDanhsachBNNopTien_VP::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CrptDanhsachBNNopTien_VP::OnAddrptDanhsachBNNopTien_VP(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CrptDanhsachBNNopTien_VP::OnEditrptDanhsachBNNopTien_VP(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CrptDanhsachBNNopTien_VP::OnDeleterptDanhsachBNNopTien_VP(){
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
 		OnCancelrptDanhsachBNNopTien_VP(); 
 	}
	return 0;
}
int CrptDanhsachBNNopTien_VP::OnSaverptDanhsachBNNopTien_VP(){
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
 		//OnrptDanhsachBNNopTien_VPListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CrptDanhsachBNNopTien_VP::OnCancelrptDanhsachBNNopTien_VP(){
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
int CrptDanhsachBNNopTien_VP::OnrptDanhsachBNNopTien_VPListLoadData(){
	return 0;
}

CString CrptDanhsachBNNopTien_VP::GetQueryString(){
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND (hfe_date) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szClerkKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hfe_staff= '%s'"), m_szClerkKey);
	}
	if (m_bLocked)
		szWhere.AppendFormat(_T(" AND hfe_locked = 'Y'"));
	
	szSQL.Format(_T(" SELECT trim(hp_surname") \
				_T("   ||' '") \
				_T("   ||hp_midname") \
				_T("   ||' '") \
				_T("   ||hp_firstname) AS pname,") \
				_T("   hd_docno        AS docno,") \
				_T("   hfe_deptid,") \
				_T("   (SELECT DISTINCT he_roomid") \
				_T("   FROM hms_exam") \
				_T("   WHERE he_docno  =hd_docno") \
				_T("   AND he_receptidx=hd_areceptidx") \
				_T("   )          AS roomid,") \
				_T("   hfe_amount AS amount") \
				_T(" FROM hms_patient") \
				_T(" LEFT JOIN hms_doc") \
				_T(" ON(hd_patientno=hp_patientno)") \
				_T(" LEFT JOIN hms_xfee_invoice") \
				_T(" ON(hd_docno=hfe_docno)") \
				_T(" WHERE hfe_status <>'C' %s") \
				_T(" ORDER BY hfe_invoiceno"), szWhere);

	return szSQL;
}