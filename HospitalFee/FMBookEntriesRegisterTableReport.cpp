#include "stdafx.h"
#include "FMBookEntriesRegisterTableReport.h"
#include "ReportDocument.h"
#include "Excel.h"
#include "HMSMainFrame.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((FMBookEntriesRegisterTableReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((FMBookEntriesRegisterTableReport * )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((FMBookEntriesRegisterTableReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((FMBookEntriesRegisterTableReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((FMBookEntriesRegisterTableReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((FMBookEntriesRegisterTableReport * )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((FMBookEntriesRegisterTableReport *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((FMBookEntriesRegisterTableReport *)pWnd)->OnClerkAddNew();
}*/

static long _OnDeptLoadDataFnc(CWnd *pWnd){
	return ((FMBookEntriesRegisterTableReport *)pWnd)->OnDeptLoadData();
}

static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((FMBookEntriesRegisterTableReport *)pWnd)->OnObjectLoadData();
}
static void _OnPrintSelectFnc(CWnd *pWnd){
	FMBookEntriesRegisterTableReport *pVw = (FMBookEntriesRegisterTableReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	FMBookEntriesRegisterTableReport *pVw = (FMBookEntriesRegisterTableReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	FMBookEntriesRegisterTableReport *pVw = (FMBookEntriesRegisterTableReport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddHMSGeneralCostReportDialogFnc(CWnd *pWnd){
	 return ((FMBookEntriesRegisterTableReport *)pWnd)->OnAddHMSGeneralCostReportDialog();
} 
static int _OnEditHMSGeneralCostReportDialogFnc(CWnd *pWnd){
	 return ((FMBookEntriesRegisterTableReport *)pWnd)->OnEditHMSGeneralCostReportDialog();
} 
static int _OnDeleteHMSGeneralCostReportDialogFnc(CWnd *pWnd){
	 return ((FMBookEntriesRegisterTableReport *)pWnd)->OnDeleteHMSGeneralCostReportDialog();
} 
static int _OnSaveHMSGeneralCostReportDialogFnc(CWnd *pWnd){
	 return ((FMBookEntriesRegisterTableReport *)pWnd)->OnSaveHMSGeneralCostReportDialog();
} 
static int _OnCancelHMSGeneralCostReportDialogFnc(CWnd *pWnd){
	 return ((FMBookEntriesRegisterTableReport *)pWnd)->OnCancelHMSGeneralCostReportDialog();
} 
FMBookEntriesRegisterTableReport ::FMBookEntriesRegisterTableReport (CWnd *pParent)
	{

	m_nDlgWidth = 445;
	m_nDlgHeight = 215;
	m_szTitle = _T("Thu \x63hi n\x1ED9i tr\xFA h\xE0ng ng\xE0y");
	SetDefaultValues();
}
FMBookEntriesRegisterTableReport ::~FMBookEntriesRegisterTableReport (){
}
void FMBookEntriesRegisterTableReport ::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 100, 55);
	m_wndYear.Create(this,102, 30, 217, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 228, 30, 318, 55);
	m_wndReportFilter.Create(this, _T("Report Condition"), 5, 5, 440, 180);
	m_wndReportPeriod.Create(this,323, 30, 438, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 100, 85);
	m_wndFromDate.Create(this,102, 60, 217, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 228, 60, 318, 85);
	m_wndToDate.Create(this,323, 60, 438, 85); 
	m_wndDeptLabel.Create(this, _T("Dept"), 10, 90, 100, 115);
	m_wndDept.Create(this,102, 90, 437, 115); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 120, 100, 145);
	m_wndObject.Create(this,102, 120, 437, 145); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 150, 100, 175);
	m_wndClerk.Create(this,102, 150, 439, 175); 
	m_wndPrint.Create(this, _T("&Print"), 255, 185, 345, 210);
	m_wndExport.Create(this, _T("&Export"), 350, 185, 440, 210);

}
void FMBookEntriesRegisterTableReport ::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	//m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	//m_wndToDate.SetCheckValue(true);
	////m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);

	m_wndReportPeriod.Format(2, 1, 17);
	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 100);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

	m_wndDept.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndDept.InsertColumn(1,_T("Description"), CFMT_TEXT,200);
	
	m_wndObject.InsertColumn(0, _T("ID"), CFMT_TEXT, 100);
	m_wndObject.InsertColumn(1,_T("Description"), CFMT_TEXT,200);

}
void FMBookEntriesRegisterTableReport ::OnSetWindowEvents(){
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

	m_wndDept.SetEvent(WE_LOADDATA, _OnDeptLoadDataFnc);

	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	SetMode(VM_EDIT);
}
void FMBookEntriesRegisterTableReport ::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_TextEx(pDX, m_wndDept.GetDlgCtrlID(), m_szDeptKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
}
void FMBookEntriesRegisterTableReport ::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void FMBookEntriesRegisterTableReport ::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void FMBookEntriesRegisterTableReport ::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();

}
int FMBookEntriesRegisterTableReport ::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, 1, 2, 3, -1); 
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
/*void FMBookEntriesRegisterTableReport ::OnYearChange(){
	
} */
/*void FMBookEntriesRegisterTableReport ::OnYearSetfocus(){
	
} */
/*void FMBookEntriesRegisterTableReport ::OnYearKillfocus(){
	
} */
int FMBookEntriesRegisterTableReport ::OnYearCheckValue(){
	return 0;
} 
void FMBookEntriesRegisterTableReport ::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void FMBookEntriesRegisterTableReport ::OnReportPeriodSelendok(){
	
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
/*void FMBookEntriesRegisterTableReport ::OnReportPeriodSetfocus(){
	
}*/
/*void FMBookEntriesRegisterTableReport ::OnReportPeriodKillfocus(){
	
}*/
long FMBookEntriesRegisterTableReport::OnReportPeriodLoadData(){
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
/*void FMBookEntriesRegisterTableReport ::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void FMBookEntriesRegisterTableReport ::OnFromDateChange(){
	
} */
/*void FMBookEntriesRegisterTableReport ::OnFromDateSetfocus(){
	
} */
/*void FMBookEntriesRegisterTableReport ::OnFromDateKillfocus(){
	
} */
int FMBookEntriesRegisterTableReport ::OnFromDateCheckValue(){
	return 0;
} 
/*void FMBookEntriesRegisterTableReport ::OnToDateChange(){
	
} */
/*void FMBookEntriesRegisterTableReport ::OnToDateSetfocus(){
	
} */
/*void FMBookEntriesRegisterTableReport ::OnToDateKillfocus(){
	
} */
int FMBookEntriesRegisterTableReport ::OnToDateCheckValue(){
	return 0;
} 
void FMBookEntriesRegisterTableReport ::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void FMBookEntriesRegisterTableReport ::OnClerkSelendok(){
	 
}
/*void FMBookEntriesRegisterTableReport ::OnClerkSetfocus(){
	
}*/
/*void FMBookEntriesRegisterTableReport ::OnClerkKillfocus(){
	
}*/
long FMBookEntriesRegisterTableReport ::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty()){
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	};
	m_wndClerk.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as userid, su_name as name FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);
		
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("userid")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void FMBookEntriesRegisterTableReport ::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */

long FMBookEntriesRegisterTableReport ::OnDeptLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDept.IsSearchKey() && !m_szDeptKey.IsEmpty()){
		szWhere.Format(_T(" and sd_id ='%s' "), m_szDeptKey);
	};
	m_wndDept.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT sd_id as id, sd_name as name FROM sys_dept WHERE sd_type='DT' %s ORDER BY sd_id"), szWhere);
		
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDept.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}

long FMBookEntriesRegisterTableReport ::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty()){
		szWhere.Format(_T(" where ho_id =%d "), ToInt(m_szObjectKey));
	};
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" select ho_type as id, ho_desc as name from hms_object %s order by cast(ho_id as integer)"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
void FMBookEntriesRegisterTableReport::OnPrintSelect(){	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	if(!rpt.Init(_T("Reports/HMS/HF_BANGKECHUNGTUGHISO.RPT")) )
	return ;
	UpdateData(true);
// viet cau lenh
	_fmsg(_T("%s"), szSQL);

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm,ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	long double nTotal[8];
	for (int i = 1; i<= 7; i++)
	{
		nTotal[i] = 0.0;
	}
	int nIndex = 1;
		CReportSection* rptDetail;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%ld"), nIndex ++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T(""),tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		
		rs.GetValue(_T(""), tmpStr);
		nTotal[1] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		
		rs.GetValue(_T(""), tmpStr);
		nTotal[2] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		nTotal[3] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		nTotal[4] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		nTotal[5] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		nTotal[6] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.GetValue(_T(""), tmpStr);
		nTotal[7] += ToDouble(tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.MoveNext();	
	}
	for (int i = 1; i<= 7; i++)
	{
	
			CString  szField;
			szField.Format(_T("s%d"), i+6);
			tmpStr.Format(_T("%.0f"), nTotal[i]);
			rpt.GetReportFooter()->SetValue(szField, tmpStr);		
	}
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
} 
void FMBookEntriesRegisterTableReport ::OnExportSelect(){	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	UpdateData(true);
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 26);
	xls.SetColumnWidth(2, 13);

	xls.SetColumnWidth(3, 21);
	xls.SetColumnWidth(4, 21);
	xls.SetColumnWidth(5, 21);
	xls.SetColumnWidth(6, 21);
	xls.SetColumnWidth(7, 21);
	xls.SetColumnWidth(8, 21);
	xls.SetColumnWidth(9, 21);
	xls.SetColumnWidth(10, 21);
	xls.SetColumnWidth(11, 21);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);
	xls.SetCellText(nCol, nRow + 1,pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 2, _T("Ph\xF2ng t\xE0i \x63h\xEDnh"), FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellMergedColumns(nCol + 4, nRow + 1, 6);
	xls.SetCellText(nCol + 4, nRow + 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellMergedColumns(nCol + 4, nRow + 2, 6);
	xls.SetCellText(nCol + 4, nRow + 2, _T("\x110\x1ED9\x63 l\x1EADp - T\x1EF1 \x64o - H\x1EA1nh ph\xFA\x63"), FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellMergedColumns(nCol, nRow + 3, 11);
	xls.SetCellText(nCol, nRow + 3, _T("\x42\x1EA2NG K\xCA \x43HI \x42\x1EC6NH NH\xC2N \x44\x1ECA\x43H V\x1EE4"), FMT_TEXT | FMT_CENTER, true, 14);

	xls.SetCellMergedColumns(nCol, nRow + 4, 11);
	xls.SetCellText(nCol, nRow + 4, _T("\x43hi tr\x1EA3 \x62\x1EC7nh nh\xE2n \x64\x1ECB\x63h v\x1EE5 y t\x1EBF"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedColumns(nCol, nRow + 5, 11);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol, nRow + 7, _T("STT"), FMT_TEXT | FMT_CENTER, true, 10, 1);
	xls.SetCellText(nCol + 1, nRow + 7, _T("H\x1ECD v\xE0 t\xEAn"), FMT_TEXT | FMT_CENTER, true, 10, 1);

	xls.SetCellText(nCol + 2, nRow + 7, _T("S\x1ED1 \x62\x1EC7nh \xE1n"), FMT_TEXT | FMT_CENTER, true, 10, 1);
	xls.SetCellText(nCol + 3, nRow + 7, _T("Kho\x61"), FMT_TEXT | FMT_CENTER, true, 10, 1);
	xls.SetCellText(nCol + 4, nRow + 7, _T("Thu \x58N ngo\x1EA1i tr\xFA"), FMT_TEXT | FMT_CENTER, true, 10, 1);

	xls.SetCellText(nCol + 5, nRow + 7, _T("\x43\xE1\x63 kho\x1EA3n ph\x1EA3i thu"), FMT_TEXT | FMT_CENTER, true, 10, 1);
	xls.SetCellText(nCol + 6, nRow + 7, _T("T\x1EA1m g\x1EEDi"), FMT_TEXT | FMT_CENTER, true, 10, 1);

	xls.SetCellText(nCol + 7, nRow + 7, _T("Vi\x1EC7n ph\xED ph\x1EA3i thu"), FMT_TEXT | FMT_CENTER, true, 10, 1);
	xls.SetCellText(nCol + 8, nRow + 7, _T("Tr\xED\x63h t\x1EA1m g\x1EEDi"), FMT_TEXT | FMT_CENTER, true, 10, 1);

	xls.SetCellText(nCol + 9, nRow + 7, _T("\x43hi tr\x1EA3"), FMT_TEXT | FMT_CENTER, true, 10, 1);
	xls.SetCellText(nCol + 10, nRow + 7, _T("T\x1ED5ng \x63hi"), FMT_TEXT | FMT_CENTER, true, 10, 1);
	
	szWhere.Empty();

	int nIndex = 1;
	long double nTotal[8];
	for(int i = 1 ; i<= 7; i++)
	{
		nTotal[i] = 0.0;
	}
	while(!rs.IsEOF())
	{
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow + 6, tmpStr, FMT_INTEGER);

		rs.GetValue(_T(""),tmpStr);
		xls.SetCellText(nCol + 1, nRow + 6, tmpStr, FMT_TEXT);
 
		rs.GetValue(_T(""),tmpStr);
		xls.SetCellText(nCol + 2, nRow + 6, tmpStr, FMT_TEXT);

		rs.GetValue(_T(""), tmpStr);
		xls.SetCellText(nCol + 3, nRow + 6, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T(""),tmpStr);
		nTotal[1] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 4, nRow + 6, tmpStr, FMT_CURRENCY1);
		
		rs.GetValue(_T(""), tmpStr);
		nTotal[2] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 5, nRow + 6, tmpStr, FMT_CURRENCY1);

		rs.GetValue(_T("Dept"), tmpStr);
		nTotal[3] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 6, nRow + 6, tmpStr, FMT_CURRENCY1);
		
		rs.GetValue(_T(""), tmpStr);
		nTotal[4] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 7, nRow + 6, tmpStr, FMT_CURRENCY1);

		rs.GetValue(_T(""), tmpStr);
		nTotal[5] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 8, nRow + 6, tmpStr, FMT_CURRENCY1);

		rs.GetValue(_T(""), tmpStr);
		nTotal[6] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 9, nRow + 6, tmpStr, FMT_CURRENCY1);

		rs.GetValue(_T(""), tmpStr);
		nTotal[7] += ToDouble(tmpStr);
		xls.SetCellText(nCol + 10, nRow + 6, tmpStr, FMT_CURRENCY1);
		nRow++;
		rs.MoveNext();
			
	}
	xls.SetCellMergedColumns(nCol, nRow + 6, 6);
	xls.SetCellText(nCol, nRow + 6, _T("T\x1ED5ng"), FMT_TEXT | FMT_CENTER, true, 11);
	for (int i = 1; i<= 7; i++)
	{
		tmpStr.Format(_T("%.0f"), nTotal[i]);
		xls.SetCellText(nCol, nRow + 6, tmpStr, FMT_CURRENCY1, true, 11);
	}
	CString szDate;
	xls.Save(_T("Exports\\BANG KE THU CHUNG TU GHI SO.xls"));
} 
void FMBookEntriesRegisterTableReport ::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int FMBookEntriesRegisterTableReport ::OnAddHMSGeneralCostReportDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int FMBookEntriesRegisterTableReport ::OnEditHMSGeneralCostReportDialog(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int FMBookEntriesRegisterTableReport ::OnDeleteHMSGeneralCostReportDialog(){
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
 		OnCancelHMSGeneralCostReportDialog(); 
 	}
	return 0;
}
int FMBookEntriesRegisterTableReport ::OnSaveHMSGeneralCostReportDialog(){
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
 		//OnHMSGeneralCostReportDialogListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int FMBookEntriesRegisterTableReport ::OnCancelHMSGeneralCostReportDialog(){
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
int FMBookEntriesRegisterTableReport ::OnHMSGeneralCostReportDialogListLoadData(){
	return 0;
}
CString FMBookEntriesRegisterTableReport::GetQueryString()
{
	CString tmpStr, szSQL, szWhere,szUserReceiver,szFromDateToDate, szTmp, szObject;
	szWhere.Empty();	
	
	if (!m_wndClerk.GetCurrent(0).IsEmpty())
		szUserReceiver.Format(_T(" and hfi_receiver ='%s' "), m_wndClerk.GetCurrent(0));
	
	if (!m_szDeptKey.IsEmpty())
		szWhere.Format(_T(" and hfe_deptid ='%s' "), m_szDeptKey);

	if (!m_szObjectKey.IsEmpty())
		szObject.AppendFormat(_T(" and ho_type ='%s' "), m_szObjectKey);
	

	szSQL.Format(_T(" SELECT ") \
				_T(" 	max(hfi_invoiceno) as invoiceno,") \
				_T(" 	date(hfi_recvdate) as ngaythuphi,") \
				_T(" 	hfe_deptid as deptid,") \
				_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as c1,") \
				_T(" 	hd_docno as c2,") \
				_T(" 	hcr_recordno as c3,") \
				_T(" 	(select distinct sd_name from sys_dept where sd_id=hfe_deptid) as c4,") \
				_T(" 	(select distinct ho_desc from hms_object where ho_id = hd_object) as c5,") \
				_T(" 	1 as c6,") \
				_T(" 	(select distinct sum(hfi_patpaid) from hms_fee_invoice where (hfi_type ='A' and hfi_category ='A' and hfi_status='O') and (hfi_docno=hd_docno and hfe_deptid=hfi_deptid)) as c7, ") \
				_T(" 	bedqty as c8,") \
				_T(" 	sum(examfee) as c9,") \
				_T(" 	sum(bedfee) as c10,") \
				_T(" 	sum(drugfee) as c11,") \
				_T(" 	sum(materialfee) as c12,") \
				_T(" 	sum(XNHH)as c13,") \
				_T(" 	sum(XNsinhhoa) as c14, ") \
				_T(" 	sum(XNViSinh) as c15,") \
				_T(" 	sum(XNGPB) as c16,") \
				_T(" 	sum(CDHASAdentrang) as c17,") \
				_T(" 	sum(CDHASAMau) as c18,") \
				_T(" 	sum(XQuang) as c19,") \
				_T(" 	sum(CtyScan) as c20,") \
				_T(" 	sum(CDHAdientim) as c21,") \
				_T(" 	sum(CDHAdiennao) as c22,") \
				_T(" 	sum(CDHAluuhuyetnao) as c23,") \
				_T(" 	sum(CDHAnoisoi) as c24,") \
				_T(" 	sum(CDHAHohap) as c25,") \
				_T(" 	sum(bloodfee) as c26,") \
				_T(" 	sum(PTTT) as c27,") \
				_T(" 	sum(BHcungchitra) as c28,") \
				_T(" 	sum(otherfee) as c29,") \
				_T(" 	(select distinct sum(hfi_patpaid) from hms_fee_invoice where hfi_type ='D' and (hfi_docno=hd_docno and hfe_deptid=hfi_deptid)) as c30, ") \
				_T(" 	sum(cost) as c31,") \
				_T(" 	(select distinct sum(hfi_patpaid) from hms_fee_invoice where hfi_type ='R' and (hfi_docno=hd_docno and hfe_deptid=hfi_deptid)) as c32,") \
				_T(" 	(select distinct sum(hfi_patpaid) from hms_fee_invoice where (hfi_type ='A' and hfi_category ='D' and hfi_status='P') and (hfi_docno=hd_docno and hfe_deptid=hfi_deptid)) as c33, ") \
				_T(" 	hfi_bookno  as c34, ") \
				_T(" 	hfi_recvno as c35,") \
				_T(" 	(select distinct sum(hfi_patpaid) from hms_fee_invoice where (hfi_type = 'P' and hfi_type not in('A','D')) and (hfi_docno=hd_docno and hfe_deptid=hfi_deptid)) as Thutructiep") \
				_T("    FROM ") \
				_T("    (") \
				_T("        SELECT ") \
				_T(" 	hfe_deptid, hfi_recvdate,") \
				_T(" 	hfi_class,") \
				_T(" 	hfi_docno,") \
				_T(" 	(select distinct  sum(hb_treatqty) from hms_bed  where hb_docno = hfe_docno and hb_deptid=hfe_deptid) as bedqty, ") \
				_T(" 	hfi_invoiceno,") \
				_T(" 	1 as soluot,") \
				_T(" 	hfi_advanceamt as tamungvaovien,") \
				_T(" 	case when hfe_group='D0000' then hfe_cost-hfe_difcost else 0 end as examfee,") \
				_T(" 	case when hfe_group ='C0000' then hfe_cost-hfe_difcost else 0 end as bedfee,	") \
				_T(" 	case when substr(hfe_group, 1, 2)='A1' and hfe_group <>'A1400' then hfe_cost-hfe_difcost else 0 end as drugfee,") \
				_T(" 	case when substr(hfe_group, 1, 2)='A9' then hfe_cost-hfe_difcost else 0 end as materialfee,") \
				_T(" 	case when hfe_group in ('B1100', 'B1700') then hfe_cost-hfe_difcost else 0 end as XNHH,") \
				_T(" 	case when hfe_group in ('B1200', 'B1300') then hfe_cost-hfe_difcost else 0 end as XNsinhhoa,	") \
				_T(" 	case when hfe_group = 'B1100' then hfe_cost-hfe_difcost else 0 end as XNViSinh,") \
				_T(" 	case when hfe_group = 'B1900' then hfe_cost-hfe_difcost else 0 end as XNGPB,") \
				_T(" 	case when substr(hfe_group, 1, 3)='B24' and hfe_hitech='N' then hfe_cost-hfe_difcost else 0 end as CDHASAdentrang,") \
				_T(" 	case when substr(hfe_group, 1, 3)='B25' and hfe_hitech='N' then hfe_cost-hfe_difcost else 0 end as CDHASAMau,") \
				_T(" 	case when substr(hfe_group, 1, 3)='B21' and hfe_hitech='N' then hfe_cost-hfe_difcost else 0 end as XQuang,") \
				_T(" 	case when substr(hfe_group, 1, 3)='B22' and hfe_hitech='N' then hfe_cost-hfe_difcost else 0 end as CtyScan,") \
				_T(" 	case when substr(hfe_group, 1, 3)='B33' and hfe_hitech='N' then hfe_cost-hfe_difcost else 0 end as CDHAdientim,") \
				_T(" 	case when substr(hfe_group, 1, 3)='B34' and hfe_hitech='N' then hfe_cost-hfe_difcost else 0 end as CDHAdiennao,") \
				_T(" 	case when substr(hfe_group, 1, 3)='B36' and hfe_hitech='N' then hfe_cost-hfe_difcost else 0 end as CDHAluuhuyetnao,") \
				_T(" 	case when substr(hfe_group, 1, 3)='B31' and hfe_hitech='N' then hfe_cost-hfe_difcost else 0 end as CDHAnoisoi,") \
				_T(" 	case when substr(hfe_group, 1, 3)='B35' and hfe_hitech='N' then hfe_cost-hfe_difcost else 0 end as CDHAHohap,") \
				_T(" 	case when substr(hfe_group, 1, 2) = 'A2' then hfe_cost-hfe_difcost else 0 end as bloodfee,") \
				_T(" 	case when substr(hfe_group, 1, 2) in('B4','B5') and hfe_hitech='N' then hfe_cost-hfe_difcost else 0 end as PTTT,") \
				_T(" 	hfe_inspaid as BHcungchitra,") \
				_T(" 	case when hfe_group = 'F0000' then hfe_cost - hfe_difcost else 0 end as otherfee,") \
				_T(" 	hfe_discount as discount,") \
				_T(" 	hfe_cost as cost,") \
				_T(" 	hfi_refundamt as hoanung,") \
				_T(" 	hfi_advpayment as tamung,") \
				_T(" 	hfi_bookno, ") \
				_T(" 	hfi_recvno") \
				_T("   FROM hmsv_fee") \
				_T("   LEFT JOIN hms_fee_invoice ON(hfi_docno=hfe_docno AND hfe_invoiceno=hfi_invoiceno)") \
				_T("   WHERE hfe_status='P' ") \
				_T(" 	and hfe_type not in('E')") \
				_T(" 	and date(hfi_recvdate) between date('%s') and date('%s')") \
				_T(" 	and hfi_type='P'") \
				_T(" 	and hfi_class in('I','O') %s %s") \
				_T("   ) as tbl ") \
				_T("   LEFT JOIN hms_doc ON(hfi_docno=hd_docno)	") \
				_T("   LEFT JOIN hms_patient ON(hd_patientno=hp_patientno) ") \
				_T("   LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno) ") \
				_T("   LEFT JOIN hms_object ON(ho_id=hd_object) ") \
				_T("   WHERE hcr_status ='T' %s ") \
				_T("   GROUP BY c1, c2, deptid, ho_type, c3, ngaythuphi, c34, c35, c8, c5") \
				_T("   HAVING sum(cost) > 0") \
				_T("   ORDER BY c1, c2"), m_szFromDate, m_szToDate, szUserReceiver,szWhere, szObject);
		return szSQL;
}
