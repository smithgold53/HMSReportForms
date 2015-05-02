#include "stdafx.h"
#include "HMSMainFrame.h"
#include "TMReportsDeptPatientList.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CTMReportsDeptPatientList *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMReportsDeptPatientList* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CTMReportsDeptPatientList *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CTMReportsDeptPatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CTMReportsDeptPatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMReportsDeptPatientList* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CTMReportsDeptPatientList *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnObjectAddNew();
}*/
static void _OnRankSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMReportsDeptPatientList* )pWnd)->OnRankSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnRankSelendokFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnRankSelendok();
}
/*static void _OnRankSetfocusFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnRankKillfocus();
}*/
/*static void _OnRankKillfocusFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnRankKillfocus();
}*/
static long _OnRankLoadDataFnc(CWnd *pWnd){
	return ((CTMReportsDeptPatientList *)pWnd)->OnRankLoadData();
}
/*static void _OnRankAddNewFnc(CWnd *pWnd){
	((CTMReportsDeptPatientList *)pWnd)->OnRankAddNew();
}*/
static void _OnAllSelectFnc(CWnd *pWnd){
	  ((CTMReportsDeptPatientList*)pWnd)->OnAllSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	  ((CTMReportsDeptPatientList*)pWnd)->OnInPatientSelect();
}
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	  ((CTMReportsDeptPatientList*)pWnd)->OnOutPatientSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CTMReportsDeptPatientList *pVw = (CTMReportsDeptPatientList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CTMReportsDeptPatientList *pVw = (CTMReportsDeptPatientList *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddTMReportsDeptPatientListFnc(CWnd *pWnd){
	 return ((CTMReportsDeptPatientList*)pWnd)->OnAddTMReportsDeptPatientList();
} 
static int _OnEditTMReportsDeptPatientListFnc(CWnd *pWnd){
	 return ((CTMReportsDeptPatientList*)pWnd)->OnEditTMReportsDeptPatientList();
} 
static int _OnDeleteTMReportsDeptPatientListFnc(CWnd *pWnd){
	 return ((CTMReportsDeptPatientList*)pWnd)->OnDeleteTMReportsDeptPatientList();
} 
static int _OnSaveTMReportsDeptPatientListFnc(CWnd *pWnd){
	 return ((CTMReportsDeptPatientList*)pWnd)->OnSaveTMReportsDeptPatientList();
} 
static int _OnCancelTMReportsDeptPatientListFnc(CWnd *pWnd){
	 return ((CTMReportsDeptPatientList*)pWnd)->OnCancelTMReportsDeptPatientList();
} 
CTMReportsDeptPatientList::CTMReportsDeptPatientList(CWnd *pWnd){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CTMReportsDeptPatientList::~CTMReportsDeptPatientList(){
}
void CTMReportsDeptPatientList::OnCreateComponents(){
	m_wndFromDateLabel.Create(this, _T("Ng\xE0y \x111i\x1EC1u tr\x1ECB :"), 10, 30, 90, 55);
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 90);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndObjectLab.Create(this, _T("Object"), 10, 60, 90, 85);
	m_wndObject.Create(this,95, 60, 215, 85); 
	m_wndRankLab.Create(this, _T("Rank"), 220, 60, 300, 85);
	m_wndRank.Create(this,305, 60, 425, 85); 
	m_wndAll.Create(this, _T("All"), 6, 95, 86, 120);
	m_wndInPatient.Create(this, _T("InPatient"), 90, 95, 170, 120);
	m_wndOutPatient.Create(this, _T("OutPatient"), 174, 95, 254, 120);
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 235, 124, 345, 149);
	m_wndExport.Create(this, _T("&Export"), 350, 124, 425, 149);

	/*m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);*/

}
void CTMReportsDeptPatientList::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	/*m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);*/
	/*m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);*/
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	/*m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);*/
	m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(35);
	m_wndRank.SetCheckValue(true);
	m_wndRank.LimitText(35);




	m_wndObject.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);


	m_wndRank.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndRank.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CTMReportsDeptPatientList::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndYear.SetEvent(WE_CHANGE, _OnYearChangeFnc);
	//m_wndYear.SetEvent(WE_SETFOCUS, _OnYearSetfocusFnc);
	//m_wndYear.SetEvent(WE_KILLFOCUS, _OnYearKillfocusFnc);
	//m_wndYear.SetEvent(WE_CHECKVALUE, _OnYearCheckValueFnc);
	//m_wndReportPeriod.SetEvent(WE_SELENDOK, _OnReportPeriodSelendokFnc);
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
	//m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndRank.SetEvent(WE_SELENDOK, _OnRankSelendokFnc);
	//m_wndRank.SetEvent(WE_SETFOCUS, _OnRankSetfocusFnc);
	//m_wndRank.SetEvent(WE_KILLFOCUS, _OnRankKillfocusFnc);
	m_wndRank.SetEvent(WE_SELCHANGE, _OnRankSelectChangeFnc);
	m_wndRank.SetEvent(WE_LOADDATA, _OnRankLoadDataFnc);
	//m_wndRank.SetEvent(WE_ADDNEW, _OnRankAddNewFnc);
	m_wndAll.SetEvent(WE_CLICK, _OnAllSelectFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	
	/*CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));*/
	m_szFromDate = pMF->GetSysDate();
	m_szFromDate += _T("23:59");
	//m_szToDate += _T("23:59");
	UpdateData(FALSE);
	m_wndAll.SetCheck(true);
	m_szCondition = _T("A");

}
void CTMReportsDeptPatientList::OnDoDataExchange(CDataExchange* pDX){
	//DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	//DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	//DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_TextEx(pDX, m_wndRank.GetDlgCtrlID(), m_szRankKey);
	DDX_Radio(pDX, m_wndAll.GetDlgCtrlID(), m_nAll);
	DDX_Radio(pDX, m_wndInPatient.GetDlgCtrlID(), m_nInPatient);
	DDX_Radio(pDX, m_wndOutPatient.GetDlgCtrlID(), m_nOutPatient);

}
void CTMReportsDeptPatientList::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CTMReportsDeptPatientList::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CTMReportsDeptPatientList::SetDefaultValues(){

	//m_nYear=0;
	//m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	//m_szToDate.Empty();
	m_szObjectKey.Empty();
	m_szRankKey.Empty();
	m_nAll=0;
	m_nInPatient=1;
	m_nOutPatient=1;

}
int CTMReportsDeptPatientList::SetMode(int nMode){
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
/*void CTMReportsDeptPatientList::OnYearChange(){
	
} */
/*void CTMReportsDeptPatientList::OnYearSetfocus(){
	
} */
/*void CTMReportsDeptPatientList::OnYearKillfocus(){
	
} */
int CTMReportsDeptPatientList::OnYearCheckValue(){
	return 0;
} 
void CTMReportsDeptPatientList::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMReportsDeptPatientList::OnReportPeriodSelendok(){
	/*CString tmpStr;
	CDate date;
	UpdateData(true);
	int nMonth = ToInt(m_szReportPeriodKey);
	if(nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/01"), m_nYear, nMonth);
		date.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), m_nYear, nMonth, date.GetMonthLastDay());
	}
	if(nMonth == 13){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/03/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59" ), m_nYear,  date.GetMonthLastDay());
	}
	if(nMonth == 14){
		m_szFromDate.Format(_T("%.4d/04/01"), m_nYear);
		tmpStr.Format(_T("%.4d/06/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), m_nYear, date.GetMonthLastDay());
	}
	if(nMonth == 15){
		m_szFromDate.Format(_T("%.4d/07/01"), m_nYear);
		tmpStr.Format(_T("%.4d/09/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), m_nYear, date.GetMonthLastDay());
	}
	if(nMonth == 16){
		m_szFromDate.Format(_T("%.4d/10/01"), m_nYear);
		tmpStr.Format(_T("%.4d/10/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, date.GetMonthLastDay());
	}
	if(nMonth == 17){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/12/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, date.GetMonthLastDay());		
	}	
	UpdateData(false);*/
}
/*void CTMReportsDeptPatientList::OnReportPeriodSetfocus(){
	
}*/
/*void CTMReportsDeptPatientList::OnReportPeriodKillfocus(){
	
}*/
long CTMReportsDeptPatientList::OnReportPeriodLoadData(){
	/*CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
	return nCount;*/
	return 0;
}
/*void CTMReportsDeptPatientList::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CTMReportsDeptPatientList::OnFromDateChange(){
	
} */
/*void CTMReportsDeptPatientList::OnFromDateSetfocus(){
	
} */
/*void CTMReportsDeptPatientList::OnFromDateKillfocus(){
	
} */
int CTMReportsDeptPatientList::OnFromDateCheckValue(){
	return 0;
} 
/*void CTMReportsDeptPatientList::OnToDateChange(){
	
} */
/*void CTMReportsDeptPatientList::OnToDateSetfocus(){
	
} */
/*void CTMReportsDeptPatientList::OnToDateKillfocus(){
	
} */
int CTMReportsDeptPatientList::OnToDateCheckValue(){
	return 0;
} 
void CTMReportsDeptPatientList::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMReportsDeptPatientList::OnObjectSelendok(){
	 
}
/*void CTMReportsDeptPatientList::OnObjectSetfocus(){
	
}*/
/*void CTMReportsDeptPatientList::OnObjectKillfocus(){
	
}*/
long CTMReportsDeptPatientList::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty())
	{
		//szWhere.Format();
	};
	szSQL.Format(_T("SELECT ho_id as id, ho_desc as description FROM hms_object ORDER BY cast(id as integer)"));
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CTMReportsDeptPatientList::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CTMReportsDeptPatientList::OnRankSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMReportsDeptPatientList::OnRankSelendok(){
	 
}
/*void CTMReportsDeptPatientList::OnRankSetfocus(){
	
}*/
/*void CTMReportsDeptPatientList::OnRankKillfocus(){
	
}*/
long CTMReportsDeptPatientList::OnRankLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndRank.IsSearchKey() && !m_szRankKey.IsEmpty()){
		szWhere.Format(_T("and sd_id = '%s'"), pMF->m_szDept);
	};
	m_wndRank.DeleteAllItems(); 
	szSQL.Format(_T("SELECT ss_code id,ss_desc name FROM sys_sel WHERE ss_id='hms_rank' and ss_code >12"), szWhere);
	_fmsg(_T("%s"),szSQL);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndRank.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CTMReportsDeptPatientList::OnRankAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CTMReportsDeptPatientList::OnAllSelect(){
	m_szCondition = _T("A");
}
void CTMReportsDeptPatientList::OnInPatientSelect(){
	m_szCondition = _T("I");
}
void CTMReportsDeptPatientList::OnOutPatientSelect(){
	m_szCondition = _T("E");
}
void CTMReportsDeptPatientList::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr,szOldGroup,szNewGroup;
	int ret = 0 ;
	szSQL = GetQueryString();
	if (!rpt.Init(_T("Reports\\HMS\\HE_DANHSACHBENHNHANDANGDIEUTRI.RPT")))
		return;
	ret = rs.ExecSQL(szSQL);
	if (ret <=0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HealthService"), _T("\x42\x1ED8 QU\x1ED0\x43 PH\xD2NG"));
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), _T("\x42\x1EC6NH VI\x1EC6N TW QU\xC2N \x110\x1ED8I \x31\x30\x38"));
	tmpStr.Format(_T("KHOA %s"),pMF->m_szDept);
	rpt.GetReportHeader()->SetValue(_T("khoa"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection *rptDetail = NULL;
	int nIndex = 1;
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("hovaten"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("age"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("capbac"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("HP_WORKPLACE"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("HD_DIAGNOSTIC"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("DOITUONG"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("htr_deptid"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.MoveNext();
	}
	CString szSysDate;
	szSysDate = pMF->GetSysDate(); 
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CTMReportsDeptPatientList::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);	CString szSQL, tmpStr, szTemp, szWhere, szNewGroup , szOldGroup;	UpdateData(TRUE);	BeginWaitCursor();	szWhere.Empty();	int nRow =0, nCol = 0;	CExcel xls;	xls.CreateSheet(1);	xls.SetWorksheet(0);	xls.SetColumnWidth(0, 8);	xls.SetColumnWidth(1, 23);	xls.SetColumnWidth(2, 15);	xls.SetColumnWidth(3, 16);	xls.SetColumnWidth(4, 13);	xls.SetColumnWidth(5, 33);	xls.SetColumnWidth(6, 15);	xls.SetColumnWidth(7, 15);	xls.SetCellText(0, 0, _T("\x42\x1ED8 QU\x1ED0\x43 PH\xD2NG"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(0, 1, _T("\x42\x1EC6NH VI\x1EC6N TW QU\xC2N \x110\x1ED8I \x31\x30\x38"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(0, 3, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N N\x1EB0M \x110I\x1EC0U TR\x1ECA T\x1EA0I VI\x1EC6N"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	tmpStr.Format(_T("Ng\xE0y \x111i\x1EC1u tr\x1ECB %s "),CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss),CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));	xls.SetCellText(0, 4, tmpStr, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(0, 6, _T("STT"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(1, 6, _T("T\xEAn \x42\x1EC7nh Nh\xE2n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(2, 6, _T("N\x103m sinh"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(3, 6, _T("\x43\x1EA5p \x62\x1EAD\x63"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(4, 6, _T("\x110\x1A1n v\x1ECB"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(5, 6, _T("\x43h\x1EA5n \x111o\xE1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(6, 6, _T("\x110\x1ED1i t\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(7, 6, _T("V\xE0o kho\x61"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetMerge(0, 0, 0, 2);	xls.SetMerge(1, 1, 0, 2);	xls.SetMerge(3, 3, 0, 7);	xls.SetMerge(4, 4, 0, 7);	szSQL = GetQueryString();	rs.ExecSQL(szSQL);	nRow = 7;	nCol = 0;	int nIdx=1;	while(!rs.IsEOF()){		tmpStr.Format(_T("%d"), nIdx++);		xls.SetCellText(nCol+0, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("hovaten"), tmpStr);		xls.SetCellText(nCol+1, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("age"), tmpStr);		xls.SetCellText(nCol+2, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("capbac"), tmpStr);		xls.SetCellText(nCol+3, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("HP_WORKPLACE"), tmpStr);		xls.SetCellText(nCol+4, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("HD_DIAGNOSTIC"), tmpStr);		xls.SetCellText(nCol+5, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("DOITUONG"), tmpStr);		xls.SetCellText(nCol+6, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("htr_deptid"), tmpStr);		xls.SetCellText(nCol+7, nRow, tmpStr, FMT_TEXT);		nRow++;		rs.MoveNext();	}	EndWaitCursor();	xls.Save(_T("Exports\\DANH SACH BN NAM DIEU TRI TAI VIEN.xls"));
} 
int CTMReportsDeptPatientList::OnAddTMReportsDeptPatientList(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CTMReportsDeptPatientList::OnEditTMReportsDeptPatientList(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CTMReportsDeptPatientList::OnDeleteTMReportsDeptPatientList(){
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
 		OnCancelTMReportsDeptPatientList();
 	}
	return 0;
}
int CTMReportsDeptPatientList::OnSaveTMReportsDeptPatientList(){
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
 		//OnTMReportsDeptPatientListListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CTMReportsDeptPatientList::OnCancelTMReportsDeptPatientList(){
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
int CTMReportsDeptPatientList::OnTMReportsDeptPatientListListLoadData(){
	return 0;
}
CString CTMReportsDeptPatientList::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere, szDept, szObject, szRank,szDept1;
	
		//Trang thai dieu tri toan vien
		if (m_szCondition == _T("A"))
			szWhere.AppendFormat(_T(" AND 1=1"));

		//Trang thai dieu tri noi tru
		if (m_szCondition == _T("I"))
			szWhere.AppendFormat(_T(" AND htr_OutPatient='N'"));

		//Trang thai dieu tri ngoai tru
		if (m_szCondition == _T("E"))
			szWhere.AppendFormat(_T(" AND htr_OutPatient='Y'"));

	if (!m_szObjectKey.IsEmpty())
		szObject.AppendFormat(_T(" AND hd_object = '%s'"), m_szObjectKey);

	if (!m_szRankKey.IsEmpty())
		szRank.AppendFormat(_T(" AND hp_rank = '%s'"), m_szRankKey);
	/*szSQL.Format(_T(" SELECT * ") \
	_T(" FROM(SELECT 1 typeid,hd_docno,") \
	_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) AS hovaten,") \
	_T("        extract(YEAR FROM hp_birthdate)       AS age,") \
	_T("        (SELECT ss_desc FROM sys_sel where ss_id='hms_rank' and ss_code=hp_rank) as capbac,") \
	_T("        hp_workplace,") \
	_T("        hd_diagnostic,") \
	_T("        hms_getobjectname(hd_object) AS doituong,") \
	_T("        trunc_date(htr_admitdate) AS admitdate,") \
	_T("        htr_deptid") \
	_T(" FROM hms_patient") \
	_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
	_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
	_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
	_T(" WHERE htr_deptid not in('A4') %s %s %s %s) %s ") \
	_T(" ORDER BY typeid,htr_deptid,capbac"), szDept, szObject, szRank, szWhere, szDept1);*/
	szSQL.Format(_T(" SELECT distinct * ") \
		_T(" FROM(SELECT distinct 1 typeid,hd_docno,") \
		_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) AS hovaten,") \
		_T("        extract(YEAR FROM hp_birthdate)       AS age,") \
		_T("        (SELECT ss_desc FROM sys_sel where ss_id='hms_rank' and ss_code=hp_rank) as capbac,") \
		_T("        hp_workplace,") \
		_T("        hd_diagnostic,") \
		_T("        hms_getobjectname(hd_object) AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_deptid='%s' AND htr_status='I' %s %s %s AND htr_admitdate < Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT distinct 1 typeid,hd_docno,") \
		_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) AS hovaten,") \
		_T("        extract(YEAR FROM hp_birthdate)       AS age,") \
		_T("        (SELECT ss_desc FROM sys_sel where ss_id='hms_rank' and ss_code=hp_rank) as capbac,") \
		_T("        hp_workplace,") \
		_T("        hd_diagnostic,") \
		_T("        hms_getobjectname(hd_object) AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_deptid='%s' AND htr_status='T' %s %s %s AND htr_admitdate < Cast_string2timestamp('%s') AND htr_dischargedate >= cast_string2timestamp('%s'))") \
		_T(" ORDER BY dischargedate"),pMF->m_szDept,szObject, szRank, szWhere,m_szFromDate,
		pMF->m_szDept,szObject, szRank, szWhere,m_szFromDate,m_szFromDate);
	return szSQL;
}