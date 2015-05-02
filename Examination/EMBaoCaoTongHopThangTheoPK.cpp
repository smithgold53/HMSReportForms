#include "stdafx.h"
#include "EMBaoCaoTongHopThangTheoPK.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMGeneralReportC1_1 *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMGeneralReportC1_1* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMGeneralReportC1_1 *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMGeneralReportC1_1 *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMGeneralReportC1_1 *)pWnd)->OnToDateCheckValue();
} 
static void _OnExamRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMGeneralReportC1_1* )pWnd)->OnExamRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnExamRoomSelendokFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnExamRoomSelendok();
}
/*static void _OnExamRoomSetfocusFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnExamRoomKillfocus();
}*/
/*static void _OnExamRoomKillfocusFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnExamRoomKillfocus();
}*/
static long _OnExamRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMGeneralReportC1_1 *)pWnd)->OnExamRoomLoadData();
}
/*static void _OnExamRoomAddNewFnc(CWnd *pWnd){
	((CEMGeneralReportC1_1 *)pWnd)->OnExamRoomAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMGeneralReportC1_1 *pVw = (CEMGeneralReportC1_1 *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMGeneralReportC1_1 *pVw = (CEMGeneralReportC1_1 *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMGeneralReportC1_1 *pVw = (CEMGeneralReportC1_1 *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddrptBaoCaoTongHopThangTheoPKFnc(CWnd *pWnd){
	 return ((CEMGeneralReportC1_1*)pWnd)->OnAddrptBaoCaoTongHopThangTheoPK();
} 
static int _OnEditrptBaoCaoTongHopThangTheoPKFnc(CWnd *pWnd){
	 return ((CEMGeneralReportC1_1*)pWnd)->OnEditrptBaoCaoTongHopThangTheoPK();
} 
static int _OnDeleterptBaoCaoTongHopThangTheoPKFnc(CWnd *pWnd){
	 return ((CEMGeneralReportC1_1*)pWnd)->OnDeleterptBaoCaoTongHopThangTheoPK();
} 
static int _OnSaverptBaoCaoTongHopThangTheoPKFnc(CWnd *pWnd){
	 return ((CEMGeneralReportC1_1*)pWnd)->OnSaverptBaoCaoTongHopThangTheoPK();
} 
static int _OnCancelrptBaoCaoTongHopThangTheoPKFnc(CWnd *pWnd){
	 return ((CEMGeneralReportC1_1*)pWnd)->OnCancelrptBaoCaoTongHopThangTheoPK();
} 
CEMGeneralReportC1_1::CEMGeneralReportC1_1(CWnd *pParent)
	{
	m_nDlgWidth = 440;
	m_nDlgHeight = 155;
	//m_szTitle = _T("Synthesis Report Month Examination");
	m_szTitle = _T("\x42\xE1o \x63\xE1o t\x1ED5ng h\x1EE3p th\xE1ng th\x65o ph\xF2ng kh\xE1m");
	SetDefaultValues();
}
CEMGeneralReportC1_1::~CEMGeneralReportC1_1(){
}
void CEMGeneralReportC1_1::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 120);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndExamRoomLabel.Create(this, _T("Exam Room"), 10, 90, 90, 115);
	m_wndExamRoom.Create(this,95, 90, 425, 115); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 240, 125, 350, 150);
	m_wndExport.Create(this, _T("&Export"), 355, 125, 430, 150);

}
void CEMGeneralReportC1_1::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	//m_wndExamRoom.SetCheckValue(true);
	m_wndExamRoom.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);


	m_wndExamRoom.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndExamRoom.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

}
void CEMGeneralReportC1_1::OnSetWindowEvents(){
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
void CEMGeneralReportC1_1::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndExamRoom.GetDlgCtrlID(), m_szExamRoomKey);

}
void CEMGeneralReportC1_1::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMGeneralReportC1_1::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMGeneralReportC1_1::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szExamRoomKey.Empty();

}
int CEMGeneralReportC1_1::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0,1,2,3,-1); 
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
/*void CEMGeneralReportC1_1::OnYearChange(){
	
} */
/*void CEMGeneralReportC1_1::OnYearSetfocus(){
	
} */
/*void CEMGeneralReportC1_1::OnYearKillfocus(){
	
} */
int CEMGeneralReportC1_1::OnYearCheckValue(){
	return 0;
} 
void CEMGeneralReportC1_1::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMGeneralReportC1_1::OnReportPeriodSelendok(){
	CString tmpStr;
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
	UpdateData(false);	 
	 
}
/*void CEMGeneralReportC1_1::OnReportPeriodSetfocus(){
	
}*/
/*void CEMGeneralReportC1_1::OnReportPeriodKillfocus(){
	
}*/
long CEMGeneralReportC1_1::OnReportPeriodLoadData(){
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
/*void CEMGeneralReportC1_1::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMGeneralReportC1_1::OnFromDateChange(){
	
} */
/*void CEMGeneralReportC1_1::OnFromDateSetfocus(){
	
} */
/*void CEMGeneralReportC1_1::OnFromDateKillfocus(){
	
} */
int CEMGeneralReportC1_1::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMGeneralReportC1_1::OnToDateChange(){
	
} */
/*void CEMGeneralReportC1_1::OnToDateSetfocus(){
	
} */
/*void CEMGeneralReportC1_1::OnToDateKillfocus(){
	
} */
int CEMGeneralReportC1_1::OnToDateCheckValue(){
	//CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//CDateTime sysdate, dte;
	//sysdate.ParseDateTime(pMF->GetSysDateTime());
	//dte.ParseDateTime(m_szToDate);
	//if (dte > sysdate)
	//	return -1;
	return 0;
} 
void CEMGeneralReportC1_1::OnExamRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMGeneralReportC1_1::OnExamRoomSelendok(){
	 
}
/*void CEMGeneralReportC1_1::OnExamRoomSetfocus(){
	
}*/
/*void CEMGeneralReportC1_1::OnExamRoomKillfocus(){
	
}*/
long CEMGeneralReportC1_1::OnExamRoomLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndExamRoom.IsSearchKey() && !m_szExamRoomKey.IsEmpty()){
		szWhere.Format(_T("and hrl_id = %d"), ToInt(m_szExamRoomKey));
	};
	m_wndExamRoom.DeleteAllItems(); 
	szSQL.Format(_T("select hrl_name as name, hrl_id as id from hms_roomlist where hrl_deptid = '%s' %s Order by id"), pMF->GetCurrentDepartmentID(), szWhere);
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
/*void CEMGeneralReportC1_1::OnExamRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMGeneralReportC1_1::OnPrintPreviewSelect(){
	if (!IsValidateData())
		return;
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString tmpStr, szSQL, szRoomID, szTemp, szWhere;
	CRecord rs(&pMF->m_db);
	CReportSection* rptDetail;
	CString szDept = pMF->GetCurrentDepartmentID();
	if (szDept != _T("C1.1"))
	{
		AfxMessageBox(_T("This report can only run by C1.1 user!"), MB_ICONASTERISK);
		return;
	}
	if(!rpt.Init(_T("Reports/HMS/HE_BAOCAOTONGHOPTHANGPHONGKHAM2.RPT")) )
		return;
	
	BeginWaitCursor();

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	int nIndex = 1;
	int nSel = m_wndExamRoom.GetCurSel();
	szRoomID = m_wndExamRoom.GetCurrent(nSel, 0);
	CStringArray arrCol;
	for (int i = 0; i < 13; i++)
	{
		tmpStr.Format(_T("c%d"), i+1);
		arrCol.Add(tmpStr);
	}
	szWhere.Empty();
	//So dang ky
	szWhere.Format(_T(" AND he_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty()){
		szWhere.AppendFormat(_T(" and he_roomid=%d"), ToInt(m_szExamRoomKey));
	}
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object=4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T("         SELECT distinct hd_docno, hd_object ") \
				_T("         FROM hms_doc") \
				_T("         LEFT JOIN hms_exam ON(he_docno=hd_docno)") \
				_T("         WHERE hd_admitdate BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T("		 %s") \
				_T(" ) tbl") \
				_T(" GROUP BY hd_docno, hd_object) tbl2"), m_szFromDate, m_szToDate, szWhere);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rptDetail->SetValue(_T("2"), _T("S\x1ED1 \x111\x103ng k\xFD"));
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			szTemp.Format(_T("%d"), j+3);
			rptDetail->SetValue(szTemp, tmpStr);
		}
	}
	//So ng kham
	szWhere.Format(_T(" AND he_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty()){
		szWhere.AppendFormat(_T(" and he_roomid=%d"), ToInt(m_szExamRoomKey));
	}
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object=4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T("		 SELECT	  he_docno          docno, ") \
				_T("                  Min(he_receptidx) receptidx ") \
				_T("           FROM   hms_exam ") \
				_T("           WHERE  he_status IN ( 'T', 'P' ) ") \
				_T("              AND he_deptid = 'C1.1' ") \
				_T("           GROUP  BY he_docno) tbl") \
				_T(" LEFT JOIN hms_exam ON( he_docno = docno ") \
				_T("                        AND he_receptidx = receptidx ) ") \
				_T(" LEFT JOIN hms_doc ON ( he_docno = hd_docno ) ") \
				_T(" WHERE he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
				_T(" %s") \
				_T(" GROUP BY hd_docno, hd_object) tbl2"), m_szFromDate, m_szToDate, szWhere);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rptDetail->SetValue(_T("2"), _T("S\x1ED1 ng\x1B0\x1EDDi kh\xE1m"));
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			szTemp.Format(_T("%d"), j+3);
			rptDetail->SetValue(szTemp, tmpStr);
		}
	}

	//So lan kham
	szWhere.Format(_T(" AND he_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	if (!m_szExamRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND he_roomid = %ld"), ToInt(m_szExamRoomKey));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_exam ON(he_docno=hd_docno)") \
				_T(" WHERE   (he_examdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T("         and he_status in ('P','T') ") \
				_T(" %s") \
				_T(" ) tbl"), m_szFromDate, m_szToDate, szWhere);
	//QueryOpener(szSQL);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rptDetail->SetValue(_T("2"), _T("S\x1ED1 l\x1B0\x1EE3t kh\xE1m"));
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			szTemp.Format(_T("%d"), j+3);
			rptDetail->SetValue(szTemp, tmpStr);
		}
	}
	//So ke don
	szWhere.Format(_T(" AND hpo_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_roomid=%d"), ToInt(m_szExamRoomKey));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_pharmaorder ON(hpo_docno=hd_docno)") \
				_T(" WHERE (hpo_orderdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T("         and hpo_orderstatus IN ( 'A', 'S')") \
				_T("     %s") \
				_T(" ) tbl"), m_szFromDate, m_szToDate, szWhere);
	
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rptDetail->SetValue(_T("2"), _T("S\x1ED1 k\xEA \x111\x1A1n"));
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			szTemp.Format(_T("%d"), j+3);
			rptDetail->SetValue(szTemp, tmpStr);
		}
	}
/////////////////
//So vao vien
/////////////////
	szWhere.Empty();
	szWhere.Format(_T(" AND hd_enddept = '%s'"), pMF->GetCurrentDepartmentID());
		if(!m_szExamRoomKey.IsEmpty())
			szWhere.AppendFormat(_T(" AND he_roomid=%ld"), ToInt(szRoomID));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_exam ON (he_docno = hd_docno AND he_doctor = hd_doctor)")
				_T(" LEFT JOIN hms_clinical_record ON (hd_docno = hcr_docno)") \
				_T(" WHERE   (hcr_admitdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T("  and hcr_numinward > 0 ") \
				_T(" %s") \
				_T(" ) tbl2"), m_szFromDate, m_szToDate, szWhere);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rptDetail->SetValue(_T("2"), _T("S\x1ED1 v\xE0o vi\x1EC7n"));
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			szTemp.Format(_T("%d"), j+3);
			rptDetail->SetValue(szTemp, tmpStr);
		}
	}
/////////////////
//So chuyen vien
/////////////////
	szWhere.Empty();
	szWhere.Format(_T(" AND hd_admitdept = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND he_roomid=%ld"), ToInt(szRoomID));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc ") \
				_T(" LEFT JOIN hms_exam ON (he_docno = hd_docno AND he_doctor = hd_doctor)")
				_T(" WHERE   (hd_admitdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s') ") \
				_T("         and hd_status ='T' and hd_suggestion='F' ") \
				_T(" %s") \
				_T(" ) tbl2"), m_szFromDate, m_szToDate,  szWhere);
	

	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rptDetail->SetValue(_T("2"), _T("S\x1ED1 \x63huy\x1EC3n vi\x1EC7n"));
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			szTemp.Format(_T("%d"), j+3);
			rptDetail->SetValue(szTemp, tmpStr);
		}
	}

/////////////////
//cho ve
/////////////////
	szWhere.Empty();
	szWhere.Format(_T(" AND hd_admitdept = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND he_roomid=%ld"), ToInt(szRoomID));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_exam ON (he_docno = hd_docno AND he_doctor = hd_doctor)")
				_T(" WHERE   (hd_admitdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T(" 	and hd_status ='T' AND hd_suggestion = 'A' ") \
				_T(" %s") \
				_T(" ) tbl2"), m_szFromDate, m_szToDate, szWhere);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rptDetail->SetValue(_T("2"), _T("S\x1ED1 \x63ho v\x1EC1"));
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			szTemp.Format(_T("%d"), j+3);
			rptDetail->SetValue(szTemp, tmpStr);
		}
	}
/////////////////
//So tu vong
/////////////////
	szWhere.Empty();
	szWhere.Format(_T(" AND hd_admitdept = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND he_roomid=%ld"), ToInt(szRoomID));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_exam ON (hd_docno = he_docno AND he_doctor = hd_doctor)")
				_T(" WHERE   (hd_admitdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T(" 	and hd_status ='T' and hd_suggestion IN ('X', 'Y') ") \
				_T(" %s") \
				_T(" ) tbl2"), m_szFromDate, m_szToDate, szWhere);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rptDetail->SetValue(_T("2"), _T("S\x1ED1 t\x1EED vong"));
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			szTemp.Format(_T("%d"), j+3);
			rptDetail->SetValue(szTemp, tmpStr);
		}
	}


/////////////////
//So phau thuat ngoai tru
/////////////////
//	szWhere.Empty();
//	if(!m_szExamRoomKey.IsEmpty())
//		//szWhere.Format(" AND hd_treatdate=%ld", atol(szRoomID));
//
//		szSQL.Format("SELECT sum(c1) as c1, \
//        sum(c2) as c2, \
//        sum(c3) as c3, \
//        sum(c4) as c4, \
//        sum(c5) as c5, \
//        sum(c6) as c6, \
//        sum(c7) as c7, \
//        sum(c8) as c8, \
//		sum(c9) as c9, \
//        sum(c10) as c10 "\
//"FROM  \
//( \
//SELECT hd_object as policy, \
//        case when hd_object=1 then 1 else 0 end as c1, \
//        case when hd_object=2 then 1 else 0 end as c2, \
//        case when hd_object=8 then 1 else 0 end as c3, \
//		case when hd_object in(6, 9) then 1 else 0 end as c4, \
//        case when hd_object=3 then 1 else 0 end as c5, \
//        case when hd_object=7 then 1 else 0 end as c6, \
//        case when hd_object=4 then 1 else 0 end as c7, \
//        case when hd_object=5 then 1 else 0 end as c8, \
//		case when hd_object=10 then 1 else 0 end as c9, \
//        1 as c10 \
//FROM hms_doc \
//WHERE   hd_sidate BETWEEN date('%s') and date('%s') \
//        and hd_ward='%s' and hd_status ='T' and hd_suggest='G' %s \
//) as tbl2", m_szFromDate, m_szToDate, m_ward, szWhere);
//	
//
//		foundFlg = rs->exec(szSQL);
//		while(foundFlg){
//			m_wndList->insertItem(&lvi);
//			m_wndList->setItemText(lvi.iItem, 0, "PhÉu thuËt ngo¹i tró");
//			for (int i =1; i <= 10; i++){
//				szCol.Format("c%d", i);
//				szData = rs->getFld(szCol);
//				m_wndList->setItemText(lvi.iItem, i, atol(szData) > 0?util.FormatNumStr(szData):"");
//			}
//			foundFlg = rs->findItem(NEXT);
//		}
	szWhere.Empty();
	szWhere.Format(_T(" AND hd_admitdept = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty())
			szWhere.AppendFormat(_T(" AND hpc_roomid =%ld"), ToInt(m_szExamRoomKey));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  hfg_name,") \
				_T("         sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hpc_groupid,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object=4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_testorder ON(hpc_docno=hd_docno)") \
				_T(" WHERE   (hpc_orderdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T("         and hpc_status <> 'O' and length(hpc_groupid) > 0 ") \
				_T(" %s") \
				_T(" GROUP BY hpc_docno, hpc_orderid,hd_object, hpc_groupid") \
				_T(" ) tbl") \
				_T(" LEFT JOIN hms_fee_group ON(hfg_id=hpc_groupid)") \
				_T(" GROUP BY hpc_groupid, hfg_name") \
				_T(" ORDER BY hpc_groupid"), m_szFromDate, m_szToDate, szWhere);

	rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("hfg_name"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			szTemp.Format(_T("%d"), j+3);
			rptDetail->SetValue(szTemp, tmpStr);
		}
		rs.MoveNext();
	}
	//for (int i = 1; i <= 10; i++)
	//{
	//	CString  szField;
	//	szField.Format(_T("s%d"), i+2);
	//	tmpStr.Format(_T("%d"), nTotal[i]);
	//	rpt.GetReportFooter()->SetValue(szField, tmpStr);		
	//}
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();
	
} 
void CEMGeneralReportC1_1::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	if (!IsValidateData())
		return;
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CString tmpStr, szSQL, szWhere, szRoomID, szTemp;
	CRecord rs(&pMF->m_db);
	CString szDept = pMF->GetCurrentDepartmentID();
	if (szDept != _T("C1.1"))
	{
		AfxMessageBox(_T("This report can only run by C1.1 user!"), MB_ICONASTERISK);
		return;
	}
	BeginWaitCursor();
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 30);
	xls.SetColumnWidth(2, 13);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(5, 13);
	xls.SetColumnWidth(6, 13);
	xls.SetColumnWidth(7, 10);
	xls.SetColumnWidth(8, 14);
	xls.SetColumnWidth(9, 14);
	xls.SetColumnWidth(10, 11);
	xls.SetColumnWidth(11, 12);

	int nCol = 0;
	int nRow = 0;
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	rs.GetValue(_T("tenkhoa"), tmpStr);
	xls.SetCellText(nCol, nRow + 2, tmpStr, FMT_TEXT, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 12);
	xls.SetCellText(nCol, nRow + 3, _T("\x42\xC1O \x43\xC1O T\x1ED4NG H\x1EE2P TH\xC1NG TH\x45O PH\xD2NG KH\xC1M"), FMT_TEXT | FMT_CENTER, true, 14);
	xls.SetCellMergedColumns(nCol, nRow + 4, 12);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol, nRow + 5, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 5, _T("Ph\xF2ng kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 5, _T("Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 3, nRow + 5, _T("BHYT Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 5, _T("\x43h\xEDnh s\xE1\x63h"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 5, nRow + 5, _T("BHYT Kh\xE1\x63"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 6, nRow + 5, _T("BHYT TNSQ"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 7, nRow + 5, _T("< 6 tu\x1ED5i"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 8, nRow + 5, _T("\x44\x1ECB\x63h v\x1EE5"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 9, nRow + 5, _T("\x42\x1EA1n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 10, nRow + 5, _T("< \x36 tu\x1ED5i(th\x1EBB)"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 11, nRow + 5, _T("BHYT Qu\xE2n \x111\x1ED9i"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 12, nRow + 5, _T("BHYT Qu\xE2n nh\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 13, nRow + 5, _T("BHYT (n\x1EE3 th\x1EBB)"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 14, nRow + 5, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 11);
	nRow = 6;
	int nIndex = 1;
	int nSel = m_wndExamRoom.GetCurSel();
	szRoomID = m_wndExamRoom.GetCurrent(nSel, 0);
	CStringArray arrCol;
	for (int i = 0; i < 13; i++)
	{
		tmpStr.Format(_T("c%d"), i+1);
		arrCol.Add(tmpStr);
	}
	//So dang ky
	szWhere.Empty();
	szWhere.Format(_T(" AND he_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty()){
		szWhere.AppendFormat(_T(" and he_roomid=%d"), ToInt(m_szExamRoomKey));
	}
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object=4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T("         SELECT distinct hd_docno, hd_status, hd_object, ") \
				_T("		 first_value(he_examdate) over (partition by hd_docno order by he_receptidx) as exdate") \
				_T("         FROM hms_doc") \
				_T("         LEFT JOIN hms_exam ON(he_docno=hd_docno)") \
				_T("		 WHERE hd_admitdate BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T("		 %s") \
				_T(" ) tbl") \
				_T(" GROUP BY hd_docno, hd_object) tbl2"), m_szFromDate, m_szToDate, szWhere);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, _T("S\x1ED1 \x111\x103ng k\xFD"), FMT_TEXT);
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			xls.SetCellText(j+2, nRow, tmpStr, FMT_INTEGER);
		}
		nRow++;
	}
	//So ng kham
	szWhere.Empty();
	szWhere.Format(_T(" AND he_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty()){
		szWhere.AppendFormat(_T(" and he_roomid=%d"), ToInt(m_szExamRoomKey));
	}
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object=4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T("		 SELECT	  he_docno          docno, ") \
				_T("                  Min(he_receptidx) receptidx ") \
				_T("           FROM   hms_exam ") \
				_T("           WHERE  he_status IN ( 'T', 'P' ) ") \
				_T("              AND he_deptid = 'C1.1' ") \
				_T("           GROUP  BY he_docno) tbl") \
				_T(" LEFT JOIN hms_exam ON( he_docno = docno ") \
				_T("                        AND he_receptidx = receptidx ) ") \
				_T(" LEFT JOIN hms_doc ON ( he_docno = hd_docno ) ") \
				_T(" WHERE he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') ") \
				_T(" %s") \
				_T(" GROUP BY hd_docno, hd_object) tbl2"), m_szFromDate, m_szToDate, szWhere);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, _T("S\x1ED1 ng\x1B0\x1EDDi kh\xE1m"), FMT_TEXT);
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			xls.SetCellText(j+2, nRow, tmpStr, FMT_INTEGER);
		}
		nRow++;
	}

	//So lan kham
	szWhere.Format(_T(" AND he_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	if (!m_szExamRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND he_roomid = %ld"), ToInt(m_szExamRoomKey));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_exam ON(he_docno=hd_docno)") \
				_T(" WHERE   (he_examdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T("         AND he_status in ('P','T') ") \
				_T(" %s") \
				_T(" ) tbl"), m_szFromDate, m_szToDate, szWhere);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, _T("S\x1ED1 l\x1B0\x1EE3t kh\xE1m"), FMT_TEXT);
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			xls.SetCellText(j+2, nRow, tmpStr, FMT_INTEGER);
		}
		nRow++;
	}
	//So ke don
	szWhere.Format(_T(" AND hpo_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_roomid=%d"), ToInt(m_szExamRoomKey));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_pharmaorder ON(hpo_docno=hd_docno)") \
				_T(" WHERE (hpo_orderdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T("         and hpo_orderstatus IN ( 'A', 'S')") \
				_T("     %s") \
				_T(" ) tbl"), m_szFromDate, m_szToDate, szWhere);
	
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, _T("S\x1ED1 k\xEA \x111\x1A1n"), FMT_TEXT);
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			xls.SetCellText(j+2, nRow, tmpStr, FMT_INTEGER);
		}
		nRow++;
	}
	/////////////////
	//So vao vien
	/////////////////
	szWhere.Empty();
	szWhere.Format(_T(" AND hd_enddept = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty())
		szWhere.AppendFormat(_T("AND he_roomid=%ld"), ToInt(szRoomID));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_exam ON (he_docno = hd_docno AND he_doctor = hd_doctor)")
				_T(" LEFT JOIN hms_clinical_record ON (hd_docno = hcr_docno)") \
				_T(" WHERE hcr_admitdate BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T("  and hcr_numinward > 0 ") \
				_T(" %s") \
				_T(" ) tbl2"), m_szFromDate, m_szToDate, szWhere);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, _T("S\x1ED1 v\xE0o vi\x1EC7n"), FMT_TEXT);
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			xls.SetCellText(j+2, nRow, tmpStr, FMT_INTEGER);
		}
		nRow++;
	}

	/////////////////
	//So chuyen vien
	/////////////////
	szWhere.Empty();
	szWhere.Format(_T(" AND hd_admitdept = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND he_roomid=%ld"), ToInt(szRoomID));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc ") \
				_T(" LEFT JOIN hms_exam ON (he_docno = hd_docno AND he_doctor = hd_doctor)")
				_T(" WHERE   (hd_admitdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s') ") \
				_T("         and hd_status ='T' and hd_suggestion='F' ") \
				_T(" %s") \
				_T(" ) tbl2"), m_szFromDate, m_szToDate,  szWhere);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, _T("S\x1ED1 \x63huy\x1EC3n vi\x1EC7n"), FMT_TEXT);
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			xls.SetCellText(j+2, nRow, tmpStr, FMT_INTEGER);
		}
		nRow++;
	}
	/////////////////
	//cho ve
	/////////////////
	szWhere.Empty();
	szWhere.Format(_T(" AND hd_admitdept = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND he_roomid=%ld"), ToInt(szRoomID));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_exam ON (he_docno = hd_docno AND he_doctor = hd_doctor)")
				_T(" WHERE   (hd_admitdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T(" 	and hd_status ='T' AND hd_suggestion = 'A' ") \
				_T(" %s") \
				_T(" ) tbl2"), m_szFromDate, m_szToDate, szWhere);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, _T("S\x1ED1 \x63ho v\x1EC1"), FMT_TEXT);
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			xls.SetCellText(j+2, nRow, tmpStr, FMT_INTEGER);
		}
		nRow++;
	}
	/////////////////
	//So tu vong
	/////////////////
	szWhere.Empty();
	szWhere.Format(_T(" AND hd_admitdept = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND he_roomid=%ld"), ToInt(szRoomID));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hd_object as policy,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object = 4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_exam e ON (hd_docno = he_docno AND he_doctor = hd_doctor)")
				_T(" WHERE   (hd_admitdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T(" 	and hd_status ='T' and hd_suggestion IN ('X', 'Y') ") \
				_T(" %s") \
				_T(" ) tbl2"), m_szFromDate, m_szToDate, szWhere);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, _T("S\x1ED1 t\x1EED vong"), FMT_TEXT);
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			xls.SetCellText(j+2, nRow, tmpStr, FMT_INTEGER);
		}
		nRow++;
	}
	szWhere.Format(_T(" AND hd_admitdept = '%s'"), pMF->GetCurrentDepartmentID());
	if(!m_szExamRoomKey.IsEmpty())
			szWhere.AppendFormat(_T(" AND hpc_roomid =%ld"), ToInt(m_szExamRoomKey));
	szSQL.Empty();
	szSQL.Format(_T(" SELECT  hfg_name,") \
				_T("         sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12, ") \
				_T("         sum(c13) as c13 ") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hpc_groupid,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=3 then 1 else 0 end as c3,") \
				_T("         case when hd_object=4 then 1 else 0 end as c4,") \
				_T("         case when hd_object=5 then 1 else 0 end as c5,") \
				_T("         case when hd_object=6 then 1 else 0 end as c6,") \
				_T("         case when hd_object=7 then 1 else 0 end as c7,") \
				_T("         case when hd_object=8 then 1 else 0 end as c8,") \
				_T("         case when hd_object=9 then 1 else 0 end as c9,") \
				_T("         case when hd_object=10 then 1 else 0 end as c10,") \
				_T("         case when hd_object=11 then 1 else 0 end as c11,") \
				_T("         case when hd_object=12 then 1 else 0 end as c12,") \
				_T("         1 as c13 ") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_testorder ON(hpc_docno=hd_docno)") \
				_T(" WHERE   (hpc_orderdate) BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T("         and hpc_status <> 'O' and length(hpc_groupid) > 0 ") \
				_T(" %s") \
				_T(" GROUP BY hpc_docno, hpc_orderid,hd_object, hpc_groupid") \
				_T(" ) tbl") \
				_T(" LEFT JOIN hms_fee_group ON(hfg_id=hpc_groupid)") \
				_T(" GROUP BY hpc_groupid, hfg_name") \
				_T(" ORDER BY hpc_groupid"), m_szFromDate, m_szToDate, szWhere);
	

	rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);
		rs.GetValue(_T("hfg_name"), tmpStr);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
		for (int j = 0; j< arrCol.GetCount(); j++)
		{
			rs.GetValue(arrCol.GetAt(j), tmpStr);
			xls.SetCellText(j+2, nRow, tmpStr, FMT_INTEGER);
		}
		nRow++;
		rs.MoveNext();
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\BAO CAO TONG HOP THANG THEO PHONG KHAM.xls"));
	
} 
void CEMGeneralReportC1_1::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CEMGeneralReportC1_1::OnAddrptBaoCaoTongHopThangTheoPK(){
	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMGeneralReportC1_1::OnEditrptBaoCaoTongHopThangTheoPK(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMGeneralReportC1_1::OnDeleterptBaoCaoTongHopThangTheoPK(){
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
 		OnCancelrptBaoCaoTongHopThangTheoPK(); 
 	}
	return 0;
}
int CEMGeneralReportC1_1::OnSaverptBaoCaoTongHopThangTheoPK(){
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
 		//OnrptBaoCaoTongHopThangTheoPKListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMGeneralReportC1_1::OnCancelrptBaoCaoTongHopThangTheoPK(){
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
int CEMGeneralReportC1_1::OnrptBaoCaoTongHopThangTheoPKListLoadData(){
	return 0;
}
CString CEMGeneralReportC1_1::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	return szSQL;
}BEGIN_MESSAGE_MAP(CEMGeneralReportC1_1, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CEMGeneralReportC1_1::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);
}