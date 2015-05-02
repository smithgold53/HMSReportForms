#include "stdafx.h"
#include "EMSoluongbenhnhankhamchuyenkhoa.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){ 
	((CEMSoluongbenhnhankhamchuyenkhoa* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnToDateCheckValue();
} 
static void _OnExamRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMSoluongbenhnhankhamchuyenkhoa* )pWnd)->OnExamRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnExamRoomSelendokFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnExamRoomSelendok();
}
/*static void _OnExamRoomSetfocusFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnExamRoomKillfocus();
}*/
/*static void _OnExamRoomKillfocusFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnExamRoomKillfocus();
}*/
static long _OnExamRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnExamRoomLoadData();
}
/*static void _OnExamRoomAddNewFnc(CWnd *pWnd){
	((CEMSoluongbenhnhankhamchuyenkhoa *)pWnd)->OnExamRoomAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMSoluongbenhnhankhamchuyenkhoa *pVw = (CEMSoluongbenhnhankhamchuyenkhoa *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMSoluongbenhnhankhamchuyenkhoa *pVw = (CEMSoluongbenhnhankhamchuyenkhoa *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMSoluongbenhnhankhamchuyenkhoa *pVw = (CEMSoluongbenhnhankhamchuyenkhoa *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddEMSoluongbenhnhankhamchuyenkhoaFnc(CWnd *pWnd){
	 return ((CEMSoluongbenhnhankhamchuyenkhoa*)pWnd)->OnAddEMSoluongbenhnhankhamchuyenkhoa();
} 
static int _OnEditEMSoluongbenhnhankhamchuyenkhoaFnc(CWnd *pWnd){
	 return ((CEMSoluongbenhnhankhamchuyenkhoa*)pWnd)->OnEditEMSoluongbenhnhankhamchuyenkhoa();
} 
static int _OnDeleteEMSoluongbenhnhankhamchuyenkhoaFnc(CWnd *pWnd){
	 return ((CEMSoluongbenhnhankhamchuyenkhoa*)pWnd)->OnDeleteEMSoluongbenhnhankhamchuyenkhoa();
} 
static int _OnSaveEMSoluongbenhnhankhamchuyenkhoaFnc(CWnd *pWnd){
	 return ((CEMSoluongbenhnhankhamchuyenkhoa*)pWnd)->OnSaveEMSoluongbenhnhankhamchuyenkhoa();
} 
static int _OnCancelEMSoluongbenhnhankhamchuyenkhoaFnc(CWnd *pWnd){
	 return ((CEMSoluongbenhnhankhamchuyenkhoa*)pWnd)->OnCancelEMSoluongbenhnhankhamchuyenkhoa();
} 
CEMSoluongbenhnhankhamchuyenkhoa::CEMSoluongbenhnhankhamchuyenkhoa(CWnd *pParent)
{
	m_nDlgWidth = 445;
	m_nDlgHeight = 155;
	m_szTitle = _T("Statistic ICD - Object");
	SetDefaultValues();
}
CEMSoluongbenhnhankhamchuyenkhoa::~CEMSoluongbenhnhankhamchuyenkhoa(){
}
void CEMSoluongbenhnhankhamchuyenkhoa::OnCreateComponents(){
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
void CEMSoluongbenhnhankhamchuyenkhoa::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	//m_wndExamRoom.SetCheckValue(true);
	m_wndExamRoom.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);


	m_wndExamRoom.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndExamRoom.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

}
void CEMSoluongbenhnhankhamchuyenkhoa::OnSetWindowEvents(){
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
void CEMSoluongbenhnhankhamchuyenkhoa::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndExamRoom.GetDlgCtrlID(), m_szExamRoomKey);
}
void CEMSoluongbenhnhankhamchuyenkhoa::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMSoluongbenhnhankhamchuyenkhoa::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMSoluongbenhnhankhamchuyenkhoa::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szExamRoomKey.Empty();

}
int CEMSoluongbenhnhankhamchuyenkhoa::SetMode(int nMode){
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
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnYearChange(){
	
} */
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnYearSetfocus(){
	
} */
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnYearKillfocus(){
	
} */
int CEMSoluongbenhnhankhamchuyenkhoa::OnYearCheckValue(){
	return 0;
} 
void CEMSoluongbenhnhankhamchuyenkhoa::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMSoluongbenhnhankhamchuyenkhoa::OnReportPeriodSelendok(){
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
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnReportPeriodSetfocus(){
	
}*/
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnReportPeriodKillfocus(){
	
}*/
long CEMSoluongbenhnhankhamchuyenkhoa::OnReportPeriodLoadData(){
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
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnFromDateChange(){
	
} */
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnFromDateSetfocus(){
	
} */
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnFromDateKillfocus(){
	
} */
int CEMSoluongbenhnhankhamchuyenkhoa::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnToDateChange(){
	
} */
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnToDateSetfocus(){
	
} */
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnToDateKillfocus(){
	
} */
int CEMSoluongbenhnhankhamchuyenkhoa::OnToDateCheckValue(){
	return 0;
} 
void CEMSoluongbenhnhankhamchuyenkhoa::OnExamRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMSoluongbenhnhankhamchuyenkhoa::OnExamRoomSelendok(){
	 
}
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnExamRoomSetfocus(){
	
}*/
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnExamRoomKillfocus(){
	
}*/
long CEMSoluongbenhnhankhamchuyenkhoa::OnExamRoomLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndExamRoom.IsSearchKey() && !m_szExamRoomKey.IsEmpty())
	{
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
/*void CEMSoluongbenhnhankhamchuyenkhoa::OnExamRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMSoluongbenhnhankhamchuyenkhoa::OnPrintPreviewSelect()
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
	if(!rpt.Init(_T("Reports/HMS/HE_SOLUONGBENHNHANKHAMCHUYENKHOA.RPT")) )
		return ;

	rpt.GetReportHeader()->SetValue(_T("HeathService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDate::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	int nIndex = 1;
	CReportSection* rptDetail;

	int  nTotal1 = 0, nTotal2 = 0, nTotal3 = 0,nTotal4 = 0, nTotal5 = 0;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();	

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		szTemp = rs.GetValue(_T("hrl_name"));
		rptDetail->SetValue(_T("2"),szTemp);

		szTemp = rs.GetValue(_T("quan"));
		nTotal1 += ToInt(szTemp);
		rptDetail->SetValue(_T("3"), szTemp);

		szTemp = rs.GetValue(_T("huu"));
		nTotal2 += ToInt(szTemp);
		rptDetail->SetValue(_T("4"), szTemp);

		szTemp = rs.GetValue(_T("bhytquan"));
		nTotal3 += ToInt(szTemp);
		rptDetail->SetValue(_T("5"), szTemp);

		szTemp = rs.GetValue(_T("captuong"));
		nTotal4 += ToInt(szTemp);
		rptDetail->SetValue(_T("6"), szTemp);

		szTemp = rs.GetValue(_T("tong"));
		nTotal5 += ToInt(szTemp);
		rptDetail->SetValue(_T("7"), szTemp);
		rs.MoveNext();
	}	
	rpt.GetReportFooter()->SetValue(_T("s3"), nTotal1);
	rpt.GetReportFooter()->SetValue(_T("s4"), nTotal2);
	rpt.GetReportFooter()->SetValue(_T("s5"), nTotal3);
	rpt.GetReportFooter()->SetValue(_T("s6"), nTotal4);
	rpt.GetReportFooter()->SetValue(_T("s7"), nTotal5);

	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
	EndWaitCursor();
			
} 
void CEMSoluongbenhnhankhamchuyenkhoa::OnExportSelect(){
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
	xls.SetColumnWidth(4, 20);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 15);

	int nRow = 0, nCol = 0;
	
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellMergedColumns(nCol, nRow + 1, 4);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(nCol, nRow + 2, 6);
	xls.SetCellMergedColumns(nCol, nRow + 3, 6);
	xls.SetCellText(nCol, nRow + 2, _T("S\x1ED0 L\x1AF\x1EE2NG \x42\x1EC6NH NH\xC2N KH\xC1M \x43HUY\xCAN KHO\x41"), FMT_TEXT | FMT_CENTER, true, 13);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y: %s \x110\x1EBFn ng\xE0y: %s"), 
		          CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss),
				  CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellText(nCol, nRow + 4, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 4, _T("Ph\xF2ng kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 4, _T("Qu\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 3, nRow + 4, _T("H\x1B0u"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 4, _T("\x42HYT Qu\xE2n nh\xE2n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 5, nRow + 4, _T("T\x1B0\x1EDBng"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 6, nRow + 4, _T("T\x1ED5ng s\x1ED1"), FMT_TEXT | FMT_CENTER, true, 11);
	
	nRow += 5;
	int nIndex = 1;
	int nTotal1 = 0,nTotal2 = 0, nTotal3 = 0,  nTotal4 = 0, nTotal5 = 0;
	while (!rs.IsEOF())
	{
		szTemp.Format(_T("%d"), nIndex);
		xls.SetCellText(nCol, nRow, szTemp, FMT_INTEGER);

		szTemp = rs.GetValue(_T("hrl_name"));
		xls.SetCellText(nCol + 1, nRow, szTemp, FMT_TEXT);

		szTemp = rs.GetValue(_T("quan"));
		nTotal1 += ToInt(szTemp);
		xls.SetCellText(nCol + 2, nRow, szTemp, FMT_NUMBER1);

		szTemp = rs.GetValue(_T("huu"));
		nTotal2 += ToInt(szTemp);
		xls.SetCellText(nCol + 3, nRow, szTemp, FMT_NUMBER1);

		szTemp = rs.GetValue(_T("bhytquan"));
		nTotal3 += ToInt(szTemp);
		xls.SetCellText(nCol + 4, nRow, szTemp, FMT_NUMBER1);

		szTemp = rs.GetValue(_T("captuong"));
		nTotal4 += ToInt(szTemp);
		xls.SetCellText(nCol + 5, nRow, szTemp, FMT_NUMBER1);

		szTemp = rs.GetValue(_T("tong"));
		nTotal5 += ToInt(szTemp);
		xls.SetCellText(nCol + 6, nRow, szTemp, FMT_NUMBER1, true, 12);

		nIndex++;
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 12);

	xls.SetCellText(nCol + 2, nRow, int2str(nTotal1), FMT_NUMBER1, true, 12);
	xls.SetCellText(nCol + 3, nRow, int2str(nTotal2), FMT_NUMBER1, true, 12);
	xls.SetCellText(nCol + 4, nRow, int2str(nTotal3), FMT_NUMBER1, true, 12);
	xls.SetCellText(nCol + 5, nRow, int2str(nTotal4), FMT_NUMBER1, true, 12);
	xls.SetCellText(nCol + 6, nRow, int2str(nTotal5), FMT_NUMBER1, true, 12);
	EndWaitCursor();
	xls.Save(_T("Exports\\So luong benh nhan kham chuyen khoa.xls"));
} 

void CEMSoluongbenhnhankhamchuyenkhoa::OnCloseSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
	
} 
int CEMSoluongbenhnhankhamchuyenkhoa::OnAddEMSoluongbenhnhankhamchuyenkhoa(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMSoluongbenhnhankhamchuyenkhoa::OnEditEMSoluongbenhnhankhamchuyenkhoa(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMSoluongbenhnhankhamchuyenkhoa::OnDeleteEMSoluongbenhnhankhamchuyenkhoa(){
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
 		OnCancelEMSoluongbenhnhankhamchuyenkhoa(); 
 	}
	return 0;
}
int CEMSoluongbenhnhankhamchuyenkhoa::OnSaveEMSoluongbenhnhankhamchuyenkhoa(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT) 
 		return -1; 
 	if(!IsValidateData()) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
 	CString szSQL; 
 	if(GetMode() == VM_ADD){ 
 		//szSQL = m_tblTbl.GetInsertSQL(); 
 	} 
 	else if(GetMode() == VM_EDIT)
	{ 
 		CString szWhere; 
 		szWhere.Format(_T(" WHERE")); 
 		//szSQL = m_tblTbl.GetUpdateSQL(_T("createdby,createddate")); 
 		szSQL += szWhere; 
 	} 
 _fmsg(_T("%s"), szSQL); 
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret > 0) 
 	{ 
 		//OnEMSoluongbenhnhankhamchuyenkhoaListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMSoluongbenhnhankhamchuyenkhoa::OnCancelEMSoluongbenhnhankhamchuyenkhoa(){
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
int CEMSoluongbenhnhankhamchuyenkhoa::OnEMSoluongbenhnhankhamchuyenkhoaListLoadData(){
	return 0;
}
CString CEMSoluongbenhnhankhamchuyenkhoa::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	if (!m_szExamRoomKey.IsEmpty())
	{
		szWhere.Format(_T(" and he_roomid = %d "), ToInt(m_szExamRoomKey));
	}
	 szSQL.Format(_T("            SELECT hrl_name,hrl_id,") \
				_T("                  SUM(tong)    AS tong,") \
				_T("                  SUM(quan)    AS quan,") \
				_T("                  SUM(huu)    AS huu, ") \
				_T("				  SUM(bhytquan) as bhytquan, ") \
				_T("                  SUM(captuong) as captuong") \
				_T(" 				   FROM") \
				_T(" 				     (") \
				_T("              SELECT hrl_id, hrl_name, ") \
				_T(" 				case when  he_deptid  <> hd_admitdept then 1 else  0 end as tong,") \
				_T("                case when  he_deptid  <> hd_admitdept and hd_object = '1' then 1 else  0 end as quan,") \
				_T("                case when  he_deptid  <> hd_admitdept and hd_object = '2' then 1 else  0 end as huu, ") \
				_T("                case when  he_deptid  <> hd_admitdept and hd_object = '11' then 1 else 0 end as bhytquan, ") \
				_T("                case when  he_deptid  <> hd_admitdept and hp_rank in('15', '16', '17', '18')  then 1 else 0 end as captuong") \
				_T("              FROM hms_doc") \
				_T(" 				     LEFT JOIN hms_patient ON (hd_patientno = hp_patientno)") \
				_T(" 				     LEFT JOIN hms_exam ON(he_docno=hd_docno)") \
				_T(" 				     LEFT JOIN hms_roomlist ON(hrl_deptid=he_deptid AND hrl_id = he_roomid)") \
				_T(" 				     WHERE Cast_Timestamp2Date(he_examdate) BETWEEN cast_string2date('%s') AND cast_string2date('%s')") \
				_T(" 				     AND he_status = 'T'	   ") \
				_T(" 				     AND hd_admitdept ='C1.2' %s") \
				_T(" 				     ) tbl") \
				_T("              Having  SUM(tong) >0") \
				_T(" 				   GROUP BY hrl_id, hrl_name") \
				_T("            order by hrl_id, hrl_name"), m_szFromDate, m_szToDate, szWhere);

	return szSQL;
}