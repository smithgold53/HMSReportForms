#include "stdafx.h"
#include "EMDrugDeliveryListReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMDrugDeliveryListReport *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMDrugDeliveryListReport* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMDrugDeliveryListReport *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMDrugDeliveryListReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMDrugDeliveryListReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnExamRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMDrugDeliveryListReport* )pWnd)->OnExamRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnExamRoomSelendokFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnExamRoomSelendok();
}
/*static void _OnExamRoomSetfocusFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnExamRoomKillfocus();
}*/
/*static void _OnExamRoomKillfocusFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnExamRoomKillfocus();
}*/
static long _OnExamRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMDrugDeliveryListReport *)pWnd)->OnExamRoomLoadData();
}
/*static void _OnExamRoomAddNewFnc(CWnd *pWnd){
	((CEMDrugDeliveryListReport *)pWnd)->OnExamRoomAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMDrugDeliveryListReport *pVw = (CEMDrugDeliveryListReport *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMDrugDeliveryListReport *pVw = (CEMDrugDeliveryListReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMDrugDeliveryListReport *pVw = (CEMDrugDeliveryListReport *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddEMDrugDeliveryListReportFnc(CWnd *pWnd){
	 return ((CEMDrugDeliveryListReport*)pWnd)->OnAddEMDrugDeliveryListReport();
} 
static int _OnEditEMDrugDeliveryListReportFnc(CWnd *pWnd){
	 return ((CEMDrugDeliveryListReport*)pWnd)->OnEditEMDrugDeliveryListReport();
} 
static int _OnDeleteEMDrugDeliveryListReportFnc(CWnd *pWnd){
	 return ((CEMDrugDeliveryListReport*)pWnd)->OnDeleteEMDrugDeliveryListReport();
} 
static int _OnSaveEMDrugDeliveryListReportFnc(CWnd *pWnd){
	 return ((CEMDrugDeliveryListReport*)pWnd)->OnSaveEMDrugDeliveryListReport();
} 
static int _OnCancelEMDrugDeliveryListReportFnc(CWnd *pWnd){
	 return ((CEMDrugDeliveryListReport*)pWnd)->OnCancelEMDrugDeliveryListReport();
} 
CEMDrugDeliveryListReport::CEMDrugDeliveryListReport(CWnd *pParent)
	{
	m_nDlgWidth = 445;
	m_nDlgHeight = 155;
	m_szTitle = _T("Drug Delivery List Report");
	SetDefaultValues();
}
CEMDrugDeliveryListReport::~CEMDrugDeliveryListReport(){
}
void CEMDrugDeliveryListReport::OnCreateComponents(){
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
void CEMDrugDeliveryListReport::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	////m_wndYear.SetCheckValue(true);
	////m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	//m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	//m_wndToDate.SetCheckValue(true);
	//m_wndExamRoom.SetCheckValue(true);
	m_wndExamRoom.LimitText(35);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 100);

	m_wndExamRoom.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndExamRoom.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);
}
void CEMDrugDeliveryListReport::OnSetWindowEvents()
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
void CEMDrugDeliveryListReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndExamRoom.GetDlgCtrlID(), m_szExamRoomKey);
}
void CEMDrugDeliveryListReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CEMDrugDeliveryListReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CEMDrugDeliveryListReport::SetDefaultValues(){
	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szExamRoomKey.Empty();
}
int CEMDrugDeliveryListReport::SetMode(int nMode){
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
/*void CEMDrugDeliveryListReport::OnYearChange(){
	
} */
/*void CEMDrugDeliveryListReport::OnYearSetfocus(){
	
} */
/*void CEMDrugDeliveryListReport::OnYearKillfocus(){
	
} */
int CEMDrugDeliveryListReport::OnYearCheckValue(){
	return 0;
} 
void CEMDrugDeliveryListReport::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
} 
void CEMDrugDeliveryListReport::OnReportPeriodSelendok(){
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
/*void CEMDrugDeliveryListReport::OnReportPeriodSetfocus(){
	
}*/
/*void CEMDrugDeliveryListReport::OnReportPeriodKillfocus(){
	
}*/
long CEMDrugDeliveryListReport::OnReportPeriodLoadData(){
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
/*void CEMDrugDeliveryListReport::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMDrugDeliveryListReport::OnFromDateChange(){
	
} */
/*void CEMDrugDeliveryListReport::OnFromDateSetfocus(){
	
} */
/*void CEMDrugDeliveryListReport::OnFromDateKillfocus(){
	
} */
int CEMDrugDeliveryListReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMDrugDeliveryListReport::OnToDateChange(){
	
} */
/*void CEMDrugDeliveryListReport::OnToDateSetfocus(){
	
} */
/*void CEMDrugDeliveryListReport::OnToDateKillfocus(){
	
} */
int CEMDrugDeliveryListReport::OnToDateCheckValue(){
	return 0;
} 
void CEMDrugDeliveryListReport::OnExamRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMDrugDeliveryListReport::OnExamRoomSelendok(){
	 
}
/*void CEMDrugDeliveryListReport::OnExamRoomSetfocus(){
	
}*/
/*void CEMDrugDeliveryListReport::OnExamRoomKillfocus(){
	
}*/
long CEMDrugDeliveryListReport::OnExamRoomLoadData(){
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
/*void CEMDrugDeliveryListReport::OnExamRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMDrugDeliveryListReport::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString tmpStr, szSQL;
	int nRetired = 0, nCount = 0, nSoldier = 0;
	int c1 = 0, c2 = 0, c3 = 0, c4 = 0;
	CRecord rs(&pMF->m_db);
	BeginWaitCursor();
	if(!rpt.Init(_T("Reports/HMS/HE_DANHSACHBENHNHANDUOCCAPTHUOC.RPT")) )
		return ;	
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	nCount = rs.GetRecordCount();
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	rs.GetValue(_T("tenkhoa"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);

	int nIndex = 1;
	CReportSection* rptDetail;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%ld"), nIndex ++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("docno"),tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		tmpStr = CDate::Convert(rs.GetValue(_T("examdate")), yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("patientname"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		
		tmpStr = CDate::Convert(rs.GetValue(_T("birthdate")), yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("5"), tmpStr);
		
		rs.GetValue(_T("rank"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("workplace"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("objectname"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("objid"), tmpStr);
		if(tmpStr == _T("1")) 
			{
				c1++;
			}
		else if(tmpStr == _T("2")) 
			{
				c2++;
			}
		else if(tmpStr == _T("11"))
			{
				c3++;
			}

		rs.MoveNext();
			
	}
	FormatCurrency(c1++, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("quan"), tmpStr);
	FormatCurrency(c2++, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("huu"), tmpStr);
	/*FormatCurrency(c3++, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("bhytquan"), tmpStr);*/

	FormatCurrency(nCount, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("sum"), tmpStr);
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();	
} 
void CEMDrugDeliveryListReport::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CString tmpStr, szSQL;
	CRecord rs(&pMF->m_db);
	int c1 = 0, nCount = 0, c2 = 0, c3 = 0, c4 = 0;
	szSQL = GetQueryString();
	BeginWaitCursor();
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	nCount = rs.GetRecordCount();
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1, 12);
	xls.SetColumnWidth(2, 12);
	xls.SetColumnWidth(3, 23);
	xls.SetColumnWidth(4, 11);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 26);
	xls.SetColumnWidth(7, 26);
	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 8);
	xls.SetCellText(nCol, nRow + 3, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N \x110\x1AF\x1EE2\x43 \x43\x1EA4P THU\x1ED0\x43"), FMT_TEXT | FMT_CENTER, true, 13);
	xls.SetCellMergedColumns(nCol, nRow + 4, 8);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	
	//xls.SetCellText(nCol + 2, nRow + 5, _T("Qu\xE2n:"), FMT_TEXT | FMT_CENTER, true, 11);
	//xls.SetCellText(nCol + 4, nRow + 5, _T("H\x1B0u:"), FMT_TEXT | FMT_CENTER, true, 11);
	//xls.SetCellText(nCol + 6, nRow + 5, _T("T\x1ED5ng \x63\x1ED9ng:"), FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol, nRow + 6, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 6, _T("S\x1ED1 HS"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 6, _T("Ng\xE0y kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	
	xls.SetCellText(nCol + 3, nRow + 6, _T("H\x1ECD v\xE0 t\xEAn"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 6, _T("Ng\xE0y sinh"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 5, nRow + 6, _T("\x43\x1EA5p \x62\x1EAD\x63"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 6, nRow + 6, _T("\x110\x1A1n v\x1ECB"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 7, nRow + 6, _T("\x110\x1ED1i t\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER, true, 11);
	int nIndex = 0;
	while(!rs.IsEOF())
	{
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow + 7, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("docno"),tmpStr);
		xls.SetCellText(nCol + 1, nRow + 7, tmpStr, FMT_TEXT);

		tmpStr = CDate::Convert(rs.GetValue(_T("examdate")), yyyymmdd, ddmmyyyy);
		xls.SetCellText(nCol + 2, nRow + 7, tmpStr, FMT_TEXT);

		rs.GetValue(_T("patientname"), tmpStr);
		xls.SetCellText(nCol + 3, nRow + 7, tmpStr, FMT_TEXT);
		
		tmpStr = CDate::Convert(rs.GetValue(_T("birthdate")), yyyymmdd, ddmmyyyy);
		xls.SetCellText(nCol + 4, nRow + 7, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("rank"), tmpStr);
		xls.SetCellText(nCol + 5, nRow + 7, tmpStr, FMT_TEXT);

		rs.GetValue(_T("workplace"), tmpStr);
		xls.SetCellText(nCol + 6, nRow + 7, tmpStr, FMT_TEXT);

		rs.GetValue(_T("objectname"), tmpStr);
		xls.SetCellText(nCol + 7, nRow + 7, tmpStr, FMT_TEXT);

		rs.GetValue(_T("objid"), tmpStr);
		
		if(tmpStr == _T("1")) 
			{
				c1++;
			}
		else if(tmpStr == _T("2")) 
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
	
	CString szTemp;

	szTemp.Format(_T("Qu\xE2n: %d"),c1++);
	xls.SetCellText(1, 5, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("H\x1B0u: %d"),c2++);
	xls.SetCellText(2, 5, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("\x42HYT Qu\xE2n nh\xE2n: %d"),c3++);
	xls.SetCellText(3, 5, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("T\x1ED5ng s\x1ED1: %d"),nIndex++);
	xls.SetCellText(4, 5, szTemp, FMT_TEXT, true, 12);





	EndWaitCursor();
	xls.Save(_T("Exports\\THONG KE BENH NHAN CAP THUOC.xls"));
} 
void CEMDrugDeliveryListReport::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CEMDrugDeliveryListReport::OnAddEMDrugDeliveryListReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMDrugDeliveryListReport::OnEditEMDrugDeliveryListReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMDrugDeliveryListReport::OnDeleteEMDrugDeliveryListReport(){
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
 		OnCancelEMDrugDeliveryListReport(); 
 	}
	return 0;
}
int CEMDrugDeliveryListReport::OnSaveEMDrugDeliveryListReport(){
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
 		//OnEMDrugDeliveryListReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMDrugDeliveryListReport::OnCancelEMDrugDeliveryListReport(){
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
int CEMDrugDeliveryListReport::OnEMDrugDeliveryListReportListLoadData(){
	return 0;
}
CString CEMDrugDeliveryListReport::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	if (!m_szExamRoomKey.IsEmpty())
		szWhere.Format(_T(" AND he_roomid = %d"), ToInt(m_szExamRoomKey));

	szSQL.Format(_T(" SELECT 	distinct hd_docno as docno,") \
				_T(" 	(he_examdate) as examdate,") \
				_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as patientname,") \
				_T(" 	hp_birthdate as birthdate,") \
				_T(" 	get_syssel_desc('hms_rank', hp_rank) as rank,") \
				_T("	get_syssel_desc('hms_position', hp_position) as position,") \
				_T(" 	hp_workplace as workplace, hd_object as objid, ") \
				_T("  	get_objectname(hd_object) as objectname") \
				_T(" FROM hms_doc ") \
				_T(" LEFT JOIN hms_patient ON (hd_patientno = hp_patientno)") \
				_T(" LEFT JOIN hms_exam ON (he_docno = hd_docno)") \
				_T(" LEFT JOIN hms_pharmaorder ON hpo_docno = hd_docno") \
				_T(" WHERE hd_status = 'T' AND hpo_orderstatus in('S', 'A')") \
				_T(" AND (hpo_orderdate) between cast_string2timestamp('%s') and cast_string2timestamp('%s') AND he_deptid = '%s' %s"), m_szFromDate, m_szToDate, pMF->GetCurrentDepartmentID(), szWhere);
	return szSQL;
}