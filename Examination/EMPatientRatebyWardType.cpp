#include "stdafx.h"
#include "EMPatientRatebyWardType.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMPatientRatebyWardType *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMPatientRatebyWardType* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMPatientRatebyWardType *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMPatientRatebyWardType *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMPatientRatebyWardType *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMPatientRatebyWardType *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMPatientRatebyWardType *pVw = (CEMPatientRatebyWardType *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMPatientRatebyWardType *pVw = (CEMPatientRatebyWardType *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMPatientRatebyWardType *pVw = (CEMPatientRatebyWardType *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMPatientRatebyWardType*)pWnd)->OnAddEMStatisticsGeneralPatientReport();
} 
static int _OnEditEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMPatientRatebyWardType*)pWnd)->OnEditEMStatisticsGeneralPatientReport();
} 
static int _OnDeleteEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMPatientRatebyWardType*)pWnd)->OnDeleteEMStatisticsGeneralPatientReport();
} 
static int _OnSaveEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMPatientRatebyWardType*)pWnd)->OnSaveEMStatisticsGeneralPatientReport();
} 
static int _OnCancelEMStatisticsGeneralPatientReportFnc(CWnd *pWnd){
	 return ((CEMPatientRatebyWardType*)pWnd)->OnCancelEMStatisticsGeneralPatientReport();
} 
CEMPatientRatebyWardType::CEMPatientRatebyWardType(CWnd *pParent)
{
	m_szTitle = _T("Patients General Statistics");
	m_nDlgWidth = 420;
	m_nDlgHeight = 125;
	SetDefaultValues();
}
CEMPatientRatebyWardType::~CEMPatientRatebyWardType()
{
}
void CEMPatientRatebyWardType::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 205, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 215, 30, 295, 55);
	m_wndReportPeriod.Create(this,300, 30, 410, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 415, 90);
	m_wndFromDate.Create(this,95, 60, 205, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 215, 60, 295, 85);
	m_wndToDate.Create(this,300, 60, 410, 85); 
	//m_wndPrint.Create(this, _T("&Print"), 250, 95, 330, 120);
	m_wndExport.Create(this, _T("&Export"), 335, 95, 415, 120);

}
void CEMPatientRatebyWardType::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(false);
	m_wndReportPeriod.SetCheckValue(false);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(false);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(false);
	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);
}
void CEMPatientRatebyWardType::OnSetWindowEvents(){
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
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
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
void CEMPatientRatebyWardType::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
}
void CEMPatientRatebyWardType::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CEMPatientRatebyWardType::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMPatientRatebyWardType::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMPatientRatebyWardType::SetMode(int nMode){
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
/*void CEMPatientRatebyWardType::OnYearChange(){
	
} */
/*void CEMPatientRatebyWardType::OnYearSetfocus(){
	
} */
/*void CEMPatientRatebyWardType::OnYearKillfocus(){
	
} */
int CEMPatientRatebyWardType::OnYearCheckValue(){
	return 0;
} 
void CEMPatientRatebyWardType::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMPatientRatebyWardType::OnReportPeriodSelendok(){
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
/*void CEMPatientRatebyWardType::OnReportPeriodSetfocus(){
	
}*/
/*void CEMPatientRatebyWardType::OnReportPeriodKillfocus(){
	
}*/
long CEMPatientRatebyWardType::OnReportPeriodLoadData()
{
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
/*void CEMPatientRatebyWardType::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMPatientRatebyWardType::OnFromDateChange(){
	
} */
/*void CEMPatientRatebyWardType::OnFromDateSetfocus(){
	
} */
/*void CEMPatientRatebyWardType::OnFromDateKillfocus(){
	
} */
int CEMPatientRatebyWardType::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMPatientRatebyWardType::OnToDateChange(){
	
} */
/*void CEMPatientRatebyWardType::OnToDateSetfocus(){
	
} */
/*void CEMPatientRatebyWardType::OnToDateKillfocus(){
	
} */
int CEMPatientRatebyWardType::OnToDateCheckValue(){
	return 0;
} 
void CEMPatientRatebyWardType::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	CRecord rs(&pMF->m_db);
	if(!rpt.Init(_T("Reports/HMS/HE_THONGKEBENHNHANCAPTUONG.RPT")) )
	return ;
	BeginWaitCursor();
	UpdateData(true);
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);

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
		rptDetail->SetValue(_T("0"), tmpStr);
		
		rs.GetValue(_T("docno"),tmpStr);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("tenBN"),tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("tencapbac"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("chucvu"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("donvi"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		
		rs.GetValue(_T("tendoituong"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("icd10"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		
		rs.GetValue(_T("benhchinh"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.MoveNext();	
	}	
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();
	
} 
void CEMPatientRatebyWardType::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CString szSQL, tmpStr;
	int nNewType = 0, nOldType = 0;
	double nTmp = 0, nGrpQty = 0, nTotal = 0;
	CRecord rs(&pMF->m_db);
	BeginWaitCursor();
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 30);
	xls.SetColumnWidth(1, 15);
	//xls.SetColumnWidth(2, 25);
	//xls.SetColumnWidth(3, 12);
	//xls.SetColumnWidth(4, 23);
	//xls.SetColumnWidth(5, 23);
	//xls.SetColumnWidth(6, 12);
	//xls.SetColumnWidth(7, 12);
	//xls.SetColumnWidth(8, 66);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 1);
	xls.SetCellMergedColumns(nCol, nRow + 2, 1);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 2);
	xls.SetCellText(nCol, nRow + 3, _T("TH\x1ED0NG K\xCA \x42\x1EC6NH NH\xC2N N\x1ED8I, NGO\x1EA0I, \x43HUY\xCAN KHO\x41"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 2);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol, nRow + 5, _T("\x43huy\xEAn kho\x61 kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 5, _T("S\x1ED1 l\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER, true, 11);

	nRow = 6;
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("hrl_type"), nNewType);
		if (nNewType != nOldType)
		{
			if ((nNewType == 2 || nNewType == 3) || (nNewType > 3 && nOldType <=3) )
			{
				if (nGrpQty > 0)
				{
					xls.SetCellText(0, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
					xls.SetCellText(1, nRow, double2str(nGrpQty), FMT_NUMBER1, true);
					nTotal += nGrpQty;
					nGrpQty = 0;
					nRow++;
				}
				tmpStr.Empty();
				if (nNewType == 2)
					tmpStr = pMF->GetSelectionString(_T("hms_exam_type"), _T("1"));
				else if (nNewType == 3)
					tmpStr = pMF->GetSelectionString(_T("hms_exam_type"), _T("2"));
				else if (nNewType > 3)
					tmpStr = pMF->GetSelectionString(_T("hms_exam_type"), _T("3"));
				if (!tmpStr.IsEmpty())
					xls.SetCellText(0, nRow, tmpStr, 4098, true);
				nOldType = nNewType;
				nRow++;
			}
		}
		rs.GetValue(_T("nme"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr);
		rs.GetValue(_T("cnt"), nTmp);
		nGrpQty += nTmp;
		FormatCurrency(nTmp, tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_RIGHT);
		nRow++;
		rs.MoveNext();		
	}
	if (nGrpQty > 0)
	{
		xls.SetCellText(0, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
		xls.SetCellText(1, nRow, double2str(nGrpQty), FMT_NUMBER1, true);
		nTotal += nGrpQty;
		nGrpQty = 0;
		nRow++;
	}
	if (nTotal > 0)
	{
		xls.SetCellText(0, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
		xls.SetCellText(1, nRow, double2str(nTotal), FMT_NUMBER1, true);
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\THONG KE BENH NHAN NOI NGOAI CHUYEN KHOA.xls"));
} 
void CEMPatientRatebyWardType::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CEMPatientRatebyWardType::OnAddEMStatisticsGeneralPatientReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMPatientRatebyWardType::OnEditEMStatisticsGeneralPatientReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMPatientRatebyWardType::OnDeleteEMStatisticsGeneralPatientReport(){
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
 		OnCancelEMStatisticsGeneralPatientReport(); 
 	}
	return 0;
}
int CEMPatientRatebyWardType::OnSaveEMStatisticsGeneralPatientReport(){
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
 		//OnEMStatisticsGeneralPatientReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMPatientRatebyWardType::OnCancelEMStatisticsGeneralPatientReport(){
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
int CEMPatientRatebyWardType::OnEMStatisticsGeneralPatientReportListLoadData(){
	return 0;
}
CString CEMPatientRatebyWardType::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND he_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	szWhere.AppendFormat(_T(" AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szSQL.Format(_T(" SELECT") \
				_T("   hrl_type,") \
				_T("   hrl_name as nme,") \
				_T("   Count(he_docno) as cnt") \
				_T(" FROM   hms_exam") \
				_T(" LEFT JOIN hms_roomlist ON (he_deptid=hrl_deptid AND he_roomid=hrl_id)") \
				_T(" WHERE  he_status IN ('P', 'T')") \
				_T("        %s") \
				_T(" GROUP  BY hrl_name,") \
				_T("           hrl_type") \
				_T(" ORDER  BY hrl_type,") \
				_T("           cast(substr(hrl_name, 1, 2) as integer)"), szWhere);
	return szSQL;
}
