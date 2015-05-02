#include "stdafx.h"
#include "EMPatientStatbyRoom.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMPatientStatbyRoom *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMPatientStatbyRoom* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMPatientStatbyRoom *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMPatientStatbyRoom *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyRoom *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMPatientStatbyRoom *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMPatientStatbyRoom *pVw = (CEMPatientStatbyRoom *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMPatientStatbyRoom *pVw = (CEMPatientStatbyRoom *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CEMPatientStatbyRoom *pVw = (CEMPatientStatbyRoom *)pWnd;
	pVw->OnCloseSelect();
} 
CEMPatientStatbyRoom::CEMPatientStatbyRoom(CWnd *pParent)
{
	m_szTitle = _T("General Information In Exam Room");
	m_nDlgWidth = 500;
	m_nDlgHeight = 125;
	SetDefaultValues();
}
CEMPatientStatbyRoom::~CEMPatientStatbyRoom(){
}
void CEMPatientStatbyRoom::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 490, 90);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 245, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 250, 30, 330, 55);
	m_wndReportPeriod.Create(this,335, 30, 485, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 59, 90, 84);
	m_wndFromDate.Create(this,95, 59, 245, 84); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 60, 330, 85);
	m_wndToDate.Create(this,335, 60, 485, 85); 
	//m_wndPrintPreview.Create(this, _T("Print Pre&view"), 300, 95, 410, 120);
	m_wndExport.Create(this, _T("&Export"), 415, 95, 490, 120);
}
void CEMPatientStatbyRoom::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	EnableButtons(TRUE, 0, 1, -1);
	m_wndYear.SetLimitText(16);
	//m_wndYear.SetCheckValue(true);
	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);
}
void CEMPatientStatbyRoom::OnSetWindowEvents(){
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
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = szSysDate + _T("00:00");
	//m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate = szSysDate + _T("23:59");
	UpdateData(false);
}
void CEMPatientStatbyRoom::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
}
void CEMPatientStatbyRoom::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);
}
void CEMPatientStatbyRoom::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMPatientStatbyRoom::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMPatientStatbyRoom::SetMode(int nMode){
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
/*void CEMPatientStatbyRoom::OnYearChange(){
	
} */
/*void CEMPatientStatbyRoom::OnYearSetfocus(){
	
} */
/*void CEMPatientStatbyRoom::OnYearKillfocus(){
	
} */
int CEMPatientStatbyRoom::OnYearCheckValue(){
	return 0;
} 
void CEMPatientStatbyRoom::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMPatientStatbyRoom::OnReportPeriodSelendok(){
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
/*void CEMPatientStatbyRoom::OnReportPeriodSetfocus(){
	
}*/
/*void CEMPatientStatbyRoom::OnReportPeriodKillfocus(){
	
}*/
long CEMPatientStatbyRoom::OnReportPeriodLoadData(){
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
/*void CEMPatientStatbyRoom::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMPatientStatbyRoom::OnFromDateChange(){
	
} */
/*void CEMPatientStatbyRoom::OnFromDateSetfocus(){
	
} */
/*void CEMPatientStatbyRoom::OnFromDateKillfocus(){
	
} */
int CEMPatientStatbyRoom::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMPatientStatbyRoom::OnToDateChange(){
	
} */
/*void CEMPatientStatbyRoom::OnToDateSetfocus(){
	
} */
/*void CEMPatientStatbyRoom::OnToDateKillfocus(){
	
} */
int CEMPatientStatbyRoom::OnToDateCheckValue(){
	return 0;
} 
void CEMPatientStatbyRoom::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	CReport rpt;
	CString tmpStr, szSQL, szSQL1, szWhere;
	CRecord rs(&pMF->m_db);
	CRecord rs1(&pMF->m_db);
	UpdateData(true);

	if(!rpt.Init(_T("Reports/HMS/RM_TONGHOPTHONGTINKHAMBENHTAICACPHONGKHAM.RPT")) ) return ;

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd||hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	int nTong = 0;
	int nIndex = 1;
	CString tmpStr1;
	CReportSection* rptDetail  ;
	long double  nTotal[4];
	for (int i = 1 ; i<= 3; i++)
	{
		nTotal[i] = 0.0;
	}

	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("4"), tmpStr);
		
		rs.GetValue(_T("name"),tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		
		rs.GetValue(_T("reg"), tmpStr);
		nTotal[1] += ToLong(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("examed"), tmpStr);
		nTotal[2] += ToLong(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		
		rs.GetValue(_T("terminated"), tmpStr);
		nTotal[3] += ToLong(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.MoveNext();		
	}

	tmpStr.Format(_T("%d"), nTong);
	rpt.GetReportFooter()->SetValue(_T("sum3"), tmpStr);
	CString szSum;
	for (int i = 1 ; i<= 3; i++)
	{
		szSum.Format(_T("s%d"), i+5);
		tmpStr.Format(_T("%.2f"), nTotal[i]);
		rpt.GetReportFooter()->SetValue(szSum, tmpStr);
	}
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
	
} 
void CEMPatientStatbyRoom::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szWhere;
	int nOldType = 0, nNewType = 0;
	double nGrpQty = 0;
	CRecord rs(&pMF->m_db);
	UpdateData(true);
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 6);
	xls.SetColumnWidth(1, 24);
	xls.SetColumnWidth(2, 14);
	xls.SetColumnWidth(3, 14);
	xls.SetColumnWidth(4, 14);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellMergedColumns(nCol, nRow + 3, 5);
	xls.SetCellText(nCol, nRow + 3, _T("T\x1ED4NG H\x1EE2P TH\xD4NG TIN KH\xC1M \x42\x1EC6NH T\x1EA0I \x43\xC1\x43 PH\xD2NG KH\xC1M"), FMT_TEXT | FMT_CENTER, true, 14);
	xls.SetCellMergedColumns(nCol, nRow + 4, 8);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol, nRow + 6, _T("STT"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 1, nRow + 6, _T("Ph\xF2ng kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 2, nRow + 6, _T("TS Ti\x1EBFp \x111\xF3n"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 3, nRow + 6, _T("TS \x111\xE3 kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol + 4, nRow + 6, _T("TS \x63h\x1B0\x61 kh\xE1m"), FMT_TEXT | FMT_CENTER, true, 11);

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	int nIndex = 1;
	
	double nTotal[3], nGrpTotal[3];
	for (int i = 0 ; i < 3; i++)
	{
		nTotal[i] = 0;
		nGrpTotal[i] = 0;
	}
	nRow = 7;
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("hrl_type"), nNewType);
		if (nNewType != nOldType)
		{
			if ((nNewType == 2 || nNewType == 3) || (nNewType > 3 && nOldType <=3) )
			{
				nGrpQty = nGrpTotal[0] + nGrpTotal[1] + nGrpTotal[2];
				if (nGrpQty > 0)
				{
					xls.SetCellText(1, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
					for (int i = 0; i < 3; i++)
					{
						xls.SetCellText(i + 2, nRow, double2str(nGrpTotal[i]), FMT_NUMBER1);
						nTotal[i] += nGrpTotal[i];
						nGrpTotal[i] = 0;
					}
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
					xls.SetCellText(1, nRow, tmpStr, 4098, true);
				nOldType = nNewType;
				nRow++;
				nIndex = 1;
			}
		}
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("name"),tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT);
	
		rs.GetValue(_T("reg"),tmpStr);
		nGrpTotal[0] += ToLong(tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("examed"), tmpStr);
		nGrpTotal[1] += ToLong(tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("remain"),tmpStr);
		nGrpTotal[2] += ToLong(tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1);
		rs.MoveNext();
		nRow++;
	}
	nGrpQty = nGrpTotal[0] + nGrpTotal[1] + nGrpTotal[2];
	if (nGrpQty > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
		for (int i = 0; i < 3; i++)
		{
			xls.SetCellText(i + 2, nRow, double2str(nGrpTotal[i]), FMT_NUMBER1);
			nTotal[i] += nGrpTotal[i];
			nGrpTotal[i] = 0;
		}
		nRow++;
	}	
	xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true, 11);
	for (int i = 0; i < 3; i++)
	{
		tmpStr.Format(_T("%.0f"), nTotal[i]); 
		xls.SetCellText(nCol + i + 2, nRow, tmpStr, FMT_NUMBER1, true, 11);

	}
	xls.Save(_T("Exports\\TONG HOP THONG TIN KHAM BENH TAI PHONG KHAM.xls"));
} 
void CEMPatientStatbyRoom::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}

CString CEMPatientStatbyRoom::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString  szSQL;
	CString szWhere;
	szSQL.Format(_T(" SELECT hrl_type, hrl_id, hrl_name as name,") \
				_T(" 	sum(reg) as reg,") \
				_T(" 	sum(examed) as examed,") \
				_T(" 	sum(remain) as remain") \
				_T("  FROM ") \
				_T("  ( ") \
				_T("  SELECT") \
				_T(" 	he_deptid,") \
				_T(" 	he_roomid,") \
				_T(" 	1 as reg,") \
				_T(" 	case when he_status IN ('P', 'T') then 1 else 0 end as examed,") \
				_T(" 	case when he_status NOT IN ('P', 'T') then 1 else 0 end as remain") \
				_T("  FROM") \
				_T(" 	hms_exam") \
				_T("  WHERE he_examdate between cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
				_T("  ) Tbl") \
				_T("  LEFT JOIN hms_roomlist ON(hrl_deptid=he_deptid AND hrl_id=he_roomid)") \
				_T("  WHERE hrl_id > 0 AND he_deptid = '%s'") \
				_T("  GROUP BY hrl_type,hrl_id, hrl_name") \
				_T("  ORDER BY hrl_type,cast(substr(hrl_name, 1, 2) as integer)"), m_szFromDate, m_szToDate, pMF->m_szDept);
	return szSQL;
}