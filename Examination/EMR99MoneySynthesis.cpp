#include "stdafx.h"
#include "EMR99MoneySynthesis.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMR99MoneySynthesis *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMR99MoneySynthesis* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMR99MoneySynthesis *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMR99MoneySynthesis *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMR99MoneySynthesis *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMR99MoneySynthesis *)pWnd)->OnToDateCheckValue();
} 
static void _OnByDaySelectFnc(CWnd *pWnd){
	 ((CEMR99MoneySynthesis*)pWnd)->OnByDaySelect();
}
static void _OnByDeptSelectFnc(CWnd *pWnd){
	 ((CEMR99MoneySynthesis*)pWnd)->OnByDeptSelect();
}
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMR99MoneySynthesis *pVw = (CEMR99MoneySynthesis *)pWnd;
	pVw->OnExportSelect();
} 
CEMR99MoneySynthesis::CEMR99MoneySynthesis(CWnd *pParent){

	m_nDlgWidth = 500;
	m_nDlgHeight = 125;
	SetDefaultValues();
}
CEMR99MoneySynthesis::~CEMR99MoneySynthesis(){
}
void CEMR99MoneySynthesis::OnCreateComponents(){
	m_wndR99MoneySynthesis.Create(this, _T("R99 Money Synthesis"), 5, 5, 490, 90);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 245, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 250, 30, 330, 55);
	m_wndReportPeriod.Create(this,335, 30, 485, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 245, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 60, 330, 85);
	m_wndToDate.Create(this,335, 60, 485, 85); 
	m_wndByDay.Create(this, _T("By Day"), 5, 95, 110, 120);
	m_wndByDept.Create(this, _T("By Dept"), 115, 95, 220, 120);
	m_wndExport.Create(this, _T("&Export"), 410, 95, 490, 120);

}
void CEMR99MoneySynthesis::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(16);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);


	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CEMR99MoneySynthesis::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
	m_wndByDay.SetEvent(WE_CLICK, _OnByDaySelectFnc);
	m_wndByDept.SetEvent(WE_CLICK, _OnByDeptSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	m_bByDay = TRUE;
	UpdateData(FALSE);

}
void CEMR99MoneySynthesis::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_Check(pDX, m_wndByDay.GetDlgCtrlID(), m_bByDay);
	DDX_Check(pDX, m_wndByDept.GetDlgCtrlID(), m_bByDept);

}
void CEMR99MoneySynthesis::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_bByDay=FALSE;
	m_bByDept=FALSE;

}
int CEMR99MoneySynthesis::SetMode(int nMode){
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
/*void CEMR99MoneySynthesis::OnYearChange(){
	
} */
/*void CEMR99MoneySynthesis::OnYearSetfocus(){
	
} */
/*void CEMR99MoneySynthesis::OnYearKillfocus(){
	
} */
int CEMR99MoneySynthesis::OnYearCheckValue(){
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
void CEMR99MoneySynthesis::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMR99MoneySynthesis::OnReportPeriodSelendok(){
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
	UpdateData(FALSE);
}
/*void CEMR99MoneySynthesis::OnReportPeriodSetfocus(){
	
}*/
/*void CEMR99MoneySynthesis::OnReportPeriodKillfocus(){
	
}*/
long CEMR99MoneySynthesis::OnReportPeriodLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
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
/*void CEMR99MoneySynthesis::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMR99MoneySynthesis::OnFromDateChange(){
	
} */
/*void CEMR99MoneySynthesis::OnFromDateSetfocus(){
	
} */
/*void CEMR99MoneySynthesis::OnFromDateKillfocus(){
	
} */
int CEMR99MoneySynthesis::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMR99MoneySynthesis::OnToDateChange(){
	
} */
/*void CEMR99MoneySynthesis::OnToDateSetfocus(){
	
} */
/*void CEMR99MoneySynthesis::OnToDateKillfocus(){
	
} */
int CEMR99MoneySynthesis::OnToDateCheckValue(){
	return 0;
} 
void CEMR99MoneySynthesis::OnByDaySelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndByDept.SetCheck(false);
}
void CEMR99MoneySynthesis::OnByDeptSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndByDay.SetCheck(false);
}
void CEMR99MoneySynthesis::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	CStringArray arrCol;
	int nRow = 0, nCol = 0, nIdx = 1;
	double nTmp = 0, nCount = 0;
	long double nCost = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data!"));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 15);
	xls.SetColumnWidth(3, 15);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 3);
	xls.SetCellMergedColumns(0, 3, 3);
	xls.SetCellText(0, 0, pMF->m_CompanyInfo.sc_pname, 4098, true);
	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_name, 4098, true);
	xls.SetCellText(0, 2, _T("T\x1ED4NG H\x1EE2P TI\x1EC0N THU PH\xD2NG Y\xCAU \x43\x1EA6U"), 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd, ddmmyyyy),
				  CDateTime::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, FMT_TEXT | FMT_CENTER);
	arrCol.Add(_T("STT"));
	if (m_bByDay)
	{
		arrCol.Add(_T("Ng\xE0y"));
		arrCol.Add(_T("S\x1ED1 \x42N"));
	}
	else
	{
		arrCol.Add(_T("Kho\x61 th\x1EF1\x63 hi\x1EC7n"));
	}
	arrCol.Add(_T("S\x1ED1 ti\x1EC1n"));
	for (int i = 0; i < arrCol.GetCount(); i++)
		xls.SetCellText(i, 4, arrCol.GetAt(i), 4098, true);
	nRow = 5;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		if (m_bByDay)
		{
			xls.SetCellText(1, nRow, CDate::Convert(rs.GetValue(_T("hfe_date")), yyyymmdd, ddmmyyyy));
			rs.GetValue(_T("patientcount"), nTmp);
			xls.SetCellText(2, nRow, double2str(nTmp), FMT_NUMBER1);
			nCount += nTmp;
			rs.GetValue(_T("cost"), nTmp);
			xls.SetCellText(3, nRow, double2str(nTmp), FMT_NUMBER1);
			nCost += nTmp;
		}
		else
		{
			xls.SetCellText(1, nRow, rs.GetValue(_T("pdeptid")), FMT_TEXT);
			rs.GetValue(_T("cost"), nTmp);
			xls.SetCellText(2, nRow, double2str(nTmp), FMT_NUMBER1);
			nCost += nTmp;
		}
		nRow++;
		rs.MoveNext();
	}
	if (nCost > 0)
	{
		xls.SetCellText(1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
		if (m_bByDay)
		{
			xls.SetCellText(2, nRow, double2str(nCount), FMT_NUMBER1);
			tmpStr.Format(_T("%f"), nCost);
			xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);
		}
		else
		{
			tmpStr.Format(_T("%f"), nCost);
			xls.SetCellText(2, nRow, tmpStr, FMT_NUMBER1);
		}
	}
	xls.Save(_T("Exports\\Tong hop tien phong yeu cau.xls"));
} 

CString CEMR99MoneySynthesis::GetQueryString(){
	CString szSQL, szWhere, szGroupBy, szSelect, szOrderBy;
	if (m_bByDay)
	{
		szSelect.Format(_T(" hfe_date, count(distinct hpc_docno) patientcount,"));
		szGroupBy.Format(_T(" hfe_date"));
		szOrderBy.Format(_T(" hfe_date"));
	}
	else
	{
		szSelect.Format(_T(" hpc_pdeptid pdeptid,"));
		szGroupBy.Format(_T(" hpc_pdeptid"));
		szOrderBy.Format(_T("hpc_pdeptid"));
	}
	szWhere.Format(_T(" AND i.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') AND hfe_category <> 'Y'"), m_szFromDate, m_szToDate);
	szSQL.Format(_T(" SELECT %s ") \
	_T("        SUM(hfe_cost)              cost ") \
	_T(" FROM   (SELECT    hpc_docno, ") \
	_T("				   trunc(i.hfe_date) hfe_date,") \
	_T("				   hfl_deptid hpc_pdeptid,") \
	_T("                   hfe_cost ") \
	_T("         FROM      hmsv_paraclinic ") \
	_T("         LEFT JOIN hms_fee f ON ( hfe_type IN ( 'P', 'T' ) ") \
	_T("                                  AND hpc_docno = hfe_docno ") \
	_T("                                  AND hpc_orderid = hfe_orderid ) ") \
	_T("		 LEFT JOIN hms_fee_list ON (hfl_feeid = hfe_itemid)") \
	_T("         LEFT JOIN hms_fee_invoice i ON ( i.hfe_docno = f.hfe_docno ") \
	_T("                                          AND i.hfe_invoiceno = f.hfe_invoiceno ) ") \
	_T("         WHERE     i.hfe_status = 'P' ") \
	_T("               AND f.hfe_status = 'P' ") \
	_T("               AND hpc_roomid = 99 %s") \
	_T("         UNION ALL ") \
	_T("         SELECT    ho_docno, ") \
	_T("				   trunc(i.hfe_date) hfe_date,") \
	_T("				   ho_pdeptid,") \
	_T("                   hfe_cost ") \
	_T("         FROM      hms_operation ") \
	_T("         LEFT JOIN hms_fee f ON ( hfe_type IN ('O') ") \
	_T("                                  AND ho_docno = hfe_docno ") \
	_T("                                  AND ho_idx = hfe_orderid ) ") \
	_T("         LEFT JOIN hms_fee_invoice i ON ( i.hfe_docno = f.hfe_docno ") \
	_T("                                          AND i.hfe_invoiceno = f.hfe_invoiceno ) ") \
	_T("         WHERE     i.hfe_status = 'P' ") \
	_T("               AND f.hfe_status = 'P' ") \
	_T("               AND ho_roomid = 99 %s") \
	_T("         UNION ALL ") \
	_T("         SELECT    hpo_docno, ") \
	_T("				   trunc(i.hfe_date) hfe_date,") \
	_T("				   hpo_deptid,") \
	_T("                   hfe_cost ") \
	_T("         FROM      hms_pharmaorder ") \
	_T("         LEFT JOIN hms_fee f ON ( hpo_docno = hfe_docno ") \
	_T("                                  AND hpo_orderid = hfe_orderid ) ") \
	_T("         LEFT JOIN hms_fee_invoice i ON ( i.hfe_docno = f.hfe_docno ") \
	_T("                                          AND i.hfe_invoiceno = f.hfe_invoiceno ) ") \
	_T("         WHERE     i.hfe_status = 'P' ") \
	_T("               AND f.hfe_status = 'P' ") \
	_T("               AND hpo_roomid = 99 %s)") \
	_T(" GROUP  BY %s ") \
	_T(" ORDER  BY %s "), szSelect, szWhere, szWhere, szWhere, szGroupBy, szOrderBy);_fmsg(_T("%s"), szSQL);
	return szSQL;
}