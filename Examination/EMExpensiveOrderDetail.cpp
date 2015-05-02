#include "stdafx.h"
#include "EMExpensiveOrderDetail.h"
#include "HMSMainFrame.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMExpensiveOrderDetail *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMExpensiveOrderDetail* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMExpensiveOrderDetail *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExpensiveOrderDetail *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExpensiveOrderDetail *)pWnd)->OnToDateCheckValue();
} 
/*static void _OnTotalCostChangeFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnTotalCostChange();
} */
/*static void _OnTotalCostSetfocusFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnTotalCostSetfocus();} */ 
/*static void _OnTotalCostKillfocusFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnTotalCostKillfocus();
} */
static int _OnTotalCostCheckValueFnc(CWnd *pWnd){
	return ((CEMExpensiveOrderDetail *)pWnd)->OnTotalCostCheckValue();
} 
static void _OnRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMExpensiveOrderDetail* )pWnd)->OnRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnRoomSelendokFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnRoomSelendok();
}
/*static void _OnRoomSetfocusFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnRoomKillfocus();
}*/
/*static void _OnRoomKillfocusFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnRoomKillfocus();
}*/
static long _OnRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMExpensiveOrderDetail *)pWnd)->OnRoomLoadData();
}
/*static void _OnRoomAddNewFnc(CWnd *pWnd){
	((CEMExpensiveOrderDetail *)pWnd)->OnRoomAddNew();
}*/
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMExpensiveOrderDetail *pVw = (CEMExpensiveOrderDetail *)pWnd;
	pVw->OnExportSelect();
} 
CEMExpensiveOrderDetail::CEMExpensiveOrderDetail(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CEMExpensiveOrderDetail::~CEMExpensiveOrderDetail(){
}
void CEMExpensiveOrderDetail::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 435, 150);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 95, 55);
	m_wndYear.Create(this,100, 30, 220, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 225, 30, 305, 55);
	m_wndReportPeriod.Create(this,310, 30, 430, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 95, 85);
	m_wndFromDate.Create(this,100, 60, 220, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 225, 60, 305, 85);
	m_wndToDate.Create(this,310, 60, 430, 85); 
	m_wndTotalCostLabel.Create(this, _T("Min Total Cost"), 10, 90, 95, 115);
	m_wndTotalCost.Create(this,100, 90, 430, 115); 
	m_wndRoomLabel.Create(this, _T("Room"), 10, 120, 95, 145);
	m_wndRoom.Create(this,100, 120, 430, 145); 
	m_wndExport.Create(this, _T("&Export"), 335, 155, 435, 180);

}
void CEMExpensiveOrderDetail::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndTotalCost.SetLimitText(1024);
	m_wndTotalCost.SetCheckValue(true);
	m_wndRoom.SetCheckValue(true);
	m_wndRoom.LimitText(1024);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Name"), CFMT_TEXT, 80);

	m_wndRoom.InsertColumn(0, _T("ID"), CFMT_TEXT|CFMT_RIGHT, 80);
	m_wndRoom.InsertColumn(1, _T("Name"), CFMT_TEXT, 220);
}
void CEMExpensiveOrderDetail::OnSetWindowEvents(){
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
	//m_wndTotalCost.SetEvent(WE_CHANGE, _OnTotalCostChangeFnc);
	//m_wndTotalCost.SetEvent(WE_SETFOCUS, _OnTotalCostSetfocusFnc);
	//m_wndTotalCost.SetEvent(WE_KILLFOCUS, _OnTotalCostKillfocusFnc);
	m_wndTotalCost.SetEvent(WE_CHECKVALUE, _OnTotalCostCheckValueFnc);
	m_wndRoom.SetEvent(WE_SELENDOK, _OnRoomSelendokFnc);
	//m_wndRoom.SetEvent(WE_SETFOCUS, _OnRoomSetfocusFnc);
	//m_wndRoom.SetEvent(WE_KILLFOCUS, _OnRoomKillfocusFnc);
	m_wndRoom.SetEvent(WE_SELCHANGE, _OnRoomSelectChangeFnc);
	m_wndRoom.SetEvent(WE_LOADDATA, _OnRoomLoadDataFnc);
	//m_wndRoom.SetEvent(WE_ADDNEW, _OnRoomAddNewFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	CString szCurDate = pMF->GetSysDate();
	CDate dte;
	dte.ParseDate(szCurDate);
	m_szFromDate = m_szToDate = szCurDate;
	m_szToDate += _T("23:59");
	m_nYear = ToInt(szCurDate.Left(4));
	m_szReportPeriodKey = int2str(dte.GetMonth());
	UpdateData(FALSE);

}
void CEMExpensiveOrderDetail::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_Text(pDX, m_wndTotalCost.GetDlgCtrlID(), m_nTotalCost);
	DDX_TextEx(pDX, m_wndRoom.GetDlgCtrlID(), m_szRoomKey);

}
void CEMExpensiveOrderDetail::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_nTotalCost=0;
	m_szRoomKey.Empty();

}
int CEMExpensiveOrderDetail::SetMode(int nMode){
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
/*void CEMExpensiveOrderDetail::OnYearChange(){
	
} */
/*void CEMExpensiveOrderDetail::OnYearSetfocus(){
	
} */
/*void CEMExpensiveOrderDetail::OnYearKillfocus(){
	
} */
int CEMExpensiveOrderDetail::OnYearCheckValue(){
	return 0;
} 
void CEMExpensiveOrderDetail::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMExpensiveOrderDetail::OnReportPeriodSelendok(){
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
/*void CEMExpensiveOrderDetail::OnReportPeriodSetfocus(){
	
}*/
/*void CEMExpensiveOrderDetail::OnReportPeriodKillfocus(){
	
}*/
long CEMExpensiveOrderDetail::OnReportPeriodLoadData(){
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
/*void CEMExpensiveOrderDetail::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CEMExpensiveOrderDetail::OnFromDateChange(){
	
} */
/*void CEMExpensiveOrderDetail::OnFromDateSetfocus(){
	
} */
/*void CEMExpensiveOrderDetail::OnFromDateKillfocus(){
	
} */
int CEMExpensiveOrderDetail::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMExpensiveOrderDetail::OnToDateChange(){
	
} */
/*void CEMExpensiveOrderDetail::OnToDateSetfocus(){
	
} */
/*void CEMExpensiveOrderDetail::OnToDateKillfocus(){
	
} */
int CEMExpensiveOrderDetail::OnToDateCheckValue(){
	return 0;
} 
/*void CEMExpensiveOrderDetail::OnTotalCostChange(){
	
} */
/*void CEMExpensiveOrderDetail::OnTotalCostSetfocus(){
	
} */
/*void CEMExpensiveOrderDetail::OnTotalCostKillfocus(){
	
} */
int CEMExpensiveOrderDetail::OnTotalCostCheckValue(){
	return 0;
} 
void CEMExpensiveOrderDetail::OnRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMExpensiveOrderDetail::OnRoomSelendok(){
	 
}
/*void CEMExpensiveOrderDetail::OnRoomSetfocus(){
	
}*/
/*void CEMExpensiveOrderDetail::OnRoomKillfocus(){
	
}*/
long CEMExpensiveOrderDetail::OnRoomLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();
	if(m_wndRoom.IsSearchKey() && !m_szRoomKey.IsEmpty())
	{
		szWhere.Format(_T("and hrl_id = %d"), ToInt(m_szRoomKey));
	}
	m_wndRoom.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("select hrl_name as name, hrl_id as id from hms_roomlist where hrl_deptid = '%s' %s Order by id"),pMF->GetCurrentDepartmentID(), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndRoom.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMExpensiveOrderDetail::OnRoomAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMExpensiveOrderDetail::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CExcel xls;
	CString szSQL, tmpStr, szOldRoom, szNewRoom;
	double nTmp = 0, nGrpAmt = 0;
	long double nTotal = 0;
	int nIdx = 1, nRow = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 10);
	xls.SetColumnWidth(1, 20);
	xls.SetColumnWidth(2, 15);
	xls.SetColumnWidth(4, 20);
	xls.SetColumnWidth(5, 15);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 4);
	xls.SetCellMergedColumns(0, 3, 4);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr = _T("\x43HI TI\x1EBET \x110\x1A0N THU\x1ED0\x43 \x110\x1EAET TI\x1EC0N");
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	xls.SetCellText(0, 4, _T("STT"), 4098, true);
	xls.SetCellText(1, 4, _T("\x42\xE1\x63 s\x1EF9"), 4098, true);
	xls.SetCellText(2, 4, _T("Ng\xE0y k\xEA"), 4098, true);
	xls.SetCellText(3, 4, _T("S\x1ED1 HS"), 4098, true);
	xls.SetCellText(4, 4, _T("T\xEAn \x42N"), 4098, true);
	xls.SetCellText(5, 4, _T("S\x1ED1 ti\x1EC1n"), 4098, true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("hrl_id"), szNewRoom);
		if (szNewRoom != szOldRoom)
		{
			if (nGrpAmt > 0)
			{
				xls.SetCellText(4, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
				xls.SetCellText(5, nRow, double2str(nGrpAmt), FMT_NUMBER1, true);
				nTotal += nGrpAmt;
				nGrpAmt = 0;
				nRow++;
			}
			xls.SetCellMergedColumns(0, nRow, 5);
			xls.SetCellText(0, nRow, rs.GetValue(_T("room_name")), FMT_TEXT, true);
			nIdx = 1;
			szOldRoom = szNewRoom;
			nRow++;
		}
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("doctor")), FMT_TEXT);
		tmpStr = CDate::Convert(rs.GetValue(_T("order_date")), yyyymmdd, ddmmyyyy);
		xls.SetCellText(2, nRow, tmpStr, FMT_TEXT);
		xls.SetCellText(3, nRow, rs.GetValue(_T("doc_no")), FMT_TEXT);
		xls.SetCellText(4, nRow, rs.GetValue(_T("patient_name")), FMT_TEXT);
		rs.GetValue(_T("order_amount"), nTmp);
		nGrpAmt += nTmp;
		xls.SetCellText(5, nRow, double2str(nTmp), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nGrpAmt > 0)
	{
		xls.SetCellText(4, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
		xls.SetCellText(5, nRow, double2str(nGrpAmt), FMT_NUMBER1, true);
		nTotal += nGrpAmt;
		nGrpAmt = 0;
		nRow++;
	}
	if (nTotal > 0)
	{
		xls.SetCellText(4, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
		tmpStr.Format(_T("%f"), nTotal);
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\Chi tiet don thuoc dat tien.xls"));

} 

CString CEMExpensiveOrderDetail::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szWhere, szCondition;
	szWhere.Format(_T(" AND hpo_deptid = '%s' AND hpo_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
				, pMF->GetCurrentDepartmentID(), m_szFromDate, m_szToDate);
	if (!m_szRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_roomid = %s"), m_szRoomKey);
	if (m_nTotalCost > 0)
		szCondition.AppendFormat(_T(" HAVING SUM(hpol_unitprice * hpol_qtyissue) > %ld"), m_nTotalCost);
	szSQL.Format(_T(" SELECT    hrl_id, hrl_name room_name, hpo_orderdate order_date, ") \
	_T("           hpo_docno doc_no, get_username(hpo_doctor) doctor,") \
	_T("           Get_patientname(hpo_docno) patient_name, ") \
	_T("           SUM(hpol_unitprice * hpol_qtyissue) order_amount ") \
	_T(" FROM      hms_pharmaorder ") \
	_T(" LEFT JOIN hms_pharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
	_T(" LEFT JOIN hms_roomlist ON (hpo_roomid = hrl_id AND hrl_deptid = '%s')") \
	_T(" WHERE     hpo_orderstatus IN ( 'S', 'A' ) %s") \
	_T(" GROUP     BY hpo_orderdate, hpo_docno, hrl_id, hpo_doctor, hrl_name %s") \
	_T(" ORDER BY cast (substr(hrl_name, 1, 2) as integer), hpo_doctor"), pMF->GetCurrentDepartmentID(), szWhere, szCondition);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}