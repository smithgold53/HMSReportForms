#include "stdafx.h"
#include "EMDiseasebyMachine.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include "ReportDocument.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CEMDiseasebyMachine *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMDiseasebyMachine* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CEMDiseasebyMachine *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMDiseasebyMachine *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMDiseasebyMachine *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMDiseasebyMachine *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMDiseasebyMachine *pVw = (CEMDiseasebyMachine *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMDiseasebyMachine *pVw = (CEMDiseasebyMachine *)pWnd;
	pVw->OnExportSelect();
} 
CEMDiseasebyMachine::CEMDiseasebyMachine(CWnd *pParent){

	m_nDlgWidth = 500;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CEMDiseasebyMachine::~CEMDiseasebyMachine(){
}
void CEMDiseasebyMachine::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 490, 90);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 245, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 250, 30, 330, 55);
	m_wndReportPeriod.Create(this,335, 30, 485, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 245, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 60, 330, 85);
	m_wndToDate.Create(this,335, 60, 485, 85); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 300, 95, 410, 120);
	m_wndExport.Create(this, _T("&Export"), 415, 95, 490, 120);

}
void CEMDiseasebyMachine::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), FMT_INTEGER, 30);
	m_wndReportPeriod.InsertColumn(1, _T("Name"), FMT_TEXT, 80);

}
void CEMDiseasebyMachine::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
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
	CDate dte;
	dte.ParseDate(pMF->GetSysDate());
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	m_nYear = dte.GetYear();
	m_szReportPeriodKey.Format(_T("%d"), dte.GetMonth());
	UpdateData(false);

}
void CEMDiseasebyMachine::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CEMDiseasebyMachine::SetDefaultValues(){

	m_nYear = 0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMDiseasebyMachine::SetMode(int nMode){
 		int nOldMode = GetMode();
 		CGuiView::SetMode(nMode);
 		CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
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
/*void CEMDiseasebyMachine::OnYearChange(){
	
} */
/*void CEMDiseasebyMachine::OnYearSetfocus(){
	
} */
/*void CEMDiseasebyMachine::OnYearKillfocus(){
	
} */
int CEMDiseasebyMachine::OnYearCheckValue(){
	return 0;
} 
void CEMDiseasebyMachine::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

} 
void CEMDiseasebyMachine::OnReportPeriodSelendok(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	pMF->OnReportPeriodSelendok(&m_szFromDate, &m_szToDate, m_szReportPeriodKey);
	UpdateData(false);	 
}
/*void CEMDiseasebyMachine::OnReportPeriodSetfocus(){
	
}*/
/*void CEMDiseasebyMachine::OnReportPeriodKillfocus(){
	
}*/
long CEMDiseasebyMachine::OnReportPeriodLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadReportPeriodList(&m_wndReportPeriod, m_szReportPeriodKey);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && !m_szReportPeriodKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szReportPeriodKey
};
	m_wndReportPeriod.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndReportPeriod.AddItems(
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CEMDiseasebyMachine::OnReportPeriodAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
/*void CEMDiseasebyMachine::OnFromDateChange(){
	
} */
/*void CEMDiseasebyMachine::OnFromDateSetfocus(){
	
} */
/*void CEMDiseasebyMachine::OnFromDateKillfocus(){
	
} */
int CEMDiseasebyMachine::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMDiseasebyMachine::OnToDateChange(){
	
} */
/*void CEMDiseasebyMachine::OnToDateSetfocus(){
	
} */
/*void CEMDiseasebyMachine::OnToDateKillfocus(){
	
} */
int CEMDiseasebyMachine::OnToDateCheckValue(){
	return 0;
} 
void CEMDiseasebyMachine::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CReportSection* rptDetail;
	if (!rpt.Init(_T("Reports\\HMS\\c15.RPT")))
		return;
	CString szSQL, szWhere;
	CString tmpStr, szPos;
	int nTemp = 0, nTmp = 0, j = 0, nExam_Count = 0;
	double nAmt = 0, nTotal = 0;
	CStringArray arrMac;
	CRecord rs(&pMF->m_db);
	//BN kham
	szSQL.Format(_T(" SELECT") \
				_T("   Count(distinct hd_docno) as examqty") \
				_T(" FROM   hms_doc") \
				_T("        LEFT JOIN hms_exam ON (hd_docno=he_docno)") \
				_T(" WHERE  he_roomid=50") \
				_T("        AND he_status IN ('T', 'P') ") \
				_T("		AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	rs.ExecSQL(szSQL);
	if (rs.GetIntValue() == 0)
		return;
	arrMac.Add(_T("CO"));
	arrMac.Add(_T("F"));
	arrMac.Add(_T("HE"));
	arrMac.Add(_T("Q"));
	arrMac.Add(_T("YA"));
	arrMac.Add(_T("YAG.X"));

	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	rptDetail = rpt.AddDetail();
	rs.GetValue(_T("examqty"), nExam_Count);
	rptDetail->SetValue(_T("1"), int2str(nExam_Count));
	//BN dieu tri
	szSQL.Format(_T(" SELECT") \
				_T("   Count(distinct ho_docno) as treatqty") \
				_T(" FROM   hms_operation") \
				_T(" WHERE  ho_pdeptid='C15'") \
				_T("		AND ho_status <> 'O'") \
				_T("		AND ho_payment = 'P'") \
				_T("		AND ho_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				_T("        AND ho_recordno IS NOT NULL"), m_szFromDate, m_szToDate);
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("treatqty")));
	//BN dieu tri chi tiet
	szSQL.Format(_T(" SELECT") \
				_T("   ho_recordno as recordno,") \
				_T("   Count(DISTINCT ho_docno) as cnt") \
				_T(" FROM   hms_operation") \
				_T(" WHERE  ho_pdeptid='C15'") \
				_T("        AND ho_recordno IS NOT NULL") \
				_T("		AND ho_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				_T("        AND ho_status <> 'O'") \
				_T("		AND ho_payment = 'P'") \
				_T(" GROUP  BY ho_recordno") \
				_T(" ORDER BY ho_recordno"), m_szFromDate, m_szToDate);
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		j = 0;
		for (int i = 0 ; i < arrMac.GetCount(); i++)
		{
			if (arrMac.GetAt(i) == _T("HE"))
				j = i - 1;
			rs.GetValue(_T("recordno"), tmpStr);
			if (tmpStr == arrMac.GetAt(i))
			{
				rs.GetValue(_T("cnt"), nTemp);
				if (tmpStr != _T("HE"))
					rptDetail->SetValue(int2str(j+3), rs.GetValue(_T("cnt")));
				rptDetail->SetValue(int2str(i+8), rs.GetValue(_T("cnt")));
				if (tmpStr == _T("CO") || tmpStr == _T("HE"))
					nTmp += nTemp;
				break;
			}
			j++;
		}
		rs.MoveNext();
	}
	rptDetail->SetValue(_T("3"), nTmp);
	//So tien
	szSQL.Format(_T(" SELECT") \
				_T("   ho_recordno as recordno,") \
				_T("   Sum(hfe_cost) as cost") \
				_T(" FROM   hms_operation") \
				_T("        LEFT JOIN hms_fee ON (hfe_docno=ho_docno") \
				_T("                              AND ho_itemid=hfe_itemid AND hfe_orderid = ho_idx)") \
				_T(" WHERE  ho_pdeptid='C15'") \
				_T("        AND ho_recordno IS NOT NULL") \
				_T("        AND ho_payment='P'") \
				_T("		AND hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				_T(" GROUP  BY ho_recordno") \
				_T(" ORDER BY ho_recordno"), m_szFromDate, m_szToDate);
	j = 0;
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		for (int i = 0 ; i < arrMac.GetCount(); i++)
		{
			rs.GetValue(_T("recordno"), tmpStr);
			if (tmpStr == arrMac.GetAt(i))
			{
				rs.GetValue(_T("cost"), nAmt);
				rptDetail->SetValue(int2str(i+14), double2str(nAmt));
				nTotal += nAmt;
				break;
			}
		}
		rs.MoveNext();
	}	
	nAmt = nExam_Count*30000;
	rptDetail->SetValue(_T("20"), double2str(nAmt));
	nTotal += nAmt;
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), double2str(nTotal));
	rpt.PrintPreview();
} 
void CEMDiseasebyMachine::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	CStringArray arrMac;
	double nAmt = 0, nTotal = 0;
	int nCol = 0, nRow = 0, j = 0, nTemp = 0, nTmp = 0;
	CellFormat cf(&xls);
	cf.SetBold(true);
	cf.SetCellStyle(FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING);
	//xls.SetColumnWidth(0, 1);
	//BN kham
	szSQL.Format(_T(" SELECT") \
				_T("   Count(distinct hd_docno) as examqty") \
				_T(" FROM   hms_doc") \
				_T("        LEFT JOIN hms_exam ON (hd_docno=he_docno)") \
				_T(" WHERE  he_roomid=50") \
				_T("        AND he_status IN ('T', 'P') ") \
				_T("		AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	rs.ExecSQL(szSQL);
	if (rs.GetIntValue() == 0)
		return;
	arrMac.Add(_T("CO"));
	arrMac.Add(_T("F"));
	arrMac.Add(_T("HE"));
	arrMac.Add(_T("Q"));
	arrMac.Add(_T("YA"));
	arrMac.Add(_T("YAG.X"));
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 3, 15);
	xls.SetCellMergedColumns(nCol, nRow + 4, 15);
	xls.SetCellText(nCol, nRow, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 2, pMF->GetDepartmentName(pMF->GetCurrentUser()), FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 3, _T("\x42\xC1O \x43\xC1O M\x1EB6T \x42\x1EC6NH TH\x45O M\xC1Y"), FMT_TEXT | FMT_CENTER, true, 12);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, false, 10);
	//Column
	xls.SetCellMergedRows(nCol, nRow + 5, 2);
	xls.SetCellMergedRows(nCol + 1, nRow + 5, 2);
	xls.SetCellMergedColumns(nCol + 2, nRow + 5, 5);
	xls.SetCellMergedColumns(nCol + 7, nRow + 5, 6);
	xls.SetCellMergedColumns(nCol + 13, nRow + 5, 7);

	xls.SetCellText(nCol, nRow + 5, _T("\x42N Kh\xE1m"), &cf);
	xls.SetCellText(nCol + 1, nRow + 5, _T("BN ÐT"), &cf);
	xls.SetCellText(nCol + 2, nRow + 5, _T("\x43\xE1\x63 m\x1EB7t \x62\x1EC7nh"), &cf);
	xls.SetCellText(nCol + 2, nRow + 6, _T("U nh\x1ECF"), &cf);
	xls.SetCellText(nCol + 3, nRow + 6, _T("S\x1EB9o"), &cf);
	xls.SetCellText(nCol + 4, nRow + 6, _T("RLST Da"), &cf);
	xls.SetCellText(nCol + 5, nRow + 6, _T("U m\xE1u"), &cf);
	xls.SetCellText(nCol + 6, nRow + 6, _T("UM+TL"), &cf);
	xls.SetCellText(nCol + 7, nRow + 5, _T("\x42\x1EC7nh nh\xE2n \x111i\x1EC1u tr\x1ECB"), &cf);
	xls.SetCellText(nCol + 7, nRow + 6, _T("CO2"), &cf);
	xls.SetCellText(nCol + 8, nRow + 6, _T("F"), &cf);
	xls.SetCellText(nCol + 9, nRow + 6, _T("H"), &cf);
	xls.SetCellText(nCol + 10, nRow + 6, _T("Q.S"), &cf);
	xls.SetCellText(nCol + 11, nRow + 6, _T("YA"), &cf);
	xls.SetCellText(nCol + 12, nRow + 6, _T("YAD"), &cf);
	xls.SetCellText(nCol + 13, nRow + 5, _T("S\x1ED1 ti\x1EC1n"), &cf);
	xls.SetCellText(nCol + 13, nRow + 6, _T("CO2"), &cf);
	xls.SetCellText(nCol + 14, nRow + 6, _T("F"), &cf);
	xls.SetCellText(nCol + 15, nRow + 6, _T("HE-NE"), &cf);
	xls.SetCellText(nCol + 16, nRow + 6, _T("Q.S"), &cf);
	xls.SetCellText(nCol + 17, nRow + 6, _T("YAG"), &cf);
	xls.SetCellText(nCol + 18, nRow + 6, _T("YAG-XD"), &cf);
	xls.SetCellText(nCol + 19, nRow + 6, _T("T\x1ED5ng"), &cf);
	nRow = 7;
	xls.SetCellText(nCol, nRow, rs.GetValue(_T("examqty")), FMT_NUMBER1);
	//BN dieu tri
	szSQL.Format(_T(" SELECT") \
				_T("   Count(distinct ho_docno) as treatqty") \
				_T(" FROM   hms_operation") \
				_T(" WHERE  ho_pdeptid='C15'") \
				_T("		AND ho_status <> 'O'") \
				_T("		AND ho_payment = 'P'") \
				_T("		AND ho_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				_T("        AND ho_recordno IS NOT NULL"), m_szFromDate, m_szToDate);
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
		xls.SetCellText(nCol + 1, nRow, rs.GetValue(_T("treatqty")), FMT_NUMBER1);
	//BN dieu tri chi tiet
	szSQL.Format(_T(" SELECT") \
				_T("   ho_recordno as recordno,") \
				_T("   Count(DISTINCT ho_docno) as cnt") \
				_T(" FROM   hms_operation") \
				_T(" WHERE  ho_pdeptid='C15'") \
				_T("        AND ho_recordno IS NOT NULL") \
				_T("		AND ho_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				_T("        AND ho_status <> 'O'") \
				_T("		AND ho_payment = 'P'") \
				_T(" GROUP  BY ho_recordno") \
				_T(" ORDER BY ho_recordno"), m_szFromDate, m_szToDate);
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		j = 0;
		for (int i = 0 ; i < arrMac.GetCount(); i++)
		{
			if (arrMac.GetAt(i) == _T("HE"))
				j = i - 1;
			rs.GetValue(_T("recordno"), tmpStr);
			if (tmpStr == arrMac.GetAt(i))
			{
				rs.GetValue(_T("cnt"), nTemp);
				if (tmpStr != _T("HE"))
					xls.SetCellText(nCol + j + 2, nRow, rs.GetValue(_T("cnt")), FMT_NUMBER1);
				xls.SetCellText(nCol + i + 7, nRow, rs.GetValue(_T("cnt")), FMT_NUMBER1);
				if (tmpStr == _T("CO") || tmpStr == _T("HE"))
					nTmp += nTemp;
				break;
			}
			j++;
		}
		rs.MoveNext();
	}
	xls.SetCellText(nCol + 2, nRow, int2str(nTmp), FMT_NUMBER1);
	//So tien
	szSQL.Format(_T(" SELECT") \
				_T("   ho_recordno as recordno,") \
				_T("   Sum(hfe_cost) as cost") \
				_T(" FROM   hms_operation") \
				_T("        LEFT JOIN hms_fee ON (hfe_docno=ho_docno") \
				_T("                              AND ho_itemid=hfe_itemid AND hfe_orderid = ho_idx)") \
				_T(" WHERE  ho_pdeptid='C15'") \
				_T("        AND ho_recordno IS NOT NULL") \
				_T("        AND ho_payment='P'") \
				_T("		AND hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				_T(" GROUP  BY ho_recordno") \
				_T(" ORDER BY ho_recordno"), m_szFromDate, m_szToDate);
	j = 0;
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		for (int i = 0 ; i < arrMac.GetCount(); i++)
		{
			rs.GetValue(_T("recordno"), tmpStr);
			if (tmpStr == arrMac.GetAt(i))
			{
				rs.GetValue(_T("cost"), nAmt);
				xls.SetCellText(nCol + i + 13, nRow, double2str(nAmt), FMT_NUMBER1);
				nTotal += nAmt;
				break;
			}
		}
		rs.MoveNext();
	}	
	xls.SetCellText(nCol + 19, nRow, double2str(nTotal), FMT_NUMBER1);
	xls.Save(_T("Exports\\Bao cao mat benh theo may.xls"));
} 