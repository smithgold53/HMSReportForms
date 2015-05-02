#include "stdafx.h"
#include "PMSServiceDrugUseofInpatient.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSServiceDrugUseofInpatient *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSServiceDrugUseofInpatient *)pWnd)->OnToDateCheckValue();
} 
static void _OnDeptSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSServiceDrugUseofInpatient* )pWnd)->OnDeptSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDeptSelendokFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnDeptSelendok();
}
/*static void _OnDeptSetfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnDeptKillfocus();
}*/
/*static void _OnDeptKillfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnDeptKillfocus();
}*/
static long _OnDeptLoadDataFnc(CWnd *pWnd){
	return ((CPMSServiceDrugUseofInpatient *)pWnd)->OnDeptLoadData();
}
/*static void _OnDeptAddNewFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnDeptAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSServiceDrugUseofInpatient *pVw = (CPMSServiceDrugUseofInpatient *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSServiceDrugUseofInpatient *pVw = (CPMSServiceDrugUseofInpatient *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSServiceDrugUseofInpatient* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMSServiceDrugUseofInpatient *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofInpatient *)pWnd)->OnStorageAddNew();
}*/
static int _OnAddPMSServiceDrugUseofOutpatientFnc(CWnd *pWnd){
	 return ((CPMSServiceDrugUseofInpatient*)pWnd)->OnAddPMSServiceDrugUseofOutpatient();
} 
static int _OnEditPMSServiceDrugUseofOutpatientFnc(CWnd *pWnd){
	 return ((CPMSServiceDrugUseofInpatient*)pWnd)->OnEditPMSServiceDrugUseofOutpatient();
} 
static int _OnDeletePMSServiceDrugUseofOutpatientFnc(CWnd *pWnd){
	 return ((CPMSServiceDrugUseofInpatient*)pWnd)->OnDeletePMSServiceDrugUseofOutpatient();
} 
static int _OnSavePMSServiceDrugUseofOutpatientFnc(CWnd *pWnd){
	 return ((CPMSServiceDrugUseofInpatient*)pWnd)->OnSavePMSServiceDrugUseofOutpatient();
} 
static int _OnCancelPMSServiceDrugUseofOutpatientFnc(CWnd *pWnd){
	 return ((CPMSServiceDrugUseofInpatient*)pWnd)->OnCancelPMSServiceDrugUseofOutpatient();
} 
CPMSServiceDrugUseofInpatient::CPMSServiceDrugUseofInpatient(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CPMSServiceDrugUseofInpatient::~CPMSServiceDrugUseofInpatient(){
}
void CPMSServiceDrugUseofInpatient::OnCreateComponents(){
	m_wndServiceDrugUseofOutpatient.Create(this, _T("Service Drug Use of Outpatient"), 5, 5, 575, 545);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndDeptLabel.Create(this, _T("Dept"), 10, 60, 90, 85);
	m_wndDept.Create(this,95, 60, 570, 85); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 385, 550, 495, 575);
	m_wndExport.Create(this, _T("&Export"), 500, 550, 575, 575);
	m_wndStorageList.Create(this,10, 90, 570, 540); 
	m_wndStorageList.SetCheckBox(true);
}
void CPMSServiceDrugUseofInpatient::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndDept.SetCheckValue(true);
	m_wndDept.LimitText(1024);

	
	m_wndDept.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndDept.InsertColumn(1, _T("Name"), CFMT_TEXT, 300); 

	m_wndStorageList.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStorageList.InsertColumn(1, _T("Name"), CFMT_TEXT, 450);

}
void CPMSServiceDrugUseofInpatient::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndDept.SetEvent(WE_SELENDOK, _OnDeptSelendokFnc);
	//m_wndDept.SetEvent(WE_SETFOCUS, _OnDeptSetfocusFnc);
	//m_wndDept.SetEvent(WE_KILLFOCUS, _OnDeptKillfocusFnc);
	m_wndDept.SetEvent(WE_SELCHANGE, _OnDeptSelectChangeFnc);
	m_wndDept.SetEvent(WE_LOADDATA, _OnDeptLoadDataFnc);
	//m_wndDept.SetEvent(WE_ADDNEW, _OnDeptAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndStorageList.SetEvent(WE_SELENDOK, _OnStorageSelendokFnc);
	//m_wndStorageList.SetEvent(WE_SETFOCUS, _OnStorageSetfocusFnc);
	//m_wndStorageList.SetEvent(WE_KILLFOCUS, _OnStorageKillfocusFnc);
	m_wndStorageList.SetEvent(WE_SELCHANGE, _OnStorageSelectChangeFnc);
	m_wndStorageList.SetEvent(WE_LOADDATA, _OnStorageLoadDataFnc);
	//m_wndStorageList.SetEvent(WE_ADDNEW, _OnStorageAddNewFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	OnStorageLoadData();
	UpdateData(false);

}
void CPMSServiceDrugUseofInpatient::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDept.GetDlgCtrlID(), m_szDeptKey);
	DDX_TextEx(pDX, m_wndStorageList.GetDlgCtrlID(), m_szStorageKey);

}
void CPMSServiceDrugUseofInpatient::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPMSServiceDrugUseofInpatient::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CPMSServiceDrugUseofInpatient::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDeptKey.Empty();
	m_szStorageKey.Empty();

}
int CPMSServiceDrugUseofInpatient::SetMode(int nMode){
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
/*void CPMSServiceDrugUseofInpatient::OnFromDateChange(){
	
} */
/*void CPMSServiceDrugUseofInpatient::OnFromDateSetfocus(){
	
} */
/*void CPMSServiceDrugUseofInpatient::OnFromDateKillfocus(){
	
} */
int CPMSServiceDrugUseofInpatient::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSServiceDrugUseofInpatient::OnToDateChange(){
	
} */
/*void CPMSServiceDrugUseofInpatient::OnToDateSetfocus(){
	
} */
/*void CPMSServiceDrugUseofInpatient::OnToDateKillfocus(){
	
} */
int CPMSServiceDrugUseofInpatient::OnToDateCheckValue(){
	return 0;
} 
void CPMSServiceDrugUseofInpatient::OnDeptSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSServiceDrugUseofInpatient::OnDeptSelendok(){
	 
}
/*void CPMSServiceDrugUseofInpatient::OnDeptSetfocus(){
	
}*/
/*void CPMSServiceDrugUseofInpatient::OnDeptKillfocus(){
	
}*/
long CPMSServiceDrugUseofInpatient::OnDeptLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDept.IsSearchKey() && !m_szDeptKey.IsEmpty()){
	};
	szSQL.Format(_T("SELECT sd_id as id, sd_name as deptname FROM sys_dept WHERE 1=1 ORDER BY sd_id"));
	m_wndDept.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDept.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("deptname")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPMSServiceDrugUseofInpatient::OnDeptAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSServiceDrugUseofInpatient::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szSysDate, szOldTit, szNewTit;
	double nOrder = 0, nSell = 0, nBought = 0 , nProfit = 0, nTemp = 0;
	long double nOrderTotal = 0,nSellTotal = 0, nBoughtTotal = 0, nProfitTotal = 0;
	if (!rpt.Init(_T("Reports/HMS/PMS_BAOCAOSUDUNGTHUOCDICHVUNOITRU.RPT")))
		return;
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		//QueryOpener(szSQL);
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection *rptDetail;
	while (!rs.IsEOF()){
		rs.GetValue(_T("title"), szNewTit);
		if (szNewTit != szOldTit)
		{
			if (nOrder > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				FormatCurrency(nOrder, tmpStr);
				rptDetail->SetValue(_T("s1"), tmpStr);
				FormatCurrency(nBought, tmpStr);
				rptDetail->SetValue(_T("s2"), tmpStr);
				FormatCurrency(nSell, tmpStr);
				rptDetail->SetValue(_T("s3"), tmpStr);
				FormatCurrency(nProfit, tmpStr);
				rptDetail->SetValue(_T("s4"), tmpStr);
				nOrderTotal += nOrder;nOrder = 0;
				nBoughtTotal += nBought;nBought = 0;
				nSellTotal += nSell;nSell = 0;
				nProfitTotal += nProfit;nProfit = 0;
			}
			szOldTit = szNewTit;
		}
		rptDetail = rpt.AddDetail();
		rs.GetValue(_T("deptname"), tmpStr);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("ordercount"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		nOrder += nTemp;
		rs.GetValue(_T("vatamount"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		nBought += nTemp;
		rs.GetValue(_T("servamount"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		nSell+= nTemp;
		rs.GetValue(_T("profitamount"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		nProfit += nTemp;
		rs.MoveNext();
	}
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	FormatCurrency(nOrder, tmpStr);
	rptDetail->SetValue(_T("s1"), tmpStr);
	FormatCurrency(nBought, tmpStr);
	rptDetail->SetValue(_T("s2"), tmpStr);
	FormatCurrency(nSell, tmpStr);
	rptDetail->SetValue(_T("s3"), tmpStr);
	FormatCurrency(nProfit, tmpStr);
	rptDetail->SetValue(_T("s4"), tmpStr);
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();	
} 
void CPMSServiceDrugUseofInpatient::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	CString szOldTit, szNewTit;
	int nCol = 0, nRow = 0;
	int nCnt = 0;
	double nOrderTtl = 0, nSellTtl = 0, nBoughtTtl = 0, nProfitTtl = 0, nTemp = 0;
	double nOrderGrp = 0, nSellGrp = 0, nBoughtGrp = 0, nProfitGrp = 0;
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 35);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 15);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 15);
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 6);
	xls.SetCellMergedColumns(nCol, nRow + 3, 6);
	xls.SetCellText(nCol, nRow, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, _T("\x42\xC1O \x43\xC1O S\x1EEC \x44\x1EE4NG THU\x1ED0\x43 \x44\x1ECA\x43H V\x1EE4 N\x1ED8I TR\xDA"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("T\xEAn kho\x61"));
	arrCol.Add(_T("T\x1ED5ng s\x1ED1 phi\x1EBFu"));
	arrCol.Add(_T("Ti\x1EC1n thu\x1ED1\x63 \x62\xE1n"));
	arrCol.Add(_T("Ti\x1EC1n thu\x1ED1\x63 mu\x61"));
	arrCol.Add(_T("Ti\x1EC1n l\xE3i"));
	nRow = 4;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10); 
	}
	nRow = 5;
	while (!rs.IsEOF()){
		rs.GetValue(_T("title"), szNewTit);
		if (szNewTit != szOldTit)
		{
			if (nOrderGrp > 0)
			{
				TranslateString(_T("Total Group"), tmpStr);
				xls.SetCellText(nCol, nRow, tmpStr, 4098, true, 10);
				xls.SetCellText(nCol + 1, nRow, double2str(nOrderGrp), FMT_NUMBER1, true, 10);
				nOrderTtl += nOrderGrp;nOrderGrp = 0;
				xls.SetCellText(nCol + 2, nRow, double2str(nSellGrp), FMT_NUMBER1, true, 10);
				nSellTtl += nSellGrp;nSellGrp = 0;
				xls.SetCellText(nCol + 3, nRow, double2str(nBoughtGrp), FMT_NUMBER1, true, 10);
				nBoughtTtl += nBoughtGrp;nBoughtGrp = 0;
				xls.SetCellText(nCol + 4, nRow, double2str(nProfitGrp), FMT_NUMBER1, true, 10);
				nProfitTtl += nProfitGrp;nProfitGrp = 0;
				nRow++;
			}
			xls.SetCellMergedColumns(0, nRow, 5);
			xls.SetCellText(0, nRow, szNewTit, FMT_TEXT, true);
			nRow++;
			szOldTit = szNewTit;
		}
		rs.GetValue(_T("deptname"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("ordercount"), nTemp);
		xls.SetCellText(nCol + 1, nRow, double2str(nTemp), FMT_NUMBER1);
		nOrderGrp += nTemp;
		rs.GetValue(_T("vatamount"), nTemp);
		xls.SetCellText(nCol + 2, nRow, double2str(nTemp), FMT_NUMBER1);
		nSellGrp += nTemp;
		rs.GetValue(_T("servamount"), nTemp);
		xls.SetCellText(nCol + 3, nRow, double2str(nTemp), FMT_NUMBER1);
		nBoughtGrp += nTemp;
		rs.GetValue(_T("profitamount"), nTemp);
		xls.SetCellText(nCol + 4, nRow, double2str(nTemp), FMT_NUMBER1);
		nProfitGrp += nTemp;
		nCnt++;
		nRow++;
		rs.MoveNext();
	}
	nOrderTtl += nOrderGrp;
	nSellTtl += nSellGrp;
	nBoughtTtl += nBoughtGrp;
	nProfitTtl += nProfitGrp;
	if (nOrderGrp > 0)
	{
		TranslateString(_T("Total Group"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT, true, 10);
		xls.SetCellText(nCol + 1, nRow, double2str(nOrderGrp), FMT_NUMBER1, true, 10);
		nOrderTtl += nOrderGrp;
		xls.SetCellText(nCol + 2, nRow, double2str(nSellGrp), FMT_NUMBER1, true, 10);
		nSellTtl += nSellGrp;
		xls.SetCellText(nCol + 3, nRow, double2str(nBoughtGrp), FMT_NUMBER1, true, 10);
		nBoughtTtl += nBoughtGrp;
		xls.SetCellText(nCol + 4, nRow, double2str(nProfitGrp), FMT_NUMBER1, true, 10);
		nProfitTtl += nProfitGrp;
	}
	TranslateString(_T("Total Group"), tmpStr);
	xls.SetCellText(nCol, nRow, tmpStr, 4098, true, 10);
	xls.SetCellText(nCol + 1, nRow, double2str(nOrderGrp), FMT_NUMBER1, true, 10);
	xls.SetCellText(nCol + 2, nRow, double2str(nSellGrp), FMT_NUMBER1, true, 10);
	xls.SetCellText(nCol + 3, nRow, double2str(nBoughtGrp), FMT_NUMBER1, true, 10);
	xls.SetCellText(nCol + 4, nRow, double2str(nProfitGrp), FMT_NUMBER1, true, 10);
	xls.Save(_T("Exports\\Bao Cao Su Dung Thuoc Dich Vu Noi Tru.xls"));	
} 
void CPMSServiceDrugUseofInpatient::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSServiceDrugUseofInpatient::OnStorageSelendok(){
	 
}
/*void CPMSServiceDrugUseofInpatient::OnStorageSetfocus(){
	
}*/
/*void CPMSServiceDrugUseofInpatient::OnStorageKillfocus(){
	
}*/
long CPMSServiceDrugUseofInpatient::OnStorageLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND msl_type = 'A'"));
	return pMF->LoadStorageIntoListCtrl(&m_wndStorageList, szFilter);
	//CRecord rs(&pMF->m_db);
	//CString szSQL, szWhere;
	//m_wndStorageList.BeginLoad();
	//int nCount = 0;
	//szSQL.Format(_T("SELECT msl_storage_id as id, msl_name as name FROM m_storagelist WHERE msl_type = 'A' ORDER BY msl_storage_id "));
	//nCount = rs.ExecSQL(szSQL);
	//while(!rs.IsEOF()){ 
	//	m_wndStorageList.AddItems(
	//		rs.GetValue(_T("id")), 
	//		rs.GetValue(_T("name")), NULL);
	//	rs.MoveNext();
	//}
	//m_wndStorageList.EndLoad();
	//return nCount;
}
/*void CPMSServiceDrugUseofInpatient::OnStorageAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
int CPMSServiceDrugUseofInpatient::OnAddPMSServiceDrugUseofOutpatient(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CPMSServiceDrugUseofInpatient::OnEditPMSServiceDrugUseofOutpatient(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CPMSServiceDrugUseofInpatient::OnDeletePMSServiceDrugUseofOutpatient(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd();
 	CString szSQL;
 	if(ShowMessage(1, MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2) == IDNO)
 		return -1;
 	szSQL.Format(_T("DELETE FROM  WHERE  AND") );
 	int ret = pMF->ExecSQL(szSQL);
 	if(ret >= 0){
 		SetMode(VM_NONE);
 		OnCancelPMSServiceDrugUseofOutpatient();
 	}
	return 0;
}
int CPMSServiceDrugUseofInpatient::OnSavePMSServiceDrugUseofOutpatient(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT)
 		return -1;
 	if(!IsValidateData())
 		return -1;
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd();
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
 		//OnPMSServiceDrugUseofOutpatientListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CPMSServiceDrugUseofInpatient::OnCancelPMSServiceDrugUseofOutpatient(){
 	if(GetMode() == VM_EDIT)
 	{
 		SetMode(VM_VIEW);
 	} 
 	else 
 	{
 		SetMode(VM_NONE);
 	} 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd();
 	return 0;
} 
int CPMSServiceDrugUseofInpatient::OnPMSServiceDrugUseofOutpatientListLoadData(){
	return 0;
}
CString CPMSServiceDrugUseofInpatient::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
	CString szSQL, szWhere, szTransactionWhere, szSaleWhere, szStorages;
	szWhere.Format(_T(" AND	hpo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szSaleWhere.Format(_T(" AND so_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szTransactionWhere.Format(_T(" AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	for (int i=0 ; i< m_wndStorageList.GetItemCount(); i++)
	{
		if(m_wndStorageList.GetCheck(i))
		{
			m_wndStorageList.SetCurSel(i);
			if(!szStorages.IsEmpty())
				szStorages += _T(", ");
			szStorages.AppendFormat(_T("%s"), m_wndStorageList.GetItemText(i, 0));		
		}		
	}
	if (!szStorages.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hpo_storage_id IN (%s)"), szStorages);
		szSaleWhere.AppendFormat(_T(" AND so_storage_id IN (%s)"), szStorages);
		szTransactionWhere.AppendFormat(_T(" AND mt_storage_to_id IN(%s)"), szStorages);
	}
	if (!m_szDeptKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hpo_deptid = '%s'"), m_szDeptKey);
		szTransactionWhere.AppendFormat(_T(" AND mt_department_to_id = '%s'"), m_szDeptKey);
	}
	szSQL.Format(_T(" SELECT title, ") \
				_T("        case when deptid = 'R' then N'\x58u\x1EA5t \x62\xE1n l\x1EBB' else sd_name end deptname,") \
				_T("        Count(DISTINCT orderno)     ordercount, ") \
				_T("        SUM(vatamount)              AS vatamount, ") \
				_T("        SUM(servamount)             AS servamount, ") \
				_T("        SUM(servamount - vatamount) profitamount ") \
				_T(" FROM   (SELECT    N'Phi\x1EBFu l\x129nh' title, Nvl(mt_orderno, hpo_orderid)     AS orderno, ") \
				_T("                   hpo_deptid deptid, ") \
				_T("                   hpol_qtyorder * product_taxprice AS vatamount, ") \
				_T("                   hpol_qtyorder * hpol_unitprice   AS servamount ") \
				_T("         FROM      hms_ipharmaorder ") \
				_T("         LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T("         LEFT JOIN m_transaction ON ( mt_transaction_id = hpo_transaction_id ) ") \
				_T("         LEFT JOIN m_productitem_view ON ( product_item_id = hpol_product_item_id ) ") \
				_T("         WHERE     hpo_orderstatus = 'A' ") \
				_T("               AND hpo_ordertype IN ('D', 'M', 'B') %s") \
				_T("		 UNION ALL") \
				_T("         SELECT    N'Phi\x1EBFu l\x129nh' title, mt_orderno                       AS orderno, ") \
				_T("                   mt_department_id, ") \
				_T("                   -hpol_qtyreverse * product_taxprice AS vatamount, ") \
				_T("                   -hpol_qtyreverse * hpol_unitprice   AS servamount ") \
				_T("		 FROM hms_ipharmaorder") \
				_T("         LEFT JOIN hms_ipharmaorderline ON (hpo_orderid = hpol_orderid)") \
				_T("         LEFT JOIN m_transaction ON ( mt_transaction_id = hpol_retorder_id ) ") \
				_T("         LEFT JOIN m_productitem_view ON ( product_item_id = hpol_product_item_id ) ") \
				_T("         WHERE hpol_qtyreverse > 0 AND mt_doctype = 'DRO' AND mt_status = 'A' %s") \
				_T("         UNION ALL ") \
				_T("         SELECT    N'\x58u\x1EA5t \x62\xE1n l\x1EBB', Cast(so_sale_order_id AS NVARCHAR2(15)), ") \
				_T("                   N'R', ") \
				_T("                   sol_qtyorder * product_taxprice AS vatamount, ") \
				_T("                   sol_qtyorder * sol_unitprice    AS servamount ") \
				_T("         FROM      sale_order ") \
				_T("         LEFT JOIN sale_orderline ON ( so_sale_order_id = sol_sale_order_id ) ") \
				_T("         LEFT JOIN m_productitem_view ON ( product_item_id = sol_product_item_id ) ") \
				_T("         WHERE     so_status = 'A' ") \
				_T("               AND so_payment = 'P' %s)") \
				_T(" LEFT JOIN sys_dept ON (sd_id = deptid)") \
				_T(" GROUP  BY title, deptid, sd_name, sd_index ") \
				_T(" ORDER  BY sd_index "), szWhere, szTransactionWhere, szSaleWhere);

	return szSQL;
}BEGIN_MESSAGE_MAP(CPMSServiceDrugUseofInpatient, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSServiceDrugUseofInpatient::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
