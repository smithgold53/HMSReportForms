#include "stdafx.h"
#include "PMSServiceDrugUseofOutpatient.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include "ReportDocument.h"
#include ".\pmsservicedruguseofoutpatient.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSServiceDrugUseofOutpatient *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSServiceDrugUseofOutpatient *)pWnd)->OnToDateCheckValue();
} 
static void _OnRoomSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSServiceDrugUseofOutpatient* )pWnd)->OnRoomSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnRoomSelendokFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnRoomSelendok();
}
/*static void _OnRoomSetfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnRoomKillfocus();
}*/
/*static void _OnRoomKillfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnRoomKillfocus();
}*/
static long _OnRoomLoadDataFnc(CWnd *pWnd){
	return ((CPMSServiceDrugUseofOutpatient *)pWnd)->OnRoomLoadData();
}
/*static void _OnRoomAddNewFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnRoomAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSServiceDrugUseofOutpatient *pVw = (CPMSServiceDrugUseofOutpatient *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSServiceDrugUseofOutpatient *pVw = (CPMSServiceDrugUseofOutpatient *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSServiceDrugUseofOutpatient* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMSServiceDrugUseofOutpatient *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMSServiceDrugUseofOutpatient *)pWnd)->OnStorageAddNew();
}*/
static int _OnAddPMSServiceDrugUseofOutpatientFnc(CWnd *pWnd){
	 return ((CPMSServiceDrugUseofOutpatient*)pWnd)->OnAddPMSServiceDrugUseofOutpatient();
} 
static int _OnEditPMSServiceDrugUseofOutpatientFnc(CWnd *pWnd){
	 return ((CPMSServiceDrugUseofOutpatient*)pWnd)->OnEditPMSServiceDrugUseofOutpatient();
} 
static int _OnDeletePMSServiceDrugUseofOutpatientFnc(CWnd *pWnd){
	 return ((CPMSServiceDrugUseofOutpatient*)pWnd)->OnDeletePMSServiceDrugUseofOutpatient();
} 
static int _OnSavePMSServiceDrugUseofOutpatientFnc(CWnd *pWnd){
	 return ((CPMSServiceDrugUseofOutpatient*)pWnd)->OnSavePMSServiceDrugUseofOutpatient();
} 
static int _OnCancelPMSServiceDrugUseofOutpatientFnc(CWnd *pWnd){
	 return ((CPMSServiceDrugUseofOutpatient*)pWnd)->OnCancelPMSServiceDrugUseofOutpatient();
} 
CPMSServiceDrugUseofOutpatient::CPMSServiceDrugUseofOutpatient(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CPMSServiceDrugUseofOutpatient::~CPMSServiceDrugUseofOutpatient(){
}
void CPMSServiceDrugUseofOutpatient::OnCreateComponents(){
	m_wndServiceDrugUseofOutpatient.Create(this, _T("Service Drug Use of Outpatient"), 5, 5, 575, 545);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndRoomLabel.Create(this, _T("Room"), 10, 60, 90, 85);
	m_wndRoom.Create(this,95, 60, 570, 85); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 385, 550, 495, 575);
	m_wndExport.Create(this, _T("&Export"), 500, 550, 575, 575);
	m_wndStorageList.Create(this,10, 90, 570, 540); 
	m_wndStorageList.SetCheckBox(true);
}
void CPMSServiceDrugUseofOutpatient::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndRoom.SetCheckValue(true);
	m_wndRoom.LimitText(1024);

	
	m_wndRoom.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndRoom.InsertColumn(1, _T("Name"), CFMT_TEXT, 200); 

	m_wndStorageList.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStorageList.InsertColumn(1, _T("Name"), CFMT_TEXT, 450);

}
void CPMSServiceDrugUseofOutpatient::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndRoom.SetEvent(WE_SELENDOK, _OnRoomSelendokFnc);
	//m_wndRoom.SetEvent(WE_SETFOCUS, _OnRoomSetfocusFnc);
	//m_wndRoom.SetEvent(WE_KILLFOCUS, _OnRoomKillfocusFnc);
	m_wndRoom.SetEvent(WE_SELCHANGE, _OnRoomSelectChangeFnc);
	m_wndRoom.SetEvent(WE_LOADDATA, _OnRoomLoadDataFnc);
	//m_wndRoom.SetEvent(WE_ADDNEW, _OnRoomAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndStorageList.SetEvent(WE_SELENDOK, _OnStorageSelendokFnc);
	//m_wndStorageList.SetEvent(WE_SETFOCUS, _OnStorageSetfocusFnc);
	//m_wndStorageList.SetEvent(WE_KILLFOCUS, _OnStorageKillfocusFnc);
	m_wndStorageList.SetEvent(WE_SELCHANGE, _OnStorageSelectChangeFnc);
	m_wndStorageList.SetEvent(WE_LOADDATA, _OnStorageLoadDataFnc);
	//m_wndStorageList.SetEvent(WE_ADDNEW, _OnStorageAddNewFnc);
	AddEvent(1, _T("Add	Ctrl+A"), _OnAddPMSServiceDrugUseofOutpatientFnc, 0, 'A', VK_CONTROL);
	AddEvent(2, _T("Edit	Ctrl+E"), _OnEditPMSServiceDrugUseofOutpatientFnc, 0, 'E', VK_CONTROL);
	AddEvent(3, _T("Delete	Ctrl+D"), _OnDeletePMSServiceDrugUseofOutpatientFnc, 0, 'D', VK_CONTROL);
	AddEvent(4, _T("Save	Ctrl+S"), _OnSavePMSServiceDrugUseofOutpatientFnc, 0, 'S', VK_CONTROL);
	AddEvent(0, _T("-"));
	AddEvent(5, _T("Cancel	Ctrl+T"), _OnCancelPMSServiceDrugUseofOutpatientFnc, 0, 'T', VK_CONTROL);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	OnStorageLoadData();
	UpdateData(false);

}
void CPMSServiceDrugUseofOutpatient::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndRoom.GetDlgCtrlID(), m_szRoomKey);
	DDX_TextEx(pDX, m_wndStorageList.GetDlgCtrlID(), m_szStorageKey);

}
void CPMSServiceDrugUseofOutpatient::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPMSServiceDrugUseofOutpatient::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CPMSServiceDrugUseofOutpatient::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szRoomKey.Empty();
	m_szStorageKey.Empty();

}
int CPMSServiceDrugUseofOutpatient::SetMode(int nMode){
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
/*void CPMSServiceDrugUseofOutpatient::OnFromDateChange(){
	
} */
/*void CPMSServiceDrugUseofOutpatient::OnFromDateSetfocus(){
	
} */
/*void CPMSServiceDrugUseofOutpatient::OnFromDateKillfocus(){
	
} */
int CPMSServiceDrugUseofOutpatient::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSServiceDrugUseofOutpatient::OnToDateChange(){
	
} */
/*void CPMSServiceDrugUseofOutpatient::OnToDateSetfocus(){
	
} */
/*void CPMSServiceDrugUseofOutpatient::OnToDateKillfocus(){
	
} */
int CPMSServiceDrugUseofOutpatient::OnToDateCheckValue(){
	return 0;
} 
void CPMSServiceDrugUseofOutpatient::OnRoomSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSServiceDrugUseofOutpatient::OnRoomSelendok(){
	 
}
/*void CPMSServiceDrugUseofOutpatient::OnRoomSetfocus(){
	
}*/
/*void CPMSServiceDrugUseofOutpatient::OnRoomKillfocus(){
	
}*/
long CPMSServiceDrugUseofOutpatient::OnRoomLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndRoom.IsSearchKey() && !m_szRoomKey.IsEmpty()){
	};
	szSQL.Format(_T("SELECT hrl_id as id, hrl_name as description FROM hms_roomlist WHERE hrl_deptid = 'C1.1' ORDER BY id"));
	m_wndRoom.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndRoom.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPMSServiceDrugUseofOutpatient::OnRoomAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSServiceDrugUseofOutpatient::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szSysDate;
	double nOrder = 0, nSell = 0, nBought = 0 , nProfit = 0, nTemp = 0;
	if (!rpt.Init(_T("Reports/HMS/PMS_BAOCAOSUDUNGTHUOCDICHVUNGOAITRU.RPT")))
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
		rptDetail = rpt.AddDetail();
		rs.GetValue(_T("room"), tmpStr);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("doctor"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("ordercount"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		nOrder += nTemp;
		rs.GetValue(_T("sellprice"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		nSell += nTemp;
		rs.GetValue(_T("boughtprice"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		nBought+= nTemp;
		rs.GetValue(_T("profit"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		nProfit += nTemp;
		rs.MoveNext();
	}
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	FormatCurrency(nOrder, tmpStr);
	rptDetail->SetValue(_T("s1"), tmpStr);
	FormatCurrency(nSell, tmpStr);
	rptDetail->SetValue(_T("s2"), tmpStr);
	FormatCurrency(nBought, tmpStr);
	rptDetail->SetValue(_T("s3"), tmpStr);
	FormatCurrency(nProfit, tmpStr);
	rptDetail->SetValue(_T("s4"), tmpStr);
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();	
} 
void CPMSServiceDrugUseofOutpatient::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	CString szOldRoom, szNewRoom;
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
	xls.SetColumnWidth(0, 20);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 15);
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 6);
	xls.SetCellMergedColumns(nCol, nRow + 3, 6);
	xls.SetCellText(nCol, nRow, _T("HEALTHSERVICE"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, _T("HOSPITALNAME"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, _T("\x42\xC1O \x43\xC1O S\x1EEC \x44\x1EE4NG THU\x1ED0\x43 \x44\x1ECA\x43H V\x1EE4 NGO\x1EA0I TR\xDA"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("Ph\xF2ng kh\xE1m"));
	arrCol.Add(_T("T\xEAn \x62\xE1\x63 s\x129"));
	arrCol.Add(_T("T\x1ED5ng s\x1ED1 \x111\x1A1n"));
	arrCol.Add(_T("Ti\x1EC1n thu\x1ED1\x63 \x62\xE1n"));
	arrCol.Add(_T("Ti\x1EC1n thu\x1ED1\x63 mu\x61"));
	arrCol.Add(_T("Ti\x1EC1n l\xE3i"));
	nRow = 4;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10); 
	}
	nRow = 5;
	rs.GetValue(_T("room"), szOldRoom);
	while (!rs.IsEOF()){
		rs.GetValue(_T("room"), szNewRoom);
		if (szNewRoom != szOldRoom)
		{
			if (nCnt > 1)
			{
				xls.SetCellText(0, nRow, _T("T\x1ED5ng nh\xF3m"), FMT_TEXT, true);
				xls.SetCellText(2, nRow, double2str(nOrderGrp),FMT_NUMBER1);
				xls.SetCellText(3, nRow, double2str(nSellGrp),FMT_NUMBER1);
				xls.SetCellText(4, nRow, double2str(nBoughtGrp),FMT_NUMBER1);
				xls.SetCellText(5, nRow, double2str(nProfitGrp),FMT_NUMBER1);
				nRow++;
			}
			szOldRoom = szNewRoom;
			nOrderTtl += nOrderGrp;
			nSellTtl += nSellGrp;
			nBoughtTtl += nBoughtGrp;
			nProfitTtl += nProfitGrp;
			nOrderGrp = nSellGrp = nBoughtGrp = nProfitGrp = nCnt = 0;
		}
		xls.SetCellText(nCol, nRow, szNewRoom, FMT_TEXT);
		rs.GetValue(_T("doctor"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT);
		rs.GetValue(_T("ordercount"), nTemp);
		xls.SetCellText(nCol + 2, nRow, double2str(nTemp), FMT_NUMBER1);
		nOrderGrp += nTemp;
		rs.GetValue(_T("sellprice"), nTemp);
		xls.SetCellText(nCol + 3, nRow, double2str(nTemp), FMT_NUMBER1);
		nSellGrp += nTemp;
		rs.GetValue(_T("boughtprice"), nTemp);
		xls.SetCellText(nCol + 4, nRow, double2str(nTemp), FMT_NUMBER1);
		nBoughtGrp += nTemp;
		rs.GetValue(_T("profit"), nTemp);
		xls.SetCellText(nCol + 5, nRow, double2str(nTemp), FMT_NUMBER1);
		nProfitGrp += nTemp;
		nCnt++;
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellText(0, nRow, _T("T\x1ED5ng nh\xF3m"), FMT_TEXT);
	xls.SetCellText(2, nRow, double2str(nOrderGrp),FMT_NUMBER1);
	xls.SetCellText(3, nRow, double2str(nSellGrp),FMT_NUMBER1);
	xls.SetCellText(4, nRow, double2str(nBoughtGrp),FMT_NUMBER1);
	xls.SetCellText(5, nRow, double2str(nProfitGrp),FMT_NUMBER1);
	nRow++;
	nOrderTtl += nOrderGrp;
	nSellTtl += nSellGrp;
	nBoughtTtl += nBoughtGrp;
	nProfitTtl += nProfitGrp;
	TranslateString(_T("Total"), tmpStr);
	xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT, true, 10);
	xls.SetCellText(nCol + 2, nRow, double2str(nOrderTtl), FMT_NUMBER1, true, 10);
	xls.SetCellText(nCol + 3, nRow, double2str(nSellTtl), FMT_NUMBER1, true, 10);
	xls.SetCellText(nCol + 4, nRow, double2str(nBoughtTtl), FMT_NUMBER1, true, 10);
	xls.SetCellText(nCol + 5, nRow, double2str(nProfitTtl), FMT_NUMBER1, true, 10);
	xls.Save(_T("Exports\\Bao Cao Su Dung Thuoc Dich Vu Ngoai Tru.xls"));	
} 
void CPMSServiceDrugUseofOutpatient::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSServiceDrugUseofOutpatient::OnStorageSelendok(){
	 
}
/*void CPMSServiceDrugUseofOutpatient::OnStorageSetfocus(){
	
}*/
/*void CPMSServiceDrugUseofOutpatient::OnStorageKillfocus(){
	
}*/
long CPMSServiceDrugUseofOutpatient::OnStorageLoadData(){
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
/*void CPMSServiceDrugUseofOutpatient::OnStorageAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
int CPMSServiceDrugUseofOutpatient::OnAddPMSServiceDrugUseofOutpatient(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CPMSServiceDrugUseofOutpatient::OnEditPMSServiceDrugUseofOutpatient(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CPMSServiceDrugUseofOutpatient::OnDeletePMSServiceDrugUseofOutpatient(){
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
int CPMSServiceDrugUseofOutpatient::OnSavePMSServiceDrugUseofOutpatient(){
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
int CPMSServiceDrugUseofOutpatient::OnCancelPMSServiceDrugUseofOutpatient(){
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
int CPMSServiceDrugUseofOutpatient::OnPMSServiceDrugUseofOutpatientListLoadData(){
	return 0;
}
CString CPMSServiceDrugUseofOutpatient::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *)AfxGetMainWnd(); 
	CString szSQL, szWhere, szStorages;
	szWhere.Format(_T(" AND	(hfe_createddate) BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
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
		szWhere.AppendFormat(_T(" AND hpo_storage_id IN (%s)"), szStorages);
	if (!m_szRoomKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_roomid = %d"), ToInt(m_szRoomKey));
	szSQL.Format(_T(" SELECT  hrl_name as room,") \
				_T("         get_username(hpo_doctor) as doctor,") \
				_T("         sum(ordercount) as ordercount,") \
				_T("         sum(sellprice) as sellprice,") \
				_T("         sum(boughtprice) as boughtprice,") \
				_T("         sum(sellprice-boughtprice) as profit") \
				_T(" FROM") \
				_T(" (") \
				_T(" SELECT  hpo_roomid,") \
				_T("         hpo_doctor,") \
				_T("         1 as ordercount,") \
				_T("         sum(hpol_qtyissue*product_saleprice1) as sellprice,") \
				_T("         sum(hpol_qtyissue*product_unitprice) as boughtprice") \
				_T(" FROM") \
				_T(" hms_pharmaorder") \
				_T(" LEFT JOIN hms_fee_invoice ON (hpo_invoiceno = hfe_invoiceno)") \
				_T(" LEFT JOIN hms_pharmaorderline ON (hpo_orderid = hpol_orderid)") \
				_T(" LEFT JOIN m_productitem_view ON (product_item_id = hpol_product_item_id)") \
				_T(" WHERE hpo_roomid > 0 AND hpo_deptid = 'C1.1'") \
				_T(" AND hfe_status = 'P' AND hfe_refstatus = 'P' %s") \
				_T(" GROUP BY hpo_orderid, hpo_doctor, hpo_roomid") \
				_T(" )") \
				_T(" LEFT JOIN hms_roomlist ON (hrl_deptid = 'C1.1' AND hrl_id = hpo_roomid)") \
				_T(" GROUP BY hrl_name, hpo_roomid, hpo_doctor") \
				_T(" ORDER BY cast(substr(hrl_name, 1, 2) as integer), doctor"), szWhere);
	return szSQL;
}BEGIN_MESSAGE_MAP(CPMSServiceDrugUseofOutpatient, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSServiceDrugUseofOutpatient::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
