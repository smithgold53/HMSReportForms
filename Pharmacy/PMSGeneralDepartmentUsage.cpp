#include "stdafx.h"
#include "PMSGeneralDepartmentUsage.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include ".\pmsgeneraldepartmentusage.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSGeneralDepartmentUsage *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSGeneralDepartmentUsage *)pWnd)->OnToDateCheckValue();
} 
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSGeneralDepartmentUsage* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMSGeneralDepartmentUsage *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnStorageAddNew();
}*/
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CPMSGeneralDepartmentUsage *)pWnd)->OnDeptListLoadData();
}
static void _OnItemGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSGeneralDepartmentUsage* )pWnd)->OnItemGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnItemGroupSelendokFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnItemGroupSelendok();
}
/*static void _OnItemGroupSetfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnItemGroupKillfocus();
}*/
/*static void _OnItemGroupKillfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnItemGroupKillfocus();
}*/
static long _OnItemGroupLoadDataFnc(CWnd *pWnd){
	return ((CPMSGeneralDepartmentUsage *)pWnd)->OnItemGroupLoadData();
}
/*static void _OnItemGroupAddNewFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage *)pWnd)->OnItemGroupAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSGeneralDepartmentUsage *pVw = (CPMSGeneralDepartmentUsage *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSGeneralDepartmentUsage *pVw = (CPMSGeneralDepartmentUsage *)pWnd;
	pVw->OnExportSelect();
} 
CPMSGeneralDepartmentUsage::CPMSGeneralDepartmentUsage(CWnd *pParent){

	m_nDlgWidth = 545;
	m_nDlgHeight = 155;
	SetDefaultValues();
}
CPMSGeneralDepartmentUsage::~CPMSGeneralDepartmentUsage(){
}
void CPMSGeneralDepartmentUsage::OnCreateComponents(){
	m_wndGeneralDepartmentUsage.Create(this, _T("General Department Usage"), 5, 5, 575, 545);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndItemGroupLabel.Create(this, _T("Item Group"), 10, 60, 90, 85);
	m_wndItemGroup.Create(this,95, 60, 570, 85); 
	m_wndOnlyDDO.Create(this, _T("Only DDO"), 5, 550, 155, 575);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 395, 550, 490, 575);
	m_wndExport.Create(this, _T("&Export"), 495, 550, 575, 575);
	m_wndDeptList.Create(this,10, 320, 570, 540); 
	m_wndDeptList.SetCheckBox(true);
	m_wndStorageList.Create(this,10, 90, 570, 315); 
	m_wndStorageList.SetCheckBox(true);
}
void CPMSGeneralDepartmentUsage::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndItemGroup.LimitText(35);


	m_wndStorageList.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStorageList.InsertColumn(1, _T("Name"), CFMT_TEXT, 380);

	m_wndDeptList.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndDeptList.InsertColumn(1, _T("Name"), CFMT_TEXT, 380);
	
	m_wndItemGroup.Format(3, 2);
	m_wndItemGroup.InsertColumn(0, _T(""), CFMT_TEXT, 0);
	m_wndItemGroup.InsertColumn(1, _T("Code"), CFMT_TEXT, 60);
	m_wndItemGroup.InsertColumn(2, _T("Name"), CFMT_TEXT, 400);


}
void CPMSGeneralDepartmentUsage::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndStorageList.SetEvent(WE_SELENDOK, _OnStorageSelendokFnc);
	//m_wndStorageList.SetEvent(WE_SETFOCUS, _OnStorageSetfocusFnc);
	//m_wndStorageList.SetEvent(WE_KILLFOCUS, _OnStorageKillfocusFnc);
	m_wndStorageList.SetEvent(WE_SELCHANGE, _OnStorageSelectChangeFnc);
	m_wndStorageList.SetEvent(WE_LOADDATA, _OnStorageLoadDataFnc);
	//m_wndStorageList.SetEvent(WE_ADDNEW, _OnStorageAddNewFnc);
	m_wndDeptList.SetEvent(WE_LOADDATA, _OnDeptListLoadDataFnc);
	m_wndItemGroup.SetEvent(WE_SELENDOK, _OnItemGroupSelendokFnc);
	//m_wndItemGroup.SetEvent(WE_SETFOCUS, _OnItemGroupSetfocusFnc);
	//m_wndItemGroup.SetEvent(WE_KILLFOCUS, _OnItemGroupKillfocusFnc);
	m_wndItemGroup.SetEvent(WE_SELCHANGE, _OnItemGroupSelectChangeFnc);
	m_wndItemGroup.SetEvent(WE_LOADDATA, _OnItemGroupLoadDataFnc);
	//m_wndItemGroup.SetEvent(WE_ADDNEW, _OnItemGroupAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T(" 00:00");
	m_szToDate += _T(" 23:59");
	OnStorageLoadData();
	OnDeptListLoadData();
	UpdateData(false);

}
void CPMSGeneralDepartmentUsage::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStorageList.GetDlgCtrlID(), m_szStorageKey);
	DDX_TextEx(pDX, m_wndItemGroup.GetDlgCtrlID(), m_szItemGroupKey);
	DDX_Check(pDX, m_wndOnlyDDO.GetDlgCtrlID(), m_bOnlyDDO);

}
void CPMSGeneralDepartmentUsage::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStorageKey.Empty();
	m_szItemGroupKey.Empty();

}
int CPMSGeneralDepartmentUsage::SetMode(int nMode){
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
/*void CPMSGeneralDepartmentUsage::OnFromDateChange(){
	
} */
/*void CPMSGeneralDepartmentUsage::OnFromDateSetfocus(){
	
} */
/*void CPMSGeneralDepartmentUsage::OnFromDateKillfocus(){
	
} */
int CPMSGeneralDepartmentUsage::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSGeneralDepartmentUsage::OnToDateChange(){
	
} */
/*void CPMSGeneralDepartmentUsage::OnToDateSetfocus(){
	
} */
/*void CPMSGeneralDepartmentUsage::OnToDateKillfocus(){
	
} */
int CPMSGeneralDepartmentUsage::OnToDateCheckValue(){
	return 0;
} 
void CPMSGeneralDepartmentUsage::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSGeneralDepartmentUsage::OnStorageSelendok(){
	 
}
/*void CPMSGeneralDepartmentUsage::OnStorageSetfocus(){
	
}*/
/*void CPMSGeneralDepartmentUsage::OnStorageKillfocus(){
	
}*/
long CPMSGeneralDepartmentUsage::OnStorageLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageIntoListCtrl(&m_wndStorageList);
	//CRecord rs(&pMF->m_db);
	//CString szSQL, szWhere;
	//m_wndStorageList.BeginLoad();
	//int nCount = 0;
	//szSQL.Format(_T("SELECT msl_storage_id as id, msl_name as name FROM m_storagelist WHERE 1=1 %s ORDER BY msl_storage_id "), szWhere);
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
/*void CPMSGeneralDepartmentUsage::OnStorageAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */

long CPMSGeneralDepartmentUsage::OnDeptListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndDeptList.BeginLoad(); 
	m_wndDeptList.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" select sd_id as id, sd_name as name from sys_dept where sd_type IN ('DT', 'KB', 'KD') OR sd_id IN ('C10', 'C8') order by  sd_name"));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDeptList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndDeptList.EndLoad(); 
	return nCount;
}
void CPMSGeneralDepartmentUsage::OnItemGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSGeneralDepartmentUsage::OnItemGroupSelendok(){
	 
}
/*void CPMSGeneralDepartmentUsage::OnItemGroupSetfocus(){
	
}*/
/*void CPMSGeneralDepartmentUsage::OnItemGroupKillfocus(){
	
}*/
long CPMSGeneralDepartmentUsage::OnItemGroupLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadProductCategoryList(&m_wndItemGroup, m_szItemGroupKey);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndItemGroup.IsSearchKey() && !m_szItemGroupKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szItemGroupKey
};
	m_wndItemGroup.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndItemGroup.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CPMSGeneralDepartmentUsage::OnItemGroupAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSGeneralDepartmentUsage::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CReportSection* rptDetail, *rptHeader;
	CString szSQL, tmpStr, szNewType, szOldType, szTemp;
	CRecord rs(&pMF->m_db);
	CStringArray arrDat;
	CString szCurDte;
	int nIdx = 1, j = 0, nSel = 0;
	long double nTotal[7], nTotal_Surgery[7];
	if (!rpt.Init(_T("Reports\\HMS\\PM_BAOCAOTONGKETSUDUNGTHEODONVI.RPT")))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONSTOP | MB_OK);
		return;
	}
	for (int i = 0; i < 7; i++)
	{
		nTotal[i] = 0;
		nTotal_Surgery[i] = 0;
	}
	//Header Section
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	tmpStr = _T("To\xE0n \x62\x1ED9");
	nSel = m_wndStorageList.GetCurSel();
	if (!m_szStorageKey.IsEmpty())
		rptHeader->SetValue(_T("StockName"), m_wndStorageList.GetItemText(nSel, 1));
	else
		rptHeader->SetValue(_T("StockName"), tmpStr);
	if (!m_szItemGroupKey.IsEmpty())
		rptHeader->SetValue(_T("Type"), m_wndItemGroup.GetCurrent(1));
	else
		rptHeader->SetValue(_T("Type"), tmpStr);
	tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), 
	CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	//Detail
	arrDat.Add(_T("deptid"));
	arrDat.Add(_T("sol"));
	arrDat.Add(_T("pol"));
	arrDat.Add(_T("solins"));
	arrDat.Add(_T("othins"));
	arrDat.Add(_T("ser"));
	arrDat.Add(_T("oth"));
	arrDat.Add(_T("ttl"));
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("cat"), szNewType);
		if (szNewType != szOldType)
		{
			if (nTotal_Surgery[6] > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("tongso"), _T("\x44\xE0nh \x63ho PTTT"));
				for (int i = 0 ; i < 7;i++)
				{
					tmpStr.Format(_T("%f"), nTotal_Surgery[i]);
					szTemp.Format(_T("s%d"), i+3);
					rptDetail->SetValue(szTemp, tmpStr);
					nTotal_Surgery[i] = 0;
				}
			}
			if (nTotal[6] > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("tongso"), _T("T\x1ED5ng \x63\x1ED9ng"));
				for (int i = 0 ; i < 7;i++)
				{
					
					tmpStr.Format(_T("%.f"), nTotal[i]);
					szTemp.Format(_T("s%d"), i+3);
					rptDetail->SetValue(szTemp, tmpStr);
					nTotal[i] = 0;
				}
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			TranslateString(szNewType, tmpStr);
			rptDetail->SetValue(_T("GroupName"), tmpStr);	
			szOldType = szNewType;
		}
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		for (int i = 0; i < arrDat.GetCount(); i++)
		{
			j = i + 2;
			rs.GetValue(arrDat.GetAt(i), tmpStr);
			if (i > 0)
			{
				nTotal[i-1] += str2long(tmpStr);
				szTemp.Format(_T("%s_surgery"), arrDat.GetAt(i));
				nTotal_Surgery[i-1] += str2long(rs.GetValue(szTemp));
			}
			if (tmpStr != _T("0"))
				rptDetail->SetValue(int2str(j), tmpStr);
		}
		rs.MoveNext();
	}
	if (nTotal_Surgery[6] > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("tongso"), _T("\x44\xE0nh \x63ho PTTT"));
		for (int i = 0 ; i < 7;i++)
		{
			tmpStr.Format(_T("%f"), nTotal_Surgery[i]);
			szTemp.Format(_T("s%d"), i+3);
			rptDetail->SetValue(szTemp, tmpStr);
		}
	}
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	rptDetail->SetValue(_T("tongso"), _T("T\x1ED5ng \x63\x1ED9ng"));
	for (int i = 0 ; i < 7;i++)
	{
		tmpStr.Format(_T("%f"), nTotal[i]);
		szTemp.Format(_T("s%d"), i+3);
		rptDetail->SetValue(szTemp, tmpStr);
	}

	//Footer
	szCurDte = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szCurDte.Right(2), szCurDte.Mid(5, 2), szCurDte.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CPMSGeneralDepartmentUsage::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr, szNewType, szOldType, szTemp;
	CRecord rs(&pMF->m_db);
	int nIdx = 1, j = 0;
	double nTemp = 0;
	long double nTotal[7], nTotal_Surgery[7], nTotal_Old[7];
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONASTERISK | MB_OK);
		return;
	}
	for (int i = 0; i < 7; i++)
	{
		nTotal[i] = 0;
		nTotal_Surgery[i] = 0;
		nTotal_Old[i] = 0;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	//Header Section
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 60);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 4);
	xls.SetCellText(nCol, nRow + 3, _T("T\x1ED4NG K\x1EBET S\x1EEC \x44\x1EE4NG THU\x1ED0\x43 T\x1EA0I \x110\x1A0N V\x1ECA"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 4);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	CStringArray arrCol, arrDat;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn \x111\x1A1n v\x1ECB"));
	arrCol.Add(_T("\x42\x1ED9 \x111\x1ED9i"));
	arrCol.Add(_T("\x43h\xEDnh s\xE1\x63h"));
	arrCol.Add(_T("\x42H Qu\xE2n"));
	arrCol.Add(_T("\x42H Kh\xE1\x63"));
	arrCol.Add(_T("\x44\x1ECB\x63h v\x1EE5"));
	arrCol.Add(_T("Kh\xE1\x63"));
	arrCol.Add(_T("\x43\x1ED9ng"));
	for (int i = 0; i < arrCol.GetCount(); i++)
		xls.SetCellText(nCol + i, nRow + 5, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 11);
	nRow = 6;
	arrDat.Add(_T("deptid"));
	arrDat.Add(_T("sol"));
	arrDat.Add(_T("pol"));
	arrDat.Add(_T("solins"));
	arrDat.Add(_T("othins"));
	arrDat.Add(_T("ser"));
	arrDat.Add(_T("oth"));
	arrDat.Add(_T("ttl"));
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("cat"), szNewType);
		if (szNewType != szOldType)
		{
			if (nTotal_Surgery[6] > 0)
			{
				xls.SetCellText(nCol + 1, nRow, _T("\x44\xE0nh \x63ho PTTT"), FMT_TEXT | FMT_CENTER, true);
				for (int i = 0; i < 7; i ++)
				{
					tmpStr.Format(_T("%f"), nTotal_Surgery[i]);
					xls.SetCellText(nCol + i + 2, nRow, tmpStr, FMT_NUMBER1, true);
					nTotal_Surgery[i] = 0;
				}
				nRow++;
			}
			if (nTotal[6] > 0)
			{
				xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
				for (int i = 0; i < 7; i ++)
				{
					tmpStr.Format(_T("%f"), nTotal[i]);
					xls.SetCellText(nCol + i + 2, nRow, tmpStr, FMT_NUMBER1, true);
					nTotal[i] = 0;
				}
				nRow++;
			}
			xls.SetCellMergedColumns(nCol, nRow, 9);
			TranslateString(szNewType, tmpStr);
			xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT, true);
			szOldType = szNewType;
			nRow++;
		}
		xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, rs.GetValue(_T("deptid")), FMT_TEXT);
		for (int i = 1; i < arrDat.GetCount(); i++)
		{
			j = i + 1;
			xls.SetCellText(nCol+i+1, nRow, rs.GetValue(arrDat.GetAt(i)), FMT_NUMBER1);
			szTemp.Format(_T("%s_surgery"), arrDat.GetAt(i));
			rs.GetValue(szTemp, nTemp);
			nTotal_Surgery[i-1] += nTemp;
			nTemp = 0;
			rs.GetValue(arrDat.GetAt(i), nTemp);
			nTotal_Old[i-1] += str2long(rs.GetValue(arrDat.GetAt(i)));
			nTotal[i-1] += nTemp;
			_debug(_T("%f|%f|%f"), nTemp, nTotal[i-1], nTotal_Old[i-1]);
		}
		nRow++;
		rs.MoveNext();
	}
	
	if (nTotal_Surgery[6] > 0)
	{
		xls.SetCellText(nCol + 1, nRow, _T("\x44\xE0nh \x63ho PTTT"), FMT_TEXT | FMT_CENTER, true);
		for (int i = 0; i < 7; i ++)
		{
			tmpStr.Format(_T("%f"), nTotal_Surgery[i]);
			xls.SetCellText(nCol + i + 2, nRow, tmpStr, FMT_NUMBER1, true);
		}
		nRow++;
	}
	if (nTotal[6] > 0)
	{
		xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
		for (int i = 0; i < 7; i ++)
		{
			tmpStr.Format(_T("%f"), nTotal[i]);
			xls.SetCellText(nCol + i + 2, nRow, tmpStr, FMT_NUMBER1, true);
		}
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\TONGKETSUDUNGTHUOCTAIDONVI.xls"));
} 

CString CPMSGeneralDepartmentUsage::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szWhere1, szWhere2, szWhere3, szWhere31, szWhere4, szWhere5, szStorageStr, szDeptStr, szIsDDO, szTemp;
	for (int i = 0; i < m_wndDeptList.GetItemCount(); i++)
	{
		if (m_wndDeptList.GetCheck(i))
		{
			m_wndDeptList.SetCurSel(i);
			szTemp = m_wndDeptList.GetItemText(i, 0);
			if (!szDeptStr.IsEmpty())
				szDeptStr += _T(", ");
			szDeptStr.AppendFormat(_T("'%s'"), szTemp);
			
		}
	}
	for (int i = 0; i < m_wndStorageList.GetItemCount(); i++)
	{
		if (m_wndStorageList.GetCheck(i))
		{
			m_wndStorageList.SetCurSel(i);
			if (!szStorageStr.IsEmpty())
				szStorageStr += _T(", ");
			szStorageStr.AppendFormat(_T("'%s'"), m_wndStorageList.GetItemText(i, 0));
		}
	}
	if (m_bOnlyDDO)
		szIsDDO.Format(_T(" AND 0=1"));
	else
		szIsDDO.Empty();
	szWhere1.Format(_T(" AND mt_approveddate BETWEEN to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss')"), m_szFromDate, m_szToDate);
	szWhere2.Format(_T(" AND hpo_approvedate BETWEEN to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss')"), m_szFromDate, m_szToDate);
	szWhere3.Format(_T(" AND so_approveddate BETWEEN to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss')"), m_szFromDate, m_szToDate);
	szWhere4.Format(_T(" AND pro.hpo_approveddate BETWEEN to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss')"), m_szFromDate, m_szToDate);
	szWhere5.Format(_T(" AND mt_approveddate BETWEEN to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd|hh24:mi:ss')"), m_szFromDate, m_szToDate);
	if (!szStorageStr.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" AND mt_storage_id IN (%s)"), szStorageStr);
		szWhere2.AppendFormat(_T(" AND hpo_storage_id IN (%s)"), szStorageStr);
		szWhere3.AppendFormat(_T(" AND so_storage_id IN (%s)"), szStorageStr);
		szWhere4.AppendFormat(_T(" AND pro.hpo_storage_id IN (%s)"), szStorageStr);
		szWhere5.AppendFormat(_T(" AND mt_storage_to_id IN (%s)"), szStorageStr);
	}
	if (!szDeptStr.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" AND mt_department_to_id IN (%s)"), szDeptStr);
		szWhere2.AppendFormat(_T(" AND hpo_deptid IN (%s)"), szDeptStr);
		szWhere31.AppendFormat(_T(" AND deptid IN (%s)"), szDeptStr);
		szWhere4.AppendFormat(_T(" AND pro.hpo_deptid IN (%s)"), szDeptStr);
		szWhere5.AppendFormat(_T(" AND mt_department_id IN (%s)"), szDeptStr);
	}
	if (!m_szItemGroupKey.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" AND product_categoryid = %d"), ToInt(m_szItemGroupKey));
		szWhere2.AppendFormat(_T(" AND product_categoryid = %d"), ToInt(m_szItemGroupKey));
		szWhere3.AppendFormat(_T(" AND product_categoryid = %d"), ToInt(m_szItemGroupKey));
		szWhere4.AppendFormat(_T(" AND product_categoryid = %d"), ToInt(m_szItemGroupKey));
		szWhere5.AppendFormat(_T(" AND product_categoryid = %d"), ToInt(m_szItemGroupKey));
	}
	szWhere1.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	szWhere2.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	szWhere3.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	szWhere4.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	szWhere5.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	szSQL.Format(_T(" SELECT cat, deptid,") \
				_T("   get_departmentname(deptid) as deptname,") \
				_T("   Sum(solamt)    AS sol,") \
				_T("   Sum(polamt)    AS pol,") \
				_T("   Sum(solinsamt) AS solins,") \
				_T("   Sum(othinsamt) AS othins,") \
				_T("   Sum(seramt)    AS ser,") \
				_T("   Sum(othamt)    AS oth,") \
				_T("   Sum(solamt+polamt+solinsamt+othinsamt+seramt") \
				_T("       +othamt)   AS ttl,") \
				_T("   Sum(solamt_surgery)    AS sol_surgery,") \
				_T("   Sum(polamt_surgery)    AS pol_surgery,") \
				_T("   Sum(solinsamt_surgery) AS solins_surgery,") \
				_T("   Sum(othinsamt_surgery) AS othins_surgery,") \
				_T("   Sum(seramt_surgery)    AS ser_surgery,") \
				_T("   Sum(othamt_surgery)    AS oth_surgery,") \
				_T("   Sum(solamt_surgery+polamt_surgery+solinsamt_surgery+othinsamt_surgery+seramt_surgery") \
				_T("       +othamt_surgery)   AS ttl_surgery") \
				_T(" FROM   (SELECT") \
				_T("   'Delivery' as cat,") \
				_T("   hpo_deptid deptid,") \
				_T("   Sum(solamt)    AS solamt,") \
				_T("   Sum(polamt)    AS polamt,") \
				_T("   Sum(solinsamt) AS solinsamt,") \
				_T("   Sum(othinsamt) AS othinsamt,") \
				_T("   Sum(seramt)    AS seramt,") \
				_T("   Sum(otheramt)  AS othamt,") \
				_T("   Sum(solamt_surgery)    AS solamt_surgery,") \
				_T("   Sum(polamt_surgery)    AS polamt_surgery,") \
				_T("   Sum(solinsamt_surgery) AS solinsamt_surgery,") \
				_T("   Sum(othinsamt_surgery) AS othinsamt_surgery,") \
				_T("   Sum(seramt_surgery)    AS seramt_surgery,") \
				_T("   Sum(othamt_surgery)  AS othamt_surgery") \
				_T(" FROM   (SELECT") \
				_T("           hpo_deptid,") \
				_T("           CASE WHEN hpo_object_id=1 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS solamt,") \
				_T("           CASE WHEN hpo_object_id=3 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS polamt,") \
				_T("           CASE WHEN hpo_object_id=2 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS solinsamt,") \
				_T("           CASE WHEN hpo_object_id=4 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS othinsamt,") \
				_T("           CASE WHEN hpo_object_id=7 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS seramt,") \
				_T("           CASE WHEN hpo_object_id IN (5, 6, 8, 9,") \
				_T("                              10, 11, 12) THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS otheramt,") \
				_T("           CASE WHEN hpo_object_id=1 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS solamt_surgery,") \
				_T("           CASE WHEN hpo_object_id=3 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS polamt_surgery,") \
				_T("           CASE WHEN hpo_object_id=2 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS solinsamt_surgery,") \
				_T("           CASE WHEN hpo_object_id=4 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS othinsamt_surgery,") \
				_T("           CASE WHEN hpo_object_id=7 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS seramt_surgery,") \
				_T("           CASE WHEN hpo_object_id IN (5, 6, 8, 9,") \
				_T("                              10, 11, 12) AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS othamt_surgery") \
				_T("         FROM   hms_ipharmaorder") \
				_T("         LEFT JOIN hms_doc ON (hd_docno=hpo_docno)") \
				_T("         LEFT JOIN hms_ipharmaorderline ON (hpo_orderid=hpol_orderid)") \
				_T("		 LEFT JOIN m_transaction ON (mt_transaction_id = hpo_transaction_id)") \
				_T("		 LEFT JOIN m_product_view ON (product_id = hpol_product_id)") \
				_T("         WHERE  hpo_orderstatus = 'A' AND hpo_ordertype IN ('D', 'M', 'B') ") \
				_T("			%s %s) tbl1") \
				_T(" GROUP  BY hpo_deptid") \
				_T("		 UNION ALL") \
				_T(" SELECT") \
				_T("   'Delivery' as cat,") \
				_T("   hpo_deptid,") \
				_T("   Sum(solamt)    AS solamt,") \
				_T("   Sum(polamt)    AS polamt,") \
				_T("   Sum(solinsamt) AS solinsamt,") \
				_T("   Sum(othinsamt) AS othinsamt,") \
				_T("   Sum(seramt)    AS seramt,") \
				_T("   Sum(otheramt)  AS otheramt,") \
				_T("   Sum(solamt_surgery)    AS solamt_surgery,") \
				_T("   Sum(polamt_surgery)    AS polamt_surgery,") \
				_T("   Sum(solinsamt_surgery) AS solinsamt_surgery,") \
				_T("   Sum(othinsamt_surgery) AS othinsamt_surgery,") \
				_T("   Sum(seramt_surgery)    AS seramt_surgery,") \
				_T("   Sum(othamt_surgery)  AS othamt_surgery") \
				_T(" FROM   (SELECT") \
				_T("           hpo_deptid,") \
				_T("           CASE WHEN hpo_object_id=1 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS solamt,") \
				_T("           CASE WHEN hpo_object_id=3 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS polamt,") \
				_T("           CASE WHEN hpo_object_id=2 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS solinsamt,") \
				_T("           CASE WHEN hpo_object_id=4 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS othinsamt,") \
				_T("           CASE WHEN hpo_object_id=7 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS seramt,") \
				_T("           CASE WHEN hpo_object_id IN (5, 6, 8, 9,") \
				_T("                              10, 11, 12) THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS otheramt,") \
				_T("           CASE WHEN hpo_object_id=1 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS solamt_surgery,") \
				_T("           CASE WHEN hpo_object_id=3 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS polamt_surgery,") \
				_T("           CASE WHEN hpo_object_id=2 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS solinsamt_surgery,") \
				_T("           CASE WHEN hpo_object_id=4 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS othinsamt_surgery,") \
				_T("           CASE WHEN hpo_object_id=7 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS seramt_surgery,") \
				_T("           CASE WHEN hpo_object_id IN (5, 6, 8, 9,") \
				_T("                              10, 11, 12) AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS othamt_surgery") \
				_T("         FROM   hms_pharmaorder") \
				_T("         LEFT JOIN hms_doc ON (hd_docno=hpo_docno)") \
				_T("         LEFT JOIN hms_pharmaorderline ON (hpo_orderid=hpol_orderid)") \
				_T("		 LEFT JOIN m_product_view ON (product_id = hpol_product_id)") \
				_T("         WHERE  hpo_orderstatus = 'A' ") \
				_T("			%s %s") \
				_T("		 UNION ALL ") \
				_T("		 SELECT") \
				_T("           hpo_deptid,") \
				_T("           CASE WHEN hpo_object_id=1 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS solamt,") \
				_T("           CASE WHEN hpo_object_id=3 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS polamt,") \
				_T("           CASE WHEN hpo_object_id=2 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS solinsamt,") \
				_T("           CASE WHEN hpo_object_id=4 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS othinsamt,") \
				_T("           CASE WHEN hpo_object_id=7 THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS seramt,") \
				_T("           CASE WHEN hpo_object_id IN (5, 6, 8, 9,") \
				_T("                              10, 11, 12) THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS otheramt,") \
				_T("           CASE WHEN hpo_object_id=1 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS solamt_surgery,") \
				_T("           CASE WHEN hpo_object_id=3 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS polamt_surgery,") \
				_T("           CASE WHEN hpo_object_id=2 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS solinsamt_surgery,") \
				_T("           CASE WHEN hpo_object_id=4 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS othinsamt_surgery,") \
				_T("           CASE WHEN hpo_object_id=7 AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS seramt_surgery,") \
				_T("           CASE WHEN hpo_object_id IN (5, 6, 8, 9,") \
				_T("                              10, 11, 12) AND hpo_ordertype = 'M' THEN hpol_qtyorder*hpol_unitprice") \
				_T("           ELSE 0") \
				_T("           END AS othamt_surgery") \
				_T("         FROM   hms_ipharmaorder") \
				_T("         LEFT JOIN hms_doc ON (hd_docno=hpo_docno)") \
				_T("         LEFT JOIN hms_ipharmaorderline ON (hpo_orderid=hpol_orderid)") \
				_T("		 LEFT JOIN m_transaction ON (mt_transaction_id = hpo_transaction_id)") \
				_T("		 LEFT JOIN m_product_view ON (product_id = hpol_product_id)") \
				_T("         WHERE  hpo_orderstatus = 'A' AND hpo_ordertype = 'C' ") \
				_T("			%s %s) tbl") \
				_T(" GROUP  BY hpo_deptid ") \
				_T(" UNION ALL") \
				_T(" SELECT") \
				_T("   'Delivery' as cat,") \
				_T("   mt_department_to_id,") \
				_T("   Sum(mtl_qtysold*mtl_saleprice)    AS solamt,") \
				_T("   Sum(mtl_qtypolicy*mtl_saleprice)    AS polamt,") \
				_T("   Sum(mtl_qtysoldins*mtl_saleprice) AS solinsamt,") \
				_T("   Sum(mtl_qtyotherins*mtl_saleprice) AS othinsamt,") \
				_T("   Sum(mtl_qtyservice*mtl_saleprice)    AS seramt,") \
				_T("   Sum(mtl_qtyother*mtl_saleprice)  AS otheramt,") \
				_T("   0, 0, 0, 0, 0, 0") \
				_T(" FROM m_transaction") \
				_T(" LEFT JOIN m_transactionline ON (mt_transaction_id = mtl_transaction_id)") \
				_T(" LEFT JOIN m_product_view ON (mtl_product_id = product_id)") \
				_T(" WHERE mt_doctype = 'DDO' AND mt_status = 'A' ") \
				_T("	%s") \
				_T(" GROUP BY mt_department_to_id") \
				_T(" UNION ALL") \
				_T(" SELECT cat, deptid, 0, 0, 0, 0, sum(seramt) seramt, 0, ") \
				_T(" 0, 0, 0, 0, 0, 0") \
				_T(" FROM(") \
				_T("	SELECT") \
				_T("		'Delivery' as cat,") \
				_T("		case when NVL(so_partner_type_id, 'W') = 'W' THEN hd_admitdept else so_partneraddress end deptid,") \
				_T("		0 AS solamt,") \
				_T("		0 AS polamt,") \
				_T("		0 solinsamt,") \
				_T("		0 othinsamt,") \
				_T("		sol_qtydelivery*sol_unitprice AS seramt,") \
				_T("		0  AS otheramt,") \
				_T("		0, 0, 0, 0, 0, 0") \
				_T(" FROM sale_order") \
				_T(" LEFT JOIN hms_doc ON (hd_docno = so_docno)") \
				_T(" LEFT JOIN sale_orderline ON (so_sale_order_id = sol_sale_order_id)") \
				_T(" LEFT JOIN m_product_view ON (sol_product_id = product_id)") \
				_T(" WHERE so_doctype = 'SOO' AND so_status = 'A' ") \
				_T("	%s %s)") \
				_T(" WHERE 1=1 ") \
				_T("	%s") \
				_T(" GROUP BY deptid ") \
				_T(" UNION ALL ") \
				_T(" SELECT 'Return' AS cat, ") \
				_T("        hpo_deptid, ") \
				_T("        SUM(solamt) AS solamt, ") \
				_T("        SUM(polamt) AS polamt, ") \
				_T("        SUM(solinsamt) AS solinsamt, ") \
				_T("        SUM(othinsamt) AS othinsamt, ") \
				_T("        SUM(seramt) AS seramt, ") \
				_T("        SUM(otheramt) AS otheramt, ") \
				_T("        SUM(solamt_surgery) AS solamt_surgery, ") \
				_T("        SUM(polamt_surgery) AS polamt_surgery, ") \
				_T("        SUM(solinsamt_surgery) AS solinsamt_surgery, ") \
				_T("        SUM(othinsamt_surgery) AS othinsamt_surgery, ") \
				_T("        SUM(seramt_surgery) AS seramt_surgery, ") \
				_T("        SUM(othamt_surgery) AS othamt_surgery ") \
				_T(" FROM   (SELECT    pro.hpo_deptid, ") \
				_T("                   CASE WHEN hpo_object_id = 1 THEN hpol_qtyreturn * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS solamt, ") \
				_T("                   CASE WHEN hpo_object_id = 3 THEN hpol_qtyreturn * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS polamt, ") \
				_T("                   CASE WHEN hpo_object_id = 2 THEN hpol_qtyreturn * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS solinsamt, ") \
				_T("                   CASE WHEN hpo_object_id = 4 THEN hpol_qtyreturn * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS othinsamt, ") \
				_T("                   CASE WHEN hpo_object_id = 7 THEN hpol_qtyreturn * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS seramt, ") \
				_T("                   CASE WHEN hpo_object_id IN ( 5, 6, 8, 9, ") \
				_T("                                           10, 11, 12 ) THEN hpol_qtyreturn * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS otheramt, ") \
				_T("                   CASE WHEN hpo_object_id = 1 AND hpo_ordertype = 'M' THEN hpol_qtyreturn * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS solamt_surgery, ") \
				_T("                   CASE WHEN hpo_object_id = 3 AND hpo_ordertype = 'M' THEN hpol_qtyreturn * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS polamt_surgery, ") \
				_T("                   CASE WHEN hpo_object_id = 2 AND hpo_ordertype = 'M' THEN hpol_qtyreturn * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS solinsamt_surgery, ") \
				_T("                   CASE WHEN hpo_object_id = 4 AND hpo_ordertype = 'M' THEN hpol_qtyreturn * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS othinsamt_surgery, ") \
				_T("                   CASE WHEN hpo_object_id = 7 AND hpo_ordertype = 'M' THEN hpol_qtyreturn * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS seramt_surgery, ") \
				_T("                   CASE WHEN hpo_object_id IN ( 5, 6, 8, 9, ") \
				_T("                                           10, 11, 12 ) AND hpo_ordertype = 'M' THEN hpol_qtyreturn * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS othamt_surgery ") \
				_T("         FROM      hms_pharmareturnorder pro ") \
				_T("         LEFT JOIN hms_pharmaorder po ON ( po.hpo_docno = pro.hpo_docno AND po.hpo_orderid = pro.hpo_orderid ) ") \
				_T("         LEFT JOIN hms_doc ON ( hd_docno = pro.hpo_docno ) ") \
				_T("         LEFT JOIN hms_pharmareturnorder_line prol ON ( pro.hpo_orderid = hpol_orderid ) ") \
				_T("         LEFT JOIN m_productitem_view ON ( product_item_id = prol.hpol_product_item_id ) ") \
				_T("         WHERE     pro.hpo_status = 'A' ") \
				_T("			%s %s") \
				_T("         UNION ALL ") \
				_T("         SELECT    hpo_deptid deptid, ") \
				_T("                   CASE WHEN hpo_object_id = 1 THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS solamt, ") \
				_T("                   CASE WHEN hpo_object_id = 3 THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS polamt, ") \
				_T("                   CASE WHEN hpo_object_id = 2 THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS solinsamt, ") \
				_T("                   CASE WHEN hpo_object_id = 4 THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS othinsamt, ") \
				_T("                   CASE WHEN hpo_object_id = 7 THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS seramt, ") \
				_T("                   CASE WHEN hpo_object_id IN ( 5, 6, 8, 9, ") \
				_T("                                           10, 11, 12 ) THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS otheramt, ") \
				_T("                   CASE WHEN hpo_object_id = 1 AND hpo_ordertype = 'M' THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS solamt_surgery, ") \
				_T("                   CASE WHEN hpo_object_id = 3 AND hpo_ordertype = 'M' THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS polamt_surgery, ") \
				_T("                   CASE WHEN hpo_object_id = 2 AND hpo_ordertype = 'M' THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS solinsamt_surgery, ") \
				_T("                   CASE WHEN hpo_object_id = 4 AND hpo_ordertype = 'M' THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS othinsamt_surgery, ") \
				_T("                   CASE WHEN hpo_object_id = 7 AND hpo_ordertype = 'M' THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS seramt_surgery, ") \
				_T("                   CASE WHEN hpo_object_id IN ( 5, 6, 8, 9, ") \
				_T("                                           10, 11, 12 ) AND hpo_ordertype = 'M' THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                   ELSE 0 ") \
				_T("                   END AS othamt_surgery ") \
				_T("         FROM      hms_ipharmaorder ") \
				_T("         LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T("         LEFT JOIN m_productitem_view ON ( hpol_product_item_id = product_item_id ) ") \
				_T("         LEFT JOIN m_transaction ON ( mt_transaction_id = hpol_retorder_id ) ") \
				_T("         WHERE     hpol_qtyreverse > 0 AND mt_status = 'A' ") \
				_T("			%s %s) ") \
				_T(" GROUP  BY hpo_deptid )") \
				_T(" GROUP  BY cat, deptid") \
				_T(" ORDER  BY cat, deptid"), szWhere1, szIsDDO, szWhere2, szIsDDO, szWhere2, szIsDDO, 
				szWhere1, szWhere3, szIsDDO, szWhere31, szWhere4, szIsDDO, szWhere5, szIsDDO);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}BEGIN_MESSAGE_MAP(CPMSGeneralDepartmentUsage, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSGeneralDepartmentUsage::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
