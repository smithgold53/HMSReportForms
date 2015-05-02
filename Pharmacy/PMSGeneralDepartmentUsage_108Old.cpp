#include "stdafx.h"
#include "PMSGeneralDepartmentUsage_108Old.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include ".\pmsgeneraldepartmentusage.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnToDateCheckValue();
} 
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSGeneralDepartmentUsage_108Old* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnStorageAddNew();
}*/
static long _OnDeptListLoadDataFnc(CWnd *pWnd){
	return ((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnDeptListLoadData();
}
static void _OnItemGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSGeneralDepartmentUsage_108Old* )pWnd)->OnItemGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnItemGroupSelendokFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnItemGroupSelendok();
}
/*static void _OnItemGroupSetfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnItemGroupKillfocus();
}*/
/*static void _OnItemGroupKillfocusFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnItemGroupKillfocus();
}*/
static long _OnItemGroupLoadDataFnc(CWnd *pWnd){
	return ((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnItemGroupLoadData();
}
/*static void _OnItemGroupAddNewFnc(CWnd *pWnd){
	((CPMSGeneralDepartmentUsage_108Old *)pWnd)->OnItemGroupAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSGeneralDepartmentUsage_108Old *pVw = (CPMSGeneralDepartmentUsage_108Old *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSGeneralDepartmentUsage_108Old *pVw = (CPMSGeneralDepartmentUsage_108Old *)pWnd;
	pVw->OnExportSelect();
} 
CPMSGeneralDepartmentUsage_108Old::CPMSGeneralDepartmentUsage_108Old(CWnd *pParent){

	m_nDlgWidth = 545;
	m_nDlgHeight = 155;
	SetDefaultValues();
}
CPMSGeneralDepartmentUsage_108Old::~CPMSGeneralDepartmentUsage_108Old(){
}
void CPMSGeneralDepartmentUsage_108Old::OnCreateComponents(){
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
void CPMSGeneralDepartmentUsage_108Old::OnInitializeComponents(){
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
void CPMSGeneralDepartmentUsage_108Old::OnSetWindowEvents(){
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
void CPMSGeneralDepartmentUsage_108Old::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStorageList.GetDlgCtrlID(), m_szStorageKey);
	DDX_TextEx(pDX, m_wndItemGroup.GetDlgCtrlID(), m_szItemGroupKey);
	DDX_Check(pDX, m_wndOnlyDDO.GetDlgCtrlID(), m_bOnlyDDO);

}
void CPMSGeneralDepartmentUsage_108Old::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStorageKey.Empty();
	m_szItemGroupKey.Empty();

}
int CPMSGeneralDepartmentUsage_108Old::SetMode(int nMode){
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
/*void CPMSGeneralDepartmentUsage_108Old::OnFromDateChange(){
	
} */
/*void CPMSGeneralDepartmentUsage_108Old::OnFromDateSetfocus(){
	
} */
/*void CPMSGeneralDepartmentUsage_108Old::OnFromDateKillfocus(){
	
} */
int CPMSGeneralDepartmentUsage_108Old::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSGeneralDepartmentUsage_108Old::OnToDateChange(){
	
} */
/*void CPMSGeneralDepartmentUsage_108Old::OnToDateSetfocus(){
	
} */
/*void CPMSGeneralDepartmentUsage_108Old::OnToDateKillfocus(){
	
} */
int CPMSGeneralDepartmentUsage_108Old::OnToDateCheckValue(){
	return 0;
} 
void CPMSGeneralDepartmentUsage_108Old::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSGeneralDepartmentUsage_108Old::OnStorageSelendok(){
	 
}
/*void CPMSGeneralDepartmentUsage_108Old::OnStorageSetfocus(){
	
}*/
/*void CPMSGeneralDepartmentUsage_108Old::OnStorageKillfocus(){
	
}*/
long CPMSGeneralDepartmentUsage_108Old::OnStorageLoadData(){
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
/*void CPMSGeneralDepartmentUsage_108Old::OnStorageAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */

long CPMSGeneralDepartmentUsage_108Old::OnDeptListLoadData(){
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
void CPMSGeneralDepartmentUsage_108Old::OnItemGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSGeneralDepartmentUsage_108Old::OnItemGroupSelendok(){
	 
}
/*void CPMSGeneralDepartmentUsage_108Old::OnItemGroupSetfocus(){
	
}*/
/*void CPMSGeneralDepartmentUsage_108Old::OnItemGroupKillfocus(){
	
}*/
long CPMSGeneralDepartmentUsage_108Old::OnItemGroupLoadData(){
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
/*void CPMSGeneralDepartmentUsage_108Old::OnItemGroupAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSGeneralDepartmentUsage_108Old::OnPrintPreviewSelect(){
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
	int nCount = rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONSTOP | MB_OK);
		if (nCount < 0)
			_msg(_T("%s"), szSQL);
		else
			_fmsg(_T("%s"), szSQL);
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
	arrDat.Add(_T("dept_id"));
	arrDat.Add(_T("sol"));
	arrDat.Add(_T("pol"));
	arrDat.Add(_T("solins"));
	arrDat.Add(_T("othins"));
	arrDat.Add(_T("ser"));
	arrDat.Add(_T("other"));
	arrDat.Add(_T("total"));
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("transaction_type"), szNewType);
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
			if (szNewType == _T("1"))
				tmpStr = _T("\x43\x1EA5p ph\xE1t");
			else
				tmpStr = _T("Tr\x1EA3 l\x1EA1i");
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
				_debug(_T("%s|%ld"), rs.GetValue(szTemp), str2long(rs.GetValue(szTemp)));
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
void CPMSGeneralDepartmentUsage_108Old::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr, szNewType, szOldType, szTemp;
	CRecord rs(&pMF->m_db);
	int nIdx = 1, j = 0;
	long double nTotal[7], nTotal_Surgery[7];
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
	arrDat.Add(_T("dept_id"));
	arrDat.Add(_T("sol"));
	arrDat.Add(_T("pol"));
	arrDat.Add(_T("solins"));
	arrDat.Add(_T("othins"));
	arrDat.Add(_T("ser"));
	arrDat.Add(_T("other"));
	arrDat.Add(_T("total"));
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("transaction_type"), szNewType);
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
			if (szNewType == _T("1"))
				tmpStr = _T("\x43\x1EA5p ph\xE1t");
			else
				tmpStr = _T("Tr\x1EA3 l\x1EA1i");
			xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT, true);
			szOldType = szNewType;
			nRow++;
		}
		xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, rs.GetValue(_T("dept_id")), FMT_TEXT);
		for (int i = 1; i < arrDat.GetCount(); i++)
		{
			j = i + 1;
			xls.SetCellText(nCol+i+1, nRow, rs.GetValue(arrDat.GetAt(i)), FMT_NUMBER1);
			szTemp.Format(_T("%s_surgery"), arrDat.GetAt(i));
			nTotal_Surgery[i-1] += str2long(rs.GetValue(szTemp));
			nTotal[i-1] += str2long(rs.GetValue(arrDat.GetAt(i)));
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

CString CPMSGeneralDepartmentUsage_108Old::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szWhere, szStorageStr, szDeptStr, szTemp, szIssueSQL, szReturnSQL;
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
	szWhere.Format(_T(" AND stt = 'A' AND approval_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')")
					, m_szFromDate, m_szToDate);
	if (!szStorageStr.IsEmpty())
		szWhere.AppendFormat(_T(" AND storage_id IN (%s)"), szStorageStr);
	if (!szDeptStr.IsEmpty())
		szWhere.AppendFormat(_T(" AND dept_id IN (%s)"), szDeptStr);
	if (!m_szItemGroupKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND product_categoryid = '%s'"), m_szItemGroupKey);
	szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	if (m_bOnlyDDO)
		szWhere.AppendFormat(_T(" AND doc_type = 'DDO'"));
	szIssueSQL.Format(_T(" SELECT    1 AS transaction_type, ") \
	_T("           Dept_id, ") \
	_T("           SUM(sol_amt) sol, ") \
	_T("           SUM(pol_amt) pol, ") \
	_T("           SUM(solins_amt) solins, ") \
	_T("           SUM(othins_amt) othins, ") \
	_T("           SUM(ser_amt) ser, ") \
	_T("           SUM(other_amt) other, ") \
	_T("           SUM(total_amt) total, ") \
	_T("           SUM(sol_surgeryamt) sol_surgery, ") \
	_T("           SUM(pol_surgeryamt) pol_surgery, ") \
	_T("           SUM(solins_surgeryamt) solins_surgery, ") \
	_T("           SUM(othins_surgeryamt) othins_surgery, ") \
	_T("           SUM(ser_surgeryamt) ser_surgery, ") \
	_T("           SUM(other_surgeryamt) other_surgery, ") \
	_T("           SUM(total_surgeryamt) total_surgery ") \
	_T(" FROM     (SELECT Dept_id, ") \
	_T("                  Storage_id, ") \
	_T("                  Approval_date, ") \
	_T("                  CASE WHEN object_id = 1 THEN line_qty * unit_price ") \
	_T("                  ELSE 0 ") \
	_T("                  END sol_amt, ") \
	_T("                  CASE WHEN object_id = 3 THEN line_qty * unit_price ") \
	_T("                  ELSE 0 ") \
	_T("                  END pol_amt, ") \
	_T("                  CASE WHEN object_id = 2 THEN line_qty * unit_price ") \
	_T("                  ELSE 0 ") \
	_T("                  END solins_amt, ") \
	_T("                  CASE WHEN object_id = 4 THEN line_qty * unit_price ") \
	_T("                  ELSE 0 ") \
	_T("                  END othins_amt, ") \
	_T("                  CASE WHEN object_id = 7 THEN line_qty * unit_price ") \
	_T("                  ELSE 0 ") \
	_T("                  END ser_amt, ") \
	_T("                  CASE WHEN object_id IN ( 5, 6, 8, 9, ") \
	_T("                                      10, 11, 12 ) THEN line_qty * unit_price ") \
	_T("                  ELSE 0 ") \
	_T("                  END other_amt, ") \
	_T("                  line_qty * unit_price total_amt, ") \
	_T("                  CASE WHEN object_id = 1 AND order_type = 'M' THEN line_qty * unit_price ") \
	_T("                  ELSE 0 ") \
	_T("                  END sol_surgeryamt, ") \
	_T("                  CASE WHEN object_id = 3 AND order_type = 'M' THEN line_qty * unit_price ") \
	_T("                  ELSE 0 ") \
	_T("                  END pol_surgeryamt, ") \
	_T("                  CASE WHEN object_id = 2 AND order_type = 'M' THEN line_qty * unit_price ") \
	_T("                  ELSE 0 ") \
	_T("                  END solins_surgeryamt, ") \
	_T("                  CASE WHEN object_id = 4 AND order_type = 'M' THEN line_qty * unit_price ") \
	_T("                  ELSE 0 ") \
	_T("                  END othins_surgeryamt, ") \
	_T("                  CASE WHEN object_id = 7 AND order_type = 'M' THEN line_qty * unit_price ") \
	_T("                  ELSE 0 ") \
	_T("                  END ser_surgeryamt, ") \
	_T("                  CASE WHEN object_id IN ( 5, 6, 8, 9, ") \
	_T("                                      10, 11, 12 ) AND order_type = 'M' THEN line_qty * unit_price ") \
	_T("                  ELSE 0 ") \
	_T("                  END other_surgeryamt, ") \
	_T("                  CASE WHEN order_type = 'M' THEN line_qty * unit_price ELSE 0 END total_surgeryamt, ") \
	_T("                  sitem_id, ") \
	_T("				  doc_type,") \
	_T("				  stt") \
	_T("           FROM   (SELECT    so_partneraddress AS dept_id, ") \
	_T("                             7 AS object_id, ") \
	_T("                             Sol_unitprice AS unit_price, ") \
	_T("                             Sol_qtyorder AS line_qty, ") \
	_T("                             So_storage_id storage_id, ") \
	_T("                             So_approveddate approval_date, ") \
	_T("                             Sol_product_item_id sitem_id, ") \
	_T("                             Cast('0' AS NVARCHAR2(1)) AS order_type, ") \
	_T("							 so_doctype doc_type,")
	_T("							 so_status stt") \
	_T("                   FROM      sale_order ") \
	_T("                   LEFT JOIN sale_orderline ON ( so_sale_order_id = sol_sale_order_id ) ") \
	_T("                   WHERE     so_doctype = 'SOO' AND so_partner_type_id = 'I'") \
	_T("                   UNION ALL ") \
	_T("                   SELECT    hpo_deptid, ") \
	_T("                             Hpo_object_id, ") \
	_T("                             Hpol_unitprice, ") \
	_T("                             Hpol_qtyorder, ") \
	_T("                             CASE WHEN coalesce(hpo_reforder_id, 0) > 0 THEN mt_storage_id ELSE hpo_storage_id END, ") \
	_T("                             CASE WHEN coalesce(hpo_reforder_id, 0) > 0 THEN mt_approveddate ELSE hpo_approvedate END, ") \
	_T("                             Hpol_product_item_id, ") \
	_T("                             hpo_ordertype, ") \
	_T("							 NVL(mt_doctype, 'P'),") \
	_T("							 CASE WHEN coalesce(hpo_reforder_id, 0) > 0 THEN mt_status ELSE hpo_orderstatus END") \
	_T("                   FROM      hms_pharmaorder_view ") \
	_T("                   LEFT JOIN hms_pharmaorderline_view2 ON ( hpo_orderid = hpol_orderid ) ") \
	_T("                   LEFT JOIN m_transaction ON ( mt_transaction_id = hpo_reforder_id )") \
	_T("				   WHERE hpo_orderstatus = 'A' AND hpo_ordertype <> 'C') tbl_pharma ") \
	_T("           UNION ALL ") \
	_T("           SELECT    mt_department_to_id, ") \
	_T("                     Mt_storage_id, ") \
	_T("                     Mt_approveddate, ") \
	_T("                     mtl_qtysold * mtl_saleprice, ") \
	_T("                     mtl_qtypolicy * mtl_saleprice, ") \
	_T("                     mtl_qtysoldins * mtl_saleprice, ") \
	_T("                     mtl_qtyotherins * mtl_saleprice, ") \
	_T("                     mtl_qtyservice * mtl_saleprice, ") \
	_T("                     mtl_qtyother * mtl_saleprice, ") \
	_T("                     mtl_qtyorder * mtl_saleprice, ") \
	_T("                     0, ") \
	_T("                     0, ") \
	_T("                     0, ") \
	_T("                     0, ") \
	_T("                     0, ") \
	_T("                     0, ") \
	_T("					 0, ") \
	_T("                     mtl_product_item_id, ") \
	_T("					 mt_doctype,") \
	_T("					 mt_status") \
	_T("           FROM      m_transaction ") \
	_T("           LEFT JOIN m_transactionline ON ( mt_transaction_id = mtl_transaction_id ) ") \
	_T("           WHERE     mt_doctype NOT IN ( 'CRO', 'DRO', 'SRO', 'PRO', ") \
	_T("                                         'DMO', 'MOV', 'BIO' ) AND mt_status = 'A' AND mtl_qtydelivery > 0) tbl_issue ") \
	_T(" LEFT JOIN m_productitem_view ON ( sitem_id = product_item_id ) ") \
	_T(" WHERE     dept_id IS NOT NULL %s") \
	_T(" GROUP     BY dept_id "), szWhere);

	szReturnSQL.Format(_T(" SELECT    16 AS transaction_type, ") \
	_T("           Dept_id, ") \
	_T("           SUM(sol) sol, ") \
	_T("           SUM(pol) pol, ") \
	_T("           SUM(solins) solins, ") \
	_T("           SUM(othins) othins, ") \
	_T("           SUM(ser) ser, ") \
	_T("           SUM(other) other, ") \
	_T("           SUM(total) total, ") \
	_T("           SUM(sol_surgery) sol_surgery, ") \
	_T("           SUM(pol_surgeryamt) pol_surgeryamt, ") \
	_T("           SUM(solins_surgery) solins_surgery, ") \
	_T("           SUM(othins_surgery) othins_surgery, ") \
	_T("           SUM(ser_surgery) ser_surgery, ") \
	_T("           SUM(other_surgery) other_surgery, ") \
	_T("           SUM(total_surgery) total_surgery ") \
	_T(" FROM (") \
	_T("		SELECT sitem_id,") \
	_T("			   dept_id,") \
	_T("			   approval_date,") \
	_T("			   storage_id,") \
	_T("			   doc_type,") \
	_T("			   stt,") \
	_T("           CASE WHEN object_id = 1 THEN line_qty * unit_price ") \
	_T("               ELSE 0 ") \
	_T("               END sol, ") \
	_T("           CASE WHEN object_id = 3 THEN line_qty * unit_price ") \
	_T("               ELSE 0 ") \
	_T("               END pol, ") \
	_T("           CASE WHEN object_id = 2 THEN line_qty * unit_price ") \
	_T("               ELSE 0 ") \
	_T("               END solins, ") \
	_T("           CASE WHEN object_id = 4 THEN line_qty * unit_price ") \
	_T("               ELSE 0 ") \
	_T("               END othins, ") \
	_T("           CASE WHEN object_id = 7 THEN line_qty * unit_price ") \
	_T("               ELSE 0 ") \
	_T("               END ser, ") \
	_T("           CASE WHEN object_id IN ( 5, 6, 8, 9, ") \
	_T("                                   10, 11, 12 ) THEN line_qty * unit_price ") \
	_T("               ELSE 0 ") \
	_T("               END other, ") \
	_T("           line_qty * unit_price total, ") \
	_T("           CASE WHEN object_id = 1 AND hpo_ordertype = 'M' THEN line_qty * unit_price ") \
	_T("               ELSE 0 ") \
	_T("               END sol_surgery, ") \
	_T("           CASE WHEN object_id = 3 AND hpo_ordertype = 'M' THEN line_qty * unit_price ") \
	_T("               ELSE 0 ") \
	_T("               END pol_surgeryamt, ") \
	_T("           CASE WHEN object_id = 2 AND hpo_ordertype = 'M' THEN line_qty * unit_price ") \
	_T("               ELSE 0 ") \
	_T("               END solins_surgery, ") \
	_T("           CASE WHEN object_id = 4 AND hpo_ordertype = 'M' THEN line_qty * unit_price ") \
	_T("               ELSE 0 ") \
	_T("               END othins_surgery, ") \
	_T("           CASE WHEN object_id = 7 AND hpo_ordertype = 'M' THEN line_qty * unit_price ") \
	_T("               ELSE 0 ") \
	_T("               END ser_surgery, ") \
	_T("           CASE WHEN object_id IN ( 5, 6, 8, 9, ") \
	_T("                                   10, 11, 12 ) AND hpo_ordertype = 'M' THEN line_qty * unit_price ") \
	_T("               ELSE 0 ") \
	_T("               END other_surgery, ") \
	_T("           CASE WHEN hpo_ordertype = 'M' THEN line_qty * unit_price ELSE 0 END total_surgery ") \
	_T("	FROM     (SELECT    hpo.Hpo_deptid AS dept_id, ") \
	_T("                     Hpo.Hpo_object_id AS object_id, ") \
	_T("                     Hpol_unitprice unit_price, ") \
	_T("                     Hpol_qtyreturn line_qty, ") \
	_T("                     Hpo_approveddate approval_date, ") \
	_T("                     hpo.Hpo_storage_id storage_id, ") \
	_T("                     hpol_product_item_id sitem_id, ") \
	_T("                     Hpo_ordertype, ") \
	_T("					 cast('P' as NVARCHAR2(1)) doc_type,") \
	_T("					 hpro.hpo_status stt") \
	_T("           FROM      hms_pharmareturnorder hpro ") \
	_T("           LEFT JOIN hms_pharmareturnorder_line ON ( hpro.hpo_pharmareturnorder_id = hpol_pharmareturnorder_id ) ") \
	_T("           LEFT JOIN hms_pharmaorder hpo ON ( hpo.hpo_docno = hpro.hpo_docno AND hpo.hpo_orderid = hpro.hpo_orderid ) ") \
	_T("           UNION ALL ") \
	_T("           SELECT    Hpo_deptid, ") \
	_T("                     Hpo_object_id, ") \
	_T("                     Hpol_unitprice, ") \
	_T("                     Hpol_qtyreverse, ") \
	_T("                     Mt_approveddate, ") \
	_T("                     Hpo_storage_id, ") \
	_T("                     Hpol_product_item_id sitem_id, ") \
	_T("                     Hpo_ordertype, ") \
	_T("					 mt_doctype,") \
	_T("					 mt_status") \
	_T("           FROM      hms_ipharmaorder ") \
	_T("           LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
	_T("           LEFT JOIN m_transaction ON ( mt_transaction_id = hpol_retorder_id ) ") \
	_T("           WHERE     hpol_qtyreverse > 0) tbl_patient ") \
	_T("  UNION ALL ") \
	_T("  SELECT    mtl_product_item_id, ") \
	_T("   mt_department_id, ") \
	_T("   mt_approveddate,") \
	_T("   Mt_storage_to_id, ") \
	_T("   mt_doctype,") \
	_T("   mt_status,") \
	_T("   mtl_qtysold * mtl_saleprice, ") \
	_T("   mtl_qtypolicy * mtl_saleprice, ") \
	_T("   mtl_qtysoldins * mtl_saleprice, ") \
	_T("   mtl_qtyotherins * mtl_saleprice, ") \
	_T("   mtl_qtyservice * mtl_saleprice, ") \
	_T("   mtl_qtyother * mtl_saleprice, ") \
	_T("   mtl_qtyorder * mtl_saleprice, ") \
	_T("   0, ") \
	_T("   0, ") \
	_T("   0, ") \
	_T("   0, ") \
	_T("   0, ") \
	_T("   0, ") \
	_T("   0 ") \
	_T("  FROM      m_transaction ") \
	_T("  LEFT JOIN m_transactionline ON ( mt_transaction_id = mtl_transaction_id ) ") \
	_T("  WHERE     mt_doctype = 'CRO' AND mt_status = 'A' AND mtl_qtydelivery > 0) tbl_return ") \
	_T(" LEFT JOIN m_productitem_view ON ( sitem_id = product_item_id ) ") \
	_T(" WHERE     dept_id IS NOT NULL %s") \
	_T(" GROUP     BY dept_id "), szWhere);

	szSQL.Format(_T("SELECT * FROM (") \
				_T(" %s") \
				_T(" UNION ALL") \
				_T(" %s)") \
				_T(" ORDER BY transaction_type, dept_id"), szIssueSQL, szReturnSQL);

	return szSQL;
}BEGIN_MESSAGE_MAP(CPMSGeneralDepartmentUsage_108Old, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSGeneralDepartmentUsage_108Old::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
