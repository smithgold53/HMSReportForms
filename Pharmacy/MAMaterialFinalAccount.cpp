#include "stdafx.h"
#include "MAMaterialFinalAccount.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CMAMaterialFinalAccount *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CMAMaterialFinalAccount *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CMAMaterialFinalAccount *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CMAMaterialFinalAccount *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CMAMaterialFinalAccount *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CMAMaterialFinalAccount *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CMAMaterialFinalAccount *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CMAMaterialFinalAccount *)pWnd)->OnToDateCheckValue();
} 
static long _OnStorageListLoadDataFnc(CWnd *pWnd){
	return ((CMAMaterialFinalAccount*)pWnd)->OnStorageListLoadData();
} 
static void _OnStorageListDblClickFnc(CWnd *pWnd){
	((CMAMaterialFinalAccount*)pWnd)->OnStorageListDblClick();
} 
static void _OnStorageListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CMAMaterialFinalAccount*)pWnd)->OnStorageListSelectChange(nOldItem, nNewItem);
} 
static int _OnStorageListDeleteItemFnc(CWnd *pWnd){
	 return ((CMAMaterialFinalAccount*)pWnd)->OnStorageListDeleteItem();
} 
static int _OnStorageListCheckAllFnc(CWnd *pWnd){
	 ((CMAMaterialFinalAccount*)pWnd)->OnStorageListCheckAll();
	 return 0;
}
static int _OnStorageListUncheckAllFnc(CWnd *pWnd){
	 ((CMAMaterialFinalAccount*)pWnd)->OnStorageListUncheckAll();
	 return 0;
} 
static int _OnDepartmentListCheckAllFnc(CWnd *pWnd){
	 ((CMAMaterialFinalAccount*)pWnd)->OnDepartmentListCheckAll();
	 return 0;
} 
static int _OnDepartmentListUncheckAllFnc(CWnd *pWnd){
	((CMAMaterialFinalAccount*)pWnd)->OnDepartmentListUncheckAll();
	return 0;
} 
static long _OnDepartmentListLoadDataFnc(CWnd *pWnd){
	return ((CMAMaterialFinalAccount*)pWnd)->OnDepartmentListLoadData();
} 
static void _OnDepartmentListDblClickFnc(CWnd *pWnd){
	((CMAMaterialFinalAccount*)pWnd)->OnDepartmentListDblClick();
} 
static void _OnDepartmentListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CMAMaterialFinalAccount*)pWnd)->OnDepartmentListSelectChange(nOldItem, nNewItem);
} 
static int _OnDepartmentListDeleteItemFnc(CWnd *pWnd){
	 return ((CMAMaterialFinalAccount*)pWnd)->OnDepartmentListDeleteItem();
} 
static long _OnTypeListLoadDataFnc(CWnd *pWnd){
	return ((CMAMaterialFinalAccount*)pWnd)->OnTypeListLoadData();
}
static void _OnTypeListDblClickFnc(CWnd *pWnd){
	((CMAMaterialFinalAccount*)pWnd)->OnTypeListDblClick();
}
static int _OnTypeListCheckAllFnc(CWnd *pWnd){
	((CMAMaterialFinalAccount*)pWnd)->OnTypeListCheckAll();
	return 0;
}
static int _OnTypeListUncheckAllFnc(CWnd *pWnd){
	((CMAMaterialFinalAccount*)pWnd)->OnTypeListUncheckAll();
	return 0;
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CMAMaterialFinalAccount *pVw = (CMAMaterialFinalAccount *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CMAMaterialFinalAccount *pVw = (CMAMaterialFinalAccount *)pWnd;
	pVw->OnExportSelect();
} 
CMAMaterialFinalAccount::CMAMaterialFinalAccount(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CMAMaterialFinalAccount::~CMAMaterialFinalAccount(){
}
void CMAMaterialFinalAccount::OnCreateComponents(){
	m_wndMaterialFinalAccount.Create(this, _T("Material Final Account"), 5, 5, 575, 545);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 390, 550, 490, 575);
	m_wndExport.Create(this, _T("&Export"), 495, 550, 575, 575);
	m_wndStorageList.Create(this,10, 60, 290, 295); 
	m_wndStorageList.SetCheckBox(true);
	m_wndTypeList.Create(this,295, 60, 570, 295); 
	m_wndTypeList.SetCheckBox(true);
	m_wndDepartmentList.Create(this,10, 300, 570, 540); 
	m_wndDepartmentList.SetCheckBox(true);
	
	
}
void CMAMaterialFinalAccount::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

	m_wndStorageList.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStorageList.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

	m_wndDepartmentList.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndDepartmentList.InsertColumn(1, _T("Name"), CFMT_TEXT, 320);

	m_wndTypeList.InsertColumn(0, _T("ID"), CFMT_TEXT, 60);
	m_wndTypeList.InsertColumn(1, _T("Name"), CFMT_TEXT, 190);
}

void CMAMaterialFinalAccount::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndStorageList.SetEvent(WE_SELCHANGE, _OnStorageListSelectChangeFnc);
	m_wndStorageList.SetEvent(WE_LOADDATA, _OnStorageListLoadDataFnc);
	m_wndStorageList.SetEvent(WE_DBLCLICK, _OnStorageListDblClickFnc);
	m_wndStorageList.AddEvent(1, _T("Check All"), _OnStorageListCheckAllFnc);
	m_wndStorageList.AddEvent(2, _T("UnCheck All"), _OnStorageListUncheckAllFnc);
	m_wndDepartmentList.SetEvent(WE_SELCHANGE, _OnDepartmentListSelectChangeFnc);
	m_wndDepartmentList.SetEvent(WE_LOADDATA, _OnDepartmentListLoadDataFnc);
	m_wndDepartmentList.SetEvent(WE_DBLCLICK, _OnDepartmentListDblClickFnc);
	m_wndDepartmentList.AddEvent(1, _T("Check All"), _OnDepartmentListCheckAllFnc);
	m_wndDepartmentList.AddEvent(2, _T("UnCheck All"), _OnDepartmentListUncheckAllFnc);
	m_wndTypeList.SetEvent(WE_LOADDATA, _OnTypeListLoadDataFnc);
	m_wndTypeList.SetEvent(WE_DBLCLICK, _OnTypeListDblClickFnc);
	m_wndTypeList.AddEvent(1, _T("Check All"), _OnTypeListCheckAllFnc);
	m_wndTypeList.AddEvent(2, _T("Uncheck All"), _OnTypeListUncheckAllFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");

	UpdateData(false);
	OnStorageListLoadData();
	OnDepartmentListLoadData();
	OnTypeListLoadData();

}
void CMAMaterialFinalAccount::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CMAMaterialFinalAccount::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();

}

int CMAMaterialFinalAccount::SetMode(int nMode){
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

/*void CMAMaterialFinalAccount::OnFromDateChange(){
	
} */
/*void CMAMaterialFinalAccount::OnFromDateSetfocus(){
	
} */
/*void CMAMaterialFinalAccount::OnFromDateKillfocus(){
	
} */
int CMAMaterialFinalAccount::OnFromDateCheckValue(){
	return 0;
}
 
/*void CMAMaterialFinalAccount::OnToDateChange(){
	
} */
/*void CMAMaterialFinalAccount::OnToDateSetfocus(){
	
} */
/*void CMAMaterialFinalAccount::OnToDateKillfocus(){
	
} */
int CMAMaterialFinalAccount::OnToDateCheckValue(){
	return 0;
}
 
void CMAMaterialFinalAccount::OnStorageListDblClick(){
	
} 
void CMAMaterialFinalAccount::OnStorageListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10* ) AfxGetMainWnd();
	
} 
int CMAMaterialFinalAccount::OnStorageListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10* ) AfxGetMainWnd();
	 return 0;
} 
long CMAMaterialFinalAccount::OnStorageListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10* ) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND msl_org_id = 'MA'"));
	return pMF->LoadStorageIntoListCtrl(&m_wndStorageList, szFilter);
}
void CMAMaterialFinalAccount::OnDepartmentListDblClick(){
	
} 
void CMAMaterialFinalAccount::OnDepartmentListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10* ) AfxGetMainWnd();
	
} 
int CMAMaterialFinalAccount::OnDepartmentListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10* ) AfxGetMainWnd();
	 return 0;
} 
long CMAMaterialFinalAccount::OnDepartmentListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndDepartmentList.BeginLoad(); 
	m_wndDepartmentList.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" select sd_id as id, sd_name as name from sys_dept where sd_id <> '15' order by  sd_name"));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDepartmentList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndDepartmentList.EndLoad(); 
	return nCount;
}

void CMAMaterialFinalAccount::OnTypeListDblClick(){
	
}
long CMAMaterialFinalAccount::OnTypeListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndTypeList.BeginLoad(); 
	m_wndTypeList.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT mpt_product_type_id as id, mpt_name as name ") \
		_T("FROM m_product_type ") \
		_T("WHERE mpt_isactive='Y' AND mpt_org_id in('GL', '%s') ORDER BY mpt_product_type_id "), pMF->GetModuleID());
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndTypeList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndTypeList.EndLoad(); 
	return nCount;
	
}
void CMAMaterialFinalAccount::OnPrintPreviewSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szPrintDate, szPos;
	int nIdx = 1;
	double nCost = 0;
	long double nTtlCost[14];
	CReport rpt;
	CReportSection *rptDetail = NULL;
	if (!rpt.Init(_T("Reports\\HMS\\MA_BANGTHANHQUYETTOANDUNGCUVTTH.RPT")))
	{
		return;
	}
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		if (nCount < 0)
			_msg(_T("%s"), szSQL);
		else
			_fmsg(_T("%s"), szSQL);
		return;
	}
	for (int i = 0; i < 14; i++)
		nTtlCost[i] = 0;
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		for (int i = 0; i < rs.GetFieldCount(); i++)
		{
			if (i > 0)
			{
				rs.GetValue(rs.GetFieldName(i), nCost);
				nTtlCost[i-1] += nCost;
			}
			rptDetail->SetValue(int2str(i+2), rs.GetValue(rs.GetFieldName(i)));
		}
		rs.MoveNext();
	}
	if (nTtlCost[13] > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
		for (int i = 0; i < 14;i++)
		{
			tmpStr.Format(_T("%f"), nTtlCost[i]);
			szPos.Format(_T("s%d"), i+3);
			rptDetail->SetValue(szPos, tmpStr);
		}
	}
	szPrintDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szPrintDate.Mid(8, 2), szPrintDate.Mid(5, 2), szPrintDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}
 
void CMAMaterialFinalAccount::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr, szNewType, szOldType;
	CRecord rs(&pMF->m_db);
	int nIdx = 1, j = 0;
	long double nTotal[14];
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONASTERISK | MB_OK);
		return;
	}
	for (int i = 0; i < 14; i++)
		nTotal[i] = 0;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	//Header Section
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 50);
	xls.SetColumnWidth(2, 15);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 15);
	xls.SetColumnWidth(7, 15);
	xls.SetColumnWidth(8, 15);
	xls.SetColumnWidth(9, 15);
	xls.SetColumnWidth(10, 15);
	xls.SetColumnWidth(11, 15);
	xls.SetColumnWidth(12, 15);
	xls.SetColumnWidth(13, 15);
	xls.SetColumnWidth(14, 15);
	xls.SetColumnWidth(15, 15);
	xls.SetColumnWidth(16, 15);
	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 4);
	xls.SetCellText(nCol, nRow + 3, _T("\x42\x1EA2NG TH\x41NH QUY\x1EBET TO\xC1N \x44\x1EE4NG \x43\x1EE4 V\xC0 V\x1EACT T\x1AF TI\xCAU H\x41O"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 4);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	CStringArray arrCol, arrDat;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn khoa"));
	arrCol.Add(_T("Qu\xE2n"));
	arrCol.Add(_T("BHYT Qu\xE2n"));
	arrCol.Add(_T("\x43h\xEDnh s\xE1\x63h"));
	arrCol.Add(_T("BHYT kh\xE1\x63"));
	arrCol.Add(_T("BHYT Th\xE2n nh\xE2n s\x1EF9 qu\x61n"));
	arrCol.Add(_T("Tr\x1EBB \x65m \x64\x1B0\x1EDBi \x36 tu\x1ED5i"));
	arrCol.Add(_T("\x44\x1ECB\x63h v\x1EE5"));
	arrCol.Add(_T("\x42\x1EA1n"));
	arrCol.Add(_T("Tr\x1EBB \x65m \x64\x1B0\x1EDBi \x36 tu\x1ED5i(\x63\xF3 th\x1EBB)"));
	arrCol.Add(_T("BHYT Qu\xE2n \x111\x1ED9i"));
	arrCol.Add(_T("BHYT Qu\xE2n nh\xE2n"));
	arrCol.Add(_T("BHYT(N\x1EE3 th\x1EBB)"));
	arrCol.Add(_T("Kh\xE1\x63"));
	arrCol.Add(_T("T\x1ED5ng kinh ph\xED \x111\xE3 s\x1EED \x64\x1EE5ng"));
	for (int i = 0; i < arrCol.GetCount(); i++)
		xls.SetCellText(nCol + i, nRow + 5, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);
	nRow = 6;
	arrDat.Add(_T("solamt"));
	arrDat.Add(_T("solinsamt"));
	arrDat.Add(_T("polamt"));
	arrDat.Add(_T("otherinsamt"));
	arrDat.Add(_T("relativeinsamt"));
	arrDat.Add(_T("kidamt"));
	arrDat.Add(_T("servamt"));
	arrDat.Add(_T("palamt"));
	arrDat.Add(_T("kidinsamt"));
	arrDat.Add(_T("armyamt"));
	arrDat.Add(_T("officeramt"));
	arrDat.Add(_T("carddebtamt"));
	arrDat.Add(_T("otheramt"));
	arrDat.Add(_T("totalamt"));
	while (!rs.IsEOF())
	{
		xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, rs.GetValue(_T("deptname")), FMT_TEXT);
		for (int i = 0; i < arrDat.GetCount(); i++)
		{
			xls.SetCellText(nCol+i+2, nRow, rs.GetValue(arrDat.GetAt(i)), FMT_NUMBER1);
			nTotal[i] += str2double(rs.GetValue(arrDat.GetAt(i)));
		}
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
	for (int i = 0; i < 14; i ++)
	{
		tmpStr.Format(_T("%f"), nTotal[i]);
		xls.SetCellText(nCol + i + 2, nRow, tmpStr, FMT_NUMBER1, true);
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\BANG THANH QUYET TOAN DUNG CU VA VAT TU TIEU HAO.xls"));
}
 
CString CMAMaterialFinalAccount::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szTransactionWhere, szPharmaWhere, szTransactionWhere1;
	CString szStorageStr, szDepts, szTypes;
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
	for (int i = 0; i < m_wndDepartmentList.GetItemCount(); i++)
	{
		if (m_wndDepartmentList.GetCheck(i))
		{
			m_wndDepartmentList.SetCurSel(i);
			if (!szDepts.IsEmpty())
				szDepts += _T(",");
			szDepts.AppendFormat(_T("'%s'"), m_wndDepartmentList.GetItemText(i, 0));
		}
	}
	for (int i = 0; i < m_wndTypeList.GetItemCount(); i++)
	{
		if (m_wndTypeList.GetCheck(i))
		{
			m_wndTypeList.SetCurSel(i);
			if (!szTypes.IsEmpty())
				szTypes += _T(",");
			szTypes.AppendFormat(_T("'%s'"), m_wndTypeList.GetItemText(i, 0));
		}
	}
	szTransactionWhere.Format(_T(" AND mt_status = 'A' AND product_org_id = '%s'") \
							  _T(" AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), 
							  pMF->GetModuleID(), m_szFromDate, m_szToDate);
	szTransactionWhere1.Format(_T(" AND mt_status = 'A' AND product_org_id = '%s'") \
							  _T(" AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), 
							  pMF->GetModuleID(), m_szFromDate, m_szToDate);
	szPharmaWhere.Format(_T(" AND hpo_orderstatus = 'A' AND product_org_id = '%s'") \
							  _T(" AND hpo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), 
							  pMF->GetModuleID(), m_szFromDate, m_szToDate);
	
	if (!szStorageStr.IsEmpty())
	{
		szTransactionWhere.AppendFormat(_T(" AND mt_storage_id IN (%s)"), szStorageStr);
		szTransactionWhere1.AppendFormat(_T(" AND mt_storage_to_id IN (%s)"), szStorageStr);
		szPharmaWhere.AppendFormat(_T(" AND hpo_storage_id IN (%s)"), szStorageStr);
	}	
	if (!szDepts.IsEmpty())
	{
		szTransactionWhere.AppendFormat(_T(" and mt_department_to_id in(%s) "), szDepts);
		szTransactionWhere1.AppendFormat(_T(" and mt_department_to_id in(%s) "), szDepts);
		szPharmaWhere.AppendFormat(_T(" and hpo_deptid in(%s) "), szDepts);
	}
	if (!szTypes.IsEmpty())
	{
		szTransactionWhere.AppendFormat(_T(" AND product_producttype IN (%s)"), szTypes);
		szTransactionWhere1.AppendFormat(_T(" and product_producttype in(%s) "), szDepts);
		szPharmaWhere.AppendFormat(_T(" AND product_producttype IN (%s)"), szTypes);
	}
	szSQL.Format(_T(" SELECT hpo_deptid deptname,") \
				_T("        SUM(solamt)         solamt, ") \
				_T("        SUM(solinsamt)      solinsamt, ") \
				_T("        SUM(polamt)         polamt, ") \
				_T("        SUM(otherinsamt)    otherinsamt, ") \
				_T("        SUM(relativeinsamt) relativeinsamt, ") \
				_T("        SUM(kidamt)         kidamt, ") \
				_T("        SUM(servamt)        servamt, ") \
				_T("        SUM(palamt)         palamt, ") \
				_T("        SUM(kidinsamt)      kidinsamt, ") \
				_T("        SUM(armyamt)        armyamt, ") \
				_T("        SUM(officeramt)     officeramt, ") \
				_T("        SUM(carddebtamt)    carddebtamt, ") \
				_T("        SUM(otheramt)       otheramt, ") \
				_T("        SUM(solamt + solinsamt + polamt + otherinsamt ") \
				_T("            + relativeinsamt + kidamt + servamt + palamt ") \
				_T("            + kidinsamt + armyamt + officeramt + carddebtamt ") \
				_T("            + otheramt)     totalamt ") \
				_T(" FROM   (SELECT hpo_deptid, ") \
				_T("                CASE WHEN hpo_object_id = 1 THEN hpol_unitprice * hpol_qtyissue ") \
				_T("                ELSE 0 ") \
				_T("                END AS solamt, ") \
				_T("                CASE WHEN hpo_object_id = 2 THEN hpol_unitprice * hpol_qtyissue ") \
				_T("                ELSE 0 ") \
				_T("                END AS solinsamt, ") \
				_T("                CASE WHEN hpo_object_id = 3 THEN hpol_unitprice * hpol_qtyissue ") \
				_T("                ELSE 0 ") \
				_T("                END AS polamt, ") \
				_T("                CASE WHEN hpo_object_id = 4 THEN hpol_unitprice * hpol_qtyissue ") \
				_T("                ELSE 0 ") \
				_T("                END AS otherinsamt, ") \
				_T("                CASE WHEN hpo_object_id = 5 THEN hpol_unitprice * hpol_qtyissue ") \
				_T("                ELSE 0 ") \
				_T("                END AS relativeinsamt, ") \
				_T("                CASE WHEN hpo_object_id = 6 THEN hpol_unitprice * hpol_qtyissue ") \
				_T("                ELSE 0 ") \
				_T("                END AS kidamt, ") \
				_T("                CASE WHEN hpo_object_id = 7 THEN hpol_unitprice * hpol_qtyissue ") \
				_T("                ELSE 0 ") \
				_T("                END AS servamt, ") \
				_T("                CASE WHEN hpo_object_id = 8 THEN hpol_unitprice * hpol_qtyissue ") \
				_T("                ELSE 0 ") \
				_T("                END AS palamt, ") \
				_T("                CASE WHEN hpo_object_id = 9 THEN hpol_unitprice * hpol_qtyissue ") \
				_T("                ELSE 0 ") \
				_T("                END AS kidinsamt, ") \
				_T("                CASE WHEN hpo_object_id = 10 THEN hpol_unitprice * hpol_qtyissue ") \
				_T("                ELSE 0 ") \
				_T("                END AS armyamt, ") \
				_T("                CASE WHEN hpo_object_id = 11 THEN hpol_unitprice * hpol_qtyissue ") \
				_T("                ELSE 0 ") \
				_T("                END AS officeramt, ") \
				_T("                CASE WHEN hpo_object_id = 12 THEN hpol_unitprice * hpol_qtyissue ") \
				_T("                ELSE 0 ") \
				_T("                END AS carddebtamt, ") \
				_T("                0   AS otheramt ") \
				_T("         FROM   (SELECT    hpo_deptid, ") \
				_T("                           hpo_object_id, ") \
				_T("                           hpol_unitprice, ") \
				_T("                           hpol_qtyissue ") \
				_T("                 FROM      hms_ipharmaorder ") \
				_T("                 LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T("                 LEFT JOIN m_product_view ON ( product_id = hpol_product_id ) ") \
				_T("                 WHERE     hpo_ordertype IN ( 'D', 'M', 'B' ) ") \
				_T("				 %s") \
				_T("                 UNION ALL ") \
				_T("                 SELECT    hpo_deptid, ") \
				_T("                           hpo_object_id, ") \
				_T("                           hpol_unitprice, ") \
				_T("                           hpol_qtyissue ") \
				_T("                 FROM      m_transaction ") \
				_T("                 LEFT JOIN purchase_orderline2 ON ( mt_transaction_id = pol_transaction_id ) ") \
				_T("                 LEFT JOIN hms_ipharmaorder ON ( hpo_orderid = pol_orderid ) ") \
				_T("                 LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid AND pol_product_id = hpol_product_id) ") \
				_T("                 LEFT JOIN m_product_view ON ( product_id = hpol_product_id ) ") \
				_T("                 WHERE     mt_doctype = 'CON' ") \
				_T("				 %s) tbl1 ") \
				_T("         UNION ALL ") \
				_T("         SELECT    mt_department_to_id, ") \
				_T("                   mtl_saleprice * mtl_qtysold     solamt, ") \
				_T("                   mtl_saleprice * mtl_qtysoldins  solinsamt, ") \
				_T("                   mtl_saleprice * mtl_qtypolicy   polamt, ") \
				_T("                   mtl_saleprice * mtl_qtyotherins otherinsamt, ") \
				_T("                   0                               relativeinsamt, ") \
				_T("                   0                               kidamt, ") \
				_T("                   mtl_saleprice * mtl_qtyservice  servamt, ") \
				_T("                   0                               palamt, ") \
				_T("                   0                               kidinsamt, ") \
				_T("                   0                               armyamt, ") \
				_T("                   0                               officeramt, ") \
				_T("                   0                               carddebtamt, ") \
				_T("                   mtl_saleprice * mtl_qtyother    otheramt ") \
				_T("         FROM      m_transaction ") \
				_T("         LEFT JOIN m_transactionline ON ( mt_transaction_id = mtl_transaction_id ) ") \
				_T("         LEFT JOIN m_product_view ON ( product_id = mtl_product_id ) ") \
				_T("         WHERE     (mt_doctype = 'DDO' ") \
				_T("		 OR (mt_doctype = 'CSO' AND NVL(mtl_client_id, 'XX') <> 'TT'))") \
				_T("		 %s") \
				_T("         UNION ALL ") \
				_T("         SELECT    mt_department_to_id, ") \
				_T("                   mtl_saleprice * mtl_qtysold     solamt, ") \
				_T("                   mtl_saleprice * mtl_qtysoldins  solinsamt, ") \
				_T("                   mtl_saleprice * mtl_qtypolicy   polamt, ") \
				_T("                   mtl_saleprice * mtl_qtyotherins otherinsamt, ") \
				_T("                   0                               relativeinsamt, ") \
				_T("                   0                               kidamt, ") \
				_T("                   mtl_saleprice * mtl_qtyservice  servamt, ") \
				_T("                   0                               palamt, ") \
				_T("                   0                               kidinsamt, ") \
				_T("                   0                               armyamt, ") \
				_T("                   0                               officeramt, ") \
				_T("                   0                               carddebtamt, ") \
				_T("                   mtl_saleprice * mtl_qtyother    otheramt ") \
				_T("         FROM      m_transaction ") \
				_T("         LEFT JOIN m_transactionline_c ON ( mt_transaction_id = mtl_transaction_id ) ") \
				_T("         LEFT JOIN m_product_view ON ( product_id = mtl_product_id ) ") \
				_T("         WHERE     mt_doctype = 'CSO' AND NVL(mtl_client_id, 'XXX') <> 'REP'") \
				_T("		 %s ") \
				_T("		) tbl2 ") \
				_T(" GROUP  BY hpo_deptid ORDER BY hpo_deptid"), szPharmaWhere, szTransactionWhere, szTransactionWhere, szTransactionWhere1);
	return szSQL;
}

void CMAMaterialFinalAccount::OnStorageListCheckAll()
{
	for (int i = 0; i < m_wndStorageList.GetItemCount(); i++)
	{
		if (!m_wndStorageList.GetCheck(i))
		{
			m_wndStorageList.SetCheck(i, TRUE);
		}
	}
	return;
}

void CMAMaterialFinalAccount::OnStorageListUncheckAll()
{
	for (int i = 0; i < m_wndStorageList.GetItemCount(); i++)
	{
		if (m_wndStorageList.GetCheck(i))
		{
			m_wndStorageList.SetCheck(i, FALSE);
		}
	}
	return;
}

void CMAMaterialFinalAccount::OnDepartmentListCheckAll()
{
	for (int i = 0; i < m_wndDepartmentList.GetItemCount(); i++)
	{
		if (!m_wndDepartmentList.GetCheck(i))
		{
			m_wndDepartmentList.SetCheck(i, TRUE);
		}
	}
	return;
}

void CMAMaterialFinalAccount::OnDepartmentListUncheckAll()
{
	for (int i = 0; i < m_wndDepartmentList.GetItemCount(); i++)
	{
		if (m_wndDepartmentList.GetCheck(i))
		{
			m_wndDepartmentList.SetCheck(i, FALSE);
		}
	}
	return;
}

void CMAMaterialFinalAccount::OnTypeListCheckAll()
{
	for (int i = 0; i < m_wndTypeList.GetItemCount(); i++)
	{
		if (!m_wndTypeList.GetCheck(i))
		{
			m_wndTypeList.SetCheck(i, TRUE);
		}
	}
	return;
}

void CMAMaterialFinalAccount::OnTypeListUncheckAll()
{
	for (int i = 0; i < m_wndTypeList.GetItemCount(); i++)
	{
		if (m_wndTypeList.GetCheck(i))
		{
			m_wndTypeList.SetCheck(i, FALSE);
		}
	}
	return;
}