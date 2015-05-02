#include "stdafx.h"
#include "MAMaterialUsagebyDrug.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "Excel.h"
#include "StringToken.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CMAMaterialUsagebyDrug *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CMAMaterialUsagebyDrug *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CMAMaterialUsagebyDrug *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CMAMaterialUsagebyDrug *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CMAMaterialUsagebyDrug *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CMAMaterialUsagebyDrug *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CMAMaterialUsagebyDrug *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CMAMaterialUsagebyDrug *)pWnd)->OnToDateCheckValue();
} 
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CMAMaterialUsagebyDrug*)pWnd)->OnStorageLoadData();
} 
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CMAMaterialUsagebyDrug*)pWnd)->OnObjectLoadData();
}
static int _OnDepartmentListCheckAllFnc(CWnd *pWnd){
	 ((CMAMaterialUsagebyDrug*)pWnd)->OnDepartmentListCheckAll();
	 return 0;
} 
static int _OnDepartmentListUncheckAllFnc(CWnd *pWnd){
	((CMAMaterialUsagebyDrug*)pWnd)->OnDepartmentListUncheckAll();
	return 0;
} 
static long _OnDepartmentListLoadDataFnc(CWnd *pWnd){
	return ((CMAMaterialUsagebyDrug*)pWnd)->OnDepartmentListLoadData();
} 
static void _OnDepartmentListDblClickFnc(CWnd *pWnd){
	((CMAMaterialUsagebyDrug*)pWnd)->OnDepartmentListDblClick();
} 
static void _OnDepartmentListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CMAMaterialUsagebyDrug*)pWnd)->OnDepartmentListSelectChange(nOldItem, nNewItem);
} 
static int _OnDepartmentListDeleteItemFnc(CWnd *pWnd){
	 return ((CMAMaterialUsagebyDrug*)pWnd)->OnDepartmentListDeleteItem();
} 
static long _OnTypeListLoadDataFnc(CWnd *pWnd){
	return ((CMAMaterialUsagebyDrug*)pWnd)->OnTypeListLoadData();
}
static void _OnTypeListDblClickFnc(CWnd *pWnd){
	((CMAMaterialUsagebyDrug*)pWnd)->OnTypeListDblClick();
}
static int _OnTypeListCheckAllFnc(CWnd *pWnd){
	((CMAMaterialUsagebyDrug*)pWnd)->OnTypeListCheckAll();
	return 0;
}
static int _OnTypeListUncheckAllFnc(CWnd *pWnd){
	((CMAMaterialUsagebyDrug*)pWnd)->OnTypeListUncheckAll();
	return 0;
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CMAMaterialUsagebyDrug *pVw = (CMAMaterialUsagebyDrug *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CMAMaterialUsagebyDrug *pVw = (CMAMaterialUsagebyDrug *)pWnd;
	pVw->OnExportSelect();
} 
CMAMaterialUsagebyDrug::CMAMaterialUsagebyDrug(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CMAMaterialUsagebyDrug::~CMAMaterialUsagebyDrug(){
}
void CMAMaterialUsagebyDrug::OnCreateComponents(){
	m_wndMaterialFinalAccount.Create(this, _T("Material Usage by Drug"), 5, 5, 575, 545);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndStorageLabel.Create(this, _T("Storage"), 10, 60, 90, 85);
	m_wndStorage.SetCheckBox(true);
	m_wndStorage.Create(this,95, 60, 290, 85); 
	m_wndObjectLabel.Create(this, _T("Object"), 295, 60, 375, 85);
	m_wndObject.SetCheckBox(true);
	m_wndObject.Create(this,381, 60, 570, 85); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 390, 550, 490, 575);
	m_wndExport.Create(this, _T("&Export"), 495, 550, 575, 575);
	m_wndDepartmentList.Create(this,10, 90, 570, 330); 
	m_wndDepartmentList.SetCheckBox(true);
	m_wndTypeList.Create(this,10, 335, 570, 540); 
	m_wndTypeList.SetCheckBox(true);
}
void CMAMaterialUsagebyDrug::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

	m_wndTypeList.InsertColumn(0, _T("ID"), CFMT_TEXT, 85);
	m_wndTypeList.InsertColumn(1, _T("Name"), CFMT_TEXT, 380);

	m_wndStorage.InsertColumn(0, _T("ID"), CFMT_NUMBER, 40);
	m_wndStorage.InsertColumn(1, _T("Name"), CFMT_TEXT, 120);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 40);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 120);
	m_wndObject.Format(2, 1, 12);

	m_wndDepartmentList.InsertColumn(0, _T("ID"), CFMT_TEXT, 85);
	m_wndDepartmentList.InsertColumn(1, _T("Name"), CFMT_TEXT, 380);

}

void CMAMaterialUsagebyDrug::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndStorage.SetEvent(WE_LOADDATA, _OnStorageLoadDataFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
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
	OnObjectLoadData();
	OnStorageLoadData();
	OnTypeListLoadData();
	OnDepartmentListLoadData();

}
void CMAMaterialUsagebyDrug::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStorage.GetDlgCtrlID(), m_szStorageKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
}
void CMAMaterialUsagebyDrug::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();

}

int CMAMaterialUsagebyDrug::SetMode(int nMode){
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

/*void CMAMaterialUsagebyDrug::OnFromDateChange(){
	
} */
/*void CMAMaterialUsagebyDrug::OnFromDateSetfocus(){
	
} */
/*void CMAMaterialUsagebyDrug::OnFromDateKillfocus(){
	
} */
int CMAMaterialUsagebyDrug::OnFromDateCheckValue(){
	return 0;
}
 
/*void CMAMaterialUsagebyDrug::OnToDateChange(){
	
} */
/*void CMAMaterialUsagebyDrug::OnToDateSetfocus(){
	
} */
/*void CMAMaterialUsagebyDrug::OnToDateKillfocus(){
	
} */
int CMAMaterialUsagebyDrug::OnToDateCheckValue(){
	return 0;
}

void CMAMaterialUsagebyDrug::OnStorageSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10* ) AfxGetMainWnd();
	
} 
int CMAMaterialUsagebyDrug::OnStorageListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10* ) AfxGetMainWnd();
	 return 0;
} 
long CMAMaterialUsagebyDrug::OnStorageLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10* ) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND msl_org_id = 'MA'"));
	return pMF->LoadStorageList(&m_wndStorage, m_szStorageKey, szFilter);
}
long CMAMaterialUsagebyDrug::OnObjectLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT ho_id id, ho_desc descr FROM hms_object ORDER BY cast(ho_id as integer)"));
	int nCount = rs.ExecSQL(szSQL);
	m_wndObject.DeleteAllItems();
	while (!rs.IsEOF())
	{
		m_wndObject.AddItems(
			rs.GetValue(_T("id")),
			rs.GetValue(_T("descr")),
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
void CMAMaterialUsagebyDrug::OnDepartmentListDblClick(){
	
} 
void CMAMaterialUsagebyDrug::OnDepartmentListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10* ) AfxGetMainWnd();
	
} 
int CMAMaterialUsagebyDrug::OnDepartmentListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10* ) AfxGetMainWnd();
	 return 0;
} 
long CMAMaterialUsagebyDrug::OnDepartmentListLoadData(){
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

long CMAMaterialUsagebyDrug::OnTypeListLoadData(){
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

void CMAMaterialUsagebyDrug::OnTypeListDblClick(){
}
void CMAMaterialUsagebyDrug::OnPrintPreviewSelect(){
	CMainFrame_E10* pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	if (!rpt.Init(_T("Reports\\HMS\\MA_BAOCAOSUDUNGVATTUYTE.RPT")))
		return;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szSysDate;
	CReportSection *rptDetail = NULL, *rptHeader = NULL;
	int nIdx = 1;
	double nCost = 0;
	long double nTtlCost = 0;
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
	rptHeader = rpt.GetReportHeader();
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	rptHeader->SetValue(_T("Stockname"), m_szStorageStr);
	rptHeader->SetValue(_T("Object"), m_szObjectStr);
	rptHeader->SetValue(_T("DrugType"), m_szTypeStr);
	rptHeader->SetValue(_T("Department"), m_szDeptStr);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("product_name")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("uom")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("unitprice")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("qtyissue")));
		rs.GetValue(_T("amount"), nCost);
		nTtlCost += nCost;
		rptDetail->SetValue(_T("6"), double2str(nCost));
		rs.MoveNext();
	}
	tmpStr.Format(_T("%f"), nTtlCost);
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	rptDetail->SetValue(_T("s6"), tmpStr);
	szSysDate = pMF->GetSysDateTime();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Mid(11, 5),szSysDate.Mid(8,2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}
 
void CMAMaterialUsagebyDrug::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr, szNewType, szOldType;
	CRecord rs(&pMF->m_db);
	int nIdx = 1, j = 0;
	long double nTotal = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONASTERISK | MB_OK);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	//Header Section
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 50);
	xls.SetColumnWidth(10, 15);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 15);
	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 4);
	xls.SetCellText(nCol, nRow + 3, _T("T\xCCNH H\xCCNH S\x1EEC \x44\x1EE4NG V\x1EACT T\x1AF Y T\x1EBE"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 4);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn v\x1EADt t\x1B0"));
	arrCol.Add(_T("\x110VT"));
	arrCol.Add(_T("\x110\x1A1n gi\xE1"));
	arrCol.Add(_T("S\x1ED1 l\x1B0\x1EE3ng"));
	arrCol.Add(_T("Th\xE0nh ti\x1EC1n"));
	for (int i = 0; i < arrCol.GetCount(); i++)
		xls.SetCellText(nCol + i, nRow + 5, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);
	nRow = 6;
	while (!rs.IsEOF())
	{
		xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, rs.GetValue(_T("product_name")), FMT_TEXT);
		xls.SetCellText(nCol+2, nRow, rs.GetValue(_T("uom")), FMT_TEXT);
		xls.SetCellText(nCol+3, nRow, rs.GetValue(_T("unitprice")), FMT_NUMBER1);
		xls.SetCellText(nCol+4, nRow, rs.GetValue(_T("qtyissue")), FMT_NUMBER1);
		xls.SetCellText(nCol+5, nRow, rs.GetValue(_T("amount")), FMT_NUMBER1);
		nTotal += str2double(rs.GetValue(_T("amount")));
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellText(nCol + 4, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
	tmpStr.Format(_T("%.2f"), nTotal);
	xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1, true);
	EndWaitCursor();
	xls.Save(_T("Exports\\TINH HINH SU DUNG VAT TU Y TE.xls"));
}
 
CString CMAMaterialUsagebyDrug::GetQueryString(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szWhere, tmpStr, szQtyExpr, szObjectWhere, szObjectID, szObjects;
	CString szStorageStr, szDepts, szTypes;
	int nObjectID = 0;
	m_szObjectStr.Empty();
	m_szTypeStr.Empty();
	m_szDeptStr.Empty();
	m_szStorageStr.Empty();
	CStringArray arrObj;
	arrObj.SetSize(13);
	arrObj.SetAt(0, _T("mtl_qtyother"));
	arrObj.SetAt(1, _T("mtl_qtysold"));
	arrObj.SetAt(2, _T("mtl_qtysoldins"));
	arrObj.SetAt(3, _T("mtl_qtypolicy"));
	arrObj.SetAt(4, _T("mtl_qtyotherins"));
	arrObj.SetAt(5, _T("mtl_qtyother"));
	arrObj.SetAt(6, _T("mtl_qtyother"));
	arrObj.SetAt(7, _T("mtl_qtyservice"));
	arrObj.SetAt(8, _T("mtl_qtyother"));
	arrObj.SetAt(9, _T("mtl_qtyother"));
	arrObj.SetAt(10, _T("mtl_qtyother"));
	arrObj.SetAt(11, _T("mtl_qtyother"));
	arrObj.SetAt(12, _T("mtl_qtyother"));
	for (int i = 0; i < m_wndStorage.GetItemCount(); i++)
	{
		if (m_wndStorage.GetCheck(i))
		{
			m_wndStorage.SetCurSel(i);
			if (!szStorageStr.IsEmpty())
			{
				szStorageStr += _T(", ");
				m_szStorageStr += _T(", ");
			}
			szStorageStr.AppendFormat(_T("'%s'"), m_wndStorage.GetCurrent(0));
			m_szStorageStr.AppendFormat(_T("%s"), m_wndStorage.GetCurrent(1));
		}
	}
	if (m_szStorageStr.IsEmpty())
		m_szStorageStr = _T("To\xE0n \x62\x1ED9");
	CStringToken StorageToken(m_szStorageStr);
	if (StorageToken.GetSize() > 3)
		m_szStorageStr = _T("Nhi\x1EC1u kho");
	for (int i = 0; i < m_wndObject.GetItemCount(); i++)
	{
		if (m_wndObject.GetCheck(i))
		{
			 m_wndObject.SetCurSel(i);
			if (!szObjects.IsEmpty())
			{
				szObjects += _T(", ");
				m_szObjectStr += _T(", ");
			}
			szObjects.AppendFormat(_T("'%s'"), m_wndObject.GetCurrent(0));
			m_szObjectStr.AppendFormat(_T("%s"), m_wndObject.GetCurrent(1));
		}
	}
	if (m_szObjectStr.IsEmpty())
		m_szObjectStr = _T("To\xE0n \x62\x1ED9");
	CStringToken ObjectToken(m_szObjectStr);
	if (ObjectToken.GetSize() > 3)
		m_szObjectStr = _T("Nhi\x1EC1u \x111\x1ED1i t\x1B0\x1EE3ng");
	for (int i = 0; i < m_wndDepartmentList.GetItemCount(); i++)
	{
		if (m_wndDepartmentList.GetCheck(i))
		{
			if (!szDepts.IsEmpty())
			{
				szDepts += _T(",");
				m_szDeptStr += _T(", ");
			}
			szDepts.AppendFormat(_T("'%s'"), m_wndDepartmentList.GetItemText(i, 0));
			m_szDeptStr.AppendFormat(_T("%s"), m_wndDepartmentList.GetItemText(i, 1));
		}
	}
	if (m_szDeptStr.IsEmpty())
		m_szDeptStr = _T("To\xE0n \x62\x1ED9");
	CStringToken DeptToken(m_szDeptStr);
	if (DeptToken.GetSize() > 3)
		m_szDeptStr = _T("Nhi\x1EC1u kho\x61");
	for (int i = 0; i < m_wndTypeList.GetItemCount(); i++)
	{
		if (m_wndTypeList.GetCheck(i))
		{
			if (!szTypes.IsEmpty())
			{
				szTypes += _T(",");
				m_szTypeStr += _T(", ");
			}
			szTypes.AppendFormat(_T("'%s'"), m_wndTypeList.GetItemText(i, 0));
			m_szTypeStr.AppendFormat(_T("%s"), m_wndTypeList.GetItemText(i, 1));
		}
	}
	if (m_szTypeStr.IsEmpty())
		m_szTypeStr = _T("To\xE0n \x62\x1ED9");
	CStringToken TypeToken(m_szTypeStr);
	if (TypeToken.GetSize() > 3);
		m_szTypeStr = _T("Nhi\x1EC1u lo\x1EA1i");
	szWhere.Format(_T(" AND status = 'A' AND approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!szStorageStr.IsEmpty())
		szWhere.AppendFormat(_T(" AND storageid IN (%s)"), szStorageStr);
	if (!szDepts.IsEmpty())
		szWhere.AppendFormat(_T(" and deptid in (%s) "), szDepts);
	if (!szTypes.IsEmpty())
		szWhere.AppendFormat(_T(" AND product_producttype IN (%s)"), szTypes);
	szQtyExpr = _T("0");

	if (!szObjects.IsEmpty())
	{
		szObjectWhere.Format(_T(" AND hpo_object_id IN (%s)"), szObjects);
		CStringToken token(szObjects);
		for (int i = 0; i < token.GetSize(); i++)
		{
			token.GetAt(i, szObjectID);
			szObjectID.Remove('\'');
			nObjectID = str2int(szObjectID);
			tmpStr = arrObj.GetAt(nObjectID);
			if (tmpStr != _T("mtl_qtyother"))
				szQtyExpr.AppendFormat(_T("+ %s"), arrObj.GetAt(nObjectID));
			else
			{
				if (szQtyExpr.Find(_T("mtl_qtyother")) == -1)
					szQtyExpr.AppendFormat(_T("+ %s"), arrObj.GetAt(nObjectID));
			}
		}
	}
	else
		szQtyExpr = _T("mtl_qtyorder");
	szSQL.Format(_T(" SELECT    product_name, ") \
				_T("			get_uomname(product_uomid) uom,") \
				_T("           unitprice, ") \
				_T("           SUM(qtyissue)             qtyissue, ") \
				_T("           SUM(unitprice * qtyissue) amount ") \
				_T(" FROM      (SELECT    hpol_product_id productid, ") \
				_T("                      hpol_unitprice unitprice, ") \
				_T("                      hpol_qtyissue qtyissue, ") \
				_T("                      hpo_storage_id storageid, ") \
				_T("					  hpo_orderstatus status,") \
				_T("					  hpo_approvedate approveddate,") \
				_T("                      hpo_deptid deptid") \
				_T("            FROM      hms_ipharmaorder ") \
				_T("            LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T("            WHERE     hpo_ordertype IN ( 'D', 'M', 'B' ) ") \
				_T("			%s") \
				_T("            UNION ALL ") \
				_T("            SELECT    mtl_product_id, ") \
				_T("                      mtl_saleprice, ") \
				_T("                      %s, ") \
				_T("                      mt_storage_id, ") \
				_T("					  mt_status status,") \
				_T("					  mt_approveddate approveddate,") \
				_T("                      mt_department_to_id ") \
				_T("            FROM      m_transaction ") \
				_T("            LEFT JOIN m_transactionline ON ( mt_transaction_id = mtl_transaction_id ) ") \
				_T("            WHERE     mt_doctype = 'DDO' ") \
				_T("			OR (mt_doctype = 'CSO' AND NVL(mtl_client_id, 'XX') <> 'TT')") \
				_T("            UNION ALL ") \
				_T("            SELECT    mtl_product_id, ") \
				_T("                      mtl_saleprice, ") \
				_T("                      %s, ") \
				_T("                      mt_storage_to_id, ") \
				_T("					  mt_status status,") \
				_T("					  mt_approveddate approveddate,") \
				_T("                      mt_department_to_id ") \
				_T("            FROM      m_transaction ") \
				_T("            LEFT JOIN m_transactionline_c ON ( mt_transaction_id = mtl_transaction_id ) ") \
				_T("            WHERE     mt_doctype = 'CSO' AND NVL(mtl_client_id , 'XXX') <> 'REP'") \
				_T("            UNION ALL ") \
				_T("            SELECT    hpol_product_id, ") \
				_T("                      hpol_unitprice, ") \
				_T("                      hpol_qtyissue, ") \
				_T("                      mt_storage_id, ") \
				_T("					  mt_status status,") \
				_T("					  mt_approveddate approveddate,") \
				_T("                      hpo_deptid ") \
				_T("            FROM      m_transaction ") \
				_T("            LEFT JOIN purchase_orderline2 ON (pol_transaction_id = mt_transaction_id) ") \
				_T("            LEFT JOIN hms_ipharmaorder ON ( hpo_orderid = pol_orderid ) ") \
				_T("            LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ") \
				_T("                                               AND pol_product_id = hpol_product_id ) ") \
				_T("            WHERE     mt_doctype = 'CON' ") \
				_T("			%s)") \
				_T(" LEFT JOIN m_product_view ON ( productid = product_id ) ") \
				_T(" WHERE     product_org_id = 'MA' ") \
				_T(" %s") \
				_T(" GROUP     BY product_name, ") \
				_T("              unitprice, product_uomid ") \
				_T(" ORDER     BY product_name "), szObjectWhere, szQtyExpr, szQtyExpr, szObjectWhere, szWhere);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}

void CMAMaterialUsagebyDrug::OnDepartmentListCheckAll()
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

void CMAMaterialUsagebyDrug::OnDepartmentListUncheckAll()
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

void CMAMaterialUsagebyDrug::OnTypeListCheckAll()
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

void CMAMaterialUsagebyDrug::OnTypeListUncheckAll()
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