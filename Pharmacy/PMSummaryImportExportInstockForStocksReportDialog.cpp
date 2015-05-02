#include "stdafx.h"
#include "PMSummaryImportExportInstockForStocksReportDialog.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include ".\pmsummaryimportexportinstockforstocksreportdialog.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnToDateCheckValue();
} 
static void _OnTypeSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSSummaryImportExportInstockForStocksReportDialog* )pWnd)->OnTypeSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnTypeSelendokFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnTypeSelendok();
}
/*static void _OnTypeSetfocusFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnTypeKillfocus();
}*/
/*static void _OnTypeKillfocusFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnTypeKillfocus();
}*/
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnTypeLoadData();
}
/*static void _OnTypeAddNewFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnTypeAddNew();
}*/
static void _OnOriginalSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSSummaryImportExportInstockForStocksReportDialog* )pWnd)->OnOriginalSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnOriginalSelendokFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnOriginalSelendok();
}
/*static void _OnOriginalSetfocusFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnOriginalKillfocus();
}*/
/*static void _OnOriginalKillfocusFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnOriginalKillfocus();
}*/
static long _OnOriginalLoadDataFnc(CWnd *pWnd){
	return ((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnOriginalLoadData();
}
/*static void _OnOriginalAddNewFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd)->OnOriginalAddNew();
}*/
static long _OnStockListLoadDataFnc(CWnd *pWnd){
	return ((CPMSSummaryImportExportInstockForStocksReportDialog*)pWnd)->OnStockListLoadData();
} 
static void _OnStockListDblClickFnc(CWnd *pWnd){
	((CPMSSummaryImportExportInstockForStocksReportDialog*)pWnd)->OnStockListDblClick();
} 
static void _OnStockListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPMSSummaryImportExportInstockForStocksReportDialog*)pWnd)->OnStockListSelectChange(nOldItem, nNewItem);
} 
static int _OnStockListDeleteItemFnc(CWnd *pWnd){
	 return ((CPMSSummaryImportExportInstockForStocksReportDialog*)pWnd)->OnStockListDeleteItem();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CPMSSummaryImportExportInstockForStocksReportDialog *pVw = (CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportXLSSelectFnc(CWnd *pWnd){
	CPMSSummaryImportExportInstockForStocksReportDialog *pVw = (CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd;
	pVw->OnExportXLSSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CPMSSummaryImportExportInstockForStocksReportDialog *pVw = (CPMSSummaryImportExportInstockForStocksReportDialog *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddPMSSummaryImportExportInstockForStocksReportDialogFnc(CWnd *pWnd){
	 return ((CPMSSummaryImportExportInstockForStocksReportDialog*)pWnd)->OnAddPMSSummaryImportExportInstockForStocksReportDialog();
} 
static int _OnEditPMSSummaryImportExportInstockForStocksReportDialogFnc(CWnd *pWnd){
	 return ((CPMSSummaryImportExportInstockForStocksReportDialog*)pWnd)->OnEditPMSSummaryImportExportInstockForStocksReportDialog();
} 
static int _OnDeletePMSSummaryImportExportInstockForStocksReportDialogFnc(CWnd *pWnd){
	 return ((CPMSSummaryImportExportInstockForStocksReportDialog*)pWnd)->OnDeletePMSSummaryImportExportInstockForStocksReportDialog();
} 
static int _OnSavePMSSummaryImportExportInstockForStocksReportDialogFnc(CWnd *pWnd){
	 return ((CPMSSummaryImportExportInstockForStocksReportDialog*)pWnd)->OnSavePMSSummaryImportExportInstockForStocksReportDialog();
} 
static int _OnCancelPMSSummaryImportExportInstockForStocksReportDialogFnc(CWnd *pWnd){
	 return ((CPMSSummaryImportExportInstockForStocksReportDialog*)pWnd)->OnCancelPMSSummaryImportExportInstockForStocksReportDialog();
} 
CPMSSummaryImportExportInstockForStocksReportDialog::CPMSSummaryImportExportInstockForStocksReportDialog(CWnd *pParent)
{

	m_nDlgWidth = 585;
	m_nDlgHeight = 585;
	SetDefaultValues();
}
CPMSSummaryImportExportInstockForStocksReportDialog::~CPMSSummaryImportExportInstockForStocksReportDialog(){
}
void CPMSSummaryImportExportInstockForStocksReportDialog::OnCreateComponents(){
	m_wndReportParameters.Create(this, _T("Report Parameters"), 5, 5, 574, 550);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 285, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 290, 30, 370, 55);
	m_wndToDate.Create(this,375, 30, 570, 55); 
	m_wndTypeLabel.Create(this, _T("Type"), 10, 60, 90, 85);
	m_wndType.Create(this,95, 60, 570, 85); 
	m_wndOriginalLabel.Create(this, _T("Original"), 10, 90, 90, 115);
	m_wndOriginal.Create(this,95, 90, 570, 115); 
	m_wndStockList.Create(this,10, 120, 570, 545); 
	m_wndCheck.Create(this, _T("Sort by group"), 5, 555, 160, 580);
	m_wndExportXLS.Create(this, _T("Export XLS"), 495, 555, 575, 580);



}
void CPMSSummaryImportExportInstockForStocksReportDialog::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
//	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndType.SetCheckValue(true);
	m_wndType.LimitText(35);
	m_wndOriginal.SetCheckValue(true);
	m_wndOriginal.LimitText(35);


	m_wndType.InsertColumn(0, _T("ID"), CFMT_NUMBER, 70);
	m_wndType.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);


	m_wndOriginal.InsertColumn(0, _T("ID"), CFMT_NUMBER, 30);
	m_wndOriginal.InsertColumn(1, _T("Description"), CFMT_TEXT, 125);


	m_wndStockList.InsertColumn(0, _T("Stock ID"), CFMT_NUMBER, 80);
	m_wndStockList.InsertColumn(1, _T("Stock Name"), CFMT_TEXT, 340);
	m_wndStockList.SetCheckBox(true);

}
void CPMSSummaryImportExportInstockForStocksReportDialog::OnSetWindowEvents(){
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndType.SetEvent(WE_SELENDOK, _OnTypeSelendokFnc);
	//m_wndType.SetEvent(WE_SETFOCUS, _OnTypeSetfocusFnc);
	//m_wndType.SetEvent(WE_KILLFOCUS, _OnTypeKillfocusFnc);
	m_wndType.SetEvent(WE_SELCHANGE, _OnTypeSelectChangeFnc);
	m_wndType.SetEvent(WE_LOADDATA, _OnTypeLoadDataFnc);
	//m_wndType.SetEvent(WE_ADDNEW, _OnTypeAddNewFnc);
	m_wndOriginal.SetEvent(WE_SELENDOK, _OnOriginalSelendokFnc);
	//m_wndOriginal.SetEvent(WE_SETFOCUS, _OnOriginalSetfocusFnc);
	//m_wndOriginal.SetEvent(WE_KILLFOCUS, _OnOriginalKillfocusFnc);
	m_wndOriginal.SetEvent(WE_SELCHANGE, _OnOriginalSelectChangeFnc);
	m_wndOriginal.SetEvent(WE_LOADDATA, _OnOriginalLoadDataFnc);
	//m_wndOriginal.SetEvent(WE_ADDNEW, _OnOriginalAddNewFnc);
	m_wndStockList.SetEvent(WE_SELCHANGE, _OnStockListSelectChangeFnc);
	m_wndStockList.SetEvent(WE_LOADDATA, _OnStockListLoadDataFnc);
	m_wndStockList.SetEvent(WE_DBLCLICK, _OnStockListDblClickFnc);
	m_wndStockList.AddEvent(1, _T("Delete"), _OnStockListDeleteItemFnc, 0, VK_DELETE, 0);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExportXLS.SetEvent(WE_CLICK, _OnExportXLSSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);

	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_szFromDate.Format(_T("%s 00:00"), pMF->GetSysDate());
	m_szToDate.Format(_T("%s 23:59"), pMF->GetSysDate());
	OnStockListLoadData();
	SetMode(VM_EDIT);
}
void CPMSSummaryImportExportInstockForStocksReportDialog::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_TextEx(pDX, m_wndOriginal.GetDlgCtrlID(), m_szOriginalKey);
	DDX_Check(pDX, m_wndCheck.GetDlgCtrlID(), m_bCheck);

}
void CPMSSummaryImportExportInstockForStocksReportDialog::GetDataToScreen(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPMSSummaryImportExportInstockForStocksReportDialog::GetScreenToData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

}
void CPMSSummaryImportExportInstockForStocksReportDialog::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szTypeKey.Empty();
	m_szOriginalKey.Empty();

}
int CPMSSummaryImportExportInstockForStocksReportDialog::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, 1, 2, -1); 
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
/*void CPMSSummaryImportExportInstockForStocksReportDialog::OnFromDateChange(){
	
} */
/*void CPMSSummaryImportExportInstockForStocksReportDialog::OnFromDateSetfocus(){
	
} */
/*void CPMSSummaryImportExportInstockForStocksReportDialog::OnFromDateKillfocus(){
	
} */
int CPMSSummaryImportExportInstockForStocksReportDialog::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMSSummaryImportExportInstockForStocksReportDialog::OnToDateChange(){
	
} */
/*void CPMSSummaryImportExportInstockForStocksReportDialog::OnToDateSetfocus(){
	
} */
/*void CPMSSummaryImportExportInstockForStocksReportDialog::OnToDateKillfocus(){
	
} */
int CPMSSummaryImportExportInstockForStocksReportDialog::OnToDateCheckValue(){
	return 0;
} 
void CPMSSummaryImportExportInstockForStocksReportDialog::OnTypeSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSSummaryImportExportInstockForStocksReportDialog::OnTypeSelendok(){
	 
}
/*void CPMSSummaryImportExportInstockForStocksReportDialog::OnTypeSetfocus(){
	
}*/
/*void CPMSSummaryImportExportInstockForStocksReportDialog::OnTypeKillfocus(){
	
}*/
long CPMSSummaryImportExportInstockForStocksReportDialog::OnTypeLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndType.IsSearchKey() && !m_szTypeKey.IsEmpty()){
		//szWhere.Format(_T(" WHERE pmdt_id='%s' "), m_szTypeKey);
	};
	//szWhere.Format(_T(" AND hfg_moduleid = '%s'"), pMF->GetModuleID());
	m_wndType.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT hfg_id as id, hfg_name as name FROM hms_fee_group WHERe substr(hfg_id,1, 1)='A' %s ORDER BY hfg_index, hfg_id "), szWhere);

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndType.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), 
			rs.GetValue(_T("group")), 
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPMSSummaryImportExportInstockForStocksReportDialog::OnTypeAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSSummaryImportExportInstockForStocksReportDialog::OnOriginalSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSSummaryImportExportInstockForStocksReportDialog::OnOriginalSelendok(){
	 
}
/*void CPMSSummaryImportExportInstockForStocksReportDialog::OnOriginalSetfocus(){
	
}*/
/*void CPMSSummaryImportExportInstockForStocksReportDialog::OnOriginalKillfocus(){
	
}*/
long CPMSSummaryImportExportInstockForStocksReportDialog::OnOriginalLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndOriginal.IsSearchKey() && str2int(m_szOriginalKey) > 0){
		szWhere.Format(_T(" and mpr_product_resource_id=%d"), str2int(m_szOriginalKey));
	};
	szWhere.AppendFormat(_T(" AND mpr_org_id = '%s'"), pMF->GetModuleID());
	m_wndOriginal.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT * FROM m_product_resource WHERE mpr_isactive='Y' %s ORDER BY mpr_product_resource_id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndOriginal.AddItems(
			rs.GetValue(_T("mpr_product_resource_id")), 
			rs.GetValue(_T("mpr_name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CPMSSummaryImportExportInstockForStocksReportDialog::OnOriginalAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CPMSSummaryImportExportInstockForStocksReportDialog::OnStockListDblClick(){
	
} 
void CPMSSummaryImportExportInstockForStocksReportDialog::OnStockListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CPMSSummaryImportExportInstockForStocksReportDialog::OnStockListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
} 
long CPMSSummaryImportExportInstockForStocksReportDialog::OnStockListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	return pMF->LoadStorageIntoListCtrl(&m_wndStockList);

}
void CPMSSummaryImportExportInstockForStocksReportDialog::OnPrintSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CPMSSummaryImportExportInstockForStocksReportDialog::OnExportXLSSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere, tmpStr, szSubSQL, szSumGroup, szOrderby, szTemp;
	CString szTypeID, szItemID;						  
	double nPrice=0;
	int nStockID;
	BeginWaitCursor();
	CString szType, szTypeName;
	for (int i=0 ; i<= m_wndType.GetCount(); i++)
	{
		 if(m_wndType.GetCheck(i))
		 {
			 m_wndType.SetCurSel(i);
			 if(!szType.IsEmpty())
			 {
				szType += _T(",");
				szTypeName +=_T(", ");
			 }
			 szType.AppendFormat(_T("'%s'"), m_wndType.GetCurrent(0));				
			 szTypeName.AppendFormat(_T("%s"), m_wndType.GetCurrent(1));
		}		
	 }
	

	if(!szType.IsEmpty())	
		szWhere.AppendFormat(_T(" and product_producttype in(%s)"), szType);	
	else
		szTypeName.Format(_T("T\x1EA5t \x63\x1EA3 lo\x1EA1i thu\x1ED1\x63"));

	if(str2int(m_szOriginalKey) > 0)
	{
		szWhere.AppendFormat(_T(" and product_resourceid=%d "), str2int(m_szOriginalKey));
	}
	szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	if (m_bCheck)
		szOrderby.Format(_T("ORDER BY product_categoryid, product_categoryname, name, unit, unitprice"));
	else
		szOrderby.Format(_T("ORDER BY name, unit, unitprice"));
	
	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	CellFormat df(&xls), cf(&xls), hf(&xls);
	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	df.SetFontSize(9);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER);
	hf.SetFontSize(9);
	cf.SetFontName(_T("Segoe UI"));
	cf.SetFontSize(9);
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 5);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(5, 10);
	xls.SetColumnWidth(6, 12);
	int col = 3;
	for (int i =0; i < m_wndStockList.GetItemCount(); i++){
		if(m_wndStockList.GetCheck(i))
		{
			xls.SetColumnWidth(col++, 8);
			xls.SetColumnWidth(col++, 10);
		}
	}

	if(col == 3){
		ShowMessageBox(_T("You must select a stock"));
		return;
	}

	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0,2, 15);	
	xls.SetCellText(0, 0, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 9);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 9);
	TranslateString(_T("Summary of import - export - instock for multiple stock report"), szTemp);
	StringUpper(szTemp, tmpStr);
	xls.SetCellText(0, 2, tmpStr, FMT_TEXT|FMT_CENTER, true, 10);
	xls.SetCellMergedColumns(0,3,15);
	tmpStr.Format(_T("T\x1EEB ngày %s \x111\x1EBFn ngày %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(0, 3, tmpStr, &df);
	int nRow = 3;

	if(!m_szTypeKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("Lo\x1EA1i: %s"), szTypeName);
		xls.SetCellMergedColumns(0, nRow, 15);
		xls.SetCellText(0, nRow, tmpStr, &cf);
	}

	if(!m_szOriginalKey.IsEmpty())
	{
		nRow++;
		tmpStr.Format(_T("Ngu\x1ED3n: %s"), m_wndOriginal.GetCurrent(1));
		xls.SetCellMergedColumns(0, nRow, 15);
		xls.SetCellText(0, nRow, tmpStr, &cf);
	}
	nRow++;
	xls.SetCellMergedRows(0, nRow, 1);
	xls.SetCellText(0, nRow, _T("STT"), &hf);
	xls.SetCellMergedRows(1, nRow, 1);
	xls.SetCellText(1, nRow, _T("Tên thu\x1ED1\x63/ HL"), &hf);
	xls.SetCellMergedRows(2, nRow, 1);
	xls.SetCellText(2, nRow, _T("Ho\x1EA1t \x63h\x1EA5t"), &hf);
	xls.SetCellMergedRows(3, nRow, 1);
	xls.SetCellText(3, nRow, _T("\x110\x1A1n v\x1ECB"), &hf); 
	xls.SetCellMergedRows(4, nRow, 1);
	xls.SetCellText(4, nRow, _T("H\xE3ng SX"), &hf);  
	xls.SetCellMergedRows(6, nRow, 1);
	xls.SetCellText(5, nRow, _T("N\x1B0\x1EDB\x63 SX"), &hf);  
	xls.SetCellMergedRows(6, nRow, 1);
	xls.SetCellText(6, nRow, _T("\x110\x1A1n giá"), &hf);      
	col = 7;
	for (int i =0; i < m_wndStockList.GetItemCount(); i++){

		if(m_wndStockList.GetCheck(i))
		{
			tmpStr.Format(_T("%s"), m_wndStockList.GetItemText(i, 1));
			xls.SetCellMergedColumns(col, nRow, 4);
			xls.SetCellText(col, nRow, tmpStr, &hf); 

			//Ton day ky
			//xls.SetCellMergedColumns(col, nRow+1, 2);
			xls.SetColumnWidth(col, 7);
			xls.SetCellText(col, nRow+1, _T("T\x1ED3n \x111\x1EA7u"), &hf);      
			//xls.SetCellText(col, nRow+2, _T("S\x1ED1 l\x1B0\x1EE3ng"), &hf);      
			//xls.SetCellText(col+1, nRow+2, _T("S\x1ED1 ti\x1EC1n"), &hf);

			col+=1;			
			//Nhap trong ky
			//xls.SetCellMergedColumns(col, nRow+1, 2);
			xls.SetColumnWidth(col, 7);
			xls.SetCellText(col, nRow+1, _T("Nh\x1EADp"), &hf);      
			//xls.SetCellText(col, nRow+2, _T("S\x1ED1 l\x1B0\x1EE3ng"), &hf);      
			//xls.SetCellText(col+1, nRow+2, _T("S\x1ED1 ti\x1EC1n"), &hf); 
			col+=1;		
			//Xuat trong ky
			//xls.SetCellMergedColumns(col, nRow+1, 2);
			xls.SetColumnWidth(col, 7);
			xls.SetCellText(col, nRow+1, _T("Xu\x1EA5t"), &hf);      
			//xls.SetCellText(col, nRow+2, _T("S\x1ED1 l\x1B0\x1EE3ng"), &hf);      
			//xls.SetCellText(col+1, nRow+2, _T("S\x1ED1 ti\x1EC1n"), &hf); 
			col+=1;			
			//Ton cuoi ky
			//xls.SetCellMergedColumns(col, nRow+1, 2);
			xls.SetColumnWidth(col, 7);
			xls.SetCellText(col, nRow+1, _T("T\x1ED3n \x63u\x1ED1i"), &hf);      
			//xls.SetCellText(col, nRow+2, _T("S\x1ED1 l\x1B0\x1EE3ng"), &hf);      
			//xls.SetCellText(col+1, nRow+2, _T("S\x1ED1 ti\x1EC1n"), &hf); 
			col+=1;
		}
	}

	if(col > 8)
	{
		xls.SetCellMergedColumns(col, nRow, 4);
		xls.SetCellText(col, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), &hf);      
//Ton day ky
			//xls.SetCellMergedColumns(col, nRow+1, 2);
			xls.SetColumnWidth(col, 7);
			xls.SetCellText(col, nRow+1, _T("T\x1ED3n \x111\x1EA7u"), &hf);      
			//xls.SetCellText(col, nRow+2, _T("S\x1ED1 l\x1B0\x1EE3ng"), &hf);      
			//xls.SetCellText(col+1, nRow+2, _T("S\x1ED1 ti\x1EC1n"), &hf);
			

			col+=1;			
			//Nhap trong ky
			//xls.SetCellMergedColumns(col, nRow+1, 2);
			xls.SetColumnWidth(col, 7);
			xls.SetCellText(col, nRow+1, _T("Nh\x1EADp"), &hf);      
			//xls.SetCellText(col, nRow+2, _T("S\x1ED1 l\x1B0\x1EE3ng"), &hf);      
			//xls.SetCellText(col+1, nRow+2, _T("S\x1ED1 ti\x1EC1n"), &hf); 
			col+=1;		
			//Xuat trong ky
			//xls.SetCellMergedColumns(col, nRow+1, 2);
			xls.SetColumnWidth(col, 7);
			xls.SetCellText(col, nRow+1, _T("Xu\x1EA5t"), &hf);      
			//xls.SetCellText(col, nRow+2, _T("S\x1ED1 l\x1B0\x1EE3ng"), &hf);      
			//xls.SetCellText(col+1, nRow+2, _T("S\x1ED1 ti\x1EC1n"), &hf); 
			col+=1;			
			//Ton cuoi ky
			//xls.SetCellMergedColumns(col, nRow+1, 2);
			xls.SetColumnWidth(col, 7);
			xls.SetCellText(col, nRow+1, _T("T\x1ED3n \x63u\x1ED1i"), &hf);      
			//xls.SetCellText(col, nRow+2, _T("S\x1ED1 l\x1B0\x1EE3ng"), &hf);      
			//xls.SetCellText(col+1, nRow+2, _T("S\x1ED1 ti\x1EC1n"), &hf); 
			col+=1;
	}
	nRow += 2;	

	
	bool bFirst=true;
	int n=0;
	for (int i =0; i < m_wndStockList.GetItemCount(); i++){

		if(m_wndStockList.GetCheck(i))
		{
			n = str2int(m_wndStockList.GetItemText(i, 0));
			szSumGroup.AppendFormat(_T(", sum(begin_qty%d) as beginqty%d"), n, n);
			szSumGroup.AppendFormat(_T(", sum(imp_qty%d) as impqty%d"), n, n);
			szSumGroup.AppendFormat(_T(", sum(exp_qty%d) as expqty%d"), n, n);
			szSumGroup.AppendFormat(_T(", sum(begin_qty%d+imp_qty%d-exp_qty%d) as endqty%d"), n, n, n, n);

			nStockID = str2int(m_wndStockList.GetItemText(i, 0));
			if(bFirst){
				szSubSQL.Format(_T(" SELECT *") \
								_T(" FROM") \
								_T(" (") \
								_T(" 	SELECT 	sitemid, ") \
								_T(" 		sum(begin_qty) as  begin_qty%d,") \
								_T(" 		sum(imp_qty) as imp_qty%d,") \
								_T(" 		sum(exp_qty) as exp_qty%d ") \
								_T(" 	FROM") \
								_T(" 	(") \
								_T("  		SELECT * FROM") \
								_T("  		(") \
								_T("  			SELECT impstockid as stockid, impdate as iodate, sitemid, 0 as begin_qty, impqty as imp_qty, 0 as exp_qty") \
								_T("  			FROM miv") \
								_T("  			UNION ALL") \
								_T("  			SELECT expstockid, expdate, sitemid, 0, 0, expqty") \
								_T("  			FROM mev") \
								_T("  		) iotbl") \
								_T("  		WHERE iodate between cast_string2timestamp ('%s' )and cast_string2timestamp ('%s' )") \
								_T("  		UNION ALL") \
								_T("  		SELECT stockid, iodate, sitemid, coalesce(impqty-expqty, 0), 0, 0 FROM") \
								_T("  		(") \
								_T("  			SELECT impstockid as stockid, impdate as iodate, sitemid,  0 as periodqty, impqty, 0 as expqty") \
								_T("  			FROM miv") \
								_T("  			UNION ALL") \
								_T("  			SELECT expstockid, expdate, sitemid, 0, 0, expqty") \
								_T("  			FROM mev") \
								_T("  		) ptbl") \
								_T("  		WHERE iodate < cast_string2timestamp ('%s' )") \
								_T(" 	 ) tblx%d") \
								_T(" 	 WHERE stockid=%d ") \
								_T(" 	 and sitemid > 0") \
								_T(" 	 GROUP BY sitemid") \
								_T(" 	 HAVING sum(begin_qty+imp_qty+exp_qty) > 0") \
								_T(" ) tbl%d"), n, n, n, m_szFromDate, m_szToDate, m_szFromDate, n, nStockID, n);
				bFirst = false;
			}
			else
			{
				szSubSQL.AppendFormat(_T(" full join") \
								_T(" (") \
								_T(" 	SELECT 	sitemid, ") \
								_T(" 		sum(begin_qty) as  begin_qty%d,") \
								_T(" 		sum(imp_qty) as imp_qty%d,") \
								_T(" 		sum(exp_qty) as exp_qty%d") \
								_T(" 	FROM") \
								_T(" 	(") \
								_T("  		SELECT * FROM") \
								_T("  		(") \
								_T("  			SELECT impstockid as stockid, impdate as iodate, sitemid, 0 as begin_qty, impqty as imp_qty, 0 as exp_qty") \
								_T("  			FROM miv") \
								_T("  			UNION ALL") \
								_T("  			SELECT impstockid, expdate, sitemid, 0, 0, expqty") \
								_T("  			FROM mev") \
								_T("  		) iotbl") \
								_T("  		WHERE iodate between cast_string2timestamp('%s') and cast_string2timestamp('%s') ") \
								_T("  		UNION ALL") \
								_T("  		SELECT stockid, iodate, sitemid, coalesce(impqty-expqty, 0), 0, 0 FROM") \
								_T("  		(") \
								_T("  			SELECT impstockid as stockid, impdate as iodate, sitemid,  0 as periodqty, impqty, 0 as expqty") \
								_T("  			FROM miv") \
								_T("  			UNION ALL") \
								_T("  			SELECT impstockid, expdate, sitemid, 0, 0, expqty") \
								_T("  			FROM mev") \
								_T("  		) ptbl") \
								_T("  		WHERE iodate < cast_string2timestamp('%s') ") \
								_T(" 	 ) tblx%d") \
								_T(" 	 WHERE stockid=%d") \
								_T(" 	 and sitemid > 0") \
								_T(" 	 GROUP BY sitemid") \
								_T(" 	 HAVING sum(begin_qty+imp_qty+exp_qty) > 0") \
								_T(" ) tbl%d using(sitemid)"), n, n, n, m_szFromDate, m_szToDate, m_szFromDate, n, nStockID, n);
			}
		}
	}
				szSQL.Format(_T(" SELECT") \
				_T("   product_categoryid                                         AS idx,") \
				_T("   product_categoryname                                                        AS groupname,") \
				_T("   product_code                                                         AS id,") \
				_T("   product_name                                                         AS NAME,") \
				_T("   product_uomname                                          AS unit,") \
				_T("   product_manufacturename                         AS Origin,") \
				_T("   product_countryname AS country,") \
				_T("   product_taxprice                                                    AS unitprice ") \
				_T("		 %s ") \
				_T(" FROM (%s) tblreport  ") \
				_T(" LEFT JOIN m_productitem_view ON (sitemid=product_item_id)  ") \
				_T(" WHERE length(product_code) > 0 %s") \
				_T(" GROUP BY product_categoryid, product_categoryname, product_code, product_name, product_uomname, product_manufacturename, product_countryname, product_taxprice") \
				_T(" %s"), szSumGroup, szSubSQL, szWhere, szOrderby);
	//QueryOpener(szSQL);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	CString szNewLine, szOldLine;
	long nGroupBeginQty=0, nGroupImportQty=0, nGroupExportQty=0, nGroupEndQty=0;
	long nBeginQty=0, nImportQty=0, nExportQty=0, nEndQty=0;
	int nItem = 1;
	while(!rs.IsEOF())
	{
		if (m_bCheck)
		{
			rs.GetValue(_T("idx"), szNewLine);
			if(szNewLine != szOldLine && !szNewLine.IsEmpty()){				
				CString szField, szAmount;			
				tmpStr.Format(_T("%s"), szNewLine);
				xls.SetCellMergedColumns(0,nRow, 3);
				tmpStr.Format(_T("%s. %s"), szNewLine, rs.GetValue(_T("groupname")));
				xls.SetCellText(0, nRow, tmpStr, FMT_TEXT,true, 9,0);
				szOldLine = szNewLine;
				nItem=1;
				nRow ++;
			}
		}
		tmpStr.Format(_T("%d"), nItem++);
		xls.SetCellText(0,nRow,tmpStr, FMT_INTEGER, false, 9);		
		rs.GetValue(_T("name"), tmpStr);
		xls.SetCellText(1,nRow,tmpStr, FMT_TEXT, false, 9);
		rs.GetValue(_T("genericname"), tmpStr);
		xls.SetCellText(2,nRow,tmpStr, FMT_TEXT, false, 9);
		rs.GetValue(_T("unit"), tmpStr);
		xls.SetCellText(3,nRow,tmpStr, FMT_TEXT, false, 9);
		rs.GetValue(_T("Origin"), tmpStr);
		xls.SetCellText(4,nRow,tmpStr, FMT_TEXT, false, 9);
		rs.GetValue(_T("country"), tmpStr);
		xls.SetCellText(5,nRow,tmpStr, FMT_TEXT, false, 9);
		rs.GetValue(_T("unitprice"), tmpStr);		
		xls.SetCellText(6,nRow,tmpStr, FMT_NUMBER1, false, 9);
		col = 7;
		int nIndex=11;
		CString szFieldName;
		int nStockID = 0;
		long nTotalBeginQty=0, nTotalImportQty=0, nTotalExportQty=0, nTotalEndQty=0;
		
		for (int i =0; i < m_wndStockList.GetItemCount(); i++){
			if(m_wndStockList.GetCheck(i)){
				
				nStockID = str2int(m_wndStockList.GetItemText(i, 0));
				szFieldName.Format(_T("beginqty%d"), nStockID);
				rs.GetValue(szFieldName, tmpStr);
				nTotalBeginQty += str2long(tmpStr);
				if(str2float(tmpStr) == 0)
					tmpStr.Empty();

				xls.SetCellText(col++, nRow, tmpStr, FMT_NUMBER1, false, 9);
				szFieldName.Format(_T("impqty%d"), nStockID);
				rs.GetValue(szFieldName, tmpStr);
				nTotalImportQty += str2long(tmpStr);
				if(str2float(tmpStr) == 0)
					tmpStr.Empty();
				xls.SetCellText(col++, nRow, tmpStr, FMT_NUMBER1, false, 9);
		
				szFieldName.Format(_T("expqty%d"), nStockID);
				rs.GetValue(szFieldName, tmpStr);
				nTotalExportQty += str2long(tmpStr);
				if(str2float(tmpStr) == 0)
					tmpStr.Empty();
				xls.SetCellText(col++, nRow, tmpStr, FMT_NUMBER1, false, 9);
				szFieldName.Format(_T("endqty%d"), nStockID);
				rs.GetValue(szFieldName, tmpStr);
				nTotalEndQty += str2long(tmpStr);
				if(str2float(tmpStr) == 0)
					tmpStr.Empty();
				xls.SetCellText(col++, nRow, tmpStr, FMT_NUMBER1, false, 9);
				nIndex ++;
			}
		}
		if(col > 8){
			if(nTotalBeginQty > 0)
				tmpStr.Format(_T("%ld"), nTotalBeginQty);
			else
				tmpStr.Empty();
			xls.SetCellText(col++, nRow, tmpStr, FMT_NUMBER1, false, 9);
			if(nTotalImportQty > 0)
				tmpStr.Format(_T("%ld"), nTotalImportQty);
			else
				tmpStr.Empty();
			xls.SetCellText(col++, nRow, tmpStr, FMT_NUMBER1, false, 9);
			if(nTotalExportQty > 0)
				tmpStr.Format(_T("%ld"), nTotalExportQty);
			else
				tmpStr.Empty();
			xls.SetCellText(col++, nRow, tmpStr, FMT_NUMBER1, false, 9);
			if(nTotalEndQty > 0)
				tmpStr.Format(_T("%ld"), nTotalEndQty);
			else
				tmpStr.Empty();
			xls.SetCellText(col++, nRow, tmpStr, FMT_NUMBER1, false, 9);
		}

		nRow++;
		rs.MoveNext();
	}

	xls.Save(_T("Exports\\TONGHOPNHAPXUATTONTHEONHIEUKHO.XLS"));
	EndWaitCursor();

} 
void CPMSSummaryImportExportInstockForStocksReportDialog::OnCloseSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();

} 
int CPMSSummaryImportExportInstockForStocksReportDialog::OnAddPMSSummaryImportExportInstockForStocksReportDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CPMSSummaryImportExportInstockForStocksReportDialog::OnEditPMSSummaryImportExportInstockForStocksReportDialog(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CPMSSummaryImportExportInstockForStocksReportDialog::OnDeletePMSSummaryImportExportInstockForStocksReportDialog(){
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
 		OnCancelPMSSummaryImportExportInstockForStocksReportDialog(); 
 	}
	return 0;
}
int CPMSSummaryImportExportInstockForStocksReportDialog::OnSavePMSSummaryImportExportInstockForStocksReportDialog(){
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
 		//OnPMSSummaryImportExportInstockForStocksReportDialogListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CPMSSummaryImportExportInstockForStocksReportDialog::OnCancelPMSSummaryImportExportInstockForStocksReportDialog(){
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
int CPMSSummaryImportExportInstockForStocksReportDialog::OnPMSSummaryImportExportInstockForStocksReportDialogListLoadData(){
	return 0;
}
BEGIN_MESSAGE_MAP(CPMSSummaryImportExportInstockForStocksReportDialog, CGuiView)
	ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSSummaryImportExportInstockForStocksReportDialog::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
