#include "stdafx.h"
#include "PMImportExportInstockbyCategory.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include "StringToken.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMImportExportInstockbyCategory *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMImportExportInstockbyCategory *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMImportExportInstockbyCategory *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMImportExportInstockbyCategory *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMImportExportInstockbyCategory *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMImportExportInstockbyCategory *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMImportExportInstockbyCategory *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMImportExportInstockbyCategory *)pWnd)->OnToDateCheckValue();
} 
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CPMImportExportInstockbyCategory*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CPMImportExportInstockbyCategory*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPMImportExportInstockbyCategory*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CPMImportExportInstockbyCategory*)pWnd)->OnListDeleteItem();
} 
static void _OnServiceSelectFnc(CWnd *pWnd){
	  ((CPMImportExportInstockbyCategory*)pWnd)->OnServiceSelect();
}
static void _OnInsuranceSelectFnc(CWnd *pWnd){
	  ((CPMImportExportInstockbyCategory*)pWnd)->OnInsuranceSelect();
}
static void _OnPolicySelectFnc(CWnd *pWnd){
	  ((CPMImportExportInstockbyCategory*)pWnd)->OnPolicySelect();
}
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMImportExportInstockbyCategory *pVw = (CPMImportExportInstockbyCategory *)pWnd;
	pVw->OnExportSelect();
} 
CPMImportExportInstockbyCategory::CPMImportExportInstockbyCategory(CWnd *pParent){

	m_szStorageCategory = _T("I,S,P");
	m_szCondition = _T("S");
	m_szStorageID = _T("3, 14");
	m_szTitle = _T("\x44\x1ECA\x43H V\x1EE4");
	m_szPrice = _T("product_saleprice1");
	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CPMImportExportInstockbyCategory::~CPMImportExportInstockbyCategory(){
}
void CPMImportExportInstockbyCategory::OnCreateComponents(){
	m_wndImportExportInstockByCategory.Create(this, _T("Import Export Instock by Category"), 5, 5, 480, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 240, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 245, 30, 325, 55);
	m_wndToDate.Create(this,330, 30, 475, 55); 
	m_wndList.Create(this,5, 95, 480, 250); 
	m_wndService.Create(this, _T("Service"), 10, 60, 160, 85);
	m_wndInsurance.Create(this, _T("Insurance"), 165, 60, 315, 85);
	m_wndPolicy.Create(this, _T("Policy"), 320, 60, 470, 85);
	m_wndExport.Create(this, _T("&Export"), 400, 255, 480, 280);

}
void CPMImportExportInstockbyCategory::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);


	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndList.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);

}
void CPMImportExportInstockbyCategory::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);
	m_wndList.AddEvent(1, _T("Delete"), _OnListDeleteItemFnc, 0, VK_DELETE, 0);
	m_wndService.SetEvent(WE_CLICK, _OnServiceSelectFnc);
	m_wndInsurance.SetEvent(WE_CLICK, _OnInsuranceSelectFnc);
	m_wndPolicy.SetEvent(WE_CLICK, _OnPolicySelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	OnListLoadData();
	m_mapType[_T("I")] = _T("\x42H");
	m_mapType[_T("S")] = _T("\x44V");
	m_mapType[_T("P")] = _T("\x42\x110");
	UpdateData(false);

}
void CPMImportExportInstockbyCategory::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_Radio(pDX, m_wndService.GetDlgCtrlID(), m_nService);
	DDX_Radio(pDX, m_wndInsurance.GetDlgCtrlID(), m_nInsurance);
	DDX_Radio(pDX, m_wndPolicy.GetDlgCtrlID(), m_nPolicy);

}
void CPMImportExportInstockbyCategory::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_nService=0;
	m_nInsurance=1;
	m_nPolicy=1;

}
int CPMImportExportInstockbyCategory::SetMode(int nMode){
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
/*void CPMImportExportInstockbyCategory::OnFromDateChange(){
	
} */
/*void CPMImportExportInstockbyCategory::OnFromDateSetfocus(){
	
} */
/*void CPMImportExportInstockbyCategory::OnFromDateKillfocus(){
	
} */
int CPMImportExportInstockbyCategory::OnFromDateCheckValue(){
	return 0;
} 
/*void CPMImportExportInstockbyCategory::OnToDateChange(){
	
} */
/*void CPMImportExportInstockbyCategory::OnToDateSetfocus(){
	
} */
/*void CPMImportExportInstockbyCategory::OnToDateKillfocus(){
	
} */
int CPMImportExportInstockbyCategory::OnToDateCheckValue(){
	return 0;
} 
void CPMImportExportInstockbyCategory::OnListDblClick(){
	
} 
void CPMImportExportInstockbyCategory::OnListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CPMImportExportInstockbyCategory::OnListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
} 
long CPMImportExportInstockbyCategory::OnListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND msl_category = '%s' AND msl_type = 'A'"), m_szCondition);
	return pMF->LoadStorageIntoListCtrl(&m_wndList, szFilter);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndList.BeginLoad(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
*/
	return 0;
}
void CPMImportExportInstockbyCategory::OnServiceSelect(){
	m_szCondition = _T("S");
	m_szStorageID = _T("3, 14");
	m_szTitle = _T("\x44\x1ECA\x43H V\x1EE4");
	m_szPrice = _T("product_saleprice1");
	OnListLoadData();
}
void CPMImportExportInstockbyCategory::OnInsuranceSelect(){
	m_szCondition = _T("I");
	m_szStorageID = _T("2, 7");
	m_szTitle = _T("\x42\x1EA2O HI\x1EC2M");
	m_szPrice = _T("product_saleprice2");
	OnListLoadData();
}
void CPMImportExportInstockbyCategory::OnPolicySelect(){
	m_szCondition = _T("P");
	m_szStorageID = _T("8, 9");
	m_szTitle = _T("\x42\x1ED8 \x110\x1ED8I");
	m_szPrice = _T("product_saleprice3");
	OnListLoadData();
}
void CPMImportExportInstockbyCategory::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CExcel xls;
	CString szSQL, tmpStr, tmpStr2, tmpStr3;
	CStringToken token(m_szStorageCategory);
	double nTmp = 0;
	long double nTotal[10];
	int nIdx = 1, nCol = 0, nRow = 0, k = 0, j = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	for (int i = 0; i < 10; i++)
	{
		nTotal[i] = 0;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 16);
	xls.SetCellMergedColumns(0, 3, 16);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr.Format(_T("\x42\xC1O \x43\xC1O NH\x1EACP, \x58U\x1EA4T T\x1ED2N KHO %s"), m_szTitle);
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	for (int i = 3; i <= 21; i+=2)
	{
		xls.SetCellMergedColumns(i, 4, 2);
	}
	xls.SetCellText(0, 4, _T("STT"), 4098, true);
	xls.SetCellText(1, 4, _T("T\xEAn thu\x1ED1\x63"), 4098, true);
	xls.SetCellText(2, 4, _T("\x110V"), 4098, true);
	xls.SetCellText(3, 4, _T("T\x1ED3n \x111\x1EA7u"), 4098, true);
	xls.SetCellText(3, 5, _T("SL"), 4098, true);
	xls.SetCellText(4, 5, _T("Gi\xE1 tr\x1ECB"), 4098, true);
	xls.SetCellText(5, 4, _T("Nh\x1EADp kh\xE1\x63"), 4098, true);
	xls.SetCellText(5, 5, _T("SL"), 4098, true);
	xls.SetCellText(6, 5, _T("Gi\xE1 tr\x1ECB"), 4098, true);
	xls.SetCellText(7, 4, _T("Nh\x1EADp t\x1EEB nh\xE0 \x63ung \x63\x1EA5p"), 4098, true);
	xls.SetCellText(7, 5, _T("SL"), 4098, true);
	xls.SetCellText(8, 5, _T("Gi\xE1 tr\x1ECB"), 4098, true);
	/**/
	nCol = 9;
	nRow = 4;
	for (int i = 0; i < token.GetSize(); i++)
	{
		token.GetAt(i, tmpStr2);
		//_debug(_T("%s|%s"), tmpStr2, m_szCondition);
		if (tmpStr2 == m_szCondition)
			continue;
		m_mapType.Lookup(tmpStr2, tmpStr3);
		tmpStr.Format(_T("Nh\x1EADp t\x1EEB kho %s"), tmpStr3);
		xls.SetCellText(nCol + k*4, nRow, tmpStr, 4098, true);
		xls.SetCellText(nCol + k*4, nRow + 1, _T("SL"), 4098, true);
		xls.SetCellText(nCol + k*4 + 1, nRow + 1, _T("Gi\xE1 tr\x1ECB"), 4098, true);
		tmpStr.Format(_T("\x58u\x1EA5t \x63ho kho %s"), tmpStr3);
		xls.SetCellText(nCol + k*4 + 2, nRow, tmpStr, 4098, true);
		xls.SetCellText(nCol + k*4 + 2, nRow + 1, _T("SL"), 4098, true);
		xls.SetCellText(nCol + k*4 + 3, nRow + 1, _T("Gi\xE1 tr\x1ECB"), 4098, true);
		k++;
	}
	xls.SetCellText(17, 4, _T("\x58u\x1EA5t s\x1EED \x64\x1EE5ng"), 4098, true);
	xls.SetCellText(17, 5, _T("SL"), 4098, true);
	xls.SetCellText(18, 5, _T("Gi\xE1 tr\x1ECB"), 4098, true);
	xls.SetCellText(19, 4, _T("\x58u\x1EA5t kh\xE1\x63"), 4098, true);
	xls.SetCellText(19, 5, _T("SL"), 4098, true);
	xls.SetCellText(20, 5, _T("Gi\xE1 tr\x1ECB"), 4098, true);
	xls.SetCellText(21, 4, _T("T\x1ED3n \x63u\x1ED1i"), 4098, true);
	xls.SetCellText(21, 5, _T("SL"), 4098, true);
	xls.SetCellText(22, 5, _T("Gi\xE1 tr\x1ECB"), 4098, true);
	nCol = 0;
	nRow = 6;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("product_name")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("product_uomname")), FMT_TEXT);
		xls.SetCellText(3, nRow, rs.GetValue(_T("period_qty")), FMT_NUMBER1);
		rs.GetValue(_T("period_amt"), nTmp);
		nTotal[0] += nTmp;
		xls.SetCellText(4, nRow, double2str(nTmp), FMT_NUMBER1);
		xls.SetCellText(5, nRow, rs.GetValue(_T("other_import")), FMT_NUMBER1);
		rs.GetValue(_T("other_amt"), nTmp);
		nTotal[1] += nTmp;
		xls.SetCellText(6, nRow, double2str(nTmp), FMT_NUMBER1);
		xls.SetCellText(7, nRow, rs.GetValue(_T("invoice_import")), FMT_NUMBER1);
		rs.GetValue(_T("invoice_amt"), nTmp);
		nTotal[2] += nTmp;
		xls.SetCellText(8, nRow, double2str(nTmp), FMT_NUMBER1);
		
		//xls.SetCellText(6, nRow, rs.GetValue(_T("s_import")), FMT_NUMBER1);
		//xls.SetCellText(7, nRow, rs.GetValue(_T("s_export")), FMT_NUMBER1);
		//xls.SetCellText(8, nRow, rs.GetValue(_T("i_import")), FMT_NUMBER1);
		//xls.SetCellText(9, nRow, rs.GetValue(_T("i_export")), FMT_NUMBER1);
		//xls.SetCellText(10, nRow, rs.GetValue(_T("p_import")), FMT_NUMBER1);
		//xls.SetCellText(11, nRow, rs.GetValue(_T("p_export")), FMT_NUMBER1);
		k = 0;
		j = 3;
		nCol = 9;
		for (int i = 0; i < token.GetSize(); i++)
		{
			token.GetAt(i, tmpStr2);
			if (tmpStr2 == m_szCondition)
				continue;
			tmpStr.Format(_T("%s_import"), tmpStr2);
			xls.SetCellText(nCol + k*4, nRow, rs.GetValue(tmpStr), FMT_NUMBER1);
			tmpStr.Format(_T("%s_import_amt"), tmpStr2);
			rs.GetValue(tmpStr, nTmp);
			nTotal[j++] += nTmp;
			xls.SetCellText(nCol + k*4 + 1, nRow, double2str(nTmp), FMT_NUMBER1);
			tmpStr.Format(_T("%s_export"), tmpStr2);
			xls.SetCellText(nCol + k*4 + 2, nRow, rs.GetValue(tmpStr), FMT_NUMBER1);
			tmpStr.Format(_T("%s_export_amt"), tmpStr2);
			rs.GetValue(tmpStr, nTmp);
			nTotal[j++] += nTmp;
			xls.SetCellText(nCol + k*4 + 3, nRow, double2str(nTmp), FMT_NUMBER1);
			k++;
		}
		xls.SetCellText(17, nRow, rs.GetValue(_T("patient_export")), FMT_NUMBER1);
		rs.GetValue(_T("patient_export_amt"), nTmp);
		nTotal[7] += nTmp;
		xls.SetCellText(18, nRow, double2str(nTmp), FMT_NUMBER1);
		xls.SetCellText(19, nRow, rs.GetValue(_T("other_export")), FMT_NUMBER1);
		rs.GetValue(_T("other_export_amt"), nTmp);
		nTotal[8] += nTmp;
		xls.SetCellText(20, nRow, double2str(nTmp), FMT_NUMBER1);
		xls.SetCellText(21, nRow, rs.GetValue(_T("instock_qty")), FMT_NUMBER1);
		rs.GetValue(_T("instock_qty_amt"), nTmp);
		nTotal[9] += nTmp;
		xls.SetCellText(22, nRow, double2str(nTmp), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellText(1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
	for (int i = 0; i < 10; i++)
	{
		tmpStr.Format(_T("%f"), nTotal[i]);
		_debug(_T("%s"), tmpStr);
		xls.SetCellText((i+2)*2, nRow, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\Nhap xuat ton theo loai kho.xls"));

} 

CString CPMImportExportInstockbyCategory::GetQueryString(){
	CString szSQL, szWhere1, szWhere2, szWhere3, szWhere4;
	szWhere1.Format(_T(" AND impstockid IN (%s) AND impdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
		, m_szStorageID, m_szFromDate, m_szToDate);
	szWhere2.Format(_T(" AND expstockid IN (%s) AND expdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
		, m_szStorageID, m_szFromDate, m_szToDate);
	szWhere3.Format(_T(" AND impstockid IN (%s) AND impdate < cast_string2timestamp('%s')")
		, m_szStorageID, m_szFromDate);
	szWhere4.Format(_T(" AND expstockid IN (%s) AND expdate < cast_string2timestamp('%s')")
		, m_szStorageID, m_szFromDate);
	szSQL.Format(_T(" SELECT   product_groupid, ") \
				_T("           product_groupname, ") \
				_T("           product_id, ") \
				_T("           product_name, ") \
				_T("           product_uomname, ") \
				_T("           SUM(period_qty) period_qty, ") \
				_T("           SUM(period_qty*%s) period_amt, ") \
				_T("           SUM(invoice_import) invoice_import, ") \
				_T("           SUM(invoice_import*%s) invoice_amt, ") \
				_T("           SUM(other_import) other_import, ") \
				_T("           SUM(other_import*%s) other_amt, ") \
				_T("           SUM(s_import) s_import, ") \
				_T("           SUM(s_import*%s) s_import_amt, ") \
				_T("           SUM(s_export) s_export, ") \
				_T("           SUM(s_export*%s) s_export_amt, ") \
				_T("           SUM(i_import) i_import, ") \
				_T("           SUM(i_import*%s) i_import_amt, ") \
				_T("           SUM(i_export) i_export, ") \
				_T("           SUM(i_export*%s) i_export_amt, ") \
				_T("           SUM(p_import) p_import, ") \
				_T("           SUM(p_import*%s) p_import_amt, ") \
				_T("           SUM(p_export) p_export, ") \
				_T("           SUM(p_export*%s) p_export_amt, ") \
				_T("           SUM(patient_export) patient_export, ") \
				_T("           SUM(patient_export*%s) patient_export_amt, ") \
				_T("           SUM(other_export) other_export, ") \
				_T("           SUM(other_export*%s) other_export_amt, ") \
				_T("           SUM(total_import) total_import, ") \
				_T("           SUM(total_export) total_export, ") \
				_T("		   SUM(period_qty+total_import-total_export) instock_qty,") \
				_T("		   SUM((period_qty+total_import-total_export)*%s) instock_qty_amt") \
				_T(" FROM      (SELECT sitemid, ") \
				_T("                   0 period_qty, ") \
				_T("                   CASE WHEN ( iotype IN ( 'RRO', 'DRO', 'CRO', 'ADO' ) OR ( iotype = 'MOV' AND storage_id NOT IN ( 3, 14, 2, 7, ") \
				_T("                                                        8, 9 ) ) ) THEN impqty ") \
				_T("                   ELSE 0 ") \
				_T("                   END other_import, ") \
				_T("                   CASE WHEN iotype = 'POO' THEN impqty ") \
				_T("                   ELSE 0 ") \
				_T("                   END invoice_import, ") \
				_T("                   CASE WHEN ( iotype = 'MOV' AND storage_id IN ( 3, 14 ) ) THEN impqty ") \
				_T("                   ELSE 0 ") \
				_T("                   END s_import, ") \
				_T("                   CASE WHEN ( iotype = 'MOV' AND storage_id IN ( 3, 14 ) ) THEN expqty ") \
				_T("                   ELSE 0 ") \
				_T("                   END s_export, ") \
				_T("                   CASE WHEN ( iotype = 'MOV' AND storage_id IN ( 2, 7 ) ) THEN impqty ") \
				_T("                   ELSE 0 ") \
				_T("                   END i_import, ") \
				_T("                   CASE WHEN ( iotype = 'MOV' AND storage_id IN ( 2, 7 ) ) THEN expqty ") \
				_T("                   ELSE 0 ") \
				_T("                   END i_export, ") \
				_T("                   CASE WHEN ( iotype = 'MOV' AND storage_id IN ( 8, 9 ) ) THEN impqty ") \
				_T("                   ELSE 0 ") \
				_T("                   END p_import, ") \
				_T("                   CASE WHEN ( iotype = 'MOV' AND storage_id IN ( 8, 9 ) ) THEN expqty ") \
				_T("                   ELSE 0 ") \
				_T("                   END p_export, ") \
				_T("                   CASE WHEN ( iotype IN ( 'DMO', 'PPO', 'CSO', 'DDO', 'SOO' )  OR ( iotype = 'MOV' AND storage_id NOT IN ( 3, 14, 2, 7, ") \
				_T("                                                        8, 9 ) ) ) THEN expqty ") \
				_T("                   ELSE 0 ") \
				_T("                   END patient_export, ") \
				_T("                   CASE WHEN iotype IN ( 'DOO', 'LOO', 'EOO', 'BOO', ") \
				_T("                                    'SLO', 'EXO', 'EXP', 'EOH', 'ADO' ) THEN expqty ") \
				_T("                   ELSE 0 ") \
				_T("                   END other_export, ") \
				_T("                   impqty total_import, ") \
				_T("                   expqty total_export ") \
				_T("            FROM   (SELECT sitemid, ") \
				_T("                           expstockid storage_id, ") \
				_T("                           iotype, ") \
				_T("                           impqty, ") \
				_T("                           0 expqty ") \
				_T("                    FROM   miv ") \
				_T("                    WHERE  1=1 %s ") \
				_T("                    UNION ALL ") \
				_T("                    SELECT sitemid, ") \
				_T("                           impstockid, ") \
				_T("                           iotype, ") \
				_T("                           0, ") \
				_T("                           expqty ") \
				_T("                    FROM   mev ") \
				_T("                    WHERE  1=1 %s) e_table1 ") \
				_T("            UNION ALL ") \
				_T("            SELECT sitemid, ") \
				_T("                   Coalesce(impqty - expqty, 0), ") \
				_T("                   0, ") \
				_T("                   0, ") \
				_T("                   0, ") \
				_T("                   0, ") \
				_T("                   0, ") \
				_T("                   0, ") \
				_T("                   0, ") \
				_T("                   0, ") \
				_T("                   0, ") \
				_T("                   0, ") \
				_T("                   0, ") \
				_T("                   0 ") \
				_T("            FROM   (SELECT sitemid, ") \
				_T("                           expstockid storage_id, ") \
				_T("                           iotype, ") \
				_T("                           impqty, ") \
				_T("                           0 expqty ") \
				_T("                    FROM   miv ") \
				_T("                    WHERE  1=1 %s ") \
				_T("                    UNION ALL ") \
				_T("                    SELECT sitemid, ") \
				_T("                           impstockid, ") \
				_T("                           iotype, ") \
				_T("                           0, ") \
				_T("                           expqty ") \
				_T("                    FROM   mev ") \
				_T("                    WHERE  1= 1 %s) e_table2) m_table ") \
				_T(" LEFT JOIN m_productitem_view ON ( sitemid = product_item_id ) ") \
				_T(" GROUP BY product_groupid,product_groupname,product_id,product_name,product_uomname ") \
				_T(" HAVING SUM(period_qty+total_import+total_export) > 0") \
				_T(" ORDER BY product_name"), 
				m_szPrice, m_szPrice, m_szPrice, m_szPrice, m_szPrice, m_szPrice, 
				m_szPrice, m_szPrice, m_szPrice, m_szPrice, m_szPrice, m_szPrice, 
				szWhere1, szWhere2, szWhere3, szWhere4);

_msg(_T("%s"), szSQL);
	return szSQL;
}