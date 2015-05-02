#include "stdafx.h"
#include "EMDepartmentUsageinDetail.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMDepartmentUsageinDetail *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMDepartmentUsageinDetail *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMDepartmentUsageinDetail *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMDepartmentUsageinDetail *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMDepartmentUsageinDetail *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMDepartmentUsageinDetail *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMDepartmentUsageinDetail *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMDepartmentUsageinDetail *)pWnd)->OnToDateCheckValue();
} 
static void _OnOrderListSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMDepartmentUsageinDetail* )pWnd)->OnOrderListSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnOrderListSelendokFnc(CWnd *pWnd){
	((CEMDepartmentUsageinDetail *)pWnd)->OnOrderListSelendok();
}
/*static void _OnOrderListSetfocusFnc(CWnd *pWnd){
	((CEMDepartmentUsageinDetail *)pWnd)->OnOrderListKillfocus();
}*/
/*static void _OnOrderListKillfocusFnc(CWnd *pWnd){
	((CEMDepartmentUsageinDetail *)pWnd)->OnOrderListKillfocus();
}*/
static long _OnOrderListLoadDataFnc(CWnd *pWnd){
	return ((CEMDepartmentUsageinDetail *)pWnd)->OnOrderListLoadData();
}
/*static void _OnOrderListAddNewFnc(CWnd *pWnd){
	((CEMDepartmentUsageinDetail *)pWnd)->OnOrderListAddNew();
}*/
static void _OnGroupbyDeptSelectFnc(CWnd *pWnd){
	 ((CEMDepartmentUsageinDetail*)pWnd)->OnGroupbyDeptSelect();
}
static void _OnDrugSelectFnc(CWnd *pWnd){
	  ((CEMDepartmentUsageinDetail*)pWnd)->OnDrugSelect();
}
static void _OnMaterialSelectFnc(CWnd *pWnd){
	  ((CEMDepartmentUsageinDetail*)pWnd)->OnMaterialSelect();
}
static void _OnAllSelectFnc(CWnd *pWnd){
	  ((CEMDepartmentUsageinDetail*)pWnd)->OnAllSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CEMDepartmentUsageinDetail *pVw = (CEMDepartmentUsageinDetail *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMDepartmentUsageinDetail *pVw = (CEMDepartmentUsageinDetail *)pWnd;
	pVw->OnExportSelect();
} 
CEMDepartmentUsageinDetail::CEMDepartmentUsageinDetail(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CEMDepartmentUsageinDetail::~CEMDepartmentUsageinDetail(){
}
void CEMDepartmentUsageinDetail::OnCreateComponents(){
	m_wndDepartmentUsageinDetail.Create(this, _T("Department Usage in Detail"), 5, 5, 575, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 105, 55);
	m_wndFromDate.Create(this,110, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndOrderList.SetCheckBox(true);
	m_wndOrderListLabel.Create(this, _T("Order List"), 10, 60, 105, 85);
	m_wndOrderList.Create(this,110, 60, 290, 85); 
	m_wndGroupbyDept.Create(this, _T("Group by Dept"), 5, 95, 110, 120);
	m_wndDrug.Create(this, _T("Drug"), 115, 95, 200, 120);
	m_wndMaterial.Create(this, _T("Material"), 205, 95, 290, 120);
	m_wndAll.Create(this, _T("All"), 295, 95, 380, 120);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 395, 95, 490, 120);
	m_wndExport.Create(this, _T("&Export"), 495, 95, 575, 120);

}
void CEMDepartmentUsageinDetail::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndOrderList.SetCheckValue(true);
	m_wndOrderList.LimitText(35);


	m_wndOrderList.InsertColumn(0, _T("ID"), CFMT_TEXT | CFMT_RIGHT, 80);
	m_wndOrderList.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CEMDepartmentUsageinDetail::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndOrderList.SetEvent(WE_SELENDOK, _OnOrderListSelendokFnc);
	//m_wndOrderList.SetEvent(WE_SETFOCUS, _OnOrderListSetfocusFnc);
	//m_wndOrderList.SetEvent(WE_KILLFOCUS, _OnOrderListKillfocusFnc);
	m_wndOrderList.SetEvent(WE_SELCHANGE, _OnOrderListSelectChangeFnc);
	m_wndOrderList.SetEvent(WE_LOADDATA, _OnOrderListLoadDataFnc);
	//m_wndOrderList.SetEvent(WE_ADDNEW, _OnOrderListAddNewFnc);
	m_wndGroupbyDept.SetEvent(WE_CLICK, _OnGroupbyDeptSelectFnc);
	m_wndDrug.SetEvent(WE_CLICK, _OnDrugSelectFnc);
	m_wndMaterial.SetEvent(WE_CLICK, _OnMaterialSelectFnc);
	m_wndAll.SetEvent(WE_CLICK, _OnAllSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);
	OnOrderListLoadData();
}
void CEMDepartmentUsageinDetail::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndOrderList.GetDlgCtrlID(), m_szOrderListKey);
	DDX_Check(pDX, m_wndGroupbyDept.GetDlgCtrlID(), m_bGroupbyDept);
	DDX_Radio(pDX, m_wndDrug.GetDlgCtrlID(), m_nDrug);
	DDX_Radio(pDX, m_wndMaterial.GetDlgCtrlID(), m_nMaterial);
	DDX_Radio(pDX, m_wndAll.GetDlgCtrlID(), m_nAll);

}
void CEMDepartmentUsageinDetail::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szOrderListKey.Empty();
	m_bGroupbyDept=FALSE;
	m_nDrug=1;
	m_nMaterial=1;
	m_nAll=0;
	m_szItemType = _T("'PM', 'MA'");

}
int CEMDepartmentUsageinDetail::SetMode(int nMode){
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
/*void CEMDepartmentUsageinDetail::OnFromDateChange(){
	
} */
/*void CEMDepartmentUsageinDetail::OnFromDateSetfocus(){
	
} */
/*void CEMDepartmentUsageinDetail::OnFromDateKillfocus(){
	
} */
int CEMDepartmentUsageinDetail::OnFromDateCheckValue(){
	OnOrderListLoadData();
	return 0;
} 
/*void CEMDepartmentUsageinDetail::OnToDateChange(){
	
} */
/*void CEMDepartmentUsageinDetail::OnToDateSetfocus(){
	
} */
/*void CEMDepartmentUsageinDetail::OnToDateKillfocus(){
	
} */
int CEMDepartmentUsageinDetail::OnToDateCheckValue(){
	OnOrderListLoadData();
	return 0;
} 
void CEMDepartmentUsageinDetail::OnOrderListSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMDepartmentUsageinDetail::OnOrderListSelendok(){
	 
}
/*void CEMDepartmentUsageinDetail::OnOrderListSetfocus(){
	
}*/
/*void CEMDepartmentUsageinDetail::OnOrderListKillfocus(){
	
}*/
long CEMDepartmentUsageinDetail::OnOrderListLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndOrderList.IsSearchKey() && !m_szOrderListKey.IsEmpty()){
		szWhere.Format(_T(" and id='%s' "), m_szOrderListKey);
	}
	szWhere.AppendFormat(_T(" AND mt_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), 
						m_szFromDate, m_szToDate);
	m_wndOrderList.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT mt_transaction_id as id, mt_orderno as name FROM m_transaction WHERE mt_doctype = 'CSO' AND mt_department_to_id = 'C1.3' %s ORDER BY id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndOrderList.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMDepartmentUsageinDetail::OnOrderListAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMDepartmentUsageinDetail::OnGroupbyDeptSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CEMDepartmentUsageinDetail::OnDrugSelect(){
	m_szItemType = _T("'PM'");
}
void CEMDepartmentUsageinDetail::OnMaterialSelect(){
	m_szItemType = _T("'MA'");	
}
void CEMDepartmentUsageinDetail::OnAllSelect(){
	m_szItemType = _T("'PM', 'MA'");
}
void CEMDepartmentUsageinDetail::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString szSQL, tmpStr, szNewType, szOldType, szTemp, szOldDept, szNewDept;
	CRecord rs(&pMF->m_db);
	CReportSection *rptHeader, *rptDetail;
	CStringArray arrDat;
	CString szCurDte;
	int nIdx = 1, j = 0, k = 0, nSel = 0, nCount = 0;
	long double nTtCost[7], nGrpCost[7], nTemp = 0;
	long double nTtQty[7], nGrpQty[7];
	if (!rpt.Init(_T("Reports\\HMS\\PMS_CHITIETTINHHINHSUDUNGTAIDONVI.RPT")))
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
		nGrpQty[i] = 0;
		nGrpCost[i] = 0;
		nTtCost[i] = 0;
		nTtQty[i] = 0;
	}
	//Header Section
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), 
	CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	//Detail
	arrDat.Add(_T("pname"));
	arrDat.Add(_T("product_uomname"));
	arrDat.Add(_T("price"));
	arrDat.Add(_T("solqty"));
	arrDat.Add(_T("solamt"));
	arrDat.Add(_T("polqty"));
	arrDat.Add(_T("polamt"));
	arrDat.Add(_T("solinsqty"));
	arrDat.Add(_T("solinsamt"));
	arrDat.Add(_T("othinsqty"));
	arrDat.Add(_T("othinsamt"));
	arrDat.Add(_T("serqty"));
	arrDat.Add(_T("seramt"));
	arrDat.Add(_T("othqty"));
	arrDat.Add(_T("othamt"));
	arrDat.Add(_T("ttlqty"));
	arrDat.Add(_T("ttlamt"));
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("cat"), szNewType);
		rs.GetValue(_T("deptid"), szNewDept);
		if (szNewType != szOldType)
		{
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			TranslateString(szNewType, tmpStr);
			rptDetail->SetValue(_T("GroupName"), tmpStr);
			szOldType = szNewType;
		}
		if (m_bGroupbyDept)
			if (szNewDept != szOldDept)
			{
				nTemp = 0;
				for (int i = 0; i < 7; i++)
				{
					nTemp += nGrpQty[i];
				}
				if (nTemp > 0)
				{
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
					for (int i = 0; i < 7; i++)
					{
						tmpStr.Format(_T("s%d"), (i+2)*2+1);
						szTemp.Format(_T("%f"), nGrpQty[i]);
						nTtQty[i] += nGrpQty[i];
						nGrpQty[i] = 0;
						rptDetail->SetValue(tmpStr, szTemp);
						tmpStr.Format(_T("s%d"), (i+3)*2);
						szTemp.Format(_T("%.2f"), nGrpCost[i]);
						nTtCost[i] += nGrpCost[i];
						nGrpCost[i] = 0;
						rptDetail->SetValue(tmpStr, szTemp);
					}
				}
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				rptDetail->SetValue(_T("GroupName"), pMF->GetDepartmentName(szNewDept));
				szOldDept = szNewDept;
				nIdx = 1;
			}
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));	
		//i: 0..16
		//j: 2..18
		//k: 0..7
		k = 0;
		for (int i = 0; i < arrDat.GetCount(); i++)
		{
			j = i + 2;
			rs.GetValue(arrDat.GetAt(i), tmpStr);
			if (tmpStr != _T("0"))
				rptDetail->SetValue(int2str(j), tmpStr);
			if ((j%2 == 1) && (j >= 5))
			{
				if (m_bGroupbyDept)
					nGrpQty[k] += str2double(tmpStr);
				else
					nTtQty[k] += str2double(tmpStr);
			}
			if ((j%2 == 0) && (j >= 6))
			{
				if (m_bGroupbyDept)
					nGrpCost[k] += str2double(tmpStr);
				else
					nTtCost[k] += str2double(tmpStr);
				k++;
			}
		}
		rs.MoveNext();
	}
	if (m_bGroupbyDept)
	{
		nTemp = 0;
		for (int i = 0; i < 7; i++)
		{
			nTemp += nGrpQty[i];
		}
		if (nTemp > 0)
		{
			rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
			for (int i = 0; i < 7; i++)
			{
				tmpStr.Format(_T("s%d"), (i+2)*2+1);
				szTemp.Format(_T("%f"), nGrpQty[i]);
				nTtQty[i] += nGrpQty[i];
				nGrpQty[i] = 0;
				rptDetail->SetValue(tmpStr, szTemp);
				tmpStr.Format(_T("s%d"), (i+3)*2);
				szTemp.Format(_T("%.2f"), nGrpCost[i]);
				nTtCost[i] += nGrpCost[i];
				nGrpCost[i] = 0;
				rptDetail->SetValue(tmpStr, szTemp);
			}
		}
	}
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	for (int i = 0; i < 7; i++)
	{
		tmpStr.Format(_T("s%d"), (i+2)*2+1);
		szTemp.Format(_T("%f"), nTtQty[i]);
		rptDetail->SetValue(tmpStr, szTemp);
		tmpStr.Format(_T("s%d"), (i+3)*2);
		szTemp.Format(_T("%.2f"), nTtCost[i]);
		rptDetail->SetValue(tmpStr, szTemp);
	}
	szCurDte = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szCurDte.Right(2), szCurDte.Mid(5, 2), szCurDte.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
	
} 
void CEMDepartmentUsageinDetail::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr, szNewType, szOldType, szNewDept, szOldDept, szTemp;
	CRecord rs(&pMF->m_db);
	int nIdx = 1, j = 0, k = 0;
	long double nTtCost[7], nGrpCost[7], nTemp = 0;
	long double nTtQty[7], nGrpQty[7];
	CStringArray arrDat;
	CellFormat hf(&xls);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING);
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONASTERISK | MB_OK);
		return;
	}
	for (int i = 0; i < 7; i++)
	{
		nGrpQty[i] = 0;
		nGrpCost[i] = 0;
		nTtQty[i] = 0;
		nTtCost[i] = 0;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	//Header Section
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 30);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 17);
	xls.SetCellText(nCol, nRow + 3, _T("\x43HI TI\x1EBET S\x1EEC \x44\x1EE4NG THU\x1ED0\x43 T\x1EA0I \x110\x1A0N V\x1ECA"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 17);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	nRow = 5;
	xls.SetCellMergedRows(nCol, nRow, 2);
	xls.SetCellMergedRows(nCol+1, nRow, 2);
	xls.SetCellMergedRows(nCol+2, nRow, 2);
	xls.SetCellMergedRows(nCol+3, nRow, 2);
	xls.SetCellMergedColumns(nCol+4, nRow, 2);
	xls.SetCellMergedColumns(nCol+6, nRow, 2);
	xls.SetCellMergedColumns(nCol+8, nRow, 2);
	xls.SetCellMergedColumns(nCol+10, nRow, 2);
	xls.SetCellMergedColumns(nCol+12, nRow, 2);
	xls.SetCellMergedColumns(nCol+14, nRow, 2);
	xls.SetCellMergedColumns(nCol+16, nRow, 2);

	xls.SetCellText(nCol, nRow, _T("STT"), &hf);
	xls.SetCellText(nCol+1, nRow, _T("T\xEAn thu\x1ED1\x63, h\xE0m l\x1B0\x1EE3ng"));
	xls.SetCellText(nCol+2, nRow, _T("\x110VT"));
	xls.SetCellText(nCol+3, nRow, _T("\x110\x1A1n gi\xE1"));
	xls.SetCellText(nCol+4, nRow, _T("\x42\x1ED9 \x111\x1ED9i"));
	xls.SetCellText(nCol+4, nRow+1, _T("SL"));
	xls.SetCellText(nCol+5, nRow+1, _T("Th\xE0nh ti\x1EC1n"));
	xls.SetCellText(nCol+6, nRow, _T("\x43h\xEDnh s\xE1\x63h"));
	xls.SetCellText(nCol+6, nRow+1, _T("SL"));
	xls.SetCellText(nCol+7, nRow+1, _T("Th\xE0nh ti\x1EC1n"));
	xls.SetCellText(nCol+8, nRow, _T("\x42H Qu\xE2n"));
	xls.SetCellText(nCol+8, nRow+1, _T("SL"));
	xls.SetCellText(nCol+9, nRow+1, _T("Th\xE0nh ti\x1EC1n"));
	xls.SetCellText(nCol+10, nRow, _T("\x42H Kh\xE1\x63"));
	xls.SetCellText(nCol+10, nRow+1, _T("SL"));
	xls.SetCellText(nCol+11, nRow+1, _T("Th\xE0nh ti\x1EC1n"));
	xls.SetCellText(nCol+12, nRow, _T("\x44\x1ECB\x63h v\x1EE5"));
	xls.SetCellText(nCol+12, nRow+1, _T("SL"));
	xls.SetCellText(nCol+13, nRow+1, _T("Th\xE0nh ti\x1EC1n"));
	xls.SetCellText(nCol+14, nRow, _T("\x110T kh\xE1\x63"));
	xls.SetCellText(nCol+14, nRow+1, _T("SL"));
	xls.SetCellText(nCol+15, nRow+1, _T("Th\xE0nh ti\x1EC1n"));
	xls.SetCellText(nCol+16, nRow, _T("T\x1ED5ng \x63\x1ED9ng"));
	xls.SetCellText(nCol+16, nRow+1, _T("SL"));
	xls.SetCellText(nCol+17, nRow+1, _T("Th\xE0nh ti\x1EC1n"));

	nRow = 7;
	arrDat.Add(_T("pname"));
	arrDat.Add(_T("product_uomname"));
	arrDat.Add(_T("price"));
	arrDat.Add(_T("solqty"));
	arrDat.Add(_T("solamt"));
	arrDat.Add(_T("polqty"));
	arrDat.Add(_T("polamt"));
	arrDat.Add(_T("solinsqty"));
	arrDat.Add(_T("solinsamt"));
	arrDat.Add(_T("othinsqty"));
	arrDat.Add(_T("othinsamt"));
	arrDat.Add(_T("serqty"));
	arrDat.Add(_T("seramt"));
	arrDat.Add(_T("othqty"));
	arrDat.Add(_T("othamt"));
	arrDat.Add(_T("ttlqty"));
	arrDat.Add(_T("ttlamt"));
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("cat"), szNewType);
		rs.GetValue(_T("deptid"), szNewDept);
		if (szNewType != szOldType)
		{
			xls.SetCellMergedColumns(nCol, nRow, 9);
			TranslateString(szNewType, tmpStr);
			xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT, true);
			szOldType = szNewType;
			nRow++;
		}
		if (m_bGroupbyDept)
			if (szNewDept != szOldDept)
			{
				nTemp = 0;
				for (int i = 0; i < 7; i++)
				{
					nTemp += nGrpQty[i];
				}
				if (nTemp > 0)
				{
					xls.SetCellText(3, nRow, _T("T\x1ED5ng nh\xF3m:"), FMT_TEXT | FMT_CENTER, true);
					for (int i = 0; i < 7; i++)
					{
						tmpStr.Format(_T("%d"), (i + 2)*2);
						szTemp.Format(_T("%f"), nGrpQty[i]);
						nTtQty[i] += nGrpQty[i];
						nGrpQty[i] = 0;
						xls.SetCellText(str2int(tmpStr), nRow, szTemp, FMT_NUMBER1);
						tmpStr.Format(_T("%d"), (i + 2)*2+1);
						szTemp.Format(_T("%.2f"), nGrpCost[i]);
						nTtCost[i] += nGrpCost[i];
						nGrpCost[i] = 0;
						xls.SetCellText(str2int(tmpStr), nRow, szTemp, FMT_NUMBER1);
					}
					nRow++;
				}
				xls.SetCellMergedColumns(0, nRow, 10);
				xls.SetCellText(0, nRow, pMF->GetDepartmentName(szNewDept), FMT_TEXT, true);
				szOldDept = szNewDept;
				nIdx = 1;
				nRow++;
			}
		xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, rs.GetValue(_T("pname")), FMT_TEXT);
		xls.SetCellText(nCol+2, nRow, rs.GetValue(_T("product_uomname")), FMT_TEXT);
		k = 0;
		for (int i = 2; i < arrDat.GetCount(); i++)
		{
			j = i + 1;
			xls.SetCellText(nCol+j, nRow, rs.GetValue(arrDat.GetAt(i)), FMT_NUMBER1);
			if ((j%2 == 0) && (j >=4))
			{
				if (m_bGroupbyDept)
					nGrpQty[k] += str2double(rs.GetValue(arrDat.GetAt(i)));
				else
					nTtQty[k] += str2double(rs.GetValue(arrDat.GetAt(i)));
			}
			if ((j%2 == 1) && (j >=5))
			{
				if (m_bGroupbyDept)
					nGrpCost[k] += str2double(rs.GetValue(arrDat.GetAt(i)));
				else
					nTtCost[k] += str2double(rs.GetValue(arrDat.GetAt(i)));
				k++;
			}
		}
		nRow++;
		rs.MoveNext();
	}
	if (m_bGroupbyDept)
	{
		nTemp = 0;
		for (int i = 0; i < 7; i++)
		{
			nTemp += nGrpQty[i];
		}
		if (nTemp > 0)
		{
			xls.SetCellText(3, nRow, _T("T\x1ED5ng nh\xF3m:"), FMT_TEXT | FMT_CENTER, true);
			for (int i = 0; i < 7; i++)
			{
				tmpStr.Format(_T("%d"), (i + 2)*2);
				szTemp.Format(_T("%f"), nGrpQty[i]);
				nTtQty[i] += nGrpQty[i];
				nGrpQty[i] = 0;
				xls.SetCellText(str2int(tmpStr), nRow, szTemp, FMT_NUMBER1);
				tmpStr.Format(_T("%d"), (i + 2)*2+1);
				szTemp.Format(_T("%.2f"), nGrpCost[i]);
				nTtCost[i] += nGrpCost[i];
				nGrpCost[i] = 0;
				xls.SetCellText(str2int(tmpStr), nRow, szTemp, FMT_NUMBER1);
			}
			nRow++;
		}
	}
	xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_RIGHT, true);
	for (int i = 0; i < 7; i ++)
	{
		tmpStr.Format(_T("%f"), nTtQty[i]);
		xls.SetCellText((i + 2)*2, nRow, tmpStr, FMT_NUMBER1, true);
		tmpStr.Format(_T("%.2f"), nTtCost[i]);
		xls.SetCellText((i + 2)*2+1, nRow, tmpStr, FMT_NUMBER1, true);
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\CHITIETSUDUNGTHUOCTAIDONVI.xls"));
	
} 

CString CEMDepartmentUsageinDetail::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szWhere, szWhere1, szWhere2, szWhere3, szTemp, szDept, szStorageStr, szField, szGroupBy, szOrderBy, szOrders;
	szOrders.Empty();
	for (int i = 0; i < m_wndOrderList.GetItemCount(); i++)
	{
		if (m_wndOrderList.GetCheck(i))
		{
			m_wndOrderList.SetCurSel(i);
			if (!szOrders.IsEmpty())
				szOrders += _T(", ");
			szOrders += m_wndOrderList.GetCurrent(0);
		}
	}
	if (!szOrders.IsEmpty())
		szWhere2.Format(_T(" AND hpo_orderid IN (SELECT distinct hmt_orderid FROM hms_medical_transaction_view ") \
							_T(" WHERE hmt_reftransaction_id IN (%s))"), szOrders);
	szWhere.AppendFormat(_T(" AND product_org_id IN (%s)"), m_szItemType);
	//mtl_saleprice fixed price
	if (m_bGroupbyDept)
	{
		szField.Format(_T("deptid,"));
		szGroupBy.Format(_T("deptid,"));
		szOrderBy.Format(_T("deptid,"));
	}
	szSQL.Format(_T(" SELECT cat,") \
				_T("   product_id,") \
				_T("   product_name                     AS pname,") \
				_T("   %s") \
				_T("   product_uomname,") \
				_T("   Sum(solqty)                      AS solqty,") \
				_T("   Sum(solqty)*0    AS solamt,") \
				_T("   Sum(polqty)                      AS polqty,") \
				_T("   Sum(polqty)*0    AS polamt,") \
				_T("   Sum(solinsqty)                   AS solinsqty,") \
				_T("   Sum(solinsqty)*0 AS solinsamt,") \
				_T("   Sum(othinsqty)                   AS othinsqty,") \
				_T("   Sum(othinsqty)*0 AS othinsamt,") \
				_T("   Sum(serqty)                      AS serqty,") \
				_T("   Sum(serqty)*0    AS seramt,") \
				_T("   Sum(othqty)                      AS othqty,") \
				_T("   Sum(othqty)*0    AS othamt,") \
				_T("   Sum(solqty+polqty+solinsqty+othinsqty+serqty") \
				_T("       +othqty)                     AS ttlqty,") \
				_T("   Sum(solqty+polqty+solinsqty+othinsqty+serqty") \
				_T("       +othqty)*0   AS ttlamt") \
				_T(" FROM   (SELECT") \
				_T("   'Delivery' as cat,") \
				_T("   hpol_product_id,") \
				_T("   deptid,") \
				_T("   Sum(solqty)    AS solqty,") \
				_T("   Sum(polqty)    AS polqty,") \
				_T("   Sum(solinsqty) AS solinsqty,") \
				_T("   Sum(othinsqty) AS othinsqty,") \
				_T("   Sum(serqty)    AS serqty,") \
				_T("   Sum(otherqty)  AS othqty") \
				_T(" FROM   (SELECT") \
				_T("           hpol_product_id,") \
				_T("           hd_object,") \
				_T("		   hpo_deptid deptid,") \
				_T("           CASE WHEN hd_object=1 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS solqty,") \
				_T("           CASE WHEN hd_object=3 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS polqty,") \
				_T("           CASE WHEN hd_object=2 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS solinsqty,") \
				_T("           CASE WHEN hd_object=4 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS othinsqty,") \
				_T("           CASE WHEN hd_object=7 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS serqty,") \
				_T("           CASE WHEN hd_object IN (5, 6, 8, 9,") \
				_T("                              10, 11, 12) THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS otherqty") \
				_T("         FROM   hms_pharmaorder") \
				_T("                LEFT JOIN hms_doc ON (hd_docno=hpo_docno)") \
				_T("                LEFT JOIN hms_pharmaorderline ON (hpo_orderid=hpol_orderid)") \
				_T("         WHERE  hpo_orderstatus = 'A' %s) tbl") \
				_T(" GROUP  BY hpol_product_id, deptid)") \
				_T(" LEFT JOIN m_product_view ON (product_id=hpol_product_id)") \
				_T(" WHERE 1=1 %s") \
				_T(" GROUP  BY cat,") \
				_T("		   product_id,") \
				_T("           product_name,") \
				_T("           product_saleprice2,%s") \
				_T("		   product_uomname") \
				_T(" ORDER  BY cat, %s product_name "), szField, szWhere2, szWhere, szGroupBy, szOrderBy);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}