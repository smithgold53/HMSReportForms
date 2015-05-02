#include "stdafx.h"
#include "FMProcessedPrescriptionList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMProcessedPrescriptionList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMProcessedPrescriptionList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMProcessedPrescriptionList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMProcessedPrescriptionList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMProcessedPrescriptionList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMProcessedPrescriptionList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMProcessedPrescriptionList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMProcessedPrescriptionList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMProcessedPrescriptionList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMProcessedPrescriptionList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMProcessedPrescriptionList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMProcessedPrescriptionList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMProcessedPrescriptionList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMProcessedPrescriptionList *)pWnd)->OnClerkAddNew();
}*/
static void _OnDischargedDateSelectFnc(CWnd *pWnd){
	 ((CFMProcessedPrescriptionList*)pWnd)->OnDischargedDateSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMProcessedPrescriptionList *pVw = (CFMProcessedPrescriptionList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMProcessedPrescriptionList *pVw = (CFMProcessedPrescriptionList *)pWnd;
	pVw->OnExportSelect();
} 
CFMProcessedPrescriptionList::CFMProcessedPrescriptionList(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CFMProcessedPrescriptionList::~CFMProcessedPrescriptionList(){
}
void CFMProcessedPrescriptionList::OnCreateComponents(){
	m_wndProcessedPrescriptionList.Create(this, _T("Processed Prescription List"), 5, 5, 490, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 485, 85); 
	//m_wndDischargedDate.Create(this, _T("By Discharged Date"), 5, 95, 245, 120);
	//m_wndPrintPreview.Create(this, _T("&Print Preview"), 300, 95, 410, 120);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 380, 95, 490, 120);

}
void CFMProcessedPrescriptionList::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(65);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 270);

}

void CFMProcessedPrescriptionList::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndClerk.SetEvent(WE_SELENDOK, _OnClerkSelendokFnc);
	//m_wndClerk.SetEvent(WE_SETFOCUS, _OnClerkSetfocusFnc);
	//m_wndClerk.SetEvent(WE_KILLFOCUS, _OnClerkKillfocusFnc);
	m_wndClerk.SetEvent(WE_SELCHANGE, _OnClerkSelectChangeFnc);
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	//m_wndClerk.SetEvent(WE_ADDNEW, _OnClerkAddNewFnc);
	//m_wndDischargedDate.SetEvent(WE_CLICK, _OnDischargedDateSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);	

}
void CFMProcessedPrescriptionList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	//DDX_Check(pDX, m_wndDischargedDate.GetDlgCtrlID(), m_bDischargedDate);

}
void CFMProcessedPrescriptionList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bDischargedDate=FALSE;

}

int CFMProcessedPrescriptionList::SetMode(int nMode){
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

/*void CFMProcessedPrescriptionList::OnFromDateChange(){
	
} */
/*void CFMProcessedPrescriptionList::OnFromDateSetfocus(){
	
} */
/*void CFMProcessedPrescriptionList::OnFromDateKillfocus(){
	
} */
int CFMProcessedPrescriptionList::OnFromDateCheckValue(){
	return 0;
}
 
/*void CFMProcessedPrescriptionList::OnToDateChange(){
	
} */
/*void CFMProcessedPrescriptionList::OnToDateSetfocus(){
	
} */
/*void CFMProcessedPrescriptionList::OnToDateKillfocus(){
	
} */
int CFMProcessedPrescriptionList::OnToDateCheckValue(){
	return 0;
}
 
void CFMProcessedPrescriptionList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
void CFMProcessedPrescriptionList::OnClerkSelendok(){
	 
}

/*void CFMProcessedPrescriptionList::OnClerkSetfocus(){
	
}*/
/*void CFMProcessedPrescriptionList::OnClerkKillfocus(){
	
}*/
long CFMProcessedPrescriptionList::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	}
	m_wndClerk.DeleteAllItems(); 
	szSQL.Format(_T("SELECT su_userid as userid, su_name as name FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("userid")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}

/*void CFMProcessedPrescriptionList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMProcessedPrescriptionList::OnDischargedDateSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMProcessedPrescriptionList::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CReportSection *rptDetail;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szWhere, szWhere1, szDate;
	CString szOldTit, szNewTit;
	double nDrugAmt = 0, nMatAmt = 0;
	int nIdx = 1;
	if (!rpt.Init(_T("Reports\\HMS\\PMS_DANHSACHDUYETDONTHUOC.RPT")))
		return;
	szWhere.Format(_T(" AND mt_processeddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"),m_szFromDate, m_szToDate);
	szWhere1.Format(_T(" AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"),m_szFromDate, m_szToDate);
	if (!m_szClerkKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND mt_processedby = '%s'"), m_szClerkKey);
		szWhere1.AppendFormat(_T(" AND mt_approvedby= '%s'"), m_szClerkKey);
	}
	szSQL.Format(_T(" SELECT tit, appdate, ") \
				_T("        orderno, ") \
				_T("        deptname, ") \
				_T("        SUM(drugamt)     drugamt, ") \
				_T("        SUM(materialamt) materialamt ") \
				_T(" FROM   (SELECT    'Drug Daily' as tit, mtb_transaction_batch_id                 orderid, ") \
				_T("                   Trunc(mt_processeddate)                 appdate, ") \
				_T("                   mtb_orderno                              orderno, ") \
				_T("                   Get_departmentname(mtb_department_to_id) deptname, ") \
				_T("                   CASE ") \
				_T("                     WHEN product_org_id = 'PM' THEN mtl_saleprice * mtl_qtyorder ") \
				_T("                     ELSE 0 ") \
				_T("                   END                                      AS drugamt, ") \
				_T("                   CASE ") \
				_T("                     WHEN product_org_id = 'MA' THEN mtl_saleprice * mtl_qtyorder ") \
				_T("                     ELSE 0 ") \
				_T("                   END                                      AS materialamt ") \
				_T("         FROM      m_transaction_batch ") \
				_T("         LEFT JOIN m_transaction ON ( mt_orderno = mtb_orderno ) ") \
				_T("         LEFT JOIN m_transactionline ON ( mt_transaction_id = mtl_transaction_id ) ") \
				_T("         LEFT JOIN m_product_view ON ( product_id = mtl_product_id ) ") \
				_T("         WHERE     mt_processed = 'Y' %s") \
				_T("         UNION ALL ") \
				_T("         SELECT    'Return Order' as tit, mt_transaction_id                    orderid, ") \
				_T("                   Trunc(mt_approveddate)               appdate, ") \
				_T("                   mt_orderno                           orderno, ") \
				_T("                   Get_departmentname(mt_department_id) deptname, ") \
				_T("                   CASE ") \
				_T("                     WHEN product_org_id = 'PM' THEN hpol_unitprice * hpol_qtyreverse ") \
				_T("                     ELSE 0 ") \
				_T("                   END                                  AS drugamt, ") \
				_T("                   CASE ") \
				_T("                     WHEN product_org_id = 'MA' THEN hpol_unitprice * hpol_qtyreverse ") \
				_T("                     ELSE 0 ") \
				_T("                   END                                  AS materialamt ") \
				_T("         FROM      m_transaction ") \
				_T("         LEFT JOIN hms_ipharmaorderline ON ( mt_transaction_id = hpol_retorder_id) ") \
				_T("         LEFT JOIN m_product_view ON ( product_id = hpol_product_id ) ") \
				_T("         WHERE     mt_status = 'A' AND mt_doctype = 'DRO' AND mt_storage_to_id IN (14, 15) %s) ") \
				_T(" GROUP  BY tit, orderid, ") \
				_T("           appdate, ") \
				_T("           orderno, ") \
				_T("           deptname ") \
				_T(" ORDER  BY tit, appdate, ") \
				_T("           orderno "), szWhere, szWhere1);

	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("tit"), szNewTit);
		if (szNewTit != szOldTit)
		{
			if (nDrugAmt + nMatAmt > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				rptDetail->SetValue(_T("s7"), double2str(nDrugAmt));
				rptDetail->SetValue(_T("s8"), double2str(nMatAmt));
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			TranslateString(szNewTit, tmpStr);
			rptDetail->SetValue(_T("GroupName"), tmpStr);
			szOldTit = szNewTit;
			nIdx = 1;
			nDrugAmt = 0;
			nMatAmt = 0;
		}
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), CDate::Convert(rs.GetValue(_T("appdate")), yyyymmdd, ddmmyyyy));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("orderno")));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("deptname")));
		rs.GetValue(_T("drugamt"), tmpStr);
		nDrugAmt += str2double(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rs.GetValue(_T("materialamt"), tmpStr);
		nMatAmt += str2double(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.MoveNext();
	}
	if (nDrugAmt + nMatAmt > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
		rptDetail->SetValue(_T("s7"), double2str(nDrugAmt));
		rptDetail->SetValue(_T("s8"), double2str(nMatAmt));
	}
	szDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5, 2), szDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
	
}
 
void CFMProcessedPrescriptionList::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0, nIdx = 0;
	double nAmount = 0;
	long double nTotal = 0;

	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}

	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 25);
	xls.SetColumnWidth(3, 15);

	xls.SetRowHeight(4, 40);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);

	xls.SetCellMergedColumns(nCol + 3, nRow, 4);
	xls.SetCellMergedColumns(nCol + 3, nRow + 1, 4);

	xls.SetCellMergedColumns(nCol, nRow + 2, 4);
	xls.SetCellMergedColumns(nCol, nRow + 3, 4);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 3, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 3, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol, nRow + 2, _T("T\x1ED4NG H\x1EE2P TH\x45O NG\x1AF\x1EDCI THU"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("M\xE3"));
	arrCol.Add(_T("T\xEAn ng\x1B0\x1EDDi thu"));
	arrCol.Add(_T("S\x1ED1 ti\x1EC1n"));
	nRow = 4;

	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10); 
	}

	nRow = 5;

	while (!rs.IsEOF())
	{
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("userid"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("username"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("amount"), nAmount);
		tmpStr.Format(_T("%.2f"), nAmount);
		nTotal += nAmount;
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);
		nRow++;
		rs.MoveNext();
	}
	
	xls.SetCellText(nCol + 2, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 10);
	tmpStr.Format(_T("%.2Lf"), nTotal);
	xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
	xls.Save(_T("Exports\\Tong Hop Theo Nguoi Thu.xls"));
	
}
 
CString CFMProcessedPrescriptionList::GetQueryString(){
	CString szSQL;
	return szSQL;
}