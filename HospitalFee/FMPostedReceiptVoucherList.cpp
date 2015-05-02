#include "stdafx.h"
#include "FMPostedReceiptVoucherList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMPostedReceiptVoucherList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMPostedReceiptVoucherList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMPostedReceiptVoucherList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMPostedReceiptVoucherList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMPostedReceiptVoucherList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMPostedReceiptVoucherList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMPostedReceiptVoucherList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMPostedReceiptVoucherList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMPostedReceiptVoucherList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMPostedReceiptVoucherList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMPostedReceiptVoucherList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMPostedReceiptVoucherList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMPostedReceiptVoucherList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMPostedReceiptVoucherList *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMPostedReceiptVoucherList *pVw = (CFMPostedReceiptVoucherList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMPostedReceiptVoucherList *pVw = (CFMPostedReceiptVoucherList *)pWnd;
	pVw->OnExportSelect();
} 
CFMPostedReceiptVoucherList::CFMPostedReceiptVoucherList(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CFMPostedReceiptVoucherList::~CFMPostedReceiptVoucherList(){
}
void CFMPostedReceiptVoucherList::OnCreateComponents(){
	m_wndPostedReceiptVoucherList.Create(this, _T("Posted Receipt Voucher List"), 5, 5, 450, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 225, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 230, 30, 310, 55);
	m_wndToDate.Create(this,315, 30, 445, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 445, 85); 
	m_wndService.Create(this, _T("Service"), 5, 95, 85, 120);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 265, 95, 365, 120);
	m_wndExport.Create(this, _T("&Export"), 370, 95, 450, 120);

}
void CFMPostedReceiptVoucherList::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

}
void CFMPostedReceiptVoucherList::OnSetWindowEvents(){
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
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CFMPostedReceiptVoucherList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndService.GetDlgCtrlID(), m_bService);

}
void CFMPostedReceiptVoucherList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bService = FALSE;
	m_nEAmount = 0;
	m_nIAmount = 0;

}
int CFMPostedReceiptVoucherList::SetMode(int nMode){
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
/*void CFMPostedReceiptVoucherList::OnFromDateChange(){
	
} */
/*void CFMPostedReceiptVoucherList::OnFromDateSetfocus(){
	
} */
/*void CFMPostedReceiptVoucherList::OnFromDateKillfocus(){
	
} */
int CFMPostedReceiptVoucherList::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMPostedReceiptVoucherList::OnToDateChange(){
	
} */
/*void CFMPostedReceiptVoucherList::OnToDateSetfocus(){
	
} */
/*void CFMPostedReceiptVoucherList::OnToDateKillfocus(){
	
} */
int CFMPostedReceiptVoucherList::OnToDateCheckValue(){
	return 0;
} 
void CFMPostedReceiptVoucherList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMPostedReceiptVoucherList::OnClerkSelendok(){
	 
}
/*void CFMPostedReceiptVoucherList::OnClerkSetfocus(){
	
}*/
/*void CFMPostedReceiptVoucherList::OnClerkKillfocus(){
	
}*/
long CFMPostedReceiptVoucherList::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty()){
	 szWhere.Format(_T(" and su_userid='%s' "), m_szClerkKey);
	}
	m_wndClerk.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as id, su_name as name FROM sys_user WHERE su_deptid = 'PTC' ORDER BY su_userid "));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMPostedReceiptVoucherList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMPostedReceiptVoucherList::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szPos, szSysDate, szTemp;
	CStringArray arrCol;
	CReportSection *rptDetail = NULL, *rptFooter = NULL;
	long double nTtl[6];
	double nCost = 0, nHitechEAmount = 0, nHitechIAmount = 0;
	int nIdx = 1;
	if (!rpt.Init(_T("Reports\\HMS\\HF_BANGKECHUNGTUTHU.RPT")))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	for (int i = 0 ; i < 6; i++){
		nTtl[i] = 0;
	}
	arrCol.Add(_T("eamount"));
	arrCol.Add(_T("deposit"));
	arrCol.Add(_T("iamount"));
	arrCol.Add(_T("inpatient_deposit_paid"));
	arrCol.Add(_T("inpatient_payment"));
	arrCol.Add(_T("totalamount"));
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), 
		CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	if (m_bService)
		rpt.GetReportHeader()->SetValue(_T("Object"), _T("\x44\x1ECB\x63h v\x1EE5"));
	else
		rpt.GetReportHeader()->SetValue(_T("Object"), _T("\x42\x1EA3o hi\x1EC3m"));
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("staff")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("invoiceno")));
		for (int i = 0; i < arrCol.GetCount(); i++)
		{
			rs.GetValue(arrCol.GetAt(i), nCost);
			nTtl[i] += nCost;
			rptDetail->SetValue(int2str(i+4), double2str(nCost));
		}
		rs.MoveNext();
	}
	rptFooter = rpt.GetReportFooter();
	for (int i = 0; i < 6; i++)
	{
		szPos.Format(_T("s%d"), i+4);
		tmpStr.Format(_T("%f"), nTtl[i]);
		rptFooter->SetValue(szPos, tmpStr);
	}
	m_nEAmount = nTtl[0];
	m_nIAmount = nTtl[2];
	szSQL = GetQueryString1();
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{
		rs.GetValue(_T("eamount"), nHitechEAmount);
		rs.GetValue(_T("iamount"), nHitechIAmount);
	}
	tmpStr.Empty();
	if (nHitechEAmount > 0)
	{
		FormatCurrency(nHitechEAmount, szTemp);
		tmpStr.AppendFormat(_T("Ngo\x1EA1i tr\xFA: %s"), szTemp);
	}
	if (nHitechIAmount > 0)
	{
		FormatCurrency(nHitechIAmount, szTemp);
		tmpStr.AppendFormat(_T(" N\x1ED9i tr\xFA: %s"), szTemp);
	}
	if (!tmpStr.IsEmpty())
		rptFooter->SetValue(_T("cnc"), tmpStr);
	tmpStr.Empty();
	if (m_nEAmount > 0)
	{
		FormatCurrency(m_nEAmount - nHitechEAmount, szTemp);
		tmpStr.AppendFormat(_T("Ngo\x1EA1i tr\xFA: %s"), szTemp);
	}
	if (m_nIAmount > 0)
	{
		FormatCurrency(m_nIAmount - nHitechIAmount, szTemp);
		tmpStr.AppendFormat(_T(" N\x1ED9i tr\xFA: %s"), szTemp);
	}
	if (!tmpStr.IsEmpty())
		rptFooter->SetValue(_T("conlai"), tmpStr);
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), szSysDate.Mid(8, 2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rptFooter->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CFMPostedReceiptVoucherList::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0;
	double nCost = 0;
	long double nTtlCost[6];
	int nIdx = 1;
	CStringArray arrField;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	for (int i = 0; i < 6; i++)
		nTtlCost[i] = 0;
	arrField.Add(_T("staff"));
	arrField.Add(_T("invoiceno"));
	arrField.Add(_T("eamount"));
	arrField.Add(_T("deposit"));
	arrField.Add(_T("iamount"));
	arrField.Add(_T("inpatient_deposit_paid"));
	arrField.Add(_T("inpatient_payment"));
	arrField.Add(_T("totalamount"));
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 30);
	for (int i = 3; i < 9; i++)
		xls.SetColumnWidth(i, 15);
	//Header
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 9);
	xls.SetCellMergedColumns(0, 3, 9);
	xls.SetCellText(0, 0, pMF->m_CompanyInfo.sc_name, 4098, true);
	xls.SetCellText(0, 1, _T("Ph\xF2ng t\xE0i \x63h\xEDnh"), 4098, true);
	xls.SetCellText(0, 2, _T("\x42\x1EA2NG K\xCA \x43H\x1EE8NG T\x1EEA GHI S\x1ED4 - THU"), 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), m_szFromDate, m_szToDate);
	xls.SetCellText(0, 3, tmpStr, 4098);
	for (int i = 0; i < 9; i++)
	{
		if (i < 5 || i > 7)
			xls.SetCellMergedRows(nCol + i, 4, 2);
		else if (i == 5)
			xls.SetCellMergedColumns(nCol + i, 4, 3);
	}
	xls.SetCellText(0, 4, _T("STT"), 528386, true);
	xls.SetCellText(1, 4, _T("H\x1ECD v\xE0 t\xEAn nh\xE2n vi\xEAn"), 528386, true);
	xls.SetCellText(2, 4, _T("S\x1ED1 phi\x1EBFu"), 528386, true);
	xls.SetCellText(3, 4, _T("\x43\xE1\x63 kho\x1EA3n ph\x1EA3i thu \x42N ngo\x1EA1i tr\xFA"), 528386, true);
	xls.SetCellText(4, 4, _T("\x42N t\x1EA1m g\x1EEDi"), 528386, true);
	xls.SetCellText(5, 4, _T("\x43\xE1\x63 kho\x1EA3n ph\x1EA3i thu \x42N n\x1ED9i tr\xFA"), 528386, true);
	xls.SetCellText(5, 5, _T("\x43\x1ED9ng"), 528386, true);
	xls.SetCellText(6, 5, _T("Tr\xED\x63h t\x1EA1m g\x1EEDi"), 528386, true);
	xls.SetCellText(7, 5, _T("Thu th\xEAm"), 528386, true);
	xls.SetCellText(8, 4, _T("T\x1ED5ng thu"), 528386, true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		for (int i = 0; i < arrField.GetCount(); i++)
		{
			if (i < 2)
				xls.SetCellText(nCol + i + 1, nRow, rs.GetValue(arrField.GetAt(i)), FMT_TEXT);
			else
			{
				rs.GetValue(arrField.GetAt(i), nCost);
				nTtlCost[i-2] += nCost;
				xls.SetCellText(nCol + i + 1, nRow, double2str(nCost), FMT_NUMBER1);
			}
		}
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellText(2, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
	for (int i = 0 ; i < 6;i++)	
	{
		tmpStr.Format(_T("%f"), nTtlCost[i]);
		xls.SetCellText(nCol+i+3, nRow, tmpStr, FMT_NUMBER1);
	}	
	xls.Save(_T("Exports\\Bang ke chung tu ghi so - thu.xls"));
} 

CString CFMPostedReceiptVoucherList::GetQueryString(){
	CString szSQL, szWhere, szWhere1, szWhere2;
	szWhere.AppendFormat(_T(" AND hfe_status = 'P' AND fac_posted = 'Y'") \
						_T(" AND fac_posteddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
						_T(" AND fac_invoicetype = 'R'"), m_szFromDate, m_szToDate);
	szWhere1.AppendFormat(_T(" AND hfe_status IN ('P', 'R') AND fac_posted = 'Y'") \
						_T(" AND fac_posteddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
						_T(" AND fac_invoicetype = 'R'"), m_szFromDate, m_szToDate);
	if (m_bService)
		szWhere2.AppendFormat(_T(" AND fac_cashbook_id = 'DV'"));
	else
		szWhere2.AppendFormat(_T(" AND fac_cashbook_id <> 'DV'"));
	if (!m_szClerkKey.IsEmpty())
	{
		szWhere.AppendFormat(_T("  AND fac_postedby = '%s'"), m_szClerkKey);
		szWhere1.AppendFormat(_T(" AND fac_postedby = '%s'"), m_szClerkKey);
	}
	szSQL.Format(_T(" SELECT    get_username(hfe_staff) staff, ") \
				_T("		   fac_invoiceno invoiceno,") \
				_T("           SUM(patpaid)                               AS eamount, ") \
				_T("           SUM(deposit)                               AS deposit, ") \
				_T("           SUM(inpatient_cost)                        AS iamount, ") \
				_T("           SUM(inpatient_depositpaid)                 AS inpatient_deposit_paid, ") \
				_T("           SUM(inpatient_payment)                     AS inpatient_payment, ") \
				_T("           ROUND(SUM(patpaid + inpatient_payment + deposit)) AS totalamount ") \
				_T(" FROM      (SELECT fac_invoiceno, ") \
				_T("				   fac_cashbook_id,") \
				_T("				   hfe_staff, ") \
				_T("                   hfe_cash_id, ") \
				_T("                   hfe_patpaid AS patpaid, ") \
				_T("                   hfe_deposit AS deposit, ") \
				_T("                   0           AS inpatient_cost, ") \
				_T("                   0           inpatient_depositpaid, ") \
				_T("                   0           inpatient_payment ") \
				_T("            FROM   hms_fee_invoice ") \
				_T("			LEFT JOIN fa_cash ON (hfe_cash_id = fac_cash_id)") \
				_T("            WHERE  hfe_class IN ( 'E', 'X' ) ") \
				_T("			%s") \
				_T("            UNION ALL ") \
				_T("            SELECT fac_invoiceno, ") \
				_T("				   fac_cashbook_id,") \
				_T("				   hfe_staff, ") \
				_T("                   hfe_cash_id, ") \
				_T("                   0 patpaid, ") \
				_T("                   hfe_amount deposit, ") \
				_T("                   0 AS inpatient_cost, ") \
				_T("                   0 inpatient_depositpaid, ") \
				_T("                   0 inpatient_payment ") \
				_T("            FROM   hms_fee_deposit ") \
				_T("			LEFT JOIN fa_cash ON (fac_cash_id = hfe_cash_id)") \
				_T("            WHERE  hfe_class IN('I', 'A') ") \
				_T("			%s") \
				_T("            UNION ALL ") \
				_T("            SELECT fac_invoiceno, ") \
				_T("				   fac_cashbook_id,") \
				_T("				   hfe_staff, ") \
				_T("                   hfe_cash_id, ") \
				_T("                   0, ") \
				_T("                   0, ") \
				_T("                   hfe_amount  AS inpatient_cost, ") \
				_T("                   CASE WHEN hfe_deposit > hfe_patpaid THEN hfe_patpaid ") \
				_T("                   ELSE hfe_deposit ") \
				_T("                   END         AS inpatient_depositpaid, ") \
				_T("                   hfe_payment AS inpatient_payment ") \
				_T("            FROM   hms_fee_invoice ") \
				_T("			LEFT JOIN fa_cash ON (fac_cash_id = hfe_cash_id)") \
				_T("            WHERE  hfe_class IN( 'A', 'I' ) AND hfe_payment > 0 ") \
				_T("			%s)") \
				_T(" WHERE 1=1 ") \
				_T(" %s") \
				_T(" GROUP BY  hfe_staff, fac_invoiceno") \
				_T(" HAVING    SUM(patpaid + inpatient_payment + deposit) > 0 ORDER BY fac_invoiceno "), szWhere, szWhere1, szWhere, szWhere2);
	return szSQL;
}

CString CFMPostedReceiptVoucherList::GetQueryString1(){
	CString szSQL, szWhere, szWhere1;
	szWhere.AppendFormat(_T(" AND i.hfe_status = 'P' AND fac_posted = 'Y'") \
						_T(" AND fac_posteddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
						_T(" AND fac_invoicetype = 'R'"), m_szFromDate, m_szToDate);
	if (m_bService)
		szWhere1.AppendFormat(_T(" AND fac_cashbook_id = 'DV'"));
	else
		szWhere1.AppendFormat(_T(" AND fac_cashbook_id <> 'DV'"));
	szSQL.Format(_T(" SELECT ") \
	_T("		   SUM(CASE WHEN i.hfe_class IN ( 'E', 'S' ) AND Nvl(hfl_hitechmachine, 'X') = 'Y' THEN f.hfe_cost ") \
	_T("               ELSE 0 END) eamount, ") \
	_T("           SUM(CASE WHEN i.hfe_class IN ( 'A', 'I' ) AND Nvl(hfl_hitechmachine, 'X') = 'Y' THEN f.hfe_cost ") \
	_T("               ELSE 0 END) iamount ") \
	_T(" FROM      fa_cash ") \
	_T(" LEFT JOIN hms_fee_invoice i ON ( i.hfe_cash_id = fac_cash_id ) ") \
	_T(" LEFT JOIN hms_fee f ON ( f.hfe_docno = i.hfe_docno AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
	_T(" LEFT JOIN hms_fee_list ON ( hfl_feeid = hfe_itemid ) ") \
	_T(" WHERE 1 = 1 ") \
	_T(" %s %s "), szWhere, szWhere1);
	return szSQL;
}