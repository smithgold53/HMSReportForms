#include "stdafx.h"
#include "FMPostedPaymentVoucherList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMPostedPaymentVoucherList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMPostedPaymentVoucherList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMPostedPaymentVoucherList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMPostedPaymentVoucherList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMPostedPaymentVoucherList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMPostedPaymentVoucherList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMPostedPaymentVoucherList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMPostedPaymentVoucherList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMPostedPaymentVoucherList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMPostedPaymentVoucherList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMPostedPaymentVoucherList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMPostedPaymentVoucherList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMPostedPaymentVoucherList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMPostedPaymentVoucherList *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMPostedPaymentVoucherList *pVw = (CFMPostedPaymentVoucherList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMPostedPaymentVoucherList *pVw = (CFMPostedPaymentVoucherList *)pWnd;
	pVw->OnExportSelect();
} 
CFMPostedPaymentVoucherList::CFMPostedPaymentVoucherList(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CFMPostedPaymentVoucherList::~CFMPostedPaymentVoucherList(){
}
void CFMPostedPaymentVoucherList::OnCreateComponents(){
	m_wndPostedPaymentVoucherList.Create(this, _T("Posted Payment Voucher List"), 5, 5, 450, 90);
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
void CFMPostedPaymentVoucherList::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

}

void CFMPostedPaymentVoucherList::OnSetWindowEvents(){
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
void CFMPostedPaymentVoucherList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndService.GetDlgCtrlID(), m_bService);

}
void CFMPostedPaymentVoucherList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bService = FALSE;
	m_nIAmount = 0;
	m_nEAmount = 0;
	m_nHitechEAmount = 0;
}

int CFMPostedPaymentVoucherList::SetMode(int nMode){
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

/*void CFMPostedPaymentVoucherList::OnFromDateChange(){
	
} */
/*void CFMPostedPaymentVoucherList::OnFromDateSetfocus(){
	
} */
/*void CFMPostedPaymentVoucherList::OnFromDateKillfocus(){
	
} */
int CFMPostedPaymentVoucherList::OnFromDateCheckValue(){
	return 0;
}
 
/*void CFMPostedPaymentVoucherList::OnToDateChange(){
	
} */
/*void CFMPostedPaymentVoucherList::OnToDateSetfocus(){
	
} */
/*void CFMPostedPaymentVoucherList::OnToDateKillfocus(){
	
} */
int CFMPostedPaymentVoucherList::OnToDateCheckValue(){
	return 0;
}
 
void CFMPostedPaymentVoucherList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
void CFMPostedPaymentVoucherList::OnClerkSelendok(){
	 
}

/*void CFMPostedPaymentVoucherList::OnClerkSetfocus(){
	
}*/
/*void CFMPostedPaymentVoucherList::OnClerkKillfocus(){
	
}*/
long CFMPostedPaymentVoucherList::OnClerkLoadData(){
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

/*void CFMPostedPaymentVoucherList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMPostedPaymentVoucherList::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szPos, szSysDate, szTemp;
	CStringArray arrCol;
	CReportSection *rptDetail = NULL, *rptFooter = NULL;
	long double nTtl[6];
	double nCost = 0, nHitechIAmount = 0;
	int nIdx = 1;
	if (!rpt.Init(_T("Reports\\HMS\\HF_BANGKECHUNGTUCHI.RPT")))
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
	arrCol.Add(_T("hitechreturn"));
	arrCol.Add(_T("eamount"));
	arrCol.Add(_T("iamount"));
	arrCol.Add(_T("deposit"));
	arrCol.Add(_T("refund"));
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
	m_nHitechEAmount = nTtl[0];
	m_nEAmount = nTtl[1];
	m_nIAmount = nTtl[2];
	szSQL = GetQueryString1();
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{
		rs.GetValue(_T("iamount"), nHitechIAmount);
	}
	tmpStr.Empty();
	if (m_nHitechEAmount > 0)
	{
		FormatCurrency(m_nHitechEAmount, szTemp);
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
		FormatCurrency(m_nEAmount, szTemp);
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
 
void CFMPostedPaymentVoucherList::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	CStringArray arrField;
	double nCost = 0;
	long double nTtlCost[6];
	int nCol = 0, nRow = 0, nIdx = 1;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	for (int i = 0; i < 6; i++)
		nTtlCost[i] = 0;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 30);
	for (int i = 3; i < 9; i++)
		xls.SetColumnWidth(i, 15);
	arrField.Add(_T("staff"));
	arrField.Add(_T("invoiceno"));
	arrField.Add(_T("hitechreturn"));
	arrField.Add(_T("eamount"));
	arrField.Add(_T("iamount"));
	arrField.Add(_T("deposit"));
	arrField.Add(_T("refund"));
	arrField.Add(_T("totalamount"));
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 9);
	xls.SetCellMergedColumns(0, 3, 9);
	xls.SetCellText(0, 0, pMF->m_CompanyInfo.sc_name, 4098, true);
	xls.SetCellText(0, 1, _T("Ph\xF2ng t\xE0i \x63h\xEDnh"), 4098, true);
	xls.SetCellText(0, 2, _T("\x42\x1EA2NG K\xCA \x43H\x1EE8NG T\x1EEA GHI S\x1ED4 - CHI"), 4098, true);
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
	xls.SetCellText(3, 4, _T("\x43\xE1\x63 kho\x1EA3n \x63hi tr\x1EA3 m\xE1y \x43N\x43"), 528386, true);
	xls.SetCellText(4, 4, _T("\x43\xE1\x63 kho\x1EA3n \x63hi tr\x1EA3"), 528386, true);
	xls.SetCellText(5, 4, _T("\x43\xE1\x63 kho\x1EA3n ph\x1EA3i chi \x42N n\x1ED9i tr\xFA"), 528386, true);
	xls.SetCellText(5, 5, _T("\x43\x1ED9ng"), 528386, true);
	xls.SetCellText(6, 5, _T("Tr\xED\x63h t\x1EA1m g\x1EEDi"), 528386, true);
	xls.SetCellText(7, 5, _T("\x43hi tr\x1EA3"), 528386, true);
	xls.SetCellText(8, 4, _T("T\x1ED5ng chi"), 528386, true);
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
	xls.Save(_T("Exports\\Bang ke chung tu ghi so - chi.xls"));
}
 
CString CFMPostedPaymentVoucherList::GetQueryString(){
	CString szSQL, szWhere, szRefWhere, szWhere1;
	szWhere.AppendFormat(_T(" AND hfe_status = 'P' AND fac_invoicetype IN ('P')") \
						_T(" AND fac_posteddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')")
						, m_szFromDate, m_szToDate);
	szRefWhere.Format(_T(" AND rf.hfe_status = 'P' AND fac_invoicetype = 'P' ") \
					_T(" AND fac_posteddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), 
					m_szFromDate, m_szToDate);
	if (m_bService)
		szWhere1.AppendFormat(_T(" AND fac_cashbook_id = 'DV'"));
	else
		szWhere1.AppendFormat(_T(" AND fac_cashbook_id <> 'DV'"));

	if (!m_szClerkKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND fac_posted = 'Y' AND fac_postedby = '%s'"), m_szClerkKey);
		szRefWhere.AppendFormat(_T(" AND fac_posted = 'Y' AND fac_postedby = '%s'"), m_szClerkKey);
	}
	szSQL.Format(_T(" SELECT    get_username(hfe_staff) staff, ") \
				_T("		   fac_invoiceno invoiceno,") \
				_T("		   SUM(hitechreturn) hitechreturn,") \
				_T("           SUM(eamount) eamount, ") \
				_T("           SUM(iamount) iamount, ") \
				_T("           SUM(deposit) deposit, ") \
				_T("           SUM(refund)   AS refund, ") \
				_T("		   ROUND(SUM(hitechreturn+eamount+refund)) AS totalamount")
				_T(" FROM      (SELECT    fac_invoiceno, ") \
				_T("					  fac_cashbook_id,") \
				_T("					  rf.hfe_staff, ") \
				_T("                      rf.hfe_cash_id, ") \
				_T("                      SUM(CASE WHEN NVL(hfl_hitechmachine, 'X') = 'Y' THEN hfe_cost ELSE 0 END) hitechreturn, ") \
				_T("                      SUM(CASE WHEN NVL(hfl_hitechmachine, 'X') <> 'Y' THEN hfe_cost ELSE 0 END) eamount, ") \
				_T("                      0 iamount, ") \
				_T("                      0 deposit, ") \
				_T("                      0 refund") \
				_T("            FROM      hms_fee_refund rf ") \
				_T("			LEFT JOIN fa_cash ON (fac_cash_id = hfe_cash_id)") \
				_T("            LEFT JOIN hms_fee_refundline rfl ON ( rf.hfe_docno = rfl.hfe_docno ") \
				_T("                                                  AND rf.hfe_invoiceno = rfl.hfe_invoiceno) ") \
				_T("            LEFT JOIN hms_fee_list ON ( hfe_itemid = hfl_feeid ) ") \
				_T("            WHERE     rf.hfe_class IN ('E', 'X') ") \
				_T("			%s") \
				_T("            GROUP     BY fac_invoiceno, ") \
				_T("						 rf.hfe_staff, ") \
				_T("						 fac_cashbook_id,") \
				_T("                         rf.hfe_cash_id ") \
				_T("			UNION ALL") \
				_T("			SELECT fac_invoiceno, ") \
				_T("				   fac_cashbook_id,") \
				_T("				   rf.hfe_staff, ") \
				_T("                   rf.hfe_cash_id, ") \
				_T("				   0,") \
				_T("                   0 AS eamount, ") \
				_T("                   0 AS iamount, ") \
				_T("                   0 AS deposit, ") \
				_T("                   rf.hfe_amount refund ") \
				_T("            FROM   hms_fee_refund rf") \
				_T("			LEFT JOIN hms_fee_deposit d ON (rf.hfe_docno = d.hfe_docno AND rf.hfe_refidx = d.hfe_invoiceno)") \
				_T("			LEFT JOIN fa_cash ON (fac_cash_id = rf.hfe_cash_id)") \
				_T("            WHERE  rf.hfe_class IN ( 'A', 'I' ) AND (d.hfe_status = 'R' OR rf.hfe_type = 'E') ") \
				_T("			%s") \
				_T("			UNION ALL") \
				_T("			SELECT fac_invoiceno, ") \
				_T("				   fac_cashbook_id,") \
				_T("				   rf.hfe_staff, ") \
				_T("                   rf.hfe_cash_id, ") \
				_T("				   0,") \
				_T("                   0		  AS eamount, ") \
				_T("                   i.hfe_amount AS iamount, ") \
				_T("                   i.hfe_deposit AS deposit, ") \
				_T("                   rf.hfe_amount refund ") \
				_T("            FROM   hms_fee_refund rf") \
				_T("			LEFT JOIN hms_fee_invoice i ON (rf.hfe_docno = i.hfe_docno AND rf.hfe_refidx = i.hfe_invoiceno)") \
				_T("			LEFT JOIN fa_cash ON (fac_cash_id = rf.hfe_cash_id)") \
				_T("            WHERE  rf.hfe_class IN ( 'A', 'I' ) AND rf.hfe_type = 'F' AND i.hfe_refund > 0 ") \
				_T("			%s)") \
				_T(" WHERE 1=1 ") \
				_T(" %s") \
				_T(" GROUP     BY hfe_staff, fac_invoiceno ") \
				_T(" HAVING sum(hitechreturn+eamount+refund) > 0") \
				_T(" ORDER     BY fac_invoiceno "), szRefWhere, szRefWhere, szRefWhere, szWhere1);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}

CString CFMPostedPaymentVoucherList::GetQueryString1(){
	CString szSQL, szWhere, szWhere1;
	szWhere.Format(_T(" AND rf.hfe_status = 'P' AND fac_invoicetype = 'P' ") \
					_T(" AND fac_posteddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), 
					m_szFromDate, m_szToDate);
	if (m_bService)
		szWhere1.AppendFormat(_T(" AND fac_cashbook_id = 'DV'"));
	else
		szWhere1.AppendFormat(_T(" AND fac_cashbook_id <> 'DV'"));
	szSQL.Format(_T(" SELECT    SUM(CASE WHEN rf.hfe_class IN ( 'A', 'I' ) AND Nvl(hfl_hitechmachine, 'X') = 'Y' THEN rfl.hfe_cost ") \
				_T("               ELSE 0 ") \
				_T("               END) iamount ") \
				_T(" FROM      fa_cash ") \
				_T(" LEFT JOIN hms_fee_refund rf ON ( rf.hfe_cash_id = fac_cash_id ) ") \
				_T(" LEFT JOIN hms_fee_refundline rfl ON ( rf.hfe_docno = rfl.hfe_docno AND rf.hfe_invoiceno = rfl.hfe_invoiceno ) ") \
				_T(" LEFT JOIN hms_fee_list ON ( hfl_feeid = hfe_itemid ) ") \
				_T(" WHERE     1 = 1 ") \
				_T(" %s %s "), szWhere, szWhere1);

	return szSQL; 
}
