#include "stdafx.h"
#include "PrintForms.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include "afxtempl.h"
#include "HMSMainFrame.h"
#include "StringToken.h"
const int m_nMaxCol = 20;

typedef struct tagDrugInfo
{
	CString szItemID;
	CString szName;
	int nIndex;
}COLDRUGINFO;


typedef struct tagFEEITEM{
	CString	szGroupID;
	CString szID;
	CString szName;
	CString szUnit;
	float	nQuantity;
	double	nPrice;
	double	nInsPrice;
	double  nCost;
	double  nInsPaid;
	double	nDiscount;
	double  nDiffPaid;
	double	nPatPaid;
	double	nPatDebt;
	double	nUnpaid;
	bool	bVisible;
}FEEITEM;


typedef struct tagRECEIPTINFO{
			CString deptid;
			CString admitdate;
			CString dischargedate;
			CString mainicd;
			CString maindisease;
}RECEIPTINFO;


CPrintForms::CPrintForms(void)
{

}

CPrintForms::~CPrintForms(void)
{

}

void CPrintForms::PrintCashSheet(CString szVoucherNo, CString szVoucherType){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString szSQL, tmpStr, szSysDate;
	CRecord rs(&pMF->m_db);
	CReportSection *rptDetail = NULL;
	if (szVoucherType == _T("R"))
		tmpStr.Format(_T("Reports\\FA\\CashReceipt.RPT"));
	else if (szVoucherType == _T("P"))
		tmpStr.Format(_T("Reports\\FA\\CashPayment.RPT"));
	if(!rpt.Init(tmpStr))
		return;
	szSysDate = pMF->GetSysDate();
	szSQL = GetQueryString(_T("INVOICE"), szVoucherNo);
	rs.ExecSQL(szSQL);
	//rptDetail->SetValue(_T("COMPANYNAME"), rs.GetValue(_T("org")));
	//rptDetail->SetValue(_T("DEPARTMENT"), rs.GetValue(_T("client")));
	rpt.GetReportHeader()->SetValue(_T("COMPANYNAME"), _T("COMPANY NAME"));
	rpt.GetReportHeader()->SetValue(_T("DEPARTMENT"), _T("DEPARTMENT"));
	if (!rs.IsEOF())
	{	
		rpt.GetReportHeader()->SetValue(_T("BookNo"), _T("..."));
		rpt.GetReportHeader()->SetValue(_T("SerialNo"), rs.GetValue(_T("invoiceno")));
		if (szVoucherType == _T("R"))
			rpt.GetReportHeader()->SetValue(_T("debitaccount"), rs.GetValue(_T("debit")));
		else if (szVoucherType == _T("P"))
			rpt.GetReportHeader()->SetValue(_T("account"), rs.GetValue(_T("credit")));
		rs.GetValue(_T("transactor"), tmpStr);
		tmpStr.MakeUpper();
		rpt.GetReportHeader()->SetValue(_T("PATIENTNAME"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Address"), _T("Address"));
		rpt.GetReportHeader()->SetValue(_T("Reason"), rs.GetValue(_T("reason")));

		rpt.GetReportHeader()->SetValue(_T("TotalAmount"), rs.GetValue(_T("totalamount")));
		CString szXyz;
		MoneyToString(rs.GetValue(_T("totalamount")), szXyz);
		rpt.GetReportHeader()->SetValue(_T("SumInWord"), szXyz);
		tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));  
		rpt.GetReportHeader()->SetValue(_T("PrintDate"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("SumInWord1"), szXyz);
		szXyz.Empty();
		while (!rs.IsEOF())
		{
			if (!szXyz.IsEmpty())
				szXyz.Append(_T(", "));
			if (szVoucherType == _T("R"))
				szXyz.Append(rs.GetValue(_T("credit")));
			else if (szVoucherType == _T("P"))
				szXyz.Append(rs.GetValue(_T("debit")));
			rs.MoveNext();
		}
		if (szVoucherType == _T("R"))
			rpt.GetReportHeader()->SetValue(_T("account"), szXyz);
		else if (szVoucherType == _T("P"))
			rpt.GetReportHeader()->SetValue(_T("debitaccount"), szXyz);

	}
	rpt.PrintPreview();
}

void CPrintForms::PrintSheetList(){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString szSQL, tmpStr, szSysDate = pMF->GetSysDate();
	CRecord rs(&pMF->m_db);
	CReportSection *rptDetail = NULL;
	if (m_szType == _T("R"))
		tmpStr.Format(_T("Reports\\FA\\ReceiptList.RPT"));
	else if (m_szType == _T("P"))
		tmpStr.Format(_T("Reports\\FA\\PaymentList.RPT"));
	if(!rpt.Init(tmpStr))
		return;	
	szSQL = GetQueryString(_T("SHEET_LIST"), m_szID);
	int nRes = rs.ExecSQL(szSQL);
	if (nRes <= 0)
	{
//		_msg(_T("%s"), szSQL);
		return;
	}
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate), yyyymmdd, ddmmyyyy);
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), CDate::Convert(rs.GetValue(_T("accdate")), yyyymmdd, ddmmyyyy));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("invno")));
		rptDetail->SetValue(_T("3"), CDate::Convert(rs.GetValue(_T("invdate")), yyyymmdd, ddmmyyyy));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("partner")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("transactor")));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("Reason")));
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("totalamt")));
		//rptDetail->SetValue(_T("8"), rs.GetValue(_T("accdate")));			
		rs.MoveNext();
	}
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}

void CPrintForms::PrintBankReceipts(CString szVoucherNo, CString szVoucherType){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString szSQL, tmpStr, szSysDate;
	CRecord rs(&pMF->m_db);
	CReportSection *rptDetail = NULL;
	if (szVoucherType == _T("R"))
		tmpStr.Format(_T("Reports\\FA\\CashReceipt.RPT"));
	else if (szVoucherType == _T("P"))
		tmpStr.Format(_T("Reports\\FA\\CashPayment.RPT"));
	if(!rpt.Init(tmpStr))
		return;	
	szSysDate = pMF->GetSysDate();
	rptDetail = rpt.AddDetail();
	szSQL = GetQueryString(_T("Bank"), szVoucherNo);
	rs.ExecSQL(szSQL);
	//Rare
	/*for (int i = 0; i < rs.GetFieldCount(); i++){
		CString szVn;
		szVn = rs.GetFieldName(i);
		_msg(_T("%s-%s"), szVn, rs.GetValue(szVn));
	}*/
	rptDetail->SetValue(_T("COMPANYNAME"), rs.GetValue(_T("org")));
	rptDetail->SetValue(_T("DEPARTMENT"), rs.GetValue(_T("client")));
	if (!rs.IsEOF())
	{	
		rptDetail->SetValue(_T("BookNo"), _T("..."));
		rptDetail->SetValue(_T("SerialNo"), rs.GetValue(_T("invoiceno")));
		if (szVoucherType == _T("R"))
			rptDetail->SetValue(_T("debitaccount"), rs.GetValue(_T("debit")));
		else if (szVoucherType == _T("P"))
			rptDetail->SetValue(_T("account"), rs.GetValue(_T("credit")));
		rptDetail->SetValue(_T("PATIENTNAME"), rs.GetValue(_T("transactor")));
		rptDetail->SetValue(_T("Address"), _T("Address"));
		rptDetail->SetValue(_T("Reason"), rs.GetValue(_T("reason")));

		rptDetail->SetValue(_T("TotalAmount"), rs.GetValue(_T("totalamount")));
		CString szXyz;
		MoneyToString(rs.GetValue(_T("totalamount")), szXyz);
		rptDetail->SetValue(_T("SumInWord"), szXyz);
		tmpStr.Format(rptDetail->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));  
		rptDetail->SetValue(_T("PrintDate"), tmpStr);
		rptDetail->SetValue(_T("SumInWord1"), szXyz);
		szXyz.Empty();
		while (!rs.IsEOF())
		{
			if (!szXyz.IsEmpty())
				szXyz.Append(_T(", "));
			if (szVoucherType == _T("R"))
				szXyz.Append(rs.GetValue(_T("credit")));
			else if (szVoucherType == _T("P"))
				szXyz.Append(rs.GetValue(_T("debit")));
			rs.MoveNext();
		}
		if (szVoucherType == _T("R"))
			rptDetail->SetValue(_T("account"), szXyz);
		else if (szVoucherType == _T("P"))
			rptDetail->SetValue(_T("debitaccount"), szXyz);

	}
	rpt.PrintPreview();
}
void CPrintForms::PrintVoucher(CString szVoucherNo, CString szVoucherType){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString szSQL, tmpStr, szAmt, szID;
	int nID=1;
	CRecord rs(&pMF->m_db);
	CReportSection *rptDetail = NULL;
	if (!rpt.Init(_T("Reports\\FA\\Voucher.RPT")))
		return;
	szSQL = GetQueryString(_T("CASH_VOUCHER"), szVoucherNo);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_OK);
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("CompanyName"), rs.GetValue(_T("org")));
	rpt.GetReportHeader()->SetValue(_T("Department"), rs.GetValue(_T("department")));
	rpt.GetReportHeader()->SetValue(_T("PatientName"), rs.GetValue(_T("obj")));
	rpt.GetReportHeader()->SetValue(_T("Address"), rs.GetValue(_T("Addr")));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), CDate::Convert(rs.GetValue(_T("invoicedate")), yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("No"), rs.GetValue(_T("invoiceno")));
	rs.GetValue(_T("amt"), szAmt);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		szID.Format(_T("%d"), nID++);
		rptDetail->SetValue(_T("1"), szID);
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("descr")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("debit")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("credit")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("base")));
		rs.GetValue(_T("tax"), tmpStr);
		if (!tmpStr.IsEmpty())
		{	
			rptDetail = rpt.AddDetail();
			szID.Format(_T("%d"), nID++);
			rptDetail->SetValue(_T("1"), szID);
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("desrc")));
			if (szVoucherType == _T("R"))
			{
				rptDetail->SetValue(_T("3"), rs.GetValue(_T("debit")));
				rptDetail->SetValue(_T("4"), rs.GetValue(_T("tax")));	
			}
			else if (szVoucherType == _T("P"))
			{
				rptDetail->SetValue(_T("3"), rs.GetValue(_T("tax")));
				rptDetail->SetValue(_T("4"), rs.GetValue(_T("credit")));			
			}
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("taxamt")));
		}
		rs.MoveNext();
	}
	CString szXyz;
	MoneyToString(szAmt, szXyz);
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), szXyz);
	rpt.PrintPreview();
}

void CPrintForms::PrintBankVoucher(CString szVoucherNo, CString szVoucherType){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString szSQL, tmpStr, szAmt, szID;
	int nID=1;
	CRecord rs(&pMF->m_db);
	CReportSection *rptDetail = NULL;
	if (!rpt.Init(_T("Reports\\FA\\Voucher.RPT")))
		return;

	szSQL = GetQueryString(_T("BANK_VOUCHER"), szVoucherNo);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_OK);
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("CompanyName"), rs.GetValue(_T("org")));
	rpt.GetReportHeader()->SetValue(_T("Department"), rs.GetValue(_T("department")));
	rpt.GetReportHeader()->SetValue(_T("PatientName"), rs.GetValue(_T("obj")));
	rpt.GetReportHeader()->SetValue(_T("Address"), rs.GetValue(_T("Addr")));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), CDate::Convert(rs.GetValue(_T("invoicedate")), yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("No"), rs.GetValue(_T("invoiceno")));
	rs.GetValue(_T("amt"), szAmt);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		szID.Format(_T("%d"), nID++);
		rptDetail->SetValue(_T("1"), szID);
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("descr")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("debit")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("credit")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("base")));
		rs.GetValue(_T("tax"), tmpStr);
		if (!tmpStr.IsEmpty())
		{	
			rptDetail = rpt.AddDetail();
			szID.Format(_T("%d"), nID++);
			rptDetail->SetValue(_T("1"), szID);
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("descr")));
			if (szVoucherType == _T("C"))
			{
				rptDetail->SetValue(_T("3"), rs.GetValue(_T("debit")));
				rptDetail->SetValue(_T("4"), rs.GetValue(_T("tax")));	
			}
			else if (szVoucherType == _T("D"))
			{
				rptDetail->SetValue(_T("3"), rs.GetValue(_T("tax")));
				rptDetail->SetValue(_T("4"), rs.GetValue(_T("credit")));			
			}
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("taxamt")));
		}
		rs.MoveNext();
	}
	CString szXyz;
	MoneyToString(szAmt, szXyz);
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), szXyz);
	rpt.PrintPreview();
}
void CPrintForms::PrintPaymentOrder(CString szInvoiceNo){

	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString szSQL, tmpStr, szAmt, szID;
	CRecord rs(&pMF->m_db);
	if (!rpt.Init(_T("Reports\\FA\\Uy Nhiem Chi.RPT")))
		return;
	szSQL = GetQueryString(_T("PAYMENT_ORDER"), szInvoiceNo);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_OK);
		return;
	}
	//Lien 1
	rpt.GetReportHeader()->SetValue(_T("SerialNo"), rs.GetValue(_T("SerialNo")));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), rs.GetValue(_T("PrintDate")));
	rpt.GetReportHeader()->SetValue(_T("Payer"), rs.GetValue(_T("Payer")));
	rpt.GetReportHeader()->SetValue(_T("AccNo"), rs.GetValue(_T("AccNo")));
	rpt.GetReportHeader()->SetValue(_T("Bank"), rs.GetValue(_T("Bank")));
	rpt.GetReportHeader()->SetValue(_T("Receiver"), rs.GetValue(_T("Receiver")));
	rpt.GetReportHeader()->SetValue(_T("AccNo0"), rs.GetValue(_T("AccNo0")));
	rpt.GetReportHeader()->SetValue(_T("Bank0"), rs.GetValue(_T("Bank0")));
	rpt.GetReportHeader()->SetValue(_T("TotalAmount"), rs.GetValue(_T("TotalAmount")));
	rpt.GetReportHeader()->SetValue(_T("SumInWord"), rs.GetValue(_T("SumInWord")));
	rpt.GetReportHeader()->SetValue(_T("Reason"), rs.GetValue(_T("Reason")));
	//Lien 2
	rpt.GetReportHeader()->SetValue(_T("SerialNo0"), rs.GetValue(_T("SerialNo")));
	rpt.GetReportHeader()->SetValue(_T("PrintDate0"), rs.GetValue(_T("PrintDate")));
	rpt.GetReportHeader()->SetValue(_T("Payer0"), rs.GetValue(_T("Payer")));
	rpt.GetReportHeader()->SetValue(_T("AccNo00"), rs.GetValue(_T("AccNo")));
	rpt.GetReportHeader()->SetValue(_T("Bank1"), rs.GetValue(_T("Bank")));
	rpt.GetReportHeader()->SetValue(_T("Receiver0"), rs.GetValue(_T("Receiver")));
	rpt.GetReportHeader()->SetValue(_T("AccNo01"), rs.GetValue(_T("AccNo0")));
	rpt.GetReportHeader()->SetValue(_T("Bank2"), rs.GetValue(_T("Bank0")));
	rpt.GetReportHeader()->SetValue(_T("TotalAmount0"), rs.GetValue(_T("TotalAmount")));
	rpt.GetReportHeader()->SetValue(_T("SumInWord0"), rs.GetValue(_T("SumInWord")));
	rpt.GetReportHeader()->SetValue(_T("Reason0"), rs.GetValue(_T("Reason")));
	
	rpt.PrintPreview();	

}

void CPrintForms::PrintCreditNote(CString szNoteID, CString szVoucherType){
	CGuiMainFrame* pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection *rptDetail;
	CString szReportName, szSQL;
	if (szVoucherType == _T("C"))
		szReportName.Format(_T("Reports\\FA\\CreditNote.RPT"));
	else
		szReportName.Format(_T("Reports\\FA\\DebitNote.RPT"));
	if (!rpt.Init(szReportName))
		return;
	szSQL = GetQueryString(_T("CREDIT_NOTE"), szNoteID);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
//		_msg(_T("%s"), szSQL);
		ShowMessageBox(_T("No Data"), MB_OK);
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("Funder"), rs.GetValue(_T("Funder")));
	rpt.GetReportHeader()->SetValue(_T("Address"), rs.GetValue(_T("Address")));
	rpt.GetReportHeader()->SetValue(_T("Reason"), rs.GetValue(_T("Reason")));
	rpt.GetReportHeader()->SetValue(_T("Currency"), rs.GetValue(_T("Currency")));
	rpt.GetReportHeader()->SetValue(_T("Amount"), rs.GetValue(_T("Amount")));
	rpt.GetReportHeader()->SetValue(_T("AmountInWord"), rs.GetValue(_T("AmountInWord")));
	rpt.GetReportHeader()->SetValue(_T("No"), rs.GetValue(_T("No")));
	rpt.GetReportHeader()->SetValue(_T("Date"), rs.GetValue(_T("dte")));
	rpt.GetReportHeader()->SetValue(_T("AccNo"), rs.GetValue(_T("AccNo")));
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), rs.GetValue(_T("c1")));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("c2")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("c3")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("c4")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("c5")));	
		rs.MoveNext();
	}
	rpt.PrintPreview();
}

void CPrintForms::PrintIESheet(CString szOrderNo, CString szOrderType, CString szOrderDate){
	CGuiMainFrame* pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection *rptDetail;
	CString szReportName, szSQL, tmpStr, szDate;
	bool bImport = false, bExport = false;
	double nTotal = 0, nAmt = 0;
	int nIdx = 1;
	if (szOrderType == _T("I"))
	{
		bImport = true;
		szReportName.Format(_T("Reports\\FA\\ImportOrder.RPT"));
		szSQL = GetQueryString(_T("PURCHASE_ORDER"), szOrderNo);
	}
	else
	{
		bExport = true;
		szReportName.Format(_T("Reports\\FA\\ExportOrder.RPT"));
		szSQL = GetQueryString(_T("EXPORT_ORDER"), _T(""));
	}
	if (!rpt.Init(szReportName))
		return;
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_OK);
		return;
	}
	//rpt.GetReportHeader()->SetValue(_T("CompanyName"), );
	//rpt.GetReportHeader()->SetValue(_T("Department"), );
	_debug(_T("%s"),szOrderDate);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("SheetDate")), szOrderDate.Right(2), szOrderDate.Mid(5, 2), szOrderDate.Left(4));
	rpt.GetReportHeader()->SetValue(_T("SheetDate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("SheetNo"), rs.GetValue(_T("ShtNo")));
	rpt.GetReportHeader()->SetValue(_T("DebitAccount"), rs.GetValue(_T("DebAcc")));
	rpt.GetReportHeader()->SetValue(_T("CreditAccount"), rs.GetValue(_T("CreAcc")));
	rpt.GetReportHeader()->SetValue(_T("SupplierName"), rs.GetValue(_T("SupNme")));
	rpt.GetReportHeader()->SetValue(_T("StockName"), rs.GetValue(_T("StkNme")));
	rpt.GetReportHeader()->SetValue(_T("Address"), rs.GetValue(_T("Ads")));
	if (bExport)
		rpt.GetReportHeader()->SetValue(_T("Unit"), rs.GetValue(_T("Unt")));
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		CString szFld, szCol;
		tmpStr.Format(_T("%d"), nIdx++);
		rptDetail->SetValue(_T("1"), tmpStr);
		for (int i = 1; i < 8; i++)
		{
			szCol.Format(_T("%d"), i+1);
			szFld.Format(_T("c%d"), i+1);
			rs.GetValue(szFld, tmpStr);
			if (i == 7)
			{
				rs.GetValue(szFld, nAmt);
				nTotal += nAmt;
			}
			rptDetail->SetValue(szCol, tmpStr);
		}
		rs.MoveNext();
	}
	tmpStr.Format(_T("%.2f"), nTotal);
	rpt.GetReportFooter()->SetValue(_T("SubTotal"), tmpStr);
	MoneyToString(ToString(nTotal), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	szDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5, 2), szDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}

void CPrintForms::PrintTestReport(){
	CGuiMainFrame* pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection *rptDetail;
	CReportSection *rptHeader;
	CString szReportName, szSQL, tmpStr, szMem, szTit, szDate;
	if (!rpt.Init(_T("Reports\\FA\\FA_TestReport.RPT")))
		return;
	//Step to undercover field name
	//for (int i = 0; i < 50; i++)
	//{
	//	_debug(_T("%s"), rpt.GetDetail()->GetItem(i)->GetName());
	//}
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("companyname"), _T("COMPNAY NAME"));
	rptHeader->SetValue(_T("department"), _T("DEPARTMENT"));
	tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), _T("05/08/2013"), _T("05/08/2013"));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	for (int i = 1; i <= 10; i++)
	{
		szMem.Format(_T("Member%d"), i);
		szTit.Format(_T("Title%d"), i);
		rptHeader->SetValue(szMem, szMem);
		rptHeader->SetValue(szTit, szTit);
	}
	rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
	rptDetail->SetValue(_T("GroupName"), _T("GroupName"));
	rptDetail = rpt.AddDetail();
	for (int i = 1; i <= 9 ; i++)
		rptDetail->SetValue(i, i);
	szDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5, 2), szDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}

void CPrintForms::PrintPurchaseList(){
	CGuiMainFrame* pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection *rptDetail;
	CReportSection *rptHeader;
	CString szReportName, szSQL, tmpStr, szDate;
	if (!rpt.Init(_T("Reports\\FA\\FA_PurchaseList.RPT")))
		return;
	//Step to undercover field name
	//for (int i = 0; i < 50; i++)
	//{
	//	_debug(_T("%s"), rpt.GetDetail()->GetItem(i)->GetName());
	//}
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("companyname"), _T("COMPNAY NAME"));
	rptHeader->SetValue(_T("department"), _T("DEPARTMENT"));
	tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), _T("05/08/2013"), _T("05/08/2013"));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
	rptDetail->SetValue(_T("GroupName"), _T("GroupName"));
	rptDetail = rpt.AddDetail();
	for (int i = 1; i <= 9 ; i++)
		rptDetail->SetValue(i, i);
	szDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5, 2), szDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}

void CPrintForms::PrintMaterialReport(){
	CGuiMainFrame* pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection *rptDetail;
	CReportSection *rptHeader;
	CString szReportName, szSQL, tmpStr, szMem, szTit, szDate;
	if (!rpt.Init(_T("Reports\\FA\\FA_MaterialReport.RPT")))
		return;
	//Step to undercover field name
	//for (int i = 0; i < 50; i++)
	//{
	//	_debug(_T("%s"), rpt.GetDetail()->GetItem(i)->GetName());
	//}
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("companyname"), _T("COMPNAY NAME"));
	rptHeader->SetValue(_T("department"), _T("DEPARTMENT"));
	tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), _T("05/08/2013"), _T("05/08/2013"));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	for (int i = 1; i <= 10; i++)
	{
		szMem.Format(_T("Member%d"), i);
		szTit.Format(_T("Title%d"), i);
		rptHeader->SetValue(szMem, szMem);
		rptHeader->SetValue(szTit, szTit);
	}
	rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
	rptDetail->SetValue(_T("GroupName"), _T("GroupName"));
	rptDetail = rpt.AddDetail();
	for (int i = 1; i <= 16 ; i++)
		rptDetail->SetValue(i, i);
	szDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5, 2), szDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}

void CPrintForms::PrintMaterialRemain(){
	CGuiMainFrame* pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection *rptDetail;
	CReportSection *rptHeader;
	CString szReportName, szSQL, tmpStr;
	if (!rpt.Init(_T("Reports\\FA\\FA_MaterialRemain.RPT")))
		return;
	//Step to undercover field name
	//for (int i = 0; i < 50; i++)
	//{
	//	_debug(_T("%s"), rpt.GetDetail()->GetItem(i)->GetName());
	//}
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("companyname"), _T("COMPANY NAME"));
	rptHeader->SetValue(_T("department"), _T("DEPARTMENT"));
	tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), _T("05/08/2013"), _T("05/08/2013"));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
	rptDetail->SetValue(_T("GroupName"), _T("GroupName"));
	rptDetail = rpt.AddDetail();
	for (int i = 1; i <= 6 ; i++)
		rptDetail->SetValue(i, i);
	rpt.PrintPreview();
}

void CPrintForms::PrintDistribution(){
	CGuiMainFrame* pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection *rptDetail;
	CReportSection *rptHeader;
	CString szReportName, szSQL, tmpStr, szDate;
	if (!rpt.Init(_T("Reports\\FA\\FA_Distribution.RPT")))
		return;
	//Step to undercover field name
	//for (int i = 0; i < 50; i++)
	//{
	//	_debug(_T("%s"), rpt.GetDetail()->GetItem(i)->GetName());
	//}
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("companyname"), _T("COMPANY NAME"));
	rptHeader->SetValue(_T("department"), _T("DEPARTMENT"));
	tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), _T("05/08/2013"), _T("05/08/2013"));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
	rptDetail->SetValue(_T("GroupName"), _T("GroupName"));
	rptDetail = rpt.AddDetail();
	for (int i = 1; i <= 8 ; i++)
		rptDetail->SetValue(i, i);
	szDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5, 2), szDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}

void CPrintForms::PrintPurchaseOrder(long nOrderID){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, szOrderDate;
	CString tmpStr, tmpStr2;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);

	//if(!pMF->CheckPermission(_T("01.05")))
	//{
	//	ShowMessageBox(_T("Permission denined"), MB_OK);
	//	return;
	//}

	
	szSQL.Format(_T(" SELECT po_orderno as orderno, ") \
				_T("	po_invoiceno as invoiceno,") \
				_T("	Get_Taxrate(po_tax_id)*100 as vat, ") \
				_T(" 	po_invoicedate as invoicedate,") \
				_T("	(po_approveddate) as importdate,") \
				_T("	(po_orderdate) as orderdate,") \
				_T(" 	po_status as status,") \
				_T(" 	GET_PARTNERNAME(po_partner_id) as suppliername,") \
				_T(" 	po_deliveryby as deliverby,") \
				_T(" 	po_receivedby as receiveby,") \
				_T(" 	po_debitacct as debitaccount,") \
				_T(" 	po_creditacct as creditaccount,") \
				_T("	GET_PRODUCTRESOURCE_DESC(po_resource_id) as original , ") \
				_T("	GET_USERNAME(po_createdby) as createdby, ") \
				_T("	po_description as description,") \
				_T("	po_reference as voucher,") \
				_T(" 	GET_STORAGENAME(po_storage_id) as storagename,	") \
				_T("	po_documentno as documentno ") \
				_T(" FROM purchase_order") \
				_T(" WHERE po_purchase_order_id=%ld "), nOrderID);
_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);
	
	if(rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"));
		return;
	}
	
	CReport rpt;
	CString szDate, szSysDate;

	rs.GetValue(_T("status"), tmpStr);
	if(tmpStr != _T("A"))
	{
		return;
	}

	if(!rpt.Init(_T("Reports/HMS/PM_PHIEUNHAPKHO.RPT")) )
	{
		return;
	}
	//Step to undercover field name
	//for (int i = 0; i < 50; i++)
	//{
	//	_debug(_T("%s"), rpt.GetReportHeader()->GetItem(i)->GetName());
	//}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	rpt.GetReportHeader()->SetValue(_T("Department"), pMF->GetDepartmentName(pMF->GetDepartmentID()));
	rs.GetValue(_T("orderdate"), szOrderDate);
	//szOrderDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),tmpStr.Mid(8, 2), tmpStr.Mid(5,2),tmpStr.Left(4));
	szDate = CDateTime::Convert(szOrderDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss);
	rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);

	rpt.GetReportHeader()->SetValue(_T("ID"), nOrderID);
	rs.GetValue(_T("Invoiceno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Orderno"), rs.GetValue(_T("orderno")));
	rs.GetValue(_T("invoicedate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	rs.GetValue(_T("Status"), tmpStr);
	if(tmpStr == _T("A")) 
		TranslateString(_T("Approved"), tmpStr);
	else
		TranslateString(_T("Openning"), tmpStr);

	rpt.GetReportHeader()->SetValue(_T("Status"), tmpStr);
	rs.GetValue(_T("suppliername"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("suppliername"), tmpStr);
	rs.GetValue(_T("deliverby"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("deliverby"), tmpStr);
	rs.GetValue(_T("ReceiveBy"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("receiveby"), tmpStr);
	rs.GetValue(_T("DebitAccount"), tmpStr);
	if(tmpStr == _T("0"))
		tmpStr.Empty();
	rpt.GetReportHeader()->SetValue(_T("DebitAccount"), tmpStr);
	rs.GetValue(_T("CreditAccount"), tmpStr);
	if(tmpStr == _T("0"))
		tmpStr.Empty();
	rpt.GetReportHeader()->SetValue(_T("CreditAccount"), tmpStr);
	rs.GetValue(_T("StorageName"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("StockName"), tmpStr);
	//rs.GetValue(_T("Importdate"), tmpStr);
	//rpt.GetReportHeader()->SetValue(_T("ImportDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	rs.GetValue(_T("voucher"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("voucher"), tmpStr);

	if (pMF->GetModuleID() == _T("MA"))
	{
		rs.GetValue(_T("documentno"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("ReferenceNo"), tmpStr);
	}

	rs.GetValue(_T("description"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Reason"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Source"), rs.GetValue(_T("original")));
	//Page Header
	//Report Detail

	int nCount = 0;
	int nVAT;
	double nTotalAmount = 0;
	double nTotalVatAmount = 0;
	CString szMoney, szExprDate, szItemType;
	rs.GetValue(_T("vat"), nVAT);
	
	tmpStr.Format(_T("%d %s"), nVAT, _T("%"));

	rpt.GetReportHeader()->SetValue(_T("VAT"), tmpStr);



	szSQL.Format(_T(" SELECT ") \
				_T(" product_id as id, ") \
				_T(" product_name as name,") \
				_T(" 	product_purchase_uomname as unit,") \
				_T(" 	pol_expdate as expdate,") \
				_T(" 	pol_lotno as lotno,") \
				_T("	product_countryname as madein,") \
				_T("	product_manufacturename as mfgname,") \
				_T(" 	pol_qtyorder as qty,") \
				_T(" 	pol_unitprice as price,") \
				_T("	product_groupname,") \
				_T(" 	pol_unitprice+pol_unitprice*pol_taxrate/100.0 as vatprice,") \
				_T("	(pol_unitprice+pol_unitprice*pol_taxrate/100.0)*pol_qtyorder vatamount,") \
				_T(" 	pol_qtyorder*pol_unitprice as amount ") \
				_T(" FROM purchase_orderline") \
				_T(" LEFT JOIN m_productitem_view") \
				_T(" on(product_item_id              =pol_product_item_id) ") \
				_T(" WHERE pol_purchase_order_id=%ld ORDER BY pol_line "), nOrderID);
	
	rs2.ExecSQL(szSQL);
	
	CReportSection* rptDetail; 
	while(!rs2.IsEOF()){ 
		rptDetail = rpt.AddDetail();
		rs2.GetValue(_T("expdate"), tmpStr);
		if(CDate::IsValid(tmpStr))
			szExprDate = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		else
			szExprDate.Empty();
		nTotalAmount += ToDouble(rs2.GetValue(_T("amount")));
		nTotalVatAmount += ToDouble(rs2.GetValue(_T("vatamt")));
		nCount++;	
		rs2.GetValue(_T("product_groupname"), tmpStr);
		if (szItemType.Find(tmpStr) == -1)
		{
			if (!szItemType.IsEmpty())
				szItemType += _T(", ");
			szItemType.AppendFormat(_T("%s"), tmpStr);
		}

		tmpStr.Format(_T("%d"), nCount);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs2.GetValue(_T("name"), tmpStr) ;
		rptDetail->SetValue(_T("2"), tmpStr);
		rs2.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rptDetail->SetValue(_T("4"), szExprDate);
		rs2.GetValue(_T("lotno"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		FormatCurrencyStr(rs2.GetValue(_T("qty")), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		FormatCurrencyStr(rs2.GetValue(_T("price")), tmpStr);
		rptDetail-> SetValue(_T("7"), tmpStr);
		FormatCurrencyStr(rs2.GetValue(_T("vatprice")), tmpStr);
		rptDetail-> SetValue(_T("8"), tmpStr);

		FormatCurrencyStr(rs2.GetValue(_T("amount")), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);
		rs2.GetValue(_T("madein"), tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);
		rs2.GetValue(_T("mfgname"), tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);
		FormatCurrencyStr(rs2.GetValue(_T("vatamount")), tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);
		rs2.MoveNext();
	}
	CStringToken token(szItemType);
	int nTokenSize = token.GetSize();
	switch (nTokenSize)
	{
		case 0:
			break;
		case 1:
			token.GetAt(0, tmpStr2);
			TranslateString(tmpStr2, tmpStr);
			break;
		case 6:
			tmpStr = _T("To\xE0n \x62\x1ED9");
			break;
		default:
			tmpStr = _T("\x110\x61 lo\x1EA1i");
			break;
	}
	rpt.GetReportHeader()->SetValue(_T("ItemType"), tmpStr);
//	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	double nVatAmt = (double)(nTotalAmount*nVAT)/100.0;
//	FormatCurrency(nTotalAmount, tmpStr);
//	rptDetail->SetValue(_T("s8"), tmpStr);
	//Page Footer
	//Report Footer
	rs.GetValue(_T("createdby"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("CreatedBy"), tmpStr);	
	
	tmpStr.Format(_T("%d"), nCount);
	rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
	FormatCurrency(nTotalAmount, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("SubTotal"), tmpStr);
	FormatCurrency(nVatAmt, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalVAT"), tmpStr);
	FormatCurrency(nTotalAmount+nVatAmt, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	FormatCurrency(nTotalVatAmount, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalPrice"), tmpStr);
	szMoney.Format(_T("%.2f"), nTotalAmount+nVatAmt);
	MoneyToString(szMoney, tmpStr);
#ifdef UNICODE
	if(!tmpStr.IsEmpty()
		)
		tmpStr += _T(" \x111\x1ED3ng.");
#else
	if(!tmpStr.IsEmpty())
		tmpStr += _T(" ®ång.");
#endif
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	//szSysDate = pMF->GetSysDate(); 
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szOrderDate.Mid(8, 2),szOrderDate.Mid(5,2),szOrderDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}

void CPrintForms::PrintCheckinPurchaseOrder(long nOrderID){
	
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CString szSQL;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CRecord mrs(&pMF->m_db);
	//CPMSMemberDialog dlg(pMF);
	//
	//if(dlg.DoModal() != IDOK)
	//	return;


	szSQL.Format(_T(" SELECT po_orderno as invoiceno, ") \
				_T(" GET_TAXRATE(po_tax_id)*100 as vat, ") \
				_T(" 	po_orderdate as invoicedate,") \
				_T(" 	po_status as status,") \
				_T(" 	GET_PARTNERNAME(po_partner_id) as suppliername,") \
				_T(" 	po_deliveryby as deliverby,") \
				_T(" 	po_receivedby as receiveby,") \
				_T(" 	po_debitacct as deibtaccount,") \
				_T(" 	po_creditacct as creditaccount,") \
				_T("	po_description as description,") \
				_T(" 	GET_STORAGENAME(po_storage_id) as stockname	") \
				_T(" FROM purchase_order") \
				_T(" WHERE po_purchase_order_id=%ld "), nOrderID);
	rs.ExecSQL(szSQL);
	
	if(rs.IsEOF())
		return;
	
	szSQL.Format(_T("SELECT * FROM m_member ORDER BY mm_id"));
	mrs.ExecSQL(szSQL);

	CReport rpt;
	CString tmpStr;
	CString szDate, szSysDate, szName;

	if(!rpt.Init(_T("Reports/HMS/PMS_CHECKINPURCHASEINVOICE.RPT")) )
		return;
	//Step to undercover field name
	//for (int i = 0; i < 50; i++)
	//{
	//	_debug(_T("%s"), rpt.GetReportHeader()->GetItem(i)->GetName());
	//}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), szDate);
	
	
	while(!mrs.IsEOF()){
		int n;
		mrs.GetValue(_T("mm_id"), n);
		szName.Format(_T("Member%d"), n);
		mrs.GetValue(_T("mm_name"), tmpStr);
		rpt.GetReportHeader()->SetValue(szName, tmpStr);
		szName.Format(_T("Title%d"), n);
		mrs.GetValue(_T("mm_title"), tmpStr);
		rpt.GetReportHeader()->SetValue(szName, tmpStr);
		mrs.MoveNext();
	}
	rpt.GetReportHeader()->SetValue(_T("ID"), nOrderID);
	rs.GetValue(_T("InvoiceNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);
	rs.GetValue(_T("Invoicedate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	rs.GetValue(_T("Status"), tmpStr);
	if(tmpStr == _T("A")) 
		TranslateString(_T("Approved"), tmpStr);
	else
		TranslateString(_T("Openning"), tmpStr);

	rpt.GetReportHeader()->SetValue(_T("Status"), tmpStr);

	rs.GetValue(_T("suppliername"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("suppliername"), tmpStr);
	rs.GetValue(_T("deliverby"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("deliverby"), tmpStr);
	rs.GetValue(_T("ReceiveBy"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("receiveby"), tmpStr);
	rs.GetValue(_T("DebitAccount"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DebitAccount"), tmpStr);
	rs.GetValue(_T("CreditAccount"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CreditAccount"), tmpStr);
	rs.GetValue(_T("StockName"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("StockName"), tmpStr);
	rs.GetValue(_T("description"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Comment"), tmpStr);
	//Page Header
	//Report Detail

	int nCount = 0;
	int nVAT;
	double nTotalAmount = 0;
	double nTotalVatAmount = 0;
	CString szMoney, szExprDate;

	rs.GetValue(_T("vat"), nVAT);

	szSQL.Format(_T(" SELECT ") \
				_T(" product_id as id, ") \
				_T(" GET_FEEGROUPNAME(product_producttype)  as drugtype, ") \
				_T(" product_name as name,") \
				_T(" 	product_purchase_uomname as unit,") \
				_T(" 	pol_expdate as expdate,") \
				_T(" 	pol_lotno as lotno,") \
				_T(" 	pol_qtyorder as qty,") \
				_T(" 	pol_unitprice as price,") \
				/*_T(" 	mpi_insprice as vatprice,") \*/
				_T(" 	pol_qtyorder*pol_unitprice as amount, ") \
				/*_T(" 	pol_order_qty*mpi_insprice as vatamt, ") \*/
				/*_T(" 	GET_COUNTRYNAME(mp_country_id) as orgcountry, ") \*/
				_T("	GET_MANUFACTURENAME(pol_manufacture_id) as mfgcountry ") \
				_T(" FROM purchase_orderline") \
				_T(" LEFT JOIN m_product_view") \
				_T(" on(product_id              =pol_product_id)") \
				_T(" WHERE pol_purchase_order_id=%ld ORDER BY pol_line "), nOrderID);

	rs2.ExecSQL(szSQL);
	if (rs2.IsEOF())
		return;
	
	CReportSection* rptDetail = rpt.GetDetail(0); 
	while(!rs2.IsEOF())
	{	
		rptDetail = rpt.AddDetail();
		rs2.GetValue(_T("expdate"), tmpStr);
		if(CDate::IsValid(tmpStr))
			szExprDate = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		else
			szExprDate.Empty();
		nTotalAmount += ToDouble(rs2.GetValue(_T("amount")));
		nTotalVatAmount += ToDouble(rs2.GetValue(_T("vatamt")));
		nCount++;	
		

		tmpStr.Format(_T("%d"), nCount);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs2.GetValue(_T("drugtype"), tmpStr) ;
		rptDetail->SetValue(_T("2"), tmpStr);
	//	rs2.GetValue(_T("id"), tmpStr);
		rs2.GetValue(_T("name"), tmpStr) ;
		rptDetail->SetValue(_T("3"), tmpStr);
		rs2.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		
		rptDetail->SetValue(_T("5"), szExprDate);		

		rs2.GetValue(_T("lotno"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		FormatCurrencyStr(rs2.GetValue(_T("qty")), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		FormatCurrencyStr(rs2.GetValue(_T("price")), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);
		FormatCurrencyStr(rs2.GetValue(_T("amount")), tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);
		rs2.GetValue(_T("orgcountry"), tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);
		FormatCurrencyStr(rs2.GetValue(_T("vatprice")), tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);
		FormatCurrencyStr(rs2.GetValue(_T("vatamt")), tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);
		rs2.GetValue(_T("mfgcountry"),tmpStr);
		rptDetail->SetValue(_T("14"),tmpStr);		
		rs2.MoveNext();
	}
	//Page Footer
	//Report Footer
	tmpStr.Format(_T("%d"), nCount);
	rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
	FormatCurrency(nTotalAmount, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("SubTotal"), tmpStr);
	double nVatAmt = (double)(nTotalAmount*nVAT)/100.0;
	FormatCurrency(nVatAmt, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalVAT"), tmpStr);
	FormatCurrency(nTotalAmount+nVatAmt, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	FormatCurrency(nTotalVatAmount, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("Total4Price"), tmpStr);
	szMoney.Format(_T("%.2f"), nTotalAmount+nVatAmt);
	MoneyToString(szMoney, tmpStr);
	if(!tmpStr.IsEmpty())
		tmpStr += _T(" \x111\x1ED3ng.");
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);

	mrs.MoveFirst();
	while(!mrs.IsEOF()){
		int n;
		mrs.GetValue(_T("mm_id"), n);
		szName.Format(_T("UserName%d"), n);
		mrs.GetValue(_T("mm_name"), tmpStr);
		rpt.GetReportFooter()->SetValue(szName, tmpStr);
		szName.Format(_T("Office%d"), n);
		mrs.GetValue(_T("mm_title"), tmpStr);
		rpt.GetReportFooter()->SetValue(szName, tmpStr);
		mrs.MoveNext();
	}

	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);


	rpt.PrintPreview();
}

void CPrintForms::PrintStorageMovementOrder(long nOrderID)
{
	CGuiMainFrame *pMF = (CGuiMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szOrderBy;
	CRecord rs(&pMF->m_db);
	szSQL.Format(
		_T(" SELECT  mt_transaction_id as invoiceid,") \
		_T("	mt_orderno as invoiceno,") \
		_T("  	mt_posteddate as issuedate ,") \
		_T("  	mt_orderdate as invoicedate,") \
		_T(" 	GET_STORAGENAME(mt_storage_id) as storagename,	") \
		_T("  	GET_STORAGENAME(mt_storage_to_id) as storagename_to,") \
		_T("  	GET_PARTNERNAME(mt_partner_id) as partnername,") \
		_T("  	GET_SYSSEL_DESC('pms_export_type', mt_exptype) as exptype,") \
		_T("  	GET_USERNAME(mt_deliveryby) as deliverer,") \
		_T("  	GET_USERNAME(mt_receiveby) as receiver, ") \
		_T("	mt_description as note, ") \
		_T("	mt_status as status, ") \
		_T("	mt_doctype as doctype ") \
		_T(" FROM m_transaction") \
		_T(" WHERE mt_transaction_id=%ld  "), nOrderID);

	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	
	CReport rpt;
	CString tmpStr, szAmount;
	CString szDate, szSysDate;
	CString szDocType;

	if(!rpt.Init(_T("Reports/HMS/PMS_STOCKTRANSFER.RPT")) )
		return;

	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	rs.GetValue(_T("InvoiceNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);

	rs.GetValue(_T("invoicedate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("IssueDate"),CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), szDate);
	rs.GetValue(_T("doctype"), szDocType);
	tmpStr.Empty();
	if(szDocType == _T("SMO") || szDocType == _T("SRO"))
		rs.GetValue(_T("storagename_to"), tmpStr);
	else if(szDocType == _T("SPO"))
		rs.GetValue(_T("partnername"), tmpStr);
	else if(szDocType == _T("SOO"))
		rs.GetValue(_T("exptype"), tmpStr);
	rs.GetValue(_T("storagename"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("StockName"), tmpStr);
	rs.GetValue(_T("storagename_to"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("StockName1"), tmpStr);
	rs.GetValue(_T("note"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Comment"), tmpStr);
	rs.GetValue(_T("deliverer"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("deliverer"), tmpStr);
	rs.GetValue(_T("receiver"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("receiver"), tmpStr);

	tmpStr.Empty();
	//pMF->GetStatusString(rs.GetValue(_T("status")), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Status"), tmpStr);

	//Page Header
	//Report Detail

	int nCount = 0;
	double dMoney = 0;
	double nTotalAmount = 0;
	CString szMoney, szExprDate;
	szSQL.Format(_T(" SELECT") \
	_T("   product_id                          AS id,") \
	_T("   product_name                        AS name,") \
	_T("   product_uomname                     AS unit,") \
	_T("   product_expdate                     AS expdate,") \
	_T("   product_lotno                       AS lotno,") \
	_T("   Sum(mtl_qtyorder)                   AS qty,") \
	_T("   product_taxprice                   AS price,") \
	_T("   Sum(mtl_qtyorder*product_taxprice) AS amount") \
	_T(" FROM   m_transactionline") \
	_T("        LEFT JOIN m_productitem_view ON(mtl_product_item_id=product_item_id)") \
	_T(" WHERE mtl_transaction_id = %ld") \
	_T(" GROUP  BY product_id,") \
	_T("           product_name,") \
	_T("           product_uomname,") \
	_T("           product_expdate,") \
	_T("           product_lotno,") \
	_T("           product_taxprice,") \
	_T("           mtl_line ORDER BY mtl_line "), nOrderID);


	rs.ExecSQL(szSQL);
	CReportSection* rptDetail = rpt.GetDetail(0); 
	while(!rs.IsEOF()){ 
		rptDetail = rpt.AddDetail();
		rs.GetValue(_T("expdate"), tmpStr);
		szExprDate = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);

		dMoney = ToDouble(rs.GetValue(_T("amount")));
		nTotalAmount += dMoney;
		szMoney.Format(_T("%.2f"), dMoney);
		nCount++;	
		tmpStr.Format(_T("%d"), nCount);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("name"), tmpStr) ;
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("id"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rptDetail->SetValue(_T("5"), szExprDate);	
		rs.GetValue(_T("lotno"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		FormatCurrencyStr(rs.GetValue(_T("qty")), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		FormatCurrencyStr(rs.GetValue(_T("price")), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		//thanhtien
		FormatCurrencyStr(rs.GetValue(_T("amount")), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);
		rs.MoveNext();
	}
	//Page Footer
	
	//Report Footer
	tmpStr.Format(_T("%d"), nCount);
	rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);	
	FormatCurrency(nTotalAmount, szAmount);
	rpt.GetReportFooter()->SetValue(_T("SubTotal"), szAmount);
	szAmount.Format(_T("%.2f"), nTotalAmount);
	MoneyToString(szAmount, tmpStr);	
#ifdef UNICODE
	tmpStr += _T(" \x111\x1ED3ng.");
#endif
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	rpt.PrintPreview();
}

void CPrintForms::PrintStorageExportOrder(long nOrderID)
{
	CGuiMainFrame *pMF = (CGuiMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szOrderBy;
	CRecord rs(&pMF->m_db);
	szSQL.Format(
		_T(" SELECT  mt_transaction_id as invoiceid,") \
		_T("	mt_orderno as invoiceno,") \
		_T("  	mt_posteddate as issuedate ,") \
		_T("  	mt_orderdate as invoicedate,") \
		_T(" 	GET_STORAGENAME(mt_storage_id) as storagename,	") \
		_T("  	GET_STORAGENAME(mt_storage_to_id) as storagename_to,") \
		_T("  	GET_DEPARTMENTNAME(mt_department_id) as department,") \
		_T("  	GET_SYSSEL_DESC('pms_export_type', mt_exptype) as exptype,") \
		_T("  	GET_USERNAME(mt_deliveryby) as deliverer,") \
		_T("  	GET_USERNAME(mt_receiveby) as receiver, ") \
		_T("	mt_description as note, ") \
		_T("	mt_status as status, ") \
		_T("	mt_doctype as doctype, ") \
		_T("	mt_documentno documentno") \
		_T(" FROM m_transaction") \
		_T(" WHERE mt_transaction_id=%ld  "), nOrderID);

	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	
	CReport rpt;
	CString tmpStr, szAmount;
	CString szDate, szSysDate;
	CString szDocType;

	if(!rpt.Init(_T("Reports/HMS/PMS_STORAGEEXPORT.RPT")) )
		return;

	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	rs.GetValue(_T("InvoiceNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);

	rs.GetValue(_T("invoicedate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("IssueDate"),CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), szDate);
	rs.GetValue(_T("doctype"), szDocType);
	tmpStr.Empty();

	rs.GetValue(_T("storagename"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("StockName"), tmpStr);
	rs.GetValue(_T("storagename_to"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("StockName1"), tmpStr);

	rs.GetValue(_T("Department"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);

	rs.GetValue(_T("note"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Comment"), tmpStr);
	rs.GetValue(_T("deliverer"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("deliverer"), tmpStr);
	rs.GetValue(_T("receiver"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("receiver"), tmpStr);
	
	tmpStr.Empty();
	//pMF->GetStatusString(rs.GetValue(_T("status")), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Status"), tmpStr);
	if (pMF->GetModuleID() == _T("MA"))
		rpt.GetReportHeader()->SetValue(_T("ReferenceNo"), rs.GetValue(_T("documentno")));
	//Page Header
	//Report Detail

	int nCount = 0;
	int nQty = 0;
	double dMoney = 0;
	long double nTotalAmount = 0;

	CString szMoney, szExprDate;
	szSQL.Format(_T(" SELECT product_id as id,") \
				_T(" mtl_xproduct_id as xproduct_id, ") \
				_T("        product_name AS name,") \
				_T("        product_uomname AS unit,") \
				_T("        product_expdate AS expdate,") \
				_T("        product_lotno AS lotno,") \
				_T("        SUM(mtl_qtyorder) AS qtyorder,") \
				_T("        SUM(mtl_qtydelivery) AS qtyissue,") \
				_T("        product_taxprice AS price,") \
				_T("        SUM(mtl_qtyorder*product_taxprice) AS orderamount,") \
				_T("        SUM(mtl_qtydelivery*product_taxprice) AS issueamount") \
				_T(" FROM m_transactionline") \
				_T(" LEFT JOIN m_productitem_view ON(mtl_product_item_id=product_item_id)") \
				_T(" WHERE mtl_transaction_id=%ld and mtl_qtydelivery > 0 ") \
				_T(" GROUP BY product_uomindex, product_id, product_name, product_uomname,") \
				_T("          product_expdate, product_lotno, mtl_xproduct_id, ") \
				_T("          product_taxprice, mtl_line ORDER BY product_uomindex, product_name, product_uomname "), nOrderID);

	//_msg(_T("%s"), szSQL);

	rs.ExecSQL(szSQL);
	CReportSection* rptDetail = rpt.GetDetail(0); 
	long nProductID, nXProductID;
	CString szName;
	CRecord rsx(&pMF->m_db);
	while (!rs.IsEOF())
	{ 
		rptDetail = rpt.AddDetail();

		nCount++;

		tmpStr.Format(_T("%d"), nCount);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("id"), nProductID);
		rs.GetValue(_T("xproduct_id"), nXProductID);
		szName.Empty();
		rs.GetValue(_T("name"), szName) ;
		if(nProductID != nXProductID)
		{
			CString szName2;
			tmpStr.Format(_T("SELECT mp_name FROM m_Product where mp_product_id = %ld"), nXProductID);
			rsx.ExecSQL(tmpStr);
			rsx.GetValue(_T("mp_name"), szName2);

			szName.AppendFormat(_T(" [%s]"), szName2);
		}

		rptDetail->SetValue(_T("2"), szName);

		rs.GetValue(_T("id"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("expdate"), tmpStr);
		szExprDate = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("5"), szExprDate);

		rs.GetValue(_T("lotno"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("qtyorder"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("qtyissue"), tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.GetValue(_T("price"), dMoney);
		FormatCurrency(dMoney, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("issueamount"), dMoney);
		nTotalAmount += dMoney;
		FormatCurrency(dMoney, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);


		rs.MoveNext();
	}
	//Page Footer
	
	//Report Footer
	tmpStr.Format(_T("%d"), nCount);
	rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);	
	FormatCurrency(nTotalAmount, szAmount);
	rpt.GetReportFooter()->SetValue(_T("SubTotal"), szAmount);
	szAmount.Format(_T("%.2f"), nTotalAmount);
	MoneyToString(szAmount, tmpStr);	
#ifdef UNICODE
	tmpStr += _T(" \x111\x1ED3ng.");
#endif
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);

	GetDoctorName(pMF->GetCurrentUser());

	rpt.PrintPreview();
}


void CPrintForms::PrintStorageExportOrder2(long nOrderID)
{
	CGuiMainFrame *pMF = (CGuiMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szOrderBy;
	CRecord rs(&pMF->m_db);
	szSQL.Format(_T(" SELECT  mt_transaction_id as invoiceid,") \
		_T("	mt_orderno as invoiceno,") \
		_T("  	mt_posteddate as issuedate ,") \
		_T("  	mt_orderdate as invoicedate,") \
		_T(" 	GET_STORAGENAME(mt_storage_id) as storagename,	") \
		_T("  	GET_STORAGENAME(mt_storage_to_id) as storagename_to,") \
		_T("  	GET_DEPARTMENTNAME(mt_department_id) as department,") \
		_T("  	GET_SYSSEL_DESC('pms_export_type', mt_exptype) as exptype,") \
		_T("  	GET_USERNAME(mt_deliveryby) as deliverer,") \
		_T("  	GET_USERNAME(mt_receiveby) as receiver, ") \
		_T("	mt_description as note, ") \
		_T("	mt_status as status, ") \
		_T("	mt_doctype as doctype, ") \
		_T("	mt_documentno documentno") \
		_T(" FROM m_transaction") \
		_T(" WHERE mt_transaction_id=%ld  "), nOrderID);

	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	
	CReport rpt;
	CString tmpStr, szAmount;
	CString szDate, szSysDate;
	CString szDocType;

	if(!rpt.Init(_T("Reports/HMS/PMS_STORAGEEXPORT.RPT")) )
		return;

	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	rs.GetValue(_T("InvoiceNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);

	rs.GetValue(_T("invoicedate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("IssueDate"),CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), szDate);
	rs.GetValue(_T("doctype"), szDocType);
	tmpStr.Empty();

	rs.GetValue(_T("storagename"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("StockName"), tmpStr);
	rs.GetValue(_T("storagename_to"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("StockName1"), tmpStr);

	rs.GetValue(_T("Department"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);

	rs.GetValue(_T("note"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Comment"), tmpStr);
	rs.GetValue(_T("deliverer"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("deliverer"), tmpStr);
	rs.GetValue(_T("receiver"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("receiver"), tmpStr);

	tmpStr.Empty();
	//pMF->GetStatusString(rs.GetValue(_T("status")), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Status"), tmpStr);
	if (pMF->GetModuleID() == _T("MA"))
		rpt.GetReportHeader()->SetValue(_T("ReferenceNo"), rs.GetValue(_T("documentno")));
	//Page Header
	//Report Detail

	int nCount = 0;
	int nQty = 0;
	double dMoney = 0;
	long double nTotalAmount = 0;

	CString szMoney, szExprDate;
	szSQL.Format(_T(" SELECT product_id AS id,") \
				_T("        product_name AS name,") \
				_T("        product_uomname AS unit,") \
				_T("        product_expdate AS expdate,") \
				_T("        product_lotno AS lotno,") \
				_T("        SUM(hpol_qtyissue) AS qtyorder,") \
				_T("        hpol_unitprice AS price,") \
				_T("        SUM(hpol_qtyissue*hpol_unitprice) AS amount") \
				_T(" FROM purchase_orderline2") \
				_T(" LEFT JOIN m_productitem_view ON(pol_product_item_id=product_item_id)") \
				_T(" LEFT JOIN hms_ipharmaorderline ON (hpol_orderid = pol_orderid AND hpol_product_id = pol_product_id)") \
				_T(" WHERE pol_transaction_id=%ld") \
				_T(" GROUP BY product_uomindex, product_id, product_name, product_uomname,") \
				_T("          product_expdate, product_lotno,") \
				_T("          hpol_unitprice ORDER BY product_uomindex, product_name, product_uomname"), nOrderID);


	rs.ExecSQL(szSQL);
	CReportSection* rptDetail = rpt.GetDetail(0); 
	while (!rs.IsEOF())
	{ 
		rptDetail = rpt.AddDetail();

		nCount++;

		tmpStr.Format(_T("%d"), nCount);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("name"), tmpStr) ;
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("id"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("expdate"), tmpStr);
		szExprDate = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("5"), szExprDate);

		rs.GetValue(_T("lotno"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("qtyorder"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

//		rs.GetValue(_T("qtyissue"), tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.GetValue(_T("price"), dMoney);
		FormatCurrency(dMoney, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("amount"), dMoney);
		nTotalAmount += dMoney;
		FormatCurrency(dMoney, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);


		rs.MoveNext();
	}
	//Page Footer
	
	//Report Footer
	tmpStr.Format(_T("%d"), nCount);
	rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);	
	FormatCurrency(nTotalAmount, szAmount);
	rpt.GetReportFooter()->SetValue(_T("SubTotal"), szAmount);
	szAmount.Format(_T("%.2f"), nTotalAmount);
	MoneyToString(szAmount, tmpStr);	
#ifdef UNICODE
	tmpStr += _T(" \x111\x1ED3ng.");
#endif
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);

	GetDoctorName(pMF->GetCurrentUser());

	rpt.PrintPreview();
}


/*
void CPrintForms::PrintDeliveryOrder(long nOrderID)
{
	CGuiMainFrame *pMF = (CGuiMainFrame *) AfxGetMainWnd(); 
	CString szSQL, szWhere, szOrderBy;

	//if(pMF->m_szDrugOrderByName == _T("Y")){
	//	szOrderBy.Format(_T(" ORDER BY p_name, p_unit, pi_insprice "));
	//}
	//else
	//	szOrderBy.Format(_T(" ORDER BY pmssel_index "));

	CRecord rs(&pMF->m_db);
	//if(!bOpenStatus){
	//	szWhere.Format(_T(" AND pmsse_status <>'O' "));
	//}
	szSQL.Format(_T("  SELECT mt_orderno as invoiceno,") \
	_T("  	mt_orderdate as orderdate,") \
	_T("  	get_syssel_desc('pms_order_status', mt_status) as status,") \
	_T("  	mt_deliveryby as deliverby,") \
	_T("  	mt_receiveby as receiveby,") \
	_T(" 	get_doctype(mt_doctype) as type,") \
	_T(" 	get_storagename(mt_storage_id) as stockname,	") \
	_T("	mt_description as note ") \
	_T("  FROM m_transaction") \
	_T("  WHERE mt_transaction_id= %d %s "), nOrderID, szWhere);
	//_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	
	CReport rpt;
	CString tmpStr, szNameReport;
	rs.GetValue(_T("type"), tmpStr);
	if (tmpStr!=_T("W"))
		szNameReport.Format(_T("%s"), _T("Reports/HMS/PM_PHIEUXUATKHO.RPT") );
	else
		szNameReport.Format(_T("%s"), _T("Reports/HMS/PMS_WAREHOUSERECEIPT.RPT"));

	if(!rpt.Init(szNameReport) )
		return;
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);

	rs.GetValue(_T("InvoiceNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);
	rs.GetValue(_T("orderdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("OrderDate"), CDate::Convert(rs.GetValue(_T("orderdate")), yyyymmdd, ddmmyyyy));
	CString szDate, szSysDate;
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), szDate);
	rs.GetValue(_T("Status"), tmpStr);
	
	rpt.GetReportHeader()->SetValue(_T("Status"), tmpStr);
	rs.GetValue(_T("StockName"), tmpStr);	
	rpt.GetReportHeader()->SetValue(_T("StockName"), tmpStr);
	rs.GetValue(_T("note"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Comment"), tmpStr);
	rs.GetValue(_T("type"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Reason"), tmpStr);
	rs.GetValue(_T("deliverby"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DeliverBy"), tmpStr);
	rs.GetValue(_T("ReceiveBy"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ReceiveBy"), tmpStr);
	//Page Header
	//Report Detail

	int nCount = 0;
	double dMoney = 0;
	double nTotalAmount = 0;
	CString szMoney, szExprDate;
	szSQL.Format(_T("SELECT mp_product_id as pid, mp_name as pnme, get_uomname(mp_uom_id) as uom,") \
		_T(" mpi_lotno as lot, mpi_expdate as expdte, sum(mtl_qtyorder) as qtyord, mpi_unitprice as untprc ") \
		_T(" FROM m_product ") \
		_T(" LEFT JOIN m_product_item ON(mpi_product_id=mp_product_id) ") \
		_T(" LEFT JOIN m_transactionline ON(mpi_product_item_id=mtl_product_item_id) ") \
		_T(" WHERE mtl_transaction_id=%d GROUP BY mp_product_id, mp_name, mp_uom_id, mpi_lotno, mpi_expdate, mpi_unitprice %s "), nOrderID, szOrderBy);
	rs.ExecSQL(szSQL);

	CReportSection* rptDetail = rpt.GetDetail(0); 
	while(!rs.IsEOF()){ 
		rptDetail = rpt.AddDetail();
		rs.GetValue(_T("expdte"), tmpStr);
		szExprDate = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		dMoney = ToDouble(rs.GetValue(_T("qtyord")))*ToDouble(rs.GetValue(_T("untprc")));
		nTotalAmount += dMoney;
		szMoney.Format(_T("%.2f"), dMoney);
		nCount++;	
		tmpStr.Format(_T("%d"), nCount);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("pnme"), tmpStr) ;
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("uom"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("lot"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rptDetail->SetValue(_T("5"), szExprDate);
		FormatCurrencyStr(rs.GetValue(_T("qtyord")), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		FormatCurrencyStr(rs.GetValue(_T("untprc")), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		FormatCurrencyStr(szMoney, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.MoveNext();
	}
	//Page Footer
	//Report Footer
	tmpStr.Format(_T("%d"), nCount);
	rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
	FormatCurrency(nTotalAmount, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	szMoney.Format(_T("%.2f"), nTotalAmount);
	MoneyToString(szMoney, tmpStr);
#ifdef UNICODE
	if(!tmpStr.IsEmpty())
		tmpStr += _T(" \x111\x1ED3ng");
#else
	if(!tmpStr.IsEmpty())
		tmpStr += _T(" ®ång");
#endif
	CString szDte;
	szDte = pMF->GetSysDate();
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("ReportDate")), szDte.Right(2), szDte.Mid(5, 2), szDte.Left(4));
	rpt.GetReportFooter()->SetValue(_T("ReportDate"), tmpStr);
	rpt.PrintPreview();	// TODO: Add your command handler code here

}	


*/

void CPrintForms::PrintDDO(long nOrderID)
{
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CString tmpStr, tmpStr2, szDate, szSQL,szDesc, szOrderBy;
	CString szFloor;
	
	double	nTotalAmount = 0, nAmount=0;
	int		nQty, nTotalItem=1;
	int		nPrint=0;
	bool bDirect = false;

	//if(m_szDrugOrderByName == _T("Y")){
	//	szOrderBy.Format(_T(" ORDER BY name, unit, price "));
	//}
	//else
	//	szOrderBy.Format(_T(" ORDER BY pmssel_index "));
	szOrderBy.Format(_T(" ORDER BY name, unit, price "));

	szSQL.Format(_T(" SELECT mt_orderno as ordno,") \
				_T("   mt_doctype          AS doctype,") \
				_T("   mt_deliveryby       AS deliveryby,") \
				_T("   mt_description      AS descr,") \
				_T("   mt_orderdate        AS orddte,") \
				_T("   mt_status           AS stt,") \
				_T("   get_storagename(mt_storage_id)       AS stock,") \
				_T("   get_storagename(mt_storage_to_id)    AS stock_to,") \
				_T("   get_departmentname(mt_department_to_id) AS dept_to") \
				_T(" FROM   m_transaction") \
				_T(" WHERE mt_transaction_id = %ld"), nOrderID);

	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	CString szType, szReportName, szSendBy;
	rs.GetValue(_T("doctype"), szType);
	if(szType == _T("B"))
	{
		szOrderBy.Format(_T(" ORDER BY name, unit, price "));
		szReportName = _T("Reports/HMS/PMS_DRUGORDER2.RPT");
	}
	else
		szReportName = _T("Reports/HMS/PMS_DRUGORDER2.RPT");
	rs.GetValue(_T("deliveryby"), szSendBy);
	rs.GetValue(_T("descr"), szDesc);
	
	
	rs.GetValue(_T("pmsse_floor"), szFloor);
	//szFloor = GetSelectionString(_T("hms_floor"), tmpStr);


	if(szType == _T("DDO") || szType == _T("DPO") || szType == _T("C") || szType == _T("DMO"))
	{
		CRecord rst(&pMF->m_db);
		CRecord rsp(&pMF->m_db);
		CString szGroup, szTitle, szPrintType, szNote;
		szSQL.Format(_T("SELECT * FROM pms_drugtype2 ORDER BY pmdt_id"));
		rst.ExecSQL(szSQL);
		while(!rst.IsEOF())
		{
			//_debug(_T("nOrderID:%ld : szGroup:%s"), nOrderID, szGroup);
			rst.GetValue(_T("pmdt_name"), tmpStr);
			StringUpper(tmpStr, szTitle);
			rst.GetValue(_T("pmdt_desc"), szGroup);
			szGroup.Trim();
			rst.GetValue(_T("pmdt_printtype"), szPrintType);

			if(szPrintType == _T("2"))
			{
				rst.GetValue(_T("pmdt_reportname"), tmpStr);
				if(tmpStr.IsEmpty())
					szReportName = _T("Reports/HMS/PMS_DRUGORDER2.RPT");
				else
					szReportName.Format(_T("Reports/HMS/%s"), tmpStr);
				_debug(_T("%s"), szReportName);
				if(!rpt.Init(szReportName) )
					return;   			
			
				//Report Header
				rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
				rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
				rpt.GetReportHeader()->Format(_T("TITLE"), szTitle);
				rs.GetValue(_T("orddte"), tmpStr);
				szDate.Format(rpt.GetReportHeader()->GetValue(_T("OrderDate")), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
				rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);
				rpt.GetReportHeader()->SetValue(_T("OrderID"), rs.GetValue(_T("ordno")));
				rs.GetValue(_T("stt"), tmpStr);
				//tmpStr = GetStatusString(tmpStr);
				rpt.GetReportHeader()->SetValue(_T("Status"), tmpStr);
				rs.GetValue(_T("stock"), tmpStr);	
				//tmpStr = GetStockName(ToInt(tmpStr));
				rpt.GetReportHeader()->SetValue(_T("FromStock"), tmpStr);
				rpt.GetReportHeader()->SetValue(_T("FromStock2"), tmpStr);

				rpt.GetReportHeader()->SetValue(_T("Comment"), szDesc);
				rpt.GetReportHeader()->SetValue(_T("Comment2"), szDesc);

				
				//rs.GetValue(_T("stock_to"), tmpStr);
				//if(!tmpStr.IsEmpty())
				//{
				//	//tmpStr = GetStockName(ToInt(tmpStr));
				//	rpt.GetReportHeader()->SetValue(_T("ToStock"), tmpStr);
				//	rpt.GetReportHeader()->SetValue(_T("ToStock2"), tmpStr);
				//}

				if(!tmpStr.IsEmpty())
				{
					tmpStr = rs.GetValue(_T("dept_to"));
					rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);
					rpt.GetReportHeader()->SetValue(_T("Department2"), tmpStr);
				}
				tmpStr.Empty();
				rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("1"));

				rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE2"), pMF->m_szHealthService);
				rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME2"), pMF->m_szHospitalName);
				rpt.GetReportHeader()->Format(_T("TITLE2"), szTitle);
				rs.GetValue(_T("orddte"), tmpStr);
				szDate.Format(rpt.GetReportHeader()->GetValue(_T("OrderDate")), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
				rpt.GetReportHeader()->SetValue(_T("OrderDate2"), szDate);
				rpt.GetReportHeader()->SetValue(_T("OrderID2"), rs.GetValue(_T("ordno")));
				rpt.GetReportHeader()->SetValue(_T("Sheet2"), _T("2"));

				rpt.GetReportHeader()->SetValue(_T("Floor"), szFloor);
				rpt.GetReportHeader()->SetValue(_T("Floor2"), szFloor);

			}
			else
			{
				if(!rpt.Init(szReportName) )
					return;
			
			

				//Report Header
				rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
				rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
				rpt.GetReportHeader()->Format(_T("TITLE"), szTitle);
				rs.GetValue(_T("mt_orderdate"), tmpStr);
				szDate.Format(rpt.GetReportHeader()->GetValue(_T("OrderDate")), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
				rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);
				rpt.GetReportHeader()->SetValue(_T("OrderID"), rs.GetValue(_T("ordno")));
				rs.GetValue(_T("mt_expstockid"), tmpStr);
				//rpt.GetReportHeader()->SetValue(_T("FromStock"), GetStockName(ToInt(tmpStr)));
				rs.GetValue(_T("mt_impstockid"), tmpStr);
				if(!tmpStr.IsEmpty())
					//rpt.GetReportHeader()->SetValue(_T("ToStock"), GetStockName(ToInt(tmpStr)));

				rs.GetValue(_T("mt_deptid"), tmpStr);
				if(!tmpStr.IsEmpty())
					//rpt.GetReportHeader()->SetValue(_T("Department"), GetDepartmentName(tmpStr));
				tmpStr.Empty();
				rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("1"));


				rpt.GetReportHeader()->SetValue(_T("Floor"), szFloor);

			}

			
			//Page Header
			//Report Detail

			szSQL.Format(_T(" SELECT") \
						_T("   product_id                                  AS id,") \
						_T("   product_name                                        AS name,") \
						_T("   product_uomname                         AS unit,") \
						_T("   Sum(mtl_qtyorder)                              AS qty,") \
						_T("   product_taxprice                                  AS price,") \
						_T("   Sum(mtl_qtyorder*product_taxprice)                AS amount,") \
						_T("   product_manufacturename AS prod_country") \
						_T(" FROM   m_transactionline") \
						_T("        LEFT JOIN m_productitem_view ON(product_item_id=mtl_product_item_id)") \
						_T(" WHERE mtl_transaction_id = %d") \
						_T(" AND product_producttype IN (%s)") \
						_T(" GROUP  BY product_id,") \
						_T("           product_name,") \
						_T("           product_uomname,") \
						_T("           product_taxprice,") \
						_T("           product_manufacturename %s"), nOrderID, szGroup, szOrderBy);
			//_msg(_T("%s"), szSQL);
			rs2.ExecSQL(szSQL);
			if(rs2.IsEOF()){
				rst.MoveNext();
				continue;
			}
			CReportSection* rptDetail = rpt.GetDetail(0); 
			rpt.ResetContent();
			nTotalItem=1;
			nTotalAmount = 0;
			CString szItemId;
			while(!rs2.IsEOF())
			{
				rptDetail = rpt.AddDetail();
				rs2.GetValue(_T("id"), szItemId);
				if(szPrintType == _T("2"))
				{
					tmpStr.Format(_T("%d"), nTotalItem++);
					rptDetail->SetValue(_T("1"), tmpStr);
					rptDetail->SetValue(_T("12"), tmpStr);
					rs2.GetValue(_T("name"), tmpStr);
					rptDetail->SetValue(_T("2"), tmpStr);
					rptDetail->SetValue(_T("22"), tmpStr);
					rs2.GetValue(_T("unit"), tmpStr);
					rptDetail->SetValue(_T("3"), tmpStr);
					rptDetail->SetValue(_T("32"), tmpStr);
					rs2.GetValue(_T("qty"), nQty);
					if(szGroup == _T("'A1130'") || szGroup == _T("'A1140'") || szGroup == _T("'A2000'"))
						MoneyToString(ToString(nQty), szNote);
					else
						szNote.Empty();
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4"), tmpStr2);
					rptDetail->SetValue(_T("42"), tmpStr2);
					rs2.GetValue(_T("Price"), tmpStr);
					FormatCurrencyStr(tmpStr, tmpStr2);
					rptDetail->SetValue(_T("5"), tmpStr2);
					rptDetail->SetValue(_T("52"), tmpStr2);
					rs2.GetValue(_T("amount"), nAmount);
					nTotalAmount += nAmount;
					
					rptDetail->SetValue(_T("6"), szNote);
					rptDetail->SetValue(_T("62"), szNote);
				}
				else
				{
					tmpStr.Format(_T("%d"), nTotalItem++);
					rptDetail->SetValue(_T("1"), tmpStr);
					rs2.GetValue(_T("name"), tmpStr);
					rptDetail->SetValue(_T("2"), tmpStr);
					rs2.GetValue(_T("unit"), tmpStr);
					
					rptDetail->SetValue(_T("3"), tmpStr);
					rs2.GetValue(_T("qty"), nQty);
					if(szGroup == _T("'A1130'") || szGroup == _T("'A1140'")|| szGroup == _T("'A2000'"))
						MoneyToString(ToString(nQty), szNote);
					else
						szNote.Empty();
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4"), tmpStr2);
					rs2.GetValue(_T("Price"), tmpStr);
					FormatCurrencyStr(tmpStr, tmpStr2);
					rptDetail->SetValue(_T("5"), tmpStr2);
					rs2.GetValue(_T("amount"), nAmount);
					nTotalAmount += nAmount;
					rptDetail->SetValue(_T("6"), szNote);
				}
				
				if(szGroup.Find(_T("'A1130'")) != -1 || szGroup.Find(_T("'A1140'")) != -1 || szGroup.Find(_T("'A2000'")) != -1)
				{
					//Ne du tru thuoc(phieu linh)
					if(szType == _T("DDO") || szType == _T("DPO") || szType == _T("DMO"))
					{
						szSQL.Format(_T(" SELECT") \
						_T("   hpo_docno            AS docno,") \
						_T("   Trim(hp_surname||' '||hp_midname||' '||hp_firstname) AS pname,") \
						_T("   Sum(hpol_qtyorder)   AS orderqty") \
						_T(" FROM   hms_ipharmaorder") \
						_T("        LEFT JOIN hms_ipharmaorderline ON(hpo_orderid=hpol_orderid)") \
						_T("        LEFT JOIN hms_patient ON(hp_patientno=hpo_patientno)") \
						_T(" WHERE  hpo_transaction_id=%d") \
						_T("        AND hpol_product_id='%s'") \
						_T(" GROUP  BY hpo_docno,") \
						_T("           hp_surname,") \
						_T("           hp_midname,") \
						_T("           hp_firstname") \
						_T(" ORDER  BY pname"), nOrderID, szItemId);
						//_msg(_T("%s"), szSQL);
					}
					else	//Neu bo sung tu truc
					{
						szSQL.Format(_T(" SELECT 	hmc_docno as docno,") \
						_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
						_T(" 	hmc_qty as orderqty") \
						_T(" FROM hms_medicalcabinet_issue") \
						_T(" LEFT JOIN hms_doc ON(hd_docno=hmc_docno)") \
						_T(" LEFT JOIN hms_patient ON(hp_patientno=hd_patientno)") \
						_T(" WHERE hmc_supplementid=%d ") \
						_T(" and hmc_itemid='%s' ORDER BY docno,pname "), nOrderID, szItemId);

					}
					rsp.ExecSQL(szSQL);
					while(!rsp.IsEOF()){
						rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
/*
						rptDetail->GetItem(_T("1"))->SetItalic(true);	
						rptDetail->GetItem(_T("2"))->SetItalic(true);	
						rptDetail->GetItem(_T("3"))->SetItalic(true);	
						rptDetail->GetItem(_T("4"))->SetItalic(true);	
						rptDetail->GetItem(_T("5"))->SetItalic(true);	
						rptDetail->GetItem(_T("6"))->SetItalic(true);	
*/
						rptDetail->SetValue(_T("1"), _T(""));
						rsp.GetValue(_T("pname"), tmpStr);
						rptDetail->SetValue(_T("2"), tmpStr);
						rsp.GetValue(_T("docno"), tmpStr);
						rptDetail->SetValue(_T("3"), tmpStr);
						rsp.GetValue(_T("orderqty"), nQty);
						FormatCurrency(nQty, tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);
						rptDetail->SetValue(_T("5"), _T(""));
						MoneyToString(ToString(nQty), szNote);
						rptDetail->SetValue(_T("6"), szNote);
						if(szPrintType == _T("2")){
/*
							rptDetail->GetItem(_T("12"))->SetItalic(true);	
							rptDetail->GetItem(_T("22"))->SetItalic(true);	
							rptDetail->GetItem(_T("32"))->SetItalic(true);	
							rptDetail->GetItem(_T("42"))->SetItalic(true);	
							rptDetail->GetItem(_T("52"))->SetItalic(true);	
							rptDetail->GetItem(_T("62"))->SetItalic(true);	
  */
							rptDetail->SetValue(_T("12"), _T(""));
							rsp.GetValue(_T("pname"), tmpStr);
							rptDetail->SetValue(_T("22"), tmpStr);
							rsp.GetValue(_T("docno"), tmpStr);
							rptDetail->SetValue(_T("32"), tmpStr);
							rsp.GetValue(_T("orderqty"), nQty);
							FormatCurrency(nQty, tmpStr);
							rptDetail->SetValue(_T("42"), tmpStr);
							rptDetail->SetValue(_T("52"), _T(""));
							MoneyToString(ToString(nQty), szNote);
							rptDetail->SetValue(_T("62"), szNote);
						}
						rsp.MoveNext();
					}
				}
				rs2.MoveNext();
			}
			//Page Footer
			//Report Footer
			if(szPrintType == _T("2"))
			{
				tmpStr.Format(_T("%d"), nTotalItem-1);
				rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
				rpt.GetReportFooter()->SetValue(_T("TotalItem2"), tmpStr);
				tmpStr.Format(_T("%.0f"), nTotalAmount);
				FormatCurrencyStr(tmpStr, tmpStr2);
				rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr2);
				rpt.GetReportFooter()->SetValue(_T("TotalAmount2"), tmpStr2);
				CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
				if(img)
				{
//					img->SetHBITMAP(pMF->GetSignatureImage(szSendBy));
					img->SetFixedHeight(false);
				}

//				rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(szSendBy));
				
				CReportItem *img2 = rpt.GetReportFooter()->GetItem(_T("Signature2"));
				if(img2)
				{
//					img2->SetHBITMAP(pMF->GetSignatureImage(szSendBy));
					img2->SetFixedHeight(false);
				}

//				rpt.GetReportFooter()->SetValue(_T("DoctorName2"), GetDoctorName(szSendBy));


				tmpStr = pMF->GetSysDate();	
				szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
				rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
				rpt.GetReportFooter()->SetValue(_T("PrintDate2"), szDate);
			//	rpt.Print(bDirect);
			//	bDirect=true;
				//rpt.PrintPreview();
			}
			else
			{
				tmpStr.Format(_T("%d"), nTotalItem-1);
				rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
				tmpStr.Format(_T("%.0f"), nTotalAmount);
				FormatCurrencyStr(tmpStr, tmpStr2);
				rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr2);
				CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
				if(img)
				{
//					img->SetHBITMAP(pMF->GetSignatureImage(szSendBy));
					img->SetFixedHeight(false);
				}
//				rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(szSendBy));

				tmpStr = pMF->GetSysDate();	
				szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
				rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
				rpt.PrintPreview();
			//	bDirect=true;
			//	rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("2"));
			//	rpt.Print(bDirect);
			}
			//rpt.PrintPreview();
			rst.MoveNext();
		}
	}
	else
	{
		if(!rpt.Init(szReportName) )
			return;
		//Report Header
		rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
		rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
		rs.GetValue(_T("pmsse_orderdate"), tmpStr);
		szDate.Format(rpt.GetReportHeader()->GetValue(_T("OrderDate")), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);
		rpt.GetReportHeader()->SetValue(_T("OrderID"), rs.GetValue(_T("ordno")));
		rs.GetValue(_T("pmsse_expstockid"), tmpStr);
//		rpt.GetReportHeader()->SetValue(_T("FromStock"), GetStockName(ToInt(tmpStr)));
		rs.GetValue(_T("pmsse_impstockid"), tmpStr);
		if(!tmpStr.IsEmpty())
//			rpt.GetReportHeader()->SetValue(_T("ToStock"), GetStockName(ToInt(tmpStr)));

		rs.GetValue(_T("pmsse_deptid"), tmpStr);
		if(!tmpStr.IsEmpty())
//			rpt.GetReportHeader()->SetValue(_T("Department"), GetDepartmentName(tmpStr));
		tmpStr.Empty();
		rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("1"));
		
		rs.GetValue(_T("pmsse_deptid"), tmpStr);
		if(!tmpStr.IsEmpty())
//			rpt.GetReportHeader()->SetValue(_T("Department"), GetDepartmentName(tmpStr));

		rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("1"));

		rpt.GetReportHeader()->SetValue(_T("Floor"), szFloor);
		rpt.GetReportHeader()->SetValue(_T("Floor2"), szFloor);


		//Page Header
		//Report Detail
		szSQL.Format(_T(" SELECT ") \
				_T(" 	pmi_id as id,") \
				_T(" 	pmi_name as name,") \
				_T(" 	pmi_unit	as unit,") \
				_T("	sum(pmssel_qty) as qty, ") \
				_T(" 	pmsi_vatprice as price,") \
				_T(" 	sum(pmssel_qty*pmsi_vatprice) as amount, ") \
				_T(" 	(SELECT sc_name FROM sys_country WHERE sc_id=pmsi_countryid) as prod_country") \
				_T(" FROM pms_stockexport_line") \
				_T(" LEFT JOIN pms_stockitems ON(pmsi_id=pmssel_sitemid)") \
				_T(" LEFT JOIN pms_items ON(pmi_id=pmsi_itemid)") \
				_T(" WHERE pmssel_id=%d ") \
				_T(" GROUP BY pmi_id, pmi_name, pmi_unit, pmsi_vatprice, pmsi_countryid %s ") ,	nOrderID, szOrderBy);

		rs2.ExecSQL(szSQL);

		CReportSection* rptDetail = rpt.GetDetail(0); 
		while(!rs2.IsEOF())
		{
			rptDetail = rpt.AddDetail();
			tmpStr.Format(_T("%d"), nTotalItem++);
			rptDetail->SetValue(_T("1"), tmpStr);
			rs2.GetValue(_T("name"), tmpStr);
			rptDetail->SetValue(_T("2"), tmpStr);
			rs2.GetValue(_T("unit"), tmpStr);
			
			rptDetail->SetValue(_T("3"), tmpStr);
			rs2.GetValue(_T("qty"), nQty);
			FormatCurrency(nQty, tmpStr2);
			rptDetail->SetValue(_T("4"), tmpStr2);
			rs2.GetValue(_T("Price"), tmpStr);
			FormatCurrencyStr(tmpStr, tmpStr2);
			rptDetail->SetValue(_T("5"), tmpStr2);
			rs2.GetValue(_T("amount"), nAmount);
			nTotalAmount += nAmount;
			tmpStr.Empty();
			rptDetail->SetValue(_T("6"), tmpStr);
			rs2.MoveNext();
		}
		//Page Footer
		//Report Footer
		tmpStr.Format(_T("%d"), nTotalItem-1);
		rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
		tmpStr.Format(_T("%.0f"), nTotalAmount);
		FormatCurrencyStr(tmpStr, tmpStr2);
		rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr2);
		
		tmpStr = pMF->GetSysDate();	
		szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
		rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);

		
		CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
		if(img)
		{
//			img->SetHBITMAP(pMF->GetSignatureImage(szSendBy));
			img->SetFixedHeight(false);
		}
//		rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(szSendBy));

/*
		CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
		if(img)
		{
			img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
			img->SetFixedHeight(false);
		}
*/
		rpt.PrintPreview();
		//rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("2"));
		//rpt.Print(bDirect);
		
	}

}
void CPrintForms::PrintDDO2(long nOID, bool bPreviewOnly)
{
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rsReport(&pMF->m_db);
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CRecord rst(&pMF->m_db);
	CRecord rsp(&pMF->m_db);
	CString szGroup, szTitle, szPrintType, szNote;
	CString tmpStr, tmpStr2, szDate, szSQL,szDesc, szOrderBy, szStorageCategory, szOrderType;
	int nStorage_id =0;
	CString szID;
	long nProductID = 0, nProductItemID = 0;

	szID.Format(_T("%ld"), nOID);

	double	nTotalAmount = 0, nAmount=0, nAmountServ = 0, nTotalAmountServ =0;
	int		nQty, nTotalItem=1;
	int		nPrint=0, nBreak = 0;
	bool bDirect = false;
	
	szOrderBy.Format(_T(" ORDER BY name, unit, price "));
	szSQL.Format(_T(" SELECT distinct mt_orderno, ") \
				_T("        mt_orderdate, ") \
				_T("        mt_approveddate, ") \
				_T("        mt_postedby, ") \
				_T("        mt_storage_id, ") \
				_T("        Get_storagename(mt_storage_id)          AS mt_storage_name, ") \
				_T("        Get_storagename(mt_storage_to_id)       AS mt_storageto_name, ") \
				_T("        Get_departmentname(mt_department_id)    AS mt_department_name, ") \
				_T("        Get_departmentname(mt_department_to_id) AS mt_departmentto_name, ") \
				_T("        mt_doctype, ") \
				_T("        mt_description, ") \
				_T("		mt_objecttype, ") 
				_T("		hpo_ordertype, mt_documentno ") \
				_T(" FROM   m_transaction ") \
				_T(" LEFT JOIN hms_ipharmaorder ON (mt_doctype = 'DMO' AND hpo_transaction_id = mt_transaction_id)") \
				_T(" WHERE  mt_transaction_id = %ld ") \
				_T(" AND mt_status NOT IN ('C')"), nOID);

	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	CString szType, szReportName, szSendBy, szObjectType;
	rs.GetValue(_T("mt_doctype"), szType);
	rs.GetValue(_T("mt_orderno"), szID);
	rs.GetValue(_T("hpo_ordertype"), szOrderType);
	if(szType == _T("DDO") || szType == _T("DPO") || szType == _T("DMO") || szType==_T("CSO"))
		szOrderBy.Format(_T(" ORDER BY product_uomindex, product_name, product_uomname "));
	rs.GetValue(_T("mt_postedby"), szSendBy);
	rs.GetValue(_T("mt_description"), szDesc);
	
	rs.GetValue(_T("mt_storage_id"), nStorage_id);
	rs.GetValue(_T("mt_objecttype"), szObjectType);
	
	//Kiem tra kho dich vu
	if (szObjectType == _T("S"))
		szStorageCategory = _T("S");
	else if (szObjectType == _T("I"))
		szStorageCategory = _T("I");
	else if (szObjectType == _T("P"))
		szStorageCategory = _T("P");
	else
	{
		szSQL.Format(_T("SELECT msl_category category FROM m_storagelist WHERE msl_storage_id = %d"), nStorage_id);
		rsReport.ExecSQL(szSQL);
		rsReport.GetValue(_T("category"), szStorageCategory);
	}
	if(szType == _T("DPO") || szType == _T("DMO") || szType == _T("DDO") || szType == _T("CSO"))
	{
		szSQL.Format(_T(" SELECT DISTINCT pmdt_name,") \
			_T("   pmdt_desc,") \
			_T("   pmdt_reportname, pmdt_index, pmdt_printtype  ") \
			_T(" FROM pms_drugtype,") \
			_T("   m_transactionline,") \
			_T("   m_productitem_view") \
			_T(" WHERE mtl_transaction_id =%ld") \
			_T(" AND mtl_product_item_id  =product_item_id") \
			_T(" AND instr(pmdt_desc, product_producttype) > 0") \
			_T(" ORDER BY pmdt_index "), nOID);
		rst.ExecSQL(szSQL);
		while(!rst.IsEOF())
		{
			rst.GetValue(_T("pmdt_name"), tmpStr);
			StringUpper(tmpStr, szTitle);
			rst.GetValue(_T("pmdt_desc"), szGroup);
			szGroup.Trim();
			
			rst.GetValue(_T("pmdt_reportname"), tmpStr);
			if(!tmpStr.IsEmpty())
				if(szStorageCategory == _T("S"))
					szReportName.Format(_T("Reports/HMS/%sDV.RPT"), tmpStr);
				else if(szStorageCategory == _T("P"))
					szReportName.Format(_T("Reports/HMS/%sBD.RPT"), tmpStr);
				else if(szStorageCategory == _T("I"))
					szReportName.Format(_T("Reports/HMS/%sBH.RPT"), tmpStr);
				else
					szReportName.Format(_T("Reports/HMS/%s.RPT"), tmpStr);
			
			if(!rpt.Init(szReportName) )
				return;   			
		
			//Report Header
			rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
			rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
			if (szStorageCategory != _T("S"))
				rpt.GetReportHeader()->Format(_T("TITLE"), szTitle);
			rs.GetValue(_T("mt_orderdate"), tmpStr);
			szDate.Format(rpt.GetReportHeader()->GetValue(_T("OrderDate")), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
			rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);
			rpt.GetReportHeader()->SetValue(_T("OrderID"), szID);
			if (szOrderType == _T("M"))
				rpt.GetReportHeader()->SetValue(_T("Purpose"), _T("\x44\xE0nh \x63ho \x63\xE1\x63 \x63\x61 ph\x1EABu thu\x1EADt, th\x1EE7 thu\x1EADt"));
			rs.GetValue(_T("mt_statusdesc"), tmpStr);
			//tmpStr = GetStatusString(tmpStr);
			rpt.GetReportHeader()->SetValue(_T("Status"), tmpStr);
			rs.GetValue(_T("mt_storage_name"), tmpStr);
			//tmpStr = GetStorageName(ToInt(tmpStr));
			rpt.GetReportHeader()->SetValue(_T("FromStock"), tmpStr);
			rpt.GetReportHeader()->SetValue(_T("Comment"), szDesc);
			
			rs.GetValue(_T("mt_storageto_name"), tmpStr);
			if(!tmpStr.IsEmpty())
				rpt.GetReportHeader()->SetValue(_T("ToStock"), tmpStr);

			rs.GetValue(_T("mt_department_name"), tmpStr);
			if(!tmpStr.IsEmpty())
				rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);

			rs.GetValue(_T("mt_departmentto_name"), tmpStr);
			if(!tmpStr.IsEmpty())
				rpt.GetReportHeader()->SetValue(_T("DepartmentTo"), tmpStr);
			tmpStr.Empty();
			rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("1"));

			if (pMF->GetModuleID() == _T("MA"))
			{
				rs.GetValue(_T("mt_documentno"), tmpStr);
				rpt.GetReportHeader()->SetValue(_T("ReferenceNo"), tmpStr);
			}

			//*?Lay gia ban tu line thuoc
			//Page Header
			//Report Detail
			if (szStorageCategory == _T("S") && (szGroup.Find(_T("'A1130'")) == -1 && szGroup.Find(_T("'A1140'")) == -1))
				szSQL.Format(_T(" SELECT product_id                             AS id, ") \
				_T("		product_item_id item_id,") \
				_T("        product_name                           AS name, ") \
				_T("        product_uomname                        AS unit, ") \
				_T("        Sum(mtl_qtyorder)                      AS qtyorder, ") \
				_T("        Sum(mtl_qtysold)                       AS qtysold, ") \
				_T("        Sum(mtl_qtypolicy)                     AS qtypolicy, ") \
				_T("        Sum(mtl_qtysoldins)                    AS qtysoldins, ") \
				_T("        Sum(mtl_qtyotherins)                   AS qtyotherins, ") \
				_T("        Sum(mtl_qtyservice)                    AS qtyservice, ") \
				_T("        Sum(mtl_qtyother)                      AS qtyother, ") \
				_T("        mtl_saleprice                      AS price, ") \
				_T("        Sum(mtl_qtyorder * mtl_saleprice) AS amount ") \
				_T(" FROM   m_transactionline ") \
				_T(" LEFT JOIN m_productitem_view ON( product_item_id = mtl_product_item_id ) ") \
				_T(" WHERE  mtl_transaction_id = %ld ") \
				_T("        AND product_producttype NOT IN('A1130', 'A1140') ") \
				_T(" GROUP  BY product_id, ") \
				_T("		   product_item_id,") \
				_T("           product_name, ") \
				_T("           product_uomname, ") \
				_T("		   mtl_saleprice,") \
				_T("           product_uomindex ") \
				_T("		   %s "), nOID, szOrderBy);
			else
				szSQL.Format(_T(" SELECT product_id                             AS id, ") \
				_T("		product_item_id item_id, ") \
				_T("        product_name                           AS name, ") \
				_T("        product_uomname                        AS unit, ") \
				_T("        Sum(mtl_qtyorder)                      AS qtyorder, ") \
				_T("        Sum(mtl_qtysold)                       AS qtysold, ") \
				_T("        Sum(mtl_qtypolicy)                     AS qtypolicy, ") \
				_T("        Sum(mtl_qtysoldins)                    AS qtysoldins, ") \
				_T("        Sum(mtl_qtyotherins)                   AS qtyotherins, ") \
				_T("        Sum(mtl_qtyservice)                    AS qtyservice, ") \
				_T("        Sum(mtl_qtyother)                      AS qtyother, ") \
				_T("        mtl_saleprice                      AS price, ") \
				_T("        Sum(mtl_qtyorder * mtl_saleprice) AS amount ") \
				_T(" FROM   m_transactionline ") \
				_T(" LEFT JOIN m_productitem_view ON( product_item_id = mtl_product_item_id ) ") \
				_T(" WHERE  mtl_transaction_id = %ld ") \
				_T("        AND product_producttype IN(%s) ") \
				_T(" GROUP  BY product_id, ") \
				_T("		   product_item_id,") \
				_T("           product_name, ") \
				_T("           product_uomname, ") \
				_T("           mtl_saleprice, ") \
				_T("           product_uomindex ") \
				_T("		   %s "), nOID, szGroup, szOrderBy);

			rs2.ExecSQL(szSQL);
_fmsg(_T("%s"), szSQL);
			if(rs2.IsEOF()){
				rst.MoveNext();
				continue;

			}
			if (szStorageCategory == _T("S") && (szGroup.Find(_T("'A1130'")) == -1 && szGroup.Find(_T("'A1140'")) == -1))
			{
				nBreak++;
			}
			CReportSection* rptDetail = rpt.GetDetail(0); 
			rpt.ResetContent();
			nTotalItem=1;
			nTotalAmount = 0;
			CString szItemId;
			while(!rs2.IsEOF())
			{
				rptDetail = rpt.AddDetail();
				rs2.GetValue(_T("id"), nProductID);
				rs2.GetValue(_T("item_id"), nProductItemID);
				tmpStr.Format(_T("%d"), nTotalItem++);
				rptDetail->SetValue(_T("1"), tmpStr);
				rs2.GetValue(_T("name"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rs2.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("3"), tmpStr);
				rs2.GetValue(_T("qtyorder"), nQty);

				if(szGroup.Find(_T("'A1130'")) != -1 || szGroup.Find(_T("'A1140'")) != -1 || szGroup.Find(_T("'A2000'")) != -1 || szOrderType == _T("M"))
					MoneyToString(ToString(nQty), szNote);
				else
					szNote.Empty();
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4"), tmpStr2);

				rs2.GetValue(_T("qtysold"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.1"), tmpStr2);

				rs2.GetValue(_T("qtypolicy"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.2"), tmpStr2);
				
				rs2.GetValue(_T("qtysoldins"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.3"), tmpStr2);
				rs2.GetValue(_T("qtyotherins"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.4"), tmpStr2);
				
				rs2.GetValue(_T("qtyservice"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.5"), tmpStr2);

				rs2.GetValue(_T("qtyother"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.6"), tmpStr2);

				rs2.GetValue(_T("price"), tmpStr);
				rs2.GetValue(_T("amount"), nAmount);

				FormatCurrencyStr(tmpStr, tmpStr2);
				rptDetail->SetValue(_T("5"), tmpStr2);

				rptDetail->SetValue(_T("8"), tmpStr2);
				nTotalAmount += nAmount;

				FormatCurrency(nAmount, tmpStr2);
				rptDetail->SetValue(_T("7"), tmpStr2);
				
				rptDetail->SetValue(_T("9"), tmpStr2);

				rptDetail->SetValue(_T("6"), szNote);

				if(szGroup.Find(_T("'A1130'")) != -1 || szGroup.Find(_T("'A1140'")) != -1 || 
					szGroup.Find(_T("'A2000'")) != -1 || szOrderType == _T("M"))
				{
					szSQL.Format(_T(" SELECT    hpo_docno            AS docno, ") \
								_T("           get_patientname(hpo_docno) AS pname, ") \
								_T("           SUM(hpol_qtyorder)   AS orderqty ") \
								_T(" FROM      hms_ipharmaorder ") \
								_T(" LEFT JOIN hms_ipharmaorderline ON( hpo_orderid = hpol_orderid ) ") \
								_T(" WHERE     hpo_transaction_id = %ld ") \
								_T("       AND hpol_product_item_id = %ld ") \
								_T(" GROUP     BY hpo_docno ") \
								_T(" ORDER     BY pname "), nOID, nProductItemID);
					rsp.ExecSQL(szSQL);
					while(!rsp.IsEOF()){
						rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
						rptDetail->SetValue(_T("1"), _T(""));
						rsp.GetValue(_T("pname"), tmpStr);
						rptDetail->SetValue(_T("2"), tmpStr);
						rsp.GetValue(_T("docno"), tmpStr);
						rptDetail->SetValue(_T("3"), tmpStr);
						rsp.GetValue(_T("orderqty"), nQty);
						FormatCurrency(nQty, tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);
						rptDetail->SetValue(_T("5"), _T(""));
						MoneyToString(ToString(nQty), szNote);
						rptDetail->SetValue(_T("6"), szNote);
						rsp.MoveNext();
					}
				}
				rs2.MoveNext();
			}
			//Page Footer
			//Report Footer
			rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1)); 
			tmpStr.Format(_T("%d"), nTotalItem-1);
			rptDetail->SetValue(_T("TotalItem"), tmpStr);
			tmpStr.Format(_T("%.0f"), nTotalAmount);
			FormatCurrencyStr(tmpStr, tmpStr2);
			rptDetail->SetValue(_T("TotalAmount"), tmpStr2);

			MoneyToString(tmpStr, tmpStr2);
			tmpStr.Format(_T("%s \x111\x1ED3ng."), tmpStr2);
			rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);

			CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
			if(img)
			{
				//img->SetHBITMAP(pMF->GetSignatureImage(szSendBy));
				//img->SetFixedHeight(false);
			}

			rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(szSendBy));

			tmpStr = pMF->GetSysDate();	
			szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
			rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
			
			rpt.PrintPreview(bPreviewOnly);
			//_msg(_T("%s %s"), GroupStr, szGroup);
			if (nBreak > 0)
				break;
			rst.MoveNext();
		}
	}
}

void CPrintForms::PM_PrintOldCabinetSheet(long nOID)
{
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rsReport(&pMF->m_db);
	CRecord rs(&pMF->m_db);
	CRecord rsm(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CString tmpStr, tmpStr2, szDate, szSQL,szDesc, szOrderBy;
	CString szFloor;
	int nStorage_id =0;
	CString szID;
	long nProductID;


	szID.Format(_T("%ld"), nOID);

	double	nTotalAmount = 0, nAmount=0, nAmountServ = 0, nTotalAmountServ =0;
	int		nQty, nTotalItem=1, nPrintStorage=0;
	int		nPrint=0;
	bool bDirect = false;
	
	szOrderBy.Format(_T(" ORDER BY name, unit, price "));
	szSQL.Format(_T("SELECT mt_orderno, mt_orderdate, mt_approveddate, mt_postedby, mt_storage_id, ") \
		_T(" get_storagename(mt_storage_id) as mt_storage_name, ") \
		_T(" get_storagename(mt_storage_to_id) as mt_storageto_name, ") \
		_T(" get_departmentname(mt_department_id) as mt_department_name, ") \
		_T(" get_departmentname(mt_department_to_id) as mt_departmentto_name, ") \
		_T(" mt_doctype, mt_description ") \
		_T("FROM m_transaction WHERE mt_transaction_id=%ld "), nOID);
	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	CString szType, szReportName, szSendBy;
	rs.GetValue(_T("mt_doctype"), szType);
	rs.GetValue(_T("mt_orderno"), szID);
	
	if(szType == _T("MOV"))
		szReportName = _T("Reports/HMS/PMS_STOCKTRANSFER.RPT");
	else if(szType == _T("DDO") || szType == _T("DPO") || szType == _T("DMO") || szType==_T("CSO"))
	{
		szOrderBy.Format(_T(" ORDER BY product_uomindex, product_name, product_uomname "));
		szReportName = _T("Reports/HMS/PMS_DRUGORDER.RPT");
	}
	else if(szType == _T("MCO"))
		szReportName = _T("Reports/HMS/PMS_CABINETORDER.RPT");
	else
		szReportName = _T("Reports/HMS/PMS_DRUGEXPORT.RPT");
	rs.GetValue(_T("mt_postedby"), szSendBy);
	rs.GetValue(_T("mt_description"), szDesc);
	
	
	rs.GetValue(_T("mt_floor"), tmpStr);
	//szFloor = pMF->GetSelectionString(_T("hms_floor"), tmpStr);
	szFloor.Empty();

	
	if(szType == _T("DPO") || szType == _T("DMO") || szType == _T("DDO") || szType == _T("CSO"))
	{
		CRecord rst(&pMF->m_db);
		CRecord rsp(&pMF->m_db);
		CString szGroup, szTitle, szPrintType, szNote;
		szSQL.Format(_T(" SELECT DISTINCT pmdt_name,") \
_T("   pmdt_desc,") \
_T("   pmdt_reportname, pmdt_index,pmdt_printtype  ") \
_T(" FROM pms_drugtype,") \
_T("   m_transactionline,") \
_T("   m_productitem_view") \
_T(" WHERE mtl_transaction_id =%ld") \
_T(" AND mtl_product_item_id  =product_item_id") \
_T(" AND instr(pmdt_desc, product_producttype) > 0") \
_T(" ORDER BY pmdt_index "), nOID);
		rst.ExecSQL(szSQL);
		while(!rst.IsEOF())
		{
			rst.GetValue(_T("pmdt_name"), tmpStr);
			StringUpper(tmpStr, szTitle);
			rst.GetValue(_T("pmdt_desc"), szGroup);
			szGroup.Trim();
			rst.GetValue(_T("pmdt_printtype"), szPrintType);


			//Kiem tra neu la kho thiet lap cho doi tuong DICH VU thi goi mau rieng
			rs.GetValue(_T("mt_storage_id"), nStorage_id);
			szSQL.Format(_T(" SELECT msd_printtype FROM m_storage_default WHERE msd_storage_id = %d "), nStorage_id);
			rsReport.ExecSQL(szSQL);
			
			nPrintStorage = rsReport.GetIntValue();


			if(szPrintType == _T("2"))
			{
				rst.GetValue(_T("pmdt_reportname"), tmpStr);
				if(tmpStr.IsEmpty())
					szReportName = _T("Reports/HMS/PMS_DRUGORDER2.RPT");
				else
				{
					if(nPrintStorage == 1)
						szReportName.Format(_T("Reports/HMS/%sDV.RPT"), tmpStr);
					else
						szReportName.Format(_T("Reports/HMS/%s.RPT"), tmpStr);
				}
//_msg(_T("%s, %s"), szSQL, szReportName);
				if(!rpt.Init(szReportName) )
					return;   			
			
				//Report Header
				rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
				rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
				rpt.GetReportHeader()->Format(_T("TITLE"), szTitle);

				rs.GetValue(_T("mt_orderdate"), tmpStr);
				szDate.Format(rpt.GetReportHeader()->GetValue(_T("OrderDate")), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
				rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);
				rpt.GetReportHeader()->SetValue(_T("OrderID"), szID);
				rs.GetValue(_T("mt_statusdesc"), tmpStr);
				//tmpStr = GetStatusString(tmpStr);
				rpt.GetReportHeader()->SetValue(_T("Status"), tmpStr);
				rs.GetValue(_T("mt_storage_name"), tmpStr);
				//tmpStr = GetStorageName(ToInt(tmpStr));
				rpt.GetReportHeader()->SetValue(_T("FromStock"), tmpStr);
				rpt.GetReportHeader()->SetValue(_T("FromStock2"), tmpStr);

				rpt.GetReportHeader()->SetValue(_T("Comment"), szDesc);
				rpt.GetReportHeader()->SetValue(_T("Comment2"), szDesc);

				
				rs.GetValue(_T("mt_storageto_name"), tmpStr);
				if(!tmpStr.IsEmpty())
				{
					//tmpStr = GetStorageName(ToInt(tmpStr));
					rpt.GetReportHeader()->SetValue(_T("ToStock"), tmpStr);
					rpt.GetReportHeader()->SetValue(_T("ToStock2"), tmpStr);
				}

				rs.GetValue(_T("mt_department_name"), tmpStr);
				if(!tmpStr.IsEmpty())
				{
					//tmpStr = GetDepartmentName(tmpStr);
					rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);
					rpt.GetReportHeader()->SetValue(_T("Department2"), tmpStr);
				}

				rs.GetValue(_T("mt_departmentto_name"), tmpStr);
				if(!tmpStr.IsEmpty())
				{
					//tmpStr = GetDepartmentName(tmpStr);
					tmpStr.Append(_T(" - "));
					tmpStr.Append(rs.GetValue(_T("mt_storageto_name")));
					rpt.GetReportHeader()->SetValue(_T("DepartmentTo"), tmpStr);
					rpt.GetReportHeader()->SetValue(_T("DepartmentTo2"), tmpStr);
				}
				tmpStr.Empty();
				rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("1"));

				rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE2"), pMF->m_szHealthService);
				rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME2"), pMF->m_szHospitalName);
				rpt.GetReportHeader()->Format(_T("TITLE2"), szTitle);
				rs.GetValue(_T("mt_orderdate"), tmpStr);
				szDate.Format(rpt.GetReportHeader()->GetValue(_T("OrderDate")), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
				rpt.GetReportHeader()->SetValue(_T("OrderDate2"), szDate);
				rpt.GetReportHeader()->SetValue(_T("OrderID2"), szID);
				rpt.GetReportHeader()->SetValue(_T("Sheet2"), _T("2"));

				rpt.GetReportHeader()->SetValue(_T("Floor"), szFloor);
				rpt.GetReportHeader()->SetValue(_T("Floor2"), szFloor);

				szSQL.Format(_T("SELECT MAX(TRUNC_DATE(hpo_orderdate)) as maxorderdate, MIN(TRUNC_DATE(hpo_orderdate)) as minorderdate ") \
					_T(" FROM hms_medical_transaction_view ")\
					_T(" LEFT JOIN hms_ipharmaorder ON(hmt_orderid=hpo_orderid) ") \
					_T(" WHERE hmt_reftransaction_id=%ld"), nOID);
				_fmsg(_T("%s"), szSQL);
				rsm.ExecSQL(szSQL);
				if(!rsm.IsEOF()){
					rpt.GetReportHeader()->SetValue(_T("MaxOrderDate"), CDate::Convert(rsm.GetValue(_T("MaxOrderDate")), yyyymmdd, ddmmyyyy));
					rpt.GetReportHeader()->SetValue(_T("MinOrderDate"), CDate::Convert(rsm.GetValue(_T("MinOrderDate")), yyyymmdd, ddmmyyyy));
				}

			}
			else
			{
				if(!rpt.Init(szReportName) )
					return;
			
			

				//Report Header
				rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
				rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
				rpt.GetReportHeader()->Format(_T("TITLE"), szTitle);
				rs.GetValue(_T("mt_orderdate"), tmpStr);
				szDate.Format(rpt.GetReportHeader()->GetValue(_T("OrderDate")), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
				rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);
				rpt.GetReportHeader()->SetValue(_T("OrderID"), szID);
				rs.GetValue(_T("mt_storage_name"), tmpStr);
				rpt.GetReportHeader()->SetValue(_T("FromStock"), tmpStr);
				rs.GetValue(_T("mt_storageto_name"), tmpStr);
				if(!tmpStr.IsEmpty())
					rpt.GetReportHeader()->SetValue(_T("ToStock"), tmpStr);

				rs.GetValue(_T("mt_department_to_name"), tmpStr);
				if(!tmpStr.IsEmpty())
					rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);
				tmpStr.Empty();
				rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("1"));
				rpt.GetReportHeader()->SetValue(_T("Floor"), szFloor);

				szSQL.Format(_T("SELECT MAX(TRUNC_DATE(hpo_orderdate)) as maxorderdate, MIN(TRUNC_DATE(hpo_orderdate)) as minorderdate ") \
					_T(" FROM hms_medical_transaction_view ")\
					_T(" LEFT JOIN hms_ipharmaorder ON(hmt_orderid=hpo_orderid) ") \
					_T(" WHERE hmt_reftransaction_id=%ld"), nOID);
				//_fmsg(_T("%s"), szSQL);
				rsm.ExecSQL(szSQL);
				if(!rsm.IsEOF()){
					rpt.GetReportHeader()->SetValue(_T("MaxOrderDate"), CDate::Convert(rsm.GetValue(_T("MaxOrderDate")), yyyymmdd, ddmmyyyy));
					rpt.GetReportHeader()->SetValue(_T("MinOrderDate"), CDate::Convert(rsm.GetValue(_T("MinOrderDate")), yyyymmdd, ddmmyyyy));
				}

			}

			
			//Page Header
			//Report Detail

			szSQL.Format(_T(" SELECT ") \
					_T(" 	product_id as id,") \
					_T(" 	product_name as name,") \
					_T(" 	product_uomname	as unit,") \
					_T("	sum(mtl_qtyorder) as qtyorder, ") \
					_T("	sum(mtl_qtysold) as qtysold, ") \
					_T("	sum(mtl_qtypolicy) as qtypolicy, ") \
					_T("	sum(mtl_qtysoldins) as qtysoldins, ") \
					_T("	sum(mtl_qtyotherins) as qtyotherins, ") \
					_T("	sum(mtl_qtyservice) as qtyservice, ") \
					_T("	sum(mtl_qtyother) as qtyother, ") \
					_T(" 	mtl_saleprice as price,") \
					_T(" 	sum(mtl_qtyorder*mtl_saleprice) as amount, ") \
					_T(" 	product_countryname as hpol_productcountry ") \
					_T(" FROM m_transactionline") \
					_T(" LEFT JOIN m_productitem_view ON(product_item_id=mtl_product_item_id)") \
					_T(" WHERE mtl_transaction_id=%ld AND product_producttype in(%s) ") \
					_T(" GROUP BY product_id, product_name, product_uomname, mtl_saleprice, product_countryname, product_uomindex %s "), nOID, szGroup, szOrderBy);
//_msg(_T("%s"), szSQL);
			rs2.ExecSQL(szSQL);
			if(rs2.IsEOF()){
				rst.MoveNext();
				continue;

			}
			CReportSection* rptDetail = rpt.GetDetail(0); 
			rpt.ResetContent();
			nTotalItem=1;
			nTotalAmount = 0;
			CString szItemId;
			while(!rs2.IsEOF())
			{
				rptDetail = rpt.AddDetail();
				rs2.GetValue(_T("id"), nProductID);
				if(szPrintType == _T("2"))
				{
					tmpStr.Format(_T("%d"), nTotalItem++);
					rptDetail->SetValue(_T("1"), tmpStr);
					rptDetail->SetValue(_T("12"), tmpStr);
					rs2.GetValue(_T("name"), tmpStr);
					rptDetail->SetValue(_T("2"), tmpStr);
					rptDetail->SetValue(_T("22"), tmpStr);
					rs2.GetValue(_T("unit"), tmpStr);
					rptDetail->SetValue(_T("3"), tmpStr);
					rptDetail->SetValue(_T("32"), tmpStr);
					rs2.GetValue(_T("qtyorder"), nQty);

					if(szGroup.Find(_T("'A1130'")) != -1 || szGroup.Find(_T("'A1140'")) != -1 || szGroup.Find(_T("'A2000'")) != -1)
						MoneyToString(ToString(nQty), szNote);
					else
						szNote.Empty();
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4"), tmpStr2);
					rptDetail->SetValue(_T("42"), tmpStr2);

					rs2.GetValue(_T("qtysold"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.1"), tmpStr2);

					rs2.GetValue(_T("qtypolicy"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.2"), tmpStr2);
					
					rs2.GetValue(_T("qtysoldins"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.3"), tmpStr2);
					rs2.GetValue(_T("qtyotherins"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.4"), tmpStr2);
					
					rs2.GetValue(_T("qtyservice"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.5"), tmpStr2);

					rs2.GetValue(_T("qtyother"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.6"), tmpStr2);

					
					rs2.GetValue(_T("Price"), tmpStr);
					FormatCurrencyStr(tmpStr, tmpStr2);
					rptDetail->SetValue(_T("5"), tmpStr2);
					rptDetail->SetValue(_T("52"), tmpStr2);
					rs2.GetValue(_T("amount"), nAmount);
					nTotalAmount += nAmount;

					FormatCurrency(nAmount, tmpStr2);
					rptDetail->SetValue(_T("7"), tmpStr2);
					rptDetail->SetValue(_T("72"), tmpStr2);
					
					rptDetail->SetValue(_T("6"), szNote);
					rptDetail->SetValue(_T("62"), szNote);

					rs2.GetValue(_T("price"), tmpStr);
					FormatCurrencyStr(tmpStr, tmpStr2);
					rptDetail->SetValue(_T("8"), tmpStr2);
					rptDetail->SetValue(_T("82"), tmpStr2);
					rs2.GetValue(_T("amount"), nAmountServ);
					nTotalAmountServ += nAmountServ;

					FormatCurrency(nAmountServ, tmpStr2);
					rptDetail->SetValue(_T("9"), tmpStr2);
					rptDetail->SetValue(_T("92"), tmpStr2);
				}
				else
				{
					tmpStr.Format(_T("%d"), nTotalItem++);
					rptDetail->SetValue(_T("1"), tmpStr);
					rs2.GetValue(_T("name"), tmpStr);
					rptDetail->SetValue(_T("2"), tmpStr);
					rs2.GetValue(_T("unit"), tmpStr);
					
					rptDetail->SetValue(_T("3"), tmpStr);
					rs2.GetValue(_T("qtyorder"), nQty);
					if(szGroup.Find(_T("'A1130'")) != -1 || szGroup.Find( _T("'A1140'")) != -1 || szGroup.Find(_T("'A2000'")) != -1)
						MoneyToString(ToString(nQty), szNote);
					else
						szNote.Empty();
					
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4"), tmpStr2);

					rs2.GetValue(_T("qtysold"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.1"), tmpStr2);

					rs2.GetValue(_T("qtypolicy"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.2"), tmpStr2);

					rs2.GetValue(_T("qtysoldins"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.3"), tmpStr2);
					rs2.GetValue(_T("qtyotherins"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.4"), tmpStr2);
					
					rs2.GetValue(_T("qtyservice"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.5"), tmpStr2);

					rs2.GetValue(_T("qtyother"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.6"), tmpStr2);

					rs2.GetValue(_T("Price"), tmpStr);
					FormatCurrencyStr(tmpStr, tmpStr2);
					rptDetail->SetValue(_T("5"), tmpStr2);
					rs2.GetValue(_T("amount"), nAmount);
					nTotalAmount += nAmount;

					rs2.GetValue(_T("price"), tmpStr);
					FormatCurrencyStr(tmpStr, tmpStr2);
					rptDetail->SetValue(_T("8"), tmpStr2);
					rptDetail->SetValue(_T("82"), tmpStr2);
					rs2.GetValue(_T("amount"), nAmountServ);
					nTotalAmountServ += nAmountServ;

					FormatCurrency(nAmountServ, tmpStr2);
					rptDetail->SetValue(_T("9"), tmpStr2);
					rptDetail->SetValue(_T("92"), tmpStr2);

					rptDetail->SetValue(_T("6"), szNote);
				}
				
				if(szGroup.Find(_T("'A1130'")) != -1 || szGroup.Find(_T("'A1140'")) != -1 || szGroup.Find(_T("'A2000'")) != -1)
				{
					//Neu bo sung tu truc
					if(szType == _T("CSO"))
					{
						szSQL.Format(_T(" SELECT 	hmt_docno as docno,") \
									_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
									_T(" 	sum(hmt_qtyissue) as orderqty") \
									_T(" FROM hms_medical_transaction_view") \
									_T(" LEFT JOIN hms_doc ON(hd_docno=hmt_docno)") \
									_T(" LEFT JOIN hms_patient ON(hp_patientno=hd_patientno)") \
									_T(" WHERE hmt_reftransaction_id=%ld ") \
									_T(" and (hmt_suppproduct_id=%ld or hmt_product_id=%ld) ") \
									_T(" GROUP BY hmt_docno, hp_firstname, hp_midname, hp_surname ") \
									_T(" ORDER BY pname"), nOID, nProductID, nProductID);

					}
					else
					{
						szSQL.Format(_T(" SELECT 	hpo_docno as docno,") \
							_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
							_T(" 	sum(hpol_qtyissue) as orderqty") \
							_T(" FROM hms_ipharmaorder") \
							_T(" LEFT JOIN hms_ipharmaorderline ON(hpo_orderid=hpol_orderid)") \
							_T(" LEFT JOIN hms_patient ON(hp_patientno=hpo_patientno)") \
							_T(" WHERE hpo_transaction_id=%ld AND hpol_product_id=%ld ") \
							_T(" GROUP BY hpo_docno, hp_firstname, hp_midname, hp_surname ") \
							_T(" ORDER BY pname"), nOID, nProductID);
					}
					rsp.ExecSQL(szSQL);
					while(!rsp.IsEOF()){
						rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
/*
						rptDetail->GetItem(_T("1"))->SetItalic(true);	
						rptDetail->GetItem(_T("2"))->SetItalic(true);	
						rptDetail->GetItem(_T("3"))->SetItalic(true);	
						rptDetail->GetItem(_T("4"))->SetItalic(true);	
						rptDetail->GetItem(_T("5"))->SetItalic(true);	
						rptDetail->GetItem(_T("6"))->SetItalic(true);	
*/
						rptDetail->SetValue(_T("1"), _T(""));
						rsp.GetValue(_T("pname"), tmpStr);
						rptDetail->SetValue(_T("2"), tmpStr);
						rsp.GetValue(_T("docno"), tmpStr);
						rptDetail->SetValue(_T("3"), tmpStr);
						rsp.GetValue(_T("orderqty"), nQty);
						FormatCurrency(nQty, tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);
						rptDetail->SetValue(_T("5"), _T(""));
						MoneyToString(ToString(nQty), szNote);
						rptDetail->SetValue(_T("6"), szNote);
						if(szPrintType == _T("2")){
/*
							rptDetail->GetItem(_T("12"))->SetItalic(true);	
							rptDetail->GetItem(_T("22"))->SetItalic(true);	
							rptDetail->GetItem(_T("32"))->SetItalic(true);	
							rptDetail->GetItem(_T("42"))->SetItalic(true);	
							rptDetail->GetItem(_T("52"))->SetItalic(true);	
							rptDetail->GetItem(_T("62"))->SetItalic(true);	
  */
							rptDetail->SetValue(_T("12"), _T(""));
							rsp.GetValue(_T("pname"), tmpStr);
							rptDetail->SetValue(_T("22"), tmpStr);
							rsp.GetValue(_T("docno"), tmpStr);
							rptDetail->SetValue(_T("32"), tmpStr);
							rsp.GetValue(_T("orderqty"), nQty);
							FormatCurrency(nQty, tmpStr);
							rptDetail->SetValue(_T("42"), tmpStr);
							rptDetail->SetValue(_T("52"), _T(""));
							MoneyToString(ToString(nQty), szNote);
							rptDetail->SetValue(_T("62"), szNote);
						}
						rsp.MoveNext();
					}
				}
				rs2.MoveNext();
			}
			//Page Footer
			//Report Footer
			if(szPrintType == _T("2"))
			{
				tmpStr.Format(_T("%d"), nTotalItem-1);
				rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
				rpt.GetReportFooter()->SetValue(_T("TotalItem2"), tmpStr);
				tmpStr.Format(_T("%.0f"), nTotalAmount);
				FormatCurrencyStr(tmpStr, tmpStr2);
				rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr2);
				rpt.GetReportFooter()->SetValue(_T("TotalAmount2"), tmpStr2);

				MoneyToString(tmpStr, tmpStr2);
				rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr2);

				CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
				if(img)
				{
		//			img->SetHBITMAP(pMF->GetSignatureImage(szSendBy));
		//			img->SetFixedHeight(false);
				}

				rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(szSendBy));
				
				CReportItem *img2 = rpt.GetReportFooter()->GetItem(_T("Signature2"));
				if(img2)
				{
					//img2->SetHBITMAP(pMF->GetSignatureImage(szSendBy));
					img2->SetFixedHeight(false);
				}

				rpt.GetReportFooter()->SetValue(_T("DoctorName2"), GetDoctorName(szSendBy));


				tmpStr = pMF->GetSysDate();	
				szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
				rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
				rpt.GetReportFooter()->SetValue(_T("PrintDate2"), szDate);
			//	rpt.Print(bDirect);
			//	bDirect=true;
				rpt.PrintPreview();
			}
			else
			{
				tmpStr.Format(_T("%d"), nTotalItem-1);
				rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
				tmpStr.Format(_T("%.0f"), nTotalAmount);
				FormatCurrencyStr(tmpStr, tmpStr2);
				rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr2);

				MoneyToString(tmpStr, tmpStr2);
				rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr2);

				CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
				if(img)
				{
	//				img->SetHBITMAP(pMF->GetSignatureImage(szSendBy));
					img->SetFixedHeight(false);
				}
				rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(szSendBy));

				tmpStr = pMF->GetSysDate();	
				szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
				rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
				rpt.Print(bDirect);
				bDirect=true;
				rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("2"));
				rpt.Print(bDirect);
			}
			
			rst.MoveNext();
		}
	}
	else if (szType == _T("CRO"))
		PrintDRO(nOID);
	else
	{
		if(!rpt.Init(szReportName) )
			return;
		//Report Header
		rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
		rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
		rs.GetValue(_T("mt_orderdate"), tmpStr);
		szDate.Format(rpt.GetReportHeader()->GetValue(_T("OrderDate")), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);
		rpt.GetReportHeader()->SetValue(_T("OrderID"), szID);
		rs.GetValue(_T("mt_storage_name"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("FromStock"), tmpStr);
		rs.GetValue(_T("mt_storage_to_name"), tmpStr);
		if(!tmpStr.IsEmpty())
			rpt.GetReportHeader()->SetValue(_T("ToStock"), tmpStr);

		rs.GetValue(_T("mt_department_to_name"), tmpStr);
		if(!tmpStr.IsEmpty())
			rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);
		tmpStr.Empty();
		rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("1"));
		
		rs.GetValue(_T("mt_department_to_name"), tmpStr);
		if(!tmpStr.IsEmpty())
			rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);

		rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("1"));

		rpt.GetReportHeader()->SetValue(_T("Floor"), szFloor);
		rpt.GetReportHeader()->SetValue(_T("Floor2"), szFloor);

		szSQL.Format(_T("SELECT MAX(TRUNC_DATE(hpo_orderdate)) as maxorderdate, MIN(TRUNC_DATE(hpo_orderdate)) as minorderdate ") \
			_T(" FROM hms_medical_transaction_view ")\
			_T(" LEFT JOIN hms_ipharmaorder ON(hmt_orderid=hpo_orderid) ") \
			_T(" WHERE hmt_reftransaction_id=%ld"), nOID);
		//_fmsg(_T("%s"), szSQL);
		rsm.ExecSQL(szSQL);
		if(!rsm.IsEOF()){
			rpt.GetReportHeader()->SetValue(_T("MaxOrderDate"), CDate::Convert(rsm.GetValue(_T("MaxOrderDate")), yyyymmdd, ddmmyyyy));
			rpt.GetReportHeader()->SetValue(_T("MinOrderDate"), CDate::Convert(rsm.GetValue(_T("MinOrderDate")), yyyymmdd, ddmmyyyy));
		}

		//Page Header
		//Report Detail
		szSQL.Format(_T(" SELECT product_id                     AS id,") \
_T("   product_name                        AS name,") \
_T("   product_uomname                    AS unit,") \
_T("	sum(mtl_qtyorder) as qtyorder, ") \
_T("	sum(mtl_qtysold) as qtysold, ") \
_T("	sum(mtl_qtypolicy) as qtypolicy, ") \
_T("	sum(mtl_qtysoldins) as qtysoldins, ") \
_T("	sum(mtl_qtyotherins) as qtyotherins, ") \
_T("	sum(mtl_qtyservice) as qtyservice, ") \
_T("	sum(mtl_qtyother) as qtyother, ") \
_T("   mtl_saleprice                   AS price,") \
_T("   SUM(mtl_qtyorder*mtl_saleprice) AS amount,") \
_T("   product_countryname                AS hpol_productcountry") \
_T(" FROM m_transactionline") \
_T(" LEFT JOIN m_productitem_view") \
_T(" ON(product_item_id      =mtl_product_item_id)") \
_T(" WHERE mtl_transaction_id=%ld ") \
_T(" GROUP BY product_id,") \
_T("   product_name,") \
_T("   product_uomname,") \
_T("   mtl_saleprice,") \
_T("   product_countryname") \
_T(" ORDER BY name,  unit, price") ,	nOID, szOrderBy);

		rs2.ExecSQL(szSQL);

		CReportSection* rptDetail = rpt.GetDetail(0);
		while(!rs2.IsEOF())
		{
			rptDetail = rpt.AddDetail();
			tmpStr.Format(_T("%d"), nTotalItem++);
			rptDetail->SetValue(_T("1"), tmpStr);
			rs2.GetValue(_T("name"), tmpStr);
			rptDetail->SetValue(_T("2"), tmpStr);
			rs2.GetValue(_T("unit"), tmpStr);
			
			rptDetail->SetValue(_T("3"), tmpStr);
			
			rs2.GetValue(_T("qty"), nQty);
			FormatCurrency(nQty, tmpStr2);
			
			FormatCurrency(nQty, tmpStr2);
			rptDetail->SetValue(_T("4"), tmpStr2);

			rs2.GetValue(_T("qtysold"), nQty);
			FormatCurrency(nQty, tmpStr2);
			rptDetail->SetValue(_T("4.1"), tmpStr2);

			rs2.GetValue(_T("qtypolicy"), nQty);
			FormatCurrency(nQty, tmpStr2);
			rptDetail->SetValue(_T("4.2"), tmpStr2);

			rs2.GetValue(_T("qtysoldins"), nQty);
			FormatCurrency(nQty, tmpStr2);
			rptDetail->SetValue(_T("4.3"), tmpStr2);
			rs2.GetValue(_T("qtyotherins"), nQty);
			FormatCurrency(nQty, tmpStr2);
			rptDetail->SetValue(_T("4.4"), tmpStr2);
			
			rs2.GetValue(_T("qtyservice"), nQty);
			FormatCurrency(nQty, tmpStr2);
			rptDetail->SetValue(_T("4.5"), tmpStr2);

			rs2.GetValue(_T("qtyother"), nQty);
			FormatCurrency(nQty, tmpStr2);
			rptDetail->SetValue(_T("4.6"), tmpStr2);

			rs2.GetValue(_T("Price"), tmpStr);
			FormatCurrencyStr(tmpStr, tmpStr2);
			rptDetail->SetValue(_T("5"), tmpStr2);
			rs2.GetValue(_T("amount"), nAmount);
			nTotalAmount += nAmount;
			tmpStr.Empty();
			rptDetail->SetValue(_T("6"), tmpStr);
			rs2.MoveNext();
		}
		//Page Footer
		//Report Footer
		tmpStr.Format(_T("%d"), nTotalItem-1);
		rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
		if(nPrintStorage == 1)
			tmpStr.Format(_T("%.0f"), nTotalAmountServ);
		else
			tmpStr.Format(_T("%.0f"), nTotalAmount);
		FormatCurrencyStr(tmpStr, tmpStr2);
		rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr2);
		

		MoneyToString(tmpStr, tmpStr2);
		rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr2);

		tmpStr = pMF->GetSysDate();	
		szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
		rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
		

		CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
		if(img)
		{
	//		img->SetHBITMAP(pMF->GetSignatureImage(szSendBy));
			img->SetFixedHeight(false);
		}
		rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(szSendBy));

/*
		CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
		if(img)
		{
			img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
			img->SetFixedHeight(false);
		}
*/
		rpt.PrintPreview();
		//rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("2"));
		//rpt.Print(bDirect);
		
	}
}

void CPrintForms::PM_PrintCabinetSheet(long nOID)
{
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rsReport(&pMF->m_db);
	CRecord rs(&pMF->m_db);
	CRecord rsm(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CString tmpStr, tmpStr2, szDate, szSQL,szDesc, szOrderBy;
	CString szFloor;
	int nStorage_id =0;
	CString szID;
	long nProductID;


	szID.Format(_T("%ld"), nOID);

	double	nTotalAmount = 0, nAmount=0, nAmountServ = 0, nTotalAmountServ =0, nQty = 0;
	int		nTotalItem=1;
	CString szPrintStorage;
	int		nPrint=0;
	bool bDirect = false;
	
	szOrderBy.Format(_T(" ORDER BY name, unit, price "));
	szSQL.Format(_T("SELECT mt_orderno, mt_orderdate, mt_approveddate, mt_postedby, mt_storage_id, ") \
				_T(" get_storagename(mt_storage_id) as mt_storage_name, ") \
				_T(" get_storagename(mt_storage_to_id) as mt_storageto_name, ") \
				_T(" get_departmentname(mt_department_id) as mt_department_name, ") \
				_T(" get_departmentname(mt_department_to_id) as mt_departmentto_name, ") \
				_T(" mt_doctype, mt_description ") \
				_T("FROM m_transaction ") \
				_T("WHERE mt_transaction_id=%ld "), nOID);
	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	CString szType, szReportName, szSendBy;
	rs.GetValue(_T("mt_doctype"), szType);
	rs.GetValue(_T("mt_orderno"), szID);
	
	szOrderBy.Format(_T(" ORDER BY product_uomindex, product_name, product_uomname "));

	rs.GetValue(_T("mt_postedby"), szSendBy);
	rs.GetValue(_T("mt_description"), szDesc);
	
	
	rs.GetValue(_T("mt_floor"), tmpStr);
	//szFloor = pMF->GetSelectionString(_T("hms_floor"), tmpStr);
	szFloor.Empty();

	
	if(szType == _T("DPO") || szType == _T("DMO") || szType == _T("DDO") || szType == _T("CSO"))
	{
		CRecord rst(&pMF->m_db);
		CRecord rsp(&pMF->m_db);
		CString szGroup, szTitle, szPrintType, szNote;
		szSQL.Format(_T(" SELECT DISTINCT pmdt_id, pmdt_name,") \
					_T("   pmdt_desc,") \
					_T("   pmdt_reportname, pmdt_index,pmdt_printtype  ") \
					_T(" FROM pms_drugtype,") \
					_T("   m_transactionline,") \
					_T("   m_productitem_view") \
					_T(" WHERE mtl_transaction_id =%ld") \
					_T(" AND mtl_product_item_id  =product_item_id") \
					_T(" AND instr(pmdt_desc, product_producttype) > 0") \
					_T(" ORDER BY pmdt_index "), nOID);
		rst.ExecSQL(szSQL);
		while(!rst.IsEOF())
		{
			rst.GetValue(_T("pmdt_name"), tmpStr);
			StringUpper(tmpStr, szTitle);
			rst.GetValue(_T("pmdt_desc"), szGroup);
			szGroup.Trim();
			rst.GetValue(_T("pmdt_printtype"), szPrintType);


			//Kiem tra neu la kho thiet lap cho doi tuong DICH VU thi goi mau rieng
			rs.GetValue(_T("mt_storage_id"), nStorage_id);
			//szSQL.Format(_T(" SELECT msd_printtype FROM m_storage_default WHERE msd_storage_id = %d "), nStorage_id);
			szSQL.Format(_T("SELECT msl_category FROM m_storagelist WHERE msl_storage_id = %d"), nStorage_id);
			rsReport.ExecSQL(szSQL);
			
			szPrintStorage = rsReport.GetStringValue();

			tmpStr = _T("PMS_BOSUNGTUTRUC");
			if(szPrintStorage == _T("I"))
				tmpStr += _T("_BH");
			else if (szPrintStorage == _T("S"))
				tmpStr += _T("_DV");
			else if (szPrintStorage == _T("P"))
				tmpStr += _T("_QUAN");
			rst.GetValue(_T("pmdt_id"), tmpStr2);
			if (tmpStr2 == _T("YC") || tmpStr2 == _T("VT"))
				tmpStr += _T("TB");
			else if (tmpStr2 == _T("GN") || tmpStr2 == _T("HT"))
				tmpStr += _T("GNHT");
			szReportName.Format(_T("Reports/HMS/%s.RPT"), tmpStr);
//_msg(_T("%s, %s"), szSQL, szReportName);
			if(!rpt.Init(szReportName) )
				return;   			
		
			//Report Header
			rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
			rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
			rpt.GetReportHeader()->Format(_T("TITLE"), szTitle);

			rs.GetValue(_T("mt_orderdate"), tmpStr);
			szDate.Format(rpt.GetReportHeader()->GetValue(_T("OrderDate")), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
			rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);
			rpt.GetReportHeader()->SetValue(_T("OrderID"), szID);
			rs.GetValue(_T("mt_statusdesc"), tmpStr);
			//tmpStr = GetStatusString(tmpStr);
			rpt.GetReportHeader()->SetValue(_T("Status"), tmpStr);
			rs.GetValue(_T("mt_storage_name"), tmpStr);
			//tmpStr = GetStorageName(ToInt(tmpStr));
			rpt.GetReportHeader()->SetValue(_T("FromStock"), tmpStr);
			rpt.GetReportHeader()->SetValue(_T("Comment"), szDesc);
			
			rs.GetValue(_T("mt_storageto_name"), tmpStr);
			if(!tmpStr.IsEmpty())
				rpt.GetReportHeader()->SetValue(_T("ToStock"), tmpStr);

			rs.GetValue(_T("mt_department_name"), tmpStr);
			if(!tmpStr.IsEmpty())
				rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);

			rs.GetValue(_T("mt_departmentto_name"), tmpStr);
			if(!tmpStr.IsEmpty())
			{
				//tmpStr = GetDepartmentName(tmpStr);
				tmpStr.Append(_T(" - "));
				tmpStr.Append(rs.GetValue(_T("mt_storageto_name")));
				rpt.GetReportHeader()->SetValue(_T("DepartmentTo"), tmpStr);
			}
			tmpStr.Empty();
			rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("1"));

			rpt.GetReportHeader()->SetValue(_T("Floor"), szFloor);

			szSQL.Format(_T("SELECT MAX(TRUNC_DATE(hpo_orderdate)) as maxorderdate, MIN(TRUNC_DATE(hpo_orderdate)) as minorderdate ") \
				_T(" FROM hms_medical_transaction_view ")\
				_T(" LEFT JOIN hms_ipharmaorder ON(hmt_orderid=hpo_orderid) ") \
				_T(" WHERE hmt_reftransaction_id=%ld"), nOID);
			_fmsg(_T("%s"), szSQL);
			rsm.ExecSQL(szSQL);
			if(!rsm.IsEOF()){
				rpt.GetReportHeader()->SetValue(_T("MaxOrderDate"), CDate::Convert(rsm.GetValue(_T("MaxOrderDate")), yyyymmdd, ddmmyyyy));
				rpt.GetReportHeader()->SetValue(_T("MinOrderDate"), CDate::Convert(rsm.GetValue(_T("MinOrderDate")), yyyymmdd, ddmmyyyy));
			}			
			//Page Header
			//Report Detail
			szSQL.Format(_T(" SELECT ") \
					_T(" 	product_id as id,") \
					_T(" mtl_xproduct_id as xproduct_id, ") \
					_T(" 	product_name as name,") \
					_T(" 	product_uomname	as unit,") \
					_T("	sum(mtl_qtydelivery) as qtyrealorder, ") \
					_T("	sum(mtl_qtysold) as qtysold, ") \
					_T("	sum(mtl_qtypolicy) as qtypolicy, ") \
					_T("	sum(mtl_qtysoldins) as qtysoldins, ") \
					_T("	sum(mtl_qtyotherins) as qtyotherins, ") \
					_T("	sum(mtl_qtyservice) as qtyservice, ") \
					_T("	sum(mtl_qtyother) as qtyother, ") \
					_T("	sum(mtl_qtyorder) qtyorder,") \
					_T(" 	mtl_saleprice as price,") \
					_T(" 	sum(mtl_qtyorder*mtl_saleprice) as amount, ") \
					_T(" 	product_countryname as hpol_productcountry ") \
					_T(" FROM m_transactionline") \
					_T(" LEFT JOIN m_productitem_view ON(product_item_id=mtl_product_item_id)") \
					_T(" WHERE mtl_transaction_id=%ld ") \
					_T(" AND product_producttype in(%s) ") \
					_T(" and nvl(mtl_client_id,'XXX') <> 'REP' ") \
					_T(" GROUP BY product_id, mtl_xproduct_id, product_name, product_uomname, mtl_saleprice, product_countryname, product_uomindex ") \
					_T(" %s "), nOID, szGroup, szOrderBy);
			rs2.ExecSQL(szSQL);
			if(rs2.IsEOF()){
				rst.MoveNext();
				continue;

			}
			CReportSection* rptDetail = rpt.GetDetail(0); 
			rpt.ResetContent();
			nTotalItem=1;
			nTotalAmount = 0;
			CString szItemId;
			long nProductID, nXProductID;
			CString szName, szName2;
			CRecord rsx(&pMF->m_db);
			while(!rs2.IsEOF())
			{
				rptDetail = rpt.AddDetail();
				rs2.GetValue(_T("id"), nProductID);
				tmpStr.Format(_T("%d"), nTotalItem++);
				rptDetail->SetValue(_T("1"), tmpStr);
				rs2.GetValue(_T("name"), szName);
				rs2.GetValue(_T("id"), nProductID);
				rs2.GetValue(_T("xproduct_id"), nXProductID);
				if (nProductID != nXProductID)
				{
					tmpStr.Format(_T("SELECT mp_name FROM m_product WHERE mp_product_id = %ld "), nXProductID);
					rsx.ExecSQL(tmpStr);
					rsx.GetValue(_T("mp_name"), szName2);
					szName.AppendFormat(_T(" [%s] "), szName2);
				}
				rptDetail->SetValue(_T("2"), szName);
				rs2.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("3"), tmpStr);
				rs2.GetValue(_T("qtyorder"), nQty);

				if(szGroup.Find(_T("'A1130'")) != -1 || szGroup.Find(_T("'A1140'")) != -1 || szGroup.Find(_T("'A2000'")) != -1)
					MoneyToString(ToString(nQty), szNote);
				else
					szNote.Empty();
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4"), tmpStr2);

				rs2.GetValue(_T("qtysold"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.1"), tmpStr2);

				rs2.GetValue(_T("qtypolicy"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.2"), tmpStr2);
				
				rs2.GetValue(_T("qtysoldins"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.3"), tmpStr2);
				rs2.GetValue(_T("qtyotherins"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.4"), tmpStr2);
				
				rs2.GetValue(_T("qtyservice"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.5"), tmpStr2);

				rs2.GetValue(_T("qtyother"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.6"), tmpStr2);

				
				rs2.GetValue(_T("qtyrealorder"), tmpStr);
				FormatCurrencyStr(tmpStr, tmpStr2);
				rptDetail->SetValue(_T("5"), tmpStr2);
				rs2.GetValue(_T("amount"), nAmount);
				nTotalAmount += nAmount;

				FormatCurrency(nAmount, tmpStr2);
				rptDetail->SetValue(_T("7"), tmpStr2);
				
				rptDetail->SetValue(_T("6"), szNote);

				rs2.GetValue(_T("price"), tmpStr);
				FormatCurrencyStr(tmpStr, tmpStr2);
				rptDetail->SetValue(_T("8"), tmpStr2);
				rs2.GetValue(_T("amount"), nAmountServ);
				nTotalAmountServ += nAmountServ;

				FormatCurrency(nAmountServ, tmpStr2);
				rptDetail->SetValue(_T("9"), tmpStr2);
				
				if(szGroup.Find(_T("'A1130'")) != -1 || szGroup.Find(_T("'A1140'")) != -1 || szGroup.Find(_T("'A2000'")) != -1)
				{
					szSQL.Format(_T(" SELECT 	hmt_docno as docno,") \
								_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
								_T(" 	sum(hmt_qtyissue) as orderqty") \
								_T(" FROM hms_medical_transaction_view") \
								_T(" LEFT JOIN hms_doc ON(hd_docno=hmt_docno)") \
								_T(" LEFT JOIN hms_patient ON(hp_patientno=hd_patientno)") \
								_T(" WHERE hmt_reftransaction_id=%ld ") \
								_T(" and hmt_product_id=%ld ") \
								_T(" GROUP BY hmt_docno, hp_firstname, hp_midname, hp_surname ") \
								_T(" ORDER BY pname"), nOID, nProductID);
					rsp.ExecSQL(szSQL);
					while(!rsp.IsEOF()){
						rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
/*
						rptDetail->GetItem(_T("1"))->SetItalic(true);	
						rptDetail->GetItem(_T("2"))->SetItalic(true);	
						rptDetail->GetItem(_T("3"))->SetItalic(true);	
						rptDetail->GetItem(_T("4"))->SetItalic(true);	
						rptDetail->GetItem(_T("5"))->SetItalic(true);	
						rptDetail->GetItem(_T("6"))->SetItalic(true);	
*/
						rptDetail->SetValue(_T("1"), _T(""));
						rsp.GetValue(_T("pname"), tmpStr);
						rptDetail->SetValue(_T("2"), tmpStr);
						rsp.GetValue(_T("docno"), tmpStr);
						rptDetail->SetValue(_T("3"), tmpStr);
						rsp.GetValue(_T("orderqty"), nQty);
						FormatCurrency(nQty, tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);
						rptDetail->SetValue(_T("5"), _T(""));
						MoneyToString(ToString(nQty), szNote);
						rptDetail->SetValue(_T("6"), szNote);
/*
						rptDetail->GetItem(_T("12"))->SetItalic(true);	
						rptDetail->GetItem(_T("22"))->SetItalic(true);	
						rptDetail->GetItem(_T("32"))->SetItalic(true);	
						rptDetail->GetItem(_T("42"))->SetItalic(true);	
						rptDetail->GetItem(_T("52"))->SetItalic(true);	
						rptDetail->GetItem(_T("62"))->SetItalic(true);	
*/
						rsp.MoveNext();
					}
				}
				rs2.MoveNext();
			}
			//Page Footer
			//Report Footer
			tmpStr.Format(_T("%d"), nTotalItem-1);
			rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
			tmpStr.Format(_T("%.0f"), nTotalAmount);
			FormatCurrencyStr(tmpStr, tmpStr2);
			rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr2);

			MoneyToString(tmpStr, tmpStr2);
			rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr2);

			CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
			if(img)
			{
	//			img->SetHBITMAP(pMF->GetSignatureImage(szSendBy));
	//			img->SetFixedHeight(false);
			}

			rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(szSendBy));

			tmpStr = pMF->GetSysDate();	
			szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
			rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
		//	rpt.Print(bDirect);
		//	bDirect=true;
			rpt.PrintPreview();
			rst.MoveNext();
		}
	}
}





void CPrintForms::PM_PrintExportCabinetSheet(long nOID)
{
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rsReport(&pMF->m_db);
	CRecord rs(&pMF->m_db);
	CRecord rsm(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CString tmpStr, tmpStr2, szDate, szSQL,szDesc, szOrderBy;
	CString szFloor;
	int nStorage_id =0;
	CString szID;
	long nProductID;


	szID.Format(_T("%ld"), nOID);

	double	nTotalAmount = 0, nAmount=0, nAmountServ = 0, nTotalAmountServ =0, nQty = 0;
	int		nTotalItem=1;
	CString szPrintStorage;
	int		nPrint=0;
	bool bDirect = false;
	
	szOrderBy.Format(_T(" ORDER BY name, unit, price "));
	szSQL.Format(_T("SELECT mt_orderno, mt_orderdate, mt_approveddate, mt_postedby, mt_storage_id, ") \
				_T(" get_storagename(mt_storage_id) as mt_storage_name, ") \
				_T(" get_storagename(mt_storage_to_id) as mt_storageto_name, ") \
				_T(" get_departmentname(mt_department_id) as mt_department_name, ") \
				_T(" get_departmentname(mt_department_to_id) as mt_departmentto_name, ") \
				_T(" mt_doctype, mt_description ") \
				_T("FROM m_transaction ") \
				_T("WHERE mt_transaction_id=%ld "), nOID);
	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	CString szType, szReportName, szSendBy;
	rs.GetValue(_T("mt_doctype"), szType);
	rs.GetValue(_T("mt_orderno"), szID);
	
	szOrderBy.Format(_T(" ORDER BY product_uomindex, product_name, product_uomname "));

	rs.GetValue(_T("mt_postedby"), szSendBy);
	rs.GetValue(_T("mt_description"), szDesc);
	
	
	rs.GetValue(_T("mt_floor"), tmpStr);
	//szFloor = pMF->GetSelectionString(_T("hms_floor"), tmpStr);
	szFloor.Empty();

	
	if(szType == _T("DPO") || szType == _T("DMO") || szType == _T("DDO") || szType == _T("CSO"))
	{
		CRecord rst(&pMF->m_db);
		CRecord rsp(&pMF->m_db);
		CString szGroup, szTitle, szPrintType, szNote;
		szSQL.Format(_T(" SELECT DISTINCT pmdt_id, pmdt_name,") \
					_T("   pmdt_desc,") \
					_T("   pmdt_reportname, pmdt_index,pmdt_printtype  ") \
					_T(" FROM pms_drugtype,") \
					_T("   m_transactionline,") \
					_T("   m_productitem_view") \
					_T(" WHERE mtl_transaction_id =%ld") \
					_T(" AND mtl_product_item_id  =product_item_id") \
					_T(" AND instr(pmdt_desc, product_producttype) > 0") \
					_T(" ORDER BY pmdt_index "), nOID);
		rst.ExecSQL(szSQL);
		while(!rst.IsEOF())
		{
			rst.GetValue(_T("pmdt_name"), tmpStr);
			StringUpper(tmpStr, szTitle);
			rst.GetValue(_T("pmdt_desc"), szGroup);
			szGroup.Trim();
			rst.GetValue(_T("pmdt_printtype"), szPrintType);


			//Kiem tra neu la kho thiet lap cho doi tuong DICH VU thi goi mau rieng
			rs.GetValue(_T("mt_storage_id"), nStorage_id);
			//szSQL.Format(_T(" SELECT msd_printtype FROM m_storage_default WHERE msd_storage_id = %d "), nStorage_id);
			szSQL.Format(_T("SELECT msl_category FROM m_storagelist WHERE msl_storage_id = %d"), nStorage_id);
			rsReport.ExecSQL(szSQL);
			
			szPrintStorage = rsReport.GetStringValue();

			tmpStr = _T("PMS_BOSUNGTUTRUC");
			if(szPrintStorage == _T("I"))
				tmpStr += _T("_BH");
			else if (szPrintStorage == _T("S"))
				tmpStr += _T("_DV");
			else if (szPrintStorage == _T("P"))
				tmpStr += _T("_QUAN");
			rst.GetValue(_T("pmdt_id"), tmpStr2);
			if (tmpStr2 == _T("YC") || tmpStr2 == _T("VT"))
				tmpStr += _T("TB");
			else if (tmpStr2 == _T("GN") || tmpStr2 == _T("HT"))
				tmpStr += _T("GNHT");
			szReportName.Format(_T("Reports/HMS/%s.RPT"), tmpStr);
//_msg(_T("%s, %s"), szSQL, szReportName);
			if(!rpt.Init(szReportName) )
				return;   			
		
			//Report Header
			rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
			rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
			rpt.GetReportHeader()->Format(_T("TITLE"), szTitle);

			rs.GetValue(_T("mt_orderdate"), tmpStr);
			szDate.Format(rpt.GetReportHeader()->GetValue(_T("OrderDate")), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
			rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);
			rpt.GetReportHeader()->SetValue(_T("OrderID"), szID);
			rs.GetValue(_T("mt_statusdesc"), tmpStr);
			//tmpStr = GetStatusString(tmpStr);
			rpt.GetReportHeader()->SetValue(_T("Status"), tmpStr);
			rs.GetValue(_T("mt_storage_name"), tmpStr);
			//tmpStr = GetStorageName(ToInt(tmpStr));
			rpt.GetReportHeader()->SetValue(_T("FromStock"), tmpStr);
			rpt.GetReportHeader()->SetValue(_T("Comment"), szDesc);
			
			rs.GetValue(_T("mt_storageto_name"), tmpStr);
			if(!tmpStr.IsEmpty())
				rpt.GetReportHeader()->SetValue(_T("ToStock"), tmpStr);

			rs.GetValue(_T("mt_department_name"), tmpStr);
			if(!tmpStr.IsEmpty())
				rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);

			rs.GetValue(_T("mt_departmentto_name"), tmpStr);
			if(!tmpStr.IsEmpty())
			{
				//tmpStr = GetDepartmentName(tmpStr);
				tmpStr.Append(_T(" - "));
				tmpStr.Append(rs.GetValue(_T("mt_storageto_name")));
				rpt.GetReportHeader()->SetValue(_T("DepartmentTo"), tmpStr);
			}
			tmpStr.Empty();
			rpt.GetReportHeader()->SetValue(_T("Sheet"), _T("1"));

			rpt.GetReportHeader()->SetValue(_T("Floor"), szFloor);

			szSQL.Format(_T("SELECT MAX(TRUNC_DATE(hpo_orderdate)) as maxorderdate, MIN(TRUNC_DATE(hpo_orderdate)) as minorderdate ") \
				_T(" FROM hms_medical_transaction_view ")\
				_T(" LEFT JOIN hms_ipharmaorder ON(hmt_orderid=hpo_orderid) ") \
				_T(" WHERE hmt_reftransaction_id=%ld"), nOID);
			_fmsg(_T("%s"), szSQL);
			rsm.ExecSQL(szSQL);
			if(!rsm.IsEOF()){
				rpt.GetReportHeader()->SetValue(_T("MaxOrderDate"), CDate::Convert(rsm.GetValue(_T("MaxOrderDate")), yyyymmdd, ddmmyyyy));
				rpt.GetReportHeader()->SetValue(_T("MinOrderDate"), CDate::Convert(rsm.GetValue(_T("MinOrderDate")), yyyymmdd, ddmmyyyy));
			}			
			//Page Header
			//Report Detail
			szSQL.Format(_T(" SELECT mtl_transactionline_id, NVL(mtl_client_id,'XX') as client_id, ") \
					_T(" 	product_id as id,") \
					_T(" mtl_xproduct_id as xproduct_id, ") \
					_T(" 	product_name as name,") \
					_T(" 	product_uomname	as unit,") \
					_T("	sum(mtl_qtydelivery) as qtyrealorder, ") \
					_T("	sum(mtl_qtysold) as qtysold, ") \
					_T("	sum(mtl_qtypolicy) as qtypolicy, ") \
					_T("	sum(mtl_qtysoldins) as qtysoldins, ") \
					_T("	sum(mtl_qtyotherins) as qtyotherins, ") \
					_T("	sum(mtl_qtyservice) as qtyservice, ") \
					_T("	sum(mtl_qtyother) as qtyother, ") \
					_T("	sum(mtl_qtyorder) qtyorder,") \
					_T(" 	mtl_saleprice as price,") \
					_T(" 	sum(mtl_qtyorder*mtl_saleprice) as amount, ") \
					_T(" 	product_countryname as hpol_productcountry ") \
					_T(" FROM m_transactionline") \
					_T(" LEFT JOIN m_productitem_view ON(product_item_id=mtl_product_item_id)") \
					_T(" WHERE mtl_transaction_id=%ld ") \
					_T(" AND product_producttype in(%s) and NVL(mtl_client_id, 'XX') <> 'REP' ") \
					_T(" GROUP BY mtl_transactionline_id, mtl_client_id, product_id, mtl_xproduct_id, product_name, product_uomname, mtl_saleprice, product_countryname, product_uomindex ") \
					_T(" %s "), nOID, szGroup, szOrderBy);
			rs2.ExecSQL(szSQL);
			if(rs2.IsEOF()){
				rst.MoveNext();
				continue;

			}
			CReportSection* rptDetail = rpt.GetDetail(0); 
			rpt.ResetContent();
			nTotalItem=1;
			nTotalAmount = 0;
			CString szItemId;
			long nProductID, nXProductID;
			CString szClientID;
			long	nTransactionLine_ID;

			CString szName, szName2;
			CRecord rsx(&pMF->m_db);
			CRecord rsl(&pMF->m_db);
			while(!rs2.IsEOF())
			{
				rptDetail = rpt.AddDetail();
				rs2.GetValue(_T("id"), nProductID);
				tmpStr.Format(_T("%d"), nTotalItem++);
				rptDetail->SetValue(_T("1"), tmpStr);
				rs2.GetValue(_T("name"), szName);
				rs2.GetValue(_T("id"), nProductID);
				rs2.GetValue(_T("xproduct_id"), nXProductID);
				rs2.GetValue(_T("client_id"), szClientID);
				if (nProductID != nXProductID)
				{
					tmpStr.Format(_T("SELECT mp_name FROM m_product WHERE mp_product_id = %ld "), nXProductID);
					rsx.ExecSQL(tmpStr);
					rsx.GetValue(_T("mp_name"), szName2);
					szName.AppendFormat(_T(" [%s] "), szName2);
				}
				rptDetail->SetValue(_T("2"), szName);
				rs2.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("3"), tmpStr);
				rs2.GetValue(_T("qtyorder"), nQty);
				if(szClientID == _T("TT"))
				{
					rs2.GetValue(_T("qtyrealorder"), nQty);

				}

				if(szGroup.Find(_T("'A1130'")) != -1 || szGroup.Find(_T("'A1140'")) != -1 || szGroup.Find(_T("'A2000'")) != -1)
					MoneyToString(ToString(nQty), szNote);
				else
					szNote.Empty();
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4"), tmpStr2);

				rs2.GetValue(_T("qtysold"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.1"), tmpStr2);

				rs2.GetValue(_T("qtypolicy"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.2"), tmpStr2);
				
				rs2.GetValue(_T("qtysoldins"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.3"), tmpStr2);
				rs2.GetValue(_T("qtyotherins"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.4"), tmpStr2);
				
				rs2.GetValue(_T("qtyservice"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.5"), tmpStr2);

				rs2.GetValue(_T("qtyother"), nQty);
				FormatCurrency(nQty, tmpStr2);
				rptDetail->SetValue(_T("4.6"), tmpStr2);

				
				rs2.GetValue(_T("qtyrealorder"), tmpStr);
				FormatCurrencyStr(tmpStr, tmpStr2);
				rptDetail->SetValue(_T("5"), tmpStr2);
				rs2.GetValue(_T("amount"), nAmount);
				nTotalAmount += nAmount;

				FormatCurrency(nAmount, tmpStr2);
				rptDetail->SetValue(_T("7"), tmpStr2);
				
				rptDetail->SetValue(_T("6"), szNote);

				rs2.GetValue(_T("price"), tmpStr);
				FormatCurrencyStr(tmpStr, tmpStr2);
				
				rs2.GetValue(_T("amount"), nAmountServ);


				if(szClientID == _T("TT"))
				{
					nAmountServ = 0;
					tmpStr2.Empty();

				}
				rptDetail->SetValue(_T("8"), tmpStr2);

				nTotalAmountServ += nAmountServ;

				FormatCurrency(nAmountServ, tmpStr2);
				rptDetail->SetValue(_T("9"), tmpStr2);
				
				if(szGroup.Find(_T("'A1130'")) != -1 || szGroup.Find(_T("'A1140'")) != -1 || szGroup.Find(_T("'A2000'")) != -1)
				{
					szSQL.Format(_T(" SELECT 	hmt_docno as docno,") \
								_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
								_T(" 	sum(hmt_qtyissue) as orderqty") \
								_T(" FROM hms_medical_transaction_view") \
								_T(" LEFT JOIN hms_doc ON(hd_docno=hmt_docno)") \
								_T(" LEFT JOIN hms_patient ON(hp_patientno=hd_patientno)") \
								_T(" WHERE hmt_reftransaction_id=%ld ") \
								_T(" and hmt_product_id=%ld ") \
								_T(" GROUP BY hmt_docno, hp_firstname, hp_midname, hp_surname ") \
								_T(" ORDER BY pname"), nOID, nProductID);
					rsp.ExecSQL(szSQL);
					while(!rsp.IsEOF()){
						rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
/*
						rptDetail->GetItem(_T("1"))->SetItalic(true);	
						rptDetail->GetItem(_T("2"))->SetItalic(true);	
						rptDetail->GetItem(_T("3"))->SetItalic(true);	
						rptDetail->GetItem(_T("4"))->SetItalic(true);	
						rptDetail->GetItem(_T("5"))->SetItalic(true);	
						rptDetail->GetItem(_T("6"))->SetItalic(true);	
*/
						rptDetail->SetValue(_T("1"), _T(""));
						rsp.GetValue(_T("pname"), tmpStr);
						rptDetail->SetValue(_T("2"), tmpStr);
						rsp.GetValue(_T("docno"), tmpStr);
						rptDetail->SetValue(_T("3"), tmpStr);
						rsp.GetValue(_T("orderqty"), nQty);
						FormatCurrency(nQty, tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);
						rptDetail->SetValue(_T("5"), _T(""));
						MoneyToString(ToString(nQty), szNote);
						rptDetail->SetValue(_T("6"), szNote);
/*
						rptDetail->GetItem(_T("12"))->SetItalic(true);	
						rptDetail->GetItem(_T("22"))->SetItalic(true);	
						rptDetail->GetItem(_T("32"))->SetItalic(true);	
						rptDetail->GetItem(_T("42"))->SetItalic(true);	
						rptDetail->GetItem(_T("52"))->SetItalic(true);	
						rptDetail->GetItem(_T("62"))->SetItalic(true);	
*/
						rsp.MoveNext();
					}
				}

				rs2.GetValue(_T("client_id"), szClientID);
				if(szClientID == _T("TT"))
				{
					rs2.GetValue(_T("mtl_transactionline_id"), nTransactionLine_ID);


					szSQL.Format(_T(" SELECT mtl_transactionline_id, NVL(mtl_client_id,'XX') as client_id, ") \
					_T(" 	product_id as id,") \
					_T(" mtl_xproduct_id as xproduct_id, ") \
					_T(" 	product_name as name,") \
					_T(" 	product_uomname	as unit,") \
					_T("	sum(mtl_qtydelivery) as qtyrealorder, ") \
					_T("	sum(mtl_qtysold) as qtysold, ") \
					_T("	sum(mtl_qtypolicy) as qtypolicy, ") \
					_T("	sum(mtl_qtysoldins) as qtysoldins, ") \
					_T("	sum(mtl_qtyotherins) as qtyotherins, ") \
					_T("	sum(mtl_qtyservice) as qtyservice, ") \
					_T("	sum(mtl_qtyother) as qtyother, ") \
					_T("	sum(mtl_qtyorder) qtyorder,") \
					_T(" 	mtl_saleprice as price,") \
					_T(" 	sum(mtl_qtyorder*mtl_saleprice) as amount, ") \
					_T(" 	product_countryname as hpol_productcountry ") \
					_T(" FROM m_transactionline") \
					_T(" LEFT JOIN m_productitem_view ON(product_item_id=mtl_product_item_id)") \
					_T(" WHERE mtl_transaction_id=%ld ") \
					_T(" AND product_producttype in(%s) and NVL(mtl_client_id, 'XX')='REP' ") \
					_T(" and  mtl_transactionline_id in(select mtlr_reftransactionline_id from m_transactionline_ref where mtlr_transaction_id=%ld and mtlr_transactionline_id=%ld) ") \
					_T(" GROUP BY mtl_transactionline_id, mtl_client_id, product_id, mtl_xproduct_id, product_name, product_uomname, mtl_saleprice, product_countryname, product_uomindex ") \
					_T(" %s "), nOID, szGroup, nOID, nTransactionLine_ID, szOrderBy);

				rsl.ExecSQL(szSQL);
				while(!rsl.IsEOF())
				{
					rptDetail = rpt.AddDetail();
					rsl.GetValue(_T("id"), nProductID);
					tmpStr.Format(_T("%d"), nTotalItem++);
					rptDetail->SetValue(_T("1"), tmpStr);
					rsl.GetValue(_T("name"), szName);
					rsl.GetValue(_T("id"), nProductID);
					rsl.GetValue(_T("xproduct_id"), nXProductID);
					if (nProductID != nXProductID)
					{
						tmpStr.Format(_T("SELECT mp_name FROM m_product WHERE mp_product_id = %ld "), nXProductID);
						rsx.ExecSQL(tmpStr);
						rsx.GetValue(_T("mp_name"), szName2);
						szName.AppendFormat(_T(" [%s] "), szName2);
					}
					rptDetail->SetValue(_T("2"), szName);
					rsl.GetValue(_T("unit"), tmpStr);
					rptDetail->SetValue(_T("3"), tmpStr);
					rsl.GetValue(_T("qtyorder"), nQty);

					if(szGroup.Find(_T("'A1130'")) != -1 || szGroup.Find(_T("'A1140'")) != -1 || szGroup.Find(_T("'A2000'")) != -1)
						MoneyToString(ToString(nQty), szNote);
					else
						szNote.Empty();
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4"), tmpStr2);

					rsl.GetValue(_T("qtysold"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.1"), tmpStr2);

					rsl.GetValue(_T("qtypolicy"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.2"), tmpStr2);
					
					rsl.GetValue(_T("qtysoldins"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.3"), tmpStr2);
					rsl.GetValue(_T("qtyotherins"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.4"), tmpStr2);
					
					rsl.GetValue(_T("qtyservice"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.5"), tmpStr2);

					rsl.GetValue(_T("qtyother"), nQty);
					FormatCurrency(nQty, tmpStr2);
					rptDetail->SetValue(_T("4.6"), tmpStr2);

					
					rsl.GetValue(_T("qtyrealorder"), tmpStr);
					FormatCurrencyStr(tmpStr, tmpStr2);
					rptDetail->SetValue(_T("5"), tmpStr2);
					rsl.GetValue(_T("amount"), nAmount);
					nTotalAmount += nAmount;

					FormatCurrency(nAmount, tmpStr2);
					rptDetail->SetValue(_T("7"), tmpStr2);
					
					rptDetail->SetValue(_T("6"), szNote);

					rsl.GetValue(_T("price"), tmpStr);
					FormatCurrencyStr(tmpStr, tmpStr2);
					rptDetail->SetValue(_T("8"), tmpStr2);
					rsl.GetValue(_T("amount"), nAmountServ);
					nTotalAmountServ += nAmountServ;

					FormatCurrency(nAmountServ, tmpStr2);
					rptDetail->SetValue(_T("9"), tmpStr2);
				
				
					rsl.MoveNext();
				}


				}
				rs2.MoveNext();
			}
			//Page Footer
			//Report Footer
			tmpStr.Format(_T("%d"), nTotalItem-1);
			rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
			tmpStr.Format(_T("%.0f"), nTotalAmount);
			FormatCurrencyStr(tmpStr, tmpStr2);
			rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr2);

			MoneyToString(tmpStr, tmpStr2);
			rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr2);

			CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
			if(img)
			{
	//			img->SetHBITMAP(pMF->GetSignatureImage(szSendBy));
	//			img->SetFixedHeight(false);
			}

			rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(szSendBy));

			tmpStr = pMF->GetSysDate();	
			szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
			rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
		//	rpt.Print(bDirect);
		//	bDirect=true;
			rpt.PrintPreview();
			rst.MoveNext();
		}
	}
}




void CPrintForms::PrintPurchaseReturnOrder(long nOrderID, CString szFilter){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString szSQL, szWhere, tmpStr, szOrderDate, szDeliver, szReceiver;
	CRecord rs(&pMF->m_db);
	CReportSection *rptHeader, *rptDetail, *rptFooter;
	int nIdx = 0;
	double nS6 = 0, nS8 = 0;
	if (!rpt.Init(_T("Reports\\HMS\\PM_HANGMUATRALAI.RPT")))
		return;
	szWhere.AppendFormat(_T(" AND %s"), szFilter);
	szSQL.Format(_T(" SELECT") \
				_T("   mt_orderno       AS orderno,") \
				_T("   Cast_Timestamp2Date(mt_orderdate) as orderdate,") \
				_T("   mt_deliveryby    AS deliver,") \
				_T("   mt_receiveby		AS receiver,") \
				_T("   Cast_Timestamp2Date(mt_approveddate)  AS dte,") \
				_T("   get_storagename(mt_storage_id) AS stock,") \
				_T("   get_departmentname(mt_department_id) as dept,") \
				_T("   get_partnername(mt_partner_id) AS partner,") \
				_T("   mt_description   AS descr") \
				_T(" FROM   m_transaction") \
				_T(" WHERE mt_transaction_id = %ld"), nOrderID);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONASTERISK | MB_OK);
		return;
	}
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	rptHeader->SetValue(_T("dept"), rs.GetValue(_T("dept")));
	rptHeader->SetValue(_T("Invoiceno"), rs.GetValue(_T("orderno")));
	rs.GetValue(_T("deliver"), szDeliver);
	rs.GetValue(_T("receiver"), szReceiver);
	rptHeader->SetValue(_T("StockName"), rs.GetValue(_T("stock")));
	rptHeader->SetValue(_T("Reason"), rs.GetValue(_T("descr")));
	rs.GetValue(_T("orderdate"), szOrderDate);
	rs.GetValue(_T("dte"), tmpStr);
	rptHeader->SetValue(_T("ReturnDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	rptHeader->SetValue(_T("SupplierName"), rs.GetValue(_T("partner")));
	rptHeader->SetValue(_T("clerk"), pMF->GetCurrentUser());
	szSQL.Format(_T(" SELECT") \
				_T("   product_name   AS pname,") \
				_T("   mtl_qtydelivery	AS qty,") \
				_T("   product_uomname	AS uom,") \
				_T("   product_expdate	AS expdte,") \
				_T("   product_taxprice AS price,") \
				_T("   product_lotno	AS lot,") \
				_T("   mtl_qtydelivery*product_taxprice AS amt") \
				_T(" FROM m_transactionline") \
				_T(" LEFT JOIN m_productitem_view ON (product_item_id=mtl_product_item_id)") \
				_T(" WHERE mtl_transaction_id = %ld"), nOrderID);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;
	while (!rs.IsEOF())
	{
		nIdx++;
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("pname")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("uom")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("lot")));
		rs.GetValue(_T("expdte"), tmpStr);
		rptDetail->SetValue(_T("5"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		
		rs.GetValue(_T("qty"), tmpStr);
		nS6 += str2double(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("price")));
		
		rs.GetValue(_T("amt"), tmpStr);
		nS8 += str2double(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.MoveNext();
	}
	rptFooter = rpt.AddDetail(rpt.GetGroupFooter(1));
	rptFooter->SetValue(_T("s8"), double2str(nS8));
	tmpStr.Format(rptFooter->GetValue(_T("tongso")), nIdx);
	rptFooter->SetValue(_T("tongso"), tmpStr);
	MoneyToString(ToString(nS8), tmpStr);
	if (!tmpStr.IsEmpty())
	{
		tmpStr += _T(" \x111\x1ED3ng.");
		rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	}
	rptFooter = rpt.GetReportFooter();
	rptFooter->SetValue(_T("TotalItem"), int2str(nIdx));
	tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), szOrderDate.Mid(8, 2), szOrderDate.Mid(5,2), szOrderDate.Left(4));
	rptFooter->SetValue(_T("PrintDate"), tmpStr);
	rptFooter->SetValue(_T("User1"), _T(""));
	rptFooter->SetValue(_T("User2"), szDeliver);
	rptFooter->SetValue(_T("User3"), szReceiver);
	rptFooter->SetValue(_T("User4"), pMF->GetCurrentUser());
	rpt.PrintPreview();

}

void CPrintForms::PrintDRO(long nOrderID, CString szFilter, CString szDepttype){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString szSQL, szWhere, tmpStr, tmpStr2, szGroup, szReportName, szTitle, szObjectType, szStorageCategory, szNote;
	CRecord rs(&pMF->m_db), rst(&pMF->m_db), rsp(&pMF->m_db), rsReport(&pMF->m_db), rsl(&pMF->m_db);

	CReportSection *rptHeader, *rptDetail, *rptFooter;
	int nTotalItem = 0, nBreak = 0, nStorage_id = 0, nQty = 0;
	double nS6 = 0, nS8 = 0;
	long nProductID = 0;
	CString szSysDte;
	szWhere.AppendFormat(_T(" AND %s"), szFilter);
	szSQL.Format(_T(" SELECT") \
				_T("   mt_orderno       AS orderno,") \
				_T("   mt_storage_to_id,") \
				_T("   get_username(mt_approvedby)    AS clerk,") \
				_T("   Cast_Timestamp2Date(mt_approveddate)  AS dte,") \
				_T("   get_storagename(mt_storage_to_id) AS stock,") \
				_T("   get_departmentname(mt_department_id) AS dept,") \
				_T("   mt_description   AS descr,") \
				_T("   mt_documentno AS documentno") \
				/*_T("   mt_objecttype AS objecttype") \*/
				_T(" FROM   m_transaction") \
				_T(" WHERE mt_transaction_id = %ld AND mt_status <> 'O'"), nOrderID);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONASTERISK | MB_OK);
		return;
	}

	//rs.GetValue(_T("mt_objecttype"), szObjectType);
	rs.GetValue(_T("mt_storage_to_id"), nStorage_id);

	//Kiem tra kho dich vu
	//if (szObjectType == _T("S"))
	//	szStorageCategory = _T("S");
	//else
	{
		szSQL.Format(_T("SELECT msl_category category FROM m_storagelist WHERE msl_storage_id = %d"), nStorage_id);
		rsReport.ExecSQL(szSQL);
		rsReport.GetValue(_T("category"), szStorageCategory);
	}
	szSQL.Format(_T(" SELECT DISTINCT pmdt_name2,") \
		_T("   pmdt_desc,") \
		_T("   pmdt_reportname2, pmdt_index, pmdt_printtype  ") \
		_T(" FROM pms_drugtype,") \
		_T("   m_transactionline,") \
		_T("   m_productitem_view") \
		_T(" WHERE mtl_transaction_id =%ld") \
		_T(" AND mtl_product_item_id  =product_item_id") \
		_T(" AND instr(pmdt_desc, product_producttype) > 0") \
		_T(" ORDER BY pmdt_index "), nOrderID);
	rst.ExecSQL(szSQL);
	while (!rst.IsEOF())
	{
		rst.GetValue(_T("pmdt_name2"), tmpStr);
		StringUpper(tmpStr, szTitle);
		rst.GetValue(_T("pmdt_desc"), szGroup);
		szGroup.Trim();
		
		rst.GetValue(_T("pmdt_reportname2"), tmpStr);
		if(!tmpStr.IsEmpty())
			if(szStorageCategory == _T("S"))
			{
				szReportName.Format(_T("Reports\\HMS\\DV%s"), tmpStr);
			}
			else
				szReportName.Format(_T("Reports\\HMS\\%s"), tmpStr);
		if(!rpt.Init(szReportName) )
			return;   		

		rptHeader = rpt.GetReportHeader();
		rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
		rptHeader->SetValue(_T("REPORTTITLE"), szTitle);
		rptHeader->SetValue(_T("Orderno"), rs.GetValue(_T("orderno")));
		rptHeader->SetValue(_T("DeliverBy"), rs.GetValue(_T("clerk")));
		rptHeader->SetValue(_T("StockName"), rs.GetValue(_T("stock")));
		rptHeader->SetValue(_T("Reason"), rs.GetValue(_T("descr")));
		rs.GetValue(_T("dte"), tmpStr);
		rptHeader->SetValue(_T("ReturnDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rptHeader->SetValue(_T("Department"), rs.GetValue(_T("dept")));
		//So phieu nhap tu dong
		if (pMF->GetModuleID() == _T("MA"))
			rptHeader->SetValue(_T("ReferenceNo"), rs.GetValue(_T("documentno")));

		// Neu la ngoai tru thi lay trong bang hms_pharmaorderline_r
		if(szDepttype == _T("E"))
		{
			if (szStorageCategory == _T("S") && (szGroup.Find(_T("'A1130'")) == -1 && szGroup.Find(_T("'A1140'")) == -1))
			szSQL.Format(_T(" SELECT productid, ") \
						_T("        pname, ") \
						_T("		price,") \
						_T("		uom,") \
						_T("        SUM(solqty)    solqty, ") \
						_T("        SUM(polqty)    polqty, ") \
						_T("        SUM(solinsqty) solinsqty, ") \
						_T("        SUM(othinsqty) othinsqty, ") \
						_T("        SUM(palqty)    palqty, ") \
						_T("        SUM(servqty)    servqty, ") \
						_T("		SUM(solqty+polqty+solinsqty+othinsqty+palqty+servqty) reverseqty,") \
						_T("        SUM(amt)       amt ") \
						_T(" FROM   (SELECT    hpolr_product_id                  productid, ") \
						_T("                   product_name                     AS pname, ") \
						_T("				   get_uomname(product_uomid) uom,") \
						_T("                   CASE WHEN hpo_object_id = 1 THEN hpolr_qtyreverse ") \
						_T("                   ELSE 0 ") \
						_T("                   END                              AS solqty, ") \
						_T("                   CASE WHEN hpo_object_id = 3 THEN hpolr_qtyreverse ") \
						_T("                   ELSE 0 ") \
						_T("                   END                              AS polqty, ") \
						_T("                   CASE WHEN hpo_object_id IN ( 2, 10, 11 ) THEN hpolr_qtyreverse ") \
						_T("                   ELSE 0 ") \
						_T("                   END                              AS solinsqty, ") \
						_T("                   CASE WHEN hpo_object_id IN (4, 5) THEN hpolr_qtyreverse ") \
						_T("                   ELSE 0 ") \
						_T("                   END                              AS othinsqty, ") \
						_T("                   CASE WHEN hpo_object_id = 8 THEN hpolr_qtyreverse ") \
						_T("                   ELSE 0 ") \
						_T("                   END                              AS palqty, ") \
						_T("                   CASE WHEN hpo_object_id = 7 THEN hpolr_qtyreverse ") \
						_T("                   ELSE 0 ") \
						_T("                   END                              AS servqty, ") \
						_T("                   hpolr_unitprice                   AS price, ") \
						_T("                   hpolr_qtyreverse * hpolr_unitprice AS amt ") \
						_T("         FROM      hms_pharmaorder ") \
						_T("         LEFT JOIN hms_pharmaorderline_r ON ( hpo_orderid = hpolr_orderid ) ") \
						_T("         LEFT JOIN m_productitem_view ON ( product_item_id = hpolr_product_item_id ) ") \
						_T("         WHERE     hpolr_retorder_id = %ld ") \
						_T("               AND product_producttype NOT IN ('A1130', 'A1140') ") \
						_T("			   AND hpolr_status = 'A')") \
						_T(" GROUP  BY productid, uom,") \
						_T("           pname, ") \
						_T("           price "), nOrderID);
		else
			szSQL.Format(_T(" SELECT productid, ") \
						_T("        pname, ") \
						_T("		price,") \
						_T("		uom,") \
						_T("        SUM(solqty)    solqty, ") \
						_T("        SUM(polqty)    polqty, ") \
						_T("        SUM(solinsqty) solinsqty, ") \
						_T("        SUM(othinsqty) othinsqty, ") \
						_T("        SUM(palqty)    palqty, ") \
						_T("        SUM(servqty)    servqty, ") \
						_T("		SUM(solqty+polqty+solinsqty+othinsqty+palqty+servqty) reverseqty,") \
						_T("        SUM(amt)       amt ") \
						_T(" FROM   (SELECT    hpolr_product_id                  productid, ") \
						_T("                   product_name                     AS pname, ") \
						_T("				   get_uomname(product_uomid) uom,") \
						_T("                   CASE WHEN hpo_object_id = 1 THEN hpolr_qtyreverse ") \
						_T("                   ELSE 0 ") \
						_T("                   END                              AS solqty, ") \
						_T("                   CASE WHEN hpo_object_id = 3 THEN hpolr_qtyreverse ") \
						_T("                   ELSE 0 ") \
						_T("                   END                              AS polqty, ") \
						_T("                   CASE WHEN hpo_object_id IN ( 2, 10, 11 ) THEN hpolr_qtyreverse ") \
						_T("                   ELSE 0 ") \
						_T("                   END                              AS solinsqty, ") \
						_T("                   CASE WHEN hpo_object_id IN (4, 5, 6, 9, 12) THEN hpolr_qtyreverse ") \
						_T("                   ELSE 0 ") \
						_T("                   END                              AS othinsqty, ") \
						_T("                   CASE WHEN hpo_object_id = 8 THEN hpolr_qtyreverse ") \
						_T("                   ELSE 0 ") \
						_T("                   END                              AS palqty, ") \
						_T("                   CASE WHEN hpo_object_id = 7 THEN hpolr_qtyreverse ") \
						_T("                   ELSE 0 ") \
						_T("                   END                              AS servqty, ") \
						_T("                   hpolr_unitprice                   AS price, ") \
						_T("                   hpolr_qtyreverse * hpolr_unitprice AS amt ") \
						_T("         FROM      hms_pharmaorder ") \
						_T("         LEFT JOIN hms_pharmaorderline_r ON (hpo_orderid = hpolr_orderid ) ") \
						_T("         LEFT JOIN m_productitem_view ON ( product_item_id = hpolr_product_item_id ) ") \
						_T("         WHERE     hpolr_retorder_id = %ld ") \
						_T("               AND product_producttype IN (%s) ") \
						_T("			   AND hpolr_status = 'A')") \
						_T(" GROUP  BY productid, uom,") \
						_T("           pname, ") \
						_T("           price "), nOrderID, szGroup);
		}
		else
		{
			if (szStorageCategory == _T("S") && (szGroup.Find(_T("'A1130'")) == -1 && szGroup.Find(_T("'A1140'")) == -1))
				szSQL.Format(_T(" SELECT productid, ") \
							_T("        pname, ") \
							_T("		price,") \
							_T("		uom,") \
							_T("        SUM(solqty)    solqty, ") \
							_T("        SUM(polqty)    polqty, ") \
							_T("        SUM(solinsqty) solinsqty, ") \
							_T("        SUM(othinsqty) othinsqty, ") \
							_T("        SUM(palqty)    palqty, ") \
							_T("        SUM(servqty)    servqty, ") \
							_T("		SUM(solqty+polqty+solinsqty+othinsqty+palqty+servqty) reverseqty,") \
							_T("        SUM(amt)       amt ") \
							_T(" FROM   (SELECT    hpolr_product_id                  productid, ") \
							_T("                   product_name                     AS pname, ") \
							_T("				   get_uomname(product_uomid) uom,") \
							_T("                   CASE WHEN hpo_object_id = 1 THEN hpolr_qtyreverse ") \
							_T("                   ELSE 0 ") \
							_T("                   END                              AS solqty, ") \
							_T("                   CASE WHEN hpo_object_id = 3 THEN hpolr_qtyreverse ") \
							_T("                   ELSE 0 ") \
							_T("                   END                              AS polqty, ") \
							_T("                   CASE WHEN hpo_object_id IN ( 2, 10, 11 ) THEN hpolr_qtyreverse ") \
							_T("                   ELSE 0 ") \
							_T("                   END                              AS solinsqty, ") \
							_T("                   CASE WHEN hpo_object_id IN (4, 5) THEN hpolr_qtyreverse ") \
							_T("                   ELSE 0 ") \
							_T("                   END                              AS othinsqty, ") \
							_T("                   CASE WHEN hpo_object_id = 8 THEN hpolr_qtyreverse ") \
							_T("                   ELSE 0 ") \
							_T("                   END                              AS palqty, ") \
							_T("                   CASE WHEN hpo_object_id = 7 THEN hpolr_qtyreverse ") \
							_T("                   ELSE 0 ") \
							_T("                   END                              AS servqty, ") \
							_T("                   hpolr_unitprice                   AS price, ") \
							_T("                   hpolr_qtyreverse * hpolr_unitprice AS amt ") \
							_T("         FROM      hms_ipharmaorder ") \
							_T("         LEFT JOIN hms_ipharmaorderline_r ON ( hpo_orderid = hpolr_orderid ) ") \
							_T("         LEFT JOIN m_productitem_view ON ( product_item_id = hpolr_product_item_id ) ") \
							_T("         WHERE     hpolr_retorder_id = %ld ") \
							_T("               AND product_producttype NOT IN ('A1130', 'A1140') ") \
							_T("			   AND hpolr_status = 'A')") \
							_T(" GROUP  BY productid, uom,") \
							_T("           pname, ") \
							_T("           price "), nOrderID);
			else
				szSQL.Format(_T(" SELECT productid, ") \
							_T("        pname, ") \
							_T("		price,") \
							_T("		uom,") \
							_T("        SUM(solqty)    solqty, ") \
							_T("        SUM(polqty)    polqty, ") \
							_T("        SUM(solinsqty) solinsqty, ") \
							_T("        SUM(othinsqty) othinsqty, ") \
							_T("        SUM(palqty)    palqty, ") \
							_T("        SUM(servqty)    servqty, ") \
							_T("		SUM(solqty+polqty+solinsqty+othinsqty+palqty+servqty) reverseqty,") \
							_T("        SUM(amt)       amt ") \
							_T(" FROM   (SELECT    hpolr_product_id                  productid, ") \
							_T("                   product_name                     AS pname, ") \
							_T("				   get_uomname(product_uomid) uom,") \
							_T("                   CASE WHEN hpo_object_id = 1 THEN hpolr_qtyreverse ") \
							_T("                   ELSE 0 ") \
							_T("                   END                              AS solqty, ") \
							_T("                   CASE WHEN hpo_object_id = 3 THEN hpolr_qtyreverse ") \
							_T("                   ELSE 0 ") \
							_T("                   END                              AS polqty, ") \
							_T("                   CASE WHEN hpo_object_id IN ( 2, 10, 11 ) THEN hpolr_qtyreverse ") \
							_T("                   ELSE 0 ") \
							_T("                   END                              AS solinsqty, ") \
							_T("                   CASE WHEN hpo_object_id IN (4, 5, 6, 9, 12) THEN hpolr_qtyreverse ") \
							_T("                   ELSE 0 ") \
							_T("                   END                              AS othinsqty, ") \
							_T("                   CASE WHEN hpo_object_id = 8 THEN hpolr_qtyreverse ") \
							_T("                   ELSE 0 ") \
							_T("                   END                              AS palqty, ") \
							_T("                   CASE WHEN hpo_object_id = 7 THEN hpolr_qtyreverse ") \
							_T("                   ELSE 0 ") \
							_T("                   END                              AS servqty, ") \
							_T("                   hpolr_unitprice                   AS price, ") \
							_T("                   hpolr_qtyreverse * hpolr_unitprice AS amt ") \
							_T("         FROM      hms_ipharmaorder ") \
							_T("         LEFT JOIN hms_ipharmaorderline_r ON (hpo_orderid = hpolr_orderid ) ") \
							_T("         LEFT JOIN m_productitem_view ON ( product_item_id = hpolr_product_item_id ) ") \
							_T("         WHERE     hpolr_retorder_id = %ld ") \
							_T("               AND product_producttype IN (%s) ") \
							_T("			   AND hpolr_status = 'A')") \
							_T(" GROUP  BY productid, uom,") \
							_T("           pname, ") \
							_T("           price "), nOrderID, szGroup);
		}
		rsl.ExecSQL(szSQL);

		//if(rs.IsEOF()){
		//	rst.MoveNext();
		//	continue;

		//}
		if (szStorageCategory == _T("S") && (szGroup.Find(_T("'A1130'")) == -1 && szGroup.Find(_T("'A1140'")) == -1))
		{
			nBreak++;
		}
		nTotalItem = 1;
		nS8 = 0;
		while (!rsl.IsEOF())
		{
			rsl.GetValue(_T("productid"), nProductID);
			rptDetail = rpt.AddDetail();
			rptDetail->SetValue(_T("1"), int2str(nTotalItem++));
			rptDetail->SetValue(_T("2"), rsl.GetValue(_T("pname")));
			rptDetail->SetValue(_T("3"), rsl.GetValue(_T("uom")));
			rptDetail->SetValue(_T("4.1"), rsl.GetValue(_T("solqty")));
			rptDetail->SetValue(_T("4.2"), rsl.GetValue(_T("polqty")));
			rptDetail->SetValue(_T("4.3"), rsl.GetValue(_T("solinsqty")));
			rptDetail->SetValue(_T("4.4"), rsl.GetValue(_T("othinsqty")));
			rptDetail->SetValue(_T("4.5"), rsl.GetValue(_T("Servqty")));
			rptDetail->SetValue(_T("4.6"), rsl.GetValue(_T("palqty")));
			rptDetail->SetValue(_T("4"), rsl.GetValue(_T("reverseqty")));
			rptDetail->SetValue(_T("5"), rsl.GetValue(_T("price")));
			rsl.GetValue(_T("amt"), tmpStr);
			nS8 += str2double(tmpStr);
			rptDetail->SetValue(_T("7"), tmpStr);
			if(szGroup.Find(_T("'A1130'")) != -1 || szGroup.Find(_T("'A1140'")) != -1)
			{
				szSQL.Format(_T(" SELECT    hpo_docno            AS docno, ") \
							_T("           Trim(hp_surname||' '||hp_midname||' '||hp_firstname) AS pname, ") \
							_T("           SUM(hpolr_qtyreverse)   AS orderqty ") \
							_T(" FROM      hms_ipharmaorder ") \
							_T(" LEFT JOIN hms_ipharmaorderline _r ON( hpo_orderid = hpolr_orderid ) ") \
							_T(" LEFT JOIN hms_patient ON( hp_patientno = hpo_patientno ) ") \
							_T(" WHERE     hpolr_retorder_id = %ld ") \
							_T("       AND hpolr_product_id = %ld ") \
							_T(" GROUP     BY hpo_docno, ") \
							_T("              hp_firstname, ") \
							_T("              hp_midname, ") \
							_T("              hp_surname ") \
							_T(" ORDER     BY pname "), nOrderID, nProductID);
				//Neu bo sung tu truc
				//{
				//	szSQL.Format(_T(" SELECT 	hmc_docno as docno,") \
				//	_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				//	_T(" 	hmc_qty as orderqty") \
				//	_T(" FROM hms_medicalcabinet_issue") \
				//	_T(" LEFT JOIN hms_doc ON(hd_docno=hmc_docno)") \
				//	_T(" LEFT JOIN hms_patient ON(hp_patientno=hd_patientno)") \
				//	_T(" WHERE hmc_supplementid='%s' ") \
				//	_T(" and hmc_itemid='%s' ORDER BY docno,pname "), szID, szItemId);
				//}
				rsp.ExecSQL(szSQL);
				while(!rsp.IsEOF()){
					rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
					rptDetail->SetValue(_T("1"), _T(""));
					rsp.GetValue(_T("pname"), tmpStr);
					rptDetail->SetValue(_T("2"), tmpStr);
					rsp.GetValue(_T("docno"), tmpStr);
					rptDetail->SetValue(_T("3"), tmpStr);
					rsp.GetValue(_T("orderqty"), nQty);
					FormatCurrency(nQty, tmpStr);
					rptDetail->SetValue(_T("4"), tmpStr);
					rptDetail->SetValue(_T("5"), _T(""));
					MoneyToString(ToString(nQty), szNote);
					rptDetail->SetValue(_T("6"), szNote);
					rsp.MoveNext();
				}
			}
			rsl.MoveNext();
		}
		rptFooter = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptFooter->SetValue(_T("s6"), double2str(nS6));
		rptFooter->SetValue(_T("s8"), double2str(nS8));
		rptFooter->SetValue(_T("TotalItem"), int2str(nTotalItem - 1));
		tmpStr.Format(_T("%f"), nS8);
		FormatCurrencyStr(tmpStr, tmpStr2);
		rptFooter->SetValue(_T("TotalAmount"), tmpStr2);
		MoneyToString(ToString(nS8), tmpStr);
		if (!tmpStr.IsEmpty())
		{
			tmpStr += _T(" \x111\x1ED3ng");
			rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
		}
		szSysDte = pMF->GetSysDate();
		tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDte.Right(2), szSysDte.Mid(5,2), szSysDte.Left(4));
		rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
		rpt.PrintPreview();
		if (nBreak > 0)
			break;
		rst.MoveNext();
	}
}

void CPrintForms::PrintCompositeExportSheet(long nSheetID){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	CString szSysDte;
	int nIdx = 0;
	double nS7 = 0, nS8 = 0;
	CReportSection *rptDetail, *rptHeader, *rptFooter;
	m_szReportName = _T("Reports\\HMS\\PM_PHIEUXUATKHO.RPT");
	if (!rpt.Init(m_szReportName))
		return;
	szSQL.Format(_T(" SELECT") \
				_T("   mt_orderno       AS orderno,") \
				_T("   get_username(mt_approvedby)    AS clerk,") \
				_T("   get_partnername(mt_partner_id) as deliverto,") \
				_T("   Cast_Timestamp2Date(mt_approveddate)  AS dte,") \
				_T("   get_storagename(mt_storage_id) AS stock,") \
				_T("   get_departmentname(mt_department_id) AS dept,") \
				_T("   mt_description   AS descr,") \
				_T("   mt_documentno documentno") \
				_T(" FROM m_transaction") \
				_T(" WHERE mt_transaction_id = %ld"), nSheetID);	
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONASTERISK | MB_OK);
		return;
	}
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	rptHeader->SetValue(_T("Invoiceno"), rs.GetValue(_T("orderno")));
	rptHeader->SetValue(_T("DeliverBy"), rs.GetValue(_T("clerk")));
	rptHeader->SetValue(_T("StockName"), rs.GetValue(_T("stock")));
	rptHeader->SetValue(_T("Reason"), rs.GetValue(_T("descr")));
	rs.GetValue(_T("dte"), szSysDte);
	tmpStr.Format(rptHeader->GetValue(_T("OrderDate")), szSysDte.Mid(8, 2), szSysDte.Mid(5, 2), szSysDte.Left(4));
	rptHeader->SetValue(_T("OrderDate"), tmpStr);
	rptHeader->SetValue(_T("Department"), rs.GetValue(_T("dept")));
	rptHeader->SetValue(_T("DeliveryTo"), rs.GetValue(_T("deliverto")));
	rptHeader->SetValue(_T("ReferenceNo"), rs.GetValue(_T("documentno")));
	szSQL.Format(_T(" SELECT") \
				_T("   product_name   AS pname,") \
				_T("   mtl_qtydelivery	AS qty,") \
				_T("   product_uomname	AS uom,") \
				_T("   product_expdate	AS expdte,") \
				_T("   product_taxprice AS price,") \
				_T("   product_lotno	AS lot,") \
				_T("   mtl_qtydelivery*product_taxprice AS amt") \
				_T(" FROM m_transactionline") \
				_T(" LEFT JOIN m_productitem_view ON (product_item_id=mtl_product_item_id)") \
				_T(" WHERE mtl_transaction_id = %ld"), nSheetID);	
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;
	while (!rs.IsEOF())
	{
		nIdx++;
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("pname")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("uom")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("lot")));
		rs.GetValue(_T("expdte"), tmpStr);
		rptDetail->SetValue(_T("5"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("price")));
		rs.GetValue(_T("qty"), tmpStr);
		nS7 += str2double(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rs.GetValue(_T("amt"), tmpStr);
		nS8 += str2double(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.MoveNext();
	}
	rptFooter = rpt.AddDetail(rpt.GetGroupFooter(1));
	rptFooter->SetValue(_T("s7"), double2str(nS7));
	rptFooter->SetValue(_T("s8"), double2str(nS8));
	rpt.GetReportFooter()->SetValue(_T("TotalItem"), int2str(nIdx));
	MoneyToString(ToString(nS8), tmpStr);
	if (!tmpStr.IsEmpty())
	{
		tmpStr += _T(" \x111\x1ED3ng");
		rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	}
	szSysDte = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDte.Right(2), szSysDte.Mid(5,2), szSysDte.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();	
}

void CPrintForms::PrintDepartmentExportSheet(long nSheetID, CString szFilter){
	CGuiMainFrame *pMF = (CGuiMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szOrderBy;
	CRecord rs(&pMF->m_db);
	szSQL.Format(
		_T(" SELECT  mt_transaction_id as invoiceid,") \
		_T("	mt_orderno as invoiceno,") \
		_T("  	mt_posteddate as issuedate ,") \
		_T("  	mt_orderdate as invoicedate,") \
		_T(" 	GET_STORAGENAME(mt_storage_id) as storagename,	") \
		_T("  	GET_STORAGENAME(mt_storage_to_id) as storagename_to,") \
		_T("  	GET_DEPARTMENTNAME(mt_department_to_id) as department,") \
		_T("  	GET_SYSSEL_DESC('pms_export_type', mt_exptype) as exptype,") \
		_T("  	GET_USERNAME(mt_deliveryby) as deliverer,") \
		_T("  	GET_USERNAME(mt_receiveby) as receiver, ") \
		_T("	mt_description as note, ") \
		_T("	mt_status as status, ") \
		_T("	mt_doctype as doctype, mt_documentno ") \
		_T(" FROM m_transaction") \
		_T(" WHERE mt_transaction_id=%ld  "), nSheetID);

	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	
	CReport rpt;
	CString tmpStr, szAmount;
	CString szDate, szSysDate;
	CString szDocType;

	if(!rpt.Init(_T("Reports/HMS/PMS_DRUGEXPORT.RPT")) )
		return;

	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	rs.GetValue(_T("InvoiceNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);

	rs.GetValue(_T("invoicedate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("OrderDate"),CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), szDate);
	rs.GetValue(_T("doctype"), szDocType);
	tmpStr.Empty();

	rs.GetValue(_T("storagename"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("StockName"), tmpStr);
	rs.GetValue(_T("storagename_to"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("StockName1"), tmpStr);

	rs.GetValue(_T("Department"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);

	rs.GetValue(_T("note"), tmpStr);	
	rpt.GetReportHeader()->SetValue(_T("Comment"), tmpStr);
	rs.GetValue(_T("deliverer"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("deliverer"), tmpStr);
	rs.GetValue(_T("receiver"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("receiver"), tmpStr);

	tmpStr.Empty();
	//pMF->GetStatusString(rs.GetValue(_T("status")), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Status"), tmpStr);


	if (pMF->GetModuleID() == _T("MA"))
	{
		rs.GetValue(_T("mt_documentno"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("ReferenceNo"), tmpStr);
	}

	//Page Header
	//Report Detail

	int nCount = 0;
	int nQty = 0;
	double dMoney = 0;
	long double nTotalAmount = 0;

	CString szMoney, szExprDate;
	szSQL.Format(_T(" SELECT product_id AS id,") \
				_T("        product_name AS name,") \
				_T("        product_uomname AS unit,") \
				_T("        product_expdate AS expdate,") \
				_T("        product_lotno AS lotno,") \
				_T("        SUM(mtl_qtyorder) AS qtyorder,") \
				_T("        SUM(mtl_qtydelivery) AS qtyissue,") \
				_T("        mtl_saleprice AS price,") \
				_T("        SUM(mtl_qtyorder*mtl_saleprice) AS orderamount,") \
				_T("        SUM(mtl_qtydelivery*mtl_saleprice) AS issueamount") \
				_T(" FROM m_transactionline") \
				_T(" LEFT JOIN m_productitem_view ON(mtl_product_item_id=product_item_id)") \
				_T(" WHERE mtl_transaction_id=%ld") \
				_T(" GROUP BY product_id, product_name, product_uomname,") \
				_T("          product_expdate, product_lotno,") \
				_T("          mtl_saleprice, mtl_line") \
				_T(" ORDER BY mtl_line"), nSheetID);

	//_msg(_T("%s"), szSQL);

	rs.ExecSQL(szSQL);
	CReportSection* rptDetail = rpt.GetDetail(0); 
	while (!rs.IsEOF())
	{ 
		rptDetail = rpt.AddDetail();

		nCount++;

		tmpStr.Format(_T("%d"), nCount);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("name"), tmpStr) ;
		rptDetail->SetValue(_T("2"), tmpStr);
		
		rs.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("expdate"), tmpStr);
		szExprDate = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("4"), szExprDate);

		rs.GetValue(_T("lotno"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("qtyorder"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("qtyissue"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("price"), dMoney);
		FormatCurrency(dMoney, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("issueamount"), dMoney);
		nTotalAmount += dMoney;
		FormatCurrency(dMoney, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);


		rs.MoveNext();
	}
	//Page Footer
	
	//Report Footer
	tmpStr.Format(_T("%d"), nCount);
	rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);	
	FormatCurrency(nTotalAmount, szAmount);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), szAmount);
	szAmount.Format(_T("%.2f"), nTotalAmount);
	MoneyToString(szAmount, tmpStr);	
#ifdef UNICODE
	tmpStr += _T(" \x111\x1ED3ng.");
#endif
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);

	tmpStr = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(11,5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	//GetDoctorName(pMF->GetCurrentUser());

	rpt.PrintPreview();
}

CString CPrintForms::GetQueryString(LPCTSTR szQueryID, LPCTSTR szInvoiceID){
	CString szSQL, szWhere;
	if(szQueryID == _T("INVOICE"))
	{
		szWhere.Format(_T(" AND fac_invoiceno = '%s'"), szInvoiceID);
		szSQL.Format(_T(" SELECT") \
					_T("   fac_cash_id      AS id,") \
					_T("   fac_client_id    AS client,") \
					_T("   fac_org_id       AS org,") \
					_T("   fac_user_id      AS userid,") \
					_T("   fac_invoicedate  AS invoicedate,") \
					_T("   facl_debit_acct  AS debit,") \
					_T("   facl_credit_acct AS credit,") \
					_T("   fac_invoiceno    AS invoiceno,") \
					_T("   fac_acctdate     AS acctdate,") \
					_T("   fac_partner_id   AS partnerid,") \
					_T("   fac_transactor   AS transactor,") \
					_T("   fac_reason       AS reason,") \
					_T("   fac_amount       AS totalamount") \
					_T(" FROM   fa_cash") \
					_T(" LEFT JOIN fa_cashline ON (fac_cash_id=facl_cash_id)") \
					_T(" WHERE  fac_isactive='Y' %s") \
					_T(" ORDER  BY fac_cash_id "), szWhere);
	}
	//if(szQueryID == _T("Bank"))
	//{
	//	szWhere.Format(_T(" AND fac_invoiceno = '%s'"), szInvoiceID);
	//	szSQL.Format(_T("SELECT fac_cash_id as id, fac_client_id as client, fac_org_id as org, fac_user_id as uid, fac_invoicedate as invoicedate, ") \
	//			_T("facl_debit_acct as debit, facl_credit_acct as credit,") \
	//			_T("fac_invoiceno as invoiceno, fac_acctdate as acctdate, ") \
	//			_T("fac_partner_id as partnerid, fac_transactor as transactor, ") \
	//			_T("fac_reason as reason, fac_amount as totalamount ") \
	//			_T(" FROM fa_cash LEFT JOIN fa_cashline ON (fac_cash_id = facl_cash_id) WHERE fac_isactive='Y' %s ORDER BY fac_cash_id"), szWhere);
	//}
	if (szQueryID == _T("CASH_VOUCHER"))
	{
		szWhere.Format(_T(" AND fac_invoiceno = '%s'"), szInvoiceID);
		szSQL.Format(_T(" SELECT") \
					_T("   'COMPANY NAME'                AS org,") \
					_T("   'DEPARTMENT'                  AS department,") \
					_T("   (SELECT") \
					_T("      adp_name") \
					_T("    FROM   ad_partner") \
					_T("    WHERE  adp_partner_id=fac_partner_id) AS obj,") \
					_T("   (SELECT") \
					_T("      adp_address") \
					_T("    FROM   ad_partner") \
					_T("    WHERE  adp_partner_id=fac_partner_id) AS addr,") \
					_T("   fac_invoicedate               AS invoicedate,") \
					_T("   facl_debit_acct               AS debit,") \
					_T("   facl_credit_acct              AS credit,") \
					_T("   fac_invoiceno                 AS invoiceno,") \
					_T("   facl_description              AS descr,") \
					_T("   facl_baseamt                  AS base,") \
					_T("   facl_taxamt                   AS taxamt,") \
					_T("   fac_amount                    AS amt,") \
					_T("   facl_tax_acct                 AS tax") \
					_T(" FROM   fa_cash") \
					_T("        LEFT JOIN fa_cashline ON (fac_cash_id=facl_cash_id)") \
					_T(" WHERE  fac_isactive='Y' %s") \
					_T(" ORDER  BY fac_cash_id "), szWhere);
	}
	if (szQueryID == _T("BANK_VOUCHER"))
	{
		szWhere.Format(_T(" AND fabs_invoiceno = '%s'"), szInvoiceID);
		szSQL.Format(_T(" SELECT") \
					_T("   'COMPANY NAME'                          AS org,") \
					_T("   'DEPARTMENT'                            AS department,") \
					_T("   (SELECT") \
					_T("      fap_name") \
					_T("    FROM   fa_partner") \
					_T("    WHERE  fap_partner_id=fabs_partner_id) AS obj,") \
					_T("   (SELECT") \
					_T("      fap_name") \
					_T("    FROM   fa_partner") \
					_T("    WHERE  fap_partner_id=fabs_partner_id) AS addr,") \
					_T("   fabs_invoicedate                        AS invoicedate,") \
					_T("   fabsl_debit_acct                        AS debit,") \
					_T("   fabsl_credit_acct                       AS credit,") \
					_T("   fabs_invoiceno                          AS invoiceno,") \
					_T("   fabsl_description                       AS descr,") \
					_T("   fabsl_baseamt                           AS base,") \
					_T("   fabsl_taxamt                            AS taxamt,") \
					_T("   fabs_amount                             AS amt,") \
					_T("   fabsl_tax_acct                          AS tax") \
					_T(" FROM   fa_bankstatement") \
					_T("        LEFT JOIN fa_bankstatementline ON (fabs_bankstatement_id=fabsl_bankstatement_id)") \
					_T(" WHERE  fabs_isactive='Y' %s") \
					_T(" ORDER  BY fabs_bankstatement_id "), szWhere);
	}
	if (szQueryID == _T("PAYMENT_ORDER"))
	{
		szSQL.Format(_T("SELECT 'SerialNo' as SerialNo, 'PrintDate' as PrintDate,") \
					_T(" 'Payer' as Payer, 'AccNo' as AccNo, 'Bank' as Bank, 'Receiver' as Receiver, 'AccNo0' as AccNo0,") \
					_T(" 'Bank0' as Bank0, 'TotalAmount' as TotalAmount, 'SumInWord' as SumInWord, 'Reason' as Reason") \
					_T(" FROM dual"));
	}
	if (szQueryID == _T("CREDIT_NOTE"))
	{
		szSQL.Format(_T(" SELECT 'Funder' as Funder, 'Address' as Address,") \
					_T(" 'Reason' as Reason, 'Currency' as Currency, 'Amount' as Amount,") \
					_T(" 'AmountInWord' as AmountInWord, 'No' as No, 'Date' as dte, 'AccNo' as AccNo, 1 as c1,") \
					_T(" 2 as c2, 3 as c3, 4 as c4, 5 as c5") \
					_T(" FROM dual"));
	}
	if (szQueryID == _T("PURCHASE_ORDER"))
		//szSQL.Format(_T("SELECT 'DebitAccount' as DebAcc, 'CreditAccount' as CreAcc, 'SheetNo' as ShtNo, 'SupplierName' as SupNme, 'StockName' as StkNme, 'Address' as Ads, 1 as c1, 2 as c2, 3 as c3 , 4 as c4, 5 as c5, 6 as c6, 7 as c7, 8 as c8 FROM dual"));
		szSQL.Format(_T(" SELECT") \
					_T("   po_debitacct                          AS debacc,") \
					_T("   po_creditacct                         AS creacc,") \
					_T("   po_orderno                            AS shtno,") \
					_T("   Get_partnername(po_partner_id)        AS supnme,") \
					_T("   Get_storagename(po_storage_id)        AS stknme,") \
					_T("   'Address'                             AS ads,") \
					_T("   1                                     AS c1,") \
					_T("   (SELECT") \
					_T("      mp_name") \
					_T("    FROM   m_product") \
					_T("    WHERE  mp_product_id=pol_product_id) AS c2,") \
					_T("   pol_product_id                        AS c3,") \
					_T("   Get_uomname(pol_uom_id)               AS c4,") \
					_T("   pol_qtyorder                          AS c5,") \
					_T("   pol_qtyinvoice                        AS c6,") \
					_T("   pol_unitprice                         AS c7,") \
					_T("   pol_netamount                         AS c8") \
					_T(" FROM   purchase_order") \
					_T("        LEFT JOIN purchase_orderline ON (po_purchase_order_id=pol_purchase_order_id)") \
					_T(" WHERE  po_status='A' AND po_orderno = '%s'"), szInvoiceID);
	if (szQueryID == _T("EXPORT_ORDER"))
		szSQL.Format(_T("SELECT 'DebitAccount' as DebAcc, 'CreditAccount' as CreAcc, 'SheetNo' as ShtNo, 'SupplierName' as SupNme, 'StockName' as StkNme, 'Address' as Ads, 'Unit' as Unt, 1 as c1, 2 as c2, 3 as c3 , 4 as c4, 5 as c5, 6 as c6, 7 as c7, 8 as c8 FROM dual"));
	if (szQueryID == _T("SHEET_LIST"))
	{
		if(!m_szType.IsEmpty())
			szWhere.Format(_T(" and fac_invoicetype='%s' "), m_szType);
		if (!m_szFromDate.IsEmpty() && !m_szToDate.IsEmpty())
			szWhere.AppendFormat(_T(" and Cast_Timestamp2Date(fac_acctdate) between cast_string2date('%s') and cast_string2date('%s') "), m_szFromDate, m_szToDate);
		if(!m_szID.IsEmpty())
		{
			szWhere.AppendFormat(_T(" and fac_invoiceno='%s' "), m_szID);
		}
		szSQL.Format(_T(" SELECT") \
					_T("   fac_cash_id                            AS id,") \
					_T("   fac_client_id,") \
					_T("   fac_org_id,") \
					_T("   fac_user_id,") \
					_T("   fac_invoicedate                        AS invdate,") \
					_T("   fac_invoiceno                          AS invno,") \
					_T("   fac_acctdate                           AS accdate,") \
					_T("   (SELECT") \
					_T("      DISTINCT") \
					_T("      fap_name") \
					_T("    FROM   fa_partner") \
					_T("    WHERE  fap_partner_id=fac_partner_id) AS partner,") \
					_T("   fac_transactor                         AS transactor,") \
					_T("   fac_reason                             AS reason,") \
					_T("   fac_amount                             AS totalamt") \
					_T(" FROM   fa_cash") \
					_T(" WHERE  fac_isactive='Y' %s") \
					_T(" ORDER  BY fac_cash_id "), szWhere);
	}
	return szSQL;
}

void CPrintForms::SetReference(CString szSheetType, CString szSheetID, CString szFromDate, CString szToDate){
	m_szType = szSheetType;
	if (!szFromDate.IsEmpty())
		m_szFromDate = szFromDate;
	else
		m_szFromDate.Empty();
	if (!szToDate.IsEmpty())
		m_szToDate = szToDate;
	else
		m_szToDate.Empty();
	if (!szSheetID.IsEmpty())
		m_szID = szSheetID;
	else
		m_szID.Empty();
}

void CPrintForms::PMS_PrintDrugDelivery(long nOrderID, bool bPreview, bool bDirect){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr;

	//Report Header

	CString szSQL, szUsage, szDeptType;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	long nDocumentNo;
	int nRefIndex;			
	double dblTotalAmount=0;

	szSQL.Format(_T(" SELECT 	hd_patientno as patientno,  ") \
		_T(" 	hd_docno as docno,") \
		_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
		_T(" 	TO_CHAR(hp_birthdate, 'YYYY') as birthyear,") \
		_T(" 	hms_getage(cast_timestamp2date(hd_admitdate), hp_birthdate) as age,") \
		_T(" 	hp_sex as sexid,") \
		_T(" 	sys_sel_getname('sys_sex', hp_sex) as sex,") \
		_T(" 	sys_sel_getname('hms_rank', hp_rank) as rank,") \
		_T(" 	hp_occupation as occupationid,") \
		_T(" 	hp_dtladdr as detailaddress,") \
		_T(" 	hms_getaddress(hp_provid,hp_distid, hp_villid) as address,") \
		_T(" 	hp_workplaceid as workplaceid,") \
		_T(" 	hp_workplace as workplace,") \
		_T(" 	hd_object as objectid,") \
		_T(" 	ho_desc as objectname,") \
		_T(" 	ho_type as objecttype,") \
		_T(" 	hd_cardno as cardno,") \
		_T(" 	hd_cardidx as cardidx, ") \
		_T("	hd_reldisease as relationdisease, ") \
		_T("	hd_diagnostic ||' ['||hd_icd||']' as diagnostic, ") \
		_T(" 	hd_status, ") \
		_T(" 	hdf_acceptedfee, ") \
		_T("	Cast_Timestamp2Date(hd_admitdate) as examdate, ") \
		_T(" 	hpo_orderdate as orderdate, ") \
		_T(" 	hpo_doctor as doctor, ") \
		_T(" 	hpo_storage_id as stockid, ") \
		_T(" 	hpo_deptid as deptid, ") \
		_T(" 	hpo_roomid as roomid, ") \
		_T(" 	HMS_GETROOMNAME(hpo_deptid, hpo_roomid)  as roomname, ") \
		_T(" 	hpo_bedid as bedid, ") \
		_T(" 	hpo_depttype as depttype, ") \
		_T(" 	hpo_refidx as refidx, ") \
		_T(" 	hpo_printed as printed ") \
		_T(" FROM hms_ipharmaorder") \
		_T(" LEFT JOIN hms_patient ON(hp_patientno=hpo_patientno)") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno and hd_docno=hpo_docno)") \
		_T(" LEFT JOIN hms_object ON(ho_id=hd_object) ") \
		_T(" LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx) ") \
		_T(" WHERE hpo_orderid=%ld AND hpo_orderstatus NOT IN('O','C') "), nOrderID);
	rs.ExecSQL(szSQL);

	if(rs.IsEOF())		
	{
		_debug(_T("Order not found."));
		return;					
	}
	

	szSQL.Format(_T(" SELECT hpol_line,") \
	_T("   hpol_product_id,") \
	_T("   hpol_productcode,") \
	_T("   hpol_productname,") \
	_T("   hpol_productuom,") \
	_T("   hpol_unitprice,") \
	_T("   SUM(hpol_qtyissue) AS hpol_qtyissue,") \
	_T("   SUM(hpol_qtyissue*hpol_unitprice) AS hpol_amount,") \
	_T("   hpol_productcountry, ") \
	_T(" HMS_GETDOUSAGE(hpol_orderid, hpol_product_id) as hpol_usage ") \
	_T(" FROM hms_ipharmaorderline_view") \
	_T(" WHERE hpol_orderid =%ld ") \
	_T(" GROUP BY hpol_orderid, hpol_line,") \
	_T("   hpol_product_id,") \
	_T("   hpol_productcode,") \
	_T("   hpol_productname,") \
	_T("   hpol_productuom,") \
	_T("   hpol_unitprice,") \
	_T("   hpol_productcountry ") \
	_T(" HAVING SUM(hpol_qtyissue) > 0 ")\
	_T(" ORDER BY hpol_productname,hpol_line"), nOrderID);
	
	rsl.ExecSQL(szSQL);
	if(rsl.IsEOF()){	

		_debug(_T("\r\nData not found"));
		return;
	}
	
	CString szReportName;
	if(rs.GetValue(_T("objectid")) == _T("7"))
		szReportName.Format(_T("Reports/HMS/PMS_DRUGDELIVERYVOTE_DV.RPT"));
	else
		if(rs.GetValue(_T("objectid")) == _T("1") || rs.GetValue(_T("objectid")) == _T("3") ||
			rs.GetValue(_T("objectid")) == _T("8"))
			szReportName.Format(_T("Reports/HMS/PMS_DRUGDELIVERYVOTE_D.RPT"));
		else
			szReportName.Format(_T("Reports/HMS/PMS_DRUGDELIVERYVOTE.RPT"));


	if(!rpt.Init(szReportName, false, false, 1))
		return;
_tprintf(_T("\r\n%s"), szReportName);

	//rpt.SetPrintCallback(_OnCheckPrintPrescriptionFnc);

	tmpStr.Empty();
	/*rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);*/

	StringUpper(pMF->m_szHealthService, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HealthService"), tmpStr);
	StringUpper(pMF->m_szHospitalName, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), tmpStr);
	rs.GetValue(_T("deptid"), tmpStr);
//	rpt.GetReportHeader()->SetValue(_T("Department"), GetDepartmentName(tmpStr));
	rs.GetValue(_T("roomname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RoomName"), tmpStr);

	rs.GetValue(_T("docno"), nDocumentNo);
	tmpStr.Format(_T("%ld"), nDocumentNo);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);

	rpt.GetReportHeader()->SetValue(_T("Barcode[128A]"),tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128B]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128C]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[39]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[93]"), tmpStr);

	tmpStr.Format(_T("%ld"), nOrderID);
	rpt.GetReportHeader()->SetValue(_T("OrderID"), tmpStr);

	CString szPatientName;
	rs.GetValue(_T("pname"), szPatientName);
	rpt.GetReportHeader()->SetValue(_T("PATIENTNAME"), szPatientName);
//In ra tuoi	
	rs.GetValue(_T("age"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), tmpStr);
//Chi In ra nam sinh.
	rs.GetValue(_T("birthyear"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);

	rs.GetValue(_T("sex"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Sex"), tmpStr);
	rs.GetValue(_T("address"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Address"), tmpStr);
	rs.GetValue(_T("workplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Workplace"), tmpStr);
	rs.GetValue(_T("objectname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Object"), tmpStr);

	rs.GetValue(_T("examdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ExamDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

	rs.GetValue(_T("rank"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Rank"), tmpStr);
	
	rs.GetValue(_T("cardno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CardNo"), tmpStr);
	int nStockId;
	rs.GetValue(_T("stockid"), nStockId);
//	rpt.GetReportHeader()->SetValue(_T("Stock"), GetStorageName(nStockId));
	rs.GetValue(_T("refidx"), nRefIndex);
	//rs.GetValue(_T("roomid"), pMF->m_nRoomID);
	//tmpStr.Format(_T("%d"), m_nRoomID);
	//rpt.GetReportHeader()->SetValue(_T("RoomID"), tmpStr);
	//rs.GetValue(_T("bedid"), tmpStr);
	//rs.GetValue(_T("depttype"), szDeptType);
	//szDeptType.Trim();
	
	//szSQL.Format(_T("SELECT he_icd10 as Icd,he_diagnostic || ' [' || he_icd10 || ']' as Diagnostic FROM hms_exam WHERE he_docno=%ld AND he_roomid=%d"), nDocumentNo, m_nRoomID);
	//rs2.ExecSQL(szSQL);
	//if(!rs2.IsEOF())
	//{
	//	rpt.GetReportHeader()->SetValue(_T("Diagnostic"), rs2.GetValue(_T("Diagnostic")));			

	//}
	//else
	//{
	//	rs.GetValue(_T("Diagnostic"), tmpStr);
	//	rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);
	//}
	//
	rs.GetValue(_T("relationdisease"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("relationdisease"), tmpStr);
	
	
	//Page Header
	//Report Detail
	CReportSection* rptDetail = rpt.GetDetail(0); 
	int nItem=1;

	CString szWhere, tmpStr2, szQty;

	nItem = 1;
	dblTotalAmount = 0;
	double money=0;
	while(!rsl.IsEOF()){ 
		rptDetail = rpt.AddDetail();
		rsl.GetValue(_T("hpol_usage"), szUsage);
		
		rsl.GetValue(_T("hpol_amount"), money);
		dblTotalAmount += money;
		tmpStr.Format(_T("%d"), nItem++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rsl.GetValue(_T("hpol_productname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rsl.GetValue(_T("hpol_productuom"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rsl.GetValue(_T("hpol_qtyissue"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rptDetail->SetValue(_T("4.1"), _T(""));

		rsl.GetValue(_T("hpol_unitprice"), tmpStr);
		FormatCurrencyStr(tmpStr, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rsl.GetValue(_T("hpol_amount"), tmpStr);
		FormatCurrencyStr(tmpStr, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rsl.GetValue(_T("hpol_productcountry"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rsl.MoveNext();
	}

	//Page Footer
	//Report Footer
	CString tmpVal;
	tmpStr.Format(_T("%d"), nItem-1);
	rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
	FormatCurrency(dblTotalAmount, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	tmpVal.Format(_T("%2.f"), dblTotalAmount);
	MoneyToString(tmpVal, tmpStr);
#ifdef UNICODE
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr+ _T(" \x111\x1ED3ng"));
#endif
	CString szDate;
	//rs.GetValue(_T("orderdate"), tmpStr);

	tmpStr = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(11,5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	
	
	rpt.GetReportFooter()->SetValue(_T("PatientSign"), szPatientName);

	rs.GetValue(_T("doctor"), tmpStr);

	/*rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(tmpStr));
	CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
	if(img)
	{
		img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img->SetFixedHeight(false);
	}*/

	//rpt.GetReportFooter()->SetValue(_T("IssueBy"), pMF->GetDoctorName(GetCurrentUser()));
	CString szStatus;
	CString szAcceptedFee;
	CString szObjectType;
	int nPrinted=0;

	rs.GetValue(_T("hd_status"), szStatus);
	rs.GetValue(_T("objecttype"), szObjectType);
	rs.GetValue(_T("printed"), nPrinted);
	rs.GetValue(_T("hd_acceptedfee"), szAcceptedFee);

	//if(CheckPermission(_T("10.10")))
	//{
	//	if(bPreview)
	//		rpt.PrintPreview();
	//	else
	//	{
	//		rpt.Print(bDirect);		
	//	}
	//}
	//else
	//{
	//	if(bPreview)
	//	{
	//		if((szObjectType == _T("I") || szObjectType == _T("C")) && (szAcceptedFee != _T("Y")) )
	//		{
	//			nPrinted = 1;
	//		}

	//		if(szObjectType == _T("D") && szStatus != _T("T"))
	//			nPrinted = 1;

	//		if(nPrinted > 0 && !CheckPermission(_T("10.10")))
	//			rpt.PrintPreview(false);
	//		else
	//			rpt.PrintPreview();

	//	}
	//	else
	//	{
	//		
	//		if(nPrinted > 0){
	//			
	//			CString szMsg;
	//			szMsg.Format(_T("\x110\x1A1n thu\x1ED1\x63 \x111\xE3 \x111\x1B0\x1EE3\x63 in [%\x64] l\x1EA7n. Kh\xF4ng th\x1EC3 in l\x1EA1i!"), nPrinted);
	//			ShowMessageBox(szMsg, MB_OK);

	//			return;
	//		}
	//		if(szObjectType == _T("D") && szStatus != _T("T"))
	//		{
	//			ShowMessageBox(_T("Patient profile not finish. Cannot allow print prescription"));
	//			return;
	//		}

	//		if((szObjectType == _T("I") || szObjectType == _T("C")) && (szAcceptedFee != _T("Y"))  && GetModuleID() != _T("FM") )
	//		{
	//			ShowMessageBox(_T("\x42\x1EC7nh nh\xE2n \x42HYT \x63h\x1B0\x61 \x64uy\x1EC7t ph\xED. Kh\xF4ng \x111\x1B0\x1EE3\x63 in \x111\x1A1n thu\x1ED1\x63!"));
	//			return;
	//		}
	//		

	//		rpt.Print(bDirect);		
	//	}
	//}

	rpt.PrintPreview(true);
}

void CPrintForms::PrintDrugReturn(long nOrderID){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr;

	//Report Header
	
	CString szSQL, szUsage, szDeptType;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	long nDocumentNo;
	double dblTotalAmount=0, nTemp = 0;

	szSQL.Format(_T(" SELECT") \
		_T("   hd_patientno                                                AS patientno,") \
		_T("   hd_docno                                                    AS docno,") \
		_T("   get_patientname(hpo_docno) AS pname,") \
		_T("   To_char(hp_birthdate, 'YYYY')                               AS birthyear,") \
		_T("   Hms_getage(Cast_timestamp2date(hd_admitdate), hp_birthdate) AS age,") \
		_T("   Get_syssel_desc('sys_sex', hp_sex)                          AS sex,") \
		_T("   hp_dtladdr                                                  AS detailaddress,") \
		_T("   CASE WHEN hpo_doctype = 'DRO' THEN CAST(Hms_getaddress(hp_provid, hp_distid, hp_villid) AS NVARCHAR2(254))") \
		_T("   ELSE (SELECT so_partneraddress FROM sale_order WHERE so_sale_order_id = hpo_orderid) END AS address,") \
		_T("   hp_workplaceid                                              AS workplaceid,") \
		_T("   hp_workplace                                                AS workplace,") \
		_T("   hd_object                                                   AS objectid,") \
		_T("   Hms_getobjectname(hd_object)                                AS objectname,") \
		_T("   hd_cardno                                                   AS cardno,") \
		_T("   hd_cardidx                                                  AS cardidx,") \
		_T("   hd_reldisease                                               AS relationdisease,") \
		_T("   hd_diagnostic||' ['||hd_icd||']'                            AS diagnostic,") \
		_T("   hc_regdate                                                  AS regdate,") \
		_T("   hc_expdate                                                  AS expdate,") \
		_T("   hpo_approvedby                                              AS doctor,") \
		_T("   get_storagename(hpo_storage_id)							   AS stock,") \
		_T("   CASE WHEN hpo_doctype = 'DRO' THEN CAST(hpo_orderid AS NVARCHAR2(8)) ") \
		_T("   ELSE (select so_orderno from sale_order where so_sale_order_id = hpo_orderid) END AS forderid,") \
		_T("   hpo_approveddate AS retdate,") \
		_T("   get_departmentname(hpo_deptid)                                                  AS dept") \
		_T(" FROM   hms_pharmareturnorder") \
		_T("        LEFT JOIN hms_doc ON(hd_docno=hpo_docno)") \
		_T("        LEFT JOIN hms_patient ON(hd_patientno=hp_patientno)") \
		_T("        LEFT JOIN hms_card ON(hc_patientno=hd_patientno") \
		_T("                              AND hc_idx=hd_cardidx) ") \
		_T(" WHERE hpo_pharmareturnorder_id=%ld AND hpo_status <>'O' "), nOrderID);
	rs.ExecSQL(szSQL);
	if(rs.IsEOF())														   
		return;					

	szSQL.Format(_T(" SELECT ") \
				_T("   product_item_id,") \
				_T("   product_name,") \
				_T("   product_uomname,") \
				_T("   SUM(hpol_qtyreturn)                AS product_qtyreverse,") \
				_T("   hpol_unitprice AS price,") \
				_T("   SUM(hpol_qtyreturn*hpol_unitprice) AS amount,") \
				_T("   product_countryname AS mfg") \
				_T(" FROM hms_pharmareturnorder_line") \
				_T(" LEFT JOIN m_productitem_view") \
				_T(" ON(product_item_id             =hpol_product_item_id)") \
				_T(" WHERE hpol_pharmareturnorder_id=%ld ") \
				_T(" GROUP BY ") \
				_T("   product_item_id,") \
				_T("   product_name,") \
				_T("   product_uomname,") \
				_T("   hpol_unitprice,") \
				_T("   product_countryname") \
				_T(" ORDER BY product_name "), nOrderID);
	rsl.ExecSQL(szSQL);
	if(rsl.IsEOF()){
		return;
	}

	if(!rpt.Init(_T("Reports/HMS/PMS_DRUGRETURN.RPT")))
		return;
	

	tmpStr.Empty();
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	rs.GetValue(_T("dept"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);
	rs.GetValue(_T("roomname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RoomName"), tmpStr);

	rs.GetValue(_T("docno"), nDocumentNo);
	tmpStr.Format(_T("%ld"), nDocumentNo);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	tmpStr.Format(_T("%ld"), nOrderID);
	rpt.GetReportHeader()->SetValue(_T("OrderID"), tmpStr);
	rs.GetValue(_T("forderid"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("FOrderID"), tmpStr);
	rs.GetValue(_T("stock"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Stock"), tmpStr);
	CString szPatientName;
	rs.GetValue(_T("pname"), szPatientName);
	rpt.GetReportHeader()->SetValue(_T("PATIENTNAME"), szPatientName);
//In ra tuoi	
	rs.GetValue(_T("age"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), tmpStr);
//Chi In ra nam sinh.
	rs.GetValue(_T("birthyear"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);

	rs.GetValue(_T("sex"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Sex"), tmpStr);
	rs.GetValue(_T("address"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Address"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReturnDate")), CDate::Convert(rs.GetValue(_T("retdate")), yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReturnDate"), tmpStr);
	
	//Page Header
	//Report Detail
	CReportSection* rptDetail = rpt.GetDetail(0); 
	int nItem=1;

	CString szWhere, tmpStr2, szQty;

	nItem = 1;
	dblTotalAmount = 0;
	while(!rsl.IsEOF()){ 
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nItem++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rsl.GetValue(_T("product_name"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rsl.GetValue(_T("product_uomname"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rsl.GetValue(_T("product_qtyreverse"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rsl.GetValue(_T("price"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rsl.GetValue(_T("amount"), nTemp);
		dblTotalAmount += nTemp;
		tmpStr.Format(_T("%f"), nTemp);
		rptDetail->SetValue(_T("6"), tmpStr);
		rsl.GetValue(_T("mfg"), tmpStr);	
		rptDetail->SetValue(_T("7"), tmpStr);
		rsl.MoveNext();
	}

	//Page Footer
	//Report Footer
	CString tmpVal;
	tmpStr.Format(_T("%d"), nItem-1);
	rpt.GetReportFooter()->SetValue(_T("TotalItem"), tmpStr);
	FormatCurrency(dblTotalAmount, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	tmpVal.Format(_T("%2.f"), dblTotalAmount);
	MoneyToString(tmpVal, tmpStr);
#ifdef UNICODE
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr+ _T(" \x111\x1ED3ng"));
#endif
	CString szDate;
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	
	
	rpt.GetReportFooter()->SetValue(_T("PatientSign"), szPatientName);

	rs.GetValue(_T("doctor"), tmpStr);

	rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(tmpStr));
	CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
	if(img)
	{
//		img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img->SetFixedHeight(false);
	}

	rpt.GetReportFooter()->SetValue(_T("IssueBy"), GetDoctorName(pMF->GetCurrentUser()));
	rpt.PrintPreview();
}


void CPrintForms::PM_PrintDrugDeliveryOrder(long nOrderID){

}

CString CPrintForms::GetDoctorName(CString szDoctorID)
{
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CString szSQL, szRet;
	CRecord rs(&pMF->m_db);
	szRet.Empty();
	szSQL.Format(_T("SELECT su_name ") \
				_T(" FROM sys_user") \
				_T(" WHERE lower(su_userid)=lower('%s')"), szDoctorID);

	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return _T("");
	rs.GetValue(_T("su_name"), szRet);
	return szRet;
}


void CPrintForms::PrintDeliveryOrder(long nOID){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CRecord ors(&pMF->m_db);
	CString szSQL, szStatus, szDoctor, tmpStr;
	szSQL.Format(_T("SELECT mt_status, mt_updatedby, mt_orderdate  ") \
		_T("FROM m_transaction ") \
		_T("WHERE mt_transaction_id=%ld "), nOID);
	ors.ExecSQL(szSQL);
	if(ors.IsEOF())
		return;
	ors.GetValue(_T("mt_status"), szStatus);
	ors.GetValue(_T("mt_updatedby"), szDoctor);
	ors.GetValue(_T("mt_orderdate"), tmpStr);
	if(szStatus == _T("O"))
		return;

	CString szPrintType;
	szSQL.Format(_T("SELECT hi_deliveryvote_printtype FROM hms_config "));
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF())
	{
	//	rs.GetValue(_T("hi_autoprint_deliveryvote"), m_szAutoPrintDeliveryVote);
		rs.GetValue(_T("hi_deliveryvote_printtype"), szPrintType);
	}
	if(szPrintType == _T("P"))
	{
		OnPrintDelivery(nOID, tmpStr.Left(10), szDoctor);
		return;
	}
	else
	{
		szSQL.Format(_T(" SELECT distinct Cast_Timestamp2Date(hpo_orderdate) as orderdate ") \
		_T(" FROM hms_pharmaorder ") \
		_T(" WHERE 	hpo_transaction_id=%ld ") \
		_T(" ORDER BY Cast_Timestamp2Date(hpo_orderdate) DESC "), nOID);
		rs.ExecSQL(szSQL);

		while(!rs.IsEOF()){
			rs.GetValue(_T("orderdate"), tmpStr);
			OnPrintDelivery(nOID, tmpStr, szDoctor);
			rs.MoveNext();
		}
	}

}
void CPrintForms::OnPrintDelivery(long nOID, LPCTSTR lpszDate, LPCTSTR lpszDoctor)
{
	CGuiMainFrame *pMF = (CGuiMainFrame *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szPrintType, szDeptID;
	CString szTransactionOrderNo;
	CRecord rs(&pMF->m_db);
	CRecord prs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	CRecord rsx(&pMF->m_db);
	int nRoomId, nBedIdx;
	CString	szOldDate, szNewDate, szReportDate, szPrintDate;
	CString szDepartment;
	CString szRoomName, szBedName;
	szSQL.Format(_T("SELECT hi_deliveryvote_printtype FROM hms_config "));
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF())
	{
	//	rs.GetValue(_T("hi_autoprint_deliveryvote"), m_szAutoPrintDeliveryVote);
		rs.GetValue(_T("hi_deliveryvote_printtype"), szPrintType);
	}

	if(szPrintType == _T("P"))
	{
		CReportSection* rptDetail = NULL;
		CString szItemID;
		long nOrderID;

		if(!rpt.Init(_T("Reports/HMS/HT_DAILYDELIVERYDRUGPORTRAIT.RPT"), true, false) )
			return;


		rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));

		tmpStr.Empty();
		//StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
		rptDetail->SetValue(_T("HealthService"), pMF->m_szHealthService);
		//StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
		rptDetail->SetValue(_T("HospitalName"), pMF->m_szHospitalName);

		szNewDate.Empty();
		szOldDate.Empty();
		szReportDate = rptDetail->GetValue(_T("ReportDate"));

		tmpStr.Format(szReportDate, CDate::Convert(szNewDate, yyyymmdd, ddmmyyyy));
		rptDetail->SetValue(_T("ReportDate"), tmpStr);
		rptDetail->SetValue(_T("SheetNo"), szTransactionOrderNo);
		//Page Header
		szSQL.Format(_T(" SELECT  hpo_docno as docno, ") \
			_T(" hpo_deptid as deptid, ") \
			_T(" hpo_orderid as orderid, ") \
			_T(" Cast_Timestamp2Date(hpo_orderdate) as orderdate, ") \
			_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
			_T(" hms_getage(Cast_Timestamp2Date(hd_admitdate), hp_birthdate) as age, ") \
			_T(" date_part('year', hp_birthdate) as yearofbirth, ") \
			_T(" hpo_roomid as roomid,	hpo_bedid as bedid ") \
			_T(" FROM hms_patient ") \
			_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno) ") \
			_T(" LEFT JOIN hms_ipharmaorder ON(hpo_docno=hd_docno)") \
			_T(" WHERE 	hpo_transaction_id=%ld ") \
			_T(" ORDER BY orderdate, roomid, bedid, pname, docno"), nOID);
		prs.ExecSQL(szSQL);
		

		int i = 1;
		CString szDate, szDoctorName;

		tmpStr = pMF->GetSysDate();
		
		szDoctorName = GetDoctorName(lpszDoctor)  ;

		szDate = rpt.GetGroupFooter(1)->GetValue(_T("PrintDate"));
		tmpStr = pMF->GetSysDateTime();
		szPrintDate.Format(szDate, tmpStr.Mid(10, 6), tmpStr.Mid(8,2), tmpStr.Mid(5, 2), tmpStr.Left(4));

		while(!prs.IsEOF())
		{

			prs.GetValue(_T("orderdate"), szNewDate);
			if(szOldDate.IsEmpty()){
				szOldDate = szNewDate;
				tmpStr.Format(szReportDate, CDate::Convert(szNewDate, yyyymmdd, ddmmyyyy));
				rptDetail->SetValue(_T("ReportDate"), tmpStr);

			}
			if(rptDetail && szNewDate != szOldDate){
				szOldDate = szNewDate;
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("PrintDate"), szPrintDate);
			//	rs.GetValue(_T("doctor"), tmpStr);
				rptDetail->SetValue(_T("Doctor"), szDoctorName);
				CReportItem *img = rptDetail->GetItem(_T("Signature"));
				if(img)
				{
	//				img->SetHBITMAP(pMF->GetSignatureImage(lpszDoctor));
					img->SetFixedHeight(false);
				}

				rptDetail->SetPageBreak(true);


				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
				tmpStr.Empty();
				//StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
				rptDetail->SetValue(_T("HealthService"), pMF->m_szHealthService);
				//StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
				rptDetail->SetValue(_T("HospitalName"), pMF->m_szHospitalName);

				tmpStr.Format(szReportDate, CDate::Convert(szNewDate, yyyymmdd, ddmmyyyy));
				rptDetail->SetValue(_T("ReportDate"), tmpStr);
				rptDetail->SetValue(_T("SheetNo"), szTransactionOrderNo);

				
			}

			rptDetail = rpt.AddDetail();
			rptDetail->GetItem(_T("1"))->SetBold(true);
			rptDetail->GetItem(_T("2"))->SetBold(true);
			rptDetail->GetItem(_T("3"))->SetBold(true);
			rptDetail->GetItem(_T("4"))->SetBold(true);
			
			prs.GetValue(_T("deptid"), szDepartment);
			rpt.GetReportHeader()->SetValue(_T("Department"), szDepartment);
			
			rptDetail->SetValue(_T("1"), CDate::Convert(szNewDate, yyyymmdd, ddmmyyyy));
			prs.GetValue(_T("pname"), tmpStr);
			CString szPatientName;
			szPatientName.Format(_T("[%s] %s  - %s"), prs.GetValue(_T("docno")), tmpStr,prs.GetValue(_T("age")));
			rptDetail->SetValue(_T("2"), szPatientName);

			prs.GetValue(_T("age"), tmpStr);
			rptDetail->SetValue(_T("age"), tmpStr);
			prs.GetValue(_T("yearofbirth"), tmpStr);
			rptDetail->SetValue(_T("yearofbirth"), tmpStr);


			prs.GetValue(_T("roomname"), nRoomId);
			rptDetail->SetValue(_T("3"), szRoomName);
			prs.GetValue(_T("bedname"), nBedIdx);
			rptDetail->SetValue(_T("4"), szBedName);

			prs.GetValue(_T("orderid"), nOrderID);

			szSQL.Format(_T(" SELECT 	product_id as itemid,") \
			_T(" product_name as name, product_uomname as unit, ") \
			_T(" 	sum(hpol_qtyorder) as qty, HMS_GETDOUSAGE(hpol_orderid, hpol_product_id) as usage ") \
			_T(" FROM hms_ipharmaorderline ") \
			_T(" LEFT JOIN m_productitem_view  ON(hpol_product_item_id=product_item_id)") \
			_T(" WHERE hpol_orderid=%ld ") \
			_T(" GROUP BY product_id, product_name, product_uomname, hpol_orderid, hpol_product_id ") \
			_T(" ORDER BY name"), nOrderID);
			rsl.ExecSQL(szSQL);
			int n = 1;
			while(!rsl.IsEOF()){
				rptDetail = rpt.AddDetail();
				tmpStr.Format(_T("%d"), n++);
				rptDetail->SetValue(_T("1"), tmpStr);
				rsl.GetValue(_T("name"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rsl.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("3"), tmpStr);
				rsl.GetValue(_T("qty"), tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				rsl.GetValue(_T("usage"), tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				rsl.MoveNext();
			}
			

			prs.MoveNext();
		}
		
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("PrintDate"), szPrintDate);
	//	rs.GetValue(_T("doctor"), tmpStr);
		rptDetail->SetValue(_T("Doctor"), szDoctorName);
		CReportItem *img = rptDetail->GetItem(_T("Signature"));
		if(img)
		{
//			img->SetHBITMAP(pMF->GetSignatureImage(lpszDoctor));
			img->SetFixedHeight(false);
		}
		rpt.PrintPreview();


	}
	else
	{
		if(!rpt.Init(_T("Reports/HMS/HT_DAILYDELIVERYDRUG.RPT")) )
			return;

		tmpStr.Empty();
		//StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
		rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_szHealthService);
		//StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
		rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		rpt.GetReportHeader()->SetValue(_T("Department"), szDepartment);
		tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(lpszDate, yyyymmdd, ddmmyyyy));
		rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("SheetNo"), szTransactionOrderNo);
		

		//Page Header

		szSQL.Format(_T(" SELECT distinct product_name as drugname, product_id as itemid ") \
		_T(" FROM hms_ipharmaorder ") \
		_T(" LEFT JOIN hms_ipharmaorderline ON(hpo_orderid=hpol_orderid)") \
		_T(" LEFT JOIN m_productitem_view  ON(hpol_product_item_id=product_item_id)") \
		_T(" WHERE 	hpo_transaction_id=%ld ") \
		_T(" 	and Cast_Timestamp2Date(hpo_orderdate)='%s' ORDER BY drugname "), nOID, lpszDate);
		int nCount = rs.ExecSQL(szSQL);
		if(rs.IsEOF())
			return;
		CString szName;
		int i =1;
		while(!rs.IsEOF())
		{
			szName.Format(_T("%d"), i++);
			rs.GetValue(_T("drugname"), tmpStr);
			rpt.GetPageHeader()->SetValue(szName, tmpStr);
			rs.MoveNext();
		}
		for (; i < 37; i++){
			szName.Format(_T("%d"), i++);
			rpt.GetPageHeader()->SetValue(szName, _T(""));
		}
		szSQL.Format(_T(" SELECT  hpo_docno as docno, hpo_orderid as orderid, ") \
			_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
			_T(" hms_getage(Cast_Timestamp2Date(hd_admitdate), hp_birthdate) as age, ") \
			_T(" date_part('year', hp_birthdate) as yearofbirth, ") \
			_T(" 	hpo_deptid as deptid, hpo_roomid as roomid, hpo_bedid as bedid ") \
			_T(" FROM hms_patient ") \
			_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno) ") \
			_T(" LEFT JOIN hms_ipharmaorder ON(hpo_docno=hd_docno)") \
			_T(" WHERE 	hpo_transaction_id=%ld ") \
			_T(" 	and Cast_Timestamp2Date(hpo_orderdate)='%s'") \
			_T(" ORDER BY pname, docno"), nOID, lpszDate);
		prs.ExecSQL(szSQL);
		CReportSection* rptDetail;
		CString szItemID;
		long nOrderID;
		i = 1;
		while(!prs.IsEOF())
		{
			rptDetail = rpt.AddDetail();
			szName.Format(_T("%d"), i++);
			rptDetail->SetValue(_T("Index"), szName);
			prs.GetValue(_T("docno"), tmpStr);
			rptDetail->SetValue(_T("DocNo"), tmpStr);
			prs.GetValue(_T("pname"), tmpStr);
			rptDetail->SetValue(_T("PatientName"), tmpStr);
			prs.GetValue(_T("age"), tmpStr);
			rptDetail->SetValue(_T("age"), tmpStr);
			prs.GetValue(_T("yearofbirth"), tmpStr);
			rptDetail->SetValue(_T("yearofbirth"), tmpStr);

			prs.GetValue(_T("department"), szDepartment);
			prs.GetValue(_T("roomname"), szRoomName);
			rptDetail->SetValue(_T("roomname"), szRoomName);
			prs.GetValue(_T("bedname"), szBedName);
			rptDetail->SetValue(_T("bedname"), szBedName);

			rptDetail->SetValue(_T("SignLabel"), _T(""));
			
			prs.GetValue(_T("orderid"), nOrderID);

			szSQL.Format(_T(" SELECT 	product_id as itemid,") \
			_T(" 	sum(hpol_qtyorder) as qty") \
			_T(" FROM hms_ipharmaorderline ") \
			_T(" LEFT JOIN m_productitem_view  ON(hpol_product_item_id=product_item_id)") \
			_T(" WHERE hpol_orderid=%ld ") \
			_T(" GROUP BY product_id"), nOrderID);
			rsl.ExecSQL(szSQL);
			while(!rsl.IsEOF()){
				int n = 0;
				rsl.GetValue(_T("itemid"), szItemID);
				rs.MoveFirst();
				while(!rs.IsEOF())
				{
					n++;
					rs.GetValue(_T("itemid"), tmpStr);
					if(tmpStr == szItemID){
						break;
					}
					
					rs.MoveNext();
				}
				szName.Format(_T("%d"), n);
				rsl.GetValue(_T("qty"), tmpStr);
				rptDetail->SetValue(szName, tmpStr);
				rsl.MoveNext();
			}
			prs.MoveNext();
		}

		CString szDate;
		tmpStr = pMF->GetSysDate();
		szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
		rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
		rpt.GetReportFooter()->SetValue(_T("Doctor"), GetDoctorName(lpszDoctor));
		CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
		if(img)
		{
	//		img->SetHBITMAP(pMF->GetSignatureImage(lpszDoctor));
			img->SetFixedHeight(false);
		}
		rpt.PrintPreview();
	}
}

CString CPrintForms::GetReportName(){
	return m_szReportName;
}

void CPrintForms::OnPrintDailyViewDetail(LPCTSTR lpszOrderNo, LPCTSTR lpszDate, LPCTSTR lpszTransactionString, long nProduct_ID, long nProductItem_ID, LPCTSTR lpszOrderType, int nStorage_ID, CString szOrgID){
	CGuiMainFrame *pMF = (CGuiMainFrame *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL;
	CString szTransactionOrderNo;
	CRecord rs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	int nProduct_item_id;
	CString	szOldDate, szNewDate, szReportDate, szPrintDate;
	CString szDepartment;
	CString szRoomName, szBedName;
	CString szTableName;
	CReportSection* rptDetail = NULL;
	
	if(!rpt.Init(_T("Reports/HMS/HT_DAILYDELIVERYDRUGVIEWDETAIL.RPT"), true, false) )
		return;
	
	if(szOrgID == _T("TM"))
		szTableName = _T("hms_ipharmaorder");
	else
		szTableName = _T("hms_pharmaorder");

	rptDetail = rpt.GetReportHeader();

	tmpStr.Empty();
	StringUpper( pMF->m_szHealthService, tmpStr);
	rptDetail->SetValue(_T("HealthService"), tmpStr);
	StringUpper(pMF->m_szHospitalName, tmpStr);
	rptDetail->SetValue(_T("HospitalName"), tmpStr);

	szNewDate.Empty();
	szOldDate.Empty();
	szReportDate = rptDetail->GetValue(_T("ReportDate"));

	tmpStr.Format(szReportDate, CDate::Convert(szNewDate, yyyymmdd, ddmmyyyy));
	rptDetail->SetValue(_T("ReportDate"), tmpStr);
	rptDetail->SetValue(_T("SheetNo"), szTransactionOrderNo);
	//Page Header	
	CString szWhere;
	szWhere.Empty();
	if(nProductItem_ID > 0)
		szWhere.Format(_T(" and product_item_id=%ld "), nProductItem_ID);

	szSQL.Format(_T(" SELECT Get_StorageName(hpo_storage_id) as storagename, hpo_storage_id,") \
	_T("	product_id, product_item_id, product_name as product_name, ") \
	_T(" 	product_uomname as product_unit,") \
	_T(" 	hpol_unitprice as product_taxprice,") \
	_T(" 	sum(hpol_qtyorder) as product_qtyorder,") \
	_T(" 	sum(hpol_qtyissue) as product_qtyissue,") \
	_T(" 	sum(hpol_qtyissue*hpol_unitprice) as product_amount ") \
	_T(" FROM %s ") \
	_T(" LEFT JOIN %sline ON(hpo_orderid=hpol_orderid)") \
	_T(" LEFT JOIN m_productitem_view ON(product_item_id=hpol_product_item_id) ") \
	_T(" WHERE hpo_ordertype in(%s) ") \
	_T("	AND hpo_transaction_id in(%s)  ") \
	_T("	AND hpo_storage_id = %ld AND product_id=%ld  %s ") \
	_T("	AND hpo_org_id NOT IN('SMM') ") \
	_T(" GROUP BY hpo_storage_id, product_id, product_item_id, ") \
	_T("	product_name, product_uomname, hpol_unitprice ") \
	_T(" ORDER BY hpo_storage_id, product_name, product_unit, product_taxprice  ") , 
	szTableName, szTableName,
	lpszOrderType,
	lpszTransactionString, 
	nStorage_ID, 
	nProduct_ID,
	szWhere
	);
	
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
	if(rs.IsEOF()){
		ShowMessageBox(_T("No data"));
		return ;
	}

	rs.GetValue(_T("product_item_id"), nProduct_item_id);	
	rptDetail->SetValue(_T("Storage"), rs.GetValue(_T("storagename")));
	rptDetail->SetValue(_T("SheetNo"), lpszOrderNo);
	rptDetail->SetValue(_T("DrugName"), rs.GetValue(_T("product_name")));
	rptDetail->SetValue(_T("Unit"), rs.GetValue(_T("product_unit")));
	rptDetail->SetValue(_T("Qty"), rs.GetValue(_T("product_qtyorder")));

	szSQL.Format(_T(" SELECT DISTINCT orderdate,") \
	_T("   pname,") \
	_T("   docno,") \
	_T("   recordno,") \
	_T("   get_departmentname(deptid) deptname,") \
	_T("   cardno, ") \
	_T("   objectname, ") \
	_T("   address, ") 
	_T("   yearofbirth, ") \
	_T("   hpo_storage_id,") \
	_T("   product_name,") \
	_T("   product_unit,") \
	_T("   product_taxprice,") \
	_T("   SUM(product_qtyorder) AS QtyOrder,") \
	_T("   SUM(product_qtyissue) AS QtyIssue") \
	_T(" FROM") \
	_T("   (SELECT hpo_docno      AS docno,") \
	_T("	 hcr_recordno recordno,") \
	_T("     hpo_deptid           AS deptid,") \
	_T("     hpo_orderid          AS orderid,") \
	_T("     Cast_Timestamp2Date(hpo_orderdate) AS orderdate,") \
	_T("     trim(hp_surname") \
	_T("     ||' '") \
	_T("     ||hp_midname") \
	_T("     ||' '") \
	_T("     ||hp_firstname)                               AS pname,") \
	_T("     hms_getage(Cast_Timestamp2Date(hd_admitdate), hp_birthdate) AS age,") \
	_T("     extract(year from hp_birthdate)               AS yearofbirth,") \
	_T(" 	 hms_getaddress(hp_provid,hp_distid,hp_villid) as address,") \
	_T(" (SELECT hrl_name") \
	_T("     FROM hms_roomlist") \
	_T("     WHERE hrl_id  =hpo_roomid") \
	_T("     AND hrl_deptid=hpo_deptid") \
	_T("     ) AS roomname,") \
	_T("     (SELECT hbl_name") \
	_T("     FROM hms_bedlist") \
	_T("     WHERE hbl_id  = hpo_bedid") \
	_T("     AND hbl_roomid=hpo_roomid") \
	_T("     AND hbl_deptid=hpo_deptid") \
	_T("     ) AS bedname,") \
	_T("	 ho_desc as objectname, ") \
	_T(" 	 hd_cardno as cardno,") \
	_T("     hpo_storage_id,") \
	_T("     product_id,") \
	_T("     product_item_id,") \
	_T("     product_name    AS product_name,") \
	_T("     product_uomname AS product_unit,") \
	_T("     hpol_unitprice  AS product_taxprice,") \
	_T("     hpol_qtyorder   AS product_qtyorder,") \
	_T("     hpol_qtyissue   AS product_qtyissue") \
	_T("   FROM %s ") \
	_T("   LEFT JOIN %sline") \
	_T("   ON(hpo_orderid=hpol_orderid)") \
	_T("   LEFT JOIN m_productitem_view") \
	_T("   ON(product_item_id   =hpol_product_item_id)") \
	_T("   LEFT JOIN hms_doc") \
	_T("   ON(hpo_docno=hd_docno)") \
	_T("   LEFT JOIN hms_clinical_record ON (hd_docno = hcr_docno)")
	_T(" LEFT JOIN hms_patient ") \
	_T("   ON(hd_patientno=hp_patientno)") \
	_T("   LEFT JOIN hms_object ON(ho_id=hd_object) ") \
	_T("   LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx) ") \
	_T("   WHERE hpo_ordertype  IN(%s)") \
	_T("   AND hpo_transaction_id IN(%s) AND hpo_org_id NOT IN('SMM')") \
	_T(" and hpo_storage_id=%ld ") \
	_T("   AND product_id  =%ld  %s ") \
	_T("   )") \
	_T(" GROUP BY orderdate,") \
	_T("   pname,") \
	_T("   docno,") \
	_T("   recordno,") \
	_T("   deptid,") \
	_T("   cardno, ") \
	_T("   objectname, ") \
	_T("   address, ") \
	_T("   yearofbirth, ") \
	_T("   hpo_storage_id,") \
	_T("   product_name,") \
	_T("   product_unit,") \
	_T("   product_taxprice") \
	_T(" ORDER BY docno,") \
	_T("   pname,") \
	_T("   orderdate,") \
	_T("   hpo_storage_id,") \
	_T("   product_name,") \
	_T("   product_unit,") \
	_T("   product_taxprice"), 
	szTableName, szTableName,
	lpszOrderType, lpszTransactionString, nStorage_ID, nProduct_ID, szWhere);

_fmsg(_T("%s"), szSQL);
	rsl.ExecSQL(szSQL);
	int nIdx=0;
	while(!rsl.IsEOF())
	{
		nIdx++;
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("0"), tmpStr);
		rsl.GetValue(_T("OrderDate"), tmpStr);		
		rptDetail->SetValue(_T("1"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rsl.GetValue(_T("pname"), tmpStr);		
		rptDetail->SetValue(_T("2"), tmpStr);		
		rsl.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rsl.GetValue(_T("recordno"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rsl.GetValue(_T("deptname"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rsl.GetValue(_T("QtyOrder"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rsl.GetValue(_T("CardNo"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rsl.GetValue(_T("ObjectName"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rsl.GetValue(_T("Address"), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rsl.MoveNext();
	}
	
	CString szDate, szSysDate;	
	szSysDate = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Mid(11, 5),szSysDate.Mid(8,2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
}

void CPrintForms::PM_PrintCabinetUsageDetail(LPCTSTR lpszOrderNo, LPCTSTR lpszDate, LPCTSTR lpszTransactionString, long nProduct_id, LPCTSTR lpszOrderType, int nStorage_ID, CString szOrgID){
	CGuiMainFrame *pMF = (CGuiMainFrame *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL;
	CString szTransactionOrderNo;
	CRecord rs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	int nProduct_item_id;
	CString	szOldDate, szNewDate, szReportDate, szPrintDate;
	CString szDepartment;
	CString szRoomName, szBedName;
	CReportSection* rptDetail = NULL;
	
	if(!rpt.Init(_T("Reports/HMS/HT_DAILYDELIVERYDRUGVIEWDETAIL.RPT"), true, false) )
		return;
	
	rptDetail = rpt.GetReportHeader();

	tmpStr.Empty();
	StringUpper( pMF->m_szHealthService, tmpStr);
	rptDetail->SetValue(_T("HealthService"), tmpStr);
	StringUpper(pMF->m_szHospitalName, tmpStr);
	rptDetail->SetValue(_T("HospitalName"), tmpStr);

	szNewDate.Empty();
	szOldDate.Empty();
	szReportDate = rptDetail->GetValue(_T("ReportDate"));

	tmpStr.Format(szReportDate, CDate::Convert(szNewDate, yyyymmdd, ddmmyyyy));
	rptDetail->SetValue(_T("ReportDate"), tmpStr);
	rptDetail->SetValue(_T("SheetNo"), szTransactionOrderNo);
	//Page Header	

	szSQL.Format(_T("SELECT Get_storagename(mt_storage_id)        AS storagename, ") \
	_T("       mt_storage_id, ") \
	_T("	   get_storagename(mt_storage_to_id) AS storagetoname,")
	_T("       product_id, ") \
	_T("       product_name                          AS product_name, ") \
	_T("	   case when hmt_product_id <> hmt_suppproduct_id and hmt_suppproduct_id > 0 then (select distinct mp_name from m_product where mp_product_id = hmt_suppproduct_id) else cast('' as nvarchar2(254)) end as xproduct_name, ") \
	_T("       product_uomname                       AS product_unit, ") \
	_T("       Sum(hmt_qtyissue)                     AS product_qtyorder, ") \
	_T("       Sum(hmt_qtyissue)                     AS product_qtyissue, ") \
	_T("       Sum(hmt_qtyissue * product_unitprice) AS product_amount ") \
	_T("FROM   hms_medical_transaction_view ") \
	_T("       LEFT JOIN m_product_view ") \
	_T("              ON( product_id = hmt_product_id ) ") \
	_T("       LEFT JOIN m_transaction ") \
	_T("              ON ( mt_transaction_id = hmt_reftransaction_id ) ") \
	_T("WHERE hmt_reftransaction_id = %s ") \
	_T("       AND (hmt_suppproduct_id = %d or hmt_product_id=%ld) ") \
	_T("       AND mt_storage_id = %d ") \
	_T("GROUP  BY mt_storage_id, ") \
	_T("		  mt_storage_to_id,") \
	_T("          product_id, ") \
	_T("          hmt_product_id, ") \
	_T("          hmt_suppproduct_id, ") \
	_T("          product_name, ") \
	_T("          product_uomname ") \
	_T("ORDER  BY mt_storage_id, ") \
	_T("          product_name, ") \
	_T("          product_unit "), 
	lpszTransactionString, nProduct_id, nProduct_id, nStorage_ID);
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
	if(rs.IsEOF()){
		ShowMessageBox(_T("No data"));
		return ;
	}

	CString  szName;
	rs.GetValue(_T("product_item_id"), nProduct_item_id);	
	rptDetail->SetValue(_T("Storage"), rs.GetValue(_T("storagename")));
	rptDetail->SetValue(_T("ToStorage"), rs.GetValue(_T("storagetoname")));
	rptDetail->SetValue(_T("SheetNo"), lpszOrderNo);
	rs.GetValue(_T("xproduct_name"), tmpStr);
	rs.GetValue(_T("product_name"), szName);
	if(!tmpStr.IsEmpty())
		szName.AppendFormat(_T(" ->[%s]"), tmpStr);
	rptDetail->SetValue(_T("DrugName"), szName);
	rptDetail->SetValue(_T("Unit"), rs.GetValue(_T("product_unit")));
	rptDetail->SetValue(_T("Qty"), rs.GetValue(_T("product_qtyorder")));

	if(szOrgID != _T("EM"))
	{
	szSQL.Format(_T("SELECT DISTINCT orderdate, ") \
	_T("                roomname, ") \
	_T("                bedname, ") \
	_T("                pname, ") \
	_T("                docno, ") \
	_T("				recordno,") \
	_T("				get_departmentname(deptid) deptname,") \
	_T("                cardno, ") \
	_T("                objectname, ") \
	_T("                address, ") \
	_T("                yearofbirth, ") \
	_T("                hpo_storage_id, ") \
	_T("                product_name, ") \
	_T("                product_unit, ") \
	_T("                Sum(product_qtyorder) AS QtyOrder, ") \
	_T("                Sum(product_qtyissue) AS QtyIssue ") \
	_T("FROM   (SELECT hmt_docno                                       AS docno, ") \
	_T("			   hcr_recordno recordno,") \
	_T("               (select htr_deptid from hms_treatment_record where htr_docno = hmt_docno and htr_status = 'I') deptid, ") \
	_T("               hmt_medical_transaction_id                      AS orderid, ") \
	_T("               Cast_Timestamp2Date(hpo_orderdate)                            AS orderdate, ") \
	_T("               Trim(hp_surname ") \
	_T("                    || ' ' ") \
	_T("                    || hp_midname ") \
	_T("                    || ' ' ") \
	_T("                    || hp_firstname)                           AS pname, ") \
	_T("               Hms_getage(Cast_Timestamp2Date(hd_admitdate), hp_birthdate)   AS age, ") \
	_T("               Extract(YEAR FROM hp_birthdate)                 AS yearofbirth, ") \
	_T("               Hms_getaddress(hp_provid, hp_distid, hp_villid) AS address, ") \
	_T("               (SELECT hrl_name ") \
	_T("                FROM   hms_roomlist ") \
	_T("                WHERE  hrl_id = hpo_roomid ") \
	_T("                       AND hrl_deptid = hpo_deptid)            AS roomname, ") \
	_T("               (SELECT hbl_name ") \
	_T("                FROM   hms_bedlist ") \
	_T("                WHERE  hbl_id = hpo_bedid ") \
	_T("                       AND hbl_roomid = hpo_roomid ") \
	_T("                       AND hbl_deptid = hpo_deptid)            AS bedname, ") \
	_T("               ho_desc                                         AS objectname, ") \
	_T("               hd_cardno                                       AS cardno, ") \
	_T("               hpo_storage_id, ") \
	_T("               product_id, ") \
	_T("               product_name                                    AS product_name, ") \
	_T("               product_uomname                                 AS product_unit, ") \
	_T("               hmt_qtyissue                                    AS product_qtyorder, ") \
	_T("               hmt_qtyissue                                    AS product_qtyissue ") \
	_T("        FROM   hms_patient ") \
	_T("               LEFT JOIN hms_doc ") \
	_T("                      ON( hd_patientno = hp_patientno ) ") \
	_T("			   LEFT JOIN hms_clinical_record ON (hcr_docno = hd_docno)") \
	_T("               LEFT JOIN hms_medical_transaction_view ") \
	_T("                      ON( hmt_docno = hd_docno ) ") \
	_T("               LEFT JOIN m_product_view ") \
	_T("                      ON( product_id = hmt_suppproduct_id ) ") \
	_T("               LEFT JOIN hms_object ") \
	_T("                      ON( ho_id = hd_object ) ") \
	_T("               LEFT JOIN hms_card ") \
	_T("                      ON( hc_patientno = hd_patientno ") \
	_T("                          AND hc_idx = hd_cardidx ) ") \
	_T("               LEFT JOIN hms_ipharmaorder ") \
	_T("                      ON ( hpo_orderid = hmt_orderid ) ") \
	_T("        WHERE  hmt_reftransaction_id = %s ") \
	_T("              and (hmt_suppproduct_id = %ld or hmt_product_id=%ld) ) tbl ") \
	_T("GROUP  BY orderdate, ") \
	_T("          roomname, ") \
	_T("          bedname, ") \
	_T("          pname, ") \
	_T("          docno, ") \
	_T("		  recordno,") \
	_T("		  deptid,") \
	_T("          cardno, ") \
	_T("          objectname, ") \
	_T("          address, ") \
	_T("          yearofbirth, ") \
	_T("          hpo_storage_id, ") \
	_T("          product_name, ") \
	_T("          product_unit ") \
	_T("ORDER  BY docno, ") \
	_T("          roomname, ") \
	_T("          bedname, ") \
	_T("          pname, ") \
	_T("          orderdate, ") \
	_T("          hpo_storage_id, ") \
	_T("          product_name, ") \
	_T("          product_unit ")
	, lpszTransactionString, nProduct_id, nProduct_id);
	}
	else
	{
szSQL.Format(_T("SELECT DISTINCT orderdate, ") \
	_T("                roomname, ") \
	_T("                bedname, ") \
	_T("                pname, ") \
	_T("                docno, ") \
	_T("				get_departmentname(deptid) deptname,") \
	_T("                cardno, ") \
	_T("                objectname, ") \
	_T("                address, ") \
	_T("                yearofbirth, ") \
	_T("                hpo_storage_id, ") \
	_T("                product_name, ") \
	_T("                product_unit, ") \
	_T("                Sum(product_qtyorder) AS QtyOrder, ") \
	_T("                Sum(product_qtyissue) AS QtyIssue ") \
	_T("FROM   (SELECT hmt_docno                                       AS docno, ") \
	_T("               (select htr_deptid from hms_treatment_record where htr_docno = hmt_docno and htr_status = 'I') AS deptid, ") \
	_T("               hmt_medical_transaction_id                      AS orderid, ") \
	_T("               Cast_Timestamp2Date(hpo_orderdate)                            AS orderdate, ") \
	_T("               Trim(hp_surname ") \
	_T("                    || ' ' ") \
	_T("                    || hp_midname ") \
	_T("                    || ' ' ") \
	_T("                    || hp_firstname)                           AS pname, ") \
	_T("               Hms_getage(Cast_Timestamp2Date(hd_admitdate), hp_birthdate)   AS age, ") \
	_T("               Extract(YEAR FROM hp_birthdate)                 AS yearofbirth, ") \
	_T("               Hms_getaddress(hp_provid, hp_distid, hp_villid) AS address, ") \
	_T("               (SELECT hrl_name ") \
	_T("                FROM   hms_roomlist ") \
	_T("                WHERE  hrl_id = hpo_roomid ") \
	_T("                       AND hrl_deptid = hpo_deptid)            AS roomname, ") \
	_T("               (SELECT hbl_name ") \
	_T("                FROM   hms_bedlist ") \
	_T("                WHERE  hbl_id = hpo_bedid ") \
	_T("                       AND hbl_roomid = hpo_roomid ") \
	_T("                       AND hbl_deptid = hpo_deptid)            AS bedname, ") \
	_T("               ho_desc                                         AS objectname, ") \
	_T("               hd_cardno                                       AS cardno, ") \
	_T("               hpo_storage_id, ") \
	_T("               product_id, ") \
	_T("               product_name                                    AS product_name, ") \
	_T("               product_uomname                                 AS product_unit, ") \
	_T("               hmt_qtyissue                                    AS product_qtyorder, ") \
	_T("               hmt_qtyissue                                    AS product_qtyissue ") \
	_T("        FROM   hms_patient ") \
	_T("               LEFT JOIN hms_doc ") \
	_T("                      ON( hd_patientno = hp_patientno ) ") \
	_T("               LEFT JOIN hms_medical_transaction_view ") \
	_T("                      ON( hmt_docno = hd_docno ) ") \
	_T("               LEFT JOIN m_product_view ") \
	_T("                      ON( product_id = hmt_suppproduct_id ) ") \
	_T("               LEFT JOIN hms_object ") \
	_T("                      ON( ho_id = hd_object ) ") \
	_T("               LEFT JOIN hms_card ") \
	_T("                      ON( hc_patientno = hd_patientno ") \
	_T("                          AND hc_idx = hd_cardidx ) ") \
	_T("               LEFT JOIN hms_pharmaorder ") \
	_T("                      ON ( hpo_orderid = hmt_orderid ) ") \
	_T("        WHERE  hmt_reftransaction_id = %s") \
	_T("               and  (hmt_suppproduct_id = %ld or hmt_product_id=%ld) ) tbl ") \
	_T("GROUP  BY orderdate, ") \
	_T("		  deptid,") \
	_T("          roomname, ") \
	_T("          bedname, ") \
	_T("          pname, ") \
	_T("          docno, ") \
	_T("          cardno, ") \
	_T("          objectname, ") \
	_T("          address, ") \
	_T("          yearofbirth, ") \
	_T("          hpo_storage_id, ") \
	_T("          product_name, ") \
	_T("          product_unit ") \
	_T("ORDER  BY docno, ") \
	_T("          roomname, ") \
	_T("          bedname, ") \
	_T("          pname, ") \
	_T("          orderdate, ") \
	_T("          hpo_storage_id, ") \
	_T("          product_name, ") \
	_T("          product_unit "),
	 lpszTransactionString, nProduct_id, nProduct_id);
	}
_fmsg(_T("%s"), szSQL);
	rsl.ExecSQL(szSQL);
	int nIdx=0;
	while(!rsl.IsEOF())
	{
		nIdx++;
		rptDetail = rpt.AddDetail();		
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("0"), tmpStr);
		rsl.GetValue(_T("OrderDate"), tmpStr);		
		rptDetail->SetValue(_T("1"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
		rsl.GetValue(_T("pname"), tmpStr);		
		rptDetail->SetValue(_T("2"), tmpStr);		
		rsl.GetValue(_T("yearofbirth"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rsl.GetValue(_T("roomname"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rsl.GetValue(_T("bedname"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rsl.GetValue(_T("QtyOrder"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rsl.GetValue(_T("CardNo"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rsl.GetValue(_T("ObjectName"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rsl.GetValue(_T("Address"), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);
		rptDetail->SetValue(_T("10"), rsl.GetValue(_T("docno")));
		rptDetail->SetValue(_T("11"), rsl.GetValue(_T("recordno")));
		rptDetail->SetValue(_T("12"), rsl.GetValue(_T("deptname")));
		rsl.MoveNext();
	}
	
	CString szDate, szSysDate;	
	szSysDate = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Mid(11, 5),szSysDate.Mid(8,2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
}



void CPrintForms::OnPrintProcessedPrescription(LPCTSTR lpszFilter){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CReport rpt;
	CReportSection *rptDetail;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nIdx = 1;
	if (!rpt.Init(_T("Reports\\HMS\\PMS_DANHSACHDUYETDONTHUOC.RPT")))
		return;
	szSQL.Format(_T(" SELECT") \
				_T("   hpo_processeddate              AS approveddate,") \
				_T("   hpo_orderid                    AS orderid,") \
				_T("   hpo_docno                      AS docno,") \
				_T("   Get_patientname(hpo_docno)     AS patname,") \
				_T("   Get_departmentname(hpo_deptid) AS dept,") \
				_T("   hpo_amount                     AS amount") \
				_T(" FROM   hms_ipharmaorder "));
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), tmpStr);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("approveddate")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("orderid")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("docno")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("patname")));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("dept")));
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("amount")));
		rs.MoveNext();
	}
	rpt.PrintPreview();
}

void CPrintForms::TM_PrintDrugReturnPatientList(long nOrderID)
{
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere, tmpStr, szOldDoc, szNewDoc, szDate;
	CReportSection* rptHeader, *rptDetail;
	int nIdx = 1;
	double nTtlAmt = 0, nGrpAmt = 0, nAmt = 0;
	if (!rpt.Init(_T("Reports\\HMS\\PMS_DANHSACHTHUOCTRALAI.RPT")))
		return;
	szSQL.Format(_T("SELECT get_storagename(mt_storage_to_id) as stockname, ") \
				_T("       get_departmentname(mt_department_id) as deptname, ") \
				_T("       mt_approveddate as approveddate ") \
				_T("FROM   m_transaction ") \
				_T("WHERE  mt_transaction_id = %ld AND mt_status <> 'O' AND mt_doctype = 'DRO'"), nOrderID);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."), MB_OK);
		return;
	}
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	rptHeader->SetValue(_T("PrintDate"), CDate::Convert(rs.GetValue(_T("approveddate")), yyyymmdd, ddmmyyyy));
	rptHeader->SetValue(_T("StockName"), rs.GetValue(_T("stockname")));
	rptHeader->SetValue(_T("Department"), rs.GetValue(_T("deptname")));
	szSQL.Format(_T("SELECT hpolr_docno                  AS docno, ") \
				_T("       Get_patientname(hpolr_docno) AS patname, ") \
				_T("       product_name                AS pname, ") \
				_T("       product_uomname             AS uom, ") \
				_T("       hpolr_unitprice              AS price, ") \
				_T("       Sum(hpolr_qtyreverse)        AS qty, ") \
				_T("	   SUM(hpolr_qtyreverse)*hpolr_unitprice AS amt ") \
				_T("FROM   hms_ipharmaorderline_r ") \
				_T("LEFT JOIN m_product_view ON ( hpolr_product_id = product_id ) ") \
				_T("WHERE  hpolr_retorder_id = %ld ") \
				_T("GROUP  BY hpolr_docno, ") \
				_T("          product_name, ") \
				_T("          product_uomname, ") \
				_T("          hpolr_unitprice ") \
				_T("ORDER  BY docno, ") \
				_T("          pname "), nOrderID);
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("docno"), szNewDoc);
		if (szNewDoc != szOldDoc)
		{
			if (nGrpAmt > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("TotalGroup"), _T("T\x1ED5ng nh\xF3m"));
				nTtlAmt += nGrpAmt;
				rptDetail->SetValue(_T("s6"), double2str(nGrpAmt));
				nGrpAmt = 0;
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			tmpStr.Format(_T("[%s] %s"), rs.GetValue(_T("docno")), rs.GetValue(_T("patname")));
			rptDetail->SetValue(_T("PName"), tmpStr);
			szOldDoc = szNewDoc;
			nIdx = 1;
		}
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIdx++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("pname")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("uom")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("price")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("qty")));
		rs.GetValue(_T("amt"), nAmt);
		nGrpAmt += nAmt;
		rptDetail->SetValue(_T("6"), double2str(nAmt));
		rs.MoveNext();
	}
	if (nGrpAmt > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), _T("T\x1ED5ng nh\xF3m"));
		nTtlAmt += nGrpAmt;
		rptDetail->SetValue(_T("s6"), double2str(nGrpAmt));
	}
	if (nTtlAmt > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("TotalGroup"), _T("T\x1ED5ng \x63\x1ED9ng"));
		rptDetail->SetValue(_T("s6"), double2str(nTtlAmt));
	}
	szDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szDate.Mid(8,2), szDate.Mid(5, 2), szDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}

void CPrintForms::FM_PrintGeneralApprovedInOrder(long nTransactionID, bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;

	CString szSQL, szWhere;
	CString tmpStr, szDate, szSysDate, tmpStr2, szFacList;
	CString szObject, szGroups, szObjectName, szType, szTypeName, szTableName;
	CString szStocks, szStockNames;
	CRecord rs(&pMF->m_db);
	szType.Empty();
	szTypeName.Empty();

	int nCount = 0;

	//if (!rpt.Init(_T("Reports/HMS/PTTT/QUYETTOANSUDUNGTHUOC-VTYTBN.RPT")))
	if (!rpt.Init(_T("Reports/HMS/PTTT/HF_PHIEULINHDV.RPT")))
	{
		return;
	}
	szSQL.Format(_T("SELECT mt_approveddate appdate, ") \
				_T("       get_storagename(mt_storage_id) stockname, ") \
				_T("	   get_departmentname(mt_department_to_id) deptname,") \
				_T("	   mt_orderno orderno") \
				_T(" FROM   m_transaction ") \
				_T(" WHERE  mt_transaction_id = %ld"), nTransactionID);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		return;
	}

	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	rpt.GetReportHeader()->SetValue(_T("Department"), rs.GetValue(_T("deptname")));
	rpt.GetReportHeader()->SetValue(_T("OrderNo"), rs.GetValue(_T("orderno")));
	//tmpStr.Format(_T("\x42\xE1o \x63\xE1o quy\x1EBFt to\xE1n s\x1EED \x64\x1EE5ng thu\x1ED1\x63"));
	//rpt.GetReportHeader()->SetValue(_T("ReportName"),tmpStr);
	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
	              CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), CDate::Convert(rs.GetValue(_T("appdate")), yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("StockName"), rs.GetValue(_T("stockname")));
	rpt.GetReportHeader()->SetValue(_T("Object"), szObjectName);
	rpt.GetReportHeader()->SetValue(_T("DrugType"), szTypeName);
	rpt.GetReportHeader()->SetValue(_T("TableName"), szTableName);
	

	szSQL.Format(_T("SELECT pname, ") \
				_T("       age, ") \
				_T("       docno, ") \
				_T("       Sum(drugamt)     drugamt, ") \
				_T("       Sum(materialamt) materialamt, ") \
				_T("       Sum(amount)      amount ") \
				_T("FROM   (SELECT Get_patientname(hpol_docno)    pname, ") \
				_T("               Hms_getagebydoc(hpol_docno)    age, ") \
				_T("               hpol_docno                     docno, ") \
				_T("               CASE ") \
				_T("                 WHEN product_org_id = 'PM' THEN hpol_qtyorder * hpol_unitprice ") \
				_T("                 ELSE 0 ") \
				_T("               END                            AS drugamt, ") \
				_T("               CASE ") \
				_T("                 WHEN product_org_id = 'MA' THEN hpol_qtyorder * hpol_unitprice ") \
				_T("                 ELSE 0 ") \
				_T("               END                            AS materialamt, ") \
				_T("               hpol_qtyorder * hpol_unitprice AS amount ") \
				_T("        FROM   hms_ipharmaorder ") \
				_T("               LEFT JOIN hms_ipharmaorderline ") \
				_T("                      ON ( hpo_orderid = hpol_orderid ) ") \
				_T("               LEFT JOIN m_product_view ") \
				_T("                      ON ( hpol_product_id = product_id ) ") \
				_T("        WHERE  hpo_transaction_id = %ld) tbl ") \
				_T("GROUP  BY pname, ") \
				_T("          age, ") \
				_T("          docno "), nTransactionID);


	//_msg(_T("%s"), szSQL);
			
		CReportSection* rptDetail;
		CString szOldLine, szNewLine, szAmount;
		
		long double grpCost[8];
		long double ttlCost[8];
		double cost = 0;
		int nItem = 1;

		_fmsg(_T("%s"), szSQL);

		rs.ExecSQL(szSQL);

		if (rs.IsEOF())
		{
			_msg(_T("%s"), szSQL);
			ShowMessageBox(_T("No Data"), MB_OK);
			return;
		}

		for (int i = 0; i < 8; i++)
		{
			grpCost[i] = 0;
			ttlCost[i] = 0;
		}
				
		while(!rs.IsEOF())
		{

			rptDetail = rpt.AddDetail();
			tmpStr.Format(_T("%d"), nItem++);
			rptDetail->SetValue(_T("1"), tmpStr);
			rs.GetValue(_T("id"), tmpStr);				
			rs.GetValue(_T("pname"), tmpStr);
			rptDetail->SetValue(_T("2"), tmpStr);
			rs.GetValue(_T("docno"), tmpStr);
			rptDetail->SetValue(_T("3"), tmpStr);
			rs.GetValue(_T("age"), tmpStr);				
			rptDetail->SetValue(_T("4"), tmpStr);

			rs.GetValue(_T("docno"), tmpStr);
			rptDetail->SetValue(_T("5"), tmpStr);

			rs.GetValue(_T("materialamt"), cost);	
			grpCost[5] += cost;
			FormatCurrency(cost, tmpStr);
			rptDetail->SetValue(_T("6"), tmpStr);

			rs.GetValue(_T("drugamt"), cost);	
			grpCost[6] += cost;
			FormatCurrency(cost, tmpStr);
			rptDetail->SetValue(_T("7"), tmpStr);

			rs.GetValue(_T("amount"), cost);	
			grpCost[7] += cost;
			FormatCurrency(cost, tmpStr);
			rptDetail->SetValue(_T("8"), tmpStr);

			rs.MoveNext();
		}
					
		
		if (grpCost[7] > 0)
		{	
			TranslateString(_T("Total Group"), szAmount);
			rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));	
			rptDetail->SetValue(_T("TotalGroup"), szAmount);
			for (int i = 5; i < 8; i++)
			{
				FormatCurrency(grpCost[i], tmpStr);
				tmpStr2.Format(_T("s%d"), i + 1);
				rptDetail->SetValue(tmpStr2, tmpStr);
				ttlCost[i] += grpCost[i];					
				grpCost[i] = 0;
			}			
		}
		if (ttlCost[7] > 0)
		{				
			TranslateString(_T("Total Amount"), szAmount);
			rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
			rptDetail->SetValue(_T("TotalAmount"), szAmount);
			for (int i = 5; i < 8; i++)
			{
				FormatCurrency(ttlCost[i], tmpStr);
				tmpStr2.Format(_T("s%d"), i + 1);
				rptDetail->SetValue(tmpStr2, tmpStr);
			}
		}
		CString szTime;
		szTime = pMF->GetSysDateTime();
		szSysDate = pMF->GetSysDate(); 
		szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
			          szTime.Right(8), szSysDate.Right(2), szSysDate.Mid(5,2), szSysDate.Left(4));
		rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
		if (bPreview)
			rpt.PrintPreview();
		else
			rpt.Print(true);
}

void CPrintForms::FM_PrintDetailedApprovedInOrder(long nDocno, long nTransactionID, bool bPreview){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CRecord rss(&pMF->m_db);

	static CReport rpt;

	CString szSQL;
	CString tmpStr, tmpStr1, szType;

	if (!rpt.Init(_T("Reports/HMS/PTTT/HF_PHIEULINHDV_CT.RPT")) )
	{
		//ShowMessageBox(_T("Can not open file report."), MB_OK);
		return;
	}
	szSQL.Format(_T("SELECT hcr_docno                                       docno, ") \
				_T("       hcr_recordno                                    recordno, ") \
				_T("       Get_departmentname(hcr_admitdept)               deptname, ") \
				_T("       Get_patientname(hcr_docno)                      patname, ") \
				_T("       Hms_getagebydoc(hcr_docno)                      age, ") \
				_T("       Get_syssel_desc('sys_sex', hp_sex) sex, ") \
				_T("	   hd_diagnostic diagno,") \
				_T("       Hms_getaddress(hp_provid, hp_distid, hp_villid) addr ") \
				_T("FROM   hms_clinical_record ") \
				_T("       LEFT JOIN hms_patient ") \
				_T("              ON ( hcr_patientno = hp_patientno ) ") \
				_T("	   LEFT JOIN hms_doc ") \
				_T("			  ON (hcr_docno = hd_docno)") \
				_T("WHERE  hcr_docno = %ld"), nDocno);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	rpt.GetReportHeader()->SetValue(_T("Department"), rs.GetValue(_T("deptname")));

	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), rs.GetValue(_T("docno")));

	rpt.GetReportHeader()->SetValue(_T("RecordNo"), rs.GetValue(_T("recordno")));

	rpt.GetReportHeader()->SetValue(_T("PATIENTNAME"), rs.GetValue(_T("patname")));

	rpt.GetReportHeader()->SetValue(_T("Age"), rs.GetValue(_T("age")));

	rpt.GetReportHeader()->SetValue(_T("Sex"), rs.GetValue(_T("sex")));

	rpt.GetReportHeader()->SetValue(_T("Address"), rs.GetValue(_T("addr")));

	rpt.GetReportHeader()->SetValue(_T("Diagnostic"), rs.GetValue(_T("diagno")));
	szSQL.Format(_T("SELECT mt_orderno orderno FROM m_transaction WHERE mt_transaction_id = %ld"), nTransactionID);
	rs.ExecSQL(szSQL);
	rpt.GetReportHeader()->SetValue(_T("OrderNo"), rs.GetValue(_T("orderno")));
	//Page Header
	//Report Detail
	szSQL.Format(_T("SELECT product_name                        pname, ") \
				_T("       product_uomname                     uom, ") \
				_T("       Sum(hpol_qtyorder)                  qty, ") \
				_T("       hpol_unitprice                      price, ") \
				_T("       Sum(hpol_unitprice * hpol_qtyorder) amount ") \
				_T("FROM   hms_ipharmaorder ") \
				_T("       LEFT JOIN hms_ipharmaorderline ") \
				_T("              ON ( hpo_orderid = hpol_orderid ) ") \
				_T("       LEFT JOIN m_product_view ") \
				_T("              ON ( product_id = hpol_product_id ) ") \
				_T("WHERE  hpo_transaction_id = %ld ") \
				_T("       AND hpol_docno = %ld ") \
				_T("GROUP  BY product_name, ") \
				_T("          product_uomname, ") \
				_T("          hpol_unitprice "), nTransactionID, nDocno);
	rs.ExecSQL(szSQL);
	CReportSection* rptDetail = NULL;

	int m_nIdx = 0;
	double Cost = 0;
	long double ttCost = 0, grpCost = 0;
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		m_nIdx++;
		tmpStr.Format(_T("%d"), m_nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("pname")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("uom")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("price")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("qty")));
		rs.GetValue(_T("amount"), tmpStr);
		Cost += str2double(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);	
		rs.MoveNext();
	}		
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
	rptDetail->SetValue(_T("s6"), double2str(Cost));
	//Page Footer
	//Report Footer
	CDate sysDate;
	sysDate.ParseDate(pMF->GetSysDate());
	rpt.GetReportFooter()->Format(_T("PrintDate"), sysDate.GetDay(), sysDate.GetMonth(), sysDate.GetYear());
	if (bPreview)
		rpt.PrintPreview();
	else 
		rpt.Print(true);
}

void CPrintForms::FM_PrintDetailedApprovedInOrderbyDrug(long nTransactionID){
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szDate, szSysDate, szSQL, tmpStr2, m_szFacList, szWhere;
	CString  szObject, m_szGroups, szObjectName, szType, szTypeName, szTableName;
	CString szStocks, szStockNames;
	int nCount = 0;

	if (!rpt.Init(_T("Reports/HMS/PTTT/PMS_BAOCAOXUATTHUOCCHOKHOATHEODOITUONG.RPT")))
	{
		//CString MsgReport;
		//MsgReport.Format(_T("Mau bao cao [%s] chua co. Lien he voi Adminstrator"),  _T("PMS_BAOCAOXUATTHUOCCHOKHOATHEODOITUONG.RPT"));
		//AfxMessageBox(MsgReport);
		return;
	}

	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	szSysDate = pMF->GetSysDateTime(); 
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm ),
	CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	rpt.GetReportHeader()->SetValue(_T("StockName"), szStockNames);
	rpt.GetReportHeader()->SetValue(_T("TableName"), szTableName);
	rpt.GetReportHeader()->SetValue(_T("DrugType"), szTypeName);
	rpt.GetReportHeader()->SetValue(_T("Object"), szObjectName);

	szSQL.Format(_T("SELECT product_id                          id, ") \
				_T("       product_name                        name, ") \
				_T("       product_uomname                     unit, ") \
				_T("       hpol_unitprice                      price, ") \
				_T("       Sum(hpol_qtyissue)                  AS qty, ") \
				_T("       Sum(hpol_unitprice * hpol_qtyissue) AS amount ") \
				_T("FROM   hms_ipharmaorder ") \
				_T("       LEFT JOIN hms_ipharmaorderline ") \
				_T("              ON ( hpo_orderid = hpol_orderid ) ") \
				_T("       LEFT JOIN m_product_view ") \
				_T("              ON ( product_id = hpol_product_id ) ") \
				_T("WHERE  hpo_transaction_id = %ld ") \
				_T("GROUP  BY product_id, ") \
				_T("          product_name, ") \
				_T("          product_uomname, ") \
				_T("          hpol_unitprice ") \
				_T("ORDER  BY product_name "), nTransactionID);

			
			CReportSection* rptDetail;
			CString szOldLine, szNewLine, szAmount;
			CRecord rs(&pMF->m_db);
			long double grpCost=0.0;
			long double ttlCost=0.0;
			double cost = 0;
			int nItem = 1;
			_fmsg(_T("%s"), szSQL);
			rs.ExecSQL(szSQL);

			if(rs.IsEOF())
			{
				ShowMessageBox(_T("No Data"), MB_OK);
				return ;
			}

					
			while(!rs.IsEOF())
			{
				rptDetail = rpt.AddDetail();
				tmpStr.Format(_T("%d"), nItem++);
				rptDetail->SetValue(_T("1"), tmpStr);
				rs.GetValue(_T("id"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rs.GetValue(_T("name"), tmpStr);
				rptDetail->SetValue(_T("3"), tmpStr);
				rs.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				rs.GetValue(_T("price"), cost);
				FormatCurrency(cost, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);

				rs.GetValue(_T("qty"), tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);
				rs.GetValue(_T("amount"), cost);	
				grpCost += cost;
				FormatCurrency(cost, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				rs.MoveNext();
			}
						
			ttlCost += grpCost;
			
			if (grpCost > 0)
			{	
				TranslateString(_T("Total Group"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));	
				rptDetail->SetValue(_T("TotalGroup"), szAmount);
				FormatCurrency(grpCost, tmpStr);
				rptDetail->SetValue(_T("s7"), tmpStr);			
			}
			if (ttlCost > 0)
			{				
				TranslateString(_T("Total Amount"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("TotalGroup"), szAmount);
				FormatCurrency(ttlCost, tmpStr);
				rptDetail->SetValue(_T("s7"), tmpStr);
				ttlCost += grpCost;	
			}
			
			CString szTime;
			szTime = pMF->GetSysDateTime();
			szSysDate = pMF->GetSysDate(); 
			szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szTime.Right(8),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
			rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
			//	rpt.GetReportFooter()->SetValue(_T("username"), tmpStr);
		rpt.PrintPreview();			
}

void CPrintForms::TM_PrintInSurgeryPayment(long nSOrderID, long nDocumentNo, CString szDeptID){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	CRecord rss(&pMF->m_db);

	static CReport rpt;
	CReportSection *grp;
	CString tmpStr, tmpStrIdx, szSQL, szGroups, szReportName, szObjectKey;

	CString lpszChapter[] = {_T("I"), _T("II"), _T("III"), _T("IV"), _T("V"),
		                    _T("VI"), _T("VII"), _T("VIII"), _T("IX"), _T("X"), _T("XI")};
	int nChapter = 0;

	szSQL.Format(_T(" SELECT    hd_patientno                    AS patientno, ") \
				_T("           hcr_recordno                    AS recordno, ") \
				_T("           Trim(hp_surname ") \
				_T("                ||' ' ") \
				_T("                ||hp_midname ") \
				_T("                ||' ' ") \
				_T("                ||hp_firstname)            AS pname, ") \
				_T("           Extract(YEAR FROM hp_birthdate) AS birthdate, ") \
				_T("           hp_birthdate                    AS birthdate1, ") \
				_T("           get_syssel_desc('sys_sex', hp_sex)                          AS sex, ") \
				_T("           sp_name                         AS provname, ") \
				_T("           sys_dist.sd_name                AS distname, ") \
				_T("           sv_name                         AS villname, ") \
				_T("           hp_dtladdr                      AS address, ") \
				_T("           hp_workplace                    AS workplace, ") \
				_T("           hp_rank                         AS rank, ") \
				_T("           hd_object                       AS obj, ") \
				_T("		   get_objectname(hd_object)	   AS objname,") \
				_T("           hd_doctor                       AS doctor, ") \
				_T("           hd_indept                       AS deptid, ") \
				_T("		   get_departmentname(hd_indept)   AS deptname,") \
				_T("           hc_cardno                       AS cardno, ") \
				_T("           hd_disrate                      AS insrate, ") \
				_T("           hd_isinpatient                  AS inpatient, ") \
				_T("           hcr_admitdate                   AS admitdate, ") \
				_T("           hd_diagnostic                   AS diagnostic ") \
				_T(" FROM      hms_patient ") \
				_T(" LEFT JOIN hms_doc ON( hd_patientno = hp_patientno ) ") \
				_T(" LEFT JOIN hms_clinical_record ON( hcr_docno = hd_docno ) ") \
				_T(" LEFT JOIN hms_card ON( hc_patientno = hp_patientno ") \
				_T("                        AND hc_idx = hd_cardidx ") \
				_T("                        AND hc_cardno = hd_cardno ) ") \
				_T(" LEFT JOIN sys_dept ON( sys_dept.sd_id = hd_indept ) ") \
				_T(" LEFT JOIN sys_prov ON( sp_id = hp_provid ) ") \
				_T(" LEFT JOIN sys_dist ON( sd_provid = hp_provid ") \
				_T("                        AND sys_dist.sd_id = hp_distid ) ") \
				_T(" LEFT JOIN sys_vill ON( sv_id = hp_villid ") \
				_T("                        AND sv_distid = hp_distid ") \
				_T("                        AND sv_id = hp_villid ) ") \
				_T(" WHERE     hd_docno = %ld "), nDocumentNo);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;
	rs.GetValue(_T("obj"), szObjectKey);
	if (szObjectKey == _T("1") || szObjectKey ==_T("7") || szObjectKey ==_T("8")
		/*|| m_bExpireFlag*/)	
	{
		szReportName =_T("Reports/HMS/PTTT/HF_BANGKEPTTT_QUAN_DICHVU.RPT");
	}
	else
	{
		szReportName =_T("Reports/HMS/PTTT/HF_INSDISCHARGEPAYMENT.RPT");
	}
		

	if (!rpt.Init(szReportName))
	{
		//ShowMessageBox(_T("Can not open file report."), MB_OK);
		return;
	}
	CString szDate;	
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")),
		          tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), szDate);

	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	
	
	tmpStr.Format(_T("%ld"), nDocumentNo);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);

	//tmpStr.Format(_T("%ld"), m_nRecordNo);
	tmpStr.Format(_T("%s"), rs.GetValue(_T("recordno")));

	rpt.GetReportHeader()->SetValue(_T("RecordNo"), tmpStr);
	StringUpper(rs.GetValue(_T("pname")), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("PATIENTNAME"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), rs.GetValue(_T("birthdate")));
	rpt.GetReportHeader()->SetValue(_T("Sex"), rs.GetValue(_T("sex")));
	tmpStr.Empty();
	tmpStr.AppendFormat(_T("%s"), rs.GetValue(_T("villname")));
	if(!tmpStr.IsEmpty()) 
		tmpStr += _T(" - ");
	tmpStr.AppendFormat(_T("%s"), rs.GetValue(_T("distname")));
	if(!tmpStr.IsEmpty()) 
		tmpStr += _T(" - ");
	tmpStr.AppendFormat(_T("%s"), rs.GetValue(_T("provname")));
	
	rpt.GetReportHeader()->SetValue(_T("Address"), tmpStr);

	rpt.GetReportHeader()->SetValue(_T("Object"), rs.GetValue(_T("objname")));

	rpt.GetReportHeader()->SetValue(_T("CardNo"), rs.GetValue(_T("cardno")));

	rpt.GetReportHeader()->SetValue(_T("Department"), rs.GetValue(_T("deptname")));
	//Page Header
	//Report Detail

	int nIdx = 1, nIdxl = 1, nNhom = 0;
	int nNewIndex = 0, nOldIndex = 0;
	long nOrderID;
	double cost = 0;
	CReportSection* rptDetail = NULL;
	double optcost=0, inscost=0, patpaid=0;	
	double grpoptcost=0, grpinscost=0, grppatpaid=0, grpcost=0;	
	double grp2optcost=0, grp2inscost=0, grp2patpaid=0, grp2cost=0;	
	double ttloptcost=0, ttlinscost=0, ttlpatpaid=0, ttlcost=0, ttSurgeryCost = 0;	
	double unitprice=0, qty=0;
	long nDocNo;
	CString szOldLine, szNewLine, szOldGroup, szNewGroup, szLineName, szLineNameOld;
    CString szComment, szTotalComment;
	
	szGroups.Empty();
	CString szTmp;

	szSQL.Format(_T("SELECT ho_idx, ho_docno, hfl_name, ho_comment, hfl_unit, hfl_servprice, ") \
				 _T(" ho_diagnostic diagnostic, (select hst_name from hms_surgery_table where hst_idx = ho_opera_table) table_name") \
				 _T(" FROM hms_operation ") \
				 _T(" LEFT JOIN hms_fee_list ON(hfl_feeid = ho_itemid)") \
				 _T(" WHERE ho_idx = %ld AND ho_status in('A','T','S')") \
				 _T(" ORDER BY ho_idx"), nSOrderID);


	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;

	rpt.GetReportHeader()->SetValue(_T("room"), rs.GetValue(_T("table_name")));
	rpt.GetReportHeader()->SetValue(_T("Diagnostic"), rs.GetValue(_T("diagnostic")));
	szComment.Empty();
	szGroups.Empty();

	grp = rpt.GetGroupHeader(1);
	rptDetail = rpt.AddDetail(grp);
	tmpStr.Format(_T("%s%s"), lpszChapter[nChapter], _T("."));
	rptDetail->SetValue(_T("SubGroupID"), tmpStr);		
	rs.GetValue(_T("ho_comment"), szComment);
	
	tmpStr = rs.GetValue(_T("hfl_name"));
	rptDetail->SetValue(_T("SubGroupName"), tmpStr);
	rs.GetValue(_T("hfl_servprice"), cost);
	ttSurgeryCost += cost;
	FormatCurrency(cost, tmpStr);
	
	rptDetail->SetValue(_T("SubGroupCost"), tmpStr);

	if (!szComment.IsEmpty())
	{
		rptDetail = rpt.AddDetail(grp);
		tmpStr.Format(_T("* %s"), szComment);
		rptDetail->SetValue(_T("SubGroupName"), tmpStr);
	}

	rs.GetValue(_T("ho_idx"), nOrderID);
	rs.GetValue(_T("ho_docno"), nDocNo);
	
	szSQL.Format(_T(" SELECT hsd_feeidx AS grpidx,") \
				_T("   hsd_name        AS grpname,") \
				_T("   hfe_desc,") \
				_T("   hfe_unit,") \
				_T("   hpol_unitprice as hfe_unitprice,") \
				_T("   hfe_category,") \
				_T("   SUM(hfe_quantity) AS qty,") \
				_T("   CASE") \
				_T("     WHEN hfe_category='Y'") \
				_T("     THEN SUM(hfe_cost)") \
				_T("     ELSE 0") \
				_T("   END AS optcost,") \
				_T("   CASE") \
				_T("     WHEN hfe_category='N'") \
				_T("     THEN SUM(hfe_inspaid)") \
				_T("     ELSE 0") \
				_T("   END                              AS inscost,") \
				_T("   SUM(hfe_diffcost)                AS diffcost,") \
				_T("   SUM(hfe_quantity*hpol_unitprice) AS cost") \
				_T(" FROM hms_fee") \
				_T(" LEFT JOIN hms_surgery_drugtype") \
				_T(" ON(hsd_id =hfe_group)") \
				_T(" LEFT JOIN hms_ipharmaorder") \
				_T(" ON(hfe_docno    = hpo_docno") \
				_T(" AND hfe_orderid = hpo_orderid)") \
				_T(" LEFT JOIN hms_ipharmaorderline") \
				_T(" ON(hpol_orderid      = hpo_orderid") \
				_T(" AND hpol_orderlineid = hfe_orderline)") \
				_T(" WHERE hfe_docno      =%ld") \
				_T(" AND hpo_reference_id =%ld") \
				_T(" AND hfe_type        IN('D','P','M')") \
				_T(" GROUP BY hsd_feeidx,") \
				_T("   hsd_name,") \
				_T("   hfe_desc,") \
				_T("   hfe_unit,") \
				_T("   hfe_unitprice,") \
				_T("   hfe_category,") \
				_T("   hpol_unitprice") \
				_T(" ORDER BY hsd_feeidx,") \
				_T("   hfe_desc"), nDocumentNo, nOrderID);
	int nNewFeeIdx=0, nOldFeeIdx=0;
	
	rsl.ExecSQL(szSQL);
	if(rsl.IsEOF())
		return;
	while(!rsl.IsEOF())
	{

		rsl.GetValue(_T("grpidx"), nNewFeeIdx);
		if(nNewFeeIdx != nOldFeeIdx)
		{
			if(grp2cost > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(3));
				tmpStr.Format(_T("\x43\x1ED9ng(%d):"), nIdx);
				rptDetail->SetValue(_T("TotalAmountLabel"), tmpStr);
				FormatCurrency(grp2optcost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);					
				FormatCurrency(grp2inscost, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);					
				FormatCurrency(grp2patpaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);
				FormatCurrency(grp2cost, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				grp2cost = grp2inscost=grp2optcost=grp2patpaid = 0;
			}

			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(2));
			tmpStr.Format(_T("%d"), ++nIdx);
			rptDetail->SetValue(_T("TypeID"), tmpStr);
			if(nNewFeeIdx == 1)
				tmpStr.Format(_T("Thu\x1ED1\x63, h\xF3\x61 \x63h\x1EA5t, m\xE1u, \x64\x1ECB\x63h truy\x1EC1n"));
			else if(nNewFeeIdx == 2)
				tmpStr.Format(_T("V\x1EADt t\x1B0 y t\x1EBF ti\xEAu h\x61o"));
			else
			{
				rsl.GetValue(_T("grpname"), tmpStr);
			}
			rptDetail->SetValue(_T("TypeName"), tmpStr);			
			
			nIdxl = 1;
			nOldFeeIdx = nNewFeeIdx;
		}
		rptDetail = rpt.AddDetail();

		tmpStrIdx.Format(_T("%d)   "), nIdxl++);
		rsl.GetValue(_T("hfe_desc"), tmpStr);
		rptDetail->SetValue(_T("1"), _T(""));
		rptDetail->SetValue(_T("2"), tmpStrIdx + tmpStr);

		rsl.GetValue(_T("hfe_unit"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rsl.GetValue(_T("qty"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rsl.GetValue(_T("hfe_unitprice"), unitprice);
		FormatCurrency(unitprice, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rsl.GetValue(_T("optcost"), optcost);
		grp2optcost += optcost;
		grpoptcost += optcost;
		ttloptcost += optcost;
		FormatCurrency(optcost, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rsl.GetValue(_T("inscost"), inscost);
		grp2inscost += inscost;
		grpinscost += inscost;
		ttlinscost += inscost;
		FormatCurrency(inscost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rsl.GetValue(_T("patpaid"), patpaid);
		grp2patpaid += patpaid;
		grppatpaid += patpaid;
		ttlpatpaid += patpaid;
		FormatCurrency(patpaid, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);


		rsl.GetValue(_T("cost"),cost);			
		grp2cost += cost;
		grpcost += cost;
		ttlcost += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rsl.MoveNext();
	}


	if(grp2cost > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupHeader(3));
		tmpStr.Format(_T("\x43\x1ED9ng(%d):"), nIdx);
		rptDetail->SetValue(_T("TotalAmountLabel"), tmpStr);
		FormatCurrency(grp2optcost, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);					
		FormatCurrency(grp2inscost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);					
		FormatCurrency(grp2patpaid, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		FormatCurrency(grp2cost, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);
		grp2cost = grp2inscost=grp2optcost=grp2patpaid = 0;
	}

	rptDetail = rpt.AddDetail(rpt.GetGroupHeader(3));
	for (int i =0; i < nIdx; i++){
		if(!szTotalComment.IsEmpty())
			szTotalComment.AppendFormat(_T("+"));
		szTotalComment.AppendFormat(_T("%d"), i+1);
	}
	tmpStr.Format(_T("T\x1ED5ng (%s)=(%s):"),lpszChapter[nChapter],szTotalComment);
	szTotalComment.Empty();
	rptDetail->SetValue(_T("TotalAmountLabel"), tmpStr);
	
	FormatCurrency(grpoptcost, tmpStr);
	rptDetail->SetValue(_T("6"), tmpStr);					
	FormatCurrency(grpinscost, tmpStr);
	rptDetail->SetValue(_T("7"), tmpStr);					
	FormatCurrency(grppatpaid, tmpStr);
	rptDetail->SetValue(_T("8"), tmpStr);
	FormatCurrency(grpcost, tmpStr);
	rptDetail->SetValue(_T("9"), tmpStr);
	
	
	CString szSumInWord;
	szSumInWord.Format(_T("\x110\x1ED3ng"));
	FormatCurrency(ttlcost, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	FormatCurrency(ttlinscost, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("InsurancePaid"), tmpStr );
	FormatCurrency(ttlpatpaid, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("PatientPaid"), tmpStr);

	//Page Footer
	//Report Footer
	CDate sysDate;
	sysDate.ParseDate(pMF->GetSysDate());
	rpt.GetReportFooter()->Format(_T("PrintDate"), sysDate.GetDay(), sysDate.GetMonth(), sysDate.GetYear());
	//rpt.GetReportFooter()->SetValue(_T("DoctorName"), tmpStr);
	rpt.PrintPreview();	
	
}



void CPrintForms::PrintOperationHitechMaterial(long nDocumentNo, long nOperationIdx)
{
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	static CReport rpt;
	CString tmpStr, szSQL;
	if(!rpt.Init(_T("Reports/HMS/HF_OPERATIONMATERIAL_HITECH.RPT")) )
	{
		//ShowMessageBox(_T("Can not open file report."), MB_OK);
		return;
	}


		//Page Header
	//Report Detail
	CReportSection* rptDetail = NULL, *rptHeader = NULL, *rptHead = NULL;
	int nIdx = 1;
	double nTtCost = 0, nGrpCost = 0, nGrpCostAB = 0, nCost = 0, nOutInsCost = 0;
	double nTotalInsLimit = 0, nTotalDiscount = 0, nTotalPatpaid = 0;
	bool bInsOnly = true;
	float nAmount = 0;

	CRecord rsl(&pMF->m_db);
	CString szOldType=_T(""), szNewType, szSysTime = pMF->GetSysTime(), szSysDate = pMF->GetSysDate();
	//Patient Info query
	szSQL.Format(_T(" SELECT hcr_recordno AS recordno,") \
_T("   hd_docno docno,") \
_T("   Trim(hp_surname") \
_T("   ||' '") \
_T("   ||hp_midname") \
_T("   ||' '") \
_T("   ||hp_firstname)                                         AS pname,") \
_T("   Extract(YEAR FROM hp_birthdate)                         AS birthdate,") \
_T("   hms_getage(hd_admitdate, hp_birthdate)                  AS age,") \
_T("   Get_syssel_desc('sys_sex', hp_sex)                      AS sex,") \
_T("   Hms_getaddress(hp_provid, hp_distid, hp_villid)         AS address,") \
_T("   get_objectname(hd_object)                               AS obj,") \
_T("   get_username(ho_practitioner)                                 AS doctor,") \
_T("   hc_cardno                                               AS cardno,") \
_T("   hd_disrate                                              AS insrate,") \
_T("   hd_diagnostic                                           AS diagnostic,") \
_T("   Get_departmentname(ho_deptid)                           AS dept,") \
_T("   ho_performdate                                          AS performdate,") \
_T("   hfe_desc                                                AS optname ,") \
_T("   hfe_quantity                                            AS qty,") \
_T("   (hfe_unitprice-hfe_diffcost)/hfe_quantity                 AS insprice,") \
_T("   (hms_fee.hfe_cost-hfe_diffcost) AS amount,") \
_T("   hms_fee.hfe_discount as discount, ") \
_T("   hfe_diffcost as  diffamount,") \
_T(" (hfe_patdebt+hfe_patpaid) as patpaid, ") \
_T("   ho_maxinspaid") \
_T(" FROM hms_patient") \
_T(" LEFT JOIN hms_doc") \
_T(" ON( hd_patientno = hp_patientno )") \
_T(" LEFT JOIN hms_clinical_record") \
_T(" ON( hcr_docno = hd_docno )") \
_T(" LEFT JOIN hms_card") \
_T(" ON( hc_patientno = hp_patientno") \
_T(" AND hc_idx       = hd_cardidx") \
_T(" AND hc_cardno    = hd_cardno )") \
_T(" LEFT JOIN hms_operation") \
_T(" ON ( hd_docno = ho_docno)") \
_T(" LEFT JOIN hms_fee") \
_T(" ON (hfe_docno    = ho_docno") \
_T(" AND hfe_orderid  = ho_idx)") \
_T(" WHERE hd_docno   = %ld") \
_T(" AND ho_idx       = %ld") \
_T(" AND hfe_type     ='O'") \
_T(" AND hfe_quantity > 0"), nDocumentNo, nOperationIdx);

	
	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}	

	rs.GetValue(_T("ho_maxinspaid"), nTotalInsLimit);

	//Header Data
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
	rptHead = rpt.GetReportHeader();
	rptHead->SetValue(_T("RecordNo"), rs.GetValue(_T("recordno")));
	rptHead->SetValue(_T("DocumentNo"), rs.GetValue(_T("docno")));
	tmpStr.Format(rptHead->GetValue(_T("PrintDate")), szSysTime, szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rptHead->SetValue(_T("PrintDate"), tmpStr);
	rptHead->SetValue(_T("PatientName"), rs.GetValue(_T("pname")));
	rptHead->SetValue(_T("Age"), rs.GetValue(_T("age")));
	rptHead->SetValue(_T("Sex"), rs.GetValue(_T("sex")));
	rptHead->SetValue(_T("address"), rs.GetValue(_T("address")));
	rptHead->SetValue(_T("cardno"), rs.GetValue(_T("cardno")));
	rptHead->SetValue(_T("Object"), rs.GetValue(_T("obj")));
	rptHead->SetValue(_T("Department"), rs.GetValue(_T("dept")));
	rptHead->SetValue(_T("%"), rs.GetValue(_T("insrate")));
	rptHead->SetValue(_T("perfomdate"), CDate::Convert(rs.GetValue(_T("performdate")), yyyymmdd, ddmmyyyy));
	rptHead->SetValue(_T("doctor"), rs.GetValue(_T("doctor")));
	rptHead->SetValue(_T("diagnostic"), rs.GetValue(_T("diagnostic")));

	rs.GetValue(_T("diffamount"), nOutInsCost);
	
	rptHeader = rpt.AddDetail(rpt.GetGroupHeader(1));
	rptHeader->SetValue(_T("SubGroupID"), _T("A"));
	rptHeader->SetValue(_T("SubGroupName"), _T("\x43hi ph\xED trong quy \x111\x1ECBnh \x63\x1EE7\x61 \x42H\x58H"));
	int nIndex = 1;
	double opt_discount = 0;
	double opt_patpaid = 0;
	rs.GetValue(_T("discount"), opt_discount);
	rs.GetValue(_T("patpaid"), opt_patpaid);


	rptHeader = rpt.AddDetail(rpt.GetGroupHeader(2));	
	rptHeader->SetValue(_T("TypeID"), _T("I"));
	rptHeader->SetValue(_T("TypeName"), _T("\x44\x1ECB\x63h v\x1EE5 k\x1EF9 thu\x1EADt \x63\x61o (\x44VKT\x43)"));

	if (opt_discount > 0)
	{
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("optname")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("unit")));
		rptDetail->SetValue(_T("4"), _T(""));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("qty")));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("qty")));
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("insprice")));
		rs.GetValue(_T("amount"), nCost);
		rptDetail->SetValue(_T("8"), double2str(nCost));

		nGrpCostAB += nCost;

		if (nCost > 0)
		{
			rptHeader = rpt.AddDetail(rpt.GetGroupHeader(3));
			rptHeader->SetValue(_T("TotalAmount"), _T("\x43\x1ED9ng I"));
			rptHeader->SetValue(_T("8"), double2str(nCost));
			
			nCost = 0;
		}
	}
	else
	{
		
	}

	rs.GetValue(_T("discount"), nAmount);
	nTotalDiscount += nAmount;
	rs.GetValue(_T("patpaid"), nAmount);
	nTotalPatpaid += nAmount;

	//Material Info query
	szSQL.Format(_T(" SELECT product_type,") \
_T("   product_name,") \
_T("   product_uomname,") \
_T("   manufacture,") \
_T("   idx,") \
_T("   unitprice,") \
_T("   SUM(qty)      AS qty,") \
_T("   SUM(cost)     AS amount,") \
_T(" SUM(inspaid) as inspaid, ") \
_T("   SUM(discount) AS discount, ") \
_T(" SUM(diffpaid) AS diffpaid ") \
_T(" FROM") \
_T("   (SELECT mp_producttype AS product_type,") \
_T("     mp_name              AS product_name,") \
_T("     adu_name             AS product_uomname,") \
_T("     get_manufacturename(mp_manufacture_id) manufacture,") \
_T("     CASE") \
_T("       WHEN hpol_inspaid = 'Y'") \
_T("       THEN 0") \
_T("       ELSE 1") \
_T("     END                                     AS idx,") \
_T("     hfe_insprice AS unitprice,") \
_T("     hfe_quantity                            AS qty,") \
_T("     hfe_cost                   AS cost,") \
_T("     hfe_inspaid                   AS inspaid,") \
_T("     hfe_discount                            AS discount, ") \
_T("     hfe_diffcost                            AS diffpaid ") \
_T("   FROM hms_ipharmaorder") \
_T("   LEFT JOIN hms_ipharmaorderline") \
_T("   ON(hpol_orderid=hpo_orderid)") \
_T("   LEFT JOIN hms_fee") \
_T("   ON(hfe_docno      = hpo_docno") \
_T("   AND hfe_orderid   = hpo_orderid") \
_T("   AND hfe_orderline = hpol_orderlineid") \
_T("   AND hfe_type     IN('D','M') )") \
_T("   LEFT JOIN m_product") \
_T("   ON(mp_product_id=hpol_product_id)") \
_T("   LEFT JOIN ad_uom") \
_T("   ON(adu_uom_id        =mp_product_uom_id)") \
_T("   WHERE hpo_docno      =%ld") \
_T("   AND hpo_reference_id =%ld") \
_T("   AND hpo_ordertype   IN('D','M')") \
_T("   AND mp_org_id     IN('MA','BB') ") \
_T("   AND hpol_inoperation<>'Y' AND hfe_discount > 0  and hpol_hitech='Y' ") \
_T("   )") \
_T(" GROUP BY product_type,") \
_T("   product_uomname,") \
_T("   product_name,") \
_T("   manufacture,") \
_T("   unitprice,") \
_T("   idx") \
_T(" ORDER BY idx, product_name, manufacture, unitprice") , nDocumentNo, nOperationIdx);


	rsl.ExecSQL(szSQL);
	if (!rsl.IsEOF())
	{
	}

	rptHeader = rpt.AddDetail(rpt.GetGroupHeader(2));
	rptHeader->SetValue(_T("TypeID"), _T("II"));
	rptHeader->SetValue(_T("TypeName"), _T("V\x1EADt t\x1EF1 ti\xEAu h\x61o \x111\x1EB7\x63 \x62i\x1EC7t"));

	nIdx = 1;
	double nDiffPaid = 0;
	double nAmt = 0;
	double nInsPaid = 0;
	double nTotalInsPaid = 0;
	

	bool bApplyMaterialLimit = false;

	nTotalInsPaid += nGrpCostAB;

	while(!rsl.IsEOF())
	{
		rsl.GetValue(_T("idx"), szNewType);
		if (szNewType != szOldType)
		{
			
			nIdx = 1;
			szOldType = szNewType;		
		}
		rptDetail = rpt.AddDetail();	
		tmpStr.Format(_T("%d"), nIdx++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rsl.GetValue(_T("product_name"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rsl.GetValue(_T("product_uomname"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rsl.GetValue(_T("manufacture"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rsl.GetValue(_T("qty"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rsl.GetValue(_T("unitprice"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rsl.GetValue(_T("amount"), nCost);
		rsl.GetValue(_T("inspaid"), nInsPaid);
		rsl.GetValue(_T("diffpaid"), nAmt);
		
		nTotalInsPaid += nInsPaid;

	//	nDiffPaid += nAmt;
		nGrpCost += nCost;
		tmpStr.Format(_T("%f"), nCost);
		rptDetail->SetValue(_T("8"), tmpStr);
		if(szNewType == _T("0"))
		{
			rsl.GetValue(_T("discount"), nAmount);
			nTotalDiscount += nAmount;
		}
		if (nInsPaid > 0 && nInsPaid != nCost)
		{
			rptDetail = rpt.AddDetail();
			rptDetail->SetValue(_T("1"), _T(""));
			//rptDetail->SetValue(_T("2"), _T(""));
			rptDetail->SetValue(_T("3"), _T(""));
			rptDetail->SetValue(_T("4"), _T(""));
			rptDetail->SetValue(_T("5"), _T(""));
			rptDetail->SetValue(_T("6"), _T(""));
			rptDetail->SetValue(_T("7"), _T(""));
			rptDetail->SetValue(_T("2"), _T("\x110\x1ECBnh m\x1EE9\x63 \x42H\x58H"));
			tmpStr.Format(_T("%.2f"), nInsPaid);
			rptDetail->SetValue(_T("8"), tmpStr);
			
			bApplyMaterialLimit = true;

		}

		rsl.MoveNext();
	}

	if(nGrpCost > 0)
	{
		rptHeader = rpt.AddDetail(rpt.GetGroupHeader(3));
		rptHeader->SetValue(_T("TotalAmount"), _T("\x43\x1ED9ng II"));
		tmpStr.Format(_T("%f"), nGrpCost);
		rptHeader->SetValue(_T("8"), tmpStr);
		nGrpCostAB += nGrpCost;
	}

	nTtCost += nGrpCostAB;
	nGrpCostAB = 0;
//	if(nTtCost > 0)
	{
		rptHeader = rpt.AddDetail(rpt.GetGroupHeader(3));
		rptHeader->SetValue(_T("TotalAmount"), _T("\x43\x1ED9ng A(I+II)"));
		tmpStr.Format(_T("%f"), nTtCost);
		nGrpCostAB = 0;
		rptHeader->SetValue(_T("8"), tmpStr);

	}



	//B:NGOAI QUY DINH BHXH
	rptHeader = rpt.AddDetail(rpt.GetGroupHeader(1));
	rptHeader->SetValue(_T("SubGroupID"), _T("B"));
	rptHeader->SetValue(_T("SubGroupName"), _T("\x43hi ph\xED ngo\xE0i quy \x111\x1ECBnh \x42H\x58H"));

	nGrpCostAB = 0;
	if(opt_discount <= 0 && opt_patpaid > 0)
	{
		rptHeader = rpt.AddDetail(rpt.GetGroupHeader(2));	
		rptHeader->SetValue(_T("TypeID"), _T("I"));
		rptHeader->SetValue(_T("TypeName"), _T("\x44\x1ECB\x63h v\x1EE5 k\x1EF9 thu\x1EADt \x63\x61o (\x44VKT\x43)"));

		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("optname")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("unit")));
		rptDetail->SetValue(_T("4"), _T(""));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("qty")));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("qty")));
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("insprice")));
		rs.GetValue(_T("amount"), nCost);
		rptDetail->SetValue(_T("8"), double2str(nCost));
		if (nCost > 0)
		{
			rptHeader = rpt.AddDetail(rpt.GetGroupHeader(3));
			rptHeader->SetValue(_T("TotalAmount"), _T("\x43\x1ED9ng I"));
			rptHeader->SetValue(_T("8"), double2str(nCost));
			nGrpCostAB += nCost;
			nCost = 0;
		}

		

	}
	else if (nOutInsCost > 0)
	{

		nOutInsCost += nDiffPaid;
		rptHeader = rpt.AddDetail(rpt.GetGroupHeader(3));
		rptHeader->SetValue(_T("TotalID"), _T("I"));
		rptHeader->SetValue(_T("TotalAmount"), _T("\x43hi ph\xED kh\xE1\x63 ngo\xE0i quy \x111\x1ECBnh \x42H\x58H"));
		tmpStr.Format(_T("%.2f"), ceil(nOutInsCost));
		rptHeader->SetValue(_T("8"), tmpStr);

		nGrpCostAB += nOutInsCost;
	}

	//Material Info query
	szSQL.Format(_T(" SELECT product_type,") \
_T("   product_name,") \
_T("   product_uomname,") \
_T("   manufacture,") \
_T("   idx,") \
_T("   unitprice,") \
_T("   SUM(qty)      AS qty,") \
_T("   SUM(cost)     AS amount,") \
_T(" SUM(inspaid) as inspaid, ") \
_T("   SUM(discount) AS discount, ") \
_T(" SUM(diffpaid) AS diffpaid ") \
_T(" FROM") \
_T("   (SELECT mp_producttype AS product_type,") \
_T("     mp_name              AS product_name,") \
_T("     adu_name             AS product_uomname,") \
_T("     get_manufacturename(mp_manufacture_id) manufacture,") \
_T("     CASE") \
_T("       WHEN hpol_inspaid = 'Y'") \
_T("       THEN 0") \
_T("       ELSE 1") \
_T("     END                                     AS idx,") \
_T("     hfe_unitprice AS unitprice,") \
_T("     hfe_quantity                            AS qty,") \
_T("     hfe_cost                   AS cost,") \
_T("     hfe_inspaid                   AS inspaid,") \
_T("     hfe_discount                            AS discount, ") \
_T("     hfe_diffcost                            AS diffpaid ") \
_T("   FROM hms_ipharmaorder") \
_T("   LEFT JOIN hms_ipharmaorderline") \
_T("   ON(hpol_orderid=hpo_orderid)") \
_T("   LEFT JOIN hms_fee") \
_T("   ON(hfe_docno      = hpo_docno") \
_T("   AND hfe_orderid   = hpo_orderid") \
_T("   AND hfe_orderline = hpol_orderlineid") \
_T("   AND hfe_type     IN('D','M') )") \
_T("   LEFT JOIN m_product") \
_T("   ON(mp_product_id=hpol_product_id)") \
_T("   LEFT JOIN ad_uom") \
_T("   ON(adu_uom_id        =mp_product_uom_id)") \
_T("   WHERE hpo_docno      =%ld") \
_T("   AND hpo_reference_id =%ld") \
_T("   AND hpo_ordertype   IN('D','M')") \
_T("   AND mp_org_id     IN('MA','BB') ") \
_T("   AND hpol_inoperation<>'Y' AND hfe_discount = 0 and hpol_hitech='Y' ") \
_T("   )") \
_T(" GROUP BY product_type,") \
_T("   product_uomname,") \
_T("   product_name,") \
_T("   manufacture,") \
_T("   unitprice,") \
_T("   idx") \
_T(" ORDER BY idx, product_name, manufacture, unitprice") , nDocumentNo, nOperationIdx);

//_msg(_T("%s"), szSQL);
	rsl.ExecSQL(szSQL);
	if (!rsl.IsEOF())
	{
	}
	
	rptHeader = rpt.AddDetail(rpt.GetGroupHeader(2));
	rptHeader->SetValue(_T("TypeID"), _T("II"));
	rptHeader->SetValue(_T("TypeName"), _T("V\x1EADt t\x1EF1 ti\xEAu h\x61o \x111\x1EB7\x63 \x62i\x1EC7t"));


	nIdx = 1;
	nDiffPaid = 0;
	nAmt = 0;
	nInsPaid = 0;
	
	

	bApplyMaterialLimit = false;

	

	nGrpCost = 0;
	while(!rsl.IsEOF())
	{
		rsl.GetValue(_T("idx"), szNewType);
		if (szNewType != szOldType)
		{
			
			nIdx = 1;
			szOldType = szNewType;		
		}
		rptDetail = rpt.AddDetail();	
		tmpStr.Format(_T("%d"), nIdx++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rsl.GetValue(_T("product_name"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rsl.GetValue(_T("product_uomname"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rsl.GetValue(_T("manufacture"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rsl.GetValue(_T("qty"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rsl.GetValue(_T("unitprice"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rsl.GetValue(_T("amount"), nCost);
		rsl.GetValue(_T("inspaid"), nInsPaid);
		rsl.GetValue(_T("diffpaid"), nAmt);
		
		

	//	nDiffPaid += nAmt;
		nGrpCost += nCost;
		tmpStr.Format(_T("%f"), nCost);
		rptDetail->SetValue(_T("8"), tmpStr);
		if(szNewType == _T("0"))
		{
			rsl.GetValue(_T("discount"), nAmount);
			nTotalDiscount += nAmount;
		}
		if (nInsPaid > 0 && nInsPaid != nCost)
		{
			rptDetail = rpt.AddDetail();
			rptDetail->SetValue(_T("1"), _T(""));
			//rptDetail->SetValue(_T("2"), _T(""));
			rptDetail->SetValue(_T("3"), _T(""));
			rptDetail->SetValue(_T("4"), _T(""));
			rptDetail->SetValue(_T("5"), _T(""));
			rptDetail->SetValue(_T("6"), _T(""));
			rptDetail->SetValue(_T("7"), _T(""));
			rptDetail->SetValue(_T("2"), _T("\x110\x1ECBnh m\x1EE9\x63 \x42H\x58H"));
			tmpStr.Format(_T("%.2f"), nInsPaid);
			rptDetail->SetValue(_T("8"), tmpStr);
			
			bApplyMaterialLimit = true;

		}

		rsl.MoveNext();
	}




	if (nGrpCost > 0)
	{
		rptHeader = rpt.AddDetail(rpt.GetGroupHeader(3));
		rptHeader->SetValue(_T("TotalAmount"), _T("\x43\x1ED9ng II"));
		tmpStr.Format(_T("%f"), nGrpCost);
		rptHeader->SetValue(_T("8"), tmpStr);
		nGrpCostAB += nGrpCost;
	}
	
	nTtCost += nGrpCostAB;
	if(nGrpCostAB > 0)
	{
		rptHeader = rpt.AddDetail(rpt.GetGroupHeader(3));
		rptHeader->SetValue(_T("TotalAmount"), _T("\x43\x1ED9ng B (I + II)"));
		tmpStr.Format(_T("%f"), nGrpCostAB);
		rptHeader->SetValue(_T("8"), tmpStr);
	}
		


	if (nTtCost > 0)
	{
		rptHeader = rpt.AddDetail(rpt.GetGroupHeader(3));
		rptHeader->SetValue(_T("TotalID"), _T("C"));
		rptHeader->SetValue(_T("TotalAmount"), _T("C = A + B"));
		tmpStr.Format(_T("%f"), nTtCost);
		rptHeader->SetValue(_T("8"), tmpStr);
	}

	
	//Page Footer
	//Report Footer
	rptDetail = rpt.GetReportFooter();

	szSQL.Format(_T("SELECT hfe_inspaid, hfe_discount  FROM hms_fee WHERE hfe_docno = %ld and hfe_orderid=%ld and hfe_type='V' and hfe_itemid='VT000002' "),
		nDocumentNo, nOperationIdx);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF())
	{
		rs.GetValue(_T("hfe_inspaid"), nTotalInsPaid);
		rs.GetValue(_T("hfe_discount"), nTotalDiscount);
	}

	if(opt_discount <= 0)
	{
		nTotalInsPaid = 0;
	}
	
	tmpStr.Format(_T("%.f"), ceil(nTotalInsPaid));
	rptDetail->SetValue(_T("TotalInsLimit"), tmpStr);

	tmpStr.Format(_T("%.f"), ceil(nTotalDiscount));
	rptDetail->SetValue(_T("TotalInsPaid"), tmpStr);
	nTotalPatpaid = nTtCost - nTotalDiscount;

	tmpStr.Format(_T("%.f"), ceil(nTotalPatpaid));
	rptDetail->SetValue(_T("TotalPatpaid"), tmpStr);
	rpt.PrintPreview();	
}




BOOL CPrintForms::CheckCardExpire(long nDocumentNo, long nRefIdx)
{
	CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	CString szExpDate, szInsRegDate;

	bool bCardExpire = false;

	szSQL.Format(_T(" SELECT hc_expdate as expdate,") \
				_T("        hd_insregdate as insregdate") \
				_T(" FROM hms_doc") \
				_T(" LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx)") \
				_T(" WHERE hd_docno=%ld AND length(hd_cardno)>1"), nDocumentNo);

	rs.ExecSQL(szSQL);

	if (!rs.IsEOF())
	{
		rs.GetValue(_T("expdate"), szExpDate);
		rs.GetValue(_T("insregdate"), szInsRegDate);
	}
	else
	{
		return bCardExpire;
	}

	szSQL.Format(_T(" SELECT trunc_date(ho_orderdate) AS orderdate") \
				_T(" FROM hms_operation") \
				_T(" WHERE ho_docno=%ld AND ho_hasfee IN('Y','M')") \
				_T("       AND hfe_refstatus IN('O') AND ho_idx=%ld"),
				nDocumentNo, nRefIdx);

	rs.ExecSQL(szSQL);

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("orderdate"), tmpStr);

		if (!bCardExpire)
		{
			if (CompareDate(tmpStr, szInsRegDate) < 0 || CompareDate(tmpStr, szExpDate) > 0)
			{
				bCardExpire = true;
			}
		}
		rs.MoveNext();
	}

	if (!bCardExpire)
	{
		szSQL.Format(_T(" SELECT trunc_date(hpo_orderdate) AS orderdate") \
					_T(" FROM hms_ipharmaorder") \
					_T(" LEFT JOIN hms_ipharmaorderline ON(hpol_orderid=hpo_orderid)") \
					_T(" WHERE hpo_docno=%ld AND hfe_refstatus IN('O')") \
					_T("       AND hpo_reference_id=%ld AND hpo_deptid IN('%s')"),
					nDocumentNo, nRefIdx, pMF->GetDepartmentID());

		rs.ExecSQL(szSQL);

		while (!rs.IsEOF())
		{
			rs.GetValue(_T("orderdate"), tmpStr);

			if (!bCardExpire)
			{
				if (CompareDate(tmpStr, szInsRegDate) < 0 || CompareDate(tmpStr, szExpDate) > 0)
				{
					bCardExpire = true;
				}
			}
			rs.MoveNext();
		}
	}
	return bCardExpire;
}

//void CPrintForms::TM_PrintDrugCabinetUsed(CString szFromDate, CString szToDate)
//{
//	CGuiMainFrame *pMF = (CGuiMainFrame *) AfxGetMainWnd();
//
//	CStringArray strObjects;
//	CStringArray strOrgID;
//
//	strObjects.Add(_T("'1','3','8'"));
//	strObjects.Add(_T("'7'"));
//	strObjects.Add(_T("'2','4','5','6','9','10','11','12'"));
//
//	strOrgID.Add(_T("'PM'"));
//	strOrgID.Add(_T("'MA'"));
//
//	CRecord rs(&pMF->m_db);
//	CRecord rsl(&pMF->m_db);
//	CRecord rsc(&pMF->m_db);
//
//	static CReport rpt;
//	CString szSQL;
//	CString tmpStr, szTemp;
//
//	if (!rpt.Init(_T("Reports/HMS/TM_DANHSACHBENHNHANSUDUNGTHUOCTUTRUC.RPT")))
//		return;
//
//	StringUpper(pMF->m_szHealthService, tmpStr);
//	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
//	StringUpper(pMF->m_szHospitalName, tmpStr);
//	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);
//	
//	rpt.GetReportHeader()->Format(_T("OrderDate"),
//		                          CDateTime::Convert(szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
//								  CDateTime::Convert(szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
//
//	CReportSection* rptDetail;
//	CString szObjects, szOrgIDs;
//	CString szOldDoc, szNewDoc;
//	CString szOldType, szNewType;
//
//	for (int p = 0; p < strObjects.GetSize(); p++)
//	{
//		szObjects.Format(_T("%s"), strObjects[p]);
//
//		for (int q = 0; q < strOrgID.GetSize(); q++)
//		{
//			szOrgIDs.Format(_T("%s"), strOrgID[q]);
//
//			szSQL.Format(_T(" SELECT * ") \
//						_T(" FROM") \
//						_T(" (") \
//						_T("  SELECT distinct hpo_orderid as orderid,") \
//						_T("         hd_docno as docno,") \
//						_T("         hcr_recordno as recordno,") \
//						_T(" 			  trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
//						_T("         hpo_ordertype as ordertype,") \
//						_T("         case when hpo_ordertype in('C') then 1") \
//						_T("              when hpo_ordertype in('M') then 2") \
//						_T("              when hpo_ordertype in('B') then 3") \
//						_T("         else 4 end as idx") \
//						_T("  FROM hms_patient") \
//						_T("  LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
//						_T("  LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
//						_T("  LEFT JOIN hms_ipharmaorder ON(hpo_docno=hd_docno)") \
//						_T("  LEFT JOIN hms_ipharmaorderline ON(hpo_orderid=hpol_orderid)") \
//						_T("  LEFT JOIN m_product ON(mp_product_id=hpol_product_id)") \
//						_T("  LEFT JOIN m_product_item ON(mpi_product_item_id=hpol_product_item_id)") \
//						_T("  LEFT JOIN m_storagelist ON(msl_storage_id=hpo_storage_id)") \
//						_T("  WHERE msl_type IN('C')") \
//						_T("        AND hpo_orderstatus IN('A')") \
//						_T("        AND hd_object IN(%s)") \
//						_T("        AND mp_org_id IN(%s)") \
//						_T("        AND hpo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
//						_T(" ) Tbl") \
//						_T(" ORDER BY idx, ordertype, pname"),
//						szObjects, szOrgIDs,
//						szFromDate, szToDate);
//
//			rs.ExecSQL(szSQL);
//
//			if (rs.IsEOF())
//				continue;
//
//			double nQty = 0;
//			int nIndex = 1, nIdx = 1;
//			//int nRoomID, nBedID;
//			CString szOrderID, szDocNo;
//			CString szTemp;
//			COLDRUGINFO colInfo;
//			CArray<COLDRUGINFO, COLDRUGINFO&> m_arrItem;
//
//			m_arrItem.RemoveAll();
//
//			while (!rs.IsEOF())
//			{
//				if (!szOrderID.IsEmpty())
//					szOrderID += _T(",");
//				szOrderID.AppendFormat(_T("%ld"), str2long(rs.GetValue(_T("orderid"))));
//				rs.MoveNext();
//			}
//
//			if (!szOrderID.IsEmpty())
//			{
//				szSQL.Format(_T(" SELECT distinct mp_product_id as id,") \
//							_T("        mp_name as drugname") \
//							_T(" FROM hms_ipharmaorder") \
//							_T(" LEFT JOIN hms_ipharmaorderline ON(hpo_orderid=hpol_orderid)") \
//							_T(" LEFT JOIN m_product ON(mp_product_id=hpol_product_id)") \
//							_T(" LEFT JOIN m_product_item ON(mpi_product_item_id=hpol_product_item_id)") \
//							_T(" WHERE hpo_orderid IN(%s)") \
//							_T("       AND length(mp_name)>0 AND mp_org_id IN(%s)") \
//							_T(" ORDER BY drugname"),
//							szOrderID, szOrgIDs);
//
//				rsc.ExecSQL(szSQL);
//
//				_fmsg(_T("%s"), szSQL);
//
//				while (!rsc.IsEOF())
//				{ 
//					//MessageBox(_T("OK"));
//					/*if (nIndex <= m_nMaxCol)
//					{
//						szTemp.Format(_T("%d"), nIndex);
//						rsc.GetValue(_T("drugname"), tmpStr);	
//						rptDetail->SetValue(szTemp, tmpStr);
//						nIndex++;
//					}*/
//					if (nIdx % m_nMaxCol == 1)
//						nIdx = 1;
//
//					rsc.GetValue(_T("id"), tmpStr);
//					colInfo.szItemID = tmpStr;
//
//					rsc.GetValue(_T("drugname"), tmpStr);
//					colInfo.szName = tmpStr;
//
//					colInfo.nIndex = nIdx;
//					//colInfo.bIsRep = true;
//					m_arrItem.Add(colInfo);
//
//					nIdx++;
//					rsc.MoveNext();
//				}
//			}
//
//			int nCount = 0;
//			bool bCheck = false;
//			//_msg(_T("%d"), m_arrItem.GetSize());
//			int nMaxVal;
//
//			while (nCount * m_nMaxCol < m_arrItem.GetSize())
//			{
//				rs.MoveFirst();
//
//				nIndex = 1;
//
//				szOldDoc.Empty();
//				szNewDoc.Empty();
//				szNewType.Empty();
//				szOldType.Empty();
//				szOrderID.Empty();
//
//				if ((nCount + 1) * m_nMaxCol < m_arrItem.GetSize())
//					nMaxVal = (nCount + 1) * m_nMaxCol;
//				else
//					nMaxVal = m_arrItem.GetSize();
//
//				if (nCount > 0)
//				{
//					if (rptDetail)
//						rptDetail->SetPageBreak(true);
//				}
//
//				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
//
//				for (int j = nCount * m_nMaxCol; j < nMaxVal; j++)
//				{
//					/*if (j > nCount * m_nMaxCol && j % m_nMaxCol == 0)
//						break;*/
//					colInfo = m_arrItem[j];
//					szTemp.Format(_T("l%d"), colInfo.nIndex);
//					rptDetail->SetValue(szTemp, colInfo.szName);
//				}
//
//				while (!rs.IsEOF())
//				{
//					rs.GetValue(_T("ordertype"), szNewType);
//
//					if (!szNewType.IsEmpty() && szOldType != szNewType)
//					{
//						rptDetail = rpt.AddDetail(rpt.GetGroupHeader(2));
//						if (szNewType == _T("C"))
//							rptDetail->SetValue(_T("GroupName"), _T(""));
//						else if (szNewType == _T("M"))
//							rptDetail->SetValue(_T("GroupName"), _T(""));
//						else if (szNewType == _T("B"))
//							rptDetail->SetValue(_T("GroupName"), _T(""));
//						else
//							rptDetail->SetValue(_T("GroupName"), _T(""));
//
//						szOldType = szNewType;
//						szOldDoc.Empty();
//					}
//
//					rs.GetValue(_T("docno"), szNewDoc);
//
//					if (!szNewDoc.IsEmpty() && szOldDoc != szNewDoc)
//					{
//						if (!szOrderID.IsEmpty())
//						{
//							szSQL.Format(_T(" SELECT mp_product_id as id,") \
//										_T("        mp_name as drugname,") \
//										_T("        sum(hpol_qtyissue) as qty") \
//										_T(" FROM hms_ipharmaorder") \
//										_T(" LEFT JOIN hms_ipharmaorderline ON(hpo_orderid=hpol_orderid)") \
//										_T(" LEFT JOIN m_product ON(mp_product_id=hpol_product_id)") \
//										_T(" LEFT JOIN m_product_item ON(mpi_product_item_id=hpol_product_item_id)") \
//										_T(" LEFT JOIN ad_uom ON(adu_uom_id=mp_product_uom_id)") \
//										_T(" WHERE hpo_orderid IN(%s) AND mp_org_id IN(%s)") \
//										_T(" GROUP BY mp_product_id, mp_name") \
//										_T(" ORDER BY drugname, id"),
//										szOrderID, szOrgIDs);
//
//							rsl.ExecSQL(szSQL);
//							//double nTotalAmount = 0;
//							
//							while (!rsl.IsEOF())
//							{
//								rsl.GetValue(_T("id"), tmpStr);
//								for (int i = nCount * m_nMaxCol; i < nMaxVal; i++)
//								{
//									colInfo = m_arrItem[i];
//									if (colInfo.szItemID == tmpStr)
//									{
//										szTemp.Format(_T("%d"), colInfo.nIndex);
//										rsl.GetValue(_T("qty"), nQty);
//										tmpStr.Format(_T("%d"), nQty);
//										rptDetail->SetValue(szTemp, tmpStr);
//										/*rsl.GetValue(_T("qty"), nQty);
//										colInfo.nTotal += nQty;
//										m_arrItem.SetAt(i, colInfo);*/
//										//_msg(_T("%ld"), colInfo.nTotal);
//										break;
//									}
//								}
//								rsl.MoveNext();
//							}
//						}
//
//						rptDetail = rpt.AddDetail();
//
//						tmpStr.Format(_T("%d"), nIndex);
//						rptDetail->SetValue(_T("Index"), tmpStr);
//
//						rs.GetValue(_T("recordno"), tmpStr);
//						rptDetail->SetValue(_T("RecordNo"), szDocNo);
//
//						rs.GetValue(_T("pname"), tmpStr);
//						rptDetail->SetValue(_T("PatientName"), tmpStr);
//
//						szOrderID.Empty();
//						szOldDoc = szNewDoc;
//						nIndex++;
//					}
//
//					if (!szOrderID.IsEmpty())
//						szOrderID += _T(",");
//					szOrderID.AppendFormat(_T("%ld"), str2long(rs.GetValue(_T("orderid"))));
//					
//					rs.MoveNext();
//				}
//
//				if (!szOrderID.IsEmpty())
//				{
//					szSQL.Format(_T(" SELECT mp_product_id as id,") \
//										_T("        mp_name as drugname,") \
//										_T("        sum(hpol_qtyissue) as qty") \
//										_T(" FROM hms_ipharmaorder") \
//										_T(" LEFT JOIN hms_ipharmaorderline ON(hpo_orderid=hpol_orderid)") \
//										_T(" LEFT JOIN m_product ON(mp_product_id=hpol_product_id)") \
//										_T(" LEFT JOIN m_product_item ON(mpi_product_item_id=hpol_product_item_id)") \
//										_T(" LEFT JOIN ad_uom ON(adu_uom_id=mp_product_uom_id)") \
//										_T(" WHERE hpo_orderid IN(%s) AND mp_org_id IN(%s)") \
//										_T(" GROUP BY mp_product_id, mp_name") \
//										_T(" ORDER BY drugname, id"),
//										szOrderID, szOrgIDs);
//
//					rsl.ExecSQL(szSQL);
//					//double nTotalAmount = 0;
//					while (!rsl.IsEOF())
//					{
//						rsl.GetValue(_T("id"), tmpStr);
//						for (int l = nCount * m_nMaxCol; l < nMaxVal; l++)
//						{
//							colInfo = m_arrItem[l];
//							if (colInfo.szItemID == tmpStr)
//							{
//								szTemp.Format(_T("%d"), colInfo.nIndex);
//								rsl.GetValue(_T("qty"), nQty);
//								tmpStr.Format(_T("%d"), nQty);
//								rptDetail->SetValue(szTemp, tmpStr);
//								break;
//							}
//						}
//						rsl.MoveNext();
//					}
//				}
//
//				bCheck = false;
//				nCount++;
//			}
//
//			rpt.PrintPreview();
//
//		}
//	}
//}


typedef struct tagDrugItemData{
		long product_id;
		CString product_name;
		CString product_uomname;
		float	qty[365+1];
		int		qtywarning;
		int		page;
		int		col_filter;

}DrugItemData;

void CPrintForms::PM_PrintDailyMaterialAndDrugsOfPatient(long nDocumentNo, CString szPrintType, CString szTransactionIDS){
	CGuiMainFrame *pMF = (CGuiMainFrame *) AfxGetMainWnd();
	
	CReportSection* rptDetail = NULL;
	CRecord rs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	CString szTableName, szSQL, szDate, szReportTitle, szSysDate, tmpStr;

	

	CArray<DrugItemData, DrugItemData> items;

	bool bInpatient = true;
	
	
	if(bInpatient)
	{
		szTableName = _T("hms_ipharmaorder");
	
		szSQL.Format(_T(" SELECT hcr_docno as docno,") \
				_T("    trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("    hms_getage( trunc_date(hd_admitdate), hp_birthdate) as age,") \
				_T(" 	extract(year from hp_birthdate) as yearofbirth,") \
				_T("    sys_sel_getname('sys_sex', hp_sex) as sex,") \
				_T("	hms_getobjectname(hd_object)                       AS object, ") \
				_T("    hcr_maindisease as diagnostic,") \
				_T("    trunc_date(hcr_admitdate) as admitdate,") \
				_T("    trunc_date(hcr_dischargedate) as dischargedate,") \
				_T("    htr_deptid,") \
				_T("    hb_roomid,") \
				_T("    hb_bedid,") \
				_T("    GET_DEPARTMENTNAME(htr_deptid) as deptname,") \
				_T("    hrl_name as roomname,") \
				_T("    hbl_name as bedname") \
				_T(" FROM hms_patient") \
				_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
				_T(" LEFT JOIN hms_clinical_record ON(hd_docno=hcr_docno)") \
				_T(" LEFT JOIN hms_treatment_record ON(htr_docno=hcr_docno AND htr_idx=hcr_refidx)") \
				_T(" LEFT JOIN hms_bed ON(hb_docno=hcr_docno AND hb_deptid=htr_deptid AND hb_refidx=hcr_refidx)") \
				_T(" LEFT JOIN hms_roomlist ON(hrl_deptid=htr_deptid and hrl_id=hb_roomid)") \
				_T(" LEFT JOIN hms_bedlist ON(hbl_deptid=htr_deptid and hbl_roomid=hb_roomid and hbl_id=hb_bedid)") \
				_T(" WHERE hd_docno=%ld"), nDocumentNo);
	}
	else
	{
		//Kham benh
		
		szSQL.Format(_T(" SELECT hcr_docno as docno,") \
				_T("    trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("    hms_getage( trunc_date(hd_admitdate), hp_birthdate) as age,") \
				_T(" 	extract(year from hp_birthdate) as yearofbirth,") \
				_T("    sys_sel_getname('sys_sex', hp_sex) as sex,") \
				_T("	hms_getobjectname(hd_object)                       AS object, ") \
				_T("    hd_diagnostic as diagnostic,") \
				_T("    trunc_date(hd_admitdate) as admitdate,") \
				_T("    trunc_date(hd_enddate) as dischargedate,") \
				_T("    hd_deptid,") \
				_T("    hb_roomid,") \
				_T("    hb_bedid,") \
				_T("    GET_DEPARTMENTNAME(hd_deptid) as deptname,") \
				_T("    '' as roomname,") \
				_T("    '' as bedname") \
				_T(" FROM hms_patient") \
				_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
				_T(" WHERE hd_docno=%ld"), nDocumentNo);
		szTableName = _T("hms_pharmaorder");

	}

	rs.ExecSQL(szSQL);

	CString szName, szFromDate, szToDate, szMinDate, szMaxDate;
	CDate dte;
	
	//Print Detail
	int nItem = 1;
	long nProduct_ID;
	CRecord rsc(&pMF->m_db);
	CString szWhere;
	
	
	if(!szTransactionIDS.IsEmpty())
		szWhere.AppendFormat(_T(" and hpo_transaction_id in(%s) and hpo_orderstatus IN('O', 'A','S')  "), szTransactionIDS);
	else
		szWhere.AppendFormat(_T(" and hpo_orderstatus in('A','S') "));
	if(!szPrintType.IsEmpty())
	{
		if(szPrintType == _T("GL"))
			szWhere.AppendFormat(_T(" and msl_org_id in('GL','PM','MA')  "), szPrintType);
		else
			szWhere.AppendFormat(_T(" and msl_org_id IN ('GL', '%s') "), szPrintType);

	}
	//if(szPrintType.IsEmpty())
	//	szWhere.Empty();

	szSQL.Format(_T("SELECT trunc_date(min(hpo_orderdate)) as min_orderdate, trunc_date(max(hpo_orderdate)) max_orderdate ") \
		_T(" from hms_ipharmaorder ") \
		_T(" left join m_storagelist on(msl_storage_id=hpo_storage_id) ") \
		_T(" where hpo_docno=%ld and hpo_orderstatus in('O', 'A','S') ") \
		_T(" %s "), nDocumentNo, szWhere);
	rsl.ExecSQL(szSQL);
	rsl.GetValue(_T("min_orderdate"), tmpStr);
	szMinDate = tmpStr.Left(10);
	rsl.GetValue(_T("max_orderdate"), tmpStr);
	szMaxDate = tmpStr.Left(10);
	int n = CompareDate(szMaxDate, szMinDate);

	if(!szPrintType.IsEmpty() && n > 29)
	{
		szSQL.Format(_T("SELECT TO_DATE('%s','YYYY-MM-DD')-29 as min_date FROM DUAL"), szMaxDate);
		rsl.ExecSQL(szSQL);
		rsl.GetValue(_T("min_date"), tmpStr);
		szMinDate = tmpStr.Left(10);
	}

	szFromDate = szMinDate;
	dte.ParseDate(szFromDate);
	rptDetail = NULL;
	szWhere.Empty();
	if(!szTransactionIDS.IsEmpty())
		szWhere.AppendFormat(_T(" and hpo_transaction_id in(%s) and hpo_orderstatus IN('O', 'A','S')  "), szTransactionIDS);
	else
		szWhere.AppendFormat(_T(" and hpo_orderstatus in('A','S') "));


	if(!szPrintType.IsEmpty())
	{
		if(szPrintType == _T("GL"))
			szWhere.AppendFormat(_T(" and product_org_id in('GL','PM','MA')  "), szPrintType);
		else
			szWhere.AppendFormat(_T(" and product_org_id='%s' "), szPrintType);
		szWhere.AppendFormat(_T(" and trunc(hpo_orderdate) between cast_string2date('%s')-30 and cast_string2date('%s') "), szMaxDate, szMaxDate);

	}
	//_msg(_T("%s"), szPrintType);
	//if(szPrintType.IsEmpty())
	//	szWhere.Empty();

	szSQL.Format(_T(" SELECT product_name,") \
_T("   product_uomname,") \
_T("   product_id,") \
_T("   product_qtywarning, ") \
_T("   SUM(hpol_qtyissue) AS qtyissue") \
_T(" FROM hms_ipharmaorder") \
_T(" LEFT JOIN hms_ipharmaorderline") \
_T(" ON(hpol_orderid=hpo_orderid)") \
_T(" LEFT JOIN m_productitem_view") \
_T(" ON(hpol_product_item_id=product_item_id)") \
_T(" WHERE hpo_docno        =%ld  ") \
_T(" %s ") \
_T(" AND hpo_ordertype <>'M'") \
_T(" AND hpol_qtyissue  > 0") \
_T(" GROUP BY product_name,") \
_T("   product_uomname,") \
_T("   product_id,") \
_T("   product_uomindex, product_qtywarning ") \
_T(" ORDER BY ") \
_T("   product_uomindex, product_name "), nDocumentNo, szWhere);



	items.RemoveAll();
	rsl.ExecSQL(szSQL);
	
	int nMaxPage = 0;
	while(!rsl.IsEOF())
	{
		DrugItemData dta;
		dta.page = 1;
		rsl.GetValue(_T("product_id"), dta.product_id);
		rsl.GetValue(_T("product_name"), dta.product_name);
		rsl.GetValue(_T("product_uomname"), dta.product_uomname);
		rsl.GetValue(_T("product_qtywarning"), dta.qtywarning);
		nProduct_ID = dta.product_id;
		for (int i =0; i <= 365; i++)
			dta.qty[i] = 0;

		szSQL.Format(_T(" SELECT ") \
_T("   trunc_date(hpo_orderdate) as orderdate,") \
_T("   SUM(hpol_qtyissue) AS qtyissue") \
_T(" FROM hms_ipharmaorder") \
_T(" LEFT JOIN hms_ipharmaorderline") \
_T(" ON(hpol_orderid=hpo_orderid)") \
_T(" LEFT JOIN m_productitem_view") \
_T(" ON(hpol_product_item_id=product_item_id)") \
_T(" WHERE hpo_docno        =%ld  ") \
_T(" and product_id=%ld ") \
_T(" %s ") \
_T(" AND hpo_ordertype <>'M'") \
_T(" AND hpol_qtyissue  > 0") \
_T(" GROUP BY trunc_date(hpo_orderdate)") \
_T(" ORDER BY trunc_date(hpo_orderdate) "), nDocumentNo, nProduct_ID, szWhere);

		rsc.ExecSQL(szSQL);
		
		while(!rsc.IsEOF())
		{
			rsc.GetValue(_T("orderdate"), szDate);
			szDate.Replace(_T("-"), _T("/"));

			dte.ParseDate(szMinDate);
			for (int i =0; i <= 365; i++)
			{
				tmpStr = dte.GetDate();
				dte ++;
				if(tmpStr == szDate.Left(10))
				{
					rsc.GetValue(_T("qtyissue"), tmpStr);
					dta.qty[i] = str2float(tmpStr);
					if (dta.qty[i] > 0)
					{
						dta.page = (int)floor(i/15.0);
					}
					else
						dta.page = 0;
					nMaxPage = max(nMaxPage, dta.page);
					break;
				}
			}
			rsc.MoveNext();
		}

		items.Add(dta);

		rsl.MoveNext();
	}

	int nLast = 0;
	dte.ParseDate(szMinDate);



	for (int nPage =0; nPage <= nMaxPage; nPage++)
	{
		nItem = 1;

		CReport rpt;
			
		if(!rpt.Init(_T("Reports/HMS/TM_NHATTRINHVACONGKHAITHUOC.RPT"),true) )
		{
			return;
		}
		//Print Report Header
		rptDetail = rpt.GetPageHeader();

		rptDetail->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
		rptDetail->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
		rptDetail->SetValue(_T("ReportTitle"), szReportTitle);

		rs.GetValue(_T("deptname"), tmpStr);
		rptDetail->SetValue(_T("DEPARTMENT"), tmpStr);


		//Print page header
		

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("PatientName"), tmpStr);

		rs.GetValue(_T("age"), tmpStr);
		rptDetail->SetValue(_T("Age"), tmpStr);

		rs.GetValue(_T("sex"), tmpStr);
		rptDetail->SetValue(_T("Sex"), tmpStr);

		rs.GetValue(_T("bedname"), tmpStr);
		rptDetail->SetValue(_T("Bednumber"), tmpStr);

		rs.GetValue(_T("roomname"), tmpStr);
		rptDetail->SetValue(_T("Room"), tmpStr);

		rptDetail->SetValue(_T("object"), rs.GetValue(_T("object")));
		rptDetail->SetValue(_T("yearofbirth"), rs.GetValue(_T("yearofbirth")));

		tmpStr = CDate::Convert(rs.GetValue(_T("Admitdate")), yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("AdmitDate"), tmpStr);

		rs.GetValue(_T("diagnostic"), tmpStr);
		rptDetail->SetValue(_T("Diagnostic"), tmpStr);

		tmpStr = CDate::Convert(rs.GetValue(_T("Dischargedate")), yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("DischargeDate"), tmpStr);


		rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
		
		for (int i =0; i < 15; i++){
			tmpStr = dte.GetDate(ddmmyyyy);
			szName.Format(_T("l%d"), i+1);
			rptDetail->SetValue(szName, tmpStr.Left(5));
			dte ++;
		}

		bool bFound = false;
		for (int i = 0; i < items.GetCount(); i++)
		{
			float ttlqty=0;
			DrugItemData dta = items[i];
		//	if(dta.page > nPage)
		//		continue;

			nLast = 0;
			for (int n = 365; n >= 0; n--)
			{
				if(dta.qty[n] > 0)
				{
					nLast = n;
					break;
				}
				
			}

			

			for (int j =0; j < 15; j++)
				ttlqty += dta.qty[nPage*15+j];

			
			if (ttlqty > 0 )
			{
				int x = 0;

				dta.col_filter = 0;
			//	if (dta.qtywarning > 0 )
				{

					for (int k =0; k < nLast; k++){

						if (dta.qty[k] <= 0)
						{
							dta.col_filter = 0;
							x = 0;
						}
						else if (dta.qty[k] > 0)
						{
							x++;
						}
						if(x >= dta.qtywarning)
						{
							dta.col_filter = k;
							x=0;
							break;
						}
					}
				}

				rptDetail = rpt.AddDetail();
				bFound = true;
				
				tmpStr.Format(_T("%d"), nItem++);
				rptDetail->SetValue(_T("Index"), tmpStr);	
				rptDetail->SetValue(_T("DrugName"), dta.product_name);
				rptDetail->SetValue(_T("Unit"), dta.product_uomname);
				for (int j =0; j < 15; j++)
				{
					tmpStr.Format(_T("%.2f"), dta.qty[nPage*15+j]);
					szName.Format(_T("%d"), j+1);
					rptDetail->SetValue(szName, tmpStr);
					if (nPage*15+j > dta.col_filter && dta.col_filter > 0)
					{
						CReportItem* obj = rptDetail->GetItem(szName);
						if(obj)
						{
							obj->SetBkColor(RGB(200, 200, 200));
						}
					}
				}

				if(nLast >= nPage*15 && nLast <= nPage*15+15)
				{
					tmpStr.Format(_T("%.2f"), ttlqty);
					rptDetail->SetValue(_T("32"), tmpStr);
				}
				

			}

		}

		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		szSysDate = pMF->GetSysDate(); 
		szDate.Format(rptDetail->GetValue(_T("PrintDate")),szSysDate.Mid(8,2),szSysDate.Mid(5,2),szSysDate.Left(4));
		rptDetail->SetValue(_T("PrintDate"), szDate);

		if(bFound) rpt.PrintPreview();
	}
	
	

}

void CPrintForms::FM_PrintCollectDetailbyItem(long nCashID){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szSysDate;

	int nIdx = 0;

	double nTemp = 0;
	long double nGrpTimes = 0, nGrpExam = 0, nGrpIn = 0, nGrpAmount = 0;
	long double nTotalExam = 0, nTotalIn = 0, nTotalAmount = 0;
	szSQL.Format(_T("SELECT min(hfe_date) fromdate, max(hfe_date) todate FROM hms_fee_invoice WHERE hfe_cash_id = %ld"), nCashID);
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{
		rs.GetValue(_T("fromdate"), m_szFromDate);
		rs.GetValue(_T("todate"), m_szToDate);
	}
	szSQL.Format(_T(" SELECT    hfe_depttype depttype, ") \
				_T("           hfe_deptid deptid, ") \
				_T("           hfg_index                     AS idx, ") \
				_T("           sd_name                       AS deptname, ") \
				_T("           hfe_desc itemname, ") \
				_T("           SUM(count_patient)            AS total, ") \
				_T("           SUM(exam_fee)                 AS exam_fee, ") \
				_T("           SUM(inpatient_fee)            AS inpatient_fee, ") \
				_T("           SUM(exam_fee + inpatient_fee) AS total_amount ") \
				_T(" FROM      (SELECT    f.hfe_docno, ") \
				_T("                      f.hfe_itemid, ") \
				_T("                      f.hfe_desc, ") \
				_T("                      CASE WHEN f.hfe_deptid IN( 'C1.1', 'C1.2', 'C1.3' ) THEN Cast('KB' AS NVARCHAR2(7)) ") \
				_T("                        WHEN f.hfe_group = 'F0000' THEN Cast('XX' AS NVARCHAR2(7)) ") \
				_T("                        WHEN f.hfe_deptid IN( 'KD' ) THEN Cast('KD' AS NVARCHAR2(7)) ") \
				_T("                        WHEN f.hfe_deptid IN( 'C10' ) THEN Cast('KTB' AS NVARCHAR2(7)) ") \
				_T("                      ELSE Cast('DT' AS NVARCHAR2(7)) ") \
				_T("                      END AS hfe_depttype, ") \
				_T("                      f.hfe_deptid, ") \
				_T("                      f.hfe_class, ") \
				_T("                      f.hfe_type, ") \
				_T("                      f.hfe_group, ") \
				_T("                      CASE WHEN i.hfe_class IN( 'E', 'X' ) THEN f.hfe_patpaid ") \
				_T("                      ELSE 0 ") \
				_T("                      END AS exam_fee, ") \
				_T("                      CASE WHEN i.hfe_class NOT IN( 'E', 'X' ) THEN f.hfe_patpaid ") \
				_T("                      ELSE 0 ") \
				_T("                      END AS inpatient_fee, ") \
				_T("                      CASE WHEN f.hfe_group = 'C0000' THEN f.hfe_quantity ") \
				_T("                      ELSE 1 ") \
				_T("                      END AS count_patient ") \
				_T("            FROM      hms_fee_invoice i ") \
				_T("            LEFT JOIN (SELECT    hfe_invoiceno, ") \
				_T("                                 hfe_docno, ") \
				_T("                                 hfe_class, ") \
				_T("                                 hfe_type, ") \
				_T("                                 hfe_group, ") \
				_T("                                 CASE WHEN hfe_type IN ('P', 'T') THEN hfl_deptid ") \
				_T("								 ELSE CASE WHEN hfe_type = 'O' THEN ho_pdeptid ELSE hfe_deptid ") \
				_T("                                 END END hfe_deptid, ") \
				_T("                                 hfe_itemid, ") \
				_T("                                 CASE WHEN he_deptid = 'C1.1' AND hfe_type = 'E' AND he_roomid IN (50, 54) ") \
				_T("								 THEN hfe_desc||'-'||(select hrl_name from hms_roomlist where hrl_deptid = 'C1.1' and hrl_id = he_roomid) ELSE hfe_desc END hfe_desc, ") \
				_T("                                 hfe_quantity, ") \
				_T("                                 hfe_patpaid ") \
				_T("                       FROM      hms_fee ") \
				_T("					   LEFT JOIN hms_operation ON (hfe_type = 'O' AND hfe_docno = ho_docno AND hfe_orderid = ho_idx)") \
				_T("                       LEFT JOIN hms_fee_list ON( hfl_feeid = hfe_itemid ) ") \
				_T("					   LEFT JOIN hms_exam ON (he_docno = hfe_docno AND he_receptidx = hfe_orderid)") \
				_T("                       WHERE     hfe_status IN ('P', 'R') AND hfe_type IN( 'E', 'V', 'O', 'P', ") \
				_T("                                              'T', 'X', 'F' ) ") \
				_T("                       UNION ALL ") \
				_T("                       SELECT hfe_invoiceno, ") \
				_T("                              hfe_docno, ") \
				_T("                              hfe_class, ") \
				_T("                              hfe_type, ") \
				_T("                              hfe_group, ") \
				_T("                              hfe_deptid, ") \
				_T("                              hfe_itemid, ") \
				_T("                              hfe_desc, ") \
				_T("                              hfe_quantity, ") \
				_T("                              hfe_patpaid ") \
				_T("                       FROM   hms_fee ") \
				_T("                       WHERE  hfe_type IN( 'B', 'J' ) ") \
				_T("                       UNION ALL ") \
				_T("                       SELECT    hfe_invoiceno, ") \
				_T("                                 hfe_docno, ") \
				_T("                                 hfe_class, ") \
				_T("                                 hfe_type, ") \
				_T("                                 hfe_group, ") \
				_T("                                 CASE WHEN mp_org_id = 'PM' THEN Cast('KD' AS NVARCHAR2(7)) ") \
				_T("                                   WHEN mp_org_id = 'MA' THEN Cast('C10' AS NVARCHAR2(7)) ") \
				_T("                                 ELSE hfe_deptid ") \
				_T("                                 END AS hfe_deptid, ") \
				_T("                                 hfe_itemid, ") \
				_T("                                 hfe_desc, ") \
				_T("                                 hfe_quantity, ") \
				_T("                                 hfe_patpaid ") \
				_T("                       FROM      hms_fee ") \
				_T("                       LEFT JOIN hms_pharmaretailorder_view ON( hfe_docno = hpo_docno ") \
				_T("                                                      AND hfe_orderid = hpo_orderid ) ") \
				_T("                       LEFT JOIN m_product_item ON( Cast(mpi_product_item_id AS NVARCHAR2(15)) = hfe_itemid ) ") \
				_T("                       LEFT JOIN m_product ON( mp_product_id = mpi_product_id ) ") \
				_T("                       WHERE     hfe_status IN ('P', 'R') AND hfe_type IN( 'R', 'D', 'M' ) AND hpo_orderid > 0) f ON( f.hfe_docno = i.hfe_docno ") \
				_T("                                                                     AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
				_T("            WHERE hfe_cash_id = %ld ") \
				_T("			AND f.hfe_patpaid > 0 AND hfe_payment > 0") \
				_T("                  AND i.hfe_status = 'P') tbl ") \
				_T(" LEFT JOIN hms_fee_group ON( hfg_id = tbl.hfe_group ) ") \
				_T(" LEFT JOIN sys_dept ON( sd_id = hfe_deptid ) ") \
				_T(" GROUP     BY hfe_depttype, ") \
				_T("              hfe_deptid, ") \
				_T("              hfg_index, ") \
				_T("              sd_name, ") \
				_T("              hfe_desc ") \
				_T(" ORDER     BY hfe_depttype, ") \
				_T("              hfe_deptid, ") \
				_T("              hfg_index, ") \
				_T("              hfe_desc "), nCashID);

	int nCount = rs.ExecSQL(szSQL);

	if (nCount <=0 )
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/HF_CHITIETTHUTIENTHEODANHMUC.RPT")))
		return;

	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_szHealthService);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_szHospitalName);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection *rptDetail;
	CReportSection *rptSector;
	CString szOldFeeGrp, szNewFeeGrp;

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("deptid"), szNewFeeGrp);
		if (szNewFeeGrp != szOldFeeGrp)
		{
			if (nGrpTimes > 0)
			{
				rptSector = rpt.AddDetail(rpt.GetGroupFooter(1));
				FormatCurrency(nGrpTimes, tmpStr);
				rptSector->SetValue(_T("s3"), tmpStr);
				FormatCurrency(nGrpExam, tmpStr);
				rptSector->SetValue(_T("s4"), tmpStr);
				FormatCurrency(nGrpIn, tmpStr);
				rptSector->SetValue(_T("s5"), tmpStr);
				FormatCurrency(nGrpAmount, tmpStr);
				rptSector->SetValue(_T("s6"), tmpStr);
			}
			nIdx++;

			if (szNewFeeGrp != _T("X") && szNewFeeGrp != _T("NA"))
				tmpStr.Format(_T("%d. %s"), nIdx, rs.GetValue(_T("deptname")));
			else if (szNewFeeGrp == _T("X"))
				tmpStr.Format(_T("%d. Thu ph\xED kh\xE1\x63"), nIdx);
			else if (szNewFeeGrp == _T("NA"))
				tmpStr.Format(_T("%d. \x43h\x1B0\x61 thi\x1EBFt l\x1EADp kho\x61 th\x1EF1\x63 hi\x1EC7n"), nIdx);

			rptSector = rpt.AddDetail(rpt.GetGroupHeader(0));
			rptSector->SetValue(_T("GroupName"), tmpStr);
			szOldFeeGrp = szNewFeeGrp;
			nGrpTimes = 0;
			nTotalExam += nGrpExam;
			nGrpExam = 0;
			nTotalIn += nGrpIn;
			nGrpIn = 0;
			nTotalAmount += nGrpAmount;
			nGrpAmount = 0;
		}
		rptDetail = rpt.AddDetail();

		rs.GetValue(_T("itemname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("total"), nTemp);
		//FormatCurrency(nTemp, tmpStr);
		tmpStr.Format(_T("%.0f"), nTemp);
		rptDetail->SetValue(_T("3"), tmpStr);
		nGrpTimes += nTemp;

		rs.GetValue(_T("exam_fee"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		nGrpExam += nTemp;

		rs.GetValue(_T("inpatient_fee"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		nGrpIn += nTemp;

		rs.GetValue(_T("total_amount"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		nGrpAmount += nTemp;

		rs.MoveNext();
	}
	if (nGrpTimes > 0)
	{
		rptSector = rpt.AddDetail(rpt.GetGroupFooter(1));
		FormatCurrency(nGrpTimes, tmpStr);
		rptSector->SetValue(_T("s3"), tmpStr);
		FormatCurrency(nGrpExam, tmpStr);
		rptSector->SetValue(_T("s4"), tmpStr);
		FormatCurrency(nGrpIn, tmpStr);
		rptSector->SetValue(_T("s5"), tmpStr);
		FormatCurrency(nGrpAmount, tmpStr);
		rptSector->SetValue(_T("s6"), tmpStr);

			//Vertical total
		nTotalExam += nGrpExam;
		nTotalIn += nGrpIn;
		nTotalAmount += nGrpAmount;
	}

	FormatCurrency(nTotalExam, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("t4"), tmpStr);
	FormatCurrency(nTotalIn, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("t5"), tmpStr);
	FormatCurrency(nTotalAmount, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("t6"), tmpStr);
	szSysDate = pMF->GetSysDate();

	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();	
}

#include "HMSFeeManager.h"
void CPrintForms::FM_PrintDischargePaymentInvoice(long nDocumentNo, long nInvoiceNo)
{

	CHMSFeeManager fm(nDocumentNo);
	fm.PrintDischargePayment(nInvoiceNo);
	return;
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szObjectType;
	szSQL.Format(_T("select ho_type  as objecttype from hms_fee_invoice left join hms_object on(ho_id=hfe_objectid) where hfe_docno=%ld and hfe_invoiceno=%ld"), nDocumentNo, nInvoiceNo);
	rs.ExecSQL(szSQL);
	rs.GetValue(_T("objecttype"), szObjectType);
	
	
	

	if(szObjectType == _T("I") || szObjectType == _T("C"))
	{
		FM_PrintInsuranceDischargePaymentInvoice(nDocumentNo, nInvoiceNo);
		szSQL.Format(_T("SELECT count(*) FROM hms_fee WHERE hfe_invoiceno=%ld and hfe_docno = %ld and hfe_object=7 "), nInvoiceNo, nDocumentNo);
		rs.ExecSQL(szSQL);
		if(rs.GetIntValue() > 0)
		{
			FM_PrintServiceDischargePaymentInvoice(nDocumentNo, nInvoiceNo);
		}
	}
	else if(szObjectType == _T("D") || szObjectType == _T("P"))
	{
		FM_PrintPolicyDischargePaymentInvoice(nDocumentNo, nInvoiceNo);
		szSQL.Format(_T("SELECT count(*) FROM hms_fee WHERE hfe_invoiceno=%ld and hfe_docno = %ld and hfe_object=7 "), nInvoiceNo, nDocumentNo);
		rs.ExecSQL(szSQL);
		if(rs.GetIntValue() > 0)
		{
			FM_PrintServiceDischargePaymentInvoice(nDocumentNo, nInvoiceNo);
		}

	}
	else if(szObjectType == _T("S"))
	{
		FM_PrintServiceDischargePaymentInvoice(nDocumentNo, nInvoiceNo);
		szSQL.Format(_T("SELECT count(*) FROM hms_fee WHERE hfe_invoiceno=%ld and hfe_docno=%ld and hfe_object<>7 "), nInvoiceNo, nDocumentNo);
		rs.ExecSQL(szSQL);
		if(rs.GetIntValue() > 0)
		{
			FM_PrintInsuranceDischargePaymentInvoice(nDocumentNo, nInvoiceNo);
		}
	}

}

void CPrintForms::FM_PrintServiceDischargePaymentInvoice(long nDocumentNo, long nInvoiceNo)
{
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
//	if (!pMF->IsInPatient())
//		return;

	CReport rpt;
	CString tmpStr, szSQL, szWhere, szAdmitDate, szEndDate,szDepartments, szAddonedayoutofhospital,szNumbertreat;
	CRecord rs(&pMF->m_db);
	CRecord rsd(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	int		nDeskNo=0;
	CString szRecvDate;
	CString	szObjectType;
	CString	szDescription;
	CString szPrintMaterialOperation = _T("Y");
	CString	szOutPatient;


	double	nTotalCost=0,				//Tong chi phi
			nTotalInsCost=0,
			nTotalDiscount = 0,			//Tong Mien giam
			nTotalPatPaid =0,			//Tong benh nhan tra
			nTotalDiffPaid = 0,	//Tong so tra chenh lech
			nTotalPayable=0,
			nTotalDeposit = 0,			//Tong tam ung
			nTotalFree = 0,				//Tong mien phi
			nTotalRefund=0,				//Tong hoan tra
			nDepositPayable=0,			//So tien tra tu tam ung
			nTotalPayment = 0;
	
	
	double nSumFoodAmount=0;			


	szWhere.AppendFormat(_T(" and hfe_invoiceno=%ld "), nInvoiceNo);

	szSQL.Format(_T(" SELECT *") \
				_T(" FROM") \
				_T(" (") \
				_T("  SELECT hd_patientno as patientno,") \
				_T("         hd_docno as docno,") \
				_T("         trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("         hp_birthdate as birthdate,") \
				_T("         hms_getage(trunc_date(hd_admitdate), hp_birthdate) as age,") \
				_T("         extract(year from hp_birthdate) as yearofbirth, ") \
				_T("         Get_Selection_Desc('sys_sex', hp_sex) as sex,") \
				_T("         hp_sex as sex_id,") \
				_T("         hp_ethnic as ethnic,") \
				_T("         Get_Selection_Desc('sys_ethnic', hp_ethnic) as ethnicdesc, ") \
				_T("         hp_dtladdr as detailaddress,") \
				_T("         hms_getaddress(hp_provid,hp_distid, hp_villid) as address,") \
				_T("         hp_workplace as workplace,") \
				_T("         hd_doctor as doctor,") \
				_T("         hd_createdby as createdby,") \
				_T("         hd_icd as icd10,") \
				_T("         hd_diagnostic as diagnostic,") \
				_T("         hd_reldisease as reldisease,") \
				_T("         hd_emergency as emergency, ") \
				_T("         Get_Selection_Desc('hms_result',hd_result) as result,") \
				_T("         hd_admitdate as examdate,") \
				_T("         hd_enddate as enddate,") \
				_T(" hd_outpatient, ") \
				_T("         hcr_recordno as recordno ,") \
				_T("         hcr_admitdate as admitdate,") \
				_T("         hcr_dischargedate as dischargedate,") \
				_T("         hcr_sumtreat as TreatmentDay,") \
				_T("         hcr_mainicd as mainicd ,") \
				_T("         hcr_maindisease as maindisease ,") \
				_T("         hcr_treatdoctor as treatdoctor, ") \
				_T("         hcr_reldisease as ireldisease,") \
				_T("         hcr_suggestion as isuggestion,") \
				_T("         hd_suggestion as esuggestion,") \
				_T("         Get_Selection_Desc('hms_result' ,hcr_result)  as iresult,") \
				_T("         hfe_deptid as deptid,") \
				_T("         hfe_serialno as serialno,") \
				_T("         hfe_receiptno as recvno,") \
				_T("         hfe_date as recvdate, ") \
				_T("         hfe_staff as receiver,") \
				_T("         hfe_desc as reason,") \
				_T("         hfe_amount as cost,") \
				_T("         hfe_discount as discount,") \
				_T("         hfe_patpaid as patpaid,") \
				_T("         hfe_deposit as deposit_amount, ") \
				_T("         hfe_payment as payment_amount,") \
				_T("         hfe_refund as refund_amount, ") \
				_T("         hfe_freeamount as free_amount,") \
				_T("         hfe_eat_amount as food_amount,") \
				_T("         row_number() over (partition by hd_docno order by hd_docno) as dn") \
				_T(" FROM hms_fee_invoice ") \
				_T(" LEFT JOIN hms_doc ON(hd_docno=hfe_docno)") \
				_T(" LEFT JOIN hms_patient ON(hp_patientno=hd_patientno)") \
				_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno) ") \
				_T(" LEFT JOIN hms_treatment_record ON(hfe_docno=htr_docno) ") \
				_T(" LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx) ") \
				_T(" WHERE hfe_docno=%ld and hfe_type<>'E' %s ") \
				_T(" ) Tbl") \
				_T(" WHERE dn=1"), nDocumentNo, szWhere);
	
	
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		return;
	}

	rs.GetValue(_T("cost"), nTotalCost);
	rs.GetValue(_T("discount"), nTotalDiscount);
	rs.GetValue(_T("patpaid"), nTotalPatPaid);
	rs.GetValue(_T("deposit_amount"), nTotalDeposit);
	rs.GetValue(_T("free_amount"), nTotalFree);
	rs.GetValue(_T("refund_amount"), nTotalRefund);
	rs.GetValue(_T("payment_amount"), nTotalPayment);
	rs.GetValue(_T("reason"), szDescription);
	rs.GetValue(_T("deskno"), nDeskNo);

	rs.GetValue(_T("food_amount"), nSumFoodAmount);

	rs.GetValue(_T("hd_outpatient"), szOutPatient);
	if(szOutPatient == _T("Y"))
	{
		if(!rpt.Init(_T("Reports/HMS/HF_DISCHARGEPAYMENTOUT.RPT"), false) )
			return;
	}
	else
	{
		if(!rpt.Init(_T("Reports/HMS/HF_DISCHARGEPAYMENT.RPT"), false) )
			return;
	}
	//Report Header
	tmpStr.Empty();
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);

	rs.GetValue(_T("docno"), nDocumentNo);
	tmpStr.Format(_T("%ld"), nDocumentNo);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	tmpStr.Format(_T("%s-%s-%s"), rs.GetValue(_T("serialno")), rs.GetValue(_T("bookno")),  rs.GetValue(_T("recvno")));
	rpt.GetReportHeader()->SetValue(_T("ReceiptNo"), tmpStr);
	tmpStr.Format(_T("%ld"), nInvoiceNo);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);


	CString szDate;
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")), tmpStr.Mid(11, 5),tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), szDate);


	rs.GetValue(_T("recordno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RecordNo"), tmpStr);
	//rs.GetValue(_T("pname"), tmpStr);
	

	CString szPatientName;
	StringUpper(rs.GetValue(_T("pname")), szPatientName);
	rpt.GetReportHeader()->SetValue(_T("PATIENTNAME"), szPatientName);
	rs.GetValue(_T("age"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), tmpStr);
	rs.GetValue(_T("yearofbirth"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);

	rs.GetValue(_T("birthdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("BirthDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

	rs.GetValue(_T("sex"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Sex"), tmpStr);

	rs.GetValue(_T("sex_id"), tmpStr);
	if(tmpStr == _T("F"))
		rpt.GetReportHeader()->SetValue(_T("Female"), _T("X"));
	else
		rpt.GetReportHeader()->SetValue(_T("Male"), _T("X"));


	rs.GetValue(_T("ethnic"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ethnic"), tmpStr);
	rs.GetValue(_T("ethnicdesc"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ethnicdesc"), tmpStr);
	
	rs.GetValue(_T("workplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Workplace"), tmpStr);

	CString szAddress;
	rs.GetValue(_T("detailaddress"), tmpStr);
	szAddress = tmpStr;
	rs.GetValue(_T("address"), tmpStr);
	if(!szAddress.IsEmpty())
		szAddress += _T(" - ");
	szAddress += tmpStr;
	rpt.GetReportHeader()->SetValue(_T("Address"), szAddress);
	

	rs.GetValue(_T("transplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("TransferHospital"), tmpStr);
	
	
	rs.GetValue(_T("emergency"), tmpStr);
	if (tmpStr == _T("Y") )
		tmpStr = _T("X");
	else
		tmpStr.Empty();

	rpt.GetReportHeader()->SetValue(_T("Emergency"), tmpStr);

	rs.GetValue(_T("maindisease"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);
		rs.GetValue(_T("ireldisease"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("RelationDisease"), tmpStr);
		rs.GetValue(_T("mainicd"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("ICD10"), tmpStr);
		rs.GetValue(_T("iresult"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Result"), tmpStr);
		rs.GetValue(_T("isuggestion"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Suggestion"), pMF->GetSelectionString(_T("hms_suggestion"), tmpStr));
	
	
	
	CString szDischargeDate;
	rs.GetValue(_T("admitdate"), tmpStr);

	rpt.GetReportHeader()->SetValue(_T("admitdate"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	
	rs.GetValue(_T("dischargedate"), tmpStr);
	szDischargeDate = tmpStr;
	
	if(szDischargeDate.Left(4) != _T("1752"))
			rpt.GetReportHeader()->SetValue(_T("dischargedate"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	else
			rpt.GetReportHeader()->SetValue(_T("dischargedate"), _T(""));

	rs.GetValue(_T("TreatmentDay"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("TreatmentDay"), tmpStr);	

	CString szData;
	tmpStr.Format(_T("%s-%s-%s"), rs.GetValue(_T("bookno")), rs.GetValue(_T("serialno")), rs.GetValue(_T("recvno")));
	rpt.GetReportFooter()->SetValue(_T("SerialNo"), tmpStr);
	rs.GetValue(_T("recvdate"), szRecvDate);	
	rpt.GetReportFooter()->SetValue(_T("ReceiptDate"), CDateTime::Convert(szRecvDate, yyyymmdd|hhmm,ddmmyyyy|hhmm));
	szData.Format(rpt.GetReportFooter()->GetValue(_T("ReceiptDate1")), szRecvDate.Mid(11,5),szRecvDate.Mid(8,2),szRecvDate.Mid(5,2),szRecvDate.Left(4));	
	rpt.GetReportFooter()->SetValue(_T("ReceiptDate1"),szData);
	szData.Format(rpt.GetReportFooter()->GetValue(_T("ReceiptDate2")), szRecvDate.Mid(11,5),szRecvDate.Mid(8,2),szRecvDate.Mid(5,2),szRecvDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("ReceiptDate2"),szData);
	
	rs.GetValue(_T("treatdoctor"), tmpStr);
	tmpStr.Trim();
	rpt.GetReportFooter()->SetValue(_T("Doctor"), GetDoctorName(tmpStr));

	CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
	if(img)
	{
		img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img->SetFixedHeight(true);
	}
	tmpStr.Empty();
	rs.GetValue(_T("receiver"), tmpStr);
	tmpStr.Trim();
	rpt.GetReportFooter()->SetValue(_T("ReceiverBy"), GetDoctorName(tmpStr));
	CReportItem *img2 = rpt.GetReportFooter()->GetItem(_T("Signature2"));
	if(img2)
	{
		img2->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img2->SetFixedHeight(false);
	}	
	
	
	// Nguoi tiep don
	tmpStr.Empty();
	rs.GetValue(_T("createdby"), tmpStr);
	tmpStr.Trim();
	rpt.GetReportFooter()->SetValue(_T("createdby"), GetDoctorName(tmpStr));
	CReportItem *img3 = rpt.GetReportFooter()->GetItem(_T("Signature3"));
	if(img3)
	{
		img3->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img3->SetFixedHeight(false);
	}	


	CReportSection* rptDetail=NULL;
	CReportSection *grp;
	CRecord grs(&pMF->m_db);
	CString szFeeID, szID;
	CString szSubItem, szType;

	TCHAR *lpszChapter[] = {_T("I"), _T("II"), _T("III"), _T("IV"), _T("V"), _T("VI"), _T("VII"), _T("VIII"), _T("IX"), _T("X"), _T("XI")};


	CRecord drs(&pMF->m_db);
	CString szDeptID;
	CStringArray arDepartmentList;

	szSQL.Format(_T(" SELECT deptid, admitdate, dischargedate, mainicd, maindisease") \
					_T(" FROM") \
					_T(" (") \
					_T("  select htr_deptid as deptid,") \
					_T("         htr_admitdate as admitdate,") \
					_T("         htr_dischargedate as dischargedate,") \
					_T("         htr_mainicd as mainicd,") \
					_T("         htr_maindisease as maindisease,") \
					_T("         htr_idx as idx") \
					_T("  from hms_treatment_record") \
					_T("  where htr_docno=%ld") \
					_T("  union all") \
					_T("  select hd_admitdept,") \
					_T("         hd_admitdate,") \
					_T("         hd_enddate,") \
					_T("         hd_icd,") \
					_T("         hd_diagnostic,") \
					_T("         0") \
					_T("  from hms_doc") \
					_T("  where hd_docno=%ld") \
					_T(" ) tbl") \
					_T(" WHERE deptid in(select distinct hfe_deptid from hms_fee where hfe_invoiceno=%ld)") \
					_T(" ORDER BY idx"), nDocumentNo, nDocumentNo, nInvoiceNo);

		drs.ExecSQL(szSQL);

		RECEIPTINFO pi;
		CString szDepts;
		szDepts.Empty();
		CString szTreatDeptID;
		while (!drs.IsEOF())
		{
			drs.GetValue(_T("deptid"), szTreatDeptID);
			
			bool bFound = false;
			for (int i =0; i < arDepartmentList.GetCount(); i++)
			{
				tmpStr  = arDepartmentList[i];
				if (szTreatDeptID == tmpStr)
				{
					bFound = true;
					break;
				}
			}
			if(!szDepts.IsEmpty())
				szDepts += _T(",  ");
			szDepts.AppendFormat(_T("%s"), szTreatDeptID);

			if (!bFound)
			{
				arDepartmentList.Add(szTreatDeptID);	
			}
			
			drs.MoveNext();
		}

		rpt.GetReportHeader()->SetValue(_T("Department"), szDepts);	

		double nSumCost=0, nSumInsCost=0, nSumDiscount=0, nSumDiffPaid=0, nSumPatPaid=0;

		szDepts.Empty();


		int nChapter = 0;
		int nCount = 0, nIndex, nChapterID = 0;
		int nItem = 0, nOldItem = 0, nItem2 = 0;
		bool bDeleteParent = false;
		bool bLoadItems = false;
		CString szNewGroup, szOldGroup, szParentGroup;

		
	double nGroup0Cost=0, nGroup0InsCost= 0, nGroup0Discount=0, nGroup0DiffPaid=0, nGroup0PatPaid=0;
	double nGroup1Cost=0, nGroup1InsCost = 0, nGroup1Discount=0, nGroup1DiffPaid = 0, nGroup1PatPaid=0;
	double nGroup2Cost=0, nGroup2InsCost = 0, nGroup2Discount=0, nGroup2DiffPaid = 0, nGroup2PatPaid=0;	
	double nCost=0, nInsCost =0,  nDiscount=0, nDiffPaid  = 0, nPatPaid=0;

	
	nTotalCost = nTotalDiffPaid = nTotalInsCost = nTotalPatPaid = nTotalDiscount =0;

	szSQL.Format(_T("SELECT * FROM hms_fee_type ORDER BY hft_id "));
	grs.ExecSQL(szSQL);
	
	szWhere.Format(_T(" AND hfe_invoiceno=%ld"), nInvoiceNo);

	if (szOutPatient != _T("Y")) 
		szWhere.AppendFormat(_T(" and hfe_groupid <>'D0000' "));
	szWhere.AppendFormat(_T(" and hfe_class in('E', 'I','X') "));
	szWhere.AppendFormat(_T(" and (hfe_category <> 'Y' or hfe_category is null) "));

	szWhere.AppendFormat(_T(" and hfe_object in(7) "));

	szSQL.Format(_T(" SELECT hfe_typeindex AS typeindex,") \
_T("   hfe_status         AS status,") \
_T("   hfe_type           AS feetype,") \
_T("   hfe_groupid        AS groupid,") \
_T("   hfe_groupid3       AS groupid3,") \
_T("   hfe_itemid         AS itemid,") \
_T("   hfe_name           AS name,") \
_T("   hfe_unit           AS unit,") \
_T("   hfe_hitech         AS hitech,") \
_T("   hfe_unitprice      AS unitprice,") \
_T("   SUM(hfe_quantity)       AS qty,") \
_T("   SUM(hfe_cost)      AS cost,") \
_T("   SUM(hfe_patpaid)   AS patpaid") \
_T(" FROM hms_fee_invoiceline_view ") \
_T(" WHERE hfe_docno  =%ld %s ") \
_T(" GROUP BY hfe_typeindex,") \
_T("   hfe_status ,") \
_T("   hfe_type ,") \
_T("   hfe_groupid ,") \
_T("   hfe_groupid3 ,") \
_T("   hfe_itemid ,") \
_T("   hfe_name ,") \
_T("   hfe_unit ,") \
_T("   hfe_hitech ,") \
_T("   hfe_unitprice") \
_T(" ORDER BY hfe_typeindex, hfe_groupid3, hfe_name, hfe_unit, hfe_unitprice"), nDocumentNo, szWhere);

	
	nIndex = 1;
	int nSubItem = 1;
	int nIDX;
	nChapterID = 0;
	CArray<FEEITEM, FEEITEM&> feeList;
	FEEITEM fee;
	CString szName;
	CString szNewIndex, szOldIndex;
	CString szIndex;
	bool bFound = false;
	bool bInList = false;
	bool bOutList = false;
	bool bKList = false;
	int nSubIndex = 1;
	
	double nTtlCost=0, nTtlInsCost=0, nTtlDiscount=0, nTtlDiffPaid=0, nTtlPatPaid=0;
	rs.ExecSQL(szSQL);
_fmsg(_T("%s"), szSQL);

	
	int nNumberTreat = 0;
	while(!grs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex);
		grs.GetValue(_T("hft_name"), szName);
		grs.GetValue(_T("hft_id"), szNewIndex);
		
		fee.szGroupID = _T("Type");
		fee.szID = tmpStr;
		fee.szName = szName;
		nItem = feeList.Add(fee);
		
		nIDX  = nItem;
		bFound = false;
		nGroup1Cost=nGroup1InsCost=nGroup1Discount=nGroup1DiffPaid=nGroup1PatPaid=0;
		nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;

		rs.MoveFirst();
		while(!rs.IsEOF())
		{
			rs.GetValue(_T("typeindex"), tmpStr);
			
			if(tmpStr == szNewIndex){
				bFound = true;		
				rs.GetValue(_T("groupid"), szNewGroup);
				if(szNewGroup != szOldGroup)
				{
					szOldGroup = szNewGroup;
					if(szNewGroup.Left(2) == _T("B1")){
						szSQL.Format(_T("SELECT hfg_name FROM hms_fee_group WHERE rtrim(hfg_id,'0')='%s' "), szNewGroup);
						CRecord rs2(&pMF->m_db);
						rs2.ExecSQL(szSQL);
						if(nGroup2Cost > 0){
							TranslateString(_T("Sub Amount"), tmpStr);
							fee.szGroupID = _T("SubAmount");
							fee.szName = tmpStr;
							fee.nCost = nGroup2Cost;
							fee.nInsPaid = nGroup2InsCost;
							fee.nDiscount = nGroup2Discount;
							fee.nDiffPaid = nGroup2DiffPaid;
							fee.nPatPaid = nGroup2PatPaid;
							int nItem2 = feeList.Add(fee);
						}

						szIndex.Format(_T("%d.%d"), nIndex, nSubIndex++);
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = rs2.GetValue(_T("hfg_name"));
						int nItem2 = feeList.Add(fee);
						nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;

					}
					
				}

				if(szOldIndex != szNewIndex)
				{
					szOldIndex = szNewIndex;
					bInList = false;
					bOutList = false;
				}


				nInsCost = 0;
				nDiscount = 0;
				rs.GetValue(_T("cost"), nCost);
				rs.GetValue(_T("patpaid"), nPatPaid);
				
				nTotalCost += nCost;
				nGroup1Cost += nCost;
				nGroup2Cost += nCost;

				nTotalInsCost += nInsCost;
				nGroup1InsCost += nInsCost;
				nGroup2InsCost += nInsCost;

				nTotalDiscount += nDiscount;
				nGroup1Discount += nDiscount;
				nGroup2Discount += nDiscount;

				nTotalDiffPaid += nDiffPaid;
				nGroup1DiffPaid += nDiffPaid;
				nGroup2DiffPaid += nDiffPaid;

				nTotalPatPaid += nPatPaid;
				nGroup1PatPaid += nPatPaid;
				nGroup2PatPaid += nPatPaid;

				fee.szID.Empty();
				fee.szGroupID = _T("Item");
				fee.szName = rs.GetValue(_T("name"));
				fee.szUnit = rs.GetValue(_T("unit"));
				fee.nCost = nCost;
				fee.nInsPaid = nInsCost;
				fee.nDiscount = nDiscount;
				fee.nQuantity = str2float(rs.GetValue(_T("qty")));
				fee.nPrice = str2double(rs.GetValue(_T("unitprice")));
				fee.nInsPrice = 0;
				fee.nDiffPaid = nDiffPaid;
				fee.nPatPaid = nPatPaid;
				feeList.Add(fee);

				if(szNewGroup.Left(1) == _T("C"))
				{
					nNumberTreat += (int) fee.nQuantity;

				}

				if(szNewGroup.Left(2) == _T("B4") || szNewGroup.Left(2) == _T("B5") ){
/*
					CString szItemID;
					rs.GetValue(_T("itemid"), szItemID);
					
					szSQL.Format(_T("SELECT hfe_itemid, hfe_desc, hfe_quantity, hfe_unitprice, hfe_cost, hfe_discount, hfe_diffcost, hfe_patpaid ") \
								_T("FROM hms_fee ") \
								_T("WHERE hfe_docno=%ld and hfe_type='V' and hfe_unit='%s' ORDER BY hfe_itemid DESC "), 
								nDocumentNo, szItemID);
					rsl.ExecSQL(szSQL);
					nCost = nInsCost = nPatPaid  = 0;
					rsl.MoveFirst();
					while(!rsl.IsEOF())
					{
								CString szItemID;
								rsl.GetValue(_T("hfe_itemid"), szItemID);
								
								fee.szID = _T("*");
								fee.szGroupID = _T("Item");
								fee.szName = rsl.GetValue(_T("hfe_desc"));
								fee.nQuantity = str2double(rsl.GetValue(_T("hfe_quantity")));
								fee.nCost = 0;
								fee.nDiscount = 0;
								fee.nPrice = str2double(rsl.GetValue(_T("hfe_unitprice")));

								if(m_szObjectType == _T("S"))
								{
									rsl.GetValue(_T("hfe_patpaid"), nPatPaid);
									fee.nCost = nPatPaid;
									fee.nDiffPaid = 0;
									fee.nPatPaid = nPatPaid;

									nCost = nPatPaid;
									nDiscount =0;
									nInsCost =0;
									feeList.Add(fee);
									nTtlCost += nCost;
									nTtlPatPaid += nPatPaid;
									
								}
								else
								{
									rsl.GetValue(_T("hfe_cost"), nCost);
									rsl.GetValue(_T("hfe_discount"), nDiscount);
									rsl.GetValue(_T("hfe_diffcost"), nDiffPaid);
									//rsl.GetValue(_T("hfe_patpaid"), nPatPaid);
									nPatPaid =0;
									fee.nDiffPaid = nDiffPaid;
									fee.nPatPaid = nPatPaid;
									fee.nCost = nCost;
									fee.nDiscount = nDiscount;
									nInsCost = nCost;
									feeList.Add(fee);
									nTtlDiffPaid += nDiffPaid;
									nTtlCost += nCost;
									nTtlInsCost += nInsCost;
									nTtlDiscount += nDiscount;
									nTtlPatPaid += nPatPaid;
								}
								if(nDiscount +nDiffPaid+nPatPaid > 0)
								{
									nTotalCost += nCost;
									
									nGroup1Cost += nCost;
									nGroup2Cost += nCost;
									nTotalInsCost += nInsCost;
									nGroup1InsCost += nInsCost;
									nGroup2InsCost += nInsCost;
									nTotalDiscount += nDiscount;
									nGroup1Discount += nDiscount;
									nGroup2Discount += nDiscount;
									nTotalDiffPaid += nDiffPaid;
									nGroup1DiffPaid += nDiffPaid;
									nGroup2DiffPaid += nDiffPaid;
									nTotalPatPaid += nPatPaid;
									nGroup1PatPaid += nPatPaid;
									nGroup2PatPaid += nPatPaid;
								}
						rsl.MoveNext();

					}
*/

				}	//end B4, B5



			}
			rs.MoveNext();
		}
		if(!bFound)
		{
			feeList.RemoveAt(nIDX);
			nGroup1Cost=nGroup1InsCost=nGroup1Discount=nGroup1DiffPaid=nGroup1PatPaid=0;
			nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;
		}
		else
		{
			if(nGroup1Cost > 0){
				if(nGroup2Cost > 0 && nGroup1Cost != nGroup2Cost){
					TranslateString(_T("Sub Amount"), tmpStr);
					fee.szGroupID = _T("SubAmount");
					fee.szID.Empty();;
					fee.szName = tmpStr;
					fee.nCost = nGroup2Cost;
					fee.nInsPaid = nGroup2InsCost;
					fee.nDiscount = nGroup2Discount;
					fee.nDiffPaid = nGroup2DiffPaid;
					fee.nPatPaid = nGroup2PatPaid;
					feeList.Add(fee);
				}

					TranslateString(_T("Total"), tmpStr);
					tmpStr.AppendFormat(_T(" (%d)"), nIndex);
					fee.szGroupID = _T("GrandAmount");
					fee.szID.Empty();;
					fee.szName = tmpStr;
					fee.nCost = nGroup1Cost;
					fee.nInsPaid = nGroup1InsCost;
					fee.nDiscount = nGroup1Discount;
					fee.nDiffPaid = nGroup1DiffPaid;
					fee.nPatPaid = nGroup1PatPaid;
					feeList.Add(fee);
			}
			nGroup1Cost=nGroup1InsCost=nGroup1Discount=nGroup1DiffPaid=nGroup1PatPaid=0;
			nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;

			nIndex++;
		}



		grs.MoveNext();
	}

	nTotalCost  = ceil(nTotalCost);
	nTotalInsCost  = ceil(nTotalInsCost);
	nTotalDiscount  = ceil(nTotalDiscount);
	nTotalDiffPaid  = ceil(nTotalDiffPaid);
	nTotalPatPaid  = ceil(nTotalPatPaid);
	nTotalPayable = nTotalCost-nTotalDiscount;
	
	TranslateString(_T("Total Amount"), szName);
	fee.szGroupID = _T("TotalAmount");
	fee.szName.Format(_T("%s "), szName);
	fee.nCost = nTotalCost;
	fee.nInsPaid = nTotalInsCost;
	fee.nDiscount = nTotalDiscount;
	fee.nDiffPaid = nTotalDiffPaid;
	fee.nPatPaid = nTotalPatPaid;
	feeList.Add(fee);
	
	nSumCost += nTotalCost;
	nSumInsCost += nTotalInsCost;
	nSumDiscount += nTotalDiscount;
	nSumDiffPaid += nTotalDiffPaid;
	nSumPatPaid += nTotalPatPaid;


		for (int i =0; i < feeList.GetCount(); i++){
			fee = feeList[i];
			szType = fee.szGroupID;

			if(szType == _T("Type"))
			{
				grp = rpt.GetGroupHeader(1);
				rptDetail = rpt.AddDetail(grp);
				tmpStr = fee.szID;
				rptDetail->SetValue(_T("TypeID"), tmpStr);
				StringUpper(fee.szName, tmpStr);
				rptDetail->SetValue(_T("TypeName"), tmpStr);
			}
			else if(szType == _T("Group")){
				grp = rpt.GetGroupHeader(1);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("TypeID"), fee.szID);
				rptDetail->SetValue(_T("TypeName"), fee.szName);
			}
			else if(szType == _T("SubGroup")){
				grp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("SubGroupID"), fee.szID);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);
				
				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupInsCost"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);
				
				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);
				

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupPatPaid"), tmpStr);


				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("SubGroupInsUnPaid"), tmpStr);
				}
			}
			else if(szType == _T("SubAmount")){
				grp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(grp);
				CReportItem *obj ;
				obj = rptDetail->GetItem(_T("SubGroupName")); if(obj) obj->SetBold(true);
				obj = rptDetail->GetItem(_T("SubGroupCost")); if(obj) obj->SetBold(true);
				obj = rptDetail->GetItem(_T("SubGroupDiscount")); if(obj) obj->SetBold(true);
				obj = rptDetail->GetItem(_T("SubGroupPatpaid")); if(obj) obj->SetBold(true);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);


				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("SubGroupInsUnPaid"), tmpStr);
				}

			}
			else if(szType == _T("GrandAmount")){
				grp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(grp);
			//	rptDetail->GetItem(_T("SubGroupName"))->SetBold(true);
			//	rptDetail->GetItem(_T("SubGroupCost"))->SetBold(true);
			//	rptDetail->GetItem(_T("SubGroupDiscount"))->SetBold(true);
			//	rptDetail->GetItem(_T("SubGroupPatpaid"))->SetBold(true);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("SubGroupInsUnPaid"), tmpStr);
				}

			}

			else if(szType == _T("Item")){
				rptDetail = rpt.AddDetail();
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), fee.szUnit);
				FormatCurrency(fee.nQuantity, tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				
			}
			else if(szType == _T("SubItemGroup")){
				rptDetail = rpt.AddDetail();
				CReportItem *obj = NULL;
				for (int i = 1; i <= 9; i++)
				{
					obj = rptDetail->GetItem(i);
					if(obj) 
					{
						obj->SetItalic(true);
						obj->SetBold(true);
					}

				}
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), _T(""));
				rptDetail->SetValue(_T("4"), _T(""));
				rptDetail->SetValue(_T("5"), _T(""));
				
			}

			else if(szType == _T("SubItem")){
				rptDetail = rpt.AddDetail();
				CReportItem *obj = NULL;
				for (int i = 1; i <= 9; i++)
				{
					obj = rptDetail->GetItem(i);
					if(obj) obj->SetItalic(true);
				}

				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), fee.szUnit);
				FormatCurrency(fee.nQuantity, tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("11"), tmpStr);
				}
				
			}
			else if(szType == _T("SubItemAmount")){
				rptDetail = rpt.AddDetail();
				CReportItem *obj = NULL;
				for (int i = 1; i <= 9; i++)
				{
					obj = rptDetail->GetItem(i);
					if(obj) 
					{
						obj->SetItalic(true);
						obj->SetBold(true);
					}

				}
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), _T(""));
				rptDetail->SetValue(_T("4"), _T(""));
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), _T(""));
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("11"), tmpStr);
				}
				
			}

			else if(szType == _T("Dousage")){
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(4));
				rptDetail->SetValue(_T("usage"), fee.szName);
			}
			else if(szType == _T("TotalAmount")){
				grp = rpt.GetGroupHeader(3);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("TotalAmountLabel"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("TotalAmount"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("TotalInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("TotalDiscount"), tmpStr);
				
				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("TotalDiffPaid"), tmpStr);
				//fee.nPatPaid= fee.nCost-fee.nDiscount;
				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("TotalPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("TotalInsUnPaid"), tmpStr);
				}
			}
		}	
	

	if(nTotalCost != nSumCost){
		CString szName;
		TranslateString(_T("Total Amount"), tmpStr);
		szName.Format(_T("%s"), tmpStr);
			grp = rpt.GetGroupHeader(3);
			rptDetail = rpt.AddDetail(grp);
			rptDetail->SetValue(_T("TotalAmountLabel"), szName);
			FormatCurrency(nSumCost, tmpStr);
			rptDetail->SetValue(_T("TotalAmount"), tmpStr);

			FormatCurrency(nSumInsCost, tmpStr);
			rptDetail->SetValue(_T("TotalInsCost"), tmpStr);

			FormatCurrency(nSumDiscount, tmpStr);
			rptDetail->SetValue(_T("TotalDiscount"), tmpStr);
				
			FormatCurrency(nSumDiffPaid, tmpStr);
			rptDetail->SetValue(_T("TotalDiffPaid"), tmpStr);
			FormatCurrency(nSumPatPaid, tmpStr);
			rptDetail->SetValue(_T("TotalPatPaid"), tmpStr);
	}


	szSQL.Format(_T(" SELECT hfe_amount as hfe_amount,") \
		_T("   hfe_inspaid,") \
		_T("   hfe_discount,") \
		_T("   hfe_patpaid,") \
		_T("   hfe_payment,") \
		_T("   hfe_diffcost,") \
		_T("   hfe_diffpaid,") \
		_T("   hfe_deposit,") \
		_T("   hfe_refund,") \
		_T("   hfe_freeamount") \
		_T(" FROM hms_fee_invoice") \
		_T(" WHERE hfe_invoiceno=%ld and hfe_status='P' "), nInvoiceNo);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF())
	{
		double nTotalDiffPaid;
		double nTotalPayment;
		double nTotalPatPaid;

		rs.GetValue(_T("hfe_amount"), nTotalCost);
		rs.GetValue(_T("hfe_inspaid"), nTotalInsCost);
		rs.GetValue(_T("hfe_discount"), nTotalDiscount);
		rs.GetValue(_T("hfe_patpaid"), nTotalPatPaid);
		rs.GetValue(_T("hfe_diffcost"), nTotalDiffPaid);
		rs.GetValue(_T("hfe_deposit"), nTotalDeposit);
		rs.GetValue(_T("hfe_refund"), nTotalRefund);
		rs.GetValue(_T("hfe_freeamount"), nTotalFree);
		rs.GetValue(_T("hfe_payment"), nTotalPayment);
	}
	else
	{
		double nDepositAmount = 0;
		szSQL.Format(_T("select sum(hfe_amount-hfe_patpaid) as deposit_amount ") \
			_T("from hms_fee_deposit ") \
			_T("where hfe_docno=%ld and hfe_status<>'C' ") \
			_T("and hfe_class IN('E','I','A') and hfe_objectid=7  "), nDocumentNo);
		rs.ExecSQL(szSQL);
		if(!rs.IsEOF())
		{
			rs.GetValue(_T("deposit_amount"), nDepositAmount);
		}
		nTotalDeposit = nDepositAmount;
		if (nTotalCost < nTotalDeposit)
		{
			nTotalRefund = nTotalDeposit - nTotalCost;
			nTotalPayment = 0;
		}
		else
		{
			nTotalPayment = nTotalCost - nTotalDeposit;
			nTotalRefund = 0;
		}
	}

	if (nNumberTreat > 0)
	{
		tmpStr.Format(_T("%d"), nNumberTreat);
		rpt.GetReportHeader()->SetValue(_T("TreatmentDay"), tmpStr);
	}
		rpt.GetReportFooter()->SetValue(_T("PatientSign"), szPatientName);
		FormatCurrency((nTotalCost), tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumAmount"), tmpStr);
		CString szSumInWord;
		tmpStr.Format(_T("%.0f"), nTotalCost);
		MoneyToString(tmpStr, szSumInWord);
		szSumInWord += _T(" \x111\x1ED3ng.");
		rpt.GetReportFooter()->SetValue(_T("SumInWord"), szSumInWord);
		
		FormatCurrency(nTotalInsCost, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumInsCost"), tmpStr);


		FormatCurrency(nTotalDiscount, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumInsDiscount"), tmpStr);

		FormatCurrency(nTotalPatPaid, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumInsUnPaid"), tmpStr);

		FormatCurrency((nTotalDiffPaid), tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumDiffPaid"), tmpStr);


		FormatCurrency(nTotalPatPaid, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumPatPaid"), tmpStr);



		FormatCurrency((nTotalDeposit), tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumDeposit"), tmpStr);
		
		FormatCurrency(nTotalFree, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumDiscount"), tmpStr);

		FormatCurrency(nSumFoodAmount, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumFoodAmount"), tmpStr);

		if(nTotalRefund > 0)
		{
		//	FormatCurrency(nTotalRefund, tmpStr);
		//	rpt.GetReportFooter()->SetValue(_T("SumPayment"), tmpStr);

			FormatCurrency(nTotalRefund, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("SumRefund"), tmpStr);
		}
		else
		{
			FormatCurrency(nTotalPayment, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("SumPayment"), tmpStr);
		}
		
	
	rpt.PrintPreview();
}









void CPrintForms::FM_PrintAllServiceDischargePayment(long nDocumentNo)
{
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
//	if (!pMF->IsInPatient())
//		return;

	CReport rpt;
	CString tmpStr, szSQL, szWhere, szAdmitDate, szEndDate,szDepartments, szAddonedayoutofhospital,szNumbertreat;
	CRecord rs(&pMF->m_db);
	CRecord rsd(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	int		nDeskNo=0;
	CString szRecvDate;
	CString	szObjectType;
	CString	szDescription;
	CString szPrintMaterialOperation = _T("Y");
	CString szOutPatient;

	double	nTotalCost=0,				//Tong chi phi
			nTotalInsCost=0,
			nTotalDiscount = 0,			//Tong Mien giam
			nTotalPatPaid =0,			//Tong benh nhan tra
			nTotalDiffPaid = 0,	//Tong so tra chenh lech
			nTotalPayable=0,
			nTotalDeposit = 0,			//Tong tam ung
			nTotalFree = 0,				//Tong mien phi
			nTotalRefund=0,				//Tong hoan tra
			nDepositPayable=0;			//So tien tra tu tam ung
	
	
	double nSumFoodAmount=0;			


	szWhere.AppendFormat(_T(" and hfe_invoiceno > 0 and hfe_status ='P' "));

	szSQL.Format(_T(" SELECT *") \
				_T(" FROM") \
				_T(" (") \
				_T("  SELECT hd_patientno as patientno,") \
				_T("         hd_docno as docno,") \
				_T("         trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("         hp_birthdate as birthdate,") \
				_T("         hms_getage(trunc_date(hd_admitdate), hp_birthdate) as age,") \
				_T("         extract(year from hp_birthdate) as yearofbirth, ") \
				_T("         Get_Selection_Desc('sys_sex', hp_sex) as sex,") \
				_T("         hp_sex as sex_id,") \
				_T("         hp_ethnic as ethnic,") \
				_T("         Get_Selection_Desc('sys_ethnic', hp_ethnic) as ethnicdesc, ") \
				_T("         hp_dtladdr as detailaddress,") \
				_T("         hms_getaddress(hp_provid,hp_distid, hp_villid) as address,") \
				_T("         hp_workplace as workplace,") \
				_T("         hd_doctor as doctor,") \
				_T("         hd_createdby as createdby,") \
				_T("         hd_icd as icd10,") \
				_T("         hd_diagnostic as diagnostic,") \
				_T("         hd_reldisease as reldisease,") \
				_T("         hd_emergency as emergency, ") \
				_T("         Get_Selection_Desc('hms_result',hd_result) as result,") \
				_T("         hd_admitdate as examdate,") \
				_T("         hd_enddate as enddate,") \
				_T("		hd_outpatient, ") \
				_T("         hcr_recordno as recordno ,") \
				_T("         hcr_admitdate as admitdate,") \
				_T("         hcr_dischargedate as dischargedate,") \
				_T("         hcr_sumtreat as TreatmentDay,") \
				_T("         hcr_mainicd as mainicd ,") \
				_T("         hcr_maindisease as maindisease ,") \
				_T("         hcr_treatdoctor as treatdoctor, ") \
				_T("         hcr_reldisease as ireldisease,") \
				_T("         hcr_suggestion as isuggestion,") \
				_T("         hd_suggestion as esuggestion,") \
				_T("         Get_Selection_Desc('hms_result' ,hcr_result)  as iresult,") \
				_T("         hfe_deptid as deptid,") \
				_T("         hfe_serialno as serialno,") \
				_T("         hfe_receiptno as recvno,") \
				_T("         hfe_date as recvdate, ") \
				_T("         hfe_staff as receiver,") \
				_T("         hfe_desc as reason,") \
				_T("         hfe_amount as cost,") \
				_T("         hfe_discount as discount,") \
				_T("         hfe_patpaid as patpaid,") \
				_T("         hfe_deposit as deposit_amount, ") \
				_T("         0 as deposit_payable,") \
				_T("         hfe_refund as refund_amount, ") \
				_T("         hfe_freeamount as free_amount,") \
				_T("         hfe_eat_amount as food_amount,") \
				_T("         row_number() over (partition by hd_docno order by hd_docno) as dn") \
				_T(" FROM hms_fee_invoice ") \
				_T(" LEFT JOIN hms_doc ON(hd_docno=hfe_docno)") \
				_T(" LEFT JOIN hms_patient ON(hp_patientno=hd_patientno)") \
				_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno) ") \
				_T(" LEFT JOIN hms_treatment_record ON(hfe_docno=htr_docno) ") \
				_T(" LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx) ") \
				_T(" WHERE hfe_docno=%ld and hfe_type<>'E' %s ") \
				_T(" ) Tbl") \
				_T(" WHERE dn=1"), nDocumentNo, szWhere);
	
	
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		return;
	}

	rs.GetValue(_T("cost"), nTotalCost);
	rs.GetValue(_T("discount"), nTotalDiscount);
	rs.GetValue(_T("patpaid"), nTotalPatPaid);
	rs.GetValue(_T("deposit_amount"), nTotalDeposit);
	rs.GetValue(_T("free_amount"), nTotalFree);
	rs.GetValue(_T("refund_amount"), nTotalRefund);
	rs.GetValue(_T("deposit_payable"), nDepositPayable);
	rs.GetValue(_T("reason"), szDescription);
	rs.GetValue(_T("deskno"), nDeskNo);
	
	rs.GetValue(_T("food_amount"), nSumFoodAmount);
	rs.GetValue(_T("hd_outpatient"), szOutPatient);



	if(!rpt.Init(_T("Reports/HMS/HF_DISCHARGEPAYMENTALL.RPT"), false) )
			return;

	//Report Header
	tmpStr.Empty();
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);

	rs.GetValue(_T("docno"), nDocumentNo);
	tmpStr.Format(_T("%ld"), nDocumentNo);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	tmpStr.Format(_T("%s-%s-%s"), rs.GetValue(_T("serialno")), rs.GetValue(_T("bookno")),  rs.GetValue(_T("recvno")));
	rpt.GetReportHeader()->SetValue(_T("ReceiptNo"), tmpStr);
//	tmpStr.Format(_T("%ld"), nInvoiceNo);
//	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);


	CString szDate;
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")), tmpStr.Mid(11, 5),tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), szDate);


	rs.GetValue(_T("recordno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RecordNo"), tmpStr);
	//rs.GetValue(_T("pname"), tmpStr);
	

	CString szPatientName;
	StringUpper(rs.GetValue(_T("pname")), szPatientName);
	rpt.GetReportHeader()->SetValue(_T("PATIENTNAME"), szPatientName);
	rs.GetValue(_T("age"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), tmpStr);
	rs.GetValue(_T("yearofbirth"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);

	rs.GetValue(_T("birthdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("BirthDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

	rs.GetValue(_T("sex"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Sex"), tmpStr);

	rs.GetValue(_T("sex_id"), tmpStr);
	if(tmpStr == _T("F"))
		rpt.GetReportHeader()->SetValue(_T("Female"), _T("X"));
	else
		rpt.GetReportHeader()->SetValue(_T("Male"), _T("X"));


	rs.GetValue(_T("ethnic"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ethnic"), tmpStr);
	rs.GetValue(_T("ethnicdesc"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ethnicdesc"), tmpStr);
	
	rs.GetValue(_T("workplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Workplace"), tmpStr);

	CString szAddress;
	rs.GetValue(_T("detailaddress"), tmpStr);
	szAddress = tmpStr;
	rs.GetValue(_T("address"), tmpStr);
	if(!szAddress.IsEmpty())
		szAddress += _T(" - ");
	szAddress += tmpStr;
	rpt.GetReportHeader()->SetValue(_T("Address"), szAddress);
	

	rs.GetValue(_T("transplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("TransferHospital"), tmpStr);
	
	
	rs.GetValue(_T("emergency"), tmpStr);
	if (tmpStr == _T("Y") )
		tmpStr = _T("X");
	else
		tmpStr.Empty();

	rpt.GetReportHeader()->SetValue(_T("Emergency"), tmpStr);

	rs.GetValue(_T("maindisease"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);
		rs.GetValue(_T("ireldisease"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("RelationDisease"), tmpStr);
		rs.GetValue(_T("mainicd"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("ICD10"), tmpStr);
		rs.GetValue(_T("iresult"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Result"), tmpStr);
		rs.GetValue(_T("isuggestion"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Suggestion"), pMF->GetSelectionString(_T("hms_suggestion"), tmpStr));
	
	
	
	CString szDischargeDate;
	rs.GetValue(_T("admitdate"), tmpStr);

	rpt.GetReportHeader()->SetValue(_T("admitdate"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	
	rs.GetValue(_T("dischargedate"), tmpStr);
	szDischargeDate = tmpStr;
	
	if(szDischargeDate.Left(4) != _T("1752"))
			rpt.GetReportHeader()->SetValue(_T("dischargedate"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	else
			rpt.GetReportHeader()->SetValue(_T("dischargedate"), _T(""));

	rs.GetValue(_T("TreatmentDay"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("TreatmentDay"), tmpStr);	

	CString szData;
	tmpStr.Format(_T("%s-%s-%s"), rs.GetValue(_T("bookno")), rs.GetValue(_T("serialno")), rs.GetValue(_T("recvno")));
	rpt.GetReportFooter()->SetValue(_T("SerialNo"), tmpStr);
	rs.GetValue(_T("recvdate"), szRecvDate);	
	rpt.GetReportFooter()->SetValue(_T("ReceiptDate"), CDateTime::Convert(szRecvDate, yyyymmdd|hhmm,ddmmyyyy|hhmm));
	szData.Format(rpt.GetReportFooter()->GetValue(_T("ReceiptDate1")), szRecvDate.Mid(11,5),szRecvDate.Mid(8,2),szRecvDate.Mid(5,2),szRecvDate.Left(4));	
	rpt.GetReportFooter()->SetValue(_T("ReceiptDate1"),szData);
	szData.Format(rpt.GetReportFooter()->GetValue(_T("ReceiptDate2")), szRecvDate.Mid(11,5),szRecvDate.Mid(8,2),szRecvDate.Mid(5,2),szRecvDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("ReceiptDate2"),szData);
	
	rs.GetValue(_T("treatdoctor"), tmpStr);
	tmpStr.Trim();
	rpt.GetReportFooter()->SetValue(_T("Doctor"), GetDoctorName(tmpStr));

	CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
	if(img)
	{
		img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img->SetFixedHeight(true);
	}
	tmpStr.Empty();
	rs.GetValue(_T("receiver"), tmpStr);
	tmpStr.Trim();
	rpt.GetReportFooter()->SetValue(_T("ReceiverBy"), GetDoctorName(tmpStr));
	CReportItem *img2 = rpt.GetReportFooter()->GetItem(_T("Signature2"));
	if(img2)
	{
		img2->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img2->SetFixedHeight(false);
	}	
	
	
	// Nguoi tiep don
	tmpStr.Empty();
	rs.GetValue(_T("createdby"), tmpStr);
	tmpStr.Trim();
	rpt.GetReportFooter()->SetValue(_T("createdby"), GetDoctorName(tmpStr));
	CReportItem *img3 = rpt.GetReportFooter()->GetItem(_T("Signature3"));
	if(img3)
	{
		img3->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img3->SetFixedHeight(false);
	}	


	CReportSection* rptDetail=NULL;
	CReportSection *grp;
	CRecord grs(&pMF->m_db);
	CString szFeeID, szID;
	CString szSubItem, szType;

	TCHAR *lpszChapter[] = {_T("I"), _T("II"), _T("III"), _T("IV"), _T("V"), _T("VI"), _T("VII"), _T("VIII"), _T("IX"), _T("X"), _T("XI")};


	CRecord drs(&pMF->m_db);
	CString szDeptID;
	CStringArray arDepartmentList;

	szSQL.Format(_T(" SELECT deptid, admitdate, dischargedate, mainicd, maindisease") \
					_T(" FROM") \
					_T(" (") \
					_T("  select htr_deptid as deptid,") \
					_T("         htr_admitdate as admitdate,") \
					_T("         htr_dischargedate as dischargedate,") \
					_T("         htr_mainicd as mainicd,") \
					_T("         htr_maindisease as maindisease,") \
					_T("         htr_idx as idx") \
					_T("  from hms_treatment_record") \
					_T("  where htr_docno=%ld") \
					_T("  union all") \
					_T("  select hd_admitdept,") \
					_T("         hd_admitdate,") \
					_T("         hd_enddate,") \
					_T("         hd_icd,") \
					_T("         hd_diagnostic,") \
					_T("         0") \
					_T("  from hms_doc") \
					_T("  where hd_docno=%ld") \
					_T(" ) tbl") \
					_T(" WHERE deptid in(select distinct hfe_deptid from hms_fee where hfe_invoiceno > 0 and hfe_status ='P' )") \
					_T(" ORDER BY idx"), nDocumentNo, nDocumentNo);

		drs.ExecSQL(szSQL);

		RECEIPTINFO pi;
		CString szDepts;
		szDepts.Empty();
		CString szTreatDeptID;
		while (!drs.IsEOF())
		{
			drs.GetValue(_T("deptid"), szTreatDeptID);
			
			bool bFound = false;
			for (int i =0; i < arDepartmentList.GetCount(); i++)
			{
				tmpStr  = arDepartmentList[i];
				if (szTreatDeptID == tmpStr)
				{
					bFound = true;
					break;
				}
			}
			if(!szDepts.IsEmpty())
				szDepts += _T(",  ");
			szDepts.AppendFormat(_T("%s"), szTreatDeptID);

			if (!bFound)
			{
				arDepartmentList.Add(szTreatDeptID);	
			}
			
			drs.MoveNext();
		}

		rpt.GetReportHeader()->SetValue(_T("Department"), szDepts);	

		double nSumCost=0, nSumInsCost=0, nSumDiscount=0, nSumDiffPaid=0, nSumPatPaid=0;

		szDepts.Empty();


		int nChapter = 0;
		int nCount = 0, nIndex, nChapterID = 0;
		int nItem = 0, nOldItem = 0, nItem2 = 0;
		bool bDeleteParent = false;
		bool bLoadItems = false;
		CString szNewGroup, szOldGroup, szParentGroup;

		
	double nGroup0Cost=0, nGroup0InsCost= 0, nGroup0Discount=0, nGroup0DiffPaid=0, nGroup0PatPaid=0;
	double nGroup1Cost=0, nGroup1InsCost = 0, nGroup1Discount=0, nGroup1DiffPaid = 0, nGroup1PatPaid=0;
	double nGroup2Cost=0, nGroup2InsCost = 0, nGroup2Discount=0, nGroup2DiffPaid = 0, nGroup2PatPaid=0;	
	double nCost=0, nInsCost =0,  nDiscount=0, nDiffPaid  = 0, nPatPaid=0;

	
	nTotalCost = nTotalDiffPaid = nTotalInsCost = nTotalPatPaid = nTotalDiscount =0;

	szSQL.Format(_T("SELECT * FROM hms_fee_type ORDER BY hft_id "));
	grs.ExecSQL(szSQL);
	
	szWhere.Format(_T(" AND hfe_invoiceno > 0 and hfe_status='P' "));

	if (szOutPatient != _T("Y")) 
		szWhere.AppendFormat(_T(" and hfe_groupid <>'D0000' "));
	szWhere.AppendFormat(_T(" and hfe_class in('E', 'I','X') "));
	//szWhere.AppendFormat(_T(" and hfe_category<>'Y' "));
	szWhere.AppendFormat(_T(" and (hfe_category <> 'Y' or hfe_category is null) "));


	szSQL.Format(_T(" SELECT hfe_typeindex AS typeindex,") \
_T("   hfe_status         AS status,") \
_T("   hfe_type           AS feetype,") \
_T("   hfe_groupid        AS groupid,") \
_T("   hfe_groupid3       AS groupid3,") \
_T("   hfe_itemid         AS itemid,") \
_T("   hfe_name           AS name,") \
_T("   hfe_unit           AS unit,") \
_T("   hfe_hitech         AS hitech,") \
_T("   hfe_inlist         AS inlist,") \
_T("   hfe_unitprice      AS unitprice,") \
_T("   SUM(hfe_quantity)       AS qty,") \
_T("   SUM(hfe_amount)      AS cost,") \
_T("   SUM(hfe_discount)  AS discount,") \
_T("   SUM(hfe_diffpaid)  AS DiffPaid,") \
_T("   SUM(hfe_copayment) AS copayment,") \
_T("   SUM(hfe_patpaid)   AS patpaid") \
_T(" FROM hms_fee_invoiceline_view ") \
_T(" WHERE hfe_docno  =%ld %s ") \
_T(" GROUP BY hfe_typeindex,") \
_T("   hfe_status ,") \
_T("   hfe_type ,") \
_T("   hfe_groupid ,") \
_T("   hfe_groupid3 ,") \
_T("   hfe_itemid ,") \
_T("   hfe_name ,") \
_T("   hfe_unit ,") \
_T("   hfe_hitech ,") \
_T("   hfe_inlist,") \
_T("   hfe_unitprice") \
_T(" ORDER BY hfe_typeindex, hfe_inlist, hfe_groupid3, hfe_name, hfe_unit, hfe_unitprice"), nDocumentNo, szWhere);



_fmsg(_T("%s"), szSQL);		
	
	nIndex = 1;
	int nSubItem = 1;
	int nIDX;
	nChapterID = 0;
	CArray<FEEITEM, FEEITEM&> feeList;
	FEEITEM fee;
	CString szName;
	CString szNewIndex, szOldIndex;
	CString szIndex;
	bool bFound = false;
	bool bInList = false;
	bool bOutList = false;
	bool bKList = false;
	int nSubIndex = 1;
	
	double nTtlCost=0, nTtlInsCost=0, nTtlDiscount=0, nTtlDiffPaid=0, nTtlPatPaid=0;
	rs.ExecSQL(szSQL);
_fmsg(_T("%s"), szSQL);

	

	while(!grs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex);
		grs.GetValue(_T("hft_name"), szName);
		grs.GetValue(_T("hft_id"), szNewIndex);
		
		fee.szGroupID = _T("Type");
		fee.szID = tmpStr;
		fee.szName = szName;
		nItem = feeList.Add(fee);
		
		nIDX  = nItem;
		bFound = false;
		nGroup1Cost=nGroup1InsCost=nGroup1Discount=nGroup1DiffPaid=nGroup1PatPaid=0;
		nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;

		rs.MoveFirst();
		while(!rs.IsEOF())
		{
			rs.GetValue(_T("typeindex"), tmpStr);
			
			if(tmpStr == szNewIndex){
				bFound = true;		
				rs.GetValue(_T("groupid"), szNewGroup);
				if(szNewGroup != szOldGroup)
				{
					szOldGroup = szNewGroup;
					if(szNewGroup.Left(2) == _T("B1")){
						szSQL.Format(_T("SELECT hfg_name FROM hms_fee_group WHERE rtrim(hfg_id,'0')='%s' "), szNewGroup);
						CRecord rs2(&pMF->m_db);
						rs2.ExecSQL(szSQL);
						if(nGroup2Cost > 0){
							TranslateString(_T("Sub Amount"), tmpStr);
							fee.szGroupID = _T("SubAmount");
							fee.szName = tmpStr;
							fee.nCost = nGroup2Cost;
							fee.nInsPaid = nGroup2InsCost;
							fee.nDiscount = nGroup2Discount;
							fee.nDiffPaid = nGroup2DiffPaid;
							fee.nPatPaid = nGroup2PatPaid;
							int nItem2 = feeList.Add(fee);
						}

						szIndex.Format(_T("%d.%d"), nIndex, nSubIndex++);
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = rs2.GetValue(_T("hfg_name"));
						int nItem2 = feeList.Add(fee);
						nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;

					}
					
				}

				if(szOldIndex != szNewIndex)
				{
					szOldIndex = szNewIndex;
					bInList = false;
					bOutList = false;
				}



				rs.GetValue(_T("cost"), nCost);
				rs.GetValue(_T("inscost"), nInsCost);
				rs.GetValue(_T("discount"), nDiscount);
				rs.GetValue(_T("DiffPaid"), nDiffPaid);
				rs.GetValue(_T("copayment"), nPatPaid);
				
				nTotalCost += nCost;
				nGroup1Cost += nCost;
				nGroup2Cost += nCost;

				nTotalInsCost += nInsCost;
				nGroup1InsCost += nInsCost;
				nGroup2InsCost += nInsCost;

				nTotalDiscount += nDiscount;
				nGroup1Discount += nDiscount;
				nGroup2Discount += nDiscount;

				nTotalDiffPaid += nDiffPaid;
				nGroup1DiffPaid += nDiffPaid;
				nGroup2DiffPaid += nDiffPaid;

				nTotalPatPaid += nPatPaid;
				nGroup1PatPaid += nPatPaid;
				nGroup2PatPaid += nPatPaid;

				fee.szID.Empty();
				fee.szGroupID = _T("Item");
				fee.szName = rs.GetValue(_T("name"));
				fee.szUnit = rs.GetValue(_T("unit"));
				fee.nCost = nCost;
				fee.nInsPaid = nInsCost;
				fee.nDiscount = nDiscount;
				fee.nQuantity = str2float(rs.GetValue(_T("qty")));
				fee.nPrice = str2double(rs.GetValue(_T("unitprice")));
				fee.nInsPrice = str2double(rs.GetValue(_T("insprice")));
				fee.nDiffPaid = nDiffPaid;
				fee.nPatPaid = nPatPaid;
				feeList.Add(fee);

				if(szNewGroup.Left(2) == _T("B4") || szNewGroup.Left(2) == _T("B5") ){
/*
					CString szItemID;
					rs.GetValue(_T("itemid"), szItemID);
					
					szSQL.Format(_T("SELECT hfe_itemid, hfe_desc, hfe_quantity, hfe_unitprice, hfe_cost, hfe_discount, hfe_diffcost, hfe_patpaid ") \
								_T("FROM hms_fee ") \
								_T("WHERE hfe_docno=%ld and hfe_type='V' and hfe_unit='%s' ORDER BY hfe_itemid DESC "), 
								nDocumentNo, szItemID);
					rsl.ExecSQL(szSQL);
					nCost = nInsCost = nPatPaid  = 0;
					rsl.MoveFirst();
					while(!rsl.IsEOF())
					{
								CString szItemID;
								rsl.GetValue(_T("hfe_itemid"), szItemID);
								
								fee.szID = _T("*");
								fee.szGroupID = _T("Item");
								fee.szName = rsl.GetValue(_T("hfe_desc"));
								fee.nQuantity = str2double(rsl.GetValue(_T("hfe_quantity")));
								fee.nCost = 0;
								fee.nDiscount = 0;
								fee.nPrice = str2double(rsl.GetValue(_T("hfe_unitprice")));

								if(m_szObjectType == _T("S"))
								{
									rsl.GetValue(_T("hfe_patpaid"), nPatPaid);
									fee.nCost = nPatPaid;
									fee.nDiffPaid = 0;
									fee.nPatPaid = nPatPaid;

									nCost = nPatPaid;
									nDiscount =0;
									nInsCost =0;
									feeList.Add(fee);
									nTtlCost += nCost;
									nTtlPatPaid += nPatPaid;
									
								}
								else
								{
									rsl.GetValue(_T("hfe_cost"), nCost);
									rsl.GetValue(_T("hfe_discount"), nDiscount);
									rsl.GetValue(_T("hfe_diffcost"), nDiffPaid);
									//rsl.GetValue(_T("hfe_patpaid"), nPatPaid);
									nPatPaid =0;
									fee.nDiffPaid = nDiffPaid;
									fee.nPatPaid = nPatPaid;
									fee.nCost = nCost;
									fee.nDiscount = nDiscount;
									nInsCost = nCost;
									feeList.Add(fee);
									nTtlDiffPaid += nDiffPaid;
									nTtlCost += nCost;
									nTtlInsCost += nInsCost;
									nTtlDiscount += nDiscount;
									nTtlPatPaid += nPatPaid;
								}
								if(nDiscount +nDiffPaid+nPatPaid > 0)
								{
									nTotalCost += nCost;
									
									nGroup1Cost += nCost;
									nGroup2Cost += nCost;
									nTotalInsCost += nInsCost;
									nGroup1InsCost += nInsCost;
									nGroup2InsCost += nInsCost;
									nTotalDiscount += nDiscount;
									nGroup1Discount += nDiscount;
									nGroup2Discount += nDiscount;
									nTotalDiffPaid += nDiffPaid;
									nGroup1DiffPaid += nDiffPaid;
									nGroup2DiffPaid += nDiffPaid;
									nTotalPatPaid += nPatPaid;
									nGroup1PatPaid += nPatPaid;
									nGroup2PatPaid += nPatPaid;
								}
						rsl.MoveNext();

					}
*/

				}	//end B4, B5



			}
			rs.MoveNext();
		}
		if(!bFound)
		{
			feeList.RemoveAt(nIDX);
			nGroup1Cost=nGroup1InsCost=nGroup1Discount=nGroup1DiffPaid=nGroup1PatPaid=0;
			nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;
		}
		else
		{
			if(nGroup1Cost > 0){
				if(nGroup2Cost > 0 && nGroup1Cost != nGroup2Cost){
					TranslateString(_T("Sub Amount"), tmpStr);
					fee.szGroupID = _T("SubAmount");
					fee.szID.Empty();;
					fee.szName = tmpStr;
					fee.nCost = nGroup2Cost;
					fee.nInsPaid = nGroup2InsCost;
					fee.nDiscount = nGroup2Discount;
					fee.nDiffPaid = nGroup2DiffPaid;
					fee.nPatPaid = nGroup2PatPaid;
					feeList.Add(fee);
				}

					TranslateString(_T("Total"), tmpStr);
					tmpStr.AppendFormat(_T(" (%d)"), nIndex);
					fee.szGroupID = _T("GrandAmount");
					fee.szID.Empty();;
					fee.szName = tmpStr;
					fee.nCost = nGroup1Cost;
					fee.nInsPaid = nGroup1InsCost;
					fee.nDiscount = nGroup1Discount;
					fee.nDiffPaid = nGroup1DiffPaid;
					fee.nPatPaid = nGroup1PatPaid;
					feeList.Add(fee);
			}
			nGroup1Cost=nGroup1InsCost=nGroup1Discount=nGroup1DiffPaid=nGroup1PatPaid=0;
			nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;

			nIndex++;
		}



		grs.MoveNext();
	}

	nTotalCost  = ceil(nTotalCost);
	nTotalInsCost  = ceil(nTotalInsCost);
	nTotalDiscount  = ceil(nTotalDiscount);
	nTotalDiffPaid  = ceil(nTotalDiffPaid);
	nTotalPatPaid  = ceil(nTotalPatPaid);
	nTotalPayable = nTotalCost-nTotalDiscount;
	nTotalDiscount = nTotalDiscount;

	TranslateString(_T("Total Amount"), szName);
	fee.szGroupID = _T("TotalAmount");
	fee.szName.Format(_T("%s "), szName);
	fee.nCost = nTotalCost;
	fee.nInsPaid = nTotalInsCost;
	fee.nDiscount = nTotalDiscount;
	fee.nDiffPaid = nTotalDiffPaid;
	fee.nPatPaid = nTotalPatPaid;
	feeList.Add(fee);
	
	nSumCost += nTotalCost;
	nSumInsCost += nTotalInsCost;
	nSumDiscount += nTotalDiscount;
	nSumDiffPaid += nTotalDiffPaid;
	nSumPatPaid += nTotalPatPaid;


		for (int i =0; i < feeList.GetCount(); i++){
			fee = feeList[i];
			szType = fee.szGroupID;

			if(szType == _T("Type"))
			{
				grp = rpt.GetGroupHeader(1);
				rptDetail = rpt.AddDetail(grp);
				tmpStr = fee.szID;
				rptDetail->SetValue(_T("TypeID"), tmpStr);
				StringUpper(fee.szName, tmpStr);
				rptDetail->SetValue(_T("TypeName"), tmpStr);
			}
			else if(szType == _T("Group")){
				grp = rpt.GetGroupHeader(1);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("TypeID"), fee.szID);
				rptDetail->SetValue(_T("TypeName"), fee.szName);
			}
			else if(szType == _T("SubGroup")){
				grp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("SubGroupID"), fee.szID);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);
				
				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupInsCost"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);
				
				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);
				

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupPatPaid"), tmpStr);


				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("SubGroupInsUnPaid"), tmpStr);
				}
			}
			else if(szType == _T("SubAmount")){
				grp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(grp);
				CReportItem *obj ;
				obj = rptDetail->GetItem(_T("SubGroupName")); if(obj) obj->SetBold(true);
				obj = rptDetail->GetItem(_T("SubGroupCost")); if(obj) obj->SetBold(true);
				obj = rptDetail->GetItem(_T("SubGroupDiscount")); if(obj) obj->SetBold(true);
				obj = rptDetail->GetItem(_T("SubGroupPatpaid")); if(obj) obj->SetBold(true);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);


				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("SubGroupInsUnPaid"), tmpStr);
				}

			}
			else if(szType == _T("GrandAmount")){
				grp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(grp);
			//	rptDetail->GetItem(_T("SubGroupName"))->SetBold(true);
			//	rptDetail->GetItem(_T("SubGroupCost"))->SetBold(true);
			//	rptDetail->GetItem(_T("SubGroupDiscount"))->SetBold(true);
			//	rptDetail->GetItem(_T("SubGroupPatpaid"))->SetBold(true);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("SubGroupInsUnPaid"), tmpStr);
				}

			}

			else if(szType == _T("Item")){
				rptDetail = rpt.AddDetail();
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), fee.szUnit);
				FormatCurrency(fee.nQuantity, tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("11"), tmpStr);
				}
				
			}
			else if(szType == _T("SubItemGroup")){
				rptDetail = rpt.AddDetail();
				CReportItem *obj = NULL;
				for (int i = 1; i <= 9; i++)
				{
					obj = rptDetail->GetItem(i);
					if(obj) 
					{
						obj->SetItalic(true);
						obj->SetBold(true);
					}

				}
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), _T(""));
				rptDetail->SetValue(_T("4"), _T(""));
				rptDetail->SetValue(_T("5"), _T(""));
				
			}

			else if(szType == _T("SubItem")){
				rptDetail = rpt.AddDetail();
				CReportItem *obj = NULL;
				for (int i = 1; i <= 9; i++)
				{
					obj = rptDetail->GetItem(i);
					if(obj) obj->SetItalic(true);
				}

				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), fee.szUnit);
				FormatCurrency(fee.nQuantity, tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("11"), tmpStr);
				}
				
			}
			else if(szType == _T("SubItemAmount")){
				rptDetail = rpt.AddDetail();
				CReportItem *obj = NULL;
				for (int i = 1; i <= 9; i++)
				{
					obj = rptDetail->GetItem(i);
					if(obj) 
					{
						obj->SetItalic(true);
						obj->SetBold(true);
					}

				}
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), _T(""));
				rptDetail->SetValue(_T("4"), _T(""));
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), _T(""));
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("11"), tmpStr);
				}
				
			}

			else if(szType == _T("Dousage")){
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(4));
				rptDetail->SetValue(_T("usage"), fee.szName);
			}
			else if(szType == _T("TotalAmount")){
				grp = rpt.GetGroupHeader(3);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("TotalAmountLabel"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("TotalAmount"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("TotalInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("TotalDiscount"), tmpStr);
				
				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("TotalDiffPaid"), tmpStr);
				//fee.nPatPaid= fee.nCost-fee.nDiscount;
				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("TotalPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("TotalInsUnPaid"), tmpStr);
				}
			}
		}	
	

	if(nTotalCost != nSumCost){
		CString szName;
		TranslateString(_T("Total Amount"), tmpStr);
		szName.Format(_T("%s"), tmpStr);
			grp = rpt.GetGroupHeader(3);
			rptDetail = rpt.AddDetail(grp);
			rptDetail->SetValue(_T("TotalAmountLabel"), szName);
			FormatCurrency(nSumCost, tmpStr);
			rptDetail->SetValue(_T("TotalAmount"), tmpStr);

			FormatCurrency(nSumInsCost, tmpStr);
			rptDetail->SetValue(_T("TotalInsCost"), tmpStr);

			FormatCurrency(nSumDiscount, tmpStr);
			rptDetail->SetValue(_T("TotalDiscount"), tmpStr);
				
			FormatCurrency(nSumDiffPaid, tmpStr);
			rptDetail->SetValue(_T("TotalDiffPaid"), tmpStr);
			FormatCurrency(nSumPatPaid, tmpStr);
			rptDetail->SetValue(_T("TotalPatPaid"), tmpStr);
	}

	szSQL.Format(_T(" SELECT coalesce(sum(hfe_amount),0) as amount,") \
_T("   coalesce(sum(hfe_inspaid), 0) as inspaid,") \
_T("   coalesce(sum(hfe_discount), 0) as discount,") \
_T("   coalesce(sum(hfe_patpaid), 0) as patpaid,") \
_T("   coalesce(sum(hfe_payment), 0) as payment,") \
_T("   coalesce(sum(hfe_diffcost), 0) as diffcost,") \
_T("   coalesce(sum(hfe_deposit), 0) as deposit,") \
_T("   coalesce(sum(hfe_refund), 0) as refund,") \
_T("   coalesce(sum(hfe_freeamount), 0) as freeamount ") \
_T(" FROM hms_fee_invoice") \
_T(" WHERE hfe_docno=%ld and hfe_invoiceno > 0 and hfe_status='P' "), nDocumentNo);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF())
	{
		double nTotalDiffPaid;
		double nTotalPayment;
		double nTotalPatPaid;

		rs.GetValue(_T("amount"), nTotalCost);
		rs.GetValue(_T("inspaid"), nTotalInsCost);
		rs.GetValue(_T("discount"), nTotalDiscount);
		rs.GetValue(_T("patpaid"), nTotalPatPaid);
		rs.GetValue(_T("diffcost"), nTotalDiffPaid);
		rs.GetValue(_T("deposit"), nTotalDeposit);
		rs.GetValue(_T("refund"), nTotalRefund);
		rs.GetValue(_T("freeamount"), nTotalFree);
		rs.GetValue(_T("payment"), nTotalPayment);

		


		rpt.GetReportFooter()->SetValue(_T("PatientSign"), szPatientName);
		FormatCurrency((nTotalCost), tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumAmount"), tmpStr);
		CString szSumInWord;
		tmpStr.Format(_T("%.0f"), nTotalCost);
		MoneyToString(tmpStr, szSumInWord);
		szSumInWord += _T(" \x111\x1ED3ng.");
		rpt.GetReportFooter()->SetValue(_T("SumInWord"), szSumInWord);
		
		FormatCurrency(nTotalInsCost, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumInsCost"), tmpStr);


		FormatCurrency(nTotalDiscount, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumInsDiscount"), tmpStr);

		FormatCurrency(nTotalPatPaid, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumInsUnPaid"), tmpStr);

		FormatCurrency((nTotalDiffPaid), tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumDiffPaid"), tmpStr);


		FormatCurrency(nTotalPatPaid, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumPatPaid"), tmpStr);



		FormatCurrency((nTotalDeposit), tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumDeposit"), tmpStr);
		
		FormatCurrency(nTotalFree, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumDiscount"), tmpStr);

		FormatCurrency(nSumFoodAmount, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumFoodAmount"), tmpStr);

		if(nTotalRefund > 0)
		{
		//	FormatCurrency(nTotalRefund, tmpStr);
		//	rpt.GetReportFooter()->SetValue(_T("SumPayment"), tmpStr);

			FormatCurrency(nTotalRefund, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("SumRefund"), tmpStr);
		}
		else
		{
			FormatCurrency(nTotalPayment, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("SumPayment"), tmpStr);

		}
		
	}
	rpt.PrintPreview();
}





void CPrintForms::FM_PrintInsuranceDischargePaymentInvoice(long nDocumentNo, long nInvoiceNo)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
//	if (!pMF->IsInPatient())
//		return;

	CReport rpt;
	CString tmpStr, szSQL, szWhere, szAdmitDate, szEndDate,szDepartments, szAddonedayoutofhospital,szNumbertreat;
	CRecord rs(&pMF->m_db);
	CRecord rsd(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	CRecord rsx(&pMF->m_db);
	int		nDeskNo=0;
	CString szRecvDate;
	CString	szObjectType;
	CString	szDescription;
	CString szPrintMaterialOperation = _T("Y");
	CString	szOutPatient;


	double	nTotalCost=0,				//Tong chi phi
			nTotalInsCost=0,
			nTotalDiscount = 0,			//Tong Mien giam
			nTotalPatPaid =0,			//Tong benh nhan tra
			nTotalDiffPaid = 0,	//Tong so tra chenh lech
			nTotalPayable=0,
			nTotalDeposit = 0,			//Tong tam ung
			nTotalFree = 0,				//Tong mien phi
			nTotalRefund=0,				//Tong hoan tra
			nDepositPayable=0,			//So tien tra tu tam ung
			nTotalPayment = 0;
	
	
	double nSumFoodAmount=0;			


	szWhere.AppendFormat(_T(" and hfe_invoiceno=%ld "), nInvoiceNo);

	szSQL.Format(_T(" SELECT *") \
				_T(" FROM") \
				_T(" (") \
				_T("  SELECT hd_patientno as patientno,") \
				_T("         hd_docno as docno,") \
				_T("         trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("         hp_birthdate as birthdate,") \
				_T("         hms_getage(trunc_date(hd_admitdate), hp_birthdate) as age,") \
				_T("         extract(year from hp_birthdate) as yearofbirth, ") \
				_T("         Get_Selection_Desc('sys_sex', hp_sex) as sex,") \
				_T("         hp_sex as sex_id,") \
				_T("         hp_ethnic as ethnic,") \
				_T("         Get_Selection_Desc('sys_ethnic', hp_ethnic) as ethnicdesc, ") \
				_T("         hp_dtladdr as detailaddress,") \
				_T("         hms_getaddress(hp_provid,hp_distid, hp_villid) as address,") \
				_T("         hp_workplace as workplace,") \
				_T("		 hd_rank as rank, ") \
				_T("		 get_selection_desc('hms_rank', hd_rank) as rankname, ") \
				_T("         hd_doctor as doctor,") \
				_T("         hd_createdby as createdby,") \
				_T("         hd_icd as icd10,") \
				_T("         hd_diagnostic as diagnostic,") \
				_T("         hd_reldisease as reldisease,") \
				_T("         hd_xobject as xobject, ") \
				_T("         hd_xcardno as xcardno, ") \
				_T("         hd_xissuedate as xissuedate, ") \
				_T("         hd_emergency as emergency, ") \
				_T("         Get_Selection_Desc('hms_result',hd_result) as result,") \
				_T("         hd_admitdate as examdate,") \
				_T("         hd_enddate as enddate,") \
				_T("	hd_outpatient, ") \
				_T("         hcr_recordno as recordno ,") \
				_T("         htr_admitdate as admitdate,") \
				_T("         htr_dischargedate as dischargedate,") \
				_T("         hcr_sumtreat as TreatmentDay,") \
				_T("         hcr_mainicd as mainicd ,") \
				_T("         hcr_maindisease as maindisease ,") \
				_T("         hcr_treatdoctor as treatdoctor, ") \
				_T("         hcr_reldisease as ireldisease,") \
				_T("         hcr_suggestion as isuggestion,") \
				_T("         hd_suggestion as esuggestion,") \
				_T("         Get_Selection_Desc('hms_result' ,hcr_result)  as iresult,") \
				_T("         Hms_GetObjectType(hd_object) as objecttype, ") \
				_T("         cast(hd_disrate as integer) as disrate, ") \
				_T("         hd_insline as insline, ") \
				_T("         hd_insregdate as insregdate, ") \
				_T("         hd_transplace as transplace, ") \
				_T("         hc_cardno as cardno, ") \
				_T("         hc_regdate as regdate, ") \
				_T("         hc_expdate as expdate, ") \
				_T("         hc_regcode as regcode, ") \
				_T("         HMS_GETHOSPITALNAME(hc_regcode) as regplace, ") \
				_T("         hfe_deptid as deptid,") \
				_T("         hfe_serialno as serialno,") \
				_T("         hfe_receiptno as recvno,") \
				_T("         hfe_date as recvdate, ") \
				_T("         hfe_staff as receiver,") \
				_T("         hfe_desc as reason,") \
				_T("         hfe_amount as cost,") \
				_T("         hfe_discount as discount,") \
				_T("         hfe_patpaid as patpaid,") \
				_T("         hfe_deposit as deposit_amount, ") \
				_T("         hfe_payment as payment_amount,") \
				_T("         hfe_refund as refund_amount, ") \
				_T("         hfe_freeamount as free_amount,") \
				_T("         hfe_eat_amount as food_amount,") \
				_T("         row_number() over (partition by hd_docno order by hd_docno) as dn") \
				_T(" FROM hms_fee_invoice ") \
				_T(" LEFT JOIN hms_doc ON(hd_docno=hfe_docno)") \
				_T(" LEFT JOIN hms_patient ON(hp_patientno=hd_patientno)") \
				_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno) ") \
				_T(" LEFT JOIN hms_treatment_record ON(hfe_docno=htr_docno) ") \
				_T(" LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx) ") \
				_T(" WHERE hfe_docno=%ld %s ") \
				_T(" ) Tbl") \
				_T(" WHERE dn=1"), nDocumentNo, szWhere);
	
	_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		return;
	}

	rs.GetValue(_T("cost"), nTotalCost);
	rs.GetValue(_T("discount"), nTotalDiscount);
	rs.GetValue(_T("patpaid"), nTotalPatPaid);
	rs.GetValue(_T("deposit_amount"), nTotalDeposit);
	rs.GetValue(_T("free_amount"), nTotalFree);
	rs.GetValue(_T("refund_amount"), nTotalRefund);
	rs.GetValue(_T("paymennt_amount"), nTotalPayment);
	rs.GetValue(_T("reason"), szDescription);
	rs.GetValue(_T("deskno"), nDeskNo);

	rs.GetValue(_T("food_amount"), nSumFoodAmount);

	rs.GetValue(_T("objecttype"), szObjectType);
	
	rs.GetValue(_T("hd_outpatient"), szOutPatient);

	if (szOutPatient == _T("Y"))
	{
		if(!rpt.Init(_T("Reports/HMS/HF_INSDISCHARGEPAYMENTOUT.RPT"), true) )
			return;
	}
	else
	{
		if(!rpt.Init(_T("Reports/HMS/HF_INSDISCHARGEPAYMENT.RPT"), true) )
			return;
	}

	//Report Header
	tmpStr.Empty();
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);

	rs.GetValue(_T("docno"), nDocumentNo);
	tmpStr.Format(_T("%ld"), nDocumentNo);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	tmpStr.Format(_T("%s-%s-%s"), rs.GetValue(_T("serialno")), rs.GetValue(_T("bookno")),  rs.GetValue(_T("recvno")));
	rpt.GetReportHeader()->SetValue(_T("ReceiptNo"), tmpStr);
	tmpStr.Format(_T("%ld"), nInvoiceNo);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);


	CString szDate;
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")), tmpStr.Mid(11, 5),tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), szDate);


	rs.GetValue(_T("recordno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RecordNo"), tmpStr);
	//rs.GetValue(_T("pname"), tmpStr);
	

	CString szPatientName;
	StringUpper(rs.GetValue(_T("pname")), szPatientName);
	rpt.GetReportHeader()->SetValue(_T("PATIENTNAME"), szPatientName);
	rs.GetValue(_T("age"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), tmpStr);
	rs.GetValue(_T("yearofbirth"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);

	rs.GetValue(_T("birthdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("BirthDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

	rs.GetValue(_T("sex"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Sex"), tmpStr);

	rs.GetValue(_T("sex_id"), tmpStr);
	if(tmpStr == _T("F"))
		rpt.GetReportHeader()->SetValue(_T("Female"), _T("X"));
	else
		rpt.GetReportHeader()->SetValue(_T("Male"), _T("X"));


	rs.GetValue(_T("ethnic"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ethnic"), tmpStr);
	rs.GetValue(_T("ethnicdesc"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ethnicdesc"), tmpStr);
	
	rs.GetValue(_T("workplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Workplace"), tmpStr);


	rs.GetValue(_T("rankname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RankName"), tmpStr);

	CString szAddress;
	rs.GetValue(_T("detailaddress"), tmpStr);
	szAddress = tmpStr;
	rs.GetValue(_T("address"), tmpStr);
	if(!szAddress.IsEmpty())
		szAddress += _T(" - ");
	szAddress += tmpStr;
	rpt.GetReportHeader()->SetValue(_T("Address"), szAddress);
	

	rs.GetValue(_T("transplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("TransferHospital"), tmpStr);
	

	rs.GetValue(_T("rankame"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Rank"), tmpStr);
	
	rs.GetValue(_T("objecttype"), szObjectType);


	
	rs.GetValue(_T("cardno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CardNo"), tmpStr.Left(15));
	if(!tmpStr.IsEmpty())
	{
		rpt.GetReportHeader()->SetValue(_T("HasCardNo"), _T("X"));
	}
	else
	{
		rpt.GetReportHeader()->SetValue(_T("NoCardNo"), _T("X"));
	}

	rs.GetValue(_T("disrate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Insrate"), tmpStr);
	

	rs.GetValue(_T("insline"), tmpStr);	
	if (tmpStr == _T("Y"))
	{		
		rpt.GetReportHeader()->SetValue(_T("InsLine"), _T("X"));
		rpt.GetReportHeader()->SetValue(_T("NotInsLine"), _T(""));
	}
	else
	{
		rpt.GetReportHeader()->SetValue(_T("InsLine"), _T(""));
		rpt.GetReportHeader()->SetValue(_T("NotInsLine"), _T("X"));
	}
	
	rs.GetValue(_T("insregdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InsRegDate"), CDate::Convert(tmpStr));


	rs.GetValue(_T("regdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RegDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

	rs.GetValue(_T("expdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ExpiryDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	rs.GetValue(_T("regplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RegistrationPlace"), tmpStr);


	CString szRegCode;
	rs.GetValue(_T("regcode"), szRegCode);
	rpt.GetReportHeader()->SetValue(_T("RegCode"), szRegCode);

	CReportItem *obj = rpt.GetReportHeader()->GetItem(_T("InsuranceCompany"));
	if (obj)
	{
		CString insname;
		CRecord rsx(&pMF->m_db);
		
		tmpStr.Format(_T("select sp_name as insname ") \
			          _T("from sys_prov left join hms_hospital on(hh_provid=sp_id) where trim(hh_id)=trim('%s') "), szRegCode);
		rsx.ExecSQL(tmpStr);
		if (!rsx.IsEOF())
		{
			rsx.GetValue(_T("insname"), tmpStr);
			rpt.GetReportHeader()->Format(_T("InsuranceCompany"), tmpStr);
		}
	}

	rs.GetValue(_T("emergency"), tmpStr);
	if (tmpStr == _T("Y") )
	{
		tmpStr = _T("X");
		rpt.GetReportHeader()->SetValue(_T("InsLine"), _T(""));
		rpt.GetReportHeader()->SetValue(_T("NotInsLine"), _T(""));
	}
	else
		tmpStr.Empty();

	rpt.GetReportHeader()->SetValue(_T("Emergency"), tmpStr);

	rs.GetValue(_T("maindisease"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);
		rs.GetValue(_T("ireldisease"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("RelationDisease"), tmpStr);
		rs.GetValue(_T("mainicd"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("ICD10"), tmpStr);
		rs.GetValue(_T("iresult"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Result"), tmpStr);
		rs.GetValue(_T("isuggestion"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Suggestion"), pMF->GetSelectionString(_T("hms_suggestion"), tmpStr));
	
	
	
	CString szDischargeDate;
	rs.GetValue(_T("admitdate"), tmpStr);

	rpt.GetReportHeader()->SetValue(_T("admitdate"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	
	rs.GetValue(_T("dischargedate"), tmpStr);
	szDischargeDate = tmpStr;
	
	if(szDischargeDate.Left(4) != _T("1752"))
			rpt.GetReportHeader()->SetValue(_T("dischargedate"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	else
			rpt.GetReportHeader()->SetValue(_T("dischargedate"), _T(""));

	rs.GetValue(_T("TreatmentDay"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("TreatmentDay"), tmpStr);	

	CString szData;
	tmpStr.Format(_T("%s-%s-%s"), rs.GetValue(_T("bookno")), rs.GetValue(_T("serialno")), rs.GetValue(_T("recvno")));
	rpt.GetReportFooter()->SetValue(_T("SerialNo"), tmpStr);
	rs.GetValue(_T("recvdate"), szRecvDate);	
	rpt.GetReportFooter()->SetValue(_T("ReceiptDate"), CDateTime::Convert(szRecvDate, yyyymmdd|hhmm,ddmmyyyy|hhmm));
	szData.Format(rpt.GetReportFooter()->GetValue(_T("ReceiptDate1")), szRecvDate.Mid(11,5),szRecvDate.Mid(8,2),szRecvDate.Mid(5,2),szRecvDate.Left(4));	
	rpt.GetReportFooter()->SetValue(_T("ReceiptDate1"),szData);
	szData.Format(rpt.GetReportFooter()->GetValue(_T("ReceiptDate2")), szRecvDate.Mid(11,5),szRecvDate.Mid(8,2),szRecvDate.Mid(5,2),szRecvDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("ReceiptDate2"),szData);
	
	rs.GetValue(_T("treatdoctor"), tmpStr);
	tmpStr.Trim();
	rpt.GetReportFooter()->SetValue(_T("Doctor"), GetDoctorName(tmpStr));

	CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
	if(img)
	{
		img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img->SetFixedHeight(true);
	}
	tmpStr.Empty();
	rs.GetValue(_T("receiver"), tmpStr);
	tmpStr.Trim();
	rpt.GetReportFooter()->SetValue(_T("ReceiverBy"), GetDoctorName(tmpStr));
	CReportItem *img2 = rpt.GetReportFooter()->GetItem(_T("Signature2"));
	if(img2)
	{
		img2->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img2->SetFixedHeight(false);
	}	
	
	
	// Nguoi tiep don
	tmpStr.Empty();
	rs.GetValue(_T("createdby"), tmpStr);
	tmpStr.Trim();
	rpt.GetReportFooter()->SetValue(_T("createdby"), GetDoctorName(tmpStr));
	CReportItem *img3 = rpt.GetReportFooter()->GetItem(_T("Signature3"));
	if(img3)
	{
		img3->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img3->SetFixedHeight(false);
	}	


	CReportSection* rptDetail=NULL;
	CReportSection *grp;
	CRecord grs(&pMF->m_db);
	CString szFeeID, szID;
	CString szSubItem, szType;

	TCHAR *lpszChapter[] = {_T("I"), _T("II"), _T("III"), _T("IV"), _T("V"), _T("VI"), _T("VII"), _T("VIII"), _T("IX"), _T("X"), _T("XI")};


	CRecord drs(&pMF->m_db);
	CString szDeptID;
		
	CArray<RECEIPTINFO, RECEIPTINFO& > payInfo;

		szSQL.Format(_T(" SELECT deptid, admitdate, dischargedate, mainicd, maindisease") \
					_T(" FROM") \
					_T(" (") \
					_T("  select htr_deptid as deptid,") \
					_T("         htr_admitdate as admitdate,") \
					_T("         htr_dischargedate as dischargedate,") \
					_T("         htr_mainicd as mainicd,") \
					_T("         htr_maindisease as maindisease,") \
					_T("         htr_idx as idx") \
					_T("  from hms_treatment_record") \
					_T("  where htr_docno=%ld") \
					_T("  union all") \
					_T("  select hd_admitdept,") \
					_T("         hd_admitdate,") \
					_T("         hd_enddate,") \
					_T("         hd_icd,") \
					_T("         hd_diagnostic,") \
					_T("         0") \
					_T("  from hms_doc") \
					_T("  where hd_docno=%ld") \
					_T(" ) tbl") \
					_T(" WHERE deptid in(select distinct hfe_deptid from hms_fee where hfe_invoiceno=%ld)") \
					_T(" ORDER BY idx"), nDocumentNo, nDocumentNo, nInvoiceNo);

		drs.ExecSQL(szSQL);

		RECEIPTINFO pi;
		CString szDepts;
		szDepts.Empty();

		while (!drs.IsEOF())
		{
			drs.GetValue(_T("deptid"), tmpStr);
			
			bool bFound = false;
			for (int i =0; i < payInfo.GetCount(); i++)
			{
				pi = payInfo[i];
				if (pi.deptid == tmpStr)
				{
					bFound = true;
					break;
				}
			}
			pi.deptid = tmpStr;
			if(!szDepts.IsEmpty())
				szDepts += _T(",  ");
			szDepts.AppendFormat(_T("%s"), tmpStr);

			drs.GetValue(_T("dischargedate"), pi.dischargedate);
			drs.GetValue(_T("mainicd"), pi.mainicd);
			drs.GetValue(_T("maindisease"), pi.maindisease);

			if (bFound)
			{
				payInfo.SetAt(i, pi);
			}
			else
			{
				drs.GetValue(_T("admitdate"), pi.admitdate);
				payInfo.Add(pi);
				
			}

			drs.MoveNext();
		}

		rpt.GetReportHeader()->SetValue(_T("Department"), szDepts);	


		int nInsRate;
	rs.GetValue(_T("disrate"), nInsRate);
	if(nInsRate > 0)
	{
		int nUnRate = 100 - nInsRate;
		if(nUnRate ==0)
			nUnRate = nInsRate;

		tmpStr.Format(_T("(%d%s)"), nUnRate, _T("%"));
		rpt.GetReportFooter()->SetValue(_T("UnRate"), tmpStr);
	}



		double nSumCost=0, nSumInsCost=0, nSumDiscount=0, nSumDiffPaid=0, nSumPatPaid=0;

		szDepts.Empty();


		int nChapter = 0;
		int nCount = 0, nIndex, nChapterID = 0;
		int nItem = 0, nOldItem = 0, nItem2 = 0;
		bool bDeleteParent = false;
		bool bLoadItems = false;
		CString szNewGroup, szOldGroup, szParentGroup;

		
	double nGroup0Cost=0, nGroup0InsCost= 0, nGroup0Discount=0, nGroup0DiffPaid=0, nGroup0PatPaid=0;
	double nGroup1Cost=0, nGroup1InsCost = 0, nGroup1Discount=0, nGroup1DiffPaid = 0, nGroup1PatPaid=0;
	double nGroup2Cost=0, nGroup2InsCost = 0, nGroup2Discount=0, nGroup2DiffPaid = 0, nGroup2PatPaid=0;	
	double nCost=0, nInsCost =0,  nDiscount=0, nDiffPaid  = 0, nPatPaid=0;

	
	nTotalCost = nTotalDiffPaid = nTotalInsCost = nTotalPatPaid = nTotalDiscount =0;

	szSQL.Format(_T("SELECT * FROM hms_fee_type ORDER BY hft_id "));
	grs.ExecSQL(szSQL);
	
	szWhere.Format(_T(" AND (hfe_invoiceno=%ld ) "), nInvoiceNo);

	if (szOutPatient != _T("Y")) 
		szWhere.AppendFormat(_T(" and hfe_groupid <>'D0000' "));
	szWhere.AppendFormat(_T(" and hfe_class in('E', 'I','X','A') "));
//	szWhere.AppendFormat(_T(" and hfe_subgroup='XX' "));
	szWhere.AppendFormat(_T(" and (hfe_category <> 'Y' or hfe_category is null) "));

	szWhere.AppendFormat(_T(" and hfe_object not in(7) "));

//	szWhere.AppendFormat(_T(" and hfe_feegroup not in('OPT','OPT_L','HITECH','HITECH_L') "));

	szSQL.Format(_T(" SELECT hfe_typeindex AS typeindex,") \
_T("   hfe_status         AS status,") \
_T("   hfe_type           AS feetype,") \
_T("   hfe_groupid        AS groupid,") \
_T("   hfe_groupid3       AS groupid3,") \
_T("	hfe_feegroup as feegroup, ") \
_T("   hfe_itemid         AS itemid,") \
_T("   hfe_name           AS name,") \
_T("   hfe_unit           AS unit,") \
_T("   hfe_hitech         AS hitech,") \
_T("   hfe_inlist         AS inlist,") \
_T("   hfe_unitprice      AS unitprice,") \
_T("   SUM(hfe_quantity)       AS qty,") \
_T("   SUM(hfe_amount)      AS cost,") \
_T("   SUM(hfe_discount)  AS discount,") \
_T("   SUM(hfe_diffpaid)  AS DiffPaid,") \
_T("   SUM(hfe_copayment) AS copayment,") \
_T("   SUM(hfe_patpaid)   AS patpaid") \
_T(" FROM hms_fee_invoiceline_view ") \
_T(" WHERE hfe_docno  =%ld %s ") \
_T(" and hfe_feegroup not in('OPT','OPT_L','HITECH','HITECH_L')  ") \
_T(" GROUP BY hfe_typeindex,") \
_T("   hfe_status ,") \
_T("   hfe_type ,") \
_T("   hfe_groupid ,") \
_T("   hfe_groupid3 ,") \
_T(" hfe_feegroup, ") \
_T("   hfe_itemid ,") \
_T("   hfe_name ,") \
_T("   hfe_unit ,") \
_T("   hfe_hitech ,") \
_T("   hfe_inlist,") \
_T("   hfe_unitprice") \
_T(" ORDER BY hfe_typeindex, hfe_inlist, hfe_groupid3, hfe_inlist, hfe_name, hfe_unit, hfe_unitprice"), 
	nDocumentNo, szWhere);

//	_msg(_T("%s"), szSQL);		
	
	nIndex = 1;
	int nSubItem = 1;
	int nIDX;
	nChapterID = 0;
	CArray<FEEITEM, FEEITEM&> feeList;
	FEEITEM fee;
	CString szName;
	CString szNewIndex, szOldIndex;
	CString szIndex;
	bool bFound = false;
	bool bInList = false;
	bool bOutList = false;
	bool bKList = false;
	int nSubIndex = 1;
	int nLineIndex=0;
	int nNumberTreat = 0;
	
	double nTtlCost=0, nTtlInsCost=0, nTtlDiscount=0, nTtlDiffPaid=0, nTtlPatPaid=0;
	rs.ExecSQL(szSQL);
//_fmsg(_T("%s"), szSQL);


	while(!grs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex);
		grs.GetValue(_T("hft_name"), szName);
		grs.GetValue(_T("hft_id"), szNewIndex);
		
		fee.szGroupID = _T("Type");
		fee.szID = tmpStr;
		fee.szName = szName;
		nItem = feeList.Add(fee);
		
		nIDX  = nItem;
		bFound = false;
		nGroup1Cost=nGroup1InsCost=nGroup1Discount=nGroup1DiffPaid=nGroup1PatPaid=0;
		nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;

		rs.MoveFirst();
		while(!rs.IsEOF())
		{
			rs.GetValue(_T("typeindex"), tmpStr);
			
			if(tmpStr == szNewIndex){
				bFound = true;		
				rs.GetValue(_T("groupid"), szNewGroup);
				if(szNewGroup != szOldGroup)
				{
					
					szOldGroup = szNewGroup;
					/*
					if(szNewGroup.Left(2) == _T("B1") ){
						szSQL.Format(_T("SELECT hfg_name FROM hms_fee_group WHERE rtrim(hfg_id,'0')='%s' "), szNewGroup);
						CRecord rs2(&pMF->m_db);
						rs2.ExecSQL(szSQL);
						if(nGroup2Cost+nGroup2DiffPaid > 0){
							TranslateString(_T("Sub Amount"), tmpStr);
							fee.szGroupID = _T("SubAmount");
							fee.szName = tmpStr;
							fee.nCost = nGroup2Cost;
							fee.nInsPaid = nGroup2InsCost;
							fee.nDiscount = nGroup2Discount;
							fee.nDiffPaid = nGroup2DiffPaid;
							fee.nPatPaid = nGroup2PatPaid;
							int nItem2 = feeList.Add(fee);
						}

						szIndex.Format(_T("%d.%d"), nIndex, nSubIndex++);
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = rs2.GetValue(_T("hfg_name"));
						int nItem2 = feeList.Add(fee);
						nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;

					}

					*/
					
					
				}

				if(szOldIndex != szNewIndex)
				{
					szOldIndex = szNewIndex;
					bInList = false;
					bOutList = false;
				}

				

				if(szNewIndex == _T("800")){
					

					rs.GetValue(_T("inlist"), tmpStr);
					
					if(tmpStr == _T("1") && !bInList){
						bInList = true;
						TranslateString(_T("Inside insurance list"), tmpStr);
						szIndex.Format(_T("%d.1"), nIndex);	
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = tmpStr;
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nDiffPaid = 0;
						fee.nPatDebt = 0;

						int nItem2 = feeList.Add(fee);

					}
					if(tmpStr == _T("2") && !bOutList){
						bOutList = true;
						TranslateString(_T("Outside insurance list"), tmpStr);
						if(bInList)
							szIndex.Format(_T("%d.2"), nIndex);	
						else
							szIndex.Format(_T("%d.1"), nIndex);	

						szIndex.Format(_T("%d.1"), nIndex);	
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = tmpStr;
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nDiffPaid = 0;
						fee.nPatDebt = 0;
						int nItem2 = feeList.Add(fee);
					}

					if(tmpStr == _T("3") && !bKList){
						bKList = true;
						if(bInList && bOutList)
							szIndex.Format(_T("%d.3"), nIndex);	
						else if(bInList || bOutList)
							szIndex.Format(_T("%d.2"), nIndex);	
						else
							szIndex.Format(_T("%d.1"), nIndex);	

						TranslateString(_T("Inside cancer list"), tmpStr);
						szIndex.Format(_T("%d.1"), nIndex);	
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = tmpStr;
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nDiffPaid = 0;
						fee.nPatDebt = 0;
						int nItem2 = feeList.Add(fee);
					}
					

				}

				if(szNewIndex == _T("900")){
					
					rs.GetValue(_T("inlist"), tmpStr);
					
					if(tmpStr == _T("1") && !bInList){
						bInList = true;
						TranslateString(_T("Inside insurance list"), tmpStr);
						szIndex.Format(_T("%d.1"), nIndex);
						szIndex.Format(_T("%d.1"), nIndex);	
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = tmpStr;
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nDiffPaid = 0;
						fee.nPatDebt = 0;
						int nItem2 = feeList.Add(fee);
					}
					if(tmpStr == _T("2") && !bOutList){
						bOutList = true;
						TranslateString(_T("Outside insurance list"), tmpStr);
						if(!bInList)
							szIndex.Format(_T("%d.1"), nIndex);
						else
							szIndex.Format(_T("%d.2"), nIndex);

						szIndex.Format(_T("%d.1"), nIndex);	
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = tmpStr;
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nDiffPaid = 0;
						fee.nPatDebt = 0;
						int nItem2 = feeList.Add(fee);
					}
				}

				
				//IN CHI PHI PHAU THUAT KY THUAT CAO
				CString szFeeGroup;
				rs.GetValue(_T("feegroup"), szFeeGroup);
				if (szFeeGroup == _T("HITECH") || szFeeGroup == _T("HITECH_L") )
				{
					CString szItemID;
					CString tmpStr;
					CString szOptGroup;


					int nCount = 0;
					double nTtlCost = 0, nCostA = 0, nCostB=0;
					double nTtlDiffCost = 0;
					double nOptDiffCost = 0;
					double nTtlInsPaid = 0;
					double nAmount = 0;

					if (szFeeGroup == _T("HITECH"))
						szOptGroup = _T("OPT");
					else
						szOptGroup = _T("OPT_L");

					rs.GetValue(_T("itemid"), szItemID);
					rs.GetValue(_T("cost"), nTtlCost);
					rs.GetValue(_T("diffpaid"), nTtlDiffCost);
					nOptDiffCost = nTtlDiffCost;
					rs.GetValue(_T("discount"), nAmount);
					nTtlInsPaid += nAmount;


					fee.szID.Format(_T("%d"), nLineIndex++);
					fee.szGroupID = _T("Item");
					fee.szName = rs.GetValue(_T("name"));
					fee.szUnit = rs.GetValue(_T("unit"));
					fee.nCost = nTtlCost;
					fee.nInsPaid = 0;
					fee.nDiscount = 0;
					fee.nQuantity = str2float(rs.GetValue(_T("qty")));
					fee.nPrice = str2double(rs.GetValue(_T("unitprice")));
					fee.nInsPrice = str2double(rs.GetValue(_T("insprice")));
					fee.nDiffPaid = str2double(rs.GetValue(_T("diffpaid")));
					fee.nDiscount = str2double(rs.GetValue(_T("discount")));
					fee.nPatPaid = 0;
					feeList.Add(fee);

					
					szSQL.Format(_T("SELECT hfe_desc as hfe_name, hfe_unit, hfe_unitprice, sum(hfe_quantity) as hfe_quantity, sum(hfe_inspaid) as hfe_cost, sum(hfe_discount) as hfe_discount, sum(hfe_diffpaid) as hfe_diffpaid ") \
						_T(" FROM hms_fee ") \
						_T(" WHERE hfe_docno=%ld and hfe_type in('D','M') ") \
						_T(" and hfe_feegroup IN('%s') and hfe_subgroup='%s' and hfe_discount > 0 " )\
						_T(" and hfe_category <> 'Y' ") \
						_T(" GROUP BY hfe_desc, hfe_unit, hfe_unitprice ") \
						_T(" ORDER BY hfe_desc "),
						nDocumentNo, szOptGroup, szItemID);
					rsl.ExecSQL(szSQL);

					int n = 1;
					if (rsl.GetRecordCount() > 0 ) 
					{
//						nItem = pList->AddItems(_T("a."), _T("Trong \x64\x61nh m\x1EE5\x63 \x42HYT"), NULL);
//						pList->MergeCell(nItem, 1, nItem, 8, DT_LEFT|DT_VCENTER|DT_SINGLELINE, false, true);
//						pList->SetItemBkColor(nItem, RGB(220, 220, 220), FALSE);

						fee.szID.Format(_T("a."));
						fee.szGroupID = _T("SubItemGroup");
						fee.szName.Format(_T("Trong \x64\x61nh m\x1EE5\x63 \x42HYT"));
						fee.szUnit.Empty();
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nQuantity = 0;
						fee.nPrice = 0;
						fee.nInsPrice = 0;
						fee.nDiffPaid = 0;
						fee.nPatPaid = 0;
						feeList.Add(fee);

					}
					
					while(!rsl.IsEOF())
					{
						//tmpStr.Format(_T("%d.%d - %s"), nIndex2, n++, rsl.GetValue(_T("hfe_name")));
						fee.szID.Empty();
						fee.szGroupID = _T("Item");
						fee.szName.Format(_T("%s"), rsl.GetValue(_T("hfe_name")));
						fee.szUnit = rsl.GetValue(_T("hfe_unit"));
						rsl.GetValue(_T("hfe_cost"), nAmount);
						rsl.GetValue(_T("hfe_diffpaid"), fee.nDiffPaid);
						fee.nPrice = str2double(rsl.GetValue(_T("hfe_unitprice")));
						fee.nInsPrice = str2double(rsl.GetValue(_T("hfe_unitprice")));
						fee.nCost = nAmount;
						fee.nInsPaid = 0;
						fee.nDiscount = str2double(rsl.GetValue(_T("hfe_discount")));
						fee.nQuantity = str2float(rsl.GetValue(_T("hfe_quantity")));
						//fee.nDiffPaid = 0;
						fee.nPatPaid = 0;
						
						
						nCostA += nAmount;

						rsl.GetValue(_T("hfe_discount"), nAmount);
						nTtlInsPaid += nAmount;

						rsl.GetValue(_T("hfe_diffpaid"), nAmount);
						fee.nDiffPaid = nAmount;
						nTtlDiffCost += nAmount;

						feeList.Add(fee);

						rsl.MoveNext();
						nCount ++;
					}
					
					if(nCount > 0)
					{
						tmpStr.Format(_T("%.2f"), nCostA);
					/*	nItem = pList->AddItems(_T(""), _T("\x43\x1ED9ng (a)"), _T(""), _T(""), _T(""), tmpStr, NULL);
						pList->MergeCell(nItem, 1, nItem, 4, DT_LEFT|DT_VCENTER|DT_SINGLELINE, false, true);
						pList->SetItemBkColor(nItem, RGB(230, 230, 230), FALSE);*/

						fee.szID.Empty();
						fee.szGroupID = _T("SubItemAmount");
						fee.szName.Format(_T("\x43\x1ED9ng (a)"));
						fee.szUnit.Empty();
						fee.nCost = nCostA;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nQuantity = 0;
						fee.nPrice = 0;
						fee.nInsPrice = 0;
						fee.nDiffPaid = 0;
						fee.nPatPaid = 0;
						feeList.Add(fee);
						
					}
					nTtlCost += nCostA;
					
					
					tmpStr.Format(_T("%.2f"), nTtlCost);



					//nItem = pList->AddItems(_T("+"), _T("T\x1ED5ng \x63hi ph\xED trong quy \x111\x1ECBnh \x42HYT"), _T(""), _T(""), _T(""), tmpStr, NULL);
					//pList->MergeCell(nItem, 1, nItem, 4, DT_LEFT|DT_VCENTER|DT_SINGLELINE, false, true);
					//pList->SetItemBkColor(nItem, RGB(200, 220, 220), FALSE);


					fee.szID.Format(_T("+"));
						fee.szGroupID = _T("SubItemAmount");
						fee.szName.Format(_T("T\x1ED5ng \x63hi ph\xED trong quy \x111\x1ECBnh \x42HYT"));
						fee.szUnit.Empty();
						fee.nCost = nTtlCost;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nQuantity = 0;
						fee.nPrice = 0;
						fee.nInsPrice = 0;
						fee.nDiffPaid = 0;
						fee.nPatPaid = 0;
						feeList.Add(fee);

					//Ngoai danh muc
					szSQL.Format(_T("SELECT hfe_desc as hfe_name, hfe_unit, hfe_unitprice, sum(hfe_quantity) as hfe_quantity, sum(hfe_cost) as hfe_cost ") \
						_T("FROM hms_fee ") \
						_T("WHERE hfe_docno=%ld and hfe_type in('D','M') ") \
						_T("and hfe_feegroup IN('%s') and hfe_subgroup='%s' and hfe_discount <= 0 and hfe_category <> 'Y' ") \
						_T("GROUP BY hfe_desc, hfe_unit, hfe_unitprice ORDER BY hfe_desc "),
						nDocumentNo, szOptGroup, szItemID);
					rsl.ExecSQL(szSQL);

					
					if (rsl.GetRecordCount() > 0 ) 
					{
						/*nItem = pList->AddItems(_T("b."), _T("Ngo\xE0i \x64\x61nh m\x1EE5\x63 \x42HYT"), NULL);
						pList->MergeCell(nItem, 1, nItem, 8, DT_LEFT|DT_VCENTER|DT_SINGLELINE, false, true);
						pList->SetItemBkColor(nItem, RGB(220, 220, 220), FALSE);*/

						fee.szID.Format(_T("b."));
						fee.szGroupID = _T("SubItemGroup");
						fee.szName.Format(_T("Ngo\xE0i \x64\x61nh m\x1EE5\x63 \x42HYT"));
						fee.szUnit.Empty();
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nQuantity = 0;
						fee.nPrice = 0;
						fee.nInsPrice = 0;
						fee.nDiffPaid = 0;
						fee.nPatPaid = 0;
						feeList.Add(fee);

					}
					nCount = 0;
					while(!rsl.IsEOF())
					{
						/*tmpStr.Format(_T("%d.%d - %s"), nIndex2, n++, rsl.GetValue(_T("hfe_name")));
						nItem = pList->AddItems(
							_T(""),
							tmpStr,
							rsl.GetValue(_T("hfe_unit")),
							rsl.GetValue(_T("hfe_quantity")),
							rsl.GetValue(_T("hfe_unitprice")),
							rsl.GetValue(_T("hfe_cost")),
							_T(""),
							_T(""),
							_T(""),
							_T("subitem"),
							NULL);*/

						fee.szID.Empty();
						fee.szGroupID = _T("Item");
						fee.szName.Format(_T("%s"), rsl.GetValue(_T("hfe_name")));
						fee.szUnit = rsl.GetValue(_T("hfe_unit"));
						rsl.GetValue(_T("hfe_cost"), nAmount);
						fee.nPrice = str2double(rsl.GetValue(_T("hfe_unitprice")));
						fee.nInsPrice = fee.nPrice;
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nQuantity = str2float(rsl.GetValue(_T("hfe_quantity")));
						fee.nDiffPaid = nAmount;
						fee.nPatPaid = 0;
						feeList.Add(fee);
						nCostB += nAmount;
						nCount ++;

						rsl.MoveNext();
					}

					if(nCount > 0)
					{
						/*tmpStr.Format(_T("%.2f"), nCostB);
						nItem = pList->AddItems(_T(""), _T("\x43\x1ED9ng (b)"), _T(""), _T(""), _T(""), tmpStr,  NULL);
						pList->MergeCell(nItem, 1, nItem, 4, DT_LEFT|DT_VCENTER|DT_SINGLELINE, false, true);
						pList->SetItemBkColor(nItem, RGB(220, 220, 220), FALSE);*/

						fee.szID.Empty();
						fee.szGroupID = _T("SubItemAmount");
						fee.szName.Format(_T("\x43\x1ED9ng (b)"));
						fee.szUnit.Empty();
						fee.nCost = nCostB;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nQuantity = 0;
						fee.nPrice = 0;
						fee.nInsPrice = 0;
						fee.nDiffPaid = 0;
						fee.nPatPaid = 0;
						feeList.Add(fee);

					}
					if(nTtlDiffCost > 0)
					{
						double nMatDiffCost = nTtlDiffCost-nOptDiffCost;
						if (nOptDiffCost > 0)
						{
							fee.szID.Empty();
							fee.szGroupID = _T("SubItemAmount");
							fee.szName.Format(_T("\x43hi ph\xED \x63h\xEAnh l\x1EC7\x63h PTTT"));
							fee.szUnit.Empty();
							fee.nCost = nOptDiffCost;
							fee.nInsPaid = 0;
							fee.nDiscount = 0;
							fee.nQuantity = 0;
							fee.nPrice = 0;
							fee.nInsPrice = 0;
							fee.nDiffPaid = 0;
							fee.nPatPaid = 0;
							feeList.Add(fee);
						}
						if (nMatDiffCost > 0)
						{
							fee.szID.Empty();
							fee.szGroupID = _T("SubItemAmount");
							fee.szName.Format(_T("\x43hi ph\xED \x63h\xEAnh l\x1EC7\x63h thu\x1ED1\x63, v\x1EADt t\x1B0"));
							fee.szUnit.Empty();
							fee.nCost = nMatDiffCost;
							fee.nInsPaid = 0;
							fee.nDiscount = 0;
							fee.nQuantity = 0;
							fee.nPrice = 0;
							fee.nInsPrice = 0;
							fee.nDiffPaid = 0;
							fee.nPatPaid = 0;
							feeList.Add(fee);
						}

					}
					double nTtlCostOut = nTtlDiffCost+nCostB;
					if(nTtlCostOut > 0)
					{
						/*tmpStr.Format(_T("%.2f"), nTtlCostOut);
						nItem = pList->AddItems(_T("+"), _T("T\x1ED5ng \x63hi ph\xED ngo\xE0i quy \x111\x1ECBnh \x42HYT kh\xF4ng th\x61nh to\xE1n"), _T(""), _T(""), _T(""),  tmpStr, NULL);
						pList->MergeCell(nItem, 1, nItem, 4, DT_LEFT|DT_VCENTER|DT_SINGLELINE, false, true);
						pList->SetItemBkColor(nItem, RGB(200, 220, 220), FALSE);*/
						fee.szID.Format(_T("+"));
						fee.szGroupID = _T("SubItemAmount");
						fee.szName.Format(_T("T\x1ED5ng \x63hi ph\xED ngo\xE0i quy \x111\x1ECBnh \x42HYT kh\xF4ng th\x61nh to\xE1n"));
						fee.szUnit.Empty();
						fee.nCost = nTtlCostOut;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nQuantity = 0;
						fee.nPrice = 0;
						fee.nInsPrice = 0;
						fee.nDiffPaid = 0;
						fee.nPatPaid = 0;
						feeList.Add(fee);

					}
					szSQL.Format(_T("SELECT hfe_discount FROM hms_fee ") \
						_T("WHERE hfe_docno = %ld and hfe_type='V' and hfe_itemid='VT000002' and hfe_subgroup='%s'"),
						nDocumentNo, szItemID);
					rsx.ExecSQL(szSQL);
					
					if(!rsx.IsEOF())
					{
						rsx.GetValue(_T("hfe_discount"), nAmount);
						nTtlInsPaid = nAmount;
					}
					else
					{

					}
					nTtlDiffCost += nCostB;


					/*tmpStr.Format(_T("%.2f"), nTtlCost);
					nItem = pList->AddItems(_T("*"), _T("T\x1ED5ng s\x1ED1 ti\x1EC1n \x64\x1ECB\x63h v\x1EE5 v\xE0 \x64\x1EE5ng \x63\x1EE5"), _T(""), _T(""), _T(""), tmpStr, ToString(nTtlInsPaid), ToString(nTtlDiffCost), ToString(nTtlCost-nTtlInsPaid), NULL);
					pList->MergeCell(nItem, 1, nItem, 4, DT_LEFT|DT_VCENTER|DT_SINGLELINE, false, true);
					pList->SetItemBkColor(nItem, RGB(200, 220, 220), FALSE);*/

					fee.szID = _T("*");
						fee.szGroupID = _T("SubItemAmount");
						fee.szName.Format(_T("S\x1ED1 ti\x1EC1n \x64\x1ECB\x63h v\x1EE5 v\xE0 \x64\x1EE5ng \x63\x1EE5 th\x61nh to\xE1n"));
						fee.szUnit.Empty();
						fee.nCost = nTtlCost;
						fee.nInsPaid = nTtlCost;
						fee.nDiscount = nTtlInsPaid;
						fee.nQuantity = 0;
						fee.nPrice = 0;
						fee.nInsPrice = 0;
						fee.nDiffPaid = nTtlDiffCost;
						fee.nPatPaid = nTtlCost-nTtlInsPaid;
						feeList.Add(fee);

					nTotalCost += nTtlCost;
					nGroup1Cost += nTtlCost;
					nGroup2Cost += nTtlCost;

					/*nTotalInsCost += nInsCost;
					nGroup1InsCost += nInsCost;
					nGroup2InsCost += nInsCost;*/

					nTotalDiscount += nTtlInsPaid;
					nGroup1Discount += nTtlInsPaid;
					nGroup2Discount += nTtlInsPaid;

					nTotalDiffPaid += nTtlDiffCost;
					nGroup1DiffPaid += nTtlDiffCost;
					nGroup2DiffPaid += nTtlDiffCost;

					nAmount = (nTtlCost-nTtlInsPaid);
					nTotalPatPaid += nAmount;
					nGroup1PatPaid += nAmount;
					nGroup2PatPaid += nAmount;
					nAmount = 0;


				}
				else
				{
					rs.GetValue(_T("cost"), nCost);
					rs.GetValue(_T("inscost"), nInsCost);
					rs.GetValue(_T("discount"), nDiscount);
					rs.GetValue(_T("DiffPaid"), nDiffPaid);
					rs.GetValue(_T("copayment"), nPatPaid);
					
					if(nInsCost <=0 )
					{
					//	nDiffPaid = nCost;
					//	nCost = 0;
					}
					nTotalCost += nCost;
					nGroup1Cost += nCost;
					nGroup2Cost += nCost;

					nTotalInsCost += nInsCost;
					nGroup1InsCost += nInsCost;
					nGroup2InsCost += nInsCost;

					nTotalDiscount += nDiscount;
					nGroup1Discount += nDiscount;
					nGroup2Discount += nDiscount;

					nTotalDiffPaid += nDiffPaid;
					nGroup1DiffPaid += nDiffPaid;
					nGroup2DiffPaid += nDiffPaid;

					nTotalPatPaid += nPatPaid;
					nGroup1PatPaid += nPatPaid;
					nGroup2PatPaid += nPatPaid;


					fee.szID.Format(_T("%d"), nLineIndex++);
					fee.szGroupID = _T("Item");
					fee.szName = rs.GetValue(_T("name"));
					fee.szUnit = rs.GetValue(_T("unit"));
					fee.nCost = nCost;
					fee.nInsPaid = nInsCost;
					fee.nDiscount = nDiscount;
					fee.nQuantity = str2float(rs.GetValue(_T("qty")));
					fee.nPrice = str2double(rs.GetValue(_T("unitprice")));
					fee.nInsPrice = str2double(rs.GetValue(_T("insprice")));
					fee.nDiffPaid = nDiffPaid;
					fee.nPatPaid = nPatPaid;
					feeList.Add(fee);

					if(szNewGroup.Left(1) == _T("C"))
					{
						nNumberTreat += (int)fee.nQuantity;

					}

				}

			}
			else
				nLineIndex = 1;
			rs.MoveNext();
		}
		if(!bFound)
		{
			feeList.RemoveAt(nIDX);
			nGroup1Cost=nGroup1InsCost=nGroup1Discount=nGroup1DiffPaid=nGroup1PatPaid=0;
			nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;
		}
		else
		{
			if(nGroup1Cost+nGroup1DiffPaid > 0)
			{
				if(nGroup2Cost > 0 && nGroup1Cost != nGroup2Cost){
					TranslateString(_T("Sub Amount"), tmpStr);
					fee.szGroupID = _T("SubAmount");
					fee.szID.Empty();;
					fee.szName = tmpStr;
					fee.nCost = nGroup2Cost;
					fee.nInsPaid = nGroup2InsCost;
					fee.nDiscount = nGroup2Discount;
					fee.nDiffPaid = nGroup2DiffPaid;
					fee.nPatPaid = nGroup2PatPaid;
					feeList.Add(fee);
				}

				TranslateString(_T("Total"), tmpStr);
				tmpStr.AppendFormat(_T(" (%d)"), nIndex);
				fee.szGroupID = _T("GrandAmount");
				fee.szID.Empty();;
				fee.szName = tmpStr;
				fee.nCost = nGroup1Cost;
				fee.nInsPaid = nGroup1InsCost;
				fee.nDiscount = nGroup1Discount;
				fee.nDiffPaid = nGroup1DiffPaid;
				fee.nPatPaid = nGroup1PatPaid;
				feeList.Add(fee);
			}
			nGroup1Cost=nGroup1InsCost=nGroup1Discount=nGroup1DiffPaid=nGroup1PatPaid=0;
			nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;

			nIndex++;
		}



		grs.MoveNext();
	}


	nTotalCost  = ceil(nTotalCost);

	if(szObjectType == _T("I") || szObjectType == _T("C"))
	{
		nTotalInsCost  = (nTotalCost);
	}
	nTotalDiscount  = (nTotalDiscount);
	nTotalDiffPaid  = (nTotalDiffPaid);
	nTotalPatPaid  = (nTotalPatPaid);
	nTotalPayable = (nTotalCost-nTotalDiscount);
	

	TranslateString(_T("Total Amount"), szName);
	fee.szGroupID = _T("TotalAmount");
	fee.szName.Format(_T("%s (II)"), szName);
	fee.nCost = nTotalCost;
	fee.nInsPaid = nTotalInsCost;
	fee.nDiscount = nTotalDiscount;
	fee.nDiffPaid = nTotalDiffPaid;
	fee.nPatPaid = nTotalPatPaid;
	feeList.Add(fee);
	

	//In phan dich vu ky thuat cao va chi phi lon


	szWhere.Format(_T(" AND (hfe_invoiceno=%ld ) "), nInvoiceNo);

	if (szOutPatient != _T("Y")) 
		szWhere.AppendFormat(_T(" and hfe_group <>'D0000' "));
	szWhere.AppendFormat(_T(" and hfe_class in('E', 'I','X','A') "));
//	szWhere.AppendFormat(_T(" and hfe_subgroup='XX' "));
	szWhere.AppendFormat(_T(" and (hfe_category <> 'Y' or hfe_category is null) "));

	szWhere.AppendFormat(_T(" and hfe_object not in(7) "));

	szSQL.Format(_T(" SELECT hfe_status   AS status,") \
_T("   hfe_type          AS feetype,") \
_T("   hfe_feegroup      AS feegroup,") \
_T("   hfe_itemid        AS itemid,") \
_T("   hfe_desc          AS name,") \
_T("   hfe_unit          AS unit,") \
_T("   hfe_hitech        AS hitech,") \
_T("   hfe_unitprice     AS unitprice,") \
_T("   hfe_insprice     AS insprice,") \
_T("   SUM(hfe_quantity) AS qty,") \
_T("   SUM(hfe_inspaid)     AS cost,") \
_T("   SUM(hfe_discount) AS discount,") \
_T("   SUM(hfe_diffpaid) AS DiffPaid,") \
_T("   SUM(hfe_patpaid+hfe_patdebt)  AS patpaid") \
_T(" FROM hms_fee") \
_T(" WHERE hfe_docno     =%ld %s ") \
_T(" AND hfe_feegroup   IN('HITECH','HITECH_L')") \
_T(" GROUP BY hfe_status ,") \
_T("   hfe_type ,") \
_T("   hfe_feegroup,") \
_T("   hfe_itemid,") \
_T("   hfe_desc ,") \
_T("   hfe_unit ,") \
_T("   hfe_hitech ,") \
_T("   hfe_unitprice, ") \
_T("   hfe_insprice ") \
_T(" ORDER BY hfe_desc,") \
_T("   hfe_unit,") \
_T("   hfe_unitprice"), nDocumentNo, szWhere);


	rs.ExecSQL(szSQL);

	if(!rs.IsEOF())
	{
		fee.szGroupID = _T("HITECH_HEADER");
		fee.szID = _T("");
		fee.szName.Format(_T(""));
		fee.nCost = 0;
		fee.nInsPaid = 0;
		fee.nDiscount = 0;
		fee.nDiffPaid = 0;
		fee.nPatPaid = 0;
		feeList.Add(fee);


		double nCostA = 0, nCostB=0;
		double nCostAIns = 0;
		double nOptDiffCost = 0;
		double nAmount = 0;

		double ttl_amount = 0;
		double ttl_diffcost = 0;
		double ttl_inspaid = 0;
		double ttl_discount = 0;
		double ttl_patpaid = 0;


		double ttl2_amount = 0;
		double ttl2_diffcost = 0;
		double ttl2_inspaid = 0;
		double ttl2_discount = 0;
		double ttl2_patpaid = 0;

		double grp_cost =0, grp_inspaid = 0, grp_discount = 0, grp_diffcost = 0, grp_patpaid =0;
		CString szItemID;
		CString szItems;
		CString szOptGroup;
		long nOrderID;


		int nHitechCount = 0;
		while(!rs.IsEOF())
		{
			nCostA  = 0;
			nCostB = 0;
			nOptDiffCost = 0;
			ttl_amount  = 0;
			ttl_inspaid = 0;
			ttl_discount = 0;
			ttl_diffcost = 0;
			ttl_patpaid = 0;
			nLineIndex = 1;

			nHitechCount ++;
			if(nHitechCount > 1)
			{
				fee.szGroupID = _T("HITECH_TYPE");
				fee.szID = _T("");
				fee.szName.Format(_T("************************************"));			
				feeList.Add(fee);
			}
		
			fee.szGroupID = _T("HITECH_TYPE");
			fee.szID = _T("A");
			fee.szName.Format(_T("Trong quy \x111\x1ECBnh \x42H\x58H"));
			fee.nCost = 0;
			fee.nInsPaid = 0;
			fee.nDiscount = 0;
			fee.nDiffPaid = 0;
			fee.nPatPaid = 0;
			feeList.Add(fee);


			CString szFeeGroup;
			rs.GetValue(_T("feegroup"), szFeeGroup);
			rs.GetValue(_T("hfe_orderid"), nOrderID);
			
			
			

			if (szFeeGroup == _T("HITECH"))
				szOptGroup = _T("OPT");
			else
				szOptGroup = _T("OPT_L");


			rs.GetValue(_T("itemid"), szItemID);

			rs.GetValue(_T("cost"), nAmount);
			
			fee.szID.Format(_T("%d"), nLineIndex++);
			fee.szGroupID = _T("HITECH_ITEM");
			fee.szName = rs.GetValue(_T("name"));
			fee.szUnit = rs.GetValue(_T("unit"));
			rs.GetValue(_T("patpaid"), fee.nPatPaid);
			
			
			nCostA += nAmount;
			fee.nCost = nAmount;
			ttl_amount += nAmount;
			ttl_inspaid += nAmount;
			grp_cost += nAmount;

			rs.GetValue(_T("discount"), nAmount);
			fee.nDiscount = nAmount;
			ttl_discount += nAmount;
			grp_discount += nAmount;
			
			fee.nQuantity = str2float(rs.GetValue(_T("qty")));
			fee.nPrice = str2double(rs.GetValue(_T("insprice")));
			
			
			rs.GetValue(_T("diffpaid"), nOptDiffCost);
			fee.nDiffPaid = nOptDiffCost;
			ttl_diffcost += nOptDiffCost;
			grp_diffcost += nOptDiffCost;

			nAmount = fee.nCost-fee.nDiscount;

			fee.nPatPaid = nAmount;
			ttl_patpaid += nAmount;
			grp_patpaid += nAmount;

			feeList.Add(fee);
			
			
			nCostAIns = fee.nInsPaid;

			
			fee.szID.Empty();
			fee.szGroupID = _T("HITECH_ITEM_GROUP");
			fee.szName.Format(_T("\x43\x1ED9ng (1)"));
			fee.szUnit.Empty();
			fee.nCost = nCostA;
			fee.nInsPaid = nCostA;
			fee.nDiscount = ttl_discount;
			fee.nQuantity = 0;
			fee.nPrice = 0;
			fee.nInsPrice = 0;
			fee.nDiffPaid = ttl_diffcost;
			fee.nPatPaid = ttl_patpaid;
			feeList.Add(fee);

			
						
			szSQL.Format(_T("SELECT hfe_desc as hfe_name, hfe_unit, hfe_unitprice, sum(hfe_quantity) as hfe_quantity, ") \
						_T("sum(hfe_cost) as hfe_cost, sum(hfe_inspaid) as hfe_inspaid, sum(hfe_discount) as hfe_discount, sum(hfe_diffcost) as hfe_diffcost ") \
							_T(" FROM hms_fee ") \
							_T(" WHERE hfe_docno=%ld and hfe_type in('D','M') ") \
							_T(" and hfe_feegroup IN('%s') and hfe_subgroup='%s' and hfe_discount > 0 " )\
							_T(" and hfe_category <> 'Y' ") \
							_T(" GROUP BY hfe_desc, hfe_unit, hfe_unitprice ") \
							_T(" ORDER BY hfe_desc "),
							nDocumentNo, szOptGroup, szItemID);

			rsl.ExecSQL(szSQL);
			
			int n = 1;
			if (rsl.GetRecordCount() > 0 ) 
			{
				fee.szID.Format(_T("2"));
				fee.szGroupID = _T("HITECH_TYPE");
				fee.szName.Format(_T("V\x1EADt t\x1B0 ti\xEAu h\x61o \x111\x1EB7\x63 \x62i\x1EC7t"));
				fee.szUnit.Empty();
				feeList.Add(fee);
							
			}
			
			int nIdx = 1;
			nCostA = 0;
			nCount = 0;
			grp_cost = grp_inspaid = grp_discount = grp_diffcost = grp_patpaid = 0;
			while(!rsl.IsEOF())
			{
				fee.szID.Format(_T("2.%d"), nIdx++);
				fee.szGroupID = _T("HITECH_ITEM");
				fee.szName.Format(_T("%s"), rsl.GetValue(_T("hfe_name")));
				fee.szUnit = rsl.GetValue(_T("hfe_unit"));
				
				fee.nPrice = str2double(rsl.GetValue(_T("hfe_unitprice")));
				fee.nInsPrice = str2double(rsl.GetValue(_T("hfe_unitprice")));
				fee.nQuantity = str2float(rsl.GetValue(_T("hfe_quantity")));

				rsl.GetValue(_T("hfe_inspaid"), nAmount);
				fee.nCost = nAmount;
				grp_cost += nAmount;
				ttl_amount += nAmount;


				rsl.GetValue(_T("hfe_inspaid"), nAmount);
				fee.nInsPaid = nAmount;
				grp_inspaid += nAmount;
				ttl_inspaid += nAmount;
				

				rsl.GetValue(_T("hfe_discount"), nAmount);
				fee.nDiscount = nAmount;
				grp_discount += nAmount;
				ttl_discount += nAmount;
				
				rsl.GetValue(_T("hfe_diffcost"), nAmount);
				fee.nDiffPaid = nAmount;
				grp_diffcost += nAmount;
				ttl_diffcost += nAmount;
				nAmount = 0;

				fee.nPatPaid = fee.nInsPaid-fee.nDiscount;
				grp_patpaid += fee.nPatPaid;
				ttl_patpaid += fee.nPatPaid;



				feeList.Add(fee);

				rsl.MoveNext();
				nCount ++;
			}
						
			if(nCount > 0)
			{
				tmpStr.Format(_T("%.2f"), nCostA);
				fee.szID.Empty();
				fee.szGroupID = _T("HITECH_ITEM_GROUP");
				fee.szName.Format(_T("\x43\x1ED9ng (2)"));
				fee.szUnit.Empty();
				fee.nCost = grp_cost;
				fee.nInsPaid = grp_inspaid;
				fee.nDiscount = grp_discount;
				fee.nQuantity = 0;
				fee.nPrice = 0;
				fee.nInsPrice = 0;
				fee.nDiffPaid = grp_diffcost;
				fee.nPatPaid = grp_patpaid;
				feeList.Add(fee);
							
			}
				
			fee.szID.Format(_T(""));
			fee.szGroupID = _T("HITECH_ITEM_GROUP");
			fee.szName.Format(_T("\x43\x1ED9ng \x41(\x31+\x32)"));
			fee.szUnit.Empty();
			fee.nCost = ttl_amount;
			fee.nInsPaid = ttl_inspaid;
			fee.nDiscount = ttl_discount;
			fee.nDiffPaid = ttl_diffcost;
			fee.nPatPaid = ttl_patpaid;
			feeList.Add(fee);


				//Ngoai danh muc
			szSQL.Format(_T("SELECT hfe_desc as hfe_name, hfe_unit, hfe_unitprice, sum(hfe_quantity) as hfe_quantity, ") \
				_T("sum(hfe_cost) as hfe_cost ") \
							_T("FROM hms_fee ") \
							_T("WHERE hfe_docno=%ld and hfe_type in('D','M') ") \
							_T("and hfe_feegroup IN('%s') and hfe_subgroup='%s'  and hfe_discount <= 0 and hfe_category <> 'Y' ") \
							_T("GROUP BY hfe_desc, hfe_unit, hfe_unitprice ORDER BY hfe_desc "),
							nDocumentNo, szOptGroup, szItemID);
				rsl.ExecSQL(szSQL);

				bool bOutList = false;

						
				if (rsl.GetRecordCount() > 0) 
				{
					fee.szID.Format(_T("B."));
					fee.szGroupID = _T("HITECH_TYPE");
					fee.szName.Format(_T("Ngo\xE0i \x64\x61nh m\x1EE5\x63 \x42HYT"));
					fee.szUnit.Empty();
					fee.nCost = 0;
					fee.nInsPaid = 0;
					fee.nDiscount = 0;
					fee.nQuantity = 0;
					fee.nPrice = 0;
					fee.nInsPrice = 0;
					fee.nDiffPaid = 0;
					fee.nPatPaid = 0;
					feeList.Add(fee);
					bOutList = true;
				}


				nCount = 0;
				grp_diffcost = 0;

				while(!rsl.IsEOF())
				{
					fee.szID.Empty();
					fee.szGroupID = _T("HITECH_ITEM");
					fee.szName.Format(_T("%s"), rsl.GetValue(_T("hfe_name")));
					fee.szUnit = rsl.GetValue(_T("hfe_unit"));
					
					fee.nPrice = str2double(rsl.GetValue(_T("hfe_unitprice")));
					
					rsl.GetValue(_T("hfe_cost"), nAmount);
					fee.nInsPrice = fee.nPrice;
					fee.nDiffPaid = nAmount;
					fee.nCost = 0;
					fee.nDiscount = 0;
					fee.nInsPaid = 0;
					fee.nPatPaid = 0;
					
					ttl_diffcost += nAmount;
					grp_diffcost += nAmount;

					fee.nQuantity = str2float(rsl.GetValue(_T("hfe_quantity")));
					feeList.Add(fee);
					nAmount = 0;
					nCount ++;
					rsl.MoveNext();
				}

				if(nCount > 0)
				{
						fee.szID.Empty();
						fee.szGroupID = _T("HITECH_ITEM_GROUP");
						fee.szName.Format(_T("\x43\x1ED9ng B"));
						fee.szUnit.Empty();
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nQuantity = 0;
						fee.nPrice = 0;
						fee.nInsPrice = 0;
						fee.nDiffPaid = grp_diffcost;
						fee.nPatPaid = 0;
						feeList.Add(fee);
				}

				if (bOutList)
				{
					fee.szID = _T("C");
					fee.szGroupID = _T("HITECH_ITEM_GROUP");
					fee.szName.Format(_T("T\x1ED5ng \x63\x1ED9ng(A+B)"));
					fee.szUnit.Empty();
					fee.nCost = ttl_amount;

					fee.nInsPaid = ttl_inspaid;
					fee.nDiscount = ttl_discount;
					fee.nQuantity = 0;
					fee.nPrice = 0;
					fee.nInsPrice = 0;
					fee.nDiffPaid = ttl_diffcost;
					fee.nPatPaid = ttl_patpaid;
					feeList.Add(fee);
					
					
				}
				

				szSQL.Format(_T("SELECT hfe_cost, hfe_inspaid, hfe_discount, hfe_patpaid+hfe_patdebt as hfe_patpaid, hfe_diffcost ") \
						_T("FROM hms_fee WHERE hfe_type='V' and hfe_docno=%ld and hfe_itemid='VT000002' and hfe_subgroup='%s'  "),
						nDocumentNo, szItemID);
					rsl.ExecSQL(szSQL);


					if(!rsl.IsEOF())
					{
						rsl.GetValue(_T("hfe_inspaid"), nAmount);
						ttl_inspaid = nAmount;
						ttl_amount = nAmount;
						rsl.GetValue(_T("hfe_discount"), nAmount);
						ttl_discount = nAmount;
						rsl.GetValue(_T("hfe_patpaid"), nAmount);
						rsl.GetValue(_T("hfe_diffcost"), nAmount);
						ttl_diffcost = nAmount;
						ttl_patpaid = ttl_amount-ttl_discount;
						
						fee.szID = _T("");
						fee.szGroupID = _T("HITECH_ITEM_GROUP");
						fee.szName.Format(_T("T\x1ED5ng s\x1ED1 ti\x1EC1n \x64\x1ECB\x63h v\x1EE5 s\x61u khi \xE1p m\x1EE9\x63 tr\x1EA7n \x42H\x58H"));
						fee.szUnit.Empty();
						fee.nCost = ttl_amount;

						fee.nInsPaid = ttl_inspaid;
						fee.nDiscount = ttl_discount;
						fee.nQuantity = 0;
						fee.nPrice = 0;
						fee.nInsPrice = 0;
						fee.nDiffPaid = ttl_diffcost;
						fee.nPatPaid = ttl_patpaid;
						feeList.Add(fee);


					}

					nTotalCost += ttl_amount;
					nTotalInsCost += ttl_inspaid;
					nTotalDiscount += ttl_discount;
					nTotalDiffPaid += ttl_diffcost;
					nTotalPatPaid += ttl_patpaid;
					ttl2_amount += ttl_amount;
					ttl2_inspaid += ttl_inspaid;
					ttl2_discount += ttl_discount;
					ttl2_diffcost += ttl_diffcost;
					ttl2_patpaid += ttl_patpaid;


			rs.MoveNext();
		}
		
	

		if(ttl2_amount > ttl_amount)
		{
			fee.szID = _T("");
			fee.szGroupID = _T("HITECH_ITEM_GROUP");
			fee.szName.Format(_T("T\x1ED5ng s\x1ED1 ti\x1EC1n \x64\x1ECB\x63h v\x1EE5"));
			fee.szUnit.Empty();
			fee.nCost = ttl2_amount;

			fee.nInsPaid = ttl2_inspaid;
			fee.nDiscount = ttl2_discount;
			fee.nQuantity = 0;
			fee.nPrice = 0;
			fee.nInsPrice = 0;
			fee.nDiffPaid = ttl2_diffcost;
			fee.nPatPaid = ttl2_patpaid;
			feeList.Add(fee);
		}


	}

	

	nSumCost += nTotalCost;
	nSumInsCost += nTotalInsCost;
	nSumDiscount += nTotalDiscount;
	nSumDiffPaid += nTotalDiffPaid;
	nSumPatPaid += nTotalPatPaid;



	for (int i =0; i < feeList.GetCount(); i++){
			fee = feeList[i];
			szType = fee.szGroupID;

			if(szType == _T("Type") || szType == _T("Section"))
			{
				grp = rpt.GetGroupHeader(1);
				rptDetail = rpt.AddDetail(grp);
				tmpStr = fee.szID;
				rptDetail->SetValue(_T("TypeID"), tmpStr);
				StringUpper(fee.szName, tmpStr);
				rptDetail->SetValue(_T("TypeName"), tmpStr);
			}
			else if(szType == _T("Group")){
				grp = rpt.GetGroupHeader(1);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("TypeID"), fee.szID);
				rptDetail->SetValue(_T("TypeName"), fee.szName);
			}
			else if(szType == _T("SubGroup")){
				grp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("SubGroupID"), fee.szID);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);
				
				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupInsCost"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);
				
				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);
				

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupPatPaid"), tmpStr);


				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("SubGroupInsUnPaid"), tmpStr);
				}
			}
			else if(szType == _T("SubAmount")){
				grp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(grp);
				CReportItem *obj ;
				obj = rptDetail->GetItem(_T("SubGroupName")); if(obj) obj->SetBold(true);
				obj = rptDetail->GetItem(_T("SubGroupCost")); if(obj) obj->SetBold(true);
				obj = rptDetail->GetItem(_T("SubGroupDiscount")); if(obj) obj->SetBold(true);
				obj = rptDetail->GetItem(_T("SubGroupPatpaid")); if(obj) obj->SetBold(true);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);


				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("SubGroupInsUnPaid"), tmpStr);
				}

			}
			else if(szType == _T("GrandAmount")){
				grp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(grp);
			//	rptDetail->GetItem(_T("SubGroupName"))->SetBold(true);
			//	rptDetail->GetItem(_T("SubGroupCost"))->SetBold(true);
			//	rptDetail->GetItem(_T("SubGroupDiscount"))->SetBold(true);
			//	rptDetail->GetItem(_T("SubGroupPatpaid"))->SetBold(true);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("SubGroupInsUnPaid"), tmpStr);
				}

			}

			else if(szType == _T("Item")){
				rptDetail = rpt.AddDetail();
				CReportItem* obj  = rptDetail->GetItem(_T("1"));
				if(obj){
					
					obj->SetTextAlign(ES_RIGHT);
				}
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), fee.szUnit);
				FormatCurrency(fee.nQuantity, tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("11"), tmpStr);
				}
				
			}
			else if(szType == _T("SubItemGroup")){
				rptDetail = rpt.AddDetail();
				CReportItem *obj = NULL;
				for (int i = 1; i <= 9; i++)
				{
					obj = rptDetail->GetItem(i);
					if(obj) 
					{
						obj->SetItalic(true);
						obj->SetBold(true);
					}

				}
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), _T(""));
				rptDetail->SetValue(_T("4"), _T(""));
				rptDetail->SetValue(_T("5"), _T(""));
				rptDetail->SetValue(_T("6"), _T(""));
				rptDetail->SetValue(_T("7"), _T(""));
				rptDetail->SetValue(_T("8"), _T(""));
				rptDetail->SetValue(_T("9"), _T(""));
			}

			else if(szType == _T("SubItem")){
				rptDetail = rpt.AddDetail();
				CReportItem *obj = NULL;
				for (int i = 1; i <= 9; i++)
				{
					obj = rptDetail->GetItem(i);
					if(obj) obj->SetItalic(true);
				}

				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), fee.szUnit);
				FormatCurrency(fee.nQuantity, tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("11"), tmpStr);
				}
				
			}
			else if(szType == _T("SubItemAmount")){
				rptDetail = rpt.AddDetail();
				CReportItem *obj = NULL;
				for (int i = 1; i <= 9; i++)
				{
					obj = rptDetail->GetItem(i);
					if(obj) 
					{
						obj->SetItalic(true);
						obj->SetBold(true);
					}

				}
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), _T(""));
				rptDetail->SetValue(_T("4"), _T(""));
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), _T(""));
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("11"), tmpStr);
				}
				
			}

			else if(szType == _T("Dousage")){
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(4));
				rptDetail->SetValue(_T("usage"), fee.szName);
			}
			else if(szType == _T("TotalAmount")){
				grp = rpt.GetGroupHeader(3);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("TotalAmountLabel"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("TotalAmount"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("TotalInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("TotalDiscount"), tmpStr);
				
				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("TotalDiffPaid"), tmpStr);
				//fee.nPatPaid= fee.nCost-fee.nDiscount;
				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("TotalPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("TotalInsUnPaid"), tmpStr);
				}
			}
			else if(szType == _T("HITECH_HEADER"))
			{
				CReportSection *rptGrp = rpt.GetGroupFooter(1);
				rptDetail = rpt.AddDetail(rptGrp);
			}
			else if(szType == _T("HITECH_TYPE"))
			{
				CReportSection *rptGrp = rpt.GetGroupHeader(1);
				rptDetail = rpt.AddDetail(rptGrp);
				rptDetail->SetValue(_T("TypeID"), fee.szID);
				rptDetail->SetValue(_T("TypeName"), fee.szName);


			}
			else if(szType == _T("HITECH_GROUP"))
			{
				CReportSection *rptGrp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(rptGrp);
				rptDetail->SetValue(_T("SubGroupID"), fee.szID);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				tmpStr.Format(_T("%.2f"), fee.nCost);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);
				tmpStr.Format(_T("%.2f"), fee.nDiscount);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);
				tmpStr.Format(_T("%.2f"), fee.nDiffPaid);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);
				tmpStr.Format(_T("%.2f"), fee.nPatPaid);
				rptDetail->SetValue(_T("SubGroupPatpaid"), tmpStr);


			}
			else if(szType == _T("HITECH_ITEM"))
			{
				rptDetail = rpt.AddDetail();
				CReportItem* obj  = rptDetail->GetItem(_T("1"));
				if(obj){
					
					obj->SetTextAlign(ES_RIGHT);
				}
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), fee.szUnit);
				FormatCurrency(fee.nQuantity, tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("11"), tmpStr);
				}
				

			}
			else if(szType == _T("HITECH_ITEM_GROUP"))
			{
				grp = rpt.GetGroupHeader(3);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("TotalAmountLabel"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("TotalAmount"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("TotalInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("TotalDiscount"), tmpStr);
				
				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("TotalDiffPaid"), tmpStr);
				//fee.nPatPaid= fee.nCost-fee.nDiscount;
				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("TotalPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("TotalInsUnPaid"), tmpStr);
				}
				

			}
		}	
	

	if(nTotalCost != nSumCost){
		CString szName;
		TranslateString(_T("Total Amount"), tmpStr);
		szName.Format(_T("%s"), tmpStr);
			grp = rpt.GetGroupHeader(3);
			rptDetail = rpt.AddDetail(grp);
			rptDetail->SetValue(_T("TotalAmountLabel"), szName);
			FormatCurrency(nSumCost, tmpStr);
			rptDetail->SetValue(_T("TotalAmount"), tmpStr);

			FormatCurrency(nSumInsCost, tmpStr);
			rptDetail->SetValue(_T("TotalInsCost"), tmpStr);

			FormatCurrency(nSumDiscount, tmpStr);
			rptDetail->SetValue(_T("TotalDiscount"), tmpStr);
				
			FormatCurrency(nSumDiffPaid, tmpStr);
			rptDetail->SetValue(_T("TotalDiffPaid"), tmpStr);
			FormatCurrency(nSumPatPaid, tmpStr);
			rptDetail->SetValue(_T("TotalPatPaid"), tmpStr);
	}



	

	
	szSQL.Format(_T(" SELECT hfe_amount as hfe_amount,") \
_T("   hfe_inspaid,") \
_T("   hfe_discount,") \
_T("   hfe_patpaid,") \
_T("   hfe_payment,") \
_T("   hfe_diffcost,") \
_T("   hfe_diffpaid,") \
_T("   hfe_deposit,") \
_T("   hfe_refund,") \
_T("   hfe_freeamount") \
_T(" FROM hms_fee_invoice") \
_T(" WHERE hfe_invoiceno=%ld and hfe_status='P' "), nInvoiceNo);

	rs.ExecSQL(szSQL);
	if(!rs.IsEOF())
	{
		

		rs.GetValue(_T("hfe_amount"), nTotalCost);
		rs.GetValue(_T("hfe_inspaid"), nTotalInsCost);
		rs.GetValue(_T("hfe_discount"), nTotalDiscount);
		rs.GetValue(_T("hfe_patpaid"), nTotalPatPaid);
		rs.GetValue(_T("hfe_diffcost"), nTotalDiffPaid);
		rs.GetValue(_T("hfe_deposit"), nTotalDeposit);
		rs.GetValue(_T("hfe_refund"), nTotalRefund);
		rs.GetValue(_T("hfe_freeamount"), nTotalFree);
		rs.GetValue(_T("hfe_payment"), nTotalPayment);
		
		nTotalPayable = nTotalPatPaid;
		

	}
	else
	{
		nTotalCost += nTotalDiffPaid;
		szSQL.Format(_T("SELECT sum(hfe_amount-hfe_patpaid) as amount ") \
			_T("FROM hms_fee_deposit ") \
			_T("WHERE hfe_docno=%ld and hfe_status<>'C' and hfe_class IN('E','I','A') and hfe_objectid<>7"), 
			nDocumentNo);
		rs.ExecSQL(szSQL);
		if(!rs.IsEOF())
		{
			rs.GetValue(_T("amount"), nTotalDeposit);
		}
		nTotalPayable = nTotalPatPaid+nTotalDiffPaid;
		nTotalPayment = nTotalPayable-(nTotalDeposit+nTotalFree);
		
		nTotalRefund = 0;
		if(nTotalPayment < 0)
		{
			nTotalRefund = -1*nTotalPayment;
			nTotalPayment = 0;
		}
		//_msg(_T("%f:%f"), nTotalPayment, nTotalRefund);
	}

	//In lai so ngay dieu tri
	if (nNumberTreat > 0)
	{
		tmpStr.Format(_T("%d"), nNumberTreat);
		rpt.GetReportHeader()->SetValue(_T("TreatmentDay"), tmpStr);
	}

	/*nTotalCost = ceil(nTotalCost);
	nTotalInsCost = ceil(nTotalInsCost);
	nTotalDiscount = ceil(nTotalDiscount);
	nTotalPayable = ceil(nTotalPayable);
	nTotalDiffPaid = ceil(nTotalDiffPaid);*/

	nTotalPayable = ceil(nTotalPayable);

	rpt.GetReportFooter()->SetValue(_T("PatientSign"), szPatientName);
		FormatCurrency((nTotalCost), tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumAmount"), tmpStr);
		CString szSumInWord;
		tmpStr.Format(_T("%.0f"), nTotalCost);
		MoneyToString(tmpStr, szSumInWord);
		szSumInWord += _T(" \x111\x1ED3ng.");
		rpt.GetReportFooter()->SetValue(_T("SumInWord"), szSumInWord);
		
		FormatCurrency(nTotalInsCost, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumInsCost"), tmpStr);


		FormatCurrency(nTotalDiscount, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumInsDiscount"), tmpStr);

		FormatCurrency(nTotalInsCost-nTotalDiscount, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumInsUnPaid"), tmpStr);

		FormatCurrency((nTotalDiffPaid), tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumDiffPaid"), tmpStr);


		FormatCurrency(nTotalPayable, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumPatPaid"), tmpStr);



		FormatCurrency((nTotalDeposit), tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumDeposit"), tmpStr);
		
		FormatCurrency(nTotalFree, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumDiscount"), tmpStr);

		FormatCurrency(nSumFoodAmount, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumFoodAmount"), tmpStr);

		nTotalRefund = ceil(nTotalRefund);
		nTotalPayment = ceil(nTotalPayment);

		if(nTotalRefund > 0)
		{
		//	FormatCurrency(nTotalRefund, tmpStr);
		//	rpt.GetReportFooter()->SetValue(_T("SumPayment"), tmpStr);

			FormatCurrency(nTotalRefund, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("SumRefund"), tmpStr);
		}
		else
		{
			FormatCurrency(nTotalPayment, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("SumPayment"), tmpStr);

		}

	

//	szSQL.Format(_T("UPDATE hms_fee_invoice SET hfe_print=hfe_print+1 WHERE hfe_invoiceno=%ld "), nInvoiceNo);
//	ExecSQL(szSQL);

//	rpt.Print();
	rpt.PrintPreview();

}



void CPrintForms::FM_PrintPolicyDischargePaymentInvoice(long nDocumentNo, long nInvoiceNo)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
//	if(!pMF->IsInPatient())
//		return;
	CReport rpt;
	CString tmpStr, szSQL, szWhere, szAdmitDate, szEndDate,szDepartments, szAddonedayoutofhospital,szNumbertreat;
	CRecord rs(&pMF->m_db);
	CRecord rsd(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	int		nDeskNo=0;
	CString szRecvDate;
	CString	szObjectType;
	CString	szDescription;
	CString szPrintMaterialOperation = _T("Y");
	CString szOutPatient;


	double	nTotalCost=0,				//Tong chi phi
			nTotalInsCost=0,
			nTotalDiscount = 0,			//Tong Mien giam
			nTotalPatPaid =0,			//Tong benh nhan tra
			nTotalDiffPaid = 0,	//Tong so tra chenh lech
			nTotalPayable=0,
			nTotalDeposit = 0,			//Tong tam ung
			nTotalFree = 0,				//Tong mien phi
			nTotalRefund=0,				//Tong hoan tra
			nDepositPayable=0;			//So tien tra tu tam ung
	
	
	double nSumFoodAmount=0;			


	szWhere.AppendFormat(_T(" and hfe_invoiceno=%ld "), nInvoiceNo);

	szSQL.Format(_T(" SELECT *") \
				_T(" FROM") \
				_T(" (") \
				_T("  SELECT hd_patientno as patientno,") \
				_T("         hd_docno as docno,") \
				_T("         trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("         hp_birthdate as birthdate,") \
				_T("         hms_getage(trunc_date(hd_admitdate), hp_birthdate) as age,") \
				_T("         extract(year from hp_birthdate) as yearofbirth, ") \
				_T("         Get_Selection_Desc('sys_sex', hp_sex) as sex,") \
				_T("         hp_sex as sex_id,") \
				_T("         hp_ethnic as ethnic,") \
				_T("         Get_Selection_Desc('sys_ethnic', hp_ethnic) as ethnicdesc, ") \
				_T("         hp_dtladdr as detailaddress,") \
				_T("         hms_getaddress(hp_provid,hp_distid, hp_villid) as address,") \
				_T("         hp_workplace as workplace,") \
				_T("		 hd_rank, get_selection_desc('hms_rank', hd_rank) as rankname, ") \
				_T("         hd_doctor as doctor,") \
				_T("         hd_createdby as createdby,") \
				_T("         hd_icd as icd10,") \
				_T("         hd_diagnostic as diagnostic,") \
				_T("         hd_reldisease as reldisease,") \
				_T("         hd_xobject as xobject, ") \
				_T("         hd_xcardno as xcardno, ") \
				_T("         hd_xissuedate as xissuedate, ") \
				_T("         hd_emergency as emergency, ") \
				_T("         Get_Selection_Desc('hms_result',hd_result) as result,") \
				_T("         hd_admitdate as examdate,") \
				_T("         hd_enddate as enddate,") \
				_T(" hd_outpatient, ") \
				_T("         hcr_recordno as recordno ,") \
				_T("         hcr_admitdate as admitdate,") \
				_T("         hcr_dischargedate as dischargedate,") \
				_T("         hcr_sumtreat as TreatmentDay,") \
				_T("         hcr_mainicd as mainicd ,") \
				_T("         hcr_maindisease as maindisease ,") \
				_T("         hcr_treatdoctor as treatdoctor, ") \
				_T("         hcr_reldisease as ireldisease,") \
				_T("         hcr_suggestion as isuggestion,") \
				_T("         hd_suggestion as esuggestion,") \
				_T("         Get_Selection_Desc('hms_result' ,hcr_result)  as iresult,") \
				_T("         Hms_GetObjectType(hd_object) as objecttype, ") \
				_T("         cast(hd_disrate as integer) as disrate, ") \
				_T("         hd_insline as insline, ") \
				_T("         hd_insregdate as insregdate, ") \
				_T("         hd_transplace as transplace, ") \
				_T("         hc_cardno as cardno, ") \
				_T("         hc_regdate as regdate, ") \
				_T("         hc_expdate as expdate, ") \
				_T("         hc_regcode as regcode, ") \
				_T("         HMS_GETHOSPITALNAME(hc_regcode) as regplace, ") \
				_T("         hfe_deptid as deptid,") \
				_T("         hfe_serialno as serialno,") \
				_T("         hfe_receiptno as recvno,") \
				_T("         hfe_date as recvdate, ") \
				_T("         hfe_staff as receiver,") \
				_T("         hfe_desc as reason,") \
				_T("         hfe_amount as cost,") \
				_T("         hfe_discount as discount,") \
				_T("         hfe_patpaid as patpaid,") \
				_T("         hfe_deposit as deposit_amount, ") \
				_T("         0 as deposit_payable,") \
				_T("         hfe_refund as refund_amount, ") \
				_T("         hfe_freeamount as free_amount,") \
				_T("         hfe_eat_amount as food_amount,") \
				_T("         row_number() over (partition by hd_docno order by hd_docno) as dn") \
				_T(" FROM hms_fee_invoice ") \
				_T(" LEFT JOIN hms_doc ON(hd_docno=hfe_docno)") \
				_T(" LEFT JOIN hms_patient ON(hp_patientno=hd_patientno)") \
				_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno) ") \
				_T(" LEFT JOIN hms_treatment_record ON(hfe_docno=htr_docno) ") \
				_T(" LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx) ") \
				_T(" WHERE hfe_docno=%ld %s ") \
				_T(" ) Tbl") \
				_T(" WHERE dn=1"), nDocumentNo, szWhere);
	
	_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		return;
	}

	rs.GetValue(_T("cost"), nTotalCost);
	rs.GetValue(_T("discount"), nTotalDiscount);
	rs.GetValue(_T("patpaid"), nTotalPatPaid);
	rs.GetValue(_T("deposit_amount"), nTotalDeposit);
	rs.GetValue(_T("free_amount"), nTotalFree);
	rs.GetValue(_T("refund_amount"), nTotalRefund);
	rs.GetValue(_T("deposit_payable"), nDepositPayable);
	rs.GetValue(_T("reason"), szDescription);
	rs.GetValue(_T("deskno"), nDeskNo);

	rs.GetValue(_T("food_amount"), nSumFoodAmount);

	rs.GetValue(_T("objecttype"), szObjectType);
	rs.GetValue(_T("hd_outpatient"), szOutPatient);


	if(szObjectType == _T("I") || szObjectType == _T("C"))
	{
		if(!rpt.Init(_T("Reports/HMS/HF_INSDISCHARGEPAYMENT.RPT"), false) )
			return;
	}
	else if(szObjectType == _T("D"))
	{
		if(!rpt.Init(_T("Reports/HMS/HF_FREEDISCHARGEPAYMENT.RPT"), false) )
			return;
	}
	else
	{
		if(!rpt.Init(_T("Reports/HMS/HF_DISCHARGEPAYMENT.RPT"), false) )
			return;
	}

	//Report Header
	tmpStr.Empty();
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);

	rs.GetValue(_T("docno"), nDocumentNo);
	tmpStr.Format(_T("%ld"), nDocumentNo);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	tmpStr.Format(_T("%s-%s-%s"), rs.GetValue(_T("serialno")), rs.GetValue(_T("bookno")),  rs.GetValue(_T("recvno")));
	rpt.GetReportHeader()->SetValue(_T("ReceiptNo"), tmpStr);
	tmpStr.Format(_T("%ld"), nInvoiceNo);
	rpt.GetReportHeader()->SetValue(_T("InvoiceNo"), tmpStr);


	CString szDate;
	tmpStr = pMF->GetSysDateTime();
	szDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")), tmpStr.Mid(11, 5),tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), szDate);


	rs.GetValue(_T("recordno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RecordNo"), tmpStr);
	//rs.GetValue(_T("pname"), tmpStr);
	

	CString szPatientName;
	StringUpper(rs.GetValue(_T("pname")), szPatientName);
	rpt.GetReportHeader()->SetValue(_T("PATIENTNAME"), szPatientName);
	rs.GetValue(_T("age"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), tmpStr);
	rs.GetValue(_T("yearofbirth"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);

	rs.GetValue(_T("birthdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("BirthDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

	rs.GetValue(_T("sex"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Sex"), tmpStr);

	rs.GetValue(_T("sex_id"), tmpStr);
	if(tmpStr == _T("F"))
		rpt.GetReportHeader()->SetValue(_T("Female"), _T("X"));
	else
		rpt.GetReportHeader()->SetValue(_T("Male"), _T("X"));


	rs.GetValue(_T("ethnic"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ethnic"), tmpStr);
	rs.GetValue(_T("ethnicdesc"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ethnicdesc"), tmpStr);
	
	rs.GetValue(_T("workplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Workplace"), tmpStr);


	rs.GetValue(_T("rankname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Rank"), tmpStr);

	CString szAddress;
	rs.GetValue(_T("detailaddress"), tmpStr);
	szAddress = tmpStr;
	rs.GetValue(_T("address"), tmpStr);
	if(!szAddress.IsEmpty())
		szAddress += _T(" - ");
	szAddress += tmpStr;
	rpt.GetReportHeader()->SetValue(_T("Address"), szAddress);
	

	rs.GetValue(_T("transplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("TransferHospital"), tmpStr);
	
	
	rs.GetValue(_T("objecttype"), szObjectType);
	
	rs.GetValue(_T("cardno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CardNo"), tmpStr.Left(15));
	if(!tmpStr.IsEmpty())
	{
		rpt.GetReportHeader()->SetValue(_T("HasCardNo"), _T("X"));
	}
	else
	{
		rpt.GetReportHeader()->SetValue(_T("NoCardNo"), _T("X"));
	}

	rs.GetValue(_T("disrate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Insrate"), tmpStr);
	

	rs.GetValue(_T("insline"), tmpStr);	
	if (tmpStr == _T("Y"))
	{		
		rpt.GetReportHeader()->SetValue(_T("InsLine"), _T("X"));
		rpt.GetReportHeader()->SetValue(_T("NotInsLine"), _T(""));
	}
	else
	{
		rpt.GetReportHeader()->SetValue(_T("InsLine"), _T(""));
		rpt.GetReportHeader()->SetValue(_T("NotInsLine"), _T("X"));
	}
	
	rs.GetValue(_T("insregdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("InsRegDate"), CDate::Convert(tmpStr));


	rs.GetValue(_T("regdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RegDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));

	rs.GetValue(_T("expdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ExpiryDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	rs.GetValue(_T("regplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RegistrationPlace"), tmpStr);


	CString szRegCode;
	rs.GetValue(_T("regcode"), szRegCode);
	rpt.GetReportHeader()->SetValue(_T("RegCode"), szRegCode);

	CReportItem *obj = rpt.GetReportHeader()->GetItem(_T("InsuranceCompany"));
	if (obj)
	{
		CString insname;
		CRecord rsx(&pMF->m_db);
		
		tmpStr.Format(_T("select sp_name as insname ") \
			          _T("from sys_prov left join hms_hospital on(hh_provid=sp_id) where trim(hh_id)=trim('%s') "), szRegCode);
		rsx.ExecSQL(tmpStr);
		if (!rsx.IsEOF())
		{
			rsx.GetValue(_T("insname"), tmpStr);
			rpt.GetReportHeader()->Format(_T("InsuranceCompany"), tmpStr);
		}
	}

	rs.GetValue(_T("emergency"), tmpStr);
	if (tmpStr == _T("Y") )
		tmpStr = _T("X");
	else
		tmpStr.Empty();

	rpt.GetReportHeader()->SetValue(_T("Emergency"), tmpStr);

	rs.GetValue(_T("maindisease"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);
		rs.GetValue(_T("ireldisease"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("RelationDisease"), tmpStr);
		rs.GetValue(_T("mainicd"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("ICD10"), tmpStr);
		rs.GetValue(_T("iresult"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Result"), tmpStr);
		rs.GetValue(_T("isuggestion"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Suggestion"), pMF->GetSelectionString(_T("hms_suggestion"), tmpStr));
	
	
	
	CString szDischargeDate;
	rs.GetValue(_T("admitdate"), tmpStr);

	rpt.GetReportHeader()->SetValue(_T("admitdate"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	
	rs.GetValue(_T("dischargedate"), tmpStr);
	szDischargeDate = tmpStr;
	
	if(szDischargeDate.Left(4) != _T("1752"))
			rpt.GetReportHeader()->SetValue(_T("dischargedate"), CDateTime::Convert(tmpStr, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	else
			rpt.GetReportHeader()->SetValue(_T("dischargedate"), _T(""));

	rs.GetValue(_T("TreatmentDay"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("TreatmentDay"), tmpStr);	

	CString szData;
	tmpStr.Format(_T("%s-%s-%s"), rs.GetValue(_T("bookno")), rs.GetValue(_T("serialno")), rs.GetValue(_T("recvno")));
	rpt.GetReportFooter()->SetValue(_T("SerialNo"), tmpStr);
	rs.GetValue(_T("recvdate"), szRecvDate);	
	rpt.GetReportFooter()->SetValue(_T("ReceiptDate"), CDateTime::Convert(szRecvDate, yyyymmdd|hhmm,ddmmyyyy|hhmm));
	szData.Format(rpt.GetReportFooter()->GetValue(_T("ReceiptDate1")), szRecvDate.Mid(11,5),szRecvDate.Mid(8,2),szRecvDate.Mid(5,2),szRecvDate.Left(4));	
	rpt.GetReportFooter()->SetValue(_T("ReceiptDate1"),szData);
	szData.Format(rpt.GetReportFooter()->GetValue(_T("ReceiptDate2")), szRecvDate.Mid(11,5),szRecvDate.Mid(8,2),szRecvDate.Mid(5,2),szRecvDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("ReceiptDate2"),szData);
	
	rs.GetValue(_T("treatdoctor"), tmpStr);
	tmpStr.Trim();
	rpt.GetReportFooter()->SetValue(_T("Doctor"), GetDoctorName(tmpStr));

	CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
	if(img)
	{
		img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img->SetFixedHeight(true);
	}
	tmpStr.Empty();
	rs.GetValue(_T("receiver"), tmpStr);
	tmpStr.Trim();
	rpt.GetReportFooter()->SetValue(_T("ReceiverBy"), GetDoctorName(tmpStr));
	CReportItem *img2 = rpt.GetReportFooter()->GetItem(_T("Signature2"));
	if(img2)
	{
		img2->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img2->SetFixedHeight(false);
	}	
	
	
	// Nguoi tiep don
	tmpStr.Empty();
	rs.GetValue(_T("createdby"), tmpStr);
	tmpStr.Trim();
	rpt.GetReportFooter()->SetValue(_T("createdby"), GetDoctorName(tmpStr));
	CReportItem *img3 = rpt.GetReportFooter()->GetItem(_T("Signature3"));
	if(img3)
	{
		img3->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img3->SetFixedHeight(false);
	}	


	CReportSection* rptDetail=NULL;
	CReportSection *grp;
	CRecord grs(&pMF->m_db);
	CString szFeeID, szID;
	CString szSubItem, szType;

	TCHAR *lpszChapter[] = {_T("I"), _T("II"), _T("III"), _T("IV"), _T("V"), _T("VI"), _T("VII"), _T("VIII"), _T("IX"), _T("X"), _T("XI")};


	CRecord drs(&pMF->m_db);
	CString szDeptID;
		
		CArray<RECEIPTINFO, RECEIPTINFO& > payInfo;

		szSQL.Format(_T(" SELECT deptid, admitdate, dischargedate, mainicd, maindisease") \
					_T(" FROM") \
					_T(" (") \
					_T("  select htr_deptid as deptid,") \
					_T("         htr_admitdate as admitdate,") \
					_T("         htr_dischargedate as dischargedate,") \
					_T("         htr_mainicd as mainicd,") \
					_T("         htr_maindisease as maindisease,") \
					_T("         htr_idx as idx") \
					_T("  from hms_treatment_record") \
					_T("  where htr_docno=%ld") \
					_T("  union all") \
					_T("  select hd_admitdept,") \
					_T("         hd_admitdate,") \
					_T("         hd_enddate,") \
					_T("         hd_icd,") \
					_T("         hd_diagnostic,") \
					_T("         0") \
					_T("  from hms_doc") \
					_T("  where hd_docno=%ld") \
					_T(" ) tbl") \
					_T(" WHERE deptid in(select distinct hfe_deptid from hms_fee where hfe_invoiceno=%ld)") \
					_T(" ORDER BY idx"), nDocumentNo, nDocumentNo, nInvoiceNo);

		drs.ExecSQL(szSQL);

		RECEIPTINFO pi;
		CString szDepts;
		szDepts.Empty();

		while (!drs.IsEOF())
		{
			drs.GetValue(_T("deptid"), tmpStr);
			
			bool bFound = false;
			for (int i =0; i < payInfo.GetCount(); i++)
			{
				pi = payInfo[i];
				if (pi.deptid == tmpStr)
				{
					bFound = true;
					break;
				}
			}
			pi.deptid = tmpStr;
			if(!szDepts.IsEmpty())
				szDepts += _T(",  ");
			szDepts.AppendFormat(_T("%s"), tmpStr);

			drs.GetValue(_T("dischargedate"), pi.dischargedate);
			drs.GetValue(_T("mainicd"), pi.mainicd);
			drs.GetValue(_T("maindisease"), pi.maindisease);

			if (bFound)
			{
				payInfo.SetAt(i, pi);
			}
			else
			{
				drs.GetValue(_T("admitdate"), pi.admitdate);
				payInfo.Add(pi);
				
			}

			drs.MoveNext();
		}

		rpt.GetReportHeader()->SetValue(_T("Department"), szDepts);	

		double nSumCost=0, nSumInsCost=0, nSumDiscount=0, nSumDiffPaid=0, nSumPatPaid=0;

		szDepts.Empty();


		int nChapter = 0;
		int nCount = 0, nIndex, nChapterID = 0;
		int nItem = 0, nOldItem = 0, nItem2 = 0;
		bool bDeleteParent = false;
		bool bLoadItems = false;
		CString szNewGroup, szOldGroup, szParentGroup;

		
	double nGroup0Cost=0, nGroup0InsCost= 0, nGroup0Discount=0, nGroup0DiffPaid=0, nGroup0PatPaid=0;
	double nGroup1Cost=0, nGroup1InsCost = 0, nGroup1Discount=0, nGroup1DiffPaid = 0, nGroup1PatPaid=0;
	double nGroup2Cost=0, nGroup2InsCost = 0, nGroup2Discount=0, nGroup2DiffPaid = 0, nGroup2PatPaid=0;	
	double nCost=0, nInsCost =0,  nDiscount=0, nDiffPaid  = 0, nPatPaid=0;

	
	nTotalCost = nTotalDiffPaid = nTotalInsCost = nTotalPatPaid = nTotalDiscount =0;

	szSQL.Format(_T("SELECT * FROM hms_fee_type ORDER BY hft_id "));
	grs.ExecSQL(szSQL);
	
	szWhere.Format(_T(" AND hfe_invoiceno=%ld"), nInvoiceNo);

	if (szOutPatient != _T("Y")) 
		szWhere.AppendFormat(_T(" and hfe_groupid <>'D0000' "));
	szWhere.AppendFormat(_T(" and hfe_class in('E', 'I','X') "));
	//szWhere.AppendFormat(_T(" and hfe_subgroup='XX' "));
	szWhere.AppendFormat(_T(" and (hfe_category <> 'Y' or hfe_category is null) "));


	szSQL.Format(_T(" SELECT hfe_typeindex AS typeindex,") \
_T("   hfe_status         AS status,") \
_T("   hfe_type           AS feetype,") \
_T("   hfe_groupid        AS groupid,") \
_T("   hfe_groupid3       AS groupid3,") \
_T("   hfe_itemid         AS itemid,") \
_T("   hfe_name           AS name,") \
_T("   hfe_unit           AS unit,") \
_T("   hfe_hitech         AS hitech,") \
_T("   hfe_inlist         AS inlist,") \
_T("   hfe_unitprice      AS unitprice,") \
_T("   SUM(hfe_quantity)       AS qty,") \
_T("   SUM(hfe_amount)      AS cost,") \
_T("   SUM(hfe_discount)  AS discount,") \
_T("   SUM(hfe_diffpaid)  AS DiffPaid,") \
_T("   SUM(hfe_copayment) AS copayment,") \
_T("   SUM(hfe_patpaid)   AS patpaid") \
_T(" FROM hms_fee_invoiceline_view ") \
_T(" WHERE hfe_docno  =%ld %s ") \
_T(" GROUP BY hfe_typeindex,") \
_T("   hfe_status ,") \
_T("   hfe_type ,") \
_T("   hfe_groupid ,") \
_T("   hfe_groupid3 ,") \
_T("   hfe_itemid ,") \
_T("   hfe_name ,") \
_T("   hfe_unit ,") \
_T("   hfe_hitech ,") \
_T("   hfe_inlist,") \
_T("   hfe_unitprice") \
_T(" ORDER BY hfe_typeindex, hfe_inlist, hfe_groupid3, hfe_name, hfe_unit, hfe_unitprice"), nDocumentNo, szWhere);



_fmsg(_T("%s"), szSQL);		
	
	nIndex = 1;
	int nSubItem = 1;
	int nIDX;
	nChapterID = 0;
	CArray<FEEITEM, FEEITEM&> feeList;
	FEEITEM fee;
	CString szName;
	CString szNewIndex, szOldIndex;
	CString szIndex;
	bool bFound = false;
	bool bInList = false;
	bool bOutList = false;
	bool bKList = false;
	int nSubIndex = 1;
	
	double nTtlCost=0, nTtlInsCost=0, nTtlDiscount=0, nTtlDiffPaid=0, nTtlPatPaid=0;
	rs.ExecSQL(szSQL);
_fmsg(_T("%s"), szSQL);

	

	while(!grs.IsEOF()){
		tmpStr.Format(_T("%d"), nIndex);
		grs.GetValue(_T("hft_name"), szName);
		grs.GetValue(_T("hft_id"), szNewIndex);
		
		fee.szGroupID = _T("Type");
		fee.szID = tmpStr;
		fee.szName = szName;
		nItem = feeList.Add(fee);
		
		nIDX  = nItem;
		bFound = false;
		nGroup1Cost=nGroup1InsCost=nGroup1Discount=nGroup1DiffPaid=nGroup1PatPaid=0;
		nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;

		rs.MoveFirst();
		while(!rs.IsEOF())
		{
			rs.GetValue(_T("typeindex"), tmpStr);
			
			if(tmpStr == szNewIndex){
				bFound = true;		
				rs.GetValue(_T("groupid"), szNewGroup);
				if(szNewGroup != szOldGroup)
				{
					szOldGroup = szNewGroup;
					if(szNewGroup.Left(2) == _T("B1")){
						szSQL.Format(_T("SELECT hfg_name FROM hms_fee_group WHERE rtrim(hfg_id,'0')='%s' "), szNewGroup);
						CRecord rs2(&pMF->m_db);
						rs2.ExecSQL(szSQL);
						if(nGroup2Cost > 0){
							TranslateString(_T("Sub Amount"), tmpStr);
							fee.szGroupID = _T("SubAmount");
							fee.szName = tmpStr;
							fee.nCost = nGroup2Cost;
							fee.nInsPaid = nGroup2InsCost;
							fee.nDiscount = nGroup2Discount;
							fee.nDiffPaid = nGroup2DiffPaid;
							fee.nPatPaid = nGroup2PatPaid;
							int nItem2 = feeList.Add(fee);
						}

						szIndex.Format(_T("%d.%d"), nIndex, nSubIndex++);
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = rs2.GetValue(_T("hfg_name"));
						int nItem2 = feeList.Add(fee);
						nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;

					}
					
				}

				if(szOldIndex != szNewIndex)
				{
					szOldIndex = szNewIndex;
					bInList = false;
					bOutList = false;
				}

				

				if(szNewIndex == _T("800")){
					

					rs.GetValue(_T("inlist"), tmpStr);
					
					if(tmpStr == _T("1") && !bInList){
						bInList = true;
						TranslateString(_T("Inside insurance list"), tmpStr);
						szIndex.Format(_T("%d.1"), nIndex);	
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = tmpStr;
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nDiffPaid = 0;
						fee.nPatDebt = 0;

						int nItem2 = feeList.Add(fee);

					}
					if(tmpStr == _T("2") && !bOutList){
						bOutList = true;
						TranslateString(_T("Outside insurance list"), tmpStr);
						if(bInList)
							szIndex.Format(_T("%d.2"), nIndex);	
						else
							szIndex.Format(_T("%d.1"), nIndex);	

						szIndex.Format(_T("%d.1"), nIndex);	
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = tmpStr;
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nDiffPaid = 0;
						fee.nPatDebt = 0;
						int nItem2 = feeList.Add(fee);
					}

					if(tmpStr == _T("3") && !bKList){
						bKList = true;
						if(bInList && bOutList)
							szIndex.Format(_T("%d.3"), nIndex);	
						else if(bInList || bOutList)
							szIndex.Format(_T("%d.2"), nIndex);	
						else
							szIndex.Format(_T("%d.1"), nIndex);	

						TranslateString(_T("Inside cancer list"), tmpStr);
						szIndex.Format(_T("%d.1"), nIndex);	
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = tmpStr;
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nDiffPaid = 0;
						fee.nPatDebt = 0;
						int nItem2 = feeList.Add(fee);
					}
					

				}

				if(szNewIndex == _T("900")){
					
					rs.GetValue(_T("inlist"), tmpStr);
					
					if(tmpStr == _T("1") && !bInList){
						bInList = true;
						TranslateString(_T("Inside insurance list"), tmpStr);
						szIndex.Format(_T("%d.1"), nIndex);
						szIndex.Format(_T("%d.1"), nIndex);	
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = tmpStr;
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nDiffPaid = 0;
						fee.nPatDebt = 0;
						int nItem2 = feeList.Add(fee);
					}
					if(tmpStr == _T("2") && !bOutList){
						bOutList = true;
						TranslateString(_T("Outside insurance list"), tmpStr);
						if(!bInList)
							szIndex.Format(_T("%d.1"), nIndex);
						else
							szIndex.Format(_T("%d.2"), nIndex);

						szIndex.Format(_T("%d.1"), nIndex);	
						fee.szGroupID = _T("Group");
						fee.szID = szIndex;
						fee.szName = tmpStr;
						fee.nCost = 0;
						fee.nInsPaid = 0;
						fee.nDiscount = 0;
						fee.nDiffPaid = 0;
						fee.nPatDebt = 0;
						int nItem2 = feeList.Add(fee);
					}
				}

				



				rs.GetValue(_T("cost"), nCost);
				rs.GetValue(_T("inscost"), nInsCost);
				rs.GetValue(_T("discount"), nDiscount);
				rs.GetValue(_T("DiffPaid"), nDiffPaid);
				rs.GetValue(_T("copayment"), nPatPaid);
				
				nTotalCost += nCost;
				nGroup1Cost += nCost;
				nGroup2Cost += nCost;

				nTotalInsCost += nInsCost;
				nGroup1InsCost += nInsCost;
				nGroup2InsCost += nInsCost;

				nTotalDiscount += nDiscount;
				nGroup1Discount += nDiscount;
				nGroup2Discount += nDiscount;

				nTotalDiffPaid += nDiffPaid;
				nGroup1DiffPaid += nDiffPaid;
				nGroup2DiffPaid += nDiffPaid;

				nTotalPatPaid += nPatPaid;
				nGroup1PatPaid += nPatPaid;
				nGroup2PatPaid += nPatPaid;

				fee.szID.Empty();
				fee.szGroupID = _T("Item");
				fee.szName = rs.GetValue(_T("name"));
				fee.szUnit = rs.GetValue(_T("unit"));
				fee.nCost = nCost;
				fee.nInsPaid = nInsCost;
				fee.nDiscount = nDiscount;
				fee.nQuantity = str2float(rs.GetValue(_T("qty")));
				fee.nPrice = str2double(rs.GetValue(_T("unitprice")));
				fee.nInsPrice = str2double(rs.GetValue(_T("insprice")));
				fee.nDiffPaid = nDiffPaid;
				fee.nPatPaid = nPatPaid;
				feeList.Add(fee);

				if(szNewGroup.Left(2) == _T("B4") || szNewGroup.Left(2) == _T("B5") ){
/*
					double nOperationCost, nOperationInsCost, nOperationDiscount, nOperationDiffCost, nOperationPatPaid;
					
					bool bHasInOperation=false;
					bool bHasOutsideInsurance = false;

					nOperationCost = nCost;
					nOperationInsCost = nInsCost;
					nOperationDiscount = nDiscount;
					nOperationDiffCost = nDiffPaid;
					nOperationPatPaid = nPatPaid;

					CString szItemID;
					rs.GetValue(_T("itemid"), szItemID);
					CRecord rsi(&pMF->m_db);
					
					szSQL.Format(_T("SELECT hfe_itemid, hfe_desc, hfe_quantity, hfe_unitprice, hfe_cost, hfe_discount, hfe_diffcost, hfe_patpaid ") \
								_T("FROM hms_fee ") \
								_T("WHERE hfe_docno=%ld and hfe_type='V' and hfe_unit='%s' ORDER BY hfe_itemid DESC "), 
								nDocumentNo, szItemID);
					rsi.ExecSQL(szSQL);

					while(!rsi.IsEOF()){
						rsi.GetValue(_T("hfe_itemid"), tmpStr);
						if(tmpStr == _T("VT000001")){
							bHasInOperation = true;
						}
						if(tmpStr == _T("VT000002")){
							bHasOutsideInsurance = true;
						}
						rsi.MoveNext();
					}



					if(m_szObjectType == _T("I") || m_szObjectType == _T("C"))
					{
						szSQL.Format(_T(" SELECT hfe_group, ") \
								_T(" 		hfe_desc as name,") \
								_T(" hfe_unit as unit, ") \
								_T(" hfe_category, ") \
								_T(" 		sum(hfe_quantity) as qty,") \
								_T(" 		hfe_insprice as unitprice,") \
								_T(" 		sum(hfe_inspaid) as cost,") \
								_T(" 		case when hfe_category<>'Y' then sum(hfe_discount) else 0 end as discount,") \
								_T(" 		sum(hfe_DiffPaid) as DiffPaid,") \
								_T(" 		sum(hfe_cost-hfe_discount-hfe_diffcost) as patpaid") \
								_T(" FROM hms_fee_view ") \
								_T(" LEFT JOIN hms_fee_group ON(hfg_id=hfe_group) ") \
								_T(" WHERE hfe_docno=%ld  and hfe_cost > 0  and hfe_class='I'  ") \
								_T(" and hfe_subgroup='%s'   ") \
								_T(" GROUP BY hfg_type_id, hfe_group, hfe_desc, hfe_unit, hfe_insprice, hfe_category, hfe_status ") \
								_T(" ORDER BY hfe_category DESC, hfg_type_id, hfe_group,  name, unitprice"), 
								nDocumentNo, szItemID);
					}
					else
					{
						szSQL.Format(_T(" SELECT hfe_group, ") \
								_T(" 		hfe_desc as name,") \
								_T(" hfe_unit as unit, ") \
								_T(" hfe_category, ") \
								_T(" 		sum(hfe_quantity) as qty,") \
								_T(" 		hfe_unitprice as unitprice,") \
								_T(" 		sum(hfe_cost) as cost,") \
								_T(" 		case when hfe_category<>'Y' then sum(hfe_discount) else 0 end as discount,") \
								_T(" 		sum(hfe_DiffPaid) as DiffPaid,") \
								_T(" 		sum(hfe_cost-hfe_discount-hfe_diffcost) as patpaid") \
								_T(" FROM hms_fee_view ") \
								_T(" LEFT JOIN hms_fee_group ON(hfg_id=hfe_group) ") \
								_T(" WHERE hfe_docno=%ld  and hfe_cost > 0  and hfe_class='I'  ") \
								_T(" and hfe_subgroup='%s'   ") \
								_T(" GROUP BY hfg_type_id, hfe_group, hfe_desc, hfe_unit, hfe_unitprice, hfe_category, hfe_status") \
								_T(" ORDER BY hfe_category DESC, hfg_type_id, hfe_group,  name, unitprice"), 
								nDocumentNo, szItemID);

					}
					

							rsl.ExecSQL(szSQL);
							
							CString szCategory;
							nCost = nInsCost = nPatPaid = nPatPaid = 0;
							nTtlCost = nTtlDiscount = nTtlDiffPaid = nTtlPatPaid = 0;
							double nTtlCost2 = 0, nTtlInsCost2=0, nTtlDiscount2 = 0, nTtlDiffPaid2 = 0, nTtlPatPaid2 = 0;
							bool bInOperation=false, bOutOperation=false;
							while(!rsl.IsEOF()){
								rsl.GetValue(_T("cost"),nCost);
								rsl.GetValue(_T("inscost"),nInsCost);
								rsl.GetValue(_T("discount"), nDiscount);
								rsl.GetValue(_T("DiffPaid"), nDiffPaid);
								rsl.GetValue(_T("patpaid"), nPatPaid);
								if(nDiscount == 0)
								{
								//	nCost = 0;
									//nPatPaid = 0;
									//nDiffPaid = 0;
								}


								
								rsl.GetValue(_T("hfe_category"), szCategory);
								if(szCategory == _T("Y") && !bInOperation)
								{
									bInOperation = true;
									CString szLabel;
									TranslateString(_T("Inside operation"), szLabel);
									fee.szGroupID = _T("SubItemGroup");
									fee.szID = _T("*");
									fee.szName = szLabel;
									fee.szUnit.Empty();
									fee.nCost = 0;
									fee.nInsPaid = 0;
									fee.nDiscount = 0;
									fee.nDiffPaid = 0;
									fee.nPatDebt = 0;
									feeList.Add(fee);


								}
								if(szCategory != _T("Y") && !bOutOperation)
								{
									CString szLabel;

									bOutOperation = true;

									if(nTtlCost > 0){
										TranslateString(_T("Amount"), tmpStr);
										fee.szGroupID = _T("SubItemAmount");
										fee.szID.Empty();
										fee.szName.Format(_T("   %s"), tmpStr);;
										fee.nCost = nTtlCost;
										fee.nDiscount = nTtlDiscount;
										fee.nDiffPaid = nTtlDiffPaid;
										fee.nPatPaid = nTtlPatPaid;
										feeList.Add(fee);

										nTtlCost = nTtlDiscount = nTtlDiffPaid = nTtlPatPaid = 0;

									}


									

									TranslateString(_T("Outside operation"), szLabel);

									fee.szGroupID = _T("SubItemGroup");
									fee.szID = _T("*");
									fee.szName = szLabel;
									fee.nCost = 0;
									fee.nInsPaid = 0;
									fee.nDiscount = 0;
									fee.nDiffPaid = 0;
									fee.nPatDebt = 0;

									feeList.Add(fee);


								}


								nTtlCost += nCost;
								nTtlInsCost += nInsCost;
								nTtlDiscount += nDiscount;
								nTtlDiffPaid += nDiffPaid;
								nTtlPatPaid += nPatPaid;




								if(nDiscount +nDiffPaid+nPatPaid > 0)
								{
									nTtlCost2 += nCost;
									nTtlInsCost2 += nInsCost;
									nTtlDiscount2 += nDiscount;
									nTtlDiffPaid2 += nDiffPaid;
									nTtlPatPaid2 += nPatPaid;
								}

								
								if(szCategory != _T("Y") && !bHasOutsideInsurance){
									nTotalCost += nCost;
									
									nGroup1Cost += nCost;
									nGroup2Cost += nCost;
									nTotalInsCost += nInsCost;
									nGroup1InsCost += nInsCost;
									nGroup2InsCost += nInsCost;
									nTotalDiscount += nDiscount;
									nGroup1Discount += nDiscount;
									nGroup2Discount += nDiscount;
									nTotalDiffPaid += nDiffPaid;
									nGroup1DiffPaid += nDiffPaid;
									nGroup2DiffPaid += nDiffPaid;
									nTotalPatPaid += nPatPaid;
									nGroup1PatPaid += nPatPaid;
									nGroup2PatPaid += nPatPaid;

								}

								rsl.GetValue(_T("name"), tmpStr);
								fee.szID.Empty();
								fee.szName.Format(_T("      %s"), tmpStr);
								fee.szUnit = rsl.GetValue(_T("unit"));
								fee.szGroupID = _T("SubItem");
								fee.nQuantity = str2double(rsl.GetValue(_T("qty")));
								fee.nPrice = str2double(rsl.GetValue(_T("unitprice")));
								fee.nCost = str2double(rsl.GetValue(_T("cost")));
								fee.nInsPaid = str2double(rsl.GetValue(_T("inscost")));
								fee.nDiscount = str2double(rsl.GetValue(_T("discount")));
								fee.nDiffPaid = str2double(rsl.GetValue(_T("diffpaid")));
								fee.nPatPaid = str2double(rsl.GetValue(_T("patpaid")));
								feeList.Add(fee);
								rsl.MoveNext();
							}
							
							if(nTtlCost > 0){
								TranslateString(_T("Amount"), tmpStr);
								fee.szGroupID = _T("SubItemAmount");
								fee.szID.Empty();
								fee.szName.Format(_T("   %s"), tmpStr);;
								fee.szUnit.Empty();
								
								fee.nCost = nTtlCost;
								fee.nDiscount = nTtlDiscount;
								fee.nDiffPaid = nTtlDiffPaid;
								fee.nPatPaid = nTtlPatPaid;
								feeList.Add(fee);

								nTtlCost = nTtlDiscount = nTtlDiffPaid = nTtlPatPaid = 0;

							}

							bool bOutsideOperation = false;

							
							nCost = nInsCost = nPatPaid  = 0;
							rsi.MoveFirst();
							while(!rsi.IsEOF())
							{
								CString szItemID;
								rsi.GetValue(_T("hfe_itemid"), szItemID);
								
								fee.szID = _T("*");
								fee.szGroupID = _T("Item");
								fee.szName = rsi.GetValue(_T("hfe_desc"));
								fee.nQuantity = str2double(rsi.GetValue(_T("hfe_quantity")));
								fee.nCost = 0;
								fee.nDiscount = 0;
								fee.nPrice = str2double(rsi.GetValue(_T("hfe_unitprice")));

								if(m_szObjectType == _T("S"))
								{
									rsi.GetValue(_T("hfe_patpaid"), nPatPaid);
									fee.nCost = nPatPaid;
									fee.nDiffPaid = 0;
									fee.nPatPaid = nPatPaid;

									nCost = nPatPaid;
									nDiscount =0;
									nInsCost =0;
									feeList.Add(fee);
									nTtlCost += nCost;
									nTtlPatPaid += nPatPaid;
									
								}
								else
								{
									rsi.GetValue(_T("hfe_cost"), nCost);
									rsi.GetValue(_T("hfe_discount"), nDiscount);
									rsi.GetValue(_T("hfe_diffcost"), nDiffPaid);
									//rsi.GetValue(_T("hfe_patpaid"), nPatPaid);
									nPatPaid =0;
									fee.nDiffPaid = nDiffPaid;
									fee.nPatPaid = nPatPaid;
									fee.nCost = nCost;
									fee.nDiscount = nDiscount;
									nInsCost = nCost;
									feeList.Add(fee);
									nTtlDiffPaid += nDiffPaid;
									nTtlCost += nCost;
									nTtlInsCost += nInsCost;
									nTtlDiscount += nDiscount;
									nTtlPatPaid += nPatPaid;
								}
								if(nDiscount +nDiffPaid+nPatPaid > 0)
								{
									nTotalCost += nCost;
									
									nGroup1Cost += nCost;
									nGroup2Cost += nCost;
									nTotalInsCost += nInsCost;
									nGroup1InsCost += nInsCost;
									nGroup2InsCost += nInsCost;
									nTotalDiscount += nDiscount;
									nGroup1Discount += nDiscount;
									nGroup2Discount += nDiscount;
									nTotalDiffPaid += nDiffPaid;
									nGroup1DiffPaid += nDiffPaid;
									nGroup2DiffPaid += nDiffPaid;
									nTotalPatPaid += nPatPaid;
									nGroup1PatPaid += nPatPaid;
									nGroup2PatPaid += nPatPaid;
								}

								bOutsideOperation = true;

								rsi.MoveNext();

							}

*/					
					}	//end B4, B5
			}
			rs.MoveNext();
		}
		if(!bFound)
		{
			feeList.RemoveAt(nIDX);
			nGroup1Cost=nGroup1InsCost=nGroup1Discount=nGroup1DiffPaid=nGroup1PatPaid=0;
			nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;
		}
		else
		{
			if(nGroup1Cost > 0){
				if(nGroup2Cost > 0 && nGroup1Cost != nGroup2Cost){
					TranslateString(_T("Sub Amount"), tmpStr);
					fee.szGroupID = _T("SubAmount");
					fee.szID.Empty();;
					fee.szName = tmpStr;
					fee.nCost = nGroup2Cost;
					fee.nInsPaid = nGroup2InsCost;
					fee.nDiscount = nGroup2Discount;
					fee.nDiffPaid = nGroup2DiffPaid;
					fee.nPatPaid = nGroup2PatPaid;
					feeList.Add(fee);
				}

					TranslateString(_T("Total"), tmpStr);
					tmpStr.AppendFormat(_T(" (%d)"), nIndex);
					fee.szGroupID = _T("GrandAmount");
					fee.szID.Empty();;
					fee.szName = tmpStr;
					fee.nCost = nGroup1Cost;
					fee.nInsPaid = nGroup1InsCost;
					fee.nDiscount = nGroup1Discount;
					fee.nDiffPaid = nGroup1DiffPaid;
					fee.nPatPaid = nGroup1PatPaid;
					feeList.Add(fee);
			}
			nGroup1Cost=nGroup1InsCost=nGroup1Discount=nGroup1DiffPaid=nGroup1PatPaid=0;
			nGroup2Cost=nGroup2InsCost=nGroup2Discount=nGroup2DiffPaid=nGroup2PatPaid=0;

			nIndex++;
		}



		grs.MoveNext();
	}

	nTotalCost  = ceil(nTotalCost);
	nTotalInsCost  = ceil(nTotalInsCost);
	nTotalDiscount  = ceil(nTotalDiscount);
	nTotalDiffPaid  = ceil(nTotalDiffPaid);
	nTotalPatPaid  = ceil(nTotalPatPaid);
	nTotalPayable = nTotalCost-nTotalDiscount;
	nTotalDiscount = nTotalDiscount;

	TranslateString(_T("Total Amount"), szName);
	fee.szGroupID = _T("TotalAmount");
	fee.szName.Format(_T("%s "), szName);
	fee.nCost = nTotalCost;
	fee.nInsPaid = nTotalInsCost;
	fee.nDiscount = nTotalDiscount;
	fee.nDiffPaid = nTotalDiffPaid;
	fee.nPatPaid = nTotalPatPaid;
	feeList.Add(fee);
	
	nSumCost += nTotalCost;
	nSumInsCost += nTotalInsCost;
	nSumDiscount += nTotalDiscount;
	nSumDiffPaid += nTotalDiffPaid;
	nSumPatPaid += nTotalPatPaid;


		for (int i =0; i < feeList.GetCount(); i++){
			fee = feeList[i];
			szType = fee.szGroupID;

			if(szType == _T("Type"))
			{
				grp = rpt.GetGroupHeader(1);
				rptDetail = rpt.AddDetail(grp);
				tmpStr = fee.szID;
				rptDetail->SetValue(_T("TypeID"), tmpStr);
				StringUpper(fee.szName, tmpStr);
				rptDetail->SetValue(_T("TypeName"), tmpStr);
			}
			else if(szType == _T("Group")){
				grp = rpt.GetGroupHeader(1);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("TypeID"), fee.szID);
				rptDetail->SetValue(_T("TypeName"), fee.szName);
			}
			else if(szType == _T("SubGroup")){
				grp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("SubGroupID"), fee.szID);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);
				
				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupInsCost"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);
				
				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);
				

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupPatPaid"), tmpStr);


				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("SubGroupInsUnPaid"), tmpStr);
				}
			}
			else if(szType == _T("SubAmount")){
				grp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(grp);
				CReportItem *obj ;
				obj = rptDetail->GetItem(_T("SubGroupName")); if(obj) obj->SetBold(true);
				obj = rptDetail->GetItem(_T("SubGroupCost")); if(obj) obj->SetBold(true);
				obj = rptDetail->GetItem(_T("SubGroupDiscount")); if(obj) obj->SetBold(true);
				obj = rptDetail->GetItem(_T("SubGroupPatpaid")); if(obj) obj->SetBold(true);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);


				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("SubGroupInsUnPaid"), tmpStr);
				}

			}
			else if(szType == _T("GrandAmount")){
				grp = rpt.GetGroupHeader(2);
				rptDetail = rpt.AddDetail(grp);
			//	rptDetail->GetItem(_T("SubGroupName"))->SetBold(true);
			//	rptDetail->GetItem(_T("SubGroupCost"))->SetBold(true);
			//	rptDetail->GetItem(_T("SubGroupDiscount"))->SetBold(true);
			//	rptDetail->GetItem(_T("SubGroupPatpaid"))->SetBold(true);
				rptDetail->SetValue(_T("SubGroupName"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("SubGroupCost"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiscount"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupDiffPaid"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("SubGroupPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("SubGroupInsUnPaid"), tmpStr);
				}

			}

			else if(szType == _T("Item")){
				rptDetail = rpt.AddDetail();
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), fee.szUnit);
				FormatCurrency(fee.nQuantity, tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("11"), tmpStr);
				}
				
			}
			else if(szType == _T("SubItemGroup")){
				rptDetail = rpt.AddDetail();
				CReportItem *obj = NULL;
				for (int i = 1; i <= 9; i++)
				{
					obj = rptDetail->GetItem(i);
					if(obj) 
					{
						obj->SetItalic(true);
						obj->SetBold(true);
					}

				}
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), _T(""));
				rptDetail->SetValue(_T("4"), _T(""));
				rptDetail->SetValue(_T("5"), _T(""));
				
			}

			else if(szType == _T("SubItem")){
				rptDetail = rpt.AddDetail();
				CReportItem *obj = NULL;
				for (int i = 1; i <= 9; i++)
				{
					obj = rptDetail->GetItem(i);
					if(obj) obj->SetItalic(true);
				}

				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), fee.szUnit);
				FormatCurrency(fee.nQuantity, tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("11"), tmpStr);
				}
				
			}
			else if(szType == _T("SubItemAmount")){
				rptDetail = rpt.AddDetail();
				CReportItem *obj = NULL;
				for (int i = 1; i <= 9; i++)
				{
					obj = rptDetail->GetItem(i);
					if(obj) 
					{
						obj->SetItalic(true);
						obj->SetBold(true);
					}

				}
				rptDetail->SetValue(_T("1"), fee.szID);
				rptDetail->SetValue(_T("2"), fee.szName);
				rptDetail->SetValue(_T("3"), _T(""));
				rptDetail->SetValue(_T("4"), _T(""));
				FormatCurrency(fee.nPrice, tmpStr);
				rptDetail->SetValue(_T("5"), _T(""));
				
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("10"), tmpStr);
				
				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("7"), tmpStr);

				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("8"), tmpStr);

				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("9"), tmpStr);
				

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("11"), tmpStr);
				}
				
			}

			else if(szType == _T("Dousage")){
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(4));
				rptDetail->SetValue(_T("usage"), fee.szName);
			}
			else if(szType == _T("TotalAmount")){
				grp = rpt.GetGroupHeader(3);
				rptDetail = rpt.AddDetail(grp);
				rptDetail->SetValue(_T("TotalAmountLabel"), fee.szName);
				FormatCurrency(fee.nCost, tmpStr);
				rptDetail->SetValue(_T("TotalAmount"), tmpStr);

				FormatCurrency(fee.nInsPaid, tmpStr);
				rptDetail->SetValue(_T("TotalInsCost"), tmpStr);

				FormatCurrency(fee.nDiscount, tmpStr);
				rptDetail->SetValue(_T("TotalDiscount"), tmpStr);
				
				FormatCurrency(fee.nDiffPaid, tmpStr);
				rptDetail->SetValue(_T("TotalDiffPaid"), tmpStr);
				//fee.nPatPaid= fee.nCost-fee.nDiscount;
				FormatCurrency(fee.nPatPaid, tmpStr);
				rptDetail->SetValue(_T("TotalPatPaid"), tmpStr);

				if(fee.nDiscount > 0)
				{
					FormatCurrency(fee.nPatPaid-fee.nDiffPaid, tmpStr);
					rptDetail->SetValue(_T("TotalInsUnPaid"), tmpStr);
				}
			}
		}	
	

	if(nTotalCost != nSumCost){
		CString szName;
		TranslateString(_T("Total Amount"), tmpStr);
		szName.Format(_T("%s"), tmpStr);
			grp = rpt.GetGroupHeader(3);
			rptDetail = rpt.AddDetail(grp);
			rptDetail->SetValue(_T("TotalAmountLabel"), szName);
			FormatCurrency(nSumCost, tmpStr);
			rptDetail->SetValue(_T("TotalAmount"), tmpStr);

			FormatCurrency(nSumInsCost, tmpStr);
			rptDetail->SetValue(_T("TotalInsCost"), tmpStr);

			FormatCurrency(nSumDiscount, tmpStr);
			rptDetail->SetValue(_T("TotalDiscount"), tmpStr);
				
			FormatCurrency(nSumDiffPaid, tmpStr);
			rptDetail->SetValue(_T("TotalDiffPaid"), tmpStr);
			FormatCurrency(nSumPatPaid, tmpStr);
			rptDetail->SetValue(_T("TotalPatPaid"), tmpStr);
	}

	szSQL.Format(_T(" SELECT hfe_amount as hfe_amount,") \
	_T("   hfe_inspaid,") \
	_T("   hfe_discount,") \
	_T("   hfe_patpaid,") \
	_T("   hfe_payment,") \
	_T("   hfe_diffcost,") \
	_T("   hfe_diffpaid,") \
	_T("   hfe_deposit,") \
	_T("   hfe_refund,") \
	_T("   hfe_freeamount") \
	_T(" FROM hms_fee_invoice") \
	_T(" WHERE hfe_invoiceno=%ld"), nInvoiceNo);
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF())
	{
		double nTotalDiffPaid;
		double nTotalPayment;
		double nTotalPatPaid;

		rs.GetValue(_T("hfe_amount"), nTotalCost);
		rs.GetValue(_T("hfe_inspaid"), nTotalInsCost);
		rs.GetValue(_T("hfe_discount"), nTotalDiscount);
		rs.GetValue(_T("hfe_patpaid"), nTotalPatPaid);
		rs.GetValue(_T("hfe_diffcost"), nTotalDiffPaid);
		rs.GetValue(_T("hfe_deposit"), nTotalDeposit);
		rs.GetValue(_T("hfe_refund"), nTotalRefund);
		rs.GetValue(_T("hfe_freeamount"), nTotalFree);
		rs.GetValue(_T("hfe_payment"), nTotalPayment);

		


		rpt.GetReportFooter()->SetValue(_T("PatientSign"), szPatientName);
		FormatCurrency((nTotalCost), tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumAmount"), tmpStr);
		CString szSumInWord;
		tmpStr.Format(_T("%.0f"), nTotalCost);
		MoneyToString(tmpStr, szSumInWord);
		szSumInWord += _T(" \x111\x1ED3ng.");
		rpt.GetReportFooter()->SetValue(_T("SumInWord"), szSumInWord);
		
		FormatCurrency(nTotalInsCost, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumInsCost"), tmpStr);


		FormatCurrency(nTotalDiscount, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumInsDiscount"), tmpStr);

		FormatCurrency(nTotalPatPaid, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumInsUnPaid"), tmpStr);

		FormatCurrency((nTotalDiffPaid), tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumDiffPaid"), tmpStr);


		FormatCurrency(nTotalPatPaid, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumPatPaid"), tmpStr);



		FormatCurrency((nTotalDeposit), tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumDeposit"), tmpStr);
		
		FormatCurrency(nTotalFree, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumDiscount"), tmpStr);

		FormatCurrency(nSumFoodAmount, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumFoodAmount"), tmpStr);

		if(nTotalRefund > 0)
		{
		//	FormatCurrency(nTotalRefund, tmpStr);
		//	rpt.GetReportFooter()->SetValue(_T("SumPayment"), tmpStr);

			FormatCurrency(nTotalRefund, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("SumRefund"), tmpStr);
		}
		else
		{
			FormatCurrency(nTotalPayment, tmpStr);
			rpt.GetReportFooter()->SetValue(_T("SumPayment"), tmpStr);

		}
		
	}

	int nInsRate;
	rs.GetValue(_T("disrate"), nInsRate);
	if(nInsRate > 0)
	{
		int nUnRate = 100 - nInsRate;
		if(nUnRate ==0)
			nUnRate = nInsRate;

		tmpStr.Format(_T("(%d%s)"), nUnRate, _T("%"));
		rpt.GetReportFooter()->SetValue(_T("UnRate"), tmpStr);
	}

//	szSQL.Format(_T("UPDATE hms_fee_invoice SET hfe_print=hfe_print+1 WHERE hfe_invoiceno=%ld "), nInvoiceNo);
//	ExecSQL(szSQL);

//	rpt.Print();
	rpt.PrintPreview();

}

void CPrintForms::E10_PrintSaleExport(long nOID){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CString szSQL, tmpStr, szMoney, szSysDate;
	CReport rpt;
	CReportSection *rptHeader = NULL, *rptDetail = NULL, *rptFooter = NULL;
	int nIdx = 1, nQty = 0, nTmp = 0;
	double nAmt = 0;
	long double nTtlAmt = 0;
	CRecord rs(&pMF->m_db);
	if (!rpt.Init(_T("Reports\\HMS\\PMS_DRUGEXPORT.RPT")))
		return;
	szSQL.Format(_T(" SELECT so_documentno                  documentno, ") \
				_T("        so_orderno                     orderno, ") \
				_T("        N''                            department, ") \
				_T("        so_orderdate                   orderdate, ") \
				_T("        Get_storagename(so_storage_id) storagename, ") \
				_T("        so_description descr") \
				_T(" FROM   sale_order ") \
				_T(" WHERE so_sale_order_id = %ld"), nOID);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		rptHeader->SetValue(_T("ReferenceNo"), rs.GetValue(_T("documentno")));
		rptHeader->SetValue(_T("InvoiceNo"), rs.GetValue(_T("orderno")));
		rptHeader->SetValue(_T("orderdate"), CDate::Convert(rs.GetValue(_T("orderdate")), yyyymmdd, ddmmyyyy));
		rptHeader->SetValue(_T("stockname"), rs.GetValue(_T("storagename")));
		rptHeader->SetValue(_T("note"), rs.GetValue(_T("descr")));
	}
	szSQL.Format(_T(" SELECT    product_name product_name, ") \
				_T("           product_uomname uom_name, ") \
				_T("           product_lotno lot_no, ") \
				_T("           case when product_hasexp = 'Y' then product_expdate else null end exp_date, ") \
				_T("           SUM(sol_qtyorder) qty, ") \
				_T("           sol_unitprice price, ") \
				_T("           SUM(sol_qtyorder) * sol_unitprice amount ") \
				_T(" FROM      sale_orderline ") \
				_T(" LEFT JOIN m_productitem_view ON ( sol_product_item_id = product_item_id ) ") \
				_T(" WHERE     sol_sale_order_id = %ld ") \
				_T(" GROUP     BY product_name, ") \
				_T("              product_uomname, ") \
				_T("              product_lotno, ") \
				_T("              product_expdate, ") \
				_T("			  product_hasexp,") \
				_T("              sol_unitprice "), nOID);
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("product_name")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("uom_name")));
		rptDetail->SetValue(_T("4"), CDate::Convert(rs.GetValue(_T("exp_date")), yyyymmdd, ddmmyyyy));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("lot_no")));
		rs.GetValue(_T("qty"), nTmp);
		nQty += nTmp;
		rptDetail->SetValue(_T("6"), int2str(nTmp));
		rptDetail->SetValue(_T("8"), rs.GetValue(_T("price")));
		rs.GetValue(_T("amount"), nAmt);
		nTtlAmt += nAmt;
		rptDetail->SetValue(_T("9"), double2str(nAmt));
		rs.MoveNext();
	}
	nIdx--;
	rptFooter = rpt.GetReportFooter();
	if (rptFooter)
	{
		rptFooter->SetValue(_T("TotalItem"), int2str(nIdx));
		tmpStr.Format(_T("%f"), nTtlAmt);
		MoneyToString(tmpStr, szMoney);
		szMoney += _T(" \x111\x1ED3ng.");
		rptFooter->SetValue(_T("TotalAmount"), tmpStr);
		rptFooter->SetValue(_T("SumInWord"), szMoney);
		szSysDate = pMF->GetSysDate();
		tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), pMF->GetSysTime().Left(5), szSysDate.Mid(8, 2), szSysDate.Mid(5, 2), szSysDate.Left(4));
		rptFooter->SetValue(_T("PrintDate"), tmpStr);
	}
	rpt.PrintPreview();
}

void CPrintForms::E10_PrintSaleImport(long nOID){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CReportSection *rptDetail = NULL, *rptHeader = NULL, *rptFooter = NULL;
	int nIdx = 1;
	double nCost = 0;
	long double nTtlCost = 0;
	if (!rpt.Init(_T("Reports\\HMS\\MA_PHIEUNHAPKHO.RPT")))
		return;
	CString szSQL, tmpStr, tmpStr2;
	szSQL.Format(_T("SELECT hpo_pharmareturnorder_id order_no,") \
				_T("		hpo_createddate order_date, get_storagename(hpo_storage_id) storage_name,") \
				_T("		hpo_documentno document_no") \
				_T(" FROM	hms_pharmareturnorder") \
				_T(" WHERE	hpo_status = 'A'") \
				_T(" AND	hpo_pharmareturnorder_id = %ld"), nOID);
_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		rptHeader->SetValue(_T("ReferenceNo"), rs.GetValue(_T("document_no")));
		rptHeader->SetValue(_T("OrderNo"), rs.GetValue(_T("order_no")));
		rs.GetValue(_T("order_date"), tmpStr2);
		tmpStr = CDate::Convert(tmpStr2, yyyymmdd, ddmmyyyy);
		rptHeader->SetValue(_T("OrderDate"), tmpStr);
		rptHeader->SetValue(_T("StockName"), rs.GetValue(_T("storage_name")));
		rptHeader->SetValue(_T("Reason"), _T(""));
	}
	szSQL.Format(_T("SELECT product_name,") \
				_T("		product_uomname,") \
				_T("		case when product_hasexp = 'Y' then product_expdate else null end product_expdate,") \
				_T("		product_lotno,") \
				_T("		sum(hpol_qtyreturn) hpol_qtyreverse,") \
				_T("		hpol_unitprice,")
				_T("		sum(hpol_qtyreturn*hpol_unitprice) amount") \
				_T(" FROM hms_pharmareturnorder_line") \
				_T(" LEFT JOIN m_productitem_view ON (hpol_product_item_id = product_item_id)") \
				_T(" WHERE hpol_pharmareturnorder_id = %ld") \
				_T(" GROUP BY product_name, product_uomname, product_hasexp, product_expdate, product_lotno, hpol_unitprice"), nOID);
_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("product_name")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("product_uomname")));
		rs.GetValue(_T("product_expdate"), tmpStr2);
		tmpStr = CDate::Convert(tmpStr2, yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("4"), tmpStr);
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("product_lotno")));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("hpol_qtyreverse")));
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("hpol_unitprice")));
		rptDetail->SetValue(_T("8"), rs.GetValue(_T("")));
		rs.GetValue(_T("amount"), nCost);
		nTtlCost += nCost;
		rptDetail->SetValue(_T("9"), double2str(nCost));
		rs.MoveNext();
	}
	rptDetail = NULL;
	tmpStr2.Format(_T("%f"), nTtlCost);
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	if (rptDetail)
		rptDetail->SetValue(_T("s9"), tmpStr2);
	rptFooter = rpt.GetReportFooter();
	nIdx--;
	rptFooter->SetValue(_T("TotalItem"), int2str(nIdx));	
	MoneyToString(tmpStr2, tmpStr);
	tmpStr += _T(" \x111\x1ED3ng");
	rptFooter->SetValue(_T("SumInWord"), tmpStr);
	tmpStr2 = pMF->GetSysDate();
	tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), tmpStr2.Mid(8, 2), tmpStr2.Mid(5, 2), tmpStr2.Left(4));
	rptFooter->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}

void CPrintForms::E10_PrintReturnImport(long nOID){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CReportSection *rptDetail = NULL, *rptHeader = NULL, *rptFooter = NULL;
	int nIdx = 1;
	double nCost = 0;
	long double nTtlCost = 0;
	if (!rpt.Init(_T("Reports\\HMS\\MA_PHIEUNHAPKHO.RPT")))
		return;
	CString szSQL, tmpStr, tmpStr2;	
	szSQL.Format(_T("SELECT mt_orderno order_no,") \
				_T("		mt_orderdate order_date,") \
				_T("		get_storagename(mt_storage_to_id) storage_name,") \
				_T("		mt_description descr,") \
				_T("		mt_documentno document_no,") \
				_T("		mt_deliveryby deliver_by,") \
				_T("		mt_receiveby receive_by") \
				_T(" FROM m_transaction") \
				_T(" WHERE mt_transaction_id = %ld"), nOID);
_fmsg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		rptHeader->SetValue(_T("ReferenceNo"), rs.GetValue(_T("document_no")));
		rptHeader->SetValue(_T("OrderNo"), rs.GetValue(_T("order_no")));
		rs.GetValue(_T("order_date"), tmpStr2);
		tmpStr = CDate::Convert(tmpStr2, yyyymmdd, ddmmyyyy);
		rptHeader->SetValue(_T("OrderDate"), tmpStr);
		rptHeader->SetValue(_T("StockName"), rs.GetValue(_T("storage_name")));
		rptHeader->SetValue(_T("Reason"), rs.GetValue(_T("descr")));
		rptHeader->SetValue(_T("DeliverBy"), rs.GetValue(_T("deliver_by")));
		rptHeader->SetValue(_T("ReceiveBy"), rs.GetValue(_T("receive_by")));
	}
	szSQL.Format(_T("SELECT product_name,") \
				_T("		product_uomname,") \
				_T("		case when product_hasexp = 'Y' then product_expdate else null end product_expdate,") \
				_T("		product_lotno,") \
				_T("		sum(mtl_qtyorder) mtl_qtyreverse,") \
				_T("		DECODE(0, mtl_saleprice, mtl_taxprice, mtl_saleprice) mtl_saleprice,")
				_T("		sum(mtl_qtyorder*DECODE(0, mtl_saleprice, mtl_taxprice, mtl_saleprice)) amount") \
				_T(" FROM m_transactionline") \
				_T(" LEFT JOIN m_productitem_view ON (mtl_product_item_id = product_item_id)") \
				_T(" WHERE  mtl_transaction_id = %ld") \
				_T(" GROUP BY product_name, product_uomname, product_hasexp, product_expdate, product_lotno, mtl_saleprice, mtl_taxprice"), nOID);
_msg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("product_name")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("product_uomname")));
		rs.GetValue(_T("product_expdate"), tmpStr2);
		tmpStr = CDate::Convert(tmpStr2, yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("4"), tmpStr);
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("product_lotno")));
		rptDetail->SetValue(_T("6"), rs.GetValue(_T("mtl_qtyreverse")));
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("mtl_saleprice")));
		rptDetail->SetValue(_T("8"), rs.GetValue(_T("")));
		rs.GetValue(_T("amount"), nCost);
		nTtlCost += nCost;
		rptDetail->SetValue(_T("9"), double2str(nCost));
		rs.MoveNext();
	}
	rptDetail = NULL;
	tmpStr2.Format(_T("%f"), nTtlCost);
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	if (rptDetail)
		rptDetail->SetValue(_T("s9"), tmpStr2);
	rptFooter = rpt.GetReportFooter();
	nIdx--;
	rptFooter->SetValue(_T("TotalItem"), int2str(nIdx));	
	MoneyToString(tmpStr2, tmpStr);
	tmpStr += _T(" \x111\x1ED3ng");
	rptFooter->SetValue(_T("SumInWord"), tmpStr);
	tmpStr2 = pMF->GetSysDate();
	tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), tmpStr2.Mid(8, 2), tmpStr2.Mid(5, 2), tmpStr2.Left(4));
	rptFooter->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}

void CPrintForms::TM_PrintDailyValuableMaterial(long nDocNo, long nOpIdx){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CReport rpt;
	CReportSection * rptDetail = NULL, *rptHeader = NULL;
	CString szSQL, szOldType, szNewType, tmpStr, szSysDate;
	int nIdx = 1, nGroup = 0;
	double nGrpAmt = 0, nAmt = 0;
	CRecord rs(&pMF->m_db);
	CStringArray arrGroup;
	arrGroup.Add(_T("I"));
	arrGroup.Add(_T("II"));
	arrGroup.Add(_T("III"));
	arrGroup.Add(_T("IV"));
	arrGroup.Add(_T("V"));
	arrGroup.Add(_T("VI"));
	if (!rpt.Init(_T("Reports\\HMS\\TM_SURGERYPRESCRIPTIONORDER.RPT")))
		return;
	szSQL.Format(_T(" SELECT   hd_docno doc_no, ") \
				_T("           hcr_recordno record_no, ") \
				_T("           get_departmentname(ho_deptid) dept_name, ") \
				_T("           Get_patientname(hd_docno) patient_name, ") \
				_T("           Extract(year FROM hp_birthdate) yob, ") \
				_T("           get_syssel_desc('sys_sex', hp_sex) sex, ") \
				_T("           Hms_getaddress(hp_provid, hp_distid, hp_villid) addr, ") \
				_T("           get_objectname(hd_object) obj_name, ") \
				_T("           get_syssel_desc('hms_rank', hd_rank) rank_name, ") \
				_T("           hd_cardno card_no, ") \
				_T("           hd_diagnostic diagnostic, ") \
				_T("           get_departmentname(ho_pdeptid) pdept_name , ") \
				_T("		   ho_opera_table op_table,") \
				_T("           ho_comment op_name") \
				_T(" FROM      hms_op_materialorder ") \
				_T(" LEFT JOIN hms_doc ON ( hd_docno = hopm_docno ) ") \
				_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno ) ") \
				_T(" LEFT JOIN hms_patient ON ( hp_patientno = hd_patientno ) ") \
				_T(" LEFT JOIN hms_operation ON ( ho_idx = hopm_operation_id ) ") \
				_T(" WHERE     hd_docno = %ld"), nDocNo);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		rptHeader->SetValue(_T("Department"), rs.GetValue(_T("dept_name")));
		rptHeader->SetValue(_T("DocumentNo"), rs.GetValue(_T("doc_no")));
		rptHeader->SetValue(_T("RecordNo"), rs.GetValue(_T("record_no")));
		rptHeader->SetValue(_T("PatientName"), rs.GetValue(_T("patient_name")));
		rptHeader->SetValue(_T("age"), rs.GetValue(_T("yob")));
		rptHeader->SetValue(_T("sex"), rs.GetValue(_T("sex")));
		rptHeader->SetValue(_T("address"), rs.GetValue(_T("addr")));
		rptHeader->SetValue(_T("object"), rs.GetValue(_T("obj_name")));
		rptHeader->SetValue(_T("rank"), rs.GetValue(_T("rank_name")));
		rptHeader->SetValue(_T("cardno"), rs.GetValue(_T("card_no")));
		rptHeader->SetValue(_T("diagnostic"), rs.GetValue(_T("diagnostic")));
		rptHeader->SetValue(_T("pdepartment"), rs.GetValue(_T("pdept_name")));
		rptHeader->SetValue(_T("surgeryname"), rs.GetValue(_T("op_name")));
		rptHeader->SetValue(_T("surgerytable"), rs.GetValue(_T("op_table")));
	}
	szSQL.Format(_T(" SELECT    product_producttype, ") \
				_T("		   get_producttypename(product_producttype) typename,") \
				_T("           product_name, ") \
				_T("           product_manufacturename, ") \
				_T("           product_uomname, ") \
				_T("           hopm_unitprice, ") \
				_T("           hopm_qtyissue, ") \
				_T("           hopm_amount ") \
				_T(" FROM      hms_op_materialorder ") \
				_T(" LEFT JOIN m_product_view ON ( product_id = hopm_product_id ) ") \
				_T(" WHERE     hopm_docno = %ld ") \
				_T("       AND hopm_operation_id = %ld"), nDocNo, nOpIdx);
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("product_producttype"), szNewType);
		if (szNewType != szOldType)
		{
			if (nGrpAmt > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				tmpStr.Format(_T("%s (%s)"), rptDetail->GetValue(_T("TotalAmount")), arrGroup.GetAt(nGroup));
				rptDetail->Format(_T("TotalAmount"), tmpStr);
				rptDetail->SetValue(_T("s6"), double2str(nGrpAmt));
				nGrpAmt = 0;
				nGroup++;
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			tmpStr.Format(_T("%s. %s"), arrGroup.GetAt(nGroup), rs.GetValue(_T("typename")));
			rptDetail->SetValue(_T("GroupName"), tmpStr);
			szOldType = szNewType;			
		}
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("product_name")));
		rptDetail->SetValue(_T("7"), rs.GetValue(_T("product_manufacturename")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("product_uomname")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("hopm_unitprice")));
		rptDetail->SetValue(_T("5"), rs.GetValue(_T("hopm_qtyissue")));
		rs.GetValue(_T("hopm_amount"), nAmt);
		nGrpAmt += nAmt;
		rptDetail->SetValue(_T("6"), double2str(nAmt));
		rs.MoveNext();
	}
	if (nGrpAmt > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
		tmpStr.Format(_T("%s (%s)"), rptDetail->GetValue(_T("TotalAmount")), arrGroup.GetAt(nGroup));
		rptDetail->SetValue(_T("TotalAmount"), tmpStr);
		rptDetail->SetValue(_T("s6"), double2str(nGrpAmt));
	}
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Mid(8, 2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
	
}

void CPrintForms::TM_DepositRequestForm(long nDocNo, CString szDeptID, long nAmount){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szSQL, szDate = pMF->GetSysDate();
	CRecord rs(&pMF->m_db);
	tmpStr = _T("Reports\\HMS\\TM_GIAYYEUCAUTAMUNG.RPT");
	if (!rpt.Init(tmpStr))
		return;
	szSQL.Format(_T(" SELECT    Get_patientname(hd_docno)       patient_name, ") \
				_T("           hms_getagebydoc(hd_docno) age, ") \
				_T("		   get_syssel_desc('sys_sex', hp_sex) sex,") \
				_T("           Get_objectname(hd_object)       object_name, ") \
				_T("		   hd_cardno card_no,") \
				_T("           Hms_getaddress(hp_provid, hp_distid, hp_villid) ") \
				_T("           ||' ' ") \
				_T("           || hp_dtladdr                   addr, ") \
				_T("		   hms_getroomname(hb_deptid, hb_roomid) room_name") \
				_T(" FROM      hms_doc ") \
				_T(" LEFT JOIN hms_patient ON ( hd_patientno = hd_patientno ) ") \
				_T(" LEFT JOIN hms_bed ON (hb_docno = hd_docno AND hb_status = 'O')") \
				_T(" WHERE hd_docno = %ld"), nDocNo);
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{
		rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
		rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
		rpt.GetReportHeader()->SetValue(_T("DocumentNo"), long2str(nDocNo));
		rpt.GetReportHeader()->SetValue(_T("Department"), pMF->GetDepartmentName(szDeptID));
		rpt.GetReportHeader()->SetValue(_T("PatientName"), rs.GetValue(_T("patient_name")));
		rpt.GetReportHeader()->SetValue(_T("Age"), rs.GetValue(_T("age")));
		rpt.GetReportHeader()->SetValue(_T("Sex"), rs.GetValue(_T("sex")));
		rpt.GetReportHeader()->SetValue(_T("Address"), rs.GetValue(_T("addr")));
		rpt.GetReportHeader()->SetValue(_T("Object"), rs.GetValue(_T("object_name")));
		rpt.GetReportHeader()->SetValue(_T("CardNo"), rs.GetValue(_T("card_no")));
		rpt.GetReportHeader()->SetValue(_T("RoomName"), rs.GetValue(_T("room_name")));
		rpt.GetReportHeader()->SetValue(_T("TotalAmount"), long2str(nAmount));
		MoneyToString(long2str(nAmount), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("SumInWord"), tmpStr);
		tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szDate.Mid(8,2), szDate.Mid(5, 2), szDate.Left(4));
		rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	}
	rpt.PrintPreview();

}

void CPrintForms::PrintTestOrderX(long nOrderID, bool bPreview, bool bDirect){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szGroupID, szSex, szStatus;
	CString szReportName, szOrderName;
	CString szSQL;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CRecord rs3(&pMF->m_db);
	long nDocumentNo;
	int nRefIndex, nRoomID;
	int nNumberOrder=0;
	CString tmpStr2, szDeptType;
	CString szFormID;

	if(pMF->GetModuleID() != _T("EM"))
		return;


	szSQL.Format(_T("SELECT * FROM hms_testorder WHERE hpc_orderid=%ld "), nOrderID);

	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;

	rs.GetValue(_T("hpc_groupid"), szGroupID);
	rs.GetValue(_T("hpc_docno"), nDocumentNo);
	rs.GetValue(_T("hpc_refidx"), nRefIndex);
	rs.GetValue(_T("hpc_roomid"), nRoomID);
	rs.GetValue(_T("hpc_status"), szStatus);
	rs.GetValue(_T("hpc_depttype"), szDeptType);

	szSQL.Format(_T("SELECT hfg_report as reportname FROM hms_fee_group WHERE hfg_id='%s' "), szGroupID);
	rs2.ExecSQL(szSQL);
	if(!rs2.IsEOF())
		rs2.GetValue(_T("reportname"), tmpStr);

	if(tmpStr.IsEmpty())
		tmpStr = _T("PACS_VOTE.RPT");
	tmpStr.Trim();
	tmpStr.Replace(_T("_"), _T("O_"));
	szReportName.Format(_T("Reports/HMS/%s"), tmpStr);
	
	szReportName.Trim();

	if(!rpt.Init(szReportName) )
	{
		return;
	}
	 
	
		
	szSQL.Format(_T(" SELECT 	hd_patientno as patientno,  ") \
		_T(" 	hd_docno as docno,") \
		_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
		_T(" 	CAST(TO_CHAR(hp_birthdate, 'YYYY') as integer) as birthyear,") \
		_T(" 	hms_getage(trunc_date(hd_admitdate), hp_birthdate) as age, 	hp_sex as sexid,") \
		_T(" 	sys_sel_getname('sys_sex', hp_sex) as sex,") \
		_T(" 	hp_occupation as occupationid,") \
		_T(" 	sys_sel_getname('sys_occupation', cast(hp_occupation as varchar(15))) as occupation,") \
		_T(" 	hp_dtladdr as detailaddress,") \
		_T(" 	hms_getaddress(hp_provid,hp_distid,hp_villid) as address,") \
		_T(" 	hp_workplaceid as workplaceid,") \
		_T(" 	hp_workplace as workplace,") \
		_T(" 	hd_object as objectid,") \
		_T(" 	(SELECT ho_desc FROM hms_object WHERE ho_id=hd_object) as objectname,") \
		_T(" 	hd_cardno as cardno,") \
		_T(" 	hd_cardidx as cardidx, ") \
		_T(" 	hd_reldisease as relationdisease, ") \
		_T("	trim(hd_diagnostic) ||' ['||trim(hd_icd)||']' as diagnostic, ") \
		_T(" 	hc_regdate as regdate, ") \
		_T(" 	hc_expdate as expdate ") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx) ") \
		_T(" WHERE hd_docno=%ld"), nDocumentNo);
//_fmsg(_T("%s"), szSQL);
	int ret = rs2.ExecSQL(szSQL);
	if(rs2.IsEOF())
	{
		ShowMessageBox(_T("Patient not found."));
		return;
	}

	//Report Header
	tmpStr.Empty();
	rs2.GetValue(_T("sexid"), szSex);
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HealthService"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), tmpStr);
	rs.GetValue(_T("hpc_pdeptid"), tmpStr);
	if(!tmpStr.IsEmpty())
	{
		CString szPerformDept;
		StringUpper(pMF->GetDepartmentName(tmpStr), szPerformDept);
		rpt.GetReportHeader()->SetValue(_T("Faculty"), szPerformDept);
	}
	else
		rpt.GetReportHeader()->SetValue(_T("Faculty"), _T(""));

	rs.GetValue(_T("hpc_docno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128A]"),tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128B]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128C]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[39]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[93]"), tmpStr);

	tmpStr.Format(_T("%ld"), nOrderID);
	rpt.GetReportHeader()->SetValue(_T("SheetNo"), tmpStr);
	StringUpper(rs2.GetValue(_T("pname")), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("PatientName"), tmpStr);
	rs2.GetValue(_T("age"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), tmpStr);

	//Chi In ra nam sinh.
	rs2.GetValue(_T("birthyear"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);
//	_msg(_T("%s"), tmpStr);
	rs2.GetValue(_T("sex"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Sex"), tmpStr);
	tmpStr.Format(_T("%s - %s"), rs2.GetValue(_T("detailaddress")), rs2.GetValue(_T("address")));
	rpt.GetReportHeader()->SetValue(_T("Address"), tmpStr);

	rs2.GetValue(_T("objectname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("objectname"), tmpStr);

	rs2.GetValue(_T("workplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Workplace"), tmpStr);

	rs2.GetValue(_T("cardno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CardNo"), tmpStr);
	rs2.GetValue(_T("regdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RegDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	rs2.GetValue(_T("expdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ExpiryDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	

	CString szDate;
	rs.GetValue(_T("hpc_orderdate"), tmpStr);
	tmpStr2 = rpt.GetReportHeader()->GetValue(_T("OrderDate"));
	szDate.Format(tmpStr2, tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);

	rs.GetValue(_T("hpc_doctor"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DoctorName"), GetDoctorName(tmpStr));
	CReportItem *imgHeader = rpt.GetReportHeader()->GetItem(_T("Signature"));
	if(imgHeader)
	{
		imgHeader->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		imgHeader->SetFixedHeight(false);		
	}	
	


	CString szDept;
	CRecord rsd(&pMF->m_db);
	rs.GetValue(_T("hpc_roomid"), tmpStr);
	
	rpt.GetReportHeader()->SetValue(_T("Room"), nRoomID);
	rs.GetValue(_T("hpc_deptid"), szDept);
	szSQL.Format(_T("SELECT hrl_name as roomname FROM hms_roomlist WHERE hrl_deptid='%s' AND hrl_id=%d"), szDept, nRoomID);
	rsd.ExecSQL(szSQL);
	if(!rsd.IsEOF())
	{
		rsd.GetValue(_T("roomname"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("RoomName"), tmpStr);
	}

	
	rs.GetValue(_T("hpc_deptid"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), pMF->GetDepartmentName(tmpStr));
	
	rs.GetValue(_T("hpc_bedid"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Bed"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("BedName"), _T(""));

	
	szSQL.Format(_T("SELECT he_prediagnostic as Diagnostic FROM hms_exam WHERE he_docno=%ld AND he_roomid=%d"), nDocumentNo, nRoomID);
	rs3.ExecSQL(szSQL);
	if(!rs3.IsEOF())		  
	{
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), rs3.GetValue(_T("Diagnostic")));
	}
	
	rs2.GetValue(_T("diagnostic"), tmpStr);
	if(tmpStr.GetLength() > 5)
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);

	rs2.GetValue(_T("relationdisease"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("relationdisease"), tmpStr);
	
	

	szSQL.Format(_T(" SELECT hfl_name   AS name,") \
	_T("      hfl_unit     AS unit,") \
	_T("      hfl_index1   AS index1,") \
	_T("      hfl_index2   AS index2,") \
	_T("      hpcl_result  AS result,") \
	_T("      hpcl_note    AS note,") \
	_T("      hfl_subitem AS subitem") \
	_T(" FROM hms_testorderline") \
	_T(" LEFT JOIN hms_fee_list") \
	_T(" ON(hfl_feeid                =hpcl_itemid)") \
	_T(" WHERE hpcl_orderid          = %ld") \
	_T(" and (hpcl_status IN('P','T') or hpcl_hasfee='Y') ") \
	_T(" ORDER BY hpcl_orderlineid"), nOrderID);


	CRecord rsl(&pMF->m_db);
	ret = rsl.ExecSQL(szSQL);
//_fmsg(_T("%s"), szSQL);
	CReportSection* rptDetail;
	if(szGroupID.Left(2) == _T("B1")){
/*
		rsl.MoveFirst();
		if(ret > 30)
		{
			for (int i =0; i <= ret/2; i++)
			{
				rptDetail = rpt.AddDetail();
				rsl.GetValue(_T("name"), tmpStr);
				rptDetail->SetValue(_T("1"), tmpStr);
				rsl.GetValue(_T("index1"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rsl.GetValue(_T("result"), tmpStr);
				rptDetail->SetValue(_T("3"), tmpStr);
				rptDetail->SetValue(_T("4"), _T(""));
				rptDetail->SetValue(_T("5"), _T(""));
				rptDetail->SetValue(_T("6"), _T(""));
				rsl.MoveNext();
			}
			for (int i =0; i <= ret/2 ; i++)
			{
				if(rsl.IsEOF())
					break;
				rptDetail = rpt.GetDetail(i+1);
				rsl.GetValue(_T("name"), tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				rsl.GetValue(_T("index1"), tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				rsl.GetValue(_T("result"), tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);
				rsl.MoveNext();
			}
		}
		else
		{
*/
			int nItem = 1;
			while(!rsl.IsEOF())
			{							
				rsl.GetValue(_T("subitem"), tmpStr);
				if (tmpStr.GetLength() <= 1)
				{	
					//rpt.GetDetail(0)->GetItem(_T("1"))->SetFaceSize(10);
					//rpt.GetDetail(0)->GetItem(_T("1"))->SetItalic(false);
					tmpStr.Format(_T(" %d. %s"), nItem++, rsl.GetValue(_T("name")));
				}
				else
				{	
					rpt.GetDetail(0)->GetItem(_T("1"))->SetFaceSize(10);
					rpt.GetDetail(0)->GetItem(_T("1"))->SetItalic(true);
					tmpStr.Format(_T("     - %s"), rsl.GetValue(_T("name")));
				}						

				rptDetail = rpt.AddDetail();
				rptDetail->SetValue(_T("1"), tmpStr);		
				if(szSex == _T("F"))
					rsl.GetValue(_T("index2"), tmpStr);
				else
					rsl.GetValue(_T("index1"), tmpStr);
				if(tmpStr == _T("0") || tmpStr == _T("0 - 0") || tmpStr == _T("0-0"))
					tmpStr.Empty();
				rptDetail->SetValue(_T("2"), tmpStr);
				
				tmpStr.Empty();
				
				rsl.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				
				
				rsl.MoveNext();
			}
	}
	else
	{
			int nItem = 1;
			while(!rsl.IsEOF())
			{
				rptDetail = rpt.AddDetail();
				tmpStr.Format(_T("%d. %s"), nItem++, rsl.GetValue(_T("name")));
				rptDetail->SetValue(_T("1"), tmpStr);
				rsl.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rsl.MoveNext();
			}

	}
	

	//Page Footer
	//Report Footer
	
	rs.GetValue(_T("hpc_orderdate"), tmpStr);
	tmpStr2 = rpt.GetReportFooter()->GetValue(_T("OrderDate"));
	szDate.Format(tmpStr2, tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("OrderDate"), szDate);
	
	CReportItem *Picture = NULL;
	
	Picture = rpt.GetReportFooter()->GetItem(_T("Picture"));
	if(Picture)
	{	
		HBITMAP hBitmap = pMF->GetPACSImage(pMF->m_szCurrentFileNamePACS);
		if (hBitmap)
		{
			Picture->SetHBITMAP(hBitmap);
			Picture->SetFixedHeight(false);
			::DeleteObject(hBitmap);
		}
	}
	Picture = rpt.GetReportHeader()->GetItem(_T("Picture"));
	
	if(Picture)
	{	
		HBITMAP hBitmap = pMF->GetPACSImage(pMF->m_szCurrentFileNamePACS);
		if (hBitmap)
		{
			Picture->SetHBITMAP(hBitmap);
			Picture->SetFixedHeight(false);
			::DeleteObject(hBitmap);
		}
	}

	if(szStatus != _T("T"))
		rs.GetValue(_T("hpc_doctor"), tmpStr);
	else
		rs.GetValue(_T("hpc_practitioner"), tmpStr);

	rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(tmpStr));
	CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
	if(img)
	{
		img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img->SetFixedHeight(false);		
	}	
	

	rs.GetValue(_T("hpc_performdate"), tmpStr);
	tmpStr2 = rpt.GetReportFooter()->GetValue(_T("PerformDate"));
	if(CDateTime::IsValid(tmpStr))
		szDate.Format(tmpStr2,tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	else
		szDate.Format(tmpStr2, _T("....."), _T("....."), _T("........"));
	rpt.GetReportFooter()->SetValue(_T("PerformDate"), szDate);
	rs.GetValue(_T("hpc_practitioner"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("Practitioner"), GetDoctorName(tmpStr));
	CReportItem *img2 = rpt.GetReportFooter()->GetItem(_T("Signature2"));
	if(img2)
	{
		img2->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img2->SetFixedHeight(false);
	}


	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print(bDirect);
}



void CPrintForms::PrintTestResultX(long nOrderID, bool bPreview, bool bDirect){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szGroupID, szSex, szStatus;
	CString szReportName, szOrderName;
	CString szSQL;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CRecord rs3(&pMF->m_db);
	long nDocumentNo;
	int nRefIndex, nRoomID;
	int nNumberOrder=0;
	CString tmpStr2, szDeptType;
	CString szFormID;

	if(pMF->GetModuleID() != _T("EM"))
		return;


	szSQL.Format(_T("SELECT * FROM hms_testorder WHERE hpc_orderid=%ld "), nOrderID);

	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;

	rs.GetValue(_T("hpc_groupid"), szGroupID);
	rs.GetValue(_T("hpc_docno"), nDocumentNo);
	rs.GetValue(_T("hpc_refidx"), nRefIndex);
	rs.GetValue(_T("hpc_roomid"), nRoomID);
	rs.GetValue(_T("hpc_status"), szStatus);
	rs.GetValue(_T("hpc_depttype"), szDeptType);

	szSQL.Format(_T("SELECT hfg_report as reportname FROM hms_fee_group WHERE hfg_id='%s' "), szGroupID);
	rs2.ExecSQL(szSQL);
	if(!rs2.IsEOF())
		rs2.GetValue(_T("reportname"), tmpStr);

	if(tmpStr.IsEmpty())
		tmpStr = _T("PACS_VOTE.RPT");
	tmpStr.Trim();

	szReportName.Format(_T("Reports/HMS/%s"), tmpStr);

	szReportName.Trim();

	if(!rpt.Init(szReportName) )
	{
		return;
	}
	 
	
		
	szSQL.Format(_T(" SELECT 	hd_patientno as patientno,  ") \
		_T(" 	hd_docno as docno,") \
		_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
		_T(" 	CAST(TO_CHAR(hp_birthdate, 'YYYY') as integer) as birthyear,") \
		_T(" 	hms_getage(trunc_date(hd_admitdate), hp_birthdate) as age, 	hp_sex as sexid,") \
		_T(" 	sys_sel_getname('sys_sex', hp_sex) as sex,") \
		_T(" 	hp_occupation as occupationid,") \
		_T(" 	sys_sel_getname('sys_occupation', cast(hp_occupation as varchar(15))) as occupation,") \
		_T(" 	hp_dtladdr as detailaddress,") \
		_T(" 	hms_getaddress(hp_provid,hp_distid,hp_villid) as address,") \
		_T(" 	hp_workplaceid as workplaceid,") \
		_T(" 	hp_workplace as workplace,") \
		_T(" 	hd_object as objectid,") \
		_T(" 	(SELECT ho_desc FROM hms_object WHERE ho_id=hd_object) as objectname,") \
		_T(" 	hd_cardno as cardno,") \
		_T(" 	hd_cardidx as cardidx, ") \
		_T(" 	hd_reldisease as relationdisease, ") \
		_T("	trim(hd_diagnostic) ||' ['||trim(hd_icd)||']' as diagnostic, ") \
		_T(" 	hc_regdate as regdate, ") \
		_T(" 	hc_expdate as expdate ") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx) ") \
		_T(" WHERE hd_docno=%ld"), nDocumentNo);
//_fmsg(_T("%s"), szSQL);
	int ret = rs2.ExecSQL(szSQL);
	if(rs2.IsEOF())
	{
		ShowMessageBox(_T("Patient not found."));
		return;
	}

	//Report Header
	tmpStr.Empty();
	rs2.GetValue(_T("sexid"), szSex);
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HealthService"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), tmpStr);
	rs.GetValue(_T("hpc_pdeptid"), tmpStr);
	if(!tmpStr.IsEmpty())
	{
		CString szPerformDept;
		StringUpper(pMF->GetDepartmentName(tmpStr), szPerformDept);
		rpt.GetReportHeader()->SetValue(_T("Faculty"), szPerformDept);
	}
	else
		rpt.GetReportHeader()->SetValue(_T("Faculty"), _T(""));

	rs.GetValue(_T("hpc_docno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128A]"),tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128B]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128C]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[39]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[93]"), tmpStr);

	tmpStr.Format(_T("%ld"), nOrderID);
	rpt.GetReportHeader()->SetValue(_T("SheetNo"), tmpStr);
	StringUpper(rs2.GetValue(_T("pname")), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("PatientName"), tmpStr);
	rs2.GetValue(_T("age"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), tmpStr);

	//Chi In ra nam sinh.
	rs2.GetValue(_T("birthyear"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);
//	_msg(_T("%s"), tmpStr);
	rs2.GetValue(_T("sex"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Sex"), tmpStr);
	tmpStr.Format(_T("%s - %s"), rs2.GetValue(_T("detailaddress")), rs2.GetValue(_T("address")));
	rpt.GetReportHeader()->SetValue(_T("Address"), tmpStr);

	rs2.GetValue(_T("objectname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("objectname"), tmpStr);

	rs2.GetValue(_T("workplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Workplace"), tmpStr);

	rs2.GetValue(_T("cardno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CardNo"), tmpStr);
	rs2.GetValue(_T("regdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RegDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	rs2.GetValue(_T("expdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ExpiryDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	

	CString szDate;
	rs.GetValue(_T("hpc_orderdate"), tmpStr);
	tmpStr2 = rpt.GetReportHeader()->GetValue(_T("OrderDate"));
	szDate.Format(tmpStr2, tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);

	rs.GetValue(_T("hpc_doctor"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DoctorName"), GetDoctorName(tmpStr));
	CReportItem *imgHeader = rpt.GetReportHeader()->GetItem(_T("Signature"));
	if(imgHeader)
	{
		imgHeader->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		imgHeader->SetFixedHeight(false);		
	}	
	


	CString szDept;
	CRecord rsd(&pMF->m_db);
	rs.GetValue(_T("hpc_roomid"), tmpStr);
	
	rpt.GetReportHeader()->SetValue(_T("Room"), nRoomID);
	rs.GetValue(_T("hpc_deptid"), szDept);
	szSQL.Format(_T("SELECT hrl_name as roomname FROM hms_roomlist WHERE hrl_deptid='%s' AND hrl_id=%d"), szDept, nRoomID);
	rsd.ExecSQL(szSQL);
	if(!rsd.IsEOF())
	{
		rsd.GetValue(_T("roomname"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("RoomName"), tmpStr);
	}

	
	rs.GetValue(_T("hpc_deptid"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), pMF->GetDepartmentName(tmpStr));
	
	rs.GetValue(_T("hpc_bedid"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Bed"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("BedName"), _T(""));

	
	szSQL.Format(_T("SELECT he_prediagnostic as Diagnostic FROM hms_exam WHERE he_docno=%ld AND he_roomid=%d"), nDocumentNo, nRoomID);
	rs3.ExecSQL(szSQL);
	if(!rs3.IsEOF())		  
	{
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), rs3.GetValue(_T("Diagnostic")));
	}
	
	rs2.GetValue(_T("diagnostic"), tmpStr);
	if(tmpStr.GetLength() > 5)
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);

	rs2.GetValue(_T("relationdisease"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("relationdisease"), tmpStr);
	
	

	szSQL.Format(_T(" SELECT hfl_name   AS name,") \
	_T("      hfl_unit     AS unit,") \
	_T("      hfl_index1   AS index1,") \
	_T("      hfl_index2   AS index2,") \
	_T("      hpcl_result  AS result,") \
	_T("      hpcl_note    AS note,") \
	_T("      hfl_subitem AS subitem") \
	_T(" FROM hms_testorderline") \
	_T(" LEFT JOIN hms_fee_list") \
	_T(" ON(hfl_feeid                =hpcl_itemid)") \
	_T(" WHERE hpcl_orderid          = %ld") \
	_T(" and (hpcl_status IN('P','T') or hpcl_hasfee='Y') ") \
	_T(" ORDER BY hpcl_orderlineid"), nOrderID);


	CRecord rsl(&pMF->m_db);
	ret = rsl.ExecSQL(szSQL);
//_fmsg(_T("%s"), szSQL);
	CReportSection* rptDetail;
	if(szGroupID.Left(2) == _T("B1")){
/*
		rsl.MoveFirst();
		if(ret > 30)
		{
			for (int i =0; i <= ret/2; i++)
			{
				rptDetail = rpt.AddDetail();
				rsl.GetValue(_T("name"), tmpStr);
				rptDetail->SetValue(_T("1"), tmpStr);
				rsl.GetValue(_T("index1"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rsl.GetValue(_T("result"), tmpStr);
				rptDetail->SetValue(_T("3"), tmpStr);
				rptDetail->SetValue(_T("4"), _T(""));
				rptDetail->SetValue(_T("5"), _T(""));
				rptDetail->SetValue(_T("6"), _T(""));
				rsl.MoveNext();
			}
			for (int i =0; i <= ret/2 ; i++)
			{
				if(rsl.IsEOF())
					break;
				rptDetail = rpt.GetDetail(i+1);
				rsl.GetValue(_T("name"), tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				rsl.GetValue(_T("index1"), tmpStr);
				rptDetail->SetValue(_T("5"), tmpStr);
				rsl.GetValue(_T("result"), tmpStr);
				rptDetail->SetValue(_T("6"), tmpStr);
				rsl.MoveNext();
			}
		}
		else
		{
*/
			int nItem = 1;
			while(!rsl.IsEOF())
			{							
				rsl.GetValue(_T("subitem"), tmpStr);
				if (tmpStr.GetLength() <= 1)
				{	
					//rpt.GetDetail(0)->GetItem(_T("1"))->SetFaceSize(10);
					//rpt.GetDetail(0)->GetItem(_T("1"))->SetItalic(false);
					tmpStr.Format(_T(" %d. %s"), nItem++, rsl.GetValue(_T("name")));
				}
				else
				{	
					rpt.GetDetail(0)->GetItem(_T("1"))->SetFaceSize(10);
					rpt.GetDetail(0)->GetItem(_T("1"))->SetItalic(true);
					tmpStr.Format(_T("     - %s"), rsl.GetValue(_T("name")));
				}						

				rptDetail = rpt.AddDetail();
				rptDetail->SetValue(_T("1"), tmpStr);		
				if(szSex == _T("F"))
					rsl.GetValue(_T("index2"), tmpStr);
				else
					rsl.GetValue(_T("index1"), tmpStr);
				if(tmpStr == _T("0") || tmpStr == _T("0 - 0") || tmpStr == _T("0-0"))
					tmpStr.Empty();
				rptDetail->SetValue(_T("2"), tmpStr);
				
				rsl.GetValue(_T("result"), tmpStr);
				rptDetail->SetValue(_T("3"), tmpStr);

				rsl.GetValue(_T("note"), tmpStr);
				if(tmpStr == _T("H") || tmpStr == _T("L"))
				{
					rptDetail->SetValue(_T("5"), tmpStr);
				}
				
				rsl.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("4"), tmpStr);
				
				
				rsl.MoveNext();
			}
	}
	else
	{
			int nItem = 1;
			while(!rsl.IsEOF())
			{
				rptDetail = rpt.AddDetail();
				tmpStr.Format(_T("%d. %s"), nItem++, rsl.GetValue(_T("name")));
				rptDetail->SetValue(_T("1"), tmpStr);
				rsl.GetValue(_T("unit"), tmpStr);
				rptDetail->SetValue(_T("2"), tmpStr);
				rsl.MoveNext();
			}

	}
	

	//Page Footer
	//Report Footer
	
	rs.GetValue(_T("hpc_orderdate"), tmpStr);
	tmpStr2 = rpt.GetReportFooter()->GetValue(_T("OrderDate"));
	szDate.Format(tmpStr2, tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("OrderDate"), szDate);
	
	CReportItem *Picture = NULL;
	
	Picture = rpt.GetReportFooter()->GetItem(_T("Picture"));
	if(Picture)
	{	
		HBITMAP hBitmap = pMF->GetPACSImage(pMF->m_szCurrentFileNamePACS);
		if (hBitmap)
		{
			Picture->SetHBITMAP(hBitmap);
			Picture->SetFixedHeight(false);
			::DeleteObject(hBitmap);
		}
	}
	Picture = rpt.GetReportHeader()->GetItem(_T("Picture"));
	
	if(Picture)
	{	
		HBITMAP hBitmap = pMF->GetPACSImage(pMF->m_szCurrentFileNamePACS);
		if (hBitmap)
		{
			Picture->SetHBITMAP(hBitmap);
			Picture->SetFixedHeight(false);
			::DeleteObject(hBitmap);
		}
	}

	if(szStatus != _T("T"))
		rs.GetValue(_T("hpc_doctor"), tmpStr);
	else
		rs.GetValue(_T("hpc_practitioner"), tmpStr);

	rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(tmpStr));
	CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
	if(img)
	{
		img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img->SetFixedHeight(false);		
	}	
	

	rs.GetValue(_T("hpc_performdate"), tmpStr);
	tmpStr2 = rpt.GetReportFooter()->GetValue(_T("PerformDate"));
	if(CDateTime::IsValid(tmpStr))
		szDate.Format(tmpStr2,tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	else
		szDate.Format(tmpStr2, _T("....."), _T("....."), _T("........"));
	rpt.GetReportFooter()->SetValue(_T("PerformDate"), szDate);
	rs.GetValue(_T("hpc_practitioner"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("Practitioner"), GetDoctorName(tmpStr));
	CReportItem *img2 = rpt.GetReportFooter()->GetItem(_T("Signature2"));
	if(img2)
	{
		img2->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img2->SetFixedHeight(false);
	}


	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print(bDirect);
}


void CPrintForms::PrintPACSOrderX(long nOrderID,  CString szItemID, bool bPreview, bool bDirect, bool bDownload){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szGroupID, szSex, szStatus;
	CString szReportName, szOrderName;
	CString szSQL;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CRecord rs3(&pMF->m_db);
	long nDocumentNo;
	int nRefIndex, nRoomID;
	
	CString tmpStr2, szDeptType;
	CString szFormID;
	CString szOrderDate;
	szSQL.Format(_T("SELECT * FROM hms_pacsorder LEFT JOIN hms_pacsorderline ON(hpcl_orderid=hpc_orderid) WHERE hpc_orderid=%ld and hpcl_itemid='%s' "), nOrderID, szItemID);
//_msg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	
	
	rs.GetValue(_T("hpc_groupid"), szGroupID);
	rs.GetValue(_T("hpc_docno"), nDocumentNo);
	rs.GetValue(_T("hpc_refidx"), nRefIndex);
	rs.GetValue(_T("hpc_status"), szStatus);
	rs.GetValue(_T("hpc_roomid"), nRoomID);
	rs.GetValue(_T("hpc_depttype"), szDeptType);
	rs.GetValue(_T("hpc_orderdate"), szOrderDate);

	szItemID.Empty();
	
	szSQL.Format(_T("SELECT hfg_report as reportname FROM hms_fee_group WHERE hfg_id='%s' "), szGroupID);
	rs2.ExecSQL(szSQL);
	tmpStr.Empty();
	if(!rs2.IsEOF())
		rs2.GetValue(_T("reportname"), tmpStr);

	if(tmpStr.IsEmpty())
		tmpStr = _T("PACS_VOTE.RPT");
	tmpStr.Trim();

	tmpStr.Replace(_T("_"), _T("O_"));
	szReportName.Format(_T("Reports/HMS/%s"), tmpStr);
	

	szSQL.Format(_T(" SELECT hfl_name as name, hfl_unit as unit") \
		_T(" FROM hms_pacsorderline") \
		_T(" LEFT JOIN hms_fee_list ON(hpcl_itemid=hfl_feeid)") \
		_T(" WHERE 	hpcl_orderid=%ld ") \
		_T(" 	and hpcl_result='%s'") \
		_T(" ORDER BY hpcl_orderlineid "), nOrderID, szFormID);

	rs2.ExecSQL(szSQL);

	while(!rs2.IsEOF())
	{
		rs2.GetValue(_T("name"), tmpStr2);
		if(!szOrderName.IsEmpty())
			szOrderName += _T("\r\n");
		szOrderName.AppendFormat(_T("%s"), tmpStr2);
		rs2.MoveNext();		
	}
	
	szReportName.Trim();

	if(!rpt.Init(szReportName) )
	{
		return;
	}
	 

		
	szSQL.Format(_T(" SELECT 	hd_patientno as patientno,  ") \
		_T(" 	hd_docno as docno,") \
		_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
		_T(" 	CAST(TO_CHAR(hp_birthdate, 'YYYY') as integer) as birthyear,") \
		_T(" 	hms_getage(trunc_date(hd_admitdate), hp_birthdate) as age, 	hp_sex as sexid,") \
		_T(" 	sys_sel_getname('sys_sex', hp_sex) as sex,") \
		_T(" 	hp_occupation as occupationid,") \
		_T(" 	sys_sel_getname('sys_occupation', cast(hp_occupation as nvarchar2(15))) as occupation,") \
		_T(" 	hp_dtladdr as detailaddress,") \
		_T(" 	hms_getaddress(hp_provid,hp_distid,hp_villid) as address,") \
		_T(" 	hp_workplaceid as workplaceid,") \
		_T(" 	hp_workplace as workplace,") \
		_T(" 	hd_object as objectid,") \
		_T("	trim(hd_diagnostic) ||' ['||trim(hd_icd)||']' as diagnostic, ") \
		_T(" 	(SELECT ho_desc FROM hms_object WHERE ho_id=hd_object) as objectname,") \
		_T(" 	hd_cardno as cardno,") \
		_T(" 	hd_cardidx as cardidx, ") \
		_T(" 	hd_reldisease as relationdisease, ") \
		_T(" 	hc_regdate as regdate, ") \
		_T(" 	hc_expdate as expdate ") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx) ") \
		_T(" WHERE hd_docno=%ld"), nDocumentNo);
//_fmsg(_T("%s"), szSQL);
	int ret = rs2.ExecSQL(szSQL);		
	if(rs2.IsEOF())
	{
		ShowMessageBox(_T("Patient not found."));
		return;
	}


	//Report Header
	tmpStr.Empty();
	rs2.GetValue(_T("sexid"), szSex);
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HealthService"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), tmpStr);
	rs.GetValue(_T("hpc_pdeptid"), tmpStr);
	if(!tmpStr.IsEmpty())
	{
		CString szPerformDept;
		StringUpper(pMF->GetDepartmentName(tmpStr), szPerformDept);
		rpt.GetReportHeader()->SetValue(_T("Faculty"), szPerformDept);
	}
	else
		rpt.GetReportHeader()->SetValue(_T("Faculty"), _T(""));

	rs.GetValue(_T("hpc_docno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128A]"),tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128B]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128C]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[39]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[93]"), tmpStr);


	tmpStr.Format(_T("%ld"), nOrderID);
	rpt.GetReportHeader()->SetValue(_T("SheetNo"), tmpStr);
	StringUpper(rs2.GetValue(_T("pname")), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("PatientName"), tmpStr);
	rs2.GetValue(_T("age"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), tmpStr);

	//Chi In ra nam sinh.
	rs2.GetValue(_T("birthyear"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);

	rs2.GetValue(_T("sex"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Sex"), tmpStr);
	tmpStr.Format(_T("%s - %s"), rs2.GetValue(_T("detailaddress")), rs2.GetValue(_T("address")));
	rpt.GetReportHeader()->SetValue(_T("Address"), tmpStr);

	rs2.GetValue(_T("objectname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("objectname"), tmpStr);

	rs2.GetValue(_T("workplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Workplace"), tmpStr);

	rs2.GetValue(_T("cardno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CardNo"), tmpStr);
	rs2.GetValue(_T("regdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RegDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	rs2.GetValue(_T("expdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ExpiryDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	

	CString szDate;
	rs.GetValue(_T("hpc_orderdate"), tmpStr);
	tmpStr2 = rpt.GetReportHeader()->GetValue(_T("OrderDate"));

	szDate.Format(tmpStr2, tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);

	rs.GetValue(_T("hpc_doctor"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DoctorName"), GetDoctorName(tmpStr));
	CReportItem *imgHeader = rpt.GetReportHeader()->GetItem(_T("Signature"));
	if(imgHeader)
	{
		imgHeader->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		imgHeader->SetFixedHeight(false);		
	}	
	


	CString szDept;
	CRecord rsd(&pMF->m_db);	
	
	rpt.GetReportHeader()->SetValue(_T("Room"), nRoomID);

	rs.GetValue(_T("hpc_deptid"), szDept);
	szSQL.Format(_T("SELECT hrl_name as roomname FROM hms_roomlist WHERE hrl_deptid='%s' AND hrl_id=%d"), szDept, nRoomID);
	rsd.ExecSQL(szSQL);
	if(!rsd.IsEOF())
	{
		rsd.GetValue(_T("roomname"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("RoomName"), tmpStr);
	}

	
	rs.GetValue(_T("hpc_deptid"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), pMF->GetDepartmentName(tmpStr));


	szSQL.Format(_T("SELECT he_prediagnostic as Diagnostic FROM hms_exam WHERE he_docno=%ld AND he_roomid=%d"), nDocumentNo, nRoomID);
	rs3.ExecSQL(szSQL);

	if(!rs3.IsEOF())
	{
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), rs3.GetValue(_T("Diagnostic")));		
	}	
	

	rs2.GetValue(_T("diagnostic"), tmpStr);
	if(tmpStr.GetLength() > 5)
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);

	rs2.GetValue(_T("relationdisease"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("relationdisease"), tmpStr);
	
	//Page Header

	//Report Detail
	CRecord rsl(&pMF->m_db);

	szSQL.Format(_T(" SELECT 	hfl_name as name, ") \
		_T(" 	hfl_unit as unit, ") \
		_T(" 	hfl_index1 as index1, ") \
		_T(" 	hfl_index2 as index2,") \
		_T(" 	hpcl_result as result,") \
		_T(" 	hpcl_note as note,hfl_subitem as subitem, hpcl_qty as qty, hpcl_note ") \
		_T(" FROM hms_pacsorderline ") \
		_T(" LEFT JOIN hms_fee_list ON(hfl_feeid=hpcl_itemid)") \
		_T(" WHERE hpcl_orderid=%ld ") \
		_T(" and (hpcl_hasfee='Y') ") \
		_T(" ORDER BY hpcl_orderlineid "), nOrderID);


	ret = rsl.ExecSQL(szSQL);
	CReportSection* rptDetail;
	int nItem = 1;
	int nQty;
	CString szNote, szUnit;
	while(!rsl.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rsl.GetValue(_T("qty"), nQty);
		rsl.GetValue(_T("hpcl_note"), szNote);
		szNote.Trim();

		tmpStr.Format(_T("%d. %s"), nItem++, rsl.GetValue(_T("name")));
		if (!szNote.IsEmpty())
		{
			tmpStr.AppendFormat(_T(" (%s)"), szNote);
		}
		if(nQty > 1)
		{
			tmpStr.AppendFormat(_T("; S\x1ED1 l\x1B0\x1EE3ng: %d"), nQty);
		}
		rptDetail->SetValue(_T("1"), tmpStr);
		rsl.GetValue(_T("unit"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rsl.MoveNext();
	}
//_fmsg(_T("%s"), szSQL);
	
	

	//Page Footer
	//Report Footer
	
	rs.GetValue(_T("hpc_orderdate"), tmpStr);
	tmpStr2 = rpt.GetReportFooter()->GetValue(_T("OrderDate"));
	szDate.Format(tmpStr2, tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("OrderDate"), szDate);
	
	CReportItem *Picture = NULL;
	
	Picture = rpt.GetReportFooter()->GetItem(_T("Picture"));

	HBITMAP hBitmap = NULL;
	if(Picture && bDownload)
	{	

		hBitmap = pMF->DownloadImage(nDocumentNo, nOrderID, szItemID, Picture->GetItemRect().Width(), Picture->GetItemRect().Height(), Picture->m_nCol1, Picture->m_nRow1);
		if (hBitmap)
		{
			Picture->SetHBITMAP(hBitmap);
			Picture->SetFixedHeight(false);
		}
	}
	Picture = rpt.GetReportHeader()->GetItem(_T("Picture"));	
	if(Picture && bDownload)
	{	

		//HBITMAP hBitmap = pMF->GetPACSImage(pMF->m_szCurrentFileNamePACS);
		if(hBitmap == NULL) 
			hBitmap = pMF->DownloadImage(nDocumentNo, nOrderID, szItemID, Picture->GetItemRect().Width(), Picture->GetItemRect().Height(), Picture->m_nCol1, Picture->m_nRow1);
		if (hBitmap)
		{			
			Picture->SetHBITMAP(hBitmap);
			Picture->SetFixedHeight(false);
			
		}
	}
	if(hBitmap) ::DeleteObject(hBitmap);


	rs.GetValue(_T("hpc_doctor"), tmpStr);



	rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(tmpStr));
	CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
	if(img)
	{
		img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img->SetFixedHeight(false);		
	}	
	

	rs.GetValue(_T("hpc_performdate"), tmpStr);
	tmpStr2 = rpt.GetReportFooter()->GetValue(_T("PerformDate"));
	if(CDateTime::IsValid(tmpStr))
		szDate.Format(tmpStr2,tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	else
		szDate.Format(tmpStr2, _T("....."), _T("....."), _T("........"));
	rpt.GetReportFooter()->SetValue(_T("PerformDate"), szDate);
	rs.GetValue(_T("hpcl_practitioner"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("Practitioner"), GetDoctorName(tmpStr));
	CReportItem *img2 = rpt.GetReportFooter()->GetItem(_T("Signature2"));
	if(img2)
	{
		img2->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img2->SetFixedHeight(false);
	}
	
	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print(bDirect);
}

void CPrintForms::PrintPACSResultX(long nOrderID,  CString szItemID, bool bPreview, bool bDirect, bool bDownload){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szGroupID, szSex, szStatus;
	CString szReportName, szOrderName;
	CString szSQL;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CRecord rs3(&pMF->m_db);
	long nDocumentNo;
	int nRefIndex, nRoomID;
	
	CString tmpStr2, szDeptType;
	CString szFormID;
	CString szOrderDate;
	szSQL.Format(_T("SELECT * FROM hms_pacsorder LEFT JOIN hms_pacsorderline ON(hpcl_orderid=hpc_orderid) WHERE hpc_orderid=%ld and hpcl_itemid='%s' "), nOrderID, szItemID);
//_msg(_T("%s"), szSQL);
	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	
	
	rs.GetValue(_T("hpc_groupid"), szGroupID);
	rs.GetValue(_T("hpc_docno"), nDocumentNo);
	rs.GetValue(_T("hpc_refidx"), nRefIndex);
	rs.GetValue(_T("hpc_status"), szStatus);
	rs.GetValue(_T("hpc_roomid"), nRoomID);
	rs.GetValue(_T("hpc_depttype"), szDeptType);
	rs.GetValue(_T("hpc_orderdate"), szOrderDate);

	szSQL.Format(_T(" SELECT hpf_id as id, hpf_reportname as reportname") \
	_T(" FROM hms_pacsorderline") \
	_T(" LEFT JOIN hms_pacs_form ON(hpf_id=hpcl_result)") \
	_T(" WHERE hpcl_orderid=%ld and hpcl_itemid='%s'"), nOrderID, szItemID);
	
	rs2.ExecSQL(szSQL);
	if(!rs2.IsEOF())
	{
		rs2.GetValue(_T("id"), szFormID);
	}

	tmpStr.Empty();
	if(!rs2.IsEOF())
		rs2.GetValue(_T("reportname"), tmpStr);

	if(tmpStr.IsEmpty())
		tmpStr = _T("PACS_VOTE.RPT");
	tmpStr.Trim();

	tmpStr.Replace(_T("_"), _T("O_"));
	szReportName.Format(_T("Reports/HMS/%s"), tmpStr);

	szSQL.Format(_T(" SELECT hfl_name as name, hfl_unit as unit") \
		_T(" FROM hms_pacsorderline") \
		_T(" LEFT JOIN hms_fee_list ON(hpcl_itemid=hfl_feeid)") \
		_T(" WHERE 	hpcl_orderid=%ld ") \
		_T(" 	and hpcl_result='%s'") \
		_T(" ORDER BY hpcl_orderlineid "), nOrderID, szFormID);
	rs2.ExecSQL(szSQL);

	while(!rs2.IsEOF())
	{
		rs2.GetValue(_T("name"), tmpStr2);
		if(!szOrderName.IsEmpty())
			szOrderName += _T("\r\n");
		szOrderName.AppendFormat(_T("%s"), tmpStr2);
		rs2.MoveNext();		
	}
	
	szReportName.Trim();

	if(!rpt.Init(szReportName) )
	{
		return;
	}
		
	szSQL.Format(_T(" SELECT 	hd_patientno as patientno,  ") \
		_T(" 	hd_docno as docno,") \
		_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
		_T(" 	CAST(TO_CHAR(hp_birthdate, 'YYYY') as integer) as birthyear,") \
		_T(" 	hms_getage(trunc_date(hd_admitdate), hp_birthdate) as age, 	hp_sex as sexid,") \
		_T(" 	sys_sel_getname('sys_sex', hp_sex) as sex,") \
		_T(" 	hp_occupation as occupationid,") \
		_T(" 	sys_sel_getname('sys_occupation', cast(hp_occupation as nvarchar2(15))) as occupation,") \
		_T(" 	hp_dtladdr as detailaddress,") \
		_T(" 	hms_getaddress(hp_provid,hp_distid,hp_villid) as address,") \
		_T(" 	hp_workplaceid as workplaceid,") \
		_T(" 	hp_workplace as workplace,") \
		_T(" 	hd_object as objectid,") \
		_T("	trim(hd_diagnostic) ||' ['||trim(hd_icd)||']' as diagnostic, ") \
		_T(" 	(SELECT ho_desc FROM hms_object WHERE ho_id=hd_object) as objectname,") \
		_T(" 	hd_cardno as cardno,") \
		_T(" 	hd_cardidx as cardidx, ") \
		_T(" 	hd_reldisease as relationdisease, ") \
		_T(" 	hc_regdate as regdate, ") \
		_T(" 	hc_expdate as expdate ") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx) ") \
		_T(" WHERE hd_docno=%ld"), nDocumentNo);
//_fmsg(_T("%s"), szSQL);
	int ret = rs2.ExecSQL(szSQL);		
	if(rs2.IsEOF())
	{
		ShowMessageBox(_T("Patient not found."));
		return;
	}


	//Report Header
	tmpStr.Empty();
	rs2.GetValue(_T("sexid"), szSex);
	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HealthService"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), tmpStr);
	rs.GetValue(_T("hpc_pdeptid"), tmpStr);
	if(!tmpStr.IsEmpty())
	{
		CString szPerformDept;
		StringUpper(pMF->GetDepartmentName(tmpStr), szPerformDept);
		rpt.GetReportHeader()->SetValue(_T("Faculty"), szPerformDept);
	}
	else
		rpt.GetReportHeader()->SetValue(_T("Faculty"), _T(""));

	rs.GetValue(_T("hpc_docno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128A]"),tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128B]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128C]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[39]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[93]"), tmpStr);


	tmpStr.Format(_T("%ld"), nOrderID);
	rpt.GetReportHeader()->SetValue(_T("SheetNo"), tmpStr);
	StringUpper(rs2.GetValue(_T("pname")), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("PatientName"), tmpStr);
	rs2.GetValue(_T("age"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), tmpStr);

	//Chi In ra nam sinh.
	rs2.GetValue(_T("birthyear"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);

	rs2.GetValue(_T("sex"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Sex"), tmpStr);
	tmpStr.Format(_T("%s - %s"), rs2.GetValue(_T("detailaddress")), rs2.GetValue(_T("address")));
	rpt.GetReportHeader()->SetValue(_T("Address"), tmpStr);

	rs2.GetValue(_T("objectname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("objectname"), tmpStr);

	rs2.GetValue(_T("workplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Workplace"), tmpStr);

	rs2.GetValue(_T("cardno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CardNo"), tmpStr);
	rs2.GetValue(_T("regdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RegDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	rs2.GetValue(_T("expdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ExpiryDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	

	CString szDate;
	rs.GetValue(_T("hpc_orderdate"), tmpStr);
	tmpStr2 = rpt.GetReportHeader()->GetValue(_T("OrderDate"));

	szDate.Format(tmpStr2, tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);

	rs.GetValue(_T("hpc_doctor"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DoctorName"), GetDoctorName(tmpStr));
	CReportItem *imgHeader = rpt.GetReportHeader()->GetItem(_T("Signature"));
	if(imgHeader)
	{
		imgHeader->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		imgHeader->SetFixedHeight(false);		
	}	
	


	CString szDept;
	CRecord rsd(&pMF->m_db);	
	
	rpt.GetReportHeader()->SetValue(_T("Room"), nRoomID);

	rs.GetValue(_T("hpc_deptid"), szDept);
	szSQL.Format(_T("SELECT hrl_name as roomname FROM hms_roomlist WHERE hrl_deptid='%s' AND hrl_id=%d"), szDept, nRoomID);
	rsd.ExecSQL(szSQL);
	if(!rsd.IsEOF())
	{
		rsd.GetValue(_T("roomname"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("RoomName"), tmpStr);
	}

	
	rs.GetValue(_T("hpc_deptid"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), pMF->GetDepartmentName(tmpStr));


	szSQL.Format(_T("SELECT he_prediagnostic as Diagnostic FROM hms_exam WHERE he_docno=%ld AND he_roomid=%d"), nDocumentNo, nRoomID);
	rs3.ExecSQL(szSQL);

	if(!rs3.IsEOF())
	{
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), rs3.GetValue(_T("Diagnostic")));		
	}	
	

	rs2.GetValue(_T("diagnostic"), tmpStr);
	if(tmpStr.GetLength() > 5)
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);

	rs2.GetValue(_T("relationdisease"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("relationdisease"), tmpStr);
	
	//Page Header

	//Report Detail
	CRecord rsl(&pMF->m_db);


	if(szFormID.IsEmpty())
	{
		szSQL.Format(_T("SELECT hpr_name as name, hpr_desc as describe ") \
			_T(" FROM hms_pacsorderline") \
			_T(" ") \
			_T(" LEFT JOIN hms_pacs_result ON(hpr_name=hpl_name AND hpr_orderid=%ld AND hpr_itemid='%s') ") \
			_T(" LEFT JOIN hms_fee_list ON(hfl_feeid=hpr_itemid) ") \
			_T("ORDER BY hfl_line"), nOrderID, szItemID);
	}
	else
	{
		szSQL.Format(_T(" SELECT  ") \
		_T(" 	hpl_name as name, ") \
		_T(" 	hpr_desc as describe  ") \
		_T(" from hms_pacsorderline  ") \
		_T(" left join hms_pacs_form on(hpf_id=hpcl_result)") \
		_T(" left join hms_pacs_layout on(hpl_id=hpf_id)") \
		_T(" left join hms_pacs_result on(hpr_orderid=hpcl_orderid and hpr_name=hpl_name and hpr_itemid=hpcl_itemid)") \
		_T(" WHERE hpcl_orderid=%ld AND hpcl_itemid='%s' and hpl_id='%s' and length(hpr_desc) > 0 "), nOrderID, szItemID, szFormID);
	}
//_fmsg(_T("%s"), szSQL);
	ret = rsl.ExecSQL(szSQL);

	CReportSection *pDetail=NULL;
	CString szDescribe, szName;
	rpt.GetReportHeader()->SetValue(_T("OrderName"), szOrderName);
	while(!rsl.IsEOF())
	{
		rsl.GetValue(_T("name"), szName);
		rsl.GetValue(_T("describe"), szDescribe);
		rpt.GetReportHeader()->SetValue(szName, szDescribe);
		rpt.GetReportFooter()->SetValue(szName, szDescribe);
		rsl.MoveNext();
	}	
	//Page Footer
	//Report Footer
	
	rs.GetValue(_T("hpc_orderdate"), tmpStr);
	tmpStr2 = rpt.GetReportFooter()->GetValue(_T("OrderDate"));
	szDate.Format(tmpStr2, tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("OrderDate"), szDate);
	
	CReportItem *Picture = NULL;
	
	Picture = rpt.GetReportFooter()->GetItem(_T("Picture"));

	HBITMAP hBitmap = NULL;
	if(Picture && bDownload)
	{	

		hBitmap = pMF->DownloadImage(nDocumentNo, nOrderID, szItemID, Picture->GetItemRect().Width(), Picture->GetItemRect().Height(), Picture->m_nCol1, Picture->m_nRow1);
		if (hBitmap)
		{
			Picture->SetHBITMAP(hBitmap);
			Picture->SetFixedHeight(false);
		}
	}
	Picture = rpt.GetReportHeader()->GetItem(_T("Picture"));	
	if(Picture && bDownload)
	{	

		//HBITMAP hBitmap = pMF->GetPACSImage(pMF->m_szCurrentFileNamePACS);
		if(hBitmap == NULL) 
			hBitmap = pMF->DownloadImage(nDocumentNo, nOrderID, szItemID, Picture->GetItemRect().Width(), Picture->GetItemRect().Height(), Picture->m_nCol1, Picture->m_nRow1);
		if (hBitmap)
		{			
			Picture->SetHBITMAP(hBitmap);
			Picture->SetFixedHeight(false);
			
		}
	}
	if(hBitmap) ::DeleteObject(hBitmap);

	rs.GetValue(_T("hpcl_practitioner"), tmpStr);


	rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(tmpStr));
	CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
	if(img)
	{
		img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img->SetFixedHeight(false);		
	}	
	

	rs.GetValue(_T("hpc_performdate"), tmpStr);
	tmpStr2 = rpt.GetReportFooter()->GetValue(_T("PerformDate"));
	if(CDateTime::IsValid(tmpStr))
		szDate.Format(tmpStr2,tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	else
		szDate.Format(tmpStr2, _T("....."), _T("....."), _T("........"));
	rpt.GetReportFooter()->SetValue(_T("PerformDate"), szDate);
	rs.GetValue(_T("hpcl_practitioner"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("Practitioner"), GetDoctorName(tmpStr));
	CReportItem *img2 = rpt.GetReportFooter()->GetItem(_T("Signature2"));
	if(img2)
	{
		img2->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img2->SetFixedHeight(false);
	}
	
	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print(bDirect);
}





void CPrintForms::PrintTestOrder(long nOrderID, bool bPreview, bool bDirect){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szGroupID, szSex, szStatus;
	CString szReportName, szOrderName;
	CString szSQL;
	CRecord rs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);

	long nDocumentNo;
	int nNumberOrder=0;
	CString tmpStr2, szDeptType;
	CString szFormID;


	szSQL.Format(_T(" SELECT hpc_deptid as deptid, hpc_docno AS docno,") \
_T("   (SELECT DISTINCT hfg_name2 FROM hms_fee_group WHERE hfg_id=hpc_groupid") \
_T("   ) AS group_name,") \
_T("   trim(hp_surname") \
_T("   ||' '") \
_T("   ||hp_midname") \
_T("   ||' '") \
_T("   ||hp_firstname) AS pname,") \
_T("   date_part('year', hp_birthdate) as yob, ") \
_T("   HMS_GETSEX(hp_sex)           AS gender,") \
_T("   HMS_GETOBJECTNAME(hd_object) AS object_desc,") \
_T("   hd_cardno as cardno, hd_diagnostic, ") \
_T(" hpc_orderdate as orderdate , ") \
_T(" get_username(hpc_doctor) as doctorname,  ") \
_T(" hpc_roomid as roomid, hpc_class ") \
_T(" FROM hms_patient,") \
_T("   hms_doc,") \
_T("   hms_testorder") \
_T(" WHERE hpc_orderid = %ld ") \
_T(" AND hpc_docno     = hd_docno") \
_T(" AND hd_patientno  = hp_patientno"), nOrderID);

	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
	{
		return;
	}


	if(!rpt.Init(_T("REPORTS/HMS/HMS_TESTORDER.RPT")) )
	{
		return;
	}

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HealthService"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), tmpStr);
	
	CString szDeptID;

	rs.GetValue(_T("deptid"), szDeptID);
	rpt.GetReportHeader()->SetValue(_T("Department"), pMF->GetDepartmentName(szDeptID));

	
	StringUpper(rs.GetValue(_T("group_name")), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("TITLE"), tmpStr);
	
	rs.GetValue(_T("docno"), nDocumentNo);
	tmpStr.Format(_T("%ld"), nDocumentNo);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128A]"),tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128B]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128C]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[39]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[93]"), tmpStr);


	tmpStr.Format(_T("%ld"), nOrderID);
	rpt.GetReportHeader()->SetValue(_T("OrderNo"), tmpStr);

	rs.GetValue(_T("pname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("PatientName"), tmpStr);
	//Chi In ra nam sinh.
	rs.GetValue(_T("yob"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);

	rs.GetValue(_T("Gender"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Gender"), tmpStr);
	rs.GetValue(_T("object_desc"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Object"), tmpStr);
	rs.GetValue(_T("cardno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CardNo"), tmpStr);

	rs.GetValue(_T("roomid"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RoomID"), tmpStr);

	CString szDate;
	rs.GetValue(_T("orderdate"), tmpStr);
	tmpStr2 = rpt.GetReportFooter()->GetValue(_T("PrintDate"));
	szDate.Format(tmpStr2, tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);

	rs.GetValue(_T("doctorname"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("DoctorName"), tmpStr);

	rs.GetValue(_T("hpc_class"), tmpStr);
	if(tmpStr == _T("E"))
	{
		rs.GetValue(_T("hd_diagnostic"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);
	}
	else
	{
		CRecord rsx(&pMF->m_db);
		CString szSQL;
		szSQL.Format(_T("SELECT htr_maindisease FROM hms_treatment_record ") \
			_T("WHERE htr_docno =%ld and htr_deptid='%s' and rownum <=1 ORDER BY htr_idx DESC "), 
			nDocumentNo, szDeptID);
		rsx.ExecSQL(szSQL);
		rsx.GetValue(_T("htr_maindisease"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);
	}


	
	szSQL.Format(_T(" SELECT hfl_name AS testname,") \
_T("   hfl_unit      AS unit, hpcl_qty, hpcl_note ") \
_T(" FROM hms_testorderline") \
_T(" LEFT JOIN hms_fee_list") \
_T(" ON(hfl_feeid       = hpcl_itemid)") \
_T(" WHERE hpcl_orderid = %ld") \
_T(" AND hpcl_hasfee    ='Y'") \
_T(" ORDER BY hfl_line "), nOrderID);


	rsl.ExecSQL(szSQL);
	
	CReportSection* rptDetail = NULL;
	int nItem = 1;
	int nQty;
	CString szNote, szUnit;
	
	while(!rsl.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nItem++);
		rptDetail->SetValue(_T("IDX"), tmpStr);
		rsl.GetValue(_T("testname"), tmpStr);
		rsl.GetValue(_T("hpcl_note"), szNote);
		rsl.GetValue(_T("hpcl_qty"), nQty);
		szNote.Trim();
		if(!szNote.IsEmpty())
		{
			tmpStr.AppendFormat(_T(" (%s)"), szNote);
		}
		if(nQty > 1)
		{
			rsl.GetValue(_T("unit"),szUnit);
			tmpStr.AppendFormat(_T("; S\x1ED1 l\x1B0\x1EE3ng: %d %s"), nQty, szUnit);
		}

		rptDetail->SetValue(_T("TestName"), tmpStr);
		rsl.MoveNext();
	}

	//Page Footer
	//Report Footer
	

	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print(bDirect);
}

void CPrintForms::PrintTestResult(long nOrderID, bool bPreview, bool bDirect)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	CString szModuleID = pMF->GetModuleID();
	pMF->m_szModuleID = _T("EM");
	pMF->PrintTestOrderX(nOrderID, true);
	pMF->m_szModuleID = szModuleID;
	return;

	CReport rpt;
	CString tmpStr, szGroupID, szSex, szStatus;
	CString szReportName, szOrderName;
	CString szSQL;
	CRecord rs(&pMF->m_db);
	CRecord rs2(&pMF->m_db);
	CRecord rs3(&pMF->m_db);
	long nDocumentNo;
	int nRefIndex, nRoomID;
	int nNumberOrder=0;
	CString tmpStr2, szDeptType;
	CString szFormID;


	szSQL.Format(_T("SELECT * FROM hms_testorder WHERE hpc_orderid=%ld "), nOrderID);

	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;

	rs.GetValue(_T("hpc_groupid"), szGroupID);
	rs.GetValue(_T("hpc_docno"), nDocumentNo);
	rs.GetValue(_T("hpc_refidx"), nRefIndex);
	rs.GetValue(_T("hpc_status"), szStatus);
	rs.GetValue(_T("hpc_depttype"), szDeptType);

	szSQL.Format(_T("SELECT hfg_report as reportname FROM hms_fee_group WHERE hfg_id='%s' "), szGroupID);
	rs2.ExecSQL(szSQL);
	tmpStr.Empty();
	if(!rs2.IsEOF())
		rs2.GetValue(_T("reportname"), tmpStr);

	if(tmpStr.IsEmpty())
		tmpStr = _T("PACS_VOTE.RPT");
	tmpStr.Trim();

	szReportName.Format(_T("Reports/HMS/%s"), tmpStr);
	


	szSQL.Format(_T(" SELECT hfl_name as name, hfl_unit as unit") \
		_T(" FROM hms_testorderline") \
		_T(" LEFT JOIN hms_fee_list ON(hpcl_itemid=hfl_feeid)") \
		_T(" WHERE 	hpcl_orderid=%ld ") \
		_T(" 	and hpcl_result='%s'"), nOrderID, szFormID);
	rs2.ExecSQL(szSQL);
	
	while(!rs2.IsEOF())
	{
		rs2.GetValue(_T("name"), tmpStr2);
		if(!szOrderName.IsEmpty())
			szOrderName += _T("\r\n");
		szOrderName.AppendFormat(_T("%s"), tmpStr2);
		rs2.MoveNext();
		nNumberOrder++;
	}
	
	szReportName.Trim();

	if(!rpt.Init(szReportName) )
	{
		return;
	}
	 
	
		
	szSQL.Format(_T(" SELECT 	hd_patientno as patientno,  ") \
		_T(" 	hd_docno as docno,") \
		_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
		_T(" 	CAST(TO_CHAR(hp_birthdate, 'YYYY') as integer) as birthyear,") \
		_T(" 	hms_getage(trunc_date(hd_admitdate), hp_birthdate) as age, 	hp_sex as sexid,") \
		_T(" 	sys_sel_getname('sys_sex', hp_sex) as sex,") \
		_T(" 	hp_occupation as occupationid,") \
		_T(" 	sys_sel_getname('sys_occupation', cast(hp_occupation as varchar(15))) as occupation,") \
		_T(" 	hp_dtladdr as detailaddress,") \
		_T(" 	hms_getaddress(hp_provid,hp_distid,hp_villid) as address,") \
		_T(" 	hp_workplaceid as workplaceid,") \
		_T(" 	hp_workplace as workplace,") \
		_T(" 	hd_object as objectid,") \
		_T(" 	(SELECT ho_desc FROM hms_object WHERE ho_id=hd_object) as objectname,") \
		_T(" 	hd_cardno as cardno,") \
		_T(" 	hd_cardidx as cardidx, ") \
		_T(" 	hd_reldisease as relationdisease, ") \
		_T(" 	hc_regdate as regdate, ") \
		_T(" 	hc_expdate as expdate ") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_card ON(hc_patientno=hd_patientno AND hc_idx=hd_cardidx) ") \
		_T(" WHERE hd_docno=%ld"), nDocumentNo);
//_fmsg(_T("%s"), szSQL);
	int ret = rs2.ExecSQL(szSQL);
	if(rs2.IsEOF())
	{
		ShowMessageBox(_T("Patient not found."));
		return;
	}

	//Report Header
	tmpStr.Empty();
	rs2.GetValue(_T("sexid"), szSex);
	//StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HealthService"), tmpStr);
	//StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), tmpStr);
	rs.GetValue(_T("hpc_pdeptid"), tmpStr);
	if(!tmpStr.IsEmpty())
	{
		CString szPerformDept;
		StringUpper(pMF->GetDepartmentName(tmpStr), szPerformDept);
		rpt.GetReportHeader()->SetValue(_T("PerformDept"), szPerformDept);
	}
	else
		rpt.GetReportHeader()->SetValue(_T("PerformDept"), _T(""));

	rs.GetValue(_T("hpc_docno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	tmpStr.Format(_T("%ld"), nOrderID);
	rpt.GetReportHeader()->SetValue(_T("SheetNo"), tmpStr);
	rs2.GetValue(_T("pname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("PatientName"), tmpStr);
	rs2.GetValue(_T("age"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), tmpStr);

	//Chi In ra nam sinh.
	rs2.GetValue(_T("birthyear"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);

	rs2.GetValue(_T("sex"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Sex"), tmpStr);
	tmpStr.Format(_T("%s - %s"), rs2.GetValue(_T("detailaddress")), rs2.GetValue(_T("address")));
	rpt.GetReportHeader()->SetValue(_T("Address"), tmpStr);

	rs2.GetValue(_T("objectname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("objectname"), tmpStr);

	rs2.GetValue(_T("workplace"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Workplace"), tmpStr);

	rs2.GetValue(_T("cardno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CardNo"), tmpStr);
	rs2.GetValue(_T("regdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RegDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	rs2.GetValue(_T("expdate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ExpiryDate"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));
	

	CString szDate;
	rs.GetValue(_T("hpc_orderdate"), tmpStr);
	tmpStr2 = rpt.GetReportHeader()->GetValue(_T("OrderDate"));
	szDate.Format(tmpStr2, tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportHeader()->SetValue(_T("OrderDate"), szDate);

	rs.GetValue(_T("hpc_doctor"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DoctorName"), GetDoctorName(tmpStr));
	CReportItem *imgHeader = rpt.GetReportHeader()->GetItem(_T("Signature"));
	if(imgHeader)
	{
		imgHeader->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		imgHeader->SetFixedHeight(false);		
	}	
	


	CString szDept;
	CRecord rsd(&pMF->m_db);
	rs.GetValue(_T("hpc_roomid"), tmpStr);
	nRoomID = ToInt(tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Room"), tmpStr);

	rs.GetValue(_T("hpc_deptid"), szDept);
	szSQL.Format(_T("SELECT hrl_name as roomname FROM hms_roomlist WHERE hrl_deptid='%s' AND hrl_id=%d"), szDept, ToInt(tmpStr));
	rsd.ExecSQL(szSQL);
	if(!rsd.IsEOF())
	{
		rsd.GetValue(_T("roomname"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("RoomName"), tmpStr);
	}

	
	rs.GetValue(_T("hpc_deptid"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Department"), pMF->GetDepartmentName(tmpStr));
	
	rs.GetValue(_T("hpc_bedid"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Bed"), tmpStr);

	szSQL.Format(_T("SELECT hbl_name as bedname FROM hms_Bedlist WHERE hbl_deptid='%s' AND hbl_roomid =%d AND hbl_id=%d"), szDept, nRoomID,ToInt(tmpStr));
	rsd.ExecSQL(szSQL);
	if(!rsd.IsEOF())
	{
		rsd.GetValue(_T("Bedname"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("BedName"), tmpStr);
	}

	szSQL.Format(_T("SELECT htr_maindisease as maindisease, hcr_reldisease as reldisease FROM hms_clinical_record LEFT JOIN hms_treatment_record ON(htr_docno=hcr_docno) WHERE htr_docno=%ld and htr_idx=%d"), nDocumentNo, nRefIndex);
	rs3.ExecSQL(szSQL);
	if(!rs3.IsEOF())
	{
		rs3.GetValue(_T("maindisease"), tmpStr);

		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);
		rs3.GetValue(_T("reldisease"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("relationdisease"), tmpStr);
	}
	else
	{
		tmpStr.Empty();
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("relationdisease"), tmpStr);
	}

	
	szSQL.Format(_T(" SELECT 	hfl_name as name, ") \
				_T(" 	hfl_unit as unit, ") \
				_T(" 	hfl_index1 as index1, ") \
				_T(" 	hfl_index2 as index2,") \
				_T(" 	hpcl_result as result,") \
				_T(" 	hpcl_note as note,hfl_subitem as subitem") \
				_T(" FROM hms_testorderline ") \
				_T(" LEFT JOIN hms_fee_list ON(hfl_feeid=hpcl_itemid)") \
				_T(" WHERE hpcl_orderid=%ld  and (hfl_printsubitem is NULL or hfl_printsubitem <>'N') ") \
				_T(" ORDER BY hfl_line"), nOrderID);


			ret = rs2.ExecSQL(szSQL);
			CReportSection* rptDetail;
			if(szGroupID.Left(2) == _T("B1")){
		/*
				rs2.MoveFirst();
				if(ret > 30)
				{
					for (int i =0; i <= ret/2; i++)
					{
						rptDetail = rpt.AddDetail();
						rs2.GetValue(_T("name"), tmpStr);
						rptDetail->SetValue(_T("1"), tmpStr);
						rs2.GetValue(_T("index1"), tmpStr);
						rptDetail->SetValue(_T("2"), tmpStr);
						rs2.GetValue(_T("result"), tmpStr);
						rptDetail->SetValue(_T("3"), tmpStr);
						rptDetail->SetValue(_T("4"), _T(""));
						rptDetail->SetValue(_T("5"), _T(""));
						rptDetail->SetValue(_T("6"), _T(""));
						rs2.MoveNext();
					}
					for (int i =0; i <= ret/2 ; i++)
					{
						if(rs2.IsEOF())
							break;
						rptDetail = rpt.GetDetail(i+1);
						rs2.GetValue(_T("name"), tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);
						rs2.GetValue(_T("index1"), tmpStr);
						rptDetail->SetValue(_T("5"), tmpStr);
						rs2.GetValue(_T("result"), tmpStr);
						rptDetail->SetValue(_T("6"), tmpStr);
						rs2.MoveNext();
					}
				}
				else
				{
		*/
					int nItem = 1;
					while(!rs2.IsEOF())
					{							
						rs2.GetValue(_T("subitem"), tmpStr);
						if (tmpStr.GetLength() <= 1)
						{	
							//rpt.GetDetail(0)->GetItem(_T("1"))->SetFaceSize(10);
							//rpt.GetDetail(0)->GetItem(_T("1"))->SetItalic(false);
							tmpStr.Format(_T(" %d. %s"), nItem++, rs2.GetValue(_T("name")));
						}
						else
						{	
							rpt.GetDetail(0)->GetItem(_T("1"))->SetFaceSize(8);
							rpt.GetDetail(0)->GetItem(_T("1"))->SetItalic(true);
							tmpStr.Format(_T("     - %s"), rs2.GetValue(_T("name")));
						}						

						rptDetail = rpt.AddDetail();
						rptDetail->SetValue(_T("1"), tmpStr);		
						if(szSex == _T("F"))
							rs2.GetValue(_T("index2"), tmpStr);
						else
							rs2.GetValue(_T("index1"), tmpStr);
						if(tmpStr == _T("0") || tmpStr == _T("0 - 0") || tmpStr == _T("0-0"))
							tmpStr.Empty();
						rptDetail->SetValue(_T("2"), tmpStr);
						
						tmpStr.Empty();
						rs2.GetValue(_T("result"), tmpStr);

						rptDetail->SetValue(_T("3"), tmpStr);
						rs2.GetValue(_T("unit"), tmpStr);
						rptDetail->SetValue(_T("4"), tmpStr);
						rs2.MoveNext();
					}
			}
			else
			{
					int nItem = 1;
					while(!rs2.IsEOF())
					{
						rptDetail = rpt.AddDetail();
						tmpStr.Format(_T("%d. %s"), nItem++, rs2.GetValue(_T("name")));
						rptDetail->SetValue(_T("1"), tmpStr);
						rs2.GetValue(_T("unit"), tmpStr);
						rptDetail->SetValue(_T("2"), tmpStr);
						rs2.MoveNext();
					}

			}
	

	//Page Footer
	//Report Footer
	
	rs.GetValue(_T("hpc_orderdate"), tmpStr);
	tmpStr2 = rpt.GetReportFooter()->GetValue(_T("OrderDate"));
	szDate.Format(tmpStr2, tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("OrderDate"), szDate);
	
	CReportItem *Picture = NULL;
	
	Picture = rpt.GetReportFooter()->GetItem(_T("Picture"));
	if(Picture)
	{	
		HBITMAP hBitmap = pMF->GetPACSImage(pMF->m_szCurrentFileNamePACS);
		if (hBitmap)
		{
			Picture->SetHBITMAP(hBitmap);
			Picture->SetFixedHeight(false);
			::DeleteObject(hBitmap);
		}
	}
	Picture = rpt.GetReportHeader()->GetItem(_T("Picture"));
	
	if(Picture)
	{	
		HBITMAP hBitmap = pMF->GetPACSImage(pMF->m_szCurrentFileNamePACS);
		if (hBitmap)
		{
			Picture->SetHBITMAP(hBitmap);
			Picture->SetFixedHeight(false);
			::DeleteObject(hBitmap);
		}
	}

	rs.GetValue(_T("hpc_doctor"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("DoctorName"), GetDoctorName(tmpStr));
	CReportItem *img = rpt.GetReportFooter()->GetItem(_T("Signature"));
	if(img)
	{
		img->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img->SetFixedHeight(false);		
	}	
	

	rs.GetValue(_T("hpc_performdate"), tmpStr);
	tmpStr2 = rpt.GetReportFooter()->GetValue(_T("PerformDate"));
	if(CDateTime::IsValid(tmpStr))
		szDate.Format(tmpStr2,tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	else
		szDate.Format(tmpStr2, _T("....."), _T("....."), _T("........"));
	rpt.GetReportFooter()->SetValue(_T("PerformDate"), szDate);
	rs.GetValue(_T("hpc_practitioner"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("Practitioner"), GetDoctorName(tmpStr));
	CReportItem *img2 = rpt.GetReportFooter()->GetItem(_T("Signature2"));
	if(img2)
	{
		img2->SetHBITMAP(pMF->GetSignatureImage(tmpStr));
		img2->SetFixedHeight(false);
	}


	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print(bDirect);
}

void CPrintForms::PrintPACSOrder(long nOrderID,  CString szItemID, bool bPreview, bool bDirect){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CString tmpStr, szGroupID, szSex, szStatus;
	CString szReportName, szOrderName;
	CString szSQL;
	CRecord rs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);

	long nDocumentNo;
	int nNumberOrder=0;
	CString tmpStr2, szDeptType;
	CString szFormID;

	szSQL.Format(_T(" SELECT hpc_deptid as deptid, hpc_docno AS docno,") \
_T("   (SELECT DISTINCT hfg_name2 FROM hms_fee_group WHERE hfg_id=hpc_groupid") \
_T("   ) AS group_name,") \
_T("   trim(hp_surname") \
_T("   ||' '") \
_T("   ||hp_midname") \
_T("   ||' '") \
_T("   ||hp_firstname) AS pname,") \
_T("   date_part('year', hp_birthdate) as yob, ") \
_T("   HMS_GETSEX(hp_sex)           AS gender,") \
_T("   HMS_GETOBJECTNAME(hd_object) AS object_desc,") \
_T("   hd_cardno as cardno, hd_diagnostic, ") \
_T(" hpc_orderdate as orderdate , ") \
_T(" get_username(hpc_doctor) as doctorname, ") \
_T(" hpc_roomid as roomid, hpc_class, hd_rank ") \
_T(" FROM hms_patient,") \
_T("   hms_doc,") \
_T("   hms_pacsorder") \
_T(" WHERE hpc_orderid = %ld ") \
_T(" AND hpc_docno     = hd_docno") \
_T(" AND hd_patientno  = hp_patientno"), nOrderID);

	rs.ExecSQL(szSQL);
	if(rs.IsEOF())
	{
		return;
	}


	if(!rpt.Init(_T("REPORTS/HMS/HMS_TESTORDER.RPT")) )
	{
		return;
	}

	StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HealthService"), tmpStr);
	StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), tmpStr);
	

	CString szDeptID;
	rs.GetValue(_T("deptid"), szDeptID);
	rpt.GetReportHeader()->SetValue(_T("Department"), pMF->GetDepartmentName(szDeptID));


	StringUpper(rs.GetValue(_T("group_name")), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("TITLE"), tmpStr);

	rs.GetValue(_T("docno"), nDocumentNo);
	tmpStr.Format(_T("%ld"), nDocumentNo);
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128A]"),tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128B]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[128C]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[39]"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Barcode[93]"), tmpStr);


	tmpStr.Format(_T("%ld"), nOrderID);
	rpt.GetReportHeader()->SetValue(_T("OrderNo"), tmpStr);
	rs.GetValue(_T("pname"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("PatientName"), tmpStr);
	//Chi In ra nam sinh.
	rs.GetValue(_T("yob"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("YearOfBirth"), tmpStr);

	rs.GetValue(_T("Gender"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Gender"), tmpStr);
	rs.GetValue(_T("object_desc"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Object"), tmpStr);
	rs.GetValue(_T("cardno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CardNo"), tmpStr);

	int nRoomID;
	rs.GetValue(_T("roomid"), tmpStr);
	nRoomID = str2int(tmpStr);
	rpt.GetReportHeader()->SetValue(_T("RoomID"), tmpStr);

	rs.GetValue(_T("hd_rank"), tmpStr);
	if(!tmpStr.IsEmpty())
	{
		rpt.GetReportHeader()->SetValue(_T("Rank"), pMF->GetSelectionString(_T("hms_rank"), tmpStr));
	}


	CString szDate;
	rs.GetValue(_T("orderdate"), tmpStr);
	tmpStr2 = rpt.GetReportFooter()->GetValue(_T("PrintDate"));
	szDate.Format(tmpStr2, tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);

	rs.GetValue(_T("doctorname"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("DoctorName"), tmpStr);

	rs.GetValue(_T("hpc_class"), tmpStr);
	if(tmpStr == _T("E"))
	{
		rs.GetValue(_T("hd_diagnostic"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);
		szSQL.Format(_T("SELECT max(he_receptno) as receptno FROM hms_exam WHERe he_docno=%ld and he_roomid=%d"), nDocumentNo, nRoomID);
		CRecord rsx(&pMF->m_db);
		rsx.ExecSQL(szSQL);
		if(!rsx.IsEOF())
		{
			rsx.GetValue(_T("receptno"), tmpStr);
			rpt.GetReportHeader()->SetValue(_T("ReceptNo"), tmpStr);
		}

	}
	else
	{
		CRecord rsx(&pMF->m_db);
		szSQL.Format(_T("SELECT htr_maindisease FROM hms_treatment_record ") \
			_T("WHERE htr_docno =%ld and htr_deptid='%s' and rownum <=1 ORDER BY htr_idx DESC "), 
			nDocumentNo, szDeptID);
		rsx.ExecSQL(szSQL);
		rsx.GetValue(_T("htr_maindisease"), tmpStr);
		rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);
	}

	
	szSQL.Format(_T(" SELECT hfl_name AS testname, hpcl_note, ") \
_T("   hfl_unit      AS unit, hpcl_qty as qty ") \
_T(" FROM hms_pacsorderline ") \
_T(" LEFT JOIN hms_fee_list") \
_T(" ON(hfl_feeid       = hpcl_itemid)") \
_T(" WHERE hpcl_orderid = %ld") \
_T(" AND hpcl_hasfee    ='Y'") \
_T(" ORDER BY hfl_line "), nOrderID);


	rsl.ExecSQL(szSQL);
	
	CReportSection* rptDetail = NULL;
	int nItem = 1;
	int nQty;
	CString szNote, szUnit;
	while(!rsl.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nItem++);
		rptDetail->SetValue(_T("IDX"), tmpStr);
		rsl.GetValue(_T("qty"), nQty);
		rsl.GetValue(_T("testname"), tmpStr);
		rsl.GetValue(_T("hpcl_note"), szNote);
		szNote.Trim();
		if(!szNote.IsEmpty())
		{
			tmpStr.AppendFormat(_T(" (%s)"), szNote);
		}
		if(nQty > 1)
		{
			rsl.GetValue(_T("hfl_unit"), szUnit);
			tmpStr.AppendFormat(_T("; S\x1ED1 l\x1B0\x1EE3ng: %d %s"), nQty, szUnit);
		}

		rptDetail->SetValue(_T("TestName"), tmpStr);
		rsl.MoveNext();
	}

	//Page Footer
	//Report Footer
	

	if(bPreview)
		rpt.PrintPreview();
	else
		rpt.Print(bDirect);
}

void CPrintForms::PrintPACSResult(long nOrderID,  CString szItemID, bool bPreview, bool bDirect){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	pMF->PrintPACSOrderX(nOrderID, szItemID, true, false);
	
}

void CPrintForms::PM_PrintPurchaseInvoice(long nOrderID)
{
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL;
	double nTmp = 0;
	int nIdx = 1;
	CString szSQL, tmpStr, szReportName, szDate, szOldLev1, szNewLev1, szFromDate, szToDate, szField, szTemp;
	long double nTotal = 0, nGroupTotal1 = 0;
	if (!rpt.Init(_T("Reports/HMS/PM_BANGKETHANHTOANCONGNO.RPT")))
		return;
	szSQL.Format(_T(" SELECT min(po_approveddate) from_date, max(po_approveddate) to_date") \
				_T(" FROM   purchase_order") \
				_T(" WHERE  po_payment = 'Y' AND po_refpayment_id = %ld"), nOrderID);
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{
		rs.GetValue(_T("from_date"), szFromDate);
		rs.GetValue(_T("to_date"), szToDate);
	}
	szSQL.Format(_T(" SELECT") \
				_T("   po_purchase_order_id as id,") \
				_T("   po_invoiceno                   AS invoiceno,") \
				_T("   Cast_Timestamp2Date(po_invoicedate)                 AS invoicedate,") \
				_T("   po_orderno                     AS orderno,") \
				_T("   Cast_Timestamp2Date(po_approveddate)                AS importdate,") \
				_T("   po_totalamount                 AS amount,") \
				_T("   Get_partnername(po_partner_id) AS partnername") \
				_T(" FROM   purchase_order") \
				_T(" LEFT JOIN purchase_payment ON (po_refpayment_id=pp_purchase_payment_id)") \
				_T(" WHERE  po_payment = 'Y' AND po_refpayment_id = %ld") \
				_T(" ORDER  BY po_partner_id "), nOrderID);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(szFromDate, yyyymmdd, ddmmyyyy), 
		CDate::Convert(szToDate), yyyymmdd, ddmmyyyy);
		rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	}
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("invoiceno")));
			rs.GetValue(_T("invoicedate"), tmpStr);
			rptDetail->SetValue(_T("3"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));		
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("orderno")));
			rs.GetValue(_T("importdate"), tmpStr);
			rptDetail->SetValue(_T("5"), CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy));		
			rs.GetValue(_T("amount"), nTmp);
			nGroupTotal1+= nTmp;
			rptDetail->SetValue(_T("6"), double2str(nTmp));
		}
		rs.MoveNext();
	}
	if (nGroupTotal1 > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		rptDetail->SetValue(_T("s6"), double2str(nGroupTotal1));
		nTotal += nGroupTotal1;
	}
	rptDetail = rpt.GetReportFooter();
	rptDetail->SetValue(_T("TotalItem"), int2str(nIdx++));
	tmpStr.Format(_T("%f"), nTotal);
	rptDetail->SetValue(_T("TotalAmount"), tmpStr);
	szSQL.Format(_T(" SELECT SUM(thuoc) drug, ") \
				_T("        SUM(hoachat) chemicals, ") \
				_T("        SUM(bbg) cotton, ") \
				_T("        SUM(duoclieu) herb,") \
				_T("        SUM(tpdd) compound,") \
				_T("        SUM(ycu) instrument") \
				_T(" FROM   (SELECT CASE WHEN (instr(product_producttype, 'A11') > 0 OR product_producttype IN ('A1400', 'A6000'))") \
				_T("                     THEN pol_totalamount ") \
				_T("                ELSE 0 ") \
				_T("                END AS thuoc, ") \
				_T("                CASE WHEN instr(product_producttype, 'A121') > 0 THEN pol_totalamount ") \
				_T("                ELSE 0 ") \
				_T("                END AS hoachat, ") \
				_T("                CASE WHEN product_producttype = 'A1500' THEN pol_totalamount ") \
				_T("                ELSE 0 ") \
				_T("                END AS bbg, ") \
				_T("                CASE WHEN product_producttype = 'A1302' THEN pol_totalamount ") \
				_T("                ELSE 0 ") \
				_T("                END AS duoclieu, ") \
				_T("				CASE WHEN product_producttype = 'A1301' THEN pol_totalamount") \
				_T("				ELSE 0") \
				_T("				END AS tpdd,") \
				_T("				CASE WHEN product_producttype = 'A1600' THEN pol_totalamount") \
				_T("				ELSE 0") \
				_T("				END AS ycu") \
				_T("         FROM   purchase_orderline, ") \
				_T("                m_product_view, ") \
				_T("				purchase_order,") \
				_T("				purchase_payment") \
				_T("         WHERE  pol_product_id = product_id ") \
				_T("			AND po_purchase_order_id = pol_purchase_order_id") \
				_T("			AND po_refpayment_id=pp_purchase_payment_id AND po_payment = 'Y' AND po_refpayment_id = %ld) "), nOrderID);
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{
		//l1..l4
		int nIdx = 1;
		for (int i = 0; i < rs.GetFieldCount(); i++)
		{
			szField = rs.GetFieldName(i);
			rs.GetValue(szField, nTmp);
			if (nTmp > 0)
			{
				tmpStr.Format(_T("l%d"), nIdx);
				TranslateString(szField, szTemp);
				rptDetail->SetValue(tmpStr, szTemp);
				tmpStr.Format(_T("s%d"), nIdx);
				FormatCurrency(nTmp, szTemp);
				szTemp += _T(" \x111\x1ED3ng.");
				rptDetail->SetValue(tmpStr, szTemp);
				nIdx++;
			}
		}
	}
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
}

void CPrintForms::E10_PrintAdjustmentExport(long nOrderId){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	/*Declaration Section*/
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL, * rptFooter = NULL;
	double nTmp = 0;
	int nIdx = 1;
	CString szSQL, tmpStr, szReportName, szDate, szTemp;
	long double nTotal = 0;
	szReportName = _T("Reports\\HMS\\PMS_EXPORTORDER.RPT");
	if (!rpt.Init(szReportName))
		return;

	szSQL.Format(_T(" SELECT mi_orderno                     invoice_no, ") \
	_T("        mi_date                        inventory_date, ") \
	_T("        mi_status                      status, ") \
	_T("        Get_storagename(mi_storage_id) storage_name, ") \
	_T("        mi_description                 note ") \
	_T(" FROM   m_inventory") \
	_T(" WHERE mi_inventory_id = %ld ") \
	_T(" AND mi_status = 'A'"), nOrderId);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		szDate = pMF->GetSysDate();
		tmpStr.Format(rptHeader->GetValue(_T("PrintDate")), szDate.Mid(8, 2), szDate.Mid(5, 2), szDate.Left(4));
		rptHeader->SetValue(_T("PrintDate"), tmpStr);
		rptHeader->SetValue(_T("InvoiceNo"), rs.GetValue(_T("invoice_no")));
		rptHeader->SetValue(_T("OrderDate"), rs.GetValue(_T("inventory_date")));																		
		rptHeader->SetValue(_T("Status"), _T("\x110\xE3 \x64uy\x1EC7t"));
		rptHeader->SetValue(_T("StockName"), rs.GetValue(_T("storage_name")));
		rptHeader->SetValue(_T("Comment"), rs.GetValue(_T("note")));
		rptHeader->SetValue(_T("Reason"), _T("\x110i\x1EC1u \x63h\x1EC9nh"));
	}
	szSQL.Format(_T(" SELECT    product_name, ") \
	_T("           product_countryname, ") \
	_T("           product_uomname, ") \
	_T("           product_lotno, ") \
	_T("           mil_qtyadj, ") \
	_T("           product_taxprice, ") \
	_T("           mil_qtyadj * product_taxprice amount ") \
	_T(" FROM      m_inventoryline ") \
	_T(" LEFT JOIN m_productitem_view ON ( mil_product_item_id = product_item_id ) ") \
	_T(" WHERE mil_inventory_id = %ld ") \
	_T(" AND mil_qtyadj > 0"), nOrderId);
	
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	while (!rs.IsEOF())
	{
		/*Group Seperation*/
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("product_name")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("product_countryname")));
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("product_uomname")));
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("product_lotno")));
			rptDetail->SetValue(_T("6"), rs.GetValue(_T("mil_qtyadj")));
			rptDetail->SetValue(_T("7"), rs.GetValue(_T("product_taxprice")));
			rs.GetValue(_T("amount"), nTmp);
			nTotal += nTmp;
			rptDetail->SetValue(_T("8"), double2str(nTmp));
		}
		rs.MoveNext();
	}
	nIdx--;
	rptFooter = rpt.GetReportFooter();
	if (rptFooter)
	{
		rptFooter->SetValue(_T("TotalItem"), int2str(nIdx));
		tmpStr.Format(_T("%f"), nTotal);
		rptFooter->SetValue(_T("TotalAmount"), tmpStr);
		MoneyToString(tmpStr, szTemp);
		rptFooter->SetValue(_T("SumInWord"), szTemp);
	}
	rpt.PrintPreview();
}

void CPrintForms::E10_PrintAdjustmentImport(long nOrderId){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	/*Declaration Section*/
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL, * rptFooter = NULL;
	double nTmp = 0;
	int nIdx = 1;
	CString szSQL, tmpStr, szReportName, szDate, szTemp;
	long double nTotal = 0;
	szReportName = _T("Reports\\HMS\\PMS_IMPORTORDER.RPT");
	if (!rpt.Init(szReportName))
		return;

	szSQL.Format(_T(" SELECT mi_orderno                     invoice_no, ") \
	_T("        mi_date                        inventory_date, ") \
	_T("        mi_status                      status, ") \
	_T("        Get_storagename(mi_storage_id) storage_name, ") \
	_T("        mi_description                 note ") \
	_T(" FROM   m_inventory") \
	_T(" WHERE mi_inventory_id = %ld ") \
	_T(" AND mi_status = 'A'"), nOrderId);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		szDate = pMF->GetSysDate();
		tmpStr.Format(rptHeader->GetValue(_T("PrintDate")), szDate.Mid(8, 2), szDate.Mid(5, 2), szDate.Left(4));
		rptHeader->SetValue(_T("PrintDate"), tmpStr);
		rptHeader->SetValue(_T("InvoiceNo"), rs.GetValue(_T("invoice_no")));
		rptHeader->SetValue(_T("OrderDate"), rs.GetValue(_T("inventory_date")));																		
		rptHeader->SetValue(_T("Status"), _T("\x110\xE3 \x64uy\x1EC7t"));
		rptHeader->SetValue(_T("StockName"), rs.GetValue(_T("storage_name")));
		rptHeader->SetValue(_T("Comment"), rs.GetValue(_T("note")));
		rptHeader->SetValue(_T("Reason"), _T("\x110i\x1EC1u \x63h\x1EC9nh"));
	}
	szSQL.Format(_T(" SELECT    product_name, ") \
	_T("           product_countryname, ") \
	_T("           product_uomname, ") \
	_T("           product_lotno, ") \
	_T("           mil_qtyadj, ") \
	_T("           product_taxprice, ") \
	_T("           mil_qtyadj * product_taxprice amount ") \
	_T(" FROM      m_inventoryline ") \
	_T(" LEFT JOIN m_productitem_view ON ( mil_product_item_id = product_item_id ) ") \
	_T(" WHERE mil_inventory_id = %ld ") \
	_T(" AND mil_qtyadj < 0"), nOrderId);
	
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	while (!rs.IsEOF())
	{
		/*Group Seperation*/
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("product_name")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("product_countryname")));
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("product_uomname")));
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("product_lotno")));
			rptDetail->SetValue(_T("6"), rs.GetValue(_T("mil_qtyadj")));
			rptDetail->SetValue(_T("7"), rs.GetValue(_T("product_taxprice")));
			rs.GetValue(_T("amount"), nTmp);
			nTotal += nTmp;
			rptDetail->SetValue(_T("8"), double2str(nTmp));
		}
		rs.MoveNext();
	}
	nIdx--;
	rptFooter = rpt.GetReportFooter();
	if (rptFooter)
	{
		rptFooter->SetValue(_T("TotalItem"), int2str(nIdx));
		tmpStr.Format(_T("%f"), nTotal);
		rptFooter->SetValue(_T("TotalAmount"), tmpStr);
		MoneyToString(tmpStr, szTemp);
		rptFooter->SetValue(_T("SumInWord"), szTemp);
	}
	rpt.PrintPreview();
}

void CPrintForms::HMS_InwardDrugUsage(long nDocNo){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	/*Declaration Section*/
	CRecord rs(&pMF->m_db);
	int nRefIdx = 0;
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL;
	double nTmp = 0;
	int nIdx = 1;
	CString szSQL, tmpStr, szReportName, szDate, szTemp;
	long double nTotal = 0;
	szReportName = _T("Reports\\HMS\\TM_THONGKESUDUNGTHUOCNOITRU.RPT");
	if (!rpt.Init(szReportName))
		return;
	if (nRefIdx <= 0)
	{
		AfxMessageBox(_T("No reference idx"));
		return;
	}
	szSQL.Format(_T(" SELECT    get_departmentname(hcr_dischargedept) dept, ") \
	_T("           hcr_recordno record_no, ") \
	_T("		   hd_docno doc_no,") \
	_T("           Get_patientname(hd_docno) patient_name, ") \
	_T("           Extract(YEAR FROM hp_birthdate) yob, ") \
	_T("           get_syssel_desc('hms_rank', DECODE(0, hd_rank, hp_rank, hp_rank)) rank, ") \
	_T("           NVL(hp_workplace, hms_getaddress(hp_provid, hp_distid, hp_villid)) address, ") \
	_T("           hcr_sumtreat numbertr, ") \
	_T("           hcr_admitdate admit_date, ") \
	_T("           hcr_dischargedate discharge_date, ") \
	_T("           get_username(hcr_treatdoctor) doctor_name, ") \
	_T("           hcr_maindisease diagnostic ") \
	_T(" FROM      hms_patient ") \
	_T(" LEFT JOIN hms_doc ON ( hp_patientno = hd_patientno ) ") \
	_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno ) ") \
	_T(" WHERE     hd_docno = %ld"), nDocNo);

	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
		rptHeader->SetValue(_T("Department"), rs.GetValue(_T("dept")));
		rptHeader->SetValue(_T("DocNo"), rs.GetValue(_T("doc_no")));
		rptHeader->SetValue(_T("RecordNo"), rs.GetValue(_T("record_no")));
		rptHeader->SetValue(_T("PatientName"), rs.GetValue(_T("patient_name")));
		rptHeader->SetValue(_T("yearofbirth"), rs.GetValue(_T("yob")));
		rptHeader->SetValue(_T("rank"), rs.GetValue(_T("rank")));
		rptHeader->SetValue(_T("address"), rs.GetValue(_T("address")));
		rptHeader->SetValue(_T("numbertreat"), rs.GetValue(_T("numbertr")));
		rptHeader->SetValue(_T("admitdate"), CDate::Convert(rs.GetValue(_T("admit_date")), yyyymmdd, ddmmyyyy));
		rptHeader->SetValue(_T("dischargedate"), CDate::Convert(rs.GetValue(_T("discharge_date")), yyyymmdd, ddmmyyyy));
		rptHeader->SetValue(_T("doctorname"), rs.GetValue(_T("doctor_name")));
		rptHeader->SetValue(_T("diagnostic"), rs.GetValue(_T("diagnostic")));
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
		CDate::Convert(m_szToDate), yyyymmdd, ddmmyyyy);
		rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	}
	szSQL.Format(_T(" SELECT   product_name, product_uomname, sum(qty) qty, ") \
				_T("		   pol_price, group_id, sum(amount) amount") \
				_T(" FROM (SELECT    hfe_desc                         product_name, ") \
				_T("           hfe_unit                         product_uomname, ") \
				_T("           hfe_quantity                qty, ") \
				_T("           case when hfe_polprice > 0 then hfe_polprice else product_saleprice3 end pol_price, ") \
				_T("           hfe_group group_id, ") \
				_T("           case when hfe_polprice > 0 then hfe_polprice * hfe_quantity else product_saleprice3*hfe_quantity end amount ") \
				_T(" FROM      hms_clinical_record ") \
				_T(" LEFT JOIN hms_fee ON (hcr_docno = hfe_docno)") \
				_T(" LEFT JOIN m_productitem_view ON (product_item_id= hfe_itemid)") \
				_T(" LEFT JOIN hms_fee_group ON (hfg_id = hfe_group)") \
				_T(" WHERE     hfe_docno = %ld ") \
				_T("	   AND hfe_treattime = %d") \
				_T("       AND hfe_object IN ( 1, 3, 8 ) ") \
				_T("	   AND hfe_category <> 'Y'") \
				_T("       AND hfg_type_id = 800 ") \
				_T("       AND hcrf_acceptedfee IN ( 'Y', 'P' ) )") \
				_T(" GROUP BY	  product_name, ") \
				_T("              product_uomname, ") \
				_T("              pol_price, ") \
				_T("              group_id ") \
				_T(" ORDER     BY group_id, ") \
				_T("              product_name"), nDocNo, nRefIdx);

	int nCount = rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		if (nCount < 0)
			_msg(_T("%s"), szSQL);
		else
			_fmsg(_T("%s"), szSQL);
		return;
	}
	while (!rs.IsEOF())
	{
		/*Group Seperation*/
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("product_name")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("product_uomname")));
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("qty")));
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("pol_price")));
			rs.GetValue(_T("amount"), nTmp);
			nTotal+= nTmp;
			rptDetail->SetValue(_T("6"), double2str(nTmp));
		}
		rs.MoveNext();
	}
	szTemp.Format(_T("%f"), nTotal);
	FormatCurrency(nTotal, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();	
}

void CPrintForms::HMS_InwardDrugUsage(long nDocNo, int nRefIdx){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd(); 
	/*Declaration Section*/
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL;
	double nTmp = 0;
	int nIdx = 1;
	CString szSQL, tmpStr, szReportName, szDate, szTemp;
	long double nTotal = 0;
	szReportName = _T("Reports\\HMS\\TM_THONGKESUDUNGTHUOCNOITRU.RPT");
	if (!rpt.Init(szReportName))
		return;
	if (nRefIdx < 0)
	{
		AfxMessageBox(_T("No reference idx"));
		return;
	}
	szSQL.Format(_T(" SELECT    get_departmentname(hcr_dischargedept) dept, ") \
	_T("           hcr_recordno record_no, ") \
	_T("		   hd_docno doc_no,") \
	_T("           Get_patientname(hd_docno) patient_name, ") \
	_T("           Extract(YEAR FROM hp_birthdate) yob, ") \
	_T("           get_syssel_desc('hms_rank', DECODE(0, hd_rank, hp_rank, hp_rank)) rank, ") \
	_T("           NVL(hp_workplace, hms_getaddress(hp_provid, hp_distid, hp_villid)) address, ") \
	_T("           hcr_sumtreat numbertr, ") \
	_T("           hcr_admitdate admit_date, ") \
	_T("           hcr_dischargedate discharge_date, ") \
	_T("           get_username(hcr_treatdoctor) doctor_name, ") \
	_T("           hcr_maindisease diagnostic ") \
	_T(" FROM      hms_patient ") \
	_T(" LEFT JOIN hms_doc ON ( hp_patientno = hd_patientno ) ") \
	_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno ) ") \
	_T(" WHERE     hd_docno = %ld"), nDocNo);

	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
		rptHeader->SetValue(_T("Department"), rs.GetValue(_T("dept")));
		rptHeader->SetValue(_T("DocNo"), rs.GetValue(_T("doc_no")));
		rptHeader->SetValue(_T("RecordNo"), rs.GetValue(_T("record_no")));
		rptHeader->SetValue(_T("PatientName"), rs.GetValue(_T("patient_name")));
		rptHeader->SetValue(_T("yearofbirth"), rs.GetValue(_T("yob")));
		rptHeader->SetValue(_T("rank"), rs.GetValue(_T("rank")));
		rptHeader->SetValue(_T("address"), rs.GetValue(_T("address")));
		rptHeader->SetValue(_T("numbertreat"), rs.GetValue(_T("numbertr")));
		rptHeader->SetValue(_T("admitdate"), CDate::Convert(rs.GetValue(_T("admit_date")), yyyymmdd, ddmmyyyy));
		rptHeader->SetValue(_T("dischargedate"), CDate::Convert(rs.GetValue(_T("discharge_date")), yyyymmdd, ddmmyyyy));
		rptHeader->SetValue(_T("doctorname"), rs.GetValue(_T("doctor_name")));
		rptHeader->SetValue(_T("diagnostic"), rs.GetValue(_T("diagnostic")));
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
		CDate::Convert(m_szToDate), yyyymmdd, ddmmyyyy);
		rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	}
	szSQL.Format(_T(" SELECT   product_name, product_uomname, sum(qty) qty, ") \
				_T("		   pol_price, group_id, sum(amount) amount") \
				_T(" FROM (SELECT    hfe_desc                         product_name, ") \
				_T("           hfe_unit                         product_uomname, ") \
				_T("           hfe_quantity                qty, ") \
				_T("           case when hfe_polprice > 0 then hfe_polprice else product_saleprice3 end pol_price, ") \
				_T("           hfe_group group_id, ") \
				_T("           case when hfe_polprice > 0 then hfe_polprice * hfe_quantity else product_saleprice3*hfe_quantity end amount ") \
				_T(" FROM      hms_clinical_record") \
				_T(" LEFT JOIN hms_fee ON (hcr_docno = hfe_docno)") \
				_T(" LEFT JOIN m_productitem_view ON (product_item_id = hfe_itemid)") \
				_T(" LEFT JOIN hms_fee_group ON (hfg_id = hfe_group)") \
				_T(" WHERE     hfe_docno = %ld ") \
				_T("	   AND hfe_treattime = %d") \
				_T("       AND hfe_object IN ( 1, 3, 8 ) ") \
				_T("	   AND hfe_category <> 'Y'") \
				_T("       AND hfg_type_id = 800 ") \
				_T("       AND hcrf_acceptedfee IN ( 'Y', 'P' ) )") \
				_T(" GROUP BY	  product_name, ") \
				_T("              product_uomname, ") \
				_T("              pol_price, ") \
				_T("              group_id ") \
				_T(" ORDER     BY group_id, ") \
				_T("              product_name"), nDocNo, nRefIdx);
	int nCount = rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		if (nCount < 0)
			_msg(_T("%s"), szSQL);
		else
			_fmsg(_T("%s"), szSQL);
		return;
	}
	while (!rs.IsEOF())
	{
		/*Group Seperation*/
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("product_name")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("product_uomname")));
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("qty")));
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("pol_price")));
			rs.GetValue(_T("amount"), nTmp);
			nTotal+= nTmp;
			rptDetail->SetValue(_T("6"), double2str(nTmp));
		}
		rs.MoveNext();
	}
	szTemp.Format(_T("%f"), nTotal);
	FormatCurrency(nTotal, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();	
}

void CPrintForms::OnPrintCabinetReturnOrder(long nOrderID){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	CString szSysDte;
	int nIdx = 0;
	double nS7 = 0, nS8 = 0;
	CReportSection *rptDetail, *rptHeader, *rptFooter;
	m_szReportName = _T("Reports\\HMS\\MA_RETURNORDERVOTE.RPT");
	if (!rpt.Init(m_szReportName))
		return;
	szSQL.Format(_T(" SELECT") \
				_T("   mt_orderno       AS orderno,") \
				_T("   get_username(mt_approvedby)    AS clerk,") \
				_T("   get_partnername(mt_partner_id) as deliverto,") \
				_T("   Cast_Timestamp2Date(mt_approveddate)  AS dte,") \
				_T("   get_storagename(mt_storage_id) AS stock,") \
				_T("   get_departmentname(mt_department_id) AS dept,") \
				_T("   mt_description   AS descr,") \
				_T("   mt_documentno documentno") \
				_T(" FROM m_transaction") \
				_T(" WHERE mt_transaction_id = %ld"), nOrderID);	
	int nRet = rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONASTERISK | MB_OK);
		if (nRet < 0)
			_msg(_T("%s"), szSQL);
		return;
	}
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	rptHeader->SetValue(_T("OrderID"), rs.GetValue(_T("orderno")));
	rptHeader->SetValue(_T("DeliverBy"), rs.GetValue(_T("clerk")));
	rptHeader->SetValue(_T("FromStock"), rs.GetValue(_T("stock")));
	rptHeader->SetValue(_T("Comment"), rs.GetValue(_T("descr")));
	rs.GetValue(_T("dte"), szSysDte);
	rptHeader->SetValue(_T("OrderDate"), CDate::Convert(szSysDte, yyyymmdd, ddmmyyyy));
	rptHeader->SetValue(_T("Department"), rs.GetValue(_T("dept")));
	rptHeader->SetValue(_T("DeliveryTo"), rs.GetValue(_T("deliverto")));
	rptHeader->SetValue(_T("ReferenceNo"), rs.GetValue(_T("documentno")));
	szSQL.Format(_T(" SELECT") \
				_T("   product_name   AS pname,") \
				_T("   mtl_qtydelivery	AS qty,") \
				_T("   product_uomname	AS uom,") \
				_T("   product_expdate	AS expdte,") \
				_T("   mtl_saleprice AS price,") \
				_T("   product_lotno	AS lot,") \
				_T("   mtl_qtydelivery*mtl_saleprice AS amt") \
				_T(" FROM m_transactionline") \
				_T(" LEFT JOIN m_productitem_view ON (product_item_id=mtl_product_item_id)") \
				_T(" WHERE mtl_transaction_id = %ld"), nOrderID);	
	nRet = rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		if (nRet < 0)
			_msg(_T("%s"), szSQL);
		return;
	}
	while (!rs.IsEOF())
	{
		nIdx++;
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("pname")));
		rptDetail->SetValue(_T("3"), rs.GetValue(_T("uom")));
		rptDetail->SetValue(_T("4"), rs.GetValue(_T("qty")));
		rptDetail->SetValue(_T("8"), rs.GetValue(_T("price")));
		rs.GetValue(_T("amt"), tmpStr);
		nS8 += str2double(tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);
		rs.MoveNext();
	}
	rptFooter = rpt.AddDetail(rpt.GetGroupFooter(1));
	rptFooter->SetValue(_T("TotalItem"), int2str(nIdx));
	rptFooter->SetValue(_T("TotalAmount"), double2str(nS8));
	MoneyToString(ToString(nS8), tmpStr);
	if (!tmpStr.IsEmpty())
	{
		tmpStr += _T(" \x111\x1ED3ng");
		rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	}
	szSysDte = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDte.Right(2), szSysDte.Mid(5,2), szSysDte.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();	
}

void CPrintForms::FM_OnPrintServiceIncomeList(CString szReceiptIDStr){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp, szWhere, szWhere1;
	CString szSysDate, szReceiptStr;
	//Get Report Date
	szSQL.Format(_T("SELECT min(fac_createddate) from_date, max(fac_createddate) to_date FROM fa_cash WHERE fac_cash_id IN (%s)"), szReceiptIDStr);
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{
		rs.GetValue(_T("from_date"), m_szFromDate);
		rs.GetValue(_T("to_date"), m_szToDate);
	}
	//Get Receipt String
	szSQL.Format(_T("SELECT fac_invoiceno invoice_no FROM fa_cash WHERE fac_cash_id IN (%s)"), szReceiptIDStr);
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		if (!szReceiptStr.IsEmpty())
			szReceiptStr += _T(", ");
		szReceiptStr += rs.GetValue(_T("invoice_no"));
		rs.MoveNext();
	}
	szWhere.Format(_T(" AND hfe_cash_id IN (%s) AND hfe_status = 'P' AND hfe_objectid = 7"), szReceiptIDStr);
	szWhere1.Format(_T(" AND hfe_cash_id IN (%s) AND hfe_status IN ('P', 'R') AND hfe_objectid = 7"), szReceiptIDStr);
	szSQL.Format(_T(" SELECT docno, ") \
				_T("		get_patientname(docno) pname,") \
				_T("		hcr_recordno recordno,") \
				_T("		ho_type object_type,") \
				_T("        deptid, ") \
				_T("		hfe_date,") \
				_T("        SUM(patpaid)               AS eamount, ") \
				_T("        SUM(deposit)               AS deposit, ") \
				_T("        SUM(inpatient_cost)        AS iamount, ") \
				_T("        SUM(inpatient_depositpaid) AS inpatient_deposit_paid, ") \
				_T("        SUM(inpatient_payment)     AS inpatient_payment, ") \
				_T("		SUM(patpaid+inpatient_payment+deposit) AS totalamount") \
				_T(" FROM   (SELECT hfe_docno   AS docno, ") \
				_T("				hfe_objectid,") \
				_T("                hfe_deptid  AS deptid, ") \
				_T("				hfe_date,") \
				_T("                hfe_patpaid AS patpaid, ") \
				_T("                hfe_deposit AS deposit, ") \
				_T("                0           AS inpatient_cost, ") \
				_T("                0           inpatient_depositpaid, ") \
				_T("                0           inpatient_payment ") \
				_T("         FROM   hms_fee_invoice ") \
				_T("         WHERE  hfe_class IN ('E', 'X') %s") \
				_T("         UNION ALL ") \
				_T("         SELECT hfe_docno, ") \
				_T("				hfe_objectid,") \
				_T("                hfe_deptid, ") \
				_T("				hfe_date,") \
				_T("                0, ") \
				_T("                hfe_amount, ") \
				_T("                0 AS inpatient_cost, ") \
				_T("                0 inpatient_depositpaid, ") \
				_T("                0 inpatient_payment ") \
				_T("         FROM   hms_fee_deposit ") \
				_T("         WHERE  hfe_class IN( 'I', 'A' ) %s") \
				_T("         UNION ALL ") \
				_T("         SELECT hfe_docno   AS docno, ") \
				_T("				hfe_objectid,") \
				_T("                hfe_deptid  AS deptid, ") \
				_T("				hfe_date,") \
				_T("                0, ") \
				_T("                0, ") \
				_T("                hfe_amount  AS inpatient_cost, ") \
				_T("                CASE WHEN hfe_deposit > hfe_patpaid THEN hfe_patpaid ") \
				_T("                ELSE hfe_deposit ") \
				_T("                END         AS inpatient_depositpaid, ") \
				_T("                hfe_payment AS inpatient_payment ") \
				_T("         FROM   hms_fee_invoice ") \
				_T("         WHERE  hfe_class IN( 'A', 'I' ) %s) ") \
				_T(" LEFT JOIN hms_clinical_record ON (hcr_docno = docno)")\
				_T(" LEFT JOIN hms_doc ON (hd_docno = docno) ") \
				_T(" LEFT JOIN hms_object ON (hfe_objectid = ho_id)") \
				_T(" WHERE 1=1") \
				_T(" GROUP  BY docno, ") \
				_T("		   hcr_recordno,") \
				_T("		   ho_type,") \
				_T("           deptid,") \
				_T("		   hfe_date ") \
				_T(" HAVING SUM(patpaid+inpatient_payment+deposit) > 0") \
				_T(" ORDER BY hfe_date"), szWhere, szWhere1, szWhere);
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}
	
	if (!rpt.Init(_T("Reports/HMS/HF_BANGKETHUBNDICHVU.RPT")))
		return;

	StringUpper(pMF->m_szHealthService, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);

	StringUpper(pMF->m_szHospitalName, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ReceiptStr"), szReceiptStr);
	szTemp.Empty();
	//szTemp stores distinct object type
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("object_type"), tmpStr);
		if (szTemp.Find(tmpStr) <= 0)
		{
			if (!szTemp.IsEmpty())
				szTemp += _T(", ");
			szTemp += tmpStr;
		}
		rs.MoveNext();
	}
	//Object Type Token
	tmpStr.Empty();
	CStringToken token(szTemp, _T(","));
	szTemp.Empty();
	for (int i = 0; i < token.GetSize(); i++)
	{
		if (!tmpStr.IsEmpty())
			tmpStr += _T(", ");
		token.GetAt(i, szTemp);
		tmpStr += pMF->GetObjectType();
	}
	rs.MoveFirst();
	CReportSection *rptDetail;
	CString szOldLine, szNewLine;
	CStringArray arrField;
	long double nTotal[6];
	double nCost;
	int nIndex = 1;

	for (int i = 0; i < 6; i++)
	{
		nTotal[i] = 0;
	}
	arrField.Add(_T("eamount"));
	arrField.Add(_T("deposit"));
	arrField.Add(_T("iamount"));
	arrField.Add(_T("inpatient_deposit_paid"));
	arrField.Add(_T("inpatient_payment"));
	arrField.Add(_T("totalamount"));
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("recordno"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		//rs.GetValue(_T("deptid"), tmpStr);
		//rptDetail->SetValue(_T("5"), tmpStr);

		for (int j = 0; j < 6; j++)
		{
			rs.GetValue(arrField.GetAt(j), nCost);
			nTotal[j] += nCost;
			FormatCurrency(nCost, tmpStr);
			szTemp.Format(_T("%d"), j+5);
			rptDetail->SetValue(szTemp, tmpStr);	
		}		
		rs.MoveNext();
	}

	if (nTotal[5] > 0)
	{
		for (int i = 0; i < 6; i++)
		{
			FormatCurrency(nTotal[i], tmpStr);
			szTemp.Format(_T("s%d"), i + 5);
			rpt.GetReportFooter()->SetValue(szTemp, tmpStr);
		}
	}

	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("User"), pMF->GetDoctorName(pMF->GetCurrentUser()));
	rpt.PrintPreview();
}

void CPrintForms::FM_OnPrintServiceOutlayList(CString szPaymentIDStr)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp, szWhere, szRefWhere, szRefWhere0;
	CString szSysDate, szPaymentStr;
	//Get Report Date
	szSQL.Format(_T("SELECT min(fac_createddate) from_date, max(fac_createddate) to_date FROM fa_cash WHERE fac_cash_id IN (%s)"), szPaymentIDStr);
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{
		rs.GetValue(_T("from_date"), m_szFromDate);
		rs.GetValue(_T("to_date"), m_szToDate);
	}
	//Get Receipt String
	szSQL.Format(_T("SELECT fac_invoiceno invoice_no FROM fa_cash WHERE fac_cash_id IN (%s)"), szPaymentIDStr);
	rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		if (!szPaymentStr.IsEmpty())
			szPaymentStr += _T(", ");
		szPaymentStr += rs.GetValue(_T("invoice_no"));
		rs.MoveNext();
	}
	szWhere.Format(_T(" AND hfe_cash_id IN (%s) AND hfe_status = 'P'"), szPaymentIDStr);
	szRefWhere.Format(_T(" AND (r.hfe_cash_id IN (%s) AND r.hfe_status = 'P')"), szPaymentIDStr);
	szRefWhere0.Format(_T(" AND rf.hfe_cash_id IN (%s) AND rf.hfe_status = 'P'"), szPaymentIDStr);
	szSQL.Format(_T(" SELECT    Get_patientname(docno) pname, ") \
				_T("           hcr_recordno recordno, ") \
				_T("           docno, ") \
				_T("           deptid, ") \
				_T("		   hfe_date,") \
				_T("		   SUM(hitechreturn) hitechreturn,") \
				_T("           SUM(eamount) eamount, ") \
				_T("           SUM(iamount) iamount, ") \
				_T("           SUM(deposit) deposit, ") \
				_T("           SUM(refund)   AS refund, ") \
				_T("		   SUM(hitechreturn+eamount+refund) AS totalamount")
				_T(" FROM      (SELECT    rf.hfe_docno AS docno, ") \
				_T("                      rf.hfe_deptid AS deptid, ") \
				_T("                      rf.hfe_date hfe_date, ") \
				_T("                      SUM(CASE WHEN NVL(hfl_hitechmachine, 'X') = 'Y' THEN hfe_cost ELSE 0 END) AS hitechreturn, ") \
				_T("                      SUM(CASE WHEN NVL(hfl_hitechmachine, 'X') <> 'Y' THEN hfe_cost ELSE 0 END) AS eamount, ") \
				_T("                      0 AS iamount, ") \
				_T("                      0 AS deposit, ") \
				_T("                      0 AS refund") \
				_T("            FROM      hms_fee_refund rf ") \
				_T("            LEFT JOIN hms_fee_refundline rfl ON ( rf.hfe_docno = rfl.hfe_docno ") \
				_T("                                                  AND rf.hfe_invoiceno = rfl.hfe_invoiceno ) ") \
				_T("            LEFT JOIN hms_fee_list ON ( hfe_itemid = hfl_feeid ) ") \
				_T("            WHERE     rf.hfe_class IN ('E', 'X') %s") \
				_T("            GROUP     BY rf.hfe_docno, ") \
				_T("                         rf.hfe_deptid, ") \
				_T("                         rf.hfe_date ") \
				_T("			UNION ALL") \
				_T("			SELECT r.hfe_docno  AS docno, ") \
				_T("                   r.hfe_deptid AS deptid, ") \
				_T("                   r.hfe_date, ") \
				_T("				   0,") \
				_T("                   0		  AS eamount, ") \
				_T("                   0          AS iamount, ") \
				_T("                   0          AS deposit, ") \
				_T("                   r.hfe_amount AS refund ") \
				_T("            FROM   hms_fee_refund r") \
				_T("			LEFT JOIN hms_fee_deposit d ON (r.hfe_refidx = d.hfe_invoiceno)") \
				_T("            WHERE  r.hfe_class IN ( 'A', 'I' ) AND (d.hfe_status = 'R' OR r.hfe_type = 'E') %s") \
				_T("            UNION ALL ") \
				_T("            SELECT i.hfe_docno   AS docno, ") \
				_T("                   i.hfe_deptid  AS deptid, ") \
				_T("                   i.hfe_date, ") \
				_T("				   0,") \
				_T("                   0, ") \
				_T("                   i.hfe_amount, ") \
				_T("                   i.hfe_deposit, ") \
				_T("                   r.hfe_amount ") \
				_T("            FROM   hms_fee_refund r") \
				_T("			LEFT JOIN hms_fee_invoice i ON (i.hfe_docno = r.hfe_docno AND i.hfe_invoiceno = r.hfe_refidx)") \
				_T("            WHERE  r.hfe_class IN( 'A', 'I' ) AND r.hfe_type = 'F' AND i.hfe_refund > 0 %s) ") \
				_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = docno ) ") \
				_T(" LEFT JOIN hms_doc ON ( hd_docno = docno ) ") \
				_T(" WHERE     1 = 1") \
				_T(" GROUP     BY docno, ") \
				_T("              hcr_recordno, ") \
				_T("              deptid, ") \
				_T("              hfe_date ") \
				_T(" HAVING sum(hitechreturn+eamount+refund) > 0") \
				_T(" ORDER     BY hfe_date "), szRefWhere0, szRefWhere, szRefWhere);
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/HF_BANGKECHIBNDICHVU.RPT")))
		return;

	StringUpper(pMF->m_szHealthService, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), tmpStr);

	StringUpper(pMF->m_szHospitalName, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), tmpStr);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("PaymentStr"), szPaymentStr);
	CReportSection *rptDetail;
	CString szOldLine, szNewLine;
	CStringArray arrField;
	long double nTotal[6];
	double nCost;
	int nIndex = 1;

	for (int i = 0; i < 6; i++)
	{
		nTotal[i] = 0;
	}
	arrField.Add(_T("hitechreturn"));
	arrField.Add(_T("eamount"));
	arrField.Add(_T("iamount"));
	arrField.Add(_T("deposit"));
	arrField.Add(_T("refund"));
	arrField.Add(_T("totalamount"));
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("recordno"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		//rs.GetValue(_T("deptid"), tmpStr);
		//rptDetail->SetValue(_T("5"), tmpStr);

		for (int j = 0; j < 6; j++)
		{
			rs.GetValue(arrField.GetAt(j), nCost);
			nTotal[j] += nCost;
			FormatCurrency(nCost, tmpStr);
			szTemp.Format(_T("%d"), j+5);
			rptDetail->SetValue(szTemp, tmpStr);	
		}		

		rs.MoveNext();
	}

	if (nTotal[5] > 0)
	{
		for (int i = 0; i < 6; i++)
		{
			FormatCurrency(nTotal[i], tmpStr);
			szTemp.Format(_T("s%d"), i + 5);
			rpt.GetReportFooter()->SetValue(szTemp, tmpStr);
		}
	}

	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.GetReportFooter()->SetValue(_T("User"), pMF->GetDoctorName(pMF->GetCurrentUser()));
	rpt.PrintPreview();	
}

void CPrintForms::TM_OnPrintAppointmentSheet(long nDocNo, CString szDeptId, bool bPreview, bool bDirect){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CReportSection *rptHeader;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, tmpStr2;
	if (!rpt.Init(_T("Reports\\HMS\\GIAYHEN.RPT")))
		return;
	rptHeader = rpt.GetReportHeader();
	szSQL.Format(_T(" SELECT hcr_recordno record_no, ") \
				 _T("	get_patientname(hcr_docno) patient_name, ") \
				 _T("	extract(year from hp_birthdate) yob, ") \
				 _T("	get_objectname(hd_object) obj_name,") \
				 _T("	hms_getaddress(hp_provid, hp_distid, hp_villid) address") \
				 _T(" FROM hms_clinical_record ") \
				 _T(" LEFT JOIN hms_doc ON (hd_docno = hcr_docno)") \
				 _T(" LEFT JOIN hms_patient ON (hcr_patientno = hp_patientno) ") \
				 _T(" WHERE hcr_docno = %ld"), nDocNo);
	int nRet = rs.ExecSQL(szSQL);
	if (nRet < 0)
		_msg(_T("%s"), szSQL);
	if (!rs.IsEOF())
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
		rptHeader->SetValue(_T("Department"), pMF->GetDepartmentName(szDeptId));
		rptHeader->SetValue(_T("documentno"), nDocNo);
		rptHeader->SetValue(_T("recordno"), rs.GetValue(_T("record_no")));
		rs.GetValue(_T("patient_name"), tmpStr2);
		StringUpper(tmpStr2, tmpStr);
		rptHeader->SetValue(_T("patientname"), tmpStr);
		rptHeader->SetValue(_T("yearofbirth"), rs.GetValue(_T("yob")));
		rptHeader->SetValue(_T("object"), rs.GetValue(_T("obj_name")));
		rptHeader->SetValue(_T("address"), rs.GetValue(_T("address")));
	}
	if (bPreview)
		rpt.PrintPreview();
	else
		rpt.Print(bDirect);
}

void CPrintForms::TM_OnPrintOperationMinutes(long nOrderID){
	
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CReport rpt;
	CString szSQL, tmpStr;
	CRecord rs(&pMF->m_db);
	CString szDate, szSysDate, szAdmit, szPerformDate;
	szSysDate = pMF->GetSysDateTime(); 

	if(pMF->GetModuleID() == _T("EM") && !pMF->CheckPermission(_T("04.04")))
	{
		ShowMessageBox(_T("Permission Denined."), 0);
		return;
	}

	if(pMF->GetModuleID() == _T("TM") && !pMF->CheckPermission(_T("06.04")))
	{
		ShowMessageBox(_T("Permission Denined."), 0);
		return;
	}


	szSQL.Format(_T(" SELECT 	hd_docno as docno,") \
			_T("  trim(hp_surname||' '||hp_midname||' '||hp_firstname) as PatientName,") \
			_T("  trunc_date(hp_birthdate) as birthdate,") \
			_T("  hms_getage(trunc_date(hd_admitdate), hp_birthdate) as Age,") \
			_T("  hp_sex as sexid,") \
			_T("  sys_sel_getname('sys_sex', hp_sex) as Sex,") \
			_T("  ho_deptid as deptid,") \
			_T("  ho_roomid as roomid,") \
			_T("  ho_bedid as bedid,") \
			_T("  (select sd_name from sys_dept where sd_id=ho_deptid) as Department,") \
			_T("  hms_getaddress(hp_provid, hp_distid, hp_villid) as address,") \
			_T("  hd_cardno as cardno,") \
			_T("  hd_cardidx as cardidx, ") \
			_T("  hc_expdate as expdate,") \
			_T("  hd_admitdate as admitdate, ") \
			_T("  ho_beforeopera  as  BeforeOpera, ho_afteropera  as  AfterOpera, ") \
			_T("  ho_performdate as PerformDate,") \
			_T("  ho_practitioner as practitioner, ho_assistant as assistant,") \
			_T("  ho_anesthetist as anesthetist, ") \
			_T("  hd_diagnostic as Diagnostic, ho_inmethod as InMethod,") \
			_T("  ho_itemid as itemid, ") \
			_T("  hfl_name as operationname ") \
			_T("  FROM hms_operation") \
			_T("  RIGHT JOIN hms_patient ON(hp_patientno=ho_patientno)") \
			_T("  RIGHT JOIN hms_doc on (hd_docno=ho_docno) ") \
			_T("  LEFT JOIN hms_card ON (hd_patientno=hc_patientno AND hd_cardno = hc_cardno AND hd_cardidx=hc_idx)") \
			_T("  LEFT JOIN hms_fee_list ON(hfl_feeid=ho_itemid)") \
			_T("  WHERE ho_idx=%ld"), nOrderID);
	int ret = rs.ExecSQL(szSQL);
//_fmsg(_T("%s"), szSQL);

	if(rs.IsEOF())
	{
		ShowMessageBox(_T("Patient not found."));
		return;
	}
	if(!rpt.Init(_T("Reports/HMS/HMS_OPERATIONSHEET.RPT")) )
		return ;
	//Report Header

	//StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HEALTH_SERVICE"), pMF->m_CompanyInfo.sc_pname);
	//StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("HOSPITAL_NAME"), pMF->m_CompanyInfo.sc_name);
	rs.GetValue(_T("docno"),tmpStr);
	rpt.GetReportHeader()->SetValue(_T("docno"), tmpStr);
	tmpStr.Format(_T("%ld"), nOrderID);
	rpt.GetReportHeader()->SetValue(_T("orderid"), tmpStr);

	rs.GetValue(_T("PatientName"),tmpStr); //benh nhan
	tmpStr.MakeUpper();
	rpt.GetReportHeader()->SetValue(_T("PATIENTNAME"), tmpStr);
	rs.GetValue(_T("Age"),tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Age"), tmpStr);
	rs.GetValue(_T("Sex"),tmpStr);
	rpt.GetReportHeader()->SetValue(_T("Sex"), tmpStr);
	
	rs.GetValue(_T("CardNo"),tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CardNo"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("ppvocam"), _T(""));
	rs.GetValue(_T("ExpDate"),tmpStr);// ngay het han
	rpt.GetReportHeader()->SetValue(_T("ExpDate"), CDate::Convert(rs.GetValue(_T("ExpDate")), yyyymmdd, ddmmyyyy));
	rs.GetValue(_T("Practitioner"),tmpStr);// bac sy phau thuat
	rpt.GetReportHeader()->SetValue(_T("Practitioner"), pMF->GetDoctorName(tmpStr));
	rs.GetValue(_T("Assistant"),tmpStr);// thu thuat vien
	rpt.GetReportHeader()->SetValue(_T("Assistant"), pMF->GetDoctorName(tmpStr));
	rs.GetValue(_T("Anesthetist"),tmpStr);//bac sy gay me
	rpt.GetReportHeader()->SetValue(_T("Anesthetist"), pMF->GetDoctorName(tmpStr));
	rs.GetValue(_T("address"),tmpStr);//dia chi
	rpt.GetReportHeader()->SetValue(_T("Address"), tmpStr);
	rs.GetValue(_T("Department"),tmpStr);//khoa
	rpt.GetReportHeader()->SetValue(_T("Department"), tmpStr);
	rs.GetValue(_T("roomid"),tmpStr);// ten phong
	rpt.GetReportHeader()->SetValue(_T("Room"), tmpStr);
	rs.GetValue(_T("bedid"),tmpStr);//giuong
	rpt.GetReportHeader()->SetValue(_T("Bed"), tmpStr);

	tmpStr=rpt.GetReportHeader()->GetValue(_T("AdmitDate"));
	rs.GetValue(_T("AdmitDate"),szAdmit);// vao khoa luc			
	szAdmit.Format(tmpStr,szAdmit.Mid(11,2),szAdmit.Mid(14,2),CDate::Convert(szAdmit.Left(10), yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("AdmitDate"), szAdmit);//

	tmpStr=rpt.GetReportHeader()->GetValue(_T("PerformDate"));
	rs.GetValue(_T("PerformDate"),szPerformDate);// thuc hien luc		
	szPerformDate.Format(tmpStr,szPerformDate.Mid(11,2),szPerformDate.Mid(14,2),CDate::Convert(szPerformDate.Left(10), yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("PerformDate"), szPerformDate);//
	rs.GetValue(_T("Diagnostic"),tmpStr);// chan doan
	rpt.GetReportHeader()->SetValue(_T("Diagnostic"), tmpStr);
	rs.GetValue(_T("BeforeOpera"),tmpStr);// truoc phau thuat thu thuat
	rpt.GetReportHeader()->SetValue(_T("BeforeOpera"), tmpStr);
	rs.GetValue(_T("AfterOpera"),tmpStr);// sau phau thuat thu thuat
	rpt.GetReportHeader()->SetValue(_T("AfterOpera"), tmpStr);
	rs.GetValue(_T("operationname"),tmpStr);// phuong phap phau thuat thu thuat
	rpt.GetReportHeader()->SetValue(_T("InMethod"), tmpStr);
	rs.GetValue(_T("OperationType"),tmpStr);// loai phau thuat thu thuat
	rpt.GetReportHeader()->SetValue(_T("OperationType"), _T(""));
	//phuong phap vo cam			
			
	//Page Header
	//Report Detail
	CReportSection* rptDetail = rpt.GetDetail(0); 
		
	//Page Footer
	//Report Footer
			
	rs.GetValue(_T("Assistant"),tmpStr);
	rpt.GetReportFooter()->SetValue(_T("Assistant"), pMF->GetDoctorName(tmpStr));
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Mid(11,5),szSysDate.Mid(8,2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	
	rpt.PrintPreview();

} 

void CPrintForms::TM_OnPrintOperationSheet(long nOrderID){
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CReport rpt;
	CReportSection *rptHeader = NULL;
	CString szSQL, tmpStr;
	CRecord rs(&pMF->m_db);
	CString szDate, szSysDate, szPerformDate;
	szSysDate = pMF->GetSysDateTime(); 

	if(pMF->GetModuleID() == _T("EM") && !pMF->CheckPermission(_T("04.04")))
	{
		ShowMessageBox(_T("Permission Denined."), 0);
		return;
	}

	if(pMF->GetModuleID() == _T("TM") && !pMF->CheckPermission(_T("06.04")))
	{
		ShowMessageBox(_T("Permission Denined."), 0);
		return;
	}


	szSQL.Format(_T(" SELECT   hd_docno AS docno, ") \
				_T("           Trim(hp_surname||' '||hp_midname||' '||hp_firstname) AS PatientName, ") \
				_T("           Trunc_date(hp_birthdate) AS birthdate, ") \
				_T("           Hms_getage(Trunc_date(hd_admitdate), hp_birthdate) AS Age, ") \
				_T("           Sys_sel_getname('sys_sex', hp_sex) AS Sex, ") \
				_T("           ho_deptid AS deptid, ") \
				_T("           (SELECT sd_name FROM sys_dept WHERE  sd_id = ho_deptid)AS Department, ") \
				_T("           Hms_getaddress(hp_provid, hp_distid, hp_villid)    AS address, ") \
				_T("           hcr_admitdate AS admitdate, ") \
				_T("           ho_performdate AS PerformDate, ") \
				_T("           ho_practitioner AS practitioner, ") \
				_T("           ho_assistant AS assistant, ") \
				_T("           ho_anesthetist AS anesthetist, ") \
				_T("           ho_diagnostic AS Diagnostic, ") \
				_T("           (SELECT ham_name FROM hms_anesthesia_method WHERE ham_idx = ho_anes_method) AS InMethod, ") \
				_T("           ho_comment AS operationname, ") \
				_T("           hcr_recordno record_no, ") \
				_T("           Extract (year FROM hp_birthdate) yob, ") \
				_T("           Get_objectname(hd_object) obj_name, ") \
				_T("           Get_syssel_desc('sys_occupation', hp_occupation) occupation, ") \
				_T("           hcr_dischargedate discharge_date, ") \
				_T("           hcr_reason reason ") \
				_T(" FROM      hms_operation ") \
				_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = ho_docno ) ") \
				_T(" LEFT JOIN hms_patient ON( hp_patientno = ho_patientno ) ") \
				_T(" LEFT JOIN hms_doc ON ( hd_docno = ho_docno ) ") \
				_T(" WHERE     ho_idx =% ld "), nOrderID);

	int ret = rs.ExecSQL(szSQL);
//_fmsg(_T("%s"), szSQL);

	if(rs.IsEOF())
	{
		ShowMessageBox(_T("Patient not found."));
		return;
	}
	if(!rpt.Init(_T("Reports/HMS/HMS_OPERATIONSHEET2.RPT")) )
		return ;
	//Report Header
	rptHeader = rpt.GetReportHeader();
	//StringUpper(pMF->m_CompanyInfo.sc_pname, tmpStr);
	rptHeader->SetValue(_T("HEALTH_SERVICE"), pMF->m_CompanyInfo.sc_pname);
	//StringUpper(pMF->m_CompanyInfo.sc_name, tmpStr);
	rptHeader->SetValue(_T("HOSPITAL_NAME"), pMF->m_CompanyInfo.sc_name);
	rptHeader->SetValue(_T("docno"), rs.GetValue(_T("docno")));
	rptHeader->SetValue(_T("recordno"), rs.GetValue(_T("record_no")));

	StringUpper(rs.GetValue(_T("PatientName")), tmpStr);
	rptHeader->SetValue(_T("PATIENTNAME"), tmpStr);
	rptHeader->SetValue(_T("Age"), rs.GetValue(_T("yob")));
	rptHeader->SetValue(_T("Sex"), rs.GetValue(_T("sex")));
	rptHeader->SetValue(_T("object"), rs.GetValue(_T("obj_name")));
	rptHeader->SetValue(_T("occupation"), rs.GetValue(_T("occupation")));
	rptHeader->SetValue(_T("Practitioner"), pMF->GetDoctorName(rs.GetValue(_T("practitioner"))));
	rptHeader->SetValue(_T("Address"), rs.GetValue(_T("address")));
	rptHeader->SetValue(_T("Department"), rs.GetValue(_T("Department")));

	rs.GetValue(_T("AdmitDate"),szDate);		
	tmpStr.Format(rptHeader->GetValue(_T("admitdate")),szDate.Mid(11,2),szDate.Mid(14,2)
		,CDate::Convert(szDate, yyyymmdd, ddmmyyyy));
	rptHeader->SetValue(_T("AdmitDate"), tmpStr);
	rs.GetValue(_T("Discharge_date"), szDate);
	tmpStr.Format(rptHeader->GetValue(_T("DischargeDate")), szDate.Mid(11, 2), szDate.Mid(14, 2), 
		CDate::Convert(szDate, yyyymmdd, ddmmyyyy));
	rptHeader->SetValue(_T("DischargeDate"), tmpStr);
	rs.GetValue(_T("PerformDate"),szDate);	
	tmpStr.Format(rpt.GetPageHeader()->GetValue(_T("PerformDate")),szDate.Mid(11,2),szDate.Mid(14,2)
		,CDate::Convert(szDate, yyyymmdd, ddmmyyyy));
	rptHeader->SetValue(_T("PerformDate"), tmpStr);
	rptHeader->SetValue(_T("Diagnostic"), rs.GetValue(_T("diagnostic")));
	rs.GetValue(_T("operationname"),tmpStr);// cach mo
	rptHeader->SetValue(_T("TreatMethod"), tmpStr);
	rs.GetValue(_T("inmethod"),tmpStr);// phuong phap gay me
	rptHeader->SetValue(_T("InMethod"), tmpStr);
	rptHeader->SetValue(_T("Reason"), rs.GetValue(_T("reason")));		
			
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Mid(11,5),szSysDate.Mid(8,2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	
	rpt.PrintPreview();	
}

void CPrintForms::EM_OnPrintActivityTreatment(long nOrder_id){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	if (!rpt.Init(_T("Reports\\HMS\\DIEUTRIVANDONG.RPT")))
		return;
	CString szSQL, tmpStr, szCurDate;
	CReportSection *rptHeader = NULL, *rptGroup = NULL;
	CRecord rs(&pMF->m_db);
	szSQL.Format(_T(" SELECT get_patientname(ho_docno) patient_name, ho_docno doc_no, extract(year from hp_birthdate) yob, get_objectname(hd_object) object_name,") \
				 _T("		get_syssel_desc('hms_rank', NVL(hd_rank, hp_rank)) rank, get_syssel_desc('sys_sex', hp_sex) sex, ") \
				 _T("		hp_workplace work_place, hms_getaddress(hp_provid, hp_distid, hp_villid) address, ") \
				 _T("		hd_cardno card_no, hcr_admitdate admit_date, hd_suggestion, hcr_recordno record_no, hd_conclusion conclusion ") \
				 _T(" FROM hms_operation ") \
				 _T(" LEFT JOIN hms_doc ON (hd_docno = ho_docno) ") \
				 _T(" LEFT JOIN hms_patient ON (hd_patientno = hp_patientno)") \
				 _T(" LEFT JOIN hms_clinical_record ON (hcr_docno = hd_docno)") \
				 _T(" WHERE ho_idx = %ld"), nOrder_id);
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{
		rptHeader = rpt.GetReportHeader();
		if (rptHeader)
		{
			//rptHeader->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
			//rptHeader->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
			//rptHeader->SetValue(_T("Department"), pMF->GetCurrentDepartmentName());
			rptHeader->SetValue(_T("DocumentNo"), rs.GetValue(_T("doc_no")));
			StringUpper(rs.GetValue(_T("patient_name")), tmpStr);
			rptHeader->SetValue(_T("PatientName"), tmpStr);
			rptHeader->SetValue(_T("yearofbirth"), rs.GetValue(_T("yob")));
			rptHeader->SetValue(_T("rank"), rs.GetValue(_T("rank")));
			rptHeader->SetValue(_T("sex"), rs.GetValue(_T("sex")));
			rptHeader->SetValue(_T("Workplace"), rs.GetValue(_T("work_place")));
			rptHeader->SetValue(_T("address"), rs.GetValue(_T("address")));
			rptHeader->SetValue(_T("objectname"), rs.GetValue(_T("object_name")));
			rptHeader->SetValue(_T("cardno"), rs.GetValue(_T("card_no")));
			rs.GetValue(_T("hd_suggestion"), tmpStr);
			if (tmpStr == _T("C") || tmpStr == _T("D"))
			{
				rs.GetValue(_T("admit_date"), tmpStr);
				rptHeader->SetValue(_T("Admitdate"), CDateTime::Convert(tmpStr, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
				//rptHeader->SetValue(_T("RecordNo"), rs.GetValue(_T("RecordNo")));
			}
			else
			{
				rptHeader->SetValue(_T("Diagnostic"), rs.GetValue(_T("conclusion")));
			}
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));
			rptGroup->SetPageBreak();
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(2));
		}
		else
		{
			_debug(_T("X"));
		}
	}
	else
		_debug(_T("Y"));

	rpt.PrintPreview();
}

void CPrintForms::EM_OnPrintPhysicalTreatment(long nOrder_id){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	if (!rpt.Init(_T("Reports\\HMS\\DIEUTRIVATLY.RPT")))
		return;
	CString szSQL, tmpStr, szCurDate;
	CReportSection *rptHeader = NULL, *rptGroup = NULL;
	CRecord rs(&pMF->m_db);
	szSQL.Format(_T(" SELECT get_patientname(ho_docno) patient_name, ho_docno doc_no, extract(year from hp_birthdate) yob, get_objectname(hd_object) object_name,") \
				 _T("		get_syssel_desc('hms_rank', NVL(hd_rank, hp_rank)) rank, get_syssel_desc('sys_sex', hp_sex) sex, ") \
				 _T("		hp_workplace work_place, hms_getaddress(hp_provid, hp_distid, hp_villid) address, ") \
				 _T("		hd_cardno card_no, hcr_admitdate admit_date, hd_suggestion, hcr_recordno record_no, hd_conclusion conclusion ") \
				 _T(" FROM hms_operation ") \
				 _T(" LEFT JOIN hms_doc ON (hd_docno = ho_docno) ") \
				 _T(" LEFT JOIN hms_patient ON (hd_patientno = hp_patientno) ") \
				 _T(" LEFT JOIN hms_clinical_record ON (hcr_docno = hd_docno) ") \
				 _T(" WHERE ho_idx = %ld"), nOrder_id);
	rs.ExecSQL(szSQL);
	if (!rs.IsEOF())
	{
		rptHeader = rpt.GetReportHeader();
		if (rptHeader)
		{
			//rptHeader->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
			//rptHeader->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
			//rptHeader->SetValue(_T("Department"), pMF->GetCurrentDepartmentName());
			rptHeader->SetValue(_T("DocumentNo"), rs.GetValue(_T("doc_no")));
			StringUpper(rs.GetValue(_T("patient_name")), tmpStr);
			rptHeader->SetValue(_T("PatientName"), tmpStr);
			rptHeader->SetValue(_T("yearofbirth"), rs.GetValue(_T("yob")));
			rptHeader->SetValue(_T("rank"), rs.GetValue(_T("rank")));
			rptHeader->SetValue(_T("sex"), rs.GetValue(_T("sex")));
			rptHeader->SetValue(_T("Workplace"), rs.GetValue(_T("work_place")));
			rptHeader->SetValue(_T("address"), rs.GetValue(_T("address")));
			rptHeader->SetValue(_T("objectname"), rs.GetValue(_T("object_name")));
			rptHeader->SetValue(_T("cardno"), rs.GetValue(_T("card_no")));
			rs.GetValue(_T("hd_suggestion"), tmpStr);
			if (tmpStr == _T("C") || tmpStr == _T("D"))
			{
				rs.GetValue(_T("admit_date"), tmpStr);
				rptHeader->SetValue(_T("Admitdate"), CDateTime::Convert(tmpStr, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
				//rptHeader->SetValue(_T("RecordNo"), rs.GetValue(_T("RecordNo")));
			}
			else
			{
				rptHeader->SetValue(_T("Diagnostic"), rs.GetValue(_T("conclusion")));
			}
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(2));
			rptGroup->SetPageBreak();
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));
		}
	}
	rpt.PrintPreview();

}

void CPrintForms::TM_OnPrintRadioactiveTreatmentFee(long nDocNo){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL;
	double nTmp = 0;
	int nIdx = 1;
	CString szSQL, tmpStr, szReportName, szDate, szTemp;
	long double nTotal = 0;
	szReportName = _T("Reports\\HMS\\BANGKECHIPHIDICHVUKHAMDIEUTRIXATRIGIATOC.rpt");
	if (!rpt.Init(szReportName))
		return;

	szSQL.Format(_T("SELECT get_patientname(hd_docno) as patient_name, hd_docno as doc_no, hcr_recordno as record_no, ") \
				_T("		hd_conclusion as conclusion, ") \
				_T("		(select sp_name from sys_prov where sp_id = hp_provid) as address, hd_object obj, hd_cardno card_no, ") \
				_T("		extract(year from hp_birthdate) as yob") \
				_T(" FROM hms_doc ") \
				_T(" LEFT JOIN hms_patient ON (hp_patientno = hd_patientno) ") \
				_T(" LEFT JOIN hms_clinical_record ON (hcr_docno = hd_docno) ") \
				_T(" WHERE hd_docno = %ld"), nDocNo);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("PATIENTNAME"), rs.GetValue(_T("patient_name")));
	rpt.GetReportHeader()->SetValue(_T("Yearofbirth"), rs.GetValue(_T("yob")));
	rpt.GetReportHeader()->SetValue(_T("Address"), rs.GetValue(_T("address")));
	rpt.GetReportHeader()->SetValue(_T("Diagnostic"), rs.GetValue(_T("conclusion")));
	rpt.GetReportHeader()->SetValue(_T("Recordno"), rs.GetValue(_T("record_no")));
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), rs.GetValue(_T("doc_no")));
	rpt.GetReportHeader()->SetValue(_T("ObjectName"), rs.GetValue(_T("obj")));
	rpt.GetReportHeader()->SetValue(_T("CardNo"), rs.GetValue(_T("card_no")));

	szSQL.Format(_T("select hfe_desc, hfe_unit, hfe_quantity, hfe_unitprice, diff_cost from ") \
				_T(" (select hfe_desc, hfe_unit, hfe_quantity, hfe_unitprice, ") \
				_T(" case when hfe_type = 'X' then hfe_cost else hfe_diffcost end as diff_cost ") \
				_T(" from hms_fee ") \
				_T(" where hfe_docno = %ld ") \
				_T(" and hfe_type in ('X', 'O') and hfe_feegroup <> 'HITECH' and hfe_deptid = 'A6-C' and hfe_status in ('A','P') ") \
				_T(" and hfe_discount = 0) ") \
				_T(" where diff_cost > 0 "), nDocNo);

	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
		//tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
		//CDate::Convert(m_szToDate), yyyymmdd, ddmmyyyy);
		//rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	}
	while (!rs.IsEOF())
	{
		/*Group Seperation*/
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("hfe_desc")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("hfe_unit")));
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("hfe_quantity")));
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("hfe_unitprice")));
			rs.GetValue(_T("diff_cost"), nTmp);
			nTotal+= nTmp;
			rptDetail->SetValue(_T("6"), double2str(nTmp));
		}
		rs.MoveNext();
	}
	tmpStr.Format(_T("%Lf"), nTotal);
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);
	MoneyToString(tmpStr, szTemp);
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), szTemp + _T(" \x111\x1ED3ng."));
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
}

void CPrintForms::TM_OnPrintOutpatientTreatmentbyBicarbonat(long nDocNo, int nRefIdx){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CReportSection *rptHeader = NULL;
	CString szSQL, tmpStr, szCurrent_date;
	int nDay_count = 0;
	struct TREATMENT_INFO{
		int nDate[31];
	 	int nCount[31];
	};
	TREATMENT_INFO treatment_info;
	for (int i = 0; i < 31; i++)
	{
		treatment_info.nDate[i] = 0;
		treatment_info.nCount[i] = 0;
	}
	if (!rpt.Init(_T("Reports\\HMS\\PHIEUTHEODOIBENHNHANDIEUTRINGOAITRU.RPT")))
		return;
	//Patient Info
	rptHeader = rpt.GetReportHeader();
	szSQL.Format(_T(" SELECT hd_docno doc_no, get_patientname(hd_docno) patient_name,") \
				 _T("	hms_getaddress(hp_provid, hp_distid, hp_villid) address,") \
				 _T("	get_objectname(hd_object) object_name,") \
				 _T("	hd_cardno card_no, extract(year from hp_birthdate) yob") \
				 _T(" FROM hms_doc ") \
				 _T(" LEFT JOIN hms_patient ON (hd_patientno = hp_patientno) ") \
				 _T(" WHERE hd_docno = %ld"), nDocNo);
	int nRet = rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		if (nRet < 0)
			_msg(_T("%s"), szSQL);
		return;
	}
	if (rptHeader)
	{
		szCurrent_date = pMF->GetSysDate();
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		rptHeader->SetValue(_T("Department"), pMF->GetCurrentDepartmentName());
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), szCurrent_date.Mid(5, 2), szCurrent_date.Left(4));
		rptHeader->SetValue(_T("ReportDate"), tmpStr);
		rptHeader->SetValue(_T("PatientName"), rs.GetValue(_T("patient_name")));
		rptHeader->SetValue(_T("Yob"), rs.GetValue(_T("yob")));
		rptHeader->SetValue(_T("DocumentNo"), rs.GetValue(_T("doc_no")));
		rptHeader->SetValue(_T("Address"), rs.GetValue(_T("address")));
		rptHeader->SetValue(_T("ObjectName"), rs.GetValue(_T("object_name")));
		rptHeader->SetValue(_T("CardNo"), rs.GetValue(_T("card_no")));

	}
	//Treatment Info
	szSQL.Format(_T(" SELECT extract(day from ho_performdate) perform_date") \
				 _T(" FROM hms_treatment_record ") \
				 _T(" LEFT JOIN hms_operation ON (ho_docno = htr_docno AND ho_treattime = htr_treattime) ") \
				 _T(" WHERE ho_docno = %ld AND htr_treattime = %d AND htrf_acceptedfee IN ('Y', 'P')") \
				 _T(" AND ho_itemid = 'B40001699'") \
				 _T(" ORDER BY ho_performdate"), nDocNo, nRefIdx);
	nRet = rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		treatment_info.nDate[nDay_count] = ToInt(rs.GetValue(_T("perform_date")));
		treatment_info.nCount[nDay_count] = nDay_count+1;
		nDay_count++;
		rs.MoveNext();
	}
	for (int i = 0; i < 31; i++)
	{
		_debug(_T("%d"), treatment_info.nDate[i]);
		if (treatment_info.nDate[i] == 0)
			break;
		tmpStr.Format(_T("l%d"), i+1);
		rptHeader->SetValue(tmpStr, ToString(treatment_info.nDate[i]));
		tmpStr.Format(_T("%d"), i+1);
		rptHeader->SetValue(tmpStr, ToString(treatment_info.nCount[i]));
	}
	rpt.PrintPreview();
}

void CPrintForms::EM_OnPrintLaserReceiptList(long nDocNo){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CReportSection *rptHeader = NULL,* rptDetail = NULL;
	CString szSQL, tmpStr, szMoney, szDate;
	double nFeeAmount = 0;
	int nIdx = 1;
	if (!rpt.Init(_T("Reports\\HMS\\BANGKECHIPHIDIEUTRILASER.RPT")))
		return;
	szSQL.Format(_T(" SELECT   hd_docno doc_no, ") \
				_T("           upper(Get_patientname(hd_docno)) patient_name, ") \
				_T("           Extract(year FROM hp_birthdate) yob, ") \
				_T("           get_syssel_desc('sys_sex', hp_sex) sex, ") \
				_T("           Hms_getaddress(hp_provid, hp_distid, hp_villid) addr, ") \
				_T("           hd_diagnostic diagnostic ") \
				_T(" FROM      hms_doc ") \
				_T(" LEFT JOIN hms_patient ON ( hp_patientno = hd_patientno ) ") \
				_T(" WHERE     hd_docno = %ld"), nDocNo);
	rs.ExecSQL(szSQL);
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
		rptHeader->SetValue(_T("Department"), _T("Kho\x61 Kh\xE1m \x42\x1EC7nh \x43\x31.\x31"));
		tmpStr.Format(_T("%ld"), nDocNo);
		rptHeader->SetValue(_T("DocumentNo"), tmpStr);
		rptHeader->SetValue(_T("PatientName"), rs.GetValue(_T("patient_name")));
		rptHeader->SetValue(_T("yearofbirth"), rs.GetValue(_T("yob")));
		rptHeader->SetValue(_T("sex"), rs.GetValue(_T("sex")));
		rptHeader->SetValue(_T("PatientName"), rs.GetValue(_T("patient_name")));
		rptHeader->SetValue(_T("address"), rs.GetValue(_T("addr")));
		rptHeader->SetValue(_T("diagnostic"), rs.GetValue(_T("diagnostic")));
	}
	szSQL.Format(_T(" SELECT    hfe_desc descr, ") \
				_T("           hfe_unitprice unit_price, ") \
				_T("           hfe_quantity qty, ") \
				_T("           hfe_cost amount ") \
				_T(" FROM      hms_fee ") \
				_T(" LEFT JOIN hms_operation ON ( hfe_docno = ho_docno AND hfe_orderid = ho_idx AND hfe_type = 'O' ) ") \
				_T(" WHERE     hfe_class = 'E' AND ( hfe_type = 'E'  OR ( hfe_type = 'O' AND ho_pdeptid = 'C15' ) ) ") \
				_T(" AND hfe_status = 'O' AND hfe_docno = %ld ") \
				_T(" ORDER BY hfe_type"), nDocNo);
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("descr")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("unit_price")));
			rptDetail->SetValue(_T("4"), rs.GetValue(_T("qty")));
			rptDetail->SetValue(_T("5"), rs.GetValue(_T("amount")));
			nFeeAmount += str2double(rs.GetValue(_T("amount")));
		}
		rs.MoveNext();
	}
	rptHeader = rpt.AddDetail(rpt.GetGroupFooter(1));
	if (rptHeader)
		rptHeader->SetValue(_T("s5"), double2str(nFeeAmount));
	MoneyToString(ToString(nFeeAmount), tmpStr);
	tmpStr += _T(" \x111\x1ED3ng.");
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	szDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szDate.Mid(8, 2), szDate.Mid(5, 2), szDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
}