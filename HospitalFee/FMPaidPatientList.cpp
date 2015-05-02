#include "stdafx.h"
#include "FMPaidPatientList.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMPaidPatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMPaidPatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMPaidPatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMPaidPatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMPaidPatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMPaidPatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMPaidPatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMPaidPatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMPaidPatientList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMPaidPatientList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMPaidPatientList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMPaidPatientList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMPaidPatientList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMPaidPatientList *)pWnd)->OnClerkAddNew();
}*/
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	 ((CFMPaidPatientList*)pWnd)->OnOutPatientSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	 ((CFMPaidPatientList*)pWnd)->OnInPatientSelect();
}
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMPaidPatientList*)pWnd)->OnLockedSelect();
}
static void _OnWorkingDateSelectFnc(CWnd *pWnd){
	 ((CFMPaidPatientList*)pWnd)->OnWorkingDateSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMPaidPatientList *pVw = (CFMPaidPatientList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPrintInvoiceSelectFnc(CWnd *pWnd){
	CFMPaidPatientList *pVw = (CFMPaidPatientList *)pWnd;
	pVw->OnPrintInvoiceSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMPaidPatientList *pVw = (CFMPaidPatientList *)pWnd;
	pVw->OnExportSelect();
}
static void _OnCloseOutSelectFnc(CWnd *pWnd){
	CFMPaidPatientList *pVw = (CFMPaidPatientList *)pWnd;
	pVw->OnCloseOutSelect();
}
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CFMPaidPatientList*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CFMPaidPatientList*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMPaidPatientList*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMPaidPatientList*)pWnd)->OnListDeleteItem();
}
static int _OnAddFMPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMPaidPatientList*)pWnd)->OnAddFMPaidPatientList();
} 
static int _OnEditFMPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMPaidPatientList*)pWnd)->OnEditFMPaidPatientList();
} 
static int _OnDeleteFMPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMPaidPatientList*)pWnd)->OnDeleteFMPaidPatientList();
} 
static int _OnSaveFMPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMPaidPatientList*)pWnd)->OnSaveFMPaidPatientList();
} 
static int _OnCancelFMPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMPaidPatientList*)pWnd)->OnCancelFMPaidPatientList();
}


static int _OnListCheckAllFnc(CWnd *pWnd){
	return ((CFMPaidPatientList*)pWnd)->OnListCheckAll();
}

static int _OnListUncheckAllFnc(CWnd *pWnd){
	return ((CFMPaidPatientList*)pWnd)->OnListUncheckAll();
} 

CFMPaidPatientList::CFMPaidPatientList(CWnd *pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 520;
	SetDefaultValues();
}

CFMPaidPatientList::~CFMPaidPatientList(){
}
void CFMPaidPatientList::OnCreateComponents()
{
	/*m_wndPaidPatientList.Create(this, _T("Paid Patient List"), 5, 5, 490, 150);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 485, 85); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 285, 155, 385, 180);
	m_wndExport.Create(this, _T("&Export"), 390, 155, 490, 180);
	m_wndCloseOut.Create(this, _T("&Close out"), 5, 155, 105, 180);
	m_wndOutPatient.Create(this, _T("OutPatient"), 320, 90, 400, 115);
	m_wndInPatient.Create(this, _T("InPatient"), 405, 90, 485, 115);
	m_wndLocked.Create(this, _T("Locked"), 320, 120, 485, 145);*/

	m_wndPaidPatientList.Create(this, _T("Paid Patient List"), 5, 5, 430, 480);
	m_wndStockInfo.Create(this, _T("Stock Information"), 10, 90, 425, 415);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 30, 300, 55);
	m_wndToDate.Create(this,305, 30, 425, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 425, 85); 
	m_wndOutPatient.Create(this, _T("OutPatient"), 260, 420, 340, 445);
	m_wndInPatient.Create(this, _T("InPatient"), 345, 420, 425, 445);
	m_wndLocked.Create(this, _T("Locked"), 260, 450, 425, 475);
	m_wndCloseOut.Create(this, _T("&Close out"), 5, 485, 105, 510);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 225, 485, 325, 510);
	m_wndExport.Create(this, _T("&Export"), 330, 485, 430, 510);
	m_wndList.Create(this,15, 115, 420, 410); 
	m_wndCloseOut.ShowWindow(SW_HIDE);
}

void CFMPaidPatientList::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 270);

	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 90);
	m_wndList.InsertColumn(1, _T("Name"), CFMT_TEXT, 340);
	
	m_wndList.SetCheckBox(TRUE);
	m_wndList.ModifyStyle(0, LVS_NOSORTHEADER);

}
void CFMPaidPatientList::OnSetWindowEvents(){
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
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);
	m_wndWorkingDate.SetEvent(WE_CLICK, _OnWorkingDateSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndPrintInvoice.SetEvent(WE_CLICK, _OnPrintInvoiceSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndCloseOut.SetEvent(WE_CLICK, _OnCloseOutSelectFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);

	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);

	m_wndList.SetWindowText(_T("Stock List"));
	m_wndList.AddEvent(1, _T("Check All"), _OnListCheckAllFnc);
	m_wndList.AddEvent(2, _T("UnCheck All"), _OnListUncheckAllFnc);
	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	m_szClerkKey = pMF->GetCurrentUser();
	UpdateData(false);

	OnListLoadData();
}
void CFMPaidPatientList::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndOutPatient.GetDlgCtrlID(), m_bOutPatient);
	DDX_Check(pDX, m_wndInPatient.GetDlgCtrlID(), m_bInPatient);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);
	//DDX_Check(pDX, m_wndWorkingDate.GetDlgCtrlID(), m_bWorkingDate);
}
void CFMPaidPatientList::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMPaidPatientList::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMPaidPatientList::SetDefaultValues()
{

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bOutPatient=TRUE;
	m_bInPatient=FALSE;
	m_bLocked=FALSE;
	m_bWorkingDate=FALSE;
	m_nTotal = 0;

}
int CFMPaidPatientList::SetMode(int nMode){
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
		m_wndCloseOut.EnableWindow(FALSE);
 		UpdateData(FALSE); 
 		return nOldMode; 
}
/*void CFMPaidPatientList::OnFromDateChange(){
	
} */
/*void CFMPaidPatientList::OnFromDateSetfocus(){
	
} */
/*void CFMPaidPatientList::OnFromDateKillfocus(){
	
} */
int CFMPaidPatientList::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMPaidPatientList::OnToDateChange(){
	
} */
/*void CFMPaidPatientList::OnToDateSetfocus(){
	
} */
/*void CFMPaidPatientList::OnToDateKillfocus(){
	
} */
int CFMPaidPatientList::OnToDateCheckValue(){
	return 0;
} 
void CFMPaidPatientList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMPaidPatientList::OnClerkSelendok(){
	 
}
/*void CFMPaidPatientList::OnClerkSetfocus(){
	
}*/
/*void CFMPaidPatientList::OnClerkKillfocus(){
	
}*/
long CFMPaidPatientList::OnClerkLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();
	
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	}

	m_wndClerk.DeleteAllItems(); 
	szSQL.Format(_T("SELECT su_userid as userid, su_name as name ") \
		         _T("FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);

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
/*void CFMPaidPatientList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMPaidPatientList::OnOutPatientSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bInPatient = FALSE;
	UpdateData(FALSE);
}
void CFMPaidPatientList::OnInPatientSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bOutPatient = FALSE;
	UpdateData(FALSE);
}
void CFMPaidPatientList::OnLockedSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	//m_bLocked = TRUE;
	m_bWorkingDate = FALSE;
	UpdateData(FALSE);
}
void CFMPaidPatientList::OnWorkingDateSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	m_bLocked = FALSE;
	//m_bWorkingDate = TRUE;
	UpdateData(FALSE);
}

void CFMPaidPatientList::OnListDblClick(){
	
} 
void CFMPaidPatientList::OnListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMPaidPatientList::OnListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMPaidPatientList::OnListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndList.BeginLoad(); 
	int nCount = 0;

	szSQL.Format(_T("SELECT msl_storage_id as ID,") \
		         _T(" msl_name as Name") \
		         _T(" FROM m_storagelist") \
				 _T(" WHERE msl_type in('A','B') and msl_isactive='Y'") \
				 _T(" ORDER BY msl_storage_id"));

	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
}

void CFMPaidPatientList::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	UpdateData(true);

	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szSysDate;
	CString szTemp, szClass;

	int nIdx = 1;
	double nAmount = 0;
	long double nTotal = 0;
	long double nGroupTotal = 0;
	m_nTotal = 0;

	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);

	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/HF_DANHSACHBENHNHANNOPTIEN.rpt"), false, false, 2))
		return;

	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	CReportSection *rptDetail;
	CString szGroup[] = {_T("I"), _T("II")};
	int nIndex = 0;

	while (!rs.IsEOF())
	{
		if (szClass != rs.GetValue(_T("fclass")))
		{
			if (nGroupTotal > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
				TranslateString(_T("Total Amt"), tmpStr);
				tmpStr.AppendFormat(_T(" %s"), szGroup[nIndex++]);
				rptDetail->SetValue(_T("GroupName"), tmpStr);
				//tmpStr.Format(_T("%.2lf"), nGroupTotal);
				FormatCurrency(nGroupTotal, tmpStr);
				rptDetail->SetValue(_T("SumGroupName"), tmpStr);

				nTotal += nGroupTotal;
				nGroupTotal = 0;
			}

			nIdx = 1;

			rs.GetValue(_T("fclass"), szClass);
			if (szClass == _T("X"))
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
				rptDetail->SetValue(_T("GroupName"), _T("\x43\xE1\x63 kho\x61 n\x1ED9i tr\xFA"));
			}
		}

		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIdx++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		if (!rs.GetValue(_T("roomid")).IsEmpty())
			tmpStr.Format(_T("%s - %s"), rs.GetValue(_T("deptid")), rs.GetValue(_T("roomid")));
		else
			tmpStr.Format(_T("%s"), rs.GetValue(_T("deptid")));

		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("amount"), nAmount);
		nGroupTotal += nAmount;
		FormatCurrency(nAmount, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.MoveNext();
	}

	if (nGroupTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
		TranslateString(_T("Total Amt"), tmpStr);
		tmpStr.AppendFormat(_T(" %s"), szGroup[nIndex++]);
		rptDetail->SetValue(_T("GroupName"), tmpStr);
		//tmpStr.Format(_T("%.2lf"), nGroupTotal);
		FormatCurrency(nGroupTotal, tmpStr);
		rptDetail->SetValue(_T("SumGroupName"), tmpStr);

		nTotal += nGroupTotal;
		nGroupTotal = 0;
	}

	//rptDetail = rpt.AddDetail(rpt.GetReportFooter());
	tmpStr.Format(_T("%s"), _T("T\x1ED5ng \x63\x1ED9ng"));
	rpt.GetReportFooter()->SetValue(_T("TotalAmount"), tmpStr);

	nTotal += 0.5;
	long double nTemp = floor(nTotal);

	m_nTotal = nTemp;

	FormatCurrency(nTemp, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("SumTotalAmount"), tmpStr);

	tmpStr.Format(_T("%.0lf"), nTemp);
	CString szMoney = tmpStr;
	MoneyToString(szMoney, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));

	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	if (!m_szClerkKey.IsEmpty())
		rpt.GetReportFooter()->SetValue(_T("ReceiverBy"), m_wndClerk.GetCurrent(1));
	rpt.PrintPreview();

	if (!m_szClerkKey.IsEmpty() && m_bLocked && m_nTotal > 0)
	{
		int ret = OnPrintInvoice();
	}
} 
void CFMPaidPatientList::OnPrintInvoiceSelect()
{
	
	
} 
void CFMPaidPatientList::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0, nIdx = 0;
	double nAmount = 0;
	long double nTotal = 0;
	szSQL = GetQueryString();
	//MessageBox(szSQL);
	rs.ExecSQL(szSQL);

	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}

	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 15);

	xls.SetRowHeight(4, 40);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);

	xls.SetCellMergedColumns(nCol + 3, nRow, 4);
	xls.SetCellMergedColumns(nCol + 3, nRow + 1, 4);

	xls.SetCellMergedColumns(nCol, nRow + 2, 5);
	xls.SetCellMergedColumns(nCol, nRow + 3, 5);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 3, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 3, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol, nRow + 2, _T("\x42\x1EA2NG TH\x1ED0NG K\xCA \x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N N\x1ED8P TI\x1EC0N"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn \x62\x1EC7nh nh\xE2n"));
	arrCol.Add(_T("S\x1ED1 h\x1ED3 s\x1A1"));
	arrCol.Add(_T("Ph\xF2ng kh\xE1m"));
	arrCol.Add(_T("S\x1ED1 ti\x1EC1n"));

	nRow = 4;

	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol + i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10); 
	}
	nRow = 5;

	while (!rs.IsEOF())
	{
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		if (!rs.GetValue(_T("roomid")).IsEmpty())
			tmpStr.Format(_T("%s - %s"), rs.GetValue(_T("deptid")), rs.GetValue(_T("roomid")));
		else
			tmpStr.Format(_T("%s"), rs.GetValue(_T("deptid")));

		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("amount"), nAmount);
		tmpStr.Format(_T("%.2f"), nAmount);
		nTotal += nAmount;
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);
		nRow++;
		rs.MoveNext();
	}
	
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 10);
	tmpStr.Format(_T("%.2Lf"), nTotal);
	xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);

	xls.Save(_T("Exports\\Danh Sach Benh Nhan Nop Tien.xls"));
	
} 
int CFMPaidPatientList::OnAddFMPaidPatientList(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}

#include "HMSInvoiceDateSettingDialog.h"

void CFMPaidPatientList::OnCloseOutSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
	CRecord rs(&pMF->m_db);
	CString szSQL, szReceiptDate;
	CString tmpStr;
	
	UpdateData(TRUE);
	if(m_szClerkKey.IsEmpty())
	{
		ShowMessageBox(_T("You much select user name"));
		return;
	}
	if(m_szClerkKey != pMF->GetCurrentUser())
	{
		ShowMessageBox(_T("Cannot close out of other user"));
		return;
	}
	szSQL.Format(_T("SELECT Cast_Timestamp2Date(hfc_lastlockeddate) AS lockeddate FROM hms_fee_control where hfc_userid='%s' "), pMF->GetCurrentUser());
	rs.ExecSQL(szSQL);
	if(rs.GetIntValue() > 0)
	{
		rs.GetValue(_T("lockeddate"), tmpStr);
		if (CompareDate(tmpStr, pMF->GetSysDate()) == 0)
		{
			ShowMessageBox(_T("Cannot perform close out two time in day"));
			return;
		}
	}

	szReceiptDate = pMF->GetSysDate();
	szSQL.Format(_T("HMS_FEE_INVOICE_LOCK('%s', '%s', '%s', %ld, '%s') "),
		pMF->GetCurrentUser(), szReceiptDate, _T(""), 0, pMF->GetSysDate());
	int res = str2long(pMF->ExecDML(szSQL));
	CString szMsg;
	szMsg.Format(_T("Close out [%ld] receipt"), res);
	ShowMessageBox(szMsg);
	CHMSInvoiceDateSettingDialog dlg(this);
	if (dlg.DoModal() == IDOK)
	{
		CString szRecvDate = dlg.m_szWorkingDate;
		CString szStatusLabel, szStatus;
		TranslateString(_T("UserID: %s - Working Date: %s"), szStatusLabel);
		szStatus.Format(szStatusLabel, pMF->GetCurrentUser(), CDate::Convert(szRecvDate));
		szStatus.AppendFormat(_T(" - Server:%s"), pMF->GetHost());
		pMF->SetStatusText(szStatus, 1);
	}

}
int CFMPaidPatientList::OnEditFMPaidPatientList(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMPaidPatientList::OnDeleteFMPaidPatientList(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
 	CString szSQL; 
 	if(ShowMessage(1, MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2) == IDNO) 
 		return -1; 
 	szSQL.Format(_T("DELETE FROM  WHERE  AND") ); 
 	int ret = pMF->ExecSQL(szSQL); 
 	if(ret >= 0){ 
 		SetMode(VM_NONE); 
 		OnCancelFMPaidPatientList();
 	}
	return 0;
}
int CFMPaidPatientList::OnSaveFMPaidPatientList(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT) 
 		return -1; 
 	if(!IsValidateData()) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
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
 		//OnFMPaidPatientListListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMPaidPatientList::OnCancelFMPaidPatientList(){
 	if(GetMode() == VM_EDIT) 
 	{ 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 		SetMode(VM_NONE); 
 	} 
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
 	return 0;
} 
int CFMPaidPatientList::OnFMPaidPatientListListLoadData(){
	return 0;
}

CString CFMPaidPatientList::GetQueryString()
{
	CString szSQL, szWhere;
	CString szStocks;
	CString szDrugCondition;

	szWhere.Empty();
	szStocks.Empty();
	szDrugCondition.Empty();
	
	for (int i = 0; i < m_wndList.GetItemCount();i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (!szStocks.IsEmpty())
				szStocks += _T(", ");
			szStocks += m_wndList.GetItemText(i, 0);
		}
	}
	if (!m_szClerkKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_staff='%s' "), m_szClerkKey);
	}

	if (m_bLocked)
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_locked='Y' "));
		szWhere.AppendFormat(_T(" AND fi.hfe_lockeddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}
	else
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_locked<>'Y' "));
		//szWhere.AppendFormat(_T(" AND fi.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
		//	                 m_szFromDate, m_szToDate);
	}

	if (m_bOutPatient)
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_class IN('E','X') "));
	}
	else
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_class NOT IN('E','X') "));
	}
	if (!szStocks.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpo_storage_id IN (%s)"), szStocks);
	szSQL.Format(_T(" SELECT docno,") \
				_T("		 get_patientname(docno) pname,") \
				_T("		 deptid,") \
				_T("		 SUM(amount) amount") \
				_T(" FROM") \
				_T(" (") \
				_T("  SELECT fi.hfe_docno as docno,") \
				_T("         fi.hfe_deptid AS deptid,") \
				_T("         fe.hfe_patpaid AS amount") \
				_T("  FROM hms_fee_invoice fi") \
				_T("  LEFT JOIN hms_fee fe ON (fe.hfe_docno = fi.hfe_docno AND fe.hfe_invoiceno = fi.hfe_invoiceno)") \
				_T("  LEFT JOIN hms_pharmaorder_view ON (fe.hfe_docno = hpo_docno AND hfe_orderid = hpo_orderid)") \
				_T("  %s") \
				_T("  WHERE fe.hfe_patpaid > 0 AND fi.hfe_status='P' AND fe.hfe_status IN ('P', 'R') %s") \
				_T(" ) Tbl") \
				_T("  GROUP BY docno, deptid") \
				_T(" ORDER BY docno"),	szDrugCondition, szWhere);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}

int CFMPaidPatientList::OnPrintInvoice()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	UpdateData(true);

	CReport rpt;
	CRecord rs(&pMF->m_db), ors(&pMF->m_db);
	CRecord rss(&pMF->m_db);

	CString szSQL, tmpStr, szSysDate, szTemp;

	int nIdx = 0;
	double nAmount = 0;
	long double nTotal = 0;

	//szSQL = GetQueryString();
	//int nCount = rs.ExecSQL(szSQL);

	/*if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return -1;
	}*/

	if (!rpt.Init(_T("Reports/HMS/HF_PHIEUTHUC30BB.rpt"), false, false, 3))
		return 0;

	CReportSection *rptDetail;

	rptDetail = rpt.AddDetail();

	rptDetail->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rptDetail->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);

	tmpStr.Format(rptDetail->GetValue(_T("ReportDate")),
		         CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				 CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rptDetail->SetValue(_T("ReportDate"), tmpStr);

	if (!m_szClerkKey.IsEmpty())
	{
		szSQL.Format(_T(" SELECT su_name AS staff,") \
				_T("        sd_name AS deptname") \
				_T(" FROM sys_user") \
				_T(" LEFT JOIN sys_dept ON(sd_id=su_deptid)") \
				_T(" WHERE su_userid='%s'"), m_szClerkKey);

		rs.ExecSQL(szSQL);

		if (!rs.IsEOF())
		{
			rs.GetValue(_T("deptname"), tmpStr);
			rptDetail->SetValue(_T("address"), tmpStr);

			rs.GetValue(_T("staff"), tmpStr);
			rptDetail->SetValue(_T("ReceiveBy"), tmpStr);
			rptDetail->SetValue(_T("patientname"), tmpStr);
		}
	}
	/*else
	{
		rptDetail->SetValue(_T("address"), pMF->GetDepartmentName(pMF->GetCurrentDepartmentID()));
		tmpStr = pMF->GetCurrentUser();
	}*/

	/*while (!rs.IsEOF())
	{
		rs.GetValue(_T("amount"), nAmount);
		nTotal += nAmount;
		rs.MoveNext();
	}*/

	//rptDetail->SetValue(_T("patientname"), tmpStr);

	tmpStr = _T("Thu ti\x1EC1n vi\x1EC7n ph\xED");
	szSQL.Format(_T("SELECT count(*) FROM sys_userperm WHERE sup_moduleid = 'FM' AND sup_userid = '%s' AND (sup_permid = '01.15' OR sup_permid = '01.16')"), m_szClerkKey);
	ors.ExecSQL(szSQL);

	if (ors.GetIntValue() > 0)
		tmpStr = _T("Thu ti\x1EC1n vi\x1EC7n ph\xED");
	else 
	{
		szSQL.Format(_T(" select sup_permid as permid") \
					_T(" from sys_userperm") \
					_T(" where sup_permid in('01.01.E','01.01.P','01.01.O','01.01.D','01.01.X')") \
					_T(" and sup_userid='%s'"), m_szClerkKey);
		ors.ExecSQL(szSQL);
		if (ors.GetRecordCount() == 1)
		{
			szSQL.Format(_T(" select sp_name as itemname") \
						_T(" from sys_permission") \
						_T(" where sp_moduleid='FM' and sp_id='%s'"),
						ors.GetValue(_T("permid")));
			rss.ExecSQL(szSQL);

			if (!rss.IsEOF())
				rss.GetValue(_T("itemname"), tmpStr);
			
		}
		else
			tmpStr = _T("\x43\xE1\x63 kho\x1EA3n vi\x1EC7n ph\xED ph\x1EA3i thu");
	}
	rptDetail->SetValue(_T("reason"), tmpStr);
	
	//nTotal += 0.5;
	//long double nTemp = floor(nTotal);

	FormatCurrency(m_nTotal, tmpStr);
	rptDetail->SetValue(_T("TotalAmount"), tmpStr);

	szTemp.Format(_T("%.0Lf"), m_nTotal);
	MoneyToString(szTemp, tmpStr);

	rptDetail->SetValue(_T("SumInWord"), tmpStr);

	CDateTime dt;
	dt.ParseDateTime(m_szToDate);
	CDate date = dt.GetDate();
	szSysDate = date.GetDate();

	tmpStr.Format(rptDetail->GetValue(_T("ReportDate")),
		          szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));

	rptDetail->SetValue(_T("PrintDate"), tmpStr);

	rpt.PrintPreview();

	return 1;
}


int CFMPaidPatientList::OnListCheckAll()
{
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (!m_wndList.GetCheck(i))
		{
			m_wndList.SetCheck(i, TRUE);
		}
	}
	return 0;
}

int CFMPaidPatientList::OnListUncheckAll()
{
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			m_wndList.SetCheck(i, FALSE);
		}
	}
	return 0;
}