#include "stdafx.h"
#include "FMAdvancePaymentList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMAdvancePaymentList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMAdvancePaymentList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMAdvancePaymentList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMAdvancePaymentList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMAdvancePaymentList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMAdvancePaymentList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMAdvancePaymentList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMAdvancePaymentList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMAdvancePaymentList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMAdvancePaymentList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMAdvancePaymentList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMAdvancePaymentList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMAdvancePaymentList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMAdvancePaymentList *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMAdvancePaymentList *pVw = (CFMAdvancePaymentList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMAdvancePaymentList *pVw = (CFMAdvancePaymentList *)pWnd;
	pVw->OnExportSelect();
}
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	  ((CFMAdvancePaymentList*)pWnd)->OnOutPatientSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	  ((CFMAdvancePaymentList*)pWnd)->OnInPatientSelect();
}
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMAdvancePaymentList*)pWnd)->OnLockedSelect();
}
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CFMAdvancePaymentList*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CFMAdvancePaymentList*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMAdvancePaymentList*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMAdvancePaymentList*)pWnd)->OnListDeleteItem();
}
static int _OnAddFMAdvancePaymentListFnc(CWnd *pWnd){
	 return ((CFMAdvancePaymentList*)pWnd)->OnAddFMAdvancePaymentList();
} 
static int _OnEditFMAdvancePaymentListFnc(CWnd *pWnd){
	 return ((CFMAdvancePaymentList*)pWnd)->OnEditFMAdvancePaymentList();
} 
static int _OnDeleteFMAdvancePaymentListFnc(CWnd *pWnd){
	 return ((CFMAdvancePaymentList*)pWnd)->OnDeleteFMAdvancePaymentList();
} 
static int _OnSaveFMAdvancePaymentListFnc(CWnd *pWnd){
	 return ((CFMAdvancePaymentList*)pWnd)->OnSaveFMAdvancePaymentList();
} 
static int _OnCancelFMAdvancePaymentListFnc(CWnd *pWnd){
	 return ((CFMAdvancePaymentList*)pWnd)->OnCancelFMAdvancePaymentList();
}

static int _OnListCheckAllFnc(CWnd *pWnd){
	return ((CFMAdvancePaymentList*)pWnd)->OnListCheckAll();
}
static int _OnListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMAdvancePaymentList*)pWnd)->OnListUnCheckAll();
}

CFMAdvancePaymentList::CFMAdvancePaymentList(CWnd *pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 505;
	SetDefaultValues();
}
CFMAdvancePaymentList::~CFMAdvancePaymentList(){
}
void CFMAdvancePaymentList::OnCreateComponents()
{
	/*m_wndAdvancePaymentList.Create(this, _T("Advance Payment List"), 5, 4, 490, 435);
	m_wndObjectInfo.Create(this, _T("Object"), 10, 90, 485, 400);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 485, 85); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 285, 440, 385, 465);
	m_wndExport.Create(this, _T("&Export"), 390, 440, 490, 465);
	m_wndList.Create(this,15, 115, 480, 395); 
	m_wndLocked.Create(this, _T("Locked"), 375, 405, 485, 430);*/
	m_wndAdvancePaymentList.Create(this, _T("Report Condition"), 5, 5, 430, 465);
	m_wndObjectInfo.Create(this, _T("Object"), 10, 90, 425, 400);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 30, 300, 55);
	m_wndToDate.Create(this,305, 30, 425, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 425, 85); 
	m_wndOutPatient.Create(this, _T("OutPatient"), 260, 405, 340, 430);
	m_wndInPatient.Create(this, _T("InPatient"), 345, 405, 425, 430);
	m_wndLocked.Create(this, _T("Locked"), 345, 435, 425, 460);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 225, 470, 325, 495);
	m_wndExport.Create(this, _T("&Export"), 330, 470, 430, 495);
	m_wndList.Create(this,15, 115, 420, 395); 
}

void CFMAdvancePaymentList::OnInitializeComponents()
{
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

	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	m_wndList.InsertColumn(1, _T("Desc"), CFMT_TEXT, 350);
	m_wndList.SetCheckBox(TRUE);
}
void CFMAdvancePaymentList::OnSetWindowEvents()
{
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
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);

	m_wndList.SetWindowText(_T("Object"));
	m_wndList.AddEvent(1, _T("Check All"), _OnListCheckAllFnc);
	m_wndList.AddEvent(2, _T("UnCheck All"), _OnListUnCheckAllFnc);
	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

	OnListLoadData();
}
void CFMAdvancePaymentList::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Radio(pDX, m_wndOutPatient.GetDlgCtrlID(), m_nOutPatient);
	DDX_Radio(pDX, m_wndInPatient.GetDlgCtrlID(), m_nInPatient);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);
}
void CFMAdvancePaymentList::GetDataToScreen()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMAdvancePaymentList::GetScreenToData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMAdvancePaymentList::SetDefaultValues()
{

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_nOutPatient=1;
	m_nInPatient=0;
	m_bLocked=FALSE;

}
int CFMAdvancePaymentList::SetMode(int nMode)
{
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
/*void CFMAdvancePaymentList::OnFromDateChange(){
	
} */
/*void CFMAdvancePaymentList::OnFromDateSetfocus(){
	
} */
/*void CFMAdvancePaymentList::OnFromDateKillfocus(){
	
} */
int CFMAdvancePaymentList::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMAdvancePaymentList::OnToDateChange(){
	
} */
/*void CFMAdvancePaymentList::OnToDateSetfocus(){
	
} */
/*void CFMAdvancePaymentList::OnToDateKillfocus(){
	
} */
int CFMAdvancePaymentList::OnToDateCheckValue(){
	return 0;
} 
void CFMAdvancePaymentList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMAdvancePaymentList::OnClerkSelendok(){
	 
}
/*void CFMAdvancePaymentList::OnClerkSetfocus(){
	
}*/
/*void CFMAdvancePaymentList::OnClerkKillfocus(){
	
}*/
long CFMAdvancePaymentList::OnClerkLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if (m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
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
/*void CFMAdvancePaymentList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMAdvancePaymentList::OnListDblClick(){
	
} 
void CFMAdvancePaymentList::OnListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMAdvancePaymentList::OnListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMAdvancePaymentList::OnListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndList.BeginLoad(); 
	m_wndList.DeleteAllItems(); 
	int nCount = 0;

	szSQL.Format(_T("SELECT ho_id AS ID, ho_desc AS Description ") \
		         _T("FROM hms_object ") \
				 _T("ORDER BY cast(ho_id as integer)"));

	nCount = rs.ExecSQL(szSQL);

	while(!rs.IsEOF())
	{ 
		m_wndList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Description")), NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
}
void CFMAdvancePaymentList::OnOutPatientSelect()
{
}
void CFMAdvancePaymentList::OnInPatientSelect()
{
}
void CFMAdvancePaymentList::OnLockedSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();	
}
void CFMAdvancePaymentList::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);

	CReport rpt;
	CString szSQL, tmpStr, szSysDate;
	CString szObject;

	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	double nAmount = 0;
	long double nGroupTotal = 0;
	long double nTotal = 0;
	int nCount = rs.ExecSQL(szSQL);

	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/HF_ADVANCEPAYMENTPATIENTLIST.RPT")))
		return;

	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	CReportSection *rptDetail;

	while (!rs.IsEOF())
	{
		if (rs.GetValue(_T("objid")) != szObject)
		{
			if (nGroupTotal > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				TranslateString(_T("Community"), tmpStr);
				rptDetail->SetValue(_T("SumAmountLabel"), tmpStr);
				FormatCurrency(nGroupTotal, tmpStr);
				rptDetail->SetValue(_T("SumAmount"), tmpStr);
				nTotal += nGroupTotal;
				nGroupTotal = 0;
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rs.GetValue(_T("objname"), tmpStr);
			rptDetail->SetValue(_T("GroupName"), tmpStr);
			rs.GetValue(_T("objid"), szObject);
		}

		rptDetail = rpt.AddDetail();

		tmpStr = CDate::Convert(rs.GetValue(_T("receiptdate")), yyyymmdd, ddmmyyyy);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("receiptno"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("recordno"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("age"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("sex"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("deptid"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("amount"), nAmount);
		nGroupTotal += nAmount;
		FormatCurrency(nAmount, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("clerk"), tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);
		rs.MoveNext();
	}

	if (nGroupTotal > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		TranslateString(_T("Community"), tmpStr);
		rptDetail->SetValue(_T("SumAmountLabel"), tmpStr);
		FormatCurrency(nGroupTotal, tmpStr);
		rptDetail->SetValue(_T("SumAmount"), tmpStr);
		nTotal += nGroupTotal;
		//nGroupTotal = 0;
	}

	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	FormatCurrency(nTotal, tmpStr);
	rptDetail->SetValue(_T("SumAmount"), tmpStr);
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CFMAdvancePaymentList::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	CString szObject;

	int nCol = 0, nRow = 0;
	double nAmount = 0;
	long double nGroupTotal = 0;
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
	xls.SetColumnWidth(0, 10);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 9);
	xls.SetColumnWidth(3, 6);

	xls.SetColumnWidth(4, 21);
	xls.SetColumnWidth(5, 5);
	xls.SetColumnWidth(6, 5);
	xls.SetColumnWidth(7, 20);
	xls.SetColumnWidth(8, 12);
	xls.SetColumnWidth(9, 15);

	xls.SetRowHeight(4, 40);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellMergedColumns(nCol, nRow + 1, 4);

	xls.SetCellMergedColumns(nCol + 5, nRow, 5);
	xls.SetCellMergedColumns(nCol + 5, nRow + 1, 5);

	xls.SetCellMergedColumns(nCol, nRow + 2, 10);
	xls.SetCellMergedColumns(nCol, nRow + 3, 10);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 5, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 5, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol, nRow + 2, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N T\x1EA0M \x1EE8NG"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("Ng\xE0y thu"));
	arrCol.Add(_T("Phi\x1EBFu thu"));
	arrCol.Add(_T("S\x1ED1 h\x1ED3 s\x1A1"));
	arrCol.Add(_T("S\x1ED1 \x62\x1EC7nh \xE1n"));
	arrCol.Add(_T("H\x1ECD v\xE0 t\xEAn"));
	arrCol.Add(_T("Tu\x1ED5i"));
	arrCol.Add(_T("Gi\x1EDBi"));
	arrCol.Add(_T("Kho\x61/ Ph\xF2ng"));
	arrCol.Add(_T("S\x1ED1 ti\x1EC1n"));
	arrCol.Add(_T("Ng\x1B0\x1EDDi thu"));

	nRow = 4;

	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol + i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10); 
	}

	while (!rs.IsEOF())
	{
		if (rs.GetValue(_T("objid")) != szObject)
		{
			nRow++;
			if (nGroupTotal > 0)
			{
				xls.SetCellMergedColumns(nCol, nRow, 8);
				TranslateString(_T("Community"), tmpStr);
				xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
				tmpStr.Format(_T("%.2lf"), nGroupTotal);
				xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);

				nTotal += nGroupTotal;
				nGroupTotal = 0;
				nRow++;
			}

			xls.SetCellMergedColumns(nCol, nRow, 10);
			rs.GetValue(_T("objname"), tmpStr);
			xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT, true);
			rs.GetValue(_T("objid"), szObject);
		}

		nRow++;
		tmpStr = CDate::Convert(rs.GetValue(_T("receiptdate")), yyyymmdd, ddmmyyyy);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_DATE | FMT_WRAPING);

		rs.GetValue(_T("receiptno"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("recordno"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("age"), tmpStr);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("sex"), tmpStr);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("dept"), tmpStr);
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("amount"), nAmount);
		nGroupTotal += nAmount;
		//FormatCurrency(nAmount, tmpStr);
		tmpStr.Format(_T("%.2f"), nAmount);
		xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("clerk"), tmpStr);
		xls.SetCellText(nCol + 9, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.MoveNext();
	}
	//FormatCurrency(nTotal, tmpStr);
	
	if (nGroupTotal > 0)
	{
		nRow++;
		xls.SetCellMergedColumns(nCol, nRow, 8);
		TranslateString(_T("Community"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
		tmpStr.Format(_T("%.2lf"), nGroupTotal);
		xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);

		nTotal += nGroupTotal;
	}

	nRow++;
	xls.SetCellMergedColumns(nCol, nRow, 8);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 10);
	tmpStr.Format(_T("%.2lf"), nTotal);
	xls.SetCellText(nCol + 8, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
	xls.Save(_T("Exports\\Danh Sach Benh Nhan Tam Ung.xls"));
	
} 
int CFMAdvancePaymentList::OnAddFMAdvancePaymentList(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMAdvancePaymentList::OnEditFMAdvancePaymentList(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMAdvancePaymentList::OnDeleteFMAdvancePaymentList(){
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
 		OnCancelFMAdvancePaymentList(); 
 	}
	return 0;
}
int CFMAdvancePaymentList::OnSaveFMAdvancePaymentList(){
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
 		//OnFMAdvancePaymentListListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMAdvancePaymentList::OnCancelFMAdvancePaymentList(){
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
int CFMAdvancePaymentList::OnFMAdvancePaymentListListLoadData(){
	return 0;
}
CString CFMAdvancePaymentList::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	CString szObject;

	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_staff='%s' "), m_szClerkKey);

	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (!szObject.IsEmpty())
				szObject += _T(",");
			szObject.AppendFormat(_T("'%s'"), m_wndList.GetItemText(i, 0));
		}
	}

	if (!szObject.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_objectid IN(%s) "), szObject);

	if (m_nOutPatient == 0)
	{
		szWhere.AppendFormat(_T(" AND hfe_class IN('E') "));
	}
	else
	{
		szWhere.AppendFormat(_T(" AND hfe_class NOT IN('E') "));
	}

	if (!m_bLocked)
	{
		szWhere.AppendFormat(_T(" AND hfe_locked<>'Y' "));
		szWhere.AppendFormat(_T(" AND hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}
	else
	{
		szWhere.AppendFormat(_T(" AND hfe_locked='Y' "));
		szWhere.AppendFormat(_T(" AND hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}

	/*szSQL.Format(_T(" SELECT hd_docno as docno,") \
		        _T("        hcr_recordno as recordno,") \
				_T("        hfe_deptid as deptid,") \
				_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("        hms_getage(hd_admitdate, hp_birthdate) as age,") \
				_T("        (select distinct ss_desc from sys_sel where ss_id='sys_sex' and ss_code=hp_sex) as sex,") \
				_T("        (select distinct sd_name from sys_dept where sd_id=hfe_deptid) as dept,") \
				_T("        (select distinct su_name from sys_user where su_userid=hfe_staff) as clerk,") \
				_T("        case when length(hfe_serialno) > 0 then hfe_serialno||'-'||cast(hfe_receiptno as varchar2(38)) ") \
				_T("        else cast(hfe_receiptno as varchar2(38)) end as receiptno,") \
				_T("        Cast_Timestamp2Date(hfe_date) as receiptdate,") \
				_T("        sum(hfe_amount) as amount,") \
				_T("        ho_id as objid,") \
				_T("        ho_desc as objname,") \
				_T("        case when ho_type in('I','C') then 1 else 2 end as obj") \
				_T(" FROM hms_patient") \
				_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
				_T(" LEFT JOIN hms_fee_deposit ON(hfe_docno=hd_docno)") \
				_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
				_T(" LEFT JOIN hms_object ON(ho_id=hfe_objectid)") \
				_T(" WHERE hfe_status='P' AND hd_object IS NOT NULL AND hd_object<>'0' %s") \
				_T(" GROUP BY hfe_docno, hcr_recordno, hp_surname, hp_midname, hp_firstname, hfe_serialno, hfe_receiptno, hp_birthdate,") \
				_T("          hd_admitdate, Cast_Timestamp2Date(hfe_date), ho_id, ho_desc, ho_type, hp_sex, hfe_deptid, hfe_staff") \
				_T(" ORDER BY obj, cast(ho_id as integer), receiptdate, pname"),
				szWhere);*/

	szSQL.Format(_T(" SELECT hd_docno as docno,") \
				_T("         hcr_recordno as recordno,") \
				_T("         hfe_deptid as deptid,") \
				_T("         trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("         hms_getage(hd_admitdate, hp_birthdate) as age,") \
				_T("         (select distinct ss_desc from sys_sel where ss_id='sys_sex' and ss_code=hp_sex) as sex,") \
				_T("         (select distinct sd_name from sys_dept where sd_id=hfe_deptid) as dept,") \
				_T("         (select distinct su_name from sys_user where su_userid=hfe_staff) as clerk,") \
				_T("         case when length(hfe_serialno) > 0 then hfe_serialno||'-'||cast(hfe_receiptno as varchar2(38))") \
				_T("         else cast(hfe_receiptno as varchar2(38)) end as receiptno,") \
				_T("         Cast_Timestamp2Date(hfe_date) as receiptdate,") \
				_T("         sum(hfe_amount) as amount,") \
				_T("         hfe_objectid as objid,") \
				_T("         ho_desc as objname,") \
				_T("         case when ho_type in('I','C') then 1 else 2 end as obj") \
				_T("  FROM hms_patient") \
				_T("  LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
				_T("  LEFT JOIN hms_fee_deposit ON(hfe_docno=hd_docno)") \
				_T("  LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
				_T("  LEFT JOIN hms_object ON(ho_id=hfe_objectid)") \
				_T("  WHERE hfe_status='P' AND hd_object IS NOT NULL AND hd_object<>'0' %s") \
				_T("  GROUP BY hd_docno, hcr_recordno, hp_surname, hp_midname, hp_firstname, hfe_serialno, hfe_receiptno, hp_birthdate,") \
				_T("           hd_admitdate, Cast_Timestamp2Date(hfe_date), hfe_objectid, ho_desc, ho_type, hp_sex, hfe_deptid, hfe_staff") \
				_T("  ORDER BY obj, cast(hfe_objectid as integer), receiptdate, pname"), szWhere);

	return szSQL;
}

int CFMAdvancePaymentList::OnListCheckAll()
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

int CFMAdvancePaymentList::OnListUnCheckAll()
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