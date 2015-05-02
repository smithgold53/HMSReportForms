#include "stdafx.h"
#include "FMDepositPatientList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"


/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMDepositPatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMDepositPatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMDepositPatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMDepositPatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMDepositPatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMDepositPatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMDepositPatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMDepositPatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMDepositPatientList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMDepositPatientList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMDepositPatientList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMDepositPatientList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMDepositPatientList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMDepositPatientList *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMDepositPatientList *pVw = (CFMDepositPatientList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMDepositPatientList *pVw = (CFMDepositPatientList *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CFMDepositPatientList*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CFMDepositPatientList*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMDepositPatientList*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMDepositPatientList*)pWnd)->OnListDeleteItem();
}
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	  ((CFMDepositPatientList*)pWnd)->OnOutPatientSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	  ((CFMDepositPatientList*)pWnd)->OnInPatientSelect();
}
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMDepositPatientList*)pWnd)->OnLockedSelect();
}
static int _OnAddFMDepositPatientListFnc(CWnd *pWnd){
	 return ((CFMDepositPatientList*)pWnd)->OnAddFMDepositPatientList();
} 
static int _OnEditFMDepositPatientListFnc(CWnd *pWnd){
	 return ((CFMDepositPatientList*)pWnd)->OnEditFMDepositPatientList();
} 
static int _OnDeleteFMDepositPatientListFnc(CWnd *pWnd){
	 return ((CFMDepositPatientList*)pWnd)->OnDeleteFMDepositPatientList();
} 
static int _OnSaveFMDepositPatientListFnc(CWnd *pWnd){
	 return ((CFMDepositPatientList*)pWnd)->OnSaveFMDepositPatientList();
} 
static int _OnCancelFMDepositPatientListFnc(CWnd *pWnd){
	 return ((CFMDepositPatientList*)pWnd)->OnCancelFMDepositPatientList();
}

static int _OnListCheckAllFnc(CWnd *pWnd){
	return ((CFMDepositPatientList*)pWnd)->OnListCheckAll();
}
static int _OnListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMDepositPatientList*)pWnd)->OnListUnCheckAll();
}

CFMDepositPatientList::CFMDepositPatientList(CWnd* pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 505;
	SetDefaultValues();
}
CFMDepositPatientList::~CFMDepositPatientList(){
}
void CFMDepositPatientList::OnCreateComponents()
{
	/*m_wndReportConditionInfo.Create(this, _T("Report Condition"), 5, 5, 490, 435);
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

	m_wndReportConditionInfo.Create(this, _T("Report Condition"), 5, 5, 430, 465);
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
void CFMDepositPatientList::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//EnableControls(FALSE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(65);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 270);

	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	m_wndList.InsertColumn(1, _T("Desc"), CFMT_TEXT, 350);
	m_wndList.SetCheckBox(TRUE);
}
void CFMDepositPatientList::OnSetWindowEvents()
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
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);
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
void CFMDepositPatientList::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Radio(pDX, m_wndOutPatient.GetDlgCtrlID(), m_nOutPatient);
	DDX_Radio(pDX, m_wndInPatient.GetDlgCtrlID(), m_nInPatient);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);
}
void CFMDepositPatientList::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMDepositPatientList::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMDepositPatientList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_nOutPatient=1;
	m_nInPatient=0;
	m_bLocked=FALSE;

}
int CFMDepositPatientList::SetMode(int nMode){
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
/*void CFMDepositPatientList::OnFromDateChange(){
	
} */
/*void CFMDepositPatientList::OnFromDateSetfocus(){
	
} */
/*void CFMDepositPatientList::OnFromDateKillfocus(){
	
} */
int CFMDepositPatientList::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMDepositPatientList::OnToDateChange(){
	
} */
/*void CFMDepositPatientList::OnToDateSetfocus(){
	
} */
/*void CFMDepositPatientList::OnToDateKillfocus(){
	
} */
int CFMDepositPatientList::OnToDateCheckValue(){
	return 0;
} 
void CFMDepositPatientList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMDepositPatientList::OnClerkSelendok(){
	 
}
/*void CFMDepositPatientList::OnClerkSetfocus(){
	
}*/
/*void CFMDepositPatientList::OnClerkKillfocus(){
	
}*/
long CFMDepositPatientList::OnClerkLoadData()
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
/*void CFMDepositPatientList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMDepositPatientList::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);

	CReport rpt;
	CString szSQL, tmpStr, szSysDate;
	CString szObject, szTemp;
	CString szRecord;

	CRecord rs(&pMF->m_db);
	double nCost = 0;
	long double nTotal[8];

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);

	for (int i = 0; i < 8; i++)
		nTotal[i] = 0;

	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_ICONERROR | MB_OK);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/HF_DANHSACHBENHNHANTAMGUI.RPT")))
		return;

	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	CReportSection *rptDetail;

	int nIndex = 1;

	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();

		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		szRecord.Empty();
		rs.GetValue(_T("recordno"), tmpStr);
		szRecord = tmpStr;
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("deptid"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("insamount"), nCost);
		nTotal[5] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		
		if (szRecord.Left(2) == _T("DV"))
		{
			if (nCost > 0)
			{
				rptDetail->SetValue(_T("9"), _T("X"));
			}
		}

		rs.GetValue(_T("servamount"), nCost);
		nTotal[6] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("amount"), nCost);
		nTotal[7] += nCost;
		FormatCurrency(nCost, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.MoveNext();
	}

	if (nTotal[7] > 0)
	{
		for (int  i = 5; i < 8; i++)
		{
			szTemp.Format(_T("s%d"), i + 1);
			//tmpStr.Format(_T("%.2Lf"), nTotal[i]);
			FormatCurrency(nTotal[i], tmpStr);
			rpt.GetReportFooter()->SetValue(szTemp, tmpStr);
		}
	}

	szSysDate = pMF->GetSysDate();

	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), 
		          szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CFMDepositPatientList::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	CString tmpStr, szTemp;

	UpdateData(TRUE);
	BeginWaitCursor();

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessageBox(_T("No Data"), MB_ICONERROR | MB_OK);
		return;
	}

	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 9);
	xls.SetColumnWidth(1, 12);
	xls.SetColumnWidth(2, 22);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 8);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 15);

	int nRow = 1;
	int nCol = 0;

	xls.SetRowHeight(6, 45);

	xls.SetCellMergedColumns(0, 1, 3);
	xls.SetCellMergedColumns(0, 2, 3);

	xls.SetCellMergedColumns(4, 1, 3);
	xls.SetCellMergedColumns(4, 2, 3);

	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(4, 1, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(4, 2, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellMergedColumns(nCol, nRow + 3, 6);
	xls.SetCellMergedColumns(nCol, nRow + 4, 6);

	xls.SetCellText(nCol, nRow + 3, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N T\x1EA0M G\x1EECI"),
					FMT_TEXT | FMT_CENTER, true, 16, 0);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 12, 0);

	TranslateString(_T("Document No"), tmpStr);
	xls.SetCellText(nCol, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Admission No"), tmpStr);
	xls.SetCellText(nCol + 1, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Patient Name"), tmpStr);
	xls.SetCellText(nCol + 2, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Dept"), tmpStr);
	xls.SetCellText(nCol + 3, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	xls.SetCellText(nCol + 4, nRow + 5, _T("\x110\x1ED5i \x111\x1ED1i t\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Health Insurance"), tmpStr);
	xls.SetCellText(nCol + 5, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	TranslateString(_T("Service"), tmpStr);
	xls.SetCellText(nCol + 6, nRow + 5, tmpStr, FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 11);

	nRow = 6;
	long double nTotal[7];
	double nCost = 0;
	int nIndex = -1;
	CString szRecord;

	for (int i = 0; i < 7; i++)
		nTotal[i] = 0;

	while (!rs.IsEOF())
	{
		nRow++;

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		szRecord.Empty();
		rs.GetValue(_T("recordno"), tmpStr);
		szRecord = tmpStr;
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("deptid"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("insamount"), nCost);
		nTotal[5] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		if (nCost > 0)
		{
			if (szRecord.Left(2) == _T("DV"))
				xls.SetCellText(nCol + 4, nRow, _T("X"), FMT_TEXT | FMT_WRAPING, true);
		}

		rs.GetValue(_T("servamount"), nCost);
		nTotal[6] += nCost;
		tmpStr.Format(_T("%.2f"), nCost);
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.MoveNext();
	}

	if (nTotal[5] + nTotal[6] > 0)
	{
		nRow++;
		xls.SetCellMergedColumns(nCol, nRow, 5);
		TranslateString(_T("Total Sum"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);

		for (int i = 5; i < 7; i++)
		{
			tmpStr.Format(_T("%.2Lf"), nTotal[i]);
			xls.SetCellText(nCol + i, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true);
		}
	}

	EndWaitCursor();
	xls.Save(_T("Exports\\DSBNTamGui.xls"));
} 
void CFMDepositPatientList::OnListDblClick()
{
	
} 
void CFMDepositPatientList::OnListSelectChange(int nOldItem, int nNewItem)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMDepositPatientList::OnListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMDepositPatientList::OnListLoadData()
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
void CFMDepositPatientList::OnOutPatientSelect(){
	
}
void CFMDepositPatientList::OnInPatientSelect(){
	
}
void CFMDepositPatientList::OnLockedSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
int CFMDepositPatientList::OnAddFMDepositPatientList(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMDepositPatientList::OnEditFMDepositPatientList(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMDepositPatientList::OnDeleteFMDepositPatientList(){
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
 		OnCancelFMDepositPatientList(); 
 	}
	return 0;
}
int CFMDepositPatientList::OnSaveFMDepositPatientList(){
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
 		//OnFMDepositPatientListListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMDepositPatientList::OnCancelFMDepositPatientList(){
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
int CFMDepositPatientList::OnFMDepositPatientListListLoadData(){
	return 0;
}

CString CFMDepositPatientList::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	CString szObject;

	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" and hfe_staff='%s' "), m_szClerkKey);

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
		szWhere.AppendFormat(_T(" and hfe_objectid in(%s) "), szObject);

	if (m_nOutPatient == 0)
	{
		szWhere.AppendFormat(_T(" and hfe_class in('E') "));
	}
	else
	{
		szWhere.AppendFormat(_T(" and hfe_class not in('E') "));
	}

	if (!m_bLocked)
	{
		szWhere.AppendFormat(_T(" and hfe_locked<>'Y' "));
	//	szWhere.AppendFormat(_T(" and hfe_date between cast_string2timestamp('%s') and cast_string2timestamp('%s') "),
	//		                 m_szFromDate, m_szToDate);
	}
	else
	{
		szWhere.AppendFormat(_T(" and hfe_locked='Y' "));
		szWhere.AppendFormat(_T(" and hfe_lockeddate between cast_string2timestamp('%s') and cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}

	szSQL.Format(_T(" select hd_docno as docno, ") \
		        _T("        hcr_recordno as recordno,") \
		        _T("        hfe_invoiceno as invoiceno,") \
				_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("        hfe_deptid as deptid,") \
				_T("        (select distinct sd_name from sys_dept where sd_id=hfe_deptid) as deptname,") \
				_T("        case when ho_type not in('S') then sum(hfe_amount) else 0 end as insamount,") \
				_T("        case when ho_type in('S') then sum(hfe_amount) else 0 end as servamount,") \
				_T("        sum(hfe_amount) as amount") \
				_T(" from hms_patient") \
				_T(" left join hms_doc on(hd_patientno=hp_patientno)") \
				_T(" left join hms_clinical_record on(hcr_docno=hd_docno)") \
				_T(" left join hms_fee_deposit on(hfe_docno=hd_docno)") \
				_T(" left join hms_object on(ho_id=hfe_objectid)") \
				_T(" where hfe_status='P' and hd_object is not null and hd_object<>'0' %s") \
				_T(" group by hfe_invoiceno, hd_docno, hcr_recordno, hp_surname, hp_midname, hp_firstname, hfe_deptid, ho_type") \
				_T(" order by hfe_invoiceno"),
				szWhere);

	return szSQL;
}

int CFMDepositPatientList::OnListCheckAll()
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

int CFMDepositPatientList::OnListUnCheckAll()
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