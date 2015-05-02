#include "stdafx.h"
#include "FMUnlockPatientList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMUnlockPatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMUnlockPatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMUnlockPatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMUnlockPatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMUnlockPatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMUnlockPatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMUnlockPatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMUnlockPatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMUnlockPatientList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMUnlockPatientList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMUnlockPatientList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMUnlockPatientList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMUnlockPatientList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMUnlockPatientList *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMUnlockPatientList *pVw = (CFMUnlockPatientList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMUnlockPatientList *pVw = (CFMUnlockPatientList *)pWnd;
	pVw->OnExportSelect();
}
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CFMUnlockPatientList*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CFMUnlockPatientList*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMUnlockPatientList*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CFMUnlockPatientList*)pWnd)->OnListDeleteItem();
}
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	  ((CFMUnlockPatientList*)pWnd)->OnOutPatientSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	  ((CFMUnlockPatientList*)pWnd)->OnInPatientSelect();
}
static void _OnAllHospitalsSelectFnc(CWnd *pWnd){
	  ((CFMUnlockPatientList*)pWnd)->OnAllHospitalsSelect();
}
static int _OnAddFMUnlockPatientListFnc(CWnd *pWnd){
	 return ((CFMUnlockPatientList*)pWnd)->OnAddFMUnlockPatientList();
} 
static int _OnEditFMUnlockPatientListFnc(CWnd *pWnd){
	 return ((CFMUnlockPatientList*)pWnd)->OnEditFMUnlockPatientList();
} 
static int _OnDeleteFMUnlockPatientListFnc(CWnd *pWnd){
	 return ((CFMUnlockPatientList*)pWnd)->OnDeleteFMUnlockPatientList();
} 
static int _OnSaveFMUnlockPatientListFnc(CWnd *pWnd){
	 return ((CFMUnlockPatientList*)pWnd)->OnSaveFMUnlockPatientList();
} 
static int _OnCancelFMUnlockPatientListFnc(CWnd *pWnd){
	 return ((CFMUnlockPatientList*)pWnd)->OnCancelFMUnlockPatientList();
}

static int _OnListCheckAllFnc(CWnd *pWnd){
	return ((CFMUnlockPatientList*)pWnd)->OnListCheckAll();
}
static int _OnListUnCheckAllFnc(CWnd *pWnd){
	return ((CFMUnlockPatientList*)pWnd)->OnListUnCheckAll();
}

CFMUnlockPatientList::CFMUnlockPatientList(CWnd* pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 475;
	SetDefaultValues();
}
CFMUnlockPatientList::~CFMUnlockPatientList(){
}
void CFMUnlockPatientList::OnCreateComponents()
{
	m_wndReportConditionInfo.Create(this, _T("Report Condition"), 5, 5, 490, 435);
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
	m_wndOutPatient.Create(this, _T("OutPatient"), 235, 405, 315, 430);
	m_wndInPatient.Create(this, _T("InPatient"), 320, 405, 400, 430);
	m_wndAllHospitals.Create(this, _T("All Hospitals"), 405, 405, 485, 430);
}
void CFMUnlockPatientList::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

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
	m_wndList.InsertColumn(2, _T("Type"), CFMT_TEXT, 0);

	m_wndList.SetCheckBox(TRUE);
	m_wndList.ModifyStyle(0, LVS_NOSORTHEADER);

}
void CFMUnlockPatientList::OnSetWindowEvents()
{
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
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	m_wndAllHospitals.SetEvent(WE_CLICK, _OnAllHospitalsSelectFnc);

	m_wndList.SetWindowText(_T("Object"));
	m_wndList.AddEvent(1, _T("Check All"), _OnListCheckAllFnc);
	m_wndList.AddEvent(2, _T("UnCheck All"), _OnListUnCheckAllFnc);

	m_szFromDate = _T("2000/01/01");
	m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

	OnListLoadData();

	CString tmpStr;
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		tmpStr = m_wndList.GetItemText(i, 2);

		if (tmpStr == _T("S"))
			m_wndList.SetCheck(i, TRUE);
	}
}
void CFMUnlockPatientList::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Radio(pDX, m_wndOutPatient.GetDlgCtrlID(), m_nOutPatient);
	DDX_Radio(pDX, m_wndInPatient.GetDlgCtrlID(), m_nInPatient);
	DDX_Radio(pDX, m_wndAllHospitals.GetDlgCtrlID(), m_nAllHospitals);
}
void CFMUnlockPatientList::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMUnlockPatientList::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMUnlockPatientList::SetDefaultValues()
{

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_nOutPatient=0;
	m_nInPatient=1;
	m_nAllHospitals=1;

}
int CFMUnlockPatientList::SetMode(int nMode){
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
/*void CFMUnlockPatientList::OnFromDateChange(){
	
} */
/*void CFMUnlockPatientList::OnFromDateSetfocus(){
	
} */
/*void CFMUnlockPatientList::OnFromDateKillfocus(){
	
} */
int CFMUnlockPatientList::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMUnlockPatientList::OnToDateChange(){
	
} */
/*void CFMUnlockPatientList::OnToDateSetfocus(){
	
} */
/*void CFMUnlockPatientList::OnToDateKillfocus(){
	
} */
int CFMUnlockPatientList::OnToDateCheckValue(){
	return 0;
} 
void CFMUnlockPatientList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMUnlockPatientList::OnClerkSelendok(){
	 
}
/*void CFMUnlockPatientList::OnClerkSetfocus(){
	
}*/
/*void CFMUnlockPatientList::OnClerkKillfocus(){
	
}*/
long CFMUnlockPatientList::OnClerkLoadData()
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
/*void CFMUnlockPatientList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMUnlockPatientList::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
	UpdateData(true);

	CReport rpt;
	CRecord rs(&pMF->m_db);
	CRecord rss(&pMF->m_db);

	CString szSQL, tmpStr, szSysDate;
	CString szTemp, szClass;

	int nIdx = 0;
	double nAmount = 0;
	long double nTotal = 0;
	long double nGroupTotal = 0;
	m_nTotal = 0;

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);

	szSQL = GetDepositString();
	rss.ExecSQL(szSQL);

	if (rs.GetRecordCount() + rss.GetRecordCount() <= 0)
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/HF_DANHSACHBNCHUAKHOASO.RPT")))
		return;

	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	CReportSection *rptDetail;

	CString szGroup[] = {_T("I"), _T("II"), _T("III"), _T("IV")};

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
			if (szClass == _T("NO"))
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


		rs.GetValue(_T("deptid"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);


		rs.GetValue(_T("staff"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);


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

	//szSQL = GetDepositString();
	//rs.ExecSQL(szSQL);

	if (!rss.IsEOF())
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
		rptDetail->SetValue(_T("GroupName"), _T("Thu t\x1EA1m g\x1EEDi"));

		nIdx = 1;
		while (!rss.IsEOF())
		{
			rptDetail = rpt.AddDetail();
		
			tmpStr.Format(_T("%d"), nIdx++);
			rptDetail->SetValue(_T("1"), tmpStr);

			rss.GetValue(_T("pname"), tmpStr);
			rptDetail->SetValue(_T("2"), tmpStr);

			rss.GetValue(_T("docno"), tmpStr);
			rptDetail->SetValue(_T("3"), tmpStr);

			rss.GetValue(_T("deptid"), tmpStr);
			rptDetail->SetValue(_T("4"), tmpStr);

			rss.GetValue(_T("staff"), tmpStr);
			rptDetail->SetValue(_T("6"), tmpStr);

			rss.GetValue(_T("amount"), nAmount);
			nGroupTotal += nAmount;
			FormatCurrency(nAmount, tmpStr);
			rptDetail->SetValue(_T("5"), tmpStr);

			rss.MoveNext();
		}
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
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
		          szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));

	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CFMUnlockPatientList::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
	UpdateData(true);

	CExcel xls;
	CRecord rs(&pMF->m_db);
	CRecord rss(&pMF->m_db);

	CString szSQL, tmpStr, szTemp;
	CString szClass;

	int nCol = 0, nRow = 0, nIdx = 0;
	double nAmount = 0;
	long double nTotal = 0;
	long double nGroupTotal = 0;

	m_nTotal = 0;

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);

	szSQL = GetDepositString();
	rss.ExecSQL(szSQL);

	if (rs.GetRecordCount() + rss.GetRecordCount() <= 0)
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}

	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(2, 9);
	xls.SetColumnWidth(3, 8);
	xls.SetColumnWidth(4, 18);
	xls.SetColumnWidth(5, 15);

	xls.SetRowHeight(4, 40);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);

	xls.SetCellMergedColumns(nCol + 3, nRow, 4);
	xls.SetCellMergedColumns(nCol + 3, nRow + 1, 4);

	xls.SetCellMergedColumns(nCol, nRow + 2, 6);
	xls.SetCellMergedColumns(nCol, nRow + 3, 6);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 11);

	xls.SetCellText(nCol + 3, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 3, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol, nRow + 2, _T("\x42\x1EA2NG TH\x1ED0NG K\xCA \x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N \x43H\x1AF\x41 KH\xD3\x41 S\x1ED4"),
		            FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;

	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn \x62\x1EC7nh nh\xE2n"));
	arrCol.Add(_T("S\x1ED1 h\x1ED3 s\x1A1"));
	arrCol.Add(_T("Kho\x61 / Ph\xF2ng kh\xE1m"));
	arrCol.Add(_T("Ng\x1B0\x1EDDi thu"));
	arrCol.Add(_T("S\x1ED1 ti\x1EC1n"));


	CString szGroup[] = {_T("I"), _T("II"), _T("III"), _T("IV")};
	int nIndex = 0;
	nRow = 4;

	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol + i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10); 
	}

	while (!rs.IsEOF())
	{
		if (szClass != rs.GetValue(_T("fclass")))
		{
			if (nGroupTotal > 0)
			{
				nRow++;
				TranslateString(_T("Total Amt"), tmpStr);
				tmpStr.AppendFormat(_T(" %s"), szGroup[nIndex++]);
				xls.SetCellMergedColumns(nCol, nRow, 5);

				xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

				tmpStr.Format(_T("%.2Lf"), nGroupTotal);
				xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 11);

				nTotal += nGroupTotal;
				nGroupTotal = 0;
			}
			nIdx = 1;
			rs.GetValue(_T("fclass"), szClass);
			if (szClass == _T("NO"))
			{
				nRow++;
				xls.SetCellMergedColumns(nCol, nRow, 6);
				xls.SetCellText(nCol, nRow, _T("\x43\xE1\x63 kho\x61 n\x1ED9i tr\xFA"),
					            FMT_TEXT | FMT_CENTER, true, 11);
			}
		}

		nIdx++;
		nRow++;

		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("deptid"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("staff"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("amount"), nAmount);
		tmpStr.Format(_T("%.2f"), nAmount);
		nGroupTotal += nAmount;
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.MoveNext();
	}

	if (nGroupTotal > 0)
	{
		nRow++;
		TranslateString(_T("Total Amt"), tmpStr);
		tmpStr.AppendFormat(_T(" %s"), szGroup[nIndex++]);
		xls.SetCellMergedColumns(nCol, nRow, 5);

		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

		tmpStr.Format(_T("%.2Lf"), nGroupTotal);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 11);

		nTotal += nGroupTotal;
		nGroupTotal = 0;
	}

	if (!rss.IsEOF())
	{
		nRow++;
		xls.SetCellMergedColumns(nCol, nRow, 6);
		xls.SetCellText(nCol, nRow, _T("Thu t\x1EA1m g\x1EEDi"), FMT_TEXT | FMT_CENTER, true, 11);

		nIdx = 1;

		while (!rss.IsEOF())
		{
			nRow++;
			tmpStr.Format(_T("%d"), nIdx++);
			xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

			rss.GetValue(_T("pname"), tmpStr);
			xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

			rss.GetValue(_T("docno"), tmpStr);
			xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

			rss.GetValue(_T("deptid"), tmpStr);
			xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

			rss.GetValue(_T("staff"), tmpStr);
			xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

			rss.GetValue(_T("amount"), nAmount);
			nGroupTotal += nAmount;
			tmpStr.Format(_T("%.2f"), nAmount);
			xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

			rss.MoveNext();
		}
	}

	if (nGroupTotal > 0)
	{
		nRow++;
		TranslateString(_T("Total Amt"), tmpStr);
		tmpStr.AppendFormat(_T(" %s"), szGroup[nIndex++]);
		xls.SetCellMergedColumns(nCol, nRow, 5);

		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);

		tmpStr.Format(_T("%.2Lf"), nGroupTotal);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 11);

		nTotal += nGroupTotal;
		nGroupTotal = 0;
	}
	
	if (nTotal > 0)
	{
		nRow++;
		xls.SetCellMergedColumns(nCol, nRow, 5);
		xls.SetCellText(nCol, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 11);
		tmpStr.Format(_T("%.2Lf"), nTotal);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 11);
	}

	xls.Save(_T("Exports\\DanhSachBNChuaKhoaSo.xls"));
} 
void CFMUnlockPatientList::OnListDblClick(){
	
} 
void CFMUnlockPatientList::OnListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMUnlockPatientList::OnListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMUnlockPatientList::OnListLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndList.BeginLoad(); 
	m_wndList.DeleteAllItems(); 
	int nCount = 0;

	szSQL.Format(_T("SELECT ho_id AS ID, ") \
		         _T("ho_desc AS Description, ") \
				 _T("ho_type AS Type ") \
		         _T("FROM hms_object ") \
				 _T("ORDER BY cast(ho_id as integer)"));

	nCount = rs.ExecSQL(szSQL);

	while (!rs.IsEOF())
	{ 
		m_wndList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Description")),
			rs.GetValue(_T("Type")),
			NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad();

	return nCount;
}
void CFMUnlockPatientList::OnOutPatientSelect(){
	
}
void CFMUnlockPatientList::OnInPatientSelect(){
	
}
void CFMUnlockPatientList::OnAllHospitalsSelect(){
	
}
int CFMUnlockPatientList::OnAddFMUnlockPatientList(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMUnlockPatientList::OnEditFMUnlockPatientList(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMUnlockPatientList::OnDeleteFMUnlockPatientList(){
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
 		OnCancelFMUnlockPatientList();
 	}
	return 0;
}
int CFMUnlockPatientList::OnSaveFMUnlockPatientList(){
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
 		//OnFMUnlockPatientListListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CFMUnlockPatientList::OnCancelFMUnlockPatientList(){
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
int CFMUnlockPatientList::OnFMUnlockPatientListListLoadData(){
	return 0;
}

CString CFMUnlockPatientList::GetQueryString()
{
	CString szSQL, szWhere;
	CString szObject;
	szWhere.Empty();
	
	if (!m_szClerkKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hfe_staff='%s' "), m_szClerkKey);
	}

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
		szWhere.AppendFormat(_T(" AND hd_object in(%s) "), szObject);

	if (m_nOutPatient == 0)
	{
		szWhere.AppendFormat(_T(" AND hfe_class in('E','X') "));
	}
	else if (m_nInPatient == 0)
	{
		szWhere.AppendFormat(_T(" AND hfe_class not in('E','X') "));
	}

	szSQL.Format(_T(" SELECT trim(hp_surname||' '||hp_midname||' '||hp_firstname) AS pname,") \
				_T("        hd_docno AS docno,") \
				_T("        hfe_deptid AS deptid,") \
				_T("        hfe_patpaid AS amount,") \
				_T("        su_name AS staff,") \
				_T("        case when hfe_class IN('E') then 'NG'") \
                _T("        else 'NO' end AS fclass") \
				_T(" FROM hms_patient") \
				_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
				_T(" LEFT JOIN hms_object ON(ho_id=hd_object)") \
				_T(" LEFT JOIN hms_fee_invoice ON(hd_docno=hfe_docno)") \
				_T(" LEFT JOIN sys_user ON(su_userid=hfe_staff)") \
				_T(" WHERE hfe_patpaid > 0 AND hfe_status='P'") \
				_T("       AND hfe_locked<>'Y'") \
				_T("       AND hfe_createddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
				_T(" ORDER BY fclass, hfe_staff, hfe_date, hd_docno"),
				m_szFromDate, m_szToDate, szWhere);

	return szSQL;
}

CString CFMUnlockPatientList::GetDepositString()
{
	CString szSQL, szWhere;
	CString szObjects;
	szWhere.Empty();
	szObjects.Empty();

	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" and hfe_staff='%s' "), m_szClerkKey);

	szSQL.Format(_T(" select hd_docno as docno,") \
				_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("        hfe_deptid as deptid,") \
				_T("        su_name as staff,") \
				_T("        sum(hfe_amount) as amount") \
				_T(" from hms_patient") \
				_T(" left join hms_doc on(hd_patientno=hp_patientno)") \
				_T(" left join hms_clinical_record on(hcr_docno=hd_docno)") \
				_T(" left join hms_fee_deposit on(hfe_docno=hd_docno)") \
				_T(" left join hms_object on(ho_id=hfe_objectid)") \
				_T(" left join sys_user on(su_userid=hfe_staff)") \
				_T(" where hfe_status='P' and hd_object is not null and hd_object<>'0'") \
				_T("       and (hfe_refidx=0 or hfe_refidx is null)") \
				_T("       and hfe_locked<>'Y'") \
				_T("       and hfe_createddate between cast_string2timestamp('%s') and cast_string2timestamp('%s') %s") \
				_T(" group by hd_docno, hp_surname, hp_midname, hp_firstname, hfe_deptid, hfe_staff, su_name, ho_type, hfe_date") \
				_T(" order by hfe_staff, hfe_date, hd_docno, pname"),
				m_szFromDate, m_szToDate, szWhere);

	return szSQL;
}

int CFMUnlockPatientList::OnListCheckAll()
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

int CFMUnlockPatientList::OnListUnCheckAll()
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