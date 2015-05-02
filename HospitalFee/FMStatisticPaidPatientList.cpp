#include "stdafx.h"
#include "FMStatisticPaidPatientList.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMStatisticPaidPatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMStatisticPaidPatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMStatisticPaidPatientList* )pWnd)->OnGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnGroupSelendokFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnGroupSelendok();
}
/*static void _OnGroupSetfocusFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnGroupKillfocus();
}*/
/*static void _OnGroupKillfocusFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnGroupKillfocus();
}*/
static long _OnGroupLoadDataFnc(CWnd *pWnd){
	return ((CFMStatisticPaidPatientList *)pWnd)->OnGroupLoadData();
}
/*static void _OnGroupAddNewFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnGroupAddNew();
}*/
static void _OnItemSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMStatisticPaidPatientList* )pWnd)->OnItemSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnItemSelendokFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnItemSelendok();
}
/*static void _OnItemSetfocusFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnItemKillfocus();
}*/
/*static void _OnItemKillfocusFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnItemKillfocus();
}*/
static long _OnItemLoadDataFnc(CWnd *pWnd){
	return ((CFMStatisticPaidPatientList *)pWnd)->OnItemLoadData();
}
/*static void _OnItemAddNewFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnItemAddNew();
}*/
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMStatisticPaidPatientList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMStatisticPaidPatientList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMStatisticPaidPatientList *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMStatisticPaidPatientList *pVw = (CFMStatisticPaidPatientList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMStatisticPaidPatientList *pVw = (CFMStatisticPaidPatientList *)pWnd;
	pVw->OnExportSelect();
}
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMStatisticPaidPatientList*)pWnd)->OnLockedSelect();
}
static int _OnAddFMStatisticPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMStatisticPaidPatientList*)pWnd)->OnAddFMStatisticPaidPatientList();
} 
static int _OnEditFMStatisticPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMStatisticPaidPatientList*)pWnd)->OnEditFMStatisticPaidPatientList();
} 
static int _OnDeleteFMStatisticPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMStatisticPaidPatientList*)pWnd)->OnDeleteFMStatisticPaidPatientList();
} 
static int _OnSaveFMStatisticPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMStatisticPaidPatientList*)pWnd)->OnSaveFMStatisticPaidPatientList();
} 
static int _OnCancelFMStatisticPaidPatientListFnc(CWnd *pWnd){
	 return ((CFMStatisticPaidPatientList*)pWnd)->OnCancelFMStatisticPaidPatientList();
} 
CFMStatisticPaidPatientList::CFMStatisticPaidPatientList(CWnd *pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 220;
	SetDefaultValues();
}
CFMStatisticPaidPatientList::~CFMStatisticPaidPatientList(){
}
void CFMStatisticPaidPatientList::OnCreateComponents()
{
	/*m_wndTestPatientList.Create(this, _T("Test Patient List"), 5, 5, 490, 150);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndGroupLabel.Create(this, _T("Group"), 10, 60, 90, 85);
	m_wndGroup.Create(this,95, 60, 485, 85); 
	m_wndItemLabel.Create(this, _T("Item"), 10, 90, 90, 115);
	m_wndItem.Create(this,95, 90, 485, 115); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 120, 90, 145);
	m_wndClerk.Create(this,95, 120, 485, 145); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 300, 155, 410, 180);
	m_wndExport.Create(this, _T("&Export"), 415, 155, 490, 180);*/

	m_wndTestPatientList.Create(this, _T("Statistic Paid Patient List"), 5, 5, 430, 180);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 30, 300, 55);
	m_wndToDate.Create(this,305, 30, 425, 55); 
	m_wndGroupLabel.Create(this, _T("Group"), 10, 60, 90, 85);
	m_wndGroup.Create(this,95, 60, 425, 85); 
	m_wndItemLabel.Create(this, _T("Item"), 10, 90, 90, 115);
	m_wndItem.Create(this,95, 90, 425, 115); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 120, 90, 145);
	m_wndClerk.Create(this,95, 120, 425, 145); 
	m_wndLocked.Create(this, _T("Locked"), 315, 150, 425, 175);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 205, 185, 315, 210);
	m_wndExport.Create(this, _T("&Export"), 320, 185, 430, 210);
}

void CFMStatisticPaidPatientList::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	//m_wndGroup.SetCheckValue(true);
	m_wndGroup.LimitText(62);

	//m_wndItem.SetCheckValue(true);
	m_wndItem.LimitText(1065);

	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(65);


	m_wndGroup.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	m_wndGroup.InsertColumn(1, _T("Name"), CFMT_TEXT, 280);


	m_wndItem.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndItem.InsertColumn(1, _T("Name"), CFMT_TEXT, 280);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 270);

}
void CFMStatisticPaidPatientList::OnSetWindowEvents()
{
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndGroup.SetEvent(WE_SELENDOK, _OnGroupSelendokFnc);
	//m_wndGroup.SetEvent(WE_SETFOCUS, _OnGroupSetfocusFnc);
	//m_wndGroup.SetEvent(WE_KILLFOCUS, _OnGroupKillfocusFnc);
	m_wndGroup.SetEvent(WE_SELCHANGE, _OnGroupSelectChangeFnc);
	m_wndGroup.SetEvent(WE_LOADDATA, _OnGroupLoadDataFnc);
	//m_wndGroup.SetEvent(WE_ADDNEW, _OnGroupAddNewFnc);
	m_wndItem.SetEvent(WE_SELENDOK, _OnItemSelendokFnc);
	//m_wndItem.SetEvent(WE_SETFOCUS, _OnItemSetfocusFnc);
	//m_wndItem.SetEvent(WE_KILLFOCUS, _OnItemKillfocusFnc);
	m_wndItem.SetEvent(WE_SELCHANGE, _OnItemSelectChangeFnc);
	m_wndItem.SetEvent(WE_LOADDATA, _OnItemLoadDataFnc);
	//m_wndItem.SetEvent(WE_ADDNEW, _OnItemAddNewFnc);
	m_wndClerk.SetEvent(WE_SELENDOK, _OnClerkSelendokFnc);
	//m_wndClerk.SetEvent(WE_SETFOCUS, _OnClerkSetfocusFnc);
	//m_wndClerk.SetEvent(WE_KILLFOCUS, _OnClerkKillfocusFnc);
	m_wndClerk.SetEvent(WE_SELCHANGE, _OnClerkSelectChangeFnc);
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	//m_wndClerk.SetEvent(WE_ADDNEW, _OnClerkAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);

	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CFMStatisticPaidPatientList::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndGroup.GetDlgCtrlID(), m_szGroupKey);
	DDX_TextEx(pDX, m_wndItem.GetDlgCtrlID(), m_szItemKey);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);
}
void CFMStatisticPaidPatientList::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMStatisticPaidPatientList::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMStatisticPaidPatientList::SetDefaultValues()
{

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szGroupKey.Empty();
	m_szItemKey.Empty();
	m_szClerkKey.Empty();
	m_bLocked=FALSE;

}
int CFMStatisticPaidPatientList::SetMode(int nMode){
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
/*void CFMStatisticPaidPatientList::OnFromDateChange(){
	
} */
/*void CFMStatisticPaidPatientList::OnFromDateSetfocus(){
	
} */
/*void CFMStatisticPaidPatientList::OnFromDateKillfocus(){
	
} */
int CFMStatisticPaidPatientList::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMStatisticPaidPatientList::OnToDateChange(){
	
} */
/*void CFMStatisticPaidPatientList::OnToDateSetfocus(){
	
} */
/*void CFMStatisticPaidPatientList::OnToDateKillfocus(){
	
} */
int CFMStatisticPaidPatientList::OnToDateCheckValue(){
	return 0;
} 
void CFMStatisticPaidPatientList::OnGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMStatisticPaidPatientList::OnGroupSelendok(){
	 
}
/*void CFMStatisticPaidPatientList::OnGroupSetfocus(){
	
}*/
/*void CFMStatisticPaidPatientList::OnGroupKillfocus(){
	
}*/

long CFMStatisticPaidPatientList::OnGroupLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;

	if(m_wndGroup.IsSearchKey() && !m_szGroupKey.IsEmpty())
	{
		szWhere.Format(_T(" AND hfg_id='%s' "), m_szGroupKey);
	}

	szSQL.Format(_T("SELECT hfg_id as id, hfg_name as description FROM hms_fee_group ORDER BY hfg_id"));
	m_wndGroup.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndGroup.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
}

/*void CFMStatisticPaidPatientList::OnGroupAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */

void CFMStatisticPaidPatientList::OnItemSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMStatisticPaidPatientList::OnItemSelendok(){
	 
}
/*void CFMStatisticPaidPatientList::OnItemSetfocus(){
	
}*/
/*void CFMStatisticPaidPatientList::OnItemKillfocus(){
	
}*/
long CFMStatisticPaidPatientList::OnItemLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndItem.IsSearchKey() && !m_szItemKey.IsEmpty())
	{
		szWhere.Format(_T(" AND hfl_feeid='%s' "), m_szItemKey);
	}

	if (m_szGroupKey.IsEmpty())
	{
		m_wndGroup.SetToolTipMessage(_T("Pick an item."));
		return 0;
	}

	szSQL.Format(_T("SELECT hfl_feeid as id, hfl_name as description ") \
		         _T("FROM hms_fee_list WHERE hfl_groupid='%s'"), m_szGroupKey);

	m_wndItem.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndItem.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMStatisticPaidPatientList::OnItemAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMStatisticPaidPatientList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMStatisticPaidPatientList::OnClerkSelendok(){
	 
}
/*void CFMStatisticPaidPatientList::OnClerkSetfocus(){
	
}*/
/*void CFMStatisticPaidPatientList::OnClerkKillfocus(){
	
}*/
long CFMStatisticPaidPatientList::OnClerkLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	};
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
/*void CFMStatisticPaidPatientList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMStatisticPaidPatientList::OnLockedSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMStatisticPaidPatientList::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szSysDate, szOldLine, szNewLine, szTemp;
	int nIdx = 0;
	double nAmount = 0;
	long double nGroupTotal[12];
	long double nTotal[12];

	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);

	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/HF_BANGTHONGKEDSBNNOPTIEN.RPT")))
		return;

	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);

	CReportSection* rptDetail = NULL;
	for(int i = 0; i <= 12; i++)
	{
		nGroupTotal[i] = 0;
		nTotal[i] = 0;
	}
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("receiptdate"), tmpStr);
		szNewLine = CDate::Convert(tmpStr, yyyymmdd, ddmmyyyy);
		if (!szNewLine.IsEmpty() && szNewLine != szOldLine)
		{
			if (nGroupTotal[6] > 0)
			{
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
				rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng ng\xE0y"));
				for(int i = 4; i < 10; i++)
				{
					FormatCurrency(nGroupTotal[i], tmpStr);
					szTemp.Format(_T("s%d"), i);
					rptDetail->SetValue(szTemp, tmpStr);
					nTotal[i] += nGroupTotal[i];
					nGroupTotal[i] = 0;
				}
			}
			szOldLine = szNewLine;
		}
		rptDetail = rpt.AddDetail();

		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		//tmpStr = CDate::Convert(rs.GetValue(_T("receiptdate")), yyyymmdd, ddmmyyyy);
		//rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("thungoai"), nAmount);
		nGroupTotal[4] += nAmount;
		FormatCurrency(nAmount, tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("thunoi"), nAmount);
		nGroupTotal[5] += nAmount;
		FormatCurrency(nAmount, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("tongthu"), nAmount);
		nGroupTotal[6] += nAmount;
		FormatCurrency(nAmount, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("trangoai"), nAmount);
		nGroupTotal[7] += nAmount;
		FormatCurrency(nAmount, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		//rs.GetValue(_T("tranoi"), nAmount);
		//nGroupTotal[7] += nAmount;
		FormatCurrency(0, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("trangoai"), nAmount);
		nGroupTotal[9] += nAmount;
		FormatCurrency(nAmount, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.MoveNext();
	}

	if (nGroupTotal[6] > 0)
	{
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		//TranslateString(_T("\x43\x1ED9ng kho\x61"), tmpStr);
		rptDetail->SetValue(_T("TotalGroup"), _T("\x43\x1ED9ng ng\xE0y"));
		for(int i = 4; i < 10; i++)
		{
			FormatCurrency(nGroupTotal[i], tmpStr);
			szTemp.Format(_T("s%d"), i);
			rptDetail->SetValue(szTemp, tmpStr);
			nTotal[i] += nGroupTotal[i];
			nGroupTotal[i] = 0;
		}
	}
	
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	//FormatCurrency(nTotal, tmpStr);
	//rptDetail->SetValue(_T("s5"), tmpStr);

	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
		          szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CFMStatisticPaidPatientList::OnExportSelect()
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
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}

	xls.CreateSheet(1);
	xls.SetWorksheet(0);

	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 15);
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

	xls.SetCellText(nCol, nRow + 2, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N N\x1ED8P TI\x1EC0N \x58\xC9T NGHI\x1EC6M"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn \x62\x1EC7nh nh\xE2n"));
	arrCol.Add(_T("Ng\xE0y thu"));
	arrCol.Add(_T("S\x1ED1 h\x1ED3 s\x1A1"));
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

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		tmpStr = CDate::Convert(rs.GetValue(_T("receiptdate")), yyyymmdd, ddmmyyyy);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_DATE | FMT_WRAPING);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("amount"), nAmount);
		nTotal += nAmount;
		tmpStr.Format(_T("%.2f"), nAmount);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		nRow++;
		rs.MoveNext();
	}

	//FormatCurrency(nTotal, tmpStr);
	xls.SetCellMergedColumns(nCol, nRow, 4);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 10);

	tmpStr.Format(_T("%.2Lf"), nTotal);
	xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);

	xls.Save(_T("Exports\\Danh Sach Benh Nhan Nop Tien Xet Nghiem.xls"));
	
} 
int CFMStatisticPaidPatientList::OnAddFMStatisticPaidPatientList(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMStatisticPaidPatientList::OnEditFMStatisticPaidPatientList(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMStatisticPaidPatientList::OnDeleteFMStatisticPaidPatientList(){
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
 		OnCancelFMStatisticPaidPatientList(); 
 	}
	return 0;
}
int CFMStatisticPaidPatientList::OnSaveFMStatisticPaidPatientList(){
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
 		//OnFMTestPatientListListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMStatisticPaidPatientList::OnCancelFMStatisticPaidPatientList(){
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
int CFMStatisticPaidPatientList::OnFMStatisticPaidPatientListListLoadData(){
	return 0;
}

CString CFMStatisticPaidPatientList::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere, szLock;
	szWhere.Empty();
	
	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND fi.hfe_staff='%s' "), m_szClerkKey);
	if (!m_szGroupKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND fe.hfe_group='%s' "), m_szGroupKey);
	if (!m_szItemKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND fe.hfe_itemid='%s' "), m_szItemKey);

	if (m_bLocked)
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_locked='Y' "));
		szLock.AppendFormat(_T(" AND fi.hfe_locked='Y' "));
		szWhere.AppendFormat(_T(" AND fac_posteddate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}
	else
	{
		szWhere.AppendFormat(_T(" AND fac_posteddate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}

	szSQL.Format(_T("SELECT deptid, ") \
				_T("       docno, ") \
				_T("       pname, ") \
				_T("       receiptdate, ") \
				_T("       Sum(thungoai) AS thungoai, ") \
				_T("       Sum(thunoi) AS thunoi, ") \
				_T("       Sum(thunoi + thungoai) AS tongthu, ") \
				_T("       Sum(trangoai) AS trangoai ") \
				_T("FROM   (SELECT    fi.hfe_deptid AS deptid, ") \
				_T("                  fi.hfe_docno AS docno, ") \
				_T("                  get_patientname(fi.hfe_docno) AS pname, ") \
				_T("                  Cast_timestamp2date(fac_posteddate) AS receiptdate, ") \
				_T("                  CASE WHEN fi.hfe_class = 'I' ") \
				_T("                       AND hfe_category <> 'Y' THEN fe.hfe_patpaid ") \
				_T("                  ELSE 0 ") \
				_T("                  END AS thungoai, ") \
				_T("                  CASE WHEN fi.hfe_class IN ( 'E', 'X' ) ") \
				_T("                       AND hfe_category <> 'Y' THEN fe.hfe_patpaid ") \
				_T("                  ELSE 0 ") \
				_T("                  END AS thunoi, ") \
				_T("                  0 AS trangoai ") \
				_T("        FROM      hms_fee_invoice fi ") \
				_T("        LEFT JOIN hms_fee fe ON( fi.hfe_docno = fe.hfe_docno ") \
				_T("                             AND fi.hfe_invoiceno = fe.hfe_invoiceno ) ") \
				_T("        LEFT JOIN fa_cash ON( fi.hfe_cash_id = fac_cash_id ) ") \
				_T("        LEFT JOIN hms_doc d ON ( d.hd_docno = fi.hfe_docno ) ") \
				_T("        WHERE     1=1 %s ") \
				_T("                  AND fi.hfe_status = 'P' ") \
				_T("                  AND d.hd_object = '7' ") \
				_T("        UNION ALL ") \
				_T("        SELECT    fi.hfe_deptid AS deptid, ") \
				_T("                  fi.hfe_docno AS docno, ") \
				_T("                  get_patientname(fi.hfe_docno) AS pname, ") \
				_T("                  Cast_timestamp2date(fac_posteddate) AS receiptdate, ") \
				_T("                  0 AS thungoai, ") \
				_T("                  0 AS thunoi, ") \
				_T("                  CASE WHEN fi.hfe_class = 'I' THEN fi.hfe_refund ") \
				_T("                  ELSE 0 ") \
				_T("                  END AS trangoai ") \
				_T("        FROM      hms_fee_invoice fi ") \
				_T("        LEFT JOIN fa_cash ON( fi.hfe_cash_id = fac_cash_id ) ") \
				_T("        LEFT JOIN hms_doc d ON ( d.hd_docno = fi.hfe_docno ) ") \
				_T("        WHERE     fac_posteddate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s') %s") \
				_T("                  AND fi.hfe_status = 'P' ") \
				_T("                  AND d.hd_object = '7') ") \
				_T("GROUP  BY deptid, ") \
				_T("          docno, pname,") \
				_T("          receiptdate ") \
				_T("HAVING Sum(thungoai) > 0 ") \
				_T("        OR Sum(thunoi) > 0 ") \
				_T("        OR Sum(trangoai) > 0 ") \
				_T("ORDER  BY receiptdate, ") \
				_T("          deptid, docno,") \
				_T("          pname "), szWhere, m_szFromDate, m_szToDate, szLock);

	return szSQL;
}
