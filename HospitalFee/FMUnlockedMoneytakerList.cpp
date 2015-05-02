#include "stdafx.h"
#include "FMUnlockedMoneytakerList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMUnlockedMoneytakerList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMUnlockedMoneytakerList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMUnlockedMoneytakerList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMUnlockedMoneytakerList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMUnlockedMoneytakerList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMUnlockedMoneytakerList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMUnlockedMoneytakerList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMUnlockedMoneytakerList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMUnlockedMoneytakerList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMUnlockedMoneytakerList *)pWnd)->OnClerkAddNew();
}*/
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMUnlockedMoneytakerList*)pWnd)->OnLockedSelect();
}
static void _OnPostedSelectFnc(CWnd *pWnd){
	 ((CFMUnlockedMoneytakerList*)pWnd)->OnPostedSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMUnlockedMoneytakerList *pVw = (CFMUnlockedMoneytakerList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExcelSelectFnc(CWnd *pWnd){
	CFMUnlockedMoneytakerList *pVw = (CFMUnlockedMoneytakerList *)pWnd;
	pVw->OnExcelSelect();
} 
CFMUnlockedMoneytakerList::CFMUnlockedMoneytakerList(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CFMUnlockedMoneytakerList::~CFMUnlockedMoneytakerList(){
}
void CFMUnlockedMoneytakerList::OnCreateComponents(){
	m_wndUnlockedMoneytakerList.Create(this, _T("Unlocked Money-taker List"), 5, 5, 430, 60);
	m_wndToDateLabel.Create(this, _T("To Date"), 10, 30, 90, 55);
	m_wndToDate.Create(this,95, 30, 215, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 220, 30, 300, 55);
	m_wndClerk.Create(this,305, 30, 425, 55); 
	m_wndLocked.Create(this, _T("Locked & not Posted"), 5, 65, 215, 90);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 245, 65, 345, 90);
	m_wndExcel.Create(this, _T("&Excel"), 350, 65, 430, 90);

}
void CFMUnlockedMoneytakerList::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndToDate.SetCheckValue(true);
	m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 200);

}
void CFMUnlockedMoneytakerList::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
	//m_wndPosted.SetEvent(WE_CLICK, _OnPostedSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExcel.SetEvent(WE_CLICK, _OnExcelSelectFnc);
	m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CFMUnlockedMoneytakerList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);
	//DDX_Check(pDX, m_wndPosted.GetDlgCtrlID(), m_bPosted);

}
void CFMUnlockedMoneytakerList::SetDefaultValues(){

	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bLocked=FALSE;
	m_bPosted=FALSE;

}
int CFMUnlockedMoneytakerList::SetMode(int nMode){
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
/*void CFMUnlockedMoneytakerList::OnToDateChange(){
	
} */
/*void CFMUnlockedMoneytakerList::OnToDateSetfocus(){
	
} */
/*void CFMUnlockedMoneytakerList::OnToDateKillfocus(){
	
} */
int CFMUnlockedMoneytakerList::OnToDateCheckValue(){
	return 0;
} 
void CFMUnlockedMoneytakerList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMUnlockedMoneytakerList::OnClerkSelendok(){
	 
}
/*void CFMUnlockedMoneytakerList::OnClerkSetfocus(){
	
}*/
/*void CFMUnlockedMoneytakerList::OnClerkKillfocus(){
	
}*/
long CFMUnlockedMoneytakerList::OnClerkLoadData(){
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
/*void CFMUnlockedMoneytakerList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
	void CFMUnlockedMoneytakerList::OnLockedSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
	void CFMUnlockedMoneytakerList::OnPostedSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMUnlockedMoneytakerList::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CReportSection * rptDetail = NULL, *rptFooter = NULL;
	CString szSQL, szReportName, tmpStr, szSysDate, szPos;
	CStringArray arrCol;
	long double nTtl[6];
	double nCost = 0;
	int nIdx = 1;
	if (m_bLocked)
		szReportName = _T("HF_DANHSACHTHUTIENDAKHOASOCHUANOPTIEN.RPT");
	else
		szReportName = _T("HF_DANHSACHTHUTIENCHUAKHOASO.RPT");
	tmpStr.Format(_T("Reports\\HMS\\%s"), szReportName);
	if (!rpt.Init(tmpStr))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	for (int i = 0; i < 6; i++)
		nTtl[i] = 0;
	arrCol.Add(_T("eamount"));
	arrCol.Add(_T("deposit"));
	arrCol.Add(_T("iamount"));
	arrCol.Add(_T("inpatient_deposit_paid"));
	arrCol.Add(_T("inpatient_payment"));
	arrCol.Add(_T("totalamount"));
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(_T("\x110\x1EBFn ng\xE0y %s"), m_szToDate);
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));
		rptDetail->SetValue(_T("2"), rs.GetValue(_T("staff")));
		if (m_bLocked)
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("cashid")));
		for (int i = 0; i < arrCol.GetCount(); i++)
		{
			rs.GetValue(arrCol.GetAt(i), nCost);
			nTtl[i] += nCost;
			if (m_bLocked)
				rptDetail->SetValue(int2str(i+4), double2str(nCost));
			else
				rptDetail->SetValue(int2str(i+3), double2str(nCost));
		}
		rs.MoveNext();	
	}
	rptFooter = rpt.GetReportFooter();
	for (int i = 0; i < 6; i++)
	{
		if (m_bLocked)
			szPos.Format(_T("s%d"), i+4);
		else
			szPos.Format(_T("s%d"), i+3);
		tmpStr.Format(_T("%f"), nTtl[i]);
		rptFooter->SetValue(szPos, tmpStr);
	}
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rptFooter->GetValue(_T("PrintDate")), szSysDate.Mid(8, 2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rptFooter->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CFMUnlockedMoneytakerList::OnExcelSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, szName, tmpStr;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	//
	//while ()
	if (m_bLocked)
		szName.Format(_T("Danh sach thu tien chua khoa so.xls"));
	else
		szName.Format(_T("Danh sach thu tien da khoa so chua thu tien.xls"));
	tmpStr.Format(_T("Exports\\%s"), szName);
	xls.Save(tmpStr);

} 

CString CFMUnlockedMoneytakerList::GetQueryString(){
	CString szSQL, szWhere, szGroupBy, szOrderBy;
	CString szWhere1, szField;
	szWhere.Format(_T(" AND NVL(hfe_locked, 'N') <> 'Y'"));
	szGroupBy.Format(_T(" GROUP BY hfe_staff"));
	if (m_bLocked)
	{
		szField.Format(_T(" fac_invoiceno cashid,"));
		szWhere.Format(_T(" AND NVL(hfe_locked, 'N') = 'Y'"));
		szWhere1.Format(_T(" AND fac_posted = 'N'"));
		szGroupBy.Format(_T(" GROUP BY hfe_staff, fac_invoiceno"));
		szOrderBy.Format(_T(" ORDER BY fac_invoiceno"));
	}
	szWhere.AppendFormat(_T(" AND hfe_status = 'P' AND su_deptid = 'PTC' AND hfe_date < cast_string2timestamp('%s')"), m_szToDate);
	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_staff = '%s'"), m_szClerkKey);
	szSQL.Format(_T(" SELECT    get_username(hfe_staff) staff, ") \
				_T("		   %s") \
				_T("           SUM(patpaid)                               AS eamount, ") \
				_T("           SUM(deposit)                               AS deposit, ") \
				_T("           SUM(inpatient_cost)                        AS iamount, ") \
				_T("           SUM(inpatient_depositpaid)                 AS inpatient_deposit_paid, ") \
				_T("           SUM(inpatient_payment)                     AS inpatient_payment, ") \
				_T("           SUM(patpaid + inpatient_payment + deposit) AS totalamount ") \
				_T(" FROM      (SELECT hfe_staff, ") \
				_T("                   hfe_cash_id, ") \
				_T("                   hfe_patpaid AS patpaid, ") \
				_T("                   hfe_deposit AS deposit, ") \
				_T("                   0           AS inpatient_cost, ") \
				_T("                   0           inpatient_depositpaid, ") \
				_T("                   0           inpatient_payment ") \
				_T("            FROM   hms_fee_invoice ") \
				_T("			LEFT JOIN sys_user ON (hfe_staff = su_userid)") \
				_T("            WHERE  hfe_class IN ( 'E', 'X' ) %s") \
				_T("            UNION ALL ") \
				_T("            SELECT hfe_staff, ") \
				_T("                   hfe_cash_id, ") \
				_T("                   0, ") \
				_T("                   hfe_amount, ") \
				_T("                   0 AS inpatient_cost, ") \
				_T("                   0 inpatient_depositpaid, ") \
				_T("                   0 inpatient_payment ") \
				_T("            FROM   hms_fee_deposit ") \
				_T("			LEFT JOIN sys_user ON (hfe_staff = su_userid)") \
				_T("            WHERE  hfe_class IN('I', 'A') %s") \
				_T("            UNION ALL ") \
				_T("            SELECT hfe_staff, ") \
				_T("                   hfe_cash_id, ") \
				_T("                   0, ") \
				_T("                   0, ") \
				_T("                   hfe_amount  AS inpatient_cost, ") \
				_T("                   CASE WHEN hfe_deposit > hfe_patpaid THEN hfe_patpaid ") \
				_T("                   ELSE hfe_deposit ") \
				_T("                   END         AS inpatient_depositpaid, ") \
				_T("                   hfe_payment AS inpatient_payment ") \
				_T("            FROM   hms_fee_invoice ") \
				_T("			LEFT JOIN sys_user ON (hfe_staff = su_userid)") \
				_T("            WHERE  hfe_class IN( 'A', 'I' ) %s)") \
				_T(" LEFT JOIN fa_cash ON ( hfe_cash_id = fac_cash_id ) ") \
				_T(" WHERE     1 = 1 %s") \
				_T(" %s ") \
				_T(" HAVING    SUM(patpaid + inpatient_payment + deposit) > 0 %s "),szField, szWhere, szWhere, szWhere, szWhere1, szGroupBy, szOrderBy);
	return szSQL;
}