#include "stdafx.h"
#include "FMR99PaidPatientList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMR99PaidPatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMR99PaidPatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMR99PaidPatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMR99PaidPatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMR99PaidPatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMR99PaidPatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMR99PaidPatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMR99PaidPatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMR99PaidPatientList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMR99PaidPatientList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMR99PaidPatientList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMR99PaidPatientList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMR99PaidPatientList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMR99PaidPatientList *)pWnd)->OnClerkAddNew();
}*/
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMR99PaidPatientList*)pWnd)->OnLockedSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMR99PaidPatientList *pVw = (CFMR99PaidPatientList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMR99PaidPatientList *pVw = (CFMR99PaidPatientList *)pWnd;
	pVw->OnExportSelect();
} 
CFMR99PaidPatientList::CFMR99PaidPatientList(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CFMR99PaidPatientList::~CFMR99PaidPatientList(){
}
void CFMR99PaidPatientList::OnCreateComponents(){
	m_wndPaidPatientList.Create(this, _T("Paid Patient List"), 5, 5, 430, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 30, 300, 55);
	m_wndToDate.Create(this,305, 30, 425, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 425, 85); 
	m_wndLocked.Create(this, _T("Locked"), 5, 95, 105, 120);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 225, 95, 325, 120);
	m_wndExport.Create(this, _T("&Export"), 330, 95, 430, 120);

}
void CFMR99PaidPatientList::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);
	m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(1024);


}
void CFMR99PaidPatientList::OnSetWindowEvents(){
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
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	m_szClerkKey = pMF->GetCurrentUser();
	UpdateData(false);
	m_wndPrintPreview.EnableWindow(FALSE);

}
void CFMR99PaidPatientList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);

}
void CFMR99PaidPatientList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bLocked=FALSE;

}
int CFMR99PaidPatientList::SetMode(int nMode){
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
/*void CFMR99PaidPatientList::OnFromDateChange(){
	
} */
/*void CFMR99PaidPatientList::OnFromDateSetfocus(){
	
} */
/*void CFMR99PaidPatientList::OnFromDateKillfocus(){
	
} */
int CFMR99PaidPatientList::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMR99PaidPatientList::OnToDateChange(){
	
} */
/*void CFMR99PaidPatientList::OnToDateSetfocus(){
	
} */
/*void CFMR99PaidPatientList::OnToDateKillfocus(){
	
} */
int CFMR99PaidPatientList::OnToDateCheckValue(){
	return 0;
} 
void CFMR99PaidPatientList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMR99PaidPatientList::OnClerkSelendok(){
	 
}
/*void CFMR99PaidPatientList::OnClerkSetfocus(){
	
}*/
/*void CFMR99PaidPatientList::OnClerkKillfocus(){
	
}*/
long CFMR99PaidPatientList::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szClerkKey
};
	m_wndClerk.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CFMR99PaidPatientList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
	void CFMR99PaidPatientList::OnLockedSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMR99PaidPatientList::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMR99PaidPatientList::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data!"));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 30);
	xls.SetColumnWidth(3, 15);
	int nRow = 0, nCol = 0, nIdx = 1;
	double nCost = 0;
	long double nTotal = 0;
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 4);
	xls.SetCellMergedColumns(0, 3, 4);
	xls.SetCellText(0, 0, pMF->m_CompanyInfo.sc_pname, 4098, true);
	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_name, 4098, true);
	xls.SetCellText(0, 2, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N N\x1ED8P TI\x1EC0N PH\xD2NG \x39\x39"), FMT_TEXT | FMT_CENTER, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd, ddmmyyyy),
				  CDateTime::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, FMT_TEXT | FMT_CENTER);
	nRow = 5;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("docno")), FMT_TEXT | FMT_CENTER);
		xls.SetCellText(2, nRow, rs.GetValue(_T("patientname")), FMT_TEXT);
		rs.GetValue(_T("cost"), nCost);
		xls.SetCellText(3, nRow, double2str(nCost), FMT_NUMBER1);
		nTotal+= nCost;
		nRow++;
		rs.MoveNext();
	}
	if (nTotal > 0)
	{
		xls.SetCellText(2, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
		tmpStr.Format(_T("%f"), nTotal);
		xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1, true);
	}
	xls.Save(_T("Exports\\Danh sach BN nop tien phong 99.xls"));
} 

CString CFMR99PaidPatientList::GetQueryString(){
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND i.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') AND hfe_category <> 'Y'"), m_szFromDate, m_szToDate);
	if (m_bLocked)
		szWhere.AppendFormat(_T(" AND i.hfe_locked = 'Y'"));
	szSQL.Format(_T(" SELECT Get_patientname(hpc_docno) patientname, ") \
	_T("        hpc_docno                  docno, ") \
	_T("        SUM(hfe_cost)              cost ") \
	_T(" FROM   (SELECT    hpc_docno, ") \
	_T("                   hfe_cost ") \
	_T("         FROM      hmsv_paraclinic ") \
	_T("         LEFT JOIN hms_fee f ON ( hfe_type IN ( 'P', 'T' ) ") \
	_T("                                  AND hpc_docno = hfe_docno ") \
	_T("                                  AND hpc_orderid = hfe_orderid ) ") \
	_T("         LEFT JOIN hms_fee_invoice i ON ( i.hfe_docno = f.hfe_docno ") \
	_T("                                          AND i.hfe_invoiceno = f.hfe_invoiceno ) ") \
	_T("         WHERE     i.hfe_status = 'P' ") \
	_T("               AND f.hfe_status = 'P' ") \
	_T("               AND hpc_roomid = 99 %s") \
	_T("         UNION ALL ") \
	_T("         SELECT    ho_docno, ") \
	_T("                   hfe_cost ") \
	_T("         FROM      hms_operation ") \
	_T("         LEFT JOIN hms_fee f ON ( hfe_type IN ('O') ") \
	_T("                                  AND ho_docno = hfe_docno ") \
	_T("                                  AND ho_idx = hfe_orderid ) ") \
	_T("         LEFT JOIN hms_fee_invoice i ON ( i.hfe_docno = f.hfe_docno ") \
	_T("                                          AND i.hfe_invoiceno = f.hfe_invoiceno ) ") \
	_T("         WHERE     i.hfe_status = 'P' ") \
	_T("               AND f.hfe_status = 'P' ") \
	_T("               AND ho_roomid = 99 %s") \
	_T("         UNION ALL ") \
	_T("         SELECT    hpo_docno, ") \
	_T("                   hfe_cost ") \
	_T("         FROM      hms_pharmaorder ") \
	_T("         LEFT JOIN hms_fee f ON ( hpo_docno = hfe_docno ") \
	_T("                                  AND hpo_orderid = hfe_orderid ) ") \
	_T("         LEFT JOIN hms_fee_invoice i ON ( i.hfe_docno = f.hfe_docno ") \
	_T("                                          AND i.hfe_invoiceno = f.hfe_invoiceno ) ") \
	_T("         WHERE     i.hfe_status = 'P' ") \
	_T("               AND f.hfe_status = 'P' ") \
	_T("               AND hpo_roomid = 99 %s)") \
	_T(" GROUP  BY hpc_docno ") \
	_T(" ORDER  BY hpc_docno "), szWhere, szWhere, szWhere);

	return szSQL;
}