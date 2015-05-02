#include "stdafx.h"
#include "TMOperationFosteringList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CTMOperationFosteringList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CTMOperationFosteringList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CTMOperationFosteringList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CTMOperationFosteringList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CTMOperationFosteringList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CTMOperationFosteringList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CTMOperationFosteringList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CTMOperationFosteringList *)pWnd)->OnToDateCheckValue();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMOperationFosteringList* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CTMOperationFosteringList *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CTMOperationFosteringList *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CTMOperationFosteringList *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CTMOperationFosteringList *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CTMOperationFosteringList *)pWnd)->OnObjectAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CTMOperationFosteringList *pVw = (CTMOperationFosteringList *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CTMOperationFosteringList *pVw = (CTMOperationFosteringList *)pWnd;
	pVw->OnExportSelect();
} 
CTMOperationFosteringList::CTMOperationFosteringList(CWnd *pParent){

	m_nDlgWidth = 500;
	m_nDlgHeight = 220;
	SetDefaultValues();
}
CTMOperationFosteringList::~CTMOperationFosteringList(){
}
void CTMOperationFosteringList::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 490, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndObject.SetCheckBox(true);
	m_wndObjectLabel.Create(this, _T("Object"), 10, 60, 90, 85);
	m_wndObject.Create(this,95, 60, 485, 85); 
	m_wndPaid.Create(this, _T("Paid"), 5, 95, 85, 120);
	m_wndPrint.Create(this, _T("&Print"), 335, 95, 410, 120);
	m_wndExport.Create(this, _T("&Export"), 415, 95, 490, 120);

}
void CTMOperationFosteringList::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	//m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(1024);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 40);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 260);

}
void CTMOperationFosteringList::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	UpdateData(FALSE);

}
void CTMOperationFosteringList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_Check(pDX, m_wndPaid.GetDlgCtrlID(), m_bPaid);

}
void CTMOperationFosteringList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szObjectKey.Empty();
	m_bPaid = FALSE;

}
int CTMOperationFosteringList::SetMode(int nMode){
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
/*void CTMOperationFosteringList::OnFromDateChange(){
	
} */
/*void CTMOperationFosteringList::OnFromDateSetfocus(){
	
} */
/*void CTMOperationFosteringList::OnFromDateKillfocus(){
	
} */
int CTMOperationFosteringList::OnFromDateCheckValue(){
	return 0;
} 
/*void CTMOperationFosteringList::OnToDateChange(){
	
} */
/*void CTMOperationFosteringList::OnToDateSetfocus(){
	
} */
/*void CTMOperationFosteringList::OnToDateKillfocus(){
	
} */
int CTMOperationFosteringList::OnToDateCheckValue(){
	return 0;
} 
void CTMOperationFosteringList::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMOperationFosteringList::OnObjectSelendok(){
	 
}
/*void CTMOperationFosteringList::OnObjectSetfocus(){
	
}*/
/*void CTMOperationFosteringList::OnObjectKillfocus(){
	
}*/
long CTMOperationFosteringList::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty()){
		szWhere.Format(_T(" and ho_id='%s' "), m_szObjectKey);
	}
	m_wndObject.DeleteAllItems();
	int nCount = 0;
	szSQL.Format(_T("SELECT ho_id as id, ho_desc as name FROM hms_object WHERE 1=1 %s ORDER BY cast(ho_id as integer)"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")),
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CTMOperationFosteringList::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CTMOperationFosteringList::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMOperationFosteringList::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szMoneyInWord, szTemp;
	int nIdx = 1, nRow = 0;
	double nAmount = 0, nTotalAmount = 0;
	CStringArray arrCol;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	//Header
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(3, 20);
	xls.SetColumnWidth(4, 20);
	xls.SetColumnWidth(5, 20);
	xls.SetColumnWidth(6, 30);
	xls.SetColumnWidth(7, 10);
	xls.SetColumnWidth(8, 15);
	xls.SetColumnWidth(9, 15);
	xls.SetColumnWidth(10, 15);
	xls.SetColumnWidth(11, 15);
	xls.SetCellMergedColumns(0, 0, 3);
	xls.SetCellMergedColumns(0, 1, 3);
	xls.SetCellMergedColumns(0, 2, 3);
	xls.SetCellMergedColumns(0, 3, 11);
	xls.SetCellMergedColumns(0, 4, 11);
	xls.SetCellText(0, 0, pMF->m_CompanyInfo.sc_pname, 4098, true);
	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_name, 4098, true);
	xls.SetCellText(0, 2, pMF->GetCurrentDepartmentName(), 4098, true);
	TranslateString(_T("Operation Fostering List"), szTemp);
	StringUpper(szTemp, tmpStr); 
	xls.SetCellText(0, 3, tmpStr, 4098, true, 12);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), m_szFromDate, m_szToDate);
	xls.SetCellText(0, 4, tmpStr, 4098, true);
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("S\x1ED1 h\x1ED3 s\x1A1"));
	arrCol.Add(_T("S\x1ED1 BA"));
	arrCol.Add(_T("H\x1ECD v\xE0 t\xEAn"));
	arrCol.Add(_T("N\x103m sinh"));
	arrCol.Add(_T("Th\x1EBB/Qu\xE2n h\xE0m"));
	arrCol.Add(_T("\x44\x61nh m\x1EE5\x63 PTTT"));
	arrCol.Add(_T("T\xEAn PTTT"));
	arrCol.Add(_T("Ph\xE2n lo\x1EA1i PT, TT"));
	//arrCol.Add(_T("K\xEDp th\x1EF1\x63 hi\x1EC7n"));
	arrCol.Add(_T("\x43h\xEDnh"));
	arrCol.Add(_T("Ph\x1EE5"));
	arrCol.Add(_T("Gi\xFAp vi\x1EC7\x63"));
	arrCol.Add(_T("S\x1ED1 l\x1B0\x1EE3ng"));
	arrCol.Add(_T("S\x1ED1 ti\x1EC1n"));
	for (int i = 0; i < arrCol.GetCount(); i++)
		xls.SetCellText(i, 5, arrCol.GetAt(i), 4098, true);
	//Detail
	nRow = 6;
	while (!rs.IsEOF())
	{
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("doc_no")), 4098);
		xls.SetCellText(2, nRow, rs.GetValue(_T("record_no")), 4098);
		xls.SetCellText(3, nRow, rs.GetValue(_T("patient_name")), FMT_TEXT);
		xls.SetCellText(4, nRow, rs.GetValue(_T("yob")), 4098);
		xls.SetCellText(5, nRow, rs.GetValue(_T("extra_info")), FMT_TEXT);
		xls.SetCellText(6, nRow, rs.GetValue(_T("operation_name1")), FMT_TEXT);
		xls.SetCellText(7, nRow, rs.GetValue(_T("operation_name")), FMT_TEXT);
		xls.SetCellText(8, nRow, rs.GetValue(_T("operation_type")), 4098);
		xls.SetCellText(9, nRow, rs.GetValue(_T("practitioner")), FMT_TEXT);
		xls.SetCellText(10, nRow, rs.GetValue(_T("assistant")), FMT_TEXT);
		xls.SetCellText(11, nRow, rs.GetValue(_T("anethetist")), FMT_TEXT);
		xls.SetCellText(12, nRow, rs.GetValue(_T("quantity")), FMT_NUMBER1);
		rs.GetValue(_T("amount"), nAmount);
		nTotalAmount += nAmount;
		xls.SetCellText(13, nRow, double2str(nAmount), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if (nTotalAmount > 0)
	{
		xls.SetCellMergedColumns(0, nRow, 13);
		xls.SetCellText(0, nRow, _T("\x43\x1ED9ng"), 4098, true);
		xls.SetCellText(13, nRow, double2str(nTotalAmount), FMT_NUMBER1);
		nRow++;
		MoneyToString(double2str(nTotalAmount), szMoneyInWord);
		tmpStr.Format(_T("S\x1ED1 ti\x1EC1n \x62\x1EB1ng \x63h\x1EEF: %s"), szMoneyInWord);
		xls.SetCellMergedColumns(0, nRow, 13);
		xls.SetCellText(0, nRow, tmpStr, FMT_TEXT, true);
	}

	xls.Save(_T("Exports\\Danh sach de nghi boi duong PTTT.xls"));
	
} 

CString CTMOperationFosteringList::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd(); 
	CString szSQL, szWhere, szObjects;
	szWhere.Format(_T(" AND ho_orderdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (m_bPaid)
	{
		szWhere.Format(_T(" AND i.hfe_date BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')"), m_szFromDate, m_szToDate);
		szWhere.AppendFormat(_T(" AND f.hfe_status = 'P' AND i.hfe_status = 'P'"));
	}
	for (int i = 0; i < m_wndObject.GetItemCount(); i++)
	{
		m_wndObject.SetCurSel(i);
		if (m_wndObject.GetCheck(i))
		{
			if (!szObjects.IsEmpty())
				szObjects += _T(", ");
			szObjects += m_wndObject.GetCurrent(0);
		}
	}
	if (!szObjects.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_object IN (%s)"), szObjects);
	szWhere.AppendFormat(_T(" AND ho_deptid = '%s'"), pMF->GetCurrentDepartmentID());
	szSQL.Format(_T(" SELECT   ho_docno doc_no, Get_patientname(ho_docno) patient_name, hcr_recordno record_no,") \
				_T("		   extract(year from hp_birthdate) yob,") \
				_T("           CASE WHEN hd_rank > 0 THEN (select ss_vndesc from sys_sel where ss_id = 'hms_rank' and ss_code = hd_rank)||hp_workplace ") \
				_T("           ELSE ho_cardno ") \
				_T("           END                       extra_info, ") \
				_T("		   hfl_name as operation_name1,") \
				_T("           ho_comment                operation_name, ") \
				_T("           get_feegroupname(hfl_groupid)               operation_type, ") \
				_T("           get_username(ho_practitioner)           practitioner, ") \
				_T("           get_username(ho_assistant)              assistant, ") \
				_T("           get_username(ho_anesthetist)            anethetist, ") \
				_T("           ho_qty                    quantity, ") \
				_T("           ho_qty*hfl_retprice       amount ") \
				_T(" FROM      hms_operation ") \
				_T(" LEFT JOIN hms_doc ON ( hd_docno = ho_docno ) ") \
				_T(" LEFT JOIN hms_patient ON ( hd_patientno = hp_patientno ) ") \
				_T(" LEFT JOIN hms_clinical_record ON (hcr_docno = hd_docno)") \
				_T(" LEFT JOIN hms_fee f ON ( hfe_docno = ho_docno ") \
				_T("                        AND hfe_orderid = ho_idx ) ") \
				_T(" LEFT JOIN hms_fee_invoice i ON (i.hfe_docno = f.hfe_docno AND i.hfe_invoiceno = f.hfe_invoiceno)") \
				_T(" LEFT JOIN hms_fee_list ON ( ho_itemid = hfl_feeid ) ") \
				_T(" WHERE     ho_status <> 'O' %s") \
				_T(" ORDER BY i.hfe_deptid, ho_docno"), szWhere);

	return szSQL;
}