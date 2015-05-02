#include "stdafx.h"
#include "FMRefundbyDepartment.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMRefundbyDepartment *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMRefundbyDepartment *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMRefundbyDepartment *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMRefundbyDepartment *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMRefundbyDepartment *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMRefundbyDepartment *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMRefundbyDepartment *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMRefundbyDepartment *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMRefundbyDepartment* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMRefundbyDepartment *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMRefundbyDepartment *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMRefundbyDepartment *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMRefundbyDepartment *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMRefundbyDepartment *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMRefundbyDepartment *pVw = (CFMRefundbyDepartment *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMRefundbyDepartment *pVw = (CFMRefundbyDepartment *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddFMRefundbyDepartmentFnc(CWnd *pWnd){
	 return ((CFMRefundbyDepartment*)pWnd)->OnAddFMRefundbyDepartment();
} 
static int _OnEditFMRefundbyDepartmentFnc(CWnd *pWnd){
	 return ((CFMRefundbyDepartment*)pWnd)->OnEditFMRefundbyDepartment();
} 
static int _OnDeleteFMRefundbyDepartmentFnc(CWnd *pWnd){
	 return ((CFMRefundbyDepartment*)pWnd)->OnDeleteFMRefundbyDepartment();
} 
static int _OnSaveFMRefundbyDepartmentFnc(CWnd *pWnd){
	 return ((CFMRefundbyDepartment*)pWnd)->OnSaveFMRefundbyDepartment();
} 
static int _OnCancelFMRefundbyDepartmentFnc(CWnd *pWnd){
	 return ((CFMRefundbyDepartment*)pWnd)->OnCancelFMRefundbyDepartment();
} 
CFMRefundbyDepartment::CFMRefundbyDepartment(CWnd *pParent){

	m_nDlgWidth = 500;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CFMRefundbyDepartment::~CFMRefundbyDepartment(){
}
void CFMRefundbyDepartment::OnCreateComponents()
{
	m_wndRefundbyDepartment.Create(this, _T("Refund by Department"), 5, 5, 430, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 30, 300, 55);
	m_wndToDate.Create(this,305, 30, 425, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 425, 85); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 240, 95, 350, 120);
	m_wndExport.Create(this, _T("&Export"), 355, 95, 430, 120);

}

void CFMRefundbyDepartment::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_NUMBER, 80);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 270);

}
void CFMRefundbyDepartment::OnSetWindowEvents(){
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
	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);
}
void CFMRefundbyDepartment::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);

}
void CFMRefundbyDepartment::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMRefundbyDepartment::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMRefundbyDepartment::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();

}
int CFMRefundbyDepartment::SetMode(int nMode){
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
/*void CFMRefundbyDepartment::OnFromDateChange(){
	
} */
/*void CFMRefundbyDepartment::OnFromDateSetfocus(){
	
} */
/*void CFMRefundbyDepartment::OnFromDateKillfocus(){
	
} */
int CFMRefundbyDepartment::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMRefundbyDepartment::OnToDateChange(){
	
} */
/*void CFMRefundbyDepartment::OnToDateSetfocus(){
	
} */
/*void CFMRefundbyDepartment::OnToDateKillfocus(){
	
} */
int CFMRefundbyDepartment::OnToDateCheckValue(){
	return 0;
} 
void CFMRefundbyDepartment::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMRefundbyDepartment::OnClerkSelendok(){
	 
}
/*void CFMRefundbyDepartment::OnClerkSetfocus(){
	
}*/
/*void CFMRefundbyDepartment::OnClerkKillfocus(){
	
}*/
long CFMRefundbyDepartment::OnClerkLoadData(){
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
/*void CFMRefundbyDepartment::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMRefundbyDepartment::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szSysDate;
	int nIdx = 0;
	double nAmount = 0, nTotal = 0;
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	if (!rpt.Init(_T("Reports/HMS/HF_BAOCAOHOANTRAKHOA.rpt")))
		return;
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection *rptDetail;
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();

		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("deptname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("totalpatient"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("amount"), nAmount);
		nTotal += nAmount;
		FormatCurrency(nAmount, tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.MoveNext();
	}
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	FormatCurrency(nTotal, tmpStr);
	rptDetail->SetValue(_T("s1"), tmpStr);
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
	
} 
void CFMRefundbyDepartment::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr;
	int nIdx = 0;
	int nCol = 0, nRow = 0;
	double nAmount = 0, nTotal = 0;

	CRecord rs(&pMF->m_db);
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
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 15);

	xls.SetRowHeight(4, 40);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);

	xls.SetCellMergedColumns(nCol + 3, nRow, 5);
	xls.SetCellMergedColumns(nCol + 3, nRow + 1, 5);

	xls.SetCellMergedColumns(nCol, nRow + 2, 6);
	xls.SetCellMergedColumns(nCol, nRow + 3, 6);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 3, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 3, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol, nRow + 2, _T("\x42\xC1O \x43\xC1O TR\x1EA2 L\x1EA0I TH\x45O KHO\x41"), FMT_TEXT | FMT_CENTER, true, 11);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("Khoa"));
	arrCol.Add(_T("T\x1ED5ng s\x1ED1 \x62\x1EC7nh nh\xE2n"));
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

		rs.GetValue(_T("deptname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("totalpatient"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("amount"), nAmount);
		nTotal += nAmount;
		tmpStr.Format(_T("%.2f"), nAmount);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		nRow++;
		rs.MoveNext();
	}	
	//FormatCurrency(nTotal, tmpStr);
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 10);

	tmpStr.Format(_T("%.2Lf"), nTotal);
	xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
	xls.Save(_T("Exports\\Bao Cao Tra Lai Theo Khoa.xls"));
} 
int CFMRefundbyDepartment::OnAddFMRefundbyDepartment(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMRefundbyDepartment::OnEditFMRefundbyDepartment(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMRefundbyDepartment::OnDeleteFMRefundbyDepartment(){
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
 		OnCancelFMRefundbyDepartment(); 
 	}
	return 0;
}
int CFMRefundbyDepartment::OnSaveFMRefundbyDepartment(){
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
 		//OnFMRefundbyDepartmentListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMRefundbyDepartment::OnCancelFMRefundbyDepartment(){
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
int CFMRefundbyDepartment::OnFMRefundbyDepartmentListLoadData(){
	return 0;
}

CString CFMRefundbyDepartment::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	szWhere.Empty();

	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_staff='%s' "), m_szClerkKey);

	szSQL.Format(_T(" SELECT deptid,") \
				_T("        sd_name AS deptname,") \
				_T("        COUNT(docno) AS totalpatient,") \
				_T("        SUM(amount) AS amount") \
				_T(" FROM") \
				_T(" (") \
				_T("  SELECT fi.hfe_deptid AS deptid,") \
				_T("         fi.hfe_docno AS docno,") \
				_T("         SUM(hfl_retprice * fe.hfe_quantity) AS amount") \
				_T("   FROM hms_fee_invoice fi") \
				_T("   LEFT JOIN hms_fee fe ON(fi.hfe_invoiceno=fe.hfe_invoiceno AND fi.hfe_docno=fe.hfe_docno)") \
				_T("   LEFT JOIN hms_fee_list ON(fe.hfe_itemid=hfl_feeid)") \
				_T("   WHERE fi.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')") \
				_T("         AND hfl_retprice > 0 AND fi.hfe_status='P'") \
				_T("   GROUP BY fi.hfe_deptid, fi.hfe_docno") \
				_T(" ) tbl") \
				_T(" LEFT JOIN sys_dept ON(sd_id=deptid)") \
				_T(" GROUP BY deptid, sd_name") \
				_T(" ORDER BY deptid"),
				m_szFromDate, m_szToDate, szWhere);

	return szSQL;
}