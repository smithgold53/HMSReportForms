#include "stdafx.h"
#include "FMGeneralReportbyClerk.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMGeneralReportbyClerk *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMGeneralReportbyClerk *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMGeneralReportbyClerk *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralReportbyClerk *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMGeneralReportbyClerk *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMGeneralReportbyClerk *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMGeneralReportbyClerk *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralReportbyClerk *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMGeneralReportbyClerk* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMGeneralReportbyClerk *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMGeneralReportbyClerk *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMGeneralReportbyClerk *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralReportbyClerk *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMGeneralReportbyClerk *)pWnd)->OnClerkAddNew();
}*/
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMGeneralReportbyClerk*)pWnd)->OnLockedSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMGeneralReportbyClerk *pVw = (CFMGeneralReportbyClerk *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMGeneralReportbyClerk *pVw = (CFMGeneralReportbyClerk *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddFMGeneralReportbyClerkFnc(CWnd *pWnd){
	 return ((CFMGeneralReportbyClerk*)pWnd)->OnAddFMGeneralReportbyClerk();
} 
static int _OnEditFMGeneralReportbyClerkFnc(CWnd *pWnd){
	 return ((CFMGeneralReportbyClerk*)pWnd)->OnEditFMGeneralReportbyClerk();
} 
static int _OnDeleteFMGeneralReportbyClerkFnc(CWnd *pWnd){
	 return ((CFMGeneralReportbyClerk*)pWnd)->OnDeleteFMGeneralReportbyClerk();
} 
static int _OnSaveFMGeneralReportbyClerkFnc(CWnd *pWnd){
	 return ((CFMGeneralReportbyClerk*)pWnd)->OnSaveFMGeneralReportbyClerk();
} 
static int _OnCancelFMGeneralReportbyClerkFnc(CWnd *pWnd){
	 return ((CFMGeneralReportbyClerk*)pWnd)->OnCancelFMGeneralReportbyClerk();
} 
CFMGeneralReportbyClerk::CFMGeneralReportbyClerk(CWnd *pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 165;
	SetDefaultValues();
}
CFMGeneralReportbyClerk::~CFMGeneralReportbyClerk(){
}
void CFMGeneralReportbyClerk::OnCreateComponents(){
	m_wndGeneralReportbyClerk.Create(this, _T("General Report by Clerk"), 5, 5, 430, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 30, 300, 55);
	m_wndToDate.Create(this,305, 30, 425, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 425, 85); 
	m_wndLocked.Create(this, _T("Locked"), 5, 95, 105, 120);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 240, 95, 350, 120);
	m_wndExport.Create(this, _T("&Export"), 355, 95, 430, 120);

}

void CFMGeneralReportbyClerk::OnInitializeComponents()
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

}
void CFMGeneralReportbyClerk::OnSetWindowEvents(){
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
	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CFMGeneralReportbyClerk::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);

}
void CFMGeneralReportbyClerk::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMGeneralReportbyClerk::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMGeneralReportbyClerk::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bLocked=FALSE;

}
int CFMGeneralReportbyClerk::SetMode(int nMode){
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
/*void CFMGeneralReportbyClerk::OnFromDateChange(){
	
} */
/*void CFMGeneralReportbyClerk::OnFromDateSetfocus(){
	
} */
/*void CFMGeneralReportbyClerk::OnFromDateKillfocus(){
	
} */
int CFMGeneralReportbyClerk::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMGeneralReportbyClerk::OnToDateChange(){
	
} */
/*void CFMGeneralReportbyClerk::OnToDateSetfocus(){
	
} */
/*void CFMGeneralReportbyClerk::OnToDateKillfocus(){
	
} */
int CFMGeneralReportbyClerk::OnToDateCheckValue(){
	return 0;
} 
void CFMGeneralReportbyClerk::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMGeneralReportbyClerk::OnClerkSelendok(){
	 
}
/*void CFMGeneralReportbyClerk::OnClerkSetfocus(){
	
}*/
/*void CFMGeneralReportbyClerk::OnClerkKillfocus(){
	
}*/
long CFMGeneralReportbyClerk::OnClerkLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
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
/*void CFMGeneralReportbyClerk::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
	void CFMGeneralReportbyClerk::OnLockedSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMGeneralReportbyClerk::OnPrintPreviewSelect()
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
	if (!rpt.Init(_T("Reports/HMS/HF_TONGHOPTHEONGUOITHU.rpt")))
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
		rptDetail = rpt.AddDetail();

		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("userid"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("username"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("amount"), nAmount);
		nTotal += nAmount;
		FormatCurrency(nAmount, tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.MoveNext();
	}
	FormatCurrency(nTotal, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("s4"), tmpStr);
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
	
} 
void CFMGeneralReportbyClerk::OnExportSelect()
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
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 25);
	xls.SetColumnWidth(3, 15);

	xls.SetRowHeight(4, 40);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);

	xls.SetCellMergedColumns(nCol + 3, nRow, 4);
	xls.SetCellMergedColumns(nCol + 3, nRow + 1, 4);

	xls.SetCellMergedColumns(nCol, nRow + 2, 4);
	xls.SetCellMergedColumns(nCol, nRow + 3, 4);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 3, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 3, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol, nRow + 2, _T("T\x1ED4NG H\x1EE2P TH\x45O NG\x1AF\x1EDCI THU"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("M\xE3"));
	arrCol.Add(_T("T\xEAn ng\x1B0\x1EDDi thu"));
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

		rs.GetValue(_T("userid"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("username"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("amount"), nAmount);
		tmpStr.Format(_T("%.2f"), nAmount);
		nTotal += nAmount;
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);
		nRow++;
		rs.MoveNext();
	}
	
	xls.SetCellText(nCol + 2, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 10);
	tmpStr.Format(_T("%.2Lf"), nTotal);
	xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
	xls.Save(_T("Exports\\Tong Hop Theo Nguoi Thu.xls"));
	
} 
int CFMGeneralReportbyClerk::OnAddFMGeneralReportbyClerk(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMGeneralReportbyClerk::OnEditFMGeneralReportbyClerk(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMGeneralReportbyClerk::OnDeleteFMGeneralReportbyClerk(){
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
 		OnCancelFMGeneralReportbyClerk(); 
 	}
	return 0;
}
int CFMGeneralReportbyClerk::OnSaveFMGeneralReportbyClerk(){
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
 		//OnFMGeneralReportbyClerkListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMGeneralReportbyClerk::OnCancelFMGeneralReportbyClerk()
{
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
int CFMGeneralReportbyClerk::OnFMGeneralReportbyClerkListLoadData(){
	return 0;
}

CString CFMGeneralReportbyClerk::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	szWhere.Empty();

	if (m_bLocked)
	{
		szWhere.AppendFormat(_T(" AND hfe_locked='Y' "));
		szWhere.AppendFormat(_T(" AND hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}
	else
	{
		szWhere.AppendFormat(_T(" AND hfe_locked<>'Y' "));
		szWhere.AppendFormat(_T(" AND hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}

	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_staff='%s' "), m_szClerkKey);

	szSQL.Format(_T(" SELECT hfe_staff as userid,") \
				_T("        (select distinct su_name from sys_user where su_userid=hfe_staff) as username,") \
				_T("        sum(hfe_patpaid) as amount") \
				_T(" FROM hms_fee_invoice ") \
				_T(" WHERE hfe_status='P' %s") \
				_T(" GROUP BY hfe_staff") \
				_T(" HAVING sum(hfe_patpaid) > 0") \
				_T(" ORDER BY hfe_staff"),
				szWhere);
	return szSQL;
}