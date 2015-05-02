#include "stdafx.h"
#include "FMGeneralCollectbyDepartment.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralCollectbyDepartment *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMGeneralCollectbyDepartment *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMGeneralCollectbyDepartment* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralCollectbyDepartment *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnClerkAddNew();
}*/
static void _OnDepartmentSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMGeneralCollectbyDepartment* )pWnd)->OnDepartmentSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDepartmentSelendokFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnDepartmentSelendok();
}
/*static void _OnDepartmentSetfocusFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnDepartmentKillfocus();
}*/
/*static void _OnDepartmentKillfocusFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnDepartmentKillfocus();
}*/
static long _OnDepartmentLoadDataFnc(CWnd *pWnd){
	return ((CFMGeneralCollectbyDepartment *)pWnd)->OnDepartmentLoadData();
}
/*static void _OnDepartmentAddNewFnc(CWnd *pWnd){
	((CFMGeneralCollectbyDepartment *)pWnd)->OnDepartmentAddNew();
}*/
static void _OnDetailByPatientSelectFnc(CWnd *pWnd){
	 ((CFMGeneralCollectbyDepartment*)pWnd)->OnDetailByPatientSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMGeneralCollectbyDepartment *pVw = (CFMGeneralCollectbyDepartment *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMGeneralCollectbyDepartment *pVw = (CFMGeneralCollectbyDepartment *)pWnd;
	pVw->OnExportSelect();
}
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMGeneralCollectbyDepartment*)pWnd)->OnLockedSelect();
}
static int _OnAddFMGeneralCollectbyDepartmentFnc(CWnd *pWnd){
	 return ((CFMGeneralCollectbyDepartment*)pWnd)->OnAddFMGeneralCollectbyDepartment();
} 
static int _OnEditFMGeneralCollectbyDepartmentFnc(CWnd *pWnd){
	 return ((CFMGeneralCollectbyDepartment*)pWnd)->OnEditFMGeneralCollectbyDepartment();
} 
static int _OnDeleteFMGeneralCollectbyDepartmentFnc(CWnd *pWnd){
	 return ((CFMGeneralCollectbyDepartment*)pWnd)->OnDeleteFMGeneralCollectbyDepartment();
} 
static int _OnSaveFMGeneralCollectbyDepartmentFnc(CWnd *pWnd){
	 return ((CFMGeneralCollectbyDepartment*)pWnd)->OnSaveFMGeneralCollectbyDepartment();
} 
static int _OnCancelFMGeneralCollectbyDepartmentFnc(CWnd *pWnd){
	 return ((CFMGeneralCollectbyDepartment*)pWnd)->OnCancelFMGeneralCollectbyDepartment();
} 
CFMGeneralCollectbyDepartment::CFMGeneralCollectbyDepartment(CWnd *pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 190;
	SetDefaultValues();
}
CFMGeneralCollectbyDepartment::~CFMGeneralCollectbyDepartment(){
}
void CFMGeneralCollectbyDepartment::OnCreateComponents()
{
	/*m_wndGeneralCollectbyDepartment.Create(this, _T("General Collect by Department"), 5, 5, 490, 150);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 485, 85); 
	m_wndDepartmentLabel.Create(this, _T("Department"), 10, 90, 90, 115);
	m_wndDepartment.Create(this,95, 90, 485, 115); 
	m_wndDetailByPatient.Create(this, _T("Detail by Patient"), 370, 120, 485, 145);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 300, 155, 410, 180);
	m_wndExport.Create(this, _T("&Export"), 415, 155, 490, 180);*/

	m_wndGeneralCollectbyDepartment.Create(this, _T("General Collect by Department"), 5, 5, 430, 150);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 30, 300, 55);
	m_wndToDate.Create(this,305, 30, 425, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 425, 85); 
	m_wndDepartmentLabel.Create(this, _T("Department"), 10, 90, 90, 115);
	m_wndDepartment.Create(this,95, 90, 425, 115); 
	m_wndLocked.Create(this, _T("Locked"), 190, 120, 305, 145);
	m_wndDetailByPatient.Create(this, _T("Detail by Patient"), 310, 120, 425, 145);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 205, 155, 315, 180);
	m_wndExport.Create(this, _T("&Export"), 320, 155, 430, 180);

}
void CFMGeneralCollectbyDepartment::OnInitializeComponents()
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
	//m_wndDepartment.SetCheckValue(true);
	m_wndDepartment.LimitText(81);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Description"), CFMT_TEXT, 270);


	m_wndDepartment.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndDepartment.InsertColumn(1, _T("Name"), CFMT_TEXT, 280);

}
void CFMGeneralCollectbyDepartment::OnSetWindowEvents()
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
	m_wndDepartment.SetEvent(WE_SELENDOK, _OnDepartmentSelendokFnc);
	//m_wndDepartment.SetEvent(WE_SETFOCUS, _OnDepartmentSetfocusFnc);
	//m_wndDepartment.SetEvent(WE_KILLFOCUS, _OnDepartmentKillfocusFnc);
	m_wndDepartment.SetEvent(WE_SELCHANGE, _OnDepartmentSelectChangeFnc);
	m_wndDepartment.SetEvent(WE_LOADDATA, _OnDepartmentLoadDataFnc);
	//m_wndDepartment.SetEvent(WE_ADDNEW, _OnDepartmentAddNewFnc);
	m_wndDetailByPatient.SetEvent(WE_CLICK, _OnDetailByPatientSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);
	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CFMGeneralCollectbyDepartment::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_TextEx(pDX, m_wndDepartment.GetDlgCtrlID(), m_szDepartmentKey);
	DDX_Check(pDX, m_wndDetailByPatient.GetDlgCtrlID(), m_bDetailByPatient);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);
}
void CFMGeneralCollectbyDepartment::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMGeneralCollectbyDepartment::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMGeneralCollectbyDepartment::SetDefaultValues()
{

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_szDepartmentKey.Empty();
	m_bLocked=FALSE;
	m_bDetailByPatient=FALSE;

}
int CFMGeneralCollectbyDepartment::SetMode(int nMode){
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
 			EnableControls(TRUE); 
 			EnableButtons(FALSE, 3, 4, -1); 
 			break; 
 		case VM_NONE: 
 			EnableControls(TRUE); 
 			EnableButtons(TRUE, 0, 6, -1); 
 			SetDefaultValues(); 
 			break; 
 		}; 
 		UpdateData(FALSE); 
 		return nOldMode; 
}
/*void CFMGeneralCollectbyDepartment::OnFromDateChange(){
	
} */
/*void CFMGeneralCollectbyDepartment::OnFromDateSetfocus(){
	
} */
/*void CFMGeneralCollectbyDepartment::OnFromDateKillfocus(){
	
} */
int CFMGeneralCollectbyDepartment::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMGeneralCollectbyDepartment::OnToDateChange(){
	
} */
/*void CFMGeneralCollectbyDepartment::OnToDateSetfocus(){
	
} */
/*void CFMGeneralCollectbyDepartment::OnToDateKillfocus(){
	
} */
int CFMGeneralCollectbyDepartment::OnToDateCheckValue(){
	return 0;
} 
void CFMGeneralCollectbyDepartment::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMGeneralCollectbyDepartment::OnClerkSelendok(){
	 
}
/*void CFMGeneralCollectbyDepartment::OnClerkSetfocus(){
	
}*/
/*void CFMGeneralCollectbyDepartment::OnClerkKillfocus(){
	
}*/
long CFMGeneralCollectbyDepartment::OnClerkLoadData()
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
/*void CFMGeneralCollectbyDepartment::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMGeneralCollectbyDepartment::OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMGeneralCollectbyDepartment::OnDepartmentSelendok(){
	 
}
/*void CFMGeneralCollectbyDepartment::OnDepartmentSetfocus(){
	
}*/
/*void CFMGeneralCollectbyDepartment::OnDepartmentKillfocus(){
	
}*/
long CFMGeneralCollectbyDepartment::OnDepartmentLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndDepartment.IsSearchKey() && !m_szDepartmentKey.IsEmpty())
	{
		szWhere.Format(_T(" AND sd_id='%s' "), m_szDepartmentKey);
	}
	szSQL.Format(_T("SELECT sd_id as id, sd_name as description ") \
		         _T("FROM sys_dept ") \
				 _T("WHERE sd_type IN('KB','DT') %s ORDER BY sd_id"),
				 szWhere);
	m_wndDepartment.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDepartment.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CFMGeneralCollectbyDepartment::OnDepartmentAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMGeneralCollectbyDepartment::OnDetailByPatientSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}
void CFMGeneralCollectbyDepartment::OnLockedSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMGeneralCollectbyDepartment::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp, szSysDate;
	int nIdx = 0;
	double nAmount = 0, nTemp = 0;
	long double nTotal = 0;

	if (m_bDetailByPatient)
	{
		if (!rpt.Init(_T("Reports/HMS/HF_CHITIETTHUTIENTHEOKHOA.rpt")))
			return;
	}
	else
	{
		if (!rpt.Init(_T("Reports/HMS/HF_BANGTHONGKETHUTIENTHEOKHOA.rpt")))
			return;
	}

	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	/*if (!rpt.Init(_T("Reports/HMS/HF_BANGTHONGKETHUTIENTHEOKHOA.RPT")) ||
		!rpt.Init(_T("Reports/HMS/HF_CHITIETTHUTIENTHEOKHOA.RPT")))
		return;*/
	//Header

	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection *rptDetail;

	if (!m_bDetailByPatient)
	{
		while (!rs.IsEOF())
		{
			rptDetail = rpt.AddDetail();

			nIdx++;

			rs.GetValue(_T("deptid"), szTemp);

			tmpStr.Format(_T("%d"), nIdx);
			rptDetail->SetValue(_T("1"), tmpStr);

			rs.GetValue(_T("deptname"), tmpStr);

			if (szTemp == _T("X"))
				tmpStr = _T("Thu ph\xED kh\xE1\x63");
			else if (szTemp == _T("NA"))
				tmpStr = _T("\x43h\x1B0\x61 thi\x1EBFt l\x1EADp kho\x61 th\x1EF1\x63 hi\x1EC7n");

			rptDetail->SetValue(_T("2"), tmpStr);

			rs.GetValue(_T("total"), tmpStr);
			//FormatCurrency(nTemp, tmpStr);
			rptDetail->SetValue(_T("3"), tmpStr);

			rs.GetValue(_T("exam_fee"), nTemp);
			FormatCurrency(nTemp, tmpStr);
			rptDetail->SetValue(_T("4"), tmpStr);

			rs.GetValue(_T("inpatient_fee"), nTemp);
			FormatCurrency(nTemp, tmpStr);
			rptDetail->SetValue(_T("5"), tmpStr);

			rs.GetValue(_T("total_amount"), nAmount);
			FormatCurrency(nAmount, tmpStr);
			rptDetail->SetValue(_T("6"), tmpStr);
			nTotal += nAmount;
			rs.MoveNext();
		}
		//Vertical total
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		FormatCurrency(nTotal, tmpStr);
		rptDetail->SetValue(_T("s6"), tmpStr);
		szSysDate = pMF->GetSysDate();
		tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
		rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	}
	else
	{
		CString szOldDept, szNewDept;
		CReportSection *rptSector;

		double nGrpAmt = 0;
		while (!rs.IsEOF())
		{
			rs.GetValue(_T("deptid"), szNewDept);

			if (szNewDept != szOldDept)
			{
				if (nGrpAmt > 0)
				{
					rptSector = rpt.AddDetail(rpt.GetGroupFooter(1));
					FormatCurrency(nGrpAmt, tmpStr);
					rptSector->SetValue(_T("s6"), tmpStr);
				}

				if (szNewDept != _T("X") && szNewDept != _T("NA"))
					tmpStr = pMF->GetDepartmentName(szNewDept);
				else if (szNewDept == _T("X"))
					tmpStr = _T("Thu ph\xED kh\xE1\x63");
				else if (szNewDept == _T("NA"))
					tmpStr = _T("\x43h\x1B0\x61 thi\x1EBFt l\x1EADp kho\x61 th\x1EF1\x63 hi\x1EC7n");

				rptSector = rpt.AddDetail(rpt.GetGroupHeader(0));
				rptSector->SetValue(_T("GroupName"), tmpStr);
				szOldDept = szNewDept;
				nIdx = 0;
				nTotal += nGrpAmt;
				nGrpAmt = 0;
			}
			rptDetail = rpt.AddDetail();

			nIdx++;
			tmpStr.Format(_T("%d"), nIdx);
			rptDetail->SetValue(_T("1"), tmpStr);

			rs.GetValue(_T("docno"), tmpStr);
			rptDetail->SetValue(_T("2"), tmpStr);

			rs.GetValue(_T("pname"), tmpStr);
			rptDetail->SetValue(_T("3"), tmpStr);

			rs.GetValue(_T("exam_fee"), nTemp);
			FormatCurrency(nTemp, tmpStr);
			rptDetail->SetValue(_T("4"), tmpStr);

			rs.GetValue(_T("inpatient_fee"), nTemp);
			FormatCurrency(nTemp, tmpStr);
			rptDetail->SetValue(_T("5"), tmpStr);

			rs.GetValue(_T("total_amount"), nAmount);
			FormatCurrency(nAmount, tmpStr);
			rptDetail->SetValue(_T("6"), tmpStr);
			nGrpAmt += nAmount;
			rs.MoveNext();
		}
		nTotal += nGrpAmt;
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
		FormatCurrency(nGrpAmt, tmpStr);
		rptDetail->SetValue(_T("s6"), tmpStr);
		FormatCurrency(nTotal, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("t6"), tmpStr);
		szSysDate = pMF->GetSysDate();
		tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
		rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	}
	rpt.PrintPreview();

} 
void CFMGeneralCollectbyDepartment::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	if (m_bDetailByPatient)
		OnExportDetail();
	else 
		OnExportGeneral();
} 
void CFMGeneralCollectbyDepartment::OnExportDetail()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0, nIdx = 0;
	double nAmount = 0, nTemp = 0;
	long double nGrpAmount = 0, nTotal = 0;

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
	xls.SetColumnWidth(3, 12);
	xls.SetColumnWidth(4, 12);
	xls.SetColumnWidth(5, 12);

	xls.SetRowHeight(4, 40);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);

	xls.SetCellMergedColumns(nCol + 3, nRow, 4);
	xls.SetCellMergedColumns(nCol + 3, nRow + 1, 4);

	xls.SetCellMergedColumns(nCol, nRow + 2, 6);
	xls.SetCellMergedColumns(nCol, nRow + 3, 6);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 3, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 3, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol + 3, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 3, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol, nRow + 2, _T("\x43HI TI\x1EBET THU TI\x1EC0N TH\x45O KHO\x41"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("S\x1ED1 h\x1ED3 s\x1A1"));
	arrCol.Add(_T("T\xEAn \x62\x1EC7nh nh\xE2n"));
	arrCol.Add(_T("Thu Ph\xF2ng kh\xE1m"));
	arrCol.Add(_T("Thu n\x1ED9i tr\xFA"));
	arrCol.Add(_T("T\x1ED5ng thu"));
	nRow = 4;

	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol + i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10); 
	}

	nRow = 5;
	CString szOldDept, szNewDept;

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("deptid"), szNewDept);
		if (szNewDept != szOldDept)
		{
			if (nGrpAmount > 0)
			{
				xls.SetCellMergedColumns(nCol, nRow, 5);
				xls.SetCellText(nCol, nRow, _T("T\x1ED5ng nh\xF3m:"), FMT_TEXT | FMT_CENTER, true, 10);
				//FormatCurrency(nGrpAmount, tmpStr);
				tmpStr.Format(_T("%.2Lf"), nGrpAmount);
				xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
				nRow++;
			}

			szOldDept = szNewDept;
			nIdx = 0;
			nTotal += nGrpAmount;
			nGrpAmount = 0;

			if (szNewDept != _T("X") && szNewDept != _T("NA"))
				tmpStr = pMF->GetDepartmentName(szNewDept);
			else if (szNewDept == _T("X"))
				tmpStr = _T("Thu ph\xED kh\xE1\x63");
			else if (szNewDept == _T("NA"))
				tmpStr = _T("\x43h\x1B0\x61 thi\x1EBFt l\x1EADp kho\x61 th\x1EF1\x63 hi\x1EC7n");

			xls.SetCellMergedColumns(nCol, nRow, 6);
			xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_WRAPING, true, 10);
			nRow++;
		}

		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("exam_fee"), nTemp);
		//FormatCurrency(nTemp, tmpStr);
		tmpStr.Format(_T("%.2f"), nTemp);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("inpatient_fee"), nTemp);
		//FormatCurrency(nTemp, tmpStr);
		tmpStr.Format(_T("%.2f"), nTemp);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("total_amount"), nAmount);
		nGrpAmount += nAmount;
		//FormatCurrency(nAmount, tmpStr);
		tmpStr.Format(_T("%.2f"), nAmount);
		xls.SetCellText(nCol+5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		nRow++;
		rs.MoveNext();
	}
	if (nGrpAmount > 0)
	{
		xls.SetCellMergedColumns(nCol, nRow, 5);
		xls.SetCellText(nCol, nRow, _T("T\x1ED5ng nh\xF3m:"), FMT_TEXT | FMT_CENTER, true, 10);
		//FormatCurrency(nGrpAmount, tmpStr);
		tmpStr.Format(_T("%.2Lf"), nGrpAmount);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
		nTotal += nGrpAmount;
		nRow++;
	}

	//FormatCurrency(nTotal, tmpStr);
	xls.SetCellMergedColumns(nCol, nRow, 5);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 10);
	tmpStr.Format(_T("%.2Lf"), nTotal);
	xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
	xls.Save(_T("Exports\\Chi Tiet Thu Tien Theo Khoa.xls"));	
}
void CFMGeneralCollectbyDepartment::OnExportGeneral()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp;
	int nCol = 0, nRow = 0, nIdx = 0;
	double nAmount = 0, nTemp = 0;
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
	xls.SetColumnWidth(1, 50);
	xls.SetColumnWidth(2, 13);
	xls.SetColumnWidth(3, 12);
	xls.SetColumnWidth(4, 12);
	xls.SetColumnWidth(5, 12);

	xls.SetRowHeight(4, 40);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellMergedColumns(nCol, nRow + 1, 2);

	xls.SetCellMergedColumns(nCol + 2, nRow, 4);
	xls.SetCellMergedColumns(nCol + 2, nRow + 1, 4);

	xls.SetCellMergedColumns(nCol, nRow + 2, 6);
	xls.SetCellMergedColumns(nCol, nRow + 3, 6);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 2, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 2, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol, nRow + 2, _T("T\x1ED4NG H\x1EE2P THU TI\x1EC0N TH\x45O KHO\x41"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("Khoa"));
	arrCol.Add(_T("S\x1ED1 \x62\x1EC7nh nh\xE2n"));
	arrCol.Add(_T("Thu Ph\xF2ng kh\xE1m"));
	arrCol.Add(_T("Thu n\x1ED9i tr\xFA"));
	arrCol.Add(_T("T\x1ED5ng thu"));

	nRow = 4;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10); 
	}
	nRow = 5;

	while (!rs.IsEOF())
	{
		nIdx++;

		rs.GetValue(_T("deptid"), szTemp);

		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("deptname"), tmpStr);

		if (szTemp == _T("X"))
			tmpStr = _T("Thu ph\xED kh\xE1\x63");
		else if (szTemp == _T("NA"))
			tmpStr = _T("\x43h\x1B0\x61 thi\x1EBFt l\x1EADp kho\x61 th\x1EF1\x63 hi\x1EC7n");

		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);

		rs.GetValue(_T("total"), tmpStr);
		//FormatCurrency(nTemp, tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);

		rs.GetValue(_T("exam_fee"), nTemp);
		//FormatCurrency(nTemp, tmpStr);
		tmpStr.Format(_T("%.2f"), nTemp);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("inpatient_fee"), nTemp);
		//FormatCurrency(nTemp, tmpStr);
		tmpStr.Format(_T("%.2f"), nTemp);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		rs.GetValue(_T("total_amount"), nAmount);
		nTotal += nAmount;
		//FormatCurrency(nAmount, tmpStr);
		tmpStr.Format(_T("%.2f"), nAmount);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);
		nRow++;
		rs.MoveNext();
	}
	//FormatCurrency(nTotal, tmpStr);tce
	xls.SetCellMergedColumns(nCol, nRow, 5);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng ti\x1EC1n"), FMT_TEXT | FMT_CENTER, true, 10);
	tmpStr.Format(_T("%.2Lf"), nTotal);
	xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
	xls.Save(_T("Exports\\Tong Hop Thu Tien Theo Khoa.xls"));	
}
int CFMGeneralCollectbyDepartment::OnAddFMGeneralCollectbyDepartment(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMGeneralCollectbyDepartment::OnEditFMGeneralCollectbyDepartment(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMGeneralCollectbyDepartment::OnDeleteFMGeneralCollectbyDepartment(){
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
 		OnCancelFMGeneralCollectbyDepartment(); 
 	}
	return 0;
}
int CFMGeneralCollectbyDepartment::OnSaveFMGeneralCollectbyDepartment(){
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
 		//OnFMGeneralCollectbyDepartmentListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMGeneralCollectbyDepartment::OnCancelFMGeneralCollectbyDepartment()
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
int CFMGeneralCollectbyDepartment::OnFMGeneralCollectbyDepartmentListLoadData()
{
	return 0;
}

CString CFMGeneralCollectbyDepartment::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere;
	szWhere.Empty();
	
	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_staff='%s' "), m_szClerkKey);
	if (!m_szDepartmentKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_deptid='%s' "), m_szDepartmentKey);

	if (!m_bLocked)
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_locked<>'Y' "));
		szWhere.AppendFormat(_T(" AND fi.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}
	else
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_locked='Y' "));
		szWhere.AppendFormat(_T(" AND fi.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}

	if (!m_bDetailByPatient)
	{
		szSQL.Format(_T(" SELECT CASE WHEN deptid IN('C1.1','C1.2','C1.3') THEN 'KB'") \
					_T("             WHEN deptid IN('KD') THEN 'KA'") \
					_T("             WHEN deptid IN('X') THEN 'XX'") \
					_T("             WHEN deptid IN('NA') THEN 'XZ' ELSE 'DT' END AS depttype,") \
					_T("        deptid,") \
					_T("        sd_name AS deptname,") \
					_T("        COUNT(distinct docno) AS total,") \
					_T("        SUM(examfee) AS exam_fee,") \
					_T("        SUM(inpatfee) AS inpatient_fee,") \
					_T("        SUM(examfee+inpatfee) AS total_amount") \
					_T(" FROM") \
					_T(" (") \
					_T("  SELECT CASE WHEN fe.hfe_itemid IN('D0000006') THEN cast('C1.3' as nvarchar2(140))") \
					_T("              WHEN fe.hfe_group IN('D0000') AND fe.hfe_itemid NOT IN('D0000006') THEN cast('C1.1' as nvarchar2(140))") \
					_T("              WHEN substr(fe.hfe_group,1,2) IN('A1') THEN cast('KD' as nvarchar2(140))") \
					_T("              WHEN fe.hfe_group='F0000' THEN cast('X' as nvarchar2(140))") \
					_T("              WHEN hfl_deptid IS NULL THEN cast('NA' as nvarchar2(140)) ELSE hfl_deptid END AS deptid,") \
					_T("         fe.hfe_docno AS docno,") \
					_T("         CASE WHEN fi.hfe_class='E' THEN SUM(fe.hfe_patpaid) ELSE 0 END AS examfee,") \
					_T("         CASE WHEN fi.hfe_class<>'E' THEN SUM(fe.hfe_patpaid) ELSE 0 END AS inpatfee") \
					_T("  FROM hms_fee_invoice fi") \
					_T("  LEFT JOIN hms_fee fe ON(fi.hfe_invoiceno=fe.hfe_invoiceno AND fi.hfe_docno=fe.hfe_docno)") \
					_T("  LEFT JOIN hms_fee_list ON(fe.hfe_itemid=hfl_feeid)") \
					_T("  WHERE fe.hfe_type IN('E','O','P','T','X','D') AND fi.hfe_status ='P' %s") \
					_T("  GROUP BY hfl_deptid, fe.hfe_itemid, fe.hfe_docno, fi.hfe_class, fe.hfe_group") \
					_T("  HAVING SUM(fe.hfe_patpaid) > 0") \
					_T(" ) tbl") \
					_T(" LEFT JOIN sys_dept ON(sd_id=deptid)") \
					_T(" GROUP BY deptid, sd_name") \
					_T(" ORDER BY depttype, deptid"),
					szWhere);
	}
	else
	{
		szSQL.Format(_T(" SELECT distinct docno AS docno,") \
					_T("        CASE WHEN deptid IN('C1.1','C1.2','C1.3') THEN 'KB'") \
					_T("             WHEN deptid IN('KD') THEN 'KA'") \
					_T("             WHEN deptid IN('X') THEN 'XX'") \
					_T("             WHEN deptid IN('NA') THEN 'XZ' ELSE 'DT' END AS depttype,") \
					_T("        deptid,") \
					_T("        sd_name AS deptname,") \
					_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) AS pname,") \
					_T("        SUM(examfee) AS exam_fee,") \
					_T("        SUM(inpatfee) AS inpatient_fee,") \
					_T("        SUM(examfee+inpatfee) AS total_amount") \
					_T(" FROM") \
					_T(" (") \
					_T("  SELECT CASE WHEN fe.hfe_itemid IN('D0000006') THEN cast('C1.3' as nvarchar2(140))") \
					_T("              WHEN fe.hfe_group IN('D0000') AND fe.hfe_itemid NOT IN('D0000006') THEN cast('C1.1' as nvarchar2(140))") \
					_T("              WHEN substr(fe.hfe_group,1,2) IN('A1') THEN cast('KD' as nvarchar2(140))") \
					_T("              WHEN fe.hfe_group='F0000' THEN cast('X' as nvarchar2(140))") \
					_T("              WHEN hfl_deptid IS NULL THEN cast('NA' as nvarchar2(140)) ELSE hfl_deptid END AS deptid,") \
					_T("         fe.hfe_docno AS docno,") \
					_T("         CASE WHEN fi.hfe_class='E' THEN SUM(fe.hfe_patpaid) ELSE 0 END AS examfee,") \
					_T("         CASE WHEN fi.hfe_class<>'E' THEN SUM(fe.hfe_patpaid) ELSE 0 END AS inpatfee") \
					_T("  FROM hms_fee_invoice fi") \
					_T("  LEFT JOIN hms_fee fe ON(fi.hfe_invoiceno=fe.hfe_invoiceno AND fi.hfe_docno=fe.hfe_docno)") \
					_T("  LEFT JOIN hms_fee_list ON(fe.hfe_itemid=hfl_feeid)") \
					_T("  WHERE fe.hfe_type IN('E','O','P','T','X','D') AND fi.hfe_status='P'") \
					_T("        AND fi.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') %s") \
					_T("  GROUP BY hfl_deptid, fe.hfe_itemid, fe.hfe_docno, fi.hfe_class, fe.hfe_group") \
					_T("  HAVING SUM(fe.hfe_patpaid) > 0") \
					_T(" ) tbl") \
					_T(" LEFT JOIN sys_dept ON(sd_id=deptid)") \
					_T(" LEFT JOIN hms_doc ON(hd_docno=docno)") \
					_T(" LEFT JOIN hms_patient ON(hp_patientno=hd_patientno)") \
					_T(" GROUP BY deptid, sd_name, docno, hp_surname, hp_midname, hp_firstname") \
					_T(" ORDER BY depttype, deptid, pname"),
					m_szFromDate, m_szToDate, szWhere);
	}
	//QueryOpener(szSQL);
	return szSQL;
}