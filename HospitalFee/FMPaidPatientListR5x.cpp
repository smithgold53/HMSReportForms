#include "stdafx.h"
#include "FMPaidPatientListR5x.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMPaidPatientListR5x *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMPaidPatientListR5x *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMPaidPatientListR5x *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMPaidPatientListR5x *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMPaidPatientListR5x *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMPaidPatientListR5x *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMPaidPatientListR5x *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMPaidPatientListR5x *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMPaidPatientListR5x* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMPaidPatientListR5x *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMPaidPatientListR5x *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMPaidPatientListR5x *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMPaidPatientListR5x *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMPaidPatientListR5x *)pWnd)->OnClerkAddNew();
}*/
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMPaidPatientListR5x*)pWnd)->OnLockedSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMPaidPatientListR5x *pVw = (CFMPaidPatientListR5x *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMPaidPatientListR5x *pVw = (CFMPaidPatientListR5x *)pWnd;
	pVw->OnExportSelect();
} 
CFMPaidPatientListR5x::CFMPaidPatientListR5x(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CFMPaidPatientListR5x::~CFMPaidPatientListR5x(){
}
void CFMPaidPatientListR5x::OnCreateComponents(){
	m_wndPaidPatientList.Create(this, _T("Paid Patient List"), 5, 5, 430, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 30, 300, 55);
	m_wndToDate.Create(this,305, 30, 425, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 425, 85); 
	m_wndLocked.Create(this, _T("Locked"), 5, 95, 90, 120);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 225, 95, 325, 120);
	m_wndExport.Create(this, _T("&Export"), 330, 95, 430, 120);
	m_wndExport.EnableWindow(false);

}
void CFMPaidPatientListR5x::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(35);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 270);
	
}

void CFMPaidPatientListR5x::OnSetWindowEvents(){
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

}
void CFMPaidPatientListR5x::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);

}
void CFMPaidPatientListR5x::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bLocked=FALSE;

}

int CFMPaidPatientListR5x::SetMode(int nMode){
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

/*void CFMPaidPatientListR5x::OnFromDateChange(){
	
} */
/*void CFMPaidPatientListR5x::OnFromDateSetfocus(){
	
} */
/*void CFMPaidPatientListR5x::OnFromDateKillfocus(){
	
} */
int CFMPaidPatientListR5x::OnFromDateCheckValue(){
	return 0;
}
 
/*void CFMPaidPatientListR5x::OnToDateChange(){
	
} */
/*void CFMPaidPatientListR5x::OnToDateSetfocus(){
	
} */
/*void CFMPaidPatientListR5x::OnToDateKillfocus(){
	
} */
int CFMPaidPatientListR5x::OnToDateCheckValue(){
	return 0;
}
 
void CFMPaidPatientListR5x::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
void CFMPaidPatientListR5x::OnClerkSelendok(){
	 
}

/*void CFMPaidPatientListR5x::OnClerkSetfocus(){
	
}*/
/*void CFMPaidPatientListR5x::OnClerkKillfocus(){
	
}*/
long CFMPaidPatientListR5x::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();
	
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	}

	m_wndClerk.DeleteAllItems(); 
	szSQL.Format(_T("SELECT su_userid as userid, su_name as name ") \
		         _T("FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);

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

/*void CFMPaidPatientListR5x::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMPaidPatientListR5x::OnLockedSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
}

void CFMPaidPatientListR5x::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	/*Declaration Section*/
	CRecord rs(&pMF->m_db);
	CReport rpt;
	CReportSection* rptDetail = NULL,* rptHeader = NULL, * rptGroup = NULL;
	double nTmp = 0;
	int nIdx = 1;
	CString szSQL, tmpStr, tmpStr2, szReportName, szDate, szOldLev1, szNewLev1;
	long double nTotal = 0, nGroupTotal1 = 0;
	szReportName = _T("Reports\\HMS\\HF_DANHSACHBENHNHANNOPTIEN.RPT");
	if (!rpt.Init(szReportName))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	rptHeader = rpt.GetReportHeader();
	if (rptHeader)
	{
		rptHeader->SetValue(_T("HealthService"), pMF->m_szHealthService);
		rptHeader->SetValue(_T("HospitalName"), pMF->m_szHospitalName);
		tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), 
		CDate::Convert(m_szToDate), yyyymmdd, ddmmyyyy);
		rptHeader->SetValue(_T("ReportDate"), tmpStr);
	}
	while (!rs.IsEOF())
	{
		/*Group Seperation*/
		rs.GetValue(_T("room_name"), szNewLev1);
		if (szOldLev1 != szNewLev1)
		{
			if (nGroupTotal1 > 0)
			{
				tmpStr.Format(_T("%f"), nGroupTotal1);
				rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));
				rptGroup->SetValue(_T("GroupName"), _T("T\x1ED5ng nh\xF3m"));
				rptGroup->SetValue(_T("SumGroupName"), tmpStr);
				nTotal+= nGroupTotal1;
				nGroupTotal1 = 0;
			}
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));
			if (rptGroup)
				rptGroup->SetValue(_T("GroupName"), szNewLev1);
			szOldLev1 = szNewLev1;
		}
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
			rptDetail->SetValue(_T("1"), int2str(nIdx++));
			rptDetail->SetValue(_T("2"), rs.GetValue(_T("patient_name")));
			rptDetail->SetValue(_T("3"), rs.GetValue(_T("doc_no")));
			//rptDetail->SetValue(_T("5"), rs.GetValue(_T("room_name")));
			rs.GetValue(_T("pat_paid"), nTmp);
			nGroupTotal1+= nTmp;
			FormatCurrency(nTmp, tmpStr);
			rptDetail->SetValue(_T("5"), tmpStr);
		}
		rs.MoveNext();
	}
	if (nGroupTotal1 > 0)
	{
		rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));
		rptGroup->SetValue(_T("GroupName"), _T("T\x1ED5ng nh\xF3m"));
		FormatCurrency(nGroupTotal1, tmpStr);
		rptGroup->SetValue(_T("SumGroupName"), tmpStr);
		nTotal+= nGroupTotal1;
	}
	if (nTotal > 0)
	{
		rpt.GetReportFooter()->SetValue(_T("TotalAmount"), _T("T\x1ED5ng \x63\x1ED9ng"));
		FormatCurrency(nTotal, tmpStr);
		rpt.GetReportFooter()->SetValue(_T("SumTotalAmount"), tmpStr);
	}
	tmpStr2.Format(_T("%f"), nTotal);
	MoneyToString(tmpStr2, tmpStr);
	tmpStr += _T(" \x111\x1ED3ng.");
	rpt.GetReportFooter()->SetValue(_T("SumInWord"), tmpStr);
	tmpStr = pMF->GetSysDate();
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), tmpStr.Right(2), tmpStr.Mid(5,2), tmpStr.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
}
 
void CFMPaidPatientListR5x::OnExportSelect(){

}
 
CString CFMPaidPatientListR5x::GetQueryString(){
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (m_bLocked)
		szWhere.AppendFormat(_T(" AND hfe_locked = 'Y'"));
	else
		szWhere.AppendFormat(_T(" AND NVL(hfe_locked, 'X') <> 'Y'"));
	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_staff = '%s'"), m_szClerkKey);
	szSQL.Format(_T(" SELECT   get_patientname(doc_no) patient_name, ") \
				_T("		   doc_no, ") \
				_T("           room_id, ") \
				_T("		   hrl_name room_name,") \
				_T("           SUM(pat_paid) pat_paid ") \
				_T(" FROM      (SELECT    hfe_docno doc_no, ") \
				_T("                      hfe_invoiceno invoice_no, ") \
				_T("                      ho_roomid room_id, ") \
				_T("                      hfe_patpaid pat_paid ") \
				_T("            FROM      hms_fee ") \
				_T("            LEFT JOIN hms_operation ON ( hfe_docno = ho_docno AND ho_idx = hfe_orderid ) ") \
				_T("            WHERE     hfe_type = 'O' AND ho_deptid = 'C1.1' AND ho_roomid IN ( 50, 54 ) AND hfe_status IN ( 'P', 'R' ) ") \
				_T("            UNION ALL ") \
				_T("            SELECT    hfe_docno, ") \
				_T("                      hfe_invoiceno, ") \
				_T("                      hpc_roomid, ") \
				_T("                      hfe_patpaid ") \
				_T("            FROM      hms_fee ") \
				_T("            LEFT JOIN hmsv_paraclinic ON ( hfe_docno = hpc_docno AND hfe_orderid = hpc_orderid ) ") \
				_T("            WHERE     hfe_type IN ( 'P', 'T' ) AND hpc_deptid = 'C1.1' AND hpc_roomid IN ( 50, 54 ) AND hfe_status IN ( 'P', 'R' ) ") \
				_T("            UNION ALL ") \
				_T("            SELECT    hfe_docno, ") \
				_T("                      hfe_invoiceno, ") \
				_T("                      he_roomid, ") \
				_T("                      hfe_patpaid ") \
				_T("            FROM      hms_fee ") \
				_T("            LEFT JOIN hms_exam ON ( he_docno = hfe_docno AND he_receptidx = hfe_orderid ) ") \
				_T("            WHERE     hfe_type = 'E' AND he_deptid = 'C1.1' AND he_roomid IN ( 50, 54 ) AND hfe_status IN ( 'P', 'R' ) ") \
				_T("            UNION ALL ") \
				_T("            SELECT    hfe_docno, ") \
				_T("                      hfe_invoiceno, ") \
				_T("                      hpo_roomid, ") \
				_T("                      hfe_patpaid ") \
				_T("            FROM      hms_fee ") \
				_T("            LEFT JOIN hms_pharmaorder_view ON ( hfe_docno = hpo_docno AND hpo_orderid = hfe_orderid ) ") \
				_T("            WHERE     hfe_type IN ( 'D', 'M' ) AND hpo_deptid = 'C1.1' AND hpo_roomid IN ( 50, 54 ) AND hfe_status IN ( 'P', 'R' )) ") \
				_T(" LEFT JOIN hms_fee_invoice ON ( hfe_docno = doc_no AND hfe_invoiceno = invoice_no ) ") \
				_T(" LEFT JOIN hms_roomlist ON (hrl_deptid = 'C1.1' AND hrl_id = room_id)") \
				_T(" WHERE hfe_status = 'P' ") \
				_T("		%s") \
				_T(" GROUP     BY doc_no,room_id, hrl_name ") \
				_T(" ORDER     BY room_id,doc_no "), szWhere);
	
	return szSQL;
}