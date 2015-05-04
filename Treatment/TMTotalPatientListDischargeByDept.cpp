#include "stdafx.h"
#include "TMTotalPatientListDischargeByDept.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CTMTotalPatientListDischargeByDept *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CTMTotalPatientListDischargeByDept *)pWnd)->OnToDateCheckValue();
} 
static void _OnLoadSelectFnc(CWnd *pWnd){
	CTMTotalPatientListDischargeByDept *pVw = (CTMTotalPatientListDischargeByDept *)pWnd;
	pVw->OnLoadSelect();
} 
static void _OnAllSelectFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept*)pWnd)->OnAllSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept*)pWnd)->OnInPatientSelect();
}
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept*)pWnd)->OnOutPatientSelect();
}
static void _OnApprovalFeeSelectFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept*)pWnd)->OnApprovalFeeSelect();
}
static void _OnUnapprovedFeeSelectFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept*)pWnd)->OnUnapprovedFeeSelect();
}
static void _OnPaymentSelectFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept*)pWnd)->OnPaymentSelect();
}
static void _OnPrintSelectFnc(CWnd *pWnd){
	CTMTotalPatientListDischargeByDept *pVw = (CTMTotalPatientListDischargeByDept *)pWnd;
	pVw->OnPrintSelect();
} 
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CTMTotalPatientListDischargeByDept*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CTMTotalPatientListDischargeByDept*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CTMTotalPatientListDischargeByDept*)pWnd)->OnListDeleteItem();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMTotalPatientListDischargeByDept* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CTMTotalPatientListDischargeByDept *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnObjectAddNew();
}*/
static void _OnRankSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMTotalPatientListDischargeByDept* )pWnd)->OnRankSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnRankSelendokFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnRankSelendok();
}
/*static void _OnRankSetfocusFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnRankKillfocus();
}*/
/*static void _OnRankKillfocusFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnRankKillfocus();
}*/
static long _OnRankLoadDataFnc(CWnd *pWnd){
	return ((CTMTotalPatientListDischargeByDept *)pWnd)->OnRankLoadData();
}
/*static void _OnRankAddNewFnc(CWnd *pWnd){
	((CTMTotalPatientListDischargeByDept *)pWnd)->OnRankAddNew();
}*/
static int _OnAddTMPatientListDischargeFnc(CWnd *pWnd){
	 return ((CTMTotalPatientListDischargeByDept*)pWnd)->OnAddTMPatientListDischarge();
} 
static int _OnEditTMPatientListDischargeFnc(CWnd *pWnd){
	 return ((CTMTotalPatientListDischargeByDept*)pWnd)->OnEditTMPatientListDischarge();
} 
static int _OnDeleteTMPatientListDischargeFnc(CWnd *pWnd){
	 return ((CTMTotalPatientListDischargeByDept*)pWnd)->OnDeleteTMPatientListDischarge();
} 
static int _OnSaveTMPatientListDischargeFnc(CWnd *pWnd){
	 return ((CTMTotalPatientListDischargeByDept*)pWnd)->OnSaveTMPatientListDischarge();
} 
static int _OnCancelTMPatientListDischargeFnc(CWnd *pWnd){
	 return ((CTMTotalPatientListDischargeByDept*)pWnd)->OnCancelTMPatientListDischarge();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CTMTotalPatientListDischargeByDept *pVw = (CTMTotalPatientListDischargeByDept *)pWnd;
	pVw->OnExportSelect();
} 
CTMTotalPatientListDischargeByDept::CTMTotalPatientListDischargeByDept(CWnd *pWnd){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CTMTotalPatientListDischargeByDept::~CTMTotalPatientListDischargeByDept(){
}
void CTMTotalPatientListDischargeByDept::OnCreateComponents(){
	m_wndReportFilterInfo.Create(this, _T("Report Condition"), 5, 5, 398, 144);
	m_wndFromDateLabel.Create(this, _T("From Date"), 9, 29, 81, 54);
	m_wndFromDate.Create(this,84, 29, 194, 54); 
	m_wndToDateLabel.Create(this, _T("To Date"), 197, 29, 277, 54);
	m_wndToDate.Create(this,282, 29, 392, 54); 
	m_wndRanklab.Create(this, _T("Rank"), 9, 58, 81, 83);
	m_wndRank.Create(this,84, 58, 194, 83); 
	m_wndObjectLabel.Create(this, _T("Object"), 197, 58, 277, 83);
	m_wndObject.Create(this,282, 58, 392, 83); 
	m_wndPrint.Create(this, _T("&Print"), 233, 152, 313, 177);
	m_wndExport.Create(this, _T("Export"), 317, 152, 397, 177);
	m_wndAll.Create(this, _T("All"), 11, 88, 91, 113);
	m_wndInPatient.Create(this, _T("InPatient"), 96, 88, 176, 113);
	m_wndOutPatient.Create(this, _T("OutPatient"), 181, 88, 261, 113);
	m_wndApprovalFee.Create(this, _T("Approval Fee"), 12, 116, 113, 141);
	m_wndUnapprovedFee.Create(this, _T("Unapproved Fee"), 104, 116, 217, 141);
	m_wndPayment.Create(this, _T("Payment"), 218, 116, 298, 141);

	m_wndList.InsertColumn(0, _T("STT"), CFMT_NUMBER, 50);
	m_wndList.InsertColumn(1, _T("S\x1ED1 h\x1ED3 s\x1A1"), CFMT_TEXT, 100);
	m_wndList.InsertColumn(2, _T("T\xEAn \x42\x1EC7nh Nh\xE2n"), CFMT_TEXT, 150);
	m_wndList.InsertColumn(3, _T("N\x103m sinh"), CFMT_DATE, 80);
	m_wndList.InsertColumn(4, _T("\x43\x1EA5p \x62\x1EAD\x63"), CFMT_TEXT, 100);
	m_wndList.InsertColumn(5, _T("\x110\x1A1n v\x1ECB"), CFMT_TEXT, 150);
	m_wndList.InsertColumn(6, _T("\x43h\x1EA5n \x111o\xE1n"), CFMT_TEXT, 200);
	m_wndList.InsertColumn(7, _T("\x110\x1ED1i t\x1B0\x1EE3ng"), CFMT_TEXT, 80);
	m_wndList.InsertColumn(8, _T("V\xE0o kho\x61"), CFMT_TEXT, 80);
	m_wndList.InsertColumn(9, _T("Ng\xE0y r\x61"), CFMT_DATE, 100);


	m_wndObject.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

	m_wndRank.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndRank.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CTMTotalPatientListDischargeByDept::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	//m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	//m_wndToDate.SetCheckValue(true);
	//m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(1024);
	//m_wndRank.SetCheckValue(true);
	m_wndRank.LimitText(1024);

}
void CTMTotalPatientListDischargeByDept::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndLoad.SetEvent(WE_CLICK, _OnLoadSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);
	m_wndList.AddEvent(1, _T("Delete"), _OnListDeleteItemFnc, 0, VK_DELETE, 0);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndRank.SetEvent(WE_SELENDOK, _OnRankSelendokFnc);
	//m_wndRank.SetEvent(WE_SETFOCUS, _OnRankSetfocusFnc);
	//m_wndRank.SetEvent(WE_KILLFOCUS, _OnRankKillfocusFnc);
	m_wndRank.SetEvent(WE_SELCHANGE, _OnRankSelectChangeFnc);
	m_wndRank.SetEvent(WE_LOADDATA, _OnRankLoadDataFnc);
	//m_wndRank.SetEvent(WE_ADDNEW, _OnRankAddNewFnc);
	m_wndAll.SetEvent(WE_CLICK, _OnAllSelectFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndApprovalFee.SetEvent(WE_CLICK, _OnApprovalFeeSelectFnc);
	m_wndUnapprovedFee.SetEvent(WE_CLICK, _OnUnapprovedFeeSelectFnc);
	m_wndPayment.SetEvent(WE_CLICK, _OnPaymentSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate +=_T("23:59");
	UpdateData(false);

}
void CTMTotalPatientListDischargeByDept::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_TextEx(pDX, m_wndRank.GetDlgCtrlID(), m_szRankKey);
	DDX_Radio(pDX, m_wndAll.GetDlgCtrlID(), m_nAll);
	DDX_Radio(pDX, m_wndInPatient.GetDlgCtrlID(), m_nInPatient);
	DDX_Radio(pDX, m_wndOutPatient.GetDlgCtrlID(), m_nOutPatient);
	DDX_Check(pDX, m_wndApprovalFee.GetDlgCtrlID(), m_bApprovalFee);
	DDX_Check(pDX, m_wndUnapprovedFee.GetDlgCtrlID(), m_bUnapprovedFee);
	DDX_Check(pDX, m_wndPayment.GetDlgCtrlID(), m_bPayment);

}
void CTMTotalPatientListDischargeByDept::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CTMTotalPatientListDischargeByDept::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CTMTotalPatientListDischargeByDept::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szObjectKey.Empty();
	m_szRankKey.Empty();
	m_nAll=0;
	m_nInPatient=1;
	m_nOutPatient=1;
	m_bApprovalFee=FALSE;
	m_bUnapprovedFee=FALSE;
	m_bPayment=FALSE;

}
int CTMTotalPatientListDischargeByDept::SetMode(int nMode){
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
/*void CTMTotalPatientListDischargeByDept::OnFromDateChange(){
	
} */
/*void CTMTotalPatientListDischargeByDept::OnFromDateSetfocus(){
	
} */
/*void CTMTotalPatientListDischargeByDept::OnFromDateKillfocus(){
	
} */
int CTMTotalPatientListDischargeByDept::OnFromDateCheckValue(){
	return 0;
} 
/*void CTMTotalPatientListDischargeByDept::OnToDateChange(){
	
} */
/*void CTMTotalPatientListDischargeByDept::OnToDateSetfocus(){
	
} */
/*void CTMTotalPatientListDischargeByDept::OnToDateKillfocus(){
	
} */
int CTMTotalPatientListDischargeByDept::OnToDateCheckValue(){
	return 0;
} 
void CTMTotalPatientListDischargeByDept::OnLoadSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	OnListLoadData();
} 
void CTMTotalPatientListDischargeByDept::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere, tmpStr, szNewGroup, szOldGroup;
	if (!rpt.Init(_T("Reports\\HMS\\BAOCAOTONGKETRAVIEN_KHOA.RPT")))
		return;
	szSQL = GetQueryString();
	int ret =0;
	ret = rs.ExecSQL(szSQL);
	_fmsg(_T("%s"),szSQL);
	if (ret <=0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->GetHospitalName());
	rpt.GetReportHeader()->SetValue(_T("DEPARTMENT"), pMF->GetDepartmentName(pMF->GetDepartmentID()));
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection *rptDetail = NULL, * rptGroup = NULL;
	int nIdx=1;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("loai"), szNewGroup);
		if (szNewGroup != szOldGroup && !szNewGroup.IsEmpty())
		{
			rptGroup = rpt.AddDetail(rpt.GetGroupHeader(1));
			if (rptGroup)
			{
					tmpStr = szNewGroup;
				rptGroup->SetValue(_T("GroupName"), tmpStr);
			}

			szOldGroup = szNewGroup;
			nIdx = 1;
		}
		rptDetail = rpt.AddDetail();
		if (rptDetail)
		{
		tmpStr.Format(_T("%d"), nIdx++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("hovaten"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("doituong"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("buong"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("benhan"), tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);
		rs.GetValue(_T("giayravien"), tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);
		rs.GetValue(_T("phoithanhtoan"), tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);
		rs.GetValue(_T("thebhyt"), tmpStr);
		rptDetail->SetValue(_T("13"), tmpStr);
		rs.GetValue(_T("cmt"), tmpStr);
		rptDetail->SetValue(_T("14"), tmpStr);
		rs.GetValue(_T("giaytokhac"), tmpStr);
		rptDetail->SetValue(_T("15"), tmpStr);
		}
		rs.MoveNext();
	}
	CString szSysDate;
	szSysDate = pMF->GetSysDate(); 
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
	
} 
void CTMTotalPatientListDischargeByDept::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();	CRecord rs(&pMF->m_db);	CString szSQL, tmpStr, szTemp, szWhere;	UpdateData(TRUE);	BeginWaitCursor();	szWhere.Empty();	int nRow =0, nCol = 0;	CExcel xls;	xls.CreateSheet(1);	xls.SetWorksheet(0);	xls.SetColumnWidth(0, 6);	xls.SetColumnWidth(1, 12);	xls.SetColumnWidth(2, 8);	xls.SetColumnWidth(3, 22);	xls.SetColumnWidth(4, 18);	xls.SetColumnWidth(5, 22);	xls.SetColumnWidth(6, 6);	xls.SetColumnWidth(7, 6);	xls.SetColumnWidth(8, 6);	xls.SetColumnWidth(9, 6);	xls.SetColumnWidth(10, 6);	xls.SetColumnWidth(11, 6);	xls.SetColumnWidth(12, 8);	xls.SetColumnWidth(13, 8);	xls.SetColumnWidth(14, 8);	xls.SetColumnWidth(15, 8);	xls.SetColumnWidth(16, 8);	xls.SetCellText(nCol, nRow, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 11);
	xls.SetCellText(nCol+1, nRow + 1, pMF->GetDepartmentName(pMF->GetDepartmentID()), FMT_TEXT | FMT_CENTER, true, 11);	xls.SetCellText(0, 1, _T("KHO\x41:"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(0, 3, _T("\x42\xC1O \x43\xC1O T\x1ED4NG K\x1EBET R\x41 VI\x1EC6N"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 12);	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));	xls.SetCellText(0, 4, tmpStr, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(0, 6, _T("STT"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(1, 6, _T("S\x1ED1 \x62\x1EC7nh \xE1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(2, 6, _T("S\x1ED1 h\x1ED3 s\x1A1"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(3, 6, _T("H\x1ECD t\xEAn \x62\x1EC7nh nh\xE2n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(4, 6, _T("Ng\xE0y r\x61"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(5, 6, _T("\x110\x1ED1i t\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(6, 6, _T("N\x1ED9i \x64ung n\x1ED9p"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(6, 7, _T("\x42\x1EC7nh \xE1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(7, 7, _T("Gi\x1EA5y r\x61 vi\x1EC7n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(8, 7, _T("Ph\x1A1i th\x61nh to\xE1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(9, 7, _T("Th\x1EBB \x42HYT"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(10, 7, _T("\x43MT"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(11, 7, _T("Gi\x1EA5y t\x1EDD kh\xE1\x63"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(12, 6, _T("K\xFD (Ghi r\xF5 h\x1ECD t\xEAn)"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(12, 7, _T("Ng\x1B0\x1EDDi n\x1ED9p"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(13, 7, _T("KHTH"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(14, 7, _T("T\xE0i \x63h\xEDnh"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(15, 6, _T("Lo\x1EA1i gi\x1EA5y t\x1EDD n\x1EE3"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(16, 6, _T("Ghi \x63h\xFA"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetMerge(0, 0, 0, 3);	xls.SetMerge(3, 3, 0, 16);	xls.SetMerge(4, 4, 0, 16);	xls.SetMerge(6, 7, 0, 0);	xls.SetMerge(1, 1, 1, 3);	xls.SetMerge(6, 7, 1, 1);	xls.SetMerge(6, 7, 2, 2);	xls.SetMerge(6, 7, 3, 3);	xls.SetMerge(6, 7, 4, 4);	xls.SetMerge(6, 7, 5, 5);	xls.SetMerge(6, 6, 6, 11);	xls.SetMerge(6, 6, 12, 14);	xls.SetMerge(6, 7, 15, 15);	xls.SetMerge(6, 7, 16, 16);	szSQL = GetQueryString();	rs.ExecSQL(szSQL);	nRow = 8;	nCol = 0;	int nIndex=1;	while(!rs.IsEOF()){		tmpStr.Format(_T("%d"),nIndex++);		xls.SetCellText(nCol+0, nRow, tmpStr, FMT_NUMBER1);		rs.GetValue(_T("soba"), tmpStr);		xls.SetCellText(nCol+1, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("hoso"), tmpStr);		xls.SetCellText(nCol+2, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("hovaten"), tmpStr);		xls.SetCellText(nCol+3, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("ngayra"), tmpStr);		xls.SetCellText(nCol+4, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("doituong"), tmpStr);		xls.SetCellText(nCol+5, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("benhan"), tmpStr);		xls.SetCellText(nCol+6, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("giayravien"), tmpStr);		xls.SetCellText(nCol+7, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("phoithanhtoan"), tmpStr);		xls.SetCellText(nCol+8, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("thebhyt"), tmpStr);		xls.SetCellText(nCol+9, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("cmt"), tmpStr);		xls.SetCellText(nCol+10, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("giaytokhac"), tmpStr);		xls.SetCellText(nCol+11, nRow, tmpStr, FMT_TEXT);		nRow++;		rs.MoveNext();	}
		//xls.SetCellMergedColumns(nCol+14, nRow, 3);
		//tmpStr.Format(_T("Ng\xE0y %s th\xE1ng %s n\x103m %s"));
		//xls.SetCellText(nCol+10, nRow, tmpStr, FMT_TEXT|FMT_CENTER, true);
		xls.SetCellMergedColumns(nCol+14, nRow+1, 3);
		xls.SetCellText(nCol+14, nRow+1, _T("Ng\x1B0\x1EDDi l\xE0m \x62\xE1o \x63\xE1o"), FMT_TEXT|FMT_CENTER, true);
		xls.SetCellMergedColumns(nCol+14, nRow+2, 3);
		xls.SetCellText(nCol+14, nRow+2, _T("(K\xFD v\xE0 ghi r\xF5 h\x1ECD t\xEAn)"), FMT_TEXT|FMT_CENTER);	EndWaitCursor();	xls.Save(_T("Exports\\Baocaotongketravien.xls"));

} 
void CTMTotalPatientListDischargeByDept::OnListDblClick(){
	
} 
void CTMTotalPatientListDischargeByDept::OnListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CTMTotalPatientListDischargeByDept::OnListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CTMTotalPatientListDischargeByDept::OnListLoadData(){
	UpdateData(TRUE);
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString tmpStr, szSQL, szWhere, szNewGroup, szOldGroup;
	m_wndList.BeginLoad(); 
	int nCount = 0;
	szSQL = GetQueryString();
	nCount = rs.ExecSQL(szSQL);
	_fmsg(_T("%s"),szSQL);
	int nIndex=1;
	int nItem=1;
	while(!rs.IsEOF()){
		tmpStr.Format(_T("%d"),nIndex++);
		m_wndList.AddItems(
			tmpStr,
			rs.GetValue(_T("hd_docno")), 
			rs.GetValue(_T("hovaten")), 
			rs.GetValue(_T("age")), 
			rs.GetValue(_T("CAPBAC")), 
			rs.GetValue(_T("HP_WORKPLACE")), 
			rs.GetValue(_T("HD_DIAGNOSTIC")), 
			rs.GetValue(_T("DOITUONG")), 
			rs.GetValue(_T("htr_deptid")),
			rs.GetValue(_T("htr_dischargedate")), NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
}
void CTMTotalPatientListDischargeByDept::OnAllSelect(){
	m_szCondition = _T("A");
	//OnListLoadData();
}
void CTMTotalPatientListDischargeByDept::OnInPatientSelect(){
	m_szCondition = _T("I");
	//OnListLoadData();
}
void CTMTotalPatientListDischargeByDept::OnOutPatientSelect(){
	m_szCondition = _T("E");
	//OnListLoadData();
}
void CTMTotalPatientListDischargeByDept::OnApprovalFeeSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	m_bUnapprovedFee = false ;
	m_bPayment = false;
	UpdateData(false);
}
void CTMTotalPatientListDischargeByDept::OnUnapprovedFeeSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	m_bApprovalFee = false;
	m_bPayment = false;
	UpdateData(false);
}
void CTMTotalPatientListDischargeByDept::OnPaymentSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	m_bApprovalFee = false;
	m_bUnapprovedFee = false ;
	UpdateData(false);
}
void CTMTotalPatientListDischargeByDept::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMTotalPatientListDischargeByDept::OnObjectSelendok(){
	 
}
/*void CTMTotalPatientListDischargeByDept::OnObjectSetfocus(){
	
}*/
/*void CTMTotalPatientListDischargeByDept::OnObjectKillfocus(){
	
}*/
long CTMTotalPatientListDischargeByDept::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty())
	{
		//szWhere.Format();
	};
	szSQL.Format(_T("SELECT ho_id as id, ho_desc as description FROM hms_object ORDER BY cast(id as integer)"));
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CTMTotalPatientListDischargeByDept::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CTMTotalPatientListDischargeByDept::OnRankSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMTotalPatientListDischargeByDept::OnRankSelendok(){
	 
}
/*void CTMTotalPatientListDischargeByDept::OnRankSetfocus(){
	
}*/
/*void CTMTotalPatientListDischargeByDept::OnRankKillfocus(){
	
}*/
long CTMTotalPatientListDischargeByDept::OnRankLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	m_wndRank.DeleteAllItems(); 
	szSQL.Format(_T("SELECT ss_code id,ss_desc name FROM sys_sel WHERE ss_id='hms_rank' and ss_code >12"), szWhere);
	_fmsg(_T("%s"),szSQL);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndRank.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CTMTotalPatientListDischargeByDept::OnRankAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
int CTMTotalPatientListDischargeByDept::OnAddTMPatientListDischarge(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CTMTotalPatientListDischargeByDept::OnEditTMPatientListDischarge(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CTMTotalPatientListDischargeByDept::OnDeleteTMPatientListDischarge(){
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
 		OnCancelTMPatientListDischarge();
 	}
	return 0;
}
int CTMTotalPatientListDischargeByDept::OnSaveTMPatientListDischarge(){
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
 		//OnTMPatientListDischargeListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CTMTotalPatientListDischargeByDept::OnCancelTMPatientListDischarge(){
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
int CTMTotalPatientListDischargeByDept::OnTMPatientListDischargeListLoadData(){
	return 0;
}
CString CTMTotalPatientListDischargeByDept::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString szSQL, szWhere, szDept, szObject, szRank, szDept1, szFilter;
	if (!m_szObjectKey.IsEmpty())
		szObject.AppendFormat(_T(" AND hd_object = '%s'"), m_szObjectKey);

	if (!m_szRankKey.IsEmpty())
		szRank.AppendFormat(_T(" AND hp_rank = '%s'"), m_szRankKey);

	if (m_szCondition == _T("A"))
		szWhere.AppendFormat(_T(" AND htr_outpatient IN('N','Y') "));
	//Trang thai noi tru
	if (m_szCondition == _T("I"))
		szWhere.AppendFormat(_T(" AND htr_outpatient='N' "));
	//Trang thai ngoai tru
	if (m_szCondition == _T("E"))
		szWhere.AppendFormat(_T(" AND htr_outpatient='Y' "));
	
	if (m_bApprovalFee)
		szFilter = _T(" AND htrf_acceptedfee='Y'");
	else if (m_bUnapprovedFee)
		szFilter = _T(" AND htrf_acceptedfee='N'");
	else if (m_bPayment)
		szFilter = _T(" AND htrf_acceptedfee='P'");

	szWhere.AppendFormat(_T(" AND htr_status='T' AND htr_suggestion <> 'M' AND htr_dischargedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szDept.AppendFormat(_T(" AND htr_deptid='%s'"), pMF->GetCurrentDepartmentID());
	szSQL.Format(_T(" SELECT DISTINCT ")\
		_T("	 CASE WHEN ho_type IN('S') THEN '\x110\x1ED1i t\x1B0\x1EE3ng \x64\x1ECB\x63h v\x1EE5' WHEN ho_type IN('C', 'I') THEN '\x110\x1ED1i t\x1B0\x1EE3ng \x62\x1EA3o hi\x1EC3m' WHEN ho_type IN('D', 'P') THEN '\x110\x1ED1i t\x1B0\x1EE3ng \x63h\xEDnh s\xE1\x63h' END loai,") \
		_T("	 htr_recordno soba,") \
		_T("     hd_docno hoso,") \
		_T("	 hms_getroomname(hb_deptid, last_value(hb_roomid) OVER (PARTITION BY hb_docno)) buong,") \
		_T("     get_patientname(hd_docno) hovaten,") \
		_T("     htr_dischargedate ngayra,") \
		_T("     get_objectname(hd_object) doituong,") \
		_T("     CASE WHEN hhc_record = 'Y' THEN 'X' END benhan,") \
		_T("     CASE WHEN hhc_dischargedsheet = 'Y' THEN 'X' END giayravien,") \
		_T("     CASE WHEN hhc_paymentsheet = 'Y' THEN 'X' END phoithanhtoan,") \
		_T("     CASE WHEN hhc_inscard = 'Y' THEN 'X' END thebhyt,") \
		_T("     CASE WHEN hhc_idcard = 'Y' THEN 'X' END cmt,") \
		_T("     CASE WHEN hhc_others = 'Y' THEN hhc_desc END giaytokhac") \
		_T("   FROM hms_patient") \
		_T("   LEFT JOIN hms_doc") \
		_T("   ON(hd_patientno=hp_patientno)") \
		_T("   LEFT JOIN hms_object") \
		_T("   ON(ho_id = hd_object)") \
		_T("   LEFT JOIN hms_exam") \
		_T("   ON(hd_docno=he_docno)") \
		_T("   LEFT JOIN hms_clinical_record") \
		_T("   ON(hcr_docno=hd_docno)") \
		_T("   LEFT JOIN hms_treatment_record") \
		_T("   ON(htr_docno =hd_docno") \
		_T("   AND htr_idx  =hcr_refidx)") \
		_T("   LEFT JOIN hms_bed") \
		_T("   ON( htr_docno =hb_docno") \
		_T("   AND htr_idx = hb_refidx)") \
		_T("   LEFT JOIN hms_handin_content ON(hhc_docno =htr_docno)") \
		_T("   WHERE 1             =1 %s %s %s %s %s") \
		_T("   ORDER BY loai, ngayra"), szDept, szObject, szRank, szWhere,szFilter);
	return szSQL;
}