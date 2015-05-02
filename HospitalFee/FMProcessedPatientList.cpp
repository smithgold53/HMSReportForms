#include "stdafx.h"
#include "FMProcessedPatientList.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMProcessedPatientList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMProcessedPatientList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMProcessedPatientList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMProcessedPatientList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMProcessedPatientList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMProcessedPatientList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMProcessedPatientList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMProcessedPatientList *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMProcessedPatientList* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMProcessedPatientList *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMProcessedPatientList *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMProcessedPatientList *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMProcessedPatientList *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMProcessedPatientList *)pWnd)->OnClerkAddNew();
}*/
static void _OnServiceSelectFnc(CWnd *pWnd){
	 ((CFMProcessedPatientList*)pWnd)->OnServiceSelect();
}
static void _OnInsuranceSelectFnc(CWnd *pWnd){
	 ((CFMProcessedPatientList*)pWnd)->OnInsuranceSelect();
}
static void _OnPaidSelectFnc(CWnd *pWnd){
	 ((CFMProcessedPatientList*)pWnd)->OnPaidSelect();
}
static void _OnUnpaidSelectFnc(CWnd *pWnd){
	 ((CFMProcessedPatientList*)pWnd)->OnUnpaidSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMProcessedPatientList *pVw = (CFMProcessedPatientList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnPreviewGeneralSheetFnc(CWnd *pWnd){
	CFMProcessedPatientList *pVw = (CFMProcessedPatientList *)pWnd;
	pVw->OnPreviewGeneralSheet();
} 
static long _OnDepartmentLoadDataFnc(CWnd *pWnd){
	return ((CFMProcessedPatientList*)pWnd)->OnDepartmentLoadData();
} 
static void _OnDepartmentDblClickFnc(CWnd *pWnd){
	((CFMProcessedPatientList*)pWnd)->OnDepartmentDblClick();
} 
static void _OnDepartmentSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CFMProcessedPatientList*)pWnd)->OnDepartmentSelectChange(nOldItem, nNewItem);
} 
static int _OnDepartmentDeleteItemFnc(CWnd *pWnd){
	 return ((CFMProcessedPatientList*)pWnd)->OnDepartmentDeleteItem();
} 
CFMProcessedPatientList::CFMProcessedPatientList(CWnd *pParent){

	m_nDlgWidth = 500;
	m_nDlgHeight = 585;
	SetDefaultValues();
}
CFMProcessedPatientList::~CFMProcessedPatientList(){
}
void CFMProcessedPatientList::OnCreateComponents(){
	m_wndProcessedPatientList.Create(this, _T("Processed Patient List"), 5, 5, 490, 90);
	m_wndObject.Create(this, _T("Object"), 5, 95, 245, 150);
	m_wndStatus.Create(this, _T("Status"), 250, 95, 490, 150);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 485, 85); 
	m_wndService.Create(this, _T("Service"), 10, 120, 120, 145);
	m_wndInsurance.Create(this, _T("Insurance"), 125, 120, 235, 145);
	m_wndPaid.Create(this, _T("Paid"), 255, 120, 365, 145);
	m_wndUnpaid.Create(this, _T("Unpaid"), 370, 120, 480, 145);
	m_wndPreviewGeneralSheet.Create(this, _T("Preview General Sheet"), 225, 550, 375, 575);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 380, 550, 490, 575);
	m_wndDepartment.Create(this,5, 155, 490, 545); 
	m_wndDepartment.SetCheckBox(true);

}
void CFMProcessedPatientList::OnInitializeComponents(){
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

	m_wndDepartment.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndDepartment.InsertColumn(1, _T("Name"), CFMT_TEXT, 300);
}

void CFMProcessedPatientList::OnSetWindowEvents(){
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
	m_wndService.SetEvent(WE_CLICK, _OnServiceSelectFnc);
	m_wndInsurance.SetEvent(WE_CLICK, _OnInsuranceSelectFnc);
	m_wndPaid.SetEvent(WE_CLICK, _OnPaidSelectFnc);
	m_wndUnpaid.SetEvent(WE_CLICK, _OnUnpaidSelectFnc);
	m_wndPreviewGeneralSheet.SetEvent(WE_CLICK, _OnPreviewGeneralSheetFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndDepartment.SetEvent(WE_SELCHANGE, _OnDepartmentSelectChangeFnc);
	m_wndDepartment.SetEvent(WE_LOADDATA, _OnDepartmentLoadDataFnc);
	m_wndDepartment.SetEvent(WE_DBLCLICK, _OnDepartmentDblClickFnc);
	m_wndDepartment.AddEvent(1, _T("Delete"), _OnDepartmentDeleteItemFnc, 0, VK_DELETE, 0);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	OnDepartmentLoadData();
	UpdateData(false);

}
void CFMProcessedPatientList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndService.GetDlgCtrlID(), m_bService);
	DDX_Check(pDX, m_wndInsurance.GetDlgCtrlID(), m_bInsurance);
	DDX_Check(pDX, m_wndPaid.GetDlgCtrlID(), m_bPaid);
	DDX_Check(pDX, m_wndUnpaid.GetDlgCtrlID(), m_bUnpaid);

}
void CFMProcessedPatientList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bService=FALSE;
	m_bInsurance=FALSE;
	m_bPaid=FALSE;
	m_bUnpaid=FALSE;
}

int CFMProcessedPatientList::SetMode(int nMode){
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

/*void CFMProcessedPatientList::OnFromDateChange(){
	
} */
/*void CFMProcessedPatientList::OnFromDateSetfocus(){
	
} */
/*void CFMProcessedPatientList::OnFromDateKillfocus(){
	
} */
int CFMProcessedPatientList::OnFromDateCheckValue(){
	return 0;
}
 
/*void CFMProcessedPatientList::OnToDateChange(){
	
} */
/*void CFMProcessedPatientList::OnToDateSetfocus(){
	
} */
/*void CFMProcessedPatientList::OnToDateKillfocus(){
	
} */
int CFMProcessedPatientList::OnToDateCheckValue(){
	return 0;
}
 
void CFMProcessedPatientList::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
 
void CFMProcessedPatientList::OnClerkSelendok(){
	 
}

/*void CFMProcessedPatientList::OnClerkSetfocus(){
	
}*/
/*void CFMProcessedPatientList::OnClerkKillfocus(){
	
}*/
long CFMProcessedPatientList::OnClerkLoadData(){
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

/*void CFMProcessedPatientList::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMProcessedPatientList::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CReportSection *rptDetail;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, tmpStr2, szWhere, szNewDept, szOldDept, szDate, szAmount, szSysDate;
	CString szNewTit, szOldTit;
	double nDrugAmt = 0, nMatAmt = 0;
	if (!rpt.Init(_T("Reports\\HMS\\PTTT\\PMS_DANHSACHBENHNHANDUYETTHUOC_CT.RPT")))
		return;

	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	
	long double grpCost[8];
	long double ttlCost[8];
	double cost = 0;
	int nItem = 1;

	for (int i = 0; i < 8; i++)
	{
		grpCost[i] = 0;
		ttlCost[i] = 0;
	}
			
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("tit"), szNewTit);
		if (szNewTit != szOldTit)
		{
			if (grpCost[7] > 0)
			{	
				TranslateString(_T("Total Group"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));	
				rptDetail->SetValue(_T("TotalGroup"), szAmount);
				for (int i = 5; i < 8; i++)
				{
					FormatCurrency(grpCost[i], tmpStr);
					tmpStr2.Format(_T("s%d"), i + 1);
					rptDetail->SetValue(tmpStr2, tmpStr);
					ttlCost[i] += grpCost[i];					
					grpCost[i] = 0;
				}			
			}
			if (ttlCost[7] > 0)
			{				
				TranslateString(_T("Total Amount"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				rptDetail->SetValue(_T("TotalAmount"), szAmount);
				for (int i = 5; i < 8; i++)
				{
					FormatCurrency(ttlCost[i], tmpStr);
					tmpStr2.Format(_T("s%d"), i + 1);
					rptDetail->SetValue(tmpStr2, tmpStr);
					ttlCost[i] = 0;
				}
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			TranslateString(szNewTit, tmpStr);
			rptDetail->SetValue(_T("GroupName"), tmpStr);
			szOldTit = szNewTit;			
		}
		rs.GetValue(_T("dept"), szNewDept);
		if (szNewDept != szOldDept)
		{
			if (grpCost[7] > 0)
			{	
				TranslateString(_T("Total Group"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));	
				rptDetail->SetValue(_T("TotalGroup"), szAmount);
				for (int i = 5; i < 8; i++)
				{
					FormatCurrency(grpCost[i], tmpStr);
					tmpStr2.Format(_T("s%d"), i + 1);
					rptDetail->SetValue(tmpStr2, tmpStr);
					ttlCost[i] += grpCost[i];					
					grpCost[i] = 0;
				}			
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rptDetail->SetValue(_T("GroupName"), pMF->GetDepartmentName(szNewDept));
			szOldDept = szNewDept;
		}
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nItem++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("id"), tmpStr);				
		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("recordno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("age"), tmpStr);				
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("materialamt"), cost);	
		grpCost[5] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("drugamt"), cost);	
		grpCost[6] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("amount"), cost);	
		grpCost[7] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.MoveNext();
	}
			
	if (grpCost[7] > 0)
	{	
		TranslateString(_T("Total Group"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));	
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		for (int i = 5; i < 8; i++)
		{
			FormatCurrency(grpCost[i], tmpStr);
			tmpStr2.Format(_T("s%d"), i + 1);
			rptDetail->SetValue(tmpStr2, tmpStr);
			ttlCost[i] += grpCost[i];					
			grpCost[i] = 0;
		}			
	}
	if (ttlCost[7] > 0)
	{				
		TranslateString(_T("Total Amount"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
		rptDetail->SetValue(_T("TotalAmount"), szAmount);
		for (int i = 5; i < 8; i++)
		{
			FormatCurrency(ttlCost[i], tmpStr);
			tmpStr2.Format(_T("s%d"), i + 1);
			rptDetail->SetValue(tmpStr2, tmpStr);
		}
	}
	CString szTime;
	szTime = pMF->GetSysDateTime();
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
			        szTime.Right(8), szSysDate.Right(2), szSysDate.Mid(5,2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();	
}

void CFMProcessedPatientList::OnPreviewGeneralSheet(){
	CGuiMenu menu(this);
	CString tmpStr;
	CRect rect;
	menu.CreatePopupMenu();
	m_wndPreviewGeneralSheet.GetWindowRect(&rect);
	TranslateString(_T("Export General Sheet"), tmpStr);
	menu.AppendMenu(MF_BYPOSITION, 1, tmpStr);
	TranslateString(_T("Preview General Sheet"), tmpStr);
	menu.AppendMenu(MF_SEPARATOR);
	menu.AppendMenu(MF_BYPOSITION, 2, tmpStr);
	int nPos = menu.TrackPopupMenu(TPM_RETURNCMD|TPM_RIGHTALIGN|TPM_BOTTOMALIGN|TPM_NONOTIFY, rect.right, rect.top, this);
	switch (nPos)
	{
		case 1:
			OnExportGeneralSheet();
			break;
		case 2:
			OnPrintGeneralSheet();
			break;
	}
}
void CFMProcessedPatientList::OnPrintGeneralSheet(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CReportSection *rptDetail;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, tmpStr2, szWhere, szNewDept, szOldDept, szDate, szAmount, szSysDate;
	CString szNewTit, szOldTit;
	double nDrugAmt = 0, nMatAmt = 0;
	if (!rpt.Init(_T("Reports\\HMS\\PTTT\\PMS_DANHSACHBENHNHANDUYETTHUOC_TH.RPT")))
		return;

	szSQL = GetQueryString1();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	
	long double grpCost[7];
	long double ttlCost[7];
	double cost = 0;
	int nItem = 1, nIdx = 1;

	for (int i = 0; i < 7; i++)
	{
		grpCost[i] = 0;
		ttlCost[i] = 0;
	}
			
	while(!rs.IsEOF())
	{
		rs.GetValue(_T("tit"), szNewTit);
		if (szNewTit != szOldTit)
		{
			if (grpCost[5] + grpCost[6]> 0)
			{	
				TranslateString(_T("Total Group"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));	
				rptDetail->SetValue(_T("TotalGroup"), szAmount);
				for (int i = 0; i < 7; i++)
				{
					FormatCurrency(grpCost[i], tmpStr);
					tmpStr2.Format(_T("s%d"), i + 6);
					rptDetail->SetValue(tmpStr2, tmpStr);
					ttlCost[i] += grpCost[i];					
					grpCost[i] = 0;
				}			
			}
			if (ttlCost[5] + ttlCost[6] > 0)	
			{				
				TranslateString(_T("Total Amount"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
				rptDetail->SetValue(_T("TotalAmount"), szAmount);
				for (int i = 0; i < 7; i++)
				{
					FormatCurrency(ttlCost[i], tmpStr);
					tmpStr2.Format(_T("s%d"), i + 6);
					rptDetail->SetValue(tmpStr2, tmpStr);
					ttlCost[i] = 0;
				}
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			TranslateString(szNewTit, tmpStr);
			rptDetail->SetValue(_T("GroupName"), tmpStr);
			szOldTit = szNewTit;		
			nIdx = 1;
		}
		rs.GetValue(_T("dept"), szNewDept);
		if (szNewDept != szOldDept)
		{
			if (grpCost[5] +grpCost[6] > 0)
			{	
				TranslateString(_T("Total Group"), szAmount);
				rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));	
				rptDetail->SetValue(_T("TotalGroup"), szAmount);
				for (int i = 0; i < 7; i++)
				{
					FormatCurrency(grpCost[i], tmpStr);
					tmpStr2.Format(_T("s%d"), i + 6);
					rptDetail->SetValue(tmpStr2, tmpStr);
					ttlCost[i] += grpCost[i];					
					grpCost[i] = 0;
				}			
			}
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			tmpStr.Format(_T("%d. %s"), nIdx++, pMF->GetDepartmentName(szNewDept));
			rptDetail->SetValue(_T("GroupName"), tmpStr);
			szOldDept = szNewDept;
			nItem = 1;
		}
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nItem++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("id"), tmpStr);				
		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("recordno"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("age"), tmpStr);				
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("treatmaterialamt"), cost);	
		grpCost[0] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("treatdrugamt"), cost);	
		grpCost[1] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("surgerymaterialamt"), cost);	
		grpCost[2] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("surgerydrugamt"), cost);	
		grpCost[3] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("hitechmaterialamt"), cost);
		grpCost[4] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.GetValue(_T("paidamount"), cost);	
		grpCost[5] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.GetValue(_T("debtamount"), cost);	
		grpCost[6] += cost;
		FormatCurrency(cost, tmpStr);
		rptDetail->SetValue(_T("12"), tmpStr);
		rs.MoveNext();
	}
			
	if (grpCost[5] + grpCost[6]> 0)
	{	
		TranslateString(_T("Total Group"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));	
		rptDetail->SetValue(_T("TotalGroup"), szAmount);
		for (int i = 0; i < 7; i++)
		{
			FormatCurrency(grpCost[i], tmpStr);
			tmpStr2.Format(_T("s%d"), i + 6);
			rptDetail->SetValue(tmpStr2, tmpStr);
			ttlCost[i] += grpCost[i];					
			grpCost[i] = 0;
		}			
	}
	if (ttlCost[5] + ttlCost[6]> 0)
	{				
		TranslateString(_T("Total Amount"), szAmount);
		rptDetail = rpt.AddDetail(rpt.GetGroupFooter(0));
		rptDetail->SetValue(_T("TotalAmount"), szAmount);
		for (int i = 0; i < 7; i++)
		{
			FormatCurrency(ttlCost[i], tmpStr);
			tmpStr2.Format(_T("s%d"), i + 6);
			rptDetail->SetValue(tmpStr2, tmpStr);
		}
	}
	CString szTime;
	szTime = pMF->GetSysDateTime();
	szSysDate = pMF->GetSysDate(); 
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),
			        szTime.Right(8), szSysDate.Right(2), szSysDate.Mid(5,2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	rpt.PrintPreview();
	
}

void CFMProcessedPatientList::OnExportGeneralSheet(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CExcel xls;
	CString szSQL, tmpStr, szOldDept, szNewDept, szOldTit, szNewTit;
	CStringArray arrCol;
	double nTmp = 0;
	long double nGrpCost[7], nTtCost[7];
	int nIdx = 1, nRow = 0;
	szSQL = GetQueryString1();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."));
		return;
	}
	for (int i = 0; i < 7; i++)
	{
		nGrpCost[i] = 0;
		nTtCost[i] = 0;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(1, 30);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 15);
	xls.SetColumnWidth(7, 15);
	xls.SetColumnWidth(8, 15);
	xls.SetColumnWidth(9, 15);
	xls.SetColumnWidth(10, 15);
	xls.SetColumnWidth(11, 15);

	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 11);
	xls.SetCellMergedColumns(0, 3, 11);
	xls.SetCellText(0, 0, pMF->m_szHealthService, 4098, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, 4098, true);
	//TODO: Write Excel Name
	tmpStr = _T("\x42\x1EA2NG QUY\x1EBET TO\xC1N THU\x1ED0\x43 - V\x1EACT T\x1AF T\x1ED4NG H\x1EE2P");
	xls.SetCellText(0, 2, tmpStr, 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, 4098, true);
	//TODO: Write Column Header
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("H\x1ECD v\xE0 t\xEAn"));
	arrCol.Add(_T("S\x1ED1 \x42\x41"));
	arrCol.Add(_T("Tu\x1ED5i"));
	arrCol.Add(_T("S\x1ED1 h\x1ED3 s\x1A1"));
	arrCol.Add(_T("Ti\x1EC1n VT \x110T"));
	arrCol.Add(_T("Ti\x1EC1n thu\x1ED1\x63 \x110T"));
	arrCol.Add(_T("Ti\x1EC1n VT m\x1ED5"));
	arrCol.Add(_T("Ti\x1EC1n thu\x1ED1\x63 \x110T"));
	arrCol.Add(_T("Ti\x1EC1n VTKT\x43"));
	arrCol.Add(_T("T\x1ED5ng th\x61nh to\xE1n"));
	arrCol.Add(_T("T\x1ED5ng n\x1EE3"));
	for (int i = 0; i < arrCol.GetCount(); i++)
		xls.SetCellText(i, 4, arrCol.GetAt(i), 4098, true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("tit"), szNewTit);
		if (szNewTit != szOldTit)
		{
			if ((nGrpCost[5] + nGrpCost[6]) > 0)
			{
				xls.SetCellText(4, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
				for (int i = 0; i < 7; i++)
				{
					tmpStr.Format(_T("%f"), nGrpCost[i]);
					xls.SetCellText(i+5, nRow, tmpStr, FMT_NUMBER1, true);
					nTtCost[i] += nGrpCost[i];
					nGrpCost[i] = 0;
				}
				nRow++;
			}
			if ((nTtCost[5] + nTtCost[6]) > 0)
			{
				xls.SetCellText(4, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
				for (int i = 0; i < 7; i++)
				{
					tmpStr.Format(_T("%f"), nTtCost[i]);
					xls.SetCellText(i+5, nRow, tmpStr, FMT_NUMBER1, true);
					nTtCost[i] = 0;
				}
				nRow++;
			}
			xls.SetCellMergedColumns(0, nRow, 12);
			TranslateString(szNewTit, tmpStr);
			xls.SetCellText(0, nRow, tmpStr, FMT_TEXT, true);
			szOldDept.Empty();
			szOldTit = szNewTit;
			nIdx = 1;
			nRow++;
		}
		rs.GetValue(_T("dept"), szNewDept);
		if (szNewDept != szOldDept)
		{
			if ((nGrpCost[5] + nGrpCost[6]) > 0)
			{
				xls.SetCellText(4, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
				for (int i = 0; i < 7; i++)
				{
					tmpStr.Format(_T("%f"), nGrpCost[i]);
					xls.SetCellText(i+5, nRow, tmpStr, FMT_NUMBER1, true);
					nTtCost[i] += nGrpCost[i];
					nGrpCost[i] = 0;
				}
				nRow++;
			}
			xls.SetCellMergedColumns(0, nRow, 12);
			xls.SetCellText(0, nRow, pMF->GetDepartmentName(szNewDept), FMT_TEXT, true);
			szOldDept = szNewDept;
			nIdx = 1;
			nRow++;
		}
		xls.SetCellText(0, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(1, nRow, rs.GetValue(_T("pname")), FMT_TEXT);
		xls.SetCellText(2, nRow, rs.GetValue(_T("recordno")), FMT_TEXT);
		xls.SetCellText(3, nRow, rs.GetValue(_T("age")), FMT_TEXT);
		xls.SetCellText(4, nRow, rs.GetValue(_T("docno")), FMT_TEXT);
		rs.GetValue(_T("treatmaterialamt"), nTmp);
		nGrpCost[0] += nTmp;
		xls.SetCellText(5, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("treatdrugamt"), nTmp);
		nGrpCost[1] += nTmp;
		xls.SetCellText(6, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("surgerymaterialamt"), nTmp);
		nGrpCost[2] += nTmp;
		xls.SetCellText(7, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("surgerydrugamt"), nTmp);
		nGrpCost[3] += nTmp;
		xls.SetCellText(8, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("hitechmaterialamt"), nTmp);
		nGrpCost[4] += nTmp;
		xls.SetCellText(9, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("paidamount"), nTmp);
		nGrpCost[5] += nTmp;
		xls.SetCellText(10, nRow, double2str(nTmp), FMT_NUMBER1);
		rs.GetValue(_T("debtamount"), nTmp);
		nGrpCost[6] += nTmp;
		xls.SetCellText(11, nRow, double2str(nTmp), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	if ((nGrpCost[5] + nGrpCost[6]) > 0)
	{
		xls.SetCellText(4, nRow, _T("T\x1ED5ng nh\xF3m"), 4098, true);
		for (int i = 0; i < 7; i++)
		{
			tmpStr.Format(_T("%f"), nGrpCost[i]);
			xls.SetCellText(i+5, nRow, tmpStr, FMT_NUMBER1, true);
			nTtCost[i] += nGrpCost[i];
			nGrpCost[i] = 0;
		}
		nRow++;
	}
	if ((nTtCost[5] + nTtCost[6]) > 0)
	{
		xls.SetCellText(4, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), 4098, true);
		for (int i = 0; i < 7; i++)
		{
			tmpStr.Format(_T("%f"), nTtCost[i]);
			xls.SetCellText(i+5, nRow, tmpStr, FMT_NUMBER1, true);
			nTtCost[i] = 0;
		}
		nRow++;
	}
	xls.Save(_T("Exports\\Bang quyet toan thuoc vat tu.xls"));
}

void CFMProcessedPatientList::OnDepartmentDblClick(){
	
} 
void CFMProcessedPatientList::OnDepartmentSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMProcessedPatientList::OnDepartmentDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CFMProcessedPatientList::OnDepartmentLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndDepartment.BeginLoad(); 
	m_wndDepartment.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" select sd_id as id, sd_name as name from sys_dept where sd_type = 'DT' order by  sd_id"));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDepartment.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndDepartment.EndLoad(); 
	return nCount;
}

void CFMProcessedPatientList::OnServiceSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	if (m_bService && m_bInsurance)
		m_wndInsurance.SetCheck(false);
}
void CFMProcessedPatientList::OnInsuranceSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	if (m_bInsurance && m_bService)
		m_wndService.SetCheck(false);
}
void CFMProcessedPatientList::OnPaidSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	if (m_bPaid && m_bUnpaid)
		m_wndUnpaid.SetCheck(false);
}
void CFMProcessedPatientList::OnUnpaidSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	if (m_bUnpaid && m_bPaid)
		m_wndPaid.SetCheck(false);
}

CString CFMProcessedPatientList::GetQueryString(){
	CString szSQL, szWhere, szWhere1, tmpStr;
	szWhere.Format(_T(" AND hpo_processeddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szWhere1.Format(_T(" AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"),m_szFromDate, m_szToDate);
	if (!m_szClerkKey.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hpo_processedby = '%s'"), m_szClerkKey);
		szWhere1.AppendFormat(_T(" AND mt_approvedby= '%s'"), m_szClerkKey);
	}
	for (int i = 0; i < m_wndDepartment.GetItemCount(); i++)
	{
		if (m_wndDepartment.GetCheck(i))
		{
			if (!tmpStr.IsEmpty())
				tmpStr += _T(", ");
			tmpStr.AppendFormat(_T("%s"), m_wndDepartment.GetItemText(i, 0));
		}
	}
	if (!tmpStr.IsEmpty())
	{
		szWhere.AppendFormat(_T(" AND hpo_deptid IN ('%s')"), tmpStr);
		szWhere1.AppendFormat(_T(" AND hpo_deptid IN ('%s')"), tmpStr);
	}
	szSQL.Format(_T("SELECT tit, pname, ") \
				_T("       age, ") \
				_T("       docno, ") \
				_T("	   dept,") \
				_T("	   hcr_recordno recordno,") \
				_T("       Sum(drugamt)     drugamt, ") \
				_T("       Sum(materialamt) materialamt, ") \
				_T("       Sum(amount)      amount ") \
				_T("FROM   (SELECT 'Drug Daily'	tit,") \
				_T("			   Get_patientname(hpol_docno)    pname, ") \
				_T("               Hms_getagebydoc(hpol_docno)    age, ") \
				_T("               hpol_docno                     docno, ") \
				_T("			   hpo_deptid					  dept,") \
				_T("               CASE ") \
				_T("                 WHEN product_org_id = 'PM' THEN hpol_qtyorder * hpol_unitprice ") \
				_T("                 ELSE 0 ") \
				_T("               END                            AS drugamt, ") \
				_T("               CASE ") \
				_T("                 WHEN product_org_id = 'MA' THEN hpol_qtyorder * hpol_unitprice ") \
				_T("                 ELSE 0 ") \
				_T("               END                            AS materialamt, ") \
				_T("               hpol_qtyorder * hpol_unitprice AS amount ") \
				_T("        FROM   hms_ipharmaorder ") \
				_T("               LEFT JOIN hms_ipharmaorderline ") \
				_T("                      ON ( hpo_orderid = hpol_orderid ) ") \
				_T("               LEFT JOIN m_product_view ") \
				_T("                      ON ( hpol_product_id = product_id ) ") \
				_T("        WHERE  hpo_processed = 'Y' AND coalesce(hpo_reference_id, 0) = 0 AND hpo_deptid <> 'B5' ") \
				_T("		%s ") \
				_T("		UNION ALL ") \
				_T("		SELECT 'Return'	tit,") \
				_T("			   Get_patientname(hpol_docno)    pname, ") \
				_T("               Hms_getagebydoc(hpol_docno)    age, ") \
				_T("               hpol_docno                     docno, ") \
				_T("			   hpo_deptid					  dept,") \
				_T("               CASE ") \
				_T("                 WHEN product_org_id = 'PM' THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                 ELSE 0 ") \
				_T("               END                            AS drugamt, ") \
				_T("               CASE ") \
				_T("                 WHEN product_org_id = 'MA' THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("                 ELSE 0 ") \
				_T("               END                            AS materialamt, ") \
				_T("               hpol_qtyreverse * hpol_unitprice AS amount ") \
				_T("        FROM   hms_ipharmaorder ") \
				_T("               LEFT JOIN hms_ipharmaorderline ") \
				_T("                      ON ( hpo_orderid = hpol_orderid ) ") \
				_T("               LEFT JOIN m_product_view ") \
				_T("                      ON ( hpol_product_id = product_id ) ") \
				_T("			   LEFT JOIN m_transaction ON (mt_transaction_id = hpol_retorder_id)") \
				_T("        WHERE  hpol_qtyreverse > 0 AND mt_status = 'A' AND mt_storage_to_id IN (14, 15) ") \
				_T("		%s ) tbl ") \
				_T(" LEFT JOIN hms_clinical_record ON (hcr_docno = docno)") \
				_T(" GROUP  BY tit, pname, ") \
				_T("          age, ") \
				_T("          docno, hcr_recordno, dept ") \
				_T(" ORDER BY tit, dept, docno"), szWhere, szWhere1);
	return szSQL;
}

CString CFMProcessedPatientList::GetQueryString1(){
	CString szSQL, szPharmaWhere, szPharmaWhere1, szTransactionWhere, tmpStr, szWhere;
	szPharmaWhere.Format(_T(" AND hpo_processeddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szPharmaWhere1.Format(_T(" AND hpo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szTransactionWhere.Format(_T(" AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"),m_szFromDate, m_szToDate);
	if (!m_szClerkKey.IsEmpty())
	{
		szPharmaWhere.AppendFormat(_T(" AND hpo_processedby = '%s'"), m_szClerkKey);
		szPharmaWhere1.AppendFormat(_T(" AND hpo_approveby = '%s'"), m_szClerkKey);
		szTransactionWhere.AppendFormat(_T(" AND mt_approvedby= '%s'"), m_szClerkKey);
	}
	for (int i = 0; i < m_wndDepartment.GetItemCount(); i++)
	{
		if (m_wndDepartment.GetCheck(i))
		{
			if (!tmpStr.IsEmpty())
				tmpStr += _T(", ");
			tmpStr.AppendFormat(_T("%s"), m_wndDepartment.GetItemText(i, 0));
		}
	}
	if (!tmpStr.IsEmpty())
	{
		szPharmaWhere.AppendFormat(_T(" AND hpo_deptid IN ('%s')"), tmpStr);
		szPharmaWhere1.AppendFormat(_T(" AND hpo_deptid IN ('%s')"), tmpStr);
		szTransactionWhere.AppendFormat(_T(" AND hpo_deptid IN ('%s')"), tmpStr);
	}
	szWhere.Format(_T(" WHERE 1=1"));
	if (m_bService)
		szWhere.AppendFormat(_T(" AND hpo_object_id = 7"));
	else if (m_bInsurance)
		szWhere.AppendFormat(_T(" AND hpo_object_id NOT IN (1, 3, 7, 8, 6)"));
	if (m_bPaid)
		szWhere.AppendFormat(_T(" AND HCRF_ACCEPTEDFEE = 'P'"));
	else if (m_bUnpaid)
		szWhere.AppendFormat(_T(" AND HCRF_ACCEPTEDFEE <> 'P'"));
	szSQL.Format(_T("SELECT tit, docno, ") \
				_T("		Get_patientname(docno) pname,") \
				_T("		Hms_getagebydoc(docno) age,") \
				_T("	   dept,") \
				_T("	   hcr_recordno recordno ,") \
				_T("       Sum(treatdrugamt)     treatdrugamt, ") \
				_T("       Sum(surgerydrugamt)     surgerydrugamt, ") \
				_T("       Sum(treatmaterialamt) treatmaterialamt, ") \
				_T("       Sum(surgerymaterialamt) surgerymaterialamt, ") \
				_T("	   Sum(hitechdrugamt) hitechdrugamt,") \
				_T("	   Sum(hitechmaterialamt) hitechmaterialamt,") \
				_T("       case when HCRF_ACCEPTEDFEE = 'P' ") \
				_T("	   AND HCRF_ACCEPTEDDATE between cast_string2timestamp('%s') and cast_string2timestamp('%s') ") \
				_T("	   then Sum(amount) else 0 end paidamount, ") \
				_T("       case when HCRF_ACCEPTEDFEE <> 'P' ") \
				_T("	   OR HCRF_ACCEPTEDDATE not between cast_string2timestamp('%s') and cast_string2timestamp('%s') ") \
				_T("	   then Sum(amount) else 0 end debtamount ") \
				_T("FROM   (SELECT 'Drug Daily'	tit,") \
				_T("           hpol_docno   docno, ") \
				_T("		   hpo_object_id,") \
				_T("           CASE WHEN hpo_deptid = 'B5' THEN (SELECT ho_deptid ") \
				_T("                                        FROM   hms_operation ") \
				_T("                                        WHERE  ho_idx = hpo_reference_id) ") \
				_T("           ELSE hpo_deptid ") \
				_T("           END          AS dept, ") \
				_T("           CASE WHEN product_org_id = 'PM' THEN hpol_qtyorder * hpol_unitprice ") \
				_T("           ELSE 0 ") \
				_T("           END          AS treatdrugamt, ") \
				_T("           CASE WHEN product_org_id = 'MA' THEN hpol_qtyorder * hpol_unitprice ") \
				_T("           ELSE 0 ") \
				_T("           END          AS treatmaterialamt, ") \
				_T("           0            surgerydrugamt, ") \
				_T("           0            surgerymaterialamt, ") \
				_T("		   0			hitechdrugamt,") \
				_T("		   0			hitechmaterialamt,") \
				_T("		   hpol_qtyorder*hpol_unitprice amount") \
				_T(" FROM      hms_ipharmaorder ") \
				_T(" LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T(" LEFT JOIN m_product_view ON ( hpol_product_id = product_id ) ") \
				_T(" WHERE      hpo_ordertype IN ('D', 'M', 'B') AND hpo_storage_id IN (14, 15)") \
				_T("       AND Coalesce(hpo_reference_id, 0) = 0 ") \
				_T("       AND hpo_processed = 'Y' ") \
				_T("		%s") \
				_T(" UNION ALL ") \
				_T(" SELECT    'Drug Daily' tit, ") \
				_T("           hpol_docno   docno, ") \
				_T("		   hpo_object_id,") \
				_T("           CASE WHEN hpo_deptid = 'B5' THEN (SELECT ho_deptid ") \
				_T("                                        FROM   hms_operation ") \
				_T("                                        WHERE  ho_idx = hpo_reference_id) ") \
				_T("           ELSE hpo_deptid ") \
				_T("           END          AS dept, ") \
				_T("           0            treatdrugamt, ") \
				_T("           0            treatmaterialamt, ") \
				_T("           CASE WHEN product_org_id = 'PM' THEN hpol_qtyorder * hpol_unitprice ") \
				_T("           ELSE 0 ") \
				_T("           END          surgerydrugamt, ") \
				_T("           CASE WHEN product_org_id = 'MA' THEN hpol_qtyorder * hpol_unitprice ") \
				_T("           ELSE 0 ") \
				_T("           END          surgerymaterialamt, ") \
				_T("		   0			hitechdrugamt,") \
				_T("		   0			hitechmaterialamt,") \
				_T("		   hpol_qtyorder*hpol_unitprice amount") \
				_T(" FROM      hms_ipharmaorder ") \
				_T(" LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T(" LEFT JOIN m_product_view ON ( hpol_product_id = product_id ) ") \
				_T(" LEFT JOIN m_storagelist ON (msl_storage_id = hpo_storage_id)") \
				_T(" WHERE hpo_ordertype IN ('D', 'M', 'C') AND msl_category = 'S'") \
				_T("       AND hpo_reference_id > 0 ") \
				_T("       AND hpo_orderstatus = 'A' ") \
				_T("		%s") \
				_T(" UNION ALL ") \
				_T(" SELECT    'Drug Daily' tit, ") \
				_T("           hpol_docno   docno, ") \
				_T("		   hpo_object_id,") \
				_T("           CASE WHEN hpo_deptid = 'B5' THEN (SELECT ho_deptid ") \
				_T("                                        FROM   hms_operation ") \
				_T("                                        WHERE  ho_idx = hpo_reference_id) ") \
				_T("           ELSE hpo_deptid ") \
				_T("           END          AS dept, ") \
				_T("           0            treatdrugamt, ") \
				_T("           0            treatmaterialamt, ") \
				_T("           0          surgerydrugamt, ") \
				_T("           0          surgerymaterialamt, ") \
				_T("           CASE WHEN product_org_id = 'PM' THEN hpol_qtyorder * hpol_unitprice ") \
				_T("           ELSE 0 ") \
				_T("           END			hitechdrugamt,") \
				_T("		   CASE WHEN product_org_id = 'MA' THEN hpol_qtyorder * hpol_unitprice ") \
				_T("           ELSE 0 ") \
				_T("           END			hitechmaterialamt,") \
				_T("		   hpol_qtyorder*hpol_unitprice amount") \
				_T(" FROM      hms_ipharmaorder ") \
				_T(" LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T(" LEFT JOIN m_product_view ON ( hpol_product_id = product_id ) ") \
				_T(" WHERE hpo_ordertype IN ('D', 'M') AND hpo_storage_id IN (12, 13) AND hpo_object_id = 7") \
				_T(" AND hpo_reference_id > 0 ") \
				_T(" AND hpo_orderstatus = 'A' ") \
				_T("	%s") \
				_T(" UNION ALL ") \
				_T(" SELECT 'Return'	tit,") \
				_T("         hpol_docno                     docno, ") \
				_T("		 hpo_object_id,") \
				_T("		 case when hpo_deptid = 'B5' then (select ho_deptid from hms_operation where ho_idx = hpo_reference_id) else hpo_deptid end as dept,") \
				_T("         CASE ") \
				_T("         WHEN product_org_id = 'PM' AND coalesce(hpo_reference_id, 0) = 0 THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("         ELSE 0 ") \
				_T("         END AS treatdrugamt, ") \
				_T("         CASE ") \
				_T("         WHEN product_org_id = 'MA' AND coalesce(hpo_reference_id, 0) = 0 THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("         ELSE 0 ") \
				_T("         END AS treatmaterialamt, ") \
				_T("         CASE ") \
				_T("         WHEN product_org_id = 'PM' AND hpo_reference_id > 0 THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("         ELSE 0 ") \
				_T("         END AS surgerydrugamt, ") \
				_T("         CASE ") \
				_T("         WHEN product_org_id = 'MA' AND hpo_reference_id > 0 THEN hpol_qtyreverse * hpol_unitprice ") \
				_T("         ELSE 0 ") \
				_T("         END AS surgerymaterialamt, ") \
				_T("		 0,") \
				_T("		 0,") \
				_T("		 hpol_qtyreverse*hpol_unitprice amount") \
				_T("        FROM   hms_ipharmaorder ") \
				_T("        LEFT JOIN hms_ipharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T("        LEFT JOIN m_product_view ON ( hpol_product_id = product_id ) ") \
				_T("		LEFT JOIN m_transaction ON (mt_transaction_id = hpol_retorder_id)") \
				_T("        WHERE  hpol_qtyreverse > 0 AND mt_status = 'A' AND mt_storage_to_id IN (14, 15) ") \
				_T("		%s) tbl ") \
				_T(" LEFT JOIN hms_clinical_record ON (hcr_docno = docno) ") \
				_T("	%s") \
				_T(" GROUP  BY tit, ") \
				_T("          dept,") \
				_T("          docno, hcr_recordno, HCRF_ACCEPTEDFEE, HCRF_ACCEPTEDDATE") \
				_T(" ORDER BY tit, dept, docno"), m_szFromDate, m_szToDate, m_szFromDate, m_szToDate
				, szPharmaWhere, szPharmaWhere1, szPharmaWhere1, szTransactionWhere, szWhere);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}