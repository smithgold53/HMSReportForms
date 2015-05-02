#include "stdafx.h"
#include "EMGeneralExaminationReportDialog.h"
#include "ReportDocument.h"
#include "Excel.h"
#include "HMSMainFrame.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((rptGeneralExaminationReportDialog *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((rptGeneralExaminationReportDialog* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((rptGeneralExaminationReportDialog *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((rptGeneralExaminationReportDialog *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((rptGeneralExaminationReportDialog *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((rptGeneralExaminationReportDialog *)pWnd)->OnToDateCheckValue();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	rptGeneralExaminationReportDialog *pVw = (rptGeneralExaminationReportDialog *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	rptGeneralExaminationReportDialog *pVw = (rptGeneralExaminationReportDialog *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	rptGeneralExaminationReportDialog *pVw = (rptGeneralExaminationReportDialog *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddrptGeneralExaminationReportDialogFnc(CWnd *pWnd){
	 return ((rptGeneralExaminationReportDialog*)pWnd)->OnAddrptGeneralExaminationReportDialog();
} 
static int _OnEditrptGeneralExaminationReportDialogFnc(CWnd *pWnd){
	 return ((rptGeneralExaminationReportDialog*)pWnd)->OnEditrptGeneralExaminationReportDialog();
} 
static int _OnDeleterptGeneralExaminationReportDialogFnc(CWnd *pWnd){
	 return ((rptGeneralExaminationReportDialog*)pWnd)->OnDeleterptGeneralExaminationReportDialog();
} 
static int _OnSaverptGeneralExaminationReportDialogFnc(CWnd *pWnd){
	 return ((rptGeneralExaminationReportDialog*)pWnd)->OnSaverptGeneralExaminationReportDialog();
} 
static int _OnCancelrptGeneralExaminationReportDialogFnc(CWnd *pWnd){
	 return ((rptGeneralExaminationReportDialog*)pWnd)->OnCancelrptGeneralExaminationReportDialog();
} 
rptGeneralExaminationReportDialog::rptGeneralExaminationReportDialog(CWnd *pParent)
{
	m_nDlgWidth = 430;
	m_nDlgHeight = 130;
	SetDefaultValues();
	m_bPreview = true;
}
rptGeneralExaminationReportDialog::~rptGeneralExaminationReportDialog()
{
}
void rptGeneralExaminationReportDialog::OnCreateComponents()
{
	m_wndReportConditio.Create(this, _T("Report Condition"), 5, 5, 420, 90);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 205, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 210, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 415, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 205, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 210, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 415, 85); 
	m_wndPrint.Create(this, _T("&Print"), 255, 95, 335, 120);
	m_wndExport.Create(this, _T("&Export"), 340, 95, 420, 120);
}
void rptGeneralExaminationReportDialog::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(16);
//	//m_wndYear.SetCheckValue(true);
//	//m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(35);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	//m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	//m_wndToDate.SetCheckValue(true);
	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);
}
void rptGeneralExaminationReportDialog::OnSetWindowEvents()
{
	//m_wndYear.SetEvent(WE_CHANGE, _OnYearChangeFnc);
	//m_wndYear.SetEvent(WE_SETFOCUS, _OnYearSetfocusFnc);
	//m_wndYear.SetEvent(WE_KILLFOCUS, _OnYearKillfocusFnc);
	m_wndYear.SetEvent(WE_CHECKVALUE, _OnYearCheckValueFnc);
	m_wndReportPeriod.SetEvent(WE_SELENDOK, _OnReportPeriodSelendokFnc);
	//m_wndReportPeriod.SetEvent(WE_SETFOCUS, _OnReportPeriodSetfocusFnc);
	//m_wndReportPeriod.SetEvent(WE_KILLFOCUS, _OnReportPeriodKillfocusFnc);
	m_wndReportPeriod.SetEvent(WE_SELCHANGE, _OnReportPeriodSelectChangeFnc);
	m_wndReportPeriod.SetEvent(WE_LOADDATA, _OnReportPeriodLoadDataFnc);
	//m_wndReportPeriod.SetEvent(WE_ADDNEW, _OnReportPeriodAddNewFnc);
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString tmpStr;
	tmpStr = pMF->GetSysDate();
	m_nYear = ToInt(tmpStr.Left(4));
	//m_szFromDate = m_szToDate = tmpStr;
	m_szReportPeriodKey.Format(_T("%d"), ToInt(tmpStr.Mid(5, 2)));
	SetMode(VM_EDIT);
}
void rptGeneralExaminationReportDialog::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void rptGeneralExaminationReportDialog::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void rptGeneralExaminationReportDialog::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void rptGeneralExaminationReportDialog::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int rptGeneralExaminationReportDialog::SetMode(int nMode)
{
 		int nOldMode = GetMode(); 
 		CGuiView::SetMode(nMode); 
 		CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 		CString szSQL; 
 		CRecord rs(&pMF->m_db);
		CDate dt;
		CDateTime dt1, dt2;
		dt.ParseDate(pMF->GetSysDate());
  		switch(nMode)
		{ 
 		case VM_ADD: 
 			EnableControls(TRUE); 
 			EnableButtons(TRUE, 3, 4, -1); 
 			SetDefaultValues(); 
 			break; 
 		case VM_EDIT: 
 			EnableControls(TRUE); 
 			EnableButtons(TRUE, 0, 1, 2, -1);
			dt1.SetDate(dt.GetYear(), dt.GetMonth(), dt.GetDay());
			dt1.SetTime(0, 1, 0);
			dt2.SetDate(dt.GetYear(), dt.GetMonth(), dt.GetDay());
			dt2.SetTime(23, 59, 0);
			m_szFromDate = dt1.GetDateTime();
			m_szToDate = dt2.GetDateTime();
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
/*void rptGeneralExaminationReportDialog::OnYearChange(){
	
} */
/*void rptGeneralExaminationReportDialog::OnYearSetfocus(){
	
} */
/*void rptGeneralExaminationReportDialog::OnYearKillfocus(){
	
} */
int rptGeneralExaminationReportDialog::OnYearCheckValue(){
	return 0;
} 
void rptGeneralExaminationReportDialog::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void rptGeneralExaminationReportDialog::OnReportPeriodSelendok()
{
	CString tmpStr;
	CDate dte;
	UpdateData(true);
	int nMonth = ToInt(m_szReportPeriodKey);
	if(nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/01"), m_nYear, nMonth);
		dte.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), m_nYear, nMonth, dte.GetMonthLastDay());
	}
	if(nMonth == 13)
	{
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/03/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59" ), m_nYear,  dte.GetMonthLastDay());
	}
	if(nMonth == 14)
	{
		m_szFromDate.Format(_T("%.4d/04/01"), m_nYear);
		tmpStr.Format(_T("%.4d/06/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 15)
	{
		m_szFromDate.Format(_T("%.4d/07/01"), m_nYear);
		tmpStr.Format(_T("%.4d/09/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 16){
		m_szFromDate.Format(_T("%.4d/10/01"), m_nYear);
		tmpStr.Format(_T("%.4d/10/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	if(nMonth == 17){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/12/01"), m_nYear);
		dte.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, dte.GetMonthLastDay());
	}
	UpdateData(false);
}
/*void rptGeneralExaminationReportDialog::OnReportPeriodSetfocus(){
	
}*/
/*void rptGeneralExaminationReportDialog::OnReportPeriodKillfocus(){
	
}*/
long rptGeneralExaminationReportDialog::OnReportPeriodLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0)
	{
		szWhere.Format(_T(" WHERE hpr_idx=%ld"), ToInt(m_szReportPeriodKey));
	}
	m_wndReportPeriod.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT * FROM hms_period_report %s ORDER BY hpr_idx "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndReportPeriod.AddItems(
			rs.GetValue(_T("hpr_idx")), 
			rs.GetValue(_T("hpr_name")),			
			NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void rptGeneralExaminationReportDialog::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void rptGeneralExaminationReportDialog::OnFromDateChange(){
	
} */
/*void rptGeneralExaminationReportDialog::OnFromDateSetfocus(){
	
} */
/*void rptGeneralExaminationReportDialog::OnFromDateKillfocus(){
	
} */
int rptGeneralExaminationReportDialog::OnFromDateCheckValue(){
	return 0;
} 
/*void rptGeneralExaminationReportDialog::OnToDateChange(){
	
} */
/*void rptGeneralExaminationReportDialog::OnToDateSetfocus(){
	
} */
/*void rptGeneralExaminationReportDialog::OnToDateKillfocus(){
	
} */
int rptGeneralExaminationReportDialog::OnToDateCheckValue()
{
	return 0;
} 
void rptGeneralExaminationReportDialog::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CRecord rs(&pMF->m_db);
	CString tmpStr, szSQL, szFromDate, szToDate;
	BeginWaitCursor();
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	CExcel xls;	
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 7);
	xls.SetColumnWidth(1,25);
	xls.SetColumnWidth(2, 14);
	xls.SetColumnWidth(3, 14);
	xls.SetColumnWidth(4, 14);
	xls.SetColumnWidth(5, 14);
	xls.SetColumnWidth(6, 14);
	xls.SetColumnWidth(7, 14);
	xls.SetColumnWidth(8, 14);
	xls.SetColumnWidth(9, 14); 

	xls.SetCellMergedColumns(0, 1, 3);
	xls.SetCellMergedColumns(0, 2, 3);
	xls.SetCellText(0, 1, pMF->m_CompanyInfo.sc_pname,FMT_CENTER,true,12);
	xls.SetCellText(0, 2, pMF->m_CompanyInfo.sc_name,FMT_CENTER,true,12);
	xls.SetCellMergedColumns(0,3,10);
	xls.SetCellText(0, 3, _T("\x42\xC1O \x43\xC1O T\x1ED4NG H\x1EE2P TH\x45O PH\xD2NG KH\xC1M"),FMT_CENTER, true, 16);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x110\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm), 
		          CDateTime::Convert(m_szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	xls.SetCellText(0, 4, tmpStr, FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(0,4,10);

	int nRow = 6;
	xls.SetCellText(0, nRow, _T("STT"), FMT_CENTER,true);	

	TranslateString(_T("Ph\xF2ng kh\xE1m"), tmpStr);
	xls.SetCellText(1, nRow, tmpStr, FMT_CENTER,true);
	
	TranslateString(_T("T/S l\x1EA7n kh\xE1m"), tmpStr);	
	xls.SetCellText(2, nRow, tmpStr, FMT_CENTER,true);	

	TranslateString(_T("V\xE0o vi\x1EC7n"), tmpStr);
	xls.SetCellText(3, nRow, tmpStr, FMT_CENTER,true);	

	TranslateString(_T("\x43huy\x1EC3n vi\x1EC7n"), tmpStr);
	xls.SetCellText(4, nRow, tmpStr, FMT_CENTER,true);	

	TranslateString(_T("\x44\x1ECB\x63h v\x1EE5"), tmpStr);
	xls.SetCellText(5, nRow, tmpStr, FMT_CENTER,true);
	
	TranslateString(_T("\x42\x1EA3o hi\x1EC3m y t\x1EBF"), tmpStr);
	xls.SetCellText(6, nRow, tmpStr, FMT_CENTER,true);

	TranslateString(_T("Tr\x1EBB \x65m"), tmpStr);
	xls.SetCellText(7, nRow, tmpStr, FMT_CENTER,true);	

	TranslateString(_T("\x43huy\x1EC3n kh\xE1m"), tmpStr);
	xls.SetCellText(8, nRow, tmpStr, FMT_CENTER,true);

	TranslateString(_T("Kh\xE1m s\x1EE9\x63 kh\x1ECF\x65"), tmpStr);
	xls.SetCellText(9, nRow, tmpStr, FMT_CENTER,true);
		
	int nIndex = 1;
	int nTotal[10];
	for (int i= 2 ;i<=9;i++)
	{
		nTotal[i]=0;
	}	
		while(!rs.IsEOF())
	{
		nRow++;
		tmpStr.Format(_T("%d"), nIndex++);
		xls.SetCellText(0, nRow, tmpStr, FMT_INTEGER);

		rs.GetValue(_T("roomname"),tmpStr);
		xls.SetCellText(1, nRow, tmpStr, FMT_TEXT);
 
		rs.GetValue(_T("totalexam"), tmpStr);
		nTotal[2] += ToInt(tmpStr);
		xls.SetCellText(2, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("totaladmit"), tmpStr);
		nTotal[3] += ToInt(tmpStr);
		xls.SetCellText(3, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("totaldischarge"), tmpStr);
		nTotal[4] += ToInt(tmpStr);
		xls.SetCellText(4, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("totalservice"), tmpStr);
		nTotal[5] += ToInt(tmpStr);
		xls.SetCellText(5, nRow, tmpStr, FMT_NUMBER1);

		rs.GetValue(_T("totalinsurance"), tmpStr);
		nTotal[6] += ToInt(tmpStr);
		xls.SetCellText(6, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("totalchild"), tmpStr);
		nTotal[7] += ToInt(tmpStr);
		xls.SetCellText(7, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("totalmove"), tmpStr);
		nTotal[8] += ToInt(tmpStr);
		xls.SetCellText(8, nRow, tmpStr, FMT_NUMBER1);
		
		rs.GetValue(_T("totalhealth"), tmpStr);
		nTotal[9] += ToInt(tmpStr);
		xls.SetCellText(9, nRow,tmpStr,FMT_NUMBER1);

		rs.MoveNext();
	}
	nRow++;
	xls.SetCellMergedColumns(0,nRow,2);
	TranslateString(_T("T\x1ED5ng \x63\x1ED9ng"),tmpStr);
	xls.SetCellText(0, nRow, tmpStr, FMT_CENTER, true, 12);
	for (int i =2; i <= 9; i++)
	{		
		tmpStr.Format(_T("%ld"),nTotal[i] );		
		xls.SetCellText(i, nRow, tmpStr, FMT_NUMBER1, true, 12);
	}
	xls.Save(_T("Exports\\Bao cao tong hop theo phong kham benh.xls"));
	EndWaitCursor();

} 
void rptGeneralExaminationReportDialog::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	PrintGeneralExaminationReport(m_szFromDate, m_szToDate);
} 
void rptGeneralExaminationReportDialog::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
} 
int rptGeneralExaminationReportDialog::OnAddrptGeneralExaminationReportDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int rptGeneralExaminationReportDialog::OnEditrptGeneralExaminationReportDialog(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int rptGeneralExaminationReportDialog::OnDeleterptGeneralExaminationReportDialog(){
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
 		OnCancelrptGeneralExaminationReportDialog(); 
 	}
	return 0;
}
int rptGeneralExaminationReportDialog::OnSaverptGeneralExaminationReportDialog(){
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
 		//OnrptGeneralExaminationReportDialogListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int rptGeneralExaminationReportDialog::OnCancelrptGeneralExaminationReportDialog(){
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
int rptGeneralExaminationReportDialog::OnrptGeneralExaminationReportDialogListLoadData(){
	return 0;
}
void rptGeneralExaminationReportDialog::PrintGeneralExaminationReport(CString szFromDate, CString szToDate)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szFromDateLabel, szToDateLabel, szWhere;
	CString szDate, szSysDate;
	szSysDate = pMF->GetSysDate(); 
	CReport rpt; 
	/*
	szSQL.Format(_T(" SELECT hrl_name as roomname, deptid, roomid, ") \
		_T(" 	sum(totalexam) as totalexam,") \
		_T(" 	sum(totaladmit) as totaladmit,") \
		_T(" 	sum(totaldischarge) as totaldischarge,") \
		_T(" 	sum(totalservice) as totalservice,") \
		_T(" 	sum(totalinsurance) as totalinsurance,") \
		_T(" 	sum(totalchild) as totalchild,") \
		_T(" 	sum(totalmove) as totalmove,") \
		_T(" 	sum(totalhealth) as totalhealth") \
		_T(" FROM") \
		_T(" (") \
		_T(" SELECT  he_deptid as deptid, ") \
		_T(" 	he_roomid as roomid, ") \
		_T(" 	case when date(he_examdate) BETWEEN date('%s') and date('%s') and he_status <> 'O' then 1 else 0 end as totalexam,") \
		_T(" 	case when hd_suggestion='A' and date(hd_enddate) BETWEEN date('%s') and date('%s') and hd_status ='T' and hcr_status in('I','T') and he_receptidx = 1 then 1 else 0 end as totaladmit,") \
		_T(" 	case when hd_suggestion='T' and date(hd_enddate) BETWEEN date('%s') and date('%s') and he_receptidx=hd_areceptidx then 1 else 0 end as totaldischarge,") \
		_T(" 	case when ho_type='S' and date(he_examdate) BETWEEN date('%s') and date('%s') and he_status <> 'O' then 1 else 0 end as totalservice,") \
		_T(" 	case when ho_type='I' and date(he_examdate) BETWEEN date('%s') and date('%s') and he_status <> 'O' then 1 else 0 end as totalinsurance,") \
		_T(" 	case when ho_type='C' and date(he_examdate) BETWEEN date('%s') and date('%s') and he_status <> 'O' then 1 else 0 end as totalchild,") \
		_T(" 	case when he_hasfee='M' and date(he_examdate) BETWEEN date('%s') and date('%s') then 1 else 0 end as totalmove, ") \
		_T(" 	case when hd_status ='T' and hfl_subitem='1' then 1 else 0 end as totalhealth") \
		_T(" FROM hms_doc") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_exam ON(he_docno=hd_docno)") \
		_T(" LEFT JOIN hms_fee_list ON(hfl_typeid='E' and hfl_feeid=he_examtype)") \
		_T(" LEFT JOIN hms_object ON(ho_id=hd_object)") \
		_T(" WHERE he_deptid ='%s' ") \
		_T(" ) as tbl") \
		_T(" LEFT JOIN hms_roomlist ON(deptid=hrl_deptid and roomid=hrl_id)") \
		_T(" GROUP BY deptid, roomid, roomname") \
		_T(" ORDER BY deptid, roomid") ,szFromDate, szToDate,szFromDate, szToDate,szFromDate, szToDate,
		szFromDate, szToDate,szFromDate, szToDate,szFromDate, szToDate,szFromDate, szToDate, pMF->m_szDept);
	_fmsg(_T("%s"), szSQL);
	*/
	BeginWaitCursor();
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if(!rpt.Init(_T("Reports/HMS/HE_GENERALEXAMINATIONREPORT.RPT")))
		return ;
	//Report Header
	tmpStr = rpt.GetReportHeader()->GetValue(_T("ReportDate"));
	szDate.Format(tmpStr, CDateTime::Convert(szFromDate, yyyymmdd|hhmm, ddmmyyyy|hhmm),
		          CDateTime::Convert(szToDate, yyyymmdd|hhmm, ddmmyyyy|hhmm));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), szDate);
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	//Report Detail
	CReportSection* rptDetail = rpt.GetDetail(); 
	int nIndex = 1;
	int nTotal[10];
	for (int i =0; i < 10; i++)
		nTotal[i] = 0;
	while(!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();	
		tmpStr.Format(_T("%d"), nIndex++);
		rptDetail->SetValue(_T("0"), tmpStr);
		rs.GetValue(_T("roomname"), tmpStr);
		rptDetail->SetValue(_T("roomname"), tmpStr);
		rs.GetValue(_T("totalexam"), tmpStr);
		nTotal[0] += ToInt(tmpStr);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("totaladmit"), tmpStr);
		nTotal[1] += ToInt(tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		rs.GetValue(_T("totaldischarge"), tmpStr);
		nTotal[2] += ToInt(tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		rs.GetValue(_T("totalservice"), tmpStr);
		nTotal[3] += ToInt(tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		rs.GetValue(_T("totalinsurance"), tmpStr);
		nTotal[4] += ToInt(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		rs.GetValue(_T("totalchild"), tmpStr);
		nTotal[5] += ToInt(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		rs.GetValue(_T("totalmove"), tmpStr);
		nTotal[6] += ToInt(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		rs.GetValue(_T("totalhealth"), tmpStr);
		nTotal[7] += ToInt(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);		
		rs.MoveNext();
	}
	for (int i =0; i < 10; i++){
		tmpStr.Format(_T("sum%d"), i+1);
		rpt.GetReportFooter()->SetValue(tmpStr, ToString(nTotal[i]));		
	}
	//Page Footer
	//Report Footer
	szDate.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), szDate);
	EndWaitCursor();
	rpt.PrintPreview();
}
CString rptGeneralExaminationReportDialog::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL;
	szSQL.Format(_T(" SELECT  hrl_name as roomname, deptid, roomid, ") \
				_T("         sum(totalexam) as totalexam,") \
				_T("         sum(totaladmit) as totaladmit,") \
				_T("         sum(totaldischarge) as totaldischarge,") \
				_T("         sum(totalservice) as totalservice,") \
				_T("         sum(totalinsurance) as totalinsurance,") \
				_T("         sum(totalchild) as totalchild,") \
				_T("         sum(totalmove) as totalmove,") \
				_T("         sum(totalhealth) as totalhealth") \
				_T(" FROM") \
				_T(" (") \
				_T("   SELECT  he_deptid as deptid, ") \
				_T("   he_roomid as roomid, ") \
				_T("   1 as totalexam,") \
				_T("   case when hd_suggestion='D' AND hd_status ='T'  AND he_doctor=hd_doctor AND he_receptidx=hd_areceptidx then 1 else 0 end as totaladmit,") \
				_T("   case when hd_suggestion='F'  AND hd_status ='T'  AND he_doctor=hd_doctor AND he_receptidx=hd_areceptidx then 1 else 0 end as totaldischarge,") \
				_T("   case when ho_type='S' then 1 else 0 end as totalservice,") \
				_T("   case when ho_type='I' then 1 else 0 end as totalinsurance,") \
				_T("   case when ho_type='C' then 1 else 0 end as totalchild,") \
				_T("   case when he_hasfee='M' then 1 else 0 end as totalmove, ") \
				_T("   case when hd_status ='T'  AND he_doctor=hd_doctor AND he_receptidx=hd_areceptidx and hfl_subitem='1' then 1 else 0 end as totalhealth") \
				_T("   FROM hms_doc") \
				_T("   LEFT JOIN hms_exam ON(he_docno=hd_docno)") \
				_T("   LEFT JOIN hms_fee_list ON(hfl_typeid='E' and hfl_feeid=he_examtype)") \
				_T("   LEFT JOIN hms_object ON(ho_id=hd_object)") \
				_T("   WHERE he_deptid ='%s' and he_examdate BETWEEN cast_string2timestamp('%s') and cast_string2timestamp('%s') and he_status <>'O' ") \
				_T("   ) tbl") \
				_T(" LEFT JOIN hms_roomlist ON(deptid=hrl_deptid and roomid=hrl_id)") \
				_T(" GROUP BY deptid, roomid, hrl_name") \
				_T(" ORDER BY deptid, roomid"), pMF->m_szDept, m_szFromDate, m_szToDate);
	return szSQL;
}