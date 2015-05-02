#include "stdafx.h"
#include "HMSMainFrame.h"
#include "ReceiptListBloodBank.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CReceiptListBloodBank *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CReceiptListBloodBank* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CReceiptListBloodBank *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CReceiptListBloodBank *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CReceiptListBloodBank *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CReceiptListBloodBank *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CReceiptListBloodBank *pVw = (CReceiptListBloodBank *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CReceiptListBloodBank *pVw = (CReceiptListBloodBank *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddReceiptListBloodBankFnc(CWnd *pWnd){
	 return ((CReceiptListBloodBank*)pWnd)->OnAddReceiptListBloodBank();
} 
static int _OnEditReceiptListBloodBankFnc(CWnd *pWnd){
	 return ((CReceiptListBloodBank*)pWnd)->OnEditReceiptListBloodBank();
} 
static int _OnDeleteReceiptListBloodBankFnc(CWnd *pWnd){
	 return ((CReceiptListBloodBank*)pWnd)->OnDeleteReceiptListBloodBank();
} 
static int _OnSaveReceiptListBloodBankFnc(CWnd *pWnd){
	 return ((CReceiptListBloodBank*)pWnd)->OnSaveReceiptListBloodBank();
} 
static int _OnCancelReceiptListBloodBankFnc(CWnd *pWnd){
	 return ((CReceiptListBloodBank*)pWnd)->OnCancelReceiptListBloodBank();
} 
CReceiptListBloodBank::CReceiptListBloodBank(CWnd *pWnd){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CReceiptListBloodBank::~CReceiptListBloodBank(){
}
void CReceiptListBloodBank::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 91);
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 240, 98, 350, 123);
	m_wndPrintPreview.EnableWindow(FALSE);
	m_wndExport.Create(this, _T("&Export"), 355, 98, 430, 123);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);

}
void CReceiptListBloodBank::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);



}
void CReceiptListBloodBank::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
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
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(FALSE);

}
void CReceiptListBloodBank::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CReceiptListBloodBank::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CReceiptListBloodBank::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CReceiptListBloodBank::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CReceiptListBloodBank::SetMode(int nMode){
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
/*void CReceiptListBloodBank::OnYearChange(){
	
} */
/*void CReceiptListBloodBank::OnYearSetfocus(){
	
} */
/*void CReceiptListBloodBank::OnYearKillfocus(){
	
} */
int CReceiptListBloodBank::OnYearCheckValue(){
	return 0;
} 
void CReceiptListBloodBank::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CReceiptListBloodBank::OnReportPeriodSelendok(){
	CString tmpStr;
	CDate date;
	UpdateData(true);
	int nMonth = ToInt(m_szReportPeriodKey);
	if(nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/01"), m_nYear, nMonth);
		date.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), m_nYear, nMonth, date.GetMonthLastDay());
	}
	if(nMonth == 13){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/03/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59" ), m_nYear,  date.GetMonthLastDay());
	}
	if(nMonth == 14){
		m_szFromDate.Format(_T("%.4d/04/01"), m_nYear);
		tmpStr.Format(_T("%.4d/06/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), m_nYear, date.GetMonthLastDay());
	}
	if(nMonth == 15){
		m_szFromDate.Format(_T("%.4d/07/01"), m_nYear);
		tmpStr.Format(_T("%.4d/09/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), m_nYear, date.GetMonthLastDay());
	}
	if(nMonth == 16){
		m_szFromDate.Format(_T("%.4d/10/01"), m_nYear);
		tmpStr.Format(_T("%.4d/10/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, date.GetMonthLastDay());
	}
	if(nMonth == 17){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/12/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, date.GetMonthLastDay());		
	}	
	UpdateData(false);	 
}
/*void CReceiptListBloodBank::OnReportPeriodSetfocus(){
	
}*/
/*void CReceiptListBloodBank::OnReportPeriodKillfocus(){
	
}*/
long CReceiptListBloodBank::OnReportPeriodLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0){
		szWhere.Format(_T(" WHERE hpr_idx=%ld"), ToInt(m_szReportPeriodKey));
	};
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
/*void CReceiptListBloodBank::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CReceiptListBloodBank::OnFromDateChange(){
	
} */
/*void CReceiptListBloodBank::OnFromDateSetfocus(){
	
} */
/*void CReceiptListBloodBank::OnFromDateKillfocus(){
	
} */
int CReceiptListBloodBank::OnFromDateCheckValue(){
	return 0;
} 
/*void CReceiptListBloodBank::OnToDateChange(){
	
} */
/*void CReceiptListBloodBank::OnToDateSetfocus(){
	
} */
/*void CReceiptListBloodBank::OnToDateKillfocus(){
	
} */
int CReceiptListBloodBank::OnToDateCheckValue(){
	return 0;
} 
void CReceiptListBloodBank::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CReceiptListBloodBank::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);	CString szSQL, tmpStr, szTemp, szWhere;	UpdateData(TRUE);	BeginWaitCursor();	szWhere.Empty();	int nRow =0, nCol = 0;	CExcel xls;	xls.CreateSheet(1);	xls.SetWorksheet(0);	xls.SetColumnWidth(0, 19);	xls.SetColumnWidth(1, 8);	xls.SetColumnWidth(2, 8);	xls.SetColumnWidth(3, 8);	xls.SetColumnWidth(4, 8);	xls.SetColumnWidth(5, 8);	xls.SetColumnWidth(6, 8);	xls.SetColumnWidth(7, 9);	xls.SetColumnWidth(8, 9);	xls.SetColumnWidth(9, 8);	xls.SetColumnWidth(10, 8);	xls.SetColumnWidth(11, 8);	xls.SetColumnWidth(12, 8);	xls.SetColumnWidth(13, 8);	xls.SetColumnWidth(14, 11);	xls.SetColumnWidth(15, 13);	xls.SetCellText(0, 0, _T("\x42\x1ED8 QU\x1ED0\x43 PH\xD2NG"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(0, 1, _T("\x42\x1EC6NH VI\x1EC6N TW QU\xC2N \x110\x1ED8I \x31\x30\x38"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(0, 2, _T("NH\x1EACP M\xC1U V\xC0 \x43H\x1EBE PH\x1EA8M \x43\xC1\x43 LO\x1EA0I "), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));	xls.SetCellText(0, 3, tmpStr, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(1, 5, _T("M\xE1u \x63\xE1\x63 lo\x1EA1i"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(1, 6, _T("M\xE1u to\xE0n ph\x1EA7n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(1, 7, _T("250"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(2, 7, _T("350"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(3, 6, _T("Kh\x1ED1i h\x1ED3ng \x63\x1EA7u"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(3, 7, _T("250"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(4, 7, _T("350"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(5, 6, _T("\x43\x1ED9ng(\x110\x1A1n v\x1ECB)"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(5, 7, _T("250"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(6, 7, _T("350"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(7, 6, _T("\x43\x1ED9ng(ml)"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(8, 5, _T("\x43h\x1EBF ph\x1EA9m"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(8, 6, _T("Huy\x1EBFt t\x1B0\x1A1ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(8, 7, _T("150"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(9, 7, _T("250"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(10, 6, _T("Ti\x1EC3u \x63\x1EA7u"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(10, 7, _T("TC M\xE1y"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(11, 7, _T("TC Pool"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(12, 6, _T("\x43\x1ED9ng(\x110\x1A1n v\x1ECB)"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(12, 7, _T("150"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(13, 7, _T("250"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(14, 6, _T("\x43\x1ED9ng(ml)"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(15, 6, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetMerge(0, 0, 0, 2);	xls.SetMerge(1, 1, 0, 2);	xls.SetMerge(2, 2, 0, 15);	xls.SetMerge(3, 3, 0, 15);	xls.SetMerge(5, 7, 0, 0);	xls.SetMerge(5, 5, 1, 7);	xls.SetMerge(6, 6, 1, 2);	xls.SetMerge(6, 6, 3, 4);	xls.SetMerge(6, 6, 5, 6);	xls.SetMerge(6, 7, 7, 7);	xls.SetMerge(5, 5, 8, 14);	xls.SetMerge(6, 6, 8, 9);	xls.SetMerge(6, 6, 10, 11);	xls.SetMerge(6, 6, 12, 13);	xls.SetMerge(6, 7, 14, 14);	xls.SetMerge(6, 7, 15, 15);	szSQL = GetQueryString();	rs.ExecSQL(szSQL);	nRow = 8;	nCol = 0;	while(!rs.IsEOF()){		rs.GetValue(_T("type"), tmpStr);		xls.SetCellText(nCol+0, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("m250"), tmpStr);		xls.SetCellText(nCol+1, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("m350"), tmpStr);		xls.SetCellText(nCol+2, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("k250"), tmpStr);		xls.SetCellText(nCol+3, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("k350"), tmpStr);		xls.SetCellText(nCol+4, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("congm250"), tmpStr);		xls.SetCellText(nCol+5, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("congm350"), tmpStr);		xls.SetCellText(nCol+6, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("congml"), tmpStr);		xls.SetCellText(nCol+7, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("h150"), tmpStr);		xls.SetCellText(nCol+8, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("h250"), tmpStr);		xls.SetCellText(nCol+9, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("tmay"), tmpStr);		xls.SetCellText(nCol+10, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("tpool"), tmpStr);		xls.SetCellText(nCol+11, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("congc150"), tmpStr);		xls.SetCellText(nCol+12, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("congc250"), tmpStr);		xls.SetCellText(nCol+13, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("congcml"), tmpStr);		xls.SetCellText(nCol+14, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("tongml"), tmpStr);		xls.SetCellText(nCol+15, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		nRow++;		rs.MoveNext();	}	EndWaitCursor();	xls.Save(_T("Exports\\NHAPMAUVACHEPHAMCACLOAI.xls"));
} 
int CReceiptListBloodBank::OnAddReceiptListBloodBank(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CReceiptListBloodBank::OnEditReceiptListBloodBank(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CReceiptListBloodBank::OnDeleteReceiptListBloodBank(){
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
 		OnCancelReceiptListBloodBank();
 	}
	return 0;
}
int CReceiptListBloodBank::OnSaveReceiptListBloodBank(){
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
 		//OnReceiptListBloodBankListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CReceiptListBloodBank::OnCancelReceiptListBloodBank(){
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
int CReceiptListBloodBank::OnReceiptListBloodBankListLoadData(){
	return 0;
}
CString CReceiptListBloodBank::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere;

	szSQL.Format(_T(" SELECT id, type,") \
		_T("        m250,m350,k250,k350,congm250,congm350,congml,") \
		_T("        h150,h250,tmay,tpool,congc150,congc250,congcml,congml+congcml tongml") \
		_T(" FROM(SELECT id, type,") \
		_T("        sum(m250) m250,sum(m350) m350,sum(k250) k250,sum(k350) k350,sum(m250+k250) congm250,sum(m350+k350) congm350,sum(m250+k250)*250+sum(m350+k350)*350 congml,") \
		_T("        sum(h150) h150,sum(h250) h250,sum(tmay) tmay,sum(tpool) tpool,sum(h150+tpool) congc150,sum(h250+tmay) congc250,sum(h150+tpool)*150+sum(h250+tmay)*250 congcml") \
		_T(" FROM(SELECT id,type,") \
		_T("       case when product_id=6966 then 1 else 0 end m250,") \
		_T("       case when product_id=6970 then 1 else 0 end m350,") \
		_T("       case when product_id=6972 then 1 else 0 end k250,") \
		_T("       case when product_id=6973 then 1 else 0 end k350,") \
		_T("       case when product_id=6976 then 1 else 0 end h150,") \
		_T("       case when product_id=6975 then 1 else 0 end h250,") \
		_T("       case when product_id=6978 then 1 else 0 end tpool,") \
		_T("       case when product_id=6977 then 1 else 0 end tmay") \
		_T(" FROM(SELECT product_id,id,cast(type as nvarchar2(254)) type, product_name") \
		_T(" FROM(SELECT product_id,1 id,'L\x1B0u' type,") \
		_T("        product_name") \
		_T(" FROM(SELECT *") \
		_T(" FROM(SELECT sitemid,") \
		_T("             SUM(tondau) AS tondau3,") \
		_T("             SUM(nhap)   AS nhap3,") \
		_T("             SUM(xuat)   AS xuat3") \
		_T(" FROM(SELECT *") \
		_T(" FROM(SELECT impstockid AS stockid,") \
		_T("             impdate          AS iodate,") \
		_T("             sitemid,") \
		_T("             0      AS tondau,") \
		_T("             impqty AS nhap,") \
		_T("             0      AS xuat") \
		_T("      FROM miv") \
		_T("      UNION ALL") \
		_T("      SELECT expstockid, expdate, sitemid, 0, 0, expqty ") \
		_T("      FROM mev)iotbl") \
		_T(" WHERE iodate BETWEEN cast_string2timestamp ('%s' )AND cast_string2timestamp ('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT stockid,") \
		_T("         iodate,") \
		_T("         sitemid,") \
		_T("         COALESCE(impqty-expqty, 0),") \
		_T("         0,") \
		_T("         0") \
		_T(" FROM(SELECT impstockid AS stockid,") \
		_T("           impdate          AS iodate,") \
		_T("           sitemid,") \
		_T("           0 AS periodqty,") \
		_T("           impqty,") \
		_T("           0 AS expqty") \
		_T("      FROM miv") \
		_T("      UNION ALL") \
		_T("      SELECT expstockid, expdate, sitemid, 0, 0, expqty FROM mev) ptbl") \
		_T(" WHERE iodate < cast_string2timestamp ('%s' )") \
		_T(" ) tblx3") \
		_T(" WHERE stockid=20 AND sitemid  > 0") \
		_T(" GROUP BY sitemid") \
		_T(" HAVING SUM(tondau+nhap+xuat) > 0))") \
		_T(" LEFT JOIN m_productitem_view ON (sitemid=product_item_id)") \
		_T(" WHERE LENGTH(product_code) > 0 AND product_org_id = 'BB')") \
		_T(" UNION ALL") \
		_T(" SELECT mbd_product_id product_id,2 id,(select mbo_desc from m_blood_object where mbo_id=mbd_blood_object) as type,") \
		_T("       (select mp_name from m_product where mp_org_id='BB' and mp_product_id=mbd_product_id) product_name") \
		_T(" FROM   purchase_order") \
		_T(" LEFT JOIN m_blood_donation on(mbd_purchase_orderno=po_orderno)") \
		_T(" WHERE po_status<>'O'") \
		_T("  AND (po_approveddate) BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s') and po_storage_id=20") \
		_T(" ))") \
		_T(" GROUP BY type,id") \
		_T(" UNION ALL") \
		_T(" SELECT id,cast(type as nvarchar2(254)) type,sum(m250) m250,sum(m350) m350,sum(k250) k250,sum(k350) k350,sum(m250+k250) congm250,sum(m350+k350) congm350,sum(m250+k250)*250+sum(m350+k350)*350 congml,") \
		_T("        sum(h150) h150,sum(h250) h250,sum(tmay) tmay,sum(tpool) tpool,sum(h150+tpool) congc150,sum(h250+tmay) congc250,sum(h150+tpool)*150+sum(h250+tmay)*250 congcml") \
		_T(" FROM(SELECT id,type,") \
		_T("       case when product_id=6966 then 1 else 0 end m250,") \
		_T("       case when product_id=6970 then 1 else 0 end m350,") \
		_T("       case when product_id=6972 then 1 else 0 end k250,") \
		_T("       case when product_id=6973 then 1 else 0 end k350,") \
		_T("       case when product_id=6976 then 1 else 0 end h150,") \
		_T("       case when product_id=6975 then 1 else 0 end h250,") \
		_T("       case when product_id=6978 then 1 else 0 end tpool,") \
		_T("       case when product_id=6977 then 1 else 0 end tmay") \
		_T(" FROM(SELECT mbd_product_id product_id,3 id,'T\x1ED5ng nh\x1EADp' type,") \
		_T("       (select mp_name from m_product where mp_org_id='BB' and mp_product_id=mbd_product_id) product_name") \
		_T(" FROM   purchase_order") \
		_T(" LEFT JOIN m_blood_donation on(mbd_purchase_orderno=po_orderno)") \
		_T(" WHERE po_status<>'O'") \
		_T("  AND (po_approveddate) BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s') and po_storage_id=20") \
		_T(" UNION ALL") \
		_T(" SELECT product_id,ID,type, product_name") \
		_T(" FROM(SELECT product_id,4 id,'T\x1ED5ng ph\xE1t' type,product_name") \
		_T(" FROM hms_ipharmaorder") \
		_T(" LEFT JOIN hms_ipharmaorderline ON(hpol_orderid=hpo_orderid)") \
		_T(" LEFT JOIN m_productitem_view ON(product_item_id=hpol_product_item_id)") \
		_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = hpol_product_item_id)") \
		_T(" WHERE hpo_org_id='BB' and hpo_orderstatus='A' and hpo_orderdate between cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" select mtl_product_id product_id,4 id,'T\x1ED5ng ph\xE1t' type,(select mp_name from m_product where mp_org_id='BB' and mp_product_id=mtl_product_id) product_name") \
		_T(" from m_transaction ") \
		_T(" left join m_transactionline on(mt_transaction_id=mtl_transaction_id)") \
		_T(" where mt_org_id='BB' and mt_partner_id='XA11' and mt_status='A' and mt_approveddate between cast_string2timestamp('%s') and cast_string2timestamp('%s'))") \
		_T(" UNION ALL ") \
		_T(" select mtl_product_id product_id,5 id,'Ph\xE1t vi\x1EC7n kh\xE1\x63' type,(select mp_name from m_product where mp_org_id='BB' and mp_product_id=mtl_product_id) product_name") \
		_T(" from m_transaction ") \
		_T(" left join m_transactionline on(mt_transaction_id=mtl_transaction_id)") \
		_T(" where mt_org_id='BB' and mt_doctype='EOH' and mt_partner_id not in('XA11') and mt_status='A' and mt_approveddate between cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
		_T(" UNION ALL ") \
		_T(" select mtl_product_id product_id,6 id,'H\x1EE7y' type,(select mp_name from m_product where mp_org_id='BB' and mp_product_id=mtl_product_id) product_name") \
		_T(" from m_transaction ") \
		_T(" left join m_transactionline on(mt_transaction_id=mtl_transaction_id)") \
		_T(" where mt_org_id='BB' and mt_doctype IN('BOO','EOO') and mt_status='A' and mt_approveddate between cast_string2timestamp('%s') and cast_string2timestamp('%s')") \
		_T(" UNION ALL ") \
		_T(" SELECT product_id,7 id,'\x43\xF2n l\x1EA1i' type,") \
		_T("        product_name") \
		_T(" FROM(SELECT *") \
		_T(" FROM(SELECT sitemid,") \
		_T("             SUM(tondau) AS tondau3,") \
		_T("             SUM(nhap)   AS nhap3,") \
		_T("             SUM(xuat)   AS xuat3") \
		_T(" FROM(SELECT *") \
		_T(" FROM(SELECT impstockid AS stockid,") \
		_T("             impdate          AS iodate,") \
		_T("             sitemid,") \
		_T("             0      AS tondau,") \
		_T("             impqty AS nhap,") \
		_T("             0      AS xuat") \
		_T("      FROM miv") \
		_T("      UNION ALL") \
		_T("      SELECT expstockid, expdate, sitemid, 0, 0, expqty ") \
		_T("      FROM mev)iotbl") \
		_T(" WHERE iodate BETWEEN cast_string2timestamp ('%s' )AND cast_string2timestamp ('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT stockid,") \
		_T("         iodate,") \
		_T("         sitemid,") \
		_T("         COALESCE(impqty-expqty, 0),") \
		_T("         0,") \
		_T("         0") \
		_T(" FROM(SELECT impstockid AS stockid,") \
		_T("           impdate          AS iodate,") \
		_T("           sitemid,") \
		_T("           0 AS periodqty,") \
		_T("           impqty,") \
		_T("           0 AS expqty") \
		_T("      FROM miv") \
		_T("      UNION ALL") \
		_T("      SELECT expstockid, expdate, sitemid, 0, 0, expqty FROM mev) ptbl") \
		_T(" WHERE iodate < cast_string2timestamp ('%s' )") \
		_T(" ) tblx3") \
		_T(" WHERE stockid=20 AND sitemid  > 0") \
		_T(" GROUP BY sitemid") \
		_T(" HAVING SUM(tondau+nhap+xuat) > 0))") \
		_T(" LEFT JOIN m_productitem_view ON (sitemid=product_item_id)") \
		_T(" WHERE LENGTH(product_code) > 0 AND product_org_id = 'BB'") \
		_T(" ))") \
		_T(" GROUP BY id,type)") \
		_T(" ORDER BY id"), m_szFromDate,m_szFromDate,
							m_szFromDate,
							m_szFromDate,m_szToDate,
							m_szFromDate,m_szToDate,
							m_szFromDate,m_szToDate,
							m_szFromDate,m_szToDate,
							m_szFromDate,m_szToDate,
							m_szFromDate,m_szToDate,
							m_szToDate,m_szToDate,
							m_szToDate);
	//QueryOpener(szSQL);
	return szSQL;
}