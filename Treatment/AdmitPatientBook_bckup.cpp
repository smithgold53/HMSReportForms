#include "stdafx.h"
#include "AdmitPatientBook.h"
#include "HMSMainFrame.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CAdmitPatientBook *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CAdmitPatientBook *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CAdmitPatientBook *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CAdmitPatientBook *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CAdmitPatientBook *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CAdmitPatientBook *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CAdmitPatientBook *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CAdmitPatientBook *)pWnd)->OnToDateCheckValue();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CAdmitPatientBook *pVw = (CAdmitPatientBook *)pWnd;
	pVw->OnExportSelect();
} 
CAdmitPatientBook::CAdmitPatientBook(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CAdmitPatientBook::~CAdmitPatientBook(){
}
void CAdmitPatientBook::OnCreateComponents(){
	m_wndAdmitPatientBook.Create(this, _T("Admit Patient Book"), 5, 5, 470, 60);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 235, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 240, 30, 320, 55);
	m_wndToDate.Create(this,325, 30, 465, 55); 
	m_wndExport.Create(this, _T("&Export"), 390, 65, 470, 90);

}
void CAdmitPatientBook::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);

}
void CAdmitPatientBook::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CAdmitPatientBook::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CAdmitPatientBook::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CAdmitPatientBook::SetMode(int nMode){
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
/*void CAdmitPatientBook::OnFromDateChange(){
	
} */
/*void CAdmitPatientBook::OnFromDateSetfocus(){
	
} */
/*void CAdmitPatientBook::OnFromDateKillfocus(){
	
} */
int CAdmitPatientBook::OnFromDateCheckValue(){
	return 0;
} 
/*void CAdmitPatientBook::OnToDateChange(){
	
} */
/*void CAdmitPatientBook::OnToDateSetfocus(){
	
} */
/*void CAdmitPatientBook::OnToDateKillfocus(){
	
} */
int CAdmitPatientBook::OnToDateCheckValue(){
	return 0;
} 
void CAdmitPatientBook::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	CStringArray arrField;
	int nCol = 0, nRow = 0, nIdx = 1;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(3, 30);
	xls.SetColumnWidth(6, 20);
	xls.SetColumnWidth(7, 20);
	xls.SetColumnWidth(8, 15);
	xls.SetColumnWidth(9, 20);
	xls.SetColumnWidth(10, 30);
	xls.SetColumnWidth(11, 30);
	xls.SetColumnWidth(12, 30);
	xls.SetCellMergedColumns(0, 0, 4);
	xls.SetCellMergedColumns(0, 1, 4);
	xls.SetCellMergedColumns(0, 2, 10);
	xls.SetCellMergedColumns(0, 3, 10);
	xls.SetCellText(0, 0, pMF->m_CompanyInfo.sc_name, 4098, true);
	xls.SetCellText(0, 1, pMF->GetCurrentDepartmentName(), 4098, true);
	xls.SetCellText(0, 2, _T("S\x1ED4 V\xC0O VI\x1EC6N"), 4098, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), m_szFromDate, m_szToDate);
	xls.SetCellText(0, 3, tmpStr, 4098);
	for (int i = 0; i < 13; i++)
		if (i != 4 && i != 5)
			xls.SetCellMergedRows(i, 4, 2);
	xls.SetCellMergedColumns(4, 4, 2);
	xls.SetCellText(0, 4, _T("STT"), 528386, true);
	xls.SetCellText(1, 4, _T("S\x1ED1 BA"), 528386, true);
	xls.SetCellText(2, 4, _T("S\x1ED1 KCB"), 528386, true);
	xls.SetCellText(3, 4, _T("H\x1ECD t\xEAn BN"), 528386, true);
	xls.SetCellText(4, 4, _T("Tu\x1ED5i"), 528386, true);
	xls.SetCellText(6, 4, _T("Th\x1EBB BH"), 528386, true);
	xls.SetCellText(7, 4, _T("\x110\x1ED1i t\x1B0\x1EE3ng"), 528386, true);
	xls.SetCellText(8, 4, _T("Ngh\x1EC1 nghi\x1EC7p"), 528386, true);
	xls.SetCellText(9, 4, _T("Ng\xE0y v\xE0o vi\x1EC7n"), 528386, true);
	xls.SetCellText(10, 4, _T("\x110\x1ECB\x61 \x63h\x1EC9"), 528386, true);
	xls.SetCellText(11, 4, _T("\x43h\x1EA9n \x111o\xE1n KB"), 528386, true);
	xls.SetCellText(12, 4, _T("\x43h\x1EA9n \x111o\xE1n KÐT"), 528386, true);
	xls.SetCellText(4, 5, _T("N\x61m"), 528386, true);
	xls.SetCellText(5, 5, _T("N\x1EEF"), 528386, true);

	nRow = 6;
	arrField.Add(_T("recordno"));
	arrField.Add(_T("docno"));
	arrField.Add(_T("pname"));
	arrField.Add(_T("maleage"));
	arrField.Add(_T("femaleage"));
	arrField.Add(_T("cardno"));
	arrField.Add(_T("obj"));
	arrField.Add(_T("occupation"));
	arrField.Add(_T("admitdate"));
	arrField.Add(_T("addr"));
	arrField.Add(_T("hd_diagnostic"));
	arrField.Add(_T("htr_maindisease"));
	while (!rs.IsEOF())
	{
		xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		for (int i = 0; i < arrField.GetCount(); i++)
		{
			if (i == 8)
				xls.SetCellText(nCol + i + 1, nRow, rs.GetValue(arrField.GetAt(i)), FMT_DATETIME);
			else
				xls.SetCellText(nCol + i + 1, nRow, rs.GetValue(arrField.GetAt(i)), FMT_TEXT |FMT_WRAPING);
		}
		nRow++;
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\So vao vien.xls"));
} 

CString CAdmitPatientBook::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND htr_admitdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szSQL.Format(_T(" SELECT   hd_docno docno, ") \
				_T("           htr_recordno recordno, ") \
				_T("           Get_patientname(hd_docno)                           pname, ") \
				_T("           CASE WHEN hp_sex = 'M' THEN Hms_getage(htr_admitdate, hp_birthdate) ") \
				_T("           ELSE '' ") \
				_T("           END                                              AS maleage, ") \
				_T("           CASE WHEN hp_sex = 'F' THEN Hms_getage(htr_admitdate, hp_birthdate) ") \
				_T("           ELSE '' ") \
				_T("           END                                              AS femaleage, ") \
				_T("           hd_cardno                                        cardno, ") \
				_T("           Get_objectname(hd_object)                        obj, ") \
				_T("           Get_syssel_desc('sys_occupation', hp_occupation) occupation, ") \
				_T("           htr_admitdate                                    admitdate, ") \
				_T("           Hms_getaddress(hp_provid, hp_distid, hp_villid) addr, ") \
				_T("           hd_diagnostic, ") \
				_T("           htr_maindisease ") \
				_T(" FROM      hms_treatment_record ") \
				_T(" LEFT JOIN hms_patient ON ( hp_patientno = htr_patientno ) ") \
				_T(" LEFT JOIN hms_doc ON ( htr_docno = hd_docno ) ") \
				_T(" WHERE htr_idx = 1 AND htr_deptid = '%s' AND htr_status <> 'A' %s") \
				_T(" ORDER BY htr_admitdate, docno"), pMF->GetCurrentDepartmentID(), szWhere);
	return szSQL;
}