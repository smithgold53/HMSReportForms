#include "stdafx.h"
#include "PACSSettlementChemistryMaterial.h"
#include "HMSMainFrame.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSSettlementChemistryMaterial *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPACSSettlementChemistryMaterial *)pWnd)->OnToDateCheckValue();
} 
static void _OnStockSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPACSSettlementChemistryMaterial* )pWnd)->OnStockSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStockSelendokFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnStockSelendok();
}
/*static void _OnStockSetfocusFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnStockKillfocus();
}*/
/*static void _OnStockKillfocusFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnStockKillfocus();
}*/
static long _OnStockLoadDataFnc(CWnd *pWnd){
	return ((CPACSSettlementChemistryMaterial *)pWnd)->OnStockLoadData();
}
/*static void _OnStockAddNewFnc(CWnd *pWnd){
	((CPACSSettlementChemistryMaterial *)pWnd)->OnStockAddNew();
}*/
static void _OnInPatientSelectFnc(CWnd *pWnd){
	  ((CPACSSettlementChemistryMaterial*)pWnd)->OnInPatientSelect();
}
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	  ((CPACSSettlementChemistryMaterial*)pWnd)->OnOutPatientSelect();
}
static void _OnExportSelectFnc(CWnd *pWnd){
	CPACSSettlementChemistryMaterial *pVw = (CPACSSettlementChemistryMaterial *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddPACSSettlementChemistryMaterialFnc(CWnd *pWnd){
	 return ((CPACSSettlementChemistryMaterial*)pWnd)->OnAddPACSSettlementChemistryMaterial();
} 
static int _OnEditPACSSettlementChemistryMaterialFnc(CWnd *pWnd){
	 return ((CPACSSettlementChemistryMaterial*)pWnd)->OnEditPACSSettlementChemistryMaterial();
} 
static int _OnDeletePACSSettlementChemistryMaterialFnc(CWnd *pWnd){
	 return ((CPACSSettlementChemistryMaterial*)pWnd)->OnDeletePACSSettlementChemistryMaterial();
} 
static int _OnSavePACSSettlementChemistryMaterialFnc(CWnd *pWnd){
	 return ((CPACSSettlementChemistryMaterial*)pWnd)->OnSavePACSSettlementChemistryMaterial();
} 
static int _OnCancelPACSSettlementChemistryMaterialFnc(CWnd *pWnd){
	 return ((CPACSSettlementChemistryMaterial*)pWnd)->OnCancelPACSSettlementChemistryMaterial();
} 
CPACSSettlementChemistryMaterial::CPACSSettlementChemistryMaterial(CWnd *pParent){

	m_nDlgWidth = 420;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CPACSSettlementChemistryMaterial::~CPACSSettlementChemistryMaterial(){
}
void CPACSSettlementChemistryMaterial::OnCreateComponents(){
	m_wndReportParameter.Create(this, _T("Report Parameters"), 5, 5, 410, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 205, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 210, 30, 290, 55);
	m_wndToDate.Create(this,295, 30, 405, 55); 
	m_wndStockLabel.Create(this, _T("Stock"), 10, 60, 90, 85);
	m_wndStock.Create(this,95, 60, 205, 85); 
	m_wndInPatient.Create(this, _T("InPatient"), 210, 60, 290, 85);
	m_wndOutPatient.Create(this, _T("OutPatient"), 296, 60, 376, 85);
	m_wndExport.Create(this, _T("Export"), 330, 95, 410, 120);

}
void CPACSSettlementChemistryMaterial::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndStock.SetCheckValue(true);
	m_wndStock.LimitText(35);


	m_wndStock.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndStock.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CPACSSettlementChemistryMaterial::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndStock.SetEvent(WE_SELENDOK, _OnStockSelendokFnc);
	//m_wndStock.SetEvent(WE_SETFOCUS, _OnStockSetfocusFnc);
	//m_wndStock.SetEvent(WE_KILLFOCUS, _OnStockKillfocusFnc);
	m_wndStock.SetEvent(WE_SELCHANGE, _OnStockSelectChangeFnc);
	m_wndStock.SetEvent(WE_LOADDATA, _OnStockLoadDataFnc);
	//m_wndStock.SetEvent(WE_ADDNEW, _OnStockAddNewFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	m_wndStock.SetCheckBox(true);
	SetMode(VM_EDIT);
}
void CPACSSettlementChemistryMaterial::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStock.GetDlgCtrlID(), m_szStockKey);
	DDX_Radio(pDX, m_wndInPatient.GetDlgCtrlID(), m_nInPatient);
	DDX_Radio(pDX, m_wndOutPatient.GetDlgCtrlID(), m_nOutPatient);

}
void CPACSSettlementChemistryMaterial::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPACSSettlementChemistryMaterial::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CPACSSettlementChemistryMaterial::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStockKey.Empty();
	m_nInPatient=0;
	m_nOutPatient=0;

}
int CPACSSettlementChemistryMaterial::SetMode(int nMode){
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
 			EnableButtons(TRUE, 0, -1);
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
/*void CPACSSettlementChemistryMaterial::OnFromDateChange(){
	
} */
/*void CPACSSettlementChemistryMaterial::OnFromDateSetfocus(){
	
} */
/*void CPACSSettlementChemistryMaterial::OnFromDateKillfocus(){
	
} */
int CPACSSettlementChemistryMaterial::OnFromDateCheckValue(){
	return 0;
} 
/*void CPACSSettlementChemistryMaterial::OnToDateChange(){
	
} */
/*void CPACSSettlementChemistryMaterial::OnToDateSetfocus(){
	
} */
/*void CPACSSettlementChemistryMaterial::OnToDateKillfocus(){
	
} */
int CPACSSettlementChemistryMaterial::OnToDateCheckValue(){
	return 0;
} 
void CPACSSettlementChemistryMaterial::OnStockSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CPACSSettlementChemistryMaterial::OnStockSelendok(){
	 
}
/*void CPACSSettlementChemistryMaterial::OnStockSetfocus(){
	
}*/
/*void CPACSSettlementChemistryMaterial::OnStockKillfocus(){
	
}*/
long CPACSSettlementChemistryMaterial::OnStockLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndStock.IsSearchKey() && !m_szStockKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szStockKey
};
	m_wndStock.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndStock.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CPACSSettlementChemistryMaterial::OnStockAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CPACSSettlementChemistryMaterial::OnInPatientSelect(){
	
}
void CPACSSettlementChemistryMaterial::OnOutPatientSelect(){
	
}
void CPACSSettlementChemistryMaterial::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();	CRecord rs(&pMF->m_db);	CString szSQL, tmpStr, szTemp, szWhere;	UpdateData(TRUE);	BeginWaitCursor();	szWhere.Empty();	int nRow =0, nCol = 0;	CExcel xls;	xls.CreateSheet(1);	xls.SetWorksheet(0);	xls.SetColumnWidth(0, 8);	xls.SetColumnWidth(1, 8);	xls.SetColumnWidth(2, 8);	xls.SetColumnWidth(3, 8);	xls.SetColumnWidth(4, 8);	xls.SetColumnWidth(5, 8);	xls.SetColumnWidth(6, 8);	xls.SetColumnWidth(7, 8);	xls.SetColumnWidth(8, 8);	xls.SetColumnWidth(9, 8);	xls.SetColumnWidth(10, 8);	xls.SetColumnWidth(11, 8);	xls.SetColumnWidth(12, 8);	xls.SetColumnWidth(13, 8);	xls.SetColumnWidth(14, 8);	xls.SetColumnWidth(15, 8);	xls.SetColumnWidth(16, 8);	xls.SetColumnWidth(17, 8);	xls.SetColumnWidth(18, 8);	xls.SetColumnWidth(19, 8);	xls.SetColumnWidth(20, 8);	xls.SetColumnWidth(21, 8);	xls.SetColumnWidth(22, 8);	xls.SetCellText(0, 0, _T("\x42\x1EC6NH VI\x1EC6N TWQ\x110 \x31\x30\x38"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(0, 1, _T("KHO\x41 \x43H\x1EA8N \x110O\xC1N H\xCCNH \x1EA2NH"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 13);	xls.SetCellText(0, 3, _T("\x42\xC1O \x43\xC1O QUY\x1EBET TO\xC1N HO\xC1 \x43H\x1EA4T, V\x1EACT T\x1AF TI\xCAU H\x41O \x43HO \x58\xC9T NGHI\x1EC6M"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 14);	xls.SetCellText(0, 4, _T("T\x1EEB ng\xE0y  th\xE1ng  \x111\x1EBFn ng\xE0y  th\xE1ng  n\x103m"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 14);	xls.SetCellText(0, 6, _T("T\xEAn ho\xE1 \x63h\x1EA5t"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(1, 6, _T("\x110VT"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(2, 6, _T("T\x1ED3n \x63huy\x1EC3n s\x61ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(3, 6, _T("L\x129nh trong th\xE1ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(4, 6, _T("\x43\x1ED9ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(5, 6, _T("\x110\x1A1n gi\xE1"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(6, 6, _T("Gi\xE1 tr\x1ECB l\x129nh"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(7, 0, _T("\x43\x1ED8NG HO\xC0 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(7, 1, _T("\x110\x1ED9\x63 l\x1EADp - T\x1EF1 \x64o - H\x1EA1nh ph\xFA\x63"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 13);	xls.SetCellText(7, 6, _T("Qu\xE2n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(7, 7, _T("SL"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(8, 7, _T("T.Ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(9, 6, _T("\x42H qu\xE2n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(9, 7, _T("SL"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(10, 7, _T("T.Ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(11, 6, _T("\x42H kh\xE1\x63"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(11, 7, _T("SL"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(12, 7, _T("T.Ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(13, 6, _T("\x41\x31\x31"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(13, 7, _T("SL"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(14, 7, _T("T.Ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(15, 6, _T("\x43h\xEDnh s\xE1\x63h"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(15, 7, _T("SL"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(16, 7, _T("T.Ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(17, 6, _T("\x44\x1ECB\x63h v\x1EE5"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(17, 7, _T("SL"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(18, 7, _T("T.Ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(19, 6, _T("T\x1ED5ng gi\xE1 tr\x1ECB s\x1EED \x64\x1EE5ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(19, 7, _T("SL"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(20, 7, _T("T.Ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(21, 6, _T("\x43huy\x1EC3n  s\x61ng th\xE1ng s\x61u"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(21, 7, _T("SL"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(22, 7, _T("T.Ti\x1EC1n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetMerge(0, 0, 0, 5);	xls.SetMerge(1, 1, 0, 5);	xls.SetMerge(3, 3, 0, 22);	xls.SetMerge(4, 4, 0, 22);	xls.SetMerge(6, 7, 0, 0);	xls.SetMerge(6, 7, 1, 1);	xls.SetMerge(6, 7, 2, 2);	xls.SetMerge(6, 7, 3, 3);	xls.SetMerge(6, 7, 4, 4);	xls.SetMerge(6, 7, 5, 5);	xls.SetMerge(6, 7, 6, 6);	xls.SetMerge(0, 0, 7, 15);	xls.SetMerge(1, 1, 7, 15);	xls.SetMerge(6, 6, 7, 8);	xls.SetMerge(6, 6, 9, 10);	xls.SetMerge(6, 6, 11, 12);	xls.SetMerge(6, 6, 13, 14);	xls.SetMerge(6, 6, 15, 16);	xls.SetMerge(6, 6, 17, 18);	xls.SetMerge(5, 5, 18, 22);	xls.SetMerge(6, 6, 19, 20);	xls.SetMerge(6, 6, 21, 22);	szSQL.Format(_T(""));	rs.ExecSQL(szSQL);	nRow = 8;	nCol = 0;	while(!rs.IsEOF()){		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+0, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+1, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+2, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+3, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+4, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+5, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+6, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+7, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+8, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+9, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+10, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+11, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+12, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+13, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+14, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+15, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+16, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+17, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+18, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+19, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+20, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+21, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T(""), tmpStr);		xls.SetCellText(nCol+22, nRow, tmpStr, FMT_TEXT);		nRow++;		rs.MoveNext();	}	EndWaitCursor();	xls.Save(_T("Exports\\quyettoanduocvattu.xls"));
} 
int CPACSSettlementChemistryMaterial::OnAddPACSSettlementChemistryMaterial(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CPACSSettlementChemistryMaterial::OnEditPACSSettlementChemistryMaterial(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CPACSSettlementChemistryMaterial::OnDeletePACSSettlementChemistryMaterial(){
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
 		OnCancelPACSSettlementChemistryMaterial();
 	}
	return 0;
}
int CPACSSettlementChemistryMaterial::OnSavePACSSettlementChemistryMaterial(){
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
 		//OnPACSSettlementChemistryMaterialListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CPACSSettlementChemistryMaterial::OnCancelPACSSettlementChemistryMaterial(){
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
int CPACSSettlementChemistryMaterial::OnPACSSettlementChemistryMaterialListLoadData(){
	return 0;
}
