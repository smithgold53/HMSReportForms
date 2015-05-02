#include "stdafx.h"
#include "HMSMainFrame.h"
#include "CBloodBankReceiptReport.h"
static int _OnListCheckAllFnc(CWnd *pWnd){
	return ((CCBloodBankReceiptReport*)pWnd)->OnListCheckAll();
}
static int _OnListUncheckAllFnc(CWnd *pWnd){
	return ((CCBloodBankReceiptReport*)pWnd)->OnListUncheckAll();
}
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CCBloodBankReceiptReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CCBloodBankReceiptReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CCBloodBankReceiptReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CCBloodBankReceiptReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CCBloodBankReceiptReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CCBloodBankReceiptReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CCBloodBankReceiptReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CCBloodBankReceiptReport *)pWnd)->OnToDateCheckValue();
} 
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CCBloodBankReceiptReport*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CCBloodBankReceiptReport*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CCBloodBankReceiptReport*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CCBloodBankReceiptReport*)pWnd)->OnListDeleteItem();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CCBloodBankReceiptReport *pVw = (CCBloodBankReceiptReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CCBloodBankReceiptReport *pVw = (CCBloodBankReceiptReport *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddCBloodBankReceiptReportFnc(CWnd *pWnd){
	 return ((CCBloodBankReceiptReport*)pWnd)->OnAddCBloodBankReceiptReport();
} 
static int _OnEditCBloodBankReceiptReportFnc(CWnd *pWnd){
	 return ((CCBloodBankReceiptReport*)pWnd)->OnEditCBloodBankReceiptReport();
} 
static int _OnDeleteCBloodBankReceiptReportFnc(CWnd *pWnd){
	 return ((CCBloodBankReceiptReport*)pWnd)->OnDeleteCBloodBankReceiptReport();
} 
static int _OnSaveCBloodBankReceiptReportFnc(CWnd *pWnd){
	 return ((CCBloodBankReceiptReport*)pWnd)->OnSaveCBloodBankReceiptReport();
} 
static int _OnCancelCBloodBankReceiptReportFnc(CWnd *pWnd){
	 return ((CCBloodBankReceiptReport*)pWnd)->OnCancelCBloodBankReceiptReport();
} 
CCBloodBankReceiptReport::CCBloodBankReceiptReport(CWnd *pWnd){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CCBloodBankReceiptReport::~CCBloodBankReceiptReport(){
}
void CCBloodBankReceiptReport::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 6, 5, 576, 232);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 110, 55);
	m_wndFromDate.Create(this,115, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 395, 55);
	m_wndToDate.Create(this,400, 30, 570, 55); 
	m_wndList.Create(this,10, 60, 570, 227); 
	m_wndPrint.Create(this, _T("Print"), 404, 238, 484, 263);
	m_wndPrint.EnableWindow(FALSE);
	m_wndExport.Create(this, _T("&Export"), 488, 238, 568, 263);

	m_wndList.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndList.InsertColumn(1, _T("TYPE"), CFMT_TEXT, 0);
	m_wndList.InsertColumn(2, _T("\x43h\x1EBF Ph\x1EA9m"), CFMT_TEXT, 450);
	m_wndList.SetCheckBox(true);

}
void CCBloodBankReceiptReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);



}
void CCBloodBankReceiptReport::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);
	m_wndList.AddEvent(1, _T("Check All"), _OnListCheckAllFnc);
	m_wndList.AddEvent(2, _T("UnCheck All"), _OnListUncheckAllFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);
	OnListLoadData();

}
void CCBloodBankReceiptReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
int CCBloodBankReceiptReport::OnListCheckAll(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	for (int i=0; i< m_wndList.GetItemCount(); i++){
		m_wndList.SetCheck(i, true);
	}
	return 0;	
}

int CCBloodBankReceiptReport::OnListUncheckAll(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	for (int i=0; i< m_wndList.GetItemCount(); i++){
		m_wndList.SetCheck(i, false);
	}
	return 0;	
}
void CCBloodBankReceiptReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CCBloodBankReceiptReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CCBloodBankReceiptReport::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CCBloodBankReceiptReport::SetMode(int nMode){
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
/*void CCBloodBankReceiptReport::OnFromDateChange(){
	
} */
/*void CCBloodBankReceiptReport::OnFromDateSetfocus(){
	
} */
/*void CCBloodBankReceiptReport::OnFromDateKillfocus(){
	
} */
int CCBloodBankReceiptReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CCBloodBankReceiptReport::OnToDateChange(){
	
} */
/*void CCBloodBankReceiptReport::OnToDateSetfocus(){
	
} */
/*void CCBloodBankReceiptReport::OnToDateKillfocus(){
	
} */
int CCBloodBankReceiptReport::OnToDateCheckValue(){
	return 0;
} 
void CCBloodBankReceiptReport::OnListDblClick(){
	
} 
void CCBloodBankReceiptReport::OnListSelectChange(int nOldItem, int nNewItem){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CCBloodBankReceiptReport::OnListDeleteItem(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CCBloodBankReceiptReport::OnListLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndList.BeginLoad(); 
	int nCount = 0;
	szSQL.Format(_T(" SELECT * ") \
		_T(" FROM(select '6966,6972' id,1 type,'Kh\x1ED1i h\x1ED3ng \x63\x1EA7u + M\xE1u to\xE0n ph\x1EA7n 250' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6970,6973' id,2,'Kh\x1ED1i h\x1ED3ng \x63\x1EA7u + M\xE1u to\xE0n ph\x1EA7n 350' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6976' id,3,'Huy\x1EBFt T\x1B0\x1A1ng T\x1B0\x1A1i 150ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6975' id,4,'Huy\x1EBFt T\x1B0\x1A1ng T\x1B0\x1A1i 250ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6972' id,5,'Kh\x1ED1i h\x1ED3ng \x63\x1EA7u 250ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6973' id,6,'Kh\x1ED1i h\x1ED3ng \x63\x1EA7u 350ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6966' id,7,'M\xE1u to\xE0n ph\x1EA7n 250ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6970' id,8,'M\xE1u to\xE0n ph\x1EA7n \x33\x35\x30ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6977' id,9,'Ti\x1EC3u \x43\x1EA7u M\xE1y \x32\x35\x30ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6978' id,10,'Ti\x1EC3u \x43\x1EA7u Pool \x31\x35\x30ml' name") \
		_T(" from dual)") \
		_T(" Order by type"));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndList.AddItems(
			rs.GetValue(_T("id")),
			rs.GetValue(_T("type")),
			rs.GetValue(_T("name")),NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
}
void CCBloodBankReceiptReport::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CCBloodBankReceiptReport::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);	CString szSQL, tmpStr, szTemp, szWhere;	UpdateData(TRUE);	BeginWaitCursor();	szWhere.Empty();	int nRow =0, nCol = 0;	CExcel xls;	xls.CreateSheet(1);	xls.SetWorksheet(0);	xls.SetColumnWidth(0, 17);	xls.SetColumnWidth(1, 8);	xls.SetColumnWidth(2, 8);	xls.SetColumnWidth(3, 8);	xls.SetColumnWidth(4, 8);	xls.SetColumnWidth(5, 8);	xls.SetColumnWidth(6, 8);	xls.SetColumnWidth(7, 8);	xls.SetColumnWidth(8, 8);	xls.SetCellText(0, 0, _T("\x42\x1ED8 QU\x1ED0\x43 PH\xD2NG"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(0, 1, _T("\x42\x1EC6NH VI\x1EC6N TW QU\xC2N \x110\x1ED8I \x31\x30\x38"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(0, 2, _T("\x42\xC1O \x43\xC1O NH\x1EACP \x58U\x1EA4T T\x1ED2N \x43H\x1EBE PH\x1EA8M M\xC1U"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));	xls.SetCellText(0, 3, tmpStr, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	CRecord rs1(&pMF->m_db);	CRecord rs2(&pMF->m_db);	CString tmpStr1,tmpStr2,dv1,dv2,szList,szListNameStr;	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (!szList.IsEmpty())
			{
				szList += _T(", ");
				szListNameStr += _T(", ");
			}
			szListNameStr.AppendFormat(_T("%s"), m_wndList.GetItemText(i, 1));
			szList.AppendFormat(_T("%s"), m_wndList.GetItemText(i, 0));
		}

	}	tmpStr1.Format(_T("select min(substr(substr(mp_name,instr(mp_name,' ',-1,1)+1),1,3)) id from m_product where mp_org_id='BB' and mp_product_id in(%s) Order by mp_name"),szList);	rs1.ExecSQL(tmpStr1);	tmpStr2.Format(_T("select max(substr(substr(mp_name,instr(mp_name,' ',-1,1)+1),1,3)) id from m_product where mp_org_id='BB' and mp_product_id in(%s) Order by mp_name"),szList);	rs2.ExecSQL(tmpStr2);	dv1.Format(_T("%s ml"),rs1.GetValue(rs1.GetFieldName(0)));	dv2.Format(_T("%s ml"),rs2.GetValue(rs2.GetFieldName(0)));	if(rs1.GetValue(rs1.GetFieldName(0))==_T("150")&&rs2.GetValue(rs2.GetFieldName(0))==_T("350"))
	{
		AfxMessageBox(_T("L\x1ED7i !!! Vui l\xF2ng \x63h\x1ECDn t\x1EEBng \x63\x1EB7p \x63h\x1EBF ph\x1EA9m th\x65o \x111\x1A1n v\x1ECB "), MB_ICONASTERISK | MB_OK);
		return;
	}	xls.SetCellText(1, 5, _T("Nh\xF3m A"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(1, 6, dv1, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(2, 6, dv2, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(3, 5, _T("Nh\xF3m B"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(3, 6, dv1, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(4, 6, dv2, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(5, 5, _T("Nh\xF3m O"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(5, 6, dv1, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(6, 6, dv2, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(7, 5, _T("Nh\xF3m AB"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(7, 6, dv1, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(8, 6, dv2, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetMerge(0, 0, 0, 2);	xls.SetMerge(1, 1, 0, 2);	xls.SetMerge(2, 2, 0, 8);	xls.SetMerge(3, 3, 0, 8);	xls.SetMerge(5, 5, 1, 2);	xls.SetMerge(5, 5, 3, 4);	xls.SetMerge(5, 5, 5, 6);	xls.SetMerge(5, 5, 7, 8);	szSQL=GetQueryString();	rs.ExecSQL(szSQL);	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONASTERISK | MB_OK);
		return;
	}	nRow = 7;	nCol = 0;	while(!rs.IsEOF()){		rs.GetValue(_T("type"), tmpStr);		xls.SetCellText(nCol+0, nRow, tmpStr, FMT_TEXT);		rs.GetValue(_T("a250"), tmpStr);		xls.SetCellText(nCol+1, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("a350"), tmpStr);		xls.SetCellText(nCol+2, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("b250"), tmpStr);		xls.SetCellText(nCol+3, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("b350"), tmpStr);		xls.SetCellText(nCol+4, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("o250"), tmpStr);		xls.SetCellText(nCol+5, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("o350"), tmpStr);		xls.SetCellText(nCol+6, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("ab250"), tmpStr);		xls.SetCellText(nCol+7, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("ab350"), tmpStr);		xls.SetCellText(nCol+8, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		nRow++;		rs.MoveNext();	}	EndWaitCursor();	xls.Save(_T("Exports\\BAOCAONHAPXUATTONCHEPHAMMAU.xls"));
} 
int CCBloodBankReceiptReport::OnAddCBloodBankReceiptReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CCBloodBankReceiptReport::OnEditCBloodBankReceiptReport(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CCBloodBankReceiptReport::OnDeleteCBloodBankReceiptReport(){
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
 		OnCancelCBloodBankReceiptReport();
 	}
	return 0;
}
int CCBloodBankReceiptReport::OnSaveCBloodBankReceiptReport(){
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
 		//OnCBloodBankReceiptReportListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CCBloodBankReceiptReport::OnCancelCBloodBankReceiptReport(){
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
int CCBloodBankReceiptReport::OnCBloodBankReceiptReportListLoadData(){
	return 0;
}
CString CCBloodBankReceiptReport::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere, szList,szListNameStr;
	
	for (int i = 0; i < m_wndList.GetItemCount(); i++)
	{
		if (m_wndList.GetCheck(i))
		{
			if (!szList.IsEmpty())
			{
				szList += _T(", ");
				szListNameStr += _T(", ");
			}
			szListNameStr.AppendFormat(_T("%s"), m_wndList.GetItemText(i, 1));
			szList.AppendFormat(_T("%s"), m_wndList.GetItemText(i, 0));
		}

	}
	CRecord rs1(&pMF->m_db);	CRecord rs2(&pMF->m_db);	CString tmpStr1,tmpStr2,dv1,dv2;	int nCount=0,nCount1=0;	tmpStr1.Format(_T(" SELECT min(product_id1) product_id1") \
		_T(" FROM(SELECT trim(substr(substr(name,instr(name,' ',-1,1)+1),1,3)||''||id) product_id1") \
		_T(" FROM(select '6976' id,3,'Huyet tuong tuoi 150ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6975' id,4,'Huyet tuong tuoi 250ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6972' id,5,'Khoi hong cau 250ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6973' id,6,'Khoi hong cau 350ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6966' id,7,'Mau toan phan 250ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6970' id,8,'Mau toan phan 350ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6977' id,9,'Tieu cau may 250ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6978' id,10,'Tieu cau Pool 150ml' name") \
		_T(" from dual)") \
		_T(" Where id in(%s))"),szList);	rs1.ExecSQL(tmpStr1);	_fmsg(_T("%s"),tmpStr1);	tmpStr2.Format(_T(" SELECT max(product_id1) product_id1") \
		_T(" FROM(SELECT trim(substr(substr(name,instr(name,' ',-1,1)+1),1,3)||''||id) product_id1") \
		_T(" FROM(select '6976' id,3,'Huy\x1EBFt T\x1B0\x1A1ng T\x1B0\x1A1i 150ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6975' id,4,'Huy\x1EBFt T\x1B0\x1A1ng T\x1B0\x1A1i 250ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6972' id,5,'Kh\x1ED1i H\x1ED3ng \x43\x1EA7u 250ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6973' id,6,'Kh\x1ED1i H\x1ED3ng \x43\x1EA7u 350ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6966' id,7,'M\xE1u To\xE0n Ph\x1EA7n 250ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6970' id,8,'M\xE1u To\xE0n Ph\x1EA7n 350ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6977' id,9,'Ti\x1EC3u \x43\x1EA7u M\xE1y 250ml' name") \
		_T(" from dual") \
		_T(" union") \
		_T(" select '6978' id,10,'Ti\x1EC3u \x43\x1EA7u Pool 150ml' name") \
		_T(" from dual)") \
		_T(" Where id in(%s))"),szList);	rs2.ExecSQL(tmpStr2);
	if(rs1.GetValue(rs1.GetFieldName(0))==_T("2506966")&&rs2.GetValue(rs2.GetFieldName(0))==_T("3506973"))
	{
		szSQL.Format(_T(" SELECT 'S\x1ED1 l\x1B0u' type,sum(a250) a250,sum(a350) a350,") \
			_T("        sum(b250) b250,sum(b350) b350,") \
			_T("        sum(o250) o250,sum(o350) o350,") \
			_T("        sum(ab250) ab250,sum(ab350) ab350") \
			_T(" FROM(SELECT case when product_id in(6966,6972) and mbd_blood_type IN('A+','A-') then soluu else 0 end a250,") \
			_T("        case when product_id in(6970,6973) and mbd_blood_type IN('A+','A-') then soluu else 0 end a350,") \
			_T("        case when product_id in(6966,6972) and mbd_blood_type IN('B+','B-') then soluu else 0 end b250,") \
			_T("        case when product_id in(6970,6973) and mbd_blood_type IN('B+','B-') then soluu else 0 end b350,") \
			_T("        case when product_id in(6966,6972) and mbd_blood_type IN('O+','O-') then soluu else 0 end o250,") \
			_T("        case when product_id in(6970,6973) and mbd_blood_type IN('O+','O-') then soluu else 0 end o350,") \
			_T("        case when product_id in(6966,6972) and mbd_blood_type IN('AB+','AB-') then soluu else 0 end ab250,") \
			_T("        case when product_id in(6970,6973) and mbd_blood_type IN('AB+','AB-') then soluu else 0 end ab350") \
			_T(" FROM(SELECT product_id,") \
			_T("        mbd_blood_type,") \
			_T("        cast(trim(substr(substr(product_name,instr(product_name,' ',-1,1)+1),1,3)||''||product_id) as number) product_id1,") \
			_T("        product_name                      AS NAME,") \
			_T("        SUM(tondau3+nhap3) soluu,") \
			_T("        SUM(xuat3)                     AS xuat3,") \
			_T("        SUM(tondau3+nhap3-xuat3) AS conlai3") \
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
			_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = product_item_id)") \
			_T(" WHERE LENGTH(product_code) > 0 AND product_org_id = 'BB'") \
			_T(" GROUP BY product_name,product_id,mbd_blood_type))") \
			_T(" UNION ALL") \
			_T(" SELECT 'S\x1ED1 ph\xE1t' type,sum(a250) a250,sum(a350) a350,") \
			_T("        sum(b250) b250,sum(b350) b350,") \
			_T("        sum(o250) o250,sum(o350) o350,") \
			_T("        sum(ab250) ab250,sum(ab350) ab350") \
			_T(" FROM(SELECT case when product_id in(6966,6972) and mbd_blood_type IN('A+','A-') then xuat3 else 0 end a250,") \
			_T("        case when product_id in(6970,6973) and mbd_blood_type IN('A+','A-') then xuat3 else 0 end a350,") \
			_T("        case when product_id in(6966,6972) and mbd_blood_type IN('B+','B-') then xuat3 else 0 end b250,") \
			_T("        case when product_id in(6970,6973) and mbd_blood_type IN('B+','B-') then xuat3 else 0 end b350,") \
			_T("        case when product_id in(6966,6972) and mbd_blood_type IN('O+','O-') then xuat3 else 0 end o250,") \
			_T("        case when product_id in(6970,6973) and mbd_blood_type IN('O+','O-') then xuat3 else 0 end o350,") \
			_T("        case when product_id in(6966,6972) and mbd_blood_type IN('AB+','AB-') then xuat3 else 0 end ab250,") \
			_T("        case when product_id in(6970,6973) and mbd_blood_type IN('AB+','AB-') then xuat3 else 0 end ab350") \
			_T(" FROM(SELECT product_id,") \
			_T("        mbd_blood_type,") \
			_T("        cast(trim(substr(substr(product_name,instr(product_name,' ',-1,1)+1),1,3)||''||product_id) as number) product_id1,") \
			_T("        product_name                      AS NAME,") \
			_T("        SUM(tondau3+nhap3) soluu,") \
			_T("        SUM(xuat3)                     AS xuat3,") \
			_T("        SUM(tondau3+nhap3-xuat3) AS conlai3") \
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
			_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = product_item_id)") \
			_T(" WHERE LENGTH(product_code) > 0 AND product_org_id = 'BB'") \
			_T(" GROUP BY product_name,product_id,mbd_blood_type))") \
			_T(" UNION ALL") \
			_T(" SELECT '\x43\xF2n l\x1EA1i' type,sum(a250) a250,sum(a350) a350,") \
			_T("        sum(b250) b250,sum(b350) b350,") \
			_T("        sum(o250) o250,sum(o350) o350,") \
			_T("        sum(ab250) ab250,sum(ab350) ab350") \
			_T(" FROM(SELECT case when product_id in(6966,6972) and mbd_blood_type IN('A+','A-') then conlai3 else 0 end a250,") \
			_T("        case when product_id in(6970,6973) and mbd_blood_type IN('A+','A-') then conlai3 else 0 end a350,") \
			_T("        case when product_id in(6966,6972) and mbd_blood_type IN('B+','B-') then conlai3 else 0 end b250,") \
			_T("        case when product_id in(6970,6973) and mbd_blood_type IN('B+','B-') then conlai3 else 0 end b350,") \
			_T("        case when product_id in(6966,6972) and mbd_blood_type IN('O+','O-') then conlai3 else 0 end o250,") \
			_T("        case when product_id in(6970,6973) and mbd_blood_type IN('O+','O-') then conlai3 else 0 end o350,") \
			_T("        case when product_id in(6966,6972) and mbd_blood_type IN('AB+','AB-') then conlai3 else 0 end ab250,") \
			_T("        case when product_id in(6970,6973) and mbd_blood_type IN('AB+','AB-') then conlai3 else 0 end ab350") \
			_T(" FROM(SELECT product_id,") \
			_T("        mbd_blood_type,") \
			_T("        cast(trim(substr(substr(product_name,instr(product_name,' ',-1,1)+1),1,3)||''||product_id) as number) product_id1,") \
			_T("        product_name                      AS NAME,") \
			_T("        SUM(tondau3+nhap3) soluu,") \
			_T("        SUM(xuat3)                     AS xuat3,") \
			_T("        SUM(tondau3+nhap3-xuat3) AS conlai3") \
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
			_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = product_item_id)") \
			_T(" WHERE LENGTH(product_code) > 0 AND product_org_id = 'BB'") \
			_T(" GROUP BY product_name,product_id,mbd_blood_type))"),m_szFromDate,m_szToDate,
			m_szFromDate,
			m_szFromDate,m_szToDate,
			m_szFromDate,
			m_szFromDate,m_szToDate,
			m_szFromDate);
	}
	else szSQL.Format(_T(" SELECT 'S\x1ED1 l\x1B0u' type,sum(a250) a250,sum(a350) a350,") \
		_T("        sum(b250) b250,sum(b350) b350,") \
		_T("        sum(o250) o250,sum(o350) o350,") \
		_T("        sum(ab250) ab250,sum(ab350) ab350") \
		_T(" FROM(SELECT case when product_id1 in(%s) and mbd_blood_type IN('A+','A-') then soluu else 0 end a250,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('A+','A-') then soluu else 0 end a350,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('B+','B-') then soluu else 0 end b250,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('B+','B-') then soluu else 0 end b350,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('O+','O-') then soluu else 0 end o250,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('O+','O-') then soluu else 0 end o350,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('AB+','AB-') then soluu else 0 end ab250,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('AB+','AB-') then soluu else 0 end ab350") \
		_T(" FROM(SELECT product_id,") \
		_T("        mbd_blood_type,") \
		_T("        cast(trim(substr(substr(product_name,instr(product_name,' ',-1,1)+1),1,3)||''||product_id) as number) product_id1,") \
		_T("        product_name                      AS NAME,") \
		_T("        SUM(tondau3+nhap3) soluu,") \
		_T("        SUM(xuat3)                     AS xuat3,") \
		_T("        SUM(tondau3+nhap3-xuat3) AS conlai3") \
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
		_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = product_item_id)") \
		_T(" WHERE LENGTH(product_code) > 0 AND product_org_id = 'BB'") \
		_T(" GROUP BY product_name,product_id,mbd_blood_type))") \
		_T(" UNION ALL") \
		_T(" SELECT 'S\x1ED1 ph\xE1t' type,sum(a250) a250,sum(a350) a350,") \
		_T("        sum(b250) b250,sum(b350) b350,") \
		_T("        sum(o250) o250,sum(o350) o350,") \
		_T("        sum(ab250) ab250,sum(ab350) ab350") \
		_T(" FROM(SELECT case when product_id1 in(%s) and mbd_blood_type IN('A+','A-') then xuat3 else 0 end a250,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('A+','A-') then xuat3 else 0 end a350,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('B+','B-') then xuat3 else 0 end b250,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('B+','B-') then xuat3 else 0 end b350,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('O+','O-') then xuat3 else 0 end o250,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('O+','O-') then xuat3 else 0 end o350,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('AB+','AB-') then xuat3 else 0 end ab250,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('AB+','AB-') then xuat3 else 0 end ab350") \
		_T(" FROM(SELECT product_id,") \
		_T("        mbd_blood_type,") \
		_T("        cast(trim(substr(substr(product_name,instr(product_name,' ',-1,1)+1),1,3)||''||product_id) as number) product_id1,") \
		_T("        product_name                      AS NAME,") \
		_T("        SUM(tondau3+nhap3) soluu,") \
		_T("        SUM(xuat3)                     AS xuat3,") \
		_T("        SUM(tondau3+nhap3-xuat3) AS conlai3") \
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
		_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = product_item_id)") \
		_T(" WHERE LENGTH(product_code) > 0 AND product_org_id = 'BB'") \
		_T(" GROUP BY product_name,product_id,mbd_blood_type))") \
		_T(" UNION ALL") \
		_T(" SELECT '\x43\xF2n l\x1EA1i' type,sum(a250) a250,sum(a350) a350,") \
		_T("        sum(b250) b250,sum(b350) b350,") \
		_T("        sum(o250) o250,sum(o350) o350,") \
		_T("        sum(ab250) ab250,sum(ab350) ab350") \
		_T(" FROM(SELECT case when product_id1 in(%s) and mbd_blood_type IN('A+','A-') then conlai3 else 0 end a250,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('A+','A-') then conlai3 else 0 end a350,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('B+','B-') then conlai3 else 0 end b250,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('B+','B-') then conlai3 else 0 end b350,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('O+','O-') then conlai3 else 0 end o250,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('O+','O-') then conlai3 else 0 end o350,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('AB+','AB-') then conlai3 else 0 end ab250,") \
		_T("        case when product_id1 in(%s) and mbd_blood_type IN('AB+','AB-') then conlai3 else 0 end ab350") \
		_T(" FROM(SELECT product_id,") \
		_T("        mbd_blood_type,") \
		_T("        cast(trim(substr(substr(product_name,instr(product_name,' ',-1,1)+1),1,3)||''||product_id) as number) product_id1,") \
		_T("        product_name                      AS NAME,") \
		_T("        SUM(tondau3+nhap3) soluu,") \
		_T("        SUM(xuat3)                     AS xuat3,") \
		_T("        SUM(tondau3+nhap3-xuat3) AS conlai3") \
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
		_T(" LEFT JOIN m_blood_donation ON(mbd_product_item_id = product_item_id)") \
		_T(" WHERE LENGTH(product_code) > 0 AND product_org_id = 'BB'") \
		_T(" GROUP BY product_name,product_id,mbd_blood_type))"),rs1.GetValue(rs1.GetFieldName(0)),rs2.GetValue(rs2.GetFieldName(0)),rs1.GetValue(rs1.GetFieldName(0)),rs2.GetValue(rs2.GetFieldName(0)),rs1.GetValue(rs1.GetFieldName(0)),rs2.GetValue(rs2.GetFieldName(0)),rs1.GetValue(rs1.GetFieldName(0)),rs2.GetValue(rs2.GetFieldName(0)),
		m_szFromDate,m_szToDate,
		m_szFromDate,
		rs1.GetValue(rs1.GetFieldName(0)),rs2.GetValue(rs2.GetFieldName(0)),rs1.GetValue(rs1.GetFieldName(0)),rs2.GetValue(rs2.GetFieldName(0)),rs1.GetValue(rs1.GetFieldName(0)),rs2.GetValue(rs2.GetFieldName(0)),rs1.GetValue(rs1.GetFieldName(0)),rs2.GetValue(rs2.GetFieldName(0)),
		m_szFromDate,m_szToDate,
		m_szFromDate,
		rs1.GetValue(rs1.GetFieldName(0)),rs2.GetValue(rs2.GetFieldName(0)),rs1.GetValue(rs1.GetFieldName(0)),rs2.GetValue(rs2.GetFieldName(0)),rs1.GetValue(rs1.GetFieldName(0)),rs2.GetValue(rs2.GetFieldName(0)),rs1.GetValue(rs1.GetFieldName(0)),rs2.GetValue(rs2.GetFieldName(0)),
		m_szFromDate,m_szToDate,
		m_szFromDate);
	//QueryOpener(szSQL);
	return szSQL;
}