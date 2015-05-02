#include "stdafx.h"
#include "PMSC11AidPrescriptionList.h"
#include "MainFrame_E10.h"
#include "ReportDocument.h"
#include ".\pmsc11aidprescriptionlist.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CPMSC11AidPrescriptionList *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CPMSC11AidPrescriptionList *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CPMSC11AidPrescriptionList *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSC11AidPrescriptionList *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CPMSC11AidPrescriptionList *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CPMSC11AidPrescriptionList *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CPMSC11AidPrescriptionList *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CPMSC11AidPrescriptionList *)pWnd)->OnToDateCheckValue();
} 
static void _OnDoctorSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CPMSC11AidPrescriptionList* )pWnd)->OnDoctorSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDoctorSelendokFnc(CWnd *pWnd){
	((CPMSC11AidPrescriptionList *)pWnd)->OnDoctorSelendok();
}
/*static void _OnDoctorSetfocusFnc(CWnd *pWnd){
	((CPMSC11AidPrescriptionList *)pWnd)->OnDoctorKillfocus();
}*/
/*static void _OnDoctorKillfocusFnc(CWnd *pWnd){
	((CPMSC11AidPrescriptionList *)pWnd)->OnDoctorKillfocus();
}*/
static long _OnDoctorLoadDataFnc(CWnd *pWnd){
	return ((CPMSC11AidPrescriptionList *)pWnd)->OnDoctorLoadData();
}
/*static void _OnDoctorAddNewFnc(CWnd *pWnd){
	((CPMSC11AidPrescriptionList *)pWnd)->OnDoctorAddNew();
}*/
static void _OnByServPriceSelectFnc(CWnd *pWnd){
	 ((CPMSC11AidPrescriptionList*)pWnd)->OnByServPriceSelect();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CPMSC11AidPrescriptionList *pVw = (CPMSC11AidPrescriptionList *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CPMSC11AidPrescriptionList *pVw = (CPMSC11AidPrescriptionList *)pWnd;
	pVw->OnExportSelect();
} 
CPMSC11AidPrescriptionList::CPMSC11AidPrescriptionList(CWnd *pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CPMSC11AidPrescriptionList::~CPMSC11AidPrescriptionList(){
}
void CPMSC11AidPrescriptionList::OnCreateComponents(){
	m_wndC11AidPrescriptionList.Create(this, _T("Export Sheet List"), 5, 5, 575, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndDoctorLabel.Create(this, _T("Doctor"), 10, 60, 90, 85);
	m_wndDoctor.Create(this,95, 60, 570, 85); 
	m_wndByServPrice.Create(this, _T("By Serv Price"), 5, 95, 100, 120);
	//m_wndPrintPreview.Create(this, _T("&Print Preview"), 385, 95, 495, 120);
	m_wndExport.Create(this, _T("&Export"), 500, 95, 575, 120);


}
void CPMSC11AidPrescriptionList::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	m_wndDoctor.SetCheckValue(true);
	m_wndDoctor.LimitText(35);


	m_wndDoctor.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	m_wndDoctor.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);

}

void CPMSC11AidPrescriptionList::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndDoctor.SetEvent(WE_SELENDOK, _OnDoctorSelendokFnc);
	//m_wndDoctor.SetEvent(WE_SETFOCUS, _OnDoctorSetfocusFnc);
	//m_wndDoctor.SetEvent(WE_KILLFOCUS, _OnDoctorKillfocusFnc);
	m_wndDoctor.SetEvent(WE_SELCHANGE, _OnDoctorSelectChangeFnc);
	m_wndDoctor.SetEvent(WE_LOADDATA, _OnDoctorLoadDataFnc);
	//m_wndDoctor.SetEvent(WE_ADDNEW, _OnDoctorAddNewFnc);
	m_wndByServPrice.SetEvent(WE_CLICK, _OnByServPriceSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CPMSC11AidPrescriptionList::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDoctor.GetDlgCtrlID(), m_szDoctorKey);
	DDX_Check(pDX, m_wndByServPrice.GetDlgCtrlID(), m_bByServPrice);

}
void CPMSC11AidPrescriptionList::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDoctorKey.Empty();
	m_bByServPrice = false;
}

int CPMSC11AidPrescriptionList::SetMode(int nMode){
 		int nOldMode = GetMode(); 
 		CGuiView::SetMode(nMode); 
 		CMainFrame_E10 *pMF = (CMainFrame_E10 *) AfxGetMainWnd(); 
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

/*void CPMSC11AidPrescriptionList::OnFromDateChange(){
	
} */
/*void CPMSC11AidPrescriptionList::OnFromDateSetfocus(){
	
} */
/*void CPMSC11AidPrescriptionList::OnFromDateKillfocus(){
	
} */
int CPMSC11AidPrescriptionList::OnFromDateCheckValue(){
	return 0;
}
 
/*void CPMSC11AidPrescriptionList::OnToDateChange(){
	
} */
/*void CPMSC11AidPrescriptionList::OnToDateSetfocus(){
	
} */
/*void CPMSC11AidPrescriptionList::OnToDateKillfocus(){
	
} */
int CPMSC11AidPrescriptionList::OnToDateCheckValue(){
	return 0;
}
 
void CPMSC11AidPrescriptionList::OnDoctorSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
 
void CPMSC11AidPrescriptionList::OnDoctorSelendok(){
	 
}

/*void CPMSC11AidPrescriptionList::OnDoctorSetfocus(){
	
}*/
/*void CPMSC11AidPrescriptionList::OnDoctorKillfocus(){
	
}*/
long CPMSC11AidPrescriptionList::OnDoctorLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDoctor.IsSearchKey() && ToInt(m_szDoctorKey) > 0){
		szWhere.Format(_T(" AND su_userid=%d"), ToInt(m_szDoctorKey));
	};
	//szWhere.AppendFormat(_T(" and pmssl_storage_id in(%s) "), pMF->m_szStorages);
	m_wndDoctor.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT su_userid as id, su_name as name FROM sys_user WHERE su_groupid = 'D' %s ORDER BY su_userid"), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDoctor.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}

/*void CPMSC11AidPrescriptionList::OnDoctorAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
	void CPMSC11AidPrescriptionList::OnByServPriceSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
}
void CPMSC11AidPrescriptionList::OnPrintPreviewSelect(){

}
 
void CPMSC11AidPrescriptionList::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	CRecord rs(&pMF->m_db), ors(&pMF->m_db);
	CString szSQL, szSubSQL, tmpStr;
	CString szNewDoc, szOldDoc, szOrderID;
	int nRow = 0, nCol = 0;
	int nOrderCnt = 0;
	int nIdx = 1, nCnt = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
		return;
	else
		rs.GetValue(_T("hpo_doctor"), szOldDoc);
	xls.SetColumnWidth(0, 3);
	xls.SetColumnWidth(1, 20);
	xls.SetColumnWidth(2, 20);
	xls.SetColumnWidth(3, 6);
	//xls.SetColumnWidth(5, 25);
	xls.SetColumnWidth(6, 20);
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 11);
	xls.SetCellMergedColumns(nCol, nRow + 3, 11);
	xls.SetCellMergedColumns(nCol, nRow + 4, 4);
	xls.SetCellText(nCol, nRow, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, _T("\x42\xC1O \x43\xC1O K\xCA \x110\x1A0N THU\x1ED0\x43 H\x1ED6 TR\x1EE2 T\x1EA0I PH\xD2NG KH\xC1M \x43\x31-\x31"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	//tmpStr.Format(_T("Kho : %s"), m_wndDoctor.GetCurrent(1));
	//xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT, false, 11);

	//Column Heading
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("\x42\xE1\x63 s\x1EF9"));
	arrCol.Add(_T("T\xEAn thu\x1ED1\x63"));
	arrCol.Add(_T("\x110\x1A1n v\x1ECB"));
	arrCol.Add(_T("S\x1ED1 l\x1B0\x1EE3ng"));
	arrCol.Add(_T("Th\xE0nh ti\x1EC1n"));
	arrCol.Add(_T("Nh\xF3m thu\x1ED1\x63"));
	arrCol.Add(_T("T\x42 ti\x1EC1n thu\x1ED1\x63 \x62\x1ED5/ \x111\x1A1n"));
	arrCol.Add(_T("T\x1EC9 l\x1EC7 thu\x1ED1\x63 \x62\x1ED5/ \x111\x1A1n"));
	nRow = 5;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10); 
	}
	nRow = 7;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("hpo_doctor"), szNewDoc);
		//_debug(_T("New%s:Old%s:Cnt%d"), szNewDoc, szOldDoc, nOrderCnt);
		if (szNewDoc != szOldDoc)
		{
			nRow = Exp(&xls, szOrderID, nRow, nIdx, nOrderCnt, szOldDoc);
			nOrderCnt = 0;
			szOldDoc = szNewDoc;
			szOrderID.Empty();
			nIdx++;
			tmpStr.Format(_T("%d"), nIdx);
			xls.SetCellText(0, nRow, ToString(nIdx), FMT_INTEGER);
		}
		if (!szOrderID.IsEmpty())
			szOrderID += _T(", ");
		szOrderID.AppendFormat(_T("'%s'"), rs.GetValue(_T("hpo_orderid")));
		nOrderCnt++;
		rs.MoveNext();
	}	
	nRow = Exp(&xls, szOrderID, nRow, nIdx, nOrderCnt, szOldDoc);
	xls.Save(_T("Exports\\Thuoc bo C11.xls"));
}
CString CPMSC11AidPrescriptionList::GetQueryString(){
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND hfe_createddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szSQL.Format(_T(" SELECT") \
				_T("   DISTINCT") \
				_T("   hpo_doctor,") \
				_T("   hpo_orderid") \
				_T(" FROM   hms_pharmaorder") \
				_T(" LEFT JOIN hms_fee_invoice ON (hfe_invoiceno = hpo_invoiceno)") \
				_T("        LEFT JOIN hms_pharmaorderline_view ON (hpo_orderid=hpol_orderid)") \
				_T(" WHERE  hpol_producttype IN ('A1160', 'A1170', 'A1180', 'A1190')") \
				_T("        %s") \
				_T(" ORDER  BY hpo_doctor,") \
				_T("           hpo_orderid"), szWhere);
	return szSQL;
}
CString CPMSC11AidPrescriptionList::GetSubQueryString(CString szOrderID){
	CString szSQL;
	szSQL.Format(_T(" SELECT") \
				_T("   hpol_product_id,") \
				_T("   hpol_productname,") \
				_T("   hpol_productuom,") \
				_T("   hpol_producttype,") \
				_T("   Sum(hpol_qtyissue) AS qty,") \
				_T("   Sum(aided)         AS aided,") \
				_T("   Sum(ttl)           AS ttl") \
				_T(" FROM   (SELECT") \
				_T("           hpol_product_id,") \
				_T("		   hpol_productname,") \
				_T("		   hpol_productuom,") \
				_T("           hpol_qtyissue,") \
				_T("           product_saleprice1,") \
				_T("           CASE WHEN hpol_producttype IN ('A1160', 'A1170', 'A1180', 'A1190') THEN hpol_qtyissue*product_saleprice1") \
				_T("           ELSE 0") \
				_T("           END                          AS aided,") \
				_T("           hpol_qtyissue*product_saleprice1 AS ttl,") \
				_T("           hpol_producttype") \
				_T("         FROM   hms_pharmaorderline_view") \
				_T("		 LEFT JOIN m_product_view ON (product_id = hpol_product_id)") \
				_T("         WHERE  hpol_orderid IN (%s))") \
				_T(" GROUP  BY hpol_product_id,") \
				_T("		   hpol_productname,") \
				_T("		   hpol_productuom,") \
				_T("           hpol_producttype "), szOrderID);
	return szSQL;
}
int CPMSC11AidPrescriptionList::Exp(CExcel *xls, CString szOrderID, int nRow, int nIdx, int nOrderCnt, CString szOldDoc){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord ors(&pMF->m_db);
	CString szSubSQL, tmpStr;
	double nAmt = 0;
	long double nAidedAmt = 0, nTtlAmt = 0;
	bool bShow = false;
	szSubSQL = GetSubQueryString(szOrderID);
	ors.ExecSQL(szSubSQL);
	xls->SetCellText(0, nRow, ToString(nIdx), FMT_INTEGER);
	xls->SetCellText(1, nRow, pMF->GetUserName(szOldDoc), FMT_TEXT);
	while(!ors.IsEOF())
	{
		bShow = ors.GetValue(_T("aided"))==_T("0")?false:true;
		if (bShow)
		{
			xls->SetCellText(2, nRow, ors.GetValue(_T("hpol_productname")), FMT_TEXT);
			xls->SetCellText(3, nRow, ors.GetValue(_T("hpol_productuom")), FMT_TEXT);
			ors.GetValue(_T("qty"), nAmt);
			xls->SetCellText(4, nRow, double2str(nAmt), FMT_NUMBER1);
			ors.GetValue(_T("ttl"), nAmt);
			xls->SetCellText(5, nRow, double2str(nAmt), FMT_NUMBER1);
			xls->SetCellText(6, nRow, pMF->GetProductTypeName(ors.GetValue(_T("hpol_producttype"))), FMT_TEXT);
			//xls->SetCellText(11, nRow, ToString(nCnt++), FMT_INTEGER);
			nRow++;
		}
		nAidedAmt += ToLong(ors.GetValue(_T("aided")));
		nTtlAmt += ToLong(ors.GetValue(_T("ttl")));
		//_debug(_T("Aid%ld:Ttl%ld"), nAidedAmt, nTtlAmt);
		ors.MoveNext();
	}
	//xls->SetCellText(5, nRow, _T("T\x1ED5ng ti\x1EC1n thu\x1ED1\x63 \x62\x1ED5"), FMT_TEXT);
	//xls->SetCellText(7, nRow, _T("Ti\x1EC1n T\x42"), FMT_TEXT);
	//xls->SetCellText(8, nRow, _T("%"), FMT_TEXT);
	//xls->SetCellText(9, nRow, _T("S\x1ED1 \x111\x1A1n \x63\xF3 thu\x1ED1\x63 \x62\x1ED5"), FMT_TEXT);
	//xls->SetCellText(10, nRow, _T("T\x1ED5ng ti\x1EC1n \x63\xE1\x63 \x111\x1A1n \x63\xF3 thu\x1ED1\x63 \x62\x1ED5"), FMT_TEXT);
	//nRow++;
	tmpStr.Format(_T("S\x1ED1 \x111\x1A1n: %d"), nOrderCnt);
	xls->SetCellText(1, nRow, tmpStr, FMT_TEXT);
	tmpStr.Format(_T("%f"), nAidedAmt);
	xls->SetCellText(5, nRow, tmpStr, FMT_NUMBER1);
	//Tien TB
	tmpStr.Format(_T("%f"), nAidedAmt/nOrderCnt);
	xls->SetCellText(7, nRow, tmpStr, FMT_NUMBER1);
	//%
	tmpStr.Format(_T("%f"), nAidedAmt*100/nTtlAmt);
	xls->SetCellText(8, nRow, tmpStr, FMT_NUMBER1);
	xls->SetCellText(9, nRow, int2str(nOrderCnt), FMT_INTEGER);
	//Total
	tmpStr.Format(_T("%f"), nTtlAmt);
	xls->SetCellText(10, nRow, tmpStr, FMT_NUMBER1);
	nRow++;
	return nRow;
}BEGIN_MESSAGE_MAP(CPMSC11AidPrescriptionList, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CPMSC11AidPrescriptionList::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}
