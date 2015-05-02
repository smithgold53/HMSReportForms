#include "stdafx.h"
#include "EMServiceDrugTrack.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMServiceDrugTrack *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMServiceDrugTrack *)pWnd)->OnToDateCheckValue();
} 
static void _OnDrugNameSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMServiceDrugTrack* )pWnd)->OnDrugNameSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDrugNameSelendokFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnDrugNameSelendok();
}
/*static void _OnDrugNameSetfocusFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnDrugNameKillfocus();
}*/
/*static void _OnDrugNameKillfocusFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnDrugNameKillfocus();
}*/
static long _OnDrugNameLoadDataFnc(CWnd *pWnd){
	return ((CEMServiceDrugTrack *)pWnd)->OnDrugNameLoadData();
}
/*static void _OnDrugNameAddNewFnc(CWnd *pWnd){
	((CEMServiceDrugTrack *)pWnd)->OnDrugNameAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMServiceDrugTrack *pVw = (CEMServiceDrugTrack *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMServiceDrugTrack *pVw = (CEMServiceDrugTrack *)pWnd;
	pVw->OnExportSelect();
} 
CEMServiceDrugTrack::CEMServiceDrugTrack(CWnd* pParent){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CEMServiceDrugTrack::~CEMServiceDrugTrack(){
}
void CEMServiceDrugTrack::OnCreateComponents(){
	m_wndServiceDrugTrack.Create(this, _T("Service Drug Track"), 5, 5, 490, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndDrugNameLabel.Create(this, _T("Drug Name"), 10, 60, 90, 85);
	m_wndDrugName.Create(this,95, 60, 485, 85); 
	//m_wndPrint.Create(this, _T("&Print"), 325, 100, 405, 125);
	m_wndExport.Create(this, _T("&Export"), 410, 100, 490, 125);

}
void CEMServiceDrugTrack::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndDrugName.SetCheckValue(true);
	m_wndDrugName.LimitText(1024);

	m_wndDrugName.InsertColumn(0, _T(""), CFMT_TEXT, 0);
	m_wndDrugName.InsertColumn(1, _T("Code"), CFMT_TEXT, 70);
	m_wndDrugName.InsertColumn(2, _T("Name"), CFMT_TEXT, 300);
	m_wndDrugName.Format(3, 2);
}
void CEMServiceDrugTrack::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndDrugName.SetEvent(WE_SELENDOK, _OnDrugNameSelendokFnc);
	//m_wndDrugName.SetEvent(WE_SETFOCUS, _OnDrugNameSetfocusFnc);
	//m_wndDrugName.SetEvent(WE_KILLFOCUS, _OnDrugNameKillfocusFnc);
	m_wndDrugName.SetEvent(WE_SELCHANGE, _OnDrugNameSelectChangeFnc);
	m_wndDrugName.SetEvent(WE_LOADDATA, _OnDrugNameLoadDataFnc);
	//m_wndDrugName.SetEvent(WE_ADDNEW, _OnDrugNameAddNewFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CEMServiceDrugTrack::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndDrugName.GetDlgCtrlID(), m_szDrugNameKey);

}
void CEMServiceDrugTrack::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szDrugNameKey.Empty();

}
int CEMServiceDrugTrack::SetMode(int nMode){
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
/*void CEMServiceDrugTrack::OnFromDateChange(){
	
} */
/*void CEMServiceDrugTrack::OnFromDateSetfocus(){
	
} */
/*void CEMServiceDrugTrack::OnFromDateKillfocus(){
	
} */
int CEMServiceDrugTrack::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMServiceDrugTrack::OnToDateChange(){
	
} */
/*void CEMServiceDrugTrack::OnToDateSetfocus(){
	
} */
/*void CEMServiceDrugTrack::OnToDateKillfocus(){
	
} */
int CEMServiceDrugTrack::OnToDateCheckValue(){
	return 0;
} 
void CEMServiceDrugTrack::OnDrugNameSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMServiceDrugTrack::OnDrugNameSelendok(){

}
/*void CEMServiceDrugTrack::OnDrugNameSetfocus(){
	
}*/
/*void CEMServiceDrugTrack::OnDrugNameKillfocus(){
	
}*/
long CEMServiceDrugTrack::OnDrugNameLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDrugName.IsSearchKey() && !m_szDrugNameKey.IsEmpty()){
	 szWhere.Format(_T(" and mp_product_id='%s' "), m_szDrugNameKey);
	};
	m_wndDrugName.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" SELECT") \
			_T("   DISTINCT") \
			_T("   mp_code                                   AS code,") \
			_T("   mp_name                                   AS name,") \
			_T("   Get_uomname(mp_product_uom_id)                    AS unit,") \
			_T("   Get_productclassname(mp_product_class_id) AS genericname,") \
			_T("   Get_countryname((mp_country_id))           AS orgname,") \
			_T("   mp_product_id as id") \
			_T(" FROM   m_storageline") \
			_T("        LEFT JOIN m_product ON(msl_product_id=mp_product_id)") \
			_T(" WHERE  msl_storage_id=3") \
			_T("        AND msl_product_item_id>0") \
			_T("		%s") \
			_T(" ORDER  BY name, unit "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDrugName.AddItems(
			rs.GetValue(_T("id")),
			rs.GetValue(_T("code")),
			rs.GetValue(_T("name")),
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMServiceDrugTrack::OnDrugNameAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMServiceDrugTrack::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMServiceDrugTrack::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	if (m_szDrugNameKey.IsEmpty())
	{
		m_wndDrugName.SetFocus();
		return;
	}
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szOldRoomId, szNewRoomId;
	int nCol = 0, nRow = 0, nPrevRow = 0, nIdx = 1;
	long nQty = 0, nTtlQty = 0;
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 3);
	xls.SetCellMergedColumns(nCol, nRow + 3, 3);
	xls.SetColumnWidth(1, 35);
	xls.SetCellText(nCol, nRow, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, _T("TH\x45O \x44\xD5I THU\x1ED0\x43 \x44\x1ECA\x43H V\x1EE4 NGO\x1EA0I TR\xDA"), FMT_TEXT | FMT_CENTER, true, 11);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDate::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDate::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 10);
	tmpStr.Format(_T("Thu\x1ED1\x63: %s"), m_wndDrugName.GetCurrent(2));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT, true, 10);
	xls.SetCellText(nCol, nRow + 5, _T("STT"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 1, nRow + 5, _T("\x42\xE1\x63 s\x1EF9"), FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol + 2, nRow + 5, _T("S\x1ED1 l\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER, true, 10);
	nRow = 6;
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("roomid"), szNewRoomId);
		if (szNewRoomId != szOldRoomId)
		{
			xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
			xls.SetCellText(nCol + 1, nRow, pMF->GetRoomName(_T("C1.1"), str2int(szNewRoomId)));
			if (nQty > 0)
			{
				tmpStr.Format(_T("%ld"), nQty);
				xls.SetCellText(nCol + 2, nPrevRow, tmpStr, FMT_NUMBER1, true);
				nTtlQty += nQty;
				nQty = 0;
			}
			nPrevRow = nRow;
			szOldRoomId = szNewRoomId;
			nRow++;
		}
		xls.SetCellText(nCol + 1, nRow, _T("- ")+rs.GetValue(_T("doctor")));
		rs.GetValue(_T("qtyissue"), tmpStr);
		nQty += str2long(tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}	
	if (nQty > 0)
	{
		tmpStr.Format(_T("%ld"), nQty);
		xls.SetCellText(nCol + 2, nPrevRow, tmpStr, FMT_NUMBER1, true);
		nTtlQty += nQty;
		nQty = 0;
	}
	if (nTtlQty > 0)
	{
		xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_CENTER, true);
		tmpStr.Format(_T("%ld"), nTtlQty);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_NUMBER1, true);
	}

	xls.Save(_T("Exports\\Theo doi thuoc dich vu.xls"));
} 

CString CEMServiceDrugTrack::GetQueryString(){
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND hfe_createddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szDrugNameKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hpol_product_id = %s"), m_szDrugNameKey);
	szSQL.Format(_T("SELECT get_username(hpo_doctor) as doctor, ") \
				_T("       hpo_roomid as roomid, ") \
				_T("       Sum(hpol_qtyissue) as qtyissue ") \
				_T("FROM   hms_pharmaorder ") \
				_T("       LEFT JOIN hms_pharmaorderline ON ( hpo_orderid = hpol_orderid ) ") \
				_T("	   LEFT JOIN hms_fee_invoice ON (hpo_invoiceno = hfe_invoiceno)")
				_T("WHERE  hpo_storage_id = 3 ") \
				_T("       AND hpo_deptid = 'C1.1' AND hfe_status = 'P' %s ") \
				_T("GROUP  BY hpo_doctor, ") \
				_T("          hpo_roomid ") \
				_T("HAVING SUM(hpol_qtyissue) > 0") \
				_T("ORDER  BY hpo_roomid, ") \
				_T("          hpo_doctor "), szWhere);
	return szSQL;
}