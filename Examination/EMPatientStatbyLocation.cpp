#include "stdafx.h"
#include "EMPatientStatbyLocation.h"
#include "MainFrame_E10.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMPatientStatbyLocation *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMPatientStatbyLocation *)pWnd)->OnToDateCheckValue();
} 
static void _OnProvinceSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMPatientStatbyLocation* )pWnd)->OnProvinceSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnProvinceSelendokFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnProvinceSelendok();
}
/*static void _OnProvinceSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnProvinceKillfocus();
}*/
/*static void _OnProvinceKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnProvinceKillfocus();
}*/
static long _OnProvinceLoadDataFnc(CWnd *pWnd){
	return ((CEMPatientStatbyLocation *)pWnd)->OnProvinceLoadData();
}
/*static void _OnProvinceAddNewFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnProvinceAddNew();
}*/
static void _OnDistrictSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMPatientStatbyLocation* )pWnd)->OnDistrictSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnDistrictSelendokFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnDistrictSelendok();
}
/*static void _OnDistrictSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnDistrictKillfocus();
}*/
/*static void _OnDistrictKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnDistrictKillfocus();
}*/
static long _OnDistrictLoadDataFnc(CWnd *pWnd){
	return ((CEMPatientStatbyLocation *)pWnd)->OnDistrictLoadData();
}
/*static void _OnDistrictAddNewFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnDistrictAddNew();
}*/
static void _OnVillageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMPatientStatbyLocation* )pWnd)->OnVillageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnVillageSelendokFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnVillageSelendok();
}
/*static void _OnVillageSetfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnVillageKillfocus();
}*/
/*static void _OnVillageKillfocusFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnVillageKillfocus();
}*/
static long _OnVillageLoadDataFnc(CWnd *pWnd){
	return ((CEMPatientStatbyLocation *)pWnd)->OnVillageLoadData();
}
/*static void _OnVillageAddNewFnc(CWnd *pWnd){
	((CEMPatientStatbyLocation *)pWnd)->OnVillageAddNew();
}*/
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMPatientStatbyLocation *pVw = (CEMPatientStatbyLocation *)pWnd;
	pVw->OnExportSelect();
} 
CEMPatientStatbyLocation::CEMPatientStatbyLocation(CWnd *pParent){

	m_nDlgWidth = 500;
	m_nDlgHeight = 190;
	SetDefaultValues();
}
CEMPatientStatbyLocation::~CEMPatientStatbyLocation(){
}
void CEMPatientStatbyLocation::OnCreateComponents(){
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 490, 150);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndProvinceLabel.Create(this, _T("Province"), 10, 60, 90, 85);
	m_wndProvince.Create(this,95, 60, 485, 85); 
	m_wndDistrictLabel.Create(this, _T("District"), 10, 90, 90, 115);
	m_wndDistrict.Create(this,95, 90, 485, 115); 
	m_wndVillageLabel.Create(this, _T("Village"), 10, 120, 90, 145);
	m_wndVillage.Create(this,95, 120, 485, 145); 
	m_wndExport.Create(this, _T("&Export"), 410, 155, 490, 180);

}
void CEMPatientStatbyLocation::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndProvince.SetCheckValue(true);
	m_wndProvince.LimitText(35);
	m_wndDistrict.SetCheckValue(true);
	m_wndDistrict.LimitText(35);
	m_wndVillage.SetCheckValue(true);
	m_wndVillage.LimitText(35);


	m_wndProvince.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndProvince.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);


	m_wndDistrict.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndDistrict.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);


	m_wndVillage.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndVillage.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);

}
void CEMPatientStatbyLocation::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndProvince.SetEvent(WE_SELENDOK, _OnProvinceSelendokFnc);
	//m_wndProvince.SetEvent(WE_SETFOCUS, _OnProvinceSetfocusFnc);
	//m_wndProvince.SetEvent(WE_KILLFOCUS, _OnProvinceKillfocusFnc);
	m_wndProvince.SetEvent(WE_SELCHANGE, _OnProvinceSelectChangeFnc);
	m_wndProvince.SetEvent(WE_LOADDATA, _OnProvinceLoadDataFnc);
	//m_wndProvince.SetEvent(WE_ADDNEW, _OnProvinceAddNewFnc);
	m_wndDistrict.SetEvent(WE_SELENDOK, _OnDistrictSelendokFnc);
	//m_wndDistrict.SetEvent(WE_SETFOCUS, _OnDistrictSetfocusFnc);
	//m_wndDistrict.SetEvent(WE_KILLFOCUS, _OnDistrictKillfocusFnc);
	m_wndDistrict.SetEvent(WE_SELCHANGE, _OnDistrictSelectChangeFnc);
	m_wndDistrict.SetEvent(WE_LOADDATA, _OnDistrictLoadDataFnc);
	//m_wndDistrict.SetEvent(WE_ADDNEW, _OnDistrictAddNewFnc);
	m_wndVillage.SetEvent(WE_SELENDOK, _OnVillageSelendokFnc);
	//m_wndVillage.SetEvent(WE_SETFOCUS, _OnVillageSetfocusFnc);
	//m_wndVillage.SetEvent(WE_KILLFOCUS, _OnVillageKillfocusFnc);
	m_wndVillage.SetEvent(WE_SELCHANGE, _OnVillageSelectChangeFnc);
	m_wndVillage.SetEvent(WE_LOADDATA, _OnVillageLoadDataFnc);
	//m_wndVillage.SetEvent(WE_ADDNEW, _OnVillageAddNewFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CEMPatientStatbyLocation::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndProvince.GetDlgCtrlID(), m_szProvinceKey);
	DDX_TextEx(pDX, m_wndDistrict.GetDlgCtrlID(), m_szDistrictKey);
	DDX_TextEx(pDX, m_wndVillage.GetDlgCtrlID(), m_szVillageKey);

}
void CEMPatientStatbyLocation::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szProvinceKey.Empty();
	m_szDistrictKey.Empty();
	m_szVillageKey.Empty();

}
int CEMPatientStatbyLocation::SetMode(int nMode){
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
/*void CEMPatientStatbyLocation::OnFromDateChange(){
	
} */
/*void CEMPatientStatbyLocation::OnFromDateSetfocus(){
	
} */
/*void CEMPatientStatbyLocation::OnFromDateKillfocus(){
	
} */
int CEMPatientStatbyLocation::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMPatientStatbyLocation::OnToDateChange(){
	
} */
/*void CEMPatientStatbyLocation::OnToDateSetfocus(){
	
} */
/*void CEMPatientStatbyLocation::OnToDateKillfocus(){
	
} */
int CEMPatientStatbyLocation::OnToDateCheckValue(){
	return 0;
} 
void CEMPatientStatbyLocation::OnProvinceSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
} 
void CEMPatientStatbyLocation::OnProvinceSelendok(){
	m_wndDistrict.ResetContent();
	m_wndVillage.ResetContent();
}
/*void CEMPatientStatbyLocation::OnProvinceSetfocus(){
	
}*/
/*void CEMPatientStatbyLocation::OnProvinceKillfocus(){
	
}*/
long CEMPatientStatbyLocation::OnProvinceLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndProvince.IsSearchKey() && !m_szProvinceKey.IsEmpty()){
		szWhere.Format(_T(" and sp_id='%s' "), m_szProvinceKey);
	};
	m_wndProvince.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT sp_id as id, sp_name as name FROM sys_prov WHERE 1=1 %s ORDER BY id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndProvince.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMPatientStatbyLocation::OnProvinceAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CEMPatientStatbyLocation::OnDistrictSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CEMPatientStatbyLocation::OnDistrictSelendok(){
	m_wndVillage.ResetContent();
}
/*void CEMPatientStatbyLocation::OnDistrictSetfocus(){
	
}*/
/*void CEMPatientStatbyLocation::OnDistrictKillfocus(){
	
}*/
long CEMPatientStatbyLocation::OnDistrictLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndDistrict.IsSearchKey() && !m_szDistrictKey.IsEmpty()){
		szWhere.Format(_T(" and sd_id='%s' "), m_szDistrictKey);
	}
	if (!m_szProvinceKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND sd_provid = '%s'"), m_szProvinceKey);
	m_wndDistrict.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT sd_id as id, sd_name as name FROM sys_dist WHERE 1=1 %s ORDER BY id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDistrict.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMPatientStatbyLocation::OnDistrictAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CEMPatientStatbyLocation::OnVillageSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CEMPatientStatbyLocation::OnVillageSelendok(){
	 
}
/*void CEMPatientStatbyLocation::OnVillageSetfocus(){
	
}*/
/*void CEMPatientStatbyLocation::OnVillageKillfocus(){
	
}*/
long CEMPatientStatbyLocation::OnVillageLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndVillage.IsSearchKey() && !m_szVillageKey.IsEmpty()){
		szWhere.Format(_T(" and sv_id='%s' "), m_szVillageKey);
	}
	if (!m_szProvinceKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND sv_provid = '%s'"), m_szProvinceKey);
	if (!m_szDistrictKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND sv_distid = '%s'"), m_szDistrictKey);
	m_wndVillage.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT sv_id as id, sv_name as name FROM sys_vill WHERE 1=1 %s ORDER BY id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndVillage.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMPatientStatbyLocation::OnVillageAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */
void CEMPatientStatbyLocation::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0, nIdx = 1;
	double nCount = 0, nTotal = 0;
	
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data."), MB_OK|MB_ICONSTOP);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 30);
	xls.SetColumnWidth(2, 10);
	xls.SetCellMergedColumns(0, 0, 2);
	xls.SetCellMergedColumns(0, 1, 2);
	xls.SetCellMergedColumns(0, 2, 3);
	xls.SetCellMergedColumns(0, 3, 3);
	xls.SetCellText(0, 0, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, _T("TH\x1ED0NG K\xCA \x42\x1EC6NH NH\xC2N T\x1EEANG \x110\x1ECA\x41 PH\x1AF\x1A0NG"), FMT_TEXT | FMT_CENTER, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), m_szFromDate, m_szToDate);
	xls.SetCellText(0, 3, tmpStr, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(1, 4, _T("T\x1ED5ng s\x1ED1:"), FMT_TEXT, true);
	xls.SetCellText(0, 5, _T("STT"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(1, 5, _T("T\xEAn"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(2, 5, _T("S\x1ED1 l\x1B0\x1EE3ng"), FMT_TEXT | FMT_CENTER, true);
	nRow = 6;
	while (!rs.IsEOF())
	{
		xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_TEXT | FMT_RIGHT);
		xls.SetCellText(nCol + 1, nRow, rs.GetValue(_T("locationame")), FMT_TEXT);
		rs.GetValue(_T("cnt"), nCount);
		nTotal += nCount;
		xls.SetCellText(nCol + 2, nRow, double2str(nCount), FMT_NUMBER1);
		nRow++;
		rs.MoveNext();
	}
	xls.SetCellText(nCol + 2, 4, double2str(nTotal), FMT_NUMBER1);
	xls.Save(_T("Exports\\Thong Ke BN theo que quan.xls"));
} 

CString CEMPatientStatbyLocation::GetQueryString(){
	CString szSQL, szCol, szWhere, szGroupBy;
	szWhere.Format(_T(" AND hd_admitdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szVillageKey.IsEmpty())
	{
		szCol.Format(_T(" (select sv_name from sys_vill where sv_id = %s) locationame"), m_szVillageKey);
		szWhere.AppendFormat(_T(" AND hp_villid = %s AND coalesce(hp_villid, 0) <> 0"), m_szVillageKey);
		szGroupBy.Format(_T(""));
	}
	else if (!m_szDistrictKey.IsEmpty())
	{
		szCol.Format(_T(" (select sv_name from sys_vill where sv_id = hp_villid) locationame"));
		szWhere.AppendFormat(_T(" AND hp_distid = %s AND coalesce(hp_distid, 0) <> 0"), m_szDistrictKey);
		szGroupBy.Format(_T(" GROUP BY hp_villid"));
	}
	else if (!m_szProvinceKey.IsEmpty())
	{
		szCol.Format(_T(" (select sd_name from sys_dist where sd_id = hp_distid) locationame"));
		szWhere.AppendFormat(_T(" AND hp_provid = %s AND coalesce(hp_provid, 0) <> 0"), m_szProvinceKey);
		szGroupBy.Format(_T(" GROUP BY hp_distid"));
	}
	else
	{
		szCol.Format(_T(" (select sp_name from sys_prov where sp_id = hp_provid) locationame"));
		szWhere.AppendFormat(_T(" AND coalesce(hp_provid, 0) <> 0"));
		szGroupBy.Format(_T(" GROUP BY hp_provid"));
	}
	szSQL.Format(_T(" SELECT    %s, Count(DISTINCT hp_patientno) cnt") \
				_T(" FROM      hms_patient ") \
				_T(" LEFT JOIN hms_doc ON ( hp_patientno = hd_patientno )") \
				_T(" WHERE hd_status NOT IN ('O','C') %s %s ORDER BY locationame"), szCol, szWhere, szGroupBy);

	return szSQL;
}