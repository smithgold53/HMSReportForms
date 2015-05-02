#include "stdafx.h"
#include "TMDepartmentUsageinDetail.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CTMDepartmentUsageinDetail *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CTMDepartmentUsageinDetail *)pWnd)->OnToDateCheckValue();
} 
static void _OnStorageSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMDepartmentUsageinDetail* )pWnd)->OnStorageSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnStorageSelendokFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnStorageSelendok();
}
/*static void _OnStorageSetfocusFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnStorageKillfocus();
}*/
/*static void _OnStorageKillfocusFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnStorageKillfocus();
}*/
static long _OnStorageLoadDataFnc(CWnd *pWnd){
	return ((CTMDepartmentUsageinDetail *)pWnd)->OnStorageLoadData();
}
/*static void _OnStorageAddNewFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnStorageAddNew();
}*/
static int _OnStockSetCheckFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail*) pWnd)->OnStockSetCheck();
	return 0;
}
static int _OnDeptSetCheckFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail*) pWnd)->OnDeptSetCheck();
	return 0;
}
static void _OnDepartmentSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	return ((CTMDepartmentUsageinDetail*)pWnd)->OnDepartmentSelectChange(nOldItemSel, nNewItemSel);
}
static long _OnDepartmentLoadDataFnc(CWnd *pWnd){
	return ((CTMDepartmentUsageinDetail *)pWnd)->OnDepartmentLoadData();
}
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CTMDepartmentUsageinDetail *)pWnd)->OnClerkLoadData();
}
static long _OnTypeLoadDataFnc(CWnd *pWnd){
	return ((CTMDepartmentUsageinDetail *)pWnd)->OnTypeLoadData();
}
static void _OnItemGroupSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CTMDepartmentUsageinDetail* )pWnd)->OnItemGroupSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnItemGroupSelendokFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnItemGroupSelendok();
}
/*static void _OnItemGroupSetfocusFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnItemGroupKillfocus();
}*/
/*static void _OnItemGroupKillfocusFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnItemGroupKillfocus();
}*/
static long _OnItemGroupLoadDataFnc(CWnd *pWnd){
	return ((CTMDepartmentUsageinDetail *)pWnd)->OnItemGroupLoadData();
}
/*static void _OnItemGroupAddNewFnc(CWnd *pWnd){
	((CTMDepartmentUsageinDetail *)pWnd)->OnItemGroupAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CTMDepartmentUsageinDetail *pVw = (CTMDepartmentUsageinDetail *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CTMDepartmentUsageinDetail *pVw = (CTMDepartmentUsageinDetail *)pWnd;
	pVw->OnExportSelect();
} 
CTMDepartmentUsageinDetail::CTMDepartmentUsageinDetail(CWnd* pParent){

	m_nDlgWidth = 545;
	m_nDlgHeight = 185;
	SetDefaultValues();
}
CTMDepartmentUsageinDetail::~CTMDepartmentUsageinDetail(){
}
void CTMDepartmentUsageinDetail::OnCreateComponents(){
	m_wndDepartmentUsageinDetail.Create(this, _T("Department Usage in Detail"), 5, 5, 575, 545);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 290, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 295, 30, 375, 55);
	m_wndToDate.Create(this,380, 30, 570, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 290, 85); 
	m_wndTypeLabel.Create(this, _T("Type"), 295, 60, 375, 85);
	m_wndType.Create(this,380, 60, 570, 85); 
	m_wndItemGroupLabel.Create(this, _T("Item Group"), 10, 90, 90, 115);
	m_wndItemGroup.Create(this,95, 90, 570, 115); 
	m_wndGroupbyDept.Create(this, _T("Group by Dept"), 5, 550, 110, 575);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 395, 550, 490, 575);
	m_wndExport.Create(this, _T("&Export"), 495, 550, 575, 575);
	m_wndDepartmentList.Create(this,10, 120, 570, 330); 
	m_wndDepartmentList.SetCheckBox(true);
	m_wndStorageList.Create(this,10, 335, 570, 540); 
	m_wndStorageList.SetCheckBox(true);
}
void CTMDepartmentUsageinDetail::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndItemGroup.LimitText(35);


	m_wndStorageList.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndStorageList.InsertColumn(1, _T("Name"), CFMT_TEXT, 450);


	m_wndDepartmentList.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndDepartmentList.InsertColumn(1, _T("Name"), CFMT_TEXT, 380);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 60);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 140);

	m_wndType.InsertColumn(0, _T("ID"), CFMT_TEXT, 60);
	m_wndType.InsertColumn(1, _T("Name"), CFMT_TEXT, 140);

	m_wndItemGroup.Format(3, 2);
	m_wndItemGroup.InsertColumn(0, _T(""), CFMT_TEXT, 0);
	m_wndItemGroup.InsertColumn(1, _T("Code"), CFMT_TEXT, 60);
	m_wndItemGroup.InsertColumn(2, _T("Name"), CFMT_TEXT, 400);

}
void CTMDepartmentUsageinDetail::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndStorageList.SetEvent(WE_SELENDOK, _OnStorageSelendokFnc);
	//m_wndStorageList.SetEvent(WE_SETFOCUS, _OnStorageSetfocusFnc);
	//m_wndStorageList.SetEvent(WE_KILLFOCUS, _OnStorageKillfocusFnc);
	m_wndStorageList.SetEvent(WE_SELCHANGE, _OnStorageSelectChangeFnc);
	m_wndStorageList.SetEvent(WE_LOADDATA, _OnStorageLoadDataFnc);
	//m_wndStorageList.SetEvent(WE_ADDNEW, _OnStorageAddNewFnc);
	m_wndStorageList.AddEvent(1, _T("Check/ Uncheck All"), _OnStockSetCheckFnc);
	m_wndDepartmentList.SetEvent(WE_SELCHANGE, _OnDepartmentSelectChangeFnc);
	m_wndDepartmentList.SetEvent(WE_LOADDATA, _OnDepartmentLoadDataFnc);
	m_wndDepartmentList.AddEvent(1, _T("Check/ Uncheck All"), _OnDeptSetCheckFnc);
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	m_wndType.SetEvent(WE_LOADDATA, _OnTypeLoadDataFnc);
	m_wndItemGroup.SetEvent(WE_SELENDOK, _OnItemGroupSelendokFnc);
	//m_wndItemGroup.SetEvent(WE_SETFOCUS, _OnItemGroupSetfocusFnc);
	//m_wndItemGroup.SetEvent(WE_KILLFOCUS, _OnItemGroupKillfocusFnc);
	m_wndItemGroup.SetEvent(WE_SELCHANGE, _OnItemGroupSelectChangeFnc);
	m_wndItemGroup.SetEvent(WE_LOADDATA, _OnItemGroupLoadDataFnc);
	//m_wndItemGroup.SetEvent(WE_ADDNEW, _OnItemGroupAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T(" 00:00");
	m_szToDate += _T(" 23:59");
	OnDepartmentLoadData();
	OnStorageLoadData();
	OnClerkLoadData();
	OnTypeLoadData();
	OnItemGroupLoadData();
	UpdateData(false);

}
void CTMDepartmentUsageinDetail::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndStorageList.GetDlgCtrlID(), m_szStorageKey);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_TextEx(pDX, m_wndType.GetDlgCtrlID(), m_szTypeKey);
	DDX_TextEx(pDX, m_wndItemGroup.GetDlgCtrlID(), m_szItemGroupKey);
	DDX_Check(pDX, m_wndGroupbyDept.GetDlgCtrlID(), m_bGroupbyDept);

}
void CTMDepartmentUsageinDetail::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szStorageKey.Empty();
	m_szItemGroupKey.Empty();
	m_bStockCheck = FALSE;
	m_bDeptCheck = FALSE;

}
int CTMDepartmentUsageinDetail::SetMode(int nMode){
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
/*void CTMDepartmentUsageinDetail::OnFromDateChange(){
	
} */
/*void CTMDepartmentUsageinDetail::OnFromDateSetfocus(){
	
} */
/*void CTMDepartmentUsageinDetail::OnFromDateKillfocus(){
	
} */
int CTMDepartmentUsageinDetail::OnFromDateCheckValue(){
	return 0;
} 
/*void CTMDepartmentUsageinDetail::OnToDateChange(){
	
} */
/*void CTMDepartmentUsageinDetail::OnToDateSetfocus(){
	
} */
/*void CTMDepartmentUsageinDetail::OnToDateKillfocus(){
	
} */
int CTMDepartmentUsageinDetail::OnToDateCheckValue(){
	return 0;
} 
void CTMDepartmentUsageinDetail::OnStorageSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CTMDepartmentUsageinDetail::OnStorageSelendok(){
	 
}
/*void CTMDepartmentUsageinDetail::OnStorageSetfocus(){
	
}*/
/*void CTMDepartmentUsageinDetail::OnStorageKillfocus(){
	
}*/
long CTMDepartmentUsageinDetail::OnStorageLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere, szStorageID, szStoragePerm;
	m_wndStorageList.BeginLoad();
	int nCount = 0, nItem = 0;
	szStoragePerm = pMF->m_szStoragePerm;
	if (!szStoragePerm.IsEmpty())
		szWhere.AppendFormat(_T(" AND msl_storage_id IN (%s)"), szStoragePerm);
	szSQL.Format(_T("SELECT msl_storage_id as id, msl_name as name FROM m_storagelist WHERE msl_isactive = 'Y' %s ORDER BY msl_storage_id "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		nItem = m_wndStorageList.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.GetValue(_T("id"), szStorageID);
		if (szStorageID == _T("101"))
			m_wndStorageList.SetCheck(nItem);
		rs.MoveNext();
	}
	m_wndStorageList.EndLoad();
	return nCount;
}
/*void CTMDepartmentUsageinDetail::OnStorageAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CTMDepartmentUsageinDetail::OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel){
	
}
long CTMDepartmentUsageinDetail::OnDepartmentLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndDepartmentList.BeginLoad(); 
	m_wndDepartmentList.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" select sd_id as id, sd_name as name from sys_dept where sd_type IN ('DT', 'KB') order by  sd_name"));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndDepartmentList.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	m_wndDepartmentList.EndLoad(); 
	return nCount;
}
long CTMDepartmentUsageinDetail::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndClerk.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T(" select su_userid as id, su_name as name from sys_user where su_deptid IN ('KD') order by su_name"));
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("ID")), 
			rs.GetValue(_T("Name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}

long CTMDepartmentUsageinDetail::OnTypeLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	int nCount = 0;
	szSQL.Format(_T(" SELECT DISTINCT product_type_id, ") \
	_T("                 product_type_name ") \
	_T(" FROM   (SELECT CASE WHEN mpt_product_type_id IN ( 'A1100', 'A1130', 'A1140', 'A1160', ") \
	_T("                                              'A1170', 'A1180', 'A1190', 'A1400', 'A6000' ) THEN N'A0000' ") \
	_T("                ELSE mpt_product_type_id ") \
	_T("                END product_type_id, ") \
	_T("                CASE WHEN mpt_product_type_id IN ( 'A1100', 'A1130', 'A1140', 'A1160', ") \
	_T("                                              'A1170', 'A1180', 'A1190', 'A1400', 'A6000' ) THEN N'Thu\x1ED1\x63' ") \
	_T("                ELSE mpt_name ") \
	_T("                END product_type_name ") \
	_T("         FROM   m_product_type ") \
	_T("         WHERE  mpt_org_id IN ('PM', 'MA') AND mpt_isactive = 'Y' AND mpt_product_type_id <> 'A1000') ") \
	_T(" ORDER BY product_type_id"));
	nCount = rs.ExecSQL(szSQL);
	m_wndType.DeleteAllItems();
	while (!rs.IsEOF())
	{
		m_wndType.AddItems(
			rs.GetValue(_T("product_type_id")),
			rs.GetValue(_T("product_type_name")),
			NULL);
		rs.MoveNext();
	}
	return nCount;
}

void CTMDepartmentUsageinDetail::OnItemGroupSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

} 
void CTMDepartmentUsageinDetail::OnItemGroupSelendok(){
	 
}
/*void CTMDepartmentUsageinDetail::OnItemGroupSetfocus(){
	
}*/
/*void CTMDepartmentUsageinDetail::OnItemGroupKillfocus(){
	
}*/
long CTMDepartmentUsageinDetail::OnItemGroupLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szFilter;
	szFilter.Format(_T(" AND mpc_level = 1"));
	return pMF->LoadProductCategoryList(&m_wndItemGroup, m_szItemGroupKey, szFilter);
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndItemGroup.IsSearchKey() && !m_szItemGroupKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szItemGroupKey
};
	m_wndItemGroup.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndItemGroup.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CTMDepartmentUsageinDetail::OnItemGroupAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CTMDepartmentUsageinDetail::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CString szSQL, tmpStr, szNewType, szOldType, szTemp, szOldDept, szNewDept;
	CRecord rs(&pMF->m_db);
	CReportSection *rptHeader, *rptDetail;
	CStringArray arrDat;
	CString szCurDte;
	int nIdx = 1, j = 0, k = 0, nSel = 0, nCount = 0;
	long double nTtCost[7], nGrpCost[7], nTemp = 0;
	long double nTtQty[7], nGrpQty[7];
	if (!rpt.Init(_T("Reports\\HMS\\PM_CHITIETTINHHINHSUDUNGTAIDONVI.RPT")))
		return;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONSTOP | MB_OK);
		return;
	}
	for (int i = 0; i < 7; i++)
	{
		nGrpQty[i] = 0;
		nGrpCost[i] = 0;
		nTtCost[i] = 0;
		nTtQty[i] = 0;
	}
	//Header Section
	rptHeader = rpt.GetReportHeader();
	rptHeader->SetValue(_T("HEALTHSERVICE"), pMF->m_szHealthService);
	rptHeader->SetValue(_T("HOSPITALNAME"), pMF->m_szHospitalName);
	tmpStr = _T("To\xE0n \x62\x1ED9");
	nSel = m_wndStorageList.GetCurSel();
	if (!m_szStorageKey.IsEmpty())
		rptHeader->SetValue(_T("StockName"), m_wndStorageList.GetItemText(nSel, 1));
	else
		rptHeader->SetValue(_T("StockName"), tmpStr);
	nSel = m_wndDepartmentList.GetCurSel();
	for (int i = 0; i < m_wndDepartmentList.GetItemCount(); i++)
		if (m_wndDepartmentList.GetCheck(i))
			nCount++;
	if (nCount == 0)
		rptHeader->SetValue(_T("Department"), _T("To\xE0n \x62\x1ED9"));
	else if (nCount == 1)
		rptHeader->SetValue(_T("Department"), m_wndDepartmentList.GetItemText(nSel, 1));
	else
		rptHeader->SetValue(_T("Department"), _T("Nhi\x1EC1u kho"));
	if (!m_szItemGroupKey.IsEmpty())
		rptHeader->SetValue(_T("Type"), m_wndItemGroup.GetCurrent(1));
	else
		rptHeader->SetValue(_T("Type"), tmpStr);
	tmpStr.Format(rptHeader->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), 
	CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rptHeader->SetValue(_T("ReportDate"), tmpStr);
	//Detail
	arrDat.Add(_T("pname"));
	arrDat.Add(_T("product_uomname"));
	arrDat.Add(_T("price"));
	arrDat.Add(_T("solqty"));
	arrDat.Add(_T("solamt"));
	arrDat.Add(_T("polqty"));
	arrDat.Add(_T("polamt"));
	arrDat.Add(_T("solinsqty"));
	arrDat.Add(_T("solinsamt"));
	arrDat.Add(_T("othinsqty"));
	arrDat.Add(_T("othinsamt"));
	arrDat.Add(_T("serqty"));
	arrDat.Add(_T("seramt"));
	arrDat.Add(_T("othqty"));
	arrDat.Add(_T("othamt"));
	arrDat.Add(_T("ttlqty"));
	arrDat.Add(_T("ttlamt"));
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("cat"), szNewType);
		rs.GetValue(_T("deptid"), szNewDept);
		if (szNewType != szOldType)
		{
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
			TranslateString(szNewType, tmpStr);
			rptDetail->SetValue(_T("GroupName"), tmpStr);
			szOldType = szNewType;
		}
		if (m_bGroupbyDept)
			if (szNewDept != szOldDept)
			{
				nTemp = 0;
				for (int i = 0; i < 7; i++)
				{
					nTemp += nGrpQty[i];
				}
				if (nTemp > 0)
				{
					rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
					for (int i = 0; i < 7; i++)
					{
						tmpStr.Format(_T("s%d"), (i+2)*2+1);
						szTemp.Format(_T("%f"), nGrpQty[i]);
						nTtQty[i] += nGrpQty[i];
						nGrpQty[i] = 0;
						rptDetail->SetValue(tmpStr, szTemp);
						tmpStr.Format(_T("s%d"), (i+3)*2);
						szTemp.Format(_T("%.2f"), nGrpCost[i]);
						nTtCost[i] += nGrpCost[i];
						nGrpCost[i] = 0;
						rptDetail->SetValue(tmpStr, szTemp);
					}
				}
				rptDetail = rpt.AddDetail(rpt.GetGroupHeader(0));
				rptDetail->SetValue(_T("GroupName"), pMF->GetDepartmentName(szNewDept));
				szOldDept = szNewDept;
				nIdx = 1;
			}
		rptDetail = rpt.AddDetail();
		rptDetail->SetValue(_T("1"), int2str(nIdx++));	
		//i: 0..16
		//j: 2..18
		//k: 0..7
		k = 0;
		for (int i = 0; i < arrDat.GetCount(); i++)
		{
			j = i + 2;
			rs.GetValue(arrDat.GetAt(i), tmpStr);
			if (tmpStr != _T("0"))
				rptDetail->SetValue(int2str(j), tmpStr);
			if ((j%2 == 1) && (j >= 5))
			{
				if (m_bGroupbyDept)
					nGrpQty[k] += str2double(tmpStr);
				else
					nTtQty[k] += str2double(tmpStr);
			}
			if ((j%2 == 0) && (j >= 6))
			{
				if (m_bGroupbyDept)
					nGrpCost[k] += str2double(tmpStr);
				else
					nTtCost[k] += str2double(tmpStr);
				k++;
			}
		}
		rs.MoveNext();
	}
	if (m_bGroupbyDept)
	{
		nTemp = 0;
		for (int i = 0; i < 7; i++)
		{
			nTemp += nGrpQty[i];
		}
		if (nTemp > 0)
		{
			rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
			for (int i = 0; i < 7; i++)
			{
				tmpStr.Format(_T("s%d"), (i+2)*2+1);
				szTemp.Format(_T("%f"), nGrpQty[i]);
				nTtQty[i] += nGrpQty[i];
				nGrpQty[i] = 0;
				rptDetail->SetValue(tmpStr, szTemp);
				tmpStr.Format(_T("s%d"), (i+3)*2);
				szTemp.Format(_T("%.2f"), nGrpCost[i]);
				nTtCost[i] += nGrpCost[i];
				nGrpCost[i] = 0;
				rptDetail->SetValue(tmpStr, szTemp);
			}
		}
	}
	rptDetail = rpt.AddDetail(rpt.GetGroupFooter(1));
	for (int i = 0; i < 7; i++)
	{
		tmpStr.Format(_T("s%d"), (i+2)*2+1);
		szTemp.Format(_T("%f"), nTtQty[i]);
		rptDetail->SetValue(tmpStr, szTemp);
		tmpStr.Format(_T("s%d"), (i+3)*2);
		szTemp.Format(_T("%.2f"), nTtCost[i]);
		rptDetail->SetValue(tmpStr, szTemp);
	}
	szCurDte = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szCurDte.Right(2), szCurDte.Mid(5, 2), szCurDte.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 

void CTMDepartmentUsageinDetail::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr, szNewType, szOldType, szNewDept, szOldDept, szTemp;
	CRecord rs(&pMF->m_db);
	int nIdx = 1, j = 0, k = 0;
	long double nTtCost[7], nGrpCost[7], nTemp = 0;
	long double nTtQty[7], nGrpQty[7];
	CStringArray arrDat;
	CellFormat hf(&xls);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING);
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"), MB_ICONASTERISK | MB_OK);
		return;
	}
	for (int i = 0; i < 7; i++)
	{
		nGrpQty[i] = 0;
		nGrpCost[i] = 0;
		nTtQty[i] = 0;
		nTtCost[i] = 0;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	//Header Section
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 30);

	int nCol = 0;
	int nRow = 0;	

	xls.SetCellMergedColumns(nCol, nRow + 1, 2);
	xls.SetCellMergedColumns(nCol, nRow + 2, 2);
	xls.SetCellText(nCol, nRow + 1, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellMergedColumns(nCol, nRow + 3, 17);
	xls.SetCellText(nCol, nRow + 3, _T("\x43HI TI\x1EBET S\x1EEC \x44\x1EE4NG THU\x1ED0\x43 T\x1EA0I \x110\x1A0N V\x1ECA"), FMT_TEXT | FMT_CENTER, true, 12);
	xls.SetCellMergedColumns(nCol, nRow + 4, 17);
	tmpStr.Format(_T("T\x1EEB %s \x110\x1EBFn %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 4, tmpStr, FMT_TEXT | FMT_CENTER, true, 11);
	nRow = 5;
	xls.SetCellMergedRows(nCol, nRow, 2);
	xls.SetCellMergedRows(nCol+1, nRow, 2);
	xls.SetCellMergedRows(nCol+2, nRow, 2);
	xls.SetCellMergedRows(nCol+3, nRow, 2);
	xls.SetCellMergedColumns(nCol+4, nRow, 2);
	xls.SetCellMergedColumns(nCol+6, nRow, 2);
	xls.SetCellMergedColumns(nCol+8, nRow, 2);
	xls.SetCellMergedColumns(nCol+10, nRow, 2);
	xls.SetCellMergedColumns(nCol+12, nRow, 2);
	xls.SetCellMergedColumns(nCol+14, nRow, 2);
	xls.SetCellMergedColumns(nCol+16, nRow, 2);

	xls.SetCellText(nCol, nRow, _T("STT"), &hf);
	xls.SetCellText(nCol+1, nRow, _T("T\xEAn thu\x1ED1\x63, h\xE0m l\x1B0\x1EE3ng"));
	xls.SetCellText(nCol+2, nRow, _T("\x110VT"));
	xls.SetCellText(nCol+3, nRow, _T("\x110\x1A1n gi\xE1"));
	xls.SetCellText(nCol+4, nRow, _T("\x42\x1ED9 \x111\x1ED9i"));
	xls.SetCellText(nCol+4, nRow+1, _T("SL"));
	xls.SetCellText(nCol+5, nRow+1, _T("Th\xE0nh ti\x1EC1n"));
	xls.SetCellText(nCol+6, nRow, _T("\x43h\xEDnh s\xE1\x63h"));
	xls.SetCellText(nCol+6, nRow+1, _T("SL"));
	xls.SetCellText(nCol+7, nRow+1, _T("Th\xE0nh ti\x1EC1n"));
	xls.SetCellText(nCol+8, nRow, _T("\x42H Qu\xE2n"));
	xls.SetCellText(nCol+8, nRow+1, _T("SL"));
	xls.SetCellText(nCol+9, nRow+1, _T("Th\xE0nh ti\x1EC1n"));
	xls.SetCellText(nCol+10, nRow, _T("\x42H Kh\xE1\x63"));
	xls.SetCellText(nCol+10, nRow+1, _T("SL"));
	xls.SetCellText(nCol+11, nRow+1, _T("Th\xE0nh ti\x1EC1n"));
	xls.SetCellText(nCol+12, nRow, _T("\x44\x1ECB\x63h v\x1EE5"));
	xls.SetCellText(nCol+12, nRow+1, _T("SL"));
	xls.SetCellText(nCol+13, nRow+1, _T("Th\xE0nh ti\x1EC1n"));
	xls.SetCellText(nCol+14, nRow, _T("\x110T kh\xE1\x63"));
	xls.SetCellText(nCol+14, nRow+1, _T("SL"));
	xls.SetCellText(nCol+15, nRow+1, _T("Th\xE0nh ti\x1EC1n"));
	xls.SetCellText(nCol+16, nRow, _T("T\x1ED5ng \x63\x1ED9ng"));
	xls.SetCellText(nCol+16, nRow+1, _T("SL"));
	xls.SetCellText(nCol+17, nRow+1, _T("Th\xE0nh ti\x1EC1n"));

	nRow = 7;
	arrDat.Add(_T("pname"));
	arrDat.Add(_T("product_uomname"));
	arrDat.Add(_T("price"));
	arrDat.Add(_T("solqty"));
	arrDat.Add(_T("solamt"));
	arrDat.Add(_T("polqty"));
	arrDat.Add(_T("polamt"));
	arrDat.Add(_T("solinsqty"));
	arrDat.Add(_T("solinsamt"));
	arrDat.Add(_T("othinsqty"));
	arrDat.Add(_T("othinsamt"));
	arrDat.Add(_T("serqty"));
	arrDat.Add(_T("seramt"));
	arrDat.Add(_T("othqty"));
	arrDat.Add(_T("othamt"));
	arrDat.Add(_T("ttlqty"));
	arrDat.Add(_T("ttlamt"));
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("cat"), szNewType);
		rs.GetValue(_T("deptid"), szNewDept);
		if (szNewType != szOldType)
		{
			xls.SetCellMergedColumns(nCol, nRow, 9);
			TranslateString(szNewType, tmpStr);
			xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT, true);
			szOldType = szNewType;
			nRow++;
		}
		if (m_bGroupbyDept)
			if (szNewDept != szOldDept)
			{
				nTemp = 0;
				for (int i = 0; i < 7; i++)
				{
					nTemp += nGrpQty[i];
				}
				if (nTemp > 0)
				{
					xls.SetCellText(3, nRow, _T("T\x1ED5ng nh\xF3m:"), FMT_TEXT | FMT_CENTER, true);
					for (int i = 0; i < 7; i++)
					{
						tmpStr.Format(_T("%d"), (i + 2)*2);
						szTemp.Format(_T("%f"), nGrpQty[i]);
						nTtQty[i] += nGrpQty[i];
						nGrpQty[i] = 0;
						xls.SetCellText(str2int(tmpStr), nRow, szTemp, FMT_NUMBER1);
						tmpStr.Format(_T("%d"), (i + 2)*2+1);
						szTemp.Format(_T("%.2f"), nGrpCost[i]);
						nTtCost[i] += nGrpCost[i];
						nGrpCost[i] = 0;
						xls.SetCellText(str2int(tmpStr), nRow, szTemp, FMT_NUMBER1);
					}
					nRow++;
				}
				xls.SetCellMergedColumns(0, nRow, 10);
				xls.SetCellText(0, nRow, pMF->GetDepartmentName(szNewDept), FMT_TEXT, true);
				szOldDept = szNewDept;
				nIdx = 1;
				nRow++;
			}
		xls.SetCellText(nCol, nRow, int2str(nIdx++), FMT_INTEGER);
		xls.SetCellText(nCol+1, nRow, rs.GetValue(_T("pname")), FMT_TEXT);
		xls.SetCellText(nCol+2, nRow, rs.GetValue(_T("product_uomname")), FMT_TEXT);
		k = 0;
		for (int i = 2; i < arrDat.GetCount(); i++)
		{
			j = i + 1;
			xls.SetCellText(nCol+j, nRow, rs.GetValue(arrDat.GetAt(i)), FMT_NUMBER1);
			if ((j%2 == 0) && (j >=4))
			{
				if (m_bGroupbyDept)
					nGrpQty[k] += str2double(rs.GetValue(arrDat.GetAt(i)));
				else
					nTtQty[k] += str2double(rs.GetValue(arrDat.GetAt(i)));
			}
			if ((j%2 == 1) && (j >=5))
			{
				if (m_bGroupbyDept)
					nGrpCost[k] += str2double(rs.GetValue(arrDat.GetAt(i)));
				else
					nTtCost[k] += str2double(rs.GetValue(arrDat.GetAt(i)));
				k++;
			}
		}
		nRow++;
		rs.MoveNext();
	}
	if (m_bGroupbyDept)
	{
		nTemp = 0;
		for (int i = 0; i < 7; i++)
		{
			nTemp += nGrpQty[i];
		}
		if (nTemp > 0)
		{
			xls.SetCellText(3, nRow, _T("T\x1ED5ng nh\xF3m:"), FMT_TEXT | FMT_CENTER, true);
			for (int i = 0; i < 7; i++)
			{
				tmpStr.Format(_T("%d"), (i + 2)*2);
				szTemp.Format(_T("%f"), nGrpQty[i]);
				nTtQty[i] += nGrpQty[i];
				nGrpQty[i] = 0;
				xls.SetCellText(str2int(tmpStr), nRow, szTemp, FMT_NUMBER1);
				tmpStr.Format(_T("%d"), (i + 2)*2+1);
				szTemp.Format(_T("%.2f"), nGrpCost[i]);
				nTtCost[i] += nGrpCost[i];
				nGrpCost[i] = 0;
				xls.SetCellText(str2int(tmpStr), nRow, szTemp, FMT_NUMBER1);
			}
			nRow++;
		}
	}
	xls.SetCellText(nCol + 1, nRow, _T("T\x1ED5ng \x63\x1ED9ng"), FMT_TEXT | FMT_RIGHT, true);
	for (int i = 0; i < 7; i ++)
	{
		tmpStr.Format(_T("%f"), nTtQty[i]);
		xls.SetCellText((i + 2)*2, nRow, tmpStr, FMT_NUMBER1, true);
		tmpStr.Format(_T("%.2f"), nTtCost[i]);
		xls.SetCellText((i + 2)*2+1, nRow, tmpStr, FMT_NUMBER1, true);
	}
	EndWaitCursor();
	xls.Save(_T("Exports\\CHITIETSUDUNGTHUOCTAIDONVI.xls"));
} 

CString CTMDepartmentUsageinDetail::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CString szSQL, szWhere, szWhere1, szWhere2, szWhere3, szTemp, szDept, szStorageStr, szField, szGroupBy, szOrderBy;
	for (int i = 0; i < m_wndDepartmentList.GetItemCount(); i++)
	{
		if (m_wndDepartmentList.GetCheck(i))
		{
			m_wndDepartmentList.SetCurSel(i);
			szTemp = m_wndDepartmentList.GetItemText(i, 0);
			if (!szDept.IsEmpty())
				szDept += _T(", ");
			szDept.AppendFormat(_T("'%s'"), szTemp);
			
		}
	}
	for (int i = 0; i < m_wndStorageList.GetItemCount();i++)
		if (m_wndStorageList.GetCheck(i))
		{
			m_wndStorageList.SetCurSel(i);
			if (!szStorageStr.IsEmpty())
				szStorageStr += _T(", ");
			szStorageStr.AppendFormat(_T("'%s'"), m_wndStorageList.GetItemText(i, 0));
		}

	szWhere1.Format(_T(" AND mt_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szWhere2.Format(_T(" AND hpo_approvedate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	szWhere3.Format(_T(" AND so_approveddate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!szStorageStr.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" AND mt_storage_id IN (%s)"), szStorageStr);
		szWhere2.AppendFormat(_T(" AND hpo_storage_id IN (%s)"), szStorageStr);
		szWhere3.AppendFormat(_T(" AND so_storage_id IN (%s)"), szStorageStr);
	}
	if (!szDept.IsEmpty())
	{
		szWhere2.AppendFormat(_T(" AND hpo_deptid IN (%s)"), szDept);
		szWhere1.AppendFormat(_T(" AND mt_department_to_id IN (%s)"), szDept);
		szWhere3.AppendFormat(_T(" AND hd_admitdept IN (%s)"), szDept);
	}
	if (!m_szClerkKey.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" AND mt_approvedby = '%s'"), m_szClerkKey);
		szWhere2.AppendFormat(_T(" AND hpo_approveby = '%s'"), m_szClerkKey);
		szWhere3.AppendFormat(_T(" AND so_approvedby = '%s'"), m_szClerkKey);
	}
	if (!m_szItemGroupKey.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" AND product_rootid = '%s'"), m_szItemGroupKey);
		szWhere2.AppendFormat(_T(" AND product_rootid = '%s'"), m_szItemGroupKey);
		szWhere3.AppendFormat(_T(" AND product_rootid = '%s'"), m_szItemGroupKey);
	}
	if (!m_szTypeKey.IsEmpty())
	{
		szWhere1.AppendFormat(_T(" AND product_groupid = '%s'"), m_szTypeKey);
		szWhere2.AppendFormat(_T(" AND product_groupid = '%s'"), m_szTypeKey);
		szWhere3.AppendFormat(_T(" AND product_groupid = '%s'"), m_szTypeKey);
	}
	if (pMF->GetModuleID() == _T("PM"))
		szWhere.AppendFormat(_T(" AND product_org_id = '%s'"), pMF->GetModuleID());
	//mtl_saleprice fixed price
	if (m_bGroupbyDept)
	{
		szField.Format(_T("deptid,"));
		szGroupBy.Format(_T("deptid,"));
		szOrderBy.Format(_T("deptid,"));
	}
	szSQL.Format(_T(" SELECT cat,") \
				_T("   product_id,") \
				_T("   product_name                     AS pname,") \
				_T("   price,") \
				_T("   %s") \
				_T("   product_uomname,") \
				_T("   Sum(solqty)                      AS solqty,") \
				_T("   Sum(solqty)*price    AS solamt,") \
				_T("   Sum(polqty)                      AS polqty,") \
				_T("   Sum(polqty)*price    AS polamt,") \
				_T("   Sum(solinsqty)                   AS solinsqty,") \
				_T("   Sum(solinsqty)*price AS solinsamt,") \
				_T("   Sum(othinsqty)                   AS othinsqty,") \
				_T("   Sum(othinsqty)*price AS othinsamt,") \
				_T("   Sum(serqty)                      AS serqty,") \
				_T("   Sum(serqty)*price    AS seramt,") \
				_T("   Sum(othqty)                      AS othqty,") \
				_T("   Sum(othqty)*price    AS othamt,") \
				_T("   Sum(solqty+polqty+solinsqty+othinsqty+serqty") \
				_T("       +othqty)                     AS ttlqty,") \
				_T("   Sum(solqty+polqty+solinsqty+othinsqty+serqty") \
				_T("       +othqty)*price   AS ttlamt") \
				_T(" FROM   (SELECT") \
				_T("   'Delivery' as cat,") \
				_T("   hpol_product_id,") \
				_T("   price,") \
				_T("   deptid,") \
				_T("   Sum(solqty)    AS solqty,") \
				_T("   Sum(polqty)    AS polqty,") \
				_T("   Sum(solinsqty) AS solinsqty,") \
				_T("   Sum(othinsqty) AS othinsqty,") \
				_T("   Sum(serqty)    AS serqty,") \
				_T("   Sum(otherqty)  AS othqty") \
				_T(" FROM   (SELECT") \
				_T("           hpol_product_id,") \
				_T("		   hpol_unitprice AS price,") \
				_T("           hd_object,") \
				_T("		   hpo_deptid deptid,") \
				_T("           CASE WHEN hd_object=1 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS solqty,") \
				_T("           CASE WHEN hd_object=3 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS polqty,") \
				_T("           CASE WHEN hd_object=2 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS solinsqty,") \
				_T("           CASE WHEN hd_object=4 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS othinsqty,") \
				_T("           CASE WHEN hd_object=7 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS serqty,") \
				_T("           CASE WHEN hd_object IN (5, 6, 8, 9,") \
				_T("                              10, 11, 12) THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS otherqty") \
				_T("         FROM   hms_ipharmaorder") \
				_T("         LEFT JOIN hms_doc ON (hd_docno=hpo_docno)") \
				_T("         LEFT JOIN hms_ipharmaorderline ON (hpo_orderid=hpol_orderid)") \
				_T("		 LEFT JOIN m_transaction ON (mt_transaction_id = hpo_transaction_id)") \
				_T("		 LEFT JOIN m_productitem_view ON (product_item_id = hpol_product_item_id)") \
				_T("         WHERE  hpo_orderstatus = 'A' AND hpo_ordertype IN ('D', 'M', 'B') %s) tbl1") \
				_T(" GROUP  BY hpol_product_id, price, deptid") \
				_T(" UNION ALL") \
				_T(" SELECT") \
				_T("   'Delivery' as cat,") \
				_T("   hpol_product_id,") \
				_T("   price,") \
				_T("   deptid,") \
				_T("   Sum(solqty)    AS solqty,") \
				_T("   Sum(polqty)    AS polqty,") \
				_T("   Sum(solinsqty) AS solinsqty,") \
				_T("   Sum(othinsqty) AS othinsqty,") \
				_T("   Sum(serqty)    AS serqty,") \
				_T("   Sum(otherqty)  AS othqty") \
				_T(" FROM   (SELECT") \
				_T("           hpol_product_id,") \
				_T("		   hpol_unitprice AS price,") \
				_T("           hd_object,") \
				_T("		   hpo_deptid deptid,") \
				_T("           CASE WHEN hd_object=1 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS solqty,") \
				_T("           CASE WHEN hd_object=3 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS polqty,") \
				_T("           CASE WHEN hd_object=2 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS solinsqty,") \
				_T("           CASE WHEN hd_object=4 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS othinsqty,") \
				_T("           CASE WHEN hd_object=7 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS serqty,") \
				_T("           CASE WHEN hd_object IN (5, 6, 8, 9,") \
				_T("                              10, 11, 12) THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS otherqty") \
				_T("         FROM   hms_ipharmaorder") \
				_T("         LEFT JOIN hms_doc ON (hd_docno=hpo_docno)") \
				_T("         LEFT JOIN hms_ipharmaorderline ON (hpo_orderid=hpol_orderid)") \
				_T("		 LEFT JOIN m_transaction ON (mt_transaction_id = hpo_transaction_id)") \
				_T("		 LEFT JOIN m_productitem_view ON (product_item_id = hpol_product_item_id)") \
				_T("         WHERE  hpo_orderstatus = 'A' AND hpo_ordertype = 'C' %s) tbl1") \
				_T(" GROUP  BY hpol_product_id, price, deptid") \
				_T("		 UNION ALL") \
				_T(" SELECT") \
				_T("   'Delivery' as cat,") \
				_T("   hpol_product_id,") \
				_T("   price,") \
				_T("   deptid,") \
				_T("   Sum(solqty)    AS solqty,") \
				_T("   Sum(polqty)    AS polqty,") \
				_T("   Sum(solinsqty) AS solinsqty,") \
				_T("   Sum(othinsqty) AS othinsqty,") \
				_T("   Sum(serqty)    AS serqty,") \
				_T("   Sum(otherqty)  AS otherqty") \
				_T(" FROM   (SELECT") \
				_T("           hpol_product_id,") \
				_T("		   hpol_unitprice AS price,") \
				_T("           hd_object,") \
				_T("		   hpo_deptid deptid,") \
				_T("           CASE WHEN hd_object=1 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS solqty,") \
				_T("           CASE WHEN hd_object=3 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS polqty,") \
				_T("           CASE WHEN hd_object=2 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS solinsqty,") \
				_T("           CASE WHEN hd_object=4 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS othinsqty,") \
				_T("           CASE WHEN hd_object=7 THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS serqty,") \
				_T("           CASE WHEN hd_object IN (5, 6, 8, 9,") \
				_T("                              10, 11, 12) THEN hpol_qtyorder") \
				_T("           ELSE 0") \
				_T("           END AS otherqty") \
				_T("         FROM   hms_pharmaorder") \
				_T("         LEFT JOIN hms_doc ON (hd_docno=hpo_docno)") \
				_T("         LEFT JOIN hms_pharmaorderline ON (hpo_orderid=hpol_orderid)") \
				_T("		 LEFT JOIN m_productitem_view ON (product_item_id = hpol_product_item_id)") \
				_T("         WHERE  hpo_orderstatus = 'A' %s) tbl") \
				_T(" GROUP  BY hpol_product_id, price, deptid ") \
				_T(" UNION ALL") \
				_T(" SELECT") \
				_T("   'Delivery' as cat,") \
				_T("   mtl_product_id,") \
				_T("   mtl_saleprice,") \
				_T("   mt_department_to_id deptid,") \
				_T("   Sum(mtl_qtysold)    AS solqty,") \
				_T("   Sum(mtl_qtypolicy)    AS polqty,") \
				_T("   Sum(mtl_qtysoldins) AS solinsqty,") \
				_T("   Sum(mtl_qtyotherins) AS othinsqty,") \
				_T("   Sum(mtl_qtyservice)    AS serqty,") \
				_T("   Sum(mtl_qtyother)  AS otherqty") \
				_T(" FROM m_transaction") \
				_T(" LEFT JOIN m_transactionline ON (mt_transaction_id = mtl_transaction_id)") \
				_T(" LEFT JOIN m_productitem_view ON (product_item_id = mtl_product_item_id)") \
				_T(" WHERE mt_doctype = 'DDO' AND mt_status = 'A' %s") \
				_T(" GROUP BY mtl_product_id, mtl_saleprice, mt_department_to_id") \
				_T(" UNION ALL") \
				_T(" SELECT") \
				_T("   'Delivery' as cat,") \
				_T("   sol_product_id,") \
				_T("   sol_unitprice,") \
				_T("   hd_admitdept deptid,") \
				_T("   0 AS solqty,") \
				_T("   0 AS polqty,") \
				_T("   0 solinsqty,") \
				_T("   0 othinsqty,") \
				_T("   Sum(sol_qtyorder)    AS serqty,") \
				_T("   0 otherqty") \
				_T(" FROM sale_order") \
				_T(" LEFT JOIN hms_doc ON (hd_docno = so_docno)") \
				_T(" LEFT JOIN sale_orderline ON (so_sale_order_id = sol_sale_order_id)") \
				_T(" LEFT JOIN m_productitem_view ON (product_item_id = sol_product_item_id)") \
				_T(" WHERE so_doctype = 'SOO' AND so_status = 'A' AND sol_product_id > 0 %s") \
				_T(" GROUP BY sol_product_id, sol_unitprice, hd_admitdept)") \
				_T(" LEFT JOIN m_product_view ON (product_id=hpol_product_id)") \
				_T(" WHERE 1=1 %s") \
				_T(" GROUP  BY cat,") \
				_T("		   product_id,") \
				_T("           product_name,") \
				_T("           product_saleprice2,") \
				_T("		   product_uomname, %s price") \
				_T(" ORDER  BY cat, %s product_name "), szField, szWhere1, szWhere2, szWhere2, szWhere1, szWhere3, szWhere, szGroupBy, szOrderBy);
_fmsg(_T("%s"), szSQL);
	return szSQL;
	//Tra lai
				//_T("		 UNION ALL") \
				//_T("		 SELECT ") \
				//_T("		   'Return' as cat,") \
				//_T("           mtl_product_id,") \
				//_T("           nvl(mtl_qtysold, 0)     AS solqty,") \
				//_T("           nvl(mtl_qtypolicy, 0)   AS polqty,") \
				//_T("           nvl(mtl_qtysoldins, 0)  AS solinsqty,") \
				//_T("           nvl(mtl_qtyotherins, 0) AS othinsqty,") \
				//_T("           nvl(mtl_qtyservice, 0)  AS serqty,") \
				//_T("           nvl(mtl_qtyother, 0)    AS othqty") \
				//_T("         FROM   m_transaction") \
				//_T("         LEFT JOIN m_transactionline ON (mt_transaction_id=mtl_transaction_id)") \
				//_T("         WHERE  mtl_qtyorder>0") \
				//_T("                AND mt_status='A'") \
				//_T("                AND mt_doctype IN ('DRO') %s)")
}BEGIN_MESSAGE_MAP(CTMDepartmentUsageinDetail, CGuiView)
ON_WM_SETFOCUS()
END_MESSAGE_MAP()

void CTMDepartmentUsageinDetail::OnSetFocus(CWnd* pOldWnd)
{
	CGuiView::OnSetFocus(pOldWnd);

	// TODO: Add your message handler code here
	m_wndFromDate.SetFocus();
}

void CTMDepartmentUsageinDetail::OnStockSetCheck(){
	m_bStockCheck = !m_bStockCheck;
	for (int i = 0; i < m_wndStorageList.GetItemCount();i++)
		m_wndStorageList.SetCheck(i, m_bStockCheck);
}

void CTMDepartmentUsageinDetail::OnDeptSetCheck(){
	m_bDeptCheck = !m_bDeptCheck;
	for (int i = 0; i < m_wndDepartmentList.GetItemCount();i++)
		m_wndDepartmentList.SetCheck(i, m_bDeptCheck);
}