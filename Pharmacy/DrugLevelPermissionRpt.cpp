#include "stdafx.h"
#include "DrugLevelPermissionRpt.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include ".\druglevelpermissionrpt.h"

static void _OnWardSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CDrugLevelPermissionRpt* )pWnd)->OnWardSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnWardSelendokFnc(CWnd *pWnd){
	((CDrugLevelPermissionRpt *)pWnd)->OnWardSelendok();
}
/*static void _OnWardSetfocusFnc(CWnd *pWnd){
	((CDrugLevelPermissionRpt *)pWnd)->OnWardKillfocus();
}*/
/*static void _OnWardKillfocusFnc(CWnd *pWnd){
	((CDrugLevelPermissionRpt *)pWnd)->OnWardKillfocus();
}*/
static long _OnWardLoadDataFnc(CWnd *pWnd){
	return ((CDrugLevelPermissionRpt *)pWnd)->OnWardLoadData();
}
/*static void _OnWardAddNewFnc(CWnd *pWnd){
	((CDrugLevelPermissionRpt *)pWnd)->OnWardAddNew();
}*/
static void _OnExportSelectFnc(CWnd *pWnd){
	CDrugLevelPermissionRpt *pVw = (CDrugLevelPermissionRpt *)pWnd;
	pVw->OnExportSelect();
} 
CDrugLevelPermissionRpt::CDrugLevelPermissionRpt(CWnd *pParent){
	m_nDlgWidth = 310;
	m_nDlgHeight = 100;
	SetDefaultValues();
}
CDrugLevelPermissionRpt::~CDrugLevelPermissionRpt(){
}
void CDrugLevelPermissionRpt::OnCreateComponents(){
	m_wndDrugLevelPermission.Create(this, _T("Drug Level Permission"), 5, 5, 300, 60);
	m_wndWardLabel.Create(this, _T("Ward"), 10, 30, 90, 55);
	m_wndWard.Create(this,95, 30, 295, 55); 
	m_wndExport.Create(this, _T("&Export"), 220, 65, 300, 90);

}
void CDrugLevelPermissionRpt::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndWard.SetCheckValue(true);
	m_wndWard.LimitText(35);


	m_wndWard.InsertColumn(0, _T("ID"), CFMT_TEXT, 50);
	m_wndWard.InsertColumn(1, _T("Name"), CFMT_TEXT, 150);

}
void CDrugLevelPermissionRpt::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndWard.SetEvent(WE_SELENDOK, _OnWardSelendokFnc);
	//m_wndWard.SetEvent(WE_SETFOCUS, _OnWardSetfocusFnc);
	//m_wndWard.SetEvent(WE_KILLFOCUS, _OnWardKillfocusFnc);
	m_wndWard.SetEvent(WE_SELCHANGE, _OnWardSelectChangeFnc);
	m_wndWard.SetEvent(WE_LOADDATA, _OnWardLoadDataFnc);
	//m_wndWard.SetEvent(WE_ADDNEW, _OnWardAddNewFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	SetMode(VM_NONE);

}
void CDrugLevelPermissionRpt::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndWard.GetDlgCtrlID(), m_szWardKey);

}
void CDrugLevelPermissionRpt::SetDefaultValues(){

	m_szWardKey.Empty();

}
int CDrugLevelPermissionRpt::SetMode(int nMode){
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
void CDrugLevelPermissionRpt::OnWardSelectChange(int nOldItemSel, int nNewItemSel){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
void CDrugLevelPermissionRpt::OnWardSelendok(){
	 
}
/*void CDrugLevelPermissionRpt::OnWardSetfocus(){
	
}*/
/*void CDrugLevelPermissionRpt::OnWardKillfocus(){
	
}*/
long CDrugLevelPermissionRpt::OnWardLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
/*
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndWard.IsSearchKey() && !m_szWardKey.IsEmpty()){
	 szWhere.Format(_T(" and id='%s' "), m_szWardKey
};
	m_wndWard.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT fld1 as id, fld2 as name FROM tbl WHERE 1=1 %s ORDER BY id "), szWhere
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndWard.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
*/
	return 0;
}
/*void CDrugLevelPermissionRpt::OnWardAddNew(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} */

void CDrugLevelPermissionRpt::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL;
	CString tmpStr, szTmp, szOldLine, szNewLine;
	int nRow = 0, nCol = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetCellText(nCol, nRow, _T("M\xE3 NS\x44"), FMT_TEXT, true);
	xls.SetCellText(nCol + 1, nRow, _T("M\x1EE9\x63 \x31"), FMT_TEXT, true);
	xls.SetCellText(nCol + 2, nRow, _T("M\x1EE9\x63 \x32"), FMT_TEXT, true);
	xls.SetCellText(nCol + 3, nRow, _T("M\x1EE9\x63 \x33"), FMT_TEXT, true);
	xls.SetCellText(nCol + 4, nRow, _T("Kho\x61"), FMT_TEXT, true);
	xls.SetCellText(nCol + 5, nRow, _T("T\xEAn"), FMT_TEXT, true);
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("userid"), szNewLine);
		if (szOldLine != szNewLine)
		{
			szOldLine = szNewLine;
			nRow++;
		}
		xls.SetCellText(nCol, nRow, szOldLine); 
		for (int i = 1; i < 4; i++)
		{
			szTmp.Format(_T("L%d"), i);
			rs.GetValue(szTmp, tmpStr);
			if (tmpStr == 'X')
			{
				xls.SetCellText(nCol+i, nRow, tmpStr);
			}
		}
		rs.GetValue(_T("dept"), tmpStr);
		xls.SetCellText(nCol+4, nRow, tmpStr);
		rs.GetValue(_T("name"), tmpStr);
		xls.SetCellText(nCol+5, nRow, tmpStr);
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\Quyen muc thuoc.xls"));


} 
CString CDrugLevelPermissionRpt::GetQueryString(){
	CString szSQL;
	szSQL.Format(_T(" SELECT") \
				_T("   sup_userid                    AS userid,") \
				_T("   Get_username(sup_userid)      AS name,") \
				_T("   L1,") \
				_T("   L2,") \
				_T("   L3,") \
				_T("   (SELECT") \
				_T("      su_hms_xdept") \
				_T("    FROM   sys_user") \
				_T("    WHERE  su_userid=sup_userid) AS dept") \
				_T(" FROM   (SELECT") \
				_T("           sup_userid,") \
				_T("           CASE WHEN sup_moduleid='EM'") \
				_T("                AND sup_permid='05.06' THEN 'X'") \
				_T("           ELSE 'O'") \
				_T("           END AS L1,") \
				_T("           CASE WHEN sup_moduleid='EM'") \
				_T("                AND sup_permid='05.07' THEN 'X'") \
				_T("           ELSE 'O'") \
				_T("           END AS L2,") \
				_T("           CASE WHEN sup_moduleid='EM'") \
				_T("                AND sup_permid='05.08' THEN 'X'") \
				_T("           ELSE 'O'") \
				_T("           END AS L3") \
				_T("         FROM   sys_userperm") \
				_T("         WHERE  sup_userid IN") \
				_T("                (SELECT") \
				_T("                   su_userid") \
				_T("                 FROM   sys_user") \
				_T("                 WHERE  su_hms_emmodule>0))") \
				_T(" WHERE  L1<>'O'") \
				_T("         OR L2<>'O'") \
				_T("         OR L3<>'O'") \
				_T(" ORDER  BY dept,") \
				_T("           userid "));
	return szSQL;
}


