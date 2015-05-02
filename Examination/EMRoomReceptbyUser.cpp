#include "stdafx.h"
#include "EMRoomReceptbyUser.h"
#include "MainFrame_E10.h"
#include "Excel.h"
#include "StringToken.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMRoomReceptbyUser *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMRoomReceptbyUser *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMRoomReceptbyUser *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMRoomReceptbyUser *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMRoomReceptbyUser *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMRoomReceptbyUser *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMRoomReceptbyUser *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMRoomReceptbyUser *)pWnd)->OnToDateCheckValue();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMRoomReceptbyUser *pVw = (CEMRoomReceptbyUser *)pWnd;
	pVw->OnExportSelect();
} 
static long _OnUserListLoadDataFnc(CWnd *pWnd){
	return ((CEMRoomReceptbyUser*)pWnd)->OnUserListLoadData();
} 
static void _OnUserListDblClickFnc(CWnd *pWnd){
	((CEMRoomReceptbyUser*)pWnd)->OnUserListDblClick();
} 
static void _OnUserListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CEMRoomReceptbyUser*)pWnd)->OnUserListSelectChange(nOldItem, nNewItem);
} 
static int _OnUserListDeleteItemFnc(CWnd *pWnd){
	 return ((CEMRoomReceptbyUser*)pWnd)->OnUserListDeleteItem();
} 
static int _OnUserListCheckAllFnc(CWnd *pWnd){
	((CEMRoomReceptbyUser*)pWnd)->OnUserListCheckAll();
	return 0;
}
static int _OnUserListUncheckAllFnc(CWnd *pWnd){
	((CEMRoomReceptbyUser*)pWnd)->OnUserListUncheckAll();
	return 0;
}
static long _OnRoomListLoadDataFnc(CWnd *pWnd){
	return ((CEMRoomReceptbyUser*)pWnd)->OnRoomListLoadData();
} 
static void _OnRoomListDblClickFnc(CWnd *pWnd){
	((CEMRoomReceptbyUser*)pWnd)->OnRoomListDblClick();
} 
static void _OnRoomListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CEMRoomReceptbyUser*)pWnd)->OnRoomListSelectChange(nOldItem, nNewItem);
} 
static int _OnRoomListDeleteItemFnc(CWnd *pWnd){
	 return ((CEMRoomReceptbyUser*)pWnd)->OnRoomListDeleteItem();
}
static int _OnRoomListCheckAllFnc(CWnd *pWnd){
	((CEMRoomReceptbyUser*)pWnd)->OnRoomListCheckAll();
	return 0;
}
static int _OnRoomListUncheckAllFnc(CWnd *pWnd){
	((CEMRoomReceptbyUser*)pWnd)->OnRoomListUncheckAll();
	return 0;
}
CEMRoomReceptbyUser::CEMRoomReceptbyUser(CWnd *pParent){

	m_nDlgWidth = 505;
	m_nDlgHeight = 445;
	SetDefaultValues();
}
CEMRoomReceptbyUser::~CEMRoomReceptbyUser(){
}
void CEMRoomReceptbyUser::OnCreateComponents(){
	m_wndRoomReceptbyUser.Create(this, _T("Room Recept by User"), 5, 5, 500, 410);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 250, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 255, 30, 335, 55);
	m_wndToDate.Create(this,340, 30, 495, 55); 
	m_wndExport.Create(this, _T("&Export"), 420, 415, 500, 440);
	m_wndUserList.Create(this,10, 60, 250, 405); 
	m_wndUserList.SetCheckBox(true);
	m_wndRoomList.Create(this,255, 60, 495, 405); 
	m_wndRoomList.SetCheckBox(true);
}
void CEMRoomReceptbyUser::OnInitializeComponents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetCheckValue(true);

	m_wndUserList.InsertColumn(0, _T("User ID"), CFMT_TEXT, 90);
	m_wndUserList.InsertColumn(1, _T("User Name"), CFMT_TEXT, 130);

	m_wndRoomList.InsertColumn(0, _T("ID"), CFMT_NUMBER, 40);
	m_wndRoomList.InsertColumn(1, _T("Name"), CFMT_TEXT, 180);
}
void CEMRoomReceptbyUser::OnSetWindowEvents(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndUserList.SetEvent(WE_SELCHANGE, _OnUserListSelectChangeFnc);
	m_wndUserList.SetEvent(WE_LOADDATA, _OnUserListLoadDataFnc);
	m_wndUserList.SetEvent(WE_DBLCLICK, _OnUserListDblClickFnc);
	m_wndUserList.AddEvent(1, _T("Check All"), _OnUserListCheckAllFnc);
	m_wndUserList.AddEvent(2, _T("Uncheck All"), _OnUserListUncheckAllFnc);
	m_wndRoomList.SetEvent(WE_SELCHANGE, _OnRoomListSelectChangeFnc);
	m_wndRoomList.SetEvent(WE_LOADDATA, _OnRoomListLoadDataFnc);
	m_wndRoomList.SetEvent(WE_DBLCLICK, _OnRoomListDblClickFnc);
	m_wndRoomList.AddEvent(1, _T("Check All"), _OnRoomListCheckAllFnc);
	m_wndRoomList.AddEvent(2, _T("Uncheck All"), _OnRoomListUncheckAllFnc);
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	OnUserListLoadData();
	OnRoomListLoadData();
	UpdateData(false);
}
void CEMRoomReceptbyUser::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);

}
void CEMRoomReceptbyUser::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMRoomReceptbyUser::SetMode(int nMode){
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
/*void CEMRoomReceptbyUser::OnFromDateChange(){
	
} */
/*void CEMRoomReceptbyUser::OnFromDateSetfocus(){
	
} */
/*void CEMRoomReceptbyUser::OnFromDateKillfocus(){
	
} */
int CEMRoomReceptbyUser::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMRoomReceptbyUser::OnToDateChange(){
	
} */
/*void CEMRoomReceptbyUser::OnToDateSetfocus(){
	
} */
/*void CEMRoomReceptbyUser::OnToDateKillfocus(){
	
} */
int CEMRoomReceptbyUser::OnToDateCheckValue(){
	return 0;
} 
void CEMRoomReceptbyUser::OnExportSelect(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;

	int nCol = 0, nRow = 0;
	double nCount = 0;
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		AfxMessageBox(_T("No Data"));
		return;
	}
	CStringToken token(m_szUser);
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 25);
	xls.SetCellMergedColumns(0, 2, 5);
	xls.SetCellMergedColumns(0, 3, 5);
	xls.SetCellText(0, 0, pMF->m_szHealthService, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 1, pMF->m_szHospitalName, FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(0, 2, _T("TH\x1ED0NG K\xCA PH\xC2N LO\x1EA0I \x42\x1EC6NH NH\xC2N L\xCAN \x43\xC1\x43 PH\xD2NG TH\x45O US\x45R"), FMT_TEXT | FMT_CENTER, true);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd, ddmmyyyy), CDateTime::Convert(m_szToDate, yyyymmdd, ddmmyyyy));
	xls.SetCellText(0, 3, tmpStr, FMT_TEXT | FMT_CENTER);
	xls.SetCellText(nCol, 5, _T("Ph\xF2ng"), FMT_TEXT | FMT_CENTER, true);
	for (int i = 0; i < token.GetSize(); i++)
	{
		token.GetAt(i, tmpStr);
		tmpStr.Remove('\'');
		xls.SetCellText(nCol + i + 1, 5, pMF->GetUserName(tmpStr), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true);
	}
	nRow = 6;
	while (!rs.IsEOF())
	{
		xls.SetCellText(nCol, nRow, rs.GetValue(_T("roomname")), FMT_TEXT);
		for (int i = 0; i < token.GetSize(); i++)
		{
			token.GetAt(i, tmpStr);
			tmpStr.Remove('\'');
			xls.SetCellText(nCol + i + 1, nRow, rs.GetValue(tmpStr), FMT_NUMBER1 | FMT_WRAPING);
		}
		nRow++;
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\Thong ke benh nhan len phong kham theo user.xls"));
} 
void CEMRoomReceptbyUser::OnUserListDblClick(){
	
} 
void CEMRoomReceptbyUser::OnUserListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CEMRoomReceptbyUser::OnUserListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
} 
long CEMRoomReceptbyUser::OnUserListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndUserList.BeginLoad(); 
	szSQL.Format(_T(" SELECT    DISTINCT sup_userid               userid, ") \
				_T("                    Get_username(sup_userid) username ") \
				_T(" FROM      sys_userperm ") \
				_T(" LEFT JOIN sys_user ON ( su_userid = sup_userid ) ") \
				_T(" WHERE     sup_moduleid = 'RM' ") \
				_T("       AND sup_permid = '01.01' ") \
				_T("       AND su_deptid = 'C1.1' ") \
				_T(" ORDER     BY userid"));
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndUserList.AddItems(
			rs.GetValue(_T("userid")),
			rs.GetValue(_T("username")),
			NULL);
		rs.MoveNext();
	}
	m_wndUserList.EndLoad(); 
	return nCount;
}
void CEMRoomReceptbyUser::OnRoomListDblClick(){
	
} 
void CEMRoomReceptbyUser::OnRoomListSelectChange(int nOldItem, int nNewItem){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	
} 
int CEMRoomReceptbyUser::OnRoomListDeleteItem(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	 return 0;
} 
long CEMRoomReceptbyUser::OnRoomListLoadData(){
	CMainFrame_E10 *pMF = (CMainFrame_E10*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndRoomList.BeginLoad(); 
	szSQL.Format(_T("SELECT hrl_id id, hrl_name roomname FROM hms_roomlist WHERE hrl_deptid = 'C1.1' ORDER BY cast(hrl_id as integer)"));
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndRoomList.AddItems(
			rs.GetValue(_T("id")),
			rs.GetValue(_T("roomname")),
			NULL);
		rs.MoveNext();
	}
	m_wndRoomList.EndLoad(); 
	return nCount;
}

CString CEMRoomReceptbyUser::GetQueryString(){
	CString szSQL, szWhere, szWhere1, szRoom, szCaseExpr, szSumExpr, szUserID;
	m_szUser.Empty();
	for (int i = 0; i < m_wndUserList.GetItemCount(); i++)
	{
		if (m_wndUserList.GetCheck(i))
		{
			if (!m_szUser.IsEmpty())
				m_szUser += _T(",");
			m_szUser.AppendFormat(_T("'%s'"), m_wndUserList.GetItemText(i, 0));
		}
	}
	if (m_szUser.IsEmpty())
		for (int i = 0; i < m_wndUserList.GetItemCount(); i++)
		{
			if (!m_szUser.IsEmpty())
				m_szUser += _T(",");
			m_szUser.AppendFormat(_T("'%s'"), m_wndUserList.GetItemText(i, 0));
		}
	for (int i = 0; i < m_wndRoomList.GetItemCount(); i++)
	{
		if (m_wndRoomList.GetCheck(i))
		{
			if (!szRoom.IsEmpty())
				szRoom += _T(",");
			szRoom.AppendFormat(_T("%s"), m_wndRoomList.GetItemText(i, 0));
		}
	}
	szWhere.Format(_T(" AND he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szUser.IsEmpty())
		szWhere1.AppendFormat(_T(" AND clerk IN (%s)"), m_szUser);
	if (!szRoom.IsEmpty())
		szWhere1.AppendFormat(_T(" AND roomid IN (%s)"), szRoom);
	CStringToken token(m_szUser);
	TCHAR tc = '\'';
	for (int i =0; i < token.GetSize(); i++){
		token.GetAt(i, szUserID);
		szUserID.Remove(tc);
		szCaseExpr.AppendFormat(_T("case when clerk = '%s' then 1 else 0 end as %s,"), szUserID, szUserID);
		szSumExpr.AppendFormat(_T("sum(%s) %s,"), szUserID, szUserID);
	}
	_debug(_T("%s||%s"), szCaseExpr, szSumExpr);
	

	szSQL.Format(_T(" SELECT %s (SELECT hrl_name ") \
				_T("         FROM   hms_roomlist ") \
				_T("         WHERE  hrl_id = roomid ") \
				_T("            AND hrl_deptid = 'C1.1') roomname ") \
				_T(" FROM   (SELECT hd_docno, ") \
				_T("                %s roomid ") \
				_T("         FROM   (SELECT    distinct hd_docno, ") \
				_T("                           First_value(he_createdby) ") \
				_T("                             over ( ") \
				_T("                               PARTITION BY hd_docno ") \
				_T("                               ORDER BY he_receptidx) clerk, ") \
				_T("                           First_value(he_roomid) ") \
				_T("                             over ( ") \
				_T("                               PARTITION BY hd_docno ") \
				_T("                               ORDER BY he_receptidx) roomid ") \
				_T("                 FROM      hms_doc ") \
				_T("                 LEFT JOIN hms_exam ON ( hd_docno = he_docno ) ") \
				_T("                 WHERE     he_status IN ('P', 'T') AND he_deptid = 'C1.1' %s) WHERE 1=1 %s)") \
				_T(" GROUP  BY roomid ") \
				_T(" ORDER  BY roomid "),szSumExpr, szCaseExpr, szWhere, szWhere1);
		
	return szSQL;
}

void CEMRoomReceptbyUser::OnUserListCheckAll(){
	
}

void CEMRoomReceptbyUser::OnUserListUncheckAll(){

}

void CEMRoomReceptbyUser::OnRoomListCheckAll(){

}
void CEMRoomReceptbyUser::OnRoomListUncheckAll(){
	
}