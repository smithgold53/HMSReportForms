#include "StdAfx.h"
#include "PatientProfileDrugLogs.h"
#include "GuiMainFrame.h"
static long _OnListLoadDataFnc(CWnd *pWnd){
	return ((CPatientProfileDrugLogs*)pWnd)->OnListLoadData();
} 
static void _OnListDblClickFnc(CWnd *pWnd){
	((CPatientProfileDrugLogs*)pWnd)->OnListDblClick();
} 
static void _OnListSelectChangeFnc(CWnd *pWnd, int nOldItem, int nNewItem){
	((CPatientProfileDrugLogs*)pWnd)->OnListSelectChange(nOldItem, nNewItem);
} 
static int _OnListDeleteItemFnc(CWnd *pWnd){
	 return ((CPatientProfileDrugLogs*)pWnd)->OnListDeleteItem();
} 
static int _OnAddPatientProfileDrugLogsFnc(CWnd *pWnd){
	 return ((CPatientProfileDrugLogs*)pWnd)->OnAddPatientProfileDrugLogs();
} 
static int _OnEditPatientProfileDrugLogsFnc(CWnd *pWnd){
	 return ((CPatientProfileDrugLogs*)pWnd)->OnEditPatientProfileDrugLogs();
} 
static int _OnDeletePatientProfileDrugLogsFnc(CWnd *pWnd){
	 return ((CPatientProfileDrugLogs*)pWnd)->OnDeletePatientProfileDrugLogs();
} 
static int _OnSavePatientProfileDrugLogsFnc(CWnd *pWnd){
	 return ((CPatientProfileDrugLogs*)pWnd)->OnSavePatientProfileDrugLogs();
} 
static int _OnCancelPatientProfileDrugLogsFnc(CWnd *pWnd){
	 return ((CPatientProfileDrugLogs*)pWnd)->OnCancelPatientProfileDrugLogs();
} 
CPatientProfileDrugLogs::CPatientProfileDrugLogs(){

	m_nDlgWidth = 850;
	m_nDlgHeight = 585;
	SetDefaultValues();
}
CPatientProfileDrugLogs::~CPatientProfileDrugLogs(){
}
void CPatientProfileDrugLogs::OnCreateComponents(){
	m_wndDrugLogsInformation.Create(this, _T("Information"), 5, 5, 845, 580);
	m_wndList.Create(this,10, 30, 840, 570); 

}
void CPatientProfileDrugLogs::OnInitializeComponents(){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	//m_wndList.SetAutoIndex(true);
}
void CPatientProfileDrugLogs::OnSetWindowEvents(){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	m_wndList.SetEvent(WE_SELCHANGE, _OnListSelectChangeFnc);
	m_wndList.SetEvent(WE_LOADDATA, _OnListLoadDataFnc);
	m_wndList.SetEvent(WE_DBLCLICK, _OnListDblClickFnc);
	/*
	m_wndList.AddEvent(1, _T("Delete"), _OnListDeleteItemFnc, 0, VK_DELETE, 0);
	AddEvent(1, _T("Add	Ctrl+A"), _OnAddPatientProfileDrugLogsFnc, 0, 'A', VK_CONTROL);
	AddEvent(2, _T("Edit	Ctrl+E"), _OnEditPatientProfileDrugLogsFnc, 0, 'E', VK_CONTROL);
	AddEvent(3, _T("Delete	Ctrl+D"), _OnDeletePatientProfileDrugLogsFnc, 0, 'D', VK_CONTROL);
	AddEvent(4, _T("Save	Ctrl+S"), _OnSavePatientProfileDrugLogsFnc, 0, 'S', VK_CONTROL);
	AddEvent(0, _T("-"));
	AddEvent(5, _T("Cancel	Ctrl+T"), _OnCancelPatientProfileDrugLogsFnc, 0, 'T', VK_CONTROL);
	*/
	SetMode(VM_NONE);

}
void CPatientProfileDrugLogs::OnDoDataExchange(CDataExchange* pDX){

}
void CPatientProfileDrugLogs::GetDataToScreen(){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CPatientProfileDrugLogs::GetScreenToData(){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();

}
void CPatientProfileDrugLogs::SetDefaultValues(){


}
int CPatientProfileDrugLogs::SetMode(int nMode){
 		int nOldMode = GetMode();
 		CGuiView::SetMode(nMode);
 		CGuiMainFrame *pMF = (CGuiMainFrame *) AfxGetMainWnd();
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
void CPatientProfileDrugLogs::OnListDblClick(){
	
} 
void CPatientProfileDrugLogs::OnListSelectChange(int nOldItem, int nNewItem){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	
} 
int CPatientProfileDrugLogs::OnListDeleteItem(){
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	 return 0;
} 
long CPatientProfileDrugLogs::OnListLoadData(){
/*
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	m_wndList.BeginLoad(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndList.AddItems(
			rs.GetValue(_T("Idx")), 
			rs.GetValue(_T("Description")), NULL);
		rs.MoveNext();
	}
	m_wndList.EndLoad(); 
	return nCount;
*/
	return 0;
}
int CPatientProfileDrugLogs::OnAddPatientProfileDrugLogs(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CGuiMainFrame *pMF = (CGuiMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CPatientProfileDrugLogs::OnEditPatientProfileDrugLogs(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CGuiMainFrame *pMF = (CGuiMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CPatientProfileDrugLogs::OnDeletePatientProfileDrugLogs(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CGuiMainFrame *pMF = (CGuiMainFrame *)AfxGetMainWnd();
 	CString szSQL;
 	if(ShowMessage(1, MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2) == IDNO)
 		return -1;
 	szSQL.Format(_T("DELETE FROM  WHERE  AND") );
 	int ret = pMF->ExecSQL(szSQL);
 	if(ret >= 0){
 		SetMode(VM_NONE);
 		OnCancelPatientProfileDrugLogs();
 	}
	return 0;
}
int CPatientProfileDrugLogs::OnSavePatientProfileDrugLogs(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT)
 		return -1;
 	if(!IsValidateData())
 		return -1;
 	CGuiMainFrame *pMF = (CGuiMainFrame *)AfxGetMainWnd();
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
 		//OnPatientProfileDrugLogsListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CPatientProfileDrugLogs::OnCancelPatientProfileDrugLogs(){
 	if(GetMode() == VM_EDIT)
 	{
 		SetMode(VM_VIEW);
 	} 
 	else 
 	{
 		SetMode(VM_NONE);
 	} 
 	CGuiMainFrame *pMF = (CGuiMainFrame *)AfxGetMainWnd();
 	return 0;
} 
int CPatientProfileDrugLogs::OnPatientProfileDrugLogsListLoadData(){
	return 0;
}

void CPatientProfileDrugLogs::CreateColumns(CString szType, int nIdx)
{
	CGuiMainFrame *pMF = (CGuiMainFrame *) AfxGetMainWnd();
	
	m_szType  = szType;
	m_nColumn = nIdx;

	CString tmpStr;
	int nCol = 0;

	m_wndList.DeleteAllItems();
	m_wndList.DeleteAllColumns();
	
	m_wndList.GetHeaderControl()->SetItemHeight(2);
	
	if(szType == _T("DRUGLOG"))
	{
		CRecord rs(&pMF->m_db);
		CString szSQL;

		m_wndList.InsertColumn(nCol++, _T("Stt"), CFMT_NUMBER, 0);
		m_wndList.InsertColumn(nCol++, _T("T\xEAn thu\x1ED1\x63/H\xE0m l\x1B0\x1EE3ng"), CFMT_TEXT, 150);
		m_wndList.InsertColumn(nCol++, _T("\x110\x1A1n v\x1ECB"), CFMT_TEXT, 50);

		CDate minDate;
		CDate maxDate;
		CDate curDate;
		CDate dte;

		maxDate.ParseDate(pMF->GetSysDate());
		curDate = maxDate;
		minDate = maxDate;

		szSQL.Format(_T("SELECT min(trunc(hpo_orderdate))  as min_date, max(trunc(hpo_orderdate)) as max_date ") \
			_T(" FROM hms_ipharmaorder ") \
			_T(" WHERE hpo_docno=%ld and hpo_orderstatus <> 'C' and hpo_ordertype IN('D','P') "), m_nDocumentNo);
		rs.ExecSQL(szSQL);


		if(!rs.IsEOF())
		{
				rs.GetValue(_T("min_date"), tmpStr);
				minDate.ParseDate(tmpStr);
				rs.GetValue(_T("max_date"), tmpStr);
				dte.ParseDate(tmpStr);

				if(dte > maxDate)
				{
					maxDate = dte;
				}
		}



		if(m_nColumn <= 0)
		{


		}
		else
		{
			minDate = maxDate;
			for (int i =0; i < m_nColumn; i++)
					minDate--;
		}
		m_arrDays.RemoveAll();
		for(dte = minDate; dte <= maxDate; dte++)
		{
			m_arrDays.Add(dte.GetDate());
			
		}

		m_minDate = minDate;
		m_maxDate = maxDate;
		CString szDate;
		for (int i =0; i < m_arrDays.GetCount(); i++)
		{
			tmpStr = m_arrDays[i];
			szDate.Format(_T("%s/%s"), tmpStr.Right(2), tmpStr.Mid(5, 2));
			m_wndList.InsertColumn(nCol++, szDate, CFMT_NUMBER, 45);
		}
		m_wndList.GetHeaderControl()->MergeCell(_T("Ng\xE0y s\x1EED \x64\x1EE5ng"), 0, 3, 1, nCol-1);
		m_wndList.InsertColumn(nCol++, _T("\x43\x1ED9ng"), CFMT_NUMBER, 90);
		m_wndList.InsertColumn(nCol++, _T("T\x1ED5ng s\x1ED1"), CFMT_NUMBER, 0);

	}
	else 
	{
		m_wndList.InsertColumn(nCol++, _T("Stt"), CFMT_NUMBER, 35);
		m_wndList.InsertColumn(nCol++, _T("\x44i\x1EC5n gi\x1EA3i"), CFMT_TEXT, 400);
		m_wndList.InsertColumn(nCol++, _T("\x110\x1A1n v\x1ECB"), CFMT_TEXT, 125);
		m_wndList.InsertColumn(nCol++, _T("K\x1EBFt qu\x1EA3"), CFMT_TEXT, 125);
		

	}
	m_wndList.SetSortHeader(false);
}

void CPatientProfileDrugLogs::Refresh()
{
	CGuiMainFrame *pMF = (CGuiMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CRecord rsl(&pMF->m_db);
	CString szSQL, tmpStr, szName;
	int nIndex =0, nLine= 0;
	long nOrderID;
	long nProductID=0;
	m_wndList.BeginLoad();
	if(m_szType == _T("DRUGLOG"))
	{
		szSQL.Format(_T(" SELECT distinct hpol_product_id AS product_id,") \
_T("   mp_name              AS product_name,") \
_T("   adu_name             AS uom_name,") \
_T("   adu_index            AS uom_index") \
_T(" FROM hms_ipharmaorder") \
_T(" LEFT JOIN hms_ipharmaorderline") \
_T(" ON(hpol_orderid=hpo_orderid)") \
_T(" LEFT JOIN m_product") \
_T(" ON(mp_product_id = hpol_product_id)") \
_T(" LEFT JOIN ad_uom") \
_T(" ON(adu_uom_id   = mp_product_uom_id)") \
_T(" WHERE hpo_docno = %ld") \
_T(" AND TRUNC(hpo_orderdate) BETWEEN TO_DATE('%s','YYYY-MM-DD') AND TO_DATE('%s','YYYY-MM-DD')") \
_T(" AND hpo_orderstatus<>'C' and hpo_ordertype IN('D','P') ") \
_T(" AND hpol_qtyissue  > 0") \
_T(" ORDER BY adu_index, mp_name "), m_nDocumentNo, m_minDate.GetDate(), m_maxDate.GetDate());
		rs.ExecSQL(szSQL);
		//DANH SACH THUOC

		int nLine = 1;
		float nTtlQty = 0;
		while(!rs.IsEOF())
		{
			rs.GetValue(_T("product_id"), nProductID);
			
			tmpStr.Format(_T("%d"), nLine++);

			int nItem = m_wndList.AddItems(tmpStr,
				rs.GetValue(_T("product_name")),
				rs.GetValue(_T("uom_name")),
				NULL);

			szSQL.Format(_T(" SELECT TO_CHAR(hpo_orderdate, 'YYYY/MM/DD') AS approvedate,") \
_T("   SUM(hpol_qtyissue)          AS qtyissue") \
_T(" FROM hms_ipharmaorder") \
_T(" LEFT JOIN hms_ipharmaorderline") \
_T(" ON(hpol_orderid     =hpo_orderid)") \
_T(" WHERE hpo_docno     = %ld") \
_T(" AND hpol_product_id = %ld") \
_T(" AND TRUNC(hpo_orderdate) BETWEEN TO_DATE('%s','YYYY-MM-DD') AND TO_DATE('%s','YYYY-MM-DD')") \
_T(" AND hpo_orderstatus<>'C' and hpo_ordertype IN('D','P') ") \
_T(" AND hpol_qtyissue  > 0") \
_T(" GROUP BY TO_CHAR(hpo_orderdate,'YYYY/MM/DD')"), m_nDocumentNo, nProductID, m_minDate.GetDate(), m_maxDate.GetDate());

			rsl.ExecSQL(szSQL);
			nTtlQty = 0;
			float nQty;
			int nCol = 3;
			float nSumQty = 0;
			while(!rsl.IsEOF())
			{
				rsl.GetValue(_T("approvedate"), tmpStr);
				rsl.GetValue(_T("qtyissue"), nQty);
				
				nSumQty += nQty;
				CString szDate;
				for(int i = 0; i < m_arrDays.GetCount(); i++)
				{
					szDate = m_arrDays[i];
					
					if(tmpStr == szDate)
					{
						m_wndList.SetItemText(nItem, i+3, ToString(nQty));
						nTtlQty += nQty;
					}
					
				}
				
				rsl.MoveNext();
			}
			m_wndList.SetItemText(nItem, m_wndList.GetHeaderControl()->GetItemCount()-2, ToString(nTtlQty));
			m_wndList.SetItemText(nItem, m_wndList.GetHeaderControl()->GetItemCount()-1, ToString(nSumQty));



			rs.MoveNext();
		}
		
	}
	if(m_szType == _T("TEST"))
	{
		szSQL.Format(_T(" SELECT hpc_deptid as deptid,") \
		_T("   hpc_orderid as orderid,") \
		_T("   TO_CHAR(hpc_orderdate,'DD/MM/YYYY HH24:MI') as orderdate,") \
		_T("   TO_CHAR(hpc_performdate, 'DD/MM/YYYY HH24:MI') as performdate,") \
		_T("   hpc_status as orderstatus, ") \
		_T(" HMS_GetFeeGroupName(hpc_groupid) as grpname ") \
		_T(" FROM hms_testorder") \
		_T(" WHERE hpc_docno = %ld") \
		_T(" ORDER BY hpc_orderid"), m_nDocumentNo);

		rs.ExecSQL(szSQL);
		
		while(!rs.IsEOF())
		{
			nIndex++;
			tmpStr.Format(_T("%d"), nIndex);
			rs.GetValue(_T("orderid"), nOrderID);
			szName.Format(_T("%s [%s]"), rs.GetValue(_T("grpname")), rs.GetValue(_T("deptid")));
			int nItem = m_wndList.AddItems(tmpStr,
				szName,
				rs.GetValue(_T("orderid")),
				rs.GetValue(_T("orderdate")),
				NULL);
			m_wndList.SetItemBkColor(nItem, RGB(0, 128, 192), FALSE);
			m_wndList.SetItemTextColor(nItem, RGB(255, 255, 255), FALSE);

			szSQL.Format(_T(" SELECT hfl_name AS testname,") \
			_T("   hfl_unit      AS unit,") \
			_T("   hpcl_result   AS result") \
			_T(" FROM hms_testorderline") \
			_T(" LEFT JOIN hms_fee_list") \
			_T(" ON(hfl_feeid       =hpcl_itemid)") \
			_T(" WHERE hpcl_orderid = %ld") \
			_T(" ORDER BY hfl_line"), nOrderID);
			rsl.ExecSQL(szSQL);
			nLine = 0;
			while(!rsl.IsEOF())
			{
				nLine++;
				tmpStr.Format(_T("%d"), nLine);
				m_wndList.AddItems(tmpStr,
					rsl.GetValue(_T("testname")),
					rsl.GetValue(_T("unit")),
					rsl.GetValue(_T("result")),
					NULL);

				rsl.MoveNext();
			}


			rs.MoveNext();
		}
	}

	if(m_szType == _T("PACS"))
	{
		szSQL.Format(_T(" SELECT hpc_deptid as deptid,") \
_T("   hpc_orderid as orderid,") \
_T("   TO_CHAR(hpc_orderdate,'DD/MM/YYYY HH24:MI') as orderdate,") \
_T("   TO_CHAR(hpc_performdate, 'DD/MM/YYYY HH24:MI') as performdate,") \
_T("   hpc_status as orderstatus, ") \
_T(" HMS_GetFeeGroupName(hpc_groupid) as grpname ") \
_T(" FROM hms_pacsorder") \
_T(" WHERE hpc_docno = %ld") \
_T(" ORDER BY hpc_orderid"), m_nDocumentNo);

		rs.ExecSQL(szSQL);
		
		while(!rs.IsEOF())
		{
			nIndex++;
			tmpStr.Format(_T("%d"), nIndex);
			rs.GetValue(_T("orderid"), nOrderID);
			szName.Format(_T("%s [%s]"), rs.GetValue(_T("grpname")), rs.GetValue(_T("deptid")));
			int nItem = m_wndList.AddItems(tmpStr,
				szName,
				rs.GetValue(_T("orderid")),
				rs.GetValue(_T("orderdate")),
				NULL);
			m_wndList.SetItemBkColor(nItem, RGB(0, 128, 192), FALSE);
			m_wndList.SetItemTextColor(nItem, RGB(255, 255, 255), FALSE);

			szSQL.Format(_T(" SELECT hfl_name AS testname,") \
_T("   hfl_unit      AS unit,") \
_T("   hpcl_conclusion   AS result") \
_T(" FROM hms_pacsorderline") \
_T(" LEFT JOIN hms_fee_list") \
_T(" ON(hfl_feeid       =hpcl_itemid)") \
_T(" WHERE hpcl_orderid = %ld") \
_T(" ORDER BY hfl_line"), nOrderID);
			rsl.ExecSQL(szSQL);
			nLine = 0;
			while(!rsl.IsEOF())
			{
				nLine++;
				tmpStr.Format(_T("%d"), nLine);
				m_wndList.AddItems(tmpStr,
					rsl.GetValue(_T("testname")),
					rsl.GetValue(_T("unit")),
					rsl.GetValue(_T("result")),
					NULL);

				rsl.MoveNext();
			}


			rs.MoveNext();
		}
	}

	
	if(m_szType == _T("SURGERY"))
	{
		szSQL.Format(_T(" SELECT ho_deptid as deptid, ") \
			_T("TO_CHAR(ho_orderdate,'DD/MM/YYYY HH24:MI') as orderdate, ") \
			_T("hfl_name AS optname,") \
_T("   hfl_unit      AS unit,") \
_T("   ho_anes_result   AS result") \
_T(" FROM hms_operation") \
_T(" LEFT JOIN hms_fee_list") \
_T(" ON(hfl_feeid       =ho_itemid)") \
_T(" WHERE ho_docno = %ld") \
_T(" ORDER BY hfl_line"), m_nDocumentNo);
			rsl.ExecSQL(szSQL);
			nLine = 0;
			while(!rsl.IsEOF())
			{
				nLine++;
				tmpStr.Format(_T("%d"), nLine);
				szName.Format(_T("[%s] %s (%s) "), rsl.GetValue(_T("deptid")), rsl.GetValue(_T("optname")), rsl.GetValue(_T("orderdate")));

				m_wndList.AddItems(tmpStr,
					szName,
					rsl.GetValue(_T("unit")),
					rsl.GetValue(_T("result")),
					NULL);

				rsl.MoveNext();
			}


	}


	m_wndList.EndLoad();
}