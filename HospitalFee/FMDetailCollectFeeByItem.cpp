#include "stdafx.h"
#include "FMDetailCollectFeeByItem.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "ReportDocument.h"


/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMDetailCollectFeeByItem *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMDetailCollectFeeByItem *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMDetailCollectFeeByItem *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMDetailCollectFeeByItem *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMDetailCollectFeeByItem *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMDetailCollectFeeByItem *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMDetailCollectFeeByItem *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMDetailCollectFeeByItem *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMDetailCollectFeeByItem* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMDetailCollectFeeByItem *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMDetailCollectFeeByItem *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMDetailCollectFeeByItem *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMDetailCollectFeeByItem *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMDetailCollectFeeByItem *)pWnd)->OnClerkAddNew();
}*/
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMDetailCollectFeeByItem *pVw = (CFMDetailCollectFeeByItem *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMDetailCollectFeeByItem *pVw = (CFMDetailCollectFeeByItem *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMDetailCollectFeeByItem*)pWnd)->OnLockedSelect();
}
static int _OnAddFMDetailCollectFeeByItemFnc(CWnd *pWnd){
	 return ((CFMDetailCollectFeeByItem*)pWnd)->OnAddFMDetailCollectFeeByItem();
} 
static int _OnEditFMDetailCollectFeeByItemFnc(CWnd *pWnd){
	 return ((CFMDetailCollectFeeByItem*)pWnd)->OnEditFMDetailCollectFeeByItem();
} 
static int _OnDeleteFMDetailCollectFeeByItemFnc(CWnd *pWnd){
	 return ((CFMDetailCollectFeeByItem*)pWnd)->OnDeleteFMDetailCollectFeeByItem();
} 
static int _OnSaveFMDetailCollectFeeByItemFnc(CWnd *pWnd){
	 return ((CFMDetailCollectFeeByItem*)pWnd)->OnSaveFMDetailCollectFeeByItem();
} 
static int _OnCancelFMDetailCollectFeeByItemFnc(CWnd *pWnd){
	 return ((CFMDetailCollectFeeByItem*)pWnd)->OnCancelFMDetailCollectFeeByItem();
} 
CFMDetailCollectFeeByItem::CFMDetailCollectFeeByItem(CWnd* pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 160;
	SetDefaultValues();
}
CFMDetailCollectFeeByItem::~CFMDetailCollectFeeByItem(){
}
void CFMDetailCollectFeeByItem::OnCreateComponents()
{
	m_wndCollectDetailbyItem.Create(this, _T("Collect Detail by Item"), 5, 5, 490, 120);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 485, 85); 
	//m_wndPrintPreview.Create(this, _T("&Print Preview"), 265, 125, 375, 150);
	m_wndExport.Create(this, _T("&Export"), 380, 125, 490, 150);
	m_wndLocked.Create(this, _T("Locked"), 375, 90, 485, 115);
}
void CFMDetailCollectFeeByItem::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(65);

	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 70);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 280);
}
void CFMDetailCollectFeeByItem::OnSetWindowEvents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndClerk.SetEvent(WE_SELENDOK, _OnClerkSelendokFnc);
	//m_wndClerk.SetEvent(WE_SETFOCUS, _OnClerkSetfocusFnc);
	//m_wndClerk.SetEvent(WE_KILLFOCUS, _OnClerkKillfocusFnc);
	m_wndClerk.SetEvent(WE_SELCHANGE, _OnClerkSelectChangeFnc);
	m_wndClerk.SetEvent(WE_LOADDATA, _OnClerkLoadDataFnc);
	//m_wndClerk.SetEvent(WE_ADDNEW, _OnClerkAddNewFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);
	
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);
}
void CFMDetailCollectFeeByItem::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);

}
void CFMDetailCollectFeeByItem::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMDetailCollectFeeByItem::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMDetailCollectFeeByItem::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_bLocked=FALSE;

}
int CFMDetailCollectFeeByItem::SetMode(int nMode){
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
/*void CFMDetailCollectFeeByItem::OnFromDateChange(){
	
} */
/*void CFMDetailCollectFeeByItem::OnFromDateSetfocus(){
	
} */
/*void CFMDetailCollectFeeByItem::OnFromDateKillfocus(){
	
} */
int CFMDetailCollectFeeByItem::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMDetailCollectFeeByItem::OnToDateChange(){
	
} */
/*void CFMDetailCollectFeeByItem::OnToDateSetfocus(){
	
} */
/*void CFMDetailCollectFeeByItem::OnToDateKillfocus(){
	
} */
int CFMDetailCollectFeeByItem::OnToDateCheckValue(){
	return 0;
} 
void CFMDetailCollectFeeByItem::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMDetailCollectFeeByItem::OnClerkSelendok(){
	 
}
/*void CFMDetailCollectFeeByItem::OnClerkSetfocus(){
	
}*/
/*void CFMDetailCollectFeeByItem::OnClerkKillfocus(){
	
}*/
long CFMDetailCollectFeeByItem::OnClerkLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	szWhere.Empty();

	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	}

	m_wndClerk.DeleteAllItems(); 
	szSQL.Format(_T("SELECT su_userid as userid, su_name as name FROM sys_user WHERE su_groupid in('A', 'F') %s ORDER BY su_userid"), szWhere);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndClerk.AddItems(
			rs.GetValue(_T("userid")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CFMDetailCollectFeeByItem::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CFMDetailCollectFeeByItem::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMDetailCollectFeeByItem::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

	UpdateData(true);

	BeginWaitCursor();

	CExcel xls;

	CRecord rs(&pMF->m_db);
	CRecord rss(&pMF->m_db);

	CString szSQL, szWhere;
	CString tmpStr, szTemp;
	CString szCondition;

	int nCol = 0, nRow = 0, nIdx = 0;
	double nTemp = 0;
	long double nGrpTimes = 0, nGrpExam = 0, nGrpIn = 0, nGrpAmount = 0;
	long double nTotalExam = 0, nTotalIn = 0, nTotalAmount = 0;

	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessageBox(_T("No Data"), MB_OK | MB_ICONERROR);
		return;
	}


	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_staff='%s' "), m_szClerkKey);

	if (!m_bLocked)
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_locked<>'Y' "));
		szWhere.AppendFormat(_T(" AND fi.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}
	else
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_locked='Y' "));
		szWhere.AppendFormat(_T(" AND fi.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}

	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 50);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 15);
	xls.SetColumnWidth(5, 15);

	xls.SetRowHeight(4, 40);
	
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 2);
	xls.SetCellMergedColumns(nCol, nRow + 1, 2);

	xls.SetCellMergedColumns(nCol + 2, nRow, 4);
	xls.SetCellMergedColumns(nCol + 2, nRow + 1, 4);

	xls.SetCellMergedColumns(nCol, nRow + 2, 6);
	xls.SetCellMergedColumns(nCol, nRow + 3, 6);

	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);

	xls.SetCellText(nCol + 2, nRow, _T("\x43\x1ED8NG H\xD2\x41 \x58\xC3 H\x1ED8I \x43H\x1EE6 NGH\x128\x41 VI\x1EC6T N\x41M"), FMT_TEXT | FMT_CENTER, true);
	xls.SetCellText(nCol + 2, nRow + 1, _T("\x110\x1ED8\x43 L\x1EACP - T\x1EF0 \x44O - H\x1EA0NH PH\xDA\x43"), FMT_TEXT | FMT_CENTER, true);

	xls.SetCellText(nCol, nRow + 2, _T("\x43HI TI\x1EBET THU TI\x1EC0N TH\x45O \x44\x41NH M\x1EE4\x43"),
		            FMT_TEXT | FMT_CENTER, true, 11);

	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("\x44i\x1EC5n gi\x1EA3i"));
	arrCol.Add(_T("S\x1ED1 l\x1EA7n"));
	arrCol.Add(_T("Thu Ph\xF2ng kh\xE1m"));
	arrCol.Add(_T("Thu n\x1ED9i tr\xFA"));
	arrCol.Add(_T("S\x1ED1 ti\x1EC1n"));

	nRow = 4;

	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER | FMT_VCENTER | FMT_WRAPING, true, 10); 
	}

	nRow = 5;
	CString szOldFeeGrp, szNewFeeGrp;
	int nPreviousRow = 0;
	
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("deptid"), szNewFeeGrp);
		if (szNewFeeGrp != szOldFeeGrp)
		{
			if (nGrpTimes > 0)
			{
				tmpStr.Format(_T("%.0Lf"), nGrpTimes);
				xls.SetCellText(nCol + 2, nPreviousRow, tmpStr, FMT_INTEGER | FMT_WRAPING, true, 10);

				tmpStr.Format(_T("%.2Lf"), nGrpExam);
				xls.SetCellText(nCol + 3, nPreviousRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);

				tmpStr.Format(_T("%.2Lf"), nGrpIn);
				xls.SetCellText(nCol + 4, nPreviousRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);

				tmpStr.Format(_T("%.2Lf"), nGrpAmount);
				xls.SetCellText(nCol + 5, nPreviousRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
			}

			nIdx++;
			tmpStr.Format(_T("%d"), nIdx);
			xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING, true, 10);

			if (szNewFeeGrp != _T("X"))
			{
				rs.GetValue(_T("deptname"), tmpStr);
			}
			else
			{
				tmpStr = _T("\x43hi ph\xED kh\xE1\x63");
			}

			xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING, true, 10);

			szOldFeeGrp = szNewFeeGrp;
			nPreviousRow = nRow;
			nGrpTimes = 0;
			nTotalExam += nGrpExam;
			nGrpExam = 0;
			nTotalIn += nGrpIn;
			nGrpIn = 0;
			nTotalAmount += nGrpAmount;
			nGrpAmount = 0;
			nRow++;
		}

		rs.GetValue(_T("itemname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT | FMT_WRAPING);


		rs.GetValue(_T("total"), nTemp);
		nGrpTimes += nTemp;
		tmpStr.Format(_T("%.0f"), nTemp);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);
		

		rs.GetValue(_T("exam_fee"), nTemp);
		nGrpExam += nTemp;
		tmpStr.Format(_T("%.2f"), nTemp);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);
		

		rs.GetValue(_T("inpatient_fee"), nTemp);
		nGrpIn += nTemp;
		tmpStr.Format(_T("%.2f"), nTemp);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_RIGHT);
		

		rs.GetValue(_T("total_amount"), nTemp);
		nGrpAmount += nTemp;
		tmpStr.Format(_T("%.2f"), nTemp);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		szCondition.Format(_T(" AND fe.hfe_itemid='%s' "),
			               rs.GetValue(_T("itemid")));

		szSQL.Format(_T(" SELECT fi.hfe_docno AS docno,") \
					_T("        trim(hp_surname||' '||hp_midname||' '||hp_firstname) AS pname,") \
					_T("        SUM(fe.hfe_patpaid) AS amount") \
					_T(" FROM hms_fee_invoice fi") \
					_T(" LEFT JOIN hms_fee fe ON(fi.hfe_invoiceno=fe.hfe_invoiceno AND fi.hfe_docno=fe.hfe_docno)") \
					_T(" LEFT JOIN hms_doc ON(hd_docno=fi.hfe_docno)") \
					_T(" LEFT JOIN hms_patient ON(hp_patientno=hd_patientno)") \
					_T(" WHERE fi.hfe_class IN('E','X') %s %s") \
					_T("       AND fe.hfe_type IN('E','O','P','T','X','D') AND fi.hfe_status='P'") \
					_T(" GROUP BY fi.hfe_docno, hp_surname, hp_midname, hp_firstname"),
					szWhere, szCondition);

		rss.ExecSQL(szSQL);

		while (!rss.IsEOF())
		{
			nRow++;

			szTemp.Format(_T("%s [%s - %s]"), rss.GetValue(_T("docno")),
				          rss.GetValue(_T("pname")), rss.GetValue(_T("amount")));
			xls.SetCellText(nCol + 1, nRow, szTemp, FMT_TEXT | FMT_WRAPING);

			rss.MoveNext();
		}

		nRow++;
		rs.MoveNext();
	}

	if (nGrpTimes > 0)
	{
		tmpStr.Format(_T("%.0Lf"), nGrpTimes);
		xls.SetCellText(nCol + 2, nPreviousRow, tmpStr, FMT_INTEGER | FMT_WRAPING, true, 10);
		
		tmpStr.Format(_T("%.2Lf"), nGrpExam);
		xls.SetCellText(nCol + 3, nPreviousRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
		
		tmpStr.Format(_T("%.2Lf"), nGrpIn);
		xls.SetCellText(nCol + 4, nPreviousRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
		
		tmpStr.Format(_T("%.2Lf"), nGrpAmount);
		xls.SetCellText(nCol + 5, nPreviousRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);

		nTotalExam += nGrpExam;
		nTotalIn += nGrpIn;
		nTotalAmount += nGrpAmount;
	}

	TranslateString(_T("Total"), tmpStr);
	xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT, true, 10);

	tmpStr.Format(_T("%.2Lf"), nTotalExam);
	xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
	
	tmpStr.Format(_T("%.2Lf"), nTotalIn);
	xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
	
	tmpStr.Format(_T("%.2Lf"), nTotalAmount);
	xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);

	xls.Save(_T("Exports\\CTThuTienTheoDM.xls"));
	EndWaitCursor();
} 
void CFMDetailCollectFeeByItem::OnLockedSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
int CFMDetailCollectFeeByItem::OnAddFMDetailCollectFeeByItem(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CFMDetailCollectFeeByItem::OnEditFMDetailCollectFeeByItem(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMDetailCollectFeeByItem::OnDeleteFMDetailCollectFeeByItem(){
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
 		OnCancelFMDetailCollectFeeByItem();
 	}
	return 0;
}
int CFMDetailCollectFeeByItem::OnSaveFMDetailCollectFeeByItem(){
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
 		//OnFMDetailCollectFeeByItemListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CFMDetailCollectFeeByItem::OnCancelFMDetailCollectFeeByItem(){
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
int CFMDetailCollectFeeByItem::OnFMDetailCollectFeeByItemListLoadData(){
	return 0;
}

CString CFMDetailCollectFeeByItem::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Empty();

	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_staff='%s' "), m_szClerkKey);

	if (!m_bLocked)
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_locked<>'Y' "));
		szWhere.AppendFormat(_T(" AND fi.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}
	else
	{
		szWhere.AppendFormat(_T(" AND fi.hfe_locked='Y' "));
		szWhere.AppendFormat(_T(" AND fi.hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}

	szSQL.Format(_T(" SELECT CASE WHEN deptid IN('C1.1','C1.2','C1.3') THEN 'KB'") \
				_T("             WHEN deptid IN('KA') THEN 'KD'") \
				_T("             WHEN deptid IN('X') THEN 'XX' ELSE 'DT' END AS depttype,") \
				_T("        deptid,") \
				_T("        sd_name AS deptname,") \
				_T("        SUM(count) AS total,") \
				_T("        idx,") \
				_T("        itemid,") \
				_T("        itemname,") \
				_T("        SUM(examfee) AS exam_fee,") \
				_T("        SUM(inpatfee) AS inpatient_fee,") \
				_T("        SUM(examfee+inpatfee) AS total_amount") \
				_T(" FROM") \
				_T(" (") \
				_T("  SELECT CASE WHEN fe.hfe_itemid IN('D0000006') THEN cast('C1.3' as nvarchar2(140))") \
				_T("              WHEN fe.hfe_group IN('D0000') AND fe.hfe_itemid NOT IN('D0000006') THEN cast('C1.1' as nvarchar2(140))") \
				_T("              WHEN substr(fe.hfe_group,1,2) IN('A1') THEN cast('KD' as nvarchar2(140))") \
				_T("              WHEN fe.hfe_group='F0000' THEN cast('X' as nvarchar2(140)) ELSE hfl_deptid END AS deptid,") \
				_T("         hfg_index AS idx,") \
				_T("         fe.hfe_itemid AS itemid,") \
				_T("         fe.hfe_desc AS itemname,") \
				_T("         CASE WHEN fi.hfe_class='E' THEN SUM(fe.hfe_patpaid) ELSE 0 END AS examfee,") \
				_T("         CASE WHEN fi.hfe_class<>'E' THEN SUM(fe.hfe_patpaid) ELSE 0 END AS inpatfee,") \
				_T("         SUM(1) as count") \
				_T("  FROM hms_fee_invoice fi") \
				_T("  LEFT JOIN hms_fee fe ON(fi.hfe_invoiceno=fe.hfe_invoiceno AND fi.hfe_docno=fe.hfe_docno)") \
				_T("  LEFT JOIN hms_fee_list ON(fe.hfe_itemid=hfl_feeid)") \
				_T("  LEFT JOIN hms_fee_group ON(hfg_id=fe.hfe_group)") \
				_T("  WHERE fe.hfe_type IN('E','O','P','T','X','D') AND fi.hfe_status='P' %s") \
				_T("  GROUP BY hfl_deptid, fe.hfe_itemid, fe.hfe_desc, hfg_index, fi.hfe_class, fe.hfe_group") \
				_T("  HAVING SUM(fe.hfe_patpaid) > 0") \
				_T(" ) tbl") \
				_T(" LEFT JOIN sys_dept ON(sd_id=deptid)") \
				_T(" WHERE length(deptid) > 0") \
				_T(" GROUP BY deptid, idx, itemid, itemname, sd_name") \
				_T(" ORDER BY depttype, deptid, idx, itemid, itemname"),
				szWhere);

	return szSQL;
}