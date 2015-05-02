#include "stdafx.h"
#include "FMCollectDetailbyItem.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "ReportDocument.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMCollectDetailbyItem *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMCollectDetailbyItem *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMCollectDetailbyItem *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMCollectDetailbyItem *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMCollectDetailbyItem *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMCollectDetailbyItem *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMCollectDetailbyItem *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMCollectDetailbyItem *)pWnd)->OnToDateCheckValue();
} 
static void _OnClerkSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CFMCollectDetailbyItem* )pWnd)->OnClerkSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnClerkSelendokFnc(CWnd *pWnd){
	((CFMCollectDetailbyItem *)pWnd)->OnClerkSelendok();
}
/*static void _OnClerkSetfocusFnc(CWnd *pWnd){
	((CFMCollectDetailbyItem *)pWnd)->OnClerkKillfocus();
}*/
/*static void _OnClerkKillfocusFnc(CWnd *pWnd){
	((CFMCollectDetailbyItem *)pWnd)->OnClerkKillfocus();
}*/
static long _OnClerkLoadDataFnc(CWnd *pWnd){
	return ((CFMCollectDetailbyItem *)pWnd)->OnClerkLoadData();
}
/*static void _OnClerkAddNewFnc(CWnd *pWnd){
	((CFMCollectDetailbyItem *)pWnd)->OnClerkAddNew();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CFMCollectDetailbyItem *)pWnd)->OnObjectLoadData();
}
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CFMCollectDetailbyItem *pVw = (CFMCollectDetailbyItem *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMCollectDetailbyItem *pVw = (CFMCollectDetailbyItem *)pWnd;
	pVw->OnExportSelect();
}
static void _OnLockedSelectFnc(CWnd *pWnd){
	 ((CFMCollectDetailbyItem*)pWnd)->OnLockedSelect();
}
static int _OnAddFMCollectDetailbyItemFnc(CWnd *pWnd){
	 return ((CFMCollectDetailbyItem*)pWnd)->OnAddFMCollectDetailbyItem();
} 
static int _OnEditFMCollectDetailbyItemFnc(CWnd *pWnd){
	 return ((CFMCollectDetailbyItem*)pWnd)->OnEditFMCollectDetailbyItem();
} 
static int _OnDeleteFMCollectDetailbyItemFnc(CWnd *pWnd){
	 return ((CFMCollectDetailbyItem*)pWnd)->OnDeleteFMCollectDetailbyItem();
} 
static int _OnSaveFMCollectDetailbyItemFnc(CWnd *pWnd){
	 return ((CFMCollectDetailbyItem*)pWnd)->OnSaveFMCollectDetailbyItem();
} 
static int _OnCancelFMCollectDetailbyItemFnc(CWnd *pWnd){
	 return ((CFMCollectDetailbyItem*)pWnd)->OnCancelFMCollectDetailbyItem();
} 
CFMCollectDetailbyItem::CFMCollectDetailbyItem(CWnd *pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 160;
	SetDefaultValues();
}
CFMCollectDetailbyItem::~CFMCollectDetailbyItem()
{
}
void CFMCollectDetailbyItem::OnCreateComponents()
{
	/*m_wndCollectDetailbyItem.Create(this, _T("Collect Detail by Item"), 5, 5, 490, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 485, 85); 
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 300, 95, 410, 120);
	m_wndExport.Create(this, _T("&Export"), 415, 95, 490, 120);*/

	m_wndCollectDetailbyItem.Create(this, _T("Collect Detail by Item"), 5, 5, 430, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 215, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 30, 300, 55);
	m_wndToDate.Create(this,305, 30, 425, 55); 
	m_wndClerkLabel.Create(this, _T("Clerk"), 10, 60, 90, 85);
	m_wndClerk.Create(this,95, 60, 215, 85); 
	m_wndObject.SetCheckBox(true);
	m_wndObjectLabel.Create(this, _T("Object"), 220, 60, 300, 85);
	m_wndObject.Create(this,305, 60, 425, 85); 
	m_wndLocked.Create(this, _T("Locked"), 5, 95, 115, 120);
	m_wndPrintPreview.Create(this, _T("&Print Preview"), 205, 95, 315, 120);
	m_wndExport.Create(this, _T("&Export"), 320, 95, 430, 120);
}
void CFMCollectDetailbyItem::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);
	//m_wndClerk.SetCheckValue(true);
	m_wndClerk.LimitText(65);


	m_wndClerk.InsertColumn(0, _T("ID"), CFMT_TEXT, 80);
	m_wndClerk.InsertColumn(1, _T("Name"), CFMT_TEXT, 270);

	m_wndObject.Format(2, 1, 12);
	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 140);
}
void CFMCollectDetailbyItem::OnSetWindowEvents()
{
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
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndLocked.SetEvent(WE_CLICK, _OnLockedSelectFnc);
	
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);
}
void CFMCollectDetailbyItem::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndClerk.GetDlgCtrlID(), m_szClerkKey);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_Check(pDX, m_wndLocked.GetDlgCtrlID(), m_bLocked);
}
void CFMCollectDetailbyItem::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMCollectDetailbyItem::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMCollectDetailbyItem::SetDefaultValues()
{

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szClerkKey.Empty();
	m_szObjectKey.Empty();
	m_bLocked=FALSE;

}
int CFMCollectDetailbyItem::SetMode(int nMode){
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
/*void CFMCollectDetailbyItem::OnFromDateChange(){
	
} */
/*void CFMCollectDetailbyItem::OnFromDateSetfocus(){
	
} */
/*void CFMCollectDetailbyItem::OnFromDateKillfocus(){
	
} */
int CFMCollectDetailbyItem::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMCollectDetailbyItem::OnToDateChange(){
	
} */
/*void CFMCollectDetailbyItem::OnToDateSetfocus(){
	
} */
/*void CFMCollectDetailbyItem::OnToDateKillfocus(){
	
} */
int CFMCollectDetailbyItem::OnToDateCheckValue(){
	return 0;
} 
void CFMCollectDetailbyItem::OnClerkSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMCollectDetailbyItem::OnClerkSelendok(){
	 
}
/*void CFMCollectDetailbyItem::OnClerkSetfocus(){
	
}*/
/*void CFMCollectDetailbyItem::OnClerkKillfocus(){
	
}*/
long CFMCollectDetailbyItem::OnClerkLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndClerk.IsSearchKey() && !m_szClerkKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(su_userid)=lower('%s') "), m_szClerkKey);
	};
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
/*void CFMCollectDetailbyItem::OnClerkAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
long CFMCollectDetailbyItem::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty())
	{
		szWhere.Format(_T(" and lower(ho_id)=lower('%s') "), m_szObjectKey);
	};
	m_wndObject.DeleteAllItems(); 
	szSQL.Format(_T("SELECT ho_id as id, ho_desc as name FROM hms_object WHERE 1=1 %s ORDER BY CAST(ho_id AS INTEGER)"), szWhere);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
void CFMCollectDetailbyItem::OnPrintPreviewSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szSysDate;

	UpdateData(TRUE);

	int nIdx = 0;

	double nTemp = 0;
	long double nGrpTimes = 0, nGrpExam = 0, nGrpIn = 0, nGrpAmount = 0;
	long double nTotalExam = 0, nTotalIn = 0, nTotalAmount = 0;

	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);

	if (nCount <=0 )
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}

	if (!rpt.Init(_T("Reports/HMS/HF_CHITIETTHUTIENTHEODANHMUC.RPT")))
		return;

	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);

	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")),
		          CDateTime::Convert(m_szFromDate, yyyymmdd | hhmm, ddmmyyyy | hhmm),
				  CDateTime::Convert(m_szToDate, yyyymmdd | hhmm, ddmmyyyy | hhmm));

	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection *rptDetail;
	CReportSection *rptSector;
	CString szOldFeeGrp, szNewFeeGrp;

	while (!rs.IsEOF())
	{
		rs.GetValue(_T("deptid"), szNewFeeGrp);
		if (szNewFeeGrp != szOldFeeGrp)
		{
			if (nGrpTimes > 0)
			{
				rptSector = rpt.AddDetail(rpt.GetGroupFooter(1));
				FormatCurrency(nGrpTimes, tmpStr);
				rptSector->SetValue(_T("s3"), tmpStr);
				FormatCurrency(nGrpExam, tmpStr);
				rptSector->SetValue(_T("s4"), tmpStr);
				FormatCurrency(nGrpIn, tmpStr);
				rptSector->SetValue(_T("s5"), tmpStr);
				FormatCurrency(nGrpAmount, tmpStr);
				rptSector->SetValue(_T("s6"), tmpStr);
			}
			nIdx++;

			if (szNewFeeGrp != _T("X") && szNewFeeGrp != _T("NA"))
				tmpStr.Format(_T("%d. %s"), nIdx, rs.GetValue(_T("deptname")));
			else if (szNewFeeGrp == _T("X"))
				tmpStr.Format(_T("%d. Thu ph\xED kh\xE1\x63"), nIdx);
			else if (szNewFeeGrp == _T("NA"))
				tmpStr.Format(_T("%d. \x43h\x1B0\x61 thi\x1EBFt l\x1EADp kho\x61 th\x1EF1\x63 hi\x1EC7n"), nIdx);

			rptSector = rpt.AddDetail(rpt.GetGroupHeader(0));
			rptSector->SetValue(_T("GroupName"), tmpStr);
			szOldFeeGrp = szNewFeeGrp;
			nGrpTimes = 0;
			nTotalExam += nGrpExam;
			nGrpExam = 0;
			nTotalIn += nGrpIn;
			nGrpIn = 0;
			nTotalAmount += nGrpAmount;
			nGrpAmount = 0;
		}
		rptDetail = rpt.AddDetail();

		rs.GetValue(_T("itemname"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("total"), nTemp);
		//FormatCurrency(nTemp, tmpStr);
		tmpStr.Format(_T("%.0f"), nTemp);
		rptDetail->SetValue(_T("3"), tmpStr);
		nGrpTimes += nTemp;

		rs.GetValue(_T("exam_fee"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		nGrpExam += nTemp;

		rs.GetValue(_T("inpatient_fee"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		nGrpIn += nTemp;

		rs.GetValue(_T("total_amount"), nTemp);
		FormatCurrency(nTemp, tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		nGrpAmount += nTemp;

		rs.MoveNext();
	}
	if (nGrpTimes > 0)
	{
		rptSector = rpt.AddDetail(rpt.GetGroupFooter(1));
		FormatCurrency(nGrpTimes, tmpStr);
		rptSector->SetValue(_T("s3"), tmpStr);
		FormatCurrency(nGrpExam, tmpStr);
		rptSector->SetValue(_T("s4"), tmpStr);
		FormatCurrency(nGrpIn, tmpStr);
		rptSector->SetValue(_T("s5"), tmpStr);
		FormatCurrency(nGrpAmount, tmpStr);
		rptSector->SetValue(_T("s6"), tmpStr);

			//Vertical total
		nTotalExam += nGrpExam;
		nTotalIn += nGrpIn;
		nTotalAmount += nGrpAmount;
	}

	FormatCurrency(nTotalExam, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("t4"), tmpStr);
	FormatCurrency(nTotalIn, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("t5"), tmpStr);
	FormatCurrency(nTotalAmount, tmpStr);
	rpt.GetReportFooter()->SetValue(_T("t6"), tmpStr);
	szSysDate = pMF->GetSysDate();

	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CFMCollectDetailbyItem::OnExportSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0, nIdx = 0;
	double nTemp = 0;
	long double nGrpTimes = 0, nGrpExam = 0, nGrpIn = 0, nGrpAmount = 0;
	long double nTotalExam = 0, nTotalIn = 0, nTotalAmount = 0;

	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
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
				//FormatCurrency(nGrpTimes, tmpStr);
				tmpStr.Format(_T("%.0Lf"), nGrpTimes);
				xls.SetCellText(nCol + 2, nPreviousRow, tmpStr, FMT_INTEGER | FMT_WRAPING, true, 10);

				//FormatCurrency(nGrpExam, tmpStr);
				tmpStr.Format(_T("%.2Lf"), nGrpExam);
				xls.SetCellText(nCol + 3, nPreviousRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);

				//FormatCurrency(nGrpIn, tmpStr);
				tmpStr.Format(_T("%.2Lf"), nGrpIn);
				xls.SetCellText(nCol + 4, nPreviousRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);

				//FormatCurrency(nGrpAmount, tmpStr);
				tmpStr.Format(_T("%.2Lf"), nGrpAmount);
				xls.SetCellText(nCol + 5, nPreviousRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
			}

			nIdx++;
			tmpStr.Format(_T("%d"), nIdx);
			xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING, true, 10);

			if (szNewFeeGrp != _T("X") && szNewFeeGrp != _T("NA"))
			{
				rs.GetValue(_T("deptname"), tmpStr);
			}
			else if (szNewFeeGrp == _T("X"))
			{
				tmpStr = _T("\x43hi ph\xED kh\xE1\x63");
			}
			else if (szNewFeeGrp != _T("NA"))
			{
				tmpStr = _T("\x43h\x1B0\x61 thi\x1EBFt l\x1EADp kho\x61 th\x1EF1\x63 hi\x1EC7n");
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
		//FormatCurrency(nTemp, tmpStr);
		nGrpTimes += nTemp;
		tmpStr.Format(_T("%.0f"), nTemp);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER | FMT_WRAPING);
		

		rs.GetValue(_T("exam_fee"), nTemp);
		//FormatCurrency(nTemp, tmpStr);
		nGrpExam += nTemp;
		tmpStr.Format(_T("%.2f"), nTemp);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);
		

		rs.GetValue(_T("inpatient_fee"), nTemp);
		nGrpIn += nTemp;
		tmpStr.Format(_T("%.2f"), nTemp);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_RIGHT);
		

		rs.GetValue(_T("total_amount"), nTemp);
		//FormatCurrency(nTemp, tmpStr);
		nGrpAmount += nTemp;
		tmpStr.Format(_T("%.2f"), nTemp);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING);

		nRow++;
		rs.MoveNext();
	}

	if (nGrpTimes > 0)
	{
		//FormatCurrency(nGrpTimes, tmpStr);
		tmpStr.Format(_T("%.0Lf"), nGrpTimes);
		xls.SetCellText(nCol + 2, nPreviousRow, tmpStr, FMT_INTEGER | FMT_WRAPING, true, 10);
		//FormatCurrency(nGrpExam, tmpStr);
		tmpStr.Format(_T("%.2Lf"), nGrpExam);
		xls.SetCellText(nCol + 3, nPreviousRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
		//FormatCurrency(nGrpIn, tmpStr);
		tmpStr.Format(_T("%.2Lf"), nGrpIn);
		xls.SetCellText(nCol + 4, nPreviousRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
		//FormatCurrency(nGrpAmount, tmpStr);
		tmpStr.Format(_T("%.2Lf"), nGrpAmount);
		xls.SetCellText(nCol + 5, nPreviousRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);

		nTotalExam += nGrpExam;
		nTotalIn += nGrpIn;
		nTotalAmount += nGrpAmount;
	}

	TranslateString(_T("Total"), tmpStr);
	xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT, true, 10);

	//FormatCurrency(nTotalExam, tmpStr);
	tmpStr.Format(_T("%.2Lf"), nTotalExam);
	xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
	//FormatCurrency(nTotalIn, tmpStr);
	tmpStr.Format(_T("%.2Lf"), nTotalIn);
	xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);
	//FormatCurrency(nTotalAmount, tmpStr);
	tmpStr.Format(_T("%.2Lf"), nTotalAmount);
	xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_NUMBER1 | FMT_WRAPING, true, 10);

	xls.Save(_T("Exports\\Chi Tiet Thu Tien Theo Danh Muc.xls"));
}
void CFMCollectDetailbyItem::OnLockedSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
int CFMCollectDetailbyItem::OnAddFMCollectDetailbyItem(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMCollectDetailbyItem::OnEditFMCollectDetailbyItem(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMCollectDetailbyItem::OnDeleteFMCollectDetailbyItem(){
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
 		OnCancelFMCollectDetailbyItem(); 
 	}
	return 0;
}
int CFMCollectDetailbyItem::OnSaveFMCollectDetailbyItem(){
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
 		//OnFMCollectDetailbyItemListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMCollectDetailbyItem::OnCancelFMCollectDetailbyItem(){
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
int CFMCollectDetailbyItem::OnFMCollectDetailbyItemListLoadData(){
	return 0;
}

CString CFMCollectDetailbyItem::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere, szObject;
	szWhere.Empty();
	for (int i = 0; i < m_wndObject.GetItemCount(); i++)
	{
		if (m_wndObject.GetCheck(i))
		{
			m_wndObject.SetCurSel(i);
			if (!szObject.IsEmpty())
				szObject.AppendFormat(_T(", "));
			szObject.AppendFormat(_T("%s"), m_wndObject.GetCurrent(0));
		}
	}
	if (!m_szClerkKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_staff='%s' "), m_szClerkKey);
	if (!szObject.IsEmpty())
		szWhere.AppendFormat(_T(" AND hfe_objectid IN (%s) "), szObject);
	if (!m_bLocked)
	{
		szWhere.AppendFormat(_T(" AND hfe_locked<>'Y' "));
		szWhere.AppendFormat(_T(" AND hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}
	else
	{
		szWhere.AppendFormat(_T(" AND hfe_locked='Y' "));
		szWhere.AppendFormat(_T(" AND hfe_date BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s') "),
			                 m_szFromDate, m_szToDate);
	}
	szSQL.Format(_T(" SELECT    hfe_depttype depttype, ") \
				_T("           hfe_deptid deptid, ") \
				_T("           hfg_index                     AS idx, ") \
				_T("           sd_name                       AS deptname, ") \
				_T("           hfe_desc itemname, ") \
				_T("           SUM(count_patient)            AS total, ") \
				_T("           SUM(exam_fee)                 AS exam_fee, ") \
				_T("           SUM(inpatient_fee)            AS inpatient_fee, ") \
				_T("           SUM(exam_fee + inpatient_fee) AS total_amount ") \
				_T(" FROM      (SELECT    f.hfe_docno, ") \
				_T("                      f.hfe_itemid, ") \
				_T("                      f.hfe_desc, ") \
				_T("                      CASE WHEN f.hfe_deptid IN( 'C1.1', 'C1.2', 'C1.3' ) THEN Cast('KB' AS NVARCHAR2(7)) ") \
				_T("                        WHEN f.hfe_group = 'F0000' THEN Cast('XX' AS NVARCHAR2(7)) ") \
				_T("                        WHEN f.hfe_deptid IN( 'KD' ) THEN Cast('KD' AS NVARCHAR2(7)) ") \
				_T("                        WHEN f.hfe_deptid IN( 'C10' ) THEN Cast('KTB' AS NVARCHAR2(7)) ") \
				_T("                      ELSE Cast('DT' AS NVARCHAR2(7)) ") \
				_T("                      END AS hfe_depttype, ") \
				_T("                      f.hfe_deptid, ") \
				_T("                      f.hfe_class, ") \
				_T("                      f.hfe_type, ") \
				_T("                      f.hfe_group, ") \
				_T("                      CASE WHEN i.hfe_class IN( 'E', 'X' ) THEN f.hfe_patpaid ") \
				_T("                      ELSE 0 ") \
				_T("                      END AS exam_fee, ") \
				_T("                      CASE WHEN i.hfe_class NOT IN( 'E', 'X' ) THEN f.hfe_patpaid ") \
				_T("                      ELSE 0 ") \
				_T("                      END AS inpatient_fee, ") \
				_T("                      CASE WHEN f.hfe_group = 'C0000' THEN f.hfe_quantity ") \
				_T("                      ELSE 1 ") \
				_T("                      END AS count_patient ") \
				_T("            FROM      hms_fee_invoice i ") \
				_T("            LEFT JOIN (SELECT    hfe_invoiceno, ") \
				_T("                                 hfe_docno, ") \
				_T("                                 hfe_class, ") \
				_T("                                 hfe_type, ") \
				_T("                                 hfe_group, ") \
				_T("                                 CASE WHEN hfe_type IN ('P', 'T') THEN hfl_deptid ") \
				_T("								 ELSE CASE WHEN hfe_type = 'O' THEN ho_pdeptid ELSE hfe_deptid ") \
				_T("                                 END END hfe_deptid, ") \
				_T("                                 hfe_itemid, ") \
				_T("                                 hfe_desc, ") \
				_T("                                 hfe_quantity, ") \
				_T("                                 hfe_patpaid ") \
				_T("                       FROM      hms_fee ") \
				_T("					   LEFT JOIN hms_operation ON (hfe_type = 'O' AND hfe_docno = ho_docno AND hfe_orderid = ho_idx)") \
				_T("                       LEFT JOIN hms_fee_list ON( hfl_feeid = hfe_itemid ) ") \
				_T("                       WHERE     hfe_type IN( 'E', 'V', 'O', 'P', ") \
				_T("                                              'T', 'X', 'F' ) ") \
				_T("                       UNION ALL ") \
				_T("                       SELECT hfe_invoiceno, ") \
				_T("                              hfe_docno, ") \
				_T("                              hfe_class, ") \
				_T("                              hfe_type, ") \
				_T("                              hfe_group, ") \
				_T("                              hfe_deptid, ") \
				_T("                              hfe_itemid, ") \
				_T("                              hfe_desc, ") \
				_T("                              hfe_quantity, ") \
				_T("                              hfe_patpaid ") \
				_T("                       FROM   hms_fee ") \
				_T("                       WHERE  hfe_type IN( 'B', 'J' ) ") \
				_T("                       UNION ALL ") \
				_T("                       SELECT    hfe_invoiceno, ") \
				_T("                                 hfe_docno, ") \
				_T("                                 hfe_class, ") \
				_T("                                 hfe_type, ") \
				_T("                                 hfe_group, ") \
				_T("                                 CASE WHEN mp_org_id = 'PM' THEN Cast('KD' AS NVARCHAR2(7)) ") \
				_T("                                   WHEN mp_org_id = 'MA' THEN Cast('C10' AS NVARCHAR2(7)) ") \
				_T("                                 ELSE hfe_deptid ") \
				_T("                                 END AS hfe_deptid, ") \
				_T("                                 hfe_itemid, ") \
				_T("                                 hfe_desc, ") \
				_T("                                 hfe_quantity, ") \
				_T("                                 hfe_patpaid ") \
				_T("                       FROM      hms_fee ") \
				_T("                       LEFT JOIN hms_pharmaretailorder_view ON( hfe_docno = hpo_docno ") \
				_T("                                                      AND hfe_orderid = hpo_orderid ) ") \
				_T("                       LEFT JOIN m_product_item ON( Cast(mpi_product_item_id AS NVARCHAR2(15)) = hfe_itemid ) ") \
				_T("                       LEFT JOIN m_product ON( mp_product_id = mpi_product_id ) ") \
				_T("                       WHERE     hfe_type IN( 'R', 'D', 'M' ) AND hpo_orderid > 0) f ON( f.hfe_docno = i.hfe_docno ") \
				_T("                                                                     AND f.hfe_invoiceno = i.hfe_invoiceno ) ") \
				_T("            WHERE     f.hfe_patpaid > 0 AND hfe_payment > 0 %s") \
				_T("                  AND i.hfe_status = 'P') tbl ") \
				_T(" LEFT JOIN hms_fee_group ON( hfg_id = tbl.hfe_group ) ") \
				_T(" LEFT JOIN sys_dept ON( sd_id = hfe_deptid ) ") \
				_T(" GROUP     BY hfe_depttype, ") \
				_T("              hfe_deptid, ") \
				_T("              hfg_index, ") \
				_T("              sd_name, ") \
				_T("              hfe_desc ") \
				_T(" ORDER     BY hfe_depttype, ") \
				_T("              hfe_deptid, ") \
				_T("              hfg_index, ") \
				_T("              hfe_desc "), szWhere);_fmsg(_T("%s"), szSQL);

	return szSQL;
}