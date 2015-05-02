#include "stdafx.h"
#include "EMExaminationListReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMExaminationListReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMExaminationListReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMExaminationListReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExaminationListReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMExaminationListReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMExaminationListReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMExaminationListReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExaminationListReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMExaminationListReport *pVw = (CEMExaminationListReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMExaminationListReport *pVw = (CEMExaminationListReport *)pWnd;
	pVw->OnExportSelect();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMExaminationListReport* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CEMExaminationListReport *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CEMExaminationListReport *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CEMExaminationListReport *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CEMExaminationListReport *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CEMExaminationListReport *)pWnd)->OnObjectAddNew();
}*/
static long _OnRoomLoadDataFnc(CWnd *pWnd){
	return ((CEMExaminationListReport *)pWnd)->OnRoomLoadData();
}
static int _OnAddCEMExaminationListReportFnc(CWnd *pWnd){
	 return ((CEMExaminationListReport*)pWnd)->OnAddCEMExaminationListReport();
} 
static int _OnEditCEMExaminationListReportFnc(CWnd *pWnd){
	 return ((CEMExaminationListReport*)pWnd)->OnEditCEMExaminationListReport();
} 
static int _OnDeleteCEMExaminationListReportFnc(CWnd *pWnd){
	 return ((CEMExaminationListReport*)pWnd)->OnDeleteCEMExaminationListReport();
} 
static int _OnSaveCEMExaminationListReportFnc(CWnd *pWnd){
	 return ((CEMExaminationListReport*)pWnd)->OnSaveCEMExaminationListReport();
} 
static int _OnCancelCEMExaminationListReportFnc(CWnd *pWnd){
	 return ((CEMExaminationListReport*)pWnd)->OnCancelCEMExaminationListReport();
} 
CEMExaminationListReport::CEMExaminationListReport(CWnd *pParent)
{
	m_nDlgWidth = 500;
	m_nDlgHeight = 135;
	SetDefaultValues();
}
CEMExaminationListReport::~CEMExaminationListReport()
{
}
void CEMExaminationListReport::OnCreateComponents()
{
	m_wndExaminationListReport.Create(this, _T("Examination List Report"), 5, 5, 490, 120);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 60, 90, 85);
	m_wndObject.SetCheckBox(true);
	m_wndObject.Create(this,95, 60, 485, 85); 
	m_wndRoomLabel.Create(this, _T("Room"), 10, 90, 90, 115);
	m_wndRoom.Create(this,95, 90, 485, 115); 
	m_wndPrint.Create(this, _T("&Print"), 325, 125, 405, 150);
	m_wndExport.Create(this, _T("&Export"), 410, 125, 490, 150);

}
void CEMExaminationListReport::OnInitializeComponents()
{
	CHMSMainFrame  *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
//	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(35);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndObject.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);

	m_wndRoom.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndRoom.InsertColumn(1, _T("Name"), CFMT_TEXT, 250);
}
void CEMExaminationListReport::OnSetWindowEvents()
{
	CHMSMainFrame  *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	m_wndRoom.SetEvent(WE_LOADDATA, _OnRoomLoadDataFnc);
	//AddEvent(1, _T("Add	Ctrl+A"), _OnAddCEMExaminationListReportFnc, 0, 'A', VK_CONTROL);
	//AddEvent(2, _T("Edit	Ctrl+E"), _OnEditCEMExaminationListReportFnc, 0, 'E', VK_CONTROL);
	//AddEvent(3, _T("Delete	Ctrl+D"), _OnDeleteCEMExaminationListReportFnc, 0, 'D', VK_CONTROL);
	//AddEvent(4, _T("Save	Ctrl+S"), _OnSaveCEMExaminationListReportFnc, 0, 'S', VK_CONTROL);
	//AddEvent(0, _T("-"));
	//AddEvent(5, _T("Cancel	Ctrl+T"), _OnCancelCEMExaminationListReportFnc, 0, 'T', VK_CONTROL);

	CString szSysDate = pMF->GetSysDate();
	m_szFromDate = m_szToDate = szSysDate;
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
//	UpdateData(false);
	SetMode(VM_EDIT);


}
void CEMExaminationListReport::OnDoDataExchange(CDataExchange* pDX)
{
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_TextEx(pDX, m_wndRoom.GetDlgCtrlID(), m_szRoomKey);

}
void CEMExaminationListReport::GetDataToScreen(){
	CHMSMainFrame  *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMExaminationListReport::GetScreenToData(){
	CHMSMainFrame  *pMF = (CHMSMainFrame *) AfxGetMainWnd();

}
void CEMExaminationListReport::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szObjectKey.Empty();

}
int CEMExaminationListReport::SetMode(int nMode){
 		int nOldMode = GetMode();
 		CGuiView::SetMode(nMode);
 		CHMSMainFrame  *pMF = (CHMSMainFrame  *) AfxGetMainWnd();
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
 			EnableButtons(TRUE, 0, 1,2, -1);
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
/*void CEMExaminationListReport::OnFromDateChange(){
	
} */
/*void CEMExaminationListReport::OnFromDateSetfocus(){
	
} */
/*void CEMExaminationListReport::OnFromDateKillfocus(){
	
} */
int CEMExaminationListReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMExaminationListReport::OnToDateChange(){
	
} */
/*void CEMExaminationListReport::OnToDateSetfocus(){
	
} */
/*void CEMExaminationListReport::OnToDateKillfocus(){
	
} */
int CEMExaminationListReport::OnToDateCheckValue(){
	return 0;
} 

long CEMExaminationListReport::OnRoomLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT hrl_id id, hrl_name name FROM hms_roomlist WHERE hrl_deptid = '%s' ORDER BY hrl_id"), pMF->GetCurrentDepartmentID());
	int nCount = rs.ExecSQL(szSQL);
	m_wndRoom.DeleteAllItems();
	while (!rs.IsEOF())
	{
		m_wndRoom.AddItems(
			rs.GetValue(_T("id")),
			rs.GetValue(_T("name")),
			NULL);
		rs.MoveNext();
	}
	return nCount;
}

void CEMExaminationListReport::OnPrintSelect(){
	CHMSMainFrame  *pMF = (CHMSMainFrame *) AfxGetMainWnd();

	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szSysDate;
	int nIdx = 0, nDrug = 0, nInward = 0;
	double nAmount = 0, nTotal = 0;
	szSQL = GetQueryString();
	//QueryOpener(szSQL);
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	if (!rpt.Init(_T("Reports/HMS/HE_DANHSACHBENHNHANKHAMBENH.rpt")))
		return;
	int nRes = rs.GetRecordCount();
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	FormatCurrency(nRes, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("TotalPatient"), tmpStr);
	CReportSection *rptDetail;
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);
		
		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		
		rs.GetValue(_T("pname"), tmpStr);	
		rptDetail->SetValue(_T("3"), tmpStr);
		
		rs.GetValue(_T("yob"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		
		rs.GetValue(_T("rank"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		
		rs.GetValue(_T("workplace"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		
		rs.GetValue(_T("diagno"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		
		rs.GetValue(_T("drugdeliver"), tmpStr);
		//_debug(_T("drug:%s"), tmpStr);
		if (!tmpStr.IsEmpty())
			nDrug++;
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.GetValue(_T("inward"), tmpStr);
		//_debug(_T("in:%s"), tmpStr);
		if (!tmpStr.IsEmpty())
			nInward++;
		rptDetail->SetValue(_T("9"), tmpStr);
		rs.MoveNext();
	}
	tmpStr.Format(_T("%d"), nDrug);
	rpt.GetReportHeader()->SetValue(_T("Drug"), tmpStr);
	tmpStr.Format(_T("%d"), nInward);
	rpt.GetReportHeader()->SetValue(_T("Inward"), tmpStr);
	szSysDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szSysDate.Right(2), szSysDate.Mid(5, 2), szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CEMExaminationListReport::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame  *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nIdx = 0, nCol = 0, nRow = 0;
	int c1 = 0, c2 = 0, c3 = 0, c4 = 0, c5 = 0, c6 =0;
	int c7 = 0, c8 = 0, c9 = 0, c10 = 0, c11 = 0;
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 4);
	xls.SetColumnWidth(1, 25);
	xls.SetColumnWidth(3, 15);
	xls.SetColumnWidth(4, 30);
	xls.SetColumnWidth(5, 20);

	xls.SetColumnWidth(6, 20);
	xls.SetColumnWidth(7, 20);
	xls.SetColumnWidth(8, 20);
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 9);
	xls.SetCellMergedColumns(nCol, nRow + 3, 9);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 2, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N KH\xC1M \x42\x1EC6NH"), FMT_TEXT | FMT_CENTER, true, 11);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, FMT_TEXT | FMT_CENTER, false, 11);	
	
	
	
	CStringArray arrCol;
	arrCol.Add(_T("STT"));
	arrCol.Add(_T("T\xEAn \x62\x1EC7nh nh\xE2n"));
	arrCol.Add(_T("N\x103m sinh"));
	arrCol.Add(_T("\x43\x1EA5p \x62\x1EAD\x63"));
	arrCol.Add(_T("\x110\x1A1n v\x1ECB"));
	arrCol.Add(_T("\x43h\x1EA9n \x111o\xE1n"));
	arrCol.Add(_T("Ph\xE1t thu\x1ED1\x63"));
	arrCol.Add(_T("V\xE0o kho\x61"));
	arrCol.Add(_T("\x110\x1ED1i t\x1B0\x1EE3ng"));
	arrCol.Add(_T("S\x1ED1 th\x1EBB"));

	nRow = 7;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10); 
	}
	nRow = 8;
	while (!rs.IsEOF())
	{
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_INTEGER);
		
		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol + 1, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("yob"), tmpStr);
		xls.SetCellText(nCol + 2, nRow, tmpStr, FMT_INTEGER);
		
		rs.GetValue(_T("rank"), tmpStr);
		xls.SetCellText(nCol + 3, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("workplace"), tmpStr);
		xls.SetCellText(nCol + 4, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("diagno"), tmpStr);
		xls.SetCellText(nCol + 5, nRow, tmpStr, FMT_TEXT);
		
		rs.GetValue(_T("drugdeliver"), tmpStr);
		if(!tmpStr.IsEmpty()) c4++;
		xls.SetCellText(nCol + 6, nRow, tmpStr, FMT_TEXT | FMT_CENTER);

		rs.GetValue(_T("drugquan"), tmpStr);
		if(!tmpStr.IsEmpty()) c5++;

		rs.GetValue(_T("drugbhytquan"), tmpStr);
		if(!tmpStr.IsEmpty()) c6++;
		
		rs.GetValue(_T("drugbhytquancothe"), tmpStr);
		if(!tmpStr.IsEmpty()) c7++;

		rs.GetValue(_T("inward"), tmpStr);
		if(!tmpStr.IsEmpty()) c8++;
		xls.SetCellText(nCol + 7, nRow, tmpStr, FMT_TEXT);

		rs.GetValue(_T("inwardquan"), tmpStr);
		if(!tmpStr.IsEmpty()) c9++;

		rs.GetValue(_T("inwardbhytquan"), tmpStr);
		if(!tmpStr.IsEmpty()) c10++;

		rs.GetValue(_T("inwardquancothe"), tmpStr);
		if(!tmpStr.IsEmpty()) c11++;


		rs.GetValue(_T("hd_object"), tmpStr);
		if(tmpStr == _T("1")) 
		{
			c1++;
		}
		else if(tmpStr == _T("2")) 
		{
			c2++;
		}
		else if(tmpStr == _T("11"))
		{
			c3++;
		}


		rs.GetValue(_T("obj"), tmpStr);
		xls.SetCellText(nCol + 8, nRow,tmpStr , FMT_TEXT);
		
		rs.GetValue(_T("card_no"), tmpStr);
		xls.SetCellText(nCol + 9, nRow,tmpStr , FMT_TEXT);
		nRow++;
		rs.MoveNext();
	}

	CString szTemp;

	szTemp.Format(_T("Qu\xE2n: %d"),c1++);
	xls.SetCellText(1, 4, szTemp, FMT_TEXT, true, 12);

	xls.SetCellMergedColumns(2, 4, 2);
	szTemp.Format(_T("H\x1B0u: %d"),c2++);
	xls.SetCellText(2, 4, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("\x42HYT Qu\xE2n nh\xE2n: %d"),c3++);
	xls.SetCellText(4, 4, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("T\x1ED5ng s\x1ED1: %d"),nIdx++);
	xls.SetCellText(5, 4, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("Ph\xE1t thu\x1ED1\x63 qu\xE2n: %d"),c5++);
	xls.SetCellText(1, 5, szTemp, FMT_TEXT, true, 12);

	xls.SetCellMergedColumns(2, 5, 2);
	szTemp.Format(_T("Ph\xE1t thu\x1ED1\x63 h\x1B0u: %d"),c6++);
	xls.SetCellText(2, 5, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("Ph\xE1t thu\x1ED1\x63 Q\x43T(\x31\x31): %d"),c7++);
	xls.SetCellText(4, 5, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("T\x1ED5ng ph\xE1t thu\x1ED1\x63: %d"),c4++);
	xls.SetCellText(5, 5, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("V\xE0o vi\x1EC7n qu\xE2n: %d"),c9++);
	xls.SetCellText(1, 6, szTemp, FMT_TEXT, true, 12);

	xls.SetCellMergedColumns(2, 6, 2);
	szTemp.Format(_T("V\xE0o vi\x1EC7n h\x1B0u: %d"),c10++);
	xls.SetCellText(2, 6, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("V\xE0o vi\x1EC7n Q\x43T(\x31\x31): %d"),c11++);
	xls.SetCellText(4, 6, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("T\x1ED5ng v\xE0o vi\x1EC7n: %d"),c8++);
	xls.SetCellText(5, 6, szTemp, FMT_TEXT, true, 12);


	xls.Save(_T("Exports\\Danh Sach Benh Nhan Kham Benh.xls"));

	
} 
void CEMExaminationListReport::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame  *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	
} 
void CEMExaminationListReport::OnObjectSelendok(){
	 
}
/*void CEMExaminationListReport::OnObjectSetfocus()
{
	
}*/
/*void CEMExaminationListReport::OnObjectKillfocus()
{
	
}*/
long CEMExaminationListReport::OnObjectLoadData()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty())
	{
			szWhere.Format(_T(" where ho_id='%s'"), m_szObjectKey);
	};
			m_wndObject.DeleteAllItems(); 
			szSQL.Format(_T(" select ho_id as id, ho_desc as name from hms_object %s order by cast(ho_id as integer)"), szWhere);
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF())
	{ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("name")), NULL);
		rs.MoveNext();
	}
	return nCount;

}
/*void CEMExaminationListReport::OnObjectAddNew(){
	CHMSMainFrame  *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	
} */
int CEMExaminationListReport::OnAddCEMExaminationListReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame  *pMF = (CHMSMainFrame  *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CEMExaminationListReport::OnEditCEMExaminationListReport(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame  *pMF = (CHMSMainFrame  *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMExaminationListReport::OnDeleteCEMExaminationListReport(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame  *pMF = (CHMSMainFrame  *)AfxGetMainWnd();
 	CString szSQL;
 	if(ShowMessage(1, MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2) == IDNO)
 		return -1;
 	szSQL.Format(_T("DELETE FROM  WHERE  AND") );
 	int ret = pMF->ExecSQL(szSQL);
 	if(ret >= 0){
 		SetMode(VM_NONE);
 		OnCancelCEMExaminationListReport();
 	}
	return 0;
}
int CEMExaminationListReport::OnSaveCEMExaminationListReport(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT)
 		return -1;
 	if(!IsValidateData())
 		return -1;
 	CHMSMainFrame  *pMF = (CHMSMainFrame  *)AfxGetMainWnd();
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
 		//OnCEMExaminationListReportListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CEMExaminationListReport::OnCancelCEMExaminationListReport(){
 	if(GetMode() == VM_EDIT)
 	{
 		SetMode(VM_VIEW);
 	} 
 	else 
 	{
 		SetMode(VM_NONE);
 	} 
 	CHMSMainFrame  *pMF = (CHMSMainFrame  *)AfxGetMainWnd();
 	return 0;
} 
int CEMExaminationListReport::OnCEMExaminationListReportListLoadData()
{
	return 0;
}

CString CEMExaminationListReport::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	CString szSQL, szWhere, szStatus;
	CString szObjectID, szObjectName;
	
	/*if (m_nSoldier == 0)
		szWhere.AppendFormat(_T(" AND hd_object = '1'"));
	else
		szWhere.AppendFormat(_T(" AND hd_object = '2'"));
	szStatus.Format(_T(" AND hpo_orderstatus <> 'O'"));*/
	//if (m_bOnlyOrder)
	//	szStatus.Format(_T(" AND hpo_orderstatus IN('S', 'A')"));

	szWhere.Format(_T(" WHERE he_examdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	for (int i=0 ; i<= m_wndObject.GetCount(); i++)
	 {
		 if(m_wndObject.GetCheck(i))
		 {
			 m_wndObject.SetCurSel(i);
		 	if(!szObjectID.IsEmpty())
					szObjectID += _T(",");						
					szObjectID.AppendFormat(_T("'%s'"), m_wndObject.GetCheck(i,0));			
			if(!szObjectName.IsEmpty())
					szObjectName += _T(", ");						
			szObjectName.AppendFormat(_T("%s"), m_wndObject.GetCheck(i, 0));
		 }
	 }
	
	 if(!m_szObjectKey.IsEmpty())	
		szWhere.AppendFormat(_T(" and hd_object in(%s)"), szObjectID);	
	else
		szObjectName.Format(_T("T\x1EA5t \x63\x1EA3 \x63\xE1\x63 \x111\x1ED1i t\x1B0\x1EE3ng"));
	 if (!m_szRoomKey.IsEmpty())
		 szWhere.AppendFormat(_T(" AND he_roomid = %s"), m_szRoomKey);

	/*szSQL.Format(_T(" SELECT  distinct he_docno as docno,") \
				_T("         trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
				_T("         extract(year from hp_birthdate) as yob, get_objectname(hd_object) as obj,hd_object, ") \
				_T("         (select distinct ss_desc from sys_sel where ss_id = 'hms_rank' and ss_code = hp_rank) as rank,") \
				_T("         hp_workplace as workplace,") \
				_T("         hd_diagnostic as diagno,") \
				_T("         case when hd_suggestion = 'D' and length(hd_indept) > 0 then hd_indept end as inward,") \
				_T("         case when hpo_orderid IS NOT NULL AND hpo_orderstatus IN('S', 'A') AND hd_status = 'T' then cast('X' as NVARCHAR2(1)) end as drugdeliver") \
				_T(" FROM hms_patient ") \
				_T(" LEFT JOIN hms_doc ON (hd_patientno = hp_patientno)") \
				_T(" LEFT JOIN hms_exam ON (he_docno = hd_docno)") \
				_T(" LEFT JOIN hms_pharmaorder ON (hpo_docno = he_docno)") \
				_T(" %s AND he_deptid = '%s' order by docno"),szWhere,  pMF->m_szDept);*/


	szSQL.Format(_T("  SELECT  distinct hd_docno as docno,    ") \
				_T("           trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,   ") \
				_T("           extract(year from hp_birthdate) as yob, get_objectname(hd_object) as obj,hd_object,    ") \
				_T("           get_syssel_desc('hms_rank', hp_rank) as rank, ") \
				_T("           hp_workplace as workplace,        ") \
				_T("           hd_diagnostic as diagno,     ") \
				_T("		   hd_cardno card_no,") \
				_T("           case when hcr_numinward > 0 and hd_object in ('1', '2', '11') then hcr_admitdept end as inward,   ") \
				_T("           case when hcr_numinward > 0 and hd_object in('1') then hcr_admitdept end as inwardquan,   ") \
				_T("           case when hcr_numinward > 0 and hd_object in('2') then hcr_admitdept end as inwardbhytquan,   ") \
				_T("           case when hcr_numinward > 0 and hd_object in ('11') then hcr_admitdept end as inwardquancothe,   ") \
				_T("           case when hpo_orderid IS NOT NULL then cast('X' as NVARCHAR2(1)) end as drugdeliver , ") \
				_T("           case when hpo_orderid IS NOT NULL and hd_object in ('1') then cast('X' as NVARCHAR2(1)) end as drugquan, ") \
				_T("           case when hpo_orderid IS NOT NULL and hd_object in ('2') then cast('X' as NVARCHAR2(1)) end as drugbhytquan,") \
				_T("           case when hpo_orderid IS NOT NULL and hd_object in ('11') then cast('X' as NVARCHAR2(1)) end as drugbhytquancothe") \
				_T("   FROM hms_patient ") \
				_T("   LEFT JOIN hms_doc ON (hd_patientno = hp_patientno) ") \
				_T("   LEFT JOIN hms_exam ON (hd_docno = he_docno)") \
				_T("   LEFT JOIN hms_clinical_record ON (hcr_docno = hd_docno)") \
				_T("   LEFT JOIN hms_pharmaorder ON (hpo_docno = hd_docno AND hpo_orderstatus IN ('S', 'A')) ") \
				_T("  %s AND he_deptid = '%s' AND he_status IN ('T', 'P')") \
				_T("         order by docno"), szWhere,  pMF->GetCurrentDepartmentID());
	//QueryOpener(szSQL);
	return szSQL;
}

