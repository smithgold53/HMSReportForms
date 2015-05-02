#include "stdafx.h"
#include "EMExamListReportC13.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMExamListReportC13 *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMExamListReportC13 *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMExamListReportC13 *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExamListReportC13 *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMExamListReportC13 *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMExamListReportC13 *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMExamListReportC13 *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMExamListReportC13 *)pWnd)->OnToDateCheckValue();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMExamListReportC13* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CEMExamListReportC13 *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CEMExamListReportC13 *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CEMExamListReportC13 *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CEMExamListReportC13 *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CEMExamListReportC13 *)pWnd)->OnObjectAddNew();
}*/

static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMExamListReportC13 *pVw = (CEMExamListReportC13 *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMExamListReportC13 *pVw = (CEMExamListReportC13 *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMExamListReportC13*)pWnd)->OnAddEMInwardListReport();
} 
static int _OnEditEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMExamListReportC13*)pWnd)->OnEditEMInwardListReport();
} 
static int _OnDeleteEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMExamListReportC13*)pWnd)->OnDeleteEMInwardListReport();
} 
static int _OnSaveEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMExamListReportC13*)pWnd)->OnSaveEMInwardListReport();
} 
static int _OnCancelEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMExamListReportC13*)pWnd)->OnCancelEMInwardListReport();
} 
CEMExamListReportC13::CEMExamListReportC13(CWnd *pParent){

	/*m_nDlgWidth = 500;
	m_nDlgHeight = 130;
	SetDefaultValues();*/

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CEMExamListReportC13::~CEMExamListReportC13(){
}
void CEMExamListReportC13::OnCreateComponents()
{
	/*m_wndInwardListReport.Create(this, _T("Inward List Report"), 5, 5, 490, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 60, 90, 85);
	m_wndObject.Create(this,95, 60, 245, 85); 
	m_wndSectionLabel.Create(this, _T("Section"), 250, 60, 330, 85);
	m_wndSection.Create(this,335, 60, 485, 85); 
	m_wndPrint.Create(this, _T("&Print"), 335, 95, 410, 120);
	m_wndExport.Create(this, _T("&Export"), 415, 95, 490, 120);*/

	m_wndInwardListReport.Create(this, _T("Inward List Report"), 5, 5, 490, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 60, 90, 85);
	m_wndObject.Create(this,95, 60, 485, 85); 
	m_wndPrint.Create(this, _T("&Print"), 335, 95, 410, 120);
	m_wndExport.Create(this, _T("&Export"), 415, 95, 490, 120);
}
void CEMExamListReportC13::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	//m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(35);

	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndObject.InsertColumn(1, _T("Description"), CFMT_TEXT, 250);
}
void CEMExamListReportC13::OnSetWindowEvents(){
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndObject.SetEvent(WE_SELENDOK, _OnObjectSelendokFnc);
	//m_wndObject.SetEvent(WE_SETFOCUS, _OnObjectSetfocusFnc);
	//m_wndObject.SetEvent(WE_KILLFOCUS, _OnObjectKillfocusFnc);
	m_wndObject.SetEvent(WE_SELCHANGE, _OnObjectSelectChangeFnc);
	m_wndObject.SetEvent(WE_LOADDATA, _OnObjectLoadDataFnc);
	//m_wndObject.SetEvent(WE_ADDNEW, _OnObjectAddNewFnc);
	
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	AddEvent(1, _T("Add	Ctrl+A"), _OnAddEMInwardListReportFnc, 0, 'A', VK_CONTROL);
	AddEvent(2, _T("Edit	Ctrl+E"), _OnEditEMInwardListReportFnc, 0, 'E', VK_CONTROL);
	AddEvent(3, _T("Delete	Ctrl+D"), _OnDeleteEMInwardListReportFnc, 0, 'D', VK_CONTROL);
	AddEvent(4, _T("Save	Ctrl+S"), _OnSaveEMInwardListReportFnc, 0, 'S', VK_CONTROL);
	AddEvent(0, _T("-"));
	AddEvent(5, _T("Cancel	Ctrl+T"), _OnCancelEMInwardListReportFnc, 0, 'T', VK_CONTROL);
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CEMExamListReportC13::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
}
void CEMExamListReportC13::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMExamListReportC13::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMExamListReportC13::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szObjectKey.Empty();
	

}
int CEMExamListReportC13::SetMode(int nMode){
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
/*void CEMExamListReportC13::OnFromDateChange(){
	
} */
/*void CEMExamListReportC13::OnFromDateSetfocus(){
	
} */
/*void CEMExamListReportC13::OnFromDateKillfocus(){
	
} */
int CEMExamListReportC13::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMExamListReportC13::OnToDateChange(){
	
} */
/*void CEMExamListReportC13::OnToDateSetfocus(){
	
} */
/*void CEMExamListReportC13::OnToDateKillfocus(){
	
} */
int CEMExamListReportC13::OnToDateCheckValue(){
	return 0;
} 
void CEMExamListReportC13::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMExamListReportC13::OnObjectSelendok(){
	 
}
/*void CEMExamListReportC13::OnObjectSetfocus(){
	
}*/
/*void CEMExamListReportC13::OnObjectKillfocus(){
	
}*/
long CEMExamListReportC13::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty()){
		//szWhere.Format();
	};
	szSQL.Format(_T("SELECT ho_id as id, ho_desc as description FROM hms_object ORDER BY cast(id as integer)"));
	m_wndObject.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndObject.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMExamListReportC13::OnObjectAddNew()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
} */

void CEMExamListReportC13::OnPrintSelect()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nCount = 0, ret = 0;
	int c1 = 0, c2= 0;
	szSQL = GetQueryString();
	//QueryOpener(szSQL);
	if (!rpt.Init(_T("Reports\\HMS\\HE_DANHSACHBENHNHANKHAMBENHC13.RPT")))
		return;
	ret = rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
	nCount = rs.GetRecordCount();
	if (ret <=0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	rpt.GetReportHeader()->SetValue(_T("HealthService"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection *rptDetail = NULL;
	int nIdx=0;
	while (!rs.IsEOF())
	{
		//STT
		//So hs
		//hoten
		//nam sinh
		//capbac
		//doi tuong
		//ngaykham
		//baohiem
		//chandoan
		//vao vien
		// ve nha
		//tuyen
		//docno, he_examdate, pname, yob, rank, workplace, diagno, inward, drugdeliver, tuye, cardno, 
		rptDetail = rpt.AddDetail();
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);
		//sovaovien
		rs.GetValue(_T("docno"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		//tenbn
		rs.GetValue(_T("pname"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		//namsinh
		rs.GetValue(_T("yob"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		//capbac
		rs.GetValue(_T("rank"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		//dv
		rs.GetValue(_T("obj"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		//chandoan
		rs.GetValue(_T("he_examdate"), tmpStr);
		rptDetail->SetValue(_T("7"), CDateTime::Convert(tmpStr, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
		//doituong
		rs.GetValue(_T("cardno"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
	
		rs.GetValue(_T("diagno"), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("doctor"), tmpStr);
		rptDetail->SetValue(_T("10"),  pMF->GetDoctorName(tmpStr));
		
		rs.GetValue(_T("inward"), tmpStr);
		if(!tmpStr.IsEmpty())
			c1++;
		rptDetail->SetValue(_T("11"), tmpStr);
		
		rs.GetValue(_T("home"), tmpStr);
		if(!tmpStr.IsEmpty())
			c2++;
		rptDetail->SetValue(_T("12"), tmpStr);

		rs.MoveNext();
	}
	
	
	rpt.GetReportHeader()->SetValue(_T("ts"), nIdx++);
	tmpStr.Format(_T("%d"), c1++);
	rpt.GetReportHeader()->SetValue(_T("vaovien"), tmpStr);
	
	tmpStr.Format(_T("%d"), c2++);
	rpt.GetReportHeader()->SetValue(_T("venha"), tmpStr);
	CString szSysDate;
	szSysDate = pMF->GetSysDate(); 
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CEMExamListReportC13::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0;
	int nIdx = 0;
	int c1 = 0, c2 = 0;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	BeginWaitCursor();
	int ret = rs.ExecSQL(szSQL);
		_fmsg(_T("%s"), szSQL);
	if (ret <=0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	
	//Width
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(1, 15);
	xls.SetColumnWidth(2, 25);
	xls.SetColumnWidth(3, 11);
	xls.SetColumnWidth(4, 20);
	xls.SetColumnWidth(5, 15);
	xls.SetColumnWidth(6, 18);
	xls.SetColumnWidth(7, 20);
	xls.SetColumnWidth(8, 35);
	xls.SetColumnWidth(9, 25);
	xls.SetColumnWidth(10, 15);
	xls.SetColumnWidth(11, 15);


	xls.SetCellMergedColumns(nCol, nRow , 4);
	xls.SetCellMergedColumns(nCol, nRow+1, 4);
	xls.SetCellMergedColumns(nCol, nRow+2, 10);
	xls.SetCellMergedColumns(nCol, nRow+3, 10);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow+1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow+2, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N V\xC0O VI\x1EC6N"), FMT_TEXT | FMT_CENTER, true, 11);
	tmpStr.Format(_T("T\x1EEB ng\xE0y: %s \x111\x1EBFn ng\xE0y: %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow+3, tmpStr, FMT_TEXT | FMT_CENTER, false, 10);
	

	//Column header
	nRow = 4;
	CStringArray arrCol;
	arrCol.Add(_T("Stt"));
	arrCol.Add(_T("S\x1ED1 h\x1ED3 s\x1A1"));
	arrCol.Add(_T("H\x1ECD v\xE0 t\xEAn"));
	arrCol.Add(_T("N\x103m sinh"));
	arrCol.Add(_T("\x43\x1EA5p \x62\x1EAD\x63"));
	arrCol.Add(_T("\x110\x1ED1i t\x1B0\x1EE3ng"));
	arrCol.Add(_T("Ng\xE0y kh\xE1m"));
	arrCol.Add(_T("S\x1ED5 \x42HYT"));
	arrCol.Add(_T("\x43h\x1EA9n \x111o\xE1n"));
	arrCol.Add(_T("\x42\xE1\x63 s\x1EF9"));
	arrCol.Add(_T("V\xE0o vi\x1EC7n"));
	arrCol.Add(_T("V\x1EC1 nh\xE0"));

	for (int i = 0;i<arrCol.GetCount();i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10);	
	}
	//get data

	while (!rs.IsEOF())
	{
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow+1, tmpStr, FMT_INTEGER);
		//docno
		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol+1, nRow+1, tmpStr, FMT_INTEGER);
		//tenbenhnhan
		rs.GetValue(_T("pname"), tmpStr);
		xls.SetCellText(nCol+2, nRow+1, tmpStr, FMT_TEXT);
		//namsinh
		rs.GetValue(_T("yob"), tmpStr);
		xls.SetCellText(nCol+3, nRow+1, tmpStr, FMT_INTEGER);
		//capbac
		rs.GetValue(_T("rank"), tmpStr);
		xls.SetCellText(nCol+4, nRow+1, tmpStr, FMT_TEXT);
		//doi tuong
		rs.GetValue(_T("obj"), tmpStr);
		xls.SetCellText(nCol+5, nRow+1, tmpStr, FMT_TEXT);
		// ngay kham
		rs.GetValue(_T("he_examdate"), tmpStr);
		xls.SetCellText(nCol+6, nRow+1, CDateTime::Convert(tmpStr, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), FMT_TEXT);
		// the
		rs.GetValue(_T("cardno"), tmpStr);
		xls.SetCellText(nCol+7, nRow+1, tmpStr, FMT_TEXT);
		//chandoan
		rs.GetValue(_T("diagno"), tmpStr);
		xls.SetCellText(nCol+8, nRow+1, tmpStr, FMT_TEXT);
		// bac sy
		rs.GetValue(_T("doctor"), tmpStr);
		xls.SetCellText(nCol+9, nRow+1, pMF->GetDoctorName(tmpStr), FMT_TEXT);
		//vaovien
		rs.GetValue(_T("inward"), tmpStr);
		if(!tmpStr.IsEmpty()) c1++;
		xls.SetCellText(nCol+10, nRow+1, tmpStr, FMT_TEXT);
		//venha
		rs.GetValue(_T("home"), tmpStr);
		if(!tmpStr.IsEmpty()) c2++;
		xls.SetCellText(nCol+11, nRow+1, tmpStr, FMT_TEXT);
		
		nRow++;
		rs.MoveNext();
	}
	CString szTemp;
	szTemp.Format(_T("T\x1ED5ng s\x1ED1: %d"),nIdx++);
	xls.SetCellText(nCol + 4, 0, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("V\xE0o vi\x1EC7n: %d"),c1++);
	xls.SetCellText(nCol + 5, 0, szTemp, FMT_TEXT, true, 12);

	szTemp.Format(_T("V\x1EC1 nh\xE0: %d"),c2++);
	xls.SetCellText(nCol + 6, 0, szTemp, FMT_TEXT, true, 12);

	//nRow++;
	xls.SetCellMergedColumns(0, nRow + 1, 10);
	xls.SetCellText(0, nRow + 1, _T("T\x1ED5ng s\x1ED1:"), FMT_TEXT|FMT_CENTER, true, 14);

	tmpStr.Format(_T("%d"), c1);
	xls.SetCellText(nCol + 10, nRow + 1, tmpStr, FMT_NUMBER1,true, 14);

	tmpStr.Format(_T("%d"), c2);
	xls.SetCellText(nCol + 11, nRow + 1, tmpStr, FMT_NUMBER1, true, 14);

	xls.Save(_T("Exports\\Danh sach BN vao vien.xls"));
	EndWaitCursor();
	
} 
int CEMExamListReportC13::OnAddEMInwardListReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMExamListReportC13::OnEditEMInwardListReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMExamListReportC13::OnDeleteEMInwardListReport(){
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
 		OnCancelEMInwardListReport(); 
 	}
	return 0;
}
int CEMExamListReportC13::OnSaveEMInwardListReport(){
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
 		//OnEMInwardListReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMExamListReportC13::OnCancelEMInwardListReport(){
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
int CEMExamListReportC13::OnEMInwardListReportListLoadData(){
	return 0;
}

CString CEMExamListReportC13::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND he_examdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szObjectKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hd_object = '%s'"), m_szObjectKey);

	szSQL.Format(_T(" SELECT distinct he_docno  AS docno, he_doctor as doctor,") \
				_T("   he_examdate,") \
				_T("   get_patientname(he_docno) AS pname,") \
				_T("   Extract(YEAR FROM hp_birthdate) AS yob,") \
				_T("   get_syssel_desc('hms_rank', hp_rank) AS rank,") \
				_T("   hp_workplace                    AS workplace,") \
				_T("   he_diagnostic                   AS diagno,") \
				_T("   CASE WHEN hd_suggestion IN ('C', 'D')") \
				_T("        AND Length(hd_indept)>0 THEN hd_indept") \
				_T("   END                             AS inward,") \
				_T("	CASE WHEN hd_suggestion in ('E', 'A') then 'X'  else '' end as home,") \
				_T("   hd_cardno                       AS cardno,") \
				_T("   get_objectname(hd_object) as obj") \
				_T(" FROM   hms_patient") \
				_T(" LEFT JOIN hms_doc ON (hd_patientno=hp_patientno)") \
				_T(" LEFT JOIN hms_exam ON (he_docno=hd_docno)") \
				_T(" WHERE  he_status IN ('P', 'T') AND hd_enddept = 'C1.3' %s") \
				_T(" ORDER  BY he_examdate "), szWhere);

	//QueryOpener(szSQL);
	return szSQL;
}