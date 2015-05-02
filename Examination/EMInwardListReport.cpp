#include "stdafx.h"
#include "EMInwardListReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMInwardListReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMInwardListReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnObjectSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMInwardListReport* )pWnd)->OnObjectSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnObjectSelendokFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnObjectSelendok();
}
/*static void _OnObjectSetfocusFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnObjectKillfocus();
}*/
/*static void _OnObjectKillfocusFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnObjectKillfocus();
}*/
static long _OnObjectLoadDataFnc(CWnd *pWnd){
	return ((CEMInwardListReport *)pWnd)->OnObjectLoadData();
}
/*static void _OnObjectAddNewFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnObjectAddNew();
}*/
static void _OnSectionSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CEMInwardListReport* )pWnd)->OnSectionSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnSectionSelendokFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnSectionSelendok();
}
/*static void _OnSectionSetfocusFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnSectionKillfocus();
}*/
/*static void _OnSectionKillfocusFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnSectionKillfocus();
}*/
static long _OnSectionLoadDataFnc(CWnd *pWnd){
	return ((CEMInwardListReport *)pWnd)->OnSectionLoadData();
}
/*static void _OnSectionAddNewFnc(CWnd *pWnd){
	((CEMInwardListReport *)pWnd)->OnSectionAddNew();
}*/
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMInwardListReport *pVw = (CEMInwardListReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMInwardListReport *pVw = (CEMInwardListReport *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMInwardListReport*)pWnd)->OnAddEMInwardListReport();
} 
static int _OnEditEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMInwardListReport*)pWnd)->OnEditEMInwardListReport();
} 
static int _OnDeleteEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMInwardListReport*)pWnd)->OnDeleteEMInwardListReport();
} 
static int _OnSaveEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMInwardListReport*)pWnd)->OnSaveEMInwardListReport();
} 
static int _OnCancelEMInwardListReportFnc(CWnd *pWnd){
	 return ((CEMInwardListReport*)pWnd)->OnCancelEMInwardListReport();
} 
CEMInwardListReport::CEMInwardListReport(CWnd *pParent){

	m_nDlgWidth = 500;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CEMInwardListReport::~CEMInwardListReport(){
}
void CEMInwardListReport::OnCreateComponents(){
	m_wndInwardListReport.Create(this, _T("Inward List Report"), 5, 5, 490, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndObjectLabel.Create(this, _T("Object"), 10, 60, 90, 85);
	m_wndObject.Create(this,95, 60, 245, 85); 
	m_wndSectionLabel.Create(this, _T("Section"), 250, 60, 330, 85);
	m_wndSection.Create(this,335, 60, 485, 85); 
	m_wndPrint.Create(this, _T("&Print"), 335, 95, 410, 120);
	m_wndExport.Create(this, _T("&Export"), 415, 95, 490, 120);

}
void CEMInwardListReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);
	//m_wndObject.SetCheckValue(true);
	m_wndObject.LimitText(35);
	//m_wndSection.SetCheckValue(true);
	m_wndSection.LimitText(35);


	m_wndObject.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndObject.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);


	m_wndSection.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndSection.InsertColumn(1, _T("Description"), CFMT_TEXT, 150);

}
void CEMInwardListReport::OnSetWindowEvents(){
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
	m_wndSection.SetEvent(WE_SELENDOK, _OnSectionSelendokFnc);
	//m_wndSection.SetEvent(WE_SETFOCUS, _OnSectionSetfocusFnc);
	//m_wndSection.SetEvent(WE_KILLFOCUS, _OnSectionKillfocusFnc);
	m_wndSection.SetEvent(WE_SELCHANGE, _OnSectionSelectChangeFnc);
	m_wndSection.SetEvent(WE_LOADDATA, _OnSectionLoadDataFnc);
	//m_wndSection.SetEvent(WE_ADDNEW, _OnSectionAddNewFnc);
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
void CEMInwardListReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_TextEx(pDX, m_wndObject.GetDlgCtrlID(), m_szObjectKey);
	DDX_TextEx(pDX, m_wndSection.GetDlgCtrlID(), m_szSectionKey);

}
void CEMInwardListReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMInwardListReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMInwardListReport::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_szObjectKey.Empty();
	m_szSectionKey.Empty();

}
int CEMInwardListReport::SetMode(int nMode){
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
/*void CEMInwardListReport::OnFromDateChange(){
	
} */
/*void CEMInwardListReport::OnFromDateSetfocus(){
	
} */
/*void CEMInwardListReport::OnFromDateKillfocus(){
	
} */
int CEMInwardListReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMInwardListReport::OnToDateChange(){
	
} */
/*void CEMInwardListReport::OnToDateSetfocus(){
	
} */
/*void CEMInwardListReport::OnToDateKillfocus(){
	
} */
int CEMInwardListReport::OnToDateCheckValue(){
	return 0;
} 
void CEMInwardListReport::OnObjectSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMInwardListReport::OnObjectSelendok(){
	 
}
/*void CEMInwardListReport::OnObjectSetfocus(){
	
}*/
/*void CEMInwardListReport::OnObjectKillfocus(){
	
}*/
long CEMInwardListReport::OnObjectLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndObject.IsSearchKey() && !m_szObjectKey.IsEmpty())
	{
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
/*void CEMInwardListReport::OnObjectAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMInwardListReport::OnSectionSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CEMInwardListReport::OnSectionSelendok(){
	 
}
/*void CEMInwardListReport::OnSectionSetfocus(){
	
}*/
/*void CEMInwardListReport::OnSectionKillfocus(){
	
}*/
long CEMInwardListReport::OnSectionLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndSection.IsSearchKey() && !m_szSectionKey.IsEmpty()){
	};
	szSQL.Format(_T("SELECT sd_id as id, sd_name as description FROM sys_dept WHERE sd_type = 'KB' ORDER BY sd_id"));
	m_wndSection.DeleteAllItems(); 
	int nCount = 0;
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndSection.AddItems(
			rs.GetValue(_T("id")), 
			rs.GetValue(_T("description")), NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CEMInwardListReport::OnSectionAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
void CEMInwardListReport::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr;
	int nRetired = 0, nCount = 0, ret = 0, nSoldier = 0;
	szSQL = GetQueryString();
	if (!rpt.Init(_T("Reports\\HMS\\HE_DANHSACHBENHNHANVAOVIEN.RPT")))
		return;
	ret = rs.ExecSQL(szSQL);
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
		rptDetail = rpt.AddDetail();
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		rptDetail->SetValue(_T("1"), tmpStr);
		//sovaovien
		rs.GetValue(_T("inwardno"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		//tenbn
		rs.GetValue(_T("tenbenhnhan"), tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);
		//namsinh
		rs.GetValue(_T("namsinh"), tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);
		//capbac
		rs.GetValue(_T("capbac"), tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);
		//dv
		rs.GetValue(_T("dv"), tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);
		//chandoan
		rs.GetValue(_T("chandoan"), tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);
		//doituong
		rs.GetValue(_T("doituong"), tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		rs.GetValue(_T("objid"), tmpStr);
		if (tmpStr == _T("2"))
			nRetired++;
		//vaovien
		rs.GetValue(_T("vaokhoa"), tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);
		rs.MoveNext();
	}
	nSoldier = nCount - nRetired;
	FormatCurrency(nSoldier, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("bodoi"), tmpStr);
	FormatCurrency(nRetired, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("huu"), tmpStr);
	FormatCurrency(nCount, tmpStr);
	rpt.GetReportHeader()->SetValue(_T("tongcong"), tmpStr);
	CString szSysDate;
	szSysDate = pMF->GetSysDate(); 
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
	
} 
void CEMInwardListReport::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CString szSQL, tmpStr;
	int nCol = 0, nRow = 0;
	CRecord rs(&pMF->m_db);
	szSQL = GetQueryString();
	BeginWaitCursor();
	int ret = rs.ExecSQL(szSQL);
	if (ret <=0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	
	//Width
	xls.SetColumnWidth(0, 5);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 25);
	xls.SetColumnWidth(5, 11);
	xls.SetColumnWidth(6, 30);
	xls.SetColumnWidth(7, 45);
	xls.SetColumnWidth(8, 10);

	xls.SetCellMergedColumns(nCol, nRow , 4);
	xls.SetCellMergedColumns(nCol, nRow+1, 4);
	xls.SetCellMergedColumns(nCol, nRow+2, 10);
	xls.SetCellMergedColumns(nCol, nRow+3, 10);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow+1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow+2, _T("\x44\x41NH S\xC1\x43H \x42\x1EC6NH NH\xC2N V\xC0O VI\x1EC6N"), FMT_TEXT | FMT_CENTER, true, 11);
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow+3, tmpStr, FMT_TEXT | FMT_CENTER, false, 10);
	//Column header
	nRow = 4;
	CStringArray arrCol;
	arrCol.Add(_T("Stt"));
	arrCol.Add(_T("Doc No"));
	arrCol.Add(_T("Inward No"));
	arrCol.Add(_T("T\xEAn \x62\x1EC7nh nh\xE2n"));
	arrCol.Add(_T("N\x103m sinh"));
	arrCol.Add(_T("\x43\x1EA5p \x62\x1EAD\x63"));
	arrCol.Add(_T("\x110v"));
	arrCol.Add(_T("\x43h\x1EA9n \x111o\xE1n"));
	arrCol.Add(_T("\x110\x1ED1i t\x1B0\x1EE3ng"));
	arrCol.Add(_T("V\xE0o kho\x61"));
	for (int i = 0;i<arrCol.GetCount();i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10);	
	}
	//get data
	int nIdx = 0;
	while (!rs.IsEOF())
	{
		nIdx++;
		tmpStr.Format(_T("%d"), nIdx);
		xls.SetCellText(nCol, nRow+1, tmpStr, FMT_INTEGER);
		//docno
		rs.GetValue(_T("docno"), tmpStr);
		xls.SetCellText(nCol+1, nRow+1, tmpStr, FMT_INTEGER);
		//inwardno
		rs.GetValue(_T("inwardno"), tmpStr);
		xls.SetCellText(nCol+2, nRow+1, tmpStr, FMT_INTEGER);
		//tenbenhnhan
		rs.GetValue(_T("tenbenhnhan"), tmpStr);
		xls.SetCellText(nCol+3, nRow+1, tmpStr, FMT_TEXT);
		//namsinh
		rs.GetValue(_T("namsinh"), tmpStr);
		xls.SetCellText(nCol+4, nRow+1, tmpStr, FMT_INTEGER);
		//capbac
		rs.GetValue(_T("capbac"), tmpStr);
		xls.SetCellText(nCol+5, nRow+1, tmpStr, FMT_TEXT);
		//dv
		rs.GetValue(_T("dv"), tmpStr);
		xls.SetCellText(nCol+6, nRow+1, tmpStr, FMT_TEXT);
		//chandoan
		rs.GetValue(_T("chandoan"), tmpStr);
		xls.SetCellText(nCol+7, nRow+1, tmpStr, FMT_TEXT);
		//doituong
		rs.GetValue(_T("doituong"), tmpStr);
		xls.SetCellText(nCol+8, nRow+1, tmpStr, FMT_TEXT);
		//vaokhoa
		rs.GetValue(_T("vaokhoa"), tmpStr);
		xls.SetCellText(nCol+9, nRow+1, tmpStr, FMT_TEXT);
		nRow++;
		rs.MoveNext();
	}
	xls.Save(_T("Exports\\Danh sach BN vao vien.xls"));
	EndWaitCursor();
	
} 
int CEMInwardListReport::OnAddEMInwardListReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMInwardListReport::OnEditEMInwardListReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMInwardListReport::OnDeleteEMInwardListReport(){
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
int CEMInwardListReport::OnSaveEMInwardListReport(){
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
int CEMInwardListReport::OnCancelEMInwardListReport(){
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
int CEMInwardListReport::OnEMInwardListReportListLoadData(){
	return 0;
}
CString CEMInwardListReport::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere;
	szWhere.Format(_T(" AND hcr_admitdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (!m_szObjectKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hd_object = '%s'"), m_szObjectKey);
	if (!m_szSectionKey.IsEmpty())
		szWhere.AppendFormat(_T(" AND hd_enddept = '%s'"), m_szSectionKey);
	szSQL.Format(_T(" SELECT  hcr_docno as docno,") \
				_T("         hcr_recordno as inwardno,") \
				_T("         trim(hp_surname||' '||hp_midname||' '||hp_firstname) as tenbenhnhan,") \
				_T("         extract (year from hp_birthdate) as namsinh,") \
				_T("         (select distinct ss_desc from sys_sel where ss_id = 'hms_rank' and ss_code = hp_rank) as capbac,") \
				_T("         hp_workplace as dv,") \
				_T("         hcr_admitdisease as chandoan,") \
				_T("         case when hd_object <> 2 then ") \
				_T("         (select distinct ho_desc from hms_object where ho_id = hd_object) ") \
				_T("         else (select distinct ss_desc from sys_sel where ss_id = 'sys_occupation' and cast(ss_code as integer)= 24)") \
				_T("         end as doituong, hd_object as objid,") \
				_T("         hcr_admitdept as vaokhoa") \
				_T(" FROM hms_clinical_record") \
				_T(" LEFT JOIN hms_doc ON (hd_docno = hcr_docno)") \
				_T(" LEFT JOIN hms_patient ON (hcr_patientno = hp_patientno)") \
				_T(" WHERE hcr_numinward > 0") \
				_T("  %s") \
				_T(" ORDER BY hcr_docno"), szWhere);
	//QueryOpener(szSQL);
	return szSQL;
}