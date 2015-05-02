#include "stdafx.h"
#include "EMWeekSynthesisReport.h"
#include "HMSMainFrame.h"
#include "Excel.h"
#include "ReportDocument.h"
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CEMWeekSynthesisReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CEMWeekSynthesisReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CEMWeekSynthesisReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CEMWeekSynthesisReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CEMWeekSynthesisReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CEMWeekSynthesisReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CEMWeekSynthesisReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CEMWeekSynthesisReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintSelectFnc(CWnd *pWnd){
	CEMWeekSynthesisReport *pVw = (CEMWeekSynthesisReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CEMWeekSynthesisReport *pVw = (CEMWeekSynthesisReport *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddEMWeekSynthesisReportFnc(CWnd *pWnd){
	 return ((CEMWeekSynthesisReport*)pWnd)->OnAddEMWeekSynthesisReport();
} 
static int _OnEditEMWeekSynthesisReportFnc(CWnd *pWnd){
	 return ((CEMWeekSynthesisReport*)pWnd)->OnEditEMWeekSynthesisReport();
} 
static int _OnDeleteEMWeekSynthesisReportFnc(CWnd *pWnd){
	 return ((CEMWeekSynthesisReport*)pWnd)->OnDeleteEMWeekSynthesisReport();
} 
static int _OnSaveEMWeekSynthesisReportFnc(CWnd *pWnd){
	 return ((CEMWeekSynthesisReport*)pWnd)->OnSaveEMWeekSynthesisReport();
} 
static int _OnCancelEMWeekSynthesisReportFnc(CWnd *pWnd){
	 return ((CEMWeekSynthesisReport*)pWnd)->OnCancelEMWeekSynthesisReport();
} 
CEMWeekSynthesisReport::CEMWeekSynthesisReport(CWnd *pParent){

	m_nDlgWidth = 500;
	m_nDlgHeight = 100;
	SetDefaultValues();
}
CEMWeekSynthesisReport::~CEMWeekSynthesisReport(){
}
void CEMWeekSynthesisReport::OnCreateComponents(){
	m_wndExaminationListReport.Create(this, _T("Weekly Synthesis Report"), 5, 5, 490, 60);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55); 
	m_wndPrint.Create(this, _T("&Print"), 335, 65, 410, 90);
	m_wndExport.Create(this, _T("&Export"), 415, 65, 490, 90);
	m_wndSpeciality.Create(this, _T("Speciality"), 5, 65, 90, 90);
}
void CEMWeekSynthesisReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(TRUE);
	//EnableButtons(TRUE, 0, -1);
	//m_wndFromDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndFromDate.SetCheckValue(true);
	//m_wndToDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndToDate.SetCheckValue(true);

}
void CEMWeekSynthesisReport::OnSetWindowEvents(){
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
	AddEvent(1, _T("Add	Ctrl+A"), _OnAddEMWeekSynthesisReportFnc, 0, 'A', VK_CONTROL);
	AddEvent(2, _T("Edit	Ctrl+E"), _OnEditEMWeekSynthesisReportFnc, 0, 'E', VK_CONTROL);
	AddEvent(3, _T("Delete	Ctrl+D"), _OnDeleteEMWeekSynthesisReportFnc, 0, 'D', VK_CONTROL);
	AddEvent(4, _T("Save	Ctrl+S"), _OnSaveEMWeekSynthesisReportFnc, 0, 'S', VK_CONTROL);
	AddEvent(0, _T("-"));
	AddEvent(5, _T("Cancel	Ctrl+T"), _OnCancelEMWeekSynthesisReportFnc, 0, 'T', VK_CONTROL);
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(false);

}
void CEMWeekSynthesisReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_Check(pDX, m_wndSpeciality.GetDlgCtrlID(), m_bSpeciality);
}
void CEMWeekSynthesisReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CEMWeekSynthesisReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CEMWeekSynthesisReport::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();

}
int CEMWeekSynthesisReport::SetMode(int nMode){
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
/*void CEMWeekSynthesisReport::OnFromDateChange(){
	
} */
/*void CEMWeekSynthesisReport::OnFromDateSetfocus(){
	
} */
/*void CEMWeekSynthesisReport::OnFromDateKillfocus(){
	
} */
int CEMWeekSynthesisReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CEMWeekSynthesisReport::OnToDateChange(){
	
} */
/*void CEMWeekSynthesisReport::OnToDateSetfocus(){
	
} */
/*void CEMWeekSynthesisReport::OnToDateKillfocus(){
	
} */
int CEMWeekSynthesisReport::OnToDateCheckValue(){
	return 0;
} 
void CEMWeekSynthesisReport::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szCol, szDate, szWhere;
	int nIdx = 1, nTemp = 0;
	int nTotal[12];
	szSQL = GetQueryString();
	rs.ExecSQL(szSQL);
	if (rs.IsEOF())
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	if (!rpt.Init(_T("Reports\\HMS\\HE_BAOCAOTONGHOPTUAN.RPT")))
		return;
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), pMF->m_CompanyInfo.sc_pname);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), pMF->m_CompanyInfo.sc_name);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection* rptDetail;
	for (int i = 0; i < 12; i++)
		nTotal[i] = 0;
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		tmpStr.Format(_T("%d"), nIdx++);
		rptDetail->SetValue(_T("1"), tmpStr);
		rs.GetValue(_T("roomid"), tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);
		for (int i = 0; i < 12; i++)
		{
			szCol.Format(_T("c%d"), i+1);
			rs.GetValue(szCol, nTemp);
			//FormatCurrency(nTemp, tmpStr);
			szCol.Format(_T("%d"), i+3);
			rptDetail->SetValue(szCol, int2str(nTemp));
			nTotal[i] += nTemp;
		}
		rs.MoveNext();
	}
	rptDetail = rpt.GetReportFooter();
	for (int i = 0; i < 12; i++)
	{
		szCol.Format(_T("s%d"), i+3);
		rptDetail->SetValue(szCol, int2str(nTotal[i]));
	}
	szWhere.Format(_T(" AND hcr_admitdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (pMF->GetModuleID() == _T("EM"))
		szWhere.AppendFormat(_T(" AND hd_enddept = '%s'"), pMF->GetCurrentDepartmentID());
	szSQL.Format(_T(" SELECT sum(c1) as c1,") \
				_T("        sum(c2) as c2,") \
				_T("        sum(c3) as c3,") \
				_T("        sum(c4) as c4,") \
				_T("        sum(c5) as c5,") \
				_T("        sum(c6) as c6,") \
				_T("        sum(c7) as c7,") \
				_T("        sum(c8) as c8,") \
				_T("        sum(c9) as c9,") \
				_T("        sum(c10) as c10,  ") \
				_T("        sum(c11) as c11,  ") \
				_T("        sum(c12) as c12  ") \
				_T(" FROM") \
				_T(" (") \
				_T("   SELECT") \
				_T("      case when hd_object=1 then 1 else 0 end as c1,") \
				_T("      case when hd_object=2 then 1 else 0 end as c2,") \
				_T("      case when hd_object=8 then 1 else 0 end as c3,") \
				_T("      case when hd_object in(6, 9) then 1 else 0 end as c4,") \
				_T("      case when hd_object=3 then 1 else 0 end as c5,") \
				_T("      case when hd_object=7 then 1 else 0 end as c6,") \
				_T("      case when hd_object=4 then 1 else 0 end as c7,") \
				_T("      case when hd_object=5 then 1 else 0 end as c8,") \
				_T("      case when hd_object=10 then 1 else 0 end as c9,") \
				_T("      case when hd_object=11 then 1 else 0 end as c10,") \
				_T("      case when hd_object=12 then 1 else 0 end as c11,") \
				_T("      1 as c12 ") \
				_T("    FROM hms_doc") \
				_T("    LEFT JOIN hms_exam ON(he_docno=hd_docno AND he_doctor = hd_doctor)") \
				_T("	LEFT JOIN hms_clinical_record ON (hcr_docno = hd_docno)") \
				_T("    WHERE hcr_numinward > 0 %s") \
				_T(" ) tbl"), szWhere);
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
	if (!rs.IsEOF()){
		for (int i = 0; i< 12; i++)
		{
			szCol.Format(_T("c%d"), i+1);
			rs.GetValue(szCol, nTemp);
			//FormatCurrency(nTemp, tmpStr);
			szCol.Format(_T("v%d"), i+3);
			rptDetail->SetValue(szCol, int2str(nTemp));
		}
	}
	else
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	szDate = pMF->GetSysDate();
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")), szDate.Right(2), szDate.Mid(5, 2), szDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
	
} 
void CEMWeekSynthesisReport::OnExportSelect(){
	_debug(_T("%s"), CString(typeid(this).name()));
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CExcel xls;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr, szTemp, szWhere;
	int nCol = 0, nRow = 0;
	int nTotal[12], nTemp = 0;
	szSQL = GetQueryString();
	int nCount = rs.ExecSQL(szSQL);
	if (nCount <= 0)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	xls.CreateSheet(1);
	xls.SetWorksheet(0);
	xls.SetColumnWidth(0, 30);
	xls.SetColumnWidth(1, 10);
	xls.SetColumnWidth(2, 10);
	xls.SetColumnWidth(3, 10);
	xls.SetColumnWidth(4, 10);
	xls.SetColumnWidth(5, 10);
	xls.SetColumnWidth(6, 10);
	xls.SetColumnWidth(7, 10);
	xls.SetColumnWidth(8, 10);
	xls.SetColumnWidth(9, 10);
	xls.SetColumnWidth(10, 10);
	xls.SetColumnWidth(11, 13);

	CellFormat df(&xls), hf(&xls), nf(&xls);
	df.SetItalic(true);
	df.SetCellStyle(FMT_TEXT | FMT_CENTER);
	hf.SetBold(true);
	hf.SetCellStyle(FMT_TEXT | FMT_CENTER);
	nf.SetCellStyle(FMT_NUMBER1);
	//Header
	xls.SetCellMergedColumns(nCol, nRow, 3);
	xls.SetCellMergedColumns(nCol, nRow + 1, 3);
	xls.SetCellMergedColumns(nCol, nRow + 2, 11);
	xls.SetCellMergedColumns(nCol, nRow + 3, 11);
	xls.SetCellText(nCol, nRow, pMF->m_CompanyInfo.sc_pname, FMT_TEXT | FMT_CENTER, true, 10);
	xls.SetCellText(nCol, nRow + 1, pMF->m_CompanyInfo.sc_name, FMT_TEXT | FMT_CENTER, true, 10);
	TranslateString(_T("Weekly Synthesis Report"), szTemp);
	StringUpper(szTemp, tmpStr);
	xls.SetCellText(nCol, nRow + 2, tmpStr, FMT_TEXT | FMT_CENTER, true, 12);	
	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	xls.SetCellText(nCol, nRow + 3, tmpStr, &df);	
	
	//Column Header
	CStringArray arrCol;
	arrCol.Add(_T("Ph\xF2ng"));
	arrCol.Add(_T("Qu\xE2n"));
	arrCol.Add(_T("\x42HYT Qu\xE2n"));
	arrCol.Add(_T("\x42\x1EA1n"));
	arrCol.Add(_T("Tr\x1EBB < \x36t"));
	arrCol.Add(_T("\x43h\xEDnh s\xE1\x63h"));
	arrCol.Add(_T("\x44\xE2n"));
	arrCol.Add(_T("\x42HYT kh\xE1\x63"));
	arrCol.Add(_T("\x42H TNSQ"));
	arrCol.Add(_T("\x42H Qu\xE2n \x111\x1ED9i"));
	arrCol.Add(_T("BHYT Qu\xE2n nh\xE2n"));
	arrCol.Add(_T("BHYT(N\x1EE3 th\x1EBB)"));
	arrCol.Add(_T("T\x1ED5ng \x63\x1ED9ng"));
	nRow = 4;
	for (int i = 0; i < arrCol.GetCount(); i++)
	{
		xls.SetCellText(nCol+i, nRow, arrCol.GetAt(i), FMT_TEXT | FMT_CENTER, true, 10); 
	}
	for (int i = 0; i< 12; i++)
	{
		nTotal[i] = 0;
	}
	nRow = 5;
	while (!rs.IsEOF()){
		rs.GetValue(_T("roomid"), tmpStr);
		xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT);
		for (int i = 0; i< 12; i++)
		{
			szTemp.Format(_T("c%d"), i+1);
			rs.GetValue(szTemp, nTemp);
			tmpStr.Format(_T("%d"), nTemp);
			xls.SetCellText(nCol+i+1, nRow, tmpStr, &nf);
			nTotal[i] += nTemp;
		}
		nRow++;
		rs.MoveNext();
	}
	TranslateString(_T("Total"), tmpStr);
	xls.SetCellText(nCol, nRow, tmpStr, FMT_TEXT | FMT_CENTER, true);
	for (int i = 0; i<12; i++)
	{
		tmpStr.Format(_T("%d"), nTotal[i]);
		xls.SetCellText(nCol+i+1, nRow, tmpStr, FMT_NUMBER1, true);
	}
	szWhere.Format(_T(" AND hcr_admitdate BETWEEN cast_string2timestamp('%s') AND cast_string2timestamp('%s')"), m_szFromDate, m_szToDate);
	if (pMF->GetModuleID() == _T("EM"))
		szWhere.AppendFormat(_T(" AND hd_enddept = '%s'"), pMF->GetCurrentDepartmentID());
	szSQL.Format(_T(" SELECT sum(c1) as c1,") \
				_T("        sum(c2) as c2,") \
				_T("        sum(c3) as c3,") \
				_T("        sum(c4) as c4,") \
				_T("        sum(c5) as c5,") \
				_T("        sum(c6) as c6,") \
				_T("        sum(c7) as c7,") \
				_T("        sum(c8) as c8,") \
				_T("        sum(c9) as c9,") \
				_T("        sum(c10) as c10,  ") \
				_T("        sum(c11) as c11,  ") \
				_T("        sum(c12) as c12  ") \
				_T(" FROM") \
				_T(" (") \
				_T("   SELECT") \
				_T("      case when hd_object=1 then 1 else 0 end as c1,") \
				_T("      case when hd_object=2 then 1 else 0 end as c2,") \
				_T("      case when hd_object=8 then 1 else 0 end as c3,") \
				_T("      case when hd_object in(6, 9) then 1 else 0 end as c4,") \
				_T("      case when hd_object=3 then 1 else 0 end as c5,") \
				_T("      case when hd_object=7 then 1 else 0 end as c6,") \
				_T("      case when hd_object=4 then 1 else 0 end as c7,") \
				_T("      case when hd_object=5 then 1 else 0 end as c8,") \
				_T("      case when hd_object=10 then 1 else 0 end as c9,") \
				_T("      case when hd_object=11 then 1 else 0 end as c10,") \
				_T("      case when hd_object=12 then 1 else 0 end as c11,") \
				_T("      1 as c12 ") \
				_T("    FROM hms_doc") \
				_T("	LEFT JOIN hms_exam ON (he_docno = hd_docno AND hd_doctor = he_doctor)") \
				_T("    LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
				_T("    WHERE hcr_numinward > 0 %s") \
				_T(" ) tbl"), szWhere);
	rs.ExecSQL(szSQL);
	_fmsg(_T("%s"), szSQL);
	if (!rs.IsEOF()){
		nRow++;
		xls.SetCellText(nCol, nRow, _T("V\xE0o vi\x1EC7n"), FMT_TEXT | FMT_CENTER, true);
		for (int i = 0; i< 12; i++)
		{
			szTemp.Format(_T("c%d"), i+1);
			rs.GetValue(szTemp, tmpStr);
			xls.SetCellText(nCol+i+1, nRow, tmpStr, FMT_NUMBER1, true);
		}
	}
	else
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}
	xls.Save(_T("Exports\\Bao Cao Tong Hop Tuan C1_1.xls"));
	
} 
int CEMWeekSynthesisReport::OnAddEMWeekSynthesisReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CEMWeekSynthesisReport::OnEditEMWeekSynthesisReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CEMWeekSynthesisReport::OnDeleteEMWeekSynthesisReport(){
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
 		OnCancelEMWeekSynthesisReport(); 
 	}
	return 0;
}
int CEMWeekSynthesisReport::OnSaveEMWeekSynthesisReport(){
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
 		//OnEMWeekSynthesisReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CEMWeekSynthesisReport::OnCancelEMWeekSynthesisReport(){
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
int CEMWeekSynthesisReport::OnEMWeekSynthesisReportListLoadData(){
	return 0;
}
CString CEMWeekSynthesisReport::GetQueryString(){
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CString szSQL, szWhere;
	bool bDocTerminate = false;
	szWhere.Format(_T(" AND he_examdate BETWEEN to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss') AND to_timestamp('%s', 'yyyy/mm/dd hh24:mi:ss')"), m_szFromDate, m_szToDate);
	//if (bDocTerminate)
	//	szWhere.AppendFormat(_T(" AND hd_status = 'T'"));
	
	if (pMF->GetModuleID() == _T("EM"))
	{
		szWhere.AppendFormat(_T(" AND he_deptid = '%s'"), pMF->GetCurrentDepartmentID());
		if (m_bSpeciality)
			szWhere.AppendFormat(_T(" AND hrl_section IS NULL"));
		else
			szWhere.AppendFormat(_T(" AND hrl_section = '%s'"), pMF->GetCurrentDepartmentID());
	}
	szSQL.Format(_T(" SELECT  (select hrl_name from hms_roomlist where hrl_id = roomid and hrl_deptid = he_deptid) as roomid, sum(c1) as c1,") \
				_T("         sum(c2) as c2,") \
				_T("         sum(c3) as c3,") \
				_T("         sum(c4) as c4,") \
				_T("         sum(c5) as c5,") \
				_T("         sum(c6) as c6,") \
				_T("         sum(c7) as c7,") \
				_T("         sum(c8) as c8,") \
				_T("         sum(c9) as c9,") \
				_T("         sum(c10) as c10, ") \
				_T("         sum(c11) as c11, ") \
				_T("         sum(c12) as c12 ") \
				_T(" FROM") \
				_T(" (") \
				_T("     SELECT he_deptid, he_roomid as roomid,") \
				_T("         case when hd_object=1 then 1 else 0 end as c1,") \
				_T("         case when hd_object=2 then 1 else 0 end as c2,") \
				_T("         case when hd_object=8 then 1 else 0 end as c3,") \
				_T("         case when hd_object in(6, 9) then 1 else 0 end as c4,") \
				_T("         case when hd_object=3 then 1 else 0 end as c5,") \
				_T("         case when hd_object=7 then 1 else 0 end as c6,") \
				_T("         case when hd_object=4 then 1 else 0 end as c7,") \
				_T("         case when hd_object=5 then 1 else 0 end as c8,") \
				_T("         case when hd_object=10 then 1 else 0 end as c9,") \
				_T("         case when hd_object=11 then 1 else 0 end as c10,") \
				_T("         case when hd_object=12 then 1 else 0 end as c11,") \
				_T("         1 as c12 ") \
				_T("     FROM hms_doc") \
				_T("     LEFT JOIN hms_exam ON(he_docno=hd_docno)") \
				_T("	 LEFT JOIN hms_roomlist ON (hrl_deptid = he_deptid AND hrl_id = he_roomid)") \
				_T("     WHERE he_status IN ('T', 'P')  %s") \
				_T("     and he_roomid > 0)") \
				_T(" tbl") \
				_T(" GROUP BY he_deptid, roomid") \
				_T(" ORDER BY he_deptid, roomid"), szWhere);
_fmsg(_T("%s"), szSQL);
	return szSQL;
}
