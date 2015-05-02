#include "stdafx.h"
#include "HMSMainFrame.h"
#include "ArmyReportsDept.h"
/*static void _OnYearChangeFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnYearChange();
} */
/*static void _OnYearSetfocusFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnYearSetfocus();} */ 
/*static void _OnYearKillfocusFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnYearKillfocus();
} */
static int _OnYearCheckValueFnc(CWnd *pWnd){
	return ((CArmyReportsDept *)pWnd)->OnYearCheckValue();
} 
static void _OnReportPeriodSelectChangeFnc(CWnd *pWnd, int nOldItemSel, int nNewItemSel){
	((CArmyReportsDept* )pWnd)->OnReportPeriodSelectChange(nOldItemSel, nNewItemSel);
} 
static void _OnReportPeriodSelendokFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnReportPeriodSelendok();
}
/*static void _OnReportPeriodSetfocusFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnReportPeriodKillfocus();
}*/
/*static void _OnReportPeriodKillfocusFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnReportPeriodKillfocus();
}*/
static long _OnReportPeriodLoadDataFnc(CWnd *pWnd){
	return ((CArmyReportsDept *)pWnd)->OnReportPeriodLoadData();
}
/*static void _OnReportPeriodAddNewFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnReportPeriodAddNew();
}*/
/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CArmyReportsDept *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CArmyReportsDept *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CArmyReportsDept *)pWnd)->OnToDateCheckValue();
} 
static void _OnPrintPreviewSelectFnc(CWnd *pWnd){
	CArmyReportsDept *pVw = (CArmyReportsDept *)pWnd;
	pVw->OnPrintPreviewSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CArmyReportsDept *pVw = (CArmyReportsDept *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddArmyReportsDeptFnc(CWnd *pWnd){
	 return ((CArmyReportsDept*)pWnd)->OnAddArmyReportsDept();
} 
static int _OnEditArmyReportsDeptFnc(CWnd *pWnd){
	 return ((CArmyReportsDept*)pWnd)->OnEditArmyReportsDept();
} 
static int _OnDeleteArmyReportsDeptFnc(CWnd *pWnd){
	 return ((CArmyReportsDept*)pWnd)->OnDeleteArmyReportsDept();
} 
static int _OnSaveArmyReportsDeptFnc(CWnd *pWnd){
	 return ((CArmyReportsDept*)pWnd)->OnSaveArmyReportsDept();
} 
static int _OnCancelArmyReportsDeptFnc(CWnd *pWnd){
	 return ((CArmyReportsDept*)pWnd)->OnCancelArmyReportsDept();
} 
static void _OnAllSelectFnc(CWnd *pWnd){
	((CArmyReportsDept*)pWnd)->OnAllSelect();
}
static void _OnInPatientSelectFnc(CWnd *pWnd){
	((CArmyReportsDept*)pWnd)->OnInPatientSelect();
}
static void _OnOutPatientSelectFnc(CWnd *pWnd){
	((CArmyReportsDept*)pWnd)->OnOutPatientSelect();
}
CArmyReportsDept::CArmyReportsDept(CWnd *pWnd){

	m_nDlgWidth = 1029;
	m_nDlgHeight = 773;
	SetDefaultValues();
}
CArmyReportsDept::~CArmyReportsDept(){
}
void CArmyReportsDept::OnCreateComponents(){
	m_wndYearLabel.Create(this, _T("Year"), 10, 30, 90, 55);
	m_wndYear.Create(this,95, 30, 215, 55); 
	m_wndReportPeriodLabel.Create(this, _T("Report Period"), 220, 30, 300, 55);
	m_wndReportPeriod.Create(this,305, 30, 425, 55); 
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 60, 90, 85);
	m_wndFromDate.Create(this,95, 60, 215, 85); 
	m_wndReportCondition.Create(this, _T("Report Condition"), 5, 5, 430, 91);
	m_wndToDateLabel.Create(this, _T("To Date"), 220, 60, 300, 85);
	m_wndToDate.Create(this,305, 60, 425, 85); 
	m_wndAll.Create(this, _T("All"), 6, 97, 86, 122);
	m_wndInPatient.Create(this, _T("InPatient"), 90, 97, 170, 122);
	m_wndOutPatient.Create(this, _T("OutPatient"), 174, 97, 254, 122);
	m_wndPrintPreview.Create(this, _T("Print Pre&view"), 235, 126, 345, 151);
	m_wndExport.Create(this, _T("&Export"), 350, 126, 425, 151);

	m_wndReportPeriod.InsertColumn(0, _T("ID"), CFMT_NUMBER, 50);
	m_wndReportPeriod.InsertColumn(1, _T("Description"), CFMT_TEXT, 200);
}
void CArmyReportsDept::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_wndYear.SetLimitText(1024);
	m_wndYear.SetCheckValue(true);
	m_wndReportPeriod.SetCheckValue(true);
	m_wndReportPeriod.LimitText(1024);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);

	



}
void CArmyReportsDept::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndYear.SetEvent(WE_CHANGE, _OnYearChangeFnc);
	//m_wndYear.SetEvent(WE_SETFOCUS, _OnYearSetfocusFnc);
	//m_wndYear.SetEvent(WE_KILLFOCUS, _OnYearKillfocusFnc);
	m_wndYear.SetEvent(WE_CHECKVALUE, _OnYearCheckValueFnc);
	m_wndReportPeriod.SetEvent(WE_SELENDOK, _OnReportPeriodSelendokFnc);
	//m_wndReportPeriod.SetEvent(WE_SETFOCUS, _OnReportPeriodSetfocusFnc);
	//m_wndReportPeriod.SetEvent(WE_KILLFOCUS, _OnReportPeriodKillfocusFnc);
	m_wndReportPeriod.SetEvent(WE_SELCHANGE, _OnReportPeriodSelectChangeFnc);
	m_wndReportPeriod.SetEvent(WE_LOADDATA, _OnReportPeriodLoadDataFnc);
	//m_wndReportPeriod.SetEvent(WE_ADDNEW, _OnReportPeriodAddNewFnc);
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndAll.SetEvent(WE_CLICK, _OnAllSelectFnc);
	m_wndInPatient.SetEvent(WE_CLICK, _OnInPatientSelectFnc);
	m_wndOutPatient.SetEvent(WE_CLICK, _OnOutPatientSelectFnc);
	m_wndPrintPreview.SetEvent(WE_CLICK, _OnPrintPreviewSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	CString szSysDate = pMF->GetSysDate();
	m_nYear = ToInt(szSysDate.Left(4));
	m_szReportPeriodKey.Format(_T("%d"), ToInt(szSysDate.Mid(5, 2)));
	m_szFromDate = m_szToDate = pMF->GetSysDate();
	m_szFromDate += _T("00:00");
	m_szToDate += _T("23:59");
	UpdateData(FALSE);
	m_wndAll.SetCheck(true);
	m_szCondition = _T("A");
}
void CArmyReportsDept::OnDoDataExchange(CDataExchange* pDX){
	DDX_Text(pDX, m_wndYear.GetDlgCtrlID(), m_nYear);
	DDX_TextEx(pDX, m_wndReportPeriod.GetDlgCtrlID(), m_szReportPeriodKey);
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_Radio(pDX, m_wndAll.GetDlgCtrlID(), m_nAll);
	DDX_Radio(pDX, m_wndInPatient.GetDlgCtrlID(), m_nInPatient);
	DDX_Radio(pDX, m_wndOutPatient.GetDlgCtrlID(), m_nOutPatient);

}
void CArmyReportsDept::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CArmyReportsDept::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CArmyReportsDept::SetDefaultValues(){

	m_nYear=0;
	m_szReportPeriodKey.Empty();
	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_nAll=0;
	m_nInPatient=1;
	m_nOutPatient=1;

}
int CArmyReportsDept::SetMode(int nMode){
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
/*void CArmyReportsDept::OnYearChange(){
	
} */
/*void CArmyReportsDept::OnYearSetfocus(){
	
} */
/*void CArmyReportsDept::OnYearKillfocus(){
	
} */
int CArmyReportsDept::OnYearCheckValue(){
	return 0;
} 
void CArmyReportsDept::OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CArmyReportsDept::OnReportPeriodSelendok(){
	CString tmpStr;
	CDate date;
	UpdateData(true);
	int nMonth = ToInt(m_szReportPeriodKey);
	if(nMonth > 0 && nMonth <= 12)
	{
		m_szFromDate.Format(_T("%.4d/%.2d/01"), m_nYear, nMonth);
		date.ParseDate(m_szFromDate);
		m_szToDate.Format(_T("%.4d/%.2d/%.2d 23:59"), m_nYear, nMonth, date.GetMonthLastDay());
	}
	if(nMonth == 13){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/03/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/03/%.2d 23:59" ), m_nYear,  date.GetMonthLastDay());
	}
	if(nMonth == 14){
		m_szFromDate.Format(_T("%.4d/04/01"), m_nYear);
		tmpStr.Format(_T("%.4d/06/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/06/%.2d 23:59"), m_nYear, date.GetMonthLastDay());
	}
	if(nMonth == 15){
		m_szFromDate.Format(_T("%.4d/07/01"), m_nYear);
		tmpStr.Format(_T("%.4d/09/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/09/%.2d 23:59"), m_nYear, date.GetMonthLastDay());
	}
	if(nMonth == 16){
		m_szFromDate.Format(_T("%.4d/10/01"), m_nYear);
		tmpStr.Format(_T("%.4d/10/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, date.GetMonthLastDay());
	}
	if(nMonth == 17){
		m_szFromDate.Format(_T("%.4d/01/01"), m_nYear);
		tmpStr.Format(_T("%.4d/12/01"), m_nYear);
		date.ParseDate(tmpStr);
		m_szToDate.Format(_T("%.4d/12/%.2d 23:59"), m_nYear, date.GetMonthLastDay());		
	}	
	UpdateData(false);	
}
/*void CArmyReportsDept::OnReportPeriodSetfocus(){
	
}*/
/*void CArmyReportsDept::OnReportPeriodKillfocus(){
	
}*/
long CArmyReportsDept::OnReportPeriodLoadData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL, szWhere;
	if(m_wndReportPeriod.IsSearchKey() && ToInt(m_szReportPeriodKey) > 0){
		szWhere.Format(_T(" WHERE hpr_idx=%ld"), ToInt(m_szReportPeriodKey));
	};
	m_wndReportPeriod.DeleteAllItems(); 
	int nCount = 0;
	szSQL.Format(_T("SELECT * FROM hms_period_report %s ORDER BY hpr_idx "), szWhere);
	nCount = rs.ExecSQL(szSQL);
	while(!rs.IsEOF()){ 
		m_wndReportPeriod.AddItems(
			rs.GetValue(_T("hpr_idx")), 
			rs.GetValue(_T("hpr_name")),			
			NULL);
		rs.MoveNext();
	}
	return nCount;
}
/*void CArmyReportsDept::OnReportPeriodAddNew(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} */
/*void CArmyReportsDept::OnFromDateChange(){
	
} */
/*void CArmyReportsDept::OnFromDateSetfocus(){
	
} */
/*void CArmyReportsDept::OnFromDateKillfocus(){
	
} */
int CArmyReportsDept::OnFromDateCheckValue(){
	return 0;
} 
/*void CArmyReportsDept::OnToDateChange(){
	
} */
/*void CArmyReportsDept::OnToDateSetfocus(){
	
} */
/*void CArmyReportsDept::OnToDateKillfocus(){
	
} */
int CArmyReportsDept::OnToDateCheckValue(){
	return 0;
} 
void CArmyReportsDept::OnAllSelect(){
	m_szCondition = _T("A");
}
void CArmyReportsDept::OnInPatientSelect(){
	m_szCondition = _T("I");
}
void CArmyReportsDept::OnOutPatientSelect(){
	m_szCondition = _T("E");
}
void CArmyReportsDept::OnPrintPreviewSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(true);
	CReport rpt;
	CRecord rs(&pMF->m_db);
	CString szSQL, tmpStr,szOldGroup,szNewGroup;
	int ret = 0 ;
	szSQL = GetQueryString();
	if (!rpt.Init(_T("Reports\\HMS\\HE_BAOCAOQUANSO.RPT")))
		return;
	ret = rs.ExecSQL(szSQL);
	/*if (ret <=0%%)
	{
		ShowMessage(150, MB_ICONSTOP);
		return;
	}*/
	rpt.GetReportHeader()->SetValue(_T("HealthService"), _T("\x42\x1ED8 QU\x1ED0\x43 PH\xD2NG"));
	rpt.GetReportHeader()->SetValue(_T("HospitalName"), _T("\x42\x1EC6NH VI\x1EC6N TW QU\xC2N \x110\x1ED8I \x31\x30\x38"));
	tmpStr.Format(_T("KHOA %s"),pMF->GetCurrentDepartmentID());
	rpt.GetReportHeader()->SetValue(_T("khoa"), tmpStr);
	tmpStr.Format(rpt.GetReportHeader()->GetValue(_T("ReportDate")), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));
	rpt.GetReportHeader()->SetValue(_T("ReportDate"), tmpStr);
	CReportSection *rptDetail = NULL;
	int nCONG[13];
	for(int i = 2 ; i<12; i++)
	{
		nCONG[i] = 0;
	}
	while (!rs.IsEOF())
	{
		rptDetail = rpt.AddDetail();
		rs.GetValue(_T("doituong"), tmpStr);
		rptDetail->SetValue(_T("1"), tmpStr);
	
		rs.GetValue(_T("socu"), tmpStr);
		nCONG[2] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("vaovien"), tmpStr);
		nCONG[3] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("0"), tmpStr);
		nCONG[4] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("ckden"), tmpStr);
		nCONG[5] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("ravien"), tmpStr);
		nCONG[6] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("nxv"), tmpStr);
		nCONG[7] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("chuyenvien"), tmpStr);
		nCONG[8] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);
		
		rs.GetValue(_T("ckdi"), tmpStr);
		nCONG[9] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("tuvong"), tmpStr);
		nCONG[10] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.GetValue(_T("hientai"), tmpStr);
		nCONG[11] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.MoveNext();
	}
	szSQL = GetQueryString1();
	ret = rs.ExecSQL(szSQL);
	while (!rs.IsEOF())
	{
		rs.GetValue(_T("id"), szNewGroup);
		if (szOldGroup != szNewGroup)
		{
			rptDetail = rpt.AddDetail(rpt.GetGroupHeader(1));
			rs.GetValue(_T("type"), tmpStr);
			rptDetail->SetValue(_T("groupname"), tmpStr);
			szOldGroup = szNewGroup;
		}
		rptDetail = rpt.AddDetail(rpt.GetGroupHeader(2));
		rs.GetValue(_T("doituong"), tmpStr);
		rptDetail->SetValue(_T("1"), tmpStr);

		rs.GetValue(_T("socu"), tmpStr);
		nCONG[2] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("2"), tmpStr);

		rs.GetValue(_T("vaovien"), tmpStr);
		nCONG[3] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("3"), tmpStr);

		rs.GetValue(_T("0"), tmpStr);
		nCONG[4] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("4"), tmpStr);

		rs.GetValue(_T("ckden"), tmpStr);
		nCONG[5] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("5"), tmpStr);

		rs.GetValue(_T("ravien"), tmpStr);
		nCONG[6] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("6"), tmpStr);

		rs.GetValue(_T("nxv"), tmpStr);
		nCONG[7] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("7"), tmpStr);

		rs.GetValue(_T("chuyenvien"), tmpStr);
		nCONG[8] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("8"), tmpStr);

		rs.GetValue(_T("ckdi"), tmpStr);
		nCONG[9] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("9"), tmpStr);

		rs.GetValue(_T("tuvong"), tmpStr);
		nCONG[10] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("10"), tmpStr);

		rs.GetValue(_T("hientai"), tmpStr);
		nCONG[11] +=ToInt(tmpStr);
		rptDetail->SetValue(_T("11"), tmpStr);

		rs.MoveNext();
	}
	CString  szField;
	for( int i = 2; i<=11; i++)
	{
		szField.Format(_T("s%d"), i);
		tmpStr.Format(_T("%d"), nCONG[i]);
		rpt.GetReportFooter()->SetValue(szField, tmpStr);
	}
	CString szSysDate;
	szSysDate = pMF->GetSysDate(); 
	tmpStr.Format(rpt.GetReportFooter()->GetValue(_T("PrintDate")),szSysDate.Right(2),szSysDate.Mid(5,2),szSysDate.Left(4));
	rpt.GetReportFooter()->SetValue(_T("PrintDate"), tmpStr);
	rpt.PrintPreview();
} 
void CArmyReportsDept::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);	CString szSQL, tmpStr, szTemp, szWhere,szOldGroup,szNewGroup;	UpdateData(TRUE);	BeginWaitCursor();	szWhere.Empty();	int nRow =0, nCol = 0;	CExcel xls;	xls.CreateSheet(1);	xls.SetWorksheet(0);	xls.SetColumnWidth(0, 19);	xls.SetColumnWidth(1, 9);	xls.SetColumnWidth(2, 12);	xls.SetColumnWidth(3, 19);	xls.SetColumnWidth(4, 17);	xls.SetColumnWidth(5, 14);	xls.SetColumnWidth(6, 13);	xls.SetColumnWidth(7, 16);	xls.SetColumnWidth(8, 16);	xls.SetColumnWidth(9, 8);	xls.SetColumnWidth(10, 8);	xls.SetCellText(0, 0, _T("\x42\x1ED8 QU\x1ED0\x43 PH\xD2NG"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 10);	xls.SetCellText(0, 1, _T("\x42\x1EC6NH VI\x1EC6N TW QU\xC2N \x110\x1ED8I \x31\x30\x38"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	CString title;	title.Format(_T("\x42\xC1O \x43\xC1O QU\xC2N S\x1ED0 KHOA %s"),pMF->GetCurrentDepartmentID());	xls.SetCellText(0, 3, title, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	tmpStr.Format(_T("T\x1EEB ng\xE0y %s \x111\x1EBFn ng\xE0y %s"), CDateTime::Convert(m_szFromDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss), CDateTime::Convert(m_szToDate, yyyymmdd|hhmmss, ddmmyyyy|hhmmss));	xls.SetCellText(0, 4, tmpStr, FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(0, 6, _T("Th\xE0nh Ph\x1EA7n"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(1, 6, _T("\x43\x169 \x63\xF2n l\x1EA1i"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(2, 6, _T("T\x103ng"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(2, 7, _T("\x110\x1A1n v\x1ECB \x111\x1EBFn"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(3, 7, _T("Vi\x1EC7n kh\xE1\x63 \x63huy\x1EC3n \x111\x1EBFn"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(4, 7, _T("\x43huy\x1EC3n kho\x61 \x111\x1EBFn"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(5, 6, _T("Gi\x1EA3m"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(5, 7, _T("Kh\x1ECFi v\x1EC1 \x111\x1A1n v\x1ECB"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(6, 7, _T("N\x1EB7ng \x78in v\x1EC1"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(7, 7, _T("\x43huy\x1EC3n vi\x1EC7n kh\xE1\x63"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(8, 7, _T("\x43huy\x1EC3n kho\x61 kh\xE1\x63"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(9, 7, _T("T\x1EED vong"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetCellText(10, 6, _T("\x43\xF2n l\x1EA1i"), FMT_TEXT | FMT_CENTER | FMT_WRAPING, true, 11);	xls.SetMerge(0, 0, 0, 2);	xls.SetMerge(1, 1, 0, 2);	xls.SetMerge(3, 3, 0, 10);	xls.SetMerge(4, 4, 0, 10);	xls.SetMerge(6, 7, 0, 0);	xls.SetMerge(6, 7, 1, 1);	xls.SetMerge(6, 6, 2, 4);	xls.SetMerge(6, 6, 5, 9);	xls.SetMerge(6, 7, 10, 10);	szSQL =GetQueryString();	rs.ExecSQL(szSQL);	nRow = 8;	nCol = 0;	int nTotal[11];
	for (int i = 1; i<= 10; i++)
	{
		nTotal[i] = 0;
	}	while(!rs.IsEOF()){		rs.GetValue(_T("doituong"), tmpStr);		xls.SetCellText(nCol+0, nRow, tmpStr, FMT_TEXT|FMT_WRAPING);		rs.GetValue(_T("socu"), tmpStr);		nTotal[1] += ToInt(tmpStr);		xls.SetCellText(nCol+1, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("vaovien"), tmpStr);		nTotal[2] += ToInt(tmpStr);		xls.SetCellText(nCol+2, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("0"), tmpStr);		nTotal[3] += ToInt(tmpStr);		xls.SetCellText(nCol+3, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("ckden"), tmpStr);		nTotal[4] += ToInt(tmpStr);		xls.SetCellText(nCol+4, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("ravien"), tmpStr);		nTotal[5] += ToInt(tmpStr);		xls.SetCellText(nCol+5, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("nxv"), tmpStr);		nTotal[6] += ToInt(tmpStr);		xls.SetCellText(nCol+6, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("chuyenvien"), tmpStr);		nTotal[7] += ToInt(tmpStr);		xls.SetCellText(nCol+7, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("ckdi"), tmpStr);		nTotal[8] += ToInt(tmpStr);		xls.SetCellText(nCol+8, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("tuvong"), tmpStr);		nTotal[9] += ToInt(tmpStr);		xls.SetCellText(nCol+9, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("hientai"), tmpStr);		nTotal[10] += ToInt(tmpStr);		xls.SetCellText(nCol+10, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		nRow++;		rs.MoveNext();	}	szSQL =GetQueryString1();	rs.ExecSQL(szSQL);	while(!rs.IsEOF()){		rs.GetValue(_T("id"), szNewGroup);
		if (szOldGroup != szNewGroup)
		{
			xls.SetCellText(0, nRow, rs.GetValue(_T("type")), FMT_TEXT, true);
			szOldGroup = szNewGroup;
			nRow++;
		}		rs.GetValue(_T("doituong"), tmpStr);		xls.SetCellText(nCol+0, nRow, tmpStr, FMT_TEXT|FMT_WRAPING);		rs.GetValue(_T("socu"), tmpStr);		nTotal[1] += ToInt(tmpStr);		xls.SetCellText(nCol+1, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("vaovien"), tmpStr);		nTotal[2] += ToInt(tmpStr);		xls.SetCellText(nCol+2, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("0"), tmpStr);		nTotal[3] += ToInt(tmpStr);		xls.SetCellText(nCol+3, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("ckden"), tmpStr);		nTotal[4] += ToInt(tmpStr);		xls.SetCellText(nCol+4, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("ravien"), tmpStr);		nTotal[5] += ToInt(tmpStr);		xls.SetCellText(nCol+5, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("nxv"), tmpStr);		nTotal[6] += ToInt(tmpStr);		xls.SetCellText(nCol+6, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("chuyenvien"), tmpStr);		nTotal[7] += ToInt(tmpStr);		xls.SetCellText(nCol+7, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("ckdi"), tmpStr);		nTotal[8] += ToInt(tmpStr);		xls.SetCellText(nCol+8, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("tuvong"), tmpStr);		nTotal[9] += ToInt(tmpStr);		xls.SetCellText(nCol+9, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		rs.GetValue(_T("hientai"), tmpStr);		nTotal[10] += ToInt(tmpStr);		xls.SetCellText(nCol+10, nRow, tmpStr, FMT_NUMBER1|FMT_CENTER);		nRow++;		rs.MoveNext();	}	xls.SetCellMergedColumns(nCol, nRow , 1);
	xls.SetCellText(nCol, nRow, _T("T\x1ED5ng"), FMT_TEXT | FMT_CENTER, true, 11);

	for( int i = 1; i<= 10; i++)
	{
		tmpStr.Format(_T("%d"), nTotal[i]);
		xls.SetCellText(i , nRow , tmpStr, FMT_NUMBER1|FMT_CENTER, true, 11);
	}	EndWaitCursor();	xls.Save(_T("Exports\\BAOCAOQUANSO.xls"));
} 
int CArmyReportsDept::OnAddArmyReportsDept(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CArmyReportsDept::OnEditArmyReportsDept(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CArmyReportsDept::OnDeleteArmyReportsDept(){
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
 		OnCancelArmyReportsDept();
 	}
	return 0;
}
int CArmyReportsDept::OnSaveArmyReportsDept(){
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
 		//OnArmyReportsDeptListLoadData();
 		SetMode(VM_VIEW);
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CArmyReportsDept::OnCancelArmyReportsDept(){
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
int CArmyReportsDept::OnArmyReportsDeptListLoadData(){
	return 0;
}
CString CArmyReportsDept::GetQueryString()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString  szWhere, szDept, szObject, szRank,szDept1;
	CString szSQL;
	//Trang thai dieu tri toan vien
		if (m_szCondition == _T("A"))
			szWhere.AppendFormat(_T(""));
		//Trang thai dieu tri noi tru
		if (m_szCondition == _T("I"))
			szWhere.AppendFormat(_T(" AND (htr_outpatient <> 'Y' OR htr_outpatient IS NULL) "));
		//Trang thai dieu tri ngoai tru
		if (m_szCondition == _T("E"))
			szWhere.AppendFormat(_T(" AND htr_OutPatient='Y'"));

		szSQL.Format(_T(" SELECT doituong,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
			_T(" FROM(SELECT doituong,") \
			_T("        count(distinct hd_docno) socu,0 vaovien,0 ckden,0 ravien,0 ckdi,0 chuyenvien,0 tuvong, 0 nxv") \
			_T(" FROM(SELECT hd_docno,doituong") \
			_T(" FROM(SELECT distinct hd_docno,") \
			_T("        hms_getobjectname(hd_object) AS doituong,") \
			_T("        trunc_date(htr_admitdate) AS admitdate,") \
			_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
			_T("        htr_deptid") \
			_T(" FROM hms_patient") \
			_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
			_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
			_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
			_T(" WHERE hd_object not in (1,2,12) and htr_deptid not in('A4') %s AND htr_status='I' AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s')") \
			_T(" UNION ALL") \
			_T(" SELECT distinct hd_docno,") \
			_T("        hms_getobjectname(hd_object) AS doituong,") \
			_T("        trunc_date(htr_admitdate) AS admitdate,") \
			_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
			_T("        htr_deptid") \
			_T(" FROM hms_patient") \
			_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
			_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
			_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
			_T(" WHERE hd_object not in (1,2,12) and htr_deptid not in('A4') %s AND htr_status='T' AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s') AND htr_dischargedate >= cast_string2timestamp('%s')))") \
			_T(" GROUP BY doituong") \
			_T(" UNION ALL") \
			_T(" SELECT hms_getobjectname(hd_object) AS doituong,0 socu,") \
			_T("        SUM(in_qty)       vaovien, ") \
			_T(" 		   SUM(movein_qty)   ckden, ") \
			_T(" 		   SUM(return_unit)  ravien, ") \
			_T(" 		   SUM(moveout_qty)  ckdi, ") \
			_T(" 		   SUM(transfer_qty) chuyenvien,  ") \
			_T(" 		 	 SUM(dead)         tuvong,") \
			_T("        SUM(nxv) nxv") \
			_T(" FROM(SELECT hd_object,") \
			_T("             Trunc(htr_admitdate) admit_date, ") \
			_T("             1                    in_qty, ") \
			_T(" 	          0                    movein_qty, ") \
			_T(" 	          0                    return_unit, ") \
			_T(" 	          0                    moveout_qty, ") \
			_T(" 	          0                    transfer_qty, ") \
			_T(" 	          0                    dead,") \
			_T("             0 nxv") \
			_T(" FROM   hms_doc ") \
			_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
			_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno)") \
			_T(" WHERE hd_object not in (1,2,12) and htr_status <> 'A' %s AND htr_deptid='%s' and hcr_admitdept=htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
			_T(" UNION ALL") \
			_T(" SELECT hd_object,") \
			_T("             Trunc(htr_admitdate) admit_date, ") \
			_T("             0                    in_qty, ") \
			_T(" 	          1                    movein_qty, ") \
			_T(" 	          0                    return_unit, ") \
			_T(" 	          0                    moveout_qty, ") \
			_T(" 	          0                    transfer_qty, ") \
			_T(" 	          0                    dead,") \
			_T("             0 nxv") \
			_T(" FROM   hms_doc ") \
			_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
			_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
			_T(" WHERE hd_object not in (1,2,12) and htr_status <> 'A' %s AND htr_deptid='%s' AND htr_idx > 1 AND hcr_admitdept<>htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
			_T(" UNION ALL") \
			_T(" SELECT hd_object,") \
			_T("             Trunc(htr_admitdate) admit_date, ") \
			_T("             0                    in_qty, ") \
			_T(" 	          0                    movein_qty, ") \
			_T(" 	          CASE WHEN htr_suggestion not in('M','F') AND hcr_result NOT IN ( 4, 5, 6,7 ) THEN 1 ELSE 0 END return_unit, ") \
			_T(" 	          0                    moveout_qty, ") \
			_T(" 	          CASE WHEN htr_suggestion = 'F' THEN 1 ELSE 0 END transfer_qty, ") \
			_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 5, 6 ) THEN 1 ELSE 0 END dead,") \
			_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 4, 7 ) THEN 1 ELSE 0 END nxv") \
			_T(" FROM   hms_doc ") \
			_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
			_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
			_T(" WHERE hd_object not in (1,2,12) and htr_status = 'T' %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
			_T(" UNION ALL") \
			_T(" SELECT hd_object,") \
			_T("             Trunc(htr_admitdate) admit_date, ") \
			_T("             0                    in_qty, ") \
			_T(" 	          0                    movein_qty, ") \
			_T(" 	          0 return_unit, ") \
			_T(" 	          1                    moveout_qty, ") \
			_T(" 	          0                    transfer_qty, ") \
			_T(" 	          0 dead,") \
			_T("             0 nxv") \
			_T(" FROM   hms_doc ") \
			_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
			_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
			_T(" WHERE hd_object not in (1,2,12) and ( htr_suggestion ='M') %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s'))") \
			_T(" GROUP BY hd_object)") \
			_T(" GROUP BY doituong") \
			_T(" ORDER BY doituong desc"),szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,
										 szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szFromDate,
										 szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
										 szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
										 szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
										 szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate);
	return szSQL;
}
CString CArmyReportsDept::GetQueryString1()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*)AfxGetMainWnd();
	CString  szWhere, szDept, szObject, szRank,szDept1;
	CString szSQL;
	//Trang thai dieu tri toan vien
	if (m_szCondition == _T("A"))
		szWhere.AppendFormat(_T(""));
	//Trang thai dieu tri noi tru
	if (m_szCondition == _T("I"))
		szWhere.AppendFormat(_T(" AND (htr_outpatient <> 'Y' OR htr_outpatient IS NULL) "));
	//Trang thai dieu tri ngoai tru
	if (m_szCondition == _T("E"))
		szWhere.AppendFormat(_T(" AND htr_OutPatient='Y'"));

	szSQL.Format(_T(" SELECT id,type,cast(doituong as NVARCHAR2(35)) doituong,socu,vaovien,ckden,ravien,ckdi,chuyenvien,tuvong,nxv,hientai") \
		_T(" FROM(SELECT 1 id,'Qu\xE2n' type,'T\x1B0\x1EDBng' doituong,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT doituong,hd_rank,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT doituong,hd_rank,") \
		_T("        count(distinct hd_docno) socu,0 vaovien,0 ckden,0 ravien,0 ckdi,0 chuyenvien,0 tuvong, 0 nxv") \
		_T(" FROM(SELECT hd_docno,doituong,hd_rank") \
		_T(" FROM(SELECT distinct hd_docno,hd_rank,") \
		_T("        hd_object AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_status='I' %s AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT distinct hd_docno,hd_rank,") \
		_T("        hd_object AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_status='T' %s AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s') AND htr_dischargedate >= cast_string2timestamp('%s')))") \
		_T(" GROUP BY doituong,hd_rank") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object AS doituong,hd_rank,0 socu,") \
		_T("        SUM(in_qty)       vaovien, ") \
		_T(" 		   SUM(movein_qty)   ckden, ") \
		_T(" 		   SUM(return_unit)  ravien, ") \
		_T(" 		   SUM(moveout_qty)  ckdi, ") \
		_T(" 		   SUM(transfer_qty) chuyenvien,  ") \
		_T(" 		 	 SUM(dead)         tuvong,") \
		_T("        SUM(nxv) nxv") \
		_T(" FROM(SELECT hd_object,") \
		_T("             hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             1                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          0                    return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0                    dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno)") \
		_T(" WHERE  htr_status <> 'A' %s AND htr_deptid='%s' and hcr_admitdept=htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          1                    movein_qty, ") \
		_T(" 	          0                    return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0                    dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE  htr_status <> 'A' %s AND htr_deptid='%s' AND htr_idx > 1 AND hcr_admitdept<>htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          CASE WHEN htr_suggestion not in('M','F') AND hcr_result NOT IN ( 4, 5, 6,7 ) THEN 1 ELSE 0 END return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          CASE WHEN htr_suggestion = 'F' THEN 1 ELSE 0 END transfer_qty, ") \
		_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 5, 6 ) THEN 1 ELSE 0 END dead,") \
		_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 4, 7 ) THEN 1 ELSE 0 END nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE htr_status = 'T' %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          0 return_unit, ") \
		_T(" 	          1                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0 dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE ( htr_suggestion ='M') %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s'))") \
		_T(" GROUP BY hd_object,hd_rank)") \
		_T(" GROUP BY doituong,hd_rank)") \
		_T(" WHERE doituong = 1 and hd_rank >=15") \
		_T(" GROUP BY doituong") \
		_T(" UNION ALL") \
		_T(" SELECT 1 id,'Qu\xE2n' type,'4//' doituong,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT doituong,hd_rank,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT doituong,hd_rank,") \
		_T("        count(distinct hd_docno) socu,0 vaovien,0 ckden,0 ravien,0 ckdi,0 chuyenvien,0 tuvong, 0 nxv") \
		_T(" FROM(SELECT hd_docno,doituong,hd_rank") \
		_T(" FROM(SELECT distinct hd_docno,hd_rank,") \
		_T("        hd_object AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_status='I' %s AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT distinct hd_docno,hd_rank,") \
		_T("        hd_object AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_status='T' %s AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s') AND htr_dischargedate >= cast_string2timestamp('%s')))") \
		_T(" GROUP BY doituong,hd_rank") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object AS doituong,hd_rank,0 socu,") \
		_T("        SUM(in_qty)       vaovien, ") \
		_T(" 		   SUM(movein_qty)   ckden, ") \
		_T(" 		   SUM(return_unit)  ravien, ") \
		_T(" 		   SUM(moveout_qty)  ckdi, ") \
		_T(" 		   SUM(transfer_qty) chuyenvien,  ") \
		_T(" 		 	 SUM(dead)         tuvong,") \
		_T("        SUM(nxv) nxv") \
		_T(" FROM(SELECT hd_object,") \
		_T("             hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             1                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          0                    return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0                    dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno)") \
		_T(" WHERE  htr_status <> 'A' %s AND htr_deptid='%s' and hcr_admitdept=htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          1                    movein_qty, ") \
		_T(" 	          0                    return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0                    dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE  htr_status <> 'A' %s AND htr_deptid='%s' AND htr_idx > 1 AND hcr_admitdept<>htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          CASE WHEN htr_suggestion not in('M','F') AND hcr_result NOT IN ( 4, 5, 6,7 ) THEN 1 ELSE 0 END return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          CASE WHEN htr_suggestion = 'F' THEN 1 ELSE 0 END transfer_qty, ") \
		_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 5, 6 ) THEN 1 ELSE 0 END dead,") \
		_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 4, 7 ) THEN 1 ELSE 0 END nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE htr_status = 'T' %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          0 return_unit, ") \
		_T(" 	          1                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0 dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE ( htr_suggestion ='M') %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s'))") \
		_T(" GROUP BY hd_object,hd_rank)") \
		_T(" GROUP BY doituong,hd_rank)") \
		_T(" WHERE doituong = 1 and hd_rank IN ( 13, 14 )") \
		_T(" GROUP BY doituong") \
		_T(" UNION ALL") \
		_T(" SELECT 1 id,'Qu\xE2n' type,'2//' doituong,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT doituong,hd_rank,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT doituong,hd_rank,") \
		_T("        count(distinct hd_docno) socu,0 vaovien,0 ckden,0 ravien,0 ckdi,0 chuyenvien,0 tuvong, 0 nxv") \
		_T(" FROM(SELECT hd_docno,doituong,hd_rank") \
		_T(" FROM(SELECT distinct hd_docno,hd_rank,") \
		_T("        hd_object AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_status='I' %s AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT distinct hd_docno,hd_rank,") \
		_T("        hd_object AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_status='T' %s AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s') AND htr_dischargedate >= cast_string2timestamp('%s')))") \
		_T(" GROUP BY doituong,hd_rank") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object AS doituong,hd_rank,0 socu,") \
		_T("        SUM(in_qty)       vaovien, ") \
		_T(" 		   SUM(movein_qty)   ckden, ") \
		_T(" 		   SUM(return_unit)  ravien, ") \
		_T(" 		   SUM(moveout_qty)  ckdi, ") \
		_T(" 		   SUM(transfer_qty) chuyenvien,  ") \
		_T(" 		 	 SUM(dead)         tuvong,") \
		_T("        SUM(nxv) nxv") \
		_T(" FROM(SELECT hd_object,") \
		_T("             hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             1                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          0                    return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0                    dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno)") \
		_T(" WHERE  htr_status <> 'A' %s AND htr_deptid='%s' and hcr_admitdept=htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          1                    movein_qty, ") \
		_T(" 	          0                    return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0                    dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE  htr_status <> 'A' %s AND htr_deptid='%s' AND htr_idx > 1 AND hcr_admitdept<>htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          CASE WHEN htr_suggestion not in('M','F') AND hcr_result NOT IN ( 4, 5, 6,7 ) THEN 1 ELSE 0 END return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          CASE WHEN htr_suggestion = 'F' THEN 1 ELSE 0 END transfer_qty, ") \
		_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 5, 6 ) THEN 1 ELSE 0 END dead,") \
		_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 4, 7 ) THEN 1 ELSE 0 END nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE htr_status = 'T' %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          0 return_unit, ") \
		_T(" 	          1                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0 dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE ( htr_suggestion ='M') %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s'))") \
		_T(" GROUP BY hd_object,hd_rank)") \
		_T(" GROUP BY doituong,hd_rank)") \
		_T(" WHERE doituong = 1 and hd_rank < 13") \
		_T(" GROUP BY doituong") \
		_T(" UNION ALL") \
		_T(" SELECT id,type,doituong,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT 2 id,'BHQH' type,'T\x1B0\x1EDBng' doituong,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT doituong,hd_rank,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT doituong,hd_rank,") \
		_T("        count(distinct hd_docno) socu,0 vaovien,0 ckden,0 ravien,0 ckdi,0 chuyenvien,0 tuvong, 0 nxv") \
		_T(" FROM(SELECT hd_docno,doituong,hd_rank") \
		_T(" FROM(SELECT distinct hd_docno,hd_rank,") \
		_T("        hd_object AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_status='I' %s AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT distinct hd_docno,hd_rank,") \
		_T("        hd_object AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_status='T' %s AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s') AND htr_dischargedate >= cast_string2timestamp('%s')))") \
		_T(" GROUP BY doituong,hd_rank") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object AS doituong,hd_rank,0 socu,") \
		_T("        SUM(in_qty)       vaovien, ") \
		_T(" 		   SUM(movein_qty)   ckden, ") \
		_T(" 		   SUM(return_unit)  ravien, ") \
		_T(" 		   SUM(moveout_qty)  ckdi, ") \
		_T(" 		   SUM(transfer_qty) chuyenvien,  ") \
		_T(" 		 	 SUM(dead)         tuvong,") \
		_T("        SUM(nxv) nxv") \
		_T(" FROM(SELECT hd_object,") \
		_T("             hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             1                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          0                    return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0                    dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno)") \
		_T(" WHERE  htr_status <> 'A' %s AND htr_deptid='%s' and hcr_admitdept=htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          1                    movein_qty, ") \
		_T(" 	          0                    return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0                    dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE  htr_status <> 'A' %s AND htr_deptid='%s' AND htr_idx > 1 AND hcr_admitdept<>htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          CASE WHEN htr_suggestion not in('M','F') AND hcr_result NOT IN ( 4, 5, 6,7 ) THEN 1 ELSE 0 END return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          CASE WHEN htr_suggestion = 'F' THEN 1 ELSE 0 END transfer_qty, ") \
		_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 5, 6 ) THEN 1 ELSE 0 END dead,") \
		_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 4, 7 ) THEN 1 ELSE 0 END nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE htr_status = 'T' %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          0 return_unit, ") \
		_T(" 	          1                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0 dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE ( htr_suggestion ='M') %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s'))") \
		_T(" GROUP BY hd_object,hd_rank)") \
		_T(" GROUP BY doituong,hd_rank)") \
		_T(" WHERE doituong IN(2,12) and hd_rank >=15") \
		_T(" GROUP BY doituong") \
		_T(" UNION ALL") \
		_T(" SELECT 2 id,'BHQH' type,'4//' doituong,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT doituong,hd_rank,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT doituong,hd_rank,") \
		_T("        count(distinct hd_docno) socu,0 vaovien,0 ckden,0 ravien,0 ckdi,0 chuyenvien,0 tuvong, 0 nxv") \
		_T(" FROM(SELECT hd_docno,doituong,hd_rank") \
		_T(" FROM(SELECT distinct hd_docno,hd_rank,") \
		_T("        hd_object AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') %s AND htr_status='I' AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT distinct hd_docno,hd_rank,") \
		_T("        hd_object AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_status='T' %s AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s') AND htr_dischargedate >= cast_string2timestamp('%s')))") \
		_T(" GROUP BY doituong,hd_rank") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object AS doituong,hd_rank,0 socu,") \
		_T("        SUM(in_qty)       vaovien, ") \
		_T(" 		   SUM(movein_qty)   ckden, ") \
		_T(" 		   SUM(return_unit)  ravien, ") \
		_T(" 		   SUM(moveout_qty)  ckdi, ") \
		_T(" 		   SUM(transfer_qty) chuyenvien,  ") \
		_T(" 		 	 SUM(dead)         tuvong,") \
		_T("        SUM(nxv) nxv") \
		_T(" FROM(SELECT hd_object,") \
		_T("             hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             1                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          0                    return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0                    dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno)") \
		_T(" WHERE  htr_status <> 'A' %s AND htr_deptid='%s' and hcr_admitdept=htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          1                    movein_qty, ") \
		_T(" 	          0                    return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0                    dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE  htr_status <> 'A' %s AND htr_deptid='%s' AND htr_idx > 1 AND hcr_admitdept<>htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          CASE WHEN htr_suggestion not in('M','F') AND hcr_result NOT IN ( 4, 5, 6,7 ) THEN 1 ELSE 0 END return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          CASE WHEN htr_suggestion = 'F' THEN 1 ELSE 0 END transfer_qty, ") \
		_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 5, 6 ) THEN 1 ELSE 0 END dead,") \
		_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 4, 7 ) THEN 1 ELSE 0 END nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE htr_status = 'T' %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          0 return_unit, ") \
		_T(" 	          1                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0 dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE ( htr_suggestion ='M') %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s'))") \
		_T(" GROUP BY hd_object,hd_rank)") \
		_T(" GROUP BY doituong,hd_rank)") \
		_T(" WHERE doituong IN(2,12) and hd_rank IN ( 13, 14 )") \
		_T(" GROUP BY doituong") \
		_T(" UNION ALL") \
		_T(" SELECT 2 id,'BHQH' type,'2//' doituong,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT doituong,hd_rank,sum(socu) socu,sum(vaovien) vaovien,sum(ckden) ckden,sum(ravien) ravien,sum(ckdi) ckdi,sum(chuyenvien) chuyenvien,sum(tuvong) tuvong,sum(nxv) nxv,sum((socu+vaovien+ckden)-(ravien+ckdi+chuyenvien+tuvong+nxv)) hientai") \
		_T(" FROM(SELECT doituong,hd_rank,") \
		_T("        count(distinct hd_docno) socu,0 vaovien,0 ckden,0 ravien,0 ckdi,0 chuyenvien,0 tuvong, 0 nxv") \
		_T(" FROM(SELECT hd_docno,doituong,hd_rank") \
		_T(" FROM(SELECT distinct hd_docno,hd_rank,") \
		_T("        hd_object AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_status='I' %s AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT distinct hd_docno,hd_rank,") \
		_T("        hd_object AS doituong,") \
		_T("        trunc_date(htr_admitdate) AS admitdate,") \
		_T("        trunc_date(hcr_dischargedate) AS dischargedate,") \
		_T("        htr_deptid") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_clinical_record ON(hcr_docno=hd_docno)") \
		_T(" LEFT JOIN hms_treatment_record ON(htr_docno =hd_docno)") \
		_T(" WHERE htr_deptid not in('A4') AND htr_status='T' %s AND htr_deptid='%s' AND htr_admitdate < Cast_string2timestamp('%s') AND htr_dischargedate >= cast_string2timestamp('%s')))") \
		_T(" GROUP BY doituong,hd_rank") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object AS doituong,hd_rank,0 socu,") \
		_T("        SUM(in_qty)       vaovien, ") \
		_T(" 		   SUM(movein_qty)   ckden, ") \
		_T(" 		   SUM(return_unit)  ravien, ") \
		_T(" 		   SUM(moveout_qty)  ckdi, ") \
		_T(" 		   SUM(transfer_qty) chuyenvien,  ") \
		_T(" 		 	 SUM(dead)         tuvong,") \
		_T("        SUM(nxv) nxv") \
		_T(" FROM(SELECT hd_object,") \
		_T("             hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             1                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          0                    return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0                    dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno)") \
		_T(" WHERE  htr_status <> 'A' %s AND htr_deptid='%s' and hcr_admitdept=htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          1                    movein_qty, ") \
		_T(" 	          0                    return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0                    dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE  htr_status <> 'A' %s AND htr_deptid='%s' AND htr_idx > 1 AND hcr_admitdept<>htr_deptid AND htr_admitdate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          CASE WHEN htr_suggestion not in('M','F') AND hcr_result NOT IN ( 4, 5, 6,7 ) THEN 1 ELSE 0 END return_unit, ") \
		_T(" 	          0                    moveout_qty, ") \
		_T(" 	          CASE WHEN htr_suggestion = 'F' THEN 1 ELSE 0 END transfer_qty, ") \
		_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 5, 6 ) THEN 1 ELSE 0 END dead,") \
		_T("             CASE WHEN htr_suggestion = 'T' AND hcr_result IN ( 4, 7 ) THEN 1 ELSE 0 END nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE htr_status = 'T' %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s')") \
		_T(" UNION ALL") \
		_T(" SELECT hd_object,hd_rank,") \
		_T("             Trunc(htr_admitdate) admit_date, ") \
		_T("             0                    in_qty, ") \
		_T(" 	          0                    movein_qty, ") \
		_T(" 	          0 return_unit, ") \
		_T(" 	          1                    moveout_qty, ") \
		_T(" 	          0                    transfer_qty, ") \
		_T(" 	          0 dead,") \
		_T("             0 nxv") \
		_T(" FROM   hms_doc ") \
		_T(" LEFT JOIN  hms_treatment_record ON(htr_docno=hd_docno) ") \
		_T(" LEFT JOIN hms_clinical_record ON ( hcr_docno = hd_docno )") \
		_T(" WHERE ( htr_suggestion ='M') %s AND htr_deptid='%s' AND htr_dischargedate BETWEEN Cast_string2timestamp('%s') AND Cast_string2timestamp('%s'))") \
		_T(" GROUP BY hd_object,hd_rank)") \
		_T(" GROUP BY doituong,hd_rank)") \
		_T(" WHERE doituong IN(2,12) and hd_rank < 13") \
		_T(" GROUP BY doituong)") \
		_T(" GROUP BY id,type,doituong)"),szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szFromDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szFromDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szFromDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szFromDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szFromDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szFromDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate,
									szWhere,pMF->GetCurrentDepartmentID(),m_szFromDate,m_szToDate);
	return szSQL;
}