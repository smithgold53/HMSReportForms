#include "stdafx.h"
#include "FMExaminationListReport.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"
#include "Excel.h"

/*static void _OnFromDateChangeFnc(CWnd *pWnd){
	((CFMExaminationListReport *)pWnd)->OnFromDateChange();
} */
/*static void _OnFromDateSetfocusFnc(CWnd *pWnd){
	((CFMExaminationListReport *)pWnd)->OnFromDateSetfocus();} */ 
/*static void _OnFromDateKillfocusFnc(CWnd *pWnd){
	((CFMExaminationListReport *)pWnd)->OnFromDateKillfocus();
} */
static int _OnFromDateCheckValueFnc(CWnd *pWnd){
	return ((CFMExaminationListReport *)pWnd)->OnFromDateCheckValue();
} 
/*static void _OnToDateChangeFnc(CWnd *pWnd){
	((CFMExaminationListReport *)pWnd)->OnToDateChange();
} */
/*static void _OnToDateSetfocusFnc(CWnd *pWnd){
	((CFMExaminationListReport *)pWnd)->OnToDateSetfocus();} */ 
/*static void _OnToDateKillfocusFnc(CWnd *pWnd){
	((CFMExaminationListReport *)pWnd)->OnToDateKillfocus();
} */
static int _OnToDateCheckValueFnc(CWnd *pWnd){
	return ((CFMExaminationListReport *)pWnd)->OnToDateCheckValue();
} 
static void _OnSoldierSelectFnc(CWnd *pWnd){
	 ((CFMExaminationListReport*)pWnd)->OnSoldierSelect();
}
static void _OnRetiredSelectFnc(CWnd *pWnd){
	 ((CFMExaminationListReport*)pWnd)->OnRetiredSelect();
}
static void _OnPrintSelectFnc(CWnd *pWnd){
	CFMExaminationListReport *pVw = (CFMExaminationListReport *)pWnd;
	pVw->OnPrintSelect();
} 
static void _OnExportSelectFnc(CWnd *pWnd){
	CFMExaminationListReport *pVw = (CFMExaminationListReport *)pWnd;
	pVw->OnExportSelect();
} 
static int _OnAddFMExaminationListReportFnc(CWnd *pWnd){
	 return ((CFMExaminationListReport*)pWnd)->OnAddFMExaminationListReport();
} 
static int _OnEditFMExaminationListReportFnc(CWnd *pWnd){
	 return ((CFMExaminationListReport*)pWnd)->OnEditFMExaminationListReport();
} 
static int _OnDeleteFMExaminationListReportFnc(CWnd *pWnd){
	 return ((CFMExaminationListReport*)pWnd)->OnDeleteFMExaminationListReport();
} 
static int _OnSaveFMExaminationListReportFnc(CWnd *pWnd){
	 return ((CFMExaminationListReport*)pWnd)->OnSaveFMExaminationListReport();
} 
static int _OnCancelFMExaminationListReportFnc(CWnd *pWnd){
	 return ((CFMExaminationListReport*)pWnd)->OnCancelFMExaminationListReport();
} 
CFMExaminationListReport::CFMExaminationListReport(CWnd *pParent){

	m_nDlgWidth = 500;
	m_nDlgHeight = 130;
	SetDefaultValues();
}
CFMExaminationListReport::~CFMExaminationListReport(){
}
void CFMExaminationListReport::OnCreateComponents(){
	m_wndExaminationListReport.Create(this, _T("Examination List Report"), 5, 5, 490, 90);
	m_wndFromDateLabel.Create(this, _T("From Date"), 10, 30, 90, 55);
	m_wndFromDate.Create(this,95, 30, 245, 55); 
	m_wndToDateLabel.Create(this, _T("To Date"), 250, 30, 330, 55);
	m_wndToDate.Create(this,335, 30, 485, 55);
	m_wndSoldier.Create(this, _T("Soldier"), 320, 60, 400, 85);
	m_wndRetired.Create(this, _T("Retired"), 404, 61, 484, 86);
	m_wndPrint.Create(this, _T("&Print"), 335, 95, 410, 120);
	m_wndExport.Create(this, _T("&Export"), 415, 95, 490, 120);

}
void CFMExaminationListReport::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	EnableControls(FALSE);
	//EnableButtons(TRUE, 0, -1);
	m_wndFromDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndFromDate.SetCheckValue(true);
	m_wndToDate.SetMax(CDateTime(pMF->GetSysDateTime()));
	m_wndToDate.SetCheckValue(true);

}
void CFMExaminationListReport::OnSetWindowEvents(){
	//m_wndFromDate.SetEvent(WE_CHANGE, _OnFromDateChangeFnc);
	//m_wndFromDate.SetEvent(WE_SETFOCUS, _OnFromDateSetfocusFnc);
	//m_wndFromDate.SetEvent(WE_KILLFOCUS, _OnFromDateKillfocusFnc);
	m_wndFromDate.SetEvent(WE_CHECKVALUE, _OnFromDateCheckValueFnc);
	//m_wndToDate.SetEvent(WE_CHANGE, _OnToDateChangeFnc);
	//m_wndToDate.SetEvent(WE_SETFOCUS, _OnToDateSetfocusFnc);
	//m_wndToDate.SetEvent(WE_KILLFOCUS, _OnToDateKillfocusFnc);
	m_wndToDate.SetEvent(WE_CHECKVALUE, _OnToDateCheckValueFnc);
	m_wndSoldier.SetEvent(WE_CLICK, _OnSoldierSelectFnc);
	m_wndRetired.SetEvent(WE_CLICK, _OnRetiredSelectFnc);
	m_wndPrint.SetEvent(WE_CLICK, _OnPrintSelectFnc);
	m_wndExport.SetEvent(WE_CLICK, _OnExportSelectFnc);
	AddEvent(1, _T("Add	Ctrl+A"), _OnAddFMExaminationListReportFnc, 0, 'A', VK_CONTROL);
	AddEvent(2, _T("Edit	Ctrl+E"), _OnEditFMExaminationListReportFnc, 0, 'E', VK_CONTROL);
	AddEvent(3, _T("Delete	Ctrl+D"), _OnDeleteFMExaminationListReportFnc, 0, 'D', VK_CONTROL);
	AddEvent(4, _T("Save	Ctrl+S"), _OnSaveFMExaminationListReportFnc, 0, 'S', VK_CONTROL);
	AddEvent(0, _T("-"));
	AddEvent(5, _T("Cancel	Ctrl+T"), _OnCancelFMExaminationListReportFnc, 0, 'T', VK_CONTROL);

}
void CFMExaminationListReport::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndFromDate.GetDlgCtrlID(), m_szFromDate);
	DDX_TextEx(pDX, m_wndToDate.GetDlgCtrlID(), m_szToDate);
	DDX_Check(pDX, m_wndSoldier.GetDlgCtrlID(), m_bSoldier);
	DDX_Check(pDX, m_wndRetired.GetDlgCtrlID(), m_bRetired);

}
void CFMExaminationListReport::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	rs.ExecSQL(szSQL);

}
void CFMExaminationListReport::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CFMExaminationListReport::SetDefaultValues(){

	m_szFromDate.Empty();
	m_szToDate.Empty();
	m_bSoldier=FALSE;
	m_bRetired=FALSE;

}
int CFMExaminationListReport::SetMode(int nMode){
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
/*void CFMExaminationListReport::OnFromDateChange(){
	
} */
/*void CFMExaminationListReport::OnFromDateSetfocus(){
	
} */
/*void CFMExaminationListReport::OnFromDateKillfocus(){
	
} */
int CFMExaminationListReport::OnFromDateCheckValue(){
	return 0;
} 
/*void CFMExaminationListReport::OnToDateChange(){
	
} */
/*void CFMExaminationListReport::OnToDateSetfocus(){
	
} */
/*void CFMExaminationListReport::OnToDateKillfocus(){
	
} */
int CFMExaminationListReport::OnToDateCheckValue(){
	return 0;
} 
	void CFMExaminationListReport::OnSoldierSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
	void CFMExaminationListReport::OnRetiredSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
}
void CFMExaminationListReport::OnPrintSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
void CFMExaminationListReport::OnExportSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	
} 
int CFMExaminationListReport::OnAddFMExaminationListReport(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CFMExaminationListReport::OnEditFMExaminationListReport(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CFMExaminationListReport::OnDeleteFMExaminationListReport(){
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
 		OnCancelFMExaminationListReport(); 
 	}
	return 0;
}
int CFMExaminationListReport::OnSaveFMExaminationListReport(){
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
 		//OnFMExaminationListReportListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret; 
 	return 0; 
}
int CFMExaminationListReport::OnCancelFMExaminationListReport(){
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
int CFMExaminationListReport::OnFMExaminationListReportListLoadData(){
	return 0;
}
