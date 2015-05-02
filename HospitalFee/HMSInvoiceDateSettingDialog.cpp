#include <stdafx.h>
#include "HMSInvoiceDateSettingDialog.h"
#include "HMSMainFrame.h"
/*static void _OnWorkingDateChangeFnc(CWnd *pWnd){
	((CHMSInvoiceDateSettingDialog *)pWnd)->OnWorkingDateChange();
} */
/*static void _OnWorkingDateSetfocusFnc(CWnd *pWnd){
	((CHMSInvoiceDateSettingDialog *)pWnd)->OnWorkingDateSetfocus();} */ 
/*static void _OnWorkingDateKillfocusFnc(CWnd *pWnd){
	((CHMSInvoiceDateSettingDialog *)pWnd)->OnWorkingDateKillfocus();
} */
static int _OnWorkingDateCheckValueFnc(CWnd *pWnd){
	return ((CHMSInvoiceDateSettingDialog *)pWnd)->OnWorkingDateCheckValue();
} 
/*static void _OnReceiptNoChangeFnc(CWnd *pWnd){
	((CHMSInvoiceDateSettingDialog *)pWnd)->OnReceiptNoChange();
} */
/*static void _OnReceiptNoSetfocusFnc(CWnd *pWnd){
	((CHMSInvoiceDateSettingDialog *)pWnd)->OnReceiptNoSetfocus();} */ 
/*static void _OnReceiptNoKillfocusFnc(CWnd *pWnd){
	((CHMSInvoiceDateSettingDialog *)pWnd)->OnReceiptNoKillfocus();
} */
static int _OnReceiptNoCheckValueFnc(CWnd *pWnd){
	return ((CHMSInvoiceDateSettingDialog *)pWnd)->OnReceiptNoCheckValue();
} 
/*static void _OnSerialNoChangeFnc(CWnd *pWnd){
	((CHMSInvoiceDateSettingDialog *)pWnd)->OnSerialNoChange();
} */
/*static void _OnSerialNoSetfocusFnc(CWnd *pWnd){
	((CHMSInvoiceDateSettingDialog *)pWnd)->OnSerialNoSetfocus();} */ 
/*static void _OnSerialNoKillfocusFnc(CWnd *pWnd){
	((CHMSInvoiceDateSettingDialog *)pWnd)->OnSerialNoKillfocus();
} */
static int _OnSerialNoCheckValueFnc(CWnd *pWnd){
	return ((CHMSInvoiceDateSettingDialog *)pWnd)->OnSerialNoCheckValue();
} 
static void _OnOKSelectFnc(CWnd *pWnd){
	CHMSInvoiceDateSettingDialog *pVw = (CHMSInvoiceDateSettingDialog *)pWnd;
	pVw->OnOKSelect();
} 
static void _OnCloseSelectFnc(CWnd *pWnd){
	CHMSInvoiceDateSettingDialog *pVw = (CHMSInvoiceDateSettingDialog *)pWnd;
	pVw->OnCloseSelect();
} 
static int _OnAddHMSInvoiceDateSettingDialogFnc(CWnd *pWnd){
	 return ((CHMSInvoiceDateSettingDialog*)pWnd)->OnAddHMSInvoiceDateSettingDialog();
} 
static int _OnEditHMSInvoiceDateSettingDialogFnc(CWnd *pWnd){
	 return ((CHMSInvoiceDateSettingDialog*)pWnd)->OnEditHMSInvoiceDateSettingDialog();
} 
static int _OnDeleteHMSInvoiceDateSettingDialogFnc(CWnd *pWnd){
	 return ((CHMSInvoiceDateSettingDialog*)pWnd)->OnDeleteHMSInvoiceDateSettingDialog();
} 
static int _OnSaveHMSInvoiceDateSettingDialogFnc(CWnd *pWnd){
	 return ((CHMSInvoiceDateSettingDialog*)pWnd)->OnSaveHMSInvoiceDateSettingDialog();
} 
static int _OnCancelHMSInvoiceDateSettingDialogFnc(CWnd *pWnd){
	 return ((CHMSInvoiceDateSettingDialog*)pWnd)->OnCancelHMSInvoiceDateSettingDialog();
} 
CHMSInvoiceDateSettingDialog::CHMSInvoiceDateSettingDialog(CWnd *pParent):
	CGuiDialog(pParent){

	m_nDlgWidth = 305;
	m_nDlgHeight = 155;
	SetDefaultValues();
}
CHMSInvoiceDateSettingDialog::~CHMSInvoiceDateSettingDialog(){
}
void CHMSInvoiceDateSettingDialog::OnCreateComponents(){
	m_wndWorkingDateLabel.Create(this, _T("Working Date"), 10, 30, 160, 55);
	m_wndWorkingDate.Create(this,165, 30, 290, 55); 
	m_wndSerialNoLabel.Create(this, _T("Serial No"), 10, 60, 160, 85);
	m_wndSerialNo.Create(this,165, 60, 290, 85); 
	m_wndReceiptNoLabel.Create(this, _T("Receipt No"), 10, 90, 160, 115);
	m_wndReceiptNo.Create(this,165, 90, 290, 115); 
	m_wndInvoiceDateSetting.Create(this, _T("Invoice Date Setting"), 5, 5, 295, 120);
	m_wndOK.Create(this, _T("&OK"), 120, 125, 205, 150);
	m_wndClose.Create(this, _T("&Close"), 210, 125, 295, 150);


}
void CHMSInvoiceDateSettingDialog::OnInitializeComponents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndWorkingDate.SetMax(CDate(pMF->GetSysDate()));
	m_wndWorkingDate.SetCheckValue(true);
	m_wndReceiptNo.SetLimitText(22);
	m_wndReceiptNo.SetCheckValue(true);
	m_wndSerialNo.SetLimitText(15);
	m_wndSerialNo.SetCheckValue(true);
	m_hms_fee_controlTbl.SetTableName(_T("hms_fee_control"));
	m_hms_fee_controlTbl.AddField(_T("HFC_USERID"), dfText, 15); 
	m_hms_fee_controlTbl.AddField(_T("HFC_RECVDATE"), dfDate); 
	m_hms_fee_controlTbl.AddField(_T("HFC_EXAMFEE"), dfText, 1); 
	m_hms_fee_controlTbl.AddField(_T("HFC_DRUGFEE"), dfText, 1); 
	m_hms_fee_controlTbl.AddField(_T("HFC_PCMSFEE"), dfText, 1); 
	m_hms_fee_controlTbl.AddField(_T("HFC_DEPOSIT"), dfText, 1); 
	m_hms_fee_controlTbl.AddField(_T("HFC_REFUND"), dfText, 1); 
	m_hms_fee_controlTbl.AddField(_T("HFC_REMISSION"), dfText, 1); 
	m_hms_fee_controlTbl.AddField(_T("HFC_OBJECTS"), dfText, 35); 
	m_hms_fee_controlTbl.AddField(_T("HFC_OUTPATIENT"), dfText, 1); 
	m_hms_fee_controlTbl.AddField(_T("HFC_INPATIENT"), dfText, 1); 
	m_hms_fee_controlTbl.AddField(_T("HFC_RECEIPTNO"), dfLong); 
	m_hms_fee_controlTbl.AddField(_T("HFC_SERIALNO"), dfText, 15); 

}
void CHMSInvoiceDateSettingDialog::OnSetWindowEvents(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//m_wndWorkingDate.SetEvent(WE_CHANGE, _OnWorkingDateChangeFnc);
	//m_wndWorkingDate.SetEvent(WE_SETFOCUS, _OnWorkingDateSetfocusFnc);
	//m_wndWorkingDate.SetEvent(WE_KILLFOCUS, _OnWorkingDateKillfocusFnc);
	m_wndWorkingDate.SetEvent(WE_CHECKVALUE, _OnWorkingDateCheckValueFnc);
	//m_wndReceiptNo.SetEvent(WE_CHANGE, _OnReceiptNoChangeFnc);
	//m_wndReceiptNo.SetEvent(WE_SETFOCUS, _OnReceiptNoSetfocusFnc);
	//m_wndReceiptNo.SetEvent(WE_KILLFOCUS, _OnReceiptNoKillfocusFnc);
	m_wndReceiptNo.SetEvent(WE_CHECKVALUE, _OnReceiptNoCheckValueFnc);
	//m_wndSerialNo.SetEvent(WE_CHANGE, _OnSerialNoChangeFnc);
	//m_wndSerialNo.SetEvent(WE_SETFOCUS, _OnSerialNoSetfocusFnc);
	//m_wndSerialNo.SetEvent(WE_KILLFOCUS, _OnSerialNoKillfocusFnc);
	m_wndSerialNo.SetEvent(WE_CHECKVALUE, _OnSerialNoCheckValueFnc);
	m_wndOK.SetEvent(WE_CLICK, _OnOKSelectFnc);
	m_wndClose.SetEvent(WE_CLICK, _OnCloseSelectFnc);
	GetDataToScreen();
	SetMode(VM_EDIT);

}
void CHMSInvoiceDateSettingDialog::OnDoDataExchange(CDataExchange* pDX){
	DDX_TextEx(pDX, m_wndWorkingDate.GetDlgCtrlID(), m_szWorkingDate);
	DDX_Text(pDX, m_wndReceiptNo.GetDlgCtrlID(), m_nReceiptNo);
	DDX_Text(pDX, m_wndSerialNo.GetDlgCtrlID(), m_szSerialNo);

}
void CHMSInvoiceDateSettingDialog::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM hms_fee_control where hfc_userid='%s' "), pMF->GetCurrentUser());
	rs.ExecSQL(szSQL);
	if(!rs.IsEOF())
	{
		rs.GetValue(_T("HFC_RECVDATE"), m_szWorkingDate);
		rs.GetValue(_T("HFC_RECEIPTNO"), m_nReceiptNo);
		rs.GetValue(_T("HFC_SERIALNO"), m_szSerialNo);
	}
	m_szWorkingDate.Empty();
}
void CHMSInvoiceDateSettingDialog::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_hms_fee_controlTbl.SetValue(_T("HFC_USERID"), pMF->GetCurrentUser());
	m_hms_fee_controlTbl.SetValue(_T("HFC_RECVDATE"), m_szWorkingDate);
	m_hms_fee_controlTbl.SetValue(_T("HFC_RECEIPTNO"), m_nReceiptNo);
	m_hms_fee_controlTbl.SetValue(_T("HFC_SERIALNO"), m_szSerialNo);

}
void CHMSInvoiceDateSettingDialog::SetDefaultValues(){

	m_szWorkingDate.Empty();
	m_nReceiptNo=0;
	m_szSerialNo.Empty();

}
int CHMSInvoiceDateSettingDialog::SetMode(int nMode){
 		int nOldMode = GetMode();
 		CGuiDialog::SetMode(nMode);
 		CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 		CString szSQL;
 		CRecord rs(&pMF->m_db);
  		switch(nMode){
 		case VM_ADD: 
 			EnableControls(TRUE);
 			EnableButtons(TRUE, 0, 1, -1);
 			SetDefaultValues();
 			break;
 		case VM_EDIT: 
 			EnableControls(TRUE);
 			EnableButtons(TRUE, 0, 1, -1);
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
/*void CHMSInvoiceDateSettingDialog::OnWorkingDateChange(){
	
} */
/*void CHMSInvoiceDateSettingDialog::OnWorkingDateSetfocus(){
	
} */
/*void CHMSInvoiceDateSettingDialog::OnWorkingDateKillfocus(){
	
} */
int CHMSInvoiceDateSettingDialog::OnWorkingDateCheckValue()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	UpdateData(TRUE);
	int n = CompareDate(m_szWorkingDate, pMF->GetSysDate());
	if(n < 0 || n > 7) 
		return -1;
	return 0;
} 
/*void CHMSInvoiceDateSettingDialog::OnReceiptNoChange(){
	
} */
/*void CHMSInvoiceDateSettingDialog::OnReceiptNoSetfocus(){
	
} */
/*void CHMSInvoiceDateSettingDialog::OnReceiptNoKillfocus(){
	
} */
int CHMSInvoiceDateSettingDialog::OnReceiptNoCheckValue(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	if(m_nReceiptNo < 0)
		return -1;
	return 0;
} 
/*void CHMSInvoiceDateSettingDialog::OnSerialNoChange(){
	
} */
/*void CHMSInvoiceDateSettingDialog::OnSerialNoSetfocus(){
	
} */
/*void CHMSInvoiceDateSettingDialog::OnSerialNoKillfocus(){
	
} */
int CHMSInvoiceDateSettingDialog::OnSerialNoCheckValue(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	if(m_szSerialNo.IsEmpty())
		m_szSerialNo = pMF->GetCurrentUser();
	UpdateData(FALSE);
	return 0;
} 
void CHMSInvoiceDateSettingDialog::OnOKSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	OnSaveHMSInvoiceDateSettingDialog();
} 
void CHMSInvoiceDateSettingDialog::OnCloseSelect(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CGuiDialog::OnCancel();
} 
int CHMSInvoiceDateSettingDialog::OnAddHMSInvoiceDateSettingDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_ADD);
	return 0;
}
int CHMSInvoiceDateSettingDialog::OnEditHMSInvoiceDateSettingDialog(){
 	if(GetMode() != VM_VIEW)
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
 	SetMode(VM_EDIT);
	return 0;  
}
int CHMSInvoiceDateSettingDialog::OnDeleteHMSInvoiceDateSettingDialog(){
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
 		OnCancelHMSInvoiceDateSettingDialog();
 	}
	return 0;
}
int CHMSInvoiceDateSettingDialog::OnSaveHMSInvoiceDateSettingDialog(){
 	if(GetMode() != VM_ADD && GetMode() != VM_EDIT)
 		return -1;
 	if(!IsValidateData())
 		return -1;
 	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd();
	CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT count(*) FROM hms_fee_control WHERE hfc_userid='%s' "), pMF->GetCurrentUser());
	rs.ExecSQL(szSQL);
	if(rs.GetIntValue() > 0)
	{
		CString szWhere;
		szWhere.Format(_T(" WHERE hfc_userid='%s' "), pMF->GetCurrentUser());
 		szSQL = m_hms_fee_controlTbl.GetUpdateSQL(_T("HFC_USERID"));
 		szSQL += szWhere;
	}
	else
	{
		szSQL = m_hms_fee_controlTbl.GetInsertSQL();
	}
	
 	
 
	int ret = pMF->ExecSQL(szSQL);

 	if(ret > 0)
 	{
		CGuiDialog::OnOK();
 	}
 	else
 	{
 	}
 	return ret;
 	return 0;
}
int CHMSInvoiceDateSettingDialog::OnCancelHMSInvoiceDateSettingDialog(){
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
int CHMSInvoiceDateSettingDialog::OnHMSInvoiceDateSettingDialogListLoadData(){
	return 0;
}
