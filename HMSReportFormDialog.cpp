#include "stdafx.h"
#include "HMSMainFrame.h"
#include "HMSReportFormDialog.h"


//Khai bao cac file header report o day
//Reception
#include "./Reception/RMRegistrationPatientListReport.h"
#include "./Reception/RMGeneralExaminationInformationReport.h"
#include "./Reception/RMGeneralExaminationFeeReport.h"
#include "./Reception/RMTransferPatientStatistic.h"
//Examination
#include "./Examination/EMAgeDiseaseStatisticsReport.h"
#include "./Examination/EMBaoCaoTongHopThangTheoPK.h"
#include "./Examination/EMBCHangThangKhoaKhamC1_1.h"
#include "./Examination/EMDanhsachBNVaoVien.h"
#include "./Examination/EMDrugDeliveryListReport.h"
#include "./Examination/EMExaminationActivitiesReport.h"
#include "./Examination/EMExaminationListReport.h"
#include "./Examination/EMExaminationTreatmentActReport.h"
#include "./Examination/EMGeneralExaminationReportDialog.h"
#include "./Examination/EMInwardListReport.h"
#include "./Examination/EMInwardStatisticForDepartmentReport.h"
#include "./Examination/EMInwardStatisticReport.h"
#include "./Examination/EMObjectDiseaseStatisticsReport.h"
#include "./Examination/EMStatisticsGeneralPatientReport.h"
#include "./Examination/EMStatisticsMajorDiseasesReport.h"
#include "./Examination/EMWeekSynthesisReport.h"
#include "./Examination/EMExamListReportC13.h"
#include "./Examination/EMQtyAtExam.h"
#include "./Examination/EMSoluongkhamcacngaytrongtuan.h"
#include "./Examination/EMSoluongbenhnhankhamchuyenkhoa.h"
#include "./Examination/EMKythuatlamtaikhoaC12.h"
#include "./Examination/EMTestIncome.h"
#include "./Examination/EMTestnParaRatebyRoom.h"
#include "./Examination/EMAdmitRatebyRoom.h"
#include "./Examination/EMPatientRatebyWardType.h"
#include "./Examination/EMPatientStatbyObj.h"
#include "./Examination/EMTestIncomeAverage.h"
#include "./Examination/EMOrderRate.h"
#include "./Examination/EMMonthlyDrugIncome.h"
#include "./Examination/EMExpensiveOrder.h"
#include "./Examination/EMRexamPatient.h"
#include "./Examination/EMStatisticsMajorDiseasesReportC12.h"
#include "./Examination/EMDanhSachBenhNhanVaoVien.h"
#include "./Examination/EMObject11DiseaseStatisticsReport.h"
#include "./Examination/EMTestMoneybyPatient.h"                                                                                            
#include "./Examination/EMGeneralSoldierExamination.h"
#include "./Examination/EMDiseasebyMachine.h"
#include "./Examination/EMOperationFosteringList.h"
//#include "./Examination/EMXTimeAverage.h"
#include "./Examination/EMServiceDrugTrack.h"
#include "./Examination/EMPatientStatbyLocation.h"
#include "./Examination/EMRoomReceptbyUser.h"
#include "./Examination/EMGeneralTestIncome.h"
#include "./Examination/EMInsuranceDrugAmount.h"
#include "./Examination/EMGeneralInsuranceExpense.h"
#include "./Examination/EMR99MoneySynthesis.h"
#include "./Examination/EMDepartmentUsageinDetail.h"
#include "./Examination/EMPrescriptionListReport.h"
#include "./Examination/EMCabinetUsage.h"
#include "./Examination/EMPatientStatbyRoom.h"
#include "./Examination/EMExpensiveOrderDetail.h"
#include "./Examination/EMTotalTestIncomeHightech.h"
#include "./Examination/EMExambySuggestion.h"
#include "./Examination/EMOperationFosteringListC12.h"

//Fee
#include "./HospitalFee/FMTempMoneyReport.h"
#include "./HospitalFee/FMServiceMedicalIncomingOutgoingReport.h"
#include "./HospitalFee/FMBookEntriesRegisterTableReport.h"
#include "./HospitalFee/FMDanhsachBNNopTien_VP.h"
#include "./HospitalFee/FMDanhSachHoaDonThuPhi.h"
#include "./HospitalFee/FMUnpaidAdvancePatientListReport.h"
#include "./HospitalFee/FMNotPayDischargeTempPatientListReport.h"
#include "./HospitalFee/FMPayMoneyPatientListReport.h"
#include "./HospitalFee/FMRationStrengthGeneralByDayReport.h"
#include "./HospitalFee/FMPaidPatientList.h"
#include "./HospitalFee/FMGeneralCollectbyDepartment.h"
#include "./HospitalFee/FMTestPatientList.h"
#include "./HospitalFee/FMRefundPatientList.h"
#include "./HospitalFee/FMCollectDetailbyItem.h"
#include "./HospitalFee/FMGeneralReportbyClerk.h"
#include "./HospitalFee/FMRefundbyDepartment.h"
#include "./HospitalFee/FMRefundDetailbyItem.h"
#include "./HospitalFee/FMAdvancePaymentList.h"
#include "./HospitalFee/FMGeneralCostInsuraceReport25aDialog.h"
#include "./HospitalFee/FMSudungthuocBHYT20a.h"
#include "./HospitalFee/FMInsuranceReport21a.h"
#include "./HospitalFee/FMInsuranceCost79A.h"
#include "./HospitalFee/FMGeneralInsuranceCost79ATH.h"
#include "./HospitalFee/FMOutpatientCostSynthesisbyDay.h"
#include "./HospitalFee/FMOutpatientCostSynthesisbyLine.h"
#include "./HospitalFee/FMOutpatientCostSynthesisbyDept.h"
#include "./HospitalFee/FMWeeklyReport.h"
#include "./HospitalFee/FMInsuranceFeeLossReport.h"
#include "./HospitalFee/FMDetailInsurance21aReport.h"
#include "./HospitalFee/FMFeeInvoiceCancelReport.h"
#include "./HospitalFee/FMDepositPatientList.h"
#include "./HospitalFee/FMDischargeDepositUnpaidReport.h"
#include "./HospitalFee/FMConditionDepositPatientList.h"
#include "./HospitalFee/FMGeneralServiceTreatmentCost.h"
#include "./HospitalFee/FMExpenseInsuranceTreatmentCost.h"
#include "./HospitalFee/FMReceiveInsuranceTreatmentCost.h"
#include "./HospitalFee/FMInsuranceCost80A.h"
#include "./HospitalFee/FMGeneralInsuranceCost80ATH.h"
#include "./HospitalFee/FMInsuranceTreatmentCostByDept.h"
#include "./HospitalFee/FMServiceTreatmentCostByStaff.h"
#include "./HospitalFee/FMServiceTreatmentCostByDay.h"
#include "./HospitalFee/FMFoodUnpaidPatientList.h"
#include "./HospitalFee/FMGeneralRationFeeByDay.h"
#include "./HospitalFee/FMFoodSummaryByDay.h"
#include "./HospitalFee/FMDetailCollectFeeByItem.h"
#include "./HospitalFee/FMUnlockPatientList.h"
#include "./HospitalFee/FMProcessedPrescriptionList.h"
#include "./HospitalFee/FMProcessedPatientList.h"
#include "./HospitalFee/FMUnpaidAdvancePatientList.h"
#include "./HospitalFee/FMUndischargedAdvancePatientList.h"
#include "./HospitalFee/FMPaidTreatmentCostSummary.h"
#include "./HospitalFee/FMServiceIncomeList.h"
#include "./HospitalFee/FMServiceOutlayList.h"
#include "./HospitalFee/FMUnlockedMoneytakerList.h"
#include "./HospitalFee/FMPostedReceiptVoucherList.h"
#include "./HospitalFee/FMPostedPaymentVoucherList.h"
#include "./HospitalFee/FMInsuranceIncomeList.h"
#include "./HospitalFee/FMInsuranceOutlayList.h"
#include "./HospitalFee/FMGeneralInsuranceIncomeList.h"
#include "./HospitalFee/FMGeneralInsuranceOutlayList.h"
#include "./HospitalFee/FMR99PaidPatientList.h"
#include "./HospitalFee/FMInsuranceReport21a1.h"
#include "./HospitalFee/FMUndischargedFoodSummary.h"
#include "./HospitalFee/FMFoodSummaryByLevel.h"
#include "./Examination/EMExamedPatientWithoutFee.h"
#include "./HospitalFee/FMInsurancePatientTechDetail.h"
#include "./HospitalFee/FMOutpatientInsuranceCost20A.h"
#include "./HospitalFee/FMOutpatientInsuranceReport21a1.h"
#include "./HospitalFee/FMOutpatientInsuranceCost80A.h"
#include "./HospitalFee/FMPaidPatientListR5x.h"
#include "./HospitalFee/FMTempSendPatientList.h"
#include "./HospitalFee/FMTempSendPatientListByDay.h"
#include "./HospitalFee/FMTempSendUnPaidPatientList.h"
#include "./HospitalFee/FMTempSendUnPaidPatientListByDay.h"
#include "./HospitalFee/FMTempSendNoOutPatientList.h"
#include "./HospitalFee/FMTempSendNoOutPatientListByDay.h"
#include "./HospitalFee/FMDepositUnpaidOutedPatientList.h"
#include "./HospitalFee/FMDepositUnpaidOutedPatientListByDay.h"
#include "./HospitalFee/FMAdmitedPatientList.h"
#include "./HospitalFee/FMAdmitedPatientListByDay.h"
#include "./HospitalFee/FMDischargedPatientList.h"
#include "./HospitalFee/FMDischargedPatientListByDay.h"
#include "./HospitalFee/FMDischargedPatientList.h"
#include "./HospitalFee/FMDischargedPaidPatientList.h"
#include "./HospitalFee/FMDischargedPaidPatientListByDay.h"
#include "./HospitalFee/FMDischargedUnpaidPatientList.h"
#include "./HospitalFee/FMDischargedUnpaidPatientListByDay.h"
#include "./HospitalFee/FMLyingPatientList.h"
#include "./HospitalFee/FMStatisticPaidPatientList.h"
//#include "./HospitalFee/FMExamStatistic.h"
#include "./HospitalFee/FMSudungthuocBHYT20a_Y2015.h"
#include "./HospitalFee/FMInsuranceReport21a1_Y2015.h"

//PM
#include "./Pharmacy/PMBaoCaoKiemKeKho.h"
#include "./Pharmacy/PMBangKePhieuNhap.h"
#include "./Pharmacy/PMBangGiaBanThuoc.h"
#include "./Pharmacy/PMBangKeChiTietNoPhaiThu.h"
#include "./Pharmacy/PMBangKeChiTietNoPhaiTra.h"
#include "./Pharmacy/PMBangKeNoPhaiThu.h"
#include "./Pharmacy/PMBangKeNoPhaiTra.h"
#include "./Pharmacy/PMPurchaseOrderReportDialog.h"
#include "./Pharmacy/PMDonthuoctonghop.h"
#include "./Pharmacy/PMPrescriptionListReportDialog.h"
#include "./Pharmacy/PMReportConditionDialog.h"
#include "./Pharmacy/PMPurchaseOrderReportDialog.h"
#include "./Pharmacy/PMSummaryImportExportInstockForStocksReportDialog.h"
#include "./Pharmacy/PMBaocaoxuatthuocchokhoatheodoituong.h"
#include "./Pharmacy/PMStockCardReportDialog.h"
#include "./Pharmacy/PMSC11AidPrescriptionList.h"
//#include "./Pharmacy/PMBaocaosudungthuochuongthan.h"
#include "./Pharmacy/PMSExportSheetList.h"
#include "./Pharmacy/PMSServiceDrugUseofOutpatient.h"
#include "./Pharmacy/PMSServiceDrugUseofInpatient.h"
#include "./Pharmacy/PMSDrugDetailBook.h"
#include "./Pharmacy/DrugLevelPermissionRpt.h"
#include "./Pharmacy/PMSDietPatientList.h"
#include "./Pharmacy/PMSGeneralDepartmentUsage.h"
#include "./Pharmacy/PMSDepartmentUsageinDetail.h"
#include "./Pharmacy/PMSImportSheetList.h"
#include "./Pharmacy/PMSGeneralStockExport.h"
#include "./Pharmacy/PMSSupplierPaymentList.h"
#include "./Pharmacy/PMVoucherList.h"
#include "./Pharmacy/PMPaidSupplierList.h"
#include "./Pharmacy/PMDrugListwithoutInvoice.h"
#include "./Pharmacy/DepartmentReturnList.h"
#include "./Pharmacy/MAMaterialUsage.h"
#include "./Pharmacy/PMSInwardExportSheetList.h"
#include "./Pharmacy/PMSUnpaidOrderList.h"
#include "./Pharmacy/MAMaterialFinalAccount.h"
#include "./Pharmacy/MAMaterialUsagebyDrug.h"
#include "./Pharmacy/PMServiceDrugUsageDetail.h"
#include "./Pharmacy/PMReportConditionForm.h"
#include "./Pharmacy/PMSExportSheetListForInsurance.h"
#include "./Pharmacy/MAConsignmentPatientList.h"
#include "./Pharmacy/MAConsignmentProductList.h"
#include "./Pharmacy/MAMaterialUsagebyOriginDept.h"
#include "./Pharmacy/PMMaterialUsagebyOriginDept.h"
#include "./Pharmacy/PMSGeneralStockImport.h"
#include "./Pharmacy/PMSProductSupplyPlan.h"
#include "./Pharmacy/PMSoldierDrugUsage.h"
#include "./Pharmacy/PMImportExportInstockbyCategory.h"
#include "./Pharmacy/PMDetailPurchaseOrder.h"
#include "./Pharmacy/PMLiabilitiesTrack.h"
#include "./Pharmacy/PMSDepartmentUsageinDetail_108Old.h"
#include "./Pharmacy/PMSGeneralDepartmentUsage_108Old.h"
#include "./Pharmacy/MAInwardExportSheetList.h"

//LIMS
#include "./LAB/LIMSTestStaticsbyObject.h"
#include "./LIMS/LIMSPatientList.h"
#include "./LIMS/LISTestActivities.h"

//PACS
#include "./PACS/PACSEndoscopyListReport.h"
#include "./PACS/PACSFunctionalProbeListReport.h"
#include "./PACS/PACSGeneralPatientByObjectReport.h"
#include "./PACS/PACSPatientListByResultReport.h"
#include "./PACS/PACSImageActivities.h"
#include "./PACS/PACSImageFosteringList.h"
#include "./PACS/PACSXRayStatistic.h"
#include "./PACS/PACSPatientListUseContrast.h"
#include "./PACS/PACSInPatientUseChemistry.h"
#include "./PACS/PACSOutPatientUseChemistry.h"
#include "./PACS/PACSSettlementChemistryMaterial.h"
#include "./PACS/PACSFunctionalProbeListSumeryReport.h"

//TM
#include "./Treatment/AdmitPatientBook.h"
#include "./Treatment/TMTreatmentActivity.h"
#include "./Treatment/TMIOTransbookreport.h"
#include "./Treatment/TMTreatmentActivitybyDept.h"
#include "./Treatment/TMDetailIOTransbookReport.h"
#include "./Treatment/TMOperationFosteringList.h"
#include "./Treatment/TMDepartmentUsageinDetail.h"
#include "./Treatment/ArmyReportsDept.h"
#include "./Treatment/TMReportsDeptPatientList.h"
#include "./Treatment/TMPatientListDischarge.h"
#include "./Treatment/TMMaterialFinalAccount.h"
#include "./Treatment/TMTotalPatientListDischarge.h"
#include "./Treatment/TMTotalPatientListDischargeByDept.h"

//BloodBank
#include "./BB/BloodBankListService.h"
#include "./BB/BloodBankUseReportWeek.h"
#include "./BB/ReceiptListBloodBank.h"
#include "./BB/BBUnapprovedBloodBankListReportDialog.h"
#include "./BB/CBloodBankReceiptReport.h"
#include "./BB/BloodBankDetailExportReport.h"

static void _OnReportGroupSelectChangeFnc(CWnd *pWnd, int nNew){
	((CHMSReportFormDialog*)pWnd->GetParent())->OnReportGroupSelectChange(nNew);
} 
static int _OnAddHMSReportFormDialogFnc(CWnd *pWnd){
	 return ((CHMSReportFormDialog*)pWnd)->OnAddHMSReportFormDialog();
} 
static int _OnEditHMSReportFormDialogFnc(CWnd *pWnd){
	 return ((CHMSReportFormDialog*)pWnd)->OnEditHMSReportFormDialog();
} 
static int _OnDeleteHMSReportFormDialogFnc(CWnd *pWnd){
	 return ((CHMSReportFormDialog*)pWnd)->OnDeleteHMSReportFormDialog();
} 
static int _OnSaveHMSReportFormDialogFnc(CWnd *pWnd){
	 return ((CHMSReportFormDialog*)pWnd)->OnSaveHMSReportFormDialog();
} 
static int _OnCancelHMSReportFormDialogFnc(CWnd *pWnd){
	 return ((CHMSReportFormDialog*)pWnd)->OnCancelHMSReportFormDialog();
}
CHMSReportFormDialog::CHMSReportFormDialog()
{
	m_nDlgWidth = 960;
	m_nDlgHeight = 625;
	//nIndex = 0;
}
CHMSReportFormDialog::CHMSReportFormDialog(CWnd *pParent)
{
	m_nDlgWidth = 960;
	m_nDlgHeight = 625;
	//nIndex = 0;
}
CHMSReportFormDialog::~CHMSReportFormDialog(){
}
void CHMSReportFormDialog::OnCreateComponents()
{
	
	m_wndReportGroup.Create(this, CRect(5, 5, 925, 615)); 
	CListPane* pane = m_wndReportGroup.GetSubPane();
	pane->SetFontSize(9);
	
//	m_wndReportGroup.SetAutoCenter(false);
	//m_wndReportImage.Create(_T("Image"), WS_CHILD|WS_VISIBLE|SS_BITMAP|SS_CENTERIMAGE|WS_BORDER, CRect(5, 405, 955, 625), this);
	//m_wndReportImage.ShowWindow(SW_HIDE);
}
void CHMSReportFormDialog::OnInitializeComponents()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	m_szModuleID = pMF->GetModuleID();
	//m_szHospitalID = _T("BVNINHBINH");
	EnableControls(FALSE);
	//EnableButtons(TRUE, 0, -1);
	//Module tiep don
	if (m_szModuleID == _T("PM") || m_szModuleID == _T("WM") || m_szModuleID == _T("MA") || m_szModuleID == _T("LIMS") || m_szModuleID == _T("TM")|| m_szModuleID == _T("BB"))
		m_wndReportGroup.SetPaneWidth(340);
	else if (m_szModuleID == _T("EM"))
		m_wndReportGroup.SetPaneWidth(340);
	else
		m_wndReportGroup.SetPaneWidth(420);

	

//	m_wndReportGroup.SetCurSel(0);
	//CRecord rs(&pMF->m_db);
	//CString szSQL;
	//int nHospitalID=0;
	//szSQL.Format(_T("SELECT sc_id FROM sys_company LIMIT 1"));
	//
	//rs.ExecSQL(szSQL);
	//if(!rs.IsEOF()){
	//	rs.GetValue(_T("sc_id"), m_szHospitalID);
	//	nHospitalID = ToInt(m_szHospitalID);	
	//	if(nHospitalID == 1014)
	//		InitReportGroup();
	//	else
	//		MessageBox(_T("Reports for 108 Hospital!"), 0, MB_ICONASTERISK);
	//}	
	InitReportGroup();
	m_wndReportGroup.SetSingleMode(true);
}
void CHMSReportFormDialog::OnSetWindowEvents()
{
//	m_wndReportGroup.SetEventFnc(_OnReportGroupSelectChangeFnc);
	
	m_wndReportGroup.RecalcLayout();
}	
void CHMSReportFormDialog::OnDoDataExchange(CDataExchange* pDX){

}
void CHMSReportFormDialog::GetDataToScreen(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	//CRecord rs(&pMF->m_db);
	CString szSQL;
	szSQL.Format(_T("SELECT * FROM "));
	//rs.ExecSQL(szSQL);

}
void CHMSReportFormDialog::GetScreenToData(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();

}
void CHMSReportFormDialog::SetDefaultValues(){
	
}
int CHMSReportFormDialog::SetMode(int nMode){
 		int nOldMode = GetMode(); 
 		CGuiDialog::SetMode(nMode); 
 		CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 		CString szSQL; 
 		//CRecord rs(&pMF->m_db); 
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
void CHMSReportFormDialog::OnReportGroupSelectChange(int nNew)
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	if (nNew < 0)
		return;
	//CWnd *pWnd = m_wndReportGroup.GetActiveWnd();
	//CRect rcWnd;
	//CRect rcClient;
	//GetClientRect(&rcClient);
	//pWnd->GetWindowRect(&rcWnd);
	//m_rcImage.left = rcClient.left+m_wndReportGroup.GetPaneWidth()+10;
	//m_rcImage.top = rcClient.top + rcWnd.Height()+45;
	//m_rcImage.bottom = rcClient.bottom-15;
	//m_rcImage.right = rcClient.right-15;
	//
	//m_wndReportImage.SetWindowPos(NULL, m_rcImage.left, m_rcImage.top, m_rcImage.Width(), m_rcImage.Height(), SWP_SHOWWINDOW|SWP_NOZORDER);
	//CString szFileName;
	//CString szPath;
	//m_image.Destroy();
	//int nSel = m_wndReportGroup.GetCurSel();
	//if(m_imageNames.Lookup(nSel, szFileName))
	//{
	//	szPath.Format(_T("Reports\\Images\\%s"), szFileName);
	//	m_image.Load(szPath);
	//	m_wndReportImage.SetBitmap(m_image);
	//}
	//m_wndReportImage.Invalidate();
} 

int CHMSReportFormDialog::OnAddHMSReportFormDialog(){
 	if(GetMode() == VM_ADD || GetMode() == VM_EDIT)  
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_ADD);
	return 0; 
}
int CHMSReportFormDialog::OnEditHMSReportFormDialog(){
 	if(GetMode() != VM_VIEW) 
 		return -1; 
 	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd(); 
 	SetMode(VM_EDIT);
	return 0;  
}
int CHMSReportFormDialog::OnDeleteHMSReportFormDialog(){
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
 		OnCancelHMSReportFormDialog(); 
 	}
	return 0;
}
int CHMSReportFormDialog::OnSaveHMSReportFormDialog(){
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
 		//OnHMSReportFormDialogListLoadData(); 
 		SetMode(VM_VIEW); 
 	} 
 	else 
 	{ 
 	} 
 	return ret;
 	return 0; 
}
int CHMSReportFormDialog::OnCancelHMSReportFormDialog(){
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
int CHMSReportFormDialog::OnHMSReportFormDialogListLoadData(){
	return 0;
}


void CHMSReportFormDialog::InitReportGroup()
{	
	//Kiem tra goi trong module nao
	if (m_szModuleID == _T("RM"))
	{
		CreateReceptionReports();
	}
	if (m_szModuleID == _T("EM"))
	{
		CreateExaminationReports();
	}
	if (m_szModuleID == _T("TM"))
	{
		CreateTreatmentReports();
	}
	if (m_szModuleID == _T("FM"))
	{
		CreateHospitalFeeReports();
	}
	if (m_szModuleID == _T("PM") || m_szModuleID == _T("WM"))
	{
		CreatePMSReports();
	} 
	if (m_szModuleID == _T("MA"))
	{
		CreateMAReports();
	}
	if (m_szModuleID == _T("LIMS"))
	{
		CreateLIMSReports();
	}
	if (m_szModuleID == _T("PACS"))
	{
		CreatePACSReports();
	}
	if (m_szModuleID == _T("IPA"))
	{
		CreateIPAReports();
	}
	if (m_szModuleID == _T("BB"))
	{
		CreateBBReports();
	}
}
void CHMSReportFormDialog::CreatePMSReports()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CGuiView * newReport = NULL;
	CString tmpStr, szTemp;
	tmpStr.Empty();
	szTemp.Empty();
	CDocPane::CDocPaneInfo *obj = m_wndReportGroup.AddPane(_T("Report"));
	if (obj)
	{
		TranslateString(_T("General Report Group"), szTemp);
		tmpStr.Format(_T("A. %s"), szTemp);
		obj->AddCaption(tmpStr);

		newReport = new CPMSStockCardReportDialog(&m_wndReportGroup, 1);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Stock Card"), szTemp);
		tmpStr.Format(_T("1. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSStockCardReportDialog(&m_wndReportGroup, 3);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Stock Card 2"), szTemp);
		tmpStr.Format(_T("1.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		//newReport = new CPMSReportConditionDialog(&m_wndReportGroup, 5);
		//newReport->Create(&m_wndReportGroup);
		//newReport->OnInitDialog();
		//TranslateString(_T("Bien ban kiem ke kho"), szTemp);
		//tmpStr.Format(_T("2. %s"), szTemp);
		//obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMReportConditionForm(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Bien ban kiem ke kho"), szTemp);
		tmpStr.Format(_T("2. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSDrugDetailBook(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Item Detail Book"), szTemp);
		tmpStr.Format(_T("3. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSStockCardReportDialog(&m_wndReportGroup, 2);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Unapproved Item List"), szTemp);
		tmpStr.Format(_T("4. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSStockCardReportDialog(&m_wndReportGroup, 4);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Unapproved Item List 2"), szTemp);
		tmpStr.Format(_T("4.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		// Nhom bc nhap kho
		TranslateString(_T("Import Drug Report Group"), szTemp);
		tmpStr.Format(_T("B. %s"), szTemp);
		obj->AddCaption(tmpStr);
		newReport = new CPMSPurchaseOrderReportDialog(&m_wndReportGroup, 1);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Purchase Order"), szTemp);
		tmpStr.Format(_T("5. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSGeneralStockImport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Stock Import"), szTemp);
		tmpStr.Format(_T("5.1- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSPurchaseOrderReportDialog(&m_wndReportGroup, 2);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Tong hop phieu nhap kho theo nha cung cap"), szTemp);
		tmpStr.Format(_T("6. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSPurchaseOrderReportDialog(&m_wndReportGroup, 3);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Detail Purchase Order"), szTemp);
		tmpStr.Format(_T("7. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMDetailPurchaseOrder(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Detail Purchase Order 2"), szTemp);
		tmpStr.Format(_T("7.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSPurchaseOrderReportDialog(&m_wndReportGroup, 4);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Purchase Order Control Record"), szTemp);
		tmpStr.Format(_T("8. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSPurchaseOrderReportDialog(&m_wndReportGroup, 5);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Item Import Statistic"), szTemp);
		tmpStr.Format(_T("9. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMDrugListwithoutInvoice(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Drug List without Invoice"), szTemp);
		tmpStr.Format(_T("10. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		//Nhap xuat kiem ke thuoc
		TranslateString(_T("Group for import - export and inventory reports"), szTemp);
		tmpStr.Format(_T("C. %s"), szTemp);
		obj->AddCaption(tmpStr);

		newReport = new CPMSReportConditionDialog(&m_wndReportGroup, 3);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Import Export Instock For Money Report"), szTemp);
		tmpStr.Format(_T("11. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
		
		newReport = new CPMSProductSupplyPlan(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Product Supply Plan"), szTemp);
		tmpStr.Format(_T("11.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CrptBaocaoxuatthuocchokhoatheodoituong(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General export to department"), szTemp);
		tmpStr.Format(_T("12. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSSummaryImportExportInstockForStocksReportDialog(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Summary of import-export-instock for multi-stock report"), szTemp);
		tmpStr.Format(_T("13. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMImportExportInstockbyCategory(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Import Export Instock by Category"), szTemp);
		tmpStr.Format(_T("13.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CDepartmentReturnList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Department Return List"), szTemp);
		tmpStr.Format(_T("14. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		// Nhom bc su dung thuoc
		TranslateString(_T("Item Use Report Group"), szTemp);
		tmpStr.Format(_T("D. %s"), szTemp);
		obj->AddCaption(tmpStr);
		
		newReport = new CPMSPrescriptionListReportDialog(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Prescription List Report"), szTemp);
		tmpStr.Format(_T("15. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new rptDonthuoctonghop(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Don thuoc tong hop"), szTemp);
		tmpStr.Format(_T("16. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSExportSheetList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Export Sheet List"), szTemp);
		tmpStr.Format(_T("17. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSExportSheetListForInsurance(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Export Sheet List For Insurance"), szTemp);
		tmpStr.Format(_T("17.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSInwardExportSheetList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Inward Export Sheet List"), szTemp);
		tmpStr.Format(_T("18. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSServiceDrugUseofOutpatient(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Service Drug Use of Outpatient"), szTemp);
		tmpStr.Format(_T("19. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSServiceDrugUseofInpatient(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Service Drug Use of Inpatient"), szTemp);
		tmpStr.Format(_T("19.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
		//20A BHYT Vien Phi
		newReport = new CFMSudungthuocBHYT20a(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Drug by quarter of a year"), szTemp);
		tmpStr.Format(_T("20- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSC11AidPrescriptionList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("C1.1 Aid Prescription List"), szTemp);
		tmpStr.Format(_T("21- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);	
		
		if (pMF->GetCurrentUser() == _T("duoc"))
		{
			newReport = new CDrugLevelPermissionRpt(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Drug Level Permission"), szTemp);
			tmpStr.Format(_T("21- %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);	
		}

/*			newReport = new CPMSDietPatientList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Diet Patient List"), szTemp);
		tmpStr.Format(_T("20- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);*/	

		newReport = new CPMSGeneralDepartmentUsage(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Department Usage"), szTemp);
		tmpStr.Format(_T("22- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);	

		newReport = new CPMSGeneralDepartmentUsage_108Old(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Department Usage 2"), szTemp);
		tmpStr.Format(_T("22.1- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);	

		newReport = new CPMSDepartmentUsageinDetail(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Department Usage in Detail"), szTemp);
		tmpStr.Format(_T("23- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);	

		newReport = new CPMSDepartmentUsageinDetail_108Old(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Department Usage in Detail 2"), szTemp);
		tmpStr.Format(_T("23.1- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSImportSheetList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Import Sheet List"), szTemp);
		tmpStr.Format(_T("24- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);	

		newReport = new CPMSGeneralStockExport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Stock Export"), szTemp);
		tmpStr.Format(_T("25- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSoldierDrugUsage(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Soldier Drug Usage"), szTemp);
		tmpStr.Format(_T("25.2- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSSupplierPaymentList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Supplier Payment List"), szTemp);
		tmpStr.Format(_T("26- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMVoucherList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Voucher List"), szTemp);
		tmpStr.Format(_T("27- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMPaidSupplierList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Paid Supplier List"), szTemp);
		tmpStr.Format(_T("28- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSUnpaidOrderList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Unpaid Order List"), szTemp);
		tmpStr.Format(_T("29- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMLiabilitiesTrack(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Liabilities Track"), szTemp);
		tmpStr.Format(_T("29.1- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMServiceDrugUsageDetail(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Service Drug Usage Detail"), szTemp);
		tmpStr.Format(_T("30- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMMaterialUsagebyOriginDept(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Drug Usage by Origin Dept"), szTemp);
		tmpStr.Format(_T("31- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

	}
}

void CHMSReportFormDialog::CreateMAReports(){
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CGuiView * newReport = NULL;
	CString tmpStr, szTemp;
	tmpStr.Empty();
	szTemp.Empty();
	CDocPane::CDocPaneInfo *obj = m_wndReportGroup.AddPane(_T("Report"));
	if (obj)
	{
		TranslateString(_T("General Report Group"), szTemp);
		tmpStr.Format(_T("A. %s"), szTemp);
		obj->AddCaption(tmpStr);

		newReport = new CPMSStockCardReportDialog(&m_wndReportGroup, 1);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Stock Card"), szTemp);
		tmpStr.Format(_T("1. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSStockCardReportDialog(&m_wndReportGroup, 3);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Stock Card 2"), szTemp);
		tmpStr.Format(_T("1.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		//newReport = new CPMSReportConditionDialog(&m_wndReportGroup, 5);
		//newReport->Create(&m_wndReportGroup);
		//newReport->OnInitDialog();
		//TranslateString(_T("Bien ban kiem ke kho"), szTemp);
		//tmpStr.Format(_T("2. %s"), szTemp);
		//obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMReportConditionForm(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Bien ban kiem ke kho"), szTemp);
		tmpStr.Format(_T("2. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSDrugDetailBook(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Item Detail Book"), szTemp);
		tmpStr.Format(_T("3. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSStockCardReportDialog(&m_wndReportGroup, 2);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Unapproved Item List"), szTemp);
		tmpStr.Format(_T("4. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSStockCardReportDialog(&m_wndReportGroup, 4);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Unapproved Item List 2"), szTemp);
		tmpStr.Format(_T("4.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		// Nhom bc nhap kho
		TranslateString(_T("Import Drug Report Group"), szTemp);
		tmpStr.Format(_T("B. %s"), szTemp);
		obj->AddCaption(tmpStr);
		newReport = new CPMSPurchaseOrderReportDialog(&m_wndReportGroup, 1);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Purchase Order"), szTemp);
		tmpStr.Format(_T("5. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSGeneralStockImport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Stock Import"), szTemp);
		tmpStr.Format(_T("5.1- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSPurchaseOrderReportDialog(&m_wndReportGroup, 2);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Tong hop phieu nhap kho theo nha cung cap"), szTemp);
		tmpStr.Format(_T("6. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSPurchaseOrderReportDialog(&m_wndReportGroup, 3);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Detail Purchase Order"), szTemp);
		tmpStr.Format(_T("7. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMDetailPurchaseOrder(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Detail Purchase Order 2"), szTemp);
		tmpStr.Format(_T("7.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSPurchaseOrderReportDialog(&m_wndReportGroup, 4);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Purchase Order Control Record"), szTemp);
		tmpStr.Format(_T("8. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSPurchaseOrderReportDialog(&m_wndReportGroup, 5);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Item Import Statistic"), szTemp);
		tmpStr.Format(_T("9. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMDrugListwithoutInvoice(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Drug List without Invoice"), szTemp);
		tmpStr.Format(_T("10. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		//Nhap xuat kiem ke thuoc
		TranslateString(_T("Group for import - export and inventory reports"), szTemp);
		tmpStr.Format(_T("C. %s"), szTemp);
		obj->AddCaption(tmpStr);

		newReport = new CPMSReportConditionDialog(&m_wndReportGroup, 3);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Import Export Instock For Money Report"), szTemp);
		tmpStr.Format(_T("11. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
		
		newReport = new CPMSProductSupplyPlan(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Product Supply Plan"), szTemp);
		tmpStr.Format(_T("11.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CrptBaocaoxuatthuocchokhoatheodoituong(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General export to department"), szTemp);
		tmpStr.Format(_T("12. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSSummaryImportExportInstockForStocksReportDialog(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Summary of import-export-instock for multi-stock report"), szTemp);
		tmpStr.Format(_T("13. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMImportExportInstockbyCategory(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Import Export Instock by Category"), szTemp);
		tmpStr.Format(_T("13.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CDepartmentReturnList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Department Return List"), szTemp);
		tmpStr.Format(_T("14. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		// Nhom bc su dung thuoc
		TranslateString(_T("Item Use Report Group"), szTemp);
		tmpStr.Format(_T("D. %s"), szTemp);
		obj->AddCaption(tmpStr);
		
		newReport = new CPMSExportSheetList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Export Sheet List"), szTemp);
		tmpStr.Format(_T("15. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSInwardExportSheetList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Inward Export Sheet List"), szTemp);
		tmpStr.Format(_T("15.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CMAInwardExportSheetList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Inward Export Sheet List 1"), szTemp);
		tmpStr.Format(_T("15.2 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
		
		newReport = new CPMSGeneralDepartmentUsage(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Department Usage"), szTemp);
		tmpStr.Format(_T("16- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);	

		newReport = new CPMSDepartmentUsageinDetail(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Department Usage in Detail"), szTemp);
		tmpStr.Format(_T("17- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);	

		newReport = new CPMSImportSheetList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Import Sheet List"), szTemp);
		tmpStr.Format(_T("18- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);	

		newReport = new CPMSGeneralStockExport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Stock Export"), szTemp);
		tmpStr.Format(_T("19- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMSSupplierPaymentList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Supplier Payment List"), szTemp);
		tmpStr.Format(_T("20- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMVoucherList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Voucher List"), szTemp);
		tmpStr.Format(_T("21- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPMPaidSupplierList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Paid Supplier List"), szTemp);
		tmpStr.Format(_T("22- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
	
		newReport = new CMAMaterialUsage(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Material Usage"), szTemp);
		tmpStr.Format(_T("23- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CMAMaterialUsagebyOriginDept(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Material Usage by Origin Dept"), szTemp);
		tmpStr.Format(_T("23.1- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CMAMaterialFinalAccount(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Material Final Account"), szTemp);
		tmpStr.Format(_T("24- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CMAMaterialUsagebyDrug(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Material Usage by Drug"), szTemp);
		tmpStr.Format(_T("25- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CMAConsignmentPatientList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Consignment Patient List"), szTemp);
		tmpStr.Format(_T("26- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CMAConsignmentProductList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Consignment Product List"), szTemp);
		tmpStr.Format(_T("27- %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
	}
}

void CHMSReportFormDialog::CreateReceptionReports()
{
	CGuiView *newReport = NULL;
	CString tmpStr, szTemp;
	tmpStr.Empty();
	szTemp.Empty();
	CDocPane::CDocPaneInfo *obj = m_wndReportGroup.AddPane(_T("Report"));
	//1st
	if (obj)
	{
		newReport = new CRMRegistrationPatientListReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Registration Patient List"), szTemp);
		tmpStr.Format(_T("1. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
		//2nd
		newReport = new CRMGeneralExaminationInformationReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Examination Information Report"), szTemp);
		tmpStr.Format(_T("2. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
		//3rd
		newReport = new CRMGeneralExaminationFeeReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Examination Fee Report"), szTemp);
		tmpStr.Format(_T("3. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CRMTransferPatientStatistic(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Transfer Patient Statistic"), szTemp);
		tmpStr.Format(_T("4. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
	}
}
void CHMSReportFormDialog::CreateExaminationReports()
{
	CHMSMainFrame *pMF = (CHMSMainFrame*) AfxGetMainWnd();
	CGuiView *newReport = NULL;
	CString tmpStr, szTemp;
	tmpStr.Empty();
	szTemp.Empty();
	CDocPane* pDocPane = &m_wndReportGroup;
	CDocPane::CDocPaneInfo* obj=NULL;

	obj = pDocPane->AddPane(_T("Report"));
	if (obj)
	{		
		if ((pMF->GetCurrentDepartmentID() == _T("C1.1")) || (pMF->GetCurrentDepartmentID() == _T("KHTH")))
		{
			if (pMF->GetCurrentUser() == _T("bvtan"))
			{
				//11 new rpts
				tmpStr.Format(_T("New Report for C1.1 Statistic"));
				obj->AddCaption(tmpStr);

				newReport = new CEMPatientStatbyObj(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Patient Stat by Obj"), szTemp);
				tmpStr.Format(_T("1. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);
				
				newReport = new CEMPatientRatebyWardType(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Patient Rate by Ward Type"), szTemp);
				tmpStr.Format(_T("2. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMAdmitRatebyRoom(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Admit Rate by Room"), szTemp);
				tmpStr.Format(_T("3. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMTestnParaRatebyRoom(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Test n Para Rate by Room"), szTemp);
				tmpStr.Format(_T("4. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMTestIncome(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Test Income"), szTemp);
				tmpStr.Format(_T("5. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMTestIncomeAverage(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Test Income Average"), szTemp);
				tmpStr.Format(_T("6. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMGeneralTestIncome(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("General Test Income"), szTemp);
				tmpStr.Format(_T("6.1 %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);
				
				newReport = new CEMTotalTestIncomeHightech(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Total Test Income Hightech"), szTemp);
				tmpStr.Format(_T("6.2 %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMOrderRate(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Order Rate"), szTemp);
				tmpStr.Format(_T("7. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMMonthlyDrugIncome(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Monthly Drug Income"), szTemp);
				tmpStr.Format(_T("8. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMInsuranceDrugAmount(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Insurance Drug Amount"), szTemp);
				tmpStr.Format(_T("8.1 %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMExpensiveOrder(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Expensive Order"), szTemp);
				tmpStr.Format(_T("9. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMRexamPatient(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Re-exam patient"), szTemp);
				tmpStr.Format(_T("10. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);
				
				newReport = new CEMTestMoneybyPatient(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Test Money by Patient"), szTemp);
				tmpStr.Format(_T("11. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMServiceDrugTrack(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Service Drug Track"), szTemp);
				tmpStr.Format(_T("12. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMPatientStatbyLocation(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Patient Stat by Location"), szTemp);
				tmpStr.Format(_T("13. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMRoomReceptbyUser(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Room Recept by User"), szTemp);
				tmpStr.Format(_T("14. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMGeneralInsuranceExpense(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("General Insurance Expense"), szTemp);
				tmpStr.Format(_T("15. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMPatientStatbyRoom(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Patient Stat by Room"), szTemp);
				tmpStr.Format(_T("16. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMExpensiveOrderDetail(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Expensive Order Detail"), szTemp);
				tmpStr.Format(_T("17. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);
			}
			tmpStr.Format(_T("\x42\xE1o \x63\xE1o khu \x43\x31.\x31"));
			obj->AddCaption(tmpStr);
			//1st
			newReport = new CEMInwardStatisticReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Inward Statistics Report"), szTemp);
			tmpStr.Format(_T("1. %s"), szTemp);
			obj->Add(tmpStr, _T(""),  newReport);
			//2nd
			newReport = new CEMInwardStatisticForDepartmentReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Inward Statistic for Department Report"), szTemp);
			tmpStr.Format(_T("2. %s"), szTemp);
			obj->Add(tmpStr, _T(""),  newReport);
			//3rd
			newReport = new CEMExaminationListReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Examination List"), szTemp);
			tmpStr.Format(_T("3. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
			//3.1
			newReport = new CEMExamListReportC13(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Examination List C1.3"), szTemp);
			tmpStr.Format(_T("3.1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
			//4th
			newReport = new CEMInwardListReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Admission Patient List"), szTemp);
			tmpStr.Format(_T("4. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
			//5th
			newReport = new CEMStatisticsGeneralPatientReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Statistics General Patient Report"), szTemp);
			tmpStr.Format(_T("5. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
			//6th
			newReport = new CEMDrugDeliveryListReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Drug Delivery List Report"), szTemp);
			tmpStr.Format(_T("6. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
			//7th
			newReport = new CEMStatisticsMajorDiseasesReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Statistics Major Disease Report"), szTemp);
			tmpStr.Format(_T("7. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
			//8th
			newReport = new CEMObjectDiseaseStatisticsReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Object Disease Statistics Report"), szTemp);
			tmpStr.Format(_T("8. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
			//9th
			newReport = new CEMAgeDiseaseStatisticsReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Age Disease Statistics Report"), szTemp);
			tmpStr.Format(_T("9. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
			//10th
			newReport = new CEMExaminationTreatmentActReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Examination Treatment Act Report"), szTemp);
			tmpStr.Format(_T("10. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
			//11th
			newReport = new CEMGeneralReportC1_1(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("General Monthly Report by Department"), szTemp);
			tmpStr.Format(_T("11. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
			//12th
			//newReport = new CEMWeekSynthesisReport(&m_wndReportGroup);
			//newReport->Create(&m_wndReportGroup);
			//newReport->OnInitDialog();
			//TranslateString(_T("Weekly Synthesis Report"), szTemp);
			//tmpStr.Format(_T("12. %s"), szTemp);
			//obj->Add(tmpStr, _T(""), newReport);
			//13th
			newReport = new CRptBCHangThangKhoaKhamC1_1(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("C1 Monthly Report"), szTemp);
			tmpStr.Format(_T("13. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			//14th
			newReport = new CEMGeneralSoldierExamination(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("General Soldier Examination"), szTemp);
			tmpStr.Format(_T("14. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			/*if (pMF->GetCurrentDepartmentID() == _T("C15"))
			{*/
				//15th
				newReport = new CEMDiseasebyMachine(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Disease by Machine Report"), szTemp);
				tmpStr.Format(_T("15. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMOperationFosteringList(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Operation Fostering List"), szTemp);
				tmpStr.Format(_T("16. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);

				newReport = new CEMR99MoneySynthesis(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Required Room Money Synthesis"), szTemp);
				tmpStr.Format(_T("17. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);
			//}
		}
		if ((pMF->GetCurrentDepartmentID() == _T("C1.2")) || (pMF->GetCurrentDepartmentID() == _T("KHTH")))
		{
			tmpStr.Format(_T("\x42\xE1o \x63\xE1o khu \x43\x31.\x32"));
			obj->AddCaption(tmpStr);

			newReport = new CEMStatisticsGeneralPatientReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Statistics General Patient Report"), szTemp);
			tmpStr.Format(_T("1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CEMExaminationListReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Examination List"), szTemp);
			tmpStr.Format(_T("2. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
			
			newReport = new CEMDanhSachBenhNhanVaoVien(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			szTemp = _T("\x44\x61nh S\xE1\x63h \x42\x1EC7nh Nh\xE2n V\xE0o Vi\x1EC7n");
			tmpStr.Format(_T("3. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CEMDrugDeliveryListReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Drug Delivery List Report"), szTemp);
			tmpStr.Format(_T("4. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
		
			newReport = new CEMWeekSynthesisReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Weekly Synthesis Report"), szTemp);
			tmpStr.Format(_T("5. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CEMStatisticsMajorDiseasesReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Statistics Major Disease Report"), szTemp);
			tmpStr.Format(_T("6. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			// benh tat theo nhom

			newReport = new CEMStatisticsMajorDiseasesReportC12(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Th\x1ED1ng k\xEA \x62\x1EC7nh nh\xE2n th\x65o nh\xF3m m\xE3 \x62\x1EC7nh"), szTemp);
			tmpStr.Format(_T("7. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);


			newReport = new CEMObjectDiseaseStatisticsReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Object Disease Statistics Report"), szTemp);
			tmpStr.Format(_T("8. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CEMObject11DiseaseStatisticsReport(&m_wndReportGroup); 
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("\x31\x31 nh\xF3m m\x1EB7t \x62\x1EC7nh \x63h\xEDnh"), szTemp);
			tmpStr.Format(_T("9. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);


			newReport = new CEMExaminationTreatmentActReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Examination Treatment Act Report"), szTemp);
			tmpStr.Format(_T("10. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CEMQtyAtExam(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("S\x1ED1 l\x1B0\x1EE3ng kh\xE1m t\x1EA1i \x63\xE1\x63 ph\xF2ng"), szTemp);
			tmpStr.Format(_T("11. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CEMSoluongkhamcacngaytrongtuan(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			tmpStr.Format(_T("12. S\x1ED1 l\x1B0\x1EE3ng kh\xE1m \x63\xE1\x63 ng\xE0y trong tu\x1EA7n"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CEMSoluongbenhnhankhamchuyenkhoa(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			tmpStr.Format(_T("13.S\x1ED1 l\x1B0\x1EE3ng \x62\x1EC7nh nh\xE2n kh\xE1m \x63huy\xEAn kho\x61"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CEMKythuatlamtaikhoaC12(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			tmpStr.Format(_T("14. K\x1EF9 thu\x1EADt l\xE0m t\x1EA1i kho\x61"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);


			newReport = new CEMAgeDiseaseStatisticsReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			tmpStr.Format(_T("15. Th\x1ED1ng k\xEA th\x65o m\x1EB7t \x62\x1EC7nh - \x111\x1ED9 tu\x1ED5i"));
			obj->Add(tmpStr, _T(""), newReport);
		
			newReport = new CEMOperationFosteringListC12(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Operation Fostering List"), szTemp);
			tmpStr.Format(_T("16. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
		}

		if ((pMF->GetCurrentDepartmentID() == _T("C1.3")) || (pMF->GetCurrentDepartmentID() == _T("KHTH")))
		{
			tmpStr.Format(_T("\x42\xE1o \x63\xE1o khu \x43\x31.\x33"));
			obj->AddCaption(tmpStr);

			newReport = new CEMExamListReportC13(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Examination List C1.3"), szTemp);
			tmpStr.Format(_T("1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);


			newReport = new CEMInwardStatisticForDepartmentReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Inward Statistic for Department Report"), szTemp);
			tmpStr.Format(_T("2. %s"), szTemp);
			obj->Add(tmpStr, _T(""),  newReport);

			newReport = new CEMStatisticsMajorDiseasesReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Statistics Major Disease Report"), szTemp);
			tmpStr.Format(_T("3. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CEMObjectDiseaseStatisticsReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Object Disease Statistics Report"), szTemp);
			tmpStr.Format(_T("4. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CEMExaminationTreatmentActReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Examination Treatment Act Report"), szTemp);
			tmpStr.Format(_T("5. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CEMCabinetUsage(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Prescription List Report"), szTemp);
			tmpStr.Format(_T("6. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CEMDepartmentUsageinDetail(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Department Usage in Detail"), szTemp);
			tmpStr.Format(_T("7. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
			
			newReport = new CEMExambySuggestion(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Exam by Suggestion"), szTemp);
			tmpStr.Format(_T("8. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

		}

	}
}
void CHMSReportFormDialog::CreateTreatmentReports()
{
	CGuiView *newReport = NULL;
	CString tmpStr, szTemp;
	tmpStr.Empty();
	szTemp.Empty();
	CDocPane::CDocPaneInfo *obj = m_wndReportGroup.AddPane(_T("Report"));
	
	if (obj)
	{
		newReport = new CAdmitPatientBook(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Admit Patient Book"), szTemp);
		tmpStr.Format(_T("1. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CTMTreatmentActivity(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Treatment Activity"), szTemp);
		tmpStr.Format(_T("2. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CTMTreatmentActivitybyDept(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Treatment Activity by Dept"), szTemp);
		tmpStr.Format(_T("3. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
 
		newReport = new CTMIOTransbookreport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("In Out Transfer Patient Book"), szTemp);
		tmpStr.Format(_T("4. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CTMDetailIOTransbookreport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Detail In Out Transfer Patient Book"), szTemp);
		tmpStr.Format(_T("5. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CTMDepartmentUsageinDetail(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Department Usage in Detail"), szTemp);
		tmpStr.Format(_T("6. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);


		newReport = new CTMOperationFosteringList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Operation Fostering List"), szTemp);
		tmpStr.Format(_T("7. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CArmyReportsDept(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("\x42\xE1o \x63\xE1o qu\xE2n s\x1ED1"), szTemp);
		tmpStr.Format(_T("8. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CTMReportsDeptPatientList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("\x44\x61nh s\xE1\x63h \x62\x1EC7nh nh\xE2n \x111\x61ng \x111i\x1EC1u tr\x1ECB"), szTemp);
		tmpStr.Format(_T("9. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CTMPatientListDischarge(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("S\x1ED5 r\x61 vi\x1EC7n"), szTemp);
		tmpStr.Format(_T("10. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CTMMaterialFinalAccount(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Material Final Account"), szTemp);
		tmpStr.Format(_T("11. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CTMTotalPatientListDischarge(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Total Patient List Discharge"), szTemp);
		tmpStr.Format(_T("12. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CTMTotalPatientListDischargeByDept(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Total Patient List Discharge by Dept"), szTemp);
		tmpStr.Format(_T("12.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		/*newReport = new CHMSInOutPatientTrack(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("In Out Patient Track"), szTemp);
		tmpStr.Format(_T("11. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);*/
	}
}

void CHMSReportFormDialog::CreateHospitalFeeReports()
{
	CGuiView *newReport = NULL;
	CHMSMainFrame *pMF = (CHMSMainFrame *) AfxGetMainWnd();
	CString tmpStr, szTemp;

	tmpStr.Empty();
	szTemp.Empty();
	CDocPane::CDocPaneInfo *obj = m_wndReportGroup.AddPane(_T("Report"));
	if (obj)
	{

		if (!pMF->CheckPermission(_T("15.12")))
		{
			TranslateString(_T("General Report Group"), szTemp);
			tmpStr.Format(_T("A. %s"), szTemp);
			obj->AddCaption(tmpStr);

			newReport = new CRptDanhSachHoaDonThuPhi(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Charged Invoice List"), szTemp);
			tmpStr.Format(_T("1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMGeneralReportbyClerk(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("General Report by Clerk"), szTemp);
			tmpStr.Format(_T("2. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMPaidPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Paid Patient List"), szTemp);
			tmpStr.Format(_T("3. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMR99PaidPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("R99 Paid Patient List"), szTemp);
			tmpStr.Format(_T("3.1 %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMPaidPatientListR5x(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Paid Patient List R5x"), szTemp);
			tmpStr.Format(_T("3.2 %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMGeneralCollectbyDepartment(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("General Collect by Department"), szTemp);
			tmpStr.Format(_T("4. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMCollectDetailbyItem(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Collect Detail by Item"), szTemp);
			tmpStr.Format(_T("5. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMTestPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Test Patient List"), szTemp);
			tmpStr.Format(_T("6. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			/*newReport = new CFMDetailCollectFeeByItem(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Detail Collect Fee By Item"), szTemp);
			tmpStr.Format(_T("7. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);*/

			newReport = new CFMFeeInvoiceCancelReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Fee Invoice Cancel List Report"), szTemp);
			tmpStr.Format(_T("7. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);


			newReport = new CFMFoodUnpaidPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Food Unpaid Patient List"), szTemp);
			tmpStr.Format(_T("8. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);


			newReport = new CFMGeneralRationFeeByDay(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("General Ration Fee By Day"), szTemp);
			tmpStr.Format(_T("9. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);


			newReport = new CFMFoodSummaryByDay(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Food Summary By Day"), szTemp);
			tmpStr.Format(_T("10. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			

			if (pMF->GetCurrentUser() == _T("nkan") ||
				pMF->GetCurrentUser() == _T("admin"))
			{
				newReport = new CFMUnlockPatientList(&m_wndReportGroup);
				newReport->Create(&m_wndReportGroup);
				newReport->OnInitDialog();
				TranslateString(_T("Unlock Patient List Report"), szTemp);
				tmpStr.Format(_T("11. %s"), szTemp);
				obj->Add(tmpStr, _T(""), newReport);
			}
			
			newReport = new CFMUndischargedFoodSummary(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Undischarged Food Summary"), szTemp);
			tmpStr.Format(_T("12. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMFoodSummaryByLevel(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Food Summary by Level"), szTemp);
			tmpStr.Format(_T("13. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			TranslateString(_T("Advance - Refund Report Group"), szTemp);
			tmpStr.Format(_T("B. %s"), szTemp);
			obj->AddCaption(tmpStr);

			newReport = new CFMAdvancePaymentList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Advance Payment List"), szTemp);
			tmpStr.Format(_T("1. %s %s"), szTemp, _T("(M\x1EABu \x31)"));
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMDepositPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Deposit Patient List Report"), szTemp);
			tmpStr.Format(_T("   1.1 %s %s"), szTemp, _T("(M\x1EABu \x32)"));
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMRefundPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Refund Patient List"), szTemp);
			tmpStr.Format(_T("2. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMRefundbyDepartment(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Refund by Department"), szTemp);
			tmpStr.Format(_T("3. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			//newReport = new CFMRefundDetailbyItem(&m_wndReportGroup);
			//newReport->Create(&m_wndReportGroup);
			//newReport->OnInitDialog();
			//TranslateString(_T("Refund Detail by Item"), szTemp);
			//tmpStr.Format(_T("9. %s"), szTemp);
			//obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMDischargeDepositUnpaidReport(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Discharge Deposit Unpaid Report"), szTemp);
			tmpStr.Format(_T("4. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);


			//newReport = new CFMConditionDepositPatientList(&m_wndReportGroup);
			//newReport->Create(&m_wndReportGroup);
			//newReport->OnInitDialog();
			//TranslateString(_T("Inpatient Deposit List Report"), szTemp);
			//tmpStr.Format(_T("5. %s"), szTemp);
			//obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMUnpaidAdvancePatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Unpaid Advance Patient List"), szTemp);
			tmpStr.Format(_T("5. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMUndischargedAdvancePatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Undischarged Advance Patient List"), szTemp);
			tmpStr.Format(_T("6. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			//newReport = new CFMPaidTreatmentCostSummary(&m_wndReportGroup);
			//newReport->Create(&m_wndReportGroup);
			//newReport->OnInitDialog();
			//TranslateString(_T("Paid Treatment Cost Summary"), szTemp);
			//tmpStr.Format(_T("8. %s"), szTemp);
			//obj->Add(tmpStr, _T(""), newReport);

			TranslateString(_T("Nh\xF3m \x62\xE1o \x63\xE1o \x63hi ph\xED \x62\x1EC7nh nh\xE2n \x64\x1ECB\x63h v\x1EE5 n\x1ED9i tr\xFA"), szTemp);
			tmpStr.Format(_T("C. %s"), szTemp);
			obj->AddCaption(tmpStr);

			newReport = new CFMGeneralServiceTreatmentCost(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("General Service Treatment Cost Report"), szTemp);
			tmpStr.Format(_T("1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			//newReport = new CFMServiceTreatmentCostByDay(&m_wndReportGroup);
			//newReport->Create(&m_wndReportGroup);
			//newReport->OnInitDialog();
			//TranslateString(_T("General Service Treatment Cost By Day"), szTemp);
			//tmpStr.Format(_T("2. %s"), szTemp);
			//obj->Add(tmpStr, _T(""), newReport);

			//newReport = new CFMServiceTreatmentCostByStaff(&m_wndReportGroup);
			//newReport->Create(&m_wndReportGroup);
			//newReport->OnInitDialog();
			//TranslateString(_T("General Service Treatment Cost By Staff"), szTemp);
			//tmpStr.Format(_T("3. %s"), szTemp);
			//obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMServiceIncomeList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Service Income List"), szTemp);
			tmpStr.Format(_T("2. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMServiceOutlayList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Service Outlay List"), szTemp);
			tmpStr.Format(_T("3. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMUnlockedMoneytakerList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Unlocked Money-taker List"), szTemp);
			tmpStr.Format(_T("4. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMPostedReceiptVoucherList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Posted Receipt Voucher List"), szTemp);
			tmpStr.Format(_T("5. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMPostedPaymentVoucherList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Posted Payment Voucher List"), szTemp);
			tmpStr.Format(_T("6. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMTempSendPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Temp Send Patient List By Dept"), szTemp);
			tmpStr.Format(_T("7. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMTempSendPatientListByDay(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Temp Send Patient List By Day"), szTemp);
			tmpStr.Format(_T("7.1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMTempSendUnPaidPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Temp Send Unpaid Patient List"), szTemp);
			tmpStr.Format(_T("8. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMTempSendUnPaidPatientListByDay(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Temp Send Unpaid Patient List By Day"), szTemp);
			tmpStr.Format(_T("8.1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMTempSendNoOutPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Temp Send No Out Patient List"), szTemp);
			tmpStr.Format(_T("9. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMTempSendNoOutPatientListByDay(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Temp Send No Out Patient List By Day"), szTemp);
			tmpStr.Format(_T("9.1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMDepositUnpaidOutedPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Deposit Unpaid Outed Patient List"), szTemp);
			tmpStr.Format(_T("10. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMDepositUnpaidOutedPatientListByDay(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Deposit Unpaid Outed Patient List By Day"), szTemp);
			tmpStr.Format(_T("10.1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMLyingPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Lying Patient List"), szTemp);
			tmpStr.Format(_T("11. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMAdmitedPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Admited Patient List"), szTemp);
			tmpStr.Format(_T("12. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMAdmitedPatientListByDay(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Admited Patient List By Day"), szTemp);
			tmpStr.Format(_T("12.1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMDischargedPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Discharged Patient List"), szTemp);
			tmpStr.Format(_T("13. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMDischargedPatientListByDay(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Discharged Patient List By Day"), szTemp);
			tmpStr.Format(_T("13.1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMDischargedPaidPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Discharged Paid Patient List"), szTemp);
			tmpStr.Format(_T("14. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMDischargedPaidPatientListByDay(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Discharged Paid Patient List By Day"), szTemp);
			tmpStr.Format(_T("14.1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMDischargedUnpaidPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Discharged UnPaid Patient List"), szTemp);
			tmpStr.Format(_T("15. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMDischargedUnpaidPatientListByDay(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Discharged UnPaid Patient List By Day"), szTemp);
			tmpStr.Format(_T("15.1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMStatisticPaidPatientList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Statistic Paid Patient List"), szTemp);
			tmpStr.Format(_T("16. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			TranslateString(_T("Nh\xF3m \x62\xE1o \x63\xE1o \x63hi ph\xED \x62\x1EC7nh nh\xE2n \x42\x110 – \x43S – \x42H n\x1ED9i tr\xFA"), szTemp);
			tmpStr.Format(_T("D. %s"), szTemp);
			obj->AddCaption(tmpStr);

			newReport = new CFMReceiveInsuranceTreatmentCost(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Receive Insurance Treatment Cost Report"), szTemp);
			tmpStr.Format(_T("1. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMExpenseInsuranceTreatmentCost(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Expense Insurance Treatment Cost Report"), szTemp);
			tmpStr.Format(_T("2. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMInsuranceIncomeList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Insurance Income List"), szTemp);
			tmpStr.Format(_T("3. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMInsuranceOutlayList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("Insurance Outlay List"), szTemp);
			tmpStr.Format(_T("4. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMGeneralInsuranceIncomeList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("General Insurance Income List"), szTemp);
			tmpStr.Format(_T("5. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);

			newReport = new CFMGeneralInsuranceOutlayList(&m_wndReportGroup);
			newReport->Create(&m_wndReportGroup);
			newReport->OnInitDialog();
			TranslateString(_T("General Insurance Outlay List"), szTemp);
			tmpStr.Format(_T("6. %s"), szTemp);
			obj->Add(tmpStr, _T(""), newReport);
		}
		TranslateString(_T("Outpatient Report Group"), szTemp);
		tmpStr.Format(_T("E. %s"), szTemp);
		obj->AddCaption(tmpStr);
		
		newReport = new CFMOutpatientInsuranceCost20A(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Outpatient Insurance Cost 20A"), szTemp);
		tmpStr.Format(_T("1. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMOutpatientInsuranceCost80A(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Outpatient Insurance Cost 80A"), szTemp);
		tmpStr.Format(_T("2. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMOutpatientInsuranceReport21a1(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Outpatient Insurance Cost 21A1"), szTemp);
		tmpStr.Format(_T("3. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		TranslateString(_T("Nh\xF3m \x62\xE1o \x63\xE1o \x63hi ph\xED \x42HYT"), szTemp);
		tmpStr.Format(_T("F. %s"), szTemp);
		obj->AddCaption(tmpStr);

		newReport = new CFMSudungthuocBHYT20a(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Drug by quarter of a year"), szTemp);
		tmpStr.Format(_T("1. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMSudungthuocBHYT20a_Y2015(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Drug by quarter of a year 2015"), szTemp);
		tmpStr.Format(_T("1.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMInsuranceReport21a(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General statistics and technical services for patients using medical insurance"), szTemp);
		tmpStr.Format(_T("2. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMInsuranceReport21a1(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General statistics and technical services for patients using medical insurance"), szTemp);
		tmpStr.Format(_T("2.1 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMInsuranceReport21a1_Y2015(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General statistics and technical services for patients using medical insurance 2015"), szTemp);
		tmpStr.Format(_T("2.2 %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		//newReport = new CFMDetailInsurance21aReport(&m_wndReportGroup);
		//newReport->Create(&m_wndReportGroup);
		//newReport->OnInitDialog();
		//TranslateString(_T("Detail Insurance Fee 21A Report"), szTemp);
		//tmpStr.Format(_T("2.2 %s"), szTemp);
		//obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMInsuranceCost79A(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Examination Cost 79A Report (79a - BHYT)"), szTemp);
		tmpStr.Format(_T("3. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMGeneralInsuranceCost79ATH(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("79A TH - BHYT"), szTemp);
		tmpStr.Format(_T("4. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMInsuranceCost80A(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Treatment Cost 80A Report (80a - BHYT)"), szTemp);
		tmpStr.Format(_T("5. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMGeneralInsuranceCost80ATH(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("80A TH - BHYT"), szTemp);
		tmpStr.Format(_T("6. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMWeeklyReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Weekly Hospital Fee Report"), szTemp);
		tmpStr.Format(_T("7. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMInsuranceFeeLossReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Insurance Fee Loss Report"), szTemp);
		tmpStr.Format(_T("8. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMOutpatientCostSynthesisbyDay(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Outpatient Cost Synthesis by Day"), szTemp);
		tmpStr.Format(_T("9. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMOutpatientCostSynthesisbyLine(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Outpatient Cost Synthesis by Line"), szTemp);
		tmpStr.Format(_T("10. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMOutpatientCostSynthesisbyDept(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Outpatient Cost Synthesis by Dept"), szTemp);
		tmpStr.Format(_T("11. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMInsuranceTreatmentCostByDept(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Inpatient Cost Synthesis By Dept"), szTemp);
		tmpStr.Format(_T("12. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMInsurancePatientTechDetail(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Insurance Patient Tech Detail"), szTemp);
		tmpStr.Format(_T("13. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
		
		TranslateString(_T("Expanded Report Group"), tmpStr);
		obj->AddCaption(tmpStr);

		//newReport = new CEMGeneralReportC1_1(&m_wndReportGroup);
		//newReport->Create(&m_wndReportGroup);
		//newReport->OnInitDialog();
		//TranslateString(_T("General Monthly Report by Department"), szTemp);
		//tmpStr.Format(_T("1. %s"), szTemp);
		//obj->Add(tmpStr, _T(""), newReport);

		newReport = new CEMWeekSynthesisReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Weekly Synthesis Report"), szTemp);
		tmpStr.Format(_T("1. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
		
		newReport = new CEMExamedPatientWithoutFee(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Examed Patient Without Fee"), szTemp);
		tmpStr.Format(_T("2. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

	}
}

void CHMSReportFormDialog::CreateLIMSReports(){
	CGuiView *newReport = NULL;
	CString tmpStr, szTemp;
	tmpStr.Empty();
	szTemp.Empty();
	CDocPane::CDocPaneInfo *obj = m_wndReportGroup.AddPane(_T("Report"));
	if (obj)
	{
		newReport = new CLIMSPatientList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Paraclinical Patient List"), szTemp);
		tmpStr.Format(_T("1. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);	

		newReport = new CLISTestActitivities(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Test Activities"), szTemp);
		tmpStr.Format(_T("2. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
	}
}

void CHMSReportFormDialog::CreatePACSReports()
{
	CGuiView *newReport = NULL;
	CString tmpStr, szTemp;
	tmpStr.Empty();
	szTemp.Empty();
	CDocPane::CDocPaneInfo *obj = m_wndReportGroup.AddPane(_T("Report"));

	if (obj)
	{
		TranslateString(_T("General Report Group"), szTemp);
		tmpStr.Format(_T("A. %s"), szTemp);
		obj->AddCaption(tmpStr);

		newReport = new CPACSEndoscopyListReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Endoscopy List Report"), szTemp);
		tmpStr.Format(_T("1. %s (Kho\x61 \x41\x33)"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPACSFunctionalProbeListReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Paraclinical Patient List Report"), szTemp);
		tmpStr.Format(_T("2. %s (Kho\x61 \x43\x37)"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPACSGeneralPatientByObjectReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("General Patient By Object Report"), szTemp);
		tmpStr.Format(_T("3. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPACSPatientListByResultReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Patient List By Result Report"), szTemp);
		tmpStr.Format(_T("4. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPACSImageActitivities(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Image Activities"), szTemp);
		tmpStr.Format(_T("5. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPACSImageFosteringList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Image Fostering List"), szTemp);
		tmpStr.Format(_T("6. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPACSXRayStatistic(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("XRay Statistics"), szTemp);
		tmpStr.Format(_T("7. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPACSPatientListUseContrast(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Patient List Use Contrast"), szTemp);
		tmpStr.Format(_T("8. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPACSInPatientUseChemistry(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("In Patient Use Chemistry"), szTemp);
		tmpStr.Format(_T("9. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPACSOutPatientUseChemistry(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Out Patient Use Chemistry"), szTemp);
		tmpStr.Format(_T("10. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPACSSettlementChemistryMaterial(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Settlement Chemistry Material"), szTemp);
		tmpStr.Format(_T("11. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CTMMaterialFinalAccount(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Material Final Account"), szTemp);
		tmpStr.Format(_T("12. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CPACSFunctionalProbeListSumeryReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Function Probe List Summary"), szTemp);
		tmpStr.Format(_T("13. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
	}
}
void CHMSReportFormDialog::CreateIPAReports(){
	CGuiView *newReport = NULL;
	CString tmpStr, szTemp;
	tmpStr.Empty();
	szTemp.Empty();
	CDocPane::CDocPaneInfo *obj = m_wndReportGroup.AddPane(_T("Report"));

	if (obj)
	{
		newReport = new CFMProcessedPrescriptionList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Processed Prescription List"), szTemp);
		tmpStr.Format(_T("1. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CFMProcessedPatientList(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Processed Patient List"), szTemp);
		tmpStr.Format(_T("2. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);
	}
	
}
void CHMSReportFormDialog::CreateBBReports(){
	CGuiView *newReport = NULL;
	CString tmpStr, szTemp;
	tmpStr.Empty();
	szTemp.Empty();
	CDocPane::CDocPaneInfo *obj = m_wndReportGroup.AddPane(_T("Report"));

	if (obj)
	{
		newReport = new CPMSStockCardReportDialog(&m_wndReportGroup, 1);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Stock Card"), szTemp);
		tmpStr.Format(_T("1. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CBloodBankListService(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Blood Bank List Service"), szTemp);
		tmpStr.Format(_T("2. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CReceiptListBloodBank(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("Receipt List BloodBank"), szTemp);
		tmpStr.Format(_T("3. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CBloodBankUseReportWeek(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("BloodBank Use Report Week"), szTemp);
		tmpStr.Format(_T("4. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CBBUnapprovedBloodBankListReportDialog(&m_wndReportGroup, 2);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("\x44\x61nh s\xE1\x63h m\xE1u y\xEAu \x63\x1EA7u \x63h\x1B0\x61 \x64uy\x1EC7t"), szTemp);
		tmpStr.Format(_T("5. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CCBloodBankReceiptReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("\x42\xE1o \x63\xE1o nh\x1EADp \x78u\x1EA5t t\x1ED3n \x63h\x1EBF ph\x1EA9m m\xE1u"), szTemp);
		tmpStr.Format(_T("6. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

		newReport = new CBloodBankDetailExportReport(&m_wndReportGroup);
		newReport->Create(&m_wndReportGroup);
		newReport->OnInitDialog();
		TranslateString(_T("\x42\xE1o \x63\xE1o \x63hi ti\x1EBFt \x78u\x1EA5t \x63ho \x63\xE1\x63 kho\x61 \x63h\x1EBF ph\x1EA9m m\xE1u"), szTemp);
		tmpStr.Format(_T("7. %s"), szTemp);
		obj->Add(tmpStr, _T(""), newReport);

	}

}