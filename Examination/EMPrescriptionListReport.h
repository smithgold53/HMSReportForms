#ifndef EMPRESCRIPTIONLISTREPORT_H
#define EMPRESCRIPTIONLISTREPORT_H

//////////////////////////////////////////////////////////////////////////////////////////////////////////
//	Copyright(C) Viet Nam Medical Software Join Stock Company. 2005-2010. 			
//	All rights reserved. 
//	This program is protected by Viet nam and international treaties.  
//	Unauthorized reproduction or distribution of this program, 
//	or any portion of it, may result in severe civil and criminal penalties, 
//	and will be prosecuted to the maximum extent possible under the law.
//	This file is a part of the GUI(Graphical User Interface) class library.
//	(c) 2006-2008 Hay Hoang Van, All rights reserved.
//	CONTACT INFORMATION:
//	Email  : hayhv@vimes.com.vn or hayhv@yahoo.com
//	Website: http://www.vimes.com.vn
/////////////////////////////////////////////////////////////////////////////////////////////////////////
//	Ban quyen cua Cong Ty Co Phan Phan Mem Y Te Viet Nam 2005-2010.
//	Do Cuc Ban Quyen, Bo VHTT nuoc Cong hoa xa hoi chu nghia Viet Nam cap.
//	Chuong trinh phan mem nay duoc Luat phap Viet Nam va quoc te bao ho.
//	San xuat, su dung hoac phan phoi trai phep toan bo hoac mot phan cua phan men nay se
//	chiu cac hinh phat va hinh su hoac dan su, co the len den muc toi da dung theo Luat qui dinh.
//	File nay la mot phan cua thu vien lap trinh(GUI). Ban quyen cua Hoang Van Hay. 2006-2008
//	THONG TIN LIEN HE:
//	Email  : hayhv@vimes.com.vn hoac hayhv@yahoo.com
//	Website: http://www.vimes.com.vn
//////////////////////////////////////////////////////////////////////////////////////////////////////////
#include "GuiUtils.h"
#include "GuiView.h"
#include "DbField.h"
class CEMPrescriptionListReport : public CGuiView{
protected:
public:
	CGuiGroupBox	m_wndReportCondition;
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndStockLabel;
	CGuiComboBox	m_wndStock;
	CGuiLabel		m_wndDoctorLabel;
	CGuiComboBox	m_wndDoctor;
	CGuiLabel		m_wndApprovedByLabel;
	CGuiComboBox	m_wndApprovedBy;
	CGuiLabel		m_wndMinTotalAmtLabel;
	CGuiNumberCtrl	m_wndMinTotalAmt;
	CGuiCheckBox	m_wndServicePrice;
	CGuiRadioButton	m_wndCheckAll;
	CGuiRadioButton	m_wndOutPatient;
	CGuiRadioButton	m_wndInPatient;
	CGuiButton		m_wndPrintPreview;
	CGuiButton		m_wndExport;
	CGuiButton		m_wndClose;
	CGuiLabel		m_wndObjectLabel;
	CGuiComboBox	m_wndObject;
	CGuiListCtrl	m_wndList;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szStockKey;
	CString	m_szDoctorKey;
	CString	m_szApprovedByKey;
	long	m_nMinTotalAmt;
	CString	m_szCondition;
	CString	m_szObjectKey;
	BOOL	m_bServPrice;
	//void			OnFromDateChange(); 
	//void			OnFromDateSetfocus(); 
	//void			OnFromDateKillfocus(); 
	int			OnFromDateCheckValue(); 
	//void			OnToDateChange(); 
	//void			OnToDateSetfocus(); 
	//void			OnToDateKillfocus(); 
	int			OnToDateCheckValue(); 
	void			OnStockSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnStockSelendok(); 
	//void			OnStockSetfocus(); 
	//void			OnStockKillfocus(); 
	long			OnStockLoadData(); 
	void			OnStockAddNew(); 
	void			OnDoctorSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnDoctorSelendok(); 
	//void			OnDoctorSetfocus(); 
	//void			OnDoctorKillfocus(); 
	long			OnDoctorLoadData(); 
	void			OnDoctorAddNew();
	long			OnApprovedByLoadData();
	int				OnMinTotalAmtCheckValue();
	void			OnCheckAllSelect(); 
	void			OnOutPatientSelect(); 
	void			OnInPatientSelect(); 
	void			OnPrintPreviewSelect(); 
	void			OnExportSelect(); 
	void			OnCloseSelect(); 
	void			OnObjectSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnObjectSelendok(); 
	long			OnListLoadData(); 
	void			OnListSelectChange(int nOldItem, int nNewItem); 
	void			OnListDblClick(); 
	int			OnListDeleteItem(); 

	void			OnExport();
	void			OnPrintPreview();
	//void			OnObjectSetfocus(); 
	//void			OnObjectKillfocus(); 
	long			OnObjectLoadData(); 
	void			OnObjectAddNew(); 
	CEMPrescriptionListReport(CWnd *pParent);
	~CEMPrescriptionListReport();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddPMSPrescriptionListReportDialog(); 
	int OnEditPMSPrescriptionListReportDialog(); 
	int OnDeletePMSPrescriptionListReportDialog(); 
	int OnSavePMSPrescriptionListReportDialog(); 
	int OnCancelPMSPrescriptionListReportDialog(); 
	int OnPMSPrescriptionListReportDialogListLoadData(); 
	int OnListCheckC11Room();
	int OnListCheckC12Room();
	int OnListCheckC13Room();
	int OnListUncheckAll();
	DECLARE_MESSAGE_MAP()
	afx_msg void OnSetFocus(CWnd* pOldWnd);
};
#endif
