#ifndef FMINSURANCECOST79A_H
#define FMINSURANCECOST79A_H

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
class CFMInsuranceCost79A : public CGuiView{
protected:
	BOOL			m_bIsInsPurpose;//
	BOOL			m_bShowDiff;//hien thi phi chenh BH k thanh toan(ap dung cho the QN hoac bn cap tuong)
public:
	CGuiGroupBox	m_wndReportFilter;
	CGuiGroupBox	m_wndInsCardInfo;
	CGuiGroupBox	m_wndDeptInfo;
	CGuiGroupBox	m_wndExamRoomInfo;

	CGuiLabel		m_wndYearLabel;
	CGuiNumberCtrl	m_wndYear;
	CGuiLabel		m_wndReportPeriodLabel;
	CGuiComboBox	m_wndReportPeriod;
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndClerkLabel;
	CGuiComboBox	m_wndClerk;
	CGuiLabel		m_wndNumberOfDocLabel;
	CGuiNumberCtrl	m_wndNumberOfDoc;
	CGuiLabel		m_wndDeptLabel;
	CGuiComboBox	m_wndDept;

	CGuiCheckBox	m_wndWithoutEmergency;
	CGuiCheckBox	m_wndOnlyEmergency;
	CGuiCheckBox	m_wndUnapproved;
	CGuiCheckBox	m_wndOnlyCommander;
	CGuiCheckBox	m_wndOutpatientTreatment;
	CGuiCheckBox	m_wndByDischargedDate;
	CGuiCheckBox	m_wndSoldierPatient;
	CGuiButton		m_wndPrintPreview;
	CGuiButton		m_wndExport;
	CGuiButton		m_wndExportForIns;
	//CGuiListCtrl	m_wndList;
	CGuiListCtrl	m_wndCardList;
	CGuiListCtrl	m_wndDeptList;
	CGuiListCtrl	m_wndRoomList;

	long	m_nYear;
	int		m_nNumberofDoc;
	CString	m_szReportPeriodKey;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szClerkKey;
	CString	m_szDeptKey;

	BOOL	m_bWithoutEmergency;
	BOOL	m_bOnlyEmergency;
	BOOL	m_bUnapproved;
	BOOL	m_bOnlyCommander;
	BOOL	m_bOutpatientTreament;
	BOOL	m_bByDischargedDate;
	BOOL	m_bSoldierPatient;
	//void			OnYearChange(); 
	//void			OnYearSetfocus(); 
	//void			OnYearKillfocus(); 
	int			OnYearCheckValue(); 
	void			OnReportPeriodSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnReportPeriodSelendok(); 
	//void			OnReportPeriodSetfocus(); 
	//void			OnReportPeriodKillfocus(); 
	long			OnReportPeriodLoadData(); 
	void			OnReportPeriodAddNew(); 
	//void			OnFromDateChange(); 
	//void			OnFromDateSetfocus(); 
	//void			OnFromDateKillfocus(); 
	int			OnFromDateCheckValue(); 
	//void			OnToDateChange(); 
	//void			OnToDateSetfocus(); 
	//void			OnToDateKillfocus(); 
	int			OnToDateCheckValue(); 
	void			OnClerkSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnClerkSelendok(); 
	//void			OnClerkSetfocus(); 
	//void			OnClerkKillfocus(); 
	long			OnClerkLoadData(); 
	void			OnClerkAddNew();
	void			OnDeptSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnDeptSelendok(); 
	//void			OnDeptSetfocus(); 
	//void			OnDeptKillfocus(); 
	long			OnDeptLoadData(); 
	//void			OnDeptAddNew();
	void			OnWithoutEmergencySelect(); 
	void			OnOnlyEmergencySelect(); 
	void			OnUnapprovedSelect();
	void			OnPrintPreviewSelect(); 
	void			OnExportSelect();
	void			OnExportForInsSelect();
	/*long			OnListLoadData(); 
	void			OnListSelectChange(int nOldItem, int nNewItem); 
	void			OnListDblClick(); 
	int			OnListDeleteItem();*/

	long			OnCardListLoadData(); 
	void			OnCardListSelectChange(int nOldItem, int nNewItem); 
	void			OnCardListDblClick(); 
	int			OnCardListDeleteItem(); 
	long			OnDeptListLoadData(); 
	void			OnDeptListSelectChange(int nOldItem, int nNewItem); 
	void			OnDeptListDblClick(); 
	int			OnDeptListDeleteItem();
	long			OnRoomListLoadData(); 
	void			OnRoomListSelectChange(int nOldItem, int nNewItem); 
	void			OnRoomListDblClick(); 
	int			OnRoomListDeleteItem(); 

	CFMInsuranceCost79A(CWnd *pParent);
	~CFMInsuranceCost79A();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddFMInsuranceCost79A(); 
	int OnEditFMInsuranceCost79A(); 
	int OnDeleteFMInsuranceCost79A(); 
	int OnSaveFMInsuranceCost79A(); 
	int OnCancelFMInsuranceCost79A(); 
	int OnFMInsuranceCost79AListLoadData(); 
	CString GetQueryString();

	int OnCardListCheckAll();
	int OnCardListUnCheckAll();

	int OnRoomListCheckAll();
	int OnRoomListUnCheckAll();

	int OnDeptListCheckAll();
	int OnDeptListUnCheckAll();
};
#endif
