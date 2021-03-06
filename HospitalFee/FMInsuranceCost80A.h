#ifndef FMINSURANCECOST80A_H
#define FMINSURANCECOST80A_H

//////////////////////////////////////////////////////////////////////////////////////////////////////////
//	Copyright � Viet Nam Medical Software Joint Stock Company. 2005-2012. 			
//	All rights reserved. 
//	This program is protected by Viet nam and international treaties.  
//	Unauthorized reproduction or distribution of this program, 
//	or any portion of it, may result in severe civil and criminal penalties, 
//	and will be prosecuted to the maximum extent possible under the law.
//	This file is a part of the GUI(Graphical User Interface) class library.
//	(c) 2006-2008 Hoang Van Hay, All rights reserved.
//	CONTACT INFORMATION:
//	Email  : hayhv@vimes.com.vn or hayhv@yahoo.com
//	Website: http://www.vimes.com.vn
/////////////////////////////////////////////////////////////////////////////////////////////////////////
//	Ban quyen cua Cong Ty Co Phan Phan Mem Y Te Viet Nam 2005-2012.
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
class CFMInsuranceCost80A : public CGuiView{
protected:
	//xem 79A
	BOOL			m_bIsInsPurpose;
	BOOL			m_bShowDiff;
public:
	CGuiGroupBox	m_wndReportCondition;
	CGuiGroupBox	m_wndInsuranceCardInfo;
	CGuiGroupBox	m_wndDeptInfo;

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

	CGuiCheckBox	m_wndByDischargedDate;
	CGuiCheckBox	m_wndUnapproved;
	CGuiCheckBox	m_wndOnlyCommander;
	CGuiCheckBox	m_wndSoldierPatient;
	CGuiButton		m_wndPrint;
	CGuiButton		m_wndExport;
	CGuiButton		m_wndExportForIns;
	CGuiListCtrl	m_wndCardList;
	CGuiListCtrl	m_wndDeptList;

	long	m_nYear;
	CString	m_szReportPeriodKey;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szClerkKey;
	BOOL	m_bByDischargedDate;
	BOOL	m_bUnapproved;
	BOOL	m_bOnlyCommander;
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
	//void			OnReportPeriodAddNew(); 
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
	//void			OnClerkAddNew(); 
	void			OnByDischargedDateSelect();
	void			OnUnapprovedSelect();
	void			OnPrintSelect(); 
	void			OnExportSelect();
	void			OnExportForInsSelect();
	
	long			OnCardListLoadData(); 
	void			OnCardListSelectChange(int nOldItem, int nNewItem); 
	void			OnCardListDblClick(); 
	int			OnCardListDeleteItem();

	long			OnDeptListLoadData(); 
	void			OnDeptListSelectChange(int nOldItem, int nNewItem); 
	void			OnDeptListDblClick(); 
	int			OnDeptListDeleteItem();


	CFMInsuranceCost80A(CWnd* pParent);
	~CFMInsuranceCost80A();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddFMInsuranceCost80A(); 
	int OnEditFMInsuranceCost80A(); 
	int OnDeleteFMInsuranceCost80A(); 
	int OnSaveFMInsuranceCost80A(); 
	int OnCancelFMInsuranceCost80A(); 
	int OnFMInsuranceCost80AListLoadData();

	int OnCardListCheckAll();
	int OnCardListUnCheckAll();

	int OnDeptListCheckAll();
	int OnDeptListUnCheckAll();

	CString GetQueryString();
};
#endif
