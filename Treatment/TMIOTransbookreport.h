#ifndef CTMIOTransbookreport_H
#define CTMIOTransbookreport_H

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
class CTMIOTransbookreport: public CGuiView{
protected:
public:
	CGuiGroupBox	m_wndReportCondition;
	CGuiLabel		m_wndYearLabel;
	CGuiNumberCtrl	m_wndYear;
	CGuiLabel		m_wndReportPeriodLabel;
	CGuiComboBox	m_wndReportPeriod;
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndStatusLabel;
	CGuiComboBox	m_wndStatus;
	CGuiButton		m_wndPrintPreview;
	CGuiButton		m_wndPrint;
	CGuiButton		m_wndClose;
	CGuiCheckBox	m_wndInsurance;
	CGuiCheckBox	m_wndDiscount;
	CGuiCheckBox	m_wndService;

	long	m_nYear;
	CString	m_szReportPeriodKey;
	CString	m_szFromDate;
	CString	m_szToDate;

	CString	m_szStatusKey;
	BOOL	m_bInsurance;
	BOOL	m_bDiscount;
	BOOL	m_bService;

	bool	m_bPreview;
	//void			OnYearChange(); 
	//void			OnYearSetfocus(); 
	//void			OnYearKillfocus(); 
	int			OnYearCheckValue(); 
	void OnReportInpatientpatientlistreport();
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
	void			OnStatusSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnStatusSelendok(); 
	//void			OnDoctorSetfocus(); 
	//void			OnDoctorKillfocus(); 
	long			OnStatusLoadData(); 
	void			OnStatusAddNew(); 
	void			OnPrintPreviewSelect(); 
	void			OnPrintSelect(); 
	void			OnCloseSelect(); 
	void			OnInsuranceSelect(); 
	void			OnDiscountSelect(); 
	void			OnServiceSelect(); 
	void			OnExport();

	CTMIOTransbookreport(CWnd *pParent);
	~CTMIOTransbookreport();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddCTMIOTransbookreport(); 
	int OnEditCTMIOTransbookreport(); 
	int OnDeleteCTMIOTransbookreport(); 
	int OnSaveCTMIOTransbookreport(); 
	int OnCancelCTMIOTransbookreport(); 
	int OnCTMIOTransbookreportListLoadData(); 
	CString GetQueryString();
};
#endif
