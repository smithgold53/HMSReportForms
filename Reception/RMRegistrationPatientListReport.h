#ifndef HRRegistrationPatientListReport_H
#define HRRegistrationPatientListReport_H

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
//Ngay 20/10/2010
//Sua abc
#include "GuiUtils.h"
#include "GuiView.h"
#include "DbField.h"

class AFX_EXT_CLASS CRMRegistrationPatientListReport : public CGuiView{
protected:
public:
	CGuiGroupBox	m_wndReceptionList;
	CGuiLabel		m_wndYearLabel;
	CGuiTextCtrl	m_wndYear;
	CGuiLabel		m_wndReportPeriodLabel;
	CGuiComboBox	m_wndReportPeriod;
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndExamRoomLabel;
	CGuiComboBox	m_wndExamRoom;
	CGuiLabel		m_wndObjectLabel;
	CGuiComboBox	m_wndObject;
	CGuiRadioButton	m_wndAll;
	CGuiRadioButton	m_wndExaminated;
	CGuiRadioButton	m_wndHasfee;
	CGuiButton		m_wndPrint;
	CGuiButton		m_wndExport;
	CGuiButton		m_wndClose;
	long	m_nYear;
	CString	m_szReportPeriodKey;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szObjectKey;
	CString m_szExamRoomKey;
	int	m_nAll;
	int	m_nExaminated;
	int	m_nHasfee;
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
	void			OnExamRoomSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnExamRoomSelendok(); 
	//void			OnExamRoomSetfocus(); 
	//void			OnExamRoomKillfocus(); 
	long			OnExamRoomLoadData(); 
	void			OnExamRoomAddNew(); 
	void			OnObjectSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnObjectSelendok(); 
	//void			OnObjectSetfocus(); 
	//void			OnObjectKillfocus(); 
	long			OnObjectLoadData(); 
	void			OnObjectAddNew(); 
	void			OnAllSelect(); 
	void			OnExaminatedSelect(); 
	void			OnHasfeeSelect(); 
	void			OnPrintSelect(); 
	void			OnExportSelect(); 
	void			OnCloseSelect(); 
	CRMRegistrationPatientListReport(CWnd *pParent);
	~CRMRegistrationPatientListReport();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddHRRegistrationPatientListReport(); 
	int OnEditHRRegistrationPatientListReport(); 
	int OnDeleteHRRegistrationPatientListReport(); 
	int OnSaveHRRegistrationPatientListReport(); 
	int OnCancelHRRegistrationPatientListReport(); 
	int OnHRRegistrationPatientListReportListLoadData(); 
	CString	GetQueryString();

	
};
#endif
