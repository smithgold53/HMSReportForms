#ifndef PACSPATIENTLISTBYRESULTREPORT_H
#define PACSPATIENTLISTBYRESULTREPORT_H

//////////////////////////////////////////////////////////////////////////////////////////////////////////
//	Copyright © Viet Nam Medical Software Joint Stock Company. 2005-2012. 			
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

class CPACSPatientListByResultReport : public CGuiView{
protected:
public:

	CGuiGroupBox	m_wndReportCondition;
	CGuiGroupBox	m_wndRoomInfo;
	CGuiLabel		m_wndYearLabel;
	CGuiNumberCtrl	m_wndYear;
	CGuiLabel		m_wndReportPeriodLabel;
	CGuiComboBox	m_wndReportPeriod;
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndGroupLabel;
	CGuiComboBox	m_wndGroup;
	CGuiLabel		m_wndPractitionerLabel;
	CGuiComboBox	m_wndPractitioner;
	CGuiLabel		m_wndPerformDeptLabel;
	CGuiComboBox	m_wndPerformDept;
	CGuiLabel		m_wndResultLabel;
	CGuiTextCtrl	m_wndResult;
	CGuiButton		m_wndPrint;
	CGuiButton		m_wndExport;
	CGuiListCtrl	m_wndRoomList;

	long	m_nYear;
	CString	m_szReportPeriodKey;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szGroupKey;
	CString	m_szPractitionerKey;
	CString	m_szPerformDeptKey;
	CString	m_szResult;

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
	void			OnGroupSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnGroupSelendok(); 
	//void			OnGroupSetfocus(); 
	//void			OnGroupKillfocus(); 
	long			OnGroupLoadData(); 
	//void			OnGroupAddNew(); 
	void			OnPractitionerSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnPractitionerSelendok(); 
	//void			OnPractitionerSetfocus(); 
	//void			OnPractitionerKillfocus(); 
	long			OnPractitionerLoadData(); 
	//void			OnPractitionerAddNew(); 
	void			OnPerformDeptSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnPerformDeptSelendok(); 
	//void			OnPerformDeptSetfocus(); 
	//void			OnPerformDeptKillfocus(); 
	long			OnPerformDeptLoadData(); 
	//void			OnPerformDeptAddNew(); 
	//void			OnResultChange(); 
	//void			OnResultSetfocus(); 
	//void			OnResultKillfocus(); 
	int			OnResultCheckValue(); 
	void			OnPrintSelect(); 
	void			OnExportSelect(); 
	long			OnRoomListLoadData(); 
	void			OnRoomListSelectChange(int nOldItem, int nNewItem); 
	void			OnRoomListDblClick(); 
	int			OnRoomListDeleteItem(); 
	CPACSPatientListByResultReport(CWnd* pParent);
	~CPACSPatientListByResultReport();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddPACSPatientListByResultReport(); 
	int OnEditPACSPatientListByResultReport(); 
	int OnDeletePACSPatientListByResultReport(); 
	int OnSavePACSPatientListByResultReport(); 
	int OnCancelPACSPatientListByResultReport(); 
	int OnPACSPatientListByResultReportListLoadData();

	int OnListCheckAll();
	int OnListUnCheckAll();

	CString GetQueryString();
};
#endif
