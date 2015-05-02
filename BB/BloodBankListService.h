#ifndef BLOODBANKLISTSERVICE_H
#define BLOODBANKLISTSERVICE_H

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
class CBloodBankListService : public CGuiView{
protected:
public:
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiGroupBox	m_wndReportCondition;
	CGuiListCtrl	m_wndObject;
	CGuiListCtrl	m_wndDepartment;
	CGuiListCtrl	m_wndList;
	CGuiButton		m_wndPrintPreview;
	CGuiButton		m_wndExport;
	CString	m_szFromDate;
	CString	m_szToDate;
	//void			OnFromDateChange(); 
	//void			OnFromDateSetfocus(); 
	//void			OnFromDateKillfocus(); 
	int			OnFromDateCheckValue(); 
	//void			OnToDateChange(); 
	//void			OnToDateSetfocus(); 
	//void			OnToDateKillfocus(); 
	int			OnToDateCheckValue(); 

	int			OnObjectCheckAll();
	int			OnObjectUncheckAll();

	int			OnDeptCheckAll();
	int			OnDeptUncheckAll();

	int			OnListCheckAll();
	int			OnListUncheckAll();
	
	long			OnObjectLoadData(); 
	void			OnObjectSelectChange(int nOldItem, int nNewItem); 
	void			OnObjectDblClick(); 
	int			OnObjectDeleteItem(); 
	long			OnDepartmentLoadData(); 
	void			OnDepartmentSelectChange(int nOldItem, int nNewItem); 
	void			OnDepartmentDblClick(); 
	int			OnDepartmentDeleteItem(); 
	long			OnListLoadData(); 
	void			OnListSelectChange(int nOldItem, int nNewItem); 
	void			OnListDblClick(); 
	int			OnListDeleteItem(); 
	void			OnPrintPreviewSelect(); 
	void			OnExportSelect(); 
	CBloodBankListService(CWnd *pParent);
	~CBloodBankListService();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddBloodBankListService(); 
	int OnEditBloodBankListService(); 
	int OnDeleteBloodBankListService(); 
	int OnSaveBloodBankListService(); 
	int OnCancelBloodBankListService(); 
	int OnBloodBankListServiceListLoadData(); 
	CString GetQueryString();
};
#endif
