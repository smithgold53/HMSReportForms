#ifndef EMDEPARTMENTUSAGEINDETAIL_H
#define EMDEPARTMENTUSAGEINDETAIL_H

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
class CEMDepartmentUsageinDetail : public CGuiView{
protected:
public:
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndOrderListLabel;
	CGuiComboBox	m_wndOrderList;
	CGuiGroupBox	m_wndDepartmentUsageinDetail;
	CGuiListCtrl	m_wndStorageList;
	CGuiListCtrl	m_wndDepartmentList;
	CGuiLabel		m_wndClerkLabel;
	CGuiComboBox	m_wndClerk;
	CGuiLabel		m_wndItemGroupLabel;
	CGuiComboBox	m_wndItemGroup;
	CGuiCheckBox	m_wndGroupbyDept;
	CGuiRadioButton	m_wndDrug;
	CGuiRadioButton	m_wndMaterial;
	CGuiRadioButton	m_wndAll;
	CGuiButton		m_wndPrintPreview;
	CGuiButton		m_wndExport;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString m_szOrderListKey;
	CString	m_szStorageKey;
	CString	m_szClerkKey;
	CString	m_szItemGroupKey;
	BOOL	m_bGroupbyDept;
	CString	m_szItemType;
	//void			OnFromDateChange(); 
	//void			OnFromDateSetfocus(); 
	//void			OnFromDateKillfocus(); 
	int			OnFromDateCheckValue(); 
	//void			OnToDateChange(); 
	//void			OnToDateSetfocus(); 
	//void			OnToDateKillfocus(); 
	int			OnToDateCheckValue(); 
	void			OnOrderListSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnOrderListSelendok(); 
	//void			OnOrderListSetfocus(); 
	//void			OnOrderListKillfocus(); 
	long			OnOrderListLoadData(); 
	//void			OnOrderListAddNew(); 
	void			OnStorageSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnStorageSelendok(); 
	//void			OnStorageSetfocus(); 
	//void			OnStorageKillfocus(); 
	long			OnStorageLoadData(); 
	//void			OnStorageAddNew(); 
	void			OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel);
	long			OnDepartmentLoadData();
	long			OnClerkLoadData();
	void			OnItemGroupSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnItemGroupSelendok(); 
	//void			OnItemGroupSetfocus(); 
	//void			OnItemGroupKillfocus(); 
	long			OnItemGroupLoadData(); 
	//void			OnItemGroupAddNew(); 
	void			OnDrugSelect();
	void			OnMaterialSelect();
	void			OnAllSelect();
	void			OnPrintPreviewSelect(); 
	void			OnExportSelect(); 
	CEMDepartmentUsageinDetail(CWnd *pParent);
	~CEMDepartmentUsageinDetail();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void SetDefaultValues();
	int SetMode(int nMode);
	CString GetQueryString();
	CString GetQueryString1();
	DECLARE_MESSAGE_MAP()
	afx_msg void OnSetFocus(CWnd* pOldWnd);
};
#endif
