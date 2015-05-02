#ifndef PMREPORTCONDITIONFORM_H
#define PMREPORTCONDITIONFORM_H

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
class CPMReportConditionForm : public CGuiView{
protected:
	CString	m_szStockName;
public:
	CGuiGroupBox	m_wndReportCondition;
	//CGuiLabel		m_wndFromDateLabel;
	//CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndStorageLabel;
	CGuiComboBox	m_wndStorage;
	CGuiLabel		m_wndSupplierLabel;
	CGuiComboBox	m_wndSupplier;
	CGuiLabel		m_wndTypeLabel;
	CGuiComboBox	m_wndType;
	CGuiLabel		m_wndCategoryLabel;
	CGuiComboBox	m_wndCategory;
	CGuiLabel		m_wndSourceLabel;
	CGuiComboBox	m_wndSource;
	CGuiCheckBox	m_wndChestInventory;
	CGuiCheckBox	m_wndTypeSort;
	CGuiCheckBox	m_wndCategorySort;
	CGuiCheckBox	m_wndSupplierSort;
	CGuiButton		m_wndPrintPreview;
	CGuiButton		m_wndExport;
	//CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szSupplierKey;
	CString	m_szStorageKey;
	CString	m_szTypeKey;
	CString	m_szCategoryKey;
	CString	m_szSourceKey;
	BOOL	m_bChestInventory;
	BOOL	m_bTypeSort;
	BOOL	m_bCategorySort;
	BOOL	m_bSupplierSort;
	//void			OnFromDateChange(); 
	//void			OnFromDateSetfocus(); 
	//void			OnFromDateKillfocus(); 
	//int			OnFromDateCheckValue(); 
	//void			OnToDateChange(); 
	//void			OnToDateSetfocus(); 
	//void			OnToDateKillfocus(); 
	int			OnToDateCheckValue(); 
	void			OnStorageSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnStorageSelendok(); 
	//void			OnStorageSetfocus(); 
	//void			OnStorageKillfocus(); 
	long			OnStorageLoadData(); 
	//void			OnStorageAddNew();
	long			OnSupplierLoadData();
	void			OnTypeSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnTypeSelendok(); 
	//void			OnTypeSetfocus(); 
	//void			OnTypeKillfocus(); 
	long			OnTypeLoadData(); 
	//void			OnTypeAddNew(); 
	void			OnCategorySelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnCategorySelendok(); 
	//void			OnCategorySetfocus(); 
	//void			OnCategoryKillfocus(); 
	long			OnCategoryLoadData(); 
	//void			OnCategoryAddNew(); 
	void			OnSourceSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnSourceSelendok(); 
	//void			OnSourceSetfocus(); 
	//void			OnSourceKillfocus(); 
	long			OnSourceLoadData(); 
	//void			OnSourceAddNew(); 
	void			OnChestInventorySelect(); 
	void			OnTypeSortSelect(); 
	void			OnCategorySortSelect(); 
	void			OnSupplierSortSelect();
	void			OnPrintPreviewSelect(); 
	void			OnExportSelect(); 
	void			PrintFormat1();
	void			PrintFormat2();
	CPMReportConditionForm(CWnd *pParent);
	~CPMReportConditionForm();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void SetDefaultValues();
	int SetMode(int nMode);
	CString GetQueryString();
};
#endif
