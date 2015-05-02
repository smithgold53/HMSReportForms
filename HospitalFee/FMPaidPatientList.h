#ifndef FMPAIDPATIENTLIST_H
#define FMPAIDPATIENTLIST_H

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
class CFMPaidPatientList : public CGuiView{
protected:
public:
	CGuiGroupBox	m_wndPaidPatientList;
	CGuiGroupBox	m_wndStockInfo;
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndClerkLabel;
	CGuiComboBox	m_wndClerk;
	CGuiCheckBox	m_wndOutPatient;
	CGuiCheckBox	m_wndInPatient;
	CGuiCheckBox	m_wndLocked;
	CGuiButton		m_wndCloseOut;
	CGuiListCtrl	m_wndList;

	CGuiButton		m_wndPrintPreview;
	CGuiButton		m_wndPrintInvoice;
	CGuiButton		m_wndExport;
	CGuiCheckBox	m_wndWorkingDate;

	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szClerkKey;

	BOOL	m_bOutPatient;
	BOOL	m_bInPatient;
	BOOL	m_bLocked;
	BOOL	m_bWorkingDate;

	long double m_nTotal;

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
	void			OnLockedSelect(); 
	void			OnPrintPreviewSelect(); 
	void			OnPrintInvoiceSelect(); 
	void			OnExportSelect();
	void			OnWorkingDateSelect();
	void			OnOutPatientSelect(); 
	void			OnInPatientSelect();
	void			OnCloseOutSelect();
	long			OnListLoadData(); 
	void			OnListSelectChange(int nOldItem, int nNewItem); 
	void			OnListDblClick(); 
	int			OnListDeleteItem();

	CFMPaidPatientList(CWnd *pParent);
	~CFMPaidPatientList();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddFMPaidPatientList(); 
	int OnEditFMPaidPatientList(); 
	int OnDeleteFMPaidPatientList(); 
	int OnSaveFMPaidPatientList(); 
	int OnCancelFMPaidPatientList(); 
	int OnFMPaidPatientListListLoadData(); 
	CString GetQueryString();
	int OnPrintInvoice();

	int OnListCheckAll();
	int OnListUncheckAll();
};
#endif
