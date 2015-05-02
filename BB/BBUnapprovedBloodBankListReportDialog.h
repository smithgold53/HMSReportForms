#ifndef BBUnapprovedBloodBankListReportDialog_H
#define BBUnapprovedBloodBankListReportDialog_H

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
class CBBUnapprovedBloodBankListReportDialog : public CGuiView{
protected:
public:
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiGroupBox	m_wndReportCondition;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndStockLabel;
	CGuiComboBox	m_wndStock;
	CGuiLabel		m_wndNameLabel;
	CGuiTextCtrl	m_wndName;
	CGuiListCtrl	m_wndList;
	CGuiButton		m_wndPrintPreview;
	CGuiButton		m_wndPrint;
	CGuiButton		m_wndExport;
	CGuiButton		m_wndClose;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szStockKey;
	CString	m_szName;
	//void			OnFromDateChange(); 
	//void			OnFromDateSetfocus(); 
	//void			OnFromDateKillfocus(); 
	int			OnFromDateCheckValue(); 
	int			m_nPrintID;
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
	//void			OnNameChange(); 
	//void			OnNameSetfocus(); 
	//void			OnNameKillfocus(); 
	int			OnNameCheckValue(); 
	long			OnListLoadData(); 
	void			OnListSelectChange(int nOldItem, int nNewItem); 
	void			OnListDblClick(); 
	int			OnListDeleteItem(); 
	int			OnListSelectAll();
	int			OnListUnselectAll();
	void			OnPrintPreviewSelect(); 
	void			OnPrintSelect(); 
	void			OnExportSelect(); 
	void			OnCloseSelect(); 
	void			OnPrint(bool bPreview=false);
	void			OnPrint2(bool bPreview=false);
	void			OnOrderPrint(bool bPreview=false);
	void			OnOrderPrint2(bool bPreview=false);
	CBBUnapprovedBloodBankListReportDialog(CWnd *pParent, int m_bDetail);
	~CBBUnapprovedBloodBankListReportDialog();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddBBUnapprovedBloodBankListReportDialog(); 
	int OnEditBBUnapprovedBloodBankListReportDialog(); 
	int OnDeleteBBUnapprovedBloodBankListReportDialog(); 
	int OnSaveBBUnapprovedBloodBankListReportDialog(); 
	int OnCancelBBUnapprovedBloodBankListReportDialog(); 
	int OnBBUnapprovedBloodBankListReportDialogListLoadData(); 
	DECLARE_MESSAGE_MAP()
	afx_msg void OnSetFocus(CWnd* pOldWnd);
};
#endif
