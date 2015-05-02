#ifndef rptDonthuoctonghop_H
#define rptDonthuoctonghop_H

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

class rptDonthuoctonghop : public CGuiView{
protected:
public:
	CString		m_szStockName;
	int			m_nType;
	
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;

	CGuiGroupBox	m_wndReportCondition;
	CGuiLabel		m_wndStockLabel;
	CGuiComboBox	m_wndStock;
	CGuiLabel		m_wndTypeLabel;
	CGuiComboBox	m_wndType;
	CGuiLabel		m_wndGroupLabel;
	CGuiComboBox	m_wndGroup;
	CGuiLabel		m_wndSourceLabel;
	CGuiComboBox	m_wndSource;
	CGuiButton		m_wndPrintPreview;
	CGuiButton		m_wndPrint;
	CGuiButton		m_wndExport;
	CGuiButton		m_wndClose;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szStockKey;
	CString	m_szTypeKey;
	CString	m_szGroupKey;
	CString	m_szSourceKey;
	//void			OnFromDateChange(); 
	//void			OnFromDateSetfocus(); 
	//void			OnFromDateKillfocus(); 
	int			OnFromDateCheckValue(); 
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
	void			OnTypeSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnTypeSelendok(); 
	//void			OnTypeSetfocus(); 
	//void			OnTypeKillfocus(); 
	long			OnTypeLoadData(); 
	void			OnTypeAddNew(); 
	void			OnGroupSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnGroupSelendok(); 
	//void			OnGroupSetfocus(); 
	//void			OnGroupKillfocus(); 
	long			OnGroupLoadData(); 
	void			OnGroupAddNew(); 
	void			OnSourceSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnSourceSelendok(); 
	//void			OnSourceSetfocus(); 
	//void			OnSourceKillfocus(); 
	long			OnSourceLoadData(); 
	void			OnSourceAddNew(); 
	void			OnPrintPreviewSelect(); 
	void			OnPrintSelect(); 
	void			OnExportSelect(); 
	void			OnCloseSelect(); 
	rptDonthuoctonghop(CWnd *pParent);
	~rptDonthuoctonghop();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddrptDonthuoctonghop(); 
	int OnEditrptDonthuoctonghop(); 
	int OnDeleterptDonthuoctonghop(); 
	int OnSaverptDonthuoctonghop(); 
	int OnCancelrptDonthuoctonghop(); 
	int OnrptDonthuoctonghopListLoadData(); 
	void	OnPrint(bool bPreview);
	void	OnExportToXLS(); //Ham Export To Excell
	DECLARE_MESSAGE_MAP()
	afx_msg void OnSetFocus(CWnd* pOldWnd);
};
#endif
