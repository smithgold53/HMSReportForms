#ifndef RPTBAOCAOXUATTHUOCCHOKHOATHEODOITUONG_H
#define RPTBAOCAOXUATTHUOCCHOKHOATHEODOITUONG_H

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
class CrptBaocaoxuatthuocchokhoatheodoituong : public CGuiView{
protected:
public:
	CGuiGroupBox	m_wndReportCondition;
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndStockLabel;
	CGuiComboBox	m_wndStock;
	CGuiLabel		m_wndTypeLabel;
	CGuiComboBox	m_wndType;
	CGuiLabel		m_wndOriginalLabel;
	CGuiComboBox	m_wndOriginal;
	CGuiCheckBox	m_wndCheck;
	CGuiButton		m_wndPrintPreview;
	CGuiButton		m_wndPrint;
	CGuiButton		m_wndExport;
	CGuiButton		m_wndClose;
	CGuiListCtrl	m_wndListFac;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szStockKey;
	CString	m_szTypeKey;
	CString	m_szOriginalKey;
	int		m_bCheck;
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
	void			OnOriginalSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnOriginalSelendok(); 
	//void			OnOriginalSetfocus(); 
	//void			OnOriginalKillfocus(); 
	long			OnOriginalLoadData(); 
	void			OnOriginalAddNew(); 
	void			OnPrintPreviewSelect(); 
	void			OnPrintSelect(); 
	void			OnExportSelect(); 
	void			OnCloseSelect(); 
	long			OnListFacLoadData(); 
	void			OnListFacSelectChange(int nOldItem, int nNewItem); 
	void			OnListFacDblClick(); 
	int			OnListFacDeleteItem(); 
	int			OnCheckAll();
	int			OnUnCheckAll();
	CrptBaocaoxuatthuocchokhoatheodoituong(CWnd *pParent);
	~CrptBaocaoxuatthuocchokhoatheodoituong();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	void OnExportBaocaoxuathuocchokhoatheodoituong();
	void OnPrintBaocaoxuathuocchokhoatheodoituong();
	int SetMode(int nMode);
	int OnAddrptBaocaoxuatthuocchokhoatheodoituong(); 
	int OnEditrptBaocaoxuatthuocchokhoatheodoituong(); 
	int OnDeleterptBaocaoxuatthuocchokhoatheodoituong(); 
	int OnSaverptBaocaoxuatthuocchokhoatheodoituong(); 
	int OnCancelrptBaocaoxuatthuocchokhoatheodoituong(); 
	int OnrptBaocaoxuatthuocchokhoatheodoituongListLoadData(); 
	DECLARE_MESSAGE_MAP()
	afx_msg void OnSetFocus(CWnd* pOldWnd);
};
#endif