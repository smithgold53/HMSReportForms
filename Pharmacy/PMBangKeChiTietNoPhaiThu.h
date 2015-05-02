#ifndef RptPMS108_BangKeChiTietNoPhaiThu_H
#define RptPMS108_BangKeChiTietNoPhaiThu_H

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
class CRptPMS108_BangKeChiTietNoPhaiThu : public CGuiView{
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
	CGuiLabel		m_wndStockLabel;
	CGuiComboBox	m_wndStock;
	CGuiLabel		m_wndTypeLabel;
	CGuiComboBox	m_wndType;
	CGuiLabel		m_wndSupplierLabel;
	CGuiComboBox	m_wndSuppelier;
	CGuiLabel		m_wndClerkLabel;
	CGuiComboBox	m_wndClerk;
	CGuiButton		m_wndPrint;
	CGuiButton		m_wndExport;
	CGuiButton		m_wndClose;
	long	m_nYear;
	CString	m_szReportPeriodKey;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szStockKey;
	CString	m_szTypeKey;
	CString	m_szSuppelierKey;
	CString	m_szClerkKey;
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
	void			OnSuppelierSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnSuppelierSelendok(); 
	//void			OnSuppelierSetfocus(); 
	//void			OnSuppelierKillfocus(); 
	long			OnSuppelierLoadData(); 
	void			OnSuppelierAddNew(); 
	void			OnClerkSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnClerkSelendok(); 
	//void			OnClerkSetfocus(); 
	//void			OnClerkKillfocus(); 
	long			OnClerkLoadData(); 
	void			OnClerkAddNew(); 
	void			OnPrintSelect(); 
	void			OnExportSelect(); 
	void			OnCloseSelect(); 
	CRptPMS108_BangKeChiTietNoPhaiThu(CWnd *pParent);
	~CRptPMS108_BangKeChiTietNoPhaiThu();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddRptPMS108_BangKeChiTietNoPhaiThu(); 
	int OnEditRptPMS108_BangKeChiTietNoPhaiThu(); 
	int OnDeleteRptPMS108_BangKeChiTietNoPhaiThu(); 
	int OnSaveRptPMS108_BangKeChiTietNoPhaiThu(); 
	int OnCancelRptPMS108_BangKeChiTietNoPhaiThu(); 
	int OnRptPMS108_BangKeChiTietNoPhaiThuListLoadData(); 
	CString GetQueryString();
};
#endif
