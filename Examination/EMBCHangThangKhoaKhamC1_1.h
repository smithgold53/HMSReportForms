#ifndef RPTBCHANGTHANGKHOAKHAMC1_1_H
#define RPTBCHANGTHANGKHOAKHAMC1_1_H

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
class CRptBCHangThangKhoaKhamC1_1 : public CGuiView{
protected:
public:
	CGuiLabel		m_wndYearLabel;
	CGuiNumberCtrl	m_wndYear;
	CGuiLabel		m_wndReportPeriodLabel;
	CGuiGroupBox	m_wndReportCondition;
	CGuiComboBox	m_wndReportPeriod;
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndExamRoomLabel;
	CGuiComboBox	m_wndExamRoom;
	CGuiCheckBox	m_wndPatient;
	CGuiCheckBox	m_wndTurn;
	CGuiCheckBox	m_wndDrugOrder;
	CGuiCheckBox	m_wndGoDept;
	CGuiButton		m_wndPrintPreview;
	CGuiButton		m_wndExport;
	CGuiButton		m_wndClose;
	long	m_nYear;
	CString	m_szReportPeriodKey;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szExamRoomKey;
	BOOL	m_bPatient;
	BOOL	m_bTurn;
	BOOL	m_bDrugOrder;
	BOOL	m_bGoDept;
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
	void			OnPatientSelect(); 
	void			OnTurnSelect(); 
	void			OnDrugOrderSelect(); 
	void			OnGoDeptSelect(); 
	void			OnPrintPreviewSelect(); 
	void			OnExportSelect(); 
	void			OnCloseSelect(); 
	CRptBCHangThangKhoaKhamC1_1(CWnd *pParent);
	~CRptBCHangThangKhoaKhamC1_1();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddRptBCHangThangKhoaKhamC1_1(); 
	int OnEditRptBCHangThangKhoaKhamC1_1(); 
	int OnDeleteRptBCHangThangKhoaKhamC1_1(); 
	int OnSaveRptBCHangThangKhoaKhamC1_1(); 
	int OnCancelRptBCHangThangKhoaKhamC1_1(); 
	int OnRptBCHangThangKhoaKhamC1_1ListLoadData(); 
	CString GetQueryString();
};
#endif
