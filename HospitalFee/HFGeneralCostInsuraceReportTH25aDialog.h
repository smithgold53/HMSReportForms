#ifndef HMSGeneralCostInsuraceReportTH25aDialog_H
#define HMSGeneralCostInsuraceReportTH25aDialog_H

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
class CHMSGeneralCostInsuraceReportTH25aDialog : public CGuiView{
protected:
public:
	CGuiGroupBox	m_wndReportFilter;
	CGuiLabel		m_wndYearLabel;
	CGuiNumberCtrl	m_wndYear;
	CGuiLabel		m_wndReportPeriodLabel;
	CGuiComboBox	m_wndReportPeriod;
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndClerkLabel;
	CGuiComboBox	m_wndClerk;
	CGuiButton		m_wndPrintPreview;
	CGuiButton		m_wndPrint;
	CGuiButton		m_wndExport;
	CGuiButton		m_wndClose;
	CGuiCheckBox	m_wndKhonglaybncc;
	CGuiCheckBox	m_wndLaybncc;
	CGuiCheckBox	m_wndSXtheohoso;

	long	m_nYear;
	int m_nButtonID;
	CString	m_szReportPeriodKey;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szClerkKey;
	BOOL	m_bKhonglaybncc;
	BOOL	m_bLaybncc;
	BOOL	m_bSXtheosohoso;

	CRecord *m_rs;
	CString m_szCompID;

	//void			OnYearChange(); 
	//void			OnYearSetfocus(); 
	//void			OnYearKillfocus(); 
	//Tong hop chi phi kham chua benh BHYT mau 25A
	void		PrintGeneralCost25AReport(CString szReportPeriod, CString szFromDate, CString szToDate,  bool bPreview=true);
	//Xuat ra Excell hop chi phi kham chua benh BHYT mau 25A
	void		ExportToExcell25AReport(CString szReportPeriod, CString szFromDate, CString szToDate,bool bPreview=true);
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
	void			OnClerkSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnClerkSelendok(); 
	//void			OnClerkSetfocus(); 
	//void			OnClerkKillfocus(); 
	long			OnClerkLoadData(); 
	void			OnClerkAddNew(); 
	void			OnPrintPreviewSelect(); 
	void			OnPrintSelect(); 
	void			OnExportSelect(); 
	void			OnCloseSelect(); 
	void			OnKhonglaybnccSelect(); 
	void			OnLaybnccSelect(); 
	void			OnSXtheosohosoSelect();
	CHMSGeneralCostInsuraceReportTH25aDialog(CWnd *pParent);
	~CHMSGeneralCostInsuraceReportTH25aDialog();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddHMSGeneralCostInsuraceReportTH25aDialog(); 
	int OnEditHMSGeneralCostInsuraceReportTH25aDialog(); 
	int OnDeleteHMSGeneralCostInsuraceReportTH25aDialog(); 
	int OnSaveHMSGeneralCostInsuraceReportTH25aDialog(); 
	int OnCancelHMSGeneralCostInsuraceReportTH25aDialog(); 
	int OnHMSGeneralCostInsuraceReportTH25aDialogListLoadData(); 
	CString GetQueryString();
};
#endif