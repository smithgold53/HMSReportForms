#ifndef LIMSTESTSTATICSBYOBJECT_H
#define LIMSTESTSTATICSBYOBJECT_H

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
class CLIMSTestStaticsbyObject : public CGuiView{
protected:
public:
	CGuiGroupBox	m_wndReportParameters;
	CGuiLabel		m_wndDepartmentLabel;
	CGuiComboBox	m_wndDepartment;
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndMachineNoLabel;
	CGuiComboBox	m_wndMachineNo;
	CGuiButton		m_wndPrint;
	CGuiButton		m_wndExport;
	CString	m_szDepartmentKey;
	CString	m_szFromDate;
	CString	m_szToDate;
	CString	m_szMachineNoKey;
	void			OnDepartmentSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnDepartmentSelendok(); 
	//void			OnDepartmentSetfocus(); 
	//void			OnDepartmentKillfocus(); 
	long			OnDepartmentLoadData(); 
	void			OnDepartmentAddNew(); 
	//void			OnFromDateChange(); 
	//void			OnFromDateSetfocus(); 
	//void			OnFromDateKillfocus(); 
	int			OnFromDateCheckValue(); 
	//void			OnToDateChange(); 
	//void			OnToDateSetfocus(); 
	//void			OnToDateKillfocus(); 
	int			OnToDateCheckValue(); 
	void			OnMachineNoSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnMachineNoSelendok(); 
	//void			OnMachineNoSetfocus(); 
	//void			OnMachineNoKillfocus(); 
	long			OnMachineNoLoadData(); 
	void			OnMachineNoAddNew(); 
	void			OnPrintSelect(); 
	void			OnExportSelect(); 
	CLIMSTestStaticsbyObject(CWnd *pParent);
	~CLIMSTestStaticsbyObject();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddLIMSTestStaticsbyObject(); 
	int OnEditLIMSTestStaticsbyObject(); 
	int OnDeleteLIMSTestStaticsbyObject(); 
	int OnSaveLIMSTestStaticsbyObject(); 
	int OnCancelLIMSTestStaticsbyObject(); 
	int OnLIMSTestStaticsbyObjectListLoadData(); 
	CString GetQueryString();
};
#endif
