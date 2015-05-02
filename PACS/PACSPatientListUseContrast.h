#ifndef PACSPATIENTLISTUSECONTRAST_H
#define PACSPATIENTLISTUSECONTRAST_H

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
class CPACSPatientListUseContrast : public CGuiView{
protected:
	CString	m_szType;
public:
	CGuiGroupBox	m_wndReportParameters;
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndStorageLabel;
	CGuiComboBox	m_wndStorage;
	CGuiLabel		m_wndOrderNoLabel;
	CGuiComboBox	m_wndOrderNo;
	CGuiButton		m_wndExport;
	CGuiRadioButton	m_wndDrug;
	CGuiRadioButton	m_wndMaterial;
	CString	m_szFromDate;
	CString	m_szOrderNoKey;
	CString	m_szToDate;
	CString	m_szStorageKey;
	int m_nDrug;
	int m_nMaterial;
	//void			OnDateChange(); 
	//void			OnDateSetfocus(); 
	//void			OnDateKillfocus(); 
	int			OnFromDateCheckValue();
	long		OnStorageLoadData();
	void			OnOrderNoSelectChange(int nOldItemSel, int nNewItemSel); 
	void			OnOrderNoSelendok(); 
	//void			OnOrderNoSetfocus(); 
	//void			OnOrderNoKillfocus(); 
	long			OnOrderNoLoadData(); 
	//void			OnOrderNoAddNew(); 
	void			OnExportSelect(); 
	//void			OnToDateChange(); 
	//void			OnToDateSetfocus(); 
	//void			OnToDateKillfocus(); 
	int			OnToDateCheckValue(); 
	void	OnDrugSelect();
	void	OnMaterialSelect();
	CPACSPatientListUseContrast(CWnd *pParent);
	~CPACSPatientListUseContrast();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddPACSPatientListUseContrast(); 
	int OnEditPACSPatientListUseContrast(); 
	int OnDeletePACSPatientListUseContrast(); 
	int OnSavePACSPatientListUseContrast(); 
	int OnCancelPACSPatientListUseContrast(); 
	int OnPACSPatientListUseContrastListLoadData(); 
	CString GetQueryString();
};
#endif
