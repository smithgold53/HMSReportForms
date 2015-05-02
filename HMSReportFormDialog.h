#ifndef HMSREPORTFORMDIALOG_H
#define HMSREPORTFORMDIALOG_H
#include <atlimage.h>

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
#include "GuiDialog.h"
#include "GuiView.h"
#include "DbField.h"
#include "ReportImage.h"
#include "DocPane.h"

class AFX_EXT_CLASS CHMSReportFormDialog : public CGuiDialog{
protected:
	CString			m_szModuleID;
	CString			m_szHospitalID;
	int				m_nGroupSel;

public:
	CDocPane	m_wndReportGroup;
	void			OnReportGroupSelectChange(int nNew);	
	CHMSReportFormDialog();
	CHMSReportFormDialog(CWnd *pParent);
	~CHMSReportFormDialog();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddHMSReportFormDialog(); 
	int OnEditHMSReportFormDialog(); 
	int OnDeleteHMSReportFormDialog(); 
	int OnSaveHMSReportFormDialog(); 
	int OnCancelHMSReportFormDialog(); 
	int OnHMSReportFormDialogListLoadData(); 
	void InitReportGroup();
	void InitReportName();
	void CreateReceptionReports();
	void CreateExaminationReports();
	void CreateTreatmentReports();
	void CreateHospitalFeeReports();
	void CreatePMSReports();
	void CreateMAReports();
	void CreateLIMSReports();
	void CreatePACSReports();
	void CreateIPAReports();
	void CreateBBReports();
	//CRect			m_rcImage;	
	//CReportImage	m_wndReportImage;
	//CImage		m_image;
	CMap<int, int, CString, CString&> m_imageNames;
	int				nIndex;
};
#endif
