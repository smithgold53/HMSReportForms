#ifndef PMSSUMMARYIMPORTEXPORTINSTOCKFORSTOCKSREPORTDIALOG_H
#define PMSSUMMARYIMPORTEXPORTINSTOCKFORSTOCKSREPORTDIALOG_H

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
class CPMSSummaryImportExportInstockForStocksReportDialog : public CGuiView{
protected:
	

	struct DRUGITEM{
		~DRUGITEM(){
			qtyinfo.RemoveAll();
		}
		CString	szType;
		CString	szID;
		double	szPrice;
		struct QuantityInfo{
			int		nStockId;
			int		nBeginQty;
			double	nBeginAmount;
			int		nImportQty;
			double	nImportAmount;			
			int		nExportQty;
			double	nExportAmount;
			int		nEndQty;
			double	nEndAmount;
		};
		CArray<QuantityInfo,QuantityInfo> qtyinfo;
	};

public:
	CGuiGroupBox	m_wndReportParameters;
	CGuiLabel		m_wndFromDateLabel;
	CGuiDateTimeCtrl	m_wndFromDate;
	CGuiLabel		m_wndToDateLabel;
	CGuiDateTimeCtrl	m_wndToDate;
	CGuiLabel		m_wndTypeLabel;
	CGuiComboBox	m_wndType;
	CGuiLabel		m_wndOriginalLabel;
	CGuiComboBox	m_wndOriginal;
	CGuiListCtrl	m_wndStockList;
	CGuiButton		m_wndPrint;
	CGuiButton		m_wndExportXLS;
	CGuiButton		m_wndClose;
	CGuiCheckBox	m_wndCheck;
	CString	m_szFromDate;
	CString	m_szToDate;
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
	long			OnStockListLoadData(); 
	void			OnStockListSelectChange(int nOldItem, int nNewItem); 
	void			OnStockListDblClick(); 
	int			OnStockListDeleteItem(); 
	void			OnPrintSelect(); 
	void			OnExportXLSSelect(); 
	void			OnCloseSelect(); 
	CPMSSummaryImportExportInstockForStocksReportDialog(CWnd *pParent);
	~CPMSSummaryImportExportInstockForStocksReportDialog();
	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();
	void OnDoDataExchange(CDataExchange* pDX);
	void GetDataToScreen();
	void GetScreenToData();
	void SetDefaultValues();
	int SetMode(int nMode);
	int OnAddPMSSummaryImportExportInstockForStocksReportDialog(); 
	int OnEditPMSSummaryImportExportInstockForStocksReportDialog(); 
	int OnDeletePMSSummaryImportExportInstockForStocksReportDialog(); 
	int OnSavePMSSummaryImportExportInstockForStocksReportDialog(); 
	int OnCancelPMSSummaryImportExportInstockForStocksReportDialog(); 
	int OnPMSSummaryImportExportInstockForStocksReportDialogListLoadData(); 
	DECLARE_MESSAGE_MAP()
	afx_msg void OnSetFocus(CWnd* pOldWnd);
};
#endif
