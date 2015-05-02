#ifndef PATIENTPROFILEDIALOG_H
#define PATIENTPROFILEDIALOG_H

#include "guidialog.h"
#include "ListPane.h"
#include "PatientProfileDrugLogs.h"

class AFX_EXT_CLASS CPatientProfileDialog :
	public CGuiDialog
{

public:
	long		m_nDocumentNo;

	CListPane m_wndPane;
	CPatientProfileDrugLogs		m_wndDrugLogs;
	CGuiLabel					m_wndDocNoLabel;
	CGuiNumberCtrl				m_wndDocNo;

	void	OnPaneSelectCallback(int nOldItem, int nNewItem); 
	int		OnDocNoCheckValue();
	void	OnDoDataExchange(CDataExchange *pDX);
	CPatientProfileDialog(CWnd *pParent, long nDocumentNo);
	~CPatientProfileDialog(void);

	void OnCreateComponents();
	void OnInitializeComponents();
	void OnSetWindowEvents();


};

#endif
