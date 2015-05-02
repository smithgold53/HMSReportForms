#include "stdafx.h"
#include "FMPrintForms.h"
#include "GuiMainFrame.h"
#include "ReportDocument.h"
void FMInitialize(CWnd *pMain)
{
	CGuiMainFrame *pMF = (CGuiMainFrame*) pMain;
	CRecord rs(&pMF->m_db);
	CString szSQL;
	
	HospitalName.Empty();
	HealthServiceName.Empty();
	szSQL.Format(_T("SELECT sc_name, sc_pname FROM sys_company"));
	rs.ExecSQL(szSQL);

	if(!rs.IsEOF())
	{
		rs.GetValue(_T("sc_pname"), HealthServiceName);
		rs.GetValue(_T("sc_name"), HospitalName);
	}
}
void FMPrint(CString szDept, long nDocumentNo, int nReceptIdx, CString szPrintHema){
	CGuiMainFrame *pMF = (CGuiMainFrame *)AfxGetMainWnd(); 
	static CReport rpt;
	CString szSQL, tmpStr;
	CRecord rs(&pMF->m_db);
}