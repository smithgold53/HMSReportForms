#include "stdafx.h"
#include "EMPrintForms.h"
#include "HMSMainFrame.h"
#include "ReportDocument.h"

void EMInitialize(CWnd *pMain){
	CHMSMainFrame *pMF = (CHMSMainFrame*) pMain;
	CRecord rs(&pMF->m_db);
	CString szSQL;
	
	pMF->m_CompanyInfo.sc_name.Empty();
	pMF->m_CompanyInfo.sc_pname.Empty();

	szSQL.Format(_T("SELECT sc_name, sc_pname FROM sys_company"));
	
	rs.ExecSQL(szSQL);

	if(!rs.IsEOF()){
		rs.GetValue(_T("sc_pname"), pMF->m_CompanyInfo.sc_pname);
		rs.GetValue(_T("sc_name"), pMF->m_CompanyInfo.sc_name);
	}
}

void EMPrintExaminationReceipt(){
	CHMSMainFrame *pMF = (CHMSMainFrame *)AfxGetMainWnd(); 
	static CReport rpt;
	CString szSQL, tmpStr;
	CRecord rs(&pMF->m_db);
	
}