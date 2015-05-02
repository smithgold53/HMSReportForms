#include "stdafx.h"
#include "RMPrintForms.h"
#include "GuiMainFrame.h"
#include "ReportDocument.h"

void RMInitialize(CWnd *pMain)
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
//  phieu kham
void RMPrintReceptionReceipt(CString szDept, long nDocumentNo, int nReceptIdx, CString szPrintHema){
	CGuiMainFrame *pMF = (CGuiMainFrame *)AfxGetMainWnd(); 

	static CReport rpt;
	CString szSQL, tmpStr;
	CRecord rs(&pMF->m_db);
	
	if(!rpt.Init(_T("Reports/HMS/HR_EXAMINATIONSHEET.RPT")) )
		return;

szSQL.Format(_T(" SELECT 	hd_patientno as patientno, ") \
		_T(" 	hd_docno as docno,") \
		_T(" 	trim(hp_surname||' '||hp_midname||' '||hp_firstname) as pname,") \
		_T(" 	hp_birthdate as birthdate,") \
		_T(" 	hms_getage(date(hd_admitdate), hp_birthdate) as age,") \
		_T(" 	hp_sex as sexid,") \
		_T(" 	sys_sel_getname('sys_sex', hp_sex) as sex,") \
		_T(" 	hp_occupation as occupationid,") \
		_T(" 	sys_sel_getname('sys_occupation', cast(hp_occupation as text)) as occupation,") \
		_T(" 	hp_dtladdr as detailaddress,") \
		_T(" 	hp_provid as provid,") \
		_T(" 	(select sp_name from sys_prov where sp_id = hp_provid) as provill,") \
		_T(" 	hp_distid as distid,") \
		_T(" 	(select sd_name from sys_dist where sd_provid = hp_provid and sd_id = hp_distid) as district,") \
		_T(" 	hp_villid as villid,") \
		_T(" 	(select sv_name  from sys_vill where sv_provid = hp_provid and sv_distid=hp_distid and sv_id = hp_villid) as village,") \
		_T(" 	hp_workplaceid as workplaceid,") \
		_T(" 	hp_workplace as workplace,") \
		_T(" 	hd_object as objectid,") \
		_T(" 	(SELECT ho_desc FROM hms_object WHERE ho_id=hd_object) as objectname,") \
		_T(" 	hd_cardno as cardno,") \
		_T(" 	hd_cardidx as cardidx,") \
		_T(" 	hc_regdate as regdate,") \
		_T(" 	hc_expdate as expdate,") \
		_T(" 	(SELECT hfl_name FROM hms_feelist WHERE hfl_feeid=he_examtype) as examtype,") \
		_T(" 	(SELECT hrl_name FROM hms_roomlist WHERE hrl_deptid=he_deptid AND hrl_id=he_roomid) as roomname, ") \
		_T("	he_receptno as receptno, ") \
		_T("	hfe_unitprice as amount, ") \
		_T("	hd_transplace as transplace, ") \
		_T("	hd_transdiagn as transdiagn, ")
		_T("	hd_xobject as xobject, ") \
		_T("	hd_xcardno as xcardno, ") \
		_T("	hd_xissuedate as xissuedate ") \
		_T(" FROM hms_patient") \
		_T(" LEFT JOIN hms_doc ON(hd_patientno=hp_patientno)") \
		_T(" LEFT JOIN hms_card ON(hc_patientno=hd_patientno and hc_cardno=hd_cardno and hc_idx=hd_cardidx) ") \
		_T(" LEFT JOIN hms_exam ON(he_docno=hd_docno)") \
		_T(" WHERE trim(he_deptid)='KB' AND he_docno= 10051663  AND he_receptidx= 1"));
_fmsg(_T("%s"), szSQL);
	int ret = rs.ExecSQL(szSQL);
	if(rs.IsEOF())
		return;
	
	//Report Header
	rpt.GetReportHeader()->SetValue(_T("HEALTHSERVICE"), HospitalName);
	rpt.GetReportHeader()->SetValue(_T("HOSPITALNAME"), HealthServiceName);
	tmpStr = pMF->GetSysDateTime();
	CString printDate;
	printDate.Format(rpt.GetReportHeader()->GetValue(_T("PrintDate")),tmpStr.Mid(11, 5), tmpStr.Mid(8, 2), tmpStr.Mid(5, 2), tmpStr.Left(4));
	rpt.GetReportHeader()->SetValue(_T("PrintDate"), printDate);
	rpt.GetReportHeader()->SetValue(_T("PatientNo"), rs.GetValue(_T("patientno")));
	rpt.GetReportHeader()->SetValue(_T("DocumentNo"), rs.GetValue(_T("docno")));
	StringUpper(rs.GetValue(_T("pname")), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("PatientName"), tmpStr);
	tmpStr.Empty();
	rpt.GetReportHeader()->SetValue(_T("Age"), rs.GetValue(_T("age")));
	rpt.GetReportHeader()->SetValue(_T("Sex"), rs.GetValue(_T("sex")));
	rpt.GetReportHeader()->SetValue(_T("Occupation"), rs.GetValue(_T("occupation")));
	tmpStr.Format(_T("%s - %s - %s"), rs.GetValue(_T("village")), rs.GetValue(_T("district")), rs.GetValue(_T("provill")));
	rpt.GetReportHeader()->SetValue(_T("Address"), tmpStr);
	rs.GetValue(_T("detailaddress"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("DetailAddress"), tmpStr);
	tmpStr.Empty();
	rpt.GetReportHeader()->SetValue(_T("WorkPlace"), rs.GetValue(_T("workplace")));
	rpt.GetReportHeader()->SetValue(_T("ExamRoom"), rs.GetValue(_T("roomname")));
	rpt.GetReportHeader()->SetValue(_T("SheetNo"), rs.GetValue(_T("receptno")));
	rpt.GetReportHeader()->SetValue(_T("ExamType"), rs.GetValue(_T("examtype")));
	
	rpt.GetReportHeader()->SetValue(_T("Money"), rs.GetValue(_T("amount")));
	rpt.GetReportHeader()->SetValue(_T("ObjectName"), rs.GetValue(_T("objectname")));
	rs.GetValue(_T("cardno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("CardNo"), tmpStr);
	rs.GetValue(_T("xobject"), tmpStr);	
	if(szPrintHema == _T("Y") && tmpStr == _T("H")){
		rpt.GetReportHeader()->SetValue(_T("Hema"), _T("HEMA"));
	}
	rpt.GetReportHeader()->SetValue(_T("RegDate"), CDate::Convert(rs.GetValue(_T("regdate"))));
	rpt.GetReportHeader()->SetValue(_T("ExpDate"), CDate::Convert(rs.GetValue(_T("expdate"))));
	rpt.GetReportHeader()->SetValue(_T("TransferHospital"), rs.GetValue(_T("transplace")));
	rpt.GetReportHeader()->SetValue(_T("TransferDiagnosis"), rs.GetValue(_T("transdiagn")));
	tmpStr.Empty();
	rpt.GetReportHeader()->SetValue(_T("Relason"), tmpStr);
	
	rs.GetValue(_T("xcardno"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("xcardno"), tmpStr);	
	rs.GetValue(_T("xissuedate"), tmpStr);
	rpt.GetReportHeader()->SetValue(_T("xissuedate"), CDate::Convert(tmpStr));

	
	
	//Page Header
	//Report Detail
	CReportSection* rptDetail = rpt.GetDetail(0); 
	//Page Footer
	//Report Footer
	rpt.Print();;

}