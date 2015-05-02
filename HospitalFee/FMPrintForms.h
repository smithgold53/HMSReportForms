#ifndef FMPRINTFORMS_H
#define FMPRINTFORMS_H
/*
	File nay dung de viet cac ham in dung trong module tiep don

*/


//Khoi tao cac biet toan cung ban dau
void FMInitialize(CWnd *pMain);
void FMPrint(CString szDept, long nDocumentNo, int nReceptIdx, CString szHema=_T("N"));
#endif
