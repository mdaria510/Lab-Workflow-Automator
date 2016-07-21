;AUTORUN--------------------------------------------------------
;===================================

#ClipboardTimeout 5000
#NoEnv
#LTrim
#InstallKeybdHook
SetDefaultMouseSpeed, 5
SetKeyDelay, 100
#SingleInstance

Gosub initvars
FirstRun :=

return

;AUTO REPLACE HOTSTRINGS--------------------------
; Significantly increases productivity by vastly improving report writing speed. Shorthand automatically expands to create the full phrase.
;===================================

::sp::specimen
::co::consist of
::cos::consists of
::cm::composed of
::cw::consistent with
::cs::cytology slide 
::css::cytology slides 
::cb::cell block
::cbnc::cell block is non-contributory due to scant cellularity
::cbx::core biopsy
::cscb::cytology slides and cell block consist of a 

::sh::sheet
::shs::sheets
::cl::cluster
::cls::clusters
::bg::background

::hycs::hypercellular specimen
::hocs::hypocellular specimen
::mocs::moderately cellular specimen

::hycp::hypercellular population
::hocp::hypocellular population
::mocp::moderately cellular population

::ben::benign 
::rct::reactive 
::aty::atypical
::sus::suspicious
::mal::malignant

::ucs::urothelial cells
::dcs::ductal cells
::bcs::bronchial cells
::fcs::follicular cells
::mcs::mesothelial cells
::scs::squamous cells
::gcs::glandular cells
::lcs::lymphoid cells

::aucs::aytpical urothelial cells
::adcs::aytpical ductal cells
::abcs::aytpical bronchial cells
::afcs::aytpical follicular cells
::amcs::aytpical mesothelial cells
::ascs::aytpical squamous cells
::agcs::aytpical glandular cells
::alcs::atypical lymphoid cells

::rcc::renal cell carcinoma 
::tcc::transitional cell carcinoma 
::hcc::hepatocellular carcinoma
::kscc::keratinizing squamous cell carcinoma
::scc::squamous cell carcinoma
::aca::adenocarcinoma
::smcc::small cell carcinoma
::nsmcc::non small cell carcinoma
::gist::gastrointestinal stromal tumor
::sar::sarcoma
::lya::lymphoma

::mi::mild 
::mod::moderate
::modly::moderately
::sv::severe 

::inf::inflammation
::ics::inflammatory cells
::ainf::acute inflammation
::cinf::chronic inflammation
::minf::mixed inflammation


::bl::blood
::rbc::red blood cell
::rbcs::red blood cells

::hcs::histiocytes
::fhcs::foam histiocytes
::hlms::hemosiderin-laden macrophages
::lyms::lymphocytes
::neus::neutrophils
::plas::plasma cell
::macs::macrophages

::db::debris
::col::colloid
::mng::multinodular goiter
::ro::rule out
::ho::history of
::pw::presents with
::s/p::status post
::ln::lymph node

::bac::bacteria

::degd::degenerated
::degn::degeneration
::pd::proteinaceous debris 

::homo::homogeneous
::heto::hetereogeneous

::ab::abundant
::sct::scattered

::afb+::AFB stain is positive for acid fast bacilli.
::afb-::AFB stain is negative for acid fast bacilli.

::gms+::GMS stain is positive for fungal organisms.
::gms-::GMS stain is negative for fungal organisms.



;HOTKEYS-------------------------------------------------------------------------
;===========================================

;Reset all variables per each case.

InitVars:
;--------------------------------------
Gosub UpdateButtons
Preg := 
Z331 := 
D07 := 
Z392 := 
Z01419 := 
Z124 := 
FarahaniChart := 
StJohnEpis := 
Vaginal := 
HPVOrdered := 
ComHosp := 
NoVerify := 
HPVPositive := 
FNA := 
GYN := 
NG := 
CN := 
CellBlock :=  
CoreBx := 
PreviousAbnReq := 
PreviousAbnLIS := 
Labnote := 
CTAdequacy := 
AdequacyCode := 
CTInitials := 
ClipFailed := 
FlowCaseNum := 
SpecType := 
ICDClip :=
ICDEntered := 
AdequacyEntered := 
AbnButtColor := 
Abn12Months := 
Consultation := 

CommonCase := 
CommonEntered := 


; Set images
iEditWindow := "\imagesnew\editwindow.png"
iCaseDemoWindow := "\imagesnew\casedemowindow.png"
iCaseDemoButton := "\imagesnew\casedemobutton.png"
iCaseChargesWindow := "\imagesnew\casechargeswindow.png"
iCaseChargesButton := "\imagesnew\casechargesbutton.png"
iVerifyCaseButton := "\imagesnew\verifycasebutton.png"
iFlowsheetButton := "\imagesnew\flowsheetbutton.png"
iVerifyChargesWindow := "\imagesnew\verifychargeswindow.png"
iAddChargesWindow := "\imagesnew\addchargeswindow.png"
iOutStandingreportsbutton := "\imagesnew\outStandingreportsbutton.png"
iSelectReportsWindow := "\imagesnew\selectreportsWindow.png"
iCLCytopathReviewButton := "\imagesnew\clcytopathreviewbutton.png"
iTransferReportButton := "\imagesnew\transferreportbutton.png"
iTransferReportWindow := "\imagesnew\transferreportwindow.png"
iPathologistHeader := "\imagesnew\pathologistheader.png"
iSlideCountWindow := "\imagesnew\slidecountwindow.png"
iClipboardError:= "\imagesnew\clipboarderror.png"


; Set diagnosis types
DiagTypes := ["NEG","NDIAG","BNGN","ATYP","SUSN","SUSM","POS","ATROV"]
For key, value in DiagTypes
	%value% := 


; Initialize arrays
Alphabet := "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,w,x,y,z"
SJEDocs := "schron,SJEH,beer,morrish,abbi,alapati,blake,dabri,dowvona,episcopal,khatri,liu,mahadkar,raghid,rahman,taher,ulep"
CTList := "dy,ms,xd,cc,rb1,am,sh,ds,no,th,sean,fs1,md,bv,jm,jk,vjd,tp1,ac"
MDList := "las,rede,far,azi,nav,gim,coc,cec,das,breu,shei"

return


; Establish location of buttons every time there is a potential change in interface.

UpdateButtons:
;--------------------------------------
;Find flowsheet button to establish reference point for left side


PixelSearch, FlowsheetButtonX, FlowsheetButtonY, % VerButLocX +459, % VerButLocY + 84, % VerButLocX + 1325, % VerButLocY + 131, 0x00FF00, , Fast

PrevAbnButtLocX := VerButLocX + 50
PrevAbnButtLocY := VerButLocY + 260
DemographicsCheckWhiteSpaceX := VerButLocX + 640
DemographicsCheckWhiteSpaceY := VerButLocY + 403
CommonCaseButtonX := VerButLocX + 510
CommonCaseButtonY := VerButLocY + 102
CaseChargesButtonX := VerButLocX + 30
CaseChargesButtonY := VerButLocY + 1
CaseDemographicsButtonX := VerButLocX + 255
CaseDemographicsButtonY := VerButLocY + 1

BoldX := EditWinLocX + 290
BoldY := EditWinLocY + 70

SpecDescFieldX := VerButLocX + 229
SpecDescFieldY := VerButLocY + 362

If GYN {
	ClinInfFieldX := SpecDescFieldX + 0
	ClinInfFieldY := SpecDescFieldY + 55
	AdequacyAlphaFieldX := SpecDescFieldX 
	AdequacyAlphaFieldY := SpecDescFieldY + 72
	ReasonAlphaFieldX := SpecDescFieldX + 0
	ReasonAlphaFieldY := SpecDescFieldY + 90
	EndoAlphaFieldX := SpecDescFieldX + 0
	EndoAlphaFieldY := SpecDescFieldY + 105
	StatementAdeqFieldX := SpecDescFieldX + 0 	
	StatementAdeqFieldY := SpecDescFieldY + 125
	DiagAlphaFieldX := SpecDescFieldX + 0 	
	DiagAlphaFieldY := SpecDescFieldY + 140	
	FinalDiagFieldX  := SpecDescFieldX + 0
	FinalDiagFieldY  := SpecDescFieldY + 155
	CommentFieldX := SpecDescFieldX + 0
	CommentFieldY := SpecDescFieldY + 173
	HPVFieldX := SpecDescFieldX + 0
	HPVFieldY := SpecDescFieldY + 190
}

If !GYN  {
	ClinInfFieldX := SpecDescFieldX + 0
	ClinInfFieldY := SpecDescFieldY + 40
	GrossDescFieldX := SpecDescFieldX + 0
	GrossDescFieldY := SpecDescFieldY + 50
	AdequacyAlphaFieldX := SpecDescFieldX 
	AdequacyAlphaFieldY := SpecDescFieldY + 72	
	ReasonAlphaFieldX := SpecDescFieldX + 0
	ReasonAlphaFieldY := SpecDescFieldY + 90
	StatementAdeqFieldX := SpecDescFieldX + 0 	
	StatementAdeqFieldY := SpecDescFieldY + 102	
	DiagAlphaFieldX := SpecDescFieldX + 0 	
	DiagAlphaFieldY := SpecDescFieldY + 120	
	FinalDiagFieldX  := SpecDescFieldX + 0
	FinalDiagFieldY  := SpecDescFieldY + 140
	CommentFieldX := SpecDescFieldX + 0
	CommentFieldY := SpecDescFieldY + 155
}

CaseChargesGreyspaceX := CCWinX + 85
CaseChargesGreyspaceY := CCWinY + 260
CaseChargesTasksX := CCWinX + 100
CaseChargesTasksY := CCWinY + 100

return


AStartup:
F4:: 
;Case entry startup script------------------------------

; Reset variables and ensure results window is active
Gosub InitVars
GoSub ResultWindowActive


; Ask for slide number
Inputbox, UserInput, Scan Slide, Scan Slide, , , , , , , ,

if errorlevel = 1
	exit

; Get individual pieces of case number
FullSlideNum = %UserInput%							
StringMid, SiteNum, FullSlideNum, 1, 2
StringMid, SpecClass, FullSlideNum, 3, 2
StringMid, SpecYear, FullSlideNum, 5, 2
StringMid, CaseNum, FullSlideNum, 7, 6

; Community hospitals need the test performed at LIJ comment
if SiteNum not in 10,80,99
	ComHosp := 1

; Determine proper specimen type	
if SpecClass in CT
	GYN := 1
if SpecClass in NG
	NG := 1
if SpecClass in CN
	CN := 1
if SpecClass in FN
	FNA := 1
if SpecClass in CC {
	Consultation := 1
	FNA := 1
}	
if SpecClass in CG {
	Consultation := 1
	GYN := 1
}		

; ensure FNAs cant be verified
if NG || CN || FNA
	NoVerify := 1

; Check if cytotech adequacy exists
GoSub CTAdequacyQuery

; Send full string to input box
SendInput !c%SiteNum%%SpecClass%%SpecYear%%CaseNum%{Enter}

; wait verify case button to light up to determine when specimen is loaded
WinWaitOpenByImageSearch(VerButLocX, VerButLocY, iVerifyCaseButton)
Gosub UpdateButtons ; Update buttons now that VerButLoc is known

;check for previous abnormals by seeing if button lights up --- NOT YET IMPLEMENTED
;GoSub PrevAbnCheck

; Make sure both demographics windows are open --- DEPRECATED
;PixelGetColor, color, DemographicsCheckWhiteSpaceX, DemographicsCheckWhiteSpaceY
;if color != 0xFFFFFF {
;	Msgbox, Case/Patient Demo Not open
;	exit
;}

;Open specimen description field and enter in default specimen description
Mouseclick, , SpecDescFieldX, SpecDescFieldY
SendInput {enter}
WinWaitOpenByImageSearch(EditWinLocX, EditWinLocY, iEditWindow)
SendInput ^k
Sleep 300
SendInput !o
WinWaitCloseByImageSearch(iEditWindow)

; Determine type of specimen
GoSub ClipLoop
SpecDisc = %clipboard%
Gosub SpecTypeAssignment

; Update button locations now that edit window has been opened
GoSub UpdateButtons

;Append "and cell block" to specimen description if applicable
GoSub CellBlockQuery 

;GoSub CoreBXQuery ---- NOT IMPLEMENTED YET

; Cycle through gross descriptions for non-gyns
If !GYN {
	GoSub ClipLoop
	
	Loop, 2 {
		SendInput {down}{enter}
		WinWaitOpenByImageSearch(EditWinLocX, EditWinLocY, iEditWindow)
		WinWaitCloseByImageSearch(iEditWindow)
	}
}

; Get ICD number from fields and case charges
If GYN {
	
	GoSub CaseCharges
	PixelGetColor, color, % CCWinX + 151, % CCWinY + 383 ; Check for the grey background so blank field doesnt error
	
	if color = 0xE0E0E0 {
		Traytip, No Data, No Data
		Send {enter}
		WinWaitCloseByImageSearch(iCaseChargesWindow)	
	} Else {
		Mouseclick, ,% CCWinX + 151, % CCWinY + 387
		GoSub Cliploop
		ICDclip = %clipboard%
		send {Enter}
		WinWaitCloseByImageSearch(iCaseChargesWindow)		
	}

	If sitenum != 99 {	

		SendInput {down}
		Gosub cliploop
		If clipboard contains HPV and pos
			NoVerify := 1
		If clipboard contains 33.1
			Z331 := 1
		If clipboard contains 39.2
			Z392 := 1
		If clipboard contains Z01.419
			Z01419 := 1
		If clipboard contains Z12.4
			Z124 := 1
		If clipboard contains D07
			VD07 := 1
		if clipboard contains N87.9,C53,R87.610,R87.611,R87.612,R87.613,D07,C51.9,C56.9,C57.00,LSIL,LGSIL,HGSIL,HSIL,ASCUS,ASC-H,ASC-US {
			PreviousAbnReq := 1
			NoVerify := 1	
		}
	}
	else 	{
		SendInput {down}
	}
	

	SendInput {down}
	Gosub cliploop
	If clipboard contains HPV and pos
		NoVerify := 1
	If clipboard contains 33.1
		Z331 := 1
	If clipboard contains 39.2
		Z392 := 1
	If clipboard contains Z01.419
		Z01419 := 1
	If clipboard contains Z12.4
		Z124 := 1
	If clipboard contains D07
		VD07 := 1
	if clipboard contains N87.9,C53,R87.610,R87.611,R87.612,R87.613,D07,C51.9,C56.9,C57.00,LSIL,LGSIL,HGSIL,HSIL,ASCUS,ASC-H,ASC-US {
		PreviousAbnReq := 1
		NoVerify := 1	
	}

}



;Pull doctor name from demographics and specimen desc, and add in notes if applicable
GoSub DocCheck

;Check for common cases
if NG || CN || FNA {	
	PixelGetColor, color, CommonCaseButtonX, CommonCaseButtonY
	if color = 0x00FFFF
		CommonCase := 1
}

;Enter in the comhosp note if needed
If ComHosp { 
	MouseClick, , CommentFieldX, CommentFieldY 
	SendInput {enter}
	WinWaitOpenByImageSearch(EditWinLocX, EditWinLocY, iEditWindow)
	SendInput CL
	sleep 500
	sendinput {f9}
	sleep 500
	Sendinput !o
	WinWaitCloseByImageSearch(iEditWindow)			
}

;Sendinput {down}


return


!^y::
;Safe Verify Case --------------------------------

GoSub LabNoteEnter

If NoVerify
	MsgBox, Cannot Verify
else
	sendinput ^y

return



!^f::
;Safe Perform Case --------------------------------

GoSub LabNoteEnter

if CommonCase && !CommonEntered
	MsgBox, Common case entered?

If CTAdequacy && !AdequacyEntered {
	MsgBox, Adequacy not entered
	return
}

If !ICDEntered
	MsgBox, ICD not entered
else
	SendInput ^f

return



;SUBROUTINES-------------------------------------------------------------------
;============================================


CaseCharges:
;--------------------------------------
MouseClick, , CaseChargesButtonX, CaseChargesButtonY
WinWaitOpenByImageSearch(CCWinX, CCWinY, iCaseChargesWindow)
Gosub UpdateButtons
return


VerifyDas:
;--------------------------------------
MouseClick, , CaseChargesTasksX, CaseChargesTasksY
;Sendinput {down}
Sendinput !v
;WinWaitOpenByAreaColor(VerWinLocX, VerWinLocY, 502, 364, 90, 0x00FF00)
WinWaitOpenByImageSearch(CCWinX, CCWinY, iVerifyChargesWindow)
Sendinput da
return


SpecTypeAssignment:
;--------------------------------------
If GYN {
	If SpecDisc contains Vag {
		If SpecDisc not contains Cer
			Vaginal := 1
	}
	
	If SpecDisc contains HPV 
		HPVordered := 1
	
	If SpecDisc contains ref
		HPVordered :=
}

If !GYN {
	If SpecDisc contains urine,bladder,ureteral
		SpecType := "Urine"
	
	If SpecDisc contains csf,cerebrospinal
		SpecType := "CSF"
	
	If SpecDisc contains bronch,bronchial,bal,lavage
		SpecType := "Bronch"
	
	If SpecDisc contains sputum
		SpecType := "Sputum"
	
	If SpecDisc contains pelvic
		SpecType := "Pelvic"
	
	If SpecDisc contains peritoneal
		SpecType := "Peritoneal"		
	
	If SpecDisc contains ascitic,ascities
		SpecType := "Ascitic"
	
	If SpecDisc contains pleural
		SpecType := "Pleural"
	
	If SpecDisc contains pericardial
		SpecType := "Pericardial"
	
	If SpecDisc contains ovarian cyst
		SpecType :=  "OvCyst"
	
	If SpecDisc contains lymph
		SpecType := "LymphFNA"
	
	If SpecDisc contains thyroid
		SpecType := "ThyroidFNA"
	
	If SpecDisc contains pancreas, pancreatic
		SpecType := "PancreasFNA"
	
	If SpecDisc contains soft
		SpecType := "SoftTissueFNA"
	
	If SpecDisc contains salivary,parotid,submandibular
		SpecType := "SalivaryFNA"
	
	If SpecDisc contains breast {
		
		Inputbox, UserInput, Type, 1 = Cyst 2 = Nipple Discharge 3 = FNA, , , , , , , , 1
		
		If UserInput = 1
			SpecType := "BreastCyst"

		If UserInput = 2
			SpecType := "Nipple"

		If UserInput = 3
			SpecType := "BreastFNA"

	}
	If SpecDisc contains Nipple
		SpecType := "NippleDischarge"
		
	If SpecDisc contains Bone
		SpecType := "BoneFNA"
}
return 



ICDEnter:
;--------------------------------------
if ICDEntered
	return

if SpecType = Urine {
	If NEG or ATYP or NDIAG
		Inputbox, UserInput, ICD Code,
		(
		N32.9 = Unspec Disorder Bladder
		R31.9 = Microscopic Hematuria
		N30.0 = Acute Cystitis
		B37.49 = Candida
		), , , , , , , ,N32.9
	else
		Inputbox, UserInput, ICD Code,
		(
		D49.4 = Bladder Neoplasm
		C67.9 = Bladder Malignancy
		), , , , , , , ,C67.9
	
	ICDInput(UserInput)
}
if SpecType = Pericardial {
	Inputbox, UserInput, ICD Code,I30.8 = Pericardial Effusion`nC79.89 = Malig Pericardial Effusion, , , , , , , ,I30.8
	ICDInput(UserInput)
}
if SpecType = Pleural {
	Inputbox, UserInput, ICD Code,J91.8 = Pleural Effusion`nC78.2 = Malig Pleural Effusion`nR09.1 = Pleuritis, , , , , , , ,J91.8
	ICDInput(UserInput)
}
if (SpecType = "Pelvic" or Spectype = "Peritoneal") {
	Inputbox, UserInput, ICD Code,R19.09 = Pelvic Mass`n, , , , , , , ,R19.09
	ICDInput(UserInput)
}

if SpecType = BoneFNA {
	Inputbox, UserInput, ICD Code,M89.9 = Bone Disorder NOS`n = Bone Malig Neo`nC79.51 = Bone Metastasis, , , , , , , ,M89.9
	ICDInput(UserInput)
}

if SpecType = Ascitic {
	Inputbox, UserInput, ICD Code,R18.8 = Ascites`nR18.0 = Malig Ascites, , , , , , , ,R18.8
	ICDInput(UserInput)
}

if SpecType = ThyroidFNA {
	Inputbox, UserInput, ICD Code,E04.9 = Goiter%A_Tab%E04.1 = Cyst `nE06.3 = Hashimoto%A_Tab%E04.2 = Multinod Goiter `nD44.0 = NOS Neo%A_Tab%C73 = Malig Neo, , , , , , , ,E04.9
	ICDInput(UserInput)
}
if SpecType = LymphFNA {
	Inputbox, UserInput, ICD Code,R59.9 = LN Enlargement`nC77.9 = LN Metastasis`nC85.800 = Non=Hodgkin's Lymphoma`n201.90 = Hodgkin's Lymphoma, , , , , , , ,R59.9
	ICDInput(UserInput)
}
if SpecType = SoftTissueFNA {
	Inputbox, UserInput, ICD Code,M79.9 = Soft Tissue Disorder NOS, , , , , , , ,M79.9
	ICDInput(UserInput)
}
if SpecType = BreastCyst {
	Inputbox, UserInput, ICD Code,N64.9 = Breast Disorder NOS`nN60.09 = Breast Cyst`nD24.1 = Fibroadenoma`nn63 = Breast Mass`nC50.929 = Breast Malig`n198.91 = Metastasis to Breast, , , , , , , ,N64.9
	ICDInput(UserInput)
}
if SpecType = Nipple {
	Inputbox, UserInput, ICD Code,N64.9 = Breast Disorder NOS`nD24.1 = Fibroadenoma`nn63 = Breast Mass`nC50.929 = Breast Malig`n198.91 = Metastasis to Breast, , , , , , , ,N64.9
	ICDInput(UserInput)
}
if SpecType = BreastFNA {
	Inputbox, UserInput, ICD Code,N64.9 = Breast Disorder NOS`nD24.1 = Fibroadenoma`nn63 = Breast Mass`nC50.929 = Breast Malig`n198.91 = Metastasis to Breast, , , , , , , ,N64.9
	ICDInput(UserInput)
}

if SpecType = SalivaryFNA {
	Inputbox, UserInput, ICD Code, R22.1 = Neck mass`nK11.9 = Neck disease, nos`nD11.9 = Benign neoplasm`nC08.9 = Malig Neoplasm, , , , , , , R22.1
	ICDInput(UserInput)
}

if SpecType = CSF {
	ICDInput("G93.9")
}

if SpecType = Bronch  {
	Inputbox, UserInput, ICD Code,J98.9 = Respiratory Condition NOS`n786.6 = Lung Mass`n162.9 = Lung Malignancy`n162.9 = Metastasis to Lung`n136.6=PCP pneumonia`n515 = Granuloma`n, , , , , , , ,J98.9
	ICDInput(UserInput)
}

if SpecType = Sputum {
	Inputbox, UserInput, ICD Code,J98.9 = Respiratory Condition NOS`n786.6 = Lung Mass`n162.9 = Lung Malignancy`n162.9 = Metastasis to Lung`n136.6=PCP pneumonia`n515 = Granuloma`n, , , , , , , ,J98.9
	ICDInput(UserInput)
}

if SpecType = NippleDischarge
	ICDInput("N64.52")

if SpecType = PancreasFNA {
	if POS
		ICDInput("C25.9")
	else 
		ICDInput("K86.9")	
}
	
if Spectype = OvCyst
	ICDInput("N83.29")	


if GYN && !Vaginal {
	if NEG 	{
		if PreviousAbnReq || PreviousAbnLIS
		{
			ICDInput("Z01.42")
			if PreviousAbnLIS && Abn12Months
				labnote := 1 
		}
	}
	
	if ATROV
		if ICDClip contains N95.2
			ICDEntered := 1
	else 	ICDInput("N95.2")
		
	if Z331
		if ICDClip contains Z33.1
			ICDEntered := 1
	else	ICDInput("Z33.1")
		
	if Z392
		if ICDClip contains Z39.2
			ICDEntered := 1
	else	ICDInput("Z39.2")
		
	if Z01419 
		if ICDClip contains Z01.419
			ICDEntered := 1	
	else	ICDInput("Z01.419")
		
	if Z124
		if ICDClip contains Z12.4
			ICDEntered := 1
	else 	ICDInput("Z12.4")
}

if GYN && Vaginal {
	if PreviousAbnReq || PreviousAbnLIS
		ICDInput("VZ08")
	else
		ICDInput("Z12.72")
}

if !ICDEntered && GYN
	ICDInput("Z12.4")

If !GYN {	
	PixelGetColor, color, % CCWinX + 151, % CCWinY + 383 ; Check for the grey background so blank field doesnt error
	
	if color = 0xC8D0D4 {
		Inputbox, UserInput, ICD Code,, , , , , , , ,
		ICDInput(UserInput)
		
	}
	WinWaitCloseByImageSearch(iCaseChargesWindow)	
}

return


ICDInput(ICDCode)
;--------------------------------------
{
	global ICDEntered
	global Z01419
	global NEG
	if !ICDEntered	{
		if Z01419 = 1 {	
			if !NEG
				sleep 250
		}
		else	
			Sendinput !g{tab 2}!r
		
		sleep 100
		Sendinput !g
		sleep 500
		Sendinput %ICDCode%
		sleep 100
		Send {enter 2}
		Sendinput {enter}{tab 4}
		ICDEntered = 1
	}	
}
return


CPTInput(CPTCode)
;--------------------------------------
{
	global 	
		
	{
		Sendinput !d
		WinWaitOpenByImageSearch(CPTWinLocX, CPTWinLocY, iAddChargesWindow)
		Sendinput CPT %CPTCode%
		Sendinput {enter 2}
		WinWaitCloseByImageSearch(iAddChargesWindow)
	
	}	
}
return


DocCheck:
;--------------------------------------
Mouseclick, ,CaseDemographicsButtonX, CaseDemographicsButtonY
WinWaitOpenByAreaColor(CDWinLocX, CDWinLocY, 142, 136, 30, 0x97F9FF)
Mouseclick, ,% CDWinLocX + 149, % CDWinLocY + 211
Mouseclick, ,% CDWinLocX + 149, % CDWinLocY + 251
SendInput {right}
sleep 100
GoSub ClipLoop
SendInput !c
WinWaitCloseByPixelColor(CDWinLocX, CDWinLocY, 0xFFFF00)

DocName := clipboard

If DocName contains farahani,farhani {
	MouseClick, , CommentFieldX, CommentFieldY 
	SendInput {enter}
	
	Loop {
		Inputbox, UserInput, Chart#, Enter Chart#, , , , , , , ,
		if errorlevel
			exit
		else break
	}
	
	SendInput Chart#%UserInput%!o
	WinWaitCloseByImageSearch(iEditWindow)
}


If DocName contains %SJEDocs% {
	MouseClick, , CommentFieldX, CommentFieldY
	SendInput {enter}
	
	Loop {
		Inputbox, UserInput, M#, Enter M#, , , , , , , ,
		if errorlevel
			exit
		else if UserInput contains %Alphabet%
			continue
		else break
		}
	SendInput M{#}%UserInput%{enter}
		
	Loop {
		Inputbox, UserInput, Enter V#, Enter V#, , , , , , , ,
		if errorlevel
			exit
		else if UserInput contains %Alphabet%
			continue
		else break
		}
	SendInput V{#}%Userinput%!o
	WinWaitCloseByImageSearch(iEditWindow)
}


If DocName contains MMC,HSUEH,ONER,ZHU,SILVA,MIKELL,KEPPEL,HAKIMAN,KALEYA,TUNG,SHAW,CHEN,ALEBDY,SLATKIN,CURRY,RATNER,COLANNE {
	MouseClick, , GrossDescFieldX, GrossDescFieldY 
	SendInput {enter}
	
	Loop {
		Inputbox, UserInput, Ref#, Enter Ref# G16-, , , , , , , ,
		if errorlevel
			exit
		else break
		}
	
	SendInput {ENTER}Specimen received from Maimonides Medical Center with corresponding accession G16-%UserInput%!o


	WinWaitCloseByImageSearch(iEditWindow)
}

return

LabnoteEnter:
;--------------------------------------
If Labnote {
	MouseClick, , ClinInfFieldX, ClinInfFieldY
	WinWaitOpenByImageSearch(EditWinLocX, EditWinLocY, iEditWindow)
	Sendinput {pgdn 3}
	Sendinput (Lab Note: As per lab records, patient has history of previous abnormal within the prior 12 months.)
	Sendinput !o
	WinWaitCloseByImageSearch(iEditWindow)
	Labnote :=
}
return


ReformatReport:
;--------------------------------------
Mouseclick, , SpecDescFieldX, SpecDescFieldY
gosub cliploop ; pull the specimen description

MouseClick, , FinalDiagFieldX, FinalDiagFieldY
Sendinput {enter 2}
WinWaitOpenByImageSearch(EditWinLocX, EditWinLocY, iEditWindow)
SendInput {up}
SendInput ^b^v{down}{enter}
Sleep 100
SendInput ^{end}{home}

If Spectype = LymphFNA {
	if NEG || NDIAG || BNGN {
		SendInput Cytology slides and cell block consist of a heterogeneous population of lymphoid cells. 
		sleep 100
		SendInput Favor a reactive process. No Reed-Sternberg cells, granuloma or epithelial cells present.   
		sleep 100
		SendInput Recommend tissue examination if lymphadenopathy persist and if clinically indicated.
	}
}	
Else If Spectype = ThyroidFNA {
	if BNGN	{
		Send +{up}+{up}+{up}
		Send +{end}
		Send {enter}th5{f9}
		Sleep 200
		Send ^{home}
		sleep 200

	}

	if CellBlock
		Sendinput Cytology slides and cell block are composed of a 
}

Else If Spectype = Urine {
	if NEG 
		SendInput Specimen consists mainly of mature squamous epithelium and rare benign urothelial cells.+{up 2}^b
	else if ATYP
		SendInput 249
		sleep 100
		Send {f9}
		
}
Else If Spectype in Peritoneal,Pelvic,Ascitic,Pleural {
	if NEG && CellBlock
		SendInput Cytology slides and cell block consist of sheets of mesothelial cells in a background of scattered macrophages and chronic inflammation.
	else if NEG
		SendInput Specimen consists of sheets of mesothelial cells in a background of scattered macrophages and chronic inflammation.
}


Else If Spectype = BreastCyst {
	SendInput ^{home}{down}+{down}
	sleep 200
	SendInput {delete}{enter}	
	
	Inputbox, UserInput, Diag,1=No History`n2=Collapsed`n3=Part Collapse`n4=Foam Only`n5=Complex`n6=Apocrine`n7=Inflamed`n8=Seroma`n9=Bloody, , , , , , , ,1
	
	If UserInput = 4
		SendInput BR1-4{F9}
		
}

Else If FNA && CellBlock
	SendInput Cytology slides and cell block consist of 
Else If FNA && CoreBx
	SendInput Cytology slides and core biopsy consist of 
Else If FNA && CellBlock && CoreBx
	SendInput Cytology slides, cell block and core biopsy consist of 
Else If FNA
	SendInput Specimen consists of 

return


CTFNA:
;--------------------------------------
Gosub ResultWindowActive

MouseClick, , AdequacyAlphaFieldX, AdequacyAlphaFieldY

Sleep 500
Send {enter}

;if AdequacyCode = 65


if AdequacyCode = 66 {
	Sendinput S{enter}
	MouseClick, , StatementAdeqFieldX, StatementAdeqFieldY, 2
	WinWaitOpenByImageSearch(EditWinLocX, EditWinLocY, iEditWindow)
	
	SendInput +{up}{BS}
	;Click %boldX%, %boldY%
	SendInput Immediate cytologic study for adequacy of specimen, using a Diff-Quick stain, was performed.   
	sleep 100
	SendInput %A_Space%FNA acceptable for further evaluation by 
	sleep 100
	SendInput %A_Space%%CTInitials%
	sleep 100
	Sendinput {f9}
	sleep 300
	sendinput +{UP 2}^B^B
	sleep 100
	Send !o
	WinWaitCloseByImageSearch(iEditWindow)
	GoSub CaseCharges
	CPTInput("88172-TC")
	
}

if AdequacyCode = 67 {
	Sendinput S{enter}
	MouseClick, , StatementAdeqFieldX, StatementAdeqFieldY, 2
	WinWaitOpenByImageSearch(EditWinLocX, EditWinLocY, iEditWindow)
	
	SendInput +{up}{BS}
	SendInput 67 
	sleep 200
	Sendinput {f9}	
	sleep 200
	SendInput %A_Space%%CTInitials%
	sleep 200
	sendinput {f9}
	sleep 200
	sendinput +{UP 2}^B^B
	sleep 200
	Send !o
	WinWaitCloseByImageSearch(iEditWindow)
	GoSub CaseCharges
	CPTInput("88172-TC")
}
AdequacyEntered := 1
return





CoreBXQuery:
;--------------------------------------
If FNA {
	Msgbox, 3, Core Bx, Core Bx Y/N
	
	IfMsgBox, Yes
		CoreBX = 1
	IfMsgBox, Cancel
		exit
	IfMsgBox, No
		CoreBX = 0 
}


CellBlockQuery:
;--------------------------------------
If NG || CN || FNA {
	If SpecDisc not contains Urine,CSF,Bladder {
		Msgbox, 3, Cell Block, Cell Block Y/N
		
		IfMsgBox, Yes
			CellBlock := 1
		IfMsgBox, Cancel
			exit
		IfMsgBox, No
			CellBlock := 
	}
}

return


WinWaitOpenByImageSearch(ByRef VarNameX, ByRef VarNameY, pic)
;-----------------------------------
{
	Loop {
		ImageSearch, VarNameX, VarNameY, 0, 0, A_ScreenWidth, A_ScreenHeight, %pic%
		sleep 100
	} until errorlevel = 0
	return
}

WinWaitCloseByImageSearch(pic)
;-----------------------------------
{
	Loop {
		sleep 100
		ImageSearch, VarNameX, VarNameY, 0, 0, A_ScreenWidth, A_ScreenHeight, %pic%
	} until errorlevel = 1
	return
}


; slide count
^!s::
;---------------------------------------
Click 232,60
WinWaitOpenByImageSearch(DSCWinLocX, DSCWinLocY, iSlideCountWindow)
;WinWaitOpenByAreaColor(DSCWinLocX, DSCWinLocY, 28, 157, 20, 0xFF00FF)
Send {down}
MouseClick, ,% DSCWinLocX + 150, % DSCWinLocY + 130, 2
Send {tab}
return


F3:: 
;Case charges-----------------------
GoSub ResultWindowActive

GoSub CaseCharges

If NG || CN || FNA 
	GoSub ICDEnter	

return



^!1::
;Modified NEG+ --------------------------
GoSub ResultWindowActive

SendInput !1{tab}

If NG || CN || FNA {
	NEG := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub ReformatReport
}


If GYN {
	NEG := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub VerifyDas
}

return


^!2::
;Modified NEG- ----------------------
GoSub ResultWindowActive

SendInput !2{tab}

If NG || CN || FNA {
	BNGN := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub ReformatReport
}

If GYN {
	NEG := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub VerifyDas
}
return



^!3::
;Modified HYST -------------------------
GoSub ResultWindowActive

Sendinput !3{tab}

If NG || CN || FNA {
	ATYP := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub ReformatReport
}

If GYN = 1 {
	NEG := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub VerifyDas
}

return



^!4::
;Modified ATR ------------------------
GoSub ResultWindowActive

Sendinput !4{tab}

If NG || CN || FNA {
	SIGN := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub ReformatReport
}


If GYN {	
	Click, %EndoAlphaFieldX%, %EndoAlphaFieldY%

	Send {enter}

	Sendinput {up 3}
	Click, %StatementAdeqFieldX%, %StatementAdeqFieldY%

	Sendinput {enter 2}
	WinWaitOpenByImageSearch(EditWinLocX, EditWinLocY, iEditWindow)
	Sendinput !o
	WinWaitCloseByImageSearch(iEditWindow)


	NEG := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub VerifyDas
}

return



^!5::
;Modified RECC ------------------
NoVerify := 1
GoSub ResultWindowActive

Sendinput !5{tab}

If NG || CN || FNA {
	SUSN := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub ReformatReport
}

Click, %DiagAlphaFieldX%, %DiagAlphaFieldY%
SendInput {enter} 
sleep 500

if GYN {
	Loop {
		Inputbox, UserInput, Reason?,
			(
			N - None
			I - Inflammation
			C - Candida
			P - Parakeratosis
			F - Shift in Flora
			T - Trichomonas
			), , , , , , , ,N
		if errorlevel = 1
			exit
		if UserInput not contains i,c,f,n,p,t
			continue
		If userinput = i
			SendInput u{up 9}
		If userinput = c
			SendInput u{up 11}
		If userinput = p
			SendInput u{up 7}
		If userinput = f
			SendInput u{up 10}
		If userinput = t
			SendInput u{up 4}
		If userinput = n
			break
		Break	
	}
}

Send {enter}{escape}
Send {down}{enter 2}
WinWaitOpenByImageSearch(EditWinLocX, EditWinLocY, iEditWindow)


If GYN {
	Send !o
	WinWaitCloseByImageSearch(iEditWindow)
	GoSub CaseCharges
	sleep 100
	CPTInput(88141)
	
	If Vaginal
		ICDInput(N76.89)
	else
		ICDInput("N72")
	
}
return



^!6::
;Modified ASCUS -----------------
GoSub ResultWindowActive
NoVerify := 1
Sendinput !6{tab}

If NG || CN || FNA {
	SUSM := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub ReformatReport
}

if GYN {
	Click, %HPVFieldX%, %HPVFieldY%
	Sendinput {enter}
	
	Loop {
		Inputbox, UserInput, HPV, 1-Geno 2-mRNA 3-N/A
		if errorlevel = 1
			exit
		if UserInput not contains 1,2,3
			continue
		if UserInput = 1
		{
			send {down 1}{enter} 
			break	
		}
		if UserInput = 2
		{
			send {down 2}{enter} 
			break	
		}
		if UserInput = 3
		{
			send {down 3}{enter}
			break	
		}	
	}
	
	GoSub CaseCharges
	CPTInput(88141)
	ICDInput("R87.610")
}
return


^!7::
;Modified ASCUS-H ----------------------------
GoSub ResultWindowActive
NoVerify := 1

Sendinput !7{tab}

If NG || CN || FNA {
	POS := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub ReformatReport
}

if GYN {
	GoSub CaseCharges
	CPTInput(88141)
	ICDInput("R87.611")
}

return



^!8::
;Modified LSIL ----------------------------
GoSub ResultWindowActive
NoVerify := 1

Sendinput !8{tab}

If NG || CN || FNA {
	NDIAG := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub ReformatReport
}

if GYN {
	GoSub CaseCharges
	CPTInput(88141)
	ICDInput("R87.612")
}
return



^!9::
;Modified HSIL ----------------------
GoSub ResultWindowActive
NoVerify := 1

Send !9{tab}

If NG || CN || FNA {
	ATYP := 1
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub ReformatReport
}

if GYN {
	GoSub CaseCharges
	Send !d{enter}
	Sleep 250
	SendInput CPT 88141
	Send {Enter 2}
	Sleep 250
	ICDInput("R87.613")
}
return



^!0::
;Modified UNSAT ---------------------
GoSub ResultWindowActive
NoVerify = 1

Sendinput !0{tab}

If NG || CN || FNA {
	
	GoSub CaseCharges
	GoSub ICDEnter
	GoSub ReformatReport
}

Click, %ReasonAlphaFieldX%, %ReasonAlphaFieldY%


Loop {
	Inputbox, UserInput, Reason?, 
		(
		S - Scant
		I - Obs Inf
		), , , , , , , ,S
	
	if errorlevel = 1
		exit
	if UserInput not contains s,i
		continue
	If userinput = S
	{
		Sendinput {p 6}
		break
	}
	If userinput = I
	{
		Sendinput {p 5}
		break
	}
}

Send {enter}{escape}{down 2}{enter 2}
WinWaitOpenByImageSearch(EditWinLocX, EditWinLocY, iEditWindow)
Send !o

if GYN {

	GoSub CaseCharges
	CPTInput(88141)
	ICDInput("R87.615")
}
return


^!e::
;Change to ENDO- -------------------------

Setkeydelay,75
Click, %EndoAlphaFieldX%, %EndoAlphaFieldY%
Sendinput {enter}{down}{enter}{down}{enter 2}!o
return


^!i::
;Infection NEG+ ------------------------------
GoSub ResultWindowActive
Sendinput !1{tab}

Click, %DiagAlphaFieldX%, %DiagAlphaFieldY%
SendInput {enter} 

Loop {
	Inputbox, UserInput, Infection?,
		(
		I - Inflammation
		C - Candida
		F - Shift in Flora
		T - Trichomonas
		V - Atrophic Vaginitis
		), , , , , , , ,I
	
	if errorlevel = 1
		exit
	if UserInput not contains i,c,f,t,v
		continue
	If userinput = i
	{
		sendinput {n 12}
		Send {enter}{escape}
	}
	If userinput = c
	{
		sendinput {n 4}
		Send {enter}{escape}
	}
	If userinput = f
	{
		sendinput {n 8}
		Send {enter}{escape}
	}
	If userinput = t
	{	
		sendinput u{up}
		Send {enter}{escape}
	}
	If userinput = v
	{
		sendinput {n 3}{enter}{escape}	
		Click, %EndoAlphaFieldX%, %EndoAlphaFieldY%
		sleep 200
		Sendinput {enter}
		Sendinput {up}		
		Sendinput {up}		
		Sendinput {up}		
		Sendinput {enter}		
		sendinput {escape}{down}{enter 2}!o
		WinWaitCloseByImageSearch(iEditWindow)
		Sendinput {down}
		ATROV := 1
	}
	break
}

Send {down}{enter 2}
WinWaitOpenByImageSearch(EditWinLocX, EditWinLocY, iEditWindow)
Send !o
WinWaitCloseByImageSearch(iEditWindow)
GoSub CaseCharges
GoSub ICDEnter
GoSub VerifyDas
return



;^b::
;Bold button shortcut--------------------------------------------------
;click %BoldX%, %BoldY%
;return



^!m::
;Common case shortcut--------------------------------------------------
Sendinput Please also refer to specimen number _.
Sendinput {left}{bs}
CommonEntered = 1

if clipboard contains -14-
	Sendinput %clipboard%

Inputbox, UserInput, Enter Common Case, Enter Common Case, , , , , , , ,

if errorlevel = 1
	exit

; Get individual pieces of case number
FullSlideNum = %UserInput%							
StringMid, SiteNum, FullSlideNum, 1, 2
StringMid, SpecClass, FullSlideNum, 3, 2
StringMid, SpecYear, FullSlideNum, 5, 2
StringMid, CaseNum, FullSlideNum, 7, 6

Loop { 
	StringLeft, CaseNumLeadingChar, CaseNum, 1
	If CaseNumLeadingChar = 0
		StringTrimLeft, CaseNum, CaseNum, 1
	else 
		break
}

Sendinput %Sitenum%-%SpecClass%-%SpecYear%-%Casenum%

return



F6::
Gosub CTFNA
return

F7::
; Assign case

;click %TTButLocX%, %TTButLocY%
sendinput !tt
sleep 200
Sendinput {tab}

If !MDName
	Loop {
		Inputbox, UserInput, Enter MD, MD, wait until ORV open, , , , , , , ,%MDName%
		if errorlevel = 1
			exit
		if UserInput not contains %MDList%
			continue
		MDName = %UserInput%			
		break
	}

Sendinput %MDName%{enter}
sleep 200
Sendinput {down}
return


F8::
; Open case assignment------------------------------------------
Click 307,63
WinWaitOpenByImageSearch(SelectReportsWindowX, SelectReportsWindowY, iSelectReportsWindow)
Sendinput {tab}CL Cyto Path Review{enter}
WinWaitOpenByImageSearch(CLCytopathReviewButtonX, CLCytopathReviewButtonY, iCLCytopathReviewButton)
WinWaitOpenByImageSearch(TTButLocX, TTButLocY, iTransferReportButton)
Mouseclick, L, % TTButLocX + 620, % TTButLocY + 90, 2
Sendinput {end}
return




ResultWindowActive:
;--------------------------------------
WinActivate, Appbar P819 NSHS_NY - Citrix online plug-in
WinMaximize, Appbar P819 NSHS_NY - Citrix online plug-in
PixelSearch,,, SpecDescFieldX, SpecDescFieldY, % SpecDescFieldX + 2, % SpecDescFieldY+ 2, 0xC8D0D4, , Fast
if !ErrorLevel ; if grey found
	MouseClick, , SpecDescFieldX, % SpecDescFieldY - 50
if ErrorLevel ; if grey not found
	MouseClick, , SpecDescFieldX, SpecDescFieldY
return



HistoryWindowActive:
;--------------------------------------
WinActivate, Appbar Prod NSHS_NY - Citrix online plug-in
WinMaximize, Appbar Prod NSHS_NY - Citrix online plug-in
Click 843, 276
return


CTAdequacyQuery:
;--------------------------------------
if FNA {
	Msgbox, 3, CT Adequacy, CT Adequacy Y/N
	IfMsgBox, Yes
		CTAdequacy := 1
	IfMsgBox, No
		CTAdequacy := 0
	IfMsgBox, Cancel
		exit
}


if CTAdequacy {
	; Enter adequacy code
	Loop {
		Inputbox, UserInput, Adequacy Code, Adequacy Code, , , , , , , ,66
		if errorlevel = 1
			exit
		if UserInput not in 65,66,67
			continue
		if UserInput in 65,66,67 {
			AdequacyCode = %UserInput%			
			break
		}
	}
	
	; Enter initials of cytotech
	Loop	{
		Inputbox, UserInput, CT Initials, CT Initials, , , , , , , ,md
		if errorlevel = 1
			exit
		if UserInput not in %CTList%
			continue
		if UserInput in %CTList%
		{
			CTInitials = %UserInput%			
			break
		}
	}
}
return

PrevAbnCheck:
;--------------------------------------
if GYN {
	PixelGetColor, color, PrevAbnButtLocX, PrevAbnButtLocY
	if color = 0x0000FF
	{
		Click %PrevAbnButtLocX%, %PrevAbnButtLocY%
		Gosub HistoryWindowActive
		Sleep 250
		Send {home}{right}
		gosub cliploop
		
		;CurrentDate = %A_Now%
		;FormatTime, CurrentDate,, ShortDate
		
		;if clipboard = less than a year 
		;PreviousAbnLIS = 1
		NoVerify = 1
		GoSub ResultWindowActive
	}
}
return


^!r::
reload
return


;EXPERIMENTAL-------------------------------------------------------------------
;===========================================================

^!h::
ImageSearch, XX, YY, 0, 0, 1200, 1200, editwindow.png
if errorlevel = 1
	msgbox, not found
if errorlevel = 0 
	msgbox, found
return


; Example: On-screen display (OSD) via transparent window

^!o::
CustomColor = EEAA99  
Gui +LastFound +AlwaysOnTop
Gui, Color, %CustomColor%
Gui, Font, s16  ; Set a large font size (32-point).
Gui, Show, X700 Y15 H100 W500 NoActivate
Gui, Add, Text, vCaseNumOSD, %CaseNum%
Gui, Add, Text, vHPVPos, %CaseNum%
; Make all pixels of this color transparent and make the text itself translucent (150):
WinSet, TransColor, %CustomColor% 150
SetTimer, UpdateOSD, 500
Gosub UpdateOSD  ; Make the first update immediate rather than waiting for the timer.
return

UpdateOSD:
GuiControl,, CaseNumOSD, %CaseNum%
If HPVPositive = 1
	Gui, Add, Text, , HPV Positive
return


^!g::
Gui, Add, Button, default, Cell
Gui, Show,, Simple Input Example

return

;DEPRECATED
;===========================================================

!^n::
Gui, Add, Radio, vMyRadio, Sample radio1
Gui, Add, Radio,, Sample radio2
;Gui, Tab  ; i.e. subsequently-added controls will not belong to the tab control.
Gui, Add, Button, default xm, OK  ; xm puts it at the bottom left corner.
Gui, Show
return

ButtonOK:
GuiClose: 
GuiEscape:
Gui, Submit  ; Save each control's contents to its associated variable.
MsgBox You entered:`n%MyCheckbox%`n%MyRadio%`n%MyEdit%
Exit


5MinAlarm:
;--------------------------------------
ElapsedTime := A_TickCount - StartTime
If ElapsedTime > 300000 {
	SoundBeep, , 250
	Sleep 100
	SoundBeep, , 250
}
return

HistCheck:
;---------------------------------------
Gosub HistoryWindowActive
Send {home}{pgup}
Loop {
	Sleep 100
	Send {right}^c
	;Sleep 100
	If clipboard = %HistDate%
		SameDate = 1
	HistDate = %clipboard%
	Send {right}^c
	;Sleep 100
	If clipboard = %HistDiag%
		SameDiag = 1
	If SameDate = 1	
		If SameDiag = 1
		{
			GoSub ResultWindowActive
			break
		}
	HistDiag = %clipboard%
	SameDate = 0
	SameDiag = 0
	If HistDiag contains POS, ECA
	{
		PreviousAbnLIS = 1
		NoVerify = 1
		; If HistDate = Less than a year ago
		; Labnote = 1
	}
	Send {down}{home}
}		
return

CalibrateMouseWithButtonColor(ButtonName, ButtonColor)
;--------------------------------------
{
	If %ButtonName%LocX =
	{
		Traytip, Calibrate, Move mouse to %ButtonColor% pixel on %ButtonName%
		
		Loop
		{
			sleep 10
			MouseGetPos, MouseX, MouseY
			PixelGetColor, color, %MouseX%, %MouseY%
		} Until color = ButtonColor
		%ButtonName%LocX := MouseX
		%ButtonName%LocY := MouseY
		return 
	} 
}


WinWaitByPixelColor(LocX, LocY, PixelColor)
;--------------------------------------
{
	Loop
	{
		if a_index > 3000
		{	
			traytip, Win Not Found, Exiting
			exit
		}
		PixelGetColor, color, LocX, LocY
		sleep 100
	} Until color = PixelColor
	return
}


WinWaitOpenByAreaColor(ByRef AreaVarNameX, ByRef AreaVarNameY, AreaCenterLocX, AreaCenterLocY, AreaSize, AreaColor)
;-----------------------------------
{
	
	Loop
	{
		PixelSearch, AreaVarNameX, AreaVarNameY, AreaCenterLocX - AreaSize, AreaCenterLocY - AreaSize, AreaCenterLocX + AreaSize, AreaCenterLocY + AreaSize, AreaColor, , Fast
		sleep 100
	} until errorlevel = 0
	return
}

WinWaitCloseByPixelColor(PixelX, PixelY, PixelColor)
;-----------------------------------
{
	Loop
	{
		sleep 100
		PixelGetColor, color, PixelX, PixelY
	} until color != PixelColor
	return	
}


TrimTrailingSpaces()
;-----------------------------------
{
	Loop
	{ 
		Sendinput +^{HOME}
		Gosub Cliploop
		;sleep 300
		;pause
		;ImageSearch, , , 550, 300, 650, 400, %pic%
		;if errorlevel = 0
		;{
		;	sendinput {enter}
		;	continue
		;}
		StringRight, TrailingChar, clipboard, 2
		If TrailingChar contains %A_SPACE%,`r,`n
		{
			Sendinput ^{end}{BS}
			sleep 10
		}
		else		
			break
	}
	Sendinput ^v
}


ClipLoop:
;--------------------------------------
;Empty the clipboard
Loop
{
	clipboard = 
	SendInput ^c
	ClipWait, 1
	
	if ErrorLevel = 1
	{
		SendInput ^c
		ClipWait, 1
	}
	
	if ErrorLevel = 1
	{
		MsgBox, 3, Copy to clipboard failed. Try again?
		IfMsgBox Yes
			gosub cliploop
		IfMsgBox No
			break
		IfMsgBox Cancel
			break
	}	
} until ErrorLevel = 0

if clipboard is space
{
	TrayTip, clipboard is space, clipboard is space
; 	msgbox, clipboard is space
	gosub cliploop
}

if clipboard = 
{
;	msgbox, clipboard empty
	gosub cliploop
}


return
