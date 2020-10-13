// ---------------------------------------------------------------------------------
// CallingLabVIEWFromCPPUsingActiveX.cpp : 
//
// Sample program to show how to call some simple VIs in LabVIEW by way of ActiveX.
// This source code is an automation client and it calls VIs in LabVIEW way way of
// the LabVIEW automation server.
// ---------------------------------------------------------------------------------

// NOTE: LabVIEW must be installed and the #import line below must point to the
// correct LabVIEW.tlb.

#import "C:\Program Files (x86)\National Instruments\LabVIEW 2020\resource\labview.tlb"

#include <iostream>
#include <main.h>
#include "Labview_Interface.h"

extern "C" double *fn_netsim_labview_run();
extern "C" void fn_netsim_labview_init();
extern "C" void fn_netsim_labview_finish();


LabVIEW::_ApplicationPtr pLVApp;




extern "C" void fn_netsim_labview_init()
{
	CoInitialize(NULL);
	pLVApp.CreateInstance("LabVIEW.Application");
	if (NULL == pLVApp) {
		std::cout << "CallsVIThroughActiveX : An error occurred getting pLVApp";
	}

}

extern "C" void fn_netsim_labview_finish()
{
	if (NULL != pLVApp)
		pLVApp->Quit();
	CoUninitialize();
}

extern "C" double *fn_netsim_labview_run()
{
	if (NULL == pLVApp) {
		std::cout << "CallsVIThroughActiveX : An error occurred getting pLVApp";
		return 0;
	}
	static const char* kSineRelativeVIPath = "\\examples\\Random_Number.vi";
	HRESULT hr = S_OK;

	char password[60] = "";
	//VARIANT_BOOL reserveForCall = VARIANT_FALSE; // reserveForCall must be true when using "Call" or "Call2" to call a VI
	bool reserveForCall = false;
	unsigned long options = 0x01;

	char viPath[MAX_PATH];
	strcpy(viPath, pLVApp->ApplicationDirectory);
	strcat(viPath, kSineRelativeVIPath);

	// Assign an object reference to pVI.
	LabVIEW::VirtualInstrumentPtr pVI;
	pVI.CreateInstance("LabVIEW.VirtualInstrument");
	//pVI = pLVApp->GetVIReference(LPCTSTR(viPath), LPCTSTR(password), 0, 0x01);
	pVI = pLVApp->GetVIReference(viPath, "", 0, 0x01);
	if (NULL == pVI) {
		std::cout << "AddTwoNumbers : An error occurred getting pVI";
		return 0;
	}

	hr = pVI->OpenFrontPanel(VARIANT_TRUE, LabVIEW::eVisible);

	//hr = pVI->SetControlValue("NUM1", x);
	//double value = 10;
	//hr = pVI->SetControlValue("Units", value);
	VARIANT_BOOL asynchronousCall = VARIANT_TRUE;
	hr = pVI->Run(asynchronousCall);

	VARIANT result = pVI->GetControlValue("Output");
	double val_result = result.dblVal;
	//pstruEventDetails->dEventTime = result.dblVal;

	
	return &val_result;


}