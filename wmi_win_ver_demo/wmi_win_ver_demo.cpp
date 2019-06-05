// wmi_win_ver_demo.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <Windows.h>
#include <Wbemidl.h>
#include <atlstr.h>
#include <comdef.h>

#pragma comment(lib, "wbemuuid.lib")

bool get_system_internal_ver(CString &ver)
{
	HRESULT hres;

	hres = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(hres))
	{
		OutputDebugString(L"CoInitializeEx failed");
		return false;
	}

	hres = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT,
		RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
	if (FAILED(hres))
	{
		OutputDebugString(L"CoInitializeSecurity failed");
		CoUninitialize();
		return false;
	}

	IWbemLocator *pLoc = NULL;
	hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc);
	if (FAILED(hres))
	{
		OutputDebugString(L"CoCreateInstance failed");
		CoUninitialize();
		return false;
	}

	IWbemServices *pSvc = NULL;
	hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL,
		0, 0, &pSvc);
	if (FAILED(hres))
	{
		OutputDebugString(L"ConnectServer failed");
		pLoc->Release();
		CoUninitialize();
		return false;
	}

	hres = CoSetProxyBlanket(pSvc, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, NULL,
		RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE,NULL, EOAC_NONE);
	if (FAILED(hres))
	{
		OutputDebugString(L"CoSetProxyBlanket failed");
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return false;
	}

	IEnumWbemClassObject *pEnumerator = NULL;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t("SELECT * FROM Win32_OperatingSystem"),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator
	);
	if (FAILED(hres))
	{
		OutputDebugString(L"ExecQuery failed");
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return false;
	}

	IWbemClassObject *pclsObj = NULL;
	ULONG uReturn = 0;
	while (pEnumerator)
	{
		HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
		if (0 == uReturn)
			break;

		VARIANT vtProp;
		hr = pclsObj->Get(L"Version", 0, &vtProp, 0, 0);
		ver = vtProp.bstrVal;
		VariantClear(&vtProp);
		pclsObj->Release();
	}

	pSvc->Release();
	pLoc->Release();
	pEnumerator->Release();
	CoUninitialize();

	return true;
}

int main()
{
	CString ver;

	if (get_system_internal_ver(ver))
		OutputDebugString(ver);

	return 0;
}
