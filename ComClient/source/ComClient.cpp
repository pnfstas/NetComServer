#import "C:\\inetpub\\wwwroot\\netcomserver\\ComInprocServer\\lib\\NetComServer.tlb" named_guids raw_interfaces_only

#include <Windows.h>
#include <stdexcept>
#include <typeinfo>
#include "../../../../../../inetpub/wwwroot/netcomserver/ComInprocServer/include/NetComServer.h"

using namespace std;

typedef IComServer *LPCOMSERVER;

BOOL CheckResultAndShowError(HRESULT hr)
{
	BOOL success = !FAILED(hr);
	if(!success)
	{
		DWORD dwErrorCode = (DWORD)hr;
		WCHAR lpBuffer[256];
		DWORD dwSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, NULL, dwErrorCode,
			MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), lpBuffer, sizeof(lpBuffer) / sizeof(WCHAR), NULL);
		if(dwSize > 0)
		{
			wstring wstrErrorMessage(lpBuffer, dwSize);
			wstrErrorMessage = L"Create ComServer instanse interrupted with the following error:\n" + wstrErrorMessage;
			_cputws(wstrErrorMessage.c_str());
		}
	}
	return success;
}
void CreateComServer(LPCOMSERVER* ppComServer)
{
	if(ppComServer != NULL)
	{
		*ppComServer = NULL;
		try
		{
			//HRESULT hr = CoCreateInstance(CLSID_ComServer, NULL, CLSCTX_LOCAL_SERVER, IID_IComServer, (LPVOID*)ppComServer);
			HRESULT hr = CoCreateInstance(CLSID_ComServer, NULL, CLSCTX_INPROC_SERVER, IID_IComServer, (LPVOID*)ppComServer);
			if(!CheckResultAndShowError(hr))
			{
				*ppComServer = NULL;
			}
			/*
			IClassFactory* pClassFactory = NULL;
			HRESULT hr = CoGetClassObject(CLSID_ComServer, CLSCTX_INPROC_SERVER, NULL, IID_IClassFactory, (LPVOID*)&pClassFactory);
			if(!FAILED(hr))
			{
				hr = pClassFactory->CreateInstance(NULL, IID_IComServer, (LPVOID*)ppComServer);
				if(!CheckResultAndShowError(hr))
				{
					*ppComServer = NULL;
				}
			}
			*/
		}
		catch(std::exception e)
		{
			*ppComServer = NULL;
			OutputDebugStringA(e.what());
			throw;
		}
	}
}
int main(int argc, char *argv[])
{
	try
	{
		HWND hwndConsole = GetConsoleWindow();
		if(hwndConsole != NULL)
		{
			ShowWindow(hwndConsole, SW_MINIMIZE);
		}
		HRESULT hr = CoInitializeEx(0, COINITBASE_MULTITHREADED);
		if(CheckResultAndShowError(hr))
		{
			IComServer* pComServer = NULL;
			//HRESULT hr = CoCreateInstance(CLSID_ComServer, NULL, CLSCTX_LOCAL_SERVER, IID_IComServer, (LPVOID*)&pComServer);
			HRESULT hr = CoCreateInstance(CLSID_ComServer, NULL, CLSCTX_INPROC_SERVER, IID_IComServer, (LPVOID*)&pComServer);
			if(CheckResultAndShowError(hr) && pComServer != NULL)
			{
				BSTR lpwszAssembly = SysAllocString(L"C:\\inetpub\\wwwroot\\forumengine\\ForumEngine.dll");
				BSTR lpwszTypeName = SysAllocString(L"ForumEngine.ForumEngineUser");
				hr = pComServer->LoadAssembly(lpwszAssembly);
				if(!FAILED(hr))
				{
					IDispatch* pdisp = NULL;
					hr = pComServer->CreateNetInstance(lpwszTypeName, &pdisp);
				}
				if(FAILED(hr))
				{
					DISPID dispid1 = -1, dispid2 = -1;
					BSTR lpwszName1 = SysAllocString(L"LoadAssembly");
					BSTR lpwszName2 = SysAllocString(L"CreateNetInstance");
					HRESULT hr1 = pComServer->GetIDsOfNames(IID_NULL, &lpwszName1, 1, LOCALE_SYSTEM_DEFAULT, &dispid1);
					HRESULT hr2 = pComServer->GetIDsOfNames(IID_NULL, &lpwszName2, 1, LOCALE_SYSTEM_DEFAULT, &dispid2);
					hr = FAILED(hr1) ? hr1 : FAILED(hr2) ? hr2 : S_OK;
					if(!FAILED(hr) && dispid1 > 0  && dispid2 > 0)
					{
						VARIANTARG varg;
						VariantInit(&varg);
						varg.vt = VT_BSTR;
						varg.bstrVal = SysAllocString(L"C:\\inetpub\\wwwroot\\forumengine\\ForumEngine.dll");
						DISPPARAMS dispparams = { &varg, NULL, 1, 0 };
						VARIANT varResult;
						EXCEPINFO excepinfo;
						UINT uiArgErr = 0;
						VariantInit(&varResult);
						hr = pComServer->Invoke(dispid1, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispparams, &varResult, &excepinfo, &uiArgErr);
						SysFreeString(varg.bstrVal);
						if(!FAILED(hr))
						{
							VariantInit(&varg);
							varg.vt = VT_BSTR;
							varg.bstrVal = SysAllocString(L"ForumEngine.ForumEngineUser");
							//dispparams = { &varg, NULL, 1, 0 };
							uiArgErr = 0;
							VariantInit(&varResult);
							hr = pComServer->Invoke(dispid2, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispparams, &varResult, &excepinfo, &uiArgErr);
							SysFreeString(varg.bstrVal);
							if(!FAILED(hr))
							{
							}
						}
					}
				}
			}
			CoUninitialize();
		}
	}
	catch(std::exception e)
	{
		OutputDebugStringA(e.what());
		throw;
	}
}
