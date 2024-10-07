#import "C:\\inetpub\\wwwroot\\netcomserver\\ComInprocServer\\lib\\NetComServer.tlb" named_guids raw_interfaces_only

#include <Windows.h>
#include <stdexcept>
#include <typeinfo>
#include "../../ComInprocServer/include/NetComServer.h"

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
			HRESULT hr = CoCreateInstance(CLSID_ComServer, NULL, CLSCTX_LOCAL_SERVER, IID_IComServer, (LPVOID*)&pComServer);
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
				SysFreeString(lpwszAssembly);
				SysFreeString(lpwszTypeName);
			}
			CoUninitialize();
		}
	}
	catch(exception e)
	{
		OutputDebugStringA(e.what());
		throw;
	}
}
