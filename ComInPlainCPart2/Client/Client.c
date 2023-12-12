#include<windows.h>
#include<stdio.h>
#include "../ComInPlainCPart2/IExample.h"

VOID OutputHrMessage(HRESULT res) {
	void* pMsg;

	FormatMessageA(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, NULL, res,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPSTR)&pMsg, 0, NULL);

	printf("[HRESULT %d] %s\n",res, (char*)pMsg);
}


int main() {

	HRESULT hr = CoInitialize(NULL);
	if (SUCCEEDED(hr)) {
		IClassFactory* classFactory;
		hr = CoGetClassObject(&CLSID_IExample, CLSCTX_LOCAL_SERVER, 0, &IID_IClassFactory, (LPVOID*)&classFactory);
		 
		if (FAILED(hr)) {
			MessageBox(0, L"Can't get IClassFactory", L"CoGetClassObject error", MB_OK | MB_ICONEXCLAMATION);
		}
		else {
			IExample* example;
			hr = classFactory->lpVtbl->CreateInstance(classFactory, 0, &IID_IExample,&example);
			// 现在已经用不到factoryClass，释放掉即可
			classFactory->lpVtbl->Release(classFactory);
			if (FAILED(hr)) {
				MessageBox(0, L"Can't get IExample", L"CreateInstance error", MB_OK | MB_ICONEXCLAMATION);
			}
			else {
				// 参数类型需要进行修改
				BSTR strPtr;
				char inputString[] = "this is test text!";

				strPtr = SysAllocStringLen(0, lstrlenA(inputString));
				MultiByteToWideChar(CP_ACP, 0, inputString, lstrlenA(inputString), strPtr, lstrlenA(inputString));

				example->lpVtbl->SetString(example, strPtr);
				DWORD64 whatEver = 0;
				// 只要能装得下8bytes的地址就行
				BSTR* outStr = &whatEver;
				example->lpVtbl->GetString(example, outStr);
				
				wprintf(L"[buffer from COM object] %s\n", (WCHAR*)(*outStr));
				example->lpVtbl->Release(example);
				SysFreeString(strPtr);
			}
		}
		 CoUninitialize();
		
	}
	else
		MessageBox(0, L"Can't initialize COM", L"CoInitialize error", MB_OK | MB_ICONEXCLAMATION);
	return 0;
}