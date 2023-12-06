#include<windows.h>
#include<stdio.h>
#include "../ComInPlainCPart1/IExample.h"

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
		hr = CoGetClassObject(&CLSID_IExample, CLSCTX_INPROC_SERVER, 0, &IID_IClassFactory, (LPVOID*)&classFactory);
		 
		if (FAILED(hr)) {
			MessageBox(0, "Can't get IClassFactory", "CoGetClassObject error", MB_OK | MB_ICONEXCLAMATION);
		}
		else {
			IExample* example;
			hr = classFactory->lpVtbl->CreateInstance(classFactory, 0, &IID_IExample,&example);
			// 现在已经用不到factoryClass，释放掉即可
			classFactory->lpVtbl->Release(classFactory);
			if (FAILED(hr)) {
				MessageBox(0, "Can't get IExample", "CreateInstance error", MB_OK | MB_ICONEXCLAMATION);
			}
			else {
				char buffer[80];
				
				example->lpVtbl->SetString(example, "this is test text!");
				
				example->lpVtbl->GetString(example, buffer, sizeof(buffer));
				
				printf("[buffer from COM object] %s\n", buffer);
				example->lpVtbl->Release(example);
				
			}
		}
		 CoUninitialize();
		
	}
	else
		MessageBox(0, "Can't initialize COM", "CoInitialize error", MB_OK | MB_ICONEXCLAMATION);
	return 0;
}