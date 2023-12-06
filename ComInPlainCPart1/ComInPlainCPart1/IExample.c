#include<windows.h>
#include<objbase.h>
#include "IExample.h"


HRESULT STDMETHODCALLTYPE QueryInterface(LPVOID, REFIID, LPVOID*);
HRESULT STDMETHODCALLTYPE AddRef(LPVOID);
HRESULT STDMETHODCALLTYPE Release(LPVOID);

HRESULT STDMETHODCALLTYPE FactoryQueryInterface(LPVOID, REFIID, LPVOID*);
HRESULT STDMETHODCALLTYPE FactoryAddRef(LPVOID);
HRESULT STDMETHODCALLTYPE FactoryRelease(LPVOID);
HRESULT STDMETHODCALLTYPE FactoryCreateInstance(LPVOID, LPVOID, REFIID, LPVOID*);
HRESULT STDMETHODCALLTYPE FactoryLockServer(LPVOID, BOOL);


HRESULT STDMETHODCALLTYPE SetString(LPVOID, char*);
HRESULT STDMETHODCALLTYPE GetString(LPVOID, char*, size_t);


static const IClassFactoryVtbl IClassFactory_Vtbl = {
	FactoryQueryInterface,
	FactoryAddRef,
	FactoryRelease,
	FactoryCreateInstance,
	FactoryLockServer,
};
static IClassFactory classFactory;
static const IExampleVtbl IExample_Vtbl = { QueryInterface,AddRef,Release,SetString,GetString };
static DWORD OutstandingObjects;
static DWORD LockCount;

typedef struct {
	IExampleVtbl* lpVtbl;
	DWORD count;
	char buffer[80];
}IExampleForDll, * IExamplePtrForDll;


HRESULT STDMETHODCALLTYPE SetString(LPVOID this, char* buffer) {
	size_t i = lstrlenA(buffer);
	if (i > 79)
		i = 79;
	memcpy(((IExamplePtrForDll)this)->buffer, buffer, i);
	((IExamplePtrForDll)this)->buffer[i] = 0;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE GetString(LPVOID this, char* buffer, size_t size) {
	size_t i = lstrlenA(((IExamplePtrForDll)this)->buffer);
	if (size > i)
		size = i;
	memcpy(buffer, ((IExamplePtrForDll)this)->buffer, size);
	buffer[size] = 0;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE QueryInterface(LPVOID this, REFIID riid, LPVOID* ppv) {
	if (!(IsEqualCLSID(&IID_IUnknown, riid) || IsEqualCLSID(&IID_IExample, riid))) {
		*ppv = 0;
		return E_NOINTERFACE;
	}
	*ppv = this;
	((IExamplePtrForDll)this)->lpVtbl->AddRef(this);
	return S_OK;
}
HRESULT STDMETHODCALLTYPE AddRef(LPVOID this) {
	((IExamplePtrForDll)this)->count++;
	return ((IExamplePtrForDll)this)->count;
}
HRESULT STDMETHODCALLTYPE Release(LPVOID this) {
	((IExamplePtrForDll)this)->count--;
	if (((IExamplePtrForDll)this)->count == 0) {
		GlobalFree(this);
		InterlockedDecrement(&OutstandingObjects);
		return S_OK;
	}
	return ((IExamplePtrForDll)this)->count;
}



HRESULT STDMETHODCALLTYPE FactoryQueryInterface(LPVOID this, REFIID riid, LPVOID* ppv) {
	if (!(IsEqualCLSID(&IID_IUnknown, riid) || IsEqualCLSID(&IID_IClassFactory, riid))) {
		*ppv = 0;
		return E_NOINTERFACE;
	}
	*ppv = this;
	((IClassFactory*)this)->lpVtbl->AddRef((IClassFactory*)this);
	return S_OK;
}
ULONG STDMETHODCALLTYPE FactoryAddRef(LPVOID this) {
	InterlockedIncrement(&OutstandingObjects);
	return 1;
}
ULONG STDMETHODCALLTYPE FactoryRelease(LPVOID this) {
	return(InterlockedDecrement(&OutstandingObjects));
}
HRESULT STDMETHODCALLTYPE FactoryCreateInstance(LPVOID this, LPVOID pUnkOuter, REFIID riid, LPVOID* ppv) {
	HRESULT hr;
	IExamplePtrForDll example;
	*ppv = 0;
	if (pUnkOuter) {
		OutputDebugStringA("aggregation is not supported for now\n");
		return CLASS_E_NOAGGREGATION;
	}
	if (!(example = GlobalAlloc(GMEM_FIXED | GMEM_ZEROINIT, sizeof(IExampleForDll)))) {
		OutputDebugStringA("[ComInPlainCPart1 -- FactoryCreateInstance] memory allocate failed for IExample struct\n");
		hr = E_OUTOFMEMORY;
	}
	else {
		example->lpVtbl = &IExample_Vtbl;
		example->lpVtbl->AddRef(example);
		hr = example->lpVtbl->QueryInterface(example, riid, ppv);
		// 上面的QueryInterface函数会再次增加example的引用计数，如果发生错误，则引用计数仍然是1，下面的Release函数将会销毁example
		example->lpVtbl->Release(example);
		if (SUCCEEDED(hr))
			InterlockedIncrement(&OutstandingObjects);
	}
	return hr;
}

// 如果客户端想把COM dll锁在自己内存里，就可以调用这个函数
HRESULT STDMETHODCALLTYPE FactoryLockServer(LPVOID this, BOOL fLock) {
	if (fLock)
		InterlockedIncrement(&LockCount);
	else
		InterlockedDecrement(&LockCount);
	return S_OK;
}

HRESULT PASCAL DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
	HRESULT hr;
	if (IsEqualCLSID(&CLSID_IExample, rclsid))
		hr = FactoryQueryInterface(&classFactory, riid, ppv);
	else {
		*ppv = 0;
		hr = CLASS_E_CLASSNOTAVAILABLE;
	}
	return hr;
}

HRESULT PASCAL DllCanUnloadNow(void) {
	return ((OutstandingObjects | LockCount) ? S_FALSE : S_OK);
}

BOOL WINAPI DllMain(HINSTANCE instance, DWORD fdwReason, LPVOID lpvReserved) {
	switch (fdwReason) {
	case DLL_PROCESS_ATTACH: {
		// MessageBoxA(NULL, "hold on for windbg", "debug", MB_OK);
		OutstandingObjects = LockCount = 0;
		classFactory.lpVtbl = &IClassFactory_Vtbl;

		DisableThreadLibraryCalls(instance);
	}
	}
	return TRUE;
}