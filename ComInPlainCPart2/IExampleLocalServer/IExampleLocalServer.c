#include<windows.h>
#include<objbase.h>
#include "../ComInPlainCPart2/IExample.h"

// Global variable area
ITypeInfo* gTypeInfo;


HRESULT STDMETHODCALLTYPE QueryInterface(LPVOID, REFIID, LPVOID*);
HRESULT STDMETHODCALLTYPE AddRef(LPVOID);
HRESULT STDMETHODCALLTYPE Release(LPVOID);

HRESULT STDMETHODCALLTYPE GetTypeInfoCount(LPVOID this, UINT* pCount);
HRESULT STDMETHODCALLTYPE GetTypeInfo(LPVOID this, UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
HRESULT STDMETHODCALLTYPE GetIDsOfNames(LPVOID this, REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
HRESULT STDMETHODCALLTYPE Invoke(LPVOID this, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr);

HRESULT STDMETHODCALLTYPE FactoryQueryInterface(LPVOID, REFIID, LPVOID*);
HRESULT STDMETHODCALLTYPE FactoryAddRef(LPVOID);
HRESULT STDMETHODCALLTYPE FactoryRelease(LPVOID);
HRESULT STDMETHODCALLTYPE FactoryCreateInstance(LPVOID, LPVOID, REFIID, LPVOID*);
HRESULT STDMETHODCALLTYPE FactoryLockServer(LPVOID, BOOL);

HRESULT STDMETHODCALLTYPE SetString(LPVOID, BSTR);
HRESULT STDMETHODCALLTYPE GetString(LPVOID, BSTR*);


static const IClassFactoryVtbl IClassFactory_Vtbl = {
	FactoryQueryInterface,
	FactoryAddRef,
	FactoryRelease,
	FactoryCreateInstance,
	FactoryLockServer,
};
static IClassFactory classFactory;
static const IExampleVtbl IExample_Vtbl = { QueryInterface,AddRef,Release,GetTypeInfoCount,GetTypeInfo,GetIDsOfNames,Invoke,SetString,GetString };
static int OutstandingObjects;
static DWORD LockCount;

// 修改为自动变量
typedef struct {
	IExampleVtbl* lpVtbl;
	DWORD count;
	BSTR string;
}IExampleForDll, * IExamplePtrForDll;

void ShowErrorMessage(LPCTSTR header, HRESULT hr)
{
	void* pMsg;

	FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, NULL, hr,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPTSTR)&pMsg, 0, NULL);

	MessageBoxA(NULL, pMsg, "COM server message", MB_ICONERROR);

	LocalFree(pMsg);
}

HRESULT STDMETHODCALLTYPE SetString(LPVOID this, BSTR str) {
	if (!str)
		return E_POINTER;
	// 如果原来string成员变量里面有东西，就先释放掉
	if (((IExamplePtrForDll)this)->string)
		SysFreeString(((IExamplePtrForDll)this)->string);

	if (!(((IExamplePtrForDll)this)->string = SysAllocStringLen(str, SysStringLen(str))))
		return TBSIMP_E_OUT_OF_MEMORY;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE GetString(LPVOID this, BSTR* strPtr) {
	if (!strPtr)
		return E_POINTER;


	if (!(*strPtr = SysAllocStringLen(((IExamplePtrForDll)this)->string, SysStringLen(((IExamplePtrForDll)this)->string))))
		return TBSIMP_E_OUT_OF_MEMORY;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE QueryInterface(LPVOID this, REFIID riid, LPVOID* ppv) {
	// 这里之前忘了加IID_IDispatch的判断了，坑死我了，vbs脚本一直报错
	if (!(IsEqualCLSID(&IID_IUnknown, riid) || IsEqualCLSID(&IID_IExample, riid) || IsEqualCLSID(&IID_IDispatch, riid))) {
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
		// 释放成员变量
		if (((IExamplePtrForDll)this)->string)
			SysFreeString(((IExamplePtrForDll)this)->string);

		GlobalFree(this);
		InterlockedDecrement(&OutstandingObjects);
		if (!(InterlockedAnd(&OutstandingObjects, 0xFFFFFFFF) | InterlockedAnd(&LockCount, 0xFFFFFFFF)))
			PostQuitMessage(0);
		return S_OK;
	}
	return ((IExamplePtrForDll)this)->count;
}

HRESULT loadTypeInfo() {
	HRESULT hr;
	LPTYPELIB pTypeLib;

	// 调用系统API来加载TypeLib，该函数会增加typelib的引用计数，获取到一个TypeLib对象
	if (SUCCEEDED(hr = LoadRegTypeLib(&CLSID_TypeLib, 1, 0, 0, &pTypeLib))) {
		// 将IID_IExample，即IExample的vtable信息加载到gTypeInfo中
		if (SUCCEEDED(hr = pTypeLib->lpVtbl->GetTypeInfoOfGuid(pTypeLib, &IID_IExample, &gTypeInfo))) {
			// 我们获取pTypeLib的主要目的就是调用它的GetTypeInfoOfGuid函数，现在这个函数已经完成了他的使命
			// 那么pTypeLib指针也就没用了
			pTypeLib->lpVtbl->Release(pTypeLib);
			// 需要增加gTypeInfo的引用计数
			gTypeInfo->lpVtbl->AddRef(gTypeInfo);
		}
	}

	return hr;
}

// 实现IDispatch的4个函数
HRESULT STDMETHODCALLTYPE GetTypeInfoCount(LPVOID this, UINT* pCount) {
	// 有typelib就填充1，没有就填充0
	*pCount = 1;
	return S_OK;
}
HRESULT STDMETHODCALLTYPE GetTypeInfo(LPVOID this, UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) {
	HRESULT hr;
	*ppTInfo = 0;

	// 这个参数值应该是0，因为我们并不打算使用这个参数
	if (iTInfo) hr = ResultFromScode(DISP_E_BADINDEX);

	if (gTypeInfo) {
		gTypeInfo->lpVtbl->AddRef(gTypeInfo);
		hr = S_OK;
	}
	else {
		// 加载TypeLib
		hr = loadTypeInfo();
	}

	if (SUCCEEDED(hr))
		*ppTInfo = gTypeInfo;

	return hr;
}

HRESULT STDMETHODCALLTYPE GetIDsOfNames(LPVOID this, REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) {
	HRESULT hr;
	if (!gTypeInfo) {
		// 加载TypeLib
		if (FAILED(hr = loadTypeInfo()))
			// 不为0，说明发生了错误
			return hr;
	}
	// 调用系统函数来进行实现
	hr = DispGetIDsOfNames(gTypeInfo, rgszNames, cNames, rgDispId);
	return hr;
}

// Invoke函数会被脚本引擎间接调用，通过传入dispatchID来调用我们在IExample中实现的GetString和SetString函数
HRESULT STDMETHODCALLTYPE Invoke(LPVOID this, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) {
	/*
	HRESULT hr;
	switch (dispIdMember) {
		// GetString
	case 1: {
		hr = S_OK;
		break;
	}
		  // SetString
	case 2: {
		// 对参数类型进行判断
		if (pDispParams->rgvarg->vt != VT_BSTR) {
			// 判断对方传进来的是不是一个DWORD类型，如果是的话，就给他转换成字符串

			// // WORD最长也就32，加上terminating 0，就是33
			// WCHAR temp[33];
			// wsprintfW(temp, L"%d", pDispParams->rgvarg->lVal);
			// // 然后赋值给bstr成员
			// pDispParams->rgvarg->pbstrVal = SysAllocStringLen(temp, SysStringLen(temp));


			// windows提供了类型转换API，所以上面的代码可以用一个API来完成
			if ((hr = VariantChangeType(pDispParams->rgvarg, pDispParams->rgvarg, 0, VT_BSTR))) {
				hr = ((IExamplePtrForDll)this)->lpVtbl->SetString(this, pDispParams->rgvarg->pbstrVal);
			}
		}
		else {
			hr = ResultFromScode(DISP_E_TYPEMISMATCH);
		}
	}

	default:
		break;
	}
	return hr;
	*/

	// 更好的做法是使用OLE32.dll提供的DispInvoke函数来替我们完成工作，就像GetIDsOfNames中的做法一样
	// 我们只需要确保TypeLib已经被正确加载
	if (!IsEqualIID(riid, &IID_NULL)) {
		// 根据微软官方文档 https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-getidsofnames
		// Reserved for future use. Must be IID_NULL.
		return ResultFromScode(DISP_E_UNKNOWNINTERFACE);
	}

	HRESULT hr;
	if (!gTypeInfo) {
		// 加载TypeLib
		if (FAILED(hr = loadTypeInfo()))
			// 不为0，说明发生了错误
			return hr;
	}
	// 调用系统函数来进行实现
	hr = DispInvoke(this, gTypeInfo, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	return hr;
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
	DWORD count = InterlockedDecrement(&OutstandingObjects);
	if (!(InterlockedAnd(&OutstandingObjects, 0xFFFFFFFF) | InterlockedAnd(&LockCount, 0xFFFFFFFF)))
		PostQuitMessage(0);
	return count;
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
		// 初始化成员变量
		example->string = 0;
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


int APIENTRY WinMain(HINSTANCE hInstance,
	HINSTANCE hPrevInstance,
	LPSTR     lpCmdLine,
	int       nCmdShow) {

	// CoRegisterClassObject会增加factory的引用计数，但是要到调用CoRevokeClassObject的时候才会释放
	// 因此我们一开始要把值设置为-1，不然quit message永远都不会发送
	OutstandingObjects=-1;
	LockCount = 0;
	classFactory.lpVtbl = &IClassFactory_Vtbl;

	HRESULT hr;
	MessageBox(NULL, "COM local server is initilizing", "COM server message", MB_OK | MB_SETFOREGROUND);
	CoInitialize(NULL);

	// Let's see if we were started by SCM.
	if (strstr(lpCmdLine, "/Embedding") || strstr(lpCmdLine, "-Embedding"))
	{
		// for debug
		MessageBox(NULL, "COM local server is registering IExample class", "COM server message", MB_OK | MB_SETFOREGROUND);

		// Register the MyCar Factory.
		DWORD regID = 0;
		// 第二个参数一定要进行类型转换，不然加载的时候会报错，妈的坑死我了
		hr = CoRegisterClassObject(&CLSID_IExample, (IUnknown*)&classFactory,
			CLSCTX_SERVER, REGCLS_MULTIPLEUSE, &regID);
		if (FAILED(hr))
		{
			ShowErrorMessage("CoRegisterClassObject()", hr);
			CoUninitialize();
			exit(1);
		}
		MessageBox(NULL, "Begin message loop", "COM server message", MB_OK | MB_SETFOREGROUND);
		MSG ms;
		while (GetMessage(&ms, 0, 0, 0))
		{
			TranslateMessage(&ms);
			DispatchMessage(&ms);
		}

		// All done, so remove class object.
		CoRevokeClassObject(regID);
	}

	// Terminate COM.
	CoUninitialize();
	MessageBox(NULL, "COM local server quit", "COM server message", MB_OK | MB_SETFOREGROUND);
	return 0;
}