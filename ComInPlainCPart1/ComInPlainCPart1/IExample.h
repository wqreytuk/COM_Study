#pragma once
#ifndef _IEXAMPLE_H_
#define _IEXAMPLE_H_

#include<INITGUID.h>
// {F99BFE04-0816-40AA-A833-09454DE4617D}
DEFINE_GUID(CLSID_IExample,
	0xf99bfe04, 0x816, 0x40aa, 0xa8, 0x33, 0x9, 0x45, 0x4d, 0xe4, 0x61, 0x7d);

// {8E1791F0-237D-4685-953C-CA146792058E}
DEFINE_GUID(IID_IExample,
	0x8e1791f0, 0x237d, 0x4685, 0x95, 0x3c, 0xca, 0x14, 0x67, 0x92, 0x5, 0x8e);

// 我们需要向客户端隐藏成员变量，并声明出一种C++/C都能用的定义，微软提供了相应的宏
#undef INTERFACE
#define INTERFACE	IExample

DECLARE_INTERFACE_(INTERFACE, IUnknown) {
	STDMETHOD(QueryInterface) (THIS_ REFIID, LPVOID*) PURE;
	// 这两个需要加下划线是因为返回值不是HRESULT，THIS没有下划线是因为他俩只有这一个参数
	STDMETHOD_(ULONG, AddRef) (THIS) PURE;
	STDMETHOD_(ULONG, Release) (THIS) PURE;
	STDMETHOD(SetString) (THIS_ char*) PURE;
	STDMETHOD(GetString) (THIS_ char*, size_t) PURE;
};
#endif