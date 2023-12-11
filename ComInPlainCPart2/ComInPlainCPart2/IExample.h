#pragma once
#ifndef _IEXAMPLE_H_
#define _IEXAMPLE_H_

#include<INITGUID.h>
// {912889EE-309E-4CD4-90C3-44BC20698575}
DEFINE_GUID(CLSID_IExample,
	0x912889ee, 0x309e, 0x4cd4, 0x90, 0xc3, 0x44, 0xbc, 0x20, 0x69, 0x85, 0x75);


// {75ADDF86-4362-45FC-A990-4335C7E45FAA}
DEFINE_GUID(IID_IExample,
	0x75addf86, 0x4362, 0x45fc, 0xa9, 0x90, 0x43, 0x35, 0xc7, 0xe4, 0x5f, 0xaa);

// type lib
// {21189F08-3F6A-43D3-9E25-69F516AD66FC}
DEFINE_GUID(CLSID_TypeLib,
	0x21189f08, 0x3f6a, 0x43d3, 0x9e, 0x25, 0x69, 0xf5, 0x16, 0xad, 0x66, 0xfc);


// 我们需要向客户端隐藏成员变量，并声明出一种C++/C都能用的定义，微软提供了相应的宏
#undef INTERFACE
#define INTERFACE	IExample

// 脚本引擎通过IDispatch接口的这几个函数来获取type lib
// type lib就是用来保存我们在这个头文件中声明的这一堆乱七八糟的定义的
DECLARE_INTERFACE_(INTERFACE, IDispatch) {
	STDMETHOD(QueryInterface) (THIS_ REFIID, LPVOID*) PURE;
	// 这两个需要加下划线是因为返回值不是HRESULT，THIS没有下划线是因为他俩只有这一个参数
	STDMETHOD_(ULONG, AddRef) (THIS) PURE;
	STDMETHOD_(ULONG, Release) (THIS) PURE;

	// IDispatch的四个函数
	STDMETHOD(GetTypeInfoCount) (THIS_ UINT*) PURE;
	// 我们手动实现了GetTypeInfoCount函数，但是下面的GetTypeInfo并不需要我们自己去实现
	// GetTypeInfo函数的主要作用是给脚本引擎返回一个ITypeInfo指针,脚本引擎在拿到ITypeInfo指针之后
	// 即可对齐进行解析，通过调用ITypeInfo的GetFuncDesc函数来获取到函数签名
	// windows已经实现了通用的ITypeInfo函数
	// LCID参数就是language id，用来标识描述信息使用的人类语言
	// 我们只需要一个ITypeInfo对象，因为他的内容并不会改变，我们只要负责把TypeLib加载上去就行
	STDMETHOD(GetTypeInfo) (THIS_ UINT, LCID, ITypeInfo**) PURE;
	STDMETHOD(GetIDsOfNames) (THIS_ REFIID, LPOLESTR*, UINT, LCID, DISPID*) PURE;
	STDMETHOD(Invoke) (THIS_ DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) PURE;

	STDMETHOD(SetString) (THIS_ BSTR) PURE;
	STDMETHOD(GetString) (THIS_ BSTR*) PURE;
};
#endif