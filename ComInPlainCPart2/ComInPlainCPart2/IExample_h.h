

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 11:14:07 2038
 */
/* Compiler settings for IExample.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.01.0628 
    protocol : all , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */



/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 500
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */


#ifndef __IExample_h_h__
#define __IExample_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef DECLSPEC_XFGVIRT
#if defined(_CONTROL_FLOW_GUARD_XFG)
#define DECLSPEC_XFGVIRT(base, func) __declspec(xfg_virtual(base, func))
#else
#define DECLSPEC_XFGVIRT(base, func)
#endif
#endif

/* Forward Declarations */ 

#ifndef __IExampleVtbl_FWD_DEFINED__
#define __IExampleVtbl_FWD_DEFINED__
typedef interface IExampleVtbl IExampleVtbl;

#endif 	/* __IExampleVtbl_FWD_DEFINED__ */


#ifndef __IExmaple_FWD_DEFINED__
#define __IExmaple_FWD_DEFINED__

#ifdef __cplusplus
typedef class IExmaple IExmaple;
#else
typedef struct IExmaple IExmaple;
#endif /* __cplusplus */

#endif 	/* __IExmaple_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 



#ifndef __IExample_LIBRARY_DEFINED__
#define __IExample_LIBRARY_DEFINED__

/* library IExample */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_IExample;

#ifndef __IExampleVtbl_INTERFACE_DEFINED__
#define __IExampleVtbl_INTERFACE_DEFINED__

/* interface IExampleVtbl */
/* [object][nonextensible][hidden][oleautomation][dual][uuid] */ 


EXTERN_C const IID IID_IExampleVtbl;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("75ADDF86-4362-45FC-A990-4335C7E45FAA")
    IExampleVtbl : public IDispatch
    {
    public:
        virtual /* [propput][id][helpstring] */ HRESULT STDMETHODCALLTYPE put_string( 
            /* [in] */ BSTR __MIDL__IExampleVtbl0000) = 0;
        
        virtual /* [propget][id][helpstring] */ HRESULT STDMETHODCALLTYPE get_string( 
            /* [retval][out] */ BSTR *__MIDL__IExampleVtbl0001) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IExampleVtblVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IExampleVtbl * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IExampleVtbl * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IExampleVtbl * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IExampleVtbl * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IExampleVtbl * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IExampleVtbl * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IExampleVtbl * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        DECLSPEC_XFGVIRT(IExampleVtbl, put_string)
        /* [propput][id][helpstring] */ HRESULT ( STDMETHODCALLTYPE *put_string )( 
            IExampleVtbl * This,
            /* [in] */ BSTR __MIDL__IExampleVtbl0000);
        
        DECLSPEC_XFGVIRT(IExampleVtbl, get_string)
        /* [propget][id][helpstring] */ HRESULT ( STDMETHODCALLTYPE *get_string )( 
            IExampleVtbl * This,
            /* [retval][out] */ BSTR *__MIDL__IExampleVtbl0001);
        
        END_INTERFACE
    } IExampleVtblVtbl;

    interface IExampleVtbl
    {
        CONST_VTBL struct IExampleVtblVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IExampleVtbl_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IExampleVtbl_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IExampleVtbl_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IExampleVtbl_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IExampleVtbl_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IExampleVtbl_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IExampleVtbl_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IExampleVtbl_put_string(This,__MIDL__IExampleVtbl0000)	\
    ( (This)->lpVtbl -> put_string(This,__MIDL__IExampleVtbl0000) ) 

#define IExampleVtbl_get_string(This,__MIDL__IExampleVtbl0001)	\
    ( (This)->lpVtbl -> get_string(This,__MIDL__IExampleVtbl0001) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IExampleVtbl_INTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_IExmaple;

#ifdef __cplusplus

class DECLSPEC_UUID("912889EE-309E-4CD4-90C3-44BC20698575")
IExmaple;
#endif
#endif /* __IExample_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


