

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Tue Jan 19 06:14:07 2038
 */
/* Compiler settings for C:/Users/Станислав/Documents/Visual Studio Projects/NetComServer/ComInprocServer/source/ComServer.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=IA64 8.01.0628 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif /* __RPCNDR_H_VERSION__ */

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __NetComServer_h__
#define __NetComServer_h__

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

#ifndef __IComServer_FWD_DEFINED__
#define __IComServer_FWD_DEFINED__
typedef interface IComServer IComServer;

#endif 	/* __IComServer_FWD_DEFINED__ */


#ifndef __ComServer_FWD_DEFINED__
#define __ComServer_FWD_DEFINED__

#ifdef __cplusplus
typedef class ComServer ComServer;
#else
typedef struct ComServer ComServer;
#endif /* __cplusplus */

#endif 	/* __ComServer_FWD_DEFINED__ */


#ifndef __ComServerClassFactory_FWD_DEFINED__
#define __ComServerClassFactory_FWD_DEFINED__

#ifdef __cplusplus
typedef class ComServerClassFactory ComServerClassFactory;
#else
typedef struct ComServerClassFactory ComServerClassFactory;
#endif /* __cplusplus */

#endif 	/* __ComServerClassFactory_FWD_DEFINED__ */


#ifndef __ComServerConstants_FWD_DEFINED__
#define __ComServerConstants_FWD_DEFINED__

#ifdef __cplusplus
typedef class ComServerConstants ComServerConstants;
#else
typedef struct ComServerConstants ComServerConstants;
#endif /* __cplusplus */

#endif 	/* __ComServerConstants_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IComServer_INTERFACE_DEFINED__
#define __IComServer_INTERFACE_DEFINED__

/* interface IComServer */
/* [uuid][oleautomation][object] */ 


EXTERN_C const IID IID_IComServer;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("E62A1CB0-86A7-40AE-AFE4-75562C32A498")
    IComServer : public IDispatch
    {
    public:
        virtual HRESULT STDMETHODCALLTYPE LoadAssembly( 
            /* [in] */ const wchar_t *assemblyName) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE GetCurrentAssembly( 
            /* [retval][out] */ IDispatch **ret) = 0;
        
        virtual HRESULT STDMETHODCALLTYPE CreateNetInstance( 
            /* [in] */ const wchar_t *typeName,
            /* [retval][out] */ IDispatch **ret) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IComServerVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IComServer * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IComServer * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IComServer * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IComServer * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IComServer * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IComServer * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IComServer * This,
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
        
        DECLSPEC_XFGVIRT(IComServer, LoadAssembly)
        HRESULT ( STDMETHODCALLTYPE *LoadAssembly )( 
            IComServer * This,
            /* [in] */ const wchar_t *assemblyName);
        
        DECLSPEC_XFGVIRT(IComServer, GetCurrentAssembly)
        HRESULT ( STDMETHODCALLTYPE *GetCurrentAssembly )( 
            IComServer * This,
            /* [retval][out] */ IDispatch **ret);
        
        DECLSPEC_XFGVIRT(IComServer, CreateNetInstance)
        HRESULT ( STDMETHODCALLTYPE *CreateNetInstance )( 
            IComServer * This,
            /* [in] */ const wchar_t *typeName,
            /* [retval][out] */ IDispatch **ret);
        
        END_INTERFACE
    } IComServerVtbl;

    interface IComServer
    {
        CONST_VTBL struct IComServerVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IComServer_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IComServer_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IComServer_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IComServer_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IComServer_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IComServer_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IComServer_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IComServer_LoadAssembly(This,assemblyName)	\
    ( (This)->lpVtbl -> LoadAssembly(This,assemblyName) ) 

#define IComServer_GetCurrentAssembly(This,ret)	\
    ( (This)->lpVtbl -> GetCurrentAssembly(This,ret) ) 

#define IComServer_CreateNetInstance(This,typeName,ret)	\
    ( (This)->lpVtbl -> CreateNetInstance(This,typeName,ret) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IComServer_INTERFACE_DEFINED__ */



#ifndef __NetComServer_LIBRARY_DEFINED__
#define __NetComServer_LIBRARY_DEFINED__

/* library NetComServer */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_NetComServer;

EXTERN_C const CLSID CLSID_ComServer;

#ifdef __cplusplus

class DECLSPEC_UUID("C6819A04-046E-4EA2-9750-949760CB26F9")
ComServer;
#endif

EXTERN_C const CLSID CLSID_ComServerClassFactory;

#ifdef __cplusplus

class DECLSPEC_UUID("2A8CABCA-0EBD-40F3-96A4-04054B3ED79D")
ComServerClassFactory;
#endif

EXTERN_C const CLSID CLSID_ComServerConstants;

#ifdef __cplusplus

class DECLSPEC_UUID("9812DE38-ACE5-426F-9C54-652322461AFC")
ComServerConstants;
#endif
#endif /* __NetComServer_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


