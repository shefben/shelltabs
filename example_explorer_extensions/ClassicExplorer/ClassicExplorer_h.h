

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0628 */
/* at Mon Jan 18 21:14:07 2038
 */
/* Compiler settings for ClassicExplorer.idl:
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

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __ClassicExplorer_h_h__
#define __ClassicExplorer_h_h__

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

#ifndef __IExplorerBand_FWD_DEFINED__
#define __IExplorerBand_FWD_DEFINED__
typedef interface IExplorerBand IExplorerBand;

#endif 	/* __IExplorerBand_FWD_DEFINED__ */


#ifndef __IClassicCopyExt_FWD_DEFINED__
#define __IClassicCopyExt_FWD_DEFINED__
typedef interface IClassicCopyExt IClassicCopyExt;

#endif 	/* __IClassicCopyExt_FWD_DEFINED__ */


#ifndef __IExplorerBHO_FWD_DEFINED__
#define __IExplorerBHO_FWD_DEFINED__
typedef interface IExplorerBHO IExplorerBHO;

#endif 	/* __IExplorerBHO_FWD_DEFINED__ */


#ifndef __IShareOverlay_FWD_DEFINED__
#define __IShareOverlay_FWD_DEFINED__
typedef interface IShareOverlay IShareOverlay;

#endif 	/* __IShareOverlay_FWD_DEFINED__ */


#ifndef __ExplorerBand_FWD_DEFINED__
#define __ExplorerBand_FWD_DEFINED__

#ifdef __cplusplus
typedef class ExplorerBand ExplorerBand;
#else
typedef struct ExplorerBand ExplorerBand;
#endif /* __cplusplus */

#endif 	/* __ExplorerBand_FWD_DEFINED__ */


#ifndef __ClassicCopyExt_FWD_DEFINED__
#define __ClassicCopyExt_FWD_DEFINED__

#ifdef __cplusplus
typedef class ClassicCopyExt ClassicCopyExt;
#else
typedef struct ClassicCopyExt ClassicCopyExt;
#endif /* __cplusplus */

#endif 	/* __ClassicCopyExt_FWD_DEFINED__ */


#ifndef __ExplorerBHO_FWD_DEFINED__
#define __ExplorerBHO_FWD_DEFINED__

#ifdef __cplusplus
typedef class ExplorerBHO ExplorerBHO;
#else
typedef struct ExplorerBHO ExplorerBHO;
#endif /* __cplusplus */

#endif 	/* __ExplorerBHO_FWD_DEFINED__ */


#ifndef __ShareOverlay_FWD_DEFINED__
#define __ShareOverlay_FWD_DEFINED__

#ifdef __cplusplus
typedef class ShareOverlay ShareOverlay;
#else
typedef struct ShareOverlay ShareOverlay;
#endif /* __cplusplus */

#endif 	/* __ShareOverlay_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IExplorerBand_INTERFACE_DEFINED__
#define __IExplorerBand_INTERFACE_DEFINED__

/* interface IExplorerBand */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IExplorerBand;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("BC4C1B8F-0BDE-4E42-9583-E072B2A28E0D")
    IExplorerBand : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IExplorerBandVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IExplorerBand * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IExplorerBand * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IExplorerBand * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IExplorerBand * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IExplorerBand * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IExplorerBand * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IExplorerBand * This,
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
        
        END_INTERFACE
    } IExplorerBandVtbl;

    interface IExplorerBand
    {
        CONST_VTBL struct IExplorerBandVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IExplorerBand_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IExplorerBand_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IExplorerBand_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IExplorerBand_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IExplorerBand_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IExplorerBand_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IExplorerBand_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IExplorerBand_INTERFACE_DEFINED__ */


#ifndef __IClassicCopyExt_INTERFACE_DEFINED__
#define __IClassicCopyExt_INTERFACE_DEFINED__

/* interface IClassicCopyExt */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IClassicCopyExt;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("6E00B97F-A4D4-4062-98E4-4F66FC96F32F")
    IClassicCopyExt : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IClassicCopyExtVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IClassicCopyExt * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IClassicCopyExt * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IClassicCopyExt * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IClassicCopyExt * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IClassicCopyExt * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IClassicCopyExt * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IClassicCopyExt * This,
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
        
        END_INTERFACE
    } IClassicCopyExtVtbl;

    interface IClassicCopyExt
    {
        CONST_VTBL struct IClassicCopyExtVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IClassicCopyExt_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IClassicCopyExt_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IClassicCopyExt_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IClassicCopyExt_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IClassicCopyExt_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IClassicCopyExt_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IClassicCopyExt_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IClassicCopyExt_INTERFACE_DEFINED__ */


#ifndef __IExplorerBHO_INTERFACE_DEFINED__
#define __IExplorerBHO_INTERFACE_DEFINED__

/* interface IExplorerBHO */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IExplorerBHO;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("A1678625-A011-4B7C-A1FA-D691E4CDDB79")
    IExplorerBHO : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IExplorerBHOVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IExplorerBHO * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IExplorerBHO * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IExplorerBHO * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IExplorerBHO * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IExplorerBHO * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IExplorerBHO * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IExplorerBHO * This,
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
        
        END_INTERFACE
    } IExplorerBHOVtbl;

    interface IExplorerBHO
    {
        CONST_VTBL struct IExplorerBHOVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IExplorerBHO_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IExplorerBHO_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IExplorerBHO_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IExplorerBHO_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IExplorerBHO_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IExplorerBHO_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IExplorerBHO_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IExplorerBHO_INTERFACE_DEFINED__ */


#ifndef __IShareOverlay_INTERFACE_DEFINED__
#define __IShareOverlay_INTERFACE_DEFINED__

/* interface IShareOverlay */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IShareOverlay;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("2576496C-B58A-4995-8878-8B68F9E8D1FC")
    IShareOverlay : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IShareOverlayVtbl
    {
        BEGIN_INTERFACE
        
        DECLSPEC_XFGVIRT(IUnknown, QueryInterface)
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IShareOverlay * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        DECLSPEC_XFGVIRT(IUnknown, AddRef)
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IShareOverlay * This);
        
        DECLSPEC_XFGVIRT(IUnknown, Release)
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IShareOverlay * This);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfoCount)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IShareOverlay * This,
            /* [out] */ UINT *pctinfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetTypeInfo)
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IShareOverlay * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        DECLSPEC_XFGVIRT(IDispatch, GetIDsOfNames)
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IShareOverlay * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        DECLSPEC_XFGVIRT(IDispatch, Invoke)
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IShareOverlay * This,
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
        
        END_INTERFACE
    } IShareOverlayVtbl;

    interface IShareOverlay
    {
        CONST_VTBL struct IShareOverlayVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IShareOverlay_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IShareOverlay_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IShareOverlay_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IShareOverlay_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IShareOverlay_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IShareOverlay_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IShareOverlay_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IShareOverlay_INTERFACE_DEFINED__ */



#ifndef __ClassicExplorerLib_LIBRARY_DEFINED__
#define __ClassicExplorerLib_LIBRARY_DEFINED__

/* library ClassicExplorerLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_ClassicExplorerLib;

EXTERN_C const CLSID CLSID_ExplorerBand;

#ifdef __cplusplus

class DECLSPEC_UUID("553891B7-A0D5-4526-BE18-D3CE461D6310")
ExplorerBand;
#endif

EXTERN_C const CLSID CLSID_ClassicCopyExt;

#ifdef __cplusplus

class DECLSPEC_UUID("8C83ACB1-75C3-45D2-882C-EFA32333491C")
ClassicCopyExt;
#endif

EXTERN_C const CLSID CLSID_ExplorerBHO;

#ifdef __cplusplus

class DECLSPEC_UUID("449D0D6E-2412-4E61-B68F-1CB625CD9E52")
ExplorerBHO;
#endif

EXTERN_C const CLSID CLSID_ShareOverlay;

#ifdef __cplusplus

class DECLSPEC_UUID("594D4122-1F87-41E2-96C7-825FB4796516")
ShareOverlay;
#endif
#endif /* __ClassicExplorerLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


