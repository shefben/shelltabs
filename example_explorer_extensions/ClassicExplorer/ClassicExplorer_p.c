

/* this ALWAYS GENERATED file contains the proxy stub code */


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

#if defined(_M_AMD64)


#if _MSC_VER >= 1200
#pragma warning(push)
#endif

#pragma warning( disable: 4211 )  /* redefine extern to static */
#pragma warning( disable: 4232 )  /* dllimport identity*/
#pragma warning( disable: 4024 )  /* array to pointer mapping*/
#pragma warning( disable: 4152 )  /* function/data pointer conversion in expression */

#define USE_STUBLESS_PROXY


/* verify that the <rpcproxy.h> version is high enough to compile this file*/
#ifndef __REDQ_RPCPROXY_H_VERSION__
#define __REQUIRED_RPCPROXY_H_VERSION__ 475
#endif


#include "rpcproxy.h"
#include "ndr64types.h"
#ifndef __RPCPROXY_H_VERSION__
#error this stub requires an updated version of <rpcproxy.h>
#endif /* __RPCPROXY_H_VERSION__ */


#include "ClassicExplorer_h.h"

#define TYPE_FORMAT_STRING_SIZE   3                                 
#define PROC_FORMAT_STRING_SIZE   1                                 
#define EXPR_FORMAT_STRING_SIZE   1                                 
#define TRANSMIT_AS_TABLE_SIZE    0            
#define WIRE_MARSHAL_TABLE_SIZE   0            

typedef struct _ClassicExplorer_MIDL_TYPE_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ TYPE_FORMAT_STRING_SIZE ];
    } ClassicExplorer_MIDL_TYPE_FORMAT_STRING;

typedef struct _ClassicExplorer_MIDL_PROC_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ PROC_FORMAT_STRING_SIZE ];
    } ClassicExplorer_MIDL_PROC_FORMAT_STRING;

typedef struct _ClassicExplorer_MIDL_EXPR_FORMAT_STRING
    {
    long          Pad;
    unsigned char  Format[ EXPR_FORMAT_STRING_SIZE ];
    } ClassicExplorer_MIDL_EXPR_FORMAT_STRING;


static const RPC_SYNTAX_IDENTIFIER  _RpcTransferSyntax_2_0 = 
{{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}};

static const RPC_SYNTAX_IDENTIFIER  _NDR64_RpcTransferSyntax_1_0 = 
{{0x71710533,0xbeba,0x4937,{0x83,0x19,0xb5,0xdb,0xef,0x9c,0xcc,0x36}},{1,0}};

#if defined(_CONTROL_FLOW_GUARD_XFG)
#define XFG_TRAMPOLINES(ObjectType)\
NDR_SHAREABLE unsigned long ObjectType ## _UserSize_XFG(unsigned long * pFlags, unsigned long Offset, void * pObject)\
{\
return  ObjectType ## _UserSize(pFlags, Offset, (ObjectType *)pObject);\
}\
NDR_SHAREABLE unsigned char * ObjectType ## _UserMarshal_XFG(unsigned long * pFlags, unsigned char * pBuffer, void * pObject)\
{\
return ObjectType ## _UserMarshal(pFlags, pBuffer, (ObjectType *)pObject);\
}\
NDR_SHAREABLE unsigned char * ObjectType ## _UserUnmarshal_XFG(unsigned long * pFlags, unsigned char * pBuffer, void * pObject)\
{\
return ObjectType ## _UserUnmarshal(pFlags, pBuffer, (ObjectType *)pObject);\
}\
NDR_SHAREABLE void ObjectType ## _UserFree_XFG(unsigned long * pFlags, void * pObject)\
{\
ObjectType ## _UserFree(pFlags, (ObjectType *)pObject);\
}
#define XFG_TRAMPOLINES64(ObjectType)\
NDR_SHAREABLE unsigned long ObjectType ## _UserSize64_XFG(unsigned long * pFlags, unsigned long Offset, void * pObject)\
{\
return  ObjectType ## _UserSize64(pFlags, Offset, (ObjectType *)pObject);\
}\
NDR_SHAREABLE unsigned char * ObjectType ## _UserMarshal64_XFG(unsigned long * pFlags, unsigned char * pBuffer, void * pObject)\
{\
return ObjectType ## _UserMarshal64(pFlags, pBuffer, (ObjectType *)pObject);\
}\
NDR_SHAREABLE unsigned char * ObjectType ## _UserUnmarshal64_XFG(unsigned long * pFlags, unsigned char * pBuffer, void * pObject)\
{\
return ObjectType ## _UserUnmarshal64(pFlags, pBuffer, (ObjectType *)pObject);\
}\
NDR_SHAREABLE void ObjectType ## _UserFree64_XFG(unsigned long * pFlags, void * pObject)\
{\
ObjectType ## _UserFree64(pFlags, (ObjectType *)pObject);\
}
#define XFG_BIND_TRAMPOLINES(HandleType, ObjectType)\
static void* ObjectType ## _bind_XFG(HandleType pObject)\
{\
return ObjectType ## _bind((ObjectType) pObject);\
}\
static void ObjectType ## _unbind_XFG(HandleType pObject, handle_t ServerHandle)\
{\
ObjectType ## _unbind((ObjectType) pObject, ServerHandle);\
}
#define XFG_TRAMPOLINE_FPTR(Function) Function ## _XFG
#define XFG_TRAMPOLINE_FPTR_DEPENDENT_SYMBOL(Symbol) Symbol ## _XFG
#else
#define XFG_TRAMPOLINES(ObjectType)
#define XFG_TRAMPOLINES64(ObjectType)
#define XFG_BIND_TRAMPOLINES(HandleType, ObjectType)
#define XFG_TRAMPOLINE_FPTR(Function) Function
#define XFG_TRAMPOLINE_FPTR_DEPENDENT_SYMBOL(Symbol) Symbol
#endif



extern const ClassicExplorer_MIDL_TYPE_FORMAT_STRING ClassicExplorer__MIDL_TypeFormatString;
extern const ClassicExplorer_MIDL_PROC_FORMAT_STRING ClassicExplorer__MIDL_ProcFormatString;
extern const ClassicExplorer_MIDL_EXPR_FORMAT_STRING ClassicExplorer__MIDL_ExprFormatString;

#ifdef __cplusplus
namespace {
#endif

extern const MIDL_STUB_DESC Object_StubDesc;
#ifdef __cplusplus
}
#endif


extern const MIDL_SERVER_INFO IExplorerBand_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IExplorerBand_ProxyInfo;

#ifdef __cplusplus
namespace {
#endif

extern const MIDL_STUB_DESC Object_StubDesc;
#ifdef __cplusplus
}
#endif


extern const MIDL_SERVER_INFO IClassicCopyExt_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IClassicCopyExt_ProxyInfo;

#ifdef __cplusplus
namespace {
#endif

extern const MIDL_STUB_DESC Object_StubDesc;
#ifdef __cplusplus
}
#endif


extern const MIDL_SERVER_INFO IExplorerBHO_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IExplorerBHO_ProxyInfo;

#ifdef __cplusplus
namespace {
#endif

extern const MIDL_STUB_DESC Object_StubDesc;
#ifdef __cplusplus
}
#endif


extern const MIDL_SERVER_INFO IShareOverlay_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IShareOverlay_ProxyInfo;



#if !defined(__RPC_WIN64__)
#error  Invalid build platform for this stub.
#endif

static const ClassicExplorer_MIDL_PROC_FORMAT_STRING ClassicExplorer__MIDL_ProcFormatString =
    {
        0,
        {

			0x0
        }
    };

static const ClassicExplorer_MIDL_TYPE_FORMAT_STRING ClassicExplorer__MIDL_TypeFormatString =
    {
        0,
        {
			NdrFcShort( 0x0 ),	/* 0 */

			0x0
        }
    };


/* Object interface: IUnknown, ver. 0.0,
   GUID={0x00000000,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IDispatch, ver. 0.0,
   GUID={0x00020400,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IExplorerBand, ver. 0.0,
   GUID={0xBC4C1B8F,0x0BDE,0x4E42,{0x95,0x83,0xE0,0x72,0xB2,0xA2,0x8E,0x0D}} */

#pragma code_seg(".orpc")
static const unsigned short IExplorerBand_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };



/* Object interface: IClassicCopyExt, ver. 0.0,
   GUID={0x6E00B97F,0xA4D4,0x4062,{0x98,0xE4,0x4F,0x66,0xFC,0x96,0xF3,0x2F}} */

#pragma code_seg(".orpc")
static const unsigned short IClassicCopyExt_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };



/* Object interface: IExplorerBHO, ver. 0.0,
   GUID={0xA1678625,0xA011,0x4B7C,{0xA1,0xFA,0xD6,0x91,0xE4,0xCD,0xDB,0x79}} */

#pragma code_seg(".orpc")
static const unsigned short IExplorerBHO_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };



/* Object interface: IShareOverlay, ver. 0.0,
   GUID={0x2576496C,0xB58A,0x4995,{0x88,0x78,0x8B,0x68,0xF9,0xE8,0xD1,0xFC}} */

#pragma code_seg(".orpc")
static const unsigned short IShareOverlay_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };



#endif /* defined(_M_AMD64)*/



/* this ALWAYS GENERATED file contains the proxy stub code */


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

#if defined(_M_AMD64)




#if !defined(__RPC_WIN64__)
#error  Invalid build platform for this stub.
#endif


#include "ndr64types.h"
#include "pshpack8.h"
#ifdef __cplusplus
namespace {
#endif


typedef 
NDR64_FORMAT_UINT32
__midl_frag1_t;
extern const __midl_frag1_t __midl_frag1;

static const __midl_frag1_t __midl_frag1 =
(NDR64_UINT32) 0 /* 0x0 */;
#ifdef __cplusplus
}
#endif


#include "poppack.h"



/* Object interface: IUnknown, ver. 0.0,
   GUID={0x00000000,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IDispatch, ver. 0.0,
   GUID={0x00020400,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IExplorerBand, ver. 0.0,
   GUID={0xBC4C1B8F,0x0BDE,0x4E42,{0x95,0x83,0xE0,0x72,0xB2,0xA2,0x8E,0x0D}} */

#pragma code_seg(".orpc")
static const FormatInfoRef IExplorerBand_Ndr64ProcTable[] =
    {
    (FormatInfoRef)(LONG_PTR) -1,
    (FormatInfoRef)(LONG_PTR) -1,
    (FormatInfoRef)(LONG_PTR) -1,
    (FormatInfoRef)(LONG_PTR) -1,
    0
    };


static const MIDL_SYNTAX_INFO IExplorerBand_SyntaxInfo [  2 ] = 
    {
    {
    {{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}},
    0,
    ClassicExplorer__MIDL_ProcFormatString.Format,
    &IExplorerBand_FormatStringOffsetTable[-3],
    ClassicExplorer__MIDL_TypeFormatString.Format,
    0,
    0,
    0
    }
    ,{
    {{0x71710533,0xbeba,0x4937,{0x83,0x19,0xb5,0xdb,0xef,0x9c,0xcc,0x36}},{1,0}},
    0,
    0 ,
    (unsigned short *) &IExplorerBand_Ndr64ProcTable[-3],
    0,
    0,
    0,
    0
    }
    };

static const MIDL_STUBLESS_PROXY_INFO IExplorerBand_ProxyInfo =
    {
    &Object_StubDesc,
    ClassicExplorer__MIDL_ProcFormatString.Format,
    &IExplorerBand_FormatStringOffsetTable[-3],
    (RPC_SYNTAX_IDENTIFIER*)&_RpcTransferSyntax_2_0,
    2,
    (MIDL_SYNTAX_INFO*)IExplorerBand_SyntaxInfo
    
    };


static const MIDL_SERVER_INFO IExplorerBand_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    ClassicExplorer__MIDL_ProcFormatString.Format,
    (unsigned short *) &IExplorerBand_FormatStringOffsetTable[-3],
    0,
    (RPC_SYNTAX_IDENTIFIER*)&_NDR64_RpcTransferSyntax_1_0,
    2,
    (MIDL_SYNTAX_INFO*)IExplorerBand_SyntaxInfo
    };
CINTERFACE_PROXY_VTABLE(7) _IExplorerBandProxyVtbl = 
{
    0,
    &IID_IExplorerBand,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


EXTERN_C DECLSPEC_SELECTANY const PRPC_STUB_FUNCTION IExplorerBand_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _IExplorerBandStubVtbl =
{
    &IID_IExplorerBand,
    &IExplorerBand_ServerInfo,
    7,
    &IExplorerBand_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};


/* Object interface: IClassicCopyExt, ver. 0.0,
   GUID={0x6E00B97F,0xA4D4,0x4062,{0x98,0xE4,0x4F,0x66,0xFC,0x96,0xF3,0x2F}} */

#pragma code_seg(".orpc")
static const FormatInfoRef IClassicCopyExt_Ndr64ProcTable[] =
    {
    (FormatInfoRef)(LONG_PTR) -1,
    (FormatInfoRef)(LONG_PTR) -1,
    (FormatInfoRef)(LONG_PTR) -1,
    (FormatInfoRef)(LONG_PTR) -1,
    0
    };


static const MIDL_SYNTAX_INFO IClassicCopyExt_SyntaxInfo [  2 ] = 
    {
    {
    {{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}},
    0,
    ClassicExplorer__MIDL_ProcFormatString.Format,
    &IClassicCopyExt_FormatStringOffsetTable[-3],
    ClassicExplorer__MIDL_TypeFormatString.Format,
    0,
    0,
    0
    }
    ,{
    {{0x71710533,0xbeba,0x4937,{0x83,0x19,0xb5,0xdb,0xef,0x9c,0xcc,0x36}},{1,0}},
    0,
    0 ,
    (unsigned short *) &IClassicCopyExt_Ndr64ProcTable[-3],
    0,
    0,
    0,
    0
    }
    };

static const MIDL_STUBLESS_PROXY_INFO IClassicCopyExt_ProxyInfo =
    {
    &Object_StubDesc,
    ClassicExplorer__MIDL_ProcFormatString.Format,
    &IClassicCopyExt_FormatStringOffsetTable[-3],
    (RPC_SYNTAX_IDENTIFIER*)&_RpcTransferSyntax_2_0,
    2,
    (MIDL_SYNTAX_INFO*)IClassicCopyExt_SyntaxInfo
    
    };


static const MIDL_SERVER_INFO IClassicCopyExt_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    ClassicExplorer__MIDL_ProcFormatString.Format,
    (unsigned short *) &IClassicCopyExt_FormatStringOffsetTable[-3],
    0,
    (RPC_SYNTAX_IDENTIFIER*)&_NDR64_RpcTransferSyntax_1_0,
    2,
    (MIDL_SYNTAX_INFO*)IClassicCopyExt_SyntaxInfo
    };
CINTERFACE_PROXY_VTABLE(7) _IClassicCopyExtProxyVtbl = 
{
    0,
    &IID_IClassicCopyExt,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


EXTERN_C DECLSPEC_SELECTANY const PRPC_STUB_FUNCTION IClassicCopyExt_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _IClassicCopyExtStubVtbl =
{
    &IID_IClassicCopyExt,
    &IClassicCopyExt_ServerInfo,
    7,
    &IClassicCopyExt_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};


/* Object interface: IExplorerBHO, ver. 0.0,
   GUID={0xA1678625,0xA011,0x4B7C,{0xA1,0xFA,0xD6,0x91,0xE4,0xCD,0xDB,0x79}} */

#pragma code_seg(".orpc")
static const FormatInfoRef IExplorerBHO_Ndr64ProcTable[] =
    {
    (FormatInfoRef)(LONG_PTR) -1,
    (FormatInfoRef)(LONG_PTR) -1,
    (FormatInfoRef)(LONG_PTR) -1,
    (FormatInfoRef)(LONG_PTR) -1,
    0
    };


static const MIDL_SYNTAX_INFO IExplorerBHO_SyntaxInfo [  2 ] = 
    {
    {
    {{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}},
    0,
    ClassicExplorer__MIDL_ProcFormatString.Format,
    &IExplorerBHO_FormatStringOffsetTable[-3],
    ClassicExplorer__MIDL_TypeFormatString.Format,
    0,
    0,
    0
    }
    ,{
    {{0x71710533,0xbeba,0x4937,{0x83,0x19,0xb5,0xdb,0xef,0x9c,0xcc,0x36}},{1,0}},
    0,
    0 ,
    (unsigned short *) &IExplorerBHO_Ndr64ProcTable[-3],
    0,
    0,
    0,
    0
    }
    };

static const MIDL_STUBLESS_PROXY_INFO IExplorerBHO_ProxyInfo =
    {
    &Object_StubDesc,
    ClassicExplorer__MIDL_ProcFormatString.Format,
    &IExplorerBHO_FormatStringOffsetTable[-3],
    (RPC_SYNTAX_IDENTIFIER*)&_RpcTransferSyntax_2_0,
    2,
    (MIDL_SYNTAX_INFO*)IExplorerBHO_SyntaxInfo
    
    };


static const MIDL_SERVER_INFO IExplorerBHO_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    ClassicExplorer__MIDL_ProcFormatString.Format,
    (unsigned short *) &IExplorerBHO_FormatStringOffsetTable[-3],
    0,
    (RPC_SYNTAX_IDENTIFIER*)&_NDR64_RpcTransferSyntax_1_0,
    2,
    (MIDL_SYNTAX_INFO*)IExplorerBHO_SyntaxInfo
    };
CINTERFACE_PROXY_VTABLE(7) _IExplorerBHOProxyVtbl = 
{
    0,
    &IID_IExplorerBHO,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


EXTERN_C DECLSPEC_SELECTANY const PRPC_STUB_FUNCTION IExplorerBHO_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _IExplorerBHOStubVtbl =
{
    &IID_IExplorerBHO,
    &IExplorerBHO_ServerInfo,
    7,
    &IExplorerBHO_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};


/* Object interface: IShareOverlay, ver. 0.0,
   GUID={0x2576496C,0xB58A,0x4995,{0x88,0x78,0x8B,0x68,0xF9,0xE8,0xD1,0xFC}} */

#pragma code_seg(".orpc")
static const FormatInfoRef IShareOverlay_Ndr64ProcTable[] =
    {
    (FormatInfoRef)(LONG_PTR) -1,
    (FormatInfoRef)(LONG_PTR) -1,
    (FormatInfoRef)(LONG_PTR) -1,
    (FormatInfoRef)(LONG_PTR) -1,
    0
    };


static const MIDL_SYNTAX_INFO IShareOverlay_SyntaxInfo [  2 ] = 
    {
    {
    {{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}},
    0,
    ClassicExplorer__MIDL_ProcFormatString.Format,
    &IShareOverlay_FormatStringOffsetTable[-3],
    ClassicExplorer__MIDL_TypeFormatString.Format,
    0,
    0,
    0
    }
    ,{
    {{0x71710533,0xbeba,0x4937,{0x83,0x19,0xb5,0xdb,0xef,0x9c,0xcc,0x36}},{1,0}},
    0,
    0 ,
    (unsigned short *) &IShareOverlay_Ndr64ProcTable[-3],
    0,
    0,
    0,
    0
    }
    };

static const MIDL_STUBLESS_PROXY_INFO IShareOverlay_ProxyInfo =
    {
    &Object_StubDesc,
    ClassicExplorer__MIDL_ProcFormatString.Format,
    &IShareOverlay_FormatStringOffsetTable[-3],
    (RPC_SYNTAX_IDENTIFIER*)&_RpcTransferSyntax_2_0,
    2,
    (MIDL_SYNTAX_INFO*)IShareOverlay_SyntaxInfo
    
    };


static const MIDL_SERVER_INFO IShareOverlay_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    ClassicExplorer__MIDL_ProcFormatString.Format,
    (unsigned short *) &IShareOverlay_FormatStringOffsetTable[-3],
    0,
    (RPC_SYNTAX_IDENTIFIER*)&_NDR64_RpcTransferSyntax_1_0,
    2,
    (MIDL_SYNTAX_INFO*)IShareOverlay_SyntaxInfo
    };
CINTERFACE_PROXY_VTABLE(7) _IShareOverlayProxyVtbl = 
{
    0,
    &IID_IShareOverlay,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


EXTERN_C DECLSPEC_SELECTANY const PRPC_STUB_FUNCTION IShareOverlay_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _IShareOverlayStubVtbl =
{
    &IID_IShareOverlay,
    &IShareOverlay_ServerInfo,
    7,
    &IShareOverlay_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};

#ifdef __cplusplus
namespace {
#endif
static const MIDL_STUB_DESC Object_StubDesc = 
    {
    0,
    NdrOleAllocate,
    NdrOleFree,
    0,
    0,
    0,
    0,
    0,
    ClassicExplorer__MIDL_TypeFormatString.Format,
    1, /* -error bounds_check flag */
    0x60001, /* Ndr library version */
    0,
    0x8010274, /* MIDL Version 8.1.628 */
    0,
    0,
    0,  /* notify & notify_flag routine table */
    0x2000001, /* MIDL flag */
    0, /* cs routines */
    0,   /* proxy/server info */
    0
    };
#ifdef __cplusplus
}
#endif

const CInterfaceProxyVtbl * const _ClassicExplorer_ProxyVtblList[] = 
{
    ( CInterfaceProxyVtbl *) &_IExplorerBHOProxyVtbl,
    ( CInterfaceProxyVtbl *) &_IShareOverlayProxyVtbl,
    ( CInterfaceProxyVtbl *) &_IClassicCopyExtProxyVtbl,
    ( CInterfaceProxyVtbl *) &_IExplorerBandProxyVtbl,
    0
};

const CInterfaceStubVtbl * const _ClassicExplorer_StubVtblList[] = 
{
    ( CInterfaceStubVtbl *) &_IExplorerBHOStubVtbl,
    ( CInterfaceStubVtbl *) &_IShareOverlayStubVtbl,
    ( CInterfaceStubVtbl *) &_IClassicCopyExtStubVtbl,
    ( CInterfaceStubVtbl *) &_IExplorerBandStubVtbl,
    0
};

PCInterfaceName const _ClassicExplorer_InterfaceNamesList[] = 
{
    "IExplorerBHO",
    "IShareOverlay",
    "IClassicCopyExt",
    "IExplorerBand",
    0
};

const IID *  const _ClassicExplorer_BaseIIDList[] = 
{
    &IID_IDispatch,
    &IID_IDispatch,
    &IID_IDispatch,
    &IID_IDispatch,
    0
};


#define _ClassicExplorer_CHECK_IID(n)	IID_GENERIC_CHECK_IID( _ClassicExplorer, pIID, n)

int __stdcall _ClassicExplorer_IID_Lookup( const IID * pIID, int * pIndex )
{
    IID_BS_LOOKUP_SETUP

    IID_BS_LOOKUP_INITIAL_TEST( _ClassicExplorer, 4, 2 )
    IID_BS_LOOKUP_NEXT_TEST( _ClassicExplorer, 1 )
    IID_BS_LOOKUP_RETURN_RESULT( _ClassicExplorer, 4, *pIndex )
    
}

EXTERN_C const ExtendedProxyFileInfo ClassicExplorer_ProxyFileInfo = 
{
    (PCInterfaceProxyVtblList *) & _ClassicExplorer_ProxyVtblList,
    (PCInterfaceStubVtblList *) & _ClassicExplorer_StubVtblList,
    (const PCInterfaceName * ) & _ClassicExplorer_InterfaceNamesList,
    (const IID ** ) & _ClassicExplorer_BaseIIDList,
    & _ClassicExplorer_IID_Lookup, 
    4,
    2,
    0, /* table of [async_uuid] interfaces */
    0, /* Filler1 */
    0, /* Filler2 */
    0  /* Filler3 */
};
#if _MSC_VER >= 1200
#pragma warning(pop)
#endif


#endif /* defined(_M_AMD64)*/

