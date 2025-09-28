

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


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



#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        EXTERN_C __declspec(selectany) const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif // !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, IID_IExplorerBand,0xBC4C1B8F,0x0BDE,0x4E42,0x95,0x83,0xE0,0x72,0xB2,0xA2,0x8E,0x0D);


MIDL_DEFINE_GUID(IID, IID_IClassicCopyExt,0x6E00B97F,0xA4D4,0x4062,0x98,0xE4,0x4F,0x66,0xFC,0x96,0xF3,0x2F);


MIDL_DEFINE_GUID(IID, IID_IExplorerBHO,0xA1678625,0xA011,0x4B7C,0xA1,0xFA,0xD6,0x91,0xE4,0xCD,0xDB,0x79);


MIDL_DEFINE_GUID(IID, IID_IShareOverlay,0x2576496C,0xB58A,0x4995,0x88,0x78,0x8B,0x68,0xF9,0xE8,0xD1,0xFC);


MIDL_DEFINE_GUID(IID, LIBID_ClassicExplorerLib,0xBF8D124A,0xA4E0,0x402F,0x81,0x52,0x4E,0xF3,0x77,0xE6,0x25,0x86);


MIDL_DEFINE_GUID(CLSID, CLSID_ExplorerBand,0x553891B7,0xA0D5,0x4526,0xBE,0x18,0xD3,0xCE,0x46,0x1D,0x63,0x10);


MIDL_DEFINE_GUID(CLSID, CLSID_ClassicCopyExt,0x8C83ACB1,0x75C3,0x45D2,0x88,0x2C,0xEF,0xA3,0x23,0x33,0x49,0x1C);


MIDL_DEFINE_GUID(CLSID, CLSID_ExplorerBHO,0x449D0D6E,0x2412,0x4E61,0xB6,0x8F,0x1C,0xB6,0x25,0xCD,0x9E,0x52);


MIDL_DEFINE_GUID(CLSID, CLSID_ShareOverlay,0x594D4122,0x1F87,0x41E2,0x96,0xC7,0x82,0x5F,0xB4,0x79,0x65,0x16);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



