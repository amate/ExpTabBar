

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.01.0622 */
/* at Tue Jan 19 12:14:07 2038
 */
/* Compiler settings for ExpTabBar.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.01.0622 
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

#ifndef __ExpTabBar_i_h__
#define __ExpTabBar_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IExpTabBand_FWD_DEFINED__
#define __IExpTabBand_FWD_DEFINED__
typedef interface IExpTabBand IExpTabBand;

#endif 	/* __IExpTabBand_FWD_DEFINED__ */


#ifndef __ExpTabBand_FWD_DEFINED__
#define __ExpTabBand_FWD_DEFINED__

#ifdef __cplusplus
typedef class ExpTabBand ExpTabBand;
#else
typedef struct ExpTabBand ExpTabBand;
#endif /* __cplusplus */

#endif 	/* __ExpTabBand_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IExpTabBand_INTERFACE_DEFINED__
#define __IExpTabBand_INTERFACE_DEFINED__

/* interface IExpTabBand */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IExpTabBand;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("2625D19C-D2BD-4EF1-B904-0DE1BEF7ACE1")
    IExpTabBand : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IExpTabBandVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IExpTabBand * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IExpTabBand * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IExpTabBand * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IExpTabBand * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IExpTabBand * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IExpTabBand * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IExpTabBand * This,
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
    } IExpTabBandVtbl;

    interface IExpTabBand
    {
        CONST_VTBL struct IExpTabBandVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IExpTabBand_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IExpTabBand_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IExpTabBand_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IExpTabBand_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IExpTabBand_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IExpTabBand_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IExpTabBand_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IExpTabBand_INTERFACE_DEFINED__ */



#ifndef __ExpTabBarLib_LIBRARY_DEFINED__
#define __ExpTabBarLib_LIBRARY_DEFINED__

/* library ExpTabBarLib */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_ExpTabBarLib;

EXTERN_C const CLSID CLSID_ExpTabBand;

#ifdef __cplusplus

class DECLSPEC_UUID("6BE37F72-F2B9-4887-B819-5F086BE270CF")
ExpTabBand;
#endif
#endif /* __ExpTabBarLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


