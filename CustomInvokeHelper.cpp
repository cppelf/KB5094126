/////////////////////////////////////////////////////////////////////////////
// CustomInvokeHelper - replacement for COleDispatchDriver::InvokeHelper
//
// This parses the same MFC VTS_* parameter descriptor string used by
// InvokeHelper, builds an explicit VARIANTARG array + DISPPARAMS, and calls
// IDispatch::Invoke directly.
//
// The KB workaround lives in the VTS_PVARIANT branch: instead of passing the
// caller's VARIANT* as (VT_VARIANT | VT_BYREF), the *actual* value is copied
// into the argument via VariantCopyInd(). A NULL source pointer is passed as a
// missing optional argument (VT_ERROR / DISP_E_PARAMNOTFOUND).
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <afxdisp.h>   // COleDispatchDriver + VTS_* macros

// ---------------------------------------------------------------------------
// IMPORTANT: MFC's VTS_* tokens are STRING-LITERAL macros, e.g.
//     #define VTS_I2  "\x02"
// They concatenate into a byte string for the 'parms' descriptor, but they
// are NOT integers, so they cannot be used as switch/case labels.
// We therefore switch on the underlying byte VALUES, defined below.
//
// These byte values are taken from MFC's <afxdisp.h>. They have been stable
// across VC++ versions, but if you hit a mismatch, open your afxdisp.h and
// confirm (e.g. VTS_I2 == "\x02", VTS_PVARIANT == "\x4c").
// ---------------------------------------------------------------------------
enum VtsByte {
    kVTS_I2 = 0x02,
    kVTS_I4 = 0x03,
    kVTS_R4 = 0x04,
    kVTS_R8 = 0x05,
    kVTS_BSTR = 0x08,
    kVTS_DISPATCH = 0x09,
    kVTS_BOOL = 0x0B,
    kVTS_VARIANT = 0x0C,
    kVTS_UNKNOWN = 0x0D,

    kVTS_PI2 = 0x42,
    kVTS_PI4 = 0x43,
    kVTS_PR4 = 0x44,
    kVTS_PR8 = 0x45,
    kVTS_PBSTR = 0x48,
    kVTS_PDISPATCH = 0x49,
    kVTS_PBOOL = 0x4B,
    kVTS_PVARIANT = 0x4C,
    kVTS_PUNKNOWN = 0x4D
};

namespace {
// Maximum number of arguments we expect any single wrapper to pass.
const int kMaxInvokeArgs = 64;

// Returns TRUE if this dispid/method is one of the special "close"-style
// calls (Document.Close, Documents.Close, Application.Quit) for which we
// (a) drop trailing missing optional args, (b) coerce SaveChanges to VT_I4,
// and (c) tolerate an already-released dispatch.
inline BOOL IsCloseOrQuitDispId(DISPID dispid)
{
    // 0x6 == Close (Document / Documents), 0x453 == Quit (Application)
    // These are the standard Word object-model dispids emitted by ClassWizard.
    return (dispid == (DISPID)0x6) || (dispid == (DISPID)0x453);
}
} // namespace

/////////////////////////////////////////////////////////////////////////////
// The actual helper. It is implemented as a free function so it can be shared
// by every wrapper class without editing the generated headers.

static void AFX_CDECL CustomInvokeHelperImpl(
    LPDISPATCH        pDispatch,   // m_lpDispatch of the calling wrapper
    DISPID            dwDispID,
    WORD              wFlags,
    VARTYPE           vtRet,
    void*             pvRet,
    const BYTE*       pbParamInfo,
    va_list           argList)
{
    //---------------------------------------------------------------------
    // 0) Special handling when the dispatch pointer is gone.
    //    If ReleaseDispatch() has already been called and this is a
    //    Close/Quit call, return silently (the object is effectively gone).
    //---------------------------------------------------------------------
    if (pDispatch == NULL) {
        if (wFlags == DISPATCH_METHOD && IsCloseOrQuitDispId(dwDispID))
            return; // silent no-op for Close/Quit
        AfxThrowOleException(E_POINTER);
        return;
    }

    //---------------------------------------------------------------------
    // 1) Walk the VTS_* descriptor and collect the raw argument pointers/
    //    values from the va_list, in source (left-to-right) order.
    //---------------------------------------------------------------------
    VARIANTARG args[kMaxInvokeArgs];
    int        nArgs = 0;
    BOOL       bArgIsMissing[kMaxInvokeArgs];

    for (int i = 0; i < kMaxInvokeArgs; ++i) {
        VariantInit(&args[i]);
        bArgIsMissing[i] = FALSE;
    }

    const BYTE* p = pbParamInfo;
    while (p != NULL && *p != 0 /* VTS_NONE terminator */) {
        ASSERT(nArgs < kMaxInvokeArgs);
        VARIANTARG& v = args[nArgs];
        BYTE vts = *p++;

        switch (vts) {
            case kVTS_I2:        // short by value
                v.vt = VT_I2;
                v.iVal = (SHORT)va_arg(argList, int);
                break;

            case kVTS_I4:        // long by value
                v.vt = VT_I4;
                v.lVal = va_arg(argList, long);
                break;

            case kVTS_R4:        // float by value
                v.vt = VT_R4;
                v.fltVal = (float)va_arg(argList, double);
                break;

            case kVTS_R8:        // double by value
                v.vt = VT_R8;
                v.dblVal = va_arg(argList, double);
                break;

            case kVTS_BOOL:      // BOOL by value -> VARIANT_BOOL
                v.vt = VT_BOOL;
                v.boolVal = (va_arg(argList, int) != 0) ? VARIANT_TRUE : VARIANT_FALSE;
                break;

            case kVTS_BSTR:      // LPCOLESTR / CString -> BSTR
            {
                LPCOLESTR psz = va_arg(argList, LPCOLESTR);
                v.vt = VT_BSTR;
                v.bstrVal = (psz != NULL) ? SysAllocString(psz) : NULL;
            }
            break;

            case kVTS_DISPATCH:  // LPDISPATCH by value
            {
                LPDISPATCH pd = va_arg(argList, LPDISPATCH);
                v.vt = VT_DISPATCH;
                v.pdispVal = pd;
                if (pd != NULL)
                    pd->AddRef();
            }
            break;

            case kVTS_UNKNOWN:   // LPUNKNOWN by value
            {
                LPUNKNOWN pu = va_arg(argList, LPUNKNOWN);
                v.vt = VT_UNKNOWN;
                v.punkVal = pu;
                if (pu != NULL)
                    pu->AddRef();
            }
            break;

            case kVTS_VARIANT:   // VARIANT by value (caller passed a VARIANT*)
            {
                const VARIANT* psrc = va_arg(argList, const VARIANT*);
                if (psrc != NULL)
                    VariantCopy(&v, (VARIANTARG*)psrc);
                else {
                    v.vt = VT_ERROR;
                    v.scode = DISP_E_PARAMNOTFOUND;
                    bArgIsMissing[nArgs] = TRUE;
                }
            }
            break;

            //=================================================================
            // THE KB WORKAROUND
            //
            // VTS_PVARIANT: ClassWizard passes a VARIANT* and stock MFC forwards
            // it as (VT_VARIANT | VT_BYREF). Post-KB Office rejects that with
            // DISP_E_TYPEMISMATCH. Instead, de-reference and copy the VALUE.
            //=================================================================
            case kVTS_PVARIANT:
            {
                VARIANT* psrc = va_arg(argList, VARIANT*);
                if (psrc != NULL) {
                    // Copy by value, indirecting through any BYREF the caller
                    // may itself have set up.
                    VariantCopyInd(&v, psrc);

                    // If, after indirection, this is a "missing optional"
                    // (VT_EMPTY / VT_ERROR=PARAMNOTFOUND), mark it so trailing
                    // copies can be dropped for Close/Quit.
                    if (v.vt == VT_EMPTY ||
                        (v.vt == VT_ERROR && v.scode == DISP_E_PARAMNOTFOUND)) {
                        v.vt = VT_ERROR;
                        v.scode = DISP_E_PARAMNOTFOUND;
                        bArgIsMissing[nArgs] = TRUE;
                    }
                } else {
                    // NULL pointer => missing optional argument.
                    v.vt = VT_ERROR;
                    v.scode = DISP_E_PARAMNOTFOUND;
                    bArgIsMissing[nArgs] = TRUE;
                }
            }
            break;

            // By-ref scalar descriptors (rarely emitted, but handle gracefully).
            case kVTS_PI2:
                v.vt = VT_I2 | VT_BYREF;  v.piVal = va_arg(argList, SHORT*);     break;
            case kVTS_PI4:
                v.vt = VT_I4 | VT_BYREF;  v.plVal = va_arg(argList, long*);      break;
            case kVTS_PR4:
                v.vt = VT_R4 | VT_BYREF;  v.pfltVal = va_arg(argList, float*);     break;
            case kVTS_PR8:
                v.vt = VT_R8 | VT_BYREF;  v.pdblVal = va_arg(argList, double*);    break;
            case kVTS_PBOOL:
                v.vt = VT_BOOL | VT_BYREF; v.pboolVal = va_arg(argList, VARIANT_BOOL*); break;
            case kVTS_PBSTR:
                v.vt = VT_BSTR | VT_BYREF; v.pbstrVal = va_arg(argList, BSTR*);    break;
            case kVTS_PDISPATCH:
                v.vt = VT_DISPATCH | VT_BYREF; v.ppdispVal = va_arg(argList, LPDISPATCH*); break;
            case kVTS_PUNKNOWN:
                v.vt = VT_UNKNOWN | VT_BYREF; v.ppunkVal = va_arg(argList, LPUNKNOWN*); break;

            default:
                // Unknown descriptor: skip a pointer-sized arg to stay aligned.
                (void)va_arg(argList, void*);
                v.vt = VT_ERROR;
                v.scode = DISP_E_PARAMNOTFOUND;
                bArgIsMissing[nArgs] = TRUE;
                break;
        }

        ++nArgs;
    }

    //---------------------------------------------------------------------
    // 2) Special handling for Close / Quit:
    //    - coerce a leading SaveChanges (VT_BOOL) to VT_I4 (-1 / 0)
    //    - drop trailing MISSING optional args entirely
    //---------------------------------------------------------------------
    int nEffectiveArgs = nArgs;

    if (IsCloseOrQuitDispId(dwDispID) && wFlags == DISPATCH_METHOD) {
        // SaveChanges is the first positional arg of Close/Quit.
        if (nEffectiveArgs >= 1) {
            VARIANTARG& sc = args[0];
            if (sc.vt == VT_BOOL) {
                long lVal = (sc.boolVal != VARIANT_FALSE) ? -1L : 0L;
                VariantClear(&sc);
                sc.vt = VT_I4;
                sc.lVal = lVal;
            }
        }

        // Drop trailing missing optional arguments.
        while (nEffectiveArgs > 0 && bArgIsMissing[nEffectiveArgs - 1])
            --nEffectiveArgs;
    }

    //---------------------------------------------------------------------
    // 3) Build DISPPARAMS. rgvarg must be in REVERSE argument order.
    //---------------------------------------------------------------------
    VARIANTARG  rgvarg[kMaxInvokeArgs];
    DISPID      dispidNamed = DISPID_PROPERTYPUT;
    DISPPARAMS  dispparams;
    memset(&dispparams, 0, sizeof(dispparams));

    for (int j = 0; j < nEffectiveArgs; ++j)
        rgvarg[j] = args[nEffectiveArgs - 1 - j];

    dispparams.rgvarg = (nEffectiveArgs > 0) ? rgvarg : NULL;
    dispparams.cArgs = (UINT)nEffectiveArgs;
    dispparams.rgdispidNamedArgs = NULL;
    dispparams.cNamedArgs = 0;

    // For property put / putref, the single value uses the named arg
    // DISPID_PROPERTYPUT.
    if (wFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) {
        dispparams.cNamedArgs = 1;
        dispparams.rgdispidNamedArgs = &dispidNamed;
    }

    //---------------------------------------------------------------------
    // 4) Invoke.
    //---------------------------------------------------------------------
    VARIANT     vaResult;
    VariantInit(&vaResult);
    EXCEPINFO   excepInfo;
    memset(&excepInfo, 0, sizeof(excepInfo));
    UINT        uArgErr = (UINT)-1;

    VARIANT*    pvaResult = (vtRet == VT_EMPTY) ? NULL : &vaResult;

    HRESULT hr = pDispatch->Invoke(
        dwDispID,
        IID_NULL,
        LOCALE_USER_DEFAULT,
        wFlags,
        &dispparams,
        pvaResult,
        &excepInfo,
        &uArgErr);

    //---------------------------------------------------------------------
    // 5) Clean up the argument VARIANTs we own.
    //    (Free everything we allocated/AddRef'd via VariantClear.)
    //---------------------------------------------------------------------
    for (int k = 0; k < nArgs; ++k)
        VariantClear(&args[k]);

    //---------------------------------------------------------------------
    // 6) Error handling.
    //---------------------------------------------------------------------
    if (FAILED(hr)) {
        // Free EXCEPINFO BSTR members before throwing.
        if (hr == DISP_E_EXCEPTION) {
            if (excepInfo.pfnDeferredFillIn != NULL)
                (*excepInfo.pfnDeferredFillIn)(&excepInfo);
        }

        // Capture the strings, then free the BSTRs.
        SysFreeString(excepInfo.bstrSource);
        SysFreeString(excepInfo.bstrDescription);
        SysFreeString(excepInfo.bstrHelpFile);
        excepInfo.bstrSource = NULL;
        excepInfo.bstrDescription = NULL;
        excepInfo.bstrHelpFile = NULL;

        VariantClear(&vaResult);

        // For Close/Quit, tolerate "object disconnected" style failures.
        if (IsCloseOrQuitDispId(dwDispID) && wFlags == DISPATCH_METHOD &&
            (hr == RPC_E_DISCONNECTED || hr == CO_E_OBJNOTCONNECTED ||
                hr == RPC_E_SERVER_DIED || hr == RPC_E_SERVER_DIED_DNE)) {
            return; // silent
        }

        if (hr == DISP_E_EXCEPTION)
            AfxThrowOleDispatchException(0, _T("OLE Dispatch exception during late-bound Invoke."));
        else
            AfxThrowOleException(hr);
        return;
    }

    //---------------------------------------------------------------------
    // 7) Convert the result back into the MFC-expected return type.
    //---------------------------------------------------------------------
    if (vtRet != VT_EMPTY && pvRet != NULL) {
        switch (vtRet) {
            case VT_I2:   // short
            {
                VARIANT v; VariantInit(&v);
                if (SUCCEEDED(VariantChangeType(&v, &vaResult, 0, VT_I2)))
                    *(short*)pvRet = v.iVal;
                VariantClear(&v);
            }
            break;

            case VT_I4:   // long
            {
                VARIANT v; VariantInit(&v);
                if (SUCCEEDED(VariantChangeType(&v, &vaResult, 0, VT_I4)))
                    *(long*)pvRet = v.lVal;
                VariantClear(&v);
            }
            break;

            case VT_R4:   // float
            {
                VARIANT v; VariantInit(&v);
                if (SUCCEEDED(VariantChangeType(&v, &vaResult, 0, VT_R4)))
                    *(float*)pvRet = v.fltVal;
                VariantClear(&v);
            }
            break;

            case VT_R8:   // double
            {
                VARIANT v; VariantInit(&v);
                if (SUCCEEDED(VariantChangeType(&v, &vaResult, 0, VT_R8)))
                    *(double*)pvRet = v.dblVal;
                VariantClear(&v);
            }
            break;

            case VT_BOOL: // BOOL
            {
                VARIANT v; VariantInit(&v);
                if (SUCCEEDED(VariantChangeType(&v, &vaResult, 0, VT_BOOL)))
                    *(BOOL*)pvRet = (v.boolVal != VARIANT_FALSE);
                VariantClear(&v);
            }
            break;

            case VT_BSTR: // CString
            {
                VARIANT v; VariantInit(&v);
                if (SUCCEEDED(VariantChangeType(&v, &vaResult, 0, VT_BSTR)))
                    *(CString*)pvRet = v.bstrVal;  // CString::operator=(BSTR) copies
                VariantClear(&v);
            }
            break;

            case VT_DISPATCH: // LPDISPATCH (ownership transferred to caller)
            {
                if (vaResult.vt == VT_DISPATCH) {
                    *(LPDISPATCH*)pvRet = vaResult.pdispVal;
                    vaResult.pdispVal = NULL; // prevent VariantClear from releasing
                    vaResult.vt = VT_EMPTY;
                } else {
                    *(LPDISPATCH*)pvRet = NULL;
                }
            }
            break;

            case VT_UNKNOWN: // LPUNKNOWN (ownership transferred to caller)
            {
                if (vaResult.vt == VT_UNKNOWN) {
                    *(LPUNKNOWN*)pvRet = vaResult.punkVal;
                    vaResult.punkVal = NULL;
                    vaResult.vt = VT_EMPTY;
                } else {
                    *(LPUNKNOWN*)pvRet = NULL;
                }
            }
            break;

            case VT_VARIANT: // VARIANT (caller owns and must VariantClear)
            {
                VariantClear((VARIANT*)pvRet);
                VariantCopy((VARIANT*)pvRet, &vaResult);
            }
            break;

            default:
                // Unhandled return type: best-effort coerce into a VARIANT*.
                break;
        }
    }

    VariantClear(&vaResult);
}

// Thin variadic shim so the generated methods can keep their familiar
// "CustomInvokeHelper(dispid, flags, vtRet, pvRet, parms, ...)" call shape,
// mirroring COleDispatchDriver::InvokeHelper. It forwards m_lpDispatch from
// the COleDispatchDriver base.
void CustomInvokeHelper_InvokeForward(
    COleDispatchDriver* pThis,
    DISPID dwDispID, WORD wFlags, VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ...)
{
    va_list argList;
    va_start(argList, pbParamInfo);
    CustomInvokeHelperImpl(pThis->m_lpDispatch, dwDispID, wFlags, vtRet, pvRet, pbParamInfo, argList);
    va_end(argList);
}

