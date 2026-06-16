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

// Thin variadic shim so the generated methods can keep their familiar
// "CustomInvokeHelper(dispid, flags, vtRet, pvRet, parms, ...)" call shape,
// mirroring COleDispatchDriver::InvokeHelper. It forwards m_lpDispatch from
// the COleDispatchDriver base.
void CustomInvokeHelper_InvokeForward(
    COleDispatchDriver* pThis,
    DISPID dwDispID, WORD wFlags, VARTYPE vtRet, void* pvRet, const BYTE* pbParamInfo, ...);

// Convenience macro used inside the wrapper methods. Because these methods are
// members of classes derived from COleDispatchDriver, "this" is available.
#define CustomInvokeHelper(dispid, flags, vtRet, pvRet, parms, ...) CustomInvokeHelper_InvokeForward(this, (dispid), (flags), (vtRet), (pvRet), (parms), __VA_ARGS__)

// Compatibility alias used by wrappers that explicitly opt in only for
// VTS_PVARIANT signatures.
#define CustomInvokerHelper(dispid, flags, vtRet, pvRet, parms, ...) CustomInvokeHelper(dispid, flags, vtRet, pvRet, parms, __VA_ARGS__)


