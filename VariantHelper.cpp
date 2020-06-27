#include "VariantHelper.h"
// #include "win32_Unicode.h"
#include <iostream>

std::wstring TypeAsString(VARTYPE vt){
  switch (vt) {
    case VT_EMPTY            : return L"VT_EMPTY"             ;
    case VT_NULL             : return L"VT_NULL"              ;
    case VT_I2               : return L"VT_I2"                ;
    case VT_I4               : return L"VT_I4"                ;
    case VT_R4               : return L"VT_R4"                ;
    case VT_R8               : return L"VT_R8"                ;
    case VT_CY               : return L"VT_CY"                ;
    case VT_DATE             : return L"VT_DATE"              ;
    case VT_BSTR             : return L"VT_BSTR"              ;
    case VT_DISPATCH         : return L"VT_DISPATCH"          ;
    case VT_ERROR            : return L"VT_ERROR"             ;
    case VT_BOOL             : return L"VT_BOOL"              ;
    case VT_VARIANT          : return L"VT_VARIANT"           ;
    case VT_DECIMAL          : return L"VT_DECIMAL"           ;
    case VT_RECORD           : return L"VT_RECORD"            ;
    case VT_UNKNOWN          : return L"VT_UNKNOWN"           ;
    case VT_I1               : return L"VT_I1"                ;
    case VT_UI1              : return L"VT_UI1"               ;
    case VT_UI2              : return L"VT_UI2"               ;
    case VT_UI4              : return L"VT_UI4"               ;
    case VT_INT              : return L"VT_INT"               ;
    case VT_UINT             : return L"VT_UINT"              ;
    case VT_VOID             : return L"VT_VOID"              ;
    case VT_HRESULT          : return L"VT_HRESULT"           ;
    case VT_PTR              : return L"VT_PTR"               ;
    case VT_SAFEARRAY        : return L"VT_SAFEARRAY"         ;
    case VT_CARRAY           : return L"VT_CARRAY"            ;
    case VT_USERDEFINED      : return L"VT_USERDEFINED"       ;
    case VT_LPSTR            : return L"VT_LPSTR"             ;
    case VT_LPWSTR           : return L"VT_LPWSTR"            ;
    case VT_BLOB             : return L"VT_BLOB"              ;
    case VT_STREAM           : return L"VT_STREAM"            ;
    case VT_STORAGE          : return L"VT_STORAGE"           ;
    case VT_STREAMED_OBJECT  : return L"VT_STREAMED_OBJECT"   ;
    case VT_STORED_OBJECT    : return L"VT_STORED_OBJECT"     ;
    case VT_BLOB_OBJECT      : return L"VT_BLOB_OBJECT"       ;
    case VT_CF               : return L"VT_CF"                ; // Clipboard Format 
    case VT_CLSID            : return L"VT_CLSID"             ;
    case VT_VECTOR           : return L"VT_VECTOR"            ;
    case VT_ARRAY            : return L"VT_ARRAY"             ;
  //case VT_BYREF            : return L"VT_BYREF"             ;
    case VT_BYREF|VT_DECIMAL : return L"VT_BYREF|VT_DECIMAL"  ;
  //case VT_ARRAY|*:        // return L"VT_ARRAY|*"           ;
    case VT_BYREF|VT_UI1     : return L"VT_BYREF|VT_UI1"      ;
    case VT_BYREF|VT_I2      : return L"VT_BYREF|VT_I2"       ;
    case VT_BYREF|VT_I4      : return L"VT_BYREF|VT_I4"       ;
    case VT_BYREF|VT_R4      : return L"VT_BYREF|VT_R4"       ;
    case VT_BYREF|VT_R8      : return L"VT_BYREF|VT_R8"       ;
    case VT_BYREF|VT_BOOL    : return L"VT_BYREF|VT_BOOL"     ;
    case VT_BYREF|VT_ERROR   : return L"VT_BYREF|VT_ERROR"    ;
    case VT_BYREF|VT_CY      : return L"VT_BYREF|VT_CY"       ;
    case VT_BYREF|VT_DATE    : return L"VT_BYREF|VT_DATE"     ;
    case VT_BYREF|VT_BSTR    : return L"VT_BYREF|VT_BSTR"     ;
    case VT_BYREF|VT_UNKNOWN : return L"VT_BYREF|VT_UNKNOWN"  ;
    case VT_BYREF|VT_DISPATCH: return L"VT_BYREF|VT_DISPATCH" ;

  //case    VT_ARRAY|*: return L"VT_ARRAY|*";

    case    VT_BYREF|VT_VARIANT: return L"VT_BYREF|VT_VARIANT";

  //Generic case ByRef:return L"ByRef";

    case    VT_BYREF|VT_I1   : return L"VT_BYREF|VT_I1"       ;
    case    VT_BYREF|VT_UI2  : return L"VT_BYREF|VT_UI2"      ;
    case    VT_BYREF|VT_UI4  : return L"VT_BYREF|VT_UI4"      ;
    case    VT_BYREF|VT_INT  : return L"VT_BYREF|VT_INT"      ;
    case    VT_BYREF|VT_UINT : return L"VT_BYREF|VT_UINT"     ;

    default: return L"Unknown type";
  }
}

std::wstring VariantTypeAsString(const VARIANT& v) {
  return TypeAsString(v.vt);
}

variant::variant(VARIANT v) : v_(v) {
}

variant::~variant() {
  ::VariantClear(&v_);
}

std::wstring variant::ValueAsString() {
  if (v_.vt != VT_BSTR) {
     ChangeType(VT_BSTR);
  }
//std::wcout << L"ret" << std::endl;
//std::wstring ret;
//std::wcout << L"going to assign" << std::endl;
//ret = v_.bstrVal;
//std::wcout << L"going to return v_.bstrVal" << std::endl;
//return ret;
//std::wcout << L"return v_.bstrVal" << std::endl;
//std::wstring ret = v_.bstrVal;
//std::wcout << L"ret = " << std::endl;
//std::wcout << ret << std::endl;
//return ret;
  return v_.bstrVal;

//return ws2s(v_.bstrVal);
}

void variant::ChangeType(VARTYPE vt) {
   ::VariantChangeType(
     &v_              , // pvargDest (if same as pvarSrc: in-place conversion)
     &v_              , // pvarSrc
     VARIANT_ALPHABOOL, // VARIANT_ALPHABOOL: VT_BOOL as string will be 'True' or 'False'
     vt                 // Dest-type
   );
}
