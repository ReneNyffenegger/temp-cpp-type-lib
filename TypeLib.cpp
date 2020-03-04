#include "TypeLib.h"

#include <oaidl.h>
#include "VariantHelper.h"

TypeLib::TypeLib() :
    typeLib_      ( 0), 
    nofTypeInfos_ ( 0),
    curTypeInfo_  (-1),
    curITypeInfo_ ( 0), 
    curTypeAttr_  ( 0),
    curFuncDesc_  ( 0),
    curVarDesc_   ( 0),
    funcNames_    ( 0)
{
  ::OleInitialize(0);
}

bool TypeLib::Open(std::wstring const& type_lib_file) {
//std::wstring ws_type_lib_file = s2ws(type_lib_file);
//HRESULT hr = ::LoadTypeLibW(ws_type_lib_file.c_str(), &typeLib_);
  HRESULT hr = ::LoadTypeLib(type_lib_file.c_str(), &typeLib_); // There is no LoadTypeLibW variant.

  if (hr != S_OK) {
    std::cout << "Error with LoadTypeLibrary" << std::endl;

    return false;
  }
  nofTypeInfos_ = typeLib_->GetTypeInfoCount();
  return true;
}

bool TypeLib::NextTypeInfo() {
  curTypeInfo_++;

  curFunc_     = -1;
  curVar_      = -1;

  if (curTypeInfo_ >= nofTypeInfos_) return false;

  // TODO: warum curTypeAttr_
  if (curTypeAttr_ && curITypeInfo_) {
    curITypeInfo_ -> ReleaseTypeAttr(curTypeAttr_);
  }

  if (typeLib_) {
    HRESULT hr=typeLib_-> GetTypeInfo(curTypeInfo_, &curITypeInfo_);
    if (hr == S_OK) {
      curITypeInfo_ -> GetTypeAttr(&curTypeAttr_);

      return true;
    }
    return false;
  }
  return false;
}

bool TypeLib::IsTypeEnum() {
  return IsTypeKind(TKIND_ENUM);
}
bool TypeLib::IsTypeRecord() {
  return IsTypeKind(TKIND_RECORD);
}
bool TypeLib::IsTypeModule() {
  return IsTypeKind(TKIND_MODULE);
}
bool TypeLib::IsTypeInterface() {
  return IsTypeKind(TKIND_INTERFACE);
}
bool TypeLib::IsTypeDispatch() {
  return IsTypeKind(TKIND_DISPATCH);
}
bool TypeLib::IsTypeCoClass() {
  return IsTypeKind(TKIND_COCLASS);
}
bool TypeLib::IsTypeAlias() {
  return IsTypeKind(TKIND_ALIAS);
}
bool TypeLib::IsTypeUnion() {
  return IsTypeKind(TKIND_UNION);
}
bool TypeLib::IsTypeMax() {
  return IsTypeKind(TKIND_MAX);
}
bool TypeLib::IsTypeKind(int i) {
  TYPEKIND tk = curTypeAttr_->typekind;

  return i == tk;
}

std::wstring TypeLib::LibDocumentation() {
    if (! typeLib_) return L"No Type Library open!";
 
    BSTR name;
    BSTR doc;
    unsigned long ctx;
    
    HRESULT hr = typeLib_->GetDocumentation(
      -1,
      &name,
      &doc,
      &ctx,
      0  // Help File
      );
 
    if (hr != S_OK) {
      std::wcout << L"GetDocumentation failed" << std::endl;
      return L"";
    }
    
    std::wstring sName = name;
    std::wstring sDoc;
    if (doc) sDoc = doc;
 // std::string sName = ws2s(name);
 // std::string sDoc;
 // if (doc) {
 //   sDoc = ws2s(doc);
 // }
    
    ::SysFreeString(name);
    ::SysFreeString(doc );
    
    return sName + L": " + sDoc;
}

std::wstring TypeLib::TypeDocumentation() {
  return TypeDocumentation_(curITypeInfo_);
}

std::wstring TypeLib::TypeDocumentation_(ITypeInfo* i) {
  BSTR name;
  unsigned long ctx;
  
  HRESULT hr = i->GetDocumentation(
    MEMBERID_NIL, //idx,
    &name,
    0,//&doc,
    &ctx,
    0  // Help File
    );

  if (hr != S_OK) {
    std::wcout << L"GetDocumentation failed" << std::endl;
    return L"";
  }
  
//std::string sName = ws2s(name);
  std::wstring sName = name;
  
  ::SysFreeString(name);
  return sName;
}

TypeLib::INVOKEKIND TypeLib::InvokeKind() {
  return static_cast<INVOKEKIND>(curFuncDesc_->invkind);
}

TypeLib::VARIABLEKIND TypeLib::VariableKind() {
  return static_cast<VARIABLEKIND>(curVarDesc_->varkind);
}

bool TypeLib::NextFunction() {
  ReleaseFuncNames();

  curFunc_ ++;
  curFuncParam_= -1;

  if (curFunc_ >= curTypeAttr_->cFuncs) return false;

  if (curFuncDesc_) {
    curITypeInfo_->ReleaseFuncDesc(curFuncDesc_);
  }

  HRESULT hr = curITypeInfo_-> GetFuncDesc(curFunc_, &curFuncDesc_);
  GetFuncNames();

  if (hr != S_OK) {
    std::wcout << L"GetFuncDesc failed" << std::endl;
    return false;
  }

  // Is that correct? What is actually a IMPLTYPEFLAG?
  curITypeInfo_ -> GetImplTypeFlags (curFunc_, &curImplTypeFlags_);

  return true;
}

bool TypeLib::HasFunctionTypeFlag(TYPEFLAG tf) {
  return curImplTypeFlags_ & static_cast<int>(tf);
}

std::wstring TypeLib::ConstValue() {
  if (VariableKind() != const_) {
    return L"VariableKind must be const_";
  }
  //VARIANT v=*(curVarDesc_->lpvarValue);
  variant v=*(curVarDesc_->lpvarValue);

  return v.ValueAsString();
}

std::wstring TypeLib::UserdefinedType(HREFTYPE hrt) {
  std::wstring tp=L"";

  ITypeInfo* itpi;
  curITypeInfo_ -> GetRefTypeInfo(hrt, &itpi);
  std::wstring ref_type = TypeDocumentation_(itpi);
  tp += ref_type;

  return tp;
}

std::wstring TypeLib::Type(ELEMDESC const& ed) {
  TYPEDESC td = ed.tdesc;

  std::wstring tp = TypeAsString(td.vt);

  if (td.vt==VT_PTR) {
    TYPEDESC ptr_td = *(td.lptdesc);
    tp += L" (";
    tp += TypeAsString(ptr_td.vt);
    if (ptr_td.vt == VT_USERDEFINED) {
      tp += L" (";
      HREFTYPE hrt = ptr_td.hreftype;
      tp += UserdefinedType(hrt);
      tp += L")";
    }
    tp += L")";
  }
  else if (td.vt == VT_USERDEFINED) {
    tp += L" (";
    tp += UserdefinedType(td.hreftype);
    tp += L")";
  }
  else if (td.vt == VT_SAFEARRAY) {
    // TODO
  }

  return tp;
}

std::wstring TypeLib::ParameterType() {
  ELEMDESC ed = curFuncDesc_->lprgelemdescParam[curFuncParam_];

  return Type(ed);
}

std::wstring TypeLib::VariableType() {
  ELEMDESC ed = curVarDesc_->elemdescVar;

  return Type(ed);
}

std::wstring TypeLib::ReturnType() {
  ELEMDESC ed = curFuncDesc_->elemdescFunc;

  return Type(ed);
}

bool TypeLib::NextVariable() {
  curVar_ ++;
  // TODO: curVar_->cVars accessed through NofVariables().
  if (curVar_ >= curTypeAttr_->cVars) return false;

  if (curVarDesc_) {
    curITypeInfo_->ReleaseVarDesc(curVarDesc_);
  }

  HRESULT hr = curITypeInfo_->GetVarDesc(curVar_, &curVarDesc_);

  if (hr != S_OK) {
    std::cout << L"GetVarDesc failed" << std::endl;
    return false;
  }

  return true;
}

bool TypeLib::NextParameter() {
  curFuncParam_ ++;
  if (curFuncParam_ >= NofParameters()) return false;

  return true;
}

std::wstring TypeLib::VariableName() {
  BSTR varName;

  unsigned int dummy;

  HRESULT hr = curITypeInfo_->GetNames(curVarDesc_->memid, &varName, 1, &dummy);

  if (hr!= S_OK) {
    return L"GetNames failed";
  }

  std::wstring ret = varName;
//std::string ret = ws2s(varName);

  ::SysFreeString(varName);

  return ret;
}

std::wstring TypeLib::ParameterName() {
  if (1+curFuncParam_ >= static_cast<int>(nofFuncNames_)) return L"<nameless>";

//BSTR paramName = funcNames_[1+curFuncParam_];

//std::wstring ret = paramName;
//std::string ret = ws2s(paramName);

//return ret;
  return funcNames_[1+curFuncParam_];
}

// http://doc.ddart.net/msdn/header/include/oaidl.h.html
//  return ParameterIsHasX(PARAMFLAG_NONE);
bool TypeLib::ParameterIsIn() {
  return ParameterIsHasX(PARAMFLAG_FIN);
}
bool TypeLib::ParameterIsOut() {
  return ParameterIsHasX(PARAMFLAG_FOUT);
}
bool TypeLib::ParameterIsFLCID() {
  return ParameterIsHasX(PARAMFLAG_FLCID);
}
bool TypeLib::ParameterIsReturnValue() {
  return ParameterIsHasX(PARAMFLAG_FRETVAL);
}
bool TypeLib::ParameterIsOptional() {
  return ParameterIsHasX(PARAMFLAG_FOPT);
}
bool TypeLib::ParameterHasDefault() {
  return ParameterIsHasX(PARAMFLAG_FHASDEFAULT);
}
bool TypeLib::ParameterHasCustumData() {
  return ParameterIsHasX(0x40 /*PARAMFLAG_FHASCUSTDATA*/);
}
bool TypeLib::ParameterIsHasX(int flag) {
  ELEMDESC e = curFuncDesc_->lprgelemdescParam[curFuncParam_];
  return e.paramdesc.wParamFlags & flag;
}

void TypeLib::ReleaseFuncNames() {
  if (funcNames_) {
    for (int i=0;i<static_cast<int>(nofFuncNames_); i++) {
      ::SysFreeString(funcNames_[i]);
    }

    delete [] funcNames_;
    funcNames_=0;
  }
}

void TypeLib::GetFuncNames() {

  unsigned int nof_names = 1 + NofParameters();

  funcNames_ = new BSTR[nof_names];

  HRESULT hr = curITypeInfo_->GetNames(curFuncDesc_->memid, funcNames_, nof_names, &nofFuncNames_);

  if (hr!= S_OK) {
    std::cout << L"GetNames failed" << std::endl;
  }
}

std::wstring TypeLib::FunctionName() {
//BSTR funcName = funcNames_[0];

//std::string ret = ws2s(funcName);

  return funcNames_[0];
//return ret;
}

int TypeLib::NofVariables() {
  return curTypeAttr_->cVars;
}

int TypeLib::NofFunctions() {
  return curTypeAttr_->cFuncs;
}

int TypeLib::NofTypeInfos() {
  return nofTypeInfos_;
}

short TypeLib::NofParameters() {
  return curFuncDesc_->cParams;
}

short TypeLib::NofOptionalParameters() {
  return curFuncDesc_->cParamsOpt;
}
  
