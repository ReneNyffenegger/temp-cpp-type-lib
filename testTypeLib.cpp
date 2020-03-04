// g++ testTypeLib.cpp TypeLib.cpp VariantHelper.cpp -lole32 -loleaut32 -municode
// 
// a.exe "C:\Program Files\Common Files\microsoft shared\OFFICE14\MSO.DLL"

#include <windows.h>
#include <iostream>

#include "TypeLib.h"

int wmain(int argc, wchar_t *argv[]) {

  TypeLib t;

  if (argc < 2) {
      std::wcout << L"Specify typelib" << std::endl;
      return 1;
  }

  if (!t.Open(
     // "C:\\Program Files\\Microsoft Office\\Office\\MSWORD9.olb"
     // "C:\\Program Files\\Microsoft Office\\root\\Office16\\MSWORD.OLB"
     // "C:\\Program Files\\Common Files\\microsoft shared\\OFFICE14\\MSO.DLL"
     argv[1]
  )) {
    std::wcout << L"Couldn't open type library" << std::endl;
    return -1;
  }

  std::wcout << t.LibDocumentation() << std::endl;

  int nofTypeInfos = t.NofTypeInfos();


  std::wcout << L"Nof Type Infos: " << nofTypeInfos << std::endl;
  
  while (t.NextTypeInfo()) {
    std::wstring type_doc = t.TypeDocumentation(); 

    std::wcout <<                                    std::endl;
    std::wcout <<   type_doc                      << std::endl;
    std::wcout << L"----------------------------" << std::endl;

    std::wcout << "  Interface: ";
    if (t.IsTypeEnum     ())   std::wcout << L"Enum"; 
    if (t.IsTypeRecord   ())   std::wcout << L"Record"; 
    if (t.IsTypeModule   ())   std::wcout << L"Module"; 
    if (t.IsTypeInterface())   std::wcout << L"Interface"; 
    if (t.IsTypeDispatch ())   std::wcout << L"Dispatch"; 
    if (t.IsTypeCoClass  ())   std::wcout << L"CoClass"; 
    if (t.IsTypeAlias    ())   std::wcout << L"Alias"; 
    if (t.IsTypeUnion    ())   std::wcout << L"Union"; 
    if (t.IsTypeMax      ())   std::wcout << L"Max"; 
    std::wcout << std::endl;

    int nofFunctions=t.NofFunctions();
    int nofVariables=t.NofVariables();

    std::wcout << L"  functions: " << nofFunctions << std::endl;
    std::wcout << L"  variables: " << nofVariables << std::endl;

    while (t.NextFunction()) {
      std::wcout << std::endl;
      std::wcout << "  Function     : " << t.FunctionName()          << std::endl;

      std::wcout << "    returns    : " << t.ReturnType() << std::endl;

      std::wcout << "    flags      : ";
      if (t.HasFunctionTypeFlag(TypeLib::FDEFAULT      )) std::wcout << L"FDEFAULT "      ;
      if (t.HasFunctionTypeFlag(TypeLib::FSOURCE       )) std::wcout << L"FSOURCE "       ;
      if (t.HasFunctionTypeFlag(TypeLib::FRESTRICTED   )) std::wcout << L"FRESTRICTED "   ;
      if (t.HasFunctionTypeFlag(TypeLib::FDEFAULTVTABLE)) std::wcout << L"FDEFAULTVTABLE ";
      std::wcout << std::endl;

      TypeLib::INVOKEKIND ik = t.InvokeKind();
      switch (ik) {
        case TypeLib::func: 
          std::wcout <<L"    invoke kind: function" << std::endl;
          break;
        case TypeLib::put: 
          std::wcout <<L"    invoke kind: put" << std::endl;
          break;
        case TypeLib::get: 
          std::wcout <<L"    invoke kind: get" << std::endl;
          break;
        case TypeLib::putref: 
          std::wcout <<L"    invoke kind: put ref" << std::endl;
          break;
        default:
          std::wcout <<L"    invoke kind: ???" << std::endl;
      }

      std::wcout << L"    params     : " << t.NofParameters()         << std::endl;
      std::wcout << L"    params opt : " << t.NofOptionalParameters() << std::endl;

      while (t.NextParameter()) {
        std::wcout << L"    Parameter  : " << t.ParameterName();
        std::wcout << L" type = " << t.ParameterType();
        if (t.ParameterIsIn         ()) std::wcout << L" in";
        if (t.ParameterIsOut        ()) std::wcout << L" out";
        if (t.ParameterIsFLCID      ()) std::wcout << L" flcid";
        if (t.ParameterIsReturnValue()) std::wcout << L" ret";
        if (t.ParameterIsOptional   ()) std::wcout << L" opt";
        if (t.ParameterHasDefault   ()) std::wcout << L" def";
        if (t.ParameterHasCustumData()) std::wcout << L" cust";
        std::wcout << std::endl;
      }
    }
    while (t.NextVariable()) {
      std::wcout << L"  Variable : " << t.VariableName() << std::endl;
      std::wcout << L"      Type : " << t.VariableType();
      TypeLib::VARIABLEKIND vk = t.VariableKind();
      switch (vk) {
        case TypeLib::instance: std::wcout << L" (instance)" << std::endl; break;
        case TypeLib::static_ : std::wcout << L" (static)"   << std::endl; break;
        case TypeLib::const_  : std::wcout << L" (const ";
             std::wcout << t.ConstValue() << L")" << std::endl;            break;
        case TypeLib::dispatch: std::wcout << L" (dispatch)" << std::endl; break;
        default:
          std::wcout << L"    variable kind: unknown" << std::endl;
      }
    }
  }

  return 0;
}
