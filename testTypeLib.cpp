// g++ testTypeLib.cpp TypeLib.cpp win32_Unicode.cpp VariantHelper.cpp -lole32 -loleaut32

#include <windows.h>
#include <iostream>

#include "TypeLib.h"

int main(int argc, char *argv[]) {

  TypeLib t;

  if (argc < 2) {
      std::cout << "Specify typelib" << std::endl;
      return 1;
  }

  if (!t.Open(
     // "C:\\Program Files\\Microsoft Office\\Office\\MSWORD9.olb"
     // "C:\\Program Files\\Microsoft Office\\root\\Office16\\MSWORD.OLB"
     // "C:\\Program Files\\Common Files\\microsoft shared\\OFFICE14\\MSO.DLL"
     argv[1]
  )) {
    std::cout << "Couldn't open type library" << std::endl;
    return -1;
  }

  std::cout << t.LibDocumentation() << std::endl;

  int nofTypeInfos = t.NofTypeInfos();


  std::cout << "Nof Type Infos: " << nofTypeInfos << std::endl;
  
  while (t.NextTypeInfo()) {
    std::string type_doc = t.TypeDocumentation(); 

    std::cout <<                                   std::endl;
    std::cout << type_doc                       << std::endl;
    std::cout << "----------------------------" << std::endl;

    std::cout << "  Interface: ";
    if (t.IsTypeEnum     ())   std::cout << "Enum"; 
    if (t.IsTypeRecord   ())   std::cout << "Record"; 
    if (t.IsTypeModule   ())   std::cout << "Module"; 
    if (t.IsTypeInterface())   std::cout << "Interface"; 
    if (t.IsTypeDispatch ())   std::cout << "Dispatch"; 
    if (t.IsTypeCoClass  ())   std::cout << "CoClass"; 
    if (t.IsTypeAlias    ())   std::cout << "Alias"; 
    if (t.IsTypeUnion    ())   std::cout << "Union"; 
    if (t.IsTypeMax      ())   std::cout << "Max"; 
    std::cout << std::endl;

    int nofFunctions=t.NofFunctions();
    int nofVariables=t.NofVariables();

    std::cout << "  functions: " << nofFunctions << std::endl;
    std::cout << "  variables: " << nofVariables << std::endl;

    while (t.NextFunction()) {
      std::cout << std::endl;
      std::cout << "  Function     : " << t.FunctionName()          << std::endl;

      std::cout << "    returns    : " << t.ReturnType() << std::endl;

      std::cout << "    flags      : ";
      if (t.HasFunctionTypeFlag(TypeLib::FDEFAULT      )) std::cout << "FDEFAULT "      ;
      if (t.HasFunctionTypeFlag(TypeLib::FSOURCE       )) std::cout << "FSOURCE "       ;
      if (t.HasFunctionTypeFlag(TypeLib::FRESTRICTED   )) std::cout << "FRESTRICTED "   ;
      if (t.HasFunctionTypeFlag(TypeLib::FDEFAULTVTABLE)) std::cout << "FDEFAULTVTABLE ";
      std::cout << std::endl;

      TypeLib::INVOKEKIND ik = t.InvokeKind();
      switch (ik) {
        case TypeLib::func: 
          std::cout <<"    invoke kind: function" << std::endl;
          break;
        case TypeLib::put: 
          std::cout <<"    invoke kind: put" << std::endl;
          break;
        case TypeLib::get: 
          std::cout <<"    invoke kind: get" << std::endl;
          break;
        case TypeLib::putref: 
          std::cout <<"    invoke kind: put ref" << std::endl;
          break;
        default:
          std::cout <<"    invoke kind: ???" << std::endl;
      }

      std::cout << "    params     : " << t.NofParameters()         << std::endl;
      std::cout << "    params opt : " << t.NofOptionalParameters() << std::endl;

      while (t.NextParameter()) {
        std::cout << "    Parameter  : " << t.ParameterName();
        std::cout << " type = " << t.ParameterType();
        if (t.ParameterIsIn         ()) std::cout << " in";
        if (t.ParameterIsOut        ()) std::cout << " out";
        if (t.ParameterIsFLCID      ()) std::cout << " flcid";
        if (t.ParameterIsReturnValue()) std::cout << " ret";
        if (t.ParameterIsOptional   ()) std::cout << " opt";
        if (t.ParameterHasDefault   ()) std::cout << " def";
        if (t.ParameterHasCustumData()) std::cout << " cust";
        std::cout << std::endl;
      }
    }
    while (t.NextVariable()) {
      std::cout << "  Variable : " << t.VariableName() << std::endl;
      std::cout << "      Type : " << t.VariableType();
      TypeLib::VARIABLEKIND vk = t.VariableKind();
      switch (vk) {
        case TypeLib::instance: std::cout << " (instance)" << std::endl; break;
        case TypeLib::static_ : std::cout << " (static)"   << std::endl; break;
        case TypeLib::const_  : std::cout << " (const ";
             std::cout << t.ConstValue() << ")" << std::endl;            break;
        case TypeLib::dispatch: std::cout << " (dispatch)" << std::endl; break;
        default:
          std::cout << "    variable kind: unknown" << std::endl;
      }
    }
  }

  return 0;
}
