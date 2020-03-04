// See
// http://msdn.microsoft.com/library/default.asp?url=/library/en-us/automat/htm/chap6_7zdz.asp
#include <ObjBase.h>

#include <string>

std::wstring VariantTypeAsString(const VARIANT& v);
std::wstring        TypeAsString(VARTYPE);

class variant {
  public:
    variant(VARIANT);
   ~variant();

    std::wstring ValueAsString();

    void ChangeType(VARTYPE);

  private:
    VARIANT v_;
};
