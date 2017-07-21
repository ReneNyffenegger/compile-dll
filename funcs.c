#include <stdint.h>
#include <string.h>
#include <stdio.h>
#include <windows.h>

BSTR BSTR_concat(BSTR a, BSTR b) {
  
 //
 // https://stackoverflow.com/a/19977944/180275
 //


    int lengthA = SysStringLen(a);
    int lengthB = SysStringLen(b);

    char buf[100];

 // Note: SysStringLen returns only half of the actual length!
    sprintf(buf, "a: %d, b:%d", lengthA, lengthB);

    MessageBox(0, buf, 0, 0);

    BSTR result = SysAllocStringLen(NULL, lengthA + lengthB);

    memcpy(result          , a, lengthA * sizeof(OLECHAR));
    memcpy(result + lengthA, b, lengthB * sizeof(OLECHAR));

//  result[lengthA + lengthB] = 0;
    return result;
}


__declspec(dllexport) int32_t __stdcall pass_integer(int32_t a, int32_t b) { return a+b; }

__declspec(dllexport) double  __stdcall pass_double (double  a, double  b) { return a+b; }

__declspec(dllexport) BSTR    __stdcall pass_string (BSTR*   a, BSTR*   b) {

  // https://stackoverflow.com/questions/39404028/passing-strings-from-vba-to-c-dll
  //
  //     If you are genuinely trying to pass a BSTR or LPCWSTR, type it as
  //     ByVal param As Long and pass StrPtr(variableName) to pass the
  //     string in - 32bit only.
  //


  return BSTR_concat(*a, *b); // It probably leaks memory
  
}



