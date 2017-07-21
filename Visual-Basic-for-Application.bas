'
'   Call the dll with Visual Basic for Applications (VBA)
'
'   This code must be placed into a VBA-module. Otherwise, it doesn't
'   seem to run.
'
'   The dll must be found via the %PATH% environment variable:
'     c:\path\to\compile-dll> PATH=%CD%;%PATH%
'     c>\path\to\compile-dll> call-dll.xlsm
'

option explicit

'
' VBA Integer is 4 bytes.
'
declare function pass_integer lib "funcs.dll" (byVal a as integer, byVal b as integer) as integer
declare function pass_double  lib "funcs.dll" (byVal a as double , byVal b as double ) as double
declare function pass_string  lib "funcs.dll" (byRef a as string , byRef b as string ) as string

sub main()
  
  msgBox (pass_integer(20   , 22   ))
  msgBox (pass_double (20.20, 22.22))

  dim a as string
  dim b as string

  a="12345"
  b="abcde"
  msgBox (">" & pass_string (a, b) & "<")

' msgBox (">" & pass_string ("ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz") & "<")

end sub
