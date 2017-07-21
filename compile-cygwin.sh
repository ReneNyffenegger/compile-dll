i686-w64-mingw32-gcc -c funcs.c
i686-w64-mingw32-gcc -shared -o funcs.dll funcs.c -Wl,--add-stdcall-alias -loleaut32 -static-libgcc
# i686-w64-mingw32-gcc -shared -o funcs.dll funcs.c -static-libgcc
