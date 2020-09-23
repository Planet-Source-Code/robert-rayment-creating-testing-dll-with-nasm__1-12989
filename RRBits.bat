c:\nasm\nasmw -f coff RRBits.asm -l List.txt -E Err.txt
Link.exe /dll /export:RRBits /entry:DllMain RRBits.o
del RRBits.exp
del RRBits.lib
del RRBits.o