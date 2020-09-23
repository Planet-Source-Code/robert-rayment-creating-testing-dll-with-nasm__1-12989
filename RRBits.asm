;RRBits . asm   For NASM

;VB
;Private Declare Sub RRBits Lib "RRBits" _
;(ByVal BitType As Long, ByVal BitPosNum As Long, _
;ByVal InNum As Long, ByRef Result As Long)

;BitPosNum is bit position to set,clr or test
;or number of times to rotate or shift to a max of 255

;InNum is unchanged
;Note Result is ByRef

;Const BITTEST As Long = 0   'Result = 0 bit BitPos was 0, 1 bit BitPos was 1
;Const BITSHL As Long = 1    'Result = InNum shift left BitPosNum times 0 in low bit
;Const BITSHR As Long = 2    'Result = InNum shift right BitPosNum times 0 in high bit
;Const BITROL As Long = 3    'Result = InNum rotate left BitPosNum times
;Const BITROR As Long = 4    'Result = InNum rotate right BitPosNum times
;Const BITCLR As Long = 5    'Result = InNum with bit BitPos = 0
;Const BITSET As Long = 6    'Result = InNum with bit BitPos = 1

	SEGMENT code Use32
	GLOBAL _DllMain
_DllMain:	mov eax,1
	retn 12

	GLOBAL RRBits
RRBits:
	enter 0,0
	push ebx

	mov eax,[ebp+8]	;BitType
	cmp eax,0
	je NEAR BitTEST

	cmp eax, 1
	je BitSHL

	cmp eax, 2
	je BitSHR

	cmp eax, 3
	je BitROL

	cmp eax, 4
	je BitROR

	cmp eax, 5
	je BitCLR
;-----------------------------------
	;eax=6  BitSET	
	mov edx,1
	mov ecx,[ebp+12]	;BitPosNum
	rol edx,cl		;Rotate edx BitPos times
	mov eax,[ebp+16]	;InNum
	or eax,edx	;eax is result
	jmp GETOUT
;-----------------------------------
BitCLR:
	mov edx,1
	mov ecx,[ebp+12]	;BitPosNum
	rol edx,cl		;Rotate edx BitPos times
	neg edx		;2's complement+1
	dec edx		;2's complement
	mov eax,[ebp+16]	;InNum
	and eax,edx	;eax is result
	jmp GETOUT
;-----------------------------------
BitROR:
	mov ecx,[ebp+12]	;BitPosNum
	mov eax,[ebp+16]	;InNum
	ror eax,cl		;Rotate right eax cl times
	jmp GETOUT
;-----------------------------------
BitROL:
	mov ecx,[ebp+12]	;BitPosNum
	mov eax,[ebp+16]	;InNum
	rol eax,cl		;Rotate left eax cl times
	jmp GETOUT
;-----------------------------------
BitSHR:
	mov ecx,[ebp+12]	;BitPosNum
	mov eax,[ebp+16]	;InNum
	shr eax,cl		;Shift right eax cl times
	jmp GETOUT
;-----------------------------------
BitSHL:
	mov ecx,[ebp+12]	;BitPosNum
	mov eax,[ebp+16]	;InNum
	shl eax,cl		;Shift left eax cl times
	jmp GETOUT
;-----------------------------------
BitTEST:
	mov ecx,[ebp+12]	;BitPosNum
	mov eax,[ebp+16]	;InNum
	bt eax,ecx
	mov eax,0
	jnc GETOUT	;bit not set ie 0
	inc eax		;bit set ie 1
;-----------------------------------
GETOUT:	;result in eax

	mov ebx,[ebp+20]	; ->Result
	mov [ebx],eax	;Result

	pop ebx
	leave
	retn 16
ENDS