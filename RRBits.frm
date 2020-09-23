VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'RRBits by Robert Rayment

'Creating & testing RRBits.DLL

'Using NASM. Netwide Assembler freeware from
'www.phoenix.gb/net
'www.cryogen.com/Nasm
'& other sites
'see also
'www.geocities.com/SunsetStrip/Stage/8513/assembly.html

'NASM needs to be installed before you can create a DLL
'To compile run RRBits.bat
'Error message DlMain not standard - can ignore
'Paths in RRBat will need adjusting to your own
'also a Path to link.exe needs putting in Autoexec.bat

Private Declare Sub RRBits Lib "RRBits" _
(ByVal BitType As Long, ByVal BitPosNum As Long, _
ByVal InNum As Long, ByRef Result As Long)

'BitPosNum is bit position to set, clr or test
'or number of times to rotate or shift to a max of 32

'InNum is unchanged
'Note Result is ByRef

'BitType:-
Const BITTEST As Long = 0   'Result = 0 bit BitPos was 0, 1 bit BitPos was 1
Const BITSHL As Long = 1    'Result = InNum shift left BitPosNum times 0 in low bit
Const BITSHR As Long = 2    'Result = InNum shift right BitPosNum times 0 in high bit
Const BITROL As Long = 3    'Result = InNum rotate left BitPosNum times
Const BITROR As Long = 4    'Result = InNum rotate right BitPosNum times
Const BITCLR As Long = 5    'Result = InNum with bit BitPos = 0
Const BITSET As Long = 6    'Result = InNum with bit BitPos = 1

Private Sub Form_Load()
Dim BitType As Long
Dim BitPosNum As Long
Dim InNum As Long
Dim Result As Long

Print "Testing RRBits.DLL"
Print

BitType = BITTEST
BitPosNum = 0
InNum = 4
RRBits BitType, BitPosNum, InNum, Result
Print "BITTEST"; " Number="; InNum; " Bit pos="; BitPosNum; " Result="; Result
Print

BitType = BITTEST
BitPosNum = 2
InNum = 4
RRBits BitType, BitPosNum, InNum, Result
Print "BITTEST"; " Number="; InNum; " Bit pos="; BitPosNum; " Result="; Result
Print

BitType = BITSHL
BitPosNum = 2
InNum = 4
RRBits BitType, BitPosNum, InNum, Result
Print "BITSHL"; " Number="; InNum; " Shift num="; BitPosNum; " Result="; Result
Print

BitType = BITSHR
BitPosNum = 1
InNum = 4
RRBits BitType, BitPosNum, InNum, Result
Print "BITSHR"; " Number="; InNum; " Shift num="; BitPosNum; " Result="; Result
Print

BitType = BITROL
BitPosNum = 1
InNum = &H80000001
RRBits BitType, BitPosNum, InNum, Result
Print "BITROL"; " Number="; Hex$(InNum); " Roll num="; BitPosNum; " Result="; Hex$(Result)
Print

BitType = BITROR
BitPosNum = 1
InNum = &H80000001
RRBits BitType, BitPosNum, InNum, Result
Print "BITROR"; " Number="; Hex$(InNum); " Roll num="; BitPosNum; " Result="; Hex$(Result)
Print

BitType = BITCLR
BitPosNum = 0
InNum = &H80000001
RRBits BitType, BitPosNum, InNum, Result
Print "BITCLR"; " Number="; Hex$(InNum); " Bit pos="; BitPosNum; " Result="; Hex$(Result)
Print

BitType = BITSET
BitPosNum = 0
InNum = &H80000000
RRBits BitType, BitPosNum, InNum, Result
Print "BITSET"; " Number="; Hex$(InNum); " Bit pos="; BitPosNum; " Result="; Hex$(Result)
Print

End Sub

Sub RRBit(ByVal RBitType As Long, ByVal RBitPosNum As Long, ByVal RInNum As Long, ByRef RResult As Long)
RRBits RBitType, RBitPosNum, RInNum, RResult
End Sub

Private Sub Form_Resize()
Width = 4320
Height = 4800
Left = 2000
Top = 2000
End Sub
