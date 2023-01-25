VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LogicSymbolsUpdated 
   Caption         =   "UserForm1"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4365
   OleObjectBlob   =   "LogicSymbolsUpdated.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LogicSymbolsUpdated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Const NOT_unicode As Integer = 172
Public Const X_unicode As Integer = 215
Public Const LAMBDA_unicode As Integer = 955
Public Const IMPLY_unicode As Integer = 8594
Public Const IFF_unicode As Integer = 8596
Public Const EXISTS_unicode As Integer = 8707
Public Const ALL_unicode As Integer = 8704
Public Const AND_unicode As Integer = 8743
Public Const OR_unicode As Integer = 8744
Public Const XOR_unicode As Integer = 8853
Public Const INTERSECTION_unicode As Integer = 8745
Public Const UNION_unicode As Integer = 8746
Public Const EMPTY_SET_unicode As Integer = 8709
Public Const NOT_EQUAL_unicode As Integer = 8800
Public Const EQUIVALENCE_unicode As Integer = 8801
Public Const GREATER_OR_EQUAL_unicode As Integer = 8805
Public Const LESS_OR_EQUAL_unicode As Integer = 8804
Public Const ELEMENT_OF_unicode As Integer = 8712
Public Const NOT_ELEMENT_OF_unicode As Integer = 8713
Public Const NOT_SUBSET_OR_EQUAL_unicode As Integer = 8840
Public Const SUPERSET_unicode As Integer = 8835
Public Const SUBSET_OR_EQUAL_unicode As Integer = 8838
Public Const SUPERSET_OR_EQUAL_unicode As Integer = 8839
Public Const LEFT_CEILING_unicode As Integer = 8968
Public Const RIGHT_CEILING_unicode As Integer = 8969

Private Sub Insert_Character(unicode_symbol_number)
Selection.InsertSymbol CharacterNumber:=unicode_symbol_number, Font:="Cambria Math", Unicode:=True, Bias:=0
Word.Application.Activate
End Sub

' LOGIC SYMBOLS

Private Sub CommandButton1_Click()
Insert_Character(NOT_unicode)
End Sub

Private Sub CommandButton3_Click()
Insert_Character(OR_unicode)
End Sub

Private Sub CommandButton4_Click()
Insert_Character(AND_unicode)
End Sub

Private Sub CommandButton5_Click()
Insert_Character(XOR_unicode)
End Sub

Private Sub CommandButton6_Click()
Insert_Character(NOT_EQUAL_unicode)
End Sub

Private Sub CommandButton7_Click()
Insert_Character(IMPLY_unicode)
End Sub

Private Sub CommandButton8_Click()
Insert_Character(IFF_unicode)
End Sub

Private Sub CommandButton9_Click()
Insert_Character(EQUIVALENCE_unicode)
End Sub

Private Sub CommandButton10_Click()
Insert_Character(EXISTS_unicode)
End Sub

Private Sub CommandButton11_Click()
Insert_Character(ALL_unicode)
End Sub

Private Sub CommandButton12_Click()
Insert_Character(GREATER_OR_EQUAL_unicode)
End Sub

Private Sub CommandButton13_Click()
Insert_Character(LESS_OR_EQUAL_unicode)
End Sub

' SET SYMBOLS

Private Sub CommandButton2_Click()
Insert_Character(UNION_unicode)
End Sub

Private Sub CommandButton14_Click()
Insert_Character(ELEMENT_OF_unicode)
End Sub

Private Sub CommandButton15_Click()
Insert_Character(NOT_ELEMENT_OF_unicode)
End Sub

Private Sub CommandButton16_Click()
Insert_Character(NOT_SUBSET_OR_EQUAL_unicode)
End Sub

Private Sub CommandButton17_Click()
Insert_Character(SUPERSET_unicode)
End Sub

Private Sub CommandButton18_Click()
Insert_Character(SUBSET_OR_EQUAL_unicode)
End Sub

Private Sub CommandButton19_Click()
Insert_Character(SUPERSET_OR_EQUAL_unicode)
End Sub

Private Sub CommandButton20_Click()
Insert_Character(EMPTY_SET_unicode)
End Sub

Private Sub CommandButton21_Click()
Insert_Character(INTERSECTION_unicode)
End Sub

Private Sub CommandButton22_Click()
Insert_Character(X_unicode)
End Sub

' COUNTING SYMBOLS

Private Sub CommandButton23_Click()
Insert_Character(LEFT_CEILING_unicode)
Insert_Character(RIGHT_CEILING_unicode)
End Sub

Private Sub CommandButton24_Click()
Insert_Character(LAMBDA_unicode)
End Sub

' CAN REMOVE?

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Click()

End Sub
