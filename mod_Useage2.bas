Attribute VB_Name = "mod_Useage2"
Option Explicit
Public Enum SearchRange
  WholeMod
  DecOnly
  CodeOnly
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private WholeMod, DecOnly, CodeOnly
#End If
Public Enum CallAnalysis
  CallsZero
  CallsControlOnly
  CallsInternal
  CallsClassOnly
  CallsClassForm
  CallsClassMod
  CallsClassModForm
  CallsModonly
  CallsModMod
  CallsModModForm
  CallsModForm
  CallsModModFormForm
  CallsFormOnly
  CallsFormMod
  CallsFormForm
  CallsFormFormMod
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private CallsZero, CallsInternal, CallsClassOnly, CallsModForm, CallsModMod, CallsModonly, CallsFormForm, CallsFormOnly, CallsClassMod, CallsClassForm, CallsClassModForm
#End If

Public Function GetWholeCaseMatchCodeLine(VarProjName As Variant, _
                                          VarModName As Variant, _
                                          varFind As Variant, _
                                          strCode As String, _
                                          Optional lngStartLine As Long, _
                                          Optional SearchIn As SearchRange = WholeMod) As Boolean

  Dim Comp    As VBComponent
  Dim CompMod As CodeModule
  Dim EndLine As Long

  'Check that a Found line is still in the code
  If Len(varFind) Then
    Set Comp = VBInstance.VBProjects(VarProjName).VBComponents(VarModName)
    Set CompMod = Comp.CodeModule
    Select Case SearchIn
     Case WholeMod
      EndLine = -1
     Case DecOnly
      EndLine = CompMod.CountOfDeclarationLines + 1
      If EndLine < lngStartLine Then
        GoTo NoActionExit
      End If
     Case CodeOnly
      EndLine = -1
    End Select
    If Not Comp Is Nothing Then
      GetWholeCaseMatchCodeLine = CompMod.Find(varFind, lngStartLine, 1, EndLine, -1, True, True, False)
      If GetWholeCaseMatchCodeLine Then
        strCode = CompMod.Lines(lngStartLine, 1)
       Else
        strCode = ""
      End If
    End If
  End If
NoActionExit:

End Function

Public Function ProcedureUseageArray(strFind As String) As Variant

  Dim Proj       As VBProject
  Dim Comp       As VBComponent
  Dim StartLine  As Long
  Dim GuardLine  As Long
  Dim L_CodeLine As String
  Dim strArr     As String
  Dim UseCount   As Long
  Dim TPos       As Long

  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If Len(Comp.Name) Then
        StartLine = 1
        UseCount = 0
        Do While GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, strFind, L_CodeLine, StartLine)
          If GuardLine > 0 Then
            If GuardLine > StartLine Then
              Exit Do
            End If
          End If
          TPos = InStrWholeWordRX(L_CodeLine, strFind)
          Do While TPos
            If InCode(L_CodeLine, TPos) Then
              If Not isProcHead(L_CodeLine) Then
                UseCount = UseCount + 1
              End If
            End If
            TPos = InStrWholeWordRX(L_CodeLine, strFind, TPos + 1)
          Loop
          StartLine = StartLine + 1
          GuardLine = StartLine
        Loop
        If UseCount Then
          strArr = AccumulatorString(strArr, strInBrackets(UseCount) & Comp.Name & "[" & Comp.Type & "]")
        End If
      End If
    Next Comp
  Next Proj
  ProcedureUseageArray = Split(strArr, ",")

End Function

Public Function VariableUseageArray(strFind As String, _
                                    strOrig As String) As Variant

  Dim Proj       As VBProject
  Dim Comp       As VBComponent
  Dim StartLine  As Long
  Dim GuardLine  As Long
  Dim L_CodeLine As String
  Dim strArr     As String
  Dim UseCount   As Long
  Dim TPos       As Long

  For Each Proj In VBInstance.VBProjects
    For Each Comp In Proj.VBComponents
      If Len(Comp.Name) Then
        StartLine = 1
        UseCount = 0
        Do While GetWholeCaseMatchCodeLine(Proj.Name, Comp.Name, strFind, L_CodeLine, StartLine)
          If GuardLine > 0 Then
            If GuardLine > StartLine Then
              Exit Do
            End If
          End If
          If Not MultiLeft(strOrig, True, L_CodeLine) Then
            'because the strORig may have been rebuilt by detector to no linecont code line
            TPos = InStrWholeWordRX(L_CodeLine, strFind)
            Do While TPos
              If InCode(L_CodeLine, TPos) Then
                UseCount = UseCount + 1
              End If
              TPos = InStrWholeWordRX(L_CodeLine, strFind, TPos + 1)
            Loop
          End If
          StartLine = StartLine + 1
          GuardLine = StartLine
        Loop
        If UseCount Then
          strArr = AccumulatorString(strArr, strInBrackets(UseCount) & Comp.Name & "[" & Comp.Type & "]")
        End If
      End If
    Next Comp
  Next Proj
  VariableUseageArray = Split(strArr, ",")

End Function

':)Code Fixer V2.9.6 (9/02/2005 3:08:47 AM) 30 + 126 = 156 Lines Thanks Ulli for inspiration and lots of code.

