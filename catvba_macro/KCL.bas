Attribute VB_Name = "KCL"
'vba Kantoku_CATVBA_Library ver0.1.0
'KCL.bas - �W��Ӽޭ��
Option Explicit

Private mSW& '���Ԍv���p

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

'�J���p-��޼ު������
Sub CATMain()
    Dim msg$: msg = "�I�����ĉ����� : ESC�� �I��"
    Dim SI As AnyObject
    Dim doc As Document: Set doc = CATIA.ActiveDocument
    Do
        Set SI = SelectItem(msg)
        If IsNothing(SI) Then Exit Do
        Stop
    Loop
End Sub

'*****CATIA�Ȋ֐�*****
'ϸ۽�������
''' @param:DocTypes-array(string),string ϸێ��s���������޷���Ă�����
''' @return:Boolean
Function CanExecute(ByVal docTypes As Variant) As Boolean
    CanExecute = False
    
    If CATIA.Windows.count < 1 Then
        MsgBox "̧�ق��J����Ă��܂���"
        Exit Function
    End If
    
    If VarType(docTypes) = vbString Then docTypes = Split(docTypes, ",")
    If Not IsFilterType(docTypes) Then Exit Function
    
    Dim ErrMsg As String
    ErrMsg = "̧�ق����߂��قȂ�܂��B" + vbNewLine + "(" + Join(docTypes, ",") + " �݂̂ł�)"
    
    Dim ActDoc As Document
    On Error Resume Next
        Set ActDoc = CATIA.ActiveDocument
    On Error GoTo 0
    If ActDoc Is Nothing Then
        MsgBox ErrMsg, vbExclamation + vbOKOnly
        Exit Function
    End If
    
    If UBound(filter(docTypes, TypeName(ActDoc))) < 0 Then
        MsgBox ErrMsg, vbExclamation + vbOKOnly
        Exit Function
    End If
    
    CanExecute = True
End Function

'�I��
''' @param:Msg-ү����
''' @param:Filter-array(string),string �I��̨���(�w�薳����AnyObject)
''' @return:AnyObject
Function SelectItem(ByVal msg$, _
                           Optional ByVal filter As Variant = Empty) _
                           As AnyObject
    Dim SE As SelectedElement
    Set SE = SelectElement(msg, filter)
    
    If IsNothing(SE) Then
        Set SelectItem = SE
    Else
        Set SelectItem = SE.Value
    End If
End Function

'�I��
''' @param:Msg-ү����
''' @param:Filter-array(string),string �I��̨���(�w�薳����AnyObject)
''' @return:SelectedElement
Function SelectElement(ByVal msg$, _
                           Optional ByVal filter As Variant = Empty) _
                           As SelectedElement
    If IsEmpty(filter) Then filter = Array("AnyObject")
    If VarType(filter) = vbString Then filter = ToStrVriAry(filter)
    If Not IsFilterType(filter) Then Exit Function
    
    Dim sel As Variant: Set sel = CATIA.ActiveDocument.selection
    sel.Clear
    Select Case sel.SelectElement2(filter, msg, False)
        Case "Cancel", "Undo", "Redo"
            Exit Function
    End Select
    Set SelectElement = sel.Item(1)
    sel.Clear
End Function

'InternalName
''' @param:AOj-AnyObject
''' @return:String
Function GetInternalName$(ByVal aoj As AnyObject)
    If IsNothing(aoj) Then
        GetInternalName = Empty: Exit Function
    End If
    GetInternalName = aoj.GetItem("ModelElement").InternalName
End Function

'T�^��Parent�擾 Name�ł��������K�v
''' @param:AOj-AnyObject
''' @param:T-String
''' @return:AnyObject
Function GetParent_Of_T( _
    ByVal aoj As AnyObject, _
    ByVal T As String) _
    As AnyObject
    
    
    Dim aojName As String
    Dim parentName As String
    
    On Error Resume Next
        Set aoj = asDisp(aoj)
        aojName = aoj.name
        parentName = aoj.Parent.name
    On Error GoTo 0

    If TypeName(aoj) = TypeName(aoj.Parent) And _
       aojName = parentName Then
        Set GetParent_Of_T = Nothing
        Exit Function
    End If
    If TypeName(aoj) = T Then
        Set GetParent_Of_T = aoj
    Else
        Set GetParent_Of_T = GetParent_Of_T(aoj.Parent, T)
    End If
End Function

Private Function asDisp(o As INFITF.CATBaseDispatch) As INFITF.CATBaseDispatch
    Set asDisp = o
End Function

'BrepName�̎擾
''' @param:MyBRepName-String
''' @return:String
Function GetBrepName(MyBRepName As String) As String
    MyBRepName = Replace(MyBRepName, "Selection_", "")
    MyBRepName = Left(MyBRepName, InStrRev(MyBRepName, "));"))
    MyBRepName = MyBRepName + ");WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)"
    GetBrepName = MyBRepName
End Function

'����擾
'return-ISO 639-1 code
'https://ja.wikipedia.org/wiki/ISO_639-1%E3%82%B3%E3%83%BC%E3%83%89%E4%B8%80%E8%A6%A7
Function GetLanguage() As String
    GetLanguage = "non"
    If CATIA.Windows.count < 1 Then Exit Function
    GetLanguage = "other"
    CATIA.ActiveDocument.selection.Clear
    Dim st As String: st = CATIA.StatusBar
    Select Case True
        Case ExistsKey(st, "object")
            '�p��-Select an object or a command
            GetLanguage = "en"
        Case ExistsKey(st, "objet")
            '�t�����X��-Selectionnez un objet ou une commande
            GetLanguage = "fr"
        Case ExistsKey(st, "Objekt")
            '�h�C�c��-Ein Objekt oder einen Befehl auswahlen
            GetLanguage = "de"
        Case ExistsKey(st, "oggetto")
            '�C�^���A��-Selezionare un oggetto o un comando
            GetLanguage = "it"
        Case ExistsKey(st, "��޼ު��")
            '���{��-��޼ު�Ă܂��ͺ���ނ�I�����Ă�������
            GetLanguage = "ja"
        Case ExistsKey(st, "���q���u�{��")
            '���V�A��-�B���q�u���y���u ���q���u�{�� �y�|�y �{���}�p�~�t��
            GetLanguage = "ru"
        Case ExistsKey(st, "�ۈ�")
            '������-???�ۈ�?��
            GetLanguage = "zh"
        Case Else
            Select Case Len(st)
                Case 13
                    '�؍���-???? ?? ?? ??�@unicode���Ή��̈�
                    GetLanguage = "ko"
                Case 23
                    '���{��-���{��ňȊO�̂���
                    GetLanguage = "ja"
                Case Else
                    '����ȊO
            End Select
    End Select
End Function

'��������Ɏw�蕶�������݂��邩�H
'�啶���������͖���
Private Function ExistsKey(ByVal txt As String, ByVal key As String) As Boolean
    ExistsKey = IIf(InStr(LCase(txt), LCase(key)) > 0, True, False)
End Function

'�����^�z��?
Private Function IsStringAry(ByVal ary As Variant) As Boolean
    IsStringAry = False
    
    If Not IsArray(ary) Then Exit Function
    Dim i&
    For i = 0 To UBound(ary)
        If Not VarType(ary(i)) = vbString Then Exit Function
    Next
    
    IsStringAry = True
End Function

'̨������߂Ƃ���OK?
Private Function IsFilterType(ByVal ary As Variant) As Boolean
    IsFilterType = False
    Dim ErrMsg$: ErrMsg = "̨��������޷�������߂̎w���" + vbNewLine + _
                          "Variant(String)�^�z��ōs���Ă�������" + vbNewLine + _
                          "(ϸۺ��ނ̴װ�ł�)"
    
    If Not IsStringAry(ary) Then
        MsgBox ErrMsg
        Exit Function
    End If
    
    IsFilterType = True
End Function

'�����^������ر�Ĕz�񐶐�(CATIA�ׂ̈ɂ��������ʥ��)
Private Function ToStrVriAry(ByVal s$) As Variant
    Dim ary As Variant: ary = Split(s, ",")
    Dim vriary() As Variant: ReDim vriary(UBound(ary))
    Dim i&
    For i = 0 To UBound(ary)
        vriary(i) = ary(i)
    Next
    ToStrVriAry = vriary
End Function

'*****���тȊ֐�*****
'Nothing �������ɓ��ꊴ��������
''' @param:OJ-Variant(Of Object)
''' @return:Boolean
Function IsNothing(ByVal oj As Variant) As Boolean
    IsNothing = oj Is Nothing
End Function

'Scripting.Dictionary
''' @param:CompareMode-Long
''' @return:Object(Of Dictionary)
Function InitDic(Optional CompareMode As Long = vbBinaryCompare) As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = CompareMode
    Set InitDic = dic
End Function

'ArrayList
''' @return:Object(Of ArrayList)Public
Function InitLst() As Object
    Set InitLst = CreateObject("System.Collections.ArrayList")
End Function

'�^����
''' @param:OJ-Object
''' @param:T-String
''' @return:Boolean
Function IsType_Of_T(ByVal oj As Object, ByVal T$) As Boolean
    IsType_Of_T = IIf(TypeName(oj) = T, True, False)
End Function


'*****�z��Ȋ֐�*****
'�z��̘A��
''' @param:Ary1-Variant(Of Array)
''' @param:Ary2-Variant(Of Array)
''' @return:Variant(Of Array)
Function JoinAry(ByVal Ary1 As Variant, ByVal Ary2 As Variant)
    Select Case True
        Case Not IsArray(Ary1) And Not IsArray(Ary2)
            JoinAry = Empty: Exit Function
        Case Not IsArray(Ary1)
            JoinAry = Ary2: Exit Function
        Case Not IsArray(Ary2)
            JoinAry = Ary1: Exit Function
    End Select
    Dim StCount&: StCount = UBound(Ary1)
    ReDim Preserve Ary1(UBound(Ary1) + UBound(Ary2) + 1)
    Dim i&
    If IsObject(Ary2(0)) Then
        For i = StCount + 1 To UBound(Ary1)
            Set Ary1(i) = Ary2(i - StCount - 1)
        Next
    Else
        For i = StCount + 1 To UBound(Ary1)
            Ary1(i) = Ary2(i - StCount - 1)
        Next
    End If
    JoinAry = Ary1
End Function

'�z��̒��o
''' @param:Ary-Variant(Of Array)
''' @param:StartIdx-Long
''' @param:EndIdx-Long
''' @return:Variant(Of Array)
Function GetRangeAry(ByVal ary As Variant, ByVal StartIdx&, ByVal EndIdx&) As Variant
    If Not IsArray(ary) Then Exit Function
    If EndIdx - StartIdx < 0 Then Exit Function
    If StartIdx < 0 Then Exit Function
    If EndIdx > UBound(ary) Then Exit Function
    
    Dim RngAry() As Variant: ReDim RngAry(EndIdx - StartIdx)
    Dim i&
    For i = StartIdx To EndIdx
        RngAry(i - StartIdx) = ary(i)
    Next
    GetRangeAry = RngAry
End Function

'�z��̸۰�
''' @param:Ary-Variant(Of Array)
''' @return:Variant(Of Array)
Function CloneAry(ByVal ary As Variant) As Variant
    If Not IsArray(ary) Then Exit Function
    CloneAry = GetRangeAry(ary, 0, UBound(ary))
End Function

'�z��̒l����v���邩?
''' @param:Ary1-Variant(Of Array)
''' @param:Ary2-Variant(Of Array)
''' @return:Boolean
Function IsAryEqual(ByVal Ary1 As Variant, ByVal Ary2 As Variant) As Boolean
    IsAryEqual = False
    If Not IsArray(Ary1) Or Not IsArray(Ary2) Then Exit Function
    If Not UBound(Ary1) = UBound(Ary2) Then Exit Function
    Dim i&
    For i = 0 To UBound(Ary1)
        If Not Ary1(i) = Ary2(i) Then Exit Function
    Next
    IsAryEqual = True
End Function


'*****IO�Ȋ֐�*****
'FileSystemObject
''' @return:Object(Of FileSystemObject)
Function GetFSO() As Object
    Set GetFSO = CreateObject("Scripting.FileSystemObject")
End Function

'�߽/̧�ٖ�/�g���q ����
''' @param:FullPath-̧���߽
''' @return:Variant(Of Array(Of String)) (0-Path 1-BaseName 2-Extension)
Function SplitPathName(ByVal FullPath$) As Variant
    Dim path(2) As String
    With GetFSO
        path(0) = .GetParentFolderName(FullPath)
        path(1) = .GetBaseName(FullPath)
        path(2) = .GetExtensionName(FullPath)
    End With
    SplitPathName = path
End Function

'�߽/̧�ٖ�/�g���q �A��
''' @param:Path-Variant(Of Array(Of String)) (0-Path 1-BaseName 2-Extension)
''' @return:̧���߽
Function JoinPathName$(ByVal path As Variant)
    If Not IsArray(path) Then Stop '���Ή�
    If Not UBound(path) = 2 Then Stop '���Ή�
    JoinPathName = path(0) + "\" + path(1) + "." + path(2)
End Function

'̧��,̫��ނ̗L��
''' @param:Path-�߽
''' @return:Boolean
Function IsExists(ByVal path$) As Boolean
    IsExists = False
    Dim FSO As Object: Set FSO = GetFSO
    If FSO.FileExists(path) Then
        IsExists = True: Exit Function '̧��
    ElseIf FSO.FolderExists(path) Then
        IsExists = True: Exit Function '̫���
    End If
    Set FSO = Nothing
End Function

'�d�����Ȃ����O�擾
''' @param:Path-̧���߽
''' @return:�V����̧���߽
Function GetNewName$(ByVal oldPath$)
    Dim path As Variant
    path = SplitPathName(oldPath)
    path(2) = "." & path(2)
    Dim newPath$: newPath = path(0) + "\" + path(1)
    If Not IsExists(newPath + path(2)) Then
        GetNewName = newPath + path(2)
        Exit Function
    End If
    Dim TempName$, i&: i = 0
    Do
        i = i + 1
        TempName = newPath + "_" + CStr(i) + path(2)
        If Not IsExists(TempName) Then
            GetNewName = TempName
            Exit Function
        End If
    Loop
End Function

'̧�ق̏�������
''' @param:Path-̧���߽
''' @param:Txt-String
Sub WriteFile(ByVal path$, ByVal txt) '$)
    On Error Resume Next
        Call GetFSO.OpenTextFile(path, 2, True).Write(txt)
    On Error GoTo 0
End Sub

'̧�ٓǂݍ���
''' @param:Path-̧���߽
''' @return:Variant(Of Array(Of String))
Function ReadFile(ByVal path$) As Variant
    On Error Resume Next
    With GetFSO.GetFile(path).OpenAsTextStream
        ReadFile = Split(.ReadAll, vbNewLine)
        .Close
    End With
    On Error GoTo 0
End Function


'*****�į�߳����Ȋ֐�*****
'���Ԍv���X�^�[�g
Sub SW_Start()
    mSW = timeGetTime
End Sub

'�v���擾
''' @return:Double(Unit:s)
Function SW_GetTime#()
    SW_GetTime = IIf(mSW = 0, -1, (timeGetTime - mSW) * 0.001)
End Function
