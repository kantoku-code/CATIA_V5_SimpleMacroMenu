Attribute VB_Name = "KCL"
'vba Kantoku_CATVBA_Library ver0.1.0
'KCL.bas - 標準ﾓｼﾞｭｰﾙ
Option Explicit

Private mSW& '時間計測用

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

'開発用-ｵﾌﾞｼﾞｪｸﾄﾁｪｯｸ
Sub CATMain()
    Dim msg$: msg = "選択して下さい : ESCｷｰ 終了"
    Dim SI As AnyObject
    Dim doc As Document: Set doc = CATIA.ActiveDocument
    Do
        Set SI = SelectItem(msg)
        If IsNothing(SI) Then Exit Do
        Stop
    Loop
End Sub

'*****CATIAな関数*****
'ﾏｸﾛｽﾀｰﾄﾁｪｯｸ
''' @param:DocTypes-array(string),string ﾏｸﾛ実行を許可するﾄﾞｷｭﾒﾝﾄのﾀｲﾌﾟ
''' @return:Boolean
Function CanExecute(ByVal docTypes As Variant) As Boolean
    CanExecute = False
    
    If CATIA.Windows.count < 1 Then
        MsgBox "ﾌｧｲﾙが開かれていません"
        Exit Function
    End If
    
    If VarType(docTypes) = vbString Then docTypes = Split(docTypes, ",")
    If Not IsFilterType(docTypes) Then Exit Function
    
    Dim ErrMsg As String
    ErrMsg = "ﾌｧｲﾙのﾀｲﾌﾟが異なります。" + vbNewLine + "(" + Join(docTypes, ",") + " のみです)"
    
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

'選択
''' @param:Msg-ﾒｯｾｰｼﾞ
''' @param:Filter-array(string),string 選択ﾌｨﾙﾀｰ(指定無し時AnyObject)
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

'選択
''' @param:Msg-ﾒｯｾｰｼﾞ
''' @param:Filter-array(string),string 選択ﾌｨﾙﾀｰ(指定無し時AnyObject)
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

'T型のParent取得 Nameでのﾁｪｯｸも必要
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

'BrepNameの取得
''' @param:MyBRepName-String
''' @return:String
Function GetBrepName(MyBRepName As String) As String
    MyBRepName = Replace(MyBRepName, "Selection_", "")
    MyBRepName = Left(MyBRepName, InStrRev(MyBRepName, "));"))
    MyBRepName = MyBRepName + ");WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)"
    GetBrepName = MyBRepName
End Function

'言語取得
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
            '英語-Select an object or a command
            GetLanguage = "en"
        Case ExistsKey(st, "objet")
            'フランス語-Selectionnez un objet ou une commande
            GetLanguage = "fr"
        Case ExistsKey(st, "Objekt")
            'ドイツ語-Ein Objekt oder einen Befehl auswahlen
            GetLanguage = "de"
        Case ExistsKey(st, "oggetto")
            'イタリア語-Selezionare un oggetto o un comando
            GetLanguage = "it"
        Case ExistsKey(st, "ｵﾌﾞｼﾞｪｸﾄ")
            '日本語-ｵﾌﾞｼﾞｪｸﾄまたはｺﾏﾝﾄﾞを選択してください
            GetLanguage = "ja"
        Case ExistsKey(st, "объект")
            'ロシア語-Выберите объект или команду
            GetLanguage = "ru"
        Case ExistsKey(st, "象或")
            '中国語-???象或?羅
            GetLanguage = "zh"
        Case Else
            Select Case Len(st)
                Case 13
                    '韓国語-???? ?? ?? ??　unicode未対応の為
                    GetLanguage = "ko"
                Case 23
                    '日本語-日本語版以外のため
                    GetLanguage = "ja"
                Case Else
                    'それ以外
            End Select
    End Select
End Function

'文字列内に指定文字が存在するか？
'大文字小文字は無視
Private Function ExistsKey(ByVal txt As String, ByVal key As String) As Boolean
    ExistsKey = IIf(InStr(LCase(txt), LCase(key)) > 0, True, False)
End Function

'文字型配列?
Private Function IsStringAry(ByVal ary As Variant) As Boolean
    IsStringAry = False
    
    If Not IsArray(ary) Then Exit Function
    Dim i&
    For i = 0 To UBound(ary)
        If Not VarType(ary(i)) = vbString Then Exit Function
    Next
    
    IsStringAry = True
End Function

'ﾌｨﾙﾀｰﾀｲﾌﾟとしてOK?
Private Function IsFilterType(ByVal ary As Variant) As Boolean
    IsFilterType = False
    Dim ErrMsg$: ErrMsg = "ﾌｨﾙﾀｰ又はﾄﾞｷｭﾒﾝﾄﾀｲﾌﾟの指定は" + vbNewLine + _
                          "Variant(String)型配列で行ってください" + vbNewLine + _
                          "(ﾏｸﾛｺｰﾄﾞのｴﾗｰです)"
    
    If Not IsStringAry(ary) Then
        MsgBox ErrMsg
        Exit Function
    End If
    
    IsFilterType = True
End Function

'文字型からﾊﾞﾘｱﾝﾄ配列生成(CATIAの為にすごく無駄･･･)
Private Function ToStrVriAry(ByVal s$) As Variant
    Dim ary As Variant: ary = Split(s, ",")
    Dim vriary() As Variant: ReDim vriary(UBound(ary))
    Dim i&
    For i = 0 To UBound(ary)
        vriary(i) = ary(i)
    Next
    ToStrVriAry = vriary
End Function

'*****ｼｽﾃﾑな関数*****
'Nothing 書き方に統一感が無い為
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

'型ﾁｪｯｸ
''' @param:OJ-Object
''' @param:T-String
''' @return:Boolean
Function IsType_Of_T(ByVal oj As Object, ByVal T$) As Boolean
    IsType_Of_T = IIf(TypeName(oj) = T, True, False)
End Function


'*****配列な関数*****
'配列の連結
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

'配列の抽出
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

'配列のｸﾛｰﾝ
''' @param:Ary-Variant(Of Array)
''' @return:Variant(Of Array)
Function CloneAry(ByVal ary As Variant) As Variant
    If Not IsArray(ary) Then Exit Function
    CloneAry = GetRangeAry(ary, 0, UBound(ary))
End Function

'配列の値が一致するか?
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


'*****IOな関数*****
'FileSystemObject
''' @return:Object(Of FileSystemObject)
Function GetFSO() As Object
    Set GetFSO = CreateObject("Scripting.FileSystemObject")
End Function

'ﾊﾟｽ/ﾌｧｲﾙ名/拡張子 分割
''' @param:FullPath-ﾌｧｲﾙﾊﾟｽ
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

'ﾊﾟｽ/ﾌｧｲﾙ名/拡張子 連結
''' @param:Path-Variant(Of Array(Of String)) (0-Path 1-BaseName 2-Extension)
''' @return:ﾌｧｲﾙﾊﾟｽ
Function JoinPathName$(ByVal path As Variant)
    If Not IsArray(path) Then Stop '未対応
    If Not UBound(path) = 2 Then Stop '未対応
    JoinPathName = path(0) + "\" + path(1) + "." + path(2)
End Function

'ﾌｧｲﾙ,ﾌｫﾙﾀﾞの有無
''' @param:Path-ﾊﾟｽ
''' @return:Boolean
Function IsExists(ByVal path$) As Boolean
    IsExists = False
    Dim FSO As Object: Set FSO = GetFSO
    If FSO.FileExists(path) Then
        IsExists = True: Exit Function 'ﾌｧｲﾙ
    ElseIf FSO.FolderExists(path) Then
        IsExists = True: Exit Function 'ﾌｫﾙﾀﾞ
    End If
    Set FSO = Nothing
End Function

'重複しない名前取得
''' @param:Path-ﾌｧｲﾙﾊﾟｽ
''' @return:新たなﾌｧｲﾙﾊﾟｽ
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

'ﾌｧｲﾙの書き込み
''' @param:Path-ﾌｧｲﾙﾊﾟｽ
''' @param:Txt-String
Sub WriteFile(ByVal path$, ByVal txt) '$)
    On Error Resume Next
        Call GetFSO.OpenTextFile(path, 2, True).Write(txt)
    On Error GoTo 0
End Sub

'ﾌｧｲﾙ読み込み
''' @param:Path-ﾌｧｲﾙﾊﾟｽ
''' @return:Variant(Of Array(Of String))
Function ReadFile(ByVal path$) As Variant
    On Error Resume Next
    With GetFSO.GetFile(path).OpenAsTextStream
        ReadFile = Split(.ReadAll, vbNewLine)
        .Close
    End With
    On Error GoTo 0
End Function


'*****ｽﾄｯﾌﾟｳｫｯﾁな関数*****
'時間計測スタート
Sub SW_Start()
    mSW = timeGetTime
End Sub

'計測取得
''' @return:Double(Unit:s)
Function SW_GetTime#()
    SW_GetTime = IIf(mSW = 0, -1, (timeGetTime - mSW) * 0.001)
End Function
