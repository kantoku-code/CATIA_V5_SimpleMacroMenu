Attribute VB_Name = "Cat_Macro_Menu_Model"
'vba CATIA V5用　ﾏｸﾛ起動(ﾃﾇｷ)ﾒﾆｭｰ Ver0.0.1  by Kantoku
'Cat_Macro_Menu_Model.bas
'using-'KCL0.0.12'
'ﾀｸﾞを追記したﾓｼﾞｭｰﾙをﾌﾟﾛｼﾞｪｸﾄに追加するだけで
'自動的にﾎﾞﾀﾝを追加します

Const FormTitle = "Macro"

'----- ﾒﾆｭｰの仕様 お好みで ---------------------------------------

'ﾒﾆｭｰの表示方法
'True-ﾓｰﾀﾞﾚｽ表示  False-ﾓｰﾀﾞﾙ表示
Private Const MENU_SHOW_TYPE = True

'ﾒﾆｭｰを閉じるﾀｲﾐﾝｸﾞ
'True-ﾎﾞﾀﾝｸﾘｯｸ後閉じます  False-ﾌｫｰﾑのXﾎﾞﾀﾝを押して閉じます
Private Const MENU_HIDE_TYPE = True

'ﾒﾆｭｰのﾏﾙﾁﾍﾟｰｼﾞの設定
'変更する際は
'{ ﾀｸﾞのｸﾞﾙｰﾌﾟ番号 : ﾍﾟｰｼﾞのﾀｲﾄﾙ文字 }
'の状態にして下さい
Private Const GROUP_NUMBER_CAPTION = _
            "{ 1 : Part }" & _
            "{11 : Assy }" & _
            "{21 : Draw }" & _
            "{51 : Other }"
'-----------------------------------------------------------------

Option Explicit

'----- 設定用 変更しない方が良いと思います -----------------------

'ﾏﾙﾁﾍﾟｰｼﾞ用 ｸﾞﾙｰﾌﾟ設定
Private PageMap As Object

'ﾀｸﾞ
Private TagMap As Object                    'ｸﾞﾙｰﾌﾟ管理用

Private Const TAG_S = "{"                   'ﾀｸﾞｽﾀｰﾄ
Private Const TAG_D = ":"                   'ﾀｸﾞ区切り
Private Const TAG_E = "}"                   'ﾀｸﾞｴﾝﾄﾞ

Private Const TAG_GROUP = "gp"              'ﾀｸﾞｸﾞﾙｰﾌﾟ
Private Const TAG_ENTRYPNT = "ep"           'ﾀｸﾞｴﾝﾄﾘｰﾎﾟｲﾝﾄ
Private Const TAG_ENTRY_DEF = "CATMain"     'ﾀｸﾞｴﾝﾄﾘｰﾎﾟｲﾝﾄ 未定義時
Private Const TAG_PJTPATH = "pjt_path"      'ﾀｸﾞﾌﾟﾛｼﾞｪｸﾄﾊﾟｽ
Private Const TAG_MDLNAME = "mdl_name"      'ﾀｸﾞﾓｼﾞｭｰﾙ名
'-----------------------------------------------------------------

'ﾒﾆｭｰ ｴﾝﾄﾘｰﾎﾟｲﾝﾄ
Sub CATMain()
    
    'ﾏﾙﾁﾍﾟｰｼﾞ用 ｸﾞﾙｰﾌﾟ設定
    Set PageMap = Get_KeyValue(GROUP_NUMBER_CAPTION, True)
    
    'ﾎﾞﾀﾝ用情報取得
    Dim ButtonInfos As Object
    Set ButtonInfos = Get_ButtonInfo()
    If ButtonInfos Is Nothing Then
        MsgBox "ﾒﾆｭｰに表示するﾏｸﾛがありません", vbExclamation + vbOKOnly
        Exit Sub
    End If
    
    'ｸﾞﾙｰﾌﾟ順にｿｰﾄ
    Dim SoLst As Object
    Set SoLst = To_SortedList(ButtonInfos)
    If SoLst Is Nothing Then Exit Sub
    
    'View表示
    Dim Menu As Cat_Macro_Menu_View
    Set Menu = New Cat_Macro_Menu_View
    Call Menu.Set_FormInfo(SoLst, PageMap, FormTitle, MENU_HIDE_TYPE)
    
    If MENU_SHOW_TYPE Then
        Menu.Show vbModeless
    Else
        Menu.Show vbModal
    End If
End Sub


'******* ｻﾎﾟｰﾄ関数 *********

'ﾓｼﾞｭｰﾙからﾎﾞﾀﾝ用情報取得
'pram  :
'return: lst(Dict)
Private Function Get_ButtonInfo() As Object
    Set Get_ButtonInfo = Nothing
    
    Dim Apc As Object: Set Apc = GetApc()
    Dim ExecPjt As Object: Set ExecPjt = Apc.ExecutingProject
    Dim PjtPath As String: PjtPath = ExecPjt.DisplayName
    
    Dim AllComps As Object
    Set AllComps = GetModuleLst(ExecPjt.ProjectItems.VBComponents)
    If AllComps Is Nothing Then Exit Function
    
    Dim Comp As Object 'VBComponent
    Dim Mdl As Object 'CodeModule
    Dim DecCode As String
    Dim DecCnt As Long
    Dim MdlInfo As Object
    Dim CanExecMethod As String
    Dim BtnInfos As Object: Set BtnInfos = KCL.InitLst()
    
    For Each Comp In AllComps
        Set Mdl = Comp.CodeModule
        
        '宣言部位置
        DecCnt = Mdl.CountOfDeclarationLines
        If DecCnt < 1 Then GoTo Continue
        
        '宣言部ｺｰﾄﾞ
        DecCode = Mdl.Lines(1, Mdl.CountOfDeclarationLines)
        
        'ﾀｸﾞ取得
        Set MdlInfo = Get_KeyValue(DecCode)
        If MdlInfo Is Nothing Then GoTo Continue
        
        'Groupﾁｪｯｸ
        If Not MdlInfo.Exists(TAG_GROUP) Then GoTo Continue
        If IsNumeric(MdlInfo(TAG_GROUP)) Then
            MdlInfo(TAG_GROUP) = CLng(MdlInfo(TAG_GROUP))
        Else
            GoTo Continue
        End If
        Debug.Print TypeName(MdlInfo(TAG_GROUP)) & " : " & MdlInfo(TAG_GROUP)
        If Not PageMap.Exists(MdlInfo(TAG_GROUP)) Then GoTo Continue
        
        'ｴﾝﾄﾘｰﾎﾟｲﾝﾄﾁｪｯｸ
        CanExecMethod = vbNullString
        If MdlInfo.Exists(TAG_ENTRYPNT) Then
            If Exist_Method(Mdl, MdlInfo(TAG_ENTRYPNT)) Then
                CanExecMethod = MdlInfo(TAG_ENTRYPNT)
            Else
                GoTo Try_TAG_ENTRY_DEF
            End If
        Else
Try_TAG_ENTRY_DEF:
            If Exist_Method(Mdl, TAG_ENTRY_DEF) Then
                 CanExecMethod = TAG_ENTRY_DEF
            End If
        End If
        If CanExecMethod = vbNullString Then GoTo Continue
        Set MdlInfo = Push_Dic(MdlInfo, TAG_ENTRYPNT, CanExecMethod)
        
        Set MdlInfo = Push_Dic(MdlInfo, TAG_PJTPATH, PjtPath)
        Set MdlInfo = Push_Dic(MdlInfo, TAG_MDLNAME, Mdl.Name)
        
        BtnInfos.Add MdlInfo
Continue:
    Next
    
    If BtnInfos.Count < 1 Then Exit Function
    
    Set Get_ButtonInfo = BtnInfos
End Function

'Dictionaryに押し込み
'pram  : Dict,vri,vri
'return: Dict
Private Function Push_Dic(ByVal Dic As Object, _
                          ByVal Key As Variant, _
                          ByVal Val As Variant) As Object
    If Dic.Exists(Key) Then
        Dic(Key) = Val
    Else
        Dic.Add Key, Val
    End If
    Set Push_Dic = Dic
End Function

'ﾀｸﾞっぽいもの取得-ｵﾌﾟｼｮﾝはKeyをLong化
'pram  : str,Opt_bool
'return: Dict
Private Function Get_KeyValue( _
                    ByVal txt As String, _
                    Optional ByVal KeyToLong As Boolean = False) _
                    As Object
    Set Get_KeyValue = Nothing

    Dim Reg As Object
    Set Reg = CreateObject("VBScript.RegExp")
    With Reg
        .Pattern = TAG_S & "(.*?)" & TAG_D & "(.*?)" & TAG_E
        .Global = True
    End With
    
    Dim Matches As Object
    Set Matches = Reg.Execute(txt)
    Set Reg = Nothing
    
    If Matches.Count < 1 Then Exit Function
    
    Dim Dic As Object: Set Dic = KCL.InitDic(vbTextCompare)
    Dim Match As Object, SubMatchs As Object
    Dim Key As Variant, Var As Variant
    
    For Each Match In Matches
        Set SubMatchs = Match.SubMatches
        
        If SubMatchs.Count < 2 Then GoTo Continue
        
        Key = Trim(Replace(SubMatchs(0), """", ""))
        If Len(Key) < 1 Then GoTo Continue
        If KeyToLong Then Key = CLng(Key)
        
        Var = Trim(Replace(SubMatchs(1), """", ""))
        If Len(Var) < 1 Then GoTo Continue
        
        Set Dic = Push_Dic(Dic, Key, Var)
Continue:
    Next
    
    If Dic.Count < 1 Then Exit Function
    
    Set Get_KeyValue = Dic
End Function

'ｸﾞﾙｰﾌﾟをｷｰとしたｿｰﾄ済みﾘｽﾄ
'pram  :lst(Dict)
'return: Dict(lst(Dict))
Private Function To_SortedList(ByVal Infos As Object) As Object
    Set To_SortedList = Nothing
    
    Dim SoLst As Object
    Set SoLst = CreateObject("System.Collections.SortedList")
    Dim Lst As Object
    
    Dim Info As Object
    For Each Info In Infos
        If SoLst.ContainsKey(Info(TAG_GROUP)) = True Then
            SoLst(Info(TAG_GROUP)).Add Info
        Else
            Set Lst = KCL.InitLst()
            Lst.Add Info
            SoLst.Add Info(TAG_GROUP), Lst
        End If
    Next
    
    If SoLst.Count < 1 Then Exit Function
    
    'ﾓｼﾞｭｰﾙ名でｿｰﾄ
    Dim i As Long
    Dim InfoDic As Object: Set InfoDic = KCL.InitDic(vbTextCompare)
    For i = 0 To SoLst.Count - 1
        InfoDic.Add SoLst.GetKey(i), Sort_by(SoLst.GetByIndex(i))
    Next
    
    Set To_SortedList = InfoDic
End Function

'ﾓｼﾞｭｰﾙ名でｿｰﾄ
'pram  :lst(Dict)
'return: lst(Dict)
Private Function Sort_by(ByVal Lst As Object) As Object
    Dim tmp As Object
    Dim i As Long, j As Long
    Set tmp = Lst(0)
    For i = 0 To Lst.Count - 1
        For j = Lst.Count - 1 To i Step -1
            If Lst(i)(TAG_MDLNAME) > Lst(j)(TAG_MDLNAME) Then
                Set tmp = Lst(i)
                Set Lst(i) = Lst(j)
                Set Lst(j) = tmp
            End If
        Next j
    Next i
    Set Sort_by = Lst
End Function


'******* APC/VBE *********

'Apc取得
'pram  :
'return: obj-IApc
Private Function GetApc() As Object
    Set GetApc = Nothing
    
    'VBAﾊﾞｰｼﾞｮﾝﾁｪｯｸ
    Dim COMObjectName$
    #If VBA7 Then
        COMObjectName = "MSAPC.Apc.7.1"
    #ElseIf VBA6 Then
        COMObjectName = "MSAPC.Apc.6.2"
    #Else
        MsgBox "VBAのﾊﾞｰｼﾞｮﾝが未対応です"
        Exit Function
    #End If
    
    'APC取得
    Dim Apc As Object: Set Apc = Nothing
    On Error Resume Next
        Set Apc = CreateObject(COMObjectName)
    On Error GoTo 0
    
    If Apc Is Nothing Then
        MsgBox "MSAPC.Apcが取得できませんでした"
        Exit Function
    End If
    
    Set GetApc = Apc
End Function

'ﾌﾟﾛｼｰｼﾞｬの存在ﾁｪｯｸ - Privateの判断していない
'pram  : obj-CodeModule,str
'return: Boolean
Private Function Exist_Method(ByVal CodeMdl As Object, _
                              ByVal Name As String) As Boolean
    Dim tmp As Long
    On Error Resume Next
        tmp = CodeMdl.ProcBodyLine(Name, 0)
    On Error GoTo 0
    Exist_Method = tmp > 0
    Err.Number = 0
End Function

'ﾓｼﾞｭｰﾙ取得
'pram  : obj-VBComponents
'return: lst(obj-VBComponent)

'vbext_ComponentType
'1-vbext_ct_StdModule 2-vbext_ct_ClassModule 3-vbext_ct_MSForm
Private Function GetModuleLst(ByVal Itms As Object) As Object
    Set GetModuleLst = Nothing
    Dim Lst As Object: Set Lst = KCL.InitLst()
    Dim Itm As Object
    For Each Itm In Itms
        If Not Itm.Type = 1 Then GoTo Continue 'vbext_ComponentType
        Lst.Add Itm
Continue:
    Next
    If Lst.Count < 1 Then Exit Function
    Set GetModuleLst = Lst
End Function
