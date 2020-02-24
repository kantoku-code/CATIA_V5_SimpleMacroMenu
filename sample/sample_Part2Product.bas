Attribute VB_Name = "sample_Part2Product"
'vba sample_Part2Product_ver0.0.3  using-'KCL0.0.12'  by Kantoku
'PartからProduct化します
'ﾎﾞﾃﾞｨ･形状ｾｯﾄ･時系列形状ｾｯﾄで表示されているもののみ
'(ｽｹｯﾁは対象外)

'{GP:1}
'{Caption:Pt→Pro}
'{ControlTipText:PartからProductに変換します}

Option Explicit

Sub CATMain()
    'ﾄﾞｷｭﾒﾝﾄのﾁｪｯｸ
    If Not CanExecute("PartDocument") Then Exit Sub
    
    'Part
    Dim BaseDoc As PartDocument: Set BaseDoc = CATIA.ActiveDocument
    
    Dim BasePath As Variant: BasePath = Array(BaseDoc.FullName)
    
    Dim Pt As Part: Set Pt = BaseDoc.Part
    
    Dim LeafItems As Collection: Set LeafItems = Get_LeafItemLst(Pt.Bodies)
    
    Dim Msg As String
    If LeafItems Is Nothing Then
        Msg = "ｺﾋﾟｰ可能な要素がありません!"
        MsgBox Msg, vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    '確認
    Msg = LeafItems.Count & "個のｺﾋﾟｰ可能な要素があります。" & vbNewLine & _
          "ﾍﾟｰｽﾄするﾀｲﾌﾟを指定してください" & vbNewLine & vbNewLine & _
          "はい : ﾘﾝｸの結果として(As Result With Link)" & vbNewLine & _
          "いいえ : 結果として(As Result)" & vbNewLine & _
          "キャンセル : マクロ中止"
    Dim PasteType As String
    Select Case MsgBox(Msg, vbQuestion + vbYesNoCancel)
        Case vbYes
            PasteType = "CATPrtResult"
        Case vbNo
            PasteType = "CATPrtResultWithOutLink"
        Case Else
            Exit Sub
    End Select
    
    KCL.SW_Start
    
    Dim BaseScene As Variant: BaseScene = GetScene3D(GetViewPnt3D())
    
    'Assy
    Dim TopDoc As ProductDocument: Set TopDoc = CATIA.Documents.Add("Product")
    
    Call ToProduct(TopDoc, LeafItems, PasteType)
    
    Call UpdateScene(BaseScene)
    TopDoc.Product.Update
    
    Debug.Print "Time:" & KCL.SW_GetTime & "s"
    MsgBox "Done"
End Sub

'Productにﾍﾟｰｽﾄ
Private Sub ToProduct(ByVal TopDoc As ProductDocument, _
                      ByVal LeafItems As Collection, _
                      ByVal PasteType As String)
    Dim TopSel As Selection
    Set TopSel = TopDoc.Selection
    
    Dim BaseSel As Selection
    Set BaseSel = KCL.GetParent_Of_T(LeafItems(1), "PartDocument").Selection
    
    Dim Prods As Products
    Set Prods = TopDoc.Product.Products
    
    Dim Itm As AnyObject
    Dim TgtDoc As PartDocument
    Dim ProdsNameDic As Object: Set ProdsNameDic = KCL.InitDic()
    
    CATIA.HSOSynchronized = False
    
    For Each Itm In LeafItems
        If ProdsNameDic.Exists(Itm.Name) Then
            Set TgtDoc = ProdsNameDic.Item(Itm.Name)
        Else
            Set TgtDoc = Init_Part(Prods, Itm.Name)
            ProdsNameDic.Add Itm.Name, TgtDoc
        End If
        
        Call Preparing_Copy(BaseSel, Itm)
        With BaseSel
            .Copy
            .Clear
        End With
        
        With TopSel
            .Clear
            .Add TgtDoc.Part
            .PasteSpecial PasteType
        End With
    Next
    BaseSel.Clear
    TopSel.Clear
    CATIA.HSOSynchronized = True
End Sub

'ｺﾋﾟｰ要素を選択状態にする
Private Sub Preparing_Copy(ByVal Sel As Selection, ByVal Itm As AnyObject)
    Sel.Clear
    
    'Body
    If TypeName(Itm) = "Body" Then
        Sel.Add Itm
        Exit Sub
    End If
    
    'HybridBody
    Dim ShpsLst As Collection: Set ShpsLst = New Collection
    ShpsLst.Add Itm.HybridShapes
    
    Select Case TypeName(Itm)
        Case "HybridBody"
            Set ShpsLst = Get_All_HbShapes(Itm, ShpsLst)
        Case "OrderedGeometricalSet"
            Set ShpsLst = Get_All_OdrGeoSetShapes(Itm, ShpsLst)
    End Select
    
    Dim Shps As HybridShapes, Shp As HybridShape
    For Each Shps In ShpsLst
        For Each Shp In Shps
            Sel.Add Shp
        Next
    Next
End Sub

'指定時系列形状ｾｯﾄ以下の全てHybridShapesを取得
'再帰　かなり効率悪いかも
Private Function Get_All_OdrGeoSetShapes(ByVal OdrGeoSet As OrderedGeometricalSet, _
                                         ByVal Lst As Collection) As Collection
    Dim Child As OrderedGeometricalSet
    For Each Child In OdrGeoSet.OrderedGeometricalSets
        Lst.Add Child.HybridShapes
        If Child.OrderedGeometricalSets.Count > 0 Then
            Set Lst = Get_All_OdrGeoSetShapes(Child, Lst)
        End If
    Next
    Set Get_All_OdrGeoSetShapes = Lst
End Function

'指定形状ｾｯﾄ以下の全てHybridShapesを取得
'再帰　かなり効率悪いかも
Private Function Get_All_HbShapes(ByVal Hbdy As HybridBody, _
                                  ByVal Lst As Collection) As Collection
    Dim Child As HybridBody
    For Each Child In Hbdy.hybridBodies
        Lst.Add Child.HybridShapes
        If Child.hybridBodies.Count > 0 Then
            Set Lst = Get_All_HbShapes(Child, Lst)
        End If
    Next
    Set Get_All_HbShapes = Lst
End Function

'ﾙｰﾄで、空ではなく表示されているﾎﾞﾃﾞｨ・形状ｾｯﾄ取得
Private Function Get_LeafItemLst(ByVal Pt As Part) As Collection
    Set Get_LeafItemLst = Nothing
    
    Dim Sel As Selection: Set Sel = Pt.Parent.Selection
    Dim TmpLst As Collection: Set TmpLst = New Collection
    Dim i As Long

    Dim Filter As String
    Filter = "(CATPrtSearch.BodyFeature.Visibility=Shown " & _
            "+ CATPrtSearch.OpenBodyFeature.Visibility=Shown" & _
            "+ CATPrtSearch.MMOrderedGeometricalSet.Visibility=Shown),sel"
    
    CATIA.HSOSynchronized = False
    
    With Sel
        .Clear
        .Add Pt
        .Search Filter
        For i = 1 To .Count2
            TmpLst.Add .Item(i).Value
        Next
        .Clear
    End With
    
    CATIA.HSOSynchronized = True
    
    If TmpLst.Count < 1 Then Exit Function
    
    Dim LeafHBdys As Object: Set LeafHBdys = KCL.InitDic()
    Dim Hbdy As AnyObject 'HybridBody & OrderedGeometricalSets
    For Each Hbdy In Pt.hybridBodies
        LeafHBdys.Add Hbdy, 0
    Next
    For Each Hbdy In Pt.OrderedGeometricalSets
        LeafHBdys.Add Hbdy, 0
    Next
    
    
    Dim Itm As AnyObject
    Dim Lst As Collection: Set Lst = New Collection
    For Each Itm In TmpLst
        Select Case TypeName(Itm)
            Case "Body"
                If Is_LeafBody(Itm) Then Lst.Add Itm
            Case Else 'HybridBody & OrderedGeometricalSets
                If Is_LeafHybridBody(Itm, LeafHBdys) Then Lst.Add Itm
        End Select
    Next
    If Lst.Count < 1 Then Exit Function
    
    Set Get_LeafItemLst = Lst
End Function

'ﾙｰﾄのﾎﾞﾃﾞｨで中身が有るか?
Private Function Is_LeafBody(ByVal Bdy As Body) As Boolean
    Is_LeafBody = Bdy.InBooleanOperation = False And Bdy.Shapes.Count > 0
End Function

'ﾙｰﾄの形状ｾｯﾄで中身が有り可視か?As HybridBody
'prm:Hbdy - HybridBody & OrderedGeometricalSets
Private Function Is_LeafHybridBody(ByVal Hbdy As AnyObject, _
                                   ByVal Dic As Object) As Boolean
    Is_LeafHybridBody = False
    If Not Dic.Exists(Hbdy) Then Exit Function
    
    CATIA.HSOSynchronized = False
    
    Dim Sel As Selection
    Set Sel = KCL.GetParent_Of_T(Hbdy, "PartDocument").Selection
    
    Dim Cnt As Long
    
    With Sel
        .Clear
        .Add Hbdy
        .Search "Visibility=Shown,sel"
        Cnt = .Count2
        .Clear
    End With
    
    CATIA.HSOSynchronized = True
    
    If Cnt > 1 Then Is_LeafHybridBody = True
End Function

'Part新作
Private Function Init_Part(ByVal Prods As Variant, _
                           ByVal PtNum As String) As PartDocument
    Dim Prod As Product
    On Error Resume Next
        Set Prod = Prods.AddNewComponent("Part", PtNum)
    On Error GoTo 0
    Set Init_Part = Prods.Item(Prods.Count).ReferenceProduct.Parent
End Function

'*** camera ***
'表示のUpdate
Private Sub UpdateScene(ByVal Scene As Variant)
    Dim Viewer As Viewer3D: Set Viewer = CATIA.ActiveWindow.ActiveViewer
    Dim VPnt3D As Variant 'Viewpoint3D '
    Set VPnt3D = Viewer.Viewpoint3D
    
    Dim ary As Variant
    ary = GetRangeAry(Scene, 0, 2)
    Call VPnt3D.PutOrigin(ary)
    
    ary = GetRangeAry(Scene, 3, 5)
    Call VPnt3D.PutSightDirection(ary)
    
    ary = GetRangeAry(Scene, 6, 8)
    Call VPnt3D.PutUpDirection(ary)
    
    VPnt3D.FieldOfView = Scene(9)
    VPnt3D.FocusDistance = Scene(10)
    
    Call Viewer.Update
End Sub

'Viewpoint3Dからｼｰﾝ取得
Private Function GetScene3D(ViewPnt3D As Viewpoint3D) As Variant
    Dim vp As Variant: Set vp = ViewPnt3D
    
    Dim origin(2) As Variant: Call vp.GetOrigin(origin)
    
    Dim sight(2) As Variant: Call vp.GetSightDirection(sight)
    GetScene3D = KCL.JoinAry(origin, sight)
    
    Dim up(2) As Variant: Call vp.GetUpDirection(up)
    GetScene3D = KCL.JoinAry(GetScene3D, up)
    
    Dim FieldOfView(0) As Variant: FieldOfView(0) = vp.FieldOfView
    GetScene3D = KCL.JoinAry(GetScene3D, FieldOfView)
    
    Dim FocusDist(0) As Variant: FocusDist(0) = vp.FocusDistance
    GetScene3D = KCL.JoinAry(GetScene3D, FocusDist)
End Function

'現状の視点取得
Private Function GetViewPnt3D() As Viewpoint3D
    Set GetViewPnt3D = CATIA.ActiveWindow.ActiveViewer.Viewpoint3D
End Function

