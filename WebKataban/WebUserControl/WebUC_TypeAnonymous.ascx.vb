Imports WebKataban.DS_TankaTableAdapters
Imports WebKataban.DS_TypeAnonymousTableAdapters

Public Class WebUC_TypeAnonymous
    Inherits KHBase

#Region "コンストラクタ"

    Sub New()
        Me.BllType = New TypeBLL()
    End Sub

#End Region

#Region "プロパティ"

    ''' <summary>
    '''     機種選択ビジネスロジック
    ''' </summary>
    ''' <returns></returns>
    Private ReadOnly Property BllType As TypeBLL

#End Region

#Region "イベント"

    ''' <summary>
    '''     要素選択画面へ遷移イベント
    ''' </summary>
    Public Event GotoYouso()

#End Region

#Region "メソッド"

    ''' <summary>
    '''     初期化
    ''' </summary>
    Public Sub FrmInit()
        Try
            If TreeViewSeries.Nodes.Count > 0 Then
                TreeViewSeries.Nodes.Clear()
            End If

            Using da As New kh_series_hierarchyTableAdapter

                'ツリーメニューの設定
                Dim dtSeriesCatalog = da.GetData()
                'Root nodes
                Dim drRoots = dtSeriesCatalog.Select("upperlevel_id=0", "display_order")
                'Not root nodes
                Dim drNotRoots = dtSeriesCatalog.Select("upperlevel_id<>0", "display_order")

                '機種大分類を追加
                For Each dr As DataRow In drRoots
                    'Root
                    Dim rootNode As New TreeNode

                    'rootNode.Text = GetTextByLanguage(drCatalog, selLang.SelectedValue)
                    rootNode.Text = String.Empty
                    rootNode.Value = dr("id")
                    rootNode.ImageUrl = "~/KHImage/ckd/" & selLang.SelectedValue & "/" & dr("image_name") &
                                        ".png"

                    TreeViewSeries.Nodes.Add(rootNode)
                Next

                '各階層の追加
                For Each dr As DataRow In drNotRoots
                    Dim parentnode =
                            GetNodeById(TreeViewSeries.Nodes, dr("upperlevel_id")).FirstOrDefault()

                    If parentnode IsNot Nothing Then
                        Dim leafNode As New TreeNode

                        leafNode.Text = GetTextByLanguage(dr, selLang.SelectedValue)
                        leafNode.Value = dr("id")

                        If _
                            IsDBNull(dr("series_kataban")) OrElse
                            String.IsNullOrEmpty(dr("series_kataban").ToString()) Then
                            'LeafNodeではない場合は展開画像を表示
                            leafNode.ImageUrl = "~/KHImage/ckd/Collapsing.gif"
                            parentnode.ChildNodes.Add(leafNode)
                        Else
                            'LeafNodeの場合は、機種マスタの有効期限により表示かどうかを判断
                            If dr("in_effective_date") <= Now AndAlso dr("out_effective_date") > Now _
                                Then
                                '有効期限以内の場合は表示
                                parentnode.ChildNodes.Add(leafNode)
                            End If

                        End If
                    End If
                Next

            End Using

        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    '''     IDによりノードを取得
    ''' </summary>
    ''' <param name="nodes"></param>
    ''' <param name="id"></param>
    ''' <returns></returns>
    Private Iterator Function GetNodeById(nodes As TreeNodeCollection, id As String) As IEnumerable(Of TreeNode)
        For Each node As TreeNode In nodes
            If node.Value = id Then
                Yield node
            End If

            For Each child In GetNodeById(node.ChildNodes, id)
                Yield child
            Next
        Next
    End Function

    ''' <summary>
    '''     対象言語の表示名を取得
    ''' </summary>
    ''' <param name="dr"></param>
    ''' <param name="lang"></param>
    ''' <returns></returns>
    Private Function GetTextByLanguage(dr As DataRow, lang As String) As String
        Dim result As String = String.Empty

        Select Case selLang.SelectedValue
            Case CdCst.LanguageCd.DefaultLang
                '英語
                result = dr("name_en")

            Case CdCst.LanguageCd.Japanese
                '日本語
                result = dr("name_ja")

            Case CdCst.LanguageCd.SimplifiedChinese
                '簡体字
                result = dr("name_zh")

            Case CdCst.LanguageCd.TraditionalChinese
                '繁体字
                result = dr("name_tw")

            Case CdCst.LanguageCd.Korean
                '韓国語
                result = dr("name_ko")
        End Select
        Return result
    End Function

#End Region

#Region "イベント"

    ''' <summary>
    '''     製品リンククリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub TreeViewSeries_OnSelectedNodeChanged(sender As Object, e As EventArgs)
        If TreeViewSeries.SelectedNode.ChildNodes.Count = 0 Then

            Using daSeriesTree As New kh_series_hierarchyTableAdapter
                Dim dtSeriesTree = daSeriesTree.GetDataById(TreeViewSeries.SelectedValue)

                If dtSeriesTree.Rows.Count > 0 AndAlso Not IsDBNull(dtSeriesTree.Rows(0)("series_kataban")) Then

                    With dtSeriesTree.Rows(0)
                        Using daSeriesNmMst As New kh_series_nm_mstTableAdapter
                            Dim dtSeriesNmMst = daSeriesNmMst.GetDataByKeys(selLang.SelectedValue,
                                                                            dtSeriesTree.Rows(0)("series_kataban"),
                                                                            dtSeriesTree.Rows(0)("key_kataban"),
                                                                            Now)

                            '選択情報の設定
                            Dim strSeries As String = dtSeriesTree.Rows(0)("series_kataban")
                            Dim strKeyKataban As String = dtSeriesTree.Rows(0)("key_kataban")
                            Dim strGoodsNum As String = dtSeriesNmMst.Rows(0)("series_nm")

                            Using daSeriesKataban As New kh_series_katabanTableAdapter

                                Dim dtSeriesKataban =
                                        daSeriesKataban.GetDataBySeriesAndKey(dtSeriesTree.Rows(0)("series_kataban"),
                                                                              dtSeriesTree.Rows(0)("key_kataban"),
                                                                              Now)
                                Dim strCurrency = "JPY"
                                If dtSeriesKataban.Rows.Count > 0 Then
                                    strCurrency = dtSeriesKataban.Rows(0).Item("currency_cd")
                                End If


                                '引当シリーズ形番追加(機種)
                                Call BllType.subInsertSelSrsKtbnMdl(objCon,
                                                                    Me.objUserInfo.UserId,
                                                                    Me.objLoginInfo.SessionId,
                                                                    strSeries,
                                                                    strKeyKataban,
                                                                    strGoodsNum,
                                                                    strCurrency)
                                'ページ遷移(形番引当画面)
                                RaiseEvent GotoYouso()
                            End Using
                        End Using
                    End With
                End If
            End Using
        Else
            With TreeViewSeries.SelectedNode

                If .Expanded Then
                    '収束
                    .Collapse()
                    If Not String.IsNullOrEmpty(.Text) Then
                        'ルート以外はイメージを変更
                        .ImageUrl = "~/KHImage/ckd/Collapsing.gif"
                    End If
                Else
                    '展開
                    .Expand()
                    If Not String.IsNullOrEmpty(.Text) Then
                        'ルート以外はイメージを変更
                        .ImageUrl = "~/KHImage/ckd/Expanding.gif"

                    End If
                End If
            End With

        End If
        TreeViewSeries.SelectedNode.Selected = False
    End Sub

#End Region
End Class