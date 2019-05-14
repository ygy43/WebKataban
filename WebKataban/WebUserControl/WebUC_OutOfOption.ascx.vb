Imports System.Data.SqlClient
Imports WebKataban.ClsCommon

Public Class WebUC_OutOfOption
    Inherits KHBase

#Region "プロパティ"
    Public Event BacktoYouso()
    Private objOutOp As KHOutOfOptionCstm
#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        Call ClearHid()
        Me.OnLoad(Nothing)
        Call SetAllFontName(Me)
        'Me.HidMessage.Value = ClsCommon.fncGetMsg( selLang.SelectedValue, "W1002")
    End Sub

    ''' <summary>
    ''' ロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.Visible Then Exit Sub
        If Me.objUserInfo.UserId Is Nothing Then Exit Sub
        Try
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
            Me.lblSeriesNm.Text = objKtbnStrc.strcSelection.strGoodsNm

            'ページ初期設定
            Call subInitPage()
            'オプション外指定クラスインスタンス作成
            objOutOp = New KHOutOfOptionCstm(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                      selLang.SelectedValue, objKtbnStrc.strcSelection.strSeriesKataban, _
                                      objKtbnStrc.strcSelection.strKeyKataban)

            'オプション外指定情報取得
            Call objOutOp.subOutOpInfoGet(objCon, objKtbnStrc.strcSelection.strOpSymbol)
            '画面設定
            Call subListMake(objOutOp)

            Call KHLabelCtl.subSetLabel(objCon, CdCst.PgmId.KHOutOFOption, selLang.SelectedValue, Me) 'Label取得
            'Label設定
            Dim strConvert As String = String.Empty
            'タイロッド延長寸法ラジオ変換部分
            If objKtbnStrc.strcSelection.strSeriesKataban.Equals("JSC3") _
            AndAlso objKtbnStrc.strcSelection.strKeyKataban.Equals("1") Then
                strConvert = "MM"
            Else
                strConvert = "MX"
            End If
            Call ReplaceLabel(pnlTieRod, strConvert)
            '↓RM1401080 2014/01/29
            If objKtbnStrc.strcSelection.strSeriesKataban.Equals("SCS2") Then
                Me.Label14.Text = String.Empty
            End If
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' クリア
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearHid()
        Me.HdnActionType.Value = String.Empty
        Me.HdnPortPlace1.Value = String.Empty
        Me.HdnPortPlace2.Value = String.Empty
        Me.HdnPortPlace3.Value = String.Empty
        Me.HdnPortPlace4.Value = String.Empty
        Me.HdnPtnCnt.Value = String.Empty
        Me.HdnSelClevis.Value = String.Empty
        Me.HdnSelcmbTieRodCstm.Value = String.Empty
        Me.HdnSelFluoroRub.Value = String.Empty
        Me.HdnSelJM.Value = String.Empty
        Me.HdnSelMounting.Value = String.Empty
        Me.HdnSelPort.Value = String.Empty
        Me.HdnSelPortCushon.Value = String.Empty
        Me.HdnSelPortPlace.Value = String.Empty
        Me.HdnSelPortSize.Value = String.Empty
        Me.HdnSelSUS.Value = String.Empty
        Me.HdnSelTieRod.Value = String.Empty
        Me.HdnSelTieRodDefault.Value = String.Empty
        Me.HdnSelTrunnion.Value = String.Empty
        Me.HdnSeltxtTieRodCstm.Value = String.Empty
        Me.HdnTieRodRdio.Value = String.Empty
    End Sub

    ''' <summary>
    ''' ページ初期設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subInitPage()
        Try
            'ラジオボタンの選択イベント追加
            rdoRK1.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoRK1');")
            rdoRK2.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoRK2');")
            rdoRK3.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoRK3');")
            rdoRK4.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoRK4');")
            rdoRC1.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoRC1');")
            rdoRC2.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoRC2');")
            rdoRC3.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoRC3');")
            rdoRC4.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoRC4');")
            rdoHK1.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoHK1');")
            rdoHK2.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoHK2');")
            rdoHK3.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoHK3');")
            rdoHK4.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoHK4');")
            rdoHC1.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoHC1');")
            rdoHC2.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoHC2');")
            rdoHC3.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoHC3');")
            rdoHC4.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_PortChk('" & strParent & Me.ID & "','rdoHC4');")
            rdoTie.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_TieRodChk('" & strParent & Me.ID & "','1');")
            rdoTieR.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_TieRodChk('" & strParent & Me.ID & "','2');")
            rdoTieR1.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_TieRodChk('" & strParent & Me.ID & "','3');")
            rdoTieR2.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_TieRodChk('" & strParent & Me.ID & "','4');")
            rdoTieH.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_TieRodChk('" & strParent & Me.ID & "','5');")
            rdoTieH1.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_TieRodChk('" & strParent & Me.ID & "','6');")
            rdoTieH2.Attributes.Add(CdCst.JavaScript.OnClick, "f_OutOfOption_TieRodChk('" & strParent & Me.ID & "','7');")
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' オプション外特注画面を作成する
    ''' </summary>
    ''' <param name="objOutOp">オプション外特注詳細情報</param>
    ''' <remarks></remarks>
    Private Sub subListMake(ByVal objOutOp As KHOutOfOptionCstm)

        Dim strPlaceRK As Integer = 0
        Dim strPlaceRC As Integer = 0
        Dim strPlaceHK As Integer = 0
        Dim strPlaceHC As Integer = 0
        Try
            'ポート位置ラジオの選択情報を取得
            If objOutOp.SelPortCuPlace.Length > 0 Then
                strPlaceRK = Mid(objOutOp.SelPortCuPlace, 1, 1)
                strPlaceRC = Mid(objOutOp.SelPortCuPlace, 2, 1)
                strPlaceHK = Mid(objOutOp.SelPortCuPlace, 3, 1)
                strPlaceHC = Mid(objOutOp.SelPortCuPlace, 4, 1)
            End If

            '対象ラジオボタンリスト作成
            Dim hasRdoRK As New Hashtable
            hasRdoRK.Add("1", "rdoRK1")
            hasRdoRK.Add("2", "rdoRK2")
            hasRdoRK.Add("3", "rdoRK3")
            hasRdoRK.Add("4", "rdoRK4")
            Dim hasRdoRC As New Hashtable
            hasRdoRC.Add("1", "rdoRC1")
            hasRdoRC.Add("2", "rdoRC2")
            hasRdoRC.Add("3", "rdoRC3")
            hasRdoRC.Add("4", "rdoRC4")
            Dim hasRdoHK As New Hashtable
            hasRdoHK.Add("1", "rdoHK1")
            hasRdoHK.Add("2", "rdoHK2")
            hasRdoHK.Add("3", "rdoHK3")
            hasRdoHK.Add("4", "rdoHK4")
            Dim hasRdoHC As New Hashtable
            hasRdoHC.Add("1", "rdoHC1")
            hasRdoHC.Add("2", "rdoHC2")
            hasRdoHC.Add("3", "rdoHC3")
            hasRdoHC.Add("4", "rdoHC4")

            'ポート・クッションニードル位置
            Select Case objOutOp.intPortCushion
                Case -1          '非表示
                    Me.pnlPortCushon.Visible = False
                Case 0           '使用不可
                    Me.pnlPortCushon.Visible = True
                    Me.cmbPortCushon.Enabled = False
                    Me.cmbPortCushon.DataSource = objOutOp.PortCushion
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbPortCushon.DataTextField = "ITEM1"
                    Me.cmbPortCushon.DataValueField = "ITEM2"
                    Me.cmbPortCushon.DataBind()
                    Me.cmbPortCushon.SelectedIndex = objOutOp.SelPortCushion

                    'ポート位置ラジオの編集
                    subEditRdo(hasRdoRK, False, strPlaceRK)
                    subEditRdo(hasRdoRC, False, strPlaceRC)
                    subEditRdo(hasRdoHK, False, strPlaceHK)
                    subEditRdo(hasRdoHC, False, strPlaceHC)

                    'ポート位置テキスト
                    Me.txtPortCuchon.Text = objOutOp.SelPortCuPlace
                Case 1           '使用可
                    Me.pnlPortCushon.Visible = True
                    Me.cmbPortCushon.Enabled = True
                    Me.cmbPortCushon.DataSource = objOutOp.PortCushion
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbPortCushon.DataTextField = "ITEM1"
                    Me.cmbPortCushon.DataValueField = "ITEM2"
                    Me.cmbPortCushon.DataBind()
                    Me.cmbPortCushon.SelectedIndex = objOutOp.SelPortCushion

                    'ポート位置ラジオの編集
                    subEditRdo(hasRdoRK, True, strPlaceRK)
                    subEditRdo(hasRdoRC, True, strPlaceRC)
                    subEditRdo(hasRdoHK, True, strPlaceHK)
                    subEditRdo(hasRdoHC, True, strPlaceHC)

                    'ポート位置テキスト
                    Me.txtPortCuchon.Text = objOutOp.SelPortCuPlace
            End Select

            'ポート２箇所
            Select Case objOutOp.isPort
                Case -1          '非表示
                    Me.pnlPort.Visible = False
                Case 0           '使用不可
                    Me.pnlPort.Visible = True
                    Me.cmbPort.Enabled = False
                    Me.cmbPort.DataSource = objOutOp.Port
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbPort.DataTextField = "ITEM1"
                    Me.cmbPort.DataValueField = "ITEM2"
                    Me.cmbPort.DataBind()
                    Me.cmbPort.SelectedIndex = objOutOp.SelPort
                Case 1           '使用可
                    Me.pnlPort.Visible = True
                    Me.cmbPort.Enabled = True
                    Me.cmbPort.DataSource = objOutOp.Port
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbPort.DataTextField = "ITEM1"
                    Me.cmbPort.DataValueField = "ITEM2"
                    Me.cmbPort.DataBind()
                    Me.cmbPort.SelectedIndex = objOutOp.SelPort
            End Select

            'ポートサイズダウン
            Select Case objOutOp.isPortSize
                Case -1          '非表示
                    Me.pnlPortSize.Visible = False
                Case 0           '使用不可
                    Me.pnlPortSize.Visible = True
                    Me.cmbPortSize.Enabled = False
                    Me.cmbPortSize.DataSource = objOutOp.PortSize
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbPortSize.DataTextField = "ITEM1"
                    Me.cmbPortSize.DataValueField = "ITEM2"
                    Me.cmbPortSize.DataBind()
                    Me.cmbPortSize.SelectedIndex = objOutOp.SelPortSize
                Case 1           '使用可
                    Me.pnlPortSize.Visible = True
                    Me.cmbPortSize.Enabled = True
                    Me.cmbPortSize.DataSource = objOutOp.PortSize
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbPortSize.DataTextField = "ITEM1"
                    Me.cmbPortSize.DataValueField = "ITEM2"
                    Me.cmbPortSize.DataBind()
                    Me.cmbPortSize.SelectedIndex = objOutOp.SelPortSize
            End Select

            '支持金具回転
            Select Case objOutOp.isMounting
                Case -1          '非表示
                    Me.pnlMounting.Visible = False
                Case 0           '使用不可
                    Me.pnlMounting.Visible = True
                    Me.cmbMounting.Enabled = False
                    Me.cmbMounting.DataSource = objOutOp.Mounting
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbMounting.DataTextField = "ITEM1"
                    Me.cmbMounting.DataValueField = "ITEM2"
                    Me.cmbMounting.DataBind()
                    Me.cmbMounting.SelectedIndex = objOutOp.SelMounting
                Case 1           '使用可
                    Me.pnlMounting.Visible = True
                    Me.cmbMounting.Enabled = True
                    Me.cmbMounting.DataSource = objOutOp.Mounting
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbMounting.DataTextField = "ITEM1"
                    Me.cmbMounting.DataValueField = "ITEM2"
                    Me.cmbMounting.DataBind()
                    Me.cmbMounting.SelectedIndex = objOutOp.SelMounting
            End Select

            'トラニオン位置指定
            Select Case objOutOp.isTrunnion
                Case -1          '非表示
                    Me.pnlTrunnion.Visible = False
                Case 0           '使用不可
                    Me.pnlTrunnion.Visible = True
                    Me.txtTrunnion.Enabled = False
                    Me.txtTrunnion.CssClass = ""
                    Me.txtTrunnion.BackColor = Drawing.Color.LightGray
                    Me.txtTrunnion.Text = objOutOp.SelTrunnion
                Case 1           '使用可
                    Me.pnlTrunnion.Visible = True
                    Me.txtTrunnion.Enabled = True
                    Me.txtTrunnion.CssClass = "textBox"
                    Me.txtTrunnion.Text = objOutOp.SelTrunnion
            End Select

            '二山ナックル・二山クレビス
            Select Case objOutOp.isClevis
                Case -1          '非表示
                    Me.pnlClevis.Visible = False
                Case 0           '使用不可
                    Me.pnlClevis.Visible = True
                    Me.cmbClevis.Enabled = False
                    Me.cmbClevis.DataSource = objOutOp.Clevis
                    Me.cmbClevis.DataTextField = "ITEM1"
                    Me.cmbClevis.DataValueField = "ITEM2"
                    Me.cmbClevis.DataBind()
                    Me.cmbClevis.SelectedIndex = objOutOp.SelClevis
                Case 1           '使用可
                    Me.pnlClevis.Visible = True
                    Me.cmbClevis.Enabled = True
                    Me.cmbClevis.DataSource = objOutOp.Clevis
                    Me.cmbClevis.DataTextField = "ITEM1"
                    Me.cmbClevis.DataValueField = "ITEM2"
                    Me.cmbClevis.DataBind()
                    Me.cmbClevis.SelectedIndex = objOutOp.SelClevis
            End Select

            'タイロッド延長寸法
            '対象ラジオボタンリスト作成
            Dim hasRdo As New Hashtable
            hasRdo.Add("1", "rdoTie")
            hasRdo.Add("2", "rdoTieR")
            hasRdo.Add("3", "rdoTieR1")
            hasRdo.Add("4", "rdoTieR2")
            hasRdo.Add("5", "rdoTieH")
            hasRdo.Add("6", "rdoTieH1")
            hasRdo.Add("7", "rdoTieH2")

            '選択情報を取得
            Dim intSelTieRod As Integer = 0
            If objOutOp.SelTieRodRadio.Length > 0 Then
                intSelTieRod = objOutOp.SelTieRodRadio
                Me.HdnTieRodRdio.Value = objOutOp.SelTieRodRadio
            End If
            Dim intSelTieCus As Integer = 0
            If objOutOp.SelTieRodCstm.Length > 0 Then
                intSelTieCus = objOutOp.SelTieRodCstm
            End If

            Select Case objOutOp.isTieRod
                Case -1          '非表示
                    Me.pnlTieRod.Visible = False
                Case 0           '使用不可
                    Me.pnlTieRod.Visible = True
                    'タイロッド延長寸法テーブル
                    For i As Integer = 0 To objOutOp.TieRodRadio.Count - 1
                        'タイロッド延長寸法テーブルの表示列
                        Me.tblTieRod.Rows(CInt(objOutOp.TieRodRadio(i))).Visible = True
                    Next

                    'タイロッド延長寸法ラジオ編集
                    subEditRdo(hasRdo, False, intSelTieRod)

                    'タイロッド標準
                    Me.lblDefault.Text = objOutOp.TieRodDefl

                    'タイロッド特注リストの場合
                    If objOutOp.TieRodCstm.Count = 0 OrElse _
                    (objOutOp.TieRodCstm.Count > 0 AndAlso _
                        objOutOp.TieRodCstm.Item(0) <> "") Then
                        Me.txtTieRodCstm.Visible = False
                        Me.cmbTieRodCstm.Visible = True
                        Me.cmbTieRodCstm.Enabled = False
                    Else
                        Me.txtTieRodCstm.Visible = True
                        Me.txtTieRodCstm.Enabled = False
                        Me.txtTieRodCstm.BackColor = Drawing.ColorTranslator.FromHtml("#FFFFCC")
                        Me.txtTieRodCstm.Text = objOutOp.SelTieRodCstm
                        Me.cmbTieRodCstm.Visible = False
                    End If
                Case 1           '使用可
                    Me.pnlTieRod.Visible = True
                    'タイロッド延長寸法テーブル
                    For i As Integer = 0 To objOutOp.TieRodRadio.Count - 1
                        'タイロッド延長寸法テーブルの表示列
                        Me.tblTieRod.Rows(CInt(objOutOp.TieRodRadio(i))).Visible = True
                    Next

                    'タイロッド延長寸法ラジオ編集
                    subEditRdo(hasRdo, True, intSelTieRod)

                    'タイロッド標準
                    Me.lblDefault.Text = objOutOp.TieRodDefl
                    'タイロッド特注リストの場合
                    If objOutOp.TieRodCstm.Count = 0 OrElse _
                    (objOutOp.TieRodCstm.Count > 0 AndAlso _
                        objOutOp.TieRodCstm.Item(0) <> "") Then
                        Me.txtTieRodCstm.Visible = False
                        Me.cmbTieRodCstm.Visible = True
                        Me.cmbTieRodCstm.Enabled = True
                        Me.cmbTieRodCstm.DataSource = objOutOp.TieRodCstm
                        Me.cmbTieRodCstm.DataBind()
                        Me.cmbTieRodCstm.SelectedIndex = intSelTieCus
                    Else
                        Me.txtTieRodCstm.Visible = True
                        Me.txtTieRodCstm.Enabled = True
                        Me.txtTieRodCstm.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFFFCC")
                        Me.txtTieRodCstm.Text = objOutOp.SelTieRodCstm
                        Me.cmbTieRodCstm.Visible = False
                    End If
            End Select

            'タイロッド材質SUS
            Select Case objOutOp.isSUS
                Case -1          '非表示
                    Me.pnlSUS.Visible = False
                Case 0           '使用不可
                    Me.pnlSUS.Visible = True
                    Me.cmbSUS.Enabled = False
                    Me.cmbSUS.DataSource = objOutOp.SUS
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbSUS.DataTextField = "ITEM1"
                    Me.cmbSUS.DataValueField = "ITEM2"
                    Me.cmbSUS.DataBind()
                    Me.cmbSUS.SelectedIndex = objOutOp.SelSUS
                Case 1           '使用可
                    Me.pnlSUS.Visible = True
                    Me.cmbSUS.Enabled = True
                    Me.cmbSUS.DataSource = objOutOp.SUS
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbSUS.DataTextField = "ITEM1"
                    Me.cmbSUS.DataValueField = "ITEM2"
                    Me.cmbSUS.DataBind()
                    Me.cmbSUS.SelectedIndex = objOutOp.SelSUS
            End Select

            'ピストンロッドはジャバラ付寸法でジャバラなし
            Select Case objOutOp.isJM
                Case -1          '非表示
                    Me.pnlJM.Visible = False
                Case 0           '使用不可
                    Me.pnlJM.Visible = True
                    Me.cmbJM.Enabled = False
                    Me.cmbJM.DataSource = objOutOp.JM
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbJM.DataTextField = "ITEM1"
                    Me.cmbJM.DataValueField = "ITEM2"
                    Me.cmbJM.DataBind()
                    Me.cmbJM.SelectedIndex = objOutOp.SelJM
                Case 1           '使用可
                    Me.pnlJM.Visible = True
                    Me.cmbJM.Enabled = True
                    Me.cmbJM.DataSource = objOutOp.JM
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbJM.DataTextField = "ITEM1"
                    Me.cmbJM.DataValueField = "ITEM2"
                    Me.cmbJM.DataBind()
                    Me.cmbJM.SelectedIndex = objOutOp.SelJM
            End Select

            'スクレーバー、ロッドパッキンのみフッ素ゴム
            Select Case objOutOp.isFluoroRub
                Case -1          '非表示
                    Me.pnlFluoroRub.Visible = False
                Case 0           '使用不可
                    Me.pnlFluoroRub.Visible = True
                    Me.cmbFluoroRub.Enabled = False
                    Me.cmbFluoroRub.DataSource = objOutOp.FluoroRub
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbFluoroRub.DataTextField = "ITEM1"
                    Me.cmbFluoroRub.DataValueField = "ITEM2"
                    Me.cmbFluoroRub.DataBind()
                    Me.cmbFluoroRub.SelectedIndex = objOutOp.SelFluoroRub
                Case 1           '使用可
                    Me.pnlFluoroRub.Visible = True
                    Me.cmbFluoroRub.Enabled = True
                    Me.cmbFluoroRub.DataSource = objOutOp.FluoroRub
                    'データテーブルに変更したことにより追加  2017/04/06 追加
                    Me.cmbFluoroRub.DataTextField = "ITEM1"
                    Me.cmbFluoroRub.DataValueField = "ITEM2"
                    Me.cmbFluoroRub.DataBind()
                    Me.cmbFluoroRub.SelectedIndex = objOutOp.SelFluoroRub
            End Select
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' ラベルの入れ替え
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="strConvert"></param>
    ''' <remarks></remarks>
    Private Sub ReplaceLabel(obj As Object, strConvert As String)
        For inti As Integer = 0 To obj.Controls.Count - 1
            If obj.Controls(inti).Controls.count > 0 Then
                ReplaceLabel(obj.Controls(inti), strConvert)
            Else
                Select Case obj.Controls(inti).GetType.Name.ToUpper
                    Case "LABEL"
                        obj.Controls(inti).text = Replace(obj.Controls(inti).text, "@KEY", strConvert)
                End Select
            End If
        Next
    End Sub

    ''' <summary>
    ''' タイロッド延長寸法位置ラジオ編集
    ''' </summary>
    ''' <param name="tblRadio">対象コントロール一覧</param>
    ''' <param name="isEnable">使用可否フラグ</param>
    ''' <param name="intChk">チェックNo</param>
    ''' <remarks></remarks>
    Private Sub subEditRdo(ByVal tblRadio As Hashtable, ByVal isEnable As Boolean, ByVal intChk As Integer)
        Dim rdo As New RadioButton
        Try
            '対象コントロール一覧分ループ
            For i As Integer = 1 To tblRadio.Count
                '対象コントロールを取得
                rdo = CType(Me.FindControl(tblRadio(CStr(i))), RadioButton)
                '使用可否設定
                rdo.Enabled = isEnable
                'チェック
                If i = intChk Then
                    rdo.Checked = True
                Else
                    rdo.Checked = False
                End If
            Next
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' キャンセル
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        '引当情報削除
        Call subDeleteSelOutOp()
        Call ClearHid()
        RaiseEvent BacktoYouso()
    End Sub

    ''' <summary>
    ''' OKボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        'オプション外指定クラスインスタンス作成
        objOutOp = New KHOutOfOptionCstm(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                  selLang.SelectedValue, _
                                  objKtbnStrc.strcSelection.strSeriesKataban, _
                                  objKtbnStrc.strcSelection.strKeyKataban)
        Call objOutOp.subOutOpInfoGet(objCon, objKtbnStrc.strcSelection.strOpSymbol)

        Dim strControlNM As String = String.Empty
        Dim strMsg As String = String.Empty
        Dim strMsgOption As String = String.Empty
        '入力チェック
        If Not fncInputCheck(objOutOp, strControlNM, strMsg, strMsgOption) Then
            'エラーメッセージ出力
            If strMsg.Length > 0 Then
                AlertMessage(strMsg, strControlNM)
            Else
                AlertMessage("E001", strMsgOption)
            End If
            Exit Sub
        End If

        '選択情報を保存
        If fncSetInfo(objOutOp) Then
            '引当情報更新
            Call objOutOp.subUpdateSelOutOp(objCon, objKtbnStrc)
        End If
        RaiseEvent BacktoYouso()
    End Sub

    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <param name="obj">オプション外特注クラス   </param>
    ''' <param name="strControlNM"></param>
    ''' <param name="strMsg"></param>
    ''' <param name="strMsgOption"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncInputCheck(ByVal obj As KHOutOfOptionCstm, _
                                    ByRef strControlNM As String, ByRef strMsg As String, _
                                    ByRef strMsgOption As String) As Boolean
        fncInputCheck = False
        Try
            'ポート・クッションニードル位置指定
            If Me.HdnSelPortCushon.Value = "1" Then
                'R指定時
                '位置指定は、4項目すべての指定が必要
                If Me.HdnSelPortPlace.Value.Length <> 4 Then
                    strControlNM = "cmbPortCushon"
                    strMsg = "W0910"
                    Exit Function
                End If

                '支持金具の併用不可
                If Me.HdnSelPortPlace.Value <> "1212" AndAlso Me.HdnSelMounting.Value > 0 Then
                    strControlNM = "cmbMounting"
                    strMsg = "W0870"
                    Exit Function
                End If
                'ポート２個所
                Dim intChk(1) As Integer
                Select Case Me.HdnSelPort.Value
                    Case "1"
                        'E指定の場合、ロッド指定、ヘッド指定ともにチェック対象
                        intChk(0) = 1
                        intChk(1) = 3
                    Case "2"
                        'E1指定の場合、ロッド指定のみチェック対象
                        intChk(0) = 1
                    Case "3"
                        'E2指定の場合、ヘッド指定のみチェック対象
                        intChk(0) = 3
                End Select
                'チェック
                For i As Integer = 0 To 1
                    'チェック終了確認
                    If intChk(i) = 0 Then
                        Exit For
                    End If
                    '以下の指定の場合、エラー
                    Select Case Mid(Me.HdnSelPortPlace.Value, intChk(i), 2)
                        Case "11", "13", "31", "33", "22", "24", "42", "44"
                            strControlNM = "cmbPort"
                            strMsg = "W0930"
                            Exit Function
                    End Select

                Next
            End If
            'トラニオン位置指定
            ''半角チェック
            If Not obj.fncHankakuCheck(Me.HdnSelTrunnion.Value, strMsg) Then
                strControlNM = "txtTrunnion"
                strMsgOption = KHLabelCtl.fncSelectLabelById(objCon, "KHOutOFOptionMain", selLang.SelectedValue, CdCst.Lbl.Division.Label, 24)
                Exit Function
            End If
            ''数値チェック
            If Not obj.fncNumericCheck(Me.HdnSelTrunnion.Value, strMsg) Then
                strControlNM = "txtTrunnion"
                strMsgOption = KHLabelCtl.fncSelectLabelById(objCon, "KHOutOFOptionMain", selLang.SelectedValue, CdCst.Lbl.Division.Label, 24)
                Exit Function
            End If

            '特注寸法
            Dim strTieRodCstm As String = ""
            If Me.HdnSeltxtTieRodCstm.Value.Length > 0 Then
                '特注寸法テキストより設定
                strControlNM = "txtTieRodCstm"
                strTieRodCstm = Me.HdnSeltxtTieRodCstm.Value

            ElseIf Me.HdnSelcmbTieRodCstm.Value.Length > 0 Then
                '特注寸法コンボより設定
                strControlNM = "cmbTieRodCstm"
                strTieRodCstm = Me.HdnSeltxtTieRodCstm.Value

            End If

            '特注寸法に設定がある場合
            If strTieRodCstm.Length > 0 Then
                ''半角チェック
                If Not obj.fncHankakuCheck(strTieRodCstm, strMsg) Then
                    strMsgOption = KHLabelCtl.fncSelectLabelById(objCon, "KHOutOFOptionMain", selLang.SelectedValue, CdCst.Lbl.Division.Label, 45)
                    Exit Function
                End If
                ''数値チェック
                If Not obj.fncNumericCheck(strTieRodCstm, strMsg) Then
                    strMsgOption = KHLabelCtl.fncSelectLabelById(objCon, "KHOutOFOptionMain", selLang.SelectedValue, CdCst.Lbl.Division.Label, 45)
                    Exit Function
                End If

                'タイロッド寸法ラジオ選択有りの場合
                If Me.HdnSelTieRod.Value.Length > 0 Then

                    'タイロッド寸法最小チェック
                    If CDec(Me.HdnSelTieRodDefault.Value) > CDec(strTieRodCstm) Then
                        strMsg = "W0880"
                        Exit Function
                    End If

                    '標準寸法毎に最大値を算出
                    Dim decMax As Decimal = 0
                    Select Case Me.HdnSelTieRodDefault.Value
                        'Case "20"
                        '    decMax = 118
                        'Case "23"
                        '    decMax = 120
                        'Case "26"
                        '    decMax = 123
                        'Case "27"
                        '    decMax = 124
                        'Case "32"
                        '    decMax = 128
                        Case "11"
                            decMax = 118
                        Case "13"
                            decMax = 120
                        Case "15"
                            decMax = 123
                        Case "16"
                            decMax = 124
                        Case "19"
                            decMax = 128
                    End Select

                    'タイロッド寸法最大チェック
                    If decMax < CDec(strTieRodCstm) Then
                        strMsg = "W0880"
                        Exit Function
                    End If
                End If
                'クリア
                strControlNM = ""
            End If
            '正常
            fncInputCheck = True
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    ''' 選択情報の保存
    ''' </summary>
    ''' <param name="obj">オプション外特注クラス</param>
    ''' <returns>True:変更あり,False:変更なし</returns>
    ''' <remarks></remarks>
    Private Function fncSetInfo(ByVal obj As KHOutOfOptionCstm) As Boolean

        Dim dalOutOfOption As New OutOfOptionDAL   '2017/04/14 追加
        Dim dt As New DataTable                    '2017/04/14 追加

        fncSetInfo = False

        'セッション呼び出し用に宣言  2017/04/11 Upd Matsubara
        Dim httpCon As System.Web.HttpContext = System.Web.HttpContext.Current

        '初期化処理  2017/04/10 追加
        ReDim objKtbnStrc.strcSelection.strOutofOpCountryDiv(12)
        Dim intCnt As Integer = 0  '2017/04/07 追加

        Try
            'オプション外特注クラス.選択情報に設定
            'ポートクッションニードル
            If Me.HdnSelPortCushon.Value <> "0" OrElse Me.HdnSelPortCushon.Value <> obj.SelPortCushion Then
                fncSetInfo = True
                obj.SelPortCushion = CInt(Me.HdnSelPortCushon.Value)
                '変数に値を入れる 2017/04/07 追加
                objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = CInt(Me.HdnValPortCushon.Value)
                intCnt = intCnt + 1
            End If
            'ポート位置
            If Me.HdnSelPortPlace.Value <> "" OrElse Me.HdnSelPortPlace.Value <> obj.SelPortCuPlace Then
                fncSetInfo = True
                obj.SelPortCuPlace = Me.HdnSelPortPlace.Value
            End If
            'ポート二箇所
            If Me.HdnSelPort.Value <> "0" OrElse Me.HdnSelPort.Value <> obj.SelPort Then
                fncSetInfo = True
                obj.SelPort = CInt(Me.HdnSelPort.Value)
                '変数に値を入れる 2017/04/07 追加
                objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = CInt(Me.HdnValPort.Value)
                intCnt = intCnt + 1
            End If
            'ポートサイズ
            If Me.HdnSelPortSize.Value <> "0" OrElse Me.HdnSelPortSize.Value <> obj.SelPortSize Then
                fncSetInfo = True
                obj.SelPortSize = CInt(Me.HdnSelPortSize.Value)
                '変数に値を入れる 2017/04/07 追加
                objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = CInt(Me.HdnValPortSize.Value)
                intCnt = intCnt + 1
            End If
            '支持金具回転
            If Me.HdnSelMounting.Value <> "0" OrElse Me.HdnSelMounting.Value <> obj.SelMounting Then
                fncSetInfo = True
                obj.SelMounting = CInt(Me.HdnSelMounting.Value)
                '変数に値を入れる 2017/04/07 追加
                objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = CInt(Me.HdnValMounting.Value)
                intCnt = intCnt + 1
            End If
            'トラニオン位置
            If Me.HdnSelTrunnion.Value <> "" OrElse Me.HdnSelTrunnion.Value <> obj.SelTrunnion Then
                fncSetInfo = True
                obj.SelTrunnion = Me.HdnSelTrunnion.Value
                '変数に規定値を入れる 2017/04/07 追加
                objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = 5
                intCnt = intCnt + 1
            End If
            '二山ナックル・二山クレビス
            If Me.HdnSelClevis.Value <> "0" OrElse Me.HdnSelClevis.Value <> obj.SelClevis Then
                fncSetInfo = True
                obj.SelClevis = CInt(Me.HdnSelClevis.Value)
                '変数に値を入れる 2017/04/07 追加
                objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = CInt(Me.HdnValClevis.Value)
                intCnt = intCnt + 1
            End If
            'タイロッド延長寸法
            If Me.HdnSelTieRod.Value <> "" Then
                fncSetInfo = True
                obj.SelTieRodRadio = Me.HdnSelTieRod.Value
                '変数に規定値を入れる 2017/04/07 追加
                ''シリーズ型番による制御を追加  2017/04/14 追加

                '選択値とのマッチングのためのデータ取得
                '引当シリーズ形番更新(オプション情報)
                dt = dalOutOfOption.fncOutofOpDataChack(objCon, objKtbnStrc.strcSelection.strSeriesKataban, objKtbnStrc.strcSelection.strKeyKataban, _
                                                        Me.HdnSelTieRod.Value)
                If dt.Rows.Count <> 0 Then
                    objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = dt.Rows(0).Item(0)
                Else
                    objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = 1
                End If

                'If objKtbnStrc.strcSelection.strSeriesKataban.Equals("SCS2") Then
                '    objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = 5
                '    intCnt = intCnt + 1
                'Else
                '    objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = 1
                '    intCnt = intCnt + 1
                'End If
            End If
            '標準寸法
            obj.SelTieRodDefl = CInt(Me.HdnSelTieRodDefault.Value)

            '特注寸法
            Dim strTieRodCstm As String = ""
            If Me.HdnSeltxtTieRodCstm.Value.Length > 0 Then
                '特注寸法テキストより設定
                strTieRodCstm = Me.HdnSeltxtTieRodCstm.Value

            ElseIf Me.HdnSelcmbTieRodCstm.Value.Length > 0 Then
                '特注寸法コンボより設定
                strTieRodCstm = Me.HdnSelcmbTieRodCstm.Value

            End If
            If strTieRodCstm <> "" OrElse strTieRodCstm <> obj.SelTieRodCstm Then
                fncSetInfo = True
                obj.SelTieRodCstm = strTieRodCstm
            End If
            'タイロッド材質SUS
            If Me.HdnSelSUS.Value <> "0" OrElse Me.HdnSelSUS.Value <> obj.SelSUS Then
                fncSetInfo = True
                obj.SelSUS = CInt(Me.HdnSelSUS.Value)
                '変数に値を入れる 2017/04/07 追加
                objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = CInt(Me.HdnValSUS.Value)
                intCnt = intCnt + 1
            End If
            'ジャバラ
            If Me.HdnSelJM.Value <> "0" OrElse Me.HdnSelJM.Value <> obj.SelJM Then
                fncSetInfo = True
                obj.SelJM = CInt(Me.HdnSelJM.Value)
                '変数に値を入れる 2017/04/07 追加
                objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = CInt(Me.HdnValJM.Value)
                intCnt = intCnt + 1
            End If
            'フッ素ゴム
            If Me.HdnSelFluoroRub.Value <> "0" OrElse Me.HdnSelFluoroRub.Value <> obj.SelFluoroRub Then
                fncSetInfo = True
                obj.SelFluoroRub = CInt(Me.HdnSelFluoroRub.Value)
                '変数に値を入れる 2017/04/07 追加
                objKtbnStrc.strcSelection.strOutofOpCountryDiv(intCnt) = CInt(Me.HdnValFluoroRub.Value)
                intCnt = intCnt + 1
            End If

            '生産国レベルの判断をここで行う  2017/04/10 追加 松原
            obj.SelPlacelvl = fncGetPlaceLevel("")

            fncSetInfo = fncSetInfo
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    ''' 引当オプション外テーブル/引当シリーズ形番テーブルからデータを削除する
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub subDeleteSelOutOp()
        Dim bolReturn As Boolean
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            '引当オプション外特注クリア
            bolReturn = objOutOp.fncSPSelOutOpDel(objCon)
            '引当シリーズ形番更新(オプション情報)
            Call dalKtbnStrc.subSelSrsKtbnOptionUpd(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                                    objKtbnStrc.strcSelection.strRodEndOption, CdCst.Sign.Blank)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' オプションの生産国レベルを取得
    ''' </summary>
    ''' <param name="strCountryCd">ユーザー国コード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetPlaceLevel(ByVal strCountryCd As String) As Integer
        '生産国レベルを取得する(一番小さい数字を取得する)
        Dim intPlacelvl As Integer = 1024

        For inti As Integer = 0 To objKtbnStrc.strcSelection.strOutofOpCountryDiv.Length - 1
            If objKtbnStrc.strcSelection.strOutofOpCountryDiv(inti) <= 0 Then Continue For
            '単純な比較ではない、リストを展開して比較する
            If intPlacelvl <> CLng(objKtbnStrc.strcSelection.strOutofOpCountryDiv(inti)) Then
                Dim intReal As Integer = 0
                If intPlacelvl <> "1024" Then
                    Dim strMaxOne() As String = KHCountry.fncGetStroke_Logic(objKtbnStrc.strcSelection.strOutofOpCountryDiv(inti)).Split(",")
                    Dim strMinOne() As String = KHCountry.fncGetStroke_Logic(intPlacelvl).Split(",")
                    If strMaxOne.Length > 0 And strMinOne.Length > 0 Then
                        For intl As Integer = 0 To strMaxOne.Length - 1
                            For intk As Integer = 0 To strMinOne.Length - 1
                                If strMaxOne(intl) = strMinOne(intk) Then
                                    intReal += CLng(strMaxOne(intl))
                                End If
                            Next
                        Next
                    End If
                    intPlacelvl = intReal
                Else
                    intPlacelvl = objKtbnStrc.strcSelection.strOutofOpCountryDiv(inti)
                End If
            End If
        Next

        Return intPlacelvl
    End Function

End Class