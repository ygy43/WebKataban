Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class WebUC_RodEnd
    Inherits KHBase

#Region "プロパティ"
    Public Event BacktoYouso()
    Private objRod As KHRodEndCstm
#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        Me.HdnPtnCnt.Value = String.Empty
        Me.HdnSelProdSize.Value = String.Empty
        Me.OnLoad(Nothing)
        Me.HidMessage.Value = ClsCommon.fncGetMsg(selLang.SelectedValue, "W1002")
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
            'ロッド先端特注クラスインスタンス作成
            objRod = New KHRodEndCstm(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                      objKtbnStrc.strcSelection.strSeriesKataban, _
                                      objKtbnStrc.strcSelection.strKeyKataban)
            'ロッド先端特注情報取得
            Call objRod.subRodInfoGet(objCon, objKtbnStrc.strcSelection.strOpSymbol)
            '画面設定
            Call subListMake(objRod)
            Call SetAllFontName(Me)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' ロッド先端特注画面を作成する
    ''' </summary>
    ''' <param name="objRod">ロッド先端特注詳細情報</param>
    ''' <remarks></remarks>
    Private Sub subListMake(ByVal objRod As KHRodEndCstm)
        Dim objCtrl As WebUCRodEnd
        Dim strAppPath As String = System.Web.HttpContext.Current.Request.ApplicationPath
        Dim objRow As TableRow
        Dim objCell As TableCell
        Dim objRadio As RadioButton
        Dim intLoopCnt As Integer
        Dim intAccPos As Integer

        Try
            objRow = New TableRow

            For intLoopCnt = 1 To objRod.RodPtnCnt - 1

                objCell = New TableCell
                objRadio = New RadioButton
                objCtrl = LoadControl(strAppPath & "/WebUserControl/WebUCRodEnd.ascx")

                objCell.Style.Add("text-align", "left")
                'ラジオボタン基本情報設定
                objRadio.GroupName = CdCst.RodEndCstmOrder.RdoGroupNm
                objRadio.Text = objRod.RodPtn(intLoopCnt)
                objRadio.ID = "Rdo" & intLoopCnt
                objRadio.Font.Name = GetFontName(selLang.SelectedValue)
                objRadio.Font.Bold = True

                '寸法表基本情報設定
                With objCtrl
                    'ID
                    .ID = "wucRodEndOrder" & intLoopCnt
                    '選択言語
                    .LangCd = selLang.SelectedValue
                    'ロッド先端パターン記号
                    .RodPtn = objRod.RodPtn(intLoopCnt)
                    'パターンNo.
                    .PtnNo = intLoopCnt
                    'イメージURL
                    .ImageUrl = objRod.KHImageUrl(intLoopCnt)
                    '外径種類
                    .ExtFrm = objRod.ExtFrm(intLoopCnt)
                    '表示外径種類
                    .DispExtFrm = objRod.DispExtFrm(intLoopCnt)
                    '標準寸法
                    .NormalVal = objRod.NormalVal(intLoopCnt)
                    '実際標準寸法
                    .ActNormalVal = objRod.ActNormalVal(intLoopCnt)
                    '入力区分
                    .InputDiv = objRod.InputDiv(intLoopCnt)
                    '選択可能寸法
                    .SltVal = objRod.SltVal(intLoopCnt)
                    '実際選択可能寸法
                    .ActSltVal = objRod.ActSltVal(intLoopCnt)
                    'JavaScript名
                    .JSName = objRod.JsName(intLoopCnt)
                    'メッセージ表示可否
                    .VisibleMsg1 = False
                    .VisibleMsg2 = False
                    '編集区分
                    .EditDiv = Me.objUserInfo.EditDiv

                    '再検索の際に値が存在する場合は選択値をセット
                    If objRod.SelPtn IsNot Nothing Then
                        '現在選択している口径と再検索された口径の値が異なる場合は選択値をセットしない
                        If objRod.SelBoreSize = objRod.BoreSize Then
                            If objRod.RodPtn(intLoopCnt) = objRod.SelPtn Then
                                objRadio.Checked = True
                                If objRod.RodPtn(intLoopCnt).Trim = CdCst.RodEndCstmOrder.OtherSize Then
                                    .SelOtherVal = objRod.SelOtherVal
                                Else
                                    .SelValInfo = objRod.SelValInfo
                                End If
                            End If
                        End If
                    End If
                End With

                '選択可否制御設定
                If objRod.RodPtn(intLoopCnt).Trim = CdCst.RodEndCstmOrder.OtherSize Then
                    'その他寸法エリア設定
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "SSD"
                            Select Case objKtbnStrc.strcSelection.strKeyKataban
                                Case "D"
                                    objRadio.Enabled = False
                                    objCtrl.EnableFlg = False
                            End Select
                        Case "SCA2"
                            Select Case objKtbnStrc.strcSelection.strKeyKataban
                                Case "D"
                                    objRadio.Enabled = False
                                    objCtrl.EnableFlg = False
                            End Select
                        Case "SCS"
                            Select Case objKtbnStrc.strcSelection.strKeyKataban
                                Case "D"
                                    objRadio.Enabled = False
                                    objCtrl.EnableFlg = False
                            End Select
                        Case "CMK2"
                            Select Case objKtbnStrc.strcSelection.strKeyKataban
                                Case "D"
                                    objRadio.Enabled = False
                                    objCtrl.EnableFlg = False
                            End Select
                    End Select
                Else
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "SSD"
                            'メッセージ表示
                            Select Case objRod.RodPtn(intLoopCnt).Trim
                                Case CdCst.RodEndCstmOrder.RodPtnN13N11
                                    objCtrl.VisibleMsg1 = True
                                Case CdCst.RodEndCstmOrder.RodPtnN11N13
                                    objCtrl.VisibleMsg2 = True
                            End Select

                            '口径が12,16の場合N3/N31/N2/N21選択不可
                            Select Case objRod.BoreSize()
                                Case "12", "16"
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN3, CdCst.RodEndCstmOrder.RodPtnN31, _
                                             CdCst.RodEndCstmOrder.RodPtnN2, CdCst.RodEndCstmOrder.RodPtnN21
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                            End Select

                            Select Case objKtbnStrc.strcSelection.strKeyKataban
                                Case ""
                                    '付属品I/Y/I2/Y2が選択された場合N13/N15以外選択不可
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(21)
                                        Case "I", "Y", "I2", "Y2"
                                            Select Case objRod.RodPtn(intLoopCnt).Trim
                                                Case CdCst.RodEndCstmOrder.RodPtnN13, CdCst.RodEndCstmOrder.RodPtnN15
                                                Case Else
                                                    objRadio.Enabled = False
                                                    objCtrl.EnableFlg = False
                                            End Select
                                    End Select
                                    'N13-N11/N11-N13選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN13N11, CdCst.RodEndCstmOrder.RodPtnN11N13
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                                Case "K"
                                    '付属品I/Y/I2/Y2が選択された場合N13/N15以外選択不可
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(19)
                                        Case "I", "Y", "I2", "Y2"
                                            Select Case objRod.RodPtn(intLoopCnt).Trim
                                                Case CdCst.RodEndCstmOrder.RodPtnN13, CdCst.RodEndCstmOrder.RodPtnN15
                                                Case Else
                                                    objRadio.Enabled = False
                                                    objCtrl.EnableFlg = False
                                            End Select
                                    End Select
                                    'バリエーション「U」を含んでいる場合N12/N14選択不可
                                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(1), "U") <> 0 Then
                                        Select Case objRod.RodPtn(intLoopCnt).Trim
                                            Case CdCst.RodEndCstmOrder.RodPtnN12, CdCst.RodEndCstmOrder.RodPtnN14
                                                objRadio.Enabled = False
                                                objCtrl.EnableFlg = False
                                        End Select
                                    End If
                                    'N13-N11/N11-N13選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN13N11, CdCst.RodEndCstmOrder.RodPtnN11N13
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                                Case "D"
                                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(1), "Q") <> 0 Or _
                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "R" Then
                                    Else
                                        Select Case objRod.RodPtn(intLoopCnt).Trim
                                            Case CdCst.RodEndCstmOrder.RodPtnN11N13
                                                objRadio.Enabled = False
                                                objCtrl.EnableFlg = False
                                        End Select
                                    End If
                                    '中間ストロークの場合N11-N13選択可
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        Case "5", "10", "15", "20", "25", "30", "40", "50", _
                                             "60", "70", "80", "90", "100", "110", "120", _
                                             "130", "140", "150", "160", "170", "180", "190", _
                                             "200", "210", "220", "230", "240", "250", "260", _
                                             "270", "280", "290", "300"
                                            Select Case objRod.RodPtn(intLoopCnt).Trim
                                                Case CdCst.RodEndCstmOrder.RodPtnN11N13
                                                    objRadio.Enabled = False
                                                    objCtrl.EnableFlg = False
                                            End Select
                                    End Select
                                    '支持金具に「FA」を含む場合N11-N13選択可
                                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(12), "FA") <> 0 Then
                                        Select Case objRod.RodPtn(intLoopCnt).Trim
                                            Case CdCst.RodEndCstmOrder.RodPtnN11N13
                                                objRadio.Enabled = True
                                                objCtrl.EnableFlg = True
                                        End Select
                                    End If
                                    'N13-N11/N11-N13以外選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN13N11, CdCst.RodEndCstmOrder.RodPtnN11N13
                                        Case Else
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                            End Select
                        Case "JSC3"
                            '付属品I/Y/I2/Y2が選択された場合N13/N15以外選択不可
                            If InStr(objKtbnStrc.strcSelection.strOpSymbol(14), "I") <> 0 Or _
                               InStr(objKtbnStrc.strcSelection.strOpSymbol(14), "Y") <> 0 Then
                                Select Case objRod.RodPtn(intLoopCnt).Trim
                                    Case CdCst.RodEndCstmOrder.RodPtnN13, CdCst.RodEndCstmOrder.RodPtnN15
                                    Case Else
                                        objRadio.Enabled = False
                                        objCtrl.EnableFlg = False
                                End Select
                            End If
                        Case "JSC4"
                            '付属品I/Y/I2/Y2が選択された場合N13/N15以外選択不可
                            If InStr(objKtbnStrc.strcSelection.strOpSymbol(14), "I") <> 0 Or _
                               InStr(objKtbnStrc.strcSelection.strOpSymbol(14), "Y") <> 0 Then
                                Select Case objRod.RodPtn(intLoopCnt).Trim
                                    Case CdCst.RodEndCstmOrder.RodPtnN13, CdCst.RodEndCstmOrder.RodPtnN15
                                    Case Else
                                        objRadio.Enabled = False
                                        objCtrl.EnableFlg = False
                                End Select
                            End If
                        Case "SCA2"
                            'メッセージ表示
                            Select Case objRod.RodPtn(intLoopCnt).Trim
                                Case CdCst.RodEndCstmOrder.RodPtnN13N11
                                    objCtrl.VisibleMsg1 = True
                                Case CdCst.RodEndCstmOrder.RodPtnN11N13
                                    objCtrl.VisibleMsg2 = True
                            End Select

                            '付属品初期値設定
                            intAccPos = 0

                            '付属品I/Y/I2/Y2が選択された場合N13/N15以外選択不可
                            Select Case objKtbnStrc.strcSelection.strKeyKataban
                                Case "", "V"
                                    intAccPos = 14
                                Case "B"
                                    intAccPos = 18
                                Case "D"
                                    intAccPos = 13
                                Case "2"
                                    intAccPos = 15
                                Case "C"
                                    intAccPos = 19
                                Case "E"
                                    intAccPos = 14
                            End Select
                            If InStr(objKtbnStrc.strcSelection.strOpSymbol(intAccPos), "I") <> 0 Or _
                               InStr(objKtbnStrc.strcSelection.strOpSymbol(intAccPos), "Y") <> 0 Then
                                Select Case objRod.RodPtn(intLoopCnt).Trim
                                    Case CdCst.RodEndCstmOrder.RodPtnN13, CdCst.RodEndCstmOrder.RodPtnN15
                                    Case Else
                                        objRadio.Enabled = False
                                        objCtrl.EnableFlg = False
                                End Select
                            End If

                            Select Case objKtbnStrc.strcSelection.strKeyKataban
                                Case "", "V", "B", "2", "C"
                                    'N13-N11/N11-N13を選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN13N11, CdCst.RodEndCstmOrder.RodPtnN11N13
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                                Case "D", "E"
                                    ' バリエーションに「Q」を含み、落下防止機構で「HR」を選択しない場合は「N11-N13」は選択可
                                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(1), "Q") <> 0 And _
                                       InStr(objKtbnStrc.strcSelection.strOpSymbol(8), "HR") = 0 Then
                                    Else
                                        Select Case objRod.RodPtn(intLoopCnt).Trim
                                            Case CdCst.RodEndCstmOrder.RodPtnN11N13
                                                objRadio.Enabled = False
                                                objCtrl.EnableFlg = False
                                        End Select
                                    End If
                                    'N13-N11/N11-N13以外を選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN13N11, CdCst.RodEndCstmOrder.RodPtnN11N13
                                        Case Else
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                            End Select
                        Case "SCS"
                            Select Case objKtbnStrc.strcSelection.strKeyKataban
                                Case "", "B"
                                    'オプション「A2」を含む、もしくは付属品「I」「Y」を含む場合はN13/N15以外を選択不可
                                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(17), "A2") <> 0 Or _
                                       InStr(objKtbnStrc.strcSelection.strOpSymbol(18), "I") <> 0 Or _
                                       InStr(objKtbnStrc.strcSelection.strOpSymbol(18), "Y") <> 0 Then
                                        Select Case objRod.RodPtn(intLoopCnt).Trim
                                            Case CdCst.RodEndCstmOrder.RodPtnN13, CdCst.RodEndCstmOrder.RodPtnN15
                                            Case Else
                                                objRadio.Enabled = False
                                                objCtrl.EnableFlg = False
                                        End Select
                                    End If
                                    'N13-N11は選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN13N11
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                                Case "D"
                                    'N13-N11以外は選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN13N11
                                        Case Else
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                            End Select
                        Case "SCS2"
                            Select Case objKtbnStrc.strcSelection.strKeyKataban
                                Case "", "B", "F"
                                    'N3,N31,N2,N21は選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN3, CdCst.RodEndCstmOrder.RodPtnN31, CdCst.RodEndCstmOrder.RodPtnN2, CdCst.RodEndCstmOrder.RodPtnN21
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                                    ' オプション「A2」が選択されていた場合は「N13」「N15」以外は非表示
                                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(18), "A2") <> 0 Then
                                        Select Case objRod.RodPtn(intLoopCnt).Trim
                                            Case CdCst.RodEndCstmOrder.RodPtnN13, CdCst.RodEndCstmOrder.RodPtnN15
                                            Case Else
                                                objRadio.Enabled = False
                                                objCtrl.EnableFlg = False
                                        End Select
                                    End If
                                    ' 付属品「I」「Y」が選択されていた場合は「N13」「N15」以外は非表示
                                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(19), "I") <> 0 Or _
                                       InStr(objKtbnStrc.strcSelection.strOpSymbol(19), "Y") <> 0 Then
                                        Select Case objRod.RodPtn(intLoopCnt).Trim
                                            Case CdCst.RodEndCstmOrder.RodPtnN13, CdCst.RodEndCstmOrder.RodPtnN15
                                            Case Else
                                                objRadio.Enabled = False
                                                objCtrl.EnableFlg = False
                                        End Select
                                    End If
                                    'N13-N11は選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN13N11
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                                Case "D", "G"
                                    'N3,N31,N2,N21は選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN3, CdCst.RodEndCstmOrder.RodPtnN31, CdCst.RodEndCstmOrder.RodPtnN2, CdCst.RodEndCstmOrder.RodPtnN21
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                                    'N13-N11以外は選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN13N11
                                        Case Else
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                            End Select

                        Case "CMK2"
                            'メッセージ表示
                            Select Case objRod.RodPtn(intLoopCnt).Trim
                                Case CdCst.RodEndCstmOrder.RodPtnN13N11
                                    objCtrl.VisibleMsg1 = True
                                Case CdCst.RodEndCstmOrder.RodPtnN11N13
                                    objCtrl.VisibleMsg2 = True
                            End Select

                            Select Case objKtbnStrc.strcSelection.strKeyKataban
                                Case ""
                                    '付属品「I」「Y」を含む場合はN13/N15以外を選択不可
                                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(16), "I") <> 0 Or _
                                       InStr(objKtbnStrc.strcSelection.strOpSymbol(16), "Y") <> 0 Then
                                        Select Case objRod.RodPtn(intLoopCnt).Trim
                                            Case CdCst.RodEndCstmOrder.RodPtnN13, CdCst.RodEndCstmOrder.RodPtnN15
                                            Case Else
                                                objRadio.Enabled = False
                                                objCtrl.EnableFlg = False
                                        End Select
                                    End If
                                    'N13-N11,N11-N13は選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN13N11, CdCst.RodEndCstmOrder.RodPtnN11N13
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                                Case "D"
                                    'N13-N11以外は選択不可
                                    Select Case objRod.RodPtn(intLoopCnt).Trim
                                        Case CdCst.RodEndCstmOrder.RodPtnN13N11
                                        Case CdCst.RodEndCstmOrder.RodPtnN11N13
                                            'バリエーションに「Q」を含む場合(この場合「DQ」のみ)は「N11-N13」は選択可
                                            If InStr(objKtbnStrc.strcSelection.strOpSymbol(1), "Q") = 0 Then
                                                objRadio.Enabled = False
                                                objCtrl.EnableFlg = False
                                            End If
                                        Case Else
                                            objRadio.Enabled = False
                                            objCtrl.EnableFlg = False
                                    End Select
                            End Select
                    End Select
                End If

                objCell.Controls.Add(objRadio)
                objCell.Controls.Add(objCtrl)

                Dim clearLabel As New Label
                clearLabel.Height = 20
                objCell.Controls.Add(clearLabel)

                objRow.Cells.Add(objCell)
                objRow.Style.Add("padding-bottom", "10px")
                If intLoopCnt Mod 2 <> 0 Then
                    If intLoopCnt = objRod.RodPtnCnt - 1 Then
                        Me.TblRodLst.Rows.Add(objRow)
                    End If
                Else
                    Me.TblRodLst.Rows.Add(objRow)
                    objRow = New TableRow
                End If
            Next
            Me.HdnPtnCnt.Value = objRod.RodPtnCnt - 1
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' キャンセルイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        'ロッド先端引当情報削除
        subDeleteSelRod(objKtbnStrc.strcSelection.strOtherOption)
        Me.HdnPtnCnt.Value = String.Empty
        Me.HdnSelProdSize.Value = String.Empty
        RaiseEvent BacktoYouso()
    End Sub

    ''' <summary>
    ''' OKボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        'ロッド先端特注クラスインスタンス作成
        objRod = New KHRodEndCstm(Me.objUserInfo.UserId, Me.objLoginInfo.SessionId, _
                                  objKtbnStrc.strcSelection.strSeriesKataban, _
                                  objKtbnStrc.strcSelection.strKeyKataban)
        'ロッド先端特注標準情報取得
        Call objRod.subRodInfoGet(objCon, objKtbnStrc.strcSelection.strOpSymbol)
        '入力チェック
        If Not fncInpRodCheck() Then Exit Sub
        '引当情報更新
        Call objRod.subUpdateSelRod(objCon, objKtbnStrc)
        RaiseEvent BacktoYouso()
    End Sub

    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncInpRodCheck() As Boolean
        Dim strSelPtnNo As String = String.Empty
        Dim strProdSize As String = String.Empty
        Dim strSelProdSize() As String
        Dim intLoopCnt As Integer
        fncInpRodCheck = False

        Try
            '選択情報取得
            Dim intCellsCount As Integer = 0

            For Each dr As WebControls.TableRow In TblRodLst.Rows
                intCellsCount += dr.Cells.Count
            Next

            For inti As Integer = 1 To intCellsCount
                Dim obj As RadioButton = TblRodLst.FindControl("Rdo" & inti)
                If Not obj Is Nothing AndAlso obj.Checked Then
                    strSelPtnNo = inti
                    Exit For
                End If
            Next

            If Len(strSelPtnNo) <> 0 Then
                strSelProdSize = Split(Me.HdnSelProdSize.Value, CdCst.Sign.Delimiter.Pipe)
                '特注寸法編集
                For intLoopCnt = 0 To strSelProdSize.Length - 1
                    If Right(strSelProdSize(intLoopCnt).Trim, 2) = ".0" Then
                        strSelProdSize(intLoopCnt) = Left(strSelProdSize(intLoopCnt), InStr(strSelProdSize(intLoopCnt), ".0") - 1)
                    End If
                Next

                'ラジオボタンの選択がある場合
                If Not objRod.fncInpCheck(strSelPtnNo, strSelProdSize, objKtbnStrc.strcSelection.strOpSymbol) Then
                    AlertMessage(objRod.ErrCd, objRod.ErrOption)
                    'subSetErrScript(objRod.ErrPtnNo, objRod.ErrPtn, objRod.ErrFocusNo, objRod.ErrCd, objRod.ErrOption)
                    Exit Function
                End If
            Else
                AlertMessage("W8480")
                Exit Function
            End If
            fncInpRodCheck = True
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Function

    ''' <summary>
    ''' 選択した情報の削除
    ''' </summary>
    ''' <param name="strOtherOP"></param>
    ''' <remarks></remarks>
    Public Sub subDeleteSelRod(ByVal strOtherOP As String)
        Dim bolReturn As Boolean
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            '引当ロッド先端特注クリア
            bolReturn = objRod.fncSPSelRodDel(objCon)
            '引当シリーズ形番更新(オプション情報)
            Call dalKtbnStrc.subSelSrsKtbnOptionUpd(objCon, Me.objUserInfo.UserId, _
                                                    Me.objLoginInfo.SessionId, CdCst.Sign.Blank, strOtherOP)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

End Class