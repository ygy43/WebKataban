Imports System.Data.SqlClient
Imports WebKataban.ClsCommon

Public Class KHOutOfOptionCstm

#Region " Fixed Value "

    Private Const CST_BLANK As String = CdCst.Sign.Blank
    Private Const CST_INT_DEFAULT As Integer = 0
    Private Const CST_DISP_DEFAULT As Integer = 9
    Private Const CST_DISP_HIDE As Integer = -1
    Private Const CST_DISP_UNENABLE As Integer = 0
    Private Const CST_DISP_ENABLE As Integer = 1
    Private Const CST_COMMA As String = CdCst.Sign.Delimiter.Comma

#End Region

#Region " Definition "

    '引当形番情報
    Private Structure Selection
        Public strSeriesKataban As String                      'シリーズ形番
        Public strKeyKataban As String                         'キー形番
        Public strBoreSize As String                           '引当口径
        Public strFullKtbn As String                           'フル形番
        Public strUserID As String                             'ユーザID
        Public strLang As String                               '言語
        Public strSessionID As String                          'セッションID
    End Structure
    Private strcSelection As Selection

    'イメージパス情報
    Private Structure ImagePathInfo
        Public strPortPath As String             'ポート・クッションニードル位置用イメージ
        Public strPortExePath As String          'ポート・クッションニードル位置例用イメージ
        Public strMountingPath As String         '支持金具用イメージ
        Public strTrunnionPath As String         'トラニオン位置用イメージ
        Public strTieRodPath As String           'タイロッド延長寸法用イメージ
    End Structure
    Private strcImagePathInfo As ImagePathInfo

    'オプション外指定表示情報
    Private Structure DataInfo
        'ArreyList→Datatabeleへ変更  2017/04/06 
        Public intPortCushion As Integer         'ポート・クッションニードル位置表示状態(-1:非表示、0:使用不可,1:使用可,9:初期値)
        Public lstPortCushion As DataTable       'ポート・クッションニードル位置リスト
        Public intPort As Integer                'ポート２箇所表示状態(-1:非表示、0:使用不可,1:使用可,9:初期値)
        Public lstPort As DataTable              'ポート２箇所リスト
        Public intPortSize As Integer            'ポートサイズダウン表示状態(-1:非表示、0:使用不可,1:使用可,9:初期値)
        Public lstPortSize As DataTable          'ポートサイズダウンリスト
        Public intMounting As Integer            '支持金具回転表示状態(-1:非表示、0:使用不可,1:使用可,9:初期値)
        Public lstMounting As DataTable          '支持金具回転リスト
        Public intTrunnion As Integer            'トラニオン位置表示状態(-1:非表示、0:使用不可,1:使用可,9:初期値)
        Public intClevis As Integer              '二山ナックル・二山クレビス表示状態(-1:非表示、0:使用不可,1:使用可,9:初期値)
        Public lstClevis As DataTable            '二山ナックル・二山クレビスリスト
        Public intTieRod As Integer              'タイロッド延長寸法表示状態(-1:非表示、0:使用不可,1:使用可,9:初期値)
        Public lstTieRodRadio As ArrayList       'タイロッド延長寸法ラジオボタンリスト
        Public strTieRodDefl As String           'タイロッド延長寸法標準寸法
        Public lstTieRodCstm As ArrayList        'タイロッド延長寸法特注寸法リスト
        Public intSUS As Integer                 'タイロッド材質SUS表示状態(-1:非表示、0:使用不可,1:使用可,9:初期値)
        Public lstSUS As DataTable               'タイロッド材質SUSリスト
        Public intJM As Integer                  'ジャバラ表示状態(-1:非表示、0:使用不可,1:使用可,9:初期値)
        Public lstJM As DataTable                'ジャバラリスト
        Public intFluoroRub As Integer           'フッ素ゴム表示状態(-1:非表示、0:使用不可,1:使用可,9:初期値)
        Public lstFluoroRub As DataTable         'フッ素ゴムリスト
    End Structure
    Private strcDataInfo As DataInfo

    'オプション外指定選択情報
    Private Structure SelDataInfo
        Public intSelPortCushion As Integer      'ポート・クッションニードル位置
        Public strSelPortCuPlace As String       'ポート・クッションニードル位置ラジオ連結
        Public intSelPort As Integer             'ポート２箇所
        Public intSelPortSize As Integer         'ポートサイズダウン
        Public intSelMounting As Integer         '支持金具回転
        Public strSelTrunnion As String          'トラニオン位置
        Public intSelClevis As Integer           '二山ナックル・二山クレビス
        Public strSelTieRodRadio As String       'タイロッド延長寸法ラジオボタン
        Public intSelTieRodDefl As Integer       'タイロッド延長寸法標準寸法
        Public strSelTieRodCstm As String        'タイロッド延長寸法特注寸法
        Public intSelSUS As Integer              'タイロッド材質SUS
        Public intSelJM As Integer               'ジャバラ
        Public intSelFluoroRub As Integer        'フッ素ゴム
        Public intPlacelvl As Integer            '国コード   2017/04/10 追加 松原
    End Structure
    Private strcSelDataInfo As SelDataInfo

    'エラー情報
    Private Structure ErrInfo
        Public strErrCd As String                               'エラーコード
        Public strErrFocus As String                            'エラーコントロールID
    End Structure
    Private strcErrInfo As ErrInfo

#End Region

    ''' <summary>
    ''' フィールドの初期設定
    ''' </summary>
    ''' <param name="strAUserID">ユーザーID</param>
    ''' <param name="strASessionID">セッションID</param>
    ''' <param name="strASelectLang"></param>
    ''' <param name="strASeriesKataban">シリーズ形番</param>
    ''' <param name="strAKeyKataban">キー形番</param>
    ''' <remarks>ユーザーID/セッションID/シリーズ形番/キー形番を保持</remarks>
    Public Sub New(ByVal strAUserID As String, _
                   ByVal strASessionID As String, _
                   ByVal strASelectLang As String, _
                   ByVal strASeriesKataban As String, _
                   ByVal strAKeyKataban As String)

        'フィールド初期設定
        With Me.strcSelection
            .strSeriesKataban = CST_BLANK
            .strKeyKataban = CST_BLANK
            .strBoreSize = CST_BLANK
            .strFullKtbn = CST_BLANK
            .strUserID = CST_BLANK
            .strLang = CST_BLANK
            .strSessionID = CST_BLANK
        End With

        With Me.strcImagePathInfo
            .strPortPath = CST_BLANK
            .strPortExePath = CST_BLANK
            .strMountingPath = CST_BLANK
            .strTrunnionPath = CST_BLANK
            .strTieRodPath = CST_BLANK
        End With

        With Me.strcDataInfo
            .intPortCushion = CST_DISP_DEFAULT
            .lstPortCushion = Nothing
            .intPort = CST_DISP_DEFAULT
            .lstPort = Nothing
            .intPortSize = CST_DISP_DEFAULT
            .lstPortSize = Nothing
            .intMounting = CST_DISP_DEFAULT
            .lstMounting = Nothing
            .intTrunnion = CST_DISP_DEFAULT
            .intClevis = CST_DISP_DEFAULT
            .lstClevis = Nothing
            .intTieRod = CST_DISP_DEFAULT
            .lstTieRodRadio = Nothing
            .strTieRodDefl = CST_BLANK
            .lstTieRodCstm = Nothing
            .intSUS = CST_DISP_DEFAULT
            .lstSUS = Nothing
            .intJM = CST_DISP_DEFAULT
            .lstJM = Nothing
            .intFluoroRub = CST_DISP_DEFAULT
            .lstFluoroRub = Nothing
        End With

        With Me.strcSelDataInfo
            .intSelPortCushion = CST_INT_DEFAULT
            .strSelPortCuPlace = CST_BLANK
            .intSelPort = CST_INT_DEFAULT
            .intSelPortSize = CST_INT_DEFAULT
            .intSelMounting = CST_INT_DEFAULT
            .strSelTrunnion = CST_BLANK
            .intSelClevis = CST_INT_DEFAULT
            .strSelTieRodRadio = CST_BLANK
            .intSelTieRodDefl = CST_INT_DEFAULT
            .strSelTieRodCstm = CST_BLANK
            .intSelSUS = CST_INT_DEFAULT
            .intSelJM = CST_INT_DEFAULT
            .intSelFluoroRub = CST_INT_DEFAULT
        End With

        With Me.strcErrInfo
            .strErrCd = CST_BLANK
            .strErrFocus = CST_BLANK
        End With

        Me.strcSelection.strUserID = strAUserID
        Me.strcSelection.strSessionID = strASessionID
        Me.strcSelection.strLang = strASelectLang
        Me.strcSelection.strSeriesKataban = strASeriesKataban
        Me.strcSelection.strKeyKataban = strAKeyKataban

    End Sub

    ''' <summary>
    ''' オプション外指定情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strOpSymbol">形番引当画面の引当オプション </param>
    ''' <remarks>オプション外指定画面生成に必要な情報を取得する</remarks>
    Public Sub subOutOpInfoGet(ByVal objCon As SqlConnection, ByVal strOpSymbol As String())
        Try
            '口径クリア
            Me.strcSelection.strBoreSize = CST_BLANK

            '表示データクリア
            With Me.strcDataInfo
                '変数をデータテーブルに変更したための修正　2017/04/06
                '必要な各データテーブルに項目設定を行う

                .intPortCushion = CST_DISP_HIDE
                .lstPortCushion = New DataTable
                .lstPortCushion.NewRow()
                .lstPortCushion.Columns.Add("ITEM1", Type.GetType("System.String"))
                .lstPortCushion.Columns.Add("ITEM2", Type.GetType("System.String"))

                .intPort = CST_DISP_HIDE
                .lstPort = New DataTable
                .lstPort.NewRow()
                .lstPort.Columns.Add("ITEM1", Type.GetType("System.String"))
                .lstPort.Columns.Add("ITEM2", Type.GetType("System.String"))

                .intPortSize = CST_DISP_HIDE
                .lstPortSize = New DataTable
                .lstPortSize.NewRow()
                .lstPortSize.Columns.Add("ITEM1", Type.GetType("System.String"))
                .lstPortSize.Columns.Add("ITEM2", Type.GetType("System.String"))

                .intMounting = CST_DISP_HIDE
                .lstMounting = New DataTable
                .lstMounting.NewRow()
                .lstMounting.Columns.Add("ITEM1", Type.GetType("System.String"))
                .lstMounting.Columns.Add("ITEM2", Type.GetType("System.String"))

                .intTrunnion = CST_DISP_HIDE

                .intClevis = CST_DISP_HIDE
                .lstClevis = New DataTable
                .lstClevis.NewRow()
                .lstClevis.Columns.Add("ITEM1", Type.GetType("System.String"))
                .lstClevis.Columns.Add("ITEM2", Type.GetType("System.String"))

                .intTieRod = CST_DISP_HIDE
                .lstTieRodRadio = New ArrayList

                .strTieRodDefl = CST_BLANK
                .lstTieRodCstm = New ArrayList

                .intSUS = CST_DISP_HIDE
                .lstSUS = New DataTable
                .lstSUS.NewRow()
                .lstSUS.Columns.Add("ITEM1", Type.GetType("System.String"))
                .lstSUS.Columns.Add("ITEM2", Type.GetType("System.String"))

                .intJM = CST_DISP_HIDE
                .lstJM = New DataTable
                .lstJM.NewRow()
                .lstJM.Columns.Add("ITEM1", Type.GetType("System.String"))
                .lstJM.Columns.Add("ITEM2", Type.GetType("System.String"))

                .intFluoroRub = CST_DISP_HIDE
                .lstFluoroRub = New DataTable
                .lstFluoroRub.NewRow()
                .lstFluoroRub.Columns.Add("ITEM1", Type.GetType("System.String"))
                .lstFluoroRub.Columns.Add("ITEM2", Type.GetType("System.String"))

            End With

            ''選択情報クリア
            With Me.strcSelDataInfo
                .intSelPortCushion = CST_INT_DEFAULT
                .strSelPortCuPlace = CST_BLANK
                .intSelPort = CST_INT_DEFAULT
                .intSelPortSize = CST_INT_DEFAULT
                .intSelMounting = CST_INT_DEFAULT
                .strSelTrunnion = CST_BLANK
                .intSelClevis = CST_INT_DEFAULT
                .strSelTieRodRadio = CST_BLANK
                .intSelTieRodDefl = CST_INT_DEFAULT
                .strSelTieRodCstm = CST_BLANK
                .intSelSUS = CST_INT_DEFAULT
                .intSelJM = CST_INT_DEFAULT
                .intSelFluoroRub = CST_INT_DEFAULT
            End With

            '口径取得
            Call subBoreSizeSelect(objCon, strOpSymbol)

            '画面表示パターン詳細情報取得
            Call subOutOpDtlSet(objCon, strOpSymbol)

            '画面選択情報取得
            Call subSelOutOfOpSelect(objCon)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当口径検索取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strOpSymbol">形番引当画面の引当オプション</param>
    ''' <remarks></remarks>
    Private Sub subBoreSizeSelect(ByVal objCon As SqlConnection, ByVal strOpSymbol() As String)
        Dim dt As New DataTable
        Dim dalOutOfOption As New OutOfOptionDAL

        Try
            dt = dalOutOfOption.fncBoreSizeSelect(objCon, Me.strcSelection.strSeriesKataban, Me.strcSelection.strKeyKataban)

            For Each dr As DataRow In dt.Rows
                Me.strcSelection.strBoreSize = IIf(IsDBNull(dr("ktbn_strc_seq_no")), CST_BLANK, strOpSymbol(dr("ktbn_strc_seq_no")))
            Next

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 画面表示詳細情報セット
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strOpSymbol">選択オプション</param>
    ''' <remarks>画面表示詳細情報をメンバ変数にセットする</remarks>
    Private Sub subOutOpDtlSet(ByVal objCon As SqlConnection, ByVal strOpSymbol As String())
        Dim dt As New DataTable
        Dim dtRow As DataRow

        Try
            '画面詳細情報取得
            If Not subOutofOpDataSelect(objCon, dt) Then
                Exit Sub
            End If

            If dt IsNot Nothing Then
                For Each dr As DataRow In dt.Rows
                    Select Case dr(0)
                        'データテーブル型に変更したためにここも変更  2017/04/06 
                        Case 0
                            Me.strcDataInfo.intPortCushion = CST_DISP_ENABLE

                            dtRow = Me.strcDataInfo.lstPortCushion.NewRow()
                            dtRow("ITEM1") = dr(2)
                            dtRow("ITEM2") = dr(3)
                            Me.strcDataInfo.lstPortCushion.Rows.Add(dtRow)
                            'Me.strcDataInfo.lstPortCushion.Add(dr(2))
                        Case 1
                            Me.strcDataInfo.intPort = CST_DISP_ENABLE
                            dtRow = Me.strcDataInfo.lstPort.NewRow()
                            dtRow("ITEM1") = dr(2)
                            dtRow("ITEM2") = dr(3)
                            Me.strcDataInfo.lstPort.Rows.Add(dtRow)
                            'Me.strcDataInfo.lstPort.Add(dr(2))
                        Case 2
                            Me.strcDataInfo.intPortSize = CST_DISP_ENABLE

                            dtRow = Me.strcDataInfo.lstPortSize.NewRow()
                            dtRow("ITEM1") = dr(2)
                            dtRow("ITEM2") = dr(3)
                            Me.strcDataInfo.lstPortSize.Rows.Add(dtRow)
                            'Me.strcDataInfo.lstPortSize.Add(dr(2))
                        Case 3
                            Me.strcDataInfo.intMounting = CST_DISP_ENABLE

                            dtRow = Me.strcDataInfo.lstMounting.NewRow()
                            dtRow("ITEM1") = dr(2)
                            dtRow("ITEM2") = dr(3)
                            Me.strcDataInfo.lstMounting.Rows.Add(dtRow)
                            'Me.strcDataInfo.lstMounting.Add(dr(2))
                        Case 4
                            Me.strcDataInfo.intTrunnion = CST_DISP_ENABLE
                        Case 5
                            Me.strcDataInfo.intClevis = CST_DISP_ENABLE

                            dtRow = Me.strcDataInfo.lstClevis.NewRow()
                            dtRow("ITEM1") = dr(2)
                            dtRow("ITEM2") = dr(3)
                            Me.strcDataInfo.lstClevis.Rows.Add(dtRow)
                            'Me.strcDataInfo.lstClevis.Add(dr(2))

                        Case 6
                            Me.strcDataInfo.intTieRod = CST_DISP_ENABLE
                            Me.strcDataInfo.lstTieRodRadio.Add(dr(2))
                        Case 7
                            Me.strcDataInfo.strTieRodDefl = dr(2)
                        Case 8
                            Me.strcDataInfo.lstTieRodCstm.Add(dr(2))
                        Case 9
                            Me.strcDataInfo.intSUS = CST_DISP_ENABLE

                            dtRow = Me.strcDataInfo.lstSUS.NewRow()
                            dtRow("ITEM1") = dr(2)
                            dtRow("ITEM2") = dr(3)
                            Me.strcDataInfo.lstSUS.Rows.Add(dtRow)
                            'Me.strcDataInfo.lstSUS.Add(dr(2))
                        Case 10
                            Me.strcDataInfo.intJM = CST_DISP_ENABLE

                            dtRow = Me.strcDataInfo.lstJM.NewRow()
                            dtRow("ITEM1") = dr(2)
                            dtRow("ITEM2") = dr(3)
                            Me.strcDataInfo.lstJM.Rows.Add(dtRow)
                            'Me.strcDataInfo.lstJM.Add(dr(2))
                        Case 11
                            Me.strcDataInfo.intFluoroRub = CST_DISP_ENABLE

                            dtRow = Me.strcDataInfo.lstFluoroRub.NewRow()
                            dtRow("ITEM1") = dr(2)
                            dtRow("ITEM2") = dr(3)
                            Me.strcDataInfo.lstFluoroRub.Rows.Add(dtRow)
                            'Me.strcDataInfo.lstFluoroRub.Add(dr(2))
                    End Select
                Next
            End If

            Select Case Me.strcSelection.strSeriesKataban
                '2012/07/23 オプション外指定追加
                Case "SCA2"
                    '画像ファイル指定
                    With Me.strcImagePathInfo
                        .strPortPath = "../KHImage/outOp1.gif"
                        .strPortExePath = "../KHImage/outOp2.gif"
                        .strMountingPath = "../KHImage/outOp3.gif"
                        .strTrunnionPath = "../KHImage/outOp4.gif"
                        .strTieRodPath = "../KHImage/outOp5_SCS.gif"
                    End With

                    Select Case Me.strcSelection.strKeyKataban
                        '基本形の場合
                        Case "", "2"
                            '引当画面.クッション＝Ｂ以外は使用不可
                            If Trim(strOpSymbol(6)) <> "B" Or _
                               Trim(strOpSymbol(13)) = "S" Or _
                               Trim(strOpSymbol(13)) = "T" Then
                                'ポートクッションニードル位置指定を使用不可
                                Me.strcDataInfo.intPortCushion = CST_DISP_UNENABLE
                            End If
                            'ポート２箇所を非表示
                            Me.strcDataInfo.intPort = CST_DISP_HIDE

                            'ポートサイズダウンを非表示
                            Me.strcDataInfo.intPortSize = CST_DISP_HIDE

                            '支持金具回転
                            Select Case strOpSymbol(3)
                                Case "LB", "FB", "CA", "CB", "TA", "TB", "TC", "TD", "TE", "FA", "TF"
                                Case Else
                                    Me.strcDataInfo.intMounting = CST_DISP_UNENABLE
                            End Select

                            'トラニオン位置指定を使用不可
                            If Trim(strOpSymbol(3)) <> "TC" And _
                               Trim(strOpSymbol(3)) <> "TF" Then
                                Me.strcDataInfo.intTrunnion = CST_DISP_UNENABLE
                            End If

                            'タイロッド延長寸法を非表示
                            Me.strcDataInfo.intTieRod = CST_DISP_UNENABLE

                            'タイロッド材質ＳＵＳを使用不可
                            If InStr(strOpSymbol(1), "K") <> 0 Then
                                Me.strcDataInfo.intSUS = CST_DISP_UNENABLE
                            End If

                            'ピストンロッドはジャバラ付寸法でジャバラなし
                            If InStr(strOpSymbol(1), "O") <> 0 Or InStr(strOpSymbol(1), "Q2") <> 0 Or _
                               InStr(strOpSymbol(1), "U") <> 0 Or InStr(strOpSymbol(1), "G") <> 0 Or _
                               InStr(strOpSymbol(13), "J") <> 0 Or InStr(strOpSymbol(13), "K") <> 0 Or _
                               InStr(strOpSymbol(13), "L") <> 0 Then
                                Me.strcDataInfo.intJM = CST_DISP_UNENABLE
                            End If
                        Case "B", "C"
                            'ポートクッションニードル位置指定を非表示
                            Me.strcDataInfo.intPortCushion = CST_DISP_HIDE
                            'ポート２箇所を非表示
                            Me.strcDataInfo.intPort = CST_DISP_HIDE
                            'ポートサイズダウンを非表示
                            Me.strcDataInfo.intPortSize = CST_DISP_HIDE
                            '支持金具回転は使用不可
                            Select Case strOpSymbol(3)
                                Case "LB", "FB", "CA", "CB", "TA", "TB", "TC", "TD", "TE", "FA", "TF"
                                Case Else
                                    Me.strcDataInfo.intMounting = CST_DISP_UNENABLE
                            End Select
                            'トラニオン位置指定は非表示
                            Me.strcDataInfo.intTrunnion = CST_DISP_HIDE

                            'タイロッド延長寸法を非表示
                            Me.strcDataInfo.intTieRod = CST_DISP_UNENABLE
                            'タイロッド材質ＳＵＳ
                            If InStr(strOpSymbol(1), "K") <> 0 Then
                                Me.strcDataInfo.intSUS = CST_DISP_UNENABLE
                            End If
                            'ピストンロッドはジャバラ付寸法でジャバラなし
                            If InStr(strOpSymbol(1), "O") <> 0 Or InStr(strOpSymbol(1), "G") <> 0 Or _
                               InStr(strOpSymbol(1), "G1") <> 0 Or InStr(strOpSymbol(1), "G2") <> 0 Or _
                               InStr(strOpSymbol(1), "G3") <> 0 Or InStr(strOpSymbol(1), "G4") <> 0 Or _
                               InStr(strOpSymbol(17), "J") <> 0 Or InStr(strOpSymbol(17), "L") <> 0 Or _
                               InStr(strOpSymbol(17), "K") <> 0 Then
                                Me.strcDataInfo.intJM = CST_DISP_UNENABLE
                            End If
                            'スクレーパ、ロッドパッキンのみフッ素ゴム
                            If InStr(strOpSymbol(1), "O") = 0 Or InStr(strOpSymbol(1), "H") = 0 Or _
                               InStr(strOpSymbol(1), "T") = 0 Or InStr(strOpSymbol(1), "T1") = 0 Or _
                               InStr(strOpSymbol(1), "T2") = 0 Or InStr(strOpSymbol(1), "G") = 0 Or _
                               InStr(strOpSymbol(1), "G1") = 0 Or InStr(strOpSymbol(1), "G2") = 0 Or _
                               InStr(strOpSymbol(1), "G3") = 0 Or InStr(strOpSymbol(1), "G4") = 0 Then
                                Me.strcDataInfo.intFluoroRub = CST_DISP_UNENABLE
                            End If
                        Case "D", "E"
                            '引当画面.クッション＝Ｂ以外は使用不可
                            If Trim(strOpSymbol(6)) <> "B" Or _
                               Trim(strOpSymbol(13)) = "S" Or _
                               Trim(strOpSymbol(13)) = "T" Then
                                'ポートクッションニードル位置指定を使用不可
                                Me.strcDataInfo.intPortCushion = CST_DISP_UNENABLE
                            End If
                            'ポート２箇所を非表示
                            Me.strcDataInfo.intPort = CST_DISP_HIDE

                            'ポートサイズダウンを非表示
                            Me.strcDataInfo.intPortSize = CST_DISP_HIDE

                            '支持金具回転
                            Select Case strOpSymbol(3)
                                Case "LB", "FB", "CA", "CB", "TA", "TB", "TC", "TD", "TE", "FA", "TF"
                                Case Else
                                    Me.strcDataInfo.intMounting = CST_DISP_UNENABLE
                            End Select

                            'トラニオン位置指定を使用不可
                            If Trim(strOpSymbol(3)) <> "TC" And _
                               Trim(strOpSymbol(3)) <> "TF" Then
                                Me.strcDataInfo.intTrunnion = CST_DISP_UNENABLE
                            End If

                            'タイロッド延長寸法を非表示
                            Me.strcDataInfo.intTieRod = CST_DISP_UNENABLE

                            'タイロッド材質ＳＵＳを使用不可
                            If InStr(strOpSymbol(1), "K") <> 0 Then
                                Me.strcDataInfo.intSUS = CST_DISP_UNENABLE
                            End If

                            'ピストンロッドはジャバラ付寸法でジャバラなし
                            If InStr(strOpSymbol(1), "O") <> 0 Or InStr(strOpSymbol(1), "Q2") <> 0 Or _
                               InStr(strOpSymbol(1), "U") <> 0 Or InStr(strOpSymbol(1), "G") <> 0 Or _
                               InStr(strOpSymbol(13), "J") <> 0 Or InStr(strOpSymbol(13), "K") <> 0 Or _
                               InStr(strOpSymbol(13), "L") <> 0 Then
                                Me.strcDataInfo.intJM = CST_DISP_UNENABLE
                            End If
                        Case "V"
                            'ポートクッションニードル位置指定を非表示
                            Me.strcDataInfo.intPortCushion = CST_DISP_HIDE
                            'ポート２箇所を非表示
                            Me.strcDataInfo.intPort = CST_DISP_HIDE
                            'ポートサイズダウンを非表示
                            Me.strcDataInfo.intPortSize = CST_DISP_HIDE
                            ''支持金具回転は使用不可

                            'トラニオン位置指定は使用可
                            If Trim(strOpSymbol(3)) <> "TC" And _
                               Trim(strOpSymbol(3)) <> "TF" Then
                                Me.strcDataInfo.intTrunnion = CST_DISP_UNENABLE
                            End If

                            'タイロッド延長寸法を非表示
                            Me.strcDataInfo.intTieRod = CST_DISP_UNENABLE
                            'タイロッド材質ＳＵＳ
                            If InStr(strOpSymbol(1), "K") <> 0 Then
                                Me.strcDataInfo.intSUS = CST_DISP_UNENABLE
                            End If
                            'ピストンロッドはジャバラ付寸法でジャバラなし
                            If InStr(strOpSymbol(1), "G") <> 0 Or InStr(strOpSymbol(1), "G1") <> 0 Or _
                               InStr(strOpSymbol(1), "G4") <> 0 Or InStr(strOpSymbol(13), "J") <> 0 Or _
                               InStr(strOpSymbol(13), "L") <> 0 Or InStr(strOpSymbol(13), "K") <> 0 Then
                                Me.strcDataInfo.intJM = CST_DISP_UNENABLE
                            End If

                            'スクレーパ、ロッドパッキンのみフッ素ゴム
                            If InStr(strOpSymbol(1), "G") <> 0 Or InStr(strOpSymbol(1), "G1") <> 0 Or _
                               InStr(strOpSymbol(1), "G4") <> 0 Then
                                Me.strcDataInfo.intFluoroRub = CST_DISP_UNENABLE
                            End If
                    End Select
                Case "SCS"
                    '画像ファイル指定
                    With Me.strcImagePathInfo
                        .strPortPath = "../KHImage/outOp1.gif"
                        .strPortExePath = "../KHImage/outOp2.gif"
                        .strMountingPath = "../KHImage/outOp3.gif"
                        .strTrunnionPath = "../KHImage/outOp4.gif"
                        .strTieRodPath = "../KHImage/outOp5_SCS.gif"
                    End With

                    Select Case Me.strcSelection.strKeyKataban
                        '基本形の場合
                        Case ""
                            '引当画面.オプションに"S"を指定
                            If InStr(strOpSymbol(17), "S") > 0 Then
                                'ポートクッションニードル位置指定を使用不可
                                Me.strcDataInfo.intPortCushion = CST_DISP_UNENABLE
                                'ポート２箇所指定を使用不可
                                Me.strcDataInfo.intPort = CST_DISP_UNENABLE

                                '引当画面.オプションに"T"を指定
                            ElseIf InStr(strOpSymbol(17), "T") > 0 Then
                                'ポートクッションニードル位置指定を使用不可
                                Me.strcDataInfo.intPortCushion = CST_DISP_UNENABLE
                            End If

                        Case "B"
                            'ポートクッションニードル位置指定を使用不可
                            Me.strcDataInfo.intPortCushion = CST_DISP_UNENABLE
                            'ポート２箇所指定を使用不可
                            Me.strcDataInfo.intPort = CST_DISP_UNENABLE

                        Case "D"
                            '引当画面.オプションに"S"か"T"を指定
                            If InStr(strOpSymbol(17), "S") > 0 OrElse InStr(strOpSymbol(17), "T") > 0 Then
                                'ポートクッションニードル位置指定を使用不可
                                Me.strcDataInfo.intPortCushion = CST_DISP_UNENABLE
                            End If
                    End Select
                    '引当画面.オプションに"J"か"K"か"L"を指定
                    If InStr(strOpSymbol(17), "J") > 0 OrElse InStr(strOpSymbol(17), "K") > 0 _
                    OrElse InStr(strOpSymbol(17), "L") > 0 Then
                        'ジャバラ指定を使用不可
                        Me.strcDataInfo.intJM = CST_DISP_UNENABLE
                    End If

                    '引当画面.バリエーションが"T"を含む
                    If InStr(strOpSymbol(1), "T") > 0 Then
                        'フッ素ゴム指定を使用不可
                        Me.strcDataInfo.intFluoroRub = CST_DISP_UNENABLE
                    End If
                Case "SCS2"
                    '画像ファイル指定
                    With Me.strcImagePathInfo
                        .strPortPath = "../KHImage/outOp1.gif"
                        .strPortExePath = "../KHImage/outOp2.gif"
                        .strMountingPath = "../KHImage/outOp3.gif"
                        .strTrunnionPath = "../KHImage/outOp4.gif"
                        .strTieRodPath = "../KHImage/outOp5_SCS.gif"
                    End With

                    Select Case Me.strcSelection.strKeyKataban
                        '基本形の場合
                        Case "", "F"
                            'ポート２箇所を使用不可
                            Me.strcDataInfo.intPort = CST_DISP_UNENABLE

                            If InStr(strOpSymbol(18), "S") > 0 Or _
                               InStr(strOpSymbol(18), "T") > 0 Then
                                'ポートクッションニードル位置指定を使用不可
                                Me.strcDataInfo.intPortCushion = CST_DISP_UNENABLE
                            End If

                            If InStr(strOpSymbol(18), "S") > 0 Then
                                'ポート２箇所を使用不可
                                Me.strcDataInfo.intPort = CST_DISP_UNENABLE
                            End If

                        Case "B"
                            'ポートクッションニードル位置指定を使用不可
                            Me.strcDataInfo.intPortCushion = CST_DISP_UNENABLE
                            'ポート２箇所指定を使用不可
                            Me.strcDataInfo.intPort = CST_DISP_UNENABLE

                        Case "D", "G"
                            'ポート２箇所指定を使用不可
                            Me.strcDataInfo.intPort = CST_DISP_UNENABLE
                            If InStr(strOpSymbol(18), "S") > 0 Or _
                               InStr(strOpSymbol(18), "T") > 0 Then
                                'ポートクッションニードル位置指定を使用不可
                                Me.strcDataInfo.intPortCushion = CST_DISP_UNENABLE
                            End If
                    End Select

                    '引当画面.オプションに"J"か"K"か"L"を指定
                    If InStr(strOpSymbol(18), "J") > 0 OrElse InStr(strOpSymbol(18), "K") > 0 _
                    OrElse InStr(strOpSymbol(18), "L") > 0 Then
                        'ジャバラ指定を使用不可
                        Me.strcDataInfo.intJM = CST_DISP_UNENABLE
                    End If

                    '引当画面.バリエーションが"T"を含む
                    If InStr(strOpSymbol(1), "T") > 0 Then
                        'フッ素ゴム指定を使用不可
                        Me.strcDataInfo.intFluoroRub = CST_DISP_UNENABLE
                    End If
                Case "JSC3"
                    Dim isSUS As Boolean = True
                    Dim isJM As Boolean = True
                    Select Case Me.strcSelection.strKeyKataban
                        '40-100
                        Case "1"
                            '画像ファイル指定
                            With Me.strcImagePathInfo
                                .strPortPath = "../KHImage/outOp1.gif"
                                .strPortExePath = "../KHImage/outOp2.gif"
                                .strMountingPath = "../KHImage/outOp3.gif"
                                .strTrunnionPath = "../KHImage/outOp4.gif"
                                .strTieRodPath = "../KHImage/outOp5_JSC3_1.gif"
                            End With

                            '引当画面.バリエーションが"K"を含む場合
                            If InStr(strOpSymbol(2), "K") > 0 Then
                                isSUS = False
                            End If

                            '引当画面.バリエーションが"G"を含む場合、もしくは、オプション'J','K'を含む場合
                            If InStr(strOpSymbol(2), "G") > 0 _
                                   OrElse InStr(strOpSymbol(13), "J") > 0 OrElse InStr(strOpSymbol(13), "K") > 0 Then
                                isJM = False
                            End If
                            '125-180
                        Case "2"
                            '画像ファイル指定
                            With Me.strcImagePathInfo
                                .strPortPath = "../KHImage/outOp1.gif"
                                .strPortExePath = "../KHImage/outOp2.gif"
                                .strMountingPath = "../KHImage/outOp3.gif"
                                .strTrunnionPath = "../KHImage/outOp4.gif"
                                .strTieRodPath = "../KHImage/outOp5_JSC3_2.gif"
                            End With
                            '引当画面.バリエーションが"G"を含む場合、もしくは、オプション'J','K','L'を含む場合
                            If InStr(strOpSymbol(2), "G") > 0 OrElse InStr(strOpSymbol(13), "J") > 0 _
                                   OrElse InStr(strOpSymbol(13), "K") > 0 OrElse InStr(strOpSymbol(13), "L") > 0 Then
                                isJM = False
                            End If
                    End Select
                    '引当画面.支持形式
                    Me.strcDataInfo.intTrunnion = CST_DISP_UNENABLE
                    Select Case strOpSymbol(4)
                        Case "FB", "CA", "CB"
                            Me.strcDataInfo.intTieRod = CST_DISP_UNENABLE
                        Case "TC", "TF"
                            Me.strcDataInfo.intTrunnion = CST_DISP_ENABLE
                    End Select
                    'SUS
                    If Not isSUS Then Me.strcDataInfo.intSUS = CST_DISP_UNENABLE
                    'ジャバラなし
                    If Not isJM Then Me.strcDataInfo.intJM = CST_DISP_UNENABLE
                Case "JSC4"
                    Dim isSUS As Boolean = True
                    Dim isJM As Boolean = True
                    Select Case Me.strcSelection.strKeyKataban
                            '125-180
                        Case "2"
                            '画像ファイル指定
                            With Me.strcImagePathInfo
                                .strPortPath = "../KHImage/outOp1.gif"
                                .strPortExePath = "../KHImage/outOp2.gif"
                                .strMountingPath = "../KHImage/outOp3.gif"
                                .strTrunnionPath = "../KHImage/outOp4.gif"
                                .strTieRodPath = "../KHImage/outOp5_JSC3_2.gif"
                            End With
                            '引当画面.バリエーションが"G"を含む場合、もしくは、オプション'J','K','L'を含む場合
                            If InStr(strOpSymbol(2), "G") > 0 OrElse InStr(strOpSymbol(13), "J") > 0 _
                                   OrElse InStr(strOpSymbol(13), "K") > 0 OrElse InStr(strOpSymbol(13), "L") > 0 Then
                                isJM = False
                            End If
                    End Select
                    '引当画面.支持形式
                    Me.strcDataInfo.intTrunnion = CST_DISP_UNENABLE
                    Select Case strOpSymbol(4)
                        Case "FB", "CA", "CB"
                            Me.strcDataInfo.intTieRod = CST_DISP_UNENABLE
                        Case "TC", "TF"
                            Me.strcDataInfo.intTrunnion = CST_DISP_ENABLE
                    End Select
                    'SUS
                    If Not isSUS Then Me.strcDataInfo.intSUS = CST_DISP_UNENABLE
                    'ジャバラなし
                    If Not isJM Then Me.strcDataInfo.intJM = CST_DISP_UNENABLE
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 画面表示情報検索取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="dtResult">取得結果</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function subOutofOpDataSelect(ByVal objCon As SqlConnection, ByRef dtResult As DataTable) As Boolean
        Dim dalOutOfOption As New OutOfOptionDAL

        Try
            dtResult = dalOutOfOption.fncOutofOpDataSelect(objCon, Me.strcSelection.strLang, Me.strcSelection.strSeriesKataban, _
                                                        Me.strcSelection.strKeyKataban, Me.strcSelection.strBoreSize)

            subOutofOpDataSelect = True
        Catch ex As Exception
            subOutofOpDataSelect = False
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 引当オプション外特注取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks>引当オプション外特注情報を取得し、メンバ変数にセットする</remarks>
    Private Sub subSelOutOfOpSelect(ByVal objCon As SqlConnection)
        Dim dt As New DataTable
        Dim dalOutOfOption As New OutOfOptionDAL

        Try
            dt = dalOutOfOption.fncSelOutOfOpSelect(objCon, Me.strcSelection.strUserID, Me.strcSelection.strSessionID)
            If dt.Rows.Count > 0 Then
                With Me.strcSelDataInfo
                    .intSelPortCushion = dt.Rows(0)("port_cushion")
                    .strSelPortCuPlace = dt.Rows(0)("port_cushion_place")
                    .intSelPort = dt.Rows(0)("port")
                    .intSelPortSize = dt.Rows(0)("port_size")
                    .intSelMounting = dt.Rows(0)("mounting")
                    .strSelTrunnion = dt.Rows(0)("trunnion")
                    .intSelClevis = dt.Rows(0)("clevis")
                    .strSelTieRodRadio = dt.Rows(0)("tierod_radio")
                    .intSelTieRodDefl = dt.Rows(0)("tierod_default")
                    .strSelTieRodCstm = dt.Rows(0)("tierod_custom")
                    .intSelSUS = dt.Rows(0)("sus")
                    .intSelJM = dt.Rows(0)("jm")
                    .intSelFluoroRub = dt.Rows(0)("fluororub")
                End With
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当オプション外特注テーブル削除処理
    ''' </summary>
    ''' <param name="objCon">DB接続オブジェクト</param>
    ''' <returns></returns>
    ''' <remarks>引当オプション外特注テーブルからデータを削除する</remarks>
    Public Function fncSPSelOutOpDel(ByVal objCon As SqlConnection) As Boolean
        Dim dalOutOfOption As New OutOfOptionDAL

        fncSPSelOutOpDel = False
        Try
            If dalOutOfOption.fncSPSelOutOpDel(objCon, Me.strcSelection.strUserID, _
                                               Me.strcSelection.strSessionID) Then
                fncSPSelOutOpDel = True
            Else
                fncSPSelOutOpDel = False
            End If
            
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 半角チェック
    ''' </summary>
    ''' <param name="strChk">チェック対象</param>
    ''' <param name="strErrCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncHankakuCheck(ByVal strChk As String, ByRef strErrCd As String) As Boolean
        Dim sjisEnc As Encoding = Encoding.GetEncoding("Shift_JIS")
        fncHankakuCheck = True
        Try
            Dim num As Integer = sjisEnc.GetByteCount(strChk)
            '桁数比較
            If num = strChk.Length Then
                Return True
            Else
                'エラーメッセージ設定
                strErrCd = "W0920"
                Return False
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 入力チェック(数値が入っているかどうかをチェックする)
    ''' </summary>
    ''' <param name="strChk">チェック対象</param>
    ''' <param name="strErrCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncNumericCheck(ByVal strChk As String, ByRef strErrCd As String) As Boolean
        Dim intLoopCnt As Integer
        Dim lenInt As Integer = 0
        Dim lenDec As Integer = 0
        fncNumericCheck = False
        Try
            'チェック対象文字が空値の場合、正常終了
            If strChk.Length = 0 Then
                Return True
            End If
            'エラーメッセージ設定
            strErrCd = "W0920"
            '桁数算出
            If InStr(strChk, ".") > 0 Then
                '小数あり
                lenInt = InStr(strChk, ".") - 1
                lenDec = strChk.Length - (lenInt + 1)
            Else
                '小数なし
                lenInt = strChk.Length
            End If

            '整数部チェック
            For intLoopCnt = 1 To lenInt
                '数値のみOK
                If Mid(strChk.Trim, intLoopCnt, 1) = "0" Or _
                   Mid(strChk.Trim, intLoopCnt, 1) = "1" Or _
                   Mid(strChk.Trim, intLoopCnt, 1) = "2" Or _
                   Mid(strChk.Trim, intLoopCnt, 1) = "3" Or _
                   Mid(strChk.Trim, intLoopCnt, 1) = "4" Or _
                   Mid(strChk.Trim, intLoopCnt, 1) = "5" Or _
                   Mid(strChk.Trim, intLoopCnt, 1) = "6" Or _
                   Mid(strChk.Trim, intLoopCnt, 1) = "7" Or _
                   Mid(strChk.Trim, intLoopCnt, 1) = "8" Or _
                   Mid(strChk.Trim, intLoopCnt, 1) = "9" Then
                    'Do Nothing
                Else
                    '数値以外、エラー
                    Return False
                End If
            Next
            'エラーメッセージ設定
            strErrCd = "W0890"
            '小数部
            Select Case lenDec
                Case 0
                    'Do Nothing
                Case 1
                    '小数点以下に0,5以外は、エラー
                    If Mid(strChk, lenInt + 2) = "0" _
                    OrElse Mid(strChk, lenInt + 2) = "5" Then
                        'Do Nothing
                    Else
                        Return False
                    End If
                Case Else
                    '小数点以下１桁のみ有効
                    Return False
            End Select
            '正常
            strErrCd = ""
            Return True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 引当情報更新
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <remarks>入力データで引当オプション外特注テーブル/引当シリーズ形番を更新する</remarks>
    Public Sub subUpdateSelOutOp(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc)
        Dim bolReturn As Boolean
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            'フル形番クリア
            Me.strcSelection.strFullKtbn = CST_BLANK
            '引当オプション外特注クリア
            bolReturn = fncSPSelOutOpDel(objCon)
            '引当オプション外特注注更新
            bolReturn = fncSPSelOutOpIns(objCon)
            'フル形番生成
            Call subFullKtbnCreate(objCon)
            '引当シリーズ形番更新(オプション情報)
            Call dalKtbnStrc.subSelSrsKtbnOptionUpd(objCon, Me.strcSelection.strUserID, _
                                                    Me.strcSelection.strSessionID, _
                                                    objKtbnStrc.strcSelection.strRodEndOption, _
                                                    Me.strcSelection.strFullKtbn)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当オプション外特注テーブル追加処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSPSelOutOpIns(ByVal objCon As SqlConnection) As Boolean
        Dim objCmd As SqlCommand = Nothing
        fncSPSelOutOpIns = False
        Try
            objCmd = objCon.CreateCommand
            With objCmd
                .CommandType = CommandType.StoredProcedure
                .CommandText = CdCst.DB.SPL.KHSelOutOfOpIns
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10)
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88)
                .Parameters.Add("@Port_cushion", SqlDbType.Int)
                .Parameters.Add("@Port_cushion_place", SqlDbType.VarChar, 4)
                .Parameters.Add("@Port", SqlDbType.Int)
                .Parameters.Add("@Port_size", SqlDbType.Int)
                .Parameters.Add("@Mounting", SqlDbType.Int)
                .Parameters.Add("@Trunnion", SqlDbType.VarChar, 50)
                .Parameters.Add("@Clevis", SqlDbType.Int)
                .Parameters.Add("@Tierod_radio", SqlDbType.VarChar, 1)
                .Parameters.Add("@Tierod_default", SqlDbType.Int)
                .Parameters.Add("@Tierod_custom", SqlDbType.VarChar, 30)
                .Parameters.Add("@Sus", SqlDbType.Int)
                .Parameters.Add("@JM", SqlDbType.Int)
                .Parameters.Add("@FluoroRub", SqlDbType.Int)
                .Parameters.Add("@Place_lvl", SqlDbType.Int)     '2017/04/10 追加
                .Parameters.Add("@RegPerson", SqlDbType.VarChar, 10)
                .Parameters.Add("@RegDate", SqlDbType.DateTime, 88)
                .Parameters.Add("@CurPerson", SqlDbType.VarChar, 10)
                .Parameters.Add("@CurDate", SqlDbType.DateTime, 88)

                .Parameters("@UserId").Value = Me.strcSelection.strUserID
                .Parameters("@SessionId").Value = Me.strcSelection.strSessionID
                .Parameters("@Port_cushion").Value = Me.strcSelDataInfo.intSelPortCushion
                .Parameters("@Port_cushion_place").Value = Me.strcSelDataInfo.strSelPortCuPlace.Trim
                .Parameters("@Port").Value = Me.strcSelDataInfo.intSelPort
                .Parameters("@Port_size").Value = Me.strcSelDataInfo.intSelPortSize
                .Parameters("@Mounting").Value = Me.strcSelDataInfo.intSelMounting
                .Parameters("@Trunnion").Value = Me.strcSelDataInfo.strSelTrunnion.Trim
                .Parameters("@Clevis").Value = Me.strcSelDataInfo.intSelClevis
                .Parameters("@Tierod_radio").Value = Me.strcSelDataInfo.strSelTieRodRadio
                .Parameters("@Tierod_default").Value = Me.strcSelDataInfo.intSelTieRodDefl
                .Parameters("@Tierod_custom").Value = Me.strcSelDataInfo.strSelTieRodCstm
                .Parameters("@Sus").Value = Me.strcSelDataInfo.intSelSUS
                .Parameters("@JM").Value = Me.strcSelDataInfo.intSelJM
                .Parameters("@FluoroRub").Value = Me.strcSelDataInfo.intSelFluoroRub
                .Parameters("@Place_lvl").Value = Me.strcSelDataInfo.intPlacelvl    '2107/04/10 追加
                .Parameters("@RegPerson").Value = Me.strcSelection.strUserID
                .Parameters("@RegDate").Value = Now()
                .Parameters("@CurPerson").Value = DBNull.Value
                .Parameters("@CurDate").Value = DBNull.Value
                '実行
                objCmd.ExecuteNonQuery()
            End With
            fncSPSelOutOpIns = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objCmd IsNot Nothing Then
                objCmd.Dispose()
                objCmd = Nothing
            End If
        End Try
    End Function

    ''' <summary>
    ''' フル形番生成
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks>入力したオプション外特注データよりフル形番を生成する</remarks>
    Private Sub subFullKtbnCreate(ByVal objCon As SqlConnection)
        Dim dtResult As New DataTable
        Dim sbFullKata As New StringBuilder
        Dim dalOutOfOption As New OutOfOptionDAL

        Try
            dtResult = dalOutOfOption.fncFullKtbnCreate(objCon, Me.strcSelection.strUserID, Me.strcSelection.strSessionID, Me.strcSelection.strBoreSize)
            'フル形番生成
            If dtResult.Rows.Count <> 0 Then
                'ポート・クッションニードル位置
                If dtResult.Rows(0).Item("port_cushion") = 1 Then
                    If Not (dtResult.Rows(0).Item("port_cushion_place").Equals("1212")) Then
                        sbFullKata.Append("R")
                        sbFullKata.Append(dtResult.Rows(0).Item("port_cushion_place"))
                    End If
                End If
                'ポート２箇所
                If dtResult.Rows(0).Item("port") > 0 Then sbFullKata.Append(dtResult.Rows(0).Item("K_port").ToString.Trim)
                'ポートサイズ
                If dtResult.Rows(0).Item("port_size") > 0 Then sbFullKata.Append(dtResult.Rows(0).Item("K_portSize").ToString.Trim)
                '支持金具
                If dtResult.Rows(0).Item("mounting") > 0 Then sbFullKata.Append(dtResult.Rows(0).Item("K_mounting").ToString.Trim)
                'トラニオン位置指定
                If dtResult.Rows(0).Item("trunnion").ToString.Length > 0 Then
                    sbFullKata.Append("AQ")
                    sbFullKata.Append(dtResult.Rows(0).Item("trunnion").ToString.Trim)
                End If
                '二山ナックル・二山クレビス
                If dtResult.Rows(0).Item("clevis") > 0 Then sbFullKata.Append(dtResult.Rows(0).Item("K_clevis").ToString.Trim)
                'タイロッド寸法
                If dtResult.Rows(0).Item("tierod_radio").ToString.Length > 0 _
                AndAlso dtResult.Rows(0).Item("tierod_custom").ToString.Length > 0 Then
                    'シリーズごとに変換
                    If Me.strcSelection.strSeriesKataban.Equals("JSC3") _
                    AndAlso Me.strcSelection.strKeyKataban.Equals("1") Then
                        sbFullKata.Append("MM")
                        sbFullKata.Append(Me.strcDataInfo.lstTieRodCstm.Item( _
                                            CInt(dtResult.Rows(0).Item("tierod_custom").ToString)).ToString)
                    Else
                        sbFullKata.Append("MX")
                        sbFullKata.Append(dtResult.Rows(0).Item("tierod_custom").ToString.Trim)
                    End If
                    Select Case dtResult.Rows(0).Item("tierod_radio")
                        Case "2"
                            sbFullKata.Append("R")
                        Case "3"
                            sbFullKata.Append("R1")
                        Case "4"
                            sbFullKata.Append("R2")
                        Case "5"
                            sbFullKata.Append("H")
                        Case "6"
                            sbFullKata.Append("H1")
                        Case "7"
                            sbFullKata.Append("H2")
                    End Select
                End If
                'タイロッド材質
                If dtResult.Rows(0).Item("sus") > 0 Then sbFullKata.Append("M1")
                'ピストンロッド
                If dtResult.Rows(0).Item("jm") > 0 Then sbFullKata.Append("J9")
                'スクレーバ、ロッドパッキン
                If dtResult.Rows(0).Item("fluororub") > 0 Then sbFullKata.Append("T9")
            End If
            Me.strcSelection.strFullKtbn = sbFullKata.ToString
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

#Region " Property "

    '**********************************************************************************************
    '*【プロパティ】BoreSize
    '*  ロッド先端パターン記号の設定・取得
    '**********************************************************************************************
    Public Property BoreSize() As String
        Get
            Return Me.strcSelection.strBoreSize
        End Get
        Set(ByVal value As String)
            Me.strcSelection.strBoreSize = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】PortPath
    '*  ポート・クッションニードル位置用イメージパスの設定・取得
    '**********************************************************************************************
    Public Property PortPath() As String
        Get
            Return Me.strcImagePathInfo.strPortPath
        End Get
        Set(ByVal value As String)
            Me.strcImagePathInfo.strPortPath = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】PortPath
    '*  ポート・クッションニードル位置例用イメージパスの設定・取得
    '**********************************************************************************************
    Public Property PortPathExe() As String
        Get
            Return Me.strcImagePathInfo.strPortExePath
        End Get
        Set(ByVal value As String)
            Me.strcImagePathInfo.strPortExePath = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】MountingPath
    '*  支持金具回転用イメージパスの設定・取得
    '**********************************************************************************************
    Public Property MountingPath() As String
        Get
            Return Me.strcImagePathInfo.strMountingPath
        End Get
        Set(ByVal value As String)
            Me.strcImagePathInfo.strMountingPath = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】TrunnionPath
    '*  トラニオン位置用イメージパスの設定・取得
    '**********************************************************************************************
    Public Property TrunnionPath() As String
        Get
            Return Me.strcImagePathInfo.strTrunnionPath
        End Get
        Set(ByVal value As String)
            Me.strcImagePathInfo.strTrunnionPath = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】TieRodPath
    '*  タイロッド延長寸法用イメージパスの設定・取得
    '**********************************************************************************************
    Public Property TieRodPath() As String
        Get
            Return Me.strcImagePathInfo.strTieRodPath
        End Get
        Set(ByVal value As String)
            Me.strcImagePathInfo.strTieRodPath = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】isPortCushion
    '*  ポート・クッションニードル位置表示状態(-1:非表示、0:使用不可,1:使用可)の設定・取得
    '**********************************************************************************************
    Public Property intPortCushion() As Integer
        Get
            Return Me.strcDataInfo.intPortCushion
        End Get
        Set(ByVal value As Integer)
            Me.strcDataInfo.intPortCushion = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】PortCushion
    '*  ポート・クッションニードル位置リストの設定・取得
    '**********************************************************************************************
    Public Property PortCushion() As DataTable
        Get
            Return Me.strcDataInfo.lstPortCushion
        End Get
        Set(ByVal value As DataTable)
            Me.strcDataInfo.lstPortCushion = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】isPort
    '*  ポート２箇所表示状態(-1:非表示、0:使用不可,1:使用可)の設定・取得
    '**********************************************************************************************
    Public Property isPort() As Integer
        Get
            Return Me.strcDataInfo.intPort
        End Get
        Set(ByVal value As Integer)
            Me.strcDataInfo.intPort = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】Port
    '*  ポート２箇所リストの設定・取得
    '**********************************************************************************************
    Public Property Port() As DataTable
        Get
            Return Me.strcDataInfo.lstPort
        End Get
        Set(ByVal value As DataTable)
            Me.strcDataInfo.lstPort = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】isPortSize
    '* ポートサイズダウン表示状態(-1:非表示、0:使用不可,1:使用可)の設定・取得
    '**********************************************************************************************
    Public Property isPortSize() As Integer
        Get
            Return Me.strcDataInfo.intPortSize
        End Get
        Set(ByVal value As Integer)
            Me.strcDataInfo.intPortSize = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】PortSize
    '*  ポートサイズダウンリストの設定・取得
    '**********************************************************************************************
    Public Property PortSize() As DataTable
        Get
            Return Me.strcDataInfo.lstPortSize
        End Get
        Set(ByVal value As DataTable)
            Me.strcDataInfo.lstPortSize = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】isMounting
    '*  支持金具回転表示状態(-1:非表示、0:使用不可,1:使用可)の設定・取得
    '**********************************************************************************************
    Public Property isMounting() As Integer
        Get
            Return Me.strcDataInfo.intMounting
        End Get
        Set(ByVal value As Integer)
            Me.strcDataInfo.intMounting = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】Mounting
    '*  支持金具回転リストの設定・取得
    '**********************************************************************************************
    Public Property Mounting() As DataTable
        Get
            Return Me.strcDataInfo.lstMounting
        End Get
        Set(ByVal value As DataTable)
            Me.strcDataInfo.lstMounting = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】isTrunnion
    '*  トラニオン位置表示状態(-1:非表示、0:使用不可,1:使用可)の設定・取得
    '**********************************************************************************************
    Public Property isTrunnion() As Integer
        Get
            Return Me.strcDataInfo.intTrunnion
        End Get
        Set(ByVal value As Integer)
            Me.strcDataInfo.intTrunnion = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】isClevis
    '*  二山ナックル・二山クレビス表示状態(-1:非表示、0:使用不可,1:使用可)の設定・取得
    '**********************************************************************************************
    Public Property isClevis() As Integer
        Get
            Return Me.strcDataInfo.intClevis
        End Get
        Set(ByVal value As Integer)
            Me.strcDataInfo.intClevis = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】Clevis
    '*  二山ナックル・二山クレビスリストの設定・取得
    '**********************************************************************************************
    Public Property Clevis() As DataTable
        '全体をdatatableに置き換え
        Get
            Return Me.strcDataInfo.lstClevis
        End Get
        Set(ByVal value As DataTable)
            Me.strcDataInfo.lstClevis = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】isTieRod
    '*  タイロッド延長寸法表示状態(-1:非表示、0:使用不可,1:使用可)の設定・取得
    '**********************************************************************************************
    Public Property isTieRod() As Integer
        Get
            Return Me.strcDataInfo.intTieRod
        End Get
        Set(ByVal value As Integer)
            Me.strcDataInfo.intTieRod = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】TieRodRadio
    '*  タイロッド延長寸法ラジオボタンリストの設定・取得
    '**********************************************************************************************
    Public Property TieRodRadio() As ArrayList
        Get
            Return Me.strcDataInfo.lstTieRodRadio
        End Get
        Set(ByVal value As ArrayList)
            Me.strcDataInfo.lstTieRodRadio = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】TieRodDefl
    '*  タイロッド標準寸法の設定・取得
    '**********************************************************************************************
    Public Property TieRodDefl() As String
        Get
            Return Me.strcDataInfo.strTieRodDefl
        End Get
        Set(ByVal value As String)
            Me.strcDataInfo.strTieRodDefl = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】TieRodCstm
    '*  タイロッド特注寸法の設定・取得
    '**********************************************************************************************
    Public Property TieRodCstm() As ArrayList
        Get
            Return Me.strcDataInfo.lstTieRodCstm
        End Get
        Set(ByVal value As ArrayList)
            Me.strcDataInfo.lstTieRodCstm = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】isSUS
    '*  タイロッド材質SUS表示状態(-1:非表示、0:使用不可,1:使用可)の設定・取得
    '**********************************************************************************************
    Public Property isSUS() As Integer
        Get
            Return Me.strcDataInfo.intSUS
        End Get
        Set(ByVal value As Integer)
            Me.strcDataInfo.intSUS = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SUS
    '*  タイロッド材質SUSリストの設定・取得
    '**********************************************************************************************
    Public Property SUS() As DataTable
        Get
            Return Me.strcDataInfo.lstSUS
        End Get
        Set(ByVal value As DataTable)
            Me.strcDataInfo.lstSUS = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】isJM
    '*  ジャバラ表示状態(-1:非表示、0:使用不可,1:使用可)の設定・取得
    '**********************************************************************************************
    Public Property isJM() As Integer
        Get
            Return Me.strcDataInfo.intJM
        End Get
        Set(ByVal value As Integer)
            Me.strcDataInfo.intJM = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】JM
    '*  ジャバラリストの設定・取得
    '**********************************************************************************************
    Public Property JM() As DataTable
        Get
            Return Me.strcDataInfo.lstJM
        End Get
        Set(ByVal value As DataTable)
            Me.strcDataInfo.lstJM = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】isFluoroRub
    '*  フッ素ゴム表示状態(-1:非表示、0:使用不可,1:使用可)の設定・取得
    '**********************************************************************************************
    Public Property isFluoroRub() As Integer
        Get
            Return Me.strcDataInfo.intFluoroRub
        End Get
        Set(ByVal value As Integer)
            Me.strcDataInfo.intFluoroRub = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】FluoroRub
    '*  フッ素ゴムの設定・取得
    '**********************************************************************************************
    Public Property FluoroRub() As DataTable
        Get
            Return Me.strcDataInfo.lstFluoroRub
        End Get
        Set(ByVal value As DataTable)
            Me.strcDataInfo.lstFluoroRub = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelPortCushion
    '*  選択ポート・クッションニードル位置の設定・取得
    '**********************************************************************************************
    Public Property SelPortCushion() As Integer
        Get
            Return Me.strcSelDataInfo.intSelPortCushion
        End Get
        Set(ByVal value As Integer)
            Me.strcSelDataInfo.intSelPortCushion = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelPortCuPlace
    '*  選択ポート・クッションニードル位置ラジオ連結の設定・取得
    '**********************************************************************************************
    Public Property SelPortCuPlace() As String
        Get
            Return Me.strcSelDataInfo.strSelPortCuPlace
        End Get
        Set(ByVal value As String)
            Me.strcSelDataInfo.strSelPortCuPlace = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelPort
    '*  選択ポート２箇所の設定・取得
    '**********************************************************************************************
    Public Property SelPort() As Integer
        Get
            Return Me.strcSelDataInfo.intSelPort
        End Get
        Set(ByVal value As Integer)
            Me.strcSelDataInfo.intSelPort = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelPortSize
    '*  選択ポートサイズダウンの設定・取得
    '**********************************************************************************************
    Public Property SelPortSize() As Integer
        Get
            Return Me.strcSelDataInfo.intSelPortSize
        End Get
        Set(ByVal value As Integer)
            Me.strcSelDataInfo.intSelPortSize = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelMounting
    '*  選択支持金具回転の設定・取得
    '**********************************************************************************************
    Public Property SelMounting() As Integer
        Get
            Return Me.strcSelDataInfo.intSelMounting
        End Get
        Set(ByVal value As Integer)
            Me.strcSelDataInfo.intSelMounting = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelTrunnion
    '*  選択トラニオン位置の設定・取得
    '**********************************************************************************************
    Public Property SelTrunnion() As String
        Get
            Return Me.strcSelDataInfo.strSelTrunnion
        End Get
        Set(ByVal value As String)
            Me.strcSelDataInfo.strSelTrunnion = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelClevis
    '*  選択二山ナックル・二山クレビスの設定・取得
    '**********************************************************************************************
    Public Property SelClevis() As Integer
        Get
            Return Me.strcSelDataInfo.intSelClevis
        End Get
        Set(ByVal value As Integer)
            Me.strcSelDataInfo.intSelClevis = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelTieRodRadio
    '*  選択タイロッド延長寸法ラジオボタンの設定・取得
    '**********************************************************************************************
    Public Property SelTieRodRadio() As String
        Get
            Return Me.strcSelDataInfo.strSelTieRodRadio
        End Get
        Set(ByVal value As String)
            Me.strcSelDataInfo.strSelTieRodRadio = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelTieRodDefl
    '*  選択タイロッド延長寸法標準寸法の設定・取得
    '**********************************************************************************************
    Public Property SelTieRodDefl() As Integer
        Get
            Return Me.strcSelDataInfo.intSelTieRodDefl
        End Get
        Set(ByVal value As Integer)
            Me.strcSelDataInfo.intSelTieRodDefl = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelTieRodCstm
    '*  選択タイロッド延長寸法特注寸法の設定・取得
    '**********************************************************************************************
    Public Property SelTieRodCstm() As String
        Get
            Return Me.strcSelDataInfo.strSelTieRodCstm
        End Get
        Set(ByVal value As String)
            Me.strcSelDataInfo.strSelTieRodCstm = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelSUS
    '*  選択タイロッド材質SUSの設定・取得
    '**********************************************************************************************
    Public Property SelSUS() As Integer
        Get
            Return Me.strcSelDataInfo.intSelSUS
        End Get
        Set(ByVal value As Integer)
            Me.strcSelDataInfo.intSelSUS = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelJM
    '*  選択ジャバラの設定・取得
    '**********************************************************************************************
    Public Property SelJM() As Integer
        Get
            Return Me.strcSelDataInfo.intSelJM
        End Get
        Set(ByVal value As Integer)
            Me.strcSelDataInfo.intSelJM = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelFluoroRub
    '*  選択フッ素ゴムの設定・取得
    '**********************************************************************************************
    Public Property SelFluoroRub() As Integer
        Get
            Return Me.strcSelDataInfo.intSelFluoroRub
        End Get
        Set(ByVal value As Integer)
            Me.strcSelDataInfo.intSelFluoroRub = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelPlacelvl
    '*  選択された項目から算出されたPlacelvlを設定・取得   2017/04/10 追加
    '**********************************************************************************************
    Public Property SelPlacelvl() As Integer
        Get
            Return Me.strcSelDataInfo.intPlacelvl
        End Get
        Set(ByVal value As Integer)
            Me.strcSelDataInfo.intPlacelvl = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ErrCd
    '*  エラーコードの設定・取得
    '**********************************************************************************************
    Public Property ErrCd() As String
        Get
            Return Me.strcErrInfo.strErrCd
        End Get
        Set(ByVal value As String)
            Me.strcErrInfo.strErrCd = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ErrFocus
    '*  エラーコントロールIDの設定・取得
    '**********************************************************************************************
    Public Property ErrFocus() As String
        Get
            Return Me.strcErrInfo.strErrFocus
        End Get
        Set(ByVal value As String)
            Me.strcErrInfo.strErrFocus = value
        End Set
    End Property

#End Region

End Class
