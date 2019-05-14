Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class KHRodEndCstm

#Region " Fixed Value "
    Private Const CST_BLANK As String = CdCst.Sign.Blank
    Private Const CST_COMMA As String = CdCst.Sign.Delimiter.Comma
#End Region

#Region " Definition "

    '引当形番情報
    Private Structure Selection
        Public strSeriesKataban As String                      'シリーズ形番
        Public strKeyKataban As String                         'キー形番
        Public strBoreSize As String                           '引当口径
        Public strRodFullKtbn As String                        'ロッドフル形番
        Public strUserID As String                             'ユーザID
        Public strSessionID As String                          'セッションID
    End Structure
    Private strcSelection As Selection

    'ロッド先端標準情報
    Private Structure RodPtnInfo
        Public strRodPtn As String                             'ロッド先端特注パターン記号
        Public strDispNo As String                             'ロッド先端特注表示順序
        Public strKHImageUrl As String                           'イメージURL
        Public strExtFrm() As String                           '外径種類
        Public strDispExtFrm() As String                       '表示外径種類
        Public strNormalVal() As String                        '標準寸法
        Public strActNormalVal() As String                     '実標準寸法
        Public strInputDiv() As String                         '入力区分
        Public strSltVal() As String                           '選択可能寸法
        Public strActSltVal() As String                        '実選択可能寸法
        Public strJsName() As String                           'javascript名
        Public strWFMaxVal As String                           'WF最大寸法
    End Structure
    Private strcRodPtnInfo() As RodPtnInfo

    'ロッド先端選択情報
    Private Structure SelDataInfo
        Public strSelBoreSize As String                         '接続口径
        Public strSelPtnNo As String                            'ロッド先端特注パターンNo.
        Public strSelPtn As String                              'ロッド先端特注パターン記号
        Public hshtSelVal As Hashtable                          '選択外径種類/特注寸法
        Public strSelWFMaxVal As String                         'WF最大寸法
        Public strSelOtherVal As String                         'その他寸法
    End Structure
    Private strcSelDataInfo As SelDataInfo

    'エラー情報
    Private Structure ErrInfo
        Public strErrCd As String                               'エラーコード
        Public strErrOption As String                              'エラーオプション
        Public strErrFocusNo As String                          'エラーフォーカスNo.
        Public strErrPtnNo As String                            'エラーパターンNo.
        Public strErrPtn As String                              'エラーパターン
    End Structure
    Private strcErrInfo As ErrInfo

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
    '*【プロパティ】ErrOption
    '*  エラーコードの設定・取得
    '**********************************************************************************************
    Public Property ErrOption() As String
        Get
            Return Me.strcErrInfo.strErrOption
        End Get
        Set(ByVal value As String)
            Me.strcErrInfo.strErrOption = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ErrFocusNo
    '*  エラーボックスNo.の設定・取得
    '**********************************************************************************************
    Public Property ErrFocusNo() As String
        Get
            Return Me.strcErrInfo.strErrFocusNo
        End Get
        Set(ByVal value As String)
            Me.strcErrInfo.strErrFocusNo = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ErrPtnNo
    '*  エラーパターンNo.の設定・取得
    '**********************************************************************************************
    Public Property ErrPtnNo() As String
        Get
            Return Me.strcErrInfo.strErrPtnNo
        End Get
        Set(ByVal value As String)
            Me.strcErrInfo.strErrPtnNo = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ErrPtn
    '*  エラーパターン記号の設定・取得
    '**********************************************************************************************
    Public Property ErrPtn() As String
        Get
            Return Me.strcErrInfo.strErrPtn
        End Get
        Set(ByVal value As String)
            Me.strcErrInfo.strErrPtn = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】RodPtnCnt
    '*  ロッド先端特注パターン数の設定
    '**********************************************************************************************
    Public Property RodPtnCnt() As Integer
        Get
            Return Me.strcRodPtnInfo.Length
        End Get
        Set(ByVal value As Integer)

        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】RodPtn
    '*  ロッド先端パターン記号の設定・取得
    '**********************************************************************************************
    Public Property RodPtn(ByVal intPatternSeq As Integer) As String
        Get
            Return Me.strcRodPtnInfo(intPatternSeq).strRodPtn
        End Get
        Set(ByVal value As String)
            Me.strcRodPtnInfo(intPatternSeq).strRodPtn = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】KHImageUrl
    '*  イメージＵＲＬの設定・取得
    '**********************************************************************************************
    Public Property KHImageUrl(ByVal intPatternSeq As Integer) As String
        Get
            Return Me.strcRodPtnInfo(intPatternSeq).strKHImageUrl
        End Get
        Set(ByVal value As String)
            Me.strcRodPtnInfo(intPatternSeq).strKHImageUrl = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ExtFrm
    '*  外径種類の設定・取得
    '**********************************************************************************************
    Public Property ExtFrm(ByVal intPatternSeq As Integer) As String()
        Get
            Return Me.strcRodPtnInfo(intPatternSeq).strExtFrm
        End Get
        Set(ByVal value As String())
            Me.strcRodPtnInfo(intPatternSeq).strExtFrm = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】DispExtFrm
    '*  表示外径種類の設定・取得
    '**********************************************************************************************
    Public Property DispExtFrm(ByVal intPatternSeq As Integer) As String()
        Get
            Return Me.strcRodPtnInfo(intPatternSeq).strDispExtFrm
        End Get
        Set(ByVal value As String())
            Me.strcRodPtnInfo(intPatternSeq).strDispExtFrm = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】NormalVal
    '*  標準寸法の設定・取得
    '**********************************************************************************************
    Public Property NormalVal(ByVal intPatternSeq As Integer) As String()
        Get
            Return Me.strcRodPtnInfo(intPatternSeq).strNormalVal
        End Get
        Set(ByVal value As String())
            Me.strcRodPtnInfo(intPatternSeq).strNormalVal = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ActNormalVal
    '*  実標準寸法の設定・取得
    '**********************************************************************************************
    Public Property ActNormalVal(ByVal intPatternSeq As Integer) As String()
        Get
            Return Me.strcRodPtnInfo(intPatternSeq).strActNormalVal
        End Get
        Set(ByVal value As String())
            Me.strcRodPtnInfo(intPatternSeq).strActNormalVal = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】InputDiv
    '*  入力区分の設定・取得
    '**********************************************************************************************
    Public Property InputDiv(ByVal intPatternSeq As Integer) As String()
        Get
            Return Me.strcRodPtnInfo(intPatternSeq).strInputDiv
        End Get
        Set(ByVal value As String())
            Me.strcRodPtnInfo(intPatternSeq).strInputDiv = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SltVal
    '*  選択可能寸法の設定・取得
    '**********************************************************************************************
    Public Property SltVal(ByVal intPatternSeq As Integer) As String()
        Get
            Return Me.strcRodPtnInfo(intPatternSeq).strSltVal
        End Get
        Set(ByVal value As String())
            Me.strcRodPtnInfo(intPatternSeq).strSltVal = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】ActSltVal
    '*  実選択可能寸法の設定・取得
    '**********************************************************************************************
    Public Property ActSltVal(ByVal intPatternSeq As Integer) As String()
        Get
            Return Me.strcRodPtnInfo(intPatternSeq).strActSltVal
        End Get
        Set(ByVal value As String())
            Me.strcRodPtnInfo(intPatternSeq).strActSltVal = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】JsName
    '*  JavaScriptの設定・取得
    '**********************************************************************************************
    Public Property JsName(ByVal intPatternSeq As Integer) As String()
        Get
            Return Me.strcRodPtnInfo(intPatternSeq).strJsName
        End Get
        Set(ByVal value As String())
            Me.strcRodPtnInfo(intPatternSeq).strJsName = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelBoreSize
    '*  選択ロッド先端特注パターンNo.の設定・取得
    '**********************************************************************************************
    Public Property SelBoreSize() As String
        Get
            Return Me.strcSelDataInfo.strSelBoreSize
        End Get
        Set(ByVal value As String)
            Me.strcSelDataInfo.strSelBoreSize = value
        End Set
    End Property

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
    '*【プロパティ】SelPtn
    '*  選択ロッド先端特注パターン記号の設定・取得
    '**********************************************************************************************
    Public Property SelPtn() As String
        Get
            Return Me.strcSelDataInfo.strSelPtn
        End Get
        Set(ByVal value As String)
            Me.strcSelDataInfo.strSelPtn = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelOtherVal
    '*  選択その他寸法の設定・取得
    '**********************************************************************************************
    Public Property SelOtherVal() As String
        Get
            Return Me.strcSelDataInfo.strSelOtherVal
        End Get
        Set(ByVal value As String)
            Me.strcSelDataInfo.strSelOtherVal = value
        End Set
    End Property

    '**********************************************************************************************
    '*【プロパティ】SelVal
    '*  選択外径種類寸法の設定・取得
    '**********************************************************************************************
    Public Property SelValInfo() As Hashtable
        Get
            Return Me.strcSelDataInfo.hshtSelVal
        End Get
        Set(ByVal value As Hashtable)
            Me.strcSelDataInfo.hshtSelVal = value
        End Set
    End Property

#End Region

#Region " Method "

    ''' <summary>
    ''' フィールドの初期設定
    ''' </summary>
    ''' <param name="strAUserID">ユーザーID</param>
    ''' <param name="strASessionID">セッションID</param>
    ''' <param name="strASeriesKataban">シリーズ形番</param>
    ''' <param name="strAKeyKataban">キー形番</param>
    ''' <remarks>ユーザーID/セッションID/シリーズ形番/キー形番を保持</remarks>
    Public Sub New(ByVal strAUserID As String, _
                   ByVal strASessionID As String, _
                   ByVal strASeriesKataban As String, _
                   ByVal strAKeyKataban As String)

        'フィールド初期設定
        With Me.strcSelection
            .strSeriesKataban = CST_BLANK
            .strKeyKataban = CST_BLANK
            .strBoreSize = CST_BLANK
            .strRodFullKtbn = CST_BLANK
            .strUserID = CST_BLANK
            .strSessionID = CST_BLANK
        End With

        ReDim Me.strcRodPtnInfo(0)
        With Me.strcRodPtnInfo(0)
            .strRodPtn = CST_BLANK
            .strDispNo = CST_BLANK
            .strKHImageUrl = CST_BLANK
            .strWFMaxVal = CST_BLANK
            ReDim .strExtFrm(0)
            ReDim .strDispExtFrm(0)
            ReDim .strNormalVal(0)
            ReDim .strActNormalVal(0)
            ReDim .strInputDiv(0)
            ReDim .strSltVal(0)
            ReDim .strActSltVal(0)
            ReDim .strJsName(0)
        End With

        With Me.strcSelDataInfo
            .strSelBoreSize = CST_BLANK
            .strSelPtnNo = CST_BLANK
            .strSelPtn = CST_BLANK
            .strSelWFMaxVal = CST_BLANK
            .strSelOtherVal = CST_BLANK
            .hshtSelVal = New Hashtable
        End With

        With Me.strcErrInfo
            .strErrCd = CST_BLANK
            .strErrFocusNo = CST_BLANK
            .strErrPtnNo = CST_BLANK
            .strErrPtn = CST_BLANK
        End With

        Me.strcSelection.strUserID = strAUserID
        Me.strcSelection.strSessionID = strASessionID
        Me.strcSelection.strSeriesKataban = strASeriesKataban
        Me.strcSelection.strKeyKataban = strAKeyKataban

    End Sub

    ''' <summary>
    ''' ロッド先端特注情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strOpSymbol">形番引当画面の引当オプション</param>
    ''' <remarks>ロッド先端特注画面生成に必要な情報を取得する</remarks>
    Public Sub subRodInfoGet(objCon As SqlConnection, ByVal strOpSymbol As String())
        Dim strSelPtnNo As Integer
        Try
            '口径クリア
            Me.strcSelection.strBoreSize = CST_BLANK

            'ロッド先端特注標準情報クリア
            ReDim Me.strcRodPtnInfo(0)
            With Me.strcRodPtnInfo(0)
                .strRodPtn = CST_BLANK
                .strDispNo = CST_BLANK
                .strKHImageUrl = CST_BLANK
                .strWFMaxVal = CST_BLANK
                ReDim .strExtFrm(0)
                ReDim .strDispExtFrm(0)
                ReDim .strNormalVal(0)
                ReDim .strActNormalVal(0)
                ReDim .strInputDiv(0)
                ReDim .strSltVal(0)
                ReDim .strActSltVal(0)
                ReDim .strJsName(0)
            End With

            'ロッド先端特注選択情報クリア
            With Me.strcSelDataInfo
                .strSelBoreSize = CST_BLANK
                .strSelPtnNo = CST_BLANK
                .strSelPtn = CST_BLANK
                .strSelWFMaxVal = CST_BLANK
                .strSelOtherVal = CST_BLANK
                .hshtSelVal = New Hashtable
            End With

            '口径取得
            Call subBoreSizeSelect(objCon, strOpSymbol)

            'ロッド先端特注マスタ情報取得
            Call subRodEndMstSelect(objCon)

            'ロッド先端特注画面表記変更
            Call subKHImageChange(strOpSymbol, strSelPtnNo)

            'ロッド先端特注パターン詳細情報取得
            Call subRodDtlSet(objCon, strOpSymbol)

            '引当ロッド先端特注情報取得
            Call subSelRodSelect(objCon)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' ロッド先端情報更新
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <remarks></remarks>
    Public Sub subUpdateSelRod(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc)
        Dim bolReturn As Boolean
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            'ロッド先端フル形番クリア
            Me.strcSelection.strRodFullKtbn = CST_BLANK

            '引当ロッド先端特注クリア
            bolReturn = fncSPSelRodDel(objCon)

            '引当ロッド先端特注更新
            bolReturn = fncSPSelRodIns(objCon)

            'ロッド先端フル形番生成
            Call subRodFullKtbnCreate(objCon)

            '引当シリーズ形番更新(オプション情報)
            Call dalKtbnStrc.subSelSrsKtbnOptionUpd(objCon, Me.strcSelection.strUserID, _
                                                    Me.strcSelection.strSessionID, _
                                                    Me.strcSelection.strRodFullKtbn, _
                                                    objKtbnStrc.strcSelection.strOtherOption)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' ロッド先端フル形番生成
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks>入力したロッド先端データよりロッドフル形番を生成する</remarks>
    Private Sub subRodFullKtbnCreate(ByVal objCon As SqlConnection)
        Dim sbSql As New StringBuilder
        Dim dtResult As New DataTable
        Dim objCmd As SqlCommand
        Dim objAdp As SqlDataAdapter
        Dim intLoopCnt1 As Integer
        Try
            'SQL Query生成
            sbSql.Append(" SELECT      a.rod_pattern_symbol, ")
            sbSql.Append("             ISNULL(a.external_form , '') as external_form, ")
            sbSql.Append("             ISNULL(a.production_value , '') as production_value, ")
            sbSql.Append("             ISNULL(a.normal_value , '') as normal_value, ")
            sbSql.Append("             ISNULL(a.other_value , '') as other_value, ")
            sbSql.Append("             ISNULL(c.input_div , '') as input_div ")
            sbSql.Append(" FROM        kh_sel_rod_end_order a ")
            sbSql.Append(" INNER JOIN  kh_sel_srs_ktbn b ")
            sbSql.Append(" ON          a.user_id               = b.user_id ")
            sbSql.Append(" AND         a.session_id            = b.session_id ")
            sbSql.Append(" LEFT JOIN   kh_rod_end_ext_frm c ")
            sbSql.Append(" ON          b.series_kataban        = c.series_kataban ")
            sbSql.Append(" AND         b.key_kataban           = c.key_kataban ")
            sbSql.Append(" AND         a.rod_pattern_symbol    = c.rod_pattern_symbol ")
            sbSql.Append(" AND         a.external_form         = c.external_form ")
            sbSql.Append(" WHERE       a.user_id               = @UserID ")
            sbSql.Append(" AND         a.session_id            = @SessionID ")
            sbSql.Append(" ORDER BY    a.external_form_seq_no")
            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@UserID", SqlDbType.VarChar, 10).Value = Me.strcSelection.strUserID
                .Parameters.Add("@SessionID", SqlDbType.VarChar, 88).Value = Me.strcSelection.strSessionID
            End With
            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dtResult)
            If dtResult.Rows.Count <> 0 Then
                'フル形番生成
                If dtResult.Rows(intLoopCnt1).Item("other_value").Trim.Length <> 0 Then
                    Me.strcSelection.strRodFullKtbn = dtResult.Rows(intLoopCnt1).Item("other_value")
                Else
                    Me.strcSelection.strRodFullKtbn = dtResult.Rows(intLoopCnt1).Item("rod_pattern_symbol")
                    For intLoopCnt1 = 0 To dtResult.Rows.Count - 1
                        If dtResult.Rows(intLoopCnt1).Item("production_value").Trim.Length <> 0 And _
                           dtResult.Rows(intLoopCnt1).Item("input_div").Trim <> CdCst.RodEndCstmOrder.Label Then
                            If dtResult.Rows(intLoopCnt1).Item("production_value").Trim <> dtResult.Rows(intLoopCnt1).Item("normal_value").Trim Then
                                Me.strcSelection.strRodFullKtbn = Me.strcSelection.strRodFullKtbn & dtResult.Rows(intLoopCnt1).Item("external_form").Trim & _
                                                               Replace(Replace(dtResult.Rows(intLoopCnt1).Item("production_value").Trim, "(", CST_BLANK), ")", CST_BLANK)
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            sbSql = Nothing
            objAdp = Nothing
            dtResult = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 引当ロッド先端特注テーブル追加処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSPSelRodIns(ByVal objCon As SqlConnection) As Boolean
        Dim objCmd As SqlCommand = Nothing
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer

        fncSPSelRodIns = False
        Try
            objCmd = objCon.CreateCommand
            With objCmd
                .CommandType = CommandType.StoredProcedure
                .CommandText = CdCst.DB.SPL.KHSelRodIns

                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10)
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88)
                .Parameters.Add("@BoreSize", SqlDbType.Int)
                .Parameters.Add("@RodPatternSymbol", SqlDbType.VarChar, 10)
                .Parameters.Add("@ExternalFormSeqNo", SqlDbType.Int)
                .Parameters.Add("@ExternalForm", SqlDbType.VarChar, 10)
                .Parameters.Add("@ProductionValue", SqlDbType.VarChar, 15)
                .Parameters.Add("@NormalValue", SqlDbType.VarChar, 15)
                .Parameters.Add("@OtherValue", SqlDbType.VarChar, 100)
                .Parameters.Add("@RegPerson", SqlDbType.VarChar, 10)
                .Parameters.Add("@RegDate", SqlDbType.DateTime, 88)
                .Parameters.Add("@CurPerson", SqlDbType.VarChar, 10)
                .Parameters.Add("@CurDate", SqlDbType.DateTime, 88)

                .Parameters("@UserId").Value = Me.strcSelection.strUserID
                .Parameters("@SessionId").Value = Me.strcSelection.strSessionID
                .Parameters("@BoreSize").Value = CInt(Me.strcSelection.strBoreSize.Trim)
                .Parameters("@RodPatternSymbol").Value = Me.strcSelDataInfo.strSelPtn
                .Parameters("@RegPerson").Value = Me.strcSelection.strUserID
                .Parameters("@RegDate").Value = Now()
                .Parameters("@CurPerson").Value = DBNull.Value
                .Parameters("@CurDate").Value = DBNull.Value

                If Me.strcSelDataInfo.strSelOtherVal IsNot Nothing Then
                    'その他ロッド先端特注パターンの場合
                    Dim isPtn As Boolean = False

                    'パターン指定をチェック
                    If InStr(Me.strcSelDataInfo.strSelOtherVal, "WF") > 0 Then
                        'WF指定
                        Dim strPtn As String = ""
                        strPtn = Mid(Me.strcSelDataInfo.strSelOtherVal, 1, InStr(Me.strcSelDataInfo.strSelOtherVal, "WF") - 1)

                        For intLoopCnt1 = 1 To UBound(Me.strcRodPtnInfo)
                            If strPtn = Me.strcRodPtnInfo(intLoopCnt1).strRodPtn Then
                                isPtn = True
                                For intLoopCnt2 = 1 To UBound(Me.strcRodPtnInfo(intLoopCnt1).strExtFrm)
                                    .Parameters("@ExternalFormSeqNo").Value = intLoopCnt2
                                    .Parameters("@ExternalForm").Value = Me.strcRodPtnInfo(intLoopCnt1).strExtFrm(intLoopCnt2)
                                    .Parameters("@ProductionValue").Value = CST_BLANK
                                    .Parameters("@NormalValue").Value = Me.strcRodPtnInfo(intLoopCnt1).strNormalVal(intLoopCnt2)
                                    .Parameters("@OtherValue").Value = Me.strcSelDataInfo.strSelOtherVal

                                    '実行
                                    objCmd.ExecuteNonQuery()
                                Next

                                Exit For
                            End If
                        Next
                    End If

                    '一致するパターンがなかった場合、空値を設定
                    If Not isPtn Then
                        .Parameters("@ExternalFormSeqNo").Value = 0
                        .Parameters("@ExternalForm").Value = CST_BLANK
                        .Parameters("@ProductionValue").Value = CST_BLANK
                        .Parameters("@NormalValue").Value = CST_BLANK
                        .Parameters("@OtherValue").Value = Me.strcSelDataInfo.strSelOtherVal
                        '実行
                        objCmd.ExecuteNonQuery()
                    End If
                ElseIf Me.strcSelDataInfo.hshtSelVal.Count = 0 Then
                    'ラジオボタンのみのロッド先端パターンの場合(N11-N13/N13-N11)
                    .Parameters("@ExternalFormSeqNo").Value = 0
                    .Parameters("@ExternalForm").Value = CST_BLANK
                    .Parameters("@ProductionValue").Value = CST_BLANK
                    .Parameters("@NormalValue").Value = CST_BLANK
                    .Parameters("@OtherValue").Value = CST_BLANK

                    '実行
                    objCmd.ExecuteNonQuery()
                Else
                    '寸法表ありの場合
                    For intLoopCnt1 = 1 To UBound(Me.strcRodPtnInfo(Me.strcSelDataInfo.strSelPtnNo).strExtFrm)
                        .Parameters("@ExternalFormSeqNo").Value = intLoopCnt1
                        .Parameters("@ExternalForm").Value = Me.strcRodPtnInfo(Me.strcSelDataInfo.strSelPtnNo).strExtFrm(intLoopCnt1)
                        .Parameters("@ProductionValue").Value = IIf(Me.strcSelDataInfo.hshtSelVal.ContainsKey(Me.strcRodPtnInfo(Me.strcSelDataInfo.strSelPtnNo).strExtFrm(intLoopCnt1)), Me.strcSelDataInfo.hshtSelVal(Me.strcRodPtnInfo(Me.strcSelDataInfo.strSelPtnNo).strExtFrm(intLoopCnt1)), CST_BLANK)
                        .Parameters("@NormalValue").Value = Me.strcRodPtnInfo(Me.strcSelDataInfo.strSelPtnNo).strNormalVal(intLoopCnt1)
                        .Parameters("@OtherValue").Value = CST_BLANK

                        '実行
                        objCmd.ExecuteNonQuery()
                    Next
                End If
            End With
            fncSPSelRodIns = True
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
    ''' 引当ロッド先端特注取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks>引当ロッド先端特注を読み込み引当ロッド先端特注情報を取得し、メンバ変数にセットする</remarks>
    Private Sub subSelRodSelect(objCon As SqlConnection)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing
        Try
            'SQL Query生成
            sbSql.Append(" SELECT      bore_size, ")
            sbSql.Append("             rod_pattern_symbol, ")
            sbSql.Append("             external_form, ")
            sbSql.Append("             production_value, ")
            sbSql.Append("             other_value ")
            sbSql.Append(" FROM        kh_sel_rod_end_order ")
            sbSql.Append(" WHERE       user_id    = @UserId ")
            sbSql.Append(" AND         session_id = @SessionId ")
            sbSql.Append(" ORDER BY    external_form_seq_no")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text

                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = Me.strcSelection.strUserID
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = Me.strcSelection.strSessionID
            End With
            objRdr = objCmd.ExecuteReader
            While objRdr.Read()
                With Me.strcSelDataInfo
                    .strSelBoreSize = objRdr.GetValue(objRdr.GetOrdinal("bore_size"))
                    .strSelPtn = objRdr.GetValue(objRdr.GetOrdinal("rod_pattern_symbol"))
                    .strSelOtherVal = objRdr.GetValue(objRdr.GetOrdinal("other_value"))
                    .hshtSelVal.Add(objRdr.GetValue(objRdr.GetOrdinal("external_form")), objRdr.GetValue(objRdr.GetOrdinal("production_value")))
                End With
            End While
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 引当口径検索取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strOpSymbol">形番引当画面の引当オプション</param>
    ''' <remarks></remarks>
    Private Sub subBoreSizeSelect(objCon As SqlConnection, ByVal strOpSymbol() As String)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing
        Try
            'SQL Query生成
            sbSql.Append(" SELECT  ktbn_strc_seq_no ")
            sbSql.Append(" FROM    kh_kataban_strc ")
            sbSql.Append(" WHERE   series_kataban  = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban     = @KeyKataban ")
            sbSql.Append(" AND     element_div = '" & CdCst.RodEndCstmOrder.EleBoreSize & "'")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = Me.strcSelection.strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = Me.strcSelection.strKeyKataban
            End With
            objRdr = objCmd.ExecuteReader
            While objRdr.Read()
                Me.strcSelection.strBoreSize = IIf(IsDBNull(objRdr.GetValue(objRdr.GetOrdinal("ktbn_strc_seq_no"))), CST_BLANK, strOpSymbol(objRdr.GetValue(objRdr.GetOrdinal("ktbn_strc_seq_no"))))
            End While
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' ロッド先端特注マスタ検索取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks></remarks>
    Private Sub subRodEndMstSelect(objCon As SqlConnection)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  pattern_seq_no, ")
            sbSql.Append("         rod_pattern_symbol, ")
            sbSql.Append("         ISNULL(url , '') AS url ")
            sbSql.Append(" FROM    kh_rod_end_mst ")
            sbSql.Append(" WHERE   series_kataban         = @SeriesKataban ")
            sbSql.Append(" AND     key_kataban            = @KeyKataban ")
            sbSql.Append(" ORDER BY  pattern_seq_no ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = Me.strcSelection.strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = Me.strcSelection.strKeyKataban
            End With

            objRdr = objCmd.ExecuteReader
            While objRdr.Read()
                '配列再定義
                ReDim Preserve Me.strcRodPtnInfo(UBound(Me.strcRodPtnInfo) + 1)
                'ロッド先端特注パターン記号
                Me.strcRodPtnInfo(UBound(Me.strcRodPtnInfo)).strRodPtn = objRdr.GetValue(objRdr.GetOrdinal("rod_pattern_symbol"))
                'ロッド先端特注表示順序
                Me.strcRodPtnInfo(UBound(Me.strcRodPtnInfo)).strDispNo = objRdr.GetValue(objRdr.GetOrdinal("pattern_seq_no"))
                'イメージURL
                Me.strcRodPtnInfo(UBound(Me.strcRodPtnInfo)).strKHImageUrl = objRdr.GetValue(objRdr.GetOrdinal("url"))
            End While
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' ロッド先端特注画面変更
    ''' </summary>
    ''' <param name="strOpSymbol">選択オプション</param>
    ''' <param name="strSelPtnNo"></param>
    ''' <remarks>選択オプション要素によって、ロッド先端特注画面を変更する</remarks>
    Private Sub subKHImageChange(ByVal strOpSymbol As String(), ByVal strSelPtnNo As String)
        Dim objRdr As SqlDataReader = Nothing
        Try
            Select Case Me.strcSelection.strSeriesKataban
                Case "JSC3"
                    If Me.strcSelection.strKeyKataban = "1" Then
                        If InStr(strOpSymbol(4), "FA") <> 0 Then
                            'イメージURL
                            ReDim Me.strcRodPtnInfo(11)
                            With Me.strcRodPtnInfo(11)
                                'ロッド先端特注画面イメージ
                                Me.strcRodPtnInfo(0).strKHImageUrl = CST_BLANK
                                Me.strcRodPtnInfo(1).strKHImageUrl = "../KHImage/JSC3RodN13(FA).gif"
                                Me.strcRodPtnInfo(2).strKHImageUrl = "../KHImage/JSC3RodN15(FA).gif"
                                Me.strcRodPtnInfo(3).strKHImageUrl = "../KHImage/JSC3RodN11(FA).gif"
                                Me.strcRodPtnInfo(4).strKHImageUrl = "../KHImage/JSC3RodN1(FA).gif"
                                Me.strcRodPtnInfo(5).strKHImageUrl = "../KHImage/JSC3RodN12(FA).gif"
                                Me.strcRodPtnInfo(6).strKHImageUrl = "../KHImage/JSC3RodN14(FA).gif"
                                Me.strcRodPtnInfo(7).strKHImageUrl = "../KHImage/JSC3RodN3(FA).gif"
                                Me.strcRodPtnInfo(8).strKHImageUrl = "../KHImage/JSC3RodN31(FA).gif"
                                Me.strcRodPtnInfo(9).strKHImageUrl = "../KHImage/JSC3RodN2(FA).gif"
                                Me.strcRodPtnInfo(10).strKHImageUrl = "../KHImage/JSC3RodN21(FA).gif"
                                Me.strcRodPtnInfo(11).strKHImageUrl = CST_BLANK
                                'ロッド先端特注パターン記号
                                Me.strcRodPtnInfo(0).strRodPtn = CST_BLANK
                                Me.strcRodPtnInfo(1).strRodPtn = "N13"
                                Me.strcRodPtnInfo(2).strRodPtn = "N15"
                                Me.strcRodPtnInfo(3).strRodPtn = "N11"
                                Me.strcRodPtnInfo(4).strRodPtn = "N1"
                                Me.strcRodPtnInfo(5).strRodPtn = "N12"
                                Me.strcRodPtnInfo(6).strRodPtn = "N14"
                                Me.strcRodPtnInfo(7).strRodPtn = "N3"
                                Me.strcRodPtnInfo(8).strRodPtn = "N31"
                                Me.strcRodPtnInfo(9).strRodPtn = "N2"
                                Me.strcRodPtnInfo(10).strRodPtn = "N21"
                                Me.strcRodPtnInfo(11).strRodPtn = "Other"
                            End With
                        End If
                    End If
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' ロッド先端特注詳細情報セット
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strOpSymbol">選択オプション</param>
    ''' <remarks></remarks>
    Private Sub subRodDtlSet(ByVal objCon As SqlConnection, ByVal strOpSymbol As String())

        Dim dtResult As New DataTable
        Dim sbDropLst As New StringBuilder
        Dim sbDropLstVal As New StringBuilder
        Dim strEtlFrm As String
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer

        Try
            'パターン記号別の外径種類情報を取得する
            dtResult = fncRodPtnDtlSelect(objCon)

            For intLoopCnt1 = 1 To UBound(Me.strcRodPtnInfo)

                '配列定義
                ReDim Me.strcRodPtnInfo(intLoopCnt1).strExtFrm(0)
                ReDim Me.strcRodPtnInfo(intLoopCnt1).strDispExtFrm(0)
                ReDim Me.strcRodPtnInfo(intLoopCnt1).strNormalVal(0)
                ReDim Me.strcRodPtnInfo(intLoopCnt1).strActNormalVal(0)
                ReDim Me.strcRodPtnInfo(intLoopCnt1).strInputDiv(0)
                ReDim Me.strcRodPtnInfo(intLoopCnt1).strSltVal(0)
                ReDim Me.strcRodPtnInfo(intLoopCnt1).strActSltVal(0)
                ReDim Me.strcRodPtnInfo(intLoopCnt1).strJsName(0)

                'WF最大寸法
                Me.strcRodPtnInfo(intLoopCnt1).strWFMaxVal = dtResult.Rows(0).Item("wf_max_value")

                '初期値設定
                strEtlFrm = CST_BLANK

                For intLoopCnt2 = 0 To dtResult.Rows.Count - 1
                    If intLoopCnt1 = dtResult.Rows(intLoopCnt2).Item("pattern_seq_no") Then
                        If strEtlFrm <> dtResult.Rows(intLoopCnt2).Item("external_form") Then
                            '配列再定義
                            ReDim Preserve Me.strcRodPtnInfo(intLoopCnt1).strExtFrm(UBound(Me.strcRodPtnInfo(intLoopCnt1).strExtFrm) + 1)
                            ReDim Preserve Me.strcRodPtnInfo(intLoopCnt1).strDispExtFrm(UBound(Me.strcRodPtnInfo(intLoopCnt1).strDispExtFrm) + 1)
                            ReDim Preserve Me.strcRodPtnInfo(intLoopCnt1).strNormalVal(UBound(Me.strcRodPtnInfo(intLoopCnt1).strNormalVal) + 1)
                            ReDim Preserve Me.strcRodPtnInfo(intLoopCnt1).strActNormalVal(UBound(Me.strcRodPtnInfo(intLoopCnt1).strActNormalVal) + 1)
                            ReDim Preserve Me.strcRodPtnInfo(intLoopCnt1).strInputDiv(UBound(Me.strcRodPtnInfo(intLoopCnt1).strInputDiv) + 1)
                            ReDim Preserve Me.strcRodPtnInfo(intLoopCnt1).strSltVal(UBound(Me.strcRodPtnInfo(intLoopCnt1).strSltVal) + 1)
                            ReDim Preserve Me.strcRodPtnInfo(intLoopCnt1).strActSltVal(UBound(Me.strcRodPtnInfo(intLoopCnt1).strActSltVal) + 1)
                            ReDim Preserve Me.strcRodPtnInfo(intLoopCnt1).strJsName(UBound(Me.strcRodPtnInfo(intLoopCnt1).strJsName) + 1)
                            '外径寸法
                            Me.strcRodPtnInfo(intLoopCnt1).strExtFrm(UBound(Me.strcRodPtnInfo(intLoopCnt1).strExtFrm)) = dtResult.Rows(intLoopCnt2).Item("external_form")
                            '表示外径寸法
                            Me.strcRodPtnInfo(intLoopCnt1).strDispExtFrm(UBound(Me.strcRodPtnInfo(intLoopCnt1).strDispExtFrm)) = dtResult.Rows(intLoopCnt2).Item("disp_external_form")
                            '標準寸法
                            Me.strcRodPtnInfo(intLoopCnt1).strNormalVal(UBound(Me.strcRodPtnInfo(intLoopCnt1).strNormalVal)) = dtResult.Rows(intLoopCnt2).Item("normal_value")
                            '実標準寸法
                            Me.strcRodPtnInfo(intLoopCnt1).strActNormalVal(UBound(Me.strcRodPtnInfo(intLoopCnt1).strActNormalVal)) = dtResult.Rows(intLoopCnt2).Item("act_normal_value")
                            '入力区分
                            Me.strcRodPtnInfo(intLoopCnt1).strInputDiv(UBound(Me.strcRodPtnInfo(intLoopCnt1).strInputDiv)) = dtResult.Rows(intLoopCnt2).Item("input_div")
                            'javascript名
                            Me.strcRodPtnInfo(intLoopCnt1).strJsName(UBound(Me.strcRodPtnInfo(intLoopCnt1).strJsName)) = dtResult.Rows(intLoopCnt2).Item("js_name")
                            If IsDBNull(dtResult.Rows(intLoopCnt2).Item("selectable_value")) Then
                                sbDropLst = Nothing
                                sbDropLstVal = Nothing
                            Else
                                If sbDropLst Is Nothing Then
                                    sbDropLst = New StringBuilder
                                    sbDropLstVal = New StringBuilder
                                    '選択可能寸法
                                    Me.strcRodPtnInfo(intLoopCnt1).strSltVal(UBound(Me.strcRodPtnInfo(intLoopCnt1).strSltVal)) = dtResult.Rows(intLoopCnt2).Item("selectable_value")
                                    '実選択可能寸法
                                    Me.strcRodPtnInfo(intLoopCnt1).strActSltVal(UBound(Me.strcRodPtnInfo(intLoopCnt1).strActSltVal)) = dtResult.Rows(intLoopCnt2).Item("act_selectable_value")
                                Else
                                    '選択可能寸法
                                    Me.strcRodPtnInfo(intLoopCnt1).strSltVal(UBound(Me.strcRodPtnInfo(intLoopCnt1).strSltVal)) = sbDropLst.ToString
                                    '実選択可能寸法
                                    Me.strcRodPtnInfo(intLoopCnt1).strActSltVal(UBound(Me.strcRodPtnInfo(intLoopCnt1).strActSltVal)) = sbDropLstVal.ToString
                                End If
                            End If
                        Else
                            sbDropLst.Append(CST_COMMA & dtResult.Rows(intLoopCnt2).Item("selectable_value"))
                            sbDropLstVal.Append(CST_COMMA & dtResult.Rows(intLoopCnt2).Item("act_selectable_value"))
                            '選択可能寸法
                            Me.strcRodPtnInfo(intLoopCnt1).strSltVal(UBound(Me.strcRodPtnInfo(intLoopCnt1).strSltVal)) = sbDropLst.ToString
                            '実選択可能寸法
                            Me.strcRodPtnInfo(intLoopCnt1).strActSltVal(UBound(Me.strcRodPtnInfo(intLoopCnt1).strActSltVal)) = sbDropLstVal.ToString
                        End If
                        strEtlFrm = dtResult.Rows(intLoopCnt2).Item("external_form")
                    End If
                Next
            Next

            'その他ロッド先端特注標準寸法を取得する
            dtResult = subOtherWFSelect(objCon)
            ReDim Me.strcRodPtnInfo(UBound(Me.strcRodPtnInfo)).strExtFrm(UBound(Me.strcRodPtnInfo(UBound(Me.strcRodPtnInfo)).strExtFrm) + 1)
            ReDim Me.strcRodPtnInfo(UBound(Me.strcRodPtnInfo)).strNormalVal(UBound(Me.strcRodPtnInfo(UBound(Me.strcRodPtnInfo)).strNormalVal) + 1)
            Me.strcRodPtnInfo(UBound(Me.strcRodPtnInfo)).strExtFrm(UBound(Me.strcRodPtnInfo(UBound(Me.strcRodPtnInfo)).strExtFrm)) = dtResult.Rows(0).Item("external_form")
            Me.strcRodPtnInfo(UBound(Me.strcRodPtnInfo)).strNormalVal(UBound(Me.strcRodPtnInfo(UBound(Me.strcRodPtnInfo)).strNormalVal)) = dtResult.Rows(0).Item("normal_value")

            'WF標準寸法変更処理
            Call subWFChange(strOpSymbol)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            dtResult = Nothing
            sbDropLst = Nothing
            sbDropLstVal = Nothing
        End Try

    End Sub

    ''' <summary>
    ''' 標準寸法変更
    ''' </summary>
    ''' <param name="strOpSymbol">選択オプション</param>
    ''' <remarks>選択オプション要素によって、WF寸法を変更する</remarks>
    Private Sub subWFChange(ByVal strOpSymbol As String())
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer
        Try
            Select Case Me.strcSelection.strSeriesKataban
                Case "SCA2"
                    If InStr(strOpSymbol(1), "Q2") <> 0 And _
                       (strOpSymbol(8) = "R" Or _
                        strOpSymbol(8) = "HR") Then
                    Else
                        For intLoopCnt1 = 1 To UBound(Me.strcRodPtnInfo)
                            For intLoopCnt2 = 1 To UBound(Me.strcRodPtnInfo(intLoopCnt1).strExtFrm)
                                If Me.strcRodPtnInfo(intLoopCnt1).strExtFrm(intLoopCnt2) = CdCst.RodEndCstmOrder.FrmWF Then
                                    Select Case Me.strcSelection.strBoreSize
                                        Case "40"
                                            Me.strcRodPtnInfo(intLoopCnt1).strNormalVal(intLoopCnt2) = "33.5"
                                        Case "50"
                                            Me.strcRodPtnInfo(intLoopCnt1).strNormalVal(intLoopCnt2) = "37"
                                        Case "63"
                                            Me.strcRodPtnInfo(intLoopCnt1).strNormalVal(intLoopCnt2) = "35"
                                        Case "80"
                                            Me.strcRodPtnInfo(intLoopCnt1).strNormalVal(intLoopCnt2) = "48"
                                        Case "100"
                                            Me.strcRodPtnInfo(intLoopCnt1).strNormalVal(intLoopCnt2) = "53"
                                    End Select
                                End If
                            Next
                        Next
                    End If
                Case "JSC3"
                    '表示外形寸法変更対応(WF→FF)
                    If Me.strcSelection.strKeyKataban = "1" Then
                        If InStr(strOpSymbol(4), "FA") <> 0 Then
                            For intLoopCnt1 = 1 To 10
                                For intLoopCnt2 = 1 To UBound(Me.strcRodPtnInfo(intLoopCnt1).strExtFrm)
                                    If Me.strcRodPtnInfo(intLoopCnt1).strExtFrm(intLoopCnt2) = CdCst.RodEndCstmOrder.FrmWF Then
                                        Me.strcRodPtnInfo(intLoopCnt1).strDispExtFrm(intLoopCnt2) = "FF"
                                    End If
                                Next
                            Next
                        End If
                    End If
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' ロッド先端特注外径種類/ロッド先端特注標準寸法検索取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks>ロッド先端特注外径種類/ロッド先端特注標準寸法テーブルからロッド先端パターンの詳細情報を取得する</remarks>
    Private Function fncRodPtnDtlSelect(ByVal objCon As SqlConnection) As DataTable
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objAdp As SqlDataAdapter
        fncRodPtnDtlSelect = New DataTable

        Try
            'SQL Query生成
            sbSql.Append(" SELECT  b.pattern_seq_no, ")
            sbSql.Append("         a.rod_pattern_symbol, ")
            sbSql.Append("         a.external_form, ")
            sbSql.Append("         a.disp_external_form, ")
            sbSql.Append("         a.input_div, ")
            sbSql.Append("         a.js_name, ")
            sbSql.Append("         c.normal_value, ")
            sbSql.Append("         c.act_normal_value , ")
            sbSql.Append("         d.selectable_value,  ")
            sbSql.Append("         d.act_selectable_value, ")
            sbSql.Append("         e.wf_max_value ")
            sbSql.Append(" FROM    kh_rod_end_ext_frm a ")
            sbSql.Append(" INNER JOIN  kh_rod_end_mst b ")
            sbSql.Append(" ON      a.series_kataban         = b.series_kataban ")
            sbSql.Append(" AND     a.key_kataban            = b.key_kataban ")
            sbSql.Append(" AND     a.rod_pattern_symbol     = b.rod_pattern_symbol ")
            sbSql.Append(" INNER JOIN  kh_rod_end_std_size c ")
            sbSql.Append(" ON      a.series_kataban         = c.series_kataban ")
            sbSql.Append(" AND     a.key_kataban            = c.key_kataban ")
            sbSql.Append(" AND     a.rod_pattern_symbol     = c.rod_pattern_symbol ")
            sbSql.Append(" AND     a.external_form          = c.external_form ")
            sbSql.Append(" AND     c.bore_size              = @BoreSize ")
            sbSql.Append(" LEFT  JOIN  kh_rod_end_selectable_size d")
            sbSql.Append(" ON      a.series_kataban         = d.series_kataban ")
            sbSql.Append(" AND     a.key_kataban            = d.key_kataban ")
            sbSql.Append(" AND     a.rod_pattern_symbol     = d.rod_pattern_symbol ")
            sbSql.Append(" AND     a.external_form          = d.external_form ")
            sbSql.Append(" AND     d.bore_size              = @BoreSize ")
            sbSql.Append(" LEFT  JOIN  kh_rod_end_wf_max_size e")
            sbSql.Append(" ON      a.series_kataban         = e.series_kataban ")
            sbSql.Append(" AND     a.key_kataban            = e.key_kataban ")
            sbSql.Append(" AND     e.bore_size              = @BoreSize ")
            sbSql.Append(" WHERE   a.series_kataban         = @SeriesKataban ")
            sbSql.Append(" AND     a.key_kataban            = @KeyKataban ")
            sbSql.Append(" ORDER BY  b.pattern_seq_no, a.external_form_seq_no, d.sel_value_seq_no")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = Me.strcSelection.strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = Me.strcSelection.strKeyKataban
                .Parameters.Add("@BoreSize", SqlDbType.Int).Value = CInt(Me.strcSelection.strBoreSize.Trim)
            End With

            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(fncRodPtnDtlSelect)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            sbSql = Nothing
            objAdp = Nothing
        End Try
    End Function

    ''' <summary>
    ''' その他ロッド先端標準寸法検索処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <returns></returns>
    ''' <remarks>ロッド先端特注標準寸法を検索し、その他ロッド先端標準寸法を取得しセットする</remarks>
    Private Function subOtherWFSelect(ByVal objCon As SqlConnection) As DataTable
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objAdp As SqlDataAdapter
        subOtherWFSelect = New DataTable
        Try

            'SQL Query生成
            sbSql.Append(" SELECT      normal_value, ")
            sbSql.Append("             external_form ")
            sbSql.Append(" FROM        kh_rod_end_std_size ")
            sbSql.Append(" WHERE       series_kataban         = @SeriesKataban ")
            sbSql.Append(" AND         key_kataban            = @KeyKataban ")
            sbSql.Append(" AND         bore_size              = @BoreSize ")
            sbSql.Append(" AND         rod_pattern_symbol     = @RodPtnSymbol ")
            sbSql.Append(" AND         external_form          = @ExtForm ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SeriesKataban", SqlDbType.VarChar, 10).Value = Me.strcSelection.strSeriesKataban
                .Parameters.Add("@KeyKataban", SqlDbType.VarChar, 2).Value = Me.strcSelection.strKeyKataban
                .Parameters.Add("@BoreSize", SqlDbType.Int).Value = CInt(Me.strcSelection.strBoreSize.Trim)
                .Parameters.Add("@RodPtnSymbol", SqlDbType.VarChar, 10).Value = CdCst.RodEndCstmOrder.OtherSize
                .Parameters.Add("@ExtForm", SqlDbType.VarChar, 10).Value = CdCst.RodEndCstmOrder.FrmWF
            End With

            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(subOtherWFSelect)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            sbSql = Nothing
            objAdp = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 引当ロッド先端特注テーブル削除処理
    ''' </summary>
    ''' <param name="objCon">DB接続オブジェクト</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSPSelRodDel(ByVal objCon As SqlConnection) As Boolean
        Dim objCmd As SqlCommand = Nothing
        fncSPSelRodDel = False
        Try
            objCmd = objCon.CreateCommand

            With objCmd
                .CommandType = CommandType.StoredProcedure
                .CommandText = CdCst.DB.SPL.KHSelRodDel
                ' 定義
                .Parameters.Add("@UserId", SqlDbType.VarChar, 10).Value = Me.strcSelection.strUserID
                .Parameters.Add("@SessionId", SqlDbType.NVarChar, 88).Value = Me.strcSelection.strSessionID
            End With

            '実行
            objCmd.ExecuteNonQuery()
            fncSPSelRodDel = True
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
    ''' 入力チェック
    ''' </summary>
    ''' <param name="strSelPtnNo">選択ロッド先端パターンNo.</param>
    ''' <param name="strSelProdSize">選択特注寸法</param>
    ''' <param name="strOpSymbol">選択オプション</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncInpCheck(ByVal strSelPtnNo As String, _
                                ByVal strSelProdSize As String(), _
                                ByVal strOpSymbol As String()) As Boolean
        fncInpCheck = False
        Try
            'ロッド先端特注選択情報クリア
            With Me.strcSelDataInfo
                .strSelBoreSize = CST_BLANK
                .strSelPtnNo = CST_BLANK
                .strSelPtn = CST_BLANK
                .strSelWFMaxVal = CST_BLANK
                .strSelOtherVal = CST_BLANK
                .hshtSelVal = New Hashtable
            End With

            'エラー情報クリア
            With Me.strcErrInfo
                .strErrCd = CST_BLANK
                .strErrFocusNo = CST_BLANK
                .strErrPtnNo = CST_BLANK
                .strErrPtn = CST_BLANK
            End With

            '選択情報をセット
            Call subSelDataSet(strSelPtnNo, strSelProdSize)

            '機種毎にチェック
            Select Case Me.strcSelection.strSeriesKataban
                Case "SSD"
                    If fncSSDInpCheck(strOpSymbol) Then
                        fncInpCheck = True
                    End If
                Case "SCA2"
                    If fncSCA2InpCheck(strOpSymbol, Me.strcSelection.strKeyKataban) Then
                        fncInpCheck = True
                    End If
                Case "JSC3", "JSC4"
                    If fncJSC3InpCheck(strOpSymbol) Then
                        fncInpCheck = True
                    End If
                Case "SCS", "SCS2"
                    If fncSCSInpCheck(strOpSymbol) Then
                        fncInpCheck = True
                    End If
                Case "CMK2"
                    If fncCMK2InpCheck(strOpSymbol) Then
                        fncInpCheck = True
                    End If
            End Select

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 選択データをセットする
    ''' </summary>
    ''' <param name="strSelPtnNo">ロッド先端パターンNo.</param>
    ''' <param name="strSelProdSize">選択特注寸法</param>
    ''' <remarks></remarks>
    Private Sub subSelDataSet(ByVal strSelPtnNo As String, ByVal strSelProdSize As String())
        Dim intLoopCnt As Integer
        Try
            '選択データセット
            With Me.strcSelDataInfo
                .strSelPtnNo = strSelPtnNo
                .strSelPtn = Me.strcRodPtnInfo(strSelPtnNo).strRodPtn
                .strSelWFMaxVal = Me.strcRodPtnInfo(strSelPtnNo).strWFMaxVal
                If .strSelPtn = CdCst.RodEndCstmOrder.OtherSize Then
                    .strSelOtherVal = strSelProdSize(1)
                Else
                    .strSelOtherVal = Nothing
                    For intLoopCnt = 1 To UBound(strSelProdSize)
                        If strSelProdSize(intLoopCnt).Trim.Length <> 0 Then
                            .hshtSelVal.Add(Me.strcRodPtnInfo(strSelPtnNo).strExtFrm(intLoopCnt), strSelProdSize(intLoopCnt))
                        End If
                    Next
                End If
            End With
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <param name="strOpSymbol">選択オプション  </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSSDInpCheck(ByVal strOpSymbol As String()) As Boolean

        Dim strActPtn As String
        Dim strActPtnNo As String
        Dim hshtStdSize As New Hashtable
        Dim hshtSelSize As New Hashtable
        Dim hshtFrmPos As New Hashtable
        Dim hshtInputDiv As New Hashtable
        Dim strErrCd As String
        Dim intFrmPos As Integer
        fncSSDInpCheck = True
        Try
            '初期設定
            strActPtn = CST_BLANK
            strActPtnNo = CST_BLANK
            strErrCd = CST_BLANK
            intFrmPos = 0

            'ロッド先端パターン設定
            Select Case Me.strcSelDataInfo.strSelPtn
                Case CdCst.RodEndCstmOrder.OtherSize
                    If fncRodPtnGet(Me.strcSelDataInfo.strSelOtherVal, strActPtn, strActPtnNo, strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSSDInpCheck = False
                        Exit Function
                    End If
                Case Else
                    strActPtn = Me.strcSelDataInfo.strSelPtn
                    strActPtnNo = Me.strcSelDataInfo.strSelPtnNo
            End Select

            '選択データセット
            Call subFrmDataSet(strActPtnNo, hshtStdSize, hshtInputDiv, hshtSelSize, hshtFrmPos, Me.strcSelDataInfo.strSelOtherVal)

            Select Case Me.strcSelDataInfo.strSelPtn
                Case CdCst.RodEndCstmOrder.OtherSize
                    'ハイフンチェック
                    If fncOthHypenChk(Me.strcSelDataInfo.strSelOtherVal, strErrCd, _
                                      CdCst.RodEndCstmOrder.RodPtnN13N11 & CST_COMMA & CdCst.RodEndCstmOrder.RodPtnN11N13) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSSDInpCheck = False
                        Exit Function
                    End If
                    'WFの後に数値がなかったらエラー
                    If fncNumericCheck(Me.strcSelDataInfo.strSelOtherVal, CdCst.RodEndCstmOrder.FrmWF, , strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = CdCst.RodEndCstmOrder.FrmWF
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSSDInpCheck = False
                        Exit Function
                    End If
                    'Aの後に数値がなかったらエラー
                    If fncNumericCheck(Me.strcSelDataInfo.strSelOtherVal, CdCst.RodEndCstmOrder.FrmA, , strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = CdCst.RodEndCstmOrder.FrmA
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSSDInpCheck = False
                        Exit Function
                    End If
                Case Else
                    'A/KL寸法チェック
                    If fncStdAKLChk(Me.strcSelDataInfo.strSelPtn, hshtSelSize, hshtStdSize, hshtFrmPos, strErrCd, intFrmPos) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        Me.strcErrInfo.strErrFocusNo = intFrmPos
                        fncSSDInpCheck = False
                        Exit Function
                    End If
                    'WF寸法チェック
                    If fncStdWFChk(hshtSelSize, hshtStdSize, hshtFrmPos, strErrCd, intFrmPos) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        Me.strcErrInfo.strErrFocusNo = intFrmPos
                        fncSSDInpCheck = False
                        Exit Function
                    End If
            End Select

            Select Case strActPtn
                Case CdCst.RodEndCstmOrder.RodPtnN13N11, CdCst.RodEndCstmOrder.RodPtnN11N13
                Case Else
                    'N13/N11チェック
                    If fncSelectChk(strActPtn, _
                                    CdCst.RodEndCstmOrder.RodPtnN13 & CST_COMMA & CdCst.RodEndCstmOrder.RodPtnN11, _
                                    strErrCd, hshtStdSize, hshtSelSize, hshtInputDiv, _
                                    Me.strcSelDataInfo.strSelOtherVal) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSSDInpCheck = False
                        Exit Function
                    End If
                    'WF + A寸法チェック
                    If fncStdWFAChk(strActPtn, hshtSelSize, hshtStdSize, hshtFrmPos, Me.strcSelDataInfo.strSelWFMaxVal, strErrCd, "1") = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = Me.strcSelDataInfo.strSelWFMaxVal
                        fncSSDInpCheck = False
                        Exit Function
                    End If
                    '最大WFチェック
                    If fncStdMaxWFChk(hshtSelSize, hshtStdSize, Me.strcSelDataInfo.strSelWFMaxVal, strErrCd, "1") = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = Me.strcSelDataInfo.strSelWFMaxVal
                        fncSSDInpCheck = False
                        Exit Function
                    End If
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            hshtStdSize = Nothing
            hshtSelSize = Nothing
            hshtFrmPos = Nothing
            hshtInputDiv = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <param name="strOpSymbol">選択オプション</param>
    ''' <param name="strKeykataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSCA2InpCheck(ByVal strOpSymbol As String(), ByVal strKeykataban As String) As Boolean

        Dim strActPtn As String
        Dim strActPtnNo As String
        Dim hshtStdSize As New Hashtable
        Dim hshtSelSize As New Hashtable
        Dim hshtFrmPos As New Hashtable
        Dim hshtInputDiv As New Hashtable
        Dim strErrCd As String
        Dim intFrmPos As Integer
        Dim intMinASize As Integer
        Dim intMinKLSize As Integer
        fncSCA2InpCheck = True

        Try
            '初期設定
            strActPtn = CST_BLANK
            strActPtnNo = CST_BLANK
            strErrCd = CST_BLANK
            intFrmPos = 0

            '固定値設定
            intMinASize = 15
            intMinKLSize = 5

            'ロッド先端パターン設定
            Select Case Me.strcSelDataInfo.strSelPtn
                Case CdCst.RodEndCstmOrder.OtherSize
                    If fncRodPtnGet(Me.strcSelDataInfo.strSelOtherVal, strActPtn, strActPtnNo, strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSCA2InpCheck = False
                        Exit Function
                    End If
                Case Else
                    strActPtn = Me.strcSelDataInfo.strSelPtn
                    strActPtnNo = Me.strcSelDataInfo.strSelPtnNo
            End Select

            '選択データセット
            Call subFrmDataSet(strActPtnNo, hshtStdSize, hshtInputDiv, hshtSelSize, hshtFrmPos, Me.strcSelDataInfo.strSelOtherVal)

            Select Case Me.strcSelDataInfo.strSelPtn
                Case CdCst.RodEndCstmOrder.OtherSize
                    'ハイフンチェック
                    If fncOthHypenChk(Me.strcSelDataInfo.strSelOtherVal, _
                                      strErrCd, _
                                      CdCst.RodEndCstmOrder.RodPtnN13N11 & CST_COMMA & CdCst.RodEndCstmOrder.RodPtnN11N13) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSCA2InpCheck = False
                        Exit Function
                    End If
                    'WFの後に数値がなかったらエラー
                    If fncNumericCheck(Me.strcSelDataInfo.strSelOtherVal, CdCst.RodEndCstmOrder.FrmWF, , strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = CdCst.RodEndCstmOrder.FrmWF
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSCA2InpCheck = False
                        Exit Function
                    End If
                    'Aの後に数値がなかったらエラー
                    If fncNumericCheck(Me.strcSelDataInfo.strSelOtherVal, CdCst.RodEndCstmOrder.FrmA, , strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = CdCst.RodEndCstmOrder.FrmA
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSCA2InpCheck = False
                        Exit Function
                    End If
                Case Else
                    'A/KL寸法チェック
                    If fncStdAKLChk(Me.strcSelDataInfo.strSelPtn, hshtSelSize, hshtStdSize, hshtFrmPos, strErrCd, intFrmPos, intMinASize, intMinKLSize) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        Me.strcErrInfo.strErrFocusNo = intFrmPos
                        fncSCA2InpCheck = False
                        Exit Function
                    End If
                    'WF寸法チェック
                    If fncStdWFChk(hshtSelSize, hshtStdSize, hshtFrmPos, strErrCd, intFrmPos) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        Me.strcErrInfo.strErrFocusNo = intFrmPos
                        fncSCA2InpCheck = False
                        Exit Function
                    End If
            End Select

            Select Case strActPtn
                '2012/05/30　N11追加　Y.Tachi
                Case CdCst.RodEndCstmOrder.RodPtnN13N11, CdCst.RodEndCstmOrder.RodPtnN11N13, CdCst.RodEndCstmOrder.RodPtnN11
                Case Else
                    'N13/N11チェック
                    If fncSelectChk(strActPtn, _
                                    CdCst.RodEndCstmOrder.RodPtnN13 & CST_COMMA & CdCst.RodEndCstmOrder.RodPtnN11, _
                                    strErrCd, hshtStdSize, hshtSelSize, hshtInputDiv, _
                                    Me.strcSelDataInfo.strSelOtherVal) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSCA2InpCheck = False
                        Exit Function
                    End If
                    ' WF + A寸法チェック
                    If fncStdWFAChk(strActPtn, hshtSelSize, hshtStdSize, hshtFrmPos, Me.strcSelDataInfo.strSelWFMaxVal, strErrCd, "0") = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = Me.strcSelDataInfo.strSelWFMaxVal
                        fncSCA2InpCheck = False
                        Exit Function
                    End If
                    '最大WFチェック
                    If fncStdMaxWFChk(hshtSelSize, hshtStdSize, Me.strcSelDataInfo.strSelWFMaxVal, strErrCd, "0") = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = Me.strcSelDataInfo.strSelWFMaxVal
                        fncSCA2InpCheck = False
                        Exit Function
                    End If
            End Select

            'WF寸法チェック
            If fncStdWFChk1(hshtSelSize, hshtStdSize, hshtFrmPos, strErrCd, strOpSymbol, strKeykataban, intFrmPos) = False Then
                Me.strcErrInfo.strErrCd = strErrCd
                Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                Me.strcErrInfo.strErrFocusNo = intFrmPos
                fncSCA2InpCheck = False
                Exit Function
            End If

        Catch ex As Exception

            WriteErrorLog("E001", ex)

        Finally

            hshtStdSize = Nothing
            hshtSelSize = Nothing
            hshtFrmPos = Nothing
            hshtInputDiv = Nothing

        End Try
    End Function

    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <param name="strOpSymbol">選択オプション</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncJSC3InpCheck(ByVal strOpSymbol As String()) As Boolean

        Dim strActPtn As String
        Dim strActPtnNo As String
        Dim hshtStdSize As New Hashtable
        Dim hshtSelSize As New Hashtable
        Dim hshtFrmPos As New Hashtable
        Dim hshtInputDiv As New Hashtable
        Dim strErrCd As String
        Dim intFrmPos As Integer
        Dim intMinASize As Integer
        Dim intMinKLSize As Integer
        fncJSC3InpCheck = True
        Try
            '初期設定
            strActPtn = CST_BLANK
            strActPtnNo = CST_BLANK
            strErrCd = CST_BLANK
            intFrmPos = 0

            '固定値設定
            Select Case Me.strcSelection.strKeyKataban
                Case "1"
                    intMinASize = 15
                Case "2"
                    intMinASize = 20
            End Select
            intMinKLSize = 5

            'ロッド先端パターン設定
            Select Case Me.strcSelDataInfo.strSelPtn
                Case CdCst.RodEndCstmOrder.OtherSize
                    If fncRodPtnGet(Me.strcSelDataInfo.strSelOtherVal, strActPtn, strActPtnNo, strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncJSC3InpCheck = False
                        Exit Function
                    End If
                Case Else
                    strActPtn = Me.strcSelDataInfo.strSelPtn
                    strActPtnNo = Me.strcSelDataInfo.strSelPtnNo
            End Select

            '選択データセット
            Call subFrmDataSet(strActPtnNo, hshtStdSize, hshtInputDiv, hshtSelSize, hshtFrmPos, Me.strcSelDataInfo.strSelOtherVal)

            Select Case Me.strcSelDataInfo.strSelPtn
                Case CdCst.RodEndCstmOrder.OtherSize
                    'ハイフンチェック
                    If fncOthHypenChk(Me.strcSelDataInfo.strSelOtherVal, _
                                      strErrCd, _
                                      CdCst.RodEndCstmOrder.RodPtnN13N11 & CST_COMMA & CdCst.RodEndCstmOrder.RodPtnN11N13) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncJSC3InpCheck = False
                        Exit Function
                    End If
                    'WFの後に数値がなかったらエラー
                    If fncNumericCheck(Me.strcSelDataInfo.strSelOtherVal, CdCst.RodEndCstmOrder.FrmWF, , strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = CdCst.RodEndCstmOrder.FrmWF
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncJSC3InpCheck = False
                        Exit Function
                    End If
                    'Aの後に数値がなかったらエラー
                    If fncNumericCheck(Me.strcSelDataInfo.strSelOtherVal, CdCst.RodEndCstmOrder.FrmA, , strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = CdCst.RodEndCstmOrder.FrmA
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncJSC3InpCheck = False
                        Exit Function
                    End If
                Case Else
                    'A/KL寸法チェック
                    If fncStdAKLChk(Me.strcSelDataInfo.strSelPtn, hshtSelSize, hshtStdSize, hshtFrmPos, strErrCd, intFrmPos, intMinASize, intMinKLSize) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        Me.strcErrInfo.strErrFocusNo = intFrmPos
                        fncJSC3InpCheck = False
                        Exit Function
                    End If
                    'WF寸法チェック
                    If fncStdWFChk(hshtSelSize, hshtStdSize, hshtFrmPos, strErrCd, intFrmPos) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        Me.strcErrInfo.strErrFocusNo = intFrmPos
                        fncJSC3InpCheck = False
                        Exit Function
                    End If
            End Select

            Select Case strActPtn
                Case CdCst.RodEndCstmOrder.RodPtnN13N11, CdCst.RodEndCstmOrder.RodPtnN11N13
                Case Else
                    'N13チェック
                    If fncSelectChk(strActPtn, _
                                    CdCst.RodEndCstmOrder.RodPtnN13 & CST_COMMA & CdCst.RodEndCstmOrder.RodPtnN11, _
                                    strErrCd, _
                                    hshtStdSize, _
                                    hshtSelSize, _
                                    hshtInputDiv, _
                                    Me.strcSelDataInfo.strSelOtherVal) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncJSC3InpCheck = False
                        Exit Function
                    End If
                    'WF + A寸法チェック
                    If fncStdWFAChk(strActPtn, hshtSelSize, hshtStdSize, hshtFrmPos, Me.strcSelDataInfo.strSelWFMaxVal, strErrCd, "0") = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = Me.strcSelDataInfo.strSelWFMaxVal
                        fncJSC3InpCheck = False
                        Exit Function
                    End If
                    '最大WFチェック
                    If fncStdMaxWFChk(hshtSelSize, hshtStdSize, Me.strcSelDataInfo.strSelWFMaxVal, strErrCd, "0") = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = Me.strcSelDataInfo.strSelWFMaxVal
                        fncJSC3InpCheck = False
                        Exit Function
                    End If
            End Select

        Catch ex As Exception

            WriteErrorLog("E001", ex)

        Finally

            hshtStdSize = Nothing
            hshtSelSize = Nothing
            hshtFrmPos = Nothing
            hshtInputDiv = Nothing

        End Try

    End Function

    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <param name="strOpSymbol">引当オプション情報</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSCSInpCheck(ByVal strOpSymbol As String()) As Boolean

        Dim strActPtn As String
        Dim strActPtnNo As String
        Dim hshtStdSize As New Hashtable
        Dim hshtSelSize As New Hashtable
        Dim hshtFrmPos As New Hashtable
        Dim hshtInputDiv As New Hashtable
        Dim strWFMaxSize As String
        Dim strErrCd As String
        Dim intFrmPos As Integer
        Dim intMinASize As Integer
        Dim intMinKLSize As Integer
        fncSCSInpCheck = True

        Try
            '初期設定
            strActPtn = CST_BLANK
            strActPtnNo = CST_BLANK
            strWFMaxSize = CST_BLANK
            strErrCd = CST_BLANK
            intFrmPos = 0

            '固定値設定
            intMinASize = 20
            intMinKLSize = 5

            'ロッド先端パターン設定
            Select Case Me.strcSelDataInfo.strSelPtn
                Case CdCst.RodEndCstmOrder.OtherSize
                    If fncRodPtnGet(Me.strcSelDataInfo.strSelOtherVal, strActPtn, strActPtnNo, strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSCSInpCheck = False
                        Exit Function
                    End If
                Case Else
                    strActPtn = Me.strcSelDataInfo.strSelPtn
                    strActPtnNo = Me.strcSelDataInfo.strSelPtnNo
            End Select

            '選択データセット
            Call subFrmDataSet(strActPtnNo, hshtStdSize, hshtInputDiv, hshtSelSize, hshtFrmPos, Me.strcSelDataInfo.strSelOtherVal)

            Select Case Me.strcSelDataInfo.strSelPtn
                Case CdCst.RodEndCstmOrder.OtherSize
                    'ハイフンチェック
                    If fncOthHypenChk(Me.strcSelDataInfo.strSelOtherVal, strErrCd, CdCst.RodEndCstmOrder.RodPtnN13N11) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSCSInpCheck = False
                        Exit Function
                    End If
                    'WFの後に数値がなかったらエラー
                    If fncNumericCheck(Me.strcSelDataInfo.strSelOtherVal, CdCst.RodEndCstmOrder.FrmWF, , strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = CdCst.RodEndCstmOrder.FrmWF
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSCSInpCheck = False
                        Exit Function
                    End If
                    'Aの後に数値がなかったらエラー
                    If fncNumericCheck(Me.strcSelDataInfo.strSelOtherVal, CdCst.RodEndCstmOrder.FrmA, , strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = CdCst.RodEndCstmOrder.FrmA
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSCSInpCheck = False
                        Exit Function
                    End If
                Case Else
                    'A/KL寸法チェック
                    If fncStdAKLChk(Me.strcSelDataInfo.strSelPtn, hshtSelSize, hshtStdSize, hshtFrmPos, strErrCd, intFrmPos, intMinASize, intMinKLSize) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        Me.strcErrInfo.strErrFocusNo = intFrmPos
                        fncSCSInpCheck = False
                        Exit Function
                    End If
                    'WF寸法チェック
                    If fncStdWFChk(Me.strcSelDataInfo.strSelPtn, hshtSelSize, hshtStdSize, hshtFrmPos, strErrCd, intFrmPos) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        Me.strcErrInfo.strErrFocusNo = intFrmPos
                        fncSCSInpCheck = False
                        Exit Function
                    End If
            End Select

            Select Case strActPtn
                Case CdCst.RodEndCstmOrder.RodPtnN13N11, CdCst.RodEndCstmOrder.RodPtnN11N13
                Case Else
                    'N13チェック
                    If fncSelectChk(strActPtn, CdCst.RodEndCstmOrder.RodPtnN13, _
                                    strErrCd, hshtStdSize, hshtSelSize, hshtInputDiv, _
                                    Me.strcSelDataInfo.strSelOtherVal) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncSCSInpCheck = False
                        Exit Function
                    End If
                    ' WF + A寸法チェック
                    If fncStdWFAChk(strActPtn, hshtSelSize, hshtStdSize, hshtFrmPos, Me.strcSelDataInfo.strSelWFMaxVal, strErrCd, "0") = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = Me.strcSelDataInfo.strSelWFMaxVal
                        fncSCSInpCheck = False
                        Exit Function
                    End If
                    '最大WFチェック
                    If fncStdMaxWFChk(hshtSelSize, hshtStdSize, Me.strcSelDataInfo.strSelWFMaxVal, strErrCd, "0") = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = Me.strcSelDataInfo.strSelWFMaxVal
                        fncSCSInpCheck = False
                        Exit Function
                    End If
            End Select

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            hshtStdSize = Nothing
            hshtSelSize = Nothing
            hshtFrmPos = Nothing
            hshtInputDiv = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 入力チェック
    ''' </summary>
    ''' <param name="strOpSymbol">引当オプション情報</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCMK2InpCheck(ByVal strOpSymbol As String()) As Boolean
        Dim strActPtn As String
        Dim strActPtnNo As String
        Dim hshtStdSize As New Hashtable
        Dim hshtSelSize As New Hashtable
        Dim hshtFrmPos As New Hashtable
        Dim hshtInputDiv As New Hashtable
        Dim strWFMaxSize As String
        Dim strErrCd As String
        Dim intFrmPos As Integer
        Dim intMinASize As Integer
        Dim intMinKLSize As Integer
        fncCMK2InpCheck = True

        Try
            '初期設定
            strActPtn = CST_BLANK
            strActPtnNo = CST_BLANK
            strWFMaxSize = CST_BLANK
            strErrCd = CST_BLANK
            intFrmPos = 0

            '固定値設定
            intMinASize = 20
            intMinKLSize = 5

            'ロッド先端パターン設定
            Select Case Me.strcSelDataInfo.strSelPtn
                Case CdCst.RodEndCstmOrder.OtherSize
                    If fncRodPtnGet(Me.strcSelDataInfo.strSelOtherVal, strActPtn, strActPtnNo, strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncCMK2InpCheck = False
                        Exit Function
                    End If
                Case Else
                    strActPtn = Me.strcSelDataInfo.strSelPtn
                    strActPtnNo = Me.strcSelDataInfo.strSelPtnNo
            End Select

            '選択データセット
            Call subFrmDataSet(strActPtnNo, hshtStdSize, hshtInputDiv, hshtSelSize, hshtFrmPos, Me.strcSelDataInfo.strSelOtherVal)

            Select Case Me.strcSelDataInfo.strSelPtn
                Case CdCst.RodEndCstmOrder.OtherSize
                    'ハイフンチェック
                    If fncOthHypenChk(Me.strcSelDataInfo.strSelOtherVal, strErrCd, CdCst.RodEndCstmOrder.RodPtnN13N11) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncCMK2InpCheck = False
                        Exit Function
                    End If
                    'WFの後に数値がなかったらエラー
                    If fncNumericCheck(Me.strcSelDataInfo.strSelOtherVal, CdCst.RodEndCstmOrder.FrmWF, , strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = CdCst.RodEndCstmOrder.FrmWF
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncCMK2InpCheck = False
                        Exit Function
                    End If
                    'Aの後に数値がなかったらエラー
                    If fncNumericCheck(Me.strcSelDataInfo.strSelOtherVal, CdCst.RodEndCstmOrder.FrmA, , strErrCd) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = CdCst.RodEndCstmOrder.FrmA
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncCMK2InpCheck = False
                        Exit Function
                    End If
                Case Else
                    'A/KL寸法チェック
                    If fncStdAKLChk(Me.strcSelDataInfo.strSelPtn, hshtSelSize, hshtStdSize, hshtFrmPos, strErrCd, intFrmPos, intMinASize, intMinKLSize) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        Me.strcErrInfo.strErrFocusNo = intFrmPos
                        fncCMK2InpCheck = False
                        Exit Function
                    End If
                    'WF寸法チェック
                    If fncStdWFChk(Me.strcSelDataInfo.strSelPtn, hshtSelSize, hshtStdSize, hshtFrmPos, strErrCd, intFrmPos) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        Me.strcErrInfo.strErrFocusNo = intFrmPos
                        fncCMK2InpCheck = False
                        Exit Function
                    End If
            End Select

            Select Case strActPtn
                Case CdCst.RodEndCstmOrder.RodPtnN13N11, CdCst.RodEndCstmOrder.RodPtnN11N13
                Case Else
                    'N13チェック
                    If fncSelectChk(strActPtn, CdCst.RodEndCstmOrder.RodPtnN13, _
                                    strErrCd, hshtStdSize, hshtSelSize, hshtInputDiv, _
                                    Me.strcSelDataInfo.strSelOtherVal) = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrPtnNo = Me.strcSelDataInfo.strSelPtnNo
                        Me.strcErrInfo.strErrPtn = Me.strcSelDataInfo.strSelPtn
                        fncCMK2InpCheck = False
                        Exit Function
                    End If
                    ' WF + A寸法チェック
                    If fncStdWFAChk(strActPtn, hshtSelSize, hshtStdSize, hshtFrmPos, Me.strcSelDataInfo.strSelWFMaxVal, strErrCd, "1") = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = Me.strcSelDataInfo.strSelWFMaxVal
                        fncCMK2InpCheck = False
                        Exit Function
                    End If
                    '最大WFチェック
                    If fncStdMaxWFChk(hshtSelSize, hshtStdSize, Me.strcSelDataInfo.strSelWFMaxVal, strErrCd, "1") = False Then
                        Me.strcErrInfo.strErrCd = strErrCd
                        Me.strcErrInfo.strErrOption = Me.strcSelDataInfo.strSelWFMaxVal
                        fncCMK2InpCheck = False
                        Exit Function
                    End If
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            hshtStdSize = Nothing
            hshtSelSize = Nothing
            hshtFrmPos = Nothing
            hshtInputDiv = Nothing
        End Try
    End Function

    ''' <summary>
    ''' その他寸法からロッド先端パターンを取得する
    ''' </summary>
    ''' <param name="strOtherSize">その他寸法</param>
    ''' <param name="strSelRodPtn">選択ロッドパターン記号</param>
    ''' <param name="strSelRodPtnNo">選択ロッドパターンNo.</param>
    ''' <param name="strErrCd">エラーコード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncRodPtnGet(ByVal strOtherSize As String, _
                                  ByRef strSelRodPtn As String, _
                                  ByRef strSelRodPtnNo As String, _
                                  ByRef strErrCd As String) As Boolean
        Dim strRodPtn() As String
        Dim strRodPtnNo() As String
        Dim strIntend As String
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer
        Dim intRodCnt As Integer
        fncRodPtnGet = False

        Try
            '配列定義
            ReDim strRodPtn(0)
            ReDim strRodPtnNo(0)

            intRodCnt = 0

            'エラーメッセージ設定
            strErrCd = "W8480"

            'ロッド先端特注パターン記号を確定するため、ありえるロッド先端特注パターン記号を抽出し、長さが長い方から配列に格納する
            For intLoopCnt1 = 1 To UBound(Me.strcRodPtnInfo)
                If InStr(strOtherSize, Me.strcRodPtnInfo(intLoopCnt1).strRodPtn) > 0 Then
                    ReDim Preserve strRodPtn(UBound(strRodPtn) + 1)
                    ReDim Preserve strRodPtnNo(UBound(strRodPtnNo) + 1)
                    If UBound(strRodPtn) = 1 Then
                        strRodPtn(1) = Me.strcRodPtnInfo(intLoopCnt1).strRodPtn
                        strRodPtnNo(1) = Me.strcRodPtnInfo(intLoopCnt1).strDispNo
                    Else
                        For intLoopCnt2 = UBound(strRodPtn) - 1 To 1 Step -1
                            If strRodPtn(intLoopCnt2).Trim.Length < Me.strcRodPtnInfo(intLoopCnt1).strRodPtn.Trim.Length Then
                                strRodPtn(intLoopCnt2 + 1) = strRodPtn(intLoopCnt2)
                                strRodPtnNo(intLoopCnt2 + 1) = strRodPtnNo(intLoopCnt2)
                                If intLoopCnt2 = 1 Then
                                    strRodPtn(intLoopCnt2) = Me.strcRodPtnInfo(intLoopCnt1).strRodPtn.Trim
                                    strRodPtnNo(intLoopCnt2) = Me.strcRodPtnInfo(intLoopCnt1).strDispNo.Trim
                                End If
                            Else
                                strRodPtn(intLoopCnt2 + 1) = Me.strcRodPtnInfo(intLoopCnt1).strRodPtn.Trim
                                strRodPtnNo(intLoopCnt2 + 1) = Me.strcRodPtnInfo(intLoopCnt1).strDispNo.Trim
                                Exit For
                            End If
                        Next
                    End If
                End If
            Next
            'ありえるロッド選択特注パターン中の一番文字列の長いロッド先端特注パターン記号が、その他寸法の最も左にあればOK
            If UBound(strRodPtn) > 0 Then
                If InStr(Left(strOtherSize, strRodPtn(1).Trim.Length), strRodPtn(1)) > 0 Then
                    strSelRodPtn = strRodPtn(1)
                    strSelRodPtnNo = strRodPtnNo(1)
                    'ロッド先端特注パターン記号を2つ以上入力していないかチェックする
                    strIntend = strOtherSize
                    For intLoopCnt1 = 1 To UBound(strRodPtn)
                        If InStr(strIntend, strRodPtn(intLoopCnt1)) > 0 Then
                            strIntend = Mid(strIntend, 1, InStr(1, strIntend, strRodPtn(intLoopCnt1)) - 1) & _
                                        Mid(strIntend, InStr(1, strIntend, strRodPtn(intLoopCnt1)) + strRodPtn(intLoopCnt1).Length, strIntend.Length)
                            intRodCnt = intRodCnt + 1
                        End If
                    Next
                End If
                If intRodCnt = 1 Then
                    fncRodPtnGet = True
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            strRodPtn = Nothing
            strRodPtnNo = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 選択データをセットする
    ''' </summary>
    ''' <param name="strSelRodPtnNo">ロッド先端パターンNo.</param>
    ''' <param name="hshtStdSize">標準寸法情報</param>
    ''' <param name="hshtSelSize">外径種類位置情報</param>
    ''' <param name="hshtFrmPos">入力区分情報</param>
    ''' <param name="hshtInputDiv">特注寸法情報</param>
    ''' <param name="strOtherSize">その他寸法</param>
    ''' <remarks></remarks>
    Private Sub subFrmDataSet(ByVal strSelRodPtnNo As String, _
                              ByRef hshtStdSize As Hashtable, _
                              ByRef hshtInputDiv As Hashtable, _
                              ByRef hshtSelSize As Hashtable, _
                              Optional ByRef hshtFrmPos As Hashtable = Nothing, _
                              Optional ByVal strOtherSize As String = Nothing)
        Dim strSelLength As String
        Dim intLoopCnt As Integer
        Try
            strSelLength = CST_BLANK
            If strOtherSize Is Nothing Then
                '標準寸法/外径種類位置情報/入力区分情報セット
                For intLoopCnt = 1 To UBound(Me.strcRodPtnInfo(strSelRodPtnNo).strExtFrm)
                    hshtStdSize.Add(Me.strcRodPtnInfo(strSelRodPtnNo).strExtFrm(intLoopCnt), Me.strcRodPtnInfo(strSelRodPtnNo).strNormalVal(intLoopCnt))
                    hshtFrmPos.Add(Me.strcRodPtnInfo(strSelRodPtnNo).strExtFrm(intLoopCnt), intLoopCnt)
                    hshtInputDiv.Add(Me.strcRodPtnInfo(strSelRodPtnNo).strExtFrm(intLoopCnt), Me.strcRodPtnInfo(strSelRodPtnNo).strInputDiv(intLoopCnt))
                Next
                '特注寸法セット
                hshtSelSize = Me.strcSelDataInfo.hshtSelVal
            Else
                '標準寸法セット
                For intLoopCnt = 1 To UBound(Me.strcRodPtnInfo(strSelRodPtnNo).strExtFrm)
                    hshtStdSize.Add(Me.strcRodPtnInfo(strSelRodPtnNo).strExtFrm(intLoopCnt), Me.strcRodPtnInfo(strSelRodPtnNo).strNormalVal(intLoopCnt))
                    hshtInputDiv.Add(Me.strcRodPtnInfo(strSelRodPtnNo).strExtFrm(intLoopCnt), Me.strcRodPtnInfo(strSelRodPtnNo).strInputDiv(intLoopCnt))
                Next
                '特注寸法セット
                For intLoopCnt = 1 To UBound(Me.strcRodPtnInfo(strSelRodPtnNo).strExtFrm)
                    If hshtStdSize.ContainsKey(Me.strcRodPtnInfo(strSelRodPtnNo).strExtFrm(intLoopCnt)) Then
                        '特注寸法設定
                        If fncNumericCheck(strOtherSize, Me.strcRodPtnInfo(strSelRodPtnNo).strExtFrm(intLoopCnt), strSelLength) = True Then
                            hshtSelSize(Me.strcRodPtnInfo(strSelRodPtnNo).strExtFrm(intLoopCnt)) = strSelLength
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' ・入力チェック(外径寸法の後に数値が入っているかどうかをチェックする)
    ''' ・外径寸法の後の数値を返却する
    ''' </summary>
    ''' <param name="strOtherSize">その他寸法</param>
    ''' <param name="strFrmPtn">外径寸法</param>
    ''' <param name="strLength">指定外径寸法の長さ</param>
    ''' <param name="strErrCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncNumericCheck(ByVal strOtherSize As String, _
                                     ByVal strFrmPtn As String, _
                                     Optional ByRef strLength As String = CST_BLANK, _
                                     Optional ByRef strErrCd As String = CST_BLANK) As Boolean
        Dim intLoopCnt As Integer
        Dim bolFlg As Boolean = False
        fncNumericCheck = True
        Try
            'エラーメッセージ設定
            strErrCd = "W0130"
            If InStr(strOtherSize, strFrmPtn) > 0 Then
                For intLoopCnt = InStr(1, strOtherSize.Trim, strFrmPtn) + strFrmPtn.Trim.Length To Len(strOtherSize.Trim)
                    If Mid(strOtherSize.Trim, intLoopCnt, 1) = "0" Or Mid(strOtherSize.Trim, intLoopCnt, 1) = "1" Or _
                       Mid(strOtherSize.Trim, intLoopCnt, 1) = "2" Or Mid(strOtherSize.Trim, intLoopCnt, 1) = "3" Or _
                       Mid(strOtherSize.Trim, intLoopCnt, 1) = "4" Or Mid(strOtherSize.Trim, intLoopCnt, 1) = "5" Or _
                       Mid(strOtherSize.Trim, intLoopCnt, 1) = "6" Or Mid(strOtherSize.Trim, intLoopCnt, 1) = "7" Or _
                       Mid(strOtherSize.Trim, intLoopCnt, 1) = "8" Or Mid(strOtherSize.Trim, intLoopCnt, 1) = "9" Or _
                       Mid(strOtherSize.Trim, intLoopCnt, 1) = "." Then
                        bolFlg = True
                    Else
                        Exit For
                    End If
                Next
                If bolFlg = True Then
                    strLength = Mid(strOtherSize, InStr(1, strOtherSize, strFrmPtn) + strFrmPtn.Length, (intLoopCnt - 1) - (InStr(1, strOtherSize, strFrmPtn) + strFrmPtn.Length) + 1)
                Else
                    fncNumericCheck = False
                    Exit Function
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 除外文字以外にハイフンがないかどうかチェックする
    ''' </summary>
    ''' <param name="strOtherSize">その他寸法</param>
    ''' <param name="strErrCd">エラーコード</param>
    ''' <param name="strExceptChar">除外文字</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncOthHypenChk(ByVal strOtherSize As String, ByRef strErrCd As String, _
                                    Optional ByVal strExceptChar As String = CST_BLANK) As Boolean
        Dim strIntend As String
        Dim strFrmArray() As String
        Dim intLoopCnt As Integer
        fncOthHypenChk = True
        Try
            'エラーメッセージ設定
            strErrCd = "W8570"
            strIntend = strOtherSize
            If strExceptChar <> CST_BLANK Then
                strFrmArray = Split(strExceptChar, CST_COMMA)
                For intLoopCnt = 0 To UBound(strFrmArray)
                    If InStr(1, strIntend, strFrmArray(intLoopCnt)) > 0 Then
                        strIntend = Mid(strIntend, 1, InStr(1, strIntend, strFrmArray(intLoopCnt)) - 1) & _
                                    Mid(strIntend, InStr(1, strIntend, strFrmArray(intLoopCnt)) + strFrmArray(intLoopCnt).Length, strIntend.Length)
                    End If
                Next
            End If
            If InStr(1, strIntend, CdCst.Sign.Hypen) Then
                fncOthHypenChk = False
                Exit Function
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            strFrmArray = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 特定ロッド先端パターン記号において特注寸法が指定されているかチェックする
    ''' </summary>
    ''' <param name="strSelRodPtn">ロッドパターン記号</param>
    ''' <param name="strSpecPtn">指定ロッドパターン</param>
    ''' <param name="strErrCd">エラーコード</param>
    ''' <param name="hshtStdSize">標準寸法情報</param>
    ''' <param name="hshtSelSize">特注寸法情報</param>
    ''' <param name="hshtInputDiv"></param>
    ''' <param name="strOtherSize">その他寸法</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSelectChk(ByVal strSelRodPtn As String, ByVal strSpecPtn As String, _
                                  ByRef strErrCd As String, _
                                  Optional ByVal hshtStdSize As Hashtable = Nothing, _
                                  Optional ByVal hshtSelSize As Hashtable = Nothing, _
                                  Optional ByVal hshtInputDiv As Hashtable = Nothing, _
                                  Optional ByVal strOtherSize As String = Nothing) As Boolean
        Dim strRodPtnArray() As String
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer
        Dim bolFlg As Boolean
        fncSelectChk = True
        Try
            'エラーメッセージ設定
            strErrCd = "W0110"
            strRodPtnArray = Split(strSpecPtn, CST_COMMA)
            If strOtherSize IsNot Nothing Then
                For intLoopCnt1 = 0 To UBound(strRodPtnArray)
                    If strSelRodPtn = strRodPtnArray(intLoopCnt1) Then
                        If strOtherSize.Trim.Length = strSelRodPtn.Trim.Length Then
                            fncSelectChk = False
                            Exit For
                        End If
                    End If
                Next
            Else
                bolFlg = False
                For intLoopCnt1 = 0 To UBound(strRodPtnArray)
                    If strSelRodPtn = strRodPtnArray(intLoopCnt1) Then
                        For intLoopCnt2 = 1 To UBound(Me.strcRodPtnInfo(Me.strcSelDataInfo.strSelPtnNo).strExtFrm)
                            If hshtSelSize.ContainsKey(Me.strcRodPtnInfo(Me.strcSelDataInfo.strSelPtnNo).strExtFrm(intLoopCnt2)) Then
                                If (hshtInputDiv(Me.strcRodPtnInfo(Me.strcSelDataInfo.strSelPtnNo).strExtFrm(intLoopCnt2)) = CdCst.RodEndCstmOrder.Text Or _
                                    hshtInputDiv(Me.strcRodPtnInfo(Me.strcSelDataInfo.strSelPtnNo).strExtFrm(intLoopCnt2)) = CdCst.RodEndCstmOrder.Drop) And _
                                   (hshtSelSize(Me.strcRodPtnInfo(Me.strcSelDataInfo.strSelPtnNo).strExtFrm(intLoopCnt2)) <> hshtStdSize(Me.strcRodPtnInfo(Me.strcSelDataInfo.strSelPtnNo).strExtFrm(intLoopCnt2))) Then
                                    bolFlg = True
                                    Exit Function
                                End If
                            End If
                        Next
                        If bolFlg = False Then
                            fncSelectChk = False
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            strRodPtnArray = Nothing
        End Try
    End Function

    ''' <summary>
    ''' A寸法とKL寸法をチェックする
    ''' </summary>
    ''' <param name="strSelRodPtn">ロッドパターン記号</param>
    ''' <param name="hshtSelSize">特注寸法情報</param>
    ''' <param name="hshtStdSize">標準寸法情報</param>
    ''' <param name="hshtFrmPos">外径種類位置情報</param>
    ''' <param name="strErrCd">エラーコード</param>
    ''' <param name="intFrmPos">外径寸法位置</param>
    ''' <param name="intMinASize">最小A寸法</param>
    ''' <param name="intMinKLSize">最小KL寸法</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncStdAKLChk(ByVal strSelRodPtn As String, ByVal hshtSelSize As Hashtable, _
                                  ByVal hshtStdSize As Hashtable, ByVal hshtFrmPos As Hashtable, _
                                  ByRef strErrCd As String, _
                                  Optional ByRef intFrmPos As Integer = 0, _
                                  Optional ByVal intMinASize As Integer = 0, _
                                  Optional ByVal intMinKLSize As Integer = 0) As Boolean
        Dim strStdASize As String
        Dim strStdKLSize As String
        Dim strSelASize As String
        Dim strSelKLSize As String
        fncStdAKLChk = True
        Try
            '初期設定
            strStdASize = CST_BLANK
            strStdKLSize = CST_BLANK
            strSelASize = CST_BLANK
            strSelKLSize = CST_BLANK
            '選択データセット
            If hshtSelSize.ContainsKey(CdCst.RodEndCstmOrder.FrmA) Then
                strSelASize = hshtSelSize(CdCst.RodEndCstmOrder.FrmA).trim
            End If
            If hshtStdSize.ContainsKey(CdCst.RodEndCstmOrder.FrmA) Then
                strStdASize = hshtStdSize(CdCst.RodEndCstmOrder.FrmA).trim
            End If
            If hshtSelSize.ContainsKey(CdCst.RodEndCstmOrder.FrmKL) Then
                strSelKLSize = hshtSelSize(CdCst.RodEndCstmOrder.FrmKL).trim
            End If
            If hshtStdSize.ContainsKey(CdCst.RodEndCstmOrder.FrmKL) Then
                strStdKLSize = hshtStdSize(CdCst.RodEndCstmOrder.FrmKL).trim
            End If
            Select Case strSelRodPtn
                Case CdCst.RodEndCstmOrder.RodPtnN13, CdCst.RodEndCstmOrder.RodPtnN15
                    'A寸法チェック
                    If strSelASize.Length <> 0 Then
                        If strSelASize <> strStdASize Then
                            If (intMinASize < CDbl(strSelASize)) And _
                               (CDbl(strSelASize) < CDbl(strStdASize) * 2) Then
                            Else
                                strErrCd = "W8440"
                                intFrmPos = hshtFrmPos(CdCst.RodEndCstmOrder.FrmA)
                                fncStdAKLChk = False
                                Exit Function
                            End If
                        End If
                    End If
                Case CdCst.RodEndCstmOrder.RodPtnN11, CdCst.RodEndCstmOrder.RodPtnN1
                    ' KL寸法チェック
                    If strSelKLSize.Length <> 0 Then
                        If strStdKLSize <> strSelKLSize Then
                            If (intMinKLSize < CDbl(strSelKLSize)) And _
                               (CDbl(strSelKLSize) < CDbl(strStdKLSize) * 1.5) Then
                            Else
                                strErrCd = "W8450"
                                intFrmPos = hshtFrmPos(CdCst.RodEndCstmOrder.FrmKL)
                                fncStdAKLChk = False
                                Exit Function
                            End If
                        End If
                    End If
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' WF+A寸法をチェックする
    ''' </summary>
    ''' <param name="strSelRodPtn">ロッドパターン記号</param>
    ''' <param name="hshtSelSize">特注寸法情報</param>
    ''' <param name="hshtStdSize">標準寸法情報</param>
    ''' <param name="hshtFrmPos">外径種類位置情報</param>
    ''' <param name="strWFMaxSize">最大WF</param>
    ''' <param name="strErrCd">エラーコード</param>
    ''' <param name="strMaxDiv">最大値区分(0:最大WF寸法,1:標準寸法+最大WF寸法)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncStdWFAChk(ByVal strSelRodPtn As String, ByVal hshtSelSize As Hashtable, _
                                  ByVal hshtStdSize As Hashtable, ByVal hshtFrmPos As Hashtable, _
                                  ByVal strWFMaxSize As String, ByRef strErrCd As String, _
                                  ByVal strMaxDiv As String) As Boolean
        Dim strRodStroke1 As String
        Dim strRodStroke2 As String
        Dim strStdWFSize As String
        Dim strStdASize As String
        Dim strSelWFSize As String
        Dim strSelASize As String
        fncStdWFAChk = True
        Try
            '初期設定
            strStdWFSize = CST_BLANK
            strStdASize = CST_BLANK
            strSelWFSize = CST_BLANK
            strSelASize = CST_BLANK

            'エラーメッセージ設定
            If strMaxDiv = "0" Then
                strErrCd = "W8470"
            Else
                strErrCd = "W8460"
            End If

            '選択データセット
            If hshtSelSize.ContainsKey(CdCst.RodEndCstmOrder.FrmWF) Then
                strSelWFSize = hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim
            End If
            If hshtStdSize.ContainsKey(CdCst.RodEndCstmOrder.FrmWF) Then
                strStdWFSize = hshtStdSize(CdCst.RodEndCstmOrder.FrmWF).trim
            End If
            If hshtSelSize.ContainsKey(CdCst.RodEndCstmOrder.FrmA) Then
                strSelASize = hshtSelSize(CdCst.RodEndCstmOrder.FrmA).trim
            End If
            If hshtStdSize.ContainsKey(CdCst.RodEndCstmOrder.FrmA) Then
                strStdASize = hshtStdSize(CdCst.RodEndCstmOrder.FrmA).trim
            End If

            'チェック
            If (strSelWFSize.Length <> 0 And (strSelWFSize <> strStdWFSize)) Or _
               (strSelASize.Length <> 0 And (strSelASize <> strStdASize)) Then
                'WF寸法セット
                If strSelWFSize.Length <> 0 And (strSelWFSize <> strStdWFSize) Then
                    'WFに値がある場合
                    strRodStroke1 = hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim
                Else
                    'WFに値がない場合、標準ストロークをセット
                    strRodStroke1 = hshtStdSize(CdCst.RodEndCstmOrder.FrmWF).trim
                End If
                Select Case strSelRodPtn
                    Case CdCst.RodEndCstmOrder.RodPtnN13, CdCst.RodEndCstmOrder.RodPtnN15
                        'A寸法セット
                        If strSelASize.Length <> 0 And (strSelASize <> strStdASize) Then
                            'Aに値がある場合
                            strRodStroke2 = hshtSelSize(CdCst.RodEndCstmOrder.FrmA).trim
                        Else
                            'Aに値がない場合、標準ストロークをセット
                            strRodStroke2 = hshtStdSize(CdCst.RodEndCstmOrder.FrmA).trim
                        End If
                        If CDbl(strRodStroke1) + CDbl(strRodStroke2) > CDbl(strStdWFSize) + CDbl(strStdASize) + strWFMaxSize Then
                            fncStdWFAChk = False
                            Exit Function
                        End If
                    Case Else
                        If CDbl(strRodStroke1) > CDbl(strStdWFSize) + strWFMaxSize Then
                            fncStdWFAChk = False
                            Exit Function
                        End If
                End Select
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' WF寸法の最大値をチェックする
    ''' </summary>
    ''' <param name="hshtSelSize">特注寸法情報</param>
    ''' <param name="hshtStdSize">標準寸法情報</param>
    ''' <param name="strWFMaxVal">最大WF寸法</param>
    ''' <param name="strErrCd">エラーコード</param>
    ''' <param name="strMaxDiv">最大値区分(0:最大WF寸法,1:標準寸法+最大WF寸法)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncStdMaxWFChk(ByVal hshtSelSize As Hashtable, ByVal hshtStdSize As Hashtable, _
                                    ByVal strWFMaxVal As String, ByRef strErrCd As String, _
                                    ByVal strMaxDiv As String) As Boolean
        Dim intMaxWfSize As Double
        fncStdMaxWFChk = True
        Try
            If hshtSelSize.ContainsKey(CdCst.RodEndCstmOrder.FrmWF) Then
                'エラーメッセージ/最大WF値設定
                If strMaxDiv = "0" Then
                    intMaxWfSize = CDbl(strWFMaxVal)
                    strErrCd = "W8470"
                Else
                    intMaxWfSize = CDbl(strWFMaxVal) + CDbl(hshtStdSize(CdCst.RodEndCstmOrder.FrmWF).trim)
                    strErrCd = "W8460"
                End If
                If hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim.length <> 0 Then
                    If hshtStdSize(CdCst.RodEndCstmOrder.FrmWF).trim <> hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim Then
                        If CDbl(hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim) > intMaxWfSize Then
                            fncStdMaxWFChk = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' WF寸法をチェックする(SCS用)
    ''' </summary>
    ''' <param name="strSelRodPtn">ロッドパターン記号</param>
    ''' <param name="hshtSelSize">特注寸法情報</param>
    ''' <param name="hshtStdSize">標準寸法情報</param>
    ''' <param name="hshtFrmPos">外径種類位置情報</param>
    ''' <param name="strErrCd">エラーコード</param>
    ''' <param name="intFrmPos">外径種類位置</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Function fncStdWFChk(ByVal strSelRodPtn As String, _
                                           ByVal hshtSelSize As Hashtable, _
                                           ByVal hshtStdSize As Hashtable, _
                                           ByVal hshtFrmPos As Hashtable, _
                                           ByRef strErrCd As String, _
                                           Optional ByRef intFrmPos As Integer = 0) As Boolean
        fncStdWFChk = True
        Try
            'エラーメッセージ設定
            strErrCd = "W8460"
            If hshtSelSize.ContainsKey(CdCst.RodEndCstmOrder.FrmWF) Then
                If hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim.length <> 0 Then
                    Select Case strSelRodPtn
                        Case CdCst.RodEndCstmOrder.RodPtnN13, CdCst.RodEndCstmOrder.RodPtnN15, _
                             CdCst.RodEndCstmOrder.RodPtnN11, CdCst.RodEndCstmOrder.RodPtnN1, _
                             CdCst.RodEndCstmOrder.RodPtnN31, CdCst.RodEndCstmOrder.RodPtnN2, _
                             CdCst.RodEndCstmOrder.RodPtnN21
                            ' WF寸法チェック
                            If hshtStdSize(CdCst.RodEndCstmOrder.FrmWF).trim <> hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim Then
                                If CDbl(hshtStdSize(CdCst.RodEndCstmOrder.FrmWF).trim) < CDbl(hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim) Then
                                Else
                                    intFrmPos = hshtFrmPos(CdCst.RodEndCstmOrder.FrmWF)
                                    fncStdWFChk = False
                                    Exit Function
                                End If
                            End If
                        Case CdCst.RodEndCstmOrder.RodPtnN12, CdCst.RodEndCstmOrder.RodPtnN14, CdCst.RodEndCstmOrder.RodPtnN3
                            ' WF寸法チェック
                            If hshtStdSize(CdCst.RodEndCstmOrder.FrmWF).trim <> hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim Then
                                If CDbl(hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim) >= 1 Then
                                Else
                                    intFrmPos = hshtFrmPos(CdCst.RodEndCstmOrder.FrmWF)
                                    fncStdWFChk = False
                                    Exit Function
                                End If
                            End If
                    End Select
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' WF寸法をチェックする(SCS以外用)
    ''' </summary>
    ''' <param name="hshtSelSize">特注寸法情報</param>
    ''' <param name="hshtStdSize">標準寸法情報</param>
    ''' <param name="hshtFrmPos">外径種類位置情報</param>
    ''' <param name="strErrCd">エラーコード</param>
    ''' <param name="intFrmPos">外径種類位置</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Function fncStdWFChk(ByVal hshtSelSize As Hashtable, ByVal hshtStdSize As Hashtable, _
                                           ByVal hshtFrmPos As Hashtable, ByRef strErrCd As String, _
                                           Optional ByRef intFrmPos As Integer = 0) As Boolean
        fncStdWFChk = True
        Try
            'エラーメッセージ設定
            strErrCd = "W8460"
            If hshtSelSize.ContainsKey(CdCst.RodEndCstmOrder.FrmWF) Then
                If hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim.length <> 0 Then
                    If hshtStdSize(CdCst.RodEndCstmOrder.FrmWF).trim <> hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim Then
                        If CDbl(hshtStdSize(CdCst.RodEndCstmOrder.FrmWF).trim) < CDbl(hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim) Then
                        Else
                            intFrmPos = hshtFrmPos(CdCst.RodEndCstmOrder.FrmWF)
                            fncStdWFChk = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' WF寸法をチェックする
    ''' </summary>
    ''' <param name="hshtSelSize">特注寸法情報</param>
    ''' <param name="hshtStdSize">標準寸法情報</param>
    ''' <param name="hshtFrmPos">外径種類位置情報</param>
    ''' <param name="strErrCd">エラーコード</param>
    ''' <param name="intFrmPos">外径種類位置</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Overloads Function fncStdWFChk1(ByVal hshtSelSize As Hashtable, ByVal hshtStdSize As Hashtable, _
                                          ByVal hshtFrmPos As Hashtable, ByRef strErrCd As String, _
                                          ByVal strOpSymbol As String(), ByRef strKeykataban As String, _
                                          Optional ByRef intFrmPos As Integer = 0) As Boolean
        fncStdWFChk1 = True
        Try
            'エラーメッセージ設定
            strErrCd = "W8980"
            If hshtSelSize.ContainsKey(CdCst.RodEndCstmOrder.FrmWF) Then
                If hshtSelSize(CdCst.RodEndCstmOrder.FrmWF).trim.length <> 0 Then
                    Select Case strKeykataban
                        Case "", "V", "2"
                            If InStr(1, strOpSymbol(13), "J") <> 0 Or InStr(1, strOpSymbol(13), "L") <> 0 Then
                                intFrmPos = 13
                                fncStdWFChk1 = False
                            End If
                        Case "B", "C"
                            If InStr(1, strOpSymbol(17), "J") <> 0 Or InStr(1, strOpSymbol(17), "L") <> 0 Then
                                intFrmPos = 17
                                fncStdWFChk1 = False
                            End If
                        Case "D", "E"
                            If InStr(1, strOpSymbol(12), "J") <> 0 Or InStr(1, strOpSymbol(12), "L") <> 0 Then
                                intFrmPos = 12
                                fncStdWFChk1 = False
                            End If
                    End Select

                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

#End Region

End Class
