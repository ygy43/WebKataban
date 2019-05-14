Imports System.Data.SqlClient
Imports WebKataban.ClsCommon

<Serializable()>
Public Class KHKtbnStrc

#Region " Definition "
    '全ての選択情報
    Public strcSelection As New KHInfoModel
#End Region

    ''' <summary>
    ''' 選択した形番情報を取得する
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Public Sub subSelKtbnInfoGet(objCon As SqlConnection, ByVal strUserId As String, _
                                 ByVal strSessionId As String, Optional intMode As Integer = 0)

        Try
            '初期化
            Me.strcSelection = New KHInfoModel(strUserId, strSessionId)

            '引当シリーズ形番検索
            Call Me.subSelSrsKtbnSelect(objCon)
            '引当形番構成検索
            Call Me.subSelKtbnStrcSelect(objCon)

            Select Case intMode
                Case 0 '形番引当画面用
                Case 1 'それ以外
                    '引当積上単価検索
                    Call Me.subSelAccPriceStrcSelect(objCon)

                    '仕様書機種の場合
                    If Me.strcSelection.strSpecNo.Trim <> "" And Me.strcSelection.strSpecNo.Trim <> "00" Then
                        '引当仕様書情報検索
                        Call Me.subSelSpecSelect(objCon)
                        '引当仕様書構成検索
                        Call Me.subSelSpecStrcSelect(objCon)
                    End If
                    '引当ロッド先端特注検索
                    Call Me.subSelRodWFSelect(objCon)
            End Select

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当積上単価構成追加処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strKatabanCheckDiv">形番チェック区分</param>
    ''' <param name="strPlaceCd">出荷場所コード</param>
    ''' <param name="intListPrice">定価</param>
    ''' <param name="intRegPrice">登録店価格</param>
    ''' <param name="intSsprice">SS店価格</param>
    ''' <param name="intBsprice">BS店価格</param>
    ''' <param name="intGsprice">GS店価格</param>
    ''' <param name="intPsprice">PS店価格</param>
    ''' <param name="decAmount">数量</param>
    ''' <param name="strCurrency"></param>
    ''' <param name="strMadeCountry"></param>
    ''' <remarks></remarks>
    Public Sub subInsertAccPriceStrc(objCon As SqlConnection, ByVal strUserId As String, ByVal strSessionId As String, _
                                     ByVal strKataban() As String, ByVal strKatabanCheckDiv() As String, _
                                     ByVal strPlaceCd() As String, ByVal intListPrice() As Decimal, _
                                     ByVal intRegPrice() As Decimal, ByVal intSsprice() As Decimal, _
                                     ByVal intBsprice() As Decimal, ByVal intGsprice() As Decimal, _
                                     ByVal intPsprice() As Decimal, ByVal decAmount() As Decimal, _
                                     ByVal strCurrency As String, ByVal strMadeCountry As String)
        Dim intLoopCnt As Integer
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            'オプション形番分繰り返し
            For intLoopCnt = 1 To strKataban.Length - 1
                '引当積上単価構成追加
                Call dalKtbnStrc.subAccPriceStrcMnt(objCon, strUserId, strSessionId, intLoopCnt, strKataban(intLoopCnt), _
                                           strKatabanCheckDiv(intLoopCnt), strPlaceCd(intLoopCnt), _
                                           intListPrice(intLoopCnt), intRegPrice(intLoopCnt), _
                                           intSsprice(intLoopCnt), intBsprice(intLoopCnt), _
                                           intGsprice(intLoopCnt), intPsprice(intLoopCnt), _
                                           decAmount(intLoopCnt), strCurrency, strMadeCountry)
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当積上単価構成登録処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="strOfficeCd">営業所コード</param>
    ''' <param name="strRefKataban">形番</param>
    ''' <param name="decAmount">数量</param>
    ''' <param name="strPriceDiv">価格区分</param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intMode"></param>
    ''' <param name="DS_Tab"></param>
    ''' <remarks>引当積上単価構成テーブルに追加する</remarks>
    Public Sub subAccPriceStrcReg(objCon As SqlConnection, ByVal strUserId As String, _
                                  ByVal strSessionId As String, ByVal strCountryCd As String, _
                                  ByVal strOfficeCd As String, ByVal strRefKataban() As String, _
                                  ByVal decAmount() As Decimal, ByVal strPriceDiv() As String, _
                                  ByVal strCurrency As String, _
                                  Optional ByRef objKtbnStrc As KHKtbnStrc = Nothing, _
                                  Optional intMode As Integer = 0, Optional DS_Tab As DataSet = Nothing)
        Dim objPrice As New KHUnitPrice
        If objKtbnStrc Is Nothing Then objKtbnStrc = New KHKtbnStrc

        Dim strKatabanCheckDiv As String = Nothing
        Dim strPlaceCd As String = Nothing
        Dim htPriceInfo As Hashtable = Nothing

        Dim strOpKatabanCheckDiv() As String
        Dim strOpPlaceCd() As String
        Dim intOpListPrice() As Decimal
        Dim intOpRegPrice() As Decimal
        Dim intOpSsPrice() As Decimal
        Dim intOpBsPrice() As Decimal
        Dim intOpGsPrice() As Decimal
        Dim intOpPsPrice() As Decimal

        Dim bolReturn As Boolean
        Dim intLoopCnt As Integer
        'Dim strCurrency As String = String.Empty       'Add by Zxjike 2013/05/16
        Dim strMadeCountry As String = String.Empty       'Add by Zxjike 2013/06/07
        Try

            '配列定義
            ReDim strOpKatabanCheckDiv(0)
            ReDim strOpPlaceCd(0)
            ReDim intOpListPrice(0)
            ReDim intOpRegPrice(0)
            ReDim intOpSsPrice(0)
            ReDim intOpBsPrice(0)
            ReDim intOpGsPrice(0)
            ReDim intOpPsPrice(0)

            For intLoopCnt = 1 To strRefKataban.Length - 1
                '配列再定義
                ReDim Preserve strOpKatabanCheckDiv(UBound(strOpKatabanCheckDiv) + 1)
                ReDim Preserve strOpPlaceCd(UBound(strOpPlaceCd) + 1)
                ReDim Preserve intOpListPrice(UBound(intOpListPrice) + 1)
                ReDim Preserve intOpRegPrice(UBound(intOpRegPrice) + 1)
                ReDim Preserve intOpSsPrice(UBound(intOpSsPrice) + 1)
                ReDim Preserve intOpBsPrice(UBound(intOpBsPrice) + 1)
                ReDim Preserve intOpGsPrice(UBound(intOpGsPrice) + 1)
                ReDim Preserve intOpPsPrice(UBound(intOpPsPrice) + 1)

                If (Me.strcSelection.strPriceNo.Trim = "89" And intLoopCnt <= 7) Or _
                   (Me.strcSelection.strPriceNo.Trim = "96" And intLoopCnt <= 7) Or _
                   (Me.strcSelection.strPriceNo.Trim = "D3" And intLoopCnt <= 7) Then
                    'ISO価格取得(ベース・電磁弁)
                    bolReturn = Me.fncISOPriceGet(objCon, intLoopCnt, strRefKataban(intLoopCnt), _
                                                  strKatabanCheckDiv, strPlaceCd, _
                                                  htPriceInfo, strCurrency, strMadeCountry)
                Else
                    If DS_Tab Is Nothing Then
                        '積上単価情報読み込み
                        bolReturn = objPrice.fncSelectAccumulatePrice(objCon, strRefKataban(intLoopCnt), _
                                                                      strKatabanCheckDiv, strPlaceCd, htPriceInfo, strCurrency)

                        '単価情報読み込み
                        If Not bolReturn Then
                            bolReturn = objPrice.fncSelectPrice(objCon, strRefKataban(intLoopCnt), _
                                                                strKatabanCheckDiv, strPlaceCd, _
                                                                htPriceInfo, strCurrency, strMadeCountry)
                        End If
                    Else
                        Dim dr() As DataRow = Nothing
                        Dim dt_accPrice As New DS_KatOut.kh_accumulate_priceDataTable
                        dt_accPrice = DS_Tab.Tables("dt_accPrice")
                        dr = dt_accPrice.Select("kataban='" & strRefKataban(intLoopCnt) & "'")

                        htPriceInfo = New Hashtable
                        '初期化
                        strKatabanCheckDiv = ""
                        strPlaceCd = ""
                        htPriceInfo(CdCst.UnitPrice.ListPrice) = 0
                        htPriceInfo(CdCst.UnitPrice.RegPrice) = 0
                        htPriceInfo(CdCst.UnitPrice.SsPrice) = 0
                        htPriceInfo(CdCst.UnitPrice.BsPrice) = 0
                        htPriceInfo(CdCst.UnitPrice.GsPrice) = 0
                        htPriceInfo(CdCst.UnitPrice.PsPrice) = 0
                        strCurrency = String.Empty
                        strMadeCountry = String.Empty

                        If dr.Length > 0 Then
                            strKatabanCheckDiv = dr(0)("kataban_check_div")
                            strPlaceCd = dr(0)("place_cd")
                            htPriceInfo(CdCst.UnitPrice.ListPrice) = dr(0)("ls_price")
                            htPriceInfo(CdCst.UnitPrice.RegPrice) = dr(0)("rg_price")
                            htPriceInfo(CdCst.UnitPrice.SsPrice) = dr(0)("ss_price")
                            htPriceInfo(CdCst.UnitPrice.BsPrice) = dr(0)("bs_price")
                            htPriceInfo(CdCst.UnitPrice.GsPrice) = dr(0)("gs_price")
                            htPriceInfo(CdCst.UnitPrice.PsPrice) = dr(0)("ps_price")
                            bolReturn = True
                        Else
                            Dim dt_fullPrice As New DS_KatOut.kh_priceDataTable
                            dt_fullPrice = DS_Tab.Tables("dt_fullPrice")
                            dr = dt_fullPrice.Select("kataban='" & strRefKataban(intLoopCnt) & "'")
                            If dr.Length > 0 Then
                                strKatabanCheckDiv = dr(0)("kataban_check_div")
                                strPlaceCd = dr(0)("place_cd")
                                htPriceInfo(CdCst.UnitPrice.ListPrice) = dr(0)("ls_price")
                                htPriceInfo(CdCst.UnitPrice.RegPrice) = dr(0)("rg_price")
                                htPriceInfo(CdCst.UnitPrice.SsPrice) = dr(0)("ss_price")
                                htPriceInfo(CdCst.UnitPrice.BsPrice) = dr(0)("bs_price")
                                htPriceInfo(CdCst.UnitPrice.GsPrice) = dr(0)("gs_price")
                                htPriceInfo(CdCst.UnitPrice.PsPrice) = dr(0)("ps_price")
                                strCurrency = dr(0)("currency_cd")
                                strMadeCountry = dr(0)("country_cd")
                                bolReturn = True
                            End If
                        End If
                    End If
                End If


                If bolReturn Then
                    'チェック区分設定
                    strOpKatabanCheckDiv(UBound(strOpKatabanCheckDiv)) = strKatabanCheckDiv
                    '出荷場所設定
                    strOpPlaceCd(UBound(strOpPlaceCd)) = strPlaceCd

                    '価格設定
                    If strPriceDiv Is Nothing Then
                        '定価設定
                        intOpListPrice(UBound(intOpListPrice)) = htPriceInfo(CdCst.UnitPrice.ListPrice)
                        '登録店価格設定
                        intOpRegPrice(UBound(intOpRegPrice)) = htPriceInfo(CdCst.UnitPrice.RegPrice)
                        'Ss価格設定
                        intOpSsPrice(UBound(intOpSsPrice)) = htPriceInfo(CdCst.UnitPrice.SsPrice)
                        'Bs価格設定
                        intOpBsPrice(UBound(intOpBsPrice)) = htPriceInfo(CdCst.UnitPrice.BsPrice)
                        'Gs価格設定
                        intOpGsPrice(UBound(intOpGsPrice)) = htPriceInfo(CdCst.UnitPrice.GsPrice)
                        'Ps価格設定
                        intOpPsPrice(UBound(intOpPsPrice)) = htPriceInfo(CdCst.UnitPrice.PsPrice)
                    Else
                        '価格積上区分毎に設定
                        If strPriceDiv(intLoopCnt) Is Nothing Then
                            '定価設定
                            intOpListPrice(UBound(intOpListPrice)) = htPriceInfo(CdCst.UnitPrice.ListPrice)
                            '登録店価格設定
                            intOpRegPrice(UBound(intOpRegPrice)) = htPriceInfo(CdCst.UnitPrice.RegPrice)
                            'Ss価格設定
                            intOpSsPrice(UBound(intOpSsPrice)) = htPriceInfo(CdCst.UnitPrice.SsPrice)
                            'Bs価格設定
                            intOpBsPrice(UBound(intOpBsPrice)) = htPriceInfo(CdCst.UnitPrice.BsPrice)
                            'Gs価格設定
                            intOpGsPrice(UBound(intOpGsPrice)) = htPriceInfo(CdCst.UnitPrice.GsPrice)
                            'Ps価格設定
                            intOpPsPrice(UBound(intOpPsPrice)) = htPriceInfo(CdCst.UnitPrice.PsPrice)
                        Else
                            Select Case True
                                Case strPriceDiv(intLoopCnt).IndexOf(CdCst.PriceAccDiv.C5) >= 0
                                    'C5価格
                                    If strMadeCountry = "JPN" Or strMadeCountry = "" Then
                                        '定価設定
                                        intOpListPrice(UBound(intOpListPrice)) = Int(htPriceInfo(CdCst.UnitPrice.GsPrice) * CdCst.UnitPrice.C5Rate.ListPrice / 10 + 0.5) * 10
                                        '登録店価格設定
                                        intOpRegPrice(UBound(intOpRegPrice)) = Int(htPriceInfo(CdCst.UnitPrice.GsPrice) * CdCst.UnitPrice.C5Rate.RegPrice / 10 + 0.5) * 10
                                        'Ss価格設定
                                        intOpSsPrice(UBound(intOpSsPrice)) = Int(htPriceInfo(CdCst.UnitPrice.GsPrice) * CdCst.UnitPrice.C5Rate.SsPrice / 10 + 0.5) * 10
                                        'Bs価格設定
                                        intOpBsPrice(UBound(intOpBsPrice)) = Int(htPriceInfo(CdCst.UnitPrice.GsPrice) * CdCst.UnitPrice.C5Rate.BsPrice / 10 + 0.5) * 10
                                        'Gs価格設定
                                        intOpGsPrice(UBound(intOpGsPrice)) = Int(htPriceInfo(CdCst.UnitPrice.GsPrice) * CdCst.UnitPrice.C5Rate.GsPrice / 10 + 0.5) * 10
                                        'Ps価格設定
                                        intOpPsPrice(UBound(intOpPsPrice)) = Int(htPriceInfo(CdCst.UnitPrice.GsPrice) * CdCst.UnitPrice.C5Rate.PsPrice / 10 + 0.5) * 10
                                    Else
                                        '定価設定
                                        intOpListPrice(UBound(intOpListPrice)) = htPriceInfo(CdCst.UnitPrice.GsPrice) * CdCst.UnitPrice.C5Rate.ListPrice
                                        '登録店価格設定
                                        intOpRegPrice(UBound(intOpRegPrice)) = htPriceInfo(CdCst.UnitPrice.GsPrice) * CdCst.UnitPrice.C5Rate.RegPrice
                                        'Ss価格設定
                                        intOpSsPrice(UBound(intOpSsPrice)) = htPriceInfo(CdCst.UnitPrice.GsPrice) * CdCst.UnitPrice.C5Rate.SsPrice
                                        'Bs価格設定
                                        intOpBsPrice(UBound(intOpBsPrice)) = htPriceInfo(CdCst.UnitPrice.GsPrice) * CdCst.UnitPrice.C5Rate.BsPrice
                                        'Gs価格設定
                                        intOpGsPrice(UBound(intOpGsPrice)) = htPriceInfo(CdCst.UnitPrice.GsPrice) * CdCst.UnitPrice.C5Rate.GsPrice
                                        'Ps価格設定
                                        intOpPsPrice(UBound(intOpPsPrice)) = htPriceInfo(CdCst.UnitPrice.GsPrice) * CdCst.UnitPrice.C5Rate.PsPrice
                                    End If
                                Case strPriceDiv(intLoopCnt).IndexOf(CdCst.PriceAccDiv.DINRail) >= 0
                                    'DIN Rail価格
                                    If strMadeCountry = "JPN" Or strMadeCountry = "" Then
                                        '定価設定
                                        intOpListPrice(UBound(intOpListPrice)) = Int(htPriceInfo(CdCst.UnitPrice.ListPrice) / 1000 * decAmount(intLoopCnt) / 10 + 0.5) * 10
                                        '登録店価格設定
                                        intOpRegPrice(UBound(intOpRegPrice)) = Int(htPriceInfo(CdCst.UnitPrice.RegPrice) / 1000 * decAmount(intLoopCnt) / 10 + 0.5) * 10
                                        'Ss価格設定
                                        intOpSsPrice(UBound(intOpSsPrice)) = Int(htPriceInfo(CdCst.UnitPrice.SsPrice) / 1000 * decAmount(intLoopCnt) / 10 + 0.5) * 10
                                        'Bs価格設定
                                        intOpBsPrice(UBound(intOpBsPrice)) = Int(htPriceInfo(CdCst.UnitPrice.BsPrice) / 1000 * decAmount(intLoopCnt) / 10 + 0.5) * 10
                                        'Gs価格設定
                                        intOpGsPrice(UBound(intOpGsPrice)) = Int(htPriceInfo(CdCst.UnitPrice.GsPrice) / 1000 * decAmount(intLoopCnt) / 10 + 0.5) * 10
                                        'Ps価格設定
                                        intOpPsPrice(UBound(intOpPsPrice)) = Int(htPriceInfo(CdCst.UnitPrice.PsPrice) / 1000 * decAmount(intLoopCnt) / 10 + 0.5) * 10
                                    Else
                                        '定価設定
                                        intOpListPrice(UBound(intOpListPrice)) = htPriceInfo(CdCst.UnitPrice.ListPrice) / 1000 * decAmount(intLoopCnt)
                                        '登録店価格設定
                                        intOpRegPrice(UBound(intOpRegPrice)) = htPriceInfo(CdCst.UnitPrice.RegPrice) / 1000 * decAmount(intLoopCnt)
                                        'Ss価格設定
                                        intOpSsPrice(UBound(intOpSsPrice)) = htPriceInfo(CdCst.UnitPrice.SsPrice) / 1000 * decAmount(intLoopCnt)
                                        'Bs価格設定
                                        intOpBsPrice(UBound(intOpBsPrice)) = htPriceInfo(CdCst.UnitPrice.BsPrice) / 1000 * decAmount(intLoopCnt)
                                        'Gs価格設定
                                        intOpGsPrice(UBound(intOpGsPrice)) = htPriceInfo(CdCst.UnitPrice.GsPrice) / 1000 * decAmount(intLoopCnt)
                                        'Ps価格設定
                                        intOpPsPrice(UBound(intOpPsPrice)) = htPriceInfo(CdCst.UnitPrice.PsPrice) / 1000 * decAmount(intLoopCnt)
                                    End If
                                    'DIN Rail価格
                                    'DINレールの場合は数量を1に再設定
                                    decAmount(intLoopCnt) = 1
                                Case strPriceDiv(intLoopCnt).IndexOf(CdCst.PriceAccDiv.Joint) >= 0
                                    '継手価格
                                    If strMadeCountry = "JPN" Or strMadeCountry = "" Then
                                        '定価設定
                                        intOpListPrice(UBound(intOpListPrice)) = Int(htPriceInfo(CdCst.UnitPrice.ListPrice) / 10 * decAmount(intLoopCnt) / 10 + 0.5) * 10 / decAmount(intLoopCnt)
                                        '登録店価格設定
                                        intOpRegPrice(UBound(intOpRegPrice)) = Int(htPriceInfo(CdCst.UnitPrice.RegPrice) / 10 * decAmount(intLoopCnt) / 10 + 0.5) * 10 / decAmount(intLoopCnt)
                                        'Ss価格設定
                                        intOpSsPrice(UBound(intOpSsPrice)) = Int(htPriceInfo(CdCst.UnitPrice.SsPrice) / 10 * decAmount(intLoopCnt) / 10 + 0.5) * 10 / decAmount(intLoopCnt)
                                        'Bs価格設定
                                        intOpBsPrice(UBound(intOpBsPrice)) = Int(htPriceInfo(CdCst.UnitPrice.BsPrice) / 10 * decAmount(intLoopCnt) / 10 + 0.5) * 10 / decAmount(intLoopCnt)
                                        'Gs価格設定
                                        intOpGsPrice(UBound(intOpGsPrice)) = Int(htPriceInfo(CdCst.UnitPrice.GsPrice) / 10 * decAmount(intLoopCnt) / 10 + 0.5) * 10 / decAmount(intLoopCnt)
                                        'Ps価格設定
                                        intOpPsPrice(UBound(intOpPsPrice)) = Int(htPriceInfo(CdCst.UnitPrice.PsPrice) / 10 * decAmount(intLoopCnt) / 10 + 0.5) * 10 / decAmount(intLoopCnt)
                                    Else
                                        '定価設定
                                        intOpListPrice(UBound(intOpListPrice)) = htPriceInfo(CdCst.UnitPrice.ListPrice) / 10 * decAmount(intLoopCnt) / decAmount(intLoopCnt)
                                        '登録店価格設定
                                        intOpRegPrice(UBound(intOpRegPrice)) = htPriceInfo(CdCst.UnitPrice.RegPrice) / 10 * decAmount(intLoopCnt) / decAmount(intLoopCnt)
                                        'Ss価格設定
                                        intOpSsPrice(UBound(intOpSsPrice)) = htPriceInfo(CdCst.UnitPrice.SsPrice) / 10 * decAmount(intLoopCnt) / decAmount(intLoopCnt)
                                        'Bs価格設定
                                        intOpBsPrice(UBound(intOpBsPrice)) = htPriceInfo(CdCst.UnitPrice.BsPrice) / 10 * decAmount(intLoopCnt) / decAmount(intLoopCnt)
                                        'Gs価格設定
                                        intOpGsPrice(UBound(intOpGsPrice)) = htPriceInfo(CdCst.UnitPrice.GsPrice) / 10 * decAmount(intLoopCnt) / decAmount(intLoopCnt)
                                        'Ps価格設定
                                        intOpPsPrice(UBound(intOpPsPrice)) = htPriceInfo(CdCst.UnitPrice.PsPrice) / 10 * decAmount(intLoopCnt) / decAmount(intLoopCnt)
                                    End If
                                Case strPriceDiv(intLoopCnt).IndexOf(CdCst.PriceAccDiv.Open) >= 0
                                    '定価設定
                                    intOpListPrice(UBound(intOpListPrice)) = 0
                                    '登録店価格設定
                                    intOpRegPrice(UBound(intOpRegPrice)) = htPriceInfo(CdCst.UnitPrice.RegPrice)
                                    'Ss価格設定
                                    intOpSsPrice(UBound(intOpSsPrice)) = htPriceInfo(CdCst.UnitPrice.SsPrice)
                                    'Bs価格設定
                                    intOpBsPrice(UBound(intOpBsPrice)) = htPriceInfo(CdCst.UnitPrice.BsPrice)
                                    'Gs価格設定
                                    intOpGsPrice(UBound(intOpGsPrice)) = htPriceInfo(CdCst.UnitPrice.GsPrice)
                                    'Ps価格設定
                                    intOpPsPrice(UBound(intOpPsPrice)) = htPriceInfo(CdCst.UnitPrice.PsPrice)
                                Case Else
                                    '定価設定
                                    intOpListPrice(UBound(intOpListPrice)) = htPriceInfo(CdCst.UnitPrice.ListPrice)
                                    '登録店価格設定
                                    intOpRegPrice(UBound(intOpRegPrice)) = htPriceInfo(CdCst.UnitPrice.RegPrice)
                                    'Ss価格設定
                                    intOpSsPrice(UBound(intOpSsPrice)) = htPriceInfo(CdCst.UnitPrice.SsPrice)
                                    'Bs価格設定
                                    intOpBsPrice(UBound(intOpBsPrice)) = htPriceInfo(CdCst.UnitPrice.BsPrice)
                                    'Gs価格設定
                                    intOpGsPrice(UBound(intOpGsPrice)) = htPriceInfo(CdCst.UnitPrice.GsPrice)
                                    'Ps価格設定
                                    intOpPsPrice(UBound(intOpPsPrice)) = htPriceInfo(CdCst.UnitPrice.PsPrice)
                            End Select
                            '海外ユーザーはねじ加算無し
                            'ねじ価格
                            If (strCountryCd <> CdCst.CountryCd.DefaultCountry) Or _
                               (strCountryCd = CdCst.CountryCd.DefaultCountry And _
                                strOfficeCd = CdCst.OfficeCd.Overseas) Then
                                If strPriceDiv(intLoopCnt).IndexOf(CdCst.PriceAccDiv.Screw) >= 0 Then
                                    '定価設定
                                    intOpListPrice(UBound(intOpListPrice)) = 0
                                    '登録店価格設定
                                    intOpRegPrice(UBound(intOpRegPrice)) = 0
                                    'Ss価格設定
                                    intOpSsPrice(UBound(intOpSsPrice)) = 0
                                    'Bs価格設定
                                    intOpBsPrice(UBound(intOpBsPrice)) = 0
                                    'Gs価格設定
                                    intOpGsPrice(UBound(intOpGsPrice)) = 0
                                    'Ps価格設定
                                    intOpPsPrice(UBound(intOpPsPrice)) = 0
                                End If
                            End If
                        End If
                    End If
                Else
                    '価格情報が取得出来ない場合は0設定
                    strOpKatabanCheckDiv(UBound(strOpKatabanCheckDiv)) = strKatabanCheckDiv
                    strOpPlaceCd(UBound(strOpPlaceCd)) = strPlaceCd
                    intOpListPrice(UBound(intOpListPrice)) = 0
                    intOpRegPrice(UBound(intOpRegPrice)) = 0
                    intOpSsPrice(UBound(intOpSsPrice)) = 0
                    intOpBsPrice(UBound(intOpBsPrice)) = 0
                    intOpGsPrice(UBound(intOpGsPrice)) = 0
                    intOpPsPrice(UBound(intOpPsPrice)) = 0
                End If
            Next

            If intMode = 0 Then
                If strCurrency.Length <= 0 Or strMadeCountry.Length <= 0 Then
                    '引当情報取得
                    Call objKtbnStrc.subSelKtbnInfoGet(objCon, strUserId, strSessionId)
                    If strCurrency.Length <= 0 Then strCurrency = objKtbnStrc.strcSelection.strCurrency
                    If strMadeCountry.Length <= 0 Then strMadeCountry = objKtbnStrc.strcSelection.strMadeCountry 'Add by Zxjike 2013/06/07
                End If

                'RM1805036_二次電池価格加算
                Dim bolP40 As Boolean = False
                Select Case objKtbnStrc.strcSelection.strSeriesKataban
                    Case "MW3GA2", "MW4GA2", "MW4GB2", "MW4GZ2", "MW4GB4", "MW4GZ4"     '二次電池対応機種
                        Select Case Me.strcSelection.strKeyKataban
                            Case "P", "Y"
                                bolP40 = True
                        End Select
                    Case "NW3GA2", "NW4GA2", "NW4GB2", "NW4GZ2"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "P"
                                bolP40 = True
                        End Select
                    Case "W4GB2"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "G", "U"
                                bolP40 = True
                        End Select
                    Case "W4GB4", "W4GZ4"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "U", "V"
                                bolP40 = True
                        End Select
                End Select

                If bolP40 = True Then
                    For intLoopCnt = 1 To strRefKataban.Length - 1
                        '定価設定
                        intOpListPrice(intLoopCnt) = Int(intOpListPrice(intLoopCnt) * CdCst.UnitPrice.P40Rate.ListPrice / 10 + 0.5) * 10
                        '登録店価格設定
                        intOpRegPrice(intLoopCnt) = Int(intOpRegPrice(intLoopCnt) * CdCst.UnitPrice.P40Rate.RegPrice / 10 + 0.5) * 10
                        'Ss価格設定
                        intOpSsPrice(intLoopCnt) = Int(intOpSsPrice(intLoopCnt) * CdCst.UnitPrice.P40Rate.SsPrice / 10 + 0.5) * 10
                        'Bs価格設定
                        intOpBsPrice(intLoopCnt) = Int(intOpBsPrice(intLoopCnt) * CdCst.UnitPrice.P40Rate.BsPrice / 10 + 0.5) * 10
                        'Gs価格設定
                        intOpGsPrice(intLoopCnt) = Int(intOpGsPrice(intLoopCnt) * CdCst.UnitPrice.P40Rate.GsPrice / 10 + 0.5) * 10
                        'Ps価格設定
                        intOpPsPrice(intLoopCnt) = Int(intOpPsPrice(intLoopCnt) * CdCst.UnitPrice.P40Rate.PsPrice / 10 + 0.5) * 10
                    Next
                End If

                '引当積上単価構成追加
                Call objKtbnStrc.subInsertAccPriceStrc(objCon, strUserId, strSessionId, strRefKataban, _
                                                       strOpKatabanCheckDiv, strOpPlaceCd, intOpListPrice, _
                                                       intOpRegPrice, intOpSsPrice, intOpBsPrice, intOpGsPrice, _
                                                       intOpPsPrice, decAmount, strCurrency, strMadeCountry)
            Else
                'RM1805036_二次電池価格加算
                Dim bolP40 As Boolean = False
                Select Case objKtbnStrc.strcSelection.strSeriesKataban
                    Case "MW3GA2", "MW4GA2", "MW4GB2", "MW4GZ2", "MW4GB4", "MW4GZ4"     '二次電池対応機種
                        Select Case Me.strcSelection.strKeyKataban
                            Case "P", "Y"
                                bolP40 = True
                        End Select
                    Case "NW3GA2", "NW4GA2", "NW4GB2", "NW4GZ2"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "P"
                                bolP40 = True
                        End Select
                    Case "W4GB2"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "G", "U"
                                bolP40 = True
                        End Select
                    Case "W4GB4", "W4GZ4"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "U", "V"
                                bolP40 = True
                        End Select
                End Select

                If bolP40 = True Then
                    For intLoopCnt = 1 To intOpListPrice.Count - 1
                        '定価設定
                        intOpListPrice(intLoopCnt) = Int(intOpListPrice(intLoopCnt) * CdCst.UnitPrice.P40Rate.ListPrice / 10 + 0.5) * 10
                        '登録店価格設定
                        intOpRegPrice(intLoopCnt) = Int(intOpRegPrice(intLoopCnt) * CdCst.UnitPrice.P40Rate.RegPrice / 10 + 0.5) * 10
                        'Ss価格設定
                        intOpSsPrice(intLoopCnt) = Int(intOpSsPrice(intLoopCnt) * CdCst.UnitPrice.P40Rate.SsPrice / 10 + 0.5) * 10
                        'Bs価格設定
                        intOpBsPrice(intLoopCnt) = Int(intOpBsPrice(intLoopCnt) * CdCst.UnitPrice.P40Rate.BsPrice / 10 + 0.5) * 10
                        'Gs価格設定
                        intOpGsPrice(intLoopCnt) = Int(intOpGsPrice(intLoopCnt) * CdCst.UnitPrice.P40Rate.GsPrice / 10 + 0.5) * 10
                        'Ps価格設定
                        intOpPsPrice(intLoopCnt) = Int(intOpPsPrice(intLoopCnt) * CdCst.UnitPrice.P40Rate.PsPrice / 10 + 0.5) * 10
                    Next
                End If

                objKtbnStrc.strcSelection.strOpKatabanCheckDiv = strOpKatabanCheckDiv
                objKtbnStrc.strcSelection.strOpPlaceCd = strOpPlaceCd
                objKtbnStrc.strcSelection.intOpListPrice = intOpListPrice
                objKtbnStrc.strcSelection.intOpRegPrice = intOpRegPrice
                objKtbnStrc.strcSelection.intOpSsPrice = intOpSsPrice
                objKtbnStrc.strcSelection.intOpBsPrice = intOpBsPrice
                objKtbnStrc.strcSelection.intOpGsPrice = intOpGsPrice
                objKtbnStrc.strcSelection.intOpPsPrice = intOpPsPrice

                Dim strChk As String = String.Empty
                For inti As Integer = 1 To intOpListPrice.Count - 1
                    If strOpKatabanCheckDiv(inti).ToString.Trim.Length > 0 Then
                        If strChk.Trim = "" Or strChk.Trim < strOpKatabanCheckDiv(inti).ToString.Trim Then
                            strChk = strOpKatabanCheckDiv(inti).ToString.Trim
                        End If
                    End If
                    If strOpPlaceCd(inti).ToString.Trim.Length > 0 Then
                        'CHANGED BY YGY 20140617
                        'objKtbnStrc.strcSelection.strPlaceCd = strOpPlaceCd(inti)
                        If objKtbnStrc.strcSelection.strPlaceCd Is Nothing OrElse objKtbnStrc.strcSelection.strPlaceCd.Equals(String.Empty) Then
                            objKtbnStrc.strcSelection.strPlaceCd = strOpPlaceCd(inti)
                        End If
                    End If
                    'CHANGED BY YGY 20140620 decAmountが小数である場合もある
                    'objKtbnStrc.strcSelection.intListPrice += intOpListPrice(inti) * CInt(decAmount(inti))
                    'objKtbnStrc.strcSelection.intRegPrice += intOpRegPrice(inti) * CInt(decAmount(inti))
                    'objKtbnStrc.strcSelection.intSsPrice += intOpSsPrice(inti) * CInt(decAmount(inti))
                    'objKtbnStrc.strcSelection.intBsPrice += intOpBsPrice(inti) * CInt(decAmount(inti))
                    'objKtbnStrc.strcSelection.intGsPrice += intOpGsPrice(inti) * CInt(decAmount(inti))
                    'objKtbnStrc.strcSelection.intPsPrice += intOpPsPrice(inti) * CInt(decAmount(inti))
                    objKtbnStrc.strcSelection.intListPrice += Format(intOpListPrice(inti) * decAmount(inti) / 10, "####0") * 10
                    objKtbnStrc.strcSelection.intRegPrice += Format(intOpRegPrice(inti) * decAmount(inti) / 10, "####0") * 10
                    objKtbnStrc.strcSelection.intSsPrice += Format(intOpSsPrice(inti) * decAmount(inti) / 10, "####0") * 10
                    objKtbnStrc.strcSelection.intBsPrice += Format(intOpBsPrice(inti) * decAmount(inti) / 10, "####0") * 10
                    objKtbnStrc.strcSelection.intGsPrice += Format(intOpGsPrice(inti) * decAmount(inti) / 10, "####0") * 10
                    objKtbnStrc.strcSelection.intPsPrice += Format(intOpPsPrice(inti) * decAmount(inti) / 10, "####0") * 10
                Next

                objKtbnStrc.strcSelection.strKatabanCheckDiv = strChk
                Dim objOption As New KHOptionCtl
                strKatabanCheckDiv = objOption.fncKatabanCheckDivGet(objKtbnStrc, objKtbnStrc.strcSelection.strKatabanCheckDiv)
                objKtbnStrc.strcSelection.strKatabanCheckDiv = strKatabanCheckDiv
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objPrice = Nothing
        End Try

    End Sub

    ''' <summary>
    ''' 引当形番構成追加処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="intKtbnStrcSeqNo">形番構成順序</param>
    ''' <param name="strElementDiv">要素区分</param>
    ''' <param name="strStructureDiv">構成区分</param>
    ''' <param name="strAdditionDiv">付加区分</param>
    ''' <param name="strHyphenDiv">継続ハイフン有無区分</param>
    ''' <param name="strKtbnStrcNm">形番構成名称</param>
    ''' <param name="strplace_lvl"></param>
    ''' <remarks>引当形番構成テーブルにデータを追加する</remarks>
    Public Sub subSelKtbnStrcIns(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                 ByVal strSessionId As String, ByVal intKtbnStrcSeqNo As Integer, _
                                 ByVal strElementDiv As String, ByVal strStructureDiv As String, _
                                 ByVal strAdditionDiv As String, ByVal strHyphenDiv As String, _
                                 ByVal strKtbnStrcNm As String, ByVal strplace_lvl As Integer)
        Try
            Dim dalKtbnStrc As New KtbnStrcDAL

            dalKtbnStrc.subSelKtbnStrcIns(objCon, strUserId, strSessionId, intKtbnStrcSeqNo, _
                                          strElementDiv, strStructureDiv, strAdditionDiv, _
                                          strHyphenDiv, strKtbnStrcNm, strplace_lvl)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当シリーズ形番更新処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strDivision">区分</param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <remarks></remarks>
    Public Sub subSelSrsKtbnUpdate(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, _
                                   ByVal strDivision As String, ByVal strUserId As String, _
                                   ByVal strSessionId As String)
        Try
            Dim dt As New DataTable
            Dim objOption As New KHOptionCtl
            Dim dalKtbnStrc As New KtbnStrcDAL
            Dim result As New KHAccPriceModel

            With result
                dt = dalKtbnStrc.fncSelectAccPriceStrc(objCon, strUserId, strSessionId)

                For Each dr As DataRow In dt.Rows
                    result.strCheckDiv = dr("kataban_check_div")
                    result.strCurrency = dr("currency_cd")
                    result.strMadeCountry = dr("country_cd")
                    If .strKatabanCheckDiv.Trim = "" Or .strKatabanCheckDiv.Trim < .strCheckDiv Then
                        .strKatabanCheckDiv = .strCheckDiv
                    End If

                    If .strPlaceCd.Trim = "" Then .strPlaceCd = dr("place_cd")

                    If strDivision = "0" Then
                        .intListPrice = .intListPrice + CDec(dr("ls_price"))
                        .intRegPrice = .intRegPrice + CDec(dr("rg_price"))
                        .intSsPrice = .intSsPrice + CDec(dr("ss_price"))
                        .intBsPrice = .intBsPrice + CDec(dr("bs_price"))
                        .intGsPrice = .intGsPrice + CDec(dr("gs_price"))
                        .intPsPrice = .intPsPrice + CDec(dr("ps_price"))
                    Else
                        If dr("country_cd").ToString.Equals("JPN") Then
                            .intListPrice += Format(dr("ls_price") * dr("amount") / 10, "####0") * 10
                            .intRegPrice += Format(dr("rg_price") * dr("amount") / 10, "####0") * 10
                            .intSsPrice += Format(dr("ss_price") * dr("amount") / 10, "####0") * 10
                            .intBsPrice += Format(dr("bs_price") * dr("amount") / 10, "####0") * 10
                            .intGsPrice += Format(dr("gs_price") * dr("amount") / 10, "####0") * 10
                            .intPsPrice += Format(dr("ps_price") * dr("amount") / 10, "####0") * 10
                        Else
                            .intListPrice += dr("ls_price") * dr("amount")
                            .intRegPrice += dr("rg_price") * dr("amount")
                            .intSsPrice += dr("ss_price") * dr("amount")
                            .intBsPrice += dr("bs_price") * dr("amount")
                            .intGsPrice += dr("gs_price") * dr("amount")
                            .intPsPrice += dr("ps_price") * dr("amount")
                        End If

                    End If
                Next

                .strKatabanCheckDiv = objOption.fncKatabanCheckDivGet(objKtbnStrc, .strKatabanCheckDiv)

                '原価積算No取得
                .strCostCalcNo = objOption.fncCostCalcNoGet(objKtbnStrc, .strKatabanCheckDiv)

                '引当シリーズ形番更新処理(価格)
                Call dalKtbnStrc.subSelSrsKtbnPriceUpd(objCon, strUserId, strSessionId, .strKatabanCheckDiv, .strPlaceCd, _
                                              .strCostCalcNo, .intListPrice, .intRegPrice, .intSsPrice, _
                                              .intBsPrice, .intGsPrice, .intPsPrice, .strCurrency, .strMadeCountry)
            End With
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Sub

    ''' <summary>
    ''' 引当てたオプションよりフル形番を生成する
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strGAMD0"></param>
    ''' <remarks>RM0902085 T.Yagyu NAB (4)シリンダカバーボディシール材質組合せ=無記号（標準）で(5)オプションで取付板の場合は(4)に"0"をセットする</remarks>
    Public Sub subFullKatabanCreate(ByVal objCon As SqlConnection, ByVal strUserId As String, _
                                    ByVal strSessionId As String, _
                                    Optional ByVal strGAMD0 As String = "")
        Dim objKat As New KHKataban
        Dim sbKataban As New StringBuilder
        Dim strKataban As String
        Dim intLoopCnt As Integer
        Dim strRodChangeFF As String
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            '引当情報取得
            Call Me.subSelKtbnInfoGet(objCon, strUserId, strSessionId)

            '形番を生成する
            'シリーズ形番
            sbKataban.Append(Me.strcSelection.strSeriesKataban)
            If Me.strcSelection.strHyphen = CdCst.HyphenDiv.Necessary Then
                sbKataban.Append(CdCst.Sign.Hypen)
            End If

            If Left(Me.strcSelection.strSeriesKataban.Trim, 3) = "NAB" Then
                If Me.strcSelection.strOpSymbol(4).Trim = "" And Me.strcSelection.strOpSymbol(5) = "B" Then
                    Me.strcSelection.strOpSymbol(4) = "0"
                End If
            End If

            '選択したオプションを結合
            If (Left(Me.strcSelection.strSeriesKataban.Trim, 4) = "AMD3" And (Me.strcSelection.strKeyKataban.Trim = "1" Or Me.strcSelection.strKeyKataban.Trim = "2")) Or _
            (Left(Me.strcSelection.strSeriesKataban.Trim, 4) = "AMD4" And Len(Me.strcSelection.strKeyKataban.Trim) = 0) Or _
            (Left(Me.strcSelection.strSeriesKataban.Trim, 4) = "AMD5" And Len(Me.strcSelection.strKeyKataban.Trim) = 0) Or _
            (Left(Me.strcSelection.strSeriesKataban.Trim, 4) = "AMD0" And Me.strcSelection.strKeyKataban.Trim = "1") Then
                If Len(Me.strcSelection.strOpSymbol(8)) <> 0 Then
                    For intLoopCnt = 1 To UBound(Me.strcSelection.strOpSymbol)
                        sbKataban.Append(Me.strcSelection.strOpSymbol(intLoopCnt).Replace(CdCst.Sign.Delimiter.Comma, "").Trim)

                        If Me.strcSelection.strOpHyphenDiv(intLoopCnt) = CdCst.HyphenDiv.Necessary And _
                        intLoopCnt <> 5 Then
                            sbKataban.Append(CdCst.Sign.Hypen)
                        End If
                    Next
                Else
                    For intLoopCnt = 1 To UBound(Me.strcSelection.strOpSymbol)
                        sbKataban.Append(Me.strcSelection.strOpSymbol(intLoopCnt).Replace(CdCst.Sign.Delimiter.Comma, "").Trim)

                        If Me.strcSelection.strOpHyphenDiv(intLoopCnt) = CdCst.HyphenDiv.Necessary Then
                            sbKataban.Append(CdCst.Sign.Hypen)
                        End If
                    Next
                End If
            Else
                For intLoopCnt = 1 To UBound(Me.strcSelection.strOpSymbol)
                    sbKataban.Append(Me.strcSelection.strOpSymbol(intLoopCnt).Replace(CdCst.Sign.Delimiter.Comma, "").Trim)

                    If Me.strcSelection.strOpHyphenDiv(intLoopCnt) = CdCst.HyphenDiv.Necessary Then
                        sbKataban.Append(CdCst.Sign.Hypen)
                    End If
                Next
            End If

            'ロッド先端仕様を結合
            If Me.strcSelection.strRodEndOption.Trim <> "" Then
                If dalKtbnStrc.fncSelRodSelect(objCon, strUserId, strSessionId) = False Then
                    If sbKataban.ToString.EndsWith(CdCst.Sign.Hypen) Then
                        sbKataban = New StringBuilder(Left(sbKataban.ToString, sbKataban.ToString.Length - 1))
                    End If
                    sbKataban.Append(Me.strcSelection.strRodEndOption.Trim)
                Else
                    If Left(Me.strcSelection.strSeriesKataban.Trim, 4) = "JSC3" Then
                        If Me.strcSelection.strKeyKataban = "1" Then
                            If Me.strcSelection.strOpSymbol(4).Trim = "FA" Then
                                strRodChangeFF = Me.strcSelection.strRodEndOption.Trim.Replace("WF", "FF")
                                sbKataban.Append(CdCst.Sign.Hypen & strRodChangeFF)
                            Else
                                sbKataban.Append(CdCst.Sign.Hypen & Me.strcSelection.strRodEndOption.Trim)
                            End If
                        Else
                            sbKataban.Append(CdCst.Sign.Hypen & Me.strcSelection.strRodEndOption.Trim)
                        End If
                    Else
                        sbKataban.Append(CdCst.Sign.Hypen & Me.strcSelection.strRodEndOption.Trim)
                    End If
                End If
            End If

            'オプション外仕様を結合
            If Me.strcSelection.strOtherOption.Trim <> "" Then
                sbKataban.Append(CdCst.Sign.OtherOpSymbol & Me.strcSelection.strOtherOption.Trim)
            End If

            '形番から重複するハイフンを除去する
            strKataban = KHKataban.fncHypenCut(sbKataban.ToString)

            If Me.strcSelection.strSeriesKataban.Trim = "GAMD0" Then
                If Me.strcSelection.strOpSymbol(1).Trim = "X" Then strKataban &= "-X" & strGAMD0
            End If

            '引当シリーズ形番更新処理(フル形番)
            Call dalKtbnStrc.subSelSrsKtbnFullKtbnUpd(objCon, strUserId, strSessionId, strKataban)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            sbKataban = Nothing
            objKat = Nothing
        End Try

    End Sub
    ''' <summary>
    ''' ISOバルブ価格キーの取得
    ''' </summary>
    ''' <param name="intIndex">順序</param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strRefPriceKey"></param>
    ''' <param name="decRefAmount"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncISOGetPriceKey(ByVal intIndex As Integer, ByVal strKataban As String, _
                                      ByRef strRefPriceKey() As String, ByRef decRefAmount() As Decimal) As Boolean
        Dim objPrice As New KHUnitPrice
        fncISOGetPriceKey = False
        Try
            '価格キー取得
            If Me.strcSelection.strPriceNo.Trim = "89" Then
                If intIndex = 1 Then
                    Call Me.subLMFBasePriceKeyGet(strKataban, strRefPriceKey, decRefAmount)
                Else
                    Call Me.subLMFValvePriceKeyGet(strKataban, strRefPriceKey, decRefAmount)
                End If
            Else
                If intIndex = 1 Then
                    If Left(strKataban.Trim, 3) = "CMF" Then
                        Call Me.subCMFBasePriceKeyGet(strKataban, strRefPriceKey, decRefAmount)
                    Else
                        Call Me.subGMFBasePriceKeyGet(strKataban, strRefPriceKey, decRefAmount)
                    End If
                Else
                    If Left(strKataban.Trim, 3) = "CMF" Then
                        Call Me.subCMFValvePriceKeyGet(strKataban, strRefPriceKey, decRefAmount)
                    Else
                        Call Me.subGMFValvePriceKeyGet(strKataban, strRefPriceKey, decRefAmount)
                    End If
                End If
            End If
            fncISOGetPriceKey = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objPrice = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 引当シリーズ形番を読み込み単価情報を取得し返却する
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks></remarks>
    Private Sub subSelSrsKtbnSelect(objCon As SqlConnection)
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim objRdr As SqlDataReader = Nothing
        Dim dt As New DataTable
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            dt = dalKtbnStrc.fncSelSrsKtbn(objCon, Me.strcSelection.strUserId, Me.strcSelection.strSessionId)

            If dt.Rows.Count > 0 Then
                With Me.strcSelection
                    Dim dr As DataRow = dt.Rows(0)

                    .strDivision = dr("division")
                    If IsDBNull(dr("series_kataban")) Then
                        .strSeriesKataban = ""
                    Else
                        .strSeriesKataban = dr("series_kataban")
                    End If
                    If IsDBNull(dr("key_kataban")) Then
                        .strKeyKataban = ""
                    Else
                        .strKeyKataban = dr("key_kataban")
                    End If
                    If IsDBNull(dr("hyphen_div")) Then
                        .strHyphen = ""
                    Else
                        .strHyphen = dr("hyphen_div")
                    End If
                    If IsDBNull(dr("price_no")) Then
                        .strPriceNo = ""
                    Else
                        .strPriceNo = dr("price_no")
                    End If
                    If IsDBNull(dr("spec_no")) Then
                        .strSpecNo = ""
                    Else
                        .strSpecNo = dr("spec_no")
                    End If
                    If IsDBNull(dr("full_kataban")) Then
                        .strFullKataban = ""
                    Else
                        .strFullKataban = dr("full_kataban")
                    End If
                    If IsDBNull(dr("goods_nm")) Then
                        .strGoodsNm = ""
                    Else
                        .strGoodsNm = dr("goods_nm")
                    End If
                    If IsDBNull(dr("kataban_check_div")) Then
                        .strKatabanCheckDiv = ""
                    Else
                        .strKatabanCheckDiv = dr("kataban_check_div")
                    End If
                    If IsDBNull(dr("place_cd")) Then
                        .strPlaceCd = ""
                    Else
                        .strPlaceCd = dr("place_cd")
                    End If
                    If IsDBNull(dr("cost_calc_no")) Then
                        .strCostCalcNo = ""
                    Else
                        .strCostCalcNo = dr("cost_calc_no")
                    End If
                    If IsDBNull(dr("ls_price")) Then
                        .intListPrice = 0
                    Else
                        .intListPrice = dr("ls_price")
                    End If
                    If IsDBNull(dr("rg_price")) Then
                        .intRegPrice = 0
                    Else
                        .intRegPrice = dr("rg_price")
                    End If
                    If IsDBNull(dr("ss_price")) Then
                        .intSsPrice = 0
                    Else
                        .intSsPrice = dr("ss_price")
                    End If
                    If IsDBNull(dr("bs_price")) Then
                        .intBsPrice = 0
                    Else
                        .intBsPrice = dr("bs_price")
                    End If
                    If IsDBNull(dr("gs_price")) Then
                        .intGsPrice = 0
                    Else
                        .intGsPrice = dr("gs_price")
                    End If
                    If IsDBNull(dr("ps_price")) Then
                        .intPsPrice = 0
                    Else
                        .intPsPrice = dr("ps_price")
                    End If
                    If IsDBNull(dr("factor")) Then
                        .decFactor = 0
                    Else
                        .decFactor = dr("factor")
                    End If
                    If IsDBNull(dr("unit_price")) Then
                        .intUnitPrice = 0
                    Else
                        .intUnitPrice = dr("unit_price")
                    End If
                    If IsDBNull(dr("amount")) Then
                        .intAmount = 0
                    Else
                        .intAmount = dr("amount")
                    End If
                    If IsDBNull(dr("currency_cd")) Then
                        .strCurrency = ""
                    Else
                        .strCurrency = dr("currency_cd")
                    End If
                    If IsDBNull(dr("country_cd")) Then
                        .strMadeCountry = ""
                    Else
                        .strMadeCountry = dr("country_cd")
                    End If
                    If IsDBNull(dr("rod_end_option")) Then
                        .strRodEndOption = ""
                    Else
                        .strRodEndOption = dr("rod_end_option")
                    End If
                    If IsDBNull(dr("other_option")) Then
                        .strOtherOption = ""
                    Else
                        .strOtherOption = dr("other_option")
                    End If
                    If IsDBNull(dr("position_option")) Then
                        .strPositionOption = ""
                    Else
                        .strPositionOption = dr("position_option")
                    End If
                End With
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
            objCmd = Nothing
        End Try

    End Sub

    ''' <summary>
    ''' ISO価格取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="intIndex">順序</param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strKatabanCheckDiv">形番チェック区分</param>
    ''' <param name="strPlaceCd">出荷場所</param>
    ''' <param name="htPriceInfo">価格情報</param>
    ''' <param name="strCurrency"></param>
    ''' <param name="strMadeCountry"></param>
    ''' <returns></returns>
    ''' <remarks>ISOバルブのベース及び電磁便の価格を算出する</remarks>
    Private Function fncISOPriceGet(objCon As SqlConnection, ByVal intIndex As Integer, _
                                    ByVal strKataban As String, _
                                    ByRef strKatabanCheckDiv As String, _
                                    ByRef strPlaceCd As String, _
                                    ByRef htPriceInfo As Hashtable, _
                                    ByRef strCurrency As String, ByRef strMadeCountry As String) As Boolean
        Dim objPrice As New KHUnitPrice
        Dim strRefKataban() As String = Nothing
        Dim decRefAmount() As Decimal = Nothing
        Dim strRetKatabanCheckDiv As String = Nothing
        Dim strRetPlaceCd As String = Nothing
        Dim htRetPriceInfo As Hashtable = Nothing
        Dim intLoopCnt As Integer
        Dim bolReturn As Boolean
        fncISOPriceGet = False
        Try
            htPriceInfo = New Hashtable

            '初期化
            strKatabanCheckDiv = Nothing
            strPlaceCd = Nothing
            htPriceInfo(CdCst.UnitPrice.ListPrice) = 0
            htPriceInfo(CdCst.UnitPrice.RegPrice) = 0
            htPriceInfo(CdCst.UnitPrice.SsPrice) = 0
            htPriceInfo(CdCst.UnitPrice.BsPrice) = 0
            htPriceInfo(CdCst.UnitPrice.GsPrice) = 0
            htPriceInfo(CdCst.UnitPrice.PsPrice) = 0

            '価格キー取得
            If Me.strcSelection.strPriceNo.Trim = "89" Then
                If intIndex = 1 Then
                    Call Me.subLMFBasePriceKeyGet(strKataban, strRefKataban, decRefAmount)
                Else
                    Call Me.subLMFValvePriceKeyGet(strKataban, strRefKataban, decRefAmount)
                End If
            Else
                If intIndex = 1 Then
                    If Left(strKataban.Trim, 3) = "CMF" Then
                        Call Me.subCMFBasePriceKeyGet(strKataban, strRefKataban, decRefAmount)
                    Else
                        Call Me.subGMFBasePriceKeyGet(strKataban, strRefKataban, decRefAmount)
                    End If
                Else
                    If Left(strKataban.Trim, 3) = "CMF" Then
                        Call Me.subCMFValvePriceKeyGet(strKataban, strRefKataban, decRefAmount)
                    Else
                        Call Me.subGMFValvePriceKeyGet(strKataban, strRefKataban, decRefAmount)
                    End If
                End If
            End If


            For intLoopCnt = 1 To strRefKataban.Length - 1
                ''単価情報読み込み
                'bolReturn = objPrice.fncSelectPrice(objCon, strRefKataban(intLoopCnt), _
                '                                    strRetKatabanCheckDiv, strRetPlaceCd, _
                '                                    htRetPriceInfo, strCurrency, strMadeCountry)
                ''積上単価情報読み込み
                'If Not bolReturn Then
                '    bolReturn = objPrice.fncSelectAccumulatePrice(objCon, strRefKataban(intLoopCnt), _
                '                                                  strRetKatabanCheckDiv, strRetPlaceCd, htRetPriceInfo, _
                '                                                  strCurrency)
                'End If

                '積上単価情報読み込み
                bolReturn = objPrice.fncSelectAccumulatePrice(objCon, strRefKataban(intLoopCnt), _
                                                                   strRetKatabanCheckDiv, strRetPlaceCd, htRetPriceInfo, _
                                                                   strCurrency)
                If Not bolReturn Then
                    '単価情報読み込み
                    bolReturn = objPrice.fncSelectPrice(objCon, strRefKataban(intLoopCnt), _
                                                 strRetKatabanCheckDiv, strRetPlaceCd, _
                                                 htRetPriceInfo, strCurrency, strMadeCountry)
                End If

                '価格が取得出来た場合
                If bolReturn Then
                    If strKatabanCheckDiv Is Nothing Then
                        strKatabanCheckDiv = strRetKatabanCheckDiv
                    Else
                        'ADD BY YGY 20141020
                        If fncCompareStrInteger(strKatabanCheckDiv, strRetKatabanCheckDiv) Then
                            strKatabanCheckDiv = strRetKatabanCheckDiv
                        End If
                    End If
                    If strPlaceCd Is Nothing Then
                        strPlaceCd = strRetPlaceCd
                    End If
                    htPriceInfo(CdCst.UnitPrice.ListPrice) = htPriceInfo(CdCst.UnitPrice.ListPrice) + htRetPriceInfo(CdCst.UnitPrice.ListPrice) * decRefAmount(intLoopCnt)
                    htPriceInfo(CdCst.UnitPrice.RegPrice) = htPriceInfo(CdCst.UnitPrice.RegPrice) + htRetPriceInfo(CdCst.UnitPrice.RegPrice) * decRefAmount(intLoopCnt)
                    htPriceInfo(CdCst.UnitPrice.SsPrice) = htPriceInfo(CdCst.UnitPrice.SsPrice) + htRetPriceInfo(CdCst.UnitPrice.SsPrice) * decRefAmount(intLoopCnt)
                    htPriceInfo(CdCst.UnitPrice.BsPrice) = htPriceInfo(CdCst.UnitPrice.BsPrice) + htRetPriceInfo(CdCst.UnitPrice.BsPrice) * decRefAmount(intLoopCnt)
                    htPriceInfo(CdCst.UnitPrice.GsPrice) = htPriceInfo(CdCst.UnitPrice.GsPrice) + htRetPriceInfo(CdCst.UnitPrice.GsPrice) * decRefAmount(intLoopCnt)
                    htPriceInfo(CdCst.UnitPrice.PsPrice) = htPriceInfo(CdCst.UnitPrice.PsPrice) + htRetPriceInfo(CdCst.UnitPrice.PsPrice) * decRefAmount(intLoopCnt)
                End If
                fncISOPriceGet = True
            Next

        Catch ex As Exception
            fncISOPriceGet = False
        Finally
            objPrice = Nothing
        End Try
    End Function

    ''' <summary>
    ''' LMF価格キー設定
    ''' </summary>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strRefKataban">価格キー</param>
    ''' <param name="decRefAmount">個数</param>
    ''' <remarks>LMFのベースの価格キーを設定する</remarks>
    Private Sub subLMFBasePriceKeyGet(ByVal strKataban As String, _
                                      ByRef strRefKataban() As String, _
                                      ByRef decRefAmount() As Decimal)
        Try
            ReDim strRefKataban(0)
            ReDim decRefAmount(0)
            'ベース
            If Me.strcSelection.strKeyKataban.Trim = "1" Then
                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                strRefKataban(UBound(strRefKataban)) = "LMF0-1-BASE-" & Me.strcSelection.strOpSymbol(1).Trim
                decRefAmount(UBound(decRefAmount)) = 1
            Else
                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                strRefKataban(UBound(strRefKataban)) = "LMF0-2-BASE-" & Me.strcSelection.strOpSymbol(1).Trim
                decRefAmount(UBound(decRefAmount)) = 1
            End If

            'A・Bポート接続口径
            If Me.strcSelection.intQuantity.Length > 10 AndAlso Me.strcSelection.intQuantity(10) > 0 Then
                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                strRefKataban(UBound(strRefKataban)) = "LMF0-PORT-C4"
                decRefAmount(UBound(decRefAmount)) = Me.strcSelection.intQuantity(10)
            End If
            If Me.strcSelection.intQuantity.Length > 11 AndAlso Me.strcSelection.intQuantity(11) > 0 Then
                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                strRefKataban(UBound(strRefKataban)) = "LMF0-PORT-C6"
                decRefAmount(UBound(decRefAmount)) = Me.strcSelection.intQuantity(11)
            End If
            If Me.strcSelection.intQuantity.Length > 12 AndAlso Me.strcSelection.intQuantity(12) > 0 Then
                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                strRefKataban(UBound(strRefKataban)) = "LMF0-PORT-01Z"
                decRefAmount(UBound(decRefAmount)) = Me.strcSelection.intQuantity(12)
            End If

            'P・R1・R2ポート接続口径
            Select Case Me.strcSelection.strOpSymbol(3).Trim
                Case "C8B", "C8D", "C8U"
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = "LMF0-PORT-" & Me.strcSelection.strOpSymbol(3).Trim
                    decRefAmount(UBound(decRefAmount)) = 1
            End Select

            '電気接続
            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
            strRefKataban(UBound(strRefKataban)) = "LMF0-OP-" & Me.strcSelection.strOpSymbol(4).Trim
            decRefAmount(UBound(decRefAmount)) = 1
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' LMF価格キー設定
    ''' </summary>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strRefKataban">価格キー</param>
    ''' <param name="decRefAmount">個数</param>
    ''' <remarks>LMFの電磁弁の価格キーを設定する</remarks>
    Private Sub subLMFValvePriceKeyGet(ByVal strKataban As String, ByRef strRefKataban() As String, _
                                       ByRef decRefAmount() As Decimal)
        Dim intLoopCnt As Integer
        Dim intIdx As Integer = 0
        Try
            ReDim strRefKataban(0)
            ReDim decRefAmount(0)

            If Left(strKataban, 3) = "LMF" Then
                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                strRefKataban(UBound(strRefKataban)) = strKataban.Trim
                decRefAmount(UBound(decRefAmount)) = 1
            Else
                '形番の3番目のハイフン+1桁までを切り出す
                For intLoopCnt = 1 To strKataban.Length - 1
                    If Mid(strKataban.Trim, intLoopCnt, 1) = CdCst.Sign.Hypen Then
                        intIdx = intIdx + 1
                        If intIdx = 3 Then
                            'ランプ付の場合
                            If Me.strcSelection.strOpSymbol(6).Trim = "N" Then
                                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                                strRefKataban(UBound(strRefKataban)) = Mid(strKataban.Trim, 1, intLoopCnt + 1) & "*N"
                                decRefAmount(UBound(decRefAmount)) = 1
                            Else
                                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                                strRefKataban(UBound(strRefKataban)) = Mid(strKataban.Trim, 1, intLoopCnt + 1)
                                decRefAmount(UBound(decRefAmount)) = 1
                            End If
                            Exit For
                        End If
                    End If
                Next

                '手動装置加算価格キー
                Select Case Me.strcSelection.strOpSymbol(7).Trim
                    Case "M1", "M4"
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = "4L2-4-" & Me.strcSelection.strOpSymbol(7).Trim
                        If InStr(1, strKataban, "FG-S") = 0 Then
                            decRefAmount(UBound(decRefAmount)) = 2
                        Else
                            decRefAmount(UBound(decRefAmount)) = 1
                        End If
                End Select

                '電圧加算価格キー
                If Me.strcSelection.strOpSymbol(5).Trim = "9" Then
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = "4L2-4-OPT"
                    If InStr(1, strKataban, "FG-S") = 0 Then
                        decRefAmount(UBound(decRefAmount)) = 2
                    Else
                        decRefAmount(UBound(decRefAmount)) = 1
                    End If
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' CMF価格キー設定
    ''' </summary>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strRefKataban">価格キー</param>
    ''' <param name="decRefAmount">個数</param>
    ''' <remarks>CMFのベースの価格キーを設定する</remarks>
    Private Sub subCMFBasePriceKeyGet(ByVal strKataban As String, _
                                      ByRef strRefKataban() As String, _
                                      ByRef decRefAmount() As Decimal)
        Dim intLoopCnt As Integer

        Try
            ReDim strRefKataban(0)
            ReDim decRefAmount(0)

            If Me.strcSelection.strPriceNo.Trim = "96" Then
                '基本価格キー
                If Left(strKataban.Trim, 4) = "CMFZ" Then
                    'マニホールドブロック
                    Select Case Me.strcSelection.strOpSymbol(3).Trim
                        Case "HX3"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-BLOCK-02"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-BLOCK-03"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX4"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-BLOCK-02"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-BLOCK-04"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX5"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-BLOCK-03"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-BLOCK-03"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX6"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-BLOCK-03"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-BLOCK-04"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                    End Select
                    'フート
                    Select Case Me.strcSelection.strOpSymbol(5).Trim
                        Case "HY3"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-FOOT-03"
                            decRefAmount(UBound(decRefAmount)) = 1

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-FOOT-04"
                            decRefAmount(UBound(decRefAmount)) = 1
                        Case "HY4"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-FOOT-03"
                            decRefAmount(UBound(decRefAmount)) = 1

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-FOOT-06"
                            decRefAmount(UBound(decRefAmount)) = 1
                        Case "HY5"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-FOOT-04"
                            decRefAmount(UBound(decRefAmount)) = 1

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-FOOT-04"
                            decRefAmount(UBound(decRefAmount)) = 1
                        Case "HY6"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-FOOT-04"
                            decRefAmount(UBound(decRefAmount)) = 1

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-FOOT-06"
                            decRefAmount(UBound(decRefAmount)) = 1
                    End Select

                    'ミックスブロック
                    For intLoopCnt = 1 To Me.strcSelection.strOptionKataban.Length - 1
                        Select Case Me.strcSelection.strOptionKataban(intLoopCnt).Trim
                            Case "CMFBZ-00L", "CMFBZ-00R"
                                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                                strRefKataban(UBound(strRefKataban)) = Me.strcSelection.strOptionKataban(intLoopCnt).Trim
                                decRefAmount(UBound(decRefAmount)) = Me.strcSelection.intQuantity(intLoopCnt)
                        End Select
                    Next
                Else
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = Mid(strKataban.Trim, 1, InStr(1, strKataban.Trim, "-")) & "MANIHOLD-BASE"
                    decRefAmount(UBound(decRefAmount)) = 1
                End If

                'A・Bポート口径価格キー
                '裏配管の場合
                If Me.strcSelection.strOpSymbol(4).Trim = "Z" Then
                    Select Case Me.strcSelection.strOpSymbol(3).Trim
                        Case "HX1"
                            'CMF1
                            '02加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-02Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX2"
                            'CMF2
                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '04加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-04Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX3"
                            'CMFZ
                            '02加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF1-BASE-PORT-02Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF2-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX4"
                            'CMFZ
                            '02加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF1-BASE-PORT-02Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '04加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF2-BASE-PORT-04Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX5"
                            'CMFZ
                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF1-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF2-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX6"
                            'CMFZ
                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF1-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '04加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF2-BASE-PORT-04Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case Else
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-" & _
                                                                   Me.strcSelection.strOpSymbol(3).Trim & _
                                                                   Me.strcSelection.strOpSymbol(4).Trim
                            decRefAmount(UBound(decRefAmount)) = CDec(Me.strcSelection.strOpSymbol(2).Trim)
                    End Select
                End If

                'P・Rポート口径価格キー
                'P・Rポートの加算は"06"または"HY2"の場合のみで、"06"・"HY2"はCMF2の場合にしかありえない
                If Left(strKataban.Trim, 4) = "CMF2" Then
                    Select Case Me.strcSelection.strOpSymbol(5).Trim
                        Case "04"
                        Case Else
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-" & _
                                                                   Me.strcSelection.strOpSymbol(5).Trim
                            decRefAmount(UBound(decRefAmount)) = 1
                    End Select
                End If

                '制御ユニット価格キー
                '制御ユニット付のものは加算
                Select Case Me.strcSelection.strKeyKataban.Trim
                    Case "4", "5", "9"
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = "CMF1-UNIT-" & Me.strcSelection.strOpSymbol(5).Trim
                        decRefAmount(UBound(decRefAmount)) = 1
                    Case "6", "7"
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = "CMF1-BASE-UNIT-" & Me.strcSelection.strOpSymbol(5).Trim
                        decRefAmount(UBound(decRefAmount)) = 1
                End Select

                'サイレンサボックス価格キー
                '制御ユニット付でないものは、サイレンサボックスが選択可能
                Select Case Me.strcSelection.strKeyKataban.Trim
                    Case "1", "2", "3", "8"
                        If Me.strcSelection.strOpSymbol(8).Trim <> "" Then
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF-BASE-" & Me.strcSelection.strOpSymbol(8).Trim
                            decRefAmount(UBound(decRefAmount)) = 1
                        End If
                End Select

                'CMF1*「ISOバルブ(制御ユニット付)・DIN端子箱・ベースのみ」ｼﾘｰｽﾞ
                If Me.strcSelection.strKeyKataban = "4" Then
                    'AV110,AV200,AV220の場合
                    If Me.strcSelection.strOpSymbol(9).Trim = "2" Or _
                       Me.strcSelection.strOpSymbol(9).Trim = "5" Or _
                       Me.strcSelection.strOpSymbol(9).Trim = "6" Then
                        'その他電圧価格ｼﾝｸﾞﾙ加算
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = "PV5-S-OTH"
                        decRefAmount(UBound(decRefAmount)) = CDec(Me.strcSelection.strOpSymbol(2).Trim)
                    End If
                End If
            Else
                '基本価格キー
                If Left(strKataban.Trim, 4) = "CMFZ" Then
                    'マニホールドブロック
                    Select Case Me.strcSelection.strOpSymbol(3).Trim
                        Case "HX3"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-BLOCK-02"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-BLOCK-03"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX4"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-BLOCK-02"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-BLOCK-04"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX5"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-BLOCK-03"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-BLOCK-03"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX6"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-BLOCK-03"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-BLOCK-04"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                    End Select

                    'フート
                    Select Case Me.strcSelection.strOpSymbol(5).Trim
                        Case "HY3"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-FOOT-03"
                            decRefAmount(UBound(decRefAmount)) = 1

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-FOOT-04"
                            decRefAmount(UBound(decRefAmount)) = 1
                        Case "HY4"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-FOOT-03"
                            decRefAmount(UBound(decRefAmount)) = 1

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-FOOT-06"
                            decRefAmount(UBound(decRefAmount)) = 1
                        Case "HY5"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-FOOT-04"
                            decRefAmount(UBound(decRefAmount)) = 1

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-FOOT-04"
                            decRefAmount(UBound(decRefAmount)) = 1
                        Case "HY6"
                            'CMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE1-FOOT-04"
                            decRefAmount(UBound(decRefAmount)) = 1

                            'CMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMFZ-BASE2-FOOT-06"
                            decRefAmount(UBound(decRefAmount)) = 1
                    End Select

                    'ミックスブロック
                    For intLoopCnt = 1 To Me.strcSelection.strOptionKataban.Length - 1
                        Select Case Me.strcSelection.strOptionKataban(intLoopCnt).Trim
                            Case "CMFBZ-00L", "CMFBZ-00R"
                                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                                strRefKataban(UBound(strRefKataban)) = Me.strcSelection.strOptionKataban(intLoopCnt).Trim
                                decRefAmount(UBound(decRefAmount)) = CDec(Me.strcSelection.intQuantity(intLoopCnt))
                        End Select
                    Next
                Else
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = Mid(strKataban.Trim, 1, InStr(1, strKataban.Trim, "-")) & "MANIHOLD-BASE"
                    decRefAmount(UBound(decRefAmount)) = 1
                End If

                'A・Bポート口径価格キー
                '裏配管の場合
                If Me.strcSelection.strOpSymbol(4).Trim = "Z" Then
                    Select Case Me.strcSelection.strOpSymbol(3).Trim
                        Case "HX1"
                            'CMF1
                            '02加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-02Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX2"
                            'CMF2
                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '04加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-04Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX3"
                            'CMFZ
                            '02加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF1-BASE-PORT-02Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF2-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX4"
                            'CMFZ
                            '02加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF1-BASE-PORT-02Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '04加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF2-BASE-PORT-04Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX5"
                            'CMFZ
                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF1-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF2-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX6"
                            'CMFZ
                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF1-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '04加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "CMF2-BASE-PORT-04Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case Else
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-" & _
                                                                   Me.strcSelection.strOpSymbol(3).Trim & _
                                                                   Me.strcSelection.strOpSymbol(4).Trim
                            decRefAmount(UBound(decRefAmount)) = CDec(Me.strcSelection.strOpSymbol(2).Trim)
                    End Select
                End If

                'P・Rポート口径価格キー
                'PRポートの加算は"06"または"HY2"の場合のみで、"06"・"HY2"はCMF2の場合にしかありえない
                If Left(strKataban.Trim, 4) = "CMF2" Then
                    Select Case Me.strcSelection.strOpSymbol(5).Trim
                        Case "04"
                        Case Else
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-" & _
                                                                   Me.strcSelection.strOpSymbol(5).Trim
                            decRefAmount(UBound(decRefAmount)) = 1
                    End Select
                End If

                '制御ユニット価格キー
                '制御ユニット付のものは加算
                If Me.strcSelection.strKeyKataban.Trim = "4" Or _
                   Me.strcSelection.strKeyKataban.Trim = "5" Or _
                   Me.strcSelection.strKeyKataban.Trim = "9" Then
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = "CMF1-UNIT-" & Me.strcSelection.strOpSymbol(5).Trim
                    decRefAmount(UBound(decRefAmount)) = 1
                ElseIf Me.strcSelection.strKeyKataban.Trim = "6" Or _
                       Me.strcSelection.strKeyKataban.Trim = "7" Then
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = "CMF1-BASE-UNIT-" & Me.strcSelection.strOpSymbol(5).Trim
                    decRefAmount(UBound(decRefAmount)) = 1
                End If

                'サイレンサボックス価格キー
                '制御ユニット付でないものは、サイレンサボックスが選択可能
                If Me.strcSelection.strKeyKataban.Trim = "1" Or _
                   Me.strcSelection.strKeyKataban.Trim = "2" Or _
                   Me.strcSelection.strKeyKataban.Trim = "3" Or _
                   Me.strcSelection.strKeyKataban.Trim = "8" Then
                    If Me.strcSelection.strOpSymbol(8).Trim <> "" Then
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = "CMF-BASE-" & Me.strcSelection.strOpSymbol(8).Trim
                        decRefAmount(UBound(decRefAmount)) = 1
                    End If
                End If
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Sub

    ''' <summary>
    ''' CMF価格キー設定
    ''' </summary>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strRefKataban">価格キー</param>
    ''' <param name="decRefAmount">個数</param>
    ''' <remarks>CMFの電磁弁の価格キーを設定する</remarks>
    Private Sub subCMFValvePriceKeyGet(ByVal strKataban As String, ByRef strRefKataban() As String, _
                                       ByRef decRefAmount() As Decimal)

        Dim objKataban As New KHKataban

        Try
            ReDim strRefKataban(0)
            ReDim decRefAmount(0)

            If Left(strKataban, 2) = "CM" Then
                'マスキングプレート
                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                strRefKataban(UBound(strRefKataban)) = strKataban.Trim
                decRefAmount(UBound(decRefAmount)) = 1
            ElseIf Left(strKataban, 6) = "PV5-6-" Or _
                   Left(strKataban, 6) = "PV5-8-" Then
                '旧ISOバルブ(小牧分)
                'その他電圧指定時の指定電圧部分を削除する
                If InStr(1, strKataban, "-AC") <> 0 Then
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, InStr(1, strKataban.Trim, "-AC") - 1)
                Else
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = strKataban.Trim
                End If
                '切削油対応部分を削除する
                If InStr(1, strRefKataban(UBound(strRefKataban)), "-F1AW") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(1, strKataban.Trim, "-F1AW") - 1)
                End If
                'その他電圧("-9")を"-1"へ変更する
                If InStr(1, strRefKataban(UBound(strRefKataban)), "-9") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Replace(strRefKataban(UBound(strRefKataban)), "-9", "-1")
                End If
                '不要なハイフンを削除する
                strRefKataban(UBound(strRefKataban)) = KHKataban.fncHypenCut(strRefKataban(UBound(strRefKataban)))
                decRefAmount(UBound(decRefAmount)) = 1

                '切削油対応価格加算
                If InStr(1, strKataban, "-F1AW") <> 0 Then
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 3) & "-F1AW"
                    decRefAmount(UBound(decRefAmount)) = 1
                End If

                ' その他電圧価格加算
                If InStr(1, strKataban, "-9") <> 0 Then
                    If InStr(1, strKataban, "-S") <> 0 Then
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 3) & "-S-OTH"
                        decRefAmount(UBound(decRefAmount)) = 1
                    Else
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 3) & "-D-OTH"
                        decRefAmount(UBound(decRefAmount)) = 1
                    End If
                End If
            Else
                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                strRefKataban(UBound(strRefKataban)) = strKataban.Trim
                '切削油対応部分を削除する
                If InStr(1, strRefKataban(UBound(strRefKataban)), "A-") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(1, strKataban, "A-") - 1) & Mid(strRefKataban(UBound(strRefKataban)), InStr(1, strKataban, "A-") + 1, Len(strRefKataban(UBound(strRefKataban))))
                End If
                If Right(strRefKataban(UBound(strRefKataban)), 1) = "A" Then
                    strRefKataban(UBound(strRefKataban)) = Mid(strRefKataban(UBound(strRefKataban)), 1, Len(strRefKataban(UBound(strRefKataban))) - 1)
                End If
                '電圧部分を削除する
                If InStr(6, strRefKataban(UBound(strRefKataban)), "-1") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-1")) & Mid(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-1") + 2, Len(strRefKataban(UBound(strRefKataban))))
                End If

                If InStr(6, strRefKataban(UBound(strRefKataban)), "-2") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-2")) & Mid(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-2") + 2, Len(strRefKataban(UBound(strRefKataban))))
                End If

                If InStr(6, strRefKataban(UBound(strRefKataban)), "-3") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-3")) & Mid(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-3") + 2, Len(strRefKataban(UBound(strRefKataban))))
                End If

                If InStr(6, strRefKataban(UBound(strRefKataban)), "-4") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-4")) & Mid(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-4") + 2, Len(strRefKataban(UBound(strRefKataban))))
                End If

                If InStr(6, strRefKataban(UBound(strRefKataban)), "-5") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-5")) & Mid(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-5") + 2, Len(strRefKataban(UBound(strRefKataban))))
                End If

                If InStr(6, strRefKataban(UBound(strRefKataban)), "-6") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-6")) & Mid(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-6") + 2, Len(strRefKataban(UBound(strRefKataban))))
                End If
                '不要なハイフンを削除する
                strRefKataban(UBound(strRefKataban)) = KHKataban.fncHypenCut(strRefKataban(UBound(strRefKataban)))
                decRefAmount(UBound(decRefAmount)) = 1

                '切削油対応価格加算
                If InStr(1, strKataban, "A-") <> 0 Or Right(strKataban, 1) = "A" Then
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 3) & "-A"
                    decRefAmount(UBound(decRefAmount)) = 1
                End If

                'その他電圧価格加算
                If InStr(6, strKataban, "-2") <> 0 Or _
                   InStr(6, strKataban, "-5") <> 0 Or _
                   InStr(6, strKataban, "-6") <> 0 Then
                    If InStr(1, strKataban, "-S") <> 0 Then
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 3) & "-S-OTH"
                        decRefAmount(UBound(decRefAmount)) = 1
                    Else
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 3) & "-D-OTH"
                        decRefAmount(UBound(decRefAmount)) = 1
                    End If
                End If
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objKataban = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' GMF価格キー設定
    ''' </summary>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strRefKataban">価格キー</param>
    ''' <param name="decRefAmount">個数</param>
    ''' <remarks>GMFのベースの価格キーを設定する</remarks>
    Private Sub subGMFBasePriceKeyGet(ByVal strKataban As String, _
                                      ByRef strRefKataban() As String, _
                                      ByRef decRefAmount() As Decimal)
        Dim intLoopCnt As Integer
        Try
            ReDim strRefKataban(0)
            ReDim decRefAmount(0)
            If Me.strcSelection.strPriceNo.Trim = "96" Or Me.strcSelection.strPriceNo.Trim = "D3" Then
                '基本価格キー
                If Left(strKataban.Trim, 4) = "GMFZ" Then
                    'マニホールドブロック
                    Select Case Me.strcSelection.strOpSymbol(3).Trim
                        Case "HX3"
                            'GMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE1-BLOCK-02"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            'GMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE2-BLOCK-03"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX4"
                            'GMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE1-BLOCK-02"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            'GMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE2-BLOCK-04"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX5"
                            'GMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE1-BLOCK-03"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            'GMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE2-BLOCK-03"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX6"
                            'GMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE1-BLOCK-03"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            'GMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE2-BLOCK-04"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                    End Select

                    'フート
                    Select Case Me.strcSelection.strOpSymbol(5).Trim
                        Case "HY3"
                            'GMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE1-FOOT-03"
                            decRefAmount(UBound(decRefAmount)) = 1

                            'GMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE2-FOOT-04"
                            decRefAmount(UBound(decRefAmount)) = 1
                        Case "HY4"
                            'GMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE1-FOOT-03"
                            decRefAmount(UBound(decRefAmount)) = 1

                            'GMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE2-FOOT-06"
                            decRefAmount(UBound(decRefAmount)) = 1
                        Case "HY5"
                            'GMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE1-FOOT-04"
                            decRefAmount(UBound(decRefAmount)) = 1

                            'GMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE2-FOOT-04"
                            decRefAmount(UBound(decRefAmount)) = 1
                        Case "HY6"
                            'GMF1
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE1-FOOT-04"
                            decRefAmount(UBound(decRefAmount)) = 1

                            'GMF2
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMFZ-BASE2-FOOT-06"
                            decRefAmount(UBound(decRefAmount)) = 1
                    End Select

                    'ミックスブロック
                    For intLoopCnt = 1 To Me.strcSelection.strOptionKataban.Length - 1
                        Select Case Me.strcSelection.strOptionKataban(intLoopCnt).Trim
                            Case "GMFBZ-00L", "GMFBZ-00R"
                                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                                strRefKataban(UBound(strRefKataban)) = Me.strcSelection.strOptionKataban(intLoopCnt).Trim
                                decRefAmount(UBound(decRefAmount)) = CDec(Me.strcSelection.intQuantity(intLoopCnt))
                        End Select
                    Next
                Else
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = Mid(strKataban.Trim, 1, InStr(1, strKataban.Trim, "-")) & "MANIHOLD-BASE"
                    decRefAmount(UBound(decRefAmount)) = 1
                End If

                'A・Bポート口径価格キー
                '裏配管の場合
                If Me.strcSelection.strOpSymbol(4).Trim = "Z" Then
                    Select Case Me.strcSelection.strOpSymbol(3).Trim
                        Case "HX1"
                            'GMF1
                            '02加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-02Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX2"
                            'GMF2
                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '04加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-04Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX3"
                            'GMFZ
                            '02加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMF1-BASE-PORT-02Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMF2-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX4"
                            'GMFZ
                            '02加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMF1-BASE-PORT-02Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '04加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMF2-BASE-PORT-04Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX5"
                            'GMFZ
                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMF1-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMF2-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case "HX6"
                            'GMFZ
                            '03加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMF1-BASE-PORT-03Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Mid(strKataban.Trim, Len(strKataban.Trim) - 1, 1))

                            '04加算
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = "GMF2-BASE-PORT-04Z"
                            decRefAmount(UBound(decRefAmount)) = CDec(Right(strKataban.Trim, 1))
                        Case Else
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-" & _
                                                                   Me.strcSelection.strOpSymbol(3).Trim & _
                                                                   Me.strcSelection.strOpSymbol(4).Trim
                            decRefAmount(UBound(decRefAmount)) = CDec(Me.strcSelection.strOpSymbol(2).Trim)
                    End Select
                End If

                'P・Rポート口径価格キー
                'PRポートの加算は"06"または"HY2"の場合のみで、"06"・"HY2"はGMF2の場合にしかありえない
                If Left(strKataban.Trim, 4) = "GMF2" Then
                    Select Case Me.strcSelection.strOpSymbol(5).Trim
                        Case "04"
                        Case Else
                            ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                            ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                            strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 4) & "-BASE-PORT-" & _
                                                                   Me.strcSelection.strOpSymbol(5).Trim
                            decRefAmount(UBound(decRefAmount)) = 1
                    End Select
                End If

                '↓2013/03/14
                'フィルタ価格キー
                If Me.strcSelection.strKeyKataban.Trim = "1" Or Me.strcSelection.strKeyKataban.Trim = "2" Or _
                   Me.strcSelection.strKeyKataban.Trim = "3" Then
                    If Len(Me.strcSelection.strOpSymbol(8).Trim) <> 0 Or Len(Me.strcSelection.strOpSymbol(10).Trim) <> 0 Then
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = "GMF*" & Me.strcSelection.strOpSymbol(2).Trim & "-BASE-F"
                        decRefAmount(UBound(decRefAmount)) = 1
                    End If
                End If

                '↓2013/03/14
                'サイレンサボックス価格キー
                '制御ユニット付でないものは、サイレンサボックスが選択可能
                If Me.strcSelection.strKeyKataban.Trim = "1" Or Me.strcSelection.strKeyKataban.Trim = "2" Or _
                   Me.strcSelection.strKeyKataban.Trim = "3" Then
                    If Me.strcSelection.strOpSymbol(9).Trim <> "" Then
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = "GMF-BASE-" & Me.strcSelection.strOpSymbol(9).Trim
                        decRefAmount(UBound(decRefAmount)) = 1
                    End If
                End If
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Sub

    ''' <summary>
    ''' GMF価格キー設定
    ''' </summary>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strRefKataban">価格キー</param>
    ''' <param name="decRefAmount">個数</param>
    ''' <remarks>GMFの電磁弁の価格キーを設定する</remarks>
    Private Sub subGMFValvePriceKeyGet(ByVal strKataban As String, _
                                       ByRef strRefKataban() As String, _
                                       ByRef decRefAmount() As Decimal)

        Dim objKataban As New KHKataban
        Try
            ReDim strRefKataban(0)
            ReDim decRefAmount(0)

            If Left(strKataban, 2) = "CM" Then
                'マスキングプレート
                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                strRefKataban(UBound(strRefKataban)) = strKataban.Trim
                decRefAmount(UBound(decRefAmount)) = 1
            ElseIf Left(strKataban, 6) = "PV5-6-" Or _
                   Left(strKataban, 6) = "PV5-8-" Then
                '旧ISOバルブ(小牧分)
                'その他電圧指定時の指定電圧部分を削除する
                If InStr(1, strKataban, "-AC") <> 0 Then
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, InStr(1, strKataban.Trim, "-AC") - 1)
                Else
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = strKataban.Trim
                End If
                '切削油対応部分を削除する
                If InStr(1, strRefKataban(UBound(strRefKataban)), "-F1AW") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(1, strKataban.Trim, "-F1AW") - 1)
                End If
                'その他電圧("-9")を"-1"へ変更する
                If InStr(1, strRefKataban(UBound(strRefKataban)), "-9") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Replace(strRefKataban(UBound(strRefKataban)), "-9", "-1")
                End If
                '不要なハイフンを削除する
                strRefKataban(UBound(strRefKataban)) = KHKataban.fncHypenCut(strRefKataban(UBound(strRefKataban)))
                decRefAmount(UBound(decRefAmount)) = 1

                '切削油対応価格加算
                If InStr(1, strKataban, "-F1AW") <> 0 Then
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 3) & "-F1AW"
                    decRefAmount(UBound(decRefAmount)) = 1
                End If

                ' その他電圧価格加算
                If InStr(1, strKataban, "-9") <> 0 Then
                    If InStr(1, strKataban, "-S") <> 0 Then
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 3) & "-S-OTH"
                        decRefAmount(UBound(decRefAmount)) = 1
                    Else
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 3) & "-D-OTH"
                        decRefAmount(UBound(decRefAmount)) = 1
                    End If
                End If
            Else
                ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                strRefKataban(UBound(strRefKataban)) = strKataban.Trim
                '切削油対応部分を削除する
                If InStr(1, strRefKataban(UBound(strRefKataban)), "A-") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(1, strKataban, "A-") - 1) & Mid(strRefKataban(UBound(strRefKataban)), InStr(1, strKataban, "A-") + 1, Len(strRefKataban(UBound(strRefKataban))))
                End If
                If Right(strRefKataban(UBound(strRefKataban)), 1) = "A" Then
                    strRefKataban(UBound(strRefKataban)) = Mid(strRefKataban(UBound(strRefKataban)), 1, Len(strRefKataban(UBound(strRefKataban))) - 1)
                End If
                '電圧部分を削除する
                If InStr(6, strRefKataban(UBound(strRefKataban)), "-1") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-1")) & Mid(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-1") + 2, Len(strRefKataban(UBound(strRefKataban))))
                End If

                If InStr(6, strRefKataban(UBound(strRefKataban)), "-2") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-2")) & Mid(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-2") + 2, Len(strRefKataban(UBound(strRefKataban))))
                End If

                If InStr(6, strRefKataban(UBound(strRefKataban)), "-3") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-3")) & Mid(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-3") + 2, Len(strRefKataban(UBound(strRefKataban))))
                End If

                If InStr(6, strRefKataban(UBound(strRefKataban)), "-4") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-4")) & Mid(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-4") + 2, Len(strRefKataban(UBound(strRefKataban))))
                End If

                If InStr(6, strRefKataban(UBound(strRefKataban)), "-5") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-5")) & Mid(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-5") + 2, Len(strRefKataban(UBound(strRefKataban))))
                End If

                If InStr(6, strRefKataban(UBound(strRefKataban)), "-6") <> 0 Then
                    strRefKataban(UBound(strRefKataban)) = Left(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-6")) & Mid(strRefKataban(UBound(strRefKataban)), InStr(6, strKataban, "-6") + 2, Len(strRefKataban(UBound(strRefKataban))))
                End If

                '不要なハイフンを削除する
                strRefKataban(UBound(strRefKataban)) = KHKataban.fncHypenCut(strRefKataban(UBound(strRefKataban)))
                decRefAmount(UBound(decRefAmount)) = 1

                '切削油対応価格加算
                If InStr(1, strKataban, "A-") <> 0 Or Right(strKataban, 1) = "A" Then
                    ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                    ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                    strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 3) & "-A"
                    decRefAmount(UBound(decRefAmount)) = 1
                End If

                'その他電圧価格加算
                If InStr(6, strKataban, "-2") <> 0 Or InStr(6, strKataban, "-5") <> 0 Or InStr(6, strKataban, "-6") <> 0 Then
                    If InStr(1, strKataban, "-S") <> 0 Then
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 3) & "-S-OTH"
                        decRefAmount(UBound(decRefAmount)) = 1
                    Else
                        ReDim Preserve strRefKataban(UBound(strRefKataban) + 1)
                        ReDim Preserve decRefAmount(UBound(decRefAmount) + 1)
                        strRefKataban(UBound(strRefKataban)) = Left(strKataban.Trim, 3) & "-D-OTH"
                        decRefAmount(UBound(decRefAmount)) = 1
                    End If
                End If
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objKataban = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 引当シリーズ形番取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks>引当シリーズ形番を読み込み単価情報を取得し返却する</remarks>
    Private Sub subSelKtbnStrcSelect(objCon As SqlConnection)
        Dim dt As New DataTable
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            dt = dalKtbnStrc.fncSelKtbnStrcSelect(objCon, Me.strcSelection.strUserId, Me.strcSelection.strSessionId)

            For Each dr As DataRow In dt.Rows
                With Me.strcSelection
                    ReDim Preserve .strOpSymbol(UBound(.strOpSymbol) + 1)
                    .strOpSymbol(UBound(.strOpSymbol)) = IIf(IsDBNull(dr("option_symbol")), "", dr("option_symbol"))
                    ReDim Preserve .strOpElementDiv(UBound(.strOpElementDiv) + 1)
                    .strOpElementDiv(UBound(.strOpElementDiv)) = dr("element_div")
                    ReDim Preserve .strOpStructureDiv(UBound(.strOpStructureDiv) + 1)
                    .strOpStructureDiv(UBound(.strOpStructureDiv)) = dr("structure_div")
                    ReDim Preserve .strOpAdditionDiv(UBound(.strOpAdditionDiv) + 1)
                    .strOpAdditionDiv(UBound(.strOpAdditionDiv)) = dr("addition_div")
                    ReDim Preserve .strOpHyphenDiv(UBound(.strOpHyphenDiv) + 1)
                    .strOpHyphenDiv(UBound(.strOpHyphenDiv)) = dr("hyphen_div")
                    ReDim Preserve .strOpKtbnStrcNm(UBound(.strOpKtbnStrcNm) + 1)
                    .strOpKtbnStrcNm(UBound(.strOpKtbnStrcNm)) = dr("ktbn_strc_nm")
                    ReDim Preserve .strOpCountryDiv(UBound(.strOpCountryDiv) + 1)
                    .strOpCountryDiv(UBound(.strOpCountryDiv)) = dr("place_lvl")
                End With
            Next

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当積上単価構成取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks>引当積上単価構成を読み込み単価情報を取得し返却する</remarks>
    Private Sub subSelAccPriceStrcSelect(objCon As SqlConnection)
        Dim dt As New DataTable
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            dt = dalKtbnStrc.fncSelAccPriceStrcSelect(objCon, Me.strcSelection.strUserId, Me.strcSelection.strSessionId)

            For Each dr As DataRow In dt.Rows
                With Me.strcSelection
                    ReDim Preserve .strOpKataban(UBound(.strOpKataban) + 1)
                    .strOpKataban(UBound(.strOpKataban)) = dr("kataban")
                    ReDim Preserve .strOpKatabanCheckDiv(UBound(.strOpKatabanCheckDiv) + 1)
                    .strOpKatabanCheckDiv(UBound(.strOpKatabanCheckDiv)) = dr("kataban_check_div")
                    ReDim Preserve .strOpPlaceCd(UBound(.strOpPlaceCd) + 1)
                    .strOpPlaceCd(UBound(.strOpPlaceCd)) = dr("place_cd")
                    ReDim Preserve .intOpListPrice(UBound(.intOpListPrice) + 1)
                    .intOpListPrice(UBound(.intOpListPrice)) = dr("ls_price")
                    ReDim Preserve .intOpRegPrice(UBound(.intOpRegPrice) + 1)
                    .intOpRegPrice(UBound(.intOpRegPrice)) = dr("rg_price")
                    ReDim Preserve .intOpSsPrice(UBound(.intOpSsPrice) + 1)
                    .intOpSsPrice(UBound(.intOpSsPrice)) = dr("ss_price")
                    ReDim Preserve .intOpBsPrice(UBound(.intOpBsPrice) + 1)
                    .intOpBsPrice(UBound(.intOpBsPrice)) = dr("bs_price")
                    ReDim Preserve .intOpGsPrice(UBound(.intOpGsPrice) + 1)
                    .intOpGsPrice(UBound(.intOpGsPrice)) = dr("gs_price")
                    ReDim Preserve .intOpPsPrice(UBound(.intOpPsPrice) + 1)
                    .intOpPsPrice(UBound(.intOpPsPrice)) = dr("ps_price")
                    ReDim Preserve .decOpamount(UBound(.decOpamount) + 1)
                    .decOpamount(UBound(.decOpamount)) = dr("amount")
                    .strCurrency = dr("currency_cd")
                    .strMadeCountry = dr("country_cd")
                End With
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Sub

    ''' <summary>
    ''' 引当仕様書情報取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks>引当仕様書構成を読み込み仕様書情報を取得し返却する</remarks>
    Private Sub subSelSpecSelect(objCon As SqlConnection)
        Dim dt As New DataTable
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            dt = dalKtbnStrc.fncSelSpecSelect(objCon, Me.strcSelection.strUserId, Me.strcSelection.strSessionId)

            If dt.Rows.Count > 0 Then
                With Me.strcSelection
                    .strModelNo = IIf(IsDBNull(dt.Rows(0)("model_no")), "", dt.Rows(0)("model_no"))
                    .strWiringSpec = IIf(IsDBNull(dt.Rows(0)("wiring_spec")), "", dt.Rows(0)("wiring_spec"))
                    .decDinRailLength = dt.Rows(0)("din_rail_length")
                End With
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当仕様書構成取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks></remarks>
    Private Sub subSelSpecStrcSelect(objCon As SqlConnection)
        Dim dt As New DataTable
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            dt = dalKtbnStrc.fncSelSpecStrcSelect(objCon, Me.strcSelection.strUserId, Me.strcSelection.strSessionId)

            For Each dr As DataRow In dt.Rows
                With Me.strcSelection
                    ReDim Preserve .strAttributeSymbol(UBound(.strAttributeSymbol) + 1)
                    .strAttributeSymbol(UBound(.strAttributeSymbol)) = IIf(IsDBNull(dr("attribute_symbol")), "", dr("attribute_symbol"))
                    ReDim Preserve .strOptionKataban(UBound(.strOptionKataban) + 1)
                    .strOptionKataban(UBound(.strOptionKataban)) = IIf(IsDBNull(dr("option_kataban")), "", dr("option_kataban"))
                    ReDim Preserve .strCXAKataban(UBound(.strCXAKataban) + 1)
                    .strCXAKataban(UBound(.strCXAKataban)) = IIf(IsDBNull(dr("cxa_kataban")), "", dr("cxa_kataban"))
                    ReDim Preserve .strCXBKataban(UBound(.strCXBKataban) + 1)
                    .strCXBKataban(UBound(.strCXBKataban)) = IIf(IsDBNull(dr("cxb_kataban")), "", dr("cxb_kataban"))
                    ReDim Preserve .strPositionInfo(UBound(.strPositionInfo) + 1)
                    .strPositionInfo(UBound(.strPositionInfo)) = IIf(IsDBNull(dr("position_info")), "", dr("position_info"))
                    ReDim Preserve .intQuantity(UBound(.intQuantity) + 1)
                    .intQuantity(UBound(.intQuantity)) = dr("quantity")
                End With
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当ロッド先端特注WF標準寸法取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <remarks>引当ロッド先端特注を読み込み引当ロッド先端特注WF標準寸法を取得し返却する</remarks>
    Private Sub subSelRodWFSelect(objCon As SqlConnection)
        Dim dt As New DataTable
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            '初期設定
            strcSelection.strRodEndWFStdVal = "0"

            dt = dalKtbnStrc.fncSelRodWFSelect(objCon, Me.strcSelection.strUserId, Me.strcSelection.strSessionId)

            If dt.Rows.Count > 0 Then
                strcSelection.strRodEndWFStdVal = IIf(IsDBNull(dt.Rows(0)("normal_value")), "0", dt.Rows(0)("normal_value"))
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub


End Class
