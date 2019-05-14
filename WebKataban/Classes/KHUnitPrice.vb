Imports System.Data.SqlClient
Imports WebKataban.ClsCommon
Imports WebKataban.CdCst

Public Class KHUnitPrice

    Private Property dalUnitPrice As New UnitPriceDAL

    ''' <summary>
    ''' 引当単価情報設定
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="strOfficeCd"></param>
    ''' <remarks>引当てた価格情報を引当積上単価情報に設定する</remarks>
    Public Sub subPriceInfoSet(objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, ByVal strUserId As String, _
                               ByVal strSessionId As String, ByVal strCountryCd As String, _
                               ByVal strOfficeCd As String, ByRef strStorageLocation As String, ByRef strEvaluationType As String)
        Dim strOpRefKataban() As String = Nothing
        Dim decOpAmount() As Decimal = Nothing
        Dim strPriceDiv() As String = Nothing
        Dim dalKtbnStrc As New KtbnStrcDAL

        Try
            '選択形番情報削除処理
            Call dalKtbnStrc.subAccPriceStrcDel(objCon, strUserId, strSessionId)

            '選択形番情報取得
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, strUserId, strSessionId, 1)

            '単価情報読み込み(フル形番検索)
            If Not Me.fncSelectPriceFull(objCon, objKtbnStrc, strCountryCd, strOfficeCd, Nothing, Nothing, strStorageLocation, strEvaluationType) Then
                '単価情報存在しなかった場合は積み上げ処理

                '機種毎に価格算出
                fncGetPriceInfo(objKtbnStrc.strcSelection.strPriceNo.Trim, objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv, strCountryCd, strOfficeCd)

                '引当積上単価構成登録
                Call objKtbnStrc.subAccPriceStrcReg(objCon, strUserId, strSessionId, strCountryCd, strOfficeCd, _
                                                    strOpRefKataban, decOpAmount, strPriceDiv, objKtbnStrc.strcSelection.strCurrency)
                '引当シリーズ形番更新処理
                Call objKtbnStrc.subSelSrsKtbnUpdate(objCon, objKtbnStrc, "1", strUserId, strSessionId)
            Else
                '引当シリーズ形番更新処理
                Call objKtbnStrc.subSelSrsKtbnUpdate(objCon, objKtbnStrc, "0", strUserId, strSessionId)
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Sub

    ''' <summary>
    ''' 引当単価情報設定
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="strOfficeCd"></param>
    ''' <param name="DS_Tab"></param>
    ''' <remarks>引当てた価格情報を引当積上単価情報に設定する</remarks>
    Public Sub subPriceInfoSet_ForkatOut(objCon As SqlConnection, ByRef objKtbnStrc As KHKtbnStrc, _
                                         ByVal strCountryCd As String, ByVal strOfficeCd As String, _
                                         Optional DS_Tab As DataSet = Nothing)
        Dim strOpRefKataban() As String = Nothing
        Dim decOpAmount() As Decimal = Nothing
        Dim strPriceDiv() As String = Nothing

        Try
            '単価情報読み込み(フル形番検索)
            If Not Me.fncSelectPriceFull(objCon, objKtbnStrc, strCountryCd, strOfficeCd, 1, DS_Tab) Then
                '単価情報存在しなかった場合は積み上げ処理
                '機種毎に価格算出

                '初期化されないのもを初期化する(初期値はnothingで設定するのはダメ)
                subInitObjKtbnstrc(objKtbnStrc)                                    'ADD BY YGY 20140729

                If objKtbnStrc.strcSelection.strPriceNo Is Nothing OrElse objKtbnStrc.strcSelection.strPriceNo.Equals(String.Empty) Then Exit Sub

                '機種毎に価格算出
                fncGetPriceInfo(objKtbnStrc.strcSelection.strPriceNo.Trim, objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv, strCountryCd, strOfficeCd)

                '引当積上単価構成登録
                If objKtbnStrc.strcSelection.strCurrency Is Nothing Then
                    objKtbnStrc.strcSelection.strCurrency = "JPY"
                End If
                Call objKtbnStrc.subAccPriceStrcReg(objCon, "", "", strCountryCd, strOfficeCd, _
                                                    strOpRefKataban, decOpAmount, strPriceDiv, objKtbnStrc.strcSelection.strCurrency, objKtbnStrc, 1, DS_Tab)
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Sub

    ''' <summary>
    ''' 価格情報の取得
    ''' </summary>
    ''' <param name="strPriceNo"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strOpRefKataban"></param>
    ''' <param name="decOpAmount"></param>
    ''' <param name="strPriceDiv"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strOfficeCd"></param>
    ''' <remarks></remarks>
    Private Sub fncGetPriceInfo(ByVal strPriceNo As String, _
                                ByVal objKtbnStrc As KHKtbnStrc, _
                                ByRef strOpRefKataban() As String, _
                                ByRef decOpAmount() As Decimal, _
                                ByRef strPriceDiv() As String, _
                                ByRef strCountryCd As String, _
                                ByRef strOfficeCd As String)

        Select Case strPriceNo
            Case "01"
                Call KHPrice01.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv, strCountryCd, strOfficeCd)
            Case "02"
                Call KHPrice02.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv, strCountryCd, strOfficeCd)
            Case "03"
                Call KHPrice03.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv, strCountryCd, strOfficeCd)
            Case "04"
                Call KHPrice04.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strCountryCd, strOfficeCd)
            Case "05"
                Call KHPrice05.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "06"
                Call KHPrice06.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strCountryCd, strOfficeCd)
            Case "07"
                Call KHPrice07.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "08"
                Call KHPrice08.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "09"
                Call KHPrice09.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "10"
                Call KHPrice10.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "11"
                Call KHPrice11.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "12"
                Call KHPrice12.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "13"
                Call KHPrice13.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "14"
                Call KHPrice14.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "15"
                Call KHPrice15.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "16"
                Call KHPrice16.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strCountryCd, strOfficeCd)
            Case "17"
                Call KHPrice17.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv, strCountryCd, strOfficeCd)
            Case "18"
                Call KHPrice18.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv, strCountryCd, strOfficeCd)
            Case "19"
                Call KHPrice19.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "20"
                Call KHPrice20.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "21"
                Call KHPrice21.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "22"
                Call KHPrice22.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "23"
                Call KHPrice23.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "24"
                Call KHPrice24.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "25"
                Call KHPrice25.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "26"
                Call KHPrice26.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)  'RM1306001 2013/06/06
            Case "27"
                Call KHPrice27.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "28"
                Call KHPrice28.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "29"
                Call KHPrice29.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "30"
                Call KHPrice30.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "31"
                Call KHPrice31.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "32"
                Call KHPrice32.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "33"
                Call KHPrice33.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "34"
                Call KHPrice34.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strCountryCd, strOfficeCd)
            Case "35"
                Call KHPrice35.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strCountryCd, strOfficeCd)
            Case "36"
                Call KHPrice36.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strCountryCd, strOfficeCd)
            Case "37"
                Call KHPrice37.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "38"
                Call KHPrice38.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "39"
                Call KHPrice39.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "40"
                Call KHPrice40.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "41"
                Call KHPrice41.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "42"
                Call KHPrice42.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "43"
                Call KHPrice43.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "44"
                Call KHPrice44.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "45"
                Call KHPrice45.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "46"
                Call KHPrice46.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "47"
                Call KHPrice47.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "48"
                Call KHPrice48.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "49"
                Call KHPrice49.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "50"
                Call KHPrice50.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "51"
                Call KHPrice51.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "52"
                Call KHPrice52.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "53"
                Call KHPrice53.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "54"
                Call KHPrice54.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "55"
                Call KHPrice55.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strCountryCd, strOfficeCd)
            Case "56"
                Call KHPrice56.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "57"
                Call KHPrice57.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "58"
                Call KHPrice58.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "59"
                Call KHPrice59.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "60"
                Call KHPrice60.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "61"
                Call KHPrice61.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "62"
                Call KHPrice62.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "63"
                Call KHPrice63.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "64"
                Call KHPrice64.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "65"
                Call KHPrice65.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "66"
                Call KHPrice66.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "67"
                Call KHPrice67.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "68"
                Call KHPrice68.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "69"
                Call KHPrice69.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "70"
                Call KHPrice70.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "71"
                Call KHPrice71.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "72"
                Call KHPrice72.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "73"
                Call KHPrice73.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "74"
                Call KHPrice74.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "75"
                Call KHPrice75.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "76"
                Call KHPrice76.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "77"
                Call KHPrice77.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "78"
                Call KHPrice78.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "79"
                Call KHPrice79.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "80"
                Call KHPrice80.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "81"
                Call KHPrice81.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "82"
                Call KHPrice82.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "83"
                Call KHPrice83.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "84"
                Call KHPrice84.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "85"
                Call KHPrice85.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "86"
                Call KHPrice86.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "87"
                Call KHPrice87.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "88"
                Call KHPrice88.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "89"
                Call KHPrice89.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "90"
                Call KHPrice90.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "91"
                'Call KHPrice91.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "92"
                'Call KHPrice92.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "93"
                'Call KHPrice93.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "94"
                'Call KHPrice94.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "95"
                'Call KHPrice95.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "96"
                Call KHPrice96.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "97"
                Call KHPrice97.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "98"
                Call KHPrice98.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "99"
                Call KHPrice99.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "A0"
                Call KHPriceA0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "A1"
                Call KHPriceA1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "A2"
                Call KHPriceA2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "A3"
                Call KHPriceA3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "A4"
                Call KHPriceA4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "A5"
                Call KHPriceA5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "A6"
                Call KHPriceA6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "A7"
                Call KHPriceA7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "A8"
                Call KHPriceA8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "A9"
                Call KHPriceA9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "B0"
                Call KHPriceB0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "B1"
                Call KHPriceB1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "B2"
                Call KHPriceB2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "B3"
                Call KHPriceB3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "B4"
                Call KHPriceB4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "B5"
                Call KHPriceB5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "B6"
                Call KHPriceB6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "B7"
                Call KHPriceB7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "B8"
                Call KHPriceB8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "B9"
                Call KHPriceB9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "C0"
                Call KHPriceC0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "C1"
                Call KHPriceC1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "C2"
                Call KHPriceC2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "C3"
                Call KHPriceC3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "C4"
                Call KHPriceC4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "C5"
                Call KHPriceC5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "C6"
                Call KHPriceC6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "C7"
                Call KHPriceC7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "C8"
                Call KHPriceC8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "C9"
                Call KHPriceC9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "D0"
                Call KHPriceD0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "D1"
                Call KHPriceD1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "D2"
                Call KHPriceD2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "D3"
                Call KHPriceD3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "D4"
                Call KHPriceD4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "D5"
                Call KHPriceD5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "D6"
                Call KHPriceD6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "D7"
                Call KHPriceD7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "D8"
                Call KHPriceD8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "D9"
                Call KHPriceD9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "E0"
                Call KHPriceE0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "E1"
                Call KHPriceE1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "E2"
                Call KHPriceE2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "E3"
                Call KHPriceE3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "E4"
                Call KHPriceE4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "E5"
                Call KHPriceE5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "E6"
                Call KHPriceE6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "E7"
                Call KHPriceE7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "E8"
                Call KHPriceE8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "E9"
                Call KHPriceE9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "F0"
                Call KHPriceF0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "F1"
                Call KHPriceF1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "F2"
                Call KHPriceF2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "F3"
                Call KHPriceF3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "F4"
                Call KHPriceF4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "F5"
                '    Call KHPriceF5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "F6"
                '    Call KHPriceF6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "F7"
                Call KHPriceF7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "F8"
                Call KHPriceF8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "F9"
                Call KHPriceF9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "G0"
                Call KHPriceG0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "G1"
                Call KHPriceG1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv) 'RM1610011 strPriceDiv追加
            Case "G2"
                'Call KHPriceG2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "G3"
                'Call KHPriceG3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "G4"
                'Call KHPriceG4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "G5"
                'Call KHPriceG5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "G6"
                Call KHPriceG6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "G7"
                Call KHPriceG7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "G8"
                Call KHPriceG8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "G9"
                Call KHPriceG9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "H0"
                Call KHPriceH0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "H1"
                Call KHPriceH1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "H2"
                Call KHPriceH2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "H3"
                Call KHPriceH3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "H4"
                Call KHPriceH4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "H5"
                Call KHPriceH5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "H6"
                Call KHPriceH6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "H7"
                Call KHPriceH7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "H8"
                Call KHPriceH8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "H9"
                Call KHPriceH9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "I0"
                Call KHPriceI0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "I1"
                Call KHPriceI1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "I2"
                Call KHPriceI2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "I3"
                Call KHPriceI3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "I4"
                Call KHPriceI4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "I5"
                Call KHPriceI5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "I6"
                Call KHPriceI6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "I7"
                Call KHPriceI7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "I8"
                Call KHPriceI8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "I9"
                Call KHPriceI9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "J0"
                Call KHPriceJ0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "J1"
                Call KHPriceJ1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "J2"
                Call KHPriceJ2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "J3"
                Call KHPriceJ3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "J4"
                Call KHPriceJ4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "J5"
                Call KHPriceJ5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "J6"
                Call KHPriceJ6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "J7"
                Call KHPriceJ7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "J8"
                Call KHPriceJ8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "J9"
                Call KHPriceJ9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "K0"
                Call KHPriceK0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "K1"
                Call KHPriceK1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "K2"
                Call KHPriceK2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "K3"
                Call KHPriceK3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "K4"
                Call KHPriceK4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "K5"
                Call KHPriceK5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "K6"
                Call KHPriceK6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "K7"
                Call KHPriceK7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "K8"
                Call KHPriceK8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "K9"
                Call KHPriceK9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "L0"
                Call KHPriceL0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "L1"
                Call KHPriceL1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "L2"
                Call KHPriceL2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "L3"
                Call KHPriceL3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "L4"
                Call KHPriceL4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "L5"
                Call KHPriceL5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "L6"
                Call KHPriceL6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "L7"
                Call KHPriceL7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "L8"
                Call KHPriceL8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "L9"
                Call KHPriceL9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "M0"
                Call KHPriceM0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "M1"
                Call KHPriceM1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "M2"
                Call KHPriceM2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "M3"
                Call KHPriceM3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strCountryCd, strOfficeCd)
            Case "M4"
                Call KHPriceM4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "M5"
                Call KHPriceM5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)  'RM1306001 2013/06/06 追加
            Case "M6"
                Call KHPriceM6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "M7"
                Call KHPriceM7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "M8"
                Call KHPriceM8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "M9"
                Call KHPriceM9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "N0"
                Call KHPriceN0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "N1"
                Call KHPriceN1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "N2"
                Call KHPriceN2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "N3"
                Call KHPriceN3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "N4"
                Call KHPriceN4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "N5"
                Call KHPriceN5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "N6"
                Call KHPriceN6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "N7"
                Call KHPriceN7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "N8"
                Call KHPriceN8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "N9"
                Call KHPriceN9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "O0"
                Call KHPriceO0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "O1"
                Call KHPriceO1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "O2"
                Call KHPriceO2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "O3"
                Call KHPriceO3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "O5"
                Call KHPriceO5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "O6"
                Call KHPriceO6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "O7"
                Call KHPriceO7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "O8"
                Call KHPriceO8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "O9"
                Call KHPriceO9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "OA"
                Call KHPriceOA.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "P0"
                Call KHPriceP0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "P1"
                Call KHPriceP1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "P2"
                Call KHPriceP2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "P3"
                Call KHPriceP3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "P4"
                Call KHPriceP4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "P5"
                Call KHPriceP5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "P6"
                Call KHPriceP6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "P7"
                Call KHPriceP7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "P8"
                Call KHPriceP8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "P9"
                Call KHPriceP9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "Q0"
                Call KHPriceQ0.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "Q1"
                Call KHPriceQ1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "Q2"
                Call KHPriceQ2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "Q3"
                Call KHPriceQ3.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "Q4"
                Call KHPriceQ4.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "Q5"
                Call KHPriceQ5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "Q6"
                Call KHPriceQ6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "Q7"
                Call KHPriceQ7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "Q8"
                Call KHPriceQ8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "Q9"
                Call KHPriceQ9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "R1"
                Call KHPriceR1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "R2"
                Call KHPriceR2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "R5"
                Call KHPriceR5.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "R6"
                Call KHPriceR6.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "R7"
                Call KHPriceR7.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "R8"
                Call KHPriceR8.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
            Case "R9"
                Call KHPriceR9.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "S1"
                Call KHPriceS1.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            Case "S2"
                'RM1708016 2017/8/22
                Call KHPriceS2.subPriceCalculation(objKtbnStrc, strOpRefKataban, decOpAmount)
        End Select

    End Sub


    ''' <summary>
    ''' 単価情報取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc">オブジェクト</param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="intMode"></param>
    ''' <param name="DS_Tab"></param>
    ''' <returns></returns>
    ''' <remarks>単価テーブルを読み込み単価情報を取得し返却する</remarks>
    Public Function fncSelectPriceFull(objCon As SqlConnection, ByRef objKtbnStrc As KHKtbnStrc, _
                                       ByVal strCountryCd As String, _
                                        ByVal strOfficeCd As String, _
                                       Optional intMode As Integer = 0, _
                                       Optional DS_Tab As DataSet = Nothing, _
                                       Optional ByRef strStorageLocation As String = Nothing, _
                                       Optional ByRef strEvaluationType As String = Nothing) As Boolean

        Dim bolReturn As Boolean
        Dim strKatabanCheckDiv As String = Nothing
        Dim strPlaceCd As String = Nothing
        Dim htPriceInfo As Hashtable = Nothing
        Dim htScrewPriceInfo As Hashtable = Nothing

        Dim strOpRefKataban(1) As String
        Dim strOpKatabanCheckDiv(1) As String
        Dim strOpPlaceCd(1) As String
        Dim intOpListPrice(1) As Decimal
        Dim intOpRegPrice(1) As Decimal
        Dim intOpSsPrice(1) As Decimal
        Dim intOpBsPrice(1) As Decimal
        Dim intOpGsPrice(1) As Decimal
        Dim intOpPsPrice(1) As Decimal
        Dim decOpAmount(1) As Decimal
        Dim strCurrency As String = String.Empty
        Dim strMadeCountry As String = String.Empty

        fncSelectPriceFull = False
        Try
            'ｼﾘｰｽﾞ形番検索の場合
            If objKtbnStrc.strcSelection.strDivision = "1" Then
                '指定形番の場合、処理を終了する
                Select Case objKtbnStrc.strcSelection.strSeriesKataban
                    Case "MCP-S"
                        fncSelectPriceFull = False
                        Exit Function
                End Select
            End If

            If DS_Tab Is Nothing Then 'DBから取り込み
                '単価情報読み込み
                '海外仕入れ追加のため
                strCurrency = objKtbnStrc.strcSelection.strCurrency

                If strCurrency Is Nothing OrElse strCurrency.Equals(String.Empty) Then
                    strCurrency = "JPY"
                End If

                '価格取得
                Select Case objKtbnStrc.strcSelection.strDivision
                    Case "3" '仕入品  
                        bolReturn = M_KatabanDAL.fncSelectPrice(objKtbnStrc.strcSelection.strGoodsNm, strKatabanCheckDiv, _
                                              strPlaceCd, htPriceInfo, strMadeCountry, strStorageLocation, strEvaluationType)
                    Case Else
                        bolReturn = Me.fncSelectPrice(objCon, objKtbnStrc.strcSelection.strFullKataban, strKatabanCheckDiv, _
                                              strPlaceCd, htPriceInfo, strCurrency, strMadeCountry)
                End Select

                '積上単価情報読み込み
                If Not bolReturn Then
                    bolReturn = Me.fncSelectAccumulatePrice(objCon, objKtbnStrc.strcSelection.strFullKataban, _
                                                            strKatabanCheckDiv, strPlaceCd, htPriceInfo, _
                                                            strCurrency)
                End If

                'Ｇねじ加算価格取得
                Call Me.subSelectScrewKatabanMst(objCon, objKtbnStrc.strcSelection.strFullKataban, strCountryCd, strOfficeCd, htScrewPriceInfo)
            Else

                '組合せテスト専用
                bolReturn = Me.fncManifoldTest(objKtbnStrc, htPriceInfo, htScrewPriceInfo, DS_Tab, strKatabanCheckDiv, strPlaceCd, strCountryCd, strMadeCountry, strCurrency)
            End If

            '引当積上単価構成追加
            If bolReturn Then
                strOpRefKataban(1) = objKtbnStrc.strcSelection.strFullKataban
                strOpKatabanCheckDiv(1) = strKatabanCheckDiv
                strOpPlaceCd(1) = strPlaceCd
                intOpListPrice(1) = htPriceInfo(CdCst.UnitPrice.ListPrice) - htScrewPriceInfo(CdCst.UnitPrice.ListPrice)
                intOpRegPrice(1) = htPriceInfo(CdCst.UnitPrice.RegPrice) - htScrewPriceInfo(CdCst.UnitPrice.RegPrice)
                intOpSsPrice(1) = htPriceInfo(CdCst.UnitPrice.SsPrice) - htScrewPriceInfo(CdCst.UnitPrice.SsPrice)
                intOpBsPrice(1) = htPriceInfo(CdCst.UnitPrice.BsPrice) - htScrewPriceInfo(CdCst.UnitPrice.BsPrice)
                intOpGsPrice(1) = htPriceInfo(CdCst.UnitPrice.GsPrice) - htScrewPriceInfo(CdCst.UnitPrice.GsPrice)
                intOpPsPrice(1) = htPriceInfo(CdCst.UnitPrice.PsPrice) - htScrewPriceInfo(CdCst.UnitPrice.PsPrice)
                decOpAmount(1) = 1
                If strCurrency.Length <= 0 Then strCurrency = objKtbnStrc.strcSelection.strCurrency
                If strMadeCountry.Length <= 0 Then strMadeCountry = objKtbnStrc.strcSelection.strMadeCountry

                '引当積上単価構成追加
                If intMode <> 0 Then
                    objKtbnStrc.strcSelection.strOpKataban = strOpRefKataban
                    objKtbnStrc.strcSelection.strOpKatabanCheckDiv = strOpKatabanCheckDiv
                    objKtbnStrc.strcSelection.strOpPlaceCd = strOpPlaceCd
                    objKtbnStrc.strcSelection.intOpListPrice = intOpListPrice
                    objKtbnStrc.strcSelection.intOpRegPrice = intOpRegPrice
                    objKtbnStrc.strcSelection.intOpSsPrice = intOpSsPrice
                    objKtbnStrc.strcSelection.intOpBsPrice = intOpBsPrice
                    objKtbnStrc.strcSelection.intOpGsPrice = intOpGsPrice
                    objKtbnStrc.strcSelection.intOpPsPrice = intOpPsPrice
                    objKtbnStrc.strcSelection.decOpamount = decOpAmount
                    objKtbnStrc.strcSelection.strCurrency = strCurrency
                    objKtbnStrc.strcSelection.strMadeCountry = strMadeCountry
                    objKtbnStrc.strcSelection.strFullKataban = objKtbnStrc.strcSelection.strFullKataban
                    objKtbnStrc.strcSelection.strKatabanCheckDiv = strKatabanCheckDiv
                    objKtbnStrc.strcSelection.strPlaceCd = strPlaceCd
                    objKtbnStrc.strcSelection.intListPrice = htPriceInfo(CdCst.UnitPrice.ListPrice) - htScrewPriceInfo(CdCst.UnitPrice.ListPrice)
                    objKtbnStrc.strcSelection.intRegPrice = htPriceInfo(CdCst.UnitPrice.RegPrice) - htScrewPriceInfo(CdCst.UnitPrice.RegPrice)
                    objKtbnStrc.strcSelection.intSsPrice = htPriceInfo(CdCst.UnitPrice.SsPrice) - htScrewPriceInfo(CdCst.UnitPrice.SsPrice)
                    objKtbnStrc.strcSelection.intBsPrice = htPriceInfo(CdCst.UnitPrice.BsPrice) - htScrewPriceInfo(CdCst.UnitPrice.BsPrice)
                    objKtbnStrc.strcSelection.intGsPrice = htPriceInfo(CdCst.UnitPrice.GsPrice) - htScrewPriceInfo(CdCst.UnitPrice.GsPrice)
                    objKtbnStrc.strcSelection.intPsPrice = htPriceInfo(CdCst.UnitPrice.PsPrice) - htScrewPriceInfo(CdCst.UnitPrice.PsPrice)
                    objKtbnStrc.strcSelection.intAmount = 1
                Else
                    Call objKtbnStrc.subInsertAccPriceStrc(objCon, objKtbnStrc.strcSelection.strUserId, _
                                                           objKtbnStrc.strcSelection.strSessionId, _
                                                           strOpRefKataban, strOpKatabanCheckDiv, _
                                                           strOpPlaceCd, intOpListPrice, intOpRegPrice, _
                                                           intOpSsPrice, intOpBsPrice, intOpGsPrice, intOpPsPrice, _
                                                           decOpAmount, strCurrency, strMadeCountry)
                End If
            End If

            '戻り値設定
            If bolReturn Then
                fncSelectPriceFull = True
            Else
                fncSelectPriceFull = False
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 単価情報取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strKatabanCheckDiv">形番チェック区分</param>
    ''' <param name="strPlaceCd">出荷場所</param>
    ''' <param name="htPriceInfo">価格情報</param>
    ''' <param name="strCurrency"></param>
    ''' <param name="strMadeCountry"></param>
    ''' <returns></returns>
    ''' <remarks>単価テーブルを読み込み単価情報を取得し返却する</remarks>
    Public Function fncSelectPrice(objCon As SqlConnection, ByRef strKataban As String, _
                                   ByRef strKatabanCheckDiv As String, ByRef strPlaceCd As String, _
                                   ByRef htPriceInfo As Hashtable, _
                                   ByVal strCurrency As String, ByRef strMadeCountry As String) As Boolean
        Dim blnResult As Boolean = False

        Try
            If dalUnitPrice.fncSelectPrice(objCon, strKataban, strKatabanCheckDiv, strPlaceCd, _
                                       htPriceInfo, strCurrency, strMadeCountry) Then
                blnResult = True
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return blnResult
    End Function

    ''' <summary>
    ''' 積上単価情報取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strKatabanCheckDiv">形番チェック区分</param>
    ''' <param name="strPlaceCd">出荷場所</param>
    ''' <param name="htPriceInfo">価格情報</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectAccumulatePrice(objCon As SqlConnection, ByVal strKataban As String, _
                                             ByRef strKatabanCheckDiv As String, ByRef strPlaceCd As String, _
                                             ByRef htPriceInfo As Hashtable, _
                                             ByVal strCurrency As String) As Boolean
        Dim blnResult As Boolean = False

        Try
            If dalUnitPrice.fncSelectAccumulatePrice(objCon, strKataban, strKatabanCheckDiv, strPlaceCd, htPriceInfo, strCurrency) Then
                blnResult = True
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return blnResult
    End Function

    ''' <summary>
    ''' Ｇねじ形番マスタ取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strKataban">形番</param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="htPriceInfo">価格情報</param>
    ''' <remarks>Ｇねじ形番マスタを読み込み単価情報を取得し返却する</remarks>
    Public Sub subSelectScrewKatabanMst(objCon As SqlConnection, ByVal strKataban As String, _
                                        ByVal strCountryCd As String, ByVal strOfficeCd As String, _
                                        ByRef htPriceInfo As Hashtable)
        Dim blnResult As Boolean = False

        Try
            Call dalUnitPrice.subScrewKatabanMstSelect(objCon, strKataban, strCountryCd, strOfficeCd, htPriceInfo)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 単価リスト取得処理
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="strLanguageCd">言語コード</param>
    ''' <param name="strCurrencyCd">通貨コード</param>
    ''' <param name="intPriceDispLvl">価格表示レベル</param>
    ''' <param name="strPriceList">単価リスト</param>
    ''' <param name="strPriceFCA">ＦＣＡ＃１価格</param>
    ''' <param name="strPriceFCA2">ＦＣＡ＃２価格</param>
    ''' <param name="strCurShipCd">出荷場所</param>
    ''' <param name="objKtbnStrc"></param>
    ''' <remarks></remarks>
    Public Sub subPriceListSelect(objConBase As SqlConnection, ByVal strCountryCd As String, _
                                  ByVal strLanguageCd As String, ByVal strCurrencyCd As String, _
                                  ByVal intPriceDispLvl As Integer, ByRef strPriceList(,) As String, _
                                  ByRef strPriceFCA As String, ByRef strPriceFCA2 As String, _
                                  ByVal strCurShipCd As String, objKtbnStrc As KHKtbnStrc)

        Dim strPrice(,) As String = Nothing
        Dim strPricePrev() As String
        Dim strCurrency() As String
        Dim intLoopCnt As Integer

        Try
            '価格リスト取得
            Call subPriceListGet(objConBase, intPriceDispLvl, strLanguageCd, strPrice)
            Dim dt_CurrMath As DataTable = fncGetCurrMathAll(objConBase)

            '配列定義
            ReDim strPricePrev(0)
            ReDim strCurrency(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case Else
                    For intLoopCnt = 1 To UBound(strPrice)
                        Select Case strPrice(intLoopCnt, 1)
                            Case CdCst.UnitPrice.ListPrice
                                ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                                strPricePrev(UBound(strPricePrev)) = fncGetCurrMathFromDt(objKtbnStrc.strcSelection.strCurrency, _
                                                                     objKtbnStrc.strcSelection.intListPrice.ToString, dt_CurrMath)
                                ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                            Case CdCst.UnitPrice.RegPrice
                                ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                                strPricePrev(UBound(strPricePrev)) = fncGetCurrMathFromDt(objKtbnStrc.strcSelection.strCurrency, objKtbnStrc.strcSelection.intRegPrice.ToString, dt_CurrMath)
                                ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                            Case CdCst.UnitPrice.SsPrice
                                ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                                strPricePrev(UBound(strPricePrev)) = fncGetCurrMathFromDt(objKtbnStrc.strcSelection.strCurrency, objKtbnStrc.strcSelection.intSsPrice.ToString, dt_CurrMath)
                                ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                            Case CdCst.UnitPrice.BsPrice
                                ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                                strPricePrev(UBound(strPricePrev)) = fncGetCurrMathFromDt(objKtbnStrc.strcSelection.strCurrency, objKtbnStrc.strcSelection.intBsPrice.ToString, dt_CurrMath)
                                ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                            Case CdCst.UnitPrice.GsPrice
                                ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                                strPricePrev(UBound(strPricePrev)) = fncGetCurrMathFromDt(objKtbnStrc.strcSelection.strCurrency, objKtbnStrc.strcSelection.intGsPrice.ToString, dt_CurrMath)
                                ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                            Case CdCst.UnitPrice.PsPrice
                                ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                                strPricePrev(UBound(strPricePrev)) = fncGetCurrMathFromDt(objKtbnStrc.strcSelection.strCurrency, objKtbnStrc.strcSelection.intPsPrice.ToString, dt_CurrMath)
                                ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                            Case CdCst.UnitPrice.APrice
                                ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)

                                '仕入品の場合は表示しない
                                'If objKtbnStrc.strcSelection.strKatabanCheckDiv = "5" Then
                                If objKtbnStrc.strcSelection.strDivision = "3" Then
                                    strPricePrev(UBound(strPricePrev)) = 0
                                    ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                    strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                                Else
                                    '為替レートを取得する  (生産国→ログイン国の為替レートを取得する)
                                    '画面の出荷場所＝ログインユーザーの国コード＝PRC かつ　基準通貨<>JPY
                                    'RM1805001_ListPriceが0以外の場合処理追加
                                    If objKtbnStrc.strcSelection.strMadeCountry = strCountryCd And strCountryCd = "PRC" And objKtbnStrc.strcSelection.strCurrency <> "JPY" And _
                                        objKtbnStrc.strcSelection.intListPrice > 0 Then
                                        strPricePrev(UBound(strPricePrev)) = fncGetCurrMathFromDt(objKtbnStrc.strcSelection.strCurrency, objKtbnStrc.strcSelection.intListPrice.ToString, dt_CurrMath)
                                        ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                        strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                                    ElseIf objKtbnStrc.strcSelection.strMadeCountry.Equals("MDN") AndAlso _
                                           (objKtbnStrc.strcSelection.strSeriesKataban.Equals("LSH") AndAlso _
                                            objKtbnStrc.strcSelection.strKeyKataban.Equals("1")) Then
                                        If strCountryCd = "PRC" Then
                                            '現地定価の設定
                                            strPricePrev(UBound(strPricePrev)) = fncGetCurrMathFromDt("CNY", objKtbnStrc.strcSelection.intListPrice.ToString, dt_CurrMath)
                                            ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                            strCurrency(UBound(strCurrency)) = "CNY"
                                        Else
                                            'TOYO特殊対応    現地定価非表示
                                            strPricePrev(UBound(strPricePrev)) = 0
                                            ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                            strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                                        End If
                                    ElseIf objKtbnStrc.strcSelection.strMadeCountry = "MDN" AndAlso strCountryCd = "PRC" Then
                                        '出荷場所=MDN（金器）　かつ　ログインユーザーの国コード＝PRC     通貨：ログイン国の通貨　現地定価：定価       '2014/09/02 条件追加
                                        strPricePrev(UBound(strPricePrev)) = fncGetCurrMathFromDt(strCurrencyCd, objKtbnStrc.strcSelection.intListPrice.ToString, dt_CurrMath)
                                        ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                        strCurrency(UBound(strCurrency)) = strCurrencyCd     'kh_country_mstの通貨
                                    ElseIf objKtbnStrc.strcSelection.strMadeCountry.Equals("TYO") AndAlso _
                                           strCountryCd <> "PRC" AndAlso _
                                           (objKtbnStrc.strcSelection.strFullKataban.StartsWith("ETV") OrElse _
                                            objKtbnStrc.strcSelection.strFullKataban.StartsWith("ECS") OrElse _
                                            objKtbnStrc.strcSelection.strFullKataban.StartsWith("ECV") OrElse _
                                            objKtbnStrc.strcSelection.strFullKataban.StartsWith("ETS")) Then
                                        'TOYO特殊対応    現地定価非表示
                                        strPricePrev(UBound(strPricePrev)) = 0
                                        ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                        strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency

                                    ElseIf objKtbnStrc.strcSelection.strMadeCountry.Equals("CJA") Then
                                        'CJA 
                                        strPricePrev(UBound(strPricePrev)) = fncGetCurrMathFromDt(objKtbnStrc.strcSelection.strCurrency, _
                                                                       objKtbnStrc.strcSelection.intListPrice.ToString, dt_CurrMath)
                                        ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                        strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                                    Else
                                        '現地定価 Aprice = GS/PS/日本価格 * 掛率(list_price_rate1 or list_price_rate2) * 為替レート 
                                        'strCurrencyCd:ログイン国の通貨、strCurShipCd:画面で選択した生産国

                                        Dim dblRate As Decimal = 0D
                                        Call dalUnitPrice.fncSelectRateMstAprice(objConBase, strCurrencyCd, objKtbnStrc.strcSelection.strCurrency, dblRate)
                                        '現地定価の計算
                                        strPricePrev(UBound(strPricePrev)) = Me.fncAPriceGet(objConBase, CdCst.UnitPrice.APrice, _
                                                                                             objKtbnStrc, strCountryCd, _
                                                                                             objKtbnStrc.strcSelection.intListPrice, _
                                                                                             objKtbnStrc.strcSelection.intGsPrice, _
                                                                                             objKtbnStrc.strcSelection.intPsPrice, _
                                                                                             dblRate)
                                        ReDim Preserve strCurrency(UBound(strCurrency) + 1)

                                        strCurrency(UBound(strCurrency)) = strCurrencyCd     'kh_country_mstの通貨

                                    End If
                                End If
                                'FCA#1価格取得
                            Case CdCst.UnitPrice.FobPrice
                                ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                                Dim strViewCurr As String = String.Empty

                                '仕入品の場合は表示しない
                                'If objKtbnStrc.strcSelection.strKatabanCheckDiv = "5" Then
                                If objKtbnStrc.strcSelection.strDivision = "3" Then
                                    strPricePrev(UBound(strPricePrev)) = 0
                                    ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                    strCurrency(UBound(strCurrency)) = strViewCurr
                                    'FCA#1価格を格納　
                                    strPriceFCA = strPricePrev(UBound(strPricePrev))
                                Else
                                    '購入価格の計算
                                    strPricePrev(UBound(strPricePrev)) = Me.fncFobPriceGet(objConBase, CdCst.UnitPrice.FobPrice, _
                                                                                           objKtbnStrc, strCountryCd, objKtbnStrc.strcSelection.intGsPrice, _
                                                                                           strCurShipCd, strViewCurr)
                                    ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                    strCurrency(UBound(strCurrency)) = strViewCurr
                                    'FCA#1価格を格納
                                    strPriceFCA = strPricePrev(UBound(strPricePrev))
                                End If
                            Case CdCst.UnitPrice.CostPrice     '仕入価格
                                ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                                strPricePrev(UBound(strPricePrev)) = "0"
                                ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                                strCurrency(UBound(strCurrency)) = CdCst.CurrencyCd.DefaultCurrency
                        End Select
                    Next
            End Select

            ReDim strPriceList(UBound(strPricePrev), 4)

            For intLoopCnt = 1 To UBound(strPricePrev)
                strPriceList(intLoopCnt, 1) = strPrice(intLoopCnt, 2)
                strPriceList(intLoopCnt, 2) = strPricePrev(intLoopCnt)
                strPriceList(intLoopCnt, 3) = strCurrency(intLoopCnt)
                strPriceList(intLoopCnt, 4) = strPrice(intLoopCnt, 1)
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Sub

    ''' <summary>
    ''' ISO単価リスト取得処理
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objConBase"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessionId"></param>
    ''' <param name="strCountryCd">国コード</param>
    ''' <param name="strLanguageCd">言語コード</param>
    ''' <param name="strCurrencyCd">通貨コード</param>
    ''' <param name="intPriceDispLvl">表示区分</param>
    ''' <param name="strFullKataban">フル形番</param>
    ''' <param name="strCheckDiv">形番チェック区分</param>
    ''' <param name="intSglPrice">単価リスト(定価～ＰＳ)</param>
    ''' <param name="strCurShipCd">出荷場所</param>
    ''' <param name="strPriceList">表示単価リスト</param>
    ''' <remarks> ISO単価リストを取得し返却する</remarks>
    Public Sub subISOPriceListSelect(objCon As SqlConnection, objConBase As SqlConnection, objKtbnStrc As KHKtbnStrc, _
                                     ByVal strUserId As String, ByVal strSessionId As String, _
                                     ByVal strCountryCd As String, ByVal strLanguageCd As String, _
                                     ByVal strCurrencyCd As String, ByVal intPriceDispLvl As String, _
                                     ByVal strFullKataban As String, ByVal strCheckDiv As String, _
                                     ByVal intSglPrice() As Decimal, ByVal strCurShipCd As String, _
                                     ByRef strPriceList(,) As String)
        Dim strPrice(,) As String = Nothing
        Dim strPricePrev() As String
        Dim strCurrency() As String
        Dim intLoopCnt As Integer

        Try
            '価格リスト取得
            Call Me.subPriceListGet(objConBase, intPriceDispLvl, strLanguageCd, strPrice)

            '配列定義
            ReDim strPricePrev(0)
            ReDim strCurrency(0)

            For intLoopCnt = 1 To UBound(strPrice)
                Select Case strPrice(intLoopCnt, 1)
                    Case CdCst.UnitPrice.ListPrice
                        ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                        strPricePrev(UBound(strPricePrev)) = intSglPrice(0).ToString
                        ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                        strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                    Case CdCst.UnitPrice.RegPrice
                        ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                        strPricePrev(UBound(strPricePrev)) = intSglPrice(1).ToString
                        ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                        strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                    Case CdCst.UnitPrice.SsPrice
                        ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                        strPricePrev(UBound(strPricePrev)) = intSglPrice(2).ToString
                        ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                        strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                    Case CdCst.UnitPrice.BsPrice
                        ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                        strPricePrev(UBound(strPricePrev)) = intSglPrice(3).ToString
                        ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                        strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                    Case CdCst.UnitPrice.GsPrice
                        ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                        strPricePrev(UBound(strPricePrev)) = intSglPrice(4).ToString
                        ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                        strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                    Case CdCst.UnitPrice.PsPrice
                        ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                        strPricePrev(UBound(strPricePrev)) = intSglPrice(5).ToString
                        ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                        strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                    Case CdCst.UnitPrice.APrice
                        ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                        '為替レートを取得する  (生産国→ログイン国の為替レートを取得する)
                        '画面の出荷場所＝ログインﾕｰｻﾞｰの国コード＝PRC
                        If objKtbnStrc.strcSelection.strMadeCountry = strCountryCd And strCountryCd = "PRC" Then
                            strPricePrev(UBound(strPricePrev)) = fncGetCurrMath(objConBase, objKtbnStrc.strcSelection.strCurrency, _
                                                                                objKtbnStrc.strcSelection.intListPrice.ToString)
                            ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                            strCurrency(UBound(strCurrency)) = objKtbnStrc.strcSelection.strCurrency
                        Else
                            'strCurrencyCd:ログイン国の通貨、strCurShipCd:画面で選択した生産国
                            Dim dblRate As Decimal = 0D
                            dalUnitPrice.fncSelectRateMstAprice(objConBase, strCurrencyCd, objKtbnStrc.strcSelection.strCurrency, dblRate)

                            ''現地定価の計算
                            strPricePrev(UBound(strPricePrev)) = Me.fncAPriceGet(objConBase, CdCst.UnitPrice.APrice, _
                                                                                 objKtbnStrc, strCountryCd, _
                                                                                 intSglPrice(0), intSglPrice(4), _
                                                                                 intSglPrice(5), dblRate, strCheckDiv)
                            ReDim Preserve strCurrency(UBound(strCurrency) + 1)

                            strCurrency(UBound(strCurrency)) = strCurrencyCd     'kh_country_mstの通貨

                        End If
                    Case CdCst.UnitPrice.FobPrice
                        ReDim Preserve strPricePrev(UBound(strPricePrev) + 1)
                        Dim strViewCurr As String = String.Empty
                        '購入価格の計算
                        strPricePrev(UBound(strPricePrev)) = Me.fncFobPriceGet(objConBase, CdCst.UnitPrice.FobPrice, _
                                                             objKtbnStrc, strCountryCd, intSglPrice(4), strCurShipCd, strViewCurr)
                        ReDim Preserve strCurrency(UBound(strCurrency) + 1)
                        strCurrency(UBound(strCurrency)) = strViewCurr
                End Select
            Next

            '配列定義
            ReDim strPriceList(UBound(strPricePrev), 4)
            For intLoopCnt = 1 To UBound(strPricePrev)
                strPriceList(intLoopCnt, 1) = strPrice(intLoopCnt, 2)
                strPriceList(intLoopCnt, 2) = strPricePrev(intLoopCnt)
                strPriceList(intLoopCnt, 3) = strCurrency(intLoopCnt)
                strPriceList(intLoopCnt, 4) = strPrice(intLoopCnt, 1)
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 取引通貨マスタの取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetCurrMathAll(objConBase As SqlConnection) As DataTable
        fncGetCurrMathAll = New DataTable
        Try
            fncGetCurrMathAll = dalUnitPrice.fncSelectCurrMathAll(objConBase)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 取引通貨マスタより端数データを取得
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strCurrencyCd"></param>
    ''' <param name="decPrice"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetCurrMath(objConBase As SqlConnection, ByVal strCurrencyCd As String, ByVal decPrice As Decimal) As Decimal
        Dim dt As New DataTable

        fncGetCurrMath = decPrice
        Try
            dt = dalUnitPrice.fncSelectCurrMath(objConBase, strCurrencyCd)
            If dt.Rows.Count > 0 Then
                '端数データの計算
                fncGetCurrMath = CDec(subFractionProc(decPrice, dt.Rows(0)("math_Type"), dt.Rows(0)("math_Pos")))
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 取引通貨マスタの取得(Datatable)
    ''' </summary>
    ''' <param name="strCurrencyCd"></param>
    ''' <param name="decPrice"></param>
    ''' <param name="dt_CurrMath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncGetCurrMathFromDt(ByVal strCurrencyCd As String, ByVal decPrice As Decimal, _
                                        ByVal dt_CurrMath As DataTable) As Decimal
        fncGetCurrMathFromDt = decPrice
        Try
            Dim dr() As DataRow = dt_CurrMath.Select("currency_cd ='" & strCurrencyCd & "'")
            If dr.Length > 0 Then
                Dim intType As Integer = dr(0)("math_Type")
                Dim dblPos As Double = dr(0)("math_Pos")
                fncGetCurrMathFromDt = CDec(subFractionProc(decPrice, intType, dblPos))
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 端数処理区分取得
    ''' </summary>
    ''' <param name="decPrice">端数処理対象数値</param>
    ''' <param name="strMathType">計算タイプ（0：なし、1：四捨五入、2：切上げ、3：切捨て）</param>
    ''' <param name="intMathPos">計算位置（例)1：なし、10：整数一位、0.1：少数一位</param>
    ''' <returns></returns>
    ''' <remarks>拠点コードを元に端数処理をする</remarks>
    Public Function subFractionProc(ByVal decPrice As Decimal, ByVal strMathType As String, _
                                     ByVal intMathPos As Decimal) As String
        subFractionProc = "0"
        Try
            Select Case strMathType
                Case "1"
                    '四捨五入(丸め)
                    subFractionProc = (Math.Round(decPrice * intMathPos) / intMathPos).ToString
                Case "2"
                    '切上げ
                    subFractionProc = (Math.Ceiling(decPrice * intMathPos) / intMathPos).ToString
                Case "3"
                    '切捨て
                    subFractionProc = (Math.Truncate(decPrice * intMathPos) / intMathPos).ToString
                Case "4"
                    '四捨五入
                    'subFractionProc = (Math.Round(decPrice, intMathPos.ToString.Length - 1, MidpointRounding.AwayFromZero)).ToString
                    If intMathPos < 1 Then
                        subFractionProc = (Math.Round(decPrice * intMathPos, 0, MidpointRounding.AwayFromZero) / intMathPos).ToString
                    Else
                        subFractionProc = (Math.Round(decPrice, intMathPos.ToString.Length - 1, MidpointRounding.AwayFromZero)).ToString
                    End If
            End Select

            If intMathPos > 1 Then
                Dim str() As String = subFractionProc.Split(".")
                If str.Length = 2 Then '小数あり
                    subFractionProc = str(0) & "." & str(1).PadRight(intMathPos.ToString.Length - 1, "0")
                ElseIf str.Length = 1 Then
                    subFractionProc = str(0) & "." & "".PadRight(intMathPos.ToString.Length - 1, "0")
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' 価格リスト取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="intPriceDispLvl">価格表示レベル</param>
    ''' <param name="strLanguageCd">言語コード</param>
    ''' <param name="strPrice">単価リスト</param>
    ''' <remarks>価格表示レベルを元に価格リストを取得する</remarks>
    Private Sub subPriceListGet(objCon As SqlConnection, ByVal intPriceDispLvl As Integer, _
                                ByVal strLanguageCd As String, ByRef strPrice(,) As String)
        Dim intPriceLvl As Integer
        Dim strPriceDiv() As String
        Dim strPriceNm() As String
        Dim intDispSeq() As Integer
        Dim intLoopCnt As Integer
        Dim dt As New DataTable

        Try
            '配列初期化
            ReDim strPriceDiv(0)
            ReDim strPriceNm(0)
            ReDim intDispSeq(0)

            '価格表示レベル設定
            intPriceLvl = intPriceDispLvl

            '表示区分取得
            dt = dalUnitPrice.fncSelectDispLvl(objCon, intPriceDispLvl, strLanguageCd, strPrice)

            For Each dr In dt.Rows
                If intPriceLvl >= dr("price_lvl") Then
                    '価格レベル計算
                    intPriceLvl = intPriceLvl - dr("price_lvl")
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strPriceDiv(UBound(strPriceDiv)) = dr("price_div")
                    ReDim Preserve strPriceNm(UBound(strPriceNm) + 1)
                    strPriceNm(UBound(strPriceNm)) = dr("price_nm")
                    ReDim Preserve intDispSeq(UBound(intDispSeq) + 1)
                    intDispSeq(UBound(intDispSeq)) = dr("disp_seq_no")
                End If
            Next

            '配列を逆にする(価格区分順にソートする)
            Array.Reverse(strPriceDiv)
            Array.Reverse(strPriceNm)
            Array.Reverse(intDispSeq)

            '配列初期化
            ReDim strPrice(UBound(strPriceDiv), 3)
            For intLoopCnt = 1 To strPriceDiv.Length - 1
                strPrice(intLoopCnt, 1) = strPriceDiv(intLoopCnt - 1)
                strPrice(intLoopCnt, 2) = strPriceNm(intLoopCnt - 1)
                strPrice(intLoopCnt, 3) = intDispSeq(intLoopCnt - 1).ToString
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 現地定価の計算
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strPriceDiv"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strCheckDiv"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncAPriceGet(objConBase As SqlConnection, ByVal strPriceDiv As String, _
                                  objKtbnStrc As KHKtbnStrc, ByVal strCountryCd As String, _
                                  ByVal intListPrice As Decimal, ByVal intGsPrice As Decimal, _
                                  ByVal intPsPrice As Decimal, ByVal dblRate As Decimal, _
                                  Optional strCheckDiv As String = "") As String

        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objRdr As SqlDataReader = Nothing
        Dim objKataban As New KHKataban
        Dim strSeries As String

        fncAPriceGet = "0"

        Try
            Dim strKataban As String = objKtbnStrc.strcSelection.strFullKataban
            Dim strKatabanCheckDiv As String = String.Empty
            If strCheckDiv = String.Empty Then
                strKatabanCheckDiv = objKtbnStrc.strcSelection.strKatabanCheckDiv
            Else
                strKatabanCheckDiv = strCheckDiv
            End If
            Dim dtRate As New DataTable

            '機種形番取得
            strSeries = KHKataban.fncMdlKtbnGet(strKataban)

            '掛け率の取得
            dtRate = dalUnitPrice.fncSelectRateAprice(objConBase, strCountryCd, strSeries)

            If dtRate.Rows.Count > 0 Then
                Dim strTypeA As String = dtRate.Rows(0)("TypeA").ToString.Trim
                Dim strPosA As String = dtRate.Rows(0)("PosA").ToString.Trim

                'Dim dblRate As Decimal = 0D                                                '為替レート
                ''為替レートの取得
                'dalUnitPrice.fncSelectRateMstAprice(objConBase, strCurrencyCd, objKtbnStrc.strcSelection.strMadeCountry, dblRate)

                '現地定価 Aprice = GS/PS/日本価格 * 掛率(list_price_rate1 or list_price_rate2) * 為替レート 
                Select Case strCountryCd
                    Case "USA", "MEX", "E09"                         '欧州代理店明治対応 RM1705008  2017/05/11 更新
                        'Case "USA", "MEX"                           'メキシコ対応  RM1509001 
                        fncAPriceGet = subFractionProc(intListPrice * dtRate.Rows(0)("list_price_rate1") * dblRate, _
                                            strTypeA, strPosA).ToString
                    Case "PRC"
                        Select Case strKatabanCheckDiv
                            Case CdCst.KatabanChackDiv.Parts
                                'RM1806044_フル形番検索時、現地定価表示制御変更
                                If dtRate.Rows(0)("list_price_rate2") = 0 Then
                                    fncAPriceGet = subFractionProc(intGsPrice * dtRate.Rows(0)("list_price_rate1") * dblRate, _
                                                        strTypeA, strPosA).ToString
                                Else
                                    fncAPriceGet = subFractionProc(intGsPrice * dtRate.Rows(0)("list_price_rate2") * dblRate, _
                                                        strTypeA, strPosA).ToString
                                End If
                            Case Else
                                fncAPriceGet = subFractionProc(intGsPrice * dtRate.Rows(0)("list_price_rate1") * dblRate, _
                                                    strTypeA, strPosA).ToString
                        End Select
                    Case Else
                        fncAPriceGet = subFractionProc(intGsPrice * dtRate.Rows(0)("list_price_rate1") * dblRate, _
                                                strTypeA, strPosA).ToString
                End Select
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
        End Try
    End Function

    ''' <summary>
    ''' 購入価格の計算
    ''' </summary>
    ''' <param name="objConBase"></param>
    ''' <param name="strPriceDiv"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strSelCountry"></param>
    ''' <param name="strViewCurr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncFobPriceGet(objConBase As SqlConnection, ByVal strPriceDiv As String, objKtbnStrc As KHKtbnStrc, _
                                    ByVal strCountryCd As String, ByVal intGsPrice As Decimal, ByVal strSelCountry As String, _
                                    ByRef strViewCurr As String) As String
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objRdr As SqlDataReader = Nothing
        Dim objKataban As New KHKataban
        Dim strSeries As String
        Dim dtRate As New DataTable
        fncFobPriceGet = "0"

        Try
            Dim strKataban As String = objKtbnStrc.strcSelection.strFullKataban
            'Dim intGsPrice As Decimal = objKtbnStrc.strcSelection.intGsPrice
            Dim strMadeCountry As String = objKtbnStrc.strcSelection.strMadeCountry
            Dim strBaseCurr As String = objKtbnStrc.strcSelection.strCurrency

            '機種形番取得
            strSeries = KHKataban.fncMdlKtbnGet(strKataban)
            '掛け率の取得
            dtRate = dalUnitPrice.fncSelectRateFobprice(objConBase, strCountryCd, strSeries, strSelCountry)

            If dtRate.Rows.Count > 0 Then
                strViewCurr = dtRate.Rows(0)("currency_cd")    '変更通貨を取得する
                Dim strTypeFOB As String = dtRate.Rows(0)("TypeFOB").ToString.Trim
                Dim strPosFOB As String = dtRate.Rows(0)("PosFOB").ToString.Trim
                Dim decFOBRate As Decimal = dtRate.Rows(0)("fob_rate")
                Dim dblRate As Decimal = 0D

                '為替レートを取得する（GS基準通貨→変更後通貨）
                dalUnitPrice.fncGetRateMst(objConBase, strBaseCurr, strViewCurr, dblRate)

                '端数処理
                '購入定価 Fobprice = GS価格 * 掛率(fob_rate) * 為替レート 
                fncFobPriceGet = subFractionProc(intGsPrice * decFOBRate * dblRate, strTypeFOB, strPosFOB).ToString

                '特価決裁No
                objKtbnStrc.strcSelection.strAuthorizationNo = dtRate.Rows(0)("authorization_no")

            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objRdr Is Nothing Then If Not objRdr.IsClosed Then objRdr.Close()
            objRdr = Nothing
            sbSql = Nothing
        End Try

    End Function

    ''' <summary>
    '''  初期化されなかったものを初期化する
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <remarks></remarks>
    Private Sub subInitObjKtbnstrc(ByRef objKtbnStrc As KHKtbnStrc)
        With objKtbnStrc.strcSelection
            If .strRodEndOption Is Nothing Then
                .strRodEndOption = ""
            End If
            If .strOtherOption Is Nothing Then
                .strOtherOption = ""
            End If
            If .strPositionOption Is Nothing Then
                .strPositionOption = ""
            End If
        End With
    End Sub

    ''' <summary>
    ''' 価格表示情報を取得
    ''' </summary>
    ''' <param name="intPriceDispLvl">利用機能レベル</param>
    ''' <returns></returns>
    ''' <remarks>価格表示レベルをもとに価格表示情報を取得する</remarks>
    Public Function fncPriceDispLvlInfoGet(ByVal intPriceDispLvl As Integer) As Boolean()
        Dim intPriceDispInfo(7) As Boolean
        Dim intWkPriceDispLvl As Integer = intPriceDispLvl
        fncPriceDispLvlInfoGet = Nothing

        Try
            ''仕入価格
            'If intWkPriceDispLvl >= strcPriceDispLvl.CostPrice Then
            '    intPriceDispInfo(0) = True
            '    intWkPriceDispLvl = intWkPriceDispLvl - strcPriceDispLvl.CostPrice
            'Else
            '    intPriceDispInfo(0) = False
            'End If
            '購入価格
            If intWkPriceDispLvl >= strcPriceDispLvl.FobPrice Then
                intPriceDispInfo(0) = True
                intWkPriceDispLvl = intWkPriceDispLvl - strcPriceDispLvl.FobPrice
            Else
                intPriceDispInfo(0) = False
            End If
            '現地定価
            If intWkPriceDispLvl >= strcPriceDispLvl.APrice Then
                intPriceDispInfo(1) = True
                intWkPriceDispLvl = intWkPriceDispLvl - strcPriceDispLvl.APrice
            Else
                intPriceDispInfo(1) = False
            End If
            'PS
            If intWkPriceDispLvl >= strcPriceDispLvl.PsPrice Then
                intPriceDispInfo(2) = True
                intWkPriceDispLvl = intWkPriceDispLvl - strcPriceDispLvl.PsPrice
            Else
                intPriceDispInfo(2) = False
            End If
            'GS
            If intWkPriceDispLvl >= strcPriceDispLvl.GsPrice Then
                intPriceDispInfo(3) = True
                intWkPriceDispLvl = intWkPriceDispLvl - strcPriceDispLvl.GsPrice
            Else
                intPriceDispInfo(3) = False
            End If
            'BS
            If intWkPriceDispLvl >= strcPriceDispLvl.BsPrice Then
                intPriceDispInfo(4) = True
                intWkPriceDispLvl = intWkPriceDispLvl - strcPriceDispLvl.BsPrice
            Else
                intPriceDispInfo(4) = False
            End If
            'SS
            If intWkPriceDispLvl >= strcPriceDispLvl.SsPrice Then
                intPriceDispInfo(5) = True
                intWkPriceDispLvl = intWkPriceDispLvl - strcPriceDispLvl.SsPrice
            Else
                intPriceDispInfo(5) = False
            End If
            '登録店
            If intWkPriceDispLvl >= strcPriceDispLvl.RegPrice Then
                intPriceDispInfo(6) = True
                intWkPriceDispLvl = intWkPriceDispLvl - strcPriceDispLvl.RegPrice
            Else
                intPriceDispInfo(6) = False
            End If
            '定価
            If intWkPriceDispLvl >= strcPriceDispLvl.ListPrice Then
                intPriceDispInfo(7) = True
                intWkPriceDispLvl = intWkPriceDispLvl - strcPriceDispLvl.ListPrice
            Else
                intPriceDispInfo(7) = False
            End If

            fncPriceDispLvlInfoGet = intPriceDispInfo

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' マニホールドテスト
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="htPriceInfo"></param>
    ''' <param name="htScrewPriceInfo"></param>
    ''' <param name="DS_Tab"></param>
    ''' <param name="strKatabanCheckDiv"></param>
    ''' <param name="strPlaceCd"></param>
    ''' <param name="strCountryCd"></param>
    ''' <param name="strMadeCountry"></param>
    ''' <param name="strCurrency"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncManifoldTest(ByVal objKtbnStrc As KHKtbnStrc, ByRef htPriceInfo As Hashtable, _
                                ByRef htScrewPriceInfo As Hashtable, ByRef DS_Tab As DataSet, _
                                ByRef strKatabanCheckDiv As String, ByRef strPlaceCd As String, _
                                ByRef strCountryCd As String, ByRef strMadeCountry As String, _
                                ByRef strCurrency As String) As Boolean
        'Define
        Dim bolReturn As Boolean = False
        Dim dr() As DataRow = Nothing
        Dim dt_fullPrice As New DS_KatOut.kh_priceDataTable
        Try
            '初期化
            dt_fullPrice = DS_Tab.Tables("dt_fullPrice")
            dr = dt_fullPrice.Select("kataban='" & objKtbnStrc.strcSelection.strFullKataban & "'")
            htPriceInfo = New Hashtable
            htScrewPriceInfo = New Hashtable
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
                strCurrency = dr(0)("currency_cd")
                strMadeCountry = dr(0)("country_cd")
                bolReturn = True
            Else
                Dim dt_accPrice As New DS_KatOut.kh_accumulate_priceDataTable
                dt_accPrice = DS_Tab.Tables("dt_accPrice")
                dr = dt_accPrice.Select("kataban='" & objKtbnStrc.strcSelection.strFullKataban & "'")
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
                End If
            End If

            Dim dt_screPrice As New DS_KatOut.kh_screw_kataban_mstDataTable
            dt_screPrice = DS_Tab.Tables("dt_screPrice")
            dr = dt_screPrice.Select("kataban='" & objKtbnStrc.strcSelection.strFullKataban & "'")

            '初期化
            htScrewPriceInfo(CdCst.UnitPrice.ListPrice) = 0
            htScrewPriceInfo(CdCst.UnitPrice.RegPrice) = 0
            htScrewPriceInfo(CdCst.UnitPrice.SsPrice) = 0
            htScrewPriceInfo(CdCst.UnitPrice.BsPrice) = 0
            htScrewPriceInfo(CdCst.UnitPrice.GsPrice) = 0
            htScrewPriceInfo(CdCst.UnitPrice.PsPrice) = 0
            If dr.Length > 0 Then
                If strCountryCd <> CdCst.CountryCd.DefaultCountry Then '海外のみ
                    htScrewPriceInfo(CdCst.UnitPrice.ListPrice) = dr(0)("ls_price")
                    htScrewPriceInfo(CdCst.UnitPrice.RegPrice) = dr(0)("rg_price")
                    htScrewPriceInfo(CdCst.UnitPrice.SsPrice) = dr(0)("ss_price")
                    htScrewPriceInfo(CdCst.UnitPrice.BsPrice) = dr(0)("bs_price")
                    htScrewPriceInfo(CdCst.UnitPrice.GsPrice) = dr(0)("gs_price")
                    htScrewPriceInfo(CdCst.UnitPrice.PsPrice) = dr(0)("ps_price")
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
        Return bolReturn
    End Function

End Class
