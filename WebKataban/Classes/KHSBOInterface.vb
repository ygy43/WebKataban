Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class KHSBOInterface
    Private Structure HeaderInfo
        Public SystemDatetime As String                     'システム日付
        Public Kataban1 As String                           '受注形番1
        Public Kataban2 As String                           '受注形番2

        Public ListPrice As String                          '定価
        Public RegPrice As String                           '登録店価格
        Public SsPrice As String                            'SS店価格
        Public BsPrice As String                            'BS価格
        Public GsPrice As String                            'GS価格
        Public PsPrice As String                            'PS価格
        Public NetPrice As String                           '購入価格
        Public CurrencyCode As String                       '通貨コード

        Public SpecExistsDiv As String                      '仕様書有無区分
        Public ModelCd As String                            '機種コード
        Public WiringSpecDiv As String                      '配線仕様有無区分
        Public RailLength As String                         'レール長さ

        Public ProcDatetime As String                       '処理日
        Public FullKataban As String                        'マニホールド代表形番

        Public KatabanCheckDiv As String                    '形番チェック区分
        Public PlaceCd As String                            '出荷場所
        Public ELDiv As String                              'EL判定区分
        Public MsgPosition As String                        '位置情報
        Public CZFlag As String                             '特価フラグ
        Public Quantity As Integer                          '数量
        Public kataban As String                            '表示形番

    End Structure
    Private Shared strcHeader As HeaderInfo

    Private Structure AccessoryInfo
        Public AttributeSymbol As String                    '属性記号
        Public OptionKataban As String                      'オプション形番
        Public Quantity As String                           '使用数
    End Structure
    Private Shared strcAccessoryInfo() As AccessoryInfo

    Private Structure ManifoldInfo
        Public AttributeSymbol As String                    '属性記号
        Public OptionKataban As String                      'オプション形番
        Public PositionInfo As String                       '設置位置
        Public Quantity As String                           '使用数
        Public OrderNo As String                            '受注No.
    End Structure
    Private Shared strcManifoldInfo() As ManifoldInfo

    'Public clKatahikiInfoDto As New WebKataban.CommonDbService.KatahikiInfoDto
    'Public clKatahikiInfoDtoIso As New System.Collections.Generic.List(Of CommonDbService.KatahikiInfoDto)

    'RM1803032　追加
    Public Shared intLoopMax_01 As Integer = 30     'オプション形番部分行数
    Public Shared intLoopMax_02 As Integer = 15     '付属品部分行数

    Public Shared intSpaceCnt_01 As Integer = 1     '配線仕様書有無区分
    Public Shared intSpaceCnt_02 As Integer = 6     '取付ﾚｰﾙ長さ
    Public Shared intSpaceCnt_03 As Integer = 2     '属性記号
    Public Shared intSpaceCnt_04 As Integer = 30    '形番
    Public Shared intSpaceCnt_05 As Integer = 40    '設置位置
    Public Shared intSpaceCnt_06 As Integer = 2     '使用数量

    ''' <summary>
    ''' SBOインターフェース情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strUserID"></param>
    ''' <param name="strSessID"></param>
    ''' <param name="strPlaceCD">出荷場所（中国生産品対応)</param>
    ''' <returns></returns>
    ''' <remarks>SBOにインターフェースする情報を編集し返却する</remarks>
    Public Shared Function fncSBOInterfaceGet(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, _
                                              ByVal strFobPrice As String, ByVal strCountryCd As String, _
                                              ByVal strOfficeCd As String, strUserID As String, _
                                              strSessID As String, Optional ByVal strPlaceCD As String = "") As String

        Dim objOption As New KHOptionCtl
        Dim objKataban As New KHKataban
        Dim sbBuilder As New System.Text.StringBuilder(2737)

        Dim strOpArray() As String
        Dim strPositionInfo As String
        Dim intLoopCnt As Integer
        Dim intLoopCnt1 As Integer
        Dim intIndex As Integer = 0

        Dim strAccAttributeSymbol() As String = Nothing
        Dim strAccOptionKataban() As String = Nothing
        Dim strAccPositionInfo() As String = Nothing
        Dim intAccQuantity() As Integer = Nothing

        Dim strTmpKataban As String = ""
        Dim strTmpPositionInfo As String = ""
        Dim strTmpPositionInfo1 As String = ""
        Dim strTmpPositionInfo2 As String = ""
        Dim intTmpQuantity As Integer

        Dim strPos As String
        'Dim intLoopPos As Integer
        'Dim intRenPos As Integer
        'Dim intMPos As Integer
        'Dim intMPos2 As Integer
        Dim intPositionInfo() As Integer
        Dim intLoopPos2 As Integer

        Try
            fncSBOInterfaceGet = ""
            ReDim strcManifoldInfo(30)
            ReDim strcAccessoryInfo(15)

            '仕様書情報
            Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                Case "05", "06"
                    Dim intCount As Integer = 0
                    For intLoopCnt1 = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 2
                        If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1).Trim <> "" And _
                           objKtbnStrc.strcSelection.intQuantity(intLoopCnt1) <> 0 Then
                            '編集
                            With strcHeader
                                'システム日付
                                .SystemDatetime = Now.ToString("yyyyMMddHHmmss", New Globalization.CultureInfo("en-us"))

                                '受注形番1,受注形番2
                                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1).Trim.Length = 30 Then
                                    .Kataban1 = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1).Trim
                                    .Kataban2 = Space(30)
                                ElseIf objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1).Trim.Length < 30 Then
                                    .Kataban1 = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1).Trim & _
                                                Space(30 - objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1).Trim.Length)
                                    .Kataban2 = Space(30)
                                ElseIf objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1).Trim.Length > 30 Then
                                    .Kataban1 = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1).Trim, 30)
                                    .Kataban2 = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1).Trim, 31) & _
                                                Space(60 - objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1).Trim.Length)
                                End If

                                'PS価格
                                'FOB対応
                                If strCountryCd = "JPN" Then
                                    .PsPrice = Format(objKtbnStrc.strcSelection.intOpPsPrice(intLoopCnt1), "000000000")
                                Else

                                    Dim lstFobPrice As List(Of String) = strFobPrice.Split(",").ToList

                                    If lstFobPrice.Count - 1 >= intCount Then

                                        Dim decFobPrice As Decimal

                                        decFobPrice = IIf(Decimal.TryParse(lstFobPrice(intCount), decFobPrice), decFobPrice, 0)

                                        If decFobPrice - Fix(decFobPrice) = 0 Then
                                            .PsPrice = Format(decFobPrice, "000000000")
                                            intCount = intCount + 1
                                        Else
                                            .PsPrice = Format(decFobPrice, "000000.00")
                                            intCount = intCount + 1
                                        End If
                                    Else
                                        .PsPrice = Format(0, "000000000")
                                        intCount = intCount + 1
                                    End If

                                End If

                                'GS価格
                                .GsPrice = Format(objKtbnStrc.strcSelection.intOpGsPrice(intLoopCnt1), "000000000")
                                '定価
                                .ListPrice = Format(objKtbnStrc.strcSelection.intOpListPrice(intLoopCnt1), "000000000")
                                '仕様書有無区分
                                .SpecExistsDiv = "Y"
                                '機種コード
                                .ModelCd = Left(objKtbnStrc.strcSelection.strModelNo.Trim & Space(2), 2)
                                '配線仕様有無区分
                                .WiringSpecDiv = Left(objKtbnStrc.strcSelection.strWiringSpec.Trim & Space(1), 1)
                                'レール長さ
                                .RailLength = Format(objKtbnStrc.strcSelection.decDinRailLength * 100, "000000")
                                '処理日付＆マニホールド代表形番
                                .ProcDatetime = Format(Now, "MMddhhmmss")
                                .FullKataban = Left(objKtbnStrc.strcSelection.strFullKataban.Trim & Space(30), 30)
                                '形番チェック区分
                                .KatabanCheckDiv = "Z" & Left(objKtbnStrc.strcSelection.strOpKatabanCheckDiv(intLoopCnt1).Trim & Space(1), 1)
                                '出荷場所
                                If strPlaceCD Is Nothing OrElse strPlaceCD.Length <= 0 Then
                                    .PlaceCd = Left(objKtbnStrc.strcSelection.strOpPlaceCd(intLoopCnt1).Trim & Space(4), 4)
                                Else
                                    .PlaceCd = Left(strPlaceCD & Space(4), 4)
                                End If
                                strPos = ""
                                .MsgPosition = Left(strPos.Trim & Space(60), 60)
                                'EL品判定区分
                                If objKataban.fncELKatabanCheck(objCon, objKtbnStrc.strcSelection.strFullKataban.Trim, "1") Then
                                    .ELDiv = CdCst.ELDiv.Yes
                                Else
                                    .ELDiv = CdCst.ELDiv.No
                                End If
                            End With
                            '仕様書情報
                            Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                                Case "05", "06"
                                    '初期化
                                    For intLoopCnt = 1 To 20
                                        With strcManifoldInfo(intLoopCnt)
                                            .AttributeSymbol = Space(2)
                                            .OptionKataban = Space(30)
                                            .PositionInfo = Space(10)
                                            .Quantity = "00"
                                            .OrderNo = Space(8)
                                        End With
                                    Next
                                    For intLoopCnt = 1 To 10
                                        With strcAccessoryInfo(intLoopCnt)
                                            .AttributeSymbol = Space(2)
                                            .OptionKataban = Space(30)
                                            .Quantity = "00"
                                        End With
                                    Next
                                    '設定
                                    Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                                        Case "05"
                                            intIndex = 0
                                            For intLoopCnt = 1 To 25
                                                If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                                                    intIndex = intIndex + 1
                                                    strcManifoldInfo(intIndex).AttributeSymbol = Left(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim & Space(2), 2)
                                                    strcManifoldInfo(intIndex).OptionKataban = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & Space(30), 30)
                                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                                    strPositionInfo = Replace(strPositionInfo, ",", "")
                                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
                                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                                    strcManifoldInfo(intIndex).PositionInfo = Left(strPositionInfo & Space(10), 10)
                                                    strcManifoldInfo(intIndex).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                                End If
                                            Next
                                            '付属品取得
                                            Call subISOAccessoryGet(objKtbnStrc, strAccAttributeSymbol, strAccOptionKataban, intAccQuantity)
                                            intIndex = 0
                                            For intLoopCnt = 1 To strAccAttributeSymbol.Length - 1
                                                intIndex = intIndex + 1
                                                strcAccessoryInfo(intIndex).AttributeSymbol = Left(strAccAttributeSymbol(intLoopCnt).Trim & Space(2), 2)
                                                strcAccessoryInfo(intIndex).OptionKataban = Left(strAccOptionKataban(intLoopCnt).Trim & Space(30), 30)
                                                strcAccessoryInfo(intIndex).Quantity = Format(intAccQuantity(intLoopCnt), "00")
                                            Next
                                        Case "06"
                                            intIndex = 0
                                            For intLoopCnt = 1 To 19
                                                If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                                                    intIndex = intIndex + 1
                                                    strcManifoldInfo(intIndex).AttributeSymbol = Left(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim & Space(2), 2)
                                                    strcManifoldInfo(intIndex).OptionKataban = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & Space(30), 30)
                                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                                    strPositionInfo = Replace(strPositionInfo, ",", "")
                                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
                                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                                    'ADD BY YGY 20141016    逆の時NET版と一致するように    ↓↓↓↓↓↓
                                                    If objKtbnStrc.strcSelection.strSeriesKataban = "LMF0" AndAlso _
                                                        objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "T0D" Then
                                                        strPositionInfo = StrReverse(strPositionInfo)
                                                    End If
                                                    'ADD BY YGY 20141016    逆の時NET版と一致するように    ↑↑↑↑↑↑
                                                    strcManifoldInfo(intIndex).PositionInfo = Left(strPositionInfo & Space(10), 10)
                                                    strcManifoldInfo(intIndex).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                                    'End If
                                                End If
                                            Next
                                            '付属品取得
                                            Call subISOAccessoryGet(objKtbnStrc, strAccAttributeSymbol, strAccOptionKataban, intAccQuantity)
                                            intIndex = 0
                                            For intLoopCnt = 1 To strAccAttributeSymbol.Length - 1
                                                intIndex = intIndex + 1
                                                strcAccessoryInfo(intIndex).AttributeSymbol = Left(strAccAttributeSymbol(intLoopCnt).Trim & Space(2), 2)
                                                strcAccessoryInfo(intIndex).OptionKataban = Left(strAccOptionKataban(intLoopCnt).Trim & Space(30), 30)
                                                strcAccessoryInfo(intIndex).Quantity = Format(intAccQuantity(intLoopCnt), "00")
                                            Next
                                    End Select
                            End Select

                            '文字列結合
                            With sbBuilder
                                '.Append(strcHeader.SystemDatetime)
                                '.Append(strcHeader.Kataban1)
                                '.Append(strcHeader.Kataban2)
                                '.Append(strcHeader.PsPrice)
                                '.Append(strcHeader.GsPrice)
                                '.Append(strcHeader.ListPrice)
                                '.Append(strcHeader.SpecExistsDiv)
                                '.Append(strcHeader.ModelCd)
                                .Append(strcHeader.WiringSpecDiv)
                                .Append(strcHeader.RailLength)
                                '.Append(strcHeader.ProcDatetime)
                                .Append(strcHeader.FullKataban)
                                For intLoopCnt = 1 To 20
                                    .Append(strcManifoldInfo(intLoopCnt).AttributeSymbol)
                                    .Append(strcManifoldInfo(intLoopCnt).OptionKataban)
                                    .Append(strcManifoldInfo(intLoopCnt).PositionInfo)
                                    .Append(strcManifoldInfo(intLoopCnt).Quantity)
                                    .Append(strcManifoldInfo(intLoopCnt).OrderNo)
                                Next
                                For intLoopCnt = 1 To 10
                                    .Append(strcAccessoryInfo(intLoopCnt).AttributeSymbol)
                                    .Append(strcAccessoryInfo(intLoopCnt).OptionKataban)
                                    .Append(strcAccessoryInfo(intLoopCnt).Quantity)
                                Next
                                '.Append(strcHeader.KatabanCheckDiv)
                                '.Append(strcHeader.PlaceCd)
                                '.Append(strcHeader.ELDiv)
                                '.Append(strcHeader.MsgPosition)
                                .Append(vbCrLf)
                            End With
                        End If
                    Next
                Case Else
                    '編集
                    With strcHeader

                        'RM1803032ヘッダー部分削除

                        ''システム日付
                        '.SystemDatetime = Now.ToString("yyyyMMddHHmmss", New Globalization.CultureInfo("en-us"))

                        ''受注形番1,受注形番2
                        'If objKtbnStrc.strcSelection.strFullKataban.Trim.Length = 30 Then
                        '    .Kataban1 = objKtbnStrc.strcSelection.strFullKataban.Trim
                        '    .Kataban2 = Space(30)
                        'ElseIf objKtbnStrc.strcSelection.strFullKataban.Trim.Length < 30 Then
                        '    .Kataban1 = objKtbnStrc.strcSelection.strFullKataban.Trim & _
                        '                Space(30 - objKtbnStrc.strcSelection.strFullKataban.Trim.Length)
                        '    .Kataban2 = Space(30)
                        'ElseIf objKtbnStrc.strcSelection.strFullKataban.Trim.Length > 30 Then
                        '    .Kataban1 = Left(objKtbnStrc.strcSelection.strFullKataban.Trim, 30)
                        '    .Kataban2 = Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, 31) & _
                        '                Space(60 - objKtbnStrc.strcSelection.strFullKataban.Trim.Length)
                        'End If

                        ''PS価格
                        ''.PsPrice = Format(Me.strcSelection.intPsPrice, "000000000")
                        ''.PsPrice = Format(Integer.Parse(strFobPrice), "000000000")
                        ''FOB対応
                        'If strCountryCd = "JPN" And strOfficeCd <> "II2" Then
                        '    .PsPrice = Format(objKtbnStrc.strcSelection.intPsPrice, "000000000")
                        'Else
                        '    Dim decFobPrice As Decimal
                        '    decFobPrice = IIf(Decimal.TryParse(strFobPrice, decFobPrice), decFobPrice, 0)
                        '    If decFobPrice - Fix(decFobPrice) = 0 Then
                        '        .PsPrice = Format(decFobPrice, "000000000")
                        '    Else
                        '        .PsPrice = Format(decFobPrice, "000000.00")
                        '    End If
                        '    '.PsPrice = Format(Decimal.Parse(Replace(strFobPrice, Right(strFobPrice, 3), "")), "000000.00") & Right(strFobPrice, 3)
                        'End If
                        ''GS価格
                        '.GsPrice = Format(objKtbnStrc.strcSelection.intGsPrice, "000000000")
                        ''定価
                        '.ListPrice = Format(objKtbnStrc.strcSelection.intListPrice, "000000000")
                        ''仕様書有無区分
                        'Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                        '    Case "", "00"
                        '        .SpecExistsDiv = "N"
                        '        strPos = ""
                        '        .MsgPosition = Left(strPos.Trim & Space(60), 60)
                        '    Case "51", "52", "53", "54", "55", "56", "57", "58", "59", "60", _
                        '         "61", "62", "63", "64", "65", "66", "67", "68", "69", "70", _
                        '         "71", "72", "73", "74", "75", "76", "77", "78", "79", "80", _
                        '         "81", "82", "83", "84", "85", "86", "87", "88", "89", "90", _
                        '         "91", "92", "93", "A4", "A5", "A6", "A7", "A8", "98", _
                        '         "S", "T"
                        '        '"91", "92", "93", "A1", "A2", "42", "A7", "A8", "A4", "A5", "A6", "A9", "B1"

                        '        If objOption.fncVaccumMixCheck(objKtbnStrc) Then
                        '            .SpecExistsDiv = "Y"
                        '        Else
                        '            .SpecExistsDiv = "N"
                        '        End If
                        '        If objKtbnStrc.strcSelection.strSpecNo.Trim = "A1" Or objKtbnStrc.strcSelection.strSpecNo.Trim = "A2" _
                        '           Or objKtbnStrc.strcSelection.strSpecNo.Trim = "A9" Or objKtbnStrc.strcSelection.strSpecNo.Trim = "B1" _
                        '           Or objKtbnStrc.strcSelection.strSpecNo.Trim = "B2" Or objKtbnStrc.strcSelection.strSpecNo.Trim = "B3" _
                        '           Or objKtbnStrc.strcSelection.strSpecNo.Trim = "B4" Then
                        '            .SpecExistsDiv = "Y"
                        '        End If
                        '        strPos = ""
                        '        '簡易マニホールドの判断
                        '        If KHKataban.fncJudgeSimpleSpec(objCon, objKtbnStrc, strUserID, strSessID) = True Or _
                        '           objOption.fncVaccumMixCheck(objKtbnStrc) Then
                        '            intPositionInfo = KHKataban.fncGetMixManifoldInfo(objCon, objKtbnStrc, strUserID, strSessID)
                        '            objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullKataban & CdCst.Sign.Hypen
                        '            For intLoopPos2 = 1 To UBound(intPositionInfo)
                        '                objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & ClsCommon.fncPositionChance(intPositionInfo(intLoopPos2))
                        '            Next
                        '            If InStr(1, objKtbnStrc.strcSelection.strFullKataban, "-ST") <> 0 Then
                        '                objKtbnStrc.strcSelection.strFullManiKataban = Replace(objKtbnStrc.strcSelection.strFullManiKataban, "-ST", "")
                        '                objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & "-ST"
                        '            End If

                        '            '受注形番1,受注形番2
                        '            If objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length = 30 Then
                        '                .Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim
                        '                .Kataban2 = Space(30)
                        '            ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length < 30 Then
                        '                .Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim & _
                        '                            Space(30 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
                        '                .Kataban2 = Space(30)
                        '            ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length > 30 Then
                        '                .Kataban1 = Left(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 30)
                        '                .Kataban2 = Mid(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 31) & _
                        '                            Space(60 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
                        '            End If

                        '            'メッセージ欄に出す位置情報を作成
                        '            If objOption.fncVaccumMixCheck(objKtbnStrc) = False Then
                        '                intLoopPos = 1
                        '                While objKtbnStrc.strcSelection.strOptionKataban(intLoopPos).Trim.Length <> 0
                        '                    '記号設定
                        '                    If objKtbnStrc.strcSelection.intQuantity(intLoopPos) = 0 Then
                        '                    Else
                        '                        If strPos = "" Then
                        '                            strPos = objKtbnStrc.strcSelection.strAttributeSymbol(intLoopPos).Trim & "="
                        '                        Else
                        '                            strPos = strPos & "," & objKtbnStrc.strcSelection.strAttributeSymbol(intLoopPos).Trim & "="
                        '                        End If

                        '                        '位置情報設定
                        '                        For intMPos = 1 To 50 Step 2
                        '                            If Mid(objKtbnStrc.strcSelection.strPositionInfo(intLoopPos).Trim, intMPos, 1) = 1 Then
                        '                                intMPos2 = Int(intMPos / 2) + 1
                        '                                Select Case Right(strPos, 1)
                        '                                    Case (intMPos2 - 1)
                        '                                        strPos = strPos & "-"
                        '                                    Case "-"
                        '                                        If intRenPos <> intMPos2 - 1 Then
                        '                                            strPos = strPos & intMPos2 - 1 & "," & intMPos2
                        '                                        Else
                        '                                        End If
                        '                                    Case "="
                        '                                        strPos = strPos & intMPos2
                        '                                    Case Else
                        '                                        strPos = strPos & "," & intMPos2
                        '                                End Select
                        '                                intRenPos = intMPos2
                        '                            End If
                        '                        Next
                        '                        If Right(strPos, 1) = "-" Then
                        '                            strPos = strPos & intRenPos
                        '                        End If
                        '                    End If
                        '                    intLoopPos = intLoopPos + 1
                        '                End While
                        '            End If

                        '        End If
                        '        .MsgPosition = Left(strPos.Trim & Space(60), 60)
                        '    Case "12", "18", "19", "20", "21", "22", "23"
                        '        If objOption.fncVaccumMixCheck(objKtbnStrc) Then
                        '            .SpecExistsDiv = "Y"
                        '        Else
                        '            .SpecExistsDiv = "N"
                        '        End If
                        '        strPos = ""
                        '        .MsgPosition = Left(strPos.Trim & Space(60), 60)
                        '    Case "17"
                        '        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "X" Then
                        '            .SpecExistsDiv = "Y"
                        '        Else
                        '            .SpecExistsDiv = "N"
                        '        End If
                        '        strPos = ""
                        '        .MsgPosition = Left(strPos.Trim & Space(60), 60)
                        '    Case Else
                        '        .SpecExistsDiv = "Y"
                        '        strPos = ""
                        '        .MsgPosition = Left(strPos.Trim & Space(60), 60)
                        'End Select
                        ''機種コード
                        '.ModelCd = Left(objKtbnStrc.strcSelection.strModelNo.Trim & Space(2), 2)

                        '配線仕様有無区分
                        .WiringSpecDiv = LSet(objKtbnStrc.strcSelection.strWiringSpec.Trim, intSpaceCnt_01)
                        'レール長さ
                        .RailLength = Format(objKtbnStrc.strcSelection.decDinRailLength * 100, "000000")

                        ''処理日付＆マニホールド代表形番
                        'Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                        '    Case "", "00"
                        '        .ProcDatetime = ""
                        '        .FullKataban = ""
                        '    Case Else
                        '        Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        '            Case "CMF", "LMF0"
                        '                .ProcDatetime = Format(Now, "MMddhhmmss")
                        '                .FullKataban = Left(objKtbnStrc.strcSelection.strFullKataban.Trim & Space(30), 30)
                        '            Case Else
                        '                'その他マニホールドの場合
                        '                .ProcDatetime = ""
                        '                .FullKataban = ""
                        '        End Select
                        'End Select
                        ''形番チェック区分
                        '.KatabanCheckDiv = "Z" & Left(objKtbnStrc.strcSelection.strKatabanCheckDiv.Trim & Space(1), 1)
                        ''出荷場所
                        'If strPlaceCD Is Nothing OrElse strPlaceCD.Length <= 0 Then
                        '    .PlaceCd = Left(objKtbnStrc.strcSelection.strPlaceCd.Trim & Space(4), 4)
                        'Else
                        '    .PlaceCd = Left(strPlaceCD & Space(4), 4)
                        'End If

                        ''EL品判定区分
                        'If objKataban.fncELKatabanCheck(objCon, objKtbnStrc.strcSelection.strFullKataban.Trim, "1") Then
                        '    .ELDiv = CdCst.ELDiv.Yes
                        'Else
                        '    .ELDiv = CdCst.ELDiv.No
                        'End If
                    End With
                    '仕様書情報
                    Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                        Case "01", "02", "03", "04", "07", _
                             "08", "10", "11", "13", "14", _
                             "15", "16", "17", "96", "A1", "A2", _
                             "51", "53", "54", "55", "56", "57", _
                             "58", "59", "60", "61", "62", "63", _
                             "64", "65", "66", "67", "68", "69", _
                             "70", "71", "72", "73", "74", "75", _
                             "76", "77", "78", "79", "80", "81", _
                             "82", "83", "84", "85", "86", "87", _
                             "88", "89", "91", "92", "93", "A4", "A5", "A6", "A7", "A8", "A9", "B1", "98", _
                             "S", "T", "U", "B2", "B3", "B4"
                            For intLoopCnt = 1 To intLoopMax_01
                                With strcManifoldInfo(intLoopCnt)
                                    .AttributeSymbol = Space(intSpaceCnt_03)
                                    .OptionKataban = Space(intSpaceCnt_04)
                                    .PositionInfo = Space(intSpaceCnt_05)
                                    .Quantity = "00"
                                    .OrderNo = ""
                                End With
                            Next
                            For intLoopCnt = 1 To intLoopMax_02
                                With strcAccessoryInfo(intLoopCnt)
                                    .AttributeSymbol = Space(intSpaceCnt_03)
                                    .OptionKataban = Space(intSpaceCnt_04)
                                    .Quantity = "00"
                                End With
                            Next
                            Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                                Case "01"
                                    Dim intNo As Integer = 20      '表の行数（設置位置№が指定できる行数） ※画面変更時修正必要
                                    For intLoopCnt = 1 To intNo
                                        If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                                           objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                                            Select Case intLoopCnt
                                                Case 3
                                                    strTmpKataban = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                    strTmpPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                                    strTmpPositionInfo = Replace(strTmpPositionInfo, ",", "")
                                                    strTmpPositionInfo = Replace(strTmpPositionInfo, "0", " ")
                                                    strTmpPositionInfo = Replace(strTmpPositionInfo, "1", "Y")
                                                Case 4 To 12
                                                    If strTmpKataban.Trim = "" Then
                                                        intIndex = intIndex + 1
                                                        strcManifoldInfo(intIndex).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                        strcManifoldInfo(intIndex).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                                        strcManifoldInfo(intIndex).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                                        strcManifoldInfo(intIndex).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                                    Else
                                                        intTmpQuantity = 0
                                                        strTmpPositionInfo1 = ""
                                                        strTmpPositionInfo2 = ""
                                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")

                                                        For intLoopCnt1 = 1 To strPositionInfo.Length - 1
                                                            If Mid(strTmpPositionInfo, intLoopCnt1, 1) = "Y" Then
                                                                If Mid(strPositionInfo, intLoopCnt1, 1) = "Y" Then
                                                                    strTmpPositionInfo1 = strTmpPositionInfo1 & Mid(strTmpPositionInfo, intLoopCnt1, 1)
                                                                    strTmpPositionInfo2 = strTmpPositionInfo2 & " "
                                                                    intTmpQuantity = intTmpQuantity + 1
                                                                Else
                                                                    strTmpPositionInfo1 = strTmpPositionInfo1 & " "
                                                                    strTmpPositionInfo2 = strTmpPositionInfo2 & Mid(strPositionInfo, intLoopCnt1, 1)
                                                                End If
                                                            Else
                                                                strTmpPositionInfo1 = strTmpPositionInfo1 & Mid(strTmpPositionInfo, intLoopCnt1, 1)
                                                                strTmpPositionInfo2 = strTmpPositionInfo2 & Mid(strPositionInfo, intLoopCnt1, 1)
                                                            End If
                                                        Next
                                                        If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) - intTmpQuantity > 0 Then
                                                            intIndex = intIndex + 1
                                                            strcManifoldInfo(intIndex).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                            strcManifoldInfo(intIndex).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                                            strcManifoldInfo(intIndex).PositionInfo = LSet(strTmpPositionInfo2, intSpaceCnt_05)
                                                            strcManifoldInfo(intIndex).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt) - intTmpQuantity, "00")
                                                        End If
                                                        If intTmpQuantity > 0 Then
                                                            intIndex = intIndex + 1
                                                            strcManifoldInfo(intIndex).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                            strcManifoldInfo(intIndex).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & strTmpKataban.Trim, intSpaceCnt_04)
                                                            strcManifoldInfo(intIndex).PositionInfo = LSet(strTmpPositionInfo1, intSpaceCnt_05)
                                                            strcManifoldInfo(intIndex).Quantity = Format(intTmpQuantity, "00")
                                                        End If
                                                    End If
                                                Case Else
                                                    intIndex = intIndex + 1
                                                    strcManifoldInfo(intIndex).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcManifoldInfo(intIndex).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                                    strPositionInfo = Replace(strPositionInfo, ",", "")
                                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
                                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                                    strcManifoldInfo(intIndex).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                                    strcManifoldInfo(intIndex).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                            End Select
                                        End If
                                    Next
                                    '21行目以降の処理
                                    For intLoopCnt = 1 To 10        '※画面変更時修正必要
                                        intNo = intNo + 1       '現在の行
                                        Select Case intLoopCnt
                                            Case 1 To 4
                                                With strcAccessoryInfo(intLoopCnt)
                                                    .AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intNo).Trim, intSpaceCnt_03)
                                                    .OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intNo).Trim, intSpaceCnt_04)
                                                    .Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intNo), "00")
                                                End With
                                            Case 5 To 8
                                                With strcAccessoryInfo(intLoopCnt)
                                                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intNo).Trim
                                                        Case CdCst.Manifold.InspReportJp.SelectValue
                                                            .AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intNo).Trim, intSpaceCnt_03)
                                                            .OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
                                                            .Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intNo), "00")
                                                        Case CdCst.Manifold.InspReportEn.SelectValue
                                                            .AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intNo).Trim, intSpaceCnt_03)
                                                            .OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
                                                            .Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intNo), "00")
                                                        Case Else
                                                            .AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intNo).Trim, intSpaceCnt_03)

                                                            If Left(objKtbnStrc.strcSelection.strOptionKataban(intNo).Trim & Space(9), 9) = "検査成績書（英文）" Then

                                                                .OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)

                                                            ElseIf Left(objKtbnStrc.strcSelection.strOptionKataban(intNo).Trim & Space(9), 9) = "検査成績書（和文）" Then

                                                                .OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)

                                                            Else

                                                                .OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intNo).Trim, intSpaceCnt_04)

                                                            End If

                                                            .Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intNo), "00")
                                                    End Select
                                                End With
                                            Case 9
                                                With strcAccessoryInfo(intLoopCnt)
                                                    If objKtbnStrc.strcSelection.strOptionKataban(intNo).Trim = "1" Then
                                                        .AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intNo).Trim, intSpaceCnt_03)
                                                        .OptionKataban = Space(intSpaceCnt_04)
                                                        .Quantity = Format(0, "00")
                                                    Else
                                                        .AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intNo).Trim, intSpaceCnt_03)
                                                        .OptionKataban = LSet(CdCst.Manifold.TubeRemover.DummyValue, intSpaceCnt_04)
                                                        .Quantity = Format(1, "00")
                                                    End If
                                                End With
                                        End Select
                                    Next
                                Case "02"
                                    For intLoopCnt = 1 To 16        '各スペック№カウントの変更が必要
                                        strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                        strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    For intLoopCnt = 17 To 21
                                        strcAccessoryInfo(intLoopCnt - 16).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            Case CdCst.Manifold.InspReportJp.SelectValue
                                                strcAccessoryInfo(intLoopCnt - 16).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
                                            Case CdCst.Manifold.InspReportEn.SelectValue
                                                strcAccessoryInfo(intLoopCnt - 16).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
                                            Case Else
                                                strcAccessoryInfo(intLoopCnt - 16).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        End Select
                                        strcAccessoryInfo(intLoopCnt - 16).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next

                                    If objKtbnStrc.strcSelection.strOpSymbol(1).PadRight(2, " ").Substring(0, 2) = "80" Then
                                        intPositionInfo = KHKataban.fncGetMixManifoldInfo(objCon, objKtbnStrc, strUserID, strSessID)
                                        objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullKataban & CdCst.Sign.Hypen
                                        For intLoopPos2 = 1 To UBound(intPositionInfo)
                                            objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & ClsCommon.fncPositionChance(intPositionInfo(intLoopPos2))
                                        Next
                                        If InStr(1, objKtbnStrc.strcSelection.strFullKataban, "-ST") <> 0 Then
                                            objKtbnStrc.strcSelection.strFullManiKataban = Replace(objKtbnStrc.strcSelection.strFullManiKataban, "-ST", "")
                                            objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & "-ST"
                                        End If
                                        '受注形番1,受注形番2
                                        If objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length = 30 Then
                                            strcHeader.Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim
                                            strcHeader.Kataban2 = Space(30)
                                        ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length < 30 Then
                                            strcHeader.Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim & _
                                                        Space(30 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
                                            strcHeader.Kataban2 = Space(30)
                                        ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length > 30 Then
                                            strcHeader.Kataban1 = Left(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 30)
                                            strcHeader.Kataban2 = Mid(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 31) & _
                                                        Space(60 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
                                        End If
                                    End If

                                Case "03"   '機種　M
                                    For intLoopCnt = 1 To 15
                                        strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        If objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim = "" Then
                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        Else
                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                              objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim & _
                                                                                              objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        End If
                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                        strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    For intLoopCnt = 16 To 23
                                        If intLoopCnt <> 23 Then
                                            strcAccessoryInfo(intLoopCnt - 15).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                Case CdCst.Manifold.InspReportJp.SelectValue
                                                    strcAccessoryInfo(intLoopCnt - 15).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
                                                Case CdCst.Manifold.InspReportEn.SelectValue
                                                    strcAccessoryInfo(intLoopCnt - 15).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
                                                Case Else
                                                    strcAccessoryInfo(intLoopCnt - 15).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                            End Select
                                            strcAccessoryInfo(intLoopCnt - 15).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                        Else
                                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                Case CdCst.Manifold.TubeRemover.Necessity
                                                    strcAccessoryInfo(intLoopCnt - 15).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcAccessoryInfo(intLoopCnt - 15).OptionKataban = Space(intSpaceCnt_04)
                                                    strcAccessoryInfo(intLoopCnt - 15).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                                Case CdCst.Manifold.TubeRemover.UnNecessity
                                                    strcAccessoryInfo(intLoopCnt - 15).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcAccessoryInfo(intLoopCnt - 15).OptionKataban = LSet(CdCst.Manifold.TubeRemover.DummyValue, intSpaceCnt_04)
                                                    strcAccessoryInfo(intLoopCnt - 15).Quantity = Format(1, "00")
                                                Case Else
                                                    strcAccessoryInfo(intLoopCnt - 15).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcAccessoryInfo(intLoopCnt - 15).OptionKataban = Space(intSpaceCnt_04)
                                                    strcAccessoryInfo(intLoopCnt - 15).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                            End Select
                                        End If
                                    Next

                                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "8" Then
                                        intPositionInfo = KHKataban.fncGetMixManifoldInfo(objCon, objKtbnStrc, strUserID, strSessID)
                                        objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullKataban & CdCst.Sign.Hypen
                                        For intLoopPos2 = 1 To UBound(intPositionInfo)
                                            objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & ClsCommon.fncPositionChance(intPositionInfo(intLoopPos2))
                                        Next
                                        If InStr(1, objKtbnStrc.strcSelection.strFullKataban, "-ST") <> 0 Then
                                            objKtbnStrc.strcSelection.strFullManiKataban = Replace(objKtbnStrc.strcSelection.strFullManiKataban, "-ST", "")
                                            objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & "-ST"
                                        End If
                                        '受注形番1,受注形番2
                                        If objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length = 30 Then
                                            strcHeader.Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim
                                            strcHeader.Kataban2 = Space(30)
                                        ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length < 30 Then
                                            strcHeader.Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim & _
                                                        Space(30 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
                                            strcHeader.Kataban2 = Space(30)
                                        ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length > 30 Then
                                            strcHeader.Kataban1 = Left(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 30)
                                            strcHeader.Kataban2 = Mid(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 31) & _
                                                        Space(60 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
                                        End If
                                    End If

                                    'RM1803032_スペーサ行追加対応
                                Case "04"
                                    Dim intManiEnd As Integer = 16
                                    For intLoopCnt = 1 To intManiEnd
                                        strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        If objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim = "" Then
                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        Else
                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                              objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim & _
                                                                                              objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        End If
                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                        strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    For intLoopCnt = 17 To 25
                                        If intLoopCnt <> 25 Then
                                            strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                Case CdCst.Manifold.InspReportJp.SelectValue
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
                                                Case CdCst.Manifold.InspReportEn.SelectValue
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
                                                Case Else
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                            End Select
                                            strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                        Else
                                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                Case CdCst.Manifold.TubeRemover.Necessity
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = Space(intSpaceCnt_04)
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                                Case CdCst.Manifold.TubeRemover.UnNecessity
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.TubeRemover.DummyValue, intSpaceCnt_04)
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(1, "00")
                                                Case Else
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = Space(intSpaceCnt_04)
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                            End Select
                                        End If
                                    Next
                                    'RM1803032_スペーサ行追加対応
                                Case "07", "96"
                                    Dim intManiEnd As Integer = 21
                                    For intLoopCnt = 1 To intManiEnd
                                        strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                        strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    For intLoopCnt = 22 To 28
                                        If intLoopCnt <> 28 Then
                                            strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                Case CdCst.Manifold.InspReportJp.SelectValue
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
                                                Case CdCst.Manifold.InspReportEn.SelectValue
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
                                                Case Else
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                            End Select
                                            strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                        Else
                                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                Case CdCst.Manifold.TubeRemover.Necessity
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = Space(intSpaceCnt_04)
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                                Case CdCst.Manifold.TubeRemover.UnNecessity
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.TubeRemover.DummyValue, intSpaceCnt_04)
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(1, "00")
                                                Case Else
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = Space(intSpaceCnt_04)
                                                    strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                            End Select
                                        End If
                                    Next
                                    '2018/03/08_タグ銘板設定時値セット
                                    If strcAccessoryInfo(6).AttributeSymbol = "T6" And strcAccessoryInfo(6).Quantity <> "00" Then
                                        strcAccessoryInfo(8).AttributeSymbol = "L1"
                                        Dim main As New _Main
                                        strcAccessoryInfo(8).OptionKataban = LSet(main.Session("decDinRailLength").ToString, intSpaceCnt_04)
                                        strcAccessoryInfo(8).Quantity = "01"
                                    End If

                                Case "08"
                                    For intLoopCnt = 1 To 16
                                        strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                        strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    For intLoopCnt = 17 To 21
                                        strcAccessoryInfo(intLoopCnt - 16).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            Case CdCst.Manifold.InspReportJp.SelectValue
                                                strcAccessoryInfo(intLoopCnt - 16).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
                                            Case CdCst.Manifold.InspReportEn.SelectValue
                                                strcAccessoryInfo(intLoopCnt - 16).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
                                            Case Else
                                                strcAccessoryInfo(intLoopCnt - 16).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        End Select
                                        strcAccessoryInfo(intLoopCnt - 16).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "8" Then
                                        intPositionInfo = KHKataban.fncGetMixManifoldInfo(objCon, objKtbnStrc, strUserID, strSessID)
                                        objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullKataban & CdCst.Sign.Hypen
                                        For intLoopPos2 = 1 To UBound(intPositionInfo)
                                            objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & ClsCommon.fncPositionChance(intPositionInfo(intLoopPos2))
                                        Next
                                        If InStr(1, objKtbnStrc.strcSelection.strFullKataban, "-ST") <> 0 Then
                                            objKtbnStrc.strcSelection.strFullManiKataban = Replace(objKtbnStrc.strcSelection.strFullManiKataban, "-ST", "")
                                            objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & "-ST"
                                        End If
                                        '受注形番1,受注形番2
                                        If objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length = 30 Then
                                            strcHeader.Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim
                                            strcHeader.Kataban2 = Space(30)
                                        ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length < 30 Then
                                            strcHeader.Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim & _
                                                        Space(30 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
                                            strcHeader.Kataban2 = Space(30)
                                        ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length > 30 Then
                                            strcHeader.Kataban1 = Left(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 30)
                                            strcHeader.Kataban2 = Mid(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 31) & _
                                                        Space(60 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
                                        End If
                                    End If
                                Case "10"
                                    For intLoopCnt = 1 To 14
                                        strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                        strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    For intLoopCnt = 15 To 23
                                        If intLoopCnt <> 23 Then
                                            strcAccessoryInfo(intLoopCnt - 14).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                Case CdCst.Manifold.InspReportJp.SelectValue
                                                    strcAccessoryInfo(intLoopCnt - 14).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
                                                Case CdCst.Manifold.InspReportEn.SelectValue
                                                    strcAccessoryInfo(intLoopCnt - 14).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
                                                Case Else
                                                    strcAccessoryInfo(intLoopCnt - 14).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                            End Select
                                            strcAccessoryInfo(intLoopCnt - 14).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                        Else
                                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                Case CdCst.Manifold.TubeRemover.Necessity
                                                    strcAccessoryInfo(intLoopCnt - 14).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcAccessoryInfo(intLoopCnt - 14).OptionKataban = Space(intSpaceCnt_04)
                                                    strcAccessoryInfo(intLoopCnt - 14).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                                Case CdCst.Manifold.TubeRemover.UnNecessity
                                                    strcAccessoryInfo(intLoopCnt - 14).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcAccessoryInfo(intLoopCnt - 14).OptionKataban = LSet(CdCst.Manifold.TubeRemover.DummyValue, intSpaceCnt_04)
                                                    strcAccessoryInfo(intLoopCnt - 14).Quantity = Format(1, "00")
                                                Case Else
                                                    strcAccessoryInfo(intLoopCnt - 14).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                    strcAccessoryInfo(intLoopCnt - 14).OptionKataban = Space(intSpaceCnt_04)
                                                    strcAccessoryInfo(intLoopCnt - 14).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                            End Select
                                        End If
                                    Next
                                Case "11"
                                    For intLoopCnt = 1 To 15
                                        strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                        strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    For intLoopCnt = 16 To 18
                                        strcAccessoryInfo(intLoopCnt - 15).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        strcAccessoryInfo(intLoopCnt - 15).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        strcAccessoryInfo(intLoopCnt - 15).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                Case "13"
                                    For intLoopCnt = 1 To 17
                                        strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                        strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    For intLoopCnt = 18 To 24
                                        strcAccessoryInfo(intLoopCnt - 17).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            Case CdCst.Manifold.InspReportJp.SelectValue
                                                strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
                                            Case CdCst.Manifold.InspReportEn.SelectValue
                                                strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
                                            Case Else
                                                strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        End Select
                                        strcAccessoryInfo(intLoopCnt - 17).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                Case "14"
                                    For intLoopCnt = 1 To 6
                                        strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                        strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    For intLoopCnt = 7 To 9
                                        strcAccessoryInfo(intLoopCnt - 6).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        strcAccessoryInfo(intLoopCnt - 6).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        strcAccessoryInfo(intLoopCnt - 6).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    'RM1803032_スペーサ行追加対応
                                Case "15"
                                    Dim intManiEnd As Integer = 21
                                    For intLoopCnt = 1 To intManiEnd
                                        strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                        strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    For intLoopCnt = 22 To 29
                                        strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            Case CdCst.Manifold.InspReportJp.SelectValue
                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
                                            Case CdCst.Manifold.InspReportEn.SelectValue
                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
                                            Case Else
                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        End Select
                                        strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                Case "16"
                                    For intLoopCnt = 1 To 20
                                        strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)

                                        Select Case intLoopCnt
                                            Case 19
                                                strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                            Case 20
                                                strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                            Case Else
                                                strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        End Select
                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                        strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                        'DELETE BY YGY 20141218
                                        'If intLoopCnt = 18 Then
                                        '    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2, "00")
                                        'Else
                                        '    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                        'End If
                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                    For intLoopCnt = 21 To 25
                                        strcAccessoryInfo(intLoopCnt - 20).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            Case CdCst.Manifold.InspReportJp.SelectValue
                                                strcAccessoryInfo(intLoopCnt - 20).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
                                            Case CdCst.Manifold.InspReportEn.SelectValue
                                                strcAccessoryInfo(intLoopCnt - 20).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
                                            Case Else
                                                strcAccessoryInfo(intLoopCnt - 20).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                        End Select
                                        strcAccessoryInfo(intLoopCnt - 20).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    Next
                                Case "17"
                                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "X" Then Exit Select '2013/08/02
                                    For intLoopCnt = 1 To 5
                                        If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                                            strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                            strOpArray = Split(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, CdCst.Sign.Comma)
                                            For intIndex = 0 To strOpArray.Length - 1
                                                strcManifoldInfo(intLoopCnt).OptionKataban = strcManifoldInfo(intLoopCnt).OptionKataban.Trim & _
                                                                                             strOpArray(intIndex).Trim
                                            Next
                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(KHKataban.fncHypenCut(strcManifoldInfo(intLoopCnt).OptionKataban), intSpaceCnt_04)
                                            strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                            strPositionInfo = Replace(strPositionInfo, ",", "")
                                            strPositionInfo = Replace(strPositionInfo, "0", " ")
                                            strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                            strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                            strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                        End If
                                    Next
                                    strcAccessoryInfo(1).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(6).Trim, intSpaceCnt_03)
                                    strOpArray = Split(objKtbnStrc.strcSelection.strOptionKataban(6).Trim, CdCst.Sign.Comma)
                                    For intIndex = 0 To strOpArray.Length - 1
                                        strcAccessoryInfo(1).OptionKataban = strcAccessoryInfo(1).OptionKataban.Trim & _
                                                                             strOpArray(intIndex).Trim
                                    Next
                                    strcAccessoryInfo(1).OptionKataban = LSet(KHKataban.fncHypenCut(strcAccessoryInfo(1).OptionKataban), intSpaceCnt_04)
                                    strcAccessoryInfo(1).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(6), "00")
                                Case "A1", "A2", "51", "53", "54", "55", "56", "57", "58", "59", "60", "61", "62", "63", "64", _
                                     "65", "66", "67", "68", "69", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", _
                                     "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "91", "92", "93", "A4", "A5", "A6", "A7", "A8", "A9", "B1", _
                                     "S", "T", "U", "B2", "B3", "B4"
                                    If objOption.fncVaccumMixCheck(objKtbnStrc) Then
                                        Dim str() As String = Nothing
                                        Dim dtSpecItem As New DataTable
                                        Dim dtContent As New DataTable
                                        Call KHManifold.subInitTable(dtSpecItem, dtContent)
                                        Call SiyouDAL.subSQL_ItemMst(objCon, objKtbnStrc.strcSelection.strSpecNo.Trim, dtSpecItem, dtContent)
                                        Dim listResult As ArrayList = KHManifold.fncGetNewKataban(dtSpecItem, dtContent, objKtbnStrc.strcSelection.strSpecNo.Trim, _
                                                                        objKtbnStrc.strcSelection.strSeriesKataban, objKtbnStrc.strcSelection.strOpSymbol, objKtbnStrc.strcSelection.strKeyKataban)
                                        For intLoopCnt = 1 To listResult.Count
                                            str = listResult(intLoopCnt - 1).ToString.Split("_")
                                            If str.Length >= 4 Then
                                                strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                strcManifoldInfo(intLoopCnt).OptionKataban = LSet(str(1).Trim, intSpaceCnt_04)
                                                strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                                strPositionInfo = Replace(strPositionInfo, ",", "")
                                                strPositionInfo = Replace(strPositionInfo, "0", " ")
                                                strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                                strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                                strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                            End If
                                        Next
                                        For intLoopCnt = listResult.Count + 1 To 12
                                            If objKtbnStrc.strcSelection.strAttributeSymbol.Length <= intLoopCnt Then
                                                strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet("", intSpaceCnt_03)
                                                'strcManifoldInfo(intLoopCnt).OptionKataban = Left("", 20)
                                                strcManifoldInfo(intLoopCnt).OptionKataban = LSet("", intSpaceCnt_04)
                                                strPositionInfo = ""
                                                strPositionInfo = Replace(strPositionInfo, ",", "")
                                                strPositionInfo = Replace(strPositionInfo, "0", " ")
                                                strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                                strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                                strcManifoldInfo(intLoopCnt).Quantity = Format(0, "00")
                                            Else
                                                strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                                strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                                strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                                strPositionInfo = Replace(strPositionInfo, ",", "")
                                                strPositionInfo = Replace(strPositionInfo, "0", " ")
                                                strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                                strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                                strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                            End If

                                        Next
                                    End If
                            End Select
                        Case "05", "06"
                            '初期化
                            For intLoopCnt = 1 To 20
                                With strcManifoldInfo(intLoopCnt)
                                    .AttributeSymbol = Space(2)
                                    .OptionKataban = Space(20)
                                    .PositionInfo = Space(10)
                                    .Quantity = "00"
                                    .OrderNo = Space(8)
                                End With
                            Next
                            For intLoopCnt = 1 To 10
                                With strcAccessoryInfo(intLoopCnt)
                                    .AttributeSymbol = Space(2)
                                    .OptionKataban = Space(30)
                                    .Quantity = "00"
                                End With
                            Next
                            '設定
                            Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                                Case "05"
                                Case "06"
                            End Select
                        Case "09"
                            '初期化
                            For intLoopCnt = 1 To intLoopMax_01
                                With strcManifoldInfo(intLoopCnt)
                                    .AttributeSymbol = Space(intSpaceCnt_03)
                                    .OptionKataban = Space(intSpaceCnt_04)
                                    .PositionInfo = Space(intSpaceCnt_05)
                                    .Quantity = "00"
                                    .OrderNo = ""
                                End With
                            Next
                            For intLoopCnt = 1 To intLoopMax_02
                                With strcAccessoryInfo(intLoopCnt)
                                    .AttributeSymbol = Space(intSpaceCnt_03)
                                    .OptionKataban = Space(intSpaceCnt_04)
                                    .Quantity = "00"
                                End With
                            Next
                            '設定
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                                For intLoopCnt = 1 To 17
                                    strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                    Select Case intLoopCnt
                                        Case 16
                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "P", intSpaceCnt_04)
                                        Case 17
                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "R", intSpaceCnt_04)
                                        Case Else
                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                    End Select
                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                    strPositionInfo = Replace(strPositionInfo, ",", "")
                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                    strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                    Select Case intLoopCnt
                                        Case 17
                                            strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2, "00")
                                        Case Else
                                            strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    End Select
                                Next
                                For intLoopCnt = 18 To 23
                                    strcAccessoryInfo(intLoopCnt - 17).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                        Case CdCst.Manifold.InspReportJp.SelectValue
                                            strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
                                        Case CdCst.Manifold.InspReportEn.SelectValue
                                            strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
                                        Case Else
                                            Select Case intLoopCnt
                                                Case 20
                                                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4TB3" Then
                                                        strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet("R1/4", intSpaceCnt_04)
                                                    Else
                                                        strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet("R3/8", intSpaceCnt_04)
                                                    End If
                                                Case 21
                                                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4TB3" Then
                                                        strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet("R3/8", intSpaceCnt_04)
                                                    Else
                                                        strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet("R1/2", intSpaceCnt_04)
                                                    End If
                                                Case Else
                                                    strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
                                            End Select
                                    End Select
                                    strcAccessoryInfo(intLoopCnt - 17).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                Next
                            End If
                        Case "12", "18", "19", "20", "21", "22", "23"
                            '初期化
                            For intLoopCnt = 1 To intLoopMax_01
                                With strcManifoldInfo(intLoopCnt)
                                    .AttributeSymbol = Space(intSpaceCnt_03)
                                    .OptionKataban = Space(intSpaceCnt_04)
                                    .PositionInfo = Space(intSpaceCnt_05)
                                    .Quantity = "00"
                                    .OrderNo = ""
                                End With
                            Next
                            For intLoopCnt = 1 To intLoopMax_02
                                With strcAccessoryInfo(intLoopCnt)
                                    .AttributeSymbol = Space(intSpaceCnt_03)
                                    .OptionKataban = Space(intSpaceCnt_04)
                                    .Quantity = "00"
                                End With
                            Next
                            '設定
                            If objOption.fncVaccumMixCheck(objKtbnStrc) Then
                                intIndex = 0
                                For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                                    If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                                        intIndex = intIndex + 1
                                        strcManifoldInfo(intIndex).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
                                        strOpArray = Split(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, CdCst.Sign.Comma)
                                        For intLoopCnt1 = 0 To strOpArray.Length - 1
                                            strcManifoldInfo(intIndex).OptionKataban = strcManifoldInfo(intIndex).OptionKataban.Trim & _
                                                                                       strOpArray(intLoopCnt1).Trim
                                        Next
                                        strcManifoldInfo(intIndex).OptionKataban = LSet(KHKataban.fncHypenCut(strcManifoldInfo(intIndex).OptionKataban), intSpaceCnt_04)
                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
                                        strPositionInfo = Replace(strPositionInfo, ",", "")
                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
                                        strcManifoldInfo(intIndex).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
                                        strcManifoldInfo(intIndex).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
                                    End If
                                Next
                            End If
                        Case Else
                            '初期化
                            For intLoopCnt = 1 To intLoopMax_01
                                With strcManifoldInfo(intLoopCnt)
                                    .AttributeSymbol = Space(intSpaceCnt_03)
                                    .OptionKataban = Space(intSpaceCnt_04)
                                    .PositionInfo = Space(intSpaceCnt_05)
                                    .Quantity = "00"
                                    .OrderNo = ""
                                End With
                            Next
                            For intLoopCnt = 1 To intLoopMax_02
                                With strcAccessoryInfo(intLoopCnt)
                                    .AttributeSymbol = Space(intSpaceCnt_03)
                                    .OptionKataban = Space(intSpaceCnt_04)
                                    .Quantity = "00"
                                End With
                            Next
                    End Select

                    '文字列結合
                    With sbBuilder
                        '.Append(strcHeader.SystemDatetime)
                        '.Append(strcHeader.Kataban1)
                        '.Append(strcHeader.Kataban2)
                        '.Append(strcHeader.PsPrice)
                        '.Append(strcHeader.GsPrice)
                        '.Append(strcHeader.ListPrice)
                        '.Append(strcHeader.SpecExistsDiv)
                        '.Append(strcHeader.ModelCd)
                        .Append(strcHeader.WiringSpecDiv)
                        .Append(strcHeader.RailLength)
                        '.Append(strcHeader.ProcDatetime)
                        '.Append(strcHeader.FullKataban)
                        For intLoopCnt = 1 To intLoopMax_01
                            .Append(strcManifoldInfo(intLoopCnt).AttributeSymbol)
                            .Append(strcManifoldInfo(intLoopCnt).OptionKataban)
                            .Append(strcManifoldInfo(intLoopCnt).PositionInfo)
                            .Append(strcManifoldInfo(intLoopCnt).Quantity)
                            '.Append(strcManifoldInfo(intLoopCnt).OrderNo)
                        Next
                        For intLoopCnt = 1 To intLoopMax_02
                            .Append(strcAccessoryInfo(intLoopCnt).AttributeSymbol)
                            .Append(strcAccessoryInfo(intLoopCnt).OptionKataban)
                            .Append(strcAccessoryInfo(intLoopCnt).Quantity)
                        Next
                        '.Append(strcHeader.KatabanCheckDiv)
                        '.Append(strcHeader.PlaceCd)
                        '.Append(strcHeader.ELDiv)
                        '.Append(strcHeader.MsgPosition)
                        .Append(vbCrLf)
                    End With
            End Select

            '戻り値設定
            fncSBOInterfaceGet = sbBuilder.ToString
        Catch ex As Exception
            WriteErrorLog("E001", ex)
            fncSBOInterfaceGet = ""
        Finally
            sbBuilder = Nothing
            objKataban = Nothing
            objOption = Nothing
        End Try

    End Function

    ''' <summary>
    ''' SBOインターフェース情報取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strUserID"></param>
    ''' <param name="strSessID"></param>
    ''' <returns></returns>
    ''' <remarks>SBOにインターフェースする情報を編集し返却する</remarks>
    Public Shared Function fncJutyuEdiInterfaceGet(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, _
                                             ByVal strOfficeCd As String, strUserID As String, _
                                             strSessID As String, ByVal strFobPrice As String, _
                                            ByVal strCurrencyCode As String, ByVal strKeyInfo As String, _
                                            ByVal blCZFlag As Boolean, _
                                            ByVal intItemRow As Integer) As KHSBOInterface

    End Function


    '    Dim objOption As New KHOptionCtl
    '    Dim objKataban As New KHKataban
    '    Dim sbBuilder As New System.Text.StringBuilder(2737)

    '    Dim strOpArray() As String
    '    Dim strPositionInfo As String
    '    Dim intLoopCnt As Integer
    '    Dim intLoopCnt1 As Integer
    '    Dim intIndex As Integer = 0

    '    Dim strAccAttributeSymbol() As String = Nothing
    '    Dim strAccOptionKataban() As String = Nothing
    '    Dim strAccPositionInfo() As String = Nothing
    '    Dim intAccQuantity() As Integer = Nothing

    '    Dim strTmpKataban As String = ""
    '    Dim strTmpPositionInfo As String = ""
    '    Dim strTmpPositionInfo1 As String = ""
    '    Dim strTmpPositionInfo2 As String = ""
    '    Dim intTmpQuantity As Integer

    '    Dim strPos As String
    '    Dim intLoopPos As Integer
    '    Dim intRenPos As Integer
    '    Dim intMPos As Integer
    '    Dim intMPos2 As Integer
    '    Dim intPositionInfo() As Integer
    '    Dim intLoopPos2 As Integer
    '    Dim strShiyouInfo As String = String.Empty
    '    Dim intLineNo As Integer = 0


    '    Dim result As New KHSBOInterface

    '    Try
    '        fncJutyuEdiInterfaceGet = Nothing

    '        ReDim strcManifoldInfo(intLoopMax_01)
    '        ReDim strcAccessoryInfo(intLoopMax_02)

    '        '仕様書情報
    '        Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
    '            Case "05", "06"
    '                Dim intCount As Integer = 0
    '                For intLoopCnt1 = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 2
    '                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1).Trim <> "" And _
    '                       objKtbnStrc.strcSelection.intQuantity(intLoopCnt1) <> 0 And intLineNo < intItemRow Then

    '                        'result.clKatahikiInfoDto = New WebKataban.CommonDbService.KatahikiInfoDto

    '                        '編集
    '                        With strcHeader
    '                            'FOB対応
    '                            Dim lstFobPrice As List(Of String) = strFobPrice.Split(",").ToList
    '                            If lstFobPrice.Count - 1 >= intCount Then
    '                                Dim decFobPrice As Decimal
    '                                decFobPrice = IIf(Decimal.TryParse(lstFobPrice(intCount), decFobPrice), decFobPrice, 0)
    '                                If decFobPrice - Fix(decFobPrice) = 0 Then
    '                                    .NetPrice = Decimal.ToInt32(decFobPrice)
    '                                    'intCount = intCount + 1
    '                                Else
    '                                    .NetPrice = Format(decFobPrice, "#.00")
    '                                    'intCount = intCount + 1
    '                                End If
    '                            End If

    '                            Dim lstCurrencyCode As List(Of String) = strCurrencyCode.Split(",").ToList
    '                            If lstCurrencyCode.Count - 1 >= intCount Then
    '                                Dim strCurrency As String
    '                                strCurrency = lstCurrencyCode(intCount)
    '                                If .NetPrice > 0 Then
    '                                    .CurrencyCode = strCurrency
    '                                Else
    '                                    .CurrencyCode = Nothing
    '                                End If
    '                            End If

    '                            intCount = intCount + 1

    '                            '定価
    '                            .ListPrice = Decimal.ToInt32(objKtbnStrc.strcSelection.intOpRegPrice(intLoopCnt1))
    '                            '登録店価格
    '                            .RegPrice = Decimal.ToInt32(objKtbnStrc.strcSelection.intOpListPrice(intLoopCnt1))
    '                            'SS価格
    '                            .SsPrice = Decimal.ToInt32(objKtbnStrc.strcSelection.intOpSsPrice(intLoopCnt1))
    '                            'BS価格
    '                            .BsPrice = Decimal.ToInt32(objKtbnStrc.strcSelection.intOpBsPrice(intLoopCnt1))
    '                            'GS価格
    '                            .GsPrice = Decimal.ToInt32(objKtbnStrc.strcSelection.intOpGsPrice(intLoopCnt1))
    '                            'PS価格
    '                            .PsPrice = Decimal.ToInt32(objKtbnStrc.strcSelection.intOpPsPrice(intLoopCnt1))

    '                            '形番
    '                            .FullKataban = Left(objKtbnStrc.strcSelection.strOpKataban(intLoopCnt1) & Space(30), 30)

    '                            .Quantity = objKtbnStrc.strcSelection.intQuantity(intLoopCnt1)

    '                            '仕様書有無区分
    '                            .SpecExistsDiv = "Y"
    '                            '機種コード
    '                            .ModelCd = Left(objKtbnStrc.strcSelection.strModelNo.Trim & Space(2), 2)
    '                            '配線仕様有無区分
    '                            .WiringSpecDiv = Left(objKtbnStrc.strcSelection.strWiringSpec.Trim & Space(1), 1)
    '                            'レール長さ
    '                            .RailLength = Format(objKtbnStrc.strcSelection.decDinRailLength * 100, "000000")
    '                            '処理日付＆マニホールド代表形番
    '                            .ProcDatetime = Format(Now, "MMddhhmmss")
    '                            '.FullKataban = Left(objKtbnStrc.strcSelection.strFullKataban.Trim & Space(30), 30)
    '                            '形番チェック区分
    '                            .KatabanCheckDiv = "Z" & Left(objKtbnStrc.strcSelection.strOpKatabanCheckDiv(intLoopCnt1).Trim & Space(1), 1)
    '                            '出荷場所
    '                            .PlaceCd = Left(objKtbnStrc.strcSelection.strOpPlaceCd(intLoopCnt1).Trim & Space(4), 4)
    '                            'EL品判定区分
    '                            If objKataban.fncELKatabanCheck(objCon, objKtbnStrc.strcSelection.strFullKataban.Trim, "1") Then
    '                                .ELDiv = True
    '                            Else
    '                                .ELDiv = False
    '                            End If
    '                            strPos = ""
    '                        End With

    '                        '仕様書情報
    '                        Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
    '                            Case "05", "06"
    '                                '初期化
    '                                For intLoopCnt = 1 To 20
    '                                    With strcManifoldInfo(intLoopCnt)
    '                                        .AttributeSymbol = Space(2)
    '                                        .OptionKataban = Space(30)
    '                                        .PositionInfo = Space(10)
    '                                        .Quantity = "00"
    '                                        .OrderNo = Space(8)
    '                                    End With
    '                                Next
    '                                For intLoopCnt = 1 To 10
    '                                    With strcAccessoryInfo(intLoopCnt)
    '                                        .AttributeSymbol = Space(2)
    '                                        .OptionKataban = Space(30)
    '                                        .Quantity = "00"
    '                                    End With
    '                                Next
    '                                '設定
    '                                Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
    '                                    Case "05"
    '                                        intIndex = 0
    '                                        For intLoopCnt = 1 To 25
    '                                            If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
    '                                                intIndex = intIndex + 1
    '                                                strcManifoldInfo(intIndex).AttributeSymbol = Left(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim & Space(2), 2)
    '                                                strcManifoldInfo(intIndex).OptionKataban = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & Space(30), 30)
    '                                                strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                                strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                                strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                                strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                                strcManifoldInfo(intIndex).PositionInfo = Left(strPositionInfo & Space(10), 10)
    '                                                strcManifoldInfo(intIndex).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                            End If
    '                                        Next
    '                                        '付属品取得
    '                                        Call subISOAccessoryGet(objKtbnStrc, strAccAttributeSymbol, strAccOptionKataban, intAccQuantity)
    '                                        intIndex = 0
    '                                        For intLoopCnt = 1 To strAccAttributeSymbol.Length - 1
    '                                            intIndex = intIndex + 1
    '                                            strcAccessoryInfo(intIndex).AttributeSymbol = Left(strAccAttributeSymbol(intLoopCnt).Trim & Space(2), 2)
    '                                            strcAccessoryInfo(intIndex).OptionKataban = Left(strAccOptionKataban(intLoopCnt).Trim & Space(30), 30)
    '                                            strcAccessoryInfo(intIndex).Quantity = Format(intAccQuantity(intLoopCnt), "00")
    '                                        Next
    '                                    Case "06"
    '                                        intIndex = 0
    '                                        For intLoopCnt = 1 To 19
    '                                            If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
    '                                                intIndex = intIndex + 1
    '                                                strcManifoldInfo(intIndex).AttributeSymbol = Left(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim & Space(2), 2)
    '                                                strcManifoldInfo(intIndex).OptionKataban = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & Space(30), 30)
    '                                                strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                                strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                                strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                                strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                                If objKtbnStrc.strcSelection.strSeriesKataban = "LMF0" AndAlso _
    '                                                    objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "T0D" Then
    '                                                    strPositionInfo = StrReverse(strPositionInfo)
    '                                                End If
    '                                                strcManifoldInfo(intIndex).PositionInfo = Left(strPositionInfo & Space(10), 10)
    '                                                strcManifoldInfo(intIndex).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                            End If
    '                                        Next
    '                                        '付属品取得
    '                                        Call subISOAccessoryGet(objKtbnStrc, strAccAttributeSymbol, strAccOptionKataban, intAccQuantity)
    '                                        intIndex = 0
    '                                        For intLoopCnt = 1 To strAccAttributeSymbol.Length - 1
    '                                            intIndex = intIndex + 1
    '                                            strcAccessoryInfo(intIndex).AttributeSymbol = Left(strAccAttributeSymbol(intLoopCnt).Trim & Space(2), 2)
    '                                            strcAccessoryInfo(intIndex).OptionKataban = Left(strAccOptionKataban(intLoopCnt).Trim & Space(30), 30)
    '                                            strcAccessoryInfo(intIndex).Quantity = Format(intAccQuantity(intLoopCnt), "00")
    '                                        Next
    '                                End Select
    '                        End Select

    '                        '明細No.
    '                        intLineNo = intLineNo + 1

    '                        'result.clKatahikiInfoDto.RegistKey = strKeyInfo
    '                        'result.clKatahikiInfoDto.LineNo = intLineNo
    '                        'result.clKatahikiInfoDto.Kataban = strcHeader.FullKataban
    '                        'result.clKatahikiInfoDto.CheckKubun = strcHeader.KatabanCheckDiv
    '                        'result.clKatahikiInfoDto.DeliveryPlant = strcHeader.PlaceCd
    '                        'result.clKatahikiInfoDto.StorageLocation = Nothing
    '                        'result.clKatahikiInfoDto.EvaluationType = Nothing
    '                        'result.clKatahikiInfoDto.ListPrice = strcHeader.ListPrice
    '                        'result.clKatahikiInfoDto.RegistPrice = strcHeader.RegPrice
    '                        'result.clKatahikiInfoDto.SsPrice = strcHeader.SsPrice
    '                        'result.clKatahikiInfoDto.BsPrice = strcHeader.BsPrice
    '                        'result.clKatahikiInfoDto.GsPrice = strcHeader.GsPrice
    '                        'result.clKatahikiInfoDto.PsPrice = strcHeader.PsPrice
    '                        'result.clKatahikiInfoDto.NetPrice = strcHeader.NetPrice
    '                        'result.clKatahikiInfoDto.Currency = strcHeader.CurrencyCode
    '                        'result.clKatahikiInfoDto.KisyuCode = strcHeader.ModelCd.Trim

    '                        'マニホールド仕様情報
    '                        strShiyouInfo = Nothing
    '                        strShiyouInfo = strShiyouInfo & strcHeader.WiringSpecDiv
    '                        strShiyouInfo = strShiyouInfo & strcHeader.RailLength
    '                        'strShiyouInfo = strShiyouInfo & strcHeader.ProcDatetime
    '                        strShiyouInfo = strShiyouInfo & Left(objKtbnStrc.strcSelection.strFullKataban.Trim & Space(30), 30)

    '                        Dim intloop As Integer = 0
    '                        For intLoopCnt = 1 To 20
    '                            strShiyouInfo = strShiyouInfo & (strcManifoldInfo(intLoopCnt).AttributeSymbol)
    '                            strShiyouInfo = strShiyouInfo & (strcManifoldInfo(intLoopCnt).OptionKataban)
    '                            strShiyouInfo = strShiyouInfo & (strcManifoldInfo(intLoopCnt).PositionInfo)
    '                            strShiyouInfo = strShiyouInfo & (strcManifoldInfo(intLoopCnt).Quantity)
    '                            '   strShiyouInfo = strShiyouInfo & (strcManifoldInfo(intLoopCnt).OrderNo)

    '                            If strcManifoldInfo(intLoopCnt).OptionKataban.Trim.Length > 0 And CInt(strcManifoldInfo(intLoopCnt).Quantity) > 0 Then
    '                                If strcManifoldInfo(intLoopCnt).AttributeSymbol <> "GF" Then
    '                                    intloop += 1
    '                                    strShiyouInfo = strShiyouInfo & "*#*000" & intloop.ToString.PadLeft(2, "0")
    '                                Else
    '                                    strShiyouInfo = strShiyouInfo & (strcManifoldInfo(intLoopCnt).OrderNo)
    '                                End If
    '                            Else
    '                                strShiyouInfo = strShiyouInfo & (strcManifoldInfo(intLoopCnt).OrderNo)
    '                            End If

    '                        Next
    '                        For intLoopCnt = 1 To 10
    '                            strShiyouInfo = strShiyouInfo & (strcAccessoryInfo(intLoopCnt).AttributeSymbol)
    '                            strShiyouInfo = strShiyouInfo & (strcAccessoryInfo(intLoopCnt).OptionKataban)
    '                            strShiyouInfo = strShiyouInfo & (strcAccessoryInfo(intLoopCnt).Quantity)
    '                        Next

    '                        result.clKatahikiInfoDto.ManifoldSpecData = strShiyouInfo & Space(64)

    '                        result.clKatahikiInfoDto.ElKubun = strcHeader.ELDiv
    '                        If objKtbnStrc.strcSelection.strSalesUnit = "" Then
    '                            result.clKatahikiInfoDto.SalesUnit = "PC"
    '                        Else
    '                            result.clKatahikiInfoDto.SalesUnit = objKtbnStrc.strcSelection.strSalesUnit
    '                        End If

    '                        result.clKatahikiInfoDto.SapBaseUnit = objKtbnStrc.strcSelection.strSapBaseUnit
    '                        result.clKatahikiInfoDto.QuantityPerSalesUnit = objKtbnStrc.strcSelection.strQuantityPerSalesUnit
    '                        'If objKtbnStrc.strcSelection.strOrderLot = Nothing Then
    '                        '    result.clKatahikiInfoDto.OrderLot = 0
    '                        'Else
    '                        '    result.clKatahikiInfoDto.OrderLot = objKtbnStrc.strcSelection.strOrderLot
    '                        'End If

    '                        'オーダーロットに数量をセットする
    '                        result.clKatahikiInfoDto.OrderLot = strcHeader.Quantity
    '                        result.clKatahikiInfoDto.CzFlag = blCZFlag

    '                        'リストに追加
    '                        result.clKatahikiInfoDtoIso.Add(result.clKatahikiInfoDto)

    '                    End If
    '                Next
    '            Case Else
    '                '編集
    '                With strcHeader
    '                    '形番
    '                    .kataban = objKtbnStrc.strcSelection.strFullKataban

    '                    '仕様書有無区分
    '                    Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
    '                        Case "", "00"
    '                            .SpecExistsDiv = "N"
    '                            strPos = ""
    '                            .MsgPosition = Left(strPos.Trim & Space(60), 60)
    '                        Case "51", "52", "53", "54", "55", "56", "57", "58", "59", "60", _
    '                             "61", "62", "63", "64", "65", "66", "67", "68", "69", "70", _
    '                             "71", "72", "73", "74", "75", "76", "77", "78", "79", "80", _
    '                             "81", "82", "83", "84", "85", "86", "87", "88", "89", "90", _
    '                             "91", "92", "93", "A4", "A5", "A6", "A7", "A8", "98", _
    '                             "S", "T", "U"
    '                            If objOption.fncVaccumMixCheck(objKtbnStrc) Then
    '                                .SpecExistsDiv = "Y"
    '                            Else
    '                                .SpecExistsDiv = "N"
    '                            End If
    '                            If objKtbnStrc.strcSelection.strSpecNo.Trim = "A1" Or objKtbnStrc.strcSelection.strSpecNo.Trim = "A2" _
    '                               Or objKtbnStrc.strcSelection.strSpecNo.Trim = "A9" Or objKtbnStrc.strcSelection.strSpecNo.Trim = "B1" _
    '                               Or objKtbnStrc.strcSelection.strSpecNo.Trim = "B2" Or objKtbnStrc.strcSelection.strSpecNo.Trim = "B3" _
    '                               Or objKtbnStrc.strcSelection.strSpecNo.Trim = "B4" Then
    '                                .SpecExistsDiv = "Y"
    '                            End If
    '                            strPos = ""
    '                            '簡易マニホールドの判断
    '                            If KHKataban.fncJudgeSimpleSpec(objCon, objKtbnStrc, strUserID, strSessID) = True Or _
    '                               objOption.fncVaccumMixCheck(objKtbnStrc) Then
    '                                intPositionInfo = KHKataban.fncGetMixManifoldInfo(objCon, objKtbnStrc, strUserID, strSessID)
    '                                objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullKataban & CdCst.Sign.Hypen
    '                                For intLoopPos2 = 1 To UBound(intPositionInfo)
    '                                    objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & ClsCommon.fncPositionChance(intPositionInfo(intLoopPos2))
    '                                Next
    '                                If InStr(1, objKtbnStrc.strcSelection.strFullKataban, "-ST") <> 0 Then
    '                                    objKtbnStrc.strcSelection.strFullManiKataban = Replace(objKtbnStrc.strcSelection.strFullManiKataban, "-ST", "")
    '                                    objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & "-ST"
    '                                End If

    '                                .kataban = objKtbnStrc.strcSelection.strFullManiKataban

    '                                '受注形番1,受注形番2
    '                                If objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length = 30 Then
    '                                    .Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim
    '                                    .Kataban2 = Space(30)
    '                                ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length < 30 Then
    '                                    .Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim & _
    '                                                Space(30 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
    '                                    .Kataban2 = Space(30)
    '                                ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length > 30 Then
    '                                    .Kataban1 = Left(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 30)
    '                                    .Kataban2 = Mid(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 31) & _
    '                                                Space(60 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
    '                                End If

    '                                'メッセージ欄に出す位置情報を作成
    '                                If objOption.fncVaccumMixCheck(objKtbnStrc) = False Then
    '                                    intLoopPos = 1
    '                                    While objKtbnStrc.strcSelection.strOptionKataban(intLoopPos).Trim.Length <> 0
    '                                        '記号設定
    '                                        If objKtbnStrc.strcSelection.intQuantity(intLoopPos) = 0 Then
    '                                        Else
    '                                            If strPos = "" Then
    '                                                strPos = objKtbnStrc.strcSelection.strAttributeSymbol(intLoopPos).Trim & "="
    '                                            Else
    '                                                strPos = strPos & "," & objKtbnStrc.strcSelection.strAttributeSymbol(intLoopPos).Trim & "="
    '                                            End If

    '                                            '位置情報設定
    '                                            For intMPos = 1 To 50 Step 2
    '                                                If Mid(objKtbnStrc.strcSelection.strPositionInfo(intLoopPos).Trim, intMPos, 1) = 1 Then
    '                                                    intMPos2 = Int(intMPos / 2) + 1
    '                                                    Select Case Right(strPos, 1)
    '                                                        Case (intMPos2 - 1)
    '                                                            strPos = strPos & "-"
    '                                                        Case "-"
    '                                                            If intRenPos <> intMPos2 - 1 Then
    '                                                                strPos = strPos & intMPos2 - 1 & "," & intMPos2
    '                                                            Else
    '                                                            End If
    '                                                        Case "="
    '                                                            strPos = strPos & intMPos2
    '                                                        Case Else
    '                                                            strPos = strPos & "," & intMPos2
    '                                                    End Select
    '                                                    intRenPos = intMPos2
    '                                                End If
    '                                            Next
    '                                            If Right(strPos, 1) = "-" Then
    '                                                strPos = strPos & intRenPos
    '                                            End If
    '                                        End If
    '                                        intLoopPos = intLoopPos + 1
    '                                    End While
    '                                End If

    '                            End If
    '                            .MsgPosition = Left(strPos.Trim & Space(60), 60)
    '                        Case "12", "18", "19", "20", "21", "22", "23"
    '                            If objOption.fncVaccumMixCheck(objKtbnStrc) Then
    '                                .SpecExistsDiv = "Y"
    '                            Else
    '                                .SpecExistsDiv = "N"
    '                            End If
    '                            strPos = ""
    '                            .MsgPosition = Left(strPos.Trim & Space(60), 60)
    '                        Case "17"
    '                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "X" Then
    '                                .SpecExistsDiv = "Y"
    '                            Else
    '                                .SpecExistsDiv = "N"
    '                            End If
    '                            strPos = ""
    '                            .MsgPosition = Left(strPos.Trim & Space(60), 60)

    '                        Case "02", "03", "08"
    '                            If KHKataban.fncJudgeSimpleSpec(objCon, objKtbnStrc, strUserID, strSessID) = True Or _
    '                             objOption.fncVaccumMixCheck(objKtbnStrc) Then
    '                                intPositionInfo = KHKataban.fncGetMixManifoldInfo(objCon, objKtbnStrc, strUserID, strSessID)
    '                                objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullKataban & CdCst.Sign.Hypen
    '                                For intLoopPos2 = 1 To UBound(intPositionInfo)
    '                                    objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & ClsCommon.fncPositionChance(intPositionInfo(intLoopPos2))
    '                                Next
    '                                If InStr(1, objKtbnStrc.strcSelection.strFullKataban, "-ST") <> 0 Then
    '                                    objKtbnStrc.strcSelection.strFullManiKataban = Replace(objKtbnStrc.strcSelection.strFullManiKataban, "-ST", "")
    '                                    objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & "-ST"
    '                                End If

    '                                .kataban = objKtbnStrc.strcSelection.strFullManiKataban
    '                                .SpecExistsDiv = "Y"
    '                                strPos = ""
    '                                .MsgPosition = Left(strPos.Trim & Space(60), 60)
    '                            End If

    '                        Case Else
    '                            .SpecExistsDiv = "Y"
    '                            strPos = ""
    '                            .MsgPosition = Left(strPos.Trim & Space(60), 60)
    '                    End Select

    '                    'FOB対応 
    '                    'If strFobPrice = Nothing Then
    '                    '    .NetPrice = 0
    '                    'Else
    '                    '    .NetPrice = strFobPrice
    '                    'End If

    '                    'FOB対応
    '                    Dim decFobPrice As Decimal
    '                    decFobPrice = IIf(Decimal.TryParse(strFobPrice, decFobPrice), decFobPrice, 0)
    '                    If decFobPrice - Fix(decFobPrice) = 0 Then
    '                        .NetPrice = Decimal.ToInt32(decFobPrice)
    '                    Else
    '                        .NetPrice = Format(decFobPrice, "#.00")
    '                    End If

    '                    '定価
    '                    .ListPrice = Decimal.ToInt32(objKtbnStrc.strcSelection.intListPrice)
    '                    '登録店価格
    '                    .RegPrice = Decimal.ToInt32(objKtbnStrc.strcSelection.intRegPrice)
    '                    'SS価格
    '                    .SsPrice = Decimal.ToInt32(objKtbnStrc.strcSelection.intSsPrice)
    '                    'BS価格
    '                    .BsPrice = Decimal.ToInt32(objKtbnStrc.strcSelection.intBsPrice)
    '                    'GS価格
    '                    .GsPrice = Decimal.ToInt32(objKtbnStrc.strcSelection.intGsPrice)
    '                    'PS価格
    '                    .PsPrice = Decimal.ToInt32(objKtbnStrc.strcSelection.intPsPrice)

    '                    '通貨コード
    '                    If .NetPrice > 0 Then
    '                        .CurrencyCode = strCurrencyCode
    '                    Else
    '                        .CurrencyCode = Nothing
    '                    End If

    '                    '形番チェック区分
    '                    .KatabanCheckDiv = "Z" & Left(objKtbnStrc.strcSelection.strKatabanCheckDiv.Trim & Space(1), 1) &
    '                        IIf(objKtbnStrc.strcSelection.strCostCalcNo.Equals(""), Nothing, ("(" & objKtbnStrc.strcSelection.strCostCalcNo & ")"))
    '                    '出荷場所
    '                    .PlaceCd = Left(objKtbnStrc.strcSelection.strPlaceCd.Trim & Space(4), 4)
    '                    '機種コード
    '                    .ModelCd = Left(objKtbnStrc.strcSelection.strModelNo.Trim & Space(2), 2)

    '                    '配線仕様有無区分()
    '                    .WiringSpecDiv = LSet(objKtbnStrc.strcSelection.strWiringSpec.Trim, intSpaceCnt_01)
    '                    'レール長さ()
    '                    .RailLength = Format(objKtbnStrc.strcSelection.decDinRailLength * 100, "000000")

    '                    '処理日付＆マニホールド代表形番
    '                    Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
    '                        Case "", "00"
    '                            .ProcDatetime = ""
    '                            .FullKataban = ""
    '                        Case Else
    '                            Select Case objKtbnStrc.strcSelection.strSeriesKataban
    '                                Case "CMF", "LMF0"
    '                                    .ProcDatetime = Format(Now, "MMddhhmmss")
    '                                    .FullKataban = Left(objKtbnStrc.strcSelection.strFullKataban.Trim & Space(30), 30)
    '                                Case Else
    '                                    'その他マニホールドの場合
    '                                    .ProcDatetime = ""
    '                                    .FullKataban = ""
    '                            End Select
    '                    End Select

    '                    'EL品判定区分
    '                    If objKataban.fncELKatabanCheck(objCon, objKtbnStrc.strcSelection.strFullKataban.Trim, "1") Then
    '                        .ELDiv = True
    '                    Else
    '                        .ELDiv = False
    '                    End If

    '                End With
    '                '仕様書情報
    '                Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
    '                    Case "01", "02", "03", "04", "07", _
    '                         "08", "10", "11", "13", "14", _
    '                         "15", "16", "17", "96", "A1", "A2", _
    '                         "51", "53", "54", "55", "56", "57", _
    '                         "58", "59", "60", "61", "62", "63", _
    '                         "64", "65", "66", "67", "68", "69", _
    '                         "70", "71", "72", "73", "74", "75", _
    '                         "76", "77", "78", "79", "80", "81", _
    '                         "82", "83", "84", "85", "86", "87", _
    '                         "88", "89", "91", "92", "93", "A4", "A5", "A6", "A7", "A8", "A9", "B1", "98", _
    '                         "S", "T", "U", "B2", "B3", "B4"
    '                        For intLoopCnt = 1 To intLoopMax_01
    '                            With strcManifoldInfo(intLoopCnt)
    '                                .AttributeSymbol = Space(intSpaceCnt_03)
    '                                .OptionKataban = Space(intSpaceCnt_04)
    '                                .PositionInfo = Space(intSpaceCnt_05)
    '                                .Quantity = "00"
    '                                .OrderNo = ""
    '                            End With
    '                        Next
    '                        For intLoopCnt = 1 To intLoopMax_02
    '                            With strcAccessoryInfo(intLoopCnt)
    '                                .AttributeSymbol = Space(intSpaceCnt_03)
    '                                .OptionKataban = Space(intSpaceCnt_04)
    '                                .Quantity = "00"
    '                            End With
    '                        Next
    '                        Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
    '                            Case "01"
    '                                Dim intNo As Integer = 20      '表の行数（設置位置№が指定できる行数）
    '                                For intLoopCnt = 1 To intNo
    '                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
    '                                       objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
    '                                        Select Case intLoopCnt
    '                                            Case 3
    '                                                strTmpKataban = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                                strTmpPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                                strTmpPositionInfo = Replace(strTmpPositionInfo, ",", "")
    '                                                strTmpPositionInfo = Replace(strTmpPositionInfo, "0", " ")
    '                                                strTmpPositionInfo = Replace(strTmpPositionInfo, "1", "Y")
    '                                            Case 4 To 12
    '                                                If strTmpKataban.Trim = "" Then
    '                                                    intIndex = intIndex + 1
    '                                                    strcManifoldInfo(intIndex).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                    strcManifoldInfo(intIndex).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                                    strcManifoldInfo(intIndex).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                                    strcManifoldInfo(intIndex).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                                Else
    '                                                    intTmpQuantity = 0
    '                                                    strTmpPositionInfo1 = ""
    '                                                    strTmpPositionInfo2 = ""
    '                                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")

    '                                                    For intLoopCnt1 = 1 To strPositionInfo.Length - 1
    '                                                        If Mid(strTmpPositionInfo, intLoopCnt1, 1) = "Y" Then
    '                                                            If Mid(strPositionInfo, intLoopCnt1, 1) = "Y" Then
    '                                                                strTmpPositionInfo1 = strTmpPositionInfo1 & Mid(strTmpPositionInfo, intLoopCnt1, 1)
    '                                                                strTmpPositionInfo2 = strTmpPositionInfo2 & " "
    '                                                                intTmpQuantity = intTmpQuantity + 1
    '                                                            Else
    '                                                                strTmpPositionInfo1 = strTmpPositionInfo1 & " "
    '                                                                strTmpPositionInfo2 = strTmpPositionInfo2 & Mid(strPositionInfo, intLoopCnt1, 1)
    '                                                            End If
    '                                                        Else
    '                                                            strTmpPositionInfo1 = strTmpPositionInfo1 & Mid(strTmpPositionInfo, intLoopCnt1, 1)
    '                                                            strTmpPositionInfo2 = strTmpPositionInfo2 & Mid(strPositionInfo, intLoopCnt1, 1)
    '                                                        End If
    '                                                    Next
    '                                                    If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) - intTmpQuantity > 0 Then
    '                                                        intIndex = intIndex + 1
    '                                                        strcManifoldInfo(intIndex).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                        strcManifoldInfo(intIndex).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                                        strcManifoldInfo(intIndex).PositionInfo = LSet(strTmpPositionInfo2, intSpaceCnt_05)
    '                                                        strcManifoldInfo(intIndex).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt) - intTmpQuantity, "00")
    '                                                    End If
    '                                                    If intTmpQuantity > 0 Then
    '                                                        intIndex = intIndex + 1
    '                                                        strcManifoldInfo(intIndex).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                        strcManifoldInfo(intIndex).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & strTmpKataban.Trim, intSpaceCnt_04)
    '                                                        strcManifoldInfo(intIndex).PositionInfo = LSet(strTmpPositionInfo1, intSpaceCnt_05)
    '                                                        strcManifoldInfo(intIndex).Quantity = Format(intTmpQuantity, "00")
    '                                                    End If
    '                                                End If
    '                                            Case Else
    '                                                intIndex = intIndex + 1
    '                                                strcManifoldInfo(intIndex).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcManifoldInfo(intIndex).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                                strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                                strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                                strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                                strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                                strcManifoldInfo(intIndex).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                                strcManifoldInfo(intIndex).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                        End Select
    '                                    End If
    '                                Next
    '                                '21行目以降の処理
    '                                For intLoopCnt = 1 To 10
    '                                    intNo = intNo + 1       '現在の行
    '                                    Select Case intLoopCnt
    '                                        Case 1 To 4
    '                                            With strcAccessoryInfo(intLoopCnt)
    '                                                .AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intNo).Trim, intSpaceCnt_03)
    '                                                .OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intNo).Trim, intSpaceCnt_04)
    '                                                .Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intNo), "00")
    '                                            End With
    '                                        Case 5 To 8
    '                                            With strcAccessoryInfo(intLoopCnt)
    '                                                Select Case objKtbnStrc.strcSelection.strOptionKataban(intNo).Trim
    '                                                    Case CdCst.Manifold.InspReportJp.SelectValue
    '                                                        .AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intNo).Trim, intSpaceCnt_03)
    '                                                        .OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
    '                                                        .Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intNo), "00")
    '                                                    Case CdCst.Manifold.InspReportEn.SelectValue
    '                                                        .AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intNo).Trim, intSpaceCnt_03)
    '                                                        .OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
    '                                                        .Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intNo), "00")
    '                                                    Case Else
    '                                                        .AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intNo).Trim, intSpaceCnt_03)

    '                                                        If Left(objKtbnStrc.strcSelection.strOptionKataban(intNo).Trim & Space(9), 9) = "検査成績書（英文）" Then

    '                                                            .OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)

    '                                                        ElseIf Left(objKtbnStrc.strcSelection.strOptionKataban(intNo).Trim & Space(9), 9) = "検査成績書（和文）" Then
    '                                                            .OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
    '                                                        Else
    '                                                            .OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intNo).Trim, intSpaceCnt_04)
    '                                                        End If
    '                                                        .Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intNo), "00")
    '                                                End Select
    '                                            End With
    '                                        Case 9
    '                                            With strcAccessoryInfo(intLoopCnt)
    '                                                If objKtbnStrc.strcSelection.strOptionKataban(intNo).Trim = "1" Then
    '                                                    .AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intNo).Trim, intSpaceCnt_03)
    '                                                    .OptionKataban = Space(intSpaceCnt_04)
    '                                                    .Quantity = Format(0, "00")
    '                                                Else
    '                                                    .AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intNo).Trim, intSpaceCnt_03)
    '                                                    .OptionKataban = LSet(CdCst.Manifold.TubeRemover.DummyValue, intSpaceCnt_04)
    '                                                    .Quantity = Format(1, "00")
    '                                                End If
    '                                            End With
    '                                    End Select
    '                                Next
    '                            Case "02"
    '                                For intLoopCnt = 1 To 16
    '                                    strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                    strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                For intLoopCnt = 17 To 21
    '                                    strcAccessoryInfo(intLoopCnt - 16).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                        Case CdCst.Manifold.InspReportJp.SelectValue
    '                                            strcAccessoryInfo(intLoopCnt - 16).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
    '                                        Case CdCst.Manifold.InspReportEn.SelectValue
    '                                            strcAccessoryInfo(intLoopCnt - 16).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
    '                                        Case Else
    '                                            strcAccessoryInfo(intLoopCnt - 16).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    End Select
    '                                    strcAccessoryInfo(intLoopCnt - 16).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next

    '                                If objKtbnStrc.strcSelection.strOpSymbol(1).PadRight(2, " ").Substring(0, 2) = "80" Then
    '                                    intPositionInfo = KHKataban.fncGetMixManifoldInfo(objCon, objKtbnStrc, strUserID, strSessID)
    '                                    objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullKataban & CdCst.Sign.Hypen
    '                                    For intLoopPos2 = 1 To UBound(intPositionInfo)
    '                                        objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & ClsCommon.fncPositionChance(intPositionInfo(intLoopPos2))
    '                                    Next
    '                                    If InStr(1, objKtbnStrc.strcSelection.strFullKataban, "-ST") <> 0 Then
    '                                        objKtbnStrc.strcSelection.strFullManiKataban = Replace(objKtbnStrc.strcSelection.strFullManiKataban, "-ST", "")
    '                                        objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & "-ST"
    '                                    End If
    '                                    '受注形番1,受注形番2
    '                                    If objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length = 30 Then
    '                                        strcHeader.Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim
    '                                        strcHeader.Kataban2 = Space(30)
    '                                    ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length < 30 Then
    '                                        strcHeader.Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim & _
    '                                                    Space(30 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
    '                                        strcHeader.Kataban2 = Space(30)
    '                                    ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length > 30 Then
    '                                        strcHeader.Kataban1 = Left(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 30)
    '                                        strcHeader.Kataban2 = Mid(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 31) & _
    '                                                    Space(60 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
    '                                    End If
    '                                End If

    '                            Case "03"   '機種　M
    '                                For intLoopCnt = 1 To 15
    '                                    strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    If objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim = "" Then
    '                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    Else
    '                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
    '                                                                                          objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim & _
    '                                                                                          objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    End If
    '                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                    strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                For intLoopCnt = 16 To 23
    '                                    If intLoopCnt <> 23 Then
    '                                        strcAccessoryInfo(intLoopCnt - 15).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                            Case CdCst.Manifold.InspReportJp.SelectValue
    '                                                strcAccessoryInfo(intLoopCnt - 15).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
    '                                            Case CdCst.Manifold.InspReportEn.SelectValue
    '                                                strcAccessoryInfo(intLoopCnt - 15).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
    '                                            Case Else
    '                                                strcAccessoryInfo(intLoopCnt - 15).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                        End Select
    '                                        strcAccessoryInfo(intLoopCnt - 15).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                    Else
    '                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                            Case CdCst.Manifold.TubeRemover.Necessity
    '                                                strcAccessoryInfo(intLoopCnt - 15).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcAccessoryInfo(intLoopCnt - 15).OptionKataban = Space(intSpaceCnt_04)
    '                                                strcAccessoryInfo(intLoopCnt - 15).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                            Case CdCst.Manifold.TubeRemover.UnNecessity
    '                                                strcAccessoryInfo(intLoopCnt - 15).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcAccessoryInfo(intLoopCnt - 15).OptionKataban = LSet(CdCst.Manifold.TubeRemover.DummyValue, intSpaceCnt_04)
    '                                                strcAccessoryInfo(intLoopCnt - 15).Quantity = Format(1, "00")
    '                                            Case Else
    '                                                strcAccessoryInfo(intLoopCnt - 15).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcAccessoryInfo(intLoopCnt - 15).OptionKataban = Space(intSpaceCnt_04)
    '                                                strcAccessoryInfo(intLoopCnt - 15).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                        End Select
    '                                    End If
    '                                Next

    '                                If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "8" Then
    '                                    intPositionInfo = KHKataban.fncGetMixManifoldInfo(objCon, objKtbnStrc, strUserID, strSessID)
    '                                    objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullKataban & CdCst.Sign.Hypen
    '                                    For intLoopPos2 = 1 To UBound(intPositionInfo)
    '                                        objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & ClsCommon.fncPositionChance(intPositionInfo(intLoopPos2))
    '                                    Next
    '                                    If InStr(1, objKtbnStrc.strcSelection.strFullKataban, "-ST") <> 0 Then
    '                                        objKtbnStrc.strcSelection.strFullManiKataban = Replace(objKtbnStrc.strcSelection.strFullManiKataban, "-ST", "")
    '                                        objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & "-ST"
    '                                    End If
    '                                    '受注形番1,受注形番2
    '                                    If objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length = 30 Then
    '                                        strcHeader.Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim
    '                                        strcHeader.Kataban2 = Space(30)
    '                                    ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length < 30 Then
    '                                        strcHeader.Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim & _
    '                                                    Space(30 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
    '                                        strcHeader.Kataban2 = Space(30)
    '                                    ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length > 30 Then
    '                                        strcHeader.Kataban1 = Left(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 30)
    '                                        strcHeader.Kataban2 = Mid(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 31) & _
    '                                                    Space(60 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
    '                                    End If
    '                                End If

    '                                'RM1803032_スペーサ行追加対応
    '                            Case "04"
    '                                Dim intManiEnd As Integer = 16
    '                                For intLoopCnt = 1 To intManiEnd
    '                                    strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    If objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim = "" Then
    '                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    Else
    '                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
    '                                                                                          objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).Trim & _
    '                                                                                          objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    End If
    '                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                    strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                For intLoopCnt = 17 To 25
    '                                    If intLoopCnt <> 25 Then
    '                                        strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                            Case CdCst.Manifold.InspReportJp.SelectValue
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
    '                                            Case CdCst.Manifold.InspReportEn.SelectValue
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
    '                                            Case Else
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                        End Select
    '                                        strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                    Else
    '                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                            Case CdCst.Manifold.TubeRemover.Necessity
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = Space(intSpaceCnt_04)
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                            Case CdCst.Manifold.TubeRemover.UnNecessity
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.TubeRemover.DummyValue, intSpaceCnt_04)
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(1, "00")
    '                                            Case Else
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = Space(intSpaceCnt_04)
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                        End Select
    '                                    End If
    '                                Next
    '                                'RM1803032_スペーサ行追加対応
    '                            Case "07", "96"
    '                                Dim intManiEnd As Integer = 21
    '                                For intLoopCnt = 1 To intManiEnd
    '                                    strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                    strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                For intLoopCnt = 22 To 28
    '                                    If intLoopCnt <> 28 Then
    '                                        strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                            Case CdCst.Manifold.InspReportJp.SelectValue
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
    '                                            Case CdCst.Manifold.InspReportEn.SelectValue
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
    '                                            Case Else
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                        End Select
    '                                        strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                    Else
    '                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                            Case CdCst.Manifold.TubeRemover.Necessity
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = Space(intSpaceCnt_04)
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                            Case CdCst.Manifold.TubeRemover.UnNecessity
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.TubeRemover.DummyValue, intSpaceCnt_04)
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(1, "00")
    '                                            Case Else
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = Space(intSpaceCnt_04)
    '                                                strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                        End Select
    '                                    End If
    '                                Next
    '                                '2018/03/08_タグ銘板設定時値セット
    '                                If strcAccessoryInfo(6).AttributeSymbol = "T6" And strcAccessoryInfo(6).Quantity <> "00" Then
    '                                    strcAccessoryInfo(8).AttributeSymbol = "L1"
    '                                    Dim main As New _Main
    '                                    strcAccessoryInfo(8).OptionKataban = LSet(main.Session("decDinRailLength").ToString, intSpaceCnt_04)
    '                                    strcAccessoryInfo(8).Quantity = "01"
    '                                End If
    '                            Case "08"
    '                                For intLoopCnt = 1 To 16
    '                                    strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                    strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                For intLoopCnt = 17 To 21
    '                                    strcAccessoryInfo(intLoopCnt - 16).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                        Case CdCst.Manifold.InspReportJp.SelectValue
    '                                            strcAccessoryInfo(intLoopCnt - 16).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
    '                                        Case CdCst.Manifold.InspReportEn.SelectValue
    '                                            strcAccessoryInfo(intLoopCnt - 16).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
    '                                        Case Else
    '                                            strcAccessoryInfo(intLoopCnt - 16).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    End Select
    '                                    strcAccessoryInfo(intLoopCnt - 16).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "8" Then
    '                                    intPositionInfo = KHKataban.fncGetMixManifoldInfo(objCon, objKtbnStrc, strUserID, strSessID)
    '                                    objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullKataban & CdCst.Sign.Hypen
    '                                    For intLoopPos2 = 1 To UBound(intPositionInfo)
    '                                        objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & ClsCommon.fncPositionChance(intPositionInfo(intLoopPos2))
    '                                    Next
    '                                    If InStr(1, objKtbnStrc.strcSelection.strFullKataban, "-ST") <> 0 Then
    '                                        objKtbnStrc.strcSelection.strFullManiKataban = Replace(objKtbnStrc.strcSelection.strFullManiKataban, "-ST", "")
    '                                        objKtbnStrc.strcSelection.strFullManiKataban = objKtbnStrc.strcSelection.strFullManiKataban & "-ST"
    '                                    End If
    '                                    '受注形番1,受注形番2
    '                                    If objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length = 30 Then
    '                                        strcHeader.Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim
    '                                        strcHeader.Kataban2 = Space(30)
    '                                    ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length < 30 Then
    '                                        strcHeader.Kataban1 = objKtbnStrc.strcSelection.strFullManiKataban.Trim & _
    '                                                    Space(30 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
    '                                        strcHeader.Kataban2 = Space(30)
    '                                    ElseIf objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length > 30 Then
    '                                        strcHeader.Kataban1 = Left(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 30)
    '                                        strcHeader.Kataban2 = Mid(objKtbnStrc.strcSelection.strFullManiKataban.Trim, 31) & _
    '                                                    Space(60 - objKtbnStrc.strcSelection.strFullManiKataban.Trim.Length)
    '                                    End If
    '                                End If
    '                            Case "10"
    '                                For intLoopCnt = 1 To 14
    '                                    strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                    strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                For intLoopCnt = 15 To 23
    '                                    If intLoopCnt <> 23 Then
    '                                        strcAccessoryInfo(intLoopCnt - 14).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                            Case CdCst.Manifold.InspReportJp.SelectValue
    '                                                strcAccessoryInfo(intLoopCnt - 14).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
    '                                            Case CdCst.Manifold.InspReportEn.SelectValue
    '                                                strcAccessoryInfo(intLoopCnt - 14).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
    '                                            Case Else
    '                                                strcAccessoryInfo(intLoopCnt - 14).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                        End Select
    '                                        strcAccessoryInfo(intLoopCnt - 14).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                    Else
    '                                        Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                            Case CdCst.Manifold.TubeRemover.Necessity
    '                                                strcAccessoryInfo(intLoopCnt - 14).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcAccessoryInfo(intLoopCnt - 14).OptionKataban = Space(intSpaceCnt_04)
    '                                                strcAccessoryInfo(intLoopCnt - 14).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                            Case CdCst.Manifold.TubeRemover.UnNecessity
    '                                                strcAccessoryInfo(intLoopCnt - 14).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcAccessoryInfo(intLoopCnt - 14).OptionKataban = LSet(CdCst.Manifold.TubeRemover.DummyValue, intSpaceCnt_04)
    '                                                strcAccessoryInfo(intLoopCnt - 14).Quantity = Format(1, "00")
    '                                            Case Else
    '                                                strcAccessoryInfo(intLoopCnt - 14).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                                strcAccessoryInfo(intLoopCnt - 14).OptionKataban = Space(intSpaceCnt_04)
    '                                                strcAccessoryInfo(intLoopCnt - 14).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                        End Select
    '                                    End If
    '                                Next
    '                            Case "11"
    '                                For intLoopCnt = 1 To 15
    '                                    strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                    strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                For intLoopCnt = 16 To 18
    '                                    strcAccessoryInfo(intLoopCnt - 15).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    strcAccessoryInfo(intLoopCnt - 15).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    strcAccessoryInfo(intLoopCnt - 15).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                            Case "13"
    '                                For intLoopCnt = 1 To 17
    '                                    strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                    strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                For intLoopCnt = 18 To 24
    '                                    strcAccessoryInfo(intLoopCnt - 17).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                        Case CdCst.Manifold.InspReportJp.SelectValue
    '                                            strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
    '                                        Case CdCst.Manifold.InspReportEn.SelectValue
    '                                            strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
    '                                        Case Else
    '                                            strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    End Select
    '                                    strcAccessoryInfo(intLoopCnt - 17).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                            Case "14"
    '                                For intLoopCnt = 1 To 6
    '                                    strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                    strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                For intLoopCnt = 7 To 9
    '                                    strcAccessoryInfo(intLoopCnt - 6).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    strcAccessoryInfo(intLoopCnt - 6).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    strcAccessoryInfo(intLoopCnt - 6).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                'RM1803032_スペーサ行追加対応
    '                            Case "15"
    '                                Dim intManiEnd As Integer = 21
    '                                For intLoopCnt = 1 To intManiEnd
    '                                    strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                    strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                For intLoopCnt = 22 To 29
    '                                    strcAccessoryInfo(intLoopCnt - intManiEnd).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                        Case CdCst.Manifold.InspReportJp.SelectValue
    '                                            strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
    '                                        Case CdCst.Manifold.InspReportEn.SelectValue
    '                                            strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
    '                                        Case Else
    '                                            strcAccessoryInfo(intLoopCnt - intManiEnd).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    End Select
    '                                    strcAccessoryInfo(intLoopCnt - intManiEnd).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                            Case "16"
    '                                For intLoopCnt = 1 To 20
    '                                    strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)

    '                                    Select Case intLoopCnt
    '                                        Case 19
    '                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                        Case 20
    '                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                        Case Else
    '                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    End Select
    '                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                    strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                    strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                                For intLoopCnt = 21 To 25
    '                                    strcAccessoryInfo(intLoopCnt - 20).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                        Case CdCst.Manifold.InspReportJp.SelectValue
    '                                            strcAccessoryInfo(intLoopCnt - 20).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
    '                                        Case CdCst.Manifold.InspReportEn.SelectValue
    '                                            strcAccessoryInfo(intLoopCnt - 20).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
    '                                        Case Else
    '                                            strcAccessoryInfo(intLoopCnt - 20).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                    End Select
    '                                    strcAccessoryInfo(intLoopCnt - 20).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                Next
    '                            Case "17"
    '                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "X" Then Exit Select
    '                                For intLoopCnt = 1 To 5
    '                                    If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
    '                                        strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                        strOpArray = Split(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, CdCst.Sign.Comma)
    '                                        For intIndex = 0 To strOpArray.Length - 1
    '                                            strcManifoldInfo(intLoopCnt).OptionKataban = strcManifoldInfo(intLoopCnt).OptionKataban.Trim & _
    '                                                                                         strOpArray(intIndex).Trim
    '                                        Next
    '                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(KHKataban.fncHypenCut(strcManifoldInfo(intLoopCnt).OptionKataban), intSpaceCnt_04)
    '                                        strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                        strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                        strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                        strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                        strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                    End If
    '                                Next
    '                                strcAccessoryInfo(1).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(6).Trim, intSpaceCnt_03)
    '                                strOpArray = Split(objKtbnStrc.strcSelection.strOptionKataban(6).Trim, CdCst.Sign.Comma)
    '                                For intIndex = 0 To strOpArray.Length - 1
    '                                    strcAccessoryInfo(1).OptionKataban = strcAccessoryInfo(1).OptionKataban.Trim & _
    '                                                                         strOpArray(intIndex).Trim
    '                                Next
    '                                strcAccessoryInfo(1).OptionKataban = LSet(KHKataban.fncHypenCut(strcAccessoryInfo(1).OptionKataban), intSpaceCnt_04)
    '                                strcAccessoryInfo(1).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(6), "00")
    '                            Case "A1", "A2", "51", "53", "54", "55", "56", "57", "58", "59", "60", "61", "62", "63", "64", _
    '                                 "65", "66", "67", "68", "69", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", _
    '                                 "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "91", "92", "93", "A4", "A5", "A6", "A7", "A8", "A9", "B1", _
    '                                 "S", "T", "U", "B2", "B3", "B4"
    '                                If objOption.fncVaccumMixCheck(objKtbnStrc) Then
    '                                    Dim str() As String = Nothing
    '                                    Dim dtSpecItem As New DataTable
    '                                    Dim dtContent As New DataTable
    '                                    Call KHManifold.subInitTable(dtSpecItem, dtContent)
    '                                    Call SiyouDAL.subSQL_ItemMst(objCon, objKtbnStrc.strcSelection.strSpecNo.Trim, dtSpecItem, dtContent)
    '                                    Dim listResult As ArrayList = KHManifold.fncGetNewKataban(dtSpecItem, dtContent, objKtbnStrc.strcSelection.strSpecNo.Trim, _
    '                                                                    objKtbnStrc.strcSelection.strSeriesKataban, objKtbnStrc.strcSelection.strOpSymbol, objKtbnStrc.strcSelection.strKeyKataban)
    '                                    For intLoopCnt = 1 To listResult.Count
    '                                        str = listResult(intLoopCnt - 1).ToString.Split("_")
    '                                        If str.Length >= 4 Then
    '                                            strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(str(1).Trim, intSpaceCnt_04)
    '                                            strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                            strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                            strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                            strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                            strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                            strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                        End If
    '                                    Next
    '                                    For intLoopCnt = listResult.Count + 1 To 12
    '                                        If objKtbnStrc.strcSelection.strAttributeSymbol.Length <= intLoopCnt Then
    '                                            strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet("", intSpaceCnt_03)
    '                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet("", intSpaceCnt_04)
    '                                            strPositionInfo = ""
    '                                            strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                            strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                            strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                            strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                            strcManifoldInfo(intLoopCnt).Quantity = Format(0, "00")
    '                                        Else
    '                                            strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                            strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                            strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                            strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                            strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                            strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                            strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                            strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                        End If

    '                                    Next
    '                                End If
    '                        End Select
    '                    Case "05", "06"
    '                        '初期化
    '                        For intLoopCnt = 1 To 20
    '                            With strcManifoldInfo(intLoopCnt)
    '                                .AttributeSymbol = Space(2)
    '                                .OptionKataban = Space(20)
    '                                .PositionInfo = Space(10)
    '                                .Quantity = "00"
    '                                .OrderNo = Space(8)
    '                            End With
    '                        Next
    '                        For intLoopCnt = 1 To 10
    '                            With strcAccessoryInfo(intLoopCnt)
    '                                .AttributeSymbol = Space(2)
    '                                .OptionKataban = Space(30)
    '                                .Quantity = "00"
    '                            End With
    '                        Next
    '                        '設定
    '                        Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
    '                            Case "05"
    '                            Case "06"
    '                        End Select
    '                    Case "09"
    '                        '初期化
    '                        For intLoopCnt = 1 To intLoopMax_01
    '                            With strcManifoldInfo(intLoopCnt)
    '                                .AttributeSymbol = Space(intSpaceCnt_03)
    '                                .OptionKataban = Space(intSpaceCnt_04)
    '                                .PositionInfo = Space(intSpaceCnt_05)
    '                                .Quantity = "00"
    '                                .OrderNo = ""
    '                            End With
    '                        Next
    '                        For intLoopCnt = 1 To intLoopMax_02
    '                            With strcAccessoryInfo(intLoopCnt)
    '                                .AttributeSymbol = Space(intSpaceCnt_03)
    '                                .OptionKataban = Space(intSpaceCnt_04)
    '                                .Quantity = "00"
    '                            End With
    '                        Next
    '                        '設定
    '                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
    '                            For intLoopCnt = 1 To 17
    '                                strcManifoldInfo(intLoopCnt).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                Select Case intLoopCnt
    '                                    Case 16
    '                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "P", intSpaceCnt_04)
    '                                    Case 17
    '                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "R", intSpaceCnt_04)
    '                                    Case Else
    '                                        strcManifoldInfo(intLoopCnt).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                End Select
    '                                strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                strcManifoldInfo(intLoopCnt).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                Select Case intLoopCnt
    '                                    Case 17
    '                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2, "00")
    '                                    Case Else
    '                                        strcManifoldInfo(intLoopCnt).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                End Select
    '                            Next
    '                            For intLoopCnt = 18 To 23
    '                                strcAccessoryInfo(intLoopCnt - 17).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
    '                                    Case CdCst.Manifold.InspReportJp.SelectValue
    '                                        strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet(CdCst.Manifold.InspReportJp.DummyValue, intSpaceCnt_04)
    '                                    Case CdCst.Manifold.InspReportEn.SelectValue
    '                                        strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet(CdCst.Manifold.InspReportEn.DummyValue, intSpaceCnt_04)
    '                                    Case Else
    '                                        Select Case intLoopCnt
    '                                            Case 20
    '                                                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4TB3" Then
    '                                                    strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet("R1/4", intSpaceCnt_04)
    '                                                Else
    '                                                    strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet("R3/8", intSpaceCnt_04)
    '                                                End If
    '                                            Case 21
    '                                                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4TB3" Then
    '                                                    strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet("R3/8", intSpaceCnt_04)
    '                                                Else
    '                                                    strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet("R1/2", intSpaceCnt_04)
    '                                                End If
    '                                            Case Else
    '                                                strcAccessoryInfo(intLoopCnt - 17).OptionKataban = LSet(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, intSpaceCnt_04)
    '                                        End Select
    '                                End Select
    '                                strcAccessoryInfo(intLoopCnt - 17).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                            Next
    '                        End If
    '                    Case "12", "18", "19", "20", "21", "22", "23"
    '                        '初期化
    '                        For intLoopCnt = 1 To intLoopMax_01
    '                            With strcManifoldInfo(intLoopCnt)
    '                                .AttributeSymbol = Space(intSpaceCnt_03)
    '                                .OptionKataban = Space(intSpaceCnt_04)
    '                                .PositionInfo = Space(intSpaceCnt_05)
    '                                .Quantity = "00"
    '                                .OrderNo = ""
    '                            End With
    '                        Next
    '                        For intLoopCnt = 1 To intLoopMax_02
    '                            With strcAccessoryInfo(intLoopCnt)
    '                                .AttributeSymbol = Space(intSpaceCnt_03)
    '                                .OptionKataban = Space(intSpaceCnt_04)
    '                                .Quantity = "00"
    '                            End With
    '                        Next
    '                        '設定
    '                        If objOption.fncVaccumMixCheck(objKtbnStrc) Then
    '                            intIndex = 0
    '                            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
    '                                If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
    '                                    intIndex = intIndex + 1
    '                                    strcManifoldInfo(intIndex).AttributeSymbol = LSet(objKtbnStrc.strcSelection.strAttributeSymbol(intLoopCnt).Trim, intSpaceCnt_03)
    '                                    strOpArray = Split(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, CdCst.Sign.Comma)
    '                                    For intLoopCnt1 = 0 To strOpArray.Length - 1
    '                                        strcManifoldInfo(intIndex).OptionKataban = strcManifoldInfo(intIndex).OptionKataban.Trim & _
    '                                                                                   strOpArray(intLoopCnt1).Trim
    '                                    Next
    '                                    strcManifoldInfo(intIndex).OptionKataban = LSet(KHKataban.fncHypenCut(strcManifoldInfo(intIndex).OptionKataban), intSpaceCnt_04)
    '                                    strPositionInfo = objKtbnStrc.strcSelection.strPositionInfo(intLoopCnt).Trim
    '                                    strPositionInfo = Replace(strPositionInfo, ",", "")
    '                                    strPositionInfo = Replace(strPositionInfo, "0", " ")
    '                                    strPositionInfo = Replace(strPositionInfo, "1", "Y")
    '                                    strcManifoldInfo(intIndex).PositionInfo = LSet(strPositionInfo, intSpaceCnt_05)
    '                                    strcManifoldInfo(intIndex).Quantity = Format(objKtbnStrc.strcSelection.intQuantity(intLoopCnt), "00")
    '                                End If
    '                            Next
    '                        End If
    '                    Case Else
    '                        '初期化
    '                        For intLoopCnt = 1 To intLoopMax_01
    '                            With strcManifoldInfo(intLoopCnt)
    '                                .AttributeSymbol = Space(intSpaceCnt_03)
    '                                .OptionKataban = Space(intSpaceCnt_04)
    '                                .PositionInfo = Space(intSpaceCnt_05)
    '                                .Quantity = "00"
    '                                .OrderNo = ""
    '                            End With
    '                        Next
    '                        For intLoopCnt = 1 To intLoopMax_02
    '                            With strcAccessoryInfo(intLoopCnt)
    '                                .AttributeSymbol = Space(intSpaceCnt_03)
    '                                .OptionKataban = Space(intSpaceCnt_04)
    '                                .Quantity = "00"
    '                            End With
    '                        Next
    '                End Select

    '                result.clKatahikiInfoDto.RegistKey = strKeyInfo
    '                result.clKatahikiInfoDto.LineNo = 1
    '                result.clKatahikiInfoDto.Kataban = strcHeader.kataban
    '                result.clKatahikiInfoDto.CheckKubun = strcHeader.KatabanCheckDiv
    '                result.clKatahikiInfoDto.ListPrice = strcHeader.ListPrice
    '                result.clKatahikiInfoDto.RegistPrice = strcHeader.RegPrice
    '                result.clKatahikiInfoDto.SsPrice = strcHeader.SsPrice
    '                result.clKatahikiInfoDto.BsPrice = strcHeader.BsPrice
    '                result.clKatahikiInfoDto.GsPrice = strcHeader.GsPrice
    '                result.clKatahikiInfoDto.PsPrice = strcHeader.PsPrice
    '                result.clKatahikiInfoDto.NetPrice = strcHeader.NetPrice
    '                result.clKatahikiInfoDto.Currency = strcHeader.CurrencyCode
    '                result.clKatahikiInfoDto.KisyuCode = strcHeader.ModelCd.Trim

    '                'マニホールド仕様情報
    '                strShiyouInfo = Nothing
    '                strShiyouInfo = strShiyouInfo & strcHeader.WiringSpecDiv
    '                strShiyouInfo = strShiyouInfo & strcHeader.RailLength
    '                'strShiyouInfo = strShiyouInfo & strcHeader.ProcDatetime
    '                'strShiyouInfo = strShiyouInfo & strcHeader.FullKataban

    '                For intLoopCnt = 1 To intLoopMax_01
    '                    strShiyouInfo = strShiyouInfo & (strcManifoldInfo(intLoopCnt).AttributeSymbol)
    '                    strShiyouInfo = strShiyouInfo & (strcManifoldInfo(intLoopCnt).OptionKataban)
    '                    strShiyouInfo = strShiyouInfo & (strcManifoldInfo(intLoopCnt).PositionInfo)
    '                    strShiyouInfo = strShiyouInfo & (strcManifoldInfo(intLoopCnt).Quantity)
    '                    'strShiyouInfo = strShiyouInfo & (strcManifoldInfo(intLoopCnt).OrderNo)
    '                Next
    '                For intLoopCnt = 1 To intLoopMax_02
    '                    strShiyouInfo = strShiyouInfo & (strcAccessoryInfo(intLoopCnt).AttributeSymbol)
    '                    strShiyouInfo = strShiyouInfo & (strcAccessoryInfo(intLoopCnt).OptionKataban)
    '                    strShiyouInfo = strShiyouInfo & (strcAccessoryInfo(intLoopCnt).Quantity)
    '                Next

    '                result.clKatahikiInfoDto.ManifoldSpecData = strShiyouInfo
    '                result.clKatahikiInfoDto.ElKubun = strcHeader.ELDiv

    '                If objKtbnStrc.strcSelection.strSalesUnit = Nothing Then
    '                    result.clKatahikiInfoDto.SalesUnit = "PC"
    '                Else
    '                    result.clKatahikiInfoDto.SalesUnit = objKtbnStrc.strcSelection.strSalesUnit
    '                End If

    '                result.clKatahikiInfoDto.SapBaseUnit = objKtbnStrc.strcSelection.strSapBaseUnit
    '                result.clKatahikiInfoDto.QuantityPerSalesUnit = objKtbnStrc.strcSelection.strQuantityPerSalesUnit
    '                If objKtbnStrc.strcSelection.strOrderLot = Nothing Then
    '                    result.clKatahikiInfoDto.OrderLot = 0
    '                Else
    '                    result.clKatahikiInfoDto.OrderLot = objKtbnStrc.strcSelection.strOrderLot
    '                End If

    '                result.clKatahikiInfoDto.CzFlag = blCZFlag

    '        End Select

    '        Return result

    '    Catch ex As Exception
    '        WriteErrorLog("E001", ex)
    '        fncJutyuEdiInterfaceGet = Nothing
    '    Finally
    '        sbBuilder = Nothing
    '        objKataban = Nothing
    '        objOption = Nothing
    '    End Try

    'End Function

    ''' <summary>
    ''' ISO用付属品取得処理
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strAccAttributeSymbol"></param>
    ''' <param name="strAccOptionKataban"></param>
    ''' <param name="intAccQuantity"></param>
    ''' <remarks></remarks>
    Private Shared Sub subISOAccessoryGet(objKtbnStrc As KHKtbnStrc, ByRef strAccAttributeSymbol() As String, _
                                   ByRef strAccOptionKataban() As String, ByRef intAccQuantity() As Integer)
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer
        Dim intLoopCnt3 As Integer
        Dim strWkKataban As String
        Dim intBoltSize As Integer
        Dim intStdSize As Integer
        Dim intPSize As Integer
        Dim intRSize As Integer
        Dim intPCSize As Integer
        Dim intSRSize As Integer
        Dim intPosition As Integer = 0
        Dim strVariation As String = ""

        Try
            ReDim strAccAttributeSymbol(10)
            ReDim strAccOptionKataban(10)
            ReDim intAccQuantity(10)

            For intLoopCnt1 = 0 To 10
                strAccAttributeSymbol(intLoopCnt1) = ""
                strAccOptionKataban(intLoopCnt1) = ""
                intAccQuantity(intLoopCnt1) = 0
            Next

            Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                Case "05"
                    For intLoopCnt1 = 1 To strcManifoldInfo.Length - 1
                        If strcManifoldInfo(intLoopCnt1).AttributeSymbol = "G2" Then
                            If CInt(strcManifoldInfo(intLoopCnt1).Quantity) > 0 Then
                                For intLoopCnt2 = 1 To 10
                                    If Mid(strcManifoldInfo(intLoopCnt1).PositionInfo, intLoopCnt2, 1) = "Y" Then
                                        strVariation = ""

                                        ' PV5-6の場合
                                        If InStr(1, strcManifoldInfo(intLoopCnt1).OptionKataban, "PV5-6") <> 0 Or _
                                           InStr(1, strcManifoldInfo(intLoopCnt1).OptionKataban, "PV5G-6") <> 0 Then
                                            ' ﾍｯﾄﾞ形番設定
                                            strWkKataban = "CMF1-M5*"
                                            ' 基本のｻﾞｲｽﾞ設定
                                            intStdSize = 35
                                            ' 給気ｽﾍﾟｰｻのｻｲｽﾞ設定
                                            intPSize = 30
                                            ' 排気ｽﾍﾟｰｻのｻｲｽﾞ設定
                                            intRSize = 30
                                            ' ﾊﾟｲﾛｯﾄﾁｪｯｸ弁のｻｲｽﾞ設定
                                            intPCSize = 40
                                            ' ｽﾍﾟｰｻ形減圧弁のｻｲｽﾞ設定
                                            intSRSize = 40
                                        Else
                                            ' PV5-8の場合
                                            ' ﾍｯﾄﾞ形番設定
                                            strWkKataban = "CMF2-M6*"
                                            ' 基本のｻﾞｲｽﾞ設定
                                            intStdSize = 45
                                            ' 給気ｽﾍﾟｰｻのｻｲｽﾞ設定
                                            intPSize = 40
                                            ' 排気ｽﾍﾟｰｻのｻｲｽﾞ設定
                                            intRSize = 40
                                            ' ﾊﾟｲﾛｯﾄﾁｪｯｸ弁のｻｲｽﾞ設定
                                            intPCSize = 60
                                            ' ｽﾍﾟｰｻ形減圧弁のｻｲｽﾞ設定
                                            intSRSize = 55
                                        End If

                                        ' 基本ｻｲｽﾞ設定
                                        intBoltSize = intStdSize

                                        For intLoopCnt3 = 1 To strcManifoldInfo.Length - 1
                                            Select Case strcManifoldInfo(intLoopCnt3).AttributeSymbol
                                                Case "G9", "GA", "GB", "GC"
                                                    If CInt(strcManifoldInfo(intLoopCnt3).Quantity) > 0 Then
                                                        If Mid(strcManifoldInfo(intLoopCnt3).PositionInfo, intLoopCnt2, 1) = "Y" Then
                                                            Select Case strcManifoldInfo(intLoopCnt3).AttributeSymbol
                                                                Case "G9"
                                                                    intBoltSize = intBoltSize + intPSize

                                                                    ' ﾎﾞﾙﾄの組合せを設定
                                                                    If strVariation.Length = 0 Then
                                                                        strVariation = "P･R"
                                                                    Else
                                                                        strVariation = strVariation & "+P･R"
                                                                    End If
                                                                Case "GA"
                                                                    intBoltSize = intBoltSize + intRSize

                                                                    ' ﾎﾞﾙﾄの組合せを設定
                                                                    If strVariation.Length = 0 Then
                                                                        strVariation = "P･R"
                                                                    Else
                                                                        strVariation = strVariation & "+P･R"
                                                                    End If
                                                                Case "GB"
                                                                    intBoltSize = intBoltSize + intPCSize

                                                                    ' ﾎﾞﾙﾄの組合せを設定
                                                                    If strVariation.Length = 0 Then
                                                                        strVariation = "PC"
                                                                    Else
                                                                        strVariation = strVariation & "+PC"
                                                                    End If
                                                                Case "GC"
                                                                    intBoltSize = intBoltSize + intSRSize

                                                                    ' ﾎﾞﾙﾄの組合せを設定
                                                                    If strVariation.Length = 0 Then
                                                                        strVariation = "SR"
                                                                    Else
                                                                        strVariation = strVariation & "+SR"
                                                                    End If
                                                            End Select
                                                        End If
                                                    End If
                                            End Select
                                        Next

                                        ' 電磁弁+OPを2個以上組み合わせて選択した場合に取付ﾎﾞﾙﾄ形番を設定する
                                        If InStr(1, strVariation, "+") <> 0 Then
                                            For intLoopCnt3 = 1 To intPosition + 1
                                                ' 同一のOP組合せが存在する場合は個数をﾌﾟﾗｽする
                                                If strAccOptionKataban(intLoopCnt3).Trim = Trim(Left(strWkKataban & CStr(intBoltSize) & Space(12), 12) & "(" & strVariation & ")") Then
                                                    ' 個数設定
                                                    intAccQuantity(intLoopCnt3) = intAccQuantity(intLoopCnt3) + 4
                                                    Exit For
                                                End If

                                                ' 同一のOP組合せが存在しなかった場合は取付ﾎﾞﾙﾄ形番を設定する
                                                If intLoopCnt3 = intPosition + 1 Then
                                                    intPosition = intPosition + 1

                                                    ' 属性記号設定
                                                    strAccAttributeSymbol(intPosition) = "GG"
                                                    ' 形番設定
                                                    strAccOptionKataban(intPosition) = Left(strWkKataban & CStr(intBoltSize) & Space(12), 12) & "(" & strVariation & ")"
                                                    ' 個数設定
                                                    intAccQuantity(intPosition) = 4
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next
                Case "06"
                    For intLoopCnt1 = 1 To strcManifoldInfo.Length - 1
                        If strcManifoldInfo(intLoopCnt1).AttributeSymbol = "G2" Then
                            If CInt(strcManifoldInfo(intLoopCnt1).Quantity) > 0 Then
                                For intLoopCnt2 = 1 To 10
                                    If Mid(strcManifoldInfo(intLoopCnt1).PositionInfo, intLoopCnt2, 1) = "Y" Then
                                        strVariation = ""

                                        ' ﾍｯﾄﾞ形番設定
                                        strWkKataban = "LMF0-M4*"
                                        ' 基本のｻﾞｲｽﾞ設定
                                        intStdSize = 40
                                        ' 給気ｽﾍﾟｰｻのｻｲｽﾞ設定
                                        intPSize = 22
                                        ' 排気ｽﾍﾟｰｻのｻｲｽﾞ設定
                                        intRSize = 22
                                        ' ﾊﾟｲﾛｯﾄﾁｪｯｸ弁のｻｲｽﾞ設定
                                        intPCSize = 36

                                        ' 基本ｻｲｽﾞ設定
                                        intBoltSize = intStdSize

                                        For intLoopCnt3 = 1 To strcManifoldInfo.Length - 1
                                            Select Case strcManifoldInfo(intLoopCnt3).AttributeSymbol
                                                Case "G9", "GA", "GB"
                                                    If CInt(strcManifoldInfo(intLoopCnt3).Quantity) > 0 Then
                                                        If Mid(strcManifoldInfo(intLoopCnt3).PositionInfo, intLoopCnt2, 1) = "Y" Then
                                                            Select Case strcManifoldInfo(intLoopCnt3).AttributeSymbol
                                                                Case "G9"
                                                                    intBoltSize = intBoltSize + intPSize

                                                                    ' ﾎﾞﾙﾄの組合せを設定
                                                                    If strVariation.Length = 0 Then
                                                                        strVariation = "P･R"
                                                                    Else
                                                                        strVariation = strVariation & "+P･R"
                                                                    End If
                                                                Case "GA"
                                                                    intBoltSize = intBoltSize + intRSize

                                                                    ' ﾎﾞﾙﾄの組合せを設定
                                                                    If strVariation.Length = 0 Then
                                                                        strVariation = "P･R"
                                                                    Else
                                                                        strVariation = strVariation & "+P･R"
                                                                    End If
                                                                Case "GB"
                                                                    intBoltSize = intBoltSize + intPCSize

                                                                    ' ﾎﾞﾙﾄの組合せを設定
                                                                    If strVariation.Length = 0 Then
                                                                        strVariation = "PC"
                                                                    Else
                                                                        strVariation = strVariation & "+PC"
                                                                    End If
                                                            End Select
                                                        End If
                                                    End If
                                            End Select
                                        Next

                                        ' ﾎﾞﾙﾄは5mmﾋﾟｯﾁの為、再計算
                                        Select Case Right(Trim(CStr(intBoltSize)), 1)
                                            Case "1", "6"
                                                intBoltSize = intBoltSize + 4
                                            Case "2", "7"
                                                intBoltSize = intBoltSize + 3
                                            Case "3", "8"
                                                intBoltSize = intBoltSize + 2
                                            Case "4", "9"
                                                intBoltSize = intBoltSize + 1
                                        End Select

                                        ' 電磁弁+OPを2個以上組み合わせて選択した場合に取付ﾎﾞﾙﾄ形番を設定する
                                        If InStr(1, strVariation, "+") <> 0 Then
                                            For intLoopCnt3 = 1 To intPosition + 1
                                                ' 同一のOP組合せが存在する場合は個数をﾌﾟﾗｽする
                                                If strAccOptionKataban(intLoopCnt3).Trim = Trim(Left(strWkKataban & CStr(intBoltSize) & Space(12), 12) & "(" & strVariation & ")") Then
                                                    ' 個数設定
                                                    intAccQuantity(intLoopCnt3) = intAccQuantity(intLoopCnt3) + 4
                                                    Exit For
                                                End If

                                                ' 同一のOP組合せが存在しなかった場合は取付ﾎﾞﾙﾄ形番を設定する
                                                If intLoopCnt3 = intPosition + 1 Then
                                                    intPosition = intPosition + 1

                                                    ' 属性記号設定
                                                    strAccAttributeSymbol(intPosition) = "GG"
                                                    ' 形番設定
                                                    strAccOptionKataban(intPosition) = Left(strWkKataban & CStr(intBoltSize) & Space(12), 12) & "(" & strVariation & ")"
                                                    ' 個数設定
                                                    intAccQuantity(intPosition) = 4
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

End Class
