Imports WebKataban.ClsCommon
Imports Microsoft.Office.Interop
Imports System.Data.SqlClient
Imports CKDStandard.ManifoldExport.Data
Imports CKDStandard.ManifoldExport.Common
Imports CKDStandard.ManifoldExport.Constructor
Imports CKDStandard.ManifoldExport.Creator

Public Class KHExcelCtl
#Region "Define"
    'エクセル操作パス
    Private strXlDir As String
    Private strXlTemplate As String
    Private strXlUserDir As String
    Private strXlUserFile As String

    'Excelテンプレートファイル名称
    Private strXlFileName As String = "Manifold"

    'Excelテキストボックス名称
    Private strXlKata As String = "Input_Model"
    Private strXlDate As String = "Input_Date"
    Private strXlTitle As String = "Title"
    Private strXlChargePerson As String = "Charge_Person"
    Private strXlModel As String = "Model"
    Private strXlQuantitiy As String = "Quantitiy"
    Private strXlDeliveryTime As String = "Delivery_Time"
    Private strXlIssueDate As String = "Issue_Date"
    Private strXlCompanyName As String = "Company_Name"
    Private strXlPersonName As String = "Person_Name"
    Private strXlOrderNo As String = "Order_No"

    '設置位置使用文字
    Private strMarkZenkaku As String = "●"
    Private strMarkHankaku As String = "@"
    Private strMarkHyphen As String = "-"
#End Region

    ''' <summary>
    ''' 仕様書EXCEL出力関数コントロール
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strUserId">ユーザーＩＤ</param>
    ''' <param name="strSessionId">セッションＩＤ</param>
    ''' <param name="strLangCd">言語コード</param>
    ''' <param name="strSiyou"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncExcelOutput(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, ByVal strUserId As String, _
                                   ByVal strSessionId As String, ByVal strLangCd As String, _
                                   Optional ByVal strSiyou As String = "") As String
        fncExcelOutput = ""
        Try
            'エクセル作成関数コール
            '※仕様書番号より呼び出し制御する
            Dim strMode As String = String.Empty
            Select Case objKtbnStrc.strcSelection.strSpecNo
                '.NET版と一致するように
                'Case "01", "A3"
                '    strMode = "01"
                'Case "02", "03", "04", "05", "06", "08", "09", "10", "11", "13", "14", "15", "16", "17"
                '    strMode = objKtbnStrc.strcSelection.strSpecNo
                'Case "07", "96"
                '    strMode = "07"
                'Case "12", "18", "19", "20", "21", "22", "23", "94", "95"
                '    strMode = "12"
                'Case Else
                '    strMode = String.Empty
                Case "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", _
                     "11", "12", "13", "14", "15", "16", "17"
                    strMode = objKtbnStrc.strcSelection.strSpecNo
                Case "A3"
                    strMode = "01"
                Case "96"
                    strMode = "07"
                Case "18", "19", "20", "21", "22", "23", "94", "95"
                    strMode = "12"
                Case Else
                    strMode = String.Empty
            End Select
            fncExcelOutput = Me.fncMakeExcel(objCon, objKtbnStrc, strUserId, strSessionId, strLangCd, strMode, strSiyou)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    ''' <summary>
    ''' 仕様書EXCEL出力（簡易仕様書）
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strUserId"></param>
    ''' <param name="strSessId"></param>
    ''' <param name="strLangCd"></param>
    ''' <param name="strMode"></param>
    ''' <param name="strSiyou"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncMakeExcel(ByVal objCon As SqlConnection, objKtbnStrc As KHKtbnStrc, ByVal strUserId As String, _
                                  ByVal strSessId As String, ByVal strLangCd As String, strMode As String, _
                                  Optional ByVal strSiyou As String = "") As String

        Dim clsManCommon As KHManifold
        Dim dtSelSpec As DataTable
        Dim dtSpecInfo As DataTable
        Dim strFullKataban As String
        Dim objRow As DataRow
        Dim strQty As String
        Dim strCheck As String
        Dim strRailLen As String
        Dim intC As Integer
        Dim intR As Integer
        Dim strMark As String
        'CHANGED BY YGY 20141022    ↓↓↓↓↓↓
        'Dim xlTemplatePath As String = strXlDir & strXlTemplate & strXlFileName & strMode & ".xls"
        Dim xlTemplatePath As String = HttpContext.Current.Server.MapPath("Template") & "\Manifold" & strMode & ".xls"
        Dim xlUserFilePath As String

        If strSiyou.Equals(String.Empty) Then
            'CHANGED BY YGY 20141106
            'xlUserFilePath = strXlDir & strXlUserDir & strUserId & "_" & strXlUserFile
            xlUserFilePath = HttpContext.Current.Server.MapPath("TempFiles") & "\" & strUserId & "_" & CdCst.strExcelTmpFileName
        Else
            '仕様書一括出力
            If Not IO.Directory.Exists(My.Settings.ExcelOutputPathForTest & strMode & "\") Then
                IO.Directory.CreateDirectory(My.Settings.ExcelOutputPathForTest & strMode & "\")
            End If
            xlUserFilePath = My.Settings.ExcelOutputPathForTest & strMode & "\" & strSiyou & ".xls"
        End If
        'CHANGED BY YGY 20141022    ↑↑↑↑↑↑
        Dim xlApp As Excel.Application
        Try
            xlApp = New Excel.Application
        Catch ex As Exception
            WriteErrorLog("E001", ex)
            Return String.Empty
        End Try
        Dim xlBooks As Excel.Workbooks = xlApp.Workbooks
        Dim xlBook As Excel.Workbook = Nothing
        Dim xlSheets As Excel.Sheets = Nothing
        Dim xlSheet As Excel.Worksheet = Nothing
        Dim xlCells As Excel.Range = Nothing
        Dim xlRange1 As Excel.Range

        Dim intAddRows As Integer = -1
        Dim strKataban As String = String.Empty
        Dim intNo As Integer
        Dim intCnt As Integer
        Dim intC2 As Integer

        Try
            strFullKataban = objKtbnStrc.strcSelection.strFullKataban      'フル形番
            strRailLen = objKtbnStrc.strcSelection.decDinRailLength        '取付レール長さ

            clsManCommon = New KHManifold(strUserId, strSessId)

            '引当情報の取得
            dtSelSpec = clsManCommon.fncSelectSelSpec(objCon)

            '品名マスタの取得
            dtSpecInfo = fncSelSpecInfoData(objCon, objKtbnStrc.strcSelection.strSpecNo, strLangCd)

            '設置位置マークセット
            If strLangCd = CdCst.LanguageCd.Japanese Then
                strMark = strMarkZenkaku
            Else
                strMark = strMarkHankaku
            End If

            'テンプレートから処理ファイルを複製　※上書き
            System.IO.File.Copy(xlTemplatePath, xlUserFilePath, True)

            '処理ファイルを開く
            xlBook = xlBooks.Open(xlUserFilePath)
            xlSheets = xlBook.Worksheets
            xlSheet = xlSheets.Item(1)
            xlCells = xlSheet.Cells

            '共通ヘッダ部セット
            subSetHeader(objCon, xlSheet, strLangCd, strFullKataban, strMode)

            Select Case strMode
                Case "01", "A3"
                    intAddRows = 12
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(intAddRows + 1 + intR, 1)
                        xlRange1.Value = objRow("label_content")
                        If intR = CdCst.Siyou_01.Tube - 1 Then
                            If strLangCd = CdCst.LanguageCd.Japanese Then
                                xlRange1.Value = xlRange1.Value & CdCst.Manifold.UnNecessity.Japanese
                            Else
                                xlRange1.Value = xlRange1.Value & " " & CdCst.Manifold.UnNecessity.English
                            End If
                        End If
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '取付レール長さセット
                    'MT3Q0シリーズは取付レール長さなし
                    If InStr(1, strFullKataban.Trim, "MT3Q0") <> 0 Then
                    Else
                        xlRange1 = xlCells(intAddRows + CdCst.Siyou_01.Rail, 6)
                        xlRange1.Value = strRailLen
                        ClsCommon.MRComObject(xlRange1)
                    End If

                    'ﾁｭ-ﾌﾞ抜具不要セット
                    objRow = dtSelSpec.Rows(CdCst.Siyou_01.Tube - 1)
                    If objRow(CdCst.SelSpec.Kataban).ToString = "0" Then
                        xlRange1 = xlCells(intAddRows + CdCst.Siyou_01.Tube, 2)
                        xlRange1.Value = strMark
                        ClsCommon.MRComObject(xlRange1)
                    End If
                Case "02"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        xlRange1.Value = objRow("label_content")
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '取付レール長さセット
                    xlRange1 = xlCells(33, 6)
                    xlRange1.Value = strRailLen
                    ClsCommon.MRComObject(xlRange1)
                Case "03", "04"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        xlRange1.Value = objRow("label_content")
                        If intR = CdCst.Siyou_03.Tube - 1 Then
                            If strLangCd = CdCst.LanguageCd.Japanese Then
                                xlRange1.Value = xlRange1.Value & CdCst.Manifold.UnNecessity.Japanese
                            Else
                                xlRange1.Value = xlRange1.Value & " " & CdCst.Manifold.UnNecessity.English
                            End If
                        End If
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '形番・設置位置セット
                    For intI As Integer = 1 To 22
                        objRow = dtSelSpec.Rows(intI)
                        strQty = objRow(CdCst.SelSpec.Qty).ToString
                        '使用数 "0" はプリントしない
                        If strQty <> "0" Then
                            '取付レール以降は1行ずれる
                            If intI > 19 Then
                                intR = intI + 1
                            Else
                                intR = intI
                            End If
                            '形番
                            '検査成績書文字列変換
                            strKataban = fncInspReportComb(objRow(CdCst.SelSpec.Kataban).ToString, strLangCd)
                            xlRange1 = xlCells(12 + intR, 2)
                            xlRange1.Value = strKataban
                            ClsCommon.MRComObject(xlRange1)

                            '使用数
                            Select Case strMode
                                Case "03"
                                    If intI > 15 Then
                                        intC = 6
                                    Else
                                        intC = 30
                                    End If
                                Case "04"
                                    If intI > 14 Then
                                        intC = 6
                                    Else
                                        intC = 30
                                    End If
                            End Select

                            xlRange1 = xlCells(12 + intR, intC)
                            xlRange1.Value = strQty
                            ClsCommon.MRComObject(xlRange1)

                            '継手ＣＸが存在する場合
                            If Not IsDBNull(objRow(CdCst.SelSpec.CxA)) Then
                                If objRow(CdCst.SelSpec.CxA) <> "" Then
                                    xlRange1 = xlCells(12 + intR, 6)
                                    xlRange1.Value = objRow(CdCst.SelSpec.CxA).ToString
                                    ClsCommon.MRComObject(xlRange1)
                                End If
                            End If
                            If Not IsDBNull(objRow(CdCst.SelSpec.CxB)) Then
                                If objRow(CdCst.SelSpec.CxB) <> "" Then
                                    xlRange1 = xlCells(12 + intR, 8)
                                    xlRange1.Value = objRow(CdCst.SelSpec.CxB).ToString
                                    ClsCommon.MRComObject(xlRange1)
                                End If
                            End If

                            '設置位置
                            '設置位置が無い項目は対象外
                            If Not IsDBNull(objRow(CdCst.SelSpec.PosInfo)) Then
                                intC = 0
                                For intj As Integer = 0 To objRow(CdCst.SelSpec.PosInfo).ToString.Length - 1
                                    strCheck = Strings.Mid(objRow(CdCst.SelSpec.PosInfo).ToString, intj + 1, 1)
                                    If strCheck = "1" Then
                                        xlRange1 = xlCells(12 + intR, 10 + intC)
                                        xlRange1.Value = strMark
                                        ClsCommon.MRComObject(xlRange1)
                                    End If
                                    intC = intC + 1
                                Next
                            End If
                        End If
                    Next

                    '取付レール長さセット
                    If Int(strRailLen) <> 0 Then
                        '0は設定無し
                        xlRange1 = xlCells(32, 6)
                        xlRange1.Value = strRailLen
                        ClsCommon.MRComObject(xlRange1)
                    End If

                    'ﾁｭ-ﾌﾞ抜具不要セット
                    objRow = dtSelSpec.Rows(23)
                    If objRow(CdCst.SelSpec.Kataban).ToString = "0" Then
                        xlRange1 = xlCells(36, 2)
                        xlRange1.Value = strMark
                        ClsCommon.MRComObject(xlRange1)
                    End If
                Case "05"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        Select Case intR
                            Case 6, 8
                                xlRange1.Value = objRow("label_content").ToString.Split("(")(0)
                            Case Else
                                xlRange1.Value = objRow("label_content")
                        End Select
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '接続ブロック
                    'xlRange1 = xlCells(36, 1)
                    'If strLangCd = CdCst.LanguageCd.Japanese Then
                    '    xlRange1.Value = "接続ブロック"
                    'Else
                    '    xlRange1.Value = "Mix Block"
                    'End If
                    'ClsCommon.MRComObject(xlRange1)

                    '制御ユニット付の場合は1連目・2連目にハイフンをセットする
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4", "5", "6", "7", "9"
                            For intI As Integer = 1 To 24
                                Select Case intI
                                    Case 1 To 21
                                        xlRange1 = xlCells(12 + intI, 6)
                                    Case Else
                                        xlRange1 = xlCells(12 + intI, 7)
                                End Select
                                xlRange1.Value = strMarkHyphen
                                ClsCommon.MRComObject(xlRange1)
                                Select Case intI
                                    Case 1 To 21
                                        xlRange1 = xlCells(12 + intI, 8)
                                    Case Else
                                        xlRange1 = xlCells(12 + intI, 9)
                                End Select
                                xlRange1.Value = strMarkHyphen
                                ClsCommon.MRComObject(xlRange1)
                            Next
                    End Select

                    '形番・設置位置セット
                    intCnt = 0
                    Dim bolOut As Boolean = False
                    For intI As Integer = 1 To 25
                        objRow = dtSelSpec.Rows(intI)
                        strQty = objRow(CdCst.SelSpec.Qty).ToString
                        '使用数 "0" はプリントしない
                        If intCnt <> 0 Then
                            bolOut = True
                            Select Case intI
                                Case 2 To 7
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "1", "4", "6"
                                            bolOut = False
                                        Case "8"
                                            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "" Then
                                                bolOut = False
                                            End If
                                    End Select
                            End Select

                            If bolOut Then
                                intR = intI
                                '形番
                                '検査成績書文字列変換
                                If strQty <> "0" Then
                                    strKataban = fncInspReportComb(objRow(CdCst.SelSpec.Kataban).ToString, strLangCd)
                                    xlRange1 = xlCells(11 + intR, 2)
                                    xlRange1.Value = strKataban
                                    ClsCommon.MRComObject(xlRange1)
                                End If
                                '使用数
                                If strQty <> "0" Then
                                    xlRange1 = xlCells(11 + intR, 31)
                                    xlRange1.Value = strQty
                                    ClsCommon.MRComObject(xlRange1)
                                End If
                                '設置位置
                                '設置位置が無い項目は対象外
                                If Not IsDBNull(objRow(CdCst.SelSpec.PosInfo)) Then
                                    intC = 0
                                    For intj As Integer = 0 To objRow(CdCst.SelSpec.PosInfo).ToString.Length - 1
                                        strCheck = Strings.Mid(objRow(CdCst.SelSpec.PosInfo).ToString, intj + 1, 1)
                                        If strCheck = "1" Then
                                            If intR > 22 Then
                                                intC2 = 1
                                            Else
                                                intC2 = 0
                                            End If
                                            xlRange1 = xlCells(11 + intR, 6 + intC2 + intC)
                                            xlRange1.Value = strMark
                                            ClsCommon.MRComObject(xlRange1)
                                        End If
                                        intC = intC + 2
                                    Next
                                End If
                            End If
                        Else
                            intCnt = intCnt + 1
                        End If
                    Next

                    '受注No.設定
                    If strLangCd = CdCst.LanguageCd.Japanese Then
                        xlRange1 = xlCells(38, 1)
                        xlRange1.Value = "関連受注No."
                        ClsCommon.MRComObject(xlRange1)

                        xlRange1 = xlCells(38, 2)
                        xlRange1.Value = "形番"
                        ClsCommon.MRComObject(xlRange1)
                    Else
                        xlRange1 = xlCells(38, 1)
                        xlRange1.Value = "Order No."
                        ClsCommon.MRComObject(xlRange1)

                        xlRange1 = xlCells(38, 2)
                        xlRange1.Value = "Model No."
                        ClsCommon.MRComObject(xlRange1)
                    End If

                    intCnt = 0
                    For intI As Integer = 1 To 24
                        objRow = dtSelSpec.Rows(intI)
                        If objRow(CdCst.SelSpec.Kataban).ToString.Trim <> "" And objRow(CdCst.SelSpec.Qty) > 0 Then
                            intCnt = intCnt + 1

                            xlRange1 = xlCells(38 + intCnt, 2)
                            xlRange1.Value = objRow(CdCst.SelSpec.Kataban).ToString.Trim
                            ClsCommon.MRComObject(xlRange1)
                        End If
                    Next
                Case "06"
                    '設置位置Noセット(Mani06のみ降順があるため)
                    intNo = 0
                    If objKtbnStrc.strcSelection.strOpSymbol(4) = "T0D" Then
                        '降順
                        For i As Integer = 10 To 1 Step -1
                            xlRange1 = xlCells(12, 6 + intNo)
                            xlRange1.Value = CStr(i)
                            ClsCommon.MRComObject(xlRange1)
                            intNo = intNo + 2
                        Next
                    Else
                        '昇順
                        For i As Integer = 1 To 10
                            xlRange1 = xlCells(12, 6 + intNo)
                            xlRange1.Value = CStr(i)
                            ClsCommon.MRComObject(xlRange1)
                            intNo = intNo + 2
                        Next
                    End If

                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        Select Case intR
                            Case 6
                                xlRange1.Value = objRow("label_content").ToString.Split("(")(0)
                            Case Else
                                xlRange1.Value = objRow("label_content")
                        End Select
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '形番・設置位置セット
                    intCnt = 0
                    Dim bolOut As Boolean = False
                    For intI As Integer = 1 To 19
                        If intCnt <> 0 Then
                            bolOut = True
                            Select Case intI
                                Case 2 To 7
                                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                                        bolOut = False
                                    End If
                            End Select

                            If bolOut Then
                                objRow = dtSelSpec.Rows(intI)
                                strQty = objRow(CdCst.SelSpec.Qty).ToString

                                intR = intI
                                '形番
                                '検査成績書文字列変換
                                If strQty <> "0" Then
                                    strKataban = fncInspReportComb(objRow(CdCst.SelSpec.Kataban).ToString, strLangCd)
                                    xlRange1 = xlCells(11 + intR, 2)
                                    xlRange1.Value = strKataban
                                    ClsCommon.MRComObject(xlRange1)
                                End If

                                '使用数
                                If strQty <> "0" Then
                                    xlRange1 = xlCells(11 + intR, 31)
                                    xlRange1.Value = strQty
                                    ClsCommon.MRComObject(xlRange1)
                                End If

                                '設置位置
                                '設置位置が無い項目は対象外
                                If Not IsDBNull(objRow(CdCst.SelSpec.PosInfo)) Then
                                    intC = 0
                                    If objKtbnStrc.strcSelection.strSeriesKataban = "LMF0" AndAlso _
                                        objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "T0D" Then   '逆
                                        For intj As Integer = objRow(CdCst.SelSpec.PosInfo).ToString.Length - 1 To 0 Step -1
                                            strCheck = Strings.Mid(objRow(CdCst.SelSpec.PosInfo).ToString, intj + 1, 1)
                                            If strCheck = "1" Then
                                                If intR > 17 Then
                                                    intC2 = -1
                                                Else
                                                    intC2 = 0
                                                End If
                                                xlRange1 = xlCells(11 + intR, 6 + intC2 + intC)
                                                xlRange1.Value = strMark
                                                ClsCommon.MRComObject(xlRange1)
                                            End If
                                            intC = intC + 2
                                        Next
                                    Else
                                        For intj As Integer = 0 To objRow(CdCst.SelSpec.PosInfo).ToString.Length - 1
                                            strCheck = Strings.Mid(objRow(CdCst.SelSpec.PosInfo).ToString, intj + 1, 1)
                                            If strCheck = "1" Then
                                                If intR > 17 Then
                                                    intC2 = 1
                                                Else
                                                    intC2 = 0
                                                End If
                                                xlRange1 = xlCells(11 + intR, 6 + intC2 + intC)
                                                xlRange1.Value = strMark
                                                ClsCommon.MRComObject(xlRange1)
                                            End If
                                            intC = intC + 2
                                        Next
                                    End If
                                End If
                            End If
                        Else
                            intCnt = intCnt + 1
                        End If
                    Next

                    '受注No.設定
                    If strLangCd = CdCst.LanguageCd.Japanese Then
                        xlRange1 = xlCells(32, 1)
                        xlRange1.Value = "関連受注No."
                        ClsCommon.MRComObject(xlRange1)

                        xlRange1 = xlCells(32, 2)
                        xlRange1.Value = "形番"
                        ClsCommon.MRComObject(xlRange1)
                    Else
                        xlRange1 = xlCells(32, 1)
                        xlRange1.Value = "Order No."
                        ClsCommon.MRComObject(xlRange1)

                        xlRange1 = xlCells(32, 2)
                        xlRange1.Value = "Model No."
                        ClsCommon.MRComObject(xlRange1)
                    End If
                    intCnt = 0
                    For intI As Integer = 1 To 19
                        objRow = dtSelSpec.Rows(intI)
                        If objRow(CdCst.SelSpec.Kataban).ToString.Trim <> "" And objRow(CdCst.SelSpec.Qty) > 0 Then
                            intCnt = intCnt + 1

                            xlRange1 = xlCells(32 + intCnt, 2)
                            xlRange1.Value = objRow(CdCst.SelSpec.Kataban).ToString.Trim
                            ClsCommon.MRComObject(xlRange1)
                        End If
                    Next
                Case "07"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        xlRange1.Value = objRow("label_content")
                        If intR = CdCst.Siyou_07.Tube - 1 Then
                            If strLangCd = CdCst.LanguageCd.Japanese Then
                                xlRange1.Value = xlRange1.Value & CdCst.Manifold.UnNecessity.Japanese
                            Else
                                xlRange1.Value = xlRange1.Value & " " & CdCst.Manifold.UnNecessity.English
                            End If
                        End If
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '取付レール長さセット
                    xlRange1 = xlCells(36, 6)
                    xlRange1.Value = strRailLen
                    ClsCommon.MRComObject(xlRange1)

                    'ﾁｭ-ﾌﾞ抜具不要セット
                    objRow = dtSelSpec.Rows(27)
                    If objRow(CdCst.SelSpec.Kataban).ToString = "0" Then
                        xlRange1 = xlCells(40, 2)
                        xlRange1.Value = strMark
                        ClsCommon.MRComObject(xlRange1)
                    End If
                Case "08"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        xlRange1.Value = objRow("label_content")
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '取付レール長さセット
                    xlRange1 = xlCells(33, 6)
                    xlRange1.Value = strRailLen
                    ClsCommon.MRComObject(xlRange1)
                Case "09"
                    Dim strSrsKataban As String = objKtbnStrc.strcSelection.strSeriesKataban      'シリーズ形番
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        xlRange1.Value = objRow("label_content")

                        '六角穴付プラグ特殊対応(ラベル番号14、15)
                        If CInt(objRow("label_seq").ToString) = 14 Then
                            If strSrsKataban = "M4TB3" Then
                                xlRange1.Value = xlRange1.Value & " R1/4"
                            End If
                            If strSrsKataban = "M4TB4" Then
                                xlRange1.Value = xlRange1.Value & " R1/2"
                            End If
                        End If
                        If CInt(objRow("label_seq").ToString) = 15 Then
                            If strSrsKataban = "M4TB3" Then
                                xlRange1.Value = xlRange1.Value & " R3/8"
                            End If
                            If strSrsKataban = "M4TB4" Then
                                xlRange1.Value = xlRange1.Value & " R3/8"
                            End If
                        End If

                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next
                Case "10"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        xlRange1.Value = objRow("label_content")
                        If intR = 24 - 1 Then
                            If strLangCd = CdCst.LanguageCd.Japanese Then
                                xlRange1.Value = xlRange1.Value & CdCst.Manifold.UnNecessity.Japanese
                            Else
                                xlRange1.Value = xlRange1.Value & " " & CdCst.Manifold.UnNecessity.English
                            End If
                        End If
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '取付レール長さセット
                    If CDec(strRailLen) <> 0 Then
                        xlRange1 = xlCells(32, 6)
                        xlRange1.Value = strRailLen
                        ClsCommon.MRComObject(xlRange1)
                    End If

                    'ﾁｭ-ﾌﾞ抜具不要セット
                    objRow = dtSelSpec.Rows(23)
                    If objRow(CdCst.SelSpec.Kataban).ToString = "0" Then
                        xlRange1 = xlCells(36, 2)
                        xlRange1.Value = strMark
                        ClsCommon.MRComObject(xlRange1)
                    End If
                Case "11"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        xlRange1.Value = objRow("label_content")
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '取付レール長さセット
                    If CDec(strRailLen) <> 0 Then
                        xlRange1 = xlCells(31, 11)
                        xlRange1.Value = strRailLen
                        ClsCommon.MRComObject(xlRange1)
                    End If
                Case "12"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        xlRange1.Value = objRow("label_content")
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next
                Case "13"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        xlRange1.Value = objRow("label_content")
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '取付レール長さセット
                    xlRange1 = xlCells(34, 6)
                    xlRange1.Value = strRailLen
                    ClsCommon.MRComObject(xlRange1)
                Case "14"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        xlRange1.Value = objRow("label_content").ToString
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '取付レール長さセット
                    xlRange1 = xlCells(22, 6)
                    xlRange1.Value = strRailLen
                    ClsCommon.MRComObject(xlRange1)
                Case "15"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        xlRange1.Value = objRow("label_content")
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '取付レール長さセット
                    If Int(strRailLen) <> 0 Then
                        '0は設定無し
                        xlRange1 = xlCells(36, 6)
                        xlRange1.Value = strRailLen
                        ClsCommon.MRComObject(xlRange1)
                    End If
                Case "16"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        If Left(strFullKataban, 6) = "MW4GB4" AndAlso _
                            (intR = 15 OrElse intR = 16 OrElse intR = 17) Then
                            xlRange1.Value = ""
                        Else
                            xlRange1.Value = objRow("label_content")
                        End If
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next
                Case "17"
                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Rows
                        xlRange1 = xlCells(13 + intR, 1)
                        xlRange1.Value = objRow("label_content")
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next
                Case Else
                    '記号セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Select("item_div = '1'")
                        xlRange1 = xlCells(13 + intR, 1)

                        '特殊記号を作成
                        xlRange1.Value = fncConvertMarkSimple(objRow, strFullKataban, objKtbnStrc)
                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next

                    '品名部セット
                    intR = 0
                    For Each objRow In dtSpecInfo.Select("item_div = '2'")
                        'セルの取得
                        xlRange1 = xlCells(13 + intR, 2)

                        '特殊形番を作成しセルに設定
                        xlRange1.Value = fncConvertKantabanSimple(objRow, objKtbnStrc)

                        ClsCommon.MRComObject(xlRange1)
                        intR = intR + Int(objRow("item_num").ToString)
                    Next
            End Select

            Dim intColEnd As Long = 0   '列数
            Dim intUseX As Long = 0     '使用数X
            Dim intUseY As Long = 0     '使用数Y
            Dim intPosY As Long = 0     '設置位置Y
            Select Case strMode
                Case "01"
                    intColEnd = CdCst.Siyou_01.Tube - 2
                    intUseX = 12
                    intUseY = 6
                    intPosY = 6
                Case "02"
                    intColEnd = 21
                    intUseX = 12
                    intUseY = 6
                    intPosY = 6
                Case "07"
                    intColEnd = 26
                    intUseX = 12
                    intUseY = 6
                    intPosY = 6
                Case "08"
                    intColEnd = 21
                    intUseX = 12
                    intUseY = 6
                    intPosY = 6
                Case "09"
                    intColEnd = 23
                    intUseX = 12
                    intUseY = 6
                    intPosY = 6
                Case "10"
                    intColEnd = 22
                    intUseX = 12
                    intUseY = 6
                    intPosY = 6
                Case "11"
                    intColEnd = 18
                    intUseX = 12
                    intUseY = 11
                    intPosY = 11
                Case "12"
                    intColEnd = 10
                    intUseX = 12
                    intUseY = 18
                    intPosY = 6
                Case "13"
                    intColEnd = 24
                    intUseX = 12
                    intUseY = 6
                    intPosY = 6
                Case "14"
                    intColEnd = 9
                    intUseX = 12
                    intUseY = 6
                    intPosY = 6
                Case "15"
                    intColEnd = 28
                    intUseX = 12
                    intUseY = 6
                    intPosY = 6
                Case "16"
                    intColEnd = 25
                    intUseX = 12
                    intUseY = 6
                    intPosY = 6
                Case "17"
                    intColEnd = 6
                    intUseX = 12
                    intUseY = 11
                    intPosY = 6
                Case Else
                    intColEnd = 12
                    intUseX = 12
                    intUseY = 31
                    intPosY = 6
            End Select

            '形番・設置位置セット
            For intI As Integer = 1 To intColEnd
                Select Case strMode
                    Case "03", "04", "05", "06"
                        Exit For
                End Select
                If intI >= dtSelSpec.Rows.Count Then Exit For
                objRow = dtSelSpec.Rows(intI)
                strQty = objRow(CdCst.SelSpec.Qty).ToString

                'DELETED BY YGY 20141217
                '画面から対応するので特殊処理する必要がない
                'Select Case strMode
                '    Case "16"
                '        objRow = dtSelSpec.Rows(intI)
                '        If intI = 17 Then
                '            strQty = (objRow(CdCst.SelSpec.Qty) * 2).ToString
                '        Else
                '            strQty = objRow(CdCst.SelSpec.Qty).ToString
                '        End If
                'End Select

                'ｾﾝｻを追加
                If strFullKataban.StartsWith("MV3QR") Then
                    If objRow(CdCst.SelSpec.Kataban).ToString.Contains("ｾﾝｻ") OrElse _
                        objRow(CdCst.SelSpec.Kataban).ToString.Contains("Senser") Then
                        xlRange1 = xlCells(intUseX + intI, 2)
                        xlRange1.Value = objRow(CdCst.SelSpec.Kataban).ToString
                        ClsCommon.MRComObject(xlRange1)
                    End If
                End If

                '使用数 "0" はプリントしない
                If strQty <> "0" Then
                    intR = intI

                    '取付レール以降は1行ずれる
                    Select Case strMode
                        Case "01"
                            If intI >= CdCst.Siyou_01.Rail Then intR = intI + 1
                        Case "02"
                            If intI > 20 Then intR = intI + 1
                        Case "07"
                            If intI > 23 Then intR = intI + 1
                        Case "08"
                            If intI > 20 Then intR = intI + 1
                        Case "10"
                            If intI > 19 Then intR = intI + 1
                        Case "13"
                            If intI > 21 Then intR = intI + 1
                        Case "15"
                            If intI > 23 Then intR = intI + 1
                    End Select

                    Select Case strMode
                        Case ""    'ADD BY YGY 20141120
                            '品名部既にセットした、検査成績書以外の場合再設定必要がある？
                        Case "01", "02", "07", "08", "09", "10", "12", "13", "15", "16", "17"
                            '検査成績書文字列変換
                            strKataban = fncInspReportComb(objRow(CdCst.SelSpec.Kataban).ToString, strLangCd)
                            xlRange1 = xlCells(intUseX + intR, 2)
                            xlRange1.Value = KHKataban.fncHypenCut(strKataban.Replace(CdCst.Sign.Comma, ""))
                            ClsCommon.MRComObject(xlRange1)
                        Case Else
                            xlRange1 = xlCells(intUseX + intR, 2)
                            xlRange1.Value = objRow(CdCst.SelSpec.Kataban).ToString
                            ClsCommon.MRComObject(xlRange1)
                    End Select

                    '使用数
                    Select Case strMode
                        Case "01"
                            intUseY = 31
                            If intI > 20 Then intUseY = 6
                        Case "02"
                            intUseY = 31
                            If intI > 16 Then intUseY = 6
                        Case "07"
                            intUseY = 31
                            If intI > 20 Then intUseY = 6
                        Case "08"
                            intUseY = 31
                            If intI > 16 Then intUseY = 6
                        Case "09"
                            intUseY = 56
                            If intI > 17 Then intUseY = 6
                        Case "10"
                            intUseY = 31
                            If intI > 14 Then intUseY = 6
                        Case "11"
                            intUseY = 31
                            If intI > 15 Then intUseY = 11
                        Case "13"
                            intUseY = 56
                            If intI > 17 Then intUseY = 6
                        Case "14"
                            intUseY = 31
                            If intI > 6 Then intUseY = 6
                        Case "15"
                            intUseY = 31
                            If intI > 20 Then intUseY = 6
                        Case "16"
                            intUseY = 56
                            If intI > 20 Then intUseY = 6
                    End Select

                    '使用数
                    xlRange1 = xlCells(intUseX + intR, intUseY)
                    xlRange1.Value = strQty
                    Select Case strMode
                        Case "09"
                            If intI = 17 Then xlRange1.Value = strQty * 2
                    End Select
                    ClsCommon.MRComObject(xlRange1)

                    '設置位置が無い項目は対象外
                    If Not IsDBNull(objRow(CdCst.SelSpec.PosInfo)) Then
                        intC = 0
                        For intj As Integer = 0 To objRow(CdCst.SelSpec.PosInfo).ToString.Length - 1
                            strCheck = Strings.Mid(objRow(CdCst.SelSpec.PosInfo).ToString, intj + 1, 1)
                            If strCheck = "1" Then
                                intC2 = 0
                                Select Case strMode
                                    Case "09", "13"
                                        If intR > 15 Then intC2 = 1
                                        xlRange1 = xlCells(intUseX + intR, intPosY + (intC * 2) + intC2)
                                    Case "16"
                                        If intR > 16 Then intC2 = 1
                                        xlRange1 = xlCells(intUseX + intR, intPosY + (intC * 2) + intC2)
                                    Case Else
                                        xlRange1 = xlCells(intUseX + intR, intPosY + intC)
                                End Select
                                xlRange1.Value = strMark
                                ClsCommon.MRComObject(xlRange1)
                            End If
                            intC = intC + 1
                        Next
                    End If
                End If
            Next

            xlBook.Save()
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            '==================  終了処理  =====================  
            ClsCommon.MRComObject(xlCells)            'xlCells の解放
            ClsCommon.MRComObject(xlSheet)            'xlSheet の解放
            ClsCommon.MRComObject(xlSheets)           'xlSheets の解放
            If Not xlBook Is Nothing Then xlBook.Close(False) 'xlBook を閉じる
            ClsCommon.MRComObject(xlBook)             'xlBook の解放
            ClsCommon.MRComObject(xlBooks)            'xlBooks の解放
            xlApp.Quit()                    'Excelを閉じる 
            ClsCommon.MRComObject(xlApp)              'xlApp を解放
            clsManCommon = Nothing
            dtSelSpec = Nothing
            fncMakeExcel = xlUserFilePath
        End Try
    End Function

    ''' <summary>
    ''' 共通ヘッダ部分セット（簡易仕様書用）
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="xlSheet"></param>
    ''' <param name="strLangCd"></param>
    ''' <param name="strFullKataban"></param>
    ''' <param name="strMode"></param>
    ''' <remarks></remarks>
    Private Sub subSetHeader(ByVal objCon As SqlConnection, ByRef xlSheet As Excel.Worksheet, _
                             ByVal strLangCd As String, ByVal strFullKataban As String, strMode As String)
        Dim xlShapes As Excel.Shapes
        Dim xlShape As Excel.Shape
        Dim xlTextFrameas As Excel.TextFrame
        Dim xlCharacters As Excel.Characters
        Dim xlCells As Excel.Range
        Dim xlRange1 As Excel.Range = Nothing
        xlShapes = xlSheet.Shapes
        xlCells = xlSheet.Cells

        'ラベル取得
        Try
            Dim dt_Title As DataTable = KHLabelCtl.fncGetPageAllLabels(objCon, "KHExcelCtl", strLangCd)
            If dt_Title Is Nothing Then Exit Sub
            Dim dr_label() As DataRow = dt_Title.Select("label_div='" & CdCst.Lbl.Division.Label & "'")
            If dr_label.Length <= 0 Then Exit Sub

            'シート名セット
            xlSheet.Name = System.DateTime.Now.ToString("yyyyMMdd_HHmmss")

            'タイトルセット
            xlShape = xlShapes.Item(strXlTitle)
            xlTextFrameas = xlShape.TextFrame
            xlCharacters = xlTextFrameas.Characters
            xlCharacters.Text = dr_label(0)("label_content").ToString
            ClsCommon.MRComObject(xlShape)
            ClsCommon.MRComObject(xlTextFrameas)
            ClsCommon.MRComObject(xlCharacters)

            '担当セット
            xlShape = xlShapes.Item(strXlChargePerson)
            xlTextFrameas = xlShape.TextFrame
            xlCharacters = xlTextFrameas.Characters
            xlCharacters.Text = dr_label(1)("label_content").ToString
            ClsCommon.MRComObject(xlShape)
            ClsCommon.MRComObject(xlTextFrameas)
            ClsCommon.MRComObject(xlCharacters)

            '形番セット
            xlShape = xlShapes.Item(strXlModel)
            xlTextFrameas = xlShape.TextFrame
            xlCharacters = xlTextFrameas.Characters
            xlCharacters.Text = dr_label(2)("label_content").ToString & " " & strFullKataban
            ClsCommon.MRComObject(xlShape)
            ClsCommon.MRComObject(xlTextFrameas)
            ClsCommon.MRComObject(xlCharacters)

            '数量セット
            xlShape = xlShapes.Item(strXlQuantitiy)
            xlTextFrameas = xlShape.TextFrame
            xlCharacters = xlTextFrameas.Characters
            xlCharacters.Text = dr_label(3)("label_content").ToString
            ClsCommon.MRComObject(xlShape)
            ClsCommon.MRComObject(xlTextFrameas)
            ClsCommon.MRComObject(xlCharacters)

            '納期セット
            xlShape = xlShapes.Item(strXlDeliveryTime)
            xlTextFrameas = xlShape.TextFrame
            xlCharacters = xlTextFrameas.Characters
            xlCharacters.Text = dr_label(4)("label_content").ToString
            ClsCommon.MRComObject(xlShape)
            ClsCommon.MRComObject(xlTextFrameas)
            ClsCommon.MRComObject(xlCharacters)

            '発効日セット
            xlShape = xlShapes.Item(strXlIssueDate)
            xlTextFrameas = xlShape.TextFrame
            xlCharacters = xlTextFrameas.Characters
            If strLangCd = CdCst.LanguageCd.Japanese Then
                xlCharacters.Text = dr_label(5)("label_content").ToString & " " & _
                                    System.DateTime.Now.ToString("yyyy/MM/dd")
            Else
                xlCharacters.Text = dr_label(5)("label_content").ToString & " " & _
                                    System.DateTime.Now.ToString("MM/dd/yyyy")
            End If
            ClsCommon.MRComObject(xlShape)
            ClsCommon.MRComObject(xlTextFrameas)
            ClsCommon.MRComObject(xlCharacters)

            '貴社名セット
            xlShape = xlShapes.Item(strXlCompanyName)
            xlTextFrameas = xlShape.TextFrame
            xlCharacters = xlTextFrameas.Characters
            xlCharacters.Text = dr_label(6)("label_content").ToString
            ClsCommon.MRComObject(xlShape)
            ClsCommon.MRComObject(xlTextFrameas)
            ClsCommon.MRComObject(xlCharacters)

            'ご担当セット
            xlShape = xlShapes.Item(strXlPersonName)
            xlTextFrameas = xlShape.TextFrame
            xlCharacters = xlTextFrameas.Characters
            xlCharacters.Text = dr_label(7)("label_content").ToString
            ClsCommon.MRComObject(xlShape)
            ClsCommon.MRComObject(xlTextFrameas)
            ClsCommon.MRComObject(xlCharacters)

            '注文書NOセット
            xlShape = xlShapes.Item(strXlOrderNo)
            xlTextFrameas = xlShape.TextFrame
            xlCharacters = xlTextFrameas.Characters
            xlCharacters.Text = dr_label(8)("label_content").ToString
            ClsCommon.MRComObject(xlShape)
            ClsCommon.MRComObject(xlTextFrameas)
            ClsCommon.MRComObject(xlCharacters)

            '記号セット/品名セット
            xlRange1 = xlCells(12, 1)
            xlRange1.Value = dr_label(9)("label_content").ToString
            ClsCommon.MRComObject(xlRange1)

            '形番セット
            xlRange1 = xlCells(12, 2)
            xlRange1.Value = dr_label(10)("label_content").ToString
            ClsCommon.MRComObject(xlRange1)

            Select Case strMode
                Case "11"
                    '設置位置NOセット
                    xlRange1 = xlCells(11, 11)
                    xlRange1.Value = dr_label(11)("label_content").ToString
                    ClsCommon.MRComObject(xlRange1)
                Case "03", "04"
                    '継手ＣＸセット
                    xlRange1 = xlCells(11, 6)
                    xlRange1.Value = dr_label(13)("label_content").ToString
                    ClsCommon.MRComObject(xlRange1)

                    '設置位置NOセット
                    xlRange1 = xlCells(11, 10)
                    xlRange1.Value = dr_label(11)("label_content").ToString
                    ClsCommon.MRComObject(xlRange1)
                Case Else
                    '設置位置NOセット
                    xlRange1 = xlCells(11, 6)
                    xlRange1.Value = dr_label(11)("label_content").ToString
                    ClsCommon.MRComObject(xlRange1)
            End Select

            Select Case strMode
                Case "12"
                    '使用数セット
                    xlRange1 = xlCells(12, 18)
                    xlRange1.Value = dr_label(12)("label_content").ToString
                    ClsCommon.MRComObject(xlRange1)
                Case "17"
                    '使用数セット
                    xlRange1 = xlCells(12, 11)
                    xlRange1.Value = dr_label(12)("label_content").ToString
                    ClsCommon.MRComObject(xlRange1)
                Case "03", "04"
                    '使用数セット
                    xlRange1 = xlCells(12, 30)
                    xlRange1.Value = dr_label(12)("label_content").ToString
                    ClsCommon.MRComObject(xlRange1)
                Case "09", "13", "16"
                    '使用数セット
                    xlRange1 = xlCells(12, 56)
                    xlRange1.Value = dr_label(12)("label_content").ToString
                    ClsCommon.MRComObject(xlRange1)
                Case Else
                    '使用数セット
                    xlRange1 = xlCells(12, 31)
                    xlRange1.Value = dr_label(12)("label_content").ToString
                    ClsCommon.MRComObject(xlRange1)
            End Select
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            '開放処理
            ClsCommon.MRComObject(xlShapes)
            ClsCommon.MRComObject(xlCells)
            ClsCommon.MRComObject(xlRange1)
        End Try
    End Sub

    ''' <summary>
    ''' 検査成績書表示文字変換
    ''' </summary>
    ''' <param name="strVal"></param>
    ''' <param name="strLangCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncInspReportComb(ByVal strVal As String, ByVal strLangCd As String) As String
        Dim strComb As String
        Try
            'MODIFIED BY YGY 20150330
            If strVal = CdCst.Manifold.InspReportJp.SelectValue Then
                If strLangCd = CdCst.LanguageCd.Japanese Then
                    strComb = CdCst.Manifold.InspReportJp.Japanese
                Else
                    strComb = CdCst.Manifold.InspReportJp.English
                End If
            ElseIf strVal = CdCst.Manifold.InspReportEn.SelectValue Then
                If strLangCd = CdCst.LanguageCd.Japanese Then
                    strComb = CdCst.Manifold.InspReportEn.Japanese
                Else
                    strComb = CdCst.Manifold.InspReportEn.English
                End If
            Else
                strComb = strVal
            End If
            fncInspReportComb = strComb
        Catch ex As Exception
            fncInspReportComb = strVal
        End Try
    End Function

    ''' <summary>
    ''' 品名マスタ取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSpecNo"></param>
    ''' <param name="strLangCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncSelSpecInfoData(ByVal objCon As SqlConnection, ByVal strSpecNo As String, _
                                        ByVal strLangCd As String) As DataTable
        Dim objCmd As New SqlCommand
        Dim objAdp As SqlDataAdapter
        fncSelSpecInfoData = New DataTable
        Dim sbSql As New StringBuilder

        Try
            'DB接続
            objCmd = objCon.CreateCommand
            sbSql.Append("SELECT ISNULL(IT.label_content,EN.label_content) AS label_content, ")
            sbSql.Append("       ISNULL(IT.item_num,EN.item_num) AS item_num, ")
            sbSql.Append("       ISNULL(IT.item_div,EN.item_div) AS item_div, ")
            sbSql.Append("       ISNULL(IT.label_seq,EN.label_seq) AS label_seq ")
            sbSql.Append("FROM	 kh_item_mst EN ")
            sbSql.Append("LEFT OUTER JOIN kh_item_mst IT ")
            sbSql.Append("	ON	 IT.language_cd	= @LangCd ")
            sbSql.Append("	AND	 EN.spec_no		= IT.spec_no ")
            sbSql.Append("	AND	 EN.label_seq	= IT.label_seq ")
            sbSql.Append("WHERE	 EN.language_cd	= @DefaultLang ")
            sbSql.Append("AND	 EN.spec_no	    = @SpecNo ")
            sbSql.Append("ORDER BY EN.label_seq ")

            'SQL実行
            With objCmd
                .CommandText = sbSql.ToString
                .Parameters.Add("@LangCd", SqlDbType.VarChar, 2).Value = strLangCd
                .Parameters.Add("@SpecNo", SqlDbType.VarChar, 2).Value = strSpecNo
                .Parameters.Add("@DefaultLang", SqlDbType.Char, 2).Value = CdCst.LanguageCd.DefaultLang
            End With

            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(fncSelSpecInfoData)

            '<BR>を改行コードに変換
            For Each objrow As DataRow In fncSelSpecInfoData.Rows
                objrow("label_content") = objrow("label_content").ToString.Replace("<BR>", vbLf)
            Next
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
    ''' 簡易マニホールド仕様書出力時の記号変換
    ''' </summary>
    ''' <param name="objRow">品名情報</param>
    ''' <param name="strFullKataban">フル形番</param>
    ''' <returns>変換後の形番</returns>
    ''' <remarks></remarks>
    Private Function fncConvertMarkSimple(ByVal objRow As DataRow, ByVal strFullKataban As String, objKtbnStrc As KHKtbnStrc) As String
        Dim strResult As String = String.Empty
        Dim strSpecNo As String = objKtbnStrc.strcSelection.strSpecNo
        Dim strSeries As String = objKtbnStrc.strcSelection.strSeriesKataban
        Dim strOpSymbol() As String = objKtbnStrc.strcSelection.strOpSymbol

        If strSpecNo = "A1" Or strSpecNo = "A2" Or strSpecNo = "B2" Or strSpecNo = "B3" Or strSpecNo = "B4" Then
            If strOpSymbol(1) <> "8" Then
                If objRow("label_content").ToString.EndsWith("MP") Then
                    objRow("label_content") = String.Empty
                End If
            End If

            If strOpSymbol(1) = "2" Then
                Select Case strSeries
                    Case "M3QRA1", "M3QRB1"
                        If objRow("label_content") = "3QRA119" Then
                            objRow("label_content") = "3QRA129"
                        ElseIf objRow("label_content") = "3QRB119" Then
                            objRow("label_content") = "3QRB129"
                        ElseIf objRow("label_content") = "S1" Then
                            objRow("label_content") = "S2"
                        End If
                End Select
            End If
        End If

        strResult = objRow("label_content")

        Return strResult
    End Function

    ''' <summary>
    ''' 簡易マニホールド仕様書出力時の形番変換
    ''' </summary>
    ''' <param name="objRow">品名情報</param>
    ''' <returns>変換後の形番</returns>
    ''' <remarks></remarks>
    Private Function fncConvertKantabanSimple(ByVal objRow As DataRow, objKtbnStrc As KHKtbnStrc) As String
        Dim strResult As String = String.Empty
        Dim strSpecNo As String = objKtbnStrc.strcSelection.strSpecNo
        Dim strSeries As String = objKtbnStrc.strcSelection.strSeriesKataban
        Dim strFullKataban As String = objKtbnStrc.strcSelection.strFullKataban
        Dim strOpSymbol() As String = objKtbnStrc.strcSelection.strOpSymbol

        If strFullKataban.Substring(0, 3) = "M4F" Then
            Dim strM4FEnd As String = String.Empty

            strM4FEnd = fncCreateEndM4F(strFullKataban)

            If Left(objRow("label_content").ToString, 2) = "4F" Then
                strResult = Left(objRow("label_content"), 4) & strM4FEnd

            ElseIf Left(objRow("label_content").ToString, 3) = "A4F" Then
                strResult = Left(objRow("label_content"), 6) & Right(strM4FEnd, 3)

            Else
                strResult = objRow("label_content")

            End If
        ElseIf strFullKataban.Substring(0, 4) = "M4HA" Then

            Select Case Right(strFullKataban.Substring(0, 5), 1)

                Case "2"
                    strResult = objRow("label_content").ToString.Replace("4HA1", "4HA2")

                Case "3"
                    strResult = objRow("label_content").ToString.Replace("4HA1", "4HA3")

                Case Else
                    strResult = objRow("label_content")
            End Select

        ElseIf strFullKataban.Substring(0, 4) = "M4JA" Then
            Select Case Right(strFullKataban.Substring(0, 5), 1)

                Case "2"
                    strResult = objRow("label_content").ToString.Replace("4JA1", "4JA2")

                Case "3"
                    strResult = objRow("label_content").ToString.Replace("4JA1", "4JA3")

                Case Else
                    strResult = objRow("label_content")

            End Select
        ElseIf strSpecNo = "A1" Or strSpecNo = "A2" Or strSpecNo = "B2" Or strSpecNo = "B3" Or strSpecNo = "B4" Then
            If strOpSymbol(1) <> "8" Then
                If objRow("label_content").ToString.EndsWith("MP") Then
                    objRow("label_content") = String.Empty
                End If
            End If

            If strOpSymbol(1) = "2" Then
                Select Case strSeries
                    Case "M3QRA1", "M3QRB1"
                        If objRow("label_content") = "3QRA119" Then
                            objRow("label_content") = "3QRA129"
                        ElseIf objRow("label_content") = "3QRB119" Then
                            objRow("label_content") = "3QRB129"
                        ElseIf objRow("label_content") = "S1" Then
                            objRow("label_content") = "S2"
                        End If
                End Select
            End If
            strResult = objRow("label_content")
        Else
            strResult = objRow("label_content")
        End If

        Return strResult
    End Function

    ''' <summary>
    ''' M4Fの特殊形番の作成
    ''' </summary>
    ''' <param name="strFullKataban"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncCreateEndM4F(ByVal strFullKataban As String) As String
        Dim strResult As String = String.Empty

        Select Case strFullKataban.Substring(0, 4)
            Case "M4F0"
                If strFullKataban.Substring(5, 1) <> "1" Then
                    If strFullKataban.Substring(7, 2) = "06" Then
                        strResult = "0-06"
                    Else
                        strResult = "0-M5"
                    End If
                Else
                    If strFullKataban.Substring(7, 2) = "06" Then
                        strResult = "1-06"
                    Else
                        strResult = "1-M5"
                    End If
                End If
            Case "M4F1"
                If strFullKataban.Substring(5, 1) <> "1" Then
                    If strFullKataban.Substring(7, 2) = "06" Then
                        strResult = "0-06"
                    Else
                        strResult = "0-08"
                    End If
                Else
                    If strFullKataban.Substring(7, 2) = "06" Then
                        strResult = "1-06"
                    Else
                        strResult = "1-08"
                    End If
                End If
            Case "M4F2"
                If strFullKataban.Substring(5, 1) <> "8" And _
                        strFullKataban.Substring(5, 1) <> "1" Then
                    If strFullKataban.Contains("-C-") Or _
                            strFullKataban.Contains("-I-") Then
                        strResult = "9-08"
                    Else
                        strResult = "0-08"
                    End If
                Else
                    strResult = strFullKataban.Substring(5, 1) & "-08"
                End If
            Case "M4F3"
                If strFullKataban.Substring(7, 1) <> "X" Then
                    If strFullKataban.Substring(6, 1) <> "E" Then
                        If strFullKataban.Substring(5, 1) <> "8" And _
                        strFullKataban.Substring(5, 1) <> "1" Then
                            If strFullKataban.Contains("-C-") Or _
                            strFullKataban.Contains("-I-") Then
                                If strFullKataban.Substring(7, 2) = "08" Then
                                    strResult = "9-08"
                                Else
                                    strResult = "9-10"
                                End If
                            Else
                                If strFullKataban.Substring(7, 2) = "08" Then
                                    strResult = "0-08"
                                Else
                                    strResult = "0-10"
                                End If
                            End If
                        Else
                            If strFullKataban.Substring(7, 2) = "08" Then
                                strResult = strFullKataban.Substring(5, 1) & "-08"
                            Else
                                strResult = strFullKataban.Substring(5, 1) & "-10"
                            End If
                        End If
                    Else
                        strResult = "0E"
                    End If
                Else
                    strResult = "0EX"
                End If
            Case "M4F4", "M4F5"
                'RM1312084 2013/12/24
                If strFullKataban.Substring(7, 1) <> "X" Then
                    If strFullKataban.Substring(6, 1) <> "E" Then
                        If strFullKataban.Substring(5, 1) <> "8" Then
                            strResult = "9-00"
                        Else
                            strResult = "8-00"
                        End If
                    Else
                        strResult = "9E"
                    End If
                Else
                    strResult = "9EX"
                End If
            Case "M4F6"
                'RM1312084 2013/12/24
                If strFullKataban.Substring(7, 1) <> "X" Then
                    If strFullKataban.Substring(6, 1) <> "E" Then
                        If strFullKataban.Substring(5, 1) <> "8" Then
                            strResult = "9-D00"
                        Else
                            strResult = "8-D00"
                        End If
                    Else
                        strResult = "9E-D15"
                    End If
                Else
                    strResult = "9EX"
                End If
            Case "M4F7"
                'RM1312084 2013/12/24
                If strFullKataban.Substring(7, 1) <> "X" Then
                    If strFullKataban.Substring(6, 1) <> "E" Then
                        If strFullKataban.Substring(5, 1) <> "8" Then
                            strResult = "9-E00"
                        Else
                            strResult = "8-E00"
                        End If
                    Else
                        strResult = "9E-E20"
                    End If
                Else
                    strResult = "9EX"
                End If
            Case Else
        End Select

        Return strResult
    End Function

    ''' <summary>
    ''' 仕様書出力(ManifoldExport)
    ''' </summary>
    ''' <param name="language"></param>
    ''' <param name="sysId"></param>
    ''' <param name="lstData"></param>
    ''' <param name="dtItem"></param>
    ''' <param name="strUserId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncExportManifold(ByVal language As String, _
                                      ByVal sysId As Define.SystemID, _
                                      ByVal lstData As List(Of CKDStandard.ManifoldExport.Data.ManifoldBaseData), _
                                      ByVal dtItem As DataTable, _
                                      strUserId As String) As Boolean

        Dim blnResult As Boolean = False
        Dim manifold As ManifoldEntity
        Dim dataConstructor As ManifoldConstructor
        Dim fileCreator As IFileCreate

        Dim outputFile As New File(HttpContext.Current.Server.MapPath("TempFiles"), _
                                   strUserId & "_Manifold", _
                                   Define.FileType.EXCELX)

        Dim templateFile As New File(IO.Path.Combine(HttpContext.Current.Server.MapPath("ManifoldExport"), "Templates"), _
                                     "Manifold", _
                                     Define.FileType.EXCELX)

        Dim checkTemplateResult As ReturnTypeBoolean

        'データの初期化
        manifold = New ManifoldEntity(language, sysId, lstData, dtItem)

        '出力データ変換オブジェクト
        dataConstructor = ManifoldConstructorFactory.Create(manifold)

        '出力方法の初期化
        fileCreator = New FileCreatorExcel(outputFile, templateFile, dataConstructor)

        checkTemplateResult = fileCreator.CheckTemplateFile()

        If checkTemplateResult.Result Then

            'ファイルの作成
            Dim createResult As ReturnTypeBoolean = fileCreator.Create()

            If createResult.Result Then

                blnResult = True

            Else

                blnResult = False

            End If

        Else

            blnResult = False

        End If

        Return blnResult

    End Function

End Class
