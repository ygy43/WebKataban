Imports System.Drawing
Imports Microsoft.VisualBasic
Imports System.Data.SqlClient

Public Class ClsCommon

    ''' <summary>
    ''' PanelのScroll Barの位置を再設定する
    ''' </summary>
    ''' <param name="objHid"></param>
    ''' <param name="GVSelect"></param>
    ''' <param name="strName"></param>
    ''' <remarks></remarks>
    Public Shared Sub ReSetScrollBar(ByRef objHid As Web.UI.WebControls.HiddenField, _
                                     GVSelect As Web.UI.WebControls.DataGrid, _
                                     strName As String)
        If objHid.Value.ToString.Length > 0 Then
            Dim script As New System.Text.StringBuilder
            script.Append("<script language=""JavaScript"">")
            script.Append("document.getElementById('" & strName & "_PanelGrid').scrollTop = '")
            script.Append(objHid.Value)
            script.Append("';")
            script.Append("</script>")
            objHid.Value = String.Empty
            ScriptManager.RegisterStartupScript(GVSelect, GVSelect.GetType, "authenticated", script.ToString, False)
        End If
    End Sub

    ''' <summary>
    ''' PanelのScroll Barの位置を再設定する
    ''' </summary>
    ''' <param name="objHid"></param>
    ''' <param name="GVSelect"></param>
    ''' <param name="strName"></param>
    ''' <remarks></remarks>
    Public Shared Sub ReSetScrollBar(ByRef objHid As Web.UI.WebControls.HiddenField, _
                                     GVSelect As Web.UI.WebControls.GridView, _
                                     strName As String)
        If objHid.Value.ToString.Length > 0 Then
            Dim script As New System.Text.StringBuilder
            script.Append("<script language=""JavaScript"">")
            script.Append("document.getElementById('" & strName & "_PanelGrid').scrollTop = '")
            script.Append(objHid.Value)
            script.Append("';")
            script.Append("</script>")
            objHid.Value = String.Empty
            ScriptManager.RegisterStartupScript(GVSelect, GVSelect.GetType, "authenticated", script.ToString, False)
        End If
    End Sub

    ''' <summary>
    ''' 確認メッセージ
    ''' </summary>
    ''' <param name="strMsg"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function strConfirm(strMsg As String) As String
        Dim sbScript As New StringBuilder
        sbScript.Append(" if (!LogOffConfirm('" & strMsg & "')) {" & vbCrLf)
        sbScript.Append("     return false;" & vbCrLf)
        sbScript.Append(" }" & vbCrLf)
        'sbScript.Append(" return false;")
        strConfirm = sbScript.ToString
    End Function

    ''' <summary>
    ''' フォントの設定
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="strLanguage"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetFontName(obj As Object, strLanguage As String)
        Try
            obj.Font.Name = GetFontName(strLanguage)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' フォントをBoldに設定
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <remarks></remarks>
    Public Shared Sub SetFontBold(obj As Object)
        Try
            obj.Font.Bold = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 対応言語のフォントを取得
    ''' </summary>
    ''' <param name="strLanguage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetFontName(strLanguage As String) As String
        GetFontName = "Calibri"
        Select Case strLanguage.ToUpper
            Case "JA"
                GetFontName = "ＭＳ ゴシック"
            Case "TW", "ZH"
                GetFontName = "SimSun"
        End Select
    End Function

    ''' <summary>
    ''' コントロールのスタイルを設定
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="intMode">1：Label Readonly,2：Input Readonly,3：Input</param>
    ''' <remarks></remarks>
    Public Shared Sub SetAttributes(ByRef obj As Object, intMode As Integer)
        Select Case obj.GetType.Name.ToUpper
            Case "LABEL"
                Select Case intMode
                    Case 0 'WebUCRodEnd Other
                        obj.Style.Add("background-color", "#CCCCCC")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("text-align", "left")
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("font-weight", "normal")
                        obj.Style.Add("vertical-align", "middle")
                        obj.Style.Add("border-style", "none")
                        obj.Style.Add("border-width", "0px")
                    Case 1 'PriceEstimateLabel2
                        obj.Style.Add("background-color", "#008000")
                        obj.Style.Add("color", "#FFFFFF")
                        obj.Style.Add("text-align", "center")
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("font-weight", "bold")
                        obj.Style.Add("vertical-align", "bottom")
                        obj.Style.Add("border-style", "inset")
                        obj.Style.Add("border-width", "thin")
                        obj.Style.Add("height", "15px")
                        obj.Style.Add("padding-top", "2px")
                    Case 2 'DefaultNumLabel
                        obj.Style.Add("background-color", "#FFFFCC")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("text-align", "right")
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("font-weight", "normal")
                        obj.Style.Add("vertical-align", "middle")
                        obj.Style.Add("border-style", "none")
                        obj.Style.Add("border-width", "0px")
                        obj.Style.Add("height", "15px")
                        obj.Style.Add("padding-top", "2px")
                    Case 3 '各種タイトル(中央寄せ)
                        obj.Style.Add("background-color", "#008000")
                        obj.Style.Add("color", "#FFFFFF")
                        obj.Style.Add("text-align", "center")
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("font-weight", "bold")
                        obj.Style.Add("vertical-align", "middle")
                        obj.Style.Add("border-style", "none")
                        obj.Style.Add("border-width", "0px")
                        obj.Style.Add("height", "15px")
                        obj.Style.Add("padding-top", "2px")
                    Case 4 '各種タイトル(左寄せ)
                        obj.Style.Add("background-color", "#008000")
                        obj.Style.Add("color", "#FFFFFF")
                        obj.Style.Add("text-align", "left")
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("font-weight", "bold")
                        obj.Style.Add("vertical-align", "middle")
                        obj.Style.Add("border-style", "none")
                        obj.Style.Add("border-width", "0px")
                        obj.Style.Add("height", "15px")
                        obj.Style.Add("padding-top", "2px")
                        obj.Style.Add("padding-left", "2px")
                    Case 5 '各種タイトル(右寄せ)
                        obj.Style.Add("text-align", "right")
                        obj.Style.Add("font-weight", "bold")
                        obj.Style.Add("vertical-align", "middle")
                        obj.Style.Add("border-style", "none")
                        obj.Style.Add("border-width", "0px")
                        obj.Style.Add("padding-top", "2px")
                        obj.Style.Add("padding-right", "2px")
                    Case 6 'Copy Label
                        obj.Style.Add("background-color", "#FFFFFF")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("text-align", "right")
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("font-weight", "normal")
                        obj.Style.Add("vertical-align", "middle")
                        obj.Style.Add("border-style", "none")
                        obj.Style.Add("border-width", "0px")
                        obj.Style.Add("height", "15px")
                        obj.Style.Add("padding-top", "2px")
                End Select
            Case "CTLCHARTEXT" '入力項目（文字）
                obj.Style.Add("background-color", "#FFFFCC")
                obj.Style.Add("color", "#000000")
                obj.Style.Add("text-align", "left")
                obj.Style.Add("text-transform", "uppercase")
                obj.Style.Add("font-size", "10pt")
                obj.Style.Add("font-weight", "bold")
                obj.Style.Add("vertical-align", "bottom")
            Case "CTLNUMTEXT" '入力項目（数字）
                If intMode = 0 Then obj.Style.Add("background-color", "#FFFFCC")
                obj.Style.Add("color", "#000000")
                Select Case intMode
                    Case 9
                        obj.Style.Add("text-align", "center")
                    Case 8
                        obj.Style.Add("text-align", "left")
                    Case Else
                        obj.Style.Add("text-align", "right")
                End Select
                'obj.Style.Add("font-size", "12pt")
                obj.Style.Add("font-weight", "bold")
                obj.Style.Add("vertical-align", "bottom")
                obj.Style.Add("text-transform", "uppercase")
            Case "TEXTBOX"
                Select Case intMode
                    Case 1  'Title欄
                        obj.Style.Add("background-color", "#008000")
                        CType(obj, TextBox).ForeColor = Drawing.Color.White
                        obj.Style.Add("text-align", "left")
                        obj.Style.Add("text-transform", "uppercase")
                        obj.ReadOnly = True
                        obj.TabIndex = -1
                    Case 2  '表示項目（編集不可）
                        obj.Style.Add("background-color", "#CCFFCC")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("text-align", "right")
                        obj.ReadOnly = True
                        obj.TabIndex = -1
                    Case 3  '入力項目（文字欄、大文字）
                        obj.Style.Add("background-color", "#FFFFCC")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("text-align", "left")
                        obj.Style.Add("text-transform", "uppercase")
                    Case 4 'InputNumeric
                        obj.Style.Add("background-color", "#FFFFCC")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("font-weight", "bold")
                        obj.Style.Add("text-align", "right")
                        obj.Style.Add("vertical-align", "bottom")
                        Exit Select
                    Case 5 'InputNumericError
                        obj.Style.Add("background-color", "#FF0000")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("font-weight", "bold")
                        obj.Style.Add("text-align", "right")
                        obj.Style.Add("vertical-align", "bottom")
                        Exit Select
                    Case 6 'InputText
                        obj.Style.Add("background-color", "#FFFFCC")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("font-weight", "bold")
                        obj.Style.Add("text-align", "left")
                        obj.Style.Add("vertical-align", "bottom")
                    Case 7 'InputTextError
                        obj.Style.Add("background-color", "#FF0000")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("font-weight", "bold")
                        obj.Style.Add("text-align", "left")
                        obj.Style.Add("vertical-align", "bottom")
                    Case 9
                        obj.Style.Add("background-color", "#FFFFCC")
                        obj.Style.Add("color", "#000000")
                        obj.Style.Add("text-align", "center")
                        obj.Style.Add("font-size", "10pt")
                        obj.Style.Add("font-weight", "bold")
                        obj.Style.Add("vertical-align", "bottom")
                        obj.Style.Add("text-transform", "uppercase")
                End Select
                obj.Style.Add("font-size", "10pt")
                obj.Style.Add("font-weight", "bold")
                obj.Style.Add("vertical-align", "bottom")
            Case "DROPDOWNLIST"
                obj.Style.Add("background-color", "#FFFFCC")
                obj.Style.Add("color", "#000000")
                obj.Style.Add("font-size", "10pt")
                obj.Style.Add("font-weight", "bold")
                obj.Style.Add("cursor", "hand")
                obj.Style.Add("margin", "0px")
                obj.Style.Add("padding", "0px")
        End Select
    End Sub

    ''' <summary>
    '''  指定した精度の数値に四捨五入します
    ''' </summary>
    ''' <param name="dValue">丸め対象の倍精度浮動小数点数</param>
    ''' <param name="iDigits"> 戻り値の有効桁数の精度</param>
    ''' <returns>iDigits に等しい精度の数値に四捨五入された数値</returns>
    ''' <remarks></remarks>
    Public Shared Function ToHalfAjust(ByVal dValue As Decimal, ByVal iDigits As Integer) As Decimal
        Dim dCoef As Decimal = CType(System.Math.Pow(10, iDigits), Decimal)

        If dValue > 0 Then
            Return CType(System.Math.Floor((dValue * dCoef) + 0.5) / dCoef, Decimal)
        Else
            Return CType(System.Math.Ceiling((dValue * dCoef) - 0.5) / dCoef, Decimal)
        End If
    End Function

    ''' <summary>
    ''' 指定した精度の数値に切り上げします。
    ''' </summary>
    ''' <param name="dValue">丸め対象の倍精度浮動小数点数。</param>
    ''' <param name="iDigits">戻り値の有効桁数の精度。</param>
    ''' <returns>iDigits に等しい精度の数値に切り上げられた数値。</returns>

    Public Shared Function ToRoundUp(ByVal dValue As Decimal, ByVal iDigits As Integer) As Decimal
        Dim dCoef As Decimal = CType(System.Math.Pow(10, iDigits), Decimal)

        If dValue > 0 Then
            Return System.Math.Ceiling(dValue * dCoef) / dCoef
        Else
            Return System.Math.Floor(dValue * dCoef) / dCoef
        End If
    End Function

    ''' <summary>
    ''' 指定した精度の数値に切り捨てします。
    ''' </summary>
    ''' <param name="dValue">丸め対象の倍精度浮動小数点数。</param>
    ''' <param name="iDigits">戻り値の有効桁数の精度。</param>
    ''' <returns>iDigits に等しい精度の数値に切り捨てられた数値。</returns>
    Public Shared Function ToRoundDown(ByVal dValue As Decimal, ByVal iDigits As Integer) As Decimal
        Dim dCoef As Decimal = CType(System.Math.Pow(10, iDigits), Decimal)

        If dValue > 0 Then
            Return System.Math.Floor(dValue * dCoef) / dCoef
        Else
            Return System.Math.Ceiling(dValue * dCoef) / dCoef
        End If
    End Function

    ''' <summary>
    ''' 文字列が全角文字のみかどうか調べます。
    ''' </summary>
    ''' <param name="Value">対象文字列。</param>
    ''' <returns>True：全角文字のみ　False：半角文字を含む</returns>
    Public Shared Function IsZenkaku(ByVal Value As String) As Boolean
        Return Len(Value) * 2 = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(Value)
    End Function

    ''' <summary>
    ''' 文字列が半角文字のみかどうか調べます。
    ''' </summary>
    ''' <param name="Value">対象文字列。</param>
    ''' <returns>True：半角文字のみ　False：全角文字を含む</returns>
    Public Shared Function IsHankaku(ByVal Value As String) As Boolean
        Return Len(Value) = System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(Value)
    End Function

    ''' <summary>
    ''' 半角 1 バイト、全角 2 バイトとして、指定された文字列のバイト数を返します。
    ''' </summary>
    ''' <param name="stTarget">バイト数取得の対象となる文字列。</param>
    ''' <returns>半角 1 バイト、全角 2 バイトでカウントされたバイト数。</returns>
    Public Shared Function LenB(ByVal stTarget As String) As Integer
        Return System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(stTarget)
    End Function

    ''' <summary>
    ''' ファイルに使用できない文字を削除する。
    ''' </summary>
    ''' <param name="strPath">パス。</param>
    ''' <returns>エラー文字を削除したパス。</returns>
    Public Shared Function DeleteErrorString(ByVal strPath As String) As String
        strPath = strPath.Replace("\", String.Empty)
        strPath = strPath.Replace("/", String.Empty)
        strPath = strPath.Replace(":", String.Empty)
        strPath = strPath.Replace(",", String.Empty)
        strPath = strPath.Replace(";", String.Empty)
        strPath = strPath.Replace("*", String.Empty)
        strPath = strPath.Replace("?", String.Empty)
        strPath = strPath.Replace("""", String.Empty)
        strPath = strPath.Replace("<", String.Empty)
        strPath = strPath.Replace(">", String.Empty)
        strPath = strPath.Replace("|", String.Empty)

        Return strPath
    End Function

    ''' <summary>
    ''' プロセスを実行します。
    ''' </summary>
    ''' <param name="strExeName">実行ファイルパス。</param>
    ''' <param name="strParam">起動パラメータ。</param>
    ''' <param name="blnWait">True:終了まで待つ　False：終了まで待たない</param>
    ''' <param name="intWaitTime">待ち時間（ミリ秒）</param>
    ''' <param name="blnCreateNoWindow">True:ウインドウ非表示　False：ウインドウ表示</param>
    Public Shared Sub ExecuteProcess(ByVal strExeName As String, Optional ByVal strParam As String = Nothing, Optional ByVal blnWait As Boolean = True, Optional ByVal intWaitTime As String = Nothing, Optional ByVal blnCreateNoWindow As Boolean = True)
        Dim Pro As New Process
        Pro.StartInfo.FileName = strExeName                             'ファイル名
        Pro.StartInfo.UseShellExecute = False                           'シェルを使うか
        Pro.StartInfo.CreateNoWindow = blnCreateNoWindow                'ウインドウ非表示
        If Not strParam Is Nothing Then                                 'パラメータの設定
            Pro.StartInfo.Arguments = strParam
        End If
        Pro.Start()                                                     'exeを起動
        If blnWait Then
            If Not IsNothing(intWaitTime) Then
                Pro.WaitForExit(CLng(intWaitTime))                      '終了するか、時間がくるまで待つ
            Else
                Pro.WaitForExit()                                       '終了まで待つ
            End If
        End If
    End Sub

    ''' <summary>
    ''' メッセージ表示
    ''' </summary>
    ''' <param name="strMessageID"></param>
    ''' <param name="strLanguage"></param>
    ''' <param name="strMsgValue"></param>
    ''' <param name="myex"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function WriteErrorLog(strMessageID As String, strLanguage As String, _
                                       Optional ByVal strMsgValue As String = "", _
                                       Optional ByVal myex As Object = Nothing) As Boolean
        WriteErrorLog = False
        Try
            If strMessageID.Equals(String.Empty) Then Exit Function
            If strLanguage.Equals(String.Empty) Then strLanguage = "ja"

            'メッセージ内容の取得
            Dim strMessage As String = String.Empty
            If strMessage.Length <= 0 Then
                strMessage = ClsCommon.fncGetMsg(strLanguage, strMessageID)
            End If
            If strMessageID = "E001" Then strMessage = "[1]"
            If strMessage.Length <= 0 Then strMessage = strMessageID & ControlChars.NewLine

            '埋め込み文字を変換する
            Dim strMsgArray() As String = Nothing
            Dim strReplaceValue As String
            strMsgArray = Split(strMsgValue, ",")
            If strMsgArray IsNot Nothing Then
                For i As Integer = 0 To strMsgArray.Length - 1
                    strReplaceValue = "[" + (i + 1).ToString + "]"
                    strMessage = strMessage.Replace(strReplaceValue, strMsgArray(i))
                Next
            End If
            'ADD BY YGY 20141029
            'エラーメッセージを表示する
            'MsgBox(strMessage, MsgBoxStyle.OkOnly, "形引システム")
            'ScriptManager.RegisterStartupScript(Page, Page.GetType(), "Alert", "alert('" & strMessage & "');", True)
            If Not myex Is Nothing Then
                strMessage &= ControlChars.NewLine
                strMessage &= "関数名：" & myex.TargetSite.Name
                subWriteLog(" -------------------エラー 開始------------------- ")
                subWriteLog(myex.Message)
                subWriteLog(myex.StackTrace)
                subWriteLog(" -------------------エラー 終了------------------- ")
            End If

            WriteErrorLog = True
        Catch ex As Exception
            subWriteLog(" -------------------エラー 開始------------------- ")
            subWriteLog(ex.Message)
            subWriteLog(ex.StackTrace)
            subWriteLog(" -------------------エラー 終了------------------- ")
        End Try
    End Function

    ''' <summary>
    ''' メッセージ表示
    ''' </summary>
    ''' <param name="strMessageID"></param>
    ''' <param name="myex"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function WriteErrorLog(strMessageID As String, myex As Object) As Boolean
        WriteErrorLog = False
        Try
            WriteErrorLog = WriteErrorLog(strMessageID, "ja", myex.Message, myex)
        Catch ex As Exception
            subWriteLog(" -------------------エラー 開始------------------- ")
            subWriteLog(ex.Message)
            subWriteLog(ex.StackTrace)
            subWriteLog(" -------------------エラー 終了------------------- ")
        End Try
    End Function

    ''' <summary>
    ''' システム　ログファイルに出力
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub subWriteLog(ByVal strLog As String)
        Dim strPath As String = My.Settings.LogFolder
        strPath &= Now.ToString("yyyyMMdd") & "_Error_SYSTEM.txt"
        strLog = Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture) & _
                    "/" & _
                    Now.ToString("HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture) & " " & _
                    strLog & System.Environment.NewLine

        IO.File.AppendAllText(strPath, strLog, System.Text.Encoding.Default)
    End Sub

    ''' <summary>
    ''' IDによりﾒｯｾｰｼﾞ情報の取得
    ''' </summary>
    ''' <param name="MessageID"></param>
    ''' <param name="dt"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetMsgbyID(ByVal MessageID As String, ByVal dt As DataTable) As String
        GetMsgbyID = String.Empty
        Try
            Dim dr() As DataRow = dt.Select("message_cd='" & MessageID & "'")
            If dr.Length > 0 Then
                GetMsgbyID = dr(0)("message_content").ToString
            End If
        Catch ex As Exception
            GetMsgbyID = String.Empty
        End Try
    End Function

    ''' <summary>
    ''' バージョン情報の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetVerNo() As String
        Dim VersionNo As System.Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
        GetVerNo = VersionNo.Major.ToString & "." & VersionNo.Minor.ToString & "." & _
                   VersionNo.Build.ToString & "." & VersionNo.Revision.ToString
    End Function

    ''' <summary>
    ''' ｶﾝﾏ除去
    ''' </summary>
    ''' <param name="strPrice"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteKama(ByVal strPrice As String) As String
        DeleteKama = String.Empty
        For inti As Integer = 1 To Len(strPrice)
            If Mid(strPrice, inti, 1) <> "," Then
                DeleteKama &= Mid(strPrice, inti, 1)
            End If
        Next
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="strPath"></param>
    ''' <param name="strValue"></param>
    ''' <param name="flgTime"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LogFileWrite(ByVal strPath As String, ByVal strValue As String, _
                                        Optional ByVal flgTime As Boolean = True) As Boolean
        LogFileWrite = False
        Try
            If flgTime Then
                System.IO.File.AppendAllText(strPath, Now & ControlChars.Tab & strValue & ControlChars.NewLine)
            Else
                System.IO.File.AppendAllText(strPath, strValue & ControlChars.NewLine)
            End If
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxResult.Ok, "システムエラー")
        End Try
        LogFileWrite = True
    End Function

    ''' <summary>
    ''' テスト出力用
    ''' </summary>
    ''' <param name="strDir"></param>
    ''' <param name="strValue"></param>
    ''' <remarks></remarks>
    Public Shared Sub WriteLog(strDir As String, strValue As String)
        '対象フォルダを存在しない場合、新規作成する
        Dim strTime As String = Now.ToShortTimeString
        System.IO.File.AppendAllText(strDir, strTime & ControlChars.Tab & strValue & ControlChars.NewLine, System.Text.Encoding.UTF8)
    End Sub

    ''' <summary>
    ''' 列名によりテーブルを作成
    ''' </summary>
    ''' <param name="strColumnNames">列名</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncCreateTableByColumnNames(ByVal strColumnNames As List(Of String)) As DataTable
        Dim dtResult As New DataTable

        Try
            For Each strColumnName In strColumnNames
                If Not strColumnName.Equals(String.Empty) Then
                    Dim dc As New DataColumn(strColumnName)
                    dtResult.Columns.Add(dc)
                End If
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return dtResult
    End Function

    ''' <summary>
    '''  String型の数値を比較する
    ''' </summary>
    ''' <param name="strOld"></param>
    ''' <param name="strNew"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncCompareStrInteger(ByVal strOld As String, ByVal strNew As String) As Boolean
        Dim intOld As Integer
        Dim intNew As Integer

        Int32.TryParse(strOld, intOld)
        Int32.TryParse(strNew, intNew)

        If intNew > intOld Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' 文字列の「ひらがな→カタカナ」「全角カナ→半角カナ」変換
    ''' </summary>
    ''' <param name="strCS">変換対象文字列</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' 引数で渡された文字列に対し、「ひらがな→カタカナ」「全角カナ→半角カナ」変換を行う。
    ''' 変換後、全角文字が含まれている場合はエラーとする。（※漢字などが含まれた場合）
    ''' </remarks>
    Public Shared Function fncCnvNarrow(ByRef strCS As String) As Boolean
        Dim strRet As String
        Dim intLenb As Integer
        fncCnvNarrow = False
        Try
            strRet = StrConv(strCS, VbStrConv.Narrow + VbStrConv.Katakana)
            'LenBが存在しない為、代用
            intLenb = System.Text.Encoding.GetEncoding("shift-jis").GetByteCount(strRet)

            '文字数よりバイト数が多かったらエラー
            If Len(strRet) < intLenb Then Exit Function
            strCS = strRet

            fncCnvNarrow = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' メッセージ取得
    ''' </summary>
    ''' <param name="strLangCd">言語コード</param>
    ''' <param name="strMsgCd">メッセージコード</param>
    ''' <param name="strMsgValue"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' 引数で渡された言語コード、メッセージコードにてメッセージ内容（kh_message_content）を絞り込み
    ''' 取得したメッセージを返す。
    ''' </remarks>
    Public Shared Function fncGetMsg(ByVal strLangCd As String, _
                              ByVal strMsgCd As String, _
                              Optional ByVal strMsgValue() As String = Nothing) As String
        Dim cn As New SqlConnection
        Dim cm As SqlCommand = cn.CreateCommand
        Dim dr As SqlDataReader = Nothing
        Dim strReplaceValue As String
        fncGetMsg = ""

        Try
            cn.ConnectionString = My.Settings.connkhdb
            cm.CommandText = " SELECT  ISNULL(c.message_content, b.message_content) AS message_content " & _
                             " FROM    kh_message a " & _
                             " INNER JOIN  kh_message_content b " & _
                             " ON      a.message_cd = b.message_cd " & _
                             " LEFT  JOIN  kh_message_content c " & _
                             " ON      a.message_cd  = c.message_cd " & _
                             " AND     c.language_cd = @language_cd " & _
                             " WHERE   a.message_cd  = @message_cd " & _
                             " AND     b.language_cd = @def_language_cd "

            cm.Parameters.Add("@language_cd", SqlDbType.Char, 2).Value = strLangCd
            cm.Parameters.Add("@message_cd", SqlDbType.Char, 5).Value = strMsgCd
            cm.Parameters.Add("@def_language_cd", SqlDbType.Char, 2).Value = CdCst.LanguageCd.DefaultLang

            cn.Open()
            dr = cm.ExecuteReader

            If dr.HasRows Then
                dr.Read()
                fncGetMsg = dr("message_content")

                '埋め込み文字を変換する
                If strMsgValue IsNot Nothing Then
                    For i As Integer = 1 To strMsgValue.Length
                        strReplaceValue = "[" + (i).ToString + "]"
                        fncGetMsg = fncGetMsg.Replace(strReplaceValue, strMsgValue(i - 1))
                    Next
                End If
            Else
                If strLangCd = CdCst.LanguageCd.Japanese Then
                    fncGetMsg = CdCst.Message.NotFound.Japanese
                Else
                    fncGetMsg = CdCst.Message.NotFound.English
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If Not dr Is Nothing Then dr.Close()
            If Not cn Is Nothing Then cn.Close()
        End Try
    End Function

    ''' <summary>
    ''' COMオブジェクト開放処理
    ''' </summary>
    ''' <param name="objCom"></param>
    ''' <remarks></remarks>
    Public Shared Sub MRComObject(ByRef objCom As Object)
        Dim intLoopCnt As Integer
        Try
            '提供されたランタイム呼び出し可能ラッパーの参照カウントをデクリメントします
            If Not objCom Is Nothing AndAlso System.Runtime.InteropServices.Marshal.IsComObject(objCom) Then
                Do
                    intLoopCnt = System.Runtime.InteropServices.Marshal.ReleaseComObject(objCom)
                Loop Until intLoopCnt <= 0
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            objCom = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' ドット編集
    ''' </summary>
    ''' <param name="strPrice">価格</param>
    ''' <returns></returns>
    ''' <remarks>ドット編集したものを戻す</remarks>
    Public Shared Function fncPriceDot(ByVal strPrice As String) As String
        Dim intPriceLength As Integer
        Dim intPriceQut As Integer
        Dim intPriceMod As Integer
        Dim intLoopCnt As Integer
        Dim intPosition As Integer
        Dim strEndPrice As String

        fncPriceDot = ""
        strEndPrice = ""

        intPriceLength = Trim(strPrice.Length)
        intPriceQut = intPriceLength \ 3
        intPriceMod = intPriceLength Mod 3
        intPosition = intPriceMod

        If Trim(strPrice) <> "0" And Len(Trim(strPrice)) <> 0 Then
            '価格が0/スペースでない場合
            If intPriceLength <= 3 Then
                '3桁以下のときはそのままの値を返却する
                strEndPrice = strPrice
            Else
                'ドット編集
                If intPriceMod <> 0 Then
                    strEndPrice = Left(strPrice, intPriceMod)
                End If

                For intLoopCnt = 1 To intPriceQut
                    If intLoopCnt = 1 And intPriceMod = 0 Then
                        strEndPrice = Mid(strPrice, intPosition + 1, 3)
                    Else
                        strEndPrice = strEndPrice & "." & Mid(strPrice, intPosition + 1, 3)
                    End If
                    intPosition = intPosition + intLoopCnt * 3
                Next
            End If
        Else
            strEndPrice = "0"
        End If

        fncPriceDot = strEndPrice
    End Function

    ''' <summary>
    ''' 指定された文字列をダブルクォーテーションで括る
    ''' </summary>
    ''' <param name="strValue">文字列値</param>
    ''' <returns>ダブルクォーテーションで括った文字列</returns>
    ''' <remarks></remarks>
    Public Shared Function fncAddQuote(ByVal strValue As String) As String
        fncAddQuote = String.Empty
        Try
            fncAddQuote = ControlChars.Quote & strValue & ControlChars.Quote
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 対象データがnull or EMPTY値の場合、指定デフォルト値に変換する
    ''' </summary>
    ''' <param name="strData">対象データ</param>
    ''' <param name="strDef">デフォルト値</param>
    ''' <returns>対象データ or デフォルト値</returns>
    ''' <remarks></remarks>
    Public Shared Function fncIsInputed(ByVal strData As String, ByVal strDef As String) As String
        fncIsInputed = String.Empty
        Try
            'スペースの除去
            strData = strData.Trim
            If String.IsNullOrEmpty(strData) Then
                Return strDef
            Else
                Return strData
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 仕様書位置情報の変換　10以上はアルファベットに変換する
    ''' </summary>
    ''' <param name="intPosition">仕様書位置情報</param>
    ''' <returns>変換後の仕様書位置情報</returns>
    ''' <remarks></remarks>
    Public Shared Function fncPositionChance(ByVal intPosition As Integer) As String
        Select Case intPosition
            Case 10
                fncPositionChance = "A"
            Case 11
                fncPositionChance = "B"
            Case 12
                fncPositionChance = "C"
            Case 13
                fncPositionChance = "D"
            Case 14
                fncPositionChance = "E"
            Case 15
                fncPositionChance = "F"
            Case 16
                fncPositionChance = "G"
            Case 17
                fncPositionChance = "H"
            Case 18
                fncPositionChance = "I"
            Case 19
                fncPositionChance = "J"
            Case 20
                fncPositionChance = "K"
            Case 21
                fncPositionChance = "L"
            Case 22
                fncPositionChance = "M"
            Case 23
                fncPositionChance = "N"
            Case 24
                fncPositionChance = "O"
            Case 25
                fncPositionChance = "P"
            Case Else
                fncPositionChance = intPosition.ToString
        End Select
    End Function

    ''' <summary>
    ''' 指定位置の文字の入れ替え
    ''' </summary>
    ''' <param name="strOld"></param>
    ''' <param name="intInde"></param>
    ''' <param name="strVal"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function StrNewString(strOld As String, intInde As Integer, strVal As String) As String
        StrNewString = String.Empty
        For inti As Integer = 0 To strOld.Length - 1
            If inti = intInde Then
                StrNewString &= strVal
            Else
                StrNewString &= strOld(inti)
            End If
        Next
    End Function

    ''' <summary>
    ''' アルファベットと数字のチェック
    ''' </summary>
    ''' <param name="strChkData"></param>
    ''' <param name="strKigo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncAlpNumChk(ByVal strChkData As String, _
        Optional ByVal strKigo() As String = Nothing) As Boolean
        Dim strData As String
        Dim intLen As Integer

        Try
            intLen = Len(strChkData)
            '1文字ずつﾁｪｯｸする
            For idx As Integer = 1 To intLen
                strData = Mid(strChkData, idx, 1)
                '数字ｱﾙﾌｧﾍﾞｯﾄﾁｪｯｸ
                If Not IsNumeric(strData) Then
                    If strData.CompareTo("a") >= 0 And strData.CompareTo("Z") <= 0 Then
                    Else
                        Dim isKigo As Boolean = False
                        '引数．記号一覧
                        If strKigo IsNot Nothing Then
                            '引数．記号一覧を検索
                            For i As Integer = 0 To strKigo.Length - 1
                                '記号が一致した場合
                                If strData.Equals(strKigo(i)) Then
                                    isKigo = True
                                    Exit For
                                End If
                            Next
                        End If
                        Return isKigo
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 列区分の変換
    ''' </summary>
    ''' <param name="strColumnKBN"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncConvertColumnKBN(ByVal strColumnKBN As String) As String
        Dim strResult As String = String.Empty

        Select Case strColumnKBN
            Case "ListPrice"
                strResult = FileOutputColumns.ListPrice
            Case "RegPrice"
                strResult = FileOutputColumns.RegPrice
            Case "SsPrice"
                strResult = FileOutputColumns.SsPrice
            Case "BsPrice"
                strResult = FileOutputColumns.BsPrice
            Case "GsPrice"
                strResult = FileOutputColumns.GsPrice
            Case "PsPrice"
                strResult = FileOutputColumns.PsPrice
            Case "APrice"
                strResult = FileOutputColumns.APrice
            Case "FobPrice"
                strResult = FileOutputColumns.FobPrice
        End Select

        Return strResult

    End Function

    ''' <summary>
    ''' 生産レベルを分解
    ''' </summary>
    ''' <param name="intPlaceLevel">生産レベル</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSeperatePlaceLevel(ByVal intPlaceLevel As Integer) As List(Of Integer)
        '分解結果
        Dim lstResult As New List(Of Integer)
        '生産レベル
        Dim constLevel As New List(Of Integer) From {1, 2, 4, 8, 16, 32, 64, 128, 256, 512, 1024}

        For index As Integer = constLevel.Count - 1 To 0 Step -1

            If intPlaceLevel >= constLevel(index) Then

                lstResult.Add(constLevel(index))

                intPlaceLevel -= constLevel(index)

            End If

        Next

        Return lstResult

    End Function

End Class
