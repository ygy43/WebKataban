/*******************************************************************************
【関数名】  : fncGridClick
【概要】    : 各画面のリスト選択イベント
【引数】    : strName         <String>    画面名
strID           <String>    画面ID
intStartID      <Integer>   
intMode         <Integer>   0：機種画面　1：オプション　2：ISO価格　3：価格
【戻り値】  : 無し
*******************************************************************************/
function fncGridClick(strName, strID, intStartID, intMode) {
    var strRowID;
    var strNowID;
    var decPrc;
    var strPrice;
    strRowID = strID.substr(0, strID.length - 2);
    //リストの全行を一旦未選択状態にする
    for (inti = intStartID; inti <= intStartID + 200; inti++) {
        if (inti <= 9) {
            strNowID = strRowID + "0" + inti;
        } else {
            strNowID = strRowID + inti;
        }
        if (document.getElementById(strNowID)) {
            if (strNowID == strID) {
                document.getElementById(strNowID).style.backgroundColor = "#003C80";
                document.getElementById(strNowID).style.color = "white";
                document.getElementById(strName + "HidSelRowID").value = inti;
                switch (intMode) {
                    case 0: //機種画面
                        document.getElementById(strNowID).focus();
                        break;
                    case 1: //要素画面
                        document.getElementById(strNowID).focus();
                        break;
                    case 2:
                        //選択された単価を隠しエリアに保持       ISO単価
                        //CHANGED BY YGY 20141014 FIREFOX対応
                        //strPrice = document.getElementById(strNowID).cells(1).textContent.split("(");
                        strPrice = document.getElementById(strNowID).cells[1].firstChild.nodeValue.split("(");
                        strPrice = fncRemoveEmptyString(strPrice);    //ADD BY YGY 20141014

                        if (strPrice.length == 2) {
                            var strParentName;
                            if (strName.length == 47) {
                                strParentName = strName.substr(0, strName.length - 12);
                            } else {
                                strParentName = strName.substr(0, strName.length - 11);
                            }
                            if (document.getElementById(strParentName + "txt_EditNormal")) {
                                if (document.getElementById(strParentName + "txt_EditNormal").value == "0") {
                                    decPrc = fncRemoveComma(strPrice[0]);
                                } else {
                                    decPrc = fncRemoveDot(strPrice[0]);
                                }
                            }
                            //選択された単価を隠しエリアに保持
                            document.getElementById(strName + "SelUnitValue").value = decPrc;
                            document.getElementById(strName + "SelCurrValue").value = strPrice[1].substr(0, strPrice[1].length - 1);
                            //金額を再計算する
                            fncPrcLstOnchange(strName);
                        }
                        break;
                    case 3: //単価画面
                        //CHANGED BY YGY 20141014 FIREFOX対応
                        //strPrice = document.getElementById(strNowID).cells(1).textContent.split(" ");
                        document.getElementById(strName + 'GVPrice').focus();
                        strPrice = document.getElementById(strNowID).cells[1].firstChild.nodeValue.split(" ");
                        strPrice = fncRemoveEmptyString(strPrice);    //ADD BY YGY 20141014

                        if (strPrice.length == 2) {
                            if (document.getElementById(strName + "txt_EditNormal").value == "0") {
                                decPrc = fncRemoveComma(strPrice[0]);
                            } else {
                                decPrc = fncRemoveDot(strPrice[0]);
                            }
                            //選択された単価を隠しエリアに保持
                            document.getElementById(strName + "SelUnitValue").value = decPrc;
                            document.getElementById(strName + "SelCurrValue").value = strPrice[1];
                            strCurrency = strPrice[1];

                            //単価ボックスに選択価格をセット
                            document.getElementById(strName + "TextUnitPrice").value = strPrice[0];
                            if (strCurrency == "VND") {
                                document.getElementById(strName + "TextUnitPrice").value = fncSetComma(Math.floor(decPrc / 1000) * 1000);
                            }
                            //掛率・掛単価を算出
                            if (document.getElementById(strName + "txt_EditNormal").value == "0") {
                                document.getElementById(strName + "TextRateUnitPrice").value = fncSetComma(fncRound(decPrc, f_DecLen(2, strCurrency), "."));
                                document.getElementById(strName + "txt_Rate").value = "1.0000";
                            } else {
                                if (document.getElementById(strName + "TextRateUnitPrice").innerText !== undefined) {
                                    document.getElementById(strName + "TextRateUnitPrice").innerText = fncSetDot(fncRound(decPrc, f_DecLen(2, strCurrency), ","));
                                } else {
                                    document.getElementById(strName + "TextRateUnitPrice").textContent = fncSetDot(fncRound(decPrc, f_DecLen(2, strCurrency), ","));
                                }
                                document.getElementById(strName + "txt_Rate").value = "1,0000";
                            }
                            //金額・消費税・合計を計算
                            f_MoneyCal(strName, document.getElementById(strName + "txt_EditNormal").value);
                        }
                        break;
                }
            } else {
                var intCount = parseInt(inti) - parseInt(intStartID) + 1;
                if (intCount % 2 == 0) {
                    strBKcolor = "#CCCCFF";
                } else {
                    strBKcolor = "white";
                }
                document.getElementById(strNowID).style.backgroundColor = strBKcolor;
                document.getElementById(strNowID).style.color = "black";
            }
        } else {
            if (intMode == 3 || intMode == 2) {
                if (inti > 20) {
                    break;
                }
            } else {
                break;    
            }
            
        }
    }
}

/*******************************************************************************
【関数名】  : fncRemoveEmptyString
【概要】    : ブラウザーによりsplitの結果が違うので、一致するように
【詳細】　　: "9,340 JPY"の場合、IE:{"9,340","JPY"} FIREFOX:{"9,340","","JPY"}
【引数】    : strPrice         <String Array>    split結果
【戻り値】  : 空白を削除した結果
*******************************************************************************/
function fncRemoveEmptyString(strPrice) {
    var index = $.inArray("", strPrice);
    if (index != -1) {
        strPrice.splice(index, 1);
    }
    return strPrice;
}

/*******************************************************************************
【関数名】  : fncType_OnKeyup
【概要】    : 各画面のボタン押すイベント
【引数】    : strName         <String>    画面名
strID           <String>    画面ID
intStartID      <Integer>   
intMode         <Integer> 　0：機種画面　1：オプション　2：ISO価格　3：価格
【戻り値】  : 無し
*******************************************************************************/
function fncGrid_OnKeyup(event, strName, strID, intStartID, intMode, rowIndex) {
    var strRowID;
    var strNowID;
    var objEvent;
    objEvent = event || window.event;

    var intNowID;
    var strPrice;
    if (document.getElementById(strName + "HidSelRowID").value.length > 0) {
        strRowID = strID.substr(0, strID.length - 2);
        if (objEvent.keyCode == 40) {// 下矢印押下時
            intNowID = parseInt(document.getElementById(strName + "HidSelRowID").value) + 1;
            if (intNowID <= 9) {
                strNowID = strRowID + "0" + intNowID;
            } else {
                strNowID = strRowID + intNowID;
            }
            if (document.getElementById(strNowID)) {
                switch (intMode) {
                    case 0:
                        fncGridClick(strName, strNowID, intStartID, intMode);
                        break;
                    case 1:
                        fncGridClick(strName, strNowID, intStartID, intMode);
                        break;
                    case 2: //ISO単価
                        if (document.getElementById(strNowID).cells(1).innerText !== undefined) {
                            strPrice = document.getElementById(strNowID).cells(1).innerText.split("(");
                        } else {
                            strPrice = document.getElementById(strNowID).cells(1).textContent.split("(");
                        }
                        if (strPrice.length == 2) { fncGridClick(strName, strNowID, intStartID, intMode); }
                        break;
                    case 3: //単価画面
                        if (document.getElementById(strNowID).cells(1).innerText !== undefined) {
                            strPrice = document.getElementById(strNowID).cells(1).innerText.split(" ");
                        } else {
                            strPrice = document.getElementById(strNowID).cells(1).textContent.split(" ");
                        }
                        if (strPrice.length == 2) { fncGridClick(strName, strNowID, intStartID, intMode); }
                        break;
                }
            }
        } else if (objEvent.keyCode == 38) {// 上矢印押下時
            intNowID = parseInt(document.getElementById(strName + "HidSelRowID").value) - 1;
            if (intNowID <= 9) {
                strNowID = strRowID + "0" + intNowID;
            } else {
                strNowID = strRowID + intNowID;
            }
            if (document.getElementById(strNowID)) {
                fncGridClick(strName, strNowID, intStartID, intMode);
            } else {
                if (intMode == 1) {//要素画面
                    if (document.getElementById(strName + "HidLostID")) {
                        intID = document.getElementById(strName + "HidLostID").value;
                        if (document.getElementById(strName + "txt" + intID)) {
                            document.getElementById(strName + "txt" + intID).focus();
                        }
                    }
                }
            }
        } else if (objEvent.keyCode == 13) {// エンターキー押下時
            if (intMode == 1) {//要素画面
                YousoDblClick(strName, strID, rowIndex);
            }
            else if (intMode == 0) {
                document.getElementById(strName + "btnOK").focus();
            } else if ((intMode == 2) || (intMode == 3)) {//単価、ISO単価画面
                document.getElementById(strName + "txt_Rate").focus();
            }
        } else if (objEvent.keyCode == 9) {// Tabキー押下時
            if (intMode == 2) {//単価画面
                document.getElementById(strName + "txt_Rate").focus();
            }
        }
    }
}

/*******************************************************************************
【関数名】  : fncOverwriteConfirm
【概要】    : I/Fファイル出力の上書き確認
【戻り値】  : 無し
*******************************************************************************/
function fncOverwriteConfirm(strMsg, strNameB, strBtnName) {
    var strName = "ctl00_ContentTitle_";
    if (document.getElementById(strName + "AppFlg")) {
        if (confirm(strMsg)) {
            document.getElementById(strName + "AppFlg").value = true;
            document.getElementById(strNameB + strBtnName).click();
        } else {
            document.getElementById(strName + "AppFlg").value = false;
            document.getElementById(strNameB + strBtnName).click();
        }
    }
}

/*******************************************************************************
【関数名】  : fncDownload
【概要】    : ダウンロードイベント
【戻り値】  : 無し
*******************************************************************************/
function fncDownload(strName) {
    document.getElementById(strName).click();
}

/*******************************************************************************
【関数名】  : fncDisableText
【概要】    : 単価とISO単価画面用
【戻り値】  : 無し
*******************************************************************************/
function fncDisableText(strID) {
    var strKey = strID.split(",");
    for (inti = 0; inti < strKey.length; inti++) {
        if (document.getElementById(strKey[inti])) {
            document.getElementById(strKey[inti]).readOnly = true;
        }
    }
}

/*******************************************************************************
【関数名】  : frmShiftD
【概要】    : ShiftDイベント
【戻り値】  : 無し
*******************************************************************************/
function frmShiftD(strName) {
    if (document.getElementById(strName + "HidShiftD") != null) {
        if (document.getElementById(strName + "HidShiftD").value == "1") {
            document.getElementById(strName + "HidShiftD").value = "2";
            var timer = document.getElementById(strName + "btnCopy");
            if (timer) { timer.click(); }
        }
    }
}

/*******************************************************************************
【関数名】  : GetKataUse
【概要】    : 使用数を取得する
【戻り値】  : 無し
*******************************************************************************/
function GetKataUse(strUse, strName) {
    var KataUse;
    var UseList = "";
    var CXAList = "";
    var CXBList = "";
    var intStart = parseInt(strUse.substr(strUse.length - 2, 2));
    try {
    
    for (inti = intStart; inti < intStart + 50; inti = inti + 1) {
        if (inti <= 9) {
            KataUse = strUse.substr(0, strUse.length - 2) + "0" + inti;
        } else {
            KataUse = strUse.substr(0, strUse.length - 2) + inti;
        }
        //CX値を取得する（あれば）
        if (document.getElementById(KataUse + "_cmbCXA")) {
            CXAList = CXAList + document.getElementById(KataUse + "_cmbCXA").value + ",";   //CXA
        }
        if (document.getElementById(KataUse + "_cmbCXB")) {
            CXBList = CXBList + document.getElementById(KataUse + "_cmbCXB").value + ",";   //CXB
        }
        if (document.getElementById(KataUse + "_txtNum")) {
            UseList = UseList + document.getElementById(KataUse + "_txtNum").value + ",";   //TextBox
        } else {
            var grid = document.getElementById(strName + "_GridViewDetail");
            if (inti <= grid.rows.length) {
                if (grid.rows[inti - 1].cells.length > 0) {
                    //使用数位置の取得
                    if (grid.rows[inti - 1].cells[0].innerText !== undefined) {
                        if (grid.rows[0].cells[0].innerText != "CX A") {
                            //正常の場合
                            var str = grid.rows[inti - 1].cells[0].innerText;
                            UseList = UseList + str.trim() + ",";
                        } else {
                            //CXA・CXBの場合
                            var str = grid.rows[inti - 1].cells[2].innerText;
                            UseList = UseList + str.trim() + ",";
                        }
                    } else {
                        if (grid.rows[0].cells[0].textContent != "CX A") {
                            //正常の場合
                            var str = grid.rows[inti - 1].cells[0].textContent;
                            UseList = UseList + str.trim() + ",";
                        } else {
                            //CXA・CXBの場合
                            var str = grid.rows[inti - 1].cells[2].textContent;
                            UseList = UseList + str.trim() + ",";
                        }
                    }
                } else {
                    UseList = UseList + ",";
                }
            }
        }
    }
    document.getElementById(strName + "_HidUse").value = UseList;
    document.getElementById(strName + "_HidCXA").value = CXAList;
    document.getElementById(strName + "_HidCXB").value = CXBList;
} catch (err) {
    alert(err.LineText)
    }
}

/*******************************************************************************
【関数名】  : Siyou_ReSet
【概要】    : 単価画面→仕様入力、仕様入力画面の再構築
【戻り値】  : 無し
*******************************************************************************/
function Siyou_ReSet(strName) {
    var strID = document.getElementById(strName + "_HidStartID").value
    var strRowID = strID.split(",");

    var strKata = document.getElementById(strName + "_HidSelect").value
    var strKataList = strKata.split(",");
    var strUse = document.getElementById(strName + "_HidUse").value
    var strUseList = strUse.split(",");
    var strCXA = document.getElementById(strName + "_HidCXA").value
    var strCXAList = strCXA.split(",");
    var strCXB = document.getElementById(strName + "_HidCXB").value
    var strCXBList = strCXB.split(",");

    var KataCmb;
    var KataUse;
    var intStart = parseInt(strRowID[0].substr(strRowID[0].length - 2, 2));
    for (inti = intStart; inti < intStart + 50; inti = inti + 1) {
        if (inti <= 9) {
            KataCmb = strRowID[0].substr(0, strRowID[0].length - 2) + "0" + inti;
        } else {
            KataCmb = strRowID[0].substr(0, strRowID[0].length - 2) + inti;
        }
        if (document.getElementById(KataCmb + "_cmbkata")) {
            document.getElementById(KataCmb + "_cmbkata").value = strKataList[inti - intStart];
        }
    }
    intStart = parseInt(strRowID[1].substr(strRowID[1].length - 2, 2));
    for (inti = intStart; inti < intStart + 50; inti = inti + 1) {
        if (inti <= 9) {
            KataUse = strRowID[1].substr(0, strRowID[1].length - 2) + "0" + inti;
        } else {
            KataUse = strRowID[1].substr(0, strRowID[1].length - 2) + inti;
        }
        if (document.getElementById(KataUse + "_txtNum")) {
            document.getElementById(KataUse + "_txtNum").value = strUseList[inti - intStart];   //TextBox
        }
        if (document.getElementById(KataUse + "_cmbCXA")) {
            document.getElementById(KataUse + "_cmbCXA").value = strCXAList[inti - intStart];   //TextBox
        }
        if (document.getElementById(KataUse + "_cmbCXB")) {
            document.getElementById(KataUse + "_cmbCXB").value = strCXBList[inti - intStart];   //TextBox
        }
    }
}

/*******************************************************************************
【関数名】  : TubeChecked
【概要】    : Tubeチェック
【戻り値】  : 無し
*******************************************************************************/
function TubeChecked(strCheckID, strName) {
    if (document.getElementById(strCheckID)) {
        if (document.getElementById(strCheckID).checked) {
            document.getElementById(strName + "HidTube").value = "0";
        } else {
            document.getElementById(strName + "HidTube").value = "";
        }
        var timer = document.getElementById(strName + "_btnClick");
        if (timer) { timer.click(); }
    }
}

/*******************************************************************************
【関数名】  : GridViewCellSelect
【概要】    : 仕様画面形番選択すること
【戻り値】  : 無し
*******************************************************************************/
function GridViewCellSelect(strKata, strtxtID, strMode, strHidRailChangeFlgID) {
    if (document.getElementById(strKata)) {
        if (document.getElementById(strtxtID)) {
            if (strMode == "5" || strMode == "6") {//取付レール長さ
                document.getElementById(strtxtID).value = document.getElementById(strKata).value;
                document.getElementById(strHidRailChangeFlgID).value = "1";
            }
            if (strMode == "99") {                 //タブ銘板
                if (document.getElementById(strKata).value.length == 0) {
                    document.getElementById(strtxtID).value = "";
                } else {
                    document.getElementById(strtxtID).value = "1";
                }
            }
        }
    }
}

/*******************************************************************************
【関数名】  : SetButtonID
【概要】    : 画面IDをセットすること
【戻り値】  : 無し
*******************************************************************************/
function SetButtonID(strID) {
    document.getElementById("ctl00_ContentTitle_HidRunForm").value = strID;
}

/*******************************************************************************
【関数名】  : fncRoundDown
【概要】    : 少数部まるめを行う(切り捨て)
【引数】    : strArgValue        <String>  文字列
            : intDecLen          <Integer> 小数部バイト数
            : strDelim           <String>  整数部と小数部の区切り文字
【戻り値】  : strRetValue        <String>  編集された文字列
*******************************************************************************/
function fncRoundDown(strArgValue, intDecLen, strDelim) {

    var strRetValue = strArgValue;
    var intRate = 1;
    var strInt = "";
    var strDec = "";
    var intDif = 0;
    if ((strArgValue != null) && (strArgValue != "")) {
        if (strDelim == ".") {
            strRetValue = strArgValue;
        } else {
            strRetValue = strArgValue.replace(",", ".");
        }

        //小数第intDecLen桁まで四捨五入
        strRetValue = fncFloatRound(strRetValue, intDecLen);

        //文字型にキャスト
        strRetValue = strRetValue + "";

        //少数部が桁数に満たない場合、「0」埋めする
        if (strRetValue.indexOf(".") == -1) {
            strInt = strRetValue;
            strDec = "";
            intDif = 0;
        }
        else {
            strInt = strRetValue.split(".")[0];
            strDec = strRetValue.split(".")[1];
            intDif = strDec.length;
        }
        if (intDif <= intDecLen) {
            for (i = 0; i < intDecLen - intDif; i++) {
                strDec = strDec + "0";
            }
        }
        //整数部と小数部を結合
        if (intDecLen > 0) { strRetValue = strInt + strDelim + strDec; }

        if (strDelim == ".") {
        } else {
            strRetValue = strRetValue.replace(".", ",");
        }
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : f_JudgeRound
【概要】    : 単価_端数処理判定
【引数】    : strDefRound      <String>  デフォルト端数処理方法
：intValue        <Integer>  計算前の値
：intDecLen       <String>   計算後の小数部バイト数
：strComma        <String>   区切り記号
：strCurrency     <String>   通貨記号
【戻り値】  : intRetValue     <String>   端数処理後の値
*******************************************************************************/
function f_JudgeRound(strDefRound, intValue, intDecLen, strComma, strCurrency) {
    var intRetValue;
    intRetValue = intValue;
    if ((strCurrency == "USD") || (strCurrency == "THB")) {
        strDefRound = "fncRound";
    }
    if (strDefRound == "fncRoundUp") {
        intRetValue = fncRoundUp(intValue, intDecLen, strComma);
    } else if (strDefRound == "fncRoundDown") {
        intRetValue = fncRoundDown(intValue, intDecLen, strComma);
    } else {
        intRetValue = fncRound(intValue, intDecLen, strComma);
    }
    //ベトナムの場合百位切捨てにする
    if (strCurrency == "VND") {
        intRetValue = Math.floor(intRetValue / 1000) * 1000;
    }
    return intRetValue;
}

/*******************************************************************************
【関数名】  : f_DecLen
【概要】    : 小数部バイト数
【引数】    : intDefLen        <Integer>   基準小数部バイト数
            : strCurrency      <String>  　通貨記号
【戻り値】  : intDecLen        <String>  　計算後の小数部バイト数
*******************************************************************************/
function f_DecLen(intDefLen, strCurrency) {
    var intDecLen;

    intDecLen = intDefLen;
    
    if ((strCurrency == "USD") || (strCurrency == "THB")) {
        intDecLen = intDecLen + 2;
    }
    return intDecLen;
}

/*******************************************************************************
【関数名】  : Clip_Copy
【概要】    : コピー画面
*******************************************************************************/
function Clip_Copy(txtValue) {
    //クリップボードへコピー
    window.clipboardData.setData("text", txtValue);
}

/*******************************************************************************
【関数名】  : fncRoundUp
【概要】    : 少数部まるめを行う(切り上げ)
【引数】    : strArgValue        <String>  文字列
: intDecLen          <Integer> 小数部バイト数
: strDelim           <String>  整数部と小数部の区切り文字
【戻り値】  : strRetValue        <String>  編集された文字列
*******************************************************************************/
function fncRoundUp(strArgValue, intDecLen, strDelim) {
    var strRetValue = strArgValue;
    var intRate = 1;
    var strInt = "";
    var strDec = "";
    var intDif = 0;
    if ((strArgValue != null) && (strArgValue != "")) {
        if (strDelim == ".") {
            strRetValue = strArgValue;
        } else {
            strRetValue = strArgValue.replace(",", ".");
        }
        for (i = 0; i < intDecLen; i++) {
            intRate = intRate * 10;
        }
        strRetValue = Math.round(strRetValue * intRate * 10) / 10;
        strRetValue = Math.ceil(strRetValue);
        strRetValue = strRetValue / intRate;
        //文字型にキャスト
        strRetValue = strRetValue + "";

        //少数部が桁数に満たない場合、「0」埋めする
        if (strRetValue.indexOf(".") == -1) {
            strInt = strRetValue;
            strDec = "";
            intDif = 0;
        }
        else {
            strInt = strRetValue.split(".")[0];
            strDec = strRetValue.split(".")[1];
            intDif = strDec.length;
        }
        if (intDif <= intDecLen) {
            for (i = 0; i < intDecLen - intDif; i++) {
                strDec = strDec + "0";
            }
        }
        //整数部と小数部を結合
        if (intDecLen > 0) {
            strRetValue = strInt + strDelim + strDec;
        }
        if (strDelim == ".") {
        } else {
            strRetValue = strRetValue.replace(".", ",");
        }
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncRemoveComma
【概要】    : カンマを除去する
【引数】    : strArgValue         <String>  文字列
【戻り値】  : strRetValue         <String>  編集された文字列
*******************************************************************************/
function fncRemoveComma(strArgValue) {
    var strRetValue = "";
    strArgValue = new String(strArgValue);
    // 文字列をカンマで分割する
    var strSplitValue = strArgValue.split(",");
    // 分割した文字列を連結する
    for (var intLoopCnt = 0; intLoopCnt < strSplitValue.length; intLoopCnt++) {
        strRetValue += strSplitValue[intLoopCnt];
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncRemoveDot
【概要】    : ドットを除去する
【引数】    : strArgValue         <String>  文字列
【戻り値】  : strRetValue         <String>  編集された文字列
*******************************************************************************/
function fncRemoveDot(strArgValue) {
    var strRetValue = "";
    strArgValue = new String(strArgValue);
    // 文字列をドットで分割する
    var strSplitValue = strArgValue.split(".");
    // 分割した文字列を連結する
    for (var intLoopCnt = 0; intLoopCnt < strSplitValue.length; intLoopCnt++) {
        strRetValue += strSplitValue[intLoopCnt];
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncCmnTextOnFocus
【概要】    : テキストボックスフォーカス取得処理
【引数】    : field        <String>  文字列
【戻り値】  : 無し
*******************************************************************************/
function fncCmnTextOnFocus(field) {
    // 背景色を変更する
    field.style.backgroundColor = '#FFCC33';
    // フォーカス設定
    field.focus();
    // テキストを選択状態にする
    field.select();
}

/*******************************************************************************
【関数名】  : fncCmnTextOnBlur
【概要】    : テキストボックスフォーカス喪失処理
【引数】    : field        <String>  文字列
【戻り値】  : 無し
*******************************************************************************/
function fncCmnTextOnBlur(field) {
    // 背景色を変更する
    field.style.backgroundColor = '#FFFFCC';
}

/*******************************************************************************
【関数名】  : fncCheckNum
【概要】    : 引数で渡された文字列のチェックを行う
【引数】    : strCheckData        <String>  チェックフィールド
: intLen              <Integer> 整数部バイト数
: bolMinus            <Boolean> マイナス入力使用可能フラグ
:                               true:使用可／false:使用不可
: bolZero             <Boolean> ゼロ入力使用可能フラグ
:                               true:使用可／false:使用不可
: bolNull             <Boolean> Null禁止フラグ
:                               true:禁止／false:許可
: EditDiv             <String>  Edit区分
【戻り値】  :                     <Boolean> true:OK／false:NG
*******************************************************************************/
function fncCheckNum(strCheckData, intLen, bolMinus, bolZero, bolNull, EditDiv) {
    var intRet;
    var strRet;
    if (EditDiv != "0") {
        strCheckData = strCheckData.replace(",", ".");
    }
    //値がない場合は何もしない 
    if (strCheckData == "") {
        //Null禁止の場合はエラーを返す
        if (bolNull) {
            return false;
        }
        return true;
    }
    //数値チェック
    if (isNaN(strCheckData)) {
        return false;
    }
    //マイナス値のチェック
    if (!bolMinus && strCheckData < 0) {
        return false;
    }
    //ゼロ値のチェック
    if (!bolZero && strCheckData == 0) {
        return false;
    }
    if (strCheckData.indexOf(".") > -1) {
        intRet = strCheckData.split(".")[0].length
    } else {
        intRet = strCheckData.length
    }
    if ((intLen != 0) && intRet > intLen) {
        return false;
    }
    return true;
}

/*******************************************************************************
【関数名】  : fncRound
【概要】    : 少数部まるめを行う
【引数】    : strArgValue        <String>  文字列
: intDecLen          <Integer> 小数部バイト数
: strDelim           <String>  整数部と小数部の区切り文字
【戻り値】  : strRetValue        <String>  編集された文字列
*******************************************************************************/
function fncRound(strArgValue, intDecLen, strDelim) {
    var strRetValue = strArgValue;
    var intRate = 1;
    var strInt = "";
    var strDec = "";
    var intDif = 0;
    if ((strArgValue != null) && (strArgValue != "")) {
        if (strDelim == ".") {
            strRetValue = strArgValue;
        } else {
            strRetValue = strArgValue.replace(",", ".");
        }

        //小数第intDecLen桁まで四捨五入
        strRetValue = fncFloatRound(strRetValue, intDecLen);

        //文字型にキャスト
        strRetValue = strRetValue + "";
        //少数部が桁数に満たない場合、「0」埋めする
        if (strRetValue.indexOf(".") == -1) {
            strInt = strRetValue;
            strDec = "";
            intDif = 0;
        }
        else {
            strInt = strRetValue.split(".")[0];
            strDec = strRetValue.split(".")[1];
            intDif = strDec.length;
        }
        if (intDif <= intDecLen) {
            for (i = 0; i < intDecLen - intDif; i++) {
                strDec = strDec + "0";
            }
        }
        //整数部と小数部を結合
        if (intDecLen > 0) {
            strRetValue = strInt + strDelim + strDec;
        }
        if (strDelim == ".") {
        } else {
            strRetValue = strRetValue.replace(".", ",");
        }
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncSetComma
【概要】    : カンマ編集を行う
【引数】    : strArgValue         <String>  文字列
【戻り値】  : strRetValue         <String>  編集された文字列
*******************************************************************************/
function fncSetComma(strArgValue) {
    var strRetValue = new String(strArgValue);
    var bolMinusFlag = false;
    var intLoopCnt;
    var intNum;
    var intIndex;
    // 値が無い場合はそのまま返す
    if (strRetValue == "") {
        return strRetValue;
    }
    // 先頭文字がハイフンの場合は削除する
    if (strRetValue.substring(0, 1) == "-") {
        strRetValue = strRetValue.substring(1);
        // マイナス値フラグON
        bolMinusFlag = true;
    }
    // 文字列をドットで分割する
    var strSplitValue = strRetValue.split(".");
    // 一番最初に見つかった小数点の左側の文字列を処理対象とする
    strRetValue = strSplitValue[0];
    // カンマが存在している場合は削除する
    strRetValue = fncRemoveComma(strRetValue);
    // カンマの数を算出する
    var intLen = strRetValue.length;
    if (intLen % 3 == 0) {
        intNum = intLen / 3 - 1;
    } else {
        intNum = Math.floor(intLen / 3);
    }
    // 3文字・区切・カンマを付けながら連結する
    intIndex = 3;
    for (intLoopCnt = 0; intLoopCnt < intNum; intLoopCnt++) {
        strRetValue = strRetValue.substring(0, intLen - intIndex) + "," + strRetValue.substring(intLen - intIndex, intLen);
        intLen++;
        intIndex++;
        intIndex += 3;
    }
    // 小数点以下が存在した場合は少数部を連結する
    if (strSplitValue.length > 1) {
        for (intLoopCnt = 1; intLoopCnt < strSplitValue.length; intLoopCnt++) {
            strRetValue += ".";
            strRetValue += strSplitValue[intLoopCnt];
        }
    }
    // 頭がハイフンだった場合は付け直す
    if (bolMinusFlag == true) {
        strRetValue = "-" + strRetValue;
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncSetDot
【概要】    : ドット編集を行う
【引数】    : strArgValue         <String>  文字列
【戻り値】  : strRetValue         <String>  編集された文字列
*******************************************************************************/
function fncSetDot(strArgValue) {
    var strRetValue = new String(strArgValue);
    var bolMinusFlag = false;
    var intLoopCnt;
    var intNum;
    var intIndex;
    // 値が無い場合はそのまま返す
    if (strRetValue == "") {
        return strRetValue;
    }
    // 先頭文字がハイフンの場合は削除する
    if (strRetValue.substring(0, 1) == "-") {
        strRetValue = strRetValue.substring(1);
        // マイナス値フラグON
        bolMinusFlag = true;
    }
    // 文字列をカンマで分割する
    var strSplitValue = strRetValue.split(",");
    // 一番最初に見つかった小数点の左側の文字列を処理対象とする
    strRetValue = strSplitValue[0];
    // ドットが存在している場合は削除する
    strRetValue = fncRemoveDot(strRetValue);
    // カンマの数を算出する
    var intLen = strRetValue.length;
    if (intLen % 3 == 0) {
        intNum = intLen / 3 - 1;
    } else {
        intNum = Math.floor(intLen / 3);
    }
    // 3文字・区切・ドットを付けながら連結する
    intIndex = 3;
    for (intLoopCnt = 0; intLoopCnt < intNum; intLoopCnt++) {
        strRetValue = strRetValue.substring(0, intLen - intIndex) + "." + strRetValue.substring(intLen - intIndex, intLen);
        intLen++;
        intIndex++;
        intIndex += 3;
    }
    // 小数点以下が存在した場合は少数部を連結する
    if (strSplitValue.length > 1) {
        for (intLoopCnt = 1; intLoopCnt < strSplitValue.length; intLoopCnt++) {
            strRetValue += ",";
            strRetValue += strSplitValue[intLoopCnt];
        }
    }
    // 頭がハイフンだった場合は付け直す
    if (bolMinusFlag == true) {
        strRetValue = "-" + strRetValue;
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncTextTrim
【概要】    : 左右の空白を除去する
【引数】    : strArgValue        <String>  文字列
【戻り値】  : strRetValue        <String>  変換された文字列
*******************************************************************************/
function fncTextTrim(strArgValue) {
    var strRetValue = strArgValue;
    // 左側の空白を削除する
    strRetValue = fncTextLTrim(strRetValue);
    // 右側の空白を削除する
    strRetValue = fncTextRTrim(strRetValue);
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncTextRTrim
【概要】    : 右側の空白を除去する
【引数】    : strArgValue        <String>  文字列
【戻り値】  : strRetValue        <String>  変換された文字列
*******************************************************************************/
function fncTextRTrim(strArgValue) {
    // 右側の空白を消去する
    var strRetValue = strArgValue.replace(/^[ 　]+/, "");
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncTextLTrim
【概要】    : 左側の空白を除去する
【引数】    : strArgValue        <String>  文字列
【戻り値】  : strRetValue        <String>  変換された文字列
*******************************************************************************/
function fncTextLTrim(strArgValue) {
    // 左側の空白を消去する
    var strRetValue = strArgValue.replace(/[ 　]+$/, "");
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncTextLCase
【概要】    : 引数で渡された文字列を小文字に変換する
【引数】    : strArgValue        <String>  文字列
【戻り値】  : strRetValue        <String>  大文字変換された文字列
*******************************************************************************/
function fncTextLCase(strArgValue) {
    // 小文字に変換する
    var strRetValue = strArgValue.toLowerCase();
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncTextUCase
【概要】    : 大文字に変換する
【引数】    : strArgValue        <String>  文字列
【戻り値】  : strRetValue        <String>  大文字変換された文字列
*******************************************************************************/
function fncTextUCase(strArgValue) {
    // 大文字に変換する
    var strRetValue = strArgValue.toUpperCase();
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncTextToKana
【概要】    : ひらがなをカナに変換する
【引数】    : strCheckData        <String>  文字列
【戻り値】  : strRetValue         <String>  変換された文字列
*******************************************************************************/
function fncTextToKana(strCheckData) {
    var strRetValue = "";
    var intLen = strCheckData.length;
    for (i = 0; i < intLen; i++) {
        code = strCheckData.charCodeAt(i);
        if (code >= 12353 && code <= 12435) { code = code + 96; }
        strRetValue = strRetValue + String.fromCharCode(code);
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncTextToHankaku
【概要】    : 全角カナを半角カナに変換する
【引数】    : strCheckData        <String>  文字列
【戻り値】  : strRetValue         <String>  変換された文字列
*******************************************************************************/
function fncTextToHankaku(strCheckData) {
    var strRetValue = "";
    var intLen = strCheckData.length;
    var dakuten_moji = String.fromCharCode(65438);
    var handaku_moji = String.fromCharCode(65439);
    var kanahenkan = new Array();
    var hankaku_code = new Array(65383, 65393, 65384, 65394, 65385, 65395, 65386, 65396, 65387, 65397, 65398, 65398, 65399, 65399, 65400, 65400, 65401, 65401, 65402, 65402, 65403, 65403, 65404, 65404, 65405, 65405, 65406, 65406, 65407, 65407, 65408, 65408, 65409, 65409, 65391, 65410, 65410, 65411, 65411, 65412, 65412, 65413, 65414, 65415, 65416, 65417, 65418, 65418, 65418, 65419, 65419, 65419, 65420, 65420, 65420, 65421, 65421, 65421, 65422, 65422, 65422, 65423, 65424, 65425, 65426, 65427, 65388, 65428, 65389, 65429, 65390, 65430, 65431, 65432, 65433, 65434, 65435, 65436, 65436, 65394, 65396, 65382, 65437, 65395, 65398, 65401);    //半角カタカナ作成用濁点文字No.
    var dakuten_no = new Array(11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 36, 38, 40, 47, 50, 53, 56, 59, 83);
    //半角カタカナ作成用半濁点文字No.
    var handaku_no = new Array(48, 51, 54, 57, 60);
    //半角カタカナ用配列を構成する
    for (i = 0; i < 86; i++) {
        kanahenkan[i] = (hankaku_code[i]);
    }
    for (i = 0; i < 21; i++) {
        kanahenkan[dakuten_no[i]] = kanahenkan[dakuten_no[i]] + dakuten_moji;
    }
    for (i = 0; i < 5; i++) {
        kanahenkan[handaku_no[i]] = kanahenkan[handaku_no[i]] + handaku_moji
    }
    for (i = 0; i < intLen; i++) {
        code = strCheckData.charCodeAt(i);
        if (code >= 12449 && code <= 12534) {
            strRetValue = strRetValue + String.fromCharCode(kanahenkan[code - 12449]);
        } else {
            strRetValue = strRetValue + String.fromCharCode(code);
        }
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncTextToHankakuAlp
【概要】    : 全角アルファベットを半角アルファべットに変換する
【引数】    : strCheckData        <String>  文字列
【戻り値】  : strRetValue         <String>  変換された文字列
*******************************************************************************/
function fncTextToHankakuAlp(strCheckData) {
    var strRetValue = "";
    var intLen = strCheckData.length;
    for (i = 0; i < intLen; i++) {
        code = strCheckData.charCodeAt(i);
        if (code >= 65313 && code <= 65338) { code = code - 65248; }
        if (code >= 65345 && code <= 65370) { code = code - 65248; }
        strRetValue = strRetValue + String.fromCharCode(code);
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncTextToHankakuNum
【概要】    : 全角数字を半角数字に変換する
【引数】    : strCheckData        <String>  文字列
【戻り値】  : strRetValue         <String>  変換された文字列
*******************************************************************************/
function fncTextToHankakuNum(strCheckData) {
    var strRetValue = "";
    var intLen = strCheckData.length;
    for (i = 0; i < intLen; i++) {
        code = strCheckData.charCodeAt(i)
        if (code >= 65296 && code <= 65305) { code = code - 65248; }
        strRetValue = strRetValue + String.fromCharCode(code);
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncTextToHankakuKigou
【概要】    : 全角記号を半角記号に変換する
【引数】    : strCheckData        <String>  文字列
【戻り値】  : strRetValue         <String>  変換された文字列
*******************************************************************************/
function fncTextToHankakuKigou(strCheckData) {
    var strHan = "!\"#$%&'()=-~^|\\{}｢｣'@*:_;+?･><｡､";
    var strZen = "！”＃＄％＆’（）＝－～＾｜￥｛｝「」‘＠＊：＿；＋？・＞＜。、";
    var strRetValue = "";
    var intLen = strCheckData.length;
    for (i = 0; i < intLen; i++) {
        c = strCheckData.charAt(i);
        n = strZen.indexOf(c, 0);
        if (n >= 0) c = strHan.charAt(n);
        strRetValue += c;
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncCheckChar
【概要】    : 引数で渡された文字列のチェックを行う
【引数】    : strCheckData        <String>  チェックフィールド
【戻り値】  : intMinCnt           <Integer> 最小バイト数
: intMaxCnt           <Integer> 最大バイト数
: bolDword            <Boolean> 全角文字使用可能フラグ
:                               true:使用可／false:使用不可
:                     <Boolean> true:OK／false:NG
*******************************************************************************/
function fncCheckChar(strCheckData, intMinCnt, intMaxCnt, bolDword) {
    // 禁止文字の設定
    var arr_Pattrn = new Array("\'", "\"", "|", ",");
    var strPattrn;
    var intLen = strCheckData.length;
    var intByte = fncGetByteCount(strCheckData);
    // 空白の時はチェックしない
    if (intLen == 0) {
        return true;
    }
    // 禁止文字チェック
    for (var intLoopCnt = 0; intLoopCnt < intLen; intLoopCnt++) {
        for (var intLoopCnt1 = 0; intLoopCnt1 < arr_Pattrn.length; intLoopCnt1++) {
            strPattrn = arr_Pattrn[intLoopCnt1];
            if (strPattrn.indexOf(strCheckData.charAt(intLoopCnt).toUpperCase()) < 0) {
            } else {
                return false;
            }
        }
    }
    // 最小バイト数チェック
    if (intMinCnt != 0 && intByte < intMinCnt) {
        return false;
    }
    // 最大バイト数チェック
    if (intMaxCnt != 0 && intByte > intMaxCnt) {
        return false;
    }
    // 全角不可の場合はチェックする
    if (!bolDword && !fncCheckTextLength(strCheckData, 0)) {
        return false;
    }
    return true;
}

/*******************************************************************************
【関数名】  : fncGetByteCount
【概要】    : 全角を２バイト、半角を１バイトとしてカウント
【引数】    : strCheckData        <String>  文字列
【戻り値】  : count               <Integer> 大文字変換された文字列
*******************************************************************************/
function fncGetByteCount(strCheckData) {
    var intCount = 0;
    for (var intLoopCnt = 0; intLoopCnt < strCheckData.length; ++intLoopCnt) {
        var strChar = strCheckData.substring(intLoopCnt, intLoopCnt + 1);
        //半角・全角のチェック
        if (fncCheckIsZenkaku(strChar)) {
            intCount += 2;
        } else {
            intCount += 1;
        }
    }
    return intCount;
}

/*******************************************************************************
【関数名】  : fncCheckIsZenkaku
【概要】    : 引数で渡された文字列に半角または全角が含まれているかチェックする
【引数】    : strCheckData        <String>  文字列
【戻り値】  :                     <Boolean> true:全角／false:全角以外
*******************************************************************************/
function fncCheckIsZenkaku(strCheckData) {
    for (var intLoopCnt = 0; intLoopCnt < strCheckData.length; ++intLoopCnt) {
        var strChar = strCheckData.charCodeAt(intLoopCnt);
        //半角カタカナは不許可
        if (strChar < 256 || (strChar >= 0xff61 && strChar <= 0xff9f)) {
            return false;
        }
    }
    return true;
}

/*******************************************************************************
【関数名】  : fncCheckTextLength
【概要】    : 引数で渡された文字列に半角または全角が含まれているかチェックする
【引数】    : strCheckData        <String>  文字列
: intCheckFlag        <Integer> 半角・全角チェックフラグ
:                               0:使用可／1:使用不可
【戻り値】  :                     <Boolean> true:OK／false:NG
*******************************************************************************/
function fncCheckTextLength(strCheckData, intCheckFlag) {
    // 1文字ずつチェックする
    for (var intLoopCnt = 0; intLoopCnt < strCheckData.length; intLoopCnt++) {
        var strChar = strCheckData.charCodeAt(intLoopCnt);
        if ((strChar >= 0x0 && strChar < 0x81) || (strChar == 0xf8f0) || (strChar >= 0xff61 && strChar < 0xffa0) || (strChar >= 0xf8f1 && strChar < 0xf8f4)) {
            // 半角チェックの場合はNG
            if (intCheckFlag) {
                return false;
            }
        } else {
            // 全角チェックの場合はNG
            if (!intCheckFlag) {
                return false;
            }
        }
    }
    return true;
}

/*******************************************************************************
【関数名】  : fncChangeInput
【概要】    : 入力文字の変換
【引数】    : strArgValue        <String>  文字列
【戻り値】  : strRetValue        <String>  変換された文字列
*******************************************************************************/
function fncChangeInput(strArgValue) {
    var strReturn;

    strReturn = fncTextUCase(strArgValue);
    strReturn = fncTextToKana(strReturn);
    strReturn = fncTextToHankaku(strReturn);
    strReturn = fncTextToHankakuAlp(strReturn);
    //    strReturn = fncTextToHankakuNum(strReturn);
    //    strReturn = fncTextToHankakuKigou(strReturn);

    return strReturn;
}

/*******************************************************************************
【関数名】  : fncTextUCase
【概要】    : 大文字に変換する
【引数】    : strArgValue        <String>  文字列
【戻り値】  : strRetValue        <String>  大文字変換された文字列
*******************************************************************************/
function fncTextUCase(strArgValue) {

    // 大文字に変換する
    var strRetValue = strArgValue.toUpperCase();

    return strRetValue;

}
/*******************************************************************************
【関数名】  : fncTextToKana
【概要】    : ひらがなをカナに変換する
【引数】    : strCheckData        <String>  文字列
【戻り値】  : strRetValue         <String>  変換された文字列
*******************************************************************************/
function fncTextToKana(strCheckData) {

    var strRetValue = "";
    var intLen = strCheckData.length;

    for (i = 0; i < intLen; i++) {
        code = strCheckData.charCodeAt(i);
        if (code >= 12353 && code <= 12435) { code = code + 96; }
        strRetValue = strRetValue + String.fromCharCode(code);
    }
    return strRetValue;
}
/*******************************************************************************
【関数名】  : fncTextToHankaku
【概要】    : 全角カナを半角カナに変換する
【引数】    : strCheckData        <String>  文字列
【戻り値】  : strRetValue         <String>  変換された文字列
*******************************************************************************/
function fncTextToHankaku(strCheckData) {

    var strRetValue = "";
    var intLen = strCheckData.length;

    var dakuten_moji = String.fromCharCode(65438);
    var handaku_moji = String.fromCharCode(65439);
    var kanahenkan = new Array();
    var hankaku_code = new Array(65383, 65393, 65384, 65394, 65385, 65395, 65386, 65396, 65387, 65397, 65398, 65398, 65399, 65399, 65400, 65400, 65401, 65401, 65402, 65402, 65403, 65403, 65404, 65404, 65405, 65405, 65406, 65406, 65407, 65407, 65408, 65408, 65409, 65409, 65391, 65410, 65410, 65411, 65411, 65412, 65412, 65413, 65414, 65415, 65416, 65417, 65418, 65418, 65418, 65419, 65419, 65419, 65420, 65420, 65420, 65421, 65421, 65421, 65422, 65422, 65422, 65423, 65424, 65425, 65426, 65427, 65388, 65428, 65389, 65429, 65390, 65430, 65431, 65432, 65433, 65434, 65435, 65436, 65436, 65394, 65396, 65382, 65437, 65395, 65398, 65401);    //半角カタカナ作成用濁点文字No.
    var dakuten_no = new Array(11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 36, 38, 40, 47, 50, 53, 56, 59, 83);
    //半角カタカナ作成用半濁点文字No.
    var handaku_no = new Array(48, 51, 54, 57, 60);

    //半角カタカナ用配列を構成する
    for (i = 0; i < 86; i++) {
        kanahenkan[i] = (hankaku_code[i]);
    }
    for (i = 0; i < 21; i++) {
        kanahenkan[dakuten_no[i]] = kanahenkan[dakuten_no[i]] + dakuten_moji;
    }
    for (i = 0; i < 5; i++) {
        kanahenkan[handaku_no[i]] = kanahenkan[handaku_no[i]] + handaku_moji
    }

    for (i = 0; i < intLen; i++) {
        code = strCheckData.charCodeAt(i);
        if (code >= 12449 && code <= 12534) {
            strRetValue = strRetValue + String.fromCharCode(kanahenkan[code - 12449]);
        } else {
            strRetValue = strRetValue + String.fromCharCode(code);
        }
    }

    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncTextToHankakuAlp
【概要】    : 全角アルファベットを半角アルファべットに変換する
【引数】    : strCheckData        <String>  文字列
【戻り値】  : strRetValue         <String>  変換された文字列
*******************************************************************************/
function fncTextToHankakuAlp(strCheckData) {

    var strRetValue = "";
    var intLen = strCheckData.length;

    for (i = 0; i < intLen; i++) {
        code = strCheckData.charCodeAt(i);

        if (code >= 65313 && code <= 65338) { code = code - 65248; }
        if (code >= 65345 && code <= 65370) { code = code - 65248; }
        strRetValue = strRetValue + String.fromCharCode(code);
    }
    return strRetValue;
}
/*******************************************************************************
【関数名】  : fncTextToHankakuNum
【概要】    : 全角数字を半角数字に変換する
【引数】    : strCheckData        <String>  文字列
【戻り値】  : strRetValue         <String>  変換された文字列
*******************************************************************************/
function fncTextToHankakuNum(strCheckData) {

    var strRetValue = "";
    var intLen = strCheckData.length;

    for (i = 0; i < intLen; i++) {
        code = strCheckData.charCodeAt(i)
        if (code >= 65296 && code <= 65305) { code = code - 65248; }
        strRetValue = strRetValue + String.fromCharCode(code);
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : fncTextToHankakuKigou
【概要】    : 全角記号を半角記号に変換する
【引数】    : strCheckData        <String>  文字列
【戻り値】  : strRetValue         <String>  変換された文字列
*******************************************************************************/
function fncTextToHankakuKigou(strCheckData) {

    var strHan = "!\"#$%&'()=-~^|\\{}｢｣'@*:_;+?･><｡､";
    var strZen = "！”＃＄％＆’（）＝－～＾｜￥｛｝「」‘＠＊：＿；＋？・＞＜。、";
    var strRetValue = "";
    var intLen = strCheckData.length;

    for (i = 0; i < intLen; i++) {
        c = strCheckData.charAt(i);
        n = strZen.indexOf(c, 0);
        if (n >= 0) c = strHan.charAt(n);
        strRetValue += c;
    }
    return strRetValue;
}

/*******************************************************************************
【関数名】  : LogOffConfirm
【概要】    : 確認メッセージ
*******************************************************************************/
function LogOffConfirm(strMessage) {
    if (confirm(strMessage)) {
        return true;
    } else {
        return false;
    }
}

/*******************************************************************************
【関数名】  : fncTrim
【概要】    : スペースを削除
*******************************************************************************/
function fncTrim(strInput) {
    if (strInput == ' ') {
        return '';
    } else {
        return strInput.replace(/(^\s+)|(\s+$)/g, "");
    }
}

/*******************************************************************************
【関数名】  : fncTrim
【概要】    : 日付
*******************************************************************************/
Date.prototype.Format = function (fmt) { //author: meizz 
    var o = {
        "M+": this.getMonth() + 1, //月份 
        "d+": this.getDate(), //日 
        "h+": this.getHours(), //小时 
        "m+": this.getMinutes(), //分 
        "s+": this.getSeconds(), //秒 
        "q+": Math.floor((this.getMonth() + 3) / 3), //季度 
        "S": this.getMilliseconds() //毫秒 
    };
    if (/(y+)/.test(fmt)) fmt = fmt.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length));
    for (var k in o)
        if (new RegExp("(" + k + ")").test(fmt)) fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
    return fmt;
}

/*******************************************************************************
【関数名】  : fncFloatMultiplication
【概要】    : 掛け算の精度アップ
【引数】    : fltArg1          <Float>  数1
: fltArg2          <Float>  数2
【戻り値】  : 数値             <Float>
*******************************************************************************/
function fncFloatMultiplication(fltArg1, fltArg2) {
    var str1 = fltArg1.toString();
    var str2 = fltArg2.toString();
    var intCount1 = 0;               //小数の桁数を記録
    var intCount2 = 0;               //小数の桁数を記録
    var intArg1 = 0;
    var intArg2 = 0;
    var fltResult = 0;               //処理結果

    //小数の桁数を取得
    if (str1.indexOf(".") > -1) {
        intCount1 += str1.split(".")[1].length;
    }
    if (str2.indexOf(".") > -1) {
        intCount2 += str2.split(".")[1].length;
    }
    //整数に変換
    intArg1 = fltArg1 * Math.pow(10, intCount1)
    intArg2 = fltArg2 * Math.pow(10, intCount2)

    fltResult = intArg1 * intArg2 / Math.pow(10, intCount1 + intCount2);

    return fltResult;

}

/*******************************************************************************
【関数名】  : fncFloatRound
【概要】    : 小数の場合、四捨五入の精度アップ
【引数】    : floatArg           <Float>  値
: intDecLen          <Float>  小数部何桁まで四捨五入
【戻り値】  : 数値               <Float>
*******************************************************************************/
function fncFloatRound(floatArg, intDecLen) {
    var intCount = 0;
    //小数部の桁数を取得
    floatArg = floatArg + "";
    if (floatArg.indexOf(".") > -1) {
        intCount = floatArg.split(".")[1].length
    }
    //整数に変換して四捨五入する
    if (intCount > intDecLen) {
        floatArg = floatArg * Math.pow(10, intCount) / Math.pow(10, intCount - intDecLen);
    } else {
        floatArg = floatArg * Math.pow(10, intDecLen);
    }

    floatArg = Math.round(floatArg);
    floatArg = floatArg / Math.pow(10, intDecLen);

    return floatArg;
}
