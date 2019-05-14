/*******************************************************************************
【関数名】  : GridViewCellClick
【概要】    : 仕様画面に仕様選択する時(正常)
【引数】    : strName              <String>  ユーザコントロールフルネーム
            　strRow               <String>  行番号
              strCol               <String>  列番号
              strReverse           <String>  LMF0,T0Dの場合表示を逆にする
*******************************************************************************/
function GridViewCellClick(strName, strRow, strCol, strKataStart, strReverse) {
    var grid = document.getElementById(strName + "_GridViewDetail");
    var intUsedNumber;
    var intRow = Number(strRow);
    var intCol = Number(strCol);
    var intDisplayCol;

    if (grid.rows.length > 0) {
        //使用数位置の取得
        if (grid.rows[0].cells[0].innerText !== undefined) {
            if (grid.rows[0].cells[0].innerText != "CX A") {
                intUsedNumber = 0;
            } else {
                intUsedNumber = 2;
            }
        } else {
            if (grid.rows[0].cells[0].textContent != "CX A") {
                intUsedNumber = 0;
            } else {
                intUsedNumber = 2;
            }
        }

        //逆に表示する場合列番号の調整
        if (strReverse == '1') {
            intDisplayCol = grid.rows(0).cells.length - intCol;
        }
        else {
            intDisplayCol = intCol;
        }

        //使用数と仕様のセット(多ブラウザ対応)
        if (grid.rows[intRow - 1].cells[intDisplayCol].innerText !== undefined) {
            if (grid.rows[intRow - 1].cells[intDisplayCol].innerText == "●") {
                var strAll = document.getElementById(strName + "_HidClick").value
                //var strReplace = strRow + "," + strCol + ";"
                var strReplace = strRow + "," + strCol;

                //strAll = strAll.replace(new RegExp(strReplace, "gm"), "");
                strAll = ReplaceString(strAll, strReplace);
                grid.rows[intRow - 1].cells[intDisplayCol].innerText = "";
                //grid.rows[intRow - 1].cells[intDisplayCol].style.background = "#FFFFC0";
                grid.rows[intRow - 1].cells[intUsedNumber].innerText = Number(grid.rows[intRow - 1].cells[intUsedNumber].innerText) - 1;
                document.getElementById(strName + "_HidClick").value = strAll;
            }
            else {
                grid.rows[intRow - 1].cells[intDisplayCol].innerText = "●";
                //grid.rows[intRow - 1].cells[intDisplayCol].style.background = "lightBlue";
                document.getElementById(strName + "_HidClick").value += strRow + "," + strCol + ";";
                grid.rows[intRow - 1].cells[intUsedNumber].innerText = Number(grid.rows[intRow - 1].cells[intUsedNumber].innerText) + 1;
            }
        } else {
            if (grid.rows[intRow - 1].cells[intDisplayCol].textContent == "●") {
                var strAll = document.getElementById(strName + "_HidClick").value;
                var strReplace = strRow + "," + strCol;

                //strAll = strAll.replace(new RegExp(strReplace, "gm"), "");
                strAll = ReplaceString(strAll, strReplace);
                grid.rows[intRow - 1].cells[intDisplayCol].textContent = "";
                //grid.rows[intRow - 1].cells[intDisplayCol].style.background = "#FFFFC0";
                grid.rows[intRow - 1].cells[intUsedNumber].textContent = Number(grid.rows[intRow - 1].cells[intUsedNumber].textContent) - 1;
                document.getElementById(strName + "_HidClick").value = strAll;
            }
            else {
                grid.rows[intRow - 1].cells[intDisplayCol].textContent = "●";
                //grid.rows[intRow - 1].cells[intDisplayCol].style.background = "lightBlue";
                document.getElementById(strName + "_HidClick").value += strRow + "," + strCol + ";";
                grid.rows[intRow - 1].cells[intUsedNumber].textContent = Number(grid.rows[intRow - 1].cells[intUsedNumber].textContent) + 1;
            }
        }

        GetKataComb(strKataStart, strName);
    }
}

/*******************************************************************************
【関数名】  : ReplaceString
【概要】    : 入れ替え
*******************************************************************************/
function ReplaceString(strBase, strReplace) {
    var strPositions = strBase.split(";");
    var strResult = "";

    for (i = 0; i < strPositions.length; i++) {
        if (strPositions[i] != "") {
            if (strPositions[i] != strReplace) {
                if (i == strPositions.length - 1) {
                    strResult += strPositions[i];
                } else {
                    strResult += strPositions[i] + ";";
                }
            }
        }
    }
    return strResult;
}

/*******************************************************************************
【関数名】  : GridViewCellClickMid
【概要】    : 仕様画面に仕様選択する時(Mid)
【引数】    : strName              <String>  ユーザコントロールフルネーム
strMidRowID          <String>  Mid行のID
strMidRow            <String>  Mid行の番号
            　strRow               <String>  Mid行の中に選択した行番号
strCol               <String>  Mid行の中に選択した列番号
strReverse           <String>  LMF0,T0Dの場合表示を逆にする
strManifold16_18     <String>  Manifold16の18行目の使用数が2単位で増加する
*******************************************************************************/
function GridViewCellClickMid(strName, strMidRowID, strMidRow, strRow, strCol, strKataStart, strReverse, strManifold16_18) {
    var gridMid = document.getElementById(strMidRowID);
    var grid = document.getElementById(strName + "_GridViewDetail");
    var intUsedNumber;
    var intMidRow = Number(strMidRow);
    var intRow = Number(strRow);
    var intCol = Number(strCol);
    var intDisplayCol;
    var intAddUnit;

    if (grid.rows.length > 0) {
        //使用数位置の取得
        if (grid.rows[0].cells[0].innerText !== undefined) {
            if (grid.rows[0].cells[0].innerText != "CX A") {
                intUsedNumber = 0;
            } else {
                intUsedNumber = 2;
            }
        } else {
            if (grid.rows[0].cells[0].textContent != "CX A") {
                intUsedNumber = 0;
            } else {
                intUsedNumber = 2;
            }
        }

        //逆に表示する場合列番号の調整
        if (strReverse == '1') {
            intDisplayCol = grid.rows(0).cells.length - intCol - 1;
        } else {
            intDisplayCol = intCol;
        }
        //Manifold16の18行目の使用数が2単位で増加する
        if (strManifold16_18 == '1') {
            intAddUnit = 2;
        } else {
            intAddUnit = 1;
        }


        //使用数と仕様のセット
        if (gridMid.rows[intRow].cells[intDisplayCol].innerText !== undefined) {
            if (gridMid.rows[intRow].cells[intDisplayCol].innerText == "●") {
                var strAll = document.getElementById(strName + "_HidClick").value
                var strReplace = "M" + strMidRow + "," + strRow + "," + strCol + ";"

                //逆にする場合
                gridMid.rows[intRow].cells[intDisplayCol].innerText = "";
                //grid.rows[intRow - 1].cells[intDisplayCol].style.background = "#FFFFC0";
                //使用数とHiddenFieldの設定
                strAll = strAll.replace(new RegExp(strReplace, "gm"), "");
                grid.rows[intMidRow - 1 + intRow].cells[intUsedNumber].innerText = Number(grid.rows[intMidRow - 1 + intRow].cells[intUsedNumber].innerText) - intAddUnit;
                document.getElementById(strName + "_HidClick").value = strAll;
            }
            else {
                //HidClickサンプール:正常:strRow1,strColumn1;strRow2,strColumn2
                //                  :Mid:M+strMidRow,strRow1,strColumn1;M+strMidRow,strRow2,strColumn2;
                gridMid.rows[intRow].cells[intDisplayCol].innerText = "●";
                //grid.rows[intRow - 1].cells[intDisplayCol].style.background = "lightBlue";
                //使用数とHiddenFieldの設定
                document.getElementById(strName + "_HidClick").value += "M" + strMidRow + "," + strRow + "," + strCol + ";";
                grid.rows[intMidRow - 1 + intRow].cells[intUsedNumber].innerText = Number(grid.rows[intMidRow - 1 + intRow].cells[intUsedNumber].innerText) + intAddUnit;
            }
        } else {
            if (gridMid.rows[intRow].cells[intDisplayCol].textContent == "●") {
                var strAll = document.getElementById(strName + "_HidClick").value
                var strReplace = "M" + strMidRow + "," + strRow + "," + strCol + ";"

                //逆にする場合
                gridMid.rows[intRow].cells[intDisplayCol].textContent = "";
                //grid.rows[intRow - 1].cells[intDisplayCol].style.background = "#FFFFC0";
                //使用数とHiddenFieldの設定
                strAll = strAll.replace(new RegExp(strReplace, "gm"), "");
                grid.rows[intMidRow - 1 + intRow].cells[intUsedNumber].textContent = Number(grid.rows[intMidRow - 1 + intRow].cells[intUsedNumber].textContent) - intAddUnit;
                document.getElementById(strName + "_HidClick").value = strAll;
            }
            else {
                //HidClickサンプール:正常:strRow1,strColumn1;strRow2,strColumn2
                //                  :Mid:M+strMidRow,strRow1,strColumn1;M+strMidRow,strRow2,strColumn2;
                gridMid.rows[intRow].cells[intDisplayCol].textContent = "●";
                //grid.rows[intRow - 1].cells[intDisplayCol].style.background = "lightBlue";
                //使用数とHiddenFieldの設定
                document.getElementById(strName + "_HidClick").value += "M" + strMidRow + "," + strRow + "," + strCol + ";";
                grid.rows[intMidRow - 1 + intRow].cells[intUsedNumber].textContent = Number(grid.rows[intMidRow - 1 + intRow].cells[intUsedNumber].textContent) + intAddUnit;
            }
        }

        GetKataComb(strKataStart, strName);
    }
}

/*******************************************************************************
【関数名】  : GetKataComb
【概要】    : 選択した形番を取得する
【引数】    : strName              <String>  ユーザコントロールフルネーム
strKata              <String>  
*******************************************************************************/
function GetKataComb(strKata, strName) {
    var KataCmb;
    var KataList = "";
    var intStart = parseInt(strKata.substr(strKata.length - 2, 2));
    for (inti = intStart; inti < intStart + 50; inti = inti + 1) {
        if (inti <= 9) {
            KataCmb = strKata.substr(0, strKata.length - 2) + "0" + inti;
        } else {
            KataCmb = strKata.substr(0, strKata.length - 2) + inti;
        }
        //画面の形番を取得する
        if (document.getElementById(KataCmb + "_cmbkata")) {
            KataList = KataList + document.getElementById(KataCmb + "_cmbkata").value + ",";
        } else if (document.getElementById(KataCmb + "_lblkata")) {

            if (document.getElementById(KataCmb + "_lblkata").innerText !== undefined) {
                KataList = KataList + document.getElementById(KataCmb + "_lblkata").innerText + ",";
            } else {
                KataList = KataList + document.getElementById(KataCmb + "_lblkata").textContent + ",";
            }

        } else if (document.getElementById(KataCmb + "_chk1")) {
            //ADD BY YGY 20141107    チューブ抜具の設定
            if (document.getElementById(KataCmb + "_chk1").checked) {
                KataList = KataList + "1";
            } else {
                KataList = KataList + "0";
            }
        } else if (document.getElementById(KataCmb)) {
            if (document.getElementById(KataCmb).innerText !== undefined) {
                KataList = KataList + document.getElementById(KataCmb).innerText + ",";
            } else {
                KataList = KataList + document.getElementById(KataCmb).textContent + ",";
            }

        }
    }
    document.getElementById(strName + "_HidSelect").value = KataList;
}

/*******************************************************************************
【関数名】  : Siyou_OK
【概要】    : OKボタンイベント
【引数】    : strName              <String>  ユーザコントロールフルネーム
*******************************************************************************/
function Siyou_OK(strName) {
    try {
        var strID = document.getElementById(strName + "_HidStartID").value
        var strRowID = strID.split(",");
        GetKataComb(strRowID[0], strName);   //選択した形番を取得する
        GetKataUse(strRowID[1], strName);    //使用数を取得する
        //SetComputerName(strName);                   //端末名を取得する
    } catch (err) {
        alert(err.LineText)
    }
}

/*******************************************************************************
【関数名】  : GridViewCellSelCX
【概要】    : CXA,CXBの選択情報をHidSetCXに保存
【引数】    : strName              <String>  ユーザコントロールフルネーム
*******************************************************************************/
function GridViewCellSelCX(strClientID, strName, strRowID, strStartID) {
    //選択された形番
    var strSelKata = document.getElementById(strClientID).value;
    var strSolType = document.getElementById(strName + "_HidSetCX").value

    //fncKatabanOptionSet(strSelKata, strClientID, strSolType);
    
    Siyou_OK(strName);
    var timer = document.getElementById(strName + "_btnClick");
    if (timer) { timer.click(); }
}

/*******************************************************************************
【関数名】  : GridViewCellSelBlock
【概要】    : 電装ブロック選択肢の更新
【引数】    : strName              <String>  ユーザコントロールフルネーム
*******************************************************************************/
function GridViewCellSelBlock(strName) {
    //選択された形番
    Siyou_OK(strName);
    var timer = document.getElementById(strName + "_btnClick");
    if (timer) { timer.click(); }
}

/*******************************************************************************
【関数名】  : SetSiyouData
【概要】    : ﾚｰﾙ長さ選択情報設定
【引数】    : 
*******************************************************************************/
function SetSiyouData(strName, strName1, strValue) {
    if (document.getElementById(strName)) {
        document.getElementById(strName).value = strValue;
    }
    if (document.getElementById(strName1)) {
        document.getElementById(strName1).value = strValue;
    }
}

/*******************************************************************************
【関数名】  : SetSiyouData
【概要】    : レール長さが入力されたフラグの設定
【引数】    : 
*******************************************************************************/
function fncTextRailChange(strHidRailChangedFlgID) {
    document.getElementById(strHidRailChangedFlgID).value = "1";
}

/*******************************************************************************
【関数名】  : fncKatabanOptionSet
【概要】    : 形番ドロップダウン値変更時、条件によってフォーム部品を作り変える
【引数】    : strKata            <String>    選択された形番
【引数】    : strClientID        <String>    コントロールID
【戻り値】  : 無し
*******************************************************************************/
function fncKatabanOptionSet(strKata, strClientID, strSolType) {
    //CXAのIDを取得
    var strCXAId = strClientID.replace("cmbkata", "cmbCXA");
    var strCXAId = strCXAId.replace("GridViewTitle", "GridViewDetail");

    //CXBのIDを取得
    var strCXBId = strClientID.replace("cmbkata", "cmbCXB");
    var strCXBId = strCXBId.replace("GridViewTitle", "GridViewDetail");
    var strValues = "";

    //ドロップダウンをクリアする
    fncDelOptions(strCXAId);
    fncDelOptions(strCXBId);

    //選択肢の作成
    if (strKata.indexOf("-CX") >= 0) {
        //電磁弁
        if (strKata.substring(0, 4) == "3GB1" || strKata.substring(0, 4) == "3GE1") {
            if (strSolType == "0") {
                strValues = ",C4,C6,C18,CD4,CD6,X"
            } else {
                strValues = ",C4,C6,X"
            }
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);
        } else if (strKata.substring(0, 4) == "3GB2") {
            if (strSolType == "0") {
                strValues = ",C4,C6,C8,CD6,CD8,X"
            } else {
                strValues = ",C4,C6,C8,X"
            }
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);
        } else if (strKata.substring(0, 4) == "4GB1" || strKata.substring(0, 4) == "4GE1") {
            if (strKata.substring(4, 5) == "1") {
                if (strSolType == "0") {
                    strValues = ",C4,C6,C18,CD4,CD6,CF,CL4,CL6,X"
                } else {
                    strValues = ",C4,C6,CL4,CL6,X"
                }
                fncSetOptions(strCXAId, strValues);
                fncSetOptions(strCXBId, strValues);
            } else {
                if (strSolType == "0") {
                    strValues = ",C4,C6,C18,CD4,CD6,CF,X"
                } else {
                    strValues = ",C4,C6,X"
                }
                fncSetOptions(strCXAId, strValues);
                fncSetOptions(strCXBId, strValues);
            }
        } else if (strKata.substring(0, 4) == "4GB2") {
            if (strKata.substring(4, 5) == "1") {
                if (strSolType == "0") {
                    strValues = ",C4,C6,C8,CD6,CD8,CL6,CL8,X"
                } else {
                    strValues = ",C4,C6,C8,CL6,CL8,X"
                }
                fncSetOptions(strCXAId, strValues);
                fncSetOptions(strCXBId, strValues);
            } else {
                if (strSolType == "0") {
                    strValues = ",C4,C6,C8,CD6,CD8,X"
                } else {
                    strValues = ",C4,C6,C8,X"
                }
                fncSetOptions(strCXAId, strValues);
                fncSetOptions(strCXBId, strValues);
            }
        } else if (strKata.substring(0, 4) == "4GB3") {
            if (strKata.substring(4, 5) == "1") {
                if (strSolType == "0") {
                    strValues = ",C6,C8,C10,CD8,CD10,CL8,CL10,X"
                } else {
                    strValues = ",C6,C8,C10,CL8,CL10,X"
                }
                fncSetOptions(strCXAId, strValues);
                fncSetOptions(strCXBId, strValues);
            } else {
                if (strSolType == "0") {
                    strValues = ",C6,C8,C10,CD8,CD10,X"
                } else {
                    strValues = ",C6,C8,C10,X"
                }
                fncSetOptions(strCXAId, strValues);
                fncSetOptions(strCXBId, strValues);
            }
        } else if (strKata.substring(0, 4) == "4GB4" || strKata.substring(0, 4) == "4GE4") {
            strValues = ",C8,C10,C12"
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);

            //RM0904031 2009/06/24 Y.Miura↓↓
        } else if (strKata.substring(0, 4) == "3GE2") {
            if (strSolType == "0") {
                strValues = ",C4,C6,C8,X"
            } else {
                strValues = ",C4,C6,C8,X"
            }
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);
        } else if (strKata.substring(0, 4) == "4GE2") {
            if (strKata.substring(4, 5) == "1") {
                strValues = ",C4,C6,C8,CL6,CL8,X"
                fncSetOptions(strCXAId, strValues);
                fncSetOptions(strCXBId, strValues);
            } else {
                strValues = ",C4,C6,C8,X"
                fncSetOptions(strCXAId, strValues);
                fncSetOptions(strCXBId, strValues);
            }
        } else if (strKata.substring(0, 4) == "4GE3") {
            if (strKata.substring(4, 5) == "1") {
                strValues = ",C6,C8,C10,CL8,CL10,X"
                fncSetOptions(strCXAId, strValues);
                fncSetOptions(strCXBId, strValues);
            } else {
                strValues = ",C6,C8,C10,X"
                fncSetOptions(strCXAId, strValues);
                fncSetOptions(strCXBId, strValues);
            }
            //RM0904031 2009/06/24 Y.Miura↑↑

        } else if (strKata.indexOf("4G1-MP-") >= 0) {
            //マスキングプレート(個別配線)
            if (strSolType == "0") {
                strValues = ",C4,C6,C18,CD4,CD6,CF,CL4,CL6,X"
            } else {
                strValues = ",C4,C6,CL4,CL6,X"
            }
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);
        } else if (strKata.indexOf("4G2-MP-") >= 0) {
            if (strSolType == "0") {
                strValues = ",C4,C6,C8,CD6,CD8,CL6,CL8,X"
            } else {
                strValues = ",C4,C6,C8,CL6,CL8,X"
            }
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);
        } else if (strKata.indexOf("4G3-MP-") >= 0) {
            if (strSolType == "0") {
                strValues = ",C6,C8,C10,CD8,CD10,CL8,CL10,X"
            } else {
                strValues = ",C6,C8,C10,CL8,CL10,X"
            }
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);
        } else if (strKata.indexOf("4G1-MPS-") >= 0) {
            //マスキングプレート(省配線)
            if (strSolType == "0") {
                strValues = ",C4,C6,C18,CD4,CD6,CF,CL4,CL6,X"
            } else {
                strValues = ",C4,C6,CL4,CL6,X"
            }
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);
        } else if (strKata.indexOf("4G1-MPD-") >= 0) {
            if (strSolType == "0") {
                strValues = ",C4,C6,C18,CD4,CD6,CF,X"
            } else {
                strValues = ",C4,C6,X"
            }
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);
        } else if (strKata.indexOf("4G2-MPS-") >= 0) {
            if (strSolType == "0") {
                strValues = ",C4,C6,C8,CD6,CD8,CL6,CL8,X"
            } else {
                strValues = ",C4,C6,C8,CL6,CL8,X"
            }
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);
        } else if (strKata.indexOf("4G2-MPD-") >= 0) {
            if (strSolType == "0") {
                strValues = ",C4,C6,C8,CD6,CD8,X"
            } else {
                strValues = ",C4,C6,C8,X"
            }
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);
        } else if (strKata.indexOf("4G3-MPS-") >= 0) {
            if (strSolType == "0") {
                strValues = ",C6,C8,C10,CD8,CD10,CL8,CL10,X"
            } else {
                strValues = ",C6,C8,C10,CL8,CL10,X"
            }
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);
        } else if (strKata.indexOf("4G3-MPD-") >= 0) {
            if (strSolType == "0") {
                strValues = ",C6,C8,C10,CD8,CD10,X"
            } else {
                strValues = ",C6,C8,C10,X"
            }
            fncSetOptions(strCXAId, strValues);
            fncSetOptions(strCXBId, strValues);
        }
    }
}

/*******************************************************************************
【関数名】  : fncDelOptions
【概要】    : ドロップダウンのオプションをクリアする
【引数】    : objTarget      <Object>    処理対象ドロップダウンオブジェクト
【戻り値】  : 無し
*******************************************************************************/
function fncDelOptions(strID) {
    var intLen = document.getElementById(strID).options.length;

    for (intI = intLen - 1; intI >= 0; intI--) {
        document.getElementById(strID).options[intI] = null;
    }
    document.getElementById(strID).options[0] = new Option("", "");
}

/*******************************************************************************
【関数名】  : fncSetOptions
【概要】    : ドロップダウンにオプションをセットする
【引数】    : objTarget1     <Object>    処理対象ドロップダウンオブジェクト1
: objTarget2     <Object>    処理対象ドロップダウンオブジェクト2
: strValues      <String>    カンマ区切りで文字列連結したValue
: intStIdx       <int>       オプション配列インデックスのスタート値
【戻り値】  : 無し
*******************************************************************************/
function fncSetOptions(strID, strValues) {
    var strValue = new Array();

    strValue = strValues.split(",");
    for (intI = 0; intI < strValue.length; intI++) {
        document.getElementById(strID).options[intI] = new Option(strValue[intI], strValue[intI]);
    }
}

/*******************************************************************************
【関数名】  : SetComputerName
【概要】    : 端末名を取得
【戻り値】  : 無し
*******************************************************************************/
function SetComputerName(strName) {
    var computer;
    var locator = new ActiveXObject('WScript.Network');

    computer = loca.computerName;
    document.getElementById(strName + "_HidComputerName").value = computer;
}

/*******************************************************************************
【関数名】  : UpdateRail
【概要】    : 画面の更新
【戻り値】  : 無し
*******************************************************************************/
function UpdateRail(strName) {

    Siyou_OK(strName);
    var timer = document.getElementById(strName + "_btnClick");
    if (timer) { timer.click(); }
}