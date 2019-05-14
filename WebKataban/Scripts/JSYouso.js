/*******************************************************************************
【関数名】  : YousoDblClick
【概要】    : 要素ダブルクリックイベント
【引数】    : strName         <String>    画面名
            　strID           <String>    画面ID
            　rowIndex        <Integer>   選択した行No
【戻り値】  : 無し
*******************************************************************************/
function YousoDblClick(strName, strID, strStartID, rowIndex) {
    var intID;
    //全て選択された形番
    var strValue = '';

    try {
        if (document.getElementById(strID)) {
            intID = document.getElementById(strName + "HidCurrentFocus").value;
            var strMultiplcation = document.getElementById(strName + "HidMultiplcation").value.split(",");
            
            //選択した形番
            if (document.getElementById(strID).childNodes[0].innerText !== undefined) {
                strValue = fncTrim(document.getElementById(strID).childNodes[0].innerText);
            } else {
                strValue = fncTrim(document.getElementById(strID).childNodes[1].textContent);
            }

            //複数選択できる場合
            if (parseInt(strMultiplcation[intID - 1], 10) >= 4) {
                //画面オプション設定            
                var strAllValue = fncTrim(document.getElementById(strName + "txt" + intID).value + strValue);
                var intOptionsNumber = parseInt(document.getElementById(strName + "HidOptionNumber").value, 10) - 1;

                //画面オプションの設定  
                document.getElementById(strName + "txt" + intID).value = strAllValue

                //選択終了を選択した場合は次のオプションへ
                if (strValue == '') {
                    var intNext = parseInt(intID, 10) + 1;

                    if (intNext > intOptionsNumber) {
                        //最後のオプションの場合
                        document.getElementById(strName + "btnOK").focus();
                        document.getElementById(strName + "labelTitle").style.visibility = 'hidden';
                        document.getElementById(strName + "GVDetail").style.visibility = 'hidden';
                        document.getElementById(strName + "HidCurrentFocus").value = '';
                    } else {
                        //オプションがある場合
                        document.getElementById(strName + "HidSelRowID").value = '';
                        YousoGotFocus(strName, "txt" + intNext);
                    }
                } else {
                    //選択情報をHiddenFieldに反映
                    fncSetSelectedOptions(strName, strValue);

                    //複数選択オプションの設定
                    var intOptionlistNum = fncSetMultiOptions(strName, intID, strValue, rowIndex, strID, strStartID);

                    if (intOptionlistNum == 0) {
                        //オプションがない場合は次のオプションへ
                        var intNext = parseInt(intID, 10) + 1;

                        if (intNext > intOptionsNumber) {
                            //最後のオプションの場合
                            document.getElementById(strName + "btnOK").focus();
                            document.getElementById(strName + "labelTitle").style.visibility = 'hidden';
                            document.getElementById(strName + "GVDetail").style.visibility = 'hidden';
                            document.getElementById(strName + "HidCurrentFocus").value = '';
                        } else {
                            //オプションがある場合
                            document.getElementById(strName + "HidSelRowID").value = '';
                            YousoGotFocus(strName, "txt" + intNext);
                        }
                    }
                }
            } else {
                var strAllValue = strValue;
                document.getElementById(strName + "txt" + intID).value = strAllValue;
                
                //選択情報をHiddenFieldに反映
                fncSetSelectedOptions(strName);
                
                //再ロード                
                var timer = document.getElementById(strName + "btnNext");
                if (timer) { timer.click(); }
            }
        }

    } catch (err) {
        alert(err.LineText)
    }
}

/*******************************************************************************
【関数名】  : fncSetMultiOptions
【概要】    : 複数選択オプションの設定
【引数】    : strName                 <String>    画面名
            　intID                   <Integer>   FocusOnのオプションID
              strValue                <String>    選択した形番
              rowIndex                <String>    選択した行番号
【戻り値】  : 無し
*******************************************************************************/
function fncSetMultiOptions(strName, intID, strValue, rowIndex, strID, strStartID) {
    //複数選択可能な場合
    //削除行番号
    var deleteRows = new Array();

    try {
        var result = 1;
        //全てのオプションを更新する
        var hidAllMultiOptions = document.getElementById(strName + "HidAllMultiOptions").value;

        //より前のオプションを削除
        var gridViewDetail = document.getElementById(strName + "GVDetail");

        for (k = 1; k <= gridViewDetail.rows.length -1; k++) {
            var intSelRowNum = parseInt(strID.slice(-2), 10);
            var intRowNum = parseInt(gridViewDetail.rows[k].id.slice(-2), 10);

            if (intSelRowNum >= intRowNum) {
                if (fncIndexOf(deleteRows,gridViewDetail.rows[k].id) == -1) {
                    deleteRows.push(gridViewDetail.rows[k].id);
                }
            }
        }

        //同じグループの削除
        var groupOptions = hidAllMultiOptions.split(";");
        var deleteKataban = new Array();

        for (m = 0; m < groupOptions.length; m++) {
            var arrayOptions = groupOptions[m].split(",");
            var sameGroupFlg = 0;

            //同じグループの判断
            for (n = 0; n < arrayOptions.length; n++) {
                if (arrayOptions[n] == strValue) {
                    sameGroupFlg = 1;
                    break;
                }
            }
            //同じグループの保存
            if (sameGroupFlg == 1) {
                for (t = 0; t < arrayOptions.length; t++) {
                    if (fncIndexOf(deleteKataban, arrayOptions[t]) == -1) {
                        deleteKataban.push(arrayOptions[t]);
                    }
                }
            }
        }

        //形番により行番号の取得
        var rowNum = gridViewDetail.rows.length;
        for (p = 0; p < deleteKataban.length; p++) {
            for (q = 0; q < rowNum; q++) {
                var rowKataban;

                if (gridViewDetail.rows[q].childNodes[0].innerText !== undefined) {
                    rowKataban = gridViewDetail.rows[q].childNodes[0].innerText;
                } else {
                    rowKataban = gridViewDetail.rows[q].childNodes[1].textContent;
                }
                if (deleteKataban[p] == rowKataban) {
                    if (fncIndexOf(deleteRows, gridViewDetail.rows[q].id) == -1) {
                        deleteRows.push(gridViewDetail.rows[q].id);
                    }
                }
            }
        }

        //不要行の削除
        for (i = 0; i < deleteRows.length; i++) {
            var row = document.getElementById(deleteRows[i]);
            row.parentNode.removeChild(row);
        }

        if (gridViewDetail.rows.length == 0) {
            result = 0;
        } else if (gridViewDetail.rows.length == 1) {
            var rowKataban;
            if (gridViewDetail.rows[0].childNodes[0].innerText != undefined) {
                rowKataban = gridViewDetail.rows[0].childNodes[0].innerText;
            } else {
                rowKataban = gridViewDetail.rows[0].childNodes[1].textContent;
            }
            if (rowKataban.replace(/^\s+|\s+$/g, '') == '') {
                result = 0;
            }
        }

        //フォカスを設定
        if (result != 0) {
            YousoGridClick(strName, gridViewDetail.rows[0].id, strStartID, rowIndex);
        }
        
        return result;
    } catch (err) {

    }
}

/*******************************************************************************
【関数名】  : fncIndexOf
【戻り値】  : 無し
*******************************************************************************/
function fncIndexOf(deleteRows, rowID) {
    var result = -1;

    for (i = 0; i < deleteRows.length; i++) {
        if (rowID == deleteRows[i]) {
            result = 0;
        }
    }
    return result;
}

/*******************************************************************************
【関数名】  : fncDeleteRow
【概要】    : 指定した行を削除
【引数】    : rowIndex                 <String>    削除行番号

【戻り値】  : 無し
*******************************************************************************/
function fncDeleteRow(rowIndex) {
    var table = document.getElementById('<%=GVDetail.ClientID %>');
    table.deleteRow(rowIndex);
}

/*******************************************************************************
【関数名】  : YousoGridClick
【概要】    : 行クリック
【引数】    : rowIndex                 <String>    削除行番号

【戻り値】  : 無し
*******************************************************************************/
function YousoGridClick(strName, strRowClientID, intStartID, rowIndex) {
    var gridViewDetail = document.getElementById(strName + "GVDetail");
    var intRow;
    //全ての行をリセット
    for (m = 0; m < gridViewDetail.rows.length; m++) {
        var rowID = gridViewDetail.rows[m].id;

        if (rowID !== "") {
            var intCount = parseInt(m, 10) - parseInt(intStartID, 10) + 1;
            if (intCount % 2 == 0) {
                strBKcolor = "#CCCCFF";
            } else {
                strBKcolor = "white";
            }

            document.getElementById(rowID).style.backgroundColor = strBKcolor;
            document.getElementById(rowID).style.color = "black";

            if (strRowClientID == rowID) {
                intRow = parseInt(m, 10) + parseInt(intStartID, 10);
            }
        }
    }
    if (document.getElementById(strRowClientID)) {
        document.getElementById(strRowClientID).style.backgroundColor = "#003C80";
        document.getElementById(strRowClientID).style.color = "white";
        document.getElementById(strName + "HidSelRowID").value = intRow;
        document.getElementById(strRowClientID).focus();
    }

    //矢印イメージをクリックした場合
    var hidArrowClick = document.getElementById(strName + "HidArrowClick").value;
    if (hidArrowClick == "1") {
        //矢印のリセット
        document.getElementById(strName + "HidArrowClick").value = "0";
        
        //ダブルクリック
        YousoDblClick(strName, strRowClientID, intStartID, rowIndex);
    }
}

/*******************************************************************************
【関数名】  : YousoKeyDown
【概要】    : オプションキーダウンイベント
【引数】    : strName                 <String>    画面名
            　strFirstRowID           <String>    GridViewの第一行ID
              strOptionNo             <String>    オプション番号
【戻り値】  : 無し
*******************************************************************************/
function YousoKeyDown(e, strName, strFirstRowID, strOptionNo) {
    switch (e.keyCode) {
        case 13:
            var intOptionNumber = parseInt(document.getElementById(strName + "HidOptionNumber").value, 10) - 1;
            var intNextOptionNo = parseInt(strOptionNo, 10) + 1;

            if (intNextOptionNo + 1 > intOptionNumber) {
                document.getElementById(strName + "btnOK").focus();
            } else {
                document.getElementById(strName + "txt" + intNextOptionNo).focus();
            }
            event.Returnvalue = false;
            return false;
        case 40:
            // 下矢印押下時
            if (document.getElementById(strName + "GVDetail")) {
                //リスト第一行を選択
                document.getElementById(strName + "GVDetail_ctl02").click();
            }
            event.Returnvalue = false;
            return false;
        default:
            break;
    }
}

/*******************************************************************************
【関数名】  : YousoGotFocus
【概要】    : フォカスゲットイベント
【引数】    : strName         <String>    画面名
            　strID           <String>    オプションID
【戻り値】  : 無し
*******************************************************************************/
function YousoGotFocus(strName, strID) {
    //20141014 YGY FIREFOX対応
    if ($('#' + strName + 'PnlText').attr("disabled") != true) {
        var intID;

        intID = parseInt(strID.replace("txt", ""), 10);

        if ((document.getElementById(strName + "HidCurrentFocus").value != intID)) {

            //20141014 YGY FIREFOX対応
            //document.getElementById(strName + "PnlText").disabled = true;

            //手入力を無効する
            $('#' + strName + 'PnlText').attr("disabled", true);

            //document.getElementById(strName + strID).value = "";

            document.getElementById(strName + "HidCurrentFocus").value = intID;
            document.getElementById(strName + "HidSelRowID").value = "";
            document.getElementById(strName + "HidSelectedMultiOptions").value = "";

            fncSetOptionColor(strName, intID);
            fncSetSelectedOptions(strName);
            
            //再ロード
            var timer = document.getElementById(strName + "btnNext");
            if (timer) { timer.click(); }
        }
        //} 
    }
}

/*******************************************************************************
【関数名】  : fncSetSelectedOptions
【概要】    : OKボタンイベント
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function fncSetSelectedOptions(strName, strValue) {
    //選択したオプション
    var strKataban = "";

    for (inti = 1; inti <= 35; inti++) {
        if (document.getElementById(strName + "txt" + inti)) {
            if (inti > 1) { strKataban = strKataban + "|"; }
            strKataban = strKataban + document.getElementById(strName + "txt" + inti).value;
        }
    }
    document.getElementById(strName + "HidSelectedOptions").value = strKataban;

    //複数選択したオプション
    //選択したオプションをHiddenFieldに設定
    if (typeof strValue !== "undefined") {
        if (document.getElementById(strName + "HidSelectedMultiOptions").value == "") {
            document.getElementById(strName + "HidSelectedMultiOptions").value = strValue;
        } else {
            document.getElementById(strName + "HidSelectedMultiOptions").value += "," + strValue;
        }
    }
}

/*******************************************************************************
【関数名】  : fncSetBackFocus
【概要】    : 単価画面戻る時のフォカスを設定
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function fncSetBackFocus(strName) {
    document.getElementById(strName + "HidCurrentFocus").value = '1';
    document.getElementById(strName + "HidSelRowID").value = '';
    document.getElementById(strName + "HidOKClick").value = '1';
}

/*******************************************************************************
【関数名】  : fncSetOptionColor
【概要】    : オプション色を設定
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function fncSetOptionColor(strName, intSelectOption) {
    var intOptionsNumber = parseInt(document.getElementById(strName + "HidOptionNumber").value, 10) - 1;

    for (s = 0; s < intOptionsNumber; s++) {

        if (document.getElementById(strName + "txt" + s)) {
            if (s != intSelectOption) {
                document.getElementById(strName + "txt" + s).style.backgroundColor = "#FFFFC0";
            } else {
                document.getElementById(strName + "txt" + s).style.backgroundColor = "#FFFFCC";
            }
        }
    }
}

/*******************************************************************************
【関数名】  : ArrowClick
【概要】    : 矢印をクリック
【引数】    : strName         <String>    画面名
【戻り値】  : 無し
*******************************************************************************/
function ArrowClick(strName, strID, strStartID, rowIndex) {
    document.getElementById(strName + "HidArrowClick").value = "1";
}

