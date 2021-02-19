const e = require("express");
const mysql2 = require('mysql2/promise');
const { isDate } = require('moment');
const fs = require('fs');

//今日日付取得用
require('date-utils');

var coGt = require('./Common_Get.js');      //Common_Get.js読み込み

//画面入力必須項目チェック
exports.requiredCheck = function(path) {
    var inputError = "";
    try{

    //読み込み先パスの入力が無かったらエラー
        if(path.length == 0){
            inputError = "ファイル出力先パスを入力して下さい";
            console.log(inputError);
        }
    }catch(e){
        console.log(e.stack);
        console.log('画面必須項目チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}

//ファイルの存在チェック
exports.fileExistenceCheck = function(path) {
    var inputError = "";
    try{
        // ファイルが存在しない時→エラー
        if(fs.existsSync(path)) {
            //ファイル存在有り
        }else {
            inputError = "指定のファイルが見つかりませんでした";
            console.log(inputError);
        }

    }catch(e){
        console.log(e.stack);
        console.log('ファイルの存在チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}

exports.fileFormatCheck = function(path){
    var inputError = '';
    try{
    //末尾5文字の取得
        var format = path.slice( -5 );

    //末尾5文字が.xlsxでない時➡エラー
        if(format != '.xlsm'){
            inputError = 'ファイル形式が間違っています。取り込みファイルはxlsmファイルを指定してください';
            console.log(inputError);
        }

    }catch(e){
        console.log(e.stack);
        console.log('ファイルの形式チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}

//選択年月チェック
exports.selectionCheck = function(date) {
    var inputError = "";
    try{
        var today = new Date();
        today.setMonth(today.getMonth() + 1);
        var twodate = today.toFormat("YYYY/MM");
        // var twodate = '2100/12';

    //選択年月が取り込みボタン押下日＋一か月より先の年月の時エラー
        if(date > twodate){
            inputError = "2か月以上先のファイルは取り込むことができません";
            console.log(inputError);
        }

    }catch(e){
        console.log(e.stack);
        console.log('選択年月で予期せぬエラー発生');
        throw e;
    }
    return inputError;
}

//シート項目数チェック
exports.itemsCountCheck = async function(sheetName,summary,terminalColumn) {
    var count = '';
    var inputError = '';
    try{
        
        //シート名が営業名シート以外の時➡末端列P
            if(summary == 1 || summary == 0){
                count = 'P';

        //シート名が営業名シート
            }else if(summary == 2){
                count = 'S';

            }

//シートの必須項目数とシート毎の末端列の値が違う時➡エラー

            console.log("count:"+count);
            console.log("terminalColumn:"+terminalColumn);

            if(count != terminalColumn){
                inputError = sheetName + '-シートの記載形式が正しくありません';
                console.log(inputError);
            }
        
    }catch(e){
        console.log(e.stack);
        console.log('シート項目数で予期せぬエラー発生');
        throw e;
    }
    return inputError;
}

//Excel-必須項目チェック
exports.requiredItemCheck = async function(sheetName,item,value,x) {
    var inputError = '';
    try{
//入力必須項目があるか判定
        if(value == null){
            inputError = sheetName + '-' + item + '-' + x + '行目-入力必須項目です';
            console.log(inputError);
        }

    }catch(e){
        console.log(e.stack);
        console.log('Excel-必須項目チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}

//Excel-入力文字数チェック//(シート名,項目名,項目のデータ,チェックする文字数,行数)
exports.lengthCheck = async function(sheetName,item,wordValue,wordLength,x) {
    var inputError = '';
    try{

    //チェックする値がNaNの時⇒文字数チェックは飛ばす
        if(Number.isNaN(wordValue)){
        }else{
//入力があったらエラーチェックをする
    //指定の文字数より文字数が大きい時エラー
            if(String(wordValue).length > wordLength){
                inputError = sheetName + '-' + item + '-' + x + '行目-' + wordLength + '文字以下で入力してください';
                console.log(inputError);
            }
        }
    }catch(e){
        console.log(e.stack);
        console.log('Excel-入力文字数チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}


//Excel-データ数チェック//(シート名,項目名,項目のデータ,チェックするデータ数,,行数)
exports.dataCountCheck = async function(sheetName,item,wordValue,wordLength,x) {
    var inputError = '';
    try{
//値が指定のデータ数より大きかったらエラーBuffer.byteLength('buffer')
        if(String(wordValue).bytes() > wordLength){
            inputError = sheetName + '-' + item + '-' + x + '行目-半角文字:1byte 全角文字:2byte 合計' + wordLength + 'byte以内で入力してください';
            console.log(inputError);
        }
    }catch(e){
        console.log(e.stack);
        console.log('Excel-Byte数チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}

//データ数取得
String.prototype.bytes = function () {
    var length = 0;
    for (var i = 0; i < this.length; i++) {
      var c = this.charCodeAt(i);
      if ((c >= 0x0 && c < 0x81) || (c === 0xf8f0) || (c >= 0xff61 && c < 0xffa0) || (c >= 0xf8f1 && c < 0xf8f4)) {
        length += 1;
      } else {
        length += 2;
      }
    }
    return length;
  };

//Excel-全角文字チェック//(シート名,項目名,項目のデータ,行数)
exports.fullWidthCheck = async function(sheetName,item,wordValue,x) {
    var inputError = '';
    try{
//値が全角文字以外の時⇒エラー
        if(String(wordValue).match("^[^!-~｡-ﾟ]*$")){
            //正常
        }else{
            inputError = sheetName + '-' + item + '-' + x + '行目-全角文字で入力してください';
            console.log(inputError);
        }
    }catch(e){
        console.log(e.stack);
        console.log('Excel-全角文字チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}

//Excel-半角文字チェック//(シート名,項目名,項目のデータ,行数)
exports.halfWidthCheck = async function(sheetName,item,wordValue,x) {
    var inputError = '';
    try{
//値が少数以外の時⇒エラー
        if(String(wordValue).match(/^([1-9]\d*|0)(\.\d+)?$/)){
        // if(String(wordValue).match(/^[-]?([1-9]\d*|0)(\.\d+)?$/)){
            //正常
        }else{
            inputError = sheetName + '-' + item + '-' + x + '行目-半角小数または半角数字で入力してください';
            console.log(inputError);
        }
    }catch(e){
        console.log(e.stack);
        console.log('Excel-半角文字チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}


//Excel-半角文字(マイナス考慮)チェック//(シート名,項目名,項目のデータ,行数)////////////////////////////////////////////////////////////////////////////////////////(12-22追加項目)
exports.halfWidthMinusCheck = async function(sheetName,item,wordValue,x) {
    var inputError = '';
    try{
//値が少数以外の時⇒エラー
        if(String(wordValue).match(/^[-]?([1-9]\d*|0)(\.\d+)?$/)){
            //正常
        }else{
            inputError = sheetName + '-' + item + '-' + x + '行目-半角小数または半角数字で入力してください';
            console.log(inputError);
        }
    }catch(e){
        console.log(e.stack);
        console.log('Excel-半角文字チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}



//Excel-半角数字チェック//(シート名,項目名,項目のデータ,行数)
exports.halfWidthNumberCheck = async function(sheetName,item,wordValue,x) {
    var inputError = '';
    try{
//値が半角数字以外の時⇒エラー
        if(String(wordValue).match("^[0-9]+$")){
            //正常
        }else{
            inputError = sheetName + '-' + item + '-' + x + '行目-半角数字で入力してください';
            console.log(inputError);
        }
    }catch(e){
        console.log(e.stack);
        console.log('Excel-半角数字チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}

//Excel-日付妥当性チェック//(シート名,項目名,項目のデータ,行数)
exports.dateValidityCheck = async function(sheetName,item,wordValue,x) {
    var inputError = '';
    try{
//値が日付でない時⇒エラー


       //日付妥当性チェック
        var y = wordValue.split("/")[0];
        var m = wordValue.split("/")[1] - 1;
        var d = wordValue.split("/")[2];
        var date = new Date(y, m, d);

        if (date.getFullYear() != y || date.getMonth() != m || date.getDate() != d) {
            console.log("日付不正と判断");
            inputError = sheetName + '-' + item + '-' + x + '行目-日付入力形式が違います';
            console.log(inputError); 
        }

        // if(isDate(wordValue)){
        // }else{
        //     inputError = sheetName + '-' + item + '-' + x + '行目-日付入力形式が違います';
        //     console.log(inputError); 
        // }

    }catch(e){
        console.log(e.stack);
        console.log('Excel-日付妥当性チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}

//Excel-稼働終了日チェック//(シート名,項目名,項目のデータ,行数)
exports.operatingEndDateCheck = async function(sheetName,item,wordValue,x) {
    var inputError = '';
    try{
//稼働終了日が過去年月日の時⇒エラー
        var today = new Date();
        today = today.toFormat("YYYY/MM/DD");

    //日付で入力されている稼働終了日を判定
        // if(isDate(wordValue)){
            var y = wordValue.split("/")[0];
            var m = wordValue.split("/")[1] - 1;
            var d = wordValue.split("/")[2];
            var date = new Date(y, m, d);
    
            if (date.getFullYear() != y || date.getMonth() != m || date.getDate() != d) {
                console.log("日付不正と判断");
                inputError = sheetName + '-' + item + '-' + x + '行目-日付入力形式が違います';
                console.log(inputError); 
            
                //稼働終了日が取り込みボタン押下日より前の年月日の時エラー
                    // if(day < today){
                    

            }else{
                    if(date > today){
                        inputError = sheetName + '-' + item + '-' + x + '行目-当月分のファイルの稼働終了日に過去日のものがあります';
                        console.log(inputError);
                    }
            }
    }catch(e){
        console.log(e.stack);
        console.log('Excel-日付妥当性チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}

//Excel-日付大小チェック
exports.dateRangeCheck = async function(sheetName,wordValue1,wordValue2,x) {
    var inputError = '';
    try{
//日付範囲が正常でない時⇒エラー
        var y = wordValue1.split("/")[0];
        var m = wordValue1.split("/")[1] - 1;
        var d = wordValue1.split("/")[2];
        var date1 = new Date(y, m, d);

        y = wordValue2.split("/")[0];
        m = wordValue2.split("/")[1] - 1;
        d = wordValue2.split("/")[2];
        var date2 = new Date(y, m, d);



        // if(wordValue1.toFormat("YYYY/MM/DD") > wordValue2.toFormat("YYYY/MM/DD")){
        if(date1.toFormat("YYYY/MM/DD") > date2.toFormat("YYYY/MM/DD")){
            inputError = sheetName + '-' + x + '行目-稼働終了日が稼働開始日より前の日付で入力されています';
        }
    }catch(e){
        console.log(e.stack);
        console.log('Excel-日付妥当性チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}


//Excel-集計情報(営業別受注件数)チェック//(シート名,項目名,項目のデータ,行数)
exports.salesOrderCountCheck = async function(sheetName,wordValue) {
    var inputError = '';
    try{
//営業名の登録情報がない時⇒エラー
        if(wordValue != null){
            var name = wordValue.v;

//月が入っているか
            var nameMonth = String(name).indexOf( '月' );
            
    //月が無かったらエラー発生
            if(nameMonth == -1) {
                inputError = sheetName + '-集計情報-' + name + '-「〇月(営業名)受注件数」と入力してください';
                console.log(inputError);
            }else {
                var nameSp = String(name).split('月');  //月区切り
    //月区切りで文字が取得できなければエラー
                if(nameSp[1] == ''){
                    inputError = sheetName + '-集計情報-' + name + '-「〇月(営業名)受注件数」と入力してください';
                }else{


                    var nameRe = nameSp[1].replace('受注件数','');  //受注件数を空文字へ
                    
                //大倉の時は存在チェックしない
                    if(nameRe == '大倉'){
                        
                    }else{
                //営業別受注件数から取り出した営業名での登録があるか検索
                        var nameOrderCount = nameRe + '一覧';
                        var summary = await coGt.getSummary(nameOrderCount);
                        if(summary == 100){
                            inputError = sheetName + '-集計情報-' + name + '-' + '営業名の登録が無いので、登録ができません';
                            console.log(inputError);
                        }
                    }
                }
            }
        }
    }catch(e){
        console.log(e.stack);
        console.log('Excel-営業別受注件数の項目チェックで予期せぬエラー発生');
        throw e;
    }
    return inputError;
}


//Excel-データ重複チェック
exports.duplicationCheck = async function(sheetName,sheetInfo,dataCount) {
    var duplicationError = '';
    var errorResult = 0;
    try{
        for(var x = 5;x < (dataCount+5);x++){
            var errorMsg = sheetName + '-' + x + '行目';
            if(sheetInfo['F' + x] != null && sheetInfo['G' + x] != null && sheetInfo['H' + x] != null && sheetInfo['I' + x] != null && sheetInfo['J' + x] != null){
                for(var y = 5;y < (dataCount+5);y++){
                    if(x != y){     //同じ行情報は見ない
                        if(sheetInfo['F' + y] != null && sheetInfo['G' + y] != null && sheetInfo['H' + y] != null && sheetInfo['I' + y] != null && sheetInfo['J' + y] != null){
            //案件元・案件元営業・所属会社・所属会社営業・技術者名が他の行のデータと重複していたらエラー
                            if(sheetInfo['F' + x].v == sheetInfo['F' + y].v && sheetInfo['G' + x].v == sheetInfo['G' + y].v && sheetInfo['H' + x].v == sheetInfo['H' + y].v && sheetInfo['I' + x].v == sheetInfo['I' + y].v && sheetInfo['J' + x].v == sheetInfo['J' + y].v){
                                errorMsg = errorMsg + 'と' + y + '行目';
                                errorResult = 1;
                            }
                        }
                    }
                }
                if(errorResult == 1){
                    duplicationError = errorMsg + 'のデータが重複しています';
                    break;
                }
            }
        }
    }catch(e){
        console.log(e.stack);
        console.log('データの重複チェックで予期せぬエラー発生');
        throw e;
    }
    return duplicationError;
}
