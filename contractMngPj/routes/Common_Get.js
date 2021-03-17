const e = require("express");
const mysql2 = require('mysql2/promise');

//HashMapの読み込み
var HashMap = require('hashmap');

//今日日付取得用
require('date-utils');

//MySQL宣言
var mysql_setting = {
    host: 'localhost',
    user: 'root',
    password: 'passpass',
    database: 'contractPj'
}

let XLSX = require('xlsx');
const moment = require('moment');

//登録年月が履歴情報か最新情報かの判定取得
exports.getHistoryLatest = function (date) {
    var dateDecision = 0;
    try {
        var today = new Date();
        var toDate = today.toFormat("YYYY/MM");
        // var toDate = '2100/12';
        if (date < toDate) {
            //履歴情報登録項目
            dateDecision = 0;
        } else {
            //最新情報登録項目
            dateDecision = 1;
        }

    } catch (e) {
        console.log(e.stack);
        console.log('履歴情報か最新情報かの判定取得で予期せぬエラー');
        throw e;
    }
    return dateDecision;
}


// シート名から集計情報登録項目取得
exports.getSummary = async function (sheetName) {
    var summary = 100;//あり得ない数値は-1で返す
    let conn = null;
    try {
        conn = await mysql2.createConnection(mysql_setting);
        const [rows, field] = await conn.query('SELECT summary_item FROM Sheet_master WHERE sheet_name = \'' + sheetName + '\'');
        if (rows.length == 0) {

            //検索結果有りの時➡集計情報登録項目を指定
        } else {
            summary = rows[0].summary_item;
        }
    } catch (e) {
        console.log(e.stack);
        console.log('集計情報登録項目取得で予期せぬエラー');
        throw e;
    } finally {
        conn.end();
    }
    return summary;
}

// シート名からシートNo取得
exports.getSheetNo = async function (sheetName) {
    var sheetNo = 0;
    let conn = null;
    try {

        conn = await mysql2.createConnection(mysql_setting);
        const [rows, field] = await conn.query('SELECT sheet_No FROM Sheet_master WHERE sheet_name = \'' + sheetName + '\'');
        if (rows.length == 0) {

            //検索結果有りの時➡シートNo代入
        } else {
            sheetNo = rows[0].sheet_No;
        }
    } catch (e) {
        console.log(e.stack);
        console.log('シートNo取得で予期せぬエラー');
        throw e;
    } finally {
        conn.end();
    }
    return sheetNo;
}



// シートマスタ情報(シートNo,シート名,集計情報登録項目)取得
exports.getSheetMaster = async function () {
    let conn = null;
    var sheetRows = null;
    try {

        conn = await mysql2.createConnection(mysql_setting);
        const [rows, field] = await conn.query('SELECT sheet_No,sheet_name,summary_item FROM Sheet_master');
        if (rows.length == 0) {

            //検索結果有りの時➡シートNo代入
        } else {
            sheetRows = rows;
        }
    } catch (e) {
        console.log(e.stack);
        console.log('シートNo取得で予期せぬエラー');
        throw e;
    } finally {
        conn.end();
    }
    return sheetRows;
}