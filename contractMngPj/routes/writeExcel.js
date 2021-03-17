var recordAll = new Object();
var express = require('express');
var router = express.Router();
const fs = require('fs');

let xlsx = require('xlsx');
var HashMap = require('hashmap');

const XlsxPopulate = require('xlsx-populate');

//const XlsxPopulate = require('xlsx-populate');
const mysql2 = require('mysql2/promise');

var mysql_setting = {
    host: 'localhost',
    user: 'root',
    password: 'passpass',
    database: 'contractPj'
}

var k = require('./search_execute.js');
const xutil = xlsx.utils; //util変数にセットしておくと楽

/* GET users listing. */
router.get('/', function (req, res, next) {

    var msghash = new HashMap();
    var data = { errormsg: '', msghash: msghash, path: 'C:\\', fromYear: '', fromMonth: '', toYear: '', toMonth: '' }
    res.render('excelExport', data);
    console.log('初期表示');

});

router.post('/export', function (req, res, next) {

    var exportHistory = req.body.exportHistory;
    console.log("ボタン\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\" + exportHistory);
    var exportLatest = req.body.exportLatest;
    console.log("ボタン\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\" + exportLatest);
    var path = req.body.path;
    console.log('path', path);
    var fromyear = req.body.fromYear;
    var frommonth = req.body.fromMonth;
    var toyear = req.body.toYear;
    var tomonth = req.body.toMonth;
    var searchHash = jonControl(req, res, exportHistory, path, fromyear, frommonth, toyear, tomonth);
});


async function jonControl(req, res, exportHistory, path, fromyear, frommonth, toyear, tomonth) {
    //エクセルに出力する配列を宣言
    // シートマスタの検索


    try {
        mycon = await mysql2.createConnection(mysql_setting);
        console.log('接続OK');
        console.log(path);

        if (path.length != 0) {
            console.log('パスが入力されている');
            //指定のフォルダが存在するか
            if (fs.existsSync(path)) {


                const [rows, fields] = await mycon.query('select sheet_name,sheet_No,summary_item from sheet_master');
                let wb = xutil.book_new();

                var testArray = [];
                var oldsheetNo = 1;
                var newsheetNo = 1;
                var msghash = new HashMap();
                var errormsg = '';

                var searchHash = await serachInfo(req, exportHistory);
                if (excel_before_year + excel_before_month <= excel_after_year + excel_after_month) {

                    if (searchHash.get('rowsKey') != "") {
                        var M = searchHash.get('rowsKey')[0].year_month_date.split("/");
                        var YEAR = M[0];
                        var month = M[1];

                        var beforeYM = excel_before_year + '/' + excel_before_month
                        var afterYM = excel_after_year + '/' + excel_after_month;
                        var rset = '';

                        if (exportHistory == "過去情報の出力") {
                            rset = " year_month_date = \"" + beforeYM + "\"";
                        } else {
                            rset = " year_month_date = (SELECT MAX(year_month_date) FROM summary_table)";
                        }

                        //平均売上などを取得
                        mycon = await mysql2.createConnection(mysql_setting);
                        const [average] = await mycon.query(`SELECT 
                    average_sales,average_profit,operating_count,a_star,month_end_count,ordercount_new,ordercount_new_target,ordercount_GAP,sheet_No 
                    FROM summary_table WHERE `+ rset);

                        //元になるExcelファイルの読み込み


                        //選択された月の売上合計と稼働件数と粗利益
                        const [sum] = await mycon.query(`SELECT
                    SUM(earnings) AS SUM_earnings,
                    SUM(gross_profit) AS SUM_gross_profit,
                    operating_count AS operating_count
                    FROM management_history a INNER JOIN summary_table b
                    ON a.sheet_No = b.sheet_No AND a.year_month_date = b.year_month_date
                    WHERE a.sheet_No = '1' AND a. `+ rset);


                        //営業ごとの受注件数取得
                        const [ym] = await mycon.query(`select 
                    ordercount_sales,a.sheet_No,sales_name,year_month_date
                    from summary_table a INNER JOIN sheet_master b
                    on a.sheet_No = b.sheet_No
                    WHERE  (summary_item = 2 OR summary_item = 1) AND `+ rset);

                        //summaytableから受注件数などを取り出す
                        const [SUM] = await mycon.query(`SELECT 
                    SUM(month_end_count) AS SUM_month_end_count,
                    SUM(ordercount_new) AS SUM_ordercount_new,
                    SUM(ordercount_new_target) AS SUM_ordercount_new_target,
                    SUM(ordercount_GAP) AS SUM_ordercount_GAP
                    FROM summary_table WHERE `+ rset);

                        //営業ごとの受注件数取得(最新ボタン)
                        const [sales] = await mycon.query(`select 
                    ordercount_sales,a.sheet_No,sales_name,year_month_date
                    from summary_table a INNER JOIN sheet_master b
                    on a.sheet_No = b.sheet_No
                     WHERE year_month_date = (SELECT MAX(year_month_date) FROM summary_table) AND (summary_item = 2 OR summary_item = 1)`);

                        //読み込みたいファイル↓
                        XlsxPopulate.fromFileAsync("sample.xlsm")
                            .then(book => {
                                //シートの数だけ繰り返すーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーーー
                                for (var s = 0; s < rows.length; s++) {
                                    console.log('シート名' + rows[s].sheet_name);
                                    var sheetname = book.sheet(rows[s].sheet_name);

                                    var Summary_Item1 = 0;
                                    var Summary_Item2 = 0;
                                    var Summary_Item3 = 0;
                                    for (var i = 0; i < searchHash.get('rowsKey').length; i++) {

                                        if (searchHash.get('rowsKey')[i].sheet_No == s + 1) {
                                            //稼働一覧
                                            if (rows[s].summary_item == 1) {
                                                //summary_item1の結果分だけ罫線つける
                                                console.log("sheetname:" + sheetname);

                                                var deleteRow = sheetname.range('B' + (Summary_Item1 + 5) + ':P' + (Summary_Item1 + 5));
                                                var deleted = searchHash.get('rowsKey')[i].deleted;

                                                if (deleted == 1) {
                                                    deleteRow.style("fill", "a9a9a9");
                                                }

                                                var engineerRow = sheetname.range('B5:P' + (Summary_Item1 + 5));
                                                engineerRow.style("border", true);
                                                sheetname.cell("B" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].start_date);
                                                sheetname.cell("C" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].end_date);
                                                sheetname.cell("D" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].EU);
                                                sheetname.cell("E" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].general_contractor);
                                                sheetname.cell("F" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].proposition_company);
                                                sheetname.cell("G" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].proposition_manager);
                                                sheetname.cell("H" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].affiliation_company);
                                                sheetname.cell("I" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].affiliation_manager);
                                                sheetname.cell("J" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].engineer_name);
                                                sheetname.cell("K" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].skill);
                                                sheetname.cell("L" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].earnings);
                                                sheetname.cell("M" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].purchase_price);
                                                sheetname.cell("N" + (Summary_Item1 + 5)).value(searchHash.get('rowsKey')[i].legal_welfare_expense);
                                                // sheetname.cell("O"+(kadouitiran+5)).value(searchHash.get('rowsKey')[i].gross_profit);
                                                //sheetname.cell("P"+(kadouitiran+5)).value(gross_margin);
                                                Summary_Item1 += 1;
                                                newsheetNo = searchHash.get('rowsKey')[i].sheet_No;
                                            } else if (rows[s].summary_item == 2) {

                                                var deleteRow = sheetname.range('B' + (Summary_Item2 + 5) + ':S' + (Summary_Item2 + 5));
                                                var deleted = searchHash.get('rowsKey')[i].deleted;

                                                if (deleted == 1) {
                                                    deleteRow.style("fill", "a9a9a9");
                                                }

                                                //summary_item2の結果分だけ罫線つける
                                                engineerRow = sheetname.range('B5:S' + (Summary_Item2 + 5));
                                                engineerRow.style("border", true);
                                                sheetname.cell("B" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].start_date);
                                                sheetname.cell("C" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].end_date);
                                                sheetname.cell("D" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].EU);
                                                sheetname.cell("E" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].general_contractor);
                                                sheetname.cell("F" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].proposition_company);
                                                sheetname.cell("G" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].proposition_manager);
                                                sheetname.cell("H" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].affiliation_company);
                                                sheetname.cell("I" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].affiliation_manager);
                                                sheetname.cell("J" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].engineer_name);
                                                sheetname.cell("K" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].skill);
                                                sheetname.cell("L" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].earnings);
                                                sheetname.cell("M" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].purchase_price);
                                                sheetname.cell("N" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].legal_welfare_expense);
                                                //sheetname.cell("O"+(nao+5)).value(searchHash.get('rowsKey')[i].gross_profit);
                                                //sheetname.cell("P"+(nao+5)).value(gross_margin);
                                                sheetname.cell("Q" + (Summary_Item2 + 5)).value(searchHash.get('rowsKey')[i].proposition_number);
                                                //sheetname.cell("L"+(nao+5)).value(searchHash.get('rowsKey')[i].real_earnings);
                                                //sheetname.cell("S"+(nao+5)).value(searchHash.get('rowsKey')[i].real_gross_profit);
                                                Summary_Item2 += 1;

                                                newsheetNo = searchHash.get('rowsKey')[i].sheet_No;
                                            } else if (rows[s].summary_item == 0) {

                                                var deleteRow = sheetname.range('B' + (Summary_Item3 + 5) + ':P' + (Summary_Item3 + 5));
                                                var deleted = searchHash.get('rowsKey')[i].deleted;

                                                if (deleted == 1) {
                                                    deleteRow.style("fill", "a9a9a9");
                                                }

                                                //summary_item0の結果分だけ罫線つける
                                                engineerRow = sheetname.range('B5:S' + (Summary_Item3 + 5));
                                                engineerRow.style("border", true);
                                                sheetname.cell("B" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].start_date);
                                                sheetname.cell("C" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].end_date);
                                                sheetname.cell("D" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].EU);
                                                sheetname.cell("E" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].general_contractor);
                                                sheetname.cell("F" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].proposition_company);
                                                sheetname.cell("G" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].proposition_manager);
                                                sheetname.cell("H" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].affiliation_company);
                                                sheetname.cell("I" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].affiliation_manager);
                                                sheetname.cell("J" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].engineer_name);
                                                sheetname.cell("K" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].skill);
                                                sheetname.cell("L" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].earnings);
                                                sheetname.cell("M" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].purchase_price);
                                                sheetname.cell("N" + (Summary_Item3 + 5)).value(searchHash.get('rowsKey')[i].legal_welfare_expense);
                                                //sheetname.cell("O"+(nao+5)).value(searchHash.get('rowsKey')[i].gross_profit);
                                                //sheetname.cell("P"+(nao+5)).value(gross_margin),
                                                Summary_Item3 += 1;
                                                newsheetNo = searchHash.get('rowsKey')[i].sheet_No;
                                            }
                                        }
                                    }
                                    //シートNOが一つ前と変わっているかの判定
                                    if (oldsheetNo == newsheetNo) {
                                        console.log("job1");
                                        //まだシートNoが変わっていない"
                                        if (newsheetNo == 1 && oldsheetNo == 1) {
                                            console.log("job2");
                                            //区分によって表示するものが変わる
                                            if (rows[s].summary_item == 1) {
                                                console.log("job3");
                                                //選択された年月が同じかの判定
                                                if (exportHistory == "過去情報の出力") {
                                                    console.log("job4");
                                                    console.log("beforeYM:" + beforeYM + ",afterYM:" + afterYM);
                                                    //選択された年月が同じかの判定
                                                    if (beforeYM == afterYM) {
                                                        engineerRow = sheetname.range('L' + (Summary_Item1 + 7) + ':M' + (Summary_Item1 + 18));
                                                        engineerRow.style("border", true);

                                                        var speaceRow = sheetname.range('L' + (Summary_Item1 + 5) + ':M' + (Summary_Item1 + 6));
                                                        speaceRow.style("border", false);

                                                        sheetname.cell("L" + (Summary_Item1 + 5)).value("");
                                                        sheetname.cell("L" + (Summary_Item1 + 6)).value("");
                                                        sheetname.cell("M" + (Summary_Item1 + 5)).value("");
                                                        sheetname.cell("M" + (Summary_Item1 + 6)).value("");

                                                        sheetname.cell("K" + (Summary_Item1 + 5)).value("合計");
                                                        var sumRow = sheetname.range('K' + (Summary_Item1 + 5) + ':P' + (Summary_Item1 + 5));
                                                        sumRow.style("border", true);

                                                        sheetname.cell("L" + (Summary_Item1 + 7)).value("平均売上");
                                                        sheetname.cell("L" + (Summary_Item1 + 8)).value("平均利益");
                                                        sheetname.cell("L" + (Summary_Item1 + 9)).value("稼働件数");
                                                        sheetname.cell("M" + (Summary_Item1 + 9)).value(average[s].operating_count);
                                                        sheetname.cell("L" + (Summary_Item1 + 10)).value(month + "月未落ち件数");
                                                        sheetname.cell("M" + (Summary_Item1 + 10)).value(average[s].month_end_count);
                                                        sheetname.cell("L" + (Summary_Item1 + 11)).value(month + "月新規受注件数");
                                                        sheetname.cell("M" + (Summary_Item1 + 11)).value(average[s].ordercount_new);
                                                        sheetname.cell("L" + (Summary_Item1 + 12)).value(month + "月新規受注件数目標");
                                                        sheetname.cell("M" + (Summary_Item1 + 12)).value(average[s].ordercount_new_target);
                                                        sheetname.cell("L" + (Summary_Item1 + 13)).value(month + "月受注件数GAP");
                                                        sheetname.cell("M" + (Summary_Item1 + 13)).value(average[s].ordercount_GAP);
                                                        sheetname.cell("L" + (Summary_Item1 + 14)).value(month + "月渡辺受注件数");
                                                        sheetname.cell("M" + (Summary_Item1 + 14)).value(sales[1].ordercount_sales);
                                                        sheetname.cell("L" + (Summary_Item1 + 15)).value(month + "月木村受注件数");
                                                        sheetname.cell("M" + (Summary_Item1 + 15)).value(sales[2].ordercount_sales);
                                                        sheetname.cell("L" + (Summary_Item1 + 16)).value(month + "月大村受注件数");
                                                        sheetname.cell("M" + (Summary_Item1 + 16)).value(sales[3].ordercount_sales);
                                                        sheetname.cell("L" + (Summary_Item1 + 17)).value(month + "月中村受注件数");
                                                        sheetname.cell("M" + (Summary_Item1 + 17)).value(sales[4].ordercount_sales);
                                                        sheetname.cell("L" + (Summary_Item1 + 18)).value(month + "月大倉受注件数");
                                                        sheetname.cell("M" + (Summary_Item1 + 18)).value(sales[0].ordercount_sales);
                                                    }
                                                } else {
                                                    //最新情報表示
                                                    engineerRow = sheetname.range('L' + (Summary_Item1 + 7) + ':M' + (Summary_Item1 + 18));

                                                    var speaceRow = sheetname.range('L' + (Summary_Item1 + 5) + ':M' + (Summary_Item1 + 6));
                                                    speaceRow.style("border", false);

                                                    sheetname.cell("L" + (Summary_Item1 + 5)).value("");
                                                    sheetname.cell("L" + (Summary_Item1 + 6)).value("");
                                                    sheetname.cell("M" + (Summary_Item1 + 5)).value("");
                                                    sheetname.cell("M" + (Summary_Item1 + 6)).value("");

                                                    sheetname.cell("K" + (Summary_Item1 + 5)).value("合計");
                                                    var sumRow = sheetname.range('K' + (Summary_Item1 + 5) + ':P' + (Summary_Item1 + 5));
                                                    sumRow.style("border", true);

                                                    engineerRow.style("border", true);
                                                    sheetname.cell("L" + (Summary_Item1 + 7)).value("平均売上");
                                                    sheetname.cell("L" + (Summary_Item1 + 8)).value("平均利益");
                                                    sheetname.cell("L" + (Summary_Item1 + 9)).value("稼働件数");
                                                    sheetname.cell("M" + (Summary_Item1 + 9)).value(average[s].operating_count);
                                                    sheetname.cell("L" + (Summary_Item1 + 10)).value(month + "月未落ち件数");
                                                    sheetname.cell("M" + (Summary_Item1 + 10)).value(average[s].month_end_count);
                                                    sheetname.cell("L" + (Summary_Item1 + 11)).value(month + "月新規受注件数");
                                                    sheetname.cell("M" + (Summary_Item1 + 11)).value(average[s].ordercount_new);
                                                    sheetname.cell("L" + (Summary_Item1 + 12)).value(month + "月新規受注件数目標");
                                                    sheetname.cell("M" + (Summary_Item1 + 12)).value(average[s].ordercount_new_target);
                                                    sheetname.cell("L" + (Summary_Item1 + 13)).value(month + "月受注件数GAP");
                                                    sheetname.cell("M" + (Summary_Item1 + 13)).value(average[s].ordercount_GAP);
                                                    sheetname.cell("L" + (Summary_Item1 + 14)).value(month + "月渡辺受注件数");
                                                    sheetname.cell("M" + (Summary_Item1 + 14)).value(sales[1].ordercount_sales);
                                                    sheetname.cell("L" + (Summary_Item1 + 15)).value(month + "月木村受注件数");
                                                    sheetname.cell("M" + (Summary_Item1 + 15)).value(sales[2].ordercount_sales);
                                                    sheetname.cell("L" + (Summary_Item1 + 16)).value(month + "月大村受注件数");
                                                    sheetname.cell("M" + (Summary_Item1 + 16)).value(sales[3].ordercount_sales);
                                                    sheetname.cell("L" + (Summary_Item1 + 17)).value(month + "月中村受注件数");
                                                    sheetname.cell("M" + (Summary_Item1 + 17)).value(sales[4].ordercount_sales);
                                                    sheetname.cell("L" + (Summary_Item1 + 18)).value(month + "月大倉受注件数");
                                                    sheetname.cell("M" + (Summary_Item1 + 18)).value(sales[0].ordercount_sales);

                                                }
                                            } else {
                                                if (rows[s].summary_item == 2) {
                                                    if (beforeYM == afterYM) {

                                                        var speaceRow = sheetname.range('L' + (Summary_Item2 + 5) + ':M' + (Summary_Item2 + 6));
                                                        speaceRow.style("border", false);

                                                        sheetname.cell("L" + (Summary_Item2 + 5)).value("");
                                                        sheetname.cell("L" + (Summary_Item2 + 6)).value("");
                                                        sheetname.cell("M" + (Summary_Item2 + 5)).value("");
                                                        sheetname.cell("M" + (Summary_Item2 + 6)).value("");

                                                        sheetname.cell("K" + (Summary_Item2 + 5)).value("合計");
                                                        var sumRow = sheetname.range('K' + (Summary_Item2 + 5) + ':P' + (Summary_Item2 + 5));
                                                        sumRow.style("border", true);

                                                        engineerRow = sheetname.range('L' + (Summary_Item2 + 7) + ':M' + (Summary_Item2 + 10));
                                                        engineerRow.style("border", true);
                                                        sheetname.cell("L" + (Summary_Item2 + 7)).value('平均売上');
                                                        sheetname.cell("M" + (Summary_Item2 + 7)).value('');
                                                        sheetname.cell("L" + (Summary_Item2 + 8)).value('平均利益');
                                                        sheetname.cell("M" + (Summary_Item2 + 8)).value('');
                                                        sheetname.cell("L" + (Summary_Item2 + 9)).value('稼働件数');
                                                        sheetname.cell("M" + (Summary_Item2 + 9)).value('');
                                                        sheetname.cell("L" + (Summary_Item2 + 10)).value('A-STAR');
                                                        sheetname.cell("M" + (Summary_Item2 + 10)).value('');
                                                    }
                                                }
                                            }
                                            //稼働一覧以外のシートNOが同じその中の範囲選択の判定
                                        } else {
                                            if (rows[s].summary_item == 2) {
                                                if (beforeYM == afterYM) {

                                                    var speaceRow = sheetname.range('L' + (Summary_Item2 + 5) + ':M' + (Summary_Item2 + 6));
                                                    speaceRow.style("border", false);

                                                    sheetname.cell("L" + (Summary_Item2 + 5)).value("");
                                                    sheetname.cell("L" + (Summary_Item2 + 6)).value("");
                                                    sheetname.cell("M" + (Summary_Item2 + 5)).value("");
                                                    sheetname.cell("M" + (Summary_Item2 + 6)).value("");

                                                    sheetname.cell("K" + (Summary_Item2 + 5)).value("合計");
                                                    var sumRow = sheetname.range('K' + (Summary_Item2 + 5) + ':P' + (Summary_Item2 + 5));
                                                    sumRow.style("border", true);

                                                    engineerRow = sheetname.range('L' + (Summary_Item2 + 7) + ':M' + (Summary_Item2 + 10));
                                                    engineerRow.style("border", true);
                                                    sheetname.cell("L" + (Summary_Item2 + 7)).value('平均売上');
                                                    sheetname.cell("M" + (Summary_Item2 + 7)).value('');
                                                    sheetname.cell("L" + (Summary_Item2 + 8)).value('平均利益');
                                                    sheetname.cell("M" + (Summary_Item2 + 8)).value('');
                                                    sheetname.cell("L" + (Summary_Item2 + 9)).value('稼働件数');
                                                    sheetname.cell("M" + (Summary_Item2 + 9)).value('');
                                                    sheetname.cell("L" + (Summary_Item2 + 10)).value('A-STAR');
                                                    sheetname.cell("M" + (Summary_Item2 + 10)).value('');
                                                }
                                            } else if (rows[s].summary_item == 0) {
                                            }
                                        }
                                    } else {
                                        console.log("シートNoが変わった");
                                        var a = newsheetNo - oldsheetNo;
                                        oldsheetNo = newsheetNo;
                                        //oldsheetとnewsheetの差分が1以上か
                                        if (1 <= a) {
                                            if (rows[s].summary_item == 2) {
                                                if (beforeYM == afterYM) {


                                                    var speaceRow = sheetname.range('L' + (Summary_Item2 + 5) + ':M' + (Summary_Item2 + 6));
                                                    speaceRow.style("border", false);

                                                    sheetname.cell("L" + (Summary_Item2 + 5)).value("");
                                                    sheetname.cell("L" + (Summary_Item2 + 6)).value("");
                                                    sheetname.cell("M" + (Summary_Item2 + 5)).value("");
                                                    sheetname.cell("M" + (Summary_Item2 + 6)).value("");

                                                    sheetname.cell("K" + (Summary_Item2 + 5)).value("合計");
                                                    var sumRow = sheetname.range('K' + (Summary_Item2 + 5) + ':P' + (Summary_Item2 + 5));
                                                    sumRow.style("border", true);

                                                    engineerRow = sheetname.range('L' + (Summary_Item2 + 7) + ':M' + (Summary_Item2 + 10));
                                                    engineerRow.style("border", true);
                                                    sheetname.cell("L" + (Summary_Item2 + 7)).value('平均売上');
                                                    sheetname.cell("M" + (Summary_Item2 + 7)).value('');
                                                    sheetname.cell("L" + (Summary_Item2 + 8)).value('平均利益');
                                                    sheetname.cell("M" + (Summary_Item2 + 8)).value('');
                                                    sheetname.cell("L" + (Summary_Item2 + 9)).value('稼働件数');
                                                    sheetname.cell("M" + (Summary_Item2 + 9)).value(average[s].operating_count);
                                                    sheetname.cell("L" + (Summary_Item2 + 10)).value('A-STAR');
                                                    sheetname.cell("M" + (Summary_Item2 + 10)).value(average[s].a_star);
                                                }
                                            }
                                        }
                                    }
                                }
                                //↓で出力するExcelファイルの生成をしフォルダ名をつける
                                if (exportHistory == "過去情報の出力") {
                                    if (beforeYM == afterYM) {
                                        book.toFileAsync(path + '\\' + excel_before_year + "年" + excel_before_month + "月システム数字.xlsm");

                                    } else {
                                        book.toFileAsync(path + '\\' + excel_before_year + "年" + excel_before_month + "月" + "-" + excel_after_year + "年" + excel_after_month + "月システム数字.xlsm");
                                    }
                                } else {
                                    book.toFileAsync(path + '\\' + YEAR + "年" + month + "月システム数字.xlsm");
                                }
                            })

                    } else {
                        errormsg = '指定した年月で登録されている営業情報はありませんでした';
                    }
                } else {
                    errormsg = '出力年月は下限,上限正しく選択してください';
                }
            } else {
                errormsg = "出力先のパスが見つかりません";
            }
        } else {
            errormsg = '出力パスを入力して下さい';
        }
    } catch (e) {
        errormsg = 'システムエラーが発生しました';
    }


    if (errormsg == '') {
        if (exportHistory == "過去情報の出力") {
            if (beforeYM == afterYM) {
                msghash.set('MSG2', 'ファイル名:' + excel_before_year + "年" + excel_before_month + "月システム数字.xlsm");
            } else {
                msghash.set('MSG2', 'ファイル名:' + excel_before_year + "年" + excel_before_month + "月" + "-" + excel_after_year + "年" + excel_after_month + "月システム数字.xlsm");
            }
        } else {
            msghash.set('MSG2', 'ファイル名:' + YEAR + "年" + month + "月システム数字.xlsm");
        }
        msghash.set('MSG1', '出力完了しました');
        msghash.set('MSG3', '出力場所:' + path);
        errormsg = '';
    }
    fromyear, frommonth, toyear, tomonth
    var data = {
        errormsg: errormsg, msghash: msghash, path: path,
        fromYear: fromyear, fromMonth: frommonth, toYear: toyear, toMonth: tomonth
    };
    console.log(data);
    res.render('excelExport', data);
    module.exports = router;
}
//-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
async function serachInfo(req, exportHistory) {
    var searchHash = new HashMap();
    // var hmset = new HashMap();
    // var fromImpYmd = 0;
    // var toImpYmd = 0;

    if (exportHistory == "過去情報の出力") {
        console.log("過去");
        excel_before_year = req.body.fromYear;
        excel_before_month = req.body.fromMonth;
        excel_after_year = req.body.toYear;
        excel_after_month = req.body.toMonth;
    } else {
        // 最新の出力ボタン押した
        mycon = await mysql2.createConnection(mysql_setting);
        const [rows] = await mycon.query(
            `SELECT
            year_month_date,
            start_date,
            end_date,
            engineer_name, 
            EU,
            general_contractor,
            proposition_company,
            earnings,
            proposition_manager,
            affiliation_company,
            affiliation_manager,
            skill,
            legal_welfare_expense,
            purchase_price,
            gross_margin,
            gross_profit,
            deleted,
            proposition_number,
            real_earnings,
            real_gross_profit,
            sheet_No
            FROM management_history 
            WHERE year_month_date =
    (SELECT
    MAX(year_month_date)
    FROM
    management_history);`);

        var year_month = rows[0].year_month_date;
        console.log(year_month);
        var YM = year_month.split("/");
        excel_before_year = YM[0];
        excel_before_month = YM[1];
        excel_after_year = YM[0];
        excel_after_month = YM[1];
    }

    searchHash = await k.searchcheck(
        null, // キーワード
        0,    // ラジオボタン
        0,    // ジャンル
        0,    // シートサーチ
        0,    // ソートボタン
        0,    // 契約開始年(変更前)
        0,    // 契約開始月(変更前)
        0,    // 契約開始年(変更後)
        0,    // 契約開始月(変更後)
        0,    // 契約終了年(変更前)
        0,    // 契約終了月(変更前)
        0,    // 契約終了年(変更後)
        0,    // 契約終了月(変更後)
        0,    // Excel取り込み年月(ソート)
        0,    // 稼働開始日(ソート)
        0,    // 稼働終了日(ソート)
        0,    // エンド(ソート)
        0,    // 元請(ソート)
        0,    // 案件元(ソート)
        0,    // 案件元営業(ソート)
        0,    // 所属会社(ソート)
        0,    // 所属会社営業(ソート)
        0,    // 技術者名(ソート)
        0,    // スキル(ソート)
        0,    // 売上(ソート)
        0,    // 仕入(ソート)
        0,    // 法定福利(ソート)
        0,    // 粗利益(ソート)
        0,
        0,
        0,
        0,    // 粗利率(ソート)
        null,  // エラーカウント
        excel_before_year,
        excel_before_month,
        excel_after_year,
        excel_after_month,
        1//画面判定用変数（画面「０」Excel出力「１」）
        // fromImpYmd,  // 取り込み年月(FROM)
        // toImpYmd,  // 取り込み年月(TO)

    );

    return searchHash;

};

module.exports = router;
