//nodemailer読みこみ
var nodemailer = require("nodemailer");
let cron = require('node-cron');
const mysql2 = require('mysql2/promise');
var HashMap = require('hashmap');
require('date-utils');


let mycon = null;
var mysql_setting = {
    host: 'localhost',
    user: 'root',
    password: 'root',
    database: 'contractPj'
}



//SMTPサーバ基本情報設定
var smtp = nodemailer.createTransport({
    host: "210.152.149.219",
    port: 25,
    ssl: false,
    use_authentication: true,
    user: "sasaki@groove-system.jp",
    pass: "sa0701saki",
    tls: {
        rejectUnauthorized: false
    }
});

//メール情報の作成
var message = {
    from: 'uzawa@groove-system.jp',
    to: 'hokari@groove-system.jp',
    cc: 'fukutome@groove-system.jp',
    subject: 'NodeTestMail',
    text: 'testText'
};

// メール送信時間の設定
cron.schedule('47 11 * * *', () => {
    sendMailJob();
});


async function sendMailJob() {

    //メール送信当日の日付を取得
    var dt = new Date();
    var formatted = dt.toFormat("YYYY年MM月DD日");


    try {


        mycon = await mysql2.createConnection(mysql_setting);

        // メールマスタからメール送信対象者の情報を取得
        //USER,PASSの取得
        const [rows, fields] = await mycon.query(' SELECT mail_type,mail_address,mail_pass FROM mail_master WHERE mail_type="USER" ');

        for (var i = 0; i < rows.length; i++) {
            var user = rows[i].mail_address;
            smtp.user = user;
            var pass = rows[i].mail_pass;
            smtp.pass = pass;

        }

        //FROMの取得
        const [rowsFrom, fieldsFrom] = await mycon.query(' SELECT mail_type,mail_address,mail_pass FROM mail_master WHERE mail_type="FROM" ');

        for (var i = 0; i < rowsFrom.length; i++) {
            var from = rowsFrom[i].mail_address;
            message.from = from;

        }

        //TOの取得
        const [rowsTo, fieldsTo] = await mycon.query(' SELECT mail_type,mail_address,mail_pass FROM mail_master WHERE mail_type="TO" ');

        var to = '';
        for (var i = 0; i < rowsTo.length; i++) {
            if (to != '') {
                to += ';';
            }
            to += rowsTo[i].mail_address;
            message.to = to;
        }

        //CCの取得
        const [rowsCC, fieldsCC] = await mycon.query(' SELECT mail_type,mail_address,mail_pass FROM mail_master WHERE mail_type="CC" ');
        var cc = '';
        for (var i = 0; i < rowsCC.length; i++) {

            if (cc != '') {
                cc += ';';
            }
            cc += rowsCC[i].mail_address;
            message.cc = cc;

        }

        // 1か月+1週間前のメール送信対象取得
        var user = '';
        var from = '';
        var to = '';
        var subject = '';

        mycon = await mysql2.createConnection(mysql_setting);
        const [rows2, fields2] = await mycon.query('SELECT end_date,SUBDATE(STR_TO_DATE(ymd,\'%Y-%m-%d\'),INTERVAL 7 DAY) as ymd2,a.engineer_name,proposition_company,general_contractor,proposition_manager,affiliation_company,affiliation_manager,year_month_date'
            + ' FROM latest_information a,'
            + ' (SELECT SUBDATE(STR_TO_DATE(end_date,\'%Y\/%m\/%d\'),INTERVAL 1 MONTH) as ymd,engineer_name'
            + ' FROM latest_information WHERE email_history=0 AND sheet_no = 1) b'
            + ' where a.engineer_name=b.engineer_name AND ymd <= now() AND a.sheet_no = 1'
            + ' ORDER BY end_date ASC');



        if (rows2.length == 0) {
            //  検索結果が無かった場合は1ヵ月+1週間前メールは送信しないので何も設定しない
        } else {
            //件名の設定
            var i = 0;
            subject = '【契約満了通知メール】1ヵ月+1週間前　' + formatted + '分';
            message.subject = subject;


            // 取得した対象の分だけ繰り返し本文を生成
            var honbun = '稼働終了1か月+1週間前の技術者は以下の通りです\n';
            for (var i = 0; i < rows2.length; i++) {
                honbun += rows2[i].engineer_name + 'さん　　稼働終了予定日：' + rows2[i].end_date + '\n';
            }
            message.text = honbun;
        }



        if (rows2.length == 0) {
            //  検索結果が無かった場合は1ヵ月+1週間前メールは送信しない
        } else {
            // 生成した件名、本文でのメール送信処理
            smtp.sendMail(message, function (error, success) {

                if (error) {
                    console.log(error.message);
                    return;
                }
                console.log(success.messageId);
            });


            // 1ヵ月+1週間前メールを送信した分だけ、メール送信済みフラグを送信済み「１」にする
            console.log(rows.length);
            for (var i = 0; i < rows2.length; i++) {

                const [rows3, fields3] = await mycon.query('UPDATE latest_information SET email_history=1'
                    + ' WHERE engineer_name=\'' + rows2[i].engineer_name + '\''
                    + ' AND proposition_company= \'' + rows2[i].proposition_company + '\''
                    + ' AND proposition_manager=\'' + rows2[i].proposition_manager + '\''
                    + ' AND affiliation_company=\'' + rows2[i].affiliation_company + '\''
                    + ' AND affiliation_manager=\'' + rows2[i].affiliation_manager + '\''
                    + ' AND year_month_date=\'' + rows2[i].year_month_date + '\'');
            }

        }


        // 1ヶ月+1日前のメール送信対象取得
        const [rowsSelect2, fields3] = await mycon.query('SELECT end_date,SUBDATE(STR_TO_DATE(ymd,\'%Y-%m-%d\'),INTERVAL 1 DAY) as ymd2,a.engineer_name,proposition_company,proposition_manager,affiliation_company,affiliation_manager,year_month_date'
            + ' FROM latest_information a,'
            + ' (SELECT SUBDATE(STR_TO_DATE(end_date,\'%Y\/%m\/%d\'),INTERVAL 1 MONTH) as ymd,engineer_name'
            + ' FROM latest_information WHERE email_history=1 AND sheet_no = 1) b'
            + ' where a.engineer_name=b.engineer_name AND ymd <= now() AND a.sheet_no = 1');



        if (rowsSelect2.length == 0) {
            //1か月+1日前メール送信対象者検索結果なしの場合メール送信しないのでなにもしない

        } else {
            //件名の設定
            var i = 0;
            subject = '【契約満了通知メール】1ヵ月前　' + formatted + '分';
            message.subject = subject;

            // 取得した対象の分だけ繰り返し本文を生成
            var honbun = '稼働終了1か月前の技術者は以下の通りです\n';
            for (var i = 0; i < rowsSelect2.length; i++) {
                honbun += rowsSelect2[i].engineer_name + 'さん　　稼働終了予定日：' + rowsSelect2[i].end_date + '\n';
            }
            message.text = honbun;
        }


        if (rowsSelect2.length == 0) {
            //検索結果が無かった場合1か月+1日前メール送らない 
            //1か月+1週間前＆1か月+1日前の両方該当者がいないメールを送信するために件名、本文を設定
            //件名
            subject = '【契約満了通知メール】　該当者なし　' + formatted + '分';
            message.subject = subject;
            //本文
            honbun = '本日稼働終了前通知対象の技術者はいません';
            message.text = honbun;

            //生成したメッセージのメール送信処理
            smtp.sendMail(message, function (error, success) {

                if (error) {
                    console.log(error.message);
                    return;
                }
                console.log(success.messageId);
            });

        } else {
            //1か月+1週間前の検索結果ありの場合通知メール送信
            // 生成したメッセージのメール送信処理
            console.log('1か月+1日前メール送る');
            smtp.sendMail(message, function (error, success) {

                if (error) {
                    console.log(error.message);
                    return;
                }
                console.log(success.messageId);
            });
        }

        // 1ヵ月+1日前メールを送信した分だけ、メール送信済みフラグを送信済み「２」にする
        for (var i = 0; i < rowsSelect2.length; i++) {

            const [rows3, fields3] = await mycon.query('UPDATE latest_information SET email_history=2'
                + ' WHERE engineer_name=\'' + rowsSelect2[i].engineer_name + '\''
                + ' AND proposition_company= \'' + rowsSelect2[i].proposition_company + '\''
                + ' AND proposition_manager=\'' + rowsSelect2[i].proposition_manager + '\''
                + ' AND affiliation_company=\'' + rowsSelect2[i].affiliation_company + '\''
                + ' AND affiliation_manager=\'' + rowsSelect2[i].affiliation_manager + '\''
                + ' AND year_month_date=\'' + rowsSelect2[i].year_month_date + '\'');
        }

        mycon.end();

    } catch (e) {
        
        console.log("Caught Exception", e);
        //件名
        subject = '【契約満了通知メール】システムエラー　' + formatted + '分';
        message.subject = subject;
        //本文
        honbun = 'システムエラーが発生しました。';
        message.text = honbun;

        //生成したメッセージのメール送信処理
        smtp.sendMail(message, function (error, success) {

            if (error) {
                console.log(error.message);
                return;
            }
            console.log(success.messageId);
        });
    } finally {
    }
}