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

// 削除処理
exports.deletingProcess = async function (deletSql) {
    var databaseError = '';
    let conn = null;
    try {
        try {
            conn = await mysql2.createConnection(mysql_setting);
            conn.connect();
            const [rows] = await conn.query(deletSql);
            conn.commit((err) => {
                if (err) {
                    return conn.rollback(() => {
                        throw err;
                    });
                }
            });
        } catch (sqlError) {
            console.log(e.stack);
            databaseError = 'データベースエラーが発生しました';
            console.log(databaseError);
        }

    } catch (e) {
        console.log(e.stack);
        console.log('削除処理で予期せぬエラー');
        throw e;
    } finally {
        conn.end();
    }
    return databaseError;
}

// 登録処理
exports.registeringProcess = async function (preSql, insData) {
    var databaseError = '';
    let conn = null;
    try {
        try {
            conn = await mysql2.createConnection(mysql_setting);
            conn.connect();
            const [rows] = await conn.query(preSql, insData);
            conn.commit((err) => {
                if (err) {
                    return conn.rollback(() => {
                        throw err;
                    });
                }
            });

        } catch (sqlError) {
            console.log(sqlError.stack);
            databaseError = 'データベースエラーが発生しました';
            console.log(databaseError);
        }

    } catch (e) {
        console.log(e.stack);
        console.log('登録処理で予期せぬエラー');
        throw e;
    } finally {
        conn.end();
    }
    return databaseError;
}

// 更新処理
exports.updatingProcess = async function (updataSql) {
    var databaseError = '';
    let conn = null;

    try {
        try {
            conn = await mysql2.createConnection(mysql_setting);
            conn.connect();
            const [rows] = await conn.query(updataSql);
            conn.commit((err) => {
                if (err) {
                    return conn.rollback(() => {
                        throw err;
                    });
                }
            });
        } catch (sqlError) {
            console.log(sqlError.stack);
            databaseError = 'データベースエラーが発生しました';
            console.log(databaseError);
        }
    } catch (e) {
        console.log(e.stack);
        console.log('更新処理で予期せぬエラー');
        throw e;
    } finally {
        conn.end();
    }
    return databaseError;
}