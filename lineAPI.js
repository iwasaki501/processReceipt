var CHANNEL_ACCESS_TOKEN = "*********";
var userId = "*********";
var _ = Underscore.load();
var sheet_id = "*********";
// フォルダを取得
var archive = DriveApp.getFolderById("*********");
var fail = DriveApp.getFolderById("*********");
var doubt = DriveApp.getFolderById("*********");

// 何かを POST されたときにそれを受け取る関数。events データを受け取って、一つずつ replyMessage 関数に渡す
function doPost(e) {
    var contents = e.postData.contents;
    var obj = JSON.parse(contents);
    var events = obj["events"];
    for (var i = 0; i < events.length; i++) {
        if (events[i].type == "message") {
            // イベント (送られてきたメッセージ) を一つずつ処理する
            replyMessage(events[i]);
        }
    }
}

/////////////////////////////////// LINE で何か送信する ///////////////////////////////////
// データを受け取り、LINE で送信する関数
function sendData(postData, url) {
    var options = {
        method: "post",
        headers: {
            "Content-Type": "application/json",
            Authorization: "Bearer " + CHANNEL_ACCESS_TOKEN
        },
        payload: JSON.stringify(postData)
    };
    UrlFetchApp.fetch(url, options);
}

// 特定ユーザに２択の質問を送る
// 金額: \d.+円
// 「変更します」「変更しません」
// のような形式
function pushComfirm(amount) {
    var url = "https://api.line.me/v2/bot/message/push";
    var postData = {
        to: userId,
        messages: [
            {
                type: "template",
                altText: "金額を確認してください",
                template: {
                    type: "confirm",
                    actions: [
                        {
                            type: "message",
                            label: "変更します",
                            text: "変更します"
                        },
                        {
                            type: "message",
                            label: "変更しません",
                            text: "変更しません"
                        }
                    ],
                    text: "金額: " + String(amount) + "円"
                }
            }
        ]
    };
    sendData(postData, url);
}

// 特定ユーザに画像を送信する
function pushImage(imageUrl) {
    var url = "https://api.line.me/v2/bot/message/push";
    var postData = {
        to: userId,
        messages: [
            {
                type: "image",
                originalContentUrl: imageUrl,
                previewImageUrl: imageUrl
            }
        ]
    };
    sendData(postData, url);
}

// 特定ユーザにプレーンテキストを送る
function pushText(text) {
    var url = "https://api.line.me/v2/bot/message/push";
    var postData = {
        to: userId,
        messages: [
            {
                text: text,
                type: "text"
            }
        ]
    };
    sendData(postData, url);
}

/////////////////////////////////// スプレッドシートの処理 ///////////////////////////////////
// doubtSheet の Flag が格納されている列を検索し、1 が入っている行番号を返す
// EXCEL でいう VLOOKUP 関数のようなことをしたいのだが、サポートされていないようなので Underscore for GAS ライブラリを用いて行列転置した上で検索している
function searchFlag(doubtSheet) {
    var arrData = doubtSheet.getDataRange().getValues();
    // 行と列の入れ替え
    var arrTrans = _.zip.apply(_, arrData);
    // G 列、つまり Flag が格納されている列で 1 が入っている (Flag が付いている) 行番号を返す
    // 存在しない場合は 「-1」 が、存在する場合は「行番号 - 1」が返る
    return arrTrans[6].indexOf(1);
}

///////////////////////////////////////// 処理全体 /////////////////////////////////////////
// イベントを受け取り、メッセージの内容に基づいて返事あるいは処理をする
function replyMessage(event) {
    if (event.message.type != "text") {
        return false;
    }
    var text = event.message.text;
    // 「変更します」、「変更しません」、数字以外のものは無視する
    if ((text !== "変更します") & (text !== "変更しません") & isNaN(text)) {
        return false;
    }
    var spreadsheet = SpreadsheetApp.openById(sheet_id);
    var doubtSheet = spreadsheet.getSheetByName("Doubt");
    // Flag に 1 が入っている (Flag 付きの) 行番号を取得する
    var rowNumWithFlag = searchFlag(doubtSheet) + 1;
    // Flag 付きの行が存在しない場合は終了
    if (rowNumWithFlag == 0) {
        return false;
    }
    // 「変更します」の場合は「金額を入力してください」を送信する
    if (text == "変更します") {
        pushText("金額を入力してください");
        return true;
    }
    // Zaim の起動
    var service = startZaim();
    if (!service) {
        return false;
    }
    // archive, fail シートの準備
    var archiveSheet = spreadsheet.getSheetByName("Archive");
    var failSheet = spreadsheet.getSheetByName("Fail");
    // Flag 付き行を抜き出す
    var range = "a" + String(rowNumWithFlag) + ":g" + String(rowNumWithFlag);
    var rowWithFlag = doubtSheet.getRange(range).getValues()[0];
    var fileId = rowWithFlag[0];
    // ファイルの取得
    var file = DriveApp.getFileById(fileId);
    // 変更せず送ることを想定し、doubtSheet の内容に従って payment 配列の準備をしておく
    var payment = {
        date: formatDate(rowWithFlag[1]),
        amount: rowWithFlag[2],
        place: rowWithFlag[3],
        category_id: rowWithFlag[4],
        genre_id: rowWithFlag[5]
    };
    if (text === "変更しません") {
        // 「変更しません」が送られてきたとき
        doubtSheet.deleteRow(rowNumWithFlag); // doubtSheet から該当行を削除
        if (payment.amount === 0) {
            // 金額 0 円であったときは fail に移動させる
            fail.addFile(file);
            doubt.removeFile(file);
            failSheet.appendRow([fileId]);
        } else {
            // 金額が 0 円以外であったときは archive に移動させ、Zaim に送信する
            sendToZaim(service, payment);
            archive.addFile(file);
            doubt.removeFile(file);
            archiveSheet.appendRow([fileId, payment.date, payment.amount, payment.place, payment.category_id, payment.genre_id]);
        }
        pushQuestion(0);
    } else if (!isNaN(text)) {
        // 数字が送られてきたとき
        var amount = parseInt(text);
        doubtSheet.deleteRow(rowNumWithFlag); // doubtSheet から該当行を削除
        if (amount === 0) {
            // 金額 0 円であったときは fail に移動させる
            fail.addFile(file);
            doubt.removeFile(file);
            failSheet.appendRow([fileId]);
        } else {
            // 金額が 0 円以外であったときは archive に移動させ、Zaim に送信する
            payment["amount"] = amount; // 金額の変更
            sendToZaim(service, payment);
            archive.addFile(file);
            doubt.removeFile(file);
            archiveSheet.appendRow([fileId, payment.date, payment.amount, payment.place, payment.category_id, payment.genre_id]);
        }
        pushQuestion(0);
    }
}

// 必要に応じてユーザに画像や質問などを送信する
// - time-driven trigger で定期実行される
// - イベント処理が一つ終わったあと、replyMessage により 0 を引数として呼び出される
// の 2 通りの方法で実行される
function pushQuestion(trigger) {
    // doubtSheet の取得
    var spreadsheet = SpreadsheetApp.openById(sheet_id);
    var doubtSheet = spreadsheet.getSheetByName("Doubt");
    // Flag 付き行の番号を取得
    var rowNumWithFlag = searchFlag(doubtSheet) + 1;
    // flag 付き行が存在する (つまり、現在何かの行を処理中である) 場合は中止
    if (rowNumWithFlag !== 0) {
        return false;
    }
    // 以下、flag 付き行が存在しない場合の処理
    // 1. doubtSheet が空 & time-driven trigger により呼び出された -> 何もしない
    // 2. doubtSheet が空 & replyMessage により呼び出された -> 処理が全て終了したことを意味するので、"Completed!" を送信する
    // 3. doubtSheet に何か溜まっている -> 先頭行に 1 のフラグを付けて Question 送信、先頭行の処理を開始
    // doubtSheet にある最終行の行番号を取得
    var needReviewing = doubtSheet.getLastRow() - 1;
    // doubtSheet が空のとき
    if (needReviewing === 0) {
        // 上記の「2」、つまり replyMessage により呼び出されている
        if (!trigger) {
            pushText("Completed!");
        }
        // 上記の「1」、つまり定時実行されている
        return true;
    }
    // 以下は上記の「3」
    // セル G2 を 1 にする (先頭行に flag をつける)
    doubtSheet.getRange("G2").setValue(1);
    var rowWithFlag = doubtSheet.getRange("a2:g2").getValues()[0];
    var fileId = rowWithFlag[0];
    var file = DriveApp.getFileById(fileId);
    // 残りファイル数を記載した文章を送る
    pushText("Need Review: " + String(needReviewing));
    // ファイルの共有設定を変更し、リンクを知っていれば誰でも閲覧可能にする
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var imageUrl = "https://drive.google.com/uc?export=view&id=" + fileId;
    // 画像の送信
    pushImage(imageUrl);
    // ２択の質問を送る
    pushComfirm(rowWithFlag[2]);
}
