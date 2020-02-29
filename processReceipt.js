/////////////////////////////////// 日付の処理 ///////////////////////////////////
// date オブジェクトを yyyy-mm-dd 形式の文字列に変換する
function formatDate(date) {
    var y = date.getFullYear();
    var m = ("00" + (date.getMonth() + 1)).slice(-2);
    var d = ("00" + date.getDate()).slice(-2);
    var result = y + "-" + m + "-" + d;
    return result;
}

// fullAnnotation から日付を読み取り、yyyy-mm-dd 形式の文字列に変換して返す
// 日付が存在しない、あるいは不正な場合はとりあえずその日の日付にしておく
function findDate(oneline) {
    var today = new Date();
    var dateMatch = oneline.match(/20\d\d.*?[日(]/);
    // 日付らしきものが存在しない場合
    if (!dateMatch) {
        Logger.log("invalid date");
        return formatDate(today);
    }
    var dateStr = dateMatch[0];
    // 日付文字列の区切りを全部 / にする。空白は zero padding とみなして 0 とする
    dateStr = dateStr
        .replace("年", "/")
        .replace("月", "/")
        .replace("日", "")
        .replace(")", "")
        .replace("(", "")
        .replace(/\s/g, "0");
    // ここで yyyy/mm/dd の形になっているはず (mm は m でもよいし、dd は d でもよい)
    if (!dateStr.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
        Logger.log("invalid date");
        return formatDate(today);
    }
    // 日付文字列から Date オブジェクトを作成する
    var dateObj = new Date(dateStr);
    // 日付の妥当性を検証する。dateStr とそれから生成した dateObj の数字が一致していることを確認する
    if (dateObj.getFullYear() != dateStr.split("/")[0] || dateObj.getMonth() != dateStr.split("/")[1] - 1 || dateObj.getDate() != dateStr.split("/")[2]) {
        Logger.log("invalid date");
        return formatDate(today);
    }
    return formatDate(dateObj);
}

/////////////////////////////////// 金額の処理 ///////////////////////////////////
// ¥ マークの誤認識を訂正する
function correctYenMark(text) {
    var correctedText = text
        .replace(/半/g, "¥")
        .replace(/ギ/g, "¥")
        .replace(/羊/g, "¥")
        .replace(/ /g, "");
    return correctedText;
}

// キーワードの位置 (y 座標) を返す
function detectKeyWordHeight(textAnnotations, keyWord, exclusionKeyWord) {
    if (exclusionKeyWord) {
        var regExp = new RegExp("^(?=.*" + keyWord + ")(?!.*" + exclusionKeyWord + ")");
    } else {
        var regExp = new RegExp(keyWord);
    }
    for (var i = 1; i < textAnnotations.length; i++) {
        var text = textAnnotations[i].description;
        if (text.match(regExp)) {
            var keyWordUpperHeight = textAnnotations[i].boundingPoly.vertices[0].y;
            var keyWordLowerHeight = textAnnotations[i].boundingPoly.vertices[3].y;
            var keyWordHeight = (keyWordUpperHeight + keyWordLowerHeight) / 2;
            if (exclusionKeyWord && textAnnotations[i - 1].description.match(new RegExp(exclusionKeyWord + "$")) && textAnnotations[i - 1].boundingPoly.vertices[3].y > keyWordHeight) {
                continue;
            }
            return keyWordHeight;
        }
    }
    Logger.log("Failed to detect " + keyWord);
    return false;
}

// 金額の数字を抽出する
function parseAmount(textAnnotations, i) {
    var text = correctYenMark(textAnnotations[i].description);
    // ¥ だけのものはその後の数字と分離してしまっているため、一つ後ろのものを採用する
    if (text === "¥") {
        text = correctYenMark(textAnnotations[i + 1].description);
    }
    // カンマで終わっているものはそこで金額が途切れてしまっている可能性があるので、一つ後ろと連結する
    var count = 1;
    while (text.match(/,$/)) {
        text += textAnnotations[i + count].description;
        count += 1;
    }
    return parseInt(text.replace(",", "").replace("¥", ""));
}

// 「合」と同じ行 or すぐ下にある ¥ 入りの文字列を見つける
function findAmountByGoukei(textAnnotations) {
    var goukeiHeight = detectKeyWordHeight(textAnnotations, "合", "税");
    for (var i = 1; i < textAnnotations.length; i++) {
        // ¥ が入っていないものはスキップ
        if (!correctYenMark(textAnnotations[i].description).match(/\¥/)) {
            continue;
        }
        var textLowerHeight = textAnnotations[i].boundingPoly.vertices[3].y;
        // 「合」のある位置より下のものを捕捉する
        if (textLowerHeight >= goukeiHeight) {
            return parseAmount(textAnnotations, i);
        }
    }
    return false;
}

// 「費」と同じ行 or すぐ上にある ¥ 入りの文字列を見つける
function findAmountByShohi(textAnnotations) {
    var shohiHeight = detectKeyWordHeight(textAnnotations, "費", false);
    // 下から探していくので逆順にソートする
    for (var i = textAnnotations.length - 1; i > 0; i--) {
        // ¥ が入っていないものはスキップ
        if (!correctYenMark(textAnnotations[i].description).match(/\¥/)) {
            continue;
        }
        var textLowerHeight = textAnnotations[i].boundingPoly.vertices[3].y;
        // 「費」のある位置より上のものを捕捉する
        if (textLowerHeight <= shohiHeight) {
            return parseAmount(textAnnotations, i);
        }
    }
    return false;
}

// 金額を決定する
function findAmount(result) {
    var fullAnnotation = result.responses[0].textAnnotations[0].description;
    var textAnnotations = result.responses[0].textAnnotations;
    var amount = false;
    // OCR 結果で「合計」を正しく検出できているかを確認する
    // 本当は正しく「合計」を認識できていないのに、「預かり合計」「税合計」などに引っかかって存在すると間違えられるのを防ぐため、消しておく
    fullAnnotation = fullAnnotation.replace(/預.+計|税合計/g, "");
    Logger.log(fullAnnotation);
    // fullAnnotation に「合計」が存在するものだけ、まずは「合」の字で検索をかける。失敗した場合は false が返る
    if (fullAnnotation.match(/合.*計/)) {
        amount = findAmountByGoukei(textAnnotations);
    }
    // 「合計」がない場合、あるいは上の操作が失敗した場合は「費」の字で検索をかける。失敗した場合は false が返る
    if (!amount) {
        amount = findAmountByShohi(textAnnotations);
    }
    // false でなければそのまま返し、false の場合は 0 を返す
    if (amount) {
        return amount;
    } else {
        return 0;
    }
}

/////////////////////////////////// 場所の処理 ///////////////////////////////////
// 場所を決める
function findPlace(oneline, rows) {
    var place = oneline.match(/^.*?店/);
    // 「〜店」が存在する場合はそれを返す。存在しない場合は、rows の最初の３行を取って店名とする
    if (place) {
        place = place[0];
    } else {
        place = rows.slice(0, 3).join(" ");
    }
    // よくある誤認識の訂正
    place = place.replace("LAWEON", "LAWSON").replace("ライ7", "ライフ");
    return place;
}

///////////////////////////////// カテゴリの処理 /////////////////////////////////
// カテゴリ・ジャンルを決める
function decideCategory(place) {
    var shops = ["セブン", "Mart", "LAWSON", "売店", "食堂", "KIOSK", "マート", "ライフ", "スーパー", "アンスリー", "LOTTERIA", "ショップ"];
    var restaurants = ["レストラン", "カフェ", "食堂"];
    var category = {};
    // コンビニ・売店・スーパーなど
    for (var i = 0; i < shops.length; i++) {
        var regexp = new RegExp(shops[i], "i");
        if (place.match(regexp)) {
            category.category_id = 101;
            category.genre_id = 10101;
            return category;
        }
    }
    // 外食
    for (var i = 0; i < restaurants.length; i++) {
        var regexp = new RegExp(restaurants[i], "i");
        if (place.match(regexp)) {
            category.category_id = 37254802;
            category.genre_id = 17031254;
            return category;
        }
    }
    // その他
    category.category_id = 199;
    category.genre_id = 19999;
    return category;
}

/////////////////////////////// レシートの処理全体 ///////////////////////////////
// ファイルごとに payment オブジェクトを作成する
function createPayment(result) {
    var fullAnnotation = result.responses[0].textAnnotations[0].description;
    // ありがちな誤検出を修正しておく
    fullAnnotation = fullAnnotation
        .replace(/半/g, "¥")
        .replace(/羊/g, "¥")
        .replace(/ｐ/, ")")
        .replace(/, /g, "")
        .replace(/,/g, "");
    var oneline = fullAnnotation.replace(/\n/g, " ");
    var rows = fullAnnotation.split(/\n/);
    // 日付の処理
    var date = findDate(oneline);
    // 金額の処理
    var amount = findAmount(result);
    // 場所の処理
    var place = findPlace(oneline, rows);
    // カテゴリーの決定
    // 配列 category は category_id と genre_id という 2 つの key を持つ
    var category = decideCategory(place);
    // 配列 payment の作成
    var payment = {
        date: date,
        place: place,
        amount: amount,
        category_id: category.category_id,
        genre_id: category.genre_id
    };
    return payment;
}

// ディレクトリ内にあるファイルの走査
function iterateFiles(service) {
    var folder = DriveApp.getFolderById("*********");
    var archive = DriveApp.getFolderById("*********");
    var fail = DriveApp.getFolderById("*********");
    var doubt = DriveApp.getFolderById("*********");
    var sheet_id = "*********";
    var spreadsheet = SpreadsheetApp.openById(sheet_id);
    var archiveSheet = spreadsheet.getSheetByName("Archive");
    var doubtSheet = spreadsheet.getSheetByName("Doubt");
    var failSheet = spreadsheet.getSheetByName("Fail");

    files = folder.getFiles();
    while (files.hasNext()) {
        var file = files.next();
        // OCR をかける
        var result = imageAnnotate(file);
        var fileId = file.getId();
        Logger.log(file.getName());
        // 文字を読み取れない場合は、ファイルを fail フォルダに入れて次に行く
        if (!result) {
            Logger.log("No text");
            fail.addFile(file);
            folder.removeFile(file);
            failSheet.appendRow([fileId]);
            continue;
        }
        var fullAnnotation = result.responses[0].textAnnotations[0].description;
        // 「ご利用明細票」という文字列を含む場合は、ファイルを fail フォルダに入れて次に行く
        if (fullAnnotation.match(/ご利用明細票/)) {
            Logger.log("明細");
            fail.addFile(file);
            folder.removeFile(file);
            failSheet.appendRow([fileId]);
            continue;
        }
        // レシートの処理
        // 配列 payment は date, amount, place, category_id, genre_id という key を持つ
        var payment = createPayment(result);
        // 100 円以下, 10000 円以上, 100 で割り切れるものは怪しいので doubt に入れて次に行く
        if (payment.amount < 100 || payment.amount >= 10000 || payment.amount % 100 == 0) {
            doubt.addFile(file);
            folder.removeFile(file);
            doubtSheet.appendRow([fileId, payment.date, payment.amount, payment.place, payment.category_id, payment.genre_id, 0]);
            continue;
        }
        // 何もなければ、Zaim に送って archive フォルダに入れる
        sendToZaim(service, payment);
        archive.addFile(file);
        folder.removeFile(file);
        archiveSheet.appendRow([fileId, payment.date, payment.amount, payment.place, payment.category_id, payment.genre_id]);
    }
}

function run() {
    var folder = DriveApp.getFolderById("*********");
    var ssFiles = folder.getFiles();
    if (!ssFiles.hasNext()) {
        return false;
    }
    var service = startZaim();
    if (service) {
        iterateFiles(service);
    }
}
