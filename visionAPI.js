////////////////////////// Google cloud vision API の準備 //////////////////////////
function doGet(e) {
    var scriptProperties = PropertiesService.getScriptProperties();
    var accessToken = scriptProperties.getProperty("access_token");
    if (accessToken == null) {
        var param = {
            response_type: "code",
            client_id: scriptProperties.getProperty("client_id"),
            redirect_uri: getCallbackURL_(),
            state: ScriptApp.newStateToken()
                .withMethod("callback")
                .withArgument("name", "value")
                .withTimeout(2000)
                .createToken(),
            scope: "https://www.googleapis.com/auth/cloud-vision",
            access_type: "offline",
            approval_prompt: "force"
        };
        var params = [];
        for (var name in param) params.push(name + "=" + encodeURIComponent(param[name]));
        var url = "https://accounts.google.com/o/oauth2/auth?" + params.join("&");
        return HtmlService.createHtmlOutput('<a href="' + url + '" target="_blank">認証</a>');
    }
    return HtmlService.createHtmlOutput("<p>設定済です</p>");
}

function getCallbackURL_() {
    var url = ScriptApp.getService().getUrl();
    if (url.indexOf("/exec") >= 0) return url.slice(0, -4) + "usercallback";
    return url.slice(0, -3) + "usercallback";
}

function callback(e) {
    var credentials = fetchAccessToken_(e.parameter.code);
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty("access_token", credentials.access_token);
    scriptProperties.setProperty("refresh_token", credentials.refresh_token);
}

function fetchAccessToken_(code) {
    var prop = PropertiesService.getScriptProperties();
    var res = UrlFetchApp.fetch("https://accounts.google.com/o/oauth2/token", {
        method: "POST",
        payload: {
            code: code,
            client_id: prop.getProperty("client_id"),
            client_secret: prop.getProperty("client_secret"),
            redirect_uri: getCallbackURL_(),
            grant_type: "authorization_code"
        },
        muteHttpExceptions: true
    });
    return JSON.parse(res.getContentText());
}

function refreshAccessToken_() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var token = JSON.parse(
        UrlFetchApp.fetch("https://www.googleapis.com/oauth2/v4/token", {
            method: "POST",
            payload: {
                client_id: scriptProperties.getProperty("client_id"),
                client_secret: scriptProperties.getProperty("client_secret"),
                refresh_token: scriptProperties.getProperty("refresh_token"),
                grant_type: "refresh_token"
            },
            muteHttpExceptions: true
        }).getContentText()
    );
    return token.access_token;
}

////////////////////////// 認識させてみる //////////////////////////
// file を受け取り、OCR の結果を返す
function imageAnnotate(file) {
    var scriptProperties = PropertiesService.getScriptProperties();
    var accessToken = scriptProperties.getProperty("access_token");

    var payload = JSON.stringify({
        requests: [
            {
                image: {
                    content: Utilities.base64Encode(file.getBlob().getBytes())
                },
                features: [
                    {
                        type: "TEXT_DETECTION",
                        maxResults: 100
                    }
                ]
            }
        ]
    });

    var json = null;
    var requestUrl = "https://vision.googleapis.com/v1/images:annotate";
    while (true) {
        var response = UrlFetchApp.fetch(requestUrl, {
            method: "POST",
            headers: {
                authorization: "Bearer " + accessToken
            },
            contentType: "application/json",
            payload: payload,
            muteHttpExceptions: true
        });
        json = JSON.parse(response);
        if (json.error && (json.error.code == "401" || json.error.code == "403")) {
            // リフレッシュトークンを使ってアクセストークンを再取得しリトライする
            accessToken = refreshAccessToken_();
            scriptProperties.setProperty("access_token", accessToken);
            continue;
        }
        break;
    }
    return json;
}
