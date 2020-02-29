var CONSUMER_KEY = "*********";
var CONSUMER_SECRET = "*********";

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
    var service = getService();
    service.reset();
}

/**
 * Configures the service.
 */
function getService() {
    return (
        OAuth1.createService("Zaim")
            // Set the endpoint URLs.
            .setAccessTokenUrl("https://api.zaim.net/v2/auth/access")
            .setRequestTokenUrl("https://api.zaim.net/v2/auth/request")
            .setAuthorizationUrl("https://auth.zaim.net/users/auth")

            // Set the consumer key and secret.
            .setConsumerKey(CONSUMER_KEY)
            .setConsumerSecret(CONSUMER_SECRET)

            // Set the name of the callback function in the script referenced
            // above that should be invoked to complete the OAuth flow.
            .setCallbackFunction("authCallback")

            // Set the property store where authorized tokens should be persisted.
            .setPropertyStore(PropertiesService.getUserProperties())
    );
}

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
    var service = getService();
    var authorized = service.handleCallback(request);
    if (authorized) {
        return HtmlService.createHtmlOutput("認証できました！このページを閉じて再びスクリプトを実行してください。");
    } else {
        return HtmlService.createHtmlOutput("認証に失敗");
    }
}

// Zaim に支出を記録する
function sendToZaim(service, payment) {
    var url = "https://api.zaim.net/v2/home/money/payment?";
    var param = {
        mapping: 1,
        category_id: payment.category_id,
        genre_id: payment.genre_id,
        amount: payment.amount,
        date: payment.date,
        comment: payment.place
    };
    var params = [];
    for (var name in param) params.push(name + "=" + encodeURIComponent(param[name]));
    url = url + params.join("&");
    var response = service.fetch(url, {
        method: "POST"
    });
    var result = JSON.parse(response.getContentText());
    //    pushText(response);
    return result;
}

// Zaim の履歴を取得する
function getHistory(service) {
    var url = "https://api.zaim.net/v2/home/money";
    var response = service.fetch(url, {
        method: "get"
    });
    var result = JSON.parse(response.getContentText());
    Logger.log(JSON.stringify(result, null, 2));
}

// Zaim を起動する
function startZaim() {
    var service = getService();
    if (service.hasAccess()) {
        return service;
    } else {
        var authorizationUrl = service.authorize();
        Logger.log("次のURLを開いてZaimで認証したあと、再度スクリプトを実行してください。: %s", authorizationUrl);
        return false;
    }
}
