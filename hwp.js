const Lw = '\u200b'.repeat(500);
function response(room, msg, sender, isGroupChat, replier, imageDB, packageName) {
    if (msg.startsWith("!hwp 추가 ")) {
        let txt = msg.replace("!hwp 추가 ", "").split("|")
        addAccount(txt[0], txt[1])
        replier.reply("[HWP] 성공적으로 계정을 추가하였습니다.")
    }
    if (msg.startsWith("!hwp 제거 ")) {
        let txt = msg.replace("!hwp 제거 ", "")
        removeAccount(txt)
        replier.reply("[HWP] 성공적으로 계정을 제거하였습니다.")
    }
    if (msg.startsWith("!hwp 목록")) {
        replier.reply("계정 목록\n" + Lw + accountList())
    }
}
function convertObjectToString(obj) {
    const keys = Object.keys(obj);
    let result = '';
    for (let i = 0; i < keys.length; i++) {
        result += (i + 1) + "." + keys[i] + " | " + obj[keys[i]] + "\n";
    }
    return result;
}
function accountList() {
    const connection = org.jsoup.Jsoup.connect("http://34.22.72.47:5356/accountList")
        .ignoreContentType(true)
        .ignoreHttpErrors(true)
        .get()
        .text()
    return convertObjectToString(JSON.parse(connection));
}
function addAccount(id, pw) {
    const connection = org.jsoup.Jsoup.connect("http://34.22.72.47:5356/addAccount?id=" + id + "&pw=" + pw)
        .ignoreContentType(true)
        .ignoreHttpErrors(true)
        .get()
        .text()
    return true;
}
function removeAccount(id, pw) {
    const connection = org.jsoup.Jsoup.connect("http://34.22.72.47:5356/removeAccount?id=" + id)
        .ignoreContentType(true)
        .ignoreHttpErrors(true)
        .get()
        .text()
    return true;
}