// ATOM 共通スクリプト
// 画面共通の再読み込み・マニュアル表示・未保存データの確認を担います。
var AtomPage = (function () {
    "use strict";

    var isDirty = false;

    // 入力欄が変更されたら未保存とみなす
    function watchInput() {
        var content = document.querySelector(".atom-content");
        if (!content) {
            return;
        }
        content.addEventListener("input", markDirty, true);
        content.addEventListener("change", markDirty, true);
    }

    function markDirty() {
        isDirty = true;
    }

    return {
        // 再読み込み。POST の再送信を避けるため、現在の URL へ改めて遷移します。
        reload: function () {
            window.location.href = window.location.pathname + window.location.search;
            return false;
        },

        // マニュアルを新しいタブで開きます。
        openManual: function (url) {
            window.open(url, "_blank", "noopener");
        },

        // 未保存の入力があるときだけログアウトの確認を行います。
        confirmLogout: function () {
            if (!isDirty) {
                return true;
            }
            return window.confirm("入力中の内容は保存されていません。ログアウトしますか？");
        },

        // 登録処理などの完了後に、未保存の状態を解除します。
        clearUnsaved: function () {
            isDirty = false;
        },

        init: function () {
            watchInput();
        }
    };
})();

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () { AtomPage.init(); });
} else {
    AtomPage.init();
}
