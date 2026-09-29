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

        // 別の画面を新しいタブで開きます。
        // window.close() はスクリプトで開いたタブしか閉じられないため、
        // 開くときもスクリプトを通します。リンクの onclick から呼び出し、戻り値を return すること。
        // 戻り値が true のときは、リンク本来の target="_blank" の動作にまかせます。
        openScreen: function (url) {
            var opened = window.open(url, "_blank");
            // ポップアップブロックなどで開けなかったとき
            if (!opened) {
                return true;
            }
            opened.focus();
            return false;
        },

        // 自分のタブを閉じます。別のタブで開いた画面の「閉じる」から呼び出します。
        // openScreen で開かれていれば閉じられます。直接 URL を開いた場合など、
        // ブラウザの制限で閉じられないときは手で閉じてもらうよう案内します。
        // 呼び出し側は戻り値を return してポストバックを止めること。
        closeTab: function () {
            window.close();
            window.setTimeout(function () {
                if (!window.closed) {
                    window.alert("このタブを閉じて、取込画面に戻ってください。");
                }
            }, 300);
            return false;
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

        // 時間のかかる処理の開始時に、ボタンを無効化して処理中の表示を出します。
        // 連打による二重登録を防ぎます。
        // 呼び出し側では戻り値を return しないこと。
        // UseSubmitBehavior="false" のボタンは onclick の後ろに __doPostBack が続くため、
        // return するとポストバックが実行されません。
        beginProcessing: function (button) {
            var panel = document.querySelector(".atom-processing");
            if (panel) {
                panel.style.display = "block";
            }
            var buttons = document.querySelectorAll(".atom-action-button");
            for (var i = 0; i < buttons.length; i++) {
                // リンク（a 要素）のとき。disabled 属性を持たないため見た目と操作を CSS で止めます。
                if (buttons[i].tagName === "A") {
                    buttons[i].className += " atom-action-disabled";
                    buttons[i].setAttribute("aria-disabled", "true");
                // ボタンのとき
                } else {
                    buttons[i].disabled = true;
                }
            }
            if (button) {
                button.disabled = true;
            }
            return true;
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
