# docRotater-gas
Document rotate/backup tool for Google Apps Script.

Limitations:
- can handle only 'Documet' type file.
- Not acceptable Spreadsheet, Slide etc.

How to use:
- just paste 'code.js' into your Google Apps Script editor.


Googleドキュメントを自動バックアップするためのツール

制限事項:
・今のところ文書ファイルのみ対応
・スプレッドシートやスライドには使用不可

使用方法:
・Googleドキュメント上メニューの「拡張機能→AppScript」を実行
・GoogleAppScript側で code.gs という(新規の)ファイルを作成 (中身は空っぽにしておく)
・このrepositoryにある code.js の中身を、↑で作成した code.gs に貼り付ける
・CTRL+S を押してGoogleAppScriptを保存
・元のGoogleドキュメントをリロードする
・Googleドキュメントに「docRotator/議事録保存」というメニューが出現する
