# docRotater-gas
## Document rotate/backup tool for Google Apps Script.

### Features:
- Basically it's a tool to rotate/backup a meeting minute on Google Drive.
- Backup your document at everyday, everyweek or the specified day(s) of week.
- Backup files are stored in a folder of '過去の議事録.${yyyy}年' -- which can be modified on the source.
- Rename file when backup, like 'document.2023.01.01' to 'document.2023.01.07' according to execution date.
	- Date format is yyyy/mm/dd (JPN locale).

### Limitations:
- can handle only 'Document' type file.
- Not acceptable Spreadsheet, Slide etc.

### How to install:
- just paste content of 'source.dev/code.js' into your Google Apps Script editor.
- Presumably because of Google Editor's limitaion, you can not paste whole content of code.js at once.
	- Paste approx 50 lines at a time...
	- Or use [clasp](https://github.com/google/clasp) to upload the source to your Google Editor.

### How to use:
- Open your document, you'll see [docRotater/議事録保存] on the menu bar.
- On the drop-down menu...
	- Select the day(s) of week to backup your document.
	- Select the same day to clear the setting corresponding to it.
	- Select 'Delete All / 設定解除' to delete all settings.


## Googleドキュメントを自動バックアップするためのツール

### 機能概要:
- GoogleDrive上に保存してある議事録をローテーション(バックアップ)するためのツールです。
- GoogleDrive上の文書ファイルを毎日、毎週、あるいは指定した曜日にバックアップできます。
- バックアップは「過去の議事録.${yyyy}年」というフォルダに作成されます(スクリプト上で変更可能)。
- バックアップ時に ファイル名を自動変更します 
	- 例) ファイル.2023.01.01 → ファイル.2023.01.07  ※日付はバックアップ実行日時

### 制限事項:
- 今のところ文書ファイルのみ対応
- スプレッドシートやスライドには使用不可

### インストール方法:
- レポジトリ source.dev/code.js の中身を Google Apps Script のエディタに貼り付ける
- (恐らく)Googleエディタの都合により、code.js を一回で丸ごと copy&paste することは不可能な模様
	- 50行ずつ程度に分けて貼り付ければ copy&paste できるハズ
	- あるいは [clasp](https://github.com/google/clasp) を使って upload する方法もアリ

### 使用方法:
- (code.jsのインストール後) GoogleDriveのファイルを開きなおすと、画面上部のメニューバーに [docRotater/議事録保存] が出現します。
- [docRotater/議事録保存] をクリックして開いたドロップダウンメニューで...
	- 曜日を選択すれば、その曜日に自動バックアップが実行されるようになります。
	- 選択済みの曜日を再選択すると、その曜日のバックップ設定は解除されます。
	- 「Delete All / 設定解除」を選択すれば、すべての曜日での自動バックアップ設定が解除されます。

### スクリプトの設定変更情報:
```
// ドキュメントファイルのオーナーのみがスクリプトを使用可能にするフラグ 
//		true = オーナーのみ使用可能、 false=誰でも使用可能
//		※true推奨。falseにするとオーナー以外が自動実行設定できてしまうので運用がややこしくなります。
const ONLY_OWNER_CAN_RUN = true; 


// バックアップ先のフォルダ名。
//		ドキュメントの存在するフォルダと同じ場所にフォルダが作成されます。
//		${yyyy} ${mm} ${dd} はバックアップ実行日付に、
//		${filename} は元のファイル名に置換されます。
const BACKUP_FOLDER = '過去の議事録.${yyyy}年';


// バックアップ後のファイル名。
//		※ファイル名自体に日付文字列が含まれる場合、この指定は無視されます(ファイル名に含まれる日付文字列が更新されます)
//		${yyyy} ${mm} ${dd} はバックアップ実行日付に、
//		${filename} は元のファイル名に置換されます。
const BACKUP_FILENAME = '${filename}.${yyyy}.${mm}.${dd}';


// トリガー起動時間、配列で指定 (毎日 nn 時ごろに起動) ※起動失敗する場合があるので複数回設定する
//		1,3,5 と指定した場合、01:00、03:00、05:00 ごろに自動実行されます。
const TRIGGER_HOURS = [1, 3, 5];


// 設定保存用のキー名  ※変更不要
const PKEY_SETTING = 'DAYS_OF_WEEK';
const PKEY_LASTBACKUP = 'LAST_BACKUP';


// エラーリトライ回数  ※変更不要
//		※どうしてもエラーが多発する場合は増やしてみても良い(6以下を推奨)
const MAX_ERROR_RETRY = 4;
```
