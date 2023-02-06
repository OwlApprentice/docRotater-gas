# docRotater-gas
## Document rotate/backup tool for Google Apps Script.

### Limitations:
- can handle only 'Document' type file.
- Not acceptable Spreadsheet, Slide etc.

### How to use:
- just paste content of 'source.dev/code.js' into your Google Apps Script editor.
- Presumably because of Google Editor's limitaion, you can not paste whole content of code.js at once.
	- Paste approx 50 lines at a time...
	- Or use [clasp](https://github.com/google/clasp) to upload the source to your Google Editor.


## Googleドキュメントを自動バックアップするためのツール

### 制限事項:
- 今のところ文書ファイルのみ対応
- スプレッドシートやスライドには使用不可

### 使用方法:
- レポジトリ source.dev/code.js の中身を Google Apps Script のエディタに貼り付ける
- (恐らく)Googleエディタの都合により、code.js を一回で丸ごと copy&paste することは不可能な模様
	- 50行ずつ程度に分けて貼り付ければ copy&paste できるハズ
	- あるいは [clasp](https://github.com/google/clasp) を使って upload する方法もアリ

