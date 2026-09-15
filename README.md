# Submission Scanner

小学校3年生の提出物管理用QRスキャナー。

## 構成

### index.html
GitHub Pagesで公開するフロントエンド。

役割：
- QRコード読み取り
- 提出物選択
- 管理キー入力
- GAS APIへのPOST
- 登録結果の表示

提出物内部コード：
- A = 視写
- B = 音読
- C = ぴったりドリル
- D = 算数ドリル

### Code.gs
Google Apps Script側のバックエンド。

役割：
- API_KEY認証
- 生徒IDの名簿照合
- 同日重複チェック
- Google Sheetsへの提出履歴記録
- 累計提出回数計算

## Google Sheets

名簿シート：
QR用名簿

列：
A 生徒ID
B クラス
C 出席番号
D 氏名

提出履歴シート：
提出履歴

列：
A タイムスタンプ
B 日付
C 生徒ID
D 氏名
E クラス
F 提出物

## Security

API_KEYの実値はGitHubに保存しない。

Apps Script:
Project Settings → Script Properties → API_KEY

に保存する。

## Production URL

GitHub Pages:
https://ayakaito113.github.io/submission-scanner/

Apps Script URLはindex.html内のGAS_URLを使用。
