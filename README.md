# atena-generator
Googleスライドで年賀状の宛名面を作るGASです。  
年賀状のように下部にお年玉番号があるレイアウト用です。

- 使うサービス
  - Google スライド: 宛名面レイアウト
  - Google スプレッドシート: 住所録
  - Google AppsScript: もろもろの処理

![img](./atena-sample.png)

## できること・できないこと
- 自動でできること
  - スプシの住所データを元に宛名面スライドを作る
  - 宛名・差出人が単独か連名かによって位置を調整する
  - 住所データに登録した住所を縦書き用に変換する
- 手動ですること
  - スプシへの住所データの登録
  - スライドの初期設定
  - 宛名面スライドで印刷範囲外になっている部分がないかを見て配置を微調整する
  - スライドのPDF保存、印刷
- できないこと(※宛名スライド生成後に手動で調整すれば可能)
  - 宛名や差出人を3人以上指定する
  - 宛名や差出人の2人目に異なる苗字を指定する

## 準備手順
### Googleスライド
1. Googleスライドを新規で作成する
2. 分かりやすいように名前を付ける
3. はがきサイズにする
   1. メニュー「ファイル > ページ設定」を開く
   1. 1つ目プルダウンで「カスタム」を選択
   1. 幅[ 10 ] x 高さ[ 14.8 ] 単位[ cm ] に設定

### GoogleスライドにGASを設定する
1. メニュー「拡張機能 > Apps Script」を開く
2. 分かりやすいようにプロジェクト名を設定する
3. 初期ファイルを ./main.gs で置き換える

### GASを使って初期設定をする
1. 拡張メニューを表示する
   1. スライドに戻り、一度リロードする
   1. メニュー右端に「宛名作成」が増えていることを確認
2. テンプレスライド
   1. メニュー「宛名作成 > 初期設定 > テンプレスライド追加」を選ぶ
3. 住所録スプシ
   1. メニュー「宛名作成 > 初期設定 > 住所録スプシ作成」を選ぶ
      - ※トップディレクトリにスプシが自動作成され、スクリプトプロパティが設定されます
   1. 住所録リンクがダイアログで表示されるので開く
   1. 必要に応じてフォルダ移動する

## 使い方
1. 住所録に宛名・差出人情報を入力する
2. メニュー「宛名作成 > 宛名スライド作成」から作成する

印字範囲の確認がしたければ、メニュー「宛名作成 > 印字範囲: 表示」で印字範囲を示す図形を各スライドに追加します  
「宛名作成 > 印字範囲: 非表示」で印字範囲を示す図形を削除できます  
