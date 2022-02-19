# Excel->CSV直行便

ExcelファイルをCSVファイルに自動変換するWebアプリです。<br>
2018年度長野大学企業情報学部前川ゼミプロジェクトです。

前川ゼミが参与している「信州デジタルコモンズプロジェクト／オープンデータプロジェクト」は、「地域の諸データのオープンデータ化」との課題の解決を目指しています。<br>
Excelの統計データを直接、データベース化するには難しいですが、CSVへ変換するにはそれほど難しくありません。<br>
このアプリで、少しでもオープンデータ化においての「はじめてのハードル」を取り除きたいと思います。<br>

# URL

[https://e2c.ninja](https://e2c.ninja)

# 使用技術

- AWS
  - ~~Lightsail~~
  - EC2関連
  - CloudWatch Logs -> Lambda -> SNS
  - Route 53(domain, DNS)
- Python 3
  - xlrd
  - Django
  - Gunicorn
- Nginx
- Let's Encrypt
- SQLite
- Bootstrap
- Markdown
- ReCaptcha V3

# 構成図

修正中

# Features

- Excel->CSV変換機能(xlrd)
- Excelファイルアップロード機能(django.forms.ModelForm)
  - 変換出力したCSVファイルアーカイブ機能(shutil.make_archive)
  - アーカイブしたZIPファイルダウンロード機能
  - アップロード/変換出力ファイルリスト表示機能(SQLite)
  - アップロード/変換出力ファイル削除機能
- 記事表示機能(Django, Markdown)
- 記事ページング機能(Django, Bootstrap)
- 不審(Botによる)アップロード防止機能(django-recaptcha, ReCaptcha V3)
- データベースWeb管理機能(django.contrib.admin)
- SSL(Let's Encrypt)


# Tasks

- Done
  - Lightsailにデプロイ
  - HTTPS(SSL)化
  - スマホ表示対応(CSS + Bootstrap)
  - ファイル削除機能(ボタン)追加
    - 権限で削除ボタンの表示をコントロールする
  - 変換論理の書き直し
    - 変換によってデータ構造を破壊してしまうリスクのある論理を削除
      - > データ部とインデックス部の判断, etc.
    - コードの可読性向上
  - EC2に移行
  - 不審アップロード防止(ReCaptcha)
  - 投稿ページング
  - Nginx異常監視(CloudWatch Logs + Lambda + SNS)
- Doing
  - S3でのアップロードファイル保存
  - S3でのSQLiteDBファイル運用
- Todo
  - 定期的EBSスナップショット取得
  - Dockerコンテナ化
  - 変換精度向上(より高級な変換論理を掘り出す)
    - > 結合セルの解除, 空っぽセルの削除以外もっといい方法があるか？
