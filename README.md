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
  - EC2(EBS, SG, ACL, EIP)
  - IAM Policy/Role
  - CloudWatch Logs -> Lambda -> SNS
  - S3
  - Route 53(domain, DNS)
- Python 3
  - xlrd
  - Django
  - Gunicorn
  - boto3
- Nginx
- Let's Encrypt
- SQLite
- Bootstrap
- Markdown
- Google ReCaptcha V3

# 構成図

修正中

# Features

- Excel->CSV変換機能(xlrd,boto3)
- Excelファイルアップロード機能(django.forms.ModelForm)
  - 変換出力したCSVファイルアーカイブ機能(io.StringsIO, zipfile, boto3)
  - アーカイブしたZIPファイルダウンロード機能(boto3)
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
  - EC2にマイグレーション
  - 不審アップロード防止(ReCaptcha)
  - 投稿ページング機能
  - Nginx異常監視(CloudWatch Logs + Lambda + SNS)
  - S3でのDB運用(SQLite)
  - S3での`/media/`ファイル保存
  - S3バックエンドのexcel2csv機能
- Todo
  - 定期的EBSスナップショット取得設計
  - サーバ設定のユーザデータ化
    - ↑が実施済の場合、Auto-scaling/LBの導入
    - Dockerコンテナ化も検討
  - 変換精度向上(より高級な変換論理を掘り出す)
    - > 結合セルの解除, 空っぽセルの削除以外もっといい方法があるか？
