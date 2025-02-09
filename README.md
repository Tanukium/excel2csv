# Excel->CSV直行便

ExcelファイルをCSVファイルに簡易的に自動変換するWebアプリです。<br>
2018年度長野大学企業情報学部前川ゼミプロジェクトです。

前川ゼミが参与している「信州デジタルコモンズプロジェクト／オープンデータプロジェクト」は、「地域の諸データのオープンデータ化」との課題の解決を目指しています。<br>
Excelの統計データを直接、データベース化するには難しいですが、CSVへ変換するにはそれほど難しくありません。<br>
このアプリで、少しでもオープンデータ化においての「はじめてのハードル」を取り除きたいと思います。<br>

# URL

[https://e2c.ninja](https://e2c.ninja)

# 使用技術

- インフラ
  - AWS
    - EC2
    - IAM
    - CloudWatch Logs -> Lambda -> SNS
    - S3
    - Route 53(domain, DNS)
  - Terraform
  - WAF
- Backend
  - Python 3
    - xlrd
    - Django
    - Gunicorn
    - boto3
  - Nginx
  - Let's Encrypt
  - SQLite
- Frontend
  - Bootstrap
  - Markdown
  - Google ReCaptcha V3

# Features

- Excel->CSV変換(xlrd,boto3)
  - Excelファイルアップロード(django.forms.ModelForm)
  - 変換結果CSVファイル圧縮(io.StringsIO, zipfile,boto3)
  - 元Excelファイル/圧縮ファイル(ZIP)一覧表示(SQLite)
  - ファイルダウンロード(S3,boto3)
- 不審アップロード防止機能(django-recaptcha,GoogleReCaptcha V3)
- 記事表示機能(Django,Markdown)
- 記事ページング機能(Django,Bootstrap)
- SSL認証(Let's Encrypt)


# Tasks

- Done
  - Lightsailにデプロイ
  - HTTPS(SSL)化
  - スマホ表示対応(Bootstrap)
  - ファイル削除機能追加
    - 権限で削除可否をコントロール
  - 変換ロジックの書き直し
    - 変換によってデータ構造を破壊してしまうリスクのある論理を削除
      - > データ部とインデックス部の判断, etc.
    - コードの可読性向上
  - AWS EC2にマイグレーション
  - 不審アップロード防止対応(ReCaptcha)
  - 投稿ページング機能
  - Nginx異常監視(CloudWatchLogs+Lambda+SNS)
  - S3にDB格納(SQLite)
  - S3での静的ファイル保存
  - S3バックエンドのexcel2csv機能
  - 定期的EBSスナップショット取得設計(AWSBackup)
- Doing
  - IaC化(Terraform) 
- Todo
  - Dockerコンテナ化
