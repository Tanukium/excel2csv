# Excel->CSV直行便
#### alias 'def excel2csv (file: excel) -> (file: csv)'
ExcelファイルをCSVファイルに自動変換するWebアプリです。<br>
2018年度前川ゼミプロジェクトです。

前川ゼミが参与している「信州デジタルコモンズプロジェクト／オープンデータプロジェクト」は、「地域の諸データのオープンデータ化」との課題の解決を目指しています。<br>
Excelの統計データを直接、データベース化するには難しいですが、CSVへ変換するにはそれほど難しくありません。<br>
このアプリで、少しでもオープンデータ化においての「はじめてのハードル」を取り除きたいと思います。<br>

長野大学企業情報学部49期生 シュク


# URL
[https://e2c.ninja](https://e2c.ninja)

# 使用技術
- Python 3
  - xlrd(Third-party py library)
- Django >= 3.2.4
- SQLite
- Bootstrap
- Markdown
- ReCaptcha V3
- Nginx
- Gunicorn
- Let's Encrypt
- AWS
  - ~~VPC(Lightsail)~~
  - EC2 + EBS
  - CloudWatch + Amazon SNS
  - Route 53


# AWS構成図
![infra](https://github-ettnzncwtvxtk1wd.s3.ap-northeast-1.amazonaws.com/201115.drawio.png)


# 機能一覧
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


# ToDo
- 短期
  - [x] AWSにデプロイ
  - [x] HTTPS(SSL)化
  - [x] スマホ表示対応(CSS + Bootstrap)
  - [x] ファイル削除機能(ボタン)追加
    - [x] 権限で削除ボタンの表示をコントロールする
  - [x] 変換論理の書き直し(Keep it simple, stupid.)
    - [x] 変換によってデータ構造を破壊してしまうリスクのある論理を削除
      - > データ部とインデックス部の判断, etc.
    - [x] コードのメンテナンスしやすさ(可読性)向上
  - [ ] テストを書く
- 中期
  - [x] EC2に移行
  - [ ] S3に静的ファイルを保存
  - [ ] Django側の投げ出した異常をお知らせ(CloudWatch + Amazon SNS)
  - [ ] Dockerコンテナ/Lambda化
- 長期
  - [ ] 変換精度向上(より高級な変換論理を掘り出す)
    - > 結合セルの解除, 空っぽセルの削除以外もっといい方法があるか？
