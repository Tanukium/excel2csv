# Excel->CSV直行便
#### alias 'def excel2csv (file: excel) -> (file:csv)'
ExcelファイルをCSVファイルに自動変換するWebアプリです。<br>
2018年度前川ゼミプロジェクトです。

前川ゼミが参与している「信州デジタルコモンズプロジェクト／オープンデータプロジェクト」は、「地域の諸データのオープンデータ化」との課題の解決を目指しています。<br>
Excelの統計データを直接、データベース化するには難しいですが、CSVへ変換するにはそれほど難しくありません。<br>
このアプリで、少しでもオープンデータ化においての「はじめてのハードル」を取り除きたいと思います。<br>

長野大学企業情報学部49期生 シュク


# URL
[https://e2c.ninja](https://e2c.ninja)

# 使用技術
- Python 3.7
  - Xlrd(Third-party py library)
- Django 3.1
- SQLite
- Bootstrap
- Markdown
- Nginx
- Gunicorn
- Let's Encrypt
- AWS
  - VPC(Lightsail)
  - Route 53


# 機能一覧
- Excel->CSV変換機能(Xlrd)
- Excelファイルアップロード機能(django.forms.ModelForm)
  - 出力したCSVファイルダウンロード機能
  - アップロード/出力ファイルリスト表示機能(SQLite)
- 記事表示機能(Django, Markdown)
- データベースWeb管理機能(django.contrib.admin)
- SSL(Let's Encrypt)


# ToDo
- 短期
  - ~~AWSにデプロイ~~
  - ~~HTTPS(SSL)化~~
  - ~~CSSスマホ対応~~
  - ファイルリストにファイル削除機能(ボタン)追加
- 中期
  - テストを書く
  - MariaDB等のDBに切替る
  - Dockerで環境構築
- 長期
  - 変換精度向上(変換論理の修正/再構築)
