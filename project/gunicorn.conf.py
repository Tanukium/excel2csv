import os
import logging
import logging.handlers
from logging.handlers import WatchedFileHandler
import multiprocessing
from e2c import settings

BASE_DIR = str(settings.BASE_DIR)

LOG_DIR = os.path.join(BASE_DIR, 'gunicorn_log')
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

# バインドするIPとポート
bind = "127.0.0.1:8000"

# 最大待機接続数、64-2048
backlog = 512

# タイムアウト
timeout = 30

keyfile = '/etc/letsencrypt/live/e2c.ninja/privkey.pem'
certfile = '/etc/letsencrypt/live/e2c.ninja/fullchain.pem'

# デバッグモード
debug = False

# gunicornが切り替える作業ディレクトリ
chdir = BASE_DIR

# ワーカープロセスタイプ（デフォルトはsyncモード、他にeventlet、gevent、tornado、gthread、gaiohttpがある）
worker_class = 'sync'

# ワーカープロセス数
workers = multiprocessing.cpu_count()

# 各ワーカープロセスが開くスレッド数を指定
threads = multiprocessing.cpu_count() * 2

# ログレベル（エラーログのレベルを指す。アクセスログのレベルは設定できない）
loglevel = 'info'

# ログフォーマット
access_log_format = '%(t)s %(p)s %(h)s "%(r)s" %(s)s %(L)s %(b)s %(f)s" "%(a)s"'
# 各オプションの意味は以下の通り：
'''
h          リモートアドレス
l          '-'
u          現在は'-'、将来のリリースではユーザー名になる可能性がある
t          リクエストの日付
r          ステータスライン（例：``GET / HTTP/1.1``）
s          ステータス
b          レスポンスの長さまたは'-'
f          リファラー
a          ユーザーエージェント
T          リクエスト時間（秒）
D          リクエスト時間（マイクロ秒）
L          リクエスト時間（10進秒）
p          プロセスID
'''

# アクセスログファイル
accesslog = os.path.join(LOG_DIR, 'gunicorn_access.log')
# エラーログファイル
errorlog = os.path.join(LOG_DIR, 'gunicorn_error.log')
# pidファイル
pidfile = os.path.join(LOG_DIR, 'gunicorn_error.pid')

# アクセスログファイル、"-"は標準出力を意味する
accesslog = "-"
# エラーログファイル、"-"は標準出力を意味する
errorlog = "-"
