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

# 绑定的ip与端口
bind = "127.0.0.1:8000"

# 最大挂起的连接数，64-2048
backlog = 512

# 超时
timeout = 30

# 调试状态
debug = False

# gunicorn要切换到的目的工作目录
chdir = BASE_DIR

# 工作进程类型(默认的是 sync 模式，还包括 eventlet, gevent, or tornado, gthread, gaiohttp)
worker_class = 'sync'

# 工作进程数
workers = multiprocessing.cpu_count()

# 指定每个工作进程开启的线程数
threads = multiprocessing.cpu_count() * 2

# 日志级别，这个日志级别指的是错误日志的级别(debug、info、warning、error、critical)，而访问日志的级别无法设置
loglevel = 'info'

# 日志格式
access_log_format = '%(t)s %(p)s %(h)s "%(r)s" %(s)s %(L)s %(b)s %(f)s" "%(a)s"'
# 其每个选项的含义如下：
'''
h          remote address
l          '-'
u          currently '-', may be user name in future releases
t          date of the request
r          status line (e.g. ``GET / HTTP/1.1``)
s          status
b          response length or '-'
f          referer
a          user agent
T          request time in seconds
D          request time in microseconds
L          request time in decimal seconds
p          process ID
'''

# 访问日志文件
accesslog = os.path.join(LOG_DIR, 'gunicorn_access.log')
# 错误日志文件
errorlog = os.path.join(LOG_DIR, 'gunicorn_error.log')
# pid 文件
pidfile = os.path.join(LOG_DIR, 'gunicorn_error.pid')

# 访问日志文件，"-" 表示标准输出
accesslog = "-"
# 错误日志文件，"-" 表示标准输出
errorlog = "-"
