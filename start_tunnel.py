# -*- coding: utf-8 -*-
"""
一键启动本地服务 + cloudflared 公网隧道
用法：双击运行 或 命令行: python start_tunnel.py
前置：server.py 已在运行（端口 8765）
"""
import subprocess, re, time, sys, os

DIR = os.path.dirname(os.path.abspath(__file__))
cf = os.path.join(DIR, 'cloudflared.exe')

if not os.path.exists(cf):
    print('缺少 cloudflared.exe，请确认与本脚本同目录')
    sys.exit(1)

print('正在启动 cloudflared 公网隧道...')
print('（隧道 URL 每次重启会变化，本地数据保持不变）\n')

proc = subprocess.Popen(
    [cf, 'tunnel', '--url', 'http://127.0.0.1:8765', '--no-autoupdate'],
    stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1
)

url_pattern = re.compile(r'https://[a-z0-9-]+\.trycloudflare\.com')
printed = False
try:
    for line in proc.stdout:
        sys.stdout.write(line)
        sys.stdout.flush()
        if not printed:
            m = url_pattern.search(line)
            if m:
                printed = True
                url = m.group(0)
                print('\n' + '=' * 60)
                print('  公网访问地址（可分享给任何人）：')
                print('  ' + url)
                print('=' * 60 + '\n')
                print('保持此窗口与 server.py 窗口不关闭，公网即可持续访问。')
                print('按 Ctrl+C 停止隧道。\n')
except KeyboardInterrupt:
    proc.terminate()
    print('\n隧道已停止。')
