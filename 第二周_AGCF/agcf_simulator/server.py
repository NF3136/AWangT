"""本机 Web 界面；所有计算离线完成，绝不向 SIP 地址发包。"""
import argparse
import json
import os
import socket
import uuid
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
from engine import load_json, simulate, explore, sample

BASE = Path(__file__).resolve().parent


class LocalServer(ThreadingHTTPServer):
    allow_reuse_address = False
    daemon_threads = True

    def server_bind(self):
        if os.name == 'nt':
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


class Handler(BaseHTTPRequestHandler):
    def reply(self, status, content, content_type='application/json; charset=utf-8'):
        body = content if isinstance(content, bytes) else json.dumps(content, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; frame-ancestors 'none'")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urlparse(self.path).path
        if path == '/api/health':
            return self.reply(200, {'app': 'agcf-offline-lab', 'pid': os.getpid(), 'workspace': str(BASE)})
        if path == '/api/bootstrap':
            return self.reply(200, {'snapshot': load_json('real_config.json'), 'profiles': load_json('demo_profiles.json'),
                                    'sample': sample(), 'app': 'agcf-offline-lab'})
        files = {'/': ('index.html', 'text/html; charset=utf-8'), '/app.js': ('app.js', 'text/javascript; charset=utf-8'),
                 '/style.css': ('style.css', 'text/css; charset=utf-8')}
        if path not in files:
            return self.reply(404, {'error': '未找到资源'})
        name, mime = files[path]
        self.reply(200, (BASE / 'web' / name).read_bytes(), mime)

    def do_POST(self):
        try:
            origin = self.headers.get('Origin')
            if origin and origin != f'http://{self.headers.get("Host")}':
                return self.reply(403, {'error': '仅接受本地页面的同源请求'})
            size = int(self.headers.get('Content-Length', '0'))
            if not 0 < size <= 300000:
                return self.reply(413, {'error': '请求体大小不受支持'})
            body = json.loads(self.rfile.read(size).decode('utf-8'))
            if not isinstance(body, dict):
                raise ValueError('请求必须是 JSON 对象')
            path = urlparse(self.path).path
            if path == '/api/simulate':
                return self.reply(200, simulate(body))
            if path == '/api/explore':
                return self.reply(200, explore(body.get('number')))
            if path == '/api/export':
                if body.get('mode') != 'registered_demo' or not isinstance(body.get('steps'), list):
                    return self.reply(400, {'error': '只能导出本程序的模拟结果'})
                folder = BASE / 'exports'
                folder.mkdir(exist_ok=True)
                target = folder / f'AGCF-{datetime.now().strftime("%Y%m%d-%H%M%S")}-{uuid.uuid4().hex[:8]}.json'
                target.write_text(json.dumps(body, ensure_ascii=False, indent=2), encoding='utf-8')
                return self.reply(200, {'path': str(target)})
            return self.reply(404, {'error': '未找到接口'})
        except (ValueError, UnicodeError):
            return self.reply(400, {'error': '请求内容不是有效 JSON'})

    def log_message(self, format, *args):
        # 不记录报文、用户身份或参数正文。
        pass


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--port', type=int, default=18766)
    args = parser.parse_args()
    load_json('real_config.json')
    server = LocalServer(('127.0.0.1', args.port), Handler)
    print(f'AGCF 离线模拟：http://127.0.0.1:{args.port}', flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
