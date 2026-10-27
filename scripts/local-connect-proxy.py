"""本地 CONNECT 代理：只为被 hosts 劫持的 GitHub 域名做「域名 → 真实 IP」转发。

背景：这台机器的 hosts 把 github.com / api.github.com 指向 127.0.0.1，
导致 gh / git 无法连接；但 GitHub 真实 IP 的 443 是通的。
本代理接收 CONNECT 请求后，直接连真实 IP，TLS 握手仍使用原域名（SNI 与证书校验都正常），
因此 git 和 gh 都能安全地走它，不需要改 hosts，也不需要管理员权限。

用法：
    python scripts/local-connect-proxy.py --port 8899
    # 另开终端
    $env:HTTPS_PROXY = "http://127.0.0.1:8899"
    gh auth status
    git ls-remote https://github.com/OWNER/REPO.git
"""
from __future__ import annotations

import argparse
import select
import socket
import socketserver
import sys
import threading

# 实测可达的 GitHub 边缘 IP（2026-10）
UPSTREAM = {
    "github.com": "140.82.112.3",
    "api.github.com": "140.82.112.6",
    "codeload.github.com": "140.82.112.9",
    "uploads.github.com": "140.82.112.6",
    "objects.githubusercontent.com": "185.199.108.133",
    "raw.githubusercontent.com": "185.199.108.133",
    "gist.github.com": "140.82.112.3",
}


def connect_upstream(host: str, port: int) -> socket.socket:
    target = UPSTREAM.get(host)
    if target:
        return socket.create_connection((target, port), timeout=20)
    # 非 GitHub 域名走正常 DNS
    return socket.create_connection((host, port), timeout=20)


def pump(a: socket.socket, b: socket.socket) -> None:
    try:
        while True:
            r, _, _ = select.select([a, b], [], [], 60)
            if not r:
                break
            for sock in r:
                data = sock.recv(65536)
                if not data:
                    return
                (b if sock is a else a).sendall(data)
    except OSError:
        pass


class Handler(socketserver.BaseRequestHandler):
    def handle(self) -> None:
        client: socket.socket = self.request
        try:
            head = b""
            while b"\r\n\r\n" not in head:
                chunk = client.recv(4096)
                if not chunk:
                    return
                head += chunk
            line = head.split(b"\r\n", 1)[0].decode("latin-1")
            method, target, _ = line.split(" ", 2)
            if method.upper() != "CONNECT":
                client.sendall(b"HTTP/1.1 501 Not Implemented\r\n\r\n")
                return
            host, _, port_s = target.rpartition(":")
            port = int(port_s or 443)
            upstream = connect_upstream(host, port)
            client.sendall(b"HTTP/1.1 200 Connection Established\r\n\r\n")
            pump(client, upstream)
            upstream.close()
        except Exception as exc:  # 代理失败不应拖垮进程
            print(f"  ! {exc}", file=sys.stderr)
        finally:
            try:
                client.close()
            except OSError:
                pass


class Server(socketserver.ThreadingTCPServer):
    daemon_threads = True
    allow_reuse_address = True


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8899)
    ap.add_argument("--host", default="127.0.0.1")
    args = ap.parse_args()
    with Server((args.host, args.port), Handler) as srv:
        print(f"本地 CONNECT 代理已启动：http://{args.host}:{args.port}")
        print("用法：$env:HTTPS_PROXY = \"http://%s:%d\"" % (args.host, args.port))
        try:
            srv.serve_forever()
        except KeyboardInterrupt:
            print("\n已停止")
    return 0


if __name__ == "__main__":
    threading.current_thread().name = "proxy-main"
    sys.exit(main())
