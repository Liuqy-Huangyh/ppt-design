"""PPT Studio 服务端包：HTTP 服务 + 生成编排 + 档案/反馈落盘。"""

from __future__ import annotations

__version__ = "1.0.0"

__all__ = ["create_server", "__version__"]


def create_server(host: str = "127.0.0.1", port: int = 8765, store_dir=None):
    """延迟导入，避免 `import server` 时拉起整个服务端依赖。"""
    from .app import create_server as _create_server

    return _create_server(host, port, store_dir)
