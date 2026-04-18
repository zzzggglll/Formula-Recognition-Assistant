import os

from waitress import serve

from app import app


def _int_env(name: str, default: int) -> int:
    value = os.getenv(name)
    if value is None:
        return default
    try:
        return int(value)
    except ValueError:
        return default


if __name__ == '__main__':
    host = os.getenv('HOST', '0.0.0.0').strip() or '0.0.0.0'
    port = _int_env('PORT', 5050)
    threads = _int_env('WAITRESS_THREADS', 8)
    connection_limit = _int_env('WAITRESS_CONNECTION_LIMIT', 100)
    serve(app, host=host, port=port, threads=threads, connection_limit=connection_limit)
