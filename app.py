import os
from io import BytesIO
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from flask import Flask, jsonify, redirect, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.middleware.proxy_fix import ProxyFix

from services.formula_recognizer import FormulaRecognitionError, recognize_formula
from services.word_math import WordMathError, build_word_docx, prepare_word_payload


BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / '.env')
shared_env = BASE_DIR.parent / 'ai学习助手' / '.env'
if shared_env.exists():
    load_dotenv(shared_env, override=False)


app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = int(os.getenv('MAX_UPLOAD_MB', '5')) * 1024 * 1024
app.json.ensure_ascii = False

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'webp'}
ALLOWED_MIME_TYPES = {'image/png', 'image/jpeg', 'image/webp'}


def _allowed_image(filename: str, mimetype: Optional[str]) -> bool:
    if not filename or '.' not in filename:
        return False
    ext = filename.rsplit('.', 1)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        return False
    if mimetype and mimetype not in ALLOWED_MIME_TYPES:
        return False
    return True


def _parse_bool_env(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {'1', 'true', 'yes', 'on'}


if _parse_bool_env('TRUST_PROXY', default=False):
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1, x_prefix=1)


def _resolve_ssl_context():
    cert_file = os.getenv('FORMULA_OCR_SSL_CERT_FILE', '').strip() or os.getenv('SSL_CERT_FILE', '').strip()
    key_file = os.getenv('FORMULA_OCR_SSL_KEY_FILE', '').strip() or os.getenv('SSL_KEY_FILE', '').strip()

    if cert_file or key_file:
        if not cert_file or not key_file:
            raise RuntimeError('SSL_CERT_FILE 和 SSL_KEY_FILE 需要同时配置。')

        cert_path = Path(cert_file)
        key_path = Path(key_file)
        if not cert_path.is_absolute():
            cert_path = (BASE_DIR / cert_path).resolve()
        if not key_path.is_absolute():
            key_path = (BASE_DIR / key_path).resolve()

        if not cert_path.exists():
            raise RuntimeError(f'未找到 SSL 证书文件：{cert_path}')
        if not key_path.exists():
            raise RuntimeError(f'未找到 SSL 私钥文件：{key_path}')

        return (str(cert_path), str(key_path))

    if _parse_bool_env('FLASK_SSL_ADHOC', default=False):
        return 'adhoc'

    return None


def _recognize_upload(upload):
    if upload is None:
        raise FormulaRecognitionError('请先上传一张公式图片。')

    if not upload.filename:
        raise FormulaRecognitionError('未检测到文件名，请重新选择图片。')

    if not _allowed_image(upload.filename, upload.mimetype):
        raise FormulaRecognitionError('仅支持 png、jpg、jpeg、webp 图片。')

    image_bytes = upload.read()
    if not image_bytes:
        raise FormulaRecognitionError('图片内容为空，请重新上传。')

    result = recognize_formula(
        image_bytes=image_bytes,
        filename=upload.filename,
        mime_type=upload.mimetype,
    )
    return upload.filename, result


@app.get('/')
def index():
    return render_template('index.html')


@app.get('/word-addin')
def word_addin_home():
    return redirect('/word-addin/taskpane')


@app.get('/word-addin/taskpane')
def word_addin_taskpane():
    return render_template('word_addin_taskpane.html')


@app.get('/word-addin/commands')
def word_addin_commands():
    return render_template('word_addin_commands.html')


@app.get('/api/health')
def health():
    return jsonify(ok=True, service='formula-ocr-mvp')


@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(_exc):
    message = '图片超过大小限制，请上传 5MB 以内的公式截图。'
    if request.path.startswith('/api/'):
        return jsonify(ok=False, error=message), 413
    return message, 413


@app.post('/api/recognize')
def recognize():
    upload = request.files.get('file')

    try:
        filename, result = _recognize_upload(upload)
    except FormulaRecognitionError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    except Exception as exc:
        return jsonify(ok=False, error=f'识别服务异常：{exc}'), 500

    return jsonify(
        ok=True,
        latex=result['latex'],
        wrapped_latex=result['wrapped_latex'],
        provider=result['provider'],
        mode=result['mode'],
        notice=result['notice'],
        filename=filename,
    )


@app.post('/api/word/prepare')
def prepare_word():
    payload = request.get_json(silent=True) or {}
    latex = str(payload.get('latex', ''))

    try:
        result = prepare_word_payload(latex, normalize_word_ooxml=False)
    except WordMathError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    except Exception as exc:
        return jsonify(ok=False, error=f'Word 公式转换异常：{exc}'), 500

    return jsonify(
        ok=True,
        latex=result['latex'],
        word_latex=result['word_latex'],
        mathml=result['mathml'],
        omml_xml=result['omml_xml'],
        word_ooxml=result['word_ooxml'],
        clipboard_html=result['clipboard_html'],
        hint=result['hint'],
    )


@app.post('/api/word/export')
def export_word():
    payload = request.get_json(silent=True) or {}
    latex = str(payload.get('latex', ''))
    source_filename = payload.get('filename')

    try:
        docx_bytes, download_name = build_word_docx(
            latex,
            source_filename=source_filename,
            prepared_payload=payload,
        )
    except WordMathError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    except Exception as exc:
        return jsonify(ok=False, error=f'Word 文档导出异常：{exc}'), 500

    return send_file(
        BytesIO(docx_bytes),
        as_attachment=True,
        download_name=download_name,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


@app.post('/api/word-addin/recognize')
def word_addin_recognize():
    upload = request.files.get('file')

    try:
        filename, result = _recognize_upload(upload)
        word_payload = prepare_word_payload(result['latex'], normalize_word_ooxml=True)
    except FormulaRecognitionError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    except WordMathError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    except Exception as exc:
        return jsonify(ok=False, error=f'Word 插件识别异常：{exc}'), 500

    return jsonify(
        ok=True,
        filename=filename,
        latex=result['latex'],
        wrapped_latex=result['wrapped_latex'],
        provider=result['provider'],
        mode=result['mode'],
        notice=result['notice'],
        word_hint=word_payload['hint'],
        word_latex=word_payload['word_latex'],
        omml_xml=word_payload['omml_xml'],
        word_ooxml=word_payload['word_ooxml'],
    )


if __name__ == '__main__':
    host = os.getenv('HOST', '0.0.0.0').strip() or '0.0.0.0'
    port = int(os.getenv('PORT', '5050'))
    debug = _parse_bool_env('FLASK_DEBUG', default=False)
    ssl_context = _resolve_ssl_context()
    app.run(host=host, port=port, debug=debug, use_reloader=debug, ssl_context=ssl_context)


