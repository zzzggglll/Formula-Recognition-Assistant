import hashlib
import os
import tempfile
import threading
from typing import Any, Optional

from .latex_cleanup import cleanup_recognized_latex


MODE_ENV = 'FORMULA_RECOGNIZER_MODE'
DEFAULT_MODE = 'paddleocr'
PADDLE_DEVICE_ENV = 'PADDLE_OCR_DEVICE'
PADDLE_FORMULA_MODEL_NAME_ENV = 'PADDLE_OCR_FORMULA_MODEL_NAME'
PADDLE_FORMULA_MODEL_DIR_ENV = 'PADDLE_OCR_FORMULA_MODEL_DIR'
PADDLE_LAYOUT_DETECTION_ENV = 'PADDLE_OCR_USE_LAYOUT_DETECTION'
PADDLE_DOC_ORI_ENV = 'PADDLE_OCR_USE_DOC_ORIENTATION_CLASSIFY'
PADDLE_DOC_UNWARP_ENV = 'PADDLE_OCR_USE_DOC_UNWARPING'
DEFAULT_DEVICE = 'cpu'
DEFAULT_LAYOUT_DETECTION = False
DEFAULT_DOC_ORI = False
DEFAULT_DOC_UNWARP = False

DEMO_FORMULAS = [
    r'\int_0^{\infty} e^{-x^2} \, dx = \frac{\sqrt{\pi}}{2}',
    r'\nabla \cdot \mathbf{E} = \frac{\rho}{\varepsilon_0}',
    r'\hat{\beta} = (X^T X)^{-1} X^T y',
    r'\mathcal{L}(\theta) = -\sum_{i=1}^{N} y_i \log \hat{y}_i',
    r'\frac{d}{dx} \sin x = \cos x',
    r'\sum_{k=1}^{n} k = \frac{n(n+1)}{2}',
]

_PIPELINE = None
_PIPELINE_LOCK = threading.Lock()


class FormulaRecognitionError(RuntimeError):
    pass


def get_recognizer_mode() -> str:
    configured_mode = os.getenv(MODE_ENV)
    if configured_mode and configured_mode.strip():
        return configured_mode.strip().lower()
    if _is_vercel_runtime():
        return 'demo'
    return DEFAULT_MODE


def recognize_formula(image_bytes: bytes, filename: str, mime_type: Optional[str]) -> dict[str, str]:
    mode = get_recognizer_mode()
    if mode == 'demo':
        return _recognize_in_demo_mode(filename)
    if mode not in {'paddleocr', 'demo'}:
        raise FormulaRecognitionError('FORMULA_RECOGNIZER_MODE 仅支持 paddleocr 或 demo。')

    pipeline = _get_pipeline()
    latex = _recognize_with_paddleocr(
        pipeline=pipeline,
        image_bytes=image_bytes,
        filename=filename,
        mime_type=mime_type,
    )
    return {
        'latex': latex,
        'wrapped_latex': _wrap_display_math(latex),
        'provider': 'PaddleOCR',
        'mode': 'FormulaRecognitionPipeline',
        'notice': '当前为 PaddleOCR FormulaRecognitionPipeline 本地识别结果。',
    }


def _get_pipeline():
    global _PIPELINE
    if _PIPELINE is not None:
        return _PIPELINE

    with _PIPELINE_LOCK:
        if _PIPELINE is not None:
            return _PIPELINE

        os.environ.setdefault('PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK', 'True')

        try:
            from paddleocr import FormulaRecognitionPipeline
        except ImportError as exc:
            raise FormulaRecognitionError(
                '未安装 paddleocr / paddlepaddle。请先安装本地依赖后再使用 PaddleOCR 识别。'
            ) from exc

        kwargs: dict[str, Any] = {
            'device': os.getenv(PADDLE_DEVICE_ENV, DEFAULT_DEVICE).strip() or DEFAULT_DEVICE,
        }

        model_name = os.getenv(PADDLE_FORMULA_MODEL_NAME_ENV, '').strip()
        model_dir = os.getenv(PADDLE_FORMULA_MODEL_DIR_ENV, '').strip()
        if model_name:
            kwargs['formula_recognition_model_name'] = model_name
        if model_dir:
            kwargs['formula_recognition_model_dir'] = model_dir

        try:
            _PIPELINE = FormulaRecognitionPipeline(**kwargs)
        except Exception as exc:
            raise FormulaRecognitionError(f'PaddleOCR 初始化失败：{exc}') from exc

        return _PIPELINE


def _recognize_with_paddleocr(*, pipeline: Any, image_bytes: bytes, filename: str, mime_type: Optional[str]) -> str:
    suffix = _guess_suffix(filename, mime_type)
    fd, tmp_path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)

    try:
        with open(tmp_path, 'wb') as file_obj:
            file_obj.write(image_bytes)

        results = pipeline.predict(
            input=tmp_path,
            use_layout_detection=_get_env_bool(PADDLE_LAYOUT_DETECTION_ENV, DEFAULT_LAYOUT_DETECTION),
            use_doc_orientation_classify=_get_env_bool(PADDLE_DOC_ORI_ENV, DEFAULT_DOC_ORI),
            use_doc_unwarping=_get_env_bool(PADDLE_DOC_UNWARP_ENV, DEFAULT_DOC_UNWARP),
        )
    except Exception as exc:
        raise FormulaRecognitionError(f'PaddleOCR 识别失败：{exc}') from exc
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    latex = _extract_latex_from_results(results)
    if not latex:
        raise FormulaRecognitionError('PaddleOCR 已运行完成，但没有解析出 LaTeX。建议先裁剪出更纯净的公式区域再试。')
    return latex


def _extract_latex_from_results(results: Any) -> str:
    if results is None:
        return ''

    if isinstance(results, list):
        iterable = results
    else:
        try:
            iterable = list(results)
        except TypeError:
            iterable = [results]

    formulas: list[str] = []
    for item in iterable:
        formulas.extend(_extract_latex_candidates(item))

    for latex in formulas:
        normalized = _normalize_latex(latex)
        if normalized:
            return normalized
    return ''


def _extract_latex_candidates(item: Any) -> list[str]:
    candidates: list[str] = []

    if item is None:
        return candidates
    if isinstance(item, str):
        return [item]
    if isinstance(item, dict):
        candidates.extend(_extract_from_mapping(item))
        return candidates

    for attr in ('rec_formula', 'formula', 'text', 'latex'):
        value = getattr(item, attr, None)
        if isinstance(value, str) and value.strip():
            candidates.append(value.strip())

    if hasattr(item, 'res'):
        candidates.extend(_extract_latex_candidates(getattr(item, 'res')))
    if hasattr(item, 'data'):
        candidates.extend(_extract_latex_candidates(getattr(item, 'data')))

    if hasattr(item, 'to_dict'):
        try:
            candidates.extend(_extract_latex_candidates(item.to_dict()))
        except Exception:
            pass
    if hasattr(item, 'json'):
        json_value = getattr(item, 'json')
        if isinstance(json_value, dict):
            candidates.extend(_extract_latex_candidates(json_value))

    return candidates


def _extract_from_mapping(data: dict[str, Any]) -> list[str]:
    candidates: list[str] = []

    for key in ('rec_formula', 'formula', 'latex', 'text'):
        value = data.get(key)
        if isinstance(value, str) and value.strip():
            candidates.append(value.strip())

    for key in ('formula_res_list', 'dt_polys', 'rec_texts', 'pruned_result', 'res', 'data'):
        value = data.get(key)
        if isinstance(value, list):
            for entry in value:
                candidates.extend(_extract_latex_candidates(entry))
        elif isinstance(value, dict):
            candidates.extend(_extract_from_mapping(value))
        elif isinstance(value, str) and value.strip() and key in {'rec_texts'}:
            candidates.append(value.strip())

    return candidates


def _normalize_latex(value: str) -> str:
    latex = value.strip()
    latex = latex.replace('\r\n', '\n').replace('\r', '\n')
    latex = latex.strip('`')
    if latex in {'[EMPTY]', '[DOCIMG]'}:
        raise FormulaRecognitionError('这张图更像文档页面而不是单个公式，建议先裁剪出公式局部后再试。')
    return cleanup_recognized_latex(latex.strip())


def _get_env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {'1', 'true', 'yes', 'on'}


def _is_vercel_runtime() -> bool:
    if _get_env_bool('VERCEL', default=False):
        return True
    return bool((os.getenv('VERCEL_ENV') or '').strip())


def _guess_suffix(filename: str, mime_type: Optional[str]) -> str:
    if filename and '.' in filename:
        return '.' + filename.rsplit('.', 1)[1].lower()
    if mime_type == 'image/png':
        return '.png'
    if mime_type == 'image/webp':
        return '.webp'
    return '.jpg'


def _recognize_in_demo_mode(filename: str) -> dict[str, str]:
    index = int(hashlib.md5(filename.encode('utf-8')).hexdigest(), 16) % len(DEMO_FORMULAS)
    latex = DEMO_FORMULAS[index]
    return {
        'latex': latex,
        'wrapped_latex': _wrap_display_math(latex),
        'provider': 'Demo Recognizer',
        'mode': 'demo',
        'notice': '当前为 Demo 模式，仅用于界面演示，不代表真实识别结果；Vercel 部署默认不会执行本地 PaddleOCR 推理。',
    }


def _wrap_display_math(latex: str) -> str:
    return f'$$\n{latex}\n$$'



