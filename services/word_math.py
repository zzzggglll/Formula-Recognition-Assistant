import io
import os
import re
import subprocess
import tempfile
import threading
from pathlib import Path
from typing import Optional, Tuple

from .latex_cleanup import prepare_latex_for_word


MML2OMML_XSL_ENV = 'WORD_MML2OMML_XSL_PATH'
DEFAULT_MML2OMML_XSL_PATHS = (
    Path(r'C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL'),
    Path(r'C:\Program Files\Microsoft Office\root\Office15\MML2OMML.XSL'),
    Path(r'C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL'),
    Path(r'C:\Program Files (x86)\Microsoft Office\root\Office15\MML2OMML.XSL'),
)
NARY_COMMANDS = (
    'sum',
    'prod',
    'coprod',
    'bigcup',
    'bigcap',
    'bigvee',
    'bigwedge',
    'int',
    'iint',
    'iiint',
    'oint',
)
NARY_OPERATOR_PATTERN = re.compile(r'\\(?:' + '|'.join(NARY_COMMANDS) + r')(?![A-Za-z])')
_RELATION_OR_SEPARATOR_ATOM = {'+', '-', '=', ',', ';', '&'}
_TRANSFORM = None
_TRANSFORM_LOCK = threading.Lock()
_WORD_COM_LOCK = threading.Lock()
WORD_COM_ENABLED_ENV = 'WORD_COM_AUTOMATION_ENABLED'
OMML_ASCII_FALLBACKS = {
    '\u2032': "'",
    '\u2033': "''",
    '\u2034': "'''",
    '\u2057': "''''",
}


class WordMathError(RuntimeError):
    pass


def prepare_word_payload(latex: str, normalize_word_ooxml: bool = False) -> dict[str, str]:
    normalized_latex = _normalize_latex(latex)
    word_latex = _normalize_latex_for_word(normalized_latex)
    mathml = _latex_to_mathml(word_latex)
    omml_xml = _normalize_omml_for_word(_mathml_to_omml(mathml))
    word_ooxml = _build_word_insert_ooxml(omml_xml)
    if normalize_word_ooxml:
        word_ooxml = _normalize_word_ooxml_with_local_word(word_ooxml) or word_ooxml
    return {
        'latex': normalized_latex,
        'word_latex': word_latex,
        'mathml': mathml,
        'omml_xml': omml_xml,
        'word_ooxml': word_ooxml,
        'clipboard_html': _build_clipboard_html(omml_xml),
        'hint': '复制按钮现在会复制更适合粘贴进 Word 公式框的线性公式。建议先在 Word 中按 Alt+= 插入公式框，再粘贴；如果你想要最稳妥的原生格式，直接下载 Word 文档即可。',
    }

def build_word_docx(
    latex: str,
    source_filename: Optional[str] = None,
    prepared_payload: Optional[dict[str, str]] = None,
) -> Tuple[bytes, str]:
    payload = _resolve_export_payload(latex, prepared_payload)
    download_name = _build_docx_filename(source_filename)

    native_docx_bytes = _build_native_docx_with_local_word(payload['word_ooxml'])
    if native_docx_bytes is not None:
        return native_docx_bytes, download_name

    try:
        from docx import Document
        from lxml import etree
    except ImportError as exc:
        raise WordMathError('未安装 python-docx 或 lxml，暂时无法导出 Word 文档。') from exc

    doc = Document()
    paragraph = doc.add_paragraph()
    paragraph._element.append(etree.fromstring(payload['omml_xml'].encode('utf-8')))

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue(), download_name


def _resolve_export_payload(latex: str, prepared_payload: Optional[dict[str, str]] = None) -> dict[str, str]:
    normalized_latex = _normalize_latex(latex)
    if prepared_payload:
        prepared_latex = str(prepared_payload.get('latex', ''))
        prepared_omml_xml = str(prepared_payload.get('omml_xml', ''))
        prepared_word_ooxml = str(prepared_payload.get('word_ooxml', ''))
        if prepared_latex and prepared_omml_xml and prepared_word_ooxml:
            try:
                if _normalize_latex(prepared_latex) == normalized_latex:
                    return {
                        'latex': normalized_latex,
                        'omml_xml': prepared_omml_xml,
                        'word_ooxml': prepared_word_ooxml,
                    }
            except WordMathError:
                pass

    return prepare_word_payload(normalized_latex, normalize_word_ooxml=False)

def _normalize_latex(latex: str) -> str:
    value = (latex or '').strip()
    value = value.strip('`')
    if not value:
        raise WordMathError('当前没有可转换的 LaTeX 公式。')

    wrappers = (
        ('$$', '$$'),
        ('$', '$'),
        (r'\[', r'\]'),
        (r'\(', r'\)'),
    )
    for start, end in wrappers:
        if value.startswith(start) and value.endswith(end):
            value = value[len(start):-len(end)].strip()

    equation_match = re.fullmatch(
        r'\\begin\{(equation\*?|displaymath)\}(.*)\\end\{\1\}',
        value,
        flags=re.DOTALL,
    )
    if equation_match:
        value = equation_match.group(2).strip()

    if not value:
        raise WordMathError('公式内容为空，请先识别或输入 LaTeX。')
    return value


def _normalize_latex_for_word(latex: str) -> str:
    value = prepare_latex_for_word(latex)
    value = _wrap_nary_operands(value)
    return value

def _wrap_nary_operands(latex: str) -> str:
    result: list[str] = []
    cursor = 0

    while True:
        match = NARY_OPERATOR_PATTERN.search(latex, cursor)
        if match is None:
            result.append(latex[cursor:])
            break

        start = match.start()
        result.append(latex[cursor:start])

        prefix_end = _consume_nary_prefix(latex, start)
        result.append(latex[start:prefix_end])

        operand_start = _skip_spaces(latex, prefix_end)
        result.append(latex[prefix_end:operand_start])

        if operand_start >= len(latex) or latex[operand_start] == '{':
            cursor = operand_start
            continue

        operand_end = _find_nary_operand_end(latex, operand_start)
        trimmed_end = operand_end
        while trimmed_end > operand_start and latex[trimmed_end - 1].isspace():
            trimmed_end -= 1

        operand = latex[operand_start:trimmed_end]
        if not operand:
            cursor = operand_end
            continue

        result.append('{')
        result.append(operand)
        result.append('}')
        result.append(latex[trimmed_end:operand_end])
        cursor = operand_end

    return ''.join(result)


def _consume_nary_prefix(latex: str, start: int) -> int:
    position = start
    match = NARY_OPERATOR_PATTERN.match(latex, position)
    if not match:
        return start
    position = match.end()

    while position < len(latex):
        next_position = _skip_spaces(latex, position)
        if latex.startswith('\\limits', next_position):
            position = next_position + len('\\limits')
            continue
        if latex.startswith('\\nolimits', next_position):
            position = next_position + len('\\nolimits')
            continue
        if next_position < len(latex) and latex[next_position] in {'_', '^'}:
            position = _consume_latex_atom(latex, next_position + 1)
            continue
        break

    return position


def _consume_latex_atom(latex: str, start: int) -> int:
    position = _skip_spaces(latex, start)
    if position >= len(latex):
        return position

    opening = latex[position]
    if opening == '{':
        return _consume_balanced(latex, position, '{', '}')
    if opening == '(':
        return _consume_balanced(latex, position, '(', ')')
    if opening == '[':
        return _consume_balanced(latex, position, '[', ']')
    if opening == '\\':
        command_match = re.match(r'\\[A-Za-z]+|\\.', latex[position:])
        if command_match:
            return position + command_match.end()
    return position + 1


def _consume_balanced(latex: str, start: int, opening: str, closing: str) -> int:
    depth = 0
    position = start
    while position < len(latex):
        char = latex[position]
        if char == '\\':
            position += 2
            continue
        if char == opening:
            depth += 1
        elif char == closing:
            depth -= 1
            if depth == 0:
                return position + 1
        position += 1
    return len(latex)


def _find_nary_operand_end(latex: str, start: int) -> int:
    brace_depth = 0
    paren_depth = 0
    bracket_depth = 0
    position = start

    while position < len(latex):
        char = latex[position]

        if char == '\\':
            command_match = re.match(r'\\[A-Za-z]+|\\.', latex[position:])
            if command_match:
                command = command_match.group(0)
                if (
                    brace_depth == 0
                    and paren_depth == 0
                    and bracket_depth == 0
                    and command[1:] in NARY_COMMANDS
                    and position > start
                ):
                    break
                position += len(command)
                continue

        if char == '{':
            brace_depth += 1
        elif char == '}':
            if brace_depth == 0:
                break
            brace_depth -= 1
        elif char == '(':
            paren_depth += 1
        elif char == ')':
            if paren_depth == 0:
                break
            paren_depth -= 1
        elif char == '[':
            bracket_depth += 1
        elif char == ']':
            if bracket_depth == 0:
                break
            bracket_depth -= 1
        elif brace_depth == 0 and paren_depth == 0 and bracket_depth == 0 and char in _RELATION_OR_SEPARATOR_ATOM:
            break

        position += 1

    return position


def _skip_spaces(latex: str, start: int) -> int:
    position = start
    while position < len(latex) and latex[position].isspace():
        position += 1
    return position


def _latex_to_mathml(latex: str) -> str:
    try:
        from latex2mathml.converter import convert
    except ImportError as exc:
        raise WordMathError('未安装 latex2mathml，暂时无法生成 Word 公式。') from exc

    try:
        return convert(latex)
    except Exception as exc:
        raise WordMathError(f'LaTeX 转 MathML 失败：{exc}') from exc


def _mathml_to_omml(mathml: str) -> str:
    try:
        from lxml import etree
    except ImportError as exc:
        raise WordMathError('未安装 lxml，暂时无法生成 Word 公式。') from exc

    try:
        transform = _get_transform()
        mathml_root = etree.fromstring(mathml.encode('utf-8'))
        omml_tree = transform(mathml_root)
        return etree.tostring(omml_tree.getroot(), encoding='utf-8').decode('utf-8')
    except WordMathError:
        raise
    except Exception as exc:
        raise WordMathError(f'MathML 转 Word OMML 失败：{exc}') from exc


def _get_transform():
    global _TRANSFORM
    if _TRANSFORM is not None:
        return _TRANSFORM

    with _TRANSFORM_LOCK:
        if _TRANSFORM is not None:
            return _TRANSFORM

        try:
            from lxml import etree
        except ImportError as exc:
            raise WordMathError('未安装 lxml，暂时无法生成 Word 公式。') from exc

        xslt_path = _resolve_mml2omml_xsl_path()
        try:
            _TRANSFORM = etree.XSLT(etree.parse(str(xslt_path)))
        except Exception as exc:
            raise WordMathError(f'加载 Word 公式转换模板失败：{exc}') from exc
        return _TRANSFORM


def _resolve_mml2omml_xsl_path() -> Path:
    configured = os.getenv(MML2OMML_XSL_ENV, '').strip()
    candidates = []
    if configured:
        candidates.append(Path(configured))
    candidates.extend(DEFAULT_MML2OMML_XSL_PATHS)

    for candidate in candidates:
        if candidate.exists():
            return candidate

    raise WordMathError(
        '未找到 MML2OMML.XSL。请确认本机已安装 Microsoft Office，或在 .env 中配置 WORD_MML2OMML_XSL_PATH。'
    )


def _normalize_omml_for_word(omml_xml: str) -> str:
    return omml_xml


def _build_clipboard_html(omml_xml: str) -> str:
    return (
        '<html xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:w="urn:schemas-microsoft-com:office:word" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        '<head><meta charset="utf-8"></head>'
        '<body><!--StartFragment-->'
        '<p class="MsoNormal">'
        f'{omml_xml}'
        '</p>'
        '<!--EndFragment--></body>'
        '</html>'
    )


def _build_word_insert_ooxml(omml_xml: str) -> str:
    return (
        '<?xml version="1.0" standalone="yes"?>'
        '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
        '<pkg:part pkg:name="/_rels/.rels" '
        'pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" '
        'pkg:padding="512">'
        '<pkg:xmlData>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        '</Relationships>'
        '</pkg:xmlData>'
        '</pkg:part>'
        '<pkg:part pkg:name="/word/document.xml" '
        'pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
        '<pkg:xmlData>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        '<w:body>'
        '<w:p>'
        f'{omml_xml}'
        '</w:p>'
        '</w:body>'
        '</w:document>'
        '</pkg:xmlData>'
        '</pkg:part>'
        '</pkg:package>'
    )



def _normalize_word_ooxml_with_local_word(word_ooxml: str) -> Optional[str]:
    if not _should_use_local_word_automation():
        return None

    script = r'''param([string]$InputPath, [string]$OutputPath)
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {
    $xml = Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $range = $doc.Range(0, 0)
    $range.InsertXML($xml)
    if ($doc.OMaths.Count -lt 1) {
        throw 'Word failed to recognize the inserted equation.'
    }
    $normalized = $doc.OMaths.Item(1).Range.WordOpenXML
    [System.IO.File]::WriteAllText($OutputPath, $normalized, [System.Text.UTF8Encoding]::new($false))
}
finally {
    if ($doc -ne $null) { $doc.Close([ref]$false) }
    if ($word -ne $null) { $word.Quit() }
}
'''
    return _run_local_word_automation(script, word_ooxml, read_mode='text')


def _build_native_docx_with_local_word(word_ooxml: str) -> Optional[bytes]:
    if not _should_use_local_word_automation():
        return None

    script = r'''param([string]$InputPath, [string]$OutputPath)
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
try {
    $xml = Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $range = $doc.Range(0, 0)
    $range.InsertXML($xml)
    $doc.SaveAs2($OutputPath, 16)
}
finally {
    if ($doc -ne $null) { $doc.Close([ref]$false) }
    if ($word -ne $null) { $word.Quit() }
}
'''
    return _run_local_word_automation(script, word_ooxml, read_mode='bytes', output_name='normalized.docx')


def _should_use_local_word_automation() -> bool:
    if os.name != 'nt':
        return False

    value = os.getenv(WORD_COM_ENABLED_ENV)
    if value is None:
        return True
    return value.strip().lower() in {'1', 'true', 'yes', 'on'}


def _run_local_word_automation(script: str, input_text: str, *, read_mode: str, output_name: str = 'normalized.xml'):
    with _WORD_COM_LOCK:
        with tempfile.TemporaryDirectory(prefix='formula-word-') as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / 'input.xml'
            output_path = temp_path / output_name
            script_path = temp_path / 'run_word.ps1'

            input_path.write_text(input_text, encoding='utf-8')
            script_path.write_text(script, encoding='utf-8')

            command = [
                'powershell',
                '-NoProfile',
                '-ExecutionPolicy', 'Bypass',
                '-File', str(script_path),
                str(input_path),
                str(output_path),
            ]

            try:
                completed = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    timeout=120,
                    check=False,
                )
            except Exception:
                return None

            if completed.returncode != 0 or not output_path.exists():
                return None

            if read_mode == 'bytes':
                return output_path.read_bytes()
            return output_path.read_text(encoding='utf-8')


def _build_docx_filename(source_filename: Optional[str]) -> str:
    stem = 'formula'
    if source_filename:
        source_path = Path(source_filename)
        if source_path.stem:
            stem = source_path.stem

    safe_stem = re.sub(r'[^A-Za-z0-9_-]+', '-', stem).strip('-') or 'formula'
    return f'{safe_stem}-word-equation.docx'



