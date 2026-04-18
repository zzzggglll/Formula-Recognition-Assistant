import re


_TEXTUAL_COMMANDS = (
    'mathrm',
    'operatorname',
    'textrm',
    'text',
    'mbox',
)
_STYLE_COMMANDS = (
    'mathbf',
    'boldsymbol',
    'bm',
    'mathit',
    'mathsf',
    'mathtt',
    'mathcal',
    'mathbb',
    'mathfrak',
    'mathscr',
)
_TEXTUAL_PATTERN = re.compile(
    r'\\(?P<command>' + '|'.join(_TEXTUAL_COMMANDS) + r')\s*\{(?P<body>[^{}]+)\}'
)
_STYLE_PATTERN = re.compile(
    r'\\(?P<command>' + '|'.join(_STYLE_COMMANDS) + r')\s*\{(?P<body>[^{}]+)\}'
)
_SPACED_LETTERS_PATTERN = re.compile(r'(?<!\\)(?:[A-Za-z0-9]\s+){2,}[A-Za-z0-9]')
_PRIME_PATTERN = re.compile(r"\^\s*\{\s*((?:\\prime|')+)\s*\}")


def cleanup_recognized_latex(latex: str) -> str:
    value = (latex or '').strip()
    value = _normalize_newlines(value)
    value = _collapse_textual_spacing(value)
    value = _normalize_prime_groups(value)
    return value.strip()


def prepare_latex_for_word(latex: str) -> str:
    value = cleanup_recognized_latex(latex)
    value = _collapse_textual_spacing(value, word_mode=True)
    value = _strip_style_commands(value)
    value = _normalize_prime_groups(value)
    return value.strip()


def _normalize_newlines(value: str) -> str:
    return value.replace('\r\n', '\n').replace('\r', '\n')


def _collapse_textual_spacing(latex: str, word_mode: bool = False) -> str:
    previous = None
    current = latex
    while current != previous:
        previous = current
        current = _TEXTUAL_PATTERN.sub(
            lambda match: _rewrite_textual_command(match.group('command'), match.group('body'), word_mode),
            current,
        )
    return current


def _rewrite_textual_command(command: str, body: str, word_mode: bool) -> str:
    collapsed = _collapse_spaced_letters(body)
    normalized_body = collapsed if collapsed is not None else body.strip()

    if word_mode and _looks_like_plain_text(normalized_body):
        if command == 'operatorname':
            return r'\mathrm{' + normalized_body + '}'
        if command in {'textrm', 'mbox'}:
            return r'\text{' + normalized_body + '}'

    return '\\' + command + '{' + normalized_body + '}'


def _collapse_spaced_letters(body: str):
    if any(token in body for token in ('\\', '{', '}')):
        return None

    tokens = body.split()
    if len(tokens) < 2:
        return body.strip()

    if all(len(token) == 1 and token.isalnum() for token in tokens):
        return ''.join(tokens)

    if _SPACED_LETTERS_PATTERN.fullmatch(body.strip()):
        return re.sub(r'\s+', '', body.strip())

    return body.strip()


def _looks_like_plain_text(body: str) -> bool:
    if len(body) < 2:
        return False
    return bool(re.fullmatch(r'[A-Za-z0-9 .,:;!?+\-]+', body))


def _strip_style_commands(latex: str) -> str:
    previous = None
    current = latex
    while current != previous:
        previous = current
        current = _STYLE_PATTERN.sub(lambda match: match.group('body').strip(), current)
    return current


def _normalize_prime_groups(latex: str) -> str:
    return _PRIME_PATTERN.sub(lambda match: '^{' + _normalize_prime_body(match.group(1)) + '}', latex)


def _normalize_prime_body(body: str) -> str:
    normalized = re.sub(r'\s+', '', body)
    probe = normalized.replace('\\prime', '')
    if probe and set(probe) <= {"'"}:
        prime_count = normalized.count('\\prime') + probe.count("'")
        return '\\prime' * prime_count
    if not probe and '\\prime' in normalized:
        prime_count = normalized.count('\\prime')
        return '\\prime' * prime_count
    return normalized or body
