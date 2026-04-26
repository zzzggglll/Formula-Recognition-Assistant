# 公式识别助手 MVP

这个仓库现在面向 `Vercel 界面演示` 做了轻量化处理：默认使用 `demo` 模式，不再把 `PaddleOCR / PaddleX / paddlepaddle` 这类超大依赖打进部署包里。这样可以把网页界面、上传交互、LaTeX 展示与复制能力顺利部署到 Vercel。

如果你需要本地完整 OCR，请安装 `requirements.ocr.txt`，并把环境变量 `FORMULA_RECOGNIZER_MODE` 切回 `paddleocr`。

## 当前能力

- 单张公式图片上传识别
- 拖拽上传 / 粘贴截图
- LaTeX 结果展示
- KaTeX 公式预览
- 一键复制 LaTeX / 复制 `$$` 包裹结果
- 手动修正识别结果
- 提供 Word 插件版页面与接口
- 支持 `demo` 演示模式
- 可选恢复 `PaddleOCR` 本地识别模式

## 这份仓库的部署定位

- `Vercel`：用于在线界面演示，默认 `demo` 模式
- `本地 Windows / 自建服务器`：如需真实 OCR，再额外安装重依赖

当前 Vercel demo 版的取舍：

- 识别结果是演示用样例，不是实际 OCR 推理结果
- Word 导出在 Vercel demo 上默认关闭，避免因为 Office / XSL 依赖缺失导致线上报错
- 页面界面、上传交互、LaTeX 编辑、复制和预览都可以正常演示

## 目录结构

```text
公式识别git/
├─ app.py
├─ serve_public.py
├─ wsgi.py
├─ requirements.txt
├─ requirements.ocr.txt
├─ .env.example
├─ services/
├─ public/
├─ templates/
├─ static/
└─ deploy/
```

## 部署到 Vercel

把这个仓库直接连接到 Vercel 即可。当前版本已经做了这些处理：

- `requirements.txt` 只保留轻量依赖
- Vercel 上默认走 `demo` 模式
- 提供 `vercel.json` 和 `.vercelignore`
- 主页面静态资源放在 `public/`

如果你之前遇到 `Total bundle size exceeds Lambda ephemeral storage limit (500 MB)`，这份仓库就是为了解决这个问题。

## 本地运行 demo 版

```powershell
cd E:\公式识别git
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
Copy-Item .env.example .env
python app.py
```

浏览器访问：`http://127.0.0.1:5050`

## 本地恢复完整 OCR

```powershell
cd E:\公式识别git
.\.venv\Scripts\activate
pip install -r requirements.ocr.txt
```

然后把 `.env` 改成：

```env
FORMULA_RECOGNIZER_MODE=paddleocr
```

这样就会重新启用 `PaddleOCR FormulaRecognitionPipeline`。

## 关键环境变量

可直接参考 [.env.example](/e:/公式识别git/.env.example)：

- `FORMULA_RECOGNIZER_MODE=demo`：当前仓库默认用来做 Vercel 演示
- `FORMULA_RECOGNIZER_MODE=paddleocr`：切回本地 PaddleOCR 识别
- `FORMULA_WORD_EXPORT_ENABLED=false`：Vercel demo 可显式关闭 Word 导出
- `WORD_COM_AUTOMATION_ENABLED=false`：公网部署默认建议关闭

## 健康检查

服务启动后可访问：

- 首页：`/`
- 健康检查：`/api/health`
- 运行时信息：`/api/runtime-info`
