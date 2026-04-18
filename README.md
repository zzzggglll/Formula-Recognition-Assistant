# 公式识别助手 MVP

一个基于 Flask 的图片公式识别网站，当前后端默认使用 `PaddleOCR FormulaRecognitionPipeline` 做本地公式识别。现在这套网站已经具备“部署给别人访问”的基础能力，网站版支持公网发布，Word 插件版仍是本地调试形态。

## 当前能力

- 单张公式图片上传识别
- 拖拽上传 / 粘贴截图
- LaTeX 结果展示
- KaTeX 公式预览
- 一键复制 LaTeX / 复制 `$$` 包裹结果
- 手动修正识别结果，并继续用于复制和导出
- 导出包含原生 Word 公式的 `.docx`
- 提供 Word 插件版 MVP 页面与接口
- 支持 `PaddleOCR` 本地识别模式
- 支持 `demo` 演示模式

## 目录结构

```text
公式识别/
├─ app.py
├─ serve_public.py
├─ wsgi.py
├─ requirements.txt
├─ .env
├─ .env.example
├─ services/
│  ├─ formula_recognizer.py
│  └─ word_math.py
├─ templates/
├─ static/
├─ word_addin/
└─ deploy/
```

## 本地开发运行

```powershell
cd E:\公式识别
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
Copy-Item .env.example .env
python app.py
```

浏览器访问：`http://127.0.0.1:5050`

如果你只是本机调试网页，这一套就够了。

## 给别人用的部署方式

如果目标是“让别人通过浏览器访问你的网站”，建议使用下面这套方式：

1. 在服务器或可长期开机的电脑上部署项目。
2. 使用 `python serve_public.py` 启动 Waitress，而不是直接用 Flask 开发服务器。
3. 前面加一层反向代理（推荐 `Caddy` 或 `Nginx`），用 80/443 对外提供访问。
4. 如果用了反向代理，把 `.env` 里的 `TRUST_PROXY=true` 打开。

启动命令：

```powershell
cd E:\公式识别
.\.venv\Scripts\activate
python serve_public.py
```

默认监听：`0.0.0.0:5050`

对外部署详细步骤请看：[网站公网部署指南.md](/e:/公式识别/网站公网部署指南.md)

## 关键环境变量

可直接参考 [.env.example](/e:/公式识别/.env.example)。常用项如下：

- `FORMULA_RECOGNIZER_MODE=paddleocr`：使用本地 PaddleOCR 识别。
- `HOST=0.0.0.0`：监听所有网卡，便于局域网/公网访问。
- `PORT=5050`：应用监听端口。
- `MAX_UPLOAD_MB=5`：单张图片上传大小限制。
- `TRUST_PROXY=false`：如果前面有 Nginx / Caddy，改成 `true`。
- `WAITRESS_THREADS=8`：Waitress 线程数。
- `WAITRESS_CONNECTION_LIMIT=100`：Waitress 连接数上限。
- `WORD_COM_AUTOMATION_ENABLED=false`：公网部署默认建议关闭。

## Word 导出说明

当前网站支持导出 `.docx`。但有两种运行场景：

- 纯网站部署：可以正常识别、复制 LaTeX、下载 Word 文档。
- Windows + 已安装 Microsoft Word/Office：如果你希望更强的 Word 兼容性，可以按需开启 `WORD_COM_AUTOMATION_ENABLED=true`。

注意：

- `services/word_math.py` 会查找 `MML2OMML.XSL`，通常来自本机 Office 安装目录。
- 如果服务器没有安装 Office，建议先按默认配置运行，不要启用 COM 自动化。

## 健康检查

服务启动后可访问：

- 首页：`/`
- 健康检查：`/api/health`

例如：`http://127.0.0.1:5050/api/health`

## Word 插件说明

当前仓库里的 Word 插件 `manifest.xml` 仍然指向：

- `https://localhost:5050`

这表示它现在还是“本地 sideload 调试版”，不是可以直接分发给外部用户安装的公网版本。

如果后面你要把插件也分发给别人，需要单独做这几件事：

- 把插件地址改成正式域名
- 配置 HTTPS 证书
- 更新 `word_addin/manifest.xml`
- 重新测试 Word 加载、识别、插入链路

## 依赖补充

如果 OCR 相关依赖没装全，可以补装：

```powershell
pip install "paddlex[ocr]==3.4.3"
```

首次运行 PaddleOCR 时，可能会下载模型，第一次识别会慢一些。
