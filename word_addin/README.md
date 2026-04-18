# Word 插件版 MVP 使用说明

## 1. 当前实现范围

当前插件版 MVP 已完成以下主链路：

- 在 Word 任务窗格中上传公式截图
- 调用现有 Flask 后端识别公式
- 显示识别得到的 LaTeX 与公式预览
- 一键插入到当前 Word 文档
- 插入失败时，可退回下载 Word 原生公式 `.docx`

当前版本暂未实现：

- 用户在插件内手动修改 LaTeX
- 账号体系与历史记录同步
- 多公式批量识别

## 2. 相关文件

- `word_addin/manifest.xml`
- `templates/word_addin_taskpane.html`
- `templates/word_addin_commands.html`
- `static/word-addin/taskpane.js`
- `static/word-addin/taskpane.css`
- `static/word-addin/commands.js`

## 3. 本地启动

建议先开启 HTTPS 本地服务，因为 Office Add-in 本地调试通常需要 HTTPS：

```powershell
cd E:\公式识别
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
$env:FLASK_SSL_ADHOC='true'
python app.py
```

启动后，任务窗格页面地址通常为：

- `https://localhost:5050/word-addin/taskpane`

## 4. Word 中 sideload

1. 打开桌面版 Microsoft Word
2. 进入“插入 -> 我的加载项 -> 上传我的加载项”
3. 选择 `E:\公式识别\word_addin\manifest.xml`
4. 成功后，Word 功能区首页会出现“公式识别助手”按钮
5. 点击按钮即可打开任务窗格

## 5. 常见问题

### 5.1 任务窗格打不开 localhost

在 Windows 桌面 Office 中，如果加载项页面来自 `localhost`，有时需要为 WebView 打开 loopback 权限。

可尝试以管理员 PowerShell 执行：

```powershell
CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"
```

### 5.2 自签名 HTTPS 证书不受信任

如果 `ssl_context='adhoc'` 生成的本地证书不被信任，Word 可能无法正常打开任务窗格。

处理方式：

- 优先用本地受信任证书方案，例如 `mkcert`
- 或使用 dev tunnel / 反向代理，把本地 Flask 暴露为可信 HTTPS 地址
- 如果后续正式部署，建议直接使用公网 HTTPS 域名

### 5.3 插入失败

如果插件里的“插入到当前 Word 文档”失败：

- 先确认后端 `/api/word-addin/recognize` 返回正常
- 再确认本机 Office 安装完整，并且 `MML2OMML.XSL` 可被后端找到
- 如果仍失败，先使用“下载 Word 原生公式 .docx”作为兜底方案

## 6. 下一步建议

插件版 MVP 下一阶段建议按以下顺序增强：

1. 增加 LaTeX 手动修正区
2. 插件内保留最近一次识别历史
3. 接入网站版历史记录
4. 做更稳定的 Word 公式插入兼容测试
