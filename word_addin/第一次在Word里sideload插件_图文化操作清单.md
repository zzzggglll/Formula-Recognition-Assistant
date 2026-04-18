# 第一次在 Word 里 Sideload 插件

## 你要完成的目标

把 [manifest.xml](e:/公式识别/word_addin/manifest.xml) 正确加载到桌面版 Word 里，并成功打开“公式识别助手”任务窗格。

这份清单默认你使用的是：

- Windows
- 桌面版 Microsoft Word
- 当前项目目录：`E:\公式识别`

## 先记住一件事

不要双击打开 `manifest.xml`。

`manifest.xml` 不是普通 Word 文档，它是插件清单文件。  
如果你直接双击，Word 往往会把它当成 XML 文档打开，然后弹出“此文件包含 Word 不再支持的自定义 XML 元素”之类的提示。

正确方式是：  
在 Word 里走“上传我的加载项”。

---

## Step 0：准备清单

开始前，先确认下面 5 件事：

- 你已经安装桌面版 Microsoft Word
- 你的项目目录是 `E:\公式识别`
- 你已经创建并激活了项目虚拟环境 `.venv`
- 你已经安装依赖：`pip install -r requirements.txt`
- 你准备从 Word 里“上传加载项”，不是双击 `manifest.xml`

如果这 5 条都满足，再继续下一步。

---

## Step 1：启动本地 HTTPS 服务

在 PowerShell 里执行：

```powershell
cd E:\公式识别
.\.venv\Scripts\activate
$env:FLASK_SSL_ADHOC='true'
python app.py
```

如果你不想受外层环境影响，也可以直接用项目自己的解释器：

```powershell
cd E:\公式识别
$env:FLASK_SSL_ADHOC='true'
.\.venv\Scripts\python.exe app.py
```

### 你应该看到

- 终端里没有立刻报错
- 服务监听在 `https://localhost:5050`

### 如果这一步失败

- 看终端报错是不是缺依赖
- 如果提示某个包没装，重新执行：

```powershell
pip install -r requirements.txt
```

---

## Step 2：先在浏览器里打开一次插件页面

打开浏览器，访问：

```text
https://localhost:5050/word-addin/taskpane
```

### 你应该看到

- 页面标题类似“在 Word 里识别图片公式并直接插入文档”
- 页面里有上传区
- 页面里有“识别并准备插入”按钮

### 如果浏览器提示证书不安全

这是本地自签名 HTTPS 的常见现象。  
你需要先手动“继续访问此网站”或“高级 -> 继续前往”。

这一步很重要。  
如果浏览器里都打不开，Word 里的任务窗格大概率也打不开。

---

## Step 3：打开桌面版 Word

打开 Microsoft Word，建议先新建一个空白文档。

### 你应该看到

- 普通 Word 编辑界面
- 当前文档可以正常输入文字

---

## Step 4：进入“上传我的加载项”

在 Word 顶部功能区里，按下面顺序找：

1. 点击 `插入`
2. 找 `我的加载项` 或 `获取加载项`
3. 打开后进入“我的加载项”面板
4. 找到 `上传我的加载项`

有些 Word 版本入口文字会略有不同，但核心是：

- 先进入加载项面板
- 再点击“上传我的加载项”

### 你应该看到

- 一个文件选择对话框
- 或者一个“上传我的加载项”的小窗口

### 特别注意

这一步是“从 Word 内上传”。  
不是资源管理器里双击 `manifest.xml`。

---

## Step 5：选择 manifest 文件

在文件选择框里，选择这个文件：

[manifest.xml](e:/公式识别/word_addin/manifest.xml)

完整路径是：

```text
E:\公式识别\word_addin\manifest.xml
```

然后点击“上传”或“打开”。

### 你应该看到

- Word 不再把它当普通 XML 文档打开
- 功能区里出现“公式识别助手”相关按钮
- 或者右侧直接弹出任务窗格

### 如果这里又弹出“自定义 XML 元素”警告

说明你还是把 `manifest.xml` 当文档打开了。  
这时回到上一步，重新确认你走的是：

`Word -> 插入 -> 我的加载项 / 获取加载项 -> 上传我的加载项`

而不是双击文件。

---

## Step 6：点击功能区里的插件按钮

上传成功后，在 Word 顶部功能区里，找到：

- `公式识别助手`
- 或按钮文字：`识别公式`

点击它，打开右侧任务窗格。

### 你应该看到

- 右侧出现插件面板
- 面板标题类似“在 Word 里识别图片公式并直接插入文档”
- 面板里有上传区域和识别按钮

---

## Step 7：做一次最小可用测试

在任务窗格中：

1. 上传一张公式截图
2. 点击 `识别并准备插入`
3. 等待识别结果出现
4. 点击 `插入到当前 Word 文档`

### 你应该看到

- 面板里出现 LaTeX 结果
- 面板里出现公式预览
- 点击插入后，Word 当前光标位置出现公式

---

## Step 8：如果任务窗格打不开 localhost

在某些 Windows + Office 环境里，Word 的 WebView 默认不能访问本机 `localhost`。

这时用“管理员 PowerShell”执行：

```powershell
CheckNetIsolation LoopbackExempt -a -n="microsoft.win32webviewhost_cw5n1h2txyewy"
```

执行后：

1. 关闭 Word
2. 保持 Flask 服务运行
3. 重新打开 Word
4. 再次上传加载项

---

## Step 9：如果任务窗格空白

先按这个顺序检查：

1. 浏览器能不能打开 `https://localhost:5050/word-addin/taskpane`
2. 终端里的 Flask 服务有没有报错
3. 你有没有先信任本地 HTTPS 证书页面
4. Word 是否是桌面版，而不是其他受限环境

如果浏览器能打开，但 Word 任务窗格空白，优先怀疑：

- 本地 HTTPS 证书没被信任
- localhost loopback 没开

---

## Step 10：如果识别成功但插入失败

这说明插件已经加载成功，问题不在 sideload，而在“插入公式”这一步。

按下面顺序排查：

1. 看 Flask 终端有没有报错
2. 确认本机 Office 安装完整
3. 确认 `MML2OMML.XSL` 能被程序找到
4. 先点击 `下载 Word 原生公式 .docx` 做兜底

如果需要手动配置 `MML2OMML.XSL`，在 `.env` 里加：

```env
WORD_MML2OMML_XSL_PATH=C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL
```

---

## 最短成功路径

如果你只想按最短路径做一遍，照这个点：

1. 运行：

```powershell
cd E:\公式识别
$env:FLASK_SSL_ADHOC='true'
.\.venv\Scripts\python.exe app.py
```

2. 浏览器打开：

```text
https://localhost:5050/word-addin/taskpane
```

3. 信任本地证书页面

4. 打开 Word 新建空白文档

5. 点击：

`插入 -> 我的加载项 / 获取加载项 -> 上传我的加载项`

6. 选择：

[manifest.xml](e:/公式识别/word_addin/manifest.xml)

7. 点击功能区里的：

`公式识别助手 -> 识别公式`

8. 在右侧任务窗格上传图片并测试插入

---

## 你可以把这份清单和这些文件一起看

- [manifest.xml](e:/公式识别/word_addin/manifest.xml)
- [README.md](e:/公式识别/word_addin/README.md)
- [app.py](e:/公式识别/app.py)

---

## 参考资料

以下是我整理这份清单时参考的官方文档：

- Microsoft Learn: Office Add-ins overview  
  https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins
- Microsoft Learn: Sideload Office Add-ins to Office on the web  
  https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing
- Microsoft Learn: Cannot open add-in from localhost  
  https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/office-suite-issues/cannot-open-add-in-from-localhost
