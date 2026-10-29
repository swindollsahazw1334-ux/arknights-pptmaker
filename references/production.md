# 本地制作与场景格式

助手在后台运行这些命令，用户只需在对话中提供资料和回答需求。使用 skill 的绝对路径定位脚本，命令不要求当前目录是某个代码仓库。

## 读取资料

`sources.py` 使用 Python 标准库，提取 DOCX 主正文的标题、段落、表格和内嵌图片，保留换行与来源定位；也可接收 TXT、Markdown、PNG/JPEG/WebP。原始文件保持不变。PDF 等其他格式使用环境中的文档/PDF 工具读取，不要先搭建上传服务。

```powershell
python '<skill-dir>/scripts/sources.py' extract --out '<work-dir>/source.json' '<attachment.docx>' '<photo.png>'
```

结果的 `blocks` 包含 `id`、`kind`、`text`、`source.fileName`、`source.location`，表格有 `rows`，图片有 `assetId`；`assets[].path` 相对 `source.json` 所在目录，提取图片存放于旁边独立的素材子目录。检查 `warnings` 并实际看图，再自行理解、摘要和安排各页；不要机械按字符数分页。

不依赖 Word 的分页计算，因此来源用段落/表格序号，不生成虚假 Word 页码。页眉、页脚、脚注等特殊结构需要另外读取原始 DOCX；错误或警告不能被“已完整读取”掩盖。

## 创作视觉素材

完成内容计划与主题研究后，按 [生图与素材分层](imagegen.md) 实际调用内置 imagegen，制作适合资料内容的背景、透明装饰和必要插画。把选用图片与完整提示词保存在本次工作目录，再由场景文件引用。下方示例中的图片路径代表已生成并检查的实际素材，导出 helper 本身不会执行生图。

## 准备场景

将助手设计的对象保存为 UTF-8 `scene.json`。逻辑画布固定为 1280×720，坐标/字号采用像素，导出为 13⅓×7½ 英寸。数组顺序即从底层到顶层。图像路径相对场景文件或使用绝对路径。

```json
{
  "title": "海岸清洁方案",
  "theme": "孤星",
  "brief": { "pageCount": 1, "density": "balanced" },
  "slides": [{
    "id": "p1",
    "title": "海岸清洁方案",
    "background": "#F2EDD8",
    "sourceIds": ["f1-b1", "f1-b2"],
    "elements": [
      { "id": "p1-bg", "name": "生图背景：海岸档案纸景", "type": "image", "path": "visual-assets/bg-coast-archive.png", "x": 0, "y": 0, "w": 1280, "h": 720, "fit": "cover" },
      { "id": "p1-deco", "name": "生图透明装饰：轨道纸饰", "type": "image", "path": "visual-assets/deco-orbit.png", "x": 940, "y": 90, "w": 280, "h": 400, "fit": "contain" },
      { "id": "p1-title", "name": "主题标题", "type": "text", "x": 96, "y": 120, "w": 800, "h": 90, "text": "海岸清洁方案", "fontSize": 56, "bold": true, "color": "#242825" },
      { "id": "p1-summary", "name": "预算说明", "type": "text", "x": 100, "y": 300, "w": 780, "h": 160, "text": "活动预算为1200元。", "fontSize": 28, "sourceIds": ["f1-b2"] }
    ]
  }]
}
```

基础字段：

| 对象 | 字段 |
| --- | --- |
| 共同 | 唯一 `id`、`name`、`type`、`x/y/w/h`；可选 `rotation`（度）、`opacity`（0–1） |
| 文字 | `text`、`fontSize`、`fontFace`、`bold`、`italic`、`color`、`align`（left/center/right）、`lineHeight` |
| 形状 | `shape`（rect/roundRect/ellipse/line）、`fill`（十六进制颜色或 none）、`stroke`、`strokeWidth` |
| 图片 | `path`、`fit`（contain/cover）；相框、说明和图片本身各自独立 |
| 页面 | `id`、`title`、`background`、`sourceIds`、`elements`，可选 `notes` |

`brief.pageCount` 必须来自本次问答，实际页数必须与之相等。页数更改由用户授权；不能为通过检查而反向改写 brief。

背景、装饰和插画引用独立的生成素材；形状字段用于需要可编辑的表格、数据图形和流程关系等信息对象。按实际图像确定留白与叠放位置，不只依据提示词假定文字安全区。

## 导出与检查

在 Codex 桌面环境，先用 `load_workspace_dependencies` 获取 Node、Node packages 和 Python 的当前绝对路径。导出依赖 `pptxgenjs`；WebP 转 PNG 额外使用 `sharp`。现有运行时已提供这些库时直接复用，不要求用户安装网站依赖。普通 Node 环境可从脚本所在目录或 `--modules` 指定的 node_modules 加载已安装的库。

```powershell
& '<node-executable>' '<skill-dir>/scripts/export.mjs' --modules '<node-packages>' --scene '<work-dir>/scene.json' --out '<work-dir>/成品.pptx'
python '<skill-dir>/scripts/sources.py' inspect '<work-dir>/成品.pptx'
```

导出使用原生文字、独立形状和独立图片，并将页面来源写入演讲者备注。`inspect` 读取实际 PPTX 内部结构，报告页数、文字对象、非文字形状、图片和各页文本。它不证明字体、换行或遮挡正确。若环境提供演示稿渲染工具，再渲染核查；可用已安装的 PowerPoint/WPS/LibreOffice 检查，不为验证无故安装整套办公软件。

保留 `scene.json`、`source.json`、`brief.json`、生成原图、完整提示词、素材对应关系和 `.pptx` 在本次工作目录。最终交付链接指向 PPTX；场景 JSON 是后续修改用的源文件，不是让用户填写的表单。
