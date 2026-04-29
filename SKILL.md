---
name: doc2web
description: Convert large Word documents (.docx, and .doc via LibreOffice when available) into polished, complete static websites with responsive layout, section navigation, tables, images, notes, and search. Use when a user provides a Word document and wants all content preserved in a good-looking website.
metadata:
  short-description: Turn large Word docs into polished websites
---

# Doc2Web（官网风格静态站点）

当用户希望把一个 Word（`.docx` / ` .doc`）文档转换成“可直接上线的官网风格静态网站”，并且要求保留完整内容（不要做摘要/裁剪）时使用本技能。

本技能目标是：在脚本已生成的版式基础上，再进行一次“官网化增强”（视觉、结构、首页大图、SEO、可访问性与验收）。

---

## 一句话产出要求

生成的 `index.html` 应具备“官网级”观感：统一配色与排版、移动端体验可靠、导航与搜索好用，并尽量在首页 `hero` 区域加入一张“大图”（如果文档里有图片优先使用文档首图/封面图）。

---

## Primary Workflow（建议流程）

1. 将用户的 `.docx` 或 `.doc` 文件放入工作区可访问的路径。
2. 运行转换器，生成静态站点：

```bash
python3 scripts/doc2web.py path/to/document.docx --out path/to/site
```

3. 如果输入是老的 `.doc`：
   - 脚本会尝试 `soffice` / `libreoffice` 转成 `.docx`
   - 若本机没有 LibreOffice，请让用户先导出 `.docx` 再继续

4. 打开并检查：
   - `path/to/site/index.html`
   - `path/to/site/assets/`（图片与媒体资源）

---

## 官网化增强（必须做的“第二步”）

### A. 首页 Hero 加“大图”（优先满足）

执行下列策略（按优先级）：

1. 从正文中自动取“第一张真实图片”作为 hero 大图来源：
   - 在生成的 `index.html` 中找到第一处正文图片（例如 `figure.doc-figure img`）
   - 读取它的 `src`（通常是 `assets/..` 下的路径）
2. 若文档中完全没有可用图片：
   - 使用 CSS 渐变/纹理背景并叠加一个简洁的“封面占位图”（可用纯 CSS 或轻量 SVG）
   - 但仍需保证 hero 区域视觉上像“有大图的封面板”

实现方式（推荐，便于一致性）：
- 将 hero 区域从“纯文字”升级为“图文布局”（例如左侧文案+右侧/背景大图，或图片作为 hero 背景）
- 大图应具备：
  - `alt` 文本（来自 Word 图片描述或回退为 `Hero image`）
  - 响应式处理（移动端不应挤压/溢出）
  - 视觉层级（必要时加遮罩层，保证标题可读）

### B. 结构与导航“更像官网”

在不破坏现有内容的前提下，建议补齐官网常见元素：

- 顶部品牌栏（可选）：在侧边栏品牌上再加一个顶部轻量品牌条（或至少保证品牌一致性）
- 页脚（Footer）：加入版权/生成时间/文档来源信息（从文档名或 `--title` 推断）
- 强化 ToC（目录）：
  - 当前脚本已提供侧边目录与层级缩进
  - 建议让“点击后能滚动到对应位置”保持稳定（anchors 已生成就不需要改内容，只要确保 URL hash 正确）

### C. SEO 与分享（可选但推荐）

在 `index.html` 的 `<head>` 中补齐以下元信息（尽量从标题/文档摘要推断，不要凭空捏造内容）：

- `meta name="description"`：用标题+前 1-2 段正文自然拼接（不要摘要式改写太大）
- `og:title` / `og:description` / `og:type`
- `twitter:card`

如果有 hero 大图，则为分享卡片优先使用它作为 `og:image`。

### D. 可访问性（A11y）与键盘体验

建议确保：
- 页面存在清晰的焦点样式（focus outline 不要被移除）
- 搜索输入有可读的 `aria-label`
- 目录 `nav`、正文 `article`、hero `section` 等语义标签保留
- 表格可横向滚动且不造成移动端不可用（脚本已做 `overflow-x: auto`，验收要覆盖）

### E. 图片与资源优化（轻量化要求）

若文档图片较多且体积过大：
- 允许做“轻量压缩”（例如转 WebP 或缩放到合理宽度）
- 若环境不确定或缺少压缩工具，不强制；但要在验收中记录“是否存在明显卡顿”

---

## Design Guidance（设计风格约束）

- 保留全部内容：不做摘要、不删除章节、不合并段落（除非用户明确要求）。
- 优先“编辑型排版”质感：清晰层级、充足留白、可读字号、统一链接颜色与下划线规则。
- 大文档可导航是关键：anchors + 侧边目录 + 搜索 + 滚动进度条缺一不可（脚本已有，验收需确认）。
- 如果原始 Word 样式很差导致层级不理想：
  - 只在确认生成结构后，对 HTML 进行最小必要修正（例如调整标题层级、修复 anchors）

---

## 验收标准（Checklist）

在浏览器中确认以下项（至少覆盖桌面宽度与移动端宽度）：

- 首页 `hero`：标题可读，hero 大图（或占位封面）在首屏清晰可见
- 侧边目录：主要标题均出现在目录里，层级层级缩进合理
- 点击目录：跳转到对应章节不偏移、无 404/错误 hash
- 搜索：能命中预期片段；无匹配时提示文案合理
- 表格：移动端可横向滚动，内容不被截断到不可读
- 图片：从 `assets/` 正常加载，alt 不为空（至少 hero 与首图建议有）
- 备注/脚注：底部展示在合理位置（脚本会将 footnotes/endnotes 作为 aside 块输出）
- Lighthouse（可选）：性能与可访问性不过度拉跨（如果无法跑 Lighthouse，至少人工检查主要页面交互与布局稳定性）

---

## Script Notes（脚本能力假设）

- `scripts/doc2web.py` 使用 Python 标准库解析 `.docx`，不依赖网络与第三方包。
- 若遇到 `.doc`：脚本会尝试 LibreOffice/soffice 转换为 `.docx`。
- Word 内嵌图片会被提取到站点 `assets/` 下。
- 会尽可能包含：超链接、标题、段落、表格、页眉/页脚、脚注/尾注。
- 输出是纯静态站点，可本地打开或上传到任意静态站点托管。
