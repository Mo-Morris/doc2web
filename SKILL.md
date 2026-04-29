---
name: doc2web
description: Convert Word documents (.docx, and .doc via LibreOffice when available) into premium official-style static websites. Use when a user provides a document and wants a polished, launch-ready website that preserves the full source content while first curating, grouping, and summarizing it from multiple compelling dimensions instead of presenting it like an ebook.
metadata:
  short-description: Turn Word docs into premium official-style websites
---

# Doc2Web（官网级内容策展站点）

当用户希望把 Word（`.docx` / `.doc`）转换成“可直接上线的官网级静态网站”时使用本技能。输出不能只是电子书、阅读器或文档搬运页；必须先把内容策展成有吸引力的官网叙事，再保留完整原文档作为可搜索、可核验的下层内容。

本技能目标：把文档变成有品牌感、结构感和传播力的网站。首页负责吸引与归纳，正文负责完整与可信。

---

## 一句话标准

生成的 `index.html` 必须像一个正式官网：首屏大气、有视觉冲击、有清晰价值主张；中段从多个维度归类总结内容；底部才是完整文档库、目录、搜索、表格、图片和注释。

---

## Primary Workflow

1. 将用户的 `.docx` 或 `.doc` 文件放入工作区可访问的路径。
2. 先为首页生成一张官网级 hero 主视觉图：
   - 首选图片生成模型（如 image2 / 可用的 image generation tool），基于文档标题、行业、关键词、情绪和目标受众生成。
   - 画面必须大气、有视觉冲击、有格调，适合作为官网首屏背景；不要做廉价插画、普通封面、电子书封面、PPT 背景或纯抽象渐变。
   - 建议输出宽屏比例（16:9、21:9 或至少 1600px 宽），主体留出文字叠加空间。
   - 如果图片生成失败，才降级使用文档封面/首图；如果仍不可用，再用高质量 CSS 背景作为最后兜底。
3. 运行转换器，生成静态站点。若已生成 hero 图，必须传给脚本：

```bash
python3 scripts/doc2web.py path/to/document.docx --out path/to/site --hero-image path/to/generated-hero.png
```

没有生成图时才省略 `--hero-image`：

```bash
python3 scripts/doc2web.py path/to/document.docx --out path/to/site
```

4. 如果输入是老的 `.doc`：
   - 脚本会尝试 `soffice` / `libreoffice` 转成 `.docx`
   - 若本机没有 LibreOffice，请让用户先导出 `.docx` 再继续

5. 打开并检查：
   - `path/to/site/index.html`
   - `path/to/site/assets/`（图片与媒体资源）
6. 必须进行人工官网化复核：检查首页叙事、生成主视觉、维度归类、视觉质感、移动端和完整内容是否都达标。脚本只是第一版，不是最终审美责任的终点。

---

## 内容策展要求（必须）

不要把 Word 的章节顺序原样堆到首屏。生成或调整网站时，先把内容重组为 3-6 个有吸引力的入口维度。可用维度包括：

- 核心主张：这份文档最想让读者相信或理解什么
- 关键数据/证据：数字、表格、案例、图示、调研结论
- 受众/场景：不同角色为什么应该关心
- 产品/方案/能力：文档中可被包装成官网模块的能力点
- 进展/路线图：时间线、阶段、里程碑、下一步
- 风险/挑战/机会：把复杂内容变成判断框架
- 方法论/流程：把操作性内容变成清晰步骤
- 成果/影响：对业务、用户、组织、社会或市场的价值

每个维度都应有官网式标题和 1-3 句高度概括。概括要忠于原文，不能编造事实；不确定时使用“文档呈现/文档显示/可从章节中看到”等措辞。

完整内容仍要保留，但完整文档应放在“Document / 原文档库 / 深度阅读”区域，而不是把它作为唯一体验。

---

## 视觉与信息架构要求

- 首屏：必须有强标题、短副标题、来源/章节数/阅读时间等元信息；hero 主视觉优先使用图片生成模型生成的大图，并作为沉浸式背景。
- Hero 图片：必须大气、有冲击力、官网质感强。生成失败才使用文档首图/封面图；不要把“文档里第一张图”作为默认首选。
- 中段：必须有“策展维度”模块和“精选主题/章节”模块，用卡片、分栏、时间线、对比、数据条等官网组件组织。
- 完整文档：保留目录、搜索、锚点、表格横向滚动、图片、脚注/尾注、页眉/页脚。
- 页脚：标明来源文档或生成信息。
- 观感：大气、克制、现代、有格调。避免廉价渐变、装饰性气泡、过圆卡片、电子书式长栏正文首屏。
- 卡片圆角保持克制（8px 或更小），文字不能溢出或互相遮挡。
- 配色不要单一色系；避免整页只有米色、深蓝、紫蓝或棕橙。

---

## SEO 与分享

在 `index.html` 的 `<head>` 中补齐：

- `meta name="description"`：标题 + 原文前段或忠实概括
- `og:title` / `og:description` / `og:type`
- `twitter:card`
- 如果有 hero 大图，优先补 `og:image`

不要为了营销感凭空添加原文没有的信息。

---

## 可访问性与交互

- 页面存在清晰的焦点样式（focus outline 不要被移除）
- 搜索输入有可读的 `aria-label`
- 目录 `nav`、正文 `article`、hero `section` 等语义标签保留
- 表格可横向滚动且不造成移动端不可用（脚本已做 `overflow-x: auto`，验收要覆盖）
- 移动端首屏、卡片、目录、搜索结果不能重叠

---

## 图片与资源

首页 hero 图生成优先级：

1. 图片生成模型（首选，必须先尝试）：用文档标题、主题、行业、关键词、受众和情绪生成官网主视觉。
2. 文档封面/首图（仅当生成失败）：确保裁切、遮罩和文字可读性。
3. CSS 背景（最后兜底）：只能作为失败兜底，不能作为常规方案。

若文档图片较多且体积过大：
- 允许做“轻量压缩”（例如转 WebP 或缩放到合理宽度）
- 若环境不确定或缺少压缩工具，不强制；但要在验收中记录“是否存在明显卡顿”

---

## Script Notes

- `scripts/doc2web.py` 使用 Python 标准库解析 `.docx`，不依赖网络与第三方包。
- 若遇到 `.doc`：脚本会尝试 LibreOffice/soffice 转换为 `.docx`。
- Word 内嵌图片会被提取到站点 `assets/` 下。
- 会尽可能包含：超链接、标题、段落、表格、页眉/页脚、脚注/尾注。
- 当前默认输出已经包含官网化首页、策展维度、精选主题、完整文档、目录、搜索、进度条与 SEO 基础信息。
- `--hero-image path/to/image.png` 会把生成的 hero 主视觉复制到 `assets/` 并优先用于首屏背景和分享图。

---

## 验收标准（Checklist）

在浏览器中确认以下项（至少覆盖桌面宽度与移动端宽度）：

- 首屏：像官网，不像电子书；标题、副标题、元信息清晰，图片生成模型产出的 hero 大图有视觉冲击
- 内容策展：至少 3 个维度对原文进行归类总结，并能链接或引导到完整内容
- 精选主题：从多个角度呈现重点章节，而不是只按原文顺序罗列
- 完整保留：原文主要标题、段落、表格、图片、脚注/尾注均在下方文档区可见
- 侧边目录：主要标题均出现在目录里，层级缩进合理
- 点击目录：跳转到对应章节不偏移、无 404/错误 hash
- 搜索：能命中预期片段；无匹配时提示文案合理
- 表格：移动端可横向滚动，内容不被截断到不可读
- 图片：从 `assets/` 正常加载，alt 不为空（至少 hero 与首图建议有）
- 备注/脚注：底部展示在合理位置（脚本会将 footnotes/endnotes 作为 aside 块输出）
- 响应式：移动端无重叠、无横向页面溢出（表格自身滚动除外）
- Lighthouse（可选）：性能与可访问性不过度拉跨
