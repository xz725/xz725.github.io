# AGENTS.md

个人学术主页（GitHub Pages / Jekyll，Lanyon 主题），部署于 https://xz725.github.io 。

## 架构

- **单一数据源**：`raw_data/about.xlsx`。sheet 2 到最后一个 sheet 与网站侧边栏标签页一一对应
  （工作经历 / 学术兼职 / 教育经历 / 荣誉奖项 / 期刊文章 / 会议文章 / 美国专利 / 中国专利）。
- **数据流**：`about.xlsx` → `.agents/skills/update-website/scripts/build_data.py` →
  `_data/about.json` → `_layouts/section.html`（Liquid）→ 各标签页。
- **双语**：中文为默认（`/`），英文镜像于 `/en/`。页面 stub 仅含 frontmatter
  （`layout: section, section: <id>, lang: zh|en`），渲染逻辑全部在 `section.html`。
- **标签页定义**：`_config.yml` 的 `tabs:` 列表（id / 中文名 / 英文名 / URL）。

## 更新网站内容（常规操作）

1. 编辑 `raw_data/about.xlsx`（保持各 sheet 前两行为标题行/表头）。
2. 运行 skill `update-website`（或手动）：
   ```bash
   uv run --with openpyxl .agents/skills/update-website/scripts/build_data.py
   ```
3. `git add -A && git commit && git push` —— GitHub Pages 自动构建部署。

## 约定

- 不要手改 `_data/about.json`，只能由脚本重新生成。
- 不要新增 Jekyll 插件（GitHub Pages 白名单限制）；站点仅依赖 `_data` + Liquid。
- xlsx 数据格式约定见 `.agents/skills/update-website/SKILL.md`（占位符 `—`、`至今`、
  引用列互换检测、博士论文行编号 `T` 等）。
- 已删除全部 demo（原 `_posts/`、`projects/`、`resources/` 见 git 历史）。
