# Xiguang Zheng's Homepage

Personal academic homepage · Jekyll (Lanyon theme) · GitHub Pages · 中英双语

个人主页，部署于 <https://xz725.github.io>。中文为默认语言（`/`），英文位于 `/en/`。

## 如何更新内容 / How to update

所有页面内容来自 `raw_data/about.xlsx`（唯一数据源）。修改 xlsx 后重新生成并部署：

```bash
uv run --with openpyxl .agents/skills/update-website/scripts/build_data.py
git add -A && git commit -m "update content" && git push
```

GitHub Pages 会自动完成 Jekyll 构建与部署。详见 `AGENTS.md`。

## 目录结构

```
raw_data/about.xlsx   # 内容数据源（唯一需要维护的文件）
_data/about.json      # 由脚本生成，请勿手改
_layouts/section.html # 所有标签页的渲染模板
_includes/sidebar.html# 侧边栏（标签页 + 语言切换）
_config.yml           # 站点配置 + 标签页定义
.agents/skills/update-website/  # 内容更新 skill
```
