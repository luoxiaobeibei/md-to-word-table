# Markdown表格转Word转换器

[English](README_EN.md) | 简体中文

一个简单易用的工具，可以将Markdown格式的表格快速转换为Word兼容的HTML表格，支持直接复制粘贴。

- 在线使用网址: [https://md-to-word-table.pages.dev/](https://md-to-word-table.pages.dev/)

## 🌟 功能特点

- 一键将Markdown表格转换为Word兼容格式
- 自动添加表格边框和样式，保证在Word中显示美观
- 实时预览转换结果
- 支持中文和英文双语界面
- 响应式设计，适配各种屏幕尺寸
- 支持浅色/深色模式自适应

## 📝 使用说明

1. 将Markdown格式的表格粘贴到文本框中
2. 点击"转换并复制到剪贴板"按钮
3. 在Word文档中直接粘贴(Ctrl+V)即可得到格式化的表格

### Markdown表格示例:

```markdown
| 表头1 | 表头2 | 表头3 |
|-------|-------|-------|
| 内容1 | 内容2 | 内容3 |
| 数据A | 数据B | 数据C |
```

## 🛠️ 技术栈

- 纯前端实现，无需后端支持
- HTML5 + CSS3 + JavaScript
- [Bootstrap](https://getbootstrap.com/) - 用于UI组件和响应式设计
- [Marked.js](https://marked.js.org/) - 用于Markdown解析

## 💻 浏览器兼容性

应用使用了现代Web API，推荐在以下浏览器中使用：

- Google Chrome (最新版)
- Mozilla Firefox (最新版)
- Microsoft Edge (基于Chromium的版本)
- Safari (14+)

注意：在某些情况下，剪贴板功能可能需要HTTPS环境或特定浏览器权限

## 🚀 本地部署

只需将所有文件下载到本地，然后在浏览器中打开`index.html`文件即可。

```bash
# 克隆仓库(如果有Git仓库)
git clone https://github.com/fxaxg/md-to-word-table.git

# 或者直接下载所有文件到本地文件夹
```

也可以将文件部署到任何静态网站托管服务上。

## 🔧 开发说明

如需进行开发，可按以下步骤：

1. 修改`index.html`定制界面
2. 调整`style.css`自定义样式
3. 在`script.js`中扩展功能

## 📄 许可证

[MIT](LICENSE) 

## 贡献

欢迎提交问题和改进建议！ 