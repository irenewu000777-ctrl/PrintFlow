# PrintFlow 印流

> 浏览器端拼版打印工具 — 多页 PDF / 课件一键合并排版，节省纸张，本地处理，保护隐私。

[![在线体验](https://img.shields.io/badge/在线体验-printflow777.vercel.app-blue?style=flat-square)](https://printflow777.vercel.app/)
[![部署](https://img.shields.io/badge/部署-Vercel-black?style=flat-square&logo=vercel)](https://printflow777.vercel.app/)

面向学生与轻量办公场景：上传 PDF、PPT 或 PPTX，按纸张、方向、行列与阅读顺序自动拼版，导出可直接打印的 PDF。

**[立即使用 → printflow777.vercel.app](https://printflow777.vercel.app/)** · **[GitHub 源码](https://github.com/irenewu000777-ctrl/PrintFlow)**

---

## 怎么用

打开网站，四步完成：

| 步骤 | 操作 |
| --- | --- |
| 1 | 上传 PDF / PPT / PPTX（最大 40 MB） |
| 2 | 配置纸张尺寸、方向、行列数、间距 |
| 3 | 选择排列顺序，右侧实时预览拼版效果 |
| 4 | 点击「生成 PDF / Generate PDF」，自动下载 |

导出文件名：`{原文件名}-study-layout.pdf`

## 功能

- **格式**：PDF、PPT、PPTX
- **纸张**：A4、A5、Letter、B5，纵向 / 横向
- **拼版**：1–4 行 × 1–4 列，间距可调（mm）
- **顺序**：横向优先（Horizontal Pattern）/ 纵向优先（Vertical Pattern）
- **预览**：参数变更即时刷新（Live View）
- **隐私**：Local processing only — 文件不离开浏览器

## 工作原理

线上版本采用**纯浏览器本地处理**，无需后端转换服务，也无需配置环境变量：

```
PDF      → pdfjs-dist 渲染预览 → pdf-lib 拼版导出
PPT/PPTX → pptx-preview 渲染幻灯片 → html-to-image 截图 → pdf-lib 拼版导出
```

文件始终在用户设备上处理，不会上传到 Vercel 或任何第三方服务器。

## 常见问题

**课件预览不完整或报错？**

复杂 PPT/PPTX（动画、嵌入对象较多）可能在浏览器端渲染不完整。请先用 PowerPoint / WPS 导出为 PDF，再重新上传 — 这与网站上的提示一致：「若转换出现缺失，请将文件转换为 PDF 重试」。

**需要登录或安装吗？**

不需要。打开网页即可使用，无需注册、无需插件。

## 本地开发

环境要求：Node.js 18+、npm

```bash
git clone https://github.com/irenewu000777-ctrl/PrintFlow.git
cd PrintFlow
npm install
npm run dev
```

访问 [http://localhost:3000](http://localhost:3000)

```bash
npm run build   # 生产构建
npm run start   # 启动生产服务
npm run lint    # ESLint 检查
```

## 技术栈

Next.js 14 · React 18 · TypeScript · Tailwind CSS · pdfjs-dist · pdf-lib · pptx-preview · html-to-image · framer-motion

部署于 [Vercel](https://vercel.com) → [printflow777.vercel.app](https://printflow777.vercel.app/)

## 项目结构

```
app/
  page.tsx              # 主工作台
  layout.tsx
components/
  ControlPanel.tsx      # 上传与排版控制
  PreviewPane.tsx       # 实时预览
lib/
  pipeline.ts           # 文件 → 页面快照
  layout.ts             # 拼版布局逻辑
  exportPdf.ts          # PDF 导出
  constants.ts          # 纸张、限制等常量
  types.ts
docs/
  PRD.md                # 产品需求文档
```

## 许可证

Private — 详见 [GitHub 仓库](https://github.com/irenewu000777-ctrl/PrintFlow) 设置。
