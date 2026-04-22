# Study Imposition Engine

一个可上线的 Next.js + TypeScript Web 应用，用于学生课件打印排版（imposition）。

## 核心原则

- 页面不可变：每一页先转成 `Page[]` 预览单元
- 排版只做空间排列：行列、纸张、方向、顺序
- 不改内容，不重绘语义，不解析 PPT 结构

## 技术栈

- Next.js App Router
- React + TypeScript
- Tailwind CSS
- pdf-lib（导出拼版 PDF）
- pdfjs-dist（PDF 页面渲染预览）
- framer-motion（预览动画）

## 快速开始

```bash
npm install
npm run dev
```

访问 `http://localhost:3000`

## 功能覆盖

1. 上传文件（PDF / PPT / PPTX）
2. 统一输入管线到 `Page[]`
3. 行列布局：1-4 行、1-4 列
4. 排列方式：Z-pattern / N-pattern
5. 纸张：A4 / A5 / Letter / B5
6. 方向：Portrait / Landscape
7. 实时预览
8. 导出拼版 PDF（多页自动分页）
9. 导出后随机鼓励语

## PPT/PPTX 转换说明

当前实现支持两种转换方式（优先级从高到低）：

1. 远程转换服务（推荐生产）
2. 本机 LibreOffice 命令行转换（推荐本地开发）

### 方案 A：远程转换服务

配置环境变量：

```bash
CONVERT_API_URL=https://your-convert-service.example.com/convert
CONVERT_API_KEY=your-token-optional
```

要求该服务接受 multipart `file`，返回 `application/pdf` 二进制。

### 方案 B：本机 LibreOffice

安装 LibreOffice，并确保命令行可执行：

- Windows: `soffice.exe`
- Linux/macOS: `soffice`

接口会调用：

```bash
soffice --headless --convert-to pdf --outdir <tmp> <input.pptx>
```

只要能产出标准 PDF bytes，后续拼版流程无需改动。

## 项目结构

```text
app/
  api/convert/route.ts       # 输入文件转 PDF 接口（可替换实现）
  globals.css
  layout.tsx
  page.tsx                   # 主工作台
components/
  ControlPanel.tsx           # 左侧控制面板
  PreviewPane.tsx            # 右侧实时预览
lib/
  constants.ts               # 纸张参数、鼓励语、单位换算
  exportPdf.ts               # pdf-lib 拼版导出
  layout.ts                  # rows/columns + Z/N 排序逻辑
  pipeline.ts                # 文件 -> Page[] 管线
  types.ts                   # 类型定义
```

## GitHub 部署

```bash
git init
git add .
git commit -m "feat: build study imposition MVP"
git branch -M main
git remote add origin https://github.com/<your-name>/<repo>.git
git push -u origin main
```

## Vercel 部署

1. 在 Vercel 导入 GitHub 仓库
2. Framework Preset 选择 Next.js
3. Build Command: `npm run build`
4. Output Directory: `.next`
5. 点击 Deploy

## 0 成本上线（推荐组合）

目标：`Vercel(前端)` + `Oracle Always Free(转换服务)`

### 1) 部署主站到 Vercel（免费）

按上方 Vercel 步骤部署本项目。

### 2) 部署 `converter-service` 到 Oracle Cloud Always Free

该项目已内置转换服务目录：`converter-service/`

- 提供接口：`POST /convert`（multipart 字段名：`file`）
- 健康检查：`GET /health`
- 支持环境变量：`CONVERTER_API_KEY`（可选）

在 Oracle 免费实例中安装 Docker 后执行：

```bash
cd converter-service
docker build -t pptx-converter:latest .
docker run -d --name pptx-converter \
  -p 8080:8080 \
  -e CONVERTER_API_KEY=your-secret-key \
  --restart always \
  pptx-converter:latest
```

然后把该实例通过域名或反向代理暴露为 `https://your-converter-domain/convert`。

### 3) 在 Vercel 配置环境变量

```bash
CONVERT_API_URL=https://your-converter-domain/convert
CONVERT_API_KEY=your-secret-key
```

配置后重新部署即可完成端到端 PPT/PPTX 自动转换。

如果你在 Vercel 遇到 `FUNCTION_PAYLOAD_TOO_LARGE`（大文件上传超限），请改为前端直连转换服务：

```bash
NEXT_PUBLIC_CONVERT_API_URL=https://your-converter-domain/convert
NEXT_PUBLIC_CONVERT_API_KEY=your-secret-key
```

这样上传将绕过 Vercel `/api/convert`，直接发送到转换服务。
对应地，请在 `converter-service` 配置 `CORS_ORIGIN`（如你的 Vercel 域名）。

### 生产建议

- Vercel Serverless 通常不适合直接运行 LibreOffice，推荐使用外部转换服务并配置 `CONVERT_API_URL`
- 若转换服务需要密钥，放入 Vercel Environment Variables（如 `CONVERT_API_KEY`）
