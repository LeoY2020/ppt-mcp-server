# PPT MCP Server

## Server Config 名称

`ppt-mcp-server`

## Server Config

```json
{
  "mcpServers": {
    "ppt-mcp-server": {
      "command": "npx",
      "args": [
        "-y",
        "ppt-mcp-server@1.0.0"
      ]
    }
  }
}
```

---

## 语言类型

Node.js (TypeScript)

---

## 协议类型

- **传输协议**: stdio
- **MCP 版本**: 1.0.0
- **SDK 版本**: @modelcontextprotocol/sdk ^1.6.1

---

## 工具集描述

提供三个工具：`ppt_read_presentation` 读取整个PPTX文件内容；`ppt_get_slide` 获取指定幻灯片详情；`ppt_get_info` 获取元数据和摘要。支持文本样式、形状位置提取。

---

## MCP 概要

### 功能概述

PPT MCP Server 是一个用于读取和解析 PowerPoint (PPTX) 文件的 MCP 服务器。它能够：

- 解析 PPTX 文件结构（PPTX 本质是 ZIP 压缩的 XML 文件集合）
- 提取幻灯片中的文本内容
- 保留文本样式信息（加粗、斜体、字号、颜色等）
- 获取形状的位置和尺寸
- 识别图片、图表等非文本元素
- 提取演示文稿元数据（标题、作者、创建/修改时间等）

### 技术实现

- **语言**: TypeScript / Node.js
- **解析方式**: 使用 `unzipper` 解压 PPTX，`xml2js` 解析 XML
- **MCP SDK**: @modelcontextprotocol/sdk
- **输入验证**: Zod

### 使用场景

1. **文档分析**: 快速提取 PPT 中的文字内容进行分析
2. **内容迁移**: 将 PPT 内容转换为其他格式
3. **自动化处理**: 批量处理 PPT 文件
4. **AI 辅助**: 让 AI 理解 PPT 内容并提供反馈
### 注意事项

- 仅支持 `.pptx` 格式（不支持旧版 `.ppt`）
- 文件路径必须为绝对路径
- 大型演示文稿响应可能被截断，建议使用 `ppt_get_slide` 分页读取
