# WPS MCP

**让 AI 直接操作你正在使用的 WPS。**

[演示视频](assets/showcase.mp4)

## 核心能力

- **实时编辑**：创建新文档，也能接着修改已打开的文件，随时在 WPS 中查看结果。
- **原生可编辑**：文字、表格、图表和形状保留原生结构，方便继续手动调整。
- **预览与导出**：编辑后查看页面预览，保存文档或导出 PDF。

覆盖 **45 个工具组、261 项操作**，重点支持：

| 组件 | 典型任务 |
| --- | --- |
| WPS 文字 | 报告排版、表格与图片、样式、修订、引用与邮件合并 |
| WPS 表格 | 数据整理、公式、图表、透视表、条件格式与数据验证 |
| WPS 演示 | 幻灯片制作、文字与形状排版、图片、主题、动画与导出 |

## 快速开始

需要 **Windows（Linux 和 macOS 兼容版仍在开发中）、Python 3.11+、WPS Office（已注册 COM 接口）** 和 **[uv](https://docs.astral.sh/uv/)**。

推荐通过 ***PyPI*** 接入，`uvx` 会自动安装并启动服务。**在 MCP 配置文件中添加**：

```json
{
  "mcpServers": {
    "wps": {
      "command": "uvx",
      "args": ["wps-mcp"]
    }
  }
}
```

<details>
<summary>本地接入</summary>

克隆项目后，在项目根目录安装依赖：

```powershell
uv sync --locked
```

在 MCP 配置文件中添加：

```json
{
  "mcpServers": {
    "wps": {
      "command": "uv",
      "args": [
        "--directory", "/path/to/wps-mcp", // 替换为真实项目路径
        "run", "wps-mcp"
      ]
    }
  }
}
```

</details>

## 许可证

本项目采用 **GNU Affero General Public License v3.0（仅第 3 版，SPDX：`AGPL-3.0-only`）**。
完整条款见 [LICENSE](LICENSE)。
