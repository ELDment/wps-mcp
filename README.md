<h1 align="center">WPS MCP</h1>

<p align="center">
  <a href="https://pypi.org/project/wps-mcp/"><img src="https://img.shields.io/pypi/v/wps-mcp?style=flat-square&amp;color=2563eb" alt="PyPI 版本"></a>
  <a href="#快速开始"><img src="https://img.shields.io/badge/Python-3.11%2B-3776ab?style=flat-square" alt="Python 3.11 及以上，Linux 需 CPython 3.11"></a>
  <a href="#快速开始"><img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Linux-475569?style=flat-square" alt="支持 Windows 和 Linux"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-AGPL--3.0--only-7c3aed?style=flat-square" alt="许可证：AGPL-3.0-only"></a>
</p>

<p align="center">
  <strong>🪄 让 AI 接管 WPS，自动处理文档、表格与 PPT</strong>
</p>

<p align="center">
  <strong>45</strong> 个工具组 &nbsp;·&nbsp; <strong>261</strong> 项操作 &nbsp;·&nbsp; 文字 / 表格 / 演示
</p>

<https://github.com/user-attachments/assets/59574cdb-c623-48a9-850d-c4488c5acacc>

---

## 核心能力

| 实时编辑 | 原生可编辑 | 预览与导出 |
| --- | --- | --- |
| 创建新文档，或接着修改已打开的文件，随时在 WPS 中查看结果。 | 文字、表格、图表和形状保留原生结构，方便继续手动调整。 | 编辑后查看页面预览，保存文档或导出 PDF。 |

**覆盖三大办公组件**

| 组件 | 典型任务 |
| --- | --- |
| **WPS 文字** | 报告排版、表格与图片、样式、修订、引用与邮件合并 |
| **WPS 表格** | 数据整理、公式、图表、透视表、条件格式与数据验证 |
| **WPS 演示** | 幻灯片制作、文字与形状排版、图片、主题、动画与导出 |

## 快速开始

### 1. 准备环境

安装 **WPS Office** 和 **[uv](https://docs.astral.sh/uv/)**，并确认平台环境满足以下要求：

| 平台 | Python | 原生环境 |
| --- | --- | --- |
| **Windows** | 3.11+ | WPS 已注册 COM 接口 |
| **Linux x86_64** | **CPython 3.11** | WPS Linux、Qt5、已登录的 X11 或带 XWayland 的桌面 |

### 2. 配置 MCP 客户端

将下文配置添加到 MCP 配置文件中，推荐通过 **[PyPI](https://pypi.org/project/wps-mcp/)** 接入。

<details open>
<summary><strong>从PyPI接入（Windows）</strong></summary>

```json
{
  "mcpServers": {
    "wps": {
      "command": "uvx",
      "args": ["wps-mcp@latest"]
    }
  }
}
```

如需使用预发布版，将 `args` 改为 `["--prerelease", "allow", "wps-mcp@latest"]`。

</details>

<details>
<summary><strong>从PyPI接入（Linux）</strong></summary>

Linux 需要指定 **CPython 3.11**：

```json
{
  "mcpServers": {
    "wps": {
      "command": "uvx",
      "args": ["--python", "3.11", "wps-mcp@latest"]
    }
  }
}
```

如需使用预发布版，将 `args` 改为 `["--python", "3.11", "--prerelease", "allow", "wps-mcp@latest"]`。

</details>

<details>
<summary><strong>从源码接入（可选）</strong></summary>

如需本地开发，克隆项目后按对应平台安装依赖，并将配置中的 `项目的绝对路径` 替换为实际路径。

<details>
<summary><strong>Windows</strong></summary>

在项目根目录运行：

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
        "--directory", "项目的绝对路径",
        "run", "wps-mcp"
      ]
    }
  }
}
```

</details>

<details>
<summary><strong>Linux</strong></summary>

在项目根目录运行：

```sh
uv sync --locked --python 3.11
```

在 MCP 配置文件中添加：

```json
{
  "mcpServers": {
    "wps": {
      "command": "uv",
      "args": [
        "--directory", "项目的绝对路径",
        "run", "--python", "3.11", "wps-mcp"
      ]
    }
  }
}
```

</details>
</details>

---

<p align="right">
  <sub>论坛 · <a href="https://linux.do/t/topic/3064793" title="LINUX DO 社区">LINUX DO</a></sub>
</p>
