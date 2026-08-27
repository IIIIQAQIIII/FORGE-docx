# FORGE 安装说明

**Format-Oriented Rendering & Generation Engine**

> Forge documents that don’t break.

本文件面向 FORGE 公开版本使用者。

## 1. 获取项目

可以直接 Clone 公开仓库，也可以下载对应版本的 ZIP。

```bash
git clone https://github.com/IIIIQAQIIII/forge-docx.git
cd forge-docx
```

## 2. 安装运行环境

### macOS / Linux

```bash
bash install_mcp.sh
```

### Windows PowerShell

```powershell
powershell -ExecutionPolicy Bypass -File .\install_mcp.ps1
```

安装器会：

1. 检查 Python 3.10+；
2. 创建独立 `.venv`；
3. 安装运行依赖；
4. 创建 `outputs/`；
5. 加载 FORGE MCP Server 做轻量自检；
6. 如检测到 Codex，自动注册 `forge-docx`；
7. 如检测到 DeepSeek Harness / DSH，生成 FORGE Harness MCP 配置片段；
8. 对其他客户端输出标准 stdio `command` + `args`。

## 3. Codex

安装器检测到 `codex` 后会自动注册 `forge-docx`。安装后重启 Codex 即可。

## 4. DeepSeek Harness / DSH

安装器会在项目目录生成：

```text
deepseek-harness-forge-docx.yml
```

将其中的 plugin 配置加入当前 Harness profile 的 `cordis.yml`，然后重新加载或重启 DSH。

## 5. 其他 stdio MCP 客户端

使用安装器最终打印出的两项配置：

```text
command = <项目目录>/.venv/.../python
args    = <项目目录>/server.py
```

只要客户端支持标准本地 stdio MCP，即可启动同一个 FORGE `server.py`。

## 6. 输出文件

默认输出目录：

```text
outputs/
```

调用工具时也可以明确指定其他输出路径。

## 7. 更新

Git Clone 用户：

```bash
git pull --ff-only
bash install_mcp.sh
```

ZIP 用户：将新版文件覆盖到固定安装目录后重新运行安装器。建议保留自己的 `outputs/`。

## 8. 公开版本说明

公开版本中的机构、人名和示例均为虚构或脱敏内容。若你有自己的机构模板，可以在本地替换或扩展 `templates/`，但不要将含有敏感信息的模板提交到公开仓库。
