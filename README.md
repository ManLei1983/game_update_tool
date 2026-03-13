# game_tool 第一版

`game_tool` 是一个启动前检查工具。

它现在能做 4 件事：

- 拉取 `local_report` 的 `bootstrap`
- 拉取资源清单 `manifest`
- 把返回的数据写到本地文件
- 按需下载资源，并可选启动 `QianNian.exe`

## 你现在最需要理解的 3 个命令

### 1. 初始化

```bash
python game_tool.py init
```

作用：

- 生成 `game_tool_config.json`
- 创建 `cache/`、`downloads/`、`backups/`、`runtime/` 目录

### 2. 只测试能不能连上 `local_report`

```bash
python game_tool.py bootstrap
```

作用：

- 调 `GET /api/bootstrap`
- 打印当前 VM 的任务摘要

### 3. 真正同步配置和资源

```bash
python game_tool.py sync
```

作用：

- 拉 `bootstrap`
- 拉 `manifest`
- 写本地配置文件
- 下载资源文件

## 文件说明

- `game_tool.py`
  主脚本
- `game_tool_config.example.json`
  示例配置
- `game_tool_config.json`
  实际配置，运行 `init` 后生成

## 新手测试步骤

下面按最简单的方式来。

### 第一步：准备 `local_report`

1. 先确保 `local_report` 正在运行。
2. 打开浏览器进入：

```text
http://127.0.0.1:18080/console
```

如果你设置了 `AUTH_TOKEN`，就用：

```text
http://127.0.0.1:18080/console?auth_token=你的token
```

3. 在配置台里先只填这几个字段：

- `Agent ID`: `VM-3-1`
- `区服`: `97区`
- `开始 Group`: `32`
- `结束 Group`: `80`

4. 点击 `保存 VM 配置`

保存成功后，点 `预览 Bootstrap`，如果浏览器能看到 JSON，就说明服务端这一步没问题。

### 第二步：初始化 `game_tool`

在 `D:\MyProjects\Python_Projects\my_tools\game_update_tool` 目录打开终端，执行：

```bash
python game_tool.py init
```

执行后你会看到这些内容：

- `game_tool_config.json`
- `cache/`
- `downloads/`
- `backups/`
- `runtime/`

### 第三步：修改配置文件

打开：

`D:\MyProjects\Python_Projects\my_tools\game_update_tool\game_tool_config.json`

至少确认这几个字段：

```json
{
  "server": {
    "base_url": "http://127.0.0.1:18080",
    "agent_id": "VM-3-1",
    "auth_token": "",
    "use_query_token": false,
    "timeout_seconds": 15
  }
}
```

如果 `local_report` 配了 `AUTH_TOKEN`，你就把：

```json
"auth_token": "你的token"
```

填进去。

### 第四步：测试连接

执行：

```bash
python game_tool.py bootstrap
```

如果成功，你会看到类似输出：

```text
Agent ID: VM-3-1
任务启用: True
区服: 97区
Group: 32 -> 80
```

如果这里失败，优先检查：

- `local_report` 是否启动了
- `base_url` 是否写对了
- `agent_id` 是否和配置台里一致
- `auth_token` 是否填对了

### 第五步：测试同步本地文件

执行：

```bash
python game_tool.py sync
```

成功后，你应该能在 `runtime/` 里看到：

- `bootstrap.json`
- `launch.json`
- `task_payload.json` 或 `task_payload.txt`
- `manifest.json`（如果服务端有资源清单）

先重点看：

- `runtime/bootstrap.json`

如果这个文件里有你刚才在 Web 配置台里填的 `97区`、`32`、`80`，就说明第一版已经跑通了。

## 资源下载测试方法

这是第二阶段测试，先在基础连接跑通后再做。

### 方式

1. 先准备一个可以下载的小文件，比如 `hello.txt`
2. 让这个文件能通过 HTTP 被访问到
3. 在 `local_report` 配置台里添加资源条目
4. 再执行 `python game_tool.py sync`

### 最简单的测试例子

假设你新建一个目录：

```text
D:\MyProjects\Python_Projects\my_tools\game_update_tool\sample_files
```

然后在里面手工放一个文件：

```text
hello.txt
```

内容随便写，例如：

```text
this is a test file
```

然后在这个目录开一个终端，运行：

```bash
python -m http.server 9001
```

这样你的文件就能被访问：

```text
http://127.0.0.1:9001/hello.txt
```

接着到 `local_report` 配置台，在“资源管理”里填一条：

- 资源名称：`hello.txt`
- 资源类型：`config`
- 资源版本：`v1`
- 目标路径：`runtime/hello.txt`
- 下载地址：`http://127.0.0.1:9001/hello.txt`

保存后，再执行：

```bash
python game_tool.py sync
```

如果成功，你应该能在这里看到文件：

```text
D:\MyProjects\Python_Projects\my_tools\game_update_tool\runtime\hello.txt
```

## 启动 EXE

如果你已经把 `QianNian.exe` 放在 `game_update_tool` 目录下，可以执行：

```bash
python game_tool.py launch
```

或者：

```bash
python game_tool.py run
```

区别：

- `launch`：直接启动
- `run`：先 `sync`，再启动

## 第一版限制

第一版是故意做小的，所以有这些限制：

- 资源 zip 包只下载，不自动解压
- EXE 自动更新只建议直接替换 `.exe` 文件
- 不做热更新
- 不做复杂模板写入，只先把服务端返回内容写成标准 JSON / 文本文件

## 建议你现在怎么用

第一天只做下面这件事：

1. 在 `local_report` 里录入一台 VM 配置
2. 跑 `python game_tool.py bootstrap`
3. 跑 `python game_tool.py sync`
4. 打开 `runtime/bootstrap.json` 看内容是不是你想要的

只要这一步成功，后面再接资源下载和 EXE 启动就会轻松很多。
