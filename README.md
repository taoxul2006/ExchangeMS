# Exchange 目录说明

## 1. 目录用途

本目录是一套可直接拷贝到域环境机器使用的本地 Exchange 可视化操作站。

作用有三类：

1. 提供本地网页界面，用于填写连接参数、搜索条件、目标账号和删除确认。
2. 通过 PowerShell 本地服务调用 Exchange 远程 PowerShell，会话连接、搜索邮件、删除邮件。
3. 保存最少量的本地配置和账号列表，便于重复使用。

## 2. 推荐保留的目录结构

```text
exchange/
|-- exchange-web-server.ps1
|-- exchange-web-config.json
|-- leaverlist.ps1
|-- start-exchange-web.ps1
|-- start-exchange-web.cmd
|-- README.md
`-- web-ui/
    |-- index.html
    |-- styles.css
    `-- app.js
```

## 3. 各文件作用

| 文件 | 作用 | 是否运行必需 |
|---|---|---|
| `exchange-web-server.ps1` | 本地 Web 服务后端，负责提供页面、处理 API、连接 Exchange、执行搜索和删除 | 是 |
| `exchange-web-config.json` | 默认连接配置，提供用户名、连接地址、认证方式、目标文件夹默认值 | 是 |
| `leaverlist.ps1` | 目标账号列表，页面会读取和保存这个文件 | 是 |
| `start-exchange-web.ps1` | PowerShell 启动入口 | 是 |
| `start-exchange-web.cmd` | 双击启动入口，给不习惯 PowerShell 的使用者使用 | 是 |
| `README.md` | 目录说明、运行方式、依赖关系、优化建议 | 建议保留 |
| `web-ui/index.html` | 页面结构 | 是 |
| `web-ui/styles.css` | 页面样式 | 是 |
| `web-ui/app.js` | 页面交互逻辑，负责调用后端 API、渲染结果 | 是 |

## 4. 文件之间的直接关联

### 4.1 启动关联

```text
start-exchange-web.cmd
        or
start-exchange-web.ps1
        |
        v
exchange-web-server.ps1
```

### 4.2 配置关联

```text
exchange-web-server.ps1
        |
        +-- 读取 exchange-web-config.json
        |
        `-- 读取 / 写入 leaverlist.ps1
```

### 4.3 前后端关联

```text
web-ui/index.html
web-ui/styles.css
web-ui/app.js
        |
        v
exchange-web-server.ps1
        |
        v
Exchange Remote PowerShell
```

### 4.4 页面与 API 的直接映射

| 页面动作 | 对应接口 | 对应后端动作 |
|---|---|---|
| 打开页面 | `GET /` | 返回 `index.html` |
| 页面初始化 | `GET /api/state` | 返回默认配置与连接状态 |
| 读取账号列表 | `GET /api/users` | 读取 `leaverlist.ps1` |
| 保存账号列表 | `POST /api/users` | 写回 `leaverlist.ps1` |
| 连接 Exchange | `POST /api/connect` | 建立 Exchange 会话并验证 `Search-Mailbox` |
| 断开状态 | `POST /api/disconnect` | 清理本地会话状态 |
| 搜索邮件 | `POST /api/search` | 执行搜索并复制到目标邮箱 |
| 删除邮件 | `POST /api/delete` | 按当前条件执行删除 |

## 5. 运行顺序

1. 在域环境机器上双击 `start-exchange-web.cmd` 或运行 `start-exchange-web.ps1`。
2. 浏览器打开 `http://127.0.0.1:3080`。
3. 页面先从 `exchange-web-config.json` 读取默认连接参数。
4. 页面从 `leaverlist.ps1` 读取目标账号列表。
5. 用户在页面中连接 Exchange。
6. 先搜索，再确认，再删除。

## 6. 已删除的非必要文件

为了让目录更适合直接交付和迁移，以下文件不再建议保留在运行目录中：

- 原始操作手册 `docx`
- 早期静态说明页 `exchange-mail-visual-guide.html`
- 旧的手工脚本 `recallmail.ps1`
- 邮箱地址清洗脚本 `zzmailaddress.ps1`
- 运行期生成的 `exchange-web-server.log`
- 仅用于解析默认值的旧连接脚本 `linkexchange.ps1`

这些文件不是当前网站直接运行所必需的，保留反而会让目录边界不清晰。

## 7. 当前目录的优化结果

本次已经完成的优化：

1. 目录只保留当前网站直接运行所需文件。
2. 默认连接参数改为从 `exchange-web-config.json` 读取，不再依赖解析旧脚本。
3. 启动脚本和服务脚本全部使用相对路径，可整体拷贝到其他机器。
4. 运行日志改为运行期自动生成，不作为交付文件的一部分。

## 8. 仍可继续优化的空间

下面这些是后续还可以继续做，但当前不是必须项：

1. 将 `exchange-web-config.json` 的敏感字段拆分成“样例配置”和“现场配置”，避免误传真实账号信息。
2. 增加“连接诊断”按钮，自动检测域登录、DNS、端口、WinRM 和 Exchange 可达性。
3. 增加搜索结果导出功能，方便留痕。
4. 增加删除前二次确认页，先展示预计命中数量再执行删除。
5. 将后端 API 和 Exchange 调用进一步拆成独立函数文件，便于维护。

## 9. 使用提醒

- 如果使用 `Kerberos`，目标机器应已加入域，并能访问域控和 Exchange。
- 如果目标机器无法使用域凭据，页面可尝试切换到 `Negotiate` 或 `Basic`，前提是 Exchange 侧允许。
- 复制目录时请保持整个目录结构不变。
