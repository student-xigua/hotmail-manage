# kang Hotmail 管理后台

这是一个本地运行的 Outlook / Hotmail OAuth2 邮箱管理和取件工具，支持批量导入账号、列表管理、单个取件、批量取件、导出和删除。

## 文件说明

- `web/`：前端管理页面。
- `scripts/hotmail_helper.py`：本地后端服务，提供网页、账号池 API、邮件读取 API。
- `start-hotmail-helper.bat`：Windows 启动脚本，默认端口 `17373`。
- `start-hotmail-helper.command`：macOS 启动脚本。
- `hotmail-utils.js`：Hotmail/Outlook 工具函数。
- `microsoft-email.js`：Microsoft Graph / Outlook API 邮件读取模块。
- `sidepanel/`：原 Chrome 扩展侧栏相关脚本。
- `data/`：本地运行数据目录，已被 `.gitignore` 忽略，不会上传到 GitHub。

## 平时如何使用

### 1. 进入项目目录

```powershell
cd D:\Hotmail\outlook-hotmail\outlook-hotmail
```

### 2. 启动服务

Windows 可以直接双击：

```text
start-hotmail-helper.bat
```

也可以在 PowerShell 中运行：

```powershell
.\start-hotmail-helper.bat
```

默认会启动在：

```text
http://127.0.0.1:17373/
```

如果要指定端口：

```powershell
.\start-hotmail-helper.bat 17373
```

也可以直接运行 Python：

```powershell
python scripts\hotmail_helper.py --port 17373
```

### 3. 打开管理页面

浏览器访问：

```text
http://127.0.0.1:17373/
```

### 4. 导入邮箱

每行一个账号，格式如下：

```text
邮箱----密码----client_id----refresh_token
```

示例：

```text
xxx@outlook.com----邮箱密码----client_id----refresh_token
```

支持两种导入方式：

- 拖拽或点击上传 `.txt` / `.csv` 文件。
- 点击“粘贴文本导入”，直接粘贴多行账号。

### 5. 取件

导入后，可以在表格里操作：

- `刷新邮件`：读取单个邮箱的最新邮件。
- `查看邮件`：查看最近一次读取结果；如果没有缓存，会自动读取。
- 勾选多个邮箱后点 `批量获取邮件`：批量读取邮件。

## API 模式说明

- `Graph`：Microsoft Graph API，推荐优先使用，稳定性更好。
- `Outlook`：旧 Outlook REST 接口，部分 token 兼容时可作为备用。

当前后端会按可用策略自动尝试 Graph / Outlook 的取件方案。

## 数据保存位置

账号数据保存在：

```text
D:\Hotmail\outlook-hotmail\outlook-hotmail\data\accounts.json
```

保存内容包括：

- 邮箱
- 密码
- Client ID
- Refresh Token
- 状态
- 最近取件时间
- 最近邮件数量
- 最近错误
- tokenEndpoint
- transport

注意：页面表格里会对密码、Client ID、Refresh Token 打码展示，但 `accounts.json` 中保存的是完整数据。

## 邮件内容是否保存

默认不保存邮件内容。

取得的邮件只会返回给前端临时显示，前端会把最近一次结果放在页面内存里。页面刷新后，需要重新取件。

如果需要长期保存邮件内容，可以再扩展保存到 `data/messages.json`。

## 本地 API

### GET `/api/accounts`

读取账号列表和统计信息。

### POST `/api/accounts/import`

批量导入账号。

请求示例：

```json
{
  "text": "xxx@outlook.com----password----client_id----refresh_token",
  "settings": {
    "delimiter": "----",
    "defaultTop": 1,
    "apiMode": "Graph"
  }
}
```

### POST `/api/accounts/{id}/messages`

读取指定账号邮件。

请求示例：

```json
{
  "mailboxes": ["INBOX", "Junk"],
  "top": 1
}
```

### POST `/api/accounts/batch-fetch`

批量读取邮件。

请求示例：

```json
{
  "ids": ["account_id"],
  "mailboxes": ["INBOX", "Junk"],
  "top": 1
}
```

### GET `/api/accounts/export`

导出账号文本。

### POST `/api/accounts/delete`

删除账号。

请求示例：

```json
{
  "ids": ["account_id"]
}
```

清空全部：

```json
{
  "mode": "all"
}
```

## 旧接口

为了兼容原工具，仍保留：

### POST `/messages`

```json
{
  "email": "your@outlook.com",
  "clientId": "your-client-id",
  "refreshToken": "your-refresh-token",
  "mailboxes": ["INBOX", "Junk"],
  "top": 5
}
```

### POST `/code`

```json
{
  "email": "your@outlook.com",
  "clientId": "your-client-id",
  "refreshToken": "your-refresh-token",
  "mailboxes": ["INBOX", "Junk"],
  "top": 5,
  "senderFilters": [],
  "subjectFilters": [],
  "excludeCodes": [],
  "filterAfterTimestamp": 0,
  "requiredKeywords": [],
  "codePatterns": []
}
```

## 注意事项

- `data/` 目录包含本地账号信息，默认不会上传到 GitHub。
- 不要把真实账号、密码、Client ID、Refresh Token 发给别人。
- 如果端口 `17373` 被占用，可以用 `start-hotmail-helper.bat 其他端口` 启动。
