# Outlook / Hotmail 提取包

从 `GuJumpgate-v0.1.4` 中提取的 Outlook / Hotmail 相关模块。

## 文件说明

- `hotmail-utils.js`：Hotmail/Outlook 账号池、验证码提取、邮件 API 响应归一化、别名容量等纯工具函数。
- `microsoft-email.js`：Microsoft Graph / Outlook API 邮件读取模块，支持 refresh token 换 access token、读取 Inbox/Junk 邮件、提取验证码。
- `scripts/hotmail_helper.py`：本地 Hotmail helper 服务，提供 `/messages`、`/code`、账号记录同步等接口。
- `start-hotmail-helper.bat`：Windows 启动脚本，默认端口 `17373`。
- `start-hotmail-helper.command`：macOS 启动脚本。
- `sidepanel/hotmail-manager.js`：原项目侧栏 Hotmail 账号池管理器。
- `sidepanel/account-pool-ui.js`：Hotmail/2925 账号池表单显隐的共享 UI helper；`hotmail-manager.js` 依赖它。
- `data/`：本地 helper 写入账号运行记录快照时使用的目录。

## 快速启动本地 helper

Windows：

```bat
start-hotmail-helper.bat
```

或指定端口：

```bat
start-hotmail-helper.bat 17373
```

直接运行 Python：

```bash
python scripts/hotmail_helper.py --port 17373
```

## 本地 helper 主要接口

### POST `/messages`

请求示例：

```json
{
  "email": "your@outlook.com",
  "clientId": "your-client-id",
  "refreshToken": "your-refresh-token",
  "mailboxes": ["INBOX", "Junk"],
  "top": 5
}
```

返回邮件列表、transport、下一次 refresh token 等信息。

### POST `/code`

请求字段同 `/messages`，并可额外传：

```json
{
  "senderFilters": [],
  "subjectFilters": [],
  "excludeCodes": [],
  "filterAfterTimestamp": 0,
  "requiredKeywords": [],
  "codePatterns": []
}
```

返回匹配到的验证码和邮件对象。

## 注意

`sidepanel/hotmail-manager.js` 是从完整 Chrome 扩展里提取出的 UI 管理模块，它依赖原扩展的 `state`、`dom`、`helpers`、`runtime` 上下文。单独使用时，需要在调用 `createHotmailManager(context)` 时自行提供这些上下文对象。
