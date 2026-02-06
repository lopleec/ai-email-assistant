# AI邮件自动回复助手

一个基于Python的智能邮件助手，可以自动读取Outlook邮箱的未读邮件，使用OpenAI兼容API生成智能回复并自动发送。

##文件过大上传不方便，解压venv文件，放在同目录下！！！！

## 功能特点

- ✅ OAuth2网页登录，无需密码（安全）
- ✅ 自动监控Outlook邮箱的未读邮件
- ✅ 支持OpenAI兼容API（可自定义base_url）
- ✅ 使用AI生成智能、专业的回复
- ✅ 自动发送回复并标记邮件为已读
- ✅ 可配置的检查间隔和AI参数
- ✅ 完整的日志记录
- ✅ 错误处理和自动重试
- ✅ Token自动刷新机制

## 安装步骤

### 1. 克隆或下载项目

```bash
cd ~/Documents/Programing☕️/ai-email-assistant
```

### 2. 创建虚拟环境（推荐）

```bash
python3 -m venv venv
source venv/bin/activate  # Mac/Linux
# 或者
venv\Scripts\activate  # Windows
```

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

### 4. 配置环境变量

复制 `.env.example` 为 `.env`：

```bash
cp .env.example .env
```

然后编辑 `.env` 文件，填入你的配置：

```ini
# Outlook邮箱配置 (OAuth2方式，无需密码)
EMAIL_ADDRESS=your_email@outlook.com

# OpenAI兼容API配置
OPENAI_API_KEY=your_openai_api_key
OPENAI_BASE_URL=https://api.openai.com/v1

# 邮件检查间隔（秒）默认60秒=1分钟
CHECK_INTERVAL=60

# AI回复设置
AI_MODEL=gpt-3.5-turbo
AI_TEMPERATURE=0.7
REPLY_LANGUAGE=zh-CN
```

## 配置说明

### Outlook邮箱设置（OAuth2）

✨ **新特性**：使用OAuth2网页登录，无需密码，更安全！

#### 首次运行时：

1. 程序会自动打开浏览器
2. 显示一个设备代码（例如：ABCD-1234）
3. 在浏览器中输入该代码
4. 使用你的Outlook账户登录并授权
5. 授权后，程序会自动获取并保存访问令牌
6. Token会自动刷新，无需重复授权

#### 优势：
- 🔐 更安全：不需要存储密码
- 🔄 自动刷新：Token过期会自动更新
- 🚀 更简单：无需创建应用密码

### OpenAI兼容API设置

✨ **支持所有OpenAI兼容的API服务**

#### 官方OpenAI：
```ini
OPENAI_API_KEY=sk-...
OPENAI_BASE_URL=https://api.openai.com/v1
```

#### 其他兼容服务（例如）：
```ini
# Azure OpenAI
OPENAI_BASE_URL=https://your-resource.openai.azure.com/

# 本地模型（Ollama）
OPENAI_BASE_URL=http://localhost:11434/v1

# 其他第三方服务
OPENAI_BASE_URL=https://api.your-service.com/v1
```

只要是支持OpenAI API格式的服务都可以使用！

### 参数调整

- `CHECK_INTERVAL`: 检查邮件的间隔（秒），建议60-600秒
- `AI_MODEL`: 模型名称，根据你使用的API服务而定
- `AI_TEMPERATURE`: 控制回复创意度（0-1），0.7为推荐值
- `REPLY_LANGUAGE`: 回复语言，如 `zh-CN`、`en-US` 等
- `OPENAI_BASE_URL`: API服务地址，支持任何OpenAI兼容的服务

## 使用方法

### 启动助手

```bash
python email_assistant.py
```

### 后台运行（Mac/Linux）

```bash
nohup python email_assistant.py > output.log 2>&1 &
```

### 停止运行

按 `Ctrl+C` 或使用：

```bash
pkill -f email_assistant.py
```

## 日志查看

所有操作都会记录在 `email_assistant.log` 文件中：

```bash
tail -f email_assistant.log
```

## 工作流程

1. 程序启动时，会弹出浏览器进行OAuth2授权（仅首次）
2. 授权成功后，按设定的间隔（默认1分钟）检查邮箱
3. 发现未读邮件时，读取发件人、主题和正文
4. 调用OpenAI兼容API生成智能回复
5. 自动发送回复邮件
6. 将原邮件标记为已读
7. 继续等待下一次检查

## 注意事项

1. **首次运行**：需要在浏览器中授权，跟随提示完成即可
2. **API费用**：OpenAI API按使用量计费，建议设置合理的检查间隔
3. **隐私安全**：不要将 `.env` 和 `outlook_token.json` 文件提交到代码仓库
4. **邮件过滤**：可以在代码中添加过滤规则，只回复特定邮件
5. **回复质量**：可以调整提示词（prompt）来优化AI回复的风格和内容
6. **兼容API**：支持任何OpenAI兼容的API服务，包括本地模型

## 自定义回复风格

在 `email_assistant.py` 中的 `generate_ai_reply` 方法里，你可以修改提示词（prompt）来调整回复风格，例如：

- 更正式的商务风格
- 更友好的客服风格
- 特定行业的专业术语
- 添加公司签名等

## 故障排除

### 无法连接邮箱
- 检查网络连接
- 确认邮箱地址正确
- 尝试重新授权：删除 `outlook_token.json` 文件后重新运行

### OAuth2授权失败
- 确保浏览器能正常访问 login.microsoftonline.com
- 检查是否正确输入了设备代码
- 确认使用的是Outlook/Hotmail/Microsoft账户

### OpenAI API错误
- 检查API密钥是否有效
- 确认 `OPENAI_BASE_URL` 配置正确
- 检查网络是否能访问API服务
- 如使用第三方服务，确认服务可用

### 程序崩溃
- 查看 `email_assistant.log` 日志文件
- 确认所有依赖已正确安装
- 检查Python版本（需要3.7+）

## 许可证

MIT License

## 贡献

欢迎提交问题和改进建议！
