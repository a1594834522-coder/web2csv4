# Web2CSV Server

这是 Web2CSV Chrome 扩展的服务器端组件，用于处理飞书文档的导出请求。

## 架构概述

采用前后端分离的架构：
- **扩展端**：只负责发送当前页面的URL到服务器
- **服务器端**：处理所有飞书API调用和文档导出逻辑

## 功能特性

- 接收来自扩展端的飞书文档URL
- 处理飞书API认证和令牌管理
- 创建和管理文档导出任务
- 返回最终的可下载文件URL

## 安装和运行

1. 安装依赖：
```bash
cd server
npm install
```

2. 启动服务器：
```bash
# 开发模式
npm run dev

# 生产模式
npm start
```

服务器将在 `http://localhost:3000` 启动。

## API 端点

### POST /api/export
接收来自扩展端的导出请求

**请求体：**
```json
{
  "url": "https://example.feishu.cn/docx/xxxx",
  "timestamp": 1234567890
}
```

**响应：**
```json
{
  "success": true,
  "downloadUrl": "https://open.feishu.cn/...",
  "message": "Export successful"
}
```

### GET /api/health
健康检查端点

## 环境变量

可以设置以下环境变量：

- `PORT` - 服务器端口（默认：3000）
- `FEISHU_APP_ID` - 飞书应用ID
- `FEISHU_APP_SECRET` - 飞书应用密钥

## 注意事项

1. 服务器会缓存飞书访问令牌，避免频繁请求
2. 令牌缓存时间为2小时（提前5分钟过期）
3. 导出任务最多轮询30次，每次间隔2秒
4. 临时文件存储在 `temp` 目录下

## 错误处理

- 400：无效的URL或参数
- 500：服务器内部错误或飞书API错误
- 所有错误响应都包含详细的错误信息