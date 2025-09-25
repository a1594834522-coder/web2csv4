const express = require('express');
const cors = require('cors');
const axios = require('axios');
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// 中间件
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true }));

// 确保临时目录存在
const tempDir = path.join(__dirname, 'temp');
if (!fs.existsSync(tempDir)) {
  fs.mkdirSync(tempDir, { recursive: true });
}

// 配置multer用于文件上传
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, tempDir);
  },
  filename: (req, file, cb) => {
    cb(null, `${uuidv4()}-${file.originalname}`);
  }
});

const upload = multer({ storage });

// 飞书应用凭据
const FEISHU_CONFIG = {
  appId: 'cli_a8588718f4af901c',
  appSecret: 'RlUiFV5bx7KldWeMUVT8rgqc7ynMWvYO'
};

// 缓存访问令牌
let accessTokenCache = {
  token: null,
  expiresAt: 0
};

// 获取飞书访问令牌
async function getFeishuAccessToken() {
  try {
    // 如果令牌未过期，直接返回
    if (accessTokenCache.token && accessTokenCache.expiresAt > Date.now()) {
      console.log('Using cached access token');
      return accessTokenCache.token;
    }

    console.log('Getting new Feishu access token');

    const response = await axios.post('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal', {
      app_id: FEISHU_CONFIG.appId,
      app_secret: FEISHU_CONFIG.appSecret
    });

    if (response.data && response.data.code === 0 && response.data.tenant_access_token) {
      // 缓存令牌，设置2小时后过期
      accessTokenCache = {
        token: response.data.tenant_access_token,
        expiresAt: Date.now() + (2 * 60 * 60 * 1000 - 5 * 60 * 1000) // 提前5分钟过期
      };

      console.log('Access token obtained successfully');
      return accessTokenCache.token;
    } else {
      throw new Error(`Failed to get access token: ${response.data.msg || 'Unknown error'}`);
    }
  } catch (error) {
    console.error('Error getting Feishu access token:', error.response?.data || error.message);
    // Preserve Feishu API error details
    if (error.response?.data) {
      const feishuError = error.response.data;
      throw new Error(`Feishu API Error: ${feishuError.msg} (code: ${feishuError.code})`);
    }
    throw error;
  }
}

// 从URL提取飞书信息
function extractFeishuInfo(url) {
  try {
    const feishuMatch = url.match(/\/docx\/([a-zA-Z0-9_-]+)/);
    const feishuSheetMatch = url.match(/\/sheets\/([a-zA-Z0-9_-]+)/);

    const token = feishuMatch ? feishuMatch[1] : (feishuSheetMatch ? feishuSheetMatch[1] : null);
    const type = feishuSheetMatch ? 'sheet' : 'docx';

    if (!token) {
      return null;
    }

    return {
      token,
      type,
      url: url
    };
  } catch (error) {
    console.error('Error extracting Feishu info:', error);
    return null;
  }
}

// 创建导出任务
async function createFeishuExportTask(token, type, subId, accessToken) {
  try {
    const body = {
      file_extension: 'xlsx',
      token: token,
      type: type
    };

    if (subId) {
      body.sub_id = subId;
    }

    console.log('Creating export task with body:', body);

    const response = await axios.post('https://open.feishu.cn/open-apis/drive/v1/export_tasks', body, {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json; charset=utf-8'
      }
    });

    const data = response.data;
    console.log('Export task creation response:', data);

    if (data.code === 0) {
      return data.data.ticket;
    } else {
      throw new Error(`Failed to create export task: ${data.msg} (code: ${data.code})`);
    }
  } catch (error) {
    console.error('Error creating Feishu export task:', error.response?.data || error.message);
    // Preserve Feishu API error details
    if (error.response?.data) {
      const feishuError = error.response.data;
      throw new Error(`Feishu API Error: ${feishuError.msg} (code: ${feishuError.code})`);
    }
    throw error;
  }
}

// 检查导出任务状态
async function checkFeishuExportTask(ticket, token, accessToken) {
  try {
    const response = await axios.get(`https://open.feishu.cn/open-apis/drive/v1/export_tasks/${ticket}?token=${token}`, {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    const data = response.data;

    if (data.code === 0) {
      return data.data.result;
    } else {
      throw new Error(data.msg || 'Failed to check export task');
    }
  } catch (error) {
    console.error('Error checking Feishu export task:', error.response?.data || error.message);
    // Preserve Feishu API error details
    if (error.response?.data) {
      const feishuError = error.response.data;
      throw new Error(`Feishu API Error: ${feishuError.msg} (code: ${feishuError.code})`);
    }
    throw error;
  }
}

// 获取飞书表格ID
async function getFeishuSheetId(spreadsheetToken, accessToken) {
  try {
    const response = await axios.get(`https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/${spreadsheetToken}/sheets/query`, {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    const data = response.data;
    console.log('Sheet ID query response:', data);

    if (data.code === 0 && data.data && data.data.sheets && data.data.sheets.length > 0) {
      return data.data.sheets[0].sheet_id;
    } else {
      throw new Error(`Failed to get sheet ID: ${data.msg || 'No sheets found'}`);
    }
  } catch (error) {
    console.error('Error getting Feishu sheet ID:', error.response?.data || error.message);
    // Preserve Feishu API error details
    if (error.response?.data) {
      const feishuError = error.response.data;
      throw new Error(`Feishu API Error: ${feishuError.msg} (code: ${feishuError.code})`);
    }
    throw error;
  }
}

// 导出飞书文档
async function exportFeishuDocument(feishuInfo) {
  try {
    console.log('Starting Feishu document export:', feishuInfo);

    const accessToken = await getFeishuAccessToken();

    // For sheet types, we need to get the sheet/sub_id first
    let subId = null;
    if (feishuInfo.type === 'sheet') {
      console.log('Sheet type detected, getting sheet ID...');
      subId = await getFeishuSheetId(feishuInfo.token, accessToken);
      console.log('Got sheet ID:', subId);
    }

    const ticket = await createFeishuExportTask(feishuInfo.token, feishuInfo.type, subId, accessToken);
    console.log('Export task created, ticket:', ticket);

    let result;
    const maxAttempts = 30;
    let attempts = 0;

    while (attempts < maxAttempts) {
      await new Promise(resolve => setTimeout(resolve, 2000));

      result = await checkFeishuExportTask(ticket, feishuInfo.token, accessToken);
      console.log('Export task status check:', result);

      if (result.job_status === 0) {
        break;
      } else if (result.job_status === 1 || result.job_status === 2) {
        console.log('Export still in progress, attempt:', attempts + 1);
      } else if (result.job_status === 3 || result.job_status === 107 || result.job_status === 108 || result.job_status === 109) {
        throw new Error(`Export failed: ${result.job_error_msg || 'Unknown error'} (status: ${result.job_status})`);
      }

      attempts++;
    }

    if (result.job_status !== 0) {
      throw new Error(`Export timeout or failed. Final status: ${result.job_status}, message: ${result.job_error_msg || 'Unknown'}`);
    }

    return {
      success: true,
      fileToken: result.file_token,
      message: 'Export successful'
    };
  } catch (error) {
    console.error('Error in Feishu document export:', error);
    // Re-throw with preserved error details
    throw error;
  }
}

// 主API端点
app.post('/api/export', async (req, res) => {
  try {
    const { url } = req.body;

    if (!url) {
      return res.status(400).json({
        success: false,
        error: 'URL is required'
      });
    }

    console.log('Received export request for URL:', url);

    // 检查是否为飞书URL
    if (!url.includes('feishu.cn')) {
      return res.status(400).json({
        success: false,
        error: 'Only Feishu URLs are supported'
      });
    }

    // 提取飞书信息
    const feishuInfo = extractFeishuInfo(url);
    if (!feishuInfo) {
      return res.status(400).json({
        success: false,
        error: 'Invalid Feishu URL format'
      });
    }

    // 导出文档
    const result = await exportFeishuDocument(feishuInfo);

    res.json({
      success: true,
      fileToken: result.fileToken,
      message: 'Export successful'
    });

  } catch (error) {
    console.error('Error processing export request:', error);

    // Determine appropriate HTTP status code based on error type
    let statusCode = 500;
    let errorMessage = error.message || 'Internal server error';

    if (errorMessage.includes('Feishu API Error')) {
      statusCode = 400; // Bad Request for Feishu API errors
    } else if (errorMessage.includes('Invalid Feishu URL format')) {
      statusCode = 400;
    } else if (errorMessage.includes('Only Feishu URLs are supported')) {
      statusCode = 400;
    } else if (errorMessage.includes('URL is required')) {
      statusCode = 400;
    }

    res.status(statusCode).json({
      success: false,
      error: errorMessage,
      details: {
        type: error.constructor.name,
        code: error.code,
        timestamp: new Date().toISOString()
      }
    });
  }
});

// 健康检查端点
app.get('/api/health', (req, res) => {
  res.json({
    success: true,
    message: 'Server is running',
    timestamp: new Date().toISOString()
  });
});

// 启动服务器
app.listen(PORT, () => {
  console.log(`🚀 Server is running on port ${PORT}`);
  console.log(`📝 Health check: http://localhost:${PORT}/api/health`);
  console.log(`📤 Export endpoint: http://localhost:${PORT}/api/export`);
});

// 优雅关闭
process.on('SIGINT', () => {
  console.log('\n🔄 Shutting down server...');
  process.exit(0);
});

process.on('SIGTERM', () => {
  console.log('\n🔄 Shutting down server...');
  process.exit(0);
});