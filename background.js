
function extractSheetInfo(url) {
  try {
    const sheetIdMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    const gidMatch = url.match(/#gid=([0-9]+)/);

    const sheetId = sheetIdMatch ? sheetIdMatch[1] : null;
    const gid = gidMatch ? gidMatch[1] : '0';

    if (!sheetId) {
      return null;
    }

    return {
      sheetId,
      gid,
      downloadUrl: `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx&gid=${gid}`
    };
  } catch (error) {
    console.error('Error extracting sheet info:', error);
    return null;
  }
}


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

    const response = await fetch('https://open.feishu.cn/open-apis/drive/v1/export_tasks', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json; charset=utf-8'
      },
      body: JSON.stringify(body)
    });

    const data = await response.json();
    console.log('Export task creation response:', data);

    if (data.code === 0) {
      return data.data.ticket;
    } else {
      throw new Error(`Failed to create export task: ${data.msg} (code: ${data.code})`);
    }
  } catch (error) {
    console.error('Error creating Feishu export task:', error);
    throw error;
  }
}

async function checkFeishuExportTask(ticket, token, accessToken) {
  try {
    const response = await fetch(`https://open.feishu.cn/open-apis/drive/v1/export_tasks/${ticket}?token=${token}`, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    const data = await response.json();

    if (data.code === 0) {
      return data.data.result;
    } else {
      throw new Error(data.msg || 'Failed to check export task');
    }
  } catch (error) {
    console.error('Error checking Feishu export task:', error);
    throw error;
  }
}

async function getFeishuDownloadUrl(fileToken, accessToken) {
  try {
    // Feishu的实际下载需要通过API调用，返回的是文件内容，不是URL
    // 这里返回实际的API端点，这是真实的下载地址
    return `https://open.feishu.cn/open-apis/drive/v1/export_tasks/file/${fileToken}/download?access_token=${accessToken}`;
  } catch (error) {
    console.error('Error getting Feishu download URL:', error);
    throw error;
  }
}

async function downloadFeishuFile(fileToken, accessToken, filename) {
  try {
    const response = await fetch(`https://open.feishu.cn/open-apis/drive/v1/export_tasks/file/${fileToken}/download`, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    if (!response.ok) {
      throw new Error('Failed to download file');
    }

    const blob = await response.blob();
    console.log('File downloaded, blob size:', blob.size, 'type:', blob.type);

    // Convert blob to base64
    const reader = new FileReader();
    return new Promise((resolve, reject) => {
      reader.onload = function() {
        const base64data = reader.result.split(',')[1]; // Remove data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64, prefix
        const dataUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64data}`;

        chrome.downloads.download({
          url: dataUrl,
          filename: filename,
          saveAs: true
        }, (downloadId) => {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
          } else {
            console.log('Download started with ID:', downloadId);
            resolve(true);
          }
        });
      };

      reader.onerror = function() {
        reject(new Error('Failed to read file blob'));
      };

      reader.readAsDataURL(blob);
    });
  } catch (error) {
    console.error('Error downloading Feishu file:', error);
    throw error;
  }
}

async function getFeishuAccessToken(appId, appSecret) {
  try {
    console.log('Getting Feishu access token with:', { appId, appSecret: appSecret.substring(0, 8) + '...' });

    const response = await fetch('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8'
      },
      body: JSON.stringify({
        app_id: appId,
        app_secret: appSecret
      })
    });

    const data = await response.json();
    console.log('Feishu API Response:', data);
    console.log('Response type:', typeof data);
    console.log('Response keys:', Object.keys(data));
    console.log('Has tenant_access_token:', 'tenant_access_token' in data);
    console.log('tenant_access_token value:', data.tenant_access_token);

    if (data && data.code === 0 && data.tenant_access_token) {
      return data.tenant_access_token;
    } else {
      console.error('Feishu API Error Response:', data);
      console.error('Code:', data.code);
      console.error('Message:', data.msg);
      console.error('Has tenant_access_token:', !!data.tenant_access_token);
      throw new Error(data.msg || `Failed to get access token (code: ${data.code})`);
    }
  } catch (error) {
    console.error('Error getting Feishu access token:', error);
    throw error;
  }
}


// 飞书应用凭据（系统内置）
const FEISHU_CONFIG = {
  appId: 'cli_a8588718f4af901c',  // 替换为你的实际App ID
  appSecret: 'RlUiFV5bx7KldWeMUVT8rgqc7ynMWvYO'  // 替换为你的实际App Secret
};

async function getFeishuAccessToken(appId, appSecret) {
  try {
    console.log('Getting Feishu access token with:', { appId, appSecret: appSecret.substring(0, 8) + '...' });

    const response = await fetch('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8'
      },
      body: JSON.stringify({
        app_id: appId,
        app_secret: appSecret
      })
    });

    const data = await response.json();
    console.log('Feishu API Response:', data);

    if (data && data.code === 0 && data.tenant_access_token) {
      return data.tenant_access_token;
    } else {
      throw new Error(data.msg || `Failed to get access token (code: ${data.code})`);
    }
  } catch (error) {
    console.error('Error getting Feishu access token:', error);
    throw error;
  }
}

async function exportFeishuDocument(feishuInfo, accessToken = null) {
  // 如果没有提供accessToken，使用内置凭据获取
  if (!accessToken) {
    accessToken = await getFeishuAccessToken(FEISHU_CONFIG.appId, FEISHU_CONFIG.appSecret);
  }
  try {
    console.log('Starting Feishu document export:', feishuInfo);

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
        // Still processing, continue waiting
        console.log('Export still in progress, attempt:', attempts + 1);
      } else if (result.job_status === 3 || result.job_status === 107 || result.job_status === 108 || result.job_status === 109) {
        throw new Error(`Export failed: ${result.job_error_msg || 'Unknown error'} (status: ${result.job_status})`);
      }

      attempts++;
    }

    if (result.job_status !== 0) {
      throw new Error(`Export timeout or failed. Final status: ${result.job_status}, message: ${result.job_error_msg || 'Unknown'}`);
    }

    const filename = `feishu_${feishuInfo.type}_${feishuInfo.token}.xlsx`;
    console.log('Downloading file with filename:', filename);

    // 获取真实的下载URL
    const downloadUrl = await getFeishuDownloadUrl(result.file_token, accessToken);

    // 启动下载
    await downloadFeishuFile(result.file_token, accessToken, filename);

    return { success: true, downloadUrl: downloadUrl };
  } catch (error) {
    console.error('Error in Feishu document export:', error);
    throw error;
  }
}

async function getFeishuSheetId(spreadsheetToken, accessToken) {
  try {
    const response = await fetch(`https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/${spreadsheetToken}/sheets/query`, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    const data = await response.json();
    console.log('Sheet ID query response:', data);

    if (data.code === 0 && data.data && data.data.sheets && data.data.sheets.length > 0) {
      // Return the first sheet's ID
      return data.data.sheets[0].sheet_id;
    } else {
      throw new Error(`Failed to get sheet ID: ${data.msg || 'No sheets found'}`);
    }
  } catch (error) {
    console.error('Error getting Feishu sheet ID:', error);
    throw error;
  }
}

// 下载带有签名的文件
async function downloadSignedFile(url, headers, filename) {
  try {
    // 在Service Worker中，我们直接使用chrome.downloads.download
    // 但是需要先获取文件内容并转换为data URL
    const response = await fetch(url, {
      method: 'GET',
      headers: headers
    });

    if (!response.ok) {
      throw new Error(`下载失败: ${response.status} ${response.statusText}`);
    }

    // 获取文件内容
    const blob = await response.blob();

    // 转换为base64
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = function() {
        const base64data = reader.result.split(',')[1]; // 移除data:...;base64,前缀
        const dataUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64data}`;

        chrome.downloads.download({
          url: dataUrl,
          filename: filename,
          saveAs: true
        }, (downloadId) => {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
          } else {
            resolve(downloadId);
          }
        });
      };

      reader.onerror = function() {
        reject(new Error('Failed to read file blob'));
      };

      reader.readAsDataURL(blob);
    });
  } catch (error) {
    console.error('下载签名文件失败:', error);
    throw error;
  }
}

// 服务器配置
const SERVER_CONFIG = {
  url: 'http://localhost:3000/api/export',  // 服务器地址
  timeout: 30000  // 30秒超时
};

// 发送URL到服务器进行处理
async function sendUrlToServer(url) {
  try {
    console.log('Sending URL to server:', url);

    const response = await fetch(SERVER_CONFIG.url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        url: url,
        timestamp: Date.now()
      })
    });

    if (!response.ok) {
      throw new Error(`Server response: ${response.status} ${response.statusText}`);
    }

    const result = await response.json();
    console.log('Server response:', result);

    if (result.success && result.fileToken) {
      return result;
    } else {
      throw new Error(result.error || 'Server processing failed');
    }
  } catch (error) {
    console.error('Error sending URL to server:', error);
    throw error;
  }
}

// 消息监听器
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  (async () => {
    try {
      if (request.action === 'extractSheetInfo') {
        // 获取当前活动标签页的URL
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        const url = tab ? tab.url : '';
        console.log('Extracting sheet info from URL:', url);

        // 检查是否为Google Sheets
        if (url.includes('docs.google.com/spreadsheets')) {
          const sheetInfo = extractSheetInfo(url);
          if (sheetInfo) {
            sendResponse({ success: true, type: 'google', sheetInfo });
            return;
          }
        }

        // 检查是否为飞书文档
        if (url.includes('feishu.cn')) {
          const feishuInfo = extractFeishuInfo(url);
          if (feishuInfo) {
            sendResponse({ success: true, type: 'feishu', feishuInfo });
            return;
          }
        }

        sendResponse({ success: false });
      }

      else if (request.action === 'downloadSheet') {
        console.log('Downloading sheet from URL:', request.downloadUrl);
        await downloadSignedFile(request.downloadUrl, {}, request.filename);
        sendResponse({ success: true });
      }

      else if (request.action === 'exportFeishuDocument') {
        console.log('Sending Feishu URL to server...');
        try {
          // 发送URL到服务器处理
          const result = await sendUrlToServer(request.feishuInfo.url);

          // 获取新的访问令牌
          const accessToken = await getFeishuAccessToken(FEISHU_CONFIG.appId, FEISHU_CONFIG.appSecret);

          // 生成下载URL（不包含access_token，因为要在header中发送）
          const downloadUrl = `https://open.feishu.cn/open-apis/drive/v1/export_tasks/file/${result.fileToken}/download`;

          // 生成文件名
          const filename = `feishu_${request.feishuInfo.type}_${request.feishuInfo.token}.xlsx`;

          // 下载文件，带有Authorization头
          await downloadSignedFile(downloadUrl, {
            'Authorization': `Bearer ${accessToken}`
          }, filename);

          sendResponse({
            success: true,
            downloadUrl: downloadUrl
          });
        } catch (error) {
          sendResponse({ success: false, error: error.message });
        }
      }

      else if (request.action === 'ping') {
        sendResponse({ success: true });
      }


    } catch (error) {
      console.error('Error handling message:', error);
      sendResponse({ success: false, error: error.message });
    }
  })();

  return true; // 保持消息通道开放以支持异步响应
});

console.log('🔧 Background Script已加载');