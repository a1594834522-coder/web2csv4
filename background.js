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
      downloadUrl: `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${gid}`
    };
  } catch (error) {
    console.error('Error extracting sheet info:', error);
    return null;
  }
}

function extractDingTalkInfo(url) {
  try {
    const dingTalkMatch = url.match(/alidocs\.dingtalk\.com\/i\/nodes\/([a-zA-Z0-9_-]+)/);
    if (!dingTalkMatch) {
      // 尝试匹配预览页面URL
      const previewMatch = url.match(/alidocs\.dingtalk\.com\/uni-preview.*dentryUuid=([a-zA-Z0-9]+)/);
      if (previewMatch) {
        return {
          dentryUuid: previewMatch[1],
          url: url,
          type: 'preview'
        };
      }
      return null;
    }
    return {
      nodeId: dingTalkMatch[1],
      url: url,
      type: 'node'
    };
  } catch (error) {
    console.error('Error extracting DingTalk info:', error);
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
      file_extension: 'csv',
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
        const base64data = reader.result.split(',')[1]; // Remove data:text/csv;base64, prefix
        const dataUrl = `data:text/csv;base64,${base64data}`;

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

async function simulateDingTalkDownload(tabId) {
  try {
    // 尝试向内容脚本发送消息，如果失败则注入脚本
    let response;
    try {
      response = await chrome.tabs.sendMessage(tabId, {
        type: "API_DOWNLOAD"
      });
    } catch (messageError) {
      // 注入内容脚本
      await chrome.scripting.executeScript({
        target: { tabId: tabId },
        files: ['dingtalk-content-script.js']
      });

      // 等待脚本初始化
      await new Promise(resolve => setTimeout(resolve, 500));

      // 重新尝试发送消息
      response = await chrome.tabs.sendMessage(tabId, {
        type: "API_DOWNLOAD"
      });
    }

    // 即使响应验证失败，我们也检查是否有有效的下载URL
    if (!response || !response.success) {
      // 如果我们有downloadUrl，即使success为false也继续
      if (response && response.downloadUrl) {
        // 使用我们得到的数据继续
      } else {
        throw new Error(response?.msg || 'Failed to get download URL from content script');
      }
    }

    // 使用内容脚本返回的下载URL
    const downloadUrl = response.downloadUrl;
    const filename = response.filename || `dingtalk_document_${Date.now()}.xlsx`;

    // 启动下载
    chrome.downloads.download({
      url: downloadUrl,
      filename: filename,
      saveAs: true
    });

    return { success: true, message: "Download initiated successfully", downloadUrl: downloadUrl };
  } catch (error) {
    throw error;
  }
}

// 监听来自内容脚本的消息
chrome.runtime.onMessage.addListener((request, _sender, sendResponse) => {
  if (request.type === "FOUND_URL") {
    console.log('Found DingTalk download URL:', request.how, request.url);
    // 不再自动下载，只记录日志。下载由用户点击按钮触发
    sendResponse({ received: true });
  }
  return true;
});

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

// 飞书应用凭据（系统内置）
const FEISHU_CONFIG = {
  appId: 'cli_a8588718f4af901c',  // 替换为你的实际App ID
  appSecret: 'RlUiFV5bx7KldWeMUVT8rgqc7ynMWvYO'  // 替换为你的实际App Secret
};

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

    const filename = `feishu_${feishuInfo.type}_${feishuInfo.token}.csv`;
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



chrome.runtime.onMessage.addListener((request, _sender, sendResponse) => {
  if (request.action === 'extractSheetInfo') {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      const url = tabs[0].url;
      const sheetInfo = extractSheetInfo(url);

      if (sheetInfo) {
        sendResponse({
          success: true,
          sheetInfo: sheetInfo,
          type: 'google'
        });
      } else {
        const feishuInfo = extractFeishuInfo(url);
        if (feishuInfo) {
          sendResponse({
            success: true,
            feishuInfo: feishuInfo,
            type: 'feishu'
          });
        } else {
          const dingTalkInfo = extractDingTalkInfo(url);
          if (dingTalkInfo) {
            sendResponse({
              success: true,
              dingTalkInfo: dingTalkInfo,
              type: 'dingtalk'
            });
          } else {
            sendResponse({
              success: false,
              error: 'No supported document found'
            });
          }
        }
      }
    });
    return true;
  }

  if (request.action === 'downloadSheet') {
    chrome.downloads.download({
      url: request.downloadUrl,
      filename: request.filename,
      saveAs: true
    });
    sendResponse({success: true});
    return true;
  }

  if (request.action === 'exportFeishuDocument') {
    exportFeishuDocument(request.feishuInfo, request.accessToken)
      .then((result) => {
        sendResponse({success: true, downloadUrl: result.downloadUrl});
      })
      .catch((error) => {
        sendResponse({success: false, error: error.message});
      });
    return true;
  }

  if (request.action === 'getFeishuAccessToken') {
    getFeishuAccessToken(request.appId, request.appSecret)
      .then((token) => {
        sendResponse({success: true, token: token});
      })
      .catch((error) => {
        console.error('Detailed error in getFeishuAccessToken:', error);
        sendResponse({success: false, error: error.message});
      });
    return true;
  }

  if (request.action === 'debugFeishuApi') {
    (async () => {
      try {
        const response = await fetch('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json; charset=utf-8'
          },
          body: JSON.stringify({
            app_id: request.appId,
            app_secret: request.appSecret
          })
        });

        const data = await response.json();
        sendResponse({success: true, response: data});
      } catch (error) {
        sendResponse({success: false, error: error.message});
      }
    })();
    return true;
  }

  
  if (request.action === 'simulateDingTalkDownload') {
    chrome.tabs.query({active: true, currentWindow: true}, (tabs) => {
      const tabId = tabs[0].id;
      simulateDingTalkDownload(tabId)
        .then((result) => {
          sendResponse(result);
        })
        .catch((error) => {
          sendResponse({ success: false, error: error.message });
        });
    });
    return true;
  }

  });