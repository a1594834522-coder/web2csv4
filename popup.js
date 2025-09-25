document.addEventListener('DOMContentLoaded', function() {
  const statusDiv = document.getElementById('status');
  const sheetInfoDiv = document.getElementById('sheetInfo');
  const feishuInfoDiv = document.getElementById('feishuInfo');
  const downloadBtn = document.getElementById('downloadBtn');
  const debugSection = document.getElementById('debugSection');
  const debugInfo = document.getElementById('debugInfo');
  const testInjectionBtn = document.getElementById('testInjectionBtn');
  const checkUrlsBtn = document.getElementById('checkUrlsBtn');

  const sheetIdSpan = document.getElementById('sheetId');
  const gidSpan = document.getElementById('gid');
  const downloadUrlSpan = document.getElementById('downloadUrl');
  const realDownloadUrlSpan = document.getElementById('realDownloadUrl');
  const feishuTypeSpan = document.getElementById('feishuType');
  const feishuTokenSpan = document.getElementById('feishuToken');
  const feishuUrlSpan = document.getElementById('feishuUrl');
  const feishuRealDownloadUrlSpan = document.getElementById('feishuRealDownloadUrl');

  let currentInfo = null;
  let currentType = null;

  
  function updateStatus(message, type = 'info') {
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;

    
    console.log(`[${type.toUpperCase()}] ${message}`);
  }

  
  
  function showSheetInfo(sheetInfo) {
    currentInfo = sheetInfo;
    currentType = 'google';

    sheetIdSpan.textContent = sheetInfo.sheetId;
    gidSpan.textContent = sheetInfo.gid;
    downloadUrlSpan.textContent = sheetInfo.downloadUrl;

    sheetInfoDiv.style.display = 'block';
    feishuInfoDiv.style.display = 'none';
    downloadBtn.style.display = 'block';
    downloadBtn.disabled = false;
    downloadBtn.textContent = 'Download as XLSX';
  }

  function showFeishuInfo(feishuInfo) {
    currentInfo = feishuInfo;
    currentType = 'feishu';

    feishuTypeSpan.textContent = feishuInfo.type;
    feishuTokenSpan.textContent = feishuInfo.token;
    feishuUrlSpan.textContent = feishuInfo.url;

    feishuInfoDiv.style.display = 'block';
    sheetInfoDiv.style.display = 'none';
    downloadBtn.style.display = 'block';
    downloadBtn.disabled = false;
    downloadBtn.textContent = 'Export as XLSX';
  }


  
  function hideDocumentInfo() {
    sheetInfoDiv.style.display = 'none';
    feishuInfoDiv.style.display = 'none';
    downloadBtn.style.display = 'none';
    // 隐藏真实下载URL
    if (realDownloadUrlSpan) {
      realDownloadUrlSpan.textContent = '';
    }
    if (feishuRealDownloadUrlSpan) {
      feishuRealDownloadUrlSpan.textContent = '';
      feishuRealDownloadUrlSpan.style.display = 'none';
    }
    currentInfo = null;
    currentType = null;
  }

  function checkCurrentPage() {
    chrome.runtime.sendMessage(
      { action: 'extractSheetInfo' },
      function(response) {
        if (chrome.runtime.lastError) {
          console.error('Runtime error:', chrome.runtime.lastError.message);
          updateStatus('无法连接到扩展后台', 'error');
          return;
        }

        if (response && response.success) {
          if (response.type === 'google') {
            updateStatus('检测到Google表格！', 'success');
            showSheetInfo(response.sheetInfo);
          } else if (response.type === 'feishu') {
            updateStatus('检测到飞书文档！', 'success');
            showFeishuInfo(response.feishuInfo);
          }
        } else {
          updateStatus('当前页面未找到支持的文档', 'error');
          hideDocumentInfo();
        }
      }
    );
  }

  
  function downloadDocument() {
    if (!currentInfo) return;

    if (currentType === 'google') {
      const filename = `sheet_${currentInfo.sheetId}_${currentInfo.gid}.xlsx`;
      updateStatus('正在下载Google表格...', 'info');

      // Google Sheets的真实下载URL就是原始URL
      realDownloadUrlSpan.textContent = currentInfo.downloadUrl;

      chrome.runtime.sendMessage(
        {
          action: 'downloadSheet',
          downloadUrl: currentInfo.downloadUrl,
          filename: filename
        },
        function(response) {
          if (response && response.success) {
            updateStatus('下载已启动！', 'success');
          } else {
            updateStatus('下载失败，请重试。', 'error');
          }
        }
      );
    } else if (currentType === 'feishu') {
      updateStatus('正在导出飞书文档...', 'info');
      downloadBtn.disabled = true;

      chrome.runtime.sendMessage(
        {
          action: 'exportFeishuDocument',
          feishuInfo: currentInfo
        },
        function(response) {
          downloadBtn.disabled = false;
          if (response && response.success) {
            updateStatus('导出完成！', 'success');
            // 显示真实的下载URL
            if (response.downloadUrl) {
              feishuRealDownloadUrlSpan.textContent = response.downloadUrl;
              feishuRealDownloadUrlSpan.style.display = 'inline';
            }
          } else {
            const errorMsg = response?.error || '未知错误';
            updateStatus(`导出失败: ${errorMsg}`, 'error');
          }
        }
      );
    }
  }

  downloadBtn.addEventListener('click', downloadDocument);

  
  
  // 初始化时检查页面
  checkCurrentPage();
});