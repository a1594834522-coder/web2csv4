const axios = require('axios');

async function debugTest() {
  const url = 'https://dcnxbo68qr50.feishu.cn/sheets/IqLdstnEyhye8DtXUKwcRsWKnK2';

  console.log('Testing with URL:', url);

  try {
    const response = await axios.post('http://localhost:3000/api/export', {
      url: url,
      timestamp: Date.now()
    });

    console.log('Full response:', JSON.stringify(response.data, null, 2));
    console.log('fileToken value:', response.data.fileToken);
    console.log('fileToken type:', typeof response.data.fileToken);
  } catch (error) {
    if (error.response) {
      console.log('Error response:', error.response.status, error.response.data);
    } else {
      console.log('Network error:', error.message);
    }
  }
}

debugTest();