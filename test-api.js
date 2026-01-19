// Simple test script for the Google Slides Opener API
// Run with: node test-api.js

const http = require('http');

const API_HOST = '127.0.0.1';
const API_PORT = 9595;

// Test presentation URL
const TEST_URL = 'https://docs.google.com/presentation/d/1rc9BSX-0TrU7c5LGeLDRyH3zRN89-uDuXEEqOpcnLVg/edit';

function makeRequest(method, path, data = null) {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: API_HOST,
      port: API_PORT,
      path: path,
      method: method,
      headers: {
        'Content-Type': 'application/json'
      },
      timeout: 5000
    };

    const req = http.request(options, (res) => {
      let responseData = '';
      res.on('data', (chunk) => {
        responseData += chunk;
      });
      res.on('end', () => {
        try {
          const response = JSON.parse(responseData);
          resolve({ status: res.statusCode, data: response });
        } catch (error) {
          reject(error);
        }
      });
    });

    req.on('error', (error) => {
      reject(error);
    });

    req.on('timeout', () => {
      req.destroy();
      reject(new Error('Request timeout'));
    });

    if (data) {
      req.write(JSON.stringify(data));
    }
    req.end();
  });
}

async function runTests() {
  console.log('=== Google Slides Opener API Test ===\n');

  // Test 1: Check status
  console.log('1. Testing GET /api/status...');
  try {
    const result = await makeRequest('GET', '/api/status');
    console.log('   ✓ Status:', result.data);
    console.log('   Response:', JSON.stringify(result.data, null, 2));
  } catch (error) {
    console.error('   ✗ Error:', error.message);
    console.error('   Make sure the Google Slides Opener app is running!');
    return;
  }

  console.log();

  // Test 2: Open presentation
  console.log('2. Testing POST /api/open-presentation...');
  try {
    const result = await makeRequest('POST', '/api/open-presentation', { url: TEST_URL });
    console.log('   ✓ Success:', result.data.message);
    console.log('   The presentation should now be opening...');
  } catch (error) {
    console.error('   ✗ Error:', error.message);
  }

  console.log();

  // Wait a bit before closing
  console.log('3. Waiting 5 seconds before closing...');
  await new Promise(resolve => setTimeout(resolve, 5000));

  // Test 3: Close presentation
  console.log('4. Testing POST /api/close-presentation...');
  try {
    const result = await makeRequest('POST', '/api/close-presentation');
    console.log('   ✓ Success:', result.data.message);
    console.log('   The presentation should now be closed.');
  } catch (error) {
    console.error('   ✗ Error:', error.message);
  }

  console.log();
  console.log('=== Tests Complete ===');
}

// Run the tests
runTests().catch(console.error);
