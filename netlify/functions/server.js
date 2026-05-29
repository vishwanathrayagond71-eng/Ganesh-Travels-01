const serverless = require('serverless-http');
const app = require('../../server'); // Import Express app
const { initAllSheets } = require('../../utils/excelService');

let initialized = false;

// Wrap the Express app in a serverless-http handler with binary support
const handler = serverless(app, {
  binary: [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/octet-stream'
  ]
});

module.exports.handler = async (event, context) => {
  // Make sure database is initialized on the first function invocation
  if (!initialized) {
    try {
      await initAllSheets();
      initialized = true;
    } catch (err) {
      console.error('❌ Failed to initialize database in Netlify Function:', err);
    }
  }
  
  // Forward request to Express app
  return handler(event, context);
};
