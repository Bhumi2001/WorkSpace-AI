const express = require('express');
const { createProxyMiddleware } = require('http-proxy-middleware');

const app = express();

app.use('/api', createProxyMiddleware({
  target: 'https://graph.microsoft.com',
  changeOrigin: true,
  pathRewrite: {
    '^/api': '', 
  },
}));


const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Proxy server listening on port ${PORT}`);
});
