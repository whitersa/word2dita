const express = require('express');
const path = require('path');
const apiRoutes = require('./server/routes/api');

const app = express();
const port = process.env.PORT || 3000;

// 中间件设置
app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// 请求日志中间件
app.use((req, res, next) => {
    console.log(`${new Date().toISOString()} ${req.method} ${req.url}`);
    next();
});

// API路由
app.use('/api', apiRoutes);

// 错误处理中间件
app.use((err, req, res, next) => {
    console.error('服务器错误:', err.stack);
    res.status(500).json({ 
        success: false,
        error: '服务器内部错误' 
    });
});

// 启动服务器
app.listen(port, () => {
    console.log(`服务器运行在 http://localhost:${port}`);
}); 