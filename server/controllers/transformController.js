const { transformContent } = require('../services/transformService');

/**
 * 内容转换控制器
 */
class TransformController {
    /**
     * 处理内容转换请求
     */
    async handleTransform(req, res) {
        try {
            const { content } = req.body;
            
            if (!content) {
                return res.status(400).json({ 
                    success: false,
                    error: '内容不能为空' 
                });
            }

            // 调用服务层处理转换逻辑
            const result = await transformContent(content);
            
            if (!result.success) {
                return res.status(500).json(result);
            }

            res.json(result);
            
        } catch (error) {
            console.error('转换处理错误:', error);
            res.status(500).json({ 
                success: false,
                error: '处理失败: ' + error.message,
                steps: []
            });
        }
    }
}

module.exports = new TransformController(); 