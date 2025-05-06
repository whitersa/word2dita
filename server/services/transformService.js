const fs = require('fs').promises;
const path = require('path');
const { cleanHtml, formatHtml, cleanEmptyTags } = require('../utils/htmlUtils');

// 定义基础路径
const BASE_DIR = path.join(__dirname, '../..');
const DEBUG_DIR = path.join(BASE_DIR, 'debug');

/**
 * HTML转换服务
 */
class TransformService {
    constructor() {
        // 确保单例
        if (!TransformService.instance) {
            TransformService.instance = this;
            // 绑定方法到实例
            this.transformContent = this.transformContent.bind(this);
            this.handleHtmlContent = this.handleHtmlContent.bind(this);
            // 确保debug目录存在
            this.ensureDebugDir();
        }
        return TransformService.instance;
    }

    /**
     * 确保debug目录存在
     */
    async ensureDebugDir() {
        try {
            await fs.access(DEBUG_DIR);
        } catch (error) {
            await fs.mkdir(DEBUG_DIR, { recursive: true });
        }
    }

    /**
     * 生成调试用的HTML文件
     * @param {string} content - HTML内容
     */
    async generateDebugHtml(content) {
        try {
            const now = new Date();
            const fileName = `debug_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}${String(now.getSeconds()).padStart(2, '0')}.html`;
            const debugHtml = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Debug HTML</title>
</head>
<body>
${content}
</body>
</html>`;
            const filePath = path.join(DEBUG_DIR, fileName);
            await fs.writeFile(filePath, debugHtml, 'utf8');
            return filePath;
        } catch (error) {
            console.error('生成调试HTML文件失败:', error);
            return null;
        }
    }

    /**
     * 获取服务实例
     */
    static getInstance() {
        if (!TransformService.instance) {
            TransformService.instance = new TransformService();
        }
        return TransformService.instance;
    }

    /**
     * 转换内容
     * @param {string} content - 要转换的内容
     * @returns {Promise<Object>} 转换结果，包含处理步骤和转换后的内容
     */
    async transformContent(content) {
        const processingSteps = [];
        try {
            // 1. 先清理HTML
            const cleanedContent = cleanHtml(content);
            processingSteps.push('1. HTML内容清理完成');

            // 2. 清理空标签
            const noEmptyTagsContent = cleanEmptyTags(cleanedContent);
            processingSteps.push('2. 空标签清理完成');

            // 生成调试用的HTML文件
            const debugFilePath = await this.generateDebugHtml(noEmptyTagsContent);
            if (debugFilePath) {
                processingSteps.push(`调试文件已生成: ${debugFilePath}`);
            }

            // 3. 格式化处理后的HTML
            const formattedContent = formatHtml(noEmptyTagsContent);
            processingSteps.push('3. HTML格式化完成');

            return {
                success: true,
                steps: processingSteps,
                html: formattedContent,
                debugFile: debugFilePath
            };
        } catch (error) {
            console.error('内容处理错误:', error);
            return {
                success: false,
                error: error.message,
                steps: processingSteps
            };
        }
    }

    /**
     * 处理HTML内容
     * @param {string} content - HTML内容
     * @returns {Promise<string>} 处理后的HTML内容
     */
    async handleHtmlContent(content) {
        return cleanHtml(content);
    }
}

// 创建并导出单例实例
const instance = new TransformService();
Object.freeze(instance); // 防止实例被修改
module.exports = instance; 