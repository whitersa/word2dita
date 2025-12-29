const fs = require('fs').promises;
const path = require('path');
const { cleanHtml, formatHtml, } = require('../utils/htmlUtilsDita');

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
            // 初始化debug目录
            this.initializeDebugDir();
        }
        return TransformService.instance;
    }

    /**
     * 初始化debug目录
     */
    async initializeDebugDir() {
        try {
            // 检查debug目录是否存在
            try {
                await fs.access(DEBUG_DIR);
            } catch (error) {
                // 如果目录不存在，创建它
                await fs.mkdir(DEBUG_DIR, { recursive: true });
                return;
            }

            // 如果目录存在，清理旧文件
            const files = await fs.readdir(DEBUG_DIR);
            if (files.length > 0) {
                // 获取所有debug文件的详细信息
                const fileStats = await Promise.all(
                    files.map(async (file) => {
                        const filePath = path.join(DEBUG_DIR, file);
                        const stats = await fs.stat(filePath);
                        return {
                            name: file,
                            path: filePath,
                            mtime: stats.mtime
                        };
                    })
                );

                // 按修改时间排序，保留最新的文件
                fileStats.sort((a, b) => b.mtime - a.mtime);

                // 删除除最新文件外的所有文件
                for (let i = 1; i < fileStats.length; i++) {
                    await fs.unlink(fileStats[i].path);
                }
            }
        } catch (error) {
            console.error('初始化debug目录失败:', error);
        }
    }

    /**
     * 生成调试用的HTML文件
     * @param {string} content - HTML内容
     */
    async generateDebugHtml(content) {
        try {
            // 先清理旧文件
            await this.initializeDebugDir();

            const now = new Date();
            const fileName = `debug_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}${String(now.getSeconds()).padStart(2, '0')}.html`;
            const debugHtml = `
                <!DOCTYPE html>
                <html>
                    <head>
                        <meta charset="UTF-8">
                        <title>Debug HTML</title>
                    </head>
                    <body>
                        ${content}
                    </body>
                </html>
            `;
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
            // 1. 清理空标签（使用JSDOM处理标准HTML结构）
            // 2. 进行HTML清理和DITA转换
            processingSteps.push('2. HTML内容清理和DITA转换完成');
            const cleanedContent = cleanHtml(content);

            processingSteps.push('1. 空标签清理完成');
            // const noEmptyTagsContent = cleanEmptyTags(cleanedContent);

            // 生成调试用的HTML文件
            const debugFilePath = await this.generateDebugHtml(cleanedContent);
            if (debugFilePath) {
                processingSteps.push(`调试文件已生成: ${debugFilePath}`);
            }

            // 3. 格式化处理后的HTML
            const formattedContent = formatHtml(cleanedContent);
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
}

// 创建并导出单例实例
const instance = new TransformService();
Object.freeze(instance); // 防止实例被修改
module.exports = instance; 