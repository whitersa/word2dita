/**
* 解析样式字符串为对象
* @param {string} styleText 样式字符串
* @returns {Object} 样式对象
*/
function parseStyle(styleText) {
    // 1. 初始化结果对象
    const styles = {};

    // 2. 如果输入为空，直接返回
    if (!styleText) {
        return styles;
    }

    try {
        // 3. 分割样式字符串
        const declarations = styleText.split(';');

        // 4. 处理每个样式声明
        declarations.forEach(declaration => {
            // 跳过空声明
            if (!declaration.trim()) {
                return;
            }

            // 分割属性名和值
            let [property, value] = declaration.split(':');

            // 清理属性名和值
            property = property.trim().toLowerCase();
            value = value ? value.trim() : '';

            // 特殊值处理
            if (value) {
                // 处理引号
                value = value.replace(/['"]/g, '');

                // 处理颜色值
                if (property.includes('color')) {
                    value = normalizeColor(value);
                }

                // 处理数值单位
                if (/^-?\d+\.?\d*$/.test(value)) {
                    value = value + 'px';
                }

                // 存储样式
                styles[property] = value;
            }
        });
    } catch (e) {
        console.error('Style parsing error:', e);
    }

    return styles;
}

/**
* 标准化颜色值
* @param {string} color 颜色值
* @returns {string} 标准化的颜色值
*/
function normalizeColor(color) {
    // 处理命名颜色
    const namedColors = {
        'windowtext': '#000000',
        'transparent': 'transparent'
    };

    if (namedColors[color]) {
        return namedColors[color];
    }

    // 处理 rgb 值
    const rgbMatch = color.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
    if (rgbMatch) {
        const [_, r, g, b] = rgbMatch;
        return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
    }

    // 处理 # 值
    if (color.startsWith('#')) {
        // 处理简写形式 #ABC
        if (color.length === 4) {
            const [_, r, g, b] = color;
            return `#${r}${r}${g}${g}${b}${b}`;
        }
        return color.toLowerCase();
    }

    return color;
}

/**
* 将数字转换为两位十六进制
* @param {number|string} n 数字
* @returns {string} 两位十六进制字符串
*/
function toHex(n) {
    const hex = parseInt(n).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
}

function isWordContent(content) {
    return (
        // 检查 Word 特有的标记
        /<font face="Times New Roman"|class="?Mso|style="[^"]*\bmso-|style='[^'']*\bmso-|w:WordDocument/i.test(content) ||
        // 检查 Google Docs 标记
        /class="OutlineElement/.test(content) ||
        // 检查 Google Docs 内部 GUID
        /id="?docs\-internal\-guid\-/.test(content)
    );
}

/**
   * 检查是否是编号列表
   * @param {string} text - 要检查的文本
   * @returns {boolean} 是否是编号列表
   */
function isNumericList(text) {
    // 移除开头的空格
    text = text.replace(/^[\u00a0 ]+/, '');

    // 定义各种列表标记的正则表达式
    const patterns = [
        /^[IVXLMCD]{1,2}\.[ \u00a0]/,      // 罗马数字大写
        /^[ivxlmcd]{1,2}\.[ \u00a0]/,      // 罗马数字小写
        /^[a-z]{1,2}[\.\)][ \u00a0]/,      // 小写字母
        /^[A-Z]{1,2}[\.\)][ \u00a0]/,      // 大写字母
        /^[0-9]+\.[ \u00a0]/,              // 数字
        /^[\u3007\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d]+\.[ \u00a0]/, // 中文数字（简体）
        /^[\u58f1\u5f10\u53c2\u56db\u4f0d\u516d\u4e03\u516b\u4e5d\u62fe]+\.[ \u00a0]/  // 中文数字（繁体）
    ];

    let isNumeric = false;

    // 检查是否匹配任一模式
    this.tools.each(patterns, (pattern) => {
        if (pattern.test(text)) {
            isNumeric = true;
            return false; // 跳出循环
        }
    });

    return isNumeric;
}

/**
 * 检查是否是项目符号列表
 * @param {string} text - 要检查的文本
 * @returns {boolean} 是否是项目符号列表
 */
function isBulletList(text) {
    return /^[\s\u00a0]*[\u2022\u00b7\u00a7\u25CF]\s*/.test(text);
}

/**
 * 处理 Word 样式
 * @param {Node} node - 当前节点
 * @param {string} styleText - 样式文本
 * @returns {string|null} 处理后的样式
 */
function processWordStyles(node, styleText) {
    const editor = this.editor;
    const dom = editor.dom;
    const styles = {};

    // 解析样式
    const rawStyles = dom.parseStyle(styleText);

    // 处理每个样式
    this.tools.each(rawStyles, (value, name) => {
        switch (name) {
            // 处理列表样式
            case 'mso-list':
                // 提取列表级别
                const levelMatch = /\w+ \w+([0-9]+)/i.exec(styleText);
                if (levelMatch) {
                    node._listLevel = parseInt(levelMatch[1], 10);
                }

                // 处理忽略标记
                if (/Ignore/i.test(value) && node.firstChild) {
                    node._listIgnore = true;
                    node.firstChild._listIgnore = true;
                }
                break;

            // 对齐方式转换
            case 'horiz-align':
                name = 'text-align';
                break;
            case 'vert-align':
                name = 'vertical-align';
                break;

            // 颜色处理
            case 'font-color':
            case 'mso-foreground':
                name = 'color';
                break;
            case 'mso-background':
            case 'mso-highlight':
                name = 'background';
                break;

            // 字体样式处理
            case 'font-weight':
            case 'font-style':
                if (value !== 'normal') {
                    styles[name] = value;
                }
                return;

            // 注释处理
            case 'mso-element':
                if (/^(comment|comment-list)$/i.test(value)) {
                    node.remove();
                    return;
                }
                break;
        }

        // 处理注释相关样式
        if (name.indexOf('mso-comment') === 0) {
            node.remove();
            return;
        }

        // 保留需要的样式
        if (name.indexOf('mso-') !== 0 &&
            (this.retainStyleProperties === 'all' ||
                (this.validStyles && this.validStyles[name]))) {
            styles[name] = value;
        }
    });

    // 处理特殊样式
    if (/(bold)/i.test(styles['font-weight'])) {
        delete styles['font-weight'];
        node.wrap(new Node('b', 1));
    }

    if (/(italic)/i.test(styles['font-style'])) {
        delete styles['font-style'];
        node.wrap(new Node('i', 1));
    }

    // 序列化样式
    const finalStyles = dom.serializeStyle(styles, node.name);
    return finalStyles ? finalStyles : null;
}



module.exports = {
    parseStyle,
    isWordContent
}
