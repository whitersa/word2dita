const { JSDOM } = require('jsdom');

/**
 * 清理和规范化 HTML 内容
 * @param {string} html - 要清理的HTML
 * @returns {string} - 清理后的HTML
 */
function cleanHtml(html) {
    if (!html) return '';

    try {
        // 1. 基础清理
        html = html
            // 移除注释
            .replace(/<!--[\s\S]*?-->/g, '')
            // 移除空行
            .replace(/^\s*[\r\n]/gm, '')
            // 规范化空格
            .replace(/\s+/g, ' ');

        // 2. 移除特殊标签和元素
        html = html
            // 移除 style 标签及其内容
            .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
            // 移除其他特殊标签
            .replace(/<(!|script[^>]*>.*?<\/script(?=[>\s])|\/?(\?xml(:\w+)?|img|meta|link|\w:\w+)(?=[\s\/>]))[^>]*>/gi, '')
            // 移除 XML 命名空间
            .replace(/<html[^>]*>/i, '<html>')
            .replace(/<\/?[a-z]*:[^>]*>/gi, '')
            // 移除 Word 特有的 class
            .replace(/class="?Mso[a-zA-Z]+"/g, '');

        // 3. 清理样式，只保留指定样式
        html = cleanSelectiveStyles(html);
        // 生成debug文件以查看第一阶段处理结果
        try {
            require('fs').writeFileSync('debug_stage1.html', html);
        } catch (err) {
            console.error('Debug file write error:', err);
        }
        // 4. 添加列表层级class
        html = addListLevelClasses(html);

        // 5. 转换为嵌套列表
        html = convertMsoListToNestedLists(html);

        // 6. 清理非列表段落标签
        html = cleanNonListParagraphs(html);

        // 7. 清理标题标签的margin-left
        html = cleanHeadingMargins(html);

        // 8. 清理align属性
        html = cleanAlignAttributes(html);

        // 9. 清理标签
        html = html
            // 移除空span标签
            .replace(/<span>\s*<\/span>/gi, '')
            // 移除多余的换行标签
            .replace(/<br>\s*<br>/gi, '<br>')
            // 移除末尾的换行
            .replace(/<br>$/i, '')
            // 处理font标签，保留内容
            .replace(/<font[^>]*>([\s\S]*?)<\/font>/gi, '$1')
            // 处理s(删除线)标签，精确匹配并保留内容
            .replace(/<s>([\s\S]*?)<\/s>/gi, '$1');

        // 10. 清理list内冗余的span标签
        html = cleanRedundantSpans(html);

        // 11. 清理特殊字符
        html = html
            // 替换 non-breaking spaces
            .replace(/&nbsp;/g, ' ')
            // 规范化引号
            .replace(/[""]/g, '"')
            .replace(/['']/g, "'");

        // 12. 移除空属性
        html = html
            .replace(/(\w+)="\s*"/g, '')
            .replace(/\s+>/g, '>');

        // 13. 清理多余的空白
        html = html
            .replace(/>\s+</g, '><')
            .trim();

        // 14. 解码HTML实体，确保结构化处理基于原始HTML标签
        html = html
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&amp;/g, '&');

        // 15. 结构化内容
        html = structureContent(html);

        // 16. 最后处理表格，避免被JSDOM再次序列化
        html = cleanTables(html);



        return html;
    } catch (error) {
        console.error('HTML清理错误:', error);
        return html;
    }
}

/**
 * 清理非列表段落标签
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function cleanNonListParagraphs(html) {
    if (!html) return '';

    try {
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
        const document = dom.window.document;

        // 处理所有p标签
        const paragraphs = document.getElementsByTagName('p');
        for (let p of paragraphs) {
            const style = p.getAttribute('style') || '';

            // 检查是否是列表项（是否包含mso-list样式）
            if (!style.includes('mso-list:')) {
                // 不是列表项，清理class和margin-left样式
                p.removeAttribute('class');

                if (style) {
                    // 移除margin-left样式，保留其他样式
                    const styles = style.split(';')
                        .map(s => s.trim())
                        .filter(s => s && !s.toLowerCase().startsWith('margin-left:'));

                    if (styles.length > 0) {
                        p.setAttribute('style', styles.join('; '));
                    } else {
                        p.removeAttribute('style');
                    }
                }
            }
        }

        return document.body.innerHTML;
    } catch (error) {
        console.error('清理非列表段落标签错误:', error);
        return html;
    }
}

/**
 * 格式化HTML（使其更易读）
 * @param {string} html - 要格式化的HTML
 * @returns {string} - 格式化后的HTML
 */
function formatHtml(html) {
    if (!html) return '';

    try {
        let formatted = '';
        let indent = 0;

        // 分割标签
        const tags = html.split(/(<\/?[^>]+>)/g);

        // 处理每个标签
        for (let tag of tags) {
            if (!tag.trim()) continue;

            // 处理结束标签
            if (tag.startsWith('</')) {
                indent--;
                formatted += '  '.repeat(Math.max(0, indent)) + tag + '\n';
            }
            // 处理自闭合标签
            else if (tag.endsWith('/>')) {
                formatted += '  '.repeat(indent) + tag + '\n';
            }
            // 处理开始标签
            else if (tag.startsWith('<')) {
                formatted += '  '.repeat(indent) + tag + '\n';
                // 不增加缩进的标签
                if (!tag.match(/<(br|hr|img|input|link|meta|area|base|col|command|embed|keygen|param|source|track|wbr)/i)) {
                    indent++;
                }
            }
            // 处理文本内容
            else {
                formatted += '  '.repeat(indent) + tag.trim() + '\n';
            }
        }

        return formatted.trim();
    } catch (error) {
        console.error('HTML格式化错误:', error);
        return html; // 如果处理出错，返回原始内容
    }
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

    // 检查是否匹配任一模式
    return patterns.some(pattern => pattern.test(text));
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
 * Clean class and ID attributes
 */
function cleanClassAndIdAttributes() {
    const attributes = ['class', 'id'];
    attributes.forEach(attr => {
        const variations = [
            ` ${attr} = `,
            ` ${attr}= `,
            ` ${attr} =`
        ];

        variations.forEach(variant => {
            HTMLTransformer.replaceText(variant, ` ${attr}=`);
        });

        this.removeAttributeContent(attr);
        HTMLTransformer.replaceText(` ${attr}=""`, '');
    });
}

/**
 * Clean empty tags from HTML
 */
function cleanEmptyTags(content) {
    function removeEmptyTags(content) {
        const chars = content.split('');
        const result = [];
        let state = 0;
        let startIndex = 0;
        let writeIndex = 0;
        let inTable = false;  // 标记是否在表格内

        for (let i = 0; i < chars.length; i++) {
            // 检查是否进入或离开表格
            if (i + 5 < chars.length && chars.slice(i, i + 6).join('') === '<table') {
                inTable = true;
            } else if (i + 7 < chars.length && chars.slice(i, i + 8).join('') === '</table>') {
                inTable = false;
            }

            if (state === 0 && chars[i] === '<' && chars[i + 1] !== '/') {
                state = 1;
                startIndex = i;
            }

            if (state === 2 && chars[i] === '>') {
                // 只有在不在表格内时才移除空标签
                if (!inTable) {
                    for (let j = 0; j <= i - startIndex; j++) {
                        result[j + startIndex] = '';
                    }
                    chars[i] = '';
                } else {
                    // 在表格内，保留所有标签
                    for (let j = startIndex; j <= i; j++) {
                        result[j] = chars[j];
                    }
                }
                state = 0;
            }

            if (state === 1 && chars[i] === '>') {
                const isEmptyTag = chars[i - 2] !== '/' &&
                    chars[i - 1] !== '/' &&
                    chars[i + 1] === '<' &&
                    chars[i + 2] === '/';
                state = isEmptyTag ? 2 : 0;
            }

            if (state !== 2 || inTable) {
                result[writeIndex] = chars[i];
                writeIndex++;
            }
        }

        return result.join('');
    }

    /**
     * Remove tags that only contain newlines
     */
    function removeNewlineTags(content) {
        const chars = content.split('');
        const result = [];
        let state = 0;
        let startIndex = 0;
        let writeIndex = 0;
        let inTable = false;  // 标记是否在表格内

        for (let i = 0; i < chars.length; i++) {
            // 检查是否进入或离开表格
            if (i + 5 < chars.length && chars.slice(i, i + 6).join('') === '<table') {
                inTable = true;
            } else if (i + 7 < chars.length && chars.slice(i, i + 8).join('') === '</table>') {
                inTable = false;
            }

            if (state === 0 && chars[i] === '<' && chars[i + 1] !== '/') {
                state = 1;
                startIndex = i;
            }

            if (state === 2 && chars[i] === '>') {
                // 只有在不在表格内时才移除只包含换行的标签
                if (!inTable) {
                    for (let j = 0; j <= i - startIndex; j++) {
                        result[j + startIndex] = '';
                    }
                    chars[i] = '';
                } else {
                    // 在表格内，保留所有标签
                    for (let j = startIndex; j <= i; j++) {
                        result[j] = chars[j];
                    }
                }
                state = 0;
            }

            if (state === 1 && chars[i] === '>') {
                const isNewlineTag = chars[i - 2] !== '/' &&
                    chars[i - 1] !== '/' &&
                    chars[i + 1] === '\\n' &&
                    chars[i + 2] === '<' &&
                    chars[i + 3] === '/';
                state = isNewlineTag ? 2 : 0;
            }

            if (state !== 2 || inTable) {
                result[writeIndex] = chars[i];
                writeIndex++;
            }
        }

        return result.join('');
    }

    function replaceText(content, searchText, replaceText) {
        return content.replace(new RegExp(searchText, 'g'), replaceText);
    }
    content = replaceText(content, '> <', '><');
    content = replaceText(content, '> \\n', '>\\n');
    content = removeEmptyTags(content);
    content = removeNewlineTags(content);
    return content;
}

/**
   * Clean &nbsp; tags
   */
function cleanNbspTags() {
    const variations = [
        '> &nbsp;<',
        '>&nbsp; <'
    ];

    variations.forEach(variant => {
        HTMLTransformer.replaceText(variant, '>&nbsp;<');
    });

    this.removeSingleNbspTags();
}

/**
 * Clean multiple &nbsp; occurrences
 */
function cleanMultipleNbsp() {
    const variations = [
        '&nbsp;&nbsp;',
        '&nbsp; ',
        ' &nbsp;'
    ];

    variations.forEach(variant => {
        HTMLTransformer.replaceText(variant, ' ');
    });
}

/**
 * Clean HTML comments
 */
function cleanComments() {
    HTMLTransformer.replaceText('\\x3c!--', '&%&%&%&%&%!--');
    this.removeTagContent('<', '>');
    HTMLTransformer.replaceText('<>', ' ');
    HTMLTransformer.replaceText('&%&%&%&%&%!--', '\\x3c!--');
}

/**
 * Remove all attributes except href and src
 */
function cleanAllAttributes() {
    const content = AppState.editor.currentText;
    const chars = content.split('');
    const result = [];
    let state = 1;
    let pos = 0;

    for (let i = 0; i < chars.length; i++) {
        const char = chars[i];

        if (char === '<') {
            state = this.determineTagState(chars, i);
        } else if (char === ' ') {
            state = this.updateStateOnSpace(state, chars, i);
        } else if (char === '"') {
            state = this.updateStateOnQuote(state);
        } else if (char === '>' || (char === '/' && chars[i + 1] === '>')) {
            state = 1;
        }

        if (this.shouldKeepChar(state)) {
            result[pos++] = char;
        }
    }
}

/**
     * Helper method to determine tag state
     * @private
     */
function determineTagState(chars, index) {
    if (chars[index + 1] === '!' &&
        chars[index + 2] === '-' &&
        chars[index + 3] === '-') {
        return 1;
    }
    if (chars[index + 1] === 'a' && chars[index + 2] === ' ') {
        return 4;
    }
    if (chars[index + 1] === 'i' &&
        chars[index + 2] === 'm' &&
        chars[index + 3] === 'g' &&
        chars[index + 4] === ' ') {
        return 14;
    }
    return 2;
}

/**
 * Helper method to update state on space character
 * @private
 */
function updateStateOnSpace(currentState, chars, index) {
    if (currentState === 2) return 3;
    if (currentState === 4 || currentState === 5) {
        if (this.isHrefAttribute(chars, index)) return 6;
        if (this.isDownloadAttribute(chars, index)) return 6;
        if (currentState === 4) return 5;
    }
    if (currentState === 14 || currentState === 15) {
        if (this.isSrcAttribute(chars, index)) return 16;
        if (currentState === 14) return 15;
    }
    if (currentState === 8 || currentState === 18) return 3;
    return currentState;
}

/**
 * Helper method to check if chars form href attribute
 * @private
 */
function isHrefAttribute(chars, index) {
    return chars[index + 1] === 'h' &&
        chars[index + 2] === 'r' &&
        chars[index + 3] === 'e' &&
        chars[index + 4] === 'f';
}

/**
 * Helper method to check if chars form src attribute
 * @private
 */
function isSrcAttribute(chars, index) {
    return chars[index + 1] === 's' &&
        chars[index + 2] === 'r' &&
        chars[index + 3] === 'c';
}

/**
 * Helper method to check if chars form download attribute
 * @private
 */
function isDownloadAttribute(chars, index) {
    return chars[index + 1] === 'd' &&
        chars[index + 2] === 'o' &&
        chars[index + 3] === 'w' &&
        chars[index + 4] === 'n' &&
        chars[index + 5] === 'l' &&
        chars[index + 6] === 'o' &&
        chars[index + 7] === 'a' &&
        chars[index + 8] === 'd';
}

/**
 * Helper method to update state on quote character
 * @private
 */
function updateStateOnQuote(state) {
    const stateMap = {
        7: 8,
        6: 7,
        17: 18,
        16: 17
    };
    return stateMap[state] || state;
}

/**
 * Helper method to determine if character should be kept
 * @private
 */
function shouldKeepChar(state) {
    return [1, 2, 4, 6, 7, 8, 14, 16, 17, 18].includes(state);
}

/**
 * Remove tags that only contain &nbsp;
 */
function removeSingleNbspTags() {
    const content = AppState.editor.currentText;
    const chars = content.split('');
    const result = [];
    let state = 0;
    let startIndex = 0;
    let writeIndex = 0;

    for (let i = 0; i < chars.length; i++) {
        if (state === 0 && chars[i] === '<' && chars[i + 1] !== '/') {
            state = 1;
            startIndex = i;
        }

        if (state === 2 && chars[i] === '>') {
            for (let j = 0; j <= i - startIndex; j++) {
                result[j + startIndex] = '';
            }
            chars[i] = '';
            state = 0;
        }

        if (state === 1 && chars[i] === '>') {
            const isNbspTag = chars[i - 2] !== '/' &&
                chars[i - 1] !== '/' &&
                chars[i + 1] === '&' &&
                chars[i + 2] === 'n' &&
                chars[i + 3] === 'b' &&
                chars[i + 4] === 's' &&
                chars[i + 5] === 'p' &&
                chars[i + 6] === ';' &&
                chars[i + 7] === '<' &&
                chars[i + 8] === '/';
            state = isNbspTag ? 2 : 0;
        }

        result[writeIndex] = chars[i];
        writeIndex++;
    }

    AppState.editor.currentText = result.join('');
}

/**
 * Remove content between specified tags
 * @param {string} startTag - Opening tag
 * @param {string} endTag - Closing tag
 * @returns {number} Number of occurrences removed
 */
function removeTagContent(startTag, endTag) {
    const content = AppState.editor.currentText;
    const chars = content.split('');
    const startChars = startTag.split('');
    const endChars = endTag.split('');
    const result = [];

    let inTag = 0;
    let writeIndex = 0;
    let occurrences = 0;
    let state = 1;

    for (let i = 0; i < chars.length; i++) {
        if (chars[i] === '<') {
            inTag = 1;
        }
        if (chars[i] === '>') {
            inTag = 0;
        }

        if (inTag === 1) {
            let isStartTag = true;
            for (let j = 0; j < startTag.length; j++) {
                if (startChars[j] !== chars[i + j]) {
                    isStartTag = false;
                    break;
                }
            }

            if (isStartTag) {
                occurrences++;
                state = -999;
                i += startTag.length;
                for (let j = 0; j < startTag.length; j++) {
                    result[writeIndex++] = startChars[j];
                }
                continue;
            }
        }

        let isEndTag = true;
        for (let j = 0; j < endTag.length; j++) {
            if (endChars[j] !== chars[i + j]) {
                isEndTag = false;
                break;
            }
        }

        if (isEndTag) {
            state = 0;
        }

        if (state !== -999 && state > 0) {
            result[writeIndex++] = chars[i];
        }

        state++;
    }

    AppState.editor.currentText = result.join('');
    return occurrences;
}

/**
 * Remove style attributes from HTML content
 */
function removeStyleAttributes() {
    const content = AppState.editor.currentText;
    const chars = content.split('');
    const result = [];
    let state = 1;
    let writeIndex = 0;

    for (let i = 0; i < chars.length; i++) {
        if (this.isStyleAttributeStart(chars, i)) {
            state = -999;
            i += 6;
            continue;
        }

        if (state === -999 && chars[i + 1] === '"') {
            state = -2;
        }

        if (state !== -999 && state !== -2) {
            result[writeIndex++] = chars[i];
        }

        state++;
    }

    AppState.editor.currentText = result.join('');
}

/**
 * Helper method to check if current position is start of style attribute
 * @private
 */
function isStyleAttributeStart(chars, index) {
    return chars[index] === 's' &&
        chars[index + 1] === 't' &&
        chars[index + 2] === 'y' &&
        chars[index + 3] === 'l' &&
        chars[index + 4] === 'e' &&
        chars[index + 5] === '=' &&
        chars[index + 6] === '"';
}

/**
 * HTML Code Formatting
 */
const CodeFormatter = {
    /**
     * HTML tag definitions
     */
    SPECIAL_TAGS: {
        doctype: ['DOCTYPE', 'doctype'],
        meta: ['META', 'meta'],
        link: ['LINK', 'link'],
        base: ['BASE', 'base'],
        br: ['BR', 'br'],
        col: ['COL', 'col'],
        command: ['command'],
        embed: ['embed'],
        hr: ['HR', 'hr'],
        img: ['IMG', 'img'],
        input: ['input'],
        param: ['param'],
        source: ['source']
    },

    /**
     * Format HTML code with proper indentation
     */
    formatCode() {
        const content = AppState.editor.currentText;
        const chars = content.split('');
        const result = [];
        let indentLevel = 0;
        let writeIndex = 0;
        let isSpecialTag = false;
        let isClosingTag = false;

        for (let i = 0; i < chars.length; i++) {
            if (chars[i] === '<') {
                isSpecialTag = this.checkForSpecialTag(chars, i);
                isClosingTag = chars[i + 1] === '/';

                if (!isSpecialTag && !isClosingTag) {
                    indentLevel++;
                }
                if (isClosingTag) {
                    indentLevel--;
                }
            }

            // Add newline and indentation
            if (chars[i] === '\\n') {
                result[writeIndex++] = chars[i];
                const targetIndent = chars[i + 1] === '/' ? indentLevel - 1 : indentLevel;
                for (let j = 0; j < targetIndent; j++) {
                    result[writeIndex++] = '\\t';
                }
            } else {
                result[writeIndex++] = chars[i];
            }
        }

        // Remove leading newline if exists
        if (result[0] === '\\n') {
            result.shift();
        }

        AppState.editor.currentText = result.join('');
    },

    /**
     * Format HTML code with proper indentation (alternative version)
     */
    formatCodeAlternative() {
        const content = AppState.editor.currentText;
        const chars = content.split('');
        const result = [];
        let indentLevel = 0;
        let writeIndex = 0;
        let isSpecialTag = false;

        for (let i = 0; i < chars.length; i++) {
            // Check for special tags that affect indentation
            if (i >= 5 && chars[i - 5] === '<') {
                isSpecialTag = this.checkForSpecialTagBackward(chars, i);
            }

            // Handle indentation
            if (chars[i] === '<') {
                if (!isSpecialTag) {
                    if (chars[i + 1] === '/') {
                        indentLevel--;
                    } else {
                        for (let j = i + 1; chars[j] !== '>' && j < chars.length; j++) {
                            if (chars[j] === '/' && chars[j + 1] === '>') {
                                isSpecialTag = true;
                                break;
                            }
                        }
                        if (!isSpecialTag) {
                            indentLevel++;
                        }
                    }
                }
            }

            // Add newline and indentation
            if (chars[i] === '\\n') {
                result[writeIndex++] = chars[i];
                if (chars[i + 1] === '/') {
                    for (let j = 0; j < indentLevel - 1; j++) {
                        result[writeIndex++] = '\\t';
                    }
                } else {
                    for (let j = 0; j < indentLevel; j++) {
                        result[writeIndex++] = '\\t';
                    }
                }
            } else {
                result[writeIndex++] = chars[i];
            }
        }

        AppState.editor.currentText = result.join('');
    },

    /**
     * Check if current position starts a special tag
     * @private
     */
    checkForSpecialTag(chars, startIndex) {
        const tagText = this.getTagText(chars, startIndex + 1);
        return Object.values(this.SPECIAL_TAGS).flat().some(tag =>
            tagText.startsWith(tag)
        );
    },

    /**
     * Check if current position is part of a special tag (backward check)
     * @private
     */
    checkForSpecialTagBackward(chars, currentIndex) {
        const lookback = 5;
        const tagStart = currentIndex - lookback;
        if (tagStart < 0) return false;

        const tagText = this.getTagText(chars, tagStart);
        return Object.values(this.SPECIAL_TAGS).flat().some(tag =>
            tagText.startsWith(tag)
        );
    },

    /**
     * Get tag text starting from a position
     * @private
     */
    getTagText(chars, startIndex) {
        let text = '';
        let i = startIndex;
        while (i < chars.length && chars[i] !== '>' && chars[i] !== ' ') {
            text += chars[i];
            i++;
        }
        return text;
    }
};

/**
 * 清理样式，只保留加粗、斜体、下划线、msolist和marginleft相关的样式
 * @param {string} html - 要清理的HTML
 * @returns {string} - 清理后的HTML
 */
function cleanSelectiveStyles(html) {
    if (!html) return '';

    try {
        // 1. 处理h1-h6标签及其内部内容
        html = html.replace(/<h([1-6])[^>]*>([\s\S]*?)<\/h\1>/gi, (match, level, content) => {
            // 清理内部的span标签
            let cleanedContent = content.replace(/<span[^>]*style="([^"]*)"[^>]*>([\s\S]*?)<\/span>/gi, 
                (spanMatch, styles, spanContent) => {
                    const styleDeclarations = styles.split(';');
                    const hasImportantStyle = styleDeclarations.some(declaration => {
                        declaration = declaration.trim();
                        return (
                            /^font-weight:\s*bold/i.test(declaration) ||
                            /^font-style:\s*italic/i.test(declaration) ||
                            /^text-decoration:(?:[^;]*)?underline/i.test(declaration)
                        );
                    });

                    // 如果span没有重要样式，只保留内容
                    return hasImportantStyle ? spanMatch : spanContent;
                }
            );

            // 清理其他可能的span标签（没有style属性的）
            cleanedContent = cleanedContent.replace(/<span[^>]*>([\s\S]*?)<\/span>/gi, '$1');

            // 返回清理后的h标签
            return `<h${level}>${cleanedContent}</h${level}>`;
        });

        // 2. 处理其他标签的style属性
        html = html.replace(/style="([^"]*)"/gi, (match, styles) => {
            const preserved = [];

            // 分割多个样式声明
            const styleDeclarations = styles.split(';');

            for (let declaration of styleDeclarations) {
                declaration = declaration.trim();
                if (!declaration) continue;

                // 检查是否是要保留的样式
                if (
                    /^mso-list:/i.test(declaration) ||
                    /^margin-left:/i.test(declaration) ||
                    /^font-weight:\s*bold/i.test(declaration) ||
                    /^font-style:\s*italic/i.test(declaration) ||
                    /^text-decoration:(?:[^;]*)?underline/i.test(declaration)
                ) {
                    preserved.push(declaration);
                }
            }

            // 如果有要保留的样式，返回新的style属性
            return preserved.length > 0 ? `style="${preserved.join('; ')}"` : '';
        });

        return html;
    } catch (error) {
        console.error('样式清理错误:', error);
        return html; // 如果处理出错，返回原始内容
    }
}

/**
 * 根据margin-left值判断列表层级
 * @param {string} marginLeft - margin-left值，例如 "22.6500pt"
 * @returns {number} - 返回层级，从1开始
 */
function getListLevel(marginLeft) {
    if (!marginLeft) return 1;

    // 提取数值部分
    const match = marginLeft.match(/(\d+(\.\d+)?)/);
    if (!match) return 1;

    const value = parseFloat(match[1]);

    // 根据margin-left值判断层级
    if (value <= 22.65) return 1;
    if (value <= 45.35) return 2;
    if (value <= 68) return 3;    // 预留第三层级
    return Math.ceil(value / 22.65); // 其他情况按比例计算
}

/**
 * 判断HTML元素是否是列表项
 * @param {string} html - HTML元素
 * @returns {Object} - 返回列表信息，包括是否是列表项、层级等
 */
function analyzeListItem(html) {
    try {
        // 提取style属性
        const styleMatch = html.match(/style="([^"]*)"/);
        if (!styleMatch) return { isList: false };

        const style = styleMatch[1];

        // 检查是否包含mso-list
        const hasMsoList = style.includes('mso-list:');
        if (!hasMsoList) return { isList: false };

        // 提取margin-left值
        const marginMatch = style.match(/margin-left:\s*([^;]+)/);
        const marginLeft = marginMatch ? marginMatch[1].trim() : '';

        // 获取层级
        const level = getListLevel(marginLeft);

        // 提取列表标记文本（如 "a.", "i.", "1." 等）
        const markerMatch = html.match(/<span[^>]*style="[^"]*mso-list:Ignore[^"]*"[^>]*>([\s\S]*?)<\/span>/);
        const marker = markerMatch ? markerMatch[1].trim() : '';

        return {
            isList: true,
            level,
            marginLeft,
            marker,
            style
        };
    } catch (error) {
        console.error('分析列表项错误:', error);
        return { isList: false };
    }
}

/**
 * 动态分析列表层级并添加对应的class
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function addListLevelClasses(html) {
    if (!html) return '';

    try {
        const marginValues = new Set(); // 存储所有不同的margin-left值
        const marginMap = new Map();    // 存储margin值到层级的映射

        // 第一遍扫描：收集所有不同的margin-left值
        let processedHtml = html.replace(/<p[^>]*style="([^"]*)"[^>]*>/gi, (match, style) => {
            if (style.includes('mso-list:')) {
                const marginMatch = style.match(/margin-left:\s*([0-9.]+)([a-z%]*)/i);
                if (marginMatch) {
                    const [, value, unit] = marginMatch;
                    marginValues.add(value + unit);
                }
            }
            return match;
        });

        // 将margin值排序并建立映射关系
        const sortedMargins = Array.from(marginValues).sort((a, b) => {
            const valueA = parseFloat(a);
            const valueB = parseFloat(b);
            return valueA - valueB;
        });

        // 建立margin值到层级的映射
        sortedMargins.forEach((margin, index) => {
            marginMap.set(margin, index + 1);
        });

        // 第二遍扫描：添加层级class
        processedHtml = processedHtml.replace(/<p([^>]*?)style="([^"]*)"([^>]*?)>/gi, (match, before, style, after) => {
            if (style.includes('mso-list:')) {
                const marginMatch = style.match(/margin-left:\s*([0-9.]+[a-z%]*)/i);
                if (marginMatch) {
                    const margin = marginMatch[1];
                    const level = marginMap.get(margin);
                    const levelClass = `list-level-${level}`;

                    // 检查是否已经有class属性
                    if (match.includes('class="')) {
                        return match.replace(/class="([^"]*)"/, `class="$1 ${levelClass}"`);
                    } else {
                        return `<p${before}style="${style}" class="${levelClass}"${after}>`;
                    }
                }
            }
            return match;
        });

        return processedHtml;
    } catch (error) {
        console.error('添加列表层级class错误:', error);
        return html;
    }
}

/**
 * 获取列表类型
 * @param {Element} listItem - 列表项元素
 * @returns {string} - 'ol' 或 'ul'
 */
function getListType(listItem) {
    const markerSpan = listItem.querySelector('span[style*="mso-list:Ignore"]');
    if (!markerSpan) return 'ul';

    const marker = markerSpan.textContent.trim();

    // 检查是否是字母列表（a. b. c. 或 A. B. C.）
    if (/^[a-zA-Z][\.\)]/.test(marker)) {
        return 'ol';
    }

    // 检查是否是数字列表（1. 2. 3.）
    if (/^[0-9]+\./.test(marker)) {
        return 'ol';
    }

    // 检查是否是罗马数字列表（i. ii. iii. 或 I. II. III.）
    if (/^[ivxlcdmIVXLCDM]+\./.test(marker)) {
        return 'ol';
    }

    // 检查是否是中文数字列表
    if (/^[\u3007\u4e00-\u4e5d\u58f1-\u62fe]+\./.test(marker)) {
        return 'ol';
    }

    // 检查是否是项目符号列表（•, ·, §, ○ 等）
    if (/^[\u2022\u00b7\u00a7\u25CF\u25CB\u25E6\u2023\u2043]/.test(marker)) {
        return 'ul';
    }

    // 默认返回无序列表
    return 'ul';
}

function convertMsoListToNestedLists(html) {
    if (!html) return '';

    try {
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
        const document = dom.window.document;

        // 获取所有带有mso-list的段落
        const listParagraphs = Array.from(document.querySelectorAll('p[style*="mso-list"]'));
        if (listParagraphs.length === 0) return html;

        // 创建根列表
        let rootList = null;
        let currentList = null;
        let previousLevel = 0;
        let currentListType = 'ul'; // 默认列表类型

        // 处理每个段落
        for (let i = 0; i < listParagraphs.length; i++) {
            const para = listParagraphs[i];

            // 获取列表级别
            const levelMatch = para.getAttribute('class')?.match(/list-level-(\d+)/);
            const level = levelMatch ? parseInt(levelMatch[1]) : 1;

            // 确定列表类型
            const listType = getListType(para);

            // 获取列表标记和内容
            const markerSpan = para.querySelector('span[style*="mso-list:Ignore"]');
            if (markerSpan) {
                markerSpan.parentNode.removeChild(markerSpan);
            }

            // 获取段落内容
            const content = para.innerHTML.trim();

            // 创建列表项
            const li = document.createElement('li');
            li.innerHTML = content;

            if (level === 1) {
                // 第一层级，创建新的根列表
                if (!rootList || previousLevel === 0) {
                    rootList = document.createElement(listType);
                    para.parentNode.insertBefore(rootList, para);
                    currentList = rootList;
                    currentListType = listType;
                }
                currentList = rootList;
                currentList.appendChild(li);
            } else {
                // 处理子层级
                if (level > previousLevel) {
                    // 创建新的子列表
                    const subList = document.createElement(listType);
                    const lastItem = currentList.lastElementChild;
                    if (lastItem) {
                        lastItem.appendChild(subList);
                        currentList = subList;
                        currentListType = listType;
                    }
                } else if (level < previousLevel) {
                    // 返回上层列表
                    for (let j = 0; j < (previousLevel - level); j++) {
                        currentList = currentList.parentElement.parentElement;
                    }
                } else if (level === previousLevel && listType !== currentListType) {
                    // 同级但类型不同，创建新列表
                    const newList = document.createElement(listType);
                    currentList.parentElement.appendChild(newList);
                    currentList = newList;
                    currentListType = listType;
                }
                currentList.appendChild(li);
            }

            // 更新前一个层级
            previousLevel = level;

            // 移除原始段落
            para.parentNode.removeChild(para);
        }

        // 返回处理后的HTML
        return document.body.innerHTML;
    } catch (error) {
        console.error('转换列表错误:', error);
        return html;
    }
}

/**
 * 清理多余的span标签
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function cleanRedundantSpans(html) {
    if (!html) return '';

    try {
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
        const document = dom.window.document;

        /**
         * 检查span是否只包含下划线样式
         * @param {Element} span - span元素
         * @returns {boolean} - 是否只包含下划线样式
         */
        function hasOnlyUnderlineStyle(span) {
            const style = span.getAttribute('style');
            if (!style) return false;

            const styles = style.split(';').map(s => s.trim()).filter(s => s);
            return styles.length === 1 && /^text-decoration:(?:[^;]*)?underline/.test(styles[0]);
        }

        /**
         * 检查span是否只包含加粗样式
         * @param {Element} span - span元素
         * @returns {boolean} - 是否只包含加粗样式
         */
        function hasOnlyBoldStyle(span) {
            const style = span.getAttribute('style');
            if (!style) return false;

            const styles = style.split(';').map(s => s.trim()).filter(s => s);
            return styles.length === 1 && /^font-weight:\s*bold/.test(styles[0]);
        }

        /**
         * 处理元素中的span标签
         * @param {Element} element - 要处理的元素
         */
        function processSpans(element) {
            // 如果元素没有子节点，直接返回
            if (!element.hasChildNodes()) {
                return;
            }

            // 处理所有子节点
            const childNodes = Array.from(element.childNodes);
            let newContent = '';

            for (let node of childNodes) {
                if (node.nodeType === 3) { // 文本节点
                    newContent += node.textContent;
                } else if (node.nodeType === 1) { // 元素节点
                    if (node.tagName === 'SPAN') {
                        const parentTag = node.parentElement.tagName;
                        const onlyUnderline = hasOnlyUnderlineStyle(node);
                        const onlyBold = hasOnlyBoldStyle(node);

                        if ((parentTag === 'U' && onlyUnderline) ||
                            ((parentTag === 'B' || parentTag === 'STRONG') && onlyBold)) {
                            // 如果span在u标签内且只有下划线样式，或在b/strong标签内且只有加粗样式
                            // 只保留内容
                            newContent += node.innerHTML;
                        } else if (node.querySelector('*')) {
                            // 如果span内有其他标签，递归处理后保留
                            processSpans(node);
                            newContent += node.outerHTML;
                        } else if (!node.textContent.trim()) {
                            // 如果是空的span，不添加任何内容
                            continue;
                        } else {
                            // 其他情况保留span的内容
                            newContent += node.textContent;
                        }
                    } else {
                        // 对于非span标签，递归处理其内容
                        processSpans(node);
                        newContent += node.outerHTML;
                    }
                }
            }

            element.innerHTML = newContent;
        }

        /**
         * 清理空的b标签
         * @param {Element} element - 要处理的元素
         */
        function cleanEmptyBTags(element) {
            const bTags = element.getElementsByTagName('b');
            const strongTags = element.getElementsByTagName('strong');

            // 从后向前遍历，这样在删除节点时不会影响到索引
            for (let i = bTags.length - 1; i >= 0; i--) {
                const tag = bTags[i];
                if (!tag.textContent.trim()) {
                    tag.parentNode.removeChild(tag);
                }
            }

            for (let i = strongTags.length - 1; i >= 0; i--) {
                const tag = strongTags[i];
                if (!tag.textContent.trim()) {
                    tag.parentNode.removeChild(tag);
                }
            }
        }

        // 处理整个文档
        processSpans(document.body);
        cleanEmptyBTags(document.body);

        return document.body.innerHTML;
    } catch (error) {
        console.error('清理list内冗余span标签错误:', error);
        return html;
    }
}

/**
 * 清理标题标签的margin-left样式
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function cleanHeadingMargins(html) {
    if (!html) return '';

    try {
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
        const document = dom.window.document;

        // 处理所有h1-h6标签
        for (let i = 1; i <= 6; i++) {
            const headings = document.getElementsByTagName(`h${i}`);
            for (let heading of headings) {
                // 清理margin-left样式
                const style = heading.getAttribute('style');
                if (style) {
                    // 移除margin-left样式，保留其他样式
                    const styles = style.split(';')
                        .map(s => s.trim())
                        .filter(s => s && !s.toLowerCase().startsWith('margin-left:'));

                    if (styles.length > 0) {
                        heading.setAttribute('style', styles.join('; '));
                    } else {
                        heading.removeAttribute('style');
                    }
                }

                // 清理标题内的b和strong标签
                const boldElements = heading.querySelectorAll('b, strong');
                boldElements.forEach(boldElement => {
                    // 创建文本节点替换b/strong标签
                    const textContent = boldElement.textContent;
                    const textNode = document.createTextNode(textContent);
                    boldElement.parentNode.replaceChild(textNode, boldElement);
                });
            }
        }

        return document.body.innerHTML;
    } catch (error) {
        console.error('清理标题margin-left错误:', error);
        return html;
    }
}

/**
 * 清理所有标签的align属性
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function cleanAlignAttributes(html) {
    if (!html) return '';

    try {
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
        const document = dom.window.document;

        // 获取所有元素
        const allElements = document.getElementsByTagName('*');
        for (let element of allElements) {
            // 移除align属性
            if (element.hasAttribute('align')) {
                element.removeAttribute('align');
            }
        }

        return document.body.innerHTML;
    } catch (error) {
        console.error('清理align属性错误:', error);
        return html;
    }
}

/**
 * 结构化内容，使用section包围标题及其内容
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function structureContent(html) {
    if (!html) return '';

    try {
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
        const document = dom.window.document;

        // 获取所有标题元素
        const headings = Array.from(document.querySelectorAll('h1, h2, h3, h4, h5, h6'));
        if (headings.length === 0) {
            return html;
        }

        // 处理每个标题及其内容
        const processedNodes = [];
        for (let i = 0; i < headings.length; i++) {
            const currentHeading = headings[i];
            const nextHeading = headings[i + 1];

            // 创建section
            const section = document.createElement('section');

            // 创建title并保持HTML结构
            const titleContent = currentHeading.innerHTML;
            section.innerHTML = `<title>${titleContent}</title>`;

            // 收集当前标题到下一个标题之间的所有内容
            let currentNode = currentHeading.nextSibling;
            let hasContent = false;

            while (currentNode && (!nextHeading || !currentNode.isSameNode(nextHeading))) {
                if (currentNode.nodeType === 1) { // 元素节点
                    const clonedNode = currentNode.cloneNode(true);
                    if (clonedNode.textContent.trim()) {
                        section.appendChild(clonedNode);
                        hasContent = true;
                    }
                }
                currentNode = currentNode.nextSibling;
            }

            // 将原始标题移除
            currentHeading.parentNode.removeChild(currentHeading);

            // 只有当section包含标题文本或其他内容时才添加
            if (titleContent || hasContent) {
                processedNodes.push(section);
            }
        }

        // 清空body
        document.body.innerHTML = '';

        // 添加所有处理后的节点
        processedNodes.forEach(node => {
            document.body.appendChild(node);
        });

        return document.body.innerHTML.trim();
    } catch (error) {
        console.error('内容结构化错误:', error);
        return html;
    }
}

/**
 * 清理表格，移除所有属性并添加必要的属性
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function cleanTables(html) {
    if (!html) return '';

    try {
        // 处理表格标签和属性
        return html.replace(/<table[^>]*>[\s\S]*?<\/table>/gi, (tableMatch) => {
            // 使用JSDOM解析表格结构
            const dom = new JSDOM(`<!DOCTYPE html><html><body>${tableMatch}</body></html>`);
            const document = dom.window.document;
            const table = document.querySelector('table');

            if (!table) return tableMatch;

            // 分析表格结构
            const tbodies = table.getElementsByTagName('tbody');
            if (tbodies.length === 0) return tableMatch;

            // 获取第一个tbody的结构信息
            const tbody = tbodies[0];
            const rows = tbody.getElementsByTagName('tr');
            if (rows.length === 0) return tableMatch;

            // 计算列数
            const firstRow = rows[0];
            const colCount = firstRow.getElementsByTagName('td').length ||
                firstRow.getElementsByTagName('th').length;

            if (colCount === 0) return tableMatch;

            // 使用字符串方式构建新的表格
            let newTable = '<table frame="all" rowsep="1" colsep="1">';

            // 构建tgroup
            newTable += `<tgroup cols="${colCount}">`;

            // 添加colspec元素
            for (let i = 1; i <= colCount; i++) {
                newTable += `<colspec colnum="${i}" colname="col${i}"/>`;
            }

            // 构建tbody
            newTable += '<tbody>';

            // 处理每一行
            for (let row of rows) {
                newTable += '<row>';
                const cells = row.getElementsByTagName('td');
                const headerCells = row.getElementsByTagName('th');

                // 处理单元格
                if (cells.length > 0) {
                    // 确保每行都有正确数量的单元格
                    for (let i = 0; i < colCount; i++) {
                        const cell = cells[i];
                        if (cell) {
                            // 检查是否是只包含空p标签的单元格
                            const pTags = cell.getElementsByTagName('p');
                            const hasOnlyEmptyPTags = pTags.length > 0 &&
                                Array.from(pTags).every(p => !p.textContent.trim() && p.children.length === 0);

                            if (hasOnlyEmptyPTags) {
                                // 如果只包含空p标签，使用<entry></entry>
                                newTable += '<entry></entry>';
                                continue;
                            }

                            // 清理单元格内容，移除不必要的包裹标签
                            let cellContent = '';

                            if (pTags.length > 0) {
                                // 如果有p标签，提取其文本内容
                                for (let p of pTags) {
                                    // 如果p标签直接包含文本或只包含简单的格式化标签(b, i, u等)
                                    if (p.children.length === 0 ||
                                        Array.from(p.children).every(child =>
                                            ['B', 'I', 'U', 'STRONG', 'EM'].includes(child.tagName))) {
                                        cellContent += p.innerHTML.trim();
                                    } else {
                                        // 如果p标签包含其他复杂结构，保留原始HTML
                                        cellContent += p.outerHTML;
                                    }
                                }
                            } else {
                                // 如果没有p标签，使用原始内容
                                cellContent = cell.innerHTML;
                            }

                            // 移除可能的连续空格和换行，但保留空内容
                            cellContent = cellContent.replace(/\s+/g, ' ').trim();

                            // 包装成entry标签
                            newTable += `<entry>${cellContent}</entry>`;
                        } else {
                            // 如果这个位置没有单元格，添加空的entry
                            newTable += '<entry></entry>';
                        }
                    }
                } else if (headerCells.length > 0) {
                    // 对表头行做同样的处理
                    for (let i = 0; i < colCount; i++) {
                        const cell = headerCells[i];
                        if (cell) {
                            // 检查是否是只包含空p标签的单元格
                            const pTags = cell.getElementsByTagName('p');
                            const hasOnlyEmptyPTags = pTags.length > 0 &&
                                Array.from(pTags).every(p => !p.textContent.trim() && p.children.length === 0);

                            if (hasOnlyEmptyPTags) {
                                // 如果只包含空p标签，使用<entry></entry>
                                newTable += '<entry></entry>';
                                continue;
                            }

                            let cellContent = '';
                            if (pTags.length > 0) {
                                for (let p of pTags) {
                                    if (p.children.length === 0 ||
                                        Array.from(p.children).every(child =>
                                            ['B', 'I', 'U', 'STRONG', 'EM'].includes(child.tagName))) {
                                        cellContent += p.innerHTML.trim();
                                    } else {
                                        cellContent += p.outerHTML;
                                    }
                                }
                            } else {
                                cellContent = cell.innerHTML;
                            }

                            cellContent = cellContent.replace(/\s+/g, ' ').trim();
                            newTable += `<entry>${cellContent}</entry>`;
                        } else {
                            // 如果这个位置没有单元格，添加空的entry
                            newTable += '<entry></entry>';
                        }
                    }
                } else {
                    // 如果这一行既没有td也没有th，添加空的entry填充
                    for (let i = 0; i < colCount; i++) {
                        newTable += '<entry></entry>';
                    }
                }
                newTable += '</row>';
            }

            newTable += '</tbody></tgroup></table>';
            return newTable;
        });
    } catch (error) {
        console.error('清理表格错误:', error);
        return html;
    }
}

module.exports = {
    cleanHtml,
    formatHtml,
    cleanClassAndIdAttributes,
    cleanEmptyTags,
    cleanSelectiveStyles,
    getListLevel,
    analyzeListItem,
    addListLevelClasses,
    convertMsoListToNestedLists,
    cleanRedundantSpans,
    cleanHeadingMargins,
    cleanAlignAttributes,
    structureContent,
    cleanTables
};

