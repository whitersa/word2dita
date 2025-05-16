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


        // 生成debug文件以查看第一阶段处理结果
        // 4. 清理样式，只保留指定样式
        html = cleanSelectiveStyles(html);

        // 5. 添加列表层级class
        html = addListLevelClasses(html);

        // 6. 清理align属性
        html = cleanAlignAttributes(html);
        html = html.replace(/<font[^>]*>([\s\S]*?)<\/font>/gi, '$1')

        // 7. 转换为嵌套列表
        html = convertMsoListToNestedLists(html);


        // 8. 清理特殊字符
        html = html
            // 替换 non-breaking spaces
            .replace(/&nbsp;/g, ' ')

        // 9. 移除空属性
        html = html
            .replace(/(\w+)="\s*"/g, '')
            .replace(/\s+>/g, '>');


        // 处理table，使用jsdom解析结构解析出一些参数信息，然后转换成dita的表格，因为jsdom会自动纠正不合法的标签，所以内容处理必须由字符串替换来实现

        // 11. 最后处理表格，避免被JSDOM再次序列化
        // html = cleanTables(html);
        // div标签转换成p标签
        html = html.replace(/<div[^>]*>/gi, '<p>');
        html = html.replace(/<\/div>/gi, '</p>');

        try {
            require('fs').writeFileSync('debug_stage1.html', html);
        } catch (err) {
            console.error('Debug file write error:', err);
        }
        // 12. 清理标签
        html = cleanEmptyTags(html);

        html = html
            // 移除多余的换行标签
            .replace(/<br>\s*<br>/gi, '<br>')
            // 移除末尾的换行
            .replace(/<br>$/i, '')
            // 处理font标签，保留内容

            // 处理span标签，保留内容
            .replace(/<span[^>]*>([\s\S]*?)<\/span>/gi, '$1')
            // 处理s(删除线)标签，精确匹配并保留内容
            .replace(/<s>([\s\S]*?)<\/s>/gi, '$1');
        // 清理属性
        html = cleanClassAndIdAttributes(html);



        // 处理jsdom无法处理的特殊情况，所有需要规范html的处理都只能在此之前
        html = doExtraTransformJsdomCantHandle(html);



        return html;
    } catch (error) {
        console.error('HTML清理错误:', error);
        return html;
    }
}


/**
 * 处理jsdom无法处理的特殊情况
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function doExtraTransformJsdomCantHandle(html) {
    html = processTables(html);
    // 处理a标签，转换成带有scope=external的xref标签, 包括闭合a标签
    html = html
        .replace(/<a[^>]*href="([^"]+)"[^>]*>/gi, '<xref scope="external" format="html" href="$1">')
        .replace(/<\/a>/gi, '</xref>');
    // 3. 处理标题标签
    html = cleanHeadingTags(html);
    return html;
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
 * 清理 class、id、align、valign、data-* 等属性
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function cleanClassAndIdAttributes(html) {
    if (!html) return '';
    // 移除 class、id、align、valign、data-* 属性（不区分单双引号/无值）
    return html
        // 匹配 class="..."、id='...'、align=...、valign=...、data-xxx=...（单双引号或无引号）
        .replace(/\s+(class|id|align|valign|data-[^=\s]*)\s*=\s*(['"]).*?\2/gi, '')
        .replace(/\s+(class|id|align|valign|data-[^=\s]*)\s*=\s*[^ >]+/gi, '');
}

/**
 * 清理空标签，但保留可能影响布局的标签
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function cleanEmptyTags(html) {
    if (!html) return '';

    try {
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
        const document = dom.window.document;

        // 需要保留的标签列表（即使为空也不删除）
        const preserveTags = new Set([
            'table', 'thead', 'tbody', 'tr', 'td', 'th',  // 表格相关
            'ul', 'ol', 'li',  // 列表相关
        ]);

        // 递归处理元素
        function processElement(element) {
            // 如果元素没有子节点，直接返回
            if (!element.hasChildNodes()) {
                return;
            }

            const children = Array.from(element.childNodes);
            for (const child of children) {
                if (child.nodeType === 1) { // 元素节点
                    const tagName = child.tagName.toLowerCase();

                    // 递归处理子元素
                    processElement(child);

                    // 检查元素是否为空（没有文本内容和子元素）
                    const isEmpty = !child.textContent.trim() &&
                        !child.querySelector('img') &&
                        !child.querySelector('br');

                    // 如果是空元素且不在保留列表中，删除它
                    if (isEmpty && !preserveTags.has(tagName)) {
                        child.parentNode.removeChild(child);
                    }
                }
            }
        }

        // 处理整个文档
        processElement(document.body);

        return document.body.innerHTML;
    } catch (error) {
        console.error('清理空标签错误:', error);
        return html;
    }
}


/**
 * 清理样式，只保留加粗、斜体、下划线、msolist相关的样式，以及msolist元素的marginleft
 * @param {string} html - 要清理的HTML
 * @returns {string} - 清理后的HTML
 */
function cleanSelectiveStyles(html) {
    if (!html) return '';

    try {
        // 处理所有带有style属性的标签
        html = html.replace(/style="([^"]*)"[^>]*>/gi, (match, styles) => {
            const styleDeclarations = styles.split(';');
            const preserved = [];
            let hasMsoList = false;

            // 先检查是否有mso-list样式
            hasMsoList = styleDeclarations.some(declaration =>
                /^mso-list:/i.test(declaration.trim())
            );

            // 处理每个样式声明
            for (let declaration of styleDeclarations) {
                declaration = declaration.trim();
                if (!declaration) continue;

                // 检查是否是要保留的样式
                if (
                    /^mso-list:/i.test(declaration) || // 列表相关
                    (hasMsoList && /^margin-left:/i.test(declaration)) // 只有mso-list元素的margin-left才保留
                ) {
                    preserved.push(declaration);
                }
            }

            // 如果有要保留的样式，返回新的style属性
            return preserved.length > 0 ? `style="${preserved.join('; ')}">` : '>';
        });

        return html;
    } catch (error) {
        console.error('样式清理错误:', error);
        return html;
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
 * 判断是否是新的列表组
 * @param {Element} currentPara - 当前段落
 * @param {Element} previousPara - 前一个段落
 * @returns {boolean} - 是否是新的列表组
 */
function isNewListGroup(currentPara, previousPara) {
    if (!previousPara) return true;

    // 获取两个段落之间的非列表内容
    let node = previousPara.nextSibling;
    let hasNonListContent = false;
    while (node && node !== currentPara) {
        if (node.nodeType === 1 && // 元素节点
            !node.getAttribute('style')?.includes('mso-list')) {
            const text = node.textContent.trim();
            if (text) {
                hasNonListContent = true;
                break;
            }
        }
        node = node.nextSibling;
    }

    return hasNonListContent;
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
        const body = document.body;

        // 新建一个数组用于收集最终的节点顺序
        const newNodes = [];
        let buffer = [];

        // 工具函数：处理buffer为嵌套列表
        function bufferToList(buffer) {
            if (buffer.length === 0) return null;
            // 复用原有的嵌套逻辑
            let rootList = null;
            let currentList = null;
            let previousLevel = 0;
            let currentListType = 'ul';
            for (let i = 0; i < buffer.length; i++) {
                const para = buffer[i];
                const levelMatch = para.getAttribute('class')?.match(/list-level-(\d+)/);
                const level = levelMatch ? parseInt(levelMatch[1]) : 1;
                const listType = getListType(para);
                const markerSpan = para.querySelector('span[style*="mso-list:Ignore"]');
                if (markerSpan) {
                    markerSpan.parentNode.removeChild(markerSpan);
                }
                const content = para.innerHTML.trim();
                const li = document.createElement('li');
                li.innerHTML = content;
                if (level === 1) {
                    if (!rootList || previousLevel === 0) {
                        rootList = document.createElement(listType);
                        currentList = rootList;
                        currentListType = listType;
                    }
                    currentList = rootList;
                    if (currentList) currentList.appendChild(li);
                } else {
                    if (level > previousLevel) {
                        const subList = document.createElement(listType);
                        const lastItem = currentList ? currentList.lastElementChild : null;
                        if (lastItem) {
                            lastItem.appendChild(subList);
                            currentList = subList;
                            currentListType = listType;
                        } else if (currentList) {
                            // fallback: 没有lastItem但有currentList
                            currentList.appendChild(subList);
                            currentList = subList;
                            currentListType = listType;
                        } else {
                            // fallback: currentList为null，直接新建根列表
                            currentList = subList;
                            currentListType = listType;
                            if (!rootList) rootList = currentList;
                        }
                    } else if (level < previousLevel) {
                        for (let j = 0; j < (previousLevel - level); j++) {
                            if (currentList && currentList.parentElement && currentList.parentElement.parentElement) {
                                currentList = currentList.parentElement.parentElement;
                            }
                        }
                    } else if (level === previousLevel && listType !== currentListType) {
                        const newList = document.createElement(listType);
                        if (currentList && currentList.parentElement) {
                            currentList.parentElement.appendChild(newList);
                            currentList = newList;
                            currentListType = listType;
                        } else if (!currentList) {
                            currentList = newList;
                            currentListType = listType;
                            if (!rootList) rootList = currentList;
                        }
                    }
                    if (currentList) {
                        currentList.appendChild(li);
                    }
                }
                previousLevel = level;
            }
            return rootList;
        }

        // 顺序遍历body的所有子节点
        let node = body.firstChild;
        while (node) {
            const nextNode = node.nextSibling; // 先保存下一个节点
            if (node.nodeType === 1 && node.tagName.toLowerCase() === 'p' && node.getAttribute('style') && node.getAttribute('style').includes('mso-list')) {
                buffer.push(node);
            } else {
                if (buffer.length) {
                    const list = bufferToList(buffer);
                    if (list) newNodes.push(list);
                    buffer = [];
                }
                newNodes.push(node);
            }
            node = nextNode;
        }
        // 处理结尾的buffer
        if (buffer.length) {
            const list = bufferToList(buffer);
            if (list) newNodes.push(list);
        }

        // 清空body，按顺序插入新节点
        body.innerHTML = '';
        for (const n of newNodes) {
            body.appendChild(n);
        }

        return body.innerHTML;
    } catch (error) {
        console.error('转换列表错误:', error);
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
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`, {
            // 设置选项以保持原始HTML格式
            includeNodeLocations: true
        });
        const document = dom.window.document;

        // 获取所有带有align属性的元素
        const elementsWithAlign = document.querySelectorAll('[align]');
        if (elementsWithAlign.length === 0) {
            return html; // 如果没有align属性，直接返回原始HTML
        }

        // 记录所有需要处理的元素的原始状态
        const modifications = Array.from(elementsWithAlign).map(element => {
            const tempDiv = document.createElement('div');
            tempDiv.appendChild(element.cloneNode(true));
            return {
                original: tempDiv.innerHTML,
                element: element
            };
        });

        // 移除align属性
        elementsWithAlign.forEach(element => {
            element.removeAttribute('align');
        });

        // 使用字符串替换而不是整个DOM序列化
        let result = html;
        for (const mod of modifications) {
            const tempDiv = document.createElement('div');
            tempDiv.appendChild(mod.element.cloneNode(true));
            const newContent = tempDiv.innerHTML;
            result = result.replace(mod.original, newContent);
        }

        return result;
    } catch (error) {
        console.error('清理align属性错误:', error);
        return html;
    }
}

/**
 * 清理表格，移除所有属性并添加必要的属性
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function processTables(html) {
    if (!html) return '';

    try {
        // 1. 匹配所有<table ...>...</table>，支持跨行
        const tableRegex = /<table[\s\S]*?<\/table>/gi;
        let tableIndex = 0;
        const tableMap = new Map(); // id => entryTableString
        let match;
        let htmlWithPlaceholders = html;

        // 2. 遍历所有表格，分析并替换为占位符
        while ((match = tableRegex.exec(html)) !== null) {
            const tableHtml = match[0];
            tableIndex++;
            const tableId = `__TABLE_PLACEHOLDER_${tableIndex}__`;
            // 用jsdom分析表格结构，生成entry格式
            let entryTableString = '';
            try {
                const dom = new JSDOM(`<!DOCTYPE html><html><body>${tableHtml}</body></html>`);
                const document = dom.window.document;
                const table = document.querySelector('table');
                if (!table) {
                    entryTableString = tableHtml; // fallback
                } else {
                    // 计算列数
                    const colCount = (() => {
                        const firstRow = table.querySelector('tr');
                        if (!firstRow) return 0;
                        // 统计所有colspan后的最大列数
                        let maxCols = 0;
                        let rows = table.querySelectorAll('tr');
                        for (let row of rows) {
                            let count = 0;
                            for (let cell of row.querySelectorAll('td,th')) {
                                count += parseInt(cell.getAttribute('colspan') || '1', 10);
                            }
                            if (count > maxCols) maxCols = count;
                        }
                        return maxCols;
                    })();
                    if (colCount === 0) {
                        entryTableString = tableHtml;
                    } else {
                        let newTable = '<table frame="all" rowsep="1" colsep="1">';
                        newTable += `<tgroup cols="${colCount}">`;
                        for (let i = 1; i <= colCount; i++) {
                            newTable += `<colspec colnum="${i}" colname="col${i}"/>`;
                        }
                        // 处理thead和tbody
                        const thead = table.querySelector('thead');
                        const tbody = table.querySelector('tbody');
                        let headRows = [];
                        let bodyRows = [];
                        if (thead) {
                            headRows = Array.from(thead.querySelectorAll('tr'));
                            if (tbody) {
                                bodyRows = Array.from(tbody.querySelectorAll('tr'));
                            } else {
                                bodyRows = Array.from(table.querySelectorAll('tr')).filter(tr => !thead.contains(tr));
                            }
                        } else {
                            bodyRows = Array.from(table.querySelectorAll('tr'));
                        }
                        // 生成thead
                        if (headRows.length) {
                            newTable += '<thead>';
                            newTable += generateCalsRows(headRows, colCount);
                            newTable += '</thead>';
                        }
                        // 生成tbody
                        if (bodyRows.length) {
                            newTable += '<tbody>';
                            newTable += generateCalsRows(bodyRows, colCount);
                            newTable += '</tbody>';
                        }
                        newTable += '</tgroup></table>';
                        entryTableString = newTable;
                    }
                }
            } catch (err) {
                entryTableString = tableHtml;
            }
            // 记录映射
            tableMap.set(tableId, entryTableString);
            // 用占位符替换原始表格
            htmlWithPlaceholders = htmlWithPlaceholders.replace(tableHtml, `<table id="${tableId}"></table>`);
        }

        // 3. 替换所有占位符为entry格式表格串
        let finalHtml = htmlWithPlaceholders;
        for (const [tableId, entryTableString] of tableMap.entries()) {
            finalHtml = finalHtml.replace(`<table id="${tableId}"></table>`, entryTableString);
        }
        return finalHtml;
    } catch (error) {
        console.error('表格处理错误:', error);
        return html;
    }
}

// 生成CALS行，支持rowspan/colspan合并
function generateCalsRows(rows, colCount) {
    // 占位矩阵，记录哪些格子被rowspan/colspan占用
    const occupied = [];
    let result = '';
    let colNames = Array.from({ length: colCount }, (_, i) => `col${i + 1}`);
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        const row = rows[rowIndex];
        result += '<row>';
        let cells = Array.from(row.querySelectorAll('td,th'));
        let col = 0;
        for (let cellIndex = 0; cellIndex < cells.length; cellIndex++) {
            // 跳过被占用的格子
            while (occupied[rowIndex]?.[col]) col++;
            const cell = cells[cellIndex];
            const rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10);
            const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
            let entryAttrs = '';
            if (rowspan > 1) entryAttrs += ` morerows="${rowspan - 1}"`;
            if (colspan > 1) entryAttrs += ` namest="${colNames[col]}" nameend="${colNames[col + colspan - 1]}"`;
            // 标记被占用的格子
            for (let r = 0; r < rowspan; r++) {
                for (let c = 0; c < colspan; c++) {
                    if (!occupied[rowIndex + r]) occupied[rowIndex + r] = [];
                    occupied[rowIndex + r][col + c] = true;
                }
            }
            // 处理内容
            const pTags = cell.getElementsByTagName('p');
            const hasOnlyEmptyPTags = pTags.length > 0 &&
                Array.from(pTags).every(p => !p.textContent.trim() && p.children.length === 0);
            let cellContent = '';
            if (hasOnlyEmptyPTags) {
                // 空内容
            } else if (pTags.length > 0) {
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
            result += `<entry${entryAttrs}>${cellContent}</entry>`;
            col += colspan;
        }
        // 补齐空格
        while (col < colCount) {
            result += '<entry></entry>';
            col++;
        }
        result += '</row>';
    }
    return result;
}

/**
 * 处理标题标签，将h1转换为title标签，h2-h6转换为b标签
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function cleanHeadingTags(html) {
    if (!html) return '';

    try {
        // 移除所有带有mso-list:Ignore的span及其内容
        html = html.replace(/<span[^>]*style="[^"]*mso-list:Ignore[^"]*"[^>]*>[\s\S]*?<\/span>/gi, '');

        // 将h2-h6转换为b
        html = html
            .replace(/<h[2-6][^>]*>([\s\S]*?)<\/h[2-6]>/gi, '<b>$1</b>');

        // 检查是否有<h1>
        const h1Match = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
        if (h1Match) {
            // 提取h1内容
            const h1Content = h1Match[1];
            // 替换第一个h1为<title>
            html = html.replace(/<h1[^>]*>[\s\S]*?<\/h1>/i, `<title>${h1Content}</title>`);
            // 用<section>包裹整体内容
            html = `<section>${html}</section>`;
        }
        return html;
    } catch (error) {
        console.error('处理标题标签错误:', error);
        return html;
    }
}

module.exports = {
    cleanHtml,
    formatHtml,

};


