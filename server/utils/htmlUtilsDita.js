const { JSDOM } = require('jsdom');

/**
 * 清理和规范化 HTML 内容
 * @param {string} html - 要清理的HTML
 * @returns {string} - 清理后的HTML
 */
function cleanHtml(html) {
    if (!html) return '';

    try {
        // =================================================================
        // 阶段 1: 基础文本清理
        // 目标: 移除注释、空行、规范化空格，为后续处理提供干净的输入
        // =================================================================
        html = basicTextCleanup(html);

        // =================================================================
        // 阶段 2: 移除不需要的标签和属性
        // 目标: 移除脚本、样式块、XML命名空间、Word特有标记等
        // =================================================================
        html = removeUnwantedTags(html);

        // 清理样式，只保留指定样式 (加粗、斜体、下划线、列表样式、宽度)
        html = cleanSelectiveStyles(html);

        // =================================================================
        // 阶段 3: 列表处理
        // 目标: 识别 Word 的 mso-list 样式，转换为标准的 HTML 列表结构
        // =================================================================
        
        // 添加列表层级class
        html = addListLevelClasses(html);

        // 清理align属性
        html = cleanAlignAttributes(html);
        
        html = removeFontTags(html);

        // 转换为嵌套列表
        html = convertMsoListToNestedLists(html);


        // =================================================================
        // 阶段 4: 进一步清理和规范化
        // =================================================================

        // 清理特殊字符
        html = cleanSpecialCharacters(html);

        // 移除空属性
        html = cleanEmptyAttributes(html);


        // =================================================================
        // 阶段 5: 结构规范化与样式清理
        // 目标: 转换 div 为 p，清理样式标签，移除空标签等
        // =================================================================

        // div标签转换成p标签
        html = convertDivToP(html);

        // 清理样式（提前执行，将样式转换为标签，防止被后续移除span的操作误删）
        // 注意: 表格宽度样式在此步骤中被保留，供 processTables 使用
        html = cleanAllStyles(html);

        // 清理标签
        html = cleanEmptyTags(html);

        html = cleanLineBreaksAndDecorations(html);
        
        // 处理span标签，保留内容
        html = removeSpans(html);

        // 清理属性
        html = cleanClassAndIdAttributes(html);

        // =================================================================
        // 阶段 6: 最终转换和还原
        // 目标: 处理 JSDOM 无法处理的特殊情况，还原 DITA 标签
        // =================================================================

        // 处理jsdom无法处理的特殊情况
        // 包括: 表格转换 (processTables), 链接转换 (xref), 标题转换 (title/section)
        html = doExtraTransformJsdomCantHandle(html);

        // 合并连续的内联标签
        html = mergeConsecutiveInlineTags(html);

        // 还原 DITA 标签
        html = restoreDitaTags(html);

        return html;
    } catch (error) {
        console.error('HTML清理错误:', error);
        return html;
    }
}


/**
 * 阶段 1: 基础文本清理
 * 移除注释、空行、规范化空格
 */
function basicTextCleanup(html) {
    return html
        .replace(/<!--[\s\S]*?-->/g, '')
        .replace(/^\s*[\r\n]/gm, '')
        .replace(/\s+/g, ' ');
}

/**
 * 阶段 2: 移除不需要的标签和属性
 * 移除脚本、样式块、XML命名空间、Word特有标记等
 */
function removeUnwantedTags(html) {
    return html
        .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
        .replace(/<(!|script[^>]*>[\s\S]*?<\/script(?=[>\s])|\/?(\?xml(:\w+)?|img|meta|link|\w:\w+)(?=[\s\/>]))[^>]*>/gi, '')
        .replace(/<html[^>]*>/i, '<html>')
        .replace(/<\/?[a-z]*:[^>]*>/gi, '')
        .replace(/class="?Mso[a-zA-Z]+"/g, '')
        .replace(/<div[^>]*tdoc-data-src="[^"]*"[^>]*>[\s\S]*?<\/div>/gi, '');
}

/**
 * 清理样式，只保留加粗、斜体、下划线、msolist相关的样式，以及msolist元素的marginleft
 * @param {string} html - 要清理的HTML
 * @returns {string} - 处理后的HTML
 */
function cleanSelectiveStyles(html) {
    if (!html) return '';

    try {
        // 处理所有带有style属性的标签 (支持单引号和双引号)
        // 只替换style属性本身，不吞掉后续属性
        html = html.replace(/style=(["'])([\s\S]*?)\1/gi, (match, quote, styles) => {
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
                    (hasMsoList && /^margin-left:/i.test(declaration)) || // 只有mso-list元素的margin-left才保留
                    /^font-weight:\s*bold/i.test(declaration) || // 加粗
                    /^font-style:\s*italic/i.test(declaration) || // 斜体
                    /^text-decoration:\s*underline/i.test(declaration) || // 下划线
                    /^width:/i.test(declaration) // 宽度
                ) {
                    preserved.push(declaration);
                }
            }

            // 如果有要保留的样式，返回新的style属性
            return preserved.length > 0 ? `style=${quote}${preserved.join('; ')}${quote}` : '';
        });

        return html;
    } catch (error) {
        console.error('样式清理错误:', error);
        return html;
    }
}

/**
 * 动态分析列表层级并添加对应的class
 * 
 * 策略:
 * 1. 扫描所有列表项，收集 margin-left 值和 mso-list level 值
 * 2. 优先使用 margin-left (视觉层级) 来确定嵌套关系
 *    - 如果存在不同的 margin-left 值，则根据 margin 大小排序映射为 level 1, 2, 3...
 * 3. 如果 margin-left 全部相同 (或缺失)，则回退到使用 mso-list level (语义层级)
 * 
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function addListLevelClasses(html) {
    if (!html) return '';

    try {
        const marginValues = new Set();
        const marginMap = new Map();
        const levelRegex = /mso-list:\s*l\d+\s+level(\d+)/i;
        const levels = new Set();

        // 辅助函数：解析长度为磅(pt)
        function parseLengthToPt(value, unit) {
            let val = parseFloat(value);
            if (isNaN(val)) return 0;
            unit = (unit || '').toLowerCase();
            if (unit === 'in') return val * 72;
            if (unit === 'cm') return val * 28.3465;
            if (unit === 'mm') return val * 2.83465;
            if (unit === 'pc') return val * 12;
            if (unit === 'px') return val * 0.75; 
            return val; 
        }

        // 第一遍扫描：收集所有 margin 和 level 信息
        html.replace(/<p[^>]*style=(["'])([\s\S]*?)\1[^>]*>/gi, (match, quote, style) => {
            if (style.includes('mso-list:')) {
                // 收集 level
                const levelMatch = style.match(levelRegex);
                if (levelMatch) {
                    levels.add(parseInt(levelMatch[1]));
                }

                // 收集 margin
                const marginMatch = style.match(/margin-left:\s*([0-9.]+)([a-z%]*)/i);
                if (marginMatch) {
                    const [, value, unit] = marginMatch;
                    marginValues.add(parseLengthToPt(value, unit));
                } else {
                    marginValues.add(0);
                }
            }
        });

        // 决策逻辑:
        // 如果有多个不同的 margin 值，说明有视觉上的缩进层级，优先使用 margin
        // 否则，如果有多个不同的 mso-list level，使用 level
        // 否则，默认为 level 1
        const useMargins = marginValues.size > 1;

        if (useMargins) {
            // 排序 margin 值并建立映射
            const sortedMargins = Array.from(marginValues).sort((a, b) => a - b);
            sortedMargins.forEach((margin, index) => {
                // 使用 toFixed(4) 避免浮点数比较问题
                marginMap.set(margin.toFixed(4), index + 1);
            });

            return html.replace(/<p([^>]*?)style=(["'])([\s\S]*?)\2([^>]*?)>/gi, (match, before, quote, style, after) => {
                if (style.includes('mso-list:')) {
                    let ptValue = 0;
                    const marginMatch = style.match(/margin-left:\s*([0-9.]+)([a-z%]*)/i);
                    if (marginMatch) {
                        const [, value, unit] = marginMatch;
                        ptValue = parseLengthToPt(value, unit);
                    }
                    
                    const level = marginMap.get(ptValue.toFixed(4)) || 1;
                    const levelClass = `list-level-${level}`;

                    // 检查是否已经有class属性
                    if (match.includes('class=')) {
                        if (/class=(["'])/.test(match)) {
                            return match.replace(/class=(["'])(.*?)\1/, `class=$1$2 ${levelClass}$1`);
                        } else {
                            return match.replace(/class=([^ >]+)/, `class="$1 ${levelClass}"`);
                        }
                    } else {
                        return `<p${before}style=${quote}${style}${quote} class="${levelClass}"${after}>`;
                    }
                }
                return match;
            });
        } else {
            // 回退到使用 mso-list level
            return html.replace(/<p([^>]*?)style=(["'])([\s\S]*?)\2([^>]*?)>/gi, (match, before, quote, style, after) => {
                if (style.includes('mso-list:')) {
                    const levelMatch = style.match(levelRegex);
                    if (levelMatch) {
                        const level = levelMatch[1];
                        const levelClass = `list-level-${level}`;

                        if (match.includes('class=')) {
                            if (/class=(["'])/.test(match)) {
                                return match.replace(/class=(["'])(.*?)\1/, `class=$1$2 ${levelClass}$1`);
                            } else {
                                return match.replace(/class=([^ >]+)/, `class="$1 ${levelClass}"`);
                            }
                        } else {
                            return `<p${before}style=${quote}${style}${quote} class="${levelClass}"${after}>`;
                        }
                    }
                }
                return match;
            });
        }
    } catch (error) {
        console.error('添加列表层级class错误:', error);
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
 * 移除 font 标签但保留内容
 */
function removeFontTags(html) {
    return html.replace(/<font[^>]*>([\s\S]*?)<\/font>/gi, '$1');
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
 * 转换 Word 的 mso-list 为标准的嵌套列表 (ul/ol)
 * 
 * 逻辑概述:
 * 1. 扫描所有带有 mso-list 样式的 p 标签
 * 2. 收集连续的列表项到 buffer
 * 3. 使用 bufferToList 函数将 buffer 转换为嵌套列表树
 *    - 根据 list-level-x class 判断层级
 *    - 根据 marker (1., a., •) 判断列表类型 (ol/ul)
 *    - 动态创建和嵌套 ul/ol/li 元素
 * 
 * @param {string} html 
 * @returns {string}
 */
function convertMsoListToNestedLists(html) {
    if (!html) return '';

    try {
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
        const document = dom.window.document;

        // 规范地处理document和section结构
        const documentDiv = document.querySelector('div.document');
        if (documentDiv) {
            const sectionDiv = documentDiv.querySelector('div.section');
            if (sectionDiv) {
                // 将section的内容移到body中
                while (sectionDiv.firstChild) {
                    documentDiv.parentNode.insertBefore(sectionDiv.firstChild, documentDiv);
                }
                // 移除空的document和section div
                documentDiv.remove();
            }
        }

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
            
            const isListPara = node.nodeType === 1 && 
                               node.tagName.toLowerCase() === 'p' && 
                               node.getAttribute('style') && 
                               node.getAttribute('style').includes('mso-list');

            if (isListPara) {
                buffer.push(node);
            } else if (node.nodeType === 3 && !node.textContent.trim()) {
                // 忽略列表项之间的空白文本节点，防止列表被打断
                if (buffer.length === 0) {
                    newNodes.push(node);
                }
                // 如果 buffer 不为空，则忽略该空白节点（视为列表项之间的间隔）
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
 * 清理特殊字符
 */
function cleanSpecialCharacters(html) {
    return html.replace(/&nbsp;/g, ' ');
}

/**
 * 移除空属性
 */
function cleanEmptyAttributes(html) {
    return html
        .replace(/(\w+)="\s*"/g, '')
        .replace(/\s+>/g, '>');
}

/**
 * 将 div 转换为 p
 */
function convertDivToP(html) {
    return html
        .replace(/<div[^>]*>/gi, '<p>')
        .replace(/<\/div>/gi, '</p>');
}

function cleanAllStyles(html) {
    if (!html) return '';

    try {
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
        const document = dom.window.document;

        // Find all elements with style attributes
        // Process in reverse order (bottom-up) to handle nested elements correctly
        const styledElements = Array.from(document.querySelectorAll('[style]')).reverse();
        const stylesWithContent = [];

        styledElements.forEach(element => {
            const style = element.getAttribute('style');
            
            // 转换样式为标签
            if (/font-weight:\s*bold/i.test(style)) {
                const b = document.createElement('b');
                b.innerHTML = element.innerHTML;
                element.innerHTML = '';
                element.appendChild(b);
            }
            if (/font-style:\s*italic/i.test(style)) {
                const i = document.createElement('i');
                i.innerHTML = element.innerHTML;
                element.innerHTML = '';
                element.appendChild(i);
            }
            if (/text-decoration:\s*underline/i.test(style)) {
                const u = document.createElement('u');
                u.innerHTML = element.innerHTML;
                element.innerHTML = '';
                element.appendChild(u);
            }

            stylesWithContent.push({
                tag: element.tagName.toLowerCase(),
                style: style,
                content: element.innerHTML,
                path: getElementPath(element)
            });

            // 保留表格相关的宽度样式，供后续 processTables 使用
            const isTableElement = ['table', 'col', 'colgroup', 'tr', 'td', 'th'].includes(element.tagName.toLowerCase());
            const hasWidth = /width:/i.test(style);

            if (!isTableElement || !hasWidth) {
                element.removeAttribute('style');
            }
        });

        // Store styles with their content in database
        if (stylesWithContent.length > 0) {
            // TODO: Implement database storage for styles and content
            console.log('Extracted style data to be stored:', stylesWithContent);
        }

        return document.body.innerHTML;
    } catch (error) {
        console.error('Style cleaning error:', error);
        return html;
    }
}

// Helper function to get unique element path
function getElementPath(element) {
    const path = [];
    while (element && element.parentElement && element.parentElement.tagName !== 'BODY') {
        let selector = element.tagName.toLowerCase();
        if (element.id) {
            selector += '#' + element.id;
        } else {
            // Get the index among siblings of same type
            let index = 1;
            let sibling = element.previousElementSibling;
            while (sibling) {
                if (sibling.tagName === element.tagName) {
                    index++;
                }
                sibling = sibling.previousElementSibling;
            }
            selector += `:nth-of-type(${index})`;
        }
        path.unshift(selector);
        element = element.parentElement;
    }
    return path.join(' > ');
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
 * 清理换行和装饰标签
 */
function cleanLineBreaksAndDecorations(html) {
    return html
        .replace(/<br>\s*<br>/gi, '<br>')
        .replace(/<br>$/i, '')
        .replace(/<s>([\s\S]*?)<\/s>/gi, '$1');
}

/**
 * 移除所有span标签但保留其内容
 * @param {string} html 
 * @returns {string}
 */
function removeSpans(html) {
    if (!html) return '';
    try {
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
        const document = dom.window.document;
        
        // Process in reverse order to handle nesting safely
        const spans = Array.from(document.querySelectorAll('span')).reverse();
        
        spans.forEach(span => {
            const parent = span.parentNode;
            if (parent) {
                while (span.firstChild) {
                    parent.insertBefore(span.firstChild, span);
                }
                parent.removeChild(span);
            }
        });
        
        return document.body.innerHTML;
    } catch (error) {
        console.error('Remove spans error:', error);
        return html;
    }
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
 * 处理jsdom无法处理的特殊情况
 * 这些转换通常涉及复杂的结构变化或自定义标签，JSDOM 可能会误处理
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function doExtraTransformJsdomCantHandle(html) {
    // 1. 处理表格: 将 HTML 表格转换为 DITA CALS 表格
    html = processTables(html);
    
    // 2. 处理链接: 将 <a> 转换为 <xref>
    html = html
        .replace(/<a[^>]*href="([^"]+)"[^>]*>/gi, '<xref scope="external" format="html" href="$1">')
        .replace(/<\/a>/gi, '</xref>');
        
    // 3. 处理标题: 将 <h1> 转换为 <title> 并包裹 <section>，h2-h6 降级为 <b>
    html = cleanHeadingTags(html);
    return html;
}

/**
 * 清理表格，移除所有属性并添加必要的属性
 * 将 HTML 表格转换为 DITA CALS 表格模型
 * 
 * 主要步骤:
 * 1. 计算列数 (考虑 colspan)
 * 2. 计算列宽 (从 colgroup 或第一行单元格提取，支持 px, pt, %, 等)
 * 3. 生成 dita-colspec 定义
 * 4. 生成 dita-thead 和 dita-tbody
 * 5. 处理单元格合并 (rowspan, colspan)
 * 
 * @param {string} html - 要处理的HTML
 * @returns {string} - 处理后的HTML
 */
function processTables(html) {
    if (!html) return '';

    try {
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
        const document = dom.window.document;

        // 1. 获取所有表格，并反转顺序（从最深层开始处理）
        // 这样可以确保嵌套表格在被父表格包含之前已经被转换为 dita-table 结构
        const tables = Array.from(document.querySelectorAll('table')).reverse();

        tables.forEach(table => {
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
                return; // 跳过无效表格
            }

            // 计算列宽
            let colWidths = new Array(colCount).fill(null);
            
            // 1. 尝试从 colgroup 获取宽度
            const colgroup = table.querySelector('colgroup');
            if (colgroup) {
                let colIndex = 0;
                const cols = colgroup.querySelectorAll('col');
                for (const col of cols) {
                    const span = parseInt(col.getAttribute('span') || '1', 10);
                    const width = col.style.width || col.getAttribute('width');
                    if (width) {
                        for (let k = 0; k < span; k++) {
                            if (colIndex + k < colCount) {
                                colWidths[colIndex + k] = width;
                            }
                        }
                    }
                    colIndex += span;
                }
            }

            // 2. 如果宽度缺失，尝试从第一行获取
            if (colWidths.some(w => !w)) {
                const firstRow = table.querySelector('tr');
                if (firstRow) {
                    let colIndex = 0;
                    const cells = firstRow.querySelectorAll('td, th');
                    for (const cell of cells) {
                        const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
                        const width = cell.style.width || cell.getAttribute('width');
                        
                        // 仅当 colspan=1 时应用宽度，避免复杂计算
                        if (colspan === 1 && width) {
                            if (colIndex < colCount && !colWidths[colIndex]) {
                                colWidths[colIndex] = width;
                            }
                        }
                        colIndex += colspan;
                    }
                }
            }

            // 规范化宽度格式
            colWidths = colWidths.map(w => {
                if (!w) return '1*'; // 默认比例
                
                // 处理百分比
                if (w.includes('%')) {
                    return w.replace('%', '*') // 20% -> 20*
                }

                // 处理固定宽度，转换为无单位的整数（像素值），以保持比例
                try {
                    const match = w.match(/^([\d.]+)([a-z]*)$/i);
                    if (match) {
                        let val = parseFloat(match[1]);
                        const unit = match[2].toLowerCase();
                        
                        // 简单的单位转换 (以px为基准)
                        if (unit === 'pt') val *= 1.3333;
                        else if (unit === 'in') val *= 96;
                        else if (unit === 'cm') val *= 37.795;
                        else if (unit === 'mm') val *= 3.7795;
                        else if (unit === 'pc') val *= 16;
                        
                        // 返回整数
                        return Math.round(val).toString();
                    }
                } catch (e) {
                    // 忽略解析错误
                }

                // 如果无法解析，保留原样或默认处理
                if (w.match(/^\d+$/)) {
                    return w; // 纯数字直接返回
                }
                return w;
            });

            // 构建 dita-table 结构
            let newTable = '<dita-table frame="all" rowsep="1" colsep="1">';
            newTable += `<dita-tgroup cols="${colCount}">`;
            
            for (let i = 1; i <= colCount; i++) {
                const widthAttr = colWidths[i-1] ? ` colwidth="${colWidths[i-1]}"` : '';
                // 使用非自闭合标签，避免 JSDOM 解析问题
                newTable += `<dita-colspec colnum="${i}" colname="col${i}"${widthAttr}></dita-colspec>`;
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
                newTable += '<dita-thead>';
                newTable += generateCalsRows(headRows, colCount);
                newTable += '</dita-thead>';
            }

            // 生成tbody
            if (bodyRows.length) {
                newTable += '<dita-tbody>';
                newTable += generateCalsRows(bodyRows, colCount);
                newTable += '</dita-tbody>';
            }

            newTable += '</dita-tgroup></dita-table>';

            // 使用 outerHTML 替换原表格
            // JSDOM 允许设置 outerHTML 为任意 HTML 字符串，包括自定义标签
            table.outerHTML = newTable;
        });

        let finalHtml = document.body.innerHTML;
        // 保持 dita-xxx 标签，以便后续 JSDOM 处理（如 mergeConsecutiveInlineTags）不会破坏结构
        // 最终在 cleanHtml 结束时统一替换回标准标签

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
        result += '<dita-row>';
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
            result += `<dita-entry${entryAttrs}>${cellContent}</dita-entry>`;
            col += colspan;
        }
        // 补齐空格
        while (col < colCount) {
            result += '<dita-entry></dita-entry>';
            col++;
        }
        result += '</dita-row>';
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
            // 替换第一个h1为<dita-title> (使用自定义标签避免JSDOM将title内容解析为纯文本)
            html = html.replace(/<h1[^>]*>[\s\S]*?<\/h1>/i, `<dita-title>${h1Content}</dita-title>`);
            
            // 将剩余的h1转换为b (避免多重标题)
            html = html.replace(/<h1[^>]*>([\s\S]*?)<\/h1>/gi, '<b>$1</b>');

            // 用<section>包裹整体内容
            html = `<section>${html}</section>`;
        }
        return html;
    } catch (error) {
        console.error('处理标题标签错误:', error);
        return html;
    }
}

/**
 * 合并连续的内联标签 (b, i, u)
 * 
 * 解决 Word 转换中常见的碎片化标签问题，例如:
 * <b>H</b><b>ello</b> -> <b>Hello</b>
 * 
 * 处理逻辑:
 * 1. 递归遍历 DOM 树
 * 2. 处理嵌套的相同标签 (unwrap inner)
 * 3. 处理相邻的相同标签 (merge)
 *    - 忽略标签之间的纯空白文本节点
 * 
 * @param {string} html 
 * @returns {string}
 */
function mergeConsecutiveInlineTags(html) {
    if (!html) return '';
    try {
        const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
        const document = dom.window.document;
        const body = document.body;
        const mergeTags = ['b', 'i', 'u'];

        function traverse(node) {
            if (node.nodeType !== 1) return; // Element node

            // 1. Handle nested identical tags (unwrap inner)
            if (mergeTags.includes(node.tagName.toLowerCase())) {
                 for (let i = 0; i < node.childNodes.length; i++) {
                    const child = node.childNodes[i];
                    if (child.nodeType === 1 && child.tagName === node.tagName) {
                        // Unwrap child: move all children of 'child' before 'child', then remove 'child'
                        while (child.firstChild) {
                            node.insertBefore(child.firstChild, child);
                        }
                        node.removeChild(child);
                        i--; // Re-check the new nodes at this position
                    }
                 }
            }

            // 2. Handle consecutive sibling tags
            for (let i = 0; i < node.childNodes.length; i++) {
                const current = node.childNodes[i];
                
                // Check if current is one of the target tags
                if (current.nodeType === 1 && mergeTags.includes(current.tagName.toLowerCase())) {
                    let nextIndex = i + 1;
                    if (nextIndex >= node.childNodes.length) break;
                    
                    let next = node.childNodes[nextIndex];
                    let whitespaceNode = null;

                    // Check if next is whitespace text node
                    if (next.nodeType === 3 && /^\s+$/.test(next.textContent)) {
                        whitespaceNode = next;
                        nextIndex++;
                        if (nextIndex >= node.childNodes.length) break;
                        next = node.childNodes[nextIndex];
                    }

                    // Check if the next element is the same tag
                    if (next.nodeType === 1 && next.tagName === current.tagName) {
                        // Merge
                        if (whitespaceNode) {
                            current.innerHTML += whitespaceNode.textContent;
                            node.removeChild(whitespaceNode);
                        }
                        current.innerHTML += next.innerHTML;
                        node.removeChild(next);
                        
                        // Stay on current index to check for more consecutive tags
                        i--; 
                    }
                }
            }
            
            // 3. Recursive for children
            for (let child of node.childNodes) {
                traverse(child);
            }
        }

        traverse(body);
        return body.innerHTML;
    } catch (error) {
        console.error('合并内联标签错误:', error);
        return html;
    }
}

/**
 * 还原 DITA 标签
 */
function restoreDitaTags(html) {
    return html
        .replace(/<dita-title>/g, '<title>').replace(/<\/dita-title>/g, '</title>')
        .replace(/<dita-table/g, '<table').replace(/<\/dita-table>/g, '</table>')
        .replace(/<dita-tgroup/g, '<tgroup').replace(/<\/dita-tgroup>/g, '</tgroup>')
        .replace(/<dita-thead/g, '<thead').replace(/<\/dita-thead>/g, '</thead>')
        .replace(/<dita-tbody/g, '<tbody').replace(/<\/dita-tbody>/g, '</tbody>')
        .replace(/<dita-row/g, '<row').replace(/<\/dita-row>/g, '</row>')
        .replace(/<dita-entry/g, '<entry').replace(/<\/dita-entry>/g, '</entry>')
        .replace(/<dita-colspec([^>]*)><\/dita-colspec>/g, '<colspec$1/>');
}

/**
 * 格式化HTML（使其更易读）
 * 使用自定义的缩进逻辑，避免破坏 DITA 的 XML 结构
 * 特别针对 'simple blocks' (只包含内联内容的块级元素) 进行了优化，使其保持在单行
 * @param {string} html - 要格式化的HTML
 * @returns {string} - 格式化后的HTML
 */
function formatHtml(html) {
    if (!html) return '';

    try {
        let formatted = '';
        let indent = 0;
        // Split by tags, capturing delimiters
        const tags = html.split(/(<\/?[^>]+>)/g);
        
        // Define block-level tags that trigger newlines
        const blockTags = new Set([
            'html', 'body', 'dita', 'topic', 'title', 'shortdesc', 'body', 'section', 
            'p', 'div', 'table', 'tgroup', 'thead', 'tbody', 'row', 'entry', 'colspec', 
            'ul', 'ol', 'li', 'dl', 'dlentry', 'dt', 'dd', 'fig', 'note', 'lines', 'pre',
            'dita-table', 'dita-tgroup', 'dita-thead', 'dita-tbody', 'dita-row', 'dita-entry', 'dita-colspec'
        ]);
        
        let currentLine = '';
        
        function flushLine() {
            if (currentLine.trim()) {
                formatted += '  '.repeat(indent) + currentLine.trim() + '\n';
            }
            currentLine = '';
        }

        // Helper to check if a block element contains only inline content
        function isSimpleBlock(startIndex, tagName) {
            let balance = 1;
            for (let j = startIndex + 1; j < tags.length; j++) {
                const nextTag = tags[j];
                if (!nextTag) continue;
                
                if (nextTag.startsWith('<')) {
                    const isNextClose = nextTag.startsWith('</');
                    const isNextSelfClosing = nextTag.endsWith('/>');
                    const nextTagNameMatch = nextTag.match(/^<\/?([^\s\/>]+)/);
                    const nextTagName = nextTagNameMatch ? nextTagNameMatch[1].toLowerCase() : '';

                    if (nextTagName === tagName) {
                        if (isNextClose) {
                            balance--;
                            if (balance === 0) return true;
                        } else if (!isNextSelfClosing) {
                            balance++;
                        }
                    } else if (blockTags.has(nextTagName)) {
                        return false; // Contains another block tag
                    }
                }
            }
            return false;
        }

        for (let i = 0; i < tags.length; i++) {
            let tag = tags[i];
            if (!tag) continue;

            const isTag = tag.startsWith('<');
            
            if (isTag) {
                const isClose = tag.startsWith('</');
                const isSelfClosing = tag.endsWith('/>');
                // Fix regex to exclude trailing slash in self-closing tags from tag name
                const tagNameMatch = tag.match(/^<\/?([^\s\/>]+)/);
                const tagName = tagNameMatch ? tagNameMatch[1].toLowerCase() : '';

                if (blockTags.has(tagName)) {
                    // Block tag
                    if (isClose) {
                        flushLine();
                        indent = Math.max(0, indent - 1);
                        formatted += '  '.repeat(indent) + tag + '\n';
                    } else if (isSelfClosing) {
                        flushLine();
                        formatted += '  '.repeat(indent) + tag + '\n';
                    } else {
                        // Open block tag
                        // Check if it's a simple block (only inline content)
                        // We specifically target entry, p, li, title, dt, dd for this optimization
                        if (['entry', 'dita-entry', 'p', 'li', 'title', 'dt', 'dd'].includes(tagName) && isSimpleBlock(i, tagName)) {
                            flushLine();
                            formatted += '  '.repeat(indent) + tag;
                            
                            // Consume until close tag
                            let balance = 1;
                            i++;
                            while (i < tags.length) {
                                const nextTag = tags[i];
                                if (nextTag) {
                                    if (nextTag.startsWith('<')) {
                                        const isNextClose = nextTag.startsWith('</');
                                        const isNextSelfClosing = nextTag.endsWith('/>');
                                        const nextTagNameMatch = nextTag.match(/^<\/?([^\s\/>]+)/);
                                        const nextTagName = nextTagNameMatch ? nextTagNameMatch[1].toLowerCase() : '';
                                        
                                        if (nextTagName === tagName) {
                                            if (isNextClose) {
                                                balance--;
                                                if (balance === 0) {
                                                    formatted += nextTag + '\n';
                                                    break;
                                                }
                                            } else if (!isNextSelfClosing) {
                                                balance++;
                                            }
                                        }
                                    }
                                    // Append content directly
                                    // Collapse multiple spaces to single space for text nodes
                                    const content = nextTag.startsWith('<') ? nextTag : nextTag.replace(/\s+/g, ' ');
                                    formatted += content;
                                }
                                i++;
                            }
                        } else {
                            flushLine();
                            formatted += '  '.repeat(indent) + tag + '\n';
                            indent++;
                        }
                    }
                } else {
                    // Inline tag
                    currentLine += tag;
                }
            } else {
                // Text node
                // Collapse multiple spaces to single space
                const text = tag.replace(/\s+/g, ' ');
                currentLine += text;
            }
        }
        flushLine();
        
        return formatted.trim();
    } catch (error) {
        console.error('HTML格式化错误:', error);
        return html;
    }
}

module.exports = {
    cleanHtml,
    formatHtml,
};
