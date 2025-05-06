document.addEventListener('DOMContentLoaded', () => {
    const pasteArea = document.getElementById('pasteArea');
    const outputArea = document.getElementById('outputArea');
    const clearBtn = document.getElementById('clearBtn');
    const transformBtn = document.getElementById('transformBtn');
    const copyBtn = document.getElementById('copyBtn');
    const errorAlert = document.getElementById('errorAlert');
    const successAlert = document.getElementById('successAlert');

    // 显示当前页面URL
    console.log('当前页面URL:', window.location.href);
    console.log('API请求将发送到:', new URL('/api/transform', window.location.href).href);

    // 显示提示信息
    function showAlert(type, message, duration = 3000) {
        const alert = type === 'error' ? errorAlert : successAlert;
        alert.textContent = message;
        alert.style.display = 'block';
        
        setTimeout(() => {
            alert.style.display = 'none';
        }, duration);
    }

    // 处理粘贴事件
    pasteArea.addEventListener('paste', (e) => {
        e.preventDefault();
        
        // 获取粘贴的内容
        const clipboardData = e.clipboardData || window.clipboardData;
        let pastedData = '';

        // 优先获取HTML格式
        if (clipboardData.types.includes('text/html')) {
            pastedData = clipboardData.getData('text/html');
        } else {
            // 如果没有HTML格式，获取纯文本并转换为HTML
            pastedData = clipboardData.getData('text/plain');
            // 转换换行符为<br>
            pastedData = pastedData.replace(/\n/g, '<br>');
        }
        
        // 将内容显示在粘贴区域，保持HTML格式
        pasteArea.innerHTML = pastedData;
    });

    // 清空按钮事件
    clearBtn.addEventListener('click', () => {
        pasteArea.innerHTML = '';
        outputArea.textContent = '';
        showAlert('success', '内容已清空');
    });

    // 转换按钮事件
    transformBtn.addEventListener('click', async () => {
        // 获取粘贴区域的HTML内容
        const content = pasteArea.innerHTML;
        
        if (!content.trim()) {
            showAlert('error', '请先粘贴内容！');
            return;
        }

        try {
            transformBtn.disabled = true;
            transformBtn.classList.add('loading');
            transformBtn.textContent = '转换中...';

            const response = await fetch('/api/transform', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ content })
            });

            if (!response.ok) {
                throw new Error(`请求失败: ${response.status}`);
            }

            const result = await response.json();
            
            // 在右侧显示原始HTML文本
            outputArea.textContent = result.html;
            showAlert('success', '转换成功！');
            
        } catch (error) {
            console.error('请求错误:', error);
            showAlert('error', '转换失败，请重试');
        } finally {
            transformBtn.disabled = false;
            transformBtn.classList.remove('loading');
            transformBtn.textContent = '转换';
        }
    });

    // 复制结果按钮事件
    copyBtn.addEventListener('click', () => {
        const content = outputArea.textContent;
        if (!content.trim()) {
            showAlert('error', '没有可复制的内容！');
            return;
        }

        try {
            // 直接复制文本内容
            navigator.clipboard.writeText(content).then(() => {
                showAlert('success', '复制成功！');
            });
        } catch (err) {
            console.error('复制失败:', err);
            showAlert('error', '复制失败，请手动复制');
        }
    });

    // 处理粘贴区域的placeholder
    pasteArea.addEventListener('focus', () => {
        if (!pasteArea.innerHTML.trim()) {
            pasteArea.innerHTML = '';
        }
    });

    pasteArea.addEventListener('blur', () => {
        if (!pasteArea.innerHTML.trim()) {
            pasteArea.innerHTML = '';
        }
    });

    // 初始化时清空内容
    pasteArea.innerHTML = '';
    console.log('页面初始化完成');
}); 