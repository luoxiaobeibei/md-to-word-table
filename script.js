document.addEventListener('DOMContentLoaded', () => {
    const markdownInput = document.getElementById('markdownInput');
    const convertButton = document.getElementById('convertButton');
    const clearButton = document.getElementById('clearButton');
    const htmlPreview = document.getElementById('htmlPreview');
    const statusMessage = document.getElementById('statusMessage');
    const langZhBtn = document.getElementById('lang-zh');
    const langEnBtn = document.getElementById('lang-en');

    // 国际化翻译
    const translations = {
        'zh': {
            'page-title': 'Markdown表格转Word',
            'main-title': 'Markdown表格转Word转换器',
            'main-subtitle': '轻松将Markdown表格转换为Word兼容的格式',
            'usage-guide': '使用说明',
            'guide-step1': '将Markdown格式的表格粘贴到下方文本框',
            'guide-step2': '点击"转换并复制"按钮',
            'guide-step3': '在Word文档中直接粘贴(Ctrl+V)即可得到格式化的表格',
            'input-title': '粘贴您的Markdown表格：',
            'convert-button': '转换并复制到剪贴板',
            'clear-button': '清空',
            'preview-title': '预览 (HTML表格)',
            'status-success': '表格HTML已复制到剪贴板！',
            'status-empty': '错误：输入为空。',
            'status-no-table': '警告：Markdown输入中未找到表格。',
            'status-error': '错误：',
            'status-api-error': '错误：剪贴板API不受支持或需要HTTPS。',
            'status-copy-error': '错误：无法复制到剪贴板。检查浏览器权限。',
            'status-fallback': '表格已复制（使用备用方法）！粘贴格式可能会有所不同。',
            'status-copy-failed': '错误：复制完全失败。'
        },
        'en': {
            'page-title': 'Markdown Table to Word',
            'main-title': 'Markdown Table to Word Converter',
            'main-subtitle': 'Easily convert Markdown tables to Word-compatible format',
            'usage-guide': 'Usage Guide',
            'guide-step1': 'Paste your Markdown table in the text area below',
            'guide-step2': 'Click the "Convert & Copy" button',
            'guide-step3': 'Paste (Ctrl+V) directly into your Word document to get the formatted table',
            'input-title': 'Paste your Markdown table:',
            'convert-button': 'Convert & Copy to Clipboard',
            'clear-button': 'Clear',
            'preview-title': 'Preview (HTML Table)',
            'status-success': 'Table HTML (with borders) copied to clipboard!',
            'status-empty': 'Error: Input is empty.',
            'status-no-table': 'Warning: No table found in the Markdown input.',
            'status-error': 'Error: ',
            'status-api-error': 'Error: Clipboard API not supported or requires HTTPS.',
            'status-copy-error': 'Error: Could not copy to clipboard. Check browser permissions.',
            'status-fallback': 'Table copied (using fallback)! Paste format might vary.',
            'status-copy-failed': 'Error: Copying failed completely.'
        }
    };

    // 设置默认语言
    let currentLang = 'zh';

    // 应用翻译
    function applyTranslations(lang) {
        const elements = document.querySelectorAll('[data-i18n]');
        elements.forEach(el => {
            const key = el.getAttribute('data-i18n');
            if (translations[lang] && translations[lang][key]) {
                if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
                    el.placeholder = translations[lang][key];
                } else {
                    el.textContent = translations[lang][key];
                }
            }
        });
        // 更新文档标题
        document.title = translations[lang]['page-title'];
        // 更新语言按钮状态
        langZhBtn.classList.toggle('active', lang === 'zh');
        langEnBtn.classList.toggle('active', lang === 'en');
        // 更新HTML属性
        document.documentElement.lang = lang === 'zh' ? 'zh-CN' : 'en';
    }

    // 语言切换事件
    langZhBtn.addEventListener('click', () => {
        currentLang = 'zh';
        applyTranslations(currentLang);
    });

    langEnBtn.addEventListener('click', () => {
        currentLang = 'en';
        applyTranslations(currentLang);
    });

    // Configure marked.js options if needed (optional)
    // marked.setOptions({ ... });

    convertButton.addEventListener('click', () => {
        const markdownText = markdownInput.value.trim();
        statusMessage.textContent = ''; // Clear previous status
        statusMessage.className = 'alert mt-3'; // Reset status style
        statusMessage.classList.add('d-none');
        htmlPreview.innerHTML = ''; // Clear previous preview

        if (!markdownText) {
            showStatus(translations[currentLang]['status-empty'], 'danger');
            return;
        }

        try {
            // 1. Parse Markdown to HTML
            const generatedHtml = marked.parse(markdownText);

            // 2. Temporarily add to DOM to manipulate styles
            // Create a temporary div to hold the parsed HTML
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = generatedHtml;

            // 3. Find the table element
            const tableElement = tempDiv.querySelector('table');

            if (!tableElement) {
                showStatus(translations[currentLang]['status-no-table'], 'warning');
                htmlPreview.innerHTML = generatedHtml; // Show preview even if no table
                return;
            }

            // --- START: Add Inline Styles ---
            // Style the table itself
            tableElement.style.border = '1px solid black';
            tableElement.style.borderCollapse = 'collapse'; // Crucial for Word table look
            tableElement.style.width = 'auto'; // Or '100%' if you prefer full width

            // Style table headers (th) and cells (td)
            const cells = tableElement.querySelectorAll('th, td');
            cells.forEach(cell => {
                cell.style.border = '1px solid black';
                cell.style.padding = '5px'; // Add some padding for better look in Word
                cell.style.textAlign = 'left'; // Default alignment
            });
            // --- END: Add Inline Styles ---

            // 4. Get the HTML string *with* inline styles
            const styledTableHtml = tableElement.outerHTML;

            // 5. Display preview (optional: show the styled version)
            // htmlPreview.innerHTML = styledTableHtml; // Preview will now show inline styles too
            htmlPreview.appendChild(tableElement); // Append the modified table element to the preview area

            // 6. Copy the styled HTML to Clipboard
            copyHtmlToClipboard(styledTableHtml);

        } catch (error) {
            console.error('Error during conversion or copying:', error);
            showStatus(`${translations[currentLang]['status-error']}${error.message}`, 'danger');
            // Display original parsed HTML in preview even on error after parsing
            if (typeof generatedHtml !== 'undefined') {
                htmlPreview.innerHTML = generatedHtml;
            }
        }
    });

    // 添加清空按钮功能
    clearButton.addEventListener('click', () => {
        markdownInput.value = '';
        htmlPreview.innerHTML = '';
        statusMessage.classList.add('d-none');
    });

    function showStatus(message, type) {
        statusMessage.textContent = message;
        statusMessage.className = `alert mt-3 alert-${type}`;
        statusMessage.classList.remove('d-none');
    }

    function copyHtmlToClipboard(htmlString) {
        // Use the modern Clipboard API
        if (!navigator.clipboard || !navigator.clipboard.write) {
            showStatus(translations[currentLang]['status-api-error'], 'danger');
            console.warn('Clipboard API not available.');
            fallbackCopy(htmlString); // Attempt fallback
            return;
        }

        const blobHtml = new Blob([htmlString], { type: 'text/html' });
        const blobText = new Blob([markdownInput.value], { type: 'text/plain' });

        const clipboardItem = new ClipboardItem({
            'text/html': blobHtml,
            'text/plain': blobText
        });

        navigator.clipboard.write([clipboardItem]).then(() => {
            showStatus(translations[currentLang]['status-success'], 'success');
        }).catch(err => {
            console.error('Failed to copy HTML to clipboard:', err);
            showStatus(translations[currentLang]['status-copy-error'], 'danger');
            fallbackCopy(htmlString); // Attempt fallback on error
        });
    }

    // Fallback using the deprecated execCommand (less reliable for HTML)
    function fallbackCopy(htmlString) {
        console.warn("Attempting fallback copy method (execCommand).");
        const tempElement = document.createElement('div');
        tempElement.contentEditable = true;
        tempElement.style.position = 'absolute';
        tempElement.style.left = '-9999px';
        tempElement.innerHTML = htmlString; // Set the HTML content

        document.body.appendChild(tempElement);

        const range = document.createRange();
        range.selectNodeContents(tempElement);

        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);

        try {
            const successful = document.execCommand('copy');
            if (successful) {
                showStatus(translations[currentLang]['status-fallback'], 'warning');
            } else {
                throw new Error('execCommand failed');
            }
        } catch (err) {
            console.error('Fallback copy failed:', err);
            showStatus(translations[currentLang]['status-copy-failed'], 'danger');
        } finally {
            document.body.removeChild(tempElement);
            selection.removeAllRanges(); // Clear selection
        }
    }

    // Initialize application translations
    applyTranslations(currentLang);
});