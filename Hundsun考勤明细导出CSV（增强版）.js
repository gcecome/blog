// ==UserScript==
// @name         考勤明细导出CSV（增强版）
// @namespace    http://tampermonkey.net/
// @version      0.4
// @description  支持延迟加载的PeopleSoft考勤表格导出
// @author       GGB
// @match https://hr.hundsun.com/psp/ps_*/EMPLOYEE/HRMS/c/DC_ABS_MENU.GPS_ABS_DAY_ESS.GBL*
// @grant        none
// @run-at       document-end
// ==/UserScript==

// 油猴脚本 导出csv，excel工具转换成xlsx后，=E2-D2 算时间差 然后 =AVERAGEIF(K2:K103,">0") 公式算去除0的平均工时 恒生电子考勤工时导出

(function () {
    'use strict';

    let win = "win1divGPS_ABS_DAY_VW2$0"; // 默认值

    // 提取 ps_X 中的 X（1, 2, 3, 4, 5...）
    const match = window.location.href.match(/\/psp\/ps_(\d+)\//);
    if (match && match[1]) {
        const psVersion = match[1];
        console.log(`当前匹配到 ps_${psVersion} 版本`);
        win = `win${psVersion}divGPS_ABS_DAY_VW2$0`;
    } else {
        console.log("未匹配到 ps_X 版本，使用默认 win1div...");
    }
    // 添加样式
    const style = document.createElement('style');
    style.innerHTML = `
        #floatingExportBtn {
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            padding: 12px 24px;
            background-color: #4CAF50;
            color: white;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            z-index: 99999;
            font-size: 14px;
            user-select: none;
            transition: all 0.2s;
        }
        #floatingExportBtn:hover {
            background-color: #45a049;
            transform: translate(-50%, -50%) scale(1.05);
        }
        #floatingExportBtn:disabled {
            background-color: #ccc;
            cursor: not-allowed;
            transform: translate(-50%, -50%);
        }
    `;
    document.head.appendChild(style);

    // 创建按钮
    const exportBtn = document.createElement('button');
    exportBtn.id = 'floatingExportBtn';
    exportBtn.textContent = '📥 导出考勤数据';
    exportBtn.disabled = true; // 初始禁用
    document.body.appendChild(exportBtn);

    // 拖动逻辑（不变）
    let isDragging = false, offsetX, offsetY;
    exportBtn.addEventListener('mousedown', e => {
        isDragging = true;
        offsetX = e.clientX - exportBtn.getBoundingClientRect().left;
        offsetY = e.clientY - exportBtn.getBoundingClientRect().top;
        exportBtn.style.cursor = 'grabbing';
        e.preventDefault();
    });
    document.addEventListener('mousemove', e => {
        if (!isDragging) return;
        exportBtn.style.left = `${e.clientX - offsetX}px`;
        exportBtn.style.top = `${e.clientY - offsetY}px`;
        exportBtn.style.transform = 'none';
    });
    document.addEventListener('mouseup', () => {
        isDragging = false;
        exportBtn.style.cursor = 'pointer';
    });

    // ========== 核心：等待全局变量出现 ==========
    function waitForGlobalVariable(varName, interval = 500, timeout = 30000) {
        return new Promise((resolve, reject) => {
            const start = Date.now();

            const check = () => {
                const obj = window[0][varName];
                console.log(`window[0]['${varName}']=`, obj);
                if (obj) {
                    console.log(`✅ 全局变量 ${varName} 已就绪`);
                    resolve(obj);
                    return;
                }

                if (Date.now() - start > timeout) {
                    reject(new Error(`${varName} 加载超时（>${timeout}ms）`));
                    return;
                }

                // 继续轮询
                setTimeout(check, interval);
            };

            check();
        });
    }

    // ========== 导出逻辑 ==========
    async function exportToCSV() {
        exportBtn.disabled = true;
        exportBtn.textContent = '🔍 检查数据...';

        try {
            // 等待变量出现
            const container = await waitForGlobalVariable(win, 500, 60000); // 最长等60秒

            if (!container || !container.querySelector) {
                throw new Error('容器无效，缺少 querySelector 方法');
            }

            const table = container.querySelector('table.PSLEVEL1GRID');
            if (!table) {
                throw new Error('未找到表格，请确认已加载考勤明细');
            }

            // 提取表头
            const headers = Array.from(table.querySelectorAll('th.PSLEVEL1GRIDCOLUMNHDR'))
                .map(th => (th.textContent || th.getAttribute('abbr') || '').trim())
                .filter(h => h); // 过滤空标题

            if (headers.length === 0) {
                throw new Error('未提取到表头，请检查表格结构');
            }

            // 提取数据行
            const rows = Array.from(table.querySelectorAll('tr[id^="trGPS_ABS_DAY_VW2"]'))
                .map(tr => {
                    return Array.from(tr.querySelectorAll('td:not(.PSGRIDFIRSTCOLUMN)'))
                        .map(td => {
                            const span = td.querySelector('span');
                            return span ? span.textContent.trim() : '';
                        });
                })
                .filter(row => row.length > 0 && row.some(cell => cell !== ''));

            if (rows.length === 0) {
                throw new Error('未提取到任何考勤记录，请确认表格中有数据');
            }

            // 构建 CSV
            let csvContent = headers.join(',') + '\n';
            rows.forEach(row => {
                csvContent += row
                    .map(cell => `"${cell.replace(/"/g, '""')}"`)
                    .join(',') + '\n';
            });

            // 下载
            const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            const dateStr = new Date().toLocaleDateString().replace(/\//g, '');
            link.download = `考勤明细_${dateStr}.csv`;
            link.style.display = 'none';
            document.body.appendChild(link);
            link.click();

            setTimeout(() => {
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
            }, 100);

            // 成功反馈
            exportBtn.textContent = '✅ 导出成功！';
            setTimeout(() => {
                exportBtn.textContent = '📥 导出考勤数据';
                exportBtn.disabled = false;
            }, 2000);

        } catch (error) {
            console.error('导出失败:', error);
            alert(`❌ 导出失败：${error.message}\n\n请尝试刷新页面或稍后重试。`);
            exportBtn.textContent = '⚠️ 失败，点击重试';
            exportBtn.disabled = false;
        }
    }

    // ========== 初始化按钮 ==========
    exportBtn.disabled = false;
    exportBtn.textContent = '📥 导出考勤数据';

    exportBtn.addEventListener('click', () => {
        if (exportBtn.disabled) return;
        exportToCSV();
    });

    console.log('[考勤导出脚本 v0.4] 已加载，等待页面数据...', window);
})();
