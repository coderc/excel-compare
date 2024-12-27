let workbook1 = null;
let workbook2 = null;

// 当前差异位置
let currentDiffIndex = -1;

// 添加更新文件名的函数
function updateFileName(input, spanId) {
    const fileName = input.files[0]?.name || '未选择文件';
    document.getElementById(spanId).textContent = fileName;
}

// 监听文件上传
document.getElementById('file1').addEventListener('change', async (e) => {
    workbook1 = await loadFile(e.target.files[0]);
});

document.getElementById('file2').addEventListener('change', async (e) => {
    workbook2 = await loadFile(e.target.files[0]);
});

// 加载Excel文件
async function loadFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            resolve(workbook);
        };
        reader.readAsArrayBuffer(file);
    });
}

// 比较文件
function compareFiles() {
    if (!workbook1 || !workbook2) {
        alert('请先选择两个Excel文件');
        return;
    }

    // 获取两个文件中的所有sheet名称
    const sheets1 = new Set(workbook1.SheetNames);
    const sheets2 = new Set(workbook2.SheetNames);

    // 获取所有sheet名称（包括不匹配的）
    const allSheets = [...new Set([...workbook1.SheetNames, ...workbook2.SheetNames])];

    if (allSheets.length === 0) {
        alert('文件中没有工作表');
        return;
    }

    // 创建标签页
    createTabs(allSheets);
    
    // 处理所有sheet
    allSheets.forEach((sheetName, index) => {
        const sheet1 = workbook1.Sheets[sheetName] || null;
        const sheet2 = workbook2.Sheets[sheetName] || null;
        
        if (sheet1 && sheet2) {
            // 如果两个文件都有这个sheet，进行对比
            const data1 = XLSX.utils.sheet_to_json(sheet1, {header: 1});
            const data2 = XLSX.utils.sheet_to_json(sheet2, {header: 1});
            displayDifferences(data1, data2, sheetName);
        } else {
            // 如果只有一个文件有这个sheet，只显示内容
            const sheet = sheet1 || sheet2;
            const data = XLSX.utils.sheet_to_json(sheet, {header: 1});
            displaySingleSheet(data, sheetName, sheet1 ? 'source' : 'incoming');
        }
    });

    // 默认显示第一个标签页
    document.querySelector('.tab-button').click();
}

// 创建标签页
function createTabs(sheetNames) {
    const tabsContainer = document.getElementById('tabs');
    const tabContents = document.getElementById('tab-contents');
    
    tabsContainer.innerHTML = sheetNames.map((name, index) => `
        <button class="tab-button" onclick="switchTab('${name}')">${name}</button>
    `).join('');

    tabContents.innerHTML = sheetNames.map(name => `
        <div id="content-${name}" class="tab-content"></div>
    `).join('');
}

// 切换标签页
function switchTab(sheetName) {
    currentDiffIndex = -1;
    // 更新标签按钮状态
    document.querySelectorAll('.tab-button').forEach(button => {
        button.classList.remove('active');
        if (button.textContent === sheetName) {
            button.classList.add('active');
        }
    });

    // 更新内容显示
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    document.getElementById(`content-${sheetName}`).classList.add('active');
}

// 显示差异
function displayDifferences(data1, data2, sheetName) {
    const contentDiv = document.getElementById(`content-${sheetName}`);
    let hasDifferences = false;
    
    // 计算最大列数
    const maxCols = Math.max(
        ...data1.map(row => row.length),
        ...data2.map(row => row.length)
    );
    
    // 生成表头
    let html = '<table class="diff-table"><thead><tr><th>行号</th>';
    // Source文件列
    for (let i = 0; i < maxCols; i++) {
        html += `<th>Source列${i + 1}</th>`;
    }
    // 添加分隔列
    html += '<th class="separator"></th>';
    // Incoming文件列
    for (let i = 0; i < maxCols; i++) {
        html += `<th>Incoming列${i + 1}</th>`;
    }
    html += '<th>选择</th></tr></thead><tbody>';

    // 创建行内容的哈希映射，用于快速查找匹配行
    const sourceMap = new Map();
    data1.forEach((row, index) => {
        const content = row.join('|'); // 使用分隔符连接行内容作为key
        if (!sourceMap.has(content)) {
            sourceMap.set(content, []);
        }
        sourceMap.get(content).push(index);
    });

    // 标记已匹配的行
    const matchedSource = new Set();
    const matchedIncoming = new Set();
    const matches = []; // 存储匹配的行对

    // 找出匹配的行
    data2.forEach((incomingRow, incomingIndex) => {
        // 在源文件中查找匹配的行
        let bestMatch = null;
        let bestMatchIndex = -1;
        let maxMatchingCols = 0;

        data1.forEach((sourceRow, sourceIndex) => {
            if (!matchedSource.has(sourceIndex)) {
                // 计算匹配的列数
                let matchingCols = 0;
                const minCols = Math.min(sourceRow.length, incomingRow.length);
                
                for (let i = 0; i < minCols; i++) {
                    if (sourceRow[i] === incomingRow[i]) {
                        matchingCols++;
                    }
                }

                // 更新最佳匹配
                if (matchingCols > maxMatchingCols) {
                    maxMatchingCols = matchingCols;
                    bestMatch = sourceRow;
                    bestMatchIndex = sourceIndex;
                }
            }
        });

        // 如果找到足够好的匹配（可以设置阈值）
        if (maxMatchingCols > 0) {
            matches.push({
                sourceIndex: bestMatchIndex,
                incomingIndex: incomingIndex,
                content: incomingRow
            });
            matchedSource.add(bestMatchIndex);
            matchedIncoming.add(incomingIndex);
        }
    });

    // 按源文件行号排序匹配结果
    matches.sort((a, b) => a.sourceIndex - b.sourceIndex);

    let lastSourceIndex = -1;
    let lastIncomingIndex = -1;

    // 处理所有行
    for (let i = 0; i < Math.max(data1.length, data2.length); i++) {
        const match = matches.find(m => m.sourceIndex === i || m.incomingIndex === i);

        if (match) {
            // 处理匹配行之前的未匹配行
            for (let j = lastSourceIndex + 1; j < match.sourceIndex; j++) {
                if (j >= 0 && j < data1.length) {
                    // 删除的行（添加复选框）
                    html += `
                        <tr>
                            <td>${j + 1}</td>
                            ${generateCells(data1[j], maxCols, 'deleted')}
                            <td class="separator"></td>
                            ${generateEmptyCells(maxCols)}
                            <td>
                                <input type="checkbox" class="row-select" 
                                       data-sheet="${sheetName}"
                                       data-type="deleted"
                                       data-index="${j}">
                            </td>
                        </tr>
                    `;
                    hasDifferences = true;
                }
            }

            for (let j = lastIncomingIndex + 1; j < match.incomingIndex; j++) {
                if (j >= 0 && j < data2.length) {
                    // 新增的行（添加复选框）
                    html += `
                        <tr>
                            <td>${j + 1}</td>
                            ${generateEmptyCells(maxCols)}
                            <td class="separator"></td>
                            ${generateCells(data2[j], maxCols, 'added')}
                            <td>
                                <input type="checkbox" class="row-select"
                                       data-sheet="${sheetName}"
                                       data-type="added"
                                       data-index="${j}">
                            </td>
                        </tr>
                    `;
                    hasDifferences = true;
                }
            }

            // 处理匹配行（保持不变或修改的行）
            const sourceRow = data1[match.sourceIndex];
            const incomingRow = data2[match.incomingIndex];
            
            if (sourceRow.length === incomingRow.length && 
                sourceRow.some((val, idx) => val !== incomingRow[idx])) {
                // 修改的行
                html += `
                    <tr>
                        <td>${match.sourceIndex + 1}</td>
                        ${generateCells(sourceRow, maxCols, 'modified')}
                        <td class="separator"></td>
                        ${generateCells(incomingRow, maxCols, 'modified')}
                        <td>
                            <label>
                                <input type="radio" name="modified_${sheetName}_${match.sourceIndex}" 
                                       class="row-select"
                                       data-sheet="${sheetName}"
                                       data-type="modified-source"
                                       data-index="${match.sourceIndex}"
                                       checked>
                                Source
                            </label>
                            <label>
                                <input type="radio" name="modified_${sheetName}_${match.sourceIndex}"
                                       class="row-select"
                                       data-sheet="${sheetName}"
                                       data-type="modified-incoming"
                                       data-index="${match.sourceIndex}">
                                Incoming
                            </label>
                        </td>
                    </tr>
                `;
                hasDifferences = true;
            } else {
                // 完全相同的行
                html += `
                    <tr>
                        <td>${match.sourceIndex + 1}</td>
                        ${generateCells(sourceRow, maxCols)}
                        <td class="separator"></td>
                        ${generateCells(incomingRow, maxCols)}
                        <td></td>
                    </tr>
                `;
            }

            lastSourceIndex = match.sourceIndex;
            lastIncomingIndex = match.incomingIndex;
        }
    }

    // 处理剩余的未匹配行
    for (let i = lastSourceIndex + 1; i < data1.length; i++) {
        if (!matchedSource.has(i)) {
            // 删除的行（添加复选框）
            html += `
                <tr>
                    <td>${i + 1}</td>
                    ${generateCells(data1[i], maxCols, 'deleted')}
                    <td class="separator"></td>
                    ${generateEmptyCells(maxCols)}
                    <td>
                        <input type="checkbox" class="row-select"
                               data-sheet="${sheetName}"
                               data-type="deleted"
                               data-index="${i}">
                    </td>
                </tr>
            `;
            hasDifferences = true;
        }
    }

    for (let i = lastIncomingIndex + 1; i < data2.length; i++) {
        if (!matchedIncoming.has(i)) {
            // 新增的行（添加复选框）
            html += `
                <tr>
                    <td>${i + 1}</td>
                    ${generateEmptyCells(maxCols)}
                    <td class="separator"></td>
                    ${generateCells(data2[i], maxCols, 'added')}
                    <td>
                        <input type="checkbox" class="row-select"
                               data-sheet="${sheetName}"
                               data-type="added"
                               data-index="${i}">
                    </td>
                </tr>
            `;
            hasDifferences = true;
        }
    }

    html += '</tbody></table>';
    
    if (!hasDifferences) {
        contentDiv.innerHTML = '<div class="no-diff">该工作表内容完全相同</div>';
    } else {
        contentDiv.innerHTML = html;
    }
}

// 添加生成单元格的辅助函数
function generateCells(row, maxCols, className = '') {
    let cells = '';
    for (let i = 0; i < maxCols; i++) {
        const value = row[i] !== undefined ? row[i] : '';
        cells += `<td class="${className}">${value}</td>`;
    }
    return cells;
}

// 添加生成空单元格的函数
function generateEmptyCells(count) {
    return Array(count).fill('<td></td>').join('');
}

// 修改行匹配逻辑
function compareRows(row1, row2) {
    if (row1.length !== row2.length) return false;
    
    for (let i = 0; i < row1.length; i++) {
        if (row1[i] !== row2[i]) return false;
    }
    return true;
}

// 合并差异并导出
function mergeDiff() {
    if (!workbook1 || !workbook2) {
        alert('请先选择并比较文件');
        return;
    }

    // 获取共同的sheet
    const commonSheets = workbook1.SheetNames.filter(name => workbook2.SheetNames.includes(name));
    
    // 生成预览HTML
    let previewHtml = '';
    
    commonSheets.forEach(sheetName => {
        const sheet1 = workbook1.Sheets[sheetName];
        const sheet2 = workbook2.Sheets[sheetName];
        
        const data1 = XLSX.utils.sheet_to_json(sheet1, {header: 1});
        const data2 = XLSX.utils.sheet_to_json(sheet2, {header: 1});

        // 获取用户选中的行
        const selectedRows = {
            deleted: new Set(
                Array.from(document.querySelectorAll(`.row-select[data-sheet="${sheetName}"][data-type="deleted"]:checked`))
                    .map(cb => parseInt(cb.dataset.index))
            ),
            added: new Set(
                Array.from(document.querySelectorAll(`.row-select[data-sheet="${sheetName}"][data-type="added"]:checked`))
                    .map(cb => parseInt(cb.dataset.index))
            ),
            modifiedSource: new Set(
                Array.from(document.querySelectorAll(`.row-select[data-sheet="${sheetName}"][data-type="modified-source"]:checked`))
                    .map(cb => parseInt(cb.dataset.index))
            ),
            modifiedIncoming: new Set(
                Array.from(document.querySelectorAll(`.row-select[data-sheet="${sheetName}"][data-type="modified-incoming"]:checked`))
                    .map(cb => parseInt(cb.dataset.index))
            )
        };

        // 生成预览内容
        previewHtml += generatePreview(data1, data2, selectedRows, sheetName);
    });

    // 显示预览
    document.getElementById('previewContent').innerHTML = previewHtml;
    document.getElementById('previewModal').style.display = 'block';
}

// 生成预览内容
function generatePreview(data1, data2, selectedRows, sheetName) {
    let html = `<h3>${sheetName}</h3>`;
    
    // 计算最大列数
    const maxCols = Math.max(
        ...data1.map(row => row.length),
        ...data2.map(row => row.length)
    );
    
    // 生成表头
    html += '<table class="diff-table"><thead><tr><th>行号</th>';
    for (let i = 0; i < maxCols; i++) {
        html += `<th>列${i + 1}</th>`;
    }
    html += '</tr></thead><tbody>';

    // 获取合并后的数据
    const mergedData = mergeSheetData(data1, data2, selectedRows);

    // 显示合并后的数据
    mergedData.forEach((row, index) => {
        html += `
            <tr>
                <td>${index + 1}</td>
                ${generatePreviewCells(row, maxCols)}
            </tr>
        `;
    });

    html += '</tbody></table>';
    return html;
}

// 生成预览单元格
function generatePreviewCells(row, maxCols) {
    let cells = '';
    for (let i = 0; i < maxCols; i++) {
        const value = row[i] !== undefined ? row[i] : '';
        cells += `<td>${value}</td>`;
    }
    return cells;
}

// 修改合并数据的函数
function mergeSheetData(data1, data2, selectedRows) {
    const mergedData = [];
    
    // 创建行内容的哈希映射
    const sourceMap = new Map();
    data1.forEach((row, index) => {
        const content = row.join('|');
        if (!sourceMap.has(content)) {
            sourceMap.set(content, []);
        }
        sourceMap.get(content).push(index);
    });

    // 标记已匹配的行
    const matchedSource = new Set();
    const matchedIncoming = new Set();
    const matches = [];

    // 找出匹配的行
    data2.forEach((incomingRow, incomingIndex) => {
        // 在源文件中查找匹配的行
        let bestMatch = null;
        let bestMatchIndex = -1;
        let maxMatchingCols = 0;

        data1.forEach((sourceRow, sourceIndex) => {
            if (!matchedSource.has(sourceIndex)) {
                // 计算匹配的列数
                let matchingCols = 0;
                const minCols = Math.min(sourceRow.length, incomingRow.length);
                
                for (let i = 0; i < minCols; i++) {
                    if (sourceRow[i] === incomingRow[i]) {
                        matchingCols++;
                    }
                }

                // 更新最佳匹配
                if (matchingCols > maxMatchingCols) {
                    maxMatchingCols = matchingCols;
                    bestMatch = sourceRow;
                    bestMatchIndex = sourceIndex;
                }
            }
        });

        // 如果找到足够好的匹配
        if (maxMatchingCols > 0) {
            matches.push({
                sourceIndex: bestMatchIndex,
                incomingIndex: incomingIndex,
                content: incomingRow
            });
            matchedSource.add(bestMatchIndex);
            matchedIncoming.add(incomingIndex);
        }
    });

    // 按源文件行号排序匹配结果
    matches.sort((a, b) => a.sourceIndex - b.sourceIndex);

    let lastSourceIndex = -1;
    let lastIncomingIndex = -1;

    // 处理所有行
    for (let i = 0; i < Math.max(data1.length, data2.length); i++) {
        const match = matches.find(m => m.sourceIndex === i || m.incomingIndex === i);

        if (match) {
            // 处理匹配行之前的未匹配行
            for (let j = lastSourceIndex + 1; j < match.sourceIndex; j++) {
                if (j >= 0 && j < data1.length && !matchedSource.has(j)) {
                    // 只添加被选中的删除行
                    if (selectedRows.deleted.has(j)) {
                        mergedData.push([...data1[j]]);
                    }
                }
            }

            for (let j = lastIncomingIndex + 1; j < match.incomingIndex; j++) {
                if (j >= 0 && j < data2.length && !matchedIncoming.has(j)) {
                    // 只添加被选中的新增行
                    if (selectedRows.added.has(j)) {
                        mergedData.push([...data2[j]]);
                    }
                }
            }

            // 处理匹配的行
            if (selectedRows.modifiedIncoming.has(match.sourceIndex)) {
                // 如果选择了incoming版本
                mergedData.push([...data2[match.incomingIndex]]);
            } else {
                // 默认使用source版本
                mergedData.push([...data1[match.sourceIndex]]);
            }

            lastSourceIndex = match.sourceIndex;
            lastIncomingIndex = match.incomingIndex;
        }
    }

    // 处理剩余的未匹配行
    for (let i = lastSourceIndex + 1; i < data1.length; i++) {
        if (!matchedSource.has(i) && selectedRows.deleted.has(i)) {
            mergedData.push([...data1[i]]);
        }
    }

    for (let i = lastIncomingIndex + 1; i < data2.length; i++) {
        if (!matchedIncoming.has(i) && selectedRows.added.has(i)) {
            mergedData.push([...data2[i]]);
        }
    }

    return mergedData;
}

// 关闭预览
function closePreview() {
    document.getElementById('previewModal').style.display = 'none';
}

// 确认导出
function confirmExport() {
    const mergedWorkbook = XLSX.utils.book_new();
    const allSheets = [...new Set([...workbook1.SheetNames, ...workbook2.SheetNames])];

    allSheets.forEach(sheetName => {
        const sheet1 = workbook1.Sheets[sheetName] || null;
        const sheet2 = workbook2.Sheets[sheetName] || null;
        
        let mergedData;
        if (sheet1 && sheet2) {
            // 如果两个文件都有这个sheet，进行合并
            const data1 = XLSX.utils.sheet_to_json(sheet1, {header: 1});
            const data2 = XLSX.utils.sheet_to_json(sheet2, {header: 1});
            const selectedRows = {
                deleted: new Set(
                    Array.from(document.querySelectorAll(`.row-select[data-sheet="${sheetName}"][data-type="deleted"]:checked`))
                        .map(cb => parseInt(cb.dataset.index))
                ),
                added: new Set(
                    Array.from(document.querySelectorAll(`.row-select[data-sheet="${sheetName}"][data-type="added"]:checked`))
                        .map(cb => parseInt(cb.dataset.index))
                ),
                modifiedSource: new Set(
                    Array.from(document.querySelectorAll(`.row-select[data-sheet="${sheetName}"][data-type="modified-source"]:checked`))
                        .map(cb => parseInt(cb.dataset.index))
                ),
                modifiedIncoming: new Set(
                    Array.from(document.querySelectorAll(`.row-select[data-sheet="${sheetName}"][data-type="modified-incoming"]:checked`))
                        .map(cb => parseInt(cb.dataset.index))
                )
            };
            mergedData = mergeSheetData(data1, data2, selectedRows);
        } else {
            // 如果只有一个文件有这个sheet，直接使用该sheet的数据
            const sheet = sheet1 || sheet2;
            mergedData = XLSX.utils.sheet_to_json(sheet, {header: 1});
        }
        
        const mergedSheet = XLSX.utils.aoa_to_sheet(mergedData);
        XLSX.utils.book_append_sheet(mergedWorkbook, mergedSheet, sheetName);
    });

    try {
        XLSX.writeFile(mergedWorkbook, 'merged_file.xlsx');
        alert('文件导出成功！');
    } catch (error) {
        alert('导出文件失败：' + error.message);
    }

    closePreview();
}

// 显示单个sheet的内容
function displaySingleSheet(data, sheetName, type) {
    const contentDiv = document.getElementById(`content-${sheetName}`);
    
    // 计算最大列数
    const maxCols = Math.max(...data.map(row => row.length));
    
    // 生成表头
    let html = '<table class="diff-table"><thead><tr><th>行号</th>';
    
    // 如果是source文件的sheet
    if (type === 'source') {
        // Source文件列
        for (let i = 0; i < maxCols; i++) {
            html += `<th>Source列${i + 1}</th>`;
        }
        html += '<th class="separator"></th>';
        // Incoming文件列（空）
        for (let i = 0; i < maxCols; i++) {
            html += `<th>Incoming列${i + 1}</th>`;
        }
    } else {
        // Source文件列（空）
        for (let i = 0; i < maxCols; i++) {
            html += `<th>Source列${i + 1}</th>`;
        }
        html += '<th class="separator"></th>';
        // Incoming文件列
        for (let i = 0; i < maxCols; i++) {
            html += `<th>Incoming列${i + 1}</th>`;
        }
    }
    
    html += '<th>选择</th></tr></thead><tbody>';

    // 显示数据
    data.forEach((row, index) => {
        html += `<tr><td>${index + 1}</td>`;
        
        if (type === 'source') {
            // 显示source数据
            html += generateCells(row, maxCols);
            html += '<td class="separator"></td>';
            html += generateEmptyCells(maxCols);
        } else {
            // 显示incoming数据
            html += generateEmptyCells(maxCols);
            html += '<td class="separator"></td>';
            html += generateCells(row, maxCols);
        }
        
        html += '<td></td></tr>';
    });

    html += '</tbody></table>';
    contentDiv.innerHTML = html;
}

// 跳转到下一个差异
function jumpToNextDiff() {
    const activeTab = document.querySelector('.tab-content.active');
    if (!activeTab) return;
    
    // 获取所有差异行（包括修改、新增和删除的行）
    const diffRows = activeTab.querySelectorAll('tr:has(td.modified), tr:has(td.added), tr:has(td.deleted)');
    if (diffRows.length === 0) return;
    
    // 更新当前索引
    currentDiffIndex = (currentDiffIndex + 1) % diffRows.length;
    
    // 获取目标行
    const targetRow = diffRows[currentDiffIndex];
    
    // 移除之前的高亮
    activeTab.querySelectorAll('.highlight-diff').forEach(el => {
        el.classList.remove('highlight-diff');
    });
    
    // 添加新的高亮
    targetRow.classList.add('highlight-diff');
    
    // 滚动到目标位置
    targetRow.scrollIntoView({
        behavior: 'smooth',
        block: 'center'
    });
} 