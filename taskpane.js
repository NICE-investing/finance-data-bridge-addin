// Office 초기화
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("select-range").onclick = selectRange;
        document.getElementById("upload-btn").onclick = uploadData;
        document.getElementById("validate-btn").onclick = validateDSL;
        document.getElementById("save-settings").onclick = saveSettings;
        
        // 탭 전환 이벤트
        document.querySelectorAll('.tab-button').forEach(button => {
            button.addEventListener('click', switchTab);
        });
        
        // 설정 로드
        loadSettings();
        
        // URL 파라미터 처리 (더 이상 필요없지만 호환성을 위해 남겨둠)
        const urlParams = new URLSearchParams(window.location.search);
        const tab = urlParams.get('tab');
        if (tab === 'validate') {
            switchToTab('validate');
        } else {
            // 기본적으로 업로드 탭 표시 (이미 활성화되어 있음)
        }
    }
});

// 프로그래밍적으로 탭 전환
function switchToTab(tabName) {
    // 모든 탭 버튼과 콘텐츠 비활성화
    document.querySelectorAll('.tab-button').forEach(btn => {
        btn.classList.remove('active');
    });
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    
    // 지정한 탭 활성화
    const targetButton = document.querySelector(`[data-tab="${tabName}"]`);
    const targetContent = document.getElementById(`${tabName}-tab`);
    
    if (targetButton && targetContent) {
        targetButton.classList.add('active');
        targetContent.classList.add('active');
    }
}

// API 설정 관리
function getApiUrl() {
    const stored = localStorage.getItem('apiUrl');
    return stored || document.getElementById('api-url').value;
}

function saveSettings() {
    const apiUrl = document.getElementById('api-url').value;
    localStorage.setItem('apiUrl', apiUrl);
    showNotification('설정이 저장되었습니다.', 'success');
}

function loadSettings() {
    const apiUrl = localStorage.getItem('apiUrl');
    if (apiUrl) {
        document.getElementById('api-url').value = apiUrl;
    }
}

// 탭 전환
function switchTab(event) {
    const targetTab = event.target.dataset.tab;
    
    // 모든 탭 버튼과 콘텐츠 비활성화
    document.querySelectorAll('.tab-button').forEach(btn => {
        btn.classList.remove('active');
    });
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    
    // 선택한 탭 활성화
    event.target.classList.add('active');
    document.getElementById(`${targetTab}-tab`).classList.add('active');
}

// 범위 선택
async function selectRange() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load("address");
            await context.sync();
            
            document.getElementById("selected-range").textContent = `선택된 범위: ${range.address}`;
            sessionStorage.setItem('selectedRange', range.address);
        });
    } catch (error) {
        console.error("범위 선택 오류:", error);
        showNotification("범위 선택 중 오류가 발생했습니다.", "error");
    }
}

// 데이터 업로드
async function uploadData() {
    const selectedRange = sessionStorage.getItem('selectedRange');
    
    if (!selectedRange) {
        showNotification('데이터 범위를 선택해주세요.', 'warning');
        return;
    }
    
    showLoading('upload-btn', true);
    
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.worksheets.getActiveWorksheet().getRange(selectedRange);
            range.load("values");
            await context.sync();
            
            // 데이터 변환
            const data = convertToAccountData(range.values);
            
            // API 호출
            const response = await fetch(`${getApiUrl()}/api/v1/accounts/bulk`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(data)
            });
            
            const result = await response.json();
            
            if (response.ok) {
                showResult('upload-result', {
                    success: true,
                    message: `${result.success_count}개 계정이 업로드되었습니다.${result.error_count > 0 ? ` (${result.error_count}개 실패)` : ''}`,
                    data: result
                });
            } else {
                throw new Error(result.error || '업로드 실패');
            }
        });
    } catch (error) {
        console.error("업로드 오류:", error);
        showResult('upload-result', {
            success: false,
            message: error.message || '업로드 중 오류가 발생했습니다.'
        });
    } finally {
        showLoading('upload-btn', false);
    }
}

// Excel 데이터를 계정 데이터로 변환
function convertToAccountData(values) {
    const accounts = [];
    const headers = values[1]; // 두번째 행을 헤더로 가정
    
    // 헤더 인덱스 찾기
    const codeIndex = headers.findIndex(h => h && h.toString().includes('계정코드') || h.toString().includes('code'));
    const nameKrIndex = headers.findIndex(h => h && h.toString().includes('계정명') || h.toString().includes('name_kr'));
    const nameEnIndex = headers.findIndex(h => h && h.toString().includes('계정명(영문)') || h.toString().includes('name_en'));
    const reportCodeIndex = headers.findIndex(h => h && h.toString().includes('보고서') || h.toString().includes('report_code'));
    const formulaIndex = headers.findIndex(h => h && h.toString().includes('NIS맵핑산식') || h.toString().includes('nis_map_formula'));
    const descriptionIndex = headers.findIndex(h => h && h.toString().includes('비고') || h.toString().includes('description'));
    const activeIndex = headers.findIndex(h => h && h.toString().includes('활성') || h.toString().includes('active'));
    const sortOrderIndex = headers.findIndex(h => h && h.toString().includes('정렬') || h.toString().includes('sort_order'));
    const parentCodeIndex = headers.findIndex(h => h && h.toString().includes('부모') || h.toString().includes('parent_code'));
    const levelIndex = headers.findIndex(h => h && h.toString().includes('레벨') || h.toString().includes('level'));
    
    // 데이터 행 처리
    for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if (!row || !row[codeIndex]) continue;
        
        const account = {
            account_code: row[codeIndex].toString(),
            account_name: row[nameKrIndex] ? row[nameKrIndex].toString() : '',
            report_code: row[reportCodeIndex] ? row[reportCodeIndex].toString() : null,
            description: row[descriptionIndex] ? row[descriptionIndex].toString() : null,
            active: activeIndex >= 0 && row[activeIndex] !== undefined ? 
                Boolean(row[activeIndex]) : true,
            sort_order: sortOrderIndex >= 0 && row[sortOrderIndex] ? 
                parseInt(row[sortOrderIndex]) : null,
            parent_code: row[parentCodeIndex] ? row[parentCodeIndex].toString() : null,
            level: levelIndex >= 0 && row[levelIndex] ? 
                parseInt(row[levelIndex]) : 1,
            formula_dsl: row[formulaIndex] ? row[formulaIndex].toString() : null
        };
        
        accounts.push(account);
    }
    
    return accounts;
}

// DSL 검증
async function validateDSL() {
    const dslInput = document.getElementById('dsl-input').value;
    
    if (!dslInput.trim()) {
        showNotification('DSL 수식을 입력해주세요.', 'warning');
        return;
    }
    
    showLoading('validate-btn', true);
    
    try {
        const response = await fetch(`${getApiUrl()}/api/v1/validate`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                formula_dsl: dslInput
            })
        });
        
        const result = await response.json();
        
        if (response.ok) {
            showResult('validate-result', {
                success: result.valid,
                message: result.valid ? 'DSL 구문이 유효합니다.' : 'DSL 구문 오류',
                data: result
            });
            
            // 유효한 경우 AST 표시
            // if (result.valid && result.ast) {
            //     displayAST(result.ast);
            // }
        } else {
            throw new Error(result.error || '검증 실패');
        }
    } catch (error) {
        console.error("DSL 검증 오류:", error);
        showResult('validate-result', {
            success: false,
            message: error.message || 'DSL 검증 중 오류가 발생했습니다.'
        });
    } finally {
        showLoading('validate-btn', false);
    }
}

// AST 트리 표시
function displayAST(ast) {
    const resultDiv = document.getElementById('validate-result');
    const astHtml = `
        <h3>구문 분석 결과 (AST)</h3>
        <pre>${JSON.stringify(ast, null, 2)}</pre>
    `;
    resultDiv.innerHTML += astHtml;
}

// 결과 표시
function showResult(elementId, result) {
    const resultDiv = document.getElementById(elementId);
    resultDiv.className = 'result-area show';
    
    if (result.success) {
        resultDiv.classList.add('success');
    } else {
        resultDiv.classList.add('error');
    }
    
    let html = `<h3>${result.message}</h3>`;
    
    if (result.data) {
        if (result.data.errors && result.data.errors.length > 0) {
            html += '<h4>오류 상세:</h4><ul>';
            result.data.errors.forEach(error => {
                html += `<li>${error}</li>`;
            });
            html += '</ul>';
        } else if (typeof result.data === 'object') {
            html += `<pre>${JSON.stringify(result.data, null, 2)}</pre>`;
        }
    }
    
    resultDiv.innerHTML = html;
}

// 알림 표시
function showNotification(message, type) {
    // 간단한 알림 구현 (실제로는 더 나은 UI 라이브러리 사용 권장)
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.textContent = message;
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 1rem;
        background: ${type === 'success' ? '#4caf50' : type === 'error' ? '#f44336' : '#ff9800'};
        color: white;
        border-radius: 4px;
        z-index: 1000;
        animation: slideIn 0.3s;
    `;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.remove();
    }, 3000);
}

// 로딩 상태 표시
function showLoading(buttonId, show) {
    const button = document.getElementById(buttonId);
    if (show) {
        button.disabled = true;
        button.innerHTML = button.textContent + '<span class="loading"></span>';
    } else {
        button.disabled = false;
        button.innerHTML = button.textContent.replace('<span class="loading"></span>', '');
    }
}

// 엑셀에 결과 쓰기 (선택적 기능)
async function writeResultToExcel(data, startCell = 'A1') {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(startCell);
            
            // 데이터를 2차원 배열로 변환
            const values = [];
            if (Array.isArray(data)) {
                data.forEach(item => {
                    values.push([item.code, item.name, item.value]);
                });
            }
            
            if (values.length > 0) {
                const dataRange = range.getResizedRange(values.length - 1, values[0].length - 1);
                dataRange.values = values;
                await context.sync();
            }
        });
    } catch (error) {
        console.error("Excel 쓰기 오류:", error);
    }
}