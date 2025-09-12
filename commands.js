// 리본 명령 처리를 위한 함수들

Office.onReady(() => {
    // Office가 준비되면 초기화
});

// Finance Data Bridge 메인 명령
function openMainPane(event) {
    Office.ribbon.requestUpdate({
        tabs: [{
            id: "TabHome",
            groups: [{
                id: "FinanceDataBridge.Group",
                controls: [{
                    id: "MainButton",
                    enabled: true
                }]
            }]
        }]
    });
    
    // 작업창 열기
    Office.addin.showAsTaskpane("https://nice-investing.github.io/finance-data-bridge-addin/taskpane.html");
    
    event.completed();
}