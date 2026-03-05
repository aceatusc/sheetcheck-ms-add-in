/**
 * officeInit.js
 * Handles Office.onReady and host detection.
 */
const OfficeInit = (() => {
    let _isReady = false;
    let _host    = null;

    async function init() {
        return new Promise((resolve) => {
            Office.onReady((info) => {
                _isReady = true;
                _host    = info.host;
                console.log('[OfficeInit] Ready. Host:', _host);
                resolve({ host: _host });
            });
        });
    }

    function isExcel() {
        return _host === Office.HostType.Excel;
    }

    return { init, isExcel };
})();
