/**
 * app.js
 * Entry point — initializes Office and wires all modules together.
 */
(async () => {
    const { host } = await OfficeInit.init();

    if (!OfficeInit.isExcel()) {
        console.warn('[App] This add-in is designed for Excel.');
        // PLACEHOLDER: show a friendly "unsupported host" message in the UI
    }

    StepNavigator.init();
    ChatManager.init();

    console.log('[App] Ready.');
})();
