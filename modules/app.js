/**
 * app.js
 * Entry point — initializes Office and wires all modules together.
 *
 * Init order matters:
 *   DagStore  (no deps)
 *   RubricManager (DOM only)
 *   StepNavigator (DOM + DagRunner + RubricManager)
 *   ChatManager   (DOM + DagRunner + RubricManager)
 */
(async () => {
    const { host } = await OfficeInit.init();

    if (!OfficeInit.isExcel()) {
        console.warn('[App] This add-in is designed for Excel.');
    }

    DagStore.init();
    RubricManager.init();
    AspectManager.init();
    StepNavigator.init();
    ChatManager.init();

    console.log('[App] Ready.');
})();
