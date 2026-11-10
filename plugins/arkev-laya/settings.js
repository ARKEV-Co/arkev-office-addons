// The only Core addresses this plugin talks to. Core puts its address and a short-lived plugin
// token in the editor config (editorConfig.plugins.options, signed with the config); any other
// address is ignored, so a changed config can't send a document's text anywhere else.
window.ARKEV_LAYA_CORE = Object.freeze(["https://api.arkev.co"]);
