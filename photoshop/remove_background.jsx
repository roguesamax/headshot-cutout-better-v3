/*
  Example JSX hook for Photoshop background removal.
  You can adapt this script and wire it via PHOTOSHOP_BG_JSX.
  Depending on Photoshop version, the exact action names for "Remove Background"
  may differ.
*/

if (app.documents.length > 0) {
    var doc = app.activeDocument;
    app.runMenuItem(stringIDToTypeID('autoCutout'));
    // Save logic can be extended for automated pipeline usage.
}
