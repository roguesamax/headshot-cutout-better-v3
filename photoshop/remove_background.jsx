/*
  Version-tolerant Photoshop background removal helper.
  Strategy:
  1) Try Remove Background IDs (new/old variants)
  2) Fallback to Select Subject IDs + reveal-selection mask
*/

app.displayDialogs = DialogModes.NO;

function unlockLayerIfNeeded(doc) {
    try {
        if (doc.activeLayer && doc.activeLayer.isBackgroundLayer) {
            doc.activeLayer.isBackgroundLayer = false;
        }
    } catch (e) {}
}

function tryRunMenu(idName) {
    try {
        app.runMenuItem(stringIDToTypeID(idName));
        return true;
    } catch (e) {
        return false;
    }
}

function tryExecAction(idName) {
    try {
        executeAction(stringIDToTypeID(idName), undefined, DialogModes.NO);
        return true;
    } catch (e) {
        return false;
    }
}

function removeBackgroundCompat() {
    if (tryExecAction('removeBackground')) return true;
    if (tryRunMenu('removeBackground')) return true;
    if (tryRunMenu('autoCutout')) return true;
    if (tryRunMenu('autoCutoutSubject')) return true;
    return false;
}

function selectSubjectCompat() {
    if (tryExecAction('selectSubject')) return true;
    if (tryRunMenu('selectSubject')) return true;
    if (tryRunMenu('autoCutoutSubject')) return true;
    return false;
}

function applySelectionMask(doc) {
    var idMk = charIDToTypeID('Mk  ');
    var desc = new ActionDescriptor();
    var idNw = charIDToTypeID('Nw  ');
    var idChnl = charIDToTypeID('Chnl');
    desc.putClass(idNw, idChnl);

    var idAt = charIDToTypeID('At  ');
    var ref = new ActionReference();
    ref.putEnumerated(charIDToTypeID('Chnl'), charIDToTypeID('Chnl'), charIDToTypeID('Msk '));
    desc.putReference(idAt, ref);

    var idUsng = charIDToTypeID('Usng');
    var idUsrM = charIDToTypeID('UsrM');
    var idRvlS = charIDToTypeID('RvlS');
    desc.putEnumerated(idUsng, idUsrM, idRvlS);

    executeAction(idMk, desc, DialogModes.NO);

    try {
        doc.selection.deselect();
    } catch (e) {}
}

if (app.documents.length > 0) {
    var doc = app.activeDocument;
    unlockLayerIfNeeded(doc);

    var removed = removeBackgroundCompat();
    if (!removed) {
        var selected = selectSubjectCompat();
        if (!selected) {
            throw new Error('No compatible Remove Background/Select Subject action found in this Photoshop build.');
        }
        applySelectionMask(doc);
    }
}
