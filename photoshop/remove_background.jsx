/*
  Version-tolerant Photoshop background removal helper.
  Strategy:
  1) Try Remove Background menu IDs.
  2) Fallback to Select Subject + reveal-selection mask.
*/

app.displayDialogs = DialogModes.NO;

function unlockLayerIfNeeded(doc) {
    try {
        if (doc.activeLayer && doc.activeLayer.isBackgroundLayer) {
            doc.activeLayer.isBackgroundLayer = false;
        }
    } catch (e) {}
}

function runRemoveBackgroundMenu() {
    try {
        app.runMenuItem(stringIDToTypeID('autoCutout'));
        return true;
    } catch (e1) {
        try {
            app.runMenuItem(stringIDToTypeID('autoCutoutSubject'));
            return true;
        } catch (e2) {
            return false;
        }
    }
}

function selectSubjectAndMask(doc) {
    var didSelect = false;
    try {
        executeAction(stringIDToTypeID('selectSubject'), undefined, DialogModes.NO);
        didSelect = true;
    } catch (s1) {
        try {
            app.runMenuItem(stringIDToTypeID('autoCutoutSubject'));
            didSelect = true;
        } catch (s2) {
            didSelect = false;
        }
    }

    if (!didSelect) {
        throw new Error('No compatible Select Subject / Remove Background action found.');
    }

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
    if (!runRemoveBackgroundMenu()) {
        selectSubjectAndMask(doc);
    }
}
