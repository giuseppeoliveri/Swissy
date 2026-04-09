#target "InDesign"

    (function () {
        var layerName = "Swissy Layer";

        if (app.documents.length === 0) {
            alert("Errore: Apri prima un documento.");
            return;
        }

        var doc = app.activeDocument;

        // Utility
        var old_x_units, old_y_units;

        function init() {
            app.scriptPreferences.enableRedraw = false;
            with (doc.viewPreferences) {
                old_x_units = horizontalMeasurementUnits;
                old_y_units = verticalMeasurementUnits;
                horizontalMeasurementUnits = MeasurementUnits.points;
                verticalMeasurementUnits = MeasurementUnits.points;
            }
        }

        function done() {
            with (doc.viewPreferences) {
                try {
                    horizontalMeasurementUnits = old_x_units;
                    verticalMeasurementUnits = old_y_units;
                } catch (e) { }
            }
            app.scriptPreferences.enableRedraw = true;
        }

        if (app.selection.length === 0) {
            alert("Attenzione: Seleziona una cornice di testo.");
            return;
        }

        // Parametri base
        var flatterzone = 0;
        var randomness = 0;
        var thicknessPct = 0;
        var initMode = 0;
        var vOffset = -30;

        // Ripristino valori salvati
        for (var n = 0; n < app.selection.length; n++) {
            if (app.selection[n].constructor.name === "TextFrame") {
                var tf = app.selection[n];
                var sf = tf.extractLabel("swissy_flatterzone");
                var sr = tf.extractLabel("swissy_randomness");
                var st = tf.extractLabel("swissy_thickness");
                var sm = tf.extractLabel("swissy_startMode");

                if (sf !== "") flatterzone = parseFloat(sf);
                if (sr !== "") randomness = parseFloat(sr);
                if (st !== "") thicknessPct = parseFloat(st);
                if (sm !== "") initMode = parseInt(sm, 10);
                break;
            }
        }

        // UI
        var w = new Window("dialog", "Swissy 1.0 Alpha");
        w.preferredSize.width = 400;
        w.orientation = "column";
        w.alignChildren = "fill";

        // Impostazione Sfondo #640000 (rosso scuro RGB = 100, 0, 0)
        var darkRed = [100 / 255, 0, 0, 1];
        w.graphics.backgroundColor = w.graphics.newBrush(w.graphics.BrushType.SOLID_COLOR, darkRed);

        // Logo
        var logoFile = new File("~/Desktop/Swissy Mascotte.png");
        if (!logoFile.exists) {
            try {
                var scriptF = app.activeScript;
                logoFile = new File(scriptF.parent.fsName + "/Swissy Mascotte.png");
            } catch (e) {
                try { logoFile = new File(new File(e.fileName).parent.fsName + "/Swissy Mascotte.png"); } catch (ex) { }
            }
        }

        if (logoFile && logoFile.exists) {
            var gLogo = w.add("group");
            gLogo.alignment = "center";
            gLogo.preferredSize = [150, 150]; // Area quadrata per non tagliare nulla

            gLogo.onDraw = function () {
                try {
                    var img = ScriptUI.newImage(logoFile);
                    if (img && img.size && img.size[0] > 0) {
                        var imgW = img.size[0];
                        var imgH = img.size[1];

                        // Calcola scala proporzionale
                        var scale = Math.min(this.size[0] / imgW, this.size[1] / imgH);
                        var newW = imgW * scale;
                        var newH = imgH * scale;

                        // Centratura automatica
                        var dX = (this.size[0] - newW) / 2;
                        var dY = (this.size[1] - newH) / 2;

                        this.graphics.drawImage(img, dX, dY, newW, newH);
                    }
                } catch (e) { }
            };
        }

        var p = w.add("panel", undefined, "By G. Oliveri");
        p.alignChildren = "fill";
        p.graphics.backgroundColor = p.graphics.newBrush(p.graphics.BrushType.SOLID_COLOR, darkRed);

        // Slider 1
        var g1 = p.add("group");
        g1.add("statictext", undefined, "Dimensione spazio bianco");
        var lbl1 = g1.add("statictext", undefined, flatterzone + "%");
        lbl1.preferredSize.width = 45;
        var slider1 = p.add("slider", undefined, flatterzone, 0, 100);
        var upd1 = function () { lbl1.text = Math.round(slider1.value) + "%"; flatterzone = slider1.value; };
        slider1.onChanging = upd1; slider1.onChange = upd1;

        // Slider 2
        var g2 = p.add("group");
        g2.add("statictext", undefined, "Casualit\u00E0");
        var lbl2 = g2.add("statictext", undefined, randomness + "%");
        lbl2.preferredSize.width = 45;
        var slider2 = p.add("slider", undefined, randomness, 0, 100);
        var upd2 = function () { lbl2.text = Math.round(slider2.value) + "%"; randomness = slider2.value; };
        slider2.onChanging = upd2; slider2.onChange = upd2;

        // Slider 3
        var g3 = p.add("group");
        g3.add("statictext", undefined, "Spessore spazio bianco");
        var lbl3 = g3.add("statictext", undefined, thicknessPct + "%");
        lbl3.preferredSize.width = 45;
        var slider3 = p.add("slider", undefined, thicknessPct, 0, 150);
        var upd3 = function () { lbl3.text = Math.round(slider3.value) + "%"; thicknessPct = slider3.value; };
        slider3.onChanging = upd3; slider3.onChange = upd3;

        // Dropdown Alternanza Righe
        var gMode = p.add("group");
        gMode.add("statictext", undefined, "Applica a righe:");
        var modeDropdown = gMode.add("dropdownlist", undefined, ["Dispari (Inizia da 1a riga)", "Pari (Inizia da 2a riga)"]);
        modeDropdown.selection = initMode;

        // Pulsanti
        var btn_group = w.add("group");
        btn_group.alignment = "center";
        btn_group.add("button", undefined, "OK", { name: "ok" });
        btn_group.add("button", undefined, "Annulla", { name: "cancel" });

        if (w.show() != 1) return;

        // Esecuzione
        init();

        var startMode = modeDropdown.selection.index;

        // Cache valori
        var save_flatter = Math.round(flatterzone);
        var save_random = Math.round(randomness);
        var save_thick = Math.round(thicknessPct);
        var save_mode = startMode;

        // Conversioni percentuali
        flatterzone *= 0.01;
        randomness *= 0.01;
        thicknessPct *= 0.01;
        vOffset *= 0.01;

        // Setup layer
        var myLayer = doc.layers.item(layerName);
        if (!myLayer.isValid) {
            myLayer = doc.layers.add({ name: layerName });
        }
        myLayer.printable = false; // Imposta il livello come Non Stampabile

        for (var n = 0; n < app.selection.length; n++) {
            var sel = app.selection[n];
            if (sel.constructor.name != "TextFrame") continue;

            // Iniezione dati
            sel.insertLabel("swissy_flatterzone", save_flatter.toString());
            sel.insertLabel("swissy_randomness", save_random.toString());
            sel.insertLabel("swissy_thickness", save_thick.toString());
            sel.insertLabel("swissy_startMode", save_mode.toString());

            var bounds = sel.geometricBounds;
            var frameRight = bounds[3];
            var frameWidth = frameRight - bounds[1];
            var tfID = sel.id.toString();

            // Pulizia vecchie guide
            for (var g = doc.graphicLines.length - 1; g >= 0; g--) {
                var oldLine = doc.graphicLines[g];
                if (oldLine.itemLayer == myLayer && oldLine.label == tfID) {
                    oldLine.remove();
                }
            }

            var container = sel.parent;

            for (var i = 0; i < sel.lines.length; i++) {
                var currentLine = sel.lines[i];

                if (!currentLine.isValid) continue;

                // Filtro righe
                if (startMode === 0 && i % 2 !== 0) continue;
                if (startMode === 1 && i % 2 === 0) continue;

                // Check valori minimi
                if (flatterzone === 0 && randomness === 0) continue;

                var fontSize = currentLine.pointSize;
                var strokeW = fontSize * thicknessPct;
                var lineY = currentLine.baseline + (fontSize * vOffset);

                var baseWidth = frameWidth * flatterzone;
                var variance = 0;

                if (randomness > 0) {
                    variance = (frameWidth * randomness) * (Math.random() - 0.5);
                }

                var x_start = frameRight;
                var x_end = frameRight - (baseWidth + variance);

                // Controllo margini
                if (x_end > frameRight) x_end = frameRight;
                if (x_end < bounds[1]) x_end = bounds[1];

                if (x_start === x_end) continue;

                // Disegno
                var line = container.graphicLines.add({ itemLayer: myLayer });
                line.geometricBounds = [lineY, x_start, lineY, x_end];
                line.label = tfID;

                // Stile
                line.strokeColor = "Magenta";
                line.strokeWeight = strokeW;
                try { line.transparencySettings.blendingSettings.opacity = 50; } catch (e) { }

                // Wrap
                line.textWrapPreferences.textWrapMode = TextWrapModes.BOUNDING_BOX_TEXT_WRAP;
                line.textWrapPreferences.textWrapOffset = [0, 0, 0, 0];
                try { line.textWrapPreferences.textWrapSide = TextWrapSideOptions.LEFT_SIDE; } catch (e) { }
            }
        }

        done();
        alert("Operazione Completata.\nPuoi eseguire ripetutamente lo script sullo stesso blocco di testo per aggiustarlo: gli spazi magenta vecchi verranno cancellati.");
    })();