
(function (window, undefined) {
    "use strict";

    function setStatus(text, isError) {
        var el = document.getElementById("status");
        if (!el) return;
        el.textContent = text;
        el.style.color = isError ? "#9b1c1c" : "#333";
    }

    window.formatCoastalLetter = function () {
        var btn = document.getElementById("formatBtn");
        if (btn) btn.disabled = true;
        setStatus("Formatting current letter…", false);

        window.Asc.plugin.callCommand(function () {
            try {

    var doc = Api.GetDocument();
    var paragraphs = doc.GetAllParagraphs();

    /*
      IMPORTANT:
      Lyrebird/Xestro is inserting genuinely EMPTY PARAGRAPHS between some
      Past Medical History diagnoses. Setting paragraph spacing to 0 does not
      remove those blank paragraphs, so delete them first.

      Do this backwards so deleting a paragraph does not disturb the indexes
      of paragraphs we still need to inspect.
    */
    var historyStart = -1;
    var historyEnd = -1;

    for (var pre = 0; pre < paragraphs.length; pre++) {
        var preText = (paragraphs[pre].GetText({
            "Numbering": false,
            "Math": true,
            "NewLineSeparator": "\n",
            "TabSymbol": "\t"
        }) || "").replace(/\u00A0/g, " ").trim().toUpperCase();

        if (preText === "PAST MEDICAL HISTORY:" || preText === "MEDICAL HISTORY:") {
            historyStart = pre;
            continue;
        }

        if (historyStart >= 0 && pre > historyStart && preText &&
            /^[A-Z][A-Z &\/\-]+:$/.test(preText)) {
            historyEnd = pre;
            break;
        }
    }

    if (historyStart >= 0) {
        if (historyEnd < 0) historyEnd = paragraphs.length;

        for (var del = historyEnd - 1; del > historyStart; del--) {
            var delText = (paragraphs[del].GetText({
                "Numbering": false,
                "Math": true,
                "NewLineSeparator": "\n",
                "TabSymbol": "\t"
            }) || "").replace(/\u00A0/g, " ").trim();

            if (!delText) {
                paragraphs[del].Delete();
            }
        }

        // Refresh after deletion.
        paragraphs = doc.GetAllParagraphs();
    }

    // 1.27 cm ≈ 0.5 inch = 720 twips.
    var INDENT = 720;

    // Common section headings used across Coastal respiratory/sleep letters.
    var SECTION_HEADINGS = {
        "PAST MEDICAL HISTORY:": true,
        "MEDICAL HISTORY:": true,
        "RESPIRATORY HISTORY:": true,
        "SLEEP HISTORY:": true,
        "SMOKING HISTORY:": true,
        "SOCIAL HISTORY:": true,
        "SOCIAL & EXPOSURE HISTORY:": true,
        "SOCIAL AND EXPOSURE HISTORY:": true,
        "OCCUPATIONAL HISTORY:": true,
        "OCCUPATIONAL & EXPOSURE HISTORY:": true,
        "EXPOSURE HISTORY:": true,
        "FAMILY HISTORY:": true,
        "MEDICATIONS:": true,
        "CURRENT MEDICATIONS:": true,
        "ALLERGIES:": true,
        "ALLERGIES AND ADR:": true,
        "ALLERGIES & ADR:": true,
        "INVESTIGATIONS:": true,
        "EXAMINATION:": true,
        "EXAMINATION FINDINGS:": true,
        "ASSESSMENT:": true,
        "ASSESSMENT AND ISSUES:": true,
        "ASSESSMENT & ISSUES:": true,
        "IMPRESSION:": true,
        "IMPRESSION / DIAGNOSIS:": true,
        "IMPRESSION/DIAGNOSIS:": true,
        "DIAGNOSIS:": true,
        "CURRENT PROBLEMS:": true,
        "MANAGEMENT PLAN:": true,
        "PLAN:": true,
        "FOLLOW-UP:": true,
        "FOLLOW UP:": true
    };

    var HISTORY_HEADINGS = {
        "PAST MEDICAL HISTORY:": true,
        "MEDICAL HISTORY:": true
    };

    var ASSESSMENT_HEADINGS = {
        "ASSESSMENT:": true,
        "ASSESSMENT AND ISSUES:": true,
        "ASSESSMENT & ISSUES:": true,
        "IMPRESSION:": true,
        "IMPRESSION / DIAGNOSIS:": true,
        "IMPRESSION/DIAGNOSIS:": true,
        "DIAGNOSIS:": true,
        "CURRENT PROBLEMS:": true
    };

    var PLAN_HEADINGS = {
        "MANAGEMENT PLAN:": true,
        "PLAN:": true
    };

    function cleanText(p) {
        return (p.GetText({
            "Numbering": false,
            "Math": true,
            "NewLineSeparator": "\n",
            "TabSymbol": "\t"
        }) || "").replace(/\u00A0/g, " ").trim();
    }

    function normHeading(s) {
        return s.replace(/\s+/g, " ").trim().toUpperCase();
    }

    function isHeading(text) {
        return !!SECTION_HEADINGS[normHeading(text)];
    }

    function isSignatureOrFooter(text) {
        var t = text.trim();
        return /^Yours sincerely[,;]?$/i.test(t) ||
               /^Kind regards[,;]?$/i.test(t) ||
               /^Regards[,;]?$/i.test(t) ||
               /^Dr\s+Jonathan\s+Banks\b/i.test(t) ||
               /^Resp\/Sleep\/Gen Med Phys\./i.test(t) ||
               /^Coastal Respiratory & Sleep Specialists$/i.test(t) ||
               /^Prov:/i.test(t) ||
               /^cc:/i.test(t) ||
               /^This letter may have been produced with the assistance of AI technology/i.test(t) ||
               /^Coastal Respiratory and Sleep Services/i.test(t);
    }

    function looksLikeInvestigation(text) {
        var t = text.trim();

        // Tests / imaging / pathology commonly used in respiratory & sleep letters.
        return /^(RFTs?|PFTs?|Lung function|Spirometry|Plethysmography|DLCO|KCO|FeNO|6[- ]?minute walk|6MWT|Walk test|CT\b|HRCT\b|CTPA\b|PET\b|PET\/CT\b|MRI\b|CXR\b|Chest X[- ]?ray|X[- ]?ray|Ultrasound\b|Echo\b|Echocardiogram\b|ECG\b|Holter\b|ABG\b|VBG\b|Blood\b|FBC\b|UEC\b|LFT\b|CRP\b|ESR\b|IgE\b|IgG\b|Eosinophils?\b|Serology\b|ANA\b|ANCA\b|ENA\b|Biopsy\b|Lung biopsy\b|Histology\b|Pathology\b|Respiratory culture\b|Sputum\b|BAL\b|Bronchoscopy\b|EBUS\b|Sleep study\b|PSG\b|Polysomnography\b|Oximetry\b|CPAP download\b|Device download\b)/i.test(t);
    }

    function makeNumbering() {
        return doc.CreateNumbering("numbered");
    }

    function applyNumbering(p, numbering) {
        p.SetNumbering(numbering.GetLevel(0));
        // Compact list: no paragraph gap between diagnoses.
        p.SetSpacingBefore(0, false);
        p.SetSpacingAfter(0, false);
        p.SetContextualSpacing(true);
    }

    function indentSubItem(p) {
        p.SetIndLeft(INDENT);
        // Compact sub-items: no paragraph gap between dash lines or before next diagnosis.
        p.SetSpacingBefore(0, false);
        p.SetSpacingAfter(0, false);
        p.SetContextualSpacing(true);
    }

    function formatHeading(p) {
        p.SetBold(true);
        p.SetItalic(false);
        p.SetSpacingAfter(0, false);
        p.SetSpacingBefore(0, false);
    }

    // State for section-sensitive formatting.
    var mode = "";
    var historyNumbering = null;
    var assessmentNumbering = null;
    var planNumbering = null;
    var investigationCount = 0;

    // Re:/patient block: bold Re line plus the next 3 non-empty patient-detail paragraphs.
    var patientBlockRemaining = 0;

    for (var i = 0; i < paragraphs.length; i++) {
        var p = paragraphs[i];
        var text = cleanText(p);

        if (!text) {
            continue;
        }

        // Patient identifier block.
        if (/^Re\s*:/i.test(text)) {
            p.SetBold(true);
            p.SetIndLeft(INDENT);
            patientBlockRemaining = 3;
            continue;
        }

        if (patientBlockRemaining > 0) {
            // Do not consume a salutation as part of patient details.
            if (/^Dear\b/i.test(text)) {
                patientBlockRemaining = 0;
            } else {
                p.SetBold(true);
                p.SetIndLeft(INDENT);
                patientBlockRemaining--;
                continue;
            }
        }

        var h = normHeading(text);

        if (isHeading(text)) {
            formatHeading(p);

            // Reset current section mode then activate the relevant one.
            mode = "";
            investigationCount = 0;

            if (HISTORY_HEADINGS[h]) {
                mode = "history";
                historyNumbering = makeNumbering();
            } else if (h === "INVESTIGATIONS:") {
                mode = "investigations";
            } else if (ASSESSMENT_HEADINGS[h]) {
                mode = "assessment";
                assessmentNumbering = makeNumbering();
            } else if (PLAN_HEADINGS[h]) {
                mode = "plan";
                planNumbering = makeNumbering();
            }
            continue;
        }

        if (isSignatureOrFooter(text)) {
            mode = "";
            continue;
        }

        if (mode === "history") {
            if (/^\s*[-–—•]\s+/.test(text)) {
                // Preserve Lyrebird's dash, simply align the sub-item.
                indentSubItem(p);
            } else {
                applyNumbering(p, historyNumbering);
            }
            continue;
        }

        if (mode === "assessment") {
            applyNumbering(p, assessmentNumbering);
            continue;
        }

        if (mode === "plan") {
            applyNumbering(p, planNumbering);
            continue;
        }

        if (mode === "investigations") {
            if (looksLikeInvestigation(text)) {
                p.SetItalic(true);
                p.SetIndLeft(INDENT);
                p.SetSpacingAfter(0, false);
                p.SetSpacingBefore(0, false);
                investigationCount++;
                continue;
            }

            // Once at least one investigation has been formatted, the first
            // ordinary narrative paragraph ends the investigations block.
            if (investigationCount > 0) {
                mode = "";
            }
        }
    }

    /*
      FINAL SPACING PASS

      1) Past/Medical History:
         - no gaps between numbered diagnoses or dash sub-items
         - ONE visual line after the final history item before the next section

      2) Investigations:
         - preserve the investigation formatting
         - TWO line-spaces in total after the final investigation result before narrative body text (one existing blank paragraph + one line-equivalent of paragraph spacing)

      ONLYOFFICE spacing values are twips:
         240 twips = 12 pt ≈ one text line
         240 twips = 12 pt ≈ one text line
    */

    paragraphs = doc.GetAllParagraphs();

    // ---- Past / Medical History ----
    var inHistory = false;
    var lastHistoryPara = null;

    for (var j = 0; j < paragraphs.length; j++) {
        var hp = paragraphs[j];
        var ht = cleanText(hp);
        if (!ht) continue;
        var hh = normHeading(ht);

        if (HISTORY_HEADINGS[hh]) {
            inHistory = true;
            lastHistoryPara = null;
            continue;
        }

        if (inHistory && isHeading(ht)) {
            // Add exactly one visual line after the last PMH item.
            if (lastHistoryPara) {
                lastHistoryPara.SetSpacingAfter(240, false);
                lastHistoryPara.SetContextualSpacing(false);
            }
            inHistory = false;
        }

        if (inHistory) {
            // Compact every history item and sub-item.
            hp.SetSpacingBefore(0, false);
            hp.SetSpacingAfter(0, false);
            hp.SetContextualSpacing(true);
            lastHistoryPara = hp;
        }
    }

    // If history happens to be the final section in a document.
    if (inHistory && lastHistoryPara) {
        lastHistoryPara.SetSpacingAfter(240, false);
        lastHistoryPara.SetContextualSpacing(false);
    }

    // ---- Investigations ----
    var inInv = false;
    var lastInvPara = null;

    for (var k = 0; k < paragraphs.length; k++) {
        var ip = paragraphs[k];
        var it = cleanText(ip);
        if (!it) continue;
        var ih = normHeading(it);

        if (ih === "INVESTIGATIONS:") {
            inInv = true;
            lastInvPara = null;
            continue;
        }

        if (inInv) {
            if (isHeading(it)) {
                // A new heading means the investigations section has ended.
                if (lastInvPara) {
                    lastInvPara.SetSpacingAfter(240, false);
                    lastInvPara.SetContextualSpacing(false);
                }
                inInv = false;
                continue;
            }

            if (looksLikeInvestigation(it)) {
                lastInvPara = ip;
                continue;
            }

            // First ordinary narrative paragraph after investigation results.
            if (lastInvPara) {
                lastInvPara.SetSpacingAfter(240, false);
                lastInvPara.SetContextualSpacing(false);
            }
            inInv = false;
        }
    }

    if (inInv && lastInvPara) {
        lastInvPara.SetSpacingAfter(240, false);
        lastInvPara.SetContextualSpacing(false);
    }

                return "OK";
            } catch (e) {
                return "ERROR: " + (e && e.message ? e.message : String(e));
            }
        }, true, true, function (result) {
            if (btn) btn.disabled = false;
            if (result && String(result).indexOf("ERROR:") === 0) {
                setStatus(result, true);
            } else {
                setStatus("Formatting complete. Review the letter, then save/send.", false);
            }
        });
    };

    window.Asc.plugin.init = function () {
        setStatus("Ready. Click Format current letter.", false);
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };
})(window, undefined);
