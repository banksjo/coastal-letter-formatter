(function (window, undefined) {
    "use strict";

    function setStatus(text, isError) {
        var el = document.getElementById("status");
        if (!el) return;

        el.textContent = text;
        el.style.color = isError ? "#9b1c1c" : "#333333";
    }

    window.formatCoastalLetter = function () {
        var btn = document.getElementById("formatBtn");

        if (btn) btn.disabled = true;

        setStatus("Formatting current letter…", false);

        window.Asc.plugin.callCommand(
            function () {
                try {
                    var doc = Api.GetDocument();
                    var paragraphs = doc.GetAllParagraphs();

                    /*
                     * Coastal Letter Formatter v13
                     *
                     * Stable formatting only.
                     * Does not rewrite clinical wording.
                     */

                    var SUBPOINT_INDENT = 1200;
                    var INVESTIGATION_INDENT = 720;

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
                        "OCCUPATIONAL AND EXPOSURE HISTORY:": true,
                        "EXPOSURE HISTORY:": true,

                        "FAMILY HISTORY:": true,
                        "TRAVEL HISTORY:": true,

                        "AGE-APPROPRIATE MALIGNANCY SCREENING:": true,

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

                    function cleanText(paragraph) {
                        return (
                            paragraph.GetText({
                                "Numbering": false,
                                "Math": true,
                                "NewLineSeparator": "\n",
                                "TabSymbol": "\t"
                            }) || ""
                        )
                        .replace(/\u00A0/g, " ")
                        .trim();
                    }

                    function normHeading(text) {
                        return text
                            .replace(/\s+/g, " ")
                            .trim()
                            .toUpperCase();
                    }

                    function isHeading(text) {
                        return !!SECTION_HEADINGS[
                            normHeading(text)
                        ];
                    }

                    function isSubPoint(text) {
                        return /^\s*[-–—•]\s+/.test(text);
                    }

                    function isSignatureOrFooter(text) {
                        var t = text.trim();

                        return (
                            /^Yours sincerely[,;]?$/i.test(t) ||
                            /^Kind regards[,;]?$/i.test(t) ||
                            /^Regards[,;]?$/i.test(t) ||

                            /^Dr\s+Jonathan\s+Banks\b/i.test(t) ||

                            /^Resp\/Sleep\/Gen Med Phys\./i.test(t) ||

                            /^Coastal Respiratory & Sleep Specialists$/i.test(t) ||

                            /^Prov:/i.test(t) ||

                            /^cc:/i.test(t) ||

                            /^This letter may have been produced with the assistance of AI technology/i.test(t) ||

                            /^Coastal Respiratory and Sleep Services/i.test(t)
                        );
                    }

                    function isClosingNarrative(text) {
                        var t = text.trim();

                        return (
                            /^Thank you\b/i.test(t) ||
                            /^Thankyou\b/i.test(t) ||
                            /^I would be happy\b/i.test(t) ||
                            /^Please do not hesitate\b/i.test(t) ||
                            /^Please contact\b/i.test(t) ||
                            /^With thanks\b/i.test(t) ||
                            /^Many thanks\b/i.test(t)
                        );
                    }

                    function looksLikeInvestigation(text) {
                        var t = text.trim();

                        return /^(RFTs?|PFTs?|Lung function|Spirometry|Plethysmography|DLCO|KCO|FeNO|6[- ]?minute walk|6MWT|Walk test|CT\b|HRCT\b|CTPA\b|PET\b|PET\/CT\b|MRI\b|CXR\b|Chest X[- ]?ray|X[- ]?ray|Ultrasound\b|Echo\b|Echocardiogram\b|ECG\b|Holter\b|ABG\b|VBG\b|Blood\b|FBC\b|UEC\b|LFT\b|CRP\b|ESR\b|IgE\b|IgG\b|Eosinophils?\b|Serology\b|ANA\b|ANCA\b|ENA\b|Biopsy\b|Lung biopsy\b|Histology\b|Pathology\b|Respiratory culture\b|Sputum\b|BAL\b|Bronchoscopy\b|EBUS\b|Sleep study\b|PSG\b|Polysomnography\b|Oximetry\b|CPAP download\b|Device download\b)/i.test(t);
                    }

                    function makeNumbering() {
                        return doc.CreateNumbering("numbered");
                    }

                    function applyMainNumbering(
                        paragraph,
                        numbering
                    ) {
                        paragraph.SetNumbering(
                            numbering.GetLevel(0)
                        );

                        paragraph.SetSpacingBefore(
                            0,
                            false
                        );

                        paragraph.SetSpacingAfter(
                            0,
                            false
                        );

                        paragraph.SetContextualSpacing(
                            true
                        );
                    }

                    function formatSubPoint(paragraph) {
                        /*
                         * More deeply indented than the main numbered text.
                         * Keep Lyrebird's existing dash/en-dash marker.
                         */

                        paragraph.SetIndLeft(
                            SUBPOINT_INDENT
                        );

                        paragraph.SetSpacingBefore(
                            0,
                            false
                        );

                        paragraph.SetSpacingAfter(
                            0,
                            false
                        );

                        paragraph.SetContextualSpacing(
                            true
                        );
                    }

                    function formatHeading(paragraph) {
                        paragraph.SetBold(true);
                        paragraph.SetItalic(false);

                        paragraph.SetSpacingBefore(
                            0,
                            false
                        );

                        paragraph.SetSpacingAfter(
                            0,
                            false
                        );
                    }

                    function formatInvestigation(paragraph) {
                        paragraph.SetItalic(true);

                        paragraph.SetIndLeft(
                            INVESTIGATION_INDENT
                        );

                        paragraph.SetSpacingBefore(
                            0,
                            false
                        );

                        paragraph.SetSpacingAfter(
                            0,
                            false
                        );
                    }

                    /*
                     * =====================================================
                     * FIRST PASS
                     * =====================================================
                     */

                    var mode = "";

                    var historyNumbering = null;
                    var assessmentNumbering = null;
                    var planNumbering = null;

                    var investigationCount = 0;

                    var patientBlockRemaining = 0;

                    for (
                        var i = 0;
                        i < paragraphs.length;
                        i++
                    ) {
                        var p = paragraphs[i];
                        var text = cleanText(p);

                        if (!text) {
                            continue;
                        }

                        /*
                         * Patient identifier.
                         */

                        if (/^Re\s*:/i.test(text)) {
                            p.SetBold(true);

                            patientBlockRemaining = 3;

                            continue;
                        }

                        if (patientBlockRemaining > 0) {
                            if (/^Dear\b/i.test(text)) {
                                patientBlockRemaining = 0;
                            } else {
                                p.SetBold(true);

                                patientBlockRemaining--;

                                continue;
                            }
                        }

                        /*
                         * Section headings.
                         */

                        var heading = normHeading(text);

                        if (isHeading(text)) {
                            formatHeading(p);

                            mode = "";
                            investigationCount = 0;

                            if (
                                HISTORY_HEADINGS[heading]
                            ) {
                                mode = "history";

                                historyNumbering =
                                    makeNumbering();

                            } else if (
                                heading === "INVESTIGATIONS:"
                            ) {
                                mode = "investigations";

                            } else if (
                                ASSESSMENT_HEADINGS[heading]
                            ) {
                                mode = "assessment";

                                assessmentNumbering =
                                    makeNumbering();

                            } else if (
                                PLAN_HEADINGS[heading]
                            ) {
                                mode = "plan";

                                planNumbering =
                                    makeNumbering();
                            }

                            continue;
                        }

                        /*
                         * Closing narrative / signature.
                         */

                        if (
                            isClosingNarrative(text) ||
                            isSignatureOrFooter(text)
                        ) {
                            mode = "";

                            continue;
                        }

                        /*
                         * Medical History.
                         */

                        if (mode === "history") {
                            if (isSubPoint(text)) {
                                formatSubPoint(p);
                            } else {
                                applyMainNumbering(
                                    p,
                                    historyNumbering
                                );
                            }

                            continue;
                        }

                        /*
                         * Assessment and Issues.
                         */

                        if (mode === "assessment") {
                            if (isSubPoint(text)) {
                                formatSubPoint(p);
                            } else {
                                applyMainNumbering(
                                    p,
                                    assessmentNumbering
                                );
                            }

                            continue;
                        }

                        /*
                         * Management Plan.
                         */

                        if (mode === "plan") {
                            if (isSubPoint(text)) {
                                formatSubPoint(p);
                            } else {
                                applyMainNumbering(
                                    p,
                                    planNumbering
                                );
                            }

                            continue;
                        }

                        /*
                         * Investigations.
                         */

                        if (mode === "investigations") {
                            if (
                                looksLikeInvestigation(text)
                            ) {
                                formatInvestigation(p);

                                investigationCount++;

                                continue;
                            }

                            if (
                                investigationCount > 0
                            ) {
                                mode = "";
                            }
                        }
                    }


                    /*
                     * =====================================================
                     * MEDICAL HISTORY SPACING
                     * =====================================================
                     */

                    paragraphs =
                        doc.GetAllParagraphs();

                    var inHistory = false;
                    var lastHistoryParagraph = null;

                    for (
                        var j = 0;
                        j < paragraphs.length;
                        j++
                    ) {
                        var hp = paragraphs[j];
                        var ht = cleanText(hp);

                        if (!ht) {
                            continue;
                        }

                        var hh =
                            normHeading(ht);

                        if (
                            HISTORY_HEADINGS[hh]
                        ) {
                            inHistory = true;

                            lastHistoryParagraph = null;

                            continue;
                        }

                        if (
                            inHistory &&
                            isHeading(ht)
                        ) {
                            if (
                                lastHistoryParagraph
                            ) {
                                /*
                                 * One line after Medical History.
                                 */

                                lastHistoryParagraph
                                    .SetSpacingAfter(
                                        240,
                                        false
                                    );

                                lastHistoryParagraph
                                    .SetContextualSpacing(
                                        false
                                    );
                            }

                            inHistory = false;
                        }

                        if (inHistory) {
                            hp.SetSpacingBefore(
                                0,
                                false
                            );

                            hp.SetSpacingAfter(
                                0,
                                false
                            );

                            hp.SetContextualSpacing(
                                true
                            );

                            lastHistoryParagraph = hp;
                        }
                    }

                    if (
                        inHistory &&
                        lastHistoryParagraph
                    ) {
                        lastHistoryParagraph
                            .SetSpacingAfter(
                                240,
                                false
                            );

                        lastHistoryParagraph
                            .SetContextualSpacing(
                                false
                            );
                    }


                    /*
                     * =====================================================
                     * INVESTIGATION SPACING
                     * =====================================================
                     */

                    var inInvestigations = false;
                    var lastInvestigationParagraph = null;

                    for (
                        var k = 0;
                        k < paragraphs.length;
                        k++
                    ) {
                        var ip = paragraphs[k];
                        var it = cleanText(ip);

                        if (!it) {
                            continue;
                        }

                        var ih =
                            normHeading(it);

                        if (
                            ih === "INVESTIGATIONS:"
                        ) {
                            inInvestigations = true;

                            lastInvestigationParagraph = null;

                            continue;
                        }

                        if (
                            inInvestigations
                        ) {
                            if (
                                isHeading(it)
                            ) {
                                if (
                                    lastInvestigationParagraph
                                ) {
                                    lastInvestigationParagraph
                                        .SetSpacingAfter(
                                            120,
                                            false
                                        );

                                    lastInvestigationParagraph
                                        .SetContextualSpacing(
                                            false
                                        );
                                }

                                inInvestigations = false;

                                continue;
                            }

                            if (
                                looksLikeInvestigation(it)
                            ) {
                                lastInvestigationParagraph =
                                    ip;

                                continue;
                            }

                            if (
                                lastInvestigationParagraph
                            ) {
                                lastInvestigationParagraph
                                    .SetSpacingAfter(
                                        120,
                                        false
                                    );

                                lastInvestigationParagraph
                                    .SetContextualSpacing(
                                        false
                                    );
                            }

                            inInvestigations = false;
                        }
                    }

                    if (
                        inInvestigations &&
                        lastInvestigationParagraph
                    ) {
                        lastInvestigationParagraph
                            .SetSpacingAfter(
                                120,
                                false
                            );

                        lastInvestigationParagraph
                            .SetContextualSpacing(
                                false
                            );
                    }

                    return "OK";

                } catch (e) {
                    return (
                        "ERROR: " +
                        (
                            e && e.message
                                ? e.message
                                : String(e)
                        )
                    );
                }
            },

            true,
            true,

            function (result) {
                if (btn) {
                    btn.disabled = false;
                }

                if (
                    result &&
                    String(result).indexOf(
                        "ERROR:"
                    ) === 0
                ) {
                    setStatus(
                        result,
                        true
                    );
                } else {
                    setStatus(
                        "Formatting complete. Review the letter, then save/send.",
                        false
                    );
                }
            }
        );
    };


    window.Asc.plugin.init = function () {
        setStatus(
            "Coastal Formatter v14 loaded — click Format current letter.",
            false
        );
    };


    window.Asc.plugin.button = function (id) {
        this.executeCommand(
            "close",
            ""
        );
    };

})(window, undefined);
