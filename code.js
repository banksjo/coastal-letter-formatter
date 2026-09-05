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

        if (btn) {
            btn.disabled = true;
        }

        setStatus("Formatting current letter…", false);

        window.Asc.plugin.callCommand(function () {

            try {

                var doc = Api.GetDocument();
                var paragraphs = doc.GetAllParagraphs();

                /*
                 * ============================================================
                 * COASTAL LETTER FORMATTER
                 * Version 2.8
                 *
                 * Main principles:
                 *
                 * - Clinical wording is preserved.
                 * - Recognised section headings are bold.
                 * - Medical History is numbered.
                 * - Assessment / Issues is numbered.
                 * - Management Plan is numbered.
                 * - Lines beginning -, –, — or • become proper en-dash
                 *   sub-list items rather than consuming a main list number.
                 * - Investigation results are italicised and indented.
                 * - Closing narrative is not numbered.
                 * - PMH remains compact with one line after the section.
                 * ============================================================
                 */

                /*
                 * ------------------------------------------------------------
                 * SECTION HEADINGS
                 * ------------------------------------------------------------
                 */

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


                /*
                 * ------------------------------------------------------------
                 * BASIC TEXT FUNCTIONS
                 * ------------------------------------------------------------
                 */

                function cleanText(p) {

                    return (
                        p.GetText({
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


                /*
                 * ------------------------------------------------------------
                 * SUBPOINT DETECTION
                 * ------------------------------------------------------------
                 *
                 * Lyrebird may use:
                 *
                 * - text
                 * – text
                 * — text
                 * • text
                 *
                 * These should become genuine sub-list items.
                 */

                function isSubPoint(text) {

                    return /^\s*[-–—•]\s+/.test(text);
                }


                function removeSubPointMarker(text) {

                    return text
                        .replace(/^\s*[-–—•]\s+/, "")
                        .trim();
                }


                /*
                 * ------------------------------------------------------------
                 * CLOSING / SIGNATURE DETECTION
                 * ------------------------------------------------------------
                 */

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


                /*
                 * ------------------------------------------------------------
                 * INVESTIGATION DETECTION
                 * ------------------------------------------------------------
                 */

                function looksLikeInvestigation(text) {

                    var t = text.trim();

                    return /^(RFTs?|PFTs?|Lung function|Spirometry|Plethysmography|DLCO|KCO|FeNO|6[- ]?minute walk|6MWT|Walk test|CT\b|HRCT\b|CTPA\b|PET\b|PET\/CT\b|MRI\b|CXR\b|Chest X[- ]?ray|X[- ]?ray|Ultrasound\b|Echo\b|Echocardiogram\b|ECG\b|Holter\b|ABG\b|VBG\b|Blood\b|FBC\b|UEC\b|LFT\b|CRP\b|ESR\b|IgE\b|IgG\b|Eosinophils?\b|Serology\b|ANA\b|ANCA\b|ENA\b|Biopsy\b|Lung biopsy\b|Histology\b|Pathology\b|Respiratory culture\b|Sputum\b|BAL\b|Bronchoscopy\b|EBUS\b|Sleep study\b|PSG\b|Polysomnography\b|Oximetry\b|CPAP download\b|Device download\b)/i.test(t);
                }


                /*
                 * ------------------------------------------------------------
                 * LIST CREATION
                 * ------------------------------------------------------------
                 */

                function createMainNumbering() {

                    var numbering =
                        doc.CreateNumbering("numbered");

                    var level =
                        numbering.GetLevel(0);

                    /*
                     * Force conventional:
                     *
                     * 1.
                     * 2.
                     * 3.
                     */

                    level.SetTemplateType("1.");

                    level.SetSuff("space");

                    return numbering;
                }


                function createDashNumbering() {

                    var numbering =
                        doc.CreateNumbering("bullet");

                    var level =
                        numbering.GetLevel(0);

                    /*
                     * Use an en dash rather than a round bullet.
                     *
                     * Desired appearance:
                     *
                     *     – CT imaging...
                     *     – Spirometry...
                     */

                    level.SetTemplateType(
                        "bullet",
                        "–"
                    );

                    level.SetSuff("space");

                    return numbering;
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


                function applyDashSubPoint(
                    paragraph,
                    text,
                    numbering
                ) {

                    /*
                     * Remove Lyrebird's literal "-" first.
                     *
                     * The actual visible en dash will then come
                     * from the ONLYOFFICE list definition.
                     */

                    var cleaned =
                        removeSubPointMarker(text);

                    paragraph.SetText(cleaned);

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


                /*
                 * ------------------------------------------------------------
                 * SECTION HEADING FORMAT
                 * ------------------------------------------------------------
                 */

                function formatHeading(p) {

                    p.SetBold(true);

                    p.SetItalic(false);

                    p.SetSpacingBefore(
                        0,
                        false
                    );

                    p.SetSpacingAfter(
                        0,
                        false
                    );
                }


                /*
                 * ------------------------------------------------------------
                 * SPECIAL MEDICAL-HISTORY NORMALISATION
                 * ------------------------------------------------------------
                 *
                 * Letter 9 specifically treats:
                 *
                 * Renal malignancy
                 *     – Partial nephrectomy approximately 3 years ago
                 *
                 * rather than:
                 *
                 * Renal malignancy - partial nephrectomy...
                 *
                 * We only use this narrow rule because globally splitting
                 * every "diagnosis - detail" would incorrectly alter things
                 * such as COPD - emphysematous phenotype and DVT - flight.
                 */

                function shouldSplitRenalMalignancy(text) {

                    return /^Renal malignancy\s*[-–—]\s*partial nephrectomy\b/i
                        .test(text);
                }


                /*
                 * ============================================================
                 * FIRST PASS
                 * ============================================================
                 */

                var mode = "";

                var historyNumbering = null;
                var historyDashNumbering = null;

                var assessmentNumbering = null;
                var assessmentDashNumbering = null;

                var planNumbering = null;
                var planDashNumbering = null;

                var investigationCount = 0;

                var assessmentCount = 0;
                var planCount = 0;

                var patientBlockRemaining = 0;


                for (
                    var i = 0;
                    i < paragraphs.length;
                    i++
                ) {

                    var p =
                        paragraphs[i];

                    var text =
                        cleanText(p);


                    /*
                     * Blank paragraphs.
                     */

                    if (!text) {

                        continue;
                    }


                    /*
                     * Patient identification block.
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


                    var heading =
                        normHeading(text);


                    /*
                     * Section heading encountered.
                     */

                    if (isHeading(text)) {

                        formatHeading(p);

                        mode = "";

                        investigationCount = 0;


                        if (
                            HISTORY_HEADINGS[heading]
                        ) {

                            mode = "history";

                            historyNumbering =
                                createMainNumbering();

                            historyDashNumbering =
                                createDashNumbering();


                        } else if (
                            heading === "INVESTIGATIONS:"
                        ) {

                            mode = "investigations";


                        } else if (
                            ASSESSMENT_HEADINGS[heading]
                        ) {

                            mode = "assessment";

                            assessmentNumbering =
                                createMainNumbering();

                            assessmentDashNumbering =
                                createDashNumbering();

                            assessmentCount = 0;


                        } else if (
                            PLAN_HEADINGS[heading]
                        ) {

                            mode = "plan";

                            planNumbering =
                                createMainNumbering();

                            planDashNumbering =
                                createDashNumbering();

                            planCount = 0;
                        }

                        continue;
                    }


                    /*
                     * Closing text always terminates list modes.
                     */

                    if (
                        isSignatureOrFooter(text) ||
                        isClosingNarrative(text)
                    ) {

                        mode = "";

                        continue;
                    }


                    /*
                     * --------------------------------------------------------
                     * MEDICAL HISTORY
                     * --------------------------------------------------------
                     */

                    if (mode === "history") {

                        /*
                         * Existing explicit subpoint.
                         */

                        if (isSubPoint(text)) {

                            applyDashSubPoint(
                                p,
                                text,
                                historyDashNumbering
                            );

                            continue;
                        }


                        /*
                         * Special Letter-9 transformation:
                         *
                         * Renal malignancy - partial nephrectomy...
                         *
                         * becomes:
                         *
                         * 3. Renal malignancy
                         *      – Partial nephrectomy...
                         */

                        if (
                            shouldSplitRenalMalignancy(text)
                        ) {

                            var parts =
                                text.split(/\s*[-–—]\s*/);

                            var parentText =
                                parts.shift().trim();

                            var childText =
                                parts.join(" - ").trim();


                            /*
                             * Make current paragraph the parent.
                             */

                            p.SetText(parentText);

                            applyMainNumbering(
                                p,
                                historyNumbering
                            );


                            /*
                             * Add child immediately afterwards.
                             */

                            var child =
                                Api.CreateParagraph();

                            child.AddText(childText);

                            child.SetNumbering(
                                historyDashNumbering
                                    .GetLevel(0)
                            );

                            child.SetSpacingBefore(
                                0,
                                false
                            );

                            child.SetSpacingAfter(
                                0,
                                false
                            );

                            child.SetContextualSpacing(
                                true
                            );


                            /*
                             * Insert immediately after parent.
                             */

                            doc.AddElement(
                                i + 1,
                                child
                            );


                            /*
                             * Refresh paragraph array because
                             * document structure changed.
                             */

                            paragraphs =
                                doc.GetAllParagraphs();

                            i++;

                            continue;
                        }


                        /*
                         * Normal medical history item.
                         */

                        applyMainNumbering(
                            p,
                            historyNumbering
                        );

                        continue;
                    }


                    /*
                     * --------------------------------------------------------
                     * ASSESSMENT / ISSUES
                     * --------------------------------------------------------
                     */

                    if (
                        mode === "assessment"
                    ) {

                        if (
                            isSubPoint(text)
                        ) {

                            applyDashSubPoint(
                                p,
                                text,
                                assessmentDashNumbering
                            );

                        } else {

                            applyMainNumbering(
                                p,
                                assessmentNumbering
                            );

                            assessmentCount++;
                        }

                        continue;
                    }


                    /*
                     * --------------------------------------------------------
                     * MANAGEMENT PLAN
                     * --------------------------------------------------------
                     */

                    if (
                        mode === "plan"
                    ) {

                        if (
                            isSubPoint(text)
                        ) {

                            applyDashSubPoint(
                                p,
                                text,
                                planDashNumbering
                            );

                        } else {

                            applyMainNumbering(
                                p,
                                planNumbering
                            );

                            planCount++;
                        }

                        continue;
                    }


                    /*
                     * --------------------------------------------------------
                     * INVESTIGATIONS
                     * --------------------------------------------------------
                     */

                    if (
                        mode === "investigations"
                    ) {

                        if (
                            looksLikeInvestigation(text)
                        ) {

                            p.SetItalic(true);

                            /*
                             * Approximately the indentation used
                             * in your preferred letter style.
                             */

                            p.SetIndLeft(720);

                            p.SetSpacingBefore(
                                0,
                                false
                            );

                            p.SetSpacingAfter(
                                0,
                                false
                            );

                            investigationCount++;

                            continue;
                        }


                        /*
                         * First normal narrative paragraph ends
                         * the investigations section.
                         */

                        if (
                            investigationCount > 0
                        ) {

                            mode = "";
                        }
                    }
                }


                /*
                 * ============================================================
                 * SECOND PASS
                 *
                 * MEDICAL HISTORY SPACING
                 * ============================================================
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

                    var hp =
                        paragraphs[j];

                    var ht =
                        cleanText(hp);


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

                        /*
                         * One line after Medical History.
                         */

                        if (
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
                 * ============================================================
                 * THIRD PASS
                 *
                 * INVESTIGATION SPACING
                 * ============================================================
                 */

                var inInvestigations = false;

                var lastInvestigationParagraph = null;


                for (
                    var k = 0;
                    k < paragraphs.length;
                    k++
                ) {

                    var ip =
                        paragraphs[k];

                    var it =
                        cleanText(ip);


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

                        /*
                         * Another heading ends investigations.
                         */

                        if (
                            isHeading(it)
                        ) {

                            if (
                                lastInvestigationParagraph
                            ) {

                                lastInvestigationParagraph
                                    .SetSpacingAfter(
                                        240,
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


                        /*
                         * Still an investigation.
                         */

                        if (
                            looksLikeInvestigation(it)
                        ) {

                            lastInvestigationParagraph =
                                ip;

                            continue;
                        }


                        /*
                         * First narrative paragraph after investigations.
                         */

                        if (
                            lastInvestigationParagraph
                        ) {

                            lastInvestigationParagraph
                                .SetSpacingAfter(
                                    240,
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
                            240,
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


        }, true, true, function (result) {


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

        });
    };


    /*
     * ================================================================
     * PLUGIN INITIALISATION
     * ================================================================
     */

    window.Asc.plugin.init = function () {

        setStatus(
            "Ready. Click Format current letter.",
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
