(function (window, undefined) {
    "use strict";

    function setStatus(text, isError) {
        var el = document.getElementById("status");

        if (!el) {
            return;
        }

        el.textContent = text;
        el.style.color = isError ? "#9b1c1c" : "#333333";
    }


    window.formatCoastalLetter = function () {

        var btn = document.getElementById("formatBtn");

        if (btn) {
            btn.disabled = true;
        }

        setStatus("Formatting current letter…", false);


        window.Asc.plugin.callCommand(
            function () {

                try {

                    var doc = Api.GetDocument();
                    var paragraphs = doc.GetAllParagraphs();


                    /*
                     * =====================================================
                     * COASTAL LETTER FORMATTER v15
                     * =====================================================
                     *
                     * PURPOSE
                     *
                     * - Preserve clinical wording.
                     * - Bold recognised section headings.
                     * - Number Medical History.
                     * - Number Assessment / Issues.
                     * - Number Management Plan.
                     * - Preserve explicit dash/bullet subpoints.
                     * - Give subpoints a modest hanging indent.
                     * - Italicise and indent Investigations.
                     * - Keep internal list spacing compact.
                     * - Prevent closing narrative from being numbered.
                     *
                     * IMPORTANT
                     *
                     * The formatter does NOT create or rewrite clinical
                     * sentences. A subpoint must already be a separate
                     * paragraph beginning with:
                     *
                     *   -
                     *   –
                     *   —
                     *   •
                     *
                     * =====================================================
                     */


                    /*
                     * -----------------------------------------------------
                     * INDENTATION
                     * -----------------------------------------------------
                     *
                     * Twips:
                     * 1440 twips = 1 inch.
                     *
                     * Subpoints use:
                     *
                     * left indent       720 = 0.5 inch
                     * first-line indent -240
                     *
                     * This produces:
                     *
                     *     - Subpoint text begins here...
                     *       continuation aligns under text
                     *
                     * rather than pushing the whole paragraph too far
                     * to the right.
                     */

                    var SUBPOINT_LEFT_INDENT = 720;
                    var SUBPOINT_FIRST_LINE = -240;

                    var INVESTIGATION_INDENT = 720;


                    /*
                     * =====================================================
                     * SECTION HEADINGS
                     * =====================================================
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
                     * =====================================================
                     * TEXT HELPERS
                     * =====================================================
                     */

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

                        /*
                         * Explicit subpoints only.
                         *
                         * Examples:
                         *
                         * - CT imaging demonstrates...
                         * – CT imaging demonstrates...
                         * — CT imaging demonstrates...
                         * • CT imaging demonstrates...
                         */

                        return /^\s*[-–—•]\s+/.test(text);
                    }


                    /*
                     * =====================================================
                     * CLOSING / FOOTER DETECTION
                     * =====================================================
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
                     * =====================================================
                     * INVESTIGATION DETECTION
                     * =====================================================
                     */

                    function looksLikeInvestigation(text) {

                        var t = text.trim();

                        return /^(RFTs?|PFTs?|Lung function|Spirometry|Plethysmography|DLCO|KCO|FeNO|6[- ]?minute walk|6MWT|Walk test|CT\b|HRCT\b|CTPA\b|PET\b|PET\/CT\b|MRI\b|CXR\b|Chest X[- ]?ray|X[- ]?ray|Ultrasound\b|Echo\b|Echocardiogram\b|ECG\b|Holter\b|ABG\b|VBG\b|Blood\b|FBC\b|UEC\b|LFT\b|CRP\b|ESR\b|IgE\b|IgG\b|Eosinophils?\b|Serology\b|ANA\b|ANCA\b|ENA\b|Biopsy\b|Lung biopsy\b|Histology\b|Pathology\b|Respiratory culture\b|Sputum\b|BAL\b|Bronchoscopy\b|EBUS\b|Sleep study\b|PSG\b|Polysomnography\b|Oximetry\b|CPAP download\b|Device download\b)/i.test(t);
                    }


                    /*
                     * =====================================================
                     * NUMBERING
                     * =====================================================
                     */

                    function makeNumbering() {

                        /*
                         * This is the numbering implementation that has
                         * already proven reliable in Xestro.
                         */

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


                    /*
                     * =====================================================
                     * SUBPOINT FORMATTING
                     * =====================================================
                     */

                    function formatSubPoint(paragraph) {

                        /*
                         * Do NOT remove or replace the dash.
                         *
                         * Preserve whatever Lyrebird supplied.
                         *
                         * Use a modest hanging indent so wrapped text
                         * aligns neatly under the subpoint text.
                         */

                        paragraph.SetIndLeft(
                            SUBPOINT_LEFT_INDENT
                        );

                        paragraph.SetIndFirstLine(
                            SUBPOINT_FIRST_LINE
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
                     * =====================================================
                     * HEADING FORMATTING
                     * =====================================================
                     */

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


                    /*
                     * =====================================================
                     * INVESTIGATION FORMATTING
                     * =====================================================
                     */

                    function formatInvestigation(paragraph) {

                        /*
                         * This matches your preferred Letter 9:
                         * investigations italic + indented.
                         */

                        paragraph.SetItalic(true);

                        paragraph.SetIndLeft(
                            INVESTIGATION_INDENT
                        );

                        paragraph.SetIndFirstLine(
                            0
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


                        /*
                         * Ignore blank paragraphs.
                         */

                        if (!text) {

                            continue;
                        }


                        /*
                         * -------------------------------------------------
                         * PATIENT IDENTIFICATION BLOCK
                         * -------------------------------------------------
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
                         * -------------------------------------------------
                         * SECTION HEADINGS
                         * -------------------------------------------------
                         */

                        var heading =
                            normHeading(text);


                        if (isHeading(text)) {

                            formatHeading(p);

                            /*
                             * A recognised heading ends the previous
                             * section mode.
                             */

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
                         * -------------------------------------------------
                         * CLOSING TEXT
                         * -------------------------------------------------
                         */

                        if (
                            isClosingNarrative(text) ||
                            isSignatureOrFooter(text)
                        ) {

                            mode = "";

                            continue;
                        }


                        /*
                         * -------------------------------------------------
                         * MEDICAL / PAST MEDICAL HISTORY
                         * -------------------------------------------------
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
                         * -------------------------------------------------
                         * ASSESSMENT / ISSUES
                         * -------------------------------------------------
                         */

                        if (mode === "assessment") {

                            if (isSubPoint(text)) {

                                /*
                                 * Explicit dash paragraph:
                                 * subordinate to preceding numbered issue.
                                 */

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
                         * -------------------------------------------------
                         * MANAGEMENT PLAN
                         * -------------------------------------------------
                         */

                        if (mode === "plan") {

                            if (isSubPoint(text)) {

                                /*
                                 * Explicit dash paragraph:
                                 * subordinate to preceding numbered plan.
                                 */

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
                         * -------------------------------------------------
                         * INVESTIGATIONS
                         * -------------------------------------------------
                         */

                        if (mode === "investigations") {

                            if (
                                looksLikeInvestigation(text)
                            ) {

                                formatInvestigation(p);

                                investigationCount++;

                                continue;
                            }


                            /*
                             * First ordinary narrative paragraph after
                             * investigations ends investigation mode.
                             */

                            if (
                                investigationCount > 0
                            ) {

                                mode = "";
                            }
                        }
                    }


                    /*
                     * =====================================================
                     * SECOND PASS
                     *
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
                             * One visual line after the Medical History
                             * section.
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

                            lastHistoryParagraph =
                                hp;
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
                     * THIRD PASS
                     *
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


                        if (inInvestigations) {

                            /*
                             * Another recognised heading ends the
                             * investigation section.
                             */

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


                            /*
                             * Still inside investigation results.
                             */

                            if (
                                looksLikeInvestigation(it)
                            ) {

                                lastInvestigationParagraph =
                                    ip;

                                continue;
                            }


                            /*
                             * Narrative begins.
                             */

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


    /*
     * ================================================================
     * PLUGIN INITIALISATION
     * ================================================================
     */

    window.Asc.plugin.init = function () {

        setStatus(
            "Coastal Formatter v15 loaded — click Format current letter.",
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
