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
                     * COASTAL LETTER FORMATTER v18
                     * =====================================================
                     *
                     * IMPORTANT CHANGE:
                     *
                     * INVESTIGATIONS ARE NOW RECOGNISED STRUCTURALLY.
                     *
                     * We do NOT try to maintain a list of every possible
                     * investigation in medicine.
                     *
                     * Once an INVESTIGATIONS heading is found:
                     *
                     * - every subsequent non-empty paragraph is formatted
                     *   as an investigation
                     * - one blank paragraph between tests is allowed
                     * - two consecutive blank paragraphs end the section
                     * - another recognised heading also ends the section
                     *
                     * Therefore this automatically handles:
                     *
                     * pathology
                     * blood tests
                     * imaging
                     * spirometry
                     * full lung function
                     * PSG
                     * CPAP titration
                     * device downloads
                     * CPET
                     * ABG
                     * VBG
                     * ECG
                     * echo
                     * Holter
                     * PET
                     * CT
                     * MRI
                     * ultrasound
                     * bronchoscopy
                     * EBUS
                     * biopsy
                     * histopathology
                     * microbiology
                     * immunology
                     * genetics
                     * serology
                     * iron studies
                     * etc.
                     *
                     * No clinical wording is changed.
                     * =====================================================
                     */


                    var SUBPOINT_LEFT_INDENT = 1200;
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
                        "INVESTIGATION RESULTS:": true,
                        "RELEVANT INVESTIGATIONS:": true,
                        "RESULTS:": true,

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


                    var INVESTIGATION_HEADINGS = {
                        "INVESTIGATIONS:": true,
                        "INVESTIGATION RESULTS:": true,
                        "RELEVANT INVESTIGATIONS:": true,
                        "RESULTS:": true
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
                     * HELPERS
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


                    /*
                     * =====================================================
                     * NUMBERING
                     * =====================================================
                     */

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


                    /*
                     * =====================================================
                     * SUBPOINT FORMAT
                     * =====================================================
                     */

                    function formatSubPoint(paragraph) {

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
                     * HEADING FORMAT
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
                     * INVESTIGATION FORMAT
                     * =====================================================
                     */

                    function formatInvestigation(paragraph) {

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

                    /*
                     * Number of consecutive blank paragraphs encountered
                     * while inside Investigations.
                     */

                    var investigationBlankCount = 0;

                    var patientBlockRemaining = 0;


                    for (
                        var i = 0;
                        i < paragraphs.length;
                        i++
                    ) {

                        var p = paragraphs[i];
                        var text = cleanText(p);


                        /*
                         * =================================================
                         * BLANK PARAGRAPHS
                         * =================================================
                         */

                        if (!text) {

                            if (
                                mode === "investigations"
                            ) {

                                investigationBlankCount++;

                                /*
                                 * One blank paragraph is normal between
                                 * separate investigations.
                                 *
                                 * Two consecutive blanks mean the
                                 * investigations section is finished.
                                 */

                                if (
                                    investigationBlankCount >= 2
                                ) {

                                    mode = "";
                                }
                            }

                            continue;
                        }


                        /*
                         * Any non-empty paragraph resets the blank counter.
                         */

                        if (
                            mode === "investigations"
                        ) {

                            investigationBlankCount = 0;
                        }


                        /*
                         * =================================================
                         * PATIENT BLOCK
                         * =================================================
                         */

                        if (/^Re\s*:/i.test(text)) {

                            p.SetBold(true);

                            patientBlockRemaining = 3;

                            continue;
                        }


                        if (
                            patientBlockRemaining > 0
                        ) {

                            if (
                                /^Dear\b/i.test(text)
                            ) {

                                patientBlockRemaining = 0;

                            } else {

                                p.SetBold(true);

                                patientBlockRemaining--;

                                continue;
                            }
                        }


                        /*
                         * =================================================
                         * SECTION HEADING
                         * =================================================
                         */

                        var heading =
                            normHeading(text);


                        if (
                            isHeading(text)
                        ) {

                            formatHeading(p);

                            mode = "";

                            investigationBlankCount = 0;


                            if (
                                HISTORY_HEADINGS[heading]
                            ) {

                                mode = "history";

                                historyNumbering =
                                    makeNumbering();


                            } else if (
                                INVESTIGATION_HEADINGS[heading]
                            ) {

                                /*
                                 * Critical v18 behaviour:
                                 *
                                 * EVERYTHING after this heading is treated
                                 * as an investigation until the section
                                 * structurally ends.
                                 */

                                mode = "investigations";

                                investigationBlankCount = 0;


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
                         * =================================================
                         * CLOSING TEXT
                         * =================================================
                         */

                        if (
                            isClosingNarrative(text) ||
                            isSignatureOrFooter(text)
                        ) {

                            mode = "";

                            continue;
                        }


                        /*
                         * =================================================
                         * INVESTIGATIONS
                         * =================================================
                         *
                         * This deliberately occurs BEFORE the other
                         * section handlers.
                         *
                         * No test-name recognition is required.
                         */

                        if (
                            mode === "investigations"
                        ) {

                            formatInvestigation(p);

                            continue;
                        }


                        /*
                         * =================================================
                         * MEDICAL HISTORY
                         * =================================================
                         */

                        if (
                            mode === "history"
                        ) {

                            if (
                                isSubPoint(text)
                            ) {

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
                         * =================================================
                         * ASSESSMENT / ISSUES
                         * =================================================
                         */

                        if (
                            mode === "assessment"
                        ) {

                            if (
                                isSubPoint(text)
                            ) {

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
                         * =================================================
                         * MANAGEMENT PLAN
                         * =================================================
                         */

                        if (
                            mode === "plan"
                        ) {

                            if (
                                isSubPoint(text)
                            ) {

                                formatSubPoint(p);

                            } else {

                                applyMainNumbering(
                                    p,
                                    planNumbering
                                );
                            }

                            continue;
                        }
                    }


                    /*
                     * =====================================================
                     * MEDICAL HISTORY SPACING PASS
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


                        if (
                            inHistory
                        ) {

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
                     * INVESTIGATION SPACING PASS
                     * =====================================================
                     *
                     * Again: structural, not based on test names.
                     */

                    var inInvestigations = false;

                    var investigationBlanks = 0;

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

                            if (
                                inInvestigations
                            ) {

                                investigationBlanks++;

                                if (
                                    investigationBlanks >= 2
                                ) {

                                    /*
                                     * End of investigation block.
                                     */

                                    if (
                                        lastInvestigationParagraph
                                    ) {

                                        lastInvestigationParagraph
                                            .SetSpacingAfter(
                                                0,
                                                false
                                            );
                                    }

                                    inInvestigations = false;
                                }
                            }

                            continue;
                        }


                        var ih =
                            normHeading(it);


                        if (
                            INVESTIGATION_HEADINGS[ih]
                        ) {

                            inInvestigations = true;

                            investigationBlanks = 0;

                            lastInvestigationParagraph = null;

                            continue;
                        }


                        if (
                            inInvestigations
                        ) {

                            /*
                             * Another heading always ends investigations.
                             */

                            if (
                                isHeading(it)
                            ) {

                                inInvestigations = false;

                                continue;
                            }


                            /*
                             * Any non-empty paragraph before the structural
                             * end is an investigation.
                             */

                            investigationBlanks = 0;

                            ip.SetSpacingBefore(
                                0,
                                false
                            );

                            ip.SetSpacingAfter(
                                0,
                                false
                            );

                            ip.SetContextualSpacing(
                                true
                            );

                            lastInvestigationParagraph =
                                ip;
                        }
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
            "Coastal Formatter v18 loaded — click Format current letter.",
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
