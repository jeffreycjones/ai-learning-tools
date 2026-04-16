/* ============================================================
 * TLI Learning Design Suite — AV Script Export Module
 * ------------------------------------------------------------
 * Overrides downloadWord() for the Lecture Video Script Writer
 * to produce a .docx matching the Mays TLI A/V Script
 * template (Template-AV-Script.docx).
 *
 * Template structure:
 *   Table 0 — Metadata (Title, Duration, Module, Date)
 *   Table 1 — Two-column A/V table:
 *     Col 1 (6020 dxa): Dialogue (Audio)
 *     Col 2 (4040 dxa): Graphics (Video)
 *     Section dividers: full-span gray rows
 *     Header row: maroon #500000 fill
 *
 * Depends on tli-download.js being loaded first (provides
 * JSZip loader and utility functions).
 * ============================================================ */
(function () {
    'use strict';

    const AV_TEMPLATE_URL = 'assets/Template-AV-Script.docx';

    /* ── XML helpers ──────────────────────────────── */

    function xe(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    /* ── OOXML building blocks ────────────────────── */

    const BORDER = '<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                 + '<w:left w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                 + '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
                 + '<w:right w:val="single" w:sz="8" w:space="0" w:color="000000"/>';

    const CELL_MAR = '<w:tcMar>'
                   + '<w:top w:w="100" w:type="dxa"/>'
                   + '<w:left w:w="100" w:type="dxa"/>'
                   + '<w:bottom w:w="100" w:type="dxa"/>'
                   + '<w:right w:w="100" w:type="dxa"/>'
                   + '</w:tcMar>';

    const RUN_FONT = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>';
    const RUN_BODY = '<w:rPr>' + RUN_FONT + '</w:rPr>';
    const RUN_BOLD = '<w:rPr>' + RUN_FONT + '<w:b/><w:bCs/></w:rPr>';
    const RUN_WHITE_BOLD = '<w:rPr>' + RUN_FONT + '<w:b/><w:bCs/><w:color w:val="FFFFFF"/></w:rPr>';
    const RUN_ITALIC = '<w:rPr>' + RUN_FONT + '<w:i/><w:iCs/><w:color w:val="808080"/></w:rPr>';

    const PARA_TIGHT = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>';

    function textRun(text, rPr) {
        return `<w:r>${rPr || RUN_BODY}<w:t xml:space="preserve">${xe(text)}</w:t></w:r>`;
    }

    function para(content, pPr) {
        return `<w:p>${pPr || PARA_TIGHT}${content}</w:p>`;
    }

    /* ── Metadata table (Table 0) ─────────────────── */

    function buildMetadataTable(title, duration, moduleSection, revDate) {
        const labelCell = (text, w) => `<w:tc><w:tcPr>`
            + `<w:tcW w:w="${w}" w:type="dxa"/>`
            + `<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>`
            + CELL_MAR
            + `</w:tcPr>${para(textRun(text, RUN_BOLD))}</w:tc>`;

        const valueCell = (text, w) => `<w:tc><w:tcPr>`
            + `<w:tcW w:w="${w}" w:type="dxa"/>`
            + `<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>`
            + CELL_MAR
            + `</w:tcPr>${para(textRun(text))}</w:tc>`;

        const tblPr = '<w:tblPr>'
            + '<w:tblW w:w="10060" w:type="dxa"/>'
            + '<w:tblLayout w:type="fixed"/>'
            + '</w:tblPr>'
            + '<w:tblGrid><w:gridCol w:w="2155"/><w:gridCol w:w="2875"/><w:gridCol w:w="1800"/><w:gridCol w:w="3230"/></w:tblGrid>';

        const row0 = `<w:tr>${labelCell('Title', 2155)}${valueCell(title, 2875)}${labelCell('Duration', 1800)}${valueCell(duration, 3230)}</w:tr>`;
        const row1 = `<w:tr>${labelCell('Module & Section', 2155)}${valueCell(moduleSection, 2875)}${labelCell('Revision Date', 1800)}${valueCell(revDate, 3230)}</w:tr>`;

        return `<w:tbl>${tblPr}${row0}${row1}</w:tbl>`;
    }

    /* ── A/V Script table (Table 1) ───────────────── */

    function headerRow() {
        const cellPr = (w) => `<w:tcPr>`
            + `<w:tcW w:w="${w}" w:type="dxa"/>`
            + `<w:tcBorders>${BORDER}</w:tcBorders>`
            + `<w:shd w:val="clear" w:color="auto" w:fill="500000"/>`
            + CELL_MAR
            + `</w:tcPr>`;
        return `<w:tr><w:trPr><w:trHeight w:val="461"/></w:trPr>`
            + `<w:tc>${cellPr(6020)}${para(textRun('Dialogue (Audio)', RUN_WHITE_BOLD))}</w:tc>`
            + `<w:tc>${cellPr(4040)}${para(textRun('Graphics (Video)', RUN_WHITE_BOLD))}</w:tc>`
            + `</w:tr>`;
    }

    function sectionDividerRow(label) {
        const cellPr = `<w:tcPr>`
            + `<w:tcW w:w="0" w:type="auto"/>`
            + `<w:gridSpan w:val="2"/>`
            + `<w:tcBorders>${BORDER}</w:tcBorders>`
            + `<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>`
            + CELL_MAR
            + `</w:tcPr>`;
        return `<w:tr><w:trPr><w:trHeight w:val="420"/></w:trPr>`
            + `<w:tc>${cellPr}${para(textRun(label, RUN_BOLD))}</w:tc>`
            + `</w:tr>`;
    }

    function contentRow(dialogueParas, graphicsParas) {
        const cellPr = (w) => `<w:tcPr>`
            + `<w:tcW w:w="${w}" w:type="dxa"/>`
            + `<w:tcBorders>${BORDER}</w:tcBorders>`
            + CELL_MAR
            + `</w:tcPr>`;

        const dContent = dialogueParas.length
            ? dialogueParas.join('')
            : para('');
        const gContent = graphicsParas.length
            ? graphicsParas.join('')
            : para('');

        return `<w:tr><w:trPr><w:trHeight w:val="420"/></w:trPr>`
            + `<w:tc>${cellPr(6020)}${dContent}</w:tc>`
            + `<w:tc>${cellPr(4040)}${gContent}</w:tc>`
            + `</w:tr>`;
    }

    function buildAVTable(sections) {
        const tblPr = '<w:tblPr>'
            + '<w:tblW w:w="10060" w:type="dxa"/>'
            + '<w:tblLayout w:type="fixed"/>'
            + '</w:tblPr>'
            + '<w:tblGrid><w:gridCol w:w="6020"/><w:gridCol w:w="4040"/></w:tblGrid>';

        let rows = headerRow();

        // Opening boilerplate
        rows += sectionDividerRow('Opening: Flex Online Course Specific');
        rows += contentRow(
            [para('')],
            [
                para(textRun('[Intro Graphic: Flex Online, Mays Business School, Course Title, Video Title]', RUN_ITALIC)),
                para(textRun('[Intro Music: leads into the video speaker audio channel]', RUN_ITALIC))
            ]
        );

        // Script sections from AI output
        for (const section of sections) {
            rows += sectionDividerRow(section.label);
            rows += contentRow(
                section.dialogue.map(d => para(textRun(d))),
                section.graphics.map(g => para(textRun(g, RUN_ITALIC)))
            );
        }

        // Closing boilerplate
        rows += sectionDividerRow('Closing: Flex Online Course Specific');
        rows += contentRow(
            [para('')],
            [
                para(textRun('[Outro Graphic: Branding of Texas A&M University]', RUN_ITALIC)),
                para(textRun('[Outro Music]', RUN_ITALIC))
            ]
        );

        return `<w:tbl>${tblPr}${rows}</w:tbl>`;
    }

    /* ── Raw text parser ──────────────────────────── */

    // Visual / production cues go to Graphics column (no end anchor -- tolerate trailing spaces)
    const GRAPHICS_CUE_RE = /^\(Visual:.*\)|^\[SLIDE:.*\]|^\[B-ROLL:.*\]|^\[TRANSITION:.*\]/i;
    // Timing markers
    const TIMING_RE = /^\*?Approx\.?\s+[\d:]+\*?\s*$/i;
    // Pacing cues stay in dialogue
    const PACE_CUE_RE = /^\[PAUSE.*\]|^\[EMPHASIS.*\]/i;
    // Instructor review flag
    const REVIEW_RE = /^\[INSTRUCTOR REVIEW:.*\]/i;

    // Keywords that indicate a structural section header
    const SECTION_KEYWORDS = /hook|cold open|section\s*\d|recap|summary|call to action|learning obj|introduction|conclusion|opening/i;
    // Keywords for production metadata (not script sections)
    const PRODUCTION_SUMMARY_RE = /production summary/i;
    const PRODUCTION_NOTES_RE = /production notes/i;

    // Strip leading/trailing markdown formatting from a line
    function stripMd(s) {
        return s.replace(/^#+\s*/, '').replace(/^\*+/, '').replace(/\*+$/, '').replace(/^_+/, '').replace(/_+$/, '').trim();
    }

    // Normalize a line by stripping numbered prefixes, bold/italic markers, and
    // horizontal rules so pattern matching works on any combination of formatting.
    function normalizeLine(s) {
        // Strip leading numbered prefix: "1. ", "2) ", etc.
        s = s.replace(/^\d+[\.\)]\s*/, '');
        // Strip bold/italic markers
        s = s.replace(/^\*{1,3}/, '').replace(/\*{1,3}$/, '');
        // Strip leading/trailing underscores
        s = s.replace(/^_+/, '').replace(/_+$/, '');
        return s.trim();
    }

    // Try to detect if a line is a section header. Returns the label text or null.
    function detectSectionHeader(trimmed) {
        // Normalize: strip "1. **" prefix and "**" suffix so all combos work
        const norm = normalizeLine(trimmed);

        // 1. Bracket headers: [HOOK / COLD OPEN], [SECTION 1: TITLE]
        let m = norm.match(/^\[([^\]]+)\]\s*$/);
        if (m) {
            const inner = m[1].trim();
            if (SECTION_KEYWORDS.test(inner) || PRODUCTION_SUMMARY_RE.test(inner) || PRODUCTION_NOTES_RE.test(inner)) {
                return inner;
            }
        }

        // 2. Markdown headers: ## Hook / Cold Open, ### Section 1: The Three Pillars
        m = trimmed.match(/^#{1,4}\s+(.+)$/);
        if (m) {
            const text = normalizeLine(m[1]).replace(/[\[\]]/g, '').trim();
            if (SECTION_KEYWORDS.test(text) || PRODUCTION_SUMMARY_RE.test(text) || PRODUCTION_NOTES_RE.test(text)) {
                return text;
            }
        }

        // 3. Plain text (after normalization) that matches section keywords
        //    Catches: "Hook / Cold Open", "Section 1: Title", "Recap / Summary"
        const plain = norm.replace(/[\[\]]/g, '').trim();
        if (plain.length > 3 && plain.length < 80) {
            if (SECTION_KEYWORDS.test(plain) || PRODUCTION_SUMMARY_RE.test(plain) || PRODUCTION_NOTES_RE.test(plain)) {
                return plain;
            }
        }

        // 4. ALL-CAPS line that matches keywords
        if (/^[A-Z][A-Z0-9 \/:\-]{4,}$/.test(norm) && SECTION_KEYWORDS.test(norm)) {
            return norm;
        }

        return null;
    }

    // Map AI section labels to A/V template section names
    function mapSectionLabel(rawLabel) {
        const upper = rawLabel.toUpperCase().replace(/[\[\]]/g, '').trim();
        if (upper.includes('HOOK') || upper.includes('COLD OPEN')) return 'Video Introduction';
        if (upper.includes('INTRODUCTION')) return 'Video Introduction';
        if (upper.includes('LEARNING OBJECTIVE')) return 'Video Introduction';
        if (upper.includes('RECAP') || upper.includes('SUMMARY')) return 'Video Conclusion';
        if (upper.includes('CONCLUSION')) return 'Video Conclusion';
        if (upper.includes('CALL TO ACTION')) return 'Video Conclusion';
        if (PRODUCTION_NOTES_RE.test(rawLabel)) return null;
        if (PRODUCTION_SUMMARY_RE.test(rawLabel)) return null;
        // SECTION N: Title -> Video Content: Title
        const secMatch = upper.match(/SECTION\s+\d+[:\s\-]*(.*)/);
        if (secMatch) {
            const title = secMatch[1].trim();
            return title ? 'Video Content: ' + title : 'Video Content';
        }
        // If nothing else matched but it got through detectSectionHeader, use it as-is
        return rawLabel;
    }

    function parseRawScript(raw) {
        const lines = raw.split('\n');
        const sections = [];
        let currentSection = null;
        let productionSummary = [];
        let productionNotes = [];
        let inProductionSummary = false;
        let inProductionNotes = false;
        let allDialogue = [];  // fallback collector

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const trimmed = line.trim();
            if (!trimmed) continue;

            // Detect section header
            const headerLabel = detectSectionHeader(trimmed);

            if (headerLabel) {
                if (PRODUCTION_SUMMARY_RE.test(headerLabel)) {
                    inProductionSummary = true;
                    inProductionNotes = false;
                    currentSection = null;
                    continue;
                }
                if (PRODUCTION_NOTES_RE.test(headerLabel)) {
                    inProductionNotes = true;
                    inProductionSummary = false;
                    currentSection = null;
                    continue;
                }
                inProductionSummary = false;
                inProductionNotes = false;
                const label = mapSectionLabel(headerLabel);
                if (label) {
                    currentSection = { label, dialogue: [], graphics: [] };
                    sections.push(currentSection);
                }
                continue;
            }

            if (inProductionSummary) {
                productionSummary.push(trimmed.replace(/^[\-\*]\s*/, ''));
                continue;
            }
            if (inProductionNotes) {
                productionNotes.push(trimmed.replace(/^[\-\*]\s*/, ''));
                continue;
            }

            // Skip horizontal rules
            if (/^-{3,}$/.test(trimmed)) continue;

            // Classify line: dialogue or graphics
            const cleanLine = trimmed.replace(/\*+/g, '').trim();
            if (!cleanLine) continue;

            // Collect for fallback regardless
            allDialogue.push(cleanLine);

            if (!currentSection) continue;

            // If the entire line is a standalone graphics cue
            if (GRAPHICS_CUE_RE.test(trimmed)) {
                currentSection.graphics.push(cleanLine);
            } else if (TIMING_RE.test(trimmed)) {
                currentSection.graphics.push(cleanLine);
            } else {
                // Extract inline graphics cues from within dialogue lines
                // e.g., "...some dialogue [SLIDE: something] more dialogue *Approx. 0:30*"
                let dialoguePart = cleanLine;
                const inlineCues = [];

                // Pull out [SLIDE:...], [B-ROLL:...], [TRANSITION:...]
                dialoguePart = dialoguePart.replace(/\[(?:SLIDE|B-ROLL|TRANSITION):[^\]]*\]/gi, function(m) {
                    inlineCues.push(m); return '';
                });
                // Pull out (Visual:...)
                dialoguePart = dialoguePart.replace(/\(Visual:[^)]*\)/gi, function(m) {
                    inlineCues.push(m); return '';
                });
                // Pull out timing markers: *Approx. 0:30* or Approx. 0:30
                dialoguePart = dialoguePart.replace(/\*?Approx\.?\s+[\d:]+\*?/gi, function(m) {
                    inlineCues.push(m.replace(/\*/g, '').trim()); return '';
                });

                // Clean up dialogue (collapse extra spaces)
                dialoguePart = dialoguePart.replace(/\s{2,}/g, ' ').trim();
                // Strip bullet prefix
                dialoguePart = dialoguePart.replace(/^[\-\*]\s*/, '');

                if (dialoguePart) {
                    currentSection.dialogue.push(dialoguePart);
                }
                if (inlineCues.length > 0) {
                    for (const cue of inlineCues) {
                        currentSection.graphics.push(cue);
                    }
                }
            }
        }

        // Merge consecutive sections with the same label
        const merged = [];
        for (const s of sections) {
            const last = merged[merged.length - 1];
            if (last && last.label === s.label) {
                last.dialogue.push('', ...s.dialogue);
                last.graphics.push(...s.graphics);
            } else {
                merged.push({ ...s, dialogue: [...s.dialogue], graphics: [...s.graphics] });
            }
        }

        // FALLBACK: if no sections were parsed, put all text in one section
        if (merged.length === 0 && allDialogue.length > 0) {
            merged.push({
                label: 'Video Content',
                dialogue: allDialogue,
                graphics: []
            });
        }

        return { sections: merged, productionSummary, productionNotes };
    }

    /* ── Extract metadata from form + AI output ───── */

    function extractMetadata(parsed) {
        const topic = (document.getElementById('topic') || {}).value || '';
        const runtime = (document.getElementById('runtime') || {}).value || '';
        const courseContext = (document.getElementById('courseContext') || {}).value || '';

        let title = topic || 'Untitled Script';
        let duration = runtime || '';

        // Try to pull from production summary
        for (const line of parsed.productionSummary) {
            if (/working title/i.test(line)) {
                const m = line.match(/:\s*(.+)/);
                if (m) title = m[1].trim();
            }
            if (/runtime|duration/i.test(line)) {
                const m = line.match(/:\s*(.+)/);
                if (m) duration = m[1].trim();
            }
        }

        const today = new Date();
        const revDate = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;

        return { title, duration, moduleSection: courseContext || '', revDate };
    }

    /* ── Template rewiring (reuse from tli-download) ── */

    function stripAttachedTemplate(xml) {
        return xml.replace(/<Relationship[^>]*attachedTemplate[^>]*\/>/gi, '');
    }
    function stripAttachedTemplateSettings(xml) {
        return xml.replace(/<w:attachedTemplate[^>]*\/>/gi, '');
    }

    /* ── Main export function ─────────────────────── */

    async function buildAVScriptDocx(rawText) {
        // 1. Load JSZip (provided by tli-download.js)
        if (!window.JSZip && window.__tliDownloadInit) {
            // JSZip loader exposed by tli-download.js on window
            const s = document.createElement('script');
            s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
            document.head.appendChild(s);
            await new Promise(r => { s.onload = r; });
        }
        const JSZip = window.JSZip;
        if (!JSZip) throw new Error('JSZip is not available');

        // 2. Fetch AV Script template
        const resp = await fetch(AV_TEMPLATE_URL, { cache: 'force-cache' });
        if (!resp.ok) throw new Error(`Could not load AV Script template (${resp.status})`);
        const buf = await resp.arrayBuffer();
        const zip = await JSZip.loadAsync(buf);

        // 2b. Fix Content_Types: template was .dotx, output must be .docx
        const ctFile = zip.file('[Content_Types].xml');
        if (ctFile) {
            let ctXml = await ctFile.async('string');
            ctXml = ctXml.replace(
                'wordprocessingml.template.main+xml',
                'wordprocessingml.document.main+xml'
            );
            zip.file('[Content_Types].xml', ctXml);
        }

        // 3. Parse the raw script text
        const parsed = parseRawScript(rawText);
        const meta = extractMetadata(parsed);

        // 4. Build new body XML
        const metaTable = buildMetadataTable(meta.title, meta.duration, meta.moduleSection, meta.revDate);
        const avTable = buildAVTable(parsed.sections);

        // Production notes as paragraphs after the table
        let notesXml = '';
        if (parsed.productionNotes.length > 0) {
            notesXml += `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>${textRun('Production Notes', RUN_BOLD)}</w:p>`;
            for (const note of parsed.productionNotes) {
                notesXml += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${textRun(note)}</w:p>`;
            }
        }

        const newBodyInner = metaTable + '<w:p/>' + avTable + '<w:p/>' + notesXml;

        // 5. Replace document.xml body
        const docXml = await zip.file('word/document.xml').async('string');
        const bodyOpen = docXml.indexOf('<w:body>');
        const bodyClose = docXml.indexOf('</w:body>');
        const bodyInner = docXml.substring(bodyOpen + '<w:body>'.length, bodyClose);
        const sectPrMatch = bodyInner.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
        const sectPr = sectPrMatch ? sectPrMatch[0] : '';
        const head = docXml.substring(0, bodyOpen + '<w:body>'.length);
        const tail = docXml.substring(bodyClose);
        zip.file('word/document.xml', head + newBodyInner + sectPr + tail);

        // 6. Strip attachedTemplate
        const relsFile = zip.file('word/_rels/settings.xml.rels');
        if (relsFile) {
            const relsXml = await relsFile.async('string');
            zip.file('word/_rels/settings.xml.rels', stripAttachedTemplate(relsXml));
        }
        const settingsFile = zip.file('word/settings.xml');
        if (settingsFile) {
            const settingsXml = await settingsFile.async('string');
            zip.file('word/settings.xml', stripAttachedTemplateSettings(settingsXml));
        }

        // 7. Generate blob
        return zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            compression: 'DEFLATE'
        });
    }

    /* ── Override downloadWord for this tool ──────── */

    window.downloadWord = async function () {
        const raw = window.__lastOutput || '';
        if (!raw.trim()) {
            alert('Nothing to download yet. Generate a script first.');
            return;
        }

        const btn = event && event.target;
        if (btn) {
            btn.disabled = true;
            const orig = btn.textContent;
            btn.textContent = 'Building A/V script...';
            btn.dataset.origLabel = orig;
        }

        try {
            const blob = await buildAVScriptDocx(raw);

            // Build filename
            const topic = (document.getElementById('topic') || {}).value || 'Script';
            const safeName = topic.replace(/[^a-z0-9]+/gi, '-').replace(/^-+|-+$/g, '');
            const d = new Date();
            const pad = n => String(n).padStart(2, '0');
            const ts = `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;

            // Trigger download
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `AV-Script-${safeName}-${ts}.docx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(url), 500);

            if (btn) {
                btn.textContent = 'Downloaded \u2713';
                btn.classList.add('copied');
                setTimeout(() => {
                    btn.textContent = btn.dataset.origLabel || 'Download Word';
                    btn.classList.remove('copied');
                    btn.disabled = false;
                }, 1800);
            }
        } catch (err) {
            console.error('[tli-script-export] build failed:', err);
            alert('Could not build the A/V Script document. ' + (err && err.message || '') + '\nSee browser console for details.');
            if (btn) {
                btn.textContent = btn.dataset.origLabel || 'Download Word';
                btn.disabled = false;
            }
        }
    };

})();
