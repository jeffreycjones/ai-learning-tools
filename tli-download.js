/* ============================================================
 * TLI Learning Design Suite — Shared Download Module
 * ------------------------------------------------------------
 * Exposes two global functions used by per-tool buttons:
 *   - downloadMarkdown()  saves output as .md
 *   - downloadWord()      saves output as a real .docx built
 *                         from the Mays Business School TLI
 *                         brand template (assets/Template.docx).
 *                         Loads the template, replaces its body
 *                         with content translated from the
 *                         tool's rendered HTML, strips the
 *                         attachedTemplate reference per the
 *                         mays-docx-brand-standards skill, and
 *                         repacks as .docx.
 *
 * Reads output from (first match wins):
 *   #outputArea, #outputPanel, .output-area, .output-panel
 *
 * Falls back to window.__lastOutput for markdown when present.
 *
 * To update brand output, swap assets/Template.docx. No code
 * changes required unless the template's style IDs change.
 * ============================================================ */
(function () {
    'use strict';

    if (window.__tliDownloadInit) return;
    window.__tliDownloadInit = true;

    const TEMPLATE_URL = 'assets/Template.docx';
    const JSZIP_URL = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';

    /* ---------- Small utilities ---------- */

    function timestamp() {
        const d = new Date();
        const pad = n => String(n).padStart(2, '0');
        return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;
    }

    function triggerDownload(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 500);
    }

    function flashButton(btn, label) {
        if (!btn) return;
        const original = btn.textContent;
        btn.textContent = label;
        btn.classList.add('copied');
        setTimeout(() => { btn.textContent = original; btn.classList.remove('copied'); }, 1800);
    }

    function outputElement() {
        return document.getElementById('outputArea')
            || document.getElementById('outputPanel')
            || document.querySelector('.output-area')
            || document.querySelector('.output-panel');
    }

    function toolTitle() {
        const t = (document.title || 'Output').split('|')[0].trim();
        return t || 'Output';
    }

    function safeFilename(s) {
        return (s || 'Output').replace(/[^a-z0-9]+/gi, '-').replace(/^-+|-+$/g, '');
    }

    function isEmptyOutput(html) {
        if (!html || !html.trim()) return true;
        return /placeholder|generating|will appear/i.test(html);
    }

    function xmlEscape(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    /* ---------- Lazy JSZip loader ---------- */

    let jszipPromise = null;
    function loadJSZip() {
        if (window.JSZip) return Promise.resolve(window.JSZip);
        if (jszipPromise) return jszipPromise;
        jszipPromise = new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = JSZIP_URL;
            s.onload = () => window.JSZip ? resolve(window.JSZip) : reject(new Error('JSZip failed to initialize'));
            s.onerror = () => reject(new Error('Failed to load JSZip from CDN'));
            document.head.appendChild(s);
        });
        return jszipPromise;
    }

    /* ---------- HTML to OOXML translator ---------- */

    const HEADING_MAP = {
        h1: 'Heading1', h2: 'Heading2', h3: 'Heading3',
        h4: 'Heading4', h5: 'Heading5', h6: 'Heading5'
    };

    function buildRuns(node, inheritedRPr) {
        const runs = [];
        const rPr = inheritedRPr || '';

        const emitText = (text, props) => {
            if (text == null || text === '') return;
            // Collapse any tab or newline characters in inline runs to single spaces.
            text = text.replace(/[\r\n\t]+/g, ' ');
            const rPrXml = props ? `<w:rPr>${props}</w:rPr>` : '';
            runs.push(`<w:r>${rPrXml}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>`);
        };

        function walk(n, props) {
            if (n.nodeType === 3) { // text node
                emitText(n.nodeValue, props);
                return;
            }
            if (n.nodeType !== 1) return;
            const tag = n.tagName.toLowerCase();

            if (tag === 'br') { runs.push('<w:r><w:br/></w:r>'); return; }

            let nextProps = props || '';
            if (tag === 'strong' || tag === 'b') nextProps += '<w:b/>';
            else if (tag === 'em' || tag === 'i') nextProps += '<w:i/>';
            else if (tag === 'u') nextProps += '<w:u w:val="single"/>';
            else if (tag === 'code') nextProps += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>';
            else if (tag === 's' || tag === 'strike' || tag === 'del') nextProps += '<w:strike/>';

            for (const child of n.childNodes) walk(child, nextProps);
        }

        walk(node, rPr);
        return runs.join('');
    }

    function buildParagraph(el, styleName, extraPPr) {
        const runs = buildRuns(el, '');
        const parts = [];
        if (styleName) parts.push(`<w:pStyle w:val="${styleName}"/>`);
        if (extraPPr) parts.push(extraPPr);
        const pPr = parts.length ? `<w:pPr>${parts.join('')}</w:pPr>` : '';
        return `<w:p>${pPr}${runs}</w:p>`;
    }

    function buildListParagraphs(listEl, styleName, parts) {
        for (const li of listEl.children) {
            if (li.tagName && li.tagName.toLowerCase() === 'li') {
                parts.push(buildParagraph(li, styleName));
                // Nested lists inside the li
                for (const child of li.children) {
                    const t = child.tagName && child.tagName.toLowerCase();
                    if (t === 'ul') buildListParagraphs(child, 'ListBullet2', parts);
                    else if (t === 'ol') buildListParagraphs(child, 'ListNumber2', parts);
                }
            }
        }
    }

    function buildTable(tableEl) {
        const rows = [];
        const trs = tableEl.querySelectorAll('tr');
        let colCount = 0;
        const headerTexts = [];
        trs.forEach(tr => {
            const cells = tr.querySelectorAll('th, td');
            if (cells.length > colCount) colCount = cells.length;
        });
        if (!colCount) return '';

        // Collect header text for smart column sizing
        const firstRow = trs[0];
        if (firstRow) {
            firstRow.querySelectorAll('th, td').forEach(c => headerTexts.push((c.textContent || '').trim().toLowerCase()));
        }

        // Smart column widths: allocate based on content type
        // Total page width ~10060 dxa for landscape-friendly tables
        const PAGE_W = 10060;
        let colWidths;
        if (colCount <= 3) {
            // Equal distribution for small tables
            const w = Math.floor(PAGE_W / colCount);
            colWidths = Array(colCount).fill(w);
        } else {
            // Weighted distribution: short labels get narrow columns, long content gets wide
            const weights = headerTexts.map(h => {
                if (/^(#|no\.?|shot\s*#?)$/i.test(h)) return 1;
                if (/^(dur|duration|time|length)$/i.test(h)) return 1.5;
                if (/^(type|shot\s*type|format)$/i.test(h)) return 1.5;
                if (/visual|description|narration|audio|script|content|detail/i.test(h)) return 4;
                if (/graphic|note|comment/i.test(h)) return 2.5;
                return 2; // default medium
            });
            // Pad to colCount if header row had fewer cells
            while (weights.length < colCount) weights.push(2);
            const totalWeight = weights.reduce((a, b) => a + b, 0);
            colWidths = weights.map(w => Math.round((w / totalWeight) * PAGE_W));
        }

        trs.forEach((tr, rowIdx) => {
            const cells = [];
            const cellEls = tr.querySelectorAll('th, td');
            cellEls.forEach((td, colIdx) => {
                const isHeader = td.tagName.toLowerCase() === 'th';
                const runs = buildRuns(td, isHeader ? '<w:b/><w:color w:val="FFFFFF"/>' : '');
                const pPrParts = [`<w:pStyle w:val="${isHeader ? 'TableTextbold' : 'TableText'}"/>`];
                // Smaller font for dense tables
                if (colCount >= 6) pPrParts.push('<w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>');
                const pPr = `<w:pPr>${pPrParts.join('')}</w:pPr>`;
                const tcPrParts = [];
                if (colWidths[colIdx]) tcPrParts.push(`<w:tcW w:w="${colWidths[colIdx]}" w:type="dxa"/>`);
                if (isHeader) tcPrParts.push('<w:shd w:val="clear" w:color="auto" w:fill="500000"/>');
                // Alternating row shading for readability on dense tables
                if (!isHeader && colCount >= 6 && rowIdx % 2 === 0) {
                    tcPrParts.push('<w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>');
                }
                const tcPr = tcPrParts.length ? `<w:tcPr>${tcPrParts.join('')}</w:tcPr>` : '';
                cells.push(`<w:tc>${tcPr}<w:p>${pPr}${runs}</w:p></w:tc>`);
            });
            rows.push(`<w:tr>${cells.join('')}</w:tr>`);
        });

        const tblPr = [
            '<w:tblPr>',
            '<w:tblStyle w:val="TableGrid"/>',
            '<w:tblW w:w="5000" w:type="pct"/>',
            '<w:tblBorders>',
                '<w:top w:val="single" w:sz="4" w:space="0" w:color="D1D1D1"/>',
                '<w:left w:val="single" w:sz="4" w:space="0" w:color="D1D1D1"/>',
                '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D1D1D1"/>',
                '<w:right w:val="single" w:sz="4" w:space="0" w:color="D1D1D1"/>',
                '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D1D1D1"/>',
                '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D1D1D1"/>',
            '</w:tblBorders>',
            '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>',
            '</w:tblPr>'
        ].join('');

        const grid = `<w:tblGrid>${colWidths.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;

        return `<w:tbl>${tblPr}${grid}${rows.join('')}</w:tbl>`;
    }

    function hasBlockChildren(el) {
        return !!el.querySelector('h1,h2,h3,h4,h5,h6,p,ul,ol,table,div,blockquote');
    }

    function translateBlocks(container, parts) {
        for (const child of container.childNodes) {
            if (child.nodeType === 3) {
                const txt = (child.nodeValue || '').trim();
                if (txt) {
                    const tmp = document.createElement('span');
                    tmp.textContent = txt;
                    parts.push(buildParagraph(tmp, null));
                }
                continue;
            }
            if (child.nodeType !== 1) continue;
            const tag = child.tagName.toLowerCase();

            if (HEADING_MAP[tag]) {
                parts.push(buildParagraph(child, HEADING_MAP[tag]));
            } else if (tag === 'p') {
                parts.push(buildParagraph(child, null));
            } else if (tag === 'ul') {
                buildListParagraphs(child, 'ListBullet', parts);
            } else if (tag === 'ol') {
                buildListParagraphs(child, 'ListNumber', parts);
            } else if (tag === 'table') {
                parts.push(buildTable(child));
                parts.push('<w:p/>');
            } else if (tag === 'blockquote') {
                parts.push(buildParagraph(child, null, '<w:ind w:left="720"/><w:pBdr><w:left w:val="single" w:sz="12" w:space="6" w:color="500000"/></w:pBdr>'));
            } else if (tag === 'hr') {
                parts.push('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="D1D1D1"/></w:pBdr></w:pPr></w:p>');
            } else if (tag === 'br') {
                parts.push('<w:p/>');
            } else if (tag === 'div' || tag === 'section' || tag === 'article') {
                if (hasBlockChildren(child)) {
                    translateBlocks(child, parts);
                } else {
                    parts.push(buildParagraph(child, null));
                }
            } else {
                // Unknown inline-ish at block level. Treat as paragraph.
                parts.push(buildParagraph(child, null));
            }
        }
    }

    function htmlToBodyXml(html) {
        const container = document.createElement('div');
        container.innerHTML = html;
        // Hoist block elements (tables, headings, lists) that ended up nested
        // inside containers that translateBlocks won't recurse into properly.
        // AI-generated HTML often produces invalid nesting like <ul><table>...
        // or <p><table>... which causes the table to be silently skipped.
        // Run repeatedly until stable because hoisting can expose new nesting.
        let hoisted;
        do {
            hoisted = false;
            container.querySelectorAll('table, h1, h2, h3, h4, h5, h6').forEach(function(block) {
                const parent = block.parentNode;
                if (parent === container) return; // already at top level
                const grandparent = parent.parentNode;
                if (!grandparent) return;
                // Hoist the block to just after its parent
                grandparent.insertBefore(block, parent.nextSibling);
                hoisted = true;
            });
        } while (hoisted);
        const parts = [];
        translateBlocks(container, parts);
        return parts.join('');
    }

    /* ---------- Template rewiring ---------- */

    function stripAttachedTemplate(settingsRelsXml) {
        return settingsRelsXml.replace(
            /<Relationship[^>]*attachedTemplate[^>]*\/>/gi,
            ''
        );
    }

    function stripAttachedTemplateFromSettings(settingsXml) {
        return settingsXml.replace(
            /<w:attachedTemplate[^>]*\/>/gi,
            ''
        );
    }

    function replaceBodyInDocumentXml(docXml, newBodyInnerXml) {
        // Preserve everything before <w:body> and the closing </w:document>.
        // Keep the final <w:sectPr> (page size, margins, headers, footers) untouched.
        const bodyOpen = docXml.indexOf('<w:body>');
        const bodyClose = docXml.indexOf('</w:body>');
        if (bodyOpen === -1 || bodyClose === -1) throw new Error('Template document.xml missing body tags');

        const bodyInner = docXml.substring(bodyOpen + '<w:body>'.length, bodyClose);
        const sectPrMatch = bodyInner.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
        const sectPr = sectPrMatch ? sectPrMatch[0] : '';

        const head = docXml.substring(0, bodyOpen + '<w:body>'.length);
        const tail = docXml.substring(bodyClose);

        return head + newBodyInnerXml + sectPr + tail;
    }

    async function buildDocxBlob(rendered, title) {
        const JSZip = await loadJSZip();

        // Fetch template. Use cache: reload the first time to ensure fresh copy,
        // but otherwise let the browser cache it between downloads.
        const resp = await fetch(TEMPLATE_URL, { cache: 'force-cache' });
        if (!resp.ok) throw new Error(`Could not load Template.docx (${resp.status})`);
        const templateBuf = await resp.arrayBuffer();
        const zip = await JSZip.loadAsync(templateBuf);

        // Build the new body XML.
        const generatedOn = new Date().toLocaleString();
        const titlePara = `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">${xmlEscape(title)}</w:t></w:r></w:p>`;
        const metaPara = `<w:p><w:pPr><w:pStyle w:val="DocSubtitle"/></w:pPr><w:r><w:t xml:space="preserve">Generated ${xmlEscape(generatedOn)}</w:t></w:r></w:p>`;
        // Hook: tools can set window.tliDocxPreprocess to a function(html) => html
        // that restructures the rendered HTML before OOXML translation. This allows
        // per-tool customization (e.g., script-specific heading hierarchy) without
        // duplicating the core .docx machinery.
        const processedHtml = typeof window.tliDocxPreprocess === 'function'
            ? window.tliDocxPreprocess(rendered)
            : rendered;
        const bodyContent = htmlToBodyXml(processedHtml);
        const newBodyInner = titlePara + metaPara + bodyContent;

        // Rewrite document.xml.
        const docXml = await zip.file('word/document.xml').async('string');
        const newDocXml = replaceBodyInDocumentXml(docXml, newBodyInner);
        zip.file('word/document.xml', newDocXml);

        // Strip attachedTemplate from both places (mandatory per skill).
        const relsFile = zip.file('word/_rels/settings.xml.rels');
        if (relsFile) {
            const relsXml = await relsFile.async('string');
            zip.file('word/_rels/settings.xml.rels', stripAttachedTemplate(relsXml));
        }
        const settingsFile = zip.file('word/settings.xml');
        if (settingsFile) {
            const settingsXml = await settingsFile.async('string');
            zip.file('word/settings.xml', stripAttachedTemplateFromSettings(settingsXml));
        }

        // Generate final .docx blob.
        return zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            compression: 'DEFLATE'
        });
    }

    /* ---------- Public API ---------- */

    window.downloadMarkdown = function () {
        const el = outputElement();
        const raw = window.__lastOutput || (el ? el.innerText : '');
        if (isEmptyOutput(raw)) {
            alert('Nothing to download yet. Generate output first.');
            return;
        }
        const title = toolTitle();
        const header = `# ${title}\n\nMays Business School | Teaching & Learning Innovation\nGenerated: ${new Date().toLocaleString()}\n\n---\n\n`;
        const blob = new Blob([header + raw], { type: 'text/markdown;charset=utf-8' });
        triggerDownload(blob, `${safeFilename(title)}-${timestamp()}.md`);
        flashButton(event && event.target, 'Downloaded \u2713');
    };

    window.downloadWord = async function () {
        const el = outputElement();
        const rendered = el ? el.innerHTML : '';
        if (isEmptyOutput(rendered)) {
            alert('Nothing to download yet. Generate output first.');
            return;
        }
        const btn = event && event.target;
        if (btn) {
            btn.disabled = true;
            const orig = btn.textContent;
            btn.textContent = 'Building .docx...';
            btn.dataset.origLabel = orig;
        }
        try {
            const title = toolTitle();
            const blob = await buildDocxBlob(rendered, title);
            triggerDownload(blob, `${safeFilename(title)}-${timestamp()}.docx`);
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
            console.error('[tli-download] .docx build failed:', err);
            alert('Could not build the Word document. ' + (err && err.message ? err.message : '') + '\n\nSee browser console for details.');
            if (btn) {
                btn.textContent = btn.dataset.origLabel || 'Download Word';
                btn.disabled = false;
            }
        }
    };
})();
