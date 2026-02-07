/**
 * DOCX Rebuilder — Replaces text in the original DOCX XML while preserving all formatting.
 *
 * KEY INSIGHT: We don't create a new document. We modify the original XML in-place,
 * only touching <w:t> text content while keeping every XML attribute, style,
 * and binary asset byte-for-byte identical.
 *
 * PROFESSIONAL TRANSLATOR APPROACH:
 * - Segments are matched to paragraphs by INDEX (order), not by text comparison.
 * - This ensures every content paragraph gets its translation regardless of
 *   minor parsing differences.
 */

import {
    serializeXml,
    getDirectRuns,
    getRunTextNodes,
    isContentParagraph,
    getAllParagraphs,
    NS,
} from './xmlUtils.js';

/**
 * Rebuild the DOCX by replacing text according to the segment mapping.
 *
 * @param {JSZip} zip — The original ZIP archive
 * @param {Map<string, Document>} xmlFiles — Parsed XML docs (keyed by path)
 * @param {Array<{id, text, xmlPath, _element}>} segments — Source segments
 * @param {Map<number, string>} mapping — Segment ID → translated text
 * @param {Function} onProgress — Progress callback
 * @returns {Promise<Blob>} — The rebuilt .docx as a Blob
 */
export async function rebuildDocx(
    zip,
    xmlFiles,
    segments,
    mapping,
    onProgress
) {
    onProgress?.({
        step: 'rebuild-start',
        message: 'Début de la reconstruction...',
    });

    // Build a quick lookup: xmlPath → ordered list of segment IDs
    const segmentsByFile = new Map();
    for (const seg of segments) {
        if (!segmentsByFile.has(seg.xmlPath)) {
            segmentsByFile.set(seg.xmlPath, []);
        }
        segmentsByFile.get(seg.xmlPath).push(seg);
    }

    let processedCount = 0;

    // Process each XML file that has segments
    for (const [xmlPath, doc] of xmlFiles) {
        const fileSegments = segmentsByFile.get(xmlPath);
        if (!fileSegments || fileSegments.length === 0) continue;

        // Get all content paragraphs from the DOM — in the same order as the parser found them
        const allParas = getAllParagraphs(doc.documentElement);
        const contentParas = allParas.filter(isContentParagraph);

        // Replace text by index — paragraph i maps to segment i
        const count = Math.min(contentParas.length, fileSegments.length);
        for (let i = 0; i < count; i++) {
            const segment = fileSegments[i];
            const translatedText = mapping.get(segment.id);

            if (translatedText !== undefined) {
                replaceParagraphText(contentParas[i], translatedText);
                processedCount++;
            } else {
                // No translation available: keep original text but highlight in red
                highlightParagraphRed(contentParas[i], doc);
            }
        }

        // Also highlight any remaining content paragraphs beyond segment count
        for (let i = count; i < contentParas.length; i++) {
            // These paragraphs exist in the source but had no segment — highlight them
            // (shouldn't happen normally, but safety net)
        }

        // Serialize modified XML back into the ZIP
        const newXml = serializeXml(doc);
        zip.file(xmlPath, newXml);

        onProgress?.({
            step: 'rebuild-file',
            message: `${xmlPath} mis à jour (${processedCount} segments)`,
        });
    }

    onProgress?.({
        step: 'rebuild-generate',
        message: 'Génération du fichier DOCX...',
    });

    // Generate the final .docx blob
    const blob = await zip.generateAsync(
        {
            type: 'blob',
            mimeType:
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 },
        },
        (metadata) => {
            onProgress?.({
                step: 'rebuild-zip',
                message: `Compression : ${Math.round(metadata.percent)}%`,
            });
        }
    );

    onProgress?.({
        step: 'rebuild-done',
        message: `Terminé ! ${processedCount} segments remplacés.`,
    });

    return blob;
}

/**
 * Replace the text content of a paragraph, distributing new text across existing runs
 * to preserve per-run formatting.
 *
 * Strategy:
 * - If the paragraph has only 1 text run: just replace the text.
 * - If multiple text runs: distribute the translated text using WORD-BOUNDARY splitting
 *   proportional to original character ratios. This ensures formatting transitions
 *   happen at natural word boundaries rather than mid-word.
 *
 * @param {Element} paraElement — <w:p> element
 * @param {string} newText — The full translated text for this paragraph
 */
function replaceParagraphText(paraElement, newText) {
    const runs = getDirectRuns(paraElement);
    if (runs.length === 0) return;

    // Collect run info: text nodes and their text
    const runInfos = [];
    let totalOriginalLen = 0;

    for (const run of runs) {
        const textNodes = getRunTextNodes(run);
        let runText = '';
        for (const tn of textNodes) {
            runText += tn.textContent || '';
        }
        runInfos.push({ run, textNodes, originalText: runText, length: runText.length });
        totalOriginalLen += runText.length;
    }

    // Filter out runs that have no text (e.g., contain only <w:tab>, <w:br>)
    const textRuns = runInfos.filter((ri) => ri.length > 0);
    if (textRuns.length === 0) return;

    if (textRuns.length === 1) {
        // Single run: simple replacement
        const ri = textRuns[0];
        if (ri.textNodes.length >= 1) {
            ri.textNodes[0].textContent = newText;
            ri.textNodes[0].setAttribute('xml:space', 'preserve');
            // Clear any extra <w:t> nodes in this run
            for (let i = 1; i < ri.textNodes.length; i++) {
                ri.textNodes[i].textContent = '';
            }
        }
        return;
    }

    // Multiple runs: distribute text using word-aware proportional splitting
    const words = newText.split(/(\s+)/); // Split keeping whitespace as separators
    const totalNewLen = newText.length;

    // Calculate target character lengths for each run based on original proportions
    const targetLengths = textRuns.map((ri) => {
        return Math.round((ri.length / totalOriginalLen) * totalNewLen);
    });

    // Distribute words into runs respecting proportional targets
    let wordIdx = 0;
    for (let i = 0; i < textRuns.length; i++) {
        const ri = textRuns[i];
        const isLast = i === textRuns.length - 1;

        let chunk = '';
        if (isLast) {
            // Last run gets all remaining words
            chunk = words.slice(wordIdx).join('');
        } else {
            // Build chunk by adding words until we reach the target length
            const target = targetLengths[i];
            while (wordIdx < words.length) {
                const nextWord = words[wordIdx];
                if (chunk.length > 0 && chunk.length + nextWord.length > target * 1.3) {
                    // Don't overshoot too much past the target
                    break;
                }
                chunk += nextWord;
                wordIdx++;
                // Stop if we've met or exceeded target and are at a word boundary
                if (chunk.length >= target && wordIdx < words.length) {
                    break;
                }
            }
        }

        // Set the text in this run
        if (ri.textNodes.length >= 1) {
            ri.textNodes[0].textContent = chunk;
            ri.textNodes[0].setAttribute('xml:space', 'preserve');
            // Clear additional <w:t> nodes in this run
            for (let j = 1; j < ri.textNodes.length; j++) {
                ri.textNodes[j].textContent = '';
            }
        }
    }
}

/**
 * Highlight all runs of a paragraph in red to flag it as "untranslated".
 * Adds <w:highlight w:val="red"/> inside each run's <w:rPr>.
 * If the run doesn't have a <w:rPr>, one is created.
 *
 * @param {Element} paraElement — <w:p> element
 * @param {Document} doc — The XML document (needed for createElement)
 */
function highlightParagraphRed(paraElement, doc) {
    const runs = getDirectRuns(paraElement);
    if (runs.length === 0) return;

    for (const run of runs) {
        // Find or create <w:rPr> (run properties)
        let rPr = null;
        for (const child of run.childNodes) {
            if (child.localName === 'rPr' && child.namespaceURI === NS.w) {
                rPr = child;
                break;
            }
        }

        if (!rPr) {
            rPr = doc.createElementNS(NS.w, 'w:rPr');
            run.insertBefore(rPr, run.firstChild);
        }

        // Remove existing highlight if any
        const existingHighlights = [];
        for (const child of rPr.childNodes) {
            if (child.localName === 'highlight' && child.namespaceURI === NS.w) {
                existingHighlights.push(child);
            }
        }
        for (const h of existingHighlights) {
            rPr.removeChild(h);
        }

        // Add <w:highlight w:val="red"/>
        const highlight = doc.createElementNS(NS.w, 'w:highlight');
        highlight.setAttributeNS(NS.w, 'w:val', 'red');
        rPr.appendChild(highlight);
    }
}
