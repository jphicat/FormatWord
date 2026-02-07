/**
 * DOCX Rebuilder — Replaces text in the original DOCX XML while preserving all formatting.
 *
 * KEY INSIGHT: We don't create a new document. We modify the original XML in-place,
 * only touching <w:t> text content while keeping every XML attribute, style,
 * and binary asset byte-for-byte identical.
 */

import {
    serializeXml,
    getDirectRuns,
    getRunTextNodes,
    extractTextFromParagraph,
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

    // Build a quick lookup: xmlPath → list of segments
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

        // Get all content paragraphs from the DOM
        const allParas = getAllParagraphs(doc.documentElement);
        const contentParas = allParas.filter(isContentParagraph);

        // Match segments to paragraphs by their text content
        let segIdx = 0;
        for (const para of contentParas) {
            if (segIdx >= fileSegments.length) break;

            const paraText = extractTextFromParagraph(para);
            const segment = fileSegments[segIdx];

            // Verify this paragraph matches our segment
            if (paraText.trim() === segment.text.trim()) {
                const translatedText = mapping.get(segment.id);
                if (translatedText !== undefined) {
                    replaceParagraphText(para, translatedText);
                    processedCount++;
                }
                segIdx++;
            }
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
 * - If the paragraph has only 1 run: just replace the text.
 * - If multiple runs: distribute the translated text proportionally based on
 *   the original character ratios, so each run keeps its formatting over
 *   a proportional chunk of the translated text.
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
        if (ri.textNodes.length === 1) {
            ri.textNodes[0].textContent = newText;
            // Preserve xml:space="preserve" to keep leading/trailing spaces
            ri.textNodes[0].setAttribute('xml:space', 'preserve');
        } else if (ri.textNodes.length > 1) {
            // Put all text in the first <w:t>, clear the rest
            ri.textNodes[0].textContent = newText;
            ri.textNodes[0].setAttribute('xml:space', 'preserve');
            for (let i = 1; i < ri.textNodes.length; i++) {
                ri.textNodes[i].textContent = '';
            }
        }
        return;
    }

    // Multiple runs: distribute text proportionally
    let remaining = newText;

    for (let i = 0; i < textRuns.length; i++) {
        const ri = textRuns[i];
        const isLast = i === textRuns.length - 1;

        let chunk;
        if (isLast) {
            // Last run gets all remaining text
            chunk = remaining;
        } else {
            // Proportional split based on original lengths
            const ratio = ri.length / totalOriginalLen;
            const chunkLen = Math.round(newText.length * ratio);

            // Try to split at a word boundary near the proportional point
            chunk = remaining.substring(0, chunkLen);
            const lastSpace = chunk.lastIndexOf(' ');
            if (lastSpace > chunkLen * 0.5) {
                chunk = remaining.substring(0, lastSpace + 1);
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

        remaining = remaining.substring(chunk.length);
    }
}
