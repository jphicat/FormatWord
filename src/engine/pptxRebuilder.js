/**
 * PPTX Rebuilder — Replaces text in the original PPTX XML while preserving all formatting.
 *
 * Same strategy as docxRebuilder.js:
 * - We modify the original XML in-place, only touching <a:t> text content
 * - Segments are matched to paragraphs by INDEX (order), not by text comparison
 * - Untranslated segments get red highlighting to flag them
 */

import {
    serializeXml,
    getPptxTextShapes,
    getPptxTxBody,
    getPptxParagraphs,
    getPptxRuns,
    getPptxTextNodes,
    isPptxContentParagraph,
    NS,
} from './xmlUtils.js';

/**
 * Rebuild the PPTX by replacing text according to the segment mapping.
 *
 * @param {JSZip} zip — The original ZIP archive
 * @param {Map<string, Document>} xmlFiles — Parsed slide XML docs
 * @param {Array<{id, text, xmlPath}>} segments — Source segments
 * @param {Map<number, string>} mapping — Segment ID → translated text
 * @param {Function} onProgress — Progress callback
 * @returns {Promise<Blob>} — The rebuilt .pptx as a Blob
 */
export async function rebuildPptx(zip, xmlFiles, segments, mapping, onProgress) {
    onProgress?.({
        step: 'rebuild-start',
        message: 'Début de la reconstruction PPTX...',
    });

    // Build lookup: xmlPath → ordered list of segments
    const segmentsByFile = new Map();
    for (const seg of segments) {
        if (!segmentsByFile.has(seg.xmlPath)) {
            segmentsByFile.set(seg.xmlPath, []);
        }
        segmentsByFile.get(seg.xmlPath).push(seg);
    }

    let processedCount = 0;

    // Process each slide XML
    for (const [xmlPath, doc] of xmlFiles) {
        const fileSegments = segmentsByFile.get(xmlPath);
        if (!fileSegments || fileSegments.length === 0) continue;

        // Get all content paragraphs from this slide in the same order as the parser
        const contentParas = [];
        const shapes = getPptxTextShapes(doc.documentElement);

        for (const shape of shapes) {
            const txBody = getPptxTxBody(shape);
            if (!txBody) continue;

            const paragraphs = getPptxParagraphs(txBody);
            for (const para of paragraphs) {
                if (isPptxContentParagraph(para)) {
                    contentParas.push(para);
                }
            }
        }

        // Replace text by index
        const count = Math.min(contentParas.length, fileSegments.length);
        for (let i = 0; i < count; i++) {
            const segment = fileSegments[i];
            const translatedText = mapping.get(segment.id);

            if (translatedText !== undefined) {
                replacePptxParagraphText(contentParas[i], translatedText);
                processedCount++;
            } else {
                // No translation: keep original text but highlight in red
                highlightPptxParagraphRed(contentParas[i], doc);
            }
        }

        // Serialize back into ZIP
        const newXml = serializeXml(doc);
        zip.file(xmlPath, newXml);

        onProgress?.({
            step: 'rebuild-file',
            message: `${xmlPath} mis à jour (${processedCount} segments)`,
        });
    }

    onProgress?.({
        step: 'rebuild-generate',
        message: 'Génération du fichier PPTX...',
    });

    // Generate the final .pptx blob
    const blob = await zip.generateAsync(
        {
            type: 'blob',
            mimeType:
                'application/vnd.openxmlformats-officedocument.presentationml.presentation',
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
 * Replace text content of a PPTX paragraph (<a:p>), distributing across runs.
 * Same word-boundary-aware logic as the DOCX rebuilder.
 */
function replacePptxParagraphText(paraElement, newText) {
    const runs = getPptxRuns(paraElement);
    if (runs.length === 0) return;

    // Collect run info
    const runInfos = [];
    let totalOriginalLen = 0;

    for (const run of runs) {
        const textNodes = getPptxTextNodes(run);
        let runText = '';
        for (const tn of textNodes) {
            runText += tn.textContent || '';
        }
        runInfos.push({ run, textNodes, originalText: runText, length: runText.length });
        totalOriginalLen += runText.length;
    }

    const textRuns = runInfos.filter((ri) => ri.length > 0);
    if (textRuns.length === 0) return;

    if (textRuns.length === 1) {
        // Single run: simple replacement
        const ri = textRuns[0];
        if (ri.textNodes.length >= 1) {
            ri.textNodes[0].textContent = newText;
            for (let i = 1; i < ri.textNodes.length; i++) {
                ri.textNodes[i].textContent = '';
            }
        }
        return;
    }

    // Multiple runs: word-boundary-aware proportional splitting
    const words = newText.split(/(\s+)/);
    const totalNewLen = newText.length;

    const targetLengths = textRuns.map((ri) => {
        return Math.round((ri.length / totalOriginalLen) * totalNewLen);
    });

    let wordIdx = 0;
    for (let i = 0; i < textRuns.length; i++) {
        const ri = textRuns[i];
        const isLast = i === textRuns.length - 1;

        let chunk = '';
        if (isLast) {
            chunk = words.slice(wordIdx).join('');
        } else {
            const target = targetLengths[i];
            while (wordIdx < words.length) {
                const nextWord = words[wordIdx];
                if (chunk.length > 0 && chunk.length + nextWord.length > target * 1.3) {
                    break;
                }
                chunk += nextWord;
                wordIdx++;
                if (chunk.length >= target && wordIdx < words.length) {
                    break;
                }
            }
        }

        if (ri.textNodes.length >= 1) {
            ri.textNodes[0].textContent = chunk;
            for (let j = 1; j < ri.textNodes.length; j++) {
                ri.textNodes[j].textContent = '';
            }
        }
    }
}

/**
 * Highlight a PPTX paragraph's runs in red to flag as untranslated.
 * Uses <a:highlight> element inside <a:rPr> (run properties).
 */
function highlightPptxParagraphRed(paraElement, doc) {
    const runs = getPptxRuns(paraElement);
    if (runs.length === 0) return;

    for (const run of runs) {
        // Find or create <a:rPr> (run properties)
        let rPr = null;
        for (const child of run.childNodes) {
            if (child.localName === 'rPr' && child.namespaceURI === NS.a) {
                rPr = child;
                break;
            }
        }

        if (!rPr) {
            rPr = doc.createElementNS(NS.a, 'a:rPr');
            run.insertBefore(rPr, run.firstChild);
        }

        // Remove existing highlight if any
        const existingHighlights = [];
        for (const child of rPr.childNodes) {
            if (child.localName === 'highlight' && child.namespaceURI === NS.a) {
                existingHighlights.push(child);
            }
        }
        for (const h of existingHighlights) {
            rPr.removeChild(h);
        }

        // Add <a:highlight><a:srgbClr val="FF0000"/></a:highlight>
        const highlight = doc.createElementNS(NS.a, 'a:highlight');
        const color = doc.createElementNS(NS.a, 'a:srgbClr');
        color.setAttribute('val', 'FF0000');
        highlight.appendChild(color);
        rPr.appendChild(highlight);
    }
}
