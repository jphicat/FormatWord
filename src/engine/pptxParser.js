/**
 * PPTX Parser — Extracts text segments from a .pptx ZIP archive.
 *
 * Same strategy as docxParser.js:
 * We extract a flat list of "segments" (paragraph-level text chunks)
 * with enough metadata to map them back to the XML later.
 *
 * PPTX structure:
 * - ppt/slides/slide1.xml, slide2.xml, ...
 * - Each slide has shapes (<p:sp>) containing text bodies (<p:txBody>)
 * - Text bodies have paragraphs (<a:p>) with runs (<a:r>) and text (<a:t>)
 */

import JSZip from 'jszip';
import {
    parseXml,
    getPptxTextShapes,
    getPptxTxBody,
    getPptxParagraphs,
    extractTextFromPptxParagraph,
    isPptxContentParagraph,
} from './xmlUtils.js';

/**
 * Parse a .pptx file (ArrayBuffer) and return parsed data.
 * @param {ArrayBuffer} buffer — The raw .pptx bytes
 * @returns {{ zip: JSZip, xmlFiles: Map<string, Document>, segments: Array, slideCount: number }}
 */
export async function parsePptx(buffer, onProgress) {
    const zip = await JSZip.loadAsync(buffer);

    onProgress?.({ step: 'unzip', message: 'Archive PPTX décompressée' });

    // Find all slide XML files (sorted numerically)
    const slideXmlPaths = [];
    for (const path of Object.keys(zip.files)) {
        if (/^ppt\/slides\/slide\d+\.xml$/i.test(path)) {
            slideXmlPaths.push(path);
        }
    }

    // Sort by slide number
    slideXmlPaths.sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)/i)[1], 10);
        const numB = parseInt(b.match(/slide(\d+)/i)[1], 10);
        return numA - numB;
    });

    // Also find notes slides
    const notesXmlPaths = [];
    for (const path of Object.keys(zip.files)) {
        if (/^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(path)) {
            notesXmlPaths.push(path);
        }
    }
    notesXmlPaths.sort((a, b) => {
        const numA = parseInt(a.match(/notesSlide(\d+)/i)[1], 10);
        const numB = parseInt(b.match(/notesSlide(\d+)/i)[1], 10);
        return numA - numB;
    });

    const slideCount = slideXmlPaths.length;

    onProgress?.({
        step: 'identify',
        message: `${slideCount} slides détectées`,
    });

    // Parse each XML file
    const xmlFiles = new Map();
    for (const path of slideXmlPaths) {
        const xmlString = await zip.file(path).async('string');
        const doc = parseXml(xmlString);
        xmlFiles.set(path, doc);
    }

    // We don't extract notes for translation (they're presenter notes, not content)
    // But we keep them in the ZIP untouched

    onProgress?.({ step: 'parse', message: 'XML des slides parsé' });

    // Extract segments from all slides
    const segments = [];
    let segmentId = 0;

    for (const [xmlPath, doc] of xmlFiles) {
        const root = doc.documentElement;
        const shapes = getPptxTextShapes(root);

        for (const shape of shapes) {
            const txBody = getPptxTxBody(shape);
            if (!txBody) continue;

            const paragraphs = getPptxParagraphs(txBody);

            for (const para of paragraphs) {
                if (!isPptxContentParagraph(para)) continue;

                const text = extractTextFromPptxParagraph(para);
                segments.push({
                    id: segmentId++,
                    text: text,
                    xmlPath: xmlPath,
                    _element: para,
                });
            }
        }
    }

    onProgress?.({
        step: 'segments',
        message: `${segments.length} segments de texte extraits de ${slideCount} slides`,
    });

    return { zip, xmlFiles, segments, slideCount };
}

/**
 * Get PPTX document metadata for display.
 */
export function getPptxInfo(xmlFiles, slideCount) {
    const info = {
        slideCount: slideCount,
        shapeCount: 0,
        hasNotes: false,
    };

    for (const [, doc] of xmlFiles) {
        const shapes = getPptxTextShapes(doc.documentElement);
        info.shapeCount += shapes.length;
    }

    return info;
}
