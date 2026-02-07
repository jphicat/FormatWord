/**
 * DOCX Parser — Extracts text segments from a .docx ZIP archive.
 *
 * Strategy: We don't try to model the DOCX in a custom data structure.
 * Instead, we extract a flat list of "segments" (paragraph‐level text chunks)
 * with enough metadata to map them back to the XML later.
 */

import JSZip from 'jszip';
import {
    parseXml,
    getAllParagraphs,
    extractTextFromParagraph,
    isContentParagraph,
} from './xmlUtils.js';

/**
 * Parse a .docx file (ArrayBuffer) and return parsed data.
 * @param {ArrayBuffer} buffer — The raw .docx bytes
 * @returns {{ zip: JSZip, xmlFiles: Map<string, Document>, segments: Array }}
 */
export async function parseDocx(buffer, onProgress) {
    const zip = await JSZip.loadAsync(buffer);

    onProgress?.({ step: 'unzip', message: 'Archive décompressée' });

    // Identify all XML files that can contain text
    const textXmlPaths = [];

    // Main document
    if (zip.file('word/document.xml')) {
        textXmlPaths.push('word/document.xml');
    }

    // Headers & footers
    for (const path of Object.keys(zip.files)) {
        if (
            /^word\/(header|footer)\d*\.xml$/i.test(path) &&
            !textXmlPaths.includes(path)
        ) {
            textXmlPaths.push(path);
        }
    }

    // Footnotes & endnotes
    if (zip.file('word/footnotes.xml')) textXmlPaths.push('word/footnotes.xml');
    if (zip.file('word/endnotes.xml')) textXmlPaths.push('word/endnotes.xml');

    onProgress?.({
        step: 'identify',
        message: `${textXmlPaths.length} fichiers XML identifiés`,
    });

    // Parse each XML file
    const xmlFiles = new Map();
    for (const path of textXmlPaths) {
        const xmlString = await zip.file(path).async('string');
        const doc = parseXml(xmlString);
        xmlFiles.set(path, doc);
    }

    onProgress?.({ step: 'parse', message: 'XML parsé' });

    // Extract segments from all XML files
    const segments = [];
    let segmentId = 0;

    for (const [xmlPath, doc] of xmlFiles) {
        const root = doc.documentElement;
        const paragraphs = getAllParagraphs(root);

        for (const para of paragraphs) {
            if (!isContentParagraph(para)) continue;

            const text = extractTextFromParagraph(para);
            segments.push({
                id: segmentId++,
                text: text,
                xmlPath: xmlPath,
                // Store a reference to the DOM element for later replacement
                _element: para,
            });
        }
    }

    onProgress?.({
        step: 'segments',
        message: `${segments.length} segments de texte extraits`,
    });

    return { zip, xmlFiles, segments };
}

/**
 * Get document metadata (page size, margins, etc.) — informational only.
 */
export function getDocumentInfo(xmlFiles) {
    const docXml = xmlFiles.get('word/document.xml');
    if (!docXml) return {};

    const info = {
        paragraphCount: 0,
        hasHeaders: false,
        hasFooters: false,
        hasFootnotes: false,
        hasTables: false,
        hasImages: false,
    };

    // Count paragraphs
    const allParas = getAllParagraphs(docXml.documentElement);
    info.paragraphCount = allParas.filter(isContentParagraph).length;

    // Check features
    for (const path of xmlFiles.keys()) {
        if (path.includes('header')) info.hasHeaders = true;
        if (path.includes('footer')) info.hasFooters = true;
        if (path.includes('footnote')) info.hasFootnotes = true;
    }

    // Check for tables
    const tables = docXml.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'tbl'
    );
    info.hasTables = tables.length > 0;

    // Check for images
    const drawings = docXml.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'drawing'
    );
    if (drawings.length === 0) {
        const pics = docXml.getElementsByTagNameNS(
            'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'pic'
        );
        info.hasImages = pics.length > 0;
    } else {
        info.hasImages = true;
    }

    return info;
}
