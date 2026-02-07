/**
 * XML Utilities for DOCX manipulation.
 * Provides namespace-aware helpers for parsing and serializing OOXML.
 */

// OOXML namespaces
const NS = {
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    v: 'urn:schemas-microsoft-com:vml',
};

/**
 * Parse an XML string into a DOM Document.
 */
export function parseXml(xmlString) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlString, 'application/xml');
    const err = doc.querySelector('parsererror');
    if (err) {
        throw new Error(`XML parse error: ${err.textContent}`);
    }
    return doc;
}

/**
 * Serialize a DOM Document back to an XML string.
 * Preserves the original XML declaration if present.
 */
export function serializeXml(doc) {
    const serializer = new XMLSerializer();
    let xml = serializer.serializeToString(doc);

    // Ensure XML declaration is present and standalone is preserved
    if (!xml.startsWith('<?xml')) {
        xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
    }

    return xml;
}

/**
 * Get all <w:t> text nodes inside a <w:r> (run) element.
 */
export function getRunTextNodes(runElement) {
    const textNodes = [];
    for (const child of runElement.childNodes) {
        if (child.localName === 't' && child.namespaceURI === NS.w) {
            textNodes.push(child);
        }
    }
    return textNodes;
}

/**
 * Get all <w:r> (run) elements inside a <w:p> (paragraph) element.
 * Excludes runs inside other nested structures like hyperlinks — we handle those separately.
 */
export function getDirectRuns(paragraphElement) {
    const runs = [];
    for (const child of paragraphElement.childNodes) {
        if (child.localName === 'r' && child.namespaceURI === NS.w) {
            runs.push(child);
        }
        // Also handle runs inside hyperlinks <w:hyperlink>
        if (child.localName === 'hyperlink' && child.namespaceURI === NS.w) {
            for (const hChild of child.childNodes) {
                if (hChild.localName === 'r' && hChild.namespaceURI === NS.w) {
                    runs.push(hChild);
                }
            }
        }
    }
    return runs;
}

/**
 * Extract plain text from a paragraph element by reading all its runs' <w:t> nodes.
 */
export function extractTextFromParagraph(paragraphElement) {
    const runs = getDirectRuns(paragraphElement);
    let text = '';
    for (const run of runs) {
        const textNodes = getRunTextNodes(run);
        for (const tn of textNodes) {
            text += tn.textContent || '';
        }
    }
    return text;
}

/**
 * Get all <w:p> paragraph elements from a document body or a container element.
 * Recursively searches inside tables (<w:tc>) as well.
 */
export function getAllParagraphs(containerElement) {
    const paragraphs = [];

    function walk(node) {
        for (const child of node.childNodes) {
            if (child.localName === 'p' && child.namespaceURI === NS.w) {
                paragraphs.push(child);
            }
            // Recurse into table cells, SDT blocks, text boxes, etc.
            if (child.childNodes && child.childNodes.length > 0) {
                if (
                    child.localName === 'tc' || // Table cell
                    child.localName === 'body' || // Document body
                    child.localName === 'tbl' || // Table
                    child.localName === 'tr' || // Table row
                    child.localName === 'sdtContent' || // Structured doc tag content
                    child.localName === 'sdt' || // Structured doc tag
                    child.localName === 'txbxContent' || // Textbox content
                    child.localName === 'hdr' || // Header
                    child.localName === 'ftr' // Footer
                ) {
                    walk(child);
                }
            }
        }
    }

    walk(containerElement);
    return paragraphs;
}

/**
 * Check if a paragraph is "content" (has text) vs structural (empty, page breaks, etc.).
 */
export function isContentParagraph(paragraphElement) {
    const text = extractTextFromParagraph(paragraphElement);
    return text.trim().length > 0;
}

// ══════════════════════════════════════════════════
// PPTX / DrawingML Helpers
// ══════════════════════════════════════════════════

/**
 * Get all shapes (<p:sp>) from a slide that contain a text body (<p:txBody>).
 * Also walks into group shapes (<p:grpSp>).
 */
export function getPptxTextShapes(slideElement) {
    const shapes = [];

    function walk(node) {
        for (const child of node.childNodes) {
            // Direct shape with text body
            if (child.localName === 'sp' && child.namespaceURI === NS.p) {
                // Check if it has a txBody
                for (const sc of child.childNodes) {
                    if (sc.localName === 'txBody' && sc.namespaceURI === NS.p) {
                        shapes.push(child);
                        break;
                    }
                }
            }
            // Group shape — recurse into it
            if (child.localName === 'grpSp' && child.namespaceURI === NS.p) {
                walk(child);
            }
        }
    }

    // Walk the shape tree (<p:cSld> → <p:spTree>)
    for (const child of slideElement.childNodes) {
        if (child.localName === 'cSld' && child.namespaceURI === NS.p) {
            for (const sc of child.childNodes) {
                if (sc.localName === 'spTree' && sc.namespaceURI === NS.p) {
                    walk(sc);
                }
            }
        }
    }

    return shapes;
}

/**
 * Get <p:txBody> from a shape element.
 */
export function getPptxTxBody(shapeElement) {
    for (const child of shapeElement.childNodes) {
        if (child.localName === 'txBody' && child.namespaceURI === NS.p) {
            return child;
        }
    }
    return null;
}

/**
 * Get all <a:p> paragraph elements from a text body.
 */
export function getPptxParagraphs(txBodyElement) {
    const paragraphs = [];
    for (const child of txBodyElement.childNodes) {
        if (child.localName === 'p' && child.namespaceURI === NS.a) {
            paragraphs.push(child);
        }
    }
    return paragraphs;
}

/**
 * Get all <a:r> run elements from a paragraph.
 */
export function getPptxRuns(paragraphElement) {
    const runs = [];
    for (const child of paragraphElement.childNodes) {
        if (child.localName === 'r' && child.namespaceURI === NS.a) {
            runs.push(child);
        }
    }
    return runs;
}

/**
 * Get <a:t> text nodes from a run element.
 */
export function getPptxTextNodes(runElement) {
    const textNodes = [];
    for (const child of runElement.childNodes) {
        if (child.localName === 't' && child.namespaceURI === NS.a) {
            textNodes.push(child);
        }
    }
    return textNodes;
}

/**
 * Extract plain text from a PPTX paragraph by reading all its runs' <a:t> nodes.
 */
export function extractTextFromPptxParagraph(paragraphElement) {
    const runs = getPptxRuns(paragraphElement);
    let text = '';
    for (const run of runs) {
        const textNodes = getPptxTextNodes(run);
        for (const tn of textNodes) {
            text += tn.textContent || '';
        }
    }
    return text;
}

/**
 * Check if a PPTX paragraph has actual text content.
 */
export function isPptxContentParagraph(paragraphElement) {
    const text = extractTextFromPptxParagraph(paragraphElement);
    return text.trim().length > 0;
}

export { NS };
