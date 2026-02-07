/**
 * Text Matcher — Maps translated text segments to source segments.
 *
 * Multi-level matching:
 * 1. Structure match (same paragraph count → 1:1)
 * 2. Sentence-level split & sequential match
 * 3. Fuzzy matching with anchors
 * 4. Manual fallback
 */

/**
 * Normalize text for comparison: collapse whitespace, trim.
 */
function normalize(text) {
    return text.replace(/\s+/g, ' ').trim();
}

/**
 * Split text into paragraphs (by newline).
 */
function splitIntoParagraphs(text) {
    return text
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter((line) => line.length > 0);
}

/**
 * Compute Levenshtein similarity ratio (0..1) between two strings.
 */
function similarity(a, b) {
    if (a === b) return 1;
    if (!a || !b) return 0;

    const la = a.length;
    const lb = b.length;

    // Quick length-based heuristic for very different strings
    if (Math.abs(la - lb) / Math.max(la, lb) > 0.8) return 0;

    // Use a simpler metric for long strings to avoid O(n²)
    if (la > 500 || lb > 500) {
        // Word overlap ratio
        const wordsA = new Set(a.toLowerCase().split(/\s+/));
        const wordsB = new Set(b.toLowerCase().split(/\s+/));
        let common = 0;
        for (const w of wordsA) {
            if (wordsB.has(w)) common++;
        }
        return (2 * common) / (wordsA.size + wordsB.size);
    }

    // Levenshtein distance
    const matrix = [];
    for (let i = 0; i <= la; i++) {
        matrix[i] = [i];
    }
    for (let j = 0; j <= lb; j++) {
        matrix[0][j] = j;
    }
    for (let i = 1; i <= la; i++) {
        for (let j = 1; j <= lb; j++) {
            const cost = a[i - 1] === b[j - 1] ? 0 : 1;
            matrix[i][j] = Math.min(
                matrix[i - 1][j] + 1,
                matrix[i][j - 1] + 1,
                matrix[i - 1][j - 1] + cost
            );
        }
    }

    const maxLen = Math.max(la, lb);
    return 1 - matrix[la][lb] / maxLen;
}

/**
 * Extract "anchors" from text — numbers, dates, proper nouns, emails, URLs.
 * These elements typically don't change during translation.
 */
function extractAnchors(text) {
    const anchors = new Set();

    // Numbers (including decimals)
    const nums = text.match(/\d+([.,]\d+)*/g);
    if (nums) nums.forEach((n) => anchors.add(n));

    // Email addresses
    const emails = text.match(/[\w.-]+@[\w.-]+\.\w+/g);
    if (emails) emails.forEach((e) => anchors.add(e));

    // URLs
    const urls = text.match(/https?:\/\/\S+/g);
    if (urls) urls.forEach((u) => anchors.add(u));

    // Words that look like proper nouns (capitalized, not at sentence start)
    // This is a heuristic — we match mid-sentence capitalized words
    const properNouns = text.match(/(?<=[.!?]\s+\w+\s+)[A-Z][a-zÀ-ÿ]+/g);
    if (properNouns) properNouns.forEach((n) => anchors.add(n));

    return anchors;
}

/**
 * Score anchor overlap between source and translated text.
 */
function anchorScore(sourceText, translatedText) {
    const srcAnchors = extractAnchors(sourceText);
    const tgtAnchors = extractAnchors(translatedText);

    if (srcAnchors.size === 0) return 0;

    let matches = 0;
    for (const a of srcAnchors) {
        if (tgtAnchors.has(a)) matches++;
    }

    return matches / srcAnchors.size;
}

/**
 * Match source segments to translated paragraphs.
 *
 * @param {Array<{id, text}>} sourceSegments — segments from DOCX
 * @param {string} translationText — the full TXT translation
 * @returns {{ mapping: Map<number, string>, unmatched: Array, stats: Object }}
 */
export function matchTexts(sourceSegments, translationText, onProgress) {
    const translatedParagraphs = splitIntoParagraphs(translationText);
    const mapping = new Map();
    const unmatched = [];
    const stats = {
        total: sourceSegments.length,
        matched: 0,
        fuzzy: 0,
        unmatched: 0,
        method: '',
    };

    onProgress?.({
        step: 'match-start',
        message: `${sourceSegments.length} segments source, ${translatedParagraphs.length} paragraphes traduits`,
    });

    // ─── Level 1: Exact structure match (same count → 1:1) ───
    if (sourceSegments.length === translatedParagraphs.length) {
        stats.method = 'structure-1:1';
        for (let i = 0; i < sourceSegments.length; i++) {
            mapping.set(sourceSegments[i].id, translatedParagraphs[i]);
        }
        stats.matched = sourceSegments.length;

        onProgress?.({
            step: 'match-done',
            message: `Correspondance 1:1 parfaite (${stats.matched} segments)`,
        });

        return { mapping, unmatched, stats };
    }

    // ─── Level 2: Sequential matching with flexibility ───
    stats.method = 'sequential-flex';

    // Try to match sequentially, allowing for added/removed paragraphs
    let tIdx = 0;
    const usedTranslations = new Set();

    for (const segment of sourceSegments) {
        const normSource = normalize(segment.text);

        // Look ahead in a window of candidates
        let bestMatch = -1;
        let bestScore = 0;
        const windowSize = Math.min(5, translatedParagraphs.length - tIdx);

        for (let w = 0; w < windowSize; w++) {
            const candidateIdx = tIdx + w;
            if (candidateIdx >= translatedParagraphs.length) break;
            if (usedTranslations.has(candidateIdx)) continue;

            const normTrans = normalize(translatedParagraphs[candidateIdx]);

            // Check anchors first (shared numbers/names)
            const aScore = anchorScore(normSource, normTrans);

            // Position bonus — closer matches are preferred
            const positionBonus = 1 - w * 0.1;

            // Check length ratio (translations should be roughly similar length)
            const lenRatio =
                Math.min(normSource.length, normTrans.length) /
                Math.max(normSource.length, normTrans.length);
            const lenBonus = lenRatio > 0.3 ? lenRatio * 0.3 : 0;

            const totalScore = aScore * 0.4 + positionBonus * 0.3 + lenBonus * 0.3;

            if (totalScore > bestScore) {
                bestScore = totalScore;
                bestMatch = candidateIdx;
            }
        }

        if (bestMatch >= 0 && bestScore > 0.2) {
            mapping.set(segment.id, translatedParagraphs[bestMatch]);
            usedTranslations.add(bestMatch);
            tIdx = bestMatch + 1;
            stats.matched++;

            if (bestScore < 0.6) stats.fuzzy++;
        } else {
            // Try direct position match as fallback
            if (tIdx < translatedParagraphs.length && !usedTranslations.has(tIdx)) {
                mapping.set(segment.id, translatedParagraphs[tIdx]);
                usedTranslations.add(tIdx);
                tIdx++;
                stats.matched++;
                stats.fuzzy++;
            } else {
                unmatched.push(segment);
                stats.unmatched++;
            }
        }
    }

    // ─── Level 3: Try to match remaining unmatched with unused translations ───
    if (unmatched.length > 0) {
        const unusedTranslations = translatedParagraphs
            .map((t, i) => ({ text: t, index: i }))
            .filter((t) => !usedTranslations.has(t.index));

        const stillUnmatched = [];
        for (const segment of unmatched) {
            let bestMatch = null;
            let bestScore = 0;

            for (const candidate of unusedTranslations) {
                const aScore = anchorScore(segment.text, candidate.text);
                const simScore = similarity(
                    normalize(segment.text).substring(0, 100),
                    normalize(candidate.text).substring(0, 100)
                );
                const totalScore = aScore * 0.5 + simScore * 0.5;

                if (totalScore > bestScore && totalScore > 0.3) {
                    bestScore = totalScore;
                    bestMatch = candidate;
                }
            }

            if (bestMatch) {
                mapping.set(segment.id, bestMatch.text);
                unusedTranslations.splice(unusedTranslations.indexOf(bestMatch), 1);
                stats.matched++;
                stats.unmatched--;
                stats.fuzzy++;
            } else {
                stillUnmatched.push(segment);
            }
        }

        // Update unmatched list
        unmatched.length = 0;
        unmatched.push(...stillUnmatched);
    }

    onProgress?.({
        step: 'match-done',
        message: `${stats.matched} segments mappés, ${stats.fuzzy} approximatifs, ${stats.unmatched} non-mappés`,
    });

    return { mapping, unmatched, stats };
}

/**
 * Get a summary of the matching result for display.
 */
export function getMatchSummary(stats) {
    const lines = [];
    lines.push(`📊 Résultat du matching :`);
    lines.push(`  Total segments : ${stats.total}`);
    lines.push(`  ✅ Mappés : ${stats.matched}`);
    if (stats.fuzzy > 0) {
        lines.push(`  ⚠️ Approximatifs : ${stats.fuzzy}`);
    }
    if (stats.unmatched > 0) {
        lines.push(`  ❌ Non-mappés : ${stats.unmatched}`);
    }
    lines.push(`  Méthode : ${stats.method}`);
    return lines.join('\n');
}
