/**
 * Text Matcher — Maps translated text segments to source segments.
 *
 * Professional Translator Profile:
 * A professional translator preserves the EXACT paragraph structure of the source.
 * Paragraph N in the source = Paragraph N in the translation.
 * This is the golden rule. When counts differ, we handle edge cases gracefully.
 *
 * Strategy:
 * 1. PRIMARY: Strict 1:1 mapping by paragraph order (translator preserves structure)
 * 2. FALLBACK: If translation has fewer paragraphs, map what we can and keep rest original
 * 3. FALLBACK: If translation has more paragraphs, merge overflow into last matched segment
 */

/**
 * Split text into paragraphs (by newline), keeping non-empty lines.
 */
function splitIntoParagraphs(text) {
    return text
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter((line) => line.length > 0);
}

/**
 * Match source segments to translated paragraphs.
 *
 * @param {Array<{id, text}>} sourceSegments — segments from DOCX
 * @param {string} translationText — the full TXT translation
 * @param {Function} onProgress — progress callback
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

    const srcCount = sourceSegments.length;
    const trgCount = translatedParagraphs.length;

    // ─── Case 1: Same count → perfect 1:1 mapping ───
    if (srcCount === trgCount) {
        stats.method = 'structure-1:1';
        for (let i = 0; i < srcCount; i++) {
            mapping.set(sourceSegments[i].id, translatedParagraphs[i]);
        }
        stats.matched = srcCount;

        onProgress?.({
            step: 'match-done',
            message: `Correspondance 1:1 parfaite (${stats.matched} segments)`,
        });

        return { mapping, unmatched, stats };
    }

    // ─── Case 2: More source than translation → map in order, rest stays as-is ───
    if (srcCount > trgCount) {
        stats.method = 'sequential-partial';

        for (let i = 0; i < srcCount; i++) {
            if (i < trgCount) {
                mapping.set(sourceSegments[i].id, translatedParagraphs[i]);
                stats.matched++;
            } else {
                // Keep the original text (no translation available for this segment)
                unmatched.push(sourceSegments[i]);
                stats.unmatched++;
            }
        }

        onProgress?.({
            step: 'match-done',
            message: `${stats.matched} segments mappés, ${stats.unmatched} sans traduction (conservés à l'original)`,
        });

        return { mapping, unmatched, stats };
    }

    // ─── Case 3: More translation than source → map 1:1, merge overflow ───
    stats.method = 'sequential-merge';

    for (let i = 0; i < srcCount; i++) {
        if (i < srcCount - 1) {
            // Not the last segment: direct 1:1 mapping
            mapping.set(sourceSegments[i].id, translatedParagraphs[i]);
        } else {
            // Last source segment: absorb all remaining translated paragraphs
            const remaining = translatedParagraphs.slice(i).join(' ');
            mapping.set(sourceSegments[i].id, remaining);
        }
        stats.matched++;
    }

    stats.fuzzy = trgCount - srcCount; // Number of extra paragraphs merged

    onProgress?.({
        step: 'match-done',
        message: `${stats.matched} segments mappés (${stats.fuzzy} paragraphes excédentaires fusionnés dans le dernier segment)`,
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
