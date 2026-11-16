package me.muban.docx;

import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * Run-level helpers for DOCX paragraphs.
 *
 * <p>Word splits text across runs for reasons a template designer never sees: spell-checking,
 * proofing marks, editing history, character formatting. Structural markers such as
 * {@code #{if ...}} therefore end up scattered across several adjacent runs, which defeats any
 * paragraph-level pattern matching. The helpers here put such markers back together.
 *
 * <p>Everything in this class operates on the XML content list in place and is limited to
 * <em>layout</em> concerns — nothing here knows about expressions, the evaluation context, or the
 * document structure above the paragraph.
 */
final class DocxRunUtils {

    private static final Logger log = LoggerFactory.getLogger(DocxRunUtils.class);

    private DocxRunUtils() {}

    /**
     * Merge adjacent runs that together spell a marker delimited by {@code openingToken} and
     * {@code closingToken}.
     *
     * <p>Only runs with identical formatting are merged. A merge attempt starts at every run whose
     * text contains a character of {@code openingToken} and stops as soon as the accumulated text
     * contains both tokens, so a paragraph may hold several split markers.
     *
     * <p>Example with {@code "#{"} / {@code "}"}:
     * <pre>
     * Before: [w:r]"#"  [w:r]"{if show}"
     * After:  [w:r]"#{if show}"
     * </pre>
     *
     * @param paragraph    the paragraph to normalise (modified in place)
     * @param openingToken marker opening, for example {@code "#{"}
     * @param closingToken marker closing, for example {@code "}"}
     */
    static void joinSplitMarker(P paragraph, String openingToken, String closingToken) {
        List<Object> content = paragraph.getContent();

        int i = 0;
        while (i < content.size()) {
            Object unwrapped = DocxXmlUtils.unwrap(content.get(i));
            if (!(unwrapped instanceof R run)) {
                i++;
                continue;
            }

            String text = DocxXmlUtils.getRunText(run);
            if (!containsAny(text, openingToken)) {
                i++;
                continue;
            }

            StringBuilder accumulated = new StringBuilder(text);
            int mergeEnd = i;

            for (int j = i + 1; j < content.size(); j++) {
                Object nextUnwrapped = DocxXmlUtils.unwrap(content.get(j));
                if (!(nextUnwrapped instanceof R nextRun)) break;

                if (!isCompatibleFormatting(run, nextRun)) break;

                accumulated.append(DocxXmlUtils.getRunText(nextRun));
                mergeEnd = j;

                String acc = accumulated.toString();
                if (acc.contains(openingToken) && acc.contains(closingToken)) {
                    DocxXmlUtils.setRunText(run, acc);
                    for (int k = mergeEnd; k > i; k--) {
                        content.remove(k);
                    }
                    break;
                }
            }

            i++;
        }
    }

    /**
     * Check whether two runs carry the same formatting, by comparing their serialised
     * {@code w:rPr} XML.
     *
     * <p>Missing formatting on both sides counts as compatible; missing on one side only does not,
     * because merging would then apply formatting to text that had none.
     */
    static boolean isCompatibleFormatting(R run1, R run2) {
        var rpr1 = run1.getRPr();
        var rpr2 = run2.getRPr();

        if (rpr1 == null && rpr2 == null) {
            return true;
        }
        if (rpr1 == null || rpr2 == null) {
            return false;
        }

        try {
            String xml1 = org.docx4j.XmlUtils.marshaltoString(rpr1);
            String xml2 = org.docx4j.XmlUtils.marshaltoString(rpr2);
            return xml1.equals(xml2);
        } catch (Exception e) {
            log.debug("Could not compare RPr for run merging: {}", e.getMessage());
            return false;
        }
    }

    private static boolean containsAny(String text, String chars) {
        for (int i = 0; i < chars.length(); i++) {
            if (text.indexOf(chars.charAt(i)) >= 0) {
                return true;
            }
        }
        return false;
    }
}
