package me.muban.docx;

import org.docx4j.XmlUtils;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Br;
import org.docx4j.wml.CTVerticalAlignRun;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STVerticalAlignRun;
import org.docx4j.wml.U;
import org.docx4j.wml.UnderlineEnumeration;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Converts a small, whitelisted subset of HTML into WordprocessingML runs.
 *
 * <p>Content management systems (e.g. Higson) store rich business text as HTML
 * fragments: legal footers use {@code <br/>}, clauses use {@code <b>}, and
 * statutory notes use {@code <sup>(1)</sup>}. Substituting such a value with
 * {@link DocxXmlUtils#setRunText(R, String)} would print the markup literally
 * and collapse line breaks into spaces. This class turns the fragment into a
 * sequence of {@link R} elements that Word and Apache FOP render correctly.
 *
 * <p><b>Supported subset:</b>
 *
 * <table border="1">
 *   <caption>Supported HTML tags</caption>
 *   <tr><th>HTML</th><th>Effect</th></tr>
 *   <tr><td>{@code <br>}, {@code <br/>}</td><td>line break ({@code w:br})</td></tr>
 *   <tr><td>{@code <b>}, {@code <strong>}</td><td>bold ({@code w:b})</td></tr>
 *   <tr><td>{@code <i>}, {@code <em>}</td><td>italic ({@code w:i})</td></tr>
 *   <tr><td>{@code <u>}</td><td>single underline ({@code w:u})</td></tr>
 *   <tr><td>{@code <sup>}, {@code <sub>}</td><td>vertical alignment ({@code w:vertAlign})</td></tr>
 *   <tr><td>{@code <p>}, {@code <div>}, {@code <li>}, {@code <tr>}</td><td>paragraph boundary (line break)</td></tr>
 *   <tr><td>{@code &amp;amp; &amp;lt; &amp;gt; &amp;quot; &amp;apos; &amp;nbsp;}
 *       and numeric entities</td><td>decoded to characters</td></tr>
 * </table>
 *
 * <p><b>Deliberately not supported:</b> CSS and inline styles, hyperlinks,
 * images, tables and nested lists. Formatting must come from the Word template,
 * not from the data — otherwise the content system starts controlling the
 * document's visual design. Images have their own channel ({@code image:key}
 * alt-text replacement in {@link DocxImageReplacer}).
 *
 * <p><b>Robustness:</b> malformed input never throws. Unknown tags are stripped
 * (inner text preserved), unbalanced tags cannot escape the call, and values
 * longer than {@link #MAX_LENGTH} fall back to plain text with tags removed.
 * Nested formatting is tracked with counters, so {@code <b><i>x</i></b>} yields
 * a single run with both attributes.
 *
 * @see MubanDocxEngine
 */
public final class DocxRichText {

    private static final Logger log = LoggerFactory.getLogger(DocxRichText.class);

    /** Maximum accepted input length; longer values fall back to plain text. */
    public static final int MAX_LENGTH = 20_000;

    /** Maximum length of an entity name/number between {@code &} and {@code ;}. */
    private static final int MAX_ENTITY_LENGTH = 10;

    private static final ObjectFactory WML = new ObjectFactory();

    private DocxRichText() {
    }

    /**
     * Convert an HTML fragment into a list of runs.
     *
     * <p>The returned runs are ready to be inserted into a {@code w:p} content list.
     * Formatting is based on {@code baseRPr}; inline HTML toggles bold, italic,
     * underline and vertical alignment on top of it, so the surrounding template
     * style is never lost.
     *
     * @param baseRPr formatting of the run that contained the placeholder (may be {@code null})
     * @param html    HTML fragment (may be {@code null})
     * @return runs in document order; never {@code null}, possibly empty
     */
    public static List<R> toRuns(RPr baseRPr, String html) {
        if (html == null || html.isEmpty()) {
            return List.of();
        }
        if (html.length() > MAX_LENGTH) {
            log.warn("Rich text value is {} characters long (limit {}); falling back to plain text",
                    html.length(), MAX_LENGTH);
            return List.of(DocxXmlUtils.textRun(baseRPr, stripTags(html)));
        }

        List<R> runs = new ArrayList<>();
        StringBuilder buffer = new StringBuilder();
        StyleState style = new StyleState();

        int i = 0;
        int n = html.length();
        while (i < n) {
            char c = html.charAt(i);

            if (c == '<') {
                int gt = findTagEnd(html, i);
                if (gt < 0) {
                    buffer.append(html, i, n);
                    break;
                }
                String tag = html.substring(i + 1, gt).trim();
                i = gt + 1;
                applyTag(tag, buffer, runs, style, baseRPr);
            } else if (c == '&') {
                int semi = html.indexOf(';', i + 1);
                if (semi > i && semi - i <= MAX_ENTITY_LENGTH) {
                    String decoded = decodeEntity(html.substring(i + 1, semi));
                    if (decoded != null) {
                        buffer.append(decoded);
                        i = semi + 1;
                        continue;
                    }
                }
                buffer.append(c);
                i++;
            } else {
                buffer.append(c);
                i++;
            }
        }

        flush(buffer, runs, style, baseRPr);
        trimTrailingBreak(runs);
        return runs;
    }

    /**
     * Remove all tags from a value and decode entities — safe fallback.
     *
     * @param html HTML fragment (may be {@code null})
     * @return plain text without markup
     */
    public static String stripTags(String html) {
        if (html == null || html.isEmpty()) {
            return "";
        }
        StringBuilder out = new StringBuilder(html.length());
        int i = 0;
        int n = html.length();
        while (i < n) {
            char c = html.charAt(i);
            if (c == '<') {
                int gt = findTagEnd(html, i);
                if (gt < 0) {
                    out.append(html, i, n);
                    break;
                }
                String tag = html.substring(i + 1, gt).trim();
                if (!tag.isEmpty() && !tag.startsWith("/")) {
                    String name = tagName(stripTrailingSlash(tag));
                    if ("br".equals(name) || "p".equals(name) || "div".equals(name)) {
                        out.append(' ');
                    }
                }
                i = gt + 1;
            } else if (c == '&') {
                int semi = html.indexOf(';', i + 1);
                if (semi > i && semi - i <= MAX_ENTITY_LENGTH) {
                    String decoded = decodeEntity(html.substring(i + 1, semi));
                    if (decoded != null) {
                        out.append(decoded);
                        i = semi + 1;
                        continue;
                    }
                }
                out.append(c);
                i++;
            } else {
                out.append(c);
                i++;
            }
        }
        return out.toString();
    }

    // ==================== TAG HANDLING ====================

    private static void applyTag(String rawTag, StringBuilder buffer, List<R> runs,
                                 StyleState style, RPr baseRPr) {
        if (rawTag.isEmpty()) {
            return;
        }
        boolean closing = rawTag.charAt(0) == '/';
        String body = closing ? rawTag.substring(1).trim() : rawTag;
        if (body.endsWith("/")) {
            body = body.substring(0, body.length() - 1).trim();
        }
        String name = tagName(body);

        switch (name) {
            case "br" -> {
                flush(buffer, runs, style, baseRPr);
                addBreak(runs, baseRPr, true);
            }
            case "p", "div", "li", "tr" -> paragraphBoundary(buffer, runs, style, baseRPr);
            case "b", "strong" -> {
                flush(buffer, runs, style, baseRPr);
                style.bold = toggle(style.bold, closing);
            }
            case "i", "em" -> {
                flush(buffer, runs, style, baseRPr);
                style.italic = toggle(style.italic, closing);
            }
            case "u" -> {
                flush(buffer, runs, style, baseRPr);
                style.underline = toggle(style.underline, closing);
            }
            case "sup" -> {
                flush(buffer, runs, style, baseRPr);
                style.superscript = toggle(style.superscript, closing);
            }
            case "sub" -> {
                flush(buffer, runs, style, baseRPr);
                style.subscript = toggle(style.subscript, closing);
            }
            default -> {
                // Unknown or unsupported tag (span, a, table, ...) — strip, keep inner text.
            }
        }
    }

    private static void paragraphBoundary(StringBuilder buffer, List<R> runs,
                                          StyleState style, RPr baseRPr) {
        flush(buffer, runs, style, baseRPr);
        addBreak(runs, baseRPr, false);
    }

    private static int toggle(int counter, boolean closing) {
        return closing ? Math.max(0, counter - 1) : counter + 1;
    }

    private static String tagName(String tagBody) {
        int end = 0;
        while (end < tagBody.length() && !Character.isWhitespace(tagBody.charAt(end))) {
            end++;
        }
        return tagBody.substring(0, end).toLowerCase(Locale.ROOT);
    }

    private static String stripTrailingSlash(String tagBody) {
        String trimmed = tagBody.trim();
        return trimmed.endsWith("/")
                ? trimmed.substring(0, trimmed.length() - 1).trim()
                : trimmed;
    }

    /**
     * Find the closing {@code >} of a tag, tolerating HTML comments.
     */
    private static int findTagEnd(String html, int ltIndex) {
        if (html.startsWith("<!--", ltIndex)) {
            int end = html.indexOf("-->", ltIndex + 4);
            return end < 0 ? -1 : end + 2;
        }
        return html.indexOf('>', ltIndex + 1);
    }

    // ==================== RUN BUILDING ====================

    private static void flush(StringBuilder buffer, List<R> runs, StyleState style, RPr baseRPr) {
        if (buffer.isEmpty()) {
            return;
        }
        String text = buffer.toString();
        buffer.setLength(0);
        runs.add(DocxXmlUtils.buildRun(effectiveRPr(baseRPr, style), List.<Object>of(textElement(text))));
    }

    /**
     * Append a line break run.
     *
     * @param force {@code true} for {@code <br>} (a blank line is intentional),
     *              {@code false} for paragraph boundaries (collapse consecutive breaks)
     */
    private static void addBreak(List<R> runs, RPr baseRPr, boolean force) {
        if (runs.isEmpty()) {
            return;
        }
        if (!force && isBreakOnly(runs.get(runs.size() - 1))) {
            return;
        }
        runs.add(DocxXmlUtils.buildRun(baseRPr, List.<Object>of(WML.createBr())));
    }

    private static void trimTrailingBreak(List<R> runs) {
        while (!runs.isEmpty() && isBreakOnly(runs.get(runs.size() - 1))) {
            runs.remove(runs.size() - 1);
        }
    }

    private static boolean isBreakOnly(R run) {
        if (run.getContent().size() != 1) {
            return false;
        }
        return DocxXmlUtils.unwrap(run.getContent().get(0)) instanceof Br;
    }

    private static org.docx4j.wml.Text textElement(String value) {
        org.docx4j.wml.Text text = new org.docx4j.wml.Text();
        text.setValue(value);
        text.setSpace("preserve");
        return text;
    }

    private static RPr effectiveRPr(RPr baseRPr, StyleState style) {
        if (!style.isActive()) {
            return baseRPr == null ? null : XmlUtils.deepCopy(baseRPr);
        }
        RPr rPr = baseRPr == null ? WML.createRPr() : XmlUtils.deepCopy(baseRPr);
        if (style.bold > 0) {
            BooleanDefaultTrue value = new BooleanDefaultTrue();
            value.setVal(Boolean.TRUE);
            rPr.setB(value);
        }
        if (style.italic > 0) {
            BooleanDefaultTrue value = new BooleanDefaultTrue();
            value.setVal(Boolean.TRUE);
            rPr.setI(value);
        }
        if (style.underline > 0) {
            U underline = WML.createU();
            underline.setVal(UnderlineEnumeration.SINGLE);
            rPr.setU(underline);
        }
        boolean sup = style.superscript > 0;
        boolean sub = style.subscript > 0;
        if (sup != sub) {
            CTVerticalAlignRun vertical = new CTVerticalAlignRun();
            vertical.setVal(sup ? STVerticalAlignRun.SUPERSCRIPT : STVerticalAlignRun.SUBSCRIPT);
            rPr.setVertAlign(vertical);
        }
        return rPr;
    }

    // ==================== ENTITIES ====================

    private static String decodeEntity(String entity) {
        if (entity.isEmpty()) {
            return null;
        }
        if (entity.charAt(0) == '#') {
            return decodeNumericEntity(entity);
        }
        return switch (entity.toLowerCase(Locale.ROOT)) {
            case "amp" -> "&";
            case "lt" -> "<";
            case "gt" -> ">";
            case "quot" -> "\"";
            case "apos" -> "'";
            case "nbsp" -> "\u00A0";
            case "ndash" -> "\u2013";
            case "mdash" -> "\u2014";
            case "bquot", "bdquo" -> "\u201E";
            case "ldquo" -> "\u201C";
            case "rdquo" -> "\u201D";
            case "lsquo" -> "\u2018";
            case "rsquo" -> "\u2019";
            case "hellip" -> "\u2026";
            case "copy" -> "\u00A9";
            case "reg" -> "\u00AE";
            case "trade" -> "\u2122";
            default -> null;
        };
    }

    private static String decodeNumericEntity(String entity) {
        try {
            int codePoint;
            if (entity.length() > 1 && (entity.charAt(1) == 'x' || entity.charAt(1) == 'X')) {
                codePoint = Integer.parseInt(entity.substring(2), 16);
            } else {
                codePoint = Integer.parseInt(entity.substring(1), 10);
            }
            if (codePoint <= 0 || codePoint > Character.MAX_CODE_POINT) {
                return null;
            }
            return new String(Character.toChars(codePoint));
        } catch (NumberFormatException e) {
            return null;
        }
    }

    // ==================== STATE ====================

    /**
     * Nesting counters for inline formatting. Counters (rather than booleans)
     * keep {@code <b>a<b>b</b>c</b>} balanced even when a closing tag is missing.
     */
    private static final class StyleState {
        private int bold;
        private int italic;
        private int underline;
        private int superscript;
        private int subscript;

        private boolean isActive() {
            return bold > 0 || italic > 0 || underline > 0 || superscript > 0 || subscript > 0;
        }
    }
}
