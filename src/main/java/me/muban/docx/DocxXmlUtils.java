package me.muban.docx;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.XmlUtils;
import org.docx4j.wml.Br;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Text;

import java.util.ArrayList;
import java.util.List;

/**
 * Low-level XML utilities for navigating and manipulating DOCX (OOXML) content trees.
 *
 * <p>Provides helpers for unwrapping JAXB elements, reading/writing run text,
 * and extracting concatenated text from arbitrary content hierarchies.
 */
public final class DocxXmlUtils {

    private DocxXmlUtils() {}

    /**
     * Unwrap a potential {@link JAXBElement} to get the underlying value.
     * If the object is not a JAXBElement, it is returned as-is.
     */
    public static Object unwrap(Object obj) {
        if (obj instanceof JAXBElement) {
            return ((JAXBElement<?>) obj).getValue();
        }
        return obj;
    }

    /**
     * Extract the concatenated text content from a run element.
     */
    public static String getRunText(R run) {
        StringBuilder sb = new StringBuilder();
        for (Object content : run.getContent()) {
            Object unwrapped = unwrap(content);
            if (unwrapped instanceof Text text) {
                sb.append(text.getValue());
            }
        }
        return sb.toString();
    }

    /**
     * Replace all text content in a run with the given string.
     * Removes all existing {@link Text} children and adds a single new one.
     */
    public static void setRunText(R run, String text) {
        run.getContent().removeIf(obj -> unwrap(obj) instanceof Text);
        if (text != null && !text.isEmpty()) {
            Text t = new Text();
            t.setValue(text);
            t.setSpace("preserve");
            run.getContent().add(t);
        }
    }

    /**
     * Recursively extract all text from a content tree.
     * Handles {@link Text}, {@link R} (runs), and any {@link ContentAccessor} (paragraphs, cells, etc.).
     */
    public static String extractAllText(List<Object> content) {
        StringBuilder sb = new StringBuilder();
        for (Object obj : content) {
            Object unwrapped = unwrap(obj);
            if (unwrapped instanceof Text text) {
                sb.append(text.getValue());
            } else if (unwrapped instanceof R run) {
                sb.append(getRunText(run));
            } else if (unwrapped instanceof ContentAccessor accessor) {
                sb.append(extractAllText(accessor.getContent()));
            }
        }
        return sb.toString();
    }

    /**
     * Split runs that contain {@code w:br} (soft line break) elements into
     * separate runs so that each run has either text-only or a break-only content.
     *
     * <p>This must be called <b>before</b> placeholder substitution because
     * {@link #getRunText(R)} / {@link #setRunText(R, String)} only handle
     * {@link Text} children and would orphan any {@code w:br} siblings,
     * causing line breaks to shift position in the rendered output.
     *
     * <p>Example transformation:
     * <pre>
     * Before: &lt;w:r&gt;&lt;w:t&gt;Line 1&lt;/w:t&gt;&lt;w:br/&gt;&lt;w:t&gt;${name}&lt;/w:t&gt;&lt;/w:r&gt;
     * After:  &lt;w:r&gt;&lt;w:t&gt;Line 1&lt;/w:t&gt;&lt;/w:r&gt;
     *         &lt;w:r&gt;&lt;w:br/&gt;&lt;/w:r&gt;
     *         &lt;w:r&gt;&lt;w:t&gt;${name}&lt;/w:t&gt;&lt;/w:r&gt;
     * </pre>
     *
     * Each new run inherits the original run's {@code w:rPr} (formatting).
     *
     * @param paragraph the paragraph whose runs should be normalized
     */
    public static void splitRunsAtBreaks(P paragraph) {
        List<Object> original = paragraph.getContent();
        List<Object> result = new ArrayList<>(original.size());
        boolean changed = false;

        for (Object obj : original) {
            Object unwrapped = unwrap(obj);
            if (!(unwrapped instanceof R run) || !containsBreak(run)) {
                result.add(obj);
                continue;
            }

            // This run has at least one w:br — split it
            changed = true;
            RPr rPr = run.getRPr();
            List<Object> currentTextParts = new ArrayList<>();

            for (Object child : run.getContent()) {
                Object unwrappedChild = unwrap(child);
                if (unwrappedChild instanceof Br) {
                    // Flush accumulated text parts into a run (if any)
                    if (!currentTextParts.isEmpty()) {
                        result.add(buildRun(rPr, currentTextParts));
                        currentTextParts = new ArrayList<>();
                    }
                    // Create a break-only run
                    result.add(buildRun(rPr, List.of(child)));
                } else {
                    currentTextParts.add(child);
                }
            }
            // Flush remaining text parts
            if (!currentTextParts.isEmpty()) {
                result.add(buildRun(rPr, currentTextParts));
            }
        }

        if (changed) {
            original.clear();
            original.addAll(result);
        }
    }

    private static boolean containsBreak(R run) {
        for (Object child : run.getContent()) {
            if (unwrap(child) instanceof Br) return true;
        }
        return false;
    }

    private static R buildRun(RPr rPr, List<Object> content) {
        R newRun = new R();
        if (rPr != null) {
            newRun.setRPr(XmlUtils.deepCopy(rPr));
        }
        newRun.getContent().addAll(content);
        return newRun;
    }
}
