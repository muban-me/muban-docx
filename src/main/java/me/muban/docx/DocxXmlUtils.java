package me.muban.docx;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

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
}
