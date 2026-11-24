package me.muban.docx;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/**
 * Writer for {@code images.json} — the image manifest shipped with an HTML
 * bundle in {@link HtmlImageMode#CID} mode.
 *
 * <p>The manifest maps each {@code cid:} reference in {@code index.html} to the
 * image file in the archive, so the system assembling the email does not have
 * to parse the HTML:</p>
 *
 * <pre>
 * {
 *   "version": 1,
 *   "images": [
 *     {
 *       "contentId": "8f14e45f...1.image1.png&#64;muban",
 *       "path": "index.html_files/image1.png",
 *       "contentType": "image/png",
 *       "size": 1204
 *     }
 *   ]
 * }
 * </pre>
 *
 * <p>{@code contentId} is the Content-ID <b>without</b> angle brackets (the HTML
 * references it as {@code cid:} + value); pass it as {@code Content-ID: <value>}
 * with a matching {@code multipart/related} part. {@code path} is relative to
 * the HTML root (the directory containing {@code index.html}).</p>
 */
public final class HtmlImageManifest {

    /** Manifest file name, written next to {@code index.html}. */
    public static final String FILE_NAME = "images.json";

    /** Manifest schema version. */
    public static final int VERSION = 1;

    /**
     * One image entry.
     *
     * @param contentId   MIME Content-ID without angle brackets
     * @param path        image path relative to the HTML root
     * @param contentType image MIME type, e.g. {@code image/png}
     * @param size        image size in bytes
     */
    public record Entry(String contentId, String path, String contentType, long size) {
    }

    private HtmlImageManifest() {
    }

    /**
     * Writes the manifest, creating parent directories as needed.
     *
     * @param targetFile the manifest file (typically {@code <htmlRoot>/images.json})
     * @param entries    image entries in document order; may be empty
     * @throws IOException if the file cannot be written
     */
    public static void write(Path targetFile, List<Entry> entries) throws IOException {
        StringBuilder json = new StringBuilder(128 + entries.size() * 128);
        json.append("{\n  \"version\": ").append(VERSION).append(",\n  \"images\": [");
        for (int i = 0; i < entries.size(); i++) {
            Entry entry = entries.get(i);
            json.append(i == 0 ? "\n" : ",\n");
            json.append("    {\"contentId\": \"").append(escape(entry.contentId()))
                    .append("\", \"path\": \"").append(escape(entry.path()))
                    .append("\", \"contentType\": \"").append(escape(entry.contentType()))
                    .append("\", \"size\": ").append(entry.size()).append('}');
        }
        if (!entries.isEmpty()) {
            json.append('\n');
        }
        json.append("  ]\n}\n");

        Path parent = targetFile.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        Files.writeString(targetFile, json.toString(), StandardCharsets.UTF_8);
    }

    private static String escape(String value) {
        if (value == null) {
            return "";
        }
        StringBuilder out = new StringBuilder(value.length() + 8);
        for (int i = 0; i < value.length(); i++) {
            char c = value.charAt(i);
            switch (c) {
                case '"' -> out.append("\\\"");
                case '\\' -> out.append("\\\\");
                case '\n' -> out.append("\\n");
                case '\r' -> out.append("\\r");
                case '\t' -> out.append("\\t");
                default -> {
                    if (c < 0x20) {
                        out.append(String.format("\\u%04x", (int) c));
                    } else {
                        out.append(c);
                    }
                }
            }
        }
        return out.toString();
    }
}
