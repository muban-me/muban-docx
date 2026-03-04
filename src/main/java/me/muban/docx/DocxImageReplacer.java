package me.muban.docx;

import org.docx4j.dml.CTBlip;
import org.docx4j.dml.CTNonVisualDrawingProps;
import org.docx4j.dml.Graphic;
import org.docx4j.dml.wordprocessingDrawing.Anchor;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.Base64;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Replaces placeholder images in DOCX templates with dynamic content.
 *
 * <p>Images are identified by a naming convention in their alt text (description):
 * the alt text must start with {@code image:} followed by a key name. For example,
 * an image with alt text {@code image:facsimile} is a placeholder for the "facsimile" image.
 *
 * <p><b>Expression support:</b></p>
 * <p>The key (text after {@code image:}) may contain {@code ${...}} SpEL expressions which are
 * evaluated against the raw context before the resolution cascade runs. This allows template
 * designers to embed conditional logic directly in the alt text:</p>
 * <pre>
 * image:${gender == 'F' ? 'assets/female.png' : 'assets/male.png'}
 * image:${risk &gt; 80 ? 'assets/exclamation.png' : 'assets/info.png'}
 * image:assets/${department}/stamp.png
 * </pre>
 *
 * <p><b>Resolution cascade (highest priority first):</b></p>
 * <ol>
 *   <li><b>Inline Base64</b> — context value starting with {@code data:image/} is decoded as a data URI</li>
 *   <li><b>URL</b> — context value starting with {@code http://} or {@code https://} is fetched via HTTP GET</li>
 *   <li><b>Relative path (from context)</b> — any other string value is treated as a relative path within the template package</li>
 *   <li><b>Key as path</b> — if no context value, the key itself is used as a relative path within the template package</li>
 *   <li><b>No replacement</b> — keeps the original embedded image unchanged</li>
 * </ol>
 *
 * <h3>Security:</h3>
 * <p>Relative paths are validated against directory traversal — resolved paths must stay
 * within the template's cache directory.
 *
 * @see DocxExpressionEvaluator
 */
public final class DocxImageReplacer {

    private static final Logger log = LoggerFactory.getLogger(DocxImageReplacer.class);

    /** Alt text prefix that marks an image as a dynamic placeholder */
    static final String IMAGE_PREFIX = "image:";

    /** Data URI prefix for inline Base64 images */
    private static final String DATA_URI_PREFIX = "data:image/";

    /** Pattern to parse data URI: data:image/{type};base64,{payload} */
    private static final Pattern DATA_URI_PATTERN =
            Pattern.compile("^data:image/[^;]+;base64,(.+)$", Pattern.DOTALL);

    /** HTTP/HTTPS URL prefixes */
    private static final String HTTP_PREFIX = "http://";
    private static final String HTTPS_PREFIX = "https://";

    /** Connection timeout for URL image fetching */
    static final Duration URL_CONNECT_TIMEOUT = Duration.ofSeconds(10);

    /** Read/request timeout for URL image fetching */
    static final Duration URL_REQUEST_TIMEOUT = Duration.ofSeconds(10);

    /** Maximum allowed response body size (10 MB) */
    static final long URL_MAX_RESPONSE_BYTES = 10L * 1024 * 1024;

    /** Allowed content types for URL-fetched images */
    private static final Set<String> ALLOWED_CONTENT_TYPES = Set.of(
            "image/png", "image/jpeg", "image/gif", "image/bmp",
            "image/tiff", "image/webp", "image/svg+xml"
    );

    /** Pattern to detect SpEL expressions: ${...} */
    private static final Pattern SPEL_PATTERN = Pattern.compile("\\$\\{(.+?)}");

    /** Lazily initialized shared HTTP client */
    private static volatile HttpClient httpClient;

    private DocxImageReplacer() {}

    /**
     * Replace placeholder images in the DOCX document (without locale).
     *
     * @param wordPackage the loaded DOCX document
     * @param rawContext  raw context map (parameters + data) for image value lookup
     * @param cacheDir    template cache directory (extracted ZIP root) for asset resolution
     * @return the number of images successfully replaced
     */
    public static int replaceImages(WordprocessingMLPackage wordPackage,
                                     Map<String, Object> rawContext,
                                     File cacheDir) {
        return replaceImages(wordPackage, rawContext, cacheDir, null);
    }

    /**
     * Replace placeholder images in the DOCX document.
     *
     * <p>Scans the document body, headers, and footers for images with alt text
     * matching the {@code image:{key}} convention. For each match, resolves the
     * replacement image bytes and swaps the binary content of the image part.
     *
     * @param wordPackage the loaded DOCX document
     * @param rawContext  raw context map (parameters + data) for image value lookup
     * @param cacheDir    template cache directory (extracted ZIP root) for asset resolution
     * @param locale      optional locale for SpEL expression evaluation, or null
     * @return the number of images successfully replaced
     */
    public static int replaceImages(WordprocessingMLPackage wordPackage,
                                     Map<String, Object> rawContext,
                                     File cacheDir,
                                     Locale locale) {
        int replacedCount = 0;
        MainDocumentPart mainPart = wordPackage.getMainDocumentPart();

        // Process main document body
        replacedCount += processImagesInContent(mainPart.getContent(), mainPart, rawContext, cacheDir, locale);

        // Process headers and footers
        try {
            if (wordPackage.getDocumentModel() != null) {
                for (var section : wordPackage.getDocumentModel().getSections()) {
                    var hfp = section.getHeaderFooterPolicy();
                    if (hfp == null) continue;
                    replacedCount += processHeaderFooterImages(hfp.getDefaultHeader(), rawContext, cacheDir, locale);
                    replacedCount += processHeaderFooterImages(hfp.getDefaultFooter(), rawContext, cacheDir, locale);
                    replacedCount += processHeaderFooterImages(hfp.getFirstHeader(), rawContext, cacheDir, locale);
                    replacedCount += processHeaderFooterImages(hfp.getFirstFooter(), rawContext, cacheDir, locale);
                    replacedCount += processHeaderFooterImages(hfp.getEvenHeader(), rawContext, cacheDir, locale);
                    replacedCount += processHeaderFooterImages(hfp.getEvenFooter(), rawContext, cacheDir, locale);
                }
            }
        } catch (Exception e) {
            log.debug("Could not process headers/footers for image replacement: {}", e.getMessage());
        }

        if (replacedCount > 0) {
            log.info("Replaced {} placeholder image(s) in DOCX template", replacedCount);
        }

        return replacedCount;
    }

    // ==================== IMAGE KEY EXTRACTION (TEMPLATE SCANNING) ====================

    /**
     * Extract image placeholder keys from a DOCX template.
     *
     * <p>Scans all images in the document (body, headers, footers) for alt text matching
     * the {@code image:{key}} convention. Returns the set of discovered image keys.
     *
     * <p><b>Extraction rules:</b></p>
     * <ul>
     *   <li>{@code image:facsimile} — direct key → extracts {@code "facsimile"}</li>
     *   <li>{@code image:${photo}} — single simple variable → extracts {@code "photo"}</li>
     *   <li>{@code image:${items.photo}} — single dot-notation variable → extracts {@code "items.photo"}</li>
     *   <li>{@code image:${gender == 'F' ? 'a.png' : 'b.png'}} — complex expression → skipped
     *       (use {@link #extractImageExpressionVariables} to get referenced variables)</li>
     *   <li>{@code image:assets/${dept}/stamp.png} — mixed literal + expression → skipped</li>
     * </ul>
     *
     * @param wordPackage the loaded DOCX template
     * @return set of image keys (without the {@code image:} prefix), in discovery order
     */
    public static Set<String> extractImageKeys(WordprocessingMLPackage wordPackage) {
        Set<String> imageKeys = new java.util.LinkedHashSet<>();
        scanImageAltTexts(wordPackage, (imageKey) -> {
            if (!imageKey.contains("${")) {
                // Direct key: image:facsimile
                imageKeys.add(imageKey);
            } else {
                // Check for single ${simpleKey} pattern (entire key is one placeholder)
                Matcher m = SPEL_PATTERN.matcher(imageKey);
                if (m.matches()) {
                    String body = m.group(1).trim();
                    if (!DocxExpressionEvaluator.isExpression(body)) {
                        imageKeys.add(body);
                    }
                }
                // Mixed content or complex expression → skipped
            }
        });
        log.debug("Extracted {} image keys from DOCX template: {}", imageKeys.size(), imageKeys);
        return imageKeys;
    }

    /**
     * Extract variable references from complex SpEL expressions in image alt texts.
     *
     * <p>Complements {@link #extractImageKeys} by returning variables referenced in
     * complex image expressions — those that {@code extractImageKeys} deliberately skips.
     * These variables are <b>not</b> image parameters; they are regular "Object" type
     * parameters that influence image selection logic.
     *
     * <p><b>Examples:</b></p>
     * <ul>
     *   <li>{@code image:${gender == 'F' ? 'a.png' : 'b.png'}} → {@code ["gender"]}</li>
     *   <li>{@code image:${risk > 80 ? 'warn.png' : 'info.png'}} → {@code ["risk"]}</li>
     *   <li>{@code image:facsimile} → nothing (direct key, handled by extractImageKeys)</li>
     *   <li>{@code image:${photo}} → nothing (simple variable, handled by extractImageKeys)</li>
     * </ul>
     *
     * @param wordPackage the loaded DOCX template
     * @return ordered set of variable names from complex image expressions
     */
    public static Set<String> extractImageExpressionVariables(WordprocessingMLPackage wordPackage) {
        Set<String> variables = new java.util.LinkedHashSet<>();
        scanImageAltTexts(wordPackage, (imageKey) -> {
            if (!imageKey.contains("${")) return;

            // Check if the entire imageKey is a single simple ${variable}
            Matcher fullMatch = SPEL_PATTERN.matcher(imageKey);
            if (fullMatch.matches()) {
                String body = fullMatch.group(1).trim();
                if (!DocxExpressionEvaluator.isExpression(body)) {
                    // Simple variable like ${photo} — handled by extractImageKeys as Image type
                    return;
                }
                // Complex expression like ${gender == 'F' ? ...} — extract variables
                variables.addAll(DocxExpressionEvaluator.extractVariableReferences(body));
            } else {
                // Mixed content like assets/${department}/stamp.png — extract all embedded variables
                Matcher finder = SPEL_PATTERN.matcher(imageKey);
                while (finder.find()) {
                    String body = finder.group(1).trim();
                    variables.addAll(DocxExpressionEvaluator.extractVariableReferences(body));
                }
            }
        });
        if (!variables.isEmpty()) {
            log.debug("Extracted {} expression variables from image alt texts: {}", variables.size(), variables);
        }
        return variables;
    }

    /**
     * Scan all image alt texts matching {@code image:{key}} convention and invoke
     * the consumer for each discovered image key (text after the prefix).
     */
    private static void scanImageAltTexts(WordprocessingMLPackage wordPackage,
                                           java.util.function.Consumer<String> keyConsumer) {
        MainDocumentPart mainPart = wordPackage.getMainDocumentPart();
        collectImageAltTexts(mainPart.getContent(), keyConsumer);

        try {
            if (wordPackage.getDocumentModel() != null) {
                for (var section : wordPackage.getDocumentModel().getSections()) {
                    var hfp = section.getHeaderFooterPolicy();
                    if (hfp == null) continue;
                    collectImageAltTextsFromHeaderFooter(hfp.getDefaultHeader(), keyConsumer);
                    collectImageAltTextsFromHeaderFooter(hfp.getDefaultFooter(), keyConsumer);
                    collectImageAltTextsFromHeaderFooter(hfp.getFirstHeader(), keyConsumer);
                    collectImageAltTextsFromHeaderFooter(hfp.getFirstFooter(), keyConsumer);
                    collectImageAltTextsFromHeaderFooter(hfp.getEvenHeader(), keyConsumer);
                    collectImageAltTextsFromHeaderFooter(hfp.getEvenFooter(), keyConsumer);
                }
            }
        } catch (Exception e) {
            log.debug("Could not scan headers/footers for image alt texts: {}", e.getMessage());
        }
    }

    private static void collectImageAltTexts(List<Object> content,
                                              java.util.function.Consumer<String> keyConsumer) {
        for (Object obj : content) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);
            if (unwrapped instanceof R run) {
                collectImageAltTextsFromRun(run, keyConsumer);
            } else if (unwrapped instanceof ContentAccessor accessor) {
                collectImageAltTexts(accessor.getContent(), keyConsumer);
            }
        }
    }

    private static void collectImageAltTextsFromHeaderFooter(Object headerOrFooterPart,
                                                              java.util.function.Consumer<String> keyConsumer) {
        if (headerOrFooterPart == null) return;
        if (headerOrFooterPart instanceof JaxbXmlPart<?> jaxbPart) {
            Object jaxbContent = jaxbPart.getJaxbElement();
            if (jaxbContent instanceof ContentAccessor accessor) {
                collectImageAltTexts(accessor.getContent(), keyConsumer);
            }
        }
    }

    private static void collectImageAltTextsFromRun(R run,
                                                     java.util.function.Consumer<String> keyConsumer) {
        for (Object obj : run.getContent()) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);
            if (!(unwrapped instanceof Drawing drawing)) continue;

            for (Object anchorOrInline : drawing.getAnchorOrInline()) {
                CTNonVisualDrawingProps docPr = null;

                if (anchorOrInline instanceof Inline inline) {
                    docPr = inline.getDocPr();
                } else if (anchorOrInline instanceof Anchor anchor) {
                    docPr = anchor.getDocPr();
                }

                if (docPr == null) continue;
                String altText = docPr.getDescr();
                if (altText == null || !altText.startsWith(IMAGE_PREFIX)) continue;

                String imageKey = altText.substring(IMAGE_PREFIX.length()).trim();
                if (!imageKey.isEmpty()) {
                    keyConsumer.accept(imageKey);
                }
            }
        }
    }

    // ==================== CONTENT TREE WALKING ====================

    private static int processHeaderFooterImages(Object headerOrFooterPart,
                                                  Map<String, Object> rawContext, File cacheDir,
                                                  Locale locale) {
        if (headerOrFooterPart == null) return 0;

        if (headerOrFooterPart instanceof JaxbXmlPart<?> part) {
            try {
                Object jaxbContent = part.getContents();
                if (jaxbContent instanceof ContentAccessor contentAccessor) {
                    return processImagesInContent(contentAccessor.getContent(), part, rawContext, cacheDir, locale);
                }
            } catch (Exception e) {
                log.debug("Error processing header/footer images: {}", e.getMessage());
            }
        }
        return 0;
    }

    private static int processImagesInContent(List<Object> content, Part owningPart,
                                               Map<String, Object> rawContext, File cacheDir,
                                               Locale locale) {
        int count = 0;
        for (Object obj : content) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);

            if (unwrapped instanceof R run) {
                count += processImagesInRun(run, owningPart, rawContext, cacheDir, locale);
            } else if (unwrapped instanceof ContentAccessor contentAccessor) {
                count += processImagesInContent(contentAccessor.getContent(), owningPart, rawContext, cacheDir, locale);
            }
        }
        return count;
    }

    private static int processImagesInRun(R run, Part owningPart,
                                           Map<String, Object> rawContext, File cacheDir,
                                           Locale locale) {
        int count = 0;
        for (Object obj : run.getContent()) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);
            if (!(unwrapped instanceof Drawing drawing)) continue;

            for (Object anchorOrInline : drawing.getAnchorOrInline()) {
                CTNonVisualDrawingProps docPr = null;
                CTBlip blip = null;

                if (anchorOrInline instanceof Inline inline) {
                    docPr = inline.getDocPr();
                    blip = extractBlip(inline.getGraphic());
                } else if (anchorOrInline instanceof Anchor anchor) {
                    docPr = anchor.getDocPr();
                    blip = extractBlip(anchor.getGraphic());
                }

                if (docPr == null || blip == null) continue;

                String altText = docPr.getDescr();
                if (altText == null || !altText.startsWith(IMAGE_PREFIX)) continue;

                String imageKey = altText.substring(IMAGE_PREFIX.length()).trim();
                if (imageKey.isEmpty()) continue;

                byte[] imageBytes;
                if (imageKey.contains("${")) {
                    String evaluated = evaluateImageExpression(imageKey, rawContext, locale);
                    if (evaluated == null || evaluated.isEmpty() || evaluated.contains("${")) {
                        log.debug("SpEL evaluation produced no usable result for '{}', keeping original", imageKey);
                        continue;
                    }
                    imageBytes = resolveImageValue(evaluated, cacheDir);
                } else {
                    imageBytes = resolveImageBytes(imageKey, rawContext, cacheDir);
                }
                if (imageBytes == null) {
                    log.debug("No image data found for key '{}', keeping original", imageKey);
                    continue;
                }

                if (replaceImageBinary(blip, owningPart, imageBytes)) {
                    log.debug("Replaced image for key '{}' ({} bytes)", imageKey, imageBytes.length);
                    count++;
                }
            }
        }
        return count;
    }

    // ==================== SPEL EXPRESSION SUPPORT ====================

    /**
     * Evaluate SpEL expressions in an image key.
     */
    static String evaluateImageExpression(String imageKey, Map<String, Object> rawContext, Locale locale) {
        Matcher matcher = SPEL_PATTERN.matcher(imageKey);
        StringBuilder result = new StringBuilder();

        while (matcher.find()) {
            String expression = matcher.group(1);
            String evaluated = DocxExpressionEvaluator.evaluate(expression, rawContext, locale);
            matcher.appendReplacement(result, Matcher.quoteReplacement(evaluated));
        }
        matcher.appendTail(result);

        String resolved = result.toString().trim();
        log.debug("Evaluated image expression '{}' → '{}'", imageKey, resolved);
        return resolved.isEmpty() ? null : resolved;
    }

    /**
     * Resolve a direct value string into image bytes (no context lookup).
     */
    static byte[] resolveImageValue(String value, File cacheDir) {
        if (value.startsWith(DATA_URI_PREFIX)) {
            return decodeDataUri(value);
        }

        if (value.startsWith(HTTP_PREFIX) || value.startsWith(HTTPS_PREFIX)) {
            return loadFromUrl(value);
        }

        return loadFromTemplatePath(value, cacheDir);
    }

    // ==================== IMAGE RESOLUTION ====================

    /**
     * Resolve image bytes from context values or template path fallback.
     */
    static byte[] resolveImageBytes(String imageKey, Map<String, Object> rawContext, File cacheDir) {
        // 1. Direct context lookup
        Object value = rawContext.get(imageKey);

        // 2. Nested images map lookup
        if (value == null) {
            Object imagesObj = rawContext.get("images");
            if (imagesObj instanceof Map<?, ?> imagesMap) {
                value = imagesMap.get(imageKey);
            }
        }

        // 3. Resolve from context value
        if (value instanceof String stringValue) {
            if (stringValue.startsWith(DATA_URI_PREFIX)) {
                return decodeDataUri(stringValue);
            }
            if (stringValue.startsWith(HTTP_PREFIX) || stringValue.startsWith(HTTPS_PREFIX)) {
                return loadFromUrl(stringValue);
            }
            return loadFromTemplatePath(stringValue, cacheDir);
        }

        if (value instanceof byte[] bytes) {
            return bytes;
        }

        // 4. Key itself as relative path in template package
        return loadFromTemplatePath(imageKey, cacheDir);
    }

    private static byte[] decodeDataUri(String dataUri) {
        Matcher matcher = DATA_URI_PATTERN.matcher(dataUri);
        if (!matcher.matches()) {
            log.warn("Malformed image data URI (expected data:image/type;base64,payload)");
            return null;
        }

        try {
            String base64Payload = matcher.group(1).replaceAll("\\s+", "");
            return Base64.getDecoder().decode(base64Payload);
        } catch (IllegalArgumentException e) {
            log.warn("Failed to decode Base64 image data: {}", e.getMessage());
            return null;
        }
    }

    private static byte[] loadFromTemplatePath(String relativePath, File cacheDir) {
        if (relativePath == null || relativePath.isBlank()) return null;

        try {
            Path resolved = cacheDir.toPath().resolve(relativePath).normalize();

            if (!isPathSafe(resolved, cacheDir.toPath())) {
                log.warn("Image path '{}' escapes template directory — ignoring (directory traversal blocked)",
                        relativePath);
                return null;
            }

            if (!Files.exists(resolved) || !Files.isRegularFile(resolved)) {
                log.debug("Image file not found in template package: {}", relativePath);
                return null;
            }

            byte[] bytes = Files.readAllBytes(resolved);
            log.debug("Loaded image from template path '{}' ({} bytes)", relativePath, bytes.length);
            return bytes;

        } catch (IOException e) {
            log.warn("Failed to read image file '{}': {}", relativePath, e.getMessage());
            return null;
        }
    }

    // ==================== URL IMAGE FETCHING ====================

    /**
     * Fetch an image from an HTTP/HTTPS URL.
     */
    static byte[] loadFromUrl(String url) {
        try {
            HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create(url))
                    .timeout(URL_REQUEST_TIMEOUT)
                    .GET()
                    .build();

            HttpResponse<byte[]> response = getHttpClient()
                    .send(request, HttpResponse.BodyHandlers.ofByteArray());

            int status = response.statusCode();
            if (status < 200 || status >= 300) {
                log.warn("Image URL '{}' returned HTTP status {}", url, status);
                return null;
            }

            String contentType = response.headers()
                    .firstValue("content-type")
                    .map(ct -> ct.split(";")[0].trim().toLowerCase())
                    .orElse("");

            if (!ALLOWED_CONTENT_TYPES.contains(contentType)) {
                log.warn("Image URL '{}' returned unsupported content type '{}' (allowed: {})",
                        url, contentType, ALLOWED_CONTENT_TYPES);
                return null;
            }

            byte[] body = response.body();
            if (body == null || body.length == 0) {
                log.warn("Image URL '{}' returned empty body", url);
                return null;
            }
            if (body.length > URL_MAX_RESPONSE_BYTES) {
                log.warn("Image URL '{}' response too large ({} bytes, max {})",
                        url, body.length, URL_MAX_RESPONSE_BYTES);
                return null;
            }

            log.debug("Fetched image from URL '{}' ({} bytes, {})", url, body.length, contentType);
            return body;

        } catch (IllegalArgumentException e) {
            log.warn("Invalid image URL '{}': {}", url, e.getMessage());
            return null;
        } catch (IOException e) {
            log.warn("Failed to fetch image from URL '{}': {}", url, e.getMessage());
            return null;
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            log.warn("Image URL fetch interrupted for '{}'", url);
            return null;
        } catch (Exception e) {
            log.warn("Unexpected error fetching image from URL '{}': {}", url, e.getMessage());
            return null;
        }
    }

    private static HttpClient getHttpClient() {
        if (httpClient == null) {
            synchronized (DocxImageReplacer.class) {
                if (httpClient == null) {
                    httpClient = HttpClient.newBuilder()
                            .connectTimeout(URL_CONNECT_TIMEOUT)
                            .followRedirects(HttpClient.Redirect.NORMAL)
                            .build();
                }
            }
        }
        return httpClient;
    }

    /**
     * Replace the shared HTTP client instance. Intended for testing only.
     */
    static void setHttpClient(HttpClient client) {
        synchronized (DocxImageReplacer.class) {
            httpClient = client;
        }
    }

    // ==================== DOCX IMAGE PART MANIPULATION ====================

    private static CTBlip extractBlip(Graphic graphic) {
        if (graphic == null) return null;

        var graphicData = graphic.getGraphicData();
        if (graphicData == null) return null;

        for (Object any : graphicData.getAny()) {
            Object unwrapped = DocxXmlUtils.unwrap(any);
            if (unwrapped instanceof org.docx4j.dml.picture.Pic pic) {
                if (pic.getBlipFill() != null && pic.getBlipFill().getBlip() != null) {
                    return pic.getBlipFill().getBlip();
                }
            }
        }

        return null;
    }

    private static boolean replaceImageBinary(CTBlip blip, Part owningPart, byte[] imageBytes) {
        String rId = blip.getEmbed();
        if (rId == null || rId.isEmpty()) return false;

        try {
            var relPart = owningPart.getRelationshipsPart();
            if (relPart == null) return false;

            Relationship rel = relPart.getRelationshipByID(rId);
            if (rel == null) {
                log.debug("No relationship found for rId '{}'", rId);
                return false;
            }

            Part targetPart = relPart.getPart(rel);
            if (targetPart instanceof BinaryPart binaryPart) {
                binaryPart.setBinaryData(imageBytes);
                return true;
            }

            log.debug("Target part for rId '{}' is not a BinaryPart: {}", rId,
                    targetPart != null ? targetPart.getClass().getSimpleName() : "null");
            return false;

        } catch (Exception e) {
            log.warn("Failed to replace image binary for rId '{}': {}", rId, e.getMessage());
            return false;
        }
    }

    // ==================== SECURITY ====================

    /**
     * Validate that a resolved path stays within the template's base directory.
     */
    static boolean isPathSafe(Path resolved, Path baseDir) {
        try {
            Path normalizedBase = baseDir.toAbsolutePath().normalize();
            Path normalizedResolved = resolved.toAbsolutePath().normalize();
            return normalizedResolved.startsWith(normalizedBase);
        } catch (Exception e) {
            return false;
        }
    }
}
