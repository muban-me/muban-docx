/**
 * <b>muban-docx</b> — a pure-Java DOCX template engine built on docx4j.
 *
 * <p>Provides SpEL-based placeholder substitution, conditional blocks,
 * table row replication, image replacement, and PDF export for Word documents.
 *
 * <h2>Quick start</h2>
 * <pre>{@code
 * String output = MubanDocxEngine.builder()
 *     .template(new File("template.docx"))
 *     .data(Map.of("name", "Jan", "amount", 1500))
 *     .locale(Locale.forLanguageTag("pl-PL"))
 *     .outputFormat("pdf")
 *     .outputDir("/tmp/out/")
 *     .build()
 *     .generate();
 * }</pre>
 *
 * <h2>License</h2>
 * <p>AGPL-3.0-or-later — see {@code LICENSE} in the project root.
 *
 * @see me.muban.docx.MubanDocxEngine
 */
package me.muban.docx;
