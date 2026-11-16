package me.muban.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Text;

import java.util.ArrayList;
import java.util.List;

/**
 * Shared builders for muban-docx tests.
 *
 * <p>Everything here is expressed in plain-text paragraphs, because that is the level at which
 * conditional blocks are defined: one marker per paragraph. Tests that need richer XML (split
 * runs, tables) build it inline.
 */
final class DocxTestSupport {

    private DocxTestSupport() {}

    /**
     * Create a paragraph containing a single run with the given text.
     */
    static P createParagraph(String text) {
        P paragraph = new P();
        R run = new R();
        Text t = new Text();
        t.setValue(text);
        t.setSpace("preserve");
        run.getContent().add(t);
        paragraph.getContent().add(run);
        return paragraph;
    }

    /**
     * Create a simple DOCX document with one paragraph per supplied string.
     */
    static WordprocessingMLPackage createDocx(String... paragraphTexts) throws Exception {
        WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
        MainDocumentPart mainPart = wordPackage.getMainDocumentPart();
        mainPart.getContent().clear();
        for (String text : paragraphTexts) {
            mainPart.getContent().add(createParagraph(text));
        }
        return wordPackage;
    }

    /**
     * Collect the text of every paragraph in a content list, skipping non-paragraph content
     * such as tables.
     */
    static List<String> extractTexts(List<Object> content) {
        List<String> texts = new ArrayList<>();
        for (Object obj : content) {
            Object unwrapped = DocxXmlUtils.unwrap(obj);
            if (unwrapped instanceof P paragraph) {
                texts.add(DocxXmlUtils.extractAllText(List.of(paragraph)));
            }
        }
        return texts;
    }
}
