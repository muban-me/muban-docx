package me.muban.docx;

import org.apache.fop.apps.FOUserAgent;
import org.apache.fop.apps.FopFactory;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.convert.out.fo.PlaceholderReplacementHandler;
import org.docx4j.convert.out.fo.renderers.FORendererApacheFOP;
import org.docx4j.openpackaging.exceptions.Docx4JException;

import javax.xml.transform.Source;
import java.io.File;
import java.io.OutputStream;

/**
 * Apache FOP renderer that writes a DOCX-to-FO conversion as PNG page images.
 *
 * <p>FOP writes the first page to the supplied output stream and subsequent pages
 * to files derived from {@link FOUserAgent#setOutputFile(File)}. Setting the output
 * file to {@code page_.png} produces {@code page_2.png}, {@code page_3.png}, etc.,
 * while the caller writes the first page to {@code page_1.png}.</p>
 */
final class PngFoRenderer extends FORendererApacheFOP {

    private final File outputFile;
    private final float targetResolution;

    PngFoRenderer(File outputFile, float zoomRatio) {
        this.outputFile = outputFile;
        this.targetResolution = 72f * zoomRatio;
    }

    @Override
    protected void render(FOSettings settings, FopFactory fopFactory, FOUserAgent foUserAgent,
                          String mime, Source source,
                          PlaceholderReplacementHandler.PlaceholderLookup placeholderLookup,
                          OutputStream outputStream) throws Docx4JException {
        FOUserAgent agent = foUserAgent != null ? foUserAgent : fopFactory.newFOUserAgent();
        agent.setOutputFile(outputFile);
        agent.setTargetResolution(targetResolution);
        super.render(settings, fopFactory, agent, mime, source, placeholderLookup, outputStream);
    }
}
