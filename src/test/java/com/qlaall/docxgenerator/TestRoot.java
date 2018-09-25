package com.qlaall.docxgenerator;

import com.qlaall.docxgenerator.action.DocxGenerator;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

/**
 * @author: qlaall
 * @Date:2018/9/25
 * @Time:20:48
 */
public class TestRoot {
    public static void main(String[] args) throws URISyntaxException, IOException {
        URL resource = TestRoot.class.getResource("/templates/template-test.docx");
        XWPFDocument xwpfDocument = new XWPFDocument(Files.newInputStream(Paths.get(resource.toURI())));
        final Map<String, BiConsumer<XWPFDocument, XWPFParagraph>> paragraphDealMap = new HashMap<>();
        final Map<String, Consumer<XWPFTableCell>> tableCellDealMap = new HashMap<>();
        byte[] bytes = DocxGenerator.fillDocx(paragraphDealMap, tableCellDealMap, xwpfDocument);
        URI uri = new URI(TestRoot.class.getResource("/").toURI() + "templates/result.docx");
        File file = new File(uri);
        if (file.exists()) {
            file.delete();
        }
        file.createNewFile();
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        fileOutputStream.write(bytes);
        fileOutputStream.flush();
        fileOutputStream.close();

    }
}
