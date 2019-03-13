package com.qlaall.docxgenerator;

import com.qlaall.docxgenerator.action.DocxGenerator;
import com.qlaall.docxgenerator.styles.Style;
import com.qlaall.docxgenerator.util.WordUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
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
public class DocxTest {
    public static void main(String[] args) throws URISyntaxException, IOException {
        URL resource = DocxTest.class.getResource("/templates/template-test.docx");
        XWPFDocument xwpfDocument = new XWPFDocument(Files.newInputStream(Paths.get(resource.toURI())));
        final Map<String, BiConsumer<XWPFDocument, XWPFParagraph>> paragraphDealMap =test();
        final Map<String, Consumer<XWPFTableCell>> tableCellDealMap = new HashMap<>();
        byte[] bytes = DocxGenerator.fillDocx(paragraphDealMap, tableCellDealMap, xwpfDocument);
        URI uri = new URI(DocxTest.class.getResource("/").toURI() + "templates/result.docx");
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
    private static Map<String, BiConsumer<XWPFDocument, XWPFParagraph>> test(){
        Map<String, BiConsumer<XWPFDocument, XWPFParagraph>> m=new HashMap<>();
        m.put("${title}",(docx,paragraph)->{
            WordUtil.clearMe(paragraph);
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.append("，另外各测风塔与长期参考数据同期的风速为");
            XWPFRun run = paragraph.createRun();
            run.setText(stringBuilder.toString());
            Style.DEFAULT_PARA_STYLE.proccess(run);
        });
//        ${title}
//        ${test-table1}
//        ${test-table1-describe}
//        ${test-pic1}
//        ${test-pic1-describe}
//        DocxGenerator simplifies the functionality of poi, making it easier to use.
//        ${test-table2}
//        ${test-table2-describe}
//        ${test-pic2}
//        ${test-pic2-describe}
//        DocxGenerator简化了poi的功能，让它更容易使用。
//        ${test-paragraph}
        return m;
    }
}
