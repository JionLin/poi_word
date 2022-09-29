package com.jiaolin;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author johnny
 * @Classname POIUtil
 * @Description
 * @Date 2022/9/30 05:04
 */
public class POIUtil {

    public static final String removeFile = "";
    public static final String outputFIle = "";

    public static void addFooter() throws Exception {
        XWPFDocument document = new XWPFDocument(new FileInputStream(removeFile));

        XWPFHeaderFooterPolicy footerPolicy = new XWPFHeaderFooterPolicy(document);

        XWPFFooter footer = footerPolicy.createFooter(STHdrFtr.DEFAULT);

        XWPFParagraph paragraph = footer.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.BOTH);

        XWPFRun run = paragraph.createRun();

        // 设置样式
        setXWPFRunStyle(run, "宋体", 9);
        String info = "合同编号: Bojj100110031004";
        run.setText(info);
        FileOutputStream outputStream = new FileOutputStream(outputFIle);
        document.write(outputStream);
        outputStream.close();
        document.close();


    }

    // 新增页脚
    private static void createFooter(XWPFDocument document, HttpServletResponse response, String textContent) {
        try {
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();

            XWPFHeaderFooterPolicy footerPolicy = new XWPFHeaderFooterPolicy(document);

            XWPFFooter footer = footerPolicy.createFooter(STHdrFtr.DEFAULT);

            XWPFParagraph paragraph = footer.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.BOTH);

            XWPFRun run = paragraph.createRun();

            // 设置样式
            setXWPFRunStyle(run, "宋体", 9);
            run.setText(textContent);


            ServletOutputStream outputStream = response.getOutputStream();
            document.write(byteArrayOutputStream);
            outputStream.write(byteArrayOutputStream.toByteArray());

            outputStream.flush();
            outputStream.close();
            byteArrayOutputStream.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * 设置页脚的字体样式
     *
     * @param r1 段落元素
     */
    private static void setXWPFRunStyle(XWPFRun r1, String font, int fontSize) {
        r1.setFontSize(fontSize);
        CTRPr rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR().addNewRPr();
        CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
        fonts.setAscii(font);
        fonts.setEastAsia(font);
        fonts.setHAnsi(font);
    }

}
