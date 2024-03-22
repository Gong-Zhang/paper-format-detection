package com.loong.ai.test;

import cn.hutool.core.io.resource.ResourceUtil;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.util.List;

public class ParagraphFontAndSizeExample {

    public static void main(String[] args) {
        try {
            String filePath = "classpath:数据科学与大数据技术.docx";
            WordprocessingMLPackage wordMLPackage = Docx4J.load(ResourceUtil.getStream(filePath));
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            List<Object> content = mainDocumentPart.getContent();

            for (Object obj : content) {
                if (obj instanceof org.docx4j.wml.P) {
                    org.docx4j.wml.P paragraph = (org.docx4j.wml.P) obj;
                    List<Object> paragraphContent = paragraph.getContent();
                    for (Object pobj : paragraphContent) {
                        if (pobj instanceof JAXBElement) {
                            JAXBElement element = (JAXBElement) pobj;
                            if (element.getDeclaredType().equals(org.docx4j.wml.R.class)) {
                                org.docx4j.wml.R run = (org.docx4j.wml.R) element.getValue();
                                RPr runProperties = run.getRPr();
                                if (runProperties != null) {
                                    String fontFamily = runProperties.getRFonts().getAscii();
                                    System.out.println("字体：" + fontFamily);
                                    if (runProperties.getSz() != null) {
                                        // 字号以半磅为单位，需要除以2获取磅数
                                        int fontSize = runProperties.getSz().getVal().intValue() / 2;
                                        System.out.println("字号：" + fontSize + "磅");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
