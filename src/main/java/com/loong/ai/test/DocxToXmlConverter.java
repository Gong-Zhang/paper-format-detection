package com.loong.ai.test;

import cn.hutool.core.io.resource.ResourceUtil;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import java.util.ArrayList;
import java.util.List;
import javax.xml.bind.JAXBElement;
import org.docx4j.wml.PPr;
import org.docx4j.wml.Jc;


public class DocxToXmlConverter {
    public static void main(String[] args) {
        try {
            // 加载.docx文档
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(ResourceUtil.getStream("classpath:数据科学与大数据技术.docx"));
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

            // 遍历所有的段落
            List<Object> paragraphs = mainDocumentPart.getContent();

            boolean inSummarySection = false;
            boolean isFirstAbstract = true;

            List<P> abstractPList = new ArrayList<>();
            for (Object p : paragraphs) {
                if (p instanceof P) {
                    P paragraph = (P) p;
                    String paragraphText = getParagraphText(paragraph);
                    // 如果处于摘要部分
                    if (isFirstAbstract && paragraphText.contains("摘要")) {
                        inSummarySection = true;
                        isFirstAbstract = false;
                    }
                    if (paragraphText.toUpperCase().contains("RECRUITMENT INFORMATION DATA ANALYSIS AND")) {
                        inSummarySection = false;
                    }
                    if (inSummarySection) {
                        abstractPList.add(paragraph);
                        System.out.println("包含‘摘要’的段落内容：\n" + paragraphText);
                        // 输出段落属性
                        printParagraphProperties(paragraph);

                        // 遍历段落中的所有运行，寻找格式和文本
                        printRunProperties(paragraph);
                    }
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String getParagraphText(P paragraph) {
        StringBuilder sb = new StringBuilder();
        List<Object> contents = paragraph.getContent();
        for (Object content : contents) {
            Text text = findText(content);
            if (text != null) {
                sb.append(text.getValue());
            }
        }
        return sb.toString();
    }

    private static Text findText(Object content) {
        if (content instanceof JAXBElement<?>) {
            JAXBElement<?> element = (JAXBElement<?>) content;
            if (element.getValue() instanceof Text) {
                return (Text) element.getValue();
            }
        } else if (content instanceof R) {
            R run = (R) content;
            for (Object runContent : run.getContent()) {
                if (runContent instanceof JAXBElement<?>) {
                    JAXBElement<?> element = (JAXBElement<?>) runContent;
                    if (element.getValue() instanceof Text) {
                        return (Text) element.getValue();
                    }
                }
            }
        }
        return null; // No text found
    }

    private static void printParagraphProperties(P paragraph) {
        PPr pPr = paragraph.getPPr();
        if (pPr != null) {
            Jc justification = pPr.getJc();
            if (justification != null) {
                System.out.println("段落对齐方式: " + justification.getVal());
            }
            // 打印行间距信息
            PPrBase.Spacing spacing = pPr.getSpacing();
            if (spacing != null) {
                if (spacing.getLine() != null) {
                    System.out.println("行间距: " + spacing.getLine().intValue() / 20.0 + " 磅");
                }
                if (spacing.getBefore() != null) {
                    System.out.println("段前间距: " + spacing.getBefore().intValue() / 20.0 + " 磅");
                }
                if (spacing.getAfter() != null) {
                    System.out.println("段后间距: " + spacing.getAfter().intValue() / 20.0 + " 磅");
                }
            }
        }
    }

    private static void printRunProperties(P paragraph) {
        for (Object r : paragraph.getContent()) {
            if (r instanceof R) {
                R run = (R) r;
                RPr rPr = run.getRPr();
                if (rPr != null) {
                    // 字体大小
                    if (rPr.getSz() != null) {
                        System.out.println("字体大小: " + rPr.getSz().getVal().intValue() / 2 + "pt");
                    }
                    // 字体类型
                    if (rPr.getRFonts() != null) {
                        System.out.println("ASCII（英文）字体: " + rPr.getRFonts().getAscii());
                        System.out.println("East Asia（东亚）字体: " + rPr.getRFonts().getEastAsia());
                        System.out.println("HAnsi（中文）字体: " + rPr.getRFonts().getHAnsi());
                    }
                }
            }
        }
    }
}
