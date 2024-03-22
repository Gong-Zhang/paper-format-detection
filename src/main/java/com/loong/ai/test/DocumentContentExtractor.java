package com.loong.ai.test;

import cn.hutool.core.io.resource.ResourceUtil;
import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.util.List;

public class DocumentContentExtractor {

    public static void main(String[] args) {
        try {
            // 加载Word文档
            String filePath = "classpath:数据科学与大数据技术.docx";
            WordprocessingMLPackage wordMLPackage = Docx4J.load(ResourceUtil.getStream(filePath));

            // 获取文档的主文档部分
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

            // 获取文档的正文内容
            Body body = mainDocumentPart.getJaxbElement().getBody();

            // 遍历正文内容，提取文本内容
            List<Object> content = body.getContent();
            for (Object obj : content) {
                if (obj instanceof P) {
                    // 如果是段落，获取段落的文本内容
                    P paragraph = (P) obj;

                    String text = paragraph.toString();
                    System.out.println("段落内容：" + text);
                } else if (obj instanceof Tbl) {
                    // 如果是表格，获取表格的文本内容
                    Tbl table = (Tbl) obj;
                    String text = table.toString();
                    System.out.println("表格内容：" + text);
                } else if (obj instanceof JAXBElement) {
                    // 如果是JAXBElement，可能是其他类型的内容，可以根据需要进行处理
                    JAXBElement element = (JAXBElement) obj;
                    String text = element.toString();
                    System.out.println("其他内容：" + text);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
