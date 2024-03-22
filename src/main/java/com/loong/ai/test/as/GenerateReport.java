package com.loong.ai.test.as;

import com.aspose.words.*;
import java.text.SimpleDateFormat;
import java.util.Date;

 
public class GenerateReport {

    public static void main(String[] args) throws Exception {
        //通过Document空的构造函数，创建一个新的空的文档,如果有参数则是根据模板加载文档。
        Document doc = new Document("output/数据科学与大数据技术.docx");

        for (Section section : doc.getSections()) {
            Body body = section.getBody();
            for (Paragraph paragraph : body.getParagraphs()) {
                System.out.println("段落："+paragraph.getText());
                if (paragraph.getText().contains("摘要")) {
                    // 检查摘要段落的格式
                    Font font = paragraph.getRuns().get(0).getFont();
                    String fontName = font.getName();
                    double fontSize = font.getSize();
                    System.out.println("摘要的字体：" + fontName + "\n字号：" + fontSize);
                    // 检查摘要字体、字号等信息
                }
            }
        }

        /*for (Paragraph paragraph : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true)) {
            System.out.println("段落："+paragraph.getText());
            if (paragraph.getText().contains("摘要")) {
                // 检查摘要段落的格式
                Font font = paragraph.getRuns().get(0).getFont();
                String fontName = font.getName();
                double fontSize = font.getSize();
                System.out.println("摘要的字体：" + fontName + "\n字号：" + fontSize);
                // 检查摘要字体、字号等信息
            }
        }*/
    }
 
 
}