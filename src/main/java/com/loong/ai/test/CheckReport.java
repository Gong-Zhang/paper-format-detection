package com.loong.ai.test;

import com.aspose.words.*;
import org.apache.commons.lang3.StringUtils;

import java.util.Date;


public class CheckReport {

    public static void main(String[] args) throws Exception {
        //通过Document空的构造函数，创建一个新的空的文档,如果有参数则是根据模板加载文档。
        Document doc = new Document("output/数据科学与大数据技术.docx");

        for (Section section : doc.getSections()) {
            Body body = section.getBody();
            for (Paragraph paragraph : body.getParagraphs()) {

                if (StringUtils.isNotBlank(paragraph.getText())) {
                    System.out.println("段落：" + paragraph.getText());
                    RunCollection runs = paragraph.getRuns();
                    for (int i = 0; i < runs.getCount(); i++) {

                        // 获取字体、字号
                        String text = runs.get(i).getText();
                        Font font = paragraph.getRuns().get(i).getFont();
                        String fontName = font.getName();
                        double fontSize = font.getSize();
                        System.out.println("run的内容：" + text + "\nrun的字体：" + fontName + "\n字号：" + fontSize);

                        // 创建一个新的批注
                        Comment comment = new Comment(doc, "John Doe", "J.D.", new Date());
                        comment.setText("This is a comment.");
                        runs.get(i).getParentNode().insertAfter(comment, runs.get(i));

                    }

                    // 获取段落格式
                    ParagraphFormat paragraphFormat = paragraph.getParagraphFormat();

                    // 获取对齐方式
                    int alignment = paragraphFormat.getAlignment();
                    String alignmentText = "未知";
                    switch (alignment) {
                        case ParagraphAlignment.LEFT:
                            alignmentText = "左对齐";
                            break;
                        case ParagraphAlignment.CENTER:
                            alignmentText = "中间";
                            break;
                        case ParagraphAlignment.RIGHT:
                            alignmentText = "右对齐";
                            break;
                        case ParagraphAlignment.JUSTIFY:
                            alignmentText = "两端对齐";
                            break;
                        case ParagraphAlignment.DISTRIBUTED:
                            alignmentText = "分散对齐";
                            break;
                    }
                    System.out.println("对齐方式: " + alignmentText);
                    // 获取缩进字符
                    double firstLineIndent = paragraphFormat.getFirstLineIndent();
                    System.out.println("首行缩进字符: " + firstLineIndent);
                    // 获取行间距
                    double lineSpacing = paragraphFormat.getLineSpacing();
                    System.out.println("行间距: " + lineSpacing);

                    // 获取大纲级别 0为一级标题，1为二级标题，9为正文
                    int outlineLevel = paragraphFormat.getOutlineLevel();
                    System.out.println("大纲级别: " + outlineLevel);
                }
            }

            // 获取节（Section）中的页眉（HeaderFooterType.HEADER_PRIMARY）、页脚（HeaderFooterType.FOOTER_PRIMARY）内容
            HeaderFooter header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY);
            HeaderFooter footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
            // 输出页眉和页脚内容
            if (header != null) {
                System.out.println("页眉: " + header.toString(SaveFormat.TEXT));
            }
            if (footer != null) {
                System.out.println("页脚: " + footer.toString(SaveFormat.TEXT));
            }
        }
        //doc.save("output/数据科学与大数据技术.docx");
    }


}