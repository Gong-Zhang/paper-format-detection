package com.loong.ai.test;

import cn.hutool.core.io.resource.ResourceUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * created by gongzhangchao on 17:16 2024/3/11
 */
public class CheckDocument {

    public static void main(String[] args) {
        String filePath = "classpath:数据科学与大数据技术.docx";
        try (InputStream fis = ResourceUtil.getStream(filePath)) {
            XWPFDocument document = new XWPFDocument(fis);
            //获取文档总页数
            int pageCount = document.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();
            int chineseCount = 0;
            int englishCount = 0;

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                String text = paragraph.getText();
                if (text != null && !text.isEmpty()) {
                    for (char c : text.toCharArray()) {
                        if (isChinese(c)) {
                            chineseCount++;
                        } else if (isEnglish(c)) {
                            englishCount++;
                        }
                    }
                }
            }

            System.out.println("页数：" + pageCount);
            System.out.println("中文字符数：" + chineseCount);
            System.out.println("英文字符数：" + englishCount);
            System.out.println("中文摘要：" + extractAbstract(document));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static boolean isChinese(char c) {
        // 判断字符是否为中文字符（Unicode 编码范围）
        return (c >= 0x4E00 && c <= 0x9FFF);
    }

    public static boolean isEnglish(char c) {
        // 判断字符是否为英文字符（ASCII 编码范围）
        return (c >= 0x0020 && c <= 0x007E);
    }

    /**
     * 获取摘要
     *
     * @param document
     * @return
     */
    public static String extractAbstract(XWPFDocument document) {
        StringBuilder abstractText = new StringBuilder();
        boolean foundAbstract = false;
        // 遍历文档中的每个段落
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            String text = paragraph.getText().trim();
            if (text.startsWith("摘要")) {
                foundAbstract = true;
                //获取摘要的字体和对齐
                List<XWPFRun> runs = paragraph.getRuns();
                // 遍历每个文本运行
                for (XWPFRun run : runs) {
                    // 获取文本运行的字体
                    String fontFamily = run.getFontFamily();
                    // 获取文本运行的字号
                    int fontSize = run.getFontSize();
                    // 输出字体和字号信息
                    System.out.println("字体：" + fontFamily + "，字号：" + fontSize);
                }
            }
            if (foundAbstract && !text.isEmpty()) {
                abstractText.append(text).append("\n");
            }
            if (text.startsWith("ABSTRACT")) {
                // 下一页，退出循环
                break;
            }
        }
        return abstractText.toString().trim();
    }

}
