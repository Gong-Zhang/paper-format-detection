//package com.loong.ai.test;
//
//import cn.hutool.core.io.resource.ResourceUtil;
//import org.docx4j.Docx4J;
//import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
//import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
//import org.docx4j.wml.P;
//import org.docx4j.wml.Text;
//
//import javax.xml.bind.JAXBElement;
//import java.io.File;
//import java.io.InputStream;
//import java.util.List;
//
//public class AbstractExtractor {
//
//    public static void main(String[] args) {
//        try {
//            String filePath = "classpath:数据科学与大数据技术.docx";
//            WordprocessingMLPackage wordMLPackage = Docx4J.load(ResourceUtil.getStream(filePath));
//            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
//            List<Object> content = mainDocumentPart.getContent();
//
//            boolean isAbstract = false;
//            StringBuilder abstractContent = new StringBuilder();
//
//            for (Object obj : content) {
//                if (obj instanceof P) {
//                    P paragraph = (P) obj;
//                    List<Object> paragraphContent = paragraph.getContent();
//                    for (Object pobj : paragraphContent) {
//                        if (pobj instanceof JAXBElement) {
//                            JAXBElement element = (JAXBElement) pobj;
//                            if (element.getName().getLocalPart().equals("r")) {
//                                List<Object> runContent = element.getContent();
//                                for (Object robj : runContent) {
//                                    if (robj instanceof Text) {
//                                        Text text = (Text) robj;
//                                        String value = text.getValue();
//                                        if (value.toLowerCase().contains("abstract")) {
//                                            isAbstract = true;
//                                            continue;
//                                        }
//                                    }
//                                }
//                            }
//                        }
//                    }
//                }
//
//                if (isAbstract) {
//                    if (obj instanceof P) {
//                        P paragraph = (P) obj;
//                        List<Object> paragraphContent = paragraph.getContent();
//                        for (Object pobj : paragraphContent) {
//                            if (pobj instanceof JAXBElement) {
//                                JAXBElement element = (JAXBElement) pobj;
//                                if (element.getName().getLocalPart().equals("r")) {
//                                    List<Object> runContent = element.getContent();
//                                    for (Object robj : runContent) {
//                                        if (robj instanceof Text) {
//                                            Text text = (Text) robj;
//                                            abstractContent.append(text.getValue());
//                                        }
//                                    }
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//
//            System.out.println("摘要内容：" + abstractContent.toString());
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }
//}
