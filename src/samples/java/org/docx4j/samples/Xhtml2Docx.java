package org.docx4j.samples;

import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.model.structure.PageSizePaper;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.awt.*;
import java.io.File;
import java.math.BigInteger;
import java.net.URL;
import java.util.List;

public class Xhtml2Docx {

    public static void main(String[] args) throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);
        wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().setPgSize(PageSizePaper.A4, false);
        wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().getPgMar().setTop(BigInteger.valueOf(0L));
        wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().getPgMar().setBottom(BigInteger.valueOf(0L));
        wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().getPgMar().setLeft(BigInteger.valueOf(10L));
        wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().getPgMar().setRight(BigInteger.valueOf(10L));
        List<Object> objects = XHTMLImporter.convert(new URL("file:///Users/yibourkelidrissi/Lab/Git/docx4j-ImportXHTML/sample-docs/xhtml/TableBorders-Separate.xhtml"));
        wordMLPackage.getMainDocumentPart().getContent().addAll(objects);
        File outputDocx = new File("/Users/yibourkelidrissi/Desktop/output.docx");
        wordMLPackage.save(outputDocx);
        //Desktop desktop = Desktop.getDesktop();
        //desktop.open(outputDocx);
    }
}
