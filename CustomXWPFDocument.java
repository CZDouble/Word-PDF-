package com.longshine.superapp.nucleicresult.util;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.io.InputStream;

/**
 * @Author Chen ZhiZhe
 * @Create 2021.08.13 15:18
 * @Description
 */
public class CustomXWPFDocument extends XWPFDocument {


    public CustomXWPFDocument(InputStream in) throws IOException {

        super(in);
    }

    public CustomXWPFDocument() {

        super();
    }

    public CustomXWPFDocument(OPCPackage pkg) throws IOException {

        super(pkg);
    }
}