/*
 * Copyright 2015 vamsi.mohan.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package mnm.buildmyreport;

import com.mongodb.MongoClient;
import com.mongodb.client.MongoCollection;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import org.bson.Document;
import org.docx4j.TraversalUtil;
import org.docx4j.TraversalUtil.CallbackImpl;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.JAXBAssociation;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.samples.AbstractSample;
import org.docx4j.wml.Body;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.w3c.dom.Node;
import static java.util.Arrays.asList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;
import java.util.Properties;
import java.util.ResourceBundle;
import static mnm.buildmyreport.DocxToXcl.sequenceExportList;
import org.docx4j.Docx4jProperties;
import org.docx4j.utils.ResourceUtils;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.Color;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;

/**
 *
 * @author vamsi.mohan
 */
public class DocxToDocx extends AbstractSample {

    //public static JAXBContext context = org.docx4j.jaxb.Context.jc;
    static int tblcount = 0;
    static int pageCount = 0;

    static MainDocumentPart mdpOut;
    static org.docx4j.wml.ObjectFactory wmlFactory;
    static List<JAXBAssociation> allJAXBTblsInDocx;
    static HashMap<Element, Integer> sequenceList = new HashMap<>();
    static HashMap<Integer, Object> sequenceExportList = new HashMap<>();
    //create two lists, one for tables and one for figures so that we have only one parse for each type
    static HashMap<String, Element> tableElements = new HashMap<>();
    static HashMap<String, Element> figureElements = new HashMap<>();
    static char type;

    /**
     * @param args the command line arguments
     * @throws java.lang.Exception
     */
    public static void main(String[] args) throws Exception {

        try {
            getInputFilePath(args);
        } catch (IllegalArgumentException e) {
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/ANTIFOAMING AGENT MARKET â€“ GLOBAL FORECAST TO 2021.doc";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/report_1467732928.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Masked - Cardiovascular Information System Market – Forecasts to 2020.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Organic Electronics Market - Global Analysis and Forecast 2020.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Data Center Networking.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Mobile 3D Market.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Casino Management Systems (CMS) Market.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/cardiovascular.docx";

            inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Air and Missile.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/report_1472123853.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Fire Resistant Glass.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Immunotherapy Drugs Market - Copy.docx";

        }
        try {
            getOutputFilePath(args);
        } catch (IllegalArgumentException e) {

            //outputfilepath = System.getProperty("user.dir") + "/output/OUT_MnM.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/test report.docx";
        }
        try {
            getElements(args);
        } catch (IllegalArgumentException e) {

            System.out.println(e.getMessage());
            /*export mode with some test data starts*/
            //mode = "export";
            //reportId = "test";
//
            /*export mode with some test data ends*/

            /*export mode with some test data starts*/
            mode = "parse";
            //type = 'F';
            //Make sure all the elements are there in test doc, as it has to be a sequence map whose key is the index/order as below.

//            tocElements.add(new Element("T", "_Toc445308482", "47"));//cardiovascular Table 1
//            tocElements.add(new Element("T", "_Toc445308483", "48"));//cardiovascular
//            
//            tocElements.add(new Element("T", "_Toc445308484", "47"));//cardiovascular
//            tocElements.add(new Element("T", "_Toc445308485", "48"));//cardiovascular
//            
//            
            tocElements.add(new Element("F", "_Toc445308580", "47"));//cardiovascular
            tocElements.add(new Element("F", "_Toc445308581", "48"));//cardiovascular

//            tocElements.add(new Element("T", "_Toc368936324", "48"));//casino management
//            tocElements.add(new Element("F", "_Toc401326474", "48"));//Organic Electronics Market
//            tocElements.add(new Element("T", "_Toc369865422", "48"));//Data Center Networking
//            tocElements.add(new Element("T", "_Toc369865423", "48"));//Data Center Networking
//            tocElements.add(new Element("T", "_Toc369865424", "48"));//Data Center Networking
//            tocElements.add(new Element("F", "_Toc348352919", "48"));//Mobile 3d market
//            tocElements.add(new Element("F", "_Toc348352920", "48"));//Mobile 3d market
//            tocElements.add(new Element("F", "_Toc348352921", "48"));//Mobile 3d market
//            tocElements.add(new Element("F", "_Toc348352922", "48"));//Mobile 3d market
//            tocElements.add(new Element("F", "_Toc348352923", "48"));//Mobile 3d market
//            tocElements.add(new Element("F", "_Toc348352924", "48"));//Mobile 3d market
            //reflinkId _Toc348352919 //reflinkId	_Toc368936325
            /*export mode with some test data ends*/
        }

        System.out.println("inputfilepath " + inputfilepath);
        outputfilepath = System.getProperty("user.dir") + "/output/OUT_MnM.docx";
        System.out.println("outputfilepath " + outputfilepath);

        System.out.println("user.dir is " + System.getProperty("user.dir"));

        if (mode.equalsIgnoreCase("parse")) {
            parse(inputfilepath, reportId);
        }
        if (mode.equalsIgnoreCase("export")) {
//            System.out.println("Type of content to export " + (elementType.equalsIgnoreCase("T") ? "Table" : (elementType.equalsIgnoreCase("F") ? "Figure" : "Undefined")));
            export(inputfilepath, tocElements);
        }
        if (mode.equalsIgnoreCase("all")) {
//            System.out.println("Type of content to export " + (elementType.equalsIgnoreCase("T") ? "Table" : (elementType.equalsIgnoreCase("F") ? "Figure" : "Undefined")));
            exportAll(inputfilepath, type);
        }
    }

    /**
     *
     * @param path
     * @param list
     * @param elementType
     * @return
     * @throws Docx4JException
     * @throws JAXBException
     */
    @SuppressWarnings("ResultOfObjectAllocationIgnored")
    public static String exportAll(String path, char type) throws Docx4JException, JAXBException {
        /*
         * You can invoke this from an OS command line with something like:
         * 
         * java -cp dist/docx4j.jar:dist/log4j-1.2.15.jar
         * org.docx4j.samples.OpenMainDocumentAndTraverse inputdocx
         * 
         * Note the minimal set of supporting jars.
         * 
         * If there are any images in the document, you will also need:
         * 
         * dist/xmlgraphics-commons-1.4.jar:dist/commons-logging-1.1.1.jar
         */

        int seq = 0;//assign sequences as they are recieved from cmd line, sequence is implicit based on the order of elements in cmd line

        System.out.println("export mode!!!");

        WordprocessingMLPackage wordMLPackageIn = WordprocessingMLPackage
                .load(new java.io.File(inputfilepath));
        final MainDocumentPart documentPart = wordMLPackageIn.getMainDocumentPart();

        System.out.println("Too Good so far!!!!!!!!!");

        WordprocessingMLPackage wordMLPackageOut = (WordprocessingMLPackage) wordMLPackageIn.clone();

        wmlFactory = new org.docx4j.wml.ObjectFactory();
        org.docx4j.wml.Document documentOut = wmlFactory.createDocument();
        final org.docx4j.wml.Body bodyOut = wmlFactory.createBody();
        documentOut.setBody(bodyOut);

        //final MainDocumentPart mdpOut = wordMLPackageOut.getMainDocumentPart();
        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart
                .getJaxbElement();
        Body body = wmlDocumentEl.getBody();
        bodyOut.setSectPr(body.getSectPr());

        if (type == 'T') {
            TableExporter tableExporter = new TableExporter();
            new TraversalUtil(body, tableExporter);
            //filterTables(tableExporter.getPTablePairs());
            for (PTablePair ptp : tableExporter.getPTablePairs()) {
                //PTablePair ptp = (PTablePair) sequenceExportList.get(i);

                bodyOut.getContent().add(prependIndex(ptp.title, "Table " + ptp.getIndex() + " "));
                bodyOut.getContent().add(ptp.tbl);
                bodyOut.getContent().add(ptp.footer);
            }

            System.out.println("Time to export to docx, we have parsed a total of " + tableExporter.getPTablePairs().size() + " tables");
        }
        if (type == 'F') {
            FigureExporterNew figureExporter = new FigureExporterNew();
            new TraversalUtil(body, figureExporter);
            //filterFigures(figureExporter.getPFigurePairs());

            for (PFigurePair pfp : figureExporter.getPFigurePairs()) {
                bodyOut.getContent().add(prependIndex(pfp.title, "Figure " + pfp.getIndex() + " "));
                bodyOut.getContent().add(pfp.figure);
                if (pfp.footer != null) {
                    for (int j = 0; j < pfp.footer.size(); j++) {
                        bodyOut.getContent().add(pfp.footer.get(j));
                    }
                }
            }

            System.out.println("Time to export to docx, we have parsed a total of " + figureExporter.getPFigurePairs().size() + " figures");
        }
        String openXML = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:pPr>"
                + "<w:pStyle w:val=\"NewFootnote\"/>"
                + "</w:pPr>"
                + "</w:p>";
        P footNoteP = (P) XmlUtils.unmarshalString(openXML);

        bodyOut.getContent().add(footNoteP);

        //System.out.println("Time to export to docx, we have parsed a total of" + sequenceExportList.size() + " tables and(or) figures");
        documentOut.setBody(bodyOut);

        mdpOut = new MainDocumentPart();
        mdpOut.setContents(documentOut);

        //wordMLPackageOut.setsetPartName(mdpOut);
        wordMLPackageOut.getMainDocumentPart().getContent().clear();
        wordMLPackageOut.getMainDocumentPart().getContent().add(bodyOut);

        outputfilepath = org.apache.commons.io.FilenameUtils.getFullPath(outputfilepath) + org.apache.commons.io.FilenameUtils.getName(outputfilepath);
        wordMLPackageOut.save(new java.io.File(outputfilepath));

        return outputfilepath;
    }

    /**
     *
     * @param path
     * @param list
     * @param elementType
     * @return
     * @throws Docx4JException
     * @throws JAXBException
     */
    @SuppressWarnings("ResultOfObjectAllocationIgnored")
    public static String export(String path, List<Element> list) throws Docx4JException, JAXBException {
        /*
         * You can invoke this from an OS command line with something like:
         * 
         * java -cp dist/docx4j.jar:dist/log4j-1.2.15.jar
         * org.docx4j.samples.OpenMainDocumentAndTraverse inputdocx
         * 
         * Note the minimal set of supporting jars.
         * 
         * If there are any images in the document, you will also need:
         * 
         * dist/xmlgraphics-commons-1.4.jar:dist/commons-logging-1.1.1.jar
         */

        int seq = 0;//assign sequences as they are recieved from cmd line, sequence is implicit based on the order of elements in cmd line
        for (Element currelement : list) {
            if (currelement.getType().equals("T")) {
                tableElements.put(currelement.getId(), currelement);
            }
            if (currelement.getType().equals("F")) {
                figureElements.put(currelement.getId(), currelement);
            }
            //create a Map with index/sequence as key and element as value. this is needed to preserve order of selection
            sequenceList.put(currelement, seq++);
        }

        System.out.println("export mode!!!");

        WordprocessingMLPackage wordMLPackageIn = WordprocessingMLPackage
                .load(new java.io.File(inputfilepath));
        final MainDocumentPart documentPart = wordMLPackageIn.getMainDocumentPart();

        System.out.println("Too Good so far!!!!!!!!!");

        WordprocessingMLPackage wordMLPackageOut = (WordprocessingMLPackage) wordMLPackageIn.clone();

        wmlFactory = new org.docx4j.wml.ObjectFactory();
        org.docx4j.wml.Document documentOut = wmlFactory.createDocument();
        final org.docx4j.wml.Body bodyOut = wmlFactory.createBody();
        documentOut.setBody(bodyOut);

        //final MainDocumentPart mdpOut = wordMLPackageOut.getMainDocumentPart();
        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart
                .getJaxbElement();
        Body body = wmlDocumentEl.getBody();
        bodyOut.setSectPr(body.getSectPr());

        if (!tableElements.isEmpty()) {
            TableExporter tableExporter = new TableExporter();
            new TraversalUtil(body, tableExporter);
            filterTables(tableExporter.getPTablePairs());

        }
        if (!figureElements.isEmpty()) {
            FigureExporter figureExporter = new FigureExporter();
            new TraversalUtil(body, figureExporter);
            filterFigures(figureExporter.getPFigurePairs());

        }
        String openXML = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:pPr>"
                + "<w:pStyle w:val=\"NewFootnote\"/>"
                + "</w:pPr>"
                + "</w:p>";
        P footNoteP = (P) XmlUtils.unmarshalString(openXML);

        if (!sequenceExportList.isEmpty()) {
            for (int i = 0; i < sequenceExportList.size(); i++) {

                if (sequenceExportList.containsKey(i)) {
                    if (sequenceExportList.get(i).getClass().getName().equals("mnm.buildmyreport.PTablePair")) {
                        PTablePair ptp = (PTablePair) sequenceExportList.get(i);

                        bodyOut.getContent().add(prependIndex(ptp.title, "Table " + ptp.getIndex() + " "));
                        bodyOut.getContent().add(ptp.tbl);
                        bodyOut.getContent().add(ptp.footer);
                    }
                    if (sequenceExportList.get(i).getClass().getName().equals("mnm.buildmyreport.PFigurePair")) {
                        PFigurePair pfp = (PFigurePair) sequenceExportList.get(i);

                        bodyOut.getContent().add(prependIndex(pfp.title, "Figure " + pfp.getIndex() + " "));
                        bodyOut.getContent().add(pfp.figure);
                        bodyOut.getContent().add(pfp.footer);
                    }
                    bodyOut.getContent().add(footNoteP);
                }
            }
        }
        System.out.println("Time to export to docx, we have parsed a total of" + sequenceExportList.size() + " tables and(or) figures");
        documentOut.setBody(bodyOut);

        mdpOut = new MainDocumentPart();
        mdpOut.setContents(documentOut);

        //wordMLPackageOut.setsetPartName(mdpOut);
        wordMLPackageOut.getMainDocumentPart().getContent().clear();
        wordMLPackageOut.getMainDocumentPart().getContent().add(bodyOut);

        outputfilepath = org.apache.commons.io.FilenameUtils.getFullPath(outputfilepath) + org.apache.commons.io.FilenameUtils.getName(outputfilepath);
        wordMLPackageOut.save(new java.io.File(outputfilepath));

        return outputfilepath;
    }

    private static void filterTables(List<PTablePair> pTablePairs) {
        for (PTablePair ptp : pTablePairs) {
            for (int j = 0; j < ptp.title.getContent().size(); j++) {
                if (ptp.title.getContent().get(j) instanceof javax.xml.bind.JAXBElement) {
                    JAXBElement jaxb = (JAXBElement) ptp.title.getContent().get(j);

                    if (jaxb.getDeclaredType().getName().equals("org.docx4j.wml.CTBookmark")) {

                        String tableId = ((org.docx4j.wml.CTBookmark) (jaxb.getValue())).getName();
                        //System.out.println("tableId for table no."+j+1+" is: "+tableId);
                        if (tableElements.containsKey(tableId)) {

                            //System.out.println("
                            ptp.setIndex(((Element) tableElements.get(tableId)).getindex());
                            sequenceExportList.put(sequenceList.get((Element) tableElements.get(tableId)), ptp);
                        }
                    }
                }
            }
        }
    }

    private static void filterFigures(List<PFigurePair> pFigurePairs) {
        System.out.println("We have parsed " + pFigurePairs.size() + " figures ");
        for (PFigurePair pfp : pFigurePairs) {
            System.out.println("we are at figure : " + pfp.title + " whose ctblist size is" + pfp.ctblist.size());
            for (CTBookmark ctb : pfp.ctblist) {
                System.out.println("bookmark for figure is: " + ctb.getName());
                if (figureElements.containsKey(ctb.getName())) {
                    pfp.setIndex(((Element) figureElements.get(ctb.getName())).getindex());
                    sequenceExportList.put(sequenceList.get((Element) figureElements.get(ctb.getName())), pfp);
                }
            }
        }
    }

    private static P prependIndex(P p, String index) {

        //P p = ptp.title;
        boolean prepended = false;
        org.docx4j.wml.Text text = null;
        //String tabletitletext = p.toString();
        for (int l = 0; l < p.getContent().size(); l++) {
            //org.docx4j.wml.P.Hyperlink hl = null;
            if (prepended) {
                break;
            }
            if ((!prepended) && p.getContent().get(l) instanceof org.docx4j.wml.R) {
                org.docx4j.wml.R r = (org.docx4j.wml.R) p.getContent().get(l);
                for (int k = 0; k < r.getContent().size(); k++) {
                    if (r.getContent().get(k) instanceof javax.xml.bind.JAXBElement) {
                        javax.xml.bind.JAXBElement jaxb = (javax.xml.bind.JAXBElement) r.getContent().get(k);
                        if (jaxb.getDeclaredType().getName().equals("org.docx4j.wml.Text")) {
                            text = (org.docx4j.wml.Text) (jaxb).getValue();
                            //System.out.println(text.getValue());
                            //if (text.getValue().startsWith("FIGURE") || text.getValue().startsWith("TABLE")) {
                            //index = text.getValue();
                            //tabletitletext = tabletitletext + text.getValue().toUpperCase();
                            text.setValue(index + text.getValue().toUpperCase());
                            prepended = true;
                            //break;
                        }
                        //if(prepended) break;
                    }
                }
            }

        }

        p.getPPr().setPStyle(null);
        ParaRPr rpr = wmlFactory.createParaRPr();
        //createRPr();
        // Create object for caps
        BooleanDefaultTrue booleandefaulttrue = wmlFactory.createBooleanDefaultTrue();
        rpr.setCaps(booleandefaulttrue);
        // Create object for rFonts
        RFonts rfonts = wmlFactory.createRFonts();
        rpr.setRFonts(rfonts);
        rfonts.setAscii("Franklin Gothic Medium Cond");
        rfonts.setHAnsi("Franklin Gothic Medium Cond");
        // Create object for sz
        HpsMeasure hpsmeasure = wmlFactory.createHpsMeasure();
        rpr.setSz(hpsmeasure);
        hpsmeasure.setVal(BigInteger.valueOf(24));
        // Create object for lang
        CTLanguage language = wmlFactory.createCTLanguage();
        rpr.setLang(language);
        language.setEastAsia("en-IN");
        // Create object for color
        Color color = wmlFactory.createColor();
        rpr.setColor(color);
        color.setVal("0D5775");
        p.getPPr().setRPr(rpr);
        //getPStyle().setRPr(rpr);
        //return rpr;
        return p;
    }
    //System.out.println(XmlUtils.marshaltoString(currP));

    protected static class TOCElementFinder extends CallbackImpl {

        String rListOfTablesDetector;
        private P p;
        boolean foundListOfTables = false;
        String indent = "";
        String TOCelementname, index;
        CharSequence cs = "List of TableS";
        List<TOCElement> tocElementList = new ArrayList<>();
        int pagecount;

        @Override

        public List<Object> apply(Object o) {

            //System.out.println("walking through elements");
            if (o instanceof P) {
                if (XmlUtils.marshaltoString(o).contains("w:type=\"page\"/")) {
                    pagecount++;
                }
                p = (P) o;

                //System.out.println(wmlText.getValue());
                //System.out.println(XmlUtils.marshaltoString(currP));
                if (!foundListOfTables) {

                    for (int l = 0; l < p.getContent().size(); l++) {
                        if (p.getContent().get(l) instanceof org.docx4j.wml.R) {
                            R row = ((org.docx4j.wml.R) (p.getContent().get(l)));
                            //System.out.println(XmlUtils.marshaltoString(row));

                            if (XmlUtils.marshaltoString(row).contains(cs)) {
                                System.out.println("foundtables!!!");
                                foundListOfTables = true;
                            }

                        }

                    }
                } else {

                    for (int l = 0; l < p.getContent().size(); l++) {
                        org.docx4j.wml.P.Hyperlink hl = null;
                        if (p.getContent().get(l) instanceof javax.xml.bind.JAXBElement) {
                            javax.xml.bind.JAXBElement jaxb = (javax.xml.bind.JAXBElement) (p.getContent().get(l));
                            if (jaxb.getDeclaredType().getName().equals("org.docx4j.wml.P$Hyperlink")) {
                                hl = (org.docx4j.wml.P.Hyperlink) jaxb.getValue();
                                String hyperlinkAnchor = hl.getAnchor();

                                org.docx4j.wml.Text text;
                                TOCelementname = "";
                                index = "";
                                for (int k = 0; k < hl.getContent().size(); k++) {

                                    if (hl.getContent().get(k) instanceof org.docx4j.wml.R) {
                                        org.docx4j.wml.R r = (org.docx4j.wml.R) hl.getContent().get(k);

                                        for (int m = 0; m < r.getContent().size(); m++) {
                                            if (r.getContent().get(m) instanceof javax.xml.bind.JAXBElement) {
                                                javax.xml.bind.JAXBElement jaxbr = (javax.xml.bind.JAXBElement) (r.getContent().get(m));
                                                if (jaxbr.getDeclaredType().getName().equals("org.docx4j.wml.Text")) {
                                                    text = (org.docx4j.wml.Text) jaxbr.getValue();
                                                    //System.out.println(text.getValue());
                                                    if (text.getValue().startsWith("FIGURE") || text.getValue().startsWith("TABLE")) {
                                                        index = text.getValue();
                                                    } else {
                                                        TOCelementname = TOCelementname.concat(text.getValue());
                                                        
                                                    }
                                                }
                                            }
                                        }

                                    }

                                }
                                if (hyperlinkAnchor != null) {
                                    String heading = TOCelementname.split("PAGEREF")[0].trim();
                                    System.out.println(heading + " is parsed");
                                    try {
                                        String indexno = index.split(" ")[1];
                                        TOCElement tocElement = new TOCElement(heading, index.startsWith("FIGURE") ? 'F' : 'T', hl.getAnchor(), Integer.parseInt(indexno));
                                        tocElementList.add(tocElement);
                                    } catch (Exception e) {
                                        //is not a table or figure so ignore it
                                    }
                                    //String heading = temp.split(temp)[0]

                                }
                                //add tocelement,type as value and its href id as key
                                //System.out.println(TOCelementname.trim());
                            }

                            //System.out.println(XmlUtils.marshaltoString(currP));
                        }

                    }

                }
            }
            //System.out.println(TOCelementname);
            return null;
        }

        @Override
        public boolean shouldTraverse(Object o
        ) {
            return true;
        }

        // Depth first
        @Override
        public void walkJAXBElements(Object parent
        ) {

            indent += "    ";

            List children = getChildren(parent);
            if (children != null) {

                children.stream().map((o) -> XmlUtils.unwrap(o)).map((o) -> {
                    this.apply(o);
                    return o;
                }).filter((o) -> (this.shouldTraverse(o))).forEach((o) -> {
                    walkJAXBElements(o);
                });
            }

            indent = indent.substring(0, indent.length() - 4);
            //System.out.println(indent);
        }

        @Override
        public List<Object> getChildren(Object o
        ) {
            return TraversalUtil.getChildrenImpl(o);
        }

        public String getTOCElements() {
            return this.TOCelementname;
        }

        public List<TOCElement> getTOCElementList() {
            return this.tocElementList;
        }
    }

    private static void parse(String inputfilepath, String reportId) throws Docx4JException, IOException {
        System.out.println("parse mode!!!");

        WordprocessingMLPackage wordMLPackageIn = WordprocessingMLPackage
                .load(new java.io.File(inputfilepath));
        final MainDocumentPart documentPart = wordMLPackageIn.getMainDocumentPart();

        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart
                .getJaxbElement();
        Body body = wmlDocumentEl.getBody();

        TOCElementFinder tocFinder = new TOCElementFinder();
        new TraversalUtil(body, tocFinder);
        System.out.println("total toc elements found " + tocFinder.tocElementList.size());
        for (TOCElement toce : tocFinder.tocElementList) {
            if (toce.type=='T') {
                System.out.println(toce.elementName + "::" + toce.type + "::" + toce.refLinkId + "::" + toce.index);
            }
        }
        //System.out.println("total pages found " + tocFinder.pagecount);

        System.out.println("--------------------------------------------");
        System.out.println("First pass done , now for second pass!!");
        System.out.println("--------------------------------------------");

        BMRUtility BMRUtilityExporter = new BMRUtility();
        BMRUtilityExporter.setWordMLPkg(wordMLPackageIn);
        new TraversalUtil(body, BMRUtilityExporter);

        Map<String, String> tableHeadings = BMRUtilityExporter.getListOfTables();

        tableHeadings.keySet();
        for (int i = 0; i < tocFinder.getTOCElementList().size(); i++) {
            TOCElement tocElement = tocFinder.getTOCElementList().get(i);
            String tableHeaders = tableHeadings.get(tocElement.elementName);

            if (tableHeaders != null) {
                String[] headers = tableHeaders.split("/");
                System.out.println(tocElement.elementName + "::" + headers[0] + "::" + headers[1] + "::" + headers[2]);

                tocElement.mainHeading = headers[0];
                tocElement.Heading1 = headers[1];
                tocElement.Heading2 = headers[2];
                tocFinder.getTOCElementList().set(i, tocElement);
            }
        }

        tocFinder.getTOCElementList().stream().map((toce) -> {
            //System.out.println(toce.elementName);
            return toce;
        }).map((toce) -> {
            //System.out.println(toce.refLinkId);
            return toce;
        }).forEach((toce) -> {
            //System.out.println(toce.type + "--" + toce.index);
        });
        if (tocFinder.getTOCElementList().isEmpty()) {
            try {
                throw new Exception("Cannot parse Elements as parser failed!!Pls check original document!!");
            } catch (Exception ex) {
                Logger.getLogger(DocxToDocx.class.getName()).log(Level.SEVERE, null, ex);
            }
        } else {
            tocFinder.getTOCElementList().stream().map((toce) -> {
                //System.out.println(toce.elementName);
                return toce;
            }).map((toce) -> {
                //System.out.println(toce.refLinkId);
                return toce;
            }).forEach((toce) -> {
                //System.out.println(toce.type + "--" + toce.index);
            });

            //Mongo DB update here
            persist(tocFinder.getTOCElementList(), reportId);
        }

    }

    private static void persist(List<TOCElement> tocElementList, String reportId) throws IOException {

        //System.out.println(ResourceBundle.getBundle("mnm.buildmyreport.MnM.properties").getKeys());
        MongoClient client;
        MongoCollection<Document> mongoCollection;
        try {
            BMRProperties bmrp = new BMRProperties();

            client = new MongoClient(bmrp.getPropertyValue("db_host"), Integer.parseInt(bmrp.getPropertyValue("db_port")));
            mongoCollection = client.getDatabase(bmrp.getPropertyValue("db_name")).getCollection("reportcontents");
        } catch (Exception e) {
            client = new MongoClient("54.165.128.223", 27017);
            mongoCollection = client.getDatabase("mnmks").getCollection("reportcontents");
        }
        //TDocument td;
        List<Document> docList = new ArrayList<>();
        tocElementList.stream().map((toce) -> {
            Document doc = new Document();
            //System.out.println(toce.elementName+"::"+toce.type+"::"+toce.refLinkId+"::"+toce.index);
            doc.append("name", toce.elementName).append("type", toce.type).append("reflinkId", toce.refLinkId).append("index", toce.index).append("mainheading", toce.mainHeading).append("heading1", toce.Heading1).append("heading2", toce.Heading2);
            System.out.println(doc.toJson());
            return doc;
        }).forEach((doc) -> {
            docList.add(doc);
        });
        System.out.println("ReportID " + reportId + " was not parsed before. So adding an entry now.");
        DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
        Calendar cal = Calendar.getInstance();
        mongoCollection.insertOne(new Document("reportid", reportId)
                .append("contents", asList(
                                docList)).append("datetime", dateFormat.format(cal.getTime())));
    }

    // TODO code application logic here
    private static List<Object> processStyles(MainDocumentPart documentPart) {
        //System.out.println("Styles used in this doc" + documentPart.getStylesInUse());
        //Set<String> stylesinUse = documentPart.getStylesInUse();
        List<Object> lotTables = new ArrayList<>();

        for (JAXBAssociation jaxbAssociation : allJAXBTblsInDocx) {
            //System.out.println(jaxbAssociation.getDomNode().getPreviousSibling().getNodeType());
            Node previousSibling = jaxbAssociation.getDomNode().getPreviousSibling();
            if (previousSibling != null && previousSibling.getLocalName().equalsIgnoreCase("p")) {
                if (previousSibling.getFirstChild() != null && previousSibling.getFirstChild().getLocalName().equalsIgnoreCase("pPr")) {
                    if (previousSibling.getFirstChild().getFirstChild().getLocalName().equalsIgnoreCase("pstyle")) {
                        if (previousSibling.getFirstChild().getFirstChild().getAttributes().getNamedItem("w:val").getNodeValue().equals("TableTitle")) {
                            lotTables.add(jaxbAssociation.getJaxbObject());

                        }
                    }
                }

            }
        }

        return lotTables;

    }

    public static class TOCElement {

        TOCElement(String elementName, char type, String refLinkId, int index) {
            this.elementName = elementName;
            this.type = type;
            this.refLinkId = refLinkId;
            this.index = index;
        }
        String elementName, refLinkId;
        char type;
        int index;
        String mainHeading, Heading1, Heading2;

    }

    private static void createTocElements() {

        //air and missile 
        String elems = "[\"F\",\"_Toc450152443\",\"1\"],[\"F\",\"_Toc450152444\",\"2\"],[\"F\",\"_Toc450152445\",\"3\"],[\"F\",\"_Toc450152446\",\"4\"],[\"F\",\"_Toc450152447\",\"5\"],[\"F\",\"_Toc450152448\",\"6\"],[\"F\",\"_Toc450152449\",\"7\"],[\"F\",\"_Toc450152450\",\"8\"],[\"F\",\"_Toc450152451\",\"9\"],[\"F\",\"_Toc450152452\",\"10\"],[\"F\",\"_Toc450152453\",\"11\"],[\"F\",\"_Toc450152454\",\"12\"],[\"F\",\"_Toc450152455\",\"13\"],[\"F\",\"_Toc450152456\",\"14\"],[\"F\",\"_Toc450152457\",\"15\"],[\"F\",\"_Toc450152458\",\"16\"],[\"F\",\"_Toc450152459\",\"17\"],[\"F\",\"_Toc450152460\",\"18\"],[\"F\",\"_Toc450152461\",\"19\"],[\"F\",\"_Toc450152462\",\"20\"],[\"F\",\"_Toc450152463\",\"21\"],[\"F\",\"_Toc450152464\",\"22\"],[\"F\",\"_Toc450152465\",\"23\"],[\"F\",\"_Toc450152466\",\"24\"],[\"F\",\"_Toc450152467\",\"25\"],[\"F\",\"_Toc450152468\",\"26\"],[\"F\",\"_Toc450152469\",\"27\"],[\"F\",\"_Toc450152470\",\"28\"],[\"F\",\"_Toc450152471\",\"29\"],[\"F\",\"_Toc450152472\",\"30\"],[\"F\",\"_Toc450152473\",\"31\"],[\"F\",\"_Toc450152474\",\"32\"],[\"F\",\"_Toc450152475\",\"33\"],[\"F\",\"_Toc450152476\",\"34\"],[\"F\",\"_Toc450152477\",\"35\"],[\"F\",\"_Toc450152478\",\"36\"],[\"F\",\"_Toc450152479\",\"37\"],[\"F\",\"_Toc450152480\",\"38\"],[\"F\",\"_Toc450152481\",\"39\"],[\"F\",\"_Toc450152482\",\"40\"],[\"F\",\"_Toc450152483\",\"41\"],[\"F\",\"_Toc450152484\",\"42\"],[\"F\",\"_Toc450152485\",\"43\"],[\"F\",\"_Toc450152486\",\"44\"],[\"F\",\"_Toc450152487\",\"45\"],[\"F\",\"_Toc450152488\",\"46\"],[\"F\",\"_Toc450152489\",\"47\"],[\"F\",\"_Toc450152490\",\"48\"],[\"F\",\"_Toc450152491\",\"49\"],[\"F\",\"_Toc450152492\",\"50\"],[\"F\",\"_Toc450152493\",\"51\"],[\"F\",\"_Toc450152494\",\"52\"],[\"F\",\"_Toc450152495\",\"53\"],[\"F\",\"_Toc450152496\",\"54\"],[\"F\",\"_Toc450152497\",\"55\"],[\"F\",\"_Toc450152498\",\"56\"],[\"F\",\"_Toc450152499\",\"57\"],[\"F\",\"_Toc450152500\",\"58\"],[\"F\",\"_Toc450152501\",\"59\"],[\"F\",\"_Toc450152502\",\"60\"],[\"F\",\"_Toc450152503\",\"61\"],[\"F\",\"_Toc450152504\",\"62\"],[\"F\",\"_Toc450152505\",\"63\"],[\"F\",\"_Toc450152506\",\"64\"],[\"F\",\"_Toc450152507\",\"65\"],[\"F\",\"_Toc450152508\",\"66\"],[\"F\",\"_Toc450152509\",\"67\"],[\"F\",\"_Toc450152510\",\"68\"],[\"F\",\"_Toc450152511\",\"69\"],[\"F\",\"_Toc450152512\",\"70\"],[\"F\",\"_Toc450152513\",\"71\"],[\"F\",\"_Toc450152514\",\"72\"],[\"F\",\"_Toc450152515\",\"73\"],[\"F\",\"_Toc450152516\",\"74\"],[\"F\",\"_Toc450152517\",\"75\"]";

        //String elems = "[\"T\",\"_Toc450152358\",\"1\"],[\"T\",\"_Toc450152355\",\"2\"],[\"T\",\"_Toc450152356\",\"3\"],[\"T\",\"_Toc450152357\",\"4\"]";//[\"T\",\"_Toc450152354\",\"1\"],[\"T\",\"_Toc450152355\",\"2\"],[\"T\",\"_Toc450152356\",\"3\"],[\"T\",\"_Toc450152357\",\"4\"]]";
        //String elems = "[\"T\",\"_Toc450152354\",\"1\"],[\"T\",\"_Toc450152355\",\"2\"],[\"T\",\"_Toc450152356\",\"3\"],[\"T\",\"_Toc450152357\",\"4\"],[\"T\",\"_Toc450152358\",\"5\"],[\"T\",\"_Toc450152359\",\"6\"],[\"T\",\"_Toc450152360\",\"7\"],[\"T\",\"_Toc450152361\",\"8\"],[\"T\",\"_Toc450152362\",\"9\"],[\"T\",\"_Toc450152363\",\"10\"],[\"T\",\"_Toc450152364\",\"11\"],[\"T\",\"_Toc450152365\",\"12\"],[\"T\",\"_Toc450152366\",\"13\"],[\"T\",\"_Toc450152367\",\"14\"],[\"T\",\"_Toc450152368\",\"15\"],[\"T\",\"_Toc450152369\",\"16\"],[\"T\",\"_Toc450152370\",\"17\"],[\"T\",\"_Toc450152371\",\"18\"],[\"T\",\"_Toc450152372\",\"19\"],[\"T\",\"_Toc450152373\",\"20\"],[\"T\",\"_Toc450152374\",\"21\"],[\"T\",\"_Toc450152375\",\"22\"],[\"T\",\"_Toc450152376\",\"23\"],[\"T\",\"_Toc450152377\",\"24\"],[\"T\",\"_Toc450152378\",\"25\"],[\"T\",\"_Toc450152379\",\"26\"],[\"T\",\"_Toc450152380\",\"27\"],[\"T\",\"_Toc450152381\",\"28\"],[\"T\",\"_Toc450152382\",\"29\"],[\"T\",\"_Toc450152383\",\"30\"],[\"T\",\"_Toc450152384\",\"31\"],[\"T\",\"_Toc450152385\",\"32\"],[\"T\",\"_Toc450152386\",\"33\"],[\"T\",\"_Toc450152387\",\"34\"],[\"T\",\"_Toc450152388\",\"35\"],[\"T\",\"_Toc450152389\",\"36\"],[\"T\",\"_Toc450152390\",\"37\"],[\"T\",\"_Toc450152391\",\"38\"],[\"T\",\"_Toc450152392\",\"39\"],[\"T\",\"_Toc450152393\",\"40\"],[\"T\",\"_Toc450152394\",\"41\"],[\"T\",\"_Toc450152395\",\"42\"],[\"T\",\"_Toc450152396\",\"43\"],[\"T\",\"_Toc450152397\",\"44\"],[\"T\",\"_Toc450152398\",\"45\"],[\"T\",\"_Toc450152399\",\"46\"],[\"T\",\"_Toc450152400\",\"47\"],[\"T\",\"_Toc450152401\",\"48\"],[\"T\",\"_Toc450152402\",\"49\"],[\"T\",\"_Toc450152403\",\"50\"],[\"T\",\"_Toc450152404\",\"51\"],[\"T\",\"_Toc450152405\",\"52\"],[\"T\",\"_Toc450152406\",\"53\"],[\"T\",\"_Toc450152407\",\"54\"],[\"T\",\"_Toc450152408\",\"55\"],[\"T\",\"_Toc450152409\",\"56\"],[\"T\",\"_Toc450152410\",\"57\"],[\"T\",\"_Toc450152411\",\"58\"],[\"T\",\"_Toc450152412\",\"59\"],[\"T\",\"_Toc450152413\",\"60\"],[\"T\",\"_Toc450152414\",\"61\"],[\"T\",\"_Toc450152415\",\"62\"],[\"T\",\"_Toc450152416\",\"63\"],[\"T\",\"_Toc450152417\",\"64\"],[\"T\",\"_Toc450152418\",\"65\"],[\"T\",\"_Toc450152419\",\"66\"],[\"T\",\"_Toc450152420\",\"67\"],[\"T\",\"_Toc450152421\",\"68\"],[\"T\",\"_Toc450152422\",\"69\"],[\"T\",\"_Toc450152423\",\"70\"],[\"T\",\"_Toc450152424\",\"71\"],[\"T\",\"_Toc450152425\",\"72\"],[\"T\",\"_Toc450152426\",\"73\"],[\"T\",\"_Toc450152427\",\"74\"],[\"T\",\"_Toc450152428\",\"75\"],[\"T\",\"_Toc450152429\",\"76\"],[\"T\",\"_Toc450152430\",\"77\"],[\"T\",\"_Toc450152431\",\"78\"],[\"T\",\"_Toc450152432\",\"79\"],[\"T\",\"_Toc450152433\",\"80\"],[\"T\",\"_Toc450152434\",\"81\"],[\"T\",\"_Toc450152435\",\"82\"],[\"T\",\"_Toc450152436\",\"83\"],[\"T\",\"_Toc450152437\",\"84\"],[\"T\",\"_Toc450152438\",\"85\"],[\"T\",\"_Toc450152439\",\"86\"],[\"T\",\"_Toc450152440\",\"87\"],[\"T\",\"_Toc450152441\",\"88\"],[\"T\",\"_Toc450152442\",\"89\"]";
        String tocElement = elems.substring(0, elems.length() - 1);
        //System.out.println("tocElements "+tocElement);
        String[] news = tocElement.split("],");
        //System.out.println("split length  is "+news.length);

        String new1;
        Element e;
        int i;
        for (i = 0; i < news.length; i++) {
            new1 = news[i];
            e = new Element(new1.substring(2, 3), new1.substring(6, new1.indexOf(",", 6) - 1), new1.substring(new1.indexOf(",", 6) + 2, new1.length() - 1));
            tocElements.add(e);
            //System.out.println("Element added is " + e.getId()+"::"+e.getType()+"::"+e.getindex());
        }

//        for (Element elm : tocElements) {
//            System.out.println("Element is " + elm.getId()+"::"+elm.getType()+"::"+elm.getindex());
//        }
    }

}
