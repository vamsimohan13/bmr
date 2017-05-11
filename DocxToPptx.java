package mnm.buildmyreport;

import java.io.File;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map.Entry;
import java.util.Objects;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import static mnm.buildmyreport.DocxToDocx.figureElements;
import static mnm.buildmyreport.DocxToDocx.sequenceExportList;
import static mnm.buildmyreport.DocxToDocx.tableElements;
import static mnm.buildmyreport.DocxToXcl.sequenceExportList;

import org.docx4j.TraversalUtil;
import org.docx4j.dml.*;
import org.docx4j.openpackaging.contenttype.ContentTypes;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;

import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.PresentationML.MainPresentationPart;
import org.docx4j.openpackaging.parts.PresentationML.SlideLayoutPart;
import org.docx4j.openpackaging.parts.PresentationML.SlidePart;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MetafileEmfPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.samples.AbstractSample;
//import static org.docx4j.samples.AbstractSample.elementType;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.Hdr;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblGridCol;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.pptx4j.pml.CTGraphicalObjectFrame;
import org.pptx4j.pml.NvPr;
import org.pptx4j.pml.Pic;
import org.pptx4j.pml.Shape;

/**
 * To see what parts comprise your docx, try the PartsList sample.
 *
 * There will always be a MainDocumentPart, usually called document.xml. This
 * sample shows you what objects are in that part.
 *
 * It also shows a general approach for traversing the JAXB object tree in the
 * Main Document part. It can also be applied to headers, footers etc.
 *
 * It is an alternative to XSLT, and doesn't require marshalling/unmarshalling.
 *
 * If many cases, the method getJAXBNodesViaXPath would be more convenient, but
 * there are 3 JAXB bugs which detract from that (see Getting Started).
 *
 * See related classes SingleTraversalUtilVisitorCallback and
 * CompoundTraversalUtilVisitorCallback
 *
 *
 *
 */
public class DocxToPptx extends AbstractSample {

    public static JAXBContext context = org.docx4j.jaxb.Context.jc;
    static int count, figcount = 0;

    //static MainDocumentPart mdpOut;
    static org.docx4j.dml.ObjectFactory dmlFactory;
    static org.pptx4j.pml.ObjectFactory pmlFactory;
    static MainDocumentPart documentPart;
    static PresentationMLPackage presentationMLPackageOut;
    static WordprocessingMLPackage wordMLPackageIn;
    //static BinaryPartAbstractImage bpai;

    static HashMap<Element, Integer> sequenceList = new HashMap<>();
    static HashMap<Integer, Object> sequenceExportList = new HashMap<>();
    static HashMap<String, Element> tableElements = new HashMap<>();
    static HashMap<String, Element> figureElements = new HashMap<>();
    static String header, outfilename = "";


    public static void main(String[] args) throws Exception {

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
        try {
            getInputFilePath(args);
        } catch (IllegalArgumentException e) {
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/ICT - Backend as A Service (Baas) Market - Global Forecast To 2020.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/report_1467732928.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Masked - Cardiovascular Information System Market – Forecasts to 2020.docx";

            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Mobile 3D Market.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Casino Management Systems (CMS) Market.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Organic Electronics Market - Global Analysis and Forecast 2020.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Data Center Networking.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Outdoor Wi-Fi Market.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Power Quality Meter Market - Global Forecast & Trends To 2021.docx";
            inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Air and Missile.docx";
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

            /*export mode with some test data starts*/
            mode = "export";
            createTocElements();
//            tocElements.add(new Element("T", "_Toc445308482", "47"));//cardiovascular
//            tocElements.add(new Element("T", "_Toc445308483", "48"));//cardiovascular
//            
//            tocElements.add(new Element("T", "_Toc445308484", "47"));//cardiovascular
//            tocElements.add(new Element("T", "_Toc445308485", "48"));//cardiovascular
//            
//            tocElements.add(new Element("T", "_Toc445308486", "47"));//cardiovascular
//            tocElements.add(new Element("T", "_Toc445308487", "48"));//cardiovascular
//            
//            tocElements.add(new Element("T", "_Toc445308488", "47"));//cardiovascular
//            tocElements.add(new Element("T", "_Toc445308489", "48"));//cardiovascular
//            
//            tocElements.add(new Element("T", "_Toc445308490", "47"));//cardiovascular
//            tocElements.add(new Element("T", "_Toc445308491", "48"));//cardiovascular
//            tocElements.add(new Element("T", "_Toc445308564", "47"));//cardiovascular
//            tocElements.add(new Element("T", "_Toc445308565", "48"));//cardiovascular
//            tocElements.add(new Element("T", "_Toc401326363", "48"));//Organic Electronics Market
//            tocElements.add(new Element("F", "_Toc356558916", "48"));//Outdoor Wi-Fi Market
//            tocElements.add(new Element("F", "_Toc356558917", "48"));//Outdoor Wi-Fi Market
//            tocElements.add(new Element("F", "_Toc356558858", "48"));//Outdoor Wi-Fi Market
//
//            tocElements.add(new Element("F", "_Toc348352919", "48"));//Mobile 3d market
//            tocElements.add(new Element("F", "_Toc348352920", "48"));//Mobile 3d market
//            tocElements.add(new Element("F", "_Toc348352921", "48"));//Mobile 3d market
//            tocElements.add(new Element("F", "_Toc348352922", "48"));//Mobile 3d market
//            tocElements.add(new Element("F", "_Toc348352923", "48"));//Mobile 3d market
//            tocElements.add(new Element("F", "_Toc348352924", "48"));//Mobile 3d market

//            tocElements.add(new Element("T", "_Toc401326364", "48"));//Organic Electronics Market
//              tocElements.add(new Element("F", "_Toc369865508", "48"));//Data Center Networking
            /*export mode with some test data ends*/
        }
        //System.out.println(System.getProperty("classpath"));
        System.out.println("inputfilepath " + inputfilepath);

        System.out.println("user.dir is " + System.getProperty("user.dir"));

        wordMLPackageIn = WordprocessingMLPackage
                .load(new java.io.File(inputfilepath));
        documentPart = wordMLPackageIn.getMainDocumentPart();

        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart
                .getJaxbElement();

        Body body = wmlDocumentEl.getBody();
        boolean exportAllTables = false;

        int seq = 0;
        for (Element currelement : tocElements) {
            if (currelement.getType().equals("T")) {
                tableElements.put(currelement.getId(), currelement);
            }
            if (currelement.getType().equals("F")) {
                figureElements.put(currelement.getId(), currelement);
            }
            //create a Map with index/sequence as key and element as value. this is needed to preserve order of selection
            sequenceList.put(currelement, seq++);
        }
        
        List<Hdr> hdrlist = new ArrayList<Hdr>();
        //System.out.println("Too Good so far!!!!!!!!!1");
        // Uncomment to see the raw XML
        //System.out.println(XmlUtils.marshaltoString(documentPart.getJaxbElement(), true, true));
        RelationshipsPart rp = documentPart.getRelationshipsPart();
        if (rp != null) {
            rp.getRelationships().getRelationship().stream().map((r) -> rp.getPart(r)).forEach((part) -> {
                if (part instanceof HeaderPart) {

                    hdrlist.add(((HeaderPart) part).getJaxbElement());
                    //finder.walkJAXBElements(hdr);
                }
            });
        }
        header = Utils.getReportTitle(hdrlist);

        if (header == null || header.isEmpty()) {
            outfilename = "MnM Report";
        } else {
            outfilename = header;
        }
        //String outputfilepath = System.getProperty("user.dir") + "/output/OUT_Table.pptx";

        // Create skeletal package, including a MainPresentationPart and a SlideLayoutPart
        presentationMLPackageOut = PresentationMLPackage.createPackage();

        //final PresentationMLPackage presentationMLPackageIn = PresentationMLPackage.load(new java.io.File(inputfilepath2));
        //final PresentationMLPackage presentationMLPackageOut = (PresentationMLPackage) presentationMLPackageIn.clone();
        // Need references to these parts to create a slide
        // Please note that these parts *already exist* - they are
        // created by createPackage() above.  See that method
        // for instruction on how to create and add a part.
        final MainPresentationPart pp = (MainPresentationPart) presentationMLPackageOut.getParts().getParts().get(
                new PartName("/ppt/presentation.xml"));

        final SlideLayoutPart layoutPart = (SlideLayoutPart) presentationMLPackageOut.getParts().getParts().get(
                new PartName("/ppt/slideLayouts/slideLayout1.xml"));

        if (!tableElements.isEmpty()) {
            TableExporterNew tableExporter = new TableExporterNew();
            new TraversalUtil(body, tableExporter);
            filterTables(tableExporter.getPTablePairs());

        }
        if (!figureElements.isEmpty()) {
            //FigureExporter figureExporter = new FigureExporter();
            FigureExporterNew figureExporter = new FigureExporterNew();

            new TraversalUtil(body, figureExporter);
            filterFigures(figureExporter.getPFigurePairs());

        }

        if (!sequenceExportList.isEmpty()) {
            for (int i = 0; i < sequenceExportList.size(); i++) {
                count++;
                if (sequenceExportList.containsKey(i)) {
                    if (sequenceExportList.get(i).getClass().getName().equals("mnm.buildmyreport.PTablePair")) {
                        PTablePair ptp = (PTablePair) sequenceExportList.get(i);

                        //tblcount++;
                        // OK, now we can create a slide
                        SlidePart slidePart = presentationMLPackageOut.createSlidePart(pp, layoutPart,
                                new PartName("/ppt/slides/slide" + count + ".xml"));

                        // Lets add title
                        slidePart.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add(getSpTitle(ptp.title, "Table " + ptp.getIndex() + " "));
                        // Lets add table
                        slidePart.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add(getGraphicFrame(ptp.tbl));
                        //slidePart.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add(getGraphicFrame2());
                        //Lets add footnote
                        if (ptp.footer != null) {
                            for (int j = 0; j < ptp.footer.size(); j++) {
                            slidePart.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add(getSpFooter(ptp.footer.get(i), i));
                            }
                        }
                    }
                    if (sequenceExportList.get(i).getClass().getName().equals("mnm.buildmyreport.PFigurePair")) {
                        PFigurePair pfp = (PFigurePair) sequenceExportList.get(i);
                        figcount++;
                        SlidePart slidePart = presentationMLPackageOut.createSlidePart(pp, layoutPart,
                                new PartName("/ppt/slides/slide" + count + ".xml"));
                        // Lets add title

                        slidePart.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add(getSpTitle(pfp.title, "Figure " + pfp.getIndex() + " "));
                        // Lets add table

                        try {
                            slidePart.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add(getPic(pfp.figure, slidePart));
                        } catch (NullPointerException e) {
                            System.out.println("Got null fig for" + pfp.title);
                        }
                        //Lets add footnote
                        if (pfp.footer != null) {
                            for (int j = 0; j < pfp.footer.size(); j++) {
                                slidePart.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add(getSpFooter(pfp.footer.get(j), j));
                            }

                        }
                        System.out.println("exported fig::" + pfp.title);
                    }

                }
            }
        }

        System.out.println("Too Good so far!!!!!!!!!1");

        System.out.println("Time to export to pptx, we have parsed a total of" + sequenceExportList.size() + " tables and(or) figures");
        System.out.println(inputfilepath);
        CharSequence outputfoldername = inputfilepath.subSequence(inputfilepath.lastIndexOf('/') + 1, inputfilepath.lastIndexOf('.'));
        String outputfilepath = System.getProperty("user.dir") + "/output/" + outputfoldername;
        String outputfilename = outfilename + ".pptx";
        File files = new File(outputfilepath);
        if (!files.exists()) {
            if (files.mkdirs()) {
                System.out.println("Multiple directories are created! :" + outputfilepath);
                //System.out.println(outputfilepath);
            } else {
                //files.delete();                
                System.out.println("Failed to create multiple directories!");
            }
        }
        //saver.save(outputfilepath + "/" + outputfilename);

        presentationMLPackageOut.save(new java.io.File(outputfilepath+ "/" + outputfilename));

        System.out.println("\n\n done .. saved " + outputfilepath);

    }

    private static void filterTables(List<PTablePair> pTablePairs) {
        for (PTablePair ptp : pTablePairs) {
            for (int j = 0; j < ptp.title.getContent().size(); j++) {
                if (ptp.title.getContent().get(j) instanceof javax.xml.bind.JAXBElement) {
                    JAXBElement jaxb = (JAXBElement) ptp.title.getContent().get(j);
                    //org.docx4j.wml.P
                    if (jaxb.getDeclaredType().getName().equals("org.docx4j.wml.CTBookmark")) {

                        String tableId = ((org.docx4j.wml.CTBookmark) (jaxb.getValue())).getName();
                        if (tableElements.containsKey(tableId)) {

                            ptp.setIndex(((Element) tableElements.get(tableId)).getindex());
                            sequenceExportList.put(sequenceList.get((Element) tableElements.get(tableId)), ptp);
                        }
                    }
                }
            }
        }
    }

    private static CTGraphicalObjectFrame getGraphicFrame(Tbl docxTbl) throws JAXBException {
        // instatiation the factory for later object creation.
        dmlFactory = new org.docx4j.dml.ObjectFactory();
        pmlFactory = new org.pptx4j.pml.ObjectFactory();

        // Node Creation
        CTGraphicalObjectFrame graphicFrame = pmlFactory
                .createCTGraphicalObjectFrame();
        org.pptx4j.pml.CTGraphicalObjectFrameNonVisual nvGraphicFramePr = pmlFactory
                .createCTGraphicalObjectFrameNonVisual();
        org.docx4j.dml.CTNonVisualDrawingProps cNvPr = dmlFactory
                .createCTNonVisualDrawingProps();
        org.docx4j.dml.CTNonVisualGraphicFrameProperties cNvGraphicFramePr = dmlFactory
                .createCTNonVisualGraphicFrameProperties();
        org.docx4j.dml.CTGraphicalObjectFrameLocking graphicFrameLocks = new org.docx4j.dml.CTGraphicalObjectFrameLocking();
        org.docx4j.dml.CTTransform2D xfrm = dmlFactory.createCTTransform2D();
        Graphic graphic = dmlFactory.createGraphic();
        GraphicData graphicData = dmlFactory.createGraphicData();

        // Build the parent-child relationship of this slides.xml
        graphicFrame.setNvGraphicFramePr(nvGraphicFramePr);
        nvGraphicFramePr.setCNvPr(cNvPr);
        //cNvPr.setName("1");
        nvGraphicFramePr.setCNvGraphicFramePr(cNvGraphicFramePr);
        cNvGraphicFramePr.setGraphicFrameLocks(graphicFrameLocks);
        graphicFrameLocks.setNoGrp(true);
        nvGraphicFramePr.setNvPr(pmlFactory.createNvPr());

        graphicFrame.setXfrm(xfrm);

        CTPositiveSize2D ext = dmlFactory.createCTPositiveSize2D();
        ext.setCx(9000000);
        ext.setCy(865760);

        xfrm.setExt(ext);

        CTPoint2D off = dmlFactory.createCTPoint2D();
        xfrm.setOff(off);
        off.setX(1430090);
        off.setY(2892275);

        graphicFrame.setGraphic(graphic);

        graphic.setGraphicData(graphicData);
        graphicData
                .setUri("http://schemas.openxmlformats.org/drawingml/2006/table");

        CTTable ctTable = dmlFactory.createCTTable();
        JAXBElement<CTTable> tbl = dmlFactory.createTbl(ctTable);
        ///// From here we build graphicData ctTbl from docxTbl

        graphicData.getAny().add(tbl);

        CTTableGrid ctTableGrid = dmlFactory.createCTTableGrid();
        CTTableCol gridCol = dmlFactory.createCTTableCol();
        org.docx4j.dml.CTTableProperties ctTableProperties = dmlFactory.createCTTableProperties();

        // now set tablegrid from dcxtbl
        ctTable.setTblGrid(ctTableGrid);
//sum should be in the range of 6000000
        for (TblGridCol gridCol1 : docxTbl.getTblGrid().getGridCol()) {
            ctTableGrid.getGridCol().add(gridCol);
            gridCol.setW(1500000);
        }

        //now add rows but first determine how many rows
        // so first get only rows (type Tr in the content of tbl)
        List<org.docx4j.wml.Tr> docxTblRows = new ArrayList<org.docx4j.wml.Tr>();
        for (int i = 0; i < docxTbl.getContent().size(); i++) {
            if ((docxTbl.getContent().get(i)) instanceof org.docx4j.wml.Tr) {
                docxTblRows.add((org.docx4j.wml.Tr) docxTbl.getContent().get(i));
            }

        }

        for (Tr docxTblRow : docxTblRows) {
            CTTableRow ctTableRow = dmlFactory.createCTTableRow();
            ctTableRow.setH(250000);
            int incr = 0;
            for (int j = 0, k = 0; j < ctTableGrid.getGridCol().size(); k++) {
                JAXBElement jaxbTr = (JAXBElement) (docxTblRow.getContent().get(k));
                if (jaxbTr.getDeclaredType().getName().equals("org.docx4j.wml.Tc")) {

                    Tc tc = (Tc) jaxbTr.getValue();

                    incr = (tc.getTcPr().getGridSpan() != null) ? tc.getTcPr().getGridSpan().getVal().intValue() : 1;
                    j = j + incr;
                    ctTableRow.getTc().add(getCell(tc));
                }

            }
            ctTable.getTr().add(ctTableRow);
        }

        //ctTableProperties.setTableStyle(pptxTblStyle);
        //this styleId is hardcoded and is equal to styled in tableStyles.xml as specified by property pptx4j.openpackaging.packages.PresentationMLPackage.DefaultTableStyle
        ctTableProperties.setTableStyleId("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
        ctTableProperties.setBandRow(Boolean.TRUE);
        ctTableProperties.setFirstCol(Boolean.TRUE);
        ctTableProperties.setFirstRow(Boolean.TRUE);

        //Uncomment below when evrything is stable else PPTX will not open even!!!
        ctTable.setTblPr(ctTableProperties);
        //System.out.println(XmlUtils.marshaltoString(graphicFrame));
        return graphicFrame;
    }

    private static CTTableCell getCell(org.docx4j.wml.Tc tc) throws JAXBException {

        CTTableCell ctTableCell = dmlFactory.createCTTableCell();
        // Create object for tcPr
        CTTableCellProperties tablecellproperties = dmlFactory.createCTTableCellProperties();
        ctTableCell.setTcPr(tablecellproperties);
        tablecellproperties.setAnchor(org.docx4j.dml.STTextAnchoringType.T);
        tablecellproperties.setMarL(new Integer(91440));
        tablecellproperties.setMarR(new Integer(91440));
        tablecellproperties.setMarT(new Integer(45720));
        tablecellproperties.setMarB(new Integer(45720));
        tablecellproperties.setVert(org.docx4j.dml.STTextVerticalType.HORZ);
        tablecellproperties.setHorzOverflow(org.docx4j.dml.STTextHorzOverflowType.CLIP);
        ctTableCell.setGridSpan(new Integer(1));
        // tcPr done
        // Create object for txBody
        CTTextBody ctTextBody = dmlFactory.createCTTextBody();
        ctTableCell.setTxBody(ctTextBody);
        //CTRegularTextRun ctRegularTextRun = dmlFactory.createCTRegularTextRun();
        //String celltext = "";
        for (int k = 0; k < tc.getContent().size(); k++) {

            CTTextParagraph ctTextParagraph = dmlFactory.createCTTextParagraph();

            org.docx4j.wml.P tcpara = (org.docx4j.wml.P) tc.getContent().get(k);

            String rowtext = "";
            for (int l = 0; l < tcpara.getContent().size(); l++) {
                //Dont assume its always a row
                org.docx4j.wml.R r;
                org.docx4j.wml.PPr pPr;

                //CTTextParagraphProperties ctTextParagraphProperties = dmlFactory.createCTTextParagraphProperties();
                if ((tcpara.getPPr() != null)) {
                    pPr = tcpara.getPPr();
                    //insert MnM bullet style if pPr had pStyle with value like %bullet%
                    //this is a hack as of now!!
                    if (pPr.getPStyle() != null) {
                        if (pPr.getPStyle().getVal().contains("bullet")) {
                            // Create object for pPr
                            CTTextParagraphProperties textparagraphproperties = dmlFactory.createCTTextParagraphProperties();
                            textparagraphproperties.setIndent(new Integer(-171450));
                            textparagraphproperties.setMarL(new Integer(171450));
                            // Create object for buFont
                            TextFont textfont = dmlFactory.createTextFont();
                            textparagraphproperties.setBuFont(textfont);
                            textfont.setTypeface("Arial");
                            textfont.setPanose("020B0604020202020204");
                            textfont.setPitchFamily(Byte.parseByte("34"));
                            textfont.setCharset(Byte.parseByte("0"));
                            textparagraphproperties.setBuFont(textfont);
                            // Create object for buChar
                            CTTextCharBullet textcharbullet = dmlFactory.createCTTextCharBullet();
                            textparagraphproperties.setBuChar(textcharbullet);
                            textcharbullet.setChar("•");
                            ctTextParagraph.setPPr(textparagraphproperties);
                        }
                    }

                }
                if ((tcpara.getContent().get(l) instanceof org.docx4j.wml.R)) {

                    r = (org.docx4j.wml.R) (tcpara.getContent().get(l));
                    for (int m = 0; m < r.getContent().size(); m++) {
                        Object o = r.getContent().get(m);
                        if ((o instanceof org.docx4j.wml.Br
                                || o instanceof org.docx4j.wml.R.Tab
                                || o instanceof org.docx4j.wml.R.LastRenderedPageBreak)) {
                            break;
                        }
                        javax.xml.bind.JAXBElement jaxb = (javax.xml.bind.JAXBElement) (o);
                        switch (jaxb.getDeclaredType().getName()) {
                            //// also check if images or drwings are the
                            case "org.docx4j.wml.Text":
                                //
                                rowtext = rowtext + ((org.docx4j.wml.Text) (jaxb.getValue())).getValue();
                                // Create object for r
                                CTRegularTextRun regulartextrun = dmlFactory.createCTRegularTextRun();
                                //ctTextParagraph.getEGTextRun().add(regulartextrun);
                                // Create object for rPr
                                CTTextCharacterProperties textcharacterproperties = dmlFactory.createCTTextCharacterProperties();
                                regulartextrun.setRPr(textcharacterproperties);
                                textcharacterproperties.setLang("en-US");
                                textcharacterproperties.setSz(new Integer(1200));
                                textcharacterproperties.setSmtId(new Long(0));
                                regulartextrun.setT(rowtext);

                                //ctTextParagraph.getEGTextRun().add(regulartextrun);
                                CTTextCharacterProperties ctTextCharacterProperties = dmlFactory.createCTTextCharacterProperties();
                                ctTextCharacterProperties.setLang("en-US");
                                ctTextCharacterProperties.setSz(1200);

                                regulartextrun.setRPr(ctTextCharacterProperties);
                                ctTextParagraph.getEGTextRun().add(regulartextrun);
                                break;
                            case "org.docx4j.wml.Drawing":
                                org.docx4j.wml.Drawing drawing = (org.docx4j.wml.Drawing) (jaxb.getValue());
                                org.docx4j.dml.wordprocessingDrawing.Inline inline = (org.docx4j.dml.wordprocessingDrawing.Inline) (drawing).getAnchorOrInline().get(0);
                                if (inline.getGraphic() != null) {
                                    org.docx4j.dml.Graphic graphic = inline.getGraphic();
                                    if (graphic.getGraphicData() != null) {
                                        String imageId = graphic.getGraphicData().getPic().getBlipFill().getBlip().getEmbed();
                                    }
                                }
                                break;
                        }

                    }

                }

            }
            ctTextBody.getP().add(ctTextParagraph);
            ctTextBody.setBodyPr(dmlFactory.createCTTextBodyProperties());
        }
        ctTableCell.setTxBody(ctTextBody);
        return ctTableCell;
    }

    private static Shape getSpTitle(P ptp, String index) {

        org.pptx4j.pml.ObjectFactory pmlObjectFactory = new org.pptx4j.pml.ObjectFactory();

        Shape shape = pmlObjectFactory.createShape();
        // Create object for nvSpPr
        Shape.NvSpPr shapenvsppr = pmlObjectFactory.createShapeNvSpPr();
        shape.setNvSpPr(shapenvsppr);
        org.docx4j.dml.ObjectFactory dmlObjectFactory = new org.docx4j.dml.ObjectFactory();
        // Create object for cNvPr
        CTNonVisualDrawingProps nonvisualdrawingprops = dmlObjectFactory.createCTNonVisualDrawingProps();
        shapenvsppr.setCNvPr(nonvisualdrawingprops);
        nonvisualdrawingprops.setDescr("");
        nonvisualdrawingprops.setName("TextBox 4");
        nonvisualdrawingprops.setId(5);
        // Create object for cNvSpPr
        CTNonVisualDrawingShapeProps nonvisualdrawingshapeprops = dmlObjectFactory.createCTNonVisualDrawingShapeProps();
        shapenvsppr.setCNvSpPr(nonvisualdrawingshapeprops);
        // Create object for nvPr
        NvPr nvpr = pmlObjectFactory.createNvPr();
        shapenvsppr.setNvPr(nvpr);
        // Create object for spPr
        CTShapeProperties shapeproperties = dmlObjectFactory.createCTShapeProperties();
        shape.setSpPr(shapeproperties);
        // Create object for noFill
        CTNoFillProperties nofillproperties = dmlObjectFactory.createCTNoFillProperties();
        shapeproperties.setNoFill(nofillproperties);
        // Create object for xfrm
        CTTransform2D transform2d = dmlObjectFactory.createCTTransform2D();
        shapeproperties.setXfrm(transform2d);
        // Create object for ext
        CTPositiveSize2D positivesize2d = dmlObjectFactory.createCTPositiveSize2D();
        transform2d.setExt(positivesize2d);
        positivesize2d.setCx(8390238);
        positivesize2d.setCy(646331);
        transform2d.setRot(new Integer(0));
        // Create object for off
        CTPoint2D point2d = dmlObjectFactory.createCTPoint2D();
        transform2d.setOff(point2d);
        point2d.setY(245202);
        point2d.setX(1430090);
        // Create object for prstGeom
        CTPresetGeometry2D presetgeometry2d = dmlObjectFactory.createCTPresetGeometry2D();
        shapeproperties.setPrstGeom(presetgeometry2d);
        // Create object for avLst
        CTGeomGuideList geomguidelist = dmlObjectFactory.createCTGeomGuideList();
        presetgeometry2d.setAvLst(geomguidelist);
        presetgeometry2d.setPrst(org.docx4j.dml.STShapeType.RECT);
        // Create object for txBody
        CTTextBody textbody = dmlObjectFactory.createCTTextBody();
        shape.setTxBody(textbody);
        // Create object for bodyPr
        CTTextBodyProperties textbodyproperties = dmlObjectFactory.createCTTextBodyProperties();
        textbody.setBodyPr(textbodyproperties);
        textbodyproperties.setWrap(org.docx4j.dml.STTextWrappingType.SQUARE);
        // Create object for spAutoFit
        CTTextShapeAutofit textshapeautofit = dmlObjectFactory.createCTTextShapeAutofit();
        textbodyproperties.setSpAutoFit(textshapeautofit);
        // Create object for lstStyle
        CTTextListStyle textliststyle = dmlObjectFactory.createCTTextListStyle();
        textbody.setLstStyle(textliststyle);
        // Create object for title
        CTTextParagraph textparagraph = dmlObjectFactory.createCTTextParagraph();
        textbody.getP().add(textparagraph);
        // Create object for pPr
        CTTextParagraphProperties textparagraphproperties = dmlObjectFactory.createCTTextParagraphProperties();
        textparagraph.setPPr(textparagraphproperties);
        textparagraphproperties.setLvl(0);
        // Create object for r
        CTRegularTextRun regulartextrun = dmlObjectFactory.createCTRegularTextRun();
        textparagraph.getEGTextRun().add(regulartextrun);
        // Create object for rPr
        CTTextCharacterProperties textcharacterproperties = dmlObjectFactory.createCTTextCharacterProperties();
        regulartextrun.setRPr(textcharacterproperties);
        // Create object for latin
        TextFont textfont = dmlObjectFactory.createTextFont();
        textcharacterproperties.setLatin(textfont);
        textfont.setTypeface("Franklin Gothic Medium Cond");
        textfont.setPanose("020B0606030402020204");
        textfont.setPitchFamily(Byte.decode("34"));
        textfont.setCharset(Byte.decode("0"));
        // Create object for ea
        TextFont textfont2 = dmlObjectFactory.createTextFont();
        textcharacterproperties.setEa(textfont2);
        textfont2.setTypeface("Times New Roman");
        textfont2.setPanose("02020603050405020304");
        textfont2.setPitchFamily(Byte.decode("18"));
        textfont2.setCharset(Byte.decode("0"));
        // Create object for cs
        TextFont textfont3 = dmlObjectFactory.createTextFont();
        textcharacterproperties.setCs(textfont3);
        textfont3.setTypeface("Times New Roman");
        textfont3.setPanose("02020603050405020304");
        textfont3.setPitchFamily(Byte.decode("18"));
        textfont3.setCharset(Byte.decode("0"));
        textcharacterproperties.setLang("en-US");
        textcharacterproperties.setU(org.docx4j.dml.STTextUnderlineType.WORDS);
        // Create object for solidFill
        CTSolidColorFillProperties solidcolorfillproperties = dmlObjectFactory.createCTSolidColorFillProperties();
        textcharacterproperties.setSolidFill(solidcolorfillproperties);
        // Create object for srgbClr
        CTSRgbColor srgbcolor = dmlObjectFactory.createCTSRgbColor();
        solidcolorfillproperties.setSrgbClr(srgbcolor);

        textcharacterproperties.setCap(org.docx4j.dml.STTextCapsType.ALL);
        textcharacterproperties.setSmtId(new Long(0));
        regulartextrun.setT(index + getText(ptp).toUpperCase());
        // Create object for br
        CTTextLineBreak textlinebreak = dmlObjectFactory.createCTTextLineBreak();
        textparagraph.getEGTextRun().add(textlinebreak);
        // Create object for rPr
        CTTextCharacterProperties textcharacterproperties2 = dmlObjectFactory.createCTTextCharacterProperties();
        textlinebreak.setRPr(textcharacterproperties2);
        // Create object for latin
        TextFont textfont4 = dmlObjectFactory.createTextFont();
        textcharacterproperties2.setLatin(textfont4);
        textfont4.setTypeface("Franklin Gothic Medium Cond");
        textfont4.setPanose("020B0606030402020204");
        textfont4.setPitchFamily(Byte.decode("34"));
        textfont4.setCharset(Byte.decode("0"));
        // Create object for ea
        TextFont textfont5 = dmlObjectFactory.createTextFont();
        textcharacterproperties2.setEa(textfont5);
        textfont5.setTypeface("Times New Roman");
        textfont5.setPanose("02020603050405020304");
        textfont5.setPitchFamily(Byte.decode("18"));
        textfont5.setCharset(Byte.decode("0"));
        // Create object for cs
        TextFont textfont6 = dmlObjectFactory.createTextFont();
        textcharacterproperties2.setCs(textfont6);
        textfont6.setTypeface("Times New Roman");
        textfont6.setPanose("02020603050405020304");
        textfont6.setPitchFamily(Byte.decode("18"));
        textfont6.setCharset(Byte.decode("0"));
        textcharacterproperties2.setLang("en-US");
        textcharacterproperties2.setU(org.docx4j.dml.STTextUnderlineType.WORDS);
        // Create object for solidFill
        CTSolidColorFillProperties solidcolorfillproperties2 = dmlObjectFactory.createCTSolidColorFillProperties();
        textcharacterproperties2.setSolidFill(solidcolorfillproperties2);
        // Create object for srgbClr
        CTSRgbColor srgbcolor2 = dmlObjectFactory.createCTSRgbColor();
        solidcolorfillproperties2.setSrgbClr(srgbcolor2);

        textcharacterproperties2.setCap(org.docx4j.dml.STTextCapsType.ALL);
        textcharacterproperties2.setSmtId(new Long(0));

        // Create object for title
        CTTextParagraph textparagraph2 = dmlObjectFactory.createCTTextParagraph();
        textbody.getP().add(textparagraph2);
        // Create object for endParaRPr
        CTTextCharacterProperties textcharacterproperties6 = dmlObjectFactory.createCTTextCharacterProperties();
        textparagraph2.setEndParaRPr(textcharacterproperties6);
        textcharacterproperties6.setLang("en-US");
        textcharacterproperties6.setSmtId(new Long(0));

        return shape;
    }

    private static Shape getSp(P p) {
        // instatiation the factory for later object creation.
        dmlFactory = new org.docx4j.dml.ObjectFactory();
        pmlFactory = new org.pptx4j.pml.ObjectFactory();

        // Node Creation
        Shape sp = pmlFactory
                .createShape();

        //sp.setSpPr(null);
        org.docx4j.dml.CTNonVisualDrawingProps cNvPr = dmlFactory
                .createCTNonVisualDrawingProps();

        CTNonVisualDrawingShapeProps ctNonVisualDrawingShapeProps = dmlFactory.createCTNonVisualDrawingShapeProps();
        //ctNonVisualDrawingShapeProps.
        NvPr nvPr = pmlFactory.createNvPr();
        cNvPr.setId(1);
        cNvPr.setName("Heading");
        Shape.NvSpPr nvSpPr = new Shape.NvSpPr();
        //nvSpPr = new Shape.NvSpPr().setCNvPr(cNvPr);

        nvSpPr.setCNvPr(cNvPr);
        nvSpPr.setCNvSpPr(ctNonVisualDrawingShapeProps);
        nvSpPr.setNvPr(nvPr);
        sp.setNvSpPr(nvSpPr);

        org.docx4j.dml.CTTransform2D xfrm = dmlFactory.createCTTransform2D();
        CTPositiveSize2D ext = dmlFactory.createCTPositiveSize2D();
        ext.setCx(1961345);
        ext.setCy(2555928);

        xfrm.setExt(ext);

        CTPoint2D off = dmlFactory.createCTPoint2D();

        off.setX(4953000);
        off.setY(535531);
        xfrm.setOff(off);
        CTShapeProperties ctShapeProperties = dmlFactory.createCTShapeProperties();

        ctShapeProperties.setXfrm(xfrm);

        CTPresetGeometry2D ctPresetGeometry2D = dmlFactory.createCTPresetGeometry2D();
        ctPresetGeometry2D.setPrst(STShapeType.RECT);
        CTGeomGuide ctGeomGuide = dmlFactory.createCTGeomGuide();
        CTGeomGuideList ctGeomGuideList = new CTGeomGuideList();
        ctGeomGuideList.getGd().add(ctGeomGuide);
        ctPresetGeometry2D.setAvLst(ctGeomGuideList);
        ctShapeProperties.setPrstGeom(ctPresetGeometry2D);
        sp.setSpPr(ctShapeProperties);

        CTTextBody ctTextBody = dmlFactory.createCTTextBody();
        CTTextBodyProperties ctTextBodyProperties = dmlFactory.createCTTextBodyProperties();
        CTTextShapeAutofit ctTextShapeAutofit = dmlFactory.createCTTextShapeAutofit();
        ctTextBodyProperties.setSpAutoFit(ctTextShapeAutofit);
        ctTextBodyProperties.setWrap(STTextWrappingType.NONE);
        ctTextBody.setBodyPr(ctTextBodyProperties);
        CTTextListStyle lstStyle = dmlFactory.createCTTextListStyle();
        ctTextBody.setLstStyle(lstStyle);
        //List<CTTextParagraph> ctTextParagraph = ctTextBody.getP();

        CTTextParagraphProperties ctTextParagraphProperties = dmlFactory.createCTTextParagraphProperties();
        CTTextSpacing spcBef = dmlFactory.createCTTextSpacing();
        CTTextSpacingPoint ctTextSpacingPointspcBef = dmlFactory.createCTTextSpacingPoint();
        ctTextSpacingPointspcBef.setVal(600);
        spcBef.setSpcPts(ctTextSpacingPointspcBef);
        CTTextSpacing spcAft = dmlFactory.createCTTextSpacing();
        CTTextSpacingPoint ctTextSpacingPointspcAft = dmlFactory.createCTTextSpacingPoint();
        ctTextSpacingPointspcAft.setVal(600);
        spcAft.setSpcPts(ctTextSpacingPointspcAft);
        ctTextParagraphProperties.setSpcAft(spcAft);
        ctTextParagraphProperties.setSpcBef(spcBef);

        ctTextParagraphProperties.setAlgn(STTextAlignType.L);
        ctTextParagraphProperties.setMarL(342900);
        ctTextParagraphProperties.setMarR(0);
        ctTextParagraphProperties.setIndent(-342900);
        ctTextParagraphProperties.setLvl(0);
        //ctTextParagraphProperties.sete
        CTTextSpacing ctTextSpacing = dmlFactory.createCTTextSpacing();
        CTTextSpacingPoint ctTextSpacingPoint = dmlFactory.createCTTextSpacingPoint();
        ctTextSpacingPoint.setVal(120000);
        ctTextSpacing.setSpcPts(ctTextSpacingPoint);
        ctTextParagraphProperties.setLnSpc(ctTextSpacing);
        //ctTextParagraphProperties.setSpcBef(null);
        //ctTextParagraphProperties.setSpcAft(null);
        CTColor ctColor = dmlFactory.createCTColor();
        CTSRgbColor CTSRgbColor = dmlFactory.createCTSRgbColor();
        CTSRgbColor.setVal(new BigInteger("E36C0A", 16).toByteArray());
        //javax.xml.bind.DatatypeConverter;
        //Integer.decode("E36C0A").byteValue()
        ctColor.setSrgbClr(CTSRgbColor);
        ctTextParagraphProperties.setBuClr(ctColor);

        TextFont textFont = dmlFactory.createTextFont();
        textFont.setCharset(Byte.decode("0"));
        textFont.setPanose("020B0706030402020204");
        textFont.setPitchFamily(Byte.decode("34"));
        textFont.setTypeface("Franklin Gothic Demi Cond");
        ctTextParagraphProperties.setBuFont(textFont);

        CTTextAutonumberBullet ctTextAutonumberBullet = dmlFactory.createCTTextAutonumberBullet();
        ctTextAutonumberBullet.setType(STTextAutonumberScheme.ARABIC_PERIOD);
        ctTextParagraphProperties.setBuAutoNum(ctTextAutonumberBullet);

        CTTextCharacterProperties ctTextCharacterProperties = dmlFactory.createCTTextCharacterProperties();
        ctTextCharacterProperties.setLang("en-US");
        ctTextCharacterProperties.setSz(1200);
        ctTextCharacterProperties.setDirty(Boolean.FALSE);
        ctTextCharacterProperties.setCap(STTextCapsType.ALL);
        ctTextCharacterProperties.setU(STTextUnderlineType.WORDS);

        CTSolidColorFillProperties ctSolidColorFillProperties = dmlFactory.createCTSolidColorFillProperties();
        //CTColor ctColor2 = dmlFactory.createCTColor();
        CTSRgbColor CTSRgbColor2 = dmlFactory.createCTSRgbColor();
        CTSRgbColor2.setVal(new BigInteger("0D5775", 16).toByteArray());
        ctSolidColorFillProperties.setSrgbClr(CTSRgbColor2);

        ctTextCharacterProperties.setSolidFill(ctSolidColorFillProperties);

        CTRegularTextRun ctRegularTextRun = dmlFactory.createCTRegularTextRun();
        //ctRegularTextRun.setRPr(null);
        ctRegularTextRun.setT(getText(p));

        ctRegularTextRun.setRPr(ctTextCharacterProperties);
        CTTextParagraph ctTextParagraph = dmlFactory.createCTTextParagraph();
        ctTextParagraph.setPPr(ctTextParagraphProperties);
        CTTextCharacterProperties ctTextCharacterProperties2 = dmlFactory.createCTTextCharacterProperties();

        //-<a:endParaRPr lang="en-US" dirty="0" sz="800" i="1">
        ctTextCharacterProperties2.setLang(mode);
        ctTextCharacterProperties2.setDirty(Boolean.FALSE);
        ctTextCharacterProperties2.setSz(800);
        ctTextCharacterProperties2.setI(Boolean.TRUE);
        ctTextParagraph.setEndParaRPr(ctTextCharacterProperties2);
        ctTextParagraph.getEGTextRun().add(ctRegularTextRun);

        ctTextBody.getP().add(ctTextParagraph);
        sp.setTxBody(ctTextBody);

        //System.out.println(XmlUtils.marshaltoString(sp));
        return sp;

    }

    private static String getText(P p) {

        String text = "";
        for (Object o : p.getContent()) {

            if (o instanceof R) {
                //return title.getContent().toString();
                R r = (R) o;
                for (int m = 0; m < r.getContent().size(); m++) {
                    Object o2 = r.getContent().get(m);
                    if ((o2 instanceof org.docx4j.wml.Br
                            || o2 instanceof org.docx4j.wml.R.Tab
                            || o2 instanceof org.docx4j.wml.R.LastRenderedPageBreak)) {
                        break;
                    }
                    javax.xml.bind.JAXBElement jaxb = (javax.xml.bind.JAXBElement) (o2);
                    if (jaxb.getDeclaredType().getName().equals("org.docx4j.wml.Text")) {
                        //Text text2= (Text)jaxb;
                        text = text.concat(((Text) (jaxb.getValue())).getValue());
                    }
                }
            }
        }
        return text;
    }

    private static Shape getSpFooter(P footerP, int footercount) {

        org.pptx4j.pml.ObjectFactory pmlObjectFactory = new org.pptx4j.pml.ObjectFactory();

        Shape shape = pmlObjectFactory.createShape();
        // Create object for nvSpPr
        Shape.NvSpPr shapenvsppr = pmlObjectFactory.createShapeNvSpPr();
        shape.setNvSpPr(shapenvsppr);
        org.docx4j.dml.ObjectFactory dmlObjectFactory = new org.docx4j.dml.ObjectFactory();
        // Create object for cNvPr
        CTNonVisualDrawingProps nonvisualdrawingprops = dmlObjectFactory.createCTNonVisualDrawingProps();
        shapenvsppr.setCNvPr(nonvisualdrawingprops);
        nonvisualdrawingprops.setDescr("");
        nonvisualdrawingprops.setName("Rectangle 3");
        nonvisualdrawingprops.setId(4);
        // Create object for cNvSpPr
        CTNonVisualDrawingShapeProps nonvisualdrawingshapeprops = dmlObjectFactory.createCTNonVisualDrawingShapeProps();
        shapenvsppr.setCNvSpPr(nonvisualdrawingshapeprops);
        // Create object for nvPr
        NvPr nvpr = pmlObjectFactory.createNvPr();
        shapenvsppr.setNvPr(nvpr);
        // Create object for spPr
        CTShapeProperties shapeproperties = dmlObjectFactory.createCTShapeProperties();
        shape.setSpPr(shapeproperties);
        // Create object for xfrm
        CTTransform2D transform2d = dmlObjectFactory.createCTTransform2D();
        shapeproperties.setXfrm(transform2d);
        // Create object for ext
        CTPositiveSize2D positivesize2d = dmlObjectFactory.createCTPositiveSize2D();
        transform2d.setExt(positivesize2d);
        positivesize2d.setCx(4530407);
        positivesize2d.setCy(215444);

        transform2d.setRot(0);
        // Create object for off
        CTPoint2D point2d = dmlObjectFactory.createCTPoint2D();
        transform2d.setOff(point2d);
        point2d.setY((long) (6385072 / (1 + (.1 * footercount))));
        point2d.setX((long) (1520706 / (1 + (.1 * footercount))));

        // Create object for prstGeom
        CTPresetGeometry2D presetgeometry2d = dmlObjectFactory.createCTPresetGeometry2D();
        shapeproperties.setPrstGeom(presetgeometry2d);
        // Create object for avLst
        CTGeomGuideList geomguidelist = dmlObjectFactory.createCTGeomGuideList();
        presetgeometry2d.setAvLst(geomguidelist);
        presetgeometry2d.setPrst(org.docx4j.dml.STShapeType.RECT);
        // Create object for txBody
        CTTextBody textbody = dmlObjectFactory.createCTTextBody();
        shape.setTxBody(textbody);
        // Create object for bodyPr
        CTTextBodyProperties textbodyproperties = dmlObjectFactory.createCTTextBodyProperties();
        textbody.setBodyPr(textbodyproperties);
        textbodyproperties.setWrap(org.docx4j.dml.STTextWrappingType.NONE);
        // Create object for spAutoFit
        CTTextShapeAutofit textshapeautofit = dmlObjectFactory.createCTTextShapeAutofit();
        textbodyproperties.setSpAutoFit(textshapeautofit);
        // Create object for lstStyle
        CTTextListStyle textliststyle = dmlObjectFactory.createCTTextListStyle();
        textbody.setLstStyle(textliststyle);
        // Create object for p
        CTTextParagraph textparagraph = dmlObjectFactory.createCTTextParagraph();
        textbody.getP().add(textparagraph);
        // Create object for pPr
        CTTextParagraphProperties textparagraphproperties = dmlObjectFactory.createCTTextParagraphProperties();
        textparagraph.setPPr(textparagraphproperties);
        // Create object for spcBef
        CTTextSpacing textspacing = dmlObjectFactory.createCTTextSpacing();
        textparagraphproperties.setSpcBef(textspacing);
        // Create object for spcPts
        CTTextSpacingPoint textspacingpoint = dmlObjectFactory.createCTTextSpacingPoint();
        textspacing.setSpcPts(textspacingpoint);
        textspacingpoint.setVal(600);
        // Create object for tabLst
        CTTextTabStopList texttabstoplist = dmlObjectFactory.createCTTextTabStopList();
        textparagraphproperties.setTabLst(texttabstoplist);
        // Create object for tab
        CTTextTabStop texttabstop = dmlObjectFactory.createCTTextTabStop();
        texttabstoplist.getTab().add(texttabstop);
        texttabstop.setPos(5669280);
        texttabstop.setAlgn(org.docx4j.dml.STTextTabAlignType.R);
        // Create object for r
        CTRegularTextRun regulartextrun = dmlObjectFactory.createCTRegularTextRun();
        textparagraph.getEGTextRun().add(regulartextrun);
        // Create object for rPr
        CTTextCharacterProperties textcharacterproperties = dmlObjectFactory.createCTTextCharacterProperties();
        regulartextrun.setRPr(textcharacterproperties);
        // Create object for latin
        TextFont textfont = dmlObjectFactory.createTextFont();
        textcharacterproperties.setLatin(textfont);
        textfont.setTypeface("Franklin Gothic Book");
        textfont.setPanose("020B0503020102020204");
        textfont.setPitchFamily(Byte.decode("34"));
        textfont.setCharset(Byte.decode("0"));
        // Create object for ea
        TextFont textfont2 = dmlObjectFactory.createTextFont();
        textcharacterproperties.setEa(textfont2);
        textfont2.setTypeface("Times New Roman");
        textfont2.setPanose("02020603050405020304");
        textfont2.setPitchFamily(Byte.decode("18"));
        textfont2.setCharset(Byte.decode("0"));
        // Create object for cs
        TextFont textfont3 = dmlObjectFactory.createTextFont();
        textcharacterproperties.setCs(textfont3);
        textfont3.setTypeface("Times New Roman");
        textfont3.setPanose("02020603050405020304");
        textfont3.setPitchFamily(Byte.decode("18"));
        textfont3.setCharset(Byte.decode("0"));
        textcharacterproperties.setLang("en-US");
        textcharacterproperties.setSz(800);
        textcharacterproperties.setSmtId(new Long(0));
        if (footerP != null) {
            regulartextrun.setT(footerP.toString());
        }
        // Create object for r
        CTRegularTextRun regulartextrun2 = dmlObjectFactory.createCTRegularTextRun();
        textparagraph.getEGTextRun().add(regulartextrun2);
        // Create object for rPr
        CTTextCharacterProperties textcharacterproperties2 = dmlObjectFactory.createCTTextCharacterProperties();
        regulartextrun2.setRPr(textcharacterproperties2);
        // Create object for latin
        TextFont textfont4 = dmlObjectFactory.createTextFont();
        textcharacterproperties2.setLatin(textfont4);
        textfont4.setTypeface("Franklin Gothic Book");
        textfont4.setPanose("020B0503020102020204");
        textfont4.setPitchFamily(Byte.decode("34"));
        textfont4.setCharset(Byte.decode("0"));
        // Create object for ea
        TextFont textfont5 = dmlObjectFactory.createTextFont();
        textcharacterproperties2.setEa(textfont5);
        textfont5.setTypeface("Times New Roman");
        textfont5.setPanose("02020603050405020304");
        textfont5.setPitchFamily(Byte.decode("18"));
        textfont5.setCharset(Byte.decode("0"));
        // Create object for cs
        TextFont textfont6 = dmlObjectFactory.createTextFont();
        textcharacterproperties2.setCs(textfont6);
        textfont6.setTypeface("Times New Roman");
        textfont6.setPanose("02020603050405020304");
        textfont6.setPitchFamily(Byte.decode("18"));
        textfont6.setCharset(Byte.decode("0"));
        textcharacterproperties2.setLang("en-US");
        textcharacterproperties2.setSz(800);
        textcharacterproperties2.setSmtId(new Long(0));
        regulartextrun2.setT("MarketsandMarkets");
        // Create object for endParaRPr
        CTRegularTextRun regulartextrun3 = dmlObjectFactory.createCTRegularTextRun();
        textparagraph.getEGTextRun().add(regulartextrun3);
        // Create object for rPr
        CTTextCharacterProperties textcharacterproperties3 = dmlObjectFactory.createCTTextCharacterProperties();
        regulartextrun3.setRPr(textcharacterproperties3);
        // Create object for latin
        TextFont textfont7 = dmlObjectFactory.createTextFont();
        textcharacterproperties3.setLatin(textfont7);
        textfont7.setTypeface("Franklin Gothic Book");
        textfont7.setPanose("020B0503020102020204");
        textfont7.setPitchFamily(Byte.decode("34"));
        textfont7.setCharset(Byte.decode("0"));
        // Create object for ea
        TextFont textfont8 = dmlObjectFactory.createTextFont();
        textcharacterproperties3.setEa(textfont8);
        textfont8.setTypeface("Times New Roman");
        textfont8.setPanose("02020603050405020304");
        textfont8.setPitchFamily(Byte.decode("18"));
        textfont8.setCharset(Byte.decode("0"));
        // Create object for cs
        TextFont textfont9 = dmlObjectFactory.createTextFont();
        textcharacterproperties3.setCs(textfont9);
        textfont9.setTypeface("Times New Roman");
        textfont9.setPanose("02020603050405020304");
        textfont9.setPitchFamily(Byte.decode("18"));
        textfont9.setCharset(Byte.decode("0"));
        textcharacterproperties3.setLang("en-US");
        textcharacterproperties3.setSz(800);
        textcharacterproperties3.setSmtId(new Long(0));
        regulartextrun3.setT(" Analysis");
        // Create object for endParaRPr
        CTTextCharacterProperties textcharacterproperties4 = dmlObjectFactory.createCTTextCharacterProperties();
        textparagraph.setEndParaRPr(textcharacterproperties4);
        // Create object for latin
        TextFont textfont10 = dmlObjectFactory.createTextFont();
        textcharacterproperties4.setLatin(textfont10);
        textfont10.setTypeface("Franklin Gothic Book");
        textfont10.setPanose("020B0503020102020204");
        textfont10.setPitchFamily(Byte.decode("34"));
        textfont10.setCharset(Byte.decode("0"));
        // Create object for ea
        TextFont textfont11 = dmlObjectFactory.createTextFont();
        textcharacterproperties4.setEa(textfont11);
        textfont11.setTypeface("Times New Roman");
        textfont11.setPanose("02020603050405020304");
        textfont11.setPitchFamily(Byte.decode("18"));
        textfont11.setCharset(Byte.decode("0"));
        // Create object for cs
        TextFont textfont12 = dmlObjectFactory.createTextFont();
        textcharacterproperties4.setCs(textfont12);
        textfont12.setTypeface("Times New Roman");
        textfont12.setPanose("02020603050405020304");
        textfont12.setPitchFamily(Byte.decode("18"));
        textfont12.setCharset(Byte.decode("0"));
        textcharacterproperties4.setLang("en-US");
        textcharacterproperties4.setSz(800);
        // Create object for effectLst
        CTEffectList effectlist = dmlObjectFactory.createCTEffectList();
        textcharacterproperties4.setEffectLst(effectlist);
        textcharacterproperties4.setSmtId(new Long(0));

        return shape;
    }

    private static Pic getPic(P figure, SlidePart slidePart) throws InvalidFormatException, Exception {
        org.pptx4j.pml.ObjectFactory pmlObjectFactory = new org.pptx4j.pml.ObjectFactory();

        Pic pic = pmlObjectFactory.createPic();
        org.docx4j.dml.ObjectFactory dmlObjectFactory = new org.docx4j.dml.ObjectFactory();
        MetafileEmfPart imagePart = (MetafileEmfPart) documentPart.getRelationshipsPart().getPart(getRelationShip(figure));
        if (imagePart == null) {
            return pic;
        }
        byte[] imageBytes = new byte[imagePart.getBuffer().limit()];
        for (int i = 0; i < imageBytes.length; i++) {
            imageBytes[i] = imagePart.getBuffer().get(i);
        }

        //create image in /media/image<n>.emf
        BinaryPartAbstractImage bpai = BinaryPartAbstractImage.createImagePart(presentationMLPackageOut, presentationMLPackageOut.getMainPresentationPart(), imageBytes, ContentTypes.IMAGE_EMF);
        Relationship rel = slidePart.addTargetPart(bpai);
            //presentationMLPackageOut.addTargetPart(bpai);
        //getRelationshipsPart().addPart(bpai, RelationshipsPart.AddPartBehaviour.REUSE_EXISTING, null);

        // Create object for blipFill
        CTBlipFillProperties blipfillproperties = dmlObjectFactory.createCTBlipFillProperties();
        pic.setBlipFill(blipfillproperties);
        // Create object for blip
        CTBlip blip = dmlObjectFactory.createCTBlip();
        blipfillproperties.setBlip(blip);
        blip.setEmbed(rel.getId());
        // Create object for extLst
        CTOfficeArtExtensionList officeartextensionlist = dmlObjectFactory.createCTOfficeArtExtensionList();
        blip.setExtLst(officeartextensionlist);
        // Create object for ext
        CTOfficeArtExtension officeartextension = dmlObjectFactory.createCTOfficeArtExtension();
        officeartextensionlist.getExt().add(officeartextension);
        officeartextension.setUri("{28A0092B-C50C-407E-A947-70E740481C1C}");
        blip.setCstate(org.docx4j.dml.STBlipCompression.PRINT);
        blip.setLink("");
        // Create object for srcRect
        CTRelativeRect relativerect = dmlObjectFactory.createCTRelativeRect();
        blipfillproperties.setSrcRect(relativerect);
        relativerect.setB(0);
        relativerect.setR(0);
        relativerect.setL(0);
        relativerect.setT(0);
        // Create object for stretch
        CTStretchInfoProperties stretchinfoproperties = dmlObjectFactory.createCTStretchInfoProperties();
        blipfillproperties.setStretch(stretchinfoproperties);
        // Create object for fillRect
        CTRelativeRect relativerect2 = dmlObjectFactory.createCTRelativeRect();
        stretchinfoproperties.setFillRect(relativerect2);
        relativerect2.setB(0);
        relativerect2.setR(0);
        relativerect2.setL(0);
        relativerect2.setT(0);
        // Create object for spPr
        CTShapeProperties shapeproperties = dmlObjectFactory.createCTShapeProperties();
        pic.setSpPr(shapeproperties);
        // Create object for noFill
        CTNoFillProperties nofillproperties = dmlObjectFactory.createCTNoFillProperties();
        shapeproperties.setNoFill(nofillproperties);
        // Create object for xfrm
        CTTransform2D transform2d = dmlObjectFactory.createCTTransform2D();
        shapeproperties.setXfrm(transform2d);
        // Create object for ext
        CTPositiveSize2D positivesize2d = dmlObjectFactory.createCTPositiveSize2D();
        transform2d.setExt(positivesize2d);
//        positivesize2d.setCx(5732145);
//        positivesize2d.setCy(2141855);
        positivesize2d.setCx(6647096);
        positivesize2d.setCy(3580327);
        transform2d.setRot(0);
        // Create object for off
        CTPoint2D point2d = dmlObjectFactory.createCTPoint2D();
        transform2d.setOff(point2d);
//        point2d.setY(2358072);
//        point2d.setX(2086927);
        point2d.setY(1416676);
        point2d.setX(1171977);
        // Create object for ln
        CTLineProperties lineproperties = dmlObjectFactory.createCTLineProperties();
        shapeproperties.setLn(lineproperties);
        // Create object for noFill
        CTNoFillProperties nofillproperties2 = dmlObjectFactory.createCTNoFillProperties();
        lineproperties.setNoFill(nofillproperties2);
        shapeproperties.setBwMode(org.docx4j.dml.STBlackWhiteMode.AUTO);
        // Create object for prstGeom
        CTPresetGeometry2D presetgeometry2d = dmlObjectFactory.createCTPresetGeometry2D();
        shapeproperties.setPrstGeom(presetgeometry2d);
        // Create object for avLst
        CTGeomGuideList geomguidelist = dmlObjectFactory.createCTGeomGuideList();
        presetgeometry2d.setAvLst(geomguidelist);
        presetgeometry2d.setPrst(org.docx4j.dml.STShapeType.RECT);
        // Create object for nvPicPr
        Pic.NvPicPr picnvpicpr = pmlObjectFactory.createPicNvPicPr();
        pic.setNvPicPr(picnvpicpr);
        // Create object for cNvPr
        CTNonVisualDrawingProps nonvisualdrawingprops = dmlObjectFactory.createCTNonVisualDrawingProps();
        picnvpicpr.setCNvPr(nonvisualdrawingprops);
        nonvisualdrawingprops.setDescr("");
        nonvisualdrawingprops.setName("Picture " + figcount);
        nonvisualdrawingprops.setId(figcount);
        // Create object for cNvPicPr
        CTNonVisualPictureProperties nonvisualpictureproperties = dmlObjectFactory.createCTNonVisualPictureProperties();
        picnvpicpr.setCNvPicPr(nonvisualpictureproperties);
        // Create object for nvPr
        NvPr nvpr = pmlObjectFactory.createNvPr();
        picnvpicpr.setNvPr(nvpr);

        return pic;
    }

    private static Relationship getRelationShip(P tcpara) {

        Relationship rel = null;
        for (int l = 0; l < tcpara.getContent().size(); l++) {
            //Dont assume its always a row
            if (!(tcpara.getContent().get(l) instanceof org.docx4j.wml.R)) {
                continue;
            }
            org.docx4j.wml.R r = (org.docx4j.wml.R) (tcpara.getContent().get(l));

            for (int m = 0; m < r.getContent().size(); m++) {
                Object o = r.getContent().get(m);
                if ((o instanceof org.docx4j.wml.Br
                        || o instanceof org.docx4j.wml.R.Tab
                        || o instanceof org.docx4j.wml.R.LastRenderedPageBreak)) {
                    break;
                }
                javax.xml.bind.JAXBElement jaxb = (javax.xml.bind.JAXBElement) (o);
                switch (jaxb.getDeclaredType().getName()) {
                    case "org.docx4j.wml.Drawing":
                        org.docx4j.wml.Drawing drawing = (org.docx4j.wml.Drawing) (jaxb.getValue());
                        org.docx4j.dml.wordprocessingDrawing.Inline inline = null;
                        org.docx4j.dml.wordprocessingDrawing.Anchor anchor = null;
                        try {
                            inline = (org.docx4j.dml.wordprocessingDrawing.Inline) (drawing).getAnchorOrInline().get(0);
                        } catch (java.lang.ClassCastException cce) {
                            anchor = (org.docx4j.dml.wordprocessingDrawing.Anchor) (drawing).getAnchorOrInline().get(0);
                        }
                        if (inline != null && inline.getGraphic() != null) {
                            //log.debug("found a:graphic");
                            org.docx4j.dml.Graphic graphic = inline.getGraphic();
                            if (graphic.getGraphicData() != null) {
                                String relId = null;
                                if (graphic.getGraphicData().getPic() != null) {
                                    relId = graphic.getGraphicData().getPic().getBlipFill().getBlip().getEmbed();
                                } else {
                                    //this is to handle charts
                                    List<Object> anyObjs = graphic.getGraphicData().getAny();
                                    for (Object any : anyObjs) {

                                        if (any instanceof javax.xml.bind.JAXBElement) {
                                            javax.xml.bind.JAXBElement jaxbe = (javax.xml.bind.JAXBElement) any;
                                            if (jaxbe.getDeclaredType().getName().equals("org.docx4j.dml.chart.CTRelId")) {

                                                relId = ((org.docx4j.dml.chart.CTRelId) jaxbe.getValue()).getId();
                                            }

                                        }
                                    }
                                }
                                rel = wordMLPackageIn.getMainDocumentPart().getRelationshipsPart().getRelationshipByID(relId);
                                //System.out.println("Row " + (i + 1) + " column" + (j + 1) + "'s " + (k + 1) + " value is image " + imageId + " mapped to file " + relationsPart.getRelationshipByID(imageId).getTarget());
                            }
                        } else if (anchor != null && anchor.getGraphic() != null) {
                            org.docx4j.dml.Graphic graphic = anchor.getGraphic();
                            if (graphic.getGraphicData() != null) {
                                String relId = null;
                                if (graphic.getGraphicData().getPic() != null) {
                                    relId = graphic.getGraphicData().getPic().getBlipFill().getBlip().getEmbed();
                                } else {
                                    //this is to handle charts
                                    List<Object> anyObjs = graphic.getGraphicData().getAny();
                                    for (Object any : anyObjs) {

                                        if (any instanceof javax.xml.bind.JAXBElement) {
                                            javax.xml.bind.JAXBElement jaxbe = (javax.xml.bind.JAXBElement) any;
                                            if (jaxbe.getDeclaredType().getName().equals("org.docx4j.dml.chart.CTRelId")) {

                                                relId = ((org.docx4j.dml.chart.CTRelId) jaxbe.getValue()).getId();
                                            }

                                        }
                                    }
                                }
                                rel = wordMLPackageIn.getMainDocumentPart().getRelationshipsPart().getRelationshipByID(relId);
                            }
                        }
                        break;
                    // also check if images or drwings are the

                    case "org.docx4j.wml.Pict":
                        System.out.println("found a w:Pict instead of a w:drawing");
                        //org.docx4j.wml.Pict pict = (org.docx4j.wml.Pict) (jaxb.getValue());
                        break;
                }
            }
        }
        return rel;
    }

    private static void filterFigures(List<PFigurePair> pFigurePairs) {
        pFigurePairs.stream().forEach((pfp) -> {
            for (CTBookmark ctb : pfp.ctblist) {
                if (figureElements.containsKey(ctb.getName())) {
                    pfp.setIndex(((Element) figureElements.get(ctb.getName())).getindex());
                    sequenceExportList.put(sequenceList.get((Element) figureElements.get(ctb.getName())), pfp);
                }
            }
        });
    }

    private static void createTocElements() {

        //air and missile 
            //String elems = "[\"T\",\"_Toc450152354\",\"1\"],[\"T\",\"_Toc450152355\",\"2\"],[\"T\",\"_Toc450152356\",\"3\"],[\"T\",\"_Toc450152357\",\"4\"],[\"T\",\"_Toc450152358\",\"5\"],[\"T\",\"_Toc450152359\",\"6\"],[\"T\",\"_Toc450152360\",\"7\"],[\"T\",\"_Toc450152361\",\"8\"],[\"T\",\"_Toc450152362\",\"9\"],[\"T\",\"_Toc450152363\",\"10\"],[\"T\",\"_Toc450152364\",\"11\"],[\"T\",\"_Toc450152365\",\"12\"],[\"T\",\"_Toc450152366\",\"13\"],[\"T\",\"_Toc450152367\",\"14\"],[\"T\",\"_Toc450152368\",\"15\"],[\"T\",\"_Toc450152369\",\"16\"],[\"T\",\"_Toc450152370\",\"17\"],[\"T\",\"_Toc450152371\",\"18\"],[\"T\",\"_Toc450152372\",\"19\"],[\"T\",\"_Toc450152373\",\"20\"],[\"T\",\"_Toc450152374\",\"21\"],[\"T\",\"_Toc450152375\",\"22\"],[\"T\",\"_Toc450152376\",\"23\"],[\"T\",\"_Toc450152377\",\"24\"],[\"T\",\"_Toc450152378\",\"25\"],[\"T\",\"_Toc450152379\",\"26\"],[\"T\",\"_Toc450152380\",\"27\"],[\"T\",\"_Toc450152381\",\"28\"],[\"T\",\"_Toc450152382\",\"29\"],[\"T\",\"_Toc450152383\",\"30\"],[\"T\",\"_Toc450152384\",\"31\"],[\"T\",\"_Toc450152385\",\"32\"],[\"T\",\"_Toc450152386\",\"33\"],[\"T\",\"_Toc450152387\",\"34\"],[\"T\",\"_Toc450152388\",\"35\"],[\"T\",\"_Toc450152389\",\"36\"],[\"T\",\"_Toc450152390\",\"37\"],[\"T\",\"_Toc450152391\",\"38\"],[\"T\",\"_Toc450152392\",\"39\"],[\"T\",\"_Toc450152393\",\"40\"],[\"T\",\"_Toc450152394\",\"41\"],[\"T\",\"_Toc450152395\",\"42\"],[\"T\",\"_Toc450152396\",\"43\"],[\"T\",\"_Toc450152397\",\"44\"],[\"T\",\"_Toc450152398\",\"45\"],[\"T\",\"_Toc450152399\",\"46\"],[\"T\",\"_Toc450152400\",\"47\"],[\"T\",\"_Toc450152401\",\"48\"],[\"T\",\"_Toc450152402\",\"49\"],[\"T\",\"_Toc450152403\",\"50\"],[\"T\",\"_Toc450152404\",\"51\"],[\"T\",\"_Toc450152405\",\"52\"],[\"T\",\"_Toc450152406\",\"53\"],[\"T\",\"_Toc450152407\",\"54\"],[\"T\",\"_Toc450152408\",\"55\"],[\"T\",\"_Toc450152409\",\"56\"],[\"T\",\"_Toc450152410\",\"57\"],[\"T\",\"_Toc450152411\",\"58\"],[\"T\",\"_Toc450152412\",\"59\"],[\"T\",\"_Toc450152413\",\"60\"],[\"T\",\"_Toc450152414\",\"61\"],[\"T\",\"_Toc450152415\",\"62\"],[\"T\",\"_Toc450152416\",\"63\"],[\"T\",\"_Toc450152417\",\"64\"],[\"T\",\"_Toc450152418\",\"65\"],[\"T\",\"_Toc450152419\",\"66\"],[\"T\",\"_Toc450152420\",\"67\"],[\"T\",\"_Toc450152421\",\"68\"],[\"T\",\"_Toc450152422\",\"69\"],[\"T\",\"_Toc450152423\",\"70\"],[\"T\",\"_Toc450152424\",\"71\"],[\"T\",\"_Toc450152425\",\"72\"],[\"T\",\"_Toc450152426\",\"73\"],[\"T\",\"_Toc450152427\",\"74\"],[\"T\",\"_Toc450152428\",\"75\"],[\"T\",\"_Toc450152429\",\"76\"],[\"T\",\"_Toc450152430\",\"77\"],[\"T\",\"_Toc450152431\",\"78\"],[\"T\",\"_Toc450152432\",\"79\"],[\"T\",\"_Toc450152433\",\"80\"],[\"T\",\"_Toc450152434\",\"81\"],[\"T\",\"_Toc450152435\",\"82\"],[\"T\",\"_Toc450152436\",\"83\"],[\"T\",\"_Toc450152437\",\"84\"],[\"T\",\"_Toc450152438\",\"85\"],[\"T\",\"_Toc450152439\",\"86\"],[\"T\",\"_Toc450152440\",\"87\"],[\"T\",\"_Toc450152441\",\"88\"],[\"T\",\"_Toc450152442\",\"89\"]";
    
        String elems = "[\"F\",\"_Toc450152443\",\"1\"],[\"F\",\"_Toc450152444\",\"2\"],[\"F\",\"_Toc450152445\",\"3\"],[\"F\",\"_Toc450152446\",\"4\"],[\"F\",\"_Toc450152447\",\"5\"],[\"F\",\"_Toc450152448\",\"6\"],[\"F\",\"_Toc450152449\",\"7\"],[\"F\",\"_Toc450152450\",\"8\"],[\"F\",\"_Toc450152451\",\"9\"],[\"F\",\"_Toc450152452\",\"10\"],[\"F\",\"_Toc450152453\",\"11\"],[\"F\",\"_Toc450152454\",\"12\"],[\"F\",\"_Toc450152455\",\"13\"],[\"F\",\"_Toc450152456\",\"14\"],[\"F\",\"_Toc450152457\",\"15\"],[\"F\",\"_Toc450152458\",\"16\"],[\"F\",\"_Toc450152459\",\"17\"],[\"F\",\"_Toc450152460\",\"18\"],[\"F\",\"_Toc450152461\",\"19\"],[\"F\",\"_Toc450152462\",\"20\"],[\"F\",\"_Toc450152463\",\"21\"],[\"F\",\"_Toc450152464\",\"22\"],[\"F\",\"_Toc450152465\",\"23\"],[\"F\",\"_Toc450152466\",\"24\"],[\"F\",\"_Toc450152467\",\"25\"],[\"F\",\"_Toc450152468\",\"26\"],[\"F\",\"_Toc450152469\",\"27\"],[\"F\",\"_Toc450152470\",\"28\"],[\"F\",\"_Toc450152471\",\"29\"],[\"F\",\"_Toc450152472\",\"30\"],[\"F\",\"_Toc450152473\",\"31\"],[\"F\",\"_Toc450152474\",\"32\"],[\"F\",\"_Toc450152475\",\"33\"],[\"F\",\"_Toc450152476\",\"34\"],[\"F\",\"_Toc450152477\",\"35\"],[\"F\",\"_Toc450152478\",\"36\"],[\"F\",\"_Toc450152479\",\"37\"],[\"F\",\"_Toc450152480\",\"38\"],[\"F\",\"_Toc450152481\",\"39\"],[\"F\",\"_Toc450152482\",\"40\"],[\"F\",\"_Toc450152483\",\"41\"],[\"F\",\"_Toc450152484\",\"42\"],[\"F\",\"_Toc450152485\",\"43\"],[\"F\",\"_Toc450152486\",\"44\"],[\"F\",\"_Toc450152487\",\"45\"],[\"F\",\"_Toc450152488\",\"46\"],[\"F\",\"_Toc450152489\",\"47\"],[\"F\",\"_Toc450152490\",\"48\"],[\"F\",\"_Toc450152491\",\"49\"],[\"F\",\"_Toc450152492\",\"50\"],[\"F\",\"_Toc450152493\",\"51\"],[\"F\",\"_Toc450152494\",\"52\"],[\"F\",\"_Toc450152495\",\"53\"],[\"F\",\"_Toc450152496\",\"54\"],[\"F\",\"_Toc450152497\",\"55\"],[\"F\",\"_Toc450152498\",\"56\"],[\"F\",\"_Toc450152499\",\"57\"],[\"F\",\"_Toc450152500\",\"58\"],[\"F\",\"_Toc450152501\",\"59\"],[\"F\",\"_Toc450152502\",\"60\"],[\"F\",\"_Toc450152503\",\"61\"],[\"F\",\"_Toc450152504\",\"62\"],[\"F\",\"_Toc450152505\",\"63\"],[\"F\",\"_Toc450152506\",\"64\"],[\"F\",\"_Toc450152507\",\"65\"],[\"F\",\"_Toc450152508\",\"66\"],[\"F\",\"_Toc450152509\",\"67\"],[\"F\",\"_Toc450152510\",\"68\"],[\"F\",\"_Toc450152511\",\"69\"],[\"F\",\"_Toc450152512\",\"70\"],[\"F\",\"_Toc450152513\",\"71\"],[\"F\",\"_Toc450152514\",\"72\"],[\"F\",\"_Toc450152515\",\"73\"],[\"F\",\"_Toc450152516\",\"74\"],[\"F\",\"_Toc450152517\",\"75\"]";
        //organic electronics
        //String elems = "[\"F\",\"_Toc401326434\",\"1\"],[\"F\",\"_Toc401326435\",\"2\"],[\"F\",\"_Toc401326436\",\"3\"],[\"F\",\"_Toc401326437\",\"4\"],[\"F\",\"_Toc401326438\",\"5\"],[\"F\",\"_Toc401326439\",\"6\"],[\"F\",\"_Toc401326440\",\"7\"],[\"F\",\"_Toc401326441\",\"8\"],[\"F\",\"_Toc401326442\",\"9\"],[\"F\",\"_Toc401326443\",\"10\"],[\"F\",\"_Toc401326444\",\"11\"],[\"F\",\"_Toc401326445\",\"12\"],[\"F\",\"_Toc401326446\",\"13\"],[\"F\",\"_Toc401326447\",\"14\"],[\"F\",\"_Toc401326448\",\"15\"],[\"F\",\"_Toc401326449\",\"16\"],[\"F\",\"_Toc401326450\",\"17\"],[\"F\",\"_Toc401326451\",\"18\"],[\"F\",\"_Toc401326452\",\"19\"],[\"F\",\"_Toc401326453\",\"20\"],[\"F\",\"_Toc401326454\",\"21\"],[\"F\",\"_Toc401326455\",\"22\"],[\"F\",\"_Toc401326456\",\"23\"],[\"F\",\"_Toc401326457\",\"24\"],[\"F\",\"_Toc401326458\",\"25\"],[\"F\",\"_Toc401326459\",\"26\"],[\"F\",\"_Toc401326460\",\"27\"],[\"F\",\"_Toc401326461\",\"28\"],[\"F\",\"_Toc401326462\",\"29\"],[\"F\",\"_Toc401326463\",\"30\"],[\"F\",\"_Toc401326464\",\"31\"],[\"F\",\"_Toc401326465\",\"32\"],[\"F\",\"_Toc401326466\",\"33\"],[\"F\",\"_Toc401326467\",\"34\"],[\"F\",\"_Toc401326468\",\"35\"],[\"F\",\"_Toc401326469\",\"36\"],[\"F\",\"_Toc401326470\",\"37\"],[\"F\",\"_Toc401326471\",\"38\"],[\"F\",\"_Toc401326472\",\"39\"],[\"F\",\"_Toc401326473\",\"40\"],[\"F\",\"_Toc401326474\",\"41\"],[\"F\",\"_Toc401326475\",\"42\"],[\"F\",\"_Toc401326476\",\"43\"],[\"F\",\"_Toc401326477\",\"44\"],[\"F\",\"_Toc401326478\",\"45\"],[\"F\",\"_Toc401326479\",\"46\"],[\"F\",\"_Toc401326480\",\"47\"],[\"F\",\"_Toc401326481\",\"48\"],[\"F\",\"_Toc401326482\",\"49\"],[\"F\",\"_Toc401326483\",\"50\"],[\"F\",\"_Toc401326484\",\"51\"],[\"F\",\"_Toc401326485\",\"52\"],[\"F\",\"_Toc401326486\",\"53\"],[\"F\",\"_Toc401326487\",\"54\"],[\"F\",\"_Toc401326488\",\"55\"],[\"F\",\"_Toc401326489\",\"56\"],[\"F\",\"_Toc401326490\",\"57\"],[\"F\",\"_Toc401326491\",\"58\"],[\"F\",\"_Toc401326492\",\"59\"],[\"F\",\"_Toc401326493\",\"60\"],[\"F\",\"_Toc401326494\",\"61\"],[\"F\",\"_Toc401326495\",\"62\"],[\"F\",\"_Toc401326496\",\"63\"],[\"F\",\"_Toc401326497\",\"64\"],[\"F\",\"_Toc401326498\",\"65\"],[\"F\",\"_Toc401326499\",\"66\"],[\"F\",\"_Toc401326500\",\"67\"],[\"F\",\"_Toc401326501\",\"68\"],[\"F\",\"_Toc401326502\",\"69\"],[\"F\",\"_Toc401326503\",\"70\"],[\"F\",\"_Toc401326504\",\"71\"],[\"F\",\"_Toc401326505\",\"72\"],[\"F\",\"_Toc401326506\",\"73\"],[\"F\",\"_Toc401326507\",\"74\"],[\"F\",\"_Toc401326508\",\"75\"],[\"F\",\"_Toc401326509\",\"76\"],[\"F\",\"_Toc401326510\",\"77\"],[\"F\",\"_Toc401326511\",\"78\"],[\"F\",\"_Toc401326512\",\"79\"],[\"F\",\"_Toc401326513\",\"80\"],[\"F\",\"_Toc401326514\",\"81\"],[\"F\",\"_Toc401326515\",\"82\"],[\"F\",\"_Toc401326516\",\"83\"],[\"F\",\"_Toc401326517\",\"84\"],[\"F\",\"_Toc401326518\",\"85\"],[\"F\",\"_Toc401326519\",\"86\"],[\"F\",\"_Toc401326520\",\"87\"],[\"F\",\"_Toc401326521\",\"88\"]";
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



//["F","_Toc401326434","1"],["F","_Toc401326435","2"],["F","_Toc401326436","3"],["F","_Toc401326437","4"],["F","_Toc401326438","5"],["F","_Toc401326439","6"],["F","_Toc401326440","7"],["F","_Toc401326441","8"],["F","_Toc401326442","9"],["F","_Toc401326443","10"],["F","_Toc401326444","11"],["F","_Toc401326445","12"],["F","_Toc401326446","13"],["F","_Toc401326447","14"],["F","_Toc401326448","15"],["F","_Toc401326449","16"],["F","_Toc401326450","17"],["F","_Toc401326451","18"],["F","_Toc401326452","19"],["F","_Toc401326453","20"],["F","_Toc401326454","21"],["F","_Toc401326455","22"],["F","_Toc401326456","23"],["F","_Toc401326457","24"],["F","_Toc401326458","25"],["F","_Toc401326459","26"],["F","_Toc401326460","27"],["F","_Toc401326461","28"],["F","_Toc401326462","29"],["F","_Toc401326463","30"],["F","_Toc401326464","31"],["F","_Toc401326465","32"],["F","_Toc401326466","33"],["F","_Toc401326467","34"],["F","_Toc401326468","35"],["F","_Toc401326469","36"],["F","_Toc401326470","37"],["F","_Toc401326471","38"],["F","_Toc401326472","39"],["F","_Toc401326473","40"],["F","_Toc401326474","41"],["F","_Toc401326475","42"],["F","_Toc401326476","43"],["F","_Toc401326477","44"],["F","_Toc401326478","45"],["F","_Toc401326479","46"],["F","_Toc401326480","47"],["F","_Toc401326481","48"],["F","_Toc401326482","49"],["F","_Toc401326483","50"],["F","_Toc401326484","51"],["F","_Toc401326485","52"],["F","_Toc401326486","53"],["F","_Toc401326487","54"],["F","_Toc401326488","55"],["F","_Toc401326489","56"],["F","_Toc401326490","57"],["F","_Toc401326491","58"],["F","_Toc401326492","59"],["F","_Toc401326493","60"],["F","_Toc401326494","61"],["F","_Toc401326495","62"],["F","_Toc401326496","63"],["F","_Toc401326497","64"],["F","_Toc401326498","65"],["F","_Toc401326499","66"],["F","_Toc401326500","67"],["F","_Toc401326501","68"],["F","_Toc401326502","69"],["F","_Toc401326503","70"],["F","_Toc401326504","71"],["F","_Toc401326505","72"],["F","_Toc401326506","73"],["F","_Toc401326507","74"],["F","_Toc401326508","75"],["F","_Toc401326509","76"],["F","_Toc401326510","77"],["F","_Toc401326511","78"],["F","_Toc401326512","79"],["F","_Toc401326513","80"],["F","_Toc401326514","81"],["F","_Toc401326515","82"],["F","_Toc401326516","83"],["F","_Toc401326517","84"],["F","_Toc401326518","85"],["F","_Toc401326519","86"],["F","_Toc401326520","87"],["F","_Toc401326521","88"]]
