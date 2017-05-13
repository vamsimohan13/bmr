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

import java.io.File;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Objects;
import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import org.docx4j.TraversalUtil;
import org.docx4j.dml.CTBlip;
import org.docx4j.dml.CTBlipFillProperties;
import org.docx4j.dml.CTGeomGuideList;
import org.docx4j.dml.CTLineProperties;
import org.docx4j.dml.CTNoFillProperties;
import org.docx4j.dml.CTNonVisualDrawingProps;
import org.docx4j.dml.CTNonVisualPictureProperties;
import org.docx4j.dml.CTOfficeArtExtension;
import org.docx4j.dml.CTOfficeArtExtensionList;
import org.docx4j.dml.CTPoint2D;
import org.docx4j.dml.CTPositiveSize2D;
import org.docx4j.dml.CTPresetGeometry2D;
import org.docx4j.dml.CTRelativeRect;
import org.docx4j.dml.CTShapeProperties;
import org.docx4j.dml.CTStretchInfoProperties;
import org.docx4j.dml.CTTransform2D;
import org.docx4j.dml.Theme;
import org.docx4j.dml.spreadsheetdrawing.CTAnchorClientData;
import org.docx4j.dml.spreadsheetdrawing.CTMarker;
import org.docx4j.dml.spreadsheetdrawing.CTPicture;
import org.docx4j.dml.spreadsheetdrawing.CTTwoCellAnchor;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import static org.docx4j.openpackaging.packages.SpreadsheetMLPackage.createPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.Styles;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorkbookPart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.samples.AbstractSample;
import org.docx4j.vml.CTShape;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.CTTrPrBase;
import org.docx4j.wml.Hdr;
import org.docx4j.wml.P;
import org.docx4j.wml.Pict;
import org.docx4j.wml.R;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;

import org.xlsx4j.sml.*;

/**
 *
 * @author vamsi.mohan
 */
public class DocxToXcl extends AbstractSample {

    public static JAXBContext context = org.docx4j.jaxb.Context.jc;
    static int tblCount = 1;
    static SpreadsheetMLPackage pkg;
    static MainDocumentPart mdpOut;
    static org.xlsx4j.sml.ObjectFactory smlObjectFactory = new org.xlsx4j.sml.ObjectFactory();
    private static final org.docx4j.wml.ObjectFactory wmlObjectFactory = Context.getWmlObjectFactory();
    static Styles styles;
    static Theme theme;
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

            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Masked - Cardiovascular Information System Market ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Å“ Forecasts to 2020.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Organic Electronics Market - Global Analysis and Forecast 2020.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Air and Missile.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Agriculture Enzymes.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Torque Sensor.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Casino Management Systems (CMS) Market.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Data Center Networking.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Mobile 3D Market.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/ANTIFOAMING AGENT MARKET ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Å“ GLOBAL FORECAST TO 2021.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Power Quality Meter Market - Global Forecast & Trends To 2021.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/MnM.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Feed Acidifiers.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Rolling Stock.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Fire Resistant Glass.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Temperature Management Market.docx";
            inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Human Insulin Market.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/BFSI Security Market.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/Immunotherapy Drugs Market - Copy.docx";
            //inputfilepath = System.getProperty("user.dir") + "/sample-docs/word/1494597741.docx";

        }
        try {
            getOutputFilePath(args);
        } catch (IllegalArgumentException e) {

        }
        try {
            getElements(args);
        } catch (IllegalArgumentException e) {


            /*parse mode with some test data starts*/
//            mode = "parse";
//            reportId = "test";

            /*parse mode with some test data ends*/
//            /*export mode with some test data starts*/
            mode = "export";
            //mode ="all";

            createTocElements();
        }
        //boolean exportAllTables = false;
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

        WordprocessingMLPackage wordMLPackageIn = WordprocessingMLPackage
                .load(new java.io.File(inputfilepath));

        final MainDocumentPart documentPart = wordMLPackageIn.getMainDocumentPart();
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
        System.out.println(header);
        System.out.println("Too Good so far!!!!!!!!!1");

        org.docx4j.wml.Document wmlDocumentEl = (org.docx4j.wml.Document) documentPart
                .getJaxbElement();

        Body body = wmlDocumentEl.getBody();

        //String outputfilepath = System.getProperty("user.dir") + "/output/OUT_Table.xlsx";
        pkg = createPackage();
        //Create styles
        styles = new Styles(new PartName("/xl/styles.xml"));
        WorkbookPart wb = pkg.getWorkbookPart();
        wb.addTargetPart(styles);
        wb.setPartShortcut(styles);
        styles.setJaxbElement(createStyles());

        //createTheme();
        //pkg.addTargetPart(theme);
        //Create Themes
        //ThemePart themePart = new ThemePart(new PartName("/xl/styles.xml"));
        //Create SST to store shared strings
        //sharedStrings = new SharedStrings();
        //xlsx4j.sml.ObjectFactory smlObjectFactory = new org.xlsx4j.sml.ObjectFactory();
        //sst.setUniqueCount(new Long(16));
        //smlObjectFactory = new org.xlsx4j.sml.ObjectFactory();
        // this block of code is for export json for tables/figures
//        BMRUtility BMRUtilityExporter = new BMRUtility();
//        BMRUtilityExporter.setWordMLPkg(wordMLPackageIn);
//        new TraversalUtil(body, BMRUtilityExporter);
        // end block of specific shit
        if (!tableElements.isEmpty()) {
            TableExporterNew tableExporter = new TableExporterNew();
            tableExporter.setWorldMPkg(wordMLPackageIn);
            new TraversalUtil(body, tableExporter);
            filterTables(tableExporter.getPTablePairs());

        }
        if (!figureElements.isEmpty()) {
            FigureExporter figureExporter = new FigureExporter();
            new TraversalUtil(body, figureExporter);
            filterFigures(figureExporter.getPFigurePairs());
        }

        if (!sequenceExportList.isEmpty()) {
            //lets create disclaimer worksheet here as there seems to be tables to export
            //createLicenseLogo();
            wb.addTargetPart(createLicenseWorkSheetPart());
            for (int i = 0; i < sequenceExportList.size(); i++) {
                //count++;
                if (sequenceExportList.containsKey(i)) {
                    if (sequenceExportList.get(i).getClass().getName().equals("mnm.buildmyreport.PTablePair")) {
                        PTablePair ptp = (PTablePair) sequenceExportList.get(i);
                        //System.out.println(ptp.title.toString());

                        try {

                            wb.addTargetPart(createWorksheetPart(ptp));
                        } catch (java.lang.IndexOutOfBoundsException iex) {
                            //i = i;
                            //P createUnnumberedP() {
// remove the added worksheet

                            P p = wmlObjectFactory.createP();
                            R r = wmlObjectFactory.createR();
                            p.getContent().add(r);
                            // Create object for t (wrapped in JAXBElement) 
                            Text text = wmlObjectFactory.createText();
                            JAXBElement<org.docx4j.wml.Text> textWrapped = wmlObjectFactory.createRT(text);
                            r.getContent().add(textWrapped);
                            text.setValue(iex.getLocalizedMessage());
                            PTablePair excPtp = new PTablePair(p, null, null, null);

                            //return p;
                            //}	
                            excPtp.setIndex(Integer.toString(i));
                            wb.addTargetPart(createWorksheetPart(excPtp));
                        }
                    }

                    if (sequenceExportList.get(i).getClass().getName().equals("mnm.buildmyreport.PFigurePair")) {
//                    PFigurePair pfp = (PFigurePair) sequenceExportList.get(i);
//                    figcount++;
//                    SlidePart slidePart = presentationMLPackageOut.createSlidePart(pp, layoutPart,
//                            new PartName("/ppt/slides/slide" + count + ".xml"));
//                    // Lets add title
//                    slidePart.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add(getSpTitle(pfp.title));
//                    // Lets add table
//                    slidePart.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add(getPic(pfp.figure, slidePart));
//                    //Lets add footnote
//                    slidePart.getJaxbElement().getCSld().getSpTree().getSpOrGrpSpOrGraphicFrame().add(getSpFooter(pfp.footer));

                    }
                }

            }
        }

        SaveToZipFile saver = new SaveToZipFile(pkg);

        System.out.println(inputfilepath);
        CharSequence outputfoldername = inputfilepath.subSequence(inputfilepath.lastIndexOf('/') + 1, inputfilepath.lastIndexOf('.'));
        String outputfilepath = System.getProperty("user.dir") + "/output/" + outputfoldername;
        String outputfilename = outfilename + ".xlsx";
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
        saver.save(outputfilepath + "/" + outputfilename);

        System.out.println("Time to export to excel, we have parsed a total of" + sequenceExportList.size() + " tables and(or) figures");
        System.out.println("\n\n done .. saved " + outputfilepath);

    }

    public static WorksheetPart createWorksheetPart(PTablePair pTablePair) throws JAXBException, InvalidFormatException, IndexOutOfBoundsException {

        // Excel Creation
        PartName partName = new PartName("/xl/worksheets/sheet" + tblCount + ".xml");
        String sheetName = "TABLE" + tblCount;
        //long sheetId;

        //Sheets sheets = wb.getJaxbElement().getSheets();
        Worksheet worksheet = smlObjectFactory.createWorksheet();

        // Create object for pageMargins
        CTPageMargins pagemargins = smlObjectFactory.createCTPageMargins();
        worksheet.setPageMargins(pagemargins);
        CTPageSetup pagesetup = smlObjectFactory.createCTPageSetup();
        worksheet.setPageSetup(pagesetup);
        pagesetup.setErrors(org.xlsx4j.sml.STPrintError.DISPLAYED);
        pagesetup.setOrientation(org.xlsx4j.sml.STOrientation.PORTRAIT);
        pagesetup.setPaperSize(new Long(1));
        pagesetup.setFirstPageNumber(new Long(1));
        pagesetup.setHorizontalDpi(new Long(1200));
        pagesetup.setVerticalDpi(new Long(1200));
        pagesetup.setCopies(new Long(1));
        pagesetup.setScale(new Long(100));
        pagesetup.setFitToWidth(new Long(1));
        pagesetup.setFitToHeight(new Long(1));
        pagesetup.setPageOrder(org.xlsx4j.sml.STPageOrder.DOWN_THEN_OVER);
        pagesetup.setCellComments(org.xlsx4j.sml.STCellComments.NONE);

        SheetData sd = smlObjectFactory.createSheetData();
        CTSheetFormatPr ctSheetFormatPr = smlObjectFactory.createCTSheetFormatPr();
        ctSheetFormatPr.setDefaultRowHeight(15);
        ctSheetFormatPr.setBaseColWidth(new Long(8));
        ctSheetFormatPr.setOutlineLevelRow(Short.decode("0"));
        ctSheetFormatPr.setOutlineLevelCol(Short.decode("0"));

        SheetViews sheetViews = smlObjectFactory.createSheetViews();
        SheetView sheetView = smlObjectFactory.createSheetView();
        sheetView.setView(org.xlsx4j.sml.STSheetViewType.NORMAL);
        sheetViews.getSheetView().add(sheetView);
        worksheet.setSheetViews(sheetViews);

        List rows = sd.getRow();
        List<org.docx4j.wml.Tr> docxTblRows = new ArrayList<>();
        //lets fill first row with title data
        Row titlerow = smlObjectFactory.createRow();
        titlerow.setR(new Long(1));
        titlerow.setHt(17.25);
        titlerow.setOutlineLevel(Short.decode("0"));
        titlerow.setS(new Long(4));
        Cell titleCell = smlObjectFactory.createCell();

        titleCell.setT(STCellType.INLINE_STR);

        CTXstringWhitespace ctx = smlObjectFactory.createCTXstringWhitespace();
        ctx.setValue("TABLE " + pTablePair.getIndex() + " " + pTablePair.title.toString().toUpperCase());

        CTRst ctrst = new CTRst();
        ctrst.setT(ctx);
        titleCell.setIs(ctrst);
        titleCell.setS(new Long(3));

        titlerow.getC().add(titleCell);
        rows.add(titlerow);
        //add a blank row now
//        
//        Row titleafterrow = smlObjectFactory.createRow();
//        Cell blanCell = smlObjectFactory.createCell();
//
//        blanCell.setT(STCellType.INLINE_STR);
//
//        CTXstringWhitespace blancctx = smlObjectFactory.createCTXstringWhitespace();
//        blancctx.setValue("");
//
//        CTRst blancctrst = new CTRst();
//        blancctrst.setT(blancctx);
//        blanCell.setIs(blancctrst);
//        //blanCell.setR("A2");
//        //blanCell.setS(new Long(2));
//        titleafterrow.setR(new Long(2));
//        titleafterrow.getC().add(blanCell);
//        
//        rows.add(titleafterrow);
        //finished add blank row
        //Lets initialize sheetData for the sheet so that it can hold docx Table data
        if (pTablePair.tbl != null) {
            for (int i = 0; i < pTablePair.tbl.getContent().size(); i++) {
                if ((pTablePair.tbl.getContent().get(i)) instanceof org.docx4j.wml.Tr) {
                    Row row = smlObjectFactory.createRow();
                    row.setR(new Long(i + 3));
                    if (i == 0) {
                        row.setThickBot(Boolean.TRUE);
                    }
                    row.setHt(17.25);
                    row.setOutlineLevel(Short.decode("0"));
                    row.setS(new Long(0));
                    //row.getSpans().add("1:" + colSize);
                    docxTblRows.add((org.docx4j.wml.Tr) pTablePair.tbl.getContent().get(i));
                    rows.add(row);
                }

            }
            int colSize = pTablePair.tbl.getTblGrid().getGridCol().size();
            //lets fill sheetdata with docx table data
            for (int i = 0; i < docxTblRows.size(); i++) {
                int incr;
                //to ignore anything that comes outside of column grids..
                int j = getGridAfter(docxTblRows.get(i));
                for (int k = 0; colSize - j > 0; k++) {
                    JAXBElement jaxbTr = (JAXBElement) (docxTblRows.get(i).getContent().get(k));
                    if (jaxbTr.getDeclaredType().getName().equals("org.docx4j.wml.Tc")) {

                        Tc tc = (Tc) jaxbTr.getValue();

                        incr = (tc.getTcPr().getGridSpan() != null) ? tc.getTcPr().getGridSpan().getVal().intValue() : 1;
                        j = j + incr;
                        Cell cell = getCell(tc);
                        //this is a hack!! if ccolumn or row is 0 i.e. set the column/row cell to be string type
                        //this will overwrite whatever has been set in getCell()
                        if (k == 0 || i == 0) {
                            if (cell.getT().equals(STCellType.N)) {
                                System.out.println("Resetting!!");
                                cell.setT(STCellType.INLINE_STR);
                                CTXstringWhitespace ctx2 = smlObjectFactory.createCTXstringWhitespace();
                                ctx2.setValue(cell.getV());

                                CTRst ctrst2 = new CTRst();
                                ctrst2.setT(ctx2);
                                cell.setIs(ctrst2);
                                cell.setS(new Long(2));
                                cell.setV(null);

                            }
                        }
                        if (cell != null) {
                            ((Row) rows.get(i + 1)).getC().add(cell);//.get(colSize).sgetTc().add(getCell(tc));
                        }
                    }

                }
            }
        } else {
            Row row = smlObjectFactory.createRow();
            row.setR(new Long(2));

            row.setThickBot(Boolean.TRUE);

            row.setHt(17.25);
            row.setOutlineLevel(Short.decode("0"));
            row.setS(new Long(0));
            rows.add(row);
        }

        //lets fill last row with footer data
        if (pTablePair.footer != null) {

            for (int i = 0; i < pTablePair.footer.size(); i++) {
                Row footerrow = smlObjectFactory.createRow();
                footerrow.setR(new Long(2 + rows.size()));
                footerrow.setHt(17.25 * (1 + i));
                footerrow.setOutlineLevel(Short.decode("0"));
                footerrow.setS(new Long(13));
                Cell footerCell = smlObjectFactory.createCell();

                footerCell.setT(STCellType.INLINE_STR);

                CTXstringWhitespace ctx2 = smlObjectFactory.createCTXstringWhitespace();

                CTRst ctrst2 = new CTRst();
                ctrst2.setT(ctx2);
                footerCell.setIs(ctrst2);

                if (pTablePair.footer != null) {
                    ctx2.setValue(getText(pTablePair.footer.get(i)));//.getContent());
                }
                footerrow.getC().add(footerCell);

                //titlerow.setS(new Long(0));
                rows.add(footerrow);
            }
        }

        //lets fill last row with table text data
//        if (pTablePair.tabletext != null) {
//
//            //for (int i = 0; i < pTablePair.footer.size(); i++) {
//            Row tabletextrow = smlObjectFactory.createRow();
//            Cell tabletextCell = smlObjectFactory.createCell();
//
//            tabletextCell.setT(STCellType.INLINE_STR);
//
//            CTXstringWhitespace ctx2 = smlObjectFactory.createCTXstringWhitespace();
//
//            CTRst ctrst2 = new CTRst();
//            ctrst2.setT(ctx2);
//            ctx2.setValue(pTablePair.tabletext.toString());
//            tabletextCell.setIs(ctrst2);
//
//            tabletextrow.getC().add(tabletextCell);
//
//            //titlerow.setS(new Long(0));
//            System.out.println(tabletextrow);
//            rows.add(tabletextrow);
//            // }
//        }

        WorksheetPart worksheetPart = pkg.createWorksheetPart(partName, sheetName, tblCount + 1);

        //worksheetPart.g
//        WorkbookPart wb = pkg.getWorkbookPart();
//
//        //createStyles2();
//        //wb.getStylesPart().setJaxbElement(createStyles2());
//        wb.addTargetPart(styles);
//        wb.setPartShortcut(styles);
        worksheet.setSheetData(sd);
        worksheetPart.setJaxbElement(worksheet);
        //Relationship r = wb.addTargetPart(worksheetPart);
        tblCount++;

        return worksheetPart;
    }

    @SuppressWarnings("empty-statement")
    private static Cell getCell(org.docx4j.wml.Tc tc) throws JAXBException {
//        if (1 == 1) {
//            return null;
//        }

        String paratext = "";
        boolean bulletTextInCell = false;
        for (int k = 0; k < tc.getContent().size(); k++) {

            org.docx4j.wml.P tcpara;
            try {
                tcpara = (org.docx4j.wml.P) tc.getContent().get(k);
            } catch (Exception e) {
                continue;
            }

            String rowtext = "";
            for (int l = 0; l < tcpara.getContent().size(); l++) {
                //Dont assume its always a row
                org.docx4j.wml.R r;
                org.docx4j.wml.PPr pPr;

                if ((tcpara.getPPr() != null)) {
                    pPr = tcpara.getPPr();
                    //insert MnM bullet style if pPr had pStyle with value like %bullet%
                    //this is a hack as of now!!
                    if (pPr.getPStyle() != null) {
                        if (pPr.getPStyle().getVal().contains("bullet")) {//typically "tablebullet"
                            bulletTextInCell = true;
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
                            continue;
                        };
                        javax.xml.bind.JAXBElement jaxb = (javax.xml.bind.JAXBElement) (o);
                        switch (jaxb.getDeclaredType().getName()) {
                            //// also check if images or drwings are the
                            case "org.docx4j.wml.Text":
                                rowtext = rowtext + ((org.docx4j.wml.Text) (jaxb.getValue())).getValue();

                                break;
                            case "org.docx4j.wml.Drawing":
                                org.docx4j.wml.Drawing drawing = (org.docx4j.wml.Drawing) (jaxb.getValue());
                                org.docx4j.dml.wordprocessingDrawing.Inline inline = (org.docx4j.dml.wordprocessingDrawing.Inline) (drawing).getAnchorOrInline().get(0);
                                if (inline.getGraphic() != null) {
                                    //log.debug("found a:graphic");
                                    org.docx4j.dml.Graphic graphic = inline.getGraphic();
                                    if (graphic.getGraphicData() != null) {
                                        String imageId = graphic.getGraphicData().getPic().getBlipFill().getBlip().getEmbed();
                                        //System.out.println("Row " + (i + 1) + " column" + (j + 1) + "'s " + (k + 1) + " value is image " + imageId + " mapped to file " + relationsPart.getRelationshipByID(imageId).getTarget());
                                    }
                                }
                                break;
                        }

                    }

                }

            }
            if (bulletTextInCell) {
                rowtext = "." + rowtext + "\n";
            }
            paratext = paratext + rowtext;

        }

        Cell xclCell = smlObjectFactory.createCell();

//        System.out.println(StringUtils.isAlphanumeric(paratext));
        //System.out.println(Utils.isNumericFormat(paratext.trim()));
        paratext = paratext.trim();
        if (Utils.getNumericFormat(paratext).equals(Utils.NumericType.Number)) {//&&!paratext.startsWith("201")) {

            xclCell.setS(new Long(4));
            xclCell.setT(STCellType.N);
            String newparatext = paratext.replace(",", "");
//            CTXstringWhitespace ctx = smlObjectFactory.createCTXstringWhitespace();
//            ctx.setValue(paratext);
//
//            CTRst ctrst = new CTRst();
//            ctrst.setT(ctx);
            xclCell.setV(newparatext);
            //xclCell.setIs(ctrst);
            System.out.println(newparatext + " ::Number");
            //xclCell.setF(null
        } else if (Utils.getNumericFormat(paratext).equals(Utils.NumericType.Percent)) {

            xclCell.setS(new Long(5));
            xclCell.setT(STCellType.N);
            String newparatext = paratext.replace("%", "");
            double decimalNumber = Double.parseDouble(newparatext) / 100.0;

            xclCell.setV(String.valueOf(decimalNumber));

            System.out.println(String.valueOf(decimalNumber) + " ::Percent");
            //xclCell.setF(null
        } else {
            xclCell.setT(STCellType.INLINE_STR);

            CTXstringWhitespace ctx = smlObjectFactory.createCTXstringWhitespace();
            ctx.setValue(paratext);

            CTRst ctrst = new CTRst();
            ctrst.setT(ctx);
            xclCell.setIs(ctrst);
            xclCell.setS(new Long(2));
            System.out.println(paratext + " ::STR");

        }
        //System.out.println(paratext);
        return xclCell;
    }

    protected static int getGridAfter(Tr tr) {
        List<JAXBElement<?>> cnfStyleOrDivIdOrGridBefore = (tr.getTrPr() != null ? tr.getTrPr().getCnfStyleOrDivIdOrGridBefore() : null);
        JAXBElement element = getElement(cnfStyleOrDivIdOrGridBefore, "gridAfter");
        CTTrPrBase.GridAfter gridAfter = (element != null ? (CTTrPrBase.GridAfter) element.getValue() : null);
        BigInteger val = (gridAfter != null ? gridAfter.getVal() : null);
        return (val != null ? val.intValue() : 0);
    }

    private static void filterTables(List<PTablePair> pTablePairs) {

        pTablePairs.stream().forEach((ptp) -> {
            boolean bookmarkFound = false;
            for (int j = 0; j < ptp.title.getContent().size(); j++) {
                if (ptp.title.getContent().get(j) instanceof javax.xml.bind.JAXBElement) {
                    JAXBElement jaxb = (JAXBElement) ptp.title.getContent().get(j);
                    //org.docx4j.wml.P
                    if (jaxb.getDeclaredType().getName().equals("org.docx4j.wml.CTBookmark")) {

                        String tableId = ((org.docx4j.wml.CTBookmark) (jaxb.getValue())).getName();
                        //System.out.println(tableId + "::" + ptp.title.toString());
                        if (tableElements.containsKey(tableId)) {

                            ptp.setIndex(((Element) tableElements.get(tableId)).getindex());
                            sequenceExportList.put(sequenceList.get((Element) tableElements.get(tableId)), ptp);
                            bookmarkFound = true;
                        }
                    }
                }

            }
            if (!bookmarkFound) {
                //sequenceExportList()
            }
        });

    }

    private static JAXBElement<?> getElement(List<JAXBElement<?>> cnfStyleOrDivIdOrGridBefore, String localName) {
        JAXBElement<?> element;
        if ((cnfStyleOrDivIdOrGridBefore != null) && (!cnfStyleOrDivIdOrGridBefore.isEmpty())) {
            for (JAXBElement<?> cnfStyleOrDivIdOrGridBefore1 : cnfStyleOrDivIdOrGridBefore) {
                element = cnfStyleOrDivIdOrGridBefore1;
                if (localName.equals(element.getName().getLocalPart())) {
                    return element;
                }
            }
        }
        return null;
    }

    private static String getText(P p) {

        String text = "";
        if (p.getContent() == null) {
            return "";
        }
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
                        text = text.concat(((Text) (jaxb.getValue())).getValue());
                    }
                }
            }
        }
        return text;
    }

    private static CTStylesheet createStyles() {
        //org.xlsx4j.sml.ObjectFactory smlObjectFactory = new org.xlsx4j.sml.ObjectFactory();

        CTStylesheet stylesheet = smlObjectFactory.createCTStylesheet();
        //JAXBElement<org.xlsx4j.sml.CTStylesheet> stylesheetWrapped = smlObjectFactory.createStyleSheet(stylesheet);
        // Create object for fonts
        CTFonts fonts = smlObjectFactory.createCTFonts();
        stylesheet.setFonts(fonts);
        // Create object for font
        CTFont font = smlObjectFactory.createCTFont();
        fonts.getFont().add(font);
        // Create object for sz (wrapped in JAXBElement) 
        CTFontSize fontsize = smlObjectFactory.createCTFontSize();
        fontsize.setVal(new Double(12.0));
        JAXBElement<org.xlsx4j.sml.CTFontSize> fontsizeWrapped = smlObjectFactory.createCTFontSz(fontsize);
        font.getNameOrCharsetOrFamily().add(fontsizeWrapped);
        // Create object for color (wrapped in JAXBElement) 
        CTColor color = smlObjectFactory.createCTColor();
        JAXBElement<org.xlsx4j.sml.CTColor> colorWrapped = smlObjectFactory.createCTFontColor(color);
        font.getNameOrCharsetOrFamily().add(colorWrapped);
        //color.setTheme(new Long(1));
        color.setTint(new Double(0.0));
        // Create object for name (wrapped in JAXBElement) 
        CTFontName fontname = smlObjectFactory.createCTFontName();
        JAXBElement<org.xlsx4j.sml.CTFontName> fontnameWrapped = smlObjectFactory.createCTFontName(fontname);
        font.getNameOrCharsetOrFamily().add(fontnameWrapped);
        fontname.setVal("Calibri");
        // Create object for family (wrapped in JAXBElement) 
        CTFontFamily fontfamily = smlObjectFactory.createCTFontFamily();
        JAXBElement<org.xlsx4j.sml.CTFontFamily> fontfamilyWrapped = smlObjectFactory.createCTFontFamily(fontfamily);
        font.getNameOrCharsetOrFamily().add(fontfamilyWrapped);
        fontfamily.setVal(2);
        // Create object for scheme (wrapped in JAXBElement) 
        CTFontScheme fontscheme = smlObjectFactory.createCTFontScheme();
        JAXBElement<org.xlsx4j.sml.CTFontScheme> fontschemeWrapped = smlObjectFactory.createCTFontScheme(fontscheme);
        font.getNameOrCharsetOrFamily().add(fontschemeWrapped);
        fontscheme.setVal(org.xlsx4j.sml.STFontScheme.MINOR);
        // Create object for font
        CTFont font2 = smlObjectFactory.createCTFont();
        fonts.getFont().add(font2);
        // Create object for sz (wrapped in JAXBElement) 
        CTFontSize fontsize2 = smlObjectFactory.createCTFontSize();
        JAXBElement<org.xlsx4j.sml.CTFontSize> fontsizeWrapped2 = smlObjectFactory.createCTFontSz(fontsize2);
        fontsize2.setVal(new Double(10.0));
        font2.getNameOrCharsetOrFamily().add(fontsizeWrapped2);
        // Create object for name (wrapped in JAXBElement) 
        CTFontName fontname2 = smlObjectFactory.createCTFontName();
        JAXBElement<org.xlsx4j.sml.CTFontName> fontnameWrapped2 = smlObjectFactory.createCTFontName(fontname2);
        font2.getNameOrCharsetOrFamily().add(fontnameWrapped2);
        fontname2.setVal("Franklin Gothic Medium");
        // Create object for family (wrapped in JAXBElement) 
        CTFontFamily fontfamily2 = smlObjectFactory.createCTFontFamily();
        JAXBElement<org.xlsx4j.sml.CTFontFamily> fontfamilyWrapped2 = smlObjectFactory.createCTFontFamily(fontfamily2);
        font2.getNameOrCharsetOrFamily().add(fontfamilyWrapped2);
        fontfamily2.setVal(2);
        fonts.setCount(new Long(2));
        // Create object for fills
        CTFills fills = smlObjectFactory.createCTFills();
        stylesheet.setFills(fills);
        // Create object for fill
        CTFill fill = smlObjectFactory.createCTFill();
        fills.getFill().add(fill);
        // Create object for patternFill
        CTPatternFill patternfill = smlObjectFactory.createCTPatternFill();
        fill.setPatternFill(patternfill);
        patternfill.setPatternType(org.xlsx4j.sml.STPatternType.NONE);
        // Create object for fill
        CTFill fill2 = smlObjectFactory.createCTFill();
        fills.getFill().add(fill2);
        // Create object for patternFill
        CTPatternFill patternfill2 = smlObjectFactory.createCTPatternFill();
        fill2.setPatternFill(patternfill2);
        patternfill2.setPatternType(org.xlsx4j.sml.STPatternType.GRAY_125);
        fills.setCount(new Long(2));
        // Create object for borders
        CTBorders borders = smlObjectFactory.createCTBorders();
        stylesheet.setBorders(borders);
        // Create object for border
        CTBorder border = smlObjectFactory.createCTBorder();
        borders.getBorder().add(border);
        // Create object for left
        CTBorderPr borderpr = smlObjectFactory.createCTBorderPr();
        border.setLeft(borderpr);
        borderpr.setStyle(org.xlsx4j.sml.STBorderStyle.NONE);
        // Create object for right
        CTBorderPr borderpr2 = smlObjectFactory.createCTBorderPr();
        border.setRight(borderpr2);
        borderpr2.setStyle(org.xlsx4j.sml.STBorderStyle.NONE);
        // Create object for diagonal
        CTBorderPr borderpr3 = smlObjectFactory.createCTBorderPr();
        border.setDiagonal(borderpr3);
        borderpr3.setStyle(org.xlsx4j.sml.STBorderStyle.NONE);
        // Create object for top
        CTBorderPr borderpr4 = smlObjectFactory.createCTBorderPr();
        border.setTop(borderpr4);
        borderpr4.setStyle(org.xlsx4j.sml.STBorderStyle.NONE);
        // Create object for bottom
        CTBorderPr borderpr5 = smlObjectFactory.createCTBorderPr();
        border.setBottom(borderpr5);
        borderpr5.setStyle(org.xlsx4j.sml.STBorderStyle.NONE);
        borders.setCount(new Long(1));
        // Create object for cellStyleXfs
        CTCellStyleXfs cellstylexfs = smlObjectFactory.createCTCellStyleXfs();
        cellstylexfs.setCount(new Long(1));

        stylesheet.setCellStyleXfs(cellstylexfs);
        // Create object for xf
        CTXf xf = smlObjectFactory.createCTXf();
        cellstylexfs.getXf().add(xf);
        xf.setFontId(new Long(0));
        xf.setFillId(new Long(0));
        xf.setBorderId(new Long(0));
        xf.setNumFmtId(new Long(0));
        // Create object for cellXfs
        CTCellXfs cellxfs = smlObjectFactory.createCTCellXfs();
        stylesheet.setCellXfs(cellxfs);
        // Create object for xf
        CTXf xf1 = smlObjectFactory.createCTXf();
        cellxfs.getXf().add(xf1);
        xf1.setFontId(new Long(0));
        xf1.setXfId(new Long(0));
        xf1.setFillId(new Long(0));
        xf1.setBorderId(new Long(0));
        xf1.setNumFmtId(new Long(0));
        // Create object for xf
        CTXf xf2 = smlObjectFactory.createCTXf();
        cellxfs.getXf().add(xf2);
        xf2.setFontId(new Long(1));
        xf2.setXfId(new Long(0));
        xf2.setFillId(new Long(0));
        xf2.setBorderId(new Long(0));
        xf2.setNumFmtId(new Long(0));
        CTCellAlignment cellalignment2 = smlObjectFactory.createCTCellAlignment();
        cellalignment2.setVertical(STVerticalAlignment.TOP);
        cellalignment2.setWrapText(Boolean.TRUE);
        xf2.setAlignment(cellalignment2);
        // Create object for xf
        CTXf xf3 = smlObjectFactory.createCTXf();
        cellxfs.getXf().add(xf3);
        xf3.setFontId(new Long(1));
        xf3.setXfId(new Long(0));
        xf3.setFillId(new Long(0));
        xf3.setBorderId(new Long(0));
        xf3.setNumFmtId(new Long(0));
        // Create object for alignment
        CTCellAlignment cellalignment3 = smlObjectFactory.createCTCellAlignment();
        xf3.setAlignment(cellalignment3);
        cellalignment3.setVertical(org.xlsx4j.sml.STVerticalAlignment.TOP);
        cellalignment3.setWrapText(Boolean.TRUE);

        // Create object for xf
        CTXf xf4 = smlObjectFactory.createCTXf();
        cellxfs.getXf().add(xf4);
        xf4.setFontId(new Long(1));
        xf4.setXfId(new Long(0));
        xf4.setFillId(new Long(0));
        xf4.setBorderId(new Long(0));
        xf4.setNumFmtId(new Long(0));
        // Create object for alignment
        CTCellAlignment cellalignment4 = smlObjectFactory.createCTCellAlignment();
        xf4.setAlignment(cellalignment4);
        cellalignment4.setVertical(org.xlsx4j.sml.STVerticalAlignment.CENTER);
        cellalignment4.setHorizontal(STHorizontalAlignment.LEFT);
        cellalignment4.setIndent(new Long(8));

        CTXf xf5 = smlObjectFactory.createCTXf();
        cellxfs.getXf().add(xf5);
        xf5.setFontId(new Long(1));
        xf5.setXfId(new Long(0));
        xf5.setFillId(new Long(0));
        xf5.setBorderId(new Long(0));
        xf5.setNumFmtId(new Long(4));
        // Create object for alignment
        CTCellAlignment cellalignment5 = smlObjectFactory.createCTCellAlignment();
        xf5.setAlignment(cellalignment5);
        xf5.setApplyNumberFormat(Boolean.TRUE);
        cellalignment5.setVertical(org.xlsx4j.sml.STVerticalAlignment.TOP);

        CTXf xf6 = smlObjectFactory.createCTXf();
        cellxfs.getXf().add(xf6);
        xf6.setFontId(new Long(1));
        xf6.setXfId(new Long(0));
        xf6.setFillId(new Long(0));
        xf6.setBorderId(new Long(0));
        xf6.setNumFmtId(new Long(10));
        // Create object for alignment
        CTCellAlignment cellalignment6 = smlObjectFactory.createCTCellAlignment();
        xf6.setAlignment(cellalignment6);
        xf6.setApplyNumberFormat(Boolean.TRUE);
        cellalignment6.setVertical(org.xlsx4j.sml.STVerticalAlignment.TOP);
        //update count here if added/deleted xfs!!!
        cellxfs.setCount(new Long(6));
        // Create object for cellStyles
        CTCellStyles cellstyles = smlObjectFactory.createCTCellStyles();
        stylesheet.setCellStyles(cellstyles);
        // Create object for cellStyle
        CTCellStyle cellstyle = smlObjectFactory.createCTCellStyle();
        cellstyles.getCellStyle().add(cellstyle);
        cellstyle.setXfId(0);
        cellstyle.setBuiltinId(new Long(0));
        cellstyle.setName("Normal");

        cellstyles.setCount(new Long(1));
        // Create object for dxfs
        CTDxfs dxfs = smlObjectFactory.createCTDxfs();
        stylesheet.setDxfs(dxfs);
        // Create object for dxf
        CTDxf dxf = smlObjectFactory.createCTDxf();
        dxfs.getDxf().add(dxf);
        // Create object for font
        CTFont font3 = smlObjectFactory.createCTFont();
        dxf.setFont(font3);
        // Create object for b (wrapped in JAXBElement) 
        CTBooleanProperty booleanproperty = smlObjectFactory.createCTBooleanProperty();
        JAXBElement<org.xlsx4j.sml.CTBooleanProperty> booleanpropertyWrapped = smlObjectFactory.createCTFontB(booleanproperty);
        font3.getNameOrCharsetOrFamily().add(booleanpropertyWrapped);
        // Create object for i (wrapped in JAXBElement) 
        CTBooleanProperty booleanproperty2 = smlObjectFactory.createCTBooleanProperty();
        JAXBElement<org.xlsx4j.sml.CTBooleanProperty> booleanpropertyWrapped2 = smlObjectFactory.createCTFontI(booleanproperty2);
        font3.getNameOrCharsetOrFamily().add(booleanpropertyWrapped2);
        // Create object for fill
        CTFill fill3 = smlObjectFactory.createCTFill();
        dxf.setFill(fill3);
        // Create object for patternFill
        CTPatternFill patternfill3 = smlObjectFactory.createCTPatternFill();
        fill3.setPatternFill(patternfill3);
        // Create object for bgColor
        CTColor color2 = smlObjectFactory.createCTColor();
        patternfill3.setBgColor(color2);

        color2.setTint(new Double(0.0));
        // Create object for dxf
        CTDxf dxf2 = smlObjectFactory.createCTDxf();
        dxfs.getDxf().add(dxf2);
        // Create object for font
        CTFont font4 = smlObjectFactory.createCTFont();
        dxf2.setFont(font4);
        // Create object for b (wrapped in JAXBElement) 
        CTBooleanProperty booleanproperty3 = smlObjectFactory.createCTBooleanProperty();
        JAXBElement<org.xlsx4j.sml.CTBooleanProperty> booleanpropertyWrapped3 = smlObjectFactory.createCTFontB(booleanproperty3);
        font4.getNameOrCharsetOrFamily().add(booleanpropertyWrapped3);
        // Create object for i (wrapped in JAXBElement) 
        CTBooleanProperty booleanproperty4 = smlObjectFactory.createCTBooleanProperty();
        JAXBElement<org.xlsx4j.sml.CTBooleanProperty> booleanpropertyWrapped4 = smlObjectFactory.createCTFontI(booleanproperty4);
        font4.getNameOrCharsetOrFamily().add(booleanpropertyWrapped4);
        // Create object for fill
        CTFill fill4 = smlObjectFactory.createCTFill();
        dxf2.setFill(fill4);
        // Create object for patternFill
        CTPatternFill patternfill4 = smlObjectFactory.createCTPatternFill();
        fill4.setPatternFill(patternfill4);
        // Create object for bgColor
        CTColor color3 = smlObjectFactory.createCTColor();
        patternfill4.setBgColor(color3);
        color3.setIndexed(new Long(65));
        color3.setTint(new Double(0.0));
        patternfill4.setPatternType(org.xlsx4j.sml.STPatternType.NONE);
        dxfs.setCount(new Long(2));
        // Create object for tableStyles
        CTTableStyles tablestyles = smlObjectFactory.createCTTableStyles();
        stylesheet.setTableStyles(tablestyles);
        tablestyles.setDefaultPivotStyle("PivotStyleMedium9");
        tablestyles.setDefaultTableStyle("TableStyleMedium2");
        // Create object for tableStyle
        CTTableStyle tablestyle = smlObjectFactory.createCTTableStyle();
        tablestyles.getTableStyle().add(tablestyle);
        // Create object for tableStyleElement
        CTTableStyleElement tablestyleelement = smlObjectFactory.createCTTableStyleElement();
        tablestyle.getTableStyleElement().add(tablestyleelement);
        tablestyleelement.setDxfId(new Long(1));
        tablestyleelement.setType(org.xlsx4j.sml.STTableStyleType.WHOLE_TABLE);
        tablestyleelement.setSize(new Long(1));

        // Create object for tableStyleElement
        CTTableStyleElement tablestyleelement2 = smlObjectFactory.createCTTableStyleElement();
        tablestyle.getTableStyleElement().add(tablestyleelement2);
        tablestyleelement2.setDxfId(new Long(0));
        tablestyleelement2.setType(org.xlsx4j.sml.STTableStyleType.HEADER_ROW);
        tablestyleelement2.setSize(new Long(1));
        tablestyle.setName("MySqlDefault");
        tablestyle.setCount(new Long(2));
        tablestyles.setCount(new Long(1));
        // Create object for extLst
        CTExtensionList extensionlist = smlObjectFactory.createCTExtensionList();
        stylesheet.setExtLst(extensionlist);
        // Create object for ext
        CTExtension extension = smlObjectFactory.createCTExtension();
        extensionlist.getExt().add(extension);
        extension.setUri("{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}");
        // Create object for ext
        CTExtension extension2 = smlObjectFactory.createCTExtension();
        extensionlist.getExt().add(extension2);
        extension2.setUri("{9260A510-F301-46a8-8635-F512D64BE5F5}");

        return stylesheet;
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

    private static SheetData createLicenseSheetData() {
        //org.xlsx4j.sml.ObjectFactory smlObjectFactory = new org.xlsx4j.sml.ObjectFactory();

        // Create object for sheetData
        SheetData sheetdata = smlObjectFactory.createSheetData();

        // Create object for pageMargins
        // Create object for dimension
        Row row = smlObjectFactory.createRow();
        sheetdata.getRow().add(row);
        // Create object for c
        Cell cell = smlObjectFactory.createCell();
        row.getC().add(cell);
        cell.setVm(new Long(0));
        cell.setCm(new Long(0));
        cell.setR("A1");
        cell.setT(STCellType.INLINE_STR);

        CTXstringWhitespace ctx = smlObjectFactory.createCTXstringWhitespace();
        ctx.setValue("");

        CTRst ctrst = new CTRst();
        ctrst.setT(ctx);
        cell.setIs(ctrst);

        cell.setS(new Long(1));
        row.setOutlineLevel(Short.decode("0"));
        row.setR(new Long(1));
        row.setS(new Long(1));
        row.setCustomFormat(Boolean.TRUE);
        // Create object for row
        Row row2 = smlObjectFactory.createRow();
        sheetdata.getRow().add(row2);
        // Create object for c
        Cell cell2 = smlObjectFactory.createCell();
        row2.getC().add(cell2);
        cell2.setVm(new Long(0));
        cell2.setCm(new Long(0));
        cell2.setR("A7");
        cell2.setT(STCellType.INLINE_STR);

        CTXstringWhitespace ctx2 = smlObjectFactory.createCTXstringWhitespace();
        if (header != null && !header.isEmpty()) {
            ctx2.setValue(header);
        } else {
            ctx2.setValue("A Report By MarketsAndMarkets");
        }

        CTRst ctrst2 = new CTRst();
        ctrst2.setT(ctx2);
        cell2.setIs(ctrst2);

        cell2.setS(new Long(1));

        row2.setOutlineLevel(Short.decode("0"));
        row2.setR(new Long(7));
        row2.setS(new Long(1));
        // Create object for row
        Row row3 = smlObjectFactory.createRow();
        sheetdata.getRow().add(row3);
        // Create object for c
        Cell cell3 = smlObjectFactory.createCell();
        row3.getC().add(cell3);
        cell3.setVm(new Long(0));
        cell3.setCm(new Long(0));
        cell3.setR("A11");
        cell3.setT(STCellType.INLINE_STR);

        CTXstringWhitespace ctx3 = smlObjectFactory.createCTXstringWhitespace();
        ctx3.setValue("Copyright Â©MarketsandMarkets");

        CTRst ctrst3 = new CTRst();
        ctrst3.setT(ctx3);
        cell3.setIs(ctrst3);

        cell3.setS(new Long(1));

        row3.setOutlineLevel(Short.decode("0"));
        row3.setR(new Long(11));
        row3.setS(new Long(1));
        // Create object for row
        Row row4 = smlObjectFactory.createRow();
        sheetdata.getRow().add(row4);
        // Create object for c
        Cell cell4 = smlObjectFactory.createCell();
        row4.getC().add(cell4);
        cell4.setVm(new Long(0));
        cell4.setCm(new Long(0));
        cell4.setR("A12");
        cell4.setT(STCellType.INLINE_STR);

        CTXstringWhitespace ctx4 = smlObjectFactory.createCTXstringWhitespace();
        ctx4.setValue("All Rights Reserved.This document contains highly confidential information and is the sole property of MarketsandMarkets. No part of it may be circulated, copied, quoted, or otherwise reproduced without the written approval of MarketsandMarkets.");

        CTRst ctrst4 = new CTRst();
        ctrst4.setT(ctx4);
        cell4.setIs(ctrst4);
        cell4.setS(new Long(2));
        row4.setOutlineLevel(Short.decode("0"));
        row4.setHt(new Double(25.5));
        row4.setR(new Long(12));
        row4.setS(new Long(0));
        // Create object for row
        Row row5 = smlObjectFactory.createRow();
        sheetdata.getRow().add(row5);
        // Create object for c
        Cell cell5 = smlObjectFactory.createCell();
        row5.getC().add(cell5);
        cell5.setVm(new Long(0));
        cell5.setCm(new Long(0));
        cell5.setR("A14");
        cell5.setT(STCellType.INLINE_STR);
        cell5.setS(new Long(2));

        CTXstringWhitespace ctx5 = smlObjectFactory.createCTXstringWhitespace();
        ctx5.setValue("Disclaimer: MarketsandMarkets strategic analysis services are limited publications containing valuable market information provided to a select group of customers in response to orders. Our customers acknowledge, when ordering that MarketsandMarkets strategic analysis services are for our customersâ€™ internal use and not for general publication or disclosure to third parties. Quantitative market information is based primarily on interviews and therefore, is subject to fluctuation.\n"
                + "\n"
                + "MarketsandMarkets does not endorse any vendor, product or service depicted in its research publications. MarketsandMarkets strategic analysis publications consist of the opinions of MarketsandMarketsâ€™ research and should not be construed as statements of fact. MarketsandMarkets disclaims all warranties, expressed or implied, with respect to this research, including any warranties of merchantability or fitness for a particular purpose.\n"
                + "\n"
                + "MarketsandMarkets takes no responsibility for any incorrect information supplied to us by manufacturers or users.\n"
                + "\n"
                + "All trademarks, copyrights and other forms of intellectual property belong to their respective owners and may be protected by copyright. Under no circumstance may any of these be reproduced in any form without the prior written agreement of their owner.\n"
                + "\n"
                + "No part of this strategic analysis service may be given, lent, resold or disclosed to non-customers without written permission.\n"
                + "Reproduction and/or transmission in any form and by any means including photocopying, mechanical, electronic, recording or otherwise, without the permission of the publisher is prohibited.\n"
                + "\n"
                + "For information regarding permission, contact: \n"
                + "Tel: 1-888-600-6441\n"
                + "Email:sales@marketsandmarkets.com");

        CTRst ctrst5 = new CTRst();
        ctrst5.setT(ctx5);
        cell5.setIs(ctrst5);
        row5.setOutlineLevel(Short.decode("0"));
        row5.setR(new Long(14));
        row5.setS(new Long(2));

        //return worksheetWrapped;
        return sheetdata;
    }

    private static WorksheetPart createLicenseWorkSheetPart() throws InvalidFormatException, JAXBException {
        PartName partName = new PartName("/xl/worksheets/disclaimer.xml");
        String sheetName = "Disclaimer";
        //long sheetId;

        //Sheets sheets = wb.getJaxbElement().getSheets();
        Worksheet worksheet = smlObjectFactory.createWorksheet();

        // Create object for pageMargins
        CTPageMargins pagemargins = smlObjectFactory.createCTPageMargins();
        worksheet.setPageMargins(pagemargins);
        CTPageSetup pagesetup = smlObjectFactory.createCTPageSetup();
        worksheet.setPageSetup(pagesetup);
        pagesetup.setErrors(org.xlsx4j.sml.STPrintError.DISPLAYED);
        pagesetup.setOrientation(org.xlsx4j.sml.STOrientation.PORTRAIT);
        pagesetup.setPaperSize(new Long(1));
        pagesetup.setFirstPageNumber(new Long(1));
        pagesetup.setHorizontalDpi(new Long(1200));
        pagesetup.setVerticalDpi(new Long(1200));
        pagesetup.setCopies(new Long(1));
        pagesetup.setScale(new Long(100));
        pagesetup.setFitToWidth(new Long(1));
        pagesetup.setFitToHeight(new Long(1));
        pagesetup.setPageOrder(org.xlsx4j.sml.STPageOrder.DOWN_THEN_OVER);
        pagesetup.setCellComments(org.xlsx4j.sml.STCellComments.NONE);

        //SheetData sd = smlObjectFactory.createSheetData();
        CTSheetFormatPr sheetformatpr = smlObjectFactory.createCTSheetFormatPr();
        sheetformatpr.setBaseColWidth(new Long(8));
        sheetformatpr.setOutlineLevelRow(Short.decode("0"));
        sheetformatpr.setOutlineLevelCol(Short.decode("0"));
        worksheet.setSheetFormatPr(sheetformatpr);

        SheetViews sheetViews = smlObjectFactory.createSheetViews();
        SheetView sheetview = smlObjectFactory.createSheetView();

        sheetview.setColorId(new Long(64));
        sheetview.setZoomScale(new Long(100));
        sheetview.setZoomScaleNormal(new Long(0));
        sheetview.setZoomScaleSheetLayoutView(new Long(0));
        sheetview.setZoomScalePageLayoutView(new Long(0));
        //sheetview.setWorkbookViewId(0);
        sheetview.setView(org.xlsx4j.sml.STSheetViewType.NORMAL);
        sheetViews.getSheetView().add(sheetview);
        worksheet.setSheetViews(sheetViews);

        org.xlsx4j.sml.ObjectFactory smlObjectFactory = new org.xlsx4j.sml.ObjectFactory();

        Cols cols = smlObjectFactory.createCols();
        // Create object for col
        Col col = smlObjectFactory.createCol();
        cols.getCol().add(col);
        col.setMin(1);
        col.setOutlineLevel(Short.decode("0"));
        col.setStyle(new Long(1));
        col.setMax(1);
        col.setWidth(new Double(167.0));
        // Create object for col
        Col col2 = smlObjectFactory.createCol();
        cols.getCol().add(col2);
        col2.setMin(2);
        col2.setOutlineLevel(Short.decode("0"));
        col2.setStyle(new Long(1));
        col2.setMax(16384);
        col2.setWidth(new Double(9.140625));
        worksheet.getCols().add(cols);
//        List rows = sd.getRow();

        WorksheetPart worksheetPart = pkg.createWorksheetPart(partName, sheetName, 1);

        PartName drawingPartName = new PartName("/xl/drawings/drawing.xml");
        //worksheetPart.addTargetPart(drawingPartName);
        //worksheetPart.g
//        WorkbookPart wb = pkg.getWorkbookPart();
//
//        //createStyles2();
//        //wb.getStylesPart().setJaxbElement(createStyles2());
//        wb.addTargetPart(styles);
//        wb.setPartShortcut(styles);
        worksheet.setSheetData(createLicenseSheetData());
//        CTDrawing drawing = smlObjectFactory.createCTDrawing(); 
//        drawing.setId( "rId1"); 
//        worksheet.setDrawing(drawing);

        worksheetPart.setJaxbElement(worksheet);
        //Relationship r = wb.addTargetPart(worksheetPart);
        //tblCount++;

        return worksheetPart;

    }

    private static org.docx4j.dml.spreadsheetdrawing.CTDrawing createLicenseLogo() {
        org.docx4j.dml.spreadsheetdrawing.ObjectFactory dmlspreadsheetdrawingObjectFactory = new org.docx4j.dml.spreadsheetdrawing.ObjectFactory();

        org.docx4j.dml.spreadsheetdrawing.CTDrawing drawing = dmlspreadsheetdrawingObjectFactory.createCTDrawing();
        //JAXBElement<org.docx4j.dml.spreadsheetdrawing.CTDrawing> drawingWrapped = smlObjectFactory.createWsDr(drawing);
        // Create object for twoCellAnchor
        CTTwoCellAnchor twocellanchor = dmlspreadsheetdrawingObjectFactory.createCTTwoCellAnchor();
        drawing.getEGAnchor().add(twocellanchor);
        // Create object for clientData
        CTAnchorClientData anchorclientdata = dmlspreadsheetdrawingObjectFactory.createCTAnchorClientData();
        twocellanchor.setClientData(anchorclientdata);
        // Create object for pic
        CTPicture picture = dmlspreadsheetdrawingObjectFactory.createCTPicture();
        twocellanchor.setPic(picture);
        org.docx4j.dml.ObjectFactory dmlObjectFactory = new org.docx4j.dml.ObjectFactory();
        // Create object for blipFill
        CTBlipFillProperties blipfillproperties = dmlObjectFactory.createCTBlipFillProperties();
        picture.setBlipFill(blipfillproperties);
        // Create object for blip
        CTBlip blip = dmlObjectFactory.createCTBlip();
        blipfillproperties.setBlip(blip);
        blip.setEmbed("rId1");
        // Create object for extLst
        CTOfficeArtExtensionList officeartextensionlist = dmlObjectFactory.createCTOfficeArtExtensionList();
        blip.setExtLst(officeartextensionlist);
        // Create object for ext
        CTOfficeArtExtension officeartextension = dmlObjectFactory.createCTOfficeArtExtension();
        officeartextensionlist.getExt().add(officeartextension);
        officeartextension.setUri("{28A0092B-C50C-407E-A947-70E740481C1C}");
        blip.setCstate(org.docx4j.dml.STBlipCompression.NONE);
        blip.setLink("");
        // Create object for srcRect
        CTRelativeRect relativerect = dmlObjectFactory.createCTRelativeRect();
        blipfillproperties.setSrcRect(relativerect);
        relativerect.setB(0);
        relativerect.setR(0);
        relativerect.setT(0);
        relativerect.setL(0);
        // Create object for stretch
        CTStretchInfoProperties stretchinfoproperties = dmlObjectFactory.createCTStretchInfoProperties();
        blipfillproperties.setStretch(stretchinfoproperties);
        // Create object for fillRect
        CTRelativeRect relativerect2 = dmlObjectFactory.createCTRelativeRect();
        stretchinfoproperties.setFillRect(relativerect2);
        relativerect2.setB(0);
        relativerect2.setR(0);
        relativerect2.setT(0);
        relativerect2.setL(0);
        picture.setMacro("");
        // Create object for spPr
        CTShapeProperties shapeproperties = dmlObjectFactory.createCTShapeProperties();
        picture.setSpPr(shapeproperties);
        // Create object for noFill
        CTNoFillProperties nofillproperties = dmlObjectFactory.createCTNoFillProperties();
        shapeproperties.setNoFill(nofillproperties);
        // Create object for xfrm
        CTTransform2D transform2d = dmlObjectFactory.createCTTransform2D();
        shapeproperties.setXfrm(transform2d);
        // Create object for ext
        CTPositiveSize2D positivesize2d = dmlObjectFactory.createCTPositiveSize2D();
        transform2d.setExt(positivesize2d);
        positivesize2d.setCx(1619250);
        positivesize2d.setCy(571500);
        transform2d.setRot(0);
        // Create object for off
        CTPoint2D point2d = dmlObjectFactory.createCTPoint2D();
        transform2d.setOff(point2d);
        point2d.setY(200025);
        point2d.setX(304800);
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
        org.docx4j.dml.spreadsheetdrawing.CTPictureNonVisual picturenonvisual = dmlspreadsheetdrawingObjectFactory.createCTPictureNonVisual();
        picture.setNvPicPr(picturenonvisual);
        // Create object for cNvPr
        CTNonVisualDrawingProps nonvisualdrawingprops = dmlObjectFactory.createCTNonVisualDrawingProps();
        picturenonvisual.setCNvPr(nonvisualdrawingprops);
        nonvisualdrawingprops.setDescr("Description: Description: Description: Description: Description: Description: Description: Description: MnM logo_17");
        nonvisualdrawingprops.setName("Picture 4");
        nonvisualdrawingprops.setId(5);
        // Create object for cNvPicPr
        CTNonVisualPictureProperties nonvisualpictureproperties = dmlObjectFactory.createCTNonVisualPictureProperties();
        picturenonvisual.setCNvPicPr(nonvisualpictureproperties);
        // Create object for to
        CTMarker marker = dmlspreadsheetdrawingObjectFactory.createCTMarker();
        twocellanchor.setTo(marker);
        marker.setCol(0);
        marker.setColOff(1924050);
        marker.setRow(4);
        marker.setRowOff(9525);
        // Create object for from
        CTMarker marker2 = dmlspreadsheetdrawingObjectFactory.createCTMarker();
        twocellanchor.setFrom(marker2);
        marker2.setCol(0);
        marker2.setColOff(304800);
        marker2.setRow(1);
        marker2.setRowOff(9525);
        twocellanchor.setEditAs(org.docx4j.dml.spreadsheetdrawing.STEditAs.ONE_CELL);

        //pkg.addTargetPart(drawing);
        return drawing;
    }

//    private static void addDisclaimer(WorkbookPart wb) throws InvalidFormatException, JAXBException {
//        WorksheetPart licenseWorkSheetPart = createLicenseWorkSheetPart();
//        wb.addTargetPart(licenseWorkSheetPart);
//        Workbook workbook = smlObjectFactory.createWorkbook();
//        JAXBElement<org.xlsx4j.sml.Workbook> workbookWrapped = smlObjectFactory.createWorkbook(workbook);
//        // Create object for sheets
//        Sheets sheets = smlObjectFactory.createSheets();
//        workbook.setSheets(sheets);
//        // Create object for sheet
//        Sheet sheet = smlObjectFactory.createSheet();
//        sheets.getSheet().add(sheet);
//        sheet.setSheetId(1);
//        sheet.setName("Sheet1");
//        sheet.setId("rId1");
//        sheet.setState(org.xlsx4j.sml.STSheetState.VISIBLE);
//        wb.setContents(workbook);
//
////        Drawing drawing = new Drawing(new PartName("/xl/drawings/drawing1.xml"));
////        drawing.setContents(createLicenseLogo());
////        wb.addTargetPart(drawing);
////        ImagePngPart imgpngPart = new ImagePngPart(new PartName("/xl/media/image1.png"));
////        File file = new File(System.getProperty("user.dir") + "/src/test/resources/images/apple_web.png");
////
////        try {
////            imgpngPart.setBinaryData(FileUtils.readFileToByteArray(file));
////        } catch (IOException ex) {
////            Logger.getLogger(DocxToXcl.class.getName()).log(Level.SEVERE, null, ex);
////        }
////        licenseWorkSheetPart.addTargetPart(imgpngPart);
//        //wb.addTargetPart(imgpngPart);
//    }
    private static void createTocElements() {

        //mobile 3d market
        //String elems = "[\"T\",\"_Toc348352750\",\"1\"],[\"T\",\"_Toc348352751\",\"2\"],[\"T\",\"_Toc348352752\",\"3\"],[\"T\",\"_Toc348352753\",\"4\"],[\"T\",\"_Toc348352754\",\"5\"],[\"T\",\"_Toc348352755\",\"6\"],[\"T\",\"_Toc348352756\",\"7\"],[\"T\",\"_Toc348352757\",\"8\"],[\"T\",\"_Toc348352758\",\"9\"],[\"T\",\"_Toc348352759\",\"10\"],[\"T\",\"_Toc348352760\",\"11\"],[\"T\",\"_Toc348352761\",\"12\"],[\"T\",\"_Toc348352762\",\"13\"],[\"T\",\"_Toc348352763\",\"14\"],[\"T\",\"_Toc348352764\",\"15\"],[\"T\",\"_Toc348352765\",\"16\"],[\"T\",\"_Toc348352766\",\"17\"],[\"T\",\"_Toc348352767\",\"18\"],[\"T\",\"_Toc348352768\",\"19\"],[\"T\",\"_Toc348352769\",\"20\"],[\"T\",\"_Toc348352770\",\"21\"],[\"T\",\"_Toc348352771\",\"22\"],[\"T\",\"_Toc348352772\",\"23\"],[\"T\",\"_Toc348352773\",\"24\"],[\"T\",\"_Toc348352774\",\"25\"],[\"T\",\"_Toc348352775\",\"26\"],[\"T\",\"_Toc348352776\",\"27\"],[\"T\",\"_Toc348352777\",\"28\"],[\"T\",\"_Toc348352778\",\"29\"],[\"T\",\"_Toc348352779\",\"30\"],[\"T\",\"_Toc348352780\",\"31\"],[\"T\",\"_Toc348352781\",\"32\"],[\"T\",\"_Toc348352782\",\"33\"],[\"T\",\"_Toc348352783\",\"34\"],[\"T\",\"_Toc348352784\",\"35\"],[\"T\",\"_Toc348352785\",\"36\"],[\"T\",\"_Toc348352786\",\"37\"],[\"T\",\"_Toc348352787\",\"38\"],[\"T\",\"_Toc348352788\",\"39\"],[\"T\",\"_Toc348352789\",\"40\"],[\"T\",\"_Toc348352790\",\"41\"],[\"T\",\"_Toc348352791\",\"42\"],[\"T\",\"_Toc348352792\",\"43\"],[\"T\",\"_Toc348352793\",\"44\"],[\"T\",\"_Toc348352794\",\"45\"],[\"T\",\"_Toc348352795\",\"46\"],[\"T\",\"_Toc348352796\",\"47\"],[\"T\",\"_Toc348352797\",\"48\"],[\"T\",\"_Toc348352798\",\"49\"],[\"T\",\"_Toc348352799\",\"50\"],[\"T\",\"_Toc348352800\",\"51\"],[\"T\",\"_Toc348352801\",\"52\"],[\"T\",\"_Toc348352802\",\"53\"],[\"T\",\"_Toc348352803\",\"54\"],[\"T\",\"_Toc348352804\",\"55\"],[\"T\",\"_Toc348352805\",\"56\"],[\"T\",\"_Toc348352806\",\"57\"],[\"T\",\"_Toc348352807\",\"58\"],[\"T\",\"_Toc348352808\",\"59\"],[\"T\",\"_Toc348352809\",\"60\"],[\"T\",\"_Toc348352810\",\"61\"],[\"T\",\"_Toc348352811\",\"62\"],[\"T\",\"_Toc348352812\",\"63\"],[\"T\",\"_Toc348352813\",\"64\"],[\"T\",\"_Toc348352814\",\"65\"],[\"T\",\"_Toc348352815\",\"66\"],[\"T\",\"_Toc348352816\",\"67\"],[\"T\",\"_Toc348352817\",\"68\"],[\"T\",\"_Toc348352818\",\"69\"],[\"T\",\"_Toc348352819\",\"70\"],[\"T\",\"_Toc348352820\",\"71\"],[\"T\",\"_Toc348352821\",\"72\"],[\"T\",\"_Toc348352822\",\"73\"],[\"T\",\"_Toc348352823\",\"74\"],[\"T\",\"_Toc348352824\",\"75\"],[\"T\",\"_Toc348352825\",\"76\"],[\"T\",\"_Toc348352826\",\"77\"],[\"T\",\"_Toc348352827\",\"78\"],[\"T\",\"_Toc348352828\",\"79\"],[\"T\",\"_Toc348352829\",\"80\"],[\"T\",\"_Toc348352830\",\"81\"],[\"T\",\"_Toc348352831\",\"82\"],[\"T\",\"_Toc348352832\",\"83\"],[\"T\",\"_Toc348352833\",\"84\"],[\"T\",\"_Toc348352834\",\"85\"],[\"T\",\"_Toc348352835\",\"86\"],[\"T\",\"_Toc348352836\",\"87\"],[\"T\",\"_Toc348352837\",\"88\"],[\"T\",\"_Toc348352838\",\"89\"],[\"T\",\"_Toc348352839\",\"90\"],[\"T\",\"_Toc348352840\",\"91\"],[\"T\",\"_Toc348352841\",\"92\"],[\"T\",\"_Toc348352842\",\"93\"],[\"T\",\"_Toc348352843\",\"94\"],[\"T\",\"_Toc348352844\",\"95\"],[\"T\",\"_Toc348352845\",\"96\"],[\"T\",\"_Toc348352846\",\"97\"],[\"T\",\"_Toc348352847\",\"98\"],[\"T\",\"_Toc348352848\",\"99\"],[\"T\",\"_Toc348352849\",\"100\"],[\"T\",\"_Toc348352850\",\"101\"],[\"T\",\"_Toc348352851\",\"102\"],[\"T\",\"_Toc348352852\",\"103\"],[\"T\",\"_Toc348352853\",\"104\"],[\"T\",\"_Toc348352854\",\"105\"],[\"T\",\"_Toc348352855\",\"106\"],[\"T\",\"_Toc348352856\",\"107\"],[\"T\",\"_Toc348352857\",\"108\"],[\"T\",\"_Toc348352858\",\"109\"],[\"T\",\"_Toc348352859\",\"110\"],[\"T\",\"_Toc348352860\",\"111\"],[\"T\",\"_Toc348352861\",\"112\"],[\"T\",\"_Toc348352862\",\"113\"],[\"T\",\"_Toc348352863\",\"114\"],[\"T\",\"_Toc348352864\",\"115\"],[\"T\",\"_Toc348352865\",\"116\"],[\"T\",\"_Toc348352866\",\"117\"],[\"T\",\"_Toc348352867\",\"118\"],[\"T\",\"_Toc348352868\",\"119\"],[\"T\",\"_Toc348352869\",\"120\"],[\"T\",\"_Toc348352870\",\"121\"],[\"T\",\"_Toc348352871\",\"122\"],[\"T\",\"_Toc348352872\",\"123\"],[\"T\",\"_Toc348352873\",\"124\"],[\"T\",\"_Toc348352874\",\"125\"],[\"T\",\"_Toc348352875\",\"126\"],[\"T\",\"_Toc348352876\",\"127\"],[\"T\",\"_Toc348352877\",\"128\"],[\"T\",\"_Toc348352878\",\"129\"],[\"T\",\"_Toc348352879\",\"130\"],[\"T\",\"_Toc348352880\",\"131\"],[\"T\",\"_Toc348352881\",\"132\"],[\"T\",\"_Toc348352882\",\"133\"],[\"T\",\"_Toc348352883\",\"134\"],[\"T\",\"_Toc348352884\",\"135\"],[\"T\",\"_Toc348352885\",\"136\"],[\"T\",\"_Toc348352886\",\"137\"],[\"T\",\"_Toc348352887\",\"138\"],[\"T\",\"_Toc348352888\",\"139\"],[\"T\",\"_Toc348352889\",\"140\"],[\"T\",\"_Toc348352890\",\"141\"],[\"T\",\"_Toc348352891\",\"142\"],[\"T\",\"_Toc348352892\",\"143\"],[\"T\",\"_Toc348352893\",\"144\"],[\"T\",\"_Toc348352894\",\"145\"],[\"T\",\"_Toc348352895\",\"146\"],[\"T\",\"_Toc348352896\",\"147\"],[\"T\",\"_Toc348352897\",\"148\"],[\"T\",\"_Toc348352898\",\"149\"],[\"T\",\"_Toc348352899\",\"150\"],[\"T\",\"_Toc348352900\",\"151\"],[\"T\",\"_Toc348352901\",\"152\"],[\"T\",\"_Toc348352902\",\"153\"],[\"T\",\"_Toc348352903\",\"154\"],[\"T\",\"_Toc348352904\",\"155\"],[\"T\",\"_Toc348352905\",\"156\"],[\"T\",\"_Toc348352906\",\"157\"],[\"T\",\"_Toc348352907\",\"158\"],[\"T\",\"_Toc348352908\",\"159\"],[\"T\",\"_Toc348352909\",\"160\"],[\"T\",\"_Toc348352910\",\"161\"],[\"T\",\"_Toc348352911\",\"162\"],[\"T\",\"_Toc348352912\",\"163\"],[\"T\",\"_Toc348352913\",\"164\"],[\"T\",\"_Toc348352914\",\"165\"],[\"T\",\"_Toc348352915\",\"166\"],[\"T\",\"_Toc348352916\",\"167\"],[\"T\",\"_Toc348352917\",\"168\"],[\"T\",\"_Toc348352918\",\"169\"]";
        //String elems = "[\"T\",\"_Toc450152358\",\"1\"],[\"T\",\"_Toc450152355\",\"2\"],[\"T\",\"_Toc450152356\",\"3\"],[\"T\",\"_Toc450152357\",\"4\"]";//[\"T\",\"_Toc450152354\",\"1\"],[\"T\",\"_Toc450152355\",\"2\"],[\"T\",\"_Toc450152356\",\"3\"],[\"T\",\"_Toc450152357\",\"4\"]]";
        //air and missile
        //String elems = "[\"T\",\"_Toc450152354\",\"1\"],[\"T\",\"_Toc450152355\",\"2\"],[\"T\",\"_Toc450152356\",\"3\"],[\"T\",\"_Toc450152357\",\"4\"],[\"T\",\"_Toc450152358\",\"5\"],[\"T\",\"_Toc450152359\",\"6\"],[\"T\",\"_Toc450152360\",\"7\"],[\"T\",\"_Toc450152361\",\"8\"],[\"T\",\"_Toc450152362\",\"9\"],[\"T\",\"_Toc450152363\",\"10\"],[\"T\",\"_Toc450152364\",\"11\"],[\"T\",\"_Toc450152365\",\"12\"],[\"T\",\"_Toc450152366\",\"13\"],[\"T\",\"_Toc450152367\",\"14\"],[\"T\",\"_Toc450152368\",\"15\"],[\"T\",\"_Toc450152369\",\"16\"],[\"T\",\"_Toc450152370\",\"17\"],[\"T\",\"_Toc450152371\",\"18\"],[\"T\",\"_Toc450152372\",\"19\"],[\"T\",\"_Toc450152373\",\"20\"],[\"T\",\"_Toc450152374\",\"21\"],[\"T\",\"_Toc450152375\",\"22\"],[\"T\",\"_Toc450152376\",\"23\"],[\"T\",\"_Toc450152377\",\"24\"],[\"T\",\"_Toc450152378\",\"25\"],[\"T\",\"_Toc450152379\",\"26\"],[\"T\",\"_Toc450152380\",\"27\"],[\"T\",\"_Toc450152381\",\"28\"],[\"T\",\"_Toc450152382\",\"29\"],[\"T\",\"_Toc450152383\",\"30\"],[\"T\",\"_Toc450152384\",\"31\"],[\"T\",\"_Toc450152385\",\"32\"],[\"T\",\"_Toc450152386\",\"33\"],[\"T\",\"_Toc450152387\",\"34\"],[\"T\",\"_Toc450152388\",\"35\"],[\"T\",\"_Toc450152389\",\"36\"],[\"T\",\"_Toc450152390\",\"37\"],[\"T\",\"_Toc450152391\",\"38\"],[\"T\",\"_Toc450152392\",\"39\"],[\"T\",\"_Toc450152393\",\"40\"],[\"T\",\"_Toc450152394\",\"41\"],[\"T\",\"_Toc450152395\",\"42\"],[\"T\",\"_Toc450152396\",\"43\"],[\"T\",\"_Toc450152397\",\"44\"],[\"T\",\"_Toc450152398\",\"45\"],[\"T\",\"_Toc450152399\",\"46\"],[\"T\",\"_Toc450152400\",\"47\"],[\"T\",\"_Toc450152401\",\"48\"],[\"T\",\"_Toc450152402\",\"49\"],[\"T\",\"_Toc450152403\",\"50\"],[\"T\",\"_Toc450152404\",\"51\"],[\"T\",\"_Toc450152405\",\"52\"],[\"T\",\"_Toc450152406\",\"53\"],[\"T\",\"_Toc450152407\",\"54\"],[\"T\",\"_Toc450152408\",\"55\"],[\"T\",\"_Toc450152409\",\"56\"],[\"T\",\"_Toc450152410\",\"57\"],[\"T\",\"_Toc450152411\",\"58\"],[\"T\",\"_Toc450152412\",\"59\"],[\"T\",\"_Toc450152413\",\"60\"],[\"T\",\"_Toc450152414\",\"61\"],[\"T\",\"_Toc450152415\",\"62\"],[\"T\",\"_Toc450152416\",\"63\"],[\"T\",\"_Toc450152417\",\"64\"],[\"T\",\"_Toc450152418\",\"65\"],[\"T\",\"_Toc450152419\",\"66\"],[\"T\",\"_Toc450152420\",\"67\"],[\"T\",\"_Toc450152421\",\"68\"],[\"T\",\"_Toc450152422\",\"69\"],[\"T\",\"_Toc450152423\",\"70\"],[\"T\",\"_Toc450152424\",\"71\"],[\"T\",\"_Toc450152425\",\"72\"],[\"T\",\"_Toc450152426\",\"73\"],[\"T\",\"_Toc450152427\",\"74\"],[\"T\",\"_Toc450152428\",\"75\"],[\"T\",\"_Toc450152429\",\"76\"],[\"T\",\"_Toc450152430\",\"77\"],[\"T\",\"_Toc450152431\",\"78\"],[\"T\",\"_Toc450152432\",\"79\"],[\"T\",\"_Toc450152433\",\"80\"],[\"T\",\"_Toc450152434\",\"81\"],[\"T\",\"_Toc450152435\",\"82\"],[\"T\",\"_Toc450152436\",\"83\"],[\"T\",\"_Toc450152437\",\"84\"],[\"T\",\"_Toc450152438\",\"85\"],[\"T\",\"_Toc450152439\",\"86\"],[\"T\",\"_Toc450152440\",\"87\"],[\"T\",\"_Toc450152441\",\"88\"],[\"T\",\"_Toc450152442\",\"89\"]";
        //agriculture enzymes
        //String elems = "[\"T\",\"_Toc386562753\",\"1\"],[\"T\",\"_Toc386562754\",\"2\"],[\"T\",\"_Toc386562755\",\"3\"],[\"T\",\"_Toc386562756\",\"4\"],[\"T\",\"_Toc386562757\",\"5\"],[\"T\",\"_Toc386562758\",\"6\"],[\"T\",\"_Toc386562759\",\"7\"],[\"T\",\"_Toc386562760\",\"8\"],[\"T\",\"_Toc386562761\",\"9\"],[\"T\",\"_Toc386562762\",\"10\"],[\"T\",\"_Toc386562763\",\"11\"],[\"T\",\"_Toc386562764\",\"12\"],[\"T\",\"_Toc386562765\",\"13\"],[\"T\",\"_Toc386562766\",\"14\"],[\"T\",\"_Toc386562767\",\"15\"],[\"T\",\"_Toc386562768\",\"16\"],[\"T\",\"_Toc386562769\",\"17\"],[\"T\",\"_Toc386562770\",\"18\"],[\"T\",\"_Toc386562771\",\"19\"],[\"T\",\"_Toc386562772\",\"20\"],[\"T\",\"_Toc386562773\",\"21\"],[\"T\",\"_Toc386562774\",\"22\"],[\"T\",\"_Toc386562775\",\"23\"],[\"T\",\"_Toc386562776\",\"24\"],[\"T\",\"_Toc386562777\",\"25\"],[\"T\",\"_Toc386562778\",\"26\"],[\"T\",\"_Toc386562779\",\"27\"],[\"T\",\"_Toc386562780\",\"28\"],[\"T\",\"_Toc386562781\",\"29\"],[\"T\",\"_Toc386562782\",\"30\"],[\"T\",\"_Toc386562783\",\"31\"],[\"T\",\"_Toc386562784\",\"32\"],[\"T\",\"_Toc386562785\",\"33\"],[\"T\",\"_Toc386562786\",\"34\"],[\"T\",\"_Toc386562787\",\"35\"],[\"T\",\"_Toc386562788\",\"36\"],[\"T\",\"_Toc386562789\",\"37\"],[\"T\",\"_Toc386562790\",\"38\"],[\"T\",\"_Toc386562791\",\"39\"],[\"T\",\"_Toc386562792\",\"40\"],[\"T\",\"_Toc386562793\",\"41\"],[\"T\",\"_Toc386562794\",\"42\"],[\"T\",\"_Toc386562795\",\"43\"],[\"T\",\"_Toc386562796\",\"44\"],[\"T\",\"_Toc386562797\",\"45\"],[\"T\",\"_Toc386562798\",\"46\"],[\"T\",\"_Toc386562799\",\"47\"],[\"T\",\"_Toc386562800\",\"48\"],[\"T\",\"_Toc386562801\",\"49\"],[\"T\",\"_Toc386562802\",\"50\"],[\"T\",\"_Toc386562803\",\"51\"],[\"T\",\"_Toc386562804\",\"52\"],[\"T\",\"_Toc386562805\",\"53\"],[\"T\",\"_Toc386562806\",\"54\"],[\"T\",\"_Toc386562807\",\"55\"],[\"T\",\"_Toc386562808\",\"56\"],[\"T\",\"_Toc386562809\",\"57\"],[\"T\",\"_Toc386562810\",\"58\"],[\"T\",\"_Toc386562811\",\"59\"],[\"T\",\"_Toc386562812\",\"60\"],[\"T\",\"_Toc386562813\",\"61\"],[\"T\",\"_Toc386562814\",\"62\"],[\"T\",\"_Toc386562815\",\"63\"],[\"T\",\"_Toc386562816\",\"64\"],[\"T\",\"_Toc386562817\",\"65\"],[\"T\",\"_Toc386562818\",\"66\"],[\"T\",\"_Toc386562819\",\"67\"],[\"T\",\"_Toc386562820\",\"68\"],[\"T\",\"_Toc386562821\",\"69\"],[\"T\",\"_Toc386562822\",\"70\"],[\"T\",\"_Toc386562823\",\"71\"],[\"T\",\"_Toc386562824\",\"72\"],[\"T\",\"_Toc386562825\",\"73\"],[\"T\",\"_Toc386562826\",\"74\"],[\"T\",\"_Toc386562827\",\"75\"],[\"T\",\"_Toc386562828\",\"76\"],[\"T\",\"_Toc386562829\",\"77\"],[\"T\",\"_Toc386562830\",\"78\"],[\"T\",\"_Toc386562831\",\"79\"],[\"T\",\"_Toc386562832\",\"80\"],[\"T\",\"_Toc386562833\",\"81\"],[\"T\",\"_Toc386562834\",\"82\"],[\"T\",\"_Toc386562835\",\"83\"],[\"T\",\"_Toc386562836\",\"84\"],[\"T\",\"_Toc386562837\",\"85\"],[\"T\",\"_Toc386562838\",\"86\"],[\"T\",\"_Toc386562839\",\"87\"],[\"T\",\"_Toc386562840\",\"88\"],[\"T\",\"_Toc386562841\",\"89\"],[\"T\",\"_Toc386562842\",\"90\"],[\"T\",\"_Toc386562843\",\"91\"],[\"T\",\"_Toc386562844\",\"92\"],[\"T\",\"_Toc386562845\",\"93\"],[\"T\",\"_Toc386562846\",\"94\"],[\"T\",\"_Toc386562847\",\"95\"],[\"T\",\"_Toc386562848\",\"96\"],[\"T\",\"_Toc386562849\",\"97\"],[\"T\",\"_Toc386562850\",\"98\"],[\"T\",\"_Toc386562851\",\"99\"],[\"T\",\"_Toc386562852\",\"100\"],[\"T\",\"_Toc386562853\",\"101\"],[\"T\",\"_Toc386562854\",\"102\"],[\"T\",\"_Toc386562855\",\"103\"],[\"T\",\"_Toc386562856\",\"104\"],[\"T\",\"_Toc386562857\",\"105\"]";
        //organic electronics
        //String elems = "[\"T\",\"_Toc401326361\",\"1\"],[\"T\",\"_Toc401326362\",\"2\"],[\"T\",\"_Toc401326363\",\"3\"],[\"T\",\"_Toc401326364\",\"4\"],[\"T\",\"_Toc401326365\",\"5\"],[\"T\",\"_Toc401326366\",\"6\"],[\"T\",\"_Toc401326367\",\"7\"],[\"T\",\"_Toc401326368\",\"8\"],[\"T\",\"_Toc401326369\",\"9\"],[\"T\",\"_Toc401326370\",\"10\"],[\"T\",\"_Toc401326371\",\"11\"],[\"T\",\"_Toc401326372\",\"12\"],[\"T\",\"_Toc401326373\",\"13\"],[\"T\",\"_Toc401326374\",\"14\"],[\"T\",\"_Toc401326375\",\"15\"],[\"T\",\"_Toc401326376\",\"16\"],[\"T\",\"_Toc401326377\",\"17\"],[\"T\",\"_Toc401326378\",\"18\"],[\"T\",\"_Toc401326379\",\"19\"],[\"T\",\"_Toc401326380\",\"20\"],[\"T\",\"_Toc401326381\",\"21\"],[\"T\",\"_Toc401326382\",\"22\"],[\"T\",\"_Toc401326383\",\"23\"],[\"T\",\"_Toc401326384\",\"24\"],[\"T\",\"_Toc401326385\",\"25\"],[\"T\",\"_Toc401326386\",\"26\"],[\"T\",\"_Toc401326387\",\"27\"],[\"T\",\"_Toc401326388\",\"28\"],[\"T\",\"_Toc401326389\",\"29\"],[\"T\",\"_Toc401326390\",\"30\"],[\"T\",\"_Toc401326391\",\"31\"],[\"T\",\"_Toc401326392\",\"32\"],[\"T\",\"_Toc401326393\",\"33\"],[\"T\",\"_Toc401326394\",\"34\"],[\"T\",\"_Toc401326395\",\"35\"],[\"T\",\"_Toc401326396\",\"36\"],[\"T\",\"_Toc401326397\",\"37\"],[\"T\",\"_Toc401326398\",\"38\"],[\"T\",\"_Toc401326399\",\"39\"],[\"T\",\"_Toc401326400\",\"40\"],[\"T\",\"_Toc401326401\",\"41\"],[\"T\",\"_Toc401326402\",\"42\"],[\"T\",\"_Toc401326403\",\"43\"],[\"T\",\"_Toc401326404\",\"44\"],[\"T\",\"_Toc401326405\",\"45\"],[\"T\",\"_Toc401326406\",\"46\"],[\"T\",\"_Toc401326407\",\"47\"],[\"T\",\"_Toc401326408\",\"48\"],[\"T\",\"_Toc401326409\",\"49\"],[\"T\",\"_Toc401326410\",\"50\"],[\"T\",\"_Toc401326411\",\"51\"],[\"T\",\"_Toc401326412\",\"52\"],[\"T\",\"_Toc401326413\",\"53\"],[\"T\",\"_Toc401326414\",\"54\"],[\"T\",\"_Toc401326415\",\"55\"],[\"T\",\"_Toc401326416\",\"56\"],[\"T\",\"_Toc401326417\",\"57\"],[\"T\",\"_Toc401326418\",\"58\"],[\"T\",\"_Toc401326419\",\"59\"],[\"T\",\"_Toc401326420\",\"60\"],[\"T\",\"_Toc401326421\",\"61\"],[\"T\",\"_Toc401326422\",\"62\"],[\"T\",\"_Toc401326423\",\"63\"],[\"T\",\"_Toc401326424\",\"64\"],[\"T\",\"_Toc401326425\",\"65\"],[\"T\",\"_Toc401326426\",\"66\"],[\"T\",\"_Toc401326427\",\"67\"],[\"T\",\"_Toc401326428\",\"68\"],[\"T\",\"_Toc401326429\",\"69\"],[\"T\",\"_Toc401326430\",\"70\"],[\"T\",\"_Toc401326431\",\"71\"],[\"T\",\"_Toc401326432\",\"72\"],[\"T\",\"_Toc401326433\",\"73\"]";
        //torque sensor
        //String elems = "[\"T\",\"_Toc357182365\",\"1\"],[\"T\",\"_Toc357182366\",\"2\"],[\"T\",\"_Toc357182367\",\"3\"],[\"T\",\"_Toc357182368\",\"4\"],[\"T\",\"_Toc357182369\",\"5\"],[\"T\",\"_Toc357182370\",\"6\"],[\"T\",\"_Toc357182371\",\"7\"],[\"T\",\"_Toc357182372\",\"8\"],[\"T\",\"_Toc357182373\",\"9\"],[\"T\",\"_Toc357182374\",\"10\"],[\"T\",\"_Toc357182375\",\"11\"],[\"T\",\"_Toc357182376\",\"12\"],[\"T\",\"_Toc357182377\",\"13\"],[\"T\",\"_Toc357182378\",\"14\"],[\"T\",\"_Toc357182379\",\"15\"],[\"T\",\"_Toc357182380\",\"16\"],[\"T\",\"_Toc357182381\",\"17\"],[\"T\",\"_Toc357182382\",\"18\"],[\"T\",\"_Toc357182383\",\"19\"],[\"T\",\"_Toc357182384\",\"20\"],[\"T\",\"_Toc357182385\",\"21\"],[\"T\",\"_Toc357182386\",\"22\"],[\"T\",\"_Toc357182387\",\"23\"],[\"T\",\"_Toc357182388\",\"24\"],[\"T\",\"_Toc357182389\",\"25\"],[\"T\",\"_Toc357182390\",\"26\"],[\"T\",\"_Toc357182391\",\"27\"],[\"T\",\"_Toc357182392\",\"28\"],[\"T\",\"_Toc357182393\",\"29\"],[\"T\",\"_Toc357182394\",\"30\"],[\"T\",\"_Toc357182395\",\"31\"],[\"T\",\"_Toc357182396\",\"32\"],[\"T\",\"_Toc357182397\",\"33\"],[\"T\",\"_Toc357182398\",\"34\"],[\"T\",\"_Toc357182399\",\"35\"],[\"T\",\"_Toc357182400\",\"36\"],[\"T\",\"_Toc357182401\",\"37\"],[\"T\",\"_Toc357182402\",\"38\"],[\"T\",\"_Toc357182403\",\"39\"],[\"T\",\"_Toc357182404\",\"40\"],[\"T\",\"_Toc357182405\",\"41\"],[\"T\",\"_Toc357182406\",\"42\"],[\"T\",\"_Toc357182407\",\"43\"],[\"T\",\"_Toc357182408\",\"44\"],[\"T\",\"_Toc357182409\",\"45\"],[\"T\",\"_Toc357182410\",\"46\"],[\"T\",\"_Toc357182411\",\"47\"],[\"T\",\"_Toc357182412\",\"48\"],[\"T\",\"_Toc357182413\",\"49\"],[\"T\",\"_Toc357182414\",\"50\"],[\"T\",\"_Toc357182415\",\"51\"],[\"T\",\"_Toc357182416\",\"52\"],[\"T\",\"_Toc357182417\",\"53\"],[\"T\",\"_Toc357182418\",\"54\"],[\"T\",\"_Toc357182419\",\"55\"],[\"T\",\"_Toc357182420\",\"56\"],[\"T\",\"_Toc357182421\",\"57\"],[\"T\",\"_Toc357182422\",\"58\"],[\"T\",\"_Toc357182423\",\"59\"],[\"T\",\"_Toc357182424\",\"60\"],[\"T\",\"_Toc357182425\",\"61\"],[\"T\",\"_Toc357182426\",\"62\"],[\"T\",\"_Toc357182427\",\"63\"],[\"T\",\"_Toc357182428\",\"64\"],[\"T\",\"_Toc357182429\",\"65\"],[\"T\",\"_Toc357182430\",\"66\"],[\"T\",\"_Toc357182431\",\"67\"],[\"T\",\"_Toc357182432\",\"68\"],[\"T\",\"_Toc357182433\",\"69\"],[\"T\",\"_Toc357182434\",\"70\"],[\"T\",\"_Toc357182435\",\"71\"],[\"T\",\"_Toc357182436\",\"72\"],[\"T\",\"_Toc357182437\",\"73\"],[\"T\",\"_Toc357182438\",\"74\"],[\"T\",\"_Toc357182439\",\"75\"],[\"T\",\"_Toc357182440\",\"76\"],[\"T\",\"_Toc357182441\",\"77\"],[\"T\",\"_Toc357182442\",\"78\"],[\"T\",\"_Toc357182443\",\"79\"]";
        //casino management
        //String elems = "[\"T\",\"_Toc368936266\",\"1\"],[\"T\",\"_Toc368936267\",\"2\"],[\"T\",\"_Toc368936268\",\"3\"],[\"T\",\"_Toc368936269\",\"4\"],[\"T\",\"_Toc368936270\",\"5\"],[\"T\",\"_Toc368936271\",\"6\"],[\"T\",\"_Toc368936272\",\"7\"],[\"T\",\"_Toc368936273\",\"8\"],[\"T\",\"_Toc368936274\",\"9\"],[\"T\",\"_Toc368936275\",\"10\"],[\"T\",\"_Toc368936276\",\"11\"],[\"T\",\"_Toc368936277\",\"12\"],[\"T\",\"_Toc368936278\",\"13\"],[\"T\",\"_Toc368936279\",\"14\"],[\"T\",\"_Toc368936280\",\"15\"],[\"T\",\"_Toc368936281\",\"16\"],[\"T\",\"_Toc368936282\",\"17\"],[\"T\",\"_Toc368936283\",\"18\"],[\"T\",\"_Toc368936284\",\"19\"],[\"T\",\"_Toc368936285\",\"20\"],[\"T\",\"_Toc368936286\",\"21\"],[\"T\",\"_Toc368936287\",\"22\"],[\"T\",\"_Toc368936288\",\"23\"],[\"T\",\"_Toc368936289\",\"24\"],[\"T\",\"_Toc368936290\",\"25\"],[\"T\",\"_Toc368936291\",\"26\"],[\"T\",\"_Toc368936292\",\"27\"],[\"T\",\"_Toc368936293\",\"28\"],[\"T\",\"_Toc368936294\",\"29\"],[\"T\",\"_Toc368936295\",\"30\"],[\"T\",\"_Toc368936296\",\"31\"],[\"T\",\"_Toc368936297\",\"32\"],[\"T\",\"_Toc368936298\",\"33\"],[\"T\",\"_Toc368936299\",\"34\"],[\"T\",\"_Toc368936300\",\"35\"],[\"T\",\"_Toc368936301\",\"36\"],[\"T\",\"_Toc368936302\",\"37\"],[\"T\",\"_Toc368936303\",\"38\"],[\"T\",\"_Toc368936304\",\"39\"],[\"T\",\"_Toc368936305\",\"40\"],[\"T\",\"_Toc368936306\",\"41\"],[\"T\",\"_Toc368936307\",\"42\"],[\"T\",\"_Toc368936308\",\"43\"],[\"T\",\"_Toc368936309\",\"44\"],[\"T\",\"_Toc368936310\",\"45\"],[\"T\",\"_Toc368936311\",\"46\"],[\"T\",\"_Toc368936312\",\"47\"],[\"T\",\"_Toc368936313\",\"48\"],[\"T\",\"_Toc368936314\",\"49\"],[\"T\",\"_Toc368936315\",\"50\"],[\"T\",\"_Toc368936316\",\"51\"],[\"T\",\"_Toc368936317\",\"52\"],[\"T\",\"_Toc368936318\",\"53\"],[\"T\",\"_Toc368936319\",\"54\"],[\"T\",\"_Toc368936320\",\"55\"],[\"T\",\"_Toc368936321\",\"56\"],[\"T\",\"_Toc368936322\",\"57\"],[\"T\",\"_Toc368936323\",\"58\"],[\"T\",\"_Toc368936324\",\"59\"],[\"T\",\"_Toc368936325\",\"60\"]";
        //Data Center Networking
        //String elems = "[\"T\",\"_Toc369865435\",\"1\"]";
        //String elems = "[\"T\",\"_Toc369865402\",\"1\"],[\"T\",\"_Toc369865403\",\"2\"],[\"T\",\"_Toc369865404\",\"3\"],[\"T\",\"_Toc369865405\",\"4\"],[\"T\",\"_Toc369865406\",\"5\"],[\"T\",\"_Toc369865407\",\"6\"],[\"T\",\"_Toc369865408\",\"7\"],[\"T\",\"_Toc369865409\",\"8\"],[\"T\",\"_Toc369865410\",\"9\"],[\"T\",\"_Toc369865411\",\"10\"],[\"T\",\"_Toc369865412\",\"11\"],[\"T\",\"_Toc369865413\",\"12\"],[\"T\",\"_Toc369865414\",\"13\"],[\"T\",\"_Toc369865415\",\"14\"],[\"T\",\"_Toc369865416\",\"15\"],[\"T\",\"_Toc369865417\",\"16\"],[\"T\",\"_Toc369865418\",\"17\"],[\"T\",\"_Toc369865419\",\"18\"],[\"T\",\"_Toc369865420\",\"19\"],[\"T\",\"_Toc369865421\",\"20\"],[\"T\",\"_Toc369865422\",\"21\"],[\"T\",\"_Toc369865423\",\"22\"],[\"T\",\"_Toc369865424\",\"23\"],[\"T\",\"_Toc369865425\",\"24\"],[\"T\",\"_Toc369865426\",\"25\"],[\"T\",\"_Toc369865427\",\"26\"],[\"T\",\"_Toc369865428\",\"27\"],[\"T\",\"_Toc369865429\",\"28\"],[\"T\",\"_Toc369865430\",\"29\"],[\"T\",\"_Toc369865431\",\"30\"],[\"T\",\"_Toc369865432\",\"31\"],[\"T\",\"_Toc369865433\",\"32\"],[\"T\",\"_Toc369865434\",\"33\"],[\"T\",\"_Toc369865435\",\"34\"],[\"T\",\"_Toc369865436\",\"35\"],[\"T\",\"_Toc369865437\",\"36\"],[\"T\",\"_Toc369865438\",\"37\"],[\"T\",\"_Toc369865439\",\"38\"],[\"T\",\"_Toc369865440\",\"39\"],[\"T\",\"_Toc369865441\",\"40\"],[\"T\",\"_Toc369865442\",\"41\"],[\"T\",\"_Toc369865443\",\"42\"],[\"T\",\"_Toc369865444\",\"43\"],[\"T\",\"_Toc369865445\",\"44\"],[\"T\",\"_Toc369865446\",\"45\"],[\"T\",\"_Toc369865447\",\"46\"],[\"T\",\"_Toc369865448\",\"47\"],[\"T\",\"_Toc369865449\",\"48\"],[\"T\",\"_Toc369865450\",\"49\"],[\"T\",\"_Toc369865451\",\"50\"],[\"T\",\"_Toc369865452\",\"51\"],[\"T\",\"_Toc369865453\",\"52\"],[\"T\",\"_Toc369865454\",\"53\"],[\"T\",\"_Toc369865455\",\"54\"],[\"T\",\"_Toc369865456\",\"55\"],[\"T\",\"_Toc369865457\",\"56\"],[\"T\",\"_Toc369865458\",\"57\"],[\"T\",\"_Toc369865459\",\"58\"],[\"T\",\"_Toc369865460\",\"59\"],[\"T\",\"_Toc369865461\",\"60\"],[\"T\",\"_Toc369865462\",\"61\"],[\"T\",\"_Toc369865463\",\"62\"],[\"T\",\"_Toc369865464\",\"63\"],[\"T\",\"_Toc369865465\",\"64\"],[\"T\",\"_Toc369865466\",\"65\"],[\"T\",\"_Toc369865467\",\"66\"],[\"T\",\"_Toc369865468\",\"67\"],[\"T\",\"_Toc369865469\",\"68\"],[\"T\",\"_Toc369865470\",\"69\"],[\"T\",\"_Toc369865471\",\"70\"],[\"T\",\"_Toc369865472\",\"71\"],[\"T\",\"_Toc369865473\",\"72\"],[\"T\",\"_Toc369865474\",\"73\"],[\"T\",\"_Toc369865475\",\"74\"],[\"T\",\"_Toc369865476\",\"75\"],[\"T\",\"_Toc369865477\",\"76\"],[\"T\",\"_Toc369865478\",\"77\"],[\"T\",\"_Toc369865479\",\"78\"],[\"T\",\"_Toc369865480\",\"79\"],[\"T\",\"_Toc369865481\",\"80\"],[\"T\",\"_Toc369865482\",\"81\"],[\"T\",\"_Toc369865483\",\"82\"],[\"T\",\"_Toc369865484\",\"83\"],[\"T\",\"_Toc369865485\",\"84\"],[\"T\",\"_Toc369865486\",\"85\"],[\"T\",\"_Toc369865487\",\"86\"],[\"T\",\"_Toc369865488\",\"87\"],[\"T\",\"_Toc369865489\",\"88\"],[\"T\",\"_Toc369865490\",\"89\"],[\"T\",\"_Toc369865491\",\"90\"],[\"T\",\"_Toc369865492\",\"91\"],[\"T\",\"_Toc369865493\",\"92\"],[\"T\",\"_Toc369865494\",\"93\"],[\"T\",\"_Toc369865495\",\"94\"],[\"T\",\"_Toc369865496\",\"95\"],[\"T\",\"_Toc369865497\",\"96\"],[\"T\",\"_Toc369865498\",\"97\"],[\"T\",\"_Toc369865499\",\"98\"],[\"T\",\"_Toc369865500\",\"99\"],[\"T\",\"_Toc369865501\",\"100\"],[\"T\",\"_Toc369865502\",\"101\"],[\"T\",\"_Toc369865503\",\"102\"],[\"T\",\"_Toc369865504\",\"103\"],[\"T\",\"_Toc369865505\",\"104\"],[\"T\",\"_Toc369865506\",\"105\"],[\"T\",\"_Toc369865507\",\"106\"]";
        //feed acidifiers
        //String elems = "[\"T\",\"_Toc455501745\",\"1\"],[\"T\",\"_Toc455501746\",\"2\"],[\"T\",\"_Toc455501747\",\"3\"],[\"T\",\"_Toc455501748\",\"4\"],[\"T\",\"_Toc455501749\",\"5\"],[\"T\",\"_Toc455501750\",\"6\"],[\"T\",\"_Toc455501751\",\"7\"],[\"T\",\"_Toc455501752\",\"8\"],[\"T\",\"_Toc455501753\",\"9\"],[\"T\",\"_Toc455501754\",\"10\"],[\"T\",\"_Toc455501755\",\"11\"],[\"T\",\"_Toc455501756\",\"12\"],[\"T\",\"_Toc455501757\",\"13\"],[\"T\",\"_Toc455501758\",\"14\"],[\"T\",\"_Toc455501759\",\"15\"],[\"T\",\"_Toc455501760\",\"16\"],[\"T\",\"_Toc455501761\",\"17\"],[\"T\",\"_Toc455501762\",\"18\"],[\"T\",\"_Toc455501763\",\"19\"],[\"T\",\"_Toc455501764\",\"20\"],[\"T\",\"_Toc455501765\",\"21\"],[\"T\",\"_Toc455501766\",\"22\"],[\"T\",\"_Toc455501767\",\"23\"],[\"T\",\"_Toc455501768\",\"24\"],[\"T\",\"_Toc455501769\",\"25\"],[\"T\",\"_Toc455501770\",\"26\"],[\"T\",\"_Toc455501771\",\"27\"],[\"T\",\"_Toc455501772\",\"28\"],[\"T\",\"_Toc455501773\",\"29\"],[\"T\",\"_Toc455501774\",\"30\"],[\"T\",\"_Toc455501775\",\"31\"],[\"T\",\"_Toc455501776\",\"32\"],[\"T\",\"_Toc455501777\",\"33\"],[\"T\",\"_Toc455501778\",\"34\"],[\"T\",\"_Toc455501779\",\"35\"],[\"T\",\"_Toc455501780\",\"36\"],[\"T\",\"_Toc455501781\",\"37\"],[\"T\",\"_Toc455501782\",\"38\"],[\"T\",\"_Toc455501783\",\"39\"],[\"T\",\"_Toc455501784\",\"40\"],[\"T\",\"_Toc455501785\",\"41\"],[\"T\",\"_Toc455501786\",\"42\"],[\"T\",\"_Toc455501787\",\"43\"],[\"T\",\"_Toc455501788\",\"44\"],[\"T\",\"_Toc455501789\",\"45\"],[\"T\",\"_Toc455501790\",\"46\"],[\"T\",\"_Toc455501791\",\"47\"],[\"T\",\"_Toc455501792\",\"48\"],[\"T\",\"_Toc455501793\",\"49\"],[\"T\",\"_Toc455501794\",\"50\"],[\"T\",\"_Toc455501795\",\"51\"],[\"T\",\"_Toc455501796\",\"52\"],[\"T\",\"_Toc455501797\",\"53\"],[\"T\",\"_Toc455501798\",\"54\"],[\"T\",\"_Toc455501799\",\"55\"],[\"T\",\"_Toc455501800\",\"56\"],[\"T\",\"_Toc455501801\",\"57\"],[\"T\",\"_Toc455501802\",\"58\"],[\"T\",\"_Toc455501803\",\"59\"],[\"T\",\"_Toc455501804\",\"60\"],[\"T\",\"_Toc455501805\",\"61\"],[\"T\",\"_Toc455501806\",\"62\"],[\"T\",\"_Toc455501807\",\"63\"],[\"T\",\"_Toc455501808\",\"64\"],[\"T\",\"_Toc455501809\",\"65\"],[\"T\",\"_Toc455501810\",\"66\"],[\"T\",\"_Toc455501811\",\"67\"],[\"T\",\"_Toc455501812\",\"68\"],[\"T\",\"_Toc455501813\",\"69\"],[\"T\",\"_Toc455501814\",\"70\"],[\"T\",\"_Toc455501815\",\"71\"],[\"T\",\"_Toc455501816\",\"72\"],[\"T\",\"_Toc455501817\",\"73\"],[\"T\",\"_Toc455501818\",\"74\"],[\"T\",\"_Toc455501819\",\"75\"],[\"T\",\"_Toc455501820\",\"76\"],[\"T\",\"_Toc455501821\",\"77\"],[\"T\",\"_Toc455501822\",\"78\"],[\"T\",\"_Toc455501823\",\"79\"],[\"T\",\"_Toc455501824\",\"80\"],[\"T\",\"_Toc455501825\",\"81\"],[\"T\",\"_Toc455501826\",\"82\"],[\"T\",\"_Toc455501827\",\"83\"],[\"T\",\"_Toc455501828\",\"84\"],[\"T\",\"_Toc455501829\",\"85\"],[\"T\",\"_Toc455501830\",\"86\"],[\"T\",\"_Toc455501831\",\"87\"],[\"T\",\"_Toc455501832\",\"88\"],[\"T\",\"_Toc455501833\",\"89\"],[\"T\",\"_Toc455501834\",\"90\"],[\"T\",\"_Toc455501835\",\"91\"],[\"T\",\"_Toc455501836\",\"92\"],[\"T\",\"_Toc455501837\",\"93\"],[\"T\",\"_Toc455501838\",\"94\"],[\"T\",\"_Toc455501839\",\"95\"],[\"T\",\"_Toc455501840\",\"96\"],[\"T\",\"_Toc455501841\",\"97\"],[\"T\",\"_Toc455501842\",\"98\"],[\"T\",\"_Toc455501843\",\"99\"],[\"T\",\"_Toc455501844\",\"100\"],[\"T\",\"_Toc455501845\",\"101\"],[\"T\",\"_Toc455501846\",\"102\"],[\"T\",\"_Toc455501847\",\"103\"],[\"T\",\"_Toc455501848\",\"104\"],[\"T\",\"_Toc455501849\",\"105\"],[\"T\",\"_Toc455501850\",\"106\"],[\"T\",\"_Toc455501851\",\"107\"],[\"T\",\"_Toc455501852\",\"108\"],[\"T\",\"_Toc455501853\",\"109\"],[\"T\",\"_Toc455501854\",\"110\"],[\"T\",\"_Toc455501855\",\"111\"],[\"T\",\"_Toc455501856\",\"112\"],[\"T\",\"_Toc455501857\",\"113\"],[\"T\",\"_Toc455501858\",\"114\"],[\"T\",\"_Toc455501859\",\"115\"],[\"T\",\"_Toc455501860\",\"116\"],[\"T\",\"_Toc455501861\",\"117\"],[\"T\",\"_Toc455501862\",\"118\"]";
        //rolling stock
        //String elems = "[\"T\",\"_Toc457837799\",\"1\"],[\"T\",\"_Toc457837800\",\"2\"],[\"T\",\"_Toc457837801\",\"3\"],[\"T\",\"_Toc457837802\",\"4\"],[\"T\",\"_Toc457837803\",\"5\"],[\"T\",\"_Toc457837804\",\"6\"],[\"T\",\"_Toc457837805\",\"7\"],[\"T\",\"_Toc457837806\",\"8\"],[\"T\",\"_Toc457837807\",\"9\"],[\"T\",\"_Toc457837808\",\"10\"],[\"T\",\"_Toc457837809\",\"11\"],[\"T\",\"_Toc457837810\",\"12\"],[\"T\",\"_Toc457837811\",\"13\"],[\"T\",\"_Toc457837812\",\"14\"],[\"T\",\"_Toc457837813\",\"15\"],[\"T\",\"_Toc457837814\",\"16\"],[\"T\",\"_Toc457837815\",\"17\"],[\"T\",\"_Toc457837816\",\"18\"],[\"T\",\"_Toc457837817\",\"19\"],[\"T\",\"_Toc457837818\",\"20\"],[\"T\",\"_Toc457837819\",\"21\"],[\"T\",\"_Toc457837820\",\"22\"],[\"T\",\"_Toc457837821\",\"23\"],[\"T\",\"_Toc457837822\",\"24\"],[\"T\",\"_Toc457837823\",\"25\"],[\"T\",\"_Toc457837824\",\"26\"],[\"T\",\"_Toc457837825\",\"27\"],[\"T\",\"_Toc457837826\",\"28\"],[\"T\",\"_Toc457837827\",\"29\"],[\"T\",\"_Toc457837828\",\"30\"],[\"T\",\"_Toc457837829\",\"31\"],[\"T\",\"_Toc457837830\",\"32\"],[\"T\",\"_Toc457837831\",\"33\"],[\"T\",\"_Toc457837832\",\"34\"],[\"T\",\"_Toc457837833\",\"35\"],[\"T\",\"_Toc457837834\",\"36\"],[\"T\",\"_Toc457837835\",\"37\"],[\"T\",\"_Toc457837836\",\"38\"],[\"T\",\"_Toc457837837\",\"39\"],[\"T\",\"_Toc457837838\",\"40\"],[\"T\",\"_Toc457837839\",\"41\"],[\"T\",\"_Toc457837840\",\"42\"],[\"T\",\"_Toc457837841\",\"43\"],[\"T\",\"_Toc457837842\",\"44\"],[\"T\",\"_Toc457837843\",\"45\"],[\"T\",\"_Toc457837844\",\"46\"],[\"T\",\"_Toc457837845\",\"47\"],[\"T\",\"_Toc457837846\",\"48\"],[\"T\",\"_Toc457837847\",\"49\"],[\"T\",\"_Toc457837848\",\"50\"],[\"T\",\"_Toc457837849\",\"51\"],[\"T\",\"_Toc457837850\",\"52\"],[\"T\",\"_Toc457837851\",\"53\"],[\"T\",\"_Toc457837852\",\"54\"],[\"T\",\"_Toc457837853\",\"55\"],[\"T\",\"_Toc457837854\",\"56\"],[\"T\",\"_Toc457837855\",\"57\"],[\"T\",\"_Toc457837856\",\"58\"],[\"T\",\"_Toc457837857\",\"59\"],[\"T\",\"_Toc457837858\",\"60\"],[\"T\",\"_Toc457837859\",\"61\"],[\"T\",\"_Toc457837860\",\"62\"],[\"T\",\"_Toc457837861\",\"63\"],[\"T\",\"_Toc457837862\",\"64\"],[\"T\",\"_Toc457837863\",\"65\"],[\"T\",\"_Toc457837864\",\"66\"],[\"T\",\"_Toc457837865\",\"67\"],[\"T\",\"_Toc457837866\",\"68\"],[\"T\",\"_Toc457837867\",\"69\"],[\"T\",\"_Toc457837868\",\"70\"],[\"T\",\"_Toc457837869\",\"71\"],[\"T\",\"_Toc457837870\",\"72\"],[\"T\",\"_Toc457837871\",\"73\"],[\"T\",\"_Toc457837872\",\"74\"],[\"T\",\"_Toc457837873\",\"75\"]";
        //fire resistant glass
        //String elems = "[\"T\",\"_Toc457834214\",\"1\"],[\"T\",\"_Toc457834215\",\"2\"],[\"T\",\"_Toc457834216\",\"3\"],[\"T\",\"_Toc457834217\",\"5\"],[\"T\",\"_Toc457834218\",\"6\"],[\"T\",\"_Toc457834219\",\"7\"],[\"T\",\"_Toc457834220\",\"8\"],[\"T\",\"_Toc457834221\",\"9\"],[\"T\",\"_Toc457834222\",\"10\"],[\"T\",\"_Toc457834223\",\"11\"],[\"T\",\"_Toc457834224\",\"12\"],[\"T\",\"_Toc457834225\",\"13\"],[\"T\",\"_Toc457834226\",\"15\"],[\"T\",\"_Toc457834227\",\"16\"],[\"T\",\"_Toc457834228\",\"17\"],[\"T\",\"_Toc457834229\",\"19\"],[\"T\",\"_Toc457834230\",\"21\"],[\"T\",\"_Toc457834231\",\"23\"],[\"T\",\"_Toc457834232\",\"24\"],[\"T\",\"_Toc457834233\",\"25\"],[\"T\",\"_Toc457834234\",\"26\"],[\"T\",\"_Toc457834235\",\"27\"],[\"T\",\"_Toc457834236\",\"28\"],[\"T\",\"_Toc457834237\",\"29\"],[\"T\",\"_Toc457834238\",\"30\"],[\"T\",\"_Toc457834239\",\"31\"],[\"T\",\"_Toc457834240\",\"32\"],[\"T\",\"_Toc457834241\",\"33\"],[\"T\",\"_Toc457834242\",\"34\"],[\"T\",\"_Toc457834243\",\"35\"],[\"T\",\"_Toc457834244\",\"36\"],[\"T\",\"_Toc457834245\",\"37\"],[\"T\",\"_Toc457834246\",\"38\"],[\"T\",\"_Toc457834247\",\"39\"],[\"T\",\"_Toc457834248\",\"40\"],[\"T\",\"_Toc457834249\",\"41\"],[\"T\",\"_Toc457834250\",\"42\"],[\"T\",\"_Toc457834251\",\"43\"],[\"T\",\"_Toc457834252\",\"44\"],[\"T\",\"_Toc457834253\",\"45\"],[\"T\",\"_Toc457834254\",\"46\"],[\"T\",\"_Toc457834255\",\"47\"],[\"T\",\"_Toc457834256\",\"48\"],[\"T\",\"_Toc457834257\",\"49\"],[\"T\",\"_Toc457834258\",\"51\"],[\"T\",\"_Toc457834259\",\"52\"],[\"T\",\"_Toc457834260\",\"53\"],[\"T\",\"_Toc457834261\",\"54\"],[\"T\",\"_Toc457834262\",\"55\"],[\"T\",\"_Toc457834263\",\"56\"],[\"T\",\"_Toc457834264\",\"57\"],[\"T\",\"_Toc457834265\",\"58\"],[\"T\",\"_Toc457834266\",\"59\"],[\"T\",\"_Toc457834267\",\"60\"],[\"T\",\"_Toc457834268\",\"61\"],[\"T\",\"_Toc457834269\",\"62\"],[\"T\",\"_Toc457834270\",\"63\"],[\"T\",\"_Toc457834271\",\"64\"],[\"T\",\"_Toc457834272\",\"65\"],[\"T\",\"_Toc457834273\",\"66\"],[\"T\",\"_Toc457834274\",\"67\"],[\"T\",\"_Toc457834275\",\"68\"],[\"T\",\"_Toc457834276\",\"69\"],[\"T\",\"_Toc457834277\",\"70\"],[\"T\",\"_Toc457834278\",\"71\"],[\"T\",\"_Toc457834279\",\"72\"],[\"T\",\"_Toc457834280\",\"73\"],[\"T\",\"_Toc457834281\",\"74\"],[\"T\",\"_Toc457834282\",\"75\"],[\"T\",\"_Toc457834283\",\"76\"],[\"T\",\"_Toc457834284\",\"77\"],[\"T\",\"_Toc457834285\",\"78\"],[\"T\",\"_Toc457834286\",\"79\"],[\"T\",\"_Toc457834287\",\"81\"],[\"T\",\"_Toc457834288\",\"82\"],[\"T\",\"_Toc457834289\",\"83\"],[\"T\",\"_Toc457834290\",\"84\"],[\"T\",\"_Toc457834291\",\"85\"],[\"T\",\"_Toc457834292\",\"86\"],[\"T\",\"_Toc457834293\",\"87\"],[\"T\",\"_Toc457834294\",\"88\"],[\"T\",\"_Toc457834295\",\"89\"],[\"T\",\"_Toc457834296\",\"90\"],[\"T\",\"_Toc457834297\",\"91\"],[\"T\",\"_Toc457834298\",\"92\"],[\"T\",\"_Toc457834299\",\"93\"],[\"T\",\"_Toc457834300\",\"94\"],[\"T\",\"_Toc457834301\",\"95\"],[\"T\",\"_Toc457834302\",\"96\"],[\"T\",\"_Toc457834303\",\"97\"],[\"T\",\"_Toc457834304\",\"98\"],[\"T\",\"_Toc457834305\",\"99\"],[\"T\",\"_Toc457834306\",\"100\"],[\"T\",\"_Toc457834307\",\"101\"],[\"T\",\"_Toc457834308\",\"102\"],[\"T\",\"_Toc457834309\",\"103\"],[\"T\",\"_Toc457834310\",\"104\"],[\"T\",\"_Toc457834311\",\"105\"],[\"T\",\"_Toc457834312\",\"106\"],[\"T\",\"_Toc457834313\",\"107\"],[\"T\",\"_Toc457834314\",\"108\"],[\"T\",\"_Toc457834315\",\"109\"],[\"T\",\"_Toc457834316\",\"110\"],[\"T\",\"_Toc457834317\",\"111\"],[\"T\",\"_Toc457834318\",\"112\"],[\"T\",\"_Toc457834319\",\"113\"],[\"T\",\"_Toc457834320\",\"114\"],[\"T\",\"_Toc457834321\",\"115\"],[\"T\",\"_Toc457834322\",\"116\"],[\"T\",\"_Toc457834323\",\"117\"],[\"T\",\"_Toc457834324\",\"118\"],[\"T\",\"_Toc457834325\",\"119\"],[\"T\",\"_Toc457834326\",\"120\"],[\"T\",\"_Toc457834327\",\"121\"],[\"T\",\"_Toc457834328\",\"122\"],[\"T\",\"_Toc457834329\",\"123\"],[\"T\",\"_Toc457834330\",\"124\"],[\"T\",\"_Toc457834331\",\"125\"],[\"T\",\"_Toc457834332\",\"126\"],[\"T\",\"_Toc457834333\",\"127\"],[\"T\",\"_Toc457834334\",\"128\"],[\"T\",\"_Toc457834335\",\"129\"],[\"T\",\"_Toc457834336\",\"130\"],[\"T\",\"_Toc457834337\",\"131\"],[\"T\",\"_Toc457834338\",\"132\"],[\"T\",\"_Toc457834339\",\"133\"],[\"T\",\"_Toc457834340\",\"134\"],[\"T\",\"_Toc457834341\",\"135\"],[\"T\",\"_Toc457834342\",\"137\"],[\"T\",\"_Toc457834343\",\"138\"],[\"T\",\"_Toc457834344\",\"139\"],[\"T\",\"_Toc457834345\",\"140\"],[\"T\",\"_Toc457834346\",\"141\"],[\"T\",\"_Toc457834347\",\"143\"],[\"T\",\"_Toc457834348\",\"144\"],[\"T\",\"_Toc457834349\",\"145\"],[\"T\",\"_Toc457834350\",\"146\"],[\"T\",\"_Toc457834351\",\"147\"],[\"T\",\"_Toc457834352\",\"148\"],[\"T\",\"_Toc457834353\",\"149\"],[\"T\",\"_Toc457834354\",\"150\"],[\"T\",\"_Toc457834355\",\"151\"],[\"T\",\"_Toc457834356\",\"152\"],[\"T\",\"_Toc457834357\",\"153\"],[\"T\",\"_Toc457834358\",\"154\"],[\"T\",\"_Toc457834359\",\"155\"],[\"T\",\"_Toc457834360\",\"156\"],[\"T\",\"_Toc457834361\",\"157\"],[\"T\",\"_Toc457834362\",\"158\"],[\"T\",\"_Toc457834363\",\"159\"],[\"T\",\"_Toc457834364\",\"160\"],[\"T\",\"_Toc457834365\",\"161\"],[\"T\",\"_Toc457834366\",\"162\"],[\"T\",\"_Toc457834367\",\"163\"],[\"T\",\"_Toc457834368\",\"164\"],[\"T\",\"_Toc457834369\",\"165\"],[\"T\",\"_Toc457834370\",\"166\"],[\"T\",\"_Toc457834371\",\"167\"],[\"T\",\"_Toc457834372\",\"168\"],[\"T\",\"_Toc457834373\",\"169\"],[\"T\",\"_Toc457834374\",\"170\"],[\"T\",\"_Toc457834375\",\"171\"],[\"T\",\"_Toc457834376\",\"172\"],[\"T\",\"_Toc457834377\",\"173\"],[\"T\",\"_Toc457834378\",\"174\"],[\"T\",\"_Toc457834379\",\"175\"],[\"T\",\"_Toc457834380\",\"176\"]";
        //Temperature Management Market
        //String elems = "[\"T\",\"_Toc438728601\",\"1\"],[\"T\",\"_Toc438728602\",\"2\"],[\"T\",\"_Toc438728603\",\"3\"],[\"T\",\"_Toc438728604\",\"4\"],[\"T\",\"_Toc438728605\",\"5\"],[\"T\",\"_Toc438728606\",\"6\"],[\"T\",\"_Toc438728607\",\"7\"],[\"T\",\"_Toc438728608\",\"8\"],[\"T\",\"_Toc438728609\",\"9\"],[\"T\",\"_Toc438728610\",\"10\"],[\"T\",\"_Toc438728611\",\"11\"],[\"T\",\"_Toc438728612\",\"12\"],[\"T\",\"_Toc438728613\",\"13\"],[\"T\",\"_Toc438728614\",\"14\"],[\"T\",\"_Toc438728615\",\"15\"],[\"T\",\"_Toc438728616\",\"16\"],[\"T\",\"_Toc438728617\",\"17\"],[\"T\",\"_Toc438728618\",\"18\"],[\"T\",\"_Toc438728619\",\"19\"],[\"T\",\"_Toc438728620\",\"20\"],[\"T\",\"_Toc438728621\",\"21\"],[\"T\",\"_Toc438728622\",\"22\"],[\"T\",\"_Toc438728623\",\"23\"],[\"T\",\"_Toc438728624\",\"24\"],[\"T\",\"_Toc438728625\",\"25\"],[\"T\",\"_Toc438728626\",\"26\"],[\"T\",\"_Toc438728627\",\"27\"],[\"T\",\"_Toc438728628\",\"28\"],[\"T\",\"_Toc438728629\",\"29\"],[\"T\",\"_Toc438728630\",\"30\"],[\"T\",\"_Toc438728631\",\"31\"],[\"T\",\"_Toc438728632\",\"32\"],[\"T\",\"_Toc438728633\",\"33\"],[\"T\",\"_Toc438728634\",\"34\"],[\"T\",\"_Toc438728635\",\"35\"],[\"T\",\"_Toc438728636\",\"36\"],[\"T\",\"_Toc438728637\",\"37\"],[\"T\",\"_Toc438728638\",\"38\"],[\"T\",\"_Toc438728639\",\"39\"],[\"T\",\"_Toc438728640\",\"40\"],[\"T\",\"_Toc438728641\",\"41\"],[\"T\",\"_Toc438728642\",\"42\"],[\"T\",\"_Toc438728643\",\"43\"],[\"T\",\"_Toc438728644\",\"44\"],[\"T\",\"_Toc438728645\",\"45\"],[\"T\",\"_Toc438728646\",\"46\"],[\"T\",\"_Toc438728647\",\"47\"],[\"T\",\"_Toc438728648\",\"48\"],[\"T\",\"_Toc438728649\",\"49\"],[\"T\",\"_Toc438728650\",\"50\"],[\"T\",\"_Toc438728651\",\"51\"],[\"T\",\"_Toc438728652\",\"52\"],[\"T\",\"_Toc438728653\",\"53\"],[\"T\",\"_Toc438728654\",\"54\"],[\"T\",\"_Toc438728655\",\"55\"],[\"T\",\"_Toc438728656\",\"56\"],[\"T\",\"_Toc438728657\",\"57\"],[\"T\",\"_Toc438728658\",\"58\"],[\"T\",\"_Toc438728659\",\"59\"],[\"T\",\"_Toc438728660\",\"60\"],[\"T\",\"_Toc438728661\",\"61\"],[\"T\",\"_Toc438728662\",\"62\"],[\"T\",\"_Toc438728663\",\"63\"],[\"T\",\"_Toc438728664\",\"64\"],[\"T\",\"_Toc438728665\",\"65\"],[\"T\",\"_Toc438728666\",\"66\"],[\"T\",\"_Toc438728667\",\"67\"],[\"T\",\"_Toc438728668\",\"68\"],[\"T\",\"_Toc438728669\",\"69\"],[\"T\",\"_Toc438728670\",\"70\"],[\"T\",\"_Toc438728671\",\"71\"],[\"T\",\"_Toc438728672\",\"72\"],[\"T\",\"_Toc438728673\",\"73\"],[\"T\",\"_Toc438728674\",\"74\"],[\"T\",\"_Toc438728675\",\"75\"],[\"T\",\"_Toc438728676\",\"76\"],[\"T\",\"_Toc438728677\",\"77\"],[\"T\",\"_Toc438728678\",\"78\"],[\"T\",\"_Toc438728679\",\"79\"],[\"T\",\"_Toc438728680\",\"80\"],[\"T\",\"_Toc438728681\",\"81\"],[\"T\",\"_Toc438728682\",\"82\"],[\"T\",\"_Toc438728683\",\"83\"],[\"T\",\"_Toc438728684\",\"84\"],[\"T\",\"_Toc438728685\",\"85\"],[\"T\",\"_Toc438728686\",\"86\"],[\"T\",\"_Toc438728687\",\"87\"],[\"T\",\"_Toc438728688\",\"88\"],[\"T\",\"_Toc438728689\",\"89\"],[\"T\",\"_Toc438728690\",\"90\"],[\"T\",\"_Toc438728691\",\"91\"],[\"T\",\"_Toc438728692\",\"92\"],[\"T\",\"_Toc438728693\",\"93\"],[\"T\",\"_Toc438728694\",\"94\"],[\"T\",\"_Toc438728695\",\"95\"],[\"T\",\"_Toc438728696\",\"96\"],[\"T\",\"_Toc438728697\",\"97\"],[\"T\",\"_Toc438728698\",\"98\"],[\"T\",\"_Toc438728699\",\"99\"],[\"T\",\"_Toc438728700\",\"100\"],[\"T\",\"_Toc438728701\",\"101\"],[\"T\",\"_Toc438728702\",\"102\"],[\"T\",\"_Toc438728703\",\"103\"],[\"T\",\"_Toc438728704\",\"104\"],[\"T\",\"_Toc438728705\",\"105\"],[\"T\",\"_Toc438728706\",\"106\"],[\"T\",\"_Toc438728707\",\"107\"],[\"T\",\"_Toc438728708\",\"108\"],[\"T\",\"_Toc438728709\",\"109\"],[\"T\",\"_Toc438728710\",\"110\"],[\"T\",\"_Toc438728711\",\"111\"],[\"T\",\"_Toc438728712\",\"112\"],[\"T\",\"_Toc438728713\",\"113\"],[\"T\",\"_Toc438728714\",\"114\"],[\"T\",\"_Toc438728715\",\"115\"],[\"T\",\"_Toc438728716\",\"116\"],[\"T\",\"_Toc438728717\",\"117\"],[\"T\",\"_Toc438728718\",\"118\"],[\"T\",\"_Toc438728719\",\"119\"],[\"T\",\"_Toc438728720\",\"120\"],[\"T\",\"_Toc438728721\",\"121\"],[\"T\",\"_Toc438728722\",\"122\"],[\"T\",\"_Toc438728723\",\"123\"],[\"T\",\"_Toc438728724\",\"124\"],[\"T\",\"_Toc438728725\",\"125\"],[\"T\",\"_Toc438728726\",\"126\"],[\"T\",\"_Toc438728727\",\"127\"],[\"T\",\"_Toc438728728\",\"128\"],[\"T\",\"_Toc438728729\",\"129\"],[\"T\",\"_Toc438728730\",\"130\"],[\"T\",\"_Toc438728731\",\"131\"],[\"T\",\"_Toc438728732\",\"132\"],[\"T\",\"_Toc438728733\",\"133\"],[\"T\",\"_Toc438728734\",\"134\"],[\"T\",\"_Toc438728735\",\"135\"],[\"T\",\"_Toc438728736\",\"136\"],[\"T\",\"_Toc438728737\",\"137\"],[\"T\",\"_Toc438728738\",\"138\"],[\"T\",\"_Toc438728739\",\"139\"],[\"T\",\"_Toc438728740\",\"140\"],[\"T\",\"_Toc438728741\",\"141\"],[\"T\",\"_Toc438728742\",\"142\"],[\"T\",\"_Toc438728743\",\"143\"],[\"T\",\"_Toc438728744\",\"144\"],[\"T\",\"_Toc438728745\",\"145\"],[\"T\",\"_Toc438728746\",\"146\"],[\"T\",\"_Toc438728747\",\"147\"],[\"T\",\"_Toc438728748\",\"148\"],[\"T\",\"_Toc438728749\",\"149\"],[\"T\",\"_Toc438728750\",\"150\"],[\"T\",\"_Toc438728751\",\"151\"],[\"T\",\"_Toc438728752\",\"152\"],[\"T\",\"_Toc438728753\",\"153\"],[\"T\",\"_Toc438728754\",\"154\"],[\"T\",\"_Toc438728755\",\"155\"],[\"T\",\"_Toc438728756\",\"156\"],[\"T\",\"_Toc438728757\",\"157\"],[\"T\",\"_Toc438728758\",\"158\"],[\"T\",\"_Toc438728759\",\"159\"],[\"T\",\"_Toc438728760\",\"160\"],[\"T\",\"_Toc438728761\",\"161\"],[\"T\",\"_Toc438728762\",\"162\"],[\"T\",\"_Toc438728763\",\"163\"],[\"T\",\"_Toc438728764\",\"164\"],[\"T\",\"_Toc438728765\",\"165\"],[\"T\",\"_Toc438728766\",\"166\"],[\"T\",\"_Toc438728767\",\"167\"],[\"T\",\"_Toc438728768\",\"168\"],[\"T\",\"_Toc438728769\",\"169\"],[\"T\",\"_Toc438728770\",\"170\"],[\"T\",\"_Toc438728771\",\"171\"],[\"T\",\"_Toc438728772\",\"172\"],[\"T\",\"_Toc438728773\",\"173\"],[\"T\",\"_Toc438728774\",\"174\"],[\"T\",\"_Toc438728775\",\"175\"],[\"T\",\"_Toc438728776\",\"176\"],[\"T\",\"_Toc438728777\",\"177\"],[\"T\",\"_Toc438728778\",\"178\"],[\"T\",\"_Toc438728779\",\"179\"],[\"T\",\"_Toc438728780\",\"180\"],[\"T\",\"_Toc438728781\",\"181\"],[\"T\",\"_Toc438728782\",\"182\"],[\"T\",\"_Toc438728783\",\"183\"],[\"T\",\"_Toc438728784\",\"184\"],[\"T\",\"_Toc438728785\",\"185\"],[\"T\",\"_Toc438728786\",\"186\"],[\"T\",\"_Toc438728787\",\"187\"],[\"T\",\"_Toc438728788\",\"188\"],[\"T\",\"_Toc438728789\",\"189\"],[\"T\",\"_Toc438728790\",\"190\"],[\"T\",\"_Toc438728791\",\"191\"],[\"T\",\"_Toc438728792\",\"192\"],[\"T\",\"_Toc438728793\",\"193\"],[\"T\",\"_Toc438728794\",\"194\"],[\"T\",\"_Toc438728795\",\"195\"],[\"T\",\"_Toc438728796\",\"196\"],[\"T\",\"_Toc438728797\",\"197\"],[\"T\",\"_Toc438728798\",\"198\"],[\"T\",\"_Toc438728799\",\"199\"],[\"T\",\"_Toc438728800\",\"200\"],[\"T\",\"_Toc438728801\",\"201\"]";
        //Human Insulin
        String elems = "[\"T\",\"_Toc439775282\",\"1\"],[\"T\",\"_Toc439775283\",\"2\"],[\"T\",\"_Toc439775284\",\"3\"],[\"T\",\"_Toc439775285\",\"4\"],[\"T\",\"_Toc439775286\",\"5\"],[\"T\",\"_Toc439775287\",\"6\"],[\"T\",\"_Toc439775288\",\"7\"],[\"T\",\"_Toc439775289\",\"8\"],[\"T\",\"_Toc439775290\",\"9\"],[\"T\",\"_Toc439775291\",\"10\"],[\"T\",\"_Toc439775292\",\"11\"],[\"T\",\"_Toc439775293\",\"12\"],[\"T\",\"_Toc439775294\",\"13\"],[\"T\",\"_Toc439775295\",\"14\"],[\"T\",\"_Toc439775296\",\"15\"],[\"T\",\"_Toc439775297\",\"16\"],[\"T\",\"_Toc439775298\",\"17\"],[\"T\",\"_Toc439775299\",\"18\"],[\"T\",\"_Toc439775300\",\"19\"],[\"T\",\"_Toc439775301\",\"20\"],[\"T\",\"_Toc439775302\",\"21\"],[\"T\",\"_Toc439775303\",\"22\"],[\"T\",\"_Toc439775304\",\"23\"],[\"T\",\"_Toc439775305\",\"24\"],[\"T\",\"_Toc439775306\",\"25\"],[\"T\",\"_Toc439775307\",\"26\"],[\"T\",\"_Toc439775308\",\"27\"],[\"T\",\"_Toc439775309\",\"28\"],[\"T\",\"_Toc439775310\",\"29\"],[\"T\",\"_Toc439775311\",\"30\"],[\"T\",\"_Toc439775312\",\"31\"],[\"T\",\"_Toc439775313\",\"32\"],[\"T\",\"_Toc439775314\",\"33\"],[\"T\",\"_Toc439775315\",\"34\"],[\"T\",\"_Toc439775316\",\"35\"],[\"T\",\"_Toc439775317\",\"36\"],[\"T\",\"_Toc439775318\",\"37\"],[\"T\",\"_Toc439775319\",\"38\"],[\"T\",\"_Toc439775320\",\"39\"],[\"T\",\"_Toc439775321\",\"40\"],[\"T\",\"_Toc439775322\",\"41\"],[\"T\",\"_Toc439775323\",\"42\"],[\"T\",\"_Toc439775324\",\"43\"],[\"T\",\"_Toc439775325\",\"44\"],[\"T\",\"_Toc439775326\",\"45\"],[\"T\",\"_Toc439775327\",\"46\"],[\"T\",\"_Toc439775328\",\"47\"],[\"T\",\"_Toc439775329\",\"48\"],[\"T\",\"_Toc439775330\",\"49\"],[\"T\",\"_Toc439775331\",\"50\"],[\"T\",\"_Toc439775332\",\"51\"],[\"T\",\"_Toc439775333\",\"52\"],[\"T\",\"_Toc439775334\",\"53\"],[\"T\",\"_Toc439775335\",\"54\"],[\"T\",\"_Toc439775336\",\"55\"],[\"T\",\"_Toc439775337\",\"56\"],[\"T\",\"_Toc439775338\",\"57\"],[\"T\",\"_Toc439775339\",\"58\"],[\"T\",\"_Toc439775340\",\"59\"],[\"T\",\"_Toc439775341\",\"60\"],[\"T\",\"_Toc439775342\",\"61\"],[\"T\",\"_Toc439775343\",\"62\"],[\"T\",\"_Toc439775344\",\"63\"],[\"T\",\"_Toc439775345\",\"64\"],[\"T\",\"_Toc439775346\",\"65\"],[\"T\",\"_Toc439775347\",\"66\"],[\"T\",\"_Toc439775348\",\"67\"],[\"T\",\"_Toc439775349\",\"68\"],[\"T\",\"_Toc439775350\",\"69\"],[\"T\",\"_Toc439775351\",\"70\"],[\"T\",\"_Toc439775352\",\"71\"],[\"T\",\"_Toc439775353\",\"72\"],[\"T\",\"_Toc439775354\",\"73\"],[\"T\",\"_Toc439775355\",\"74\"],[\"T\",\"_Toc439775356\",\"75\"],[\"T\",\"_Toc439775357\",\"76\"],[\"T\",\"_Toc439775358\",\"77\"],[\"T\",\"_Toc439775359\",\"78\"],[\"T\",\"_Toc439775360\",\"79\"],[\"T\",\"_Toc439775361\",\"80\"],[\"T\",\"_Toc439775362\",\"81\"],[\"T\",\"_Toc439775363\",\"82\"],[\"T\",\"_Toc439775364\",\"83\"],[\"T\",\"_Toc439775365\",\"84\"],[\"T\",\"_Toc439775366\",\"85\"],[\"T\",\"_Toc439775367\",\"86\"],[\"T\",\"_Toc439775368\",\"87\"],[\"T\",\"_Toc439775369\",\"88\"],[\"T\",\"_Toc439775370\",\"89\"],[\"T\",\"_Toc439775371\",\"90\"],[\"T\",\"_Toc439775372\",\"91\"],[\"T\",\"_Toc439775373\",\"92\"]";
        //String elems = "[\"T\",\"_Toc439775282\",\"1\"],[\"T\",\"_Toc439775286\",\"2\"]";
        //BFSI Security 
        //String elems = "[\"T\",\"_Toc461120388\",\"1\"],[\"T\",\"_Toc461120389\",\"2\"],[\"T\",\"_Toc461120390\",\"3\"],[\"T\",\"_Toc461120391\",\"4\"],[\"T\",\"_Toc461120392\",\"5\"],[\"T\",\"_Toc461120393\",\"6\"],[\"T\",\"_Toc461120394\",\"7\"],[\"T\",\"_Toc461120395\",\"8\"],[\"T\",\"_Toc461120396\",\"9\"],[\"T\",\"_Toc461120397\",\"10\"],[\"T\",\"_Toc461120398\",\"11\"],[\"T\",\"_Toc461120399\",\"12\"],[\"T\",\"_Toc461120400\",\"13\"],[\"T\",\"_Toc461120401\",\"14\"],[\"T\",\"_Toc461120402\",\"15\"],[\"T\",\"_Toc461120403\",\"16\"],[\"T\",\"_Toc461120404\",\"17\"],[\"T\",\"_Toc461120405\",\"18\"],[\"T\",\"_Toc461120406\",\"19\"],[\"T\",\"_Toc461120407\",\"20\"],[\"T\",\"_Toc461120408\",\"21\"],[\"T\",\"_Toc461120409\",\"22\"],[\"T\",\"_Toc461120410\",\"23\"],[\"T\",\"_Toc461120411\",\"24\"],[\"T\",\"_Toc461120412\",\"25\"],[\"T\",\"_Toc461120413\",\"26\"],[\"T\",\"_Toc461120414\",\"27\"],[\"T\",\"_Toc461120415\",\"28\"],[\"T\",\"_Toc461120416\",\"29\"],[\"T\",\"_Toc461120417\",\"30\"],[\"T\",\"_Toc461120418\",\"31\"],[\"T\",\"_Toc461120419\",\"32\"],[\"T\",\"_Toc461120420\",\"33\"],[\"T\",\"_Toc461120421\",\"34\"],[\"T\",\"_Toc461120422\",\"35\"],[\"T\",\"_Toc461120423\",\"36\"],[\"T\",\"_Toc461120424\",\"37\"],[\"T\",\"_Toc461120425\",\"38\"],[\"T\",\"_Toc461120426\",\"39\"],[\"T\",\"_Toc461120427\",\"40\"],[\"T\",\"_Toc461120428\",\"41\"],[\"T\",\"_Toc461120429\",\"42\"],[\"T\",\"_Toc461120430\",\"43\"],[\"T\",\"_Toc461120431\",\"44\"],[\"T\",\"_Toc461120432\",\"45\"],[\"T\",\"_Toc461120433\",\"46\"],[\"T\",\"_Toc461120434\",\"47\"],[\"T\",\"_Toc461120435\",\"48\"],[\"T\",\"_Toc461120436\",\"49\"],[\"T\",\"_Toc461120437\",\"50\"],[\"T\",\"_Toc461120438\",\"51\"],[\"T\",\"_Toc461120439\",\"52\"],[\"T\",\"_Toc461120440\",\"53\"],[\"T\",\"_Toc461120441\",\"54\"],[\"T\",\"_Toc461120442\",\"55\"],[\"T\",\"_Toc461120443\",\"56\"],[\"T\",\"_Toc461120444\",\"57\"],[\"T\",\"_Toc461120445\",\"58\"],[\"T\",\"_Toc461120446\",\"59\"],[\"T\",\"_Toc461120447\",\"60\"],[\"T\",\"_Toc461120448\",\"61\"],[\"T\",\"_Toc461120449\",\"62\"],[\"T\",\"_Toc461120450\",\"63\"],[\"T\",\"_Toc461120451\",\"64\"],[\"T\",\"_Toc461120452\",\"65\"],[\"T\",\"_Toc461120453\",\"66\"],[\"T\",\"_Toc461120454\",\"67\"],[\"T\",\"_Toc461120455\",\"68\"],[\"T\",\"_Toc461120456\",\"69\"],[\"T\",\"_Toc461120457\",\"70\"],[\"T\",\"_Toc461120458\",\"71\"],[\"T\",\"_Toc461120459\",\"72\"],[\"T\",\"_Toc461120460\",\"73\"],[\"T\",\"_Toc461120461\",\"74\"],[\"T\",\"_Toc461120462\",\"75\"],[\"T\",\"_Toc461120463\",\"76\"],[\"T\",\"_Toc461120464\",\"77\"],[\"T\",\"_Toc461120465\",\"78\"],[\"T\",\"_Toc461120466\",\"79\"],[\"T\",\"_Toc461120467\",\"80\"],[\"T\",\"_Toc461120468\",\"81\"],[\"T\",\"_Toc461120469\",\"82\"],[\"T\",\"_Toc461120470\",\"83\"],[\"T\",\"_Toc461120471\",\"84\"],[\"T\",\"_Toc461120472\",\"85\"],[\"T\",\"_Toc461120473\",\"86\"],[\"T\",\"_Toc461120474\",\"87\"],[\"T\",\"_Toc461120475\",\"88\"],[\"T\",\"_Toc461120476\",\"89\"]";
        //Immunotherapy Drugs Market
        //String elems = "[\"T\",\"_Toc436407972\",\"32\"]";
        //1494597741
        //String elems = "[\"T\",\"_Toc479789906\",\"1\"],[\"T\",\"_Toc479789931\",\"2\"],[\"T\",\"Toc479789937\",\"3\"],[\"T\",\"_Toc436407945\",\"4\"],[\"T\",\"_Toc436407946\",\"5\"],[\"T\",\"_Toc436407947\",\"6\"],[\"T\",\"_Toc436407948\",\"7\"],[\"T\",\"_Toc436407949\",\"8\"],[\"T\",\"_Toc436407950\",\"9\"],[\"T\",\"_Toc436407951\",\"10\"],[\"T\",\"_Toc436407952\",\"11\"],[\"T\",\"_Toc436407953\",\"12\"],[\"T\",\"_Toc436407954\",\"13\"],[\"T\",\"_Toc436407955\",\"14\"],[\"T\",\"_Toc436407956\",\"15\"],[\"T\",\"_Toc436407957\",\"16\"],[\"T\",\"_Toc436407958\",\"17\"],[\"T\",\"_Toc436407959\",\"18\"],[\"T\",\"_Toc436407960\",\"19\"],[\"T\",\"_Toc436407961\",\"20\"],[\"T\",\"_Toc436407962\",\"21\"],[\"T\",\"_Toc436407963\",\"22\"],[\"T\",\"_Toc436407964\",\"23\"],[\"T\",\"_Toc436407965\",\"24\"],[\"T\",\"_Toc436407966\",\"25\"],[\"T\",\"_Toc436407967\",\"26\"],[\"T\",\"_Toc436407968\",\"27\"],[\"T\",\"_Toc436407969\",\"28\"],[\"T\",\"_Toc436407970\",\"29\"],[\"T\",\"_Toc436407971\",\"30\"],[\"T\",\"_Toc436407972\",\"31\"],[\"T\",\"_Toc436407973\",\"32\"],[\"T\",\"_Toc436407974\",\"33\"],[\"T\",\"_Toc436407975\",\"34\"],[\"T\",\"_Toc436407976\",\"35\"],[\"T\",\"_Toc436407977\",\"36\"],[\"T\",\"_Toc436407978\",\"37\"]";
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
        }

    }

}
