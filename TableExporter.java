/*
 * Copyright 2016 vamsi.mohan.
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

import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.wml.P;
import org.docx4j.wml.Tbl;

/**
 *
 * @author vamsi.mohan
 */
class TableExporter extends TraversalUtil.CallbackImpl {

    public List<PTablePair> pairList = new ArrayList<PTablePair>();
    public List<String> listOfTables = new ArrayList<String>();
    org.docx4j.wml.PPr pPrListOfTableStyleDetector;
    private P titleP,  currP, tabletextP;
    
    private List<P> footerP;
    private Tbl tbl;
    boolean gotTableTitle, gotTable, gotTableFooter,emptyFooter= false;
    String indent = "";
    static int count = 0;

    @Override
    public List<Object> apply(Object o) {

        if (o instanceof P) {
            currP = (P) o;
            if (currP.getPPr() != null) {
                if (currP.getPPr().getPStyle() != null) {
                    if (!gotTableTitle && currP.getPPr().getPStyle().getVal().equalsIgnoreCase("TableTitle")) {
                        titleP = currP;
                        //System.out.println(titleP.toString());
                        gotTableTitle = true;
                    } else if (gotTableTitle && gotTable) {
                        if (currP.getPPr().getPStyle().getVal().contains("NewFootnote") || currP.getPPr().getPStyle().getVal().contains("Footnotenew") ) {
                            footerP.add(currP);
                            gotTableFooter = true;
                        }else emptyFooter= true;
                        
                    }
                    //org.docx4j.wml.R r = (org.docx4j.wml.R) (currP.getContent().get(l));
                }

            }
        }
        if (o instanceof Tbl) {
            tbl = (Tbl) o;

            if(gotTableTitle) gotTable = true;
            footerP = null;
            //extractMnMFormatTable(curTablePair);
            count++;
        }
        if (gotTableTitle && gotTable && (gotTableFooter||emptyFooter)) {
            PTablePair curTablePair = new PTablePair(titleP, tbl,footerP,tabletextP);
            pairList.add(curTablePair);
            gotTableFooter = false;
            gotTableTitle = false;
            gotTable = false;
            emptyFooter = false;
            titleP = null;
            footerP = null;
            tbl = null;
        }
        return null;
    }

    @Override
    public boolean shouldTraverse(Object o) {
        return false;
    }

    // Depth first
    @Override
    public void walkJAXBElements(Object parent) {
        indent += "    ";
        List children = getChildren(parent);
        if (children != null) {
            for (Object o : children) {
                // if its wrapped in javax.xml.bind.JAXBElement, get its
                // value
                o = XmlUtils.unwrap(o);
                this.apply(o);
                if (this.shouldTraverse(o)) {
                    walkJAXBElements(o);
                }
            }
        }
        indent = indent.substring(0, indent.length() - 4);
    }

    @Override
    public List<Object> getChildren(Object o) {
        return TraversalUtil.getChildrenImpl(o);
    }

    private void extractMnMFormatTable(PTablePair currTablePair) throws Docx4JException {
        if (isMnMStyleTable(currTablePair)) {
            DocxToDocx.tblcount++;
            //remove this later;
            //if (tblcount == 1) {
            //                if (isPartOfListOfTables(currTablePair)) {
            //                    bodyOut.getContent().add(currTablePair.currP);
            //                    bodyOut.getContent().add(currTablePair.tbl);
            //                }
            //ExportTblToDocx(tbl);
            //}
        }
        //System.out.println(XmlUtils.marshaltoString(currTablePair.currP));
        //System.out.println(XmlUtils.marshaltoString(currTablePair.tbl));
        pairList.add(currTablePair);
    }

    public List<PTablePair> getPTablePairs() {
        return pairList;
    }

    private void ExportTblToDocx(Tbl tbl) {
        //TR is in content of table(wml.Tbl) which is arraylist. So cast each of them to TR's
        List<Object> tablerows = tbl.getContent();
        //org.docx4j.wml.TblGrid tablegrid = tbl.getTblGrid();
        //org.docx4j.wml.Tr header = (org.docx4j.wml.Tr) content.get(0);
        //System.out.println("Table " + tblcount + " has " + tablerows.size() + " rows(including header) and " + tbl.getTblGrid().getGridCol().size() + " columns");
        for (Object tablerow : tablerows) {
            org.docx4j.wml.Tr tr = (org.docx4j.wml.Tr) tablerow;
            //System.out.println("Row " + (i + 1));
            for (int j = 0; j < tr.getContent().size(); j++) {
                javax.xml.bind.JAXBElement trcontent = (javax.xml.bind.JAXBElement) tr.getContent().get(j);
                org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) trcontent.getValue();
                for (int k = 0; k < tc.getContent().size(); k++) {
                    org.docx4j.wml.P tcpara = (org.docx4j.wml.P) tc.getContent().get(k);
                    for (int l = 0; l < tcpara.getContent().size(); l++) {
                        //Dont assume its always a row
                        if (!(tcpara.getContent().get(l) instanceof org.docx4j.wml.R)) {
                            break;
                        }
                        org.docx4j.wml.R r = (org.docx4j.wml.R) (tcpara.getContent().get(l));
                        for (int m = 0; m < r.getContent().size(); m++) {
                            Object o = r.getContent().get(m);
                            if (o instanceof org.docx4j.wml.Br || o instanceof org.docx4j.wml.R.Tab || o instanceof org.docx4j.wml.R.LastRenderedPageBreak) {
                                break;
                            }
                            javax.xml.bind.JAXBElement jaxb = (javax.xml.bind.JAXBElement) (o);
                            if (jaxb.getDeclaredType().getName().equals("org.docx4j.wml.Text")) {
                            } else if (jaxb.getDeclaredType().getName().equals("org.docx4j.wml.Drawing")) {
                                org.docx4j.wml.Drawing drawing = (org.docx4j.wml.Drawing) (jaxb.getValue());
                                org.docx4j.dml.wordprocessingDrawing.Inline inline = (org.docx4j.dml.wordprocessingDrawing.Inline) (drawing).getAnchorOrInline().get(0);
                                List<Object> artificialList = new ArrayList<Object>();
                                //                                                CTNonVisualDrawingProps drawingProps = inline.getDocPr();
                                //                                                if (drawingProps != null) {
                                //                                                    handleCTNonVisualDrawingProps(drawingProps, artificialList);
                                //                                                }
                                if (inline.getGraphic() != null) {
                                    //log.debug("found a:graphic");
                                    org.docx4j.dml.Graphic graphic = inline.getGraphic();
                                    if (graphic.getGraphicData() != null) {
                                        String imageId = graphic.getGraphicData().getPic().getBlipFill().getBlip().getEmbed();
                                        //System.out.println("Row " + (i + 1) + " column" + (j + 1) + "'s " + (k + 1) + " value is image " + imageId + " mapped to file " + relationsPart.getRelationshipByID(imageId).getTarget());
                                    }
                                }
                            }
                            // also check if images or drwings are the
                        }
                    }
                    //                                    throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
                }
            }
        }
    }

    private boolean isMnMStyleTable(PTablePair tbl) {
        //currently assuming this table MUST have previous w:currP whose w:pPr 's w:pStyle is 'TableTitle'
        // It should also be a part of TOC's List of tables - this is also custom requirement of MnM format docs
        // just return true for now
        //System.out.println(tbl.getParent());
        return true;
    }

}
