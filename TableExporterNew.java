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
import java.util.Stack;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.bind.JAXBElement;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Style;
import org.docx4j.wml.Tbl;

/**
 *
 * @author vamsi.mohan
 */
class TableExporterNew extends TraversalUtil.CallbackImpl {

    WordprocessingMLPackage wordMLPackageIn;
    public List<PTablePair> pairList = new ArrayList<>();
    public List<String> listOfTables = new ArrayList<>();
    org.docx4j.wml.PPr pPrListOfTableStyleDetector;
    private P titleP, currP, tabletextP;
    private Tbl tbl;
    String indent = "";
    static int count = 0;
    Stack currTableElemStack = new Stack();
    List<P> footnoteList;

    boolean inSequence, gotNewTableTitle = false;

    @Override
    public List<Object> apply(Object o) {
        gotNewTableTitle = false;
        if (o instanceof P) {
            currP = (P) o;
            if (isTableTitle(currP)) {
                if (!currTableElemStack.isEmpty()) {
                    empty();
                }
                currTableElemStack.push(currP);
                inSequence = true;
                gotNewTableTitle = true;
                //exit(0);
            }
            if (!gotNewTableTitle) {
                if (inSequence && currTableElemStack.size() >= 2) {
                    // we have got table title and table, so this currP can be a footer, check and add if so
                    if (isFooterLine(currP)) {
                        inSequence = true;
                    } else if (isFootNote(currP)) {
                        //if its first footnote, create footnote array and push element into stack
                        if (currTableElemStack.size() == 2) {
                            footnoteList = new ArrayList<>();
                            footnoteList.add(currP);
                            currTableElemStack.push(footnoteList);
                            //currFigElemStack.push(footnoteList);
                        } else if (currTableElemStack.size() == 3) {
                            //if its not first footnote, then elem size is 4 as atleast one footnote has been push into footnotelist object which is top of stack
                            //so just pop , add and push :)
                            footnoteList = (List<P>) currTableElemStack.pop();
                            footnoteList.add(currP);
                            currTableElemStack.push(footnoteList);
                        }
                        inSequence = true;
                    } else {
                        //check if tabletext is here,,ignore blank spaces
                        if (!(currP.toString().trim().length() == 0)) {
                            tabletextP = currP;
                            currTableElemStack.push(tabletextP);
                            inSequence = false;
                        } else {
                            inSequence = true;//there could be blank P's before , we need to skip those
                        }
                    }
                } else if (inSequence && currTableElemStack.size() == 1) {
                    //there was table title <P> followed by a <P>(or multiple Ps) rather than <Tbl>
                    // its still OK as  the Tbl may come next..just do nothing and continue
                }
            }
            //inSequence = true;
        }
        if (o instanceof Tbl) {
            tbl = (Tbl) o;

            if (currTableElemStack.size() == 1) {
                currTableElemStack.push(tbl);
                //gotTable = true;
            }
            //footerP = null;
            //extractMnMFormatTable(curTablePair);
            count++;
            inSequence = true;
        }
        if (!inSequence && !currTableElemStack.isEmpty()) {
            empty();
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

    public List<PTablePair> getPTablePairs() {
        return pairList;
    }

    private void empty() {
        if (currTableElemStack.size() == 1) {
            titleP = (P) currTableElemStack.pop();
            PTablePair ptp = new PTablePair(titleP, null, null, null);
            System.out.println(titleP + ":: IS Added but Table is not found!!");
            pairList.add(ptp);
        }
        if (currTableElemStack.size() == 2) {
            tbl = (Tbl) currTableElemStack.pop();
            //ctblist = (List<CTBookmark>) currTableElemStack.pop();
            titleP = (P) currTableElemStack.pop();

            PTablePair ptp = new PTablePair(titleP, tbl, null, null);
            System.out.println(titleP + ":: IS Added");
            pairList.add(ptp);
        } else if (currTableElemStack.size() == 3) {
            footnoteList = (List<P>) currTableElemStack.pop();
            tbl = (Tbl) currTableElemStack.pop();
            //ctblist = (List<CTBookmark>) currTableElemStack.pop();
            titleP = (P) currTableElemStack.pop();

            PTablePair ptp = new PTablePair(titleP, tbl, footnoteList, null);
            System.out.println(titleP + ":: IS Added");
            //.out.println(titleP.toString());
            pairList.add(ptp);

        } else if (currTableElemStack.size() == 4) {
            tabletextP = (P) currTableElemStack.pop();
            footnoteList = (List<P>) currTableElemStack.pop();
            tbl = (Tbl) currTableElemStack.pop();
            //ctblist = (List<CTBookmark>) currTableElemStack.pop();
            titleP = (P) currTableElemStack.pop();

            PTablePair ptp = new PTablePair(titleP, tbl, footnoteList, tabletextP);
            System.out.println(titleP + ":: IS Added");
            //.out.println(titleP.toString());
            pairList.add(ptp);

        }
    }

    private boolean isTableTitle(P currP) {
        if (currP.getPPr() != null) {
            if (currP.getPPr().getPStyle() != null) {
                try {
                    TraversalUtil.visit(null, inSequence, this);
                    //visit();
                    if (getStyle("Table_Title", currP.getPPr().getPStyle().getVal())) {
                        titleP = currP;
                        //System.out.println(titleP.toString());
                        //gotTableTitle = true;
                        return true;
                    }
                    //return true;
                } catch (Docx4JException ex) {
                    Logger.getLogger(TableExporterNew.class.getName()).log(Level.SEVERE, null, ex);
                }
            }

        }
        return false;
    }

    protected boolean getStyle(String stylename, String styleId) throws Docx4JException {

        for (Style s : this.wordMLPackageIn.getMainDocumentPart().getStyleDefinitionsPart().getContents().getStyle()) {
            if (stylename.equals(s.getName().getVal()) && styleId.equals(s.getStyleId())) {
                return true;
            }
        }
        return false;
    }

    private boolean isFooterLine(P currP) {
        if (currP.getPPr() != null && currP.getPPr().getPStyle() != null) {
            try {
                if (getStyle("Footerline1", currP.getPPr().getPStyle().getVal())) {

                    inSequence = true;
                    return true;
                }
            } catch (Docx4JException ex) {
                Logger.getLogger(TableExporterNew.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        return false;
    }

    private boolean isFootNote(P currP) {
        if (currP.getPPr() != null) {
            if (currP.getPPr().getPStyle() != null) {
                if (currP.getPPr().getPStyle().getVal().equalsIgnoreCase("NewFootnote") || currP.getPPr().getPStyle().getVal().equalsIgnoreCase("Footnotenew")) {
                    for (int l = 0; l < currP.getContent().size(); l++) {
                        //Dont assume its always a row
                        if ((currP.getContent().get(l) instanceof org.docx4j.wml.R)) {
                            return true;
                        }
                    }
                }
            } else {
                //this is tricky ..can be a footnote without style..found recently :(
                if (!currP.toString().isEmpty()) {
                    return true;
                }
            }
        }
        return false;
    }

    void setWorldMPkg(WordprocessingMLPackage wordMLPackageIn) {
        this.wordMLPackageIn = wordMLPackageIn;
        //throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private boolean isLikelyTableText(P currP) {
        System.out.println(currP);
        if (currP.getPPr() != null) {
            if (currP.getPPr().getPStyle() != null) {
                if (currP.getPPr().getPStyle().getVal().equalsIgnoreCase("NewFootnote") || currP.getPPr().getPStyle().getVal().equalsIgnoreCase("Footnotenew")) {
                    for (int l = 0; l < currP.getContent().size(); l++) {
                        //Dont assume its always a row
                        if ((currP.getContent().get(l) instanceof org.docx4j.wml.R)) {
                            return true;
                        }
                    }
                }
            } else {
                //this is tricky ..can be a footnote without style..found recently :(
                if (!currP.toString().isEmpty()) {
                    return true;
                }
            }
        }
        return false;
    }

}
