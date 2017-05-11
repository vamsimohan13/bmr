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
import javax.xml.bind.JAXBElement;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.wml.P;
import org.docx4j.wml.R;

/**
 *
 * @author vamsi.mohan
 */
class FigureExporter extends TraversalUtil.CallbackImpl {

    public FigureExporter() {
    }
    public List<PFigurePair> pFigurePairList = new ArrayList<>();
    public List<String> listOfFigures = new ArrayList<>();
    org.docx4j.wml.PPr pPrListOfFiguresStyleDetector;
    private P titleP;
    private P currP;
    private P figureP;
    private List<P> footerP;
    boolean gotTitle, gotFooter, gotFigure, emptyFooter, continueFooter, Figure = false;
    String indent = "";
    List<org.docx4j.wml.CTBookmark> ctblist;

    @Override
    public List<Object> apply(Object o) {
        if (o instanceof P) {
            currP = (P) o;
            if (!gotTitle || !gotFooter || !gotFigure) {

                if (currP.getPPr() != null&&currP.getPPr().getPStyle() != null) {
                    
                        //System.out.println(currP.getPPr().getPStyle().getVal());
                        if (currP.getPPr().getPStyle().getVal().equalsIgnoreCase("FigureTitle")) {
                            //System.out.println("found a figure title!!");
                            //check if we need to persist the previous element..happens if there are consecutive figures
                            //..i.e ,<TitleP><Fig> followed immediately witha new <TiteleP> <Fig>, with no footer in between
                            if (gotTitle && gotFigure) {

                                PFigurePair pfp = new PFigurePair(titleP, figureP, footerP, ctblist);
                                pFigurePairList.add(pfp);
                                gotTitle = false;
                                gotFigure = false;
                            }
                            ctblist = new ArrayList<>();
                            for (int i = 0; i < currP.getContent().size(); i++) {
                                if (currP.getContent().get(i) instanceof javax.xml.bind.JAXBElement) {
                                    JAXBElement jaxb = (JAXBElement) currP.getContent().get(i);
                                    if (jaxb.getDeclaredType().getName().equalsIgnoreCase("org.docx4j.wml.CTBookmark")) {
                                        org.docx4j.wml.CTBookmark ctb = (org.docx4j.wml.CTBookmark) (jaxb.getValue());
                                        ctblist.add(ctb);
                                    }
                                }
                            }
                            titleP = currP;
                            gotTitle = true;
                        } else if (gotTitle && !gotFigure) {//this logic needs to be changed to check for existence of Figure in currp
                            //.out.println("found a figure !!");
                            if (hasFigure(currP)) {
                                figureP = currP;
                                gotFigure = true;
                            } else {
                                System.out.println(titleP + " does not have a figure or pict!! Fix it!!");
                            }

                        } else if (gotTitle && gotFigure) {
                            //System.out.println("found a figure footnote!!");
                            //to skip blank footers
                            if (!currP.getPPr().getPStyle().getVal().equalsIgnoreCase("Footerline1")) {

                                if (currP.getPPr().getPStyle().getVal().equalsIgnoreCase("NewFootnote")) {
                                    footerP.add(currP);
                                }
                                gotFooter = true;
                                //continueFooter = true;
                            } else {
                                continueFooter = true;
                            }
                        }
                    
                }
            }
            if (gotTitle && gotFigure && (gotFooter || emptyFooter)) {
                PFigurePair pfp = new PFigurePair(titleP, figureP, footerP, ctblist);
                //.out.println(titleP.toString());
                pFigurePairList.add(pfp);
                gotTitle = false;
                gotFigure = false;
                gotFooter = false;
                titleP = null;
                figureP = null;
                footerP = null;
                ctblist = null;

            }

        }
        return null;
    }

    @Override
    public boolean shouldTraverse(Object o) {
        return true;
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

    List<PFigurePair> getPFigurePairs() {
        return this.pFigurePairList;
    }

    private boolean hasFigure(P currP) {
        for (Object o : currP.getContent()) {
            if (o instanceof org.docx4j.wml.R) {
                R r = (R) o;
                for (Object o2 : r.getContent()) {
                    if (o2 instanceof javax.xml.bind.JAXBElement) {
                        JAXBElement jaxb = (JAXBElement) o2;
                        //.out.println("jaxb.getDeclaredType().getName() is "+jaxb.getDeclaredType().getName());
                        if (jaxb.getDeclaredType().getName().equalsIgnoreCase("org.docx4j.wml.Drawing") || jaxb.getDeclaredType().getName().equalsIgnoreCase("org.docx4j.wml.Pict")) {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    }

}
