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
import java.util.Queue;
import java.util.Stack;
import javax.xml.bind.JAXBElement;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.P;
import org.docx4j.wml.R;

/**
 *
 * @author vamsi.mohan
 */
class FigureExporterNew extends TraversalUtil.CallbackImpl {

    public FigureExporterNew() {
    }
    public List<PFigurePair> pFigurePairList = new ArrayList<>();
    public List<String> listOfFigures = new ArrayList<>();
    org.docx4j.wml.PPr pPrListOfFiguresStyleDetector;
    private P titleP;
    private P currP;
    private P figureP;
    boolean inSequence = false;
    String indent = "";
    List<org.docx4j.wml.CTBookmark> ctblist;
    Stack currFigElemStack = new Stack();
    //Queue figElemQueue;
    List<P> footnoteList;

    @Override
    public List<Object> apply(Object o) {
        if (o instanceof P) {
            //could be a title, figure or footnote, all are type P, and we need a title,figure and an optional footnote to complete the figure element
            // title and footnote will be determined by styles plus prev state
            // stack will always have title then ctblist then fig then optional footnote(s)
            //a fig element traversal is complete ONLY when a new figureTitle is encountered!!!
            currP = (P) o;
            if (isFigureTitle(currP)) {

                if (!currFigElemStack.isEmpty()) {
                    empty();
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
                //currFigElemStack = new Stack();
                //push figure title but first check if stack is empty, 
                //cos every valid fig element(which is a collection of single title,single figure and footnote(s))has to be popped out and added to element list in serial order
                if (currFigElemStack.isEmpty()) {
                    currFigElemStack.push(currP);
                    currFigElemStack.push(ctblist);
                }
                inSequence = true;
            } else if (isFigure(currP)) {
                //there can only be 1 title <w:P> and 1 ctblist elem in the stack
                if (currFigElemStack.size() == 2) {
                    currFigElemStack.push(currP);
                }

                inSequence = true;
            } else if (inSequence) {
                if (isFooterLine(currP)) {
                    inSequence = true;
                } else if (isFootNote(currP)) {
                    //if its first footnote, create footnote array and push element into stack
                    if (currFigElemStack.size() == 3) {
                        footnoteList = new ArrayList<>();
                        footnoteList.add(currP);
                        currFigElemStack.push(footnoteList);
                        //currFigElemStack.push(footnoteList);
                    } else if (currFigElemStack.size() == 4) {
                        //if its not first footnote, then elem size is 4 as atleast one footnote has been push into footnotelist object which is top of stack
                    //so just pop , add and push :)
                        footnoteList = (List<P>) currFigElemStack.pop();
                        footnoteList.add(currP);
                        currFigElemStack.push(footnoteList);
                    }
                    inSequence = true;
                } else {
                    inSequence = false;
                }
            }
        }

        if(!inSequence && !currFigElemStack.isEmpty()){
            empty();
        }
        return null;
    }

    private boolean isFigureTitle(P currP) {

        if (currP.getPPr() != null && currP.getPPr().getPStyle() != null) {

            //System.out.println(currP.getPPr().getPStyle().getVal());
            if (currP.getPPr().getPStyle().getVal().equalsIgnoreCase("FigureTitle")) {
                return true;
            }
        }
        return false;
    }

    private boolean isFigure(P currP) {
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

    private boolean isFootNote(P currP) {
        if (currP.getPPr() != null && currP.getPPr().getPStyle() != null) {
            if (currP.getPPr().getPStyle().getVal().equalsIgnoreCase("NewFootnote")) {
                for (int l = 0; l < currP.getContent().size(); l++) {
                    //Dont assume its always a row
                    if ((currP.getContent().get(l) instanceof org.docx4j.wml.R)) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    private boolean isFooterLine(P currP) {
        if (currP.getPPr() != null && currP.getPPr().getPStyle() != null) {
            if (currP.getPPr().getPStyle().getVal().equalsIgnoreCase("Footerline1")) {
                inSequence = true;
                return true;
            }
        }
        return false;
    }

    @Override
    public boolean shouldTraverse(Object o
    ) {
        return false;
    }

    // Depth first
    @Override
    public void walkJAXBElements(Object parent
    ) {
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
            //this.currFigElemStack
        }
        indent = indent.substring(0, indent.length() - 4);
    }

    @Override
    public List<Object> getChildren(Object o
    ) {
        return TraversalUtil.getChildrenImpl(o);
    }

    List<PFigurePair> getPFigurePairs() {
        return this.pFigurePairList;
    }

    private void empty() {
        if (currFigElemStack.size() == 3) {
            figureP = (P) currFigElemStack.pop();
            ctblist = (List<CTBookmark>) currFigElemStack.pop();
            titleP = (P) currFigElemStack.pop();

            PFigurePair pfp = new PFigurePair(titleP, figureP, null, ctblist);
            System.out.println(titleP + ":: IS Added");
            pFigurePairList.add(pfp);
        } else if (currFigElemStack.size() == 4) {
            footnoteList = (List<P>) currFigElemStack.pop();
            figureP = (P) currFigElemStack.pop();
            ctblist = (List<CTBookmark>) currFigElemStack.pop();
            titleP = (P) currFigElemStack.pop();

            PFigurePair pfp = new PFigurePair(titleP, figureP, footnoteList, ctblist);
            System.out.println(titleP + ":: IS Added");
            //.out.println(titleP.toString());
            pFigurePairList.add(pfp);

        }
    }

}
