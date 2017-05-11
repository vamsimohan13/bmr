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

import java.util.List;
import java.util.regex.Pattern;
import javax.xml.bind.JAXBElement;
import org.docx4j.vml.CTShape;
import org.docx4j.wml.Hdr;
import org.docx4j.wml.P;
import org.docx4j.wml.Pict;
import org.docx4j.wml.R;

/**
 *
 * @author vamsi.mohan
 */
//1 0
//2 0.00
//3 #,##0
//4 #,##0.00
//5 $#,##0_);($#,##0)
//6 $#,##0_);[Red]($#,##0)
//7 $#,##0.00_);($#,##0.00)
//8 $#,##0.00_);[Red]($#,##0.00)
//9 0%
//10 0.00%
//11 0.00E+00
//12 # ?/?
//13 # ??/??
//14 m/d/yyyy
//15 d-mmm-yy
//16 d-mmm
//17 mmm-yy
//18 h:mm AM/PM
//19 h:mm:ss AM/PM
//20 h:mm
//21 h:mm:ss
//22 m/d/yyyy h:mm
//37 #,##0_);(#,##0)
//38 #,##0_);[Red](#,##0)
//39 #,##0.00_);(#,##0.00)
//40 #,##0.00_);[Red](#,##0.00)
//45 mm:ss
//46 [h]:mm:ss
//47 mm:ss.0
//48 ##0.0E+0
//49 @
final class Utils {

    static enum NumericType {

        Number, Percent, Year, SerialNo, NaN;
    }

    /**
     * Utility class.
     */
    private Utils() {
    }

    /**
     * Returns the length in characters of the leading white space in the given
     * char sequence.
     *
     * @param s the char sequence to look at.
     * @return the number of whitespace characters at the beginning of the
     * sequence..
     */
    public static NumericType getNumericFormat(CharSequence s) {
        if (0 == s.length()) {
            return NumericType.NaN;
        }
        Pattern percentpattern = Pattern.compile("[0-9.]+[%]");
        Pattern numberpattern = Pattern.compile("[0-9.,]+");


        if (numberpattern.matcher(s).matches()) {
            //System.out.println("isnumber pattern" + s);
            if (s.toString().split("[.]").length == 3) {
                return NumericType.NaN;
            }
            return NumericType.Number;
        } else if (percentpattern.matcher(s).matches()) {
            //System.out.println("isdecimal pattern" + s);

            return NumericType.Percent;

        }

        return NumericType.NaN;
    }
    
    
    public static String getReportTitle(List<Hdr> hdrlist) {
        String tryheader = null;
        if (hdrlist != null) {

            List<Object> content;
            for (Hdr hdr : hdrlist) {
                content = hdr.getContent();
                for (Object o : content) {
                    if (o instanceof P) {
                        P headerP = (P) o;
                        content = headerP.getContent();

                        for (Object r : content) {
                            if (r instanceof R) {
                                R headR = (R) r;
                                for (int m = 0; m < headR.getContent().size(); m++) {
                                    Object heado = headR.getContent().get(m);
                                    javax.xml.bind.JAXBElement jaxb = (javax.xml.bind.JAXBElement) (heado);
                                    switch (jaxb.getDeclaredType().getName()) {
                                        //// also check if images or drwings are the
                                        case "org.docx4j.wml.Pict":
                                            Pict p = ((org.docx4j.wml.Pict) (jaxb.getValue()));
                                            List<Object> anyobjs = p.getAnyAndAny();
                                            for (Object opict : anyobjs) {
                                                if (opict instanceof JAXBElement) {
                                                    if (((JAXBElement) opict).getDeclaredType().getName().equals("org.docx4j.vml.CTShape")) {
                                                        CTShape ctshape = (CTShape) ((JAXBElement) opict).getValue();
                                                        //Objects.requireNonNull(ctshape,StringBuilder.getTokenisedItem().getBuiltWith()));

                                                        //CTShape ctshape = (CTShape) opict;
                                                        List<JAXBElement<?>> elems = ctshape.getEGShapeElements();
                                                        for (JAXBElement jaxbshape : elems) {
                                                            if (jaxbshape.getDeclaredType().getName().equals("org.docx4j.vml.CTTextbox")) {
                                                                org.docx4j.vml.CTTextbox text = (org.docx4j.vml.CTTextbox) jaxbshape.getValue();
                                                                List<Object> headertext = text.getTxbxContent().getContent();
                                                                for (Object op : headertext) {
                                                                    if (op instanceof P) {
                                                                        //yay !!! got report name
                                                                        tryheader = ((P) op).toString();
                                                                        if (!tryheader.isEmpty()) {
                                                                            break;
                                                                        }
                                                                    }
                                                                }

                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            break;
                                        //header = header + r.toString();
                                    }
                                }
                            }

                        }
                        //hdr.toString();
                    }
                }
            }
        }
        return tryheader;
    }
}
