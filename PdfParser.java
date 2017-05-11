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
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.pdfbox.ExtractText;
import org.apache.pdfbox.pdfparser.PDFParser;

/**
 *
 * @author vamsi.mohan
 */
public class PdfParser {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        PDFParser pdf ;
        try {
            ExtractText.main(new String[]{"C:\\Users\\vamsi.mohan\\Desktop\\india-retail-report-2646.pdf", "D:\\test.txt"});
        } catch (Exception ex) {
            Logger.getLogger(PdfParser.class.getName()).log(Level.SEVERE, null, ex);
        }
       
}
}
