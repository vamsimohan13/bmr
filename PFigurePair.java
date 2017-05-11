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
import org.docx4j.wml.CTBookmark;
import org.docx4j.wml.P;

/**
 *
 * @author vamsi.mohan
 */
class PFigurePair {

    PFigurePair(P p, P figure, List<P> footer, List<CTBookmark> ctblist) {
        this.title = p;
        this.figure = figure;
        this.ctblist = ctblist;
        this.footer = footer;
    }
    P title;
    P figure;
    List<P> footer;
    List<CTBookmark> ctblist;
    private String index;
    
    public void setIndex(String index){
        this.index=index;
    }
    
    public String getIndex(){
        return this.index;
    }
}
