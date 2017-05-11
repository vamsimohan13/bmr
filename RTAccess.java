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

import com.mongodb.MongoClient;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import static java.util.Arrays.asList;
import java.util.Calendar;
import java.util.List;
import org.bson.Document;

/**
 *
 * @author vamsi.mohan
 */
public class RTAccess {
    
    public static void main(String[] args) throws Exception {
        MongoClient client;
        MongoCollection<Document> mongoCollection;
        try {
            //BMRProperties bmrp = new BMRProperties();

            client = new MongoClient("172.31.25.65", 27017);
            //mongoCollection = client.getDatabase("proj_rt2").getCollection("markets");
        } catch (Exception e) {
            client = new MongoClient("54.165.128.223", 27017);
            mongoCollection = client.getDatabase("mnmks").getCollection("reportcontents");
        }
        //TDocument td;
        //System.out.println(mongoCollection.getNamespace());
        MongoDatabase db = client.getDatabase("proj_rt2");
        System.out.println(db.getCollection("markets").count());
    }
    
}
