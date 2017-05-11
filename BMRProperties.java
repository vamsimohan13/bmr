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
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;
/**
 *
 * @author vamsi.mohan
 */
public class BMRProperties {

 

 
    private Properties prop = null;
     
    public BMRProperties() throws IOException{
         
        InputStream is = null;
        this.prop = new Properties();
        is = this.getClass().getResourceAsStream("MnM.properties");
            prop.load(is);

    }
     
    public String getPropertyValue(String key){
        return prop.getProperty(key);
    }
     
    public static void main(String a[]) throws IOException{
         
        BMRProperties bmr = new BMRProperties();
        System.out.println("db.host: "+bmr.getPropertyValue("db_host"));
        System.out.println("db.port: "+bmr.getPropertyValue("db_port"));
        System.out.println("db.name: "+bmr.getPropertyValue("db_name"));
        
        //System.out.println("db.password: "+bmr.getPropertyValue("db.password"));
    }
}
