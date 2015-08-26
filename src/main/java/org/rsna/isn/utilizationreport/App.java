/* Copyright (c) <2014>, <Radiological Society of North America>
 * All rights reserved.
 * Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
 * Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
 * Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
 * Neither the name of the <RSNA> nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND
 * CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED
 * WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A
 * PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
 * THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
 * PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF
 * USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
 * CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
 * NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE
 * USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY
 * OF SUCH DAMAGE.
 */

package org.rsna.isn.utilizationreport;

//import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.jackson.JacksonFactory;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.sql.SQLException;
import java.util.Arrays;
import java.util.List;
import java.util.Properties;
import java.util.Random;
import org.apache.log4j.Logger;
import org.rsna.isn.dao.ConfigurationDao;
import org.rsna.isn.util.Environment;

/**
 * This app sends utilization statistics to Google Docs. 
 * 
 * @author Clifton Li
 * @version 3.2.0
 */
public class App 
{
    private static final Logger logger = Logger.getLogger(App.class);
    
    public static void main( String[] args ) throws InterruptedException, SQLException,GeneralSecurityException, IOException
    {
            Environment.init("utilization-report");
        
            ConfigurationDao config = new ConfigurationDao();
            
            //Configured to not sent data
            if (!Boolean.parseBoolean(config.getConfiguration("submit-stats")))
            {
                    return;
            }  
        
            //Setting random delay to avoid simultaneous read/write to google spreadsheets
            Random rand = new Random(); 
            int sleepTime = rand.nextInt(30000); 

            logger.info("Delay set to " + sleepTime + "ms");
            Thread.sleep(sleepTime);
            

            //Load peoperties file
            Properties props = new Properties();
            File confDir = Environment.getConfDir();
            File propFile = new File(confDir, "gdoc.properties");

            
            try
            {
                    FileInputStream in = new FileInputStream(propFile);
                    props.load(in);

                    in.close();
                    
                    logger.info("Loaded properties from " + propFile.getPath());
            }
            catch (Exception ex)
            {
                    logger.info("gdoc.properties file not found in " + confDir.getPath());
                    return;
            }
            
                    
            File p12 = new File(confDir,"gdoc-key.p12");
            
            if (!p12.exists())
            {
                    System.out.println("gdoc-key.p12 file not found in " + confDir.getPath());
                    return;
            }
            
            HttpTransport httpTransport = new NetHttpTransport();
            JacksonFactory jsonFactory = new JacksonFactory();
        
            String[] SCOPESArray = {"https://spreadsheets.google.com/feeds", "https://spreadsheets.google.com/feeds/spreadsheets/private/full", "https://docs.google.com/feeds"};
            final List SCOPES = Arrays.asList(SCOPESArray);
        
            GoogleCredential credential = new GoogleCredential.Builder()
                    .setTransport(httpTransport)
                    .setJsonFactory(jsonFactory)
                    .setServiceAccountId(props.getProperty("gdoc.clientid"))
                    .setServiceAccountScopes(SCOPES)
                    .setServiceAccountPrivateKeyFromP12File(p12)
                    .build();

           
            String gDocSpreadsheet = props.getProperty("gdoc.spreadsheet");
            String gDocWorksheet = props.getProperty("gdoc.worksheet");

            String proxySet = props.getProperty("gdoc.proxyset");
            String proxyHost = props.getProperty("gdoc.proxyhost");
            String proxyPort = props.getProperty("gdoc.proxyport");

            //google doesnt like the keystore
            //removing keystore properties for now until a better solution is found
            System.getProperties().remove("javax.net.ssl.keyStore");
            System.getProperties().remove("javax.net.ssl.keyStorePassword");
            System.getProperties().remove("javax.net.ssl.trustStore");
            System.getProperties().remove("javax.net.ssl.trustStorePassword");
            
            //set proxy
            if (Boolean.parseBoolean(proxySet))
            {
                    System.getProperties().put("proxySet", proxySet);
                    System.getProperties().put("proxyHost", proxyHost);
                    System.getProperties().put("proxyPort", proxyPort);
            }

            GSpreadsheet gSheet = new GSpreadsheet();
            gSheet.getService().setOAuth2Credentials(credential);
            
            try
            {
                    gSheet.setSpreadsheet(gDocSpreadsheet);
                    gSheet.setWorksheet(gDocWorksheet);
                    gSheet.addDateToHeader();
                    gSheet.calculateStats();    
            }
            catch (Exception ex)
            {
                    logger.error("Uncaught exception", ex);
            }
    }
}
