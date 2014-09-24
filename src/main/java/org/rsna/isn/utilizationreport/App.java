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


import com.google.gdata.util.AuthenticationException;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.SQLException;
import java.util.Properties;
import java.util.Random;
import org.apache.commons.io.IOUtils;
import org.apache.log4j.Logger;
import org.rsna.isn.dao.ConfigurationDao;
import org.rsna.isn.util.Environment;
import org.rsna.isn.util.PasswordEncryption;

/**
 * This app sends utilization statistics to Google Docs. 
 * 
 * @author Clifton Li
 * @version 3.2.0
 */
public class App 
{
    private static final Logger logger = Logger.getLogger(App.class);
    
    public static void main( String[] args ) throws FileNotFoundException, IOException, InterruptedException, SQLException
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


            if (propFile.exists())
            {
                    FileInputStream in = new FileInputStream(propFile);
                    props.load(in);

                    in.close();
                    
                    logger.info("Loaded properties from " + propFile.getPath());
            }
            else
            {
                    InputStream in = App.class.getResourceAsStream("gdoc.properties");

                    byte buffer[] = IOUtils.toByteArray(in);
                    in.close();

                    props.load(new ByteArrayInputStream(buffer));

                    FileOutputStream fos = new FileOutputStream(propFile);
                    fos.write(buffer);
                    fos.close();
                    
                    logger.info("Created properties file: " + propFile.getPath());
            }

            
            String gDocUser = props.getProperty("gdoc.username");
            String gDocPass = props.getProperty("gdoc.password");
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
            
            try
            {
                     gSheet.login(gDocUser,PasswordEncryption.decrypt(gDocPass));
            }
            catch (AuthenticationException ex) 
            {
                    logger.error("Unable to authenticate to Google", ex);
                    return;
            }

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
