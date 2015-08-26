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

package org.rsna.isn.utilizationreport.util;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.jackson.JacksonFactory;
import com.google.gdata.data.spreadsheet.CellEntry;
import com.google.gdata.data.spreadsheet.CellFeed;
import com.google.gdata.util.ServiceException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.Socket;
import java.net.UnknownHostException;
import java.security.GeneralSecurityException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import org.rsna.isn.util.Environment;
import org.rsna.isn.utilizationreport.GSpreadsheet;


/**
 * This app tests the configuration and connectivity of Google Docs. 
 * Verifies the following steps:
 * 1. Load properties file
 * 2. Check proxy connectivity
 * 3. Google Doc authentication
 * 4. Edit spreadsheet
 * 
 * @author Clifton Li
 * @version 3.2.0
 */

public class TestConnection 
{
        public static void main( String[] args ) throws UnknownHostException, IOException, ServiceException, ParseException, GeneralSecurityException
        {
            
            Environment.init("utilization-report");

            Properties props = new Properties();
            File confDir = Environment.getConfDir();
            File propFile = new File(confDir, "gdoc.properties");

            try
            {
                    FileInputStream in = new FileInputStream(propFile);
                    props.load(in);

                    in.close();
                    
                    System.out.println("Loading gdoc.properties...");
            }
            catch (Exception ex)
            {
                    System.out.println("gdoc.properties file not found in " + confDir.getPath());
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
                    
                    try 
                    {
                            Socket socket = new Socket(proxyHost,Integer.parseInt(proxyPort));
                            socket.close();              
                    }
                    catch (Exception ex)
                    {
                            System.out.println("Cannot connect to " + proxyHost + ":" + proxyPort + ". Please check local and RSNA firewall. " + ex);
                            return;
                    }

                    System.out.println("Connected to " + proxyHost + ":" + proxyPort);
            }
            else
            {
                    System.out.println("Proxy not set. Enable proxySet variable to true.");
                    return;
            }

                        
            GSpreadsheet gSheet = new GSpreadsheet();
            
            gSheet.getService().setOAuth2Credentials(credential);
            
            Date date = new Date();
            DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
            String testInput = "Test entry created at " + dateFormat.format(date);
                        
            try
            {
                        gSheet.setSpreadsheet("Test");
                        gSheet.setWorksheet("Sheet1");

                        CellEntry batchOperation = gSheet.createUpdateOperation(1,1,testInput);

                        CellFeed batchRequest = new CellFeed();
                        batchRequest.getEntries().add(batchOperation);

                        CellFeed batchResponse = gSheet.submitBatch(batchRequest);
                        
                        gSheet.printBatchStatus(batchResponse);
            }
            catch (Exception ex)
            {       
                    System.out.println("Error " + ex);
                    return;
            }

            System.out.println("Successfully edited test sheet with values \"" + dateFormat.format(date) + "\"");
            System.out.println("View the spreadsheet: https://docs.google.com/spreadsheets/d/1_tc_bvbDlgnQfv_WYdhXhZnt9Vum-Ci_XoZGleDzxuM/edit?usp=sharing");           
        }
}
