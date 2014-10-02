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

import com.google.gdata.data.spreadsheet.CellEntry;
import com.google.gdata.data.spreadsheet.CellFeed;
import com.google.gdata.util.AuthenticationException;
import com.google.gdata.util.ServiceException;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.Socket;
import java.net.UnknownHostException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;
import org.apache.commons.io.IOUtils;
import org.rsna.isn.util.Environment;
import org.rsna.isn.utilizationreport.App;
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
        public static void main( String[] args ) throws UnknownHostException, IOException, ServiceException, ParseException
        {
            
            Environment.init("utilization-report");

            Properties props = new Properties();
            File confDir = Environment.getConfDir();
            File propFile = new File(confDir, "gdoc.properties");

            if (propFile.exists())
            {
                    FileInputStream in = new FileInputStream(propFile);
                    props.load(in);

                    in.close();
                    
                    System.out.println("Loading gdoc.properties...");

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
                    
                    System.out.println("Created properties file: " + propFile.getPath());
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
                    System.out.println("Proxy not set.");
            }

                        
            GSpreadsheet gSheet = new GSpreadsheet();
            
            boolean loginSuccess = true;
            try
            {
                     gSheet.login(gDocUser,gDocPass);
                     
            }
            catch (AuthenticationException ex) 
            {
                     System.out.println("Unable to authenticate to Google Docs" + ex);
                     loginSuccess = false;
            }
            finally
            {
                    if (loginSuccess)
                    {
                            System.out.println("Authenticated to Google Docs");
                    }
            }

            boolean writeToGdoc = true;
            
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
                    System.out.println("Error" + ex);
                    writeToGdoc = false;
            }
            finally
            {
                    System.out.println("Successfully edited test sheet with values \"" + dateFormat.format(date) + "\"");
                    System.out.println("View the spreadsheet: https://docs.google.com/spreadsheets/d/1_tc_bvbDlgnQfv_WYdhXhZnt9Vum-Ci_XoZGleDzxuM/edit?usp=sharing");
            }

        }
}
