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

import com.google.gdata.client.spreadsheet.CellQuery;
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.Link;
import com.google.gdata.data.batch.BatchOperationType;
import com.google.gdata.data.batch.BatchStatus;
import com.google.gdata.data.batch.BatchUtils;
import com.google.gdata.data.spreadsheet.Cell;
import com.google.gdata.data.spreadsheet.CellEntry;
import com.google.gdata.data.spreadsheet.CellFeed;
import com.google.gdata.data.spreadsheet.ListFeed;
import com.google.gdata.data.spreadsheet.SpreadsheetEntry;
import com.google.gdata.data.spreadsheet.SpreadsheetFeed;
import com.google.gdata.data.spreadsheet.WorksheetEntry;
import com.google.gdata.data.spreadsheet.WorksheetFeed;
import com.google.gdata.util.AuthenticationException;
import com.google.gdata.util.ServiceException;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import org.rsna.isn.dao.ConfigurationDao;
import org.apache.log4j.Logger;

/**
 * This class implements submitting utilization statistics to Google Docs .
 *
 * @author Clifton Li
 * @version 3.2.0
 * @since 3.2.0
 */
public class GSpreadsheet {
        
        private SpreadsheetService service = new SpreadsheetService("RSNA-Image-Sharing-Network-Utilization-Report");   
        private SpreadsheetEntry spreadsheet;
        private WorksheetEntry worksheet;
        private ListFeed listFeed;
        private CellFeed cellFeed;
        private URL cellFeedUrl;
        private static int DATE_HEADER = 3;
        private static int START_COLUMN = 1;  
        private CellEntry[][] cellEntries;
        SimpleDateFormat sdf = new SimpleDateFormat("M/d/yyyy");
        
        private static final Logger logger = Logger.getLogger(GSpreadsheet.class);

        public SpreadsheetService getService()
        {
                return this.service;
        }
        public void setSpreadsheet(String title) throws MalformedURLException, IOException, ServiceException
        {
                URL spreadsheetFeedURL = new URL("https://spreadsheets.google.com/feeds/spreadsheets/private/full");
                
                SpreadsheetFeed feed = service.getFeed(spreadsheetFeedURL,SpreadsheetFeed.class);
                List<SpreadsheetEntry> spreadsheets = feed.getEntries();
                
                boolean spreadsheetFound = false;
                
                for (int i = 0; i < spreadsheets.size(); i++)
                {
                        SpreadsheetEntry sheet = spreadsheets.get(i);
                        
                        if (sheet.getTitle().getPlainText().contentEquals(title))
                        {
                                this.spreadsheet = spreadsheets.get(i);
                                spreadsheetFound = true;
                                logger.info("Loaded sheet #" + i + " - " + sheet.getTitle().getPlainText());
                        }
                }
                
                if (!spreadsheetFound)
                {
                        logger.error("Can't find spreadsheet titled: " + title);
                }  
        }

        public void setWorksheet(String title) throws IOException, ServiceException
        {
                WorksheetFeed worksheetFeed = service.getFeed(this.spreadsheet.getWorksheetFeedUrl(), WorksheetFeed.class);
                List<WorksheetEntry> worksheets = worksheetFeed.getEntries();
    
                boolean worksheetFound = false;
                
                for (int i = 0; i < worksheets.size(); i++)
                {
                        WorksheetEntry sheet = worksheets.get(i);
                        
                        if (sheet.getTitle().getPlainText().contentEquals(title))
                        {
                                this.worksheet = worksheets.get(i);
                                this.cellFeedUrl = worksheet.getCellFeedUrl();
                                this.cellFeed = service.getFeed(cellFeedUrl, CellFeed.class);
                                worksheetFound = true;
                                
                                logger.info("Loaded sheet #" + i + " - " + sheet.getTitle().getPlainText());
                        }
                }

                if (!worksheetFound)
                {
                        logger.error("Cannot find worksheet titled: " + title);
                }
        }
        
        public void login(String username, String password) throws AuthenticationException 
        {
            //try 
            //{
                    // Authenticate
                    service.setUserCredentials(username, password);
            /*} 
            catch (AuthenticationException ex) 
            {
                    logger.error("Unable to authenticate to Google", ex);
            }*/
        }

        public void calculateStats() throws IOException, ServiceException, SQLException, ParseException
        {
                ConfigurationDao config = new ConfigurationDao();
                String siteName = config.getConfiguration("site_id");
                
                if (siteName == null)
                {
                        logger.error("Cannot update worksheet because site_id variable not set");
                        return;
                }
                
                addNewSiteToSheet(siteName);
                
                CellEntry cell = search(siteName);
                CellFeed batchRequest = new CellFeed();
                
                int patientsConsentedRow = cell.getCell().getRow()+1;
                int examSentRow = patientsConsentedRow+1;
                int patientSentRow = examSentRow+1;

                Date date = currentSaturday();
                String currentSaturday = sdf.format(date.getTime());
                
                CellEntry cellSat = search(currentSaturday);
                int column = cellSat.getCell().getCol();
                
                ReportDao report = new ReportDao();

                String patientsConsented = report.patientsConsented(date);
                String examsSent = report.examsSent(date);
                String patientsSent = report.patientsSent(date);

                CellEntry batchOperation = createUpdateOperation(patientsConsentedRow,column,patientsConsented);
                batchRequest.getEntries().add(batchOperation);

                batchOperation = createUpdateOperation(examSentRow,column,examsSent);
                batchRequest.getEntries().add(batchOperation);

                batchOperation = createUpdateOperation(patientSentRow,column,patientsSent);
                batchRequest.getEntries().add(batchOperation);
                
                CellFeed batchResponse = submitBatch(batchRequest);
                printBatchStatus(batchResponse);
                
        }
        
        private void addNewSiteToSheet(String siteName) 
        {
                CellEntry cell = search(siteName);
                int lastRow;
                
                //if date not found
                if (cell == null)
                {
                        if (isColumnEmpty(START_COLUMN))
                        {
                                lastRow = 2;
                        }
                        else
                        {
                                int maxRow = worksheet.getRowCount();
                                //int lastRow = lastSiteRow();
                                lastRow = lastSiteRow();
                                
                                //add rows if necessary
                                if (maxRow <  lastRow+5)
                                {
                                        worksheet.setEtag("*");
                                        worksheet.setRowCount(maxRow+5);
                                        
                                        try
                                        {
                                                worksheet.update();
                                        }
                                        catch (Exception ex)
                                        {
                                                logger.error("Unable to update worksheet", ex);
                                        }
                                }
                        }

                        try
                        {
                                CellEntry newcell = new CellEntry(lastRow+2,START_COLUMN,siteName);
                                cellFeed.insert(newcell);

                                newcell = new CellEntry(lastRow+3,1,"Patients consented");
                                cellFeed.insert(newcell);

                                newcell = new CellEntry(lastRow+4,1,"Exams sent to CH");
                                cellFeed.insert(newcell);

                                newcell = new CellEntry(lastRow+5,1,"Patients sent to CH");
                                cellFeed.insert(newcell);   
                        }
                        catch (Exception ex)
                        {
                                logger.error("Unable to insert cells to worksheet", ex);
                        }
                }
        }
        
        public CellEntry createUpdateOperation(int row, int col, String value) throws ServiceException, IOException 
        {
                String batchId = "R" + row + "C" + col;
                URL entryUrl = new URL(cellFeedUrl.toString() + "/" + batchId);
                CellEntry entry = service.getEntry(entryUrl, CellEntry.class);
                entry.changeInputValueLocal(value);
                BatchUtils.setBatchId(entry, batchId);
                BatchUtils.setBatchOperationType(entry, BatchOperationType.UPDATE);

                return entry;
        }
        
         // Get the batch feed URL and submit the batch request
        public CellFeed submitBatch(CellFeed batchRequest) throws IOException, ServiceException
        {
                CellFeed feed = service.getFeed(cellFeedUrl, CellFeed.class);
                Link batchLink = feed.getLink(Link.Rel.FEED_BATCH, Link.Type.ATOM);
                URL batchUrl = new URL(batchLink.getHref());
                
                CellFeed batchResponse = service.batch(batchUrl, batchRequest); 
                
                return batchResponse;
        }
        
         // Print any errors that may have happened.
        public void printBatchStatus(CellFeed batchResponse) 
        {
                boolean isSuccess = true;
                for (CellEntry entry : batchResponse.getEntries()) 
                {
                        String batchId = BatchUtils.getBatchId(entry);
                        if (!BatchUtils.isSuccess(entry)) 
                        {
                                isSuccess = false;
                                BatchStatus status = BatchUtils.getBatchStatus(entry);
                                logger.error("Batch failed on batchId: " + batchId + " (" + status.getReason()
                                    + ") " + status.getContent());
                        }
                }
                
                
                if (isSuccess) 
                {
                        logger.info("Batch operations successful");
                }
                
        }
        
        
        public void printRow(int row) throws IOException, ServiceException
        {
                CellQuery cellQuery = new CellQuery(worksheet.getCellFeedUrl());
                cellQuery.setReturnEmpty(true);
                
                cellQuery.setMinimumRow(row);
                cellQuery.setMaximumRow(row);
                //cellQuery.setMinimumCol(3);
                
                this.cellFeed = service.getFeed(cellQuery, CellFeed.class);
                
                for (CellEntry cellEntry : cellFeed.getEntries()) 
                {
                        Cell cell = cellEntry.getCell();
                }
        }
        
        public List<CellEntry> getRow(int row, boolean returnEmptyCells)
        {
                CellQuery cellQuery = new CellQuery(worksheet.getCellFeedUrl());
                cellQuery.setReturnEmpty(returnEmptyCells);
                
                cellQuery.setMinimumRow(row);
                cellQuery.setMaximumRow(row);
                
                try 
                {
                        this.cellFeed = service.getFeed(cellQuery, CellFeed.class);
                } 
                catch (Exception ex) 
                {
                        logger.error("Unable to retreive worksheet rows", ex);
                }
                
                return cellFeed.getEntries();
        }
 
        public List<CellEntry> getColumn(int column, boolean returnEmptyCells)
        {
                CellQuery cellQuery = new CellQuery(worksheet.getCellFeedUrl());
                cellQuery.setReturnEmpty(returnEmptyCells);
                
                cellQuery.setMinimumCol(column);
                cellQuery.setMaximumCol(column);
                
                try 
                {
                        this.cellFeed = service.getFeed(cellQuery, CellFeed.class);
                } 
                catch (Exception ex) 
                {
                        logger.error("Umable to retrieve worksheet column", ex);
                }
                
                return cellFeed.getEntries();
        }
                
        public boolean isRowEmpty(int row)
        {
                for (CellEntry cellEntry : getRow(row,true)) 
                {
                        if (cellEntry.getCell().getValue() != null)
                        {
                                return false;
                        }
                }
                
                return true;
        }
        
        public boolean isColumnEmpty(int column)
        {
                for (CellEntry cellEntry : getColumn(1,true)) 
                {
                        if (cellEntry.getCell().getValue() != null)
                        {
                                return false;
                        }
                }
                
                return true;
        }
        
        private void refreshCachedData() throws IOException, ServiceException 
        {
                CellQuery cellQuery = new CellQuery(worksheet.getCellFeedUrl());
                cellQuery.setReturnEmpty(true);
                
                this.cellFeed = service.getFeed(cellQuery, CellFeed.class);

                // A subtlety: Spreadsheets row,col numbers are 1-based whereas the
                // cellEntries array is 0-based. Rather than wasting an extra row and
                // column worth of cells in memory, we adjust accesses by subtracting
                // 1 from each row or column number.
                
                cellEntries = new CellEntry[worksheet.getRowCount()+1][worksheet.getColCount()+1];
                
                for (CellEntry cellEntry : cellFeed.getEntries()) 
                {
                        Cell cell = cellEntry.getCell();
                        cellEntries[cell.getRow()][cell.getCol()] = cellEntry;
                }
        }
        
        public CellEntry search(String fullTextSearchString)
        {
                CellQuery query = new CellQuery(this.cellFeedUrl);
                query.setFullTextQuery(fullTextSearchString);
                
                try
                {
                        CellFeed feed = service.query(query, CellFeed.class);

                        for (CellEntry entry : feed.getEntries()) 
                        {
                                return entry;
                        }
                }
                catch (Exception ex)
                {
                        logger.error(ex);
                }
                
                return null;
        }
        
        public int lastDateColumn()
        {
                List<CellEntry> cellEntry = getRow(DATE_HEADER,false);
                return cellEntry.get(cellEntry.size()-1).getCell().getCol();
        }
        
        public int lastSiteRow()
        {
                List<CellEntry> cellEntry = getColumn(1,false);
                return cellEntry.get(cellEntry.size()-1).getCell().getRow();
        }
        
        public void addDateToHeader() throws IOException, ServiceException, ParseException
        {       
                String currentSaturday = sdf.format(currentSaturday());
                
                if (isRowEmpty(DATE_HEADER))
                {   
                        CellEntry batchOperation = createUpdateOperation(DATE_HEADER,DATE_HEADER,currentSaturday);

                        CellFeed batchRequest = new CellFeed();
                        batchRequest.getEntries().add(batchOperation);

                        CellFeed batchResponse = submitBatch(batchRequest);
                }
                else 
                {
                        int colCount = this.worksheet.getColCount();
                        int lastDateColumn = lastDateColumn();
                        
                        if (search(currentSaturday) == null)
                        {   
                                if (colCount <= lastDateColumn)
                                {
                                    extendSpreadsheetColumn(colCount+1);
                                }
                                
                                CellEntry batchOperation = createUpdateOperation(DATE_HEADER,lastDateColumn()+1,currentSaturday);

                                CellFeed batchRequest = new CellFeed();
                                batchRequest.getEntries().add(batchOperation);

                                CellFeed batchResponse = submitBatch(batchRequest);
                                printBatchStatus(batchResponse);
                                
                                logger.info("Added header date to worksheet: " + currentSaturday);
                        }                   
                }



        }
        
        public void extendSpreadsheetColumn(int col)
        {
                worksheet.setEtag("*");
                worksheet.setColCount(col);
                
                try 
                {
                        worksheet.update();
                } 
                catch (Exception ex) 
                {
                        logger.error("Unable to update worksheet", ex);
                        //stop and throw ex
                }
        }
        public static Date currentSaturday()
        {
                Calendar cal = Calendar.getInstance();
 
                cal.add( Calendar.DAY_OF_WEEK, -(cal.get(Calendar.DAY_OF_WEEK))); 
                cal.set(Calendar.HOUR_OF_DAY, 0);
                cal.set(Calendar.MINUTE, 0);
                cal.set(Calendar.SECOND, 0);
                cal.set(Calendar.MILLISECOND, 0);
                
                return cal.getTime();
        }
        
        public CellEntry getCellValue(int row, int col)
        {
                for (CellEntry cell : cellFeed.getEntries()) 
                {
                        if (cell.getCell().getCol() == col && cell.getCell().getRow() == row)
                        {
                            return cell;
                        }
                }

                return null;
        }
        
}
        