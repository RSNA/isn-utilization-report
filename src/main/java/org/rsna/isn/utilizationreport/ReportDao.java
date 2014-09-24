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

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.util.Date;
import org.rsna.isn.dao.Dao;

/**
 *
 * @author Clifton Li
 * @version 3.2.0
 * @since 3.2.0
 */
public class ReportDao extends Dao 
{
    public String patientsConsented(Date date) throws SQLException, ParseException
    {       
            Connection con = getConnection();

            PreparedStatement stmt = con.prepareStatement("SELECT count(*) AS count FROM patients WHERE consent_timestamp IS NOT NULL and \"consent_timestamp\" <= ?");
            stmt.setDate(1, new java.sql.Date(date.getTime()));

            int report = 0;
            
            ResultSet rs = stmt.executeQuery();
            
            if(rs.next())
            {
                    report = rs.getInt("count");               
            }
            
            rs.close();
            con.close();

            return Integer.toString(report);
    }

    public String patientsConsented() throws SQLException, ParseException
    {       
            Connection con = getConnection();

            PreparedStatement stmt = con.prepareStatement("SELECT count(*) AS count FROM patients WHERE consent_timestamp IS NOT NULL");

            int report = 0;
            
            ResultSet rs = stmt.executeQuery();
            
            if(rs.next())
            {
                    report = rs.getInt("count");               
            }
            
            rs.close();
            con.close();

            return Integer.toString(report);
    }
        
    public String examsSent(Date date) throws SQLException, ParseException
    {      
            Connection con = getConnection();

            PreparedStatement stmt = con.prepareStatement("SELECT count(*) AS count FROM transactions WHERE status_code = 40 and \"modified_date\" <= ?");
            stmt.setDate(1, new java.sql.Date(date.getTime()));

            int report = 0;
            
            ResultSet rs = stmt.executeQuery();
            
            if(rs.next())
            {
                    report = rs.getInt("count");               
            }
            
            rs.close();
            con.close();

            return Integer.toString(report);        
    }

    public String examsSent() throws SQLException, ParseException
    {      
            Connection con = getConnection();

            PreparedStatement stmt = con.prepareStatement("SELECT count(*) AS count FROM transactions WHERE status_code = 40");

            int report = 0;
            
            ResultSet rs = stmt.executeQuery();
            
            if(rs.next())
            {
                    report = rs.getInt("count");               
            }
            
            rs.close();
            con.close();

            return Integer.toString(report);        
    }
        
    public String patientsSent(Date date) throws SQLException, ParseException
    {      
            Connection con = getConnection();

            PreparedStatement stmt = con.prepareStatement("SELECT COUNT(DISTINCT job_sets.patient_id) AS count " +
                    "FROM transactions, jobs, job_sets WHERE " +
                    "(transactions.status_code = 40) AND " +
                    "(transactions.job_id = jobs.job_id) AND " +
                    "(jobs.job_set_id = job_sets.job_set_id) AND transactions.modified_date <= ?");
            
            stmt.setDate(1, new java.sql.Date(date.getTime()));

            int report = 0;
            
            ResultSet rs = stmt.executeQuery();
            
            if(rs.next())
            {
                    report = rs.getInt("count");               
            }
            
            rs.close();
            con.close();

            return Integer.toString(report);        
    }

    public String patientsSent() throws SQLException, ParseException
    {      
            Connection con = getConnection();

            PreparedStatement stmt = con.prepareStatement("SELECT COUNT(DISTINCT job_sets.patient_id) AS count " +
                    "FROM transactions, jobs, job_sets WHERE " +
                    "(transactions.status_code = 40) AND " +
                    "(transactions.job_id = jobs.job_id) AND " +
                    "(jobs.job_set_id = job_sets.job_set_id)");

            int report = 0;
            
            ResultSet rs = stmt.executeQuery();
            
            if(rs.next())
            {
                    report = rs.getInt("count");               
            }
            
            rs.close();
            con.close();

            return Integer.toString(report);        
    }
    
}
