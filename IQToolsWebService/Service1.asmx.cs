using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.UI;
using System.Web.Services;
using DataLayer;
using System.Data;
using System.IO;
using System.Threading;
using ADOX;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Ionic.Zip;
using System.Net;
using System.Net.Mail;
using JRO;
using System.Collections;
using System.Configuration;
using OfficeOpenXml;
using System.Xml.Linq;

namespace IQToolsWebService
{

    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    public class Service1 : System.Web.Services.WebService
    {

        #region variables

        string ErrorMessage = "";
        string cnnstr = System.Configuration.ConfigurationManager.ConnectionStrings["msSQLconnString"].ConnectionString;
        private const string folder = "C:\\Cohort\\";
        public const string AccessOleDbConnectionStringFormat = "Data Source={0};Provider=Microsoft.Jet.OLEDB.4.0";
        string iqtoolsDB = System.Configuration.ConfigurationManager.AppSettings["IQToolsDB"].ToString();
        string mainServerType = System.Configuration.ConfigurationManager.AppSettings["mainServerType"].ToString();

        #endregion

        [WebMethod]
        public DataSet GetFacility(string sqlString)
        {
            DataSet ds = new DataSet();
            if (sqlString != "")
            {
                Entity theObject = new Entity();
                ClsUtility.Init_Hashtable();
                ds = (DataSet)theObject.ReturnObject(cnnstr, ClsUtility.theParams, sqlString, ClsUtility.ObjectEnum.DataSet, "mssql");

            }
            return ds;
        }

        [WebMethod]
        public DataTable GetData(string sql)
        {
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            Entity en = new Entity();
            ds = (DataSet)en.ReturnObject(cnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataSet, "mssql");
            return ds.Tables[0];
        }

        [WebMethod]
        public DataTable GetDataDB(string sql, string DB, string dbType)
        {
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            Entity en = new Entity();
            if (dbType == "mssql")
            {
                string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + DB;
                ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataSet, "mssql");
            }
            else if (dbType == "mysql")
            {
                String varcnnstr = getConnString("mysql") + ";database=" + DB;
                ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataSet, "mysql");
            }
            return ds.Tables[0];
        }

        [WebMethod]
        public DataTable GetIQToolsDBData(string sql, string dbType)
        {
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            Entity en = new Entity();
         
            if (dbType == "mssql")
            {
                
                string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;                
                ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataSet, "mssql");               
            }
            else if (dbType == "mysql")
            {
                String varcnnstr = getConnString("mysql") + ";database=" + iqtoolsDB;
                ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataSet, "mysql");
            }
            return ds.Tables[0];
        }

        [WebMethod]
        public DataTable GetQueries(string emr, int userID)
        {
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            Entity en = new Entity();

            if (emr.ToLower() != "isante")
            {
                string sql = "SELECT * FROM (SELECT qryName[QRYNAME], qrydescription[DESCRIPTION], '' [CATEGORY], " +
                                " '' [SUB-CATEGORY], qrydefinition[SQL SYNTAX], b.UserName [OWNER] from aa_queries a inner join " +
                                "aa_users b on a.UID = b.UID " +
                                " WHERE a.UID = "+ userID +" AND a.DeleteFlag IS NULL UNION ";
                sql += "SELECT qryName[QRYNAME], qrydescription[DESCRIPTION], c.Category[CATEGORY], " +
                        "b.sbCategory[SUB-CATEGORY], a.qrydefinition[SQL SYNTAX], 'Admin' [OWNER] from aa_queries a left join " +
                        "aa_sbcategory b on a.qryid = b.qryid left join aa_category c on b.catID = c.catID " +
                        "where (uid is null or uid = 17) AND a.DeleteFlag IS NULL ";
                sql += ")queries ORDER BY OWNER DESC";
                string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;
                ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataSet, "mssql");
            }
            else if (emr.ToLower() == "isante")
            {
                string varcnnstr = getConnString(mainServerType, iqtoolsDB);
                ClsUtility.Init_Hashtable();
                ClsUtility.AddParameters("application", SqlDbType.VarChar, "iqtoolslite");
                ClsUtility.AddParameters("userID", SqlDbType.VarChar, userID.ToString());
                ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, "pr_GetQueries_IQTools", ClsUtility.ObjectEnum.DataSet, "mysql");
            }
            return ds.Tables[0];
        }

        [WebMethod]
        public DataTable GetDBs(string countryCode, string emr)
        {
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            Entity en = new Entity();

            if (emr.ToLower() != "isante")
            {
                string sql = "SELECT Name from Sys.Databases Where Name Like 'IQTools_" + emr + countryCode + "%' AND name not IN ('IQTools_IQCare','IQTools_CTC2' " +
                            ",'IQTools_CTC','IQTools_IQChart','IQTools_ICAP') ORDER BY Name";
                string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;
                ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataSet, "mssql");
            }
            else if (emr.ToLower() == "isante")
            {
                string sql = "Select SCHEMA_NAME from INFORMATION_SCHEMA.SCHEMATA Where SCHEMA_NAME LIKE 'IQTools_ISanteHT%'";
                string varcnnstr = getConnString("mysql") + ";database=" + iqtoolsDB;
                ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataSet, "mysql");
            }
            return ds.Tables[0];
        }
        
        [WebMethod]
        public DataTable GetDataDC(List<string> parameters, string sql, string DB, string dbType)
        {
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            Entity en = new Entity();

            Hashtable hash = new Hashtable();
            for (int i = 0; i < parameters.Count; i = i + 3)
            {
                hash.Add((i+1), parameters[i]);
                hash.Add((i + 2), parameters[i + 1]);
                hash.Add((i + 3), parameters[i + 2]);
            }

            if (dbType == "mssql")
            {
                string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + DB;
                ds = (DataSet)en.ReturnObject(varcnnstr, hash, sql, ClsUtility.ObjectEnum.DataSet, "mssql");
            }
            else if (dbType == "mysql")
            {
                String varcnnstr = getConnString("mysql") + ";database=" + DB;
                ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataSet, "mysql");
            }
            return ds.Tables[0];
        }

        private SqlDbType GetTypeHelper(string type)
        {
            switch(type.ToLower()){
                case "int" :
                    return SqlDbType.Int;
                case "datetime" :
                    return SqlDbType.DateTime;
                default:
                    return SqlDbType.Text;
            }            
        }

        public DataTable GetDataByParameters(string sql, string DB, string dbType, string pars)
        {
            return null;
        }

        [WebMethod]
        public DataTable GetDataDBLog(string sql, string DB, string dbType, int UserID, string application, string method)
        {
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            Entity en = new Entity();

            if (dbType == "mssql")
            {
                string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + DB;
                ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataSet, "mssql");
            }
            else if (dbType == "mysql")
            {
                String varcnnstr = getConnString("mysql") + ";database=" + DB;
                ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataSet, "mysql");
            }
            return ds.Tables[0];
        }

        [WebMethod]
        public string Merge(string emr, string sql, string DB, string[] DBs, int UID)
        {
            Thread merger = new Thread(() => MergeThread(emr, sql, DB, DBs, UID));
            merger.SetApartmentState(ApartmentState.STA);
            merger.Start();
            return "OK"; 
        }

        [WebMethod]
        public string CheckAvailability()
        {
            return "OK";
        }
       
        [WebMethod]
        public string[] GetCredentials(string username, string password)
        {
            Entity en = new Entity();
            ClsUtility.Init_Hashtable();
            string[] results = new string[6];
            string varcnnstr = getConnString(mainServerType, iqtoolsDB);
            string SqlString = "select Distinct UserName, Password, FirstName + ' ' + LastName FullName, UID, Email, FirstName, LastName from aa_Users " +
                               "where userName = '" + username.Trim() + "' and deleteflag = 0";
            try
            {
                DataRow dr = (DataRow)en.ReturnObject(varcnnstr, ClsUtility.theParams, SqlString, ClsUtility.ObjectEnum.DataRow, mainServerType);
                if (ClsUtility.Decrypt(dr["Password"].ToString()) == password)
                {
                    results[0] = "Success";
                    results[1] = dr[2].ToString();
                    results[2] = dr[3].ToString();
                    results[3] = dr[4].ToString();
                    results[4] = dr[5].ToString();
                    results[5] = dr[6].ToString();

                    LogUserHistory(Convert.ToInt32(dr[3].ToString()), "login");

                    return results;
                }
                else
                {
                    results[0] = "WrongPassword";
                    return results;
                }

            }
            catch (Exception ex)
            {
                ErrorLogging("<<Service1.asmx.cs: GetCredentials>> : " + ex.Message, "Web Service", 0);

                if (ex.Message.ToLower() == "there is no row at position 0.")
                {

                    results[0] = "NoUser";
                    return results;
                }
                else
                {
                    results[0] = ex.Message;
                    return results;
                }

            }
        }

        [WebMethod]
        public string ChangePassword(int userid, string oldPassword, string newPassword)
        {
            Entity en = new Entity();
            ClsUtility.Init_Hashtable();
            string[] results = new string[3];
            string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;
            String SqlString = "select Password from aa_users where UID = " + userid.ToString().Trim();
            
            DataRow dr = (DataRow)en.ReturnObject(varcnnstr, ClsUtility.theParams, SqlString, ClsUtility.ObjectEnum.DataRow, "mssql");
            if (ClsUtility.Decrypt(dr["Password"].ToString()) == oldPassword)
            {
                int i = 0;
                String updateString = "update aa_users set Password = '" + ClsUtility.Encrypt(newPassword) + "' where UID = " + userid.ToString().Trim();
                try { i = (int)en.ReturnObject(varcnnstr, ClsUtility.theParams, updateString, ClsUtility.ObjectEnum.ExecuteNonQuery, "mssql"); }
                catch (Exception ex) { return ex.Message; }
                if (i == 1) { return "Success"; }
                else return "Fail";
            }
            else
            {
                return "Your old password is incorrect!!!";
            }
            
            return "Unexpected-Error, changing password!!!";
        }

        [WebMethod]
        public string ChangeDetails(int userid, string email, string fname, string lname)
        {
            Entity en = new Entity();
            ClsUtility.Init_Hashtable();
            string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;
            
            int i = 0;
            String updateString = "update aa_users set Email = '" + email + "', FirstName = '" + fname + "', LastName = '" + lname + "' where UID = " + userid.ToString().Trim();
            try { i = (int)en.ReturnObject(varcnnstr, ClsUtility.theParams, updateString, ClsUtility.ObjectEnum.ExecuteNonQuery, "mssql"); }
            catch (Exception ex) { return ex.Message; }
            if (i == 1) { return "Success"; }
            else return "Fail";
        }

        public void MergeThread(string emr, string sql, string DB, string[] DBs, int UID)
        {
            #region merge
            string link = "";
            Entity en = new Entity();
            int i = 0;
            String varcnnstr = "";
            string MergeString = "";
            int rowNum = 0;
            DropMergeTables(UID, emr,DB);


            foreach (string dbx in DBs)
            {
                if (emr != "isante")
                {
                    varcnnstr = getConnString("mssql") + ";Initial Catalog=" + dbx;
                    try
                    {
                        i = (int)en.ReturnObject(varcnnstr, ClsUtility.theParams, "SELECT * INTO " + DB + ".dbo.Mgr_" + UID.ToString() + "" + dbx + " FROM (" + sql + ")temp", ClsUtility.ObjectEnum.ExecuteNonQuery, "mssql");
                        MergeString += " Select * FROM Mgr_" + UID.ToString() + "" + dbx + " UNION ";
                    }
                    catch (Exception ex) {
                        ErrorLogging("<<Service1.asmx.cs: MergeThread>> : " + ex.Message, "Web Service", 0);
                        ErrorMessage += "Merge Error 10 " + ex.Message; } 
                }
                else if (emr == "isante")
                {
                    varcnnstr = getConnString("mysql") + ";Database=" + dbx;
                    try
                    {
                        i = (int)en.ReturnObject(varcnnstr, ClsUtility.theParams, "CREATE TABLE `" + DB + "`.`mgr_" + dbx + "` " + sql + "", ClsUtility.ObjectEnum.ExecuteNonQuery, "mysql");
                        MergeString += " Select * FROM Mgr_" + dbx + " UNION ";
                    }
                    catch (Exception ex) {
                        ErrorLogging("<<Service1.asmx.cs: MergeThread>> : " + ex.Message, "Web Service", 0);
                        ErrorMessage += "Merge Error 11 " + ex.Message; }  
                }
            }
            if (MergeString.Length > 6)
            {
                string toMerge = MergeString.Substring(0, MergeString.Length - 6);
                if (emr != "isante")
                {
                    varcnnstr = getConnString("mssql") + ";Initial Catalog=" + DB;
                    try
                    {
                        i = (int)en.ReturnObject(varcnnstr, ClsUtility.theParams, "SELECT * INTO Mgr_" + UID.ToString() + " FROM (" + toMerge + ")temp", ClsUtility.ObjectEnum.ExecuteNonQuery, "mssql");
                        DataRow n = (DataRow)en.ReturnObject(varcnnstr, ClsUtility.theParams, "SELECT Count(*) FROM Mgr_" + UID.ToString(), ClsUtility.ObjectEnum.DataRow, "mssql");
                        rowNum = Convert.ToInt32(n[0].ToString());
                    }
                    catch (Exception ex)
                    {
                        ErrorLogging("<<Service1.asmx.cs: MergeThread>> : " + ex.Message, "Web Service", 0);
                        ErrorMessage += ex.Message;
                    }
                }
                else if (emr == "isante")
                {
                    varcnnstr = getConnString("mysql") + ";Database=" + DB;
                    try
                    {
                        i = (int)en.ReturnObject(varcnnstr, ClsUtility.theParams, "CREATE TABLE mgr_" + UID.ToString() + toMerge + ";", ClsUtility.ObjectEnum.ExecuteNonQuery, "mysql");
                        DataRow n = (DataRow)en.ReturnObject(varcnnstr, ClsUtility.theParams, "SELECT Count(*) FROM Mgr_" + UID.ToString(), ClsUtility.ObjectEnum.DataRow, "mssql");
                        rowNum = Convert.ToInt32(n[0].ToString());
                    }
                    catch (Exception ex)
                    {
                        ErrorLogging("<<Service1.asmx.cs: MergeThread>> : " + ex.Message, "Web Service", 0);
                        ErrorMessage += "Merge Error 12 " + ex.Message;
                    }
                }
            }

            #endregion merge

                if (i != 0)
                {
                    if (rowNum > 1000)
                    {
                        try
                        {
                        string toUpload = ZipIt(SQLToAccess(emr, DB, "Mgr_" + UID.ToString()));                            
                            link = folder + toUpload;
                        }
                        catch (Exception ex)
                        {
                            ErrorLogging("<<Service1.asmx.cs: MergeThread>> : " + ex.Message, "Web Service", 0);
                            ErrorMessage += "Merge Error 13 " + ex.Message;
                        }
                        DropMergeTables(UID, emr, DB);

                    }
                    else
                    {
                        //SQLToExcel
                        //Set link to file path
                        DataTable exDt = (DataTable)en.ReturnObject(varcnnstr, ClsUtility.theParams, "SELECT * FROM Mgr_" + UID.ToString(), ClsUtility.ObjectEnum.DataTable, "mssql");
                        string exFileName = folder + "Mgr_" + UID.ToString() + ".xlsx";
                        try
                        {
                            CreateExcel(exDt, exFileName);
                        }
                        catch (Exception ex) { ErrorMessage += ex.Message; }
                        link = exFileName;

                    }
                    ClsUtility.Init_Hashtable();
                    DataRow dr = (DataRow)en.ReturnObject(varcnnstr, ClsUtility.theParams, "Select FirstName From aa_Users Where UID = " + UID.ToString(), ClsUtility.ObjectEnum.DataRow, "mssql");
                    EmailIt(dr[0].ToString(), link, ErrorMessage, GetEmail(iqtoolsDB, UID));
                }
            
        }

        private void DropMergeTables(int UID, string emr, string DB)
        {
            DataTable mgrDt = new DataTable();
            Entity en = new Entity();
            ClsUtility.Init_Hashtable();
            if (emr != "isante")
            {
                string sql = "SELECT name from sys.tables where name like 'mgr_" + UID.ToString() + "%'";
                string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + DB;
                int i = 0;

                mgrDt = (DataTable)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataTable, "mssql");
                DataTableReader mgrDr = mgrDt.CreateDataReader();
                while (mgrDr.Read())
                {
                    sql = "DROP TABLE [" + mgrDr[0].ToString() + "]";
                    i = (int)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.ExecuteNonQuery, "mssql");
                }
            }
            else if (emr == "isante")
            {
                string sql = "select TABLE_NAME from INFORMATION_SCHEMA.TABLES " +
                             "WHERE TABLE_SCHEMA = '"+DB+"' " +
                             "AND TABLE_NAME LIKE 'mgr_%'";
                string varcnnstr = getConnString("mysql") + ";Database=" + DB;
                int i = 0;
                mgrDt = (DataTable)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataTable, "mysql");
                DataTableReader mgrDr = mgrDt.CreateDataReader();
                while (mgrDr.Read())
                {
                    sql = "DROP TABLE `" + mgrDr[0].ToString() + "`";
                    i = (int)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.ExecuteNonQuery, "mysql");
                }
            }
        }

        private string SQLToAccess(string emr, string DB, string table)
        {
            
            DataTable dt = new DataTable();
            Entity en = new Entity();

            if (emr != "isante")
            {
                string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + DB;
                dt = (DataTable)en.ReturnObject(varcnnstr, ClsUtility.theParams, "SELECT * FROM " + table, ClsUtility.ObjectEnum.DataTable, "mssql");
            }
            else if (emr == "isante")
            {
                string varcnnstr = getConnString("mysql") + ";Database=" + DB;
                dt = (DataTable)en.ReturnObject(varcnnstr, ClsUtility.theParams, "SELECT * FROM " + table, ClsUtility.ObjectEnum.DataTable, "mysql");
            }
            Catalog cat = new Catalog();
            string str;
            string fileName = folder + table + ".mdb";

            str = "Provider=Microsoft.Jet.OLEDB.4.0;";
            str += "Data Source=" + fileName + ";Jet OLEDB:Engine Type=5";
            Table nTable = new Table();
            nTable.Name = table;
            try
            {
                //Recreate Mdb

                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
                cat.Create(str);

                //Create Columns From DataTable

                foreach (DataColumn dColumn in dt.Columns)
                {
                    nTable.Columns.Append(dColumn.ColumnName, TranslateType(dColumn.DataType), 255);        
                }
                

                foreach (Column col in nTable.Columns)
                {
                    col.Attributes = ColumnAttributesEnum.adColNullable;
                }

                cat.Tables.Append(nTable);

                //Add Data
                using (OleDbConnection conn = new OleDbConnection(str))
                {
                    conn.Open();
                    string values = "";
                    string value = "";
                    string insertsql = "";
                    foreach (DataRow row in dt.Rows)
                    {
                        
                        try
                        {
                            values = "";
                            insertsql = "INSERT INTO " + table + " VALUES(";
                            foreach (DataColumn column in dt.Columns)
                            {
                                //string dataType = column.DataType.ToString();
                                value = row[column].ToString();
                                if (value.Contains("'"))
                                { value = value.Replace("'", "''"); }

                                if (column.DataType.ToString().ToLower() == "system.datetime")
                                {
                                    if (value != "")
                                        values += "CDATE('" + value + "') ,";
                                        //values += "#" + value + "# ,";
                                    else
                                        //values += "'' ,";
                                        values += "NULL ,";
                                }
                                else if (column.DataType.ToString().ToLower().Contains("system.int") || column.DataType.ToString().ToLower() == "system.decimal" || column.DataType.ToString().ToLower() == "system.double")
                                {
                                    if (value != "")
                                        values += value + " ,";
                                    //values += "#" + value + "# ,";
                                    else
                                        //values += "'' ,";
                                        values += "NULL ,";
                                }
                                //{ values += "#" + value + "# ,"; }
                                else
                                values += "'" + value + "' ,";
                            }
                            insertsql += values.Substring(0, values.Length - 1) + ")";
                            OleDbCommand comm = new OleDbCommand(insertsql, conn);
                            comm.ExecuteNonQuery();
                            comm.Dispose();
                        }
                        catch (Exception ex) {
                            ErrorMessage += "Insert Into Access Error:   " + ex.Message;
                        }
                    }
                    conn.Close();
                    conn.Dispose();
                    //conn = null;
                    GC.Collect();
                    Marshal.FinalReleaseComObject(nTable);
                    Marshal.FinalReleaseComObject(cat.Tables);
                    Marshal.FinalReleaseComObject(cat.ActiveConnection);
                    Marshal.FinalReleaseComObject(cat);

                }
            }
            catch (Exception ex) {
                ErrorLogging("<<Service1.asmx.cs: SQLToAccess>> : " + ex.Message, "Web Service", 0);
                ErrorMessage += "Merge Error 13 " + ex.Message; }


            bool x = CompactJetDatabase(fileName);
            return fileName;
        }

        private ADOX.DataTypeEnum TranslateType(Type columnType)
        {
            switch (columnType.UnderlyingSystemType.ToString())
            {
                case "System.Boolean":
                    return ADOX.DataTypeEnum.adBoolean; //supported

                case "System.Byte":
                    return ADOX.DataTypeEnum.adUnsignedTinyInt; //supported

                //case "System.Char":
                //    return ADOX.DataTypeEnum.adChar; //****** not supported

                case "System.DateTime":
                    return ADOX.DataTypeEnum.adDate; //supported

                case "System.Decimal":
                    return ADOX.DataTypeEnum.adInteger;
                
                case "System.Double":
                    return ADOX.DataTypeEnum.adDouble; //supported

                case "System.Int16":
                    return ADOX.DataTypeEnum.adSmallInt; //supported

                case "System.Int32":
                    return ADOX.DataTypeEnum.adInteger; //supported

                case "System.Int64":
                    return ADOX.DataTypeEnum.adBigInt; //***** not known

                case "System.SByte":
                    return ADOX.DataTypeEnum.adTinyInt;  //***** not known

                case "System.Single":
                    return ADOX.DataTypeEnum.adSingle;  //supported

                case "System.UInt16":
                    return ADOX.DataTypeEnum.adUnsignedSmallInt; //***** not known

                case "System.UInt32":
                    return ADOX.DataTypeEnum.adUnsignedInt; //***** not known

                case "System.UInt64":
                    return ADOX.DataTypeEnum.adUnsignedBigInt; //***** not known

                case "System.String":
                default:
                    return ADOX.DataTypeEnum.adVarWChar;
                //return ADOX.DataTypeEnum.adVarChar; //****** not supported
            }
        }

        private string ZipIt(string toZip)
        {
            string zipName = "Merge_" + DateTime.Now.ToShortDateString() + "--" + DateTime.Now.ToShortTimeString() + ".zip";
            zipName = zipName.Replace("\\", "-");
            zipName = zipName.Replace("/", "-");
            zipName = zipName.Replace(":", "-");
            try
            {
                if (File.Exists(folder + zipName))
                {
                    File.Delete(folder + zipName);
                }
                using (ZipFile zip = new ZipFile())
                {
                    zip.AddFile(toZip);
                    zip.Save(folder + zipName);
                }
            }
            catch (Exception ex) {
                ErrorLogging("<<Service1.asmx.cs: ZipIt>> : " + ex.Message, "Web Service", 0);
                ErrorMessage += "Merge Error 14 " + ex.Message; }
            return zipName;

        }


        private void EmailIt(string userName, string dbLink, string errors, string toEmail)
        {
            var fromAddress = new MailAddress("iqtools.merge.module@gmail.com", "IQTools Merge");
            var toAddress = new MailAddress(toEmail);
            const string fromPassword = "Pl3as3fix";
            string body = "";
            string subject = "Merge " + DateTime.Now.ToShortTimeString();
            if (errors == "" && dbLink != "")
            {
                if (dbLink.Contains("xls") || dbLink.Contains("zip"))
                {
                   
                    StringWriter writer = new StringWriter();
                    HtmlTextWriter html = new HtmlTextWriter(writer);
                    html.RenderBeginTag(HtmlTextWriterTag.P);
                    html.WriteEncodedText("Hi " + userName + ",");
                    html.RenderEndTag();
                    html.RenderBeginTag(HtmlTextWriterTag.P);
                    html.WriteEncodedText("The requested merge operation is complete. Please download the attached File.");
                    html.RenderEndTag();
                    html.RenderBeginTag(HtmlTextWriterTag.P);
                    html.WriteEncodedText("Thanks");
                    
                    html.RenderEndTag();
                    html.Flush();
                    body = writer.ToString();
                  
                }
                else
                {
                    StringWriter writer = new StringWriter();
                    HtmlTextWriter html = new HtmlTextWriter(writer);
                    html.RenderBeginTag(HtmlTextWriterTag.P);
                    html.WriteEncodedText("Hi " + userName + ",");
                    html.RenderEndTag();
                    html.RenderBeginTag(HtmlTextWriterTag.P);
                    html.WriteEncodedText("The requested merge operation is complete. Please download from ");
                        html.RenderBeginTag(HtmlTextWriterTag.Link);
                            html.WriteEncodedUrl(dbLink);
                        html.RenderEndTag();
                    html.RenderEndTag();
                    html.RenderBeginTag(HtmlTextWriterTag.P);
                    html.WriteEncodedText("Thanks");
                   
                    html.RenderEndTag();
                    html.Flush();
                    body = writer.ToString();

                }
            }
            else if (errors != "")
            {
                body = "Merge Incomplete. Errors encountered..." + errors;
            }
            try
            {
                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
                };
                MailMessage message = new MailMessage(fromAddress, toAddress);
                message.Subject = subject;
                message.IsBodyHtml = true;
                message.Body = body;
               
                if (dbLink.Contains("xls") || dbLink.Contains("zip"))
                {
                    message.Attachments.Add(new Attachment(dbLink));
                }
                smtp.Send(message);
                message.Dispose();
                
            }
            catch (Exception ex)
            { //txtLog.Text += ex.Message; 
                ErrorLogging("<<Service1.asmx.cs: EmailIt>> : " + ex.Message, "Web Service", 0);
            }
            

        }

        private string GetEmail(string DB, int userID)
        {
            Entity en = new Entity();
            ClsUtility.Init_Hashtable();
            string sql = "SELECT TOP 1 EMail from aa_users where UID = " + userID + "";
            //string varcnnstr = "Data Source=.\\SQLEXPRESS;Initial Catalog=" + DB + ";uid=sa;pwd=c0nstella";
            string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + DB;
            //string varcnnstr = "Data Source=localhost;Initial Catalog=" + DB + ";uid=sa;pwd=c0nstella";
            DataRow dr = (DataRow)en.ReturnObject(varcnnstr, ClsUtility.theParams, sql, ClsUtility.ObjectEnum.DataRow, "mssql");
            return dr[0].ToString();
        }

        public static bool CompactJetDatabase(string fileName)
        {
            try
            {
                string newFileName = Path.Combine(folder, Guid.NewGuid().ToString("N") + ".mdb");
                JetEngine engine = new JetEngine();
                string sourceConnection = String.Format(AccessOleDbConnectionStringFormat, fileName);
                string destConnection = String.Format(AccessOleDbConnectionStringFormat, newFileName);
                engine.CompactDatabase(sourceConnection, destConnection);
                File.Delete(fileName);
                File.Move(newFileName, fileName);
                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }

        [WebMethod]
        public string SaveQuery(string UID, string emr, string sql, string qryName, string qryDesc)
        {
            string connString = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;
            Entity en = new Entity();
            ClsUtility.Init_Hashtable();
            int i = 0;
            string insertString = "INSERT INTO [dbo].[aa_Queries] " +
                                "([qryName] " +
                                ",[qryDefinition] " +
                                ",[qryDescription] " +
                                ",[qryType] " +
                                ",[CreateDate] " +
                                ",[UpdateDate] " +
                                ",[Deleteflag] " +
                                ",[MkTable] " +
                                ",[Decrypt] " +
                                ",[Hidden] " +
                                ",[qryGroup] " +
                                ",[UID]) " +
                                "VALUES " +
                                "( '"+qryName+"' " +
                                ", '"+sql+"' " +
                                ", '"+qryDesc+"' " +
                                ", 'UserQuery' " +
                                ", getdate() " +
                                ", null " +
                                ", null " +
                                ", null " +
                                ", null " +
                                ", null " +
                                ", null, "+UID+") ";

            try { i = (int)en.ReturnObject(connString, ClsUtility.theParams, insertString, ClsUtility.ObjectEnum.ExecuteNonQuery, "mssql"); }
            catch (Exception ex) { return ex.Message; }
            if (i == 1) { return "Success"; }
            else return "Fail";
        }

        [WebMethod]
        public string UpdateQuery(string UID, string emr, string sql, string qryName, string qryDesc)
        {
            string connString = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;
            Entity en = new Entity();
            ClsUtility.Init_Hashtable();
            int i = 0;
            DataRow dr = (DataRow)en.ReturnObject(connString,ClsUtility.theParams,"SELECT TOP 1 qryID FROM aa_queries where qryName = '"+qryName + "' AND UID = " + UID + "",ClsUtility.ObjectEnum.DataRow,"mssql");
            string qryID = dr[0].ToString();
            string updateString = "UPDATE aa_queries SET qrydefinition = '" + sql + "', qrydescription = '" + qryDesc + "' where qryID = " + qryID + "";

            try { i = (int)en.ReturnObject(connString, ClsUtility.theParams, updateString, ClsUtility.ObjectEnum.ExecuteNonQuery, "mssql"); }
            catch (Exception ex) { return ex.Message; }
            if (i == 1) { return "Success"; }
            else return "Fail";
        }

        private string getConnString(string server)
        {            
            if (server == "mssql")
            {
                return System.Configuration.ConfigurationManager.ConnectionStrings["msSQLconnString"].ConnectionString;
            }
            else if (server == "mysql")
            {
                return System.Configuration.ConfigurationManager.ConnectionStrings["mySQLconnString"].ConnectionString;
            }
            else return System.Configuration.ConfigurationManager.ConnectionStrings["msSQLconnString"].ConnectionString;
        }

        private string getConnString(string server, string DB)
        {
            if (server == "mssql")
            {
                return System.Configuration.ConfigurationManager.ConnectionStrings["msSQLconnString"].ConnectionString + ";Initial Catalog=" + DB;
            }
            else if (server == "mysql")
            {
                return System.Configuration.ConfigurationManager.ConnectionStrings["mySQLconnString"].ConnectionString + ";Database=" + DB;
            }
            else return System.Configuration.ConfigurationManager.ConnectionStrings["msSQLconnString"].ConnectionString + ";Initial Catalog=" + DB;
        }

        [WebMethod]
        public string RetrievePassword(string userName)
        {
            string msg = string.Empty;
            Entity en = new Entity();
            ClsUtility.Init_Hashtable();
            string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;
            
            String selectString = "select password, email from aa_users where UserName like '%" + userName + "%'";
            try
            {
                DataSet ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, selectString, ClsUtility.ObjectEnum.DataSet, "mssql");
                if ((ds != null) && (ds.Tables.Count > 0) && (ds.Tables[0].Rows.Count > 0))
                {
                    string pass = ds.Tables[0].Rows[0]["password"].ToString();
                    string email = ds.Tables[0].Rows[0]["email"].ToString();
                    string body = "Hi " + userName + "<br/><br/>You are receiving this email because you requested for your password for IQToolsLite application : <br/> Password : " + ClsUtility.Decrypt(pass) + "<br/><br/>Regards,";
                    SendEmail(email, "Iqtools.merge.module@gmail.com", "Password retrieval", body);
                }
                else
                {
                    return "No user found, please double check your username!!";
                }
                return string.Empty;
            }
            catch (Exception ex) { 
                ErrorLogging(ex.Message, "Webservice", 0);
                return "Problems encounted retrieving password, please contact the administrator."; 
            }
        }

        [WebMethod]
        public void LogUserHistory(int UserID, string logInOut)
        {
            string msg = string.Empty;
            Entity en = new Entity();
            ClsUtility.Init_Hashtable();
            string varcnnstr = getConnString(mainServerType, iqtoolsDB);

            string updateString = "insert into aa_UserHistory (UID, LoginTime) values(" + UserID + ", '" + DateTime.Now + "')";
            if(logInOut == "logout")
                updateString = "update aa_UserHistory set LogOutTime = '" + DateTime.Now + "' where UID = " + UserID + " and UserHistoryID = " +
                    "(select Max(UserHistoryID) UserHistoryID from aa_UserHistory where UID = " + UserID + ")";
            try
            {
                int i = (int)en.ReturnObject(varcnnstr, ClsUtility.theParams, updateString, ClsUtility.ObjectEnum.ExecuteNonQuery, mainServerType);
            }
            catch (Exception ex) {
                ErrorLogging(ex.Message, "IQToolsLite", UserID);
            }
        }

        [WebMethod]
        public bool PasswordChanged(string userName)
        {
            string msg = string.Empty;
            Entity en = new Entity();
            ClsUtility.Init_Hashtable();
            string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;

            String selectString = "select password from aa_users where UserName like '%" + userName + "%'";
            try
            {
                DataSet ds = (DataSet)en.ReturnObject(varcnnstr, ClsUtility.theParams, selectString, ClsUtility.ObjectEnum.DataSet, "mssql");
                if ((ds != null) && (ds.Tables.Count > 0) && (ds.Tables[0].Rows.Count > 0))
                {
                    string pass = ds.Tables[0].Rows[0]["password"].ToString();
                    if (ClsUtility.Decrypt(pass).ToLower().Trim() != "CrsUser2013".ToLower())
                        return true;
                }
            }
            catch (Exception ex)
            {
                ErrorLogging(ex.Message, "Webservice", 0);
            }
            return false;
        }

      
        [WebMethod]
        public IQToolsWebService.Service1.GeoLocation GetUserLocation ( string ipAddress )
        {
            try
            {
        
            string apiKey = ConfigurationManager.AppSettings["ApiKey"];
            string url = string.Format ( "http://api.ipinfodb.com/v3/ip-city/?key={0}&ip={1}&format=xml", apiKey, ipAddress );
            var result = XDocument.Load(url);

            var location = (from x in result.Descendants("Response")
                            select new GeoLocation
                            {
                                    CityName = (string)x.Element("cityName"),
                                    RegionName = (string)x.Element("regionName"),
                                    CountryName = (string)x.Element("countryName"),
                                    CountryCode = (string)x.Element("countryCode")
                                }).First();


            return location;
        }
            catch (Exception ex) { return null; }
        }

        public class GeoLocation 
          {
     
          public string IPAddress { get; set; }
          
          public string CountryName { get; set; }
          public string CountryCode { get; set; }
          public string CityName { get; set; }
          public string RegionName { get; set; }
          public string ZipCode { get; set; }

          }

        [WebMethod]
        public string ErrorLogging(string errorMsg, string application, int UserID)
        {
            string msg = string.Empty;
            Entity en = new Entity();
            ClsUtility.Init_Hashtable();
            string varcnnstr = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;

            errorMsg = errorMsg.Replace("'", "''");
            String updateString = "insert into aa_errorLogs (UID, Application, Message) values(" + UserID + ", '" + application + "', '" + errorMsg + "' )";
            try
            {
                int i = (int)en.ReturnObject(varcnnstr, ClsUtility.theParams, updateString, ClsUtility.ObjectEnum.ExecuteNonQuery, "mssql");
                ErrorEmailing(application, errorMsg);
                return string.Empty;
            }
            catch (Exception ex) { return ex.Message; }
        }
        
        private void ErrorEmailing(string application, string errorMsg)
        {
            string body = "<b>Application :</b> " + application + "</br>" +
                          "<b>Date : </b>" + DateTime.Now.ToShortDateString() + "</br></br>" +
                          errorMsg;

            SendEmail(ConfigurationManager.AppSettings["ErrorEmailTo"].ToString(),
                ConfigurationManager.AppSettings["ErrorEmailFrom"].ToString(),
                ConfigurationManager.AppSettings["ErrorEmailSubject"].ToString(), body);
        }

        private void SendEmail(string to, string from, string subject, string body)
        {           
            try
            {
                MailAddress SendFrom = new MailAddress(from);
                MailAddress SendTo = new MailAddress(to);

                const string fromPassword = "Pl3as3fix";

                try
                {
                    var smtp = new SmtpClient
                    {
                        Host = "smtp.gmail.com",
                        Port = 587,
                        EnableSsl = true,
                        DeliveryMethod = SmtpDeliveryMethod.Network,
                        UseDefaultCredentials = false,
                        Credentials = new NetworkCredential(SendFrom.Address, fromPassword)
                    };

                    using (var message = new MailMessage(SendFrom, SendTo)
                    {
                        Subject = subject,
                        Body = body,
                        IsBodyHtml = true
                        
                    })
                    {
                        smtp.Send(message);
                    }
                }
                catch (Exception ex)
                { 
                    ErrorLogging("<<Service1.asmx.cs: EmailIt>> : " + ex.Message, "Web Service", 0);
                }         
            }
            catch (Exception ex)
            {
                
            }
        }

        private void TestEmailSending(string body)
        {
            string from = "iqtools.merge.module@gmail.com";
            string to = "smkhwanazi@futuresgroup.com";

            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
            mail.To.Add(to);
            mail.From = new MailAddress(from, "IQTools admin", System.Text.Encoding.UTF8);
            mail.Subject = "IQTools Error";
            mail.SubjectEncoding = System.Text.Encoding.UTF8;
            mail.Body = body;
            mail.BodyEncoding = System.Text.Encoding.UTF8;
            mail.IsBodyHtml = true;
            mail.Priority = MailPriority.High;

            SmtpClient client = new SmtpClient();
            client.Credentials = new System.Net.NetworkCredential(from, "Pl3as3fix");
            client.Port = 587; // Gmail works on this port
            client.Host = "smtp.gmail.com";
            client.EnableSsl = true; //Gmail works on Server Secured Layer

            try
            {
                client.Send(mail);
            }
            catch (Exception ex)
            {
                Exception ex2 = ex;
                string errorMessage = string.Empty;
                while (ex2 != null)
                {
                    errorMessage += ex2.ToString();
                    ex2 = ex2.InnerException;
                }
            } 
        }

        private void CreateExcel(DataTable dt, string fName)
        {
            try
            { File.Delete(fName); }
            catch { }
            try
            {
                FileInfo newFile = new FileInfo(fName);
                using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                {
                    ExcelWorksheet workSheet = xlPackage.Workbook.Worksheets.Add("IQTools");
                    int iRow = 2;

                    for (int j = 0; j < dt.Columns.Count; j++)
                    { workSheet.Cell(1, j + 1).Value = dt.Columns[j].ColumnName; }

                    for (int rowNo = 0; rowNo < dt.Rows.Count; rowNo++)
                    {
                        for (int colNo = 0; colNo < dt.Columns.Count; colNo++)
                        {
                            workSheet.Cell(iRow, colNo + 1).Value = dt.Rows[rowNo][colNo].ToString().Replace("'","");                          
                        }
                        iRow++;
                    }
                    xlPackage.Workbook.Properties.Author = "IQTools Lite";
                    xlPackage.Workbook.Properties.Company = "Futures Group";
                    xlPackage.Save();
                    xlPackage.Dispose();
                }
            }
            catch (Exception ex)
            {
                ErrorMessage += ex.Message;                
            }
        }

        [WebMethod]
        public string SiteHandshake(string MFLCode, string SiteName)
        {
            Entity en = new Entity();
            ClsUtility.Init_Hashtable();
            string connString = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;
            string InsertString = "INSERT INTO aa_RemoteServices (ServiceName, MFLCode, SiteName, ServiceStatus, CreateDate) " +
                                  "VALUES ('SiteHandshake','" + MFLCode + "','" + SiteName + "','OK',GetDate())";
            try
            {
                int i = (int)en.ReturnObject(connString, ClsUtility.theParams, InsertString, ClsUtility.ObjectEnum.ExecuteNonQuery, "mssql");
                return "Connected";
            }
            catch (Exception ex) { return ex.Message; }
        }
      
        [WebMethod]
        public DataTable DBCompare(DataTable SiteObjects)
        {
            try
            {              
                DataTable masterObjects = new DataTable("Input");
                string objectid = string.Empty;
                Entity en = new Entity();
                ClsUtility.Init_Hashtable();
                string connString = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;
                ClsUtility.AddParameters("@WithSyntax", SqlDbType.Text, "1");
                string sp = "pr_GetQueriesForUpdate_IQTools";
              
                masterObjects = (DataTable)en.ReturnObject(connString, ClsUtility.theParams
                    , sp, ClsUtility.ObjectEnum.DataTable, "mssql");

                DataTable Changes = new DataTable("Output");

                Changes = (from rA in masterObjects.AsEnumerable()
                           join rB in SiteObjects.AsEnumerable()
                           on rA.Field<string>("ROUTINE_NAME") equals rB.Field<string>("ROUTINE_NAME") into rC
                           from c in rC.DefaultIfEmpty()
                           where c == null ? true :
                           rA.Field<DateTime>("LAST_ALTERED") > c.Field<DateTime>("LAST_ALTERED")
                           select rA).CopyToDataTable();    
             
                if(Changes != null)
                {
                    DataSet ds = new DataSet();
                    ds.Tables.Add(Changes);
                    return ds.Tables[0];
                }
                else return null;
            }
            catch (Exception ex) { return null; }

        }
        
        [WebMethod]
        public string GetDBVersion()
        {
            string SQL = "Select DBVersion FROM aa_Version";
            string DBVersion = string.Empty;
            try
            {
                string connString = getConnString("mssql") + ";Initial Catalog=" + iqtoolsDB;
                Entity en = new Entity();
                ClsUtility.Init_Hashtable();
                DataRow dr = (DataRow)en.ReturnObject(connString, ClsUtility.theParams, SQL
                    , ClsUtility.ObjectEnum.DataRow, "mssql");
                if(!dr.IsNull(0))
                {
                    DBVersion = dr[0].ToString();
                }

            }
            catch(Exception ex)
            {
                //Log this error
            }
            return DBVersion;           

        }
    
    }
}
