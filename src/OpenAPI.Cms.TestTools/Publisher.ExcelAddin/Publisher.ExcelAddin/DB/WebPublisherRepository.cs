using Newtonsoft.Json;
using OpenApi.Cms.TestTools.Client.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenApi.Cms.TestTools.Client.DB
{
    public class WebPublisherRepository
    {
        private string _connString;

        public WebPublisherRepository(string connString)
        {
            _connString = connString;
        }

        public int TestDBConnection(string reportName)
        {
            using (var conn = new SqlConnection(_connString))
            {
                try
                {
                    conn.Open();
                    var command = conn.CreateCommand();
                    command.CommandText = @"SELECT COUNT(row_id) FROM [dbo].[Record] WHERE sheetName=@sheetName";
                    command.Parameters.Add(new SqlParameter("@sheetName", SqlDbType.NVarChar)
                    {
                        Value = reportName
                    });
                    command.CommandType = CommandType.Text;
                    command.CommandTimeout = 300;
                    return (int)command.ExecuteScalar();
                }
                catch
                {
                    throw;
                }
                finally
                {
                    conn.Close();
                }
            }
        }

        public void DeleteReportData(string reportName)
        {
            using (var conn = new SqlConnection(_connString))
            {
                try
                {
                    conn.Open();
                    var command = conn.CreateCommand();
                    command.CommandText = @"DELETE FROM [dbo].[Record] WHERE sheetName=@sheetName";
                    command.Parameters.Add(new SqlParameter("@sheetName", SqlDbType.NVarChar)
                    {
                        Value = reportName
                    });
                    command.CommandType = CommandType.Text;
                    command.CommandTimeout = 60000;
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"{ex.Message} - {ex.StackTrace}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    throw;
                }
                finally
                {
                    conn.Close();
                }
            }
        }


        public GenericCellData[] GetReportData(string reportName, int topRows)
        {
            string spName = "dbo.GetLatestPositions";
            using (var conn = new SqlConnection(_connString))
            {
                try
                {
                    conn.Open();
                    var command = conn.CreateCommand();
                    command.CommandText = spName;
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandTimeout = 300;
                    command.Parameters.Add(new SqlParameter("@sheetName", SqlDbType.NVarChar)
                    {
                        Value = reportName
                    });

                    command.Parameters.Add(new SqlParameter("@topCount", SqlDbType.Int)
                    {
                        Value = topRows
                    });

                    command.Parameters.Add(new SqlParameter("@isPercentage", SqlDbType.Bit)
                    {
                        Value = false
                    });

                    var reader = command.ExecuteReader();
                    return ReadData(reader);
                }
                catch
                {
                    throw;
                }
                finally
                {
                    conn.Close();
                }
            }
        }


        public int PublishPositions(GenericCellData cellData)
        {
            string spName = "dbo.PublishPositions";
            using (var conn = new SqlConnection(_connString))
            {
                try
                {                    
                    conn.Open();
                    var command = conn.CreateCommand();
                    command.CommandText = spName;
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandTimeout = 300;
                    command.Parameters.Add(new SqlParameter("@publishTime", SqlDbType.DateTime2)
                    {
                        Value = cellData.Timestamp
                    });

                    command.Parameters.Add(new SqlParameter("@sheetName", SqlDbType.NVarChar)
                    {
                        Value = cellData.SheetName
                    });

                    command.Parameters.Add(new SqlParameter("@rows", SqlDbType.Int)
                    {
                        Value = cellData.Rows
                    });

                    command.Parameters.Add(new SqlParameter("@columns", SqlDbType.Int)
                    {
                        Value = cellData.Columns
                    });

                    var bytes = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(cellData));
                    command.Parameters.Add(new SqlParameter("@data", SqlDbType.VarBinary)
                    {
                        Value = bytes
                    });

                    if (cellData.MaxStoredRecords.HasValue)
                    {
                        command.Parameters.Add(new SqlParameter("@maxStoredRecords", SqlDbType.BigInt)
                        {
                            Value = cellData.MaxStoredRecords
                        });
                    }

                    if (!string.IsNullOrEmpty(cellData.DataTimeTag))
                    {
                        command.Parameters.Add(new SqlParameter("@dataTimeTag", SqlDbType.NVarChar)
                        {
                            Value = cellData.DataTimeTag
                        });
                    }

                    return (int)command.ExecuteScalar();
                }
                catch
                {
                    throw;
                }
                finally
                {
                    conn.Close();
                }
            }
        }



        private GenericCellData[] ReadData(SqlDataReader reader)
        {
            var ret = new List<GenericCellData>();
            while (reader.Read())
            {
                var bytes = (byte[])reader["data"];
                var json = Encoding.UTF8.GetString(bytes);
                var cellData = JsonConvert.DeserializeObject<GenericCellData>(json);

                cellData.Timestamp = (DateTime)reader["publishTime"];
                cellData.Rows = (int)reader["rows"];
                cellData.Columns = (int)reader["columns"];

                ret.Add(cellData);
            }

            return ret.ToArray();
        }

    }
}
