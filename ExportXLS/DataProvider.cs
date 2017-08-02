using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace ExportXLS
{
    class DataProvider
    {
        private string cs;

        public DataProvider(string server,string database,string user,string password)
        {
            SqlConnectionStringBuilder csb = new SqlConnectionStringBuilder();
            csb.DataSource = server;
            csb.InitialCatalog = database;
            csb.UserID = user;
            csb.Password = password;
            cs = csb.ToString();
        }

        public DataTable GetProfile(int objectCode, int itemCode, DateTime day)
        {
            DataTable result = new DataTable();
            using (SqlConnection cn = new SqlConnection(cs))
            {
                cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandText = string.Format("SELECT * FROM dbo.GetProfile({0},{1},'{2}')",
                    objectCode, itemCode, day.ToString("yyyyMMdd"));
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                try
                {
                    da.Fill(result);
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Ошибка запроса профиля по {0}.{1} за {2}",
                                                       objectCode, itemCode, day),
                                        ex);
                }
            }
            return result;
        }

        public string GetSensorName(string deviceCode, string sensorCode)
        {
            object result;
            using (SqlConnection cn = new SqlConnection(cs))
            {
                cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandText = string.Format(@"SELECT Sensors.Name FROM Sensors INNER JOIN Devices " +
                                                    "ON Sensors.StationID=Devices.ID " +
                                                    "WHERE Devices.Code={0} AND Sensors.Code={1}",
                                                deviceCode, sensorCode);
                try
                {
                    result = cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Ошибка получения имени канала по {0}.{1}",
                                                       deviceCode, sensorCode),
                                        ex);
                }
                if (result == null || Convert.IsDBNull(result))
                    throw new Exception(string.Format("Ошибка получения имени канала по {0}.{1}",
                        deviceCode,sensorCode) + Environment.NewLine + "Запрос вернул пустой набор строк");
            }
            return result.ToString();
        }

        public string GetDeviceName(string deviceCode)
        {
            object result;
            using (SqlConnection cn = new SqlConnection(cs))
            {
                cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandText = "SELECT Name FROM Devices WHERE Code=" + deviceCode;
                try
                {
                    result = cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Ошибка получения имени устройства по {0}",
                        deviceCode), ex);
                }
                if (result == null || Convert.IsDBNull(result))
                    throw new Exception(string.Format("Ошибка получения имени устройства по {0}",
                        deviceCode) + Environment.NewLine + "Запрос вернул пустой набор строк");
                return result.ToString();
            }
        }

        public int GatheredData(int objCode, int itemCode, DateTime day)
        {
            object result;
            using (SqlConnection cn = new SqlConnection(cs))
            {
                cn.Open();
                SqlCommand cmd = cn.CreateCommand();
                StringBuilder sql = new StringBuilder();
                sql.Append("SELECT count(*) res FROM DATA WHERE Parnumber=12 AND ");
                sql.AppendFormat("Object={0} AND Item={1} AND ", objCode, itemCode);
                sql.AppendFormat("Data_Date between '{0}' AND '{1}' ",
                    day.AddMinutes(30).ToString("yyyyMMdd HH:mm"),
                    day.AddDays(1).ToString("yyyyMMdd"));
                cmd.CommandText = sql.ToString();
                try
                {
                    result = cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Ощибка подсчёта количества собранных значений для {0}.{1}",
                        objCode, itemCode), ex);
                }
                if (result != null && !Convert.IsDBNull(result))
                    return Convert.ToInt32(result);
                else
                    return 0;
            }
        }



        public double GetSingleHalfhour(string deviceCode, string sensorCode, DateTime halfhour)
        {
            object result;
            using (SqlConnection cn = new SqlConnection(cs))
            {
                cn.Open();
                StringBuilder sql = new StringBuilder();
                sql.Append("SELECT value0 FROM DATA WHERE ");
                sql.AppendFormat("Object={0} AND Item={1} AND Parnumber=12 AND Data_Date='{2}'",
                    deviceCode, sensorCode, halfhour.ToString("yyyyMMdd HH:mm"));
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandText = sql.ToString();
                try
                {
                    result = cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    throw new Exception(
                        string.Format("Не удалось получить значение получасовки для {0}.{1} за {2}",
                                            deviceCode,sensorCode,halfhour),ex);
                }
                if (result == null || Convert.IsDBNull(result))
                    throw new Exception(string.Format("Ошибка получения значения получасовки для {0}.{1} за {2}",
                    deviceCode, sensorCode, halfhour) +
                    Environment.NewLine + "Запрос вернул пустой набор строк");
                return (double)result;
            }
        }

        public double GetFixedValue(string deviceCode, string sensorCode, DateTime timePoint)
        {
            object result;
            using (SqlConnection cn = new SqlConnection(cs))
            {
                cn.Open();
                StringBuilder sql = new StringBuilder();
                sql.Append("SELECT value0 FROM DATA WHERE ");
                sql.AppendFormat("Object={0} AND Item={1} AND Parnumber=101 AND Data_Date='{2}'",
                    deviceCode, sensorCode, timePoint.ToString("yyyyMMdd HH:mm"));
                SqlCommand cmd = cn.CreateCommand();
                cmd.CommandText = sql.ToString();
                try
                {
                    result = cmd.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    throw new Exception(
                        string.Format("Не удалось получить показания для {0}.{1} на {2}",
                                            deviceCode, sensorCode, timePoint), ex);
                }
                if (result == null || Convert.IsDBNull(result))
                    throw new Exception(string.Format("Ошибка получения показаний для {0}.{1} на {2}",
                    deviceCode, sensorCode, timePoint) +
                    Environment.NewLine + "Запрос вернул пустой набор строк");
                return (double)result;
            }
        }
    }
}
