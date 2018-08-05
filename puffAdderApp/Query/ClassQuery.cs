using System;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System.Reflection; 

namespace PuffAdderApplication
{
    class ClassQuery
    {
        PuffAdderApplication.DTO.LogDTO logDTO = new PuffAdderApplication.DTO.LogDTO();

        /// <summary>
        /// DB 저장 - 공통
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="dt"></param>
        public void DB_insert(string tableName, DataTable dt)
        {
            string MyconnectString = "server=127.0.0.1; port=3306 ; database= investment ; uid=root ; pwd= 5587 ; SslMode=none";
            logDTO = new PuffAdderApplication.DTO.LogDTO();
            logDTO.setDbNm(tableName);
            logDTO.setFuncNm(MethodBase.GetCurrentMethod().Name);
            MySqlConnection conn = new MySqlConnection(MyconnectString);
            conn.Open();

            try
            {
                switch (tableName)
                {
                    case "price":
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            MySqlCommand comm = conn.CreateCommand();

                            comm.CommandText = "INSERT INTO price" +
                                                    " (CATE_CD" +
                                                    ", TRADE_DT" +
                                                    ", ST_PRICE" +
                                                    ", HIGH_PRICE" +
                                                    ", LOW_PRICE" +
                                                    ", END_PRICE" +
                                                    ", TRADE_SUM) " +
                                              "VALUES(@CATE_CD" +
                                                   ", @TRADE_DT" +
                                                   ", @ST_PRICE" +
                                                   ", @HIGH_PRICE" +
                                                   ", @LOW_PRICE" +
                                                   ", @END_PRICE" +
                                                   ", @TRADE_SUM) " +
                                                   "ON DUPLICATE KEY " +
                                               "UPDATE ST_PRICE = @ST_PRICE" +
                                                    ", HIGH_PRICE = @HIGH_PRICE" +
                                                    ", LOW_PRICE = @LOW_PRICE" +
                                                    ", END_PRICE = @END_PRICE" +
                                                    ", TRADE_SUM = @TRADE_SUM";

                            logDTO.setCateCd(dt.Rows[i]["cateCd"].ToString());
                            logDTO.setTradeDt(dt.Rows[i]["tradeDt"].ToString());

                            comm.Parameters.AddWithValue("@CATE_CD", dt.Rows[i]["cateCd"]);
                            comm.Parameters.AddWithValue("@TRADE_DT", dt.Rows[i]["tradeDt"]);
                            comm.Parameters.AddWithValue("@ST_PRICE", dt.Rows[i]["stPrice"]);
                            comm.Parameters.AddWithValue("@HIGH_PRICE", dt.Rows[i]["highPrice"]);
                            comm.Parameters.AddWithValue("@LOW_PRICE", dt.Rows[i]["lowPrice"]);
                            comm.Parameters.AddWithValue("@END_PRICE", dt.Rows[i]["endPrice"]);
                            comm.Parameters.AddWithValue("@TRADE_SUM", dt.Rows[i]["tradeSum"]);

                            comm.ExecuteNonQuery();
                        }
                        break;

                    case "categ":
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            MySqlCommand comm = conn.CreateCommand();
                            comm.CommandText = "INSERT INTO categ" +
                                                    " (CATE_CD" +
                                                    ", CATE_NM" +
                                                    ", MARKET" +
                                                    ", INDU_BASIC" +
                                                    ", INDU_DTL" +
                                                    ", VALD_YN" +
                                                    ", ST_DT" +
                                                    ", END_DT) " +
                                              "VALUES(@CATE_CD" +
                                                   ", @CATE_NM" +
                                                   ", @MARKET" +
                                                   ", @INDU_BASIC" +
                                                   ", @INDU_DTL" +
                                                   ", @VALD_YN" +
                                                   ", @ST_DT" +
                                                   ", @END_DT) " +
                                                   "ON DUPLICATE KEY " +
                                               "UPDATE CATE_NM = @CATE_NM" +
                                                    ", MARKET  = @MARKET" +
                                                    ", INDU_BASIC = @INDU_BASIC" +
                                                    ", INDU_DTL = @INDU_DTL" +
                                                    ", VALD_YN = @VALD_YN" +
                                                    ", ST_DT = @ST_DT" +
                                                    ", END_DT = @END_DT";

                            comm.Parameters.AddWithValue("@CATE_CD", dt.Rows[i]["cateCd"]);
                            comm.Parameters.AddWithValue("@CATE_NM", dt.Rows[i]["cateNm"]);
                            comm.Parameters.AddWithValue("@MARKET", dt.Rows[i]["market"]);
                            comm.Parameters.AddWithValue("@INDU_BASIC", dt.Rows[i]["induBasic"]);
                            comm.Parameters.AddWithValue("@INDU_DTL", dt.Rows[i]["induDtl"]);
                            comm.Parameters.AddWithValue("@VALD_YN", "Y");
                            comm.Parameters.AddWithValue("@ST_DT", "19000101");
                            comm.Parameters.AddWithValue("@END_DT", "99991231");
                            comm.ExecuteNonQuery();
                        }
                        break;

                    case "trade":
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            if (dt.Rows[i]["tradeDt"] == null || dt.Rows[i]["tradeDt"].ToString() == "" || dt.Rows[i]["tradeDt"].ToString() == "0")
                            {
                                return;
                            }

                            MySqlCommand comm = conn.CreateCommand();
                            comm.CommandText = "INSERT INTO trade" +
                                                    " (CATE_CD" +
                                                    ", TRADE_DT" +
                                                    ", ORG_SUM" +
                                                    ", FOR_SUM" +
                                                    ", FOR_HAVE_CNT" +
                                                    ", FOR_HAVE_PNT) " +
                                              "VALUES(@CATE_CD" +
                                                   ", @TRADE_DT" +
                                                   ", @ORG_SUM" +
                                                   ", @FOR_SUM" +
                                                   ", @FOR_HAVE_CNT" +
                                                   ", @FOR_HAVE_PNT) " +
                                                   "ON DUPLICATE KEY " +
                                               "UPDATE ORG_SUM = @ORG_SUM" +
                                                    ", FOR_SUM  = @FOR_SUM" +
                                                    ", FOR_HAVE_CNT = @FOR_HAVE_CNT" +
                                                    ", FOR_HAVE_PNT = @FOR_HAVE_PNT";

                            comm.Parameters.AddWithValue("@CATE_CD", dt.Rows[i]["cateCd"]);
                            comm.Parameters.AddWithValue("@TRADE_DT", dt.Rows[i]["tradeDt"]);
                            comm.Parameters.AddWithValue("@ORG_SUM", dt.Rows[i]["orgSum"]);
                            comm.Parameters.AddWithValue("@FOR_SUM", dt.Rows[i]["forSum"]);
                            comm.Parameters.AddWithValue("@FOR_HAVE_CNT", dt.Rows[i]["forHaveCnt"]);
                            comm.Parameters.AddWithValue("@FOR_HAVE_PNT", dt.Rows[i]["forHavePnt"]);
                            comm.ExecuteNonQuery();
                        }
                        break;

                    case "success":
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            MySqlCommand comm = conn.CreateCommand();
                            comm.CommandText = "INSERT INTO success" +
                                                    " (CATE_CD " +
                                                    ", LAST_DTM )" +
                                              "VALUES(@CATE_CD " +
                                                   ", @LAST_DTM )" +
                                                   "ON DUPLICATE KEY " +
                                               "UPDATE LAST_DTM = @LAST_DTM";

                            comm.Parameters.AddWithValue("@CATE_CD", dt.Rows[i]["cateCd"]);
                            comm.Parameters.AddWithValue("@LAST_DTM", DateTime.Today.ToString("yyyyMMdd"));
                            comm.ExecuteNonQuery();
                            if (i == dt.Rows.Count - 1)
                            {

                                //memoLog.Text += "\n엑셀데이터 저장 완료";
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                logDTO.setMemo(ex.Message.ToString());
                DB_insert_log();
                //PuffAdderApplication.Client.ActiveForm.
                //memoLog.Text += "\n" + MethodBase.GetCurrentMethod().Name + "==>" + ex.Message.ToString();
            }
        }


        /// <summary>
        /// 로그 테이블 저장
        /// </summary>
        public void DB_insert_log()
        {
            string MyconnectString = "server=127.0.0.1; port=3306 ; database= investment ; uid=root ; pwd= 5587 ; SslMode=none";

            MySqlConnection conn = new MySqlConnection(MyconnectString);
            conn.Open();

            MySqlCommand comm = conn.CreateCommand();
            comm.CommandText = "INSERT INTO log" +
                                    " (LOG_SN" +
                                    ", DB_NM" +
                                    ", CATE_CD" +
                                    ", PAGE_NO" +
                                    ", TRADE_DT" +
                                    ", HTTP_ADRS" +
                                    ", FUNC_NM" +
                                    ", MEMO" +
                                    ", FRST_DTM)" +
                              "VALUES( (SELECT MAX(ll.LOG_SN) + 1 FROM log ll) " +
                                   ", @DB_NM" +
                                   ", @CATE_CD" +
                                   ", @PAGE_NO" +
                                   ", @TRADE_DT" +
                                   ", @HTTP_ADRS" +
                                   ", @FUNC_NM" +
                                   ", @MEMO" +
                                   ", @FRST_DTM) ";

            comm.Parameters.AddWithValue("@DB_NM", logDTO.getDbNm());
            comm.Parameters.AddWithValue("@CATE_CD", logDTO.getCateCd());
            comm.Parameters.AddWithValue("@PAGE_NO", logDTO.getPageNo());
            comm.Parameters.AddWithValue("@TRADE_DT", logDTO.getTradeDt());
            comm.Parameters.AddWithValue("@HTTP_ADRS", logDTO.getHttpAdrs());
            comm.Parameters.AddWithValue("@FUNC_NM", logDTO.getFuncNm());
            comm.Parameters.AddWithValue("@MEMO", logDTO.getMemo());
            comm.Parameters.AddWithValue("@FRST_DTM", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            comm.ExecuteNonQuery();
            //memoLog.Text += "\n[Exception Save]" + "==> insert in table log, check the exception";
        }


        /// <summary>
        /// DB 조회 - 공통
        /// </summary>
        public void DataTableFromDB()
        {
            string MyconnectString = "server=127.0.0.1; port=3306 ; database= investment ; uid=root ; pwd= 5587 ; SslMode=none";

            MySqlConnection conn = new MySqlConnection(MyconnectString);
            conn.Open();

            MySqlCommand selectCommand = new MySqlCommand();
            selectCommand.Connection = conn;
            selectCommand.CommandText = "SELECT * FROM catg";

            DataSet ds = new DataSet();
            MySqlDataAdapter da = new MySqlDataAdapter("SELECT *FROM categ", conn);
            da.Fill(ds);

            DataTable dt = ds.Tables[0];

            //dataGridViewExcel.DataSource = dt;

            conn.Close();
        }


        /// <summary>
        /// 성공여부 테이블 업데이트
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="cateCd"></param>
        public void UpdateDB(string table, string cateCd)
        {
            string MyconnectString = "server=127.0.0.1; port=3306 ; database= investment ; uid=root ; pwd= 5587 ; SslMode=none";

            MySqlConnection conn = new MySqlConnection(MyconnectString);
            conn.Open();
            MySqlCommand comm = conn.CreateCommand();

            switch (table)
            {
                case "price":
                    comm.CommandText = "INSERT INTO success" +
                                            " (CATE_CD" +
                                            ", PRICE_FST" +
                                            ", TRADE_FST" +
                                            ", LAST_DTM) " +
                                      "VALUES(@CATE_CD" +
                                           ", @PRICE_FST" +
                                           ", @TRADE_FST" +
                                           ", @LAST_DTM)" +
                                           "ON DUPLICATE KEY " +
                                       "UPDATE PRICE_FST = @PRICE_FST" +
                                            ", LAST_DTM  = @LAST_DTM";

                    comm.Parameters.AddWithValue("@CATE_CD", cateCd);
                    comm.Parameters.AddWithValue("@PRICE_FST", "Y");
                    comm.Parameters.AddWithValue("@TRADE_FST", "");
                    comm.Parameters.AddWithValue("@LAST_DTM", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    break;
                case "trade":
                    comm.CommandText = "INSERT INTO success" +
                                            " (CATE_CD" +
                                            ", PRICE_FST" +
                                            ", TRADE_FST" +
                                            ", LAST_DTM) " +
                                      "VALUES(@CATE_CD" +
                                           ", @PRICE_FST" +
                                           ", @TRADE_FST" +
                                           ", @LAST_DTM)" +
                                           "ON DUPLICATE KEY " +
                                       "UPDATE TRADE_FST = @TRADE_FST" +
                                            ", LAST_DTM  = @LAST_DTM";

                    comm.Parameters.AddWithValue("@CATE_CD", cateCd);
                    comm.Parameters.AddWithValue("@PRICE_FST", "");
                    comm.Parameters.AddWithValue("@TRADE_FST", "Y");
                    comm.Parameters.AddWithValue("@LAST_DTM", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    break;
            }

            comm.ExecuteNonQuery();
        }


    }
}
