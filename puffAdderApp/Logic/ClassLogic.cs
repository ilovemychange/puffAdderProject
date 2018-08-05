using System;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Reflection;



namespace PuffAdderApplication
{
    class ClassLogic
    {
        PuffAdderApplication.DTO.LogDTO logDTO = new PuffAdderApplication.DTO.LogDTO();

        /// <summary>
        /// KOSPI & KOSDAQ 개별종목 NAVER에서 HTML 코드 조회
        /// </summary>
        public string searchHtmlCode(string httpAddress)
        {
            int euckrCodepage = 51949;
            string sHtml = string.Empty;

            try
            {
                HttpWebRequest oRequest = (HttpWebRequest)WebRequest.Create(httpAddress);
                HttpWebResponse oGetResponse = (HttpWebResponse)oRequest.GetResponse();
                Encoding encode;

                switch (oGetResponse.CharacterSet.ToLower())
                {
                    case "utf-8":
                        encode = Encoding.UTF8; break;
                    case "euc-kr":
                        encode = Encoding.GetEncoding(euckrCodepage); break;
                    default:
                        encode = Encoding.Default; break;
                }

                StreamReader oStreamReader = new StreamReader(oGetResponse.GetResponseStream(), encode);
                sHtml = oStreamReader.ReadToEnd();
            }
            catch (Exception ex)
            {
                //memoLog.Text += "\n" + MethodBase.GetCurrentMethod().Name + "==>" + ex.Message.ToString();
                throw ex;
            }
            return sHtml;
        }



        /// <summary>
        /// 일별주가 HTML 파싱 from Naver
        /// </summary>
        /// <param name="htmlCode"></param>
        /// <returns></returns>
        private DataTable parseHtml_price(string httpAddress, string code, int page)
        {
            DataTable dtPrice = new DataTable();
            dtPrice.Columns.Add("cateCd", typeof(string));
            dtPrice.Columns.Add("tradeDt", typeof(string));
            dtPrice.Columns.Add("stPrice", typeof(int));
            dtPrice.Columns.Add("highPrice", typeof(int));
            dtPrice.Columns.Add("lowPrice", typeof(int));
            dtPrice.Columns.Add("endPrice", typeof(int));
            dtPrice.Columns.Add("tradeSum", typeof(int));
            dtPrice.Columns.Add("antSum", typeof(int));
            dtPrice.Columns.Add("orgSum", typeof(int));
            dtPrice.Columns.Add("forSum", typeof(int));

            logDTO.setCateCd(code);
            logDTO.setDbNm("price");
            logDTO.setHttpAdrs(httpAddress);
            logDTO.setFuncNm(MethodBase.GetCurrentMethod().Name.ToString());
            logDTO.setPageNo(page);

            try
            {
                HtmlAgilityPack.HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument doc = web.Load(httpAddress);

                DataRow newRow;

                /* 거래일 추출 */
                int j = 0;
                foreach (HtmlNode row in doc.DocumentNode.SelectNodes("//span[@class='tah p10 gray03']"))
                {
                    newRow = dtPrice.NewRow();
                    string tradeDt_yyyymmdd = row.InnerText.ToString().Replace(".", "");
                    newRow["tradeDt"] = tradeDt_yyyymmdd;
                    dtPrice.Rows.Add(newRow);
                    j++;
                }

                /* 주가정보 추출 */
                int i = 0;
                foreach (HtmlNode row in doc.DocumentNode.SelectNodes("//span[@class='tah p11']"))
                {
                    int rowNum = i / 5;

                    if (row.InnerText == "0" && i % 5 == 1)
                    {
                        continue;
                    }
                    else if (row.InnerText == "0" && i % 5 == 4)
                    {
                        // 거래량에 0이 들어왔을 경우
                        dtPrice.Rows[rowNum]["tradeSum"] = 0;
                        i++;
                        continue;
                    }
                    dtPrice.Rows[rowNum]["cateCd"] = code;

                    switch (i % 5)
                    {
                        case 0: // 종가
                            dtPrice.Rows[rowNum]["endPrice"] = row.InnerText.ToString() == "0" ? Convert.ToInt32(row.InnerText) : Convert.ToInt32(row.InnerText.Replace(",", ""));
                            break;
                        case 1: // 시가
                            dtPrice.Rows[rowNum]["stPrice"] = row.InnerText.ToString() == "0" ? Convert.ToInt32(row.InnerText) : Convert.ToInt32(row.InnerText.Replace(",", ""));
                            break;
                        case 2: // 고가
                            dtPrice.Rows[rowNum]["highPrice"] = row.InnerText.ToString() == "0" ? Convert.ToInt32(row.InnerText) : Convert.ToInt32(row.InnerText.Replace(",", ""));
                            break;
                        case 3: // 저가
                            dtPrice.Rows[rowNum]["lowPrice"] = row.InnerText.ToString() == "0" ? Convert.ToInt32(row.InnerText) : Convert.ToInt32(row.InnerText.Replace(",", ""));
                            break;
                        case 4: // 거래량
                            dtPrice.Rows[rowNum]["tradeSum"] = row.InnerText.ToString() == "0" ? Convert.ToInt32(row.InnerText) : Convert.ToInt32(row.InnerText.Replace(",", ""));
                            break;
                    }
                    i++;
                }
            }
            catch (Exception ex)
            {
                //logDTO.setMemo(ex.Message.ToString());
                //memoLog.Text += "\nMethodName =" + MethodBase.GetCurrentMethod().Name + ", code= " + code + ", page= " + page.ToString() + "==>" + ex.Message.ToString();
                //DB_insert_log();
                // logDB 에 저장
                throw ex;
            }

            return dtPrice;
        }

    }
}
