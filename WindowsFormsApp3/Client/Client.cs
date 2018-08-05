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
    public partial class Client : Form
    {


        #region ======================== 초기화 부 ========================


        public Client()
        {
            InitializeComponent();
        }


        /// <summary>
        /// 전역 변수 선언
        /// </summary>
        DataTable dtSave = new DataTable();
        PuffAdderApplication.DTO.LogDTO logDTO = new PuffAdderApplication.DTO.LogDTO();


        /// <summary>
        /// 로드 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void httpToDB_Load(object sender, EventArgs e)
        {
            //InitializeProgressBar();
        }


        /// <summary>
        /// 프로그래스바 초기화
        /// </summary>
        public void InitializeProgressBar(int step)
        {
            progressBarMain.Style = ProgressBarStyle.Continuous;
            progressBarMain.Minimum = 0;
            progressBarMain.Maximum = 100;
            progressBarMain.Step = step;
            progressBarMain.Value = 0;
        }


        /// <summary>
        /// 데이터테이블 초기화
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="kind"></param>
        /// <returns></returns>
        public DataTable InitializeDataTable(DataTable dt, string kind)
        {
            if(kind == "price")
            {
                dt.Columns.Add("cateCd", typeof(string));
                dt.Columns.Add("tradeDt", typeof(string));
                dt.Columns.Add("stPrice", typeof(int));
                dt.Columns.Add("highPrice", typeof(int));
                dt.Columns.Add("lowPrice", typeof(int));
                dt.Columns.Add("endPrice", typeof(int));
                dt.Columns.Add("tradeSum", typeof(int));
                dt.Columns.Add("antSum", typeof(int));
                dt.Columns.Add("orgSum", typeof(int));
                dt.Columns.Add("forSum", typeof(int));
            }
            else if(kind == "categ")
            {
                dt.Columns.Add("cateCd", typeof(string));
                dt.Columns.Add("cateNm", typeof(string));
                dt.Columns.Add("market", typeof(string));
                dt.Columns.Add("induBasic", typeof(string));
                dt.Columns.Add("induDtl", typeof(string));
                dt.Columns.Add("valdYn", typeof(string));
                dt.Columns.Add("stDt", typeof(string));
                dt.Columns.Add("endDt", typeof(string));
            }
            else if(kind == "trade")
            {
                dt.Columns.Add("cateCd", typeof(string));
                dt.Columns.Add("tradeDt", typeof(string));
                dt.Columns.Add("orgSum", typeof(int));
                dt.Columns.Add("forSum", typeof(int));
                dt.Columns.Add("forHaveCnt", typeof(int));
                dt.Columns.Add("forHavePnt", typeof(string));
            }
            //else if(kind == "finan")
            //{
            //    dt.Columns.Add("cateCd", typeof(string));
            //    dt.Columns.Add("termClsfCd", typeof(string));
            //    dt.Columns.Add("aggrMon", typeof(string));
            //    //dt.Columns.Add("sales", typeof());
            //    //dt.Columns.Add("oprPrf", typeof(int));
            //    //dt.Columns.Add("netPrf", typeof(string));
            //}
            else if (kind == "point")
            {
                dt.Columns.Add("cateCd", typeof(string));
                dt.Columns.Add("aggrDt", typeof(string));
                dt.Columns.Add("per", typeof(float));
                dt.Columns.Add("induPer", typeof(float));
                dt.Columns.Add("pbr", typeof(float));
                dt.Columns.Add("divRate", typeof(float));
            }
            return dt;
        }


        #endregion



        #region ======================== 이벤트 부 ========================



        #region ### 종목코드별

        /// <summary>
        /// 종목코드별 저장 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonByCode_Click(object sender, EventArgs e)
        {
            int step = 1;
            InitializeProgressBar(step);
            lblProgTitle.Text = "종목코드별 저장 실행 중";
            lblProgPercent.Text = "진행률(00%)";
            textProgCateCode.Text = textCode.Text;

            switch (cmbInfo1.SelectedItem.ToString())
            {
                case "일별주가":
                    OnAction_PriceByCode();
                    break;
                case "종목기본":
                    OnAction_CategByCode();
                    break;
                case "일별거래량":
                    OnAction_TradeByCode();
                    break;
                case "투자지표":
                    OnAction_PointByCode();
                    break;
                default:
                    MessageBox.Show(cmbInfo1.SelectedItem.ToString() + " 실행은 아직 준비되지 않은 기능입니다.");
                    break;
            }

            textProgCateCode.Text = "";
            lblProgTitle.Text = "저장완료";
            lblProgPercent.Text = "진행률(100%)";
            textToPage1.Text = "";
        }

        #endregion



        #region ### 마켓별

        /// <summary>
        /// 마켓별 저장 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Click(object sender, EventArgs e)
        {
            lblProgPercent.Text = "진행률(00%)";
            textProgCateCode.Text = textCode.Text;

            if (rdoAllMarket.Checked == true)
            {
                lblProgTitle.Text = "전종목 " + cmbInfo2.SelectedItem.ToString() + " 저장 실행 중";
                insertAction(textKospi.Text.ToString());
                insertAction(textKosdaq.Text.ToString());
            }
            else if (rdoKospi.Checked == true)
            {
                lblProgTitle.Text = "KOSPI " + cmbInfo2.SelectedItem.ToString() + " 저장 실행 중";
                insertAction(textKospi.Text.ToString());
            }
            else if (rdoKosdaq.Checked == true)
            {
                lblProgTitle.Text = "KOSDAQ " + cmbInfo2.SelectedItem.ToString()+ " 저장 실행 중";
                insertAction(textKosdaq.Text.ToString());
            }

            textProgCateCode.Text = "";
            lblProgTitle.Text = "저장완료";
            lblProgPercent.Text = "진행률(100%)";
            textToPage1.Text = "";


        }

        #endregion



        #region ### 인덱스별
        #endregion



        #region ### 추가기능

        /// <summary>
        /// HTML 코드조회 실행 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonHTML_Click(object sender, EventArgs e)
        {
            memoLog.Text += "\n" + searchHtmlCode(httpText.Text.ToString());
        }


        /// <summary>
        /// 엑셀파일 조회버튼 실행 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonExcel_Click(object sender, EventArgs e)
        {
            if (rdoAllMarket.Checked == true)
            {
                DataTable dtKospi = ReadExcelData(textKospi.Text.ToString());
                DataTable dtKosdaq = ReadExcelData(textKosdaq.Text.ToString());
            }
            else if (rdoKospi.Checked == true)
            {
                DataTable dtKospi = ReadExcelData(textKospi.Text.ToString());
            }
            else if (rdoKosdaq.Checked == true)
            {
                DataTable dtKosdaq = ReadExcelData(textKosdaq.Text.ToString());
            }
        }


        /// <summary>
        /// 최종저장일 찾기 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textCode_TextChanged(object sender, EventArgs e)
        {
            if (textCode.TextLength == 6 && cmbInfo1.SelectedItem != null && cmbInfo1.SelectedItem.ToString() != "")
            {
                OnSearchLastSaveDate();
            }
        }


        /// <summary>
        /// 최종저장일 찾기 이벤트
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbInfo1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (textCode.Text.ToString() != "" && textCode.TextLength == 6)
            {
                OnSearchLastSaveDate();
            }
        }


        #endregion



        #endregion



        #region ======================== 메인 로직 ========================



        #region ### 종목코드별


        /// <summary>
        /// 종목코드 - 일별주가 저장 실행
        /// </summary>
        private void OnAction_PriceByCode()
        {
            int page = Convert.ToInt32(textFromPage1.Text);
            string httpAddressFirst = "http://finance.naver.com/item/sise_day.nhn?code=" + textCode.Text.ToString() + "&page=1";

            if (textToPage1.Text.ToString() == "")
            {
                textToPage1.Text = parseLastPage("price", httpAddressFirst, textCode.Text.ToString()).ToString();
            }

            int lastPage = Convert.ToInt32(textToPage1.Text);

            for (page = Convert.ToInt32(textFromPage1.Text); page <= lastPage; page++)
            {
                textProgTempPage.Text = page.ToString();
                string httpAddress = "http://finance.naver.com/item/sise_day.nhn?code=" + textCode.Text.ToString() + "&page=" + page.ToString();
                DataTable dtParse = parseHtml_price(httpAddress, textCode.Text.ToString(), page);
                DB_insert("price", dtParse);

                if (page == lastPage)
                {
                    UpdateDB("trade", textCode.Text.ToString());
                    memoLog.Text += "\n[Complete] => " + textCode.Text.ToString() + " until last " + lastPage.ToString() + "p.";
                }
            }

        }


        /// <summary>
        /// 종목코드 - 종목기본 저장 실행
        /// </summary>
        private void OnAction_CategByCode()
        {
            string httpAddress = "https://companyinfo.stock.naver.com/v1/company/c1010001.aspx?cmp_cd=" + textCode.Text.ToString();
            DataTable dtParse_Categ = parseHtml_categ(httpAddress, textCode.Text.ToString());
            DB_insert("categ", dtParse_Categ);
        }


        /// <summary>
        /// 종목코드 - 거래량 저장 실행 
        /// </summary>
        private void OnAction_TradeByCode()
        {
            int page = Convert.ToInt32(textFromDate1.Text);
            string httpAddressFirst = "https://finance.naver.com/item/frgn.nhn?code=" + textCode.Text.ToString() + "&page=1";
            if (textToPage1.Text.ToString() == "")
            {
                textToPage1.Text = parseLastPage("trade", httpAddressFirst, textCode.Text.ToString()).ToString();
            }

            int lastPage = Convert.ToInt32(textToPage1.Text);
            memoLog.Text += "\n마지막 페이지: " + lastPage.ToString();
            for (page = Convert.ToInt32(textFromDate1.Text); page <= lastPage; page++)
            {
                string httpAddress = "https://finance.naver.com/item/frgn.nhn?code=" + textCode.Text.ToString() + "&page=" + page.ToString();
                DataTable dtParse = parseHtml_trade(httpAddress, textCode.Text.ToString(), page);
                DB_insert("trade", dtParse);

                if (page == lastPage)
                {
                    UpdateDB("trade", textCode.Text.ToString());
                    memoLog.Text += "\n[Complete] => " + textCode.Text.ToString() + " until last " + lastPage.ToString() + "p.";
                }
            }

        }


        private void OnAction_PointByCode()
        {
            string httpAddress = "https://finance.naver.com/item/coinfo.nhn?code=" + textCode.Text.ToString();
            DataTable dtParse_Point = parseHtml_point(httpAddress, textCode.Text.ToString());
            DB_insert("point", dtParse_Point);

        }


        #endregion



        #region ### 마켓별

        /// <summary>
        /// 마켓별 저장 실행부
        /// </summary>
        /// <param name="path"></param>
        private void insertAction(string path)
        {
            InitializeProgressBar(1);
            DataTable dtExcel = ReadExcelData(path);

            if (cmbInfo2.SelectedItem.ToString() == "일별주가")
            {
                string httpAddressFront = "http://finance.naver.com/item/sise_day.nhn?";
                foreach (DataRow drExcel in dtExcel.Rows)
                {
                    string code = drExcel["cateCd"].ToString();
                    textProgCateName.Text = drExcel["cateNm"].ToString();

                    // 시작페이지 설정
                    int FromPage = 1;
                    if (textFromPage2.Text != null)
                    {
                        FromPage = Convert.ToInt32(textFromPage2.Text);
                    }
                    string httpAddressFirst = httpAddressFront + "code=" + code + "&page=" + FromPage.ToString();

                    // 종료페이지 설정
                    int ToPage = 1;
                    int ParseToPage = Convert.ToInt32(parseLastPage("price", httpAddressFirst, code));
                    if (textToPage2.Text != null)
                    {
                        if (Convert.ToInt32(textToPage2.Text) > ParseToPage)
                        {
                            ToPage = ParseToPage;
                        }
                        else
                        {
                            ToPage = Convert.ToInt32(textToPage2.Text);
                        }
                    }
                    else
                    {
                        ToPage = ParseToPage;
                    }

                    for (int pageNow = FromPage; pageNow <= ToPage; pageNow++)
                    {
                        textProgTempPage.Text = pageNow.ToString();
                        string httpAddressFull = httpAddressFront + "code=" + code + "&page=" + pageNow.ToString();
                        DataTable dtParse = parseHtml_price(httpAddressFull, code, pageNow);
                        DB_insert("price", dtParse);

                        if (pageNow == ToPage)
                        {
                            UpdateDB("price", code);
                            memoLog.Text += "\n[Complete] => " + code + " until last " + ToPage.ToString() + "p.";
                        }
                    }
                }
            }
            else if (cmbInfo2.SelectedItem.ToString() == "종목기본")
            {
                string httpAddressFront = "https://companyinfo.stock.naver.com/v1/company/c1010001.aspx?cmp_cd=";

                int i = 0;
                foreach (DataRow drExcel in dtExcel.Rows)
                {
                    string code = drExcel["cateCd"].ToString();
                    textProgCateName.Text = drExcel["cateNm"].ToString();
                    string httpAddress = httpAddressFront + code;
                    DataTable dtParse_Categ = parseHtml_categ(httpAddress, code);
                    DB_insert("categ", dtParse_Categ);
                    i++;
                }
            }
            else if (cmbInfo2.SelectedItem.ToString() == "일별거래량")
            {
                string httpAddressFront = "https://finance.naver.com/item/frgn.nhn?";
                foreach (DataRow drExcel in dtExcel.Rows)
                {
                    string code = drExcel["cateCd"].ToString();
                    textProgCateName.Text = drExcel["cateNm"].ToString();

                    int FromPage = 1;
                    if (textFromPage2.Text != null)
                    {
                        FromPage = Convert.ToInt32(textFromPage2.Text);
                    }

                    string httpAddressFirst = httpAddressFront + "code=" + code + "&page=" + FromPage.ToString();

                    int ToPage = 1;
                    if(textFromPage2.Text == textToPage2.Text)
                    {
                        ToPage = FromPage;
                    }
                    else
                    {
                        int ParseToPage = Convert.ToInt32(parseLastPage("trade", httpAddressFirst, code));
                        if (textToPage2.Text != null)
                        {
                            if (Convert.ToInt32(textToPage2.Text) > ParseToPage)
                            {
                                ToPage = ParseToPage;
                            }
                            else
                            {
                                ToPage = Convert.ToInt32(textToPage2.Text);
                            }
                        }
                        else
                        {
                            ToPage = ParseToPage;
                        }
                    }

                    for (int pageNow = FromPage; pageNow <= ToPage; pageNow++)
                    {
                        textProgTempPage.Text  = pageNow.ToString();
                        string httpAddressFull = httpAddressFront + "code=" + code + "&page=" + pageNow.ToString();
                        DataTable dtParse_Trade = parseHtml_trade(httpAddressFull, code, pageNow);
                        DB_insert("trade", dtParse_Trade);

                        if (pageNow == ToPage)
                        {
                            UpdateDB("trade", code);
                            memoLog.Text += "\n[Complete] => " + code + " until last " + ToPage.ToString() + "p.";
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("아직 미개발");
            }

        }

        #endregion



        #region ### 인덱스별
        #endregion



        #endregion



        #region ======================== HTML 추출부 ========================


        /// <summary>
        /// 종목코드로 HTML 조회
        /// </summary>
        private string searchHtmlCode(string httpAddress)
        {
            int euckrCodepage = 51949;
            string sHtml = string.Empty;

            try
            {
                HttpWebRequest oRequest = (HttpWebRequest)WebRequest.Create(httpAddress);
                HttpWebResponse oGetResponse = (HttpWebResponse)oRequest.GetResponse();
                Encoding encode;

                switch(oGetResponse.CharacterSet.ToLower())
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
                memoLog.Text += "\n" + MethodBase.GetCurrentMethod().Name + "==>" + ex.Message.ToString();
                throw ex;
            }
            return sHtml;
        }

        
        /// <summary>
        /// 일별주가 파싱
        /// </summary>
        /// <param name="htmlCode"></param>
        /// <returns></returns>
        private DataTable parseHtml_price(string httpAddress, string code, int page)
        {
            DataTable dtPrice = new DataTable();
            string table = "price";
            dtPrice = InitializeDataTable(dtPrice, table);

            logDTO.setCateCd(code);
            logDTO.setDbNm(table);
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

                    if(row.InnerText == "0" && i % 5 == 1)
                    {
                        continue;
                    }
                    else if(row.InnerText == "0" && i % 5 == 4)
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
            catch(Exception ex)
            {
                logDTO.setMemo(ex.Message.ToString());
                memoLog.Text += "\nMethodName =" + MethodBase.GetCurrentMethod().Name + ", code= " + code +", page= " + page.ToString() + "==>" + ex.Message.ToString();
                DB_insert_log();
                // logDB 에 저장
            }

            return dtPrice;
        }


        /// <summary>
        /// 종목기본 파싱
        /// </summary>
        /// <param name="httpAddress"></param>
        /// <param name="code"></param>
        /// <returns></returns>
        private DataTable parseHtml_categ(string httpAddress, string code)
        {
            DataTable dtCateg = new DataTable();
            string table = "categ";
            dtCateg = InitializeDataTable(dtCateg, table);

            logDTO.setCateCd(code);
            logDTO.setDbNm(table);
            logDTO.setHttpAdrs(httpAddress);
            logDTO.setFuncNm(MethodBase.GetCurrentMethod().Name.ToString());

            try
            {
                HtmlAgilityPack.HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument doc = web.Load(httpAddress);

                DataRow newRow;
                int j = 0;
                foreach (HtmlNode row in doc.DocumentNode.SelectNodes("//td[@class='cmp-table-cell td0101']"))
                {
                    if (j > 0)
                        continue;

                    newRow = dtCateg.NewRow();
                    int n = 0;
                    foreach(HtmlNode rowDtl in row.SelectNodes("//span[@class='name']"))
                    {
                        if(n == 0)
                        {
                            string cateNm = rowDtl.InnerText.ToString().Replace(" ", "");
                            newRow["cateNm"] = cateNm;
                            newRow["cateCd"] = code;
                            dtCateg.Rows.Add(newRow);
                        }
                        n++;
                    }

                    int m = 0;
                    foreach(HtmlNode rowDtl in row.SelectNodes("//dt[@class='line-left']"))
                    {
                        if(m == 1)
                        {
                            string[] arr = rowDtl.InnerText.ToString().Split(':');
                            newRow["market"] = arr[0].Trim().ToString();
                            newRow["induBasic"] = arr[1].Trim().ToString();
                        }
                        else if(m == 2)
                        {
                            string[] arr = rowDtl.InnerText.ToString().Split(':');
                            newRow["induDtl"] = arr[1].Trim().ToString();
                        }
                        m++;
                    }
                    j++;
                }
            }
            catch (Exception ex)
            {
                logDTO.setMemo(ex.Message.ToString());
                memoLog.Text += "\n[categ] MethodName =" + MethodBase.GetCurrentMethod().Name + ", code= " + code +  "==>" + ex.Message.ToString();
                DB_insert_log();
            }

            return dtCateg;
        }


        /// <summary>
        /// 일별거래량 파싱
        /// </summary>
        /// <param name="httpAddress"></param>
        /// <param name="code"></param>
        /// <returns></returns>
        private DataTable parseHtml_trade(string httpAddress, string code, int page)
        {
            DataTable dtTrade = new DataTable();
            string table = "trade";
            dtTrade = InitializeDataTable(dtTrade, table);

            logDTO.setCateCd(code);
            logDTO.setDbNm(table);
            logDTO.setHttpAdrs(httpAddress);
            logDTO.setFuncNm(MethodBase.GetCurrentMethod().Name.ToString());
            logDTO.setPageNo(page);

            try
            {
                HtmlAgilityPack.HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument doc = web.Load(httpAddress);

                DataRow newRow;

                foreach (HtmlNode row1 in doc.DocumentNode.SelectNodes("//div[@id='wrap']"))
                {
                    foreach(HtmlNode row2 in row1.SelectNodes("//div[@id='middle']"))
                    {
                        foreach(HtmlNode row3 in row2.SelectNodes("//div[@id='content']"))
                        {
                            foreach (HtmlNode row4 in row3.SelectNodes("//div[@class='section inner_sub']"))
                            {
                                foreach(HtmlNode row5 in row4.SelectNodes("//td[@class='tc']"))
                                {
                                    newRow = dtTrade.NewRow();
                                    newRow["tradeDt"] = row5.InnerText.Trim().Replace(".", "").ToString();
                                    newRow["cateCd"] = code;
                                    dtTrade.Rows.Add(newRow);
                                }
                                int i = 0;
                                foreach (HtmlNode row6 in row4.SelectNodes("//td[@class='num']"))
                                {
                                    int r = i / 8;
                                    switch (i % 8)
                                    {
                                        case 0: break; //nothing
                                        case 1: break; //nothing 
                                        case 2: break; //nothing
                                        case 3: break; //nothing
                                        case 4:
                                            dtTrade.Rows[r]["orgSum"] = Convert.ToInt32(row6.InnerText.Trim().Replace(",",""));
                                            break; 
                                        case 5:
                                            dtTrade.Rows[r]["forSum"] = Convert.ToInt32(row6.InnerText.Trim().Replace(",", ""));
                                            break;
                                        case 6:
                                            dtTrade.Rows[r]["forHaveCnt"] = Convert.ToInt32(row6.InnerText.Trim().Replace(",", ""));
                                            break;
                                        case 7:
                                            dtTrade.Rows[r]["forHavePnt"] = row6.InnerText.Trim().Replace("%", "").ToString();
                                            break;
                                    }
                                    i++;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logDTO.setMemo(ex.Message.ToString());
                memoLog.Text += "\n[trade] MethodName =" + MethodBase.GetCurrentMethod().Name + ", code= " + code + "==>" + ex.Message.ToString();
                DB_insert_log();
            }
            return dtTrade;
        }


        /// <summary>
        /// 재무상태 파싱 (개발중)
        /// </summary>
        /// <param name="httpAddress"></param>
        /// <param name="code"></param>
        /// <returns></returns>
        private DataTable parseHtml_finan(string httpAddress, string code)
        {
            DataTable dtFinan = new DataTable();
            string table = "point";
            return dtFinan;
        }


        /// <summary>
        /// 투자지표 파싱 (개발중)
        /// </summary>
        /// <param name="httpAddress"></param>
        /// <param name="code"></param>
        /// <returns></returns>
        private DataTable parseHtml_point(string httpAddress, string code)
        {
            DataTable dtPoint = new DataTable();
            //string table = "point";
            //dtPoint = InitializeDataTable(dtPoint, table);

            //logDTO.setCateCd(code);
            //logDTO.setDbNm(table);
            //logDTO.setHttpAdrs(httpAddress);
            //logDTO.setFuncNm(MethodBase.GetCurrentMethod().Name.ToString());

            //try
            //{
            //    HtmlAgilityPack.HtmlWeb web = new HtmlWeb();
            //    HtmlAgilityPack.HtmlDocument doc = web.Load(httpAddress);
            //    //DataRow newRow;

            //    foreach (HtmlNode row1 in doc.DocumentNode.SelectNodes("//div[@id='wrap']"))
            //    {
            //        foreach (HtmlNode row2 in row1.SelectNodes("//div[@id='middle']"))
            //        {
            //            foreach (HtmlNode row3 in row2.SelectNodes("//div[@id='content']"))
            //            {
            //                //foreach (HtmlNode row4 in row3.SelectNodes("//div[@class='section inner_sub']"))
            //                //{
            //                    //foreach (HtmlNode row5 in row4.SelectNodes("//div[@class='body-section']"))
            //                    //{
            //                        foreach (HtmlNode row6 in row3.SelectNodes("//div[@id='all_contentWrap']"))
            //                        {
            //                            foreach (HtmlNode row7 in row6.SelectNodes("//div[@id='contentWrap']"))
            //                            {
            //                                foreach (HtmlNode row8 in row7.SelectNodes("//div[@id='pArea']"))
            //                                {
            //                                    foreach (HtmlNode row9 in row8.SelectNodes("//div[@class='wrapper-table']"))
            //                                    {
            //                                        foreach (HtmlNode row10 in row9.SelectNodes("//div[@class='c-table-div']"))
            //                                        {
            //                                            foreach (HtmlNode row11 in row10.SelectNodes("//td[@class='cmp-table-cell td0301']"))
            //                                            {
            //                                                memoLog.Text = "\n" + row11.InnerHtml;
            //                                            }
            //                                        }
            //                                    }
            //                                }
            //                            }
            //                        //}
            //                    //}
            //                }
            //            }
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    logDTO.setMemo(ex.Message.ToString());
            //    memoLog.Text += "\nMethodName =" + MethodBase.GetCurrentMethod().Name + ", code= " + code + "==>" + ex.Message.ToString();
            //    DB_insert_log();
            //}

            return dtPoint;
        }


        /// <summary>
        /// 마지막페이지 파싱
        /// </summary>
        /// <param name="httpAddress"></param>
        /// <param name="code"></param>
        /// <returns></returns>
        private int parseLastPage(string kind, string httpAddress, string code)
        {
            int lastPage = 0;
            try
            {
                HtmlAgilityPack.HtmlWeb web = new HtmlWeb();
                HtmlAgilityPack.HtmlDocument doc = web.Load(httpAddress);

                string result = string.Empty;
                if (kind == "price")
                {
                    foreach (HtmlNode row in doc.DocumentNode.SelectNodes("//td[@class='pgRR']"))
                    {
                        List<string> hrefTags = ExtractAllAHrefTags(row);
                        int size = hrefTags.Count;
                        result = hrefTags[size - 1];
                    }

                    string[] arr = result.Split('=');
                    int arrlen = arr.Length;
                    lastPage = Convert.ToInt32(arr[arrlen - 1]);
                }
                else if (kind == "trade")
                {
                    foreach (HtmlNode row in doc.DocumentNode.SelectNodes("//td[@class='pgRR']"))
                    {
                        List<string> hrefTags = ExtractAllAHrefTags(row);
                        result = hrefTags[68];
                    }
                    string[] arr = result.Split('=');
                    int arrlen = arr.Length;
                    lastPage = Convert.ToInt32(arr[arrlen - 1]);
                }
            }
            catch(Exception ex)
            {
                memoLog.Text += "\n" + MethodBase.GetCurrentMethod().Name + "==>" + ex.Message.ToString();
                // logDB 에 저장
            }
            return lastPage;
        }


        private List<string> ExtractAllAHrefTags(HtmlAgilityPack.HtmlNode htmlSnippet)
        {
            List<string> hrefTags = new List<string>();
            try
            {
                foreach (HtmlNode link in htmlSnippet.SelectNodes("//a[@href]"))
                {
                    HtmlAttribute att = link.Attributes["href"];
                    hrefTags.Add(att.Value);
                }
            }
            catch(Exception ex)
            {
                memoLog.Text += "\n" + MethodBase.GetCurrentMethod().Name + "==>" + ex.Message.ToString();
                throw ex;
            }
            return hrefTags;
        }


        #endregion



        #region ======================== D B 연동부 ========================



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
                            textProgTempDate.Text = dt.Rows[i]["tradeDt"].ToString();

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
                            if(dt.Rows[i]["tradeDt"] == null || dt.Rows[i]["tradeDt"].ToString() == "" || dt.Rows[i]["tradeDt"].ToString() == "0")
                            {
                                return;
                            }
                            textProgTempDate.Text = dt.Rows[i]["tradeDt"].ToString();
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
                        for(int i = 0; i < dt.Rows.Count; i++)
                        {
                            MySqlCommand comm = conn.CreateCommand();
                            comm.CommandText = "INSERT INTO success" +
                                                    " (CATE_CD " +
                                                    ", LAST_DTM )" +
                                              "VALUES(@CATE_CD " +
                                                   ", @LAST_DTM )" +
                                                   "ON DUPLICATE KEY " +
                                               "UPDATE LAST_DTM = @LAST_DTM" ;

                            comm.Parameters.AddWithValue("@CATE_CD", dt.Rows[i]["cateCd"]);
                            comm.Parameters.AddWithValue("@LAST_DTM", DateTime.Today.ToString("yyyyMMdd"));
                            comm.ExecuteNonQuery();
                            if(i == dt.Rows.Count -1)
                            {
                                memoLog.Text += "\n엑셀데이터 저장 완료";
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                logDTO.setMemo(ex.Message.ToString());
                DB_insert_log();
                memoLog.Text += "\n" + MethodBase.GetCurrentMethod().Name + "==>" + ex.Message.ToString();
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
            memoLog.Text += "\n[Exception Save]" + "==> insert in table log, check the exception";
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


        /// <summary>
        /// 마지막페이지 찾기 쿼리
        /// </summary>
        /// <param name="table"></param>
        /// <param name="cateCd"></param>
        public void SelectLastPageDB(string table, string cateCd)
        {
            string MyconnectString = "server=127.0.0.1; port=3306 ; database= investment ; uid=root ; pwd= 5587 ; SslMode=none";
            MySqlConnection conn = null;
            MySqlDataReader reader = null;
            DataTable dt = new DataTable();

            try
            {
                conn = new MySqlConnection(MyconnectString);
                conn.Open();

                MySqlCommand comm = conn.CreateCommand();
                comm.CommandText = "SELECT MAX(TRADE_DT) FROM " + table +
                                   " WHERE CATE_CD = '" + cateCd + "'";

                MySqlDataAdapter da = new MySqlDataAdapter();
                da.SelectCommand = new MySqlCommand(comm.CommandText, conn);
                da.Fill(dt);
                if (dt == null || dt.Rows[0][0].ToString() == "")
                    textLastDate1.Text = "없음";
                textLastDate1.Text = dt.Rows[0][0].ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (reader != null)
                    reader.Close();
                if(conn != null)
                {
                    conn.Close();
                }
            } 
        }


        #endregion



        #region ======================== 엑셀 연동부 ========================


        /// <summary>
        /// 엑셀 데이터 가져오는 함수
        /// </summary>
        /// <param name="path"></param>
        public DataTable ReadExcelData(string path)
        {
            dtSave = null;
            // path는 Excel파일의 전체 경로입니다.
            // 예. D:\test\test.xslx
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                excelApp = new Excel.Application();
                wb = excelApp.Workbooks.Open(path);
                // path 대신 문자열도 가능합니다
                // 예. Open(@"D:\test\test.xslx");
                ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;
                // 첫번째 Worksheet를 선택합니다.
                Excel.Range rng = ws.UsedRange;   // '여기'
                                                  // 현재 Worksheet에서 사용된 셀 전체를 선택합니다.
                object[,] data = rng.Value;

                string market = string.Empty;
                if(path.Trim() == "")
                {
                    MessageBox.Show("엑셀파일 경로에 값을 입력해야 합니다.");
                    return null;
                }
                else if (path.Contains("KOSPI"))
                    market = "KOSPI";
                else if (path.Contains("KOSDAQ"))
                    market = "KOSDAQ";
                else
                {
                    MessageBox.Show("엑셀파일명은 KOSPI 또는 KOSDAQ으로 지어야 합니다.");
                    return null;
                }

                /* 편의상 데이터테이블로 변환 */
                DataTable dt = ObjectToDataTable(data, market);
                //dataGridViewExcel.DataSource = dt;

                dtSave = dt;

                return dt;
            }
            catch (Exception ex)
            {
                memoLog.Text += "\n" + MethodBase.GetCurrentMethod().Name + "==>" + ex.Message.ToString();

                throw ex;
                // logDB 에 저장
            }
            finally
            {
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(excelApp);
            }
        }


        /// <summary>
        /// Object -> 데이터테이블 형변환
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="market"></param>
        /// <returns></returns>
        public DataTable ObjectToDataTable(object[,] obj, string market)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("cateNm", typeof(string));
            dt.Columns.Add("cateCd", typeof(string));
            dt.Columns.Add("market", typeof(string));

            int rowCnt = obj.GetLength(0);
            int colCnt = obj.GetLength(1);

            DataRow dr;

            for (int i = 2; i <= rowCnt; i++)
            {
                dr = dt.NewRow();
                dr["cateCd"] = obj[i, 2];
                dr["cateNm"] = obj[i, 3];
                dr["market"] = market;
                dt.Rows.Add(dr);
            }
            return dt;
        }


        /// <summary>
        /// 엑셀 관련
        /// </summary>
        /// <param name="obj"></param>
        static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
                // logDB 에 저장
            }
            finally
            {
                GC.Collect();
            }
        }



        #endregion



        /// <summary>
        /// 최종 저장일 찾기 실행 로직
        /// </summary>
        private void OnSearchLastSaveDate()
        {
            if(cmbInfo1.SelectedItem.ToString() == "일별주가")
            {
                SelectLastPageDB("price", textCode.Text.ToString());
            }
            else if(cmbInfo1.SelectedItem.ToString() == "일별거래량")
            {
                SelectLastPageDB("trade", textCode.Text.ToString());
            }
            else
            {

            }
        }

        private void rdoFromDate1_CheckedChanged(object sender, EventArgs e)
        {
            if(rdoFromDate1.Checked == true)
            {
                textFromDate1.Enabled = true;
                textFromPage1.Enabled = false;
                textToDate1.Enabled = true;
                textToPage1.Enabled = false;
            }
            else
            {
                textFromDate1.Enabled = false;
                textFromPage1.Enabled = true;
                textToDate1.Enabled = false;
                textToPage1.Enabled = true;
            }
        }

        private void rdoFromPage1_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoFromPage1.Checked == true)
            {
                textFromDate1.Enabled = false;
                textFromPage1.Enabled = true;
                textToDate1.Enabled = false;
                textToPage1.Enabled = true;
            }
            else
            {
                textFromDate1.Enabled = true;
                textFromPage1.Enabled = false;
                textToDate1.Enabled = true;
                textToPage1.Enabled = false;
            }
        }

        private void rdoFromDate2_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoFromDate2.Checked == true)
            {
                textFromDate2.Enabled = true;
                textFromPage2.Enabled = false;
                textToDate2.Enabled = true;
                textToPage2.Enabled = false;
            }
            else
            {
                textFromDate2.Enabled = false;
                textFromPage2.Enabled = true;
                textToDate2.Enabled = false;
                textToPage2.Enabled = true;
            }
        }

        private void rdoFromPage2_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoFromPage2.Checked == true)
            {
                textFromDate2.Enabled = false;
                textFromPage2.Enabled = true;
                textToDate2.Enabled = false;
                textToPage2.Enabled = true;
            }
            else
            {
                textFromDate2.Enabled = true;
                textFromPage2.Enabled = false;
                textToDate2.Enabled = true;
                textToPage2.Enabled = false;
            }

        }

        private void rdoFromDate3_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoFromDate3.Checked == true)
            {
                textFromDate3.Enabled = true;
                textFromPage3.Enabled = false;
                textToDate3.Enabled = true;
                textToPage3.Enabled = false;
            }
            else
            {
                textFromDate3.Enabled = false;
                textFromPage3.Enabled = true;
                textToDate3.Enabled = false;
                textToPage3.Enabled = true;
            }

        }

        private void rdoFromPage3_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoFromPage3.Checked == true)
            {
                textFromDate3.Enabled = false;
                textFromPage3.Enabled = true;
                textToDate3.Enabled = false;
                textToPage3.Enabled = true;
            }
            else
            {
                textFromDate3.Enabled = true;
                textFromPage3.Enabled = false;
                textToDate3.Enabled = true;
                textToPage3.Enabled = false;
            }

        }
    }




}
