using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PuffAdderApplication.DTO
{
    class LogDTO
    {
        private String dbNm;          // DB테이블명
        private String cateCd;        // 종목코드
        private int pageNo;           // 페이지번호
        private String httpAdrs;      // HTTP주소
        private String memo;          // 메모
        private String funcNm;        // 함수명
        private String tradeDt;       // 거래일

        /// <summary>
        /// 테이블명
        /// </summary>
        public void setDbNm(String dbNm)
        {
            this.dbNm = dbNm;
        }
        public String getDbNm()
        {
            return this.dbNm;
        }

        /// <summary>
        /// 종목코드
        /// </summary>
        /// <param name="cateCd"></param>
        public void setCateCd(String cateCd)
        {
            this.cateCd = cateCd;
        }
        public String getCateCd()
        {
            return this.cateCd;
        }

        /// <summary>
        /// 페이지번호
        /// </summary>
        /// <param name="pageNo"></param>
        public void setPageNo(int pageNo)
        {
            this.pageNo = pageNo;
        }
        public int getPageNo()
        {
            return this.pageNo;
        }

        /// <summary>
        /// HTTP주소
        /// </summary>
        /// <param name="httpAdrs"></param>
        public void setHttpAdrs(String httpAdrs)
        {
            this.httpAdrs = httpAdrs;
        }
        public String getHttpAdrs()
        {
            return this.httpAdrs;
        }

        /// <summary>
        /// 메모내용
        /// </summary>
        /// <param name="memo"></param>
        public void setMemo(String memo)
        {
            this.memo = memo;
        }
        public String getMemo()
        {
            return this.memo;
        }

        /// <summary>
        /// 함수명
        /// </summary>
        /// <param name="funcNm"></param>
        public void setFuncNm(String funcNm)
        {
            this.funcNm = funcNm;
        }
        public String getFuncNm()
        {
            return this.funcNm;
        }

        /// <summary>
        /// 거래일
        /// </summary>
        /// <param name="tradeDt"></param>
        public void setTradeDt(String tradeDt)
        {
            this.tradeDt = tradeDt;
        }
        public String getTradeDt()
        {
            return this.tradeDt;
        }

    }
}
