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

        // DB테이블명
        public void setDbNm(String dbNm)
        {
            this.dbNm = dbNm;
        }
        public String getDbNm()
        {
            return this.dbNm;
        }

        // 종목코드
        public void setCateCd(String cateCd)
        {
            this.cateCd = cateCd;
        }
        public String getCateCd()
        {
            return this.cateCd;
        }

        // 페이지번호
        public void setPageNo(int pageNo)
        {
            this.pageNo = pageNo;
        }
        public int getPageNo()
        {
            return this.pageNo;
        }

        // HTTP주소
        public void setHttpAdrs(String httpAdrs)
        {
            this.httpAdrs = httpAdrs;
        }
        public String getHttpAdrs()
        {
            return this.httpAdrs;
        }

        // 메모
        public void setMemo(String memo)
        {
            this.memo = memo;
        }
        public String getMemo()
        {
            return this.memo;
        }

        // 함수명
        public void setFuncNm(String funcNm)
        {
            this.funcNm = funcNm;
        }
        public String getFuncNm()
        {
            return this.funcNm;
        }

        // 함수명
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
