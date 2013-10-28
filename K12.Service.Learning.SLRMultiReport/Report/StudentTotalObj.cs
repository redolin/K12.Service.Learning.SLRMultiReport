using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace K12.Service.Learning.SLRMultiReport
{
    class StudentTotalObj
    {
        /// <summary>
        /// 學生ID
        /// </summary>
        public string ref_student_id { get; set; }

        /// <summary>
        /// 班級ID
        /// </summary>
        public string ref_class_id { get; set; }

        /// <summary>
        /// 座號
        /// </summary>
        public string seat_no { get; set; }

        /// <summary>
        /// 學號
        /// </summary>
        public string student_number { get; set; }

        /// <summary>
        /// 狀態
        /// </summary>
        public string status { get; set; }

        /// <summary>
        /// 學生姓名
        /// </summary>
        public string student_name { get; set; }

        /// <summary>
        /// 學生性別
        /// </summary>
        public string student_gender { get; set; }


        /// <summary>
        /// 服務時數, Key為學年度跟學期(ex: "90/1"), Value為時數
        /// </summary>
        public Dictionary<string, decimal> SLRDic { get; set; }

        public StudentTotalObj(DataRow row)
        {
            ref_student_id = "" + row[0];
            ref_class_id = "" + row[1];
            seat_no = "" + row[2];
            student_number = "" + row[3];
            student_name = "" + row[4];
            status = "" + row[5];
            string 性別 = "" + row[6];
            if (性別 == "1")
                student_gender = "男";
            else if (性別 == "0")
                student_gender = "女";
            else
                student_gender = "";

            SLRDic = new Dictionary<string, decimal>();
        }
    }
}
