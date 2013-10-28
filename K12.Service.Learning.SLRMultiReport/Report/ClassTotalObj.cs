using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace K12.Service.Learning.SLRMultiReport
{
    class ClassTotalObj
    {
        /// <summary>
        /// 班級ID
        /// </summary>
        public string ref_class_id { get; set; }

        /// <summary>
        /// 班級名稱
        /// </summary>
        public string class_name { get; set; }

        /// <summary>
        /// 班級序號
        /// </summary>
        public string class_index { get; set; }

        /// <summary>
        /// 年級
        /// </summary>
        public string class_grade_year { get; set; }

        /// <summary>
        /// 老師ID
        /// </summary>
        public string ref_tearch_id { get; set; }

        /// <summary>
        /// 老師姓名
        /// </summary>
        public string tearch_name { get; set; }

        /// <summary>
        /// 老師暱稱
        /// </summary>
        public string nickname { get; set; }

        ///// <summary>
        ///// 學生清單
        ///// </summary>
        //public List<StudentTotalObj> StudentObjList { get; set; }

        ///// <summary>
        ///// 學生ID清單
        ///// </summary>
        //public List<string> StudentIDList { get; set; }

        /// <summary>
        /// 學生清單, Key:學生ID; Value:學生資料
        /// </summary>
        public Dictionary<string, StudentTotalObj> StudentDic { get; set; }

        /// <summary>
        /// 學生服務時數的學年度跟學期的聯集, value沒有意義
        /// </summary>
        public Dictionary<string, int> SLRSchoolYearSemesterDic { get; set; }

        public ClassTotalObj(DataRow row)
        {
            ref_class_id = "" + row[0];
            class_name = "" + row[1];
            ref_tearch_id = "" + row[2];
            tearch_name = "" + row[3];
            nickname = "" + row[4];
            class_grade_year = "" + row[5];
            class_index = "" + row[6];

            //StudentObjList = new List<StudentTotalObj>();
            //StudentIDList = new List<string>();
            StudentDic = new Dictionary<string, StudentTotalObj>();
            SLRSchoolYearSemesterDic = new Dictionary<string,int>();
        }

        public void SetSLRInStudent(SLRecord slr)
        {
            if (StudentDic.ContainsKey(slr.RefStudentID))
            {
                string sKey = slr.SchoolYear + "/" + slr.Semester;
                Dictionary<string, decimal> SLRDic = StudentDic[slr.RefStudentID].SLRDic;

                // 假如找到同一個學年度跟學期, 就把服務時間加總
                if (SLRDic.ContainsKey(sKey))
                {
                    SLRDic[sKey] += slr.Hours;
                }
                // 沒有找到同一個學年度跟學期, 就把資料新增進去
                else
                {
                    SLRDic.Add(sKey, slr.Hours);
                }

                // 所有學生的學年度跟學期的聯集
                if(!this.SLRSchoolYearSemesterDic.ContainsKey(sKey))
                {
                    this.SLRSchoolYearSemesterDic.Add(sKey, 0);
                }
            }
        }

        /// <summary>
        /// 取得老師名稱
        /// (包含暱稱)
        /// </summary>
        public string GetTeacherName()
        {
            if (!string.IsNullOrEmpty(tearch_name))
            {
                if (!string.IsNullOrEmpty(nickname))
                {
                    return tearch_name + "(" + nickname + ")";
                }
                else
                {
                    return tearch_name;
                }
            }
            else
                return "";
        }
    }
}
