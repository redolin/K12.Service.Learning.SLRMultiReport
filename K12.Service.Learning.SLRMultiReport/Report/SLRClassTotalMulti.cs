using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Presentation;
using System.Diagnostics;
using Aspose.Words;

namespace K12.Service.Learning.SLRMultiReport
{
    public partial class SLRClassTotalMulti : BaseForm
    {
        // 建立一個背景處理物件
        BackgroundWorker _BGW = new BackgroundWorker();

        // 查詢UDT用的物件
        FISCA.UDT.AccessHelper _accessHelper = new FISCA.UDT.AccessHelper();
        FISCA.Data.QueryHelper _queryHelper = new FISCA.Data.QueryHelper();

        // 輸出word的樣板
        string _SLRClassTotalMulti_ReportCofig = "K12.Service.Learning.Modules.Config.SLRClassTotalMulti.cs";
        int _MAX_COLUMN_COUNT = 6;

        Document _doc = new Document(); // 主文件
        Run _run; // 移動使用
        Document _template;

        string _WarningMsg = "座號:{0} 姓名:{1} , 列印不完全";

        public SLRClassTotalMulti()
        {
            InitializeComponent();
        }

        private void SLRClassTotalMulti_Load(object sender, EventArgs e)
        {
            // 設定背景執行哪個method
            _BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            // 當執行結束後, 呼叫哪個method
            _BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            // 取得設定檔
            Campus.Report.ReportConfiguration reportConfig = new Campus.Report.ReportConfiguration(_SLRClassTotalMulti_ReportCofig);
            // 如果沒有設定過樣板
            if (reportConfig.Template == null)
            {
                reportConfig.Template = new Campus.Report.ReportTemplate(Properties.Resources.班級服務學習統計表_多學期_範本, Campus.Report.TemplateType.Word);
                reportConfig.Save();
            }
        }
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //取得設定檔
            Campus.Report.ReportConfiguration reportConfig = new Campus.Report.ReportConfiguration(_SLRClassTotalMulti_ReportCofig);
            //畫面內容(範本內容,預設樣式
            Campus.Report.TemplateSettingForm TemplateForm = new Campus.Report.TemplateSettingForm(reportConfig.Template, new Campus.Report.ReportTemplate(Properties.Resources.班級服務學習統計表_多學期_範本, Campus.Report.TemplateType.Word));
            //預設名稱
            TemplateForm.DefaultFileName = "班級服務學習統計表_多學期(範本)";
            //如果回傳為OK
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                //設定後樣試,回傳
                reportConfig.Template = TemplateForm.Template;
                //儲存
                reportConfig.Save();
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!_BGW.IsBusy)
            {
                _BGW.RunWorkerAsync();
            }
            else
            {
                MsgBox.Show("系統忙碌中,請稍後再試!!");
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // 主要邏輯區塊
        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //取得目前所選班級資料
            List<string> classIDList = NLDPanels.Class.SelectedSource;

            //取得班級學生資料
            List<StudentTotalObj> studentList = GetStudentOBJ(classIDList);

            //組合班級清單
            List<ClassTotalObj> classList = GetClassOBJ(classIDList);

            //取得設定檔
            Campus.Report.ReportConfiguration ConfigurationInCadre = new Campus.Report.ReportConfiguration(_SLRClassTotalMulti_ReportCofig);
            _template = ConfigurationInCadre.Template.ToDocument();
            _doc.Sections.Clear();

            // 把查詢出來的學生加入各個班級物件裡
            foreach (StudentTotalObj student in studentList)
            {
                foreach (ClassTotalObj aClass in classList)
                {
                    if (student.ref_class_id == aClass.ref_class_id)
                    {
                        if (!aClass.StudentDic.ContainsKey(student.ref_student_id))
                        {
                            aClass.StudentDic.Add(student.ref_student_id, student);
                        }
                    }
                }
            }

            // 取得學生IDs
            List<string> studentIDList = GetStudentID(studentList);
            string qu = string.Join("','", studentIDList);
            
            // 取得學生的服務時數
            List<SLRecord> SLRList = _accessHelper.Select<SLRecord>("ref_student_id in ('" + qu + "')");

            // 把服務時數加入各個班級的各個學生裡
            foreach (SLRecord slrRecord in SLRList)
            {
                foreach (ClassTotalObj aClass in classList)
                {
                    if (aClass.StudentDic.ContainsKey(slrRecord.RefStudentID))
                    {
                        aClass.SetSLRInStudent(slrRecord);
                    }
                }
            }
            
            classList.Sort(SortClass);

            // 把每個班級中的學生匯出至WORD
            foreach (ClassTotalObj aClass in classList)
            {
                if (aClass.StudentDic.Count > 0)
                {
                    Document PageOne = SetDocument(aClass);

                    if (PageOne != null)
                    {
                        _doc.Sections.Add(_doc.ImportNode(PageOne.FirstSection, true));
                    }
                }
            }

            e.Result = _doc;
        }

        // 當背景程式結束, 就會呼叫method
        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MsgBox.Show("列印作業已中止!!");
                return;
            }

            if (e.Error == null)
            {
                Document inResult = (Document)e.Result;

                try
                {
                    SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                    SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                    SaveFileDialog1.FileName = "班級服務學習統計表(多學期)";

                    if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        inResult.Save(SaveFileDialog1.FileName);
                        Process.Start(SaveFileDialog1.FileName);
                    }
                    else
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
                        return;
                    }
                }
                catch
                {
                    
                    FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                    return;
                }

                this.Close();
            }
            else
            {
                MsgBox.Show("列印發生錯誤:\n" + e.Error.Message);
                return;
            }
        }

        // WORD處理
        private Document SetDocument(ClassTotalObj aClass)
        {
            StringBuilder WarningStudents = new StringBuilder();

            // 取得班級中所有學生清單
            List<StudentTotalObj> studentObjList = new List<StudentTotalObj>(aClass.StudentDic.Values);
            
            //排序
            studentObjList.Sort(SortStudent);

            Document PageOne = (Document)_template.Clone(true);
            _run = new Run(PageOne);
            DocumentBuilder builder = new DocumentBuilder(PageOne);

            int columnCount = aClass.SLRSchoolYearSemesterDic.Count;
            int columnIndex = 1;

            // 文件移到學年度跟學期的標題
            builder.MoveToMergeField("學年度"); // 移到有學年度的標籤上
            Cell cellSchoolYear = (Cell)builder.CurrentParagraph.ParentNode;

            // 準備學年度跟學期的標題
            List<string> SchoolYearSemesterList = new List<string>(aClass.SLRSchoolYearSemesterDic.Keys);

            // 排序
            SchoolYearSemesterList.Sort(SortSchoolYearSemester);

            // 取得學年度跟學期的欄位數量
            _MAX_COLUMN_COUNT = GetRemainColumn(cellSchoolYear);

            // 確保資料的數量跟欄位一致
            if (SchoolYearSemesterList.Count > _MAX_COLUMN_COUNT)
            {
                int delCount = SchoolYearSemesterList.Count - _MAX_COLUMN_COUNT;
                for(int intI=0; intI<delCount; intI++)
                {
                    SchoolYearSemesterList.RemoveAt(0);
                }
            }

            // 輸出學年度跟學期的標題
            foreach (string each in SchoolYearSemesterList)
            {
                // 學年度
                Write(cellSchoolYear, each);

                // 不是最後一筆資料的話, 就移到下一個cell
                if(columnIndex < columnCount)
                {
                    cellSchoolYear = GetMoveRightCell(cellSchoolYear, 1);
                }
                columnIndex++;
            }
            

            builder.MoveToMergeField("資料"); // 移到有資料的標籤上
            Cell cell = (Cell)builder.CurrentParagraph.ParentNode;

            //需要Insert Row
            //取得目前Row
            Row 日3row = (Row)cell.ParentRow;

            for (int x = 1; x < studentObjList.Count; x++)
            {
                cell.ParentRow.ParentTable.InsertAfter(日3row.Clone(true), cell.ParentNode);
            }

            //學生ID
            foreach (StudentTotalObj student in studentObjList)
            {
                // 用來計算輸出了幾次學年度跟學期的服務時數
                int SLROutputCount = 0;

                // 座號
                Write(cell, student.seat_no);
                cell = GetMoveRightCell(cell, 1);
                
                // 姓名
                Write(cell, student.student_name);
                cell = GetMoveRightCell(cell, 1);
                
                // 學號
                Write(cell, student.student_number);
                cell = GetMoveRightCell(cell, 1);
                
                // 性別
                Write(cell, student.student_gender);
                cell = GetMoveRightCell(cell, 1);
                
                // 學習時數
                columnCount = SchoolYearSemesterList.Count;
                columnIndex = 1;
                foreach(string each in SchoolYearSemesterList)
                {
                    if (student.SLRDic.ContainsKey(each))
                    {
                        Write(cell, student.SLRDic[each].ToString());
                        SLROutputCount++;
                    }
                    else
                    {
                        Write(cell, "0");
                    }

                    // 不是最後一筆資料的話, 就移到後面一個cell
                    if (columnIndex < columnCount)
                    {
                        cell = GetMoveRightCell(cell, 1);
                    }
                    columnIndex++;
                }

                // 假如學生服務學習的總學年度跟學期的數量不等於列印的的數量, 加入警告清單
                if (student.SLRDic.Count != SLROutputCount )
                {
                    string msg = String.Format(_WarningMsg, student.seat_no, student.student_name);

                    WarningStudents.AppendLine(msg);
                }

                Row Nextrow = cell.ParentRow.NextSibling as Row; //取得下一行
                if (Nextrow == null)
                {
                    break;
                }
                cell = Nextrow.FirstCell; //第一格
            }

            #region MailMerge

            List<string> name = new List<string>();
            List<string> value = new List<string>();

            name.Add("班級");
            value.Add(aClass.class_name);

            name.Add("導師");
            value.Add(aClass.GetTeacherName());

            name.Add("提示");
            value.Add(WarningStudents.ToString());

            PageOne.MailMerge.Execute(name.ToArray(), value.ToArray());

            #endregion

            return PageOne;
        }

        private int GetRemainColumn(Cell cell)
        {
            if(cell == null) return 0;

            Row row = cell.ParentRow;
            int indexCol = row.IndexOf(cell);
            int totalCol = row.Count;

            return totalCol - indexCol;
        }

        // WORD處理, 以Cell為基準,向右移一格
        private Cell GetMoveRightCell(Cell cell, int count)
        {
            if (count == 0) return cell;

            Row row = cell.ParentRow;
            int col_index = row.IndexOf(cell);
            Table table = row.ParentTable;
            int row_index = table.Rows.IndexOf(row);

            try
            {
                return table.Rows[row_index].Cells[col_index + count];
            }
            catch
            {
                return null;
            }
        }

        // WORD處理, 寫入資料
        private void Write(Cell cell, string text)
        {
            if (cell.FirstParagraph == null)
                cell.Paragraphs.Add(new Paragraph(cell.Document));
            cell.FirstParagraph.Runs.Clear();
            _run.Text = text;
            _run.Font.Size = 10;
            _run.Font.Name = "標楷體";
            cell.FirstParagraph.Runs.Add(_run.Clone(true));
        }

        // 排序: 學年度/學期
        private int SortSchoolYearSemester(string obj1, string obj2)
        {
            string[] tmp = obj1.Split('/');
            string SchoolYearSem1 = tmp[0].PadLeft(3, '0') + tmp[1];

            tmp = obj2.Split('/');
            string SchoolYearSem2 = tmp[0].PadLeft(3, '0') + tmp[1];

            return SchoolYearSem1.CompareTo(SchoolYearSem2);
        }

        // 排序:年級/班級序號/班級名稱
        private int SortClass(ClassTotalObj obj1, ClassTotalObj obj2)
        {
            //年級
            string seatno1 = obj1.class_grade_year.PadLeft(1, '0');
            seatno1 += obj1.class_index.PadLeft(3, '0');
            seatno1 += obj1.class_name.PadLeft(10, '0');

            string seatno2 = obj2.class_grade_year.PadLeft(1, '0');
            seatno2 += obj2.class_index.PadLeft(3, '0');
            seatno2 += obj2.class_name.PadLeft(10, '0');

            return seatno1.CompareTo(seatno2);
        }

        // 排序:座號/姓名
        private int SortStudent(StudentTotalObj obj1, StudentTotalObj obj2)
        {
            string seatno1 = obj1.seat_no.PadLeft(3, '0');
            seatno1 += obj1.student_name.PadLeft(10, '0');

            string seatno2 = obj2.seat_no.PadLeft(3, '0');
            seatno2 += obj2.student_name.PadLeft(10, '0');

            return seatno1.CompareTo(seatno2);
        }

        // 取得學生ID
        private List<string> GetStudentID(List<StudentTotalObj> studentList)
        {
            return studentList.Select(x => x.ref_student_id).ToList();
        }

        // 取得班級資料 By SQL
        private List<ClassTotalObj> GetClassOBJ(List<string> _ClassIDList)
        {
            List<ClassTotalObj> list = new List<ClassTotalObj>();
            string classid = string.Join("','", _ClassIDList);
            string qu = "select class.id,class.class_name,teacher.id as teacher_id,teacher.teacher_name,teacher.nickname,class.grade_year,class.display_order ";
            qu += "from class LEFT join teacher on class.ref_teacher_id=teacher.id ";
            qu += "where class.id in('" + classid + "')";
            DataTable dt = _queryHelper.Select(qu);

            foreach (DataRow row in dt.Rows)
            {
                ClassTotalObj obj = new ClassTotalObj(row);
                list.Add(obj);
            }
            return list;

        }

        // 取得學生資料 By SQL
        private List<StudentTotalObj> GetStudentOBJ(List<string> _ClassIDList)
        {
            List<StudentTotalObj> list = new List<StudentTotalObj>();
            //取得班級學生資料
            string qu = "select student.id,class.id as class_id,student.seat_no,student.student_number,student.name,student.status,student.gender from student join class on class.id=student.ref_class_id where class.id in('" + string.Join("','", _ClassIDList) + "')";
            DataTable dt = _queryHelper.Select(qu);

            foreach (DataRow row in dt.Rows)
            {
                StudentTotalObj obj = new StudentTotalObj(row);
                //學生不等於 刪除 與 畢業及離校
                if (obj.status != "16" && obj.status != "256")
                {
                    list.Add(obj);
                }
            }
            return list;
        }
    }
}
