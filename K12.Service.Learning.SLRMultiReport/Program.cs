using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Presentation;
using K12.Presentation;
using FISCA.Permission;

namespace K12.Service.Learning.SLRMultiReport
{
    public class Program
    {
        [FISCA.MainMethod]
        static public void Main()
        {
            RibbonBarItem reportClassBar = NLDPanels.Class.RibbonBarItems["資料統計"];
            reportClassBar["報表"]["學務相關報表"]["班級服務學習統計表(多學期)"].Enable = Permissions.班級服務學習統計表_多學期_權限;
            reportClassBar["報表"]["學務相關報表"]["班級服務學習統計表(多學期)"].Click += delegate
            {
                SLRClassTotalMulti frm = new SLRClassTotalMulti();
                frm.ShowDialog();
            };

            Catalog ribbon = RoleAclSource.Instance["班級"]["功能按鈕"];
            ribbon.Add(new RibbonFeature(Permissions.班級服務學習統計表_多學期, "班級服務學習統計表(多學期)"));
        }
    }
}
