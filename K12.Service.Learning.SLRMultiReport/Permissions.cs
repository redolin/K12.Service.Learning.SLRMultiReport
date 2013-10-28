using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K12.Service.Learning.SLRMultiReport
{
    class Permissions
    {
        public static string 班級服務學習統計表_多學期 { get { return "K12.Service.Learning.Modules.SLRClassTotalMulti.cs"; } }
        public static bool 班級服務學習統計表_多學期_權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[班級服務學習統計表_多學期].Executable;
            }
        }
    }
}
