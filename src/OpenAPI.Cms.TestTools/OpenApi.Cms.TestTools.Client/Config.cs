using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenApi.Cms.TestTools.Client
{

    public class Config
    { 
        public static int Interval 
        {
            get
            {
                int interval;
                var val = ConfigurationManager.AppSettings["PublishInterval"];
                if (!int.TryParse(val, out interval))
                    return 100;

                return interval;
            }
        }

        public static string DbConnectionString
        {
            get
            {
                var val = ConfigurationManager.AppSettings["WebPublisherDB"]; 
                if (string.IsNullOrEmpty(val))
                    throw new Exception("Cannot find 'WebPublisherDB' in the confg file");

                return val;
            }
        }

        public static Dictionary<string, string> SheetsToPublish
        {
            get
            {
                var val = ConfigurationManager.AppSettings["PublishSheets"];
                if (string.IsNullOrEmpty(val))
                    throw new Exception("Cannot find 'PublishSheets' in the confg file");
              
                var settings = val.Split(',');
                var ret = new Dictionary<string, string>(settings.Length);
                for(int i = 0; i < settings.Length; i ++)
                {
                    var sheetRange = settings[i].Split('!');
                    ret[sheetRange[0]] = sheetRange.Length > 1 ? sheetRange[1] : string.Empty;
                }

                return ret;
            }
        }

    }
}
