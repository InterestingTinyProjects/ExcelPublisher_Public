using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenApi.Cms.TestTools.Client
{
    public class Config
    {
        private IConfiguration _config;
        public Config(IConfiguration config)
        {
            _config = config;
        }

        public string WebPublisherDB
        {
            get
            {
                var val = _config["WebPublisherDB"];
                if (string.IsNullOrEmpty(val))
                    throw new Exception("Cannot find WebPublisherDB in appsettings.json");

                return val;
            }
        }


        public int NoPublishWarningThreshold
        {
            get
            {
                int noPublishThreshold;
                var val = _config["NoPublishThreshold"];
                if (!int.TryParse(val, out noPublishThreshold))
                    return 1000;

                return noPublishThreshold;
            }
        }

        public string DefaultSheetName
        {
            get
            {
                var val = _config["DefaultSheetName"];
                if (string.IsNullOrEmpty(val))
                    throw new Exception("Cannot find DefaultSheetName in appsettings.json");

                return val;
            }
        }

        public Dictionary<string, string[]> SheetFormatter
        {
            get
            {
                var dic = _config.GetSection("SheetFormatter").Get<Dictionary<string, string[]>>();
                if (dic == null)
                    return new Dictionary<string, string[]>(0);

                return dic;
            }
        }

    }
}
