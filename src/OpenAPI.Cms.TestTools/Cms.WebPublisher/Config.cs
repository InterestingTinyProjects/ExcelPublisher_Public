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
    }
}
