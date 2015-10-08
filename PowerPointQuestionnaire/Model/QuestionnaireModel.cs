using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace PowerPointQuestionnaire.Model
{
    public class QuestionnaireModel
    {
        public string user { get; set; }

        public int choices { get; set; }

        [JsonProperty("_id")]
        public string id { get; set; }

        public QuestionnaireAttachment image { get; set; }

        public string tempFileId { get; set; }
    }

    public class QuestionnaireAttachment
    {
        public string user { get; set; }

        public int size { get; set; }

        public string md5 { get; set; }
    }
}
