using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadGmail
{
    internal class EmailModel
    {
        public EmailModel()
        {
            GatewayId = new List<string>();
        }

        public int MessageNumber { get; set; }
        public string From { get; set; }
        public string Firma { get; set; }
        public List<string> GatewayId { get; set; }
        public string GatewayList { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string BodyCopy { get; set; }
        public DateTime DateSent { get; set; }
    }
}
