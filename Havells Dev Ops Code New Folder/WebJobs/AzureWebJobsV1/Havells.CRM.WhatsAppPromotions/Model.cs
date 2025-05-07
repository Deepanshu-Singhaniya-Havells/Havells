using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.WhatsAppPromotions
{
    internal class Model
    {
    }
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
    public class Body
    {
        public List<object> placeholders { get; set; }
    }

    public class Button
    {
        public string type { get; set; }
        public string parameter { get; set; }
    }

    public class Content
    {
        public string templateName { get; set; }
        public TemplateData templateData { get; set; }
        public string language { get; set; }
    }

    public class Header
    {
        public string type { get; set; }
        public string placeholder { get; set; }
    }

    public class Message
    {
        public long from { get; set; }
        public long to { get; set; }
        public string messageId { get; set; }
        public Content content { get; set; }
        public string callbackData { get; set; }
    }

    public class MsgSend
    {
        public List<Message> messages { get; set; }
    }

    public class TemplateData
    {
        public Body body { get; set; }
        public Header header { get; set; }
        public List<Button> buttons { get; set; }
    }


}
