using System;
using System.Threading;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Args;
using Telegram.Bot.Types;
using Newtonsoft.Json;
using Aspose.Cells;
using System.IO;
using File = System.IO.File;
using System.Xml.Linq;
using Telegram.Bot.Polling;

namespace Pozdravleniya_s_prazdnikami_bot
{
    public class JsonFile
    {
        public static Dictionary<long, List<Host.ExcelData>> chatData;
        public static string ChatDataFile = "ChatData.json";
    }
    class Program
    {
        static void Main(string[] args)
        {
            Host bot = new Host("7004222316:AAGmR-srArzAFigMTHawje756qnAyb-E6JM");
            if (File.Exists(JsonFile.ChatDataFile))
            {
                var json = File.ReadAllText(JsonFile.ChatDataFile);
                JsonFile.chatData = JsonConvert.DeserializeObject<Dictionary<long, List<Host.ExcelData>>>(json);
            }
            else
            {
                JsonFile.chatData = new Dictionary<long, List<Host.ExcelData>>();
            }
            bot.Start();
            Console.ReadLine();
            var stateJson = JsonConvert.SerializeObject(JsonFile.chatData);
            File.WriteAllText(JsonFile.ChatDataFile, stateJson);
        }
    }
}
