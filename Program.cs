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
        public static Dictionary<long, Chat> chatGroups;
        public static string ChatDataFile = "ChatData.json";
        public static string ChatGroupsFile = "ChatGroups.json";
    }
    class Program
    {
        static void Main(string[] args)
        {
            Host bot = new Host("7004222316:AAHenlPnp_w5yZC3FLEf6OMthna3KTgOhhE");

            try
            {
                // Чтение файла
                if (File.Exists(JsonFile.ChatDataFile))
                {
                    using (var fileStream = File.OpenText(JsonFile.ChatDataFile))
                    {
                        var json = fileStream.ReadToEnd();
                        JsonFile.chatData = JsonConvert.DeserializeObject<Dictionary<long, List<Host.ExcelData>>>(json);
                    }
                }
                else
                {
                    JsonFile.chatData = new Dictionary<long, List<Host.ExcelData>>();
                }
                // Чтение файла
                if (File.Exists(JsonFile.ChatGroupsFile))
                {
                    using (var fileStream = File.OpenText(JsonFile.ChatGroupsFile))
                    {
                        var json = fileStream.ReadToEnd();
                        JsonFile.chatGroups = JsonConvert.DeserializeObject<Dictionary<long, Chat>>(json);
                    }
                }
                else
                {
                    JsonFile.chatGroups = new Dictionary<long, Chat>();
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при работе с файлом: {ex.Message}");
            }
            Console.WriteLine("//До запуска:");
            foreach (var pair in JsonFile.chatGroups)
                Console.WriteLine($"{pair.Key}, {pair.Value.Title}");
            foreach (var pair in JsonFile.chatData)
                Console.WriteLine($"{pair.Key}, {pair.Value}");

            bot.Start();
            Console.ReadLine();

            try
            {

                // Запись файла
                using (var fileStream = File.CreateText(JsonFile.ChatDataFile))
                {
                    var chatDataJson = JsonConvert.SerializeObject(JsonFile.chatData);
                    fileStream.Write(chatDataJson);
                }
                // Запись файла
                using (var fileStream = File.CreateText(JsonFile.ChatGroupsFile))
                {
                    var chatGroupsJson = JsonConvert.SerializeObject(JsonFile.chatGroups);
                    fileStream.Write(chatGroupsJson);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при работе с файлом: {ex.Message}");
            }
        }
    }
}
