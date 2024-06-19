using Aspose.Cells;
using Newtonsoft.Json;
using Aspose.Cells.Charts;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Telegram.Bot;
using Telegram.Bot.Requests;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;
using static System.Runtime.InteropServices.JavaScript.JSType;
using File = System.IO.File;

namespace Pozdravleniya_s_prazdnikami_bot
{
    public class Host
    {
        private static TelegramBotClient bot;
        static List<Chat> botChats = new List<Chat>();

        public Host(string token)
        {
            bot = new TelegramBotClient(token);
        }

        public void Start()
        {
            bot.StartReceiving(UpdateHandler, ErrorHandler);
            Console.WriteLine("Бот запущен");
        }

        private async Task ErrorHandler(ITelegramBotClient client, Exception exception, CancellationToken token)
        {
            Console.WriteLine("Ошибка: " + exception.Message);
            await Task.CompletedTask;
        }

        async Task SendMessage(Chat chatId, string message)
        {
            await bot.SendTextMessageAsync(
              chatId: chatId,
              text: message
            );
        }
        static Chat GetUserChat(string reqChatName)
        {
            foreach (Chat chat in botChats)
            {
                string chatName = chat.Title;

                if (chatName == reqChatName)
                {
                    return chat;
                }
            }
            return null;
        }

        private async Task UpdateHandler(ITelegramBotClient client, Update update, CancellationToken token)
        {
            var message = update.Message;//сообщение
            var user = message.From;//от кого пришло сообщение
            var chat = message.Chat;

            async Task SetSchedule(Chat reqChat)
            {
                long reqChatId = reqChat.Id;
                Console.WriteLine($"//Установка расписания");
                Console.WriteLine($"Чат группы: {reqChat.Title}, {reqChatId}");
                // Проход по всем чатам
                foreach (var chatId in JsonFile.chatData.Keys)
                {
                    Console.WriteLine($"Чат с пользователем: {chatId}");
                    if (chatId == chat.Id)
                    {
                        // Получение расписания для текущего чата
                        List<ExcelData> schedule = JsonFile.chatData[chatId];

                        // Установка расписания для каждого элемента в расписании
                        foreach (var item in schedule)
                        {
                            DateTime scheduledTime = item.Date; // Получение времени отправки из расписания

                            // Установка часов и минут
                            //scheduledTime = DateTime.Now.Date.AddHours(DateTime.Now.Hour).AddMinutes(DateTime.Now.Minute);
                            //scheduledTime = new DateTime(scheduledTime.Year, scheduledTime.Month, scheduledTime.Day, 18, 47, 0);
                            scheduledTime = new DateTime(scheduledTime.Year, scheduledTime.Month, scheduledTime.Day, DateTime.Now.Hour, DateTime.Now.Minute+1, 0);

                            // Если время отправки уже прошло, устанавливаем для следующего года
                            if (DateTime.Now > scheduledTime)
                                scheduledTime = scheduledTime.AddYears(1);

                            TimeSpan timeToWait = scheduledTime - DateTime.Now; // Вычисляем время до следующего события

                            int dueTime = (int)timeToWait.TotalSeconds;
                            int period = (int)TimeSpan.FromDays(365).TotalSeconds;

                            Timer timer = new Timer(async (obj) => await SendCongratulation(reqChatId, item.Congratulation, item.Username), null, dueTime, period); // Отправка сообщения каждый год в указанное время
                        }
                    }
                }
            }
            async Task SendCongratulation(long chatId, string text, string username)
            {
                try
                {
                    string congratulation = $"{username}\n{text}";
                   // await SendMessage(chat, congratulation);
                    await bot.SendTextMessageAsync(chatId, congratulation);
                    Console.WriteLine($"Сообщение отправлено {congratulation} в чат {chatId}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при отправке сообщения: {ex.Message}");
                }
            }

            try
            {
                switch (update.Type)
                {
                    case UpdateType.Message:
                        {
                            switch (chat.Type)
                            {
                                case ChatType.Private: //личные сообщения
                                    {
                                        switch (message.Type)
                                        {
                                            case MessageType.Text: //обработка текстовых сообщений
                                                {
                                                    Console.WriteLine($"Получено сообщение \"{message.Text}\" в чате {chat.Id} от пользователя {user.Username}.");
                                                    switch (message.Text.ToLower())
                                                    {
                                                        case string text when text == "/start":
                                                            {
                                                                await bot.SendTextMessageAsync(chat, "Рад знакомству! Я бот-помощник в праздничных делах🎊" +
                                                                     "\nЧтобы установить расписание праздников заполните эту табличку🤲");
                                                                var templatePath = $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\template.xlsx";
                                                                using (FileStream fileStream = File.OpenRead(templatePath))
                                                                {
                                                                    InputFile inputFile = new InputFileStream(fileStream, "Расписание праздников.xlsx");
                                                                    await bot.SendDocumentAsync(chat.Id, inputFile);
                                                                }
                                                                break;
                                                            }
                                                        case string text when text == "/help":
                                                            await bot.SendTextMessageAsync(chat, "Список доступных команд:" +
                                                                "\n/start -  начать знакомство с ботом сначала" +
                                                                "\n/help - вывести эту справку и руководство пользователя" +
                                                                "\n/viewschedule - просмотр сохраненного расписания");
                                                            var filePath = $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\Руководство пользователя.pdf";
                                                            using (FileStream fileStream = File.OpenRead(filePath))
                                                            {
                                                                InputFile inputFile = new InputFileStream(fileStream, "Руководство пользователя.pdf");
                                                                await bot.SendDocumentAsync(chat.Id, inputFile);
                                                            }
                                                            break;

                                                        case string text when text == "/viewschedule":
                                                            {
                                                                if (JsonFile.chatData.ContainsKey(chat.Id))
                                                                {
                                                                    var dataPath = $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\{chat.Id}.xlsx";
                                                                    using (FileStream fileStream = File.OpenRead(dataPath))
                                                                    {
                                                                        InputFile inputFile = new InputFileStream(fileStream, "Расписание праздников.xlsx");
                                                                        await bot.SendTextMessageAsync(chat, "Ваше расписание праздников: ");
                                                                        await bot.SendDocumentAsync(chat.Id, inputFile);
                                                                        await bot.SendTextMessageAsync(chat, "Если ваши планы изменились, не стесняйтесь вносить изменения!");
                                                                    }
                                                                }
                                                                else
                                                                    await SendMessage(chat, "Вы еще не создали ни одного расписания! Нажмите /start чтобы начать");
                                                                break;
                                                            }
                                                        default:
                                                            Chat reqChat = GetUserChat(message.Text);
                                                            await SetSchedule(reqChat);

                                                            await SendMessage(chat, "Расписание праздников установлено");
                                                            //await bot.SendTextMessageAsync(chat, "Я понимаю только определенные команды🥺\n/help", replyToMessageId: message.MessageId);
                                                            break;
                                                    }

                                                    break;
                                                }
                                            case MessageType.Document: //обработка файлов
                                                {
                                                    var fileId = message.Document.FileId;
                                                    var fileInfo = await bot.GetFileAsync(fileId);
                                                    var filePath = fileInfo.FilePath;

                                                    string destinationFilePath = $@"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\\{chat.Id}.xlsx";
                                                    FileStream fileStream = File.OpenWrite(destinationFilePath);
                                                    await bot.DownloadFileAsync(filePath, fileStream);
                                                    fileStream.Close();

                                                    Workbook workbook = new Workbook(destinationFilePath);

                                                    Worksheet worksheet = workbook.Worksheets[0];

                                                    Cells cells = worksheet.Cells;
                                                    int numRows = cells.MaxDataRow + 1;
                                                    await SendMessage(chat, "Записываю...");
                                                    for (int row = 1; row < numRows; row++)
                                                    {
                                                        string username = cells[row, 0].StringValue;
                                                        DateTime date = cells[row, 1].DateTimeValue;
                                                        string congratulation = cells[row, 2].StringValue;

                                                        if (!JsonFile.chatData.ContainsKey(chat.Id))
                                                        {
                                                            JsonFile.chatData[chat.Id] = new List<ExcelData>();
                                                        }

                                                        JsonFile.chatData[chat.Id].Add(new ExcelData { Username = username, Date = date, Congratulation = congratulation });
                                                    }
                                                    await bot.SendTextMessageAsync(chat, "Введите название чата, в котором нужно установить расписание");
                                                    //await SetSchedule();

                                                    //await SendMessage(chat, "Расписание праздников установлено");
                                                    break;
                                                }
                                            default:
                                                await bot.SendTextMessageAsync(chat, "Я понимаю только определенные команды🥺\n/help", replyToMessageId: message.MessageId);
                                                break;
                                        }
                                        break;
                                    }
                                case ChatType.Supergroup: //супергруппы
                                case ChatType.Group: //группы
                                    {
                                        switch (message.Type)
                                        {
                                            case MessageType.Text:
                                                {
                                                    Console.WriteLine("//Запись групповых чатов в список");
                                                    botChats.Add(chat);
                                                    Console.WriteLine($"В список чатов добавлен новый чат: {chat.Title}, {chat.Id}");
                                                    foreach(Chat item in botChats)
                                                        Console.WriteLine($"{item.Title}, {item.Id}");
                                                    Console.WriteLine($"Получено сообщение \"{message.Text}\" в чате {chat.Id} от пользователя {user.Username}.");
                                                    await bot.SendTextMessageAsync(chat, "Эту и другие команды нужно отправлять мне личным сообщением", replyToMessageId: message.MessageId);
                                                    break;
                                                }
                                        }
                                        break;
                                    }
                            }
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            Console.WriteLine($"Сообщение: {message?.Text ?? "Это не текст"}");
            await Task.CompletedTask;
        }

        public class ExcelData
        {
            public string Username { get; set; }
            public DateTime Date { get; set; }
            public string Congratulation { get; set; }

        }
    }
}