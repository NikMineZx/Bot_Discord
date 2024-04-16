using DSharpPlus.CommandsNext;
using DSharpPlus.CommandsNext.Attributes;
using DSharpPlus.Entities;
using OfficeOpenXml;
using System.Xml.Serialization;

namespace Discord_Bot_StalCraft.commands
{
    public class CommandInfo
    {
        public CommandContext Context { get; set; }
        public string PlayerName { get; set; }
        public string CommandName { get; set; }
        public float Amount { get; set; }
    }

    public class ExcelFilePathConfig
    {
        public string FilePath { get; set; }

        public void Save(string filePath)
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ExcelFilePathConfig));
                using (TextWriter writer = new StreamWriter(filePath))
                {
                    serializer.Serialize(writer, this);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving Excel file path: {ex.Message}");
            }
        }

        public static ExcelFilePathConfig Load(string filePath)
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ExcelFilePathConfig));
                using (TextReader reader = new StreamReader(filePath))
                {
                    return (ExcelFilePathConfig)serializer.Deserialize(reader);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading Excel file path: {ex.Message}");
                return null;
            }
        }
    }

    internal class TestCommands : BaseCommandModule
    {
        private readonly ulong commandChannelId = 1227340356123623434; // ID канала для команд 1219747115735978065
        private readonly ulong displayChannelId = 1227377194712432670; // ID канала для отображения информации 1219375413046935582
        private const ulong BettingChannelId = 1227377194712432670; // ID канала для ставок
        private readonly ulong BettConfirmID = 1224452570764935349; // ID канала для принятия ставок
        private ExcelFilePathConfig excelFilePathConfig;

        public TestCommands()
        {
            excelFilePathConfig = ExcelFilePathConfig.Load("ExcelFilePathConfig.xml");
            excelFilePathConfig ??= new ExcelFilePathConfig();
        }

        private Queue<CommandInfo> commandQueue = new Queue<CommandInfo>();
        private bool processingQueue = false;

        public async Task Shutdown(CommandContext ctx)
        {
            int time = 2;
            if (ctx.Channel.Id != commandChannelId)
            {
                await ctx.RespondAsync("Команды можно использовать только в определенном канале.");
                return;
            }

            var displayChannel = ctx.Guild.GetChannel(displayChannelId);
            if (displayChannel != null && displayChannel is DiscordChannel)
            {
                //var embedBuilder = new DiscordEmbedBuilder()
                //    .WithTitle("Уважаемые беттеры!")
                //    .WithDescription("**Бот перестанет принимать ставки через `10 минут`.**")
                //    .WithColor(new DiscordColor(255, 0, 0));
                //await displayChannel.SendMessageAsync(embed: embedBuilder.Build());
                //await Task.Delay(TimeSpan.FromMinutes(time));

                //embedBuilder = new DiscordEmbedBuilder()
                //    .WithTitle("Уважаемые беттеры!")
                //    .WithDescription("**Бот перестанет принимать ставки через `8 минут`.**")
                //    .WithColor(new DiscordColor(255, 0, 0));
                //await displayChannel.SendMessageAsync(embed: embedBuilder.Build());
                //await Task.Delay(TimeSpan.FromMinutes(time));

                var embedBuilder = new DiscordEmbedBuilder()
                    .WithTitle("Уважаемые беттеры!")
                    .WithDescription("**Бот перестанет примимать ставки через `6 минут`.**")
                    .WithColor(new DiscordColor(255, 0, 0));
                await displayChannel.SendMessageAsync(embed: embedBuilder.Build());
                await Task.Delay(TimeSpan.FromMinutes(time));

                embedBuilder = new DiscordEmbedBuilder()
                    .WithTitle("Уважаемые беттеры!")
                    .WithDescription("**Бот перестанет примимать ставки через `4 минут`.**")
                    .WithColor(new DiscordColor(255, 0, 0));
                await displayChannel.SendMessageAsync(embed: embedBuilder.Build());
                await Task.Delay(TimeSpan.FromMinutes(time));

                embedBuilder = new DiscordEmbedBuilder()
                    .WithTitle("Уважаемые беттеры!")
                    .WithDescription("**Бот перестанет примимать ставки через `2 минут`.**")
                    .WithColor(new DiscordColor(255, 0, 0));
                await displayChannel.SendMessageAsync(embed: embedBuilder.Build());
                await Task.Delay(TimeSpan.FromMinutes(time));

                FileInfo file = new FileInfo(excelFilePathConfig.FilePath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    Console.WriteLine("Doshel");
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    var commandName = worksheet.Cells[10, 1].GetValue<string>();
                    string coefficient = worksheet.Cells[10, 6].GetValue<string>();
                    var commandName_2 = worksheet.Cells[11, 1].GetValue<string>();
                    string coefficient_2 = worksheet.Cells[11, 6].GetValue<string>();

                    var embedBuilder_end = new DiscordEmbedBuilder()
                    .WithTitle("Уважаемые беттеры!")
                    .WithDescription($"**Бот забирает деньги и уходит. <3 **\n**Прием ставок на событие `{commandName}` vs `{commandName_2}` закрыт!**\n" +
                    "=================================================\n" +
                    "**Итоговый коэффициент по результатам ставок:**\n")
                    .WithColor(new DiscordColor(250, 225, 2));
                    embedBuilder_end.AddField($"**`{commandName}`**", $"Коэффициент: `{(float)Math.Round(Convert.ToSingle(coefficient), 2)}`");
                    embedBuilder_end.AddField($"**`{commandName_2}`**", $"Коэффициент: `{(float)Math.Round(Convert.ToSingle(coefficient_2), 2)}`\n=================================================\n **Да начнутся игры!**");
                    await displayChannel.SendMessageAsync(embed: embedBuilder_end.Build());
                    package.Save();
                }
            }
            else
            {
                await ctx.RespondAsync("Не удалось найти канал для отображения информации.");
            }
            await ctx.Client.DisconnectAsync();
        }

        [Command("set_excel")]
        public async Task SetExcelPath(CommandContext ctx, string filePath)
        {
            if (ctx.Channel.Id != commandChannelId)
            {
                await ctx.RespondAsync("Команды можно использовать только в определенном канале.");
                return;
            }

            excelFilePathConfig.FilePath = filePath;
            excelFilePathConfig.Save("ExcelFilePathConfig.xml");

            await ctx.RespondAsync($"Путь к файлу Excel успешно сохранен: {filePath}");
        }

        [Command("startbets"), Description("Начать прием ставок на событие")]
        public async Task EmbedMessage(CommandContext ctx)
        {
            if (ctx.Channel.Id != commandChannelId)
            {
                await ctx.RespondAsync("Команды можно использовать только в определенном канале.");
                return;
            }

            var embedBuilder = new DiscordEmbedBuilder()
                .WithTitle("Уважаемые беттеры!")
                .WithDescription("")
                .WithColor(new DiscordColor(255, 255, 0));

            if (!File.Exists(excelFilePathConfig.FilePath))
            {
                await ctx.RespondAsync("Файл в системе не найден, Пожалуйста добавьте новый Excel.");
                return;
            }

            await ctx.RespondAsync($"Обработка файла `{Path.GetFileName(excelFilePathConfig.FilePath)}` началась.");

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(excelFilePathConfig.FilePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        await ctx.RespondAsync("Файл Excel не содержит листов.");
                        return;
                    }
                    var commandName = worksheet.Cells[10, 1].GetValue<string>();
                    string coefficient = worksheet.Cells[10, 5].GetValue<string>();
                    string maxBet = worksheet.Cells[10, 8].GetValue<string>();

                    var commandName_2 = worksheet.Cells[11, 1].GetValue<string>();
                    string coefficient_2 = worksheet.Cells[11, 5].GetValue<string>();
                    string maxBet_2 = worksheet.Cells[11, 8].GetValue<string>();

                    embedBuilder.AddField($"Мы открываем прием ставок на событие **`{commandName}`** vs **`{commandName_2}`** открыт!", "================================================================");
                    embedBuilder.AddField("**Информация о ставках:**", $"Коэффициент будет меняться в зависимости от ставок, совершенных на данный матч с обновлением в одну ставку.\r\nИтоговый коэффициент будет написан в сообщении о закрытии приема ставок!!!\nМинимальная ставка - **`1000`** руб.\n" +
                        "================================================================");

                    embedBuilder.AddField($"**`{commandName}`**", $"Коэффициент: **`{(float)Math.Round(Convert.ToSingle(coefficient), 2)}`** | Потолок максимальной ставки: **`{(float)Math.Round(Convert.ToSingle(maxBet), 2)}`** руб.");
                    embedBuilder.AddField($"**`{commandName_2}`**", $"Коэффициент: **`{(float)Math.Round(Convert.ToSingle(coefficient_2), 2)}`** | Потолок максимальной ставки: **`{(float)Math.Round(Convert.ToSingle(maxBet_2), 2)}`**  руб.");
                }
                await ctx.RespondAsync("Файл успешно обработан. Команды были добавлены.");
            }
            catch (Exception ex)
            {
                await ctx.RespondAsync($"Произошла ошибка при обработке файла: {ex.Message}");
            }

            var channel = ctx.Guild.GetChannel(displayChannelId);
            if (channel != null && channel is DiscordChannel displayChannel)
            {
                await displayChannel.SendMessageAsync(embedBuilder);
            }
            else
            {
                await ctx.RespondAsync("Не удалось найти канал для отображения информации.");
            }
            Shutdown(ctx);
        }

        [Command("bet")]
        public async Task Bet(CommandContext ctx, string playerName, string commandName, float amount)
        {
            if (ctx.Channel.Id != BettingChannelId)
            {
                await ctx.RespondAsync("Команду можно использовать только в определенном канале.");
                return;
            }

            CommandInfo commandInfo = new CommandInfo
            {
                Context = ctx,
                PlayerName = playerName,
                CommandName = commandName,
                Amount = amount
            };
            commandQueue.Enqueue(commandInfo);

            if (!processingQueue)
            {
                processingQueue = true;
                await ProcessQueue();
            }
        }

        private async Task ProcessQueue()
        {
            while (commandQueue.Count > 0)
            {
                CommandInfo commandInfo = commandQueue.Dequeue();
                await BetInternal(commandInfo.Context, commandInfo.PlayerName, commandInfo.CommandName, commandInfo.Amount);
            }
            processingQueue = false;
        }

        private async Task BetInternal(CommandContext ctx, string playerName, string commandName, float amount)
        {
            var embedBuilder_error = new DiscordEmbedBuilder()
                .WithTitle("Ставка")
                .WithDescription("")
                .WithColor(new DiscordColor(255, 3, 179));

            if (ctx.Message.Attachments.Count == 0)
            {
                DiscordChannel displayErrorChannel = ctx.Guild.GetChannel(displayChannelId); // Переименовали переменную
                if (displayErrorChannel != null)
                {
                    embedBuilder_error.AddField("Ошибка:",$"**`{playerName}`** Прикрепите скриншот вашей ставки к сообщению.");
                    await displayErrorChannel.SendMessageAsync(embedBuilder_error.Build());
                }
                else
                {
                    await ctx.RespondAsync("Не удалось найти канал для отображения информации.");
                }
                return;
            }

            if (amount < 1000)
            {
                DiscordChannel displayErrorChannel = ctx.Guild.GetChannel(displayChannelId);
                if (displayErrorChannel != null)
                {
                    embedBuilder_error.AddField("Ошибка:", $"**`{playerName}`** Нельзя ставить меньше минимальной ставки.");
                    await displayErrorChannel.SendMessageAsync(embedBuilder_error.Build());
                }
                else
                {
                    await ctx.RespondAsync("Не удалось найти канал для отображения информации.");
                }
                return;
            }

            var embedBuilder = new DiscordEmbedBuilder()
                .WithTitle("Ставка")
                .WithDescription("")
                .WithColor(new DiscordColor(98, 247, 5));

            var channel = ctx.Guild.GetChannel(displayChannelId);
            FileInfo file = new FileInfo(excelFilePathConfig.FilePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                DiscordChannel displayErrorChannel = ctx.Guild.GetChannel(displayChannelId);
                if (commandName != worksheet.Cells["A10"].GetValue<string>() && commandName != worksheet.Cells["A11"].GetValue<string>())
                {
                    if (displayErrorChannel != null)
                    {
                        embedBuilder_error.AddField("Ошибка:", $"**`{playerName}`** Команда \"{commandName}\" не существует.");
                        await displayErrorChannel.SendMessageAsync(embedBuilder_error.Build());
                    }
                    else
                    {
                        await ctx.RespondAsync("Не удалось найти канал для отображения информации.");
                    }
                    return;
                }

                if (worksheet.Cells["I10"].GetValue<float>() > 0 && worksheet.Cells["I11"].GetValue<float>() > 0)
                {
                    if (commandName == worksheet.Cells["A10"].GetValue<string>() && amount > worksheet.Cells["I10"].GetValue<float>() ||
                        commandName == worksheet.Cells["A11"].GetValue<string>() && amount > worksheet.Cells["I11"].GetValue<float>())
                    {
                        if (displayErrorChannel != null)
                        {
                            embedBuilder_error.AddField("Ошибка", $"**`{playerName}`** Ставка превышает максимальное допустимое значение.");
                            await displayErrorChannel.SendMessageAsync(embedBuilder_error.Build());
                        }
                        else
                        {
                            await ctx.RespondAsync("Не удалось найти канал для отображения информации.");
                        }
                        return;
                    }
                }
                else
                {
                    if (commandName == worksheet.Cells["A10"].GetValue<string>() && amount > worksheet.Cells["H10"].GetValue<float>() ||
                        commandName == worksheet.Cells["A11"].GetValue<string>() && amount > worksheet.Cells["H11"].GetValue<float>())
                    {
                        if (displayErrorChannel != null)
                        {
                            embedBuilder_error.AddField("Ошибка:", $"**`{playerName}`** Ставка превышает максимальное допустимое значение.");
                            await displayErrorChannel.SendMessageAsync(embedBuilder_error.Build());
                        }
                        else
                        {
                            await ctx.RespondAsync("Не удалось найти канал для отображения информации.");
                        }
                        return;
                    }
                }
            }

            if (SaveBetToExcel(playerName, commandName, amount, false))
            {
                (float firstBet, float secondBet, float firstKoef, float secondKoef, string firstName, string secondName) = SendNewKoeficients();
                embedBuilder.AddField($"**Информация:**",$"**Ставка от « **`{playerName}`** » на команду « **`{commandName}`** » в размере « **`{amount}`** » была принята!**\n" +
                $"**Минимальная ставка « **`1000`** руб.** »\n" +
                "======================================================================\n" +
                $"**Нынешние коэффициенты на команды:**\n" +
                    $"**« **`{firstName}`** » \n Коэф. - « **`{(float)Math.Round(Convert.ToSingle(firstKoef), 2)}`** » | Потолок максимальной ставки « **`{firstBet}`** » руб.**\n" +
                        $"** « **`{secondName}`** » \n Коэф. - « **`{(float)Math.Round(Convert.ToSingle(secondKoef), 2)}`** » | Потолок максимальной ставки « **`{secondBet}`** » руб.**");
            }
            else
            {
                await ctx.RespondAsync("Не удалось сохранить ставку. Пожалуйста, попробуйте снова.");
            }

            if (channel != null && channel is DiscordChannel displayChannel)
            {
                await displayChannel.SendMessageAsync(embedBuilder);
            }
            else
            {
                await ctx.RespondAsync("Не удалось найти канал для отображения информации.");
            }
        }

        private bool SaveBetToExcel(string playerName, string commandName, float amount, bool confirmed)
        {
            try
            {
                FileInfo file = new FileInfo(excelFilePathConfig.FilePath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    int row = 4;
                    while (worksheet.Cells[row, 13].Value != null)
                    {
                        row++;
                    }

                    int betId = GetNextBetId();
                    worksheet.Cells[row, 13].Value = betId;
                    worksheet.Cells[row, 14].Value = playerName;
                    worksheet.Cells[row, 15].Value = commandName;
                    worksheet.Cells[row, 16].Value = amount;
                    worksheet.Cells[row, 18].Value = confirmed ? "Yes" : "No";

                    worksheet.Calculate();

                    package.Save();
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private (float, float, float, float, string, string) SendNewKoeficients()
        {
            float firstBet = 1;
            float secondBet = 1;
            float firstKoef = 1;
            float secondKoef = 1;
            string firstName = "";
            string secondName = "";
            try
            {
                FileInfo file = new FileInfo(excelFilePathConfig.FilePath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet.Cells["I10"].GetValue<float>() <= 0 || worksheet.Cells["I11"].GetValue<float>() <= 0)
                    {
                        firstBet = worksheet.Cells[10, 8].GetValue<float>();
                        secondBet = worksheet.Cells[11, 8].GetValue<float>();
                        firstKoef = worksheet.Cells[10, 5].GetValue<float>();
                        secondKoef = worksheet.Cells[11, 5].GetValue<float>();
                        firstName = worksheet.Cells["A10"].GetValue<string>();
                        secondName = worksheet.Cells["A11"].GetValue<string>();
                    }
                    else
                    {
                        firstBet = worksheet.Cells[10, 9].GetValue<float>();
                        secondBet = worksheet.Cells[11, 9].GetValue<float>();
                        firstKoef = worksheet.Cells[10, 6].GetValue<float>();
                        secondKoef = worksheet.Cells[11, 6].GetValue<float>();
                        firstName = worksheet.Cells["A10"].GetValue<string>();
                        secondName = worksheet.Cells["A11"].GetValue<string>();
                    }
                    package.Save();
                }
                return (firstBet, secondBet, firstKoef, secondKoef, firstName, secondName);
            }
            catch (Exception)
            {
                return (firstBet, secondBet, firstKoef, secondKoef, firstName, secondName);
            }
        }

        private int GetNextBetId()
        {
            try
            {
                FileInfo file = new FileInfo(excelFilePathConfig.FilePath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        return 1;
                    }

                    int lastBetId = 0;
                    int row = 4;
                    while (worksheet.Cells[row, 13].Value != null)
                    {
                        lastBetId = Convert.ToInt32(worksheet.Cells[row, 13].Value);
                        row++;
                    }

                    return lastBetId + 1;
                }
            }
            catch (Exception)
            {
                return -1; // Возвращаем -1 в случае ошибки
            }
        }


    [Command("confirmbet")]
        public async Task ConfirmBet(CommandContext ctx, int betId)
        {
            if (ctx.Channel.Id != BettConfirmID)
            {
                await ctx.RespondAsync("Команду можно использовать только в определенном канале.");
                return;
            }

            try
            {
                FileInfo file = new FileInfo(excelFilePathConfig.FilePath);
                if (!file.Exists)
                {
                    await ctx.RespondAsync("Файл Excel не найден.");
                    return;
                }

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        await ctx.RespondAsync("Лист Excel не найден.");
                        return;
                    }

                    bool found = false;
                    int row = 4;
                    while (worksheet.Cells[row, 12].Value != null)
                    {
                        int id = Convert.ToInt32(worksheet.Cells[row, 12].Value);
                        if (id == betId)
                        {
                            worksheet.Cells[row, 17].Value = "Yes"; // Помечаем ставку как подтвержденную
                            found = true;
                            break;
                        }
                        row++;
                    }

                    if (found)
                    {
                        package.Save(); // Сохраняем изменения в файле Excel
                        await ctx.RespondAsync($"Ставка с ID {betId} была успешно подтверждена.");
                    }
                    else
                    {
                        await ctx.RespondAsync($"Ставка с ID {betId} не найдена.");
                    }
                }
            }
            catch (Exception ex)
            {
                await ctx.RespondAsync($"Произошла ошибка при подтверждении ставки: {ex.Message}");
            }
        }

        [Command("ConfirmbetName")]
        public async Task ConfirmBetName(CommandContext ctx, string user)
        {
            if (ctx.Channel.Id != BettConfirmID)
            {
                await ctx.RespondAsync("Команду можно использовать только в определенном канале.");
                return;
            }

            try
            {
                FileInfo file = new FileInfo(excelFilePathConfig.FilePath);
                if (!file.Exists)
                {
                    await ctx.RespondAsync("Файл Excel не найден.");
                    return;
                }

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        await ctx.RespondAsync("Лист Excel не найден.");
                        return;
                    }

                    bool found = false;
                    int row = 4;
                    while (worksheet.Cells[row, 12].Value != null)
                    {
                        string name = Convert.ToString(worksheet.Cells[row, 13].Value);
                        if (name == user)
                        {
                            worksheet.Cells[row, 17].Value = "Yes"; // Помечаем ставку как подтвержденную
                            found = true;
                            break;
                        }
                        row++;
                    }

                    if (found)
                    {
                        package.Save(); // Сохраняем изменения в файле Excel
                        await ctx.RespondAsync($"Ставка Игрока {user} была успешно подтверждена.");
                    }
                    else
                    {
                        await ctx.RespondAsync($"Ставка Игрока {user} не найдена.");
                    }
                }
            }
            catch (Exception ex)
            {
                await ctx.RespondAsync($"Произошла ошибка при подтверждении ставки: {ex.Message}");
            }
        }

        [Command("viewbets")]
        public async Task ViewBets(CommandContext ctx)
        {
            if (ctx.Channel.Id != BettConfirmID)
            {
                await ctx.RespondAsync("Команду можно использовать только в определенном канале.");
                return;
            }

            List<Bet> bets = GetBetsFromExcel();
            if (bets.Count == 0)
            {
                await ctx.RespondAsync("Нет доступных ставок.");
                return;
            }

            string message = "Ставки:\n";
            foreach (var bet in bets)
            {
                message += $"ID: {bet.Id}, Игрок: {bet.PlayerName}, Команда: {bet.CommandName}, Ставка: {bet.Amount}, Confirmed: {(bet.Confirmed ? "Yes" : "No")}\n";
            }

            await ctx.RespondAsync(message);
        }
        private List<Bet> GetBetsFromExcel()
        {
            try
            {
                FileInfo file = new FileInfo(excelFilePathConfig.FilePath);
                if (!file.Exists)
                {
                    return new List<Bet>();
                }

                List<Bet> bets = new List<Bet>();

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        return new List<Bet>();
                    }

                    int row = 4;
                    while (worksheet.Cells[row, 12].Value != null)
                    {
                        Bet bet = new Bet
                        {
                            Id = Convert.ToInt32(worksheet.Cells[row, 12].Value),
                            PlayerName = Convert.ToString(worksheet.Cells[row, 13].Value),
                            CommandName = Convert.ToString(worksheet.Cells[row, 14].Value),
                            Amount = Convert.ToSingle(worksheet.Cells[row, 15].Value),
                            Confirmed = Convert.ToString(worksheet.Cells[row, 17].Value) == "Yes"
                        };
                        bets.Add(bet);
                        row++;
                    }
                }

                return bets;
            }
            catch (Exception)
            {
                return new List<Bet>();
            }
        }

    }

    public class Bet
    {
        public int Id { get; set; }
        public string PlayerName { get; set; }
        public string CommandName { get; set; }
        public float Amount { get; set; }
        public bool Confirmed { get; set; }
    }
}

