using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Threading;
using System.ServiceProcess;
using System.Runtime.InteropServices;
using SmorcIRL.TempMail;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using System.Text;
using Newtonsoft.Json.Linq;
using WindowsInput;
using WindowsInput.Native;
using Tesseract;
using UglyToad.PdfPig;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using SmorcIRL.TempMail.Models;
using System.Net.Http;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.Text.Json;
using Windows.Media.AppBroadcasting;

class Program
{
    #region NameGen

    static string[] firstWord = { "schwarzer", "flotter", "gruener", "harmonischer", "schoener", "schlauer", "durchgeknallter", "gewaschender", "sauberer", "dreckiger", "lakaka", "suechtiger", "ungeliebter", "selbstmordgefährdeter", "depressiver", "gesprengter", "schwimmendes", "geile", "gruselige", "spooky", "schummelnder", "geschwindigkeitsUeberschreitende", "kleiner", "grosser", "veganer", "fleichfressender", "linksversiffter", "gefeierter", "auserwaehlter", "radikaler", "schlechter", "boeser", "guter", "attraktive", "begeisterte", "vieler", "harter", "vollgeschissene", "verschimmelter", "verseuchter", "versiffter", "nutzloser", "angestrengter", "unnoetiger", "halal", "haram", "gespannter", "erregter", "lange", "tiefer", "steifer", "gestreckte", "enge", "feuchte", "unterschriebene", "ausgeleiherte", "fleischiger", "maennlicher", "weibliche", "Steuernhinterziehende", "bombadierter", "gejagter", "politischVerfolgter", "transgender", "homosexueller", "gebleichter", "pythonnutzender", "spielsuechtiger", "rauchender" };
    static string[] secondWord = { "Klabautermann", "Seefahrer", "BurgerKingArbeiter", "Kuenstler", "Pirat", "Entwickler", "Verb", "Apflel", "Marrokaner", "Gieskanne", "Wetterballon", "Basketball", "Mappe", "Vikinger", "Sitzsackpolster", "CDUWaehler", "Gruenenwaehler", "BMWFahrer", "Helge", "Fahrlehrer", "Huan", "DireStraitsFan", "Schaumkrone", "Wackler", "Investor", "Helium", "Shakespeare", "Meister", "Rhabarbarkuchen", "Geisterbahn", "Schimmel", "Kerzentraeger", "Bierbrauer", "Omelett", "Franzose", "Drache", "Fussende", "Wolkenkratzer", "Mathegenie", "Monkey", "Affe", "Vogelscheuche", "Baum", "Minenschacht", "StuhlTischBank", "Vater", "Yarrack", "Topografie", "Geographie", "Franzose", "Karte", "Alkohol", "Flasche", "Designerstueck", "Spiel", "Taschentuch", "Metalldetektor", "jeans", "Unterhose", "Bomber", "Islam", "schlawiner", "Palestiner", "PythonNutzer", "Gambler", "Spielsuechtiger", "CasinoBesucher", "VegasLover", "Raucher", "Kartoffel", "Rechtschreibfehler", "Check24" };
    static int numRange = 200;

    public static async Task<string> GetCoolUsername()
    {
        Random random = new Random();

        string randNum = random.Next(numRange).ToString();
        string result = await GetRandomAdjectiveNounAsync() + randNum;

        if (String.IsNullOrEmpty(result))
        {
            result = (firstWord[random.Next(firstWord.Length)] + secondWord[random.Next(secondWord.Length)] + randNum);
        }

        return result;
    }

    private static readonly HttpClient client = new HttpClient();
    private const string BaseUrl = "https://random-word-form.herokuapp.com/random";

    public static async Task<string> GetRandomAdjectiveNounAsync()
    {
        try
        {
            Random rand = new Random();
            string randomLetter = lower[rand.Next(lower.Length)].ToString();
            // Get a random adjective
            string adjective = await GetRandomWordAsync($"/adjective/{randomLetter}", false);

            // Get a random noun
            string noun = await GetRandomWordAsync($"/noun/{randomLetter}", true);

            // Combine the two words
            return $"{adjective}{noun}";
        }
        catch (Exception ex)
        {
            // Handle any exceptions that occur during the HTTP request
            Console.WriteLine($"An error occurred: {ex.Message}");
            return null;
        }
    }

    private static async Task<string> GetRandomWordAsync(string endpoint, bool CapitalLetter)
    {
        // Make the GET request to the API
        HttpResponseMessage response = await client.GetAsync(BaseUrl + endpoint);

        // Ensure the request was successful
        response.EnsureSuccessStatusCode();

        // Read the response content as a string
        string jsonResponse = await response.Content.ReadAsStringAsync();

        try
        {
            // Deserialize the JSON array to a string array
            string[] words = JsonSerializer.Deserialize<string[]>(jsonResponse);

            if (CapitalLetter)
            {
                words[0] = CapitalizeFirstLetter(words[0]);
            }

            // Return the first word (since the API returns an array)
            return words[0];
        }
        catch (JsonException)
        {
            // Handle JSON parsing errors
            Console.WriteLine("Failed to parse JSON response.");
            return null;
        }
    }

    private static string CapitalizeFirstLetter(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        // Capitalize the first letter and concatenate with the rest of the string
        return char.ToUpper(input[0]) + input.Substring(1);
    }

    #endregion

    #region Shortcut
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [StructLayout(LayoutKind.Sequential)]
    public struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public POINT pt;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int x;
        public int y;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

    private const int WM_HOTKEY = 0x0312;

    // Modifier keys
    private const uint MOD_ALT = 0x0001;
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;
    private const uint MOD_WIN = 0x0008;

    // Virtual Key Codes
    private const uint VK_F9 = 0x78;  // F9 key
    private const uint VK_W = 0x57; //w
    private const uint VK_Q = 0x51; //q
    private const uint VK_E = 0x45;
    private const uint VK_S = 0x53;
    private const uint VK_1 = 0x31;
    #endregion

    #region Window Visibility

    [DllImport("kernel32.dll")]
    static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    const int SW_HIDE = 0;
    const int SW_SHOW = 5;

    static bool minimizeOnStart = false;

    #endregion

    #region Password Generation

    const string lower = "abcdefghijklmnopqrstuvwxyz";
    const string upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    const string digits = "0123456789";
    const string symbols = "!@#$%^&*()_+[]{};:,.?";

    public static string GeneratePassword()
    {
        Random _random = new Random();

        
        const string allChars = lower + upper + digits + symbols;

        int passwordLength = _random.Next(8, 17); // Generates a length between 8 and 16

        StringBuilder password = new StringBuilder(passwordLength);

        // Ensure the password contains at least one character from each required set
        password.Append(lower[_random.Next(lower.Length)]);
        password.Append(upper[_random.Next(upper.Length)]);
        password.Append(digits[_random.Next(digits.Length)]);
        password.Append(symbols[_random.Next(symbols.Length)]);

        // Fill the rest of the password with random characters from all character sets
        for (int i = 4; i < passwordLength; i++)
        {
            password.Append(allChars[_random.Next(allChars.Length)]);
        }

        // Shuffle the password to prevent the first four characters from being predictable
        return new string(password.ToString().OrderBy(c => _random.Next()).ToArray());
    }

    #endregion

    #region Important Variables or something i guess
    public static bool runningMSGThread = false;

    public static int pasteWaitTime = 400;
    public static int threadAliveTime = 1200000; //10 minutes
    public static int generateNewAccountTime = 600000; // more that 1,5 minutes

    public static MailClient currentClient;
    public static string currentPassword;
    public static string currentUsername;
    public static string[] currentVerificationCodes;


    public static bool deleteAccountWhenAbandoned = false;

    public static string newline = "NEWLINE";
    public static string rawBody = "RAWBODY";

    public static string NewLine(int amount)
    {
        string nl = "";

        for(int i = 0; i < amount; i++)
        {
            nl += newline;
        }

        return nl;
    }
    #endregion

    #region Clipboard Action

    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWnd);

    static void BringToFront()
    {
        // Get the handle of the active window
        IntPtr hWnd = GetForegroundWindow();

        // Bring the window to the foreground
        SetForegroundWindow(hWnd);
    }

    static InputSimulator sim = new InputSimulator();

    [STAThread]
    static void PasteText(string text)
    {
        Thread staThread = new Thread(() =>
        {
            try
            {
                string previousClipBoard = Clipboard.GetText();

                Thread.Sleep(pasteWaitTime / 2);
                Clipboard.SetText(text);
                Thread.Sleep(pasteWaitTime / 2);
                //SendKeys.SendWait("^v");
                sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);

                Thread.Sleep(pasteWaitTime / 2);
                if (previousClipBoard != null && previousClipBoard != "")
                {
                    if (previousClipBoard == null) previousClipBoard = "Wait a minute Mr.Postman!";
                    Clipboard.SetText(previousClipBoard);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Pasting Error, maybe dont click that fast, or: {ex.ToString()}");
            }

        });

        staThread.SetApartmentState(ApartmentState.STA);
        staThread.Start();
        staThread.Join();
    }

    #endregion

    #region Logging

    static void LogMessage(string message)
    {
        Process loggerProcess = new Process();

        // Get the full path to ConsoleLogger.exe
        string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string loggerPath = System.IO.Path.Combine(exeDirectory, "Logger", "ConsoleLogger.exe");

        // Escape newlines for command-line arguments
        string escapedMessage = message.Replace(Environment.NewLine, "\\n");

        // Quote the message to handle newlines and spaces
        string quotedMessage = "\"" + escapedMessage.Replace("\"", "\\") + "\"";

        if (!System.IO.File.Exists(loggerPath))
        {
            Console.WriteLine("ConsoleLogger.exe does not exist at the specified path.");
            return;
        }

        loggerProcess.StartInfo.FileName = loggerPath;
        loggerProcess.StartInfo.Arguments = quotedMessage;
        loggerProcess.StartInfo.UseShellExecute = true;

        try
        {
            loggerProcess.Start();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error starting process: {ex.Message}");
        }
    }

    public static string ExtractTextFromHtml(string rawHtml)
    {
        if (string.IsNullOrEmpty(rawHtml))
        {
            return string.Empty;
        }

        // Load the HTML into an HtmlDocument
        HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
        htmlDoc.LoadHtml(rawHtml);

        // Use the HtmlDocument's DocumentNode to extract all text content
        string plainText = htmlDoc.DocumentNode.InnerText;

        // Optionally, you can also clean up extra whitespace or newlines
        plainText = System.Web.HttpUtility.HtmlDecode(plainText); // Decode HTML entities
        plainText = System.Text.RegularExpressions.Regex.Replace(plainText, @"\s+", " "); // Replace multiple spaces/newlines with a single space
        plainText = plainText.Trim();

        return plainText;
    }

    public static string GetNewlineCompatible(string text)
    {
        if (String.IsNullOrEmpty(text) || String.IsNullOrWhiteSpace(text)) return "";
        return text.Replace("\n", newline);
    }

    public static string[] ExtractVerificationCodes(string emailText, int proximity = 50, string fancyAssText = "")
    {
        // List to store codes along with their proximity
        List<(string Code, int Proximity)> codesWithProximity = new List<(string Code, int Proximity)>();

        // Convert the email text to lowercase for case-insensitive matching
        emailText = emailText.ToLower();

        // Keywords to search for around the verification code
        string[] keywords = { "verification", "verify", "code", "authentication", "authenticate", "verif", "authen"};

        // Regular expression to find 4 or 6 digit numbers
        string codePattern = @"\b\d{4,6}\b";

        bool keywordFound = false;

        foreach (string keyword in keywords)
        {
            // Regex pattern to search for the keyword and a 4-6 digit number within the proximity range
            string pattern = $@"\b{keyword}\b.{{0,{proximity}}}?" + codePattern + $@"|" + codePattern + $@".{{0,{proximity}}}?\b{keyword}\b";
            MatchCollection matches = Regex.Matches(emailText, pattern, RegexOptions.IgnoreCase);

            foreach (Match match in matches)
            {
                keywordFound = true;
                // Extract the code from the matched string
                Match codeMatch = Regex.Match(match.Value, codePattern);
                if (codeMatch.Success)
                {
                    // Calculate the proximity (distance between keyword and code)
                    int keywordIndex = match.Value.IndexOf(keyword, StringComparison.OrdinalIgnoreCase);
                    int codeIndex = match.Value.IndexOf(codeMatch.Value, StringComparison.OrdinalIgnoreCase);
                    int distance = Math.Abs(keywordIndex - codeIndex);

                    // Add the code and its proximity to the list
                    codesWithProximity.Add((codeMatch.Value, distance));
                }
            }
        }

        if (!keywordFound)
        {
            if(!String.IsNullOrEmpty(fancyAssText) && fancyAssText != "" && fancyAssText != empty)
            {
                emailText = fancyAssText.ToLower();
            }
            // If no keywords are found, fallback to finding the code closest to the center of the text
            int centerPosition = emailText.Length / 2;
            MatchCollection matches = Regex.Matches(emailText, codePattern, RegexOptions.IgnoreCase);

            foreach (Match match in matches)
            {
                // Calculate the proximity (distance to the center of the text)
                int codeIndex = match.Index;
                int distanceToCenter = Math.Abs(centerPosition - codeIndex);

                // Add the code and its distance to the center to the list
                codesWithProximity.Add((match.Value, distanceToCenter));
            }
        }

        // Sort the codes by proximity (smallest proximity first)
        var sortedCodes = codesWithProximity.OrderBy(c => c.Proximity).Select(c => c.Code).ToArray();

        // Return the sorted codes, ensuring no duplicates
        return sortedCodes.Distinct().ToArray();
    }

    public static string GetLogo()
    {
        return "LOGO";
    }

    #endregion

    static async Task<MailClient> GenerateMailClient()
    {
        Random rand = new Random();
        MailClient client = new();

        var domain = await client.GetFirstAvailableDomainName();
        string customUsername = await GetCoolUsername();
        string customEmailAddress = $"{customUsername}@{domain}";
        string password = GeneratePassword();

        try
        {
            await client.Register(customEmailAddress, password);
        }catch (Exception ex)
        {
            Console.WriteLine($"ERROR creting new account, retrying {ex.Message}");

            Console.WriteLine($"Bad Disposable Email Address: {customEmailAddress}");
            Console.WriteLine($"Bad Password: {password}");

            Thread.Sleep(5000);
            await GenerateMailClient();
            return currentClient;
        }

        var account = await client.GetAccountInfo();

        Console.WriteLine($"Disposable Email Address: {account.Address}");
        Console.WriteLine($"Password: {password}");

        currentClient = client;
        currentPassword = password;
        currentUsername = customUsername;

        return client;
    }


    

    static async Task SetupTempMailsNShit()
    {
        await GenerateMailClient();


        #region Wait for shortcut

        if (RegisterHotKey(IntPtr.Zero, 1, MOD_ALT, VK_Q))
        {
            Console.WriteLine("Hotkey Alt + Q registered for address.");
        }
        else
        {
            Console.WriteLine("Failed to register hotkey Alt + Q.");
        }

        if (RegisterHotKey(IntPtr.Zero, 2, MOD_ALT, VK_W))
        {
            Console.WriteLine("Hotkey Alt + W registered for password.");
        }
        else
        {
            Console.WriteLine("Failed to register hotkey Alt + W.");
        }

        if (RegisterHotKey(IntPtr.Zero, 3, MOD_ALT, VK_E))
        {
            Console.WriteLine("Hotkey Alt + E registered for Username.");
        }
        else
        {
            Console.WriteLine("Failed to register hotkey Alt + E.");
        }
        if (RegisterHotKey(IntPtr.Zero, 4, MOD_ALT, VK_S))
        {
            Console.WriteLine("Hotkey Alt + S registered for VeriCode.");
        }
        else
        {
            Console.WriteLine("Failed to register hotkey Alt + S.");
        }
        if (RegisterHotKey(IntPtr.Zero, 5, MOD_ALT, VK_1))
        {
            Console.WriteLine("Hotkey Alt + 1 registered for Account REgenerate.");
        }
        else
        {
            Console.WriteLine("Failed to register hotkey Alt + 1.");
        }



        MSG msg;
        while(GetMessage(out msg, IntPtr.Zero, 0, 0))
        {
            if (msg.message == WM_HOTKEY)
            {

                //OnHotKeyPressed();
                
                switch (msg.wParam.ToInt32())
                {
                    case 1:
                        // Alt + Q was pressed
                        OnHotKeyPressed();
                        break;
                    case 2:
                        // Alt + W was pressed
                        OnHotKeyPSWDPressed();
                        break;
                    case 3:
                        //Alt + E
                        OnHotKeyUsernamePressed();
                        break;
                    case 4:
                        //Alt + S
                        OnHotKeyVerificationCodePressed();
                        break;
                    case 5:
                        OnAccountRegenerate();
                        break;
                }
                
            }
        }

        #endregion
    }

    static string empty = "EMPTY";

    static async void CheckNewMessages(MailClient currentClient)
    {
        bool running = true;
        bool runningExitThread = false;
        while (running)
        {
            var messages = await currentClient.GetAllMessages();

            foreach (var message in messages)
            {
                var messageDetails = await currentClient.GetMessage(message.Id);

                MessageSource source = await currentClient.GetMessageSource(message.Id);

                string plainText = ExtractTextFromHtml(source.Data);
                string? BodyText;

                if(messageDetails != null && messageDetails.BodyText != null && messageDetails.BodyText.ToString() != null) {
                    BodyText = messageDetails.BodyText.ToString();
                    if (String.IsNullOrEmpty(BodyText))
                    {
                        BodyText = empty;
                    }
                }
                else
                {
                    BodyText = empty;
                }

                currentVerificationCodes = ExtractVerificationCodes(plainText, 50, message.Subject.ToString() + BodyText);
                verificationCodeCounter = 0;
                string verificationCode = currentVerificationCodes.Length > 0 ? currentVerificationCodes[0] : "";

                
                string messageToLog = $"{newline} " +
    $"From: {message.From.Address}{newline} " +
    $"Subject: {message.Subject}{newline} " +
    $"Code: {verificationCode}{NewLine(2)} " +
    $"Body: {newline} " +
    $"---------------------------- {NewLine(2)} " +
    $"{GetNewlineCompatible(BodyText)}{newline} " +
    $"============================ {newline}" +
    $"{rawBody}: {GetNewlineCompatible(plainText)}";

                LogMessage(messageToLog);

                await currentClient.MarkMessageAsSeen(message.Id, true);
                await currentClient.DeleteMessage(message.Id);


                if (!runningExitThread)
                {
                    Console.WriteLine($"Exiting CheckThread in {threadAliveTime} Milliseconds");
                    runningExitThread = true;
                    //Disable Account 1.5 minutes after first message
                    Thread timedThread = new Thread(async () =>
                    {
                        Thread.Sleep(generateNewAccountTime);

                        await GenerateMailClient();

                        runningMSGThread = false;

                        Thread.Sleep(threadAliveTime - generateNewAccountTime);

                        running = false;

                        if (deleteAccountWhenAbandoned)
                        {
                            await currentClient.DeleteAccount();
                        }

                        Console.WriteLine("Exited a CheckThread");

                    });
                    timedThread.Start();
                }
            }
            await Task.Delay(3000); //sleep sum time!
        }
    }



    [STAThread]
    static async Task Main()
    {

        if(minimizeOnStart)
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);
        }

        await SetupTempMailsNShit();
        //UnregisterHotKey(IntPtr.Zero, 1); //TODO do that somewhere else!
        
    }

    #region Cooldown because my code is as unstable as my mental condition

    private static DateTime lastHotKeyPress = DateTime.MinValue;
    private static readonly object cooldownLock = new object();
    private static readonly TimeSpan cooldown = TimeSpan.FromSeconds(0.5 + (pasteWaitTime/1000) * 2);

    private static bool CanExecute()
    {
        lock (cooldownLock)
        {
            if (DateTime.Now - lastHotKeyPress < cooldown)
            {
                Console.WriteLine($"Cooldown in effect, try again in: {lastHotKeyPress + cooldown}");
                return false;
            }

            lastHotKeyPress = DateTime.Now;
            return true;
        }
    }

    #endregion

    #region HotKey Functions

    [STAThread]
    private static void OnHotKeyPressed()
    {
        if (!CanExecute()) return;
        Console.WriteLine("Hotkey pressed!");
        //if (currentClient == null) await GenerateMailClient();

        string emailaddress = currentClient.Email;

        PasteText(emailaddress);


        if (runningMSGThread)
        {
            Console.WriteLine("Message Thread already running");
            return;
        }
        runningMSGThread = true;

        Thread messageCheckerThread = new Thread(() => CheckNewMessages(currentClient));
        messageCheckerThread.Start();

    }

    [STAThread]
    private static void OnHotKeyPSWDPressed()
    {
        if (!CanExecute()) return;
        Console.WriteLine("Hotkey PSWD pressed!");

        PasteText(currentPassword!);
    }

    [STAThread]
    private static void OnHotKeyUsernamePressed()
    {
        if (!CanExecute()) return;
        Console.WriteLine("Hotkey Username pressed!");

        PasteText(currentUsername!);
    }

    public static int verificationCodeCounter = 0;

    [STAThread]
    private static void OnHotKeyVerificationCodePressed()
    {
        if (!CanExecute() || currentVerificationCodes == null) return;
        Console.WriteLine("Hotkey VeriCode pressed!");

        if(!(verificationCodeCounter + 1 <= currentVerificationCodes!.Length))
        {
            verificationCodeCounter = 0;
        }

        PasteText(currentVerificationCodes[verificationCodeCounter]);
        verificationCodeCounter++;
    }

    [STAThread]
    private static async void OnAccountRegenerate()
    {
        Console.WriteLine("yur");
        if (!CanExecute()) return;
        Console.WriteLine("Hotkey Account Regenerate pressed!");

        await GenerateMailClient();
        runningMSGThread = false;
    }

    #endregion

    void OnStop()
    {
        Console.WriteLine("Stopping");
        UnregisterHotKey(IntPtr.Zero, 1);
        UnregisterHotKey(IntPtr.Zero, 2);
        UnregisterHotKey(IntPtr.Zero, 3);
        UnregisterHotKey(IntPtr.Zero, 4);
        UnregisterHotKey(IntPtr.Zero, 5);
    }
}
