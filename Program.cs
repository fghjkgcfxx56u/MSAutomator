//all .dll files needed
using System;
using System.IO;
using System.Linq;
using OpenQA.Selenium;
using Newtonsoft.Json;
using System.Threading;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Collections.Generic;


namespace Msr
{
    class MSRewards
    {
        static IWebDriver getDriver()
        {
           
            ChromeOptions options = new ChromeOptions();
            options.AddExtension("Config/Automator.crx");
            options.AddArgument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36 Edge/86.0.622.56");
            options.AddArgument("--log-level=3");
            options.BinaryLocation =  @"J:/Program Files/Google/Chrome/Application/chrome.exe";
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            IWebDriver driver = new ChromeDriver(service, options);
            return driver;
        }
        static IWebDriver getMobile()
        {
            ChromeOptions moptions = new ChromeOptions();
            moptions.AddArgument("--log-level=3");
            moptions.EnableMobileEmulation("Nexus 5");
            moptions.BinaryLocation = @"J:/Program Files/Google/Chrome/Application/chrome.exe";
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            IWebDriver mobile = new ChromeDriver(service, moptions);
            return mobile;
        }

        static void Sleep(double Seconds) 
        {
            int Timeout = Convert.ToInt32(Seconds);
            Thread.Sleep(Timeout * 1000);
        }
        
        //model to extract credentials
        public class jsonmodel 
        {
            public string email { get; set; }
            public string password { get; set; }
        }

        //main function [takes credentials extracted from the json model as arguments]
        static void msr(string email, string password) 
        {
            Console.WriteLine("Starting automation of: " + email + "\n");

            //some random words
            List<string> words = new List<string>() { "kuzndc", "upomby", "rpchnz", "dufxld", "okabsq", "xzkijt", "pngwmf", "skwvgl", "diszom", "wqgvmy", "kliryf", "ctarnz", "kxkxjh", "ggwged", "trzfhl", "vmwegq", "pnkvkp", "slhwfa", "xpeiyq", "zszwoi", "mcivyw", "uxkxjq", "ssyuzv", "haxdcl", "bgofcy", "gwwnfn", "porapw", "wkjoiv", "ksugno", "csktxf", "jycbbp", "xhaiou", "zfuxme", "nyuull", "ropjgy", "czdpms", "myzhxa", "yepdvk", "hwrphz", "ufztpn", "lvirbz", "bngffh", "kadvqn", "tlzgaw", "smruck", "hldeei", "psvmbi", "dzynwv", "thmwym", "zqqghf", "gtbjmd", "gcodsk", "bmpjza", "nldjqe", "imjibj", "zulftx", "rryoui", "uebuui", "ydxzvw", "crrowv", "nozogv", "iblymt", "eporou", "elecbe", "tkyvac", "ljqctv", "tougpd", "nvplch", "gckxdp", "hpolye", "dozqtb", "qopbcx", "wposqv", "adnqih", "dgcmkp", "yaokws", "mbvioz", "oiyvbf", "osiuqn", "lfcfrm", "peneon", "qjxocb", "zxhdxt", "gsfyci", "cahuzw", "hfeawd", "edxvvd", "pnkpbe", "jvgklb", "synvwz", "pgarcq", "hovslp", "qgidxe", "ifpinb", "mhgwuw", "yjzela", "dobfdn", "bgrfwx", "rjepmw", "umvxnd", "apafpy", "gcuyfh", "yvkzti", "whazef", "zumtxs", "qokkce", "mrfjjr", "oosfna", "fknudk", "htxjhj", "ebqqjd", "ehlzpr", "fmscnm", "fgmoro", "iassba", "ifthxc", "wtygoh", "ejngvn", "wjzsax", "doceur", "jgrjze", "zlgtme", "ylaebr", "dgsrzg", "vpqntr", "cawmjf", "qujxnx", "sonxib", "xsasnj", "shppfx", "xesyhq", "rhersd", "pytyey", "ygepbt", "owprsu", "uggksu", "zxjuti", "pyufal", "rtqhij", "ddfkgd", "rmnwau", "zemqwu", "svxrnk", "cojnkr", "kngdrv", "pkhezx", "geprqo", "ldgocc", "ombiie", "rtsufo" };
            var random = new Random();

            IWebDriver driver = MSRewards.getDriver();

            Console.WriteLine("Created new chromedriver");

            //login
            driver.Navigate().GoToUrl("https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=13&ct=1605296186&rver=6.7.6631.0&wp=MBI_SSL&wreply=https%3a%2f%2fwww.bing.com%2fsecure%2fPassport.aspx%3frequrl%3dhttps%253a%252f%252fwww.bing.com%252f%253ftoWww%253d1%2526redig%253dE994C2D23BCD4D2EA9621849E98AB3CA%2526wlsso%253d1%2526wlexpsignin%253d1%26sig%3d097B9CEBAB1D604A1B7D936AAA2A6114&lc=1040&id=264960&CSRFToken=3177a0f1-2f55-4746-adc0-7b3267c630d3&aadredir=1");

            TimeSpan time = TimeSpan.FromSeconds(200);

            IWebElement mail = new WebDriverWait(driver, time).Until(ExpectedConditions.ElementToBeClickable(By.Id("i0116")));
            mail.SendKeys(email);

            IWebElement button1 = new WebDriverWait(driver, time).Until(ExpectedConditions.ElementToBeClickable(By.Id("idSIButton9")));
            button1.Click();

            Sleep(2.0);

            IWebElement pass = new WebDriverWait(driver, time).Until(ExpectedConditions.ElementToBeClickable(By.Id("i0118")));
            pass.SendKeys(password);

            IWebElement button2 = new WebDriverWait(driver, time).Until(ExpectedConditions.ElementToBeClickable(By.Id("idSIButton9")));
            button2.Click();

            Sleep(5.0);

            IWebElement button3 = new WebDriverWait(driver, time).Until(ExpectedConditions.ElementToBeClickable(By.Id("bnp_btn_accept")));
            button3.Click();

            Console.WriteLine("Logged in");

            driver.Navigate().GoToUrl("https://account.microsoft.com/rewards/?fref=home.banner.rewards");

            //daily activities
            for (int i = 0; i < 3; i++)
            {
                while (true)
                {
                    try
                    {
                        int num = i + 1; 
                        string missions = "/html/body/div[1]/div[3]/main/div/ui-view/mee-rewards-dashboard/main/div/mee-rewards-daily-set-section/div/mee-card-group[1]/div/mee-card[" + num.ToString() + "]/div/card-content/mee-rewards-daily-set-item-content/div/div[3]/a";
                        IWebElement mission = new WebDriverWait(driver, time).Until(ExpectedConditions.ElementToBeClickable(By.XPath(missions)));
                        mission.Click();
                        break;
                    }
                    catch
                    {
                        driver.Navigate().Refresh();
                        Sleep(10.0);
                    }

                }

                driver.SwitchTo().Window(driver.WindowHandles.Last());
                Sleep(45.0);

                driver.Close();
                Sleep(1.5);


                driver.SwitchTo().Window(driver.WindowHandles.Last());
            }


            //daily activities #2
            List<Int64> list = new List<Int64>();
            Dictionary<Int64, string> dict = new Dictionary<Int64, string>();

            int count = driver.FindElements(By.ClassName("c-card-content")).Count();
            var scount = Convert.ToInt32(count);


            for (int j = scount - 20; j <= scount - 20; j++)
            {
                
                list.Add(j);
                int Num = j + 1;
                dict.Add(j, "//*[@id=\"more-activities\"]/div/mee-card[" + Convert.ToString(Num) + "]/div/card-content/mee-rewards-more-activities-card-item/div/div[3]/a");
                driver.FindElement(By.XPath(dict[j])).Click();
                driver.SwitchTo().Window(driver.WindowHandles.Last());

                Sleep(2.0);

                driver.Close();
                driver.SwitchTo().Window(driver.WindowHandles.Last());


            }

            Console.WriteLine("Finished daily activities");

            //bing searches [PC]
            for (int x = 0; x <= 51; x++)
            {
                Random SleepTime = new Random();
                int index = random.Next(words.Count);
                string word = words[index];
                driver.Navigate().GoToUrl("https://www.bing.com/search?q=" + word + "&PC=U316&FORM=CHROMN");
                IEnumerable<int> sleepTime = Enumerable.Range(1, 3).Select(x => ((x/2)+(x/3))*2);
                Sleep(SleepTime.NextDouble() * 2.5);
            }

            driver.Quit();
            

            Console.WriteLine("Finished PC searches");

            IWebDriver mobile = MSRewards.getMobile();

            Console.WriteLine("Created new chromedriver with mobile emulation");

            //login
            mobile.Navigate().GoToUrl("https://bing.com");

            IWebElement mbutton3 = new WebDriverWait(mobile, time).Until(ExpectedConditions.ElementToBeClickable(By.Id("mHamburger")));
            Sleep(0.7);
            mbutton3.Click();

            IWebElement mbutton4 = new WebDriverWait(mobile, time).Until(ExpectedConditions.ElementToBeClickable(By.Id("HBSignIn")));
            Sleep(0.7);
            mbutton4.Click();

            IWebElement mmail = new WebDriverWait(mobile, time).Until(ExpectedConditions.ElementToBeClickable(By.Id("i0116")));
            mmail.SendKeys(email);

            IWebElement mbutton1 = new WebDriverWait(mobile, time).Until(ExpectedConditions.ElementToBeClickable(By.Id("idSIButton9")));
            mbutton1.Click();

            IWebElement mpass = new WebDriverWait(mobile, time).Until(ExpectedConditions.ElementToBeClickable(By.Id("i0118")));
            mpass.SendKeys(password);

            IWebElement mbutton2 = new WebDriverWait(mobile, time).Until(ExpectedConditions.ElementToBeClickable(By.Id("idSIButton9")));
            mbutton2.Click();

            Console.WriteLine("Logged in");

            //bing searches [MOBILE]
            for (int y = 0; y <= 21; y++)
            {
                int index = random.Next(words.Count);
                mobile.Navigate().GoToUrl("https://www.bing.com/search?q=" + words[index] + "&PC=U316&FORM=CHROMN");
            }

            mobile.Quit();
           

            Console.WriteLine("Finished mobile searches\nFinished automation of: " + email + "\n");


        }

        static void Main(string[] args)
        {
            string JSONFile = "";
            List<string> mandp = new List<string>();

            //get credentials
            using (var reader = new StreamReader("Config/credentials.json"))
            {
                JSONFile = reader.ReadToEnd();
            }

            Dictionary<string, string> credentials = JsonConvert.DeserializeObject<Dictionary<string, string>>(JSONFile);

            foreach (string email in credentials.Keys)
            {
                mandp.Add(email + ":   " + credentials[email]);
            }

            Console.WriteLine("\nStarting automating rewards with accounts: \n");

            foreach (string credential in mandp) 
            {
                Console.WriteLine("    -" + credential);
            }

            Console.WriteLine("\n\n");
           
            //loop over them
            foreach (string email in credentials.Keys) 
            {
                try
                {
                    msr(email, credentials[email]);
                }
                catch (Exception Ezeption) 
                {
                    Console.WriteLine(Ezeption);
                    using (StreamWriter Fs = File.CreateText("Debug/Debug.log"))
                    {
                        Fs.WriteLine(Ezeption.ToString());
                    }
                }
                
            }
            
            //end
            Console.WriteLine("\n\n--------------------------------------------------------FINISHED--------------------------------------------------------\n\n");

            //sleeping for 20 minutes to wait until user notices and closes application
            Sleep(1200.0);
            
        }   
    }
}
 