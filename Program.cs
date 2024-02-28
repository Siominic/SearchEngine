    using System.ComponentModel; //necessary for the main functions of the code
using System.Xml;
using HtmlAgilityPack; //assists in the crawling of the code
using Microsoft.AspNetCore.Builder; //assists in the RateLimiting of the code
using Search_EngineTLS; //namespace to make the code uniform
using Microsoft.AspNetCore.Http;
//using Microsoft.AspNetCore.Mvc.ApplicationParts;
using System.Security.Cryptography.X509Certificates;
/*These Libraries are used for the different aspects of the code*/

namespace Search_EngineTLS
/*I have used a namespace so that the code is uniform and communicates correctly*/
{
    internal class Program
    /*The centralised class that refers all the classes and methods from it*/
    {
        private static Seed seed;
        private static Queue queue;
        private static Crawled crawled;

        static void Main(string[] args)
        /*this is the main method, which is the first run and calls the other methods and attributes to it*/
        {
            RateLimiter rate = new RateLimiter();
            RateLimiter.Rate_Limiter(120, 2);
        }
        public static void Initialize()
        {

            string path = Directory.GetCurrentDirectory();
            string seedPath = Path.Combine(path, "Seed.txt");
            string queuePath = Path.Combine(path, "Queue.txt");
            string crawledPath = Path.Combine(path, "Crawled.txt");

            seed = new(seedPath);
            var seedURLs = seed.Items;
            queue = new(queuePath, seedURLs);
            crawled = new(crawledPath);


             Task.Run(async () => await Crawl()).Wait();
        }
        static async Task Crawl()

        {
            do
            {
                string url = queue.Top;

                Crawl crawl = new(url);
                await crawl.Start();

                if (crawl.parsedURLs.Count > 0)
                    await ProcessURLs(crawl.parsedURLs);

                await PostCrawl(url);

            } while (queue.HasURLs);
        }
        static async Task ProcessURLs(List<string> urls)
        {
            foreach (var url in urls)
            {
                if (!crawled.HasBeenCrawled(url) && !queue.IsInQueue(url))
                    await queue.Add(url);
            }
        }
        static async Task PostCrawl(string url)
        {
            await queue.Remove(url);

            await crawled.Add(url);
        }

    }
}
class Seed
{
    /// <summary>
    /// Returns all seed URLs.
    /// </summary>
    public string[] Items
    {
        get => File.ReadAllLines(path);
    }

    private readonly string path;

    public Seed(string path)
    {
        this.path = path;

        string[] seedURLs = new string[]
        {
            "https://www.nasa.gov"

        };

        using StreamWriter file = File.CreateText(path);

        foreach (string url in seedURLs)
            file.WriteLine(url.ToCleanURL());
    }
}
class Queue
{
    /// <summary>
    /// Returns the first item in the queue.
    /// </summary>
    public string Top
    {
        get => File.ReadAllLines(path).First();
    }

    /// <summary>
    /// Returns all items in the queue;
    /// </summary>
    public string[] All
    {
        get => File.ReadAllLines(path);
    }

    /// <summary>
    /// Returns a value based on whether there are URLs in the queue.
    /// </summary>
    public bool HasURLs
    {
        get => File.ReadAllLines(path).Length > 0;
    }

    private readonly string path;

    public Queue(string path, string[] seedURLs)
    {
        this.path = path;

        using StreamWriter file = File.CreateText(path);

        foreach (string url in seedURLs)
            file.WriteLine(url.ToCleanURL());
    }

    public async Task Add(string url)
    {
        using StreamWriter file = new(path, append: true);

        await file.WriteLineAsync(url.ToCleanURL());
    }

    public async Task Remove(string url)
    {
        IEnumerable<string> filteredURLs = All.Where(u => u != url);

        await File.WriteAllLinesAsync(path, filteredURLs);
    }

    public bool IsInQueue(string url) => All.Where(u => u == url).Any();
}
class Crawled
{
    private readonly string path;

    public Crawled(string path)
    {
        this.path = path;
        File.Create(path).Close();
    }

    public bool HasBeenCrawled(string url) => File.ReadAllLines(path).Any(c => c == url.ToCleanURL());

    public async Task Add(string url)
    {
        using StreamWriter file = new(path, append: true);

        await file.WriteLineAsync(url.ToCleanURL());
    }
}
static class StringExtensions
{
    public static string ToCleanURL(this string str) => str.Trim().ToLower();
}
class Crawl
{
    public readonly string url; //this needs to be gotten from the Index class
    public string webPage; //this needs to be gotten from the index class
    public List<string> parsedURLs;
    public Crawl(string url)
    {
        this.url = url;
        webPage = null;
        parsedURLs = new List<string>();
    }
    public async Task Start()
    {
        await GetWebPage();

        if (!string.IsNullOrWhiteSpace(webPage))
        {
            ParseURLs(url);
        }
    }

    public async Task GetWebPage()
    {
        int i = 0;
        HttpClientHandler clientHandler = new HttpClientHandler();
        clientHandler.ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => cert.Verify();

        // Pass the handler to httpclient(from you are calling api)
        HttpClient client = new HttpClient(clientHandler);

        client.Timeout = TimeSpan.FromSeconds(60);


        string responseBody = await client.GetStringAsync(url);

        if (!string.IsNullOrWhiteSpace(responseBody))
            webPage = responseBody;
        Console.WriteLine("------------------------------");
        Trimming trimming = new Trimming();
        Trimming.ExtractContent(webPage);
    }

    public void ParseURLs(string url)
    {
        int outwardLinks = 0;
        HtmlDocument htmlDoc = new HtmlDocument();
        // Use this to manually input some HTML
        // Change this into a file location
        htmlDoc.LoadHtml(webPage);
        foreach (HtmlNode link in htmlDoc.DocumentNode.SelectNodes("//a[@href]"))
        {
            string hrefValue = link.GetAttributeValue("href", string.Empty);

            if (hrefValue.StartsWith("http"))
                parsedURLs.Add(hrefValue);
            outwardLinks++;
        }
        Console.WriteLine("URL    " + url);
        Console.WriteLine("Outwards Links  " + outwardLinks);
        Console.WriteLine("-------------------------------");
        Thread.Sleep(15000);


        // pass variables from here

    }
}
