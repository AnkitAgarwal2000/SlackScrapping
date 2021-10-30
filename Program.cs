using System;
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Net;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using CsvHelper;
using System.IO;
using System.Globalization;
using Newtonsoft.Json.Linq;

namespace feedbacks
{
    class Program
    {
        static void Main(string[] args)
        {   
            JObject json= GetCallResult();
            List<Feed> feeds= new List<Feed>();
            feeds=wrap(json);
            Export(feeds);
            // Console.WriteLine(feeds[0].name);
        }

        private static void Export(List<Feed> feeds)
        {
            using (var writer = new StreamWriter("./feedbacks.csv"))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                 csv.WriteRecords(feeds);
            }
        }

        static List<Feed> wrap(JObject json)
        {
            int nm = (int)json.SelectToken("$.messages.total");
            Console.WriteLine(nm);

            var values = json.SelectTokens("$.messages.matches.[*].blocks..text");
            var users = json.SelectTokens("$.messages.matches.[*].user");
            var ts = json.SelectTokens("$.messages.matches.[*].ts");

            JObject[] li = new JObject[nm];
            string[] lim = new String[nm];
            string[] lits = new String[nm];

            int k=0;
            foreach (JToken item in users)
            {
                li[k] = GetUserJson((string)item);
                k++;
            }
            k=0;
            foreach (JToken item in values)
            {
                lim[k] = (string)item;
                k++;
            }
            k=0;
            foreach (JToken item in ts)
            {
                lits[k] = (string)item;
                k++;
            }

            List<Feed> feeds = GetDetails(li, lim, lits, nm);
            return feeds;
        }

        static List<Feed> GetDetails(JObject[] li, string[] lim, string[] lits, int s)
        {
            var feeds = new List<Feed>();

            int k=0;
            foreach(JObject json in li)
            {
                var feed = new Feed();
                feed.message=lim[k];
                
                double dt=double.Parse(lits[k], System.Globalization.CultureInfo.InvariantCulture);
                System.DateTime dtDateTime = new DateTime(1970,1,1,0,0,0,0,System.DateTimeKind.Utc);
                dtDateTime = dtDateTime.AddSeconds( dt ).ToLocalTime();
                feed.TimeStamp=dtDateTime;

                feed.Zone = (string)json.SelectToken("$.user.tz");
                feed.Name = (string)json.SelectToken("$.user.real_name");
                feed.Language = (string)json.SelectToken("$.user.locale");
                feeds.Add(feed);
                k++;
            }

            return feeds;
        }

        static JObject GetUserJson(String id)
        {
            string url = "https://slack.com/api/users.info?user="+id+"&include_locale=true&pretty=1";

            var httpRequest = (HttpWebRequest)WebRequest.Create(url);
            httpRequest.Method = "POST";
            
            httpRequest.Accept = "*/*";
            httpRequest.Headers["Authorization"] = "Bearer xoxp-**************-************-********************-*********************************";
            httpRequest.ContentType = "application/x-www-form-urlencoded";
            httpRequest.Headers["Content-Length"] = "0";
            
            
            var httpResponse = (HttpWebResponse)httpRequest.GetResponse();

            string result;
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
               result = streamReader.ReadToEnd();
            }
            Console.WriteLine(httpResponse.StatusCode);
            return GetJson(result);
        }

        static JObject GetJson(string res)
        {
            JObject json = JObject.Parse(res);
            return json;
        }
        static JObject GetCallResult()
        {
            var url = "https://slack.com/api/search.messages?query=Feedback&team_id=C02KC4B3G4R&pretty=1";

            var httpRequest = (HttpWebRequest)WebRequest.Create(url);
            httpRequest.Method = "POST";

            httpRequest.Accept = "*/*";
            httpRequest.Headers["Authorization"] = "Bearer xoxp-**************-************-********************-*********************************";
            httpRequest.ContentType = "application/x-www-form-urlencoded";

            var data = "query=feedback";

            using (var streamWriter = new StreamWriter(httpRequest.GetRequestStream()))
            {
                streamWriter.Write(data);
            }

            var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
            string res;
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                res = streamReader.ReadToEnd();
            }
            Console.WriteLine(httpResponse.StatusCode);
            return(GetJson(res));

        }


    }
}
