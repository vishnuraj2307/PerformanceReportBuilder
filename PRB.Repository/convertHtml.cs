using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace PRB.Repository
{
    public class convertHtml
    {
        public string send(string url,string node)
        {
            try
            {
                if (!string.IsNullOrEmpty(url))
                {
                    WebClient webClient = new WebClient();
                    webClient.Encoding = System.Text.Encoding.UTF8;
                    string html = webClient.DownloadString(url);
                    if(!string.IsNullOrEmpty(html))
                    {
                        HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument();
                        htmlDocument.LoadHtml(html);
                        if (htmlDocument != null)
                        {
                            HtmlAgilityPack.HtmlNode htmlNode = htmlDocument.DocumentNode.SelectSingleNode(node);
                            if(htmlNode != null)
                            {
                                string innerText=htmlNode.InnerText;
                                if(!string.IsNullOrEmpty(innerText))
                                {
                                    return innerText;
                                }
                            }
                        }
                    }
                }
                return null;

            }
            catch
            {
                return null;
            }
            
        }
    }
}
