using System;
using System.IO;
using HtmlAgilityPack;

namespace IconTests
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var dir = @"H:\testdata\oletests\extracted";

            string filePath = Path.Combine(dir, "SamplePNGImage_3mb.png");
            var displayName = "Lorem ipsum dolor sit amer.png".Replace(" ", " ");
            var base64Str = ShellIcon.GetFileIconAsBase64(filePath, displayName);
            var l = base64Str.Length;
            var htmlDoc = new HtmlDocument();

            HtmlNode embeddedLinkNode = htmlDoc.CreateElement("a");
            embeddedLinkNode.SetAttributeValue("href", filePath);
            HtmlNode imgNode = htmlDoc.CreateElement("img");
            imgNode.SetAttributeValue("src", $"data:image/png;base64, {base64Str}");
            imgNode.SetAttributeValue("alt", Path.GetFileName(filePath));
            imgNode.SetAttributeValue("title", displayName);
            embeddedLinkNode.AppendChild(imgNode);

            htmlDoc.DocumentNode.AppendChild(embeddedLinkNode);

            var outputFilePath = Path.Combine(dir, "embedding.html");
            File.WriteAllText(outputFilePath, htmlDoc.DocumentNode.OuterHtml);
            Console.WriteLine("Done.");

            Console.ReadKey();
        }
    }
}
