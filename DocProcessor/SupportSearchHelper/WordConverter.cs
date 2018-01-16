using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;
using SupportSearch.Common.Dtos;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SupportSearch.Helper
{
    public class WordConverter
    {
        //protected SupportSearch.Common.Interfaces.ISupportBlobUpload _blobUploader = new SupportSearch.Helper.AzureBlobUploadHelper();
        //protected SupportSearch.Common.Interfaces.ISupportKeyWords _mlKeyWordHelper = new SupportSearch.Helper.AzureCSKeyWordHandler();

        protected SupportSearch.Common.Interfaces.ISupportBlobUpload _blobUploader = new SupportSearch.Helper.AzureFileUploadHelper();
        protected SupportSearch.Common.Interfaces.ISupportKeyWords _mlKeyWordHelper = new SupportSearch.Helper.AzureCSKeyWordHandler();

        public WordConverter(SupportSearch.Common.Interfaces.ISupportBlobUpload blobUploader, SupportSearch.Common.Interfaces.ISupportKeyWords keyWordsHelper)
        {
            _blobUploader = blobUploader;
            _mlKeyWordHelper = keyWordsHelper;
        }

        public string ConvertDocToHtml(string DocumentName, Stream WordDocument)
        {
            string result = "";
            WordprocessingDocument doc = WordprocessingDocument.Open(WordDocument, false);

            WordProcessingDocDto wordDto = new WordProcessingDocDto(DocumentName);
            GetStyles(wordDto, doc);

            ParseDocument(wordDto, doc);
            foreach (WordProcessingDocPart docPart in wordDto.WordProcessingDocParts)
            {
                FillAndWritePages(DocumentName, docPart);
            }

            // Write QnAMaker output
            WriteQnAMakerOutput(wordDto);

            return result;
        }

        protected void GetStyles(WordProcessingDocDto wordDto, WordprocessingDocument doc)
        {
            // First parse styles, get header info cached
            foreach (OpenXmlElement style in doc.MainDocumentPart.StyleDefinitionsPart.Styles.ChildElements)
            {
                string styleId = "";
                foreach (OpenXmlAttribute attr in style.GetAttributes())
                {
                    if (attr.LocalName == "styleId")
                    {
                        styleId = attr.Value;
                    }
                }

                if (styleId != "")
                {
                    foreach (OpenXmlElement styleElem in style.ChildElements)
                    {
                        if (styleElem.LocalName == "pPr")
                        {
                            foreach (OpenXmlElement prElem in styleElem.ChildElements)
                            {
                                if (prElem.LocalName == "outlineLvl")
                                {

                                    if (((OutlineLevel)prElem).Val == 0)
                                        wordDto.Heading1StyleIds.Add(styleId);
                                    else if (((OutlineLevel)prElem).Val == 1)
                                        wordDto.Heading2StyleIds.Add(styleId);
                                    else if (((OutlineLevel)prElem).Val == 2)
                                        wordDto.Heading3StyleIds.Add(styleId);
                                }
                            }
                        }
                    }
                }
            }
        }

        protected void ParseDocument(WordProcessingDocDto wordDto, WordprocessingDocument doc)
        {
            WordProcessingDocPart lastHeader = null;
            bool listRunning = false;
            // First get the top level headers
            foreach (OpenXmlElement docElem in doc.MainDocumentPart.RootElement.Descendants<Body>().First().ChildElements)
            {
                bool foundElementMeaning = false;

                foreach (ParagraphProperties subElem in docElem.Descendants<ParagraphProperties>())
                {
                    foreach (ParagraphStyleId styleElem in subElem.Descendants<ParagraphStyleId>())
                    {
                        if (wordDto.Heading1StyleIds.Contains(styleElem.Val))
                        {
                            // We have a top level header!!
                            foundElementMeaning = true;
                            WordProcessingDocPart header1 = new WordProcessingDocPart(HeaderType.Header1);
                            lastHeader = header1;
                            foreach (Run runElem in docElem.Descendants<Run>())
                            {
                                header1.HeaderText += runElem.InnerText;
                            }

                            header1.XmlElement = XElement.Parse(docElem.OuterXml);
                            wordDto.WordProcessingDocParts.Add(header1);
                            header1.HeaderNumber = wordDto.WordProcessingDocParts.Count.ToString();
                        }
                        else if (wordDto.Heading2StyleIds.Contains(styleElem.Val))
                        {
                            // We have a header 2!
                            foundElementMeaning = true;
                            WordProcessingDocPart header2 = new WordProcessingDocPart(HeaderType.Header2);
                            header2.ParentHeader1 = wordDto.WordProcessingDocParts.Last();
                            lastHeader = header2;
                            foreach (Run runElem in docElem.Descendants<Run>())
                            {
                                header2.HeaderText += runElem.InnerText;
                            }
                            header2.XmlElement = XElement.Parse(docElem.OuterXml);
                            wordDto.WordProcessingDocParts.Last().SubHeaders.Add(header2);
                            header2.HeaderNumber = $"{wordDto.WordProcessingDocParts.Count.ToString()}.{wordDto.WordProcessingDocParts.Last().SubHeaders.Count.ToString()}";
                        }
                        else if (wordDto.Heading3StyleIds.Contains(styleElem.Val))
                        {
                            // We have a header 3!
                            foundElementMeaning = true;
                            WordProcessingDocPart header3 = new WordProcessingDocPart(HeaderType.Header3);
                            header3.ParentHeader1 = wordDto.WordProcessingDocParts.Last();
                            header3.ParentHeader2 = wordDto.WordProcessingDocParts.Last().SubHeaders.Last();

                            lastHeader = header3;
                            foreach (Run runElem in docElem.Descendants<Run>())
                            {
                                header3.HeaderText += runElem.InnerText;
                            }
                            header3.XmlElement = XElement.Parse(docElem.OuterXml);
                            wordDto.WordProcessingDocParts.Last().SubHeaders.Last().SubHeaders.Add(header3);
                            header3.HeaderNumber = $"{wordDto.WordProcessingDocParts.Count.ToString()}.{wordDto.WordProcessingDocParts.Last().SubHeaders.Count.ToString()}.{wordDto.WordProcessingDocParts.Last().SubHeaders.Last().SubHeaders.Count.ToString()}";
                        }
                        else if (wordDto.Heading4StyleIds.Contains(styleElem.Val))
                        {
                            // We have a header 4!
                            foundElementMeaning = true;
                            WordProcessingDocPart header4 = new WordProcessingDocPart(HeaderType.Header4);
                            header4.ParentHeader1 = wordDto.WordProcessingDocParts.Last();
                            header4.ParentHeader2 = wordDto.WordProcessingDocParts.Last().SubHeaders.Last();
                            header4.ParentHeader3 = wordDto.WordProcessingDocParts.Last().SubHeaders.Last().SubHeaders.Last();

                            lastHeader = header4;
                            foreach (Run runElem in docElem.Descendants<Run>())
                            {
                                header4.HeaderText += runElem.InnerText;
                            }
                            header4.XmlElement = XElement.Parse(docElem.OuterXml);
                            wordDto.WordProcessingDocParts.Last().SubHeaders.Last().SubHeaders.Last().SubHeaders.Add(header4);
                            header4.HeaderNumber = $"{wordDto.WordProcessingDocParts.Count.ToString()}.{wordDto.WordProcessingDocParts.Last().SubHeaders.Count.ToString()}.{wordDto.WordProcessingDocParts.Last().SubHeaders.Last().SubHeaders.Count.ToString()}.{wordDto.WordProcessingDocParts.Last().SubHeaders.Last().SubHeaders.Last().SubHeaders.Count.ToString()}";
                        }
                        else if (styleElem.Val.Value.ToUpper().Contains("BOLD"))
                        {
                            // We have a header bold subheaders!
                            foundElementMeaning = true;
                            WordProcessingDocPart headerSub = new WordProcessingDocPart(HeaderType.SubHeader);

                            foreach (Run runElem in docElem.Descendants<Run>())
                            {
                                headerSub.HeaderText += runElem.InnerText;
                            }

                            headerSub.XmlElement = XElement.Parse(docElem.OuterXml);

                            if (lastHeader != null)
                            {
                                headerSub.HeaderNumber = lastHeader.HeaderNumber;

                                if (lastHeader.Type == HeaderType.Header1)
                                    headerSub.ParentHeader1 = wordDto.WordProcessingDocParts.Last();
                                else if (lastHeader.Type == HeaderType.Header2)
                                {
                                    headerSub.ParentHeader1 = wordDto.WordProcessingDocParts.Last();
                                    headerSub.ParentHeader2 = wordDto.WordProcessingDocParts.Last().SubHeaders.Last();
                                }
                                else if (lastHeader.Type == HeaderType.Header3)
                                {
                                    headerSub.ParentHeader1 = wordDto.WordProcessingDocParts.Last();
                                    headerSub.ParentHeader2 = wordDto.WordProcessingDocParts.Last().SubHeaders.Last();
                                    headerSub.ParentHeader3 = wordDto.WordProcessingDocParts.Last().SubHeaders.Last().SubHeaders.Last();
                                }
                                else if (lastHeader.Type == HeaderType.Header4)
                                {
                                    headerSub.ParentHeader1 = wordDto.WordProcessingDocParts.Last();
                                    headerSub.ParentHeader2 = wordDto.WordProcessingDocParts.Last().SubHeaders.Last();
                                    headerSub.ParentHeader3 = wordDto.WordProcessingDocParts.Last().SubHeaders.Last().SubHeaders.Last();
                                    headerSub.ParentHeader4 = wordDto.WordProcessingDocParts.Last().SubHeaders.Last().SubHeaders.Last().SubHeaders.Last();
                                }

                                lastHeader.SubHeaders.Add(headerSub);
                            }
                        }
                    }
                }

                if (!foundElementMeaning && lastHeader != null)
                {
                    // This means that this p is not a header, so do the loop to find out what it has
                    WordProcessingDocPart lastPart = (lastHeader != null && lastHeader.SubHeaders.Count == 0) ? lastHeader : lastHeader.SubHeaders.Last();

                    // Check if this is a list
                    if (docElem.Descendants<NumberingLevelReference>().Count() > 0 && !listRunning)
                    {
                        listRunning = true;
                        lastPart.Content.Add("UL");
                    }
                    else if (docElem.Descendants<NumberingLevelReference>().Count() == 0 && listRunning)
                    {
                        listRunning = false;
                        lastPart.Content.Add("/UL");
                    }

                    string runText = "";
                    // Check if this is a table
                    if (docElem.GetType() == typeof(DocumentFormat.OpenXml.Wordprocessing.Table))
                    {
                        DocumentFormat.OpenXml.Wordprocessing.Table table = (DocumentFormat.OpenXml.Wordprocessing.Table)docElem;
                        foreach (var row in table.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRow>())
                        {
                            if (runText == "")
                            {
                                // We don't yet have headers, so create them
                                runText += "<table><colgroup>";
                                int colCount = row.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Count();
                                int colWidth = 100 / colCount;
                                foreach (var headerCell in row.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>())
                                {
                                    runText += $"<col span='1' style='width: {colWidth}%;'>";
                                }
                                runText += "</colgroup><thead><tr>";

                                foreach (var headerCell in row.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>())
                                {
                                    string cellText = "";
                                    foreach (Run runElem in headerCell.Descendants<Run>())
                                    {
                                        cellText += $" {runElem.InnerText}";
                                    }
                                    runText += $"<th>{cellText}</th>";
                                }
                                runText += "</tr></thead><tbody>";
                            }
                            else
                            {
                                // Add normal rows
                                runText += "<tr>";
                                foreach (var headerCell in row.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>())
                                {
                                    string cellText = "";
                                    foreach (Run runElem in headerCell.Descendants<Run>())
                                    {
                                        cellText += $" {runElem.InnerText}";
                                    }
                                    runText += $"<td>{cellText}</td>";
                                }
                                runText += "</tr>";
                            }
                        }

                        runText += "</tbody></table>";
                    }
                    else {

                        // So this is content, must be added to the last header
                        foreach (Run runElem in docElem.Descendants<Run>())
                        {
                            runText += $"{runElem.InnerText}";

                            // Test for images
                            foreach (DocumentFormat.OpenXml.Drawing.Blip b in runElem.Descendants<DocumentFormat.OpenXml.Drawing.Blip>())
                            {
                                ImagePart imagePart = (ImagePart)doc.MainDocumentPart.GetPartById(b.Embed.Value);
                                Stream stream = imagePart.GetStream();
                                long length = stream.Length;
                                byte[] byteStream = new byte[length];
                                stream.Read(byteStream, 0, (int)length);

                                //string fileName = $"C:\\Temp\\docinfos3\\images\\image-{b.Embed.Value}.jpg";
                                string fileName = $"image-{wordDto.DocumentName}-{b.Embed.Value}.jpg";

                                MemoryStream imageStream = new MemoryStream();
                                imageStream.Write(byteStream, 0, (int)length);
                                imageStream.Position = 0;
                                //imageStream.Close();
                                string imageUri = _blobUploader.UploadDocImage(fileName, imageStream);
                                imageStream.Close();
                                runText += Environment.NewLine + Environment.NewLine + $"IMAGE_{lastPart.ImagePaths.Count}" + Environment.NewLine + Environment.NewLine;
                                lastPart.ImagePaths.Add(imageUri);
                            }
                        }
                    }

                    lastPart.Content.Add(runText);
                }
            }
        }

        protected void FillAndWritePages(string DocTitle, WordProcessingDocPart docPart)
        {
            string keyWordText = GetTitles(docPart);

            // Only fill pages if header keywords are found
            if (!String.IsNullOrEmpty(keyWordText))
            {
                string keyWords = _mlKeyWordHelper.GetKeyWords(keyWordText);
                //JObject jsonResult = JObject.Parse(keyWords);
                //docPart.KeyWords = jsonResult["Results"]["output1"][0]["Preprocessed Headers"].Value<string>().Replace("|", "");
                docPart.KeyWords = keyWords;

                if (docPart.Content.Count > 0)
                {
                    // Text Output
                    docPart.TextOutput = GetText(DocTitle, docPart);

                    // Html Output
                    docPart.HtmlOutput = GetHtml(DocTitle, docPart);
                    string fileName = DocTitle + "-" + docPart.HeaderNumber + "-" + docPart.HeaderText + ".html";
                    var invalids = System.IO.Path.GetInvalidFileNameChars();
                    var newName = String.Join("_", fileName.Split(invalids, StringSplitOptions.RemoveEmptyEntries)).TrimEnd('.');
                    _blobUploader.UploadDocHtml(newName, docPart.HtmlOutput, docPart.KeyWords, docPart.TextOutput);

                    // Markdown Output
                    docPart.MarkdownOutput = GetMarkdown(DocTitle, docPart);
                    fileName = DocTitle + "-" + docPart.HeaderNumber + "-" + docPart.HeaderText + docPart.HeaderText + ".md";
                    newName = String.Join("_", fileName.Split(invalids, StringSplitOptions.RemoveEmptyEntries)).TrimEnd('.');
                    _blobUploader.UploadDocMarkdown(newName, docPart.MarkdownOutput, docPart.KeyWords, docPart.TextOutput);
                }

                if (docPart.SubHeaders.Count > 0)
                {
                    foreach (WordProcessingDocPart doc in docPart.SubHeaders)
                        FillAndWritePages(DocTitle, doc);
                }
            }
        }

        protected string GetHtml(string DocTitle, WordProcessingDocPart docPart)
        {
            string result = $"<html><head><title>{DocTitle} {docPart.HeaderText}</title></head><body>";
            if (docPart.ParentHeader1 != null) result += $"<h1>{docPart.ParentHeader1.HeaderText}</h1>";
            if (docPart.ParentHeader2 != null) result += $"<h2>{docPart.ParentHeader2.HeaderText}</h2>";
            if (docPart.ParentHeader3 != null) result += $"<h3>{docPart.ParentHeader3.HeaderText}</h3>";
            if (docPart.ParentHeader4 != null) result += $"<h4>{docPart.ParentHeader4.HeaderText}</h4>";

            if (docPart.ParentHeader1 == null) result += $"<h1>{docPart.HeaderText}</h1>";
            else if (docPart.ParentHeader2 == null) result += $"<h2>{docPart.HeaderText}</h2>";
            else if (docPart.ParentHeader3 == null) result += $"<h3>{docPart.HeaderText}</h3>";
            else if (docPart.ParentHeader4 == null) result += $"<h4>{docPart.HeaderText}</h4>";
            else result += $"<h5>{docPart.HeaderText}</h5>";

            result += "<div>";
            bool inList = false;
            foreach (string con in docPart.Content)
            {
                if (con == "UL")
                {
                    inList = true;
                    result += "<div><ul>";
                }
                else if (con == "/UL")
                {
                    inList = false;
                    result += "</ul></div>";
                }
                else if (inList)
                {
                    result += $"<li>{con}</li>";
                }
                else
                {
                    result += $"{con}";
                }
            }

            for (int i = 0; i < docPart.ImagePaths.Count; i++)
            {
                result = result.Replace($"IMAGE_{i}", $"<div><img src='{docPart.ImagePaths[i]}' /></div>");
            }

            //            result += $"</div><br><b>Produkt: {DocTitle}</b><br><div style='font-style: italic;'>Link zur Dokumentation: <a href='https://microsoft-my.sharepoint.com/personal/tyayers_microsoft_com/_layouts/15/guestaccess.aspx?guestaccesstoken=1zHY2jPHDer%2f6aQLEDvn9Yje41gJjLm836mb%2fqcK%2brk%3d&docid=2_165c8b95ea4e741a890fefa0170af63f0&rev=1'>{DocTitle} - {docPart.HeaderNumber} - {docPart.HeaderText}</a></div><br><div>Schlüsselwörter:";
            foreach (string keyWord in docPart.KeyWords.Split(' '))
            {
                result += $" <a href='https://azuresearchtester20170716114728.azurewebsites.net/?search={keyWord}'  target='_blank'>{keyWord}</a>";
            }

            //            result += $"<div><b>Benutzerbewertung: 3</b></div>";
            result += "</div></body></html>";

            return result;
        }

        protected string GetMarkdown(string DocTitle, WordProcessingDocPart docPart)
        {
            string result = Environment.NewLine;
            if (docPart.ParentHeader1 != null) result += $"# {docPart.ParentHeader1.HeaderText}" + System.Environment.NewLine;
            if (docPart.ParentHeader2 != null) result += $"## {docPart.ParentHeader2.HeaderText}" + System.Environment.NewLine;
            if (docPart.ParentHeader3 != null) result += $"### {docPart.ParentHeader3.HeaderText}" + System.Environment.NewLine;
            if (docPart.ParentHeader4 != null) result += $"#### {docPart.ParentHeader4.HeaderText}" + System.Environment.NewLine;

            if (docPart.ParentHeader1 == null) result += $"# {docPart.HeaderText}" + System.Environment.NewLine;
            else if (docPart.ParentHeader2 == null) result += $"## {docPart.HeaderText}" + System.Environment.NewLine;
            else if (docPart.ParentHeader3 == null) result += $"### {docPart.HeaderText}" + System.Environment.NewLine;
            else if (docPart.ParentHeader4 == null) result += $"#### {docPart.HeaderText}" + System.Environment.NewLine;
            else result += $"##### {docPart.HeaderText}" + System.Environment.NewLine;

            result += "---" + System.Environment.NewLine + Environment.NewLine;

            bool inList = false;
            foreach (string con in docPart.Content)
            {
                if (con == "UL")
                {
                    inList = true;
                }
                else if (con == "/UL")
                {
                    inList = false;
                    result += Environment.NewLine + Environment.NewLine;
                }
                else if (inList)
                {
                    result += Environment.NewLine + $"* {con}";
                }
                else
                {
                    result += Environment.NewLine + $"{con}";
                }
            }

            for (int i = 0; i < docPart.ImagePaths.Count; i++)
            {
                result = result.Replace($"IMAGE_{i}", $"![]({docPart.ImagePaths[i]})" + Environment.NewLine + Environment.NewLine);
            }

            //result += $"<br><br>**Produkt:** {DocTitle}<br>**Link zur Dokumentation:** [{DocTitle} - {docPart.HeaderNumber} - {docPart.HeaderText}](https://microsoft-my.sharepoint.com/personal/tyayers_microsoft_com/_layouts/15/guestaccess.aspx?guestaccesstoken=1zHY2jPHDer%2f6aQLEDvn9Yje41gJjLm836mb%2fqcK%2brk%3d&docid=2_165c8b95ea4e741a890fefa0170af63f0&rev=1)<br>**Schlüsselwörter:**";
            foreach (string keyWord in docPart.KeyWords.Split(' '))
            {
                result += $" [{keyWord}](https://azuresearchtester20170716114728.azurewebsites.net/?search={keyWord})";
            }

            //result += $"<br>**Benutzerbewertung: 3**";

            return result;
        }

        protected string GetText(string DocTitle, WordProcessingDocPart docPart)
        {
            string result = "";
            if (docPart.ParentHeader1 != null) result += $" {docPart.ParentHeader1.HeaderText}";
            if (docPart.ParentHeader2 != null) result += $" {docPart.ParentHeader2.HeaderText}";
            if (docPart.ParentHeader3 != null) result += $" {docPart.ParentHeader3.HeaderText}";
            if (docPart.ParentHeader4 != null) result += $" {docPart.ParentHeader4.HeaderText}";

            if (docPart.ParentHeader1 == null) result += $" {docPart.HeaderText}";
            else if (docPart.ParentHeader2 == null) result += $" {docPart.HeaderText}";
            else if (docPart.ParentHeader3 == null) result += $" {docPart.HeaderText}";
            else if (docPart.ParentHeader4 == null) result += $" {docPart.HeaderText}";
            else result += $" {docPart.HeaderText}";

            bool inList = false;
            foreach (string con in docPart.Content)
            {
                if (con == "UL")
                {
                    inList = true;
                }
                else if (con == "/UL")
                {
                    inList = false;
                }
                else if (inList)
                {
                    result += $" {con}";
                }
                else
                {
                    result += $" {con}";
                }
            }

            return result;
        }

        protected string GetTitles(WordProcessingDocPart part)
        {
            string result = part.HeaderText;
            if (part.ParentHeader5 != null) result += $" {part.ParentHeader5.HeaderText}";
            if (part.ParentHeader4 != null) result += $" {part.ParentHeader4.HeaderText}";
            if (part.ParentHeader3 != null) result += $" {part.ParentHeader3.HeaderText}";
            if (part.ParentHeader2 != null) result += $" {part.ParentHeader2.HeaderText}";
            if (part.ParentHeader1 != null) result += $" {part.ParentHeader1.HeaderText}";

            return result;
        }

        protected void WriteQnAMakerOutput(WordProcessingDocDto doc)
        {
            string result = "Question\tAnswer\tSource";

            foreach (WordProcessingDocPart docPart in doc.WordProcessingDocParts)
            {
                result += GetQnAMakerPart(docPart);
            }

            _blobUploader.UploadDocQnAMaker("QnAMakerOutput", result);
        }

        protected string GetQnAMakerPart(WordProcessingDocPart docPart)
        {
            string result = Environment.NewLine + docPart.KeyWords + "\t" + docPart.KeyWords + "\tEditorial";

            if (docPart.SubHeaders.Count > 0)
            {
                foreach (WordProcessingDocPart doc in docPart.SubHeaders)
                    result += GetQnAMakerPart(doc);
            }

            return result;
        }
    }
}
