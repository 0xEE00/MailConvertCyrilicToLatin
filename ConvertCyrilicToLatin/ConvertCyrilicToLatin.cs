using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Runtime.InteropServices;
using System.Text;

namespace ConvertCyrilicToLatin
{
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("e60ec182-a724-454f-8743-8365d0c4d725")]
    [ComVisible(true)]
    public interface IConvertCyrilicToLatin
    {
        [DispId(1)]
        string CyrilicToLatin(string cyrTextParam);

        [DispId(2)]
        string LatinToCyrilic(string latinTextParam);

        [DispId(3)]
        bool SendMail(string hostParam, string fromParam, string toParam, string subjectParam, bool isHTMLParam);

        [DispId(4)]
        bool SendMailLatToCir(string hostParam, string fromParam, string toParam, string subjectParam, bool isHTMLParam, bool convertSubjectParam);

        [DispId(5)]
        bool SendMailCirToLat(string hostParam, string fromParam, string toParam, string subjectParam, bool isHTMLParam, bool convertSubjectParam);

        [DispId(6)]
        bool AddBody(string bodyTextParam, bool newLine);

        [DispId(7)]
        void AddAtt(string attTextParam);

        [DispId(8)]
        void AddCC(string ccTextParam);

        [DispId(9)]
        void AddBCC(string bccTextParam);

        [DispId(10)]
        void SetImage(string imageUrlParam, int widthParam, int heightParam);
    }

    [Guid("fdaae4ee-1251-4e70-984a-e785dbee9cd4")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("ConvertCyrilicToLatin")]
    public class ConvertCyrilicToLatin : IConvertCyrilicToLatin
    {
        readonly string[] lat_up = { "&#65*", "&#66", "&#67", "&#68", "&#69", "&#70", "&#71", "&#72", "&#73", "&#74", "&#75", "&#76", "&#77", "&#78", "&#79", "&#80", "&#81", "&#82", "&#83", "&#84", "&#85", "&#86", "&#87", "&#88", "&#89", "&#90", "A", "B", "V", "G", "D", "E", "Ë", "Ò", "Z", "I", "Y", "K", "L", "M", "N", "O", "P", "R", "S", "T", "U", "F", "H", "C", "Ô", "Ù", "W", "Ø", "Y", "Ä", "Ê", "Û", "Ü" };
        readonly string[] lat_low = { "&#97", "&#98", "&#99", "&#100", "&#101", "&#102", "&#103", "&#104", "&#105", "&#106", "&#107", "&#108", "&#109", "&#110", "&#111", "&#112", "&#113", "&#114", "&#115", "&#116", "&#117", "&#118", "&#119", "&#120", "&#121", "&#122", "a", "b", "v", "g", "d", "e", "ë", "ò", "z", "i", "j", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "h", "c", "ô", "ù", "w", "ø", "y", "ä", "ê", "û", "ü" };
        readonly string[] rus_up = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "А", "Б", "В", "Г", "Д", "Е", "Ё", "Ж", "З", "И", "Й", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Ъ", "Ы", "Ь", "Э", "Ю", "Я" };
        readonly string[] rus_low = { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "а", "б", "в", "г", "д", "е", "ё", "ж", "з", "и", "й", "к", "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф", "х", "ц", "ч", "ш", "щ", "ъ", "ы", "ь", "э", "ю", "я" };

        readonly string[] lat_up2 = { "A", "B", "V", "G", "D", "E", "Ë", "Ò", "Z", "I", "Y", "K", "L", "M", "N", "O", "P", "R", "S", "T", "U", "F", "H", "C", "Ô", "Ù", "W", "Ø", "Y", "Ä", "Ê", "Û", "Ü", "&#65*", "&#66", "&#67", "&#68", "&#69", "&#70", "&#71", "&#72", "&#73", "&#74", "&#75", "&#76", "&#77", "&#78", "&#79", "&#80", "&#81", "&#82", "&#83", "&#84", "&#85", "&#86", "&#87", "&#88", "&#89", "&#90" };
        readonly string[] lat_low2 = { "a", "b", "v", "g", "d", "e", "ë", "ò", "z", "i", "j", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "h", "c", "ô", "ù", "w", "ø", "y", "ä", "ê", "û", "ü", "&#97", "&#98", "&#99", "&#100", "&#101", "&#102", "&#103", "&#104", "&#105", "&#106", "&#107", "&#108", "&#109", "&#110", "&#111", "&#112", "&#113", "&#114", "&#115", "&#116", "&#117", "&#118", "&#119", "&#120", "&#121", "&#122" };
        readonly string[] rus_up2 = { "А", "Б", "В", "Г", "Д", "Е", "Ё", "Ж", "З", "И", "Й", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Ъ", "Ы", "Ь", "Э", "Ю", "Я", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
        readonly string[] rus_low2 = { "а", "б", "в", "г", "д", "е", "ё", "ж", "з", "и", "й", "к", "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф", "х", "ц", "ч", "ш", "щ", "ъ", "ы", "ь", "э", "ю", "я", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z" };

        Boolean sendStatus = false;
        StringBuilder bodyText = new StringBuilder();
        MailMessage message = new MailMessage();
        public int Width { get; set; }
        public int Height { get; set; }
        public string ImageURL { get; set; }

        public string CyrilicToLatin(string cyrTextParam)
        {
            try
            {
                for (int i = 0; i <= 58; i++)
                {
                    cyrTextParam = cyrTextParam.Replace(rus_up[i], lat_up[i]);
                    cyrTextParam = cyrTextParam.Replace(rus_low[i], lat_low[i]);
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return cyrTextParam;
        }

        public string LatinToCyrilic(string latinTextParam)
        {
            try
            {
                for (int i = 0; i <= 58; i++)
                {
                    latinTextParam = latinTextParam.Replace(lat_up2[i], rus_up2[i]);
                    latinTextParam = latinTextParam.Replace(lat_low2[i], rus_low2[i]);
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return latinTextParam;
        }

        public bool SendMail(string hostParam, string fromParam, string toParam, string subjectParam, bool isHTMLParam)
        {
            return InitAndSendMail(hostParam, fromParam, toParam, subjectParam, isHTMLParam, bodyText.ToString());
        }

        public bool SendMailLatToCir(string hostParam, string fromParam, string toParam, string subjectParam, bool isHTMLParam, bool convertSubjectParam)
        {
            if (convertSubjectParam)
                return InitAndSendMail(hostParam, fromParam, toParam, LatinToCyrilic(subjectParam), isHTMLParam, LatinToCyrilic(bodyText.ToString()));
            else
                return InitAndSendMail(hostParam, fromParam, toParam, subjectParam, isHTMLParam, LatinToCyrilic(bodyText.ToString()));
        }

        public bool SendMailCirToLat(string hostParam, string fromParam, string toParam, string subjectParam, bool isHTML, bool convertSubjectParam)
        {
            if (convertSubjectParam)
                return InitAndSendMail(hostParam, fromParam, toParam, LatinToCyrilic(subjectParam), isHTML, CyrilicToLatin(bodyText.ToString()));
            else
                return InitAndSendMail(hostParam, fromParam, toParam, subjectParam, isHTML, CyrilicToLatin(bodyText.ToString()));
        }

        private bool InitAndSendMail(string hostParam, string fromParam, string toParam, string subjectParam, bool isHTMLParam, string dataParam)
        {
            try
            {
                SmtpClient client = new SmtpClient(hostParam);

                MailAddress from = new MailAddress(fromParam);
                MailAddress to = new MailAddress(toParam);
                if (ImageURL != null)
                {
                    LinkedResource linkedImage = new LinkedResource(ImageURL, MediaTypeNames.Image.Jpeg);
                    linkedImage.ContentId = Guid.NewGuid().ToString();
                    //linkedImage.ContentType = new ContentType(MediaTypeNames.Image.Jpeg);

                    AlternateView htmlView = AlternateView.CreateAlternateViewFromString(ReplaceBr(dataParam) + "<br><br><img src=cid:" + linkedImage.ContentId + " width=" + Width + " height=" + Height + ">", null, MediaTypeNames.Text.Html);
                    htmlView.LinkedResources.Add(linkedImage);
                    message.AlternateViews.Add(htmlView);
                }
                else
                {
                    message.Body = ReplaceBr(dataParam);
                }

                message.From = from;
                message.To.Add(to);
                message.Subject = subjectParam;
                message.SubjectEncoding = Encoding.UTF8;
                message.BodyEncoding = Encoding.UTF8;
                message.IsBodyHtml = isHTMLParam;
                //Commented code!!!!!
                //client.SendCompleted += new SendCompletedEventHandler(SendCompletedCallback);
                //client.SendAsync(message, "");
                client.Send(message);
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            ConvertCyrilicToLatin s = new ConvertCyrilicToLatin();
            String token = (string)e.UserState;
            if (e.Cancelled)
                s.sendStatus = false;
            if (e.Error != null)
                s.sendStatus = false;
            else
                s.sendStatus = true;
        }

        public bool AddBody(string bodyTextParam, bool newLine)
        {
            try
            {
                if (newLine)
                {
                    if (bodyTextParam.Contains("<nobr>"))
                        bodyText.Append(bodyTextParam.Replace("<nobr>", "") + "<br>");
                    else
                        bodyText.AppendLine(bodyTextParam + "<br>");
                }
                else
                {
                    if (bodyTextParam.Contains("<nobr>"))
                        bodyText.Append(bodyTextParam.Replace("<nobr><br>", ""));
                    else
                        bodyText.Append(bodyTextParam);
                }

            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public void AddAtt(string attParam)
        {
            try
            {
                Attachment att = new Attachment(attParam);
                message.Attachments.Add(att);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void AddAtt2(string attParam)
        {
            try
            {
                Attachment att = new Attachment(attParam);
                message.Attachments.Add(att);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void AddAtt3(string attParam)
        {
            try
            {
                Attachment att = new Attachment(attParam);
                message.Attachments.Add(att);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void AddCC(string ccParam)
        {
            try
            {
                message.CC.Add(ccParam);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void AddBCC(string bccParam)
        {
            try
            {
                message.Bcc.Add(bccParam);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SetImage(string imageUrlParam, int widthParam, int heightParam)
        {
            ImageURL = imageUrlParam;
            Width = widthParam;
            Height = heightParam;
        }

        private string ReplaceBr(string cirText)
        {
            return cirText.Replace("<бр>", "<br>");
        }
    }
}
