using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;

namespace navercreatefile
{
    [ComVisible(true),
    GuidAttribute("7210001F-23DC-4393-A311-0B0A0CEA3BF1")]
    [ProgId("NaverMail.CreateFile")]
    public class navercreatefile : Inavercreatefile
    {
        public String createFile(string accesskey, string timestamp, string signature, string fileName)
        {
            string upFileName = fileName.Substring(fileName.LastIndexOf("\\")+1);

            string boundary = "---------------------------" + DateTime.Now.Ticks.ToString("x");
            byte[] boundarybytes = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");

            HttpWebRequest Req = (HttpWebRequest)WebRequest.Create("https://mail.apigw.ntruss.com/api/v1/files");
            Req.Method = "POST";
            Req.KeepAlive = true;
            Req.ContentType = "multipart/form-data; boundary=" + boundary;
            Req.Headers.Add("x-ncp-apigw-timestamp", timestamp);
            Req.Headers.Add("x-ncp-iam-access-key", accesskey);
            Req.Headers.Add("x-ncp-apigw-signature-v2", signature);

            Stream stream = Req.GetRequestStream();
            string formdataTemplate = "Content-Disposition: form-data; name=\"{0}\"\r\n\r\n{1}";
            string formitem = string.Format(formdataTemplate, "fileList", "filename");
            stream.Write(boundarybytes, 0, boundarybytes.Length);
            byte[] formitembytes = System.Text.Encoding.UTF8.GetBytes(formitem);
            stream.Write(formitembytes, 0, formitembytes.Length);
            stream.Write(boundarybytes, 0, boundarybytes.Length);

            string headerTemplate = "Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\"\r\nContent-Type: {2}\r\n\r\n";
            string header = string.Format(headerTemplate, "fileList", upFileName, mimetype(upFileName));
            byte[] headerbytes = System.Text.Encoding.UTF8.GetBytes(header);
            stream.Write(headerbytes, 0, headerbytes.Length);


            FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            byte[] buffer = new byte[4096];
            int bytesRead = 0;
            while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0)
            {
                stream.Write(buffer, 0, bytesRead);
            }
            fileStream.Close();

            byte[] trailer = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "--\r\n");
            stream.Write(trailer, 0, trailer.Length);
            stream.Close();

            try
            {
                WebResponse Res = Req.GetResponse();

                // 응답 Stream 읽기
                Stream stReadData = Res.GetResponseStream();
                StreamReader srReadData = new StreamReader(stReadData, Encoding.UTF8);

                // 응답 Stream -> 응답 String 변환
                string strResult = srReadData.ReadToEnd();

                Res.Close();

                //strResult = strResult.Substring(strResult.LastIndexOf(@"fileId"":""")+1, strResult.LastIndexOf(@"""}]}")-1);
                return strResult;
            }
            catch (Exception ex)
            {
                return ex.Message.ToString() + "++" + upFileName;
            }
        }

        static String mimetype(String fileName)
        {
            String ext = fileName.Substring(fileName.LastIndexOf(".")+1);
            switch (ext) {
                case "jpg": return "image/jpeg";
                case "gif": return "image/gif";
                case "png": return "image/png";
                case "bmp": return "image/bmp";
                case "txt": return "text/plain";
                case "pdf": return "application/pdf";
                case "xls": return "application/vnd.ms-excel";
                case "xlsx": return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                default: return "application/zip";
            }
        }
    }
}
