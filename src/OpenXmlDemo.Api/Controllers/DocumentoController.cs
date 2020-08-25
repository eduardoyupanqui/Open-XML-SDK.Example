using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using Microsoft.Net.Http.Headers;

namespace OpenXmlDemo.Api.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class DocumentoController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<DocumentoController> _logger;

        public DocumentoController(ILogger<DocumentoController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public IEnumerable<WeatherForecast> Get()
        {
            var rng = new Random();
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = rng.Next(-20, 55),
                Summary = Summaries[rng.Next(Summaries.Length)]
            })
            .ToArray();
        }

        [HttpGet("GetDocumento")]
        public ActionResult GetDocumento()
        {
            string sourceFile = Path.Combine("D:\\Word10-Prueba.docx");
            var arrayBytes = System.IO.File.ReadAllBytes(sourceFile);
            MemoryStream stream = new MemoryStream(arrayBytes, true);
            //using (MemoryStream stream = new MemoryStream(arrayBytes, true))
            //{
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
            {
                var document_ = wordDoc.MainDocumentPart.Document;
                using (OpenXmlReader reader = OpenXmlReader.Create(document_))
                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(Text))
                        {
                            Text text = (Text)reader.LoadCurrentElement();
                            //text = element.InnerText;
                            //text = reader.GetText();
                            if (text.Text.Contains("@AQUI_SALUDO"))
                            {
                                Regex regexText = new Regex(@"@AQUI_SALUDO");
                                text.Text = regexText.Replace(text.Text, "Hi Everyone!");
                                break;//Quitar si se quiere en mas de un @AQUI_SALUDO
                            }
                        }
                    }
            }
            stream.Seek(0, SeekOrigin.Begin);
            //Opcion 1
            return File(stream, "application/octet-stream", "download.docx"); //Return to download file
            //Opcion 2
            //return new FileStreamResult(stream, new MediaTypeHeaderValue("application/octet-stream"))
            //{
            //    FileDownloadName = "download.docx"
            //};
            //}
        }

        [HttpGet("GetDocumento2")]
        public ActionResult GetDocumento2()
        {
            string sourceFile = Path.Combine("D:\\Word10-Prueba.docx");
            var arrayBytes = System.IO.File.ReadAllBytes(sourceFile);
            MemoryStream stream = new MemoryStream(arrayBytes, true);
            //using (MemoryStream stream = new MemoryStream(arrayBytes, true))
            //{
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
            {
                var document_ = wordDoc.MainDocumentPart.Document;
                using (OpenXmlReader reader = OpenXmlReader.Create(document_))
                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(Text))
                        {
                            Text text = (Text)reader.LoadCurrentElement();
                            //text = element.InnerText;
                            //text = reader.GetText();
                            if (text.Text.Contains("@AQUI_SALUDO"))
                            {
                                Regex regexText = new Regex(@"@AQUI_SALUDO");
                                text.Text = regexText.Replace(text.Text, "Hi Everyone!");
                                break;//Quitar si se quiere en mas de un @AQUI_SALUDO
                            }
                        }
                    }
            }
            stream.Seek(0, SeekOrigin.Begin);
            //Opcion 1
            //return File(stream, "application/octet-stream", "download.docx"); //Return to download file
            //Opcion 2
            return new FileStreamResult(stream, new MediaTypeHeaderValue("application/octet-stream"))
            {
                FileDownloadName = "download.docx"
            };
            //}
        }
    }
}
