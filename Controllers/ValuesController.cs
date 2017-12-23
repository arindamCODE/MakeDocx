using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace MakeDocx.Controllers
{
    [Route("api/[controller]")]
    public class ValuesController : Controller
    {

        [HttpPost]
        private static byte[] HtmlToWord(string html, string fileName)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(
                memoryStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                if (mainPart == null)
                {
                    mainPart = wordDocument.AddMainDocumentPart();
                    new Document(new Body()).Save(mainPart);
                }

                HtmlConverter converter = new HtmlConverter(mainPart);
                converter.ImageProcessing = ImageProcessing.AutomaticDownload;
                Body body = mainPart.Document.Body;

                IList<OpenXmlCompositeElement> paragraphs = converter.Parse(html);
                body.Append(paragraphs);

                mainPart.Document.Save();
                return memoryStream.ToArray();
            }
        }

        // GET api/values
        /* [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        } */

        // GET api/values/5
        /* [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        } */

        // POST api/values
       /*  [HttpPost]
        public void Post([FromBody]string value)
        {
        } */

        // PUT api/values/5
        /* [HttpPut("{id}")]
        public void Put(int id, [FromBody]string value)
        {
        } */

        // DELETE api/values/5
       /*  [HttpDelete("{id}")]
        public void Delete(int id)
        {
        } */
    }
}
