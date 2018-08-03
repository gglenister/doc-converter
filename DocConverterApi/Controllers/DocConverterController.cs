using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

using DocConverterApi.Helpers;

namespace DocConverterApi.Controllers
{
    [Route("api/[controller]")]
    public class DocConverterController : Controller
    {
        // GET api/DocConverter
        [HttpGet]
        public ActionResult Get()
        {
            return Json(new { message = "I Feel happy!" });
        }

        public class FileConvert
        {
            public string fileLocation { get; set; }
            public string convertTo { get; set; }
        }

        // POST api/DocConverter
        [HttpPost]
        public ActionResult Post([FromBody]FileConvert fileConvert)
        {
            string msg = "error converting document";

            try
            {
                AsposeWordsHelper helper = new AsposeWordsHelper();
                byte[] outFile = null;
                bool hasSignatures = false;
                try
                {
                    hasSignatures = helper.CheckWordFileForDigitalSignatures(fileConvert.fileLocation);
                }
                catch (Exception ex) { }

                if (!hasSignatures)
                {
                    switch (fileConvert.convertTo.ToLower())
                    {
                        case "rtf":
                            outFile = helper.ConvertToRTF(fileConvert.fileLocation);
                            break;
                        case "pdf":
                            outFile = helper.ConvertToPDF(fileConvert.fileLocation);
                            break;
                        case "docx":
                            outFile = helper.ConvertToDOCX(fileConvert.fileLocation);
                            break;
                        default:
                            msg = "this file format is not yet supported";
                            throw new Exception("This file format is not yet supported.");
                    }
                }
                else
                {
                    throw new Exception("This file has signatures and cannot be converted.");
                }
                
                string fileStr = Convert.ToBase64String(outFile);

                return Json(new { file = fileStr });
            }
            catch (Exception ex)
            {
                return BadRequest(msg);
            }
        }
    }
}