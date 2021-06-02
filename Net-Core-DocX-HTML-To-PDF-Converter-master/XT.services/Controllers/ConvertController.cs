using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using XT.services.Model;
using static ExampleApplication.src.ConvertPDF;
using dd = ExampleApplication.src;

namespace XT.services.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ConvertController : ControllerBase
    {      

     
        [HttpPost("convert")]
        public IActionResult Post(Root data)
        {
            String response =  dd.ConvertPDF.ConvertingPDF(data);

            ClassResponse.statusResponse status = new ClassResponse.statusResponse();

            status.code = "201";

            status.detail = (response=="ok")?"Documento procesado.": response;

            return CreatedAtAction(nameof(Post), status);

        }

        
    }
}
