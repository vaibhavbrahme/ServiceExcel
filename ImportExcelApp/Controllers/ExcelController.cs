using System.Net;
using System.Threading.Tasks;
using System.Web.Http;
using ImportExcelApp.Models;
using ImportExcelApp.Services;

namespace ImportExcelApp.Controllers
{
    [RoutePrefix("api/excel")]
    public class ExcelController : ApiController
    {
        private readonly IExcelService _excelService;

        public ExcelController()
        {
            _excelService = new ExcelService();
        }

        [HttpPost]
        [Route("import")]
        public async Task<IHttpActionResult> Import([FromBody] ExcelImportRequest request)
        {
            if (request == null)
            {
                return BadRequest("Request body must contain ExcelImportRequest JSON.");
            }

            var result = await _excelService.ImportExcelAsync(request);

            if (result.Success)
            {
                return Ok(result);
            }

            return Content(HttpStatusCode.BadRequest, result);
        }
    }
}