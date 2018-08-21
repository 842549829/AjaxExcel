using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;
using ajaxexcel.Models;
using NPOI.HPSF;

namespace ajaxexcel.Controllers
{
    public class ValuesController : ApiController
    {
        private static MemoryStream GetMemoryStream(User user)
        {
            //var users = Models.User.DefaultSection.Where(user);
            var users = Models.User.DefaultSection;
            DataTable dt = ExcleHelper.ToDataTable(users);
            MemoryStream ms = ExcleHelper.BuildWorkbook(dt);
            return ms;
        }

        /// <summary>
        ///  缓存查询条件
        /// </summary>
        /// <param name="user">查询条件</param>
        /// <returns>返回缓存Id</returns>
        [HttpGet]
        [HttpPost]
        [ActionName("exportuser")]
        public User PostExport(User user)
        {
            var key = Guid.NewGuid().ToString();
            user.Id = key;
            HttpRuntimeCache.Set(key, user);
            return user;
        }

        /// <summary>
        /// 下载文件
        /// </summary>
        /// <param name="id">缓存查询条件Id</param>
        /// <returns>返回文件</returns>

        [HttpGet]
        [HttpPost]
        [ActionName("export")]
        public HttpResponseMessage PostExportData(string id)
        {
            // 查询条件
            var user = (User)HttpRuntimeCache.Get(id);
            var file = GetMemoryStream(user);
            //string csv = _service.GetData(model);
            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StreamContent(file)
            };
            //a text file is actually an octet-stream (pdf, etc)
            //result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.ms-excel");
            //we used attachment to force download
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = "file.xls";
            return result;
        }
    }
}