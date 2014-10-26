using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Threading.Tasks;
using Microsoft.Office365.OAuth;
using Microsoft.Office365.SharePoint;
using O365APIweb;

namespace O365APIweb.Controllers
{
    public class O365SitesController : Controller
    {
        /// <summary>
        /// Displays the contents of a user's OneDrive For Business "Shared With Everyone" folder.
        /// Minimal permission required: permission to read users' files.
        /// </summary>
        public async Task<ActionResult> Index()
        {
            try
            {
                //Call to get files and load into view
                var files = await SitesApiSample.GetDefaultDocumentFiles();

                return View(files);
            }
            catch (AuthenticationFailedException e)
            {
                ViewBag.ErrorMessage = e.ErrorDescription ?? e.ErrorCode;
                return View("Office365Error");
            }    
            //***Make sure you call this on the Web implementation, to prevent the redirect exception from occurring
            catch (RedirectRequiredException ex)
            {
                return Redirect(ex.RedirectUri.ToString());
            }
        }

        [HttpPost]
        public async Task<ActionResult> UploadFile()
        {
            var isSuccess = false;
            foreach(string fileName in Request.Files)
            {
                HttpPostedFileBase hpf = Request.Files[fileName];

                if (hpf.ContentLength == 0)
                    continue;
                isSuccess = await SitesApiSample.Uploaddoc(hpf);       
            }
                
            return RedirectToAction("Index");
        }
    }
}