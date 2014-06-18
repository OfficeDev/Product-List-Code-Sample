using Microsoft.SharePoint.Client;
using SharePointAppSampleWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;

namespace SharePointAppSampleWeb.Controllers
{   
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            ViewBag.Username = SharePointService.GetUserName(spContext);

            CamlQuery queryProducts = new CamlQuery();
            queryProducts.ViewXml = @"<View><ViewFields><FieldRef Name='Title'/>
                                        <FieldRef Name='ProductDescription'/>
                                    <FieldRef Name='Price'/></ViewFields></View>";

            List<Product> products = SharePointService.GetProducts(spContext, queryProducts);

            return View(products);
        }
        
        [HttpPost]
        [SharePointContextFilter]
        public ActionResult AddProduct(string title,string description,string price)
        {
            HttpStatusCodeResult httpCode = new HttpStatusCodeResult(HttpStatusCode.MethodNotAllowed);

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            Product newProduct = new Product();
            newProduct.Title = title;
            newProduct.Description = description;
            newProduct.Price = price;

            if (SharePointService.AddProduct(spContext, newProduct))
            {
                httpCode = new HttpStatusCodeResult(HttpStatusCode.Created);
            }

            return httpCode;
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}
