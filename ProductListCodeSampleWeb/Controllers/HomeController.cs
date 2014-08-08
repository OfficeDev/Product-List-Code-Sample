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

        [SharePointContextFilter]
        public ActionResult Edit(int id)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            Product product = SharePointService.GetProductDetails(spContext, id);
            return View(product);
        }

        [SharePointContextFilter]
        [HttpPost]
        public ActionResult Edit(Product product)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            SharePointService.UpdateProduct(spContext, product);
            return RedirectToAction("Index", new { SPHostUrl = SharePointContext.GetSPHostUrl(HttpContext.Request).AbsoluteUri });
        }

        [SharePointContextFilter]
        public ActionResult Delete(int id)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            Product product = SharePointService.GetProductDetails(spContext, id);
            return View(product);
        }

        [SharePointContextFilter]
        [HttpPost]
        public ActionResult Delete(Product product)
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            SharePointService.DeleteProduct(spContext, product);
            return RedirectToAction("Index", new { SPHostUrl = SharePointContext.GetSPHostUrl(HttpContext.Request).AbsoluteUri });
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
