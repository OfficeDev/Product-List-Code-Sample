using Microsoft.SharePoint.Client;
using SharePointAppSampleWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SharePointAppSampleWeb
{
    public static class SharePointService
    {
        public static string GetUserName(SharePointContext spContext)
        {
            string strUserName = null;

            User spUser = null;            

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    spUser = clientContext.Web.CurrentUser;

                    clientContext.Load(spUser, user => user.Title);

                    clientContext.ExecuteQuery();

                    strUserName = spUser.Title;
                }
            }

            return strUserName;
        }

        public static List<Product> GetProducts(SharePointContext spContext, CamlQuery camlQuery)
        {
            List<Product> products = new List<Product>();

            using (var clientContext = spContext.CreateUserClientContextForSPAppWeb())
            {
                if (clientContext != null)
                {
                    List lstProducts = clientContext.Web.Lists.GetByTitle("Products");

                    ListItemCollection lstProductItems = lstProducts.GetItems(camlQuery);

                    clientContext.Load(lstProductItems);

                    clientContext.ExecuteQuery();

                    if (lstProductItems != null)
                    {
                        foreach (var lstProductItem in lstProductItems)
                        {
                            products.Add(
                                new Product
                                {
                                    Id = lstProductItem.Id,
                                    Title = lstProductItem["Title"].ToString(),
                                    Description = lstProductItem["ProductDescription"].ToString(),
                                    Price = lstProductItem["Price"].ToString()
                                }); 
                        }
                    }
                }
            }

            return products;
        }

        public static bool AddProduct(SharePointContext spContext, Product product)
        {
            using (var clientContext = spContext.CreateUserClientContextForSPAppWeb())
            {
                if (clientContext != null)
                {
                    try
                    {
                        List lstProducts = clientContext.Web.Lists.GetByTitle("Products");

                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem newProduct = lstProducts.AddItem(itemCreateInfo);
                        newProduct["Title"] = product.Title;
                        newProduct["ProductDescription"] = product.Description;
                        newProduct["Price"] = product.Price;
                        newProduct.Update();

                        clientContext.ExecuteQuery();

                        return true;
                    }
                    catch (ServerException ex)
                    {
                        return false;
                    }
                }
            }

            return false;
        }

        internal static Product GetProductDetails(SharePointContext spContext, int id)
        {
            Product product = new Product();
            using (var clientContext = spContext.CreateUserClientContextForSPAppWeb())
            {
                if (clientContext != null)
                {
                    List lstProducts = clientContext.Web.Lists.GetByTitle("Products");
                    ListItem selectedItem = lstProducts.GetItemById(id);
                    clientContext.Load(selectedItem);

                    clientContext.ExecuteQuery();

                    if (selectedItem != null)
                    {
                        product = new Product
                        {
                            Title = selectedItem["Title"] as string,
                            Description = selectedItem["ProductDescription"] as string,
                            Price = selectedItem["Price"] as string,
                        };
                    }
                }

            }

            return product;
        }

        public static bool UpdateProduct(SharePointContext spContext, Product product)
        {

            using (var clientContext = spContext.CreateUserClientContextForSPAppWeb())
            {
                if (clientContext != null)
                {
                    try
                    {
                        List lstProducts = clientContext.Web.Lists.GetByTitle("Products");
                        ListItem selectedItem = lstProducts.GetItemById(product.Id);

                        selectedItem["Title"] = product.Title;
                        selectedItem["ProductDescription"] = product.Description;
                        selectedItem["Price"] = product.Price;
                        selectedItem.Update();

                        clientContext.ExecuteQuery();
                        return true;

                    }
                    catch (ServerException ex)
                    {
                        return false;
                    }

                }
            }
            return false;
        }

        public static void DeleteProduct(SharePointContext spContext, Product product)
        {

            using (var clientContext = spContext.CreateUserClientContextForSPAppWeb())
            {
                try
                {
                    List productsList = clientContext.Web.Lists.GetByTitle("Products");
                    ListItem itemToDelete = productsList.GetItemById(product.Id);
                    itemToDelete.DeleteObject();

                    clientContext.ExecuteQuery();
                }
                catch (ServerException ex)
                {
                    // TODO: Exception Handling
                }
            }
        }
    
    }
}