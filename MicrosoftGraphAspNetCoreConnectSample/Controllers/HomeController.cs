/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using MicrosoftGraphAspNetCoreConnectSample.Helpers;
using System.Security.Claims;
using System.Threading.Tasks;

namespace MicrosoftGraphAspNetCoreConnectSample.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly IHostingEnvironment _env;
        private readonly IGraphSdkHelper _graphSdkHelper;

        public HomeController(IConfiguration configuration, IHostingEnvironment hostingEnvironment, IGraphSdkHelper graphSdkHelper)
        {
            _configuration = configuration;
            _env = hostingEnvironment;
            _graphSdkHelper = graphSdkHelper;
        }

        //[AllowAnonymous]
        // Load user's profile.
        public async Task<IActionResult> Index(string email)
        {
            if (User.Identity.IsAuthenticated)
            {
                // Get users's email.
                email = email ?? User.FindFirst("preferred_username")?.Value;
                ViewData["Email"] = email;

                // Initialize the GraphServiceClient.
                var graphClient = _graphSdkHelper.GetAuthenticatedClient((ClaimsIdentity)User.Identity);

                ViewData["Response"] = await GraphService.GetUserJson(graphClient, email, HttpContext);

                ViewData["Picture"] = await GraphService.GetPictureBase64(graphClient, email, HttpContext);
            }

            return View();
        }

        //[Authorize]
        [HttpPost]
        // Send an email message from the current user.
        public async Task<IActionResult> SendEmail(string recipients)
        {
            if (string.IsNullOrEmpty(recipients))
            {
                TempData["Message"] = "Please add a valid email address to the recipients list!";
                return RedirectToAction("Index");
            }

            try
            {
                // Initialize the GraphServiceClient.
                var graphClient = _graphSdkHelper.GetAuthenticatedClient((ClaimsIdentity)User.Identity);

                // Send the email.
                await GraphService.SendEmail(graphClient, _env, recipients, HttpContext);

                // Reset the current user's email address and the status to display when the page reloads.
                TempData["Message"] = "Success! Your mail was sent.";
                return RedirectToAction("Index");
            }
            catch (ServiceException se)
            {
                if (se.Error.Code == "Caller needs to authenticate.")
                {
                    return new EmptyResult();
                }

                return RedirectToAction("Error", "Home", new { message = "Error: " + se.Error.Message });
            }
        }

        //[AllowAnonymous]
        public IActionResult About()
        {
            return View();
        }

        //[AllowAnonymous]
        public IActionResult Contact()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [AllowAnonymous]
        public IActionResult Error()
        {
            return View();
        }

        //[Authorize]
        // Send an email message from the current user.
        public async Task<IActionResult> ReadEmails()
        {
            try
            {
                // Initialize the GraphServiceClient.
                var graphClient = _graphSdkHelper.GetAuthenticatedClient((ClaimsIdentity)User.Identity);

                //var mailResults = await graphClient.Me.MailFolders.Inbox.Messages.Request()
                var mailResults = await graphClient.Me.Messages.Request()
                    .OrderBy("receivedDateTime DESC")
                    .Select("*")
                    .Top(10)
                    //.Skip(10)
                    //.Filter("receivedDateTime ge 1900-01-01T00:00:00Z and hasAttachments eq true")
                    .Filter("receivedDateTime ge 2019-04-03 and hasAttachments eq true")
                    .Expand("attachments")
                    .GetAsync();


                // Send the email.
                //await GraphService.SendEmail(graphClient, _env, recipients, HttpContext);

                // Reset the current user's email address and the status to display when the page reloads.
                //TempData["mailResults"] = mailResults.CurrentPage;
                //TempData["Message"] = "OK";
                //return RedirectToAction("mailResults");
                return View(mailResults.CurrentPage);
            }
            catch (ServiceException se)
            {
                if (se.Error.Code == "Caller needs to authenticate.")
                {
                    return new EmptyResult();
                }

                return RedirectToAction("Error", "Home", new { message = "Error: " + se.Error.Message });
            }
        }

        /*
             Item that is attached to the message could be requested like this (documentation)(https://docs.microsoft.com/en-us/graph/api/attachment-get?view=graph-rest-1.0#request-2):
             var attachmentRequest = graphClient.Me.MailFolders.Inbox.Messages[message.Id]
            .Attachments[attachment.Id].Request().Expand("microsoft.graph.itemattachment/item").GetAsync();
            var itemAttachment = (ItemAttachment)attachmentRequest.Result;
            var itemMessage = (Message) itemAttachment.Item;  //get attached message
            Console.WriteLine(itemMessage.Body);  //print message body

            Example
            Demonstrates how to get attachments and save it into file if attachment is a file and read the attached message if attachment is an item:

            var request = graphClient.Me.MailFolders.Inbox.Messages.Request().Expand("attachments").GetAsync();
            var messages = request.Result;
            foreach (var message in messages)
            {
                foreach(var attachment in message.Attachments)
                {
                    if (attachment.ODataType == "#microsoft.graph.itemAttachment")
                    {

                        var attachmentRequest = graphClient.Me.MailFolders.Inbox.Messages[message.Id]
                                    .Attachments[attachment.Id].Request().Expand("microsoft.graph.itemattachment/item").GetAsync();
                        var itemAttachment = (ItemAttachment)attachmentRequest.Result;
                        var itemMessage = (Message) itemAttachment.Item;  //get attached message
                        //...
                    }
                    else
                    {
                        var fileAttachment = (FileAttachment)attachment;
                        System.IO.File.WriteAllBytes(System.IO.Path.Combine(downloadPath,fileAttachment.Name), fileAttachment.ContentBytes);
                    }
                }
            }
            

        */


    }
}