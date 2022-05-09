// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;

namespace graphconsoleapp
{
  public class Program
  {
    public static void Main(string[] args)
    {
      Console.WriteLine("Hello World!");

      var config = LoadAppSettings();
      if (config == null)
      {
        Console.WriteLine("Invalid appsettings.json file.");
        return;
      }

      var client = GetAuthenticatedGraphClient(config);

      // request 1 - all users
      // var requestAllUsers = client.Users.Request();

      // var results = requestAllUsers.GetAsync().Result;
      // foreach (var user in results)
      // {
      //   Console.WriteLine(user.Id + ": " + user.DisplayName + " <" + user.Mail + ">");
      // }

      // Console.WriteLine("\nGraph Request:");
      // Console.WriteLine(requestAllUsers.GetHttpRequestMessage().RequestUri);

      // request 2 - current user
      // var requestMeUser = client.Me.Request();

      // var resultMe = requestMeUser.GetAsync().Result;
      // Console.WriteLine(resultMe.Id + ": " + resultMe.DisplayName + " <" + resultMe.Mail + ">");

      // Console.WriteLine("\nGraph Request:");
      // Console.WriteLine(requestMeUser.GetHttpRequestMessage().RequestUri);

      // request 3 - specific user
      var requestSpecificUser = client.Users["c3c76dcc-2109-49d9-a789-698b914776b0"].Request();
      var resultOtherUser = requestSpecificUser.GetAsync().Result;
      Console.WriteLine(resultOtherUser.Id + ": " + resultOtherUser.DisplayName + " <" + resultOtherUser.Mail + ">");

      Console.WriteLine("\nGraph Request:");
      Console.WriteLine(requestSpecificUser.GetHttpRequestMessage().RequestUri);
    }

    private static IConfigurationRoot? LoadAppSettings()
    {
      try
      {
        var config = new ConfigurationBuilder()
                          .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                          .AddJsonFile("appsettings.json", false, true)
                          .Build();

        if (string.IsNullOrEmpty(config["applicationId"]) ||
            string.IsNullOrEmpty(config["tenantId"]))
        {
          return null;
        }

        return config;
      }
      catch (System.IO.FileNotFoundException)
      {
        return null;
      }
    }

    private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
    {
      var clientId = config["applicationId"];
      var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

      List<string> scopes = new List<string>();
      scopes.Add("https://graph.microsoft.com/.default");

      var cca = PublicClientApplicationBuilder.Create(clientId)
                                              .WithAuthority(authority)
                                              .WithDefaultRedirectUri()
                                              .Build();
      return MsalAuthenticationProvider.GetInstance(cca, scopes.ToArray());
    }

    private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
    {
      var authenticationProvider = CreateAuthorizationProvider(config);
      var graphClient = new GraphServiceClient(authenticationProvider);
      return graphClient;
    }
  }
}