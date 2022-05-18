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
using System.Threading.Tasks;

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

      // request 1: create user
      // var resultNewUser = CreateUserAsync(client);
      // resultNewUser.Wait();
      // Console.WriteLine("New user ID: " + resultNewUser.Id);

      // request 2: update user
      // (1/2) get the user we just created
      var userToUpdate = client.Users.Request()
                                     .Select("id")
                                     .Filter("UserPrincipalName eq 'melissad@M365x23090844.onmicrosoft.com'")
                                     .GetAsync()
                                     .Result[0];
      // (2/2) update the user's phone number
      var resultUpdatedUser = UpdateUserAsync(client, userToUpdate.Id);
      resultUpdatedUser.Wait();
      Console.WriteLine("Updated user ID: " + resultUpdatedUser.Id);

      // request 3: delete user
      var deleteTask = DeleteUserAsync(client, userToUpdate.Id);
      deleteTask.Wait();
    }

    private static async Task DeleteUserAsync(GraphServiceClient client, string userIdToDelete)
    {
      await client.Users[userIdToDelete].Request().DeleteAsync();
    }

    private static async Task<Microsoft.Graph.User> UpdateUserAsync(GraphServiceClient client, string userIdToUpdate)
    {
      Microsoft.Graph.User user = new Microsoft.Graph.User()
      {
        MobilePhone = "555-555-1212"
      };
      return await client.Users[userIdToUpdate].Request().UpdateAsync(user);
    }

    private static async Task<Microsoft.Graph.User> CreateUserAsync(GraphServiceClient client)
    {
      Microsoft.Graph.User user = new Microsoft.Graph.User()
      {
        AccountEnabled = true,
        GivenName = "Melissa",
        Surname = "Darrow",
        DisplayName = "Melissa Darrow",
        MailNickname = "MelissaD",
        UserPrincipalName = "melissad@M365x23090844.onmicrosoft.com",
        PasswordProfile = new PasswordProfile()
        {
          Password = "Password1!",
          ForceChangePasswordNextSignIn = true
        }
      };
      var requestNewUser = client.Users.Request();
      return await requestNewUser.AddAsync(user);
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