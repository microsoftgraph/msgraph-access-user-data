using System;
using System.Collections.Generic;
using System.Security;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;
using System.Threading.Tasks;

namespace graphusers01
{
  class Program
  {
    static void Main(string[] args)
    {
      Console.WriteLine("Hello World!");

      var config = LoadAppSettings();
      if (config == null)
      {
        Console.WriteLine("Invalid appsettings.json file.");
        return;
      }

      var userName = ReadUsername();
      var userPassword = ReadPassword();

      var client = GetAuthenticatedGraphClient(config, userName, userPassword);

      // request 1: create user
      // var resultNewUser = CreateUserAsync(client);
      // resultNewUser.Wait();
      // Console.WriteLine("New user ID: " + resultNewUser.Id);

      // request 2: update user
      // (1/2) get the user we just created
      var userToUpdate = client.Users
                                      .Request()
                                      .Select("id")
                                      .Filter("UserPrincipalName eq 'melissad@M365x068225.onmicrosoft.com'")
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

    private static async Task<Microsoft.Graph.User> CreateUserAsync(GraphServiceClient client) {
      Microsoft.Graph.User user = new Microsoft.Graph.User() {
        AccountEnabled = true,
        GivenName = "Melissa",
        Surname = "Darrow",
        DisplayName = "Melissa Darrow",
        MailNickname = "MelissaD",
        UserPrincipalName = "melissad@M365x068225.onmicrosoft.com",
        PasswordProfile = new PasswordProfile() {
          Password = "Password1!",
          ForceChangePasswordNextSignIn = true
        }
      };
      var requestNewUser = client.Users.Request();
      return await requestNewUser.AddAsync(user);
    }

    private static async Task<Microsoft.Graph.User> UpdateUserAsync(GraphServiceClient client, string userIdToUpdate) {
      Microsoft.Graph.User user = new Microsoft.Graph.User() {
        MobilePhone = "555-555-1212"
      };
      return await client.Users[userIdToUpdate].Request().UpdateAsync(user);
    }

    private static async Task DeleteUserAsync(GraphServiceClient client, string userIdToDelete) {
      await client.Users[userIdToDelete].Request().DeleteAsync();
    }

    private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config, string userName, SecureString userPassword)
    {
      var clientId = config["applicationId"];
      var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

      List<string> scopes = new List<string>();
      scopes.Add("User.Read");
      scopes.Add("User.Read.All");
      scopes.Add("User.ReadWrite.All");

      var cca = PublicClientApplicationBuilder.Create(clientId)
                                              .WithAuthority(authority)
                                              .Build();
      return MsalAuthenticationProvider.GetInstance(cca, scopes.ToArray(), userName, userPassword);
    }

    private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config, string userName, SecureString userPassword)
    {
      var authenticationProvider = CreateAuthorizationProvider(config, userName, userPassword);
      var graphClient = new GraphServiceClient(authenticationProvider);
      return graphClient;
    }

    private static string ReadUsername()
    {
      string username;
      Console.WriteLine("Enter your username");
      username = Console.ReadLine();
      return username;
    }

    private static SecureString ReadPassword()
    {
      Console.WriteLine("Enter your password");
      SecureString password = new SecureString();
      while (true)
      {
        ConsoleKeyInfo c = Console.ReadKey(true);
        if (c.Key == ConsoleKey.Enter)
        {
          break;
        }
        password.AppendChar(c.KeyChar);
        Console.Write("*");
      }
      Console.WriteLine();
      return password;
    }

    private static IConfigurationRoot LoadAppSettings()
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
  }
}
