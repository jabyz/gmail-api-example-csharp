using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GmailAPI_GetMails
{
  class Program
  {
    // If modifying these scopes, delete your previously saved credentials
    // at ~/.credentials/gmail-dotnet-quickstart.json
    static string[] Scopes = {
      GmailService.Scope.GmailReadonly
    };
    static string ApplicationName = "Gmail API .NET Quickstart";
    static void Main(string[] args)
    {
      UserCredential credential;

      using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
      {
        // The file token.json stores the user's access and refresh tokens, and is created
        // automatically when the authorization flow completes for the first time.
        string credPath = "token.json";
        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
          GoogleClientSecrets.Load(stream).Secrets,
          Scopes,
          "user",
          CancellationToken.None,
          new FileDataStore(credPath, true)).Result;
        //Console.WriteLine("Credential file saved to: " + credPath);
      }

      // Create Gmail API service.
      var service = new GmailService(new BaseClientService.Initializer()
      {
        HttpClientInitializer = credential,
        ApplicationName = ApplicationName,
      });

      //GET EMAILS INSIDE THE INBOX AND NOT OTHER FOLDERS ONLY THOSE UNREAD
      var inboxRequest = service.Users.Messages.List("me");
      inboxRequest.LabelIds = "INBOX";
      inboxRequest.Q = "is:unread";
      var inboxResponse = inboxRequest.Execute();
      System.Threading.Thread.Sleep(100);

      //LOOP THROUGH EMAILS
      foreach (var email in inboxResponse.Messages)
      {
        
        //GET ACTUAL DATA FROM EMAIL
        Message currentMail = service.Users.Messages.Get("me", email.Id).Execute();
        System.Threading.Thread.Sleep(100);

        IList<MessagePart> CurrentEmailAttachments = currentMail.Payload.Parts;
        IList<MessagePartHeader> CurrentEmailHeader = currentMail.Payload.Headers;

        //LOOP THROUGH EMAIL PROPERTIES
        foreach (MessagePartHeader headers in CurrentEmailHeader)
        {

          //FIND SPECIFIC SENDER
          if (headers.Name.ToString().ToUpper() == "FROM" && headers.Value.ToString().ToUpper().Split('<', '>')[1] == "my-email@mail.com".ToUpper())
          {
            Console.WriteLine(headers.Name.ToString().ToUpper());
            Console.WriteLine(headers.Value.ToString().ToUpper());
            foreach (MessagePart CurrentAttachment in CurrentEmailAttachments)
            {

              //LOOP THROUGH ATTACHMENTS IF ANY
              if (CurrentAttachment.Filename != "" && CurrentAttachment.Filename.Length > 0)
              {
                Console.WriteLine(CurrentAttachment.Filename);
                string attachId = CurrentAttachment.Body.AttachmentId;
                MessagePartBody attachPart = service.Users.Messages.Attachments.Get("me", email.Id, attachId).Execute();
                System.Threading.Thread.Sleep(100);

                String attachData = attachPart.Data.Replace("-", "+");
                attachData = attachData.Replace("_", "/");
                byte[] bytes = Convert.FromBase64String(attachData);


                //SAVE ATTACHMENTS TO LOCAL FOLDER
                try
                {
                  var dir = @"C:\myLocalDir\";

                  File.WriteAllBytes(Path.Combine(dir, CurrentAttachment.Filename.ToString().Trim()), bytes);
                }
                catch (Exception ex)
                {
                  Console.WriteLine(ex.Message);
                  throw;
                }
              }
            }
          }

        }

      }

      Console.Read();
    }
  }
}