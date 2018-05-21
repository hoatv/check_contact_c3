using System;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Util.Store;
using Google.Apis.Sheets.v4;

namespace hoatv.Controllers
{
	public class AppFlowMetadata 
	{
		private static readonly IAuthorizationCodeFlow flow =
			new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
			{
				ClientSecrets = new ClientSecrets
				{
					ClientId = "621141550197-si5oo0nve6ebuqvdbrehbmkstkbpopmk.apps.googleusercontent.com",
					ClientSecret = "r3vPYqQFW-IEMVC4yqhr1Ibf"
				},
				Scopes = new[] { SheetsService.Scope.Spreadsheets },
				DataStore = new FileDataStore("Drive.Api.Auth.Store")
			});

		//public override IAuthorizationCodeFlow Flow
		//{
		//	get { return flow; }
		//}
	}
}
