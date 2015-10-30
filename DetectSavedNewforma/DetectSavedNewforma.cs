using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Net.Mail;


namespace DetectSavedNewforma
{
	public partial class NewformaAddIn
    {		
		Outlook.Items mailItems;
		const string PR_MAIL_HEADER_TAG = @"http://schemas.microsoft.com/mapi/proptag/0x007D001E";
		public string Domain { get; set; }
		public string EmailAddress { get; set; }

		private void AddInStartUp(object sender, System.EventArgs e)
		{
			EmailAddress = this.Application.ActiveExplorer().Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress;
			MailAddress address = new MailAddress(EmailAddress);
			Domain = address.Host;

			Outlook.MAPIFolder inbox = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
			mailItems = inbox.Items;

			mailItems.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(ThreadStarter);
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(AddInStartUp);
		}
		#endregion

		private void ThreadStarter(Object Item)
		{
			System.Threading.Thread incomingMailThread = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(this.InboxFolderItemAdded));
			incomingMailThread.IsBackground = true;
			incomingMailThread.Start(Item);
		}

		private void InboxFolderItemAdded(object Item)
		{
			ProcessMailItem(Item);
		}

		private void ProcessMailItem(object Item)
		{
			if (Item is Outlook.MailItem)
			{
				Outlook.MailItem mailItem = (Item as Outlook.MailItem);
				Outlook.PropertyAccessor propertyAccessor = mailItem.PropertyAccessor as Outlook.PropertyAccessor;

				if (mailItem.SenderEmailAddress.Contains(Domain))
				{
					try
					{
						string headers = (string)propertyAccessor.GetProperty(PR_MAIL_HEADER_TAG);

						if (headers.Contains("x-newforma-client-submit-time"))
						{
							string existingCategories = mailItem.Categories;
							if (String.IsNullOrEmpty(existingCategories))
							{
								mailItem.Categories = "Filed By Newforma";
							}
							else
							{
								if (mailItem.Categories.Contains("Filed By Newforma") == false)
								{
									mailItem.Categories = existingCategories + ", Filed By Newforma";
								}
							}
							mailItem.Save();
						}
					}
					catch { }
				}
			}
		}
	}
}
