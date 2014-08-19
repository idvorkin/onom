using System;
using System.Xml.Linq;
using Microsoft.Office.Interop.Outlook;

namespace OnenoteCapabilities
{
    /// <summary>
    /// Convert connect tags to mail:// tags
    /// </summary>
    public class ConnectSmartTagProcessor : ISmartTagProcessor
    {
        public bool ShouldProcess(SmartTag st, OneNotePageCursor cursor)
        {
            return st.TagName() == "connect";
        }

        private string ResolveNameViaMapi(string userName)
        {
            // TBD: Learn how to get email suffix via mapi.
            var emailSuffix = "@microsoft.com";

            var outlook = new Application();
            var mapi = outlook.GetNamespace("MAPI");
            var recipient = mapi.CreateRecipient(userName);


            bool isSimpleResolution = recipient.Resolve();
            if (isSimpleResolution)
            {
                return recipient.AddressEntry.GetExchangeUser().Alias + emailSuffix;
            }

            // show resolution dialog.
            var resolverDialog = outlook.Session.GetSelectNamesDialog();
            resolverDialog.Recipients.Add(userName);
            if (resolverDialog.Recipients.ResolveAll())
            {
                return resolverDialog.Recipients[1].AddressEntry.GetExchangeUser().Alias + emailSuffix;
            }
            
            // all resolutions failed - return null.
            return null;
        }

        public void Process(SmartTag smartTag, XDocument pageContent, SmartTagAugmenter smartTagAugmenter, OneNotePageCursor cursor)
        {
            // a weak test.
            var isAnEmailAddress = (smartTag.TextAfterTag().Contains("@"));
            string emailAddress = "";
            if (isAnEmailAddress)
            {
                emailAddress = smartTag.TextAfterTag();
            }
            else
            {
                // let mapi mapi do resolution.
                emailAddress = ResolveNameViaMapi(smartTag.TextAfterTag());
            }

            if (String.IsNullOrEmpty(emailAddress))
            {
                smartTag.AddContentAfter("<b>Error:</b> unable to resolve name");
                return;
            }

            var mailLink = String.Format("mailto:{0}", emailAddress);
            smartTag.SetLink(new Uri(mailLink));
        }

        public string HelpLine()
        {
            return "<b>#connect</b> make a connection to the person or email address";
        }
    }
}
