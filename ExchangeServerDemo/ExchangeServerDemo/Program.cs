using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeServerDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
            service.Credentials = new NetworkCredential("Domaindeki kullanıcı adı@siz", "Şifre", "DomainAdı");
            service.AutodiscoverUrl("mailadresiniz@teknohisar.com");//Mail Adresiniz @şeklinde
            service.TraceEnabled = true;//true yaptığımızda ekranda süreç ile ilgili hata alırsanız onun hakkında bilgi alırsınız.

            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, new ItemView(10)); //inboxdaki 10 adet maili getirir
            Folder rootfolder = Folder.Bind(service, WellKnownFolderName.MsgFolderRoot); //Mail Klasörlerini getirir
            rootfolder.Load();

            foreach (Item item in findResults.Items)//Mailler döner
            {
                foreach (Folder folder in rootfolder.FindFolders(new FolderView(100)))//Mail klasörü
                {
                    if (folder.DisplayName == "TaşınacakKlasörAdı")//Klasör adını buluyoruz
                    {
                        var fid = folder.Id;//folder id buluyoruz
                        item.Move(fid);//maili o folder'a taşıyoruz
                    }
                }

            }
        }
    }
}
