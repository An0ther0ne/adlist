// We can get users from AD with one line command in PowerShell, but this command-line solution run more quickly and without any human intervention.
// Just use old school syntax (vs2010)
using System;
using System.IO;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Security.Principal;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ADUsers{

    public class adduser{
        public string Name;
        public string displayName;
        public string description;
        public string mail;
        public string givenName;
        public string sn;
        public string telephoneNumber;
        public string gid;
        public string sid;
        public string primaryGroupID;
        public DateTime lastLogon;
        public int logonCount;
        public int userAccountControl;
        public bool AccountDisabled;
        public string userPrincipalName;
        public string distinguishedName;

        public void init(){
            Name = string.Empty;
            displayName = string.Empty;
            description = string.Empty;
            mail = string.Empty;
            givenName = string.Empty;
            sn = string.Empty;
            telephoneNumber = string.Empty;
            gid = string.Empty;
            sid = string.Empty;
            primaryGroupID = string.Empty;
            lastLogon = DateTime.FromFileTimeUtc(0);
            logonCount = 0;
            userAccountControl = 0;
            AccountDisabled = false;
            userPrincipalName = string.Empty;
        }
        internal static void Usage(){
            System.Console.WriteLine("Usage: adusers.exe [options]\n"+
            "\tWhere options as follow:\n"+
            "\t -help\tThis screen.\n"+
            "\t -CN:\tCommon Name. Optional parameter and may be skipped.\n"+
            "\t -OU:\tOrganisation Units, separated by a point.\n"+
            "\t -DC:\tFull Domain Controller Name, for example: DC:dc1.example.com\n"+
            "\t -out:<out.csv> -- Optional: specify output file name. 'ADUsers.csv' by default.\n"+
            "\t -usr:<mask> -- An username mask to search for. Default: '*'. Example: '-usr:50*'\n");
        }
    }

    class Program{
        static int Main(string[] args){
            string CN="",OU="",DC="";
            string entry = "";
            string username = "*";
            string fname = "ADUsers.csv";

            foreach (string sss in args){ // parse command line
                string key, val;
                String pattern = @"-{1,2}([^:^\s]+)\:(.+)";
                System.Text.RegularExpressions.Match mmm = System.Text.RegularExpressions.Regex.Match(sss, pattern);
                if (mmm.Success){
                    key = mmm.Groups[1].Value;
                    val = mmm.Groups[2].Value;
                }else{
                    key = sss;
                    val = "";
                }
                switch (key.ToLower()){
                    case "--help":
                    case "-help":
                    case "-?":
                    case "/?":
                    case "?":
                        adduser.Usage();
                        return 0;
                    case "cn":
                        CN = String.Concat("CN=",val,",");
                        break;
                    case "ou":
                        foreach (string ss in val.Split('.')){
                            OU = String.Concat(OU, "OU=", ss, ",");
                        }
                        break;
                    case "dc":
                        foreach (string ss in val.Split('.')){
                            DC = String.Concat(DC, "DC=", ss, ",");
                        }
                        DC = DC.TrimEnd(',');
                        System.Console.WriteLine(DC);
                        break;
                    case "out":
                        fname = String.Copy(val);
                        break;
                    case "usr":
                        username = String.Copy(val);
                        break;
                }
            }
            string destr = String.Concat("LDAP://", CN, OU, DC);

            DirectoryEntry de = new DirectoryEntry(destr); //"LDAP://CN=cname,OU=subOU,OU=nameOU,DC=some,DC=example,DC=com,DC=ua");
            StreamWriter fs = null;
            Domain dd = Domain.GetCurrentDomain();
            SearchResultCollection rescol;
            SearchResult res;
            adduser iUser = new adduser();
//          DirectoryEntry de = new DirectoryEntry(string.Format("LDAP://{0}", dd));
            DirectorySearcher search = new DirectorySearcher(entry);
            search.SearchRoot = de;
//          search.Filter = "(objectCategory=user)";   // otherwise groups join to result
            search.Filter = string.Format("(&(SAMAccountName={0})(objectCategory=user))", username);
            search.PropertiesToLoad.Add("name");
            search.PropertiesToLoad.Add("displayname");
            search.PropertiesToLoad.Add("description");
            search.PropertiesToLoad.Add("mail");
            search.PropertiesToLoad.Add("givenname");
            search.PropertiesToLoad.Add("sn");
            search.PropertiesToLoad.Add("telephonenumber");
            search.PropertiesToLoad.Add("objectguid");
            search.PropertiesToLoad.Add("objectsid");
            search.PropertiesToLoad.Add("primarygroupid");
//          search.PropertiesToLoad.Add("company");
//          search.PropertiesToLoad.Add("homePhone");
            search.PropertiesToLoad.Add("lastlogon");
            search.PropertiesToLoad.Add("logoncount");
            search.PropertiesToLoad.Add("useraccountcontrol");
//          search.PropertiesToLoad.Add("msds-useraccountdisabled");
            search.PropertiesToLoad.Add("userprincipalname");
            search.PropertiesToLoad.Add("distinguishedName");
//          search.PropertiesToLoad.Add("st");
//          search.PropertiesToLoad.Add("postalCode");
            rescol = search.FindAll();
            if (rescol != null){
                try{
//                  fs = File.CreateText(fname);
                    fs = new StreamWriter(fname, false, Encoding.Default);
                }
                catch (Exception e){
                    System.Console.WriteLine("Error while create file: " + e.ToString());
                    Environment.Exit(-1);
                }
                fs.WriteLine("Name;DisplayName;Description;EMail;GivenName;SN;TelephoneNumber;GID;SID;PrimaryGroupID;LastLogon;LogonCount;UserAccountControl;AccountDisabled;UserPrincipalName;distinguishedName");
                foreach (SearchResult ires in rescol){
                    iUser.init();
                    foreach (string key in ires.Properties.PropertyNames){
                        switch (key){
                            case "name":
                                iUser.Name = ires.Properties[key][0].ToString();
                                break;
                            case "displayname":
                                iUser.displayName = ires.Properties[key][0].ToString();
                                break;
                            case "description":
                                iUser.description = ires.Properties[key][0].ToString();
                                break;
                            case "mail":
                                iUser.mail = ires.Properties[key][0].ToString();
                                break;
                            case "givenname":
                                iUser.givenName = ires.Properties[key][0].ToString();
                                break;
                            case "sn":
                                iUser.sn = ires.Properties[key][0].ToString();
                                break;
                            case "telephonenumber":
                                iUser.telephoneNumber = ires.Properties[key][0].ToString();
                                break;
                            case "objectguid":
                                var gid = new Guid((byte[])@ires.Properties[key][0]);
                                iUser.gid = gid.ToString().ToUpper();
//                              System.Console.WriteLine(key + " = " + gid.ToString());
                                break;
                            case "objectsid":
                                var sid = new SecurityIdentifier((byte[])@ires.Properties[key][0], 0);
                                iUser.sid = sid.ToString();
//                              System.Console.WriteLine(key + " = " + sid.ToString());
                                break;
                            case "primarygroupid":
                                iUser.primaryGroupID = ires.Properties[key][0].ToString();
                                break;
                            case "lastlogon":
                                iUser.lastLogon = DateTime.FromFileTimeUtc((Int64)ires.Properties[key][0]);
                                break;
                            case "logoncount":
                                iUser.logonCount = (int)ires.Properties[key][0];
                                break;
                            case "useraccountcontrol":
                                int acb = (int)ires.Properties[key][0];
                                iUser.userAccountControl = acb;
                                if ((acb & 0x0002) != 0){
                                    iUser.AccountDisabled = true;
                                }else{
                                    iUser.AccountDisabled = false;
                                }
                                break;
                            case "userprincipalname":
                                iUser.userPrincipalName = ires.Properties[key][0].ToString();
                                break;
                            case "distinguishedname":
                                iUser.distinguishedName = ires.Properties[key][0].ToString();
                                break;
                        }
                    }
                    fs.Write("\"" + iUser.Name + "\";");
                    fs.Write("\"" + iUser.displayName + "\";");
                    fs.Write("\"" + iUser.description + "\";");
                    fs.Write("\"" + iUser.mail + "\";");
                    fs.Write("\"" + iUser.givenName + "\";");
                    fs.Write("\"" + iUser.sn + "\";");
                    fs.Write("\"" + iUser.telephoneNumber + "\";");
                    fs.Write("\"" + iUser.gid + "\";");
                    fs.Write("\"" + iUser.sid + "\";");
                    fs.Write("\"" + iUser.primaryGroupID + "\";");
                    fs.Write("\"" + iUser.lastLogon.ToString() + "\";");
                    fs.Write("\"" + iUser.logonCount + "\";");
                    fs.Write("\"" + iUser.userAccountControl + "\";");
                    if (iUser.AccountDisabled){
                        fs.Write("\"1\";");
                    }else{
                        fs.Write("\"0\";");
                    }
                    fs.Write("\"" + iUser.userPrincipalName + "\";");
                    fs.Write("\"" + iUser.distinguishedName + "\"");
                    fs.WriteLine();
                }
            }
            if (fs != null){
                System.Console.WriteLine("Save result to file: " + fname);
                fs.Close();
            }
        return 0;
        }
    }
}
