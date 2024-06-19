using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace TMHelper.Common
{
    public class ActiveDirectoryHelper
    {
        #region Properties
        private string RootDomainName { get; set; }
        private string DomainUserName { get; set; }
        private string DomainPassword { get; set; }
        private PrincipalContext pContext { get; set; }
        #endregion

        #region Constructor
        public ActiveDirectoryHelper(string strDomainName, string strUserName, string strPassword)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(strDomainName)) { throw new Exception("Domain name is null or empty"); }
                if (string.IsNullOrWhiteSpace(strUserName)) { throw new Exception("Domain user name is null or empty"); }
                if (string.IsNullOrWhiteSpace(strPassword)) { throw new Exception("Domain password is null or empty"); }

                this.RootDomainName = strDomainName;
                this.DomainUserName = strUserName;
                this.DomainPassword = strPassword;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Accessible Method
        public bool GetUserADStatusBySamAccName(string strSamAccName, out bool blnRecordFound, bool blnIsSearchAllSubDomain = true)
        {
            bool blnUserADStatus = false;
            UserPrincipal uPrincipal = null;

            blnRecordFound = false;

            try
            {
                uPrincipal = blnIsSearchAllSubDomain ? SearchUserFromAllSubDomain(IdentityType.SamAccountName, strSamAccName) : SearchUser(this.RootDomainName, IdentityType.SamAccountName, strSamAccName);
                if (uPrincipal != null && uPrincipal.Enabled != null && uPrincipal.Enabled.HasValue)
                {
                    blnRecordFound = true;
                    blnUserADStatus = uPrincipal.Enabled.Value;
                }
                else
                {
                    blnRecordFound = false;
                    blnUserADStatus = false;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (uPrincipal != null)
                {
                    uPrincipal.Dispose();
                    uPrincipal = null;
                }
            }

            return blnUserADStatus;
        }

        public List<string> GetSubDomainList()
        {
            List<string> listSubDomain = null;

            try
            {
                listSubDomain = new List<string>();

                string strLdapBaseDomain = string.Format("LDAP://{0}", this.RootDomainName);
                string strDirPath = string.Format("{0}/rootDSE", strLdapBaseDomain);

                DirectoryEntry deRoot = new DirectoryEntry(strDirPath, this.DomainUserName, this.DomainPassword);
                string strDomainConfigNameContext = deRoot.Properties["configurationNamingContext"][0].ToString();

                strDirPath = string.Format("{0}/{1}", strLdapBaseDomain, strDomainConfigNameContext);
                DirectoryEntry deBase = new DirectoryEntry(strDirPath, this.DomainUserName, this.DomainPassword);

                DirectorySearcher dsLookForDomain = new DirectorySearcher(deBase);
                dsLookForDomain.Filter = "(&(objectClass=crossRef)(nETBIOSName=*))";
                dsLookForDomain.SearchScope = SearchScope.Subtree;
                dsLookForDomain.PropertiesToLoad.Add("nCName");
                dsLookForDomain.PropertiesToLoad.Add("dnsRoot");

                SearchResultCollection srcDomains = dsLookForDomain.FindAll();

                if (srcDomains != null && srcDomains.Count > 0)
                {
                    foreach (SearchResult aSRDomain in srcDomains)
                    {
                        string strChildDomainName = aSRDomain.Properties["dnsRoot"][0].ToString();
                        listSubDomain.Add(strChildDomainName);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return listSubDomain;
        }

        public UserPrincipal SearchUserFromAllSubDomain(IdentityType iType, string strIdentityVal)
        {
            UserPrincipal uPrincipal = null;
            try
            {
                List<string> listSubDomain = GetSubDomainList();
                if (listSubDomain != null && listSubDomain.Count > 0)
                {
                    foreach (string strDomainNm in listSubDomain)
                    {
                        try
                        {
                            uPrincipal = SearchUser(strDomainNm, iType, strIdentityVal);
                            if (uPrincipal != null)
                                break;
                        }
                        catch (Exception ex) { }
                    }
                }
                else
                {
                    //Search root domain only if sub domain list is empty
                    uPrincipal = SearchUser(this.RootDomainName, iType, strIdentityVal); 
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return uPrincipal;
        }

        public UserPrincipal SearchUser(string strDomainName, IdentityType iType, string strIdentityVal)
        {
            PrincipalContext pCOntext = null;
            UserPrincipal uPrincipal = null;

            try
            {
                pCOntext = new PrincipalContext(ContextType.Domain, strDomainName, this.DomainUserName, this.DomainPassword);
                uPrincipal = UserPrincipal.FindByIdentity(pCOntext, iType, strIdentityVal);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (pCOntext != null)
                {
                    pCOntext.Dispose();
                    pCOntext = null;
                }
            }

            return uPrincipal;
        }
        #endregion
    }
}
