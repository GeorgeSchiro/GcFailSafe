using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using tvToolbox;

namespace GetCert2
{
    /// <summary>
    /// GcFailSafe.exe
    ///
    /// Run this program. It will prompt to create a folder with profile file:
    ///
    ///     GcFailSafe.exe.config
    ///
    /// The profile will contain help (see -Help) as well as default options.
    ///
    /// Note: This software creates its own support files (including DLLs or
    ///       other EXEs) in the folder that contains it. It will prompt you
    ///       to create its own desktop folder when you run it the first time.
    /// </summary>


    public partial class DoGcFailSafe : Form
    {
        private tvProfile       moDnsNamesProfile   = null;
        private string          msDnsName;
        private StringBuilder   msbLog = new StringBuilder();

        public DoGcFailSafe()
        {
            InitializeComponent();
        }

        private string sServerId
        {
            get
            {
                return String.Format("the \"{0}\" server", Env.sComputerName);
            }
        }


        private void DoGcFailSafe_Load(object sender, EventArgs e)
        {
            this.Hide();

            bool        lbNoPrompts         = false;
            tvProfile   loProfile           = null;
            tvProfile   loBindingsProfile   = new tvProfile();
            tvProfile   loReturnProfile     = new tvProfile();
            string      lsMsg               = null;
            string      lsMsgPromptTail     = Environment.NewLine + Environment.NewLine + "Check the log for details.";

            try
            {
                loProfile = new tvProfile(Environment.GetCommandLineArgs(), true);
                            tvProfile.oGlobal(loProfile);

                lbNoPrompts = loProfile.bValue("-NoPrompts", false);

                if ( !loProfile.bExit )
                {
                    loProfile.GetAdd("-Help",
                            @"
Introduction


This utility will get a list of all local port bound digital certificates. 
It will output the IP addresses or hostnames bound to each port as well as 
the certificate thumbprint and expiration status of each.

If a list of domain names is provided (see -DnsName below), it will be 
compared to the list of bound certificates. If a given -DnsName does not
match any of the port bound certificates, an error will be reported.

If 'stand-alone' mode is disabled (see -UseStandAloneMode below), the 
status of each bound digital certificate will be reported to the SCS 
service (see 'SafeTrust.org').
    

Command-Line Usage


Open this utility's profile file to see additional options available. It is
usually located in the same folder as '{EXE}' and has the same name
with '.config' added (see '{INI}').

Profile file options can be overridden with command-line arguments. The
keys for any '-key=value' pairs passed on the command-line must match
those that appear in the profile (with the exception of the '-ini' key).

For example, the following invokes the use of an alternative profile file
(be sure to copy an existing profile file if you do this):

    {EXE} -ini=NewProfile.txt

This tells the software to run with no UI prompts:

    {EXE} -NoPrompts


Author:  George Schiro (GeoCode@SafeTrust.org)

Date:    11/19/2020




Options and Features


The main options for this utility are listed below with their default values.
A brief description of each feature follows.

-CertificateDomainName= NO DEFAULT VALUE

    This is the subject name (ie. DNS name) of the primary certificate.

    Note: there is no need to set this value. It will be replaced by the first
            -DnsName value, if any (see -DnsName below).

-DnsName= NO DEFAULT VALUE

    This specifies a domain name for which there must be a matching bound
    digital certificate found on the server running this utility. If a matching
    certificate is not found, an error will be reported.

    The first -DnsName value will also set the -CertificateDomainName value (ie.
    the primary certificate name, see above).

    Note: This key may appear any number of times in the profile.

-FetchSource=False

    Set this switch True to fetch the source code for this utility from the EXE.
    Look in the containing folder for a ZIP file with the full project sources.

-Help= SEE PROFILE FOR DEFAULT VALUE

    This help text.

-LogEntryDateTimeFormatPrefix='yyyy-MM-dd hh:mm:ss:fff tt  '

    This format string is used to prepend a timestamp prefix to each log entry
    in the process log file (see -LogPathFile below).    

-LogFileDateFormat='-yyyy-MM-dd'

    This format string is used to form the variable part of each log file output 
    filename (see -LogPathFile below). It is inserted between the filename and 
    the extension.

-LogPathFile='Logs\Log.txt'

    This is the output path\file that will contain the process log. The profile
    filename will be prepended to the default filename (see above) and the current
    date (see -LogFileDateFormat above) will be inserted between the filename and
    the extension.

-NoPrompts=False

    Set this switch True and all pop-up prompts will be suppressed. Messages
    are written to the log instead (see -LogPathFile above). You must use this
    switch whenever the software is run via a server computer batch job or job
    scheduler (ie. where no user interaction is permitted).

-SaveProfile=True

    Set this switch False to prevent saving to the profile file by this utility
    software itself. This is not recommended since process status information is 
    written to the profile after each run.

-SaveSansCmdLine=True

    Set this switch False to allow merged command-lines to be written to
    the profile file (ie. ""{INI}""). When True, everything
    but command-line keys will be saved.

-ShowProfile=False

    Set this switch True to immediately display the entire contents of the profile
    file at startup in command-line format. This may be helpful as a diagnostic.

-UseStandAloneMode=True

    Set this switch False and the software will use the SafeTrust Secure Certificate
    Service (see 'SafeTrust.org') to communicate certificate status to various
    responsible parties.


Notes:

    There may be various other settings that can be adjusted also (user
    interface settings, etc). See the profile file ('{INI}')
    for all available options.

    To see the options related to any particular behavior, you must run that
    part of the software first. Configuration options are added 'on the fly'
    (in order of execution) to '{INI}' as the software runs.

    "
                            .Replace("{EXE}", Path.GetFileName(System.Windows.Application.ResourceAssembly.Location))
                            .Replace("{INI}", Path.GetFileName(loProfile.sActualPathFile))
                            .Replace("{{", "{")
                            .Replace("}}", "}")
                            );

                    string lsFetchName = null;

                    // Fetch license.
                    tvFetchResource.ToDisk("GetCert2", "MIT License.txt", null);

                    // Fetch simple setup.
                    tvFetchResource.ToDisk("GetCert2", String.Format("{0}{1}", Env.sFetchPrefix
                            , lsFetchName="Setup Application Folder.exe"), loProfile.sRelativeToProfilePathFile(lsFetchName));

                    // Fetch source code.
                    if ( loProfile.bValue("-FetchSource", false) )
                         tvFetchResource.ToDisk("GetCert2", System.Windows.Application.ResourceAssembly.GetName().Name + ".zip", null);

                    // Only do the following fetch on the first pass and if we're not running in stand-alone mode.
                    if ( !loProfile.ContainsKey("-UseStandAloneMode") && !loProfile.bValue("-UseStandAloneMode", true) )
                    {
                        // Discard the default profile. Replace it with the WCF version fetched below.
                        File.Delete(loProfile.sLoadedPathFile);

                        // Fetch WCF config.
                        tvFetchResource.ToDisk("GetCert2", String.Format("{0}{1}", Env.sFetchPrefix
                                , Path.GetFileName(System.Windows.Application.ResourceAssembly.Location) + ".config"), loProfile.sLoadedPathFile);

                        Env.ResetConfigMechanism(loProfile);
                    }

                    // Wait a random period each cycle to allow this client to run at times
                    // other than the standard fail-safe time (skip if in stand-alone mode).
                    int     liMaxRunTimeDelayMins = loProfile.iValue("-MaxRunTimeDelayMins", 60);
                            if ( 0 != liMaxRunTimeDelayMins && !loProfile.bValue("-UseStandAloneMode", true)
                                    && DateTime.Now.Hour < loProfile.iValue("-MaxHourToAllowRunTimeDelays", 6)
                                    )
                            {
                                int liRunTimeDelayMins = new Random().Next(liMaxRunTimeDelayMins);

                                Env.LogIt("");
                                Env.LogIt(String.Format("Waiting {0} minutes ({1} minutes max) before the fail-safe run ..."
                                                        , liRunTimeDelayMins, liMaxRunTimeDelayMins));

                                Thread.Sleep(60000 * liRunTimeDelayMins);
                            }

                    Env.LogIt("");
                    Env.LogIt("Fetching local certificate bindings ...");

                    bool            lbSkipLogPsOutput = loProfile.bValue("-SkipLogPsOutput", true);
                    bool            lbUseBindingsListOverride = loProfile.bValue("-UseBindingsListOverride", false);
                    string          lsBindingsListOverride = loProfile.sValue("-BindingsListOverride", Environment.NewLine + "No bindings have been defined." + Environment.NewLine);
                    string          lsKey = null;
                    string          lsNL = "'\r\n-";
                    string          lsBindings = null;
                                    if ( lbUseBindingsListOverride )
                                        lsBindings = lsBindingsListOverride.Replace(Environment.NewLine, "");
                                    else
                                        Env.bRunPowerScript(out lsBindings, null, loProfile.sValue("-ShowBindingsCommand", "netsh http show sslcert"), false, lbSkipLogPsOutput, true);
                    StringBuilder   lsbBindings = new StringBuilder(lsBindings);
                                    lsbBindings.Replace(lsKey="IP:port"                                                 , lsNL + " Binding=[\r\n-" + lsKey);
                                    lsbBindings.Replace(lsKey="Hostname:port"                                           , lsNL + " Binding=[\r\n-" + lsKey);
                                    lsbBindings.Replace(lsKey="Certificate Hash"                                        , lsNL + lsKey);
                                    lsbBindings.Replace(lsKey="Application ID"                                          , lsNL + lsKey);
                                    lsbBindings.Replace(lsKey="Certificate Store Name"                                  , lsNL + lsKey);
                                    lsbBindings.Replace(lsKey="Verify Client Certificate Revocation"                    , lsNL + lsKey);
                                    lsbBindings.Replace(lsKey="Verify Revocation Using Cached Client Certificate Only"  , lsNL + lsKey);
                                    lsbBindings.Replace(lsKey="Usage Check"                                             , lsNL + lsKey);
                                    lsbBindings.Replace(lsKey="Revocation Freshness Time"                               , lsNL + lsKey);
                                    lsbBindings.Replace(lsKey="URL Retrieval Timeout"                                   , lsNL + lsKey);
                                    lsbBindings.Replace(lsKey="Ctl Identifier"                                          , lsNL + lsKey);
                                    lsbBindings.Replace(lsKey="Ctl Store Name"                                          , lsNL + lsKey);
                                    lsbBindings.Replace(lsKey="DS Mapper Usage"                                         , lsNL + lsKey);
                                    lsbBindings.Replace(lsKey="Negotiate Client Certificate : Enabled"                  , lsNL + lsKey + lsNL + "Binding=]\r\n");
                                    lsbBindings.Replace(lsKey="Negotiate Client Certificate : Disabled"                 , lsNL + lsKey + lsNL + "Binding=]\r\n");
                                    lsbBindings.Replace(lsKey="Negotiate Client Certificate    : Enabled"               , lsNL + lsKey + lsNL + "Binding=]\r\n");   // 2008
                                    lsbBindings.Replace(lsKey="Negotiate Client Certificate    : Disabled"              , lsNL + lsKey + lsNL + "Binding=]\r\n");   // 2008
                                    lsbBindings.Replace("-------------------------", "");
                                    lsbBindings.Replace(" : ", "='");
                                    lsbBindings.Replace(" ", "");
                    tvProfile       loBindingsProfileRaw = new tvProfile(lsbBindings.ToString());

                                    if ( lbSkipLogPsOutput )
                                    {
                                        Env.LogSuccess();
                                    }
                                    else
                                    {
                                        Env.LogIt("");
                                        Env.LogIt("lsBindings:");
                                        Env.LogIt(lsBindings);
                                        Env.LogIt("");
                                        Env.LogIt("lsbBindings:");
                                        Env.LogIt(lsbBindings.ToString());
                                        Env.LogIt("");
                                        Env.LogIt("loBindingsProfileRaw:");
                                        Env.LogIt(loBindingsProfileRaw.ToString());
                                    }

                    Env.LogIt("");
                    Env.LogIt("Segregating IP from hostname bindings ...");

                    //Add bindings (with "-Hostname" or "IP" extracted) to the bindings profile from the raw bindings profile.
                    foreach(DictionaryEntry loEntry in loBindingsProfileRaw)
                    {
                        tvProfile   loBindingProfileRaw = new tvProfile(loEntry.Value.ToString());
                        string[]    lsSplitPortArray = loBindingProfileRaw.sValue("-Hostname:port", "").Split(':');
                        string      lsHostname = lsSplitPortArray[0];
                        tvProfile   loBindingProfile = new tvProfile();
                                    if ( !String.IsNullOrEmpty(lsHostname) )
                                    {
                                        loBindingProfile.LoadFromCommandLine(String.Format("-Hostname='{0}' -port={1}", lsHostname, lsSplitPortArray[1]), tvProfileLoadActions.Append);
                                    }
                                    else
                                    {
                                        string  lsIpAddress = null;
                                                lsSplitPortArray = loBindingProfileRaw.sValue("-IP:port", "").Split(':');

                                                // This approach handles IPv6 as well as IPv4.
                                                for (int i=0; i < lsSplitPortArray.Length - 1; i++)
                                                    lsIpAddress += (null == lsIpAddress ? "" : ":") + lsSplitPortArray[i];

                                        loBindingProfile.LoadFromCommandLine(String.Format("-IP={0} -port={1}", lsIpAddress, lsSplitPortArray[lsSplitPortArray.Length - 1]), tvProfileLoadActions.Append);
                                    }
                                    loBindingProfile.LoadFromCommandLine(loBindingProfileRaw.ToString(), tvProfileLoadActions.Append);

                        loBindingsProfile.Add("-Binding", loBindingProfile.ToString());

                        Env.LogIt("Binding:" + loBindingProfile.sCommandLine());

                        if ( !lbNoPrompts )
                        {
                            lblMessage.Text = String.Format("\"{0}\" binding found ...", !String.IsNullOrEmpty(lsHostname) ? lsHostname : loBindingProfile.sValue("-IP", "-IP not found") );
                            this.ShowMe();
                        }
                    }

                    Env.LogSuccess();

                    Env.LogIt("");
                    Env.LogIt("Checking for bound certificate expiration ...");

                    X509Store loStore = null;

                    try
                    {
                        loStore = new X509Store(StoreName.My, StoreLocation.LocalMachine);
                        loStore.Open(OpenFlags.ReadOnly);

                        for (int i=0; i < loBindingsProfile.Count; i++)
                        {
                            DictionaryEntry             loEntry = loBindingsProfile.oEntry(i);
                            bool                        lbFailure = false;
                            bool                        lbInnocuous = false;
                            bool                        lbSuccess = false;
                            tvProfile                   loBindingProfile = new tvProfile(loEntry.Value.ToString());
                                                        msDnsName = loBindingProfile.sValue("-Hostname", "");
                            bool                        lbHostnameIsIpAddress = String.IsNullOrEmpty(msDnsName);
                                                        if ( "localhost" == msDnsName.ToLower() )
                                                            lbHostnameIsIpAddress = true;
                                                        else
                                                            msDnsName = !lbHostnameIsIpAddress ? msDnsName : loBindingProfile.sValue("-IP", "Neither -Hostname nor -IP was found");
                            string                      lsPort = loBindingProfile.sValue("-port", "-port not found");
                            string                      lsCertThumbprint = loBindingProfile.sValue("-CertificateHash", "-CertificateHash (ie. thumbprint) not found");
                            tvProfile                   loResultProfile = this.oResultProfile(msDnsName, lsCertThumbprint, loBindingProfile.iValue("-port", 0), lbHostnameIsIpAddress);
                            X509Certificate2Collection  loCertCollection = loStore.Certificates.Find(X509FindType.FindByThumbprint, lsCertThumbprint, false);

                            if ( null == loCertCollection || 0 == loCertCollection.Count )
                            {
                                lbInnocuous = true;
                                lsMsg = String.Format("The \"{0}:{1}\" thumbprint is \"{2}\", yet it can't be found in the local machine personal certificate store.", msDnsName, lsPort, lsCertThumbprint);
                            }
                            else
                            {
                                X509Certificate2    loCert = loCertCollection[0];
                                string              lsCertName = null;
                                                    if ( lbHostnameIsIpAddress )
                                                    {
                                                        lsCertName = Env.sCertName(loCert);
                                                        loBindingProfile["-HostnameIsIpAddress"] = lbHostnameIsIpAddress;
                                                        loBindingProfile["-CertName"] = lsCertName;
                                                        loBindingsProfile[i] = loBindingProfile.ToString();
                                                    }
                                string              lsMsgPrefix = String.Format("The \"{0}:{1}\" thumbprint is \"{2}\"{3}. ", msDnsName, lsPort, lsCertThumbprint
                                                            , String.IsNullOrEmpty(lsCertName) ? "" : String.Format("(\"{0}\")", lsCertName));

                                if ( !loCert.Verify() )
                                {
                                    // Ignore self-signed "CN=localhost" certificates.
                                    if ( "CN=localhost" != loCert.Issuer )
                                        lbFailure = true;

                                    lsMsg = lsMsgPrefix + String.Format("The issuer is \"{0}\". This certificate is NOT VALID.", loCert.Issuer);
                                }
                                else
                                if ( loCert.NotBefore > DateTime.Now )
                                {
                                    lbFailure = true;
                                    lsMsg = lsMsgPrefix + String.Format("This certificate is NOT VALID. It will be valid after \"{0}\".", loCert.NotBefore);
                                }
                                else
                                if ( loCert.NotAfter < DateTime.Now )
                                {
                                    lbFailure = true;
                                    lsMsg = lsMsgPrefix + String.Format("This certificate is NOT VALID. It expired \"{0}\".", loCert.NotAfter);
                                }
                                else
                                if ( loCert.NotAfter < DateTime.Now.AddDays(loProfile.iValue("-MaxDaysBeforeExpiration", 14)) )
                                {
                                    int liExpirationDays = (int)loCert.NotAfter.Subtract(DateTime.Now).TotalDays;

                                    lbFailure = true;
                                    lsMsg = lsMsgPrefix + String.Format("This certificate will expire in {0} day{1}. It expires \"{2}\"."
                                                                        , liExpirationDays, 1==liExpirationDays ? "" : "s", loCert.NotAfter);
                                }
                                else
                                {
                                    lbSuccess = true;
                                    lsMsg = lsMsgPrefix + String.Format("The certificate is valid and will not expire for at least {0} days (ie. \"{1}\")."
                                                                        , loProfile.iValue("-MaxDaysBeforeExpiration", 14), loCert.NotAfter);
                                }
                            }

                            if ( lbFailure)
                                loReturnProfile.Add("-Failure", this.oResultProfile(loResultProfile, lbSuccess, lsMsg).ToString());
                            else
                            if ( lbInnocuous)
                                loReturnProfile.Add("-Warning", this.oResultProfile(loResultProfile, lbSuccess, lsMsg).ToString());
                            else
                            if ( lbSuccess)
                                loReturnProfile.Add("-Success", this.oResultProfile(loResultProfile, lbSuccess, lsMsg).ToString());

                            Env.LogIt(lsMsg);
                        }

                        Env.LogDone();
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        if ( null != loStore )
                            loStore.Close();
                    }

                    Env.LogIt("");
                    Env.LogIt("Verifying each given -DnsName has a corresponding binding ...");

                    bool lbAtLeastOneError = false;

                    moDnsNamesProfile = loProfile.oOneKeyProfile("-DnsName");

                    foreach(DictionaryEntry loEntry in moDnsNamesProfile)
                    {
                        bool    lbBindingExists = false;

                        msDnsName = loEntry.Value.ToString();

                        foreach(DictionaryEntry loEntry2 in loBindingsProfile)
                        {
                            tvProfile   loBindingProfile = new tvProfile(loEntry2.Value.ToString());
                                        lbBindingExists = msDnsName == loBindingProfile.sValue("-Hostname", "");
                                        if ( !lbBindingExists && loBindingProfile.bValue("-HostnameIsIpAddress", false) )
                                            lbBindingExists = msDnsName == loBindingProfile.sValue("-CertName", "");

                            if ( lbBindingExists )
                            {
                                lsMsg = String.Format("\"{0}\" binding match found:", msDnsName) + loBindingProfile.sCommandLine();

                                if ( !lbNoPrompts )
                                {
                                    lblMessage.Text = lsMsg;
                                    this.ShowMe();
                                }

                                break;
                            }
                        }

                        if ( !lbBindingExists )
                        {
                            lsMsg = String.Format("\"{0}\" can't be found in the list of port bindings on {1}.", msDnsName, this.sServerId);

                            loReturnProfile.Add("-Failure", this.oResultProfile(msDnsName, false, lsMsg).ToString());

                            if ( !lbNoPrompts )
                                MessageBox.Show(lsMsg + lsMsgPromptTail, System.Windows.Application.ResourceAssembly.GetName().Name);

                            lbAtLeastOneError = true;
                        }

                        Env.LogIt(lsMsg);
                    }

                    Env.LogDone();

                    if ( !lbNoPrompts && !lbAtLeastOneError )
                        MessageBox.Show("Process complete." + lsMsgPromptTail, System.Windows.Application.ResourceAssembly.GetName().Name);

                    if ( lbAtLeastOneError )
                        Environment.ExitCode = 1;
                }
            }
            catch (Exception ex)
            {
                lsMsg = String.Format("Error for DnsName=\"{0}\" occured on {1}: {2}.", msDnsName, this.sServerId, Env.sExceptionMessage(ex));

                loReturnProfile.Add("-Failure", this.oResultProfile(msDnsName, false, lsMsg).ToString());

                Env.LogIt(lsMsg);

                if ( !lbNoPrompts )
                    MessageBox.Show(lsMsg + lsMsgPromptTail, System.Windows.Application.ResourceAssembly.GetName().Name);

                Environment.ExitCode = 1;
            }
            finally
            {
                if ( !loProfile.bExit )
                {
                    bool lbSuccess = false;

                    try
                    {
                        if ( !loProfile.bValue("-UseStandAloneMode", true) )
                        {
                            // If no bindings could be found, return the log instead.
                            if ( 0 == loBindingsProfile.Count )
                                loReturnProfile["-Log"] = File.ReadAllText(Env.sLogPathFile);

                            // Add host profile (filtered).
                            loReturnProfile["-HostProfile"] = Env.sHostProfile();

                            loProfile["-CertificateDomainName"] = moDnsNamesProfile.sValue("-DnsName", loProfile.sValue("-CertificateDomainName", ""));
                            loProfile["-ContactEmailAddress"] = "not needed here, but can't be empty";
                            loProfile.Save();

                            Env.LogIt("");
                            Env.LogIt("Sending results to service ...");

                            using (GetCertService.IGetCertServiceChannel loGetCertServiceClient = Env.oGetCertServiceFactory.CreateChannel())
                            {
                                tvProfile   loMinProfile = Env.oMinProfile(loProfile);
                                byte[]      lbtArrayMinProfile = loMinProfile.btArrayZipped();
                                string      lsHash = HashClass.sHashIt(loMinProfile);

                                loGetCertServiceClient.NotifyFailSafeInternalValidationResults(lsHash, lbtArrayMinProfile, loReturnProfile.btArrayZipped());
                                if ( CommunicationState.Faulted == loGetCertServiceClient.State )
                                    loGetCertServiceClient.Abort();
                                else
                                    loGetCertServiceClient.Close();
                            }

                            Env.LogSuccess();

                            lbSuccess = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Env.LogIt(Env.sExceptionMessage(ex));
                        Environment.ExitCode = 1;
                    }

                    if ( !loProfile.bValue("-UseStandAloneMode", true) )
                    {
                        tvProfile   loGcProfile = new tvProfile(loProfile.sRelativeToProfilePathFile("GetCert2.exe.config"), true);
                                    loGcProfile.Remove("-FailSafe*");
                                    loGcProfile.Add("-FailSafe" + (lbSuccess ? "Success" : "Failure"), DateTime.Now.ToString());
                                    loGcProfile.Save();
                    }
                }

                Application.Exit();
            }
        }

        private tvProfile oResultProfile(
                  string asDnsName
                , string asCertThumbprint
                , int    aiPort
                , bool   abHostnameIsIpAddress
                , bool   abIsSuccess
                , string asMessage
                )
        {
            tvProfile   loResultProfile = new tvProfile();
                        loResultProfile["-DnsName"] = asDnsName;
                        loResultProfile["-CertThumbprint"] = asCertThumbprint;
                        loResultProfile["-Port"] = aiPort;
                        loResultProfile["-HostnameIsIpAddress"] = abHostnameIsIpAddress;
                        loResultProfile["-IsSuccess"] = abIsSuccess;
                        loResultProfile["-Message"] = asMessage;

            return loResultProfile;
        }
        private tvProfile oResultProfile(
                  string asDnsName
                , bool   abIsSuccess
                , string asMessage
                )
        {
            return this.oResultProfile(asDnsName, "", 0, false, abIsSuccess, asMessage);
        }
        private tvProfile oResultProfile(
                  string asDnsName
                , string asCertThumbprint
                , int    aiPort
                , bool   abHostnameIsIpAddress
                )
        {
            return this.oResultProfile(asDnsName, asCertThumbprint, aiPort, abHostnameIsIpAddress, false, "");
        }
        private tvProfile oResultProfile(
                  tvProfile aoResultProfile
                , bool abIsSuccess
                , string asMessage
                )
        {
            aoResultProfile["-IsSuccess"] = abIsSuccess;
            aoResultProfile["-Message"] = asMessage;

            return aoResultProfile;
        }

        private void ShowMe()
        {
            this.Activate();
            this.Opacity = 100;
            this.WindowState = FormWindowState.Normal;
            this.Show();
            this.Refresh();
        }
    }
}
