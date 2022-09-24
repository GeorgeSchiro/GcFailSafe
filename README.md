Overview
========


**GcFailSafe** will get a list of all local port bound digital certificates. It will output the IP addresses or hostnames bound to each port as well as the certificate thumbprint and expiration status of each.

If a list of domain names is provided (see -DnsName below), it will be compared to the list of bound certificates. If a given -DnsName does not match any of the port bound certificates, an error will be reported.

If 'stand-alone' mode is disabled (see -UseStandAloneMode below), the status of each bound digital certificate will be reported to the SCS service (see 'SafeTrust.org').


Features
========


-   Simple setup - try it out fast
-   Requires no input.
-   Automatically finds all port bound certificates
-   Can be command-line driven from another process
-   Software is highly configurable
-   Software is totally self-contained (EXE is its own setup)


Details
=======


**GcFailSafe** is typically used together with **GetCert2** for failsafe warnings. Its output is used to generate notifications about certificates due to expire within 14 days (the default).

It's a simple tool for tracking the expiration status of all port bound digital certificates on a single Windows server.

**GcFailSafe** is typically run daily by a job scheduler.

If you are interested in a specific certificate or set of certificates, you can pass a list of certificate names like this:

    GcFailSafe.exe -DnsName=MyDomain1.com -DnsName=MyDomain2.com

Alternatively, you can add any number of -DnsName values to the **GcFailSafe** profile file (see below).


Requirements
============


-   .Net Framework 4.5+
-   PowerShell 5.1+
-   Windows Server 2008+


Command-Line Usage
==================


    Open this utility's profile file to see additional options available. It is
    usually located in the same folder as 'GcFailSafe.exe' and has the same name
    with '.config' added (see 'GcFailSafe.exe.config').

    Profile file options can be overridden with command-line arguments. The
    keys for any '-key=value' pairs passed on the command-line must match
    those that appear in the profile (with the exception of the '-ini' key).

    For example, the following invokes the use of an alternative profile file
    (be sure to copy an existing profile file if you do this):

        GcFailSafe.exe -ini=NewProfile.txt

    This tells the software to run with no UI prompts:

        GcFailSafe.exe -NoPrompts


    Author:  George Schiro (GeoCode@SafeTrust.org)

    Date:    11/19/2020

 
Options and Features
====================


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
    the profile file (ie. 'GcFailSafe.exe.txt'). When True, everything
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
    interface settings, etc). See the profile file ('GcFailSafe.exe.config')
    for all available options.

    To see the options related to any particular behavior, you must run that
    part of the software first. Configuration options are added 'on the fly'
    (in order of execution) to 'GcFailSafe.exe.config' as the software runs.
