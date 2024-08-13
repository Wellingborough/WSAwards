Project for the automatic generation of Award Postcards
Wellingborough School


Requirements
============

The headline requirement is to generate emails with 'postcards' congratulating students on achieving five positive
points (rewards) in the current reporting period.

The Rewards and Sanctions report spreadsheet contains a workbook per department

We really should have a database with the pupils running totals in.  

Pupils may only appear in the Rewards and Sanctions report spreadsheet when they are awarded some points, but these may
be negative points of course, so they may show up but have 0[0] positives.

We can generate the initial database of students and their names and scores from the first export of the Pupil Contacts
spreadsheet; we can then scan the pupil spreadsheet on each run to check for new students.

We will have to regenerate the pupil database at the start of each year.


Status
======
Done:

Around 75% complete in the first few hours...
Database builds automatically from contacts spreadsheet (and some checking for updates done)
Rewards spreadsheet is parsed and processed with department name exceptions and department head lookup

Remaining:

Should add the email address to the database?  Although we can rebuild this from puid+@wellingboroughschool.org
Probably best to do this for now as we can mangle the email addresses (e.g., replace them with
mark.r.gamble@outlook.com)

Gmail integration done.

Should probably add some GUI for future use

Integrate with email - check how I did this previously.  Hmmm - I am not sure that I did!  See the section at the end
of this document...

And I have been an idiot.  We need a table per department with the points for each student (currently I am overwriting
the number of points each time I find a student on a sheet.  So if we have a table with all of the pupils in (puid,
fname, sname, academicyear) and then a table per department with uuid, puid, points.  Fixed



Issues
======

1.  Excel Formats
-----------------
The spreadsheets as provided by CAG are in different formats:

Pupil Contacts.xlsx (this is fine)
Senior Department Rewards and Sanctions.xls (less so - openpyxl cannot read)
(Check whether this was a choice on export)

Second spreadsheet converted.  Admin staff will have to do this.  (If we cannot change on export)

Second spreadsheet has X[Y] in the Positive[PTS] column - we are currently taking the first number (X).  Should we be
taking the second?

2.  Department Names
--------------------

There are three departments to ignore:
LD, VMT, PSHCE

There is one department to rename:
Information Technology -> Computer Science

For the moment I will just hardcode these (done in constants.py).

3.  Department Head Names
-------------------------

Not currently in the spreadsheets.

For the moment I will just hardcode these, but for a futureproof solution, should have these in a spreadsheet which can
be exported from iSAMS.

Apparently not available. so we will have to create and maintain a spreadsheet of these.

4.  Pupil Names
---------------

At present, I have no way of identifying students other than by name.  This is going to be a problem when students
change their names (which they inevitably do in terms of surname).

I said the above, but the Pupil Contacts spreadsheet has the email address which contains the unique pupil Id.  We
should use this as a reliably unique ID, updating the pupil name as necessary.

Ah, the problem arises however, because the second spreadsheet uses only pupil name, not their surname.  So, if a
surname changes from one run to the next, then I will see a new surname and a missing surname.

5. Emailing
-----------

Will require IS to sort out permissions.

An alternative is to explore SchoolPost which integrates with Firefly and is already used to distribute reports and
other communications.

Another alternative might be to go via Firefly itself (SchoolPost is a Firefly company).  If we could use an API to
'publish' the awards pages to Firefly, could we get them emailed from there?

A further problem is how to serve up the icon images - this will be a problem if we cannot embed the images.  We
probably cannot/should not use github.io for this as we may hit the limits on number of downloads.  Embedding images is
not a good approach.  I think we need a free/low-cost CDN for this.  Google Firebase is a possibility, but in all cases
we need to check usage limits.  Could we use Firefly as a CDN?

Can we use the Noun Project as a CDN?  Doesn't seem like it, but could investigate.

Software Requirements
=====================

Python
openpyxl Python library
sqlite3 Python library
(will be more depending on the approach that we use for email integration)



Integration with Office 365
===========================
Run PowerShell on the host PC
> Install-Module AzureAD
...
> 

Directory ID:  121aff59-00d6-4263-86bc-56a65d1e4d3b

> Connect-AzureAD -TenantId 121aff59-00d6-4263-86bc-56a65d1e4d3b
<enter username>
Seems to work, but I get a warning that I should use the "latest PowerShell module, the Microsoft Graph PowerShell SDK"...

OK, so now uninstall the AzureAD module as it is deprecated...

> Disconnect-AzureAD
> Uninstall-Module -Name AzureAD

Cannot uninstall as the module is currently in use, even after disconnecting as above and restarting the PowerShell.

Install the new Graph PowerShell SDK (only when the AzureAD module is uninstalled, so...)

Possible alternative and easier approach.  it appears that we are allowed to send emails within our own tenant more easily:

This is from Microsoft support:
https://answers.microsoft.com/en-us/outlook_com/forum/all/how-to-send-mail-using-python-using-an-outlook/e8f1eead-8df5-4dc7-989f-3acda7ce86c5

8<--
Also, smtp_server = "smtp.office365.com" is for smtp auth and the required port will be 587

You can't use any random port to connect; if you are agrees to use smtp auth, you should consider testing the service using the PS cmd to see if it works for you or not:  

Send-MailMessage -SmtpServer smtp.office365.com -Port 587 -From user@domain.com -To recipient@domain.com -Subject "SMTP Auth Client $(Get-Date -Format g)" -Body "This is a test email using SMTP Auth" -UseSSL
8<--

Hmm - I'm pretty sure that won't work with 2FA.

Perhaps do a proof of concept using my personal gmail account?  I can register an APP for that, and it works in a very similar way.

We could then get IT to sort out the necessary permission changes.

I think we can work around with the GraphAPI, but it may take a while.

https://learn.microsoft.com/en-us/graph/api/resources/message?view=graph-rest-1.0

Firefly do not have an official API as far as I can see (they do have some third party integerations, and clearly they do have a REST API, but details are not public, and they seem to treat this as a private API).  Given that they have been recently sold to an American company, I'm not sure what their future roadmap will look like.

Integrating with Microsoft 365 for Outlook and Teams might be a better approach.  Perhaps we should transition to Teams as a Firefly replacement.

