# README #

### What is Cliplink for? ###

* I developed Clipboard-Link from end user need.   It is a reliable method to create links to files and folders on Windows network drives. 
* Spaces and special characters break Outlook's ability to create hyperlinks to network files.  Clipboard-Link eliminates that problem.

### How do I get set up? ###

* There is no configuration
* There are no dependencies
* No database configuration
* Save the contents of this repo to a network drive or share.
* Tell users to run **setup** from there.

### How does it work ###

* Locate the file or folder on the network.
* Right-click and select **Send To** > **clipboard-link**
* Cliplink creates a safe link to the file and puts it on the Clipboard.
* Paste the link into Email, documents, etc.

### To make a friendly link ###

* Create a link to a network file or folder with Clipboard-link and paste into a message
* Select the link part and Cut
* Select the name part between the brackets
* Select Insert menu > Hyperlink
* Paste the link in the URL box
* Click Ok
* Clipboard-link copies the file / folder name and shortened link to the clipboard

```
             [Engineering Documentation]                  file://F:\ENGINE~1 

                      ^ Name ^                               ^   link  ^
```

NOTE: **Cliplink will not create links for executables [files ending with .EXE, .COM, .BAT, etc.].

### Who am I? ###

* Developed by dkurman@netsnbytes.com with lots of help and feedback.  I had jobs as I.T. Manager, sysadmin, helpdesk and programmer for several companies.
* [@davekurman on Twitter](https://twitter.com/davekurman)
* [LinkedIn](https://www.linkedin.com/in/davekurman?trk=profile-badge-cta)
 
