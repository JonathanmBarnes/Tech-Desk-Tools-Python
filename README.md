# Tech Desk Tools Python

#### This is an informative guide about Tech Desk Tools and gives an overview of the tools built into the app as well as an explanation for some of the choices made when designing the application

## Some Quick Background:

Tech Desk Tools Python was made to build on the legacy and functionality of the older Tech Desk Tools. The original tech desk tools was created by a Tech Desk student named Jared Klassen around 2019. He wrote in scripts and features that were commonly used or that could be useful for the Tech Desk, the older version combined these into a single application for easier use. The old program functioned fine and did it's role, however, it was written in AutoHotKey (AHK) which has some limitations when you may want to build or expand functionality and accessibility. It is for that reason we chose to switch it to python in December 2022. Python is easier to read but also allows for more options in its function. This new version was designed to consistently be on your screen and to be useful in many different situations. The application itself does not enable anything special you couldn't do on your own, but instead works to make it easier to accomplish these tasks.

Tech Desk Tools is designed specifically to be used at the University of St. Thomas and operates behind the scenes using things we have and maintain on campus. For that reason should you try to run this at another university or off campus the application will lock away a lot of functionality. 

# Functions Built into the app

## Front Page (Left Side Explanations)

![](https://services.stthomas.edu/TDPortal/Images/Viewer?fileName=34158c1c-2bb8-4450-a278-f10e6aaf7476.png&beidInt=373)

##### **Primary Functions**

-   Four Main Buttons:

    -   Open Main Apps - Opens Applications used commonly at the Tech Desk, the applications change based on senior student title or regular student

    -   Tech Desk Menu - Opens the standard [Tech Desk Training menu](https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=128248) in the knowledge base

    -   Enter Hours - Opens the hour enter page within Murphy Online in the employee portal. This does require you to sign in with your normal UST account

    -   Screenshot - Opens the screenshot tool for windows, should be used to aid documenting issues

**Have a Question?**

-   Contains two buttons, one that opens the KB and one that opens Bing. If there is anything typed and you press a button, it searches in desired engine, by default if you type something and press enter it will search in Bing. This is purposeful, and due to us being a Microsoft campus, the Bing search at times is able to link directly to UST specific help or information relevant to us as a campus. 

##### **Ask First Call Response**

-   Provides a textbox and button that when clicked will open and type the text in first call response, note that it does not send it in case something lags. If you just click the ask for help button it will simply open the First Call Response channel in teams  

**Tabs with further options seen below**

-   Home, Applications, Directory, Get Info, and Senior Students

## Home Tab

![](https://services.stthomas.edu/TDPortal/Images/Viewer?fileName=9c2515a7-dbbb-4f98-a2de-2fcd52a627ae.png&beidInt=373)

**Standard Request Button**

-   This button will take you directly to a form for creating a basic ticket

**Priority Zero**

-   This button will take you directly to the form for creating a Classroom or Event Emergency

**Note Space**

-   This is a temporary note space. Can be used for collecting information from clients or ranting, as you see fit. Please note that text entered does not saved at all and cannot be recovered
-   If the user who is actively signed in is not a member of the Tech Desk or the program is not able to get their permissions this space will show a different text.

## Applications Tab

![](https://services.stthomas.edu/TDPortal/Images/Viewer?fileName=73b56068-21da-4f35-b161-b625c56095c1.png&beidInt=373)

Organized buttons separated based on Tech Desk Links, General Student Links, and Password Resets. 

**Tech Desk Links -** has almost any link that might be needed for work purposes. The links will launch via incognito if needed but should not when logged in as an XST, this excludes MIM.  

-   For MIM you need to log in using your standard UST account which is prompted following clicking the button.

**Student links** - Common links used by all/most students. Included both for the use of the Tech Desk students as well as for getting URLs for referring student callers to.

**Password Resets -** Provides:

-   A quick link button to the #TD KB on resetting a password.
-   A button to open ADAC (the password reset tool/app).
-   A button linking to the Main, University public KB on password resets
-   Password Generator - This creates passwords based off words coded into the application, the passwords will always be sufficient for the UST requirements. It is not recommended to keep these passwords and should set it to something more meaningful afterwards

## Directory Tab

![](https://services.stthomas.edu/TDPortal/Images/Viewer?fileName=e027db8c-6c1d-4557-a378-5fa5cad27a73.png&beidInt=373)

##### The contact info page consists of many different departments which Tech Desk often must refer people too. It also contains UST Houston's phone number for when needed. 

##### The text under the "[email\@stthomas.edu](mailto:email@stthomas.edu){.email}" header is the leading part of the email address. All you need is the @stthomas.edu after any of them.

##### The phone numbers for campus are all **651 (unless specified)**, and the extensions for the departments are the last five in the number listed. 

## Get Info Tab

**Previously named CMD, this was renamed to Get Info to reiterate its function.** 

CMD or Command Prompt, is the default terminal for windows that can communicate directly with the operating system and system. This tab simply runs commands in a hidden terminal window and returns the output into the respective box. All of the commands can be run manually by anyone, with some additional commands requiring admin credentials. 

![](https://services.stthomas.edu/TDPortal/Images/Viewer?fileName=9f73673d-d60b-45bc-920a-64bd6074fa40.png&beidInt=373)

**User Info** - Enter the user's name and enter/return and the CMD info about the user will be output to the open window below. The username clears after entering.

-   The manual function this completes would be to Open Window Command Prompt, type "net user %Username% /domain", press enter
-   If you enter the username wrong or it does not exist an error message will show with some additional instructions

![](https://services.stthomas.edu/TDPortal/Images/Viewer?fileName=ecf1989e-bcbe-4570-8b95-48bc8709693b.png&beidInt=373)

**Ping** - Enter the asset number of the device or the web URL and pressed enter/return, availability of the device on the network or website will be output to the open window below.

-   The manual function this completes would be to Open Window Command Prompt, type "ping %AssetOrUrl%", press enter
-   As with above, if something isn't working it will show an error message and additional info

![](https://services.stthomas.edu/TDPortal/Images/Viewer?fileName=56a07315-f651-43cf-a4ae-b6c213e800ce.png&beidInt=373)

## Senior Students Tab

![](https://services.stthomas.edu/TDPortal/Images/Viewer?fileName=a502fa57-335e-4897-b4a9-f7ef3e5cc2f1.png&beidInt=373)

**Provides links to websites and apps that would be used by a Senior student. They all require additional permissions and credentials to access, some of which are not granted even when appointed too senior student.**

These include:

-   LAPS - Local Admin Password 
-   Azure - used to reset MFA's
-   Intune - Online Microsoft Device Managment software
-   Jamf - Apple Managment software
-   SalesForce -  Currently a work in progress with very few people having access. This requires completion of formal FERPA training
-   Adobe Admin - Able to check active licenses for Adobe Acrobat and Creative Cloud.

It also contains links to KBs with guidelines for the use of these sites/tools.

**The Final button opens the senior student menu which contains much more information and many more guides. It is expected that Senior Students know the information contained within the menu.** 

# Additional Notes

## Maintenance of Application - For St. Thomas ITS

### **Deployment/Installing Tech Desk Tools**

If needed, you can run the PS1 script and provided the asset list is correct it will deploy Tech Desk Tools to all assets within that list. Please be aware you need admin in order to run the script. If needed you can also simply copy the exe to the computer and run it like that, I do warn that Microsoft has been flagging it as a trojan due to it running CMD commands upon opening, the script fixes that. 

For additional information you can open the PS1, it has more instructions and more details. 

### **Updating the App**

The application is written in python and contained within a single .py file. Should you need to edit it, you will need to download the most recent version of python to a computer and an editor. I personally use VS Code. At the top of the .py file I wrote a loop that installs all needed packages within the file, I recommend creating a new .py file and running the loop out of that or just in the terminal.

The application contains instructions about what things do to the best of my ability. If questions arise just reach out to me. The layout of the app and the GUI library takes some getting used to, but it is easier once looking at it a bit, it is difficult to explain ever single line in the GUI. I should add it is still much simpler than alternatives which is why I chose to use it.

## Reason for some choices/FAQ

### **Why is it a single .exe file**

Simplicity. I tried making the app run from a directory of folders and while it launched a tiny bit faster, it made it extremely difficult to update the application. Another reason I switched it was the size reduction that came from making it a single file, it was about 50-60mb less. 

I am not a software engineer by any means and the realistic options we have for updating the application require manual placement of the app on the computers or using a script to automate it remotely. When trying to remove a directory (Folders) it would fail due to the computer not being able to close the app and its processes, so we couldn't delete it when people were working. This left me several times waiting until after 10pm or waking up at 5am to remotely restart computers to kill the processes and deploy the application, which would at times take over 2 hours to copy every file and folder to each computer.

With the single exe file, we are not only able to kill the process at anytime, but also due to the reduced size the install/update time shortens to under 4-5 minutes.

### **Why is the source code contained within a single file instead of containerized**

Growing off the above, simplicity. 

I also prefer writing several files and calling them from within my code, it makes it nicer and cleaner. I think the source code is around 1800 lines of code so it is very overwhelming initially, it also at points is messy and not the cleanest code to look at.

The downside however of creating seperate files is needing to keep track of all the dependent files, for the sake of future upgradability I chose to contain the entire application and anything dependency I would normally contain in separate files within the actual .py file. My hope is that it makes updating simpler in the future, I have found it so when needing to compile it into an exe. 

### **Why Python - More Specifics**

Python is not the first choice for desktop application development typically due to how it operates and interacts with the computer. C, C#, C++ are sort of the gold standard for their speed and at times stability. Python however unlike those is very easy to write and extremely easy to read due to its syntax being very similar to common English. It is also a language often used for introductory programming classes, its simplicity makes it easy to learn for those new to coding, but powerful for those who already do know how to code.

For what we need and use at the Tech Desk, and as I found the limits placed on us in regard to permissions, python offered a significant upgrade in functionality not possible in AutoHotKey without losing any of the functionality previously maintained by the older application, there is even several libraries that mimic AHK within python.  

### **Design Choices**

I'll be the first to admit that its not perfect, but its better.

I tried out several renditions of layouts and ideas we had relating to how it was organized and what would be available within the application. With the amount of info contained and easy of getting to it quickly the tab approach was best for what I was able to do within the python GUI library. The Home/Main page which doesn't change contains things that you are expected to use in some fashion, the application is simply there to make it easier to do them, that follows for most of the information. Nothing within the app requires the application, but it makes it quicker or simpler. And that's what the aim of the program is. 

------------------------------------------------------------------------

I hope this continues to be helpful for a while. Thank you :)
