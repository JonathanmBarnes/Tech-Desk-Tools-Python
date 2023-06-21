# Welcome to TD Tools Python edition
#
# Written by Jonathan from Autohotkey TDTools. If you have questions please reachout via linkedin or github
#
#
# ===================================================================================================== #
# ____________              __     ________                 __    __________             __             #
# \___    ___/___   ____   |  |__  \______ \   ____   _____|  | __\__    __/____   ____ |  |   ______   #
#    |    | / __ \ / ___\  |  |  \  |  | |  \_/ __ \ /  ___/  |/ /  |    | /  _ \ /  _ \|  |  /  ___/   #
#    |    | \  __/_\  \___ |   Y  \ |  |_|   \  ___/ \___ \|    <   |    |(  <_> |  <_> )  |__\___ \    #
#    |____|  \_____\\_____\|___|__/ /________/\_____>______>__|__\  |____| \____/ \____/|____/______>   #
#                                                                                                       #
# ===================================================================================================== #
#
# Had to keep it, pretty iconic
#
# ===========================================================================================================
# --------------------------------------    Used Packages     -----------------------------------------------
# ===========================================================================================================
# If you are working on TDtools and need the packages downloaded run the code right below and they will install.
# Some may be Preinstalled with python, which is why they will not be listed
#
#
# import os
# packages = ['pyperclip','subprocess','random','pyautogui','pygetwindow','datetime','PySimpleGUI', 'auto-py-to-exe', 'pyinstaller']
# for each in packages:
#    os.system(f'pip install {each}')
#
#
# Import functions once installed
import pyperclip as clip  # Functions as clipboard copying and pasting (Ctrl+C and Ctrl+V)
import os  # Run CMD Commands or get system stuff, this import isn't needed but also does not add any additional space
import time  # Gets Time or gives time
import subprocess  # Run Windows functions or CMD commands
import pyautogui as pg  # This controls a lot of the old AHK, can control mouse and keyboard input
import random  # Self Explanatory

# ===========================================================================================================
# --------------------------------------    User Permissions     --------------------------------------------
# ===========================================================================================================


def GetPermAndUser():
    # Define strings for permission checking. This is the string line from CMD net user command
    laps_string = "LAPS-ReadWrite"  # Change if LAPS string ever changes for some reason
    adac_string = (
        "ITS Tech Desk - Reset"  # Change if ADAC/Password Permission string changes
    )
    TDStudent = "TDX - Student Technic"  # Change if TDX Student string changes
    MainCMDError = False
    AdminCMDError = False
    # Actual function for getting permissions
    # For the most part this could be re-writen more efficiently but I wrote this as I was getting into python again. This does the job
    username = str(
        subprocess.check_output(
            ["echo", "%username%"],  # Returns the current username from the shell
            text=True,  # Returns the output as a string
            creationflags=0x08000000,  # Keeps shell/CMD terminal hidden during use. it runs 3ish times during the start up
            shell=True,  # Runs in windows shell
        )
    )[
        :-1
    ]  # The output has an extra 2 characters which are trimmed via this string method
    LoggedAsAdmin = False  # These strings determine if you are both logged in at an admin account and then to determine what admin
    AdminType = "None"
    XSTAdmin = "xst-"  # Change if student admin string changes
    XTRAAdmin = "xtra-"  # Change if staff admin string changes
    if "-" in username:
        # If logged in username has "-" then it checks if the begining 4 or 5 characters (XST- / XTRA-) match the above.
        # Note python starts at 0, so it checks starting at the character following the 4th or 5th
        if (
            username[:4] == XSTAdmin
        ):  # If the first 4 characters match XST then it sets the following
            LoggedAsAdmin = True
            AdminType = (
                XSTAdmin  # This is used to determine what admin account is being used
            )
            AdminUser = username  # This is used to set the admin account for the rest of the script
            username = username[
                4:
            ]  # This is used to set the username for the rest of the script
            TDUser = True  # This is used to determine if the user is a TD or ITS user

        elif username[:5] == XTRAAdmin:  # Same as above but for XTRA Admins
            LoggedAsAdmin = True
            AdminType = XTRAAdmin
            AdminUser = username
            username = username[5:]
            TDUser = True
        else:  # If the user is not an admin account it sets the following
            LoggedAsAdmin = False
            AdminType = "Other"
            AdminUser = username  # Only set to username because it is not an admin account, allows things to function properly
            username = username

    # If user is not currently logged in as an admin account it runs net user to see if an XST/Xtra account exists
    if not LoggedAsAdmin:
        AdminUser = (
            XSTAdmin + username
        )  # Due to both being strings the adding just concatates them
        # The following runs the CMD line and saves it as a text. This is done because it other would return a string output with \n and such
        # the subprocess module runs a list, any space in the sequence is indicated by a comma
        try:  # Error checking, if the username does not exist it will set the AdminUser to XTRAAdmin + username
            ADMCMD = subprocess.check_output(
                ["net", "user", AdminUser, "/domain"],
                text=True,
                creationflags=0x08000000,
                shell=True,
            )
        except (
            subprocess.CalledProcessError  # This is the error that will be returned if the username does not exist
        ) as e:  # The system will give an error and stop the function if the username does not exist, this lets it continue.
            Error = True  # Will set as error if the above has triggered
            if Error:
                AdminUser = XTRAAdmin + username
            try:
                ADMCMD = subprocess.check_output(
                    ["net", "user", AdminUser, "/domain"],
                    text=True,
                    creationflags=0x08000000,
                    shell=True,
                )
            except subprocess.CalledProcessError as e:
                Error = True
                if Error:
                    AdminUser = False  # If the Admin username does not exist or CMD cannot be run it will set the AdminUser to False
                    AdminCMDError = True  # This is used to determine if the CMD command was able to be run
    else:
        try:  # If the user is logged in as an admin account it will run the CMD command to get the permissions
            ADMCMD = subprocess.check_output(
                ["net", "user", AdminUser, "/domain"],
                text=True,
                creationflags=0x08000000,
                shell=True,
            )
            AdminCMDError = False  # This is used to determine if the CMD command was able to be run, needs to be here otherwise error is thrown later
        except subprocess.CalledProcessError as e:
            AdminCMDError = True

    try:  # Runs the CMD command to get the primary account permissions of the logged in user
        CMD = subprocess.check_output(
            ["net", "user", username, "/domain"],
            text=True,
            creationflags=0x08000000,
            shell=True,
        )
    except subprocess.CalledProcessError as e:
        MainCMDError = True

    # Pulls the data between Name and Comment in the CMD output. returns the full name of logged in user
    if (
        not MainCMDError
    ):  # Error checking, if the CMD command was not able to be run it will skip this
        S_index = CMD.find("Full Name")  # Finds the index of the string "Full Name"
        E_index = CMD.find("Comment")  # Finds the index of the string "Comment"
        Substring = CMD[
            S_index:E_index
        ]  # Sets the substring to the string between the two indexes
        Users_name = Substring[
            29:
        ]  # Sets the Users_name to the substring starting at the 29th character, this is the whole name of a user
    else:  # If the CMD command was not able to be run it will set the Users_name to the logged in username, if local account it will be that accounts name
        Users_name = username

    if username == "varn3146" or username == "drkuhns" or username == "lars1716":
        LoggedAsAdmin = True  # If the username is one of the above it will set LoggedAsAdmin to True
        # Manual override for staff accounts needing to access adac since they can use it on primary account

    # Checks if set variable strings above are contained within the CMD permissions output.
    LAPS = (
        False  # Sets LAPS to False, if the below strings are found it will set to True
    )
    MFA = False  # Sets MFA to False, if the below strings are found it will set to True
    if (
        not AdminCMDError
    ):  # Error checking, if the CMD command was not able to be run it will skip this, prevents error
        try:
            if (
                TDStudent not in ADMCMD
            ):  # If the string is not found in the CMD output it will set TDUser to False
                TDUser = False
        except UnboundLocalError as StringError:  # Error handling
            NoCMD = StringError
        try:
            if (
                laps_string in ADMCMD
            ):  # If the string is found in the CMD output it will set LAPS to True
                LAPS = True
                TDUser = True  # Sets TDUser to True, this is so FTEs can use the app, they may have LAPS but not TDStudent in their permissions
        except UnboundLocalError as StringError:
            NoCMD = StringError
        try:
            if (
                adac_string in ADMCMD
            ):  # If the string is found in the CMD output it will set MFA to True
                MFA = True
                TDUser = True  # Sets TDUser to True, this is so FTEs can use the app, they may have MFA but not TDStudent in their permissions
        except UnboundLocalError as StringError:
            NoCMD = StringError

    # The below strings are there only because the program doesn't want to run if you do not have a stthomas account. It shouldn't ever need to be run but just incase.
    # Variable is named Hold because if I say that user = user it breaks
    if not MainCMDError:
        try:  # General plug in for info gathered above
            Hold = User(
                username,
                Users_name,
                LoggedAsAdmin,
                AdminUser,
                LAPS,
                MFA,
                AdminType,
                TDUser=TDUser,
            )
        except (
            TypeError
        ) as e:  # Error handling, caused if any of the above has a null status. Prevents the program from failing to open
            Hold = User(
                username,
                f"Guest, {username}",
                LoggedAsAdmin=False,
                AdminUser=None,
                LAPSPerm=False,
                MFAPerm=False,
                AdminType="none",
            )
    else:  # This is for when things fail, its a safety net to prevent the program from failing to open. "Guest Login"
        Hold = User(
            username,
            f"Guest, {username}",
            LoggedAsAdmin=False,
            AdminUser=None,
            LAPSPerm=False,
            MFAPerm=False,
            AdminType="none",
        )

    return Hold  # Returns the User class with all the info gathered above


# ===========================================================================================================
# --------------------------------------    Functions and Class's     ---------------------------------------
# ===========================================================================================================
#
# User class used to store info from the permissions function above. it is easier to pass a user and just call user.xxx to grab info
#
class User:
    def __init__(  # Starts class definition
        User,  # Class name
        Username: str,  # Username of the logged in user
        Name: str,  # Full name of the logged in user
        LoggedAsAdmin: bool,  # If the user is logged in as an admin account
        AdminUser: str,  # The admin account the user is logged in as
        LAPSPerm: bool,  # If the user has LAPS permissions
        MFAPerm: bool,  # If the user has MFA permissions
        AdminType: str,  # The type of admin account the user is logged in as
        TDUser: bool = False,  # If the user has TDStudent permissions, defaults to False
    ):  # Ends class definition
        # Assigns all the variables to the class, allows for easy access of the info
        User.Username = Username
        User.Name = Name
        User.FName = Name.partition(",")[-1].split()[
            0
        ]  # First Name. Splits name string by the comma, then splits the first name from the rest of the name.
        User.AdminStatus = AdminType
        User.AdLog = LoggedAsAdmin
        User.AdUser = AdminUser
        User.AdEmail = (
            f"{AdminUser}@stthomas.edu"  # Sets the email of the admin account
        )
        User.Email = f"{Username}@stthomas.edu"  # Sets the email of the logged in user
        User.Laps = LAPSPerm
        User.Mfa = MFAPerm
        User.TDUser = TDUser


# Below is the brains of TD Tools and all the links
# This is the class that is called when the user clicks most buttons
class Open:
    # Default Browser launch location. Should this change or microsoft make another browser replace this line for another browser's location
    # Many people have asked for chrome, we avoid chrome due to security concerns and we are infact a microsoft school
    browser = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    Map = [browser, "https://campusmap.stthomas.edu/"]  # Map link

    def __init__(
        Open, User: User
    ):  # Starts class definition, Calls user class from above
        LoggedAsAdmin = (
            User.AdLog  # Sets LoggedAsAdmin to the AdLog variable from the User class, determines if apps open in incognito for links
        )
        # Each instance creates a list of required things for subprocess.Popen later used. I ran into problems with having that within the class unless I defined a function for each
        Open.MIM = [
            Open.browser,  # Sets default browser to open in
            "-inprivate"
            if LoggedAsAdmin
            else "",  # Determines if they are using an Admin account and if so, will change if opened in incognito.
            # For links requiring an admin login add "not" in the statement like below. examples are such are MIM or student links
            "https://mimportal.stthomas.edu/IdentityManagement/",
        ]
        Open.Tickets = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDNext",
        ]
        Open.KB = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=128248",
        ]
        Open.TDMail = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://outlook.office.com/mail/techdesk@stthomas.edu/inbox",
        ]
        Open.Azure = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://portal.azure.com/#view/Microsoft_AAD_IAM/UsersManagementMenuBlade/~/MsGraphUsers",
        ]
        Open.WTW = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://whentowork.com/logins.htm",
        ]

        Open.Intune = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://endpoint.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/overview",
        ]
        Open.Jamf = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://stthomas.jamfcloud.com/",
        ]
        Open.Laps = ["C:\Program Files\LAPS\AdmPwd.UI.exe"]

        Open.ADAC = ["C:\WINDOWS\system32\dsac.exe"]

        Open.MainPass = [  # Opens public password reset article
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=96839",
        ]

        Open.Verify = [  # Opens Tech Desk article about PW resets and process
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=141571",
        ]

        Open.TDTrain = [  # Open's main training menu
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=128248",
        ]

        Open.SSMenu = [  # Opens Senior Student Menu and Guide
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=140905",
        ]

        Open.WalkupGuide = [  # Opens main guide for the OSF walk-up
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=145519",
        ]

        Open.MailStthomas = [
            Open.browser,
            "-inprivate" if LoggedAsAdmin else "",
            "mail.stthomas.edu",
        ]

        Open.Canvas = [
            Open.browser,
            "-inprivate" if LoggedAsAdmin else "",
            "stthomas.instructure.com",
        ]

        Open.OneUst = [
            Open.browser,
            "-inprivate" if LoggedAsAdmin else "",
            "one.stthomas.edu",
        ]

        Open.Murphy = [
            Open.browser,
            "-inprivate" if LoggedAsAdmin else "",
            "banner.stthomas.edu",
        ]

        Open.Office = [
            Open.browser,
            "-inprivate" if LoggedAsAdmin else "",
            "office.com",
        ]

        Open.NewTicket = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDNext/Apps/1465/Tickets/New",
        ]

        Open.P0Ticket = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDNext/Apps/1465/Tickets/New?formId=53175",
        ]
        Open.MainDirect = [
            Open.browser,
            "-inprivate" if LoggedAsAdmin else "",
            "https://directory.aws.stthomas.edu",
        ]
        Open.AboutTDT = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=141578",
        ]
        Open.EnterTime = [
            Open.browser,
            "-inprivate" if LoggedAsAdmin else "",
            "https://banner.ampr.stthomas.edu/EmployeeSelfService/ssb/timeEntry#/teApp/timesheet/dashboard/payperiod",
        ]
        Open.MainKBHome = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/Home/",
        ]
        Open.MFAProcess = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=150370",
        ]
        Open.UsingLaps = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=150371",
        ]
        Open.JamfIntune = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=150372",
        ]
        Open.UsingAzure = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=150373",
        ]
        Open.TDXInventory = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=150374",
        ]
        Open.XSTReset = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=150375",
        ]
        Open.P0Troubleshoot = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=129557",
        ]
        Open.CollectInfo = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=136191",
        ]
        Open.MakingTicket = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=124544",
        ]
        Open.StatusDash = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://status.stthomas.edu/",
        ]
        Open.AdobeAdmin = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "adminconsole.adobe.com",
        ]
        Open.Salesforce = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://uofstthomasmn.lightning.force.com/lightning/o/Contact/list?filterName=Recent",
        ]
        Open.SalesforceGuide = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=152438",
        ]

    def Search(Open, LoggedAsAdmin: bool, Search: str, KB: bool, browser=browser):
        # Thankfully google and KB have simple search URL's this takes an input splits it then adds a plus sign then adds it to the end of the search string
        Page = "https://www.bing.com/search?q="
        URLSearch = "+".join(Search.split(" "))
        if KB:
            Page = "https://services.stthomas.edu/TDClient/1898/ClientPortal/Shared/Search/?c=all&s="
        url = f"{Page}{URLSearch}"
        subprocess.Popen([browser, "-inprivate" if not LoggedAsAdmin else "", url])

    def OpenMain(
        Open, User: User, browser=browser
    ):  # Will open normal tabs we expect people to have open to do their jobs
        time.sleep(0.8)

        subprocess.Popen(
            [
                browser,
                "-inprivate" if not User.AdLog else "",
                "https://remotesupport.stthomas.edu/saml",
            ]
        )
        time.sleep(0.4)

        subprocess.Popen(
            [
                browser,
                "-inprivate" if not User.AdLog else "",
                "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=128248",
            ]
        )
        time.sleep(0.4)

        subprocess.Popen(
            [
                browser,
                "-inprivate" if not User.AdLog else "",
                "https://outlook.office.com/mail/techdesk@stthomas.edu/inbox",
            ]
        )
        time.sleep(0.4)

        subprocess.Popen(
            [
                browser,
                "-inprivate" if not User.AdLog else "",
                "https://portal.azure.com/#view/Microsoft_AAD_IAM/UsersManagementMenuBlade/~/MsGraphUsers",
            ]
        )
        time.sleep(0.4)
        subprocess.Popen(
            [
                browser,
                "-inprivate" if not User.AdLog else "",
                "https://services.stthomas.edu/TDNext",
            ]
        )
        time.sleep(0.4)
        subprocess.Popen(
            [
                browser,
                "-inprivate" if not User.AdLog else "",
                "https://status.stthomas.edu/",
            ]
        )

        if not User.Laps:
            time.sleep(0.4)

            subprocess.Popen(
                [
                    browser,
                    "-inprivate" if not User.AdLog else "",
                    "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=140931",
                ]
            )

        if not User.AdLog:
            time.sleep(3)
            pg.typewrite(User.AdEmail, interval=0.1)
            time.sleep(0.5)
            pg.press("tab")
            pg.press("tab")
            pg.press("enter")
            return
        else:
            return

    def Bomgar(Open, User: User, Browser=browser):
        # Open Microsoft Edge in private mode to the Remote Support website

        subprocess.Popen(
            [
                Browser,
                "-inprivate" if not User.AdLog else "",
                "https://remotesupport.stthomas.edu/saml",
            ]
        )
        if not User.AdLog:
            time.sleep(2)
            # Type the admin email and press tab once
            pg.typewrite(User.AdEmail, interval=0.1)

    def Checkout(Open, User: User, Browser=browser):
        try:
            subprocess.Popen(
                [
                    Browser,
                    "-inprivate" if not User.AdLog else "",
                    "https://stthomas.webcheckout.net/wco",
                ]
            )
            time.sleep(2)
            pg.typewrite(User.AdUser, interval=0.01)
        except TypeError as e:
            test = e

    def teams(
        Open, User, Input, Browser=browser
    ):  # Will open First Call Reponse. Should this URL change copy the team URL from inside of teams and replace the section between @Thread and 19: with the different channel ID
        subprocess.Popen(
            [
                "start",
                "msteams:/l/channel/19:fd860cc4585e45f399be20b76d9ef1d8@thread.skype/",
            ],
            creationflags=0x08000000,
            shell=True,
        )  # This will copy and type any input into the box above the button into teams. It does not press enter incase of error
        time.sleep(2)
        pg.hotkey("alt", "shift", "c")
        time.sleep(1)
        pg.hotkey("alt", "shift", "c")
        time.sleep(1)
        pg.typewrite(Input, interval=0.01)
        return


# Open Individual Functions
def FindPrinter():  # Open's Find Printer application built into our windows image.
    FindPrinter = r"C:\Windows\Find Printers.qds"
    subprocess.Popen([FindPrinter], shell=True)
    return


def Snip():  # Screenshot
    pg.hotkey("win", "shift", "s")
    return


def CopyEmailOnline():  # Preforms shortcuts to copy the email from outlook inbox to the clipboard
    import pygetwindow as gw  # Allows for selecting windows

    try:
        test = gw.getAllTitles()
        Mailbox = "Mail - Tech Desk - Outlook"
        for i in test:
            if Mailbox in i:
                index = i
        win = gw.getWindowsWithTitle(index)[0]
        win.activate()
        pg.press("insert")  # Flag email
        time.sleep(0.5)
        pg.press("r")  # Open as reply
        time.sleep(0.7)
        pg.hotkey("shift", "right")
        time.sleep(1)
        pg.hotkey("ctrl", "a")
        time.sleep(0.7)
        pg.hotkey("ctrl", "c")
        time.sleep(1)
        pg.press("esc")
        pg.press("esc")
        time.sleep(1)
        pg.press("enter")  # Deletes the draft. or tries its best to delete it
    except UnboundLocalError as e:
        test = e
    return True


# Password (Legacy)
def GenPassword():
    # Generate two random words from the word list
    # Hardcoded only to avoid having to import the wordlist
    words = [
        "Horse",
        "Whale",
        "Tiger",
        "Goose",
        "Camel",
        "Zebra",
        "Moose",
        "Sheep",
        "Mouse",
        "Panda",
        "Snail",
        "Koala",
        "Otter",
        "Rhino",
        "Sloth",
        "Bison",
        "Hippo",
        "Monkey",
        "Turtle",
        "Donkey",
        "Rabbit",
        "Bobcat",
        "Walrus",
        "Badger",
        "Gopher",
        "Lizard",
        "Parrot",
        "Dolphin",
        "Gorilla",
        "Octopus",
        "Green",
        "Orange",
        "Yellow",
        "Purple",
        "Silver",
        "Bronze",
        "Indigo",
        "Emerald",
        "Magenta",
        "Scarlet",
        "Bacon",
        "Candy",
        "Peach",
        "Salad",
        "Steak",
        "Carrot",
        "Cherry",
        "Coffee",
        "Potato",
        "Cheese",
        "Cookie",
        "Hotdog",
        "Radish",
        "Shrimp",
        "Tomato",
        "Turkey",
        "Waffle",
        "Salmon",
        "Avocado",
        "Chicken",
        "Cupcake",
        "Lobster",
        "Pancake",
        "Popcorn",
        "Burger",
        "March",
        "January",
        "Monday",
        "Tuesday",
        "Thursday",
        "Friday",
        "Saturday",
        "Sunday",
        "April",
        "October",
        "November",
        "December",
        "Dance",
        "Smile",
        "Think",
    ]
    RanWord1 = random.choice(words).strip()
    RanWord2 = random.choice(words).strip()
    # Prevent the same word from being chosen twice
    while RanWord1 == RanWord2:
        RanWord2 = random.choice(words).strip()

    # Generate a random number between 1 and 9 in order to reach the 15 character minimum
    password = RanWord1
    PassHold = RanWord1 + RanWord2
    while len(PassHold) < 15:
        PassHold += "A"
        password += str(random.randint(1, 9))

    # Combine the two words and the random number
    password += RanWord2
    return password


# CMD Functions, Calls Netuser
def UserInfo(Username: str):
    try:
        NetUser = subprocess.check_output(
            ["net", "user", Username, "/domain"],
            text=True,
            creationflags=0x08000000,
        )
    except subprocess.CalledProcessError as e:
        NetUser = f"{Username} returned an error. \n \nIt is possible {Username} does not exist as a St. Thomas Microsoft Account. Please check spelling or confirm the Username. \n \nIf the client recently received an email telling them to claim their account, check MIM to see it exists. If it does ask them to wait 24 Hours"
    return NetUser


def Ping(Asset: str):
    param = "-n" if os.name == "nt" else "-c"
    command = ["ping", param, "1", Asset]
    if subprocess.call(command, creationflags=0x08000000) == 0:
        return f"{Asset} is responding and online."
    return f"No connection to {Asset}. \n It may be turned off or have no connection. It also may no longer exist"


# ===============================================================================================================
# -------------------------------------------------    GUI    ---------------------------------------------------
# ===============================================================================================================

import PySimpleGUI as sg
import datetime
from datetime import date


def TDTMain():
    # ===========================================================================================================
    # ----------------------------------   Main Variables and Get Info   ----------------------------------------
    # ===========================================================================================================

    User = GetPermAndUser()
    Apps = Open(User)
    VersionNum = "3.4"
    Date = date.today().strftime(
        "%B %d, %Y"
    )  # Grabs current date and format's it to have long spelling of year, numerical day, and year
    Time = datetime.datetime.now().strftime("%I:%M %p")

    # ===========================================================================================================
    # ------------------------------------------   Windows Dimensions   -----------------------------------------
    # ===========================================================================================================

    # App is designed to fit into one corner of the screen.
    # w, h = sg.Window.get_screen_size()  # Gives screen size as a tuple, in pixels. Leaving this just incase you desire to make it scale
    AppW = 1012
    AppH = 550
    ButS = [24, 2]
    ButSLong = [50, 2]
    ButSSmall = [16, 1]
    # Padding: ((left, right), (top, buttom))
    IPad = 3  # This gives a smaller inner padding
    BPad = 8  # Larger outside border

    # ===========================================================================================================
    # ----------------------------------------------   Colors   -------------------------------------------------
    # ===========================================================================================================

    # Provides HEX numbers for the standard UST branding colors.
    # It is eaiser to read when used as variables and provides an easy way to change the colors in the future

    LightPurple = "#482b58"
    Purple = "#28063b"
    Gray = "#98999b"
    White = "#FFFFFF"
    Black = "#000000"

    # ===========================================================================================================
    # -----------------------------------------   Fonts Variables   ---------------------------------------------
    # ===========================================================================================================

    # The font sizing is weird. I probably wouldn't adjust it.
    TFont = ["AvenirNext LT Pro Medium", 10]
    ButFont = ["AvenirNext LT Pro Regular", 6]
    HFont = ["AvenirNext LT Pro Regular", 10]
    MainFont = ["AvenirNext LT Pro Regular", 5]
    TabFont = ["AvenirNext LT Pro Regular", 6]

    # ===========================================================================================================
    # ------------------------------------   Tech Desk Theme Variable   -----------------------------------------
    # ===========================================================================================================

    # This is tweaked in some places but this sets most colors
    theme_TDT = {
        "BACKGROUND": LightPurple,
        "TEXT": White,
        "INPUT": White,
        "TEXT_INPUT": Black,
        "SCROLL": Purple,
        "BUTTON": (White, Gray),
        "PROGRESS": ("#FFFFFF", "#C7D5E0"),  # Progress bar not used
        "BORDER": 1,  # Sets border width around text box elements
        "SLIDER_DEPTH": 0,
        "PROGRESS_DEPTH": 0,
    }

    sg.LOOK_AND_FEEL_TABLE["TDT"] = theme_TDT
    sg.theme("TDT")
    # ===========================================================================================================
    # ---------------------------------------------    Text    --------------------------------------------------
    # ===========================================================================================================
    #
    NotesText = "Welcome To Tech Desk Tools! \n\nYou can delete this text and use this as a notepad to record info from calls or people at the walkup desks. Below you will find guides for important information you should collect and troubleshooting you can try, above are buttons to open new tickets.\n\n You can find a guide for the rest of the app and a lot of the functions under the Applications Tab. If you have any suggestions or find any issues please submit them to Tech Desk leadership."

    if User.TDUser == False:
        NotesText = "Welcome To Tech Desk Tools! \n\nIf you are seeing this message it means that you are not a Tech Desk User or the application failed to get your permissions.\nFor a permissions error, a common fix is simply to restart the computer\n\nThis however means you do not have access to the fully functioning Tech Desk Tools.\nIf you believe this is an error please contact leadership.\n-Jonathan"

    # Layout elements
    # Each instancce below is an individual element. This have been done in one gigantic layout but This helps with reading. the justification setting sets the text justification within the element. Not the element itself
    # ===========================================================================================================
    # -----------------------------------------   Top Bar Layout   ----------------------------------------------
    # ===========================================================================================================
    # Greeting Text
    Greeting = "Welcome"

    MainRowName = sg.T(
        f" {Greeting} {User.FName}!", font=TFont, justification="left"
    )  # Sets text in top left of app to first name of user
    MainRowTime = sg.T(  # Sets center text as a running clock
        Time,
        key="-Time-",
        font=TFont,
        justification="center",
        size=(25, 1),
    )
    MainRowDate = sg.T(
        f"{Date} ", font=TFont, justification="right"
    )  # Sets text in top corner to the current date

    MainRow = sg.Column(  # Combines the above three elements into a single column element.
        [
            [
                MainRowName,
                sg.Stretch(),
                MainRowTime,
                sg.Stretch(),
                MainRowDate,
            ],  # Stretch() acts as the seperator otherwise they would clumb on the left side regardless of justification because they are not justified within the column
        ],
        background_color=LightPurple,
        pad=((BPad, BPad), (BPad, IPad)),
        expand_x=True,
    )

    # ===========================================================================================================
    # -----------------------------------------   Left Side Layout   --------------------------------------------
    # ===========================================================================================================

    Search = sg.Column(
        [
            [sg.Text("Have a Question?", font=HFont)],
            [
                sg.Push(),
                sg.Input(
                    "",
                    enable_events=False,
                    key="-search-",
                    do_not_clear=True,  # Leave False. Due to how the clock operates if set as true it will not allow typing and constantly clear. I have added a function below that clears the text only after running
                    justification="c",
                    size=(50),
                    font=TabFont,  # appears slightly bigger
                ),
                sg.Push(),
            ],
            [
                sg.Button(
                    "Check the Knowledge Base",
                    bind_return_key=False,
                    size=ButS,
                    font=ButFont,  # Bind_return_key is false due to it being implemented in a different, better way below
                ),
                sg.Button("Search Using Bing", size=ButS, font=ButFont),
            ],
        ],
        element_justification="center",
        expand_x=True,
        expand_y=True,
    )

    DefaultButtons = sg.Column(
        [
            [
                sg.Text("Primary Functions", font=HFont),
            ],
            [
                sg.VPush(),
            ],
            [
                sg.Button(
                    "Open Main Apps",
                    size=ButS,
                    font=ButFont,
                    disabled=True if User.TDUser == False else False,
                ),
                sg.Button(
                    "Tech Desk Menu",
                    size=ButS,
                    font=ButFont,
                    disabled=True if User.TDUser == False else False,
                ),
            ],
            [
                sg.Button(
                    "Enter Hours",
                    size=ButS,
                    font=ButFont,
                    disabled=True if User.TDUser == False else False,
                ),
                sg.Button("Screenshot", size=ButS, font=ButFont),
            ],
        ],
        element_justification="c",
    )

    FirstCall = sg.Column(
        [
            [sg.Push(), sg.Text("Ask First Call Response!", font=HFont), sg.Push()],
            [
                sg.Multiline(
                    key="-FCR-",
                    size=(56, 4),
                    font=MainFont,
                    no_scrollbar=True,
                    disabled=True if User.TDUser == False else False,
                ),
            ],
            [
                sg.Push(),
                sg.Button(
                    "Ask For Help",
                    size=ButSLong,
                    font=ButFont,
                    disabled=True if User.TDUser == False else False,
                ),
                sg.Push(),
            ],
            [sg.VPush()],
        ],
        element_justification="center",
    )

    # ===========================================================================================================
    # -------------------------------------------   CMD Tab Layout   --------------------------------------------
    # ===========================================================================================================

    TabCmd = [
        [
            # sg.Push(),
            sg.Multiline(
                reroute_stdout=True,
                echo_stdout_stderr=True,
                reroute_cprint=True,
                key="-NetIn-",
                size=(61, 22),
                rstrip=True,
                justification="left",
                font=MainFont,
                no_scrollbar=True,
            ),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Input(
                default_text="Enter Username",
                key="-NetUser-",
                size=48,
                font=MainFont,
                justification="center",
            ),
            sg.Push(),
        ],
        [sg.VPush()],
        [sg.VPush()],
        [
            # sg.Push(),
            sg.Multiline(
                key="-PingIn-",
                reroute_stdout=True,
                echo_stdout_stderr=True,
                reroute_cprint=True,
                size=(61, 3),
                no_scrollbar=True,
                font=MainFont,
                justification="center",
            ),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Input(
                default_text="Ping Computer or Website",
                key="-Ping-",
                size=48,
                font=MainFont,
                justification="center",
            ),
            sg.Push(),
        ],
        [sg.VPush()],
    ]

    # ===========================================================================================================
    # ---------------------------------------   Home Tab Layout   -----------------------------------------------
    # ===========================================================================================================

    TabMain = [
        [sg.Push(), sg.T("Create a New Ticket", font=HFont), sg.Push()],
        [
            sg.Push(),
            sg.Button(
                "Standard Request",
                size=ButS,
                font=ButFont,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Button(
                "Priority Zero",
                size=ButS,
                font=ButFont,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Push(),
        ],
        [sg.VPush()],
        [
            sg.Push(),
            sg.Multiline(
                font=MainFont,
                default_text=NotesText,
                size=(60, 12),
                no_scrollbar=True,
                key="-Notes-",
                justification="center",
            ),
            sg.Push(),
        ],
        [sg.Push(), sg.T("Important Guides", font=HFont), sg.Push()],
        [
            sg.Push(),
            sg.Button(
                "Things To Ask Clients",
                size=ButS,
                font=ButFont,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Button(
                "P-Zero Troubleshooting",
                size=ButS,
                font=ButFont,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Button(
                "Walkup Guide",
                size=ButS,
                font=ButFont,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Button(
                "Email and Filling Out Tickets",
                size=ButS,
                font=ButFont,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Push(),
        ],
        [sg.VPush()],
    ]

    # ===========================================================================================================
    # ------------------------------------   Application Tab Layout   -------------------------------------------
    # ===========================================================================================================

    TabLinks = [
        [sg.Push(), sg.T("Tech Desk Links", font=HFont), sg.Push()],
        [
            sg.Push(),
            sg.Button(
                "MIM",
                font=ButFont,
                size=ButSSmall,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Button(
                "Tickets",
                font=ButFont,
                size=ButSSmall,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Button(
                "TD Mailbox",
                font=ButFont,
                size=ButSSmall,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Button(
                "Checkout",
                font=ButFont,
                size=ButSSmall,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Button(
                "Bomgar",
                font=ButFont,
                size=ButSSmall,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Button(
                "When To Work",
                font=ButFont,
                size=ButSSmall,
            ),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Button(
                "Status Dashboard",
                font=ButFont,
                size=ButSSmall,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Button("Printers", font=ButFont, size=ButSSmall),
            sg.Button(
                "About TD Tools",
                font=ButFont,
                size=ButSSmall,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Push(),
        ],
        [sg.Push(), sg.T("General Student Links", font=HFont), sg.Push()],
        [
            sg.Push(),
            sg.Button("Canvas", font=ButFont, size=ButSSmall),
            sg.Button("Outlook", font=ButFont, size=ButSSmall),
            sg.Button("One St. Thomas", font=ButFont, size=ButSSmall),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Button("Microsoft 365", font=ButFont, size=ButSSmall),
            sg.Button("Murphy", font=ButFont, size=ButSSmall),
            sg.Button("Campus Map", font=ButFont, size=ButSSmall),
            sg.Push(),
        ],
        [sg.VPush()],
        [sg.Push(), sg.T("Password Resets", font=HFont), sg.Push()],
        [sg.VPush()],
        [
            sg.Push(),
            sg.T("Generate Password", font=ButFont),
            sg.Push(),
            sg.Input(key="-RanPass-", font=ButFont, size=22, justification="c"),
            sg.Push(),
            sg.Button(
                "Generate",
                font=MainFont,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Push(),
        ],
        [sg.Push()],
        [
            sg.Push(),
            sg.Button(
                "How to Reset",
                font=ButFont,
                size=ButSSmall,
                disabled=True if User.TDUser == False else False,
            ),
            sg.Button(
                "ADAC",
                font=ButFont,
                size=ButSSmall,
                disabled=True if User.AdLog == False else False,
            ),
            sg.Button("Main Article", font=ButFont, size=ButSSmall),
            sg.Push(),
        ],
        [sg.VPush()],
    ]

    # ===========================================================================================================
    # ------------------------------------   Contact info Tab Layout   ------------------------------------------
    # ===========================================================================================================

    data = {
        "AARC": {
            "Title": "",
            "data": ["", "962-5920"],
        },
        "Admissions": {
            "Title": "MHC 250",
            "data": ["admissions", "962-6150"],
        },
        "Alumni": {
            "Title": "",
            "data": ["alumni", "962-6430"],
        },
        "Athletics": {
            "Title": "",
            "data": ["tickets", "962-5900"],
        },
        "Bookstore": {
            "Title": "LL MHC",
            "data": ["bookstore", "962-6850"],
        },
        "Business Office": {
            "Title": "MHC 105",
            "data": ["busoffice", "962-6600"],
        },
        "Dean of Students": {
            "Title": "ASC 241",
            "data": ["deanstudents", "962-6050"],
        },
        "Disability Services": {
            "Title": "MHC 110",
            "data": ["disabilityresources", "962-6315"],
        },
        "Emergency": {
            "Title": "Please Call:",
            "data": ["", "962-5555"],
        },
        "Financial Aid": {
            "Title": "MHC 152",
            "data": ["financialaid", "962-6550"],
        },
        "Health Services": {
            "Title": "CWB",
            "data": ["healthservices", "962-6750"],
        },
        "Human Resources": {
            "Title": "AQU 217",
            "data": ["humanresources", "962-6510"],
        },
        "MPLS General": {
            "Title": "",
            "data": ["tommiecentral", "962-4000"],
        },
        "Off-Campus": {
            "Title": "ASC 253",
            "data": ["offcampus", "962-6138"],
        },
        "Payroll": {
            "Title": "AQU 202",
            "data": ["payroll", "962-6494"],
        },
        "Public Safety": {
            "Title": "MOR 118",
            "data": ["publicsafety", "962-5100"],
        },
        "Registrar      ": {
            "Title": "MHC Main",
            "data": ["registrar", "962-6700"],
        },
        "Residence Life": {
            "Title": "KOC 120",
            "data": ["reslife", "962-6470"],
        },
        "STP General": {
            "Title": "ASC Main",
            "data": ["tommiecentral", "962-5000"],
        },
        "UST Texas": {
            "Title": "Houston, TX",
            "data": ["713-", "522-7911"],
        },
        "Tech Desk MPLS": {
            "Title": "TMH 203B",
            "data": ["Tech Desk", "962-6230"],
        },
        "Tech Desk STP": {
            "Title": "OSF 1FL",
            "data": ["Tech Desk", "962-6230"],
        },
    }

    treedata = sg.TreeData()
    for item in data:
        treedata.insert(parent="", key=f"-{item}-", text=item, values=[])
        treedata.insert(
            parent=f"-{item}-",
            key=f"-{item}_title-",
            text=data[item]["Title"],
            values=[data[item]["data"][0], data[item]["data"][1]],
        )

    TabContact = [
        # [sg.VPush()],
        [
            # sg.Push(),
            sg.Tree(
                treedata,
                headings=["Email@stthomas.edu", "Phone"],
                font=ButFont,
                col0_heading="Contact Info",
                col0_width=13,
                col_widths=[14, 11],
                num_rows=18,
                hide_vertical_scroll=False,
                sbar_width=3,
                sbar_arrow_width=3,
                justification="c",
                background_color=White,
                text_color=Purple,
                border_width=0,
                header_border_width=0,
                header_font=ButFont,
                header_relief="RELIEF_SUNKEN",
                header_background_color=Gray,
                header_text_color=White,
                selected_row_colors=(White, LightPurple),
            ),
            sg.Push(),
        ],
        [sg.VPush()],
        [
            sg.Push(),
            sg.Button("Main Directory", size=ButSLong, font=ButFont),
            sg.Push(),
        ],
        [sg.VPush()],
    ]

    # ===========================================================================================================
    # ----------------------------------   Senior Student Tab Layout   ------------------------------------------
    # ===========================================================================================================

    TabSenStu = [
        [sg.Push(), sg.T("Extra Permissions Required", font=HFont), sg.Push()],
        [
            sg.Push(),
            sg.Button(
                "LAPS",
                size=ButS,
                font=ButFont,
                disabled=True if User.Laps == False else False,
            ),
            sg.Button(
                "Azure",
                size=ButS,
                font=ButFont,
                disabled=True if User.Laps == False else False,
            ),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Button(
                "Intune",
                size=ButS,
                font=ButFont,
                disabled=True if User.Laps == False else False,
            ),
            sg.Button(
                "Jamf",
                size=ButS,
                font=ButFont,
                disabled=True if User.Laps == False else False,
            ),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Button(
                "Salesforce",
                size=ButS,
                font=ButFont,
                disabled=True if User.Laps == False else False,
            ),
            sg.Button(
                "Adobe Admin",
                size=ButS,
                font=ButFont,
                disabled=True if User.Laps == False else False,
            ),
            sg.Push(),
        ],
        [sg.Push(), sg.T("How To Use and Guides", font=HFont), sg.Push()],
        [
            sg.Push(),
            sg.Button("MFA Process", font=ButFont, size=ButSSmall),
            sg.Button("Using LAPS", font=ButFont, size=ButSSmall),
            sg.Button("Jamf and Intune", font=ButFont, size=ButSSmall),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Button("Salesforce Guide", font=ButFont, size=ButSSmall),
            sg.Button("Using Azure", font=ButFont, size=ButSSmall),
            sg.Button("Inventory", font=ButFont, size=ButSSmall),
            sg.Push(),
        ],
        [sg.VPush()],
        [sg.Push(), sg.T("Senior Student Reference Menu", font=HFont), sg.Push()],
        [
            sg.Push(),
            sg.Button("Senior Student Menu", size=ButSLong, font=ButFont),
            sg.Push(),
        ],
        [sg.VPush()],
    ]

    # ===========================================================================================================
    # ---------------------------------    Tab Group Combined Layout   ------------------------------------------
    # ===========================================================================================================

    TabGroups = [
        [
            sg.TabGroup(
                [
                    [
                        sg.Tab(
                            " Home Tab ",
                            TabMain,
                            font=TabFont,
                        ),
                        sg.Tab(
                            " Applications ",
                            TabLinks,
                            font=TabFont,
                        ),
                        sg.Tab(
                            " Directory ",
                            TabContact,
                            font=TabFont,
                        ),
                        sg.Tab(
                            " Get Info ",
                            TabCmd,
                            font=TabFont,
                            disabled=True if User.TDUser == False else False,
                            key="-CMD-",
                        ),
                        sg.Tab(
                            " Senior Students ",
                            TabSenStu,
                            font=TabFont,
                            disabled=True if User.TDUser == False else False,
                        ),
                    ]
                ],
                background_color=LightPurple,
                font=TabFont,
                border_width=0,
                tab_border_width=0,
                selected_background_color=Gray,
                enable_events=True,
            )
        ]
    ]
    # ===========================================================================================================
    # --------------------------------------   Creating Main Layout   -------------------------------------------
    # ===========================================================================================================

    layout = [
        [MainRow],
        [
            sg.Column(
                [
                    [sg.VPush()],
                    [
                        DefaultButtons,
                    ],
                    [
                        Search,
                    ],
                    [FirstCall],
                    [sg.VPush()],
                ],
                size=(AppW * 0.48, AppH * 0.9),
                pad=((BPad, IPad), (IPad, BPad)),
                background_color=LightPurple,
            ),
            sg.Column(
                TabGroups,
                size=(AppW * 0.51, AppH * 0.9),
                pad=((IPad, BPad), (IPad, BPad)),
            ),
        ],
    ]
    # ===========================================================================================================
    # ----------------------------------   Program Window Settings   --------------------------------------------
    # ===========================================================================================================

    window = sg.Window(
        f"Tech Desk Tools {VersionNum} Python",
        layout,
        margins=(0, 0),
        background_color=Purple,
        grab_anywhere=False,
        resizable=False,
        finalize=True,
        auto_size_text=True,
        scaling=2.7,
    )
    B64IconTest = b"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"
    # B64Icon = b"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"
    window.SetIcon(
        # icon=r"C:\Users\xst-barn5203\OneDrive - University of St. Thomas\Desktop\TDT.ico",
        pngbase64=B64IconTest,
    )
    window.set_min_size((AppW, AppH))
    # This allows for the focused element to submit enter. For example if you type in CMD but you have text in the search it will not submit a search as well as CMD
    window["-search-"].bind("<Return>", "_Enter")
    window["-NetUser-"].bind("<Return>", "_Enter")
    window["-Ping-"].bind("<Return>", "_Enter")

    # ===========================================================================================================
    # ----------------------------------   Program Logic and Controls   -----------------------------------------
    # ===========================================================================================================

    # While annoying, every button and element will need an elif statement. For text boxes you can set a key to call or update the input.

    while True:  # Event Loop
        try:
            import pyi_splash  # This is here as a backup for the hidden import. Because sometimes it will not work

            pyi_splash.close()
        except:
            pass

        event, values = window.read(timeout=1000)
        window["-Time-"].update(
            datetime.datetime.now().strftime("%I:%M %p")
        )  # This is the clock brain. The format is 12 hour, Minute, seconds, AM/PM
        window.refresh()
        if (
            event == sg.WIN_CLOSED
        ):  # If closed, program closes and will not break things
            break

        elif event == "Check the Knowledge Base" and values["-search-"] != "":
            Apps.Search(User.AdLog, values["-search-"].strip(), True, Apps.browser)
            window["-search-"].update("")

        elif (
            event == "-search-" + "_Enter"
            and values["-search-"].strip() != ""
            or event == "Search Using Bing"
        ):
            Apps.Search(User.AdLog, values["-search-"].strip(), False, Apps.browser)
            window["-search-"].update("")

        elif event == "Check the Knowledge Base" and values["-search-"] == "":
            subprocess.Popen(Apps.MainKBHome)

        elif event == "Open Main Apps":
            Apps.OpenMain(User)
            window.refresh()
        elif event == "Tech Desk Menu":
            subprocess.Popen(Apps.TDTrain)
            window.refresh()
        elif event == "Copy Email":
            CopyEmailOnline()
            window.refresh()

        elif event == "Enter Hours":
            subprocess.Popen(Apps.EnterTime)

        elif event == "Screenshot":
            Snip()
            window.refresh()

        elif event == "Ask For Help":
            Apps.teams(
                User,
                values["-FCR-"],
            )
        elif event == "Get User Info" or event == "-NetUser-" + "_Enter":
            Client = values["-NetUser-"]
            window["-NetIn-"].update(UserInfo(Client.strip()))
            window["-NetUser-"].update("")
            window.refresh()

        elif event == "Ping" or event == "-Ping-" + "_Enter":
            Asset = values["-Ping-"]
            window["-PingIn-"].update(f"Attempting to Ping {Asset.strip()}")
            window.refresh()
            window["-PingIn-"].update(Ping(Asset.strip()))
            window.refresh()

        elif event == "Generate":
            Password = GenPassword()
            clip.copy(Password)
            window["-RanPass-"].update(Password)
            window.refresh()

        elif event == "Main Article":
            subprocess.Popen(Apps.MainPass)

        elif event == "How to Reset":
            subprocess.Popen(Apps.Verify)

        elif event == "Priority Zero":
            subprocess.Popen(Apps.P0Ticket)

        elif event == "Standard Request":
            subprocess.Popen(Apps.NewTicket)

        elif event == "MIM":
            subprocess.Popen(Apps.MIM)
            if User.AdLog:
                time.sleep(1)
                pg.typewrite(User.Username, interval=0.05)
                time.sleep(0.5)
                pg.press("tab")

        elif event == "Tickets":
            subprocess.Popen(Apps.Tickets)

        elif event == "TD Mailbox":
            subprocess.Popen(Apps.TDMail)

        elif event == "Checkout":
            Apps.Checkout(User)

        elif event == "Bomgar":
            Apps.Bomgar(User)

        elif event == "ADAC":
            subprocess.Popen(Apps.ADAC, shell=True)

        elif event == "Walkup Guide":
            subprocess.Popen(Apps.WalkupGuide)

        elif event == "Printers":
            FindPrinter()

        elif event == "When To Work":
            subprocess.Popen(Apps.WTW)

        elif event == "Canvas":
            subprocess.Popen(Apps.Canvas)

        elif event == "Outlook":
            subprocess.Popen(Apps.MailStthomas)

        elif event == "One St. Thomas":
            subprocess.Popen(Apps.OneUst)

        elif event == "Microsoft 365":
            subprocess.Popen(Apps.Office)

        elif event == "Murphy":
            subprocess.Popen(Apps.Murphy)

        elif event == "Campus Map":
            subprocess.Popen(Apps.Map)

        elif event == "Azure":
            subprocess.Popen(Apps.Azure)

        elif event == "LAPS":
            subprocess.Popen(Apps.Laps, shell=True)

        elif event == "Intune":
            subprocess.Popen(Apps.Intune)
        elif event == "Jamf":
            subprocess.Popen(Apps.Jamf)

        elif event == "Senior Student Menu":
            subprocess.Popen(Apps.SSMenu)

        elif event == "XST Resets":
            subprocess.Popen(Apps.XSTReset)

        elif event == "MFA Process":
            subprocess.Popen(Apps.MFAProcess)

        elif event == "Using LAPS":
            subprocess.Popen(Apps.UsingLaps)

        elif event == "Jamf and Intune":
            subprocess.Popen(Apps.JamfIntune)

        elif event == "Inventory":
            subprocess.Popen(Apps.TDXInventory)

        elif event == "Using Azure":
            subprocess.Popen(Apps.UsingAzure)

        elif event == "Main Directory":
            subprocess.Popen(Apps.MainDirect)

        elif event == "About TD Tools":
            subprocess.Popen(Apps.AboutTDT)

        elif event == "Things To Ask Clients":
            subprocess.Popen(Apps.CollectInfo)

        elif event == "P-Zero Troubleshooting":
            subprocess.Popen(Apps.P0Troubleshoot)

        elif event == "Email and Filling Out Tickets":
            subprocess.Popen(Apps.MakingTicket)

        elif event == "Status Dashboard":
            subprocess.Popen(Apps.StatusDash)

        elif event == "Adobe Admin":
            subprocess.Popen(Apps.AdobeAdmin)

        elif event == "Salesforce":
            subprocess.Popen(Apps.Salesforce)

        elif event == "Salesforce Guide":
            subprocess.Popen(Apps.SalesforceGuide)


if (
    __name__ == "__main__"
):  # Going to be honest, not sure why this syntax matters but it does,
    TDTMain()  # I think it has to do with preventing the rest of the program from throwing
    # errors. TDTMain is the GUI function listed above


# ===========================================================================================================
# --------------------------------------   Program .exe Creation   ------------------------------------------
# ===========================================================================================================

# Run the following command in terminal to open the EXE creator
# auto-py-to-exe
# These are the settings required for TD Tools to launch as expected, Auto-py-to-exe allows a GUI interface for pyinstaller which I Prefered.
#
# Onefile - this sets whether you get a folder or a single app. The version I set onto the desktops in TD is the folder
# This is because the main app is 90mb, and with that comes a long load time, choosing directory(folder) makes launch significantly quicker
#
# Icon - This is the TDT.ico, please be sure to use the one without the white border since it shows poorly on the taskbar
#
# Advanced -> Hidden Import under What to Bundle. you need to add pyi_splash
# The splash screen is the logo that shows up when you try to launch, without the hidden import the logo will
# remain on the screen and not movable.
#
# You will also want to add the TDTSplash file to the --splash option, this is the logo that the system
# Chooses and shows whilst the program is launching. This is not neccesary for function, but is nice so we know the program is working
#
# I personally choose to do a bootloader debugger although to be honest, I don't know if it does anything
#
# Other then that you just choose the output directory under settings and let it run.
# the lines of code below are the terminal input for the setting above, but it sometimes it weird and does not work
# That is why I prefer the GUI

# pyinstaller --noconfirm --onefile --windowed --icon "C:/Users/xst-barn5203/OneDrive - University of St. Thomas/ITS Stuff/TDT.ico"
# --name "Tech Desk Tools" --splash "C:/Users/xst-barn5203/OneDrive - University of St. Thomas/ITS Stuff/TDTSplash.png"
# "C:/Users/xst-barn5203/OneDrive - University of St. Thomas/Desktop/TechDeskToolsPython/BackEnd/TechDeskTools3.py"
