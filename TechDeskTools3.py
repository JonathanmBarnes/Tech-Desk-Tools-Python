# Welcome to TD Tools Python edition
# Written by Jonathan from Autohotkey TDTools. If questions occur please ask Sarah or email me Jonathanmbarnes@outlook.com
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
# packages = ['pyperclip','subprocess','random','pyautogui','pygetwindow','pandas','datetime','PySimpleGUI', 'auto-py-to-exe', 'pyinstaller']
# for each in packages:
#    os.system(f'pip install {each}')
#
#
# Import functions once installed
import pyperclip as clip  # Functions as clipboard copying
import os  # Run CMD Commands or get system stuff
import time  # Gets Time or gives time
import subprocess  # Run Windows functions or CMD
import pyautogui as pg  # This controls a lot of the old AHK, can control mouse and keyboard input
import random  # Self Explanatory
import pandas as pd  # Often used in data, used to read excels or manipulate numbers
import pygetwindow as gw  # Allows for selecting windows, Used in Copy Email Function

# ===========================================================================================================
# --------------------------------------    User Permissions     --------------------------------------------
# ===========================================================================================================


def GetPermAndUser():
    # Define strings for permission checking. This is the string line from CMD net user command
    laps_string = "LAPS-ReadWrite"  # Change if LAPS string ever changes for some reason
    adac_string = "ITS Tech Desk - Reset"
    MainCMDError = False
    # Actual function for getting permissions
    # For the most part this could be re-writen more efficiently but I wrote this as I was getting into python again. This does the job
    username = str(
        subprocess.check_output(
            ["echo", "%username%"],  # Returns the current username from the shell
            text=True,
            creationflags=0x08000000,  # Keeps shell/CMD terminal hidden during use. it runs 3ish times during the start up
            shell=True,
        )
    )[
        :-1
    ]  # The output has an extra 2 characters which are trimmed via this string method
    LoggedAsAdmin = False  # These strings determine if you are both logged in at an admin account and then to determine what admin
    AdminType = "None"
    XSTAdmin = "xst-"
    XTRAAdmin = "xtra-"
    if "-" in username:
        # If logged in username has "-" then it checks if the begining 4 or 5 characters (XST- / XTRA-) match the above.
        # Note python starts at 0, so it checks starting at the character following the 4th or 5th
        if username[:4] == XSTAdmin:
            LoggedAsAdmin = True
            AdminType = XSTAdmin
            AdminUser = username
            username = username[4:]
            HasAdmin = True

        elif username[:5] == XTRAAdmin:
            LoggedAsAdmin = True
            AdminType = XTRAAdmin
            AdminUser = username
            username = username[5:]
            HasAdmin = True
    # If user is not currently logged in as an admin account it runs net user to see if an XST/Xtra account exists
    if not LoggedAsAdmin:
        HasAdmin = True  # Default is true mainly for error correction.
        AdminUser = (
            XSTAdmin + username
        )  # Due to both being strings the adding just concatates them
        # The following runs the CMD line and saves it as a text. This is done because it other would return a string output with \n and such
        # the subprocess module runs a list, any space in the sequence is indicated by a comma
        try:
            ADMCMD = subprocess.check_output(
                ["net", "user", AdminUser, "/domain"],
                text=True,
                creationflags=0x08000000,
                shell=True,
            )
        except (
            subprocess.CalledProcessError
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
                    AdminUser = "No Network Admin Account"
                    HasAdmin = False
    else:
        ADMCMD = subprocess.check_output(
            ["net", "user", AdminUser, "/domain"],
            text=True,
            creationflags=0x08000000,
            shell=True,
        )
    try:
        CMD = subprocess.check_output(
            ["net", "user", username, "/domain"],
            text=True,
            creationflags=0x08000000,
            shell=True,
        )
    except subprocess.CalledProcessError as e:
        MainCMDError = True

    # Pulls the data between Name and Comment in the CMD output. returns the full name of logged in user
    if not MainCMDError:
        S_index = CMD.find("Full Name")
        E_index = CMD.find("Comment")
        Substring = CMD[S_index:E_index]
        Users_name = Substring[29:]
    else:
        Users_name = username

    if username == "varn3146" or username == "drkuhns" or username == "lars1716":
        LoggedAsAdmin = True

    # Checks if set variable strings above are contained within the CMD permissions.
    LAPS = False
    MFA = False
    try:
        if laps_string in ADMCMD:
            LAPS = True
    except UnboundLocalError as StringError:
        NoCMD = StringError
    try:
        if adac_string in ADMCMD:
            MFA = True
    except UnboundLocalError as StringError:
        NoCMD = StringError

    # The below strings are there only because the program doesn't want to run if you do not have a stthomas account. It shouldn't ever need to be run but just incase.
    # Variable is named Hold because if I say that user = user it breaks
    try:
        Hold = User(
            username, Users_name, LoggedAsAdmin, AdminUser, LAPS, MFA, AdminType
        )
    except TypeError as e:
        Hold = User(
            None,
            "Guest, Guest",
            LoggedAsAdmin=False,
            AdminUser=None,
            LAPS=False,
            MFA=False,
            AdminType="none",
        )
    return Hold


# ===========================================================================================================
# --------------------------------------    Functions and Class's     ---------------------------------------
# ===========================================================================================================
#
# User class used to store info from the permissions function above. it is easier to pass a user and just call user.xxx to grab info
#
class User:
    def __init__(
        User,
        Username: str,
        Name: str,
        LoggedAsAdmin: bool,
        AdminUser: str,
        LAPSPerm: bool,
        MFAPerm: bool,
        AdminType: str,
    ):
        User.Username = Username
        User.Name = Name
        User.FName = Name.partition(",")[-1].split()[0]
        User.AdminStatus = AdminType
        User.AdLog = LoggedAsAdmin
        User.AdUser = AdminUser
        User.AdEmail = f"{AdminUser}@stthomas.edu"
        User.Email = f"{Username}@stthomas.edu"
        User.Laps = LAPSPerm
        User.Mfa = MFAPerm


# Below is the brains of TD Tools and all the links
#
#
class Open:
    # Default Browser launch location. Should this change or microsoft make another browser replace this line for another browser
    # For the most part I will not add notes into this as its just URL's
    browser = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    Map = [browser, "https://campusmap.stthomas.edu/"]

    def __init__(Open, User: User):
        LoggedAsAdmin = (
            User.AdLog
        )  # These were mainly copied from previously exist functions I wrote being using a class, you could replace LoggedAsAdmin and delete this line
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
        Open.MakingTicketEmail = [
            Open.browser,
            "-inprivate" if not LoggedAsAdmin else "",
            "https://services.stthomas.edu/TDClient/1898/ClientPortal/KB/ArticleDet?ID=115750",
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
        time.sleep(3)
        pg.hotkey("alt", "shift", "c")
        time.sleep(2)
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


def CopyEmailOnline():  # Preforms shortcuts to copy the email from outlook inbox
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
    with open(r"C:\ReferenceFiles\List.txt", "r") as word_list:
        words = word_list.readlines()
        RanWord1 = random.choice(words).strip()
        RanWord2 = random.choice(words).strip()

    while RanWord1 == RanWord2:
        RanWord2 = random.choice(words).strip()

    password = RanWord1
    PassHold = RanWord1 + RanWord2
    while len(PassHold) < 15:
        PassHold += "A"
        password += str(random.randint(1, 9))

    password += RanWord2
    return password


def ResetPasswordOld(
    Username: str, UserMustChange: bool
):  # This proved problematic. The function does work but input from the GUI will not work as intended
    Password = str(GenPassword())
    clip.copy(Password)

    # Close any existing PowerShell credential request windows
    try:
        # Run the password reset utility
        ResetScript = r"C:\\ReferenceFiles\\PowerShell\\PWReset.ps1"
        subprocess.call(["powershell.exe", f'Unblock-File "{ResetScript}"'])
        subprocess.call(
            ["powershell.exe", ResetScript] + [Username, Password, UserMustChange]
        )
    except subprocess.CalledProcessError as e:
        test = e


# CMD
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
    VersionNum = "3.2"
    Date = date.today().strftime(
        "%B %d, %Y"
    )  # Grabs current date and format's it to have long spelling of year, numerical day, and year
    Time = datetime.datetime.now().strftime("%I:%M %p")

    # ===========================================================================================================
    # ------------------------------------------   Windows Dimensions   -----------------------------------------
    # ===========================================================================================================

    # App is designed to fit into one corner of the screen.
    # w, h = sg.Window.get_screen_size()  # Gives screen size as a tuple, in pixels. Leaving this just incase you desire to make it scale
    AppW = 1010
    AppH = 540
    ButS = [24, 2]
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
    Green = "#94d600"

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
    NotesText = "Welcome To Tech Desk Tools! \n\nYou can delete this text and use this as a notepad to record info from calls or people at the walkup desks. Above you will find guides for important information you should collect and troubleshooting you can try, below are buttons to open new tickets.\n\n You can find a guide for the rest of the app and a lot of the functions under the Applications Tab. If you have any suggestions or find any issues please submit them to leadership.\n\n-Jonathan"

    # Layout elements
    # Each instancce below is an individual element. This have been done in one gigantic layout but This helps with reading. the justification setting sets the text justification within the element. Not the element itself
    # ===========================================================================================================
    # -----------------------------------------   Top Bar Layout   ----------------------------------------------
    # ===========================================================================================================
    # Greeting Text
    Greeting = "Welcome"
    if User.Username == "more9821":
        Greeting = "¡Quiubo"

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
                    "Search Knowledge Base",
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
                sg.Button("Open Main Apps", size=ButS, font=ButFont),
                sg.Button("Tech Desk Menu", size=ButS, font=ButFont),
            ],
            [
                sg.Button("Enter Hours", size=ButS, font=ButFont),
                sg.Button("Screenshot", size=ButS, font=ButFont),
            ],
        ],
        element_justification="c",
    )

    FirstCall = sg.Column(
        [
            [sg.Push(), sg.Text("Ask Your Question In Teams!", font=HFont), sg.Push()],
            [
                sg.Multiline(
                    key="-FCR-",
                    size=(56, 4),
                    font=MainFont,
                    no_scrollbar=True,
                ),
            ],
            [sg.Push(), sg.Button("Ask For Help", size=ButS, font=ButFont), sg.Push()],
            [sg.VPush()],
        ],
        element_justification="center",
    )

    # ===========================================================================================================
    # -------------------------------------------   CMD Tab Layout   --------------------------------------------
    # ===========================================================================================================

    TabCmd = [
        [sg.VPush()],
        [
            sg.Stretch(),
            sg.Text("User Info", font=ButFont),
            sg.Stretch(),
            sg.Input(key="-NetUser-", size=22, font=MainFont, justification="center"),
            sg.Stretch(),
            sg.Button("Get User Info", font=MainFont, size=ButSSmall),
            sg.Stretch(),
        ],
        [
            sg.Push(),
            sg.Multiline(
                reroute_stdout=True,
                echo_stdout_stderr=True,
                reroute_cprint=True,
                key="-NetIn-",
                size=(60, 20),
                rstrip=True,
                justification="left",
                font=MainFont,
                no_scrollbar=True,
            ),
            sg.Push(),
        ],
        [sg.VPush()],
        [
            sg.Stretch(),
            sg.Text("Ping Client", font=ButFont),
            sg.Stretch(),
            sg.Input(key="-Ping-", size=25, font=MainFont, justification="center"),
            sg.Stretch(),
            sg.Button("Ping", font=MainFont, size=ButSSmall),
            sg.Stretch(),
        ],
        [
            sg.Push(),
            sg.Multiline(
                key="-PingIn-",
                reroute_stdout=True,
                echo_stdout_stderr=True,
                reroute_cprint=True,
                size=(60, 3),
                no_scrollbar=True,
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
        [sg.Push(), sg.T("Important Guides", font=HFont), sg.Push()],
        [
            sg.Push(),
            sg.Button("Things To Ask Clients", size=ButS, font=ButFont),
            sg.Button("P-Zero Troubleshooting", size=ButS, font=ButFont),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Button("How To Fill Out a Ticket", size=ButS, font=ButFont),
            sg.Button("Processing Emails", size=ButS, font=ButFont),
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
        [sg.Push(), sg.T("Create a New Ticket", font=HFont), sg.Push()],
        [
            sg.Push(),
            sg.Button("Standard Request", size=ButS, font=ButFont),
            sg.Button("Priority Zero", size=ButS, font=ButFont),
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
            sg.Button("MIM", font=ButFont, size=ButSSmall),
            sg.Button("Tickets", font=ButFont, size=ButSSmall),
            sg.Button("TD Mailbox", font=ButFont, size=ButSSmall),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Button("Checkout", font=ButFont, size=ButSSmall),
            sg.Button("Bomgar", font=ButFont, size=ButSSmall),
            sg.Button(
                "When To Work",
                font=ButFont,
                size=ButSSmall,
            ),
            sg.Push(),
        ],
        [
            sg.Push(),
            sg.Button("Walkup Guide", font=ButFont, size=ButSSmall),
            sg.Button("Printers", font=ButFont, size=ButSSmall),
            sg.Button("About TD Tools", font=ButFont, size=ButSSmall),
            sg.Push(),
        ],
        [sg.Push(), sg.T("Student Applications", font=HFont), sg.Push()],
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
        [
            sg.Push(),
            sg.Button("How to Reset", font=ButFont, size=ButSSmall),
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
        [
            sg.Push(),
            sg.T("Generate Password", font=ButFont),
            sg.Push(),
            sg.Input(key="-RanPass-", font=ButFont, size=22, justification="c"),
            sg.Push(),
            sg.Button("Generate", font=MainFont),
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
        [sg.VPush()],
        [
            sg.Push(),
            sg.Tree(
                treedata,
                headings=["Email@stthomas.edu", "Phone"],
                font=ButFont,
                col0_heading="Contact Info",
                col0_width=13,
                col_widths=[13, 8],
                num_rows=18,
                hide_vertical_scroll=True,
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
        [sg.Push(), sg.Button("Main Directory", size=ButS, font=ButFont), sg.Push()],
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
            sg.Button("XST Resets", font=ButFont, size=ButSSmall),
            sg.Button("Using Azure", font=ButFont, size=ButSSmall),
            sg.Button("Inventory", font=ButFont, size=ButSSmall),
            sg.Push(),
        ],
        [sg.VPush()],
        [sg.Push(), sg.T("Senior Student Reference Menu", font=HFont), sg.Push()],
        [
            sg.Push(),
            sg.Button("Senior Student Menu", size=ButS, font=ButFont),
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
                        sg.Tab(" Home Tab ", TabMain, font=TabFont),
                        sg.Tab(
                            " Applications ",
                            TabLinks,
                            font=TabFont,
                        ),
                        sg.Tab(
                            " Contact Info ",
                            TabContact,
                            font=TabFont,
                        ),
                        sg.Tab(" CMD ", TabCmd, font=TabFont),
                        sg.Tab(
                            " Senior Students ",
                            TabSenStu,
                            font=TabFont,
                        ),
                    ]
                ],
                background_color=LightPurple,
                font=TabFont,
                border_width=0,
                tab_border_width=0,
                selected_background_color=Gray,
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
        icon=r"C:\Users\xst-barn5203\OneDrive - University of St. Thomas\Desktop\TDT.ico",
        titlebar_icon=r"C:\Users\xst-barn5203\OneDrive-University of St. Thomas\Desktop\TDT.ico",
    )
    window.SetIcon(
        icon=r"C:\Users\xst-barn5203\OneDrive-University of St. Thomas\Desktop\TDT.ico",
        pngbase64=None,
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
        elif event == "Search Knowledge Base" and values["-search-"] != "":
            Apps.Search(User.AdLog, values["-search-"], True, Apps.browser)
            window["-search-"].update("")

        elif (
            event == "-search-" + "_Enter"
            and values["-search-"] != ""
            or event == "Search Using Bing"
        ):
            Apps.Search(User.AdLog, values["-search-"], False, Apps.browser)
            window["-search-"].update("")

        elif event == "Search Knowledge Base" and values["-search-"] == "":
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
            User = values["-NetUser-"]
            window["-NetIn-"].update(UserInfo(User))
            window["-NetUser-"].update("")
            window.refresh()

        elif event == "Ping" or event == "-Ping-" + "_Enter":
            Asset = values["-Ping-"]
            window["-PingIn-"].update(f"Attempting to Ping {Asset}")
            window.refresh()
            window["-PingIn-"].update(Ping(Asset))
            window.refresh()

        elif event == "Generate":
            Password = GenPassword()
            clip.copy(Password)
            window["-RanPass-"].update(Password)
            window.refresh()

        elif event == "Reset Password":
            Username = values["-ResetPassword-"]
            ResetPasswordOld(Username, True)
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

        elif event == "How To Fill Out a Ticket":
            subprocess.Popen(Apps.MakingTicket)

        elif event == "Processing Emails":
            subprocess.Popen(Apps.MakingTicketEmail)


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

# pyinstaller --noconfirm --onedir --windowed --icon "C:/Users/xst-barn5203/OneDrive - University of St. Thomas/Desktop/TDT.ico"
# --splash "C:/Users/xst-barn5203/OneDrive - University of St. Thomas/ITS Stuff/TDTSplash.png" --hidden-import "pyi_splash"
# "C:/Users/xst-barn5203/OneDrive - University of St. Thomas/Desktop/TechDeskToolsPython/BackEnd/TechDeskTools3.py"
