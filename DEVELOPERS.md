# Cicero Word Add-in Development Guide

## Cicero Word Add-in Specific Information

### Development Setup

#### Prerequisites

1. Operating System: Windows or MacOSX
2. MS Word (preferably desktop version)
3. Node JS (preferably v12 as it is the latest LTS)

#### Setup Instructions

1. Clone the project using HTTPS or SSH.
    ```bash
    $ git clone https://github.com/accordproject/cicero-word-add-in.git
    ```

2. Setup the project by running the command 
   ```bash
    npm run setup
   ```
   It will install the dependencies and set up the git hooks.

   Alternatively, one can also use `npm i` to install the depedencies and `npm run prepare`
   to setup the git hooks.

3. Firing up the server diverge a little depending upon the OS you're using.
    1. **On Windows:** Run `npm start` and it will start the server and
       automatically sideload the the add-in on the desktop version of MS Word.
    2. **On MacOSX:** Run `npm run dev-server` for starting the server and run
     `npm start` to sideload the add-in.

4. To test your add-in on the web, run `npm run start:web` and this will start
   the server. You can then sideload the add-in on the web version of MS Word
   by following the steps given [here][addinwebtest].

5. Once the add-in is successfully loaded, the icon for add-in will be displayed
   on the **Home** tab. Go there and click the button with name
   "Cicero Word Add-in".

6. For more information on setting up Word add-in for development, follow this
   [documentation][addindocs].

#### Debugging Add-in

1. **Web version:** Open the browser's developer tools and see the console's
   output.

2. **Desktop version**
   -  **Windows:** [Download the developer tool](https://www.microsoft.com/en-us/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot:overviewtab)
   separately. Once you are able to start the server and sideload the add-in,
   you will see "Accord Project" as an option in the tool.
   ![edge developer tool](readme_assets/devtool.png)
   -  **Mac:**
      1. Open a terminal and run the following 4 commands (this is a one-time setup step).
      ```
      defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true
      defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true
      defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true
      defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true
      ```
      2. Sideload your add-in and open it in Word for Mac.
      3. Right click inside your add-in and choose "Inspect Element".
      4. The Safari Web Inspector console will automatically open, where you can
         debug the add-in the same way you would debug a web app. (ie. use the
         `Console` tab to view printed logs and errors, the `Element` tab to
         inspect the HTML and CSS, the `Network` tab for network calls, etc).

#### Debugging Add-in: Alternate way
If the above method to debug the add-in is not working, one can try the following steps:
   1. Delete the `C:\Users\{username}\.office-addin-dev-certs` folder.
   2. Run `npx office-addin-dev-certs install`.
   3. Now run `npm run dev-server` to startup the server and `npm run start` 
      to open the Add-in in Word.
   4. Click on Add-in to open it.
   5. Instead of attaching the debugger from add-in, open VS Code and click on 
      `Debug and run`. Click on the green arrow in top left corner that says
      `Start Debugging`. The debugger is now running. Open the debug console. 
      Happy debugging :).

## ❗ Accord Project Development Guide ❗
We'd love for you to help develop improvements to Cicero Word Add-in! Please refer to the [Accord Project Development guidelines][apdev] we'd like you to follow.

[apdev]: https://github.com/accordproject/techdocs/blob/master/DEVELOPERS.md
[addindocs]: https://docs.microsoft.com/en-us/office/dev/add-ins/quickstarts/word-quickstart
[addinwebtest]: https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing#sideload-an-office-add-in-in-office-on-the-web