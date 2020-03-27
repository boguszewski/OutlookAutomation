# OutlookAutomation
Some Automations for Mac Outlook


## Outlook2Things.scpt
  The Script read the internal Outlook ID of an selected Message and open a Things quick entry panel with a Link to the     Email

## OpenOutlookMessageWithID.scpt

  The Script handle links to Outlook Emails *Outlook://<Message-ID>*.
  While Message-ID is the internal message id of outlook.
  To make it work a binary must be created and the URL Handler have to been added in the Info.plist
  
  1. Export the script as an app package in Script Editor as Application
  2. Open Finder and locate the saved file. Right click and select 'Show Package Contents'.
  3. Open Info.plist in the Contents folder and add the following text before the last two lines.
   
   ```xml
        <key>CFBundleIdentifier</key>
        <string>org.personal.outlook</string>
        <key>CFBundleURLTypes</key>
        <array>
            <dict>
                <key>CFBundleURLName</key>
                <string>Pass To OutlookUriHandler</string>
                <key>CFBundleURLSchemes</key>
                <array>
                    <string>outlook</string>
                </array>
            </dict>
        </array>
      ```
