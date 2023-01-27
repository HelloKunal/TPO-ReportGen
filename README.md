# TPO-ReportGen

Generates branch wise, date wise reports of training and placement data from the unoffical google sheet.

# Steps to Use

## Set up your directory

- Clone the Repo
From the directory where you want this project added, run the following command.

```
https://github.com/HelloKunal/TPO-ReportGen.git
```

## Set up your environment

- Enable the API
Before using Google APIs, you need to turn them on in a Google Cloud project. You can turn on one or more APIs in a single Google Cloud project.
In the Google Cloud console, enable the Google Sheets API.

```
https://console.cloud.google.com/flows/enableapi?apiid=sheets.googleapis.com
```

- Authorize credentials for a desktop application
To authenticate as an end user and access user data in your app, you need to create one or more OAuth 2.0 Client IDs. A client ID is used to identify a single app to Google's OAuth servers. If your app runs on multiple platforms, you must create a separate client ID for each platform.
In the Google Cloud console, go to Menu menu > APIs & Services > Credentials.
1. Go to Credentials
```
https://console.cloud.google.com/apis/credentials
```
2. Click Create Credentials > OAuth client ID.
3. Click Application type > Desktop app.
4. In the Name field, type a name for the credential. This name is only shown in the Google Cloud console.
5. Click Create. The OAuth client created screen appears, showing your new Client ID and Client secret.
6. Click OK. The newly created credential appears under OAuth 2.0 Client IDs.
7. Save the downloaded JSON file as credentials.json, and move the file to your working directory.