# playlist-extractor

This simple little program will take a YouTube playlist and export it to an .xlsx Excel Spreadsheet with a thumbnail, title, hyplerlink, duration and will show a total duration at the top of the spreadsheet.

![image](https://github.com/topoulos/playlist-extractor/assets/1166117/f6b4d8dc-06b8-4d42-8ffc-eeda3f5cbbd4)

## Why?
I created this little program quickly to calculate time spent on programming tutorials and lectures.  It's a handy little tool to quickly catalogue watched vids.  
I have kept the application in just a couple of files so that it can be refactored easily into whatever style of application (oop, functional, tdd, etc) you desire.  It exists now mainly as a simple script.
With the source code you can tweak to your liking.

## Usage
You can either edit the appsettings.json file or provide the values from the command line.
```
{
    "Defaults": {
        "GoogleApiKey": "<YOUR API KEY>",
        "PlaylistId": "<PLAYLIST ID>",
        "OutputFilename": "PlayListDuration.xlsx"
    }
}
```
## Important
You will need a Google API Key which you can obtain from Google Cloud https://console.cloud.google.com/apis/credentials, which at the time of this initial commit, is free.  Give the key permissions for YouTube.
I would recommend putting the api key in the appsettings.json at least.  You can also set the playlistId and the default file name, or if you don't want to store them you can remove those keys (or the file, it's optional) and you will be prompted when you run the application.

