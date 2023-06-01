# playlist-extractor

This simple little program will take a YouTube playlist and export it to an .xlsx Excel Spreadsheet with a thumbnail, title, hyplerlink, duration and will show a total duration at the top of the spreadsheet.
I created this little program quickly to calculate time spent on programming tutorials and lectures.  It's a handy little tool to quickly catalogue watched vids.  
With the source code you can tweak to your liking.

![image](https://github.com/topoulos/playlist-extractor/assets/1166117/f6b4d8dc-06b8-4d42-8ffc-eeda3f5cbbd4)

You can either edit the appsettings.json file or provide the values from the command line.
```
{
    "Defaults": {
        "GoogleApiKey": "AIzaSyBDATcq3nTSLWLGJCBLUzYBQZs7Ayi7ugM",
        "PlaylistId": "PLvDxU9glBJXenAGK7b-GqTc112-6dy6qi",
        "OutputFilename": "PlayListDuration.xlsx"
    }
}
```
You will need a Google API Key which you can obtain from Google Cloud, which at the time of this initial commit, is free.  Give the key permissions for YouTube.
I would recommend putting the api key in the appsettings.json at least.  You can also set the playlistId and the default file name, or if you don't want to store them you can remove those keys (or the file, it's optional) and you will be prompted when you run the application.
This is just a quick job, feel free to fork and clean it up or contribute if you wish!
