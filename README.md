# Script-Formatter
Formats a paragraph-based word document into a script form

The heading of the script file takes the form below where everything but duration is entered in a gui. Duration is calculated from the source file.
```
Title of program: 
Program code: 
Version: 
Duration: 
Script Written By: 
Script Edited By: 
```

The source file is extracted paragraph by paragraph. Each paragraph forms a row in the table along with the scene number, duration (half the word count rounded up) and the visual or audio. If a paragraph begins with "Onscreen title:" it fills the visual column with the title and fills audio column with "**MUSIC**", if it begins with "Chapter heading:" it fills the visual column with the chapter heading and leaves the audio column blank. Otherwise, it just fills the audio column and ignores the visual.
