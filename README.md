# Certifypy
 Generator of poster presentation certificates, attendance at conferences, organizing committee and juror participation and awards.


## Features

This python script allows to generate certificates in a massive and automatic way, reading an excel file with the information of the participants of an event or congress. The files are renamed with a base name, the name of the participant and the type of certificate. 

In the excel file the structure of the columns such as: NAME, AFFILIATION, POSTER, TALK, AWARD, and ROLE must be preserved. If a participant has no poster, talk or award information, it should be left blank and the script will only generate the certificates with the information that exists.


## Usage

```
python certifypy.py -i input.dat
```

## Options

The configuration file must contain 3 headers:
* [settings]  
Contains global settings.

* [layout]  
Contains the options for size and orientation of the certificate.

* [info]  
Contains options for the text on the certificate.




