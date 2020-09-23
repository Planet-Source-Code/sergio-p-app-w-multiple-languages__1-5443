CREATE A MULTILANGUAGE APPLICATION

Whit this Project you can easily change the language
at Run-Time of your Applications.

The structure of the Languages File (.lan) is simple:

xxxx string
xxxx string
xxxx string

where:
--> xxxx is a number between 1 to 99999 (or more)
--> string is the string to load into the captions or ToolTipTexts

I use 4 digits numbers for Captions and 5 digits numbers
for ToolTipTexts. You can change this at your ease.

The 5 digits of ToolTipTexts are composed by the 4 digits of the
Tag of the Control and the Index of the control.
You MUST do this way.

Feedbacks are welcome, also ways to improve it!!!

earthworm@backdoor.to