# ARMSCII-8 to UTF-8 MS Office doc converter

# History

In 2005 I was helping to migrate a whole office workflows to GNU/Linux.
They had many gbs of documents in Word and Excel formats, in ARMSCII. I had a choice, whether to add a custom ARMSCII xkb map, or to convert them.
Converting seemd like a right thing to do.

But then there's a problem. It's easy to just get a character and convert. But how to parse Word, Excel documents?

I found out there's this COM way to automate MS Office programs. So I wrote basically 2 programs, for it to be Unix way: one console program, and one ui program that calls the console program.

The program opens Word or Excel, and gets one character out of it. If it is a character from ARMSCII-8 range, it replaces it with UTF-8 character.
This way formatting stays the same. Only the encoding changes.

A couple of years later someone asked me for the solution to exactly same problem. I tried my compiled binaries, and they did not work. I have recompiled the program with newer Delphi verison, and it started working with MS Office 2007.

After that I have never tested the program.
I know that contemporary MS Word is able to edit PDF, so perhaps with a small change it should be possible also to update this to handle PDF translations.

And if it doesn't work for you with contemporary office, then just recompile it with contemporary Embarcadero Delphi Community Edition.














# Old readme


A lot of encoding converters exist.
This one is different because it converts formatted documents.
It is useful in offices where there are a lot of legacy documentation in armscii-8 format.
Requirements: to have ms office installed.

hyaena contains console program and graphical frontend.

With gui frontend it is possible to choose not only one document to convert, but also a directory.
Hyaena will go through all files and directories in that directory recoursively and convert all documents. In this mode files are overwritten.

tested with ms office 2007.
