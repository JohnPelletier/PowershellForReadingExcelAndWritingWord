# PowershellForReadingExcelAndWritingWord
Sample code to show how to read from excel and write to word with some sorting, formatting and bookmarks figured out

There is a sample sanitized outputfile.docx which shows how my formatting was, just so you can better understand what the code is doing. I put in nonsense names, descriptions and product names, so that my former employer would have no complaints. There should be nothing proprietary in this sample output nor in the code itself at this point. If you find something, please let me know asap and I'll cleanse it immediately. 

It took me a while to get this powershell script right, but when I got it working, it worked for years through all different content, and saved me countless hours.

The formatting in Word and bookmarking were probably the hardest parts. The hard part about the bookmarking was that the Table of Contents (TOC) had to be written first because it's at the top of the document. But what it's bookmarking comes later. I found a simple trick for this using a "special string" to make it easy for me to look up and modify the TOC after writing the rest of the content and creating the real bookmarks.

Anyway, I hope this is helpful to others trying to automate document creations from spreadsheets. I'm pretty sure I'll use it again. 
