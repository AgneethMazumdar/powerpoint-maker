# Purpose

This script uses the python-pptx module to create a Microsoft PowerPoint presentation from a list of titles and urls and places an image onto each slide. The script reads from a two column csv file and downloads, resizes, and positions an image onto each slide with its respective title.

My primary motivation was creating a collection of images of movie characters, but this script can be used to speed up the production of academic powerpoints as well. Moreover, this script can be extended to include additional components such as bullet points, tables, ect.

# Requirements

* [Microsoft Powerpoint](https://products.office.com/en-US/powerpoint?legRedir=true&CorrelationId=76852e31-7bcc-4191-b245-119bfe95cee9)
* [Python 2.7](https://www.python.org/) 
* [Python-pptx](http://python-pptx.readthedocs.org/en/latest/index.html)
* [Pillow](http://pillow.readthedocs.org/) 

Installing Python-pptx using pip comes with several dependencies (including Pillow) but the necessary parts for running the script can be downloaded separately if the user wishes.

File Input and Downloading Images

It's expected that the first column of your csv file is filled with a title and that the cell in the neighboring rows contains a direct link to an image. Note that if you have an already existing file with the same name as one of your titles (minus any spaces) you may want to either change the file or title name or move the file to a different folder so that it doesn't get overwritten by the image from the link that the script downloads.
