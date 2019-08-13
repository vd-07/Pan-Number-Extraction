# Pan-Number-Extraction
Extracts the pan number from pan card and stores in an excel sheet

A video has been attached for better reference.

Completely built using Python 3.7,
This program loads an image directory using the "--images" switch and performs operations on each image one by one.
First there are Image pre-processing techniques like grayscale conversion, blurring, thresholding and morphological operations.
Then it reads the text content of the image using the TESSERACT-OCR Engine. After that, Regular Expression is applied to extract only the Pan Number.
Finally the Pan Number is stored in an Excel Sheet.
