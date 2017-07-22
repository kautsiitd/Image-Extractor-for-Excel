# Macro for Extracting Images from Excel and Name accordingly

This Project Contains a VBA Macro for Extracting Images from Excel into a desired folder and name them according to another column. You can follow **How It Works** section below to see how it works. After that, extracting images from Excel will be as easy as one click.

**NOTE:** In excel there is no meaning of cells for an image. They can be put any where in excel file, may be at center of four cells. This was the main reason why python was unable to extract images from excel. I used following process in my code to extract images:

1. Extract all images present in Excel file
2. Identify row number of Image using top left corner of image.
3. Pick up name from row(found out from step 2) and column number(as entered from user) in sheet and save image to directory selected by user with found out name.

Hence It will save all Images present in excel file. More functionality can be added if required, like limiting image selection to particular row number and column number(as right now it's extracting all images).

## Installation



## How It Works

It is simple 3 click process by which you can extract all the images in a selected excel file and name them.

1. Open that Excel file from which you want to extract Images in Excel
2. Click on button in Excel that you added in Excel after following process in Installation section.
3. A window would open now, which will show all files stored in your computer. You have to go to the folder where you want to store all the files. **This folder needs to be created before hand.** And then click OK.
3. Now a window will open which will ask the column number where all names for corresponding images are stored. Enter that number and click OK. **This number should be Integer and not alphabet, where A would map to 1, B to 2 and so on.**

That's it, Now all your images would be extracted in selected folder and named accordingly. You can follow below flow chart for better understanding of how it works.

![alt tag](https://github.com/kautsiitd/Image-Extractor-for-Excel/blob/master/Flow-Chart.png)

## Need Help/Issues

If you find some issue or require some help then you can report about it here: https://github.com/kautsiitd/Image-Extractor-for-Excel/issues.
