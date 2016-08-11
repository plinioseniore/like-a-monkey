## Git Data Comparison

Excel has a very poor revision control and this approach solve this using Git (it is supposed you are familiar with), it works fine with list of things (aka database in a  poor speech) that are row oriented (one row for each thing) and where you have a clear primary-key. I use this for revision control of list of field instruments and relevant related information.

### Instructions

***Step 1***

Start from two excel files that you want to compare, ensure that the order of column match and order the primary key alphabetically order, then using Git create a local repository and save the old file as CSV. Commit that file with Git and overwrite it with the newer file saved as CSV, commit again.

As result you have now a Git tracking of two CSV files, request a patch file, this will show all the changes as two lines for each primary-key. The older case is shown with a "-" in front of the primary key and the newer with a "+", that has not changed is not in the patch file.
Generally Git provide you also some lines before and after the change, this is to enhance readability of programming code. So open the patch file in Excel and using a filter you can delete whatever doesn't contains "-" or "+".

***Step 2***

Create a new column just before the primary-key and use a formula like ```=LEFT(B2,1)``` to get the "-" and "+" out of the primary-key, save the result "As Value" and use Find&Replace to remove the "-" and "+" from the primary-key.

You should get something like

|Change|Primary-Key|Others|
|---|---|---|
|-|PT1000|...|
|+|PT1000|...|

Use filters, order Ascending the *Change* column first and *Primary-Key* after. Now you are ready to run the **macro**, it will highlight all changes in lines with "+" using a blue color.

If you have a primary-key that has only "-", this has been deleted in the newer file and viceversa only "+" means that has been added. Changes are shown as two lines as in the example table before.

> A quick tip, this file can be filtered by values and color, to make things easier add in the last row a special char like "_" with the same blue color used for changes. Filtering you fill find it useful.
