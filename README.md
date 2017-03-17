# createRanking
This script uses an excel point chart to display it as a ranking over a given time range. It will loop through a couple of datapoints and add the description from a table. The description will be formatted with colors and font size. 

# pre requisites

I am using the copenhangenize index (http://copenhagenize.eu/index/#) to demostrate the usage. Every second year, this index ranks a couple of cities worldwide, regarding their bycicle friendlines politics. The first 20 positions of this ranking are public available. 

I put those rankings of three years into an excel table.
The first column contains the cities, the second the ranking in 2011, the third column in 2013 and the fourth column the ranking of each city in 2015.

The ranking only contains 20 participants, but the list has more than 30. If the city is not inside this ranking, it contains no value in the respective year column. Otherwise it's just the ranking number. 

Now I create a point chart in excel. This kind of chart is made of several data rows. Each data row has three dimensions. 

The name of the data row is the cities name. The x-values are the years from 2011 to 2015 - so its not changing. The y-values are represented by the given ranking number. 

# creating the ranking

After creating the point chart, insert the given vba code to macro area of the used worksheet and just run the code. Thats all

# options
there are several parameters to change:
   
1. the font size, background color and label width of each ranking element
   labelFontSize = 14
   labelWidth = 200
   bgColorBrightness = 0.5
   
2. configuring the coloring of the ranking elements, if fixRed/fixGreen/fixBlue is set to 0, the respective color value will be randomized, the randomize function will create a value between the given range (loRed to hiRed). 
   
   fixRed = 0
   loRed = 150
   hiRed = 200
   
   fixGreen = 0
   loGreen = 150
   hiGreen = 200
   
   fixBlue = 0
   loBlue = 150
   hiBlue = 200

3. instead of random colors you can use the index of the ranking element to create a color value for (red / green / blue) - in this case set this parameter to 0

   randomizeColors = 1
   
4. you can connect rankin elements over the time range with a line

   drawLines = 0
