# AlphaSeeker
Uses statistical analysis to measure the current valuation of a stock on a few key statistics and compare them to how they used to be valued

The function of this script is not to predict future prices of a stock, instead it is used to give the user insight to how the stock used to be valued based on its current fundementals, and compare it to its new price.

This script scrapes data from yahoo finance to get the information using beautiful soup.

Once the stocks data is collected it uses pandas to create a dataframe of the values and uses scipy to get the weighted average of all these fundementals that is most correlated with the adjacent close price.

it then runs this weighted average through a linear regression model using sklearn to give the price.

You run the code using the file alphaseekermain.py 
