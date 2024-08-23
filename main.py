# Using this file: 
# When you run this file, you will be prompted to paste in the path file of the data. This will be in a format C:\Users\....\....\.....\File name.xlsx (or csv)
# The x-axis will adjust to fit the provided data, the range for the y-axis can be adjusted in line 116


###################################

# libraries used
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt 
import matplotlib
#import mplcursors
from pylab import *
import matplotlib.patches as mpatches

# function for position depending on width
def y_pos1(width):
    # initialising a
    a = np.zeros(len(width))
    # starting from width/2
    a[0] = width[0]/2
    # for loop to get midpoint for each bar
    for i in range(1,len(width)):
        a[i] = a[i-1] + 1.0/2*width[i] + 1.0/2*width[i-1]
    return a

# defining the Mott MacDonald colour scheme
#MM_colour_list = [' #2fb6bc', '#2b86c7', '#afc236', '#ff278d', '#ffd400', '#e95a1a', ' #009695', '#216b99', '#41ac4c', '#8e3f91', '#f8ac04', '#e42925', '#c9c2bf', '#575756']
#quick fix of selecting number of colours as there are categories, and can feed in from this order
MM_colour_list_RGB = [(47, 182, 255),   #288 changed to 255
                      (43, 134, 199),
                      (175, 194, 54),
                      (225, 39, 141),
                      (255, 212, 0),
                      (233, 90, 26),
                      (0, 150, 149)]#,
                     # (33, 107, 153),
                     # (65, 172, 76),
                     # (142, 63, 145),
                     # (248, 172, 4),
                    #  (228, 41, 37),
                    #  (201, 194, 191),
                    #  (87, 87, 86)]

# Convert RGB values to 0-1 range
normalise_rgb_colours = [(r / 255, g / 255, b / 255) for r, g, b in MM_colour_list_RGB]


# function to customize legend
def leg_handles(dataframe):
    # initialisation
    a = []
    # for loop to add the different columns
    for i in range(len(dataframe)):
        patch = mpatches.Patch(color=dataframe.iloc[i,1], label=dataframe.iloc[i,0])
        # appending each category with colour
        a.append(patch)
    return a

    

# need to have Column1, Column2, Column3 and Column 4 in the following order: Category, Option, MtCO2, £/tCO2
def macc_function(path_file, sheetname, starting_rows, Column1, Column2, Column3, Column4):
    # reading the excel file, need to define path, sheet_name, starting rows
    df = pd.read_excel (path_file,sheet_name=sheetname,skiprows=starting_rows)
    # taking the names of the particular columns (need to define them and putting them in a smaller dataframe
    df1 = df[[Column1,Column2,Column3, Column4]]
    # dropping na values
    df1 = df1.dropna()
    # sorting in ascending order
    df1 = df1.sort_values(by=Column4, ascending=True)
    # hierarchy category
    category = df1.iloc[:,0:1].values.squeeze()
    # finding the values for height, bars, width and position
    height = df1.iloc[:,3:].values.squeeze()
    #bars = ["%s" %i for i in range(len(height))]
    bars = df1.iloc[:,1:2].values.squeeze()
    # Choose the width of each bar and their positions
    width = df1.iloc[:,2:3].values.squeeze()
    # turning negative values to positive
    width = np.array([abs(i) if i < 0 else i for i in width])
    #defining the position    
    y_pos = y_pos1(width)
    # number of distinct categories (to be used as colours)
    ndistinct = df1[Column1].nunique()
    # creating a dataframe for the distinct colours (Column1 = Category)
    d = {Column1: df1[Column1].unique(), 'Colour': normalise_rgb_colours}
    df2 = pd.DataFrame(data=d)
    # new df1 by merging colours with existing dataframe (left join)
    df1 = pd.merge(df1,df2,on=Column1,how='left')
    # defining different colours 
    colours = df1.iloc[:,4]
   
    # Make the plot
    # Set the desired thickness of the edge color (outline)
    edge_thickness = 0.5
    #colours = matplotlib.cm.rainbow(np.linspace(0, 1, 7))
    plt.bar(y_pos, height, width = width,color = colours, edgecolor='black', linewidth=edge_thickness)
    # add width (MtCO2e) data labels to each bar
    for i in range(len(width)):
       if i == 0:
           plt.text(y_pos[i], height[i]+1.2, f"{width[i]:.2f}", ha='center', va='bottom', fontsize=5, rotation=90)
       else:
           plt.text(y_pos[i], height[i]+0.5, f"{width[i]:.2f}", ha='center', va='bottom', fontsize=5, rotation=90)
    # plt.xticks(y_pos,bars)
    # removing width labels from x-axis
    # plt.xticks([])
    plt.axhline(0, color='black', linewidth=0.5)
    plt.xticks()
    plt.xticks(rotation=90, fontsize=6)
    plt.yticks(fontsize=8)
    # Add title and axis names
    plt.title('MACC 2019')
    # y-axis title
    plt.ylabel('GBP/tCO$_2$e', fontname="Arial")
    # x-axis title
    plt.xlabel('ktCO$_2$e')
    #Limits for the X axis
    # adjust the width of the x-axis to fit the data
    plt.xlim(y_pos.min() - 0.5* width[0], y_pos.max() + 0.5* width[-1])
    #plt.xlim(0,140)
    # Limits for the Y axis
    plt.ylim(-2,16)
    # hover on option
    #mplcursors.cursor(hover=True).connect( "add", lambda sel: sel.annotation.set(text=f"Category: {category[sel.target.index]}\nOption: {bars[sel.target.index]}\nUSD/tCO$_2$e: {height[sel.target.index]}")) 
    # legend
    legend_handles = leg_handles(df2)
    plt.legend(handles=legend_handles, fontsize=8)
    #plt.legend(loc="best")
    # Adjust the appearance of the spines
    ax = plt.gca()
    ax.spines['top'].set_visible(False)  # Hide the top spine (axis line)
    ax.spines['right'].set_visible(False)  # Hide the right spine (axis line)
    #show
    plt.show()
    plt.savefig("MACC_v1_DfT_2019.png",bbox_inches="tight")
    plt.show()



################################
# Example of inputs
path_file = input("enter the path file of the dataset")
sheetname = "csv_output_2019"
starting_rows = 0 
Column1 = "Category" 
Column2 = "Technology Option"
Column3 = "ktCO2e"
Column4 = "£/tCO2e"
 
# Example of output
example = macc_function(path_file, sheetname, starting_rows, Column1, Column2, Column3, Column4)