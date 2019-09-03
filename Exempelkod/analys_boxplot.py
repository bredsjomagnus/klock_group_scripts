import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sys
import getopt

options = ['-i', '--image', '-s', '--show', '-h', '--help', '-v', '--version', '--violin', '--fillzero']
display = "show"
diagramoutput = 'box'
nan = 'dropna'
THRESH = 9 # Value used in dropna for how many allowed non nan values in row.
VERSION = "analys_boxplot.py version 0.1, magnusandersson076@gmail.com"
HELPMSG = """
Plot boxplots from csv containing diamant diagnos results.
CSV header column with no fixt or required set of tests:
[year,termin,classyear,class,name,sva,gender,grade,test1,test2,test3,...]

USAGE
$ python analys_boxplot.py [option] <csv-file>
csv-file works with or without suffix '.csv'

Options:
    -h or --help        Display this help message
    -v or --version     Version
    -i or --image       Save plot to png file
    -s or --show        Show plot. Default.
    -vi or --violin     Violinplot instead of Boxplot
    --fillzero          Set all NaN values to zero. Default is to drop every row
                        with all tests set to NaN. And then drop all NaN values in
                        series before plotting. That way you get into account every
                        written test.
"""


# handle arguments and options (flags)
opts, args = getopt.getopt(sys.argv[1:], "hvis", ["help", "version", "image", "show", "violin", "fillzero"])


# iterate options
for opt in opts:
    print("opt", opt)
    for optvalue in opt:
        # check if option is in valid option list
        if optvalue in options:
            # if option set to -i or --image display is set to image
            if optvalue == '-i' or optvalue == "--image":
                display = 'image'
            if optvalue == '--fillzero':
                nan = 'fillzero'
            if optvalue == '-h' or optvalue == '--help':
                print(HELPMSG)
                input("Press enter...")
                exit()
            if optvalue == '--violin':
                diagramoutput = 'violin'
            if optvalue == '-v' or optvalue == '--version':
                print(VERSION)
                exit()

if len(args) == 0:
    print("Correct use: $ python analys_boxplot.py [option] <csv-file>")
    print("Need CSV-file with or without suffix")
    print()
    print("Type python --help analys_boxplot.py, for help")
    print()
    print("Closing...")
    exit()

csvfile = args[0]

if '.csv' in csvfile:
    csvfile = args[0].split('.')[0]
    print("Suffix found. Splitting on '.'. CSV-file: ", csvfile)

# file
# filename = 'diamant_ak7.csv'

# What columnindex test result start on
TESTINDEX = 8

# File prefix is csvfile that is checked and cleaned from sys.argv[1] above
FILEPREFIX = csvfile

# filename
file = FILEPREFIX+'.csv'

try:
    # create dataframe out of csv-file
    if nan == 'dropna':
        # Keep rows with at least 9 non NaN values. Same as at least one test has value.
        df = pd.read_csv(file).dropna(thresh=THRESH)
        # df.fillna(value=0, inplace=True)
    elif nan == 'fillzero':
        df = pd.read_csv(file).fillna(0)
except Exception as e:
    print("Error", e)
    print("")
    print("Exception raised by pd.read_csv('{}')".format(file))
    print("Are you sure", file, "is the correct csv-file?")
    print()
    print("Closing...")
    exit()

print(df)

# Årskursen
CLASSYEAR = str(df['classyear'].values[0])

# testlist - slices from df header start: TEXTINDEX column, stop: last column
testlist = df.columns.values[TESTINDEX:len(df.columns.values)]

fig, ax = plt.subplots(nrows=1, ncols=len(testlist))
plt.subplots_adjust(hspace=0.45, left=0.04, right=0.96, bottom=0.13)

girls = df.loc[df['gender'] == 'Flicka']
boys = df.loc[df['gender'] == 'Pojke']
sva = df.loc[df['sva'] == 'ja']

not_sva_girls = girls.loc[girls['sva'] == 'nej']
not_sva_boys = boys.loc[boys['sva'] == 'nej']
sva_girls = sva.loc[sva['gender'] == 'Flicka']
sva_boys = sva.loc[sva['gender'] == 'Pojke']

# colors = ['darkgreen', 'sienna', 'midnightblue']
colors = ['#8EBB8E', '#D1AFA0', '#5F809D', '#36A5B0']

## DETTA FUNGERAR ATT GÖRA TRE LÅDDIAGRAM MED RÄTT NAMN UNDER ###
if diagramoutput == 'box':
    for i, test in enumerate(testlist):
        # Drop all nan in series
        ax[i].set_title("Årskurs " + CLASSYEAR + " Flickor/Pojkar/SVA - " + test.upper())
        ax[i].boxplot([girls[test].dropna(), boys[test].dropna(), sva[test].dropna()], widths=0.3)
        ax[i].set_xticks(ticks=[1, 2, 3])
        ax[i].set_xticklabels(['Flickor', 'Pojkar', 'SVA'])
elif diagramoutput == 'violin':
    for i, test in enumerate(testlist):
        
        # drop all nan values left in these series
        not_sva_girls[test].dropna(inplace=True)
        not_sva_boys[test].dropna(inplace=True)
        sva_girls[test].dropna(inplace=True)
        sva_boys[test].dropna(inplace=True)

        try:
            ax[i].set_title("Årskurs " + CLASSYEAR + " Flickor/Pojkar/SVA - " + test.upper())
            vp = ax[i].violinplot([not_sva_girls[test].values, not_sva_boys[test].values, sva_girls[test].values, sva_boys[test].values], showmedians=True)
            ax[i].set_xticks(ticks=[1, 2, 3, 4])
            ax[i].set_xticklabels(['Övriga flickor', 'Övriga pojkar', 'SVA flickor', 'SVA pojkar'])
        except Exception as e:
            # will end up here if series are empty
            pass

        # Get into violine bodies and set color
        for j in range(len(vp['bodies'])):
            vp['bodies'][j].set(facecolor=colors[j])

print(diagramoutput)

if display == "show":
    # SHOW plot full screen
    plt.show()
elif display == 'image':
    # SAVE TO IMAGE WITH SET SIZE
    fig = plt.gcf()
    fig.set_size_inches((15, 7), forward=True)
    plt.savefig(FILEPREFIX+'_'+ diagramoutput +'.png', dpi=100)
    print("Plot saved to file '{}'".format(FILEPREFIX+'_'+ diagramoutput +'.png'))