def get_standard_deviation_sample(myList, myMean):
    sum = 0
    listCount = len(myList)
    for i in myList:
        sum = sum + (i - myMean) ** 2
    return (sum / (listCount - 1)) ** (1/2)

def get_standard_deviation(myList, myMean):
    sum = 0
    listCount = len(myList)
    for i in myList:
        sum = sum + (i - myMean) ** 2
    return (sum / listCount) ** (1/2)

def get_mean(myList):
    return sum(myList) / len(myList)


def get_MAD(myList, myMean):
    sum = 0
    listCount = len(myList)
    for i in myList:
        sum = sum + abs(i - myMean)
    return sum / listCount

myList = [5,4,6,39]
myMean = get_mean(myList)
print("Mean number:", myMean)

myStdDeviationSample = get_standard_deviation_sample(myList, myMean)
print("Standard Sample Deviation:", myStdDeviationSample)

myStdDeviation = get_standard_deviation(myList, myMean)
print("Standard Deviation:", myStdDeviation)

myMAD = get_MAD(myList, myMean)
print("MAD:", myMAD)

''' DS below '''

# pairs = [('a', 1), ('b', 2), ('c', 3)]
# letters, numbers = zip(*pairs)
# print(letters, numbers)

# from matplotlib import pyplot as plt

# years = [1950, 1960, 1970, 1980, 1990, 2000, 2010]
# gdp = [300.2, 543.3, 1075.9, 2862.5, 5979.6, 10289.7, 14958.3]
# # create a line chart, years on x-axis, gdp on y-axis
# plt.plot(years, gdp, color='green', marker='o', linestyle='solid')
# # add a title
# plt.title("Nominal GDP")
# # add a label to the y-axis
# plt.ylabel("Billions of $")
# plt.show()

# movies = ["Annie Hall", "Ben-Hur", "Casablanca", "Gandhi", "West Side Story"]
# num_oscars = [5, 11, 3, 8, 10]
# # bars are by default width 0.8, so we'll add 0.1 to the left coordinates
# # so that each bar is centered
# xs = [i + 0.1 for i, _ in enumerate(movies)]
# # plot bars with left x-coordinates [xs], heights [num_oscars]
# plt.bar(xs, num_oscars)
# plt.ylabel("# of Academy Awards")
# plt.title("My Favorite Movies")
# # label x-axis with movie names at bar centers
# plt.xticks([i + 0.5 for i, _ in enumerate(movies)], movies)
# plt.show()

# from bokeh.plotting import figure, output_file, show

# # prepare some data
# x = [1, 2, 3, 4, 5]
# y = [6, 7, 2, 4, 5]

# # output to static HTML file
# output_file("lines.html")

# # create a new plot with a title and axis labels
# p = figure(title="simple line example", x_axis_label='x', y_axis_label='y')

# # add a line renderer with legend and line thickness
# p.line(x, y, legend="Temp.", line_width=2)

# # show the results
# show(p)


def vector_add(v, w):
    """adds corresponding elements"""
    return [v_i + w_i
    for v_i, w_i in zip(v, w)]

newSum = vector_add([1, 2], 
           [2, 3])

print(newSum)