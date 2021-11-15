import pandas as pd

df = pd.read_csv("./Data/RSM.csv")
title1 = df.RSM.tolist()

df = pd.read_csv("./Data/RSM.csv")
title2 = df.FM.tolist()

def warning(a):
    color = ''
    if a in title1:
        color = '#ffff52'
    elif a in title2:
        color = '#ccff99'
    else:
        color = 'white'
    return color

def number_decorator(number):
    number = round(number, 2)
    number = format(number, ',')
    # number = number + 'K'
    return number