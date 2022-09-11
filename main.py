from matplotlib import pyplot as plt
import openpyxl
import math



#Defining some funcitons for statisctics
def sigma_x(x):
    ans = 0
    for elems in x:
        ans += elems
    return ans

def sigma_xy(x,y):
    ans = 0
    for i in range(0,len(x)):
        ans += x[i]*y[i]
    return ans

def sigma_x2(x):
    ans = 0
    for elems in x:
        ans += elems**2
    return ans

#Function to reverse a list
def reverse(x):
    z = []
    for i in range(len(x)-1,-1,-1):
        z.append(x[i])
    return z

#opening Excel sheet and creating an Excel Object
path = "v.xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active


#defining Some useful variables and lists
X = []
Y = []
l = []
be = 2
e = 100
MA =[]

for i in range(be,e):
    X.append(i)

for i in range(be,e):
    cell_obj = sheet_obj.cell(row = i, column = 2)
    Y.append(cell_obj.value)


Y = reverse(Y)

#Finding the slope and y intersept of least squares or regression line
n = e-be
m = (n*sigma_xy(X,Y) - (sigma_x(X)*sigma_x(Y)))/(n*sigma_x2(X) - sigma_x(X)**2)
b = (sigma_x(Y) -m*sigma_x(X))/n

mas = 5

for i in range(0,n):
    ra = []
    if i<mas:
        MA.append(Y[i])
    else:
        for j in range(i-mas+1,i+1):
            ra.append(Y[j])
        MA.append(sigma_x(ra)/mas)
    ra=[]

y1,x1 = [m+b,m*n+b],[1,n]
fig, ax = plt.subplots()
ax.plot(X,Y)
ax.plot(x1,y1,color="red")
ax.plot(X,MA,color="green")
fig.show()

