# map()

items = [1, 2, 3, 4, 5]

squared = list(map(lambda current: current**2, items))

print(squared)

def multiply(x):
    return (x*x)
def add(x):
    return (x+x)

funcs = [multiply, add]
for i in range(5):
    value = list(map(lambda x: x(i), funcs))
    print(value)

# filter()
number_list = range(-5, 5)

less_than_zero = list(filter(lambda current: current < 0, number_list))

print(less_than_zero)

# reduce()
from functools import reduce

product = reduce((lambda current, new: current + new), [1, 2, 3, 4])

print(product)

# zip(), zip(*)

a = [1,2,3]
b = [4,5,6]

a_and_b = list(zip(a,b))
print(a_and_b)

x, y = list(zip(*a_and_b))
print(x)
print(y)

print(isinstance(x, tuple))