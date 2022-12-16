import os

print(os.getcwd())

os.mkdir("Deck")
os.chdir("Deck")
print(os.getcwd())

os.chdir("..")
print(os.getcwd())

os.rmdir("Deck")

print(os.listdir())

print([x for x in os.listdir() if x[-3:] == ".py"]) 