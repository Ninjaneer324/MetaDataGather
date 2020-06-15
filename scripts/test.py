def inRange(input, index):
    return -1 < input < len(index)

input = "George Orwell 1984"
for c in range(len(input)):
    if (inRange(input, i) for i in range(c, c+4)) and input[c:c+4].isdigit():
        print(input[c:c+4])
        break