test_list = [1, 2, 3, 4, 5, 6, 7]

for i in range(10):
    for x in test_list:
        if x == 2:
            print("breaking")
            break
    print(i)