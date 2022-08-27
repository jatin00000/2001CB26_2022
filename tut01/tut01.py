def factorial(x):
    ans = 1
    for i in range(1, x+1):
        ans = ans*i
    print(ans)


x = int(input("Enter the number whose factorial is to be found"))
factorial(x)
