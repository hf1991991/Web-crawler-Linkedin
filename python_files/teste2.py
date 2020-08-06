def f1():
    return None

def f2():
    return 1

def f3():
    yield f1() or f2()

print(f3())
print(list(f3()))