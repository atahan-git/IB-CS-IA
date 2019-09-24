
n = 10
def testing ():
    global n;
    print()
    print("start")
    try:
        if(n>0):
            print("raise exception")
            n-=1;
            raise Exception("n is " + str(n+1));
    except:
        print("Error catch found")
        return ErrorCatcher(testing);

    print("finally do this")

def ErrorCatcher (callback):
    print("calling callback")
    return callback();

testing();