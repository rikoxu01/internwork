import datetime
import threading
import os


def func():
    # get current time
    now_time = datetime.datetime.now()
    print(now_time)
    if now_time.hour == 9 and now_time.minute < 10:
        print("Running script\n")
        os.system("python yieldcurve.py")

    timer = threading.Timer(600, func)
    timer.start()


# Timer: parameters are (delay time in seconds, function to execute)
timer = threading.Timer(5, func)
timer.start()
