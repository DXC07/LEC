from time import sleep

try:
    from rpa import RPA

    rpa = RPA()

    if __name__ == '__main__':
        rpa.run_RPA()
except Exception as e:
    print(str(e))
    sleep(120)
