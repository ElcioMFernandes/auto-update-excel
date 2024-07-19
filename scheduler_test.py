import schedule
import time

def job():
    with open("log.txt", "a") as log_file:
        log_file.write("Trabalhando\n")

schedule.every(10).seconds.do(job)

while True:
    schedule.run_pending()
    time.sleep(1)