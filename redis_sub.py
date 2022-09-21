import os
import redis
import json
import time
from multiprocessing import Process
from formatPrintout import testPrint, formatDocument

print("STARTED REDIS CON, SCRIPT")

redis_conn = redis.Redis(  host= 'redis-15456.c62.us-east-1-4.ec2.cloud.redislabs.com',
  port= '15456',
  password= 'TduKRzZYDhePnQvIb62w1lrZ6xLekzX6',charset="utf-8", decode_responses=True)


def sub(name: str):
    print("subscribing")
    pubsub = redis_conn.pubsub()
    pubsub.subscribe("broadcast")
    for message in pubsub.listen():
        try:
            print(message)
            data = json.loads(message["data"])
            print("Recieved data" , data)
            formatDocument(data)

        except Exception as e:
            print("Got exception while trying to print ")
            print(e)



if __name__ == "__main__":
   testPrint()
   sub("")
