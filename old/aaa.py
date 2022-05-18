import threading
import GraphDati
import LogAITA
import VisualDati

exitFlag = 0

class myThread (threading.Thread):
    def __init__(self, threadID, name, counter, run_script):
       threading.Thread.__init__(self)
       self.threadID = threadID
       self.name = name
       self.counter = counter
       self.run_script = run_script

    def run(self):
       print ("Starting " + self.name)
       if self.run_script == 0:
          LogAITA.main_log()
       if self.run_script == 1:
          GraphDati.main_graph()
       if self.run_script == 2:
          VisualDati.main_grid()

# Create new threads
thread1 = myThread(1, "Log", 1, 0)
thread2 = myThread(2, "RealTime", 2, 2)
#thread3 = myThread(3, "Thread-3", 3, 1)

# Start new Threads
thread1.start()
thread2.start()
#thread3.start()