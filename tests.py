from fastprogress import progress_bar
import time
for i in  progress_bar(range(1000)):
    time.sleep(.001)
print('Done')