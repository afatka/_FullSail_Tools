'''

Python Tricks The Book. 

First task: 

ContextManager that works with the Python "with" keyword to time functionality.
'''

import time, sys
from contextlib import contextmanager

def Fibonacci(n):
	'''Calculate the Nth number of the fibonacci sequeance'''
	if n == 0: return 0
	elif n == 1: return 1
	else: return Fibonacci(n-1) + Fibonacci(n-2)


class Timer_with:

	def __init__(self):
		self.start_times = [time.clock(), ]

	def __enter__(self):
		self.start_times.append(time.clock())
		return self

	def __exit__(self, exc_type, exc_val, exc_tb):
		start_time = self.start_times.pop()
		print("Duration: {}".format(time.clock() - start_time))

@contextmanager #activate decorator from the contextlib.contextmanager library
def timer_decorator():
	try:
		start_times = [time.clock(),]
		yield
	finally:
		start_time = start_times.pop()
		print("Decorator Duration: {}".format(time.clock() - start_time))

#This runs both context managers
def TimeIt():

	print('test timer_with')

	with Timer_with() as t:
		Fibonacci(30)

	print("completed")

	print("timer_decorator start")

	with timer_decorator() as t:
		Fibonacci(28)

	print("completed")
	# sys.exit()

TimeIt()

# print(Fibonacci(25))
