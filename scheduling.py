import copy						# For deep copy
from operator import attrgetter	# For parameters sorted()
import xlwt						# Write to excel file
# import xlrd						# Read from excel file
import os						# For opening file


class Process:

	"""
	A process consist of arrival time, burst time and an optional priority.
	"""
	def __init__(self, name, arrival_time, burst_time, priority=None):
		self.name = name
		self.arrival_time = int(arrival_time)
		self.burst_time = int(burst_time)
		self.priority = priority
		self.executed_burst_time = 0
		self.remaining_burst_time = int(burst_time)
		self.active = True

	def reset(self):
		self.name = None
		self.arrival_time = int(0)
		self.burst_time = int(0)
		self.priority = int(0)
		self.executed_burst_time = 0
		self.remaining_burst_time = int(0)
		self.active = False


	def execute(self):
		"""
		If there's remaining burst time, process is executed and remaining burst time 
		decrease by one, executed burst time increase by 1 and return true, else return
		false.
		"""
		if self.remaining_burst_time > 0:
			self.remaining_burst_time -= 1
			self.executed_burst_time += 1
		else:
			pass
	
	def executable(self):
		if self.remaining_burst_time > 0:
			return True
		else:
			return False


	def __str__(self):
		return "Name: {} Executed: {} Remaining: {}".format(self.name, self.executed_burst_time, self.remaining_burst_time)


class CPU:
	"""
	A CPU can run a set of processes with 4 different options:
	 1. FCFS (First Come First Served)-based pre-emptive Priority
	 2. Round Robin Scheduling with Priority
	 3. Three-level Queue Scheduling
	 4. Shortest Remaining Time Next (SRTN) Scheduling with Priority
	"""
	def __init__(self, processes):
		self.processes = processes
		self.total_burst_time = self.calculate_total_burst_time()

	def calculate_total_burst_time(self):
		"""
		Return the total burst time for all the processes in this CPU
		"""
		total = 0
		for p in self.processes:
			total = total + p.burst_time
		return total

	def execute(self, queue):
		"""
		Execute the first process in queue (if available), and return the queue.
		"""
		if len(queue) != 0:
			queue[0].execute()
			if not queue[0].executable():
				# Remove first process from queue
				queue.pop(0)
		return queue

	def FCFS_preemptive_priority(self):
		"""
		Scheduling is based on priority. When two or more processes have 
		the same priority, the process arrived first is selected. If a 
		new process arrives with a higher priority, it will pre-empt the 
		running process. Priority ranges from 1 to 6 with smaller number 
		indicates higher priority.

		Return the sequence of process execution in list.
		"""

		# Remaining burst time
		remaining_burst_time = self.total_burst_time

		# Current time
		current_time = 0

		# Sort all the processes by arrival time
		queue = []

		# Initialize gantt chart
		chart = []

		# Make a copy of the processes to prevent the this function from modifying it
		processes = copy.deepcopy(self.processes)

		# Go through every second until max burst time
		while remaining_burst_time > 0:
			# Add arrived process in waiting queue
			for p in processes:
				if p.arrival_time == current_time:
					queue.append(p)

			# Sort the queue and execute it, else do nothing
			if len(queue) > 0:
				queue = sorted(queue, key=attrgetter('priority', 'arrival_time'))
				chart.append(queue[0].name)
				queue = self.execute(queue)
				remaining_burst_time -= 1
					
			else:
				chart.append('-')

			current_time += 1

		return chart

	def SRTN_preemptive_priority(self):
		"""
		Return the sequence of process execution in list.
		"""
		# Remaining burst time
		remaining_burst_time = self.total_burst_time

		# Current time
		current_time = 0

		# Sort all the processes by arrival time
		queue = []

		# Initialize gantt chart
		chart = []

		# Make a copy of the processes to prevent the this function from modifying it
		processes = copy.deepcopy(self.processes)

		# Go through every second until max burst time
		while remaining_burst_time > 0:
			# Add arrived process in waiting queue
			for p in processes:
				if p.arrival_time == current_time:
					queue.append(p)
					
			# Sort the queue and execute it, else do nothing
			if len(queue) > 0:
				queue = sorted(queue, key=attrgetter('remaining_burst_time', 'priority'))
				chart.append(queue[0].name)
				queue = self.execute(queue)
				remaining_burst_time -= 1			
					
			else:
				chart.append('-')

			
			current_time += 1

		return chart

	def Round_Robin(self,quantum):
		"""
		Return the sequence of process execution in list.
		"""
		# Remaining burst time
		remaining_burst_time = self.total_burst_time

		# Current time
		current_time = 0

		# Sort all the processes by arrival time
		queue = []

		# Initialize gantt chart
		chart = []

		next = Process(name="next", arrival_time=0, burst_time=0, priority=0)
		buffer = Process(name="buffer", arrival_time=0, burst_time=0, priority=0)
		next.reset()
		buffer.reset()

		# Make a copy of the processes to prevent the this function from modifying it
		processes = copy.deepcopy(self.processes)



		# Go through every second until max burst time
		while remaining_burst_time > 0:
			# Add arrived process in waiting queue
			for p in processes:
				if p.arrival_time == current_time:
					queue.append(p)
			queue = sorted(queue, key=attrgetter('priority','arrival_time'))

			if buffer.active == True:
				queue.append(buffer)
				queue = sorted(queue, key=attrgetter('priority','arrival_time'))
				buffer = copy.deepcopy(buffer)
				buffer.reset()

			if len(queue) == 0 and next.active == False:
				chart.append('-')
			else:
				if next.active == False:
					next = copy.deepcopy(queue[0])
					del queue[0]

				next.execute()

				chart.append(next.name)
				remaining_burst_time -= 1


				if next.remaining_burst_time == 0:
					next.reset()
				elif next.executed_burst_time != 0 and next.executed_burst_time % quantum == 0:
					buffer = copy.deepcopy(next)
					if (len(queue) != 0):
						next = copy.deepcopy(queue[0])
						del queue[0]
					else:
						next.reset()

			current_time = current_time + 1

		return chart

	def three_level(self,quantum):
		"""
		Return the sequence of process execution in list.
		"""
		'''total = 0
		for p in self.processes:
			total = total + p.burst_time
		return total'''
		# Remaining burst time
		remaining_burst_time = self.total_burst_time

		# Current time
		current_time = 0

		# Sort all the processes by arrival time
		queue1 = []
		queue2 = []
		queue3 = []

		# Initialize gantt chart
		chart1 = []
		chart2 = []
		chart3 = []

		# Make a copy of the processes to prevent the this function from modifying it
		processes = copy.deepcopy(self.processes)
		
		checkQuantum = quantum

		firsttime = True

		checkprocess = processes

		previouslen = 0

		runappend = True
		#checkburst = quantum;
		# Go through every second until max burst time
		while remaining_burst_time > 0:
			# Add arrived process in waiting queue
			#if check == 0:
			if runappend == True:
				for p in processes:
					if p.arrival_time == current_time:
						if p.priority <= 2:
							queue1.append(p)
						elif p.priority >= 5:
							queue3.append(p)
						else:
							queue2.append(p)

			####################### ROUND ROBIN ###########################
			# sort the processes if queue 1 is running for the first time
			# ensure that the processes in the queue will not be rearrange
			if firsttime == True and len(queue1) > 0:
				queue1 = sorted(queue1, key=attrgetter('priority','arrival_time'))

			# run queue 1 if it is not empty
			if len(queue1) > 0:
				if checkQuantum == 0 and checkprocess == queue1[0]:
					lastProcess = queue1[0]
					queue1.pop(0)
					queue1 = sorted(queue1, key=attrgetter('priority','arrival_time'))
					queue1.append(lastProcess)
					checkQuantum = quantum
				else:
					# if previouslen != len(queue1):
					# 	#if checkQuantum == quantum:
					# 	queue1 = sorted(queue1, key=attrgetter('priority','arrival_time'))
					if len(queue1) > 0:
						if checkprocess != queue1[0]:
							checkQuantum = quantum
							queue1 = sorted(queue1, key=attrgetter('priority','arrival_time'))

				checkprocess = queue1[0]
				chart1.append(queue1[0].name)
				chart2.append('-')
				chart3.append('-')
				queue1 = self.execute(queue1)
				remaining_burst_time -= 1
				current_time += 1
				checkQuantum -= 1

				# if len(queue1) > 0:
				# 	if checkprocess != queue1[0]:
				# 		checkQuantum = quantum
					
				firsttime = False
				previouslen = len(queue1)
				runappend = True

			###################### First Come First Server ###########################
			# run queue 2 if it is not empty
			elif len(queue2) > 0:
				queue2 = sorted(queue2, key=attrgetter('arrival_time', 'priority'))
				chart1.append('-')
				chart2.append(queue2[0].name)
				chart3.append('-')
				
				queue2 = self.execute(queue2)
				remaining_burst_time -= 1
				current_time += 1
				runappend = True

			####################### Q2 waiting queue ###########################
			# move queue 3 into queue 2 if queue 2 is empty
			elif len(queue3) > 0:
				for p in queue3:
					queue2.append(p)
				for i in range(len(queue3)):
					queue3.pop(0)
				runappend = False
			
			####################### No process running ###########################
			# run if there are no processes running
			else:
				chart1.append('-')
				chart2.append('-')
				chart3.append('-')
				current_time += 1
				runappend = True
			 
		return chart1, chart2, chart3

def output_to_excel(charts, processes, file_name, quantum):
	# Initializing spreadsheet
	output = xlwt.Workbook(encoding="utf-8")
	sheet = output.add_sheet("Scheduling")

	row = 1

	# Output data
	for chart_title, chart in charts.items():
		# Print the chart title
		sheet.write(row, 0, chart_title)
		# Print the gannt chart
		for i in range(len(chart)):
			sheet.write(row+1 , i+1, chart[i], xlwt.easyxf("align: vert centre, horiz center"))
			sheet.write(row+2, i, i)
		sheet.write(row+2, len(chart), len(chart))
		row += 4

	# Adjust column width from column 2
	for i in range(len(chart)):
		sheet.col(i+1).width = 2048

	# Display processes table
	# Table header
	sheet.write(row, 1, "Process", xlwt.easyxf("align: vert centre, horiz center"))
	sheet.write(row, 2, "AT", xlwt.easyxf("align: vert centre, horiz center"))
	sheet.write(row, 3, "BT", xlwt.easyxf("align: vert centre, horiz center"))
	sheet.write(row, 4, "Priority", xlwt.easyxf("align: vert centre, horiz center"))

	for p in processes:
		sheet.write(row+1, 1, p.name, xlwt.easyxf("align: vert centre, horiz center"))
		sheet.write(row+1, 2, p.arrival_time, xlwt.easyxf("align: vert centre, horiz center"))
		sheet.write(row+1, 3, p.burst_time, xlwt.easyxf("align: vert centre, horiz center"))
		sheet.write(row+1, 4, p.priority, xlwt.easyxf("align: vert centre, horiz center"))
		row += 1

	# Print quantum time
	sheet.write(row+2, 1, "Q =", xlwt.easyxf("align: vert centre, horiz center"))
	sheet.write(row+2, 2, quantum, xlwt.easyxf("align: vert centre, horiz center"))

	output.save(file_name)



""" The main function"""
file_name = 'output.xls'

processes = []

num_of_processes = int(eval(input("Enter number of processes: ")))

for i in range(num_of_processes):
	process = input(str("Enter arrival time, burst time and priority for P" + str(i+1) + ": "))
	process = process.split(" ")
	try:
		at = int(process[0])
		bt = int(process[1])
		p = int(process[2])
		processes.append(Process(name=str("P"+str(i+1)), arrival_time=at, burst_time=bt, priority=p))
	except ValueError as e:
		print(e) # handle the invalid input
		exit()


quantum = int(input("Enter quantum time: "))

cpu = CPU(processes=processes)

SRTN = cpu.SRTN_preemptive_priority()
FCFS = cpu.FCFS_preemptive_priority()
RR = cpu.Round_Robin(quantum)
level1, level2, level3 = cpu.three_level(quantum)

charts = {"SRTN": SRTN, "FCFS": FCFS, "RR": RR, "Level 1": level1, "Level 2": level2, "Level 3": level3 }

# Output all the scheduling to an Excel file
output_to_excel(charts, processes, file_name, quantum)

# Open the Excel file
os.startfile(file_name)
