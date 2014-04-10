import ps
from ps import model
import datetime
import xml.etree.ElementTree as ET
import sys, os
import csv
import random

def setVariables(config):
	variable = {}
	config = ET.parse(config)
	root = config.getroot()
	for var in root.find('variableList'):
		variable[var.tag] = var.text
	return variable

def getTime():
	currentTime = datetime.datetime.now()
	return currentTime

def openModel(inputDir, model):
	print "\nOPENING MODEL ", model, ":"
	start = getTime()
	psModel = inputDir + '/' + model
	app = ps.app.Application.instance()
	model = app.importModel(psModel)
	end = getTime()
	print "...Model opened:\t\t\t\t", end-start, "seconds\n"
	return model

def	psSolve(model):
	print "SOLVING MODEL:"
	startSolve = getTime()
	model.solve()
	endSolve = getTime()
	solveTime = endSolve - startSolve
	print "...Solve Time:\t\t\t\t\t", solveTime, "seconds\n"	

def psSaveDxt(model, outputDir, modelName):
	name = modelName.split('.')
	dxtName = name[0] + ".dxt"
	#print dxtName, outputDir
	app = ps.app.Application.instance()
	app.saveModel(outputDir + "\\" + dxtName)

''' 						Base Functions above this line				'''

def	psGetResSched(model, outputDir):
	resSeq = {}
	res = model.getResources()
	for k, r in res.iteritems():
		resSeq[k] = model.solution.findResourceSchedule(r)
	return resSeq

'''
Currently Alternates not incorporated in test
def	psGetAlt(model):
	operations = model.getOperations()
	res = model.getResources()
	R2 = res['R2']
	for oc, oo in operations.iteritems():
		av = oo.getActivity().findDurationResourceAlternates(R2)
		for a in av:
			print oc, a.code
'''
	
def	createAiDict(model):
	aiDict = {}
	ai = model.solution.getActivityInstances()
	for a in ai:
		aiDict[(a.code, a.span[0], a.span[1])] = a
		#print a.code, a.span[0], a.span[1]
	return aiDict

def getDiff(before, after):
	differences = []
	diff = set(after.keys()) - set(before.keys())
	for i, d in enumerate(diff):
		differences.append(d)
		#print "###\t", i, d
	return differences

def psManualSchedule(model, schedule, type):
	''' 
		Iterate thruugh resources and randomly select a number of operation instances for each resources 
		and arbitrarily move +/- 2 days (currently hard coded, maybe parameterize).  Compare the results before and after.
	'''
	resSeq = {}
	res = model.getResources()
	mode = ps.ms.resequencingMode.singleStage if type == Single else ps.ms.resequencingMode.multiStage
	model.solution.repairMode = ps.ms.repairMode.unconstrained		# Set to Whiteboard mode
	manualSchedule = ps.ms.service(model.schedule)

	for k, r in res.iteritems():			# k->resource code, r->resource object
		before = createAiDict(myModel)
		schedule = model.solution.findResourceSchedule(r)
		if schedule:
			count = random.randrange(1, len(schedule))
			#count = random.randrange(1, 20)
			range = count + 1
			toMove = random.sample(xrange((len(schedule))), count)
			print "\nResource", k
			for i in toMove:
				change = random.randrange(-172800, 172800)
				newStart = schedule[i].span[0] + datetime.timedelta(seconds=change)
				print "---Moving %25s %15s from %25s to %25s" %(schedule[i].code, k, schedule[i].span[0], newStart )
				manualSchedule.pasteActivityInstances([schedule[i]], r, newStart, mode)
			model.repair()
			after = createAiDict(myModel)
			diff = getDiff(before, after)
			if len(diff) == len(toMove):
				print "PASS:\t %d changes made %d difference" % (len(toMove), len(diff))
			else:
				print "FAIL:\t", len(toMove), "changes made,", len(diff), "difference"
				for d in diff:
					print "%20s %25s" % (d[0], d[1])
					
 
###################			Main				#########################

if __name__ == "__main__":

	modelXml = sys.argv[1]
	Single = 'Single'
	Multi = 'Multi'	
	variables = setVariables('psPythonConfig.xml')
	outDir = variables['psOutputDirectory']
	myModel = openModel(variables['psInputDirectory'], modelXml)
	psSolve(myModel) 
	resSeq = psGetResSched(myModel, outDir)
	psManualSchedule(myModel, resSeq, Single)

